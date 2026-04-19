
// serverside/src/services/directMarksImporter.ts

import xlsx from "xlsx";
import mongoose from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Unit from "../models/Unit";
import MarkDirect from "../models/MarkDirect";
import { computeFinalGrade } from "./gradeCalculator";
import type { AuthenticatedRequest } from "../middleware/auth";

interface ImportResult { total: number; success: number; errors: string[]; warnings: string[] }

function stripQualifier(rawRegNo: string): string {
  // Match: slash + 4-digit year + qualifier suffix (letters + digits + optional S2)
  // e.g.  /2017RP1  →  /2017
  //        /2016RP1C →  /2016
  return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
}

// Maps ALL attempt labels used in directTemplate.ts back to DB attempt strings.
function detectAttemptType(rawCell: any): string {
  const raw = (rawCell?.toString() || "").toLowerCase().trim();
  if (!raw)                                                           return "1st";
  if (raw === "a/s" || raw.startsWith("supp"))                       return "supplementary";
  if (raw === "spec" || raw.includes("special"))                     return "special";
  if (/rp\d+c/i.test(raw) || raw === "a/cf")                        return "re-take";
  if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take";
  if (/rpu\d*/i.test(raw))                                           return "re-take";
  if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st";
  return "1st";
}

export async function importDirectMarksFromBuffer(
  buffer: Buffer, filename: string, req: AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  // console.log(`[directImporter] Starting import. File: ${filename}`);

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };
  const workbook  = xlsx.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet     = workbook.Sheets[sheetName];

  // directTemplate puts unit code at F12, year at C8
  const unitCodeRaw     = sheet["F12"]?.v;
  const yearTextRaw     = sheet["C8"]?.v;
  const unitCode        = unitCodeRaw?.toString().trim().toUpperCase();
  const yearText        = yearTextRaw?.toString() || "";
  const yearMatch       = yearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr) {
    throw new Error(
      `Metadata missing. Unit code at F12: "${unitCode}", Academic year at C8: "${academicYearStr}". Check cells F12 and C8.`,
    );
  }

  const unitDoc = await Unit.findOne({ code: unitCode }).lean();
  if (!unitDoc) throw new Error(`Unit "${unitCode}" not found in the database.`);

  const academicYearDoc = await AcademicYear.findOne({
    year:        { $regex: new RegExp(`^${academicYearStr.replace("/", "\\/")}$`, "i") },
    institution: institutionId,
  }).lean();
  if (!academicYearDoc) throw new Error(`Academic Year "${academicYearStr}" not found for this institution.`);

  const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 15 });

  const allProgramUnitsForUnit = await ProgramUnit.find({
    unit: (unitDoc as any)._id,
  }).lean();

  for (const [index, row] of rawRows.entries()) {
    // const regNo = row[1]?.toString().trim().toUpperCase();
    // if (!regNo || regNo === "REG. NO." || regNo === "REG NO") continue;

    const rawCell = row[1]?.toString().trim().toUpperCase();
    if (!rawCell || rawCell === "REG. NO." || rawCell === "REG NO") continue;
    const regNo = stripQualifier(rawCell);

    result.total++;
    const rowNum = index + 16;

    const session = await mongoose.startSession();
    try {
      let savedMarkId: mongoose.Types.ObjectId | null = null;

      await session.withTransaction(async () => {
        const student = await Student.findOne({ regNo, institution: institutionId }).lean();
        if (!student) throw new Error(`Student "${regNo}" not found for this institution`);

        let programUnit = allProgramUnitsForUnit.find(
          (pu: any) => pu.program.toString() === (student as any).program.toString(),
        );

        if (!programUnit) {
          programUnit = await ProgramUnit.findOne({
            program: (student as any).program,
            unit:    (unitDoc as any)._id,
          }).lean() as any;
        }

        if (!programUnit) {
          throw new Error(
            `Unit "${unitCode}" is not linked to the curriculum for student "${regNo}"`,
          );
        }

        const savedMark = await upsertMarkDirect(
          programUnit, student, row, rowNum,
          institutionId, academicYearDoc, req, session,
        );

        savedMarkId = savedMark._id as mongoose.Types.ObjectId;
      });

      // ── Call computeFinalGrade OUTSIDE the transaction ─────────────────────
      // computeFinalGrade opens its own session internally; nesting sessions
      // can deadlock. We call it after the MarkDirect upsert commits.
      if (savedMarkId) {
        try {
          await computeFinalGrade({ markId: savedMarkId });
          // console.log(`[directImporter] Row ${rowNum}: FinalGrade computed for ${regNo}`);
        } catch (gradeErr: any) {
          // Non-fatal: mark was saved, grade calc failed (e.g. missing institution settings)
          console.warn(
            `[directImporter] Row ${rowNum}: FinalGrade calc failed for ${regNo}: ${gradeErr.message}`,
          );
          result.warnings.push(`Row ${rowNum} (${regNo}): grade calc — ${gradeErr.message}`);
        }
      }

      result.success++;
      // console.log(`[directImporter] Row ${rowNum}: SUCCESS`);
    } catch (err: any) {
      const msg = `Row ${rowNum} (${regNo}): ${err.message}`;
      console.error(`[directImporter] FAILED — ${msg}`);
      result.errors.push(msg);
    } finally {
      await session.endSession();
    }
  }

  // console.log(`[directImporter] Done. Total=${result.total}, Success=${result.success}, Errors=${result.errors.length}`);
  return result;
}

// ─── Helper: upsert the MarkDirect document ───────────────────────────────────
// Returns the saved document so the caller can pass its _id to computeFinalGrade.

async function upsertMarkDirect(
  programUnit:    any,
  student:        any,
  row:            any[],
  rowNum:         number,
  institutionId:  any,
  academicYearDoc: any,
  req:            AuthenticatedRequest,
  session:        mongoose.ClientSession,
): Promise<any> {
  // Column layout (directTemplate.ts):
  //  A=0 S/N  B=1 REG.NO.  C=2 NAME  D=3 ATTEMPT
  //  E=4 CA(/30)  F=5 EXAM(/70)  G=6 INTERNAL  H=7 EXTERNAL  I=8 AGREED  J=9 GRADE
  const attempt   = detectAttemptType(row[3]);
  const isSpecial = attempt === "special";
  const isSupp    = attempt === "supplementary";
  const isRetake  = attempt === "re-take";

  const caTotal30   = Number(row[4]) || 0;
  const examTotal70 = Number(row[5]) || 0;

  const rawAgreed  = row[8];
  const agreedMark = (rawAgreed !== undefined && rawAgreed !== null && rawAgreed !== "")
    ? Number(rawAgreed)
    : caTotal30 + examTotal70;

  const externalRaw       = row[7];
  const externalTotal100  = (externalRaw !== undefined && externalRaw !== null && externalRaw !== "")
    ? Number(externalRaw)
    : null;

  const isMissingCA = caTotal30 === 0 && !isSupp && !isSpecial;

  const markData = {
    institution:     institutionId,
    student:         (student as any)._id,
    programUnit:     (programUnit as any)._id,
    academicYear:    (academicYearDoc as any)._id,
    caTotal30,
    examTotal70,
    externalTotal100,
    agreedMark,
    attempt,
    isSpecial,
    isSupplementary: isSupp,
    isRetake,
    isMissingCA,
    uploadedBy:      req.user._id,
    uploadedAt:      new Date(),
    deletedAt:       null,
  };

  const saved = await MarkDirect.findOneAndUpdate(
    {
      student:      (student as any)._id,
      programUnit:  (programUnit as any)._id,
      academicYear: (academicYearDoc as any)._id,
    },
    { $set: markData },
    { upsert: true, new: true, session },
  );

  // console.log(`[directImporter] Row ${rowNum}: MarkDirect upserted _id=${saved._id}` + ` attempt=${attempt} CA=${caTotal30} Exam=${examTotal70} Agreed=${agreedMark}`);

  return saved;
}