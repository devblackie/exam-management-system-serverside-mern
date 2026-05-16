
// serverside/src/services/directMarksImporter.ts — COMPLETE, PRODUCTION READY

import { randomUUID } from "node:crypto";
import xlsx from "xlsx";
import mongoose from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Unit from "../models/Unit";
import MarkDirect from "../models/MarkDirect";
import { computeFinalGrade } from "./gradeCalculator";
import type { AuthenticatedRequest } from "../middleware/auth";

interface ImportResult {
  total: number;
  success: number;
  errors: string[];
  warnings: string[];
}

function stripQualifier(rawRegNo: string): string {
  return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
}

function detectAttemptType(rawCell: unknown): string {
  const raw = (String(rawCell ?? "")).toLowerCase().trim();
  if (!raw) return "1st";
  if (raw === "a/s" || raw.startsWith("supp")) return "supplementary";
  if (raw === "spec" || raw.includes("special")) return "special";
  if (/rp\d+c/i.test(raw) || raw === "a/cf") return "re-take";
  if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take";
  if (/rpu\d*/i.test(raw)) return "re-take";
  if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st";
  return "1st";
}

export async function importDirectMarksFromBuffer(
  buffer: Buffer,
  filename: string,
  req: AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const batchId = randomUUID();
  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const unitCode = sheet["F12"]?.v?.toString().trim().toUpperCase();
  const yearText = sheet["C8"]?.v?.toString() ?? "";
  const yearMatch = yearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr) {
    throw new Error(
      `Metadata missing. Unit code at F12: "${unitCode}", Academic year at C8: "${academicYearStr}". Check cells F12 and C8.`,
    );
  }

  const [unitDoc, academicYearDoc] = await Promise.all([
    Unit.findOne({ code: unitCode }).lean(),
    AcademicYear.findOne({
      year: { $regex: new RegExp(`^${academicYearStr.replace("/", "\\/")}$`, "i") },
      institution: institutionId,
    }).lean(),
  ]);

  if (!unitDoc) throw new Error(`Unit "${unitCode}" not found in the database.`);
  if (!academicYearDoc) {
    throw new Error(`Academic Year "${academicYearStr}" not found for this institution.`);
  }

  const rawRows = xlsx.utils.sheet_to_json<unknown[][]>(sheet, { header: 1, range: 15 });

  const allProgramUnitsForUnit = await ProgramUnit.find({
    unit: (unitDoc as Record<string, unknown>)._id,
  }).lean();

  for (const [index, row] of rawRows.entries()) {
    const rowArr = row as unknown[];
    const rawCell = String(rowArr[1] ?? "").trim().toUpperCase();
    if (!rawCell || rawCell === "REG. NO." || rawCell === "REG NO") continue;
    const regNo = stripQualifier(rawCell);

    result.total++;
    const rowNum = index + 16;

    try {
      const student = await Student.findOne({
        regNo,
        institution: institutionId,
      }).lean();
      if (!student) {
        result.errors.push(`Row ${rowNum} (${regNo}): Student not found.`);
        continue;
      }

      const studentDoc = student as Record<string, unknown>;
      const unitDocObj = unitDoc as Record<string, unknown>;
      const acadYearObj = academicYearDoc as Record<string, unknown>;

      let programUnit = allProgramUnitsForUnit.find(
        (pu) =>
          String((pu as Record<string, unknown>).program) ===
          String(studentDoc.program),
      );

      if (!programUnit) {
        const found = await ProgramUnit.findOne({
          program: studentDoc.program,
          unit: unitDocObj._id,
        }).lean();

        if (!found) {
          result.errors.push(
            `Row ${rowNum} (${regNo}): Unit "${unitCode}" not linked to this student's curriculum.`,
          );
          continue;
        }
        programUnit = found;
      }

      const puObj = programUnit as Record<string, unknown>;

      const attempt = detectAttemptType(rowArr[3]);
      const isSpecial = attempt === "special";
      const isSupp = attempt === "supplementary";
      const isRetake = attempt === "re-take";

      const caTotal30 = Number(rowArr[4]) || 0;
      const examTotal70 = Number(rowArr[5]) || 0;

      const rawAgreed = rowArr[8];
      const agreedMark =
        rawAgreed !== undefined && rawAgreed !== null && rawAgreed !== ""
          ? Number(rawAgreed)
          : caTotal30 + examTotal70;

      const externalRaw = rowArr[7];
      const externalTotal100 =
        externalRaw !== undefined && externalRaw !== null && externalRaw !== ""
          ? Number(externalRaw)
          : null;

      const isMissingCA = caTotal30 === 0 && !isSupp && !isSpecial;

      const markData = {
        institution: institutionId,
        student: studentDoc._id,
        programUnit: puObj._id,
        academicYear: acadYearObj._id,
        batchId,
        caTotal30,
        examTotal70,
        externalTotal100,
        agreedMark,
        attempt,
        isSpecial,
        isSupplementary: isSupp,
        isRetake,
        isMissingCA,
        uploadedBy: req.user._id,
        uploadedAt: new Date(),
        deletedAt: null,
      };

      const saved = await MarkDirect.findOneAndUpdate(
        {
          student: studentDoc._id,
          programUnit: puObj._id,
          academicYear: acadYearObj._id,
        },
        { $set: markData },
        { upsert: true, new: true },
      );

      try {
        await computeFinalGrade({
          markId: saved._id as mongoose.Types.ObjectId,
        });
      } catch (gradeErr: unknown) {
        const msg = gradeErr instanceof Error ? gradeErr.message : String(gradeErr);
        console.warn(
          `[directImporter] Row ${rowNum} (${regNo}): grade calc warning — ${msg}`,
        );
        result.warnings.push(`Row ${rowNum} (${regNo}): grade calc — ${msg}`);
      }

      result.success++;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[directImporter] Row ${rowNum} (${regNo}):`, msg);
      result.errors.push(`Row ${rowNum} (${regNo}): ${msg}`);
    }
  }

  return result;
}