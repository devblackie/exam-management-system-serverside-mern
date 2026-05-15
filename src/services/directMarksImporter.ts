
// serverside/src/services/directMarksImporter.ts

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

  // ── Generate ONE batch ID for this entire upload ───────────────────────
  const batchId = randomUUID();

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };
  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const unitCodeRaw = sheet["F12"]?.v;
  const yearTextRaw = sheet["C8"]?.v;
  const unitCode = unitCodeRaw?.toString().trim().toUpperCase();
  const yearText = yearTextRaw?.toString() || "";
  const yearMatch = yearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr) {
    throw new Error(
      `Metadata missing. Unit code at F12: "${unitCode}", Academic year at C8: "${academicYearStr}". Check cells F12 and C8.`,
    );
  }

  const unitDoc = await Unit.findOne({ code: unitCode }).lean();
  if (!unitDoc) throw new Error(`Unit "${unitCode}" not found in the database.`);

  const academicYearDoc = await AcademicYear.findOne({
    year: { $regex: new RegExp(`^${academicYearStr.replace("/", "\\/")}$`, "i") },
    institution: institutionId,
  }).lean();
  if (!academicYearDoc)
    throw new Error(`Academic Year "${academicYearStr}" not found for this institution.`);

  const rawRows = xlsx.utils.sheet_to_json<unknown[][]>(sheet, { header: 1, range: 15 });

  const allProgramUnitsForUnit = await ProgramUnit.find({
    unit: (unitDoc as Record<string, unknown>)._id,
  }).lean();

  for (const [index, row] of rawRows.entries()) {
    const rawCell = String(row[1] ?? "").trim().toUpperCase();
    if (!rawCell || rawCell === "REG. NO." || rawCell === "REG NO") continue;
    const regNo = stripQualifier(rawCell);

    result.total++;
    const rowNum = index + 16;

    const session = await mongoose.startSession();
    try {
      let savedMarkId: mongoose.Types.ObjectId | null = null;

      await session.withTransaction(async () => {
        const student = await Student.findOne({
          regNo,
          institution: institutionId,
        }).lean();
        if (!student)
          throw new Error(`Student "${regNo}" not found for this institution`);

        // First try the pre-fetched list
        let programUnit = allProgramUnitsForUnit.find(
          (pu) =>
            String((pu as Record<string, unknown>).program) ===
            String((student as Record<string, unknown>).program),
        );

        // Fallback: query directly if not found in pre-fetched list
        if (!programUnit) {
          const found = await ProgramUnit.findOne({
            program: (student as Record<string, unknown>).program,
            unit: (unitDoc as Record<string, unknown>)._id,
          }).lean();

          if (!found) {
            throw new Error(
              `Unit "${unitCode}" is not linked to the curriculum for student "${regNo}"`,
            );
          }
          programUnit = found;
        }

        const savedMark = await upsertMarkDirect(
          programUnit,
          student as Record<string, unknown>,
          row,
          rowNum,
          institutionId,
          academicYearDoc as Record<string, unknown>,
          req,
          session,
          batchId,
        );

        savedMarkId = savedMark._id as mongoose.Types.ObjectId;
      });

      if (savedMarkId) {
        try {
          await computeFinalGrade({ markId: savedMarkId });
        } catch (gradeErr: unknown) {
          const msg =
            gradeErr instanceof Error
              ? gradeErr.message
              : "Unknown grade error";
          console.warn(
            `[directImporter] Row ${rowNum}: FinalGrade calc failed for ${regNo}: ${msg}`,
          );
          result.warnings.push(`Row ${rowNum} (${regNo}): grade calc — ${msg}`);
        }
      }

      result.success++;
    } catch (err: unknown) {
      const msg = `Row ${rowNum} (${regNo}): ${err instanceof Error ? err.message : "Unknown error"}`;
      console.error(`[directImporter] FAILED — ${msg}`);
      result.errors.push(msg);
    } finally {
      await session.endSession();
    }
  }

  return result;
}

// ─── Helper: upsert the MarkDirect document ───────────────────────────────────

async function upsertMarkDirect(
  programUnit: Record<string, unknown>,
  student: Record<string, unknown>,
  row: unknown[],
  rowNum: number,
  institutionId: unknown,
  academicYearDoc: Record<string, unknown>,
  req: AuthenticatedRequest,
  session: mongoose.ClientSession,
  batchId: string,
): Promise<Record<string, unknown>> {
  const attempt = detectAttemptType(row[3]);
  const isSpecial = attempt === "special";
  const isSupp = attempt === "supplementary";
  const isRetake = attempt === "re-take";

  const caTotal30 = Number(row[4]) || 0;
  const examTotal70 = Number(row[5]) || 0;

  const rawAgreed = row[8];
  const agreedMark =
    rawAgreed !== undefined && rawAgreed !== null && rawAgreed !== ""
      ? Number(rawAgreed)
      : caTotal30 + examTotal70;

  const externalRaw = row[7];
  const externalTotal100 =
    externalRaw !== undefined && externalRaw !== null && externalRaw !== ""
      ? Number(externalRaw)
      : null;

  const isMissingCA = caTotal30 === 0 && !isSupp && !isSpecial;

  const markData = {
    institution: institutionId,
    student: student._id,
    programUnit: programUnit._id,
    academicYear: academicYearDoc._id,
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
      student: student._id,
      programUnit: programUnit._id,
      academicYear: academicYearDoc._id,
    },
    { $set: markData },
    { upsert: true, new: true, session },
  );

  return saved as unknown as Record<string, unknown>;
}