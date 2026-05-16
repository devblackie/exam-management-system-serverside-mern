// serverside/src/services/marksImporter.ts — COMPLETE, PRODUCTION READY

import { randomUUID } from "node:crypto";
import xlsx from "xlsx";
import mongoose, { Types } from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Mark from "../models/Mark";
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

export async function importMarksFromBuffer(
  buffer: Buffer,
  filename: string,
  req: AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const batchId = randomUUID();
  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const modeIndicator = (sheet["E10"]?.v?.toString() ?? "").toUpperCase();
  let detectedUnitType: "theory" | "lab" | "workshop" = "theory";
  if (modeIndicator.includes("LAB")) detectedUnitType = "lab";
  if (modeIndicator.includes("WORKSHOP")) detectedUnitType = "workshop";

  const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
  const academicYearText = sheet["D8"]?.v?.toString() ?? "";
  const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr) {
    throw new Error(
      `Invalid Template: Missing Unit Code (H12) or Academic Year (D8). Found Unit: ${unitCode}, Year: ${academicYearStr}`,
    );
  }

  const rawRows = xlsx.utils.sheet_to_json<unknown[][]>(sheet, { header: 1, range: 16 });

  const academicYearDoc = await AcademicYear.findOne({
    year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
    institution: institutionId,
  }).lean();

  if (!academicYearDoc) {
    throw new Error(`Academic Year '${academicYearStr}' not found.`);
  }

  const acadYearObj = academicYearDoc as Record<string, unknown>;

  const allProgramUnits = await ProgramUnit.find({ institution: institutionId })
    .populate("unit")
    .lean();

  const programUnitMap = new Map<string, Record<string, unknown>>(
    allProgramUnits.map((pu) => {
      const puObj = pu as Record<string, unknown>;
      const unitObj = puObj.unit as Record<string, unknown> | undefined;
      return [
        `${String(puObj.program)}_${String(unitObj?.code ?? "").toUpperCase()}`,
        puObj,
      ];
    }),
  );

  for (const [index, row] of rawRows.entries()) {
    const rowArr = row as unknown[];
    const rawCell = String(rowArr[1] ?? "").trim().toUpperCase();
    const sn = rowArr[0];
    if (!rawCell || sn === "") continue;

    const regNo = stripQualifier(rawCell);
    result.total++;
    const rowNum = index + 17;

    try {
      const student = await Student.findOne({
        regNo,
        institution: institutionId,
      }).lean();

      if (!student) {
        result.errors.push(`Row ${rowNum} (${regNo}): Student not found.`);
        continue;
      }

      const studentObj = student as Record<string, unknown>;

      const programUnitKey = `${String(studentObj.program)}_${unitCode}`;
      const programUnit = programUnitMap.get(programUnitKey);

      if (!programUnit) {
        result.errors.push(
          `Row ${rowNum} (${regNo}): Unit "${unitCode}" not linked to student's program.`,
        );
        continue;
      }

      const excelAttemptLabel = String(rowArr[3] ?? "").trim().toLowerCase();
      let finalAttempt = "1st";
      if (excelAttemptLabel.includes("supp")) {
        finalAttempt = "supplementary";
      } else if (excelAttemptLabel.includes("special")) {
        finalAttempt = "special";
      } else if (excelAttemptLabel.includes("retake")) {
        finalAttempt = "re-take";
      }

      const markData = {
        student: studentObj._id,
        programUnit: programUnit._id,
        academicYear: acadYearObj._id,
        institution: institutionId,
        uploadedBy: req.user._id,
        deletedAt: null,
        batchId,

        cat1Raw: Number(rowArr[4]) || 0,
        cat2Raw: Number(rowArr[5]) || 0,
        cat3Raw: Number(rowArr[6]) || 0,
        assgnt1Raw: Number(rowArr[8]) || 0,
        assgnt2Raw: Number(rowArr[9]) || 0,
        assgnt3Raw: Number(rowArr[10]) || 0,
        practicalRaw: Number(rowArr[12]) || 0,
        examQ1Raw: Number(rowArr[14]) || 0,
        examQ2Raw: Number(rowArr[15]) || 0,
        examQ3Raw: Number(rowArr[16]) || 0,
        examQ4Raw: Number(rowArr[17]) || 0,
        examQ5Raw: Number(rowArr[18]) || 0,
        caTotal30: Number(rowArr[13]) || 0,
        examTotal70: Number(rowArr[19]) || 0,
        internalExaminerMark: Number(rowArr[20]) || 0,
        agreedMark: Number(rowArr[22]) || 0,

        attempt: finalAttempt,
        isSpecial: finalAttempt === "special",
        isSupplementary: finalAttempt === "supplementary",
        isRetake: finalAttempt === "re-take",
        unitType: detectedUnitType,
        examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
      };

      const mark = await Mark.findOneAndUpdate(
        {
          student: studentObj._id,
          programUnit: programUnit._id,
          academicYear: acadYearObj._id,
        },
        markData,
        { upsert: true, new: true, runValidators: true },
      );

      await computeFinalGrade({ markId: mark._id as Types.ObjectId });

      result.success++;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[marksImporter] Row ${rowNum} (${regNo}):`, msg);
      result.errors.push(`Row ${rowNum} (${regNo}): ${msg}`);
    }
  }

  return result;
}