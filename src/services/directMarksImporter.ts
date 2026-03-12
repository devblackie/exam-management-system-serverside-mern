// src/services/directMarksImporter.ts

import xlsx from "xlsx";
import mongoose from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Unit from "../models/Unit"; // Added Unit import
import MarkDirect from "../models/MarkDirect";
import type { AuthenticatedRequest } from "../middleware/auth";

interface ImportResult {
  total: number;
  success: number;
  errors: string[];
  warnings: string[];
}

export async function importDirectMarksFromBuffer(
  buffer: Buffer,
  filename: string,
  req: AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const result: ImportResult = {
    total: 0,
    success: 0,
    errors: [],
    warnings: [],
  };

  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  // 1. Meta Data Extraction
  const unitCode = sheet["F12"]?.v?.toString().trim().toUpperCase();
  const academicYearText = sheet["C8"]?.v?.toString() || "";
  const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr)
    throw new Error("Template metadata missing (Unit Code or Year).");

  // 2. Resolve Global Records (Unit and Academic Year)
  const [unitDoc, academicYearDoc] = await Promise.all([
    Unit.findOne({ code: unitCode }).lean(),
    AcademicYear.findOne({
      year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
      institution: institutionId,
    }).lean(),
  ]);

  if (!unitDoc)
    throw new Error(
      `Unit with code '${unitCode}' does not exist in the database.`,
    );
  if (!academicYearDoc)
    throw new Error(`Academic Year '${academicYearStr}' not found.`);

  // 3. Parse Rows
  const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, {
    header: 1,
    range: 15,
  });

  for (const [index, row] of rawRows.entries()) {
    const regNo = row[1]?.toString().trim().toUpperCase();
    if (!regNo) continue;

    result.total++;
    const session = await mongoose.startSession();
    try {
      await session.withTransaction(async () => {
        const student = await Student.findOne({
          regNo,
          institution: institutionId,
        }).lean();
        if (!student) throw new Error(`Student ${regNo} not found`);

        // 4. Robust ProgramUnit Lookup
        // We look for a record that links the Student's Program to the Unit from the Excel
        const programUnit = await ProgramUnit.findOne({
          institution: institutionId,
          program: student.program, // This is the ID from the error: 693534f37ea0366e6b831034
          unit: unitDoc._id,
        }).lean();

        if (!programUnit) {
          throw new Error(
            `Unit ${unitCode} is not registered in this student's curriculum (Program ID: ${student.program}).`,
          );
        }

        const markData = {
          institution: institutionId,
          student: student._id,
          programUnit: programUnit._id,
          academicYear: academicYearDoc._id,
          semester: academicYearText.toUpperCase().includes("SEMESTER 2")
            ? "SEMESTER 2"
            : "SEMESTER 1",
          caTotal30: Number(row[4]) || 0,
          examTotal70: Number(row[5]) || 0,
          externalTotal100: row[7] !== undefined && row[7] !== null ? Number(row[7]) : null,
          agreedMark: Number(row[8]) || 0,
          attempt: row[3]?.toString().toLowerCase().includes("supp")
            ? "supplementary"
            : "1st",
          uploadedBy: req.user._id,
        };

        await MarkDirect.findOneAndUpdate(
          {
            student: student._id,
            programUnit: programUnit._id,
            academicYear: academicYearDoc._id,
          },
          markData,
          { upsert: true, session },
        );
      });
      result.success++;
    } catch (err: any) {
      result.errors.push(`Row ${index + 16} (${regNo}): ${err.message}`);
    } finally {
      await session.endSession();
    }
  }
  return result;
}
