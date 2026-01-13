// // src/services/marksImporter.ts
import xlsx from "xlsx";
import mongoose, { Types } from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Mark from "../models/Mark";
import { computeFinalGrade } from "./gradeCalculator";
import { logAudit } from "../lib/auditLogger";
import type { AuthenticatedRequest } from "../middleware/auth";
import { MARKS_UPLOAD_HEADERS } from "../utils/uploadTemplate";

interface ImportResult {
  total: number;
  success: number;
  errors: string[];
  warnings: string[];
}

export async function importMarksFromBuffer(
  buffer: Buffer,
  filename: string,
  req: AuthenticatedRequest
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const result: ImportResult = {
    total: 0,
    success: 0,
    errors: [],
    warnings: [],
  };

  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // --- 1. EXTRACT META DATA FROM HEADER CELLS ---
    const unitCodeCell = sheet["I12"];
    const unitCode = unitCodeCell
      ? unitCodeCell.v.toString().trim().toUpperCase()
      : null;

    const academicYearCell = sheet["A8"];
    const academicYearText = academicYearCell
      ? academicYearCell.v.toString()
      : "";
    const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
    const academicYearStr = yearMatch ? yearMatch[0] : null;

    if (!unitCode || !academicYearStr) {
      throw new Error(
        "Could not find Unit Code (I12) or Academic Year (A8) in the Excel header."
      );
    }

    // --- 2. VALIDATE HEADERS & PARSE ROWS ---
    const rows = xlsx.utils.sheet_to_json<any>(sheet, {
      range: 14,
      defval: "",
    });
    if (rows.length === 0)
      throw new Error("No student data found in the file.");

    const fileHeaders = Object.keys(rows[0]).map((h) => h.trim());

    // Updated required columns list as requested
    const required = [
      "REG. NO.",
      "CAT 1 Out of",
      "AGREED MARKS /100",
      "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
      "TOTAL CATS",
      "TOTAL ASSGNT",
      "TOTAL EXAM OUT OF",
    ];

    const missing = required.filter((h) => !fileHeaders.includes(h));

    if (missing.length > 0) {
      throw new Error(
        `Invalid template. Missing required columns: ${missing.join(", ")}`
      );
    }

    // --- 3. PRE-LOAD REFERENCE DATA ---
    const regNos = rows
      .map((r) => r["REG. NO."]?.toString().trim().toUpperCase())
      .filter(Boolean);

    const [students, academicYearDoc, programUnits] = await Promise.all([
      Student.find({
        regNo: { $in: regNos },
        institution: institutionId,
      }).lean(),
      AcademicYear.findOne({
        year: academicYearStr,
        institution: institutionId,
      }).lean(),
      ProgramUnit.find({ institution: institutionId }).populate("unit").lean(),
    ]);

    if (!academicYearDoc)
      throw new Error(
        `Academic Year '${academicYearStr}' not found in database.`
      );

    const studentMap = new Map(students.map((s) => [s.regNo.toUpperCase(), s]));
    const programUnitMap = new Map();
    programUnits.forEach((pu: any) => {
      const key = `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`;
      programUnitMap.set(key, pu);
    });

    // --- 4. DATA PROCESSING ---
    // We remove the global session and handle it per student or in small batches
    for (const [index, row] of rows.entries()) {
      const regNo = row["REG. NO."]?.toString().trim().toUpperCase();
      if (!regNo) continue;

      result.total++;
      const rowNum = index + 17;

      // Start a fresh session for each student to prevent timeouts
      const session = await mongoose.startSession();
      try {
        await session.withTransaction(async () => {
          const student: any = studentMap.get(regNo);
          if (!student)
            throw new Error(`Student ${regNo} not found in system.`);

          const programUnitKey = `${student.program.toString()}_${unitCode}`;
          const programUnit = programUnitMap.get(programUnitKey);

          if (!programUnit)
            throw new Error(`Unit ${unitCode} not linked to program.`);

          // Map Excel columns to Database fields
          const markData = {
            student: student._id,
            programUnit: programUnit._id,
            academicYear: academicYearDoc._id,
            institution: institutionId,
            uploadedBy: req.user._id,

            // --- RAW SCORES (Optional fields can be undefined, defaults are in Schema) ---
            cat1Raw:
              row["CAT 1 Out of"] !== "" ? Number(row["CAT 1 Out of"]) : 0,
            cat2Raw:
              row["CAT 2 Out of"] !== "" ? Number(row["CAT 2 Out of"]) : 0,
            cat3Raw:
              row["CAT3 Out of"] !== ""
                ? Number(row["CAT3 Out of"])
                : undefined,
            assgnt1Raw:
              row["Assgnt 1 Out of"] !== ""
                ? Number(row["Assgnt 1 Out of"])
                : 0,

            // --- ADD THESE EXAM QUESTION MAPPINGS ---
            // Note: Ensure the string keys (e.g., "Q1 /10") match your Excel column headers exactly
            examQ1Raw:
              row["Q1 out of"] !== "" ? Number(row["Q1 out of"]) : undefined,
            examQ2Raw:
              row["Q2 out of"] !== "" ? Number(row["Q2 out of"]) : undefined,
            examQ3Raw:
              row["Q3 out of"] !== "" ? Number(row["Q3 out of"]) : undefined,
            examQ4Raw:
              row["Q4 out of"] !== "" ? Number(row["Q4 out of"]) : undefined,
            examQ5Raw:
              row["Q5 out of"] !== "" ? Number(row["Q5 out of"]) : undefined,

            // --- FINAL AUDIT FIELDS (Required by Schema - MUST NOT BE NaN) ---
            // Using '|| 0' ensures we never send NaN to a required Number field
            caTotal30:
              Number(row["CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30"]) ||
              0,
            examTotal70: Number(row["TOTAL EXAM OUT OF"]) || 0,
            internalExaminerMark:
              Number(row["INTERNAL EXAMINER MARKS /100"]) || 0,
            agreedMark: Number(row["AGREED MARKS /100"]) || 0,

            attempt: String(row["ATTEMPT"]).toLowerCase().includes("supp")
              ? "supplementary"
              : String(row["ATTEMPT"]).toLowerCase().includes("re-take")
              ? "re-take"
              : "1st",

            isSupplementary: String(row["ATTEMPT"])
              .toLowerCase()
              .includes("supp"),
            isRetake: String(row["ATTEMPT"]).toLowerCase().includes("re-take"),
          };

          const mark = await Mark.findOneAndUpdate(
            {
              student: student._id,
              programUnit: programUnit._id,
              academicYear: academicYearDoc._id,
            },
            markData,
            { upsert: true, new: true, session, runValidators: true }
          );

          // This is the heavy part - running it inside the student-level session
          await computeFinalGrade({
            markId: mark._id as Types.ObjectId,
            coordinatorReq: req,
            session,
          });
        });

        result.success++;
      } catch (rowErr: any) {
        // We log the error but the loop CONTINUES to the next student
        result.errors.push(`Row ${rowNum}: ${rowErr.message}`);
        console.error(`Import error at row ${rowNum}:`, rowErr.message);
      } finally {
        await session.endSession();
      }
    }

    await logAudit(req, {
      action: "marks_bulk_import_completed",
      details: { file: filename, total: result.total, success: result.success },
    });

    return result;
  } catch (err: any) {
    throw err;
  }
}
