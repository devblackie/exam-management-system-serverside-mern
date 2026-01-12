// // src/services/marksImporter.ts
// import papa from "papaparse";
// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student from "../models/Student";
// // import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// // import Mark from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import { logAudit } from "../lib/auditLogger";
// import type { AuthenticatedRequest } from "../middleware/auth";
// import { MARKS_UPLOAD_HEADERS } from "../utils/uploadTemplate";
// import Mark from "../models/Mark";

// interface ImportRow {
//   RegNo: string;
//   UnitCode: string;
//   CAT1?: string | number;
//   CAT2?: string | number;
//   CAT3?: string | number;
//   Assignment?: string | number;
//   Practical?: string | number;
//   Exam?: string | number;
//   IsSupplementary?: string;
//   AcademicYear: string;
// }

// interface ImportResult {
//   total: number;
//   success: number;
//   errors: string[];
//   warnings: string[];
// }

// export async function importMarksFromBuffer(
//   buffer: Buffer,
//   filename: string,
//   req: AuthenticatedRequest
// ): Promise<ImportResult> {
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = {
//     total: 0,
//     success: 0,
//     errors: [],
//     warnings: [],
//   };

//   // NO SESSION FOR LOOKUPS → prevents auto-abort
//   let students: any[] = [];
//   // let units: any[] = [];
//   let academicYears: any[] = [];

//   try {
//     // 1. Parse file (unchanged)
//     let rows: ImportRow[] = [];
//     if (filename.toLowerCase().endsWith(".csv")) {
//       const text = buffer.toString("utf-8");
//       const parsed = papa.parse<ImportRow>(text, {
//         header: true,
//         skipEmptyLines: true,
//         transformHeader: (h) => h.trim(),
//         transform: (v) => v.trim(),
//       });
//       if (parsed.errors.length)
//         throw new Error(`CSV parse error: ${parsed.errors[0].message}`);
//       rows = parsed.data;
//     } else {
//       const workbook = xlsx.read(buffer, { type: "buffer" });
//       const sheet = workbook.Sheets[workbook.SheetNames[0]];
//       rows = xlsx.utils.sheet_to_json<ImportRow>(sheet, { defval: "" });
//     }

//     if (rows.length === 0) throw new Error("File is empty");

//     // 2. Validate headers
//     const headers = Object.keys(rows[0]);
//     const missing = MARKS_UPLOAD_HEADERS.filter((h) => !headers.includes(h));
//     if (missing.length)
//       throw new Error(`Missing columns: ${missing.join(", ")}`);

//     // 3. PRE-LOAD ALL REFERENCE DATA OUTSIDE ANY TRANSACTION
//     const regNos = rows.map((r) => r.RegNo.toUpperCase());
//     const unitCodes = [...new Set(rows.map((r) => r.UnitCode.toUpperCase()))];
//     const years = [...new Set(rows.map((r) => r.AcademicYear))];

//     // Get all students, all relevant academic years, and ALL relevant ProgramUnits
//     const [studentsData, academicYearsData, programUnits] = await Promise.all([
//       // ⬅️ UPDATED PROMISE.ALL
//       Student.find({
//         regNo: { $in: regNos },
//         institution: institutionId,
//       }).lean(),
//       AcademicYear.find({
//         year: { $in: years },
//         institution: institutionId,
//       }).lean(),
//       ProgramUnit.find({ institution: institutionId }).populate("unit").lean(),
//     ]);

//     students = studentsData;
//     academicYears = academicYearsData;

//     const studentMap = new Map(students.map((s) => [s.regNo.toUpperCase(), s]));
//     // const unitMap = new Map(units.map((u) => [u.code.toUpperCase(), u])); // ⬅️ REMOVED
//     const yearMap = new Map(academicYears.map((y) => [y.year, y]));

//     // 4. NOW start a transaction ONLY for writes
//     const session = await mongoose.startSession();
//     let transactionAborted = false;

//     try {
//       session.startTransaction();

//       // This helps reduce queries inside the loop, though the lookup inside the loop is still complex.
//       const programUnitMap = new Map();
//       programUnits.forEach((pu: any) => {
//         // Key format: PROGRAMID_UNITCODE
//         const key = `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`;
//         programUnitMap.set(key, pu);
//       });

//       for (const [index, row] of rows.entries()) {
//         result.total++;
//         const rowNum = index + 2;

//         try {
//           const regNo = row.RegNo.trim().toUpperCase();
//           const unitCode = row.UnitCode.trim().toUpperCase();
//           const academicYearStr = row.AcademicYear.trim();

//           const student = studentMap.get(regNo);
//           // const unit = unitMap.get(unitCode);
//           const academicYear = yearMap.get(academicYearStr);

//           if (!student) throw new Error(`Student not found: ${regNo}`);
//           // if (!unit) throw new Error(`Unit not found: ${unitCode}`);
//           if (!academicYear)
//             throw new Error(`Academic year not found: ${academicYearStr}`);

//           const programUnitKey = `${student.program.toString()}_${unitCode}`;

//           // Assign the looked-up value to a local variable
//           const programUnit = programUnitMap.get(programUnitKey);

//           if (!programUnit) {
//             throw new Error(
//               `Curriculum link (ProgramUnit) not found for UnitCode ${unitCode} in Student's Program.`
//             );
//           }

//           const markData = {
//             student: student._id,
//             programUnit: programUnit._id,
//             academicYear: academicYear._id,
//             institution: institutionId,
//             uploadedBy: req.user._id,
//             isSupplementary:
//               String(row.IsSupplementary || "NO").toUpperCase() === "YES",
//             cat1: row.CAT1 ? Number(row.CAT1) : undefined,
//             cat2: row.CAT2 ? Number(row.CAT2) : undefined,
//             cat3: row.CAT3 ? Number(row.CAT3) : undefined,
//             assignment: row.Assignment ? Number(row.Assignment) : undefined,
//             practical: row.Practical ? Number(row.Practical) : undefined,
//             exam: row.Exam !== undefined ? Number(row.Exam) : undefined,
//           };

//           const mark = await Mark.findOneAndUpdate(
//             {
//               student: student._id,
//               programUnit: programUnit._id,
//               academicYear: academicYear._id,
//             },
//             markData,
//             { upsert: true, new: true, session }
//           );

//           await computeFinalGrade({
//             markId: mark._id as Types.ObjectId,
//             coordinatorReq: req,
//             session,
//           });

//           result.success++;
//         } catch (rowErr: any) {
//           result.errors.push(`Row ${rowNum}: ${rowErr.message}`);
//           // Keep going – we don’t abort the whole import
//         }
//       }

//       await session.commitTransaction();
//     } catch (transactionErr: any) {
//       transactionAborted = true;
//       await session.abortTransaction();
//       throw transactionErr; // re-throw so caller sees 500
//     } finally {
//       session.endSession();
//     }

//     // 5. Audit log
//     await logAudit(req, {
//       action: "marks_bulk_import_completed",
//       details: {
//         file: filename,
//         totalRows: result.total,
//         successful: result.success,
//         failed: result.errors.length,
//       },
//     });

//     return result;
//   } catch (err: any) {
//     await logAudit(req, {
//       action: "marks_bulk_import_failed",
//       details: { error: err.message, file: filename },
//     });
//     throw err;
//   }
// }

// // serverside/src/services/marksImporter.ts
// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import { logAudit } from "../lib/auditLogger";
// import type { AuthenticatedRequest } from "../middleware/auth";
// import { MARKS_UPLOAD_HEADERS } from "../utils/uploadTemplate";

// interface ImportResult {
//   total: number;
//   success: number;
//   errors: string[];
//   warnings: string[];
// }

// export async function importMarksFromBuffer(
//   buffer: Buffer,
//   filename: string,
//   req: AuthenticatedRequest
// ): Promise<ImportResult> {
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

//   try {
//     const workbook = xlsx.read(buffer, { type: "buffer" });
//     const sheetName = workbook.SheetNames[0];
//     const sheet = workbook.Sheets[sheetName];

//     // --- 1. EXTRACT META DATA FROM HEADER CELLS ---
//     // Unit Code is in I12 (Row 12, Col I)
//     const unitCodeCell = sheet['I12'];
//     const unitCode = unitCodeCell ? unitCodeCell.v.toString().trim().toUpperCase() : null;

//     // Academic Year is in A8 (Extracting the year like 2023/2024 from the string)
//     const academicYearCell = sheet['A8'];
//     const academicYearText = academicYearCell ? academicYearCell.v.toString() : "";
//     // Regex to find "20XX/20XX" pattern
//     const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
//     const academicYearStr = yearMatch ? yearMatch[0] : null;

//     if (!unitCode || !academicYearStr) {
//       throw new Error("Could not find Unit Code (I12) or Academic Year (A8) in the Excel header.");
//     }

//     // --- 2. VALIDATE HEADERS & PARSE ROWS ---
//     // Read starting from Row 15 (index 14) where headers are
//     const rows = xlsx.utils.sheet_to_json<any>(sheet, { range: 14, defval: "" });
//     if (rows.length === 0) throw new Error("No student data found in the file.");

//     const fileHeaders = Object.keys(rows[0]).map(h => h.trim());
//     const required = ["REG. NO.", "CAT 1 Out of", "AGREED MARKS /100"];
//    const missing = required.filter(h => !fileHeaders.includes(h));

//     if (missing.length > 0) {
//        throw new Error(`Invalid template. Missing required columns: ${missing.join(", ")}`);
//     }

//     // --- 3. PRE-LOAD REFERENCE DATA ---
//     const regNos = rows.map(r => r["REG. NO."]?.toString().trim().toUpperCase()).filter(Boolean);

//     const [students, academicYearDoc, programUnits] = await Promise.all([
//       Student.find({ regNo: { $in: regNos }, institution: institutionId }).lean(),
//       AcademicYear.findOne({ year: academicYearStr, institution: institutionId }).lean(),
//       ProgramUnit.find({ institution: institutionId }).populate("unit").lean(),
//     ]);

//     if (!academicYearDoc) throw new Error(`Academic Year '${academicYearStr}' not found in database.`);

//     const studentMap = new Map(students.map(s => [s.regNo.toUpperCase(), s]));
//     const programUnitMap = new Map();
//     programUnits.forEach((pu: any) => {
//       const key = `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`;
//       programUnitMap.set(key, pu);
//     });

//     // --- 4. DATA PROCESSING ---
//     const session = await mongoose.startSession();
//     try {
//       session.startTransaction();

//       for (const [index, row] of rows.entries()) {
//         const regNo = row["REG. NO."]?.toString().trim().toUpperCase();
//         if (!regNo) continue;

//         result.total++;
//         const rowNum = index + 17; // Actual Excel row number

//         try {
//           const student: any = studentMap.get(regNo);
//           if (!student) throw new Error(`Student ${regNo} not found in system.`);

//           const programUnitKey = `${student.program.toString()}_${unitCode}`;
//           const programUnit = programUnitMap.get(programUnitKey);

//           if (!programUnit) throw new Error(`Unit ${unitCode} is not linked to student's program.`);

//           // Map Excel columns to Database fields
//           const markData = {
// student: student._id,
//           programUnit: programUnit._id,
//           academicYear: academicYearDoc._id,
//           institution: institutionId,
//           uploadedBy: req.user._id,
//           // ACCESSING EXACT KEYS FROM CONSTANT
//           cat1: row["CAT 1 Out of"] !== "" ? Number(row["CAT 1 Out of"]) : undefined,
//           cat2: row["CAT 2 Out of"] !== "" ? Number(row["CAT 2 Out of"]) : undefined,
//           cat3: row["CAT3 Out of"] !== "" ? Number(row["CAT3 Out of"]) : undefined,
//           assignment: row["Assgnt 1 Out of"] !== "" ? Number(row["Assgnt 1 Out of"]) : undefined,
//           exam: row["AGREED MARKS /100"] !== "" ? Number(row["AGREED MARKS /100"]) : undefined,
//   isSupplementary: String(row["ATTEMPT"]).toUpperCase().includes("SUPP"),
//           };

//           const mark = await Mark.findOneAndUpdate(
//             { student: student._id, programUnit: programUnit._id, academicYear: academicYearDoc._id },
//             markData,
//             { upsert: true, new: true, session }
//           );

//           await computeFinalGrade({
//             markId: mark._id as Types.ObjectId,
//             coordinatorReq: req,
//             session,
//           });

//           result.success++;
//         } catch (rowErr: any) {
//           result.errors.push(`Row ${rowNum}: ${rowErr.message}`);
//         }
//       }
//       await session.commitTransaction();
//     } catch (error) {
//       await session.abortTransaction();
//       throw error;
//     } finally {
//       session.endSession();
//     }

//     await logAudit(req, {
//       action: "marks_bulk_import_completed",
//       details: { file: filename, total: result.total, success: result.success }
//     });

//     return result;
//   } catch (err: any) {
//     throw err;
//   }
// }

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
    const session = await mongoose.startSession();
    try {
      session.startTransaction();

      for (const [index, row] of rows.entries()) {
        const regNo = row["REG. NO."]?.toString().trim().toUpperCase();
        if (!regNo) continue;

        result.total++;
        const rowNum = index + 17;

        try {
          const student: any = studentMap.get(regNo);
          if (!student)
            throw new Error(`Student ${regNo} not found in system.`);

          const programUnitKey = `${student.program.toString()}_${unitCode}`;
          const programUnit = programUnitMap.get(programUnitKey);

          if (!programUnit)
            throw new Error(
              `Unit ${unitCode} is not linked to student's program.`
            );

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
  examQ1Raw: row["Q1 out of"] !== "" ? Number(row["Q1 out of"]) : undefined,
  examQ2Raw: row["Q2 out of"] !== "" ? Number(row["Q2 out of"]) : undefined,
  examQ3Raw: row["Q3 out of"] !== "" ? Number(row["Q3 out of"]) : undefined,
  examQ4Raw: row["Q4 out of"] !== "" ? Number(row["Q4 out of"]) : undefined,
  examQ5Raw: row["Q5 out of"] !== "" ? Number(row["Q5 out of"]) : undefined,


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
            { upsert: true, new: true, session }
          );

          await computeFinalGrade({
            markId: mark._id as Types.ObjectId,
            coordinatorReq: req,
            session,
          });

          result.success++;
        } catch (rowErr: any) {
          result.errors.push(`Row ${rowNum}: ${rowErr.message}`);
        }
      }
      await session.commitTransaction();
    } catch (error) {
      await session.abortTransaction();
      throw error;
    } finally {
      session.endSession();
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
