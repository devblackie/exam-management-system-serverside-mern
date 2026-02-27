// src/services/marksImporter.ts
// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import type { AuthenticatedRequest } from "../middleware/auth";

// interface ImportResult {
//   total: number;
//   success: number;
//   errors: string[]; warnings: string[];}

// /**
//  * Senior Engineer Note:
//  * This importer strictly follows the column mapping defined in uploadTemplate.ts.
//  * Column indices are 0-based (A=0, B=1, etc.)
//  */
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
//     const sheet = workbook.Sheets[workbook.SheetNames[0]];

//     const modeIndicator = sheet["E10"]?.v?.toString().toUpperCase() || "";
//     let detectedUnitType: "theory" | "lab" | "workshop" = "theory";

//     if (modeIndicator.includes("LAB")) detectedUnitType = "lab";
//     if (modeIndicator.includes("WORKSHOP")) detectedUnitType = "workshop";

//     // 1. Meta Data Extraction
//     // Unit Code: H12, Academic Year: F8 (as merged cells)
//     const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
//     const academicYearText = sheet["D8"]?.v?.toString() || "";

//     // console.log(`[Importer] Checking E8: "${academicYearText}", H12: "${unitCode}"`);

//     const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
//     const academicYearStr = yearMatch ? yearMatch[0] : null;

//     if (!unitCode || !academicYearStr) {
//       throw new Error(
//         `Invalid Template: Missing Unit Code (H12) or Academic Year (D8). Found Unit: ${unitCode}, Year: ${academicYearStr}`,
//       );
//     }

//     // 2. Parse Rows starting from Row 17 (range: 16)
//     const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 16 });

//     // 3. Pre-fetch shared data
//     const academicYearDoc = await AcademicYear.findOne({
//       year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
//       institution: institutionId
//     }).lean();

//     if (!academicYearDoc) throw new Error(`Academic Year '${academicYearStr}' not found.`);

//     const programUnits = await ProgramUnit.find({ institution: institutionId }).populate("unit").lean();
//     const programUnitMap = new Map(programUnits.map((pu: any) => [`${pu.program.toString()}_${pu.unit.code.toUpperCase()}`, pu]));

//     // 4. Processing Mapping (A=0, B=1, C=2...)
//     for (const [index, row] of rawRows.entries()) {
//       const regNo = row[1]?.toString().trim().toUpperCase(); // Col B
//       const sn = row[0]; // Col A

//       if (!regNo || sn === "") continue;

//       result.total++;
//       const rowNum = index + 17;
//       const session = await mongoose.startSession();

//       try {
//         await session.withTransaction(async () => {
//           const student = await Student.findOne({ regNo, institution: institutionId }).lean();
//           if (!student) throw new Error(`Student ${regNo} not found.`);

//           const programUnitKey = `${student.program.toString()}_${unitCode}`;
//           const programUnit = programUnitMap.get(programUnitKey);
//           if (!programUnit) throw new Error(`Unit ${unitCode} not linked to program.`);

//           // --- SMART ATTEMPT DETECTION ---
//           const previousMarks = await Mark.find({
//             student: student._id,
//             programUnit: programUnit._id
//           }).session(session).sort({ createdAt: 1 }).lean();

//           const excelAttemptLabel = row[3]?.toString().trim(); // Column D
//           let finalAttempt = "1st";

//           // Priority 1: Check Excel Label
//           if (excelAttemptLabel?.toLowerCase().includes("supp")) {
//             finalAttempt = "supplementary";
//           } else if (excelAttemptLabel?.toLowerCase().includes("special")) {
//             finalAttempt = "special";
//           }
//           // Priority 2: Auto-detect from history if Excel is generic "1st" or empty
//           else if (previousMarks.length > 0) {
//             const historyCount = previousMarks.length;
//             if (historyCount === 1) finalAttempt = "supplementary";
//             else if (historyCount === 2) finalAttempt = "re-take";
//             else if (historyCount >= 3) finalAttempt = "re-retake";
//           }

//           const markData = {
//             student: student._id,
//             programUnit: programUnit._id,
//             academicYear: academicYearDoc._id,
//             institution: institutionId,
//             uploadedBy: req.user._id,

//             // CA Scores
//             cat1Raw: Number(row[4]) || 0, // Col E
//             cat2Raw: Number(row[5]) || 0, // Col F
//             cat3Raw: Number(row[6]) || 0, // Col G

//             // Assignments
//             assgnt1Raw: Number(row[8]) || 0, // Col I
//             assgnt2Raw: Number(row[9]) || 0, // Col J
//             assgnt3Raw: Number(row[10]) || 0, // Col K

//             // Practical (NEW COLUMN)
//             practicalRaw: Number(row[12]) || 0, // Col M

//             // Exam Questions
//             examQ1Raw: Number(row[14]) || 0, // Col O
//             examQ2Raw: Number(row[15]) || 0, // Col P
//             examQ3Raw: Number(row[16]) || 0, // Col Q
//             examQ4Raw: Number(row[17]) || 0, // Col R
//             examQ5Raw: Number(row[18]) || 0, // Col S

//             // Totals and Final Fields
//             caTotal30: Number(row[13]) || 0, // Col N (CA GRAND TOTAL)
//             examTotal70: Number(row[19]) || 0, // Col T (TOTAL EXAM)
//             internalExaminerMark: Number(row[20]) || Number(row[22]), // Col U (INT. EXAMINER) or fallback to Col W (AGREED MARK) if U is empty
//             agreedMark: Number(row[22]) || 0, // Col W (AGREED MARKS)

//             attempt: finalAttempt,
//             isSpecial: finalAttempt === "special",
//             isSupplementary: finalAttempt === "supplementary",
//             isRetake: finalAttempt.includes("re-take"),

//             unitType: detectedUnitType,
//             examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
//           };

//           const mark = await Mark.findOneAndUpdate(
//             { student: student._id, programUnit: programUnit._id, academicYear: academicYearDoc._id },
//             markData,
//             { upsert: true, new: true, session }
//           );

//           // Trigger the final grade calculation using the updated logic
//           await computeFinalGrade({ markId: mark._id as Types.ObjectId, session });
//         });
//         result.success++;
//       } catch (rowErr: any) {
//         result.errors.push(`Row ${rowNum} (${regNo}): ${rowErr.message}`);
//       } finally {
//         await session.endSession();
//       }
//     }

//     return result;
//   } catch (err: any) {
//     console.error(`[Importer] Fatal Error:`, err.message);
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
import type { AuthenticatedRequest } from "../middleware/auth";

interface ImportResult {
  total: number;
  success: number;
  errors: string[];
  warnings: string[];
}

export async function importMarksFromBuffer(
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

  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const modeIndicator = sheet["E10"]?.v?.toString().toUpperCase() || "";
    let detectedUnitType: "theory" | "lab" | "workshop" = "theory";

    if (modeIndicator.includes("LAB")) detectedUnitType = "lab";
    if (modeIndicator.includes("WORKSHOP")) detectedUnitType = "workshop";

    // 1. Meta Data Extraction
    const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
    const academicYearText = sheet["D8"]?.v?.toString() || "";

    const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
    const academicYearStr = yearMatch ? yearMatch[0] : null;

    if (!unitCode || !academicYearStr) {
      throw new Error(
        `Invalid Template: Missing Unit Code (H12) or Academic Year (D8). Found Unit: ${unitCode}, Year: ${academicYearStr}`,
      );
    }

    // 2. Parse Rows
    const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, {
      header: 1,
      range: 16,
    });

    // 3. Pre-fetch shared data
    const academicYearDoc = await AcademicYear.findOne({
      year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
      institution: institutionId,
    }).lean();

    if (!academicYearDoc)
      throw new Error(`Academic Year '${academicYearStr}' not found.`);

    // Optimization: Pre-fetch all units for the institution
    const programUnits = await ProgramUnit.find({ institution: institutionId })
      .populate("unit")
      .lean();
    const programUnitMap = new Map(
      programUnits.map((pu: any) => [
        `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`,
        pu,
      ]),
    );

    // 4. Processing Mapping (A=0, B=1, C=2...)
    for (const [index, row] of rawRows.entries()) {
      const regNo = row[1]?.toString().trim().toUpperCase(); // Col B
      const sn = row[0]; // Col A

      if (!regNo || sn === "") continue;

      result.total++;
      const rowNum = index + 17;
      const session = await mongoose.startSession();

      try {
        await session.withTransaction(async () => {
          const student = await Student.findOne({
            regNo,
            institution: institutionId,
          }).lean();
          if (!student) throw new Error(`Student ${regNo} not found.`);

          const programUnitKey = `${student.program.toString()}_${unitCode}`;
          const programUnit = programUnitMap.get(programUnitKey);
          if (!programUnit)
            throw new Error(
              `Unit ${unitCode} not linked to student's program.`,
            );

          // --- FIX: SMART ATTEMPT DETECTION ---
          // Removing the pre-find (Mark.find) lookup as it might be causing issues
          // with soft-delete middleware if a previous record exists but is deleted.

          const excelAttemptLabel = row[3]?.toString().trim(); // Column D
          let finalAttempt = "1st";

          if (excelAttemptLabel?.toLowerCase().includes("supp")) {
            finalAttempt = "supplementary";
          } else if (excelAttemptLabel?.toLowerCase().includes("special")) {
            finalAttempt = "special";
          } else if (excelAttemptLabel?.toLowerCase().includes("retake")) {
            finalAttempt = "re-take";
          }

          // --- FIX: Mapping Columns and Ensuring Numbers ---
          const markData = {
            student: student._id,
            programUnit: programUnit._id,
            academicYear: academicYearDoc._id,
            institution: institutionId,
            uploadedBy: req.user._id,
            deletedAt: null, // Ensure we are updating active records

            // CA Scores
            cat1Raw: Number(row[4]) || 0, // Col E
            cat2Raw: Number(row[5]) || 0, // Col F
            cat3Raw: Number(row[6]) || 0, // Col G

            // Assignments
            assgnt1Raw: Number(row[8]) || 0, // Col I
            assgnt2Raw: Number(row[9]) || 0, // Col J
            assgnt3Raw: Number(row[10]) || 0, // Col K

            // Practical
            practicalRaw: Number(row[12]) || 0, // Col M

            // Exam Questions
            examQ1Raw: Number(row[14]) || 0, // Col O
            examQ2Raw: Number(row[15]) || 0, // Col P
            examQ3Raw: Number(row[16]) || 0, // Col Q
            examQ4Raw: Number(row[17]) || 0, // Col R
            examQ5Raw: Number(row[18]) || 0, // Col S

            // Totals and Final Fields
            caTotal30: Number(row[13]) || 0, // Col N (CA GRAND TOTAL)
            examTotal70: Number(row[19]) || 0, // Col T (TOTAL EXAM)
            internalExaminerMark: Number(row[20]) || 0, // Col U
            agreedMark: Number(row[22]) || 0, // Col W

            attempt: finalAttempt,
            isSpecial: finalAttempt === "special",
            isSupplementary: finalAttempt === "supplementary",
            isRetake: finalAttempt === "re-take",

            unitType: detectedUnitType,
            // Assuming Sheet O16 holds total marks for Q1
            examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
          };

          // --- FIX: Use findOneAndUpdate for atomic upsert ---
          const mark = await Mark.findOneAndUpdate(
            {
              student: student._id,
              programUnit: programUnit._id,
              academicYear: academicYearDoc._id,
            },
            markData,
            { upsert: true, new: true, session, runValidators: true },
          );

          // Trigger the final grade calculation
          await computeFinalGrade({
            markId: mark._id as Types.ObjectId,
            session,
          });
        });
        result.success++;
      } catch (rowErr: any) {
        console.error(`Error processing row ${rowNum}:`, rowErr);
        result.errors.push(`Row ${rowNum} (${regNo}): ${rowErr.message}`);
      } finally {
        await session.endSession();
      }
    }

    return result;
  } catch (err: any) {
    console.error(`[Importer] Fatal Error:`, err.message);
    throw err;
  }
}