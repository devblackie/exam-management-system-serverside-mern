// // src/services/marksImporter.ts

// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import type { AuthenticatedRequest } from "../middleware/auth";
// import { randomUUID } from "node:crypto";

// interface ImportResult {
//   total: number;
//   success: number;
//   errors: string[];
//   warnings: string[];
// }

// function stripQualifier(rawRegNo: string): string {
//   // Match: slash + 4-digit year + qualifier suffix (letters + digits + optional S2)
//   // e.g.  /2017RP1  →  /2017
//   //        /2016RP1C →  /2016
//   return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
// }

// export async function importMarksFromBuffer(
//   buffer: Buffer,
//   filename: string,
//   req: AuthenticatedRequest,
// ): Promise<ImportResult> {
//   const batchId = randomUUID();
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = {total: 0, success: 0, errors: [], warnings: []};
  

//   try {
//     const workbook = xlsx.read(buffer, { type: "buffer" });
//     const sheet = workbook.Sheets[workbook.SheetNames[0]];

//     const modeIndicator = sheet["E10"]?.v?.toString().toUpperCase() || "";
//     let detectedUnitType: "theory" | "lab" | "workshop" = "theory";

//     if (modeIndicator.includes("LAB")) detectedUnitType = "lab";
//     if (modeIndicator.includes("WORKSHOP")) detectedUnitType = "workshop";

//     // 1. Meta Data Extraction
//     const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
//     const academicYearText = sheet["D8"]?.v?.toString() || "";
//     const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
//     const academicYearStr = yearMatch ? yearMatch[0] : null;

//     if (!unitCode || !academicYearStr) {
//       throw new Error(
//         `Invalid Template: Missing Unit Code (H12) or Academic Year (D8). Found Unit: ${unitCode}, Year: ${academicYearStr}`,
//       );
//     }

//     // 2. Parse Rows
//     const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, {header: 1, range: 16});

//     // 3. Pre-fetch shared data
//     const academicYearDoc = await AcademicYear.findOne({
//       year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
//       institution: institutionId,
//     }).lean();

//     if (!academicYearDoc)
//       throw new Error(`Academic Year '${academicYearStr}' not found.`);

//     // Optimization: Pre-fetch all units for the institution
//     const programUnits = await ProgramUnit.find({ institution: institutionId })
//       .populate("unit").lean();
//     const programUnitMap = new Map(
//       programUnits.map((pu: any) => [`${pu.program.toString()}_${pu.unit.code.toUpperCase()}`, pu]),
//     );

//     // 4. Processing Mapping (A=0, B=1, C=2...)
//     for (const [index, row] of rawRows.entries()) {
//       // const regNo = row[1]?.toString().trim().toUpperCase(); // Col B
//       // const sn = row[0]; // Col A

//       // if (!regNo || sn === "") continue;

//       const rawCell = row[1]?.toString().trim().toUpperCase();
//       const sn = row[0];
//       if (!rawCell || sn === "") continue;
//       const regNo = stripQualifier(rawCell);
      
//       result.total++;
//       const rowNum = index + 17;
//       const session = await mongoose.startSession();

//       try {
//         await session.withTransaction(async () => {
//           const student = await Student.findOne({regNo, institution: institutionId}).lean();
//           if (!student) throw new Error(`Student ${regNo} not found.`);

//           const programUnitKey = `${student.program.toString()}_${unitCode}`;
//           const programUnit = programUnitMap.get(programUnitKey);
//           if (!programUnit)
//             throw new Error(`Unit ${unitCode} not linked to student's program.`);

//           // --- FIX: SMART ATTEMPT DETECTION ---
//           // Removing the pre-find (Mark.find) lookup as it might be causing issues
//           // with soft-delete middleware if a previous record exists but is deleted.

//           const excelAttemptLabel = row[3]?.toString().trim(); // Column D
//           let finalAttempt = "1st";

//           if (excelAttemptLabel?.toLowerCase().includes("supp")) {
//             finalAttempt = "supplementary";
//           } else if (excelAttemptLabel?.toLowerCase().includes("special")) {
//             finalAttempt = "special";
//           } else if (excelAttemptLabel?.toLowerCase().includes("retake")) {
//             finalAttempt = "re-take";
//           }

//           // --- FIX: Mapping Columns and Ensuring Numbers ---
//           const markData = {
//             student: student._id,
//             programUnit: programUnit._id,
//             academicYear: academicYearDoc._id,
//             institution: institutionId,
//             uploadedBy: req.user._id,
//             deletedAt: null, // Ensure we are updating active records
//             batchId,
//             // CA Scores
//             cat1Raw: Number(row[4]) || 0, // Col E
//             cat2Raw: Number(row[5]) || 0, // Col F
//             cat3Raw: Number(row[6]) || 0, // Col G

//             // Assignments
//             assgnt1Raw: Number(row[8]) || 0, // Col I
//             assgnt2Raw: Number(row[9]) || 0, // Col J
//             assgnt3Raw: Number(row[10]) || 0, // Col K

//             // Practical
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
//             internalExaminerMark: Number(row[20]) || 0, // Col U
//             agreedMark: Number(row[22]) || 0, // Col W

//             attempt: finalAttempt,
//             isSpecial: finalAttempt === "special",
//             isSupplementary: finalAttempt === "supplementary",
//             isRetake: finalAttempt === "re-take",

//             unitType: detectedUnitType,
//             // Assuming Sheet O16 holds total marks for Q1
//             examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
//           };

//           // --- FIX: Use findOneAndUpdate for atomic upsert ---
//           const mark = await Mark.findOneAndUpdate(
//             { student: student._id, programUnit: programUnit._id, academicYear: academicYearDoc._id},
//             markData,
//             { upsert: true, new: true, session, runValidators: true },
//           );

//           // Trigger the final grade calculation
//           await computeFinalGrade({markId: mark._id as Types.ObjectId, session});
//         });
//         result.success++;
//       } catch (rowErr: any) {
//         console.error(`Error processing row ${rowNum}:`, rowErr);
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

































// // serverside/src/services/marksImporter.ts — COMPLETE

// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student      from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit  from "../models/ProgramUnit";
// import Mark         from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import type { AuthenticatedRequest } from "../middleware/auth";
// import { randomUUID } from "node:crypto";

// interface ImportResult {
//   total:    number;
//   success:  number;
//   errors:   string[];
//   warnings: string[];
// }

// function stripQualifier(rawRegNo: string): string {
//   // e.g. /2017RP1 → /2017,  /2016RP1C → /2016
//   return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
// }

// export async function importMarksFromBuffer(
//   buffer:   Buffer,
//   filename: string,
//   req:      AuthenticatedRequest,
// ): Promise<ImportResult> {
//   const batchId = randomUUID();
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

//   const workbook = xlsx.read(buffer, { type: "buffer" });
//   const sheet    = workbook.Sheets[workbook.SheetNames[0]];

//   // Detect unit type from scoresheet header
//   const modeIndicator = (sheet["E10"]?.v?.toString() ?? "").toUpperCase();
//   let detectedUnitType: "theory" | "lab" | "workshop" = "theory";
//   if (modeIndicator.includes("LAB"))      detectedUnitType = "lab";
//   if (modeIndicator.includes("WORKSHOP")) detectedUnitType = "workshop";

//   // ── Extract metadata from header cells ──────────────────────────────────────
//   const unitCode       = sheet["H12"]?.v?.toString().trim().toUpperCase();
//   const academicYearText = sheet["D8"]?.v?.toString() ?? "";
//   const yearMatch      = academicYearText.match(/\d{4}\/\d{4}/);
//   const academicYearStr = yearMatch ? yearMatch[0] : null;

//   if (!unitCode || !academicYearStr) {
//     throw new Error(
//       `Invalid template: missing Unit Code (H12="${unitCode}") or Academic Year (D8="${academicYearStr}")`,
//     );
//   }

//   // ── Pre-fetch shared data (once, outside the loop) ───────────────────────────
//   const academicYearDoc = await AcademicYear.findOne({
//     year:        { $regex: new RegExp(`^${academicYearStr}$`, "i") },
//     institution: institutionId,
//   }).lean();

//   if (!academicYearDoc) {
//     throw new Error(`Academic Year "${academicYearStr}" not found.`);
//   }

//   const allProgramUnits = await ProgramUnit.find({ institution: institutionId })
//     .populate("unit")
//     .lean();

//   const programUnitMap = new Map(
//     allProgramUnits.map((pu: any) => [
//       `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`,
//       pu,
//     ]),
//   );

//   // ── Parse data rows (starting at row 17 = index 0 after range:16) ───────────
//   const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 16 });

//   for (const [index, row] of rawRows.entries()) {
//     const rawCell = row[1]?.toString().trim().toUpperCase();
//     const sn      = row[0];
//     if (!rawCell || sn === "") continue;

//     const regNo = stripQualifier(rawCell);
//     result.total++;
//     const rowNum = index + 17;

//     try {
//       // ── Resolve student ──────────────────────────────────────────────────────
//       const student = await Student.findOne({
//         regNo:       regNo,
//         institution: institutionId,
//       }).lean();

//       if (!student) {
//         result.errors.push(`Row ${rowNum} (${regNo}): Student not found.`);
//         continue;
//       }

//       // ── Resolve program unit ─────────────────────────────────────────────────
//       const programUnitKey = `${(student as any).program.toString()}_${unitCode}`;
//       const programUnit    = programUnitMap.get(programUnitKey);

//       if (!programUnit) {
//         result.errors.push(
//           `Row ${rowNum} (${regNo}): Unit "${unitCode}" not linked to student's program.`,
//         );
//         continue;
//       }

//       // ── Map attempt from column D ────────────────────────────────────────────
//       const excelAttemptLabel = (row[3]?.toString().trim() ?? "").toLowerCase();
//       let finalAttempt = "1st";
//       if (excelAttemptLabel.includes("supp"))    finalAttempt = "supplementary";
//       else if (excelAttemptLabel.includes("special")) finalAttempt = "special";
//       else if (excelAttemptLabel.includes("retake"))  finalAttempt = "re-take";

//       // ── Build mark payload ───────────────────────────────────────────────────
//       const markData = {
//         student:      (student as any)._id,
//         programUnit:  (programUnit as any)._id,
//         academicYear: (academicYearDoc as any)._id,
//         institution:  institutionId,
//         uploadedBy:   req.user._id,
//         deletedAt:    null,
//         batchId,

//         cat1Raw:      Number(row[4])  || 0,
//         cat2Raw:      Number(row[5])  || 0,
//         cat3Raw:      Number(row[6])  || 0,
//         assgnt1Raw:   Number(row[8])  || 0,
//         assgnt2Raw:   Number(row[9])  || 0,
//         assgnt3Raw:   Number(row[10]) || 0,
//         practicalRaw: Number(row[12]) || 0,
//         examQ1Raw:    Number(row[14]) || 0,
//         examQ2Raw:    Number(row[15]) || 0,
//         examQ3Raw:    Number(row[16]) || 0,
//         examQ4Raw:    Number(row[17]) || 0,
//         examQ5Raw:    Number(row[18]) || 0,
//         caTotal30:    Number(row[13]) || 0,
//         examTotal70:  Number(row[19]) || 0,
//         internalExaminerMark: Number(row[20]) || 0,
//         agreedMark:   Number(row[22]) || 0,

//         attempt:         finalAttempt,
//         isSpecial:       finalAttempt === "special",
//         isSupplementary: finalAttempt === "supplementary",
//         isRetake:        finalAttempt === "re-take",
//         unitType:        detectedUnitType,
//         examMode:        sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
//       };

//       // ── Atomic upsert — no session needed, findOneAndUpdate is atomic ─────────
//       const mark = await Mark.findOneAndUpdate(
//         {
//           student:      (student as any)._id,
//           programUnit:  (programUnit as any)._id,
//           academicYear: (academicYearDoc as any)._id,
//         },
//         markData,
//         { upsert: true, new: true, runValidators: true },
//       );

//       // ── Compute final grade (outside any transaction) ─────────────────────────
//       await computeFinalGrade({ markId: mark._id as Types.ObjectId });

//       result.success++;

//     } catch (rowErr: unknown) {
//       const msg = rowErr instanceof Error ? rowErr.message : String(rowErr);
//       console.error(`[marksImporter] Row ${rowNum} (${regNo}):`, msg);
//       result.errors.push(`Row ${rowNum} (${regNo}): ${msg}`);
//     }
//   }

//   return result;
// }






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