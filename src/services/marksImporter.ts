// src/services/marksImporter.ts
// import xlsx from "xlsx";
// import mongoose, { Types } from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
// import { computeFinalGrade } from "./gradeCalculator";
// import { logAudit } from "../lib/auditLogger";
// import type { AuthenticatedRequest } from "../middleware/auth";

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
//     const sheet = workbook.Sheets[workbook.SheetNames[0]];

//     // 1. Meta Data Extraction (Shifted left by 1 column)
//     // Unit Code: H12, Academic Year: F8
//     const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
//     const academicYearText = sheet["F8"]?.v?.toString() || ""; 
//     const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
//     const academicYearStr = yearMatch ? yearMatch[0] : null;

//     // console.log(`[Importer] Metadata Found: Unit=${unitCode}, Year=${academicYearStr}`);

//     if (!unitCode || !academicYearStr) {
//       throw new Error(`Invalid Template: Missing Unit Code (H12) or Academic Year (F8). Found: Unit=${unitCode}, Year=${academicYearStr}`);
//     }

    

//     // 2. Parse Rows as Raw Arrays (header: 1)
//     // range: 16 starts reading at Row 17 (the first student)
//     const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 16 });

//     // 3. Pre-fetch shared data
//     const academicYearDoc = await AcademicYear.findOne({ 
//       year: { $regex: new RegExp(`^${academicYearStr}$`, "i") }, 
//       institution: institutionId 
//     }).lean();

//     if (!academicYearDoc) throw new Error(`Academic Year '${academicYearStr}' not found in database.`);

//     const programUnits = await ProgramUnit.find({ institution: institutionId }).populate("unit").lean();
//     const programUnitMap = new Map(programUnits.map((pu: any) => [`${pu.program.toString()}_${pu.unit.code.toUpperCase()}`, pu]));

//     // 4. Row Processing (A=0, B=1, C=2...)
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
//           if (!student) throw new Error(`Student ${regNo} not found in database.`);

//           const programUnitKey = `${student.program.toString()}_${unitCode}`;
//           const programUnit = programUnitMap.get(programUnitKey);
//           if (!programUnit) throw new Error(`Unit ${unitCode} not linked to student program.`);

//           const markData = {
//             student: student._id,
//             programUnit: programUnit._id,
//             academicYear: academicYearDoc._id,
//             institution: institutionId,
//             uploadedBy: req.user._id,

//             // CA Scores (Col E - G)
//             cat1Raw: Number(row[4]) || 0,
//             cat2Raw: Number(row[5]) || 0,
//             cat3Raw: Number(row[6]) || 0,
//             // Assignments (Col I - K)
//             assgnt1Raw: Number(row[8]) || 0,
//             assgnt2Raw: Number(row[9]) || 0,
//             assgnt3Raw: Number(row[10]) || 0,

//             // Exam Questions (Col N - R)
//             examQ1Raw: Number(row[13]) || 0,
//             examQ2Raw: Number(row[14]) || 0,
//             examQ3Raw: Number(row[15]) || 0,
//             examQ4Raw: Number(row[16]) || 0,
//             examQ5Raw: Number(row[17]) || 0,

//             // Totals from Excel Formulas
//             caTotal30: Number(row[12]) || 0,   // Col M
//             examTotal70: Number(row[18]) || 0, // Col S
//             agreedMark: Number(row[21]) || 0,  // Col V

//             attempt: row[3]?.toString().toLowerCase().includes("supp") ? "supplementary" : "1st", // Col D
//             isSupplementary: row[3]?.toString().toLowerCase().includes("supp"),
//           };

//           const mark = await Mark.findOneAndUpdate(
//             { student: student._id, programUnit: programUnit._id, academicYear: academicYearDoc._id },
//             markData,
//             { upsert: true, new: true, session }
//           );

//           await computeFinalGrade({ markId: mark._id as Types.ObjectId, coordinatorReq: req, session });
//           // await computeFinalGrade({ markId: mark._id as Types.ObjectId,  session });
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

// src/services/marksImporter.ts
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

/**
 * Senior Engineer Note:
 * This importer strictly follows the column mapping defined in uploadTemplate.ts.
 * Column indices are 0-based (A=0, B=1, etc.)
 */
export async function importMarksFromBuffer(
  buffer: Buffer,
  filename: string,
  req: AuthenticatedRequest
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

  try {
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // 1. Meta Data Extraction
    // Unit Code: H12, Academic Year: F8 (as merged cells)
    const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
    const academicYearText = sheet["E8"]?.v?.toString() || ""; 

    // console.log(`[Importer] Checking E8: "${academicYearText}", H12: "${unitCode}"`);

    const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
    const academicYearStr = yearMatch ? yearMatch[0] : null;

    if (!unitCode || !academicYearStr) {
      throw new Error(`Invalid Template: Missing Unit Code (H12) or Academic Year (F8).`);
    }

    // 2. Parse Rows starting from Row 17 (range: 16)
    const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 16 });

    // 3. Pre-fetch shared data
    const academicYearDoc = await AcademicYear.findOne({ 
      year: { $regex: new RegExp(`^${academicYearStr}$`, "i") }, 
      institution: institutionId 
    }).lean();

    if (!academicYearDoc) throw new Error(`Academic Year '${academicYearStr}' not found.`);

    const programUnits = await ProgramUnit.find({ institution: institutionId }).populate("unit").lean();
    const programUnitMap = new Map(programUnits.map((pu: any) => [`${pu.program.toString()}_${pu.unit.code.toUpperCase()}`, pu]));

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
          const student = await Student.findOne({ regNo, institution: institutionId }).lean();
          if (!student) throw new Error(`Student ${regNo} not found.`);

          const programUnitKey = `${student.program.toString()}_${unitCode}`;
          const programUnit = programUnitMap.get(programUnitKey);
          if (!programUnit) throw new Error(`Unit ${unitCode} not linked to program.`);

          const markData = {
            student: student._id,
            programUnit: programUnit._id,
            academicYear: academicYearDoc._id,
            institution: institutionId,
            uploadedBy: req.user._id,

            // CA Scores
            cat1Raw: Number(row[4]) || 0, // Col E
            cat2Raw: Number(row[5]) || 0, // Col F
            cat3Raw: Number(row[6]) || 0, // Col G

            // Assignments
            assgnt1Raw: Number(row[8]) || 0, // Col I
            assgnt2Raw: Number(row[9]) || 0, // Col J
            assgnt3Raw: Number(row[10]) || 0, // Col K

            // Practical (NEW COLUMN)
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
            agreedMark: Number(row[22]) || 0, // Col W (AGREED MARKS)

            attempt: row[3]?.toString().toLowerCase().includes("supp")
              ? "supplementary"
              : "1st", // Col D
            isSupplementary: row[3]?.toString().toLowerCase().includes("supp"),
            examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
          };

          const mark = await Mark.findOneAndUpdate(
            { student: student._id, programUnit: programUnit._id, academicYear: academicYearDoc._id },
            markData,
            { upsert: true, new: true, session }
          );

          // Trigger the final grade calculation using the updated logic
          await computeFinalGrade({ markId: mark._id as Types.ObjectId, session });
        });
        result.success++;
      } catch (rowErr: any) {
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