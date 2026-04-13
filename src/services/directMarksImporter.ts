// // src/services/directMarksImporter.ts

// import xlsx from "xlsx";
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Unit from "../models/Unit"; // Added Unit import
// import MarkDirect from "../models/MarkDirect";
// import type { AuthenticatedRequest } from "../middleware/auth";

// interface ImportResult { total: number; success: number; errors: string[]; warnings: string[]; }

// export async function importDirectMarksFromBuffer(
//   buffer: Buffer,
//   filename: string,
//   req: AuthenticatedRequest,
// ): Promise<ImportResult> {
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = { total: 0, success: 0, errors: [], warnings: []};

//   const workbook = xlsx.read(buffer, { type: "buffer" });
//   const sheet = workbook.Sheets[workbook.SheetNames[0]];

//   // 1. Meta Data Extraction
//   const unitCode = sheet["F12"]?.v?.toString().trim().toUpperCase();
//   const academicYearText = sheet["C8"]?.v?.toString() || "";
//   const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
//   const academicYearStr = yearMatch ? yearMatch[0] : null;

//   if (!unitCode || !academicYearStr) throw new Error("Template metadata missing (Unit Code or Year).");

//   // 2. Resolve Global Records (Unit and Academic Year)
//   const [unitDoc, academicYearDoc] = await Promise.all([
//     Unit.findOne({ code: unitCode }).lean(),
//     AcademicYear.findOne({
//       year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
//       institution: institutionId,
//     }).lean(),
//   ]);

//   if (!unitDoc)
//     throw new Error(
//       `Unit with code '${unitCode}' does not exist in the database.`,
//     );
//   if (!academicYearDoc)
//     throw new Error(`Academic Year '${academicYearStr}' not found.`);

//   // 3. Parse Rows
//   const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, {
//     header: 1,
//     range: 15,
//   });

//   for (const [index, row] of rawRows.entries()) {
//     const regNo = row[1]?.toString().trim().toUpperCase();
//     if (!regNo) continue;

//     result.total++;
//     const session = await mongoose.startSession();
//     try {
//       await session.withTransaction(async () => {
//         const student = await Student.findOne({
//           regNo,
//           institution: institutionId,
//         }).lean();
//         if (!student) throw new Error(`Student ${regNo} not found`);

//         // 4. Robust ProgramUnit Lookup
//         // We look for a record that links the Student's Program to the Unit from the Excel
//         const programUnit = await ProgramUnit.findOne({
//           institution: institutionId,
//           program: student.program, // This is the ID from the error: 693534f37ea0366e6b831034
//           unit: unitDoc._id,
//         }).lean();

//         if (!programUnit) {
//           throw new Error(
//             `Unit ${unitCode} is not registered in this student's curriculum (Program ID: ${student.program}).`,
//           );
//         }

//         const markData = {
//           institution: institutionId,
//           student: student._id,
//           programUnit: programUnit._id,
//           academicYear: academicYearDoc._id,
//           // semester: academicYearText.toUpperCase().includes("SEMESTER 2")
//           //   ? "SEMESTER 2" : "SEMESTER 1",
//           caTotal30: Number(row[4]) || 0,
//           examTotal70: Number(row[5]) || 0,
//           externalTotal100: row[7] !== undefined && row[7] !== null ? Number(row[7]) : null,
//           agreedMark: Number(row[8]) || 0,
//           attempt: row[3]?.toString().toLowerCase().includes("supp")
//             ? "supplementary"
//             : "1st",
//           uploadedBy: req.user._id,
//         };

//         await MarkDirect.findOneAndUpdate(
//           {
//             student: student._id,
//             programUnit: programUnit._id,
//             academicYear: academicYearDoc._id,
//           },
//           markData,
//           { upsert: true, session },
//         );

        
//       });
//       result.success++;
//     } catch (err: any) {
//       result.errors.push(`Row ${index + 16} (${regNo}): ${err.message}`);
//     } finally {
//       await session.endSession();
//     }
//   }
//   return result;
// }























// // src/services/directMarksImporter.ts
// import xlsx from "xlsx";
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Unit from "../models/Unit";
// import MarkDirect from "../models/MarkDirect";
// import type { AuthenticatedRequest } from "../middleware/auth";

// interface ImportResult { total: number; success: number; errors: string[]; warnings: string[] }

// // ── Attempt detection matching all labels used in directTemplate.ts ───────────
// // Labels in ATTEMPT column D:  1st, A/S, SPEC, RP1C, A/SO, RPU1, B/S, A/RA1 etc.

// function detectAttemptType(rawCell: any): string {
//   const raw = (rawCell?.toString() || "").toLowerCase().trim();

//   if (!raw) return "1st";

//   // Supplementary (A/S, Supp, supp)
//   if (raw === "a/s" || raw.startsWith("supp")) return "supplementary";

//   // Special (SPEC, Special, special)
//   if (raw === "spec" || raw.includes("special")) return "special";

//   // Carry forward (RP1C, RP2C, RP3C, A/CF, a/cf)
//   if (/rp\d+c/i.test(raw) || raw === "a/cf") return "re-take";

//   // Stayout retake (A/SO, A/SOS, a/so)
//   if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take";

//   // Repeat unit (RPU1, RPU2, rpu)
//   if (/rpu\d*/i.test(raw)) return "re-take";

//   // Repeat year / re-admission — these are FIRST attempt in their repeated year
//   // e.g. A/RA1, RP1, B/S for a repeat year student
//   if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st";

//   // Default
//   return "1st";
// }

// export async function importDirectMarksFromBuffer(
//   buffer:   Buffer,
//   filename: string,
//   req:      AuthenticatedRequest,
// ): Promise<ImportResult> {
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

//   const workbook = xlsx.read(buffer, { type: "buffer" });
//   const sheet    = workbook.Sheets[workbook.SheetNames[0]];

//   // ── Metadata extraction ───────────────────────────────────────────────────
//   // directTemplate.ts puts unit code at F12 and year at C8
//   const unitCode       = sheet["F12"]?.v?.toString().trim().toUpperCase();
//   const yearText       = sheet["C8"]?.v?.toString() || "";
//   const yearMatch      = yearText.match(/\d{4}\/\d{4}/);
//   const academicYearStr = yearMatch ? yearMatch[0] : null;

//   if (!unitCode || !academicYearStr) {
//     throw new Error(`Template metadata missing. Found Unit: "${unitCode}", Year: "${academicYearStr}". Check cells F12 and C8.`);
//   }

//   const [unitDoc, academicYearDoc] = await Promise.all([
//     Unit.findOne({ code: unitCode }).lean(),
//     AcademicYear.findOne({
//       year:        { $regex: new RegExp(`^${academicYearStr}$`, "i") },
//       institution: institutionId,
//     }).lean(),
//   ]);

//   if (!unitDoc)         throw new Error(`Unit "${unitCode}" not found.`);
//   if (!academicYearDoc) throw new Error(`Academic Year "${academicYearStr}" not found.`);

//   // ── Parse rows (data starts at row 17 in the direct template, index 16) ───
//   const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 15 });

//   for (const [index, row] of rawRows.entries()) {
//     const regNo = row[1]?.toString().trim().toUpperCase();
//     if (!regNo) continue;

//     result.total++;
//     const session = await mongoose.startSession();

//     try {
//       await session.withTransaction(async () => {
//         const student = await Student.findOne({ regNo, institution: institutionId }).lean();
//         if (!student) throw new Error(`Student ${regNo} not found`);

//         const programUnit = await ProgramUnit.findOne({
//           institution: institutionId,
//           program:     (student as any).program,
//           unit:        (unitDoc as any)._id,
//         }).lean();

//         if (!programUnit) {
//           throw new Error(
//             `Unit ${unitCode} is not in the curriculum for ${regNo}'s programme (Program ID: ${(student as any).program}).`,
//           );
//         }

//         // ── Attempt detection ────────────────────────────────────────────
//         const rawAttempt   = row[3]; // Column D = index 3
//         const attempt      = detectAttemptType(rawAttempt);
//         const isSpecial    = attempt === "special";
//         const isSupp       = attempt === "supplementary";
//         const isRetake     = attempt === "re-take";

//         // ── Column mapping (directTemplate layout) ───────────────────────
//         //  A=0  S/N
//         //  B=1  REG. NO.
//         //  C=2  NAME
//         //  D=3  ATTEMPT
//         //  E=4  CA TOTAL (/30)
//         //  F=5  EXAM TOTAL (/70)
//         //  G=6  INTERNAL (/100)   [formula — usually ignored on import]
//         //  H=7  EXTERNAL (/100)
//         //  I=8  AGREED (/100)
//         //  J=9  GRADE

//         const caTotal30      = Number(row[4]) || 0;
//         const examTotal70    = Number(row[5]) || 0;
//         const externalTotal  = row[7] !== undefined && row[7] !== null ? Number(row[7]) : null;
//         const agreedMark     = Number(row[8]) || 0;

//         const markData = {
//           institution:      institutionId,
//           student:          (student as any)._id,
//           programUnit:      (programUnit as any)._id,
//           academicYear:     (academicYearDoc as any)._id,
//           caTotal30,
//           examTotal70,
//           externalTotal100: externalTotal,
//           agreedMark,
//           attempt,
//           isSpecial,
//           isSupplementary:  isSupp,
//           isRetake,
//           uploadedBy:       req.user._id,
//         };

//         await MarkDirect.findOneAndUpdate(
//           {
//             student:      (student as any)._id,
//             programUnit:  (programUnit as any)._id,
//             academicYear: (academicYearDoc as any)._id,
//           },
//           markData,
//           { upsert: true, new: true, session },
//         );
//       });

//       result.success++;
//     } catch (err: any) {
//       result.errors.push(`Row ${index + 16} (${regNo}): ${err.message}`);
//     } finally {
//       await session.endSession();
//     }
//   }

//   return result;
// }
















// // src/services/directMarksImporter.ts
// import xlsx from "xlsx";
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import Unit from "../models/Unit";
// import MarkDirect from "../models/MarkDirect";
// import type { AuthenticatedRequest } from "../middleware/auth";

// interface ImportResult { total: number; success: number; errors: string[]; warnings: string[] }

// // Maps ALL attempt labels used in directTemplate.ts back to DB attempt strings.
// // Column D values: B/S, A/S, Supp, SPEC, Special, RP1C, RP2C, A/CF,
// //                  A/SO, A/SOS, RPU1, RPU2, A/RA1, RP1, RP2
// function detectAttemptType(rawCell: any): string {
//   const raw = (rawCell?.toString() || "").toLowerCase().trim();
//   if (!raw)                                                      return "1st";
//   if (raw === "a/s" || raw.startsWith("supp"))                  return "supplementary";
//   if (raw === "spec" || raw.includes("special"))                 return "special";
//   if (/rp\d+c/i.test(raw) || raw === "a/cf")                    return "re-take"; // carry forward
//   if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take"; // stayout retake
//   if (/rpu\d*/i.test(raw))                                       return "re-take"; // repeat unit
//   if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st"; // repeat year / re-admission
//   return "1st";
// }

// export async function importDirectMarksFromBuffer(
//   buffer: Buffer, filename: string, req: AuthenticatedRequest,
// ): Promise<ImportResult> {
//   const institutionId = req.user.institution;
//   if (!institutionId) throw new Error("Coordinator not linked to institution");

//   const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };
//   const workbook = xlsx.read(buffer, { type: "buffer" });
//   const sheet    = workbook.Sheets[workbook.SheetNames[0]];

//   // directTemplate puts unit code at F12, year at C8
//   const unitCode        = sheet["F12"]?.v?.toString().trim().toUpperCase();
//   const yearText        = sheet["C8"]?.v?.toString() || "";
//   const yearMatch       = yearText.match(/\d{4}\/\d{4}/);
//   const academicYearStr = yearMatch ? yearMatch[0] : null;

//   if (!unitCode || !academicYearStr)
//     throw new Error(`Metadata missing. Unit: "${unitCode}", Year: "${academicYearStr}". Check cells F12 and C8.`);

//   const [unitDoc, academicYearDoc] = await Promise.all([
//     Unit.findOne({ code: unitCode }).lean(),
//     AcademicYear.findOne({ year: { $regex: new RegExp(`^${academicYearStr}$`, "i") }, institution: institutionId }).lean(),
//   ]);

//   if (!unitDoc)         throw new Error(`Unit "${unitCode}" not found.`);
//   if (!academicYearDoc) throw new Error(`Academic Year "${academicYearStr}" not found.`);

//   // Data starts at row 17 (range:15 skips the first 15 rows → index 0 = row 16 header, index 1 = row 17)
//   const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 15 });

//   for (const [index, row] of rawRows.entries()) {
//     const regNo = row[1]?.toString().trim().toUpperCase();
//     if (!regNo) continue;

//     result.total++;
//     const session = await mongoose.startSession();

//     try {
//       await session.withTransaction(async () => {
//         const student = await Student.findOne({ regNo, institution: institutionId }).lean();
//         if (!student) throw new Error(`Student ${regNo} not found`);

//         const programUnit = await ProgramUnit.findOne({
//           institution: institutionId,
//           program:     (student as any).program,
//           unit:        (unitDoc as any)._id,
//         }).lean();

//         if (!programUnit)
//           throw new Error(`Unit ${unitCode} is not in the curriculum for ${regNo} (Program: ${(student as any).program}).`);

//         // Column layout (directTemplate.ts):
//         //  A=0  S/N    B=1  REG. NO.   C=2  NAME     D=3  ATTEMPT
//         //  E=4  CA(/30) F=5 EXAM(/70)  G=6  INTERNAL H=7  EXTERNAL  I=8  AGREED  J=9  GRADE
//         const attempt      = detectAttemptType(row[3]);
//         const isSpecial    = attempt === "special";
//         const isSupp       = attempt === "supplementary";
//         const isRetake     = attempt === "re-take";

//         const markData = {
//           institution:      institutionId,
//           student:          (student as any)._id,
//           programUnit:      (programUnit as any)._id,
//           academicYear:     (academicYearDoc as any)._id,
//           caTotal30:        Number(row[4]) || 0,
//           examTotal70:      Number(row[5]) || 0,
//           externalTotal100: row[7] !== undefined && row[7] !== null ? Number(row[7]) : null,
//           agreedMark:       Number(row[8]) || 0,
//           attempt,
//           isSpecial,
//           isSupplementary:  isSupp,
//           isRetake,
//           uploadedBy:       req.user._id,
//         };

//         await MarkDirect.findOneAndUpdate(
//           { student: (student as any)._id, programUnit: (programUnit as any)._id, academicYear: (academicYearDoc as any)._id },
//           markData,
//           { upsert: true, new: true, session },
//         );
//       });
//       result.success++;
//     } catch (err: any) {
//       result.errors.push(`Row ${index + 16} (${regNo}): ${err.message}`);
//     } finally {
//       await session.endSession();
//     }
//   }

//   return result;
// }











// src/services/directMarksImporter.ts
import xlsx from "xlsx";
import mongoose from "mongoose";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import Unit from "../models/Unit";
import MarkDirect from "../models/MarkDirect";
import type { AuthenticatedRequest } from "../middleware/auth";

interface ImportResult { total: number; success: number; errors: string[]; warnings: string[] }

// Maps ALL attempt labels used in directTemplate.ts back to DB attempt strings.
// Column D values: B/S, A/S, Supp, SPEC, Special, RP1C, RP2C, A/CF,
//                  A/SO, A/SOS, RPU1, RPU2, A/RA1, RP1, RP2
function detectAttemptType(rawCell: any): string {
  const raw = (rawCell?.toString() || "").toLowerCase().trim();
  console.log(`[directImporter] detectAttemptType raw="${raw}"`);
  if (!raw)                                                           return "1st";
  if (raw === "a/s" || raw.startsWith("supp"))                       return "supplementary";
  if (raw === "spec" || raw.includes("special"))                     return "special";
  if (/rp\d+c/i.test(raw) || raw === "a/cf")                        return "re-take"; // carry forward
  if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take"; // stayout retake
  if (/rpu\d*/i.test(raw))                                           return "re-take"; // repeat unit
  if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st"; // repeat year / re-admission
  return "1st";
}

export async function importDirectMarksFromBuffer(
  buffer: Buffer, filename: string, req: AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  console.log(`[directImporter] Starting import. File: ${filename}, Institution: ${institutionId}`);

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };
  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  console.log(`[directImporter] Sheet name: "${sheetName}"`);

  // directTemplate puts unit code at F12, year at C8
  const unitCodeRaw     = sheet["F12"]?.v;
  const yearTextRaw     = sheet["C8"]?.v;
  const unitCode        = unitCodeRaw?.toString().trim().toUpperCase();
  const yearText        = yearTextRaw?.toString() || "";
  const yearMatch       = yearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  console.log(`[directImporter] Metadata — unitCode: "${unitCode}", yearText: "${yearText}", academicYearStr: "${academicYearStr}"`);

  if (!unitCode || !academicYearStr) {
    const msg = `Metadata missing. Unit code at F12: "${unitCode}", Academic year at C8: "${academicYearStr}". Check cells F12 and C8.`;
    console.error(`[directImporter] ${msg}`);
    throw new Error(msg);
  }

  // Look up the Unit document
  const unitDoc = await Unit.findOne({ code: unitCode }).lean();
  console.log(`[directImporter] Unit lookup for code="${unitCode}": ${unitDoc ? `found _id=${(unitDoc as any)._id}` : "NOT FOUND"}`);
  if (!unitDoc) throw new Error(`Unit "${unitCode}" not found in the database.`);

  // Look up the AcademicYear document
  const academicYearDoc = await AcademicYear.findOne({
    year: { $regex: new RegExp(`^${academicYearStr.replace("/", "\\/")}$`, "i") },
    institution: institutionId,
  }).lean();
  console.log(`[directImporter] AcademicYear lookup for "${academicYearStr}": ${academicYearDoc ? `found _id=${(academicYearDoc as any)._id}` : "NOT FOUND"}`);
  if (!academicYearDoc) throw new Error(`Academic Year "${academicYearStr}" not found for this institution.`);

  // Data starts at row 17 (range:15 means we skip first 15 rows, so index 0 = row 16, index 1 = row 17)
  const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 15 });
  console.log(`[directImporter] Total raw rows parsed (from row 16 onward): ${rawRows.length}`);

  // Pre-fetch all ProgramUnits for this unit across all programs in the institution
  // to avoid per-row DB queries where possible
  const allProgramUnitsForUnit = await ProgramUnit.find({
    unit: (unitDoc as any)._id,
  }).lean();
  console.log(`[directImporter] ProgramUnits found for unit ${unitCode}: ${allProgramUnitsForUnit.length}`);

  for (const [index, row] of rawRows.entries()) {
    // Row index 0 = spreadsheet row 16 (header), index 1 = row 17 (first data row)
    // Skip the header row (index 0) and any row without a reg number
    const regNo = row[1]?.toString().trim().toUpperCase();
    if (!regNo || regNo === "REG. NO." || regNo === "REG NO") {
      console.log(`[directImporter] Row ${index + 16}: skipping (no regNo or header row)`);
      continue;
    }

    result.total++;
    const rowNum = index + 16;
    console.log(`[directImporter] Processing row ${rowNum}: regNo="${regNo}"`);

    const session = await mongoose.startSession();
    try {
      await session.withTransaction(async () => {
        // Find student
        const student = await Student.findOne({ regNo, institution: institutionId }).lean();
        if (!student) {
          throw new Error(`Student "${regNo}" not found for this institution`);
        }
        console.log(`[directImporter] Row ${rowNum}: student found _id=${(student as any)._id}, program=${(student as any).program}`);

        // Find the correct ProgramUnit: must belong to student's program AND reference our unit
        const programUnit = allProgramUnitsForUnit.find(
          (pu: any) => pu.program.toString() === (student as any).program.toString()
        );

        if (!programUnit) {
          // Fallback: query directly in case the pre-fetch missed it
          console.warn(`[directImporter] Row ${rowNum}: ProgramUnit not in pre-fetched list, querying directly...`);
          const directPU = await ProgramUnit.findOne({
            program: (student as any).program,
            unit: (unitDoc as any)._id,
          }).lean();

          if (!directPU) {
            throw new Error(
              `Unit "${unitCode}" is not linked to the curriculum for program of student "${regNo}" (programId: ${(student as any).program})`
            );
          }
          console.log(`[directImporter] Row ${rowNum}: ProgramUnit found via direct query: _id=${(directPU as any)._id}`);
          await upsertMark(directPU, student, row, rowNum, institutionId, academicYearDoc, unitDoc, req, session, result);
          return;
        }

        console.log(`[directImporter] Row ${rowNum}: ProgramUnit found _id=${(programUnit as any)._id}`);
        await upsertMark(programUnit, student, row, rowNum, institutionId, academicYearDoc, unitDoc, req, session, result);
      });

      result.success++;
      console.log(`[directImporter] Row ${rowNum}: SUCCESS`);
    } catch (err: any) {
      const msg = `Row ${rowNum} (${regNo}): ${err.message}`;
      console.error(`[directImporter] FAILED — ${msg}`, err.stack || "");
      result.errors.push(msg);
    } finally {
      await session.endSession();
    }
  }

  console.log(`[directImporter] Import complete. Total=${result.total}, Success=${result.success}, Errors=${result.errors.length}`);
  return result;
}

// ─── Helper: build markData and upsert ───────────────────────────────────────
async function upsertMark(
  programUnit: any,
  student: any,
  row: any[],
  rowNum: number,
  institutionId: any,
  academicYearDoc: any,
  unitDoc: any,
  req: AuthenticatedRequest,
  session: mongoose.ClientSession,
  result: ImportResult,
): Promise<void> {
  // Column layout (directTemplate.ts):
  //  A=0  S/N    B=1  REG. NO.   C=2  NAME     D=3  ATTEMPT
  //  E=4  CA(/30) F=5 EXAM(/70)  G=6  INTERNAL H=7  EXTERNAL  I=8  AGREED  J=9  GRADE
  const attempt   = detectAttemptType(row[3]);
  const isSpecial = attempt === "special";
  const isSupp    = attempt === "supplementary";
  const isRetake  = attempt === "re-take";

  const caTotal30   = Number(row[4]) || 0;
  const examTotal70 = Number(row[5]) || 0;

  // Agreed mark: prefer explicit col I value, fall back to CA + Exam sum
  const rawAgreed = row[8];
  const agreedMark = (rawAgreed !== undefined && rawAgreed !== null && rawAgreed !== "")
    ? Number(rawAgreed)
    : caTotal30 + examTotal70;

  const externalRaw = row[7];
  const externalTotal100 = (externalRaw !== undefined && externalRaw !== null && externalRaw !== "")
    ? Number(externalRaw)
    : null;

  const isMissingCA = caTotal30 === 0 && !isSupp && !isSpecial;

  console.log(`[directImporter] Row ${rowNum}: attempt="${attempt}", CA=${caTotal30}, Exam=${examTotal70}, Agreed=${agreedMark}, isSpecial=${isSpecial}, isSupp=${isSupp}, isMissingCA=${isMissingCA}`);

  const markData = {
    institution:      institutionId,
    student:          student._id,
    programUnit:      programUnit._id,
    academicYear:     academicYearDoc._id,
    caTotal30,
    examTotal70,
    externalTotal100,
    agreedMark,
    attempt,
    isSpecial,
    isSupplementary:  isSupp,
    isRetake,
    isMissingCA,
    uploadedBy:       req.user._id,
    uploadedAt:       new Date(),
    deletedAt:        null,
  };

  const filter = {
    student:      student._id,
    programUnit:  programUnit._id,
    academicYear: academicYearDoc._id,
  };

  console.log(`[directImporter] Row ${rowNum}: upserting MarkDirect with filter:`, JSON.stringify(filter));

  await MarkDirect.findOneAndUpdate(
    filter,
    { $set: markData },
    { upsert: true, new: true, session },
  );
}