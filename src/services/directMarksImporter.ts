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

// ── Attempt detection matching all labels used in directTemplate.ts ───────────
// Labels in ATTEMPT column D:  1st, A/S, SPEC, RP1C, A/SO, RPU1, B/S, A/RA1 etc.

function detectAttemptType(rawCell: any): string {
  const raw = (rawCell?.toString() || "").toLowerCase().trim();

  if (!raw) return "1st";

  // Supplementary (A/S, Supp, supp)
  if (raw === "a/s" || raw.startsWith("supp")) return "supplementary";

  // Special (SPEC, Special, special)
  if (raw === "spec" || raw.includes("special")) return "special";

  // Carry forward (RP1C, RP2C, RP3C, A/CF, a/cf)
  if (/rp\d+c/i.test(raw) || raw === "a/cf") return "re-take";

  // Stayout retake (A/SO, A/SOS, a/so)
  if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout")) return "re-take";

  // Repeat unit (RPU1, RPU2, rpu)
  if (/rpu\d*/i.test(raw)) return "re-take";

  // Repeat year / re-admission — these are FIRST attempt in their repeated year
  // e.g. A/RA1, RP1, B/S for a repeat year student
  if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw)) return "1st";

  // Default
  return "1st";
}

export async function importDirectMarksFromBuffer(
  buffer:   Buffer,
  filename: string,
  req:      AuthenticatedRequest,
): Promise<ImportResult> {
  const institutionId = req.user.institution;
  if (!institutionId) throw new Error("Coordinator not linked to institution");

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

  const workbook = xlsx.read(buffer, { type: "buffer" });
  const sheet    = workbook.Sheets[workbook.SheetNames[0]];

  // ── Metadata extraction ───────────────────────────────────────────────────
  // directTemplate.ts puts unit code at F12 and year at C8
  const unitCode       = sheet["F12"]?.v?.toString().trim().toUpperCase();
  const yearText       = sheet["C8"]?.v?.toString() || "";
  const yearMatch      = yearText.match(/\d{4}\/\d{4}/);
  const academicYearStr = yearMatch ? yearMatch[0] : null;

  if (!unitCode || !academicYearStr) {
    throw new Error(`Template metadata missing. Found Unit: "${unitCode}", Year: "${academicYearStr}". Check cells F12 and C8.`);
  }

  const [unitDoc, academicYearDoc] = await Promise.all([
    Unit.findOne({ code: unitCode }).lean(),
    AcademicYear.findOne({
      year:        { $regex: new RegExp(`^${academicYearStr}$`, "i") },
      institution: institutionId,
    }).lean(),
  ]);

  if (!unitDoc)         throw new Error(`Unit "${unitCode}" not found.`);
  if (!academicYearDoc) throw new Error(`Academic Year "${academicYearStr}" not found.`);

  // ── Parse rows (data starts at row 17 in the direct template, index 16) ───
  const rawRows = xlsx.utils.sheet_to_json<any[]>(sheet, { header: 1, range: 15 });

  for (const [index, row] of rawRows.entries()) {
    const regNo = row[1]?.toString().trim().toUpperCase();
    if (!regNo) continue;

    result.total++;
    const session = await mongoose.startSession();

    try {
      await session.withTransaction(async () => {
        const student = await Student.findOne({ regNo, institution: institutionId }).lean();
        if (!student) throw new Error(`Student ${regNo} not found`);

        const programUnit = await ProgramUnit.findOne({
          institution: institutionId,
          program:     (student as any).program,
          unit:        (unitDoc as any)._id,
        }).lean();

        if (!programUnit) {
          throw new Error(
            `Unit ${unitCode} is not in the curriculum for ${regNo}'s programme (Program ID: ${(student as any).program}).`,
          );
        }

        // ── Attempt detection ────────────────────────────────────────────
        const rawAttempt   = row[3]; // Column D = index 3
        const attempt      = detectAttemptType(rawAttempt);
        const isSpecial    = attempt === "special";
        const isSupp       = attempt === "supplementary";
        const isRetake     = attempt === "re-take";

        // ── Column mapping (directTemplate layout) ───────────────────────
        //  A=0  S/N
        //  B=1  REG. NO.
        //  C=2  NAME
        //  D=3  ATTEMPT
        //  E=4  CA TOTAL (/30)
        //  F=5  EXAM TOTAL (/70)
        //  G=6  INTERNAL (/100)   [formula — usually ignored on import]
        //  H=7  EXTERNAL (/100)
        //  I=8  AGREED (/100)
        //  J=9  GRADE

        const caTotal30      = Number(row[4]) || 0;
        const examTotal70    = Number(row[5]) || 0;
        const externalTotal  = row[7] !== undefined && row[7] !== null ? Number(row[7]) : null;
        const agreedMark     = Number(row[8]) || 0;

        const markData = {
          institution:      institutionId,
          student:          (student as any)._id,
          programUnit:      (programUnit as any)._id,
          academicYear:     (academicYearDoc as any)._id,
          caTotal30,
          examTotal70,
          externalTotal100: externalTotal,
          agreedMark,
          attempt,
          isSpecial,
          isSupplementary:  isSupp,
          isRetake,
          uploadedBy:       req.user._id,
        };

        await MarkDirect.findOneAndUpdate(
          {
            student:      (student as any)._id,
            programUnit:  (programUnit as any)._id,
            academicYear: (academicYearDoc as any)._id,
          },
          markData,
          { upsert: true, new: true, session },
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