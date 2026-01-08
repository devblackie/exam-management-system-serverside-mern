// src/services/marksImporter.ts
import papa from "papaparse";
import xlsx from "xlsx";
import mongoose, { Types } from "mongoose";
import Student from "../models/Student";
// import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
import { computeFinalGrade } from "./gradeCalculator";
import { logAudit } from "../lib/auditLogger";
import type { AuthenticatedRequest } from "../middleware/auth";
import { MARKS_UPLOAD_HEADERS } from "../utils/uploadTemplate";
import Mark from "../models/Mark";

interface ImportRow {
  RegNo: string;
  UnitCode: string;
  CAT1?: string | number;
  CAT2?: string | number;
  CAT3?: string | number;
  Assignment?: string | number;
  Practical?: string | number;
  Exam?: string | number;
  IsSupplementary?: string;
  AcademicYear: string;
}

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

  const result: ImportResult = { total: 0, success: 0, errors: [], warnings: [] };

  // NO SESSION FOR LOOKUPS → prevents auto-abort
  let students: any[] = [];
  // let units: any[] = [];
  let academicYears: any[] = [];

  try {
    // 1. Parse file (unchanged)
    let rows: ImportRow[] = [];
    if (filename.toLowerCase().endsWith(".csv")) {
      const text = buffer.toString("utf-8");
      const parsed = papa.parse<ImportRow>(text, {
        header: true,
        skipEmptyLines: true,
        transformHeader: (h) => h.trim(),
        transform: (v) => v.trim(),
      });
      if (parsed.errors.length) throw new Error(`CSV parse error: ${parsed.errors[0].message}`);
      rows = parsed.data;
    } else {
      const workbook = xlsx.read(buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = xlsx.utils.sheet_to_json<ImportRow>(sheet, { defval: "" });
    }

    if (rows.length === 0) throw new Error("File is empty");

    // 2. Validate headers
    const headers = Object.keys(rows[0]);
    const missing = MARKS_UPLOAD_HEADERS.filter((h) => !headers.includes(h));
    if (missing.length) throw new Error(`Missing columns: ${missing.join(", ")}`);

    // 3. PRE-LOAD ALL REFERENCE DATA OUTSIDE ANY TRANSACTION
    const regNos = rows.map((r) => r.RegNo.toUpperCase());
    const unitCodes = [...new Set(rows.map((r) => r.UnitCode.toUpperCase()))];
    const years = [...new Set(rows.map((r) => r.AcademicYear))];

    // Get all students, all relevant academic years, and ALL relevant ProgramUnits
    const [studentsData, academicYearsData, programUnits] = await Promise.all([ // ⬅️ UPDATED PROMISE.ALL
      Student.find({ regNo: { $in: regNos }, institution: institutionId }).lean(),
      AcademicYear.find({ year: { $in: years }, institution: institutionId }).lean(),
      // ⬅️ CRITICAL: Find all curriculum links (ProgramUnits) that match the Unit Codes and Programs of the students.
      // We will need to figure out which programs are involved based on the students found.
      // For simplicity here, we fetch ALL ProgramUnits for all programs linked to these units, and filter later.
      // **NOTE:** This assumes ProgramUnit has been populated with Unit during creation/linking.
      ProgramUnit.find({ institution: institutionId }).populate('unit').lean(), 
    ]);

    students = studentsData;
    academicYears = academicYearsData;

 const studentMap = new Map(students.map((s) => [s.regNo.toUpperCase(), s]));
    // const unitMap = new Map(units.map((u) => [u.code.toUpperCase(), u])); // ⬅️ REMOVED
    const yearMap = new Map(academicYears.map((y) => [y.year, y]));

    // 4. NOW start a transaction ONLY for writes
    const session = await mongoose.startSession();
    let transactionAborted = false;

    try {
      session.startTransaction();

      // This helps reduce queries inside the loop, though the lookup inside the loop is still complex.
      const programUnitMap = new Map();
      programUnits.forEach((pu: any) => {
          // Key format: PROGRAMID_UNITCODE
          const key = `${pu.program.toString()}_${pu.unit.code.toUpperCase()}`;
          programUnitMap.set(key, pu);
      });

      for (const [index, row] of rows.entries()) {
        result.total++;
        const rowNum = index + 2;

        try {
          const regNo = row.RegNo.trim().toUpperCase();
          const unitCode = row.UnitCode.trim().toUpperCase();
          const academicYearStr = row.AcademicYear.trim();

          const student = studentMap.get(regNo);
          // const unit = unitMap.get(unitCode);
          const academicYear = yearMap.get(academicYearStr);

          if (!student) throw new Error(`Student not found: ${regNo}`);
          // if (!unit) throw new Error(`Unit not found: ${unitCode}`);
          if (!academicYear) throw new Error(`Academic year not found: ${academicYearStr}`);

          const programUnitKey = `${student.program.toString()}_${unitCode}`;
          
          // Assign the looked-up value to a local variable
          const programUnit = programUnitMap.get(programUnitKey);
          
if (!programUnit) {
              throw new Error(`Curriculum link (ProgramUnit) not found for UnitCode ${unitCode} in Student's Program.`);
          }

          const markData = {
            student: student._id,
            programUnit: programUnit._id,
            academicYear: academicYear._id,
            institution: institutionId,
            uploadedBy: req.user._id,
            isSupplementary: String(row.IsSupplementary || "NO").toUpperCase() === "YES",
            cat1: row.CAT1 ? Number(row.CAT1) : undefined,
            cat2: row.CAT2 ? Number(row.CAT2) : undefined,
            cat3: row.CAT3 ? Number(row.CAT3) : undefined,
            assignment: row.Assignment ? Number(row.Assignment) : undefined,
            practical: row.Practical ? Number(row.Practical) : undefined,
            exam: row.Exam !== undefined ? Number(row.Exam) : undefined,
          };

         const mark = await Mark.findOneAndUpdate(
            {
              student: student._id,
              programUnit: programUnit._id, // ⬅️ UPDATED: Query Mark by programUnit
              academicYear: academicYear._id,
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
          // Keep going – we don’t abort the whole import
        }
      }

      await session.commitTransaction();
    } catch (transactionErr: any) {
      transactionAborted = true;
      await session.abortTransaction();
      throw transactionErr; // re-throw so caller sees 500
    } finally {
      session.endSession();
    }

    // 5. Audit log
    await logAudit(req, {
      action: "marks_bulk_import_completed",
      details: {
        file: filename,
        totalRows: result.total,
        successful: result.success,
        failed: result.errors.length,
      },
    });

    return result;
  } catch (err: any) {
    await logAudit(req, {
      action: "marks_bulk_import_failed",
      details: { error: err.message, file: filename },
    });
    throw err;
  }
}