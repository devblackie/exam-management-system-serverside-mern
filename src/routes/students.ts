// src/routes/students.ts
import { Router, Response } from "express";
import ExcelJS from "exceljs";
import { logAudit } from "../lib/auditLogger";
import mongoose from "mongoose";
import { normalizeProgramName } from "../services/programNormalizer";
import Student from "../models/Student";
import Program from "../models/Program";
import AcademicYear from "../models/AcademicYear";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import config from "../config/config";

const router = Router();

// GET all students
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const students = await Student.find({ institution: req.user.institution })
      .select("regNo name program admissionAcademicYear currentYearOfStudy")
      .populate("program", "name code")
      .lean();

    res.json(students);
  })
);

// GET student statistics for dashboard
router.get(
  "/stats",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution;

    const stats = await Student.aggregate([
      // 1. Filter students by the coordinator's institution
      { $match: { institution: institutionId } },

      // 2. Group by status and count
      {
        $group: {
          _id: "$status",
          count: { $sum: 1 },
        },
      },
    ]);

    // Format the result into a clean object (e.g., { active: 100, inactive: 5, total: 105 })
    let active = 0;
    let inactive = 0;
    let total = 0;

    for (const stat of stats) {
      total += stat.count;
      if (stat._id === "active") {
        active = stat.count;
      }
      // Sum all non-active statuses into 'inactive' for the dashboard display
      if (stat._id !== "active") {
        inactive += stat.count;
      }
    }

    res.json({
      active,
      inactive, // This will include graduated, suspended, deferred, and true inactive
      total,
    });
  })
);

router.get(
  "/template",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, academicYearId } = req.query;

    // ── Data Fetching Logic ───────────────────────────────────────────
    let programFilter = {};
    let academicYearFilter = {};

    if (programId && mongoose.Types.ObjectId.isValid(programId as string)) {
      programFilter = { _id: new mongoose.Types.ObjectId(programId as string) };
    }

    if (academicYearId && mongoose.Types.ObjectId.isValid(academicYearId as string)) {
      academicYearFilter = { _id: new mongoose.Types.ObjectId(academicYearId as string) };
    }

    let programs: any[] = [];
    let selectedProgram: any = null;

    if (Object.keys(programFilter).length > 0) {
      selectedProgram = await Program.findOne(programFilter).select("code name").lean();
      if (selectedProgram) programs = [selectedProgram];
    } else {
      programs = await Program.find({ institution: req.user.institution }).select("code name").lean();
    }

    let currentYearDoc: any = await AcademicYear.findOne(academicYearFilter || {
      institution: req.user.institution,
      isCurrent: true,
    }).select("year").lean();


    const currentYearString = currentYearDoc?.year || "General";

    // ── Create Workbook & Sheet ───────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Registration");
    const fontName = "Book Antiqua";

    // ── Institution Headers (Style from Scoresheet) ───────────────────
    const centerBold = {
      alignment: { horizontal: "center" as const, vertical: "middle" as const },
      font: { bold: true, name: fontName },
    };

    // Row 1: Institution Name (Using config or req.user.institution name)
    worksheet.mergeCells("A1:D1");
    const instCell = worksheet.getCell("A1");
   
    instCell.value = config.instName.toUpperCase(); 
    instCell.style = { ...centerBold, font: { ...centerBold.font, size: 14, underline: true } };

    // Row 2: Program Header
    worksheet.mergeCells("A2:D2");
    const progCell = worksheet.getCell("A2");
    progCell.value = selectedProgram 
      ? `PROGRAM: ${selectedProgram.code} - ${selectedProgram.name.toUpperCase()}`
      : "PROGRAM: ALL PROGRAMS (Select from dropdown)";
    progCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

    // Row 3: Academic Year
    worksheet.mergeCells("A3:D3");
    const yearCellHeader = worksheet.getCell("A3");
    yearCellHeader.value = `REGISTRATION TEMPLATE - ${currentYearString} ACADEMIC YEAR`;
    yearCellHeader.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

    worksheet.addRow([]); // Spacer at row 4

    // ── Table Headers (Row 5) ──────────────────────────────────────────
    const headerRowNum = 5;
    const headers = ["Reg No *", "Full Name *", "Program *", "Year of Study *"];
    const headerRow = worksheet.getRow(headerRowNum);

    headers.forEach((header, idx) => {
      const cell = headerRow.getCell(idx + 1);
      cell.value = header;
      cell.style = {
        font: { bold: true, name: fontName, color: { argb: "FFFFFFFF" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } },
        alignment: { horizontal: "center", vertical: "middle" },
        border: { 
          top: { style: "thin" }, 
          left: { style: "thin" }, 
          bottom: { style: "thin" }, 
          right: { style: "thin" } 
        }
      };
      cell.protection = { locked: true };
    });
    headerRow.height = 25;

    // ── Data Rows & Protection ────────────────────────────────────────
    const dataStartRow = 6;
    const maxRows = 500;
    const fixedProgramValue = selectedProgram ? `${selectedProgram.name}` : "";

    for (let r = dataStartRow; r <= dataStartRow + maxRows; r++) {
      const row = worksheet.getRow(r);
      row.font = { name: fontName, size: 10 };

      // Apply borders to the 4 columns
      for (let c = 1; c <= 4; c++) {
        row.getCell(c).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }

      // Column A & B (Reg No & Name): Always UNLOCKED
      row.getCell(1).protection = { locked: false };
      row.getCell(2).protection = { locked: false };

      // Column C (Program)
      const progDataCell = row.getCell(3);
      if (selectedProgram) {
        progDataCell.value = fixedProgramValue;
        progDataCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F4F6" } };
        progDataCell.protection = { locked: true }; // Pre-filled, so lock it
      } else {
        progDataCell.protection = { locked: false }; // Let them select
        if (programs.length > 0) {
          const programOptions = programs.map((p) => `${p.name}`).join(",");
          progDataCell.dataValidation = {
            type: "list",
            allowBlank: false,
            formulae: [`"${programOptions.substring(0, 250)}"`], // String length limit safety
          };
        }
      }

      // Column D (Year of Study): Always UNLOCKED
      const yearDataCell = row.getCell(4);
      yearDataCell.protection = { locked: false };
      yearDataCell.dataValidation = {
        type: "list",
        allowBlank: false,
        formulae: ['"1,2,3,4,5,6"'],
      };
    }

    // Auto-size columns
    worksheet.columns = [
      { width: 25 }, // Reg No
      { width: 40 }, // Full Name
      { width: 50 }, // Program
      { width: 20 }, // Year of Study
    ];

    worksheet.views = [{ state: "frozen", ySplit: headerRowNum }];

    // ── Sheet Protection ──────────────────────────────────────────────
    worksheet.protect("", {
      selectLockedCells: true,
      selectUnlockedCells: true,
      formatCells: true,
      formatColumns: false,
      formatRows: false,
      insertRows: false,
      deleteRows: false,
    });

    // ── Send File ─────────────────────────────────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();
    const safeYear = currentYearString.replace("/", "-");

    const cleanName =
        `${selectedProgram?.name}`
          .replace(/[^a-zA-Z0-9]/g, "_") // Replace anything not a letter or number
          .replace(/_+/g, "_") // Collapse multiple underscores (___ -> _)
          .replace(/^_|_$/g, "") // Remove _ from start or end
          ?.toUpperCase() || "TEMPLATE";
    // const filename = `Registration_Template_${safeYear}.xlsx`;
    const filename = `Registration_Template_${cleanName}.xlsx`;

    res
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Access-Control-Expose-Headers", "Content-Disposition")
      .attachment(filename)
      .send(Buffer.from(buffer as any));

    await logAudit(req, {
      action: "template_download",
      details: { type: "student_registration", programId },
    });
  })
);

// BULK register students — NO DUPLICATES
router.post(
  "/bulk",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { students } = req.body;
    if (!Array.isArray(students) || students.length === 0) {
      return res.status(400).json({ message: "No students provided" });
    }

    const institutionId = req.user.institution;

    // CLEAN & NORMALIZE INPUT
    const incoming = students.map((s) => ({
      regNo: s.regNo?.trim().toUpperCase(),
      name: s.name?.trim(),
      rawProgram: s.program?.trim(),
      normalizedProgram: normalizeProgramName(s.program?.trim() || ""),
      yearOfStudy: Number(s.yearOfStudy) || 1,
     academicYearId: s.academicYearId, 
      admissionAcademicYearString: s.admissionAcademicYear || "2024/2025",
    }));

    // Validate required fields (Reg No, Name, Program) - unchanged
    const invalid = incoming.filter(
      (s) => !s.regNo || !s.name || !s.rawProgram
    );
    if (invalid.length > 0) {
      return res
        .status(400)
        .json({ message: "Missing Reg No, Name, or Program" });
    }

    // --- STEP 1: LOOKUP PROGRAM IDs (Unchanged) ---
  

    // Get unique normalized program names
    const normNames = [...new Set(incoming.map((s) => s.normalizedProgram))];
    // Fetch all programs for this institution
    const programs = await Program.find({ institution: institutionId }).lean();
    // Build a normalized map
    const programMap = new Map<string, string>();
    for (const p of programs) {
      const norm = normalizeProgramName(p.name);
      programMap.set(norm, p._id.toString());
    }

    // Identify missing programs
    const missingPrograms: any[] = [];
    for (const n of normNames) {
      if (!programMap.has(n)) {
        missingPrograms.push(n);
      }
    }
    if (missingPrograms.length > 0) {
      return res.status(400).json({
        message: "Some programs not found",
        notFound: missingPrograms,
      });
    }

    // --- STEP 3: LOOKUP ACADEMIC YEAR IDs (NEW LOGIC) ---

    // Logic: If frontend sent academicYearId, we use it. If not, we find/create by string.
    const academicYearMap = new Map<string, mongoose.Types.ObjectId>();

    // Filter out rows that don't have an ID already and need string resolution
    const yearsToResolve = incoming
      .filter(s => !s.academicYearId)
      .map(s => s.admissionAcademicYearString);

const uniqueYearStrings = [...new Set(yearsToResolve)];

      // Use insertMany (or bulkWrite) for efficiency
     if (uniqueYearStrings.length > 0) {
      const bulkOps = uniqueYearStrings.map((yearStr) => {
        const [startYear, endYear] = yearStr.split("/").map(Number);
        return {
          updateOne: {
            filter: { year: yearStr, institution: institutionId },
            update: {
              $setOnInsert: {
                year: yearStr,
                institution: institutionId,
                startDate: new Date(`${startYear}-08-01`),
                endDate: new Date(`${endYear}-07-31`),
                isCurrent: false,
              },
            },
            upsert: true,
          },
        };
      });

          await AcademicYear.bulkWrite(bulkOps);
      const resolvedYears = await AcademicYear.find({
        institution: institutionId,
        year: { $in: uniqueYearStrings },
      }).lean();

      for (const y of resolvedYears) {
        academicYearMap.set(y.year, y._id as mongoose.Types.ObjectId);
      }
    }

    // --- STEP 4: DETECT EXISTING STUDENTS (Unchanged) ---
    const regNos = incoming.map((s) => s.regNo);
    const existing = await Student.find({
      regNo: { $in: regNos },
      institution: institutionId,
    }).select("regNo").lean();

    const existingRegNos = new Set(existing.map((s) => s.regNo));

    // --- STEP 5: BUILD FINAL PAYLOAD  ---
    const toCreate = incoming
      .filter((s) => !existingRegNos.has(s.regNo)) // Only include non-existing students
      .map((s) => {
        // Resolve ID: Use academicYearId if provided, else look up in the map
        let finalYearId: mongoose.Types.ObjectId;
        
        if (s.academicYearId && mongoose.Types.ObjectId.isValid(s.academicYearId)) {
          finalYearId = new mongoose.Types.ObjectId(s.academicYearId);
        } else {
          finalYearId = academicYearMap.get(s.admissionAcademicYearString)!;
        }

        return {
          regNo: s.regNo,
          name: s.name,
          program: programMap.get(s.normalizedProgram)!,
          currentYearOfStudy: s.yearOfStudy,
          admissionAcademicYear: finalYearId,
          institution: institutionId,
          status: "active",
        };
  });

    // --- STEP 6: INSERT AND RESPOND (Unchanged) ---
    if (toCreate.length > 0) {
      await Student.insertMany(toCreate);

      await logAudit(req, {
        action: "bulk_student_registration",
        details: { count: toCreate.length }
      });

      return res.status(200).json({
        message: `${toCreate.length} students registered successfully.`,
        registered: toCreate.map((s) => s.regNo),
        alreadyRegistered: Array.from(existingRegNos),
      });
    }

    return res.status(200).json({
      message: "All students in the list are already registered.",
      alreadyRegistered: Array.from(existingRegNos),
    });
  })
);

export default router;
