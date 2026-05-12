
// // src/routes/students.ts
// import { Router, Response } from "express";
// import ExcelJS from "exceljs";
// import { logAudit } from "../lib/auditLogger";
// import mongoose from "mongoose";
// import { normalizeProgramName } from "../services/programNormalizer";
// import Student from "../models/Student";
// import Program from "../models/Program";
// import AcademicYear from "../models/AcademicYear";
// import { requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import type { AuthenticatedRequest } from "../middleware/auth";
// import config from "../config/config";
// import { paginate } from "../utils/paginate";
// import { loadInstitutionSettings } from "../utils/loadInstitutionSettings";

// const router = Router();

// // GET all students


// // GET /students?page=1&limit=20
// router.get("/", requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const page  = Math.max(1, parseInt(req.query.page  as string) || 1);
//     const limit = Math.max(1, parseInt(req.query.limit as string) || 20);

//     const filter: Record<string, any> = {institution: req.user.institution};

//     // Only show search results — never dump all students
//     const search = (req.query.search as string)?.trim();
//     if (!search) {
//       // Return empty until user types something
//       res.json({ students: [], total: 0, page, totalPages: 0 });
//       return;
//     }

//     filter.$or = [
//       { regNo: { $regex: search, $options: "i" } },
//       { name:  { $regex: search, $options: "i" } },
//     ];

//     const [students, total] = await Promise.all([
//       paginate(
//         Student.find(filter).select("regNo name program currentYearOfStudy status").lean(),
//         page,
//         limit,
//       ),
//       Student.countDocuments(filter),
//     ]);

//     res.json({ students, total, page, totalPages: Math.ceil(total / limit)});
//   }),
// );

// // router.get("/", requireAuth,
// //   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
// //     const students = await Student.find({ institution: req.user.institution })
// //       .select("regNo name program admissionAcademicYear currentYearOfStudy")
// //       .populate("program", "name code")
// //       .lean();

// //     res.json(students);
// //   })
// // );

// // GET student statistics for dashboard
// router.get("/stats", requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const institutionId = req.user.institution;

//     const stats = await Student.aggregate([
//       // 1. Filter students by the coordinator's institution
//       { $match: { institution: institutionId } },

//       // 2. Group by status and count
//       {
//         $group: {
//           _id: "$status",
//           count: { $sum: 1 },
//         },
//       },
//     ]);

//     // Format the result into a clean object (e.g., { active: 100, inactive: 5, total: 105 })
//     let active = 0;
//     let inactive = 0;
//     let total = 0;

//     for (const stat of stats) {
//       total += stat.count;
//       if (stat._id === "active") {
//         active = stat.count;
//       }
//       // Sum all non-active statuses into 'inactive' for the dashboard display
//       if (stat._id !== "active") {
//         inactive += stat.count;
//       }
//     }

//     res.json({
//       active,
//       inactive, // This will include graduated, suspended, deferred, and true inactive
//       total,
//     });
//   })
// );


// router.get(
//   "/template",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { programId, academicYearId } = req.query;

//     // ── Data Fetching Logic ───────────────────────────────────────────
//     let programFilter = {};
//     let academicYearFilter = {};

//     if (programId && mongoose.Types.ObjectId.isValid(programId as string)) {
//       programFilter = { _id: new mongoose.Types.ObjectId(programId as string) };
//     }

//     if (academicYearId && mongoose.Types.ObjectId.isValid(academicYearId as string)) {
//       academicYearFilter = { _id: new mongoose.Types.ObjectId(academicYearId as string) };
//     }

//     let programs: any[] = [];
//     let selectedProgram: any = null;

//     if (Object.keys(programFilter).length > 0) {
//       selectedProgram = await Program.findOne(programFilter).select("code name").lean();
//       if (selectedProgram) programs = [selectedProgram];
//     } else {
//       programs = await Program.find({ institution: req.user.institution }).select("code name").lean();
//     }

//     let currentYearDoc: any = await AcademicYear.findOne(academicYearFilter || {
//       institution: req.user.institution,
//       isCurrent: true,
//     }).select("year").lean();


//     const currentYearString = currentYearDoc?.year || "General";

//     // ── Create Workbook & Sheet ───────────────────────────────────────
//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet("Registration");
//     const fontName = "Book Antiqua";

//     // ── Institution Headers (Style from Scoresheet) ───────────────────
//     const centerBold = {
//       alignment: { horizontal: "center" as const, vertical: "middle" as const },
//       font: { bold: true, name: fontName },
//     };

//     // Row 1: Institution Name (Using config or req.user.institution name)
//     worksheet.mergeCells("A1:D1");
//     const instCell = worksheet.getCell("A1");
   
//     // instCell.value = config.instName.toUpperCase(); 
//     const settings = await loadInstitutionSettings(
//       req.user.institution.toString(),
//     );
//     instCell.value = settings.docMeta.universityName.toUpperCase();
//     instCell.style = { ...centerBold, font: { ...centerBold.font, size: 14, underline: true } };

//     // Row 2: Program Header
//     worksheet.mergeCells("A2:D2");
//     const progCell = worksheet.getCell("A2");
//     progCell.value = selectedProgram 
//       ? `PROGRAM: ${selectedProgram.code} - ${selectedProgram.name.toUpperCase()}`
//       : "PROGRAM: ALL PROGRAMS (Select from dropdown)";
//     progCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     // Row 3: Academic Year
//     worksheet.mergeCells("A3:D3");
//     const yearCellHeader = worksheet.getCell("A3");
//     yearCellHeader.value = `REGISTRATION TEMPLATE - ${currentYearString} ACADEMIC YEAR`;
//     yearCellHeader.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     worksheet.addRow([]); // Spacer at row 4

//     // ── Table Headers (Row 5) ──────────────────────────────────────────
//     const headerRowNum = 5;
//     const headers = ["Reg No", "Full Name", "Program", "Year of Study", "Intake"];
//     const headerRow = worksheet.getRow(headerRowNum);

//     headers.forEach((header, idx) => {
//       const cell = headerRow.getCell(idx + 1);
//       cell.value = header;
//       cell.style = {
//         font: { bold: true, name: fontName, color: { argb: "FFFFFFFF" } },
//         fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } },
//         alignment: { horizontal: "center", vertical: "middle" },
//         border: { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } }
//       };
//       cell.protection = { locked: true };
//     });
//     headerRow.height = 25;

//     // ── Data Rows & Protection ────────────────────────────────────────
//     const dataStartRow = 6;
//     const maxRows = 500;
//     const fixedProgramValue = selectedProgram ? `${selectedProgram.name}` : "";

//     for (let r = dataStartRow; r <= dataStartRow + maxRows; r++) {
//       const row = worksheet.getRow(r);
//       row.font = { name: fontName, size: 10 };

//       // Apply borders to the 4 columns
//       for (let c = 1; c <= 4; c++) {
//         row.getCell(c).border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }};
//       }

//       // Column A & B (Reg No & Name): Always UNLOCKED
//       row.getCell(1).protection = { locked: false };
//       row.getCell(2).protection = { locked: false };
//       row.getCell(5).protection = { locked: false };
      

//       // Column C (Program)
//       const progDataCell = row.getCell(3);
//       if (selectedProgram) {
//         progDataCell.value = fixedProgramValue;
//         progDataCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F4F6" }};
//         progDataCell.protection = { locked: true }; // Pre-filled, so lock it
//       } else {
//         progDataCell.protection = { locked: false }; // Let them select
//         if (programs.length > 0) {
//           const programOptions = programs.map((p) => `${p.name}`).join(",");
//           progDataCell.dataValidation = {
//             type: "list",
//             allowBlank: false,
//             formulae: [`"${programOptions.substring(0, 250)}"`], // String length limit safety
//           };
//         }
//       }

//       // Column D (Year of Study): Always UNLOCKED
//       const yearDataCell = row.getCell(4);
//       yearDataCell.protection = { locked: false };
//       yearDataCell.dataValidation = { type: "list", allowBlank: false, formulae: ['"1,2,3,4,5,6"']};

//       // --- COLUMN E: INTAKE SELECTION ---
//       const intakeCell = row.getCell(5);
//       intakeCell.protection = { locked: false };
//       intakeCell.style = {
//         font: { name: "Book Antiqua", size: 10 },
//         border: { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }},
//       };

//       // This creates the dropdown menu in Excel
//       intakeCell.dataValidation = {
//         type: "list",
//         allowBlank: false,
//         formulae: ['"JAN,MAY,SEPT"'],
//         showErrorMessage: true,
//         errorTitle: "Invalid Intake",
//         error: "Please select an intake from the list (JAN, MAY, or SEPT)."
//       };
  
//       // Set default value for the template
//       intakeCell.value = "SEPT";
//     }

//     // Auto-size columns
//     worksheet.columns = [
//       { width: 25 }, // Reg No
//       { width: 40 }, // Full Name
//       { width: 50 }, // Program
//       { width: 20 }, // Year of Study
//       { width: 15 }, // Intake
//     ];

//     worksheet.views = [{ state: "frozen", ySplit: headerRowNum }];

//     // ── Sheet Protection ──────────────────────────────────────────────
//     worksheet.protect("", {
//       selectLockedCells: true,
//       selectUnlockedCells: true,
//       formatCells: true,
//       formatColumns: false,
//       formatRows: false,
//       insertRows: false,
//       deleteRows: false,
//     });

//     // ── Send File ─────────────────────────────────────────────────────
//     const buffer = await workbook.xlsx.writeBuffer();
//     const safeYear = currentYearString.replace("/", "-");

//     const cleanName =
//         `${selectedProgram?.name}`
//           .replace(/[^a-zA-Z0-9]/g, "_") // Replace anything not a letter or number
//           .replace(/_+/g, "_") // Collapse multiple underscores (___ -> _)
//           .replace(/^_|_$/g, "") // Remove _ from start or end
//           ?.toUpperCase() || "TEMPLATE";
//     // const filename = `Registration_Template_${safeYear}.xlsx`;
//     const filename = `Registration_Template_${cleanName}.xlsx`;

//     res
//       .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
//       .header("Access-Control-Expose-Headers", "Content-Disposition")
//       .attachment(filename)
//       .send(Buffer.from(buffer as any));

//     await logAudit(req, {
//       action: "template_download",
//       details: { type: "student_registration", programId },
//     });
//   })
// );

// // BULK register students — NO DUPLICATES
// router.post(
//   "/bulk",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { students } = req.body;
//     if (!Array.isArray(students) || students.length === 0) return res.status(400).json({ message: "No students provided" });
    
//     const institutionId = req.user.institution;

//     // 1. CLEAN & NORMALIZE INPUT
//     const incoming = students.map((s) => ({
//       regNo: s.regNo?.trim().toUpperCase(),
//       name: s.name?.trim(),
//       rawProgram: s.program?.trim(),
//       normalizedProgram: normalizeProgramName(s.program?.trim() || ""),
//       yearOfStudy: Number(s.currentYearOfStudy) || 1, // Note: using currentYearOfStudy from frontend
//       academicYearId: s.academicYearId,
//       intake: s.intake?.trim().toUpperCase() || "JAN",
//       admissionAcademicYearString: s.admissionAcademicYearString || "2024/2025",
//     }));

//     // Validate required fields (Reg No, Name, Program) - unchanged
//     const invalid = incoming.filter((s) => !s.regNo || !s.name || !s.rawProgram);
//     if (invalid.length > 0) return res.status(400).json({ message: "Missing Reg No, Name, or Program" });
    
//     // 2. LOOKUP PROGRAMS (Retrieve Full Objects, not just IDs)
//     const normNames = [...new Set(incoming.map((s) => s.normalizedProgram))];
//     const programs = await Program.find({ institution: institutionId }).lean();
//     const programNameMap = new Map(
//       programs.map((p) => [normalizeProgramName(p.name), p])
//     );
//     const programIdMap = new Map(
//       programs.map((p) => [p._id.toString(), p])
//     );

//     // Identify missing programs
//     const missingPrograms = [...new Set(incoming.map(s => s.rawProgram))]
//       .filter(raw => !programIdMap.has(raw) && !programNameMap.has(normalizeProgramName(raw)));

//     if (missingPrograms.length > 0) return res.status(400).json({ message: "Programs not found", notFound: missingPrograms });
    
//     // 3. RESOLVE ACADEMIC YEARS
//     const academicYearMap = new Map<string, mongoose.Types.ObjectId>();
//     const yearsToResolve = incoming.filter(s => !s.academicYearId).map(s => s.admissionAcademicYearString);
//     const uniqueYearStrings = [...new Set(yearsToResolve)];

//       // Use insertMany (or bulkWrite) for efficiency
//      if (uniqueYearStrings.length > 0) {
//       const bulkOps = uniqueYearStrings.map((yearStr) => {
//         const [startYear, endYear] = yearStr.split("/").map(Number);
//         return {
//           updateOne: {
//             filter: { year: yearStr, institution: institutionId },
//             update: { $setOnInsert: { year: yearStr, institution: institutionId, startDate: new Date(`${startYear}-08-01`), endDate: new Date(`${endYear}-07-31`), isCurrent: false }},
//             upsert: true,
//           },
//         };
//       });
//       await AcademicYear.bulkWrite(bulkOps);
//       const resolvedYears = await AcademicYear.find({ institution: institutionId, year: { $in: uniqueYearStrings }}).lean();
//       resolvedYears.forEach(y => academicYearMap.set(y.year, y._id as mongoose.Types.ObjectId));
//     }

//     // 4. DETECT EXISTING STUDENTS
//     const regNos = incoming.map((s) => s.regNo);
//     const existing = await Student.find({ regNo: { $in: regNos }, institution: institutionId }).select("regNo").lean();
//     const existingRegNos = new Set(existing.map((s) => s.regNo));

//     // --- STEP 5: BUILD FINAL PAYLOAD  ---
//     const toCreate = incoming
//       .filter((s) => !existingRegNos.has(s.regNo)) // Only include non-existing students
//       .map((s) => {
//         const progDoc = programIdMap.get(s.rawProgram) || programNameMap.get(s.normalizedProgram)!;
        
//         // A. Program Type (B.Sc vs B.Ed vs Diploma)
//         let pType = "B.Sc";
//         if (progDoc.name.toLowerCase().includes("education")) pType = "B.Ed";
//         else if (progDoc.name.toLowerCase().includes("diploma")) pType = "Diploma";

//         // B. Entry Type Detection (Direct vs Mid-Entry)
//         let eType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4" = "Direct";
//         if (s.yearOfStudy === 2) eType = "Mid-Entry-Y2";
//         else if (s.yearOfStudy === 3) eType = "Mid-Entry-Y3";
//         else if (s.yearOfStudy === 4) eType = "Mid-Entry-Y4";

//         // C. Academic Year ID
//         let finalYearId: mongoose.Types.ObjectId;
//         if (s.academicYearId && mongoose.Types.ObjectId.isValid(s.academicYearId)) finalYearId = new mongoose.Types.ObjectId(s.academicYearId);
//         else { finalYearId = academicYearMap.get(s.admissionAcademicYearString)!;}
        
//         return {
//           regNo: s.regNo, name: s.name, institution: institutionId, program: progDoc._id, programType: pType, entryType: eType, intake: s.intake,
//           currentYearOfStudy: s.yearOfStudy, admissionAcademicYear: finalYearId, status: "active", initialRegistrationDate: new Date(),
//         };
//   });

//   try {
//     const result = await Student.insertMany(toCreate, { ordered: false });

//     return res.status(200).json({
//       message: `${result.length} students registered successfully.`,
//       registered: result.map((r) => r.regNo),
//     });
    
//   } catch (error: any) {
//     const insertedDocs = error.insertedDocs || [];
//     const writeErrors = error.writeErrors || [];

//     const duplicateRegNos = writeErrors.filter((err: any) => err.code === 11000).map((err: any) => err.op.regNo);

//     return res.status(207).json({
//       // 207 Multi-Status
//       message: `${insertedDocs.length} registered, ${duplicateRegNos.length} skipped (duplicates).`,
//       registered: insertedDocs.map((d: any) => d.regNo),
//       alreadyRegistered: duplicateRegNos,
//     });
    
//   }
//   })
// );

// // A. DELETE SINGLE STUDENT
// router.delete("/:id", requireAuth, requireRole("admin","coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const student = await Student.findOneAndDelete({ 
//     _id: req.params.id, 
//     institution: req.user.institution 
//   });
//   if (!student) return res.status(404).json({ message: "Student not found" });

//   await logAudit(req, { action: "delete_student", details: { regNo: student.regNo } });
//   res.json({ message: "Student deleted successfully" });
// }));

// // serverside/src/routes/students.ts

// // UPDATE student (Only name allowed)
// router.patch("/:id", requireAuth, requireRole("admin", "coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const { name } = req.body;
//   if (!name) return res.status(400).json({ message: "Name is required" });

//   const student = await Student.findOneAndUpdate(
//     { _id: req.params.id, institution: req.user.institution },
//     { $set: { name: name.trim() } },
//     { new: true }
//   ).select("name regNo");

//   if (!student) return res.status(404).json({ message: "Student not found" });

//   await logAudit(req, { action: "update_student_name", details: { regNo: student.regNo, newName: name } });
//   res.json(student);
// }));

// // B. DELETE BY PROGRAM (e.g., if a program is decommissioned)
// router.delete("/bulk/by-program", requireAuth, requireRole("admin"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const { programId } = req.body;
//   if (!programId) return res.status(400).json({ message: "Program ID required" });

//   const result = await Student.deleteMany({ 
//     program: programId, 
//     institution: req.user.institution 
//   });

//   await logAudit(req, { action: "bulk_delete_program_students", details: { programId, count: result.deletedCount } });
//   res.json({ message: `Deleted ${result.deletedCount} students from program.` });
// }));

// // C. CLEANUP GRADUATED STUDENTS (Move to an archive or delete)
// // Useful for removing old records after 10 years (ENG 19.d)
// router.delete("/bulk/cleanup-graduated", requireAuth, requireRole("admin"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const result = await Student.deleteMany({ 
//     status: "graduated", 
//     institution: req.user.institution 
//   });

//   res.json({ message: `Cleanup complete. ${result.deletedCount} records removed.` });
// }));
// export default router;



















// // serverside/src/routes/students.ts — COMPLETE
// import { Router, Response } from "express";
// import ExcelJS       from "exceljs";
// import mongoose      from "mongoose";
// import Student       from "../models/Student";
// import Program       from "../models/Program";
// import AcademicYear  from "../models/AcademicYear";
// import { requireAuth, requireRole, AuthenticatedRequest, getScopedProgramIds } from "../middleware/auth";
// import { asyncHandler }           from "../middleware/asyncHandler";
// import { logAudit }               from "../lib/auditLogger";
// import { paginate }               from "../utils/paginate";
// import { normalizeProgramName }   from "../services/programNormalizer";
// import { validateRegNo }          from "../utils/validateRegNo";
// import { loadInstitutionSettings } from "../utils/loadInstitutionSettings";
// import { ApiError }               from "../middleware/errorHandler";

// const router = Router();

// // ── GET /students?search=&page= ───────────────────────────────────────────────
// router.get(
//   "/",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const page   = Math.max(1, parseInt(req.query.page  as string) || 1);
//     const limit  = Math.max(1, parseInt(req.query.limit as string) || 20);
//     const search = (req.query.search as string | undefined)?.trim() ?? "";

//     if (!search) {
//       res.json({ students: [], total: 0, page, totalPages: 0 });
//       return;
//     }

//     const allowedProgramIds = await getScopedProgramIds(req);

//     const filter: Record<string, unknown> = {
//       institution: req.user.institution,
//       program:     { $in: allowedProgramIds },
//       $or: [
//         { regNo: { $regex: search, $options: "i" } },
//         { name:  { $regex: search, $options: "i" } },
//       ],
//     };

//     const [students, total] = await Promise.all([
//       paginate(
//         Student.find(filter)
//           .select("regNo name program currentYearOfStudy status qualifierSuffix intake")
//           .populate("program", "name code departmentCode schoolCode")
//           .lean(),
//         page,
//         limit,
//       ),
//       Student.countDocuments(filter),
//     ]);

//     res.json({
//       students,
//       total,
//       page,
//       totalPages: Math.ceil(total / Math.min(100, limit)),
//     });
//   }),
// );

// // ── GET /students/stats ───────────────────────────────────────────────────────
// router.get(
//   "/stats",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const allowedProgramIds = await getScopedProgramIds(req);

//     const [total, active, graduated, discontinued] = await Promise.all([
//       Student.countDocuments({ institution: req.user.institution, program: { $in: allowedProgramIds } }),
//       Student.countDocuments({ institution: req.user.institution, program: { $in: allowedProgramIds }, status: "active" }),
//       Student.countDocuments({ institution: req.user.institution, program: { $in: allowedProgramIds }, status: "graduated" }),
//       Student.countDocuments({ institution: req.user.institution, program: { $in: allowedProgramIds }, status: "discontinued" }),
//     ]);

//     const inactive = total - active;
//     res.json({ total, active, inactive, graduated, discontinued });
//   }),
// );

// // ── GET /students/template ────────────────────────────────────────────────────
// router.get(
//   "/template",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { programId, academicYearId } = req.query;
//     const institutionId = req.user.institution.toString();

//     // Load institution settings for header
//     const settings = await loadInstitutionSettings(institutionId);
//     const universityName = settings.docMeta.universityName || "University";

//     // Scoped programs
//     const allowedProgramIds = await getScopedProgramIds(req);

//     let programs: Array<{ _id: mongoose.Types.ObjectId; code: string; name: string }> = [];
//     let selectedProgram: { _id: mongoose.Types.ObjectId; code: string; name: string } | null = null;

//     if (programId && mongoose.Types.ObjectId.isValid(programId as string)) {
//       const prog = await Program.findOne({
//         _id:         new mongoose.Types.ObjectId(programId as string),
//         institution: req.user.institution,
//         _id_in:      allowedProgramIds,
//       }).select("code name").lean() as typeof selectedProgram;
//       if (prog) { selectedProgram = prog; programs = [prog]; }
//     } else {
//       programs = await Program.find({
//         institution: req.user.institution,
//         _id:         { $in: allowedProgramIds },
//       }).select("code name").lean() as typeof programs;
//     }

//     const yearFilter = academicYearId && mongoose.Types.ObjectId.isValid(academicYearId as string)
//       ? { _id: new mongoose.Types.ObjectId(academicYearId as string) }
//       : { institution: req.user.institution, isCurrent: true };

//     const currentYearDoc = await AcademicYear.findOne(yearFilter).select("year").lean();
//     const currentYearString = (currentYearDoc as { year?: string } | null)?.year ?? "General";

//     // Build workbook
//     const workbook  = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet("Registration");
//     const fontName  = "Book Antiqua";

//     const centerBold = {
//       alignment: { horizontal: "center" as const, vertical: "middle" as const },
//       font: { bold: true, name: fontName },
//     };

//     worksheet.mergeCells("A1:D1");
//     const instCell   = worksheet.getCell("A1");
//     instCell.value   = universityName.toUpperCase();  // ← from DB not env
//     instCell.style   = { ...centerBold, font: { ...centerBold.font, size: 14, underline: true } };

//     worksheet.mergeCells("A2:D2");
//     const progCell   = worksheet.getCell("A2");
//     progCell.value   = selectedProgram
//       ? `PROGRAM: ${selectedProgram.code} - ${selectedProgram.name.toUpperCase()}`
//       : "PROGRAM: ALL PROGRAMS (Select from dropdown)";
//     progCell.style   = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     worksheet.mergeCells("A3:D3");
//     const yearCellH  = worksheet.getCell("A3");
//     yearCellH.value  = `REGISTRATION TEMPLATE - ${currentYearString} ACADEMIC YEAR`;
//     yearCellH.style  = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     worksheet.addRow([]);

//     const headerRow = worksheet.getRow(5);
//     ["Reg No","Full Name","Program","Year of Study","Intake"].forEach((header, idx) => {
//       const cell = headerRow.getCell(idx + 1);
//       cell.value = header;
//       cell.style = {
//         font:      { bold: true, name: fontName, color: { argb: "FFFFFFFF" } },
//         fill:      { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } },
//         alignment: { horizontal: "center", vertical: "middle" },
//         border:    { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} },
//       };
//     });
//     headerRow.height = 25;

//     const dataStartRow = 6;
//     const fixedProgram = selectedProgram ? selectedProgram.name : "";

//     for (let r = dataStartRow; r <= dataStartRow + 500; r++) {
//       const row = worksheet.getRow(r);
//       row.font  = { name: fontName, size: 10 };

//       for (let c = 1; c <= 5; c++) {
//         row.getCell(c).border = {
//           top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"},
//         };
//       }

//       row.getCell(1).protection = { locked: false };
//       row.getCell(2).protection = { locked: false };
//       row.getCell(5).protection = { locked: false };

//       const progCell2 = row.getCell(3);
//       if (selectedProgram) {
//         progCell2.value      = fixedProgram;
//         progCell2.fill       = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F4F6" } };
//         progCell2.protection = { locked: true };
//       } else {
//         progCell2.protection = { locked: false };
//         if (programs.length > 0) {
//           const opts = programs.map(p => p.name).join(",").substring(0, 250);
//           progCell2.dataValidation = { type: "list", allowBlank: false, formulae: [`"${opts}"`] };
//         }
//       }

//       const yearCell = row.getCell(4);
//       yearCell.protection  = { locked: false };
//       yearCell.dataValidation = { type:"list", allowBlank:false, formulae:['"1,2,3,4,5,6"'] };

//       const intakeCell = row.getCell(5);
//       intakeCell.protection = { locked: false };
//       intakeCell.value      = "SEPT";
//       intakeCell.dataValidation = {
//         type: "list", allowBlank: false, formulae: ['"JAN,MAY,SEPT"'],
//         showErrorMessage: true, errorTitle: "Invalid Intake",
//         error: "Please select JAN, MAY, or SEPT.",
//       };
//     }

//     worksheet.columns = [
//       { width: 25 }, { width: 40 }, { width: 50 }, { width: 20 }, { width: 15 },
//     ];
//     worksheet.views = [{ state: "frozen", ySplit: 5 }];
//     worksheet.protect("", { selectLockedCells: true, selectUnlockedCells: true });

//     const buffer     = await workbook.xlsx.writeBuffer();
//     const safeYear   = currentYearString.replace("/", "-");
//     const cleanName  = selectedProgram
//       ? selectedProgram.name.replace(/[^a-zA-Z0-9]/g,"_").replace(/_+/g,"_").toUpperCase()
//       : "ALL";
//     const filename   = `Registration_Template_${cleanName}_${safeYear}.xlsx`;

//     res
//       .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
//       .header("Access-Control-Expose-Headers", "Content-Disposition")
//       .attachment(filename)
//       .send(Buffer.from(buffer as ArrayBuffer));

//     await logAudit(req, { action: "template_download", details: { type: "student_registration", programId } });
//   }),
// );

// // ── POST /students/bulk ───────────────────────────────────────────────────────
// router.post(
//   "/bulk",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     interface BulkStudentRow {
//       regNo:                       string;
//       name:                        string;
//       program:                     string;
//       currentYearOfStudy?:         number;
//       academicYearId?:             string;
//       intake?:                     string;
//       admissionAcademicYearString?: string;
//     }

//     const { students } = req.body as { students: BulkStudentRow[] };
//     if (!Array.isArray(students) || students.length === 0) {
//       res.status(400).json({ message: "No students provided" }); return;
//     }

//     const institutionId     = req.user.institution;
//     const allowedProgramIds = await getScopedProgramIds(req);

//     const incoming = students.map(s => ({
//       regNo:                        (s.regNo ?? "").trim().toUpperCase(),
//       name:                         (s.name  ?? "").trim(),
//       rawProgram:                   (s.program ?? "").trim(),
//       normalizedProgram:            normalizeProgramName((s.program ?? "").trim()),
//       yearOfStudy:                  Number(s.currentYearOfStudy) || 1,
//       academicYearId:               s.academicYearId,
//       intake:                       (s.intake ?? "SEPT").trim().toUpperCase(),
//       admissionAcademicYearString:  s.admissionAcademicYearString ?? "2024/2025",
//     }));

//     const invalid = incoming.filter(s => !s.regNo || !s.name || !s.rawProgram);
//     if (invalid.length > 0) {
//       res.status(400).json({ message: "Missing Reg No, Name, or Program" }); return;
//     }

//     // Fetch scoped programs only
//     const dbPrograms = await Program.find({
//       institution: institutionId,
//       _id:         { $in: allowedProgramIds },
//     }).lean() as Array<{
//       _id: mongoose.Types.ObjectId; name: string; code: string;
//       durationYears: number; degreeType: string;
//     }>;

//     const programNameMap = new Map(
//       dbPrograms.map(p => [normalizeProgramName(p.name), p]),
//     );
//     const programIdMap = new Map(
//       dbPrograms.map(p => [p._id.toString(), p]),
//     );

//     const missingPrograms = [...new Set(incoming.map(s => s.rawProgram))].filter(
//       raw => !programIdMap.has(raw) && !programNameMap.has(normalizeProgramName(raw)),
//     );
//     if (missingPrograms.length > 0) {
//       res.status(400).json({ message: "Programs not found or not in your scope", notFound: missingPrograms });
//       return;
//     }

//     // Resolve academic years
//     const academicYearMap = new Map<string, mongoose.Types.ObjectId>();
//     const yearStrings     = [
//       ...new Set(incoming.filter(s => !s.academicYearId).map(s => s.admissionAcademicYearString)),
//     ];

//     if (yearStrings.length > 0) {
//       const bulkOps = yearStrings.map(yearStr => {
//         const [startYear, endYear] = yearStr.split("/").map(Number);
//         return {
//           updateOne: {
//             filter: { year: yearStr, institution: institutionId },
//             update: {
//               $setOnInsert: {
//                 year: yearStr, institution: institutionId,
//                 startDate: new Date(`${startYear}-08-01`),
//                 endDate:   new Date(`${endYear}-07-31`),
//                 isCurrent: false,
//               },
//             },
//             upsert: true,
//           },
//         };
//       });
//       await AcademicYear.bulkWrite(bulkOps);
//       const resolved = await AcademicYear.find({
//         institution: institutionId, year: { $in: yearStrings },
//       }).lean() as Array<{ _id: mongoose.Types.ObjectId; year: string }>;
//       resolved.forEach(y => academicYearMap.set(y.year, y._id));
//     }

//     const regNos    = incoming.map(s => s.regNo);
//     const existing  = await Student.find({
//       regNo: { $in: regNos }, institution: institutionId,
//     }).select("regNo").lean() as Array<{ regNo: string }>;
//     const existingSet = new Set(existing.map(s => s.regNo));

//     const results: { registered: string[]; duplicates: string[]; errors: string[] } = {
//       registered: [], duplicates: [], errors: [],
//     };

//     // Validate reg numbers and build create list
//     const toCreate: Array<Record<string, unknown>> = [];

//     for (const s of incoming) {
//       if (existingSet.has(s.regNo)) { results.duplicates.push(s.regNo); continue; }

//       const progDoc = programIdMap.get(s.rawProgram) ?? programNameMap.get(s.normalizedProgram);
//       if (!progDoc)  { results.errors.push(`${s.regNo}: Program not found`); continue; }

//       // ── Reg number validation ──────────────────────────────────────────────
//       const validation = await validateRegNo(
//         s.regNo,
//         institutionId.toString(),
//         progDoc._id.toString(),
//       );
//       if (!validation.valid) {
//         results.errors.push(`${s.regNo}: ${validation.reason ?? "Invalid format"}`); continue;
//       }

//       // ── Entry type from year of study ──────────────────────────────────────
//       const entryType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4" =
//         s.yearOfStudy === 2 ? "Mid-Entry-Y2" :
//         s.yearOfStudy === 3 ? "Mid-Entry-Y3" :
//         s.yearOfStudy === 4 ? "Mid-Entry-Y4" : "Direct";

//       // ── Program type — from degreeType, not name inference ─────────────────
//       // This is the FIX: use progDoc.degreeType (e.g. "BSc", "BEd") directly
//       const programType = progDoc.degreeType ?? "BSc";

//       let finalYearId: mongoose.Types.ObjectId;
//       if (s.academicYearId && mongoose.Types.ObjectId.isValid(s.academicYearId)) {
//         finalYearId = new mongoose.Types.ObjectId(s.academicYearId);
//       } else {
//         const resolved = academicYearMap.get(s.admissionAcademicYearString);
//         if (!resolved) { results.errors.push(`${s.regNo}: Academic year not resolved`); continue; }
//         finalYearId = resolved;
//       }

//       toCreate.push({
//         regNo:                s.regNo,
//         name:                 s.name,
//         institution:          institutionId,
//         program:              progDoc._id,
//         programType,          // from Program.degreeType — no name-sniffing
//         entryType,
//         intake:               s.intake,
//         currentYearOfStudy:   s.yearOfStudy,
//         admissionAcademicYear: finalYearId,
//         status:               "active",
//       });
//     }

//     if (toCreate.length > 0) {
//       try {
//         const created = await Student.insertMany(toCreate, { ordered: false });
//         results.registered.push(...created.map(c => String(c.regNo)));
//       } catch (err: unknown) {
//         // insertMany with ordered:false throws on duplicates but still inserts others
//         interface BulkWriteError {
//           insertedDocs?: Array<{ regNo?: unknown }>;
//           writeErrors?:  Array<{ code?: number; op?: { regNo?: unknown } }>;
//         }
//         const bwe = err as BulkWriteError;
//         if (bwe.insertedDocs) {
//           results.registered.push(...bwe.insertedDocs.map(d => String(d.regNo ?? "")));
//         }
//         if (bwe.writeErrors) {
//           const dupes = bwe.writeErrors
//             .filter(e => e.code === 11000)
//             .map(e => String(e.op?.regNo ?? ""));
//           results.duplicates.push(...dupes);
//         }
//       }
//     }

//     await logAudit(req, {
//       action:  "students_bulk_registered",
//       actor:   req.user._id,
//       details: { registered: results.registered.length, duplicates: results.duplicates.length, errors: results.errors.length },
//     });

//     res.status(207).json({
//       message: `${results.registered.length} registered, ${results.duplicates.length} duplicates, ${results.errors.length} errors.`,
//       ...results,
//     });
//   }),
// );

// // ── DELETE /students/:id ──────────────────────────────────────────────────────
// router.delete(
//   "/:id",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const allowedProgramIds = await getScopedProgramIds(req);

//     const student = await Student.findOneAndDelete({
//       _id:         req.params.id,
//       institution: req.user.institution,
//       program:     { $in: allowedProgramIds },
//     });

//     if (!student) {
//       throw { statusCode: 404, message: "Student not found or outside your access scope" } as ApiError;
//     }

//     await logAudit(req, {
//       action:  "delete_student",
//       details: { regNo: student.regNo, name: student.name },
//     });
//     res.json({ message: "Student deleted successfully" });
//   }),
// );

// // ── PATCH /students/:id ───────────────────────────────────────────────────────
// router.patch(
//   "/:id",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { name } = req.body as { name?: string };
//     if (!name?.trim()) {
//       throw { statusCode: 400, message: "Name is required" } as ApiError;
//     }

//     const allowedProgramIds = await getScopedProgramIds(req);

//     const student = await Student.findOneAndUpdate(
//       {
//         _id:         req.params.id,
//         institution: req.user.institution,
//         program:     { $in: allowedProgramIds },
//       },
//       { $set: { name: name.trim() } },
//       { new: true },
//     ).select("name regNo");

//     if (!student) {
//       throw { statusCode: 404, message: "Student not found or outside your access scope" } as ApiError;
//     }

//     await logAudit(req, {
//       action:  "update_student_name",
//       details: { regNo: student.regNo, newName: name },
//     });
//     res.json(student);
//   }),
// );

// // ── DELETE /students/bulk/by-program ─────────────────────────────────────────
// router.delete(
//   "/bulk/by-program",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { programId } = req.body as { programId?: string };
//     if (!programId) {
//       throw { statusCode: 400, message: "Program ID required" } as ApiError;
//     }

//     const result = await Student.deleteMany({
//       program:     programId,
//       institution: req.user.institution,
//     });

//     await logAudit(req, {
//       action:  "bulk_delete_program_students",
//       details: { programId, count: result.deletedCount },
//     });
//     res.json({ message: `Deleted ${result.deletedCount} students from program.` });
//   }),
// );

// export default router;



































































// // serverside/src/routes/students.ts — COMPLETE, ERROR-FREE
// import { Router, Response }  from "express";
// import ExcelJS              from "exceljs";
// import mongoose             from "mongoose";
// import path from "path";
// import Student              from "../models/Student";
// import Program              from "../models/Program";
// import AcademicYear         from "../models/AcademicYear";
// import {
//   requireAuth, requireRole,
//   AuthenticatedRequest, getScopedProgramIds,
// } from "../middleware/auth";
// import { asyncHandler }           from "../middleware/asyncHandler";
// import { logAudit }               from "../lib/auditLogger";
// import { paginate }               from "../utils/paginate";
// import { normalizeProgramName }   from "../services/programNormalizer";
// import { validateRegNo }          from "../utils/validateRegNo";
// import { loadInstitutionSettings } from "../utils/loadInstitutionSettings";
// import { ApiError }               from "../middleware/errorHandler";

// const router = Router();

// // ── Lean types ────────────────────────────────────────────────────────────────
// interface ProgramLean {
//   _id:           mongoose.Types.ObjectId;
//   name:          string;
//   code:          string;
//   durationYears: number;
//   degreeType:    string;
// }

// interface ProgramRef {
//   _id:  mongoose.Types.ObjectId;
//   code: string;
//   name: string;
// }

// interface AcademicYearLean {
//   _id:  mongoose.Types.ObjectId;
//   year: string;
// }

// // ── GET /students ─────────────────────────────────────────────────────────────
// router.get(
//   "/",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const page   = Math.max(1, parseInt(req.query.page  as string) || 1);
//     const limit  = Math.max(1, parseInt(req.query.limit as string) || 20);
//     const search = ((req.query.search as string) ?? "").trim();

//     if (!search) {
//       res.json({ students: [], total: 0, page, totalPages: 0 });
//       return;
//     }

//     const allowedProgramIds = await getScopedProgramIds(req);
//     const filter: Record<string, unknown> = {
//       institution: req.user.institution,
//       program:     { $in: allowedProgramIds },
//       $or: [
//         { regNo: { $regex: search, $options: "i" } },
//         { name:  { $regex: search, $options: "i" } },
//       ],
//     };

//     const [students, total] = await Promise.all([
//       paginate(
//         Student.find(filter)
//           .select("regNo name program currentYearOfStudy status qualifierSuffix intake")
//           .populate("program", "name code departmentCode schoolCode")
//           .lean(),
//         page,
//         limit,
//       ),
//       Student.countDocuments(filter),
//     ]);

//     res.json({ students, total, page, totalPages: Math.ceil(total / Math.min(100, limit)) });
//   }),
// );

// // ── GET /students/stats ───────────────────────────────────────────────────────
// router.get(
//   "/stats",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const allowedProgramIds = await getScopedProgramIds(req);
//     const base = { institution: req.user.institution, program: { $in: allowedProgramIds } };

//     const [total, active, graduated, discontinued] = await Promise.all([
//       Student.countDocuments(base),
//       Student.countDocuments({ ...base, status: "active" }),
//       Student.countDocuments({ ...base, status: "graduated" }),
//       Student.countDocuments({ ...base, status: "discontinued" }),
//     ]);

//     res.json({ total, active, inactive: total - active, graduated, discontinued });
//   }),
// );

// // ── GET /students/template ────────────────────────────────────────────────────
// router.get(
//   "/template",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { programId, academicYearId } = req.query;
//     const institutionId = req.user.institution.toString();

//     const settings = await loadInstitutionSettings(institutionId);
//     const universityName = settings.docMeta.universityName || "University";

//     const allowedProgramIds = await getScopedProgramIds(req);

//     let programs:        ProgramRef[] = [];
//     let selectedProgram: ProgramRef | null = null;

//     if (programId && mongoose.Types.ObjectId.isValid(programId as string)) {
//       const found = await Program.findOne({
//         _id:         new mongoose.Types.ObjectId(programId as string),
//         institution: req.user.institution,
//         _id2:        { $in: allowedProgramIds },
//       })
//         .select("code name")
//         .lean<ProgramRef>();
//       selectedProgram = found ?? null;
//       if (selectedProgram) programs = [selectedProgram];
//     } else {
//       programs = await Program.find({
//         institution: req.user.institution,
//         _id:         { $in: allowedProgramIds },
//       })
//         .select("code name")
//         .lean<ProgramRef[]>();
//     }

//     const yearFilter = academicYearId && mongoose.Types.ObjectId.isValid(academicYearId as string)
//       ? { _id: new mongoose.Types.ObjectId(academicYearId as string) }
//       : { institution: req.user.institution, isCurrent: true };

//     const yearDoc = await AcademicYear.findOne(yearFilter)
//       .select("year")
//       .lean<AcademicYearLean>();
//     const currentYearString = yearDoc?.year ?? "General";

//     const workbook  = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet("Registration");
//     const fontName  = "Book Antiqua";
//     const centerBold = {
//       alignment: { horizontal: "center" as const, vertical: "middle" as const },
//       font: { bold: true, name: fontName },
//     };

//     worksheet.mergeCells("A1:D1");
//     const instCell = worksheet.getCell("A1");
//     instCell.value = universityName.toUpperCase();
//     instCell.style = { ...centerBold, font: { ...centerBold.font, size: 14, underline: true } };

//     worksheet.mergeCells("A2:D2");
//     const progCell = worksheet.getCell("A2");
//     progCell.value = selectedProgram
//       ? `PROGRAM: ${selectedProgram.code} - ${selectedProgram.name.toUpperCase()}`
//       : "PROGRAM: ALL PROGRAMS (Select from dropdown)";
//     progCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     worksheet.mergeCells("A3:D3");
//     const yearHeader = worksheet.getCell("A3");
//     yearHeader.value = `REGISTRATION TEMPLATE - ${currentYearString} ACADEMIC YEAR`;
//     yearHeader.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };

//     worksheet.addRow([]);

//     const headerRow = worksheet.getRow(5);
//     ["Reg No","Full Name","Program","Year of Study","Intake"].forEach((h, i) => {
//       const cell  = headerRow.getCell(i + 1);
//       cell.value  = h;
//       cell.style  = {
//         font:      { bold: true, name: fontName, color: { argb: "FFFFFFFF" } },
//         fill:      { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } },
//         alignment: { horizontal: "center", vertical: "middle" },
//         border:    { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} },
//       };
//     });
//     headerRow.height = 25;

//     const fixedProgramName = selectedProgram?.name ?? "";
//     const border = { top:{style:"thin" as const}, left:{style:"thin" as const}, bottom:{style:"thin" as const}, right:{style:"thin" as const} };

//     for (let r = 6; r <= 506; r++) {
//       const row = worksheet.getRow(r);
//       row.font  = { name: fontName, size: 10 };
//       for (let c = 1; c <= 5; c++) row.getCell(c).border = border;

//       row.getCell(1).protection = { locked: false };
//       row.getCell(2).protection = { locked: false };
//       row.getCell(5).protection = { locked: false };

//       const progDataCell = row.getCell(3);
//       if (selectedProgram) {
//         progDataCell.value      = fixedProgramName;
//         progDataCell.fill       = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F4F6" } };
//         progDataCell.protection = { locked: true };
//       } else {
//         progDataCell.protection = { locked: false };
//         if (programs.length > 0) {
//           const opts = programs.map(p => p.name).join(",").substring(0, 250);
//           progDataCell.dataValidation = { type: "list", allowBlank: false, formulae: [`"${opts}"`] };
//         }
//       }

//       const yearDataCell = row.getCell(4);
//       yearDataCell.protection      = { locked: false };
//       yearDataCell.dataValidation  = { type: "list", allowBlank: false, formulae: ['"1,2,3,4,5,6"'] };

//       const intakeCell = row.getCell(5);
//       intakeCell.protection = { locked: false };
//       intakeCell.value      = "SEPT";
//       intakeCell.dataValidation = {
//         type: "list", allowBlank: false, formulae: ['"JAN,MAY,SEPT"'],
//         showErrorMessage: true, errorTitle: "Invalid Intake",
//         error: "Please select JAN, MAY, or SEPT.",
//       };
//     }

//     worksheet.columns = [
//       { width: 25 }, { width: 40 }, { width: 50 }, { width: 20 }, { width: 15 },
//     ];
//     worksheet.views = [{ state: "frozen", ySplit: 5 }];
//     worksheet.protect("", { selectLockedCells: true, selectUnlockedCells: true });

//     const buffer  = await workbook.xlsx.writeBuffer();
//     const safeYear = currentYearString.replace("/", "-");
//     const cleanName = selectedProgram
//       ? selectedProgram.name.replace(/[^a-zA-Z0-9]/g, "_").replace(/_+/g, "_").toUpperCase()
//       : "ALL";

//     res
//       .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
//       .header("Access-Control-Expose-Headers", "Content-Disposition")
//       .attachment(`Registration_Template_${cleanName}_${safeYear}.xlsx`)
//       .send(Buffer.from(buffer as ArrayBuffer));

//     await logAudit(req, {
//       action: "template_download",
//       details: { type: "student_registration", programId },
//     });
//   }),
// );

// // ── POST /students/bulk ───────────────────────────────────────────────────────
// router.post(
//   "/bulk",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     interface BulkRow {
//       regNo:                       string;
//       name:                        string;
//       program:                     string;
//       currentYearOfStudy?:         number;
//       academicYearId?:             string;
//       intake?:                     string;
//       admissionAcademicYearString?: string;
//     }

//     const { students } = req.body as { students: BulkRow[] };
//     if (!Array.isArray(students) || students.length === 0) {
//       res.status(400).json({ message: "No students provided" }); return;
//     }

//     const institutionId     = req.user.institution;
//     const allowedProgramIds = await getScopedProgramIds(req);

//     const incoming = students.map(s => ({
//       regNo:                       (s.regNo   ?? "").trim().toUpperCase(),
//       name:                        (s.name    ?? "").trim(),
//       rawProgram:                  (s.program ?? "").trim(),
//       normalizedProgram:           normalizeProgramName((s.program ?? "").trim()),
//       yearOfStudy:                 Number(s.currentYearOfStudy) || 1,
//       academicYearId:              s.academicYearId,
//       intake:                      (s.intake ?? "SEPT").trim().toUpperCase(),
//       admissionAcademicYearString: s.admissionAcademicYearString ?? "2024/2025",
//     }));

//     const invalid = incoming.filter(s => !s.regNo || !s.name || !s.rawProgram);
//     if (invalid.length > 0) {
//       res.status(400).json({ message: "Missing Reg No, Name, or Program" }); return;
//     }

//     const dbPrograms = await Program.find({
//       institution: institutionId,
//       _id:         { $in: allowedProgramIds },
//     })
//       .select("name code durationYears degreeType")
//       .lean<ProgramLean[]>();

//     const programNameMap = new Map(dbPrograms.map(p => [normalizeProgramName(p.name), p]));
//     const programIdMap   = new Map(dbPrograms.map(p => [p._id.toString(), p]));

//     const missingPrograms = [...new Set(incoming.map(s => s.rawProgram))].filter(
//       raw => !programIdMap.has(raw) && !programNameMap.has(normalizeProgramName(raw)),
//     );
//     if (missingPrograms.length > 0) {
//       res.status(400).json({ message: "Programs not found or not in your scope", notFound: missingPrograms });
//       return;
//     }

//     // Resolve academic years
//     const academicYearMap = new Map<string, mongoose.Types.ObjectId>();
//     const yearStrings     = [...new Set(
//       incoming.filter(s => !s.academicYearId).map(s => s.admissionAcademicYearString),
//     )];

//     if (yearStrings.length > 0) {
//       const bulkOps = yearStrings.map(yearStr => {
//         const [startYear, endYear] = yearStr.split("/").map(Number);
//         return {
//           updateOne: {
//             filter: { year: yearStr, institution: institutionId },
//             update: { $setOnInsert: {
//               year: yearStr, institution: institutionId,
//               startDate: new Date(`${startYear}-08-01`),
//               endDate:   new Date(`${endYear}-07-31`),
//               isCurrent: false,
//             }},
//             upsert: true,
//           },
//         };
//       });
//       await AcademicYear.bulkWrite(bulkOps);

//       const resolvedYears = await AcademicYear.find({
//         institution: institutionId, year: { $in: yearStrings },
//       })
//         .select("year")
//         .lean<AcademicYearLean[]>();

//       resolvedYears.forEach(y => academicYearMap.set(y.year, y._id));
//     }

//     const regNos     = incoming.map(s => s.regNo);
//     const existing   = await Student.find({
//       regNo: { $in: regNos }, institution: institutionId,
//     })
//       .select("regNo")
//       .lean<Array<{ regNo: string }>>();
//     const existingSet = new Set(existing.map(s => s.regNo));

//     const results = { registered: [] as string[], duplicates: [] as string[], errors: [] as string[] };
//     const toCreate: Array<Record<string, unknown>> = [];

//     for (const s of incoming) {
//       if (existingSet.has(s.regNo)) { results.duplicates.push(s.regNo); continue; }

//       const progDoc = programIdMap.get(s.rawProgram) ?? programNameMap.get(s.normalizedProgram);
//       if (!progDoc) { results.errors.push(`${s.regNo}: Program not found`); continue; }

//       const validation = await validateRegNo(
//         s.regNo, institutionId.toString(), progDoc._id.toString(),
//       );
//       if (!validation.valid) {
//         results.errors.push(`${s.regNo}: ${validation.reason ?? "Invalid format"}`); continue;
//       }

//       const entryType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4" =
//         s.yearOfStudy === 2 ? "Mid-Entry-Y2" :
//         s.yearOfStudy === 3 ? "Mid-Entry-Y3" :
//         s.yearOfStudy === 4 ? "Mid-Entry-Y4" : "Direct";

//       // Read from Program.degreeType — no name inference
//       const programType = progDoc.degreeType ?? "BSc";

//       let finalYearId: mongoose.Types.ObjectId;
//       if (s.academicYearId && mongoose.Types.ObjectId.isValid(s.academicYearId)) {
//         finalYearId = new mongoose.Types.ObjectId(s.academicYearId);
//       } else {
//         const resolved = academicYearMap.get(s.admissionAcademicYearString);
//         if (!resolved) { results.errors.push(`${s.regNo}: Academic year not resolved`); continue; }
//         finalYearId = resolved;
//       }

//       toCreate.push({
//         regNo:                s.regNo,
//         name:                 s.name,
//         institution:          institutionId,
//         program:              progDoc._id,
//         programType,
//         entryType,
//         intake:               s.intake,
//         currentYearOfStudy:   s.yearOfStudy,
//         admissionAcademicYear: finalYearId,
//         status:               "active",
//       });
//     }

//     if (toCreate.length > 0) {
//       try {
//         const created = await Student.insertMany(toCreate, { ordered: false });
//         results.registered.push(...created.map(c => String(c.regNo)));
//       } catch (err: unknown) {
//         interface BulkWriteError {
//           insertedDocs?: Array<{ regNo?: unknown }>;
//           writeErrors?:  Array<{ code?: number; op?: { regNo?: unknown } }>;
//         }
//         const bwe = err as BulkWriteError;
//         if (bwe.insertedDocs) {
//           results.registered.push(...bwe.insertedDocs.map(d => String(d.regNo ?? "")));
//         }
//         if (bwe.writeErrors) {
//           bwe.writeErrors
//             .filter(e => e.code === 11000)
//             .forEach(e => results.duplicates.push(String(e.op?.regNo ?? "")));
//         }
//       }
//     }

//     await logAudit(req, {
//       action:  "students_bulk_registered",
//       details: {
//         registered: results.registered.length,
//         duplicates: results.duplicates.length,
//         errors:     results.errors.length,
//       },
//     });

//     res.status(207).json({
//       message: `${results.registered.length} registered, ${results.duplicates.length} duplicates, ${results.errors.length} errors.`,
//       ...results,
//     });
//   }),
// );

// // ── DELETE /students/:id ──────────────────────────────────────────────────────
// router.delete(
//   "/:id",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const allowedProgramIds = await getScopedProgramIds(req);

//     const student = await Student.findOneAndDelete({
//       _id:         req.params.id,
//       institution: req.user.institution,
//       program:     { $in: allowedProgramIds },
//     });

//     if (!student) {
//       throw { statusCode: 404, message: "Student not found or outside your access scope" } as ApiError;
//     }

//     await logAudit(req, { action: "delete_student", details: { regNo: student.regNo } });
//     res.json({ message: "Student deleted successfully" });
//   }),
// );

// // ── PATCH /students/:id ───────────────────────────────────────────────────────
// router.patch(
//   "/:id",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { name } = req.body as { name?: string };
//     if (!name?.trim()) {
//       throw { statusCode: 400, message: "Name is required" } as ApiError;
//     }

//     const allowedProgramIds = await getScopedProgramIds(req);
//     const student = await Student.findOneAndUpdate(
//       {
//         _id:         req.params.id,
//         institution: req.user.institution,
//         program:     { $in: allowedProgramIds },
//       },
//       { $set: { name: name.trim() } },
//       { new: true },
//     ).select("name regNo");

//     if (!student) {
//       throw { statusCode: 404, message: "Student not found or outside your access scope" } as ApiError;
//     }

//     await logAudit(req, { action: "update_student_name", details: { regNo: student.regNo, newName: name } });
//     res.json(student);
//   }),
// );

// // ── DELETE /students/bulk/by-program ─────────────────────────────────────────
// router.delete(
//   "/bulk/by-program",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { programId } = req.body as { programId?: string };
//     if (!programId) { throw { statusCode: 400, message: "Program ID required" } as ApiError; }

//     const result = await Student.deleteMany({ program: programId, institution: req.user.institution });
//     await logAudit(req, { action: "bulk_delete_program_students", details: { programId, count: result.deletedCount } });
//     res.json({ message: `Deleted ${result.deletedCount} students from program.` });
//   }),
// );

// export default router;

















































// serverside/src/routes/students.ts — GET /students/template — COMPLETE
// This is the full route handler only. Paste it inside the router file
// replacing the existing /template handler.

import { Router, Response }    from "express";
import ExcelJS                 from "exceljs";
import mongoose                from "mongoose";
import path                    from "path";
import Student                 from "../models/Student";
import Program                 from "../models/Program";
import AcademicYear            from "../models/AcademicYear";
import {
  requireAuth, requireRole,
  AuthenticatedRequest, getScopedProgramIds,
} from "../middleware/auth";
import { asyncHandler }           from "../middleware/asyncHandler";
import { logAudit }               from "../lib/auditLogger";
import { paginate }               from "../utils/paginate";
import { normalizeProgramName }   from "../services/programNormalizer";
import { validateRegNo }          from "../utils/validateRegNo";
import { loadInstitutionSettings } from "../utils/loadInstitutionSettings";
import { loadLogoBuffer }          from "../utils/loadLogoBuffer";
import { ApiError }               from "../middleware/errorHandler";

const router = Router();

// ── Lean interfaces ────────────────────────────────────────────────────────────
interface ProgramLean {
  _id:           mongoose.Types.ObjectId;
  name:          string;
  code:          string;
  durationYears: number;
  degreeType:    string;
}

interface ProgramRef {
  _id:  mongoose.Types.ObjectId;
  code: string;
  name: string;
}

interface AcademicYearLean {
  _id:  mongoose.Types.ObjectId;
  year: string;
}

// Reg-pattern shape from InstitutionSettings
interface RegNoPattern {
  prefix:       string;
  separator:    string;
  yearDigits:   number;
  example:      string;
  manualRegex?: string;
}

// ── GET /students ──────────────────────────────────────────────────────────────
router.get(
  "/",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const page   = Math.max(1, parseInt(req.query.page  as string) || 1);
    const limit  = Math.max(1, parseInt(req.query.limit as string) || 20);
    const search = ((req.query.search as string) ?? "").trim();

    if (!search) {
      res.json({ students: [], total: 0, page, totalPages: 0 });
      return;
    }

    const allowedProgramIds = await getScopedProgramIds(req);
    const filter: Record<string, unknown> = {
      institution: req.user.institution,
      program:     { $in: allowedProgramIds },
      $or: [
        { regNo: { $regex: search, $options: "i" } },
        { name:  { $regex: search, $options: "i" } },
      ],
    };

    const [students, total] = await Promise.all([
      paginate(
        Student.find(filter)
          .select("regNo name program currentYearOfStudy status qualifierSuffix intake")
          .populate("program", "name code departmentCode schoolCode")
          .lean(),
        page, limit,
      ),
      Student.countDocuments(filter),
    ]);

    res.json({ students, total, page, totalPages: Math.ceil(total / Math.min(100, limit)) });
  }),
);

// ── GET /students/stats ────────────────────────────────────────────────────────
router.get(
  "/stats",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const allowedProgramIds = await getScopedProgramIds(req);
    const base = { institution: req.user.institution, program: { $in: allowedProgramIds } };

    const [total, active, graduated, discontinued] = await Promise.all([
      Student.countDocuments(base),
      Student.countDocuments({ ...base, status: "active" }),
      Student.countDocuments({ ...base, status: "graduated" }),
      Student.countDocuments({ ...base, status: "discontinued" }),
    ]);

    res.json({ total, active, inactive: total - active, graduated, discontinued });
  }),
);

// ── GET /students/template ─────────────────────────────────────────────────────
// Generates an Excel registration template scoped to the coordinator's department.
// If programId is supplied, locks the program column to that program.
// If no programId, shows a dropdown of all programs in coordinator's scope.
// University name and logo come from InstitutionSettings (DB), not env vars.
// Reg-no patterns from InstitutionSettings.schools[schoolCode].departments[deptCode]
// are shown in the header and used for data-validation in the Reg No column.
router.get(
  "/template",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { programId, academicYearId } = req.query;
    const institutionId = req.user.institution.toString();

    // ── 1. Load institution settings + logo from DB ────────────────────────────
    const [settings, logoBuffer] = await Promise.all([
      loadInstitutionSettings(institutionId),
      loadLogoBuffer(institutionId),
    ]);

    const universityName = settings.docMeta.universityName || "University";
    const schoolName     = settings.docMeta.schoolName     || "";

    // ── 2. Resolve coordinator's department reg-no patterns ────────────────────
    // Used to: (a) show format hint in header, (b) add cell validation
    const mySchoolCode = req.user.schoolCode ?? null;
    const myDeptCode   = req.user.departmentCode ?? null;

    let deptPatterns: RegNoPattern[] = [];
    let deptName = "";

    if (mySchoolCode && myDeptCode && settings) {
      // settings.schools is ISchool[] from loadInstitutionSettings
      const schoolsRaw = (settings as unknown as {
        schools?: Array<{
          code: string; name: string;
          departments?: Array<{ code: string; name: string; regNoPatterns?: RegNoPattern[] }>;
        }>;
      }).schools ?? [];

      const school = schoolsRaw.find(s => s.code === mySchoolCode.toUpperCase());
      const dept   = school?.departments?.find(d => d.code === myDeptCode.toUpperCase());
      deptPatterns = dept?.regNoPatterns ?? [];
      deptName     = dept?.name ?? "";
    }

    const enforcePatterns = (settings as unknown as { enforceRegNoPattern?: boolean }).enforceRegNoPattern ?? false;
    const patternExample  = deptPatterns.length > 0
      ? deptPatterns.map(p => p.example).join("  or  ")
      : "";

    // ── 3. Scope programs to coordinator's department ──────────────────────────
    const allowedProgramIds = await getScopedProgramIds(req);
    let programs:        ProgramRef[] = [];
    let selectedProgram: ProgramRef | null = null;

    if (programId && mongoose.Types.ObjectId.isValid(programId as string)) {
      // FIX: was using fake `_id2` field — correct query uses only real fields
      const pid = new mongoose.Types.ObjectId(programId as string);
      // Check it's in the allowed set
      if (allowedProgramIds.map(String).includes(pid.toString())) {
        selectedProgram = await Program.findOne({
          _id:         pid,
          institution: req.user.institution,
        })
          .select("code name")
          .lean<ProgramRef>() ?? null;
      }
      if (selectedProgram) programs = [selectedProgram];
    }

    if (!selectedProgram) {
      // Load all scoped programs for the dropdown
      programs = await Program.find({
        institution: req.user.institution,
        _id:         { $in: allowedProgramIds },
        isActive:    true,
      })
        .select("code name")
        .sort({ name: 1 })
        .lean<ProgramRef[]>();
    }

    // ── 4. Resolve academic year ──────────────────────────────────────────────
    const yearFilter =
      academicYearId && mongoose.Types.ObjectId.isValid(academicYearId as string)
        ? { _id: new mongoose.Types.ObjectId(academicYearId as string) }
        : { institution: req.user.institution, isCurrent: true };

    const yearDoc = await AcademicYear.findOne(yearFilter)
      .select("year")
      .lean<AcademicYearLean>();
    const currentYearString = yearDoc?.year ?? "General";

    // ── 5. Build workbook ─────────────────────────────────────────────────────
    const workbook  = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Registration");
    const fontName  = "Book Antiqua";

    const centerBold = {
      alignment: { horizontal: "center" as const, vertical: "middle" as const },
      font:      { bold: true, name: fontName },
    };

    const thinBorder = {
      top:    { style: "thin" as const },
      left:   { style: "thin" as const },
      bottom: { style: "thin" as const },
      right:  { style: "thin" as const },
    };

    // ── Row 1: Logo (left) + University name (center) ─────────────────────────
    let currentRow = 1;

    if (logoBuffer && logoBuffer.length > 0) {
      const logoId = workbook.addImage({
        buffer:    logoBuffer as any,
        extension: "png",
      });
      worksheet.addImage(logoId, {
        tl:  { col: 0, row: 0 },   // top-left: col A, row 1
        ext: { width: 80, height: 80 },
      });
    }

    // University name — columns B to E merged, row 1
    worksheet.mergeCells("B1:E1");
    const uniCell   = worksheet.getCell("B1");
    uniCell.value   = universityName.toUpperCase();
    uniCell.style   = {
      ...centerBold,
      font: { ...centerBold.font, size: 14, underline: true },
    };
    worksheet.getRow(1).height = 45;

    // ── Row 2: School name ─────────────────────────────────────────────────────
    if (schoolName) {
      worksheet.mergeCells("B2:E2");
      const schCell  = worksheet.getCell("B2");
      schCell.value  = schoolName.toUpperCase();
      schCell.style  = { ...centerBold, font: { ...centerBold.font, size: 11 } };
      worksheet.getRow(2).height = 20;
      currentRow = 2;
    }

    // ── Row 3: Department + Program ───────────────────────────────────────────
    currentRow++;
    const progRowNum = currentRow;
    worksheet.mergeCells(`A${progRowNum}:E${progRowNum}`);
    const progCell   = worksheet.getCell(`A${progRowNum}`);
    progCell.value   = selectedProgram
      ? `${selectedProgram.code} — ${selectedProgram.name.toUpperCase()}`
      : (deptName ? `${deptName.toUpperCase()} — ALL PROGRAMS (select below)` : "ALL PROGRAMS — Select from dropdown");
    progCell.style   = { ...centerBold, font: { ...centerBold.font, size: 11 } };
    worksheet.getRow(progRowNum).height = 20;

    // ── Row 4: Academic year ───────────────────────────────────────────────────
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    const yrCell  = worksheet.getCell(`A${currentRow}`);
    yrCell.value  = `REGISTRATION TEMPLATE — ${currentYearString} ACADEMIC YEAR`;
    yrCell.style  = { ...centerBold, font: { ...centerBold.font, size: 11 } };
    worksheet.getRow(currentRow).height = 18;

    // ── Row 5: Reg-no format notice (only if patterns configured) ──────────────
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    const fmtCell  = worksheet.getCell(`A${currentRow}`);
    if (patternExample) {
      fmtCell.value = `Reg No format for ${myDeptCode}: ${patternExample}${enforcePatterns ? " (enforced)" : " (reference)"}`;
      fmtCell.style = {
        alignment: { horizontal: "center" as const, vertical: "middle" as const },
        font:      { name: fontName, size: 9, color: { argb: "FF1E40AF" }, bold: true },
        fill:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFDBEAFE" } },
      };
    } else {
      fmtCell.value = "Fill in all columns. Reg No, Full Name, and Program are required.";
      fmtCell.style = {
        alignment: { horizontal: "center" as const, vertical: "middle" as const },
        font:      { name: fontName, size: 9, color: { argb: "FF6B7280" } },
        fill:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF9FAFB" } },
      };
    }
    worksheet.getRow(currentRow).height = 16;

    // ── Empty spacer row ───────────────────────────────────────────────────────
    currentRow++;
    worksheet.getRow(currentRow).height = 6;

    // ── Header row ────────────────────────────────────────────────────────────
    const headerRowNum = currentRow + 1;
    currentRow = headerRowNum;
    const headerRow = worksheet.getRow(headerRowNum);
    ["Reg No", "Full Name", "Program", "Year of Study", "Intake"].forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.style = {
        font:      { bold: true, name: fontName, size: 10, color: { argb: "FFFFFFFF" } },
        fill:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF1E3A5F" } },
        alignment: { horizontal: "center" as const, vertical: "middle" as const },
        border:    thinBorder,
      };
    });
    headerRow.height = 22;

    // ── Data rows ─────────────────────────────────────────────────────────────
    const dataStartRow     = headerRowNum + 1;
    const dataEndRow       = dataStartRow + 500;
    const fixedProgramName = selectedProgram?.name ?? "";

    // Build reg-no regex patterns for data validation
    // Excel doesn't support JS regex natively, so we use "Custom" validation
    // with a helper formula only if manualRegex is not supplied.
    // For simple prefix patterns we use a text-contains approach.
    const firstPattern = deptPatterns[0] ?? null;

    for (let r = dataStartRow; r <= dataEndRow; r++) {
      const row = worksheet.getRow(r);
      row.font  = { name: fontName, size: 10 };

      for (let c = 1; c <= 5; c++) {
        row.getCell(c).border = thinBorder;
      }

      // Col A — Reg No
      const regCell = row.getCell(1);
      regCell.protection = { locked: false };
      // Add data validation hint if pattern configured
      if (enforcePatterns && firstPattern) {
        regCell.dataValidation = {
          type:             "textLength",
          operator:         "greaterThan",
          allowBlank:       true,
          formulae:         [2],                    // must be > 2 chars
          showErrorMessage: true,
          errorTitle:       "Invalid Reg Number",
          error:            `Format: ${patternExample}`,
        };
      }

      // Col B — Full Name
      row.getCell(2).protection = { locked: false };

      // Col C — Program
      const progDataCell = row.getCell(3);
      if (selectedProgram) {
        // Locked to selected program
        progDataCell.value      = fixedProgramName;
        progDataCell.fill       = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF0FDF4" } };
        progDataCell.font       = { name: fontName, size: 10, color: { argb: "FF166534" } };
        progDataCell.protection = { locked: true };
      } else {
        progDataCell.protection = { locked: false };
        if (programs.length > 0) {
          // Build dropdown list — max 255 chars for Excel
          const optStr = programs.map(p => p.name).join(",");
          const safeOpts = optStr.length <= 250
            ? `"${optStr}"`
            : `"${programs.slice(0, Math.floor(programs.length / 2)).map(p => p.name).join(",")}"`;
          progDataCell.dataValidation = {
            type:        "list",
            allowBlank:  false,
            formulae:    [safeOpts],
            showErrorMessage: true,
            errorTitle:  "Invalid Program",
            error:       "Please select a program from the dropdown list.",
          };
        }
      }

      // Col D — Year of Study
      const yearDataCell = row.getCell(4);
      yearDataCell.protection = { locked: false };
      yearDataCell.dataValidation = {
        type:        "list",
        allowBlank:  false,
        formulae:    ['"1,2,3,4,5,6"'],
        showErrorMessage: true,
        errorTitle:  "Invalid Year",
        error:       "Year of study must be 1–6.",
      };

      // Col E — Intake
      const intakeCell = row.getCell(5);
      intakeCell.protection = { locked: false };
      intakeCell.value      = "SEPT";    // sensible default
      intakeCell.dataValidation = {
        type:        "list",
        allowBlank:  false,
        formulae:    ['"JAN,MAY,SEPT"'],
        showErrorMessage: true,
        errorTitle:  "Invalid Intake",
        error:       "Intake must be JAN, MAY, or SEPT.",
      };
    }

    // ── Alternate row shading for readability ─────────────────────────────────
    for (let r = dataStartRow; r <= dataEndRow; r += 2) {
      const row = worksheet.getRow(r);
      for (let c = 1; c <= 5; c++) {
        const cell = row.getCell(c);
        if (!cell.fill || (cell.fill as ExcelJS.FillPattern).type !== "pattern") {
          cell.fill = {
            type: "pattern" as const, pattern: "solid" as const,
            fgColor: { argb: "FFF9FAFB" },
          };
        }
      }
    }

    // ── Column widths ─────────────────────────────────────────────────────────
    worksheet.getColumn("A").width = 28;   // Reg No — wide enough for E024-01-0001/2024
    worksheet.getColumn("B").width = 40;   // Full Name
    worksheet.getColumn("C").width = 52;   // Program name — longest field
    worksheet.getColumn("D").width = 18;   // Year of Study
    worksheet.getColumn("E").width = 12;   // Intake

    // ── Freeze header rows ────────────────────────────────────────────────────
    worksheet.views = [{ state: "frozen", ySplit: headerRowNum }];

    // ── Sheet protection — lock structure but allow data entry ────────────────
    worksheet.protect("", {
      selectLockedCells:   true,
      selectUnlockedCells: true,
      formatCells:         false,
      formatColumns:       false,
      formatRows:          false,
    });

    // ── Generate buffer and send ──────────────────────────────────────────────
    const buffer   = await workbook.xlsx.writeBuffer();
    const safeYear = currentYearString.replace(/\//g, "-");
    const cleanName = selectedProgram
      ? selectedProgram.name
          .replace(/[^a-zA-Z0-9\s]/g, "")
          .replace(/\s+/g, "_")
          .toUpperCase()
          .slice(0, 40)
      : (myDeptCode ?? "ALL");
    const filename = `Registration_Template_${cleanName}_${safeYear}.xlsx`;

    res
      .header("Content-Type",        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="${filename}"`)
      .header("Access-Control-Expose-Headers", "Content-Disposition")
      .send(Buffer.from(buffer as ArrayBuffer));

    await logAudit(req, {
      action:  "template_download",
      details: { type: "student_registration", programId, universityName, deptCode: myDeptCode },
    });
  }),
);

// ── POST /students/bulk ────────────────────────────────────────────────────────
router.post(
  "/bulk",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    interface BulkRow {
      regNo:                       string;
      name:                        string;
      program:                     string;
      currentYearOfStudy?:         number;
      academicYearId?:             string;
      intake?:                     string;
      admissionAcademicYearString?: string;
    }

    const { students } = req.body as { students: BulkRow[] };
    if (!Array.isArray(students) || students.length === 0) {
      res.status(400).json({ message: "No students provided" }); return;
    }

    const institutionId     = req.user.institution;
    const allowedProgramIds = await getScopedProgramIds(req);

    const incoming = students.map(s => ({
      regNo:                       (s.regNo   ?? "").trim().toUpperCase(),
      name:                        (s.name    ?? "").trim(),
      rawProgram:                  (s.program ?? "").trim(),
      normalizedProgram:           normalizeProgramName((s.program ?? "").trim()),
      yearOfStudy:                 Number(s.currentYearOfStudy) || 1,
      academicYearId:              s.academicYearId,
      intake:                      (s.intake ?? "SEPT").trim().toUpperCase(),
      admissionAcademicYearString: s.admissionAcademicYearString ?? "2024/2025",
    }));

    const invalid = incoming.filter(s => !s.regNo || !s.name || !s.rawProgram);
    if (invalid.length > 0) {
      res.status(400).json({ message: "Missing Reg No, Name, or Program" }); return;
    }

    const dbPrograms = await Program.find({
      institution: institutionId,
      _id:         { $in: allowedProgramIds },
    })
      .select("name code durationYears degreeType")
      .lean<ProgramLean[]>();

    const programNameMap = new Map(dbPrograms.map(p => [normalizeProgramName(p.name), p]));
    const programIdMap   = new Map(dbPrograms.map(p => [p._id.toString(), p]));

    const missingPrograms = [...new Set(incoming.map(s => s.rawProgram))].filter(
      raw => !programIdMap.has(raw) && !programNameMap.has(normalizeProgramName(raw)),
    );
    if (missingPrograms.length > 0) {
      res.status(400).json({
        message:  "Programs not found or not in your scope",
        notFound: missingPrograms,
      });
      return;
    }

    // Resolve academic years
    const academicYearMap = new Map<string, mongoose.Types.ObjectId>();
    const yearStrings     = [...new Set(
      incoming.filter(s => !s.academicYearId).map(s => s.admissionAcademicYearString),
    )];

    if (yearStrings.length > 0) {
      const bulkOps = yearStrings.map(yearStr => {
        const [startYear, endYear] = yearStr.split("/").map(Number);
        return {
          updateOne: {
            filter: { year: yearStr, institution: institutionId },
            update: { $setOnInsert: {
              year: yearStr, institution: institutionId,
              startDate: new Date(`${startYear}-08-01`),
              endDate:   new Date(`${endYear}-07-31`),
              isCurrent: false,
            }},
            upsert: true,
          },
        };
      });
      await AcademicYear.bulkWrite(bulkOps);

      const resolvedYears = await AcademicYear.find({
        institution: institutionId, year: { $in: yearStrings },
      })
        .select("year")
        .lean<AcademicYearLean[]>();

      resolvedYears.forEach(y => academicYearMap.set(y.year, y._id));
    }

    const regNos    = incoming.map(s => s.regNo);
    const existing  = await Student.find({ regNo: { $in: regNos }, institution: institutionId })
      .select("regNo")
      .lean<Array<{ regNo: string }>>();
    const existingSet = new Set(existing.map(s => s.regNo));

    const results = {
      registered: [] as string[],
      duplicates: [] as string[],
      errors:     [] as string[],
    };
    const toCreate: Array<Record<string, unknown>> = [];

    for (const s of incoming) {
      if (existingSet.has(s.regNo)) { results.duplicates.push(s.regNo); continue; }

      const progDoc = programIdMap.get(s.rawProgram) ?? programNameMap.get(s.normalizedProgram);
      if (!progDoc) { results.errors.push(`${s.regNo}: Program not found`); continue; }

      // Reg-no validation (server-side, authoritative)
      const validation = await validateRegNo(
        s.regNo, institutionId.toString(), progDoc._id.toString(),
      );
      if (!validation.valid) {
        results.errors.push(`${s.regNo}: ${validation.reason ?? "Invalid reg no format"}`);
        continue;
      }

      const entryType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4" =
        s.yearOfStudy === 2 ? "Mid-Entry-Y2" :
        s.yearOfStudy === 3 ? "Mid-Entry-Y3" :
        s.yearOfStudy === 4 ? "Mid-Entry-Y4" : "Direct";

      const programType = progDoc.degreeType ?? "BSc";

      let finalYearId: mongoose.Types.ObjectId;
      if (s.academicYearId && mongoose.Types.ObjectId.isValid(s.academicYearId)) {
        finalYearId = new mongoose.Types.ObjectId(s.academicYearId);
      } else {
        const resolved = academicYearMap.get(s.admissionAcademicYearString);
        if (!resolved) {
          results.errors.push(`${s.regNo}: Academic year not resolved`); continue;
        }
        finalYearId = resolved;
      }

      toCreate.push({
        regNo: s.regNo, name: s.name, institution: institutionId,
        program: progDoc._id, programType, entryType,
        intake: s.intake, currentYearOfStudy: s.yearOfStudy,
        admissionAcademicYear: finalYearId, status: "active",
      });
    }

    if (toCreate.length > 0) {
      try {
        const created = await Student.insertMany(toCreate, { ordered: false });
        results.registered.push(...created.map(c => String(c.regNo)));
      } catch (err: unknown) {
        interface BulkWriteError {
          insertedDocs?: Array<{ regNo?: unknown }>;
          writeErrors?:  Array<{ code?: number; op?: { regNo?: unknown } }>;
        }
        const bwe = err as BulkWriteError;
        if (bwe.insertedDocs) {
          results.registered.push(...bwe.insertedDocs.map(d => String(d.regNo ?? "")));
        }
        if (bwe.writeErrors) {
          bwe.writeErrors
            .filter(e => e.code === 11000)
            .forEach(e => results.duplicates.push(String(e.op?.regNo ?? "")));
        }
      }
    }

    await logAudit(req, {
      action:  "students_bulk_registered",
      details: {
        registered: results.registered.length,
        duplicates: results.duplicates.length,
        errors:     results.errors.length,
      },
    });

    res.status(207).json({
      message: `${results.registered.length} registered, ${results.duplicates.length} duplicates, ${results.errors.length} errors.`,
      ...results,
    });
  }),
);

// ── DELETE /students/:id ───────────────────────────────────────────────────────
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const allowedProgramIds = await getScopedProgramIds(req);
    const student = await Student.findOneAndDelete({
      _id:         req.params.id,
      institution: req.user.institution,
      program:     { $in: allowedProgramIds },
    });
    if (!student) {
      throw { statusCode: 404, message: "Student not found or outside your scope" } as ApiError;
    }
    await logAudit(req, { action: "delete_student", details: { regNo: student.regNo } });
    res.json({ message: "Student deleted successfully" });
  }),
);

// ── PATCH /students/:id ────────────────────────────────────────────────────────
router.patch(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { name } = req.body as { name?: string };
    if (!name?.trim()) {
      throw { statusCode: 400, message: "Name is required" } as ApiError;
    }
    const allowedProgramIds = await getScopedProgramIds(req);
    const student = await Student.findOneAndUpdate(
      { _id: req.params.id, institution: req.user.institution, program: { $in: allowedProgramIds } },
      { $set: { name: name.trim() } },
      { new: true },
    ).select("name regNo");
    if (!student) {
      throw { statusCode: 404, message: "Student not found or outside your scope" } as ApiError;
    }
    await logAudit(req, { action: "update_student_name", details: { regNo: student.regNo } });
    res.json(student);
  }),
);

// ── DELETE /students/bulk/by-program ──────────────────────────────────────────
router.delete(
  "/bulk/by-program",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { programId } = req.body as { programId?: string };
    if (!programId) throw { statusCode: 400, message: "Program ID required" } as ApiError;
    const result = await Student.deleteMany({
      program:     programId,
      institution: req.user.institution,
    });
    await logAudit(req, { action: "bulk_delete_program_students", details: { programId, count: result.deletedCount } });
    res.json({ message: `Deleted ${result.deletedCount} students from program.` });
  }),
);

export default router;
