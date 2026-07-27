"use strict";
// serverside/src/routes/students.ts — GET /students/template — COMPLETE
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = require("express");
const exceljs_1 = __importDefault(require("exceljs"));
const mongoose_1 = __importDefault(require("mongoose"));
const Student_1 = __importDefault(require("../models/Student"));
const Program_1 = __importDefault(require("../models/Program"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const auditLogger_1 = require("../lib/auditLogger");
const paginate_1 = require("../utils/paginate");
const programNormalizer_1 = require("../services/programNormalizer");
const validateRegNo_1 = require("../utils/validateRegNo");
const loadInstitutionSettings_1 = require("../utils/loadInstitutionSettings");
const loadLogoBuffer_1 = require("../utils/loadLogoBuffer");
const multer_1 = __importDefault(require("multer"));
const xlsxLib = __importStar(require("xlsx"));
const router = (0, express_1.Router)();
const uploadStudentExcelMiddleware = (0, multer_1.default)({
    storage: multer_1.default.memoryStorage(),
    limits: { fileSize: 10 * 1024 * 1024 },
    fileFilter: (_req, file, cb) => {
        cb(null, /\.(xlsx|xls|csv)$/i.test(file.originalname));
    },
});
// ── GET /students ──────────────────────────────────────────────────────────────
router.get("/", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.max(1, parseInt(req.query.limit) || 20);
    const search = (req.query.search ?? "").trim();
    if (!search) {
        res.json({ students: [], total: 0, page, totalPages: 0 });
        return;
    }
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    const filter = {
        institution: req.user.institution,
        program: { $in: allowedProgramIds },
        $or: [
            { regNo: { $regex: search, $options: "i" } },
            { name: { $regex: search, $options: "i" } },
        ],
    };
    const [students, total] = await Promise.all([
        (0, paginate_1.paginate)(Student_1.default.find(filter)
            .select("regNo name program currentYearOfStudy status qualifierSuffix intake")
            .populate("program", "name code departmentCode schoolCode")
            .lean(), page, limit),
        Student_1.default.countDocuments(filter),
    ]);
    res.json({ students, total, page, totalPages: Math.ceil(total / Math.min(100, limit)) });
}));
// ── GET /students/stats ────────────────────────────────────────────────────────
router.get("/stats", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    const base = { institution: req.user.institution, program: { $in: allowedProgramIds } };
    const [total, active, graduated, discontinued] = await Promise.all([
        Student_1.default.countDocuments(base),
        Student_1.default.countDocuments({ ...base, status: "active" }),
        Student_1.default.countDocuments({ ...base, status: "graduated" }),
        Student_1.default.countDocuments({ ...base, status: "discontinued" }),
    ]);
    res.json({ total, active, inactive: total - active, graduated, discontinued });
}));
// ── GET /students/template ─────────────────────────────────────────────────────
// Generates an Excel registration template scoped to the coordinator's department.
// If programId is supplied, locks the program column to that program.
// If no programId, shows a dropdown of all programs in coordinator's scope.
// University name and logo come from InstitutionSettings (DB), not env vars.
// Reg-no patterns from InstitutionSettings.schools[schoolCode].departments[deptCode]
// are shown in the header and used for data-validation in the Reg No column.
router.get("/template", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, academicYearId } = req.query;
    const institutionId = req.user.institution.toString();
    // ── 1. Load institution settings + logo from DB ────────────────────────────
    const [settings, logoBuffer] = await Promise.all([
        (0, loadInstitutionSettings_1.loadInstitutionSettings)(institutionId), (0, loadLogoBuffer_1.loadLogoBuffer)(institutionId)
    ]);
    const universityName = settings.docMeta.universityName || "University";
    const schoolName = settings.docMeta.schoolName || "";
    // ── 2. Resolve coordinator's department reg-no patterns ────────────────────
    // Used to: (a) show format hint in header, (b) add cell validation
    const mySchoolCode = req.user.schoolCode ?? null;
    const myDeptCode = req.user.departmentCode ?? null;
    let deptPatterns = [];
    let deptName = "";
    if (mySchoolCode && myDeptCode && settings) {
        // settings.schools is ISchool[] from loadInstitutionSettings
        const schoolsRaw = settings.schools ?? [];
        const school = schoolsRaw.find(s => s.code === mySchoolCode.toUpperCase());
        const dept = school?.departments?.find(d => d.code === myDeptCode.toUpperCase());
        deptPatterns = dept?.regNoPatterns ?? [];
        deptName = dept?.name ?? "";
    }
    const enforcePatterns = settings.enforceRegNoPattern ?? false;
    const patternExample = deptPatterns.length > 0
        ? deptPatterns.map(p => p.example).join("  or  ")
        : "";
    // ── 3. Scope programs to coordinator's department ──────────────────────────
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    let programs = [];
    let selectedProgram = null;
    if (programId && mongoose_1.default.Types.ObjectId.isValid(programId)) {
        const pid = new mongoose_1.default.Types.ObjectId(programId);
        if (allowedProgramIds.map(String).includes(pid.toString())) {
            selectedProgram = await Program_1.default.findOne({
                _id: pid,
                institution: req.user.institution,
            })
                .select("code name")
                .lean() ?? null;
        }
        if (selectedProgram)
            programs = [selectedProgram];
    }
    if (!selectedProgram) {
        // Load all scoped programs for the dropdown
        programs = await Program_1.default.find({
            institution: req.user.institution,
            _id: { $in: allowedProgramIds },
            isActive: true,
        })
            .select("code name")
            .sort({ name: 1 })
            .lean();
    }
    // ── 4. Resolve academic year ──────────────────────────────────────────────
    const yearFilter = academicYearId && mongoose_1.default.Types.ObjectId.isValid(academicYearId)
        ? { _id: new mongoose_1.default.Types.ObjectId(academicYearId) }
        : { institution: req.user.institution, isCurrent: true };
    const yearDoc = await AcademicYear_1.default.findOne(yearFilter)
        .select("year")
        .lean();
    const currentYearString = yearDoc?.year ?? "General";
    // ── 5. Build workbook ─────────────────────────────────────────────────────
    const workbook = new exceljs_1.default.Workbook();
    const worksheet = workbook.addWorksheet("Registration Template");
    const fontName = "Book Antiqua";
    const centerBold = {
        alignment: { horizontal: "center", vertical: "middle" },
        font: { bold: true, name: fontName },
    };
    const thinBorder = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
    };
    // ── Row 1: Logo (left) + University name (center) ─────────────────────────
    let currentRow = 1;
    // if (logoBuffer && logoBuffer.length > 0) {
    //   const logoId = workbook.addImage({ buffer: logoBuffer as any, extension: "png" });
    //   worksheet.addImage(logoId, { tl:  { col: 0, row: 0 }, ext: { width: 80, height: 80 }});
    // }
    // University name — columns B to E merged, row 1
    worksheet.mergeCells("B1:E1");
    const uniCell = worksheet.getCell("B1");
    uniCell.value = universityName.toUpperCase();
    uniCell.style = { ...centerBold, font: { ...centerBold.font, size: 14, underline: true } };
    worksheet.getRow(1).height = 45;
    // ── Row 2: School name ─────────────────────────────────────────────────────
    if (schoolName) {
        worksheet.mergeCells("B2:E2");
        const schCell = worksheet.getCell("B2");
        schCell.value = schoolName.toUpperCase();
        schCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };
        worksheet.getRow(2).height = 20;
        currentRow = 2;
    }
    // ── Row 3: Department + Program ───────────────────────────────────────────
    currentRow++;
    const progRowNum = currentRow;
    worksheet.mergeCells(`A${progRowNum}:E${progRowNum}`);
    const progCell = worksheet.getCell(`A${progRowNum}`);
    progCell.value = selectedProgram
        ? `${selectedProgram.code} — ${selectedProgram.name.toUpperCase()}`
        : (deptName ? `${deptName.toUpperCase()} — ALL PROGRAMS (select below)` : "ALL PROGRAMS — Select from dropdown");
    progCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };
    worksheet.getRow(progRowNum).height = 20;
    // ── Row 4: Academic year ───────────────────────────────────────────────────
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    const yrCell = worksheet.getCell(`A${currentRow}`);
    yrCell.value = `REGISTRATION TEMPLATE — ${currentYearString} ACADEMIC YEAR`;
    yrCell.style = { ...centerBold, font: { ...centerBold.font, size: 11 } };
    worksheet.getRow(currentRow).height = 18;
    // ── Row 5: Reg-no format notice (only if patterns configured) ──────────────
    currentRow++;
    worksheet.mergeCells(`A${currentRow}:E${currentRow}`);
    const fmtCell = worksheet.getCell(`A${currentRow}`);
    if (patternExample) {
        fmtCell.value = `Reg No format for ${myDeptCode}: ${patternExample}${enforcePatterns ? " (enforced)" : " (reference)"}`;
        fmtCell.style = {
            alignment: { horizontal: "center", vertical: "middle" },
            font: { name: fontName, size: 9, color: { argb: "FF1E40AF" }, bold: true },
            fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFDBEAFE" } },
        };
    }
    else {
        fmtCell.value = "Fill in all columns. Reg No, Full Name, and Program are required.";
        fmtCell.style = {
            alignment: { horizontal: "center", vertical: "middle" },
            font: { name: fontName, size: 9, color: { argb: "FF6B7280" } },
            fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9FAFB" } },
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
            font: { bold: true, name: fontName, size: 10, color: { argb: "FFFFFFFF" } },
            fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E3A5F" } },
            alignment: { horizontal: "center", vertical: "middle" },
            border: thinBorder,
        };
    });
    headerRow.height = 22;
    // ── Data rows ─────────────────────────────────────────────────────────────
    const dataStartRow = headerRowNum + 1;
    const dataEndRow = dataStartRow + 500;
    const fixedProgramName = selectedProgram?.name ?? "";
    // Build reg-no regex patterns for data validation
    // Excel doesn't support JS regex natively, so we use "Custom" validation
    // with a helper formula only if manualRegex is not supplied.
    // For simple prefix patterns we use a text-contains approach.
    const firstPattern = deptPatterns[0] ?? null;
    for (let r = dataStartRow; r <= dataEndRow; r++) {
        const row = worksheet.getRow(r);
        row.font = { name: fontName, size: 10 };
        for (let c = 1; c <= 5; c++) {
            row.getCell(c).border = thinBorder;
        }
        // Col A — Reg No
        const regCell = row.getCell(1);
        regCell.protection = { locked: false };
        // Add data validation hint if pattern configured
        if (enforcePatterns && firstPattern) {
            regCell.dataValidation = {
                type: "textLength",
                operator: "greaterThan",
                allowBlank: true,
                formulae: [2], // must be > 2 chars
                showErrorMessage: true,
                errorTitle: "Invalid Reg Number",
                error: `Format: ${patternExample}`,
            };
        }
        // Col B — Full Name
        row.getCell(2).protection = { locked: false };
        // Col C — Program
        const progDataCell = row.getCell(3);
        if (selectedProgram) {
            // Locked to selected program
            progDataCell.value = fixedProgramName;
            progDataCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0FDF4" } };
            progDataCell.font = { name: fontName, size: 10, color: { argb: "FF166534" } };
            progDataCell.protection = { locked: true };
        }
        else {
            progDataCell.protection = { locked: false };
            if (programs.length > 0) {
                // Build dropdown list — max 255 chars for Excel
                const optStr = programs.map(p => p.name).join(",");
                const safeOpts = optStr.length <= 250
                    ? `"${optStr}"`
                    : `"${programs.slice(0, Math.floor(programs.length / 2)).map(p => p.name).join(",")}"`;
                progDataCell.dataValidation = {
                    type: "list",
                    allowBlank: false,
                    formulae: [safeOpts],
                    showErrorMessage: true,
                    errorTitle: "Invalid Program",
                    error: "Please select a program from the dropdown list.",
                };
            }
        }
        // Col D — Year of Study
        const yearDataCell = row.getCell(4);
        yearDataCell.protection = { locked: false };
        yearDataCell.dataValidation = {
            type: "list",
            allowBlank: false,
            formulae: ['"1,2,3,4,5,6"'],
            showErrorMessage: true,
            errorTitle: "Invalid Year",
            error: "Year of study must be 1–6.",
        };
        // Col E — Intake
        const intakeCell = row.getCell(5);
        intakeCell.protection = { locked: false };
        intakeCell.value = "SEPT"; // sensible default
        intakeCell.dataValidation = {
            type: "list",
            allowBlank: false,
            formulae: ['"JAN,MAY,SEPT"'],
            showErrorMessage: true,
            errorTitle: "Invalid Intake",
            error: "Intake must be JAN, MAY, or SEPT.",
        };
    }
    // ── Alternate row shading for readability ─────────────────────────────────
    for (let r = dataStartRow; r <= dataEndRow; r += 2) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= 5; c++) {
            const cell = row.getCell(c);
            if (!cell.fill || cell.fill.type !== "pattern") {
                cell.fill = {
                    type: "pattern", pattern: "solid",
                    fgColor: { argb: "FFF9FAFB" },
                };
            }
        }
    }
    // ── Column widths ─────────────────────────────────────────────────────────
    worksheet.getColumn("A").width = 28; // Reg No — wide enough for E024-01-0001/2024
    worksheet.getColumn("B").width = 40; // Full Name
    worksheet.getColumn("C").width = 52; // Program name — longest field
    worksheet.getColumn("D").width = 18; // Year of Study
    worksheet.getColumn("E").width = 12; // Intake
    // ── Freeze header rows ────────────────────────────────────────────────────
    worksheet.views = [{ state: "frozen", ySplit: headerRowNum }];
    // ── Sheet protection — lock structure but allow data entry ────────────────
    worksheet.protect("", {
        selectLockedCells: true,
        selectUnlockedCells: true,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
    });
    // ── Generate buffer and send ──────────────────────────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();
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
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="${filename}"`)
        .header("Access-Control-Expose-Headers", "Content-Disposition")
        .send(Buffer.from(buffer));
    await (0, auditLogger_1.logAudit)(req, {
        action: "template_download",
        details: { type: "student_registration", programId, universityName, deptCode: myDeptCode },
    });
}));
// ── POST /students/bulk ────────────────────────────────────────────────────────
router.post("/bulk", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { students } = req.body;
    if (!Array.isArray(students) || students.length === 0) {
        res.status(400).json({ message: "No students provided" });
        return;
    }
    const institutionId = req.user.institution;
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    const incoming = students.map(s => ({
        regNo: (s.regNo ?? "").trim().toUpperCase(),
        name: (s.name ?? "").trim(),
        rawProgram: (s.program ?? "").trim(),
        normalizedProgram: (0, programNormalizer_1.normalizeProgramName)((s.program ?? "").trim()),
        yearOfStudy: Number(s.currentYearOfStudy) || 1,
        academicYearId: s.academicYearId,
        intake: (s.intake ?? "SEPT").trim().toUpperCase(),
        admissionAcademicYearString: s.admissionAcademicYearString ?? "2024/2025",
    }));
    const invalid = incoming.filter(s => !s.regNo || !s.name || !s.rawProgram);
    if (invalid.length > 0) {
        res.status(400).json({ message: "Missing Reg No, Name, or Program" });
        return;
    }
    const dbPrograms = await Program_1.default.find({
        institution: institutionId,
        _id: { $in: allowedProgramIds },
    })
        .select("name code durationYears degreeType")
        .lean();
    const programNameMap = new Map(dbPrograms.map(p => [(0, programNormalizer_1.normalizeProgramName)(p.name), p]));
    const programIdMap = new Map(dbPrograms.map(p => [p._id.toString(), p]));
    const missingPrograms = [...new Set(incoming.map(s => s.rawProgram))].filter(raw => !programIdMap.has(raw) && !programNameMap.has((0, programNormalizer_1.normalizeProgramName)(raw)));
    if (missingPrograms.length > 0) {
        res.status(400).json({
            message: "Programs not found or not in your scope",
            notFound: missingPrograms,
        });
        return;
    }
    // Resolve academic years
    const academicYearMap = new Map();
    const yearStrings = [...new Set(incoming.filter(s => !s.academicYearId).map(s => s.admissionAcademicYearString))];
    if (yearStrings.length > 0) {
        const bulkOps = yearStrings.map(yearStr => {
            const [startYear, endYear] = yearStr.split("/").map(Number);
            return {
                updateOne: {
                    filter: { year: yearStr, institution: institutionId },
                    update: { $setOnInsert: {
                            year: yearStr, institution: institutionId,
                            startDate: new Date(`${startYear}-08-01`),
                            endDate: new Date(`${endYear}-07-31`),
                            isCurrent: false,
                        } },
                    upsert: true,
                },
            };
        });
        await AcademicYear_1.default.bulkWrite(bulkOps);
        const resolvedYears = await AcademicYear_1.default.find({
            institution: institutionId, year: { $in: yearStrings },
        })
            .select("year")
            .lean();
        resolvedYears.forEach(y => academicYearMap.set(y.year, y._id));
    }
    const regNos = incoming.map(s => s.regNo);
    const existing = await Student_1.default.find({ regNo: { $in: regNos }, institution: institutionId })
        .select("regNo")
        .lean();
    const existingSet = new Set(existing.map(s => s.regNo));
    const results = {
        registered: [],
        duplicates: [],
        errors: [],
    };
    const toCreate = [];
    for (const s of incoming) {
        if (existingSet.has(s.regNo)) {
            results.duplicates.push(s.regNo);
            continue;
        }
        const progDoc = programIdMap.get(s.rawProgram) ?? programNameMap.get(s.normalizedProgram);
        if (!progDoc) {
            results.errors.push(`${s.regNo}: Program not found`);
            continue;
        }
        // Reg-no validation (server-side, authoritative)
        const validation = await (0, validateRegNo_1.validateRegNo)(s.regNo, institutionId.toString(), progDoc._id.toString());
        if (!validation.valid) {
            results.errors.push(`${s.regNo}: ${validation.reason ?? "Invalid reg no format"}`);
            continue;
        }
        const entryType = s.yearOfStudy === 2 ? "Mid-Entry-Y2" :
            s.yearOfStudy === 3 ? "Mid-Entry-Y3" :
                s.yearOfStudy === 4 ? "Mid-Entry-Y4" : "Direct";
        const programType = progDoc.degreeType ?? "BSc";
        let finalYearId;
        if (s.academicYearId && mongoose_1.default.Types.ObjectId.isValid(s.academicYearId)) {
            finalYearId = new mongoose_1.default.Types.ObjectId(s.academicYearId);
        }
        else {
            const resolved = academicYearMap.get(s.admissionAcademicYearString);
            if (!resolved) {
                results.errors.push(`${s.regNo}: Academic year not resolved`);
                continue;
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
            const created = await Student_1.default.insertMany(toCreate, { ordered: false });
            results.registered.push(...created.map(c => String(c.regNo)));
        }
        catch (err) {
            const bwe = err;
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
    await (0, auditLogger_1.logAudit)(req, {
        action: "students_bulk_registered",
        details: {
            registered: results.registered.length,
            duplicates: results.duplicates.length,
            errors: results.errors.length,
        },
    });
    res.status(207).json({
        message: `${results.registered.length} registered, ${results.duplicates.length} duplicates, ${results.errors.length} errors.`,
        ...results,
    });
}));
// ── DELETE /students/:id ───────────────────────────────────────────────────────
router.delete("/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    const student = await Student_1.default.findOneAndDelete({
        _id: req.params.id,
        institution: req.user.institution,
        program: { $in: allowedProgramIds },
    });
    if (!student) {
        throw { statusCode: 404, message: "Student not found or outside your scope" };
    }
    await (0, auditLogger_1.logAudit)(req, { action: "delete_student", details: { regNo: student.regNo } });
    res.json({ message: "Student deleted successfully" });
}));
// ── PATCH /students/:id ────────────────────────────────────────────────────────
router.patch("/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { name } = req.body;
    if (!name?.trim()) {
        throw { statusCode: 400, message: "Name is required" };
    }
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    const student = await Student_1.default.findOneAndUpdate({ _id: req.params.id, institution: req.user.institution, program: { $in: allowedProgramIds } }, { $set: { name: name.trim() } }, { new: true }).select("name regNo");
    if (!student) {
        throw { statusCode: 404, message: "Student not found or outside your scope" };
    }
    await (0, auditLogger_1.logAudit)(req, { action: "update_student_name", details: { regNo: student.regNo } });
    res.json(student);
}));
// ── DELETE /students/bulk/by-program ──────────────────────────────────────────
router.delete("/bulk/by-program", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId } = req.body;
    if (!programId)
        throw { statusCode: 400, message: "Program ID required" };
    const result = await Student_1.default.deleteMany({
        program: programId,
        institution: req.user.institution,
    });
    await (0, auditLogger_1.logAudit)(req, { action: "bulk_delete_program_students", details: { programId, count: result.deletedCount } });
    res.json({ message: `Deleted ${result.deletedCount} students from program.` });
}));
// ── POST /students/upload-excel ───────────────────────────────────────────────
router.post("/upload-excel", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), uploadStudentExcelMiddleware.single("file"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    if (!req.file) {
        res.status(400).json({ message: "No file uploaded." });
        return;
    }
    const institutionId = req.user.institution;
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    // ── Parse workbook ────────────────────────────────────────────────────────
    const workbook = xlsxLib.read(req.file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawAll = xlsxLib.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
    });
    // Find header row — scan for the cell that says "Reg No"
    let headerIndex = -1;
    for (let i = 0; i < rawAll.length; i++) {
        const row = rawAll[i];
        const cell = (row[0] ?? "").toString().toLowerCase().trim();
        if (cell === "reg no" || cell === "reg. no." || cell === "regno") {
            headerIndex = i;
            break;
        }
    }
    if (headerIndex === -1) {
        res.status(400).json({
            message: "Header row not found. Ensure you downloaded the registration template and that column A header is 'Reg No'.",
        });
        return;
    }
    const dataRows = rawAll.slice(headerIndex + 1).filter(r => r.some(c => String(c ?? "").trim() !== ""));
    if (dataRows.length === 0) {
        res.status(400).json({ message: "No student data found in the file." });
        return;
    }
    // ── Resolve allowed programs ──────────────────────────────────────────────
    const dbPrograms = await Program_1.default.find({
        institution: institutionId,
        _id: { $in: allowedProgramIds },
        isActive: true,
    })
        .select("name code durationYears degreeType")
        .lean();
    const programByName = new Map(dbPrograms.map(p => [(0, programNormalizer_1.normalizeProgramName)(p.name), p]));
    const programByCode = new Map(dbPrograms.map(p => [p.code.toUpperCase(), p]));
    // ── Get current academic year (fallback) ──────────────────────────────────
    const currentYear = await AcademicYear_1.default.findOne({
        institution: institutionId,
        isCurrent: true,
    })
        .select("_id year")
        .lean();
    // ── Process rows ──────────────────────────────────────────────────────────
    const results = {
        registered: [],
        duplicates: [],
        errors: [],
        skipped: 0,
    };
    const toCreate = [];
    const regNosToCreate = [];
    for (const [i, row] of dataRows.entries()) {
        const regNo = String(row[0] ?? "").trim().toUpperCase();
        const name = String(row[1] ?? "").trim();
        const progRaw = String(row[2] ?? "").trim();
        const yearRaw = parseInt(String(row[3] ?? "1")) || 1;
        const intake = String(row[4] ?? "SEPT").trim().toUpperCase();
        // Blank rows — skip silently
        if (!regNo && !name) {
            results.skipped++;
            continue;
        }
        const rowNum = headerIndex + i + 2;
        if (!regNo) {
            results.errors.push(`Row ${rowNum}: Missing Reg No`);
            continue;
        }
        if (!name) {
            results.errors.push(`Row ${rowNum} (${regNo}): Missing Name`);
            continue;
        }
        if (!progRaw) {
            results.errors.push(`Row ${rowNum} (${regNo}): Missing Program`);
            continue;
        }
        // Resolve program — by name first, then by code
        const progDoc = programByName.get((0, programNormalizer_1.normalizeProgramName)(progRaw)) ??
            programByCode.get(progRaw.toUpperCase());
        if (!progDoc) {
            results.errors.push(`Row ${rowNum} (${regNo}): Program "${progRaw}" not found in your department scope`);
            continue;
        }
        // Server-side reg-no validation
        const validation = await (0, validateRegNo_1.validateRegNo)(regNo, institutionId.toString(), progDoc._id.toString());
        if (!validation.valid) {
            results.errors.push(`Row ${rowNum} (${regNo}): ${validation.reason}`);
            continue;
        }
        const safeIntake = ["JAN", "MAY", "SEPT"].includes(intake) ? intake : "SEPT";
        const entryType = yearRaw === 2 ? "Mid-Entry-Y2" :
            yearRaw === 3 ? "Mid-Entry-Y3" :
                yearRaw === 4 ? "Mid-Entry-Y4" : "Direct";
        toCreate.push({
            regNo,
            name,
            institution: institutionId,
            program: progDoc._id,
            programType: progDoc.degreeType ?? "BSc",
            entryType,
            intake: safeIntake,
            currentYearOfStudy: yearRaw,
            admissionAcademicYear: currentYear?._id,
            status: "active",
        });
        regNosToCreate.push(regNo);
    }
    // ── Batch duplicate check ─────────────────────────────────────────────────
    if (regNosToCreate.length > 0) {
        const existing = await Student_1.default.find({
            regNo: { $in: regNosToCreate },
            institution: institutionId,
        })
            .select("regNo")
            .lean();
        const existingSet = new Set(existing.map(s => s.regNo));
        const newStudents = toCreate.filter(s => {
            if (existingSet.has(s.regNo)) {
                results.duplicates.push(s.regNo);
                return false;
            }
            return true;
        });
        if (newStudents.length > 0) {
            try {
                const created = await Student_1.default.insertMany(newStudents, { ordered: false });
                results.registered.push(...created.map(c => String(c.regNo)));
            }
            catch (err) {
                const bwe = err;
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
    }
    await (0, auditLogger_1.logAudit)(req, {
        action: "students_excel_upload",
        details: {
            filename: req.file.originalname,
            registered: results.registered.length,
            duplicates: results.duplicates.length,
            errors: results.errors.length,
            skipped: results.skipped,
        },
    });
    const total = results.registered.length +
        results.duplicates.length +
        results.errors.length;
    res.status(207).json({
        message: `${results.registered.length} registered, ${results.duplicates.length} already exist, ${results.errors.length} errors from ${total} rows.`,
        ...results,
    });
}));
exports.default = router;
