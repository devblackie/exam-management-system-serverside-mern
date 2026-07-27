"use strict";
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
// src/routes/marks.ts  — PATCHED
// Key fixes:
//   1. Added POST /marks/upload-direct route so client's templateMode="direct" upload works
//   2. Added detailed console logging around the template-detection switch
//   3. Error handler now logs the full stack before rethrowing
const express_1 = require("express");
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const uploadTemplate_1 = require("../utils/uploadTemplate");
const upload_1 = require("../middleware/upload");
const marksImporter_1 = require("../services/marksImporter");
const auditLogger_1 = require("../lib/auditLogger");
const mongoose_1 = __importDefault(require("mongoose"));
const xlsx = __importStar(require("xlsx"));
const Unit_1 = __importDefault(require("../models/Unit"));
const directTemplate_1 = require("../utils/directTemplate");
const directMarksImporter_1 = require("../services/directMarksImporter");
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const loadLogoBuffer_1 = require("../utils/loadLogoBuffer");
const auth_2 = require("../middleware/auth");
const router = (0, express_1.Router)();
// ─────────────────────────────────────────────────────────────────────────────
// GET /marks/template   — Detailed breakdown scoresheet
// ─────────────────────────────────────────────────────────────────────────────
router.get("/template", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, unitId, academicYearId, yearOfStudy, semester, examMode, unitType } = req.query;
    if (!programId || !unitId || !academicYearId || !yearOfStudy || !semester) {
        return res.status(400).json({
            error: "Missing required parameters: programId, unitId, academicYearId, yearOfStudy, semester",
        });
    }
    try {
        if (!mongoose_1.default.Types.ObjectId.isValid(programId) ||
            !mongoose_1.default.Types.ObjectId.isValid(unitId) ||
            !mongoose_1.default.Types.ObjectId.isValid(academicYearId)) {
            throw new Error("One or more provided IDs (programId, unitId, academicYearId) are invalid.");
        }
        const [unit, academicYear] = await Promise.all([
            Unit_1.default.findById(unitId).lean(),
            AcademicYear_1.default.findById(academicYearId).lean(),
        ]);
        if (!academicYear)
            throw new Error("Academic Year not found.");
        const rawCode = (unit?.code || "UNIT").trim();
        const rawName = (unit?.name || "TEMPLATE").trim();
        const yearLabel = (academicYear.year || "YEAR").trim().replace(/\//g, "-");
        const cleanName = `${rawCode}_${rawName}`
            .replace(/[^a-zA-Z0-9]/g, "_")
            .replace(/_+/g, "_")
            .replace(/^_|_$/g, "")
            ?.toUpperCase() || "TEMPLATE";
        const fileName = `Scoresheet_${cleanName}_${yearLabel}.xlsx`;
        // const logoPath  = path.join(__dirname, "../../public/institutionLogoExcel.png");
        // const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);
        const institutionId = req.user.institution.toString();
        const logoBuffer = await (0, loadLogoBuffer_1.loadLogoBuffer)(institutionId);
        const excelBuffer = await (0, uploadTemplate_1.generateFullScoresheetTemplate)(new mongoose_1.default.Types.ObjectId(programId), new mongoose_1.default.Types.ObjectId(unitId), parseInt(yearOfStudy, 10), parseInt(semester, 10), new mongoose_1.default.Types.ObjectId(academicYearId), logoBuffer, examMode, unitType || "theory");
        res
            .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            .header("Access-Control-Expose-Headers", "Content-Disposition")
            .attachment(fileName)
            .send(excelBuffer);
    }
    catch (error) {
        console.error("[GET /marks/template] Error:", error.message, error.stack);
        if (error.message === "Institution settings not found.") {
            return res.status(400).json({
                message: "Institution settings are not configured. Please contact the administrator.",
                error: error.message,
            });
        }
        const status = error.message.includes("invalid") ? 400 : 500;
        // return res.status(status).json({ message: "Failed to generate scoresheet template.", error: error.message });
        return res.status(status).json({ message: error.message, error: error.message });
    }
}));
// ─────────────────────────────────────────────────────────────────────────────
// GET /marks/direct-template   — Direct entry scoresheet
// ─────────────────────────────────────────────────────────────────────────────
router.get("/direct-template", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, unitId, academicYearId, yearOfStudy, semester } = req.query;
    if (!programId || !unitId || !academicYearId || !yearOfStudy || !semester) {
        return res.status(400).json({
            error: "Missing required parameters: programId, unitId, academicYearId, yearOfStudy, semester",
        });
    }
    try {
        if (!mongoose_1.default.Types.ObjectId.isValid(programId) ||
            !mongoose_1.default.Types.ObjectId.isValid(unitId) ||
            !mongoose_1.default.Types.ObjectId.isValid(academicYearId)) {
            throw new Error("One or more provided IDs (programId, unitId, academicYearId) are invalid.");
        }
        const [unit, academicYear] = await Promise.all([
            Unit_1.default.findById(unitId).lean(),
            AcademicYear_1.default.findById(academicYearId).lean(),
        ]);
        if (!academicYear)
            throw new Error("Academic Year not found.");
        const rawCode = (unit?.code || "UNIT").trim();
        const rawName = (unit?.name || "TEMPLATE").trim();
        const yearLabel = (academicYear.year || "YEAR").trim().replace(/\//g, "-");
        const cleanName = `${rawCode}_${rawName}`
            .replace(/[^a-zA-Z0-9]/g, "_")
            .replace(/_+/g, "_")
            .replace(/^_|_$/g, "")
            ?.toUpperCase() || "TEMPLATE";
        const fileName = `Scoresheet_${cleanName}_${yearLabel}.xlsx`;
        // const logoPath   = path.join(__dirname, "../../public/institutionLogoExcel.png");
        // const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);
        const institutionId = req.user.institution.toString();
        const logoBuffer = await (0, loadLogoBuffer_1.loadLogoBuffer)(institutionId);
        const buffer = await (0, directTemplate_1.generateDirectScoresheetTemplate)(new mongoose_1.default.Types.ObjectId(programId), new mongoose_1.default.Types.ObjectId(unitId), parseInt(yearOfStudy, 10), parseInt(semester, 10), new mongoose_1.default.Types.ObjectId(academicYearId), logoBuffer);
        res
            .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            .header("Access-Control-Expose-Headers", "Content-Disposition")
            .attachment(fileName)
            .send(buffer);
    }
    catch (error) {
        console.error("[GET /marks/direct-template] Error:", error.message, error.stack);
        if (error.message === "Institution settings not found.") {
            return res.status(400).json({
                message: "Institution settings are not configured. Please contact the administrator.",
                error: error.message,
            });
        }
        const status = error.message.includes("invalid") ? 400 : 500;
        // return res.status(status).json({ message: "Failed to generate direct template.", error: error.message });
        return res
            .status(status)
            .json({ message: error.message, error: error.message });
    }
}));
// // ─────────────────────────────────────────────────────────────────────────────
// // GET /marks/upload-stats
// // ─────────────────────────────────────────────────────────────────────────────
// router.get("/upload-stats", requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const institutionId = req.user.institution;
//     const [detailedRaw, directRaw] = await Promise.all([
//       Mark.find({ institution: institutionId, deletedAt: null })
//         .populate({ path: "programUnit", populate: [{ path: "unit", select: "code name" }, { path: "program", select: "name code" }] })
//         .populate("academicYear", "year session")
//         .populate("student", "regNo name")
//         .sort({ createdAt: -1 })
//         .lean(),
//       MarkDirect.find({ institution: institutionId, deletedAt: null })
//         .populate({ path: "programUnit", populate: [{ path: "unit", select: "code name" }, { path: "program", select: "name code" }] })
//         .populate("academicYear", "year session")
//         .populate("student", "regNo name")
//         .sort({ createdAt: -1 })
//         .lean(),
//     ]);
//     interface MarkEntry {
//       _id: string; source: "detailed" | "direct"; regNo: string; studentName: string;
//       unitCode: string; unitName: string; programName: string; programCode: string;
//       agreedMark: number; attempt: string; isSpecial: boolean;
//       academicYear: string; session: string; uploadedAt: Date;
//     }
//     const shape = (m: any, source: "detailed" | "direct"): MarkEntry => ({
//       _id:         m._id.toString(),
//       source,
//       regNo:       (m.student as any)?.regNo   || "N/A",
//       studentName: (m.student as any)?.name    || "N/A",
//       unitCode:    (m.programUnit as any)?.unit?.code    || "N/A",
//       unitName:    (m.programUnit as any)?.unit?.name    || "N/A",
//       programName: (m.programUnit as any)?.program?.name || "N/A",
//       programCode: (m.programUnit as any)?.program?.code || "N/A",
//       agreedMark:  m.agreedMark ?? 0,
//       attempt:     m.attempt ?? "1st",
//       isSpecial:   m.isSpecial ?? false,
//       academicYear:(m.academicYear as any)?.year    || "Unknown",
//       session:     (m.academicYear as any)?.session || "ORDINARY",
//       uploadedAt:  m.uploadedAt ?? m.createdAt,
//     });
//     const allEntries: MarkEntry[] = [
//       ...detailedRaw.map((m) => shape(m, "detailed")),
//       ...directRaw.map((m) => shape(m, "direct")),
//     ];
//     const grouped: Record<string, Record<string, Record<string, { programName: string; entries: MarkEntry[] }>>> = {};
//     for (const entry of allEntries) {
//       const yr  = entry.academicYear;
//       const ses = entry.session;
//       const pc  = entry.programCode;
//       if (!grouped[yr])          grouped[yr]          = {};
//       if (!grouped[yr][ses])     grouped[yr][ses]     = {};
//       if (!grouped[yr][ses][pc]) grouped[yr][ses][pc] = { programName: entry.programName, entries: [] };
//       grouped[yr][ses][pc].entries.push(entry);
//     }
//     const summary = {
//       totalRecords:  allEntries.length,
//       detailed:      detailedRaw.length,
//       direct:        directRaw.length,
//       academicYears: Object.keys(grouped).sort().reverse(),
//     };
//     res.json({ summary, grouped });
//   }),
// );
// ─────────────────────────────────────────────────────────────────────────────
// GET /marks/upload-stats
// ─────────────────────────────────────────────────────────────────────────────
router.get("/upload-stats", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institutionId = req.user.institution;
    const [detailedRaw, directRaw] = await Promise.all([
        Mark_1.default.find({ institution: institutionId, deletedAt: null })
            .populate({
            path: "programUnit",
            populate: [
                { path: "unit", select: "code name" },
                { path: "program", select: "name code" },
            ],
        })
            .populate("academicYear", "year session")
            .populate("student", "regNo name")
            .sort({ createdAt: -1 })
            .lean(),
        MarkDirect_1.default.find({ institution: institutionId, deletedAt: null })
            .populate({
            path: "programUnit",
            populate: [
                { path: "unit", select: "code name" },
                { path: "program", select: "name code" },
            ],
        })
            .populate("academicYear", "year session")
            .populate("student", "regNo name")
            .sort({ createdAt: -1 })
            .lean(),
    ]);
    const shape = (m, source) => {
        const programUnit = m.programUnit;
        const unit = programUnit?.unit;
        const program = programUnit?.program;
        const student = m.student;
        const academicYear = m.academicYear;
        const unitCode = String(unit?.code ?? "N/A");
        const uploadedDate = new Date(m.uploadedAt ?? m.createdAt ?? Date.now())
            .toISOString()
            .split("T")[0];
        return {
            _id: String(m._id ?? ""),
            source,
            batchId: typeof m.batchId === "string" && m.batchId.length > 10
                ? m.batchId
                : `${unitCode}|${uploadedDate}|${String(m.attempt ?? "1st")}`,
            regNo: String(student?.regNo ?? "N/A"),
            studentName: String(student?.name ?? "N/A"),
            unitCode,
            unitName: String(unit?.name ?? "N/A"),
            programName: String(program?.name ?? "N/A"),
            programCode: String(program?.code ?? "N/A"),
            agreedMark: Number(m.agreedMark ?? 0),
            attempt: String(m.attempt ?? "1st"),
            isSpecial: Boolean(m.isSpecial ?? false),
            academicYear: String(academicYear?.year ?? "Unknown"),
            session: String(academicYear?.session ?? "ORDINARY"),
            uploadedAt: new Date(m.uploadedAt ?? m.createdAt ?? Date.now()),
        };
    };
    const allEntries = [
        ...detailedRaw.map((m) => shape(m, "detailed")),
        ...directRaw.map((m) => shape(m, "direct")),
    ];
    const grouped = {};
    for (const entry of allEntries) {
        const yr = entry.academicYear;
        const ses = entry.session;
        const pc = entry.programCode;
        if (!grouped[yr])
            grouped[yr] = {};
        if (!grouped[yr][ses])
            grouped[yr][ses] = {};
        if (!grouped[yr][ses][pc]) {
            grouped[yr][ses][pc] = { programName: entry.programName, entries: [] };
        }
        grouped[yr][ses][pc].entries.push(entry);
    }
    const summary = {
        totalRecords: allEntries.length,
        detailed: detailedRaw.length,
        direct: directRaw.length,
        academicYears: Object.keys(grouped).sort().reverse(),
    };
    res.json({ summary, grouped });
}));
// ─────────────────────────────────────────────────────────────────────────────
// POST /marks/upload   — Auto-detects template type by inspecting cell E15
// ─────────────────────────────────────────────────────────────────────────────
router.post("/upload", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), upload_1.uploadMarksFile.single("file"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ message: "No file uploaded" });
    }
    console.log(`[POST /marks/upload] File received: "${req.file.originalname}", size=${req.file.size}`);
    try {
        const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        // Cell E15 contains "CA TOTAL (/30)" in the direct template header row
        const columnEHeader = sheet["E15"]?.v?.toString().toUpperCase() || "";
        console.log(`[POST /marks/upload] Cell E15 value: "${columnEHeader}"`);
        let result;
        let templateType;
        if (columnEHeader.includes("CA TOTAL")) {
            templateType = "direct";
            console.log(`[POST /marks/upload] Detected DIRECT ENTRY template`);
            result = await (0, directMarksImporter_1.importDirectMarksFromBuffer)(req.file.buffer, req.file.originalname, req);
        }
        else {
            templateType = "detailed";
            console.log(`[POST /marks/upload] Detected DETAILED BREAKDOWN template`);
            result = await (0, marksImporter_1.importMarksFromBuffer)(req.file.buffer, req.file.originalname, req);
        }
        console.log(`[POST /marks/upload] Import result: total=${result.total}, success=${result.success}, errors=${result.errors.length}`);
        await (0, auditLogger_1.logAudit)(req, {
            action: "marks_upload_success",
            details: { templateType, total: result.total, success: result.success },
        });
        res.json({ message: "Import completed", ...result });
    }
    catch (err) {
        console.error("[POST /marks/upload] Fatal error:", err.message, err.stack);
        await (0, auditLogger_1.logAudit)(req, {
            action: "marks_upload_failed",
            details: { error: err.message, filename: req.file.originalname },
        });
        throw err;
    }
}));
// ─────────────────────────────────────────────────────────────────────────────
// POST /marks/upload-direct  — Explicit direct-entry upload endpoint
// (The client marksApi.ts calls this when templateMode === "direct")
// ─────────────────────────────────────────────────────────────────────────────
router.post("/upload-direct", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), upload_1.uploadMarksFile.single("file"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ message: "No file uploaded" });
    }
    console.log(`[POST /marks/upload-direct] File: "${req.file.originalname}", size=${req.file.size}`);
    try {
        const result = await (0, directMarksImporter_1.importDirectMarksFromBuffer)(req.file.buffer, req.file.originalname, req);
        console.log(`[POST /marks/upload-direct] Done: total=${result.total}, success=${result.success}, errors=${result.errors.length}`);
        await (0, auditLogger_1.logAudit)(req, {
            action: "marks_upload_success",
            details: { templateType: "direct", total: result.total, success: result.success },
        });
        res.json({ message: "Import completed", ...result });
    }
    catch (err) {
        console.error("[POST /marks/upload-direct] Fatal error:", err.message, err.stack);
        await (0, auditLogger_1.logAudit)(req, {
            action: "marks_upload_failed",
            details: { error: err.message, filename: req.file.originalname },
        });
        throw err;
    }
}));
// ─────────────────────────────────────────────────────────────────────────────
// GET /marks/batch/:batchId — View marks inside a single scoresheet
// ─────────────────────────────────────────────────────────────────────────────
router.get("/batch/:batchId", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { batchId } = req.params;
    const institutionId = req.user.institution;
    if (!batchId || batchId.length < 10) {
        return res.status(400).json({ message: "Invalid batch ID." });
    }
    // Fetch both detailed and direct marks
    const [detailedMarks, directMarks] = await Promise.all([
        Mark_1.default.find({ batchId, institution: institutionId, deletedAt: null })
            .populate({ path: "programUnit", populate: [{ path: "unit", select: "code name" }, { path: "program", select: "name code departmentCode schoolCode" }] })
            .populate("academicYear", "year session")
            .populate("student", "regNo name")
            .sort({ "student.regNo": 1 })
            .lean(),
        MarkDirect_1.default.find({ batchId, institution: institutionId, deletedAt: null })
            .populate({ path: "programUnit", populate: [{ path: "unit", select: "code name" }, { path: "program", select: "name code departmentCode schoolCode" }] })
            .populate("academicYear", "year session")
            .populate("student", "regNo name")
            .sort({ "student.regNo": 1 })
            .lean(),
    ]);
    const allMarks = [...detailedMarks, ...directMarks];
    if (allMarks.length === 0) {
        return res.status(404).json({ message: "No marks found for this batch." });
    }
    // Scope check: coordinator can only view batches in their department/school
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
        const scopedProgramIds = await (0, auth_2.getScopedProgramIds)(req);
        const scopedIdStrings = scopedProgramIds.map((id) => id.toString());
        for (const mark of allMarks) {
            const programId = mark.programUnit?.program?._id?.toString();
            if (programId && !scopedIdStrings.includes(programId)) {
                return res.status(403).json({
                    message: "This batch contains marks from programs outside your department.",
                });
            }
        }
    }
    const firstMark = allMarks[0];
    const batchInfo = {
        batchId,
        unitCode: firstMark.programUnit?.unit?.code || "N/A",
        unitName: firstMark.programUnit?.unit?.name || "N/A",
        programCode: firstMark.programUnit?.program?.code || "N/A",
        programName: firstMark.programUnit?.program?.name || "N/A",
        academicYear: firstMark.academicYear?.year || "N/A",
        session: firstMark.academicYear?.session || "ORDINARY",
        totalRecords: allMarks.length,
        uploadedAt: firstMark.uploadedAt || firstMark.createdAt,
    };
    const entries = allMarks.map((m) => ({
        _id: m._id.toString(),
        regNo: m.student?.regNo || "N/A",
        studentName: m.student?.name || "N/A",
        caTotal30: m.caTotal30 ?? 0,
        examTotal70: m.examTotal70 ?? 0,
        agreedMark: m.agreedMark ?? 0,
        attempt: m.attempt ?? "1st",
        isSpecial: m.isSpecial ?? false,
    }));
    res.json({ batch: batchInfo, entries });
}));
// ─────────────────────────────────────────────────────────────────────────────
// DELETE /marks/batch/:batchId — Soft-delete a single scoresheet
// ─────────────────────────────────────────────────────────────────────────────
router.delete("/batch/:batchId", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { batchId } = req.params;
    const institutionId = req.user.institution;
    if (!batchId || batchId.length < 10) {
        return res.status(400).json({ message: "Invalid batch ID." });
    }
    // Find all marks in this batch
    const [detailedMarks, directMarks] = await Promise.all([
        Mark_1.default.find({ batchId, institution: institutionId, deletedAt: null }).populate({
            path: "programUnit",
            populate: { path: "program", select: "_id departmentCode schoolCode" },
        }),
        MarkDirect_1.default.find({ batchId, institution: institutionId, deletedAt: null }).populate({
            path: "programUnit",
            populate: { path: "program", select: "_id departmentCode schoolCode" },
        }),
    ]);
    const allMarks = [...detailedMarks, ...directMarks];
    if (allMarks.length === 0) {
        return res.status(404).json({ message: "No marks found for this batch." });
    }
    // Scope check: coordinator can only delete batches in their department/school
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
        const scopedProgramIds = await (0, auth_2.getScopedProgramIds)(req);
        const scopedIdStrings = scopedProgramIds.map((id) => id.toString());
        const outOfScope = allMarks.filter((m) => {
            const programId = m.programUnit?.program?._id?.toString();
            return programId && !scopedIdStrings.includes(programId);
        });
        if (outOfScope.length > 0) {
            return res.status(403).json({
                message: `Cannot delete this batch. ${outOfScope.length} records belong to programs outside your department.`,
            });
        }
    }
    const now = new Date();
    // Soft-delete all marks in the batch
    const [detailedResult, directResult] = await Promise.all([
        Mark_1.default.updateMany({ batchId, institution: institutionId, deletedAt: null }, { $set: { deletedAt: now } }),
        MarkDirect_1.default.updateMany({ batchId, institution: institutionId, deletedAt: null }, { $set: { deletedAt: now } }),
    ]);
    const totalDeleted = detailedResult.modifiedCount + directResult.modifiedCount;
    await (0, auditLogger_1.logAudit)(req, {
        action: "marks_batch_deleted",
        details: {
            batchId,
            detailedDeleted: detailedResult.modifiedCount,
            directDeleted: directResult.modifiedCount,
            totalDeleted,
        },
    });
    res.json({
        message: `Batch deleted successfully. ${totalDeleted} records removed.`,
        deletedCount: totalDeleted,
    });
}));
// ─────────────────────────────────────────────────────────────────────────────
// DELETE /marks/batches — Bulk soft-delete multiple scoresheets
// ─────────────────────────────────────────────────────────────────────────────
router.delete("/batches", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { batchIds } = req.body;
    const institutionId = req.user.institution;
    if (!batchIds || !Array.isArray(batchIds) || batchIds.length === 0) {
        return res.status(400).json({ message: "batchIds array is required." });
    }
    if (batchIds.length > 20) {
        return res.status(400).json({ message: "Maximum 20 batches per request." });
    }
    // Validate all batchIds are non-empty strings
    if (batchIds.some((id) => typeof id !== "string" || id.length < 10)) {
        return res.status(400).json({ message: "One or more batch IDs are invalid." });
    }
    // Fetch all marks for all batches
    const [detailedMarks, directMarks] = await Promise.all([
        Mark_1.default.find({
            batchId: { $in: batchIds },
            institution: institutionId,
            deletedAt: null,
        }).populate({
            path: "programUnit",
            populate: { path: "program", select: "_id departmentCode schoolCode" },
        }),
        MarkDirect_1.default.find({
            batchId: { $in: batchIds },
            institution: institutionId,
            deletedAt: null,
        }).populate({
            path: "programUnit",
            populate: { path: "program", select: "_id departmentCode schoolCode" },
        }),
    ]);
    const allMarks = [...detailedMarks, ...directMarks];
    if (allMarks.length === 0) {
        return res.status(404).json({ message: "No marks found for these batches." });
    }
    // Scope check
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
        const scopedProgramIds = await (0, auth_2.getScopedProgramIds)(req);
        const scopedIdStrings = scopedProgramIds.map((id) => id.toString());
        const outOfScope = allMarks.filter((m) => {
            const programId = m.programUnit?.program?._id?.toString();
            return programId && !scopedIdStrings.includes(programId);
        });
        if (outOfScope.length > 0) {
            return res.status(403).json({
                message: `Cannot delete. ${outOfScope.length} records belong to programs outside your department.`,
            });
        }
    }
    const now = new Date();
    const [detailedResult, directResult] = await Promise.all([
        Mark_1.default.updateMany({ batchId: { $in: batchIds }, institution: institutionId, deletedAt: null }, { $set: { deletedAt: now } }),
        MarkDirect_1.default.updateMany({ batchId: { $in: batchIds }, institution: institutionId, deletedAt: null }, { $set: { deletedAt: now } }),
    ]);
    const totalDeleted = detailedResult.modifiedCount + directResult.modifiedCount;
    await (0, auditLogger_1.logAudit)(req, {
        action: "marks_bulk_batch_deleted",
        details: {
            batchIds,
            batchCount: batchIds.length,
            totalDeleted,
        },
    });
    res.json({
        message: `${batchIds.length} batches deleted. ${totalDeleted} records removed.`,
        deletedCount: totalDeleted,
        batchCount: batchIds.length,
    });
}));
exports.default = router;
