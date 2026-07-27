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
// serverside/src/routes/promote.ts
const express_1 = require("express");
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const statusEngine_1 = require("../services/statusEngine");
const Program_1 = __importDefault(require("../models/Program"));
const promotionReport_1 = require("../utils/promotionReport");
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const adm_zip_1 = __importDefault(require("adm-zip"));
const Student_1 = __importDefault(require("../models/Student"));
const auditLogger_1 = require("../lib/auditLogger");
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const Mark_1 = __importDefault(require("../models/Mark"));
const consolidatedMS_1 = require("../utils/consolidatedMS");
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const undoPromotion_1 = require("../services/undoPromotion");
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const loadInstitutionSettings_1 = require("../utils/loadInstitutionSettings");
const router = (0, express_1.Router)();
// preview-promotion
router.post("/preview-promotion", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    if (!programId || !yearToPromote || !academicYearName)
        return res.status(400).json({ error: "Missing parameters" });
    const previewData = await (0, statusEngine_1.previewPromotion)(programId, yearToPromote, academicYearName);
    res.json({ success: true, data: previewData });
}));
// bulk-promote
router.post("/bulk-promote", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    if (!programId || !yearToPromote || !academicYearName)
        return res.status(400).json({ error: "Missing required promotion parameters" });
    const results = await (0, statusEngine_1.bulkPromoteClass)(programId, yearToPromote, academicYearName);
    res.json({ success: true, message: `Process completed: ${results.promoted} promoted, ${results.failed} failed.`, data: results });
}));
// download-report-progress
router.post("/download-report-progress", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");
    const sendProgress = (percent, message, file) => { const data = JSON.stringify({ percent, message, file }); res.write(`data: ${data}\n\n`); };
    try {
        sendProgress(10, "Fetching student data and raw marks...");
        // 1. Fetch Basic Data
        const preview = await (0, statusEngine_1.previewPromotion)(programId, yearToPromote, academicYearName);
        const program = await Program_1.default.findById(programId).lean();
        const academicYearDoc = await AcademicYear_1.default.findOne({ year: academicYearName }).lean();
        const targetAcadYearId = academicYearDoc?._id?.toString();
        const academicYearDocForSession = await AcademicYear_1.default.findOne({ year: academicYearName }).lean();
        const institutionSettings = await InstitutionSettings_1.default.findOne({ institution: program?.institution }).lean();
        // const passMark = institutionSettings?.passMark ?? 40;
        const settings = await (0, loadInstitutionSettings_1.loadInstitutionSettings)(program?.institution?.toString() || "");
        const passMark = settings.passMark;
        const gradingScale = institutionSettings?.gradingScale ?? [];
        const sessionExamType = academicYearDocForSession?.session === "SUPPLEMENTARY" ? "SUPPLEMENTARY" : "ORDINARY";
        const logoPath = path_1.default.join(__dirname, "../../public/institutionLogoExcel.png");
        const logoBuffer = fs_1.default.existsSync(logoPath) ? fs_1.default.readFileSync(logoPath) : Buffer.alloc(0);
        // 3. Fetch Marks
        const allStudents = [...preview.eligible, ...preview.blocked];
        const studentIds = allStudents.map((s) => { const id = s._id || s.student?._id || s.id || s.student; return id?.toString(); }).filter((id) => id && id.length >= 24);
        // const rawMarks = await Mark.find({ student: { $in: studentIds } }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" }}).lean();
        // const filteredMarks = rawMarks.filter((m: any) => { return m.programUnit && Number(m.programUnit.requiredYear) === Number(yearToPromote); });
        // Fetch from both Detailed (Mark) and Direct (MarkDirect) models
        const [detailedMarks, directMarks] = await Promise.all([
            Mark_1.default.find({ student: { $in: studentIds } }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" } }).lean(),
            MarkDirect_1.default.find({ student: { $in: studentIds } }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" } }).lean()
        ]);
        // Merge them: Direct marks often take precedence if duplicates exist
        const combinedMarks = [...detailedMarks, ...directMarks];
        // Filter marks strictly for the Year of Study being promoted
        // const filteredMarks = combinedMarks.filter((m: any) => { return m.programUnit && Number(m.programUnit.requiredYear) === Number(yearToPromote);});
        const filteredMarks = combinedMarks.filter((m) => {
            const rightYear = m.programUnit && Number(m.programUnit.requiredYear) === Number(yearToPromote);
            const rightCohort = !targetAcadYearId || (m.academicYear?.toString() === targetAcadYearId);
            return rightYear && rightCohort;
        });
        const offeredUnitsRaw = await ProgramUnit_1.default.find({ program: programId, requiredYear: yearToPromote }).populate("unit").lean();
        const offeredUnits = offeredUnitsRaw.map((pu) => ({ code: pu.unit?.code || "N/A", name: pu.unit?.name || "N/A" }));
        // 5. Prepare Data Objects (Separated for type safety)
        const wordData = {
            programName: program?.name || "Program",
            academicYear: academicYearName,
            yearOfStudy: yearToPromote,
            eligible: preview.eligible,
            blocked: preview.blocked,
            offeredUnits,
            logoBuffer,
            examType: sessionExamType,
            institutionId: program?.institution?.toString() || "",
        };
        const studentsByHistory = await Student_1.default.find({
            program: programId,
            status: {
                $nin: ["graduated", "graduand", "discontinued", "deregistered"],
            }, // ← KEY FIX
            $or: [
                { currentYearOfStudy: yearToPromote },
                {
                    academicHistory: {
                        $elemMatch: { yearOfStudy: yearToPromote, academicYear: academicYearName },
                    },
                },
            ],
        }).lean();
        // Deduplicate — preview already has blocked (on_leave, deferred etc).
        // Merge so every student who ever touched this year is on the CMS.
        const previewIds = new Set([...preview.eligible, ...preview.blocked].map((s) => (s.id || s._id)?.toString()));
        const historyOnly = studentsByHistory.filter((s) => !previewIds.has(s._id.toString()));
        const allCmsStudents = [...preview.eligible, ...preview.blocked, ...historyOnly];
        const excelData = {
            programName: program?.name || "Program", academicYear: academicYearName, yearOfStudy: yearToPromote,
            session: sessionExamType, students: allCmsStudents, marks: filteredMarks, offeredUnits, logoBuffer,
            institutionId: program?.institution?.toString() || "", programId: programId, passMark, gradingScale,
        };
        // 6. Generate and Zip reports
        const cleanAcadYear = academicYearName.replace(/\//g, "_");
        const progCode = program?.code || "PROG";
        const progName = program?.name || "Program";
        const yearPrefix = `Year_${yearToPromote}`;
        const getFileName = (reportType) => `${reportType}_${progCode}_${progName}_${cleanAcadYear}_${yearPrefix}.docx`.replace(/\s+/g, "_");
        const zip = new adm_zip_1.default();
        // Helper to conditionally add documents
        const addDocIfNotEmpty = async (list, fileName, generator, ...extraArgs) => {
            if (list && list.length > 0) {
                const buffer = await generator(wordData, ...extraArgs);
                zip.addFile(fileName, buffer);
                return true;
            }
            return false;
        };
        sendProgress(30, "Generating Main Summary & Marksheet...");
        // zip.addFile(`Summary_Ordinary_Exams_${progCode}_${progName}_${yearPrefix}_${cleanAcadYear}.docx`, await generatePromotionWordDoc(wordData));
        // zip.addFile(`${progName}__${progCode}_${cleanAcadYear}_${yearPrefix}_CMS.xlsx`, await generateConsolidatedMarkSheet(excelData));
        const summaryPrefix = sessionExamType === "SUPPLEMENTARY" ? "Summary_Supp_Special" : "Summary_Ordinary";
        zip.addFile(`${summaryPrefix}_${progCode}_${progName}_${yearPrefix}_${cleanAcadYear}.docx`, await (0, promotionReport_1.generatePromotionWordDoc)(wordData));
        zip.addFile(`${progName}__${progCode}_${cleanAcadYear}_${yearPrefix}_CMS.xlsx`, await (0, consolidatedMS_1.generateConsolidatedMarkSheet)(excelData));
        sendProgress(40, "Checking Pass List...");
        await addDocIfNotEmpty(wordData.eligible, getFileName("PASS_LIST"), promotionReport_1.generateEligibleSummaryDoc);
        sendProgress(50, "Checking Supplementary List...");
        // const suppList = wordData.blocked.filter(s => s.status.includes("SUPP"));
        // await addDocIfNotEmpty(suppList, getFileName("Supplementary_List"), generateSupplementaryExamsDoc);
        // 1. First-attempt SUPP (not from prior hurdle)
        const suppFirstAttempt = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            !s.reasons?.some((r) => /stayout|a\/so|repeat\s+year|a\/ra|readmission|readmit|carry\s*forward|a\/cf|academic\s*leave/i.test(r)));
        await addDocIfNotEmpty(suppFirstAttempt, getFileName("Supplementary_List"), promotionReport_1.generateSupplementaryExamsDoc);
        // 2. Supp after Carry Forward (A/CFS)
        const suppCF = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            s.reasons?.some((r) => /carry\s*forward|a\/cf/i.test(r)));
        await addDocIfNotEmpty(suppCF, getFileName("Supplementary_After_Carry_Forward"), promotionReport_1.generateCarryForwardSuppDoc);
        // 3. Supp after Stay Out (A/SOS)
        const suppStayout = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            s.reasons?.some((r) => /stayout|a\/so/i.test(r)));
        await addDocIfNotEmpty(suppStayout, getFileName("Supplementary_After_Stayout"), promotionReport_1.generateStayoutSuppDoc);
        // 4. Supp after Repeat Year (A/RA)
        const suppRepeat = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            s.reasons?.some((r) => /repeat\s+year|a\/ra/i.test(r)));
        await addDocIfNotEmpty(suppRepeat, getFileName("Supplementary_After_Repeat_Year"), promotionReport_1.generateRepeatSuppDoc);
        // 5. Supp after Readmission
        const suppReadmit = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            s.reasons?.some((r) => /readmission|readmit/i.test(r)));
        await addDocIfNotEmpty(suppReadmit, getFileName("Supplementary_After_Readmission"), promotionReport_1.generateReadmissionSuppDoc);
        // 6. Supp after Academic Leave
        const suppLeave = wordData.blocked.filter((s) => s.status.includes("SUPP") &&
            s.reasons?.some((r) => /academic\s*leave|on\s*leave|a\/sp/i.test(r)));
        await addDocIfNotEmpty(suppLeave, getFileName("Supplementary_After_Academic_Leave"), promotionReport_1.generateAcademicLeaveSuppDoc);
        sendProgress(60, "Checking Special Exams...");
        const getSpecialGround = (s) => {
            const grounds = (s.specialGrounds || "").toLowerCase();
            const remarks = (s.remarks || "").toLowerCase();
            const leaveType = (s.academicLeavePeriod?.type || "").toLowerCase();
            const details = (s.details || "").toLowerCase();
            return `${grounds} ${remarks} ${leaveType} ${details}`;
        };
        const isSpecialStudent = (s) => /spec/i.test(s.status);
        const finSpecials = wordData.blocked.filter((s) => isSpecialStudent(s) && getSpecialGround(s).includes("financial"));
        const compSpecials = wordData.blocked.filter((s) => isSpecialStudent(s) && /compassionate|medical|sick/.test(getSpecialGround(s)));
        const otherSpecials = wordData.blocked.filter((s) => isSpecialStudent(s) && !getSpecialGround(s).includes("financial") && !/compassionate|medical|sick/.test(getSpecialGround(s)));
        await addDocIfNotEmpty(finSpecials, getFileName("Special_Exams_Financial"), promotionReport_1.generateSpecialExamsDoc, "Financial");
        await addDocIfNotEmpty(compSpecials, getFileName("Special_Exams_Compassionate"), promotionReport_1.generateSpecialExamsDoc, "Compassionate");
        await addDocIfNotEmpty(otherSpecials, getFileName("Special_Exams_Other"), promotionReport_1.generateSpecialExamsDoc, "Other");
        sendProgress(70, "Checking Stayout & Repeat Year...");
        const stayoutList = wordData.blocked.filter(s => s.status === "STAYOUT");
        await addDocIfNotEmpty(stayoutList, getFileName("Stayout_Retake_List"), promotionReport_1.generateStayoutExamsDoc);
        const repeatList = wordData.blocked.filter(s => s.status === "REPEAT YEAR");
        await addDocIfNotEmpty(repeatList, getFileName("Repeat_Year_List"), promotionReport_1.generateRepeatYearDoc);
        sendProgress(75, "Checking Academic Exceptions...");
        // 1. INCOMPLETE LIST
        const incompleteList = wordData.blocked.filter(s => s.status.includes("INC") && !s.status.includes("SPEC"));
        await addDocIfNotEmpty(incompleteList, getFileName("Incomplete_Results_List"), promotionReport_1.generateIncompleteListDoc);
        // 2. ACADEMIC LEAVE - Financial
        const isFinancialLeave = (s) => {
            const type = (s.academicLeavePeriod?.type || "").toLowerCase();
            const remarks = (s.remarks || "").toLowerCase();
            const isLeaveStatus = ["ACADEMIC LEAVE", "ON LEAVE"].includes(s.status);
            return isLeaveStatus && (type === "financial" || remarks.includes("financial"));
        };
        const finLeave = wordData.blocked.filter(isFinancialLeave);
        await addDocIfNotEmpty(finLeave, getFileName("Academic_Leave_Financial"), promotionReport_1.generateAcademicLeaveDoc, "Financial", "ACADEMIC LEAVE");
        // 3. ACADEMIC LEAVE - Compassionate  
        const isCompassionateLeave = (s) => {
            const type = (s.academicLeavePeriod?.type || "").toLowerCase();
            const remarks = (s.remarks || "").toLowerCase();
            const isLeaveStatus = ["ACADEMIC LEAVE", "ON LEAVE"].includes(s.status);
            return isLeaveStatus && (type === "compassionate" || remarks.includes("compassionate") || remarks.includes("medical"));
        };
        const compLeave = wordData.blocked.filter(isCompassionateLeave);
        await addDocIfNotEmpty(compLeave, getFileName("Academic_Leave_Compassionate"), promotionReport_1.generateAcademicLeaveDoc, "Compassionate", "ACADEMIC LEAVE");
        sendProgress(80, "Checking Discontinuations & Deregistrations...");
        const discoList = wordData.blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED");
        await addDocIfNotEmpty(discoList, getFileName("Discontinuation_List"), promotionReport_1.generateDiscontinuationDoc);
        const deregList = wordData.blocked.filter(s => s.status === "DEREGISTERED");
        await addDocIfNotEmpty(deregList, getFileName("Deregistration_List"), promotionReport_1.generateDeregistrationDoc);
        sendProgress(85, "Checking Deferment List...");
        const defermentList = wordData.blocked.filter((s) => s.status === "DEFERMENT");
        await addDocIfNotEmpty(defermentList, getFileName("Deferment_List"), promotionReport_1.generateDefermentDoc);
        sendProgress(90, "Checking Carry Forward List...");
        const carryList = wordData.eligible.filter(s => s.reasons?.length > 0 && s.status !== "ALREADY PROMOTED");
        await addDocIfNotEmpty(carryList, getFileName("Carry_Forward_List"), promotionReport_1.generateCarryForwardDoc);
        // 7. Zip and Send
        sendProgress(95, "Creating ZIP Archive...");
        const zipBase64 = zip.toBuffer().toString("base64");
        res.write(`data: ${JSON.stringify({ percent: 100, message: "Complete!", file: zipBase64 })}\n\n`);
        res.end();
    }
    catch (err) {
        console.error("Report Generation Error:", err);
        res.write(`data: ${JSON.stringify({ error: "Failed to generate" })}\n\n`);
        res.end();
    }
}));
router.post("/download-cms", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");
    const sendProgress = (percent, message, file) => {
        res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
    };
    console.log("[CMS DEBUG] Request body:", {
        programId,
        yearToPromote,
        academicYearName,
    });
    console.log("[CMS DEBUG] User institution:", req.user.institution?.toString());
    try {
        sendProgress(10, "Fetching student data...");
        const preview = await (0, statusEngine_1.previewPromotion)(programId, yearToPromote, academicYearName);
        const program = await Program_1.default.findById(programId).lean();
        // ── Academic year document (for session type AND mark filtering) ──────
        const academicYearDoc = await AcademicYear_1.default.findOne({
            year: academicYearName,
        }).lean();
        const targetAcadYearId = academicYearDoc?._id?.toString();
        const sessionExamType = academicYearDoc?.session === "SUPPLEMENTARY"
            ? "SUPPLEMENTARY"
            : "ORDINARY";
        const logoPath = path_1.default.join(__dirname, "../../public/institutionLogoExcel.png");
        const logoBuffer = fs_1.default.existsSync(logoPath)
            ? fs_1.default.readFileSync(logoPath)
            : Buffer.alloc(0);
        sendProgress(30, "Fetching marks for current cohort...");
        // ── KEY FIX 1: Exclude graduated/discontinued students and restrict to
        //    students whose Year N history entry matches THIS academic year.
        //    This prevents 2016-intake graduates from appearing in 2017/2018 CMS.
        const studentsByHistory = await Student_1.default.find({
            program: programId,
            status: {
                $nin: ["graduated", "graduand", "discontinued", "deregistered"],
            },
            $or: [
                // Currently enrolled in this year
                { currentYearOfStudy: yearToPromote },
                // Has a history record for this specific year AND cohort
                {
                    academicHistory: {
                        $elemMatch: {
                            yearOfStudy: yearToPromote,
                            academicYear: academicYearName,
                        },
                    },
                },
            ],
        }).lean();
        const previewIds = new Set([...preview.eligible, ...preview.blocked].map((s) => (s.id || s._id)?.toString()));
        const historyOnly = studentsByHistory.filter((s) => !previewIds.has(s._id.toString()));
        const allStudents = [
            ...preview.eligible,
            ...preview.blocked,
            ...historyOnly,
        ];
        const studentIds = allStudents
            .map((s) => (s._id || s.id)?.toString())
            .filter(Boolean);
        sendProgress(50, "Loading mark records...");
        const [detailedMarks, directMarks] = await Promise.all([
            Mark_1.default.find({ student: { $in: studentIds } })
                .populate({
                path: "programUnit",
                populate: { path: "unit", select: "code name" },
            })
                .lean(),
            MarkDirect_1.default.find({ student: { $in: studentIds } })
                .populate({
                path: "programUnit",
                populate: { path: "unit", select: "code name" },
            })
                .lean(),
        ]);
        const combinedMarks = [...detailedMarks, ...directMarks];
        const filteredMarks = combinedMarks.filter((m) => {
            const rightYear = m.programUnit &&
                Number(m.programUnit.requiredYear) === Number(yearToPromote);
            const rightCohort = !targetAcadYearId ||
                m.academicYear?.toString() === targetAcadYearId ||
                m.academicYear?._id?.toString() === targetAcadYearId;
            return rightYear && rightCohort;
        });
        const institutionSettings = await InstitutionSettings_1.default.findOne({
            institution: program?.institution,
        }).lean();
        const passMark = institutionSettings?.passMark ?? 40;
        const gradingScale = institutionSettings?.gradingScale ?? [];
        const offeredUnitsRaw = await ProgramUnit_1.default.find({
            program: programId,
            requiredYear: yearToPromote,
        })
            .populate("unit")
            .lean();
        const offeredUnits = offeredUnitsRaw.map((pu) => ({
            code: pu.unit?.code || "N/A",
            name: pu.unit?.name || "N/A",
        }));
        sendProgress(70, "Generating Consolidated Mark Sheet...");
        const excelData = {
            programName: program?.name || "Program",
            academicYear: academicYearName,
            yearOfStudy: yearToPromote,
            session: sessionExamType,
            students: allStudents,
            marks: filteredMarks,
            offeredUnits,
            logoBuffer,
            institutionId: program?.institution?.toString() || "",
            programId,
            passMark,
            gradingScale,
        };
        const xlsxBuffer = await (0, consolidatedMS_1.generateConsolidatedMarkSheet)(excelData);
        sendProgress(95, "Preparing download...");
        const base64 = xlsxBuffer.toString("base64");
        res.write(`data: ${JSON.stringify({ percent: 100, message: "Complete!", file: base64 })}\n\n`);
        res.end();
        // } catch (err: any) {
        //   console.error("CMS Generation Error:", err);
        //   res.write(`data: ${JSON.stringify({ error: "Failed to generate CMS" })}\n\n`);
        //   res.end();
        // }
    }
    catch (err) {
        const message = err instanceof Error ? err.message : "Unknown error";
        console.error("[CMS DEBUG] CATCH BLOCK:", message);
        if (err instanceof Error && err.stack) {
            console.error("[CMS DEBUG] STACK:", err.stack.split("\n").slice(0, 5).join("\n"));
        }
        res.write(`data: ${JSON.stringify({ error: message || "Failed to generate CMS" })}\n\n`);
        res.end();
    }
}));
// GET /promote/award-list?programId=xxx&academicYear=optional
// Returns JSON array of eligible graduates for the frontend preview.
router.get("/award-list", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, academicYear } = req.query;
    if (!programId)
        return res.status(400).json({ error: "programId is required" });
    const { generateAwardList } = await Promise.resolve().then(() => __importStar(require("../services/graduationEngine")));
    const list = await generateAwardList(programId, academicYear);
    res.json({ success: true, count: list.length, data: list });
}));
// GET /promote/award-list-doc?programId=xxx&academicYear=optional&variant=simple|classified
//   variant=simple     → plain list (S/N, Reg No., Name) — no WAA shown
//   variant=classified → grouped by class with WAA column (default)
router.get("/award-list-doc", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, academicYear, variant = "classified" } = req.query;
    if (!programId)
        return res.status(400).json({ error: "programId is required" });
    const { generateAwardList } = await Promise.resolve().then(() => __importStar(require("../services/graduationEngine")));
    const { generateAwardListDoc } = await Promise.resolve().then(() => __importStar(require("../utils/promotionReport")));
    const list = await generateAwardList(programId, academicYear);
    if (list.length === 0) {
        return res.status(404).json({ error: "No eligible graduates found." });
    }
    const program = await Program_1.default.findById(programId).lean();
    const logoPath = path_1.default.join(__dirname, "../../public/institutionLogoExcel.png");
    const logoBuffer = fs_1.default.existsSync(logoPath) ? fs_1.default.readFileSync(logoPath) : Buffer.alloc(0);
    const docData = {
        programName: program?.name || "Program",
        academicYear: academicYear || new Date().getFullYear().toString(),
        yearOfStudy: program?.durationYears || 5,
        logoBuffer,
        awardList: list,
        institutionId: req.user.institution.toString(), // ← ADD
    };
    const buffer = (variant === "simple") ? await (0, promotionReport_1.generateSimpleAwardListDoc)(docData) : await generateAwardListDoc(docData);
    const cleanYear = (academicYear || "ALL").replace(/\//g, "_");
    const progCode = program?.code || "PROG";
    const label = variant === "simple" ? "SIMPLE" : "CLASSIFIED";
    const fileName = `Award_List_${progCode}_${cleanYear}_${label}.docx`.replace(/\s+/g, "_");
    res
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        .header("Content-Disposition", `attachment; filename="${fileName}"`)
        .send(buffer);
}));
// POST /promote/download-journey-cms
// Generates the multi-year Student Journey CMS workbook for the Board.
router.post("/download-journey-cms", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { programId, academicYearName } = req.body;
    if (!programId)
        return res.status(400).json({ error: "programId is required" });
    const program = (await Program_1.default.findById(programId).lean());
    const academicYearDoc = academicYearName
        ? await AcademicYear_1.default.findOne({ year: academicYearName }).lean()
        : null;
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");
    const send = (percent, message, file) => res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
    try {
        send(10, "Loading student data...");
        const program = await Program_1.default.findById(programId).lean();
        const logoPath = path_1.default.join(__dirname, "../../public/institutionLogoExcel.png");
        const logoBuffer = fs_1.default.existsSync(logoPath) ? fs_1.default.readFileSync(logoPath) : Buffer.alloc(0);
        send(30, "Building academic histories...");
        const { generateJourneyCMS } = await Promise.resolve().then(() => __importStar(require("../utils/journeyCMS")));
        send(60, "Generating journey workbook...");
        const buffer = await generateJourneyCMS({
            // programId,
            // programName:  program?.name || "Program",
            // academicYear: academicYearName || new Date().getFullYear().toString(),
            // logoBuffer,
            // institutionId: program?.institution?.toString() || "",
            programId,
            programName: program.name,
            academicYear: academicYearDoc.year,
            logoBuffer: Buffer.alloc(0), // not used — loadLogoBuffer handles it internally
            institutionId: req.user.institution.toString(),
        });
        send(95, "Preparing download...");
        const base64 = buffer.toString("base64");
        res.write(`data: ${JSON.stringify({ percent: 100, message: "Complete!", file: base64 })}\n\n`);
        res.end();
    }
    catch (err) {
        console.error("[Journey CMS] Error:", err.message, err.stack);
        res.write(`data: ${JSON.stringify({ error: err.message || "Failed to generate Journey CMS" })}\n\n`);
        res.end();
    }
}));
// promote individual student
router.post("/:studentId", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { studentId } = req.params;
    const result = await (0, statusEngine_1.promoteStudent)(studentId);
    if (!result.success)
        return res.status(400).json({ error: "Promotion Denied", message: result.message, details: result.details });
    await (0, auditLogger_1.logAudit)(req, { action: "individual_student_promoted", targetUser: studentId, details: { message: result.message } });
    res.json(result);
}));
// POST /promote/undo/:studentId
router.post("/undo/:studentId", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { studentId } = req.params;
    const result = await (0, undoPromotion_1.undoPromotion)(studentId);
    if (!result.success) {
        return res.status(400).json({ success: false, message: result.message });
    }
    await (0, auditLogger_1.logAudit)(req, {
        action: "promotion_reversed",
        targetUser: studentId,
        details: { message: result.message, previousYear: result.previousYear, restoredYear: result.restoredYear },
    });
    res.json(result);
}));
exports.default = router;
