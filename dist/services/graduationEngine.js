"use strict";
// serverside/src/services/graduationEngine.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateAwardList = exports.calculateGraduationStatus = void 0;
exports.resolveAgreedMark = resolveAgreedMark;
const Student_1 = __importDefault(require("../models/Student"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const Mark_1 = __importDefault(require("../models/Mark"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const weightingRegistry_1 = require("../utils/weightingRegistry");
// ─────────────────────────────────────────────────────────────────────────────
// Resolve the best agreed mark for a student + programUnit across all sources.
// Priority: FinalGrade → MarkDirect → Mark
// This is the SINGLE source-of-truth resolver used by both WAA computation
// and the Journey CMS mark display.
// ─────────────────────────────────────────────────────────────────────────────
async function resolveAgreedMark(studentId, puId) {
    // 1. FinalGrade (set by computeFinalGrade — most authoritative)
    const fgs = await FinalGrade_1.default.find({
        student: studentId,
        programUnit: puId,
    }).lean();
    if (fgs.length > 0) {
        const rank = (s) => s === "PASS" ? 3 : s === "SPECIAL" ? 2 : s === "SUPPLEMENTARY" ? 1 : 0;
        const best = fgs.sort((a, b) => rank(b.status) - rank(a.status))[0];
        const mark = best.totalMark ?? 0;
        if (mark > 0)
            return mark;
    }
    // 2. MarkDirect (direct CA+Exam entry — no FinalGrade until directMarksImporter_FIXED runs)
    const md = await MarkDirect_1.default.findOne({ student: studentId, programUnit: puId }).lean();
    if (md && (md.agreedMark ?? 0) > 0)
        return md.agreedMark;
    // 3. Mark (detailed breakdown with computeFinalGrade applied)
    const dm = await Mark_1.default.findOne({ student: studentId, programUnit: puId }).lean();
    if (dm && (dm.agreedMark ?? 0) > 0)
        return dm.agreedMark;
    return 0; // no mark found
}
// ─────────────────────────────────────────────────────────────────────────────
// Core WAA computer
//
// Priority:
//   1. student.finalWeightedAverage  — stored at graduation (most accurate)
//   2. sum of academicHistory[].weightedContribution — stored at promotion
//   3. Recompute from actual marks (FinalGrade + MarkDirect + Mark fallback)
//      and backfill the DB so future calls use Priority 1.
// ─────────────────────────────────────────────────────────────────────────────
async function computeWAA(student) {
    const program = student.program;
    const history = (student.academicHistory || []);
    const entryType = student.entryType || "Direct";
    // ── Priority 1: finalWeightedAverage ─────────────────────────────────────
    if (student.finalWeightedAverage != null) {
        const stored = parseFloat(student.finalWeightedAverage);
        if (stored > 0) {
            const yearBreakdown = history.map((h) => ({
                yearOfStudy: h.yearOfStudy,
                academicYear: h.academicYear || "",
                annualMean: h.annualMeanMark || 0,
                weight: (0, weightingRegistry_1.getYearWeight)(program, entryType, h.yearOfStudy),
                weightedContribution: h.weightedContribution || 0,
                isRepeat: h.isRepeatYear || false,
                failedUnits: h.failedUnitsCount || 0,
            }));
            return { waa: stored, yearBreakdown };
        }
    }
    // ── Priority 2: stored weightedContribution ───────────────────────────────
    const totalStoredWC = history.reduce((sum, h) => sum + (h.weightedContribution || 0), 0);
    if (totalStoredWC > 0) {
        const yearBreakdown = history.map((h) => ({
            yearOfStudy: h.yearOfStudy,
            academicYear: h.academicYear || "",
            annualMean: h.annualMeanMark || 0,
            weight: (0, weightingRegistry_1.getYearWeight)(program, entryType, h.yearOfStudy),
            weightedContribution: h.weightedContribution || 0,
            isRepeat: h.isRepeatYear || false,
            failedUnits: h.failedUnitsCount || 0,
        }));
        return { waa: totalStoredWC, yearBreakdown };
    }
    // ── Priority 3: Recompute from actual marks ───────────────────────────────
    // Reads FinalGrade first, then MarkDirect, then Mark as fallback.
    // Handles legacy students and students whose direct marks have no FinalGrade yet.
    // console.log(`[graduationEngine] Recomputing WAA from marks for ${student.regNo}`);
    const duration = program?.durationYears || 5;
    let waa = 0;
    const yearBreakdown = [];
    const historyUpdates = [];
    for (let yearOfStudy = 1; yearOfStudy <= duration; yearOfStudy++) {
        const weight = (0, weightingRegistry_1.getYearWeight)(program, entryType, yearOfStudy);
        const programUnits = await ProgramUnit_1.default.find({
            program: student.program._id || student.program,
            requiredYear: yearOfStudy,
        }).lean();
        const unitCount = programUnits.length || 1;
        let totalMark = 0;
        let resolved = 0; // how many units have a mark
        for (const pu of programUnits) {
            const puId = pu._id.toString();
            const mark = await resolveAgreedMark(student._id.toString(), puId);
            if (mark > 0) {
                totalMark += mark;
                resolved++;
            }
        }
        // If no marks at all for this year, skip (student may not have sat exams)
        const annualMean = unitCount > 0 ? totalMark / unitCount : 0;
        const wc = annualMean * weight;
        waa += wc;
        yearBreakdown.push({
            yearOfStudy,
            academicYear: history.find((h) => h.yearOfStudy === yearOfStudy)?.academicYear || "",
            annualMean: parseFloat(annualMean.toFixed(2)),
            weight,
            weightedContribution: parseFloat(wc.toFixed(4)),
            isRepeat: history.find((h) => h.yearOfStudy === yearOfStudy)?.isRepeatYear || false,
            failedUnits: history.find((h) => h.yearOfStudy === yearOfStudy)?.failedUnitsCount || 0,
        });
        historyUpdates.push({ yearOfStudy, annualMean, wc });
    }
    // Backfill DB (fire-and-forget)
    _backfillStudentWAA(student._id, waa, historyUpdates).catch((err) => console.warn(`[graduationEngine] Backfill failed for ${student.regNo}:`, err.message));
    return { waa: parseFloat(waa.toFixed(2)), yearBreakdown };
}
// ─── DB backfill ──────────────────────────────────────────────────────────────
async function _backfillStudentWAA(studentId, waa, yearUpdates) {
    const student = await Student_1.default.findById(studentId).lean();
    if (!student)
        return;
    const updatedHistory = (student.academicHistory || []).map((h) => {
        const upd = yearUpdates.find(u => u.yearOfStudy === h.yearOfStudy);
        if (!upd)
            return h;
        return { ...h, annualMeanMark: upd.annualMean, weightedContribution: upd.wc };
    });
    let classification = "PASS";
    if (waa >= 70)
        classification = "FIRST CLASS HONOURS";
    else if (waa >= 60)
        classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
    else if (waa >= 50)
        classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
    await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            finalWeightedAverage: waa.toFixed(2),
            classification,
            academicHistory: updatedHistory,
        },
    });
    // console.log(`[graduationEngine] Backfilled WAA=${waa.toFixed(2)}, class=${classification} for ${studentId}`);
}
// ─── Single student graduation status ─────────────────────────────────────────
const calculateGraduationStatus = async (studentId) => {
    const student = await Student_1.default.findById(studentId).populate("program").lean();
    if (!student)
        throw new Error(`Student ${studentId} not found`);
    const program = student.program;
    const duration = program?.durationYears || 5;
    const history = (student.academicHistory || []);
    const missingRequirements = [];
    // Eligibility checks
    const yearsPresent = new Set(history.map((h) => h.yearOfStudy));
    for (let y = 1; y <= duration; y++) {
        if (!yearsPresent.has(y))
            missingRequirements.push(`Year ${y} has no academic history record`);
    }
    history.filter((h) => (h.failedUnitsCount || 0) > 0).forEach((h) => {
        missingRequirements.push(`Year ${h.yearOfStudy} has ${h.failedUnitsCount} uncleared failed unit(s)`);
    });
    const pendingCF = (student.carryForwardUnits || []).filter((u) => u.status === "pending" || !u.status);
    if (pendingCF.length > 0) {
        missingRequirements.push(`${pendingCF.length} carry-forward unit(s) pending: ${pendingCF.map((u) => u.unitCode).join(", ")}`);
    }
    if (!["graduand", "graduated"].includes(student.status) && missingRequirements.length === 0) {
        missingRequirements.push(`Student status is "${student.status}" — must be "graduand" or "graduated"`);
    }
    const { waa, yearBreakdown } = await computeWAA(student);
    let classification = "FAIL";
    if (waa >= 70)
        classification = "FIRST CLASS HONOURS";
    else if (waa >= 60)
        classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
    else if (waa >= 50)
        classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
    else if (waa >= 40)
        classification = "PASS";
    return {
        studentId: student._id.toString(),
        regNo: student.regNo,
        name: student.name,
        weightedAggregateAverage: parseFloat(waa.toFixed(2)),
        classification,
        isEligible: missingRequirements.length === 0,
        missingRequirements,
        yearBreakdown,
    };
};
exports.calculateGraduationStatus = calculateGraduationStatus;
// ─── Award list ───────────────────────────────────────────────────────────────
const generateAwardList = async (programId, academicYear) => {
    const students = await Student_1.default.find({
        program: programId,
        status: { $in: ["graduand", "graduated"] },
    }).populate("program").lean();
    const awardList = [];
    for (const student of students) {
        if (academicYear) {
            const lastHistory = (student.academicHistory || []).at(-1);
            const lastYear = lastHistory?.academicYear || "";
            if (!lastYear.includes(academicYear) && student.graduationYear !== parseInt(academicYear))
                continue;
        }
        const { waa } = await computeWAA(student);
        let classification = student.classification || "";
        if (!classification || (classification === "PASS" && waa > 50)) {
            if (waa >= 70)
                classification = "FIRST CLASS HONOURS";
            else if (waa >= 60)
                classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
            else if (waa >= 50)
                classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
            else if (waa >= 40)
                classification = "PASS";
            else
                classification = "FAIL";
        }
        awardList.push({
            studentId: student._id.toString(),
            regNo: student.regNo,
            name: student.name,
            waa: parseFloat(waa.toFixed(2)),
            classification,
            graduationYear: student.graduationYear || new Date().getFullYear(),
        });
    }
    const classOrder = [
        "FIRST CLASS HONOURS",
        "SECOND CLASS HONOURS (UPPER DIVISION)",
        "SECOND CLASS HONOURS (LOWER DIVISION)",
        "PASS", "FAIL",
    ];
    awardList.sort((a, b) => {
        const ai = classOrder.indexOf(a.classification);
        const bi = classOrder.indexOf(b.classification);
        if (ai !== bi)
            return ai - bi;
        return b.waa - a.waa;
    });
    return awardList;
};
exports.generateAwardList = generateAwardList;
