"use strict";
// serverside/src/services/gradeCalculator.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.computeFinalGrade = computeFinalGrade;
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const Student_1 = __importDefault(require("../models/Student"));
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const gradingCore_1 = require("../utils/gradingCore");
async function computeFinalGrade({ markId, session }) {
    // ── 1. Resolve mark document ──────────────────────────────────────────────
    let markDoc = await Mark_1.default
        .findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);
    let isDirect = false;
    if (!markDoc) {
        markDoc = await MarkDirect_1.default
            .findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);
        isDirect = true;
    }
    if (!markDoc)
        throw new Error("Mark not found");
    const settings = await InstitutionSettings_1.default.findOne({ institution: markDoc.institution }).session(session || null);
    if (!settings)
        throw new Error("Institution settings not found");
    // ── Extract values from nested ruleSet ───────────────────────────────────
    const ruleSet = settings.ruleSet || {};
    const catMax = ruleSet.catMax || 20;
    const assignmentMax = ruleSet.assignmentMax || 10;
    const practicalMax = ruleSet.practicalMax || 10;
    const passMark = ruleSet.passMark || 40;
    const hasLab = ruleSet.hasLab ?? true;
    const hasPractical = ruleSet.hasPractical ?? true;
    const hasWorkshop = ruleSet.hasWorkshop ?? false;
    // Determine unit type from mark document
    let unitType = "theory";
    if (markDoc.unitType) {
        unitType = markDoc.unitType;
    }
    else if (hasWorkshop) {
        unitType = "workshop";
    }
    else if (hasLab) {
        unitType = "lab";
    }
    // ── 2. Raw calculation ────────────────────────────────────────────────────
    const result = (0, gradingCore_1.calculateFinalResult)({
        cat1: markDoc.cat1Raw || 0, cat2: markDoc.cat2Raw || 0, cat3: markDoc.cat3Raw || 0,
        ass1: markDoc.assgnt1Raw || 0, ass2: markDoc.assgnt2Raw || 0, ass3: markDoc.assgnt3Raw || 0,
        practical: markDoc.practicalRaw || 0,
        examQ1: markDoc.examQ1Raw || 0, examQ2: markDoc.examQ2Raw || 0,
        examQ3: markDoc.examQ3Raw || 0, examQ4: markDoc.examQ4Raw || 0, examQ5: markDoc.examQ5Raw || 0,
        unitType: unitType,
        examMode: markDoc.examMode || "standard",
        attempt: markDoc.attempt || "1st",
        settings: {
            catMax: catMax,
            assMax: assignmentMax,
            practicalMax: practicalMax,
            passMark: passMark,
        },
    });
    // ── 3. Preservation strategy ──────────────────────────────────────────────
    const finalCA = result.caTotal === 0 && markDoc.caTotal30 > 0 ? markDoc.caTotal30 : result.caTotal;
    const finalExam = result.examTotal === 0 && markDoc.examTotal70 > 0 ? markDoc.examTotal70 : result.examTotal;
    const finalAgreed = finalCA + finalExam;
    // ── 4. Special exam override (ENG.18) ─────────────────────────────────────
    const isSpecial = markDoc.isSpecial === true || markDoc.attempt === "special";
    let grade;
    let status;
    let attemptType;
    // Get grading scale
    const gradingScale = settings.gradingScale || [
        { min: 70, max: 100, grade: "A", label: "Excellent" },
        { min: 60, max: 69, grade: "B", label: "Good" },
        { min: 50, max: 59, grade: "C", label: "Satisfactory" },
        { min: 40, max: 49, grade: "D", label: "Pass" },
        { min: 0, max: 39, grade: "E", label: "Fail" },
    ];
    const sortedScale = [...gradingScale].sort((a, b) => b.min - a.min);
    if (isSpecial && finalExam === 0) {
        // Special approved but exam NOT yet sat — keep as pending special
        grade = "I";
        status = "SPECIAL";
        attemptType = "SPECIAL";
    }
    else if (isSpecial && finalExam > 0) {
        // Special exam has been COMPLETED — grade it like a normal exam
        grade = sortedScale.find((s) => finalAgreed >= s.min)?.grade ?? "E";
        status = finalAgreed >= passMark ? "PASS" : "SUPPLEMENTARY";
        attemptType = "1ST_ATTEMPT";
    }
    else {
        // Normal (non-special) exam
        grade = sortedScale.find((s) => finalAgreed >= s.min)?.grade ?? "E";
        status = finalAgreed >= passMark ? "PASS" : "SUPPLEMENTARY";
        attemptType = markDoc.attempt === "supplementary" ? "SUPPLEMENTARY"
            : markDoc.attempt === "re-take" ? "RETAKE"
                : "1ST_ATTEMPT";
    }
    // ── 5. Update source mark ─────────────────────────────────────────────────
    const markUpdate = {
        $set: {
            caTotal30: finalCA,
            examTotal70: (isSpecial && finalExam === 0) ? 0 : finalExam,
            agreedMark: (isSpecial && finalExam === 0) ? finalCA : finalAgreed,
        },
    };
    if (isDirect)
        await MarkDirect_1.default.updateOne({ _id: markId }, markUpdate).session(session || null);
    else
        await Mark_1.default.updateOne({ _id: markId }, markUpdate).session(session || null);
    // ── 6. Upsert FinalGrade ──────────────────────────────────────────────────
    await FinalGrade_1.default.findOneAndUpdate({
        student: markDoc.student._id || markDoc.student,
        programUnit: markDoc.programUnit._id || markDoc.programUnit,
        academicYear: markDoc.academicYear._id || markDoc.academicYear,
    }, {
        $set: {
            totalMark: (isSpecial && finalExam === 0) ? finalCA : finalAgreed,
            grade,
            caTotal30: finalCA,
            examTotal70: (isSpecial && finalExam === 0) ? 0 : finalExam,
            status,
            isSpecial: isSpecial && finalExam === 0,
            attemptType,
            institution: markDoc.institution,
            semester: markDoc.semester || "SEMESTER 1",
        },
    }, { upsert: true, session });
    // ── 7. Carry-forward resolution (ENG.14) ──────────────────────────────────
    if (status === "PASS") {
        _resolveCFUnit((markDoc.student._id || markDoc.student).toString(), (markDoc.programUnit._id || markDoc.programUnit).toString()).catch((err) => console.error("[gradeCalculator] CF resolution:", err.message));
    }
    return {
        caTotal: finalCA,
        examTotal: (isSpecial && finalExam === 0) ? 0 : finalExam,
        finalMark: (isSpecial && finalExam === 0) ? finalCA : finalAgreed,
        grade,
        status,
    };
}
async function _resolveCFUnit(studentId, programUnitId) {
    const studentDoc = await Student_1.default.findById(studentId)
        .select("carryForwardUnits qualifierSuffix").lean();
    if (!studentDoc)
        return;
    const hasCF = (studentDoc.carryForwardUnits || []).some((u) => u.programUnitId === programUnitId);
    if (!hasCF)
        return;
    await Student_1.default.findByIdAndUpdate(studentId, { $pull: { carryForwardUnits: { programUnitId } } });
    const refreshed = await Student_1.default.findById(studentId).select("carryForwardUnits").lean();
    if ((refreshed?.carryForwardUnits || []).length === 0) {
        await Student_1.default.findByIdAndUpdate(studentId, { $set: { qualifierSuffix: "" } });
    }
}
