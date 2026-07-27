"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.clearDeferredSuppUnit = exports.getDeferredSuppStudentsForUnit = exports.deferSuppToNextOrdinary = exports.getStayoutStudentsForUnit = exports.getCarryForwardStudentsForUnit = exports.clearCarryForwardUnit = exports.assessAndGrantCarryForward = void 0;
// serverside/src/services/carryForwardService.ts
const mongoose_1 = __importDefault(require("mongoose"));
const Student_1 = __importDefault(require("../models/Student"));
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const academicRules_1 = require("../utils/academicRules");
// ─── assessAndGrantCarryForward ───────────────────────────────────────────────
// Called from promoteStudent after supplementary results are finalized.
// Determines carry-forward eligibility per ENG.14 and persists to student record.
const assessAndGrantCarryForward = async (studentId, programId, yearOfStudy, academicYearName) => {
    const student = await Student_1.default.findById(studentId).lean();
    if (!student)
        throw new Error("Student not found");
    const settings = await InstitutionSettings_1.default.findOne({ institution: student.institution }).lean();
    const passMark = settings?.passMark ?? 40;
    // ENG.14a: No carry-forward to final year
    const programDoc = await mongoose_1.default.model("Program").findById(programId).lean();
    const finalYear = programDoc?.durationYears || 5;
    if (yearOfStudy >= finalYear) {
        return { granted: false, cfUnits: [], qualifier: "", reason: `ENG.14: No carry-forward to final year (Year ${finalYear}).` };
    }
    const programUnits = await ProgramUnit_1.default.find({ program: programId, requiredYear: yearOfStudy })
        .populate("unit").lean();
    const totalUnits = programUnits.length;
    const puIds = programUnits.map((pu) => pu._id);
    const [detailedMarks, directMarks] = await Promise.all([
        Mark_1.default.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
        MarkDirect_1.default.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
    ]);
    const markMap = new Map();
    [...detailedMarks, ...directMarks].forEach((m) => markMap.set(m.programUnit.toString(), m));
    const failedUnitCodes = [];
    const noCAUnitCodes = [];
    const failedDetails = [];
    for (const pu of programUnits) {
        const puId = pu._id.toString();
        const m = markMap.get(puId);
        if (!m)
            continue;
        if (m.isSpecial || m.attempt === "special")
            continue;
        const mark = m.agreedMark ?? 0;
        const hasCA = (m.caTotal30 ?? 0) > 0;
        if (mark < passMark) {
            const code = pu.unit?.code || "N/A";
            failedUnitCodes.push(code);
            if (!hasCA)
                noCAUnitCodes.push(code); // ENG.15a: missing CA → cannot CF
            failedDetails.push({ programUnitId: puId, unitCode: code, unitName: pu.unit?.name || "N/A" });
        }
    }
    const eligibility = (0, academicRules_1.assessCarryForwardEligibility)(failedUnitCodes, noCAUnitCodes, totalUnits);
    if (!eligibility.eligible)
        return { granted: false, cfUnits: [], qualifier: "", reason: eligibility.reason };
    // Determine CF cycle number from existing qualifierSuffix
    const priorQualifier = student.qualifierSuffix || "";
    const priorMatch = priorQualifier.match(/RP(\d+)C/);
    const cfNumber = priorMatch ? Math.min(parseInt(priorMatch[1]) + 1, 3) : 1;
    const qualifier = academicRules_1.REG_QUALIFIERS.carryForward(cfNumber);
    const cfUnits = eligibility.units.map((code) => {
        const detail = failedDetails.find((d) => d.unitCode === code);
        return {
            programUnitId: detail?.programUnitId || "",
            unitCode: code,
            unitName: detail?.unitName || "N/A",
            fromYear: yearOfStudy,
            fromAcademicYear: academicYearName,
            attemptNumber: cfNumber + 2,
            qualifier,
            addedAt: new Date(),
            status: "pending",
        };
    });
    await Student_1.default.findByIdAndUpdate(studentId, {
        $push: { carryForwardUnits: { $each: cfUnits } },
        $set: { qualifierSuffix: qualifier },
    });
    return {
        granted: true,
        cfUnits,
        qualifier,
        reason: `ENG.14: Carry forward granted — ${cfUnits.length} unit(s): ${cfUnits.map((u) => u.unitCode).join(", ")}`,
    };
};
exports.assessAndGrantCarryForward = assessAndGrantCarryForward;
// ─── clearCarryForwardUnit ────────────────────────────────────────────────────
// Called from gradeCalculator when a CF unit is graded PASS.
const clearCarryForwardUnit = async (studentId, programUnitId) => {
    await Student_1.default.findByIdAndUpdate(studentId, {
        $pull: { carryForwardUnits: { programUnitId } },
    });
    const updated = await Student_1.default.findById(studentId).select("carryForwardUnits").lean();
    const remaining = (updated?.carryForwardUnits || []).length;
    if (remaining === 0) {
        await Student_1.default.findByIdAndUpdate(studentId, { $set: { qualifierSuffix: "" } });
    }
};
exports.clearCarryForwardUnit = clearCarryForwardUnit;
// ─── getCarryForwardStudentsForUnit ──────────────────────────────────────────
// Used by scoresheetStudentList to include CF students on ORDINARY scoresheets.
const getCarryForwardStudentsForUnit = async (programUnitId, programId) => {
    const students = await Student_1.default.find({
        program: programId,
        "carryForwardUnits.programUnitId": programUnitId,
        "carryForwardUnits.status": "pending",
    }).lean();
    return students
        .map((student) => {
        const cfUnit = student.carryForwardUnits.find((u) => u.programUnitId === programUnitId && u.status === "pending");
        if (!cfUnit)
            return null;
        return { student, cfUnit, attemptLabel: cfUnit.qualifier || "RP1C" };
    })
        .filter((item) => item !== null);
};
exports.getCarryForwardStudentsForUnit = getCarryForwardStudentsForUnit;
// ─── getStayoutStudentsForUnit ────────────────────────────────────────────────
// ENG.15h: Stayout students retake in ORDINARY of NEXT year.
const getStayoutStudentsForUnit = async (programUnitId, programId) => {
    const pu = await ProgramUnit_1.default.findById(programUnitId).lean();
    if (!pu)
        return [];
    const expectedYear = (pu.requiredYear || 1) + 1;
    const failedGrades = await FinalGrade_1.default.find({
        programUnit: programUnitId,
        status: { $ne: "PASS" },
        attemptType: { $in: ["1ST_ATTEMPT", "SUPPLEMENTARY"] },
    }).populate("student").lean();
    const result = [];
    for (const grade of failedGrades) {
        const student = grade.student;
        if (!student)
            continue;
        if (student.program?.toString() !== programId)
            continue;
        if (student.currentYearOfStudy !== expectedYear)
            continue;
        if (student.status !== "active")
            continue;
        if ((student.qualifierSuffix || "").includes("C"))
            continue; // CF students handled separately
        result.push({ student, attemptLabel: "A/SO" });
    }
    return result;
};
exports.getStayoutStudentsForUnit = getStayoutStudentsForUnit;
// ─────────────────────────────────────────────────────────────────────────────
// deferSuppToNextOrdinary
// Called by the coordinator route POST /student/defer-supp.
// Marks specific failed/special units as deferred to the next ordinary period.
// Allows the student to be promoted despite pending units (ENG.13b / ENG.18c).
// ─────────────────────────────────────────────────────────────────────────────
const deferSuppToNextOrdinary = async (studentId, programUnitIds, academicYear, reason) => {
    const student = await Student_1.default.findById(studentId).lean();
    if (!student)
        throw new Error("Student not found");
    const programUnits = await ProgramUnit_1.default.find({
        _id: { $in: programUnitIds },
    }).populate("unit").lean();
    const entries = programUnits.map((pu) => ({
        programUnitId: pu._id.toString(),
        unitCode: pu.unit?.code || "N/A",
        unitName: pu.unit?.name || "N/A",
        fromYear: student.currentYearOfStudy,
        fromAcademicYear: academicYear,
        reason,
        addedAt: new Date(),
        status: "pending",
    }));
    // Remove any existing pending deferred entries for these units first
    await Student_1.default.findByIdAndUpdate(studentId, {
        $pull: { deferredSuppUnits: { programUnitId: { $in: programUnitIds } } },
    });
    // Add the new deferred entries
    await Student_1.default.findByIdAndUpdate(studentId, {
        $push: {
            deferredSuppUnits: { $each: entries },
            statusEvents: {
                fromStatus: student.status,
                toStatus: student.status,
                date: new Date(),
                academicYear,
                reason: `ENG.13b/ENG.18c: Deferred ${reason.replace("_", " ")} for units: ${entries.map(e => e.unitCode).join(", ")}`,
            },
        },
    });
};
exports.deferSuppToNextOrdinary = deferSuppToNextOrdinary;
// ─────────────────────────────────────────────────────────────────────────────
// getDeferredSuppStudentsForUnit
// Returns students who deferred their supp/special for this specific unit
// to the next ordinary period. Called by scoresheetStudentList.
// ─────────────────────────────────────────────────────────────────────────────
const getDeferredSuppStudentsForUnit = async (programUnitId, programId) => {
    const students = await Student_1.default.find({
        program: programId,
        "deferredSuppUnits.programUnitId": programUnitId,
        "deferredSuppUnits.status": "pending",
    }).lean();
    return students
        .map((student) => {
        const entry = (student.deferredSuppUnits || []).find((u) => u.programUnitId === programUnitId && u.status === "pending");
        if (!entry)
            return null;
        const isSpecial = entry.reason === "special_deferred";
        return {
            student,
            attemptLabel: isSpecial ? "Special" : "Supp",
            isSupp: !isSpecial,
            isSpecial,
        };
    })
        .filter((x) => x !== null);
};
exports.getDeferredSuppStudentsForUnit = getDeferredSuppStudentsForUnit;
// ─────────────────────────────────────────────────────────────────────────────
// clearDeferredSuppUnit
// Called from gradeCalculator when a deferred unit is graded PASS.
// ─────────────────────────────────────────────────────────────────────────────
const clearDeferredSuppUnit = async (studentId, programUnitId) => {
    await Student_1.default.findByIdAndUpdate(studentId, { $set: { "deferredSuppUnits.$[elem].status": "passed" } }, { arrayFilters: [{ "elem.programUnitId": programUnitId, "elem.status": "pending" }] });
};
exports.clearDeferredSuppUnit = clearDeferredSuppUnit;
