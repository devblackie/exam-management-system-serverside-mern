"use strict";
// serverside/src/services/statusEngine.ts
// KEY CHANGES vs prior version:
//   1. StudentStatusResult gains `deferredList` field
//   2. calculateStudentStatus suppresses deferred units from fail/special counts
//      and populates deferredList — so the status box reflects post-defer reality
//   3. academicYearName is always returned (fixes the 400 on defer submit)
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
exports.bulkPromoteClass = exports.previewPromotion = exports.promoteStudent = exports.calculateStudentStatus = void 0;
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const Student_1 = __importDefault(require("../models/Student"));
const Mark_1 = __importDefault(require("../models/Mark"));
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const weightingRegistry_1 = require("../utils/weightingRegistry");
const academicAudit_1 = require("./academicAudit");
const academicRules_1 = require("../utils/academicRules");
const historicalYearValidator_1 = require("./historicalYearValidator");
function _qualifierShouldClearOnPromotion(qualifier) {
    if (!qualifier || qualifier.trim() === "")
        return false;
    return /^RP\d+$/.test(qualifier.trim());
}
let _assessAndGrant = null;
async function tryAssessAndGrantCarryForward(studentId, programId, yearOfStudy, academicYearName) {
    try {
        if (!_assessAndGrant) {
            const mod = await Promise.resolve().then(() => __importStar(require("./carryForwardService")));
            _assessAndGrant = mod.assessAndGrantCarryForward;
        }
        return await _assessAndGrant(studentId, programId, yearOfStudy, academicYearName);
    }
    catch (err) {
        console.warn("[StatusEngine] carryForwardService unavailable:", err.message);
        return { granted: false, cfUnits: [], qualifier: "", reason: err.message };
    }
}
const syncTerminalStatusToDb = async (studentId, engineStatus, details, academicYear) => {
    const terminalMap = {
        "DEREGISTERED": { dbStatus: "deregistered", qualifierFn: () => "" },
        "REPEAT YEAR": { dbStatus: "repeat", qualifierFn: (s) => {
                const count = (s.academicHistory || []).filter((h) => h.isRepeatYear).length + 1;
                return academicRules_1.REG_QUALIFIERS.repeatYear(count);
            } },
        "STAYOUT": { dbStatus: "active", qualifierFn: () => "" },
        "DISCONTINUED": { dbStatus: "discontinued", qualifierFn: () => "" },
    };
    const entry = terminalMap[engineStatus];
    if (!entry)
        return;
    const student = await Student_1.default.findById(studentId).lean();
    if (!student)
        return;
    if (entry.dbStatus !== "active" && student.status === entry.dbStatus)
        return;
    const fromStatus = student.status;
    const newQualifier = entry.qualifierFn(student);
    const updatePayload = {
        $set: { status: entry.dbStatus, remarks: details },
        $push: {
            statusEvents: { fromStatus, toStatus: entry.dbStatus, date: new Date(), reason: `Auto-Sync: ${details}`, academicYear },
            statusHistory: { status: entry.dbStatus, previousStatus: fromStatus, date: new Date(), reason: details },
        },
    };
    if (newQualifier)
        updatePayload.$set.qualifierSuffix = newQualifier;
    await Student_1.default.findByIdAndUpdate(studentId, updatePayload);
};
// ── calculateStudentStatus ────────────────────────────────────────────────────
const calculateStudentStatus = async (studentId, programId, academicYearName, yearOfStudy = 1, options = {}) => {
    const settings = await InstitutionSettings_1.default.findOne().lean();
    if (!settings)
        throw new Error("Institution settings not found.");
    const passMark = settings.passMark || 40;
    const student = await Student_1.default.findById(studentId).lean();
    if (!student)
        throw new Error("Student not found");
    // ── Terminal gate ─────────────────────────────────────────────────────────
    const TERMINAL = {
        on_leave: { label: "ACADEMIC LEAVE", variant: "info" },
        deferred: { label: "DEFERMENT", variant: "info" },
        discontinued: { label: "DISCONTINUED", variant: "error" },
        deregistered: { label: "DEREGISTERED", variant: "error" },
        graduated: { label: "GRADUATED", variant: "success" },
        graduand: { label: "GRADUATED", variant: "success" },
    };
    const terminalEntry = TERMINAL[student.status ?? ""];
    if (terminalEntry) {
        const leaveType = student.academicLeavePeriod?.type || "";
        const rem = (student.remarks || "").toLowerCase();
        let grounds = "";
        if (leaveType === "financial" || rem.includes("financial"))
            grounds = "FINANCIAL";
        else if (leaveType === "compassionate" || rem.includes("compassionate") || rem.includes("medical"))
            grounds = "COMPASSIONATE";
        else if (leaveType)
            grounds = leaveType.toUpperCase();
        const count = await ProgramUnit_1.default.countDocuments({ program: programId, requiredYear: yearOfStudy });
        return {
            status: terminalEntry.label, variant: terminalEntry.variant,
            details: `Student is currently ${terminalEntry.label}.${grounds ? ` Grounds: ${grounds}.` : ""}`,
            weightedMean: "0.00", sessionState: "ORDINARY", academicYearName: academicYearName ?? "",
            summary: { totalExpected: count, passed: 0, failed: 0, missing: 0, isOnLeave: true },
            passedList: [], failedList: [], specialList: [], deferredList: [],
            missingList: [], incompleteList: [], leaveDetails: grounds,
        };
    }
    // ── Academic year ──────────────────────────────────────────────────────────
    let targetYearDoc = null;
    if (!academicYearName || academicYearName === "CURRENT" || academicYearName === "undefined") {
        targetYearDoc = (await AcademicYear_1.default.findOne({ isCurrent: true }).lean())
            || (await AcademicYear_1.default.findOne().sort({ startDate: -1 }).lean());
    }
    else {
        targetYearDoc = (await AcademicYear_1.default.findOne({ year: academicYearName }).lean())
            || (await AcademicYear_1.default.findOne({ year: { $regex: new RegExp(`^${academicYearName.replace("/", "\\/")}$`, "i") } }).lean());
        if (!targetYearDoc)
            console.warn(`[StatusEngine] AcademicYear "${academicYearName}" not found.`);
    }
    const resolvedYearName = targetYearDoc?.year ?? academicYearName ?? "";
    // ── Curriculum ────────────────────────────────────────────────────────────
    const curriculum = await ProgramUnit_1.default.find({ program: programId, requiredYear: yearOfStudy })
        .populate("unit").lean();
    if (!curriculum?.length) {
        return {
            status: "CURRICULUM NOT SET", variant: "info", details: `No units defined for Year ${yearOfStudy}.`,
            weightedMean: "0.00", sessionState: "ORDINARY", academicYearName: resolvedYearName,
            summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
            passedList: [], failedList: [], specialList: [], deferredList: [],
            missingList: [], incompleteList: [],
        };
    }
    const programUnitIds = curriculum.map((pu) => pu._id);
    const [detailedMarks, directMarks, finalGrades] = await Promise.all([
        Mark_1.default.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
        MarkDirect_1.default.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
        FinalGrade_1.default.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
    ]);
    // ── marksMap ──────────────────────────────────────────────────────────────
    const marksMap = new Map();
    finalGrades.forEach((fg) => {
        const key = fg.programUnit?.toString();
        if (!key)
            return;
        const existing = marksMap.get(key);
        if (existing?.source === "finalGrade") {
            const existingIsBetter = (existing._fgStatus === "PASS" && fg.status !== "PASS") ||
                (existing._fgStatus === fg.status && (existing.createdAt ?? 0) >= (fg.createdAt ?? 0));
            if (existingIsBetter)
                return;
        }
        marksMap.set(key, {
            agreedMark: fg.totalMark ?? 0,
            caTotal30: fg.caTotal30 != null ? fg.caTotal30 : (fg.totalMark > 0 ? 1 : 0),
            examTotal70: fg.examTotal70 != null ? fg.examTotal70 : (fg.totalMark > 0 ? 1 : 0),
            attempt: fg.attemptType === "SUPPLEMENTARY" ? "supplementary"
                : fg.attemptType === "RETAKE" ? "re-take" : "1st",
            isSpecial: (fg.isSpecial === true || fg.status === "SPECIAL") && fg.status !== "PASS",
            source: "finalGrade",
            _fgStatus: fg.status,
        });
    });
    directMarks.forEach((m) => {
        marksMap.set(m.programUnit.toString(), { ...m, source: "direct" });
    });
    detailedMarks.forEach((m) => {
        marksMap.set(m.programUnit.toString(), { ...m, source: "detailed" });
    });
    finalGrades.forEach((fg) => {
        if (fg.status !== "PASS")
            return;
        const key = fg.programUnit?.toString();
        if (!key)
            return;
        const current = marksMap.get(key);
        if (!current)
            return;
        if (current.attempt === "special" || current.isSpecial === true) {
            marksMap.set(key, {
                agreedMark: fg.totalMark ?? 0,
                caTotal30: fg.caTotal30 != null ? fg.caTotal30 : (fg.totalMark > 0 ? 1 : 0),
                examTotal70: fg.examTotal70 != null ? fg.examTotal70 : (fg.totalMark > 0 ? 1 : 0),
                attempt: "1st", isSpecial: false,
                source: "finalGrade_pass", _fgStatus: "PASS",
            });
        }
    });
    // ── Unit classification loop ──────────────────────────────────────────────
    const lists = {
        passed: [],
        failed: [],
        special: [],
        missing: [],
        incomplete: [],
    };
    let totalFirstAttemptSum = 0;
    curriculum.forEach((pUnit) => {
        const code = pUnit.unit?.code?.toUpperCase();
        const displayName = `${code}: ${pUnit.unit?.name}`;
        const programUnitId = pUnit._id.toString();
        const rawMark = marksMap.get(programUnitId);
        if (!rawMark) {
            lists.missing.push(displayName);
            return;
        }
        const hasCAT = (rawMark.caTotal30 || 0) > 0;
        const hasExam = (rawMark.examTotal70 || 0) > 0;
        const markVal = rawMark.agreedMark || 0;
        const isSupp = rawMark.attempt === "supplementary";
        const isSpc = rawMark.attempt === "special" || rawMark.isSpecial === true;
        const notation = (0, academicRules_1.getAttemptLabel)({
            markAttempt: rawMark.attempt,
            studentStatus: student.status,
            studentQualifier: student.qualifierSuffix,
        });
        if (isSpc) {
            lists.special.push({ displayName, grounds: rawMark.remarks || "Special", programUnitId });
        }
        else if (!hasCAT && !hasExam) {
            lists.missing.push(`${displayName} (Absent)`);
        }
        else if (!hasCAT && hasExam) {
            if (isSupp) {
                if (markVal >= passMark)
                    lists.passed.push({ code, mark: markVal });
                else
                    lists.failed.push({ displayName, attempt: notation, programUnitId });
                totalFirstAttemptSum += markVal;
            }
            else {
                lists.incomplete.push(`${displayName} (No CAT)`);
            }
        }
        else if (!hasExam && hasCAT) {
            lists.missing.push(`${displayName} (Missing Exam)`);
        }
        else {
            if (markVal >= passMark)
                lists.passed.push({ code, mark: markVal });
            else
                lists.failed.push({ displayName, attempt: notation, programUnitId });
            totalFirstAttemptSum += markVal;
        }
    });
    // ── Deferred-unit suppression (ENG.13b / ENG.18c) ────────────────────────
    // Units the coordinator has deferred to the next ordinary period are removed
    // from the fail/special lists so the status reflects the student's effective
    // standing (i.e. what they still need to resolve THIS year).
    const pendingDeferredUnits = [];
    const allPendingDeferred = (student.deferredSuppUnits || []).filter((u) => u.status === "pending");
    if (allPendingDeferred.length > 0) {
        const deferredIds = new Set(allPendingDeferred.map((u) => u.programUnitId));
        // Move matching entries out of failed/special into deferredList
        const newFailed = [];
        const newSpecial = [];
        for (const f of lists.failed) {
            if (deferredIds.has(f.programUnitId)) {
                const entry = allPendingDeferred.find((u) => u.programUnitId === f.programUnitId);
                pendingDeferredUnits.push({ displayName: f.displayName, programUnitId: f.programUnitId, reason: entry?.reason || "supp_deferred" });
            }
            else {
                newFailed.push(f);
            }
        }
        for (const s of lists.special) {
            if (deferredIds.has(s.programUnitId)) {
                const entry = allPendingDeferred.find((u) => u.programUnitId === s.programUnitId);
                pendingDeferredUnits.push({ displayName: s.displayName, programUnitId: s.programUnitId, reason: entry?.reason || "special_deferred" });
            }
            else {
                newSpecial.push(s);
            }
        }
        lists.failed = newFailed;
        lists.special = newSpecial;
        console.log(`[StatusEngine] deferred suppression: removed ${pendingDeferredUnits.length} unit(s) from fail/special lists`);
    }
    // ── Status decision ───────────────────────────────────────────────────────
    const totalUnits = curriculum.length;
    const failCount = lists.failed.length;
    const missingCount = lists.missing.length;
    const specialCount = lists.special.length;
    const incCount = lists.incomplete.length;
    const officialMean = totalUnits > 0 ? totalFirstAttemptSum / totalUnits : 0;
    const attemptedN = totalUnits - (specialCount + missingCount + incCount);
    const perfMean = attemptedN > 0 ? totalFirstAttemptSum / attemptedN : 0;
    const currentYearDoc = targetYearDoc?.isCurrent
        ? targetYearDoc
        : (await AcademicYear_1.default.findOne({ isCurrent: true }).lean()) || (await AcademicYear_1.default.findOne().sort({ startDate: -1 }).lean());
    const targetSession = targetYearDoc?.session ?? "ORDINARY";
    const [tStart] = (resolvedYearName || "0/0").split("/").map(Number);
    const [gStart] = currentYearDoc?.year ? currentYearDoc.year.split("/").map(Number) : [0];
    const isPastYear = targetYearDoc && gStart > 0 ? tStart < gStart : false;
    const isSessionClosed = targetSession === "CLOSED" || isPastYear;
    let status = "PASS";
    let variant = "success";
    let details = "Proceed to next year.";
    if (!options.forPromotion && targetSession === "ORDINARY" && !isPastYear) {
        status = "SESSION IN PROGRESS";
        variant = "info";
        details = "Marks are currently being entered.";
    }
    else if (missingCount >= 6 && isSessionClosed) {
        status = "DEREGISTERED";
        variant = "error";
        details = `Absent from 6+ (${missingCount}) examinations (ENG 23c).`;
    }
    else if (specialCount > 0 && failCount < totalUnits / 2) {
        const parts = [];
        if (failCount > 0)
            parts.push(`SUPP ${failCount}`);
        parts.push(`SPEC ${specialCount}`);
        if (incCount > 0)
            parts.push(`INC ${incCount}`);
        if (missingCount > 0)
            parts.push(`MISSING ${missingCount}`);
        status = parts.join("; ");
        variant = "info";
        details = `Awaiting specials. Mean: ${perfMean.toFixed(2)}`;
    }
    else if (failCount >= totalUnits / 2 || officialMean < 40) {
        status = "REPEAT YEAR";
        variant = "error";
        details = `Failed >= 50% (${failCount}/${totalUnits}) or Mean (${officialMean.toFixed(2)}) < 40% (ENG 16).`;
    }
    else if (failCount > totalUnits / 3) {
        status = "STAYOUT";
        variant = "warning";
        details = `Failed > 1/3 (${failCount}/${totalUnits}). Retake in next ordinary period (ENG 15h).`;
    }
    else if (failCount > 0 || incCount > 0 || missingCount > 0) {
        const parts = [];
        if (failCount > 0)
            parts.push(`SUPP ${failCount}`);
        if (incCount > 0)
            parts.push(`INC ${incCount}`);
        if (missingCount > 0)
            parts.push(`INC ${missingCount}`);
        status = parts.join("; ");
        variant = "warning";
        details = "Eligible for supplementary exams.";
    }
    else if (pendingDeferredUnits.length > 0 && failCount === 0 && specialCount === 0) {
        // All outstanding units are deferred — student is effectively clear for promotion
        status = "PASS";
        variant = "success";
        details = `All pending units deferred to next ordinary period (ENG.13b/18c). Eligible for promotion.`;
    }
    return {
        status, variant, details,
        weightedMean: officialMean.toFixed(2),
        sessionState: targetSession,
        academicYearName: resolvedYearName,
        summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: missingCount },
        passedList: lists.passed,
        failedList: lists.failed,
        specialList: lists.special,
        deferredList: pendingDeferredUnits,
        missingList: lists.missing,
        incompleteList: lists.incomplete,
    };
};
exports.calculateStudentStatus = calculateStudentStatus;
const promoteStudent = async (studentId) => {
    const student = await Student_1.default.findById(studentId).populate("program");
    if (!student)
        throw new Error("Student not found");
    const st = student.status;
    const currentYear = student.currentYearOfStudy || 1;
    if (["deregistered", "discontinued", "graduated", "graduand"].includes(st))
        return { success: false, message: `Action blocked: Student is ${st}` };
    if (st !== "active" && st !== "repeat")
        return { success: false, message: `Promotion blocked: Student status is ${st}` };
    const auditResult = await (0, academicAudit_1.performAcademicAudit)(studentId);
    if (auditResult.discontinued)
        return { success: false, message: `Discontinued: ${auditResult.reason}` };
    const program = student.program;
    const duration = program.durationYears || 5;
    const currentSession = await AcademicYear_1.default.findOne({ isCurrent: true }).lean();
    const completedYear = currentSession?.year || "N/A";
    // ── ENG.15(b): Historical year validation ──────────────────────────────────
    // Runs ONLY when promoting from the penultimate year into the final year.
    // For a 5-year programme: currentYear = 4 (promoting into Y5).
    // For a 4-year programme: currentYear = 3 (promoting into Y4).
    //
    // WHY HERE AND NOT EARLIER:
    // Years 1→2, 2→3, 3→4 don't need historical validation because carry-forwards
    // and supplementaries from prior years are legitimately still in-flight.
    // ENG.15(b) draws a hard line specifically at final year entry.
    if (currentYear === duration - 1) {
        const histCheck = await (0, historicalYearValidator_1.validateHistoricalYears)(studentId.toString(), student.program._id.toString(), duration - 1);
        if (!histCheck.canEnterFinalYear) {
            // Write a statusEvent so the audit trail shows WHY this student was blocked
            await Student_1.default.findByIdAndUpdate(studentId, {
                $push: {
                    statusEvents: {
                        fromStatus: st,
                        toStatus: st, // status unchanged — they stay where they are
                        date: new Date(),
                        academicYear: completedYear,
                        reason: histCheck.blockReason,
                    },
                },
            });
            return {
                success: false,
                message: histCheck.blockReason,
                eng15bBlock: true,
                blockingUnits: histCheck.blockingUnits,
                yearSummaries: histCheck.yearSummaries,
            };
        }
        // Historical years are clean — fall through to normal Y(n-1) engine check
    }
    // ── Standard current-year status engine (unchanged) ───────────────────────
    const statusResult = await (0, exports.calculateStudentStatus)(student._id, student.program, completedYear, currentYear, { forPromotion: true });
    if (["DEREGISTERED", "DISCONTINUED"].includes(statusResult.status)) {
        await syncTerminalStatusToDb(studentId, statusResult.status, statusResult.details, completedYear);
        return { success: false, message: `Promotion Blocked: ${statusResult.status}`, details: statusResult };
    }
    if (statusResult.status === "REPEAT YEAR") {
        const repeatCount = (student.academicHistory || [])
            .filter((h) => h.isRepeatYear && h.yearOfStudy === currentYear).length + 1;
        const qualifier = academicRules_1.REG_QUALIFIERS.repeatYear(repeatCount);
        await Student_1.default.findByIdAndUpdate(studentId, {
            $set: { status: "repeat", qualifierSuffix: qualifier, remarks: statusResult.details },
            $push: {
                statusEvents: { fromStatus: st, toStatus: "repeat", date: new Date(), academicYear: completedYear, reason: `ENG.16: ${statusResult.details}` },
                statusHistory: { status: "repeat", previousStatus: st, date: new Date(), reason: statusResult.details },
                academicHistory: { academicYear: completedYear, yearOfStudy: currentYear, annualMeanMark: parseFloat(statusResult.weightedMean), weightedContribution: 0, failedUnitsCount: statusResult.summary.failed, isRepeatYear: true, date: new Date() },
            },
        });
        return { success: false, message: `Repeat year required (${qualifier})`, details: statusResult };
    }
    if (statusResult.status === "STAYOUT") {
        await Student_1.default.findByIdAndUpdate(studentId, {
            $set: { remarks: `ENG.15h: ${statusResult.details}` },
            $push: { statusEvents: { fromStatus: st, toStatus: "active", date: new Date(), academicYear: completedYear, reason: `ENG.15h: ${statusResult.details}` } },
        });
        return { success: false, message: "Stay out required (ENG.15h)", details: statusResult };
    }
    if (statusResult.specialList.length > 0 && statusResult.failedList.length === 0)
        return { success: false, message: "Special examinations pending", details: statusResult };
    const rawMean = parseFloat(statusResult.weightedMean);
    const yearWeight = (0, weightingRegistry_1.getYearWeight)(program, student.entryType || "Direct", currentYear);
    const histRecord = {
        academicYear: completedYear, yearOfStudy: currentYear, annualMeanMark: rawMean,
        weightedContribution: rawMean * yearWeight, failedUnitsCount: statusResult.summary.failed,
        isRepeatYear: false, date: new Date(),
    };
    const pendingCF = (student.carryForwardUnits || []).filter((u) => u.status === "pending").length;
    // Graduation path (final year)
    if (currentYear === duration) {
        if (pendingCF > 0)
            return { success: false, message: `Cannot graduate: ${pendingCF} carry-forward unit(s) not yet cleared` };
        const fullHistory = [...(student.academicHistory || []), histRecord];
        const finalWAA = fullHistory.reduce((acc, h) => acc + (h.weightedContribution || 0), 0);
        let classification = "PASS";
        if (finalWAA >= 70)
            classification = "FIRST CLASS HONOURS";
        else if (finalWAA >= 60)
            classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
        else if (finalWAA >= 50)
            classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
        await Student_1.default.findByIdAndUpdate(studentId, {
            $set: { status: "graduand", qualifierSuffix: "", finalWeightedAverage: finalWAA.toFixed(2), classification, graduationYear: new Date().getFullYear(), currentYearOfStudy: currentYear + 1 },
            $push: { academicHistory: histRecord },
        });
        return { success: true, message: `Graduated: ${classification}`, isGraduation: true };
    }
    // Standard promotion to next year
    const nextYear = currentYear + 1;
    let cfGranted = false;
    let cfMessage = "";
    if (statusResult.status !== "PASS") {
        const cfResult = await tryAssessAndGrantCarryForward(studentId, student.program.toString(), currentYear, completedYear);
        cfGranted = cfResult.granted;
        cfMessage = cfGranted
            ? `Promoted to Year ${nextYear} with carry-forward (${cfResult.qualifier}): ${cfResult.cfUnits.map((u) => u.unitCode).join(", ")}`
            : "";
        if (!cfGranted)
            return { success: false, message: `Promotion Blocked: ${cfResult.reason}`, details: statusResult };
    }
    await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            currentYearOfStudy: nextYear, currentSemester: 1,
            ...((!cfGranted && _qualifierShouldClearOnPromotion(student.qualifierSuffix || "")) ? { qualifierSuffix: "" } : {}),
        },
        $push: {
            promotionHistory: { from: currentYear, to: nextYear, date: new Date() },
            academicHistory: histRecord,
            statusHistory: { status: "active", previousStatus: st, date: new Date(), reason: `Promoted to Year ${nextYear}` },
        },
    });
    return { success: true, message: cfGranted ? cfMessage : `Successfully promoted to Year ${nextYear}` };
};
exports.promoteStudent = promoteStudent;
const previewPromotion = async (programId, yearToPromote, academicYearName) => {
    const nextYear = yearToPromote + 1;
    const targetYearDoc = await AcademicYear_1.default.findOne({
        year: academicYearName,
    }).lean();
    if (!targetYearDoc) {
        console.warn(`[previewPromotion] AcademicYear "${academicYearName}" not found.`);
        return {
            totalProcessed: 0,
            eligibleCount: 0,
            blockedCount: 0,
            eligible: [],
            blocked: [],
        };
    }
    const program = (await (await Promise.resolve().then(() => __importStar(require("../models/Program")))).default
        .findById(programId)
        .lean());
    const duration = program?.durationYears || 5;
    // Student sets (same as original — unchanged)
    const admissionStudents = await Student_1.default.find({
        program: programId,
        currentYearOfStudy: yearToPromote,
        admissionAcademicYear: targetYearDoc._id,
    }).lean();
    const [m1, m2] = await Promise.all([
        Mark_1.default.distinct("student", { academicYear: targetYearDoc._id }),
        MarkDirect_1.default.distinct("student", {
            academicYear: targetYearDoc._id,
        }),
    ]);
    const markedIds = new Set([...m1, ...m2].map((id) => id.toString()));
    const admissionIds = new Set(admissionStudents.map((s) => s._id.toString()));
    const returningStudents = await Student_1.default.find({
        program: programId,
        currentYearOfStudy: yearToPromote,
        _id: { $in: Array.from(markedIds), $nin: Array.from(admissionIds) },
    }).lean();
    const adminStudents = await Student_1.default.find({
        program: programId,
        currentYearOfStudy: yearToPromote,
        status: { $in: ["on_leave", "deferred", "deregistered", "discontinued"] },
        $or: [
            { admissionAcademicYear: targetYearDoc._id },
            { "academicHistory.academicYear": academicYearName },
        ],
        _id: {
            $nin: [
                ...Array.from(admissionIds),
                ...returningStudents.map((s) => s._id.toString()),
            ],
        },
    }).lean();
    const allStudents = [
        ...admissionStudents,
        ...returningStudents,
        ...adminStudents,
    ];
    const ADMIN_LABELS = {
        on_leave: "ACADEMIC LEAVE",
        deferred: "DEFERMENT",
        discontinued: "DISCONTINUED",
        deregistered: "DEREGISTERED",
        graduated: "GRADUATED",
    };
    const eligible = [];
    const blocked = [];
    for (const student of allStudents) {
        const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
        const adminLabel = ADMIN_LABELS[student.status];
        const baseReport = {
            id: student._id,
            _id: student._id,
            regNo: student.regNo,
            name: student.name,
            qualifierSuffix: student.qualifierSuffix || "",
            remarks: student.remarks,
            academicLeavePeriod: student.academicLeavePeriod,
        };
        if (isAlreadyPromoted) {
            eligible.push({
                ...baseReport,
                status: "ALREADY PROMOTED",
                reasons: [],
                summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
                details: "",
            });
            continue;
        }
        if (adminLabel) {
            const leaveType = student.academicLeavePeriod?.type?.toUpperCase();
            const adminGrounds = [
                student.academicLeavePeriod?.type || "",
                student.remarks || "",
            ]
                .join(" ")
                .trim() || "other";
            blocked.push({
                ...baseReport,
                status: adminLabel,
                reasons: [leaveType ? `${adminLabel} (${leaveType})` : adminLabel],
                specialGrounds: adminGrounds,
                summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
                details: "",
            });
            continue;
        }
        // ── ENG.15(b) CHECK — runs before calculateStudentStatus when entering final year ──
        // This is the KEY addition. Without this, previewPromotion would classify a
        // student with a pending Y2 carry-forward as PASS at Y4 level and show them
        // as eligible for Y5 entry. With this, the coordinator sees the specific
        // blocking units before attempting the promotion.
        if (yearToPromote === duration - 1) {
            const histCheck = await (0, historicalYearValidator_1.validateHistoricalYears)(student._id.toString(), programId, duration - 1);
            if (!histCheck.canEnterFinalYear) {
                blocked.push({
                    ...baseReport,
                    status: "ENG.15(b) BLOCK",
                    reasons: [histCheck.blockReason],
                    specialGrounds: "eng15b",
                    summary: {
                        totalExpected: 0,
                        passed: 0,
                        failed: histCheck.blockingUnits.length,
                        missing: 0,
                    },
                    details: histCheck.blockReason,
                    eng15bBlock: true,
                    blockingUnits: histCheck.blockingUnits,
                    yearSummaries: histCheck.yearSummaries,
                });
                continue;
            }
        }
        // Standard engine check (unchanged)
        const sr = await (0, exports.calculateStudentStatus)(student._id, programId, academicYearName, yearToPromote, { forPromotion: true });
        const specialGrounds = [
            (sr.specialList || [])
                .map((s) => (s.grounds || "").toLowerCase())
                .join(" "),
            (student.remarks || "").toLowerCase(),
            (student.academicLeavePeriod?.type || "").toLowerCase(),
        ]
            .join(" ")
            .trim() || "other";
        const report = {
            ...baseReport,
            status: sr.status,
            summary: sr.summary,
            reasons: [],
            details: sr.details,
            specialGrounds,
            isEligibleForSupp: !["STAYOUT", "REPEAT YEAR", "DEREGISTERED"].includes(sr.status) &&
                (sr.failedList.length > 0 || sr.specialList.length > 0),
        };
        if (sr.status === "PASS") {
            eligible.push(report);
            continue;
        }
        if (sr.status === "STAYOUT")
            report.reasons.push("ENG 15h: > 1/3 units failed");
        if (sr.status === "REPEAT YEAR")
            report.reasons.push("ENG 16: >= 1/2 units failed or mean < 40%");
        if (sr.failedList?.length)
            report.reasons.push(`${sr.failedList.length} unit(s) require supplementary`);
        if (sr.specialList?.length)
            report.reasons.push(`${sr.specialList.length} special examination(s) pending`);
        blocked.push(report);
    }
    return {
        totalProcessed: allStudents.length,
        eligibleCount: eligible.length,
        blockedCount: blocked.length,
        eligible,
        blocked,
    };
};
exports.previewPromotion = previewPromotion;
// ── bulkPromoteClass ──────────────────────────────────────────────────────────
const bulkPromoteClass = async (programId, yearToPromote, academicYearName) => {
    const nextYear = yearToPromote + 1;
    const students = await Student_1.default.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: { $in: ["active", "repeat"] } });
    const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] };
    for (const student of students) {
        const sid = student._id.toString();
        const rNo = student.regNo;
        const curYr = student.currentYearOfStudy;
        try {
            if (curYr >= nextYear) {
                results.alreadyPromoted++;
                results.promoted++;
                continue;
            }
            const res = await (0, exports.promoteStudent)(sid);
            if (res.success)
                results.promoted++;
            else
                results.failed++;
        }
        catch (err) {
            results.errors.push(`${rNo}: ${err.message}`);
        }
    }
    return results;
};
exports.bulkPromoteClass = bulkPromoteClass;
