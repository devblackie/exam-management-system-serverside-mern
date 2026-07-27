"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.performAcademicAudit = void 0;
const Student_1 = __importDefault(require("../models/Student"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const statusEngine_1 = require("./statusEngine");
// export const performAcademicAudit = async ( studentId: string, session?: mongoose.ClientSession ) => {
//   const student = await Student.findById(studentId).session(session || null);
const performAcademicAudit = async (studentId, session) => {
    // ── Load student ───────────────────────────────────────────────────────────
    const query = Student_1.default.findById(studentId);
    if (session)
        query.session(session);
    const student = await query;
    if (!student)
        throw new Error("Student not found for audit.");
    // CHECK 1: ENG.22(b)(i) — 5th Attempt Rule
    //
    // A student is discontinued if they have a FinalGrade where:
    //   - attemptNumber >= 5  (this is their 5th attempt at this unit)
    //   - status !== "PASS"   (they failed it)
    //
    // The attemptNumber field is set by gradeCalculator.ts when computing
    // FinalGrade. It counts FinalGrade documents per (student, programUnit).
    //
    // Note: We look for >= 5 not === 5 because a corrupted import could
    // theoretically create 6 records. We catch anything at or beyond the limit.
    // ─────────────────────────────────────────────────────────────────────────
    const fgQuery = FinalGrade_1.default.findOne({
        student: studentId,
        attemptNumber: { $gte: 5 },
        status: { $ne: "PASS" },
    }).populate({ path: "programUnit", populate: { path: "unit" } });
    if (session)
        fgQuery.session(session);
    const fatalGrade = await fgQuery;
    if (fatalGrade) {
        const unitCode = fatalGrade.programUnit?.unit?.code || "Unknown Unit";
        const attemptN = fatalGrade.attemptNumber ?? 5;
        const reason = `ENG.22(b)(i): Failed unit ${unitCode} at attempt ${attemptN}. Discontinued.`;
        await Student_1.default.findByIdAndUpdate(studentId, {
            $set: { status: "discontinued", remarks: reason },
            $push: {
                statusHistory: {
                    status: "discontinued",
                    previousStatus: student.status,
                    date: new Date(),
                    reason,
                },
            },
        }, session ? { session } : {});
        return { discontinued: true, reason };
    }
    // CHECK 2: ENG.22(b)(ii) — Repeat Year Failure Rule
    //
    // A student who was ALREADY in a repeat year (ENG.16) and fails again
    // (mean < 40% OR ≥ 50% units failed) is discontinued.
    //
    // FIX from original: status should be "repeat", not "active".
    // A student actively repeating a year has status === "repeat".
    // Checking status === "active" means this rule NEVER fired.
    //
    // We also check academicHistory to confirm they are genuinely in a
    // repeat year for the current year of study, not just that their
    // status field is "repeat" from a previous year.
    // ─────────────────────────────────────────────────────────────────────────
    const isInRepeatYear = student.status === "repeat" &&
        (student.academicHistory || []).some((h) => h.yearOfStudy === student.currentYearOfStudy && h.isRepeatYear);
    if (isInRepeatYear) {
        const performance = await (0, statusEngine_1.calculateStudentStatus)(studentId, student.program.toString(), "CURRENT", // FIX: was "N/A" — "CURRENT" explicitly resolves to current year
        student.currentYearOfStudy, { forPromotion: true });
        if (performance.status === "REPEAT YEAR") {
            const reason = `ENG.22(b)(ii): Failed to clear repeat year requirements ` +
                `(Mean ${performance.weightedMean}% — ${performance.summary.failed}/${performance.summary.totalExpected} ` +
                `units failed). Discontinued.`;
            await Student_1.default.findByIdAndUpdate(studentId, {
                $set: { status: "discontinued", remarks: reason },
                $push: {
                    statusHistory: {
                        status: "discontinued",
                        previousStatus: "repeat",
                        date: new Date(),
                        reason,
                    },
                },
            }, session ? { session } : {});
            return { discontinued: true, reason };
        }
    }
    return { discontinued: false, reason: "" };
};
exports.performAcademicAudit = performAcademicAudit;
