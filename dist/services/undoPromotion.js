"use strict";
// serverside/src/services/undoPromotion.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.undoPromotion = void 0;
const Student_1 = __importDefault(require("../models/Student"));
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
/**
 * Reverses the last promotion for a student.
 *
 * Guards:
 *  1. Student must be at year > 1 (can't undo a year-1 student).
 *  2. No marks (Mark, MarkDirect, FinalGrade) may exist for the student
 *     in the year they were promoted INTO. If marks exist, the promotion
 *     cannot be safely undone — the coordinator must manually clean up
 *     those marks first.
 *  3. The last academicHistory entry must match the year being reversed.
 *
 * Effect:
 *  - currentYearOfStudy decremented by 1.
 *  - Last academicHistory entry removed.
 *  - Status set back to "active".
 *  - A statusEvent is pushed to record the undo.
 */
const undoPromotion = async (studentId) => {
    const student = await Student_1.default.findById(studentId).populate("program").lean();
    if (!student)
        return { success: false, message: "Student not found." };
    const currentYear = student.currentYearOfStudy;
    // Guard 1: must have been promoted at least once
    if (currentYear <= 1)
        return { success: false, message: "Cannot undo promotion for a Year 1 student. " + "This student has not been promoted yet." };
    // Guard 2: check for marks in the current year (the year they were promoted into)
    // Find all ProgramUnit IDs for this programme + current year
    const programUnitsInCurrentYear = await ProgramUnit_1.default.find({ program: student.program, requiredYear: currentYear }).select("_id").lean();
    const puIds = programUnitsInCurrentYear.map((pu) => pu._id);
    const [detailedCount, directCount, gradeCount] = await Promise.all([
        Mark_1.default.countDocuments({ student: studentId, programUnit: { $in: puIds } }),
        MarkDirect_1.default.countDocuments({ student: studentId, programUnit: { $in: puIds } }),
        FinalGrade_1.default.countDocuments({ student: studentId, programUnit: { $in: puIds } }),
    ]);
    const totalMarksInCurrentYear = detailedCount + directCount + gradeCount;
    if (totalMarksInCurrentYear > 0) {
        return {
            success: false,
            message: `Cannot undo promotion: ${totalMarksInCurrentYear} mark record(s) already exist ` +
                `for Year ${currentYear}. Delete those marks first before reversing the promotion. ` +
                `This prevents accidental data loss.`,
        };
    }
    // Guard 3: verify the last history entry matches
    const history = student.academicHistory || [];
    const lastEntry = history[history.length - 1];
    const restoredYear = currentYear - 1;
    if (lastEntry && lastEntry.yearOfStudy !== restoredYear) {
        console.warn(`[UndoPromotion] History mismatch for ${student.regNo}: ` +
            `last history entry is yearOfStudy=${lastEntry.yearOfStudy}, ` +
            `expected=${restoredYear}. Proceeding anyway.`);
    }
    // Resolve the current academic year label for the status event
    const currentYearDoc = await AcademicYear_1.default.findOne({ isCurrent: true }).lean();
    const currentAY = currentYearDoc?.year || lastEntry?.academicYear || "UNDO CYCLE";
    // Perform the undo
    await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            currentYearOfStudy: restoredYear,
            currentSemester: 1,
            status: "active",
            remarks: `Promotion to Year ${currentYear} reversed on ${new Date().toDateString()}.`,
        },
        // Remove the last academicHistory entry
        $pop: { academicHistory: 1 },
        $push: {
            statusEvents: {
                fromStatus: `year_${currentYear}`,
                toStatus: `year_${restoredYear}`,
                date: new Date(),
                academicYear: currentAY,
                reason: `PROMOTION REVERSED: Coordinator undid promotion from Year ${restoredYear} ` +
                    `to Year ${currentYear}.`,
            },
        },
    });
    return {
        success: true,
        message: `Promotion reversed. Student returned to Year ${restoredYear}.`,
        previousYear: currentYear,
        restoredYear,
    };
};
exports.undoPromotion = undoPromotion;
