"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.readmitStudent = exports.revertStatusToActive = exports.deferAdmission = exports.grantAcademicLeave = void 0;
// serverside/src/services/academicLeave.ts
const Student_1 = __importDefault(require("../models/Student"));
const date_fns_1 = require("date-fns");
// --- IMPLEMENTING ACADEMIC LEAVE (ENG. 19) ---
const grantAcademicLeave = async (studentId, startDate, endDate, reason, leaveType) => {
    const student = await Student_1.default.findById(studentId).populate("program");
    if (!student)
        throw new Error("Student not found");
    // Calculate duration to check max stay-out
    const yearsRequested = (0, date_fns_1.differenceInYears)(endDate, startDate);
    // ENG 19.d/e: Max stay-out check (10 yrs Eng / 8 yrs Ed)
    const program = student.program;
    const maxStayOut = program.name.includes("Engineering") ? 10 : 8;
    if ((student.totalTimeOutYears || 0) + yearsRequested > maxStayOut) {
        throw new Error(`Leave denied: Exceeds max stay-out period of ${maxStayOut} years.`);
    }
    const dateRange = `${(0, date_fns_1.format)(startDate, "MMM yyyy")} - ${(0, date_fns_1.format)(endDate, "MMM yyyy")}`;
    // We determine the current academic year string for the ledger
    const currentAY = student.academicHistory?.slice(-1)[0]?.academicYear || "Current";
    // Update student status and track time
    return await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            status: "on_leave",
            remarks: `Academic Leave: ${reason} (${dateRange})`,
            academicLeavePeriod: { startDate, endDate, reason, type: leaveType },
            totalTimeOutYears: (student.totalTimeOutYears || 0) + yearsRequested,
        },
        $push: {
            statusEvents: {
                fromStatus: student.status,
                toStatus: "on_leave",
                date: new Date(),
                academicYear: currentAY,
                reason: `${leaveType.toUpperCase()}: ${reason}`
            },
            promotionHistory: {
                from: student.currentYearOfStudy,
                to: student.currentYearOfStudy, // Stay in same year
                date: new Date(),
                remarks: `GRANTED ACADEMIC LEAVE: ${reason}`
            }
        }
    }, { new: true });
};
exports.grantAcademicLeave = grantAcademicLeave;
// --- IMPLEMENTING DEFERMENT (ENG. 20) ---
const deferAdmission = async (studentId, academicYearsToDefer) => {
    const student = await Student_1.default.findById(studentId);
    if (!student || student.currentYearOfStudy !== 1)
        throw new Error("Only new students can defer admission.");
    // ENG 20.a: Senate approval check (usually via UI workflow)
    const endDate = new Date();
    endDate.setFullYear(endDate.getFullYear() + academicYearsToDefer);
    return await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            status: "deferred",
            remarks: `Deferred for ${academicYearsToDefer} year(s).`,
            // Set end date based on academic year
            academicLeavePeriod: { startDate: new Date(), endDate: endDate, reason: "Admission Deferral", type: "other" },
        },
        $push: {
            statusEvents: {
                fromStatus: "new_admission",
                toStatus: "deferred",
                date: new Date(),
                academicYear: "Admission Year",
                reason: `Deferred admission for ${academicYearsToDefer} year(s)`
            },
            promotionHistory: {
                from: 0, // 0 indicates pre-admission/entry
                to: 1,
                date: new Date(),
                remarks: `ADMISSION DEFERRED: ${academicYearsToDefer} Year(s)`
            }
        }
    }, { new: true });
};
exports.deferAdmission = deferAdmission;
// --- UNDO/REVERT STATUS (For both Leave/Defer) ---
const revertStatusToActive = async (studentId) => {
    const now = new Date();
    const student = await Student_1.default.findById(studentId);
    if (!student)
        throw new Error("Student not found");
    const previousStatus = student.status.toUpperCase();
    const currentAY = student.academicHistory?.slice(-1)[0]?.academicYear || "Current";
    return await Student_1.default.findByIdAndUpdate(studentId, {
        $set: {
            status: "active",
            remarks: `Status manually reverted to active from ${previousStatus} on ${now.toDateString()}.`
        },
        $push: {
            statusEvents: {
                fromStatus: previousStatus,
                toStatus: "active",
                date: now,
                academicYear: currentAY,
                reason: "REINSTATEMENT: Student resumed studies."
            },
            promotionHistory: {
                from: student.currentYearOfStudy,
                to: student.currentYearOfStudy,
                date: now,
                remarks: `MANUAL: RETURNED FROM ${previousStatus}`
            }
        },
        $unset: { academicLeavePeriod: 1 }
    }, { new: true });
};
exports.revertStatusToActive = revertStatusToActive;
// --- IMPLEMENTING READMISSION (ENG. 23 / 24) ---
const readmitStudent = async (studentId, remarks) => {
    const student = await Student_1.default.findById(studentId);
    if (!student)
        throw new Error("Student not found");
    // Only allow readmission for terminal or inactive statuses
    const terminalStatuses = ["deregistered", "discontinued", "on_leave", "deferred"];
    if (!terminalStatuses.includes(student.status))
        throw new Error(`Readmission denied: Student is currently ${student.status.toUpperCase()}.`);
    const previousStatus = student.status.toUpperCase();
    const now = new Date();
    // Determine the academic year context
    const lastHistory = student.academicHistory?.slice(-1)[0];
    const currentAY = lastHistory?.academicYear || "NEW CYCLE";
    return await Student_1.default.findByIdAndUpdate(studentId, {
        $set: { status: "active", remarks: `READMITTED: ${remarks} (Prev: ${previousStatus})` },
        $push: {
            statusEvents: { fromStatus: previousStatus, toStatus: "active", date: now, academicYear: currentAY, reason: `OFFICIAL READMISSION: ${remarks}` },
            // 2. Track the return in Promotion History
            promotionHistory: { from: student.currentYearOfStudy, to: student.currentYearOfStudy, date: now, remarks: `READMISSION FROM ${previousStatus}` }
        },
        // Clear any pending leave periods
        $unset: { academicLeavePeriod: 1 }
    }, { new: true });
};
exports.readmitStudent = readmitStudent;
