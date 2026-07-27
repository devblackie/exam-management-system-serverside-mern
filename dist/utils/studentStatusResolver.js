"use strict";
// serverside/src/utils/studentStatusResolver.ts
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolveStudentStatus = void 0;
/**
 * Centrally determines a student's status.
 * Handles variations like "on_leave", "ACADEMIC LEAVE", "deferred", etc.
 */
const resolveStudentStatus = (student) => {
    // Normalize input: "on_leave" -> "ON LEAVE", "deferred" -> "DEFERRED"
    const normalized = (student.status || "").toUpperCase().replace("_", " ").trim();
    // Mapping variations to standard keys
    const statusMap = { "ON LEAVE": "ACADEMIC LEAVE", "ACADEMIC LEAVE": "ACADEMIC LEAVE", DEFERRED: "DEFERMENT", DEFERMENT: "DEFERMENT", DEREGISTERED: "DEREGISTERED", DISCONTINUED: "DISCONTINUED", GRADUAND: "GRADUAND" };
    const finalStatus = statusMap[normalized] || normalized;
    const lockedStatuses = ["ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED", "DISCONTINUED", "GRADUAND"];
    const isLocked = lockedStatuses.includes(finalStatus);
    let rawReason = student.academicLeavePeriod?.reason || student.remarks || "";
    const cleanReason = rawReason.includes(":") ? rawReason.split(":").pop()?.trim() : rawReason;
    return { status: finalStatus, reason: cleanReason || "Reason Pending", isLocked: isLocked };
    // return { status: finalStatus, reason: student.remarks || "No reason provided", isLocked: isLocked};
};
exports.resolveStudentStatus = resolveStudentStatus;
