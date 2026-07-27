"use strict";
// // serverside/src/utils/academicRules.ts
// // ─── REPLACE ATTEMPT_NOTATIONS and getAttemptLabel ──────────────────────────
Object.defineProperty(exports, "__esModule", { value: true });
exports.getExamTypeTitle = exports.ADMIN_STATUS_LABELS = exports.isLockedStatus = exports.isFirstAttempt = exports.assessCarryForwardEligibility = exports.getStudentEligibility = exports.getAttemptLabel = exports.deriveQualifierSuffix = exports.buildDisplayRegNo = exports.REG_QUALIFIERS = exports.ATTEMPT_LABELS = void 0;
// export const ATTEMPT_NOTATIONS: Record<string, string> = {
//   // Normal first sitting
//   ORDINARY: "B/S",
//   // Repeat year qualifiers (ENG 16a/c) — RP1..RP5
//   REPEAT_YEAR_1: "RP1",
//   REPEAT_YEAR_2: "RP2",
//   REPEAT_YEAR_3: "RP3",
//   REPEAT_YEAR_4: "RP4",
//   REPEAT_YEAR_5: "RP5",
//   // Repeat unit qualifiers (ENG 16b) — RPU1..RPU5 (NEW)
//   REPEAT_UNIT_1: "RPU1",
//   REPEAT_UNIT_2: "RPU2",
//   REPEAT_UNIT_3: "RPU3",
//   // Disciplinary repeat — RP1D..RP5D
//   REPEAT_DISC_1: "RP1D",
//   REPEAT_DISC_2: "RP2D",
//   REPEAT_DISC_3: "RP3D",
//   // Carry forward (ENG 14) — RP1C..RP3C (previously "carry over")
//   CARRY_FORWARD_1: "RP1C",
//   CARRY_FORWARD_2: "RP2C",
//   CARRY_FORWARD_3: "RP3C",
//   // Readmit — RA1..RA5
//   READMIT_1: "RA1",
//   READMIT_2: "RA2",
//   READMIT_3: "RA3",
//   READMIT_4: "RA4",
//   READMIT_5: "RA5",
//   // Readmit second semester
//   READMIT_S2: "RA1S2",
//   // Mid-entry
//   MID_ENTRY_Y2: "M2",
//   MID_ENTRY_Y3: "M3",
//   // Transfer from another university (start of year)
//   TRANSFER_Y2: "TF2",
//   TRANSFER_Y3: "TF3",
//   // Transfer second semester
//   TRANSFER_Y2_S2: "TF2S2",
//   TRANSFER_Y3_S2: "TF3S2",
//   // Supplementary
//   SUPPLEMENTARY: "A/S",
//   // Special exam
//   SPECIAL: "SPEC",
//   // Administrative
//   DEFERRED: "DEFERRED",
//   ACADEMIC_LEAVE: "A/L",
//   DISCONTINUED: "DISC.",
//   DEREGISTERED: "DEREG.",
// };
// export interface AttemptLabelOptions {
//   markAttempt?: string;
//   studentStatus?: string;
//   regNo?: string;
//   repeatYearCount?: number;
//   repeatUnitCount?: number; // NEW: for RPU qualifier
//   isDisciplinary?: boolean; // NEW: for RPnD qualifier
//   isCarryForward?: boolean;
//   isReadmit?: boolean;
//   readmitCount?: number;
//   isMidEntry?: boolean;
//   entryYear?: number; // 2 or 3 for M2/M3
//   isTransfer?: boolean;
//   transferYear?: number; // 2 or 3 for TF2/TF3
//   isSecondSemester?: boolean;
// }
// export const getAttemptLabel = (opts: AttemptLabelOptions): string => {
//   const {
//     markAttempt = "1st",
//     studentStatus = "active",
//     regNo = "",
//     repeatYearCount = 0,
//     repeatUnitCount = 0,
//     isDisciplinary = false,
//     isCarryForward = false,
//     isReadmit = false,
//     readmitCount = 1,
//     isMidEntry = false,
//     entryYear = 2,
//     isTransfer = false,
//     transferYear = 2,
//     isSecondSemester = false,
//   } = opts;
//   const st = studentStatus.toLowerCase().replace(/_/g, " ");
//   const regUpper = regNo.toUpperCase();
//   const attempt = (markAttempt || "1st").toLowerCase();
//   // ── 1. Hard administrative statuses ─────────────────────────────────────
//   if (st === "deferred") return ATTEMPT_NOTATIONS.DEFERRED;
//   if (st === "discontinued") return ATTEMPT_NOTATIONS.DISCONTINUED;
//   if (st === "deregistered") return ATTEMPT_NOTATIONS.DEREGISTERED;
//   if (st === "on leave") return ATTEMPT_NOTATIONS.ACADEMIC_LEAVE;
//   // ── 2. Reg number suffix detection (institutional encoding) ─────────────
//   // Pattern: regNo ends with /YYYY followed by qualifier
//   // e.g. E024-01-1234/2019RP2  →  RP2
//   //      E024-01-1234/2019TF2  →  TF2
//   //      E024-01-1234/2019M3   →  M3
//   //      E024-01-1234/2019RA2  →  RA2
//   const suffixMatch = regUpper.match(
//     /\/\d{4}(RP\d+[DC]?|RPU\d+|RA\d+S2?|TF\d+S2?|M[23]|RP\d+C)$/,
//   );
//   if (suffixMatch) {
//     return suffixMatch[1]; // Return exactly what's encoded in the regNo
//   }
//   // ── 3. Transfer students ─────────────────────────────────────────────────
//   if (isTransfer) {
//     const yr = transferYear === 3 ? "TF3" : "TF2";
//     return isSecondSemester ? `${yr}S2` : yr;
//   }
//   // ── 4. Mid-entry students ────────────────────────────────────────────────
//   if (isMidEntry) {
//     return entryYear === 3 ? "M3" : "M2";
//   }
//   // ── 5. Readmit ───────────────────────────────────────────────────────────
//   if (isReadmit) {
//     const count = Math.min(Math.max(readmitCount, 1), 5);
//     return isSecondSemester ? `RA${count}S2` : `RA${count}`;
//   }
//   // ── 6. Repeat year ───────────────────────────────────────────────────────
//   if (st === "repeat") {
//     const count = Math.min(Math.max(repeatYearCount || 1, 1), 5);
//     return isDisciplinary ? `RP${count}D` : `RP${count}`;
//   }
//   // ── 7. Repeat unit (ENG 16b) ─────────────────────────────────────────────
//   if (repeatUnitCount > 0) {
//     const count = Math.min(Math.max(repeatUnitCount, 1), 5);
//     return `RPU${count}`;
//   }
//   // ── 8. Mark-level attempt ────────────────────────────────────────────────
//   if (attempt === "special") return ATTEMPT_NOTATIONS.SPECIAL;
//   if (attempt === "supplementary") return ATTEMPT_NOTATIONS.SUPPLEMENTARY;
//   if (attempt === "re-take") {
//     if (isCarryForward) {
//       const count = Math.min(Math.max(repeatYearCount || 1, 1), 3);
//       return `RP${count}C`;
//     }
//     return "RP1C"; // Default carry forward
//   }
//   // ── 9. Default: first sitting ────────────────────────────────────────────
//   return ATTEMPT_NOTATIONS.ORDINARY; // B/S
// };
// export const isFirstAttempt = (notation: string): boolean =>
//   notation === "B/S" || notation.startsWith("TF") || notation.startsWith("M");
// export const isLockedStatus = (studentStatus: string): boolean => {
//   const s = studentStatus.toLowerCase().replace(/_/g, " ");
//   return [
//     "on leave",
//     "deferred",
//     "discontinued",
//     "deregistered",
//     "graduated",
//   ].includes(s);
// };
// export const ADMIN_STATUS_LABELS: Record<string, string> = {
//   on_leave: "ACADEMIC LEAVE",
//   deferred: "DEFERMENT",
//   discontinued: "DISCONTINUED",
//   deregistered: "DEREGISTERED",
//   graduated: "GRADUATED",
//   graduand: "GRADUATED",
// };
// export const getExamTypeTitle = (
//   session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED",
// ): string => {
//   return session === "SUPPLEMENTARY"
//     ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS"
//     : "ORDINARY EXAMINATION RESULTS";
// };
// serverside/src/utils/academicRules.ts
//
// QUALIFIER SYSTEM — mapped directly to ENG.13–23 attempt sequences
//
// The qualifier serves TWO purposes:
//   1. Appears in the ATTEMPT column of scoresheets/CMS
//   2. Is APPENDED to the student's regNo in all senate documents
//      e.g. E024-01-1339/2016 becomes E024-01-1339/2016RP1
//
// QUALIFIER LIFECYCLE:
//   - Set on student.qualifierSuffix when the triggering event occurs
//   - Read from student.qualifierSuffix for all document generation
//   - Cleared (set to "") only on graduation or full re-admission reset
//
// ENG.22(b) — The 5-attempt ladder:
//   Attempt 1: B/S        ordinary (1st sitting)
//   Attempt 2: A/S        supplementary (after fail at 1st)
//   Attempt 3: A/CF       ordinary of next year (carry forward / stayout retake)
//   Attempt 4: A/S        supplementary again (after fail at 3rd)
//   Attempt 5: A/RPU      ordinary (repeat unit — enrol for CATs too)
//   → DISCONTINUED after fail at 5th
//
// QUALIFIERS APPENDED TO REG NUMBER (ENG practice):
//   RP1..RP5    Normal repeat year (ENG.16a/c)
//   RP1D..RP5D  Disciplinary repeat
//   RP1C..RP3C  Carry forward (ENG.14) — proceeds to next year
//   RPU1..RPU5  Repeat unit (ENG.16b) — 4th attempt failed, must repeat unit
//   RA1..RA5    Re-admission (ENG.21)
//   RA1S2       Re-admission into 2nd semester
//   M2,M3       Mid-entry (Year 2 or Year 3 entry)
//   TF2,TF3     Transfer from another university
//   TF2S2,TF3S2 Transfer into 2nd semester
//
// NOTES:
//   "External repeat" (RP1E) is REMOVED per the updated qualifier table.
//   Stayout is NOT a persistent qualifier — it's a status. In the next year,
//   the student's attempt column shows their actual attempt type (A/SO, A/SOS).
// ─────────────────────────────────────────────────────────────────────────────
// ATTEMPT COLUMN VALUES (what appears in the ATTEMPT column of scoresheets)
// These are distinct from the qualifier suffixes on regNo.
// ─────────────────────────────────────────────────────────────────────────────
exports.ATTEMPT_LABELS = {
    ORDINARY: "B/S", // First ordinary sitting
    SUPPLEMENTARY: "A/S", // Supplementary (after failing ordinary)
    CARRY_FORWARD_1: "RP1C", // 1st carry forward ordinary (3rd attempt)
    CARRY_FORWARD_2: "RP2C", // 2nd carry forward
    CARRY_FORWARD_3: "RP3C",
    CF_SUPP_1: "A/S", // Supp after failing CF ordinary (4th attempt)
    REPEAT_UNIT: "RPU", // Repeat unit (5th attempt — CATs + exam again)
    STAYOUT_RETAKE: "A/SO", // Stayout student retaking in ordinary year
    STAYOUT_SUPP: "A/SOS", // Stayout student failing ordinary, then supp
    REPEAT_YEAR_1: "B/S", // Repeat year first sitting (marked out of 100%)
    SPECIAL: "SPEC",
    DEFERRED: "DEF",
    ACADEMIC_LEAVE: "A/L",
    DISCONTINUED: "DISC.",
    DEREGISTERED: "DEREG.",
};
// ─────────────────────────────────────────────────────────────────────────────
// REG NUMBER QUALIFIER SUFFIXES (appended to regNo in documents)
// e.g. E024-01-1339/2016  →  E024-01-1339/2016RP1
// ─────────────────────────────────────────────────────────────────────────────
exports.REG_QUALIFIERS = {
    // Normal repeat year — RP1 through RP5
    repeatYear: (n) => `RP${Math.min(n, 5)}`,
    // Disciplinary repeat
    disciplinaryRepeat: (n) => `RP${Math.min(n, 5)}D`,
    // Carry forward — RP1C through RP3C
    carryForward: (n) => `RP${Math.min(n, 3)}C`,
    // Repeat unit (ENG.16b — 4th attempt trigger)
    repeatUnit: (n) => `RPU${Math.min(n, 5)}`,
    // Re-admission
    readmit: (n, secondSemester = false) => secondSemester ? `RA${Math.min(n, 5)}S2` : `RA${Math.min(n, 5)}`,
    // Mid-entry
    midEntry: (year) => `M${year}`,
    // Transfer
    transfer: (year, secondSemester = false) => secondSemester ? `TF${year}S2` : `TF${year}`,
};
// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Build the display regNo with qualifier
// Used in all document generators (scoresheet, CMS, senate reports)
// ─────────────────────────────────────────────────────────────────────────────
const buildDisplayRegNo = (regNo, qualifier) => {
    if (!qualifier || qualifier.trim() === "")
        return regNo;
    return `${regNo}${qualifier}`;
};
exports.buildDisplayRegNo = buildDisplayRegNo;
const deriveQualifierSuffix = (ctx) => {
    const st = (ctx.studentStatus || "").toLowerCase().replace(/_/g, " ");
    // Administrative states — no persistent qualifier on regNo
    if (st === "on leave")
        return "";
    if (st === "deferred")
        return "";
    if (st === "discontinued")
        return "";
    if (st === "deregistered")
        return "";
    if (st === "graduated")
        return "";
    if (st === "graduand")
        return "";
    // Transfer students
    if (ctx.isTransfer && ctx.transferYear) {
        return exports.REG_QUALIFIERS.transfer(ctx.transferYear, ctx.isSecondSemester);
    }
    // Mid-entry students
    if (ctx.entryType === "Mid-Entry-Y2")
        return exports.REG_QUALIFIERS.midEntry(2);
    if (ctx.entryType === "Mid-Entry-Y3")
        return exports.REG_QUALIFIERS.midEntry(3);
    // Re-admission
    if (ctx.isReadmit) {
        return exports.REG_QUALIFIERS.readmit(ctx.readmitCount, ctx.isSecondSemester);
    }
    // Repeat year (ENG.16a/c)
    if (st === "repeat" && ctx.repeatYearCount > 0) {
        return ctx.isDisciplinary
            ? exports.REG_QUALIFIERS.disciplinaryRepeat(ctx.repeatYearCount)
            : exports.REG_QUALIFIERS.repeatYear(ctx.repeatYearCount);
    }
    // Carry forward (ENG.14) — student PASSED the year but carries ≤2 units
    if (ctx.carryForwardCount > 0) {
        return exports.REG_QUALIFIERS.carryForward(ctx.carryForwardCount);
    }
    // Repeat unit (ENG.16b — 4th attempt failed)
    if (ctx.repeatUnitCount > 0) {
        return exports.REG_QUALIFIERS.repeatUnit(ctx.repeatUnitCount);
    }
    return "";
};
exports.deriveQualifierSuffix = deriveQualifierSuffix;
const getAttemptLabel = (opts) => {
    const { markAttempt = "1st", studentStatus = "active", studentQualifier = "", isCarryForward = false, carryForwardCount = 1, isStayoutRetake = false, isRepeatUnit = false, repeatUnitCount = 1, } = opts;
    const st = (studentStatus || "").toLowerCase().replace(/_/g, " ");
    const attempt = (markAttempt || "1st").toLowerCase();
    // Administrative — these students shouldn't appear on most scoresheets
    if (st === "deferred")
        return exports.ATTEMPT_LABELS.DEFERRED;
    if (st === "discontinued")
        return exports.ATTEMPT_LABELS.DISCONTINUED;
    if (st === "deregistered")
        return exports.ATTEMPT_LABELS.DEREGISTERED;
    if (st === "on leave")
        return exports.ATTEMPT_LABELS.ACADEMIC_LEAVE;
    // Special exam
    if (attempt === "special")
        return exports.ATTEMPT_LABELS.SPECIAL;
    // Repeat unit (ENG.16b — 5th attempt ladder)
    if (isRepeatUnit) {
        return `RPU${Math.min(repeatUnitCount, 5)}`;
    }
    // Carry forward (ENG.14 — 3rd attempt, ordinary period of next year)
    if (isCarryForward) {
        return `RP${Math.min(carryForwardCount, 3)}C`;
    }
    // Stayout retake (ENG.15h — ordinary period of next year)
    if (isStayoutRetake)
        return exports.ATTEMPT_LABELS.STAYOUT_RETAKE;
    // Supplementary (2nd or 4th attempt)
    if (attempt === "supplementary")
        return exports.ATTEMPT_LABELS.SUPPLEMENTARY;
    // Repeat year (ENG.16a/c) — first sitting of the repeated year
    if (st === "repeat") {
        // In a repeat year, students sit ALL exams marked out of 100% (full B/S again)
        return exports.ATTEMPT_LABELS.REPEAT_YEAR_1; // "B/S"
    }
    // Re-take (generic re-sit not classified above)
    if (attempt === "re-take") {
        // Check qualifier to determine context
        if (studentQualifier.includes("C"))
            return studentQualifier; // RP1C etc.
        return exports.ATTEMPT_LABELS.CARRY_FORWARD_1; // default
    }
    // Default: first ordinary sitting
    return exports.ATTEMPT_LABELS.ORDINARY;
};
exports.getAttemptLabel = getAttemptLabel;
const getStudentEligibility = (studentStatus, failedFraction, // failed / totalUnits (0–1)
hasSpecial, isCarryForward, isStayoutRetake) => {
    const st = (studentStatus || "").toLowerCase().replace(/_/g, " ");
    // Administrative states never appear on scoresheets
    if (["on leave", "deferred", "discontinued", "deregistered"].includes(st)) {
        return {
            appearsOnOrdinary: false, appearsOnSupplementary: false,
            reason: `Administrative status: ${st}`, attemptLabel: "",
        };
    }
    // REPEAT YEAR (ENG.16a/c) — enrols for ALL units in ordinary
    // Does NOT appear on supplementary — they redo the full year
    if (st === "repeat") {
        return {
            appearsOnOrdinary: true, appearsOnSupplementary: false,
            reason: "ENG.16 repeat year — full re-enrolment, ordinary only",
            attemptLabel: "B/S",
        };
    }
    // STAYOUT (ENG.15h) — failed >1/3 <1/2 — retakes in ORDINARY of next year
    // Does NOT appear on supplementary immediately after failure
    if (failedFraction > 1 / 3 && failedFraction < 1 / 2) {
        return {
            appearsOnOrdinary: true, appearsOnSupplementary: false,
            reason: "ENG.15h stayout — retakes in next ordinary period",
            attemptLabel: "A/SO",
        };
    }
    // SUPP eligible (ENG.13) — failed ≤ 1/3 — appears on SUPPLEMENTARY
    if (failedFraction > 0 && failedFraction <= 1 / 3) {
        return {
            appearsOnOrdinary: false, appearsOnSupplementary: true,
            reason: "ENG.13 supplementary — failed ≤ 1/3 units",
            attemptLabel: "A/S",
        };
    }
    // SPECIAL — appears on supplementary period (ENG.18c)
    if (hasSpecial) {
        return {
            appearsOnOrdinary: false, appearsOnSupplementary: true,
            reason: "ENG.18 special examination",
            attemptLabel: "SPEC",
        };
    }
    // CARRY FORWARD (ENG.14) — appears on ORDINARY of next year
    if (isCarryForward) {
        return {
            appearsOnOrdinary: true, appearsOnSupplementary: false,
            reason: "ENG.14 carry forward — retakes in next ordinary period",
            attemptLabel: "RP1C",
        };
    }
    // Stayout retake (appearing in ordinary this year)
    if (isStayoutRetake) {
        return {
            appearsOnOrdinary: true, appearsOnSupplementary: false,
            reason: "ENG.15h stayout retake — in ordinary period",
            attemptLabel: "A/SO",
        };
    }
    // PASS — doesn't need to appear on either (but shows on CMS)
    return {
        appearsOnOrdinary: false, appearsOnSupplementary: false,
        reason: "Passed — no re-examination needed",
        attemptLabel: "B/S",
    };
};
exports.getStudentEligibility = getStudentEligibility;
const assessCarryForwardEligibility = (failedUnits, // unit codes failed at supplementary
failedDueToNoCA, // units failed because of missing coursework
totalPrescribedUnits) => {
    // ENG.14: Cannot carry forward a unit failed due to missing CA
    const carryableUnits = failedUnits.filter(u => !failedDueToNoCA.includes(u));
    // ENG.14: Maximum 2 units can be carried forward
    if (carryableUnits.length === 0) {
        return { eligible: false, units: [], reason: "No carry-forwardable units" };
    }
    if (carryableUnits.length > 2) {
        return {
            eligible: false, units: [],
            reason: `Failed ${carryableUnits.length} units — exceeds 2-unit carry-forward limit (ENG.14)`,
        };
    }
    if (failedDueToNoCA.length > 0) {
        return {
            eligible: false, units: [],
            reason: "Cannot carry forward units failed due to missing coursework (ENG.15a)",
        };
    }
    return {
        eligible: true,
        units: carryableUnits.slice(0, 2),
        reason: `ENG.14 carry forward — ${carryableUnits.length} unit(s)`,
    };
};
exports.assessCarryForwardEligibility = assessCarryForwardEligibility;
const isFirstAttempt = (label) => label === "B/S" || label.startsWith("TF") || label.startsWith("M");
exports.isFirstAttempt = isFirstAttempt;
const isLockedStatus = (status) => ["on leave", "on_leave", "deferred", "discontinued", "deregistered", "graduated"]
    .includes((status || "").toLowerCase().replace(/_/g, " "));
exports.isLockedStatus = isLockedStatus;
exports.ADMIN_STATUS_LABELS = {
    on_leave: "ACADEMIC LEAVE",
    deferred: "DEFERMENT",
    discontinued: "DISCONTINUED",
    deregistered: "DEREGISTERED",
    graduated: "GRADUATED",
    graduand: "GRADUATED",
};
const getExamTypeTitle = (session) => session === "SUPPLEMENTARY"
    ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS"
    : "ORDINARY EXAMINATION RESULTS";
exports.getExamTypeTitle = getExamTypeTitle;
