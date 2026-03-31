// serverside/src/utils/academicRules.ts
// ─── REPLACE ATTEMPT_NOTATIONS and getAttemptLabel ──────────────────────────

export const ATTEMPT_NOTATIONS: Record<string, string> = {
  // Normal first sitting
  ORDINARY: "B/S",

  // Repeat year qualifiers (ENG 16a/c) — RP1..RP5
  REPEAT_YEAR_1: "RP1",
  REPEAT_YEAR_2: "RP2",
  REPEAT_YEAR_3: "RP3",
  REPEAT_YEAR_4: "RP4",
  REPEAT_YEAR_5: "RP5",

  // Repeat unit qualifiers (ENG 16b) — RPU1..RPU5 (NEW)
  REPEAT_UNIT_1: "RPU1",
  REPEAT_UNIT_2: "RPU2",
  REPEAT_UNIT_3: "RPU3",

  // Disciplinary repeat — RP1D..RP5D
  REPEAT_DISC_1: "RP1D",
  REPEAT_DISC_2: "RP2D",
  REPEAT_DISC_3: "RP3D",

  // Carry forward (ENG 14) — RP1C..RP3C (previously "carry over")
  CARRY_FORWARD_1: "RP1C",
  CARRY_FORWARD_2: "RP2C",
  CARRY_FORWARD_3: "RP3C",

  // Readmit — RA1..RA5
  READMIT_1: "RA1",
  READMIT_2: "RA2",
  READMIT_3: "RA3",
  READMIT_4: "RA4",
  READMIT_5: "RA5",

  // Readmit second semester
  READMIT_S2: "RA1S2",

  // Mid-entry
  MID_ENTRY_Y2: "M2",
  MID_ENTRY_Y3: "M3",

  // Transfer from another university (start of year)
  TRANSFER_Y2: "TF2",
  TRANSFER_Y3: "TF3",

  // Transfer second semester
  TRANSFER_Y2_S2: "TF2S2",
  TRANSFER_Y3_S2: "TF3S2",

  // Supplementary
  SUPPLEMENTARY: "A/S",

  // Special exam
  SPECIAL: "SPEC",

  // Administrative
  DEFERRED: "DEFERRED",
  ACADEMIC_LEAVE: "A/L",
  DISCONTINUED: "DISC.",
  DEREGISTERED: "DEREG.",
};

export interface AttemptLabelOptions {
  markAttempt?: string;
  studentStatus?: string;
  regNo?: string;
  repeatYearCount?: number;
  repeatUnitCount?: number; // NEW: for RPU qualifier
  isDisciplinary?: boolean; // NEW: for RPnD qualifier
  isCarryForward?: boolean;
  isReadmit?: boolean;
  readmitCount?: number;
  isMidEntry?: boolean;
  entryYear?: number; // 2 or 3 for M2/M3
  isTransfer?: boolean;
  transferYear?: number; // 2 or 3 for TF2/TF3
  isSecondSemester?: boolean;
}

export const getAttemptLabel = (opts: AttemptLabelOptions): string => {
  const {
    markAttempt = "1st",
    studentStatus = "active",
    regNo = "",
    repeatYearCount = 0,
    repeatUnitCount = 0,
    isDisciplinary = false,
    isCarryForward = false,
    isReadmit = false,
    readmitCount = 1,
    isMidEntry = false,
    entryYear = 2,
    isTransfer = false,
    transferYear = 2,
    isSecondSemester = false,
  } = opts;

  const st = studentStatus.toLowerCase().replace(/_/g, " ");
  const regUpper = regNo.toUpperCase();
  const attempt = (markAttempt || "1st").toLowerCase();

  // ── 1. Hard administrative statuses ─────────────────────────────────────
  if (st === "deferred") return ATTEMPT_NOTATIONS.DEFERRED;
  if (st === "discontinued") return ATTEMPT_NOTATIONS.DISCONTINUED;
  if (st === "deregistered") return ATTEMPT_NOTATIONS.DEREGISTERED;
  if (st === "on leave") return ATTEMPT_NOTATIONS.ACADEMIC_LEAVE;

  // ── 2. Reg number suffix detection (institutional encoding) ─────────────
  // Pattern: regNo ends with /YYYY followed by qualifier
  // e.g. E024-01-1234/2019RP2  →  RP2
  //      E024-01-1234/2019TF2  →  TF2
  //      E024-01-1234/2019M3   →  M3
  //      E024-01-1234/2019RA2  →  RA2
  const suffixMatch = regUpper.match(
    /\/\d{4}(RP\d+[DC]?|RPU\d+|RA\d+S2?|TF\d+S2?|M[23]|RP\d+C)$/,
  );
  if (suffixMatch) {
    return suffixMatch[1]; // Return exactly what's encoded in the regNo
  }

  // ── 3. Transfer students ─────────────────────────────────────────────────
  if (isTransfer) {
    const yr = transferYear === 3 ? "TF3" : "TF2";
    return isSecondSemester ? `${yr}S2` : yr;
  }

  // ── 4. Mid-entry students ────────────────────────────────────────────────
  if (isMidEntry) {
    return entryYear === 3 ? "M3" : "M2";
  }

  // ── 5. Readmit ───────────────────────────────────────────────────────────
  if (isReadmit) {
    const count = Math.min(Math.max(readmitCount, 1), 5);
    return isSecondSemester ? `RA${count}S2` : `RA${count}`;
  }

  // ── 6. Repeat year ───────────────────────────────────────────────────────
  if (st === "repeat") {
    const count = Math.min(Math.max(repeatYearCount || 1, 1), 5);
    return isDisciplinary ? `RP${count}D` : `RP${count}`;
  }

  // ── 7. Repeat unit (ENG 16b) ─────────────────────────────────────────────
  if (repeatUnitCount > 0) {
    const count = Math.min(Math.max(repeatUnitCount, 1), 5);
    return `RPU${count}`;
  }

  // ── 8. Mark-level attempt ────────────────────────────────────────────────
  if (attempt === "special") return ATTEMPT_NOTATIONS.SPECIAL;
  if (attempt === "supplementary") return ATTEMPT_NOTATIONS.SUPPLEMENTARY;

  if (attempt === "re-take") {
    if (isCarryForward) {
      const count = Math.min(Math.max(repeatYearCount || 1, 1), 3);
      return `RP${count}C`;
    }
    return "RP1C"; // Default carry forward
  }

  // ── 9. Default: first sitting ────────────────────────────────────────────
  return ATTEMPT_NOTATIONS.ORDINARY; // B/S
};

export const isFirstAttempt = (notation: string): boolean =>
  notation === "B/S" || notation.startsWith("TF") || notation.startsWith("M");

export const isLockedStatus = (studentStatus: string): boolean => {
  const s = studentStatus.toLowerCase().replace(/_/g, " ");
  return [
    "on leave",
    "deferred",
    "discontinued",
    "deregistered",
    "graduated",
  ].includes(s);
};

export const ADMIN_STATUS_LABELS: Record<string, string> = {
  on_leave: "ACADEMIC LEAVE",
  deferred: "DEFERMENT",
  discontinued: "DISCONTINUED",
  deregistered: "DEREGISTERED",
  graduated: "GRADUATED",
  graduand: "GRADUATED",
};

export const getExamTypeTitle = (
  session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED",
): string => {
  return session === "SUPPLEMENTARY"
    ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS"
    : "ORDINARY EXAMINATION RESULTS";
};
