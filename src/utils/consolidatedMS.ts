// serverside/src/utils/consolidatedMS.ts

import * as ExcelJS from "exceljs";
import config from "../config/config";
import InstitutionSettings from "../models/InstitutionSettings";
import { resolveStudentStatus } from "./studentStatusResolver";
import { calculateStudentStatus } from "../services/statusEngine";
import { buildDisplayRegNo } from "./academicRules";
import mongoose from "mongoose";
import { buildRichRegNoCMS } from "./scoresheetStudentList";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import MarkDirect from "../models/MarkDirect";
import Mark from "../models/Mark";

interface OfferedUnit { code: string; name: string }

export interface ConsolidatedData {
  programName:   string;
  programId:     string;
  academicYear:  string;
  yearOfStudy:   number;
  session:       "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
  students:      Array<Record<string, unknown>>;
  marks:         Array<Record<string, unknown>>;
  offeredUnits:  OfferedUnit[];
  logoBuffer:    any;
  institutionId: string;
  passMark:      number;
  gradingScale:  Array<{ min: number; grade: string }>;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: derive attempt notation string for CMS ATTEMPT column.
// ─────────────────────────────────────────────────────────────────────────────
function buildAttemptNotation(
  studentStatusRaw: string,
  studentQualifier: string,
  studentMarks: any[],
  yearOfStudy: number,
  academicHistory: any[],
): string {
  const st = studentStatusRaw.toLowerCase().replace(/_/g, " ");

  // ── Hard-coded admin statuses ────────────────────────────────────────────
  if (st === "deferred" || st === "deferment") return "DEF";
  if (st === "on leave" || st === "on_leave" || st === "academic leave")
    return "A/L";
  if (st === "discontinued") return "DISC.";
  if (st === "deregistered") return "DEREG.";
  if (st === "repeat") return "A/RA1";

  // ── Qualifier-based patterns ─────────────────────────────────────────────
  if (studentQualifier) {
    const q = studentQualifier.trim().toUpperCase();

    // Disciplinary repeat: RP1D, RP2D
    if (/^RP(\d+)D$/.test(q)) {
      const n = parseInt(q.replace(/^RP(\d+)D$/, "$1"));
      return `A/RA${n}D`;
    }

    // Carry-forward: RP1C, RP2C, etc.
    if (/^RP\d+C$/i.test(q)) return q;

    // Plain repeat year: RP1, RP2 — no trailing C
    // Must be checked AFTER carry-forward so "RP1C" doesn't match this.
    if (/^RP\d+$/i.test(q)) return q;

    // Repeat unit: RPU1, RPU2
    if (/^RPU\d*/i.test(q)) return q;

    // Re-admission: RA1, RA2
    if (/^RA\d/i.test(q)) return q;

    // Mid-entry: M2, M3
    if (/^M\d$/i.test(q)) return q;

    // Transfer: TF2, TF3
    if (/^TF\d/i.test(q)) return q;
  }

  // ── Derive from mark attempt types ───────────────────────────────────────
  const attemptTypes = studentMarks.map((m: any) =>
    (m.attempt || "1st").toLowerCase(),
  );

  if (attemptTypes.length === 0) {
    const hasRepeat = (academicHistory || []).some(
      (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
    );
    return hasRepeat ? "A/RA1" : "B/S";
  }

  if (attemptTypes.every((a: string) => a === "1st" || a === "special")) {
    const hasRepeat = (academicHistory || []).some(
      (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
    );
    return hasRepeat ? "A/RA1" : "B/S";
  }

  if (attemptTypes.includes("re-take")) {
    if (studentQualifier && /^RP\d+C$/i.test(studentQualifier)) {
      return studentQualifier.trim().toUpperCase();
    }
    return "A/CF";
  }

  if (attemptTypes.includes("supplementary")) return "A/S";

  return "B/S";
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: build a per-student mark map that resolves cross-year marks.
//
// Priority order per unit:
//   1. FinalGrade for the target academic year
//   2. MarkDirect for the target academic year
//   3. Most recent PASSING FinalGrade from ANY prior year (cross-year)
//   4. Mark (detailed) for the target academic year
//   5. "INC"
// ─────────────────────────────────────────────────────────────────────────────
async function buildStudentMarkMap(
  studentId:        string,
  offeredUnits:     Array<{ code: string; name: string }>,
  allProgramUnits:  any[],
  targetAcadYearId: string,
  passMark:         number,
): Promise<Map<string, {
  value:       number | string;
  isCrossYear: boolean;
  isSpecial:   boolean;
}>> {
  const map = new Map<string, { value: number | string; isCrossYear: boolean; isSpecial: boolean }>();

  for (const offered of offeredUnits) {
    const pu = allProgramUnits.find((p: any) => p.unit?.code === offered.code);
    if (!pu) {
      map.set(offered.code, { value: "INC", isCrossYear: false, isSpecial: false });
      continue;
    }
    const puId = pu._id.toString();

    // ── Priority 1: FinalGrade for THIS academic year ─────────────────────
    const fgThisYear = await FinalGrade.findOne({
      student:      studentId,
      programUnit:  puId,
      academicYear: targetAcadYearId,
    }).lean() as any;

    if (fgThisYear) {
      const isPendingSpecial = fgThisYear.isSpecial && (fgThisYear.examTotal70 ?? 0) === 0;
      const mark = fgThisYear.totalMark ?? 0;
      map.set(offered.code, {
        value:       isPendingSpecial ? `${mark}C` : mark,
        isCrossYear: false,
        isSpecial:   fgThisYear.isSpecial || fgThisYear.status === "SPECIAL",
      });
      continue;
    }

    // ── Priority 2: MarkDirect for THIS academic year ─────────────────────
    const mdThisYear = await MarkDirect.findOne({
      student:      studentId,
      programUnit:  puId,
      academicYear: targetAcadYearId,
    }).lean() as any;

    if (mdThisYear) {
      const isPendingSpecial = (mdThisYear.isSpecial || mdThisYear.attempt === "special") &&
                               (mdThisYear.examTotal70 ?? 0) === 0;
      const mark = mdThisYear.agreedMark ?? 0;
      map.set(offered.code, {
        value:       isPendingSpecial ? `${mark}C` : (mark || "INC"),
        isCrossYear: false,
        isSpecial:   mdThisYear.isSpecial || mdThisYear.attempt === "special",
      });
      continue;
    }

    // ── Priority 3: Most recent PASSING FinalGrade from ANY prior year ────
    // Covers students who passed some units in a prior year and aren't
    // re-sitting them in the current year (e.g. special/stayout/repeat).
    const fgAnyYear = await FinalGrade.findOne({
      student:     studentId,
      programUnit: puId,
      status:      "PASS",
    }).sort({ createdAt: -1 }).lean() as any;

    if (fgAnyYear) {
      const mark = fgAnyYear.totalMark ?? 0;
      map.set(offered.code, {
        value:       mark,
        isCrossYear: true,   // grey cell in CMS to distinguish from current-year marks
        isSpecial:   false,
      });
      continue;
    }

    // ── Priority 4: Mark (detailed) for THIS academic year ────────────────
    const dmThisYear = await Mark.findOne({
      student:      studentId,
      programUnit:  puId,
      academicYear: targetAcadYearId,
    }).lean() as any;

    if (dmThisYear) {
      const mark = dmThisYear.agreedMark ?? 0;
      map.set(offered.code, {
        value:       mark || "INC",
        isCrossYear: false,
        isSpecial:   dmThisYear.isSpecial || false,
      });
      continue;
    }

    map.set(offered.code, { value: "INC", isCrossYear: false, isSpecial: false });
  }

  return map;
}

// ─────────────────────────────────────────────────────────────────────────────
// Admin status gate
// ─────────────────────────────────────────────────────────────────────────────
const ADMIN_STATUS_MAP: Record<string, string> = {
  on_leave:           "ACADEMIC LEAVE",
  deferred:           "DEFERMENT",
  discontinued:       "DISCONTINUED",
  deregistered:       "DEREGISTERED",
  graduated:          "GRADUATED",
  repeat:             "",    // run engine
  "on leave":         "ACADEMIC LEAVE",
  "academic leave":   "ACADEMIC LEAVE",
  deferment:          "DEFERMENT",
  stayout:            "",    // run engine
  "already promoted": "",    // run engine
};

// ─────────────────────────────────────────────────────────────────────────────
// MAIN EXPORT
// ─────────────────────────────────────────────────────────────────────────────
export const generateConsolidatedMarkSheet = async (
  data: ConsolidatedData,
): Promise<Buffer> => {
  const {
    programName, academicYear, yearOfStudy,
    students, marks, offeredUnits,
    logoBuffer, institutionId, programId,
  } = data;

  // Institution settings
  const settings = await InstitutionSettings.findOne({ institution: institutionId });
  const passMark = settings?.passMark || 40;

  const workbook  = new ExcelJS.Workbook();
  const sheet     = workbook.addWorksheet("CONSOLIDATED MARKSHEET");
  const fontName  = "Arial";

  const tuColIdx  = 5 + offeredUnits.length;
  const totalCols = tuColIdx + 4;

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin" }, left: { style: "thin" },
    bottom: { style: "thin" }, right: { style: "thin" },
  };
  const doubleBottomBorder: Partial<ExcelJS.Borders> = {
    ...thinBorder, bottom: { style: "double" },
  };

  // ── 1. Headers ─────────────────────────────────────────────────────────────
  const centerColIdx = Math.floor(totalCols / 2);
  if (logoBuffer && logoBuffer.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, { tl: { col: centerColIdx - 1, row: 0 }, ext: { width: 100, height: 60 } });
  }

  const setCenteredHeader = (rowNum: number, text: string, fontSize = 10) => {
    sheet.mergeCells(rowNum, 1, rowNum, totalCols);
    const cell = sheet.getCell(rowNum, 1);
    cell.value = text.toUpperCase();
    cell.style = {
      alignment: { horizontal: "center", vertical: "middle" },
      font: { bold: true, name: fontName, size: fontSize - 1 },
    };
  };

  const examPhaseLabel =
    data.session === "SUPPLEMENTARY"
      ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS"
      : "ORDINARY EXAMINATION RESULTS";

  const yrTxt =
    ["FIRST","SECOND","THIRD","FOURTH","FIFTH"][yearOfStudy - 1] ||
    `${yearOfStudy}TH`;

  setCenteredHeader(4, `${config.instName}`);
  setCenteredHeader(5, `${config.schoolName || "SCHOOL OF ENGINEERING"}`);
  setCenteredHeader(6, `${programName}`);
  setCenteredHeader(7, `CONSOLIDATED MARK SHEET - - ${examPhaseLabel} - ${yrTxt} YEAR - ${academicYear} ACADEMIC YEAR`);
  sheet.getCell(7, 1).font.underline = true;

  // ── 2. Table headers ───────────────────────────────────────────────────────
  const startRow = 9;
  const subRow   = 10;
  sheet.getRow(subRow).height = 48;

  const headers: { [key: number]: string } = {
    1: "S/N", 2: "REG. NO", 3: "NAME", 4: "ATTEMPT",
    [tuColIdx]:     "T U",
    [tuColIdx + 1]: "TOTAL",
    [tuColIdx + 2]: "MEAN",
    [tuColIdx + 3]: "RECOMM.",
    [tuColIdx + 4]: "STUDENT MATTERS",
  };

  Object.entries(headers).forEach(([col, text]) => {
    const colNum = parseInt(col);
    sheet.mergeCells(startRow, colNum, subRow, colNum);
    const cell = sheet.getCell(startRow, colNum);
    cell.value = text;
    cell.style = {
      alignment: {
        horizontal: "center", vertical: "middle",
        textRotation: colNum === 4 ? 90 : 0, wrapText: true,
      },
      font:   { bold: true, size: 7, name: fontName },
      border: doubleBottomBorder,
    };
  });

  offeredUnits.forEach((unit, i) => {
    const colIdx = 5 + i;
    sheet.getCell(startRow, colIdx).value = (i + 1).toString();
    sheet.getCell(startRow, colIdx).style = {
      alignment: { horizontal: "center", vertical: "middle" },
      font: { bold: true, size: 7, name: fontName },
      border: thinBorder,
    };
    sheet.getCell(subRow, colIdx).value = unit.code;
    sheet.getCell(subRow, colIdx).style = {
      alignment: { horizontal: "center", vertical: "middle", textRotation: 90 },
      font: { bold: true, size: 7, name: fontName },
      border: thinBorder,
    };
  });

  // ── 3. Pre-loop setup: resolve target year ID and all program units ─────────
  // Required by buildStudentMarkMap to look up marks across academic years.
  const academicYearDoc = await AcademicYear.findOne({ year: academicYear }).lean() as any;
  const targetAcadYearId = academicYearDoc?._id?.toString() || "";

  const allProgramUnits = await ProgramUnit.find({
    program:      programId,
    requiredYear: yearOfStudy,
  }).populate("unit").lean() as any[];

  // ── 4. Student data rows ───────────────────────────────────────────────────
  const sortedStudents = [...students].sort(
    (a, b) => (String(a.regNo || "")).localeCompare(String(b.regNo || "")),
  );

  let currentIndex = 0;

  for (const student of sortedStudents) {
    const rIdx = 11 + currentIndex;

    // sId extraction
    const rawId = (student as any)._id ?? (student as any).id ?? null;
    let sId = "";
    if (rawId) {
      try { sId = rawId.toString(); } catch { sId = ""; }
    }

    if (!sId || !mongoose.isValidObjectId(sId)) {
      console.warn("[CMS] Skipping invalid _id:", (student as any).regNo ?? "unknown");
      continue;
    }

    // Admin status gate
    const studentStatusRaw = ((student as any).status ?? "")
      .toString().toLowerCase().trim();

    const adminStatusLabel =
      ADMIN_STATUS_MAP[studentStatusRaw] ??
      ADMIN_STATUS_MAP[studentStatusRaw.replace(/_/g, " ")] ??
      null;

    let audit: any;

    if (typeof adminStatusLabel === "string" && adminStatusLabel.length > 0) {
      audit = {
        status:        adminStatusLabel,
        variant:       "info" as const,
        details:       adminStatusLabel,
        weightedMean:  "0.00",
        passedList:    [],
        failedList:    [],
        specialList:   [],
        missingList:   [],
        incompleteList: [],
        summary: { totalExpected: offeredUnits.length, passed: 0, failed: 0, missing: 0, isOnLeave: true },
      };
    } else {
      try {
        audit = await calculateStudentStatus(sId, programId, academicYear, yearOfStudy, { forPromotion: true });
      } catch (err: any) {
        console.error(`[CMS] Engine failed for ${(student as any).regNo}:`, err.message);
        audit = {
          status: "SESSION IN PROGRESS", variant: "info" as const,
          details: "Engine error", weightedMean: "0.00",
          passedList: [], failedList: [], specialList: [],
          missingList: [], incompleteList: [],
          summary: { totalExpected: offeredUnits.length, passed: 0, failed: 0, missing: 0 },
        };
      }
    }

    // Marks for this student (used by buildAttemptNotation)
    const studentMarks = (marks as any[]).filter(
      (m: any) =>
        (m.student?._id?.toString() || m.student?.toString()) === sId,
    );

    const attemptNotation: string = buildAttemptNotation(
      studentStatusRaw,
      (student as any).qualifierSuffix || "",
      studentMarks,
      yearOfStudy,
      (student as any).academicHistory || [],
    );

    // Display name
    const hasReturnHistory = ((student as any).statusHistory || []).some(
      (h: any) =>
        h.status === "ACTIVE" &&
        (h.previousStatus === "ACADEMIC LEAVE" || h.previousStatus === "DEFERMENT"),
    );
    const repeatCount = ((student as any).academicHistory || []).filter(
      (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
    ).length;

    const finalDisplayName = [
      (student as any).name || "",
      hasReturnHistory ? " (REINSTATED)" : "",
      repeatCount > 0  ? ` (RPT${repeatCount})` : "",
    ].join("").toUpperCase();

    // ── rowData — all primitive values ─────────────────────────────────────
    const rowData: any[] = [
      currentIndex + 1,
      buildRichRegNoCMS((student as any).regNo || "", (student as any).qualifierSuffix || ""),
      finalDisplayName,
      attemptNotation,
    ];

    // ── Unit marks: use cross-year-aware mark map ──────────────────────────
    const resolvedStatus = resolveStudentStatus(student as any);

    const markCellMap = await buildStudentMarkMap(
      sId,
      offeredUnits,
      allProgramUnits,
      targetAcadYearId,
      passMark,
    );

    // Parallel array of mark metadata for cell styling (indexed 0 = first unit column)
    const rowMarkData: Array<{ isCrossYear: boolean; isSpecial: boolean } | null> = [];

    offeredUnits.forEach((unit) => {
      if (resolvedStatus.isLocked) {
        rowData.push("");
        rowMarkData.push(null);
        return;
      }
      const m = markCellMap.get(unit.code);
      rowData.push(m ? m.value : "INC");
      rowMarkData.push(m || null);
    });

    // Recommendation
    let recomm = audit.status;
    const isEngineStatus = !adminStatusLabel || adminStatusLabel === "";
    const lockedLabels   = new Set([
      "REPEAT YEAR", "STAYOUT", "DEREGISTERED",
      "ACADEMIC LEAVE", "DEFERMENT", "DISCONTINUED", "GRADUATED",
    ]);

    if (isEngineStatus && !lockedLabels.has(audit.status)) {
      const parts: string[] = [];
      if (audit.failedList?.length)     parts.push(`SUPP ${audit.failedList.length}`);
      if (audit.specialList?.length)    parts.push(`SPEC ${audit.specialList.length}`);
      if (audit.incompleteList?.length) parts.push(`INC ${audit.incompleteList.length}`);
      if (parts.length > 0) recomm = parts.join("; ");
    }

    // Student matters
    const mattersList: string[] = [];
    const leaveType           = (student as any).academicLeavePeriod?.type;
    const remarks             = ((student as any).remarks        || "").toLowerCase();
    const specialGroundsField = ((student as any).specialGrounds || "").toLowerCase();

    if (
      (typeof adminStatusLabel === "string" && adminStatusLabel.length > 0) ||
      ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)
    ) {
      if (leaveType === "financial" || remarks.includes("financial") || specialGroundsField.includes("financial")) {
        mattersList.push("FINANCIAL");
      } else if (
        leaveType === "compassionate" || remarks.includes("compassionate") ||
        remarks.includes("medical")   || specialGroundsField.includes("compassionate")
      ) {
        mattersList.push("COMPASSIONATE");
      } else if (leaveType) {
        mattersList.push(leaveType.toUpperCase());
      }
    }

    for (const spec of (audit.specialList || [])) {
      const g = (spec.grounds || "").split(":").pop()?.trim().toUpperCase() || "";
      if (g && g !== "SPECIAL" && g !== "REASON PENDING") mattersList.push(g);
    }

    const finalMatters = Array.from(new Set(mattersList)).join(", ");

    // Totals
    const totalMarks =
      (audit.passedList || []).reduce((a: number, b: any) => a + (b.mark || 0), 0) +
      (audit.failedList || []).reduce((a: number, b: any) => a + (b.mark || 0), 0);

    const isBlocked =
      audit.summary?.isOnLeave ||
      ["ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED"].includes(audit.status);

    rowData.push(
      audit.summary?.totalExpected ?? offeredUnits.length,
      isBlocked ? "-" : totalMarks,
      isBlocked ? "-" : parseFloat(audit.weightedMean || "0").toFixed(2),
      recomm,
      finalMatters,
    );

    // Write row
    const row = sheet.getRow(rIdx);
    row.values = rowData;

    // Style row
    row.eachCell((cell, colNum) => {
      cell.border    = thinBorder;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font      = { size: 8, name: fontName };

      if (colNum === totalCols - 1) {
        cell.protection = { locked: false };
        const txt = (cell.value?.toString() || "").toUpperCase();
        if (txt.includes("ACADEMIC LEAVE") || txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        }
      }

      if (colNum === totalCols) {
        cell.protection = { locked: false };
        cell.alignment  = { horizontal: "left", vertical: "middle" };
        const txt = (cell.value?.toString() || "").toUpperCase();
        if (txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { bold: true, size: 8, name: fontName };
        }
      }

      if (colNum === 2 || colNum === 3) {
        cell.alignment = { horizontal: "left", vertical: "middle" };
      }

      if (colNum >= 5 && colNum < tuColIdx) {
        const val     = cell.value?.toString() || "";
        // 0-indexed position within the unit columns
        const unitIdx = colNum - 5;
        const mData   = rowMarkData[unitIdx];

        if (resolvedStatus.isLocked) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
        } else if (val === "INC" || val.endsWith("C")) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        } else if (typeof cell.value === "number" && cell.value < passMark) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
          cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 8, name: fontName };
        } else if (mData?.isCrossYear) {
          // Mark came from a prior academic year — light grey, italic to distinguish
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
          cell.font = { size: 8, name: fontName, italic: true };
        }
      }
    });

    row.getCell(totalCols).protection = { locked: false };
    currentIndex++;
  }

  const lastDataRow = 10 + sortedStudents.length;

  // ── 5. Unit statistics ─────────────────────────────────────────────────────
  const statsStart  = lastDataRow + 2;
  const statsLabels = [
    "Mean", "Standard Deviation", "Maximum", "Minimum",
    "No. of Candidates", "No. of Passes", "No. of Fails", "No. of Blanks",
  ];

  statsLabels.forEach((label, i) => {
    const rIdx = statsStart + i;
    const r    = sheet.getRow(rIdx);
    r.height   = 15;

    const labelCell = r.getCell(3);
    labelCell.value  = label;
    labelCell.font   = { bold: true, size: 7, name: fontName };
    labelCell.border = thinBorder;

    offeredUnits.forEach((_, uIdx) => {
      const colIdx    = 5 + uIdx;
      const colLetter = sheet.getColumn(colIdx).letter;
      const cell      = r.getCell(colIdx);
      const range     = `${colLetter}11:${colLetter}${lastDataRow}`;

      cell.border = thinBorder;
      cell.numFmt = "0.0";
      cell.font   = { size: 7, name: fontName };

      if (label === "No. of Passes") {
        cell.value  = { formula: `COUNTIF(${range}, ">=${passMark}")` };
        cell.numFmt = "0";
      } else if (label === "No. of Fails") {
        cell.value  = { formula: `COUNTIFS(${range}, "<${passMark}", ${range}, "<>")` };
        cell.numFmt = "0";
      } else if (label === "No. of Blanks") {
        cell.value  = { formula: `COUNTIF(${range}, "INC")` };
        cell.numFmt = "0";
      } else {
        const funcMap: Record<string, string> = {
          "Mean": "AVERAGE", "Standard Deviation": "STDEV.P",
          "Maximum": "MAX", "Minimum": "MIN", "No. of Candidates": "COUNT",
        };
        const fn = funcMap[label];
        if (fn) cell.value = { formula: `IFERROR(ROUND(${fn}(${range}), 1), 0)` };
      }
    });

    r.getCell(3).border = { ...r.getCell(3).border, left: { style: "thick" } };
    r.getCell(tuColIdx - 1).border = { ...r.getCell(tuColIdx - 1).border, right: { style: "thick" } };

    if (i === 0) {
      for (let c = 3; c < tuColIdx; c++)
        r.getCell(c).border = { ...r.getCell(c).border, top: { style: "thick" } };
    }
    if (i === statsLabels.length - 1) {
      for (let c = 3; c < tuColIdx; c++)
        r.getCell(c).border = { ...r.getCell(c).border, bottom: { style: "thick" } };
    }
  });

  // ── 6. Summary table ───────────────────────────────────────────────────────
  const summaryStart      = lastDataRow + 12;
  const summaryHeaderCell = sheet.getCell(`B${summaryStart}`);
  summaryHeaderCell.value = "SUMMARY";
  summaryHeaderCell.font  = { bold: true, size: 10, underline: true, name: fontName };

  const summaryData: Record<string, number> = {
    PASS: 0, SUPPLEMENTARY: 0, "REPEAT YEAR": 0, "STAY OUT": 0,
    SPECIAL: 0, INCOMPLETE: 0, "ACADEMIC LEAVE": 0,
    DEFERMENT: 0, "DEREGISTERED/DISC": 0,
  };

  sheet.getColumn(tuColIdx + 3).eachCell({ includeEmpty: false }, (cell, rowNum) => {
    if (rowNum > 10 && rowNum <= lastDataRow) {
      const txt = cell.value?.toString().toUpperCase() || "";
      if      (txt === "PASS")                                summaryData.PASS++;
      else if (txt.includes("SUPP"))                          summaryData.SUPPLEMENTARY++;
      else if (txt.includes("REPEAT"))                        summaryData["REPEAT YEAR"]++;
      else if (txt.includes("STAY OUT"))                      summaryData["STAY OUT"]++;
      else if (txt.includes("SPEC"))                          summaryData.SPECIAL++;
      else if (txt.includes("ACADEMIC LEAVE"))                summaryData["ACADEMIC LEAVE"]++;
      else if (txt.includes("DEFERMENT"))                     summaryData.DEFERMENT++;
      else if (txt.includes("INC"))                           summaryData.INCOMPLETE++;
      else if (txt.includes("DEREG") || txt.includes("DISC")) summaryData["DEREGISTERED/DISC"]++;
    }
  });

  const activeSummaryEntries = Object.entries(summaryData).filter(([, count]) => count > 0);
  activeSummaryEntries.forEach(([label, count], i) => {
    const rIdx      = summaryStart + 1 + i;
    const labelCell = sheet.getCell(`B${rIdx}`);
    const countCell = sheet.getCell(`C${rIdx}`);
    labelCell.value  = label;
    countCell.value  = count;
    labelCell.border = thinBorder;
    countCell.border = thinBorder;
    labelCell.font   = { size: 8, name: fontName, bold: true };
    countCell.font   = { size: 8, name: fontName };
  });

  // ── 7. Offered units table ─────────────────────────────────────────────────
  const unitsStart = summaryStart + activeSummaryEntries.length + 4;
  sheet.mergeCells(unitsStart, 2, unitsStart, 6);
  sheet.getCell(unitsStart, 2).value = "LIST OF UNITS OFFERED";
  sheet.getCell(unitsStart, 2).font  = { bold: true, underline: true };

  const mid = Math.ceil(offeredUnits.length / 2);
  for (let i = 0; i < mid; i++) {
    const rIdx  = unitsStart + 2 + i;
    const r     = sheet.getRow(rIdx);
    const left  = offeredUnits[i];
    const right = offeredUnits[mid + i];

    r.getCell(2).value = i + 1;
    r.getCell(3).value = left.code;
    sheet.mergeCells(rIdx, 4, rIdx, 7);
    r.getCell(4).value = left.name;

    if (right) {
      r.getCell(9).value  = mid + i + 1;
      r.getCell(10).value = right.code;
      sheet.mergeCells(rIdx, 11, rIdx, 14);
      r.getCell(11).value = right.name;
    }

    [2, 3, 4, 9, 10, 11].forEach((col) => {
      const cell = r.getCell(col);
      cell.border = { ...thinBorder };
      if (col === 2)                           cell.border.left   = { style: "thick" };
      if (col === 11 || (!right && col === 4)) cell.border.right  = { style: "thick" };
      if (i === 0)                             cell.border.top    = { style: "thick" };
      if (i === mid - 1)                       cell.border.bottom = { style: "thick" };
      cell.font = { size: 8 };
    });
  }

  // ── 8. Main table thick borders ────────────────────────────────────────────
  for (let i = startRow; i <= lastDataRow; i++) {
    sheet.getCell(i, 1).border =
      { ...sheet.getCell(i, 1).border, left:  { style: "thick" } };
    sheet.getCell(i, totalCols).border =
      { ...sheet.getCell(i, totalCols).border, right: { style: "thick" } };
  }
  sheet.getRow(startRow).eachCell(
    (c) => (c.border = { ...c.border, top:    { style: "thick" } }),
  );
  sheet.getRow(lastDataRow).eachCell(
    (c) => (c.border = { ...c.border, bottom: { style: "thick" } }),
  );

  // ── 9. Sheet formatting ────────────────────────────────────────────────────
  sheet.getColumn(1).width = 4;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 25;
  sheet.getColumn(4).width = 7;
  offeredUnits.forEach((_, i) => (sheet.getColumn(5 + i).width = 4.5));
  sheet.getColumn(tuColIdx).width     = 5;
  sheet.getColumn(tuColIdx + 1).width = 7;
  sheet.getColumn(tuColIdx + 2).width = 7;
  sheet.getColumn(tuColIdx + 3).width = 20;
  sheet.getColumn(tuColIdx + 4).width = 20;

  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 10 }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup = {
    orientation: "landscape", paperSize: 9,
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
  };

  const result = await workbook.xlsx.writeBuffer();
  return Buffer.from(result as ArrayBuffer);
};



































// import * as ExcelJS from "exceljs";
// import config from "../config/config";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { resolveStudentStatus } from "./studentStatusResolver";
// import { calculateStudentStatus } from "../services/statusEngine";
// import { buildDisplayRegNo } from "./academicRules";
// import mongoose from "mongoose";

// interface OfferedUnit { code: string; name: string }

// export interface ConsolidatedData {
//   programName:   string;
//   programId:     string;
//   academicYear:  string;
//   yearOfStudy:   number;
//   session:       "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
//   students:      Array<Record<string, unknown>>;
//   marks:         Array<Record<string, unknown>>;
//   offeredUnits:  OfferedUnit[];
//   logoBuffer:    any;
//   institutionId: string;
//   passMark:      number;
//   gradingScale:  Array<{ min: number; grade: string }>;
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // HELPER: derive attempt notation string for CMS ATTEMPT column.
// //
// // This is a REGULAR FUNCTION — call it with arguments.
// // Its RETURN VALUE (a string) goes into the Excel cell, not the function itself.
// // ─────────────────────────────────────────────────────────────────────────────
// function buildAttemptNotation(
//   studentStatusRaw: string,
//   studentQualifier: string,
//   studentMarks:     any[],
//   yearOfStudy:      number,
//   academicHistory:  any[],
// ): string {
//   const st = studentStatusRaw.toLowerCase().replace(/_/g, " ");

//   if (st === "deferred"     || st === "deferment")     return "DEF";
//   if (st === "on leave"     || st === "on_leave"
//    || st === "academic leave")                         return "A/L";
//   if (st === "discontinued")                           return "DISC.";
//   if (st === "deregistered")                           return "DEREG.";
//   if (st === "repeat")                                 return "A/RA1";

//   // Carry-forward qualifier
//   if (studentQualifier && /RP\d+C/i.test(studentQualifier)) return studentQualifier;
//   // Repeat unit
//   if (studentQualifier && studentQualifier.startsWith("RPU")) return studentQualifier;
//   // Re-admission
//   if (studentQualifier && /^RA\d/i.test(studentQualifier)) return studentQualifier;

//   // Derive from mark attempt types — also check isSpecial flag
//   const attemptTypes = studentMarks.map(
//     (m: any) => {
//       if ((m as any).isSpecial === true || (m as any).attempt === "special" || (m as any).status === "SPECIAL") return "special";
//       return (m.attempt || "1st").toLowerCase();
//     },
//   );

//   if (attemptTypes.length === 0) {
//     const hasRepeat = (academicHistory || []).some((h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy);
//     return hasRepeat ? "A/RA1" : "B/S";
//   }

//   // If ALL marks are special → "SPEC"
//   if (attemptTypes.length > 0 && attemptTypes.every((a: string) => a === "special")) return "SPEC";

//   if (attemptTypes.every((a: string) => a === "1st" || a === "special")) {
//     const hasRepeat = (academicHistory || []).some((h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy);
//     return hasRepeat ? "A/RA1" : "B/S";
//   }

//   if (attemptTypes.includes("re-take")) {
//     return studentQualifier && /RP\d+C/i.test(studentQualifier) ? studentQualifier : "A/CF";
//   }
//   if (attemptTypes.includes("supplementary")) return "A/S";

//   return "B/S";
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // Admin status gate
// // ─────────────────────────────────────────────────────────────────────────────
// const ADMIN_STATUS_MAP: Record<string, string> = {
//   on_leave:           "ACADEMIC LEAVE",
//   deferred:           "DEFERMENT",
//   discontinued:       "DISCONTINUED",
//   deregistered:       "DEREGISTERED",
//   graduated:          "GRADUATED",
//   repeat:             "",    // run engine
//   "on leave":         "ACADEMIC LEAVE",
//   "academic leave":   "ACADEMIC LEAVE",
//   deferment:          "DEFERMENT",
//   stayout:            "",    // run engine
//   "already promoted": "",    // run engine
// };

// // ─────────────────────────────────────────────────────────────────────────────
// // MAIN EXPORT
// // ─────────────────────────────────────────────────────────────────────────────
// export const generateConsolidatedMarkSheet = async (
//   data: ConsolidatedData,
// ): Promise<Buffer> => {
//   const {programName, academicYear, yearOfStudy, students, marks, offeredUnits, logoBuffer, institutionId, programId} = data;

//   // Institution settings — unchanged as instructed
//   const settings = await InstitutionSettings.findOne({ institution: institutionId });
//   const passMark = settings?.passMark || 40;

//   const workbook  = new ExcelJS.Workbook();
//   const sheet     = workbook.addWorksheet("CONSOLIDATED MARKSHEET");
//   const fontName  = "Arial";

//   const tuColIdx  = 5 + offeredUnits.length;
//   const totalCols = tuColIdx + 4;

//   const thinBorder: Partial<ExcelJS.Borders> = {top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }};
//   const doubleBottomBorder: Partial<ExcelJS.Borders> = { ...thinBorder, bottom: { style: "double" }};

//   // ── 1. Headers ─────────────────────────────────────────────────────────────
//   const centerColIdx = Math.floor(totalCols / 2);
//   if (logoBuffer && logoBuffer.length > 0) {
//     const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
//     sheet.addImage(logoId, { tl: { col: centerColIdx - 1, row: 0 }, ext: { width: 100, height: 60 }});
//   }

//   const setCenteredHeader = (rowNum: number, text: string, fontSize = 10) => {
//     sheet.mergeCells(rowNum, 1, rowNum, totalCols);
//     const cell = sheet.getCell(rowNum, 1);
//     cell.value = text.toUpperCase();
//     cell.style = {
//       alignment: { horizontal: "center", vertical: "middle" },
//       font: { bold: true, name: fontName, size: fontSize - 1 },
//     };
//   };

//   const examPhaseLabel =
//     data.session === "SUPPLEMENTARY" ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS" : "ORDINARY EXAMINATION RESULTS";

//   const yrTxt = ["FIRST","SECOND","THIRD","FOURTH","FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;

//   setCenteredHeader(4, `${config.instName}`);
//   setCenteredHeader(5, `${config.schoolName || "SCHOOL OF ENGINEERING"}`);
//   setCenteredHeader(6, `${programName}`);
//   setCenteredHeader(7, `CONSOLIDATED MARK SHEET - - ${examPhaseLabel} - ${yrTxt} YEAR - ${academicYear} ACADEMIC YEAR`);
//   sheet.getCell(7, 1).font.underline = true;

//   // ── 2. Table headers ───────────────────────────────────────────────────────
//   const startRow = 9;
//   const subRow   = 10;
//   sheet.getRow(subRow).height = 48;

//   const headers: { [key: number]: string } = {
//     1: "S/N", 2: "REG. NO", 3: "NAME", 4: "ATTEMPT",
//     [tuColIdx]:     "T U",
//     [tuColIdx + 1]: "TOTAL",
//     [tuColIdx + 2]: "MEAN",
//     [tuColIdx + 3]: "RECOMM.",
//     [tuColIdx + 4]: "STUDENT MATTERS",
//   };

//   Object.entries(headers).forEach(([col, text]) => {
//     const colNum = parseInt(col);
//     sheet.mergeCells(startRow, colNum, subRow, colNum);
//     const cell = sheet.getCell(startRow, colNum);
//     cell.value = text;
//     cell.style = {
//       alignment: {
//         horizontal: "center", vertical: "middle",
//         textRotation: colNum === 4 ? 90 : 0, wrapText: true,
//       },
//       font:   { bold: true, size: 7, name: fontName },
//       border: doubleBottomBorder,
//     };
//   });

//   offeredUnits.forEach((unit, i) => {
//     const colIdx = 5 + i;
//     sheet.getCell(startRow, colIdx).value = (i + 1).toString();
//     sheet.getCell(startRow, colIdx).style = {
//       alignment: { horizontal: "center", vertical: "middle" },
//       font: { bold: true, size: 7, name: fontName },
//       border: thinBorder,
//     };
//     sheet.getCell(subRow, colIdx).value = unit.code;
//     sheet.getCell(subRow, colIdx).style = {
//       alignment: { horizontal: "center", vertical: "middle", textRotation: 90 },
//       font: { bold: true, size: 7, name: fontName }, border: thinBorder};
//   });

//   // ── 3. Student data rows ───────────────────────────────────────────────────
//   const sortedStudents = [...students].sort(
//     (a, b) => (String(a.regNo || "")).localeCompare(String(b.regNo || "")),
//   );

//   let currentIndex = 0;

//   for (const student of sortedStudents) {
//     const rIdx = 11 + currentIndex;

//     // sId extraction
//     const rawId = (student as any)._id ?? (student as any).id ?? null;
//     let sId = "";
//     if (rawId) {
//       try { sId = rawId.toString(); } catch { sId = ""; }
//     }

//     if (!sId || !mongoose.isValidObjectId(sId)) {
//       console.warn("[CMS] Skipping invalid _id:", (student as any).regNo ?? "unknown");
//       continue;
//     }

//     // Admin status gate
//     const studentStatusRaw = ((student as any).status ?? "").toString().toLowerCase().trim();

//     const adminStatusLabel =
//       ADMIN_STATUS_MAP[studentStatusRaw] ??
//       ADMIN_STATUS_MAP[studentStatusRaw.replace(/_/g, " ")] ??
//       null;

//     let audit: any;

//     if (typeof adminStatusLabel === "string" && adminStatusLabel.length > 0) {
//       audit = {
//         status:        adminStatusLabel,
//         variant:       "info" as const,
//         details:       adminStatusLabel,
//         weightedMean:  "0.00",
//         passedList:    [],
//         failedList:    [],
//         specialList:   [],
//         missingList:   [],
//         incompleteList: [],
//         summary: {
//           totalExpected: offeredUnits.length,
//           passed: 0, failed: 0, missing: 0, isOnLeave: true,
//         },
//       };
//     } else {
//       try {
//         audit = await calculateStudentStatus(
//           sId, programId, academicYear, yearOfStudy, { forPromotion: true },
//         );
//       } catch (err: any) {
//         console.error(`[CMS] Engine failed for ${(student as any).regNo}:`, err.message);
//         audit = {
//           status: "SESSION IN PROGRESS", variant: "info" as const,
//           details: "Engine error", weightedMean: "0.00",
//           passedList: [], failedList: [], specialList: [],
//           missingList: [], incompleteList: [],
//           summary: { totalExpected: offeredUnits.length, passed: 0, failed: 0, missing: 0 },
//         };
//       }
//     }

//     // Marks for this student
//     const studentMarks = (marks as any[]).filter(
//       (m: any) =>
//         (m.student?._id?.toString() || m.student?.toString()) === sId,
//     );

//     // ── THE FIX: CALL the function, use the returned STRING ────────────────
//     const attemptNotation: string = buildAttemptNotation(
//       studentStatusRaw,
//       (student as any).qualifierSuffix || "",
//       studentMarks,
//       yearOfStudy,
//       (student as any).academicHistory || [],
//     );

//     // Display regNo with qualifier
//     const displayRegNo = buildDisplayRegNo((student as any).regNo || "", (student as any).qualifierSuffix || "");

//     // Display name
//     const hasReturnHistory = ((student as any).statusHistory || []).some(
//       (h: any) =>
//         h.status === "ACTIVE" &&
//         (h.previousStatus === "ACADEMIC LEAVE" || h.previousStatus === "DEFERMENT"),
//     );
//     const repeatCount = ((student as any).academicHistory || []).filter(
//       (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
//     ).length;

//     const finalDisplayName = [
//       (student as any).name || "",
//       hasReturnHistory ? " (REINSTATED)" : "",
//       repeatCount > 0  ? ` (RPT${repeatCount})` : "",
//     ].join("").toUpperCase();

//     // ── rowData — all primitive values ─────────────────────────────────────
//     const rowData: any[] = [
//       currentIndex + 1,
//       displayRegNo,         // string
//       finalDisplayName,     // string
//       attemptNotation,      // string — RETURN VALUE of the function call above
//     ];

//     // Unit marks
//     const resolvedStatus = resolveStudentStatus(student as any);

//     offeredUnits.forEach((unit) => {
//       if (resolvedStatus.isLocked) { rowData.push(""); return; }

//       const markObj = (marks as any[]).find(
//         (m: any) =>
//           (m.student?._id?.toString() || m.student?.toString()) === sId &&
//           m.programUnit?.unit?.code === unit.code,
//       );

//       if (!markObj) { rowData.push("INC"); return; }

//       const isSpecialMark =
//         (markObj as any).isSpecial === true ||
//         (markObj as any).attempt   === "special" ||
//         (markObj as any).status    === "SPECIAL" ||
//         ((markObj as any).remarks || "").toLowerCase().includes("special");
//       const markValue = (markObj as any).agreedMark ?? 0;

//       if (isSpecialMark) {
//         // Special exam — show mark with "C" suffix per ENG.18
//         rowData.push(markValue > 0 ? `${markValue}C` : "INC");
//       } else if (markValue > 0) {
//         // Valid agreed mark exists — always show it
//         // (do NOT gate on caTotal30/examTotal70 being non-null; those may be
//         //  absent for direct-entry marks or FinalGrade-merged records)
//         rowData.push(markValue);
//       } else {
//         // Mark is genuinely zero or missing — check if student sat the exam
//         const caTotal30   = (markObj as any).caTotal30;
//         const examTotal70 = (markObj as any).examTotal70;
//         const hasSatExam  = (caTotal30 != null && caTotal30 > 0) ||
//                             (examTotal70 != null && examTotal70 > 0);
//         rowData.push(hasSatExam ? 0 : "INC");
//       }
//     });

//     // Recommendation
//     let recomm = audit.status;
//     const isEngineStatus = !adminStatusLabel || adminStatusLabel === "";
//     const lockedLabels   = new Set([
//       "REPEAT YEAR", "STAYOUT", "DEREGISTERED",
//       "ACADEMIC LEAVE", "DEFERMENT", "DISCONTINUED", "GRADUATED",
//     ]);

//     if (isEngineStatus && !lockedLabels.has(audit.status)) {
//       const parts: string[] = [];
//       if (audit.failedList?.length)     parts.push(`SUPP ${audit.failedList.length}`);
//       if (audit.specialList?.length)    parts.push(`SPEC ${audit.specialList.length}`);
//       if (audit.incompleteList?.length) parts.push(`INC ${audit.incompleteList.length}`);
//       if (parts.length > 0) recomm = parts.join("; ");
//     }

//     // Student matters
//     const mattersList: string[] = [];
//     const leaveType           = (student as any).academicLeavePeriod?.type;
//     const remarks             = ((student as any).remarks        || "").toLowerCase();
//     const specialGroundsField = ((student as any).specialGrounds || "").toLowerCase();

//     if (
//       (typeof adminStatusLabel === "string" && adminStatusLabel.length > 0) ||
//       ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)
//     ) {
//       if (leaveType === "financial" || remarks.includes("financial") || specialGroundsField.includes("financial")) {
//         mattersList.push("FINANCIAL");
//       } else if (
//         leaveType === "compassionate" || remarks.includes("compassionate") ||
//         remarks.includes("medical")   || specialGroundsField.includes("compassionate")
//       ) {
//         mattersList.push("COMPASSIONATE");
//       } else if (leaveType) {
//         mattersList.push(leaveType.toUpperCase());
//       }
//     }

//     for (const spec of (audit.specialList || [])) {
//       const g = (spec.grounds || "").split(":").pop()?.trim().toUpperCase() || "";
//       if (g && g !== "SPECIAL" && g !== "REASON PENDING") mattersList.push(g);
//     }

//     const finalMatters = Array.from(new Set(mattersList)).join(", ");

//     // Totals
//     const totalMarks =
//       (audit.passedList || []).reduce((a: number, b: any) => a + (b.mark || 0), 0) +
//       (audit.failedList || []).reduce((a: number, b: any) => a + (b.mark || 0), 0);

//     const isBlocked =
//       audit.summary?.isOnLeave ||
//       ["ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED"].includes(audit.status);

//     rowData.push(
//       audit.summary?.totalExpected ?? offeredUnits.length,
//       isBlocked ? "-" : totalMarks,
//       isBlocked ? "-" : parseFloat(audit.weightedMean || "0").toFixed(2),
//       recomm,
//       finalMatters,
//     );

//     // Write row
//     const row = sheet.getRow(rIdx);
//     row.values = rowData;

//     // Style row
//     row.eachCell((cell, colNum) => {
//       cell.border    = thinBorder;
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.font      = { size: 8, name: fontName };

//       if (colNum === totalCols - 1) {
//         cell.protection = { locked: false };
//         const txt = (cell.value?.toString() || "").toUpperCase();
//         if (txt.includes("ACADEMIC LEAVE") || txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//         }
//       }

//       if (colNum === totalCols) {
//         cell.protection = { locked: false };
//         cell.alignment  = { horizontal: "left", vertical: "middle" };
//         const txt = (cell.value?.toString() || "").toUpperCase();
//         if (txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//           cell.font = { bold: true, size: 8, name: fontName };
//         }
//       }

//       if (colNum === 2 || colNum === 3) {
//         cell.alignment = { horizontal: "left", vertical: "middle" };
//       }

//       if (colNum >= 5 && colNum < tuColIdx) {
//         const val = cell.value?.toString() || "";
//         if (resolvedStatus.isLocked) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
//         } else if (val === "INC" || val.endsWith("C")) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//         } else if (typeof cell.value === "number" && cell.value < passMark) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
//           cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 8, name: fontName };
//         }
//       }
//     });

//     row.getCell(totalCols).protection = { locked: false };
//     currentIndex++;
//   }

//   const lastDataRow = 10 + sortedStudents.length;

//   // ── 4. Unit statistics ─────────────────────────────────────────────────────
//   const statsStart  = lastDataRow + 2;
//   const statsLabels = [
//     "Mean", "Standard Deviation", "Maximum", "Minimum",
//     "No. of Candidates", "No. of Passes", "No. of Fails", "No. of Blanks",
//   ];

//   statsLabels.forEach((label, i) => {
//     const rIdx = statsStart + i;
//     const r    = sheet.getRow(rIdx);
//     r.height   = 15;

//     const labelCell = r.getCell(3);
//     labelCell.value  = label;
//     labelCell.font   = { bold: true, size: 7, name: fontName };
//     labelCell.border = thinBorder;

//     offeredUnits.forEach((_, uIdx) => {
//       const colIdx    = 5 + uIdx;
//       const colLetter = sheet.getColumn(colIdx).letter;
//       const cell      = r.getCell(colIdx);
//       const range     = `${colLetter}11:${colLetter}${lastDataRow}`;

//       cell.border = thinBorder;
//       cell.numFmt = "0.0";
//       cell.font   = { size: 7, name: fontName };

//       if (label === "No. of Passes") {
//         cell.value  = { formula: `COUNTIF(${range}, ">=${passMark}")` };
//         cell.numFmt = "0";
//       } else if (label === "No. of Fails") {
//         cell.value  = { formula: `COUNTIFS(${range}, "<${passMark}", ${range}, "<>")` };
//         cell.numFmt = "0";
//       } else if (label === "No. of Blanks") {
//         cell.value  = { formula: `COUNTIF(${range}, "INC")` };
//         cell.numFmt = "0";
//       } else {
//         const funcMap: Record<string, string> = {
//           "Mean": "AVERAGE", "Standard Deviation": "STDEV.P",
//           "Maximum": "MAX", "Minimum": "MIN", "No. of Candidates": "COUNT",
//         };
//         const fn = funcMap[label];
//         if (fn) cell.value = { formula: `IFERROR(ROUND(${fn}(${range}), 1), 0)` };
//       }
//     });

//     r.getCell(3).border = { ...r.getCell(3).border, left: { style: "thick" } };
//     r.getCell(tuColIdx - 1).border = { ...r.getCell(tuColIdx - 1).border, right: { style: "thick" } };

//     if (i === 0) {
//       for (let c = 3; c < tuColIdx; c++)
//         r.getCell(c).border = { ...r.getCell(c).border, top: { style: "thick" } };
//     }
//     if (i === statsLabels.length - 1) {
//       for (let c = 3; c < tuColIdx; c++)
//         r.getCell(c).border = { ...r.getCell(c).border, bottom: { style: "thick" } };
//     }
//   });

//   // ── 5. Summary table ───────────────────────────────────────────────────────
//   const summaryStart      = lastDataRow + 12;
//   const summaryHeaderCell = sheet.getCell(`B${summaryStart}`);
//   summaryHeaderCell.value = "SUMMARY";
//   summaryHeaderCell.font  = { bold: true, size: 10, underline: true, name: fontName };

//   const summaryData: Record<string, number> = {
//     PASS: 0, SUPPLEMENTARY: 0, "REPEAT YEAR": 0, "STAY OUT": 0,
//     SPECIAL: 0, INCOMPLETE: 0, "ACADEMIC LEAVE": 0,
//     DEFERMENT: 0, "DEREGISTERED/DISC": 0,
//   };

//   sheet.getColumn(tuColIdx + 3).eachCell({ includeEmpty: false }, (cell, rowNum) => {
//     if (rowNum > 10 && rowNum <= lastDataRow) {
//       const txt = cell.value?.toString().toUpperCase() || "";
//       if      (txt === "PASS")                                summaryData.PASS++;
//       else if (txt.includes("SUPP"))                          summaryData.SUPPLEMENTARY++;
//       else if (txt.includes("REPEAT"))                        summaryData["REPEAT YEAR"]++;
//       else if (txt.includes("STAY OUT"))                      summaryData["STAY OUT"]++;
//       else if (txt.includes("SPEC"))                          summaryData.SPECIAL++;
//       else if (txt.includes("ACADEMIC LEAVE"))                summaryData["ACADEMIC LEAVE"]++;
//       else if (txt.includes("DEFERMENT"))                     summaryData.DEFERMENT++;
//       else if (txt.includes("INC"))                           summaryData.INCOMPLETE++;
//       else if (txt.includes("DEREG") || txt.includes("DISC")) summaryData["DEREGISTERED/DISC"]++;
//     }
//   });

//   const activeSummaryEntries = Object.entries(summaryData).filter(([, count]) => count > 0);
//   activeSummaryEntries.forEach(([label, count], i) => {
//     const rIdx      = summaryStart + 1 + i;
//     const labelCell = sheet.getCell(`B${rIdx}`);
//     const countCell = sheet.getCell(`C${rIdx}`);
//     labelCell.value  = label;
//     countCell.value  = count;
//     labelCell.border = thinBorder;
//     countCell.border = thinBorder;
//     labelCell.font   = { size: 8, name: fontName, bold: true };
//     countCell.font   = { size: 8, name: fontName };
//   });

//   // ── 6. Offered units table ─────────────────────────────────────────────────
//   const unitsStart = summaryStart + activeSummaryEntries.length + 4;
//   sheet.mergeCells(unitsStart, 2, unitsStart, 6);
//   sheet.getCell(unitsStart, 2).value = "LIST OF UNITS OFFERED";
//   sheet.getCell(unitsStart, 2).font  = { bold: true, underline: true };

//   const mid = Math.ceil(offeredUnits.length / 2);
//   for (let i = 0; i < mid; i++) {
//     const rIdx  = unitsStart + 2 + i;
//     const r     = sheet.getRow(rIdx);
//     const left  = offeredUnits[i];
//     const right = offeredUnits[mid + i];

//     r.getCell(2).value = i + 1;
//     r.getCell(3).value = left.code;
//     sheet.mergeCells(rIdx, 4, rIdx, 7);
//     r.getCell(4).value = left.name;

//     if (right) {
//       r.getCell(9).value  = mid + i + 1;
//       r.getCell(10).value = right.code;
//       sheet.mergeCells(rIdx, 11, rIdx, 14);
//       r.getCell(11).value = right.name;
//     }

//     [2, 3, 4, 9, 10, 11].forEach((col) => {
//       const cell = r.getCell(col);
//       cell.border = { ...thinBorder };
//       if (col === 2)                           cell.border.left   = { style: "thick" };
//       if (col === 11 || (!right && col === 4)) cell.border.right  = { style: "thick" };
//       if (i === 0)                             cell.border.top    = { style: "thick" };
//       if (i === mid - 1)                       cell.border.bottom = { style: "thick" };
//       cell.font = { size: 8 };
//     });
//   }

//   // ── 7. Main table thick borders ────────────────────────────────────────────
//   for (let i = startRow; i <= lastDataRow; i++) {
//     sheet.getCell(i, 1).border = { ...sheet.getCell(i, 1).border, left:  { style: "thick" } };
//     sheet.getCell(i, totalCols).border = { ...sheet.getCell(i, totalCols).border, right: { style: "thick" } };
//   }
//   sheet.getRow(startRow).eachCell((c) => (c.border = { ...c.border, top:    { style: "thick" } }));
//   sheet.getRow(lastDataRow).eachCell((c) => (c.border = { ...c.border, bottom: { style: "thick" } }));

//   // ── 8. Sheet formatting ────────────────────────────────────────────────────
//   sheet.getColumn(1).width = 4;
//   sheet.getColumn(2).width = 22;  // wider for qualifier suffix
//   sheet.getColumn(3).width = 25;
//   sheet.getColumn(4).width = 7;   // wider for "RP1C", "A/SO", "A/RA1" etc.
//   offeredUnits.forEach((_, i) => (sheet.getColumn(5 + i).width = 4.5));
//   sheet.getColumn(tuColIdx).width     = 5;
//   sheet.getColumn(tuColIdx + 1).width = 7;
//   sheet.getColumn(tuColIdx + 2).width = 7;
//   sheet.getColumn(tuColIdx + 3).width = 20;
//   sheet.getColumn(tuColIdx + 4).width = 20;

//   sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 10 }];
//   sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
//   sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

//   const result = await workbook.xlsx.writeBuffer();
//   return Buffer.from(result as ArrayBuffer);
// };