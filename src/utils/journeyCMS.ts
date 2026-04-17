
// serverside/src/utils/journeyCMS.ts

import * as ExcelJS from "exceljs";
import config from "../config/config";
import InstitutionSettings from "../models/InstitutionSettings";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import ProgramUnit from "../models/ProgramUnit";
import Program from "../models/Program";
import { getYearWeight } from "../utils/weightingRegistry";
import { buildDisplayRegNo } from "../utils/academicRules";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface JourneyCMSData {
  programId:     string;
  programName:   string;
  academicYear:  string;
  logoBuffer:    Buffer;
  institutionId: string;
}

interface UnitMark {
  code:    string;
  name:    string;
  mark:    number | string;   // number, "INC", or "65C" for specials
  grade:   string;
  attempt: string;
  status:  string;
}

interface StudentYearRecord {
  yearOfStudy:    number;
  academicYear:   string;
  isRepeatYear:   boolean;
  annualMean:     number;
  weight:         number;
  suppUnits:      string[];
  cfUnits:        string[];
  specialUnits:   string[];
  passedSupp:     string[];
  passedCF:       string[];
  failedUnits:    string[];
  unitMarks:      UnitMark[];
}

interface StudentJourneyRecord {
  regNo:           string;
  displayRegNo:    string;
  name:            string;
  entryType:       string;
  status:          string;
  qualifierSuffix: string;
  waa:             number;
  classification:  string;
  yearsData:       StudentYearRecord[];
  hurdleSummary:   string;
}

// ─── Constants ─────────────────────────────────────────────────────────────────

const FONT = "Arial";
const THIN: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }};
const DOUBLE_BOTTOM: Partial<ExcelJS.Borders> = { ...THIN, bottom: { style: "double" } };

// Subtle fills — matches existing CMS tone
const F = {
  hdr:       { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE0E0E0" } }, // light grey header
  subhdr:    { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF2F2F2" } }, // very light grey
  pass:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFD9EAD3" } }, // light green
  passed:    { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFFFFF" } }, // light green
  supp:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFF2CC" } }, // yellow
  fail:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFC7CE" } }, // red
  special:   { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE3D9F5" } }, // purple
  repeat:    { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFCE4D6" } }, // orange
  gold:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFD700" } },
  blue:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFBDD7EE" } },
  amber:     { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFCE4D6" } },
  none:      { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFFFFF" } },
  inc:       { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFCE4D6" } }, // light red for INC
};

// ─── Helpers ──────────────────────────────────────────────────────────────────

// ─── Logo Helper — same centering formula as consolidatedMS.ts ───────────────
// Uses Math.floor(totalCols / 2) - 1 to match the consolidated mark sheet exactly.
function addLogo( wb: ExcelJS.Workbook, sheet: ExcelJS.Worksheet, logoBuffer: Buffer, totalCols: number) {
  if (!logoBuffer || logoBuffer.length === 0) return;

  const imageId = wb.addImage({buffer: logoBuffer as any, extension: "png"});
  const centerColIdx = Math.floor(totalCols / 2);
  sheet.addImage(imageId, {tl: { col: centerColIdx - 1, row: 0 }, ext: { width: 100, height: 60 }});
}

// ─── Centered Header Row ────────────────────────────────────────────────────
function hdrRow(
  sheet: ExcelJS.Worksheet,
  rowNum: number,
  text: string,
  totalCols: number,
  size = 10,
  bold = true,
  underline = false,
) {
  sheet.mergeCells(rowNum, 1, rowNum, totalCols);
  const cell = sheet.getCell(rowNum, 1);
  cell.value = text.toUpperCase();
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.font = { bold, name: FONT, size, underline };
  sheet.getRow(rowNum).height = size + 8;
}

function classLabel(waa: number): string {
  if (waa >= 70) return "FIRST CLASS HONOURS";
  if (waa >= 60) return "SECOND CLASS (UPPER DIVISION)";
  if (waa >= 50) return "SECOND CLASS (LOWER DIVISION)";
  if (waa >= 40) return "PASS";
  return "FAIL";
}

function classFill(waa: number) {
  if (waa >= 70) return F.gold;
  if (waa >= 60) return F.blue;
  if (waa >= 50) return F.amber;
  if (waa >= 40) return F.pass;
  return F.fail;
}

function meanFill(mean: number) {
  if (mean >= 70) return F.gold;
  if (mean >= 60) return F.blue;
  if (mean >= 50) return F.amber;
  if (mean >= 40) return F.pass;
  return F.fail;
}

function hurdleSummary(yearsData: StudentYearRecord[]): string {
  const events: string[] = [];
  for (const yr of yearsData) {
    const lbl = `Y${yr.yearOfStudy}`;
    if (yr.isRepeatYear)        events.push(`${lbl}:REPEAT`);
    if (yr.suppUnits.length)    events.push(`${lbl}:SUPP(${yr.suppUnits.join(",")})`);
    if (yr.passedSupp.length)   events.push(`${lbl}:✓SUPP(${yr.passedSupp.join(",")})`);
    if (yr.cfUnits.length)      events.push(`${lbl}:CF→Y${yr.yearOfStudy+1}(${yr.cfUnits.join(",")})`);
    if (yr.passedCF.length)     events.push(`${lbl}:✓CF(${yr.passedCF.join(",")})`);
    if (yr.specialUnits.length) events.push(`${lbl}:SPEC(${yr.specialUnits.join(",")})`);
  }
  return events.length ? events.join(" | ") : "CLEAN";
}

// ─── Mark resolution — reads FinalGrade first, then MarkDirect, then Mark ─────
// This is the KEY FIX for INC marks: direct-entry marks live in MarkDirect,
// not in FinalGrade. We resolve all three sources.

interface ResolvedMark {
  agreedMark:  number;
  caTotal30:   number | null;
  examTotal70: number | null;
  attempt:     string;
  isSpecial:   boolean;
  isSupp:      boolean;
  isRetake:    boolean;
  status:      string;   // "PASS" | "SUPPLEMENTARY" | "SPECIAL" | "FAIL"
  source:      "finalGrade" | "direct" | "detailed";
}

async function resolveMark(
  studentId: string,
  puId: string,
  passMark: number,
): Promise<ResolvedMark | null> {

  // 1. FinalGrade (set by computeFinalGrade after detailed mark upload or grade calc)
  const fgs = await FinalGrade.find({student: studentId, programUnit: puId}).lean() as any[];

  if (fgs.length > 0) {
    // Pick best: PASS > SPECIAL > SUPPLEMENTARY > others
    const rank = (s: string) => s === "PASS" ? 3 : s === "SPECIAL" ? 2 : s === "SUPPLEMENTARY" ? 1 : 0;
    const best = fgs.sort((a: any, b: any) => rank(b.status) - rank(a.status))[0];
    return {
      agreedMark:  best.totalMark ?? 0,
      caTotal30:   best.caTotal30   ?? null,
      examTotal70: best.examTotal70 ?? null,
      attempt:     best.attemptType === "SUPPLEMENTARY" ? "supplementary" :
                   best.attemptType === "RETAKE"        ? "re-take"       : "1st",
      isSpecial:   best.isSpecial === true || best.status === "SPECIAL",
      isSupp:      best.status === "SUPPLEMENTARY",
      isRetake:    best.attemptType === "RETAKE",
      status:      best.status,
      source:      "finalGrade",
    };
  }

  // 2. MarkDirect (direct CA+Exam entry — no FinalGrade record created)
  const md = await MarkDirect.findOne({ student: studentId, programUnit: puId }).lean() as any;

  if (md) {
    const mark    = md.agreedMark ?? 0;
    const isSpec  = md.isSpecial === true || md.attempt === "special";
    const isSupp  = md.attempt === "supplementary";
    const isPassed = mark >= passMark && !isSpec;
    return {
      agreedMark:  mark,
      caTotal30:   md.caTotal30   ?? null,
      examTotal70: md.examTotal70 ?? null,
      attempt:     md.attempt || "1st",
      isSpecial:   isSpec,
      isSupp,
      isRetake:    md.isRetake === true,
      status:      isSpec ? "SPECIAL" : isPassed ? "PASS" : "SUPPLEMENTARY",
      source:      "direct",
    };
  }

  // 3. Mark (detailed breakdown — agreedMark computed by gradeCalculator)
  const dm = await Mark.findOne({ student: studentId, programUnit: puId }).lean() as any;

  if (dm) {
    const mark   = dm.agreedMark ?? 0;
    const isSpec = dm.isSpecial === true || dm.attempt === "special";
    const isSupp = dm.attempt === "supplementary";
    const isPassed = mark >= passMark && !isSpec;
    return {
      agreedMark:  mark,
      caTotal30:   dm.caTotal30   ?? null,
      examTotal70: dm.examTotal70 ?? null,
      attempt:     dm.attempt || "1st",
      isSpecial:   isSpec,
      isSupp,
      isRetake:    dm.isRetake === true,
      status:      isSpec ? "SPECIAL" : isPassed ? "PASS" : "SUPPLEMENTARY",
      source:      "detailed",
    };
  }

  return null; // genuinely no mark for this unit
}

// ─── Load all student journey data ────────────────────────────────────────────

async function loadJourneys(
  programId:    string,
  program:      any,
  passMark:     number,
  gradingScale: Array<{ min: number; grade: string }>,
): Promise<StudentJourneyRecord[]> {
  const duration   = program.durationYears || 5;
  const sortedScale = [...gradingScale].sort((a, b) => b.min - a.min);
  const gradeFor   = (mark: number) =>
    mark < passMark ? "E" : (sortedScale.find(s => mark >= s.min)?.grade ?? "E");

  const students = await Student.find({
    program: programId,
    status:  { $in: ["graduand","graduated","active","repeat",
                      "on_leave","deferred","discontinued","deregistered"] },
  }).populate("program").lean() as any[];

  const records: StudentJourneyRecord[] = [];

  for (const student of students) {
    const sid     = student._id.toString();
    const et      = student.entryType || "Direct";
    const history = (student.academicHistory || []) as any[];
    const yearsData: StudentYearRecord[] = [];

    for (let y = 1; y <= duration; y++) {
      const hist = history.find((h: any) => h.yearOfStudy === y);
      if (!hist) continue; // student hasn't completed this year

      const weight    = getYearWeight(program, et, y);
      // Use stored annualMeanMark — set correctly by the /admin/recompute-graduation-waa migration
      const annualMean = hist.annualMeanMark || 0;

      const pus = await ProgramUnit.find({ program: programId, requiredYear: y }).populate("unit").lean() as any[];

      const suppUnits:    string[] = [];
      const cfUnits:      string[] = [];
      const specialUnits: string[] = [];
      const passedSupp:   string[] = [];
      const passedCF:     string[] = [];
      const failedUnits:  string[] = [];
      const unitMarks:    UnitMark[] = [];

      for (const pu of pus) {
        const code  = (pu as any).unit?.code || "N/A";
        const uname = (pu as any).unit?.name || "N/A";
        const puId  = (pu as any)._id.toString();

        // ── RESOLVE MARK FROM ALL THREE SOURCES ──────────────────────────
        const resolved = await resolveMark(sid, puId, passMark);

        if (!resolved) {
          unitMarks.push({ code, name: uname, mark: "INC", grade: "—", attempt: "—", status: "MISSING" });
          continue;
        }

        const { agreedMark, isSpecial, isSupp, isRetake, status, attempt } = resolved;

        // Attempt label for display
        let attemptLabel = "B/S";
        if (isSpecial)      { attemptLabel = "SPEC";  specialUnits.push(code); }
        else if (isRetake)  { attemptLabel = "RP1C";
          if (status === "PASS") passedCF.push(code); else cfUnits.push(code);
        }
        else if (isSupp)    { attemptLabel = "A/S";
          if (status === "PASS") passedSupp.push(code); else suppUnits.push(code);
        }
        else if (status !== "PASS" && !isSpecial) { failedUnits.push(code); }

        const displayMark: number | string = isSpecial ? `${agreedMark}C` : agreedMark;

        unitMarks.push({
          code,
          name:    uname,
          mark:    displayMark,
          grade:   isSpecial ? "I" : gradeFor(agreedMark),
          attempt: attemptLabel,
          status,
        });
      }

      yearsData.push({
        yearOfStudy:  y,
        academicYear: hist.academicYear || "",
        isRepeatYear: hist.isRepeatYear || false,
        annualMean,
        weight,
        suppUnits,
        cfUnits,
        specialUnits,
        passedSupp,
        passedCF,
        failedUnits,
        unitMarks,
      });
    }

    // ── WAA: use finalWeightedAverage when available; recompute otherwise ──
    let waa: number;
    const storedWAA = parseFloat(student.finalWeightedAverage || "0");
    if (storedWAA > 0) {
      waa = storedWAA;
    } else {
      const histWC = (student.academicHistory || []).reduce(
        (sum: number, h: any) => sum + (h.weightedContribution || 0),
        0,
      );
      if (histWC > 0) {
        waa = histWC;
      } else {
        let recomputedWAA = 0;
        for (const yr of yearsData) {
          if (yr.unitMarks.length === 0) continue;
          recomputedWAA += yr.annualMean * yr.weight;
        }
        waa = recomputedWAA;
      }
    }

    const classification = student.classification || classLabel(waa);

    records.push({
      regNo:           student.regNo,
      displayRegNo:    buildDisplayRegNo(student.regNo, student.qualifierSuffix),
      name:            student.name,
      entryType:       et,
      status:          student.status,
      qualifierSuffix: student.qualifierSuffix || "",
      waa:             parseFloat(waa.toFixed(2)),
      classification,
      yearsData,
      hurdleSummary:   hurdleSummary(yearsData),
    });
  }

  // Sort same as award list: best class first, then WAA desc
  const ORDER = [
    "FIRST CLASS HONOURS",
    "SECOND CLASS HONOURS (UPPER DIVISION)",
    "SECOND CLASS HONOURS (LOWER DIVISION)",
    "PASS","FAIL",
  ];
  records.sort((a, b) => {
    const ai = ORDER.indexOf(a.classification);
    const bi = ORDER.indexOf(b.classification);
    if (ai !== bi) return ai - bi;
    return b.waa - a.waa;
  });

  return records;
}

// ─── SHEET: SUMMARY ───────────────────────────────────────────────────────────

function buildSummary(
  wb:          ExcelJS.Workbook,
  records:     StudentJourneyRecord[],
  program:     any,
  logoBuffer:  Buffer,
  academicYear: string,
) {
  const duration   = program.durationYears || 5;
  const TOTAL_COLS = 5 + duration + 3; // S/N + Reg + Name + Entry + Y1..YN + WAA + Class + Hurdle Log + Remarks

  const sheet = wb.addWorksheet("SUMMARY");
  addLogo(wb, sheet, logoBuffer, TOTAL_COLS);  

  // Header rows — plain, no dark fill
  hdrRow(sheet, 4, config.instName, TOTAL_COLS, 12, true, false);
  hdrRow(sheet, 5, config.schoolName || "SCHOOL OF ENGINEERING", TOTAL_COLS, 10, true, false);
  hdrRow(sheet, 6, program.name, TOTAL_COLS, 10, true, false);
  hdrRow(sheet, 7, `STUDENT ACADEMIC JOURNEY SUMMARY — ${academicYear} ACADEMIC YEAR`, TOTAL_COLS, 9, true, true);
  hdrRow(sheet, 8, "BOARD OF EXAMINERS REPORT", TOTAL_COLS, 9, false, false);

  const HDR = 10;
  sheet.getRow(HDR).height = 52;

  const hdrs = ["S/N", "REG. NO.", "NAME", "ENTRY", ...Array.from({ length: duration }, (_, i) => `Y${i+1}\n(%)`), "WAA\n(%)", "CLASSIFICATION", "HURDLE LOG", "BOARD\nREMARKS"];

  hdrs.forEach((txt, i) => {
    const cell  = sheet.getRow(HDR).getCell(i + 1);
    cell.value     = txt;
    cell.fill      = F.hdr;
    cell.border    = DOUBLE_BOTTOM;
    cell.font      = { bold: true, name: FONT, size: 8 };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  let row = HDR + 1;
  let ctr = 1;

  for (const rec of records) {
    const r = sheet.getRow(row);
    r.height = 14;

    const vals: any[] = [ctr++, rec.displayRegNo, rec.name.toUpperCase(), rec.entryType];

    for (let y = 1; y <= duration; y++) {
      const yr = rec.yearsData.find(d => d.yearOfStudy === y);
      vals.push(yr ? parseFloat(yr.annualMean.toFixed(2)) : "—");
    }

    const waaCol = 5 + duration;
    vals.push(rec.waa, rec.classification, rec.hurdleSummary, "");
    r.values = vals;

    r.eachCell((cell, colNum) => {
      cell.border    = THIN;
      cell.font      = { name: FONT, size: 8 };
      cell.alignment = { vertical: "middle", horizontal: "center" };

      if (colNum === 2 || colNum === 3) cell.alignment = { horizontal: "left", vertical: "middle" };

      // Year mean cells colour
      if (colNum >= 5 && colNum < waaCol) {
        const v = cell.value as number;
        if (typeof v === "number") cell.fill = meanFill(v);
      }

      // WAA
      if (colNum === waaCol) {
        cell.fill = classFill(rec.waa);
        cell.font = { bold: true, name: FONT, size: 8 };
        cell.numFmt = "0.00";
      }

      // Classification
      if (colNum === waaCol + 1) {
        cell.fill = classFill(rec.waa);
        cell.font = { bold: true, name: FONT, size: 8 };
      }

      // Hurdle log
      if (colNum === waaCol + 2) {
        cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        const isClean = (cell.value?.toString() || "") === "CLEAN";
        cell.fill = isClean ? F.pass : F.supp;
        cell.font = { name: FONT, size: 7, italic: !isClean };
      }

      // Board remarks — unlocked
      if (colNum === waaCol + 3) {
        cell.protection = { locked: false };
        cell.fill = F.none;
      }
    });

    row++;
  }

  const lastData = row - 1;

  // Thick outer borders
  for (let r = HDR; r <= lastData; r++) {
    sheet.getCell(r, 1).border        = { ...sheet.getCell(r,1).border,        left:   { style:"thick" } };
    sheet.getCell(r, TOTAL_COLS).border = { ...sheet.getCell(r,TOTAL_COLS).border, right:  { style:"thick" } };
  }
  sheet.getRow(HDR).eachCell(c => (c.border = { ...c.border, top: { style:"thick" } }));
  sheet.getRow(lastData).eachCell(c => (c.border = { ...c.border, bottom: { style:"thick" } }));

  // Column widths
  sheet.getColumn(1).width = 4;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 8;
  for (let y = 1; y <= duration; y++) sheet.getColumn(4 + y).width = 7;
  sheet.getColumn(5 + duration).width = 8;
  sheet.getColumn(6 + duration).width = 30;
  sheet.getColumn(7 + duration).width = 55;
  sheet.getColumn(8 + duration).width = 25;

  sheet.views = [{ state: "frozen", xSplit: 3, ySplit: HDR }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

// ─── SHEET PER YEAR ───────────────────────────────────────────────────────────

function buildYearSheet(
  wb:          ExcelJS.Workbook,
  y:           number,
  records:     StudentJourneyRecord[],
  program:     any,
  academicYear: string,
  passMark:    number,
  logoBuffer:  Buffer,
) {
  // Collect all unit codes for this year (union across all students)
  const unitMap = new Map<string, string>(); // code → name
  records.forEach((rec) => {
    const yr = rec.yearsData.find((d) => d.yearOfStudy === y);
    if (!yr) return;
    yr.unitMarks.forEach((u) => unitMap.set(u.code, u.name));
  });
  const codes = Array.from(unitMap.keys());
  const names = codes.map((c) => unitMap.get(c) || c);

  const YEAR_LABEL =
    ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][y - 1] || `${y}TH`;
  const TOTAL_COLS = 4 + codes.length + 3; // S/N + Reg + Name + Attempt + units + Mean + Recomm + Hurdle

  const sheet = wb.addWorksheet(`Year ${y}`.substring(0, 31));
  addLogo(wb, sheet, logoBuffer, TOTAL_COLS);

  hdrRow(sheet, 4, config.instName, TOTAL_COLS, 11, true);
  hdrRow(sheet, 5, program.name, TOTAL_COLS, 10, true);
  hdrRow(sheet, 6, `${YEAR_LABEL} YEAR — DETAILED MARKS & HURDLE REGISTRY — ${academicYear}`, TOTAL_COLS, 9, true, true );

  const HDR = 8;
  const CODE = 9;
  sheet.getRow(HDR).height = 15;
  sheet.getRow(CODE).height = 48;

  // Merge fixed columns across header + code rows
  ["A", "B", "C", "D"].forEach((col, i) => {
    sheet.mergeCells(`${col}${HDR}:${col}${CODE}`);
    const cell = sheet.getCell(`${col}${HDR}`);
    cell.value = ["S/N", "REG. NO.", "NAME", "ATTEMPT"][i];
    cell.fill = F.hdr;
    cell.border = DOUBLE_BOTTOM;
    cell.font = { bold: true, name: FONT, size: 8 };
    cell.alignment = { horizontal: "center", vertical: "middle", textRotation: i === 3 ? 90 : 0 };
  });

  // Unit header (number row) + code row
  codes.forEach((code, i) => {
    const colNum = 5 + i;
    const h1 = sheet.getRow(HDR).getCell(colNum);
    h1.value = (i + 1).toString();
    h1.fill = F.hdr;
    h1.border = THIN;
    h1.font = { bold: true, name: FONT, size: 8 };
    h1.alignment = { horizontal: "center", vertical: "middle" };

    const h2 = sheet.getRow(CODE).getCell(colNum);
    h2.value = code;
    h2.fill = F.hdr;
    h2.border = DOUBLE_BOTTOM;
    h2.font = { bold: true, name: FONT, size: 7 };
    h2.alignment = { horizontal: "center", vertical: "middle", textRotation: 90 };
  });

  const MEAN_COL = 5 + codes.length;
  const RECOMM_COL = MEAN_COL + 1;
  const LOG_COL = RECOMM_COL + 1;

  [
    [MEAN_COL, "MEAN (%)"],
    [RECOMM_COL, "RECOMM."],
    [LOG_COL, "HURDLE LOG"],
  ].forEach(([col, lbl]) => {
    sheet.mergeCells(HDR, col as number, CODE, col as number);
    const cell = sheet.getRow(HDR).getCell(col as number);
    cell.value = lbl;
    cell.fill = F.hdr;
    cell.border = DOUBLE_BOTTOM;
    cell.font = { bold: true, name: FONT, size: 8 };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  // Data rows
  let rowIdx = CODE + 1;
  let ctr = 1;

  for (const rec of records) {
    const yr = rec.yearsData.find((d) => d.yearOfStudy === y);
    if (!yr) continue;

    const r = sheet.getRow(rowIdx);
    r.height = 14;

    // Hurdle log
    const log: string[] = [];
    if (yr.isRepeatYear) log.push("REPEAT YEAR");
    if (yr.suppUnits.length) log.push(`SUPP: ${yr.suppUnits.join(", ")}`);
    if (yr.passedSupp.length)
      log.push(`✓ SUPP CLEARED: ${yr.passedSupp.join(", ")}`);
    if (yr.cfUnits.length) log.push(`CF→Y${y + 1}: ${yr.cfUnits.join(", ")}`);
    if (yr.passedCF.length) log.push(`✓ CF CLEARED: ${yr.passedCF.join(", ")}`);
    if (yr.specialUnits.length)
      log.push(`SPECIAL: ${yr.specialUnits.join(", ")}`);
    const hLog = log.join(" | ") || "CLEAN";

    // Attempt label
    let attempt = "B/S";
    if (yr.isRepeatYear) attempt = "A/RA1";
    else if (rec.qualifierSuffix.includes("C")) attempt = rec.qualifierSuffix;

    // Recomm
    let recomm =
      yr.failedUnits.length === 0 && !yr.isRepeatYear
        ? "PASS"
        : yr.isRepeatYear
          ? "REPEAT YEAR"
          : yr.suppUnits.length
            ? `SUPP (${yr.suppUnits.length})`
            : yr.cfUnits.length
              ? `CF (${yr.cfUnits.length})` 
              : "INC";

    const vals: any[] = [ ctr++, rec.displayRegNo, rec.name.toUpperCase(), attempt ];

    codes.forEach((code) => {
      const um = yr.unitMarks.find((u) => u.code === code);
      vals.push(um ? um.mark : "—");
    });

    vals.push(parseFloat(yr.annualMean.toFixed(2)), recomm, hLog);
    r.values = vals;

    // Style
    r.eachCell((cell, colNum) => {
      cell.border = THIN;
      cell.font = { name: FONT, size: 8 };
      cell.alignment = { vertical: "middle", horizontal: "center" };

      if (colNum === 2 || colNum === 3)
        cell.alignment = { horizontal: "left", vertical: "middle" };

      // Unit mark cells
      if (colNum >= 5 && colNum < MEAN_COL) {
        const v = cell.value;
        const um = yr.unitMarks.find((u) => u.code === codes[colNum - 5]);
        if (v === "—" || v === undefined) {
          cell.fill = F.none;
        } else if (typeof v === "string" && (v as string).endsWith("C")) {
          cell.fill = F.special;
        } else if (v === "INC") {
          cell.fill = F.inc;
          cell.font = { bold: true, name: FONT, size: 8 };
        } else if (typeof v === "number") {
          if (um?.status === "SUPPLEMENTARY" && v >= passMark) {
            cell.fill = F.supp;
            cell.font = { italic: true, name: FONT, size: 8 };
          } else if (um?.status === "RETAKE" && v >= passMark) {
            cell.fill = F.amber;
            cell.font = { italic: true, name: FONT, size: 8 };
          } else if (v >= passMark) {
            cell.fill = F.passed;
          } else {
            cell.fill = F.fail;
            cell.font = { bold: true, name: FONT, size: 8, color: { argb: "FF9C0006" }};
          }
        }
      }

      if (colNum === MEAN_COL) {
        cell.fill = meanFill(cell.value as number);
        cell.font = { bold: true, name: FONT, size: 8 };
        cell.numFmt = "0.00";
      }

      if (colNum === RECOMM_COL) {
        const txt = (cell.value?.toString() || "").toUpperCase();
        cell.fill =
          txt === "PASS" ? F.pass
            : txt.includes("SUPP") ? F.supp
              : txt.includes("REPEAT") ? F.repeat 
                : F.fail;
        cell.font = { bold: true, name: FONT, size: 8 };
      }

      if (colNum === LOG_COL) {
        const txt = cell.value?.toString() || "";
        cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        cell.font = { name: FONT, size: 7, italic: txt !== "CLEAN" };
        cell.fill = txt !== "CLEAN" ? F.supp : F.pass;
      }
    });

    rowIdx++;
  }

  const lastData = rowIdx - 1;

  // Thick borders
  for (let r = HDR; r <= lastData; r++) {
    sheet.getCell(r, 1).border = { ...sheet.getCell(r, 1).border, left: { style: "thick" }};
    sheet.getCell(r, TOTAL_COLS).border = { ...sheet.getCell(r, TOTAL_COLS).border, right: { style: "thick" }};
  }
  sheet.getRow(HDR).eachCell((c) => (c.border = { ...c.border, top: { style: "thick" } }));
  sheet.getRow(lastData).eachCell((c) => (c.border = { ...c.border, bottom: { style: "thick" } }));

  // Unit code reference table
  const REF = lastData + 2;
  hdrRow(sheet, REF, "UNITS OFFERED THIS YEAR", TOTAL_COLS, 9, true, true);

  const mid = Math.ceil(codes.length / 2);
  for (let i = 0; i < mid; i++) {
    const rIdx = REF + 2 + i;
    const r = sheet.getRow(rIdx);
    r.getCell(1).value = i + 1;
    r.getCell(2).value = codes[i];
    sheet.mergeCells(rIdx, 3, rIdx, 6);
    r.getCell(3).value = names[i];
    const right = codes[mid + i];
    if (right) {
      r.getCell(8).value = mid + i + 1;
      r.getCell(9).value = right;
      sheet.mergeCells(rIdx, 10, rIdx, 13);
      r.getCell(10).value = names[mid + i];
    }
    [1, 2, 3, 8, 9, 10].forEach((c) => {
      r.getCell(c).border = THIN;
      r.getCell(c).font = { name: FONT, size: 8 };
    });
  }

  // Column widths
  sheet.getColumn(1).width = 4;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 7;
  codes.forEach((_, i) => (sheet.getColumn(5 + i).width = 6));
  sheet.getColumn(MEAN_COL).width = 8;
  sheet.getColumn(RECOMM_COL).width = 14;
  sheet.getColumn(LOG_COL).width = 50;

  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: CODE }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

// ─── LEGEND sheet ─────────────────────────────────────────────────────────────

function buildLegend(wb: ExcelJS.Workbook, logoBuffer: Buffer) {
  const sheet = wb.addWorksheet("LEGEND");
  addLogo(wb, sheet, logoBuffer, 5);

  sheet.getColumn(1).width = 18;
  sheet.getColumn(2).width = 55;

  hdrRow(sheet, 4, "SYMBOL & COLOUR LEGEND", 2, 11, true, true);

  const items: Array<[string, string, any]> = [
    ["B/S",    "First ordinary sitting",                            F.pass],
    ["A/S",    "Supplementary examination (ENG.13)",                F.supp],
    ["RP1C",   "Carry forward — 1st cycle (ENG.14)",               F.amber],
    ["RP2C",   "Carry forward — 2nd cycle (ENG.14)",               F.amber],
    ["A/RA1",  "After Repeat Year — 1st repeat (ENG.16)",          F.repeat],
    ["A/SO",   "Stayout retake in ordinary period (ENG.15h)",       F.supp],
    ["SPEC",   "Special examination (ENG.18)",                      F.special],
    ["INC",    "Incomplete — missing mark in DB",                   F.inc],
    ["—",      "Unit not taken / not in curriculum",               F.none],
    ["CLEAN",  "No hurdles — all units passed at first attempt",   F.pass],
    ["✓ SUPP", "Unit passed at supplementary",                      F.supp],
    ["✓ CF",   "Carry-forward unit subsequently passed",           F.amber],
  ];

  sheet.getRow(6).values = ["CODE / FILL", "MEANING"];
  sheet.getRow(6).eachCell(c => { c.font = { bold: true, name: FONT, size: 9 }; c.border = THIN; c.fill = F.hdr; });

  items.forEach(([code, meaning, fill], i) => {
    const r = sheet.getRow(7 + i);
    r.height = 16;
    r.getCell(1).value = code;   r.getCell(1).fill = fill;   r.getCell(1).border = THIN; r.getCell(1).font = { bold: true, name: FONT, size: 9 }; r.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    r.getCell(2).value = meaning; r.getCell(2).border = THIN; r.getCell(2).font = { name: FONT, size: 9 };
  });

  const ruleStart = 7 + items.length + 2;
  hdrRow(sheet, ruleStart, "KEY ENG RULES", 2, 9, true, true);
  [
    ["ENG.13","Supplementary — fail ≤ 1/3 units"],
    ["ENG.14","Carry Forward — max 2 units, proceed to next year"],
    ["ENG.15h","Stayout — fail >1/3 <1/2, retake in next ordinary"],
    ["ENG.16","Repeat Year — fail ≥ 1/2 or mean < 40%"],
    ["ENG.18","Special Exams — financial / compassionate / sickness"],
    ["ENG.22","Discontinuation — 5-attempt ladder"],
    ["ENG.25","WAA classification across all years"],
  ].forEach(([rule, desc], i) => {
    const r = sheet.getRow(ruleStart + 1 + i);
    r.height = 13;
    r.getCell(1).value = rule; r.getCell(1).font = { bold: true, name: FONT, size: 8 };
    r.getCell(2).value = desc; r.getCell(2).font = { name: FONT, size: 8 };
    [1,2].forEach(c => (r.getCell(c).border = THIN));
  });
}

// ─── MAIN EXPORT ──────────────────────────────────────────────────────────────

export const generateJourneyCMS = async (data: JourneyCMSData): Promise<Buffer> => {
  const { programId, programName, academicYear, logoBuffer, institutionId } = data;

  const [programDoc, settings] = await Promise.all([
    Program.findById(programId).lean() as Promise<any>,
    InstitutionSettings.findOne({ institution: institutionId }).lean() as Promise<any>,
  ]);

  if (!programDoc) throw new Error(`Program ${programId} not found`);

  const passMark    = settings?.passMark    ?? 40;
  const gradingScale = settings?.gradingScale ?? [];
  const duration    = programDoc.durationYears || 5;

  const records = await loadJourneys(programId, programDoc, passMark, gradingScale);

  if (records.length === 0) {
    throw new Error("No students found for this program.");
  }

  const wb = new ExcelJS.Workbook();

  buildSummary(wb, records, programDoc, logoBuffer, academicYear);

  for (let y = 1; y <= duration; y++) {
    if (records.some(r => r.yearsData.some(yd => yd.yearOfStudy === y))) {
      buildYearSheet(wb, y, records, programDoc, academicYear, passMark, logoBuffer);
    }
  }

  buildLegend(wb, logoBuffer);

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf as ArrayBuffer);
};