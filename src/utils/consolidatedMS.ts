// serverside/src/utils/consolidatedMS.ts
import * as ExcelJS from "exceljs";
import config from "../config/config";
import InstitutionSettings from "../models/InstitutionSettings";
import { resolveStudentStatus } from "./studentStatusResolver";
import { calculateStudentStatus } from "../services/statusEngine";
import { buildDisplayRegNo, getAttemptLabel } from "./academicRules";
import mongoose from "mongoose";

interface OfferedUnit { code: string; name: string }

interface StudentRow {
    sId: string; regNo: string; name: string; status: string;
    academicHistory?: Array<{ yearOfStudy: number; isRepeatYear?: boolean; academicYear?: string }>;
    academicLeavePeriod?: { type?: string }; marks: Map<string, number | "INC" | "C">; // unitCode → value
    attempt: string; totalUnits: number; total: number | "-"; mean: number | "-"; recomm: string; matters: string;
  }

export interface ConsolidatedData {
    programName: string; programId: string; academicYear: string; yearOfStudy: number;
    session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED"; students: Array<Record<string, unknown>>;
    marks: Array<Record<string, unknown>>; offeredUnits: OfferedUnit[]; logoBuffer: any;
    institutionId: string; passMark: number; gradingScale:  Array<{ min: number; grade: string }>;
  }

export const generateConsolidatedMarkSheet = async ( data: ConsolidatedData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, students, marks, offeredUnits, logoBuffer, institutionId, programId } = data;

  const settings = await InstitutionSettings.findOne({ institution: institutionId });
  const passMark = settings?.passMark || 40;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("CONSOLIDATED MARKSHEET");
  const fontName = "Arial";

  const tuColIdx = 5 + offeredUnits.length;
  const totalCols = tuColIdx + 4;
  const thinBorder: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }};
  const doubleBottomBorder: Partial<ExcelJS.Borders> = { ...thinBorder, bottom: { style: "double" } };

  // 1. HEADERS (Rows 4-7)
  const centerColIdx = Math.floor(totalCols / 2);
  if (logoBuffer && logoBuffer.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, { tl: { col: centerColIdx - 1, row: 0 }, ext: { width: 100, height: 60 }});
  }

  const setCenteredHeader = ( rowNum: number, text: string, fontSize: number = 10 ) => {
    sheet.mergeCells(rowNum, 1, rowNum, totalCols);
    const cell = sheet.getCell(rowNum, 1);
    cell.value = text.toUpperCase();
    cell.style = { alignment: { horizontal: "center", vertical: "middle" }, font: { bold: true, name: fontName, size: fontSize -1 }};
  };

  const examPhaseLabel =
    data.session === "SUPPLEMENTARY" ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS" : "ORDINARY EXAMINATION RESULTS";

  const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
  setCenteredHeader(4, `${config.instName}`);
  setCenteredHeader(5, `${config.schoolName || "SCHOOL OF ENGINEERING"}`);
  setCenteredHeader(6, `${programName}`);
  setCenteredHeader(7, `CONSOLIDATED MARK SHEET - - ${examPhaseLabel} - ${yrTxt} YEAR - ${academicYear} ACADEMIC YEAR`);
  sheet.getCell(7, 1).font.underline = true;

  // 2. TABLE HEADERS (Row 9-10)
  const startRow = 9;
  const subRow = 10;
  sheet.getRow(subRow).height = 48;

  const headers: { [key: number]: string } = {
    1: "S/N", 2: "REG. NO", 3: "NAME", 4: "ATTEMPT", [tuColIdx]: "T U", [tuColIdx + 1]: "TOTAL",
     [tuColIdx + 2]: "MEAN", [tuColIdx + 3]: "RECOMM.", [tuColIdx + 4]: "STUDENT MATTERS",
  };

  Object.entries(headers).forEach(([col, text]) => {
    const colNum = parseInt(col);
    sheet.mergeCells(startRow, colNum, subRow, colNum);
    const cell = sheet.getCell(startRow, colNum);
    cell.value = text;
    cell.style = {
      alignment: { horizontal: "center", vertical: "middle", textRotation: colNum === 4 ? 90 : 0, wrapText: true },
      font: { bold: true, size: 7, name: fontName },
      border: doubleBottomBorder,
    };
  });

  offeredUnits.forEach((unit, i) => {
    const colIdx = 5 + i;
    sheet.getCell(startRow, colIdx).value = (i + 1).toString();
    sheet.getCell(startRow, colIdx).style = { alignment: { horizontal: "center", vertical: "middle" }, font: { bold: true, size: 7, name: fontName }, border: thinBorder };
    sheet.getCell(subRow, colIdx).value = unit.code;
    sheet.getCell(subRow, colIdx).style = { alignment: { horizontal: "center", vertical: "middle", textRotation: 90 }, font: { bold: true, size: 7, name: fontName }, border: thinBorder };
  });

  // 3. STUDENT DATA (Row 11+)  
  const sortedStudents = [...students].sort((a, b) => (String(a.regNo || "")).localeCompare(String(b.regNo || "")));
  
//   let currentIndex = 0;
//   for (const student of sortedStudents) {
//     // const rIdx = 11 + currentIndex;
//     // const sId = student.id?.toString() || student._id?.toString();    
//     // const audit = await calculateStudentStatus( sId, programId, academicYear, yearOfStudy, { forPromotion: true } );
//     const rIdx = 11 + currentIndex;
//     const sId = student.id?.toString() || student._id?.toString();

//     // ── ADMIN STATUS GATE ─────────────────────────────────────────────
//     // Students on leave, deferred, discontinued, or deregistered have
//     // zero marks by design. Running the engine on them produces false
//     // REPEAT YEAR or DEREGISTERED. Return their DB status directly.
//     const ADMIN_STATUS_MAP: Record<string, string> = {
//       on_leave: "ACADEMIC LEAVE",
//       deferred: "DEFERMENT",
//       discontinued: "DISCONTINUED",
//       deregistered: "DEREGISTERED",
//       graduated: "GRADUATED",
//     };

//     const adminStatus = ADMIN_STATUS_MAP[student.status];
//     let audit: any;

//     if (adminStatus) {
//       // Synthesise a minimal audit object — no engine call
//       audit = {
//         status: adminStatus,
//         variant: "info",
//         details: adminStatus,
//         weightedMean: "0.00",
//         passedList: [],
//         failedList: [],
//         specialList: [],
//         missingList: [],
//         incompleteList: [],
//         summary: {
//           totalExpected: offeredUnits.length,
//           passed: 0,
//           failed: 0,
//           missing: 0,
//           isOnLeave: true,
//         },
//       };
//     } else {
//       // Active students — run the real engine
//       audit = await calculateStudentStatus(
//         sId,
//         programId,
//         academicYear,
//         yearOfStudy,
//         { forPromotion: true },
//       );
//     }
//     const resolvedStatus = resolveStudentStatus(student);
//     // --- REINSTATED FLAG LOGIC ---
//     const hasReturnHistory = student.statusHistory?.some(
//       (h: any) => h.status === "ACTIVE" && (h.previousStatus === "ACADEMIC LEAVE" || h.previousStatus === "DEFERMENT"));

//     // --- REPEAT FLAG LOGIC (RPT) ---
//     // Count how many times they have a record for this specific year of study in their history
//     const repeatCount = student.academicHistory?.filter((h: any) => h.yearOfStudy === yearOfStudy && h.status === "REPEAT YEAR").length || 0;

//     // Determine the dynamic notation for Column 4
//     // const attemptNotation = getAttemptLabel(student.attemptNumber || 1, student.status, student.regNo);

//     // const sId = student.id?.toString() || student._id?.toString();
//     const studentMarks = marks.filter(
//       (m: any) => (m.student?._id?.toString() || m.student?.toString()) === sId,
//     );

//     const attemptNotation = (() => {
//       const st = (student.status || "").toLowerCase();

//       // Administrative overrides — these students have no ordinary marks
//       if (st === "deferred") return "DEFERRED";
//       if (st === "on_leave") return "A/SO";
//       if (st === "discontinued") return "DISC.";
//       if (st === "deregistered") return "DEREG.";

//       // Repeat year is a DB status — student is re-sitting all units
//       if (st === "repeat") return "A/RA1";

//       // Derive from marks — what type of attempt does this sitting represent?
//       const attemptTypes = studentMarks.map((m: any) =>
//         (m.attempt || "1st").toLowerCase(),
//       );

//       if (attemptTypes.every((a) => a === "1st" || a === "special")) {
//         // Check academic history to see if this is truly a first sitting
//         const hasRepeatHistory = (student.academicHistory || []).some(
//           (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
//         );
//         if (hasRepeatHistory) return "A/RA1";
//         return "B/S"; // genuine first sitting
//       }
//       if (attemptTypes.includes("re-take")) return "A/CF";
//       if (attemptTypes.includes("supplementary")) return "A/S";
//       return "B/S";
//     })();
    
//     const repeatTag = repeatCount > 0 ? ` (RPT${repeatCount})` : "";
//     const reinstatedTag = hasReturnHistory ? " (REINSTATED)" : "";

//     // Final Display Name: "John Doe (REINSTATED) (RPT1)"
//     const finalDisplayName = `${student.name}${reinstatedTag}${repeatTag}`.toUpperCase();

//     // const rowData: any[] = [ currentIndex + 1, student.regNo, finalDisplayName, "B/S" ];
//     const rowData: any[] = [ currentIndex + 1, student.regNo, finalDisplayName, attemptNotation ];

//     // const rowData: any[] = [currentIndex + 1, student.regNo, student.name, "B/S"];

//     // Fill Unit Marks
//     offeredUnits.forEach((unit) => {
//       const markObj = marks.find((m) => (m.student?._id?.toString() || m.student?.toString()) === sId && m.programUnit?.unit?.code === unit.code);

//       if (resolvedStatus.isLocked) rowData.push("");
//       else if (markObj) {
//         const isSpecial = markObj.isSpecial || markObj.remarks?.toLowerCase().includes("special");
//         const markValue = markObj.agreedMark ?? 0;
//         // const isMissingData = !markObj.caTotal30 || !markObj.examTotal70;

//         // Check for null/undefined instead of falsy 0
//         const hasCA = markObj.caTotal30 !== null && markObj.caTotal30 !== undefined;
//         const hasExam = markObj.examTotal70 !== null && markObj.examTotal70 !== undefined;

//         // In Direct Entry, if they have the markObj, they usually have the data.
//         // We only show INC if the fields are physically missing from the DB record.
//         const isMissingData = !hasCA || !hasExam;

//         if (isSpecial) rowData.push(`${markValue}C`);
//         else if (isMissingData || markValue === 0) rowData.push("INC");
//         else rowData.push(markValue);
//       } else rowData.push("INC");
//     });

//     // Formatting the Specific Recommendation Column
//     let recomm = audit.status;    
//     const isSpecialCase = audit.specialList.length > 0 || audit.failedList.length > 0 || audit.incompleteList.length > 0;

//     if ( isSpecialCase && !["REPEAT YEAR", "STAYOUT", "DEREGISTERED"].includes(audit.status)) {
//       const parts = [];
//       if (audit.failedList.length > 0) parts.push(`SUPP ${audit.failedList.length}`);
//       if (audit.specialList.length > 0) parts.push(`SPEC ${audit.specialList.length}`);
//       if (audit.incompleteList.length > 0) parts.push(`INC ${audit.incompleteList.length}`);
//       recomm = parts.length > 0 ? parts.join("; ") : audit.status;
//     }

//     const studentMattersList: string[] = [];

//     // 1. Check for Academic Leave Grounds
//     if (audit.summary.isOnLeave || ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)) {
//       // Priority 1: The structured leave period type
//       const leaveType = student.academicLeavePeriod?.type;
//       // Priority 2: The remarks field
//       const remarks = student.remarks?.toLowerCase() || "";
      
//       if (leaveType) studentMattersList.push(leaveType.toUpperCase());
//       else if (remarks.includes("financial")) studentMattersList.push("FINANCIAL");
//       else if (remarks.includes("compassionate") || remarks.includes("medical")) {
//           studentMattersList.push("COMPASSIONATE");
//       } else if (audit.leaveDetails) {
//           // Fallback to the reason string from resolveStudentStatus
//           const cleanReason = audit.leaveDetails.split(":").pop()?.trim().toUpperCase();
//           if (cleanReason && !cleanReason.includes("PENDING")) studentMattersList.push(cleanReason);
//       }
//   }

//     // 2. Check for Special Grounds (from engine/marks)
//     audit.specialList.forEach((spec: any) => {
//       if (spec.grounds) {
//         // This catches "Special Granted: Financial"
//         const cleanSpec = spec.grounds.split(":").pop()?.trim() || spec.grounds;
//         if (!["special", "reason pending"].includes(cleanSpec.toLowerCase())) {
//           studentMattersList.push(cleanSpec.toUpperCase());
//         }
//       }
//     });

//     const finalMatters = Array.from(new Set(studentMattersList)).join(", ");

//     // Totals and Recommendations from Engine
//     // const totalMarks = audit.passedList.reduce((a, b) => a + b.mark, 0) + (audit.failedList as any[]).reduce((a, b) => a + (b.mark || 0), 0);
//     const totalPassedMarks = audit.passedList.reduce((a: number, b: any) => a + b.mark, 0);
// const totalFailedMarks = (audit.failedList as any[]).reduce((a: number, b: any) => a + (b.mark || 0), 0);
// const totalMarks = totalPassedMarks + totalFailedMarks;

//     const isBlockedStatus = audit.summary.isOnLeave || ["ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED"].includes(audit.status);

//     // Prepare the display values
//     const displayTotal = isBlockedStatus ? "-" : totalMarks;
//     const displayMean = isBlockedStatus ? "-" : parseFloat(audit.weightedMean).toFixed(2);

//     rowData.push( audit.summary.totalExpected, displayTotal, displayMean, recomm, finalMatters );

//     const row = sheet.getRow(rIdx);
//     row.values = rowData;

//     // STYLING
//     row.eachCell((cell, colNum) => {
//       cell.border = thinBorder;
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.font = { size: 8, name: fontName };

//       // Highlight REINSTATED names in Blue
//       // if (colNum === 3 && (hasReturnHistory || repeatCount > 0)) cell.font = { color: { argb: "FF0000FF" }, bold: true, size: 8, name: fontName };

//       if (colNum === totalCols - 1) {
//         cell.protection = { locked: false }; // Keep editable
//         const matterText = cell.value?.toString().toUpperCase() || "";
//         // Highlight "FINANCIAL" matters in Yellow
//         if (matterText.includes("ACADEMIC LEAVE") || matterText.includes("Academic Leave")) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//         }
//       }

//       if (colNum === totalCols) {
//         cell.protection = { locked: false }; // Keep editable
//         cell.alignment = { horizontal: "left", vertical: "middle" };

//         const matterText = cell.value?.toString().toUpperCase() || "";
//         // Highlight "FINANCIAL" matters in Yellow
//         if ( matterText.includes("FINANCIAL") || matterText.includes("Financial")) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//           cell.font = { bold: true, size: 8, name: fontName };
//         }
//       }

//       if (colNum === 2 || colNum === 3) cell.alignment = { horizontal: "left", vertical: "middle" };

//       if (colNum >= 5 && colNum < tuColIdx) {
//         const val = cell.value?.toString() || "";
//         if (resolvedStatus.isLocked) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
//         else if (val === "INC" || val.includes("C")) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
//         else if (typeof cell.value === "number" && cell.value < passMark) {
//           cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
//           cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 8, name: fontName };
//         }
//       }
//     });
//     row.getCell(totalCols).protection = { locked: false };
//     currentIndex++;
//   }

// const ADMIN_STATUS_MAP: Record<string, string> = {
//   on_leave: "ACADEMIC LEAVE",
//   deferred: "DEFERMENT",
//   discontinued: "DISCONTINUED",
//   deregistered: "DEREGISTERED",
//   graduated: "GRADUATED",
// };

// const ADMIN_STATUS_MAP: Record<string, string> = {
//   // DB status values (from student.status in MongoDB)
//   on_leave:     "ACADEMIC LEAVE",
//   deferred:     "DEFERMENT",
//   discontinued: "DISCONTINUED",
//   deregistered: "DEREGISTERED",
//   graduated:    "GRADUATED",
//   // Display labels (from preview objects — previewPromotion sets these)
//   "academic leave": "ACADEMIC LEAVE",
//   "deferment":      "DEFERMENT",
//   "discontinued":   "DISCONTINUED",
//   "deregistered":   "DEREGISTERED",
//   "graduated":      "GRADUATED",
//   // Alternate forms
//   "on leave":       "ACADEMIC LEAVE",
// };

// const ADMIN_STATUS_MAP: Record<string, string> = {
//   // DB status values (from student.status in MongoDB)
//   on_leave:     "ACADEMIC LEAVE",
//   deferred:     "DEFERMENT",
//   discontinued: "DISCONTINUED",
//   deregistered: "DEREGISTERED",
//   graduated:    "GRADUATED",
//   // Display labels (from preview objects — previewPromotion sets these)
//   "academic leave": "ACADEMIC LEAVE",
//   "deferment":      "DEFERMENT",
//   "on leave":       "ACADEMIC LEAVE",
// };

const ADMIN_STATUS_MAP: Record<string, string> = {
  // DB snake_case values
  on_leave: "ACADEMIC LEAVE",
  deferred: "DEFERMENT",
  discontinued: "DISCONTINUED",
  deregistered: "DEREGISTERED",
  graduated: "GRADUATED",
  repeat: "", // run engine — repeat students have marks
  // Display labels from previewPromotion (after .toLowerCase())
  "on leave": "ACADEMIC LEAVE",
  deferment: "DEFERMENT",
  stayout: "", // run engine
  "already promoted": "", // run engine
};

let currentIndex = 0;

for (const student of sortedStudents) {
  const rIdx = 11 + currentIndex;

  // ── Bulletproof sId extraction ─────────────────────────────────────────
  // Preview objects have .id (ObjectId), DB objects have ._id (ObjectId).
  // We try every possible location and convert safely.
  // let sId = "";
  // try {
  //   const rawId = student._id ?? (student as any).id ?? null;
  //   if (rawId !== null && rawId !== undefined) {
  //     sId = String(rawId);
  //   }
  // } catch {
  //   sId = "";
  // }

  // // Skip students with no valid MongoDB ObjectId — prevents CastError
  // if (!sId || !mongoose.isValidObjectId(sId)) {
  //   console.warn(
  //     "[CMS] Skipping student — invalid _id:",
  //     (student as any).regNo ?? "unknown",
  //     "| raw value:",
  //     student._id ?? (student as any).id,
  //   );
  //   continue;
  // }

  // // ── Admin status gate ──────────────────────────────────────────────────
  // const studentStatusRaw = (student as any).status ?? "";
  // const adminStatusLabel = ADMIN_STATUS_MAP[studentStatusRaw];
  // let audit: any;

  // if (adminStatusLabel) {
  //   // Synthesise audit without touching the DB engine
  //   audit = {
  //     status: adminStatusLabel,
  //     variant: "info",
  //     details: adminStatusLabel,
  //     weightedMean: "0.00",
  //     passedList: [],
  //     failedList: [],
  //     specialList: [],
  //     missingList: [],
  //     incompleteList: [],
  //     summary: {
  //       totalExpected: offeredUnits.length,
  //       passed: 0,
  //       failed: 0,
  //       missing: 0,
  //       isOnLeave: true,
  //     },
  //   };
  // } else {
  //   // Active / repeat students — run the real engine, wrapped in try-catch
  //   try {
  //     audit = await calculateStudentStatus(
  //       sId,
  //       programId,
  //       academicYear,
  //       yearOfStudy,
  //       { forPromotion: true },
  //     );
  //   } catch (engineErr: any) {
  //     console.error(
  //       `[CMS] Status engine failed for ${(student as any).regNo}:`,
  //       engineErr.message,
  //     );
  //     // Fallback: mark as incomplete rather than crashing the whole CMS
  //     audit = {
  //       status: "SESSION IN PROGRESS",
  //       variant: "info",
  //       details: "Engine error — mark data may be incomplete",
  //       weightedMean: "0.00",
  //       passedList: [],
  //       failedList: [],
  //       specialList: [],
  //       missingList: [],
  //       incompleteList: [],
  //       summary: {
  //         totalExpected: offeredUnits.length,
  //         passed: 0,
  //         failed: 0,
  //         missing: 0,
  //       },
  //     };
  //   }
  // }

  const rawId = (student as any)._id ?? (student as any).id ?? null;
  let sId = "";
  if (rawId) {
    try {
      sId = rawId.toString();
    } catch {
      sId = "";
    }
  }

  if (!sId || !mongoose.isValidObjectId(sId)) {
    console.warn("[CMS] Skipping invalid _id:", (student as any).regNo);
    continue;
  }

  const studentStatusRaw = ((student as any).status ?? "").toString().toLowerCase().trim();

  // Normalize status — handle both DB values and preview display labels
  const rawStatus = ((student as any).status ?? "")
    .toString()
    .toLowerCase()
    .trim();
  const adminStatusLabel =
    ADMIN_STATUS_MAP[rawStatus] ??
    ADMIN_STATUS_MAP[rawStatus.replace(/_/g, " ")] ??
    null;
  // adminStatusLabel === "" means "has a mapping but run the engine anyway"
  // adminStatusLabel === null means "not in map — run the engine"
  // adminStatusLabel === "ACADEMIC LEAVE" etc means "skip engine, use this label"

  let audit: any;
  if (typeof adminStatusLabel === "string" && adminStatusLabel.length > 0) {
    // Admin status — synthesise without engine
    audit = {
      status: adminStatusLabel,
      variant: "info" as const,
      details: adminStatusLabel,
      weightedMean: "0.00",
      passedList: [],
      failedList: [],
      specialList: [],
      missingList: [],
      incompleteList: [],
      summary: {
        totalExpected: offeredUnits.length,
        passed: 0,
        failed: 0,
        missing: 0,
        isOnLeave: true,
      },
    };
  } else {
    // Active / repeat / stayout — run the engine
    try {
      audit = await calculateStudentStatus(
        sId, programId, academicYear, yearOfStudy, { forPromotion: true },
      );
    } catch (err: any) {
      console.error(
        `[CMS] Engine failed for ${(student as any).regNo}:`,
        err.message,
      );
      audit = {
        status: "SESSION IN PROGRESS", variant: "info" as const, details: "Engine error", weightedMean: "0.00",
        passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
        summary: { totalExpected: offeredUnits.length, passed: 0, failed: 0, missing: 0 },
      };
    }
  }

  // ── Attempt notation ───────────────────────────────────────────────────
  const studentMarks = marks.filter(
    (m: any) => (m.student?._id?.toString() || m.student?.toString()) === sId,
  );

  // const attemptNotation = (() => {
  //   const st = studentStatusRaw.toLowerCase();
  //   if (st === "deferred") return "DEFERRED";
  //   if (st === "on_leave") return "A/SO";
  //   if (st === "discontinued") return "DISC.";
  //   if (st === "deregistered") return "DEREG.";
  //   if (st === "repeat") return "A/RA1";

  //   const attemptTypes = studentMarks.map((m: any) =>
  //     (m.attempt || "1st").toLowerCase(),
  //   );

  //   if (attemptTypes.length === 0) {
  //     const hasRepeatHistory = ((student as any).academicHistory || []).some(
  //       (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
  //     );
  //     return hasRepeatHistory ? "A/RA1" : "B/S";
  //   }
  //   if (attemptTypes.every((a: string) => a === "1st" || a === "special")) {
  //     const hasRepeatHistory = ((student as any).academicHistory || []).some(
  //       (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
  //     );
  //     return hasRepeatHistory ? "A/RA1" : "B/S";
  //   }
  //   if (attemptTypes.includes("re-take")) return "A/CF";
  //   if (attemptTypes.includes("supplementary")) return "A/S";
  //   return "B/S";
  // })();

  const buildAttemptNotation = (
    studentStatusRaw: string,
    studentQualifier: string,
    studentMarks:     any[],
    yearOfStudy:      number,
    academicHistory:  any[],
  ): string => {
    const st = studentStatusRaw.toLowerCase().replace(/_/g, " ");
   
    // Administrative statuses (preview objects use display labels)
    if (st === "deferred" || st === "deferment") return "DEF";
    if (st === "on leave" || st === "on_leave" || st === "academic leave")  return "A/L";
    if (st === "discontinued") return "DISC.";
    if (st === "deregistered") return "DEREG.";
   
    // Repeat year — sits full ordinary again (B/S, marked out of 100%)
    if (st === "repeat") return "A/RA1";
   
    // Carry-forward student — retaking CF units this year
    if (studentQualifier && studentQualifier.includes("C")) {
      // e.g. RP1C → in CMS attempt column show RP1C
      return studentQualifier; // "RP1C", "RP2C"
    }
   
    // Repeat unit (ENG.16b)
    if (studentQualifier && studentQualifier.startsWith("RPU")) {
      return studentQualifier; // "RPU1", "RPU2"
    }
   
    // Re-admission
    if (studentQualifier && studentQualifier.startsWith("RA")) {
      return studentQualifier; // "RA1", "RA2"
    }
   
    // Derive from marks
    const attemptTypes = studentMarks.map((m: any) => (m.attempt || "1st").toLowerCase());
   
    if (attemptTypes.length === 0) {
      const hasRepeatHistory = (academicHistory || []).some(
        (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy
      );
      return hasRepeatHistory ? "A/RA1" : "B/S";
    }
   
    if (attemptTypes.every((a: string) => a === "1st" || a === "special")) {
      const hasRepeatHistory = (academicHistory || []).some(
        (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy
      );
      return hasRepeatHistory ? "A/RA1" : "B/S";
    }
   
    if (attemptTypes.includes("re-take"))       return studentQualifier.includes("C") ? studentQualifier : "A/CF";
    if (attemptTypes.includes("supplementary")) return "A/S";
   
    return "B/S";
  };

  // ── Display name ───────────────────────────────────────────────────────
  const hasReturnHistory = ((student as any).statusHistory || []).some(
    (h: any) => h.status === "ACTIVE" && (h.previousStatus === "ACADEMIC LEAVE" || h.previousStatus === "DEFERMENT"),
  );
  const repeatCount = ((student as any).academicHistory || []).filter(
    (h: any) => h.isRepeatYear && h.yearOfStudy === yearOfStudy,
  ).length;

  const finalDisplayName = [
    (student as any).name || "",
    hasReturnHistory ? " (REINSTATED)" : "",
    repeatCount > 0 ? ` (RPT${repeatCount})` : "",
  ]
    .join("")
    .toUpperCase();

  // const rowData: any[] = [
  //   currentIndex + 1,
  //   (student as any).regNo || "",
  //   finalDisplayName,
  //   attemptNotation,
  // ];

    const displayRegNo = buildDisplayRegNo((student as any).regNo || "", (student as any).qualifierSuffix || "" );
      const rowData: any[] = [
        currentIndex + 1,
        displayRegNo,         // ← qualifier-suffixed reg number
        finalDisplayName,
        buildAttemptNotation,
      ];

  // ── Unit marks ─────────────────────────────────────────────────────────
  const resolvedStatus = resolveStudentStatus(student as any);

  offeredUnits.forEach((unit) => {
    if (resolvedStatus.isLocked) {
      rowData.push("");
      return;
    }

    const markObj = marks.find(
      (m: any) =>
        (m.student?._id?.toString() || m.student?.toString()) === sId &&
        m.programUnit?.unit?.code === unit.code,
    );

    if (!markObj) {
      rowData.push("INC");
      return;
    }

    const isSpecialMark =
      (markObj as any).isSpecial ||
      ((markObj as any).remarks || "").toLowerCase().includes("special");
    const markValue = (markObj as any).agreedMark ?? 0;
    const hasCA = (markObj as any).caTotal30 != null;
    const hasExam = (markObj as any).examTotal70 != null;

    if (isSpecialMark) rowData.push(`${markValue}C`);
    else if (!hasCA || !hasExam || markValue === 0) rowData.push("INC");
    else rowData.push(markValue);
  });

  // ── Recommendation ─────────────────────────────────────────────────────
  // let recomm = audit.status;
  // const lockedStatuses = [
  //   "REPEAT YEAR",
  //   "STAYOUT",
  //   "DEREGISTERED",
  //   "ACADEMIC LEAVE",
  //   "DEFERMENT",
  //   "DISCONTINUED",
  //   "GRADUATED",
  // ];

  let recomm = audit.status; // "ACADEMIC LEAVE", "PASS", "SUPP 2", etc.

  // Only run the special-case logic for students the engine actually evaluated
  const isEngineStatus = !adminStatusLabel || adminStatusLabel === "";
  const lockedLabels = new Set([
    "REPEAT YEAR",
    "STAYOUT",
    "DEREGISTERED",
    "ACADEMIC LEAVE",
    "DEFERMENT",
    "DISCONTINUED",
    "GRADUATED",
  ]);

  if (isEngineStatus && !lockedLabels.has(audit.status)) {
    const parts: string[] = [];
    if (audit.failedList?.length) parts.push(`SUPP ${audit.failedList.length}`);
    if (audit.specialList?.length)
      parts.push(`SPEC ${audit.specialList.length}`);
    if (audit.incompleteList?.length)
      parts.push(`INC ${audit.incompleteList.length}`);
    if (parts.length > 0) recomm = parts.join("; ");
  }

  // const isSpecialCase =
  //   audit.specialList?.length > 0 ||
  //   audit.failedList?.length > 0 ||
  //   audit.incompleteList?.length > 0;

  // if (isSpecialCase && !lockedStatuses.includes(audit.status)) {
  //   const parts: string[] = [];
  //   if (audit.failedList?.length > 0)
  //     parts.push(`SUPP ${audit.failedList.length}`);
  //   if (audit.specialList?.length > 0)
  //     parts.push(`SPEC ${audit.specialList.length}`);
  //   if (audit.incompleteList?.length > 0)
  //     parts.push(`INC ${audit.incompleteList.length}`);
  //   if (parts.length > 0) recomm = parts.join("; ");
  // }

  // ── Student matters ─────────────────────────────────────────────────────
  // const mattersList: string[] = [];

  // if (
  //   audit.summary?.isOnLeave ||
  //   ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)
  // ) {
  //   const leaveType = (student as any).academicLeavePeriod?.type;
  //   const remarks = ((student as any).remarks || "").toLowerCase();
  //   if (leaveType) mattersList.push(leaveType.toUpperCase());
  //   else if (remarks.includes("financial")) mattersList.push("FINANCIAL");
  //   else if (remarks.includes("compassionate") || remarks.includes("medical"))
  //     mattersList.push("COMPASSIONATE");
  // }

  // (audit.specialList || []).forEach((spec: any) => {
  //   const g = (spec.grounds || "").split(":").pop()?.trim() || "";
  //   if (g && !["special", "reason pending"].includes(g.toLowerCase())) {
  //     mattersList.push(g.toUpperCase());
  //   }
  // });

  // const finalMatters = Array.from(new Set(mattersList)).join(", ");

  // ─── REPLACE the mattersList construction ──────────────────────────────────

  const mattersList: string[] = [];

  // Source 1: admin status students have their grounds in the student record
  const leaveType = (student as any).academicLeavePeriod?.type;
  const remarks   = ((student as any).remarks || "").toLowerCase();
  const specialGroundsField = ((student as any).specialGrounds || "").toLowerCase();

  if (
    typeof adminStatusLabel === "string" && adminStatusLabel.length > 0
    || ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)
  ) {
    if (leaveType === "financial" || remarks.includes("financial") || specialGroundsField.includes("financial")) {
      mattersList.push("FINANCIAL");
    } else if (
      leaveType === "compassionate"
      || remarks.includes("compassionate")
      || remarks.includes("medical")
      || specialGroundsField.includes("compassionate")
    ) {
      mattersList.push("COMPASSIONATE");
    } else if (leaveType) {
      mattersList.push(leaveType.toUpperCase());
    }
  }

  // Source 2: engine special list (active students with specials)
  for (const spec of (audit.specialList || [])) {
    const g = (spec.grounds || "").split(":").pop()?.trim().toUpperCase() || "";
    if (g && g !== "SPECIAL" && g !== "REASON PENDING") mattersList.push(g);
  }

  const finalMatters = Array.from(new Set(mattersList)).join(", ");
  
  // ── Totals ─────────────────────────────────────────────────────────────
  const totalMarks =
    (audit.passedList || []).reduce(
      (a: number, b: any) => a + (b.mark || 0),
      0,
    ) +
    (audit.failedList || []).reduce(
      (a: number, b: any) => a + (b.mark || 0),
      0,
    );

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

  // ── Write + style ──────────────────────────────────────────────────────
  const row = sheet.getRow(rIdx);
  row.values = rowData;

  row.eachCell((cell, colNum) => {
    cell.border = thinBorder;
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.font = { size: 8, name: fontName };

    if (colNum === totalCols - 1) {
      cell.protection = { locked: false };
      const txt = (cell.value?.toString() || "").toUpperCase();
      if (
        txt.includes("ACADEMIC LEAVE") || txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")
      ) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" }};
      }
    }

    if (colNum === totalCols) {
      cell.protection = { locked: false };
      cell.alignment = { horizontal: "left", vertical: "middle" };
      const txt = (cell.value?.toString() || "").toUpperCase();
      if (txt.includes("FINANCIAL") || txt.includes("COMPASSIONATE")) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" }};
        cell.font = { bold: true, size: 8, name: fontName };
      }
    }

    if (colNum === 2 || colNum === 3) {
      cell.alignment = { horizontal: "left", vertical: "middle" };
    }

    if (colNum >= 5 && colNum < tuColIdx) {
      const val = cell.value?.toString() || "";
      if (resolvedStatus.isLocked) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" }};
      } else if (val === "INC" || val.endsWith("C")) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" }};
      } else if (typeof cell.value === "number" && cell.value < passMark) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" }};
        cell.font = {
          color: { argb: "FF9C0006" },
          bold: true, size: 8, name: fontName};
      }
    }
  });

  row.getCell(totalCols).protection = { locked: false };
  currentIndex++;
}

  const lastDataRow = 10 + students.length;

 // 4. UNIT STATISTICS 
  const statsStart = lastDataRow + 2;
  const statsLabels = [ "Mean", "Standard Deviation", "Maximum", "Minimum", "No. of Candidates", "No. of Passes", "No. of Fails", "No. of Blanks" ];

  statsLabels.forEach((label, i) => {
    const rIdx = statsStart + i;
    const r = sheet.getRow(rIdx);
    
    // Set a smaller row height for the stats section
    r.height = 15; 
    const labelCell = r.getCell(3);
    labelCell.value = label;
    // Reduced font size to 7 and simplified labels to save horizontal space
    labelCell.font = { bold: true, size: 7, name: fontName }; 
    labelCell.border = thinBorder;

    offeredUnits.forEach((_, uIdx) => {
      const colIdx = 5 + uIdx;
      const colLetter = sheet.getColumn(colIdx).letter;
      const cell = r.getCell(colIdx);
      const range = `${colLetter}11:${colLetter}${lastDataRow}`;
      
      cell.border = thinBorder;
      cell.numFmt = "0.0"; // Reduced decimal precision to save space
      cell.font = { size: 7, name: fontName }; // Match smaller font

      if (label === "No. of Passes") {
        cell.value = { formula: `COUNTIF(${range}, ">=${passMark}")` };
        cell.numFmt = "0"; // Integers only
      } else if (label === "No. of Fails") {
        // Count numbers less than passMark but NOT empty/Locked cells
        cell.value = { formula: `COUNTIFS(${range}, "<${passMark}", ${range}, "<>")` };
        cell.numFmt = "0"; 
      } else if (label === "No. of Blanks") {
        cell.value = { formula: `COUNTIF(${range}, "INC")` };
        cell.numFmt = "0";
      } else {
        let func = "";
        switch(label) {
          case "Mean": func = "AVERAGE"; break;
          case "Standard Deviation": func = "STDEV.P"; break;
          case "Maximum": func = "MAX"; break;
          case "Minimum": func = "MIN"; break;
          case "No. of Candidates": func = "COUNT"; break;
        }
        // Wrapping the formula in ROUND(..., 1) ensures it fits in column width 6
        cell.value = { formula: `IFERROR(ROUND(${func}(${range}), 1), 0)` };
      }
    });

    // Apply thick borders only to the outer edges of the stats block to make it look compact
    r.getCell(3).border = { ...r.getCell(3).border, left: { style: "thick" } };
    r.getCell(tuColIdx - 1).border = { ...r.getCell(tuColIdx - 1).border, right: { style: "thick" }};
    
    if (i === 0) {
      for(let c = 3; c < tuColIdx; c++) {
        r.getCell(c).border = { ...r.getCell(c).border, top: { style: "thick" } };
      }
    }
    if (i === statsLabels.length - 1) {
      for(let c = 3; c < tuColIdx; c++) {
        r.getCell(c).border = { ...r.getCell(c).border, bottom: { style: "thick" } };
      }
    }
  });

  // 5. SUMMARY TABLE (Dynamic: Only shows rows with count > 0)
  const summaryStart = lastDataRow + 12;
  const summaryHeaderCell = sheet.getCell(`B${summaryStart}`);
  summaryHeaderCell.value = "SUMMARY";
  summaryHeaderCell.font = { bold: true, size: 10, underline: true, name: fontName };

  const summaryData: Record<string, number> = { "PASS": 0, "SUPPLEMENTARY": 0, "REPEAT YEAR": 0, "STAY OUT": 0, "SPECIAL": 0, "INCOMPLETE": 0, "ACADEMIC LEAVE": 0, "DEFERMENT": 0, "DEREGISTERED/DISC": 0 };

  // Tally totals from the Recommendation Column
  sheet.getColumn(tuColIdx + 3).eachCell({ includeEmpty: false }, (cell, rowNum) => {
    if (rowNum > 10 && rowNum <= lastDataRow) {
      const txt = cell.value?.toString().toUpperCase() || "";
      if (txt === "PASS") summaryData.PASS++;
      else if (txt.includes("SUPP")) summaryData.SUPPLEMENTARY++;
      else if (txt.includes("REPEAT")) summaryData["REPEAT YEAR"]++;
      else if (txt.includes("STAY OUT")) summaryData["STAY OUT"]++;
      else if (txt.includes("SPEC")) summaryData.SPECIAL++;
      else if (txt.includes("ACADEMIC LEAVE")) summaryData["ACADEMIC LEAVE"]++;
      else if (txt.includes("DEFERMENT")) summaryData["DEFERMENT"]++;
      else if (txt.includes("INC")) summaryData.INCOMPLETE++;
      else if (txt.includes("DEREG") || txt.includes("DISC")) summaryData["DEREGISTERED/DISC"]++;
    }
  });

  // Filter to only include statuses that have at least one student
  const activeSummaryEntries = Object.entries(summaryData).filter(([_, count]) => count > 0);

  activeSummaryEntries.forEach(([label, count], i) => {
    const rIdx = summaryStart + 1 + i;
    const labelCell = sheet.getCell(`B${rIdx}`);
    const countCell = sheet.getCell(`C${rIdx}`);

    labelCell.value = label;
    countCell.value = count;

    labelCell.border = thinBorder;
    countCell.border = thinBorder;
    labelCell.font = { size: 8, name: fontName, bold: true };
    countCell.font = { size: 8, name: fontName };
  });

  // 6. OFFERED UNITS TABLE (Positioned dynamically based on Summary Table height)
  const unitsStart = summaryStart + activeSummaryEntries.length + 4;
  sheet.mergeCells(unitsStart, 2, unitsStart, 6);
  sheet.getCell(unitsStart, 2).value = "LIST OF UNITS OFFERED";
  sheet.getCell(unitsStart, 2).font = { bold: true, underline: true };

  const mid = Math.ceil(offeredUnits.length / 2);
  const unitsEndRow = unitsStart + 1 + mid;
  for (let i = 0; i < mid; i++) {
    const rIdx = unitsStart + 2 + i; const r = sheet.getRow(rIdx); const left = offeredUnits[i]; const right = offeredUnits[mid + i];

    r.getCell(2).value = i + 1; r.getCell(3).value = left.code; sheet.mergeCells(rIdx, 4, rIdx, 7); r.getCell(4).value = left.name;

    if (right) { r.getCell(9).value = mid + i + 1; r.getCell(10).value = right.code; sheet.mergeCells(rIdx, 11, rIdx, 14); r.getCell(11).value = right.name; }

    [2, 3, 4, 9, 10, 11].forEach((col) => {
      const cell = r.getCell(col);
      cell.border = { ...thinBorder };
      if (col === 2) cell.border.left = { style: "thick" };
      if (col === 11 || (!right && col === 4))
        cell.border.right = { style: "thick" };
      if (i === 0) cell.border.top = { style: "thick" };
      if (i === mid - 1) cell.border.bottom = { style: "thick" };
      cell.font = { size: 8 };
    });
  }

  // 7. MAIN TABLE THICK BORDERS
  for (let i = startRow; i <= lastDataRow; i++) {
    sheet.getCell(i, 1).border = { ...sheet.getCell(i, 1).border, left: { style: "thick" }};
    sheet.getCell(i, totalCols).border = { ...sheet.getCell(i, totalCols).border, right: { style: "thick" }};
  }
  sheet.getRow(startRow).eachCell((c) => (c.border = { ...c.border, top: { style: "thick" } }));
  sheet.getRow(lastDataRow).eachCell((c) => (c.border = { ...c.border, bottom: { style: "thick" } }));

  // Sheet Formatting
  // Column Widths
  sheet.getColumn(1).width = 4; sheet.getColumn(2).width = 20; sheet.getColumn(3).width = 25;
  sheet.getColumn(4).width = 5; offeredUnits.forEach((_, i) => (sheet.getColumn(5 + i).width = 4.5));
  sheet.getColumn(tuColIdx).width = 5; sheet.getColumn(tuColIdx + 1).width = 7;
  sheet.getColumn(tuColIdx + 2).width = 7; sheet.getColumn(tuColIdx + 3).width = 20; sheet.getColumn(tuColIdx + 4).width = 20;

  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 10 }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0  };
  const result = await workbook.xlsx.writeBuffer();
  return Buffer.from(result as ArrayBuffer);
};

// serverside/src/utils/consolidatedMSMulti.ts
//
// Generates the multi-sheet CMS workbook that matches the real institution format:
//
//  Sheet 1: "Civil Yr1 2024-2025" (ORDINARY)
//    All students, B/S attempt, full marks, all statuses
//
//  Sheet 2: "ACF ASO Yr1 2024-2025" (SUPPLEMENTARY)
//    Only students who need supp/specials/carry-forward/stayout
//    Each student has multiple rows: B/S row, then A/S row, then A/CF row etc.
//
//  Sheet 3: "ARA Yr1 2024-2025" (REPEAT YEAR)
//    Only repeat-year students
//    B/S row → A/RA row
//
// This matches the format from the institution's actual spreadsheets.

// import * as ExcelJS from "exceljs";
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";
// import MarkDirect from "../models/MarkDirect";
// import FinalGrade from "../models/FinalGrade";
// import Program from "../models/Program";
// import AcademicYear from "../models/AcademicYear";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { calculateStudentStatus } from "../services/statusEngine";
// import { resolveStudentStatus } from "./studentStatusResolver";
// import { getAttemptLabel, ADMIN_STATUS_LABELS, getExamTypeTitle } from "./academicRules";
// import config from "../config/config";

// // ─── Types ────────────────────────────────────────────────────────────────────

// interface OfferedUnit { code: string; name: string }

// interface StudentRow {
//   sId:    string;
//   regNo:  string;
//   name:   string;
//   status: string;
//   academicHistory?: Array<{ yearOfStudy: number; isRepeatYear?: boolean; academicYear?: string }>;
//   academicLeavePeriod?: { type?: string };
//   marks: Map<string, number | "INC" | "C">; // unitCode → value
//   attempt:   string;
//   totalUnits: number;
//   total:     number | "-";
//   mean:      number | "-";
//   recomm:    string;
//   matters:   string;
// }

// export interface ConsolidatedData {
//   programName:   string;
//   programId:     string;
//   academicYear:  string;
//   yearOfStudy:   number;
//   session:       "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
//   students:      Array<Record<string, unknown>>;
//   marks:         Array<Record<string, unknown>>;
//   offeredUnits:  OfferedUnit[];
//   logoBuffer:    Buffer;
//   institutionId: string;
//   passMark:      number;
//   gradingScale:  Array<{ min: number; grade: string }>;
// }

// // ─── Helpers ──────────────────────────────────────────────────────────────────

// const fontName = "Arial";

// const thinBorder: Partial<ExcelJS.Borders> = {
//   top: { style: "thin" }, left: { style: "thin" },
//   bottom: { style: "thin" }, right: { style: "thin" },
// };
// const doubleBorder: Partial<ExcelJS.Borders> = {
//   ...thinBorder, bottom: { style: "double" },
// };

// const gradeFromScale = (
//   mark: number,
//   scale: Array<{ min: number; grade: string }>,
//   passMark: number
// ): string => {
//   if (mark < passMark) return "E";
//   const sorted = [...scale].sort((a, b) => b.min - a.min);
//   return sorted.find((s) => mark >= s.min)?.grade ?? "E";
// };

// const shade = (fill: string): ExcelJS.Fill => ({
//   type: "pattern", pattern: "solid", fgColor: { argb: fill },
// });

// // ─── Sheet builder ────────────────────────────────────────────────────────────

// function buildSheet(
//   wb:           ExcelJS.Workbook,
//   sheetName:    string,
//   rows:         StudentRow[],
//   units:        OfferedUnit[],
//   data:         ConsolidatedData,
//   examLabel:    string,   // e.g. "ORDINARY EXAMINATION" or "SUPPLEMENTARY AND SPECIAL EXAMINATION"
//   sheetTitle:   string,   // e.g. "ORDINARY" or "SUPPLEMENTARY/CARRY FORWARD"
//   logoBuffer:   any,
// ): void {
//   const sheet = wb.addWorksheet(sheetName.substring(0, 31));

//   const tuColIdx   = 5 + units.length;
//   const totalCols  = tuColIdx + 4;

//   const yrTxt = ["FIRST","SECOND","THIRD","FOURTH","FIFTH"][data.yearOfStudy - 1]
//     ?? `${data.yearOfStudy}TH`;

//   // ── Logo ───────────────────────────────────────────────────────────────
//   if (logoBuffer?.length > 0) {
//     const logoId = wb.addImage({ buffer: logoBuffer, extension: "png" });
//     const midCol = Math.floor(totalCols / 2);
//     sheet.addImage(logoId, { tl: { col: midCol - 1, row: 0 }, ext: { width: 100, height: 60 } });
//   }

//   const setCenteredHeader = (rowNum: number, text: string, size = 10) => {
//     sheet.mergeCells(rowNum, 1, rowNum, totalCols);
//     const cell = sheet.getCell(rowNum, 1);
//     cell.value = text.toUpperCase();
//     cell.style = {
//       alignment: { horizontal: "center", vertical: "middle" },
//       font: { bold: true, name: fontName, size: size - 1 },
//     };
//   };

//   setCenteredHeader(4, config.instName, 12);
//   setCenteredHeader(5, config.schoolName || "SCHOOL OF ENGINEERING");
//   setCenteredHeader(6, data.programName);
//   setCenteredHeader(7,
//     `CONSOLIDATED MARK SHEET FOR ${data.academicYear} A.Y. - ` +
//     `YEAR ${data.yearOfStudy} (${sheetTitle})`
//   );
//   sheet.getCell(7, 1).font = { ...sheet.getCell(7, 1).font, underline: true };

//   // ── Column headers (rows 9–10) ─────────────────────────────────────────
//   const startRow = 9;
//   const subRow   = 10;
//   sheet.getRow(subRow).height = 48;

//   const fixedHeaders: Record<number, string> = {
//     1: "S/N", 2: "REG. NO", 3: "NAME", 4: "ATTEMPT",
//     [tuColIdx]:     "T U",
//     [tuColIdx + 1]: "TOTAL",
//     [tuColIdx + 2]: "MEAN",
//     [tuColIdx + 3]: "RECOMM.",
//     [tuColIdx + 4]: "STUDENT MATTERS",
//   };

//   Object.entries(fixedHeaders).forEach(([col, text]) => {
//     const colNum = parseInt(col);
//     sheet.mergeCells(startRow, colNum, subRow, colNum);
//     const cell = sheet.getCell(startRow, colNum);
//     cell.value = text;
//     cell.style = {
//       alignment: { horizontal: "center", vertical: "middle",
//         textRotation: colNum === 4 ? 90 : 0, wrapText: true },
//       font:   { bold: true, size: 7, name: fontName },
//       border: doubleBorder,
//     };
//   });

//   units.forEach((unit, i) => {
//     const colIdx = 5 + i;
//     const h1     = sheet.getCell(startRow, colIdx);
//     h1.value     = (i + 1).toString();
//     h1.style     = { alignment: { horizontal: "center", vertical: "middle" },
//       font: { bold: true, size: 7, name: fontName }, border: thinBorder };
//     const h2     = sheet.getCell(subRow, colIdx);
//     h2.value     = unit.code;
//     h2.style     = { alignment: { horizontal: "center", vertical: "middle",
//       textRotation: 90 }, font: { bold: true, size: 7, name: fontName }, border: thinBorder };
//   });

//   // ── Data rows ──────────────────────────────────────────────────────────
//   let idx = 0;
//   for (const row of rows) {
//     const rIdx = 11 + idx;
//     const r    = sheet.getRow(rIdx);

//     const rowData: (string | number | null)[] = [
//       idx + 1,
//       row.regNo,
//       row.name.toUpperCase(),
//       row.attempt,
//     ];

//     units.forEach((unit) => {
//       const val = row.marks.get(unit.code);
//       rowData.push(val !== undefined ? (val as string | number) : "INC");
//     });

//     rowData.push(
//       row.totalUnits,
//       row.total as number | null,
//       typeof row.mean === "number" ? parseFloat(row.mean.toFixed(2)) : "-" as unknown as null,
//       row.recomm,
//       row.matters,
//     );

//     r.values = rowData;

//     r.eachCell((cell, colNum) => {
//       cell.border    = thinBorder;
//       cell.alignment = { horizontal: "center", vertical: "middle" };
//       cell.font      = { size: 8, name: fontName };

//       if (colNum >= 5 && colNum < tuColIdx) {
//         const v = cell.value?.toString() || "";
//         if (v === "INC" || v.endsWith("C")) {
//           cell.fill = shade("FFFFFF00"); // yellow for INC/Special
//         } else if (typeof cell.value === "number" && cell.value < data.passMark) {
//           cell.fill = shade("FFFFC7CE"); // red for fail
//           cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 8, name: fontName };
//         }
//       }
//       if (colNum === 2 || colNum === 3) {
//         cell.alignment = { horizontal: "left", vertical: "middle" };
//       }
//       if (colNum === totalCols - 1) cell.protection = { locked: false };
//       if (colNum === totalCols) {
//         cell.protection = { locked: false };
//         cell.alignment  = { horizontal: "left", vertical: "middle" };
//       }
//     });

//     idx++;
//   }

//   const lastDataRow = 10 + rows.length;

//   // ── Stats block ────────────────────────────────────────────────────────
//   const statsStart  = lastDataRow + 2;
//   const statsLabels = ["Mean","Standard Deviation","Maximum","Minimum",
//     "No. of Candidates","No. of Passes","No. of Fails","No. of Blanks"];

//   statsLabels.forEach((label, i) => {
//     const rIdx = statsStart + i;
//     const r    = sheet.getRow(rIdx);
//     r.height   = 15;
//     const lc   = r.getCell(3);
//     lc.value   = label;
//     lc.font    = { bold: true, size: 7, name: fontName };
//     lc.border  = thinBorder;

//     units.forEach((_, uIdx) => {
//       const colIdx    = 5 + uIdx;
//       const colLetter = sheet.getColumn(colIdx).letter;
//       const cell      = r.getCell(colIdx);
//       const range     = `${colLetter}11:${colLetter}${lastDataRow}`;
//       cell.border     = thinBorder;
//       cell.numFmt     = "0.0";
//       cell.font       = { size: 7, name: fontName };

//       if (label === "No. of Passes") {
//         cell.value  = { formula: `COUNTIF(${range}, ">=${data.passMark}")` };
//         cell.numFmt = "0";
//       } else if (label === "No. of Fails") {
//         cell.value  = { formula: `COUNTIFS(${range}, "<${data.passMark}", ${range}, "<>")` };
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
//         if (fn) cell.value = { formula: `IFERROR(ROUND(${fn}(${range}),1),0)` };
//       }
//     });
//   });

//   // ── Summary ────────────────────────────────────────────────────────────
//   const summaryStart = lastDataRow + 12;
//   sheet.getCell(`B${summaryStart}`).value = "SUMMARY";
//   sheet.getCell(`B${summaryStart}`).font  = { bold: true, size: 10, underline: true, name: fontName };

//   const summaryData: Record<string, number> = {
//     PASS: 0, SUPPLEMENTARY: 0, "REPEAT YEAR": 0, STAYOUT: 0,
//     SPECIAL: 0, INCOMPLETE: 0, "ACADEMIC LEAVE": 0, DEFERMENT: 0,
//     "DEREGISTERED/DISC": 0,
//   };
//   rows.forEach((row) => {
//     const r = row.recomm.toUpperCase();
//     if (r === "PASS")                            summaryData.PASS++;
//     else if (r.includes("SUPP"))                 summaryData.SUPPLEMENTARY++;
//     else if (r.includes("REPEAT"))               summaryData["REPEAT YEAR"]++;
//     else if (r.includes("STAYOUT"))              summaryData.STAYOUT++;
//     else if (r.includes("SPEC"))                 summaryData.SPECIAL++;
//     else if (r.includes("INC") && !r.includes("SPEC")) summaryData.INCOMPLETE++;
//     else if (r === "ACADEMIC LEAVE")             summaryData["ACADEMIC LEAVE"]++;
//     else if (r === "DEFERMENT")                  summaryData.DEFERMENT++;
//     else if (r.includes("DEREG") || r.includes("DISC")) summaryData["DEREGISTERED/DISC"]++;
//   });

//   let sRow = summaryStart + 1;
//   Object.entries(summaryData)
//     .filter(([, v]) => v > 0)
//     .forEach(([label, count]) => {
//       sheet.getCell(`B${sRow}`).value = label;
//       sheet.getCell(`C${sRow}`).value = count;
//       sheet.getCell(`B${sRow}`).border = thinBorder;
//       sheet.getCell(`C${sRow}`).border = thinBorder;
//       sheet.getCell(`B${sRow}`).font   = { size: 8, name: fontName, bold: true };
//       sheet.getCell(`C${sRow}`).font   = { size: 8, name: fontName };
//       sRow++;
//     });

//   // ── Units offered table ────────────────────────────────────────────────
//   const unitsStart = sRow + 3;
//   sheet.mergeCells(unitsStart, 2, unitsStart, 6);
//   sheet.getCell(unitsStart, 2).value = "LIST OF UNITS OFFERED";
//   sheet.getCell(unitsStart, 2).font  = { bold: true, underline: true };

//   const mid = Math.ceil(units.length / 2);
//   for (let i = 0; i < mid; i++) {
//     const ri = unitsStart + 2 + i;
//     const r  = sheet.getRow(ri);
//     const l  = units[i];
//     const rt = units[mid + i];
//     r.getCell(2).value = i + 1;
//     r.getCell(3).value = l.code;
//     sheet.mergeCells(ri, 4, ri, 7);
//     r.getCell(4).value = l.name;
//     if (rt) {
//       r.getCell(9).value  = mid + i + 1;
//       r.getCell(10).value = rt.code;
//       sheet.mergeCells(ri, 11, ri, 14);
//       r.getCell(11).value = rt.name;
//     }
//     [2, 3, 4, 9, 10, 11].forEach((c) => {
//       r.getCell(c).border = thinBorder;
//       r.getCell(c).font   = { size: 8 };
//     });
//   }

//   // ── Column widths ──────────────────────────────────────────────────────
//   sheet.getColumn(1).width = 4;
//   sheet.getColumn(2).width = 20;
//   sheet.getColumn(3).width = 25;
//   sheet.getColumn(4).width = 5;
//   units.forEach((_, i) => (sheet.getColumn(5 + i).width = 4.5));
//   sheet.getColumn(tuColIdx).width     = 5;
//   sheet.getColumn(tuColIdx + 1).width = 7;
//   sheet.getColumn(tuColIdx + 2).width = 7;
//   sheet.getColumn(tuColIdx + 3).width = 20;
//   sheet.getColumn(tuColIdx + 4).width = 20;

//   sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 10 }];
//   sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
//   sheet.pageSetup = { orientation: "landscape", paperSize: 9,
//     fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
// }

// // ─── Main export ──────────────────────────────────────────────────────────────

// export const generateConsolidatedMarkSheet = async (
//   data: ConsolidatedData
// ): Promise<Buffer> => {
//   const wb = new ExcelJS.Workbook();
//   const {
//     programName, programId, academicYear, yearOfStudy,
//     session, students, marks, offeredUnits,
//     logoBuffer, passMark, gradingScale,
//   } = data;

//   const ADMIN_KEYS = new Set(Object.keys(ADMIN_STATUS_LABELS));

//   // ── Helper: build a StudentRow from student + audit ──────────────────────
//   const buildRow = (
//     student: Record<string, unknown>,
//     audit:   Record<string, unknown>,
//     attemptOverride?: string
//   ): StudentRow => {
//     const sId   = (student._id as mongoose.Types.ObjectId)?.toString() || "";
//     const st    = (student.status as string || "").toLowerCase();
//     const regNo = student.regNo as string || "";
//     const name  = student.name  as string || "";

//     const history = (student.academicHistory as Array<{ yearOfStudy: number; isRepeatYear?: boolean }> | undefined) || [];
//     const repeatCount = history.filter(
//       (h) => h.isRepeatYear && h.yearOfStudy === yearOfStudy
//     ).length;

//     const studentMarks = (marks as Array<Record<string, unknown>>).filter(
//       (m) => {
//         const ms = (m.student as { _id?: mongoose.Types.ObjectId; toString?: () => string });
//         return ms?._id?.toString() === sId || ms?.toString?.() === sId;
//       }
//     );

//     const notation = attemptOverride ?? getAttemptLabel({
//       markAttempt: "1st",
//       studentStatus: student.status as string,
//       regNo,
//       repeatYearCount: repeatCount,
//     });

//     const marksMap = new Map<string, number | "INC" | "C">();
//     const adminLabel = ADMIN_KEYS.has(st.replace(" ", "_"))
//       ? ADMIN_STATUS_LABELS[st.replace(" ", "_")]
//       : null;

//     if (!adminLabel) {
//       offeredUnits.forEach((unit) => {
//         const m = studentMarks.find(
//           (mk) => {
//             const pu = mk.programUnit as Record<string, unknown> | null;
//             return (pu?.unit as Record<string, unknown>)?.code === unit.code;
//           }
//         );
//         if (!m) { marksMap.set(unit.code, "INC"); return; }
//         const v = (m.agreedMark as number) ?? 0;
//         const isSpec = (m.isSpecial as boolean) || (m.remarks as string || "").toLowerCase().includes("special");
//         if (isSpec)  marksMap.set(unit.code, `${v}C` as unknown as "C");
//         else if (v === 0) marksMap.set(unit.code, "INC");
//         else marksMap.set(unit.code, v);
//       });
//     }

//     const auditTyped = audit as {
//       status: string; weightedMean: string; summary: { totalExpected: number };
//       failedList: Array<{ displayName: string; attempt: number }>;
//       specialList: Array<{ displayName: string; grounds: string }>;
//       incompleteList: string[];
//     };

//     let recomm = auditTyped.status;
//     if (adminLabel) {
//       recomm = adminLabel;
//     } else {
//       const parts: string[] = [];
//       if (auditTyped.failedList?.length)    parts.push(`SUPP ${auditTyped.failedList.length}`);
//       if (auditTyped.specialList?.length)   parts.push(`SPEC ${auditTyped.specialList.length}`);
//       if (auditTyped.incompleteList?.length) parts.push(`INC ${auditTyped.incompleteList.length}`);
//       if (parts.length > 0) recomm = parts.join("; ");
//     }

//     const matters: string[] = [];
//     const leaveType = (student.academicLeavePeriod as { type?: string } | undefined)?.type;
//     if (leaveType) matters.push(leaveType.toUpperCase());
//     auditTyped.specialList?.forEach((s) => {
//       const g = (s.grounds || "").split(":").pop()?.trim().toUpperCase();
//       if (g && g !== "SPECIAL") matters.push(g);
//     });

//     const mean   = parseFloat(auditTyped.weightedMean || "0");
//     const numericMarks = Array.from(marksMap.values()).filter(
//       (v) => typeof v === "number"
//     ) as number[];
//     const total  = adminLabel ? "-" : numericMarks.reduce((a, b) => a + b, 0);

//     return {
//       sId, regNo, name,
//       status:     student.status as string,
//       attempt:    notation,
//       totalUnits: auditTyped.summary?.totalExpected ?? offeredUnits.length,
//       marks:      marksMap,
//       total:      adminLabel ? "-" : total,
//       mean:       adminLabel ? "-" : mean,
//       recomm,
//       matters:    Array.from(new Set(matters)).join(", "),
//     };
//   };

//   // ── Compute audit for all active students ────────────────────────────────
//   const sortedStudents = [...(students as Array<Record<string, unknown>>)].sort(
//     (a, b) => (a.regNo as string).localeCompare(b.regNo as string)
//   );

//   const ordinaryRows:     StudentRow[] = [];
//   const suppRows:         StudentRow[] = [];
//   const repeatYearRows:   StudentRow[] = [];

//   for (const student of sortedStudents) {
//     // const sId = (student._id as mongoose.Types.ObjectId)?.toString() || "";
//     const rawId = student._id ?? student.id;
// const sId   = typeof rawId === "object"
//   ? rawId?.toString()
//   : typeof rawId === "string"
//   ? rawId
//   : "";

// // Skip any student with an invalid ID — prevents CastError
// if (!sId || sId.length < 24) {
//   console.warn("[CMS] Skipping student with invalid _id:", student.regNo ?? "unknown");
//   continue;
// }
//     const st  = (student.status as string || "").toLowerCase().replace(/ /g, "_");

//     let audit: Record<string, unknown>;

//     if (ADMIN_KEYS.has(st)) {
//       audit = {
//         status: ADMIN_STATUS_LABELS[st] || st.toUpperCase(),
//         weightedMean: "0.00",
//         summary: { totalExpected: offeredUnits.length, passed: 0, failed: 0, missing: 0 },
//         passedList: [], failedList: [], specialList: [],
//         missingList: [], incompleteList: [],
//       };
//     } else {
//       audit = await calculateStudentStatus(
//         sId, programId, academicYear, yearOfStudy,
//         { forPromotion: true }
//       ) as unknown as Record<string, unknown>;
//     }

//     const auditTyped = audit as { status: string };

//     // ── Sheet 1 (ORDINARY): all students ──────────────────────────────────
//     ordinaryRows.push(buildRow(student, audit));

//     // ── Sheet 2 (SUPP/ACF/ASO): students who need further action ──────────
//     const suppStatuses = ["SUPP", "SPEC", "INC", "STAYOUT", "REPEAT YEAR",
//       "ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED", "DISCONTINUED"];
//     const needsSupp = suppStatuses.some((s) => auditTyped.status.toUpperCase().includes(s));

//     if (needsSupp && !["ACADEMIC LEAVE","DEFERMENT"].includes(
//       (ADMIN_STATUS_LABELS[st] || "").toUpperCase()
//     )) {
//       // Row 1: B/S (original ordinary marks)
//       suppRows.push(buildRow(student, audit, "B/S"));
//       // Row 2: A/S (supplementary attempt — marks not yet known, show INC)
//       const suppRow = buildRow(student, audit, "A/S");
//       suppRow.marks = new Map(
//         Array.from(suppRow.marks.entries()).map(([code, v]) => [
//           code,
//           // In supp, only failed units show; passed ones carry over
//           typeof v === "number" && v >= passMark ? v : "INC",
//         ])
//       );
//       suppRow.recomm = "PENDING";
//       suppRows.push(suppRow);
//     }

//     // ── Sheet 3 (A/RA): repeat year students ──────────────────────────────
//     if (auditTyped.status === "REPEAT YEAR" || st === "repeat") {
//       const history = (student.academicHistory as Array<{ yearOfStudy: number; isRepeatYear?: boolean }> | undefined) || [];
//       const repeatCount = history.filter(
//         (h) => h.isRepeatYear && h.yearOfStudy === yearOfStudy
//       ).length;
//       const raLabel = repeatCount >= 2 ? "A/RA2" : "A/RA1";

//       repeatYearRows.push(buildRow(student, audit, "B/S"));
//       repeatYearRows.push(buildRow(student, audit, raLabel));
//     }
//   }

//   const examLabel = getExamTypeTitle(session);

//   // ── Build sheets ─────────────────────────────────────────────────────────
//   const yearStr = academicYear.replace("/", "-");

//   buildSheet(wb, `${programName.split(" ").pop()} Yr${yearOfStudy} ${yearStr}`,
//     ordinaryRows, offeredUnits, data, examLabel, "ORDINARY EXAMINATION", logoBuffer);

//   if (suppRows.length > 0) {
//     buildSheet(wb, `ACF ASO Yr${yearOfStudy} ${yearStr}`,
//       suppRows, offeredUnits, data,
//       "SUPPLEMENTARY AND SPECIAL EXAMINATION",
//       "SUPPLEMENTARY/CARRY FORWARD", logoBuffer);
//   }

//   if (repeatYearRows.length > 0) {
//     buildSheet(wb, `ARA Yr${yearOfStudy} ${yearStr}`,
//       repeatYearRows, offeredUnits, data,
//       "REPEAT YEAR EXAMINATION", "REPEAT YEAR (A/RA)", logoBuffer);
//   }

//   const buf = await wb.xlsx.writeBuffer();
//   return Buffer.from(buf as ArrayBuffer);
// };



