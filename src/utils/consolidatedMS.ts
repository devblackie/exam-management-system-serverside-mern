// serverside/src/utils/consolidatedMS.ts
import * as ExcelJS from "exceljs";
import config from "../config/config";
import InstitutionSettings from "../models/InstitutionSettings";
import { resolveStudentStatus } from "./studentStatusResolver";
import { calculateStudentStatus } from "../services/statusEngine";
import { getAttemptLabel } from "./academicRules";

export interface ConsolidatedData {
  programName: string; academicYear: string; yearOfStudy: number;
  students: any[]; marks: any[];
  offeredUnits: { code: string; name: string }[];
  logoBuffer?: any; institutionId?: string;  programId: string;
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

  const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
  setCenteredHeader(4, `${config.instName}`);
  setCenteredHeader(5, `${config.schoolName || "SCHOOL OF ENGINEERING"}`);
  setCenteredHeader(6, `${programName}`);
  setCenteredHeader(7, `CONSOLIDATED MARK SHEET - ORDINARY EXAMINATION RESULTS - ${yrTxt} YEAR - ${academicYear} ACADEMIC YEAR`);
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
  const sortedStudents = [...students].sort((a, b) => (a.regNo || "").localeCompare(b.regNo || ""));
  
  let currentIndex = 0;
  for (const student of sortedStudents) {
    const rIdx = 11 + currentIndex;
    const sId = student.id?.toString() || student._id?.toString();
    const audit = await calculateStudentStatus( sId, programId, academicYear, yearOfStudy );
    const resolvedStatus = resolveStudentStatus(student);
    // --- REINSTATED FLAG LOGIC ---
    const hasReturnHistory = student.statusHistory?.some(
      (h: any) => h.status === "ACTIVE" && (h.previousStatus === "ACADEMIC LEAVE" || h.previousStatus === "DEFERMENT"));

    // --- REPEAT FLAG LOGIC (RPT) ---
    // Count how many times they have a record for this specific year of study in their history
    const repeatCount = student.academicHistory?.filter((h: any) => h.yearOfStudy === yearOfStudy && h.status === "REPEAT YEAR").length || 0;

    // Determine the dynamic notation for Column 4
    const attemptNotation = getAttemptLabel(student.attemptNumber || 1, student.status, student.regNo);
    
    const repeatTag = repeatCount > 0 ? ` (RPT${repeatCount})` : "";
    const reinstatedTag = hasReturnHistory ? " (REINSTATED)" : "";

    // Final Display Name: "John Doe (REINSTATED) (RPT1)"
    const finalDisplayName = `${student.name}${reinstatedTag}${repeatTag}`.toUpperCase();

    // const rowData: any[] = [ currentIndex + 1, student.regNo, finalDisplayName, "B/S" ];
    const rowData: any[] = [ currentIndex + 1, student.regNo, finalDisplayName, attemptNotation ];

    // const rowData: any[] = [currentIndex + 1, student.regNo, student.name, "B/S"];

    // Fill Unit Marks
    offeredUnits.forEach((unit) => {
      const markObj = marks.find((m) => (m.student?._id?.toString() || m.student?.toString()) === sId && m.programUnit?.unit?.code === unit.code);

      if (resolvedStatus.isLocked) rowData.push("");
      else if (markObj) {
        const isSpecial = markObj.isSpecial || markObj.remarks?.toLowerCase().includes("special");
        const markValue = markObj.agreedMark ?? 0;
        // const isMissingData = !markObj.caTotal30 || !markObj.examTotal70;

        // FIX: Check for null/undefined instead of falsy 0
        const hasCA = markObj.caTotal30 !== null && markObj.caTotal30 !== undefined;
        const hasExam = markObj.examTotal70 !== null && markObj.examTotal70 !== undefined;

        // In Direct Entry, if they have the markObj, they usually have the data.
        // We only show INC if the fields are physically missing from the DB record.
        const isMissingData = !hasCA || !hasExam;

        if (isSpecial) rowData.push(`${markValue}C`);
        else if (isMissingData || markValue === 0) rowData.push("INC");
        else rowData.push(markValue);
      } else rowData.push("INC");
    });

    // Formatting the Specific Recommendation Column
    let recomm = audit.status;    
    const isSpecialCase = audit.specialList.length > 0 || audit.failedList.length > 0 || audit.incompleteList.length > 0;

    if ( isSpecialCase && !["REPEAT YEAR", "STAYOUT", "DEREGISTERED"].includes(audit.status)) {
      const parts = [];
      if (audit.failedList.length > 0) parts.push(`SUPP ${audit.failedList.length}`);
      if (audit.specialList.length > 0) parts.push(`SPEC ${audit.specialList.length}`);
      if (audit.incompleteList.length > 0) parts.push(`INC ${audit.incompleteList.length}`);
      recomm = parts.length > 0 ? parts.join("; ") : audit.status;
    }

    const studentMattersList: string[] = [];

    // 1. Check for Academic Leave Grounds
    if (audit.summary.isOnLeave || ["ACADEMIC LEAVE", "DEFERMENT", "ON LEAVE"].includes(audit.status)) {
      // Priority 1: The structured leave period type
      const leaveType = student.academicLeavePeriod?.type;
      // Priority 2: The remarks field
      const remarks = student.remarks?.toLowerCase() || "";
      
      if (leaveType) studentMattersList.push(leaveType.toUpperCase());
      else if (remarks.includes("financial")) studentMattersList.push("FINANCIAL");
      else if (remarks.includes("compassionate") || remarks.includes("medical")) {
          studentMattersList.push("COMPASSIONATE");
      } else if (audit.leaveDetails) {
          // Fallback to the reason string from resolveStudentStatus
          const cleanReason = audit.leaveDetails.split(":").pop()?.trim().toUpperCase();
          if (cleanReason && !cleanReason.includes("PENDING")) studentMattersList.push(cleanReason);
      }
  }

    // 2. Check for Special Grounds (from engine/marks)
    audit.specialList.forEach((spec: any) => {
      if (spec.grounds) {
        // This catches "Special Granted: Financial"
        const cleanSpec = spec.grounds.split(":").pop()?.trim() || spec.grounds;
        if (!["special", "reason pending"].includes(cleanSpec.toLowerCase())) {
          studentMattersList.push(cleanSpec.toUpperCase());
        }
      }
    });

    const finalMatters = Array.from(new Set(studentMattersList)).join(", ");

    // Totals and Recommendations from Engine
    const totalMarks = audit.passedList.reduce((a, b) => a + b.mark, 0) + (audit.failedList as any[]).reduce((a, b) => a + (b.mark || 0), 0);

    const isBlockedStatus = audit.summary.isOnLeave || ["ACADEMIC LEAVE", "DEFERMENT", "DEREGISTERED"].includes(audit.status);

    // Prepare the display values
    const displayTotal = isBlockedStatus ? "-" : totalMarks;
    const displayMean = isBlockedStatus ? "-" : parseFloat(audit.weightedMean).toFixed(2);

    rowData.push( audit.summary.totalExpected, displayTotal, displayMean, recomm, finalMatters );

    const row = sheet.getRow(rIdx);
    row.values = rowData;

    // STYLING
    row.eachCell((cell, colNum) => {
      cell.border = thinBorder;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.font = { size: 8, name: fontName };

      // Highlight REINSTATED names in Blue
      // if (colNum === 3 && (hasReturnHistory || repeatCount > 0)) cell.font = { color: { argb: "FF0000FF" }, bold: true, size: 8, name: fontName };

      if (colNum === totalCols - 1) {
        cell.protection = { locked: false }; // Keep editable
        const matterText = cell.value?.toString().toUpperCase() || "";
        // Highlight "FINANCIAL" matters in Yellow
        if (matterText.includes("ACADEMIC LEAVE") || matterText.includes("Academic Leave")) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        }
      }

      if (colNum === totalCols) {
        cell.protection = { locked: false }; // Keep editable
        cell.alignment = { horizontal: "left", vertical: "middle" };

        const matterText = cell.value?.toString().toUpperCase() || "";
        // Highlight "FINANCIAL" matters in Yellow
        if ( matterText.includes("FINANCIAL") || matterText.includes("Financial")) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
          cell.font = { bold: true, size: 8, name: fontName };
        }
      }

      if (colNum === 2 || colNum === 3) cell.alignment = { horizontal: "left", vertical: "middle" };

      if (colNum >= 5 && colNum < tuColIdx) {
        const val = cell.value?.toString() || "";
        if (resolvedStatus.isLocked) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
        else if (val === "INC" || val.includes("C")) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        else if (typeof cell.value === "number" && cell.value < passMark) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
          cell.font = { color: { argb: "FF9C0006" }, bold: true, size: 8, name: fontName };
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



