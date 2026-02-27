// // serverside/src/utils/consolidatedMS.ts
import * as ExcelJS from "exceljs";
import config from "../config/config";
import InstitutionSettings from "../models/InstitutionSettings";

export interface ConsolidatedData {
  programName: string; academicYear: string; yearOfStudy: number;
  students: any[]; marks: any[];
  offeredUnits: { code: string; name: string }[];
  logoBuffer?: any; institutionId?: string;
}

export const generateConsolidatedMarkSheet = async ( data: ConsolidatedData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, students, marks, offeredUnits, logoBuffer, institutionId } = data;

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
    cell.style = { alignment: { horizontal: "center", vertical: "middle" }, font: { bold: true, name: fontName, size: fontSize }};
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
      font: { bold: true, size: 8, name: fontName },
      // border: thinBorder,
      border: doubleBottomBorder,
    };
  });

  offeredUnits.forEach((unit, i) => {
    const colIdx = 5 + i;
    sheet.getCell(startRow, colIdx).value = (i + 1).toString();
    sheet.getCell(startRow, colIdx).style = { alignment: { horizontal: "center", vertical: "middle" }, font: { bold: true, size: 8 }, border: thinBorder };
    sheet.getCell(subRow, colIdx).value = unit.code;
    sheet.getCell(subRow, colIdx).style = { alignment: { horizontal: "center", vertical: "middle", textRotation: 90 }, font: { bold: true, size: 8, name: fontName }, border: thinBorder };
  });

  // 3. STUDENT DATA (Row 11+)  
  const statsSummary = { PASS: 0, SUPPLEMENTARY: 0, "REPEAT YEAR": 0, "STAY OUT": 0, SPECIAL: 0, INCOMPLETE: 0, DISCONTINUED: 0, DEREGISTERED: 0 };
  students.sort((a, b) => (a.regNo || "").localeCompare(b.regNo || "")).forEach((student, index) => {
    const rIdx = 11 + index;
    const sId = student.id?.toString() || student._id?.toString();

    let totalMarks = 0;
    let unitCount = 0;
    let suppCount = 0;
    let specialCount = 0;
    let incCount = 0;

    const rowData: any[] = [index + 1, student.regNo, student.name, "B/S"];

    offeredUnits.forEach((unit) => {
      const markObj = marks.find((m) => (m.student?._id?.toString() || m.student?.toString()) === sId && m.programUnit?.unit?.code === unit.code);

      if (markObj) {
        // const isSpecial = markObj.isSpecial || markObj.attempt === "special" || (markObj.remarks?.toLowerCase().includes("financial"));
        const isSpecial = markObj.isSpecial || markObj.attempt === "special" || (markObj.remarks?.toLowerCase().includes("special"));
        const markValue = markObj.agreedMark ?? 0;

        // "C" Notation Logic
        const displayMark = (markObj.examTotal70 === 0 || !markObj.examTotal70) ? `${markValue}C` : markValue;
        rowData.push(displayMark);

        if (isSpecial) specialCount++;
        else if (markValue < passMark) suppCount++; totalMarks += markValue; unitCount++;
      } else {
        incCount++;
        rowData.push("INC"); 
      }
    });  

    // Recommendation Logic (ENG RULES)
    const mean = unitCount > 0 ? totalMarks / unitCount : 0;
    const failRate = (suppCount + incCount) / offeredUnits.length;

    let recommText = "";
    if (student.status === "DEREGISTERED") recommText = "DEREGISTERED";
    else if (student.status === "ACADEMIC LEAVE") recommText = "ACADEMIC LEAVE";
    else if (student.status === "DISCONTINUED") recommText = "DISCONTINUED";
    else if (failRate >= 0.5 || mean < 40) recommText = "REPEAT YEAR";
    else if (failRate > 0.33) recommText = "STAY OUT (RETAKE)";
    else {
      const parts = [];
      if (suppCount > 0) parts.push(`SUPP ${suppCount}`);
      if (specialCount > 0) parts.push(`SPEC ${specialCount}`);
      if (incCount > 0) parts.push(`INC ${incCount}`);
      recommText = parts.length > 0 ? parts.join("; ") : "PASS";
    }

    // Update Overall Stats
    if (recommText.includes("PASS")) statsSummary.PASS++;
    else if (recommText.includes("SUPP")) statsSummary.SUPPLEMENTARY++;
    else if (recommText.includes("SPEC")) statsSummary.SPECIAL++;
    else if (recommText.includes("REPEAT")) statsSummary["REPEAT YEAR"]++;
    else if (recommText.includes("STAY OUT")) statsSummary["STAY OUT"]++;
    else if (recommText.includes("INC")) statsSummary.INCOMPLETE++;
    else if (recommText.includes("DEREG")) statsSummary.DEREGISTERED++;

    // Extract cleaned grounds for Student Matters
    const grounds = student.reasons
      ?.filter((r: string) => r.toLowerCase().includes("ground") || r.toLowerCase().includes("special") || r.toLowerCase().includes("leave"))
      .map((r: string) => r.split(":").pop()?.trim())
      .join(", ") || "";

    rowData.push(unitCount, totalMarks, parseFloat(mean.toFixed(2)), recommText, grounds);

    const row = sheet.getRow(rIdx);
    row.values = rowData;

    // --- STYLING LOGIC ---
    row.eachCell((cell, colNum) => {
      cell.border = thinBorder;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      // new addition
      cell.font = { size: 9, name: fontName };

      // Highlight Logic for Unit Columns (5 onwards)
      if (colNum >= 5 && colNum < tuColIdx) {
        const cellValue = cell.value?.toString() || "";
        const isC = cellValue.includes("C");
        const numericMark = parseFloat(cellValue.replace("C", ""));
        
        // Find matching mark object to check Special flag
        const unitCode = offeredUnits[colNum - 5].code;
        const mObj = marks.find((m) => (m.student?._id?.toString() || m.student?.toString()) === sId && m.programUnit?.unit?.code === unitCode);
        const isConfirmedSpecial = mObj?.isSpecial || mObj?.attempt === "special";

        if (isConfirmedSpecial) {
          // Rule: YELLOW for Specials
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        } else if (numericMark < passMark || isC || cellValue === "INC") {
          // Rule: RED for Failures or "C" without Special Approval (Supp)
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
          cell.font = { color: { argb: "FF9C0006" }, bold: true };
        }
      }

      // Make Student Matters Column Unlocked (Editable)
      if (colNum === totalCols) cell.protection = { locked: false };      
    });
  });
  
  const lastDataRow = 10 + students.length;

  // 4. UNIT STATISTICS (With Thick Borders)
  const statsStart = lastDataRow + 2;
  const statsLabels = [ "Mean", "Standard Deviation", "Maximum", "Minimum", "No. of Candidates", "No. of Passes", "No. of Fails", "No. of Blanks" ];

  statsLabels.forEach((label, i) => {
    const rIdx = statsStart + i; const r = sheet.getRow(rIdx); r.getCell(3).value = label;
    r.getCell(3).font = { bold: true, size: 9 }; r.getCell(3).border = thinBorder;

    offeredUnits.forEach((_, uIdx) => {
      const colIdx = 5 + uIdx;
      const colLetter = sheet.getColumn(colIdx).letter;
      const cell = r.getCell(colIdx);
      const range = `${colLetter}11:${colLetter}${lastDataRow}`;
      cell.border = thinBorder;
      cell.numFmt = "0.00";
      if (label === "No. of Passes") cell.value = { formula: `COUNTIF(${range}, ">=${passMark}")` };
      else if (label === "No. of Fails") cell.value = { formula: `COUNTIF(${range}, "<${passMark}")` };
      else if (label === "No. of Blanks") cell.value = { formula: `COUNTIF(${range}, "INC")` };
      else {
        let func = "";
        switch(label) {
          case "Mean": func = "AVERAGE"; break;
          case "Standard Deviation": func = "STDEV.P"; break; // Using Population Std Dev
          case "Maximum": func = "MAX"; break;
          case "Minimum": func = "MIN"; break;
          case "No. of Candidates": func = "COUNT"; break;
        }
        cell.value = { formula: `${func}(${range})` };
      }
    });

    // Thick borders for Stats Table
    r.getCell(3).border = { ...r.getCell(3).border, left: { style: "thick" } };
    r.getCell(tuColIdx - 1).border = { ...r.getCell(tuColIdx - 1).border, right: { style: "thick" }};
    if (i === 0)
      r.eachCell({ includeEmpty: false }, (c) => { if ( typeof c.col === "string" ? parseInt(c.col) >= 3 && parseInt(c.col) < tuColIdx : c.col >= 3 && c.col < tuColIdx ) c.border = { ...c.border, top: { style: "thick" } };
      });
    if (i === statsLabels.length - 1)
      r.eachCell({ includeEmpty: false }, (c) => { if ( typeof c.col === "string" ? parseInt(c.col) >= 3 && parseInt(c.col) < tuColIdx : c.col >= 3 && c.col < tuColIdx ) c.border = { ...c.border, bottom: { style: "thick" } };
      });
  });

  // 5. SUMMARY TABLE (With Thick Borders)
  // const summaryStart = statsStart + 10;
  const summaryStart = lastDataRow + 12;
  sheet.getCell(`B${summaryStart}`).value = "SUMMARY";
  sheet.getCell(`B${summaryStart}`).font = { bold: true, size: 10, underline: true };

  // const summaryMap: any = { PASS: 0, SUPPLEMENTARY: 0, FAIL: 0, SPECIAL: 0, INCOMPLETE: 0 };
  // students.forEach((s) => {
  //   const st = (s.status || "").toUpperCase();
  //   if (st.match(/PASS|GOOD/)) summaryMap.PASS++;
  //   else if (st.includes("SUPP")) summaryMap.SUPPLEMENTARY++;
  //   else if (st.includes("FAIL")) summaryMap.FAIL++;
  //   else if (st.includes("SPEC")) summaryMap.SPECIAL++;
  //   else summaryMap.INCOMPLETE++;
  // });

  // Object.entries(summaryMap).forEach(([label, count], i) => {
  //   const rIdx = summaryStart + 1 + i; const r = sheet.getRow(rIdx); const c1 = r.getCell(2); 
  //   const c2 = r.getCell(3); c1.value = label; c2.value = count as number;
  //   c1.border = { ...thinBorder, left: { style: "thick" } }; c2.border = { ...thinBorder, right: { style: "thick" } };
  //   if (i === 0) { c1.border.top = { style: "thick" }; c2.border.top = { style: "thick" }; }
  //   if (i === Object.keys(summaryMap).length - 1) { c1.border.bottom = { style: "thick" }; c2.border.bottom = { style: "thick" };}
  // });
  const summaryData = { "PASS": 0, "SUPPLEMENTARY": 0, "REPEAT YEAR": 0, "STAY OUT": 0, "SPECIAL": 0, "INCOMPLETE": 0, "DEREGISTERED/DISC": 0 };
  sheet.getColumn(tuColIdx + 3).eachCell({ includeEmpty: false }, (cell, rowNum) => {
    if (rowNum > 10 && rowNum <= lastDataRow) {
      const txt = cell.value?.toString().toUpperCase() || "";
      if (txt === "PASS") summaryData.PASS++;
      else if (txt.includes("SUPP")) summaryData.SUPPLEMENTARY++;
      else if (txt.includes("REPEAT")) summaryData["REPEAT YEAR"]++;
      else if (txt.includes("STAY OUT")) summaryData["STAY OUT"]++;
      else if (txt.includes("SPEC")) summaryData.SPECIAL++;
      else if (txt.includes("INC")) summaryData.INCOMPLETE++;
      else if (txt.includes("DEREG") || txt.includes("DISC")) summaryData["DEREGISTERED/DISC"]++;
    }
  });

  Object.entries(summaryData).forEach(([label, count], i) => {
    const rIdx = summaryStart + 1 + i;
    sheet.getCell(`B${rIdx}`).value = label;
    sheet.getCell(`C${rIdx}`).value = count;
    sheet.getRow(rIdx).getCell(2).border = thinBorder;
    sheet.getRow(rIdx).getCell(3).border = thinBorder;
  });

  // 6. OFFERED UNITS TABLE (With Thick Borders)
  const unitsStart = summaryStart + 8;
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
  sheet.getColumn(tuColIdx + 2).width = 7; sheet.getColumn(tuColIdx + 3).width = 25; sheet.getColumn(tuColIdx + 4).width = 20;

  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 10 }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0  };
  const result = await workbook.xlsx.writeBuffer();
  return Buffer.from(result as ArrayBuffer);
};;;


