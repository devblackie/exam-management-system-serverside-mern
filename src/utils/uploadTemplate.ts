// src/utils/uploadTemplate.ts
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import Student from "../models/Student";
import Mark from "../models/Mark";
import ProgramUnit from "../models/ProgramUnit";
import mongoose from "mongoose";
import * as ExcelJS from "exceljs";
import config from "../config/config";

export const MARKS_UPLOAD_HEADERS = [
  "S/N", "REG. NO.", "NAME", "ATTEMPT",
  "CAT 1 Out of", "CAT 2 Out of", "CAT3 Out of", "TOTAL CATS",
  "Assgnt 1 Out of", "Assgnt 2 Out of", "Assgnt 3 Out of", "TOTAL ASSGNT",
  "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
  "Q2 out of", "Q3 out of", "Q4 out of", "Q5 out of", "TOTAL EXAM OUT OF",
  "INTERNAL EXAMINER MARKS /100", "EXTERNAL EXAMINER MARKS /100", "AGREED MARKS /100", "GRADE",
];

export const generateFullScoresheetTemplate = async (
  programId: mongoose.Types.ObjectId,
  unitId: mongoose.Types.ObjectId,
  yearOfStudy: number,
  semester: number,
  academicYearId: mongoose.Types.ObjectId,
  logoBuffer: any,
  examMode: string = "standard"
): Promise<any> => {
 
  const programUnit = await ProgramUnit.findOne({ program: programId, unit: unitId }).lean();
  const [program, unit, academicYear, eligibleStudents] = await Promise.all([
    Program.findById(programId).lean(),
    Unit.findById(unitId).lean(),
    AcademicYear.findById(academicYearId).lean(),
    Student.find({ program: programId, currentYearOfStudy: yearOfStudy, status: "active" }).sort({ regNo: 1 }).lean()
  ]);

  const isMandatoryQ1Mode = examMode === "mandatory_q1";
  const q1Max = isMandatoryQ1Mode ? 30 : 10;

  let previousMarks: any[] = [];
  if (programUnit) {
    previousMarks = await Mark.find({
      student: { $in: eligibleStudents.map(s => s._id) },
      programUnit: programUnit._id
    }).lean();
  }

  const dynamicHeaders = [...MARKS_UPLOAD_HEADERS];
  dynamicHeaders.splice(13, 0, `Q1 out of ${q1Max}`);

  // Shifted max scores: Index 0 is now null (Col A), Index 1 is null (Col B/SN)
  const dynamicMaxScores = [
    null, null, null, null, null, 20, 20, 20, 20, 10, 10, 10, 10, 30, q1Max, 20, 20, 20, 20, 70, 100, 100, 100, null,
  ];  

  const workbook = new ExcelJS.Workbook();
  const sheetName = `${unit?.code || ""} ${unit?.name?.substring(0, 15) || ""}`;
  const sheet = workbook.addWorksheet(sheetName.trim());


  const fontName = "Book Antiqua";
  const fontSize = 9;
  const greyColor = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE0E0E0" } };
  const pinkColor = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFA6C9" } };
  const purpleColor = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFC5A3FF" } };
  const thinBorder: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };

  // 3. LOGO & HEADERS (Shifted right for balance)
  if (logoBuffer?.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, { tl: { col: 9, row: 0 }, ext: { width: 120, height: 90 } });
  }

  const centerBold = { alignment: { horizontal: "center" as const, vertical: "middle" as const }, font: { bold: true, name: fontName, underline: true } };

  sheet.mergeCells("G6:P6");
  sheet.getCell("G6").value = config.instName.toUpperCase();
  sheet.getCell("G6").style = { ...centerBold, font: { ...centerBold.font, size: 12 } };

  sheet.mergeCells("F7:Q7");
  sheet.getCell("F7").value = `DEGREE: ${program?.name.toUpperCase() || "BACHELOR OF TECHNOLOGY"}`;
  sheet.getCell("F7").style = { ...centerBold, font: { ...centerBold.font } };

  const semTxt = semester === 1 ? "FIRST" : "SECOND";
  const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
  sheet.mergeCells("G8:P8");
  sheet.getCell("G8").value = `${yrTxt} YEAR ${semTxt} SEMESTER ${academicYear?.year || ""} ACADEMIC YEAR`;
  sheet.getCell("G8").style = { ...centerBold, font: { ...centerBold.font } };

  sheet.mergeCells("J10:L10");
  sheet.getCell("J10").value = "SCORESHEET";
  sheet.getCell("J10").style = { ...centerBold, font: { ...centerBold.font } };

  // Unit Info
  sheet.mergeCells("G12:H12");
  sheet.getCell("G12").value = "UNIT CODE:";
  sheet.mergeCells("I12:J12");
  const unitCodeCell = sheet.getCell("I12");
  unitCodeCell.value = unit?.code;
  // Underlining the Unit Code here:
  unitCodeCell.style = { 
    alignment: { horizontal: "center", vertical: "middle" }, 
    font: { bold: true, name: fontName, size: fontSize, underline: true } 
  };
  sheet.mergeCells("M12:N12");
  sheet.getCell("M12").value = "UNIT TITLE:";
  sheet.getCell("O12").value = unit?.name.toUpperCase();
  sheet.getRow(12).font = { name: fontName, bold: true, size: fontSize };

  // 4. TABLE HEADERS (Starting from Column B = Index 2)
  // B=2, C=3, D=4, E=5 (S/N, REG, NAME, ATTEMPT)
  sheet.mergeCells("B14:B15"); sheet.mergeCells("C14:C15");
  sheet.mergeCells("D14:D15"); sheet.mergeCells("E14:E15");
  
  sheet.mergeCells("F14:I14"); sheet.getCell("F14").value = "CONTINUOUS ASSESSMENT TESTS";
  sheet.mergeCells("J14:M14"); sheet.getCell("J14").value = "ASSIGNMENTS";
  sheet.mergeCells("O14:T14"); sheet.getCell("O14").value = "END OF SEMESTER EXAMINATION";

  // Assign header values starting from Column B
  const headerRow = sheet.getRow(15);
  headerRow.height=47;
  dynamicHeaders.forEach((val, i) => { headerRow.getCell(i + 2).value = val; });
  
  const scoreRow = sheet.getRow(16);
  dynamicMaxScores.forEach((val, i) => { if(val !== null) scoreRow.getCell(i + 1).value = val; });

  [14, 15, 16].forEach((rowNum) => {
    const row = sheet.getRow(rowNum);
    for (let c = 2; c <= 24; c++) {
      const cell = row.getCell(c);
      cell.font = { name: fontName, bold: true, size: 9 };
      cell.border = { ...thinBorder, bottom: rowNum === 16 ? { style: "double" } : thinBorder.bottom };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      
      if (rowNum === 14 && [6, 10, 15].includes(c)) cell.fill = greyColor;
      if ((rowNum === 15 && c >= 5) || (rowNum === 14 && c === 5)) cell.alignment.textRotation = 90;
      
      if (c <= 4) { // Merge aesthetics
        if (rowNum === 14) cell.border = { ...cell.border, bottom: undefined };
        if (rowNum === 15) cell.border = { ...cell.border, top: undefined };
      }
    }
  });

  // 5. DATA ROWS
  const startRow = 17;
  const endRow = startRow + Math.max(eligibleStudents.length + 15);

  for (let r = startRow; r <= endRow; r++) {
    const studentIndex = r - startRow;
    const student = eligibleStudents[studentIndex];
    const row = sheet.getRow(r);
    row.height = 13;
    row.font = { name: fontName, size: fontSize };

    let isSupp = false;
    if (student) {
      const prevMark = previousMarks.find(m => m.student.toString() === student._id.toString());
      isSupp = !!(prevMark && (prevMark.attempt === "supplementary" || prevMark.agreedMark < 40));

      row.getCell(2).value = studentIndex + 1;
      row.getCell(3).value = student.regNo;
      row.getCell(4).value = student.name.toUpperCase();
      row.getCell(4).font = { name: fontName, size: 8 };
      row.getCell(5).value = isSupp ? "Supp" : "1st";

      if (isSupp && prevMark) {
        row.getCell(6).value = prevMark.cat1Raw; row.getCell(7).value = prevMark.cat2Raw; 
        row.getCell(8).value = prevMark.cat3Raw; row.getCell(10).value = prevMark.assgnt1Raw;
      }
    }

    // Styling & Formulas Loop
    for (let c = 2; c <= 24; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;
      cell.alignment = { vertical: "middle" };

      // Conditional Logic for Locks/Colors
      if (isSupp && c >= 6 && c <= 14) {
        cell.fill = greyColor;
        cell.protection = { locked: true };
      } else {
        if (c === 9 || c === 13) cell.fill = pinkColor; // Totals CAT/Assgnt
        if (c === 14) cell.fill = greyColor; // Grand Total 30
        if (c >= 15 && c <= 20) cell.fill = purpleColor; // Exam Section
        
        // Unlock input cells
        if ([3, 4, 5, 6, 7, 8, 10, 11, 12, 15, 16, 17, 18, 19, 22].includes(c)) {
            cell.protection = { locked: false };
        }
      }
    }

    // Validation & Protection
    const validate = (range: string[], max: number) => {
      range.forEach(c => {
        sheet.getCell(`${c}${r}`).dataValidation = { type: "decimal", operator: "between", formulae: [0, max], showErrorMessage: true, errorTitle: "Invalid Mark", error: `Mark must be between 0 and ${max}` };
      });
    };
    validate(["F", "G", "H", "P", "Q", "R", "S"], 20); validate(["J", "K", "L"], 10); validate(["O"], q1Max);

    // Formulas (Shifted Column References)
    const isRowEmpty = `ISBLANK(C${r})`; // Checking Reg No in Col C
    sheet.getCell(`I${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(F${r}:H${r})>=1, AVERAGE(F${r}:H${r}), ""))` };
    sheet.getCell(`M${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(J${r}:L${r})>=1, AVERAGE(J${r}:L${r}), ""))` };
    sheet.getCell(`N${r}`).value = { formula: `IF(${isRowEmpty}, "", SUM(I${r}, M${r}))` };

    const q1 = `O${r}`;
    const others = `P${r}:S${r}`;
    const all = `O${r}:S${r}`;
    
    const examFormula = isMandatoryQ1Mode 
      ? `${q1} + IFERROR(LARGE(${others}, 1), 0) + IFERROR(LARGE(${others}, 2), 0)`
      : `SUM(${all})`;
   
    sheet.getCell(`T${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(${all})>=1, ${examFormula}, ""))` };
    sheet.getCell(`U${r}`).value = { formula: `IF(${isRowEmpty}, "", ROUND(SUM(N${r}, T${r}), 0))` };
    sheet.getCell(`W${r}`).value = { formula: `IF(${isRowEmpty}, "", IF($E${r}="Supp", MIN(40, IF(V${r}<>"", V${r}, U${r})), IF(V${r}<>"", V${r}, U${r})))` };
    sheet.getCell(`X${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(W${r}<40, "F", IF(E${r}="Supp", "D", IF(W${r}>=70, "A", IF(W${r}>=60, "B", IF(W${r}>=50, "C", "D"))))))` };

    // Protection for Formulas
    [`I${r}`, `M${r}`, `N${r}`,  `T${r}`, `U${r}`, `W${r}`, `X${r}`].forEach(ref => sheet.getCell(ref).protection = { locked: true });
    // ["F", "G", "H", "J", "K", "L", "O", "P", "Q", "R", "S", "V"].forEach(col => {
    //   row.getCell(col).protection = { locked: false };
    // });

  }

  // 6. CONDITIONAL FORMATTING (Apply once to the whole range)
  sheet.addConditionalFormatting({
    ref: `F17:S${endRow}`,
    rules: [{
      priority: 1,
      type: "expression",
      formulae: [`AND(NOT(ISBLANK($C17)), ISBLANK(F17))`],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFA6C9" } } },
    }],
  });

  // Thick Borders (B14 to X_endRow)
  for (let r = 14; r <= endRow; r++) {
    for (let c = 2; c <= 24; c++) {
      const cell = sheet.getCell(r, c);
      cell.border = {
        ...cell.border,
        left: c === 2 ? { style: "thick" } : cell.border?.left,
        right: c === 24 ? { style: "thick" } : cell.border?.right,
        top: r === 14 ? { style: "thick" } : cell.border?.top,
        bottom: r === endRow ? { style: "thick" } : cell.border?.bottom,
      };
    }
  }

  // 7. FINAL FORMATTING
  sheet.getColumn("A").width = 2; // Gutter
  sheet.getColumn("B").width = 5; 
  sheet.getColumn("C").width = 22; 
  sheet.getColumn("D").width = 35;
  ["N", "U", "V", "W", "X", "T"].forEach(col => sheet.getColumn(col).width = 10.8);
  ["F", "G", "H", "I", "J", "K", "L", "M", "O", "P", "Q", "R", "S"].forEach(col => sheet.getColumn(col).width = 6.5);

  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true, formatCells: true });
  return await workbook.xlsx.writeBuffer();
};
