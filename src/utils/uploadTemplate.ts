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

  // Max scores shifted to start from index 0 (Column A)
  const dynamicMaxScores = [
    null, null, null, null, 20, 20, 20, 20, 10, 10, 10, 10, 30, q1Max, 20, 20, 20, 20, 70, 100, 100, 100, null,
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

  // 3. LOGO & HEADERS
  if (logoBuffer?.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    // Adjusted logo position slightly left since gutter is gone
    sheet.addImage(logoId, { tl: { col: 8, row: 0 }, ext: { width: 120, height: 90 } });
  }

  const centerBold = { alignment: { horizontal: "center" as const, vertical: "middle" as const }, font: { bold: true, name: fontName, underline: true } };

  sheet.mergeCells("F6:O6");
  sheet.getCell("F6").value = config.instName.toUpperCase();
  sheet.getCell("F6").style = { ...centerBold, font: { ...centerBold.font, size: 12 } };

  sheet.mergeCells("E7:P7");
  sheet.getCell("E7").value = `DEGREE: ${program?.name.toUpperCase() || "BACHELOR OF TECHNOLOGY"}`;
  sheet.getCell("E7").style = { ...centerBold, font: { ...centerBold.font } };

  const semTxt = semester === 1 ? "FIRST" : "SECOND";
  const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
  sheet.mergeCells("F8:O8");
  sheet.getCell("F8").value = `${yrTxt} YEAR ${semTxt} SEMESTER ${academicYear?.year || ""} ACADEMIC YEAR`;
  sheet.getCell("F8").style = { ...centerBold, font: { ...centerBold.font } };

  sheet.mergeCells("I10:K10");
  sheet.getCell("I10").value = "SCORESHEET";
  sheet.getCell("I10").style = { ...centerBold, font: { ...centerBold.font } };

  // Unit Info
  sheet.mergeCells("F12:G12");
  sheet.getCell("F12").value = "UNIT CODE:";
  sheet.mergeCells("H12:I12");
  const unitCodeCell = sheet.getCell("H12");
  unitCodeCell.value = unit?.code;
  unitCodeCell.style = { 
    alignment: { horizontal: "center", vertical: "middle" }, 
    font: { bold: true, name: fontName, size: fontSize, underline: true } 
  };
  sheet.mergeCells("L12:M12");
  sheet.getCell("L12").value = "UNIT TITLE:";
  sheet.getCell("N12").value = unit?.name.toUpperCase();
  sheet.getRow(12).font = { name: fontName, bold: true, size: fontSize };

  // 4. TABLE HEADERS (Starting from Column A = Index 1)
  // A=1, B=2, C=3, D=4 (S/N, REG, NAME, ATTEMPT)
  sheet.mergeCells("A14:A15"); sheet.mergeCells("B14:B15");
  sheet.mergeCells("C14:C15"); sheet.mergeCells("D14:D15");
  
  sheet.mergeCells("E14:H14"); sheet.getCell("E14").value = "CONTINUOUS ASSESSMENT TESTS";
  sheet.mergeCells("I14:L14"); sheet.getCell("I14").value = "ASSIGNMENTS";
  sheet.mergeCells("N14:S14"); sheet.getCell("N14").value = "END OF SEMESTER EXAMINATION";

  const headerRow = sheet.getRow(15);
  headerRow.height = 47;
  dynamicHeaders.forEach((val, i) => { headerRow.getCell(i + 1).value = val; });
  
  const scoreRow = sheet.getRow(16);
  dynamicMaxScores.forEach((val, i) => { if(val !== null) scoreRow.getCell(i + 1).value = val; });

  [14, 15, 16].forEach((rowNum) => {
    const row = sheet.getRow(rowNum);
    for (let c = 1; c <= 23; c++) {
      const cell = row.getCell(c);
      cell.font = { name: fontName, bold: true, size: 9 };
      cell.border = { ...thinBorder, bottom: rowNum === 16 ? { style: "double" } : thinBorder.bottom };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      
      if (rowNum === 14 && [5, 9, 14].includes(c)) cell.fill = greyColor;
      if ((rowNum === 15 && c >= 4) || (rowNum === 14 && c === 4)) cell.alignment.textRotation = 90;
      
      if (c <= 3) {
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

      row.getCell(1).value = studentIndex + 1;
      row.getCell(2).value = student.regNo;
      row.getCell(3).value = student.name.toUpperCase();
      row.getCell(3).font = { name: fontName, size: 8 };
      row.getCell(4).value = isSupp ? "Supp" : "1st";

      if (isSupp && prevMark) {
        row.getCell(5).value = prevMark.cat1Raw; row.getCell(6).value = prevMark.cat2Raw; 
        row.getCell(7).value = prevMark.cat3Raw; row.getCell(9).value = prevMark.assgnt1Raw;
      }
    }

    for (let c = 1; c <= 23; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;
      cell.alignment = { vertical: "middle" };

      if (isSupp && c >= 5 && c <= 13) {
        cell.fill = greyColor;
        cell.protection = { locked: true };
      } else {
        if (c === 8 || c === 12) cell.fill = pinkColor; 
        if (c === 13) cell.fill = greyColor; 
        if (c >= 14 && c <= 19) cell.fill = purpleColor; 
        
        if ([2, 3, 4, 5, 6, 7, 9, 10, 11, 14, 15, 16, 17, 18, 21].includes(c)) {
            cell.protection = { locked: false };
        }
      }
    }

    const validate = (range: string[], max: number) => {
      range.forEach(col => {
        sheet.getCell(`${col}${r}`).dataValidation = { type: "decimal", operator: "between", formulae: [0, max], showErrorMessage: true, errorTitle: "Invalid Mark", error: `Mark must be 0-${max}` };
      });
    };
    validate(["E", "F", "G", "O", "P", "Q", "R"], 20); validate(["I", "J", "K"], 10); validate(["N"], q1Max);

    // FORMULAS (Corrected for No Gutter)
    const isRowEmpty = `ISBLANK(B${r})`; // Checking Reg No in Col B
    sheet.getCell(`H${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(E${r}:G${r})>=1, AVERAGE(E${r}:G${r}), ""))` };
    sheet.getCell(`L${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(I${r}:K${r})>=1, AVERAGE(I${r}:K${r}), ""))` };
    sheet.getCell(`M${r}`).value = { formula: `IF(${isRowEmpty}, "", SUM(H${r}, L${r}))` };

    const q1 = `N${r}`;
    const others = `O${r}:R${r}`;
    const all = `N${r}:R${r}`;
    
    const examFormula = isMandatoryQ1Mode 
      ? `${q1} + IFERROR(LARGE(${others}, 1), 0) + IFERROR(LARGE(${others}, 2), 0)`
      : `SUM(${all})`;
   
    sheet.getCell(`S${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(COUNT(${all})>=1, ${examFormula}, ""))` };
    sheet.getCell(`T${r}`).value = { formula: `IF(${isRowEmpty}, "", ROUND(SUM(M${r}, S${r}), 0))` };
    sheet.getCell(`V${r}`).value = { formula: `IF(${isRowEmpty}, "", IF($D${r}="Supp", MIN(40, IF(U${r}<>"", U${r}, T${r})), IF(U${r}<>"", U${r}, T${r})))` };
    sheet.getCell(`W${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(V${r}<40, "F", IF(D${r}="Supp", "D", IF(V${r}>=70, "A", IF(V${r}>=60, "B", IF(V${r}>=50, "C", "D"))))))` };

    [`H${r}`, `L${r}`, `M${r}`, `S${r}`, `T${r}`, `V${r}`, `W${r}`].forEach(ref => sheet.getCell(ref).protection = { locked: true });
  }

  // 6. CONDITIONAL FORMATTING
  sheet.addConditionalFormatting({
    ref: `E17:R${endRow}`,
    rules: [{
      priority: 1,
      type: "expression",
      formulae: [`AND(NOT(ISBLANK($B17)), ISBLANK(E17))`],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFA6C9" } } },
    }],
  });

  // Thick Borders (A14 to W_endRow)
  for (let r = 14; r <= endRow; r++) {
    for (let c = 1; c <= 23; c++) {
      const cell = sheet.getCell(r, c);
      cell.border = {
        ...cell.border,
        left: c === 1 ? { style: "thick" } : cell.border?.left,
        right: c === 23 ? { style: "thick" } : cell.border?.right,
        top: r === 14 ? { style: "thick" } : cell.border?.top,
        bottom: r === endRow ? { style: "thick" } : cell.border?.bottom,
      };
    }
  }

  // 7. FINAL FORMATTING
  sheet.getColumn("A").width = 5; 
  sheet.getColumn("B").width = 22; 
  sheet.getColumn("C").width = 35;
  ["M", "T", "U", "V", "W", "S"].forEach(col => sheet.getColumn(col).width = 10.8);
  ["E", "F", "G", "H", "I", "J", "K", "L", "N", "O", "P", "Q", "R"].forEach(col => sheet.getColumn(col).width = 6.5);

  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true, formatCells: true });
  return await workbook.xlsx.writeBuffer();
};
