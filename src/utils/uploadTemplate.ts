
// src/utils/uploadTemplate.ts
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import mongoose from "mongoose";
import * as ExcelJS from 'exceljs';
import config from "../config/config";

export const MARKS_UPLOAD_HEADERS = [
  "S/N", "REG. NO.", "NAME", "ATTEMPT",
  "CAT 1 Out of", "CAT 2 Out of", "CAT3 Out of", "TOTAL",
  "Assgnt 1 Out of", "Assgnt 2 Out of", "Assgnt 3 Out of", "TOTAL",
  "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
  "Q1 out of", "Q2 out of", "Q3 out of", "Q4 out Of", "Q5 out of", "TOTAL EXAM OUT OF",
  "INTERNAL EXAMINER MARKS /100", "EXTERNAL EXAMINER MARKS /100", "AGREED MARKS /100", "GRADE",
];

export const MARKS_MAX_SCORES_ROW = [
  null, null, null, null,
  20, 20, 20, 20, 10, 10, 10, 10, 30,
  10, 20, 20, 20, 20, 70, 100, 100, 100, null
];

export const generateFullScoresheetTemplate = async (
  programId: mongoose.Types.ObjectId,
  unitId: mongoose.Types.ObjectId,
  yearOfStudy: number,
  semester: number,
  academicYearId: mongoose.Types.ObjectId,
  logoBuffer: any
): Promise<any> => {
  const [program, unit, academicYear] = await Promise.all([
    Program.findById(programId).lean(),
    Unit.findById(unitId).lean(),
    AcademicYear.findById(academicYearId).lean(),
  ]);

  const workbook = new ExcelJS.Workbook();
  const sheetName = `${unit?.code || ''} ${unit?.name?.substring(0, 15) || ''}`;
  const sheet = workbook.addWorksheet(sheetName.trim());

  // Set global font to Book Antiqua
  const fontName = 'Book Antiqua';
  const fontSize = 10;
  
  const greyColor = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE0E0E0' } }; 
  const pinkColor = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFFA6C9' } }; 
  const purpleColor = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFC5A3FF' } };

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
  };

  // 1. LOGO
  if (logoBuffer && logoBuffer.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: 'png' });
    sheet.addImage(logoId, { tl: { col: 8, row: 0 }, ext: { width: 120, height: 90 } });
  }

  // 2. UNIVERSITY HEADERS
  const centerBold = { alignment: { horizontal: 'center' as const, vertical: 'middle' as const }, font: { bold: true, name: fontName, underline: true} };
  
  sheet.mergeCells('D6:Q6');
  sheet.getCell('D6').value = config.instName.toUpperCase();
  sheet.getCell('D6').style = { ...centerBold, font: { ...centerBold.font, size: 12 } };

  sheet.mergeCells('E7:P7');
  sheet.getCell('E7').value = `DEGREE: ${program?.name.toUpperCase() || 'BACHELOR OF TECHNOLOGY'}`;
sheet.getCell('E7').style = {...centerBold, font:{...centerBold.font}};

  const semesterText = semester === 1 ? "FIRST" : "SECOND";
  const yearText = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
  sheet.mergeCells('A8:W8');
  sheet.getCell('A8').value = `${yearText} YEAR ${semesterText} SEMESTER ${academicYear?.year || ''} ACADEMIC YEAR`;
  sheet.getCell('A8').style = {...centerBold, font:{...centerBold.font}};

  sheet.mergeCells('H10:L10');
  sheet.getCell('H10').value = 'SCORESHEET';
  sheet.getCell('H10').style = {...centerBold, font:{...centerBold.font}};

  sheet.mergeCells('G12:H12');
  sheet.getCell('G12').value = 'UNIT CODE:';
  sheet.mergeCells('I12:J12');
  sheet.getCell('I12').value = unit?.code;
  sheet.getCell('I12').style = {...centerBold, font:{...centerBold.font}};
  sheet.mergeCells('M12:N12');
  sheet.getCell('M12').value = 'UNIT TITLE:';
  sheet.getCell('O12').value = unit?.name.toUpperCase();
  sheet.getRow(12).font = { name: fontName, bold: true, size: fontSize};

  // 3. TABLE HEADERS (Row 14 & 15)
  sheet.mergeCells('A14:A15'); 
  sheet.mergeCells('B14:B15'); 
  sheet.mergeCells('C14:C15'); 
  sheet.mergeCells('D14:D15'); 

  sheet.mergeCells('E14:H14'); 
  sheet.getCell('E14').value = 'CONTINUOUS ASSESSMENT TESTS';
  sheet.getCell('E14').fill = greyColor;

  sheet.mergeCells('I14:L14'); 
  sheet.getCell('I14').value = 'ASSIGNMENTS';
  sheet.getCell('I14').fill = greyColor;

  sheet.mergeCells('N14:S14'); 
  sheet.getCell('N14').value = 'END OF SEMESTER EXAMINATION';
  sheet.getCell('N14').fill = greyColor;

  const subHeaders = [
    "S/N", "REG. NO.", "NAME", "ATTEMPT", 
    "CAT 1 Out of", "CAT 2 Out of", "CAT 3 Out of", "TOTAL",
    "Assgnt 1 Out of", "Assgnt 2 Out of", "Assgnt 3 Out of", "TOTAL",
    "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
    "Q1 out of", "Q2 out of", "Q3 out of", "Q4 out of", "Q5 out of", "TOTAL EXAM OUT OF",
    "INTERNAL EXAMINER MARKS /100", "EXTERNAL EXAMINER MARKS /100", "AGREED MARKS /100", "GRADE"
  ];
  sheet.getRow(15).values = subHeaders;
  sheet.getRow(16).values = MARKS_MAX_SCORES_ROW;

  // Header Styling Logic
  [14, 15, 16].forEach(rowNum => {
    const row = sheet.getRow(rowNum);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber <= 23) {
        cell.font = { name: fontName, bold: true, size: 9 };
        cell.border = { ...thinBorder };

        // D15 to W15 vertical + Merged Attempt D14
        if ((rowNum === 15 && colNumber >= 4) || (rowNum === 14 && colNumber === 4)) {
          cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center', wrapText: true };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        }

        // Section header logic
        if (rowNum === 14 && (colNumber === 5 || colNumber === 9 || colNumber === 14)) {
            cell.fill = greyColor;
        }

        // Remove the internal border for merged cells A, B, C (Rows 14 and 15)
       if (colNumber <= 3) {
            if (rowNum === 14) cell.border = { ...cell.border, bottom: undefined };
            if (rowNum === 15) cell.border = { ...cell.border, top: undefined };
        }
      }
    });
  });

  sheet.getRow(14).height = 30;
  sheet.getRow(15).height = 55;

  // 4. DATA ROWS
  for (let r = 17; r <= 50; r++) {
    const row = sheet.getRow(r);
    // Applying the style to the WHOLE ROW ensures it picks up correctly
    row.font = { name: fontName, size: fontSize };
    row.alignment = { vertical: 'middle', wrapText: true };
    // row.height = 20; // Added some height for "padding" feel

    for (let c = 1; c <= 23; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;
      
      if (c === 8 || c === 12) cell.fill = pinkColor;
      if (c === 13) cell.fill = greyColor;
      if (c >= 14 && c <= 19) cell.fill = purpleColor;
    }
  }

  // 5. SAMPLE DATA
  sheet.getRow(17).values = [1, 'T056-01-0049/2020', 'Gregory Onyango OWINY', '1st', 13, 16, null, null, 6, null, null, null, null, 4, 5, 8, 14, null, null, 52, null, null, 'C'];

  // 6. MEDIUM OUTWARD BORDERS
  for (let c = 1; c <= 23; c++) {
    sheet.getCell(14, c).border = { ...sheet.getCell(14, c).border, top: { style: 'medium' } };
    sheet.getCell(50, c).border = { ...sheet.getCell(50, c).border, bottom: { style: 'medium' } };
  }
  for (let r = 14; r <= 50; r++) {
    sheet.getCell(r, 1).border = { ...sheet.getCell(r, 1).border, left: { style: 'medium' } };
    sheet.getCell(r, 23).border = { ...sheet.getCell(r, 23).border, right: { style: 'medium' } };
  }

  // Column Widths
  sheet.getColumn('A').width = 7;
  sheet.getColumn('B').width = 20;
  sheet.getColumn('C').width = 35;
  ['M','T','U'].forEach(col => {
      sheet.getColumn(col).width = 10;
  });
  // Setting small columns for scores
  ['E','F','G','I','J','K','N','O','P','Q','R','S'].forEach(col => {
      sheet.getColumn(col).width = 6;
  });

  const rawData = await workbook.xlsx.writeBuffer();
  return rawData;
};

