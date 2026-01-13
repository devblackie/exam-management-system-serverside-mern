// src/utils/uploadTemplate.ts
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import mongoose from "mongoose";
import * as ExcelJS from "exceljs";
import config from "../config/config";

export const MARKS_UPLOAD_HEADERS = [
  "S/N",
  "REG. NO.",
  "NAME",
  "ATTEMPT",
  "CAT 1 Out of",
  "CAT 2 Out of",
  "CAT3 Out of",
  "TOTAL CATS",
  "Assgnt 1 Out of",
  "Assgnt 2 Out of",
  "Assgnt 3 Out of",
  "TOTAL ASSGNT",
  "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
  "Q1 out of",
  "Q2 out of",
  "Q3 out of",
  "Q4 out of",
  "Q5 out of",
  "TOTAL EXAM OUT OF",
  "INTERNAL EXAMINER MARKS /100",
  "EXTERNAL EXAMINER MARKS /100",
  "AGREED MARKS /100",
  "GRADE",
];

export const MARKS_MAX_SCORES_ROW = [
  null,
  null,
  null,
  null,
  20,
  20,
  20,
  20,
  10,
  10,
  10,
  10,
  30,
  10,
  20,
  20,
  20,
  20,
  70,
  100,
  100,
  100,
  null,
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
  const sheetName = `${unit?.code || ""} ${unit?.name?.substring(0, 15) || ""}`;
  const sheet = workbook.addWorksheet(sheetName.trim());

  // Set global font to Book Antiqua
  const fontName = "Book Antiqua";
  const fontSize = 10;

  const greyColor = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFE0E0E0" },
  };
  const pinkColor = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFFFA6C9" },
  };
  const purpleColor = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFC5A3FF" },
  };

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  // 1. LOGO
  if (logoBuffer && logoBuffer.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, {
      tl: { col: 8, row: 0 },
      ext: { width: 120, height: 90 },
    });
  }

  // 2. UNIVERSITY HEADERS
  const centerBold = {
    alignment: { horizontal: "center" as const, vertical: "middle" as const },
    font: { bold: true, name: fontName, underline: true },
  };

  sheet.mergeCells("D6:Q6");
  sheet.getCell("D6").value = config.instName.toUpperCase();
  sheet.getCell("D6").style = {
    ...centerBold,
    font: { ...centerBold.font, size: 12 },
  };

  sheet.mergeCells("E7:P7");
  sheet.getCell("E7").value = `DEGREE: ${
    program?.name.toUpperCase() || "BACHELOR OF TECHNOLOGY"
  }`;
  sheet.getCell("E7").style = { ...centerBold, font: { ...centerBold.font } };

  const semesterText = semester === 1 ? "FIRST" : "SECOND";
  const yearText =
    ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] ||
    `${yearOfStudy}TH`;
  sheet.mergeCells("A8:W8");
  sheet.getCell("A8").value = `${yearText} YEAR ${semesterText} SEMESTER ${
    academicYear?.year || ""
  } ACADEMIC YEAR`;
  sheet.getCell("A8").style = { ...centerBold, font: { ...centerBold.font } };

  sheet.mergeCells("H10:L10");
  sheet.getCell("H10").value = "SCORESHEET";
  sheet.getCell("H10").style = { ...centerBold, font: { ...centerBold.font } };

  sheet.mergeCells("G12:H12");
  sheet.getCell("G12").value = "UNIT CODE:";
  sheet.mergeCells("I12:J12");
  sheet.getCell("I12").value = unit?.code;
  sheet.getCell("I12").style = { ...centerBold, font: { ...centerBold.font } };
  sheet.mergeCells("M12:N12");
  sheet.getCell("M12").value = "UNIT TITLE:";
  sheet.getCell("O12").value = unit?.name.toUpperCase();
  sheet.getRow(12).font = { name: fontName, bold: true, size: fontSize };

  // 3. TABLE HEADERS (Row 14 & 15)
  sheet.mergeCells("A14:A15");
  sheet.mergeCells("B14:B15");
  sheet.mergeCells("C14:C15");
  sheet.mergeCells("D14:D15");

  sheet.mergeCells("E14:H14");
  sheet.getCell("E14").value = "CONTINUOUS ASSESSMENT TESTS";
  sheet.getCell("E14").fill = greyColor;

  sheet.mergeCells("I14:L14");
  sheet.getCell("I14").value = "ASSIGNMENTS";
  sheet.getCell("I14").fill = greyColor;

  sheet.mergeCells("N14:S14");
  sheet.getCell("N14").value = "END OF SEMESTER EXAMINATION";
  sheet.getCell("N14").fill = greyColor;

  sheet.getRow(15).values = MARKS_UPLOAD_HEADERS;
  sheet.getRow(16).values = MARKS_MAX_SCORES_ROW;

  // Header Styling Logic
  [14, 15, 16].forEach((rowNum) => {
    const row = sheet.getRow(rowNum);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber <= 23) {
        cell.font = { name: fontName, bold: true, size: 9 };
        cell.border = { ...thinBorder };

        // D15 to W15 vertical + Merged Attempt D14
        if (
          (rowNum === 15 && colNumber >= 4) ||
          (rowNum === 14 && colNumber === 4)
        ) {
          cell.alignment = {
            textRotation: 90,
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
          };
        } else {
          cell.alignment = {
            vertical: "middle",
            horizontal: "center",
            wrapText: true,
          };
        }

        // Section header logic
        if (
          rowNum === 14 &&
          (colNumber === 5 || colNumber === 9 || colNumber === 14)
        ) {
          cell.fill = greyColor;
        }

        // Remove the internal border for merged cells A, B, C (Rows 14 and 15)
        if (colNumber <= 3) {
          if (rowNum === 14)
            cell.border = { ...cell.border, bottom: undefined };
          if (rowNum === 15) cell.border = { ...cell.border, top: undefined };
        }
      }
    });
  });

  sheet.getRow(14).height = 30;
  sheet.getRow(15).height = 55;

  // 4. DATA ROWS
  const MAX_MANUAL_ROWS = 200;
  const startRow = 17;
  const endRow = startRow + MAX_MANUAL_ROWS - 1;

  // for (let r = 17; r <= 50; r++) {
  for (let r = startRow; r <= endRow; r++) {
    const row = sheet.getRow(r);
    // Applying the style to the WHOLE ROW ensures it picks up correctly
    row.font = { name: fontName, size: fontSize };
    row.alignment = { vertical: "middle", wrapText: true };
    // row.height = 20; // Added some height for "padding" feel

    for (let c = 1; c <= 23; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;

      if (c === 8 || c === 12) cell.fill = pinkColor;
      if (c === 13) cell.fill = greyColor;
      if (c >= 14 && c <= 19) cell.fill = purpleColor;
    }

    // // --- DATA VALIDATION LOGIC ---

    // // 1. CATs (Columns E, F, G) - Max 20
    // ['E', 'F', 'G'].forEach(col => {
    //   sheet.getCell(`${col}${r}`).dataValidation = {
    //     type: 'decimal',
    //     operator: 'between',
    //     allowBlank: true,
    //     formulae: [0, 20],
    //     showErrorMessage: true,
    //     errorTitle: 'Invalid Mark',
    //     error: 'The CAT mark must be between 0 and 20.'
    //   };
    // });

    // // 2. Assignments (Columns I, J, K) - Max 10
    // ['I', 'J', 'K'].forEach(col => {
    //   sheet.getCell(`${col}${r}`).dataValidation = {
    //     type: 'decimal',
    //     operator: 'between',
    //     allowBlank: true,
    //     formulae: [0, 10],
    //     showErrorMessage: true,
    //     errorTitle: 'Invalid Mark',
    //     error: 'The Assignment mark must be between 0 and 10.'
    //   };
    // });

    // // 3. Exam Questions (Columns N, O, P, Q, R) - Max 20
    // // (Note: Adjust these max values if your question weights vary)
    // ['O', 'P', 'Q', 'R'].forEach(col => {
    //   sheet.getCell(`${col}${r}`).dataValidation = {
    //     type: 'decimal',
    //     operator: 'between',
    //     allowBlank: true,
    //     formulae: [0, 20],
    //     showErrorMessage: true,
    //     errorTitle: 'Invalid Mark',
    //     error: 'The Exam Question mark must be between 0 and 20.'
    //   };
    // });

    //  sheet.getCell(`N${r}`).dataValidation = {
    //   type: 'decimal',
    //   operator: 'between',
    //   allowBlank: true,
    //   formulae: [0, 10],
    //   showErrorMessage: true,
    //   errorTitle: 'Invalid Mark',
    //   error: 'The Exam Question mark must be between 0 and 10.'
    // };

    // // 4. External Marks (Column U) - Max 100
    // sheet.getCell(`U${r}`).dataValidation = {
    //   type: 'decimal',
    //   operator: 'between',
    //   allowBlank: true,
    //   formulae: [0, 100],
    //   showErrorMessage: true,
    //   errorTitle: 'Invalid Mark',
    //   error: 'The External mark must be between 0 and 100.'
    // };

    // --- APPLY FORMULAS ---

    // FORMULAS WITH "GATEKEEPERS"
    // If Reg No (B) is empty, show nothing. Otherwise, calculate.
    const isRowEmpty = `ISBLANK(B${r})`;

    sheet.getCell(`B${r}`).value = {
      formula: `IF(ISBLANK(B${r}), "", UPPER(B${r}))`,
      result: undefined,
    };

    // CAT Average (Column H): Only if 2 or more inputs exist
    // Formula: =IF(COUNT(E17:G17)>=2, AVERAGE(E17:G17), "")
    const catRange = `E${r}:G${r}`;
    sheet.getCell(`H${r}`).value = {
      formula: `IF(${isRowEmpty}, "", IF(COUNT(${catRange})>=2, AVERAGE(${catRange}), ""))`,
      result: undefined,
    };

    // Assignment Average (Column L): If 1 or more inputs exist
    // Formula: =IF(COUNT(I17:K17)>=1, AVERAGE(I17:K17), "")
    const assgnRange = `I${r}:K${r}`;
    sheet.getCell(`L${r}`).value = {
      formula: `IF(${isRowEmpty}, "", IF(COUNT(${assgnRange})>=1, AVERAGE(${assgnRange}), ""))`,
      result: undefined,
    };

    // Grand Total (Column M): Sum of H and L
    sheet.getCell(`M${r}`).value = {
      formula: `IF(${isRowEmpty}, "", IF(AND(H${r}="", L${r}=""), "", SUM(H${r}, L${r})))`,
      result: undefined,
    };

    // Total Exam (Column S): Sum of N to R
    sheet.getCell(`S${r}`).value = {
      formula: `IF(${isRowEmpty}, "", IF(COUNT(N${r}:R${r})>=1, SUM(N${r}:R${r}), ""))`,
      result: undefined,
    };

    // T: INTERNAL EXAMINER MARKS (M + S) - ROUNDED
    // This sums the Coursework (30) and the Exam (70) to get /100
    // Formula: =IF(AND(M17="", S17=""), "", ROUND(SUM(M17, S17), 0))
    sheet.getCell(`T${r}`).value = {
      formula: `IF(${isRowEmpty}, "", ROUND(SUM(M${r}, S${r}), 0))`,
      result: undefined,
    };

    // U: EXTERNAL EXAMINER MARKS /100 (Usually manual input, but let's ensure it's formatted)
    // V: AGREED MARKS /100
    // Usually, Agreed Marks is the External mark if it exists, otherwise the Internal mark
    sheet.getCell(`V${r}`).value = {
      formula: `IF(${isRowEmpty}, "", IF(U${r}<>"", ROUND(U${r}, 0), T${r}))`,
      result: undefined,
    };

    // Data Validation
    const validate = (range: string[], max: number) => {
      range.forEach((c) => {
        sheet.getCell(`${c}${r}`).dataValidation = {
          type: "decimal",
          operator: "between",
          formulae: [0, max],
          showErrorMessage: true,
          errorTitle: "Invalid Mark",
          error: `Mark must be between 0 and ${max}`,
        };
      });
    };
    validate(["E", "F", "G"], 20);
    validate(["I", "J", "K"], 10);
    validate(["O", "P", "Q", "R"], 20);
    sheet.getCell(`N${r}`).dataValidation = {
      type: "decimal",
      operator: "between",
      formulae: [0, 10],
    };

    // Protection Logic
    [`H${r}`, `L${r}`, `M${r}`, `S${r}`, `T${r}`, `V${r}`].forEach(
      (ref) => (sheet.getCell(ref).protection = { locked: true })
    );
    [2, 3, 4, 5, 6, 7, 9, 10, 11, 14, 15, 16, 17, 18, 21].forEach(
      (c) => (row.getCell(c).protection = { locked: false })
    );

    // const formulaCells = [`H${r}`, `L${r}`, `M${r}`, `S${r}`, `T${r}`, `V${r}`];
    // formulaCells.forEach((ref) => {
    //   sheet.getCell(ref).protection = { locked: true };
    // });

    // // IMPORTANT: You must UNLOCK the cells where teachers need to type marks
    // // By default, Excel locks everything. We must explicitly open the input areas.
    // const inputCols = [2, 3, 4, 5, 6, 7, 9, 10, 11, 14, 15, 16, 17, 18, 21];
    // inputCols.forEach((colIndex) => {
    //   sheet.getRow(r).getCell(colIndex).protection = { locked: false };
    // });
  }

  // 5. CONDITIONAL FORMATTING
  sheet.addConditionalFormatting({
    ref: `E17:L${endRow}`,
    rules: [
      {
        priority: 1,
        type: "expression",
        formulae: [`AND(NOT(ISBLANK($B17)), ISBLANK(E17))`],
        style: {
          fill: {
            type: "pattern",
            pattern: "solid",
            // Note: Conditional formatting uses 'bgColor', but ExcelJS often
            // maps this to 'fgColor' internally. argb is correct here.
            bgColor: { argb: "FFFFA6C9" },
          },
        },
      },
    ],
  });

  // 5. SAMPLE DATA (Populating ONLY input fields to preserve formulas)
  const sampleRow = sheet.getRow(17);
  sampleRow.getCell(1).value = 1; // S/N
  sampleRow.getCell(2).value = "T056-01-0049/2020"; // REG NO
  sampleRow.getCell(3).value = "Gregory Onyango OWINY"; // NAME
  sampleRow.getCell(4).value = "1st"; // ATTEMPT

  // CAT Inputs (CAT 1 & 2) - Formula in H17 will calculate the average
  sampleRow.getCell(5).value = 13;
  sampleRow.getCell(6).value = 16;

  // Assignment Input (Assgnt 1) - Formula in L17 will calculate the average
  sampleRow.getCell(9).value = 6;

  // Exam Inputs (Q1 - Q4) - Formula in S17 will calculate the total
  sampleRow.getCell(14).value = 4;
  sampleRow.getCell(15).value = 5;
  sampleRow.getCell(16).value = 8;
  sampleRow.getCell(17).value = 14;

  // sampleRow.getCell(20).value = 52; // Internal Examiner Marks
  sampleRow.getCell(23).value = "C"; // Grade

  // 6. MEDIUM OUTWARD BORDERS
  for (let c = 1; c <= 23; c++) {
    sheet.getCell(14, c).border = {
      ...sheet.getCell(14, c).border,
      top: { style: "medium" },
    };
    sheet.getCell(endRow, c).border = {
      ...sheet.getCell(endRow, c).border,
      bottom: { style: "medium" },
    };
  }
  for (let r = 14; r <= endRow; r++) {
    sheet.getCell(r, 1).border = {
      ...sheet.getCell(r, 1).border,
      left: { style: "medium" },
    };
    sheet.getCell(r, 23).border = {
      ...sheet.getCell(r, 23).border,
      right: { style: "medium" },
    };
  }

  // Column Widths
  sheet.getColumn("A").width = 7;
  sheet.getColumn("B").width = 20;
  sheet.getColumn("C").width = 35;
  ["M", "T", "U"].forEach((col) => {
    sheet.getColumn(col).width = 10;
  });
  // Setting small columns for scores
  ["E", "F", "G", "I", "J", "K", "N", "O", "P", "Q", "R", "S"].forEach(
    (col) => {
      sheet.getColumn(col).width = 6;
    }
  );

  // 8. ENABLE SHEET PROTECTION
  // This makes the "locked" property actually work.
  // Users can still select cells, but they can't change formulas.
  sheet.protect("", {
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: true,
    formatColumns: true,
    formatRows: true,
  });

  const rawData = await workbook.xlsx.writeBuffer();
  return rawData;
};
