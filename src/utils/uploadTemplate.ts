// src/utils/uploadTemplate.ts

  import Program from "../models/Program";
  import Unit from "../models/Unit";
  import AcademicYear from "../models/AcademicYear";
  import Student from "../models/Student";
  import Mark from "../models/Mark";
  import ProgramUnit from "../models/ProgramUnit";
  import InstitutionSettings from "../models/InstitutionSettings";
  import mongoose from "mongoose";
  import * as ExcelJS from "exceljs";
  import config from "../config/config";
  
  /**
   * GENERATE SCORESHEET TEMPLATE
   * Compliant with DeKUT ENG Rules (ENG 10.b/c/e & ENG 13.f)
   * Features: Theory/Lab/Workshop Normalization, Mandatory Q1 logic, and Supplementary Capping.
   */
  export const generateFullScoresheetTemplate = async (
    programId: mongoose.Types.ObjectId,
    unitId: mongoose.Types.ObjectId,
    yearOfStudy: number,
    semester: number,
    academicYearId: mongoose.Types.ObjectId,
    logoBuffer: any,
    examMode: string = "standard",
    unitType: "theory" | "lab" | "workshop" = "theory",
  ): Promise<any> => {
    const programUnit = await ProgramUnit.findOne({ program: programId, unit: unitId }).lean();
    const [program, unit, academicYear, eligibleStudents] = await Promise.all([
      Program.findById(programId).lean(),
      Unit.findById(unitId).lean(),
      AcademicYear.findById(academicYearId).lean(),
      Student.find({ program: programId, currentYearOfStudy: yearOfStudy, status: "active" }).sort({ regNo: 1 }).lean(),
    ]);
  
    const settings = await InstitutionSettings.findOne({ institution: program?.institution });
    if (!settings) throw new Error("Institution settings missing.");
  
    const previousMarks = programUnit ? await Mark.find({
      student: { $in: eligibleStudents.map((s) => s._id) },
      programUnit: programUnit._id,
    }).lean() : [];
  
    const isMandatoryQ1Mode = examMode === "mandatory_q1";
    const q1Max = isMandatoryQ1Mode ? 30 : 10;
  
    // Determine Weights per ENG 10(c)
    const weights = {
      practical: unitType === "lab" ? 15 : unitType === "workshop" ? 40 : 0,
      assignment: unitType === "lab" ? 5 : unitType === "theory" ? 10 : 0,
      tests: unitType === "lab" ? 10 : unitType === "theory" ? 20 : 0,
      exam: unitType === "workshop" ? 60 : 70,
    };
  
    const rawPaperMax = 70; 
    const caWeightTotal = 100 - weights.exam;
  
    // Header Definition
    const headers = [ "S/N", "REG. NO.", "NAME", "ATTEMPT", "CAT 1", "CAT 2", "CAT 3", "AVG CAT", "ASSGNT 1", "ASSGNT 2", "ASSGNT 3", "AVG ASSGNT", "PRACTICAL", `CA TOTAL (${caWeightTotal}%)`, `Q1 (/${q1Max})`, "Q2", "Q3", "Q4", "Q5", `EXAM (${weights.exam}%)`, "INTERNAL (/100)", "EXTERNAL (/100)", "AGREED (/100)", "GRADE", ];
  
    const dynamicMaxScores = [
      null, null, null, null,
      unitType === "workshop" ? null : settings.cat1Max,
      unitType === "workshop" ? null : settings.cat2Max,
      unitType === "workshop" ? null : settings.cat3Max,
      null,
      unitType === "workshop" ? null : settings.assignmentMax,
      unitType === "workshop" ? null : settings.assignmentMax,
      unitType === "workshop" ? null : settings.assignmentMax,
      null,
      unitType === "theory" ? null : settings.practicalMax,
      caWeightTotal,
      q1Max, 20, 20, 20, 20, weights.exam, 100, 100, 100, null,
    ];
  
    const workbook = new ExcelJS.Workbook();
    const sheetName = `${unit?.code || "SCORESHEET"}`;
    const sheet = workbook.addWorksheet(sheetName.trim());
  
    const fontName = "Book Antiqua";
    const fontSize = 9;
    const greyFill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE0E0E0" }, };
    const pinkColor = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFA6C9" }, };
    const purpleFill = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFC5A3FF" }, };
    const thinBorder: Partial<ExcelJS.Borders> = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }, };
  
    // Logo & Headers
    if (logoBuffer?.length > 0) {
      const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
      sheet.addImage(logoId, { tl: { col: 9, row: 0 }, ext: { width: 100, height: 80 }, });
    }
  
    const centerBold = { alignment: { horizontal: "center" as const, vertical: "middle" as const }, font: { bold: true, name: fontName, underline: true }, };
  
    sheet.mergeCells("E6:P6");
    sheet.getCell("E6").value = config.instName.toUpperCase();
    sheet.getCell("E6").style = { ...centerBold, font: { ...centerBold.font, size: 12 }, };
  
    sheet.mergeCells("D7:Q7");
    sheet.getCell("D7").value = `DEGREE: ${program?.name.toUpperCase() || "BACHELOR OF TECHNOLOGY"}`;
    sheet.getCell("D7").style = { ...centerBold, font: { ...centerBold.font } };
  
    const semTxt = semester === 1 ? "FIRST" : "SECOND";
    const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;
    sheet.mergeCells("D8:Q8");
    sheet.getCell("D8").value = `${yrTxt} YEAR ${semTxt} SEMESTER ${academicYear?.year || ""} ACADEMIC YEAR`;
    sheet.getCell("D8").style = { ...centerBold, font: { ...centerBold.font } };
  
    sheet.mergeCells("F10:O10");
    sheet.getCell("F10").value = `SCORESHEET FOR ${unit?.code || "UNITS"} (${unitType.toUpperCase()}) `;
    sheet.getCell("F10").style = { ...centerBold, font: { ...centerBold.font } };
  
    // Unit Info
    sheet.mergeCells("E12:G12");
    sheet.getCell("E12").value = "UNIT CODE:";
    sheet.mergeCells("H12:J12");
    const unitCodeCell = sheet.getCell("H12");
    unitCodeCell.value = unit?.code;
    unitCodeCell.style = {
      alignment: { horizontal: "center", vertical: "middle" },
      font: { bold: true, name: fontName, size: fontSize, underline: true },
    };
    sheet.mergeCells("L12:M12");
    sheet.getCell("L12").value = "UNIT TITLE:";
    sheet.getCell("N12").value = unit?.name.toUpperCase();
    sheet.getRow(12).font = { name: fontName, bold: true, size: fontSize };
  
    // 4. TABLE HEADERS (Starting from Column A = Index 1)
    // A=1, B=2, C=3, D=4 (S/N, REG, NAME, ATTEMPT)
    sheet.mergeCells("A14:A15");
    sheet.mergeCells("B14:B15");
    sheet.mergeCells("C14:C15");
    sheet.mergeCells("D14:D15");
  
    // Table Groups
    const groupHeaderRow = sheet.getRow(14);
    groupHeaderRow.height = 25;
    sheet.mergeCells("E14:H14");
    sheet.getCell("E14").value = "CONTINUOUS ASSESSMENT TESTS";
    sheet.mergeCells("I14:L14");
    sheet.getCell("I14").value = "ASSIGNMENTS";
    sheet.mergeCells("O14:T14");
    sheet.getCell("O14").value = "END OF SEMESTER EXAMINATION";
  
    const headerRow = sheet.getRow(15);
    headerRow.height = 47;
    headers.forEach((val, i) => { headerRow.getCell(i + 1).value = val; });
  
    const scoreRow = sheet.getRow(16);
    dynamicMaxScores.forEach((val, i) => { if (val !== null) scoreRow.getCell(i + 1).value = val; });
  
    [14, 15, 16].forEach((rowNum) => {
      const row = sheet.getRow(rowNum);
      for (let c = 1; c <= 24; c++) {
        const cell = row.getCell(c);
        cell.font = { name: fontName, bold: true, size: 8 };
        cell.border = { ...thinBorder, bottom: rowNum === 16 ? { style: "double" } : thinBorder.bottom, };
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true, };
        if (rowNum === 14 && [5, 9, 15].includes(c)) cell.fill = greyFill;
        if ((rowNum === 15 && c >= 4) || (rowNum === 14 && c === 4))
          cell.alignment.textRotation = 90;
  
        if (c <= 3) {
          if (rowNum === 14) cell.border = { ...cell.border, bottom: undefined };
          if (rowNum === 15) cell.border = { ...cell.border, top: undefined };
        }
        if (unitType === "workshop" && c >= 5 && c <= 12) cell.fill = greyFill;
        if (unitType === "theory" && c === 13) cell.fill = greyFill;
      }
    });
  
    // Data Rows
    const startRow = 17;
    // const endRow = startRow + Math.max(eligibleStudents.length + 10, 20);
    const endRow = startRow + Math.max(eligibleStudents.length + 10);
  
    for (let r = startRow; r <= endRow; r++) {
      const studentIdx = r - startRow; const student = eligibleStudents[studentIdx];
      const row = sheet.getRow(r); row.height = 13;
      row.font = { name: fontName, size: fontSize };

      let attemptLabel = "1st"; let isSupp = false; let isSpecial = false;

      if (student) {
        const prevMark = previousMarks.find( (m) => m.student.toString() === student._id.toString(), );
        if (prevMark) {
          if (prevMark.isSpecial || prevMark.attempt === "special") {
            attemptLabel = "Special"; isSpecial = true;
          } else if ( prevMark.agreedMark < settings.passMark || prevMark.attempt === "supplementary")
         { attemptLabel = "Supp"; isSupp = true; }
          else if (prevMark.attempt === "re-take") { attemptLabel = "Retake"; }
        }

        sheet.getRow(r).getCell(4).value = attemptLabel; sheet.getRow(r).getCell(1).value = r - 16;
        sheet.getRow(r).getCell(2).value = student.regNo; sheet.getRow(r).getCell(3).value = student.name.toUpperCase();
        sheet.getRow(r).getCell(3).font = { name: fontName, size: 8 };

        if (isSpecial && prevMark) {
          row.getCell(5).value = prevMark.cat1Raw; row.getCell(6).value = prevMark.cat2Raw;
          row.getCell(7).value = prevMark.cat3Raw; row.getCell(9).value = prevMark.assgnt1Raw;
          row.getCell(13).value = prevMark.practicalRaw;

          for (let c = 5; c <= 13; c++) { const cell = row.getCell(c); cell.fill = greyFill; cell.protection = { locked: true }; }
        }
      }

      for (let c = 1; c <= 24; c++) {
        const cell = row.getCell(c);
        cell.border = thinBorder;
        cell.alignment = { vertical: "middle" };
        if (student) {
          const isSupp = row.getCell(4).value === "Supp";
          if (isSupp && c >= 5 && c <= 13) {
            cell.fill = greyFill; cell.protection = { locked: true };
          } else if (unitType === "workshop" && c >= 5 && c <= 12) {
            cell.fill = greyFill; cell.protection = { locked: true };
          } else if (unitType === "theory" && c === 13) {
            cell.fill = greyFill; cell.protection = { locked: true };
          } else {
            if (c === 8 || c === 12) cell.fill = pinkColor;
            if (c === 13) cell.fill = greyFill;
            if (c >= 14 && c <= 19) cell.fill = purpleFill;
            if ([2, 3, 4, 5, 6, 7, 9, 10, 11, 14, 15, 16, 17, 18, 21].includes(c)) 
             { cell.protection = { locked: false }; }
          }
        } else {
          if (c === 8 || c === 12) cell.fill = pinkColor; if (c === 13) cell.fill = greyFill; 
          if (c >= 14 && c <= 19) cell.fill = purpleFill; cell.protection = { locked: c === 13 }; // Lock practical column if no student
        }
      }

      const isRowEmpty = `ISBLANK(B${r})`;

      // Formulas
      const catDivisor = settings.cat1Max || 1; // Prevent division by zero
      const assDivisor = settings.assignmentMax || 1;
      const pracDivisor = settings.practicalMax || 1;
      const catFormula = unitType === "workshop" ? "0" : `IFERROR((AVERAGE(E${r}:G${r})/${catDivisor})*${weights.tests}, 0)`;
      sheet.getCell(`H${r}`).value = { formula: `IF(${isRowEmpty}, "", ${catFormula})`,};

      const assFormula = unitType === "workshop" ? "0" : `IFERROR((AVERAGE(I${r}:K${r})/${assDivisor})*${weights.assignment}, 0)`;
      sheet.getCell(`L${r}`).value = { formula: `IF(${isRowEmpty}, "", ${assFormula})`,};

      const pracNorm = `IFERROR((M${r}/${pracDivisor})*${weights.practical}, 0)`;
      const caMultiplier = `IF(D${r}="Supp", 0, 1)`;
      sheet.getCell(`N${r}`).value = { formula: `IF(${isRowEmpty}, "", ROUND((H${r} + L${r} + ${pracNorm})*${caMultiplier}, 2))`,};
      // Exam Formula: Q1 + Best 2 or Best 3
      const q1 = `O${r}`;
      const others = `P${r}:S${r}`;
      // Standard: Best 3 others. Mandatory Q1: Best 2 others.
      const takeCount = isMandatoryQ1Mode ? 2 : 3;

      const examFormula = `${q1} + IFERROR(LARGE(${others}, 1), 0) + IFERROR(LARGE(${others}, 2), 0)` + (takeCount === 3 ? ` + IFERROR(LARGE(${others}, 3), 0)` : "");

      // Normalize raw score (out of 70) to Weight (70 or 60)
      const normalizedExamValue = `IFERROR(((${examFormula})/${rawPaperMax})*${weights.exam}, 0)`;
      sheet.getCell(`T${r}`).value = { formula: `IF(${isRowEmpty}, "", ROUND(${normalizedExamValue}, 2))`,};
      sheet.getCell(`U${r}`).value = { formula: `IF(${isRowEmpty}, "", ROUND(N${r}+T${r}, 0))`,};
      // --- FINAL AGREED MARK (Rule ENG 13.f) ---
      // If Supp: Cap at 40 and ignore CA. Else: N + T.
      // sheet.getCell(`W${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(D${r}="Supp", MIN(${settings.passMark}, T${r}), ROUND(N${r}+T${r}, 0)))`,};
      sheet.getCell(`W${r}`).value = { formula: `IF(${isRowEmpty}, "", IF(D${r}="Supp", MIN(${settings.passMark}, IF(V${r}<>"", V${r}, U${r})), IF(V${r}<>"", V${r}, U${r})))`, };
      // Grade Nesting
      const sortedScale = [...(settings.gradingScale || [])].sort((a, b) => a.min - b.min, );
      let gradeIfs = `"E"`;
      sortedScale.forEach((scale) => { gradeIfs = `IF(W${r}>=${scale.min}, "${scale.grade}", ${gradeIfs})`; });
      sheet.getCell(`X${r}`).value = { formula: `IF(${isRowEmpty}, "", ${gradeIfs})`, };

      // Cell Formatting & THE FIX FOR DATA VALIDATION
      const applyValidation = (cellAddr: string, maxVal: number) => {
        sheet.getCell(cellAddr).dataValidation = {
          type: "decimal", operator: "between",
          allowBlank: true, formulae: [0, maxVal],
          showErrorMessage: true, errorTitle: "Invalid Score",
          error: `Value must be between 0 and ${maxVal}`,
        };
      };

      for (let c = 1; c <= 24; c++) {
        const cell = row.getCell(c);
        cell.border = thinBorder;
        const isFormulaColumn = [8, 12, 14, 20, 21, 23, 24].includes(c);
        const isCAInput = c >= 5 && c <= 12; 
        const isPracticalInput = c === 13;

        // 1. Supp Students: Lock all CA input (CATs, Assgn, and Practical)
        if ( student && row.getCell(4).value === "Supp" && (isCAInput || isPracticalInput))
        { cell.fill = greyFill; cell.protection = { locked: true }; cell.value = 0;  }
        // 2. Mode-based Locking
        // THEORY: Lock Practical (13)
        else if (unitType === "theory" && isPracticalInput) { cell.fill = greyFill; cell.protection = { locked: true }; }
        // WORKSHOP: Lock CATs/Assignments (5-12) BUT UNLOCK Practical (13)
        else if (unitType === "workshop" && isCAInput) { cell.fill = greyFill; cell.protection = { locked: true }; }
        // 3. Formula Columns
        else if (isFormulaColumn) { cell.fill = greyFill; cell.protection = { locked: true }; }
        // 4. Special Students
        else if ( student && row.getCell(4).value === "Special" && (isCAInput || isPracticalInput))
         { cell.fill = greyFill; cell.protection = { locked: true }; }
         else {
          // Editable cells
          cell.protection = { locked: false };
          if ( isPracticalInput && (unitType === "workshop" || unitType === "lab"))
           { cell.fill = { type: "pattern", pattern: "none" }; }
          // Apply validation to input cells
          if (c === 5) applyValidation(`E${r}`, settings.cat1Max);
          if (c === 6) applyValidation(`F${r}`, settings.cat2Max);
          if (c === 7 && settings.cat3Max > 0) applyValidation(`G${r}`, settings.cat3Max);
          if (c === 9 && settings.assignmentMax > 0) applyValidation(`I${r}`, settings.assignmentMax);
          if (c === 13 && settings.practicalMax > 0) applyValidation(`M${r}`, settings.practicalMax);
          if (c === 15) applyValidation(`O${r}`, q1Max);
          if (c >= 16 && c <= 19) applyValidation(`${sheet.getColumn(c).letter}${r}`, 20);
        }
        if (c >= 15 && c <= 19) cell.fill = purpleFill;
      }
    }
  
    //  CONDITIONAL FORMATTING
    let activeColumnsExpression = "";
    if (unitType === "theory") {
      // Columns E:L (CATs/Assgn) and O:S (Exam Questions) are active. M (Practical) is ignored.
      activeColumnsExpression = "OR(AND(COLUMN(E17)>=5, COLUMN(E17)<=12), AND(COLUMN(E17)>=15, COLUMN(E17)<=19))";
    } else if (unitType === "lab") {
      // E:M (CATs/Assgn/Prac) and O:S (Exam Questions) are active.
      activeColumnsExpression = "OR(AND(COLUMN(E17)>=5, COLUMN(E17)<=13), AND(COLUMN(E17)>=15, COLUMN(E17)<=19))";
    } else if (unitType === "workshop") {
      // M (Practical) and O:S (Exam Questions) are active. CATs/Assgn are ignored.
      activeColumnsExpression = "OR(COLUMN(E17)=13, AND(COLUMN(E17)>=15, COLUMN(E17)<=19))";
    }

    sheet.addConditionalFormatting({
      ref: `E17:S${endRow}`, // Expanded range to cover Practical and Exam Questions
      rules: [
        {
          priority: 1, type: "expression",
          // The formula checks: Not Blank Student AND Blank Cell AND Column is Active for this Mode
          formulae: [ `AND(NOT(ISBLANK($B17)), ISBLANK(E17), ${activeColumnsExpression})`, ],
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFA6C9" },},},
        },
      ],
    });
  
    // Thick Borders (A14 to W_endRow)
    for (let r = 14; r <= endRow; r++) {
      for (let c = 1; c <= 24; c++) {
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
  
    // Final column sizing
    sheet.getColumn("A").width = 4; sheet.getColumn("B").width = 22; sheet.getColumn("C").width = 35; sheet.getColumn("N").width = 9;
  
    ["L", "M"].forEach((col) => (sheet.getColumn(col).width = 8));
  
    ["T", "U", "V", "W"].forEach((col) => (sheet.getColumn(col).width = 7.8));
    ["E", "F", "G", "H", "I", "J", "K", "L", "O", "P", "Q", "R", "S"].forEach((col) => (sheet.getColumn(col).width = 4.5), );
  
    sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
    return await workbook.xlsx.writeBuffer();
  };