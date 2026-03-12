// src/utils/directTemplate.ts
import * as ExcelJS from "exceljs";
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import Student from "../models/Student";
import config from "../config/config";
import mongoose from "mongoose";
import InstitutionSettings from "../models/InstitutionSettings";
import MarkDirect from "../models/MarkDirect";
import ProgramUnit from "../models/ProgramUnit";

export const generateDirectScoresheetTemplate = async (
  programId: mongoose.Types.ObjectId,
  unitId: mongoose.Types.ObjectId,
  yearOfStudy: number,
  semester: number,
  academicYearId: mongoose.Types.ObjectId,
  logoBuffer: any,
): Promise<any> => {
  const programUnit = await ProgramUnit.findOne({
    program: programId,
    unit: unitId,
  }).lean();
  const [program, unit, academicYear, eligibleStudents] = await Promise.all([
    Program.findById(programId).lean(),
    Unit.findById(unitId).lean(),
    AcademicYear.findById(academicYearId).lean(),
    Student.find({
      program: programId,
      currentYearOfStudy: yearOfStudy,
      status: "active",
    })
      .sort({ regNo: 1 })
      .lean(),
  ]);

  const settings = await InstitutionSettings.findOne({
    institution: program?.institution,
  });
  if (!settings) throw new Error("Institution settings missing.");

  const previousMarks = programUnit
    ? await MarkDirect.find({
        student: { $in: eligibleStudents.map((s) => s._id) },
        programUnit: programUnit._id,
      }).lean()
    : [];

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(`${unit?.code || "SCORESHEET"}`);

  // Styles
  const fontName = "Book Antiqua";
  const fontSize = 9;
  const greyFill = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFE0E0E0" },
  };
  const pinkColor = {
    type: "pattern" as const,
    pattern: "solid" as const,
    fgColor: { argb: "FFFFA6C9" },
  };
  const purpleFill = {
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

  // 1. Logo & Branding

  if (logoBuffer?.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, {
      tl: { col: 3, row: 0 },
      ext: { width: 100, height: 80 },
    });
  }

  const centerBold = {
    alignment: { horizontal: "center" as const, vertical: "middle" as const },
    font: { bold: true, name: fontName, underline: true },
  };

  sheet.mergeCells("C6:G6");
  sheet.getCell("C6").value = config.instName.toUpperCase();
  sheet.getCell("C6").style = {
    ...centerBold,
    font: { ...centerBold.font, size: 12 },
  };

  sheet.mergeCells("C7:G7");
  sheet.getCell("C7").value = `DEGREE: ${program?.name.toUpperCase()}`;
  sheet.getCell("C7").style = centerBold;

  const yrTxt =
    ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] ||
    `${yearOfStudy}TH`;
  sheet.mergeCells("C8:G8");
  sheet.getCell("C8").value =
    `${yrTxt} YEAR | SEMESTER ${semester} | ${academicYear?.year || ""} ACADEMIC YEAR`;
  sheet.getCell("C8").style = centerBold;

  sheet.mergeCells("C10:G10");
  sheet.getCell("C10").value = `SCORESHEET FOR: ${unit?.code.toUpperCase()}`;
  sheet.getCell("C10").style = {
    ...centerBold,
    font: { ...centerBold.font, size: 10 },
  };

  // Unit Info
  sheet.getCell("B12").value = "UNIT TITLE:";
  sheet.getCell("C12").value = unit?.name.toUpperCase();
  sheet.getCell("E12").value = "UNIT CODE:";
  sheet.getCell("F12").value = unit?.code;
  sheet.getRow(12).font = { name: fontName, bold: true, size: 9 };

  // sheet.mergeCells("A15:A16");
  // sheet.mergeCells("B15:B16");
  // sheet.mergeCells("C15:C16");
  // sheet.mergeCells("D15:D16");

  // Header merging for static columns
  ["A", "B", "C", "D", "J"].forEach((col) => sheet.mergeCells(`${col}15:${col}16`));

  // 2. Table Headers
  const headers = ["S/N", "REG. NO.", "NAME", null, "CA TOTAL (/30)", "EXAM TOTAL (/70)", "INTERNAL (/100)", "EXTERNAL (/100)", "AGREED (/100)", null];
  const headerRow = sheet.getRow(15);
  headerRow.height = 47;

  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, name: fontName, size: 9 };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = thinBorder;
    if (i >= 3) cell.alignment.textRotation = 90;
  });

  const maxScoreRow = sheet.getRow(16);
  for (let c = 1; c <= 10; c++) {
    const cell = maxScoreRow.getCell(c);
    cell.font = { bold: true, name: fontName, size: 8 };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { ...thinBorder, bottom: { style: "double" } };

    // Set the specific max values for the score columns
    if (c === 4 ) {
      cell.value = "ATTEMPT";
      cell.alignment.textRotation = 90; // Added rotation here
    }
    if (c === 5) cell.value = 30;
    if (c === 6) cell.value = 70;
    if (c === 7) cell.value = 100;
    if (c === 8) cell.value = 100;
    if (c === 9) cell.value = 100;

    if ( c === 10) {
      cell.value = "GRADE";
      cell.alignment.textRotation = 90;
    }

    // Fill color only for the score-related headers
    if (c >= 5 && c <= 9) cell.fill = greyFill;
  }

  // 4. Data Rows (Starting Row 17)
  const startRow = 17;
  const extraRows = 15; // Added padding
  const endRow = startRow + eligibleStudents.length + extraRows;

  for (let r = startRow; r <= endRow; r++) {
    const studentIdx = r - startRow;
    const student = eligibleStudents[studentIdx];
    const row = sheet.getRow(r);
    row.height = 13;

    let attemptLabel = "1st";
    let isSupp = false;
    let isSpecial = false;

    if (student) {
      const prevMark = previousMarks.find(
        (m) => m.student.toString() === student._id.toString(),
      );
      if (prevMark) {
        if (prevMark.isSpecial || prevMark.attempt === "special") {
          attemptLabel = "Special"; isSpecial = true;
        } else if (prevMark.agreedMark < settings.passMark || prevMark.attempt === "supplementary") {
          attemptLabel = "Supp"; isSupp = true;
        } else if (prevMark.attempt === "re-take") {
          attemptLabel = "Retake";
        }
      }

      row.getCell(1).value = studentIdx + 1;
      row.getCell(2).value = student.regNo;
      row.getCell(3).value = student.name.toUpperCase();
      row.getCell(4).value = attemptLabel;
      row.getCell(3).font = { name: fontName, size: 8 };
    }

    const isRowEmpty = `ISBLANK(B${r})`;

    // Logic: Internal = CA + Exam
    const caVal = isSupp ? "0" : `E${r}`;
    row.getCell(7).value = {
      formula: `IF(${isRowEmpty}, "", ROUND(${caVal} + F${r}, 0))`,
    };

    // Logic: Agreed (ENG 13.f)
    // If Supp: Cap at passmark. Else: Use External if exists, otherwise Internal.
    const internal = `G${r}`;
    const external = `H${r}`;
    const effectiveMark = `IF(${external}<>"", ${external}, ${internal})`;
    const finalAgreed = `IF(D${r}="Supp", MIN(${settings.passMark}, ${effectiveMark}), ${effectiveMark})`;

    row.getCell(9).value = { formula: `IF(${isRowEmpty}, "", ${finalAgreed})` };

    // Grading Logic Nesting
    const sortedScale = [...(settings.gradingScale || [])].sort((a, b) => a.min - b.min);
    let gradeIfs = `"E"`;
    sortedScale.forEach((scale) => { gradeIfs = `IF(I${r}>=${scale.min}, "${scale.grade}", ${gradeIfs})`; });
    row.getCell(10).value = { formula: `IF(${isRowEmpty}, "", ${gradeIfs})` };

    // Formatting & Validations
    for (let c = 1; c <= 10; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder;
      cell.font = { name: fontName, size: 8 };
      cell.alignment = { vertical: "middle" };

      if (c === 5) { // CA
        cell.fill = pinkColor;
        cell.protection = { locked: isSupp || isSpecial };
        if (isSupp) cell.value = 0;
        cell.dataValidation = { type: "decimal", operator: "between", formulae: [0, 30], allowBlank: true, showErrorMessage: true, errorTitle: "Invalid CA", error: "Value must be 0-30" };
      } else if (c === 6) { // Exam
        cell.fill = purpleFill;
        cell.protection = { locked: false };
        cell.dataValidation = { type: "decimal", operator: "between", formulae: [0, 70], allowBlank: true, showErrorMessage: true, errorTitle: "Invalid Exam", error: "Value must be 0-70" };
      } else if (c === 8) { // External
        cell.protection = { locked: false };
        cell.dataValidation = { type: "decimal", operator: "between", formulae: [0, 100], allowBlank: true };
      } else if (c >= 7) { // Internal, Agreed, Grade
        cell.fill = greyFill;
        cell.protection = { locked: true };
      }
    }
  }

  // 5. Thick Outer Borders
  for (let r = 15; r <= endRow; r++) {
    for (let c = 1; c <= 10; c++) {
      const cell = sheet.getCell(r, c);
      cell.border = {
        ...cell.border,
        left: c === 1 ? { style: "thick" } : cell.border?.left,
        right: c === 9 ? { style: "thick" } : cell.border?.right,
        top: r === 15 ? { style: "thick" } : cell.border?.top,
        bottom: r === endRow ? { style: "thick" } : cell.border?.bottom,
      };
    }
  }

  // Column Widths
  sheet.getColumn(1).width = 4;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 35;
  sheet.getColumn(4).width = 6;
  sheet.getColumn(10).width = 6;
  [5, 6, 7, 8, 9].forEach((c) => (sheet.getColumn(c).width = 12));
  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 16 }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
  return await workbook.xlsx.writeBuffer();
};;