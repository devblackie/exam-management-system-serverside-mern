// src/utils/directTemplate.ts
// import * as ExcelJS from "exceljs";
// import Program from "../models/Program";
// import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import Student from "../models/Student";
// import config from "../config/config";
// import mongoose from "mongoose";
// import InstitutionSettings from "../models/InstitutionSettings";
// import MarkDirect from "../models/MarkDirect";
// import ProgramUnit from "../models/ProgramUnit";
// import Mark from "../models/Mark";

// export const generateDirectScoresheetTemplate = async (
//   programId: mongoose.Types.ObjectId,
//   unitId: mongoose.Types.ObjectId,
//   yearOfStudy: number,
//   semester: number,
//   academicYearId: mongoose.Types.ObjectId,
//   logoBuffer: any,
// ): Promise<any> => {
//   const programUnit = await ProgramUnit.findOne({
//     program: programId,
//     unit: unitId,
//   }).lean();
//   // const [program, unit, academicYear, eligibleStudents] = await Promise.all([
//   //   Program.findById(programId).lean(),
//   //   Unit.findById(unitId).lean(),
//   //   AcademicYear.findById(academicYearId).lean(),
//   //   Student.find({
//   //     program: programId,
//   //     currentYearOfStudy: yearOfStudy,
//   //     status: "active",
//   //   })
//   //     .sort({ regNo: 1 })
//   //     .lean(),
//   // ]);

//   const [program, unit, academicYear] = await Promise.all([
//     Program.findById(programId).lean(),
//     Unit.findById(unitId).lean(),
//     AcademicYear.findById(academicYearId).lean(),
//   ]);

//   // Get the programUnit record so we can look up prior marks for this unit
//   const programUnitDoc = await ProgramUnit.findOne({
//     program: programId,
//     unit: unitId,
//   }).lean();

//   // All active students in this year (base pool)
//   const allActiveStudents = await Student.find({
//     program: programId,
//     currentYearOfStudy: yearOfStudy,
//     status: { $in: ["active", "repeat"] },
//   })
//     .sort({ regNo: 1 })
//     .lean();

//   // During SUPPLEMENTARY or CLOSED: only include students who actually
//   // need this specific unit — i.e. they failed it, have a special for it,
//   // or have a supp/retake attempt recorded.
//   // During ORDINARY: include everyone (marks haven't been entered yet).
//   let eligibleStudents = allActiveStudents;

//   if (
//     programUnitDoc &&
//     (academicYear?.session === "SUPPLEMENTARY" ||
//       academicYear?.session === "CLOSED")
//   ) {
//     const studentIds = allActiveStudents.map((s) => s._id);

//     // Find all prior marks for this unit across these students
//     const [priorDetailed, priorDirect] = await Promise.all([
//       Mark.find({
//         student: { $in: studentIds },
//         programUnit: programUnitDoc._id,
//       }).lean(),
//       MarkDirect.find({
//         student: { $in: studentIds },
//         programUnit: programUnitDoc._id,
//       }).lean(),
//     ]);

//     // Build a set of student IDs who need this unit in this session
//     const needsUnitIds = new Set<string>();

//     [...priorDetailed, ...priorDirect].forEach((m: any) => {
//       const sid = m.student?.toString();
//       if (!sid) return;

//       const mark = m.agreedMark ?? 0;
//       const passed = mark >= (settings?.passMark ?? 40);

//       // Include if: failed, special pending, supplementary attempt, or retake
//       if (
//         !passed ||
//         m.isSpecial ||
//         m.attempt === "supplementary" ||
//         m.attempt === "re-take"
//       ) {
//         needsUnitIds.add(sid);
//       }
//     });

//     // Also include students who have NO prior mark at all for this unit
//     // (they may simply not have sat the exam yet — still relevant during supp)
//     const studentsWithMark = new Set<string>(
//       [...priorDetailed, ...priorDirect].map((m: any) => m.student?.toString()),
//     );

//     allActiveStudents.forEach((s) => {
//       if (!studentsWithMark.has(s._id.toString())) {
//         // No mark exists — only include if they are expected to sit this unit
//         // (i.e. the unit is in their curriculum year)
//         needsUnitIds.add(s._id.toString());
//       }
//     });

//     eligibleStudents = allActiveStudents.filter((s) =>
//       needsUnitIds.has(s._id.toString()),
//     );
//   }

//   const settings = await InstitutionSettings.findOne({
//     institution: program?.institution,
//   });
//   if (!settings) throw new Error("Institution settings missing.");

//   const previousMarks = programUnit
//     ? await MarkDirect.find({
//         student: { $in: eligibleStudents.map((s) => s._id) },
//         programUnit: programUnit._id,
//       }).lean()
//     : [];

//   const workbook = new ExcelJS.Workbook();
//   const sheet = workbook.addWorksheet(`${unit?.code || "SCORESHEET"}`);

//   // Styles
//   const fontName = "Book Antiqua";
//   const fontSize = 9;
//   const greyFill = {
//     type: "pattern" as const,
//     pattern: "solid" as const,
//     fgColor: { argb: "FFE0E0E0" },
//   };
//   const pinkColor = {
//     type: "pattern" as const,
//     pattern: "solid" as const,
//     fgColor: { argb: "FFFFA6C9" },
//   };
//   const purpleFill = {
//     type: "pattern" as const,
//     pattern: "solid" as const,
//     fgColor: { argb: "FFC5A3FF" },
//   };
//   const thinBorder: Partial<ExcelJS.Borders> = {
//     top: { style: "thin" },
//     left: { style: "thin" },
//     bottom: { style: "thin" },
//     right: { style: "thin" },
//   };

//   // 1. Logo & Branding

//   if (logoBuffer?.length > 0) {
//     const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
//     sheet.addImage(logoId, {
//       tl: { col: 3, row: 0 },
//       ext: { width: 100, height: 80 },
//     });
//   }

//   const centerBold = {
//     alignment: { horizontal: "center" as const, vertical: "middle" as const },
//     font: { bold: true, name: fontName, underline: true },
//   };

//   sheet.mergeCells("C6:G6");
//   sheet.getCell("C6").value = config.instName.toUpperCase();
//   sheet.getCell("C6").style = {
//     ...centerBold,
//     font: { ...centerBold.font, size: 12 },
//   };

//   sheet.mergeCells("C7:G7");
//   sheet.getCell("C7").value = `DEGREE: ${program?.name.toUpperCase()}`;
//   sheet.getCell("C7").style = centerBold;

//   const yrTxt =
//     ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] ||
//     `${yearOfStudy}TH`;
//   sheet.mergeCells("C8:G8");
//   sheet.getCell("C8").value =
//     `${yrTxt} YEAR | SEMESTER ${semester} | ${academicYear?.year || ""} ACADEMIC YEAR`;
//   sheet.getCell("C8").style = centerBold;

//   sheet.mergeCells("C10:G10");
//   sheet.getCell("C10").value = `SCORESHEET FOR: ${unit?.code.toUpperCase()}`;
//   sheet.getCell("C10").style = {
//     ...centerBold,
//     font: { ...centerBold.font, size: 10 },
//   };

//   // Unit Info
//   sheet.getCell("B12").value = "UNIT TITLE:";
//   sheet.getCell("C12").value = unit?.name.toUpperCase();
//   sheet.getCell("E12").value = "UNIT CODE:";
//   sheet.getCell("F12").value = unit?.code;
//   sheet.getRow(12).font = { name: fontName, bold: true, size: 9 };

//   // sheet.mergeCells("A15:A16");
//   // sheet.mergeCells("B15:B16");
//   // sheet.mergeCells("C15:C16");
//   // sheet.mergeCells("D15:D16");

//   // Header merging for static columns
//   ["A", "B", "C", "D", "J"].forEach((col) => sheet.mergeCells(`${col}15:${col}16`));

//   // 2. Table Headers
//   const headers = ["S/N", "REG. NO.", "NAME", null, "CA TOTAL (/30)", "EXAM TOTAL (/70)", "INTERNAL (/100)", "EXTERNAL (/100)", "AGREED (/100)", null];
//   const headerRow = sheet.getRow(15);
//   headerRow.height = 47;

//   headers.forEach((h, i) => {
//     const cell = headerRow.getCell(i + 1);
//     cell.value = h;
//     cell.font = { bold: true, name: fontName, size: 9 };
//     cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
//     cell.border = thinBorder;
//     if (i >= 3) cell.alignment.textRotation = 90;
//   });

//   const maxScoreRow = sheet.getRow(16);
//   for (let c = 1; c <= 10; c++) {
//     const cell = maxScoreRow.getCell(c);
//     cell.font = { bold: true, name: fontName, size: 8 };
//     cell.alignment = { vertical: "middle", horizontal: "center" };
//     cell.border = { ...thinBorder, bottom: { style: "double" } };

//     // Set the specific max values for the score columns
//     if (c === 4 ) {
//       cell.value = "ATTEMPT";
//       cell.alignment.textRotation = 90; // Added rotation here
//     }
//     if (c === 5) cell.value = 30;
//     if (c === 6) cell.value = 70;
//     if (c === 7) cell.value = 100;
//     if (c === 8) cell.value = 100;
//     if (c === 9) cell.value = 100;

//     if ( c === 10) {
//       cell.value = "GRADE";
//       cell.alignment.textRotation = 90;
//     }

//     // Fill color only for the score-related headers
//     if (c >= 5 && c <= 9) cell.fill = greyFill;
//   }

//   // 4. Data Rows (Starting Row 17)
//   const startRow = 17;
//   const extraRows = 15; // Added padding
//   const endRow = startRow + eligibleStudents.length + extraRows;

//   for (let r = startRow; r <= endRow; r++) {
//     const studentIdx = r - startRow;
//     const student = eligibleStudents[studentIdx];
//     const row = sheet.getRow(r);
//     row.height = 13;

//     let attemptLabel = "1st";
//     let isSupp = false;
//     let isSpecial = false;

//     if (student) {
//       const prevMark = previousMarks.find(
//         (m) => m.student.toString() === student._id.toString(),
//       );
//       if (prevMark) {
//         if (prevMark.isSpecial || prevMark.attempt === "special") {
//           attemptLabel = "Special";
//           isSpecial = true;
//           // ENG.18(c): special exam is marked out of 100% INCLUDING existing CA.
//           // Pre-fill the CA cell so the examiner sees the locked-in CA value.
//           if (prevMark.caTotal30 != null && prevMark.caTotal30 > 0) {
//             row.getCell(5).value = prevMark.caTotal30;
//           }
//         } else if (prevMark.agreedMark < settings.passMark || prevMark.attempt === "supplementary") {
//           attemptLabel = "Supp"; isSupp = true;
//         } else if (prevMark.attempt === "re-take") {
//           attemptLabel = "Retake";
//         }
//       }

//       row.getCell(1).value = studentIdx + 1;
//       row.getCell(2).value = student.regNo;
//       row.getCell(3).value = student.name.toUpperCase();
//       row.getCell(4).value = attemptLabel;
//       row.getCell(3).font = { name: fontName, size: 8 };
//     }

//     const isRowEmpty = `ISBLANK(B${r})`;

//     // Logic: Internal = CA + Exam
//     const caVal = isSupp ? "0" : `E${r}`;
//     row.getCell(7).value = {
//       formula: `IF(${isRowEmpty}, "", ROUND(${caVal} + F${r}, 0))`,
//     };

//     // Logic: Agreed (ENG 13.f)
//     // If Supp: Cap at passmark. Else: Use External if exists, otherwise Internal.
//     const internal = `G${r}`;
//     const external = `H${r}`;
//     const effectiveMark = `IF(${external}<>"", ${external}, ${internal})`;
//     const finalAgreed = `IF(D${r}="Supp", MIN(${settings.passMark}, ${effectiveMark}), ${effectiveMark})`;

//     row.getCell(9).value = { formula: `IF(${isRowEmpty}, "", ${finalAgreed})` };

//     // Grading Logic Nesting
//     const sortedScale = [...(settings.gradingScale || [])].sort((a, b) => a.min - b.min);
//     let gradeIfs = `"E"`;
//     sortedScale.forEach((scale) => { gradeIfs = `IF(I${r}>=${scale.min}, "${scale.grade}", ${gradeIfs})`; });
//     row.getCell(10).value = { formula: `IF(${isRowEmpty}, "", ${gradeIfs})` };

//     // Formatting & Validations
//     for (let c = 1; c <= 10; c++) {
//       const cell = row.getCell(c);
//       cell.border = thinBorder;
//       cell.font = { name: fontName, size: 8 };
//       cell.alignment = { vertical: "middle" };

//       // if (c === 5) { // CA
//       //   cell.fill = pinkColor;
//       //   cell.protection = { locked: isSupp || isSpecial };
//       //   if (isSupp) cell.value = 0;
//       //   cell.dataValidation = { type: "decimal", operator: "between", formulae: [0, 30], allowBlank: true, showErrorMessage: true, errorTitle: "Invalid CA", error: "Value must be 0-30" };
//       // }
//       if (c === 5) {
//         // Special: lock cell, show existing CA in grey
//         // Supp: lock cell, force zero
//         // Normal: pink editable input
//         if (isSpecial) {
//           cell.fill = greyFill;
//           cell.protection = { locked: true };
//           // value was already set above from prevMark.caTotal30
//           // no data validation on locked cells
//         } else if (isSupp) {
//           cell.fill = greyFill;
//           cell.protection = { locked: true };
//           cell.value = 0;
//         } else {
//           cell.fill = pinkColor;
//           cell.protection = { locked: false };
//           cell.dataValidation = {
//             type: "decimal",
//             operator: "between",
//             formulae: [0, 30],
//             allowBlank: true,
//             showErrorMessage: true,
//             errorTitle: "Invalid CA",
//             error: "CA Total must be 0–30",
//           };
//         }
//       } else if (c === 6) {
//         // Exam
//         cell.fill = purpleFill;
//         cell.protection = { locked: false };
//         cell.dataValidation = {
//           type: "decimal",
//           operator: "between",
//           formulae: [0, 70],
//           allowBlank: true,
//           showErrorMessage: true,
//           errorTitle: "Invalid Exam",
//           error: "Value must be 0-70",
//         };
//       } else if (c === 8) {
//         // External
//         cell.protection = { locked: false };
//         cell.dataValidation = {
//           type: "decimal",
//           operator: "between",
//           formulae: [0, 100],
//           allowBlank: true,
//         };
//       } else if (c >= 7) {
//         // Internal, Agreed, Grade
//         cell.fill = greyFill;
//         cell.protection = { locked: true };
//       }
//     }
//   }

//   // 5. Thick Outer Borders
//   for (let r = 15; r <= endRow; r++) {
//     for (let c = 1; c <= 10; c++) {
//       const cell = sheet.getCell(r, c);
//       cell.border = {
//         ...cell.border,
//         left: c === 1 ? { style: "thick" } : cell.border?.left,
//         right: c === 9 ? { style: "thick" } : cell.border?.right,
//         top: r === 15 ? { style: "thick" } : cell.border?.top,
//         bottom: r === endRow ? { style: "thick" } : cell.border?.bottom,
//       };
//     }
//   }

//   // Column Widths
//   sheet.getColumn(1).width = 4;
//   sheet.getColumn(2).width = 22;
//   sheet.getColumn(3).width = 35;
//   sheet.getColumn(4).width = 6;
//   sheet.getColumn(10).width = 6;
//   [5, 6, 7, 8, 9].forEach((c) => (sheet.getColumn(c).width = 12));
//   sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 16 }];
//   sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });
//   return await workbook.xlsx.writeBuffer();
// };;

// serverside/src/utils/directTemplate.ts
import * as ExcelJS from "exceljs";
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import Student, { IStudent } from "../models/Student";
import ProgramUnit from "../models/ProgramUnit";
import InstitutionSettings from "../models/InstitutionSettings";
import MarkDirect from "../models/MarkDirect";
import Mark from "../models/Mark";
import config from "../config/config";
import mongoose from "mongoose";
import { getAttemptLabel, ADMIN_STATUS_LABELS } from "./academicRules";

// ─── Typed interfaces (no implicit any) ──────────────────────────────────────

interface StudentDoc {
  // _id: mongoose.Types.ObjectId;
  _id: any;
  regNo: string;
  name: string;
  status: string;
  // program: mongoose.Types.ObjectId;
  program: any;

  

  currentYearOfStudy: number;
  academicHistory?: Array<{
    yearOfStudy: number;
    isRepeatYear?: boolean;
    academicYear: string;
  }>;
  academicLeavePeriod?: {
    type?: string;
    startDate?: Date;
    endDate?: Date;
  };
}

interface MarkRecord {
  // _id: mongoose.Types.ObjectId;
  // student: mongoose.Types.ObjectId;
  // programUnit: mongoose.Types.ObjectId;
  _id: any;
  student: any;
  programUnit: any;
  agreedMark?: number;
  caTotal30?: number;
  examTotal70?: number;
  attempt?: string;
  isSpecial?: boolean;
  remarks?: string;
}

interface SettingsDoc {
  passMark: number;
  gradingScale?: Array<{ min: number; grade: string }>;
}

// ─────────────────────────────────────────────────────────────────────────────

export const generateDirectScoresheetTemplate = async (
  programId:     mongoose.Types.ObjectId,
  unitId:        mongoose.Types.ObjectId,
  yearOfStudy:   number,
  semester:      number,
  academicYearId: mongoose.Types.ObjectId,
  logoBuffer:    any,
): Promise<any> => {

  // ── Fetch metadata ──────────────────────────────────────────────────────
  const [program, unit, academicYear] = await Promise.all([
    Program.findById(programId).lean() as Promise<{ name: string; institution: mongoose.Types.ObjectId } | null>,
    Unit.findById(unitId).lean() as Promise<{ code: string; name: string } | null>,
    AcademicYear.findById(academicYearId).lean() as Promise<{
      year: string;
      session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
    } | null>,
  ]);

  const settings = await InstitutionSettings.findOne({
    institution: program?.institution,
  }).lean() as SettingsDoc | null;

  if (!settings) throw new Error("Institution settings not found.");
  if (!academicYear) throw new Error("Academic Year not found.");
  if (!unit) throw new Error("Unit not found.");

  const programUnitDoc = await ProgramUnit.findOne({
    program: programId,
    unit:    unitId,
  }).lean() as { _id: mongoose.Types.ObjectId } | null;

  // ── Student pool ────────────────────────────────────────────────────────
  // During ORDINARY: all active students in this year
  // During SUPPLEMENTARY/CLOSED: only students who need this specific unit
  // const baseStudents = await Student.find({
  //   program:            programId,
  //   currentYearOfStudy: yearOfStudy,
  //   status:             { $in: ["active", "repeat", "on_leave"] },
  // })
  //   .sort({ regNo: 1 })
  //   .lean() as StudentDoc[];
  const [marksThisYear, directMarksThisYear] = await Promise.all([
    Mark.distinct("student", { academicYear: academicYearId }),
    MarkDirect.distinct("student", { academicYear: academicYearId }),
  ]);
  
  const markedIds = new Set<string>([
    ...marksThisYear.map((id: any) => id.toString()),
    ...directMarksThisYear.map((id: any) => id.toString()),
  ]);
  
  // Also include students admitted this academic year (no marks yet during ORDINARY)
  const admittedThisYear = await Student.distinct("_id", {
    program:              programId,
    admissionAcademicYear: academicYearId,
    currentYearOfStudy:   yearOfStudy,
  });
  admittedThisYear.forEach((id: any) => markedIds.add(id.toString()));
  
  // Base pool: active/repeat students who either have marks or were admitted this year
  const baseStudents = await Student.find({
    program:            programId,
    currentYearOfStudy: yearOfStudy,
    // status:             { $in: ["active", "repeat", "on_leave"] },
    status:             { $in: ["active", "repeat"] },
    _id:                { $in: Array.from(markedIds) },
  })
    .sort({ regNo: 1 })
    .lean() as StudentDoc[];

  let eligibleStudents: StudentDoc[] = baseStudents;

  if (
    programUnitDoc &&
    (academicYear.session === "SUPPLEMENTARY" || academicYear.session === "CLOSED")
  ) {
    const studentIds = baseStudents.map((s) => s._id);

    // Find prior marks for this specific unit
    const [priorDetailed, priorDirect] = await Promise.all([
      Mark.find({
        student:     { $in: studentIds },
        programUnit: programUnitDoc._id,
      }).lean() as Promise<MarkRecord[]>,
      MarkDirect.find({
        student:     { $in: studentIds },
        programUnit: programUnitDoc._id,
      }).lean() as Promise<MarkRecord[]>,
    ]);

    const needsUnit = new Set<string>();
    const hasMarkSet = new Set<string>();

    [...priorDetailed, ...priorDirect].forEach((m) => {
      const sid    = m.student?.toString();
      if (!sid) return;
      hasMarkSet.add(sid);
      const mark   = m.agreedMark ?? 0;
      const passed = mark >= settings.passMark;
      if (!passed || m.isSpecial || m.attempt === "supplementary" || m.attempt === "re-take") {
        needsUnit.add(sid);
      }
    });

    // Students with no prior mark also appear (they may have a new attempt)
    baseStudents.forEach((s) => {
      if (!hasMarkSet.has(s._id.toString())) {
        needsUnit.add(s._id.toString());
      }
    });

    eligibleStudents = baseStudents.filter((s) =>
      needsUnit.has(s._id.toString())
    );
  }

  // ── Fetch existing marks for pre-population ─────────────────────────────
  const previousMarks: MarkRecord[] = programUnitDoc
    ? (await MarkDirect.find({
        student:     { $in: eligibleStudents.map((s) => s._id) },
        programUnit: programUnitDoc._id,
      }).lean() as MarkRecord[])
    : [];

  // ── Excel setup ─────────────────────────────────────────────────────────
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet(
    `${unit.code || "SCORESHEET"}`.trim().substring(0, 31)
  );

  const fontName = "Book Antiqua";

  const greyFill: ExcelJS.Fill = {
    type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E0E0" },
  };
  const pinkFill: ExcelJS.Fill = {
    type: "pattern", pattern: "solid", fgColor: { argb: "FFFFA6C9" },
  };
  const purpleFill: ExcelJS.Fill = {
    type: "pattern", pattern: "solid", fgColor: { argb: "FFC5A3FF" },
  };
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin" }, left: { style: "thin" },
    bottom: { style: "thin" }, right: { style: "thin" },
  };
  const doubleBorder: Partial<ExcelJS.Borders> = {
    ...thinBorder, bottom: { style: "double" },
  };

  // ── Exam type label (changes header for SUPPLEMENTARY session) ───────────
  const examTypeLabel =
    academicYear.session === "SUPPLEMENTARY"
      ? "SUPPLEMENTARY AND SPECIAL EXAMINATION"
      : "EXAMINATION";

  // ── Logo ────────────────────────────────────────────────────────────────
  if (logoBuffer?.length > 0) {
    const logoId = workbook.addImage({ buffer: logoBuffer, extension: "png" });
    sheet.addImage(logoId, { tl: { col: 3, row: 0 }, ext: { width: 100, height: 80 } });
  }

  const centerBold = {
    alignment: { horizontal: "center" as const, vertical: "middle" as const },
    font: { bold: true, name: fontName, underline: true },
  };

  // ── Header rows ─────────────────────────────────────────────────────────
  sheet.mergeCells("C6:G6");
  sheet.getCell("C6").value = config.instName.toUpperCase();
  sheet.getCell("C6").style = { ...centerBold, font: { ...centerBold.font, size: 12 } };

  sheet.mergeCells("C7:G7");
  sheet.getCell("C7").value = `DEGREE: ${(program?.name || "").toUpperCase()}`;
  sheet.getCell("C7").style = centerBold;

  const yrTxt = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH"][yearOfStudy - 1] ?? `${yearOfStudy}TH`;
  const semTxt = semester === 1 ? "FIRST" : "SECOND";

  sheet.mergeCells("C8:G8");
  sheet.getCell("C8").value =
    `${yrTxt} YEAR | ${semTxt} SEMESTER | ${academicYear.year} ACADEMIC YEAR`;
  sheet.getCell("C8").style = centerBold;

  sheet.mergeCells("C10:G10");
  sheet.getCell("C10").value = `SCORESHEET FOR: ${unit.code.toUpperCase()} — ${examTypeLabel}`;
  sheet.getCell("C10").style = { ...centerBold, font: { ...centerBold.font, size: 10 } };

  // ── Unit info ────────────────────────────────────────────────────────────
  sheet.getCell("B12").value = "UNIT TITLE:";
  sheet.getCell("C12").value = unit.name.toUpperCase();
  sheet.getCell("E12").value = "UNIT CODE:";
  sheet.getCell("F12").value = unit.code;
  sheet.getRow(12).font = { name: fontName, bold: true, size: 9 };

  // ── Static column merges ─────────────────────────────────────────────────
  ["A", "B", "C", "D", "J"].forEach((col) =>
    sheet.mergeCells(`${col}15:${col}16`)
  );

  // ── Headers (row 15 & 16) ────────────────────────────────────────────────
  const headers = [
    "S/N", "REG. NO.", "NAME", null,
    "CA TOTAL (/30)", "EXAM TOTAL (/70)",
    "INTERNAL (/100)", "EXTERNAL (/100)", "AGREED (/100)", null,
  ];

  const headerRow = sheet.getRow(15);
  headerRow.height = 47;

  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font      = { bold: true, name: fontName, size: 9 };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border    = thinBorder;
    if (i >= 3) cell.alignment.textRotation = 90;
  });

  const maxScoreRow = sheet.getRow(16);
  const maxScores = [null, null, null, "ATTEMPT", 30, 70, 100, 100, 100, "GRADE"];
  maxScores.forEach((v, i) => {
    const cell  = maxScoreRow.getCell(i + 1);
    cell.value  = v;
    cell.font   = { bold: true, name: fontName, size: 8 };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = doubleBorder;
    if (i >= 4 && i <= 8) cell.fill = greyFill;
    if (i === 3 || i === 9) cell.alignment.textRotation = 90;
  });

  // ── Data rows (start at row 17) ──────────────────────────────────────────
  const startRow  = 17;
  const extraRows = 15;
  const endRow    = startRow + eligibleStudents.length + extraRows;

  const sortedScale = [...(settings.gradingScale || [])].sort((a, b) => a.min - b.min);
  let gradeIfs = `"E"`;
  sortedScale.forEach((s) => {
    gradeIfs = `IF(I{R}>=${s.min}, "${s.grade}", ${gradeIfs})`;
  });

  for (let r = startRow; r <= endRow; r++) {
    const idx     = r - startRow;
    const student = eligibleStudents[idx] as StudentDoc | undefined;
    const row     = sheet.getRow(r);
    row.height    = 13;

    let attemptLabel = "1st";
    let isSupp       = false;
    let isSpecial    = false;

    if (student) {
      const sid      = student._id.toString();
      const prevMark = previousMarks.find(
        (m) => m.student?.toString() === sid
      );

      // Derive attempt label using academicRules
      const repeatCount = (student.academicHistory || []).filter(
        (h) => h.isRepeatYear && h.yearOfStudy === yearOfStudy
      ).length;

      const st = (student.status || "").toLowerCase();

      if (prevMark?.isSpecial || prevMark?.attempt === "special") {
        attemptLabel = "Special";
        isSpecial    = true;
      } else if (
        prevMark?.attempt === "supplementary" ||
        (prevMark && (prevMark.agreedMark ?? 0) < settings.passMark && st === "active")
      ) {
        attemptLabel = "Supp";
        isSupp       = true;
      } else if (prevMark?.attempt === "re-take") {
        attemptLabel = getAttemptLabel({
          markAttempt: "re-take",
          studentStatus: student.status,
          regNo: student.regNo,
        });
      } else {
        attemptLabel = getAttemptLabel({
          markAttempt: "1st",
          studentStatus: student.status,
          regNo: student.regNo,
          repeatYearCount: repeatCount,
        });
      }

      row.getCell(1).value = idx + 1;
      row.getCell(2).value = student.regNo;
      row.getCell(3).value = student.name.toUpperCase();
      row.getCell(4).value = attemptLabel;
      row.getCell(3).font  = { name: fontName, size: 8 };

      // Pre-fill CA for special students (ENG.18c: marked out of 100% including CA)
      if (isSpecial && prevMark) {
        const ca = prevMark.caTotal30 ?? 0;
        if (ca > 0) row.getCell(5).value = ca;
      }
    }

    const isRowEmpty = `ISBLANK(B${r})`;

    // CA cell (E) = pre-filled and locked for supp/special
    const caVal = isSupp ? "0" : `E${r}`;

    // Internal = CA + Exam (supp: CA forced to 0 per ENG.13f)
    row.getCell(7).value = {
      formula: `IF(${isRowEmpty}, "", ROUND(${caVal} + F${r}, 0))`,
    };

    // Agreed mark (capped at passMark for supp per ENG.13f)
    const internal  = `G${r}`;
    const external  = `H${r}`;
    const effective = `IF(${external}<>"", ${external}, ${internal})`;
    const finalAgreed =
      `IF(D${r}="Supp", MIN(${settings.passMark}, ${effective}), ${effective})`;
    row.getCell(9).value = { formula: `IF(${isRowEmpty}, "", ${finalAgreed})` };

    // Grade
    let gradeFormula = `"E"`;
    sortedScale.forEach((s) => {
      gradeFormula = `IF(I${r}>=${s.min}, "${s.grade}", ${gradeFormula})`;
    });
    row.getCell(10).value = { formula: `IF(${isRowEmpty}, "", ${gradeFormula})` };

    // ── Cell styling ──────────────────────────────────────────────────────
    for (let c = 1; c <= 10; c++) {
      const cell = row.getCell(c);
      cell.border    = thinBorder;
      cell.font      = { name: fontName, size: 8 };
      cell.alignment = { vertical: "middle" };

      if (c === 5) {
        if (isSpecial) {
          // Show existing CA value locked (grey)
          cell.fill       = greyFill;
          cell.protection = { locked: true };
        } else if (isSupp) {
          // Supps: CA = 0, locked (ENG.13f)
          cell.fill       = greyFill;
          cell.protection = { locked: true };
          cell.value      = 0;
        } else {
          cell.fill       = pinkFill;
          cell.protection = { locked: false };
          cell.dataValidation = {
            type: "decimal", operator: "between",
            formulae: [0, 30], allowBlank: true,
            showErrorMessage: true,
            errorTitle: "Invalid CA", error: "CA Total must be 0–30",
          };
        }
      } else if (c === 6) {
        cell.fill       = purpleFill;
        cell.protection = { locked: false };
        cell.dataValidation = {
          type: "decimal", operator: "between",
          formulae: [0, 70], allowBlank: true,
          showErrorMessage: true,
          errorTitle: "Invalid Exam", error: "Exam Total must be 0–70",
        };
      } else if (c === 8) {
        // External examiner — editable
        cell.protection = { locked: false };
        cell.dataValidation = {
          type: "decimal", operator: "between",
          formulae: [0, 100], allowBlank: true,
        };
      } else if (c >= 7) {
        // Internal, Agreed, Grade — computed, locked
        cell.fill       = greyFill;
        cell.protection = { locked: true };
      }
    }
  }

  // ── Thick outer borders ──────────────────────────────────────────────────
  for (let r = 15; r <= endRow; r++) {
    for (let c = 1; c <= 10; c++) {
      const cell = sheet.getCell(r, c);
      cell.border = {
        ...cell.border,
        left:   c === 1   ? { style: "thick" } : cell.border?.left,
        right:  c === 9   ? { style: "thick" } : cell.border?.right,
        top:    r === 15  ? { style: "thick" } : cell.border?.top,
        bottom: r === endRow ? { style: "thick" } : cell.border?.bottom,
      };
    }
  }

  // ── Column widths ────────────────────────────────────────────────────────
  sheet.getColumn(1).width = 4;
  sheet.getColumn(2).width = 22;
  sheet.getColumn(3).width = 35;
  sheet.getColumn(4).width = 6;
  sheet.getColumn(10).width = 6;
  [5, 6, 7, 8, 9].forEach((c) => (sheet.getColumn(c).width = 12));

  sheet.views = [{ state: "frozen", xSplit: 4, ySplit: 16 }];
  sheet.protect("1234", { selectLockedCells: true, selectUnlockedCells: true });

  const buf = await workbook.xlsx.writeBuffer();
  return Buffer.from(buf as ArrayBuffer);
};