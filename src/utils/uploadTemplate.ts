// // src/utils/uploadTemplate.ts
// export const MARKS_UPLOAD_HEADERS = [
//   "RegNo",
//   "UnitCode",
//   "CAT1",
//   "CAT2",
//   "CAT3",
//   "Assignment",
//   "Practical",
//   "Exam",
//   "IsSupplementary", // "YES" or "NO"
//   "AcademicYear",    // e.g. "2024/2025"
// ] as const;

// export const generateSampleCSV = () => {
//   const header = MARKS_UPLOAD_HEADERS.join(",");
//   const sampleRows = [
//     'SC/ICT/001/2023,ICS2107,18,22,20,8,9,65,NO,2024/2025',
//     'SC/ICT/002/2023,ICS2107,15,19,,7,8,58,YES,2024/2025',
//   ].join("\n");
//   return `${header}\n${sampleRows}`;
// };

// src/utils/uploadTemplate.ts
// import Program from "../models/Program"; // Import necessary models for DB interaction
// import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import mongoose from "mongoose";

// export const MARKS_UPLOAD_HEADERS = [
//     // Student Info
//     "S/N", 
//     "REG. NO. ", // <-- Note the space
//     "NAME", 
//     "ATTEMPT", 
    
//     // Continuous Assessment
//     "CAT 1 Out of 20", 
//     "CAT 2 Out of 20", 
//     "CAT3 Out of 20",
//     "TOTAL (CA Total Out of 20)", // Renamed for clarity, original: TOTAL
//     "Assgnt 1 Out of 10",
//     "Assgnt 2 Out of 10",
//     "Assgnt 3 Out of 10",
//     "TOTAL (Assgnt Total Out of 10)", // Renamed for clarity, original: TOTAL
//     "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30", // This is the combined CA
    
//     // Exam Breakdown
//     "Q1 out of 10",
//     "Q2 out of 20", 
//     "Q3 out of 20", 
//     "Q4 out Of 20", // <-- Note the capitalization
//     "Q5 out of 20",
//     "TOTAL EXAM OUT OF 70", 
    
//     // Final Marks & Grade
//     "INTERNAL EXAMINER MARKS /100", 
//     "EXTERNAL EXAMINER MARKS /100", 
//     "AGREED MARKS /100", 
//     "GRADE",
// ] as const;

// // Dynamic header generator — pulls from DB
// export const generateScoresheetHeader = async (
//   programId: string,
//   unitId: string,
//   yearOfStudy: number,
//   semester: number,
//   academicYearId: string
// ) => {
//   // Fetch from DB
//   const [program, unit, academicYear] = await Promise.all([
//     Program.findById(programId).lean(),
//     Unit.findById(unitId).lean(),
//     AcademicYear.findById(academicYearId).lean(),
//   ]);

//   if (!program || !unit || !academicYear) throw new Error("Invalid selection");

//   const semesterText = semester === 1 ? "FIRST" : "SECOND";
//   const yearText = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH", "SIXTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;

//   return `KENYATTA UNIVERSITY
// DEGREE: ${program.name.toUpperCase()}
// ${yearText} YEAR ${semesterText} SEMESTER ${academicYear.year} ACADEMIC YEAR

// SCORESHEET
// UNIT CODE:\t${unit.code}\t\t\t\t\tUNIT TITLE: ${unit.name.toUpperCase()}

// ${MARKS_UPLOAD_HEADERS.join("\t")}
// 20\t20\t20\t20\t10\t10\t10\t10\t\t10\t20\t20\t20\t20\t70`;
// };

// // Generates a sample CSV row for template illustration
// export const generateSampleCSV = () => {
//     const header = MARKS_UPLOAD_HEADERS.join(",");
//     const sampleRow = [
//         '1', // S/N
//         'T056-01-0049/2020', // REG. NO. 
//         'Gregory Onyango OWINY', // NAME
//         '1st', // ATTEMPT
//         '13.0', // CAT 1
//         '16.00', // CAT 2
//         '', // CAT 3
//         '14.50', // CAT TOTAL (Out of 20)
//         '6.0', // Assgnt 1
//         '', // Assgnt 2
//         '', // Assgnt 3
//         '6.00', // Assgnt TOTAL (Out of 10)
//         '20.5', // CA Grand Total /30
//         '4.0', // Q1
//         '5.0', // Q2
//         '8.0', // Q3
//         '14.0', // Q4
//         '', // Q5
//         '31.0', // Total Exam /70
//         '52', // Internal /100
//         '', // External /100
//         '52', // Agreed /100
//         'C', // GRADE
//     ].join(",");
    
//     return `${header}\n${sampleRow}`;
// };



// // src/utils/uploadTemplate.ts 
// import Program from "../models/Program";
// import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import mongoose from "mongoose";

// export const MARKS_UPLOAD_HEADERS = [
//   "S/N",
//   "REG. NO.",
//   "NAME",
//   "ATTEMPT",
//   "CAT 1 Out of 20",
//   "CAT 2 Out of 20",
//   "CAT3 Out of 20",
//   "TOTAL (CA Total Out of 20)",
//   "Assgnt 1 Out of 10",
//   "Assgnt 2 Out of 10",
//   "Assgnt 3 Out of 10",
//   "TOTAL (Assgnt Total Out of 10)",
//   "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30",
//   "Q1 out of 10",
//   "Q2 out of 20",
//   "Q3 out of 20",
//   "Q4 out of 20",
//   "Q5 out of 20",
//   "TOTAL EXAM OUT OF 70",
//   "INTERNAL EXAMINER MARKS /100",
//   "EXTERNAL EXAMINER MARKS /100",
//   "AGREED MARKS /100",
//   "GRADE",
// ] as const;

// export const MARKS_MAX_SCORES_ROW = [
//   "", "", "", "",
//   20, 20, 20, 20,
//   10, 10, 10, 10,
//   "", 
//   10, 20, 20, 20, 20, 70,
//   "", "", "", ""
// ].join(",");

// export const generateFullScoresheetTemplate = async (
//   programId: mongoose.Types.ObjectId,
//   unitId: mongoose.Types.ObjectId,
//   yearOfStudy: number,
//   semester: number,
//   academicYearId: mongoose.Types.ObjectId
// ) => {
//   const [program, unit, academicYear] = await Promise.all([
//     Program.findById(programId).lean(),
//     Unit.findById(unitId).lean(),
//     AcademicYear.findById(academicYearId).lean(),
//   ]);

//   if (!program || !unit || !academicYear) {
//     throw new Error("Invalid selection: Program, Unit, or Academic Year not found.");
//   }

//   const semesterText = semester === 1 ? "FIRST" : "SECOND";
//   const yearText = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH", "SIXTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;

//   // === BOLD, CENTERED, PROFESSIONAL KU HEADER (USING PADDING FOR VISUAL BOLDNESS) ===
//   const boldHeader = [
//     // University Name - BIG & VISUALLY BOLD
//     "                                KENYATTA UNIVERSITY",
//     "",
//     // Degree - VISUALLY BOLD
//     `                                 DEGREE: ${program.name.toUpperCase()}`,
//     "",
//     // Year & Semester - VISUALLY BOLD
//     `        ${yearText} YEAR ${semesterText} SEMESTER ${academicYear.year} ACADEMIC YEAR`,
//     "",
//     "",
//     // Scoresheet Title - VISUALLY BOLD
//     "                                  SCORESHEET",
//     "",
//     // Unit Info - VISUALLY BOLD & SPACED
//     `UNIT CODE: ${unit.code}`.padStart(20) + `                UNIT TITLE: ${unit.name.toUpperCase()}`,
//     "",
//     "",
//     // Column headers
//     MARKS_UPLOAD_HEADERS.join(","),
//     // Max scores row
//     MARKS_MAX_SCORES_ROW,
//   ].join("\n");

//   // Sample data
//   const sampleRow = [
//     '1', 'T056-01-0049/2020', 'Gregory Onyango OWINY', '1st',
//     '13.0', '16.00', '', '14.50', '6.0', '', '', '6.00', '20.5',
//     '4.0', '5.0', '8.0', '14.0', '', '31.0', '52', '', '52', 'C'
//   ].join(",");

//   return `${boldHeader}\n${sampleRow}\n\n`; // Extra lines for spacing
// };

// src/utils/generateScoresheetXLSX.ts - TRUE EXCEL FILE
import Program from "../models/Program";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import mongoose from "mongoose";
import * as ExcelJS from 'exceljs';
import { Buffer } from 'buffer';

// Defines the headers and their corresponding column index letters (E=5, M=13, N=14, T=20)
// This is crucial for writing the formulas.
export const MARKS_UPLOAD_HEADERS = [
  "S/N", "REG. NO.", "NAME", "ATTEMPT", // A-D
  "CAT 1 Out of 20", "CAT 2 Out of 20", "CAT3 Out of 20", // E-G
  "TOTAL (CA Total Out of 20)", // H - FORMULA
  "Assgnt 1 Out of 10", "Assgnt 2 Out of 10", "Assgnt 3 Out of 10", // I-K
  "TOTAL (Assgnt Total Out of 10)", // L - FORMULA
  "CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30", // M - FORMULA (CA Final)
  "Q1 out of 10", "Q2 out of 20", "Q3 out of 20", "Q4 out of 20", "Q5 out of 20", // N-R
  "TOTAL EXAM OUT OF 70", // S - FORMULA
  "INTERNAL EXAMINER MARKS /100", "EXTERNAL EXAMINER MARKS /100", // T-U
  "AGREED MARKS /100", "GRADE", // V-W
];

export const MARKS_MAX_SCORES_ROW = [
  null, null, null, null,
  20, 20, 20, 20,
  10, 10, 10, 10,
  30, 
  10, 20, 20, 20, 20, 70,
  100, 100, 100, null
];

export const generateFullScoresheetTemplate = async (
  programId: mongoose.Types.ObjectId,
  unitId: mongoose.Types.ObjectId,
  yearOfStudy: number,
  semester: number,
  academicYearId: mongoose.Types.ObjectId,
  logoBuffer: Buffer
): Promise<Buffer> => {
  const [program, unit, academicYear] = await Promise.all([
    Program.findById(programId).lean(),
    Unit.findById(unitId).lean(),
    AcademicYear.findById(academicYearId).lean(),
  ]);

  if (!program || !unit || !academicYear) {
    throw new Error("Invalid selection: Program, Unit, or Academic Year not found.");
  }

  const semesterText = semester === 1 ? "FIRST" : "SECOND";
  const yearText = ["FIRST", "SECOND", "THIRD", "FOURTH", "FIFTH", "SIXTH"][yearOfStudy - 1] || `${yearOfStudy}TH`;

  // 1. Create a new Workbook and Worksheet
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Scoresheet Template");
  
  // Define common styles
  const boldStyle = { bold: true };
  const centerStyle = { alignment: { horizontal: 'center' as const } };
  const caColor = { fgColor: { argb: 'FFCCEEFF' } }; // Light Blue
  const examColor = { fgColor: { argb: 'FFFFFFCC' } }; // Light Yellow

 

  // === KU HEADER INFORMATION (Starts at Row 1) ===
  let currentRow = 4;
  sheet.mergeCells(`A${currentRow}:W${currentRow}`);
  sheet.getCell(`A${currentRow}`).value = ' UNIVERSITY  OF TECHNOLOGY';
  sheet.getCell(`A${currentRow}`).font = { bold: true, size: 14 };
  sheet.getCell(`A${currentRow}`).alignment = centerStyle.alignment;
  currentRow++;

  currentRow++;
  sheet.mergeCells(`A${currentRow}:W${currentRow}`);
  sheet.getCell(`A${currentRow}`).value = `DEGREE: ${program.name.toUpperCase()}`;
  sheet.getCell(`A${currentRow}`).font = boldStyle;
  sheet.getCell(`A${currentRow}`).alignment = centerStyle.alignment;
  currentRow++;

  currentRow++;
  sheet.mergeCells(`H${currentRow}:I${currentRow}`);
  sheet.getCell(`H${currentRow}`).value = `${yearText} YEAR`;
  sheet.getCell(`H${currentRow}`).font = boldStyle;
  sheet.getCell(`H${currentRow}`).alignment = centerStyle.alignment;

  sheet.mergeCells(`K${currentRow}:L${currentRow}`);
  sheet.getCell(`K${currentRow}`).value = ` ${semesterText} SEMESTER`;
  sheet.getCell(`K${currentRow}`).font = boldStyle;
  sheet.getCell(`K${currentRow}`).alignment = centerStyle.alignment;

  sheet.mergeCells(`N${currentRow}:P${currentRow}`);
  sheet.getCell(`N${currentRow}`).value = `${academicYear.year} ACADEMIC YEAR`;
  sheet.getCell(`N${currentRow}`).font = boldStyle;
  sheet.getCell(`N${currentRow}`).alignment = centerStyle.alignment;
  currentRow++;

  currentRow++;
  sheet.mergeCells(`A${currentRow}:W${currentRow}`);
  sheet.getCell(`A${currentRow}`).value = 'SCORESHEET';
  sheet.getCell(`A${currentRow}`).font = { bold: true, size: 12 };
  sheet.getCell(`A${currentRow}`).alignment = centerStyle.alignment;
  currentRow++;

  currentRow++;
  sheet.mergeCells(`G${currentRow}:I${currentRow}`);
  sheet.getCell(`G${currentRow}`).value = `UNIT CODE: ${unit.code}`;
  sheet.getCell(`G${currentRow}`).font = boldStyle;

  sheet.mergeCells(`M${currentRow}:W${currentRow}`);
  sheet.getCell(`M${currentRow}`).value = `UNIT TITLE: ${unit.name.toUpperCase()}`;
  sheet.getCell(`M${currentRow}`).font = boldStyle;
  currentRow++;
  
  currentRow++;
  currentRow++;

  // 2. Column Headers (Row 10)
  sheet.addRow(MARKS_UPLOAD_HEADERS);
  const headerRow = sheet.getRow(currentRow);
  // headerRow.font = boldStyle;

  headerRow.eachCell({ includeEmpty: false }, (cell) => {
    cell.font = boldStyle;
});
  
  // Apply Background Colors (CA: E-M, EXAM: N-S)
  for (let i = 5; i <= 13; i++) { // E to M (CA Columns)
    headerRow.getCell(i).fill = { type: 'pattern', pattern: 'solid', ...caColor };
  }
  for (let i = 14; i <= 19; i++) { // N to S (Exam Columns)
    headerRow.getCell(i).fill = { type: 'pattern', pattern: 'solid', ...examColor };
  }
  currentRow++;

  // 3. Max Scores Row (Row 11)
  sheet.addRow(MARKS_MAX_SCORES_ROW);
  const maxRow = sheet.getRow(currentRow);
  maxRow.font = boldStyle;
  maxRow.alignment = centerStyle.alignment;
  // Apply the same colors to the Max Scores row
  for (let i = 5; i <= 13; i++) {
    maxRow.getCell(i).fill = { type: 'pattern', pattern: 'solid', ...caColor };
  }
  for (let i = 14; i <= 19; i++) {
    maxRow.getCell(i).fill = { type: 'pattern', pattern: 'solid', ...examColor };
  }
  currentRow++; // This is the first data row (Row 12)

  // 4. Sample Data & Formulas (Data starts from Row 12)
  const dataStartRow = currentRow; 
  
  const sampleData = [
    '1', 'T056-01-0049/2020', 'Gregory Onyango OWINY', '1st',
    13.0, 16.00, 0, // CAT 1-3
    null, // H - Formula
    6.0, 0, 0, // Assgnt 1-3
    null, // L - Formula
    null, // M - Formula (Grand Total)
    4.0, 5.0, 8.0, 14.0, 0, // Q1-Q5
    null, // S - Formula
    52, 0, 52, 'C' // Examiner Marks, Agreed, Grade
  ];
  sheet.addRow(sampleData);
  
  // Set up the formulas for the first data row (Row 12)
  // H: TOTAL (CA Total Out of 20) = (CAT1 + CAT2 + CAT3) * (20/60)
  sheet.getCell(`H${dataStartRow}`).value = { formula: `((E${dataStartRow}+F${dataStartRow}+G${dataStartRow}) * 20 / 60)`, result: 9.67 };
  // L: TOTAL (Assgnt Total Out of 10) = (Assgnt1 + Assgnt2 + Assgnt3) * (10/30)
  sheet.getCell(`L${dataStartRow}`).value = { formula: `((I${dataStartRow}+J${dataStartRow}+K${dataStartRow}) * 10 / 30)`, result: 2.0 };
  // M: CATs + LABS + ASSIGNMENTS GRAND TOTAL out of 30 = H + L + (Labs if applicable, for now assume H+L)
  sheet.getCell(`M${dataStartRow}`).value = { formula: `H${dataStartRow}+L${dataStartRow}`, result: 11.67 };
  // S: TOTAL EXAM OUT OF 70 = Q1 + Q2 + Q3 + Q4 + Q5
  sheet.getCell(`S${dataStartRow}`).value = { formula: `SUM(N${dataStartRow}:R${dataStartRow})`, result: 31.0 };
  // T: INTERNAL EXAMINER MARKS /100 = M + S
sheet.getCell(`T${dataStartRow}`).value = { formula: `M${dataStartRow}+S${dataStartRow}`, result: 42.67 };

// Apply a fill to the formula cells in the data row
  sheet.getCell(`H${dataStartRow}`).fill = { type: 'pattern', pattern: 'solid', ...caColor };
  sheet.getCell(`L${dataStartRow}`).fill = { type: 'pattern', pattern: 'solid', ...caColor };
  sheet.getCell(`M${dataStartRow}`).fill = { type: 'pattern', pattern: 'solid', ...caColor };
  sheet.getCell(`S${dataStartRow}`).fill = { type: 'pattern', pattern: 'solid', ...examColor };
  
  // 5. Finalize and Write to Buffer
  const buffer = await workbook.xlsx.writeBuffer();
  
  return buffer as unknown as Buffer;
};