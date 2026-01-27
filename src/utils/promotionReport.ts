// serverside/src/utils/promotionReport.ts
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  ImageRun,
} from "docx";
import config from "../config/config";

export interface PromotionData {
  programName: string;
  academicYear: string;
  yearOfStudy: number;
  eligible: any[];
  blocked: any[];
  logoBuffer: Buffer;
}

const formatStudentName = (fullName: string): string => {
  if (!fullName) return "";
  const parts = fullName.trim().split(/\s+/);

  // If only one name exists, just uppercase it
  if (parts.length <= 1) return fullName.toUpperCase();

  // Remove the last name, uppercase it, then join back
  const lastName = parts.pop()?.toUpperCase();
  return `${parts.join(" ")} ${lastName}`;
};

export const generatePromotionWordDoc = async (
  data: PromotionData,
): Promise<Buffer> => {
  const {
    programName,
    academicYear,
    yearOfStudy,
    eligible,
    blocked,
    logoBuffer,
  } = data;

  // Formatting constant for reuse
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO (Using the correct ImageRun structure)
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 100, height: 100 },
                      // Use the correct internal type mapping for docx
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. UNIVERSITY HEADERS
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200 },
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 28,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "OFFICE OF THE REGISTRAR (ACADEMIC AFFAIRS)",
                bold: true,
                size: 20,
              }),
            ],
          }),

          // 3. REPORT TITLE
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 400 },
            children: [
              new TextRun({
                text: `PROMOTION SUMMARY REPORT: ${academicYear}`,
                bold: true,
                size: 24,
                underline: { type: BorderStyle.SINGLE },
              }),
            ],
          }),

          // 4. METADATA (Headers)
          new Paragraph({
            children: [
              new TextRun({ text: `PROGRAM: `, bold: true }),
              new TextRun({ text: programName.toUpperCase() }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `CURRENT YEAR OF STUDY: `, bold: true }),
              new TextRun({ text: `YEAR ${yearOfStudy}` }),
            ],
          }),

          // 5. EXECUTIVE SUMMARY
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({ text: "1.0 EXECUTIVE SUMMARY", bold: true }),
            ],
          }),

          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    margins: cellMargin,
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: "Description", bold: true }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    margins: cellMargin,
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: "Student Count", bold: true }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: cellMargin,
                    children: [new Paragraph("Eligible for Promotion")],
                  }),
                  new TableCell({
                    margins: cellMargin,
                    children: [new Paragraph(eligible.length.toString())],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: cellMargin,
                    children: [new Paragraph("Blocked / Action Required")],
                  }),
                  new TableCell({
                    margins: cellMargin,
                    children: [new Paragraph(blocked.length.toString())],
                  }),
                ],
              }),
            ],
          }),

          // 6. DETAILED LISTS
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({ text: "2.0 DETAILED PROMOTION LIST", bold: true }),
            ],
          }),
          createStudentTable([...eligible, ...blocked]),

          // 7. SIGNATORIES
          new Paragraph({
            spacing: { before: 1200 },
            children: [
              new TextRun({
                text: "PREPARED BY: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({ text: "FACULTY COORDINATOR" }),

          new Paragraph({
            spacing: { before: 800 },
            children: [
              new TextRun({
                text: "APPROVED BY: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({ text: "CHAIRMAN, ACADEMIC BOARD" }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

function createStudentTable(students: any[]) {
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const headerRow = new TableRow({
    children: [
      new TableCell({
        shading: { fill: "E0E0E0" },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Reg No", bold: true })],
          }),
        ],
      }),
      new TableCell({
        shading: { fill: "E0E0E0" },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Full Name", bold: true })],
          }),
        ],
      }),
      new TableCell({
        shading: { fill: "E0E0E0" },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "Decision/Reasons", bold: true })],
          }),
        ],
      }),
    ],
  });

  const dataRows = students.map(
    (s) =>
      new TableRow({
        children: [
          new TableCell({
            margins: cellMargin,
            children: [new Paragraph(s.regNo)],
          }),
          // new TableCell({ margins: cellMargin, children: [new Paragraph(s.name)] }),
          new TableCell({
            margins: cellMargin,
            children: [new Paragraph(formatStudentName(s.name))],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph(
                s.status === "IN GOOD STANDING"
                  ? "PROMOTED"
                  : s.reasons?.join(", ") || s.status,
              ),
            ],
          }),
        ],
      }),
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}

export const generateEligibleSummaryDoc = async (
  data: any,
): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, eligible, logoBuffer } = data;
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. Headers
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 24,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "ORDINARY EXAMINATION RESULTS",
                bold: true,
                size: 20,
              }),
            ],
          }),

          // 3. SPECIFIC ACADEMIC INFO
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200 },
            children: [
              new TextRun({
                text: `${academicYear} ACADEMIC YEAR`,
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `YEAR ${yearOfStudy} (PROMOTED STUDENTS)`,
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: programName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: "PASS",
                bold: true,
                size: 22,
                underline: {},
              }),
            ],
          }),

          // 4. INTRODUCTORY TEXT
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: `The following ${eligible.length} candidates satisfied the College Board of Examiners during the ${academicYear} Academic year Examinations. The College Board of examiners therefore recommends that the students proceed to the next Year of study.`,
                size: 20,
              }),
            ],
          }),

          // 5. THE PASS LIST TABLE
          createPassTable(eligible, cellMargin),

          // 6. SIGNATORIES (Aligned with the image layout)
          new Paragraph({
            spacing: { before: 600 },
            children: [
              new TextRun({
                text: `APPROVED BY THE BOARD OF EXAMINERS, ${config.schoolName.toUpperCase()}`,
                bold: true,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "SIGNED: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `\tDEAN, ${config.schoolName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),

          // new Paragraph({
          //   spacing: { before: 600 },
          //   children: [
          //     new TextRun({ text: "APPROVED BY THE COLLEGE OF HEALTH SCIENCES (COHES) BOARD OF EXAMINERS", bold: true, size: 18 }),
          //   ],
          // }),
          // new Paragraph({
          //   spacing: { before: 400 },
          //   children: [
          //     new TextRun({ text: "SIGNED: __________________________\t\tDATE: _______________", bold: true }),
          //   ],
          // }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

function createPassTable(students: any[], cellMargin: any) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        width: { size: 5, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "NO.", bold: true, size: 18 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 30, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "REG. NO.", bold: true, size: 18 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 65, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "NAME", bold: true, size: 18 })],
          }),
        ],
      }),
    ],
  });

  const dataRows = students.map(
    (s, index) =>
      new TableRow({
        children: [
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: (index + 1).toString(), size: 18 }),
                ],
              }),
            ],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [new TextRun({ text: s.regNo, size: 18 })],
              }),
            ],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: formatStudentName(s.name), size: 18 }),
                ],
              }),
            ],
          }),
        ],
      }),
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [headerRow, ...dataRows],
  });
}

export const generateIneligibleSummaryDoc = async (
  data: PromotionData,
): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  // Separate students into two lists
  const specialExamPending = blocked.filter((s) =>
    s.reasons?.some((r: string) => r.toLowerCase().includes("special")),
  );
  const failureList = blocked.filter(
    (s) => !s.reasons?.some((r: string) => r.toLowerCase().includes("special")),
  );

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. HEADERS
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 24,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "ORDINARY EXAMINATION RESULTS",
                bold: true,
                size: 20,
              }),
            ],
          }),

          // 3. SPECIFIC ACADEMIC INFO
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200 },
            children: [
              new TextRun({
                text: `${academicYear} ACADEMIC YEAR`,
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `YEAR ${yearOfStudy} (INELIGIBLE STUDENTS)`,
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: programName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 300 },
            children: [
              new TextRun({
                text: "FAIL / INCOMPLETE / INELIGIBLE",
                bold: true,
                size: 22,
                underline: {},
              }),
            ],
          }),

          // 4. INTRODUCTORY TEXT (Modified for Ineligible context)
          // new Paragraph({
          //   alignment: AlignmentType.LEFT,
          //   spacing: { after: 200 },
          //   children: [
          //     new TextRun({
          //       text: `The following ${blocked.length} candidates DID NOT satisfy the College Board of Examiners during the ${academicYear} Academic year Examinations. The College Board therefore recommends that the students NOT proceed to the next Year of study until the specified requirements are met.`,
          //       size: 20
          //     }),
          //   ],
          // }),

          // 1. SPECIAL EXAMINATIONS SECTION
          new Paragraph({
            spacing: { before: 400, after: 200 },
            children: [
              new TextRun({
                text: "LIST A: SPECIAL EXAMINATIONS PENDING",
                bold: true,
                size: 22,
                underline: {},
              }),
            ],
          }),
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "The following candidates have been granted permission to sit for Special Examinations.",
                size: 18,
                italics: true,
              }),
            ],
          }),
          createIneligibleTable(specialExamPending, cellMargin),

          // 2. FAIL / SUPPLEMENTARY SECTION
          new Paragraph({
            spacing: { before: 600, after: 200 },
            children: [
              new TextRun({
                text: "LIST B: FAIL / SUPPLEMENTARY EXAMINATIONS",
                bold: true,
                size: 22,
                underline: {},
              }),
            ],
          }),
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "The following candidates are required to sit for Supplementary Examinations or Retakes.",
                size: 18,
                italics: true,
              }),
            ],
          }),
          createIneligibleTable(failureList, cellMargin),

          // 5. THE INELIGIBLE LIST TABLE (Maintaining Reasons)
          createIneligibleTable(blocked, cellMargin),

          // 6. SIGNATORIES
          new Paragraph({
            spacing: { before: 600 },
            children: [
              new TextRun({
                text: `APPROVED BY THE BOARD OF EXAMINERS, ${config.schoolName.toUpperCase()}`,
                bold: true,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "SIGNED: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `\tDEAN, ${config.schoolName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

function createIneligibleTable(students: any[], cellMargin: any) {
  const headerRow = new TableRow({
    children: [
      new TableCell({
        width: { size: 5, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "NO.", bold: true, size: 18 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "REG. NO.", bold: true, size: 18 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 40, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [new TextRun({ text: "NAME", bold: true, size: 18 })],
          }),
        ],
      }),
      new TableCell({
        width: { size: 30, type: WidthType.PERCENTAGE },
        margins: cellMargin,
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "REASON(S)", bold: true, size: 18 }),
            ],
          }),
        ],
      }),
    ],
  });

  const dataRows = students.map(
    (s, index) =>
      new TableRow({
        children: [
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: (index + 1).toString(), size: 18 }),
                ],
              }),
            ],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [new TextRun({ text: s.regNo, size: 18 })],
              }),
            ],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: formatStudentName(s.name), size: 18 }),
                ],
              }),
            ],
          }),
          new TableCell({
            margins: cellMargin,
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: s.reasons?.join(", ") || s.status,
                    size: 16,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      // top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [headerRow, ...dataRows],
  });
}

// New function for individual ineligibility notice
interface IneligibilityNoticeData {
  programName: string;
  academicYear: string;
  yearOfStudy: number;
  logoBuffer: Buffer;
}

export const generateIneligibilityNotice = async (
  student: any,
  data: IneligibilityNoticeData,
): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, logoBuffer } = data;

  const capitalizedStudentName = formatStudentName(student.name).toUpperCase();
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. HEADERS (Consistent with Pass List)
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 24,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),

          // 3. INTERNAL MEMO STYLE ADDRESSING
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({ text: "TO: ", bold: true }),
              new TextRun({ text: capitalizedStudentName }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "REG NO: ", bold: true }),
              new TextRun({ text: student.regNo }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "PROGRAM: ", bold: true }),
              new TextRun({ text: programName.toUpperCase() }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "ACADEMIC YEAR: ", bold: true }),
              new TextRun({ text: academicYear }),
            ],
          }),

          // 4. SUBJECT LINE
          new Paragraph({
            spacing: { before: 400, after: 400 },
            children: [
              new TextRun({
                text: `RE: INELIGIBILITY FOR PROMOTION TO YEAR ${yearOfStudy + 1}`,
                bold: true,
                size: 22,
                underline: {},
              }),
            ],
          }),

          // 5. BODY TEXT
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "Following the School Board of Examiners meeting, it was noted that you did not satisfy the examiners in the following units during the current academic year:",
                size: 20,
              }),
            ],
          }),

          // 6. UNIT LIST (Dynamically rendered from student.reasons)
          ...student.reasons.map(
            (unitString: string) =>
              new Paragraph({
                bullet: { level: 0 },
                spacing: { before: 150 },
                children: [
                  new TextRun({
                    text: unitString,
                    bold: true,
                    size: 19,
                  }),
                ],
              }),
          ),

          new Paragraph({
            spacing: { before: 300, after: 200 },
            children: [
              new TextRun({
                text: "Consequently, you are not eligible for promotion. You are advised to prepare for Supplementary Examinations or register for Retakes as per the University Examination Regulations.",
                size: 20,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: `Please contact the Office of the Dean, ${config.schoolName}, for the schedule of supplementary examinations.`,
                size: 20,
              }),
            ],
          }),

          // 7. SIGNATORY
          new Paragraph({
            spacing: { before: 600, after: 400 },
            children: [
              new TextRun({
                text: `APPROVED BY THE BOARD OF EXAMINERS, ${config.schoolName.toUpperCase()}`,
                bold: true,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "SIGNED: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `\tDEAN, ${config.schoolName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),

          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "Cc: Registrar (Academic Affairs)\n    ",
                size: 16,
              }),
            ],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

// New function for Special Exam Notice
export const generateSpecialExamNotice = async (
  student: any,
  data: any,
): Promise<Buffer> => {
  const { programName, academicYear, logoBuffer } = data;
  const capitalizedStudentName = formatStudentName(student.name).toUpperCase();

  const specialUnits = student.reasons
    .filter((r: string) => r.toUpperCase().includes("SPECIAL"))
    .map((r: string) =>
      r.replace("- SPECIAL", "").replace("SPECIAL", "").trim(),
    );

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. HEADERS
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 24,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),

          // 3. ADDRESSING
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({ text: "TO: ", bold: true }),
              new TextRun({ text: capitalizedStudentName }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "REG NO: ", bold: true }),
              new TextRun({ text: student.regNo }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "PROGRAM: ", bold: true }),
              new TextRun({ text: programName.toUpperCase() }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "ACADEMIC YEAR: ", bold: true }),
              new TextRun({ text: academicYear }),
            ],
          }),

          // 4. SUBJECT
          new Paragraph({
            spacing: { before: 400, after: 400 },
            children: [
              new TextRun({
                text: "RE: APPROVAL TO SIT FOR SPECIAL EXAMINATIONS",
                bold: true,
                size: 22,
                underline: { type: BorderStyle.SINGLE },
              }),
            ],
          }),

          // 5. BODY
          new Paragraph({
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: `This is to inform you that the College Board of Examiners has approved your request to sit for Special Examinations in the following unit(s) during the ${academicYear} academic cycle:`,
                size: 20,
              }),
            ],
          }),

          // DYNAMIC UNIT LIST
          ...specialUnits.map(
            (unitName: string) =>
              new Paragraph({
                bullet: { level: 0 },
                spacing: { before: 100 },
                children: [
                  new TextRun({ text: unitName, bold: true, size: 20 }),
                ],
              }),
          ),

          new Paragraph({
            spacing: { before: 300, after: 200 },
            children: [
              new TextRun({
                text: "Please note that a Special Examination is treated as a first attempt. Failure to sit for these exams will result in the units being graded as 'Incomplete'.",
                size: 20,
              }),
            ],
          }),

          new Paragraph({
            spacing: { after: 400 },
            children: [
              new TextRun({
                text: "Check the departmental notice board for the scheduled dates and venues.",
                size: 20,
              }),
            ],
          }),

          // 7. SIGNATORY
          new Paragraph({
            spacing: { before: 800 },
            children: [
              new TextRun({
                text: "SIGNED: __________________________\t\tDATE: _______________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `\tDEAN, ${config.schoolName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 400 },
            children: [new TextRun({ text: "Cc: Exam Coordinator", size: 16 })],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

export const generateStudentTranscript = async (
  student: any,
  results: any[],
  data: any,
): Promise<Buffer> => {
    const { programName, academicYear, logoBuffer, status } = data;

  // Use the value passed from the controller, falling back to student record if needed
  const displayYear = data.yearToPromote || data.yearOfStudy || student.currentYearOfStudy || "N/A";

  const cellMargin = { top: 80, bottom: 80, left: 100, right: 100 };

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO (Same 80x80 size as Summary)
          ...(logoBuffer.length > 0
            ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ImageRun({
                      data: logoBuffer,
                      transformation: { width: 80, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. INSTITUTIONAL HEADERS
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: config.instName.toUpperCase(),
                bold: true,
                size: 24,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: true,
                size: 20,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "PROVISIONAL STATEMENT OF RESULTS",
                bold: true,
                size: 20,
                underline: {},
              }),
            ],
          }),

          // 3. STUDENT PROFILE BOX
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: `NAME: ${formatStudentName(student.name).toUpperCase()}`,
                            bold: true,
                            size: 20,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                          new TextRun({
                            text: `REG NO: ${student.regNo}`,
                            bold: true,
                            size: 20,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: `PROGRAM: ${programName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 100, after: 300 },
            children: [
              new TextRun({
                // text: `ACADEMIC YEAR: ${academicYear || "N/A"}`,
                text: `ACADEMIC YEAR: ${academicYear || "N/A"}`,
                size: 18,
              }),
              // new TextRun({ text: `\t\tYEAR OF STUDY: ${yearOfStudy || "N/A"}`, size: 18 }),
              new TextRun({
                // text: `\t\tYEAR OF STUDY: ${data.yearOfStudy !== undefined ? yearOfStudy : "N/A"}`,
                text: `\t\tYEAR OF STUDY: ${displayYear}`,
                size: 18,
              }),
            ],
          }),

          // 4. RESULTS TABLE (Units & Grades)
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              // Header Row
              new TableRow({
                tableHeader: true,
                children: [
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    margins: cellMargin,
                    shading: { fill: "F2F2F2" },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: "CODE", bold: true, size: 18 }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 70, type: WidthType.PERCENTAGE },
                    margins: cellMargin,
                    shading: { fill: "F2F2F2" },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "COURSE UNIT TITLE",
                            bold: true,
                            size: 18,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    margins: cellMargin,
                    shading: { fill: "F2F2F2" },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [
                          new TextRun({ text: "GRADE", bold: true, size: 18 }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              // Data Rows
              // Data Rows (Mapping the results)
              ...results.map((r: any) => {
                // Defensive extraction to prevent .toUpperCase() on undefined
                const unitCode = String(r?.code ?? "N/A").toUpperCase();
                const unitName = String(
                  r?.name ?? "COURSE TITLE MISSING",
                ).toUpperCase();
                const unitGrade = String(r?.grade ?? "-");

                return new TableRow({
                  children: [
                    new TableCell({
                      margins: cellMargin,
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: unitCode, size: 18 })],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: cellMargin,
                      children: [
                        new Paragraph({
                          children: [new TextRun({ text: unitName, size: 18 })],
                        }),
                      ],
                    }),
                    new TableCell({
                      margins: cellMargin,
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: unitGrade,
                              bold: true,
                              size: 18,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                });
              }),
            ],
          }),

          // 5. STATUS SUMMARY
          new Paragraph({
            spacing: { before: 300 },
            children: [
              new TextRun({ text: "RECOMMENDATION: ", bold: true, size: 18 }),
              new TextRun({
                text: (
                  status ||
                  student.status ||
                  data.status ||
                  "PENDING"
                ).toUpperCase(),
                size: 18,
                bold: true,
              }),
            ],
          }),

          // 6. GRADING KEY (Consistent with University Standards)
          new Paragraph({
            spacing: { before: 400 },
            children: [
              new TextRun({ text: "GRADING KEY: ", bold: true, size: 14 }),
              new TextRun({
                text: "A (70-100%) EXCELLENT | B (60-69%) GOOD | C (50-59%) SATISFACTORY | D (40-49%) PASS | E (0-39%) FAIL",
                size: 14,
              }),
            ],
          }),

          // 7. FOOTER & SIGNATORIES
          new Paragraph({
            spacing: { before: 800 },
            children: [
              new TextRun({
                text: "SIGNED: __________________________________________",
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `DEAN, ${config.schoolName.toUpperCase()}`,
                bold: true,
                size: 18,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: `DATE OF ISSUE: ${new Date().toLocaleDateString()}`,
                size: 14,
                italics: true,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400 },
            children: [
              new TextRun({
                text: "--- This result slip is issued without any erasures or alterations ---",
                italics: true,
                size: 14,
              }),
            ],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

// export const generateStudentTranscript = async (student: any, results: any[], data: any): Promise<Buffer> => {
//   const { programName, academicYear, yearOfStudy, logoBuffer } = data;

//   const doc = new Document({
//     sections: [{
//       properties: {},
//       children: [
//         // 1. Logo & Headers
//         ...(logoBuffer.length > 0 ? [
//           new Paragraph({
//             alignment: AlignmentType.CENTER,
//             children: [new ImageRun({ data: logoBuffer, transformation: { width: 150, height: 150 }, type: "png" })],
//           }),
//         ] : []),
//         new Paragraph({
//           spacing: { after: 200 },
//           alignment: AlignmentType.CENTER,
//           children: [new TextRun({ text: config.instName.toUpperCase(), bold: true, size: 32 })],
//         }),
//          new Paragraph({
//           spacing: { before: 100 ,after:200},
//           alignment: AlignmentType.CENTER,
//           children: [
//             new TextRun({ text: "UNDERGRADUATE ACADEMIC TRANSCRIPT", bold: true, size: 20 }),
//           ],
//         }),
//          // 2. Student Info Header
//      new Paragraph({
//   spacing: { before: 400 },
//   children: [
//     // Student Name Section
//     new TextRun({ text: "NAME: ", bold: true }),
//     new TextRun({ text: formatStudentName(student.name).toUpperCase(), bold: false }),

//     // Spacing (Tabs)
//     new TextRun({ text: "\t\t\t\t" }),

//     // Registration Number Section
//     new TextRun({ text: "REG NO: ", bold: true }),
//     new TextRun({ text: student.regNo, bold: false }),
//   ],
// }),
//  new Paragraph({
//           children: [
//             new TextRun({ text: "SCHOOL: ", bold: true}),
//             new TextRun({ text: config.schoolName.toUpperCase(), bold: false }),
//           ],
//         }),
//          new Paragraph({
//           children: [
//             new TextRun({ text: "DEGREE: ", bold: true}),
//             new TextRun({ text: programName.toUpperCase, bold: false }),
//           ],
//         }),
// new Paragraph({
//   children: [
//     // Student Name Section
//     new TextRun({ text: "ACADEMIC YEAR: ", bold: true }),
//     new TextRun({ text: academicYear.toUpperCase(), bold: false }),

//     // Spacing (Tabs)
//     new TextRun({ text: "\t\t\t\t" }),

//     // Registration Number Section
//     new TextRun({ text: "YEAR OF STUDY: ", bold: true }),
//     new TextRun({ text: yearOfStudy.toString.toUpperCase(), bold: false }),
//   ],
// }),
//          new Paragraph({
//   spacing: { before: 400,after:200 },

//           alignment: AlignmentType.CENTER,
//           children: [
//             new TextRun({ text: student.status.toUpperCase(), bold: true, size: 20, underline: true }),
//           ],
//         }),

//         // 3. Results Table
//         new Table({
//           width: { size: 100, type: WidthType.PERCENTAGE },
//           rows: [
//             // Header Row
//             new TableRow({
//               children: [
//                 new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "CODE", bold: true, size: 20 })] })] }),
//                 new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "UNIT NAME", bold: true, size: 20 })] })] }),
//                 new TableCell({
//                   shading: { fill: "f2f2f2" },
//                   alignment: AlignmentType.CENTER,
//                   children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "GRADE", bold: true, size: 20 })] })]
//                 }),
//               ],
//             }),
//             // Data Rows (Mapping the results you already calculated)
//             ...results.map((r: any) => new TableRow({
//               children: [
//                 new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.code, size: 18 })] })] }),
//                 new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r.name, size: 18 })] })] }),
//                 new TableCell({
//                   verticalAlign: VerticalAlign.CENTER,
//                   children: [new Paragraph({
//                     alignment: AlignmentType.CENTER,
//                     children: [new TextRun({ text: r.grade.toString(), bold: true, size: 18 })]
//                   })]
//                 }),
//               ],
//             })),
//           ],
//         }),

//         new Paragraph({
//           spacing: { before: 400 },
//           children: [
//             new TextRun({ text: "GRADING KEY: ", bold: true, size: 14 }),
//             new TextRun({ text: "A: 70-100 | B: 60-69 | C: 50-59 | D: 40-49 | E/F: 0-39", size: 14 }),
//           ],
//         }),

//         // 4. Authentication Footer
//         new Paragraph({
//           spacing: { before: 800 },
//           alignment: AlignmentType.CENTER,
//           children: [new TextRun({ text: "--- This is a provisional statement of results and is not a final certificate ---", italics: true, size: 16 })],
//         }),
//       ],
//     }],
//   });

//   return await Packer.toBuffer(doc);
// };
