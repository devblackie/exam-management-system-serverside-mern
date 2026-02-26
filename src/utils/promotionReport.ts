// serverside/src/utils/promotionReport.ts
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  TableLayoutType, AlignmentType, HeadingLevel, BorderStyle, ImageRun, VerticalAlign,
} from "docx";
import config from "../config/config";

export interface PromotionData {
  programName: string;  academicYear: string;  yearOfStudy: number;  
  eligible: any[];  blocked: any[];  logoBuffer: Buffer; offeredUnits?: { code: string; name: string }[];
}

const numberToWords = (num: number): string => {
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  
  if (num === 0) return "Zero";
  if (num < 20) return ones[num];
  const digit = num % 10;
  return tens[Math.floor(num / 10)] + (digit ? "-" + ones[digit] : "");
};

// Helper to convert Year 1 to "First Year", etc.
const getOrdinalYear = (year: number): string => {
  const ordinals = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth"];
  return ordinals[year] || `${year}th`;
};

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
  data: PromotionData ): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, eligible, blocked, logoBuffer, offeredUnits = [] } = data;

  const currentYearOrdinal = getOrdinalYear(yearOfStudy);

  const stats: Record<string, number> = {
    "PASS": eligible.length,
    "SUPPLEMENTARY": blocked.filter(s => s.status === "SUPPLEMENTARY").length,
    "SUPPLEMENTARY (After Readmission)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r: string) => r.toLowerCase().includes("readmission"))).length,
    "SUPPLEMENTARY (After Stayout)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r: string) => r.toLowerCase().includes("stayout"))).length,
    "SUPPLEMENTARY (After Carry Forward)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r: string) => r.toLowerCase().includes("carry forward"))).length,
    "ACADEMIC LEAVE": blocked.filter(s => s.status === "ACADEMIC LEAVE").length,
    // "SPECIALS": blocked.filter(s => s.reasons?.some((r: string) => r.toLowerCase().includes("special"))).length,
    "SPECIALS (FINANCIAL GROUNDS)": blocked.filter(s => s.reasons?.some((r: string) => r.toLowerCase().includes("special") && r.toLowerCase().includes("financial"))).length,
    "SPECIALS (COMPASSIONATE GROUNDS)": blocked.filter(s => s.reasons?.some((r: string) => r.toLowerCase().includes("special") && r.toLowerCase().includes("compassionate"))).length,
    "STAYOUT": blocked.filter(s => s.status === "STAYOUT").length,
    "DISCONTINUATION": blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED").length,
    "DEREGISTRATION": blocked.filter(s => s.status === "DEREGISTERED").length,
    "REPEAT YEAR": blocked.filter(s => s.status === "REPEAT YEAR").length,
    "INCOMPLETE": blocked.filter(s => s.status === "INCOMPLETE").length,
  };

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [ ...(logoBuffer.length > 0 ? [ new Paragraph({ alignment: AlignmentType.CENTER, children: [ new ImageRun({ data: logoBuffer, transformation: { width: 120, height: 70 }, type: "png", })] })] : []),

      // 2. Headers
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [ new TextRun({ text: config.instName.toUpperCase(), bold: true, size: 23 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.schoolName.toUpperCase(), bold: true, size: 23 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.departmentName.toUpperCase(), bold: true, size: 23 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: `PROGRAM: ${programName.toUpperCase()}`, bold: true, size: 23 })]  }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: "ORDINARY EXAMINATION RESULTS", bold: true, size: 23})] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [ new TextRun({ text: `${academicYear} ACADEMIC YEAR`, bold: true, size: 23 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:100, after: 100 }, children: [ new TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 23 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:100, after: 300 }, children: [ new TextRun({ text: "SUMMARY", bold: true, size: 23, underline: {} })] }),
      
      createSummaryTable(stats),

      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400, after: 100 }, children: [new TextRun({ text: "UNITS OFFERED", bold: true, size: 22, underline: {} })] }),
        createOfferedUnitsTable(offeredUnits),
        ...createDocFooter(),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

function createSummaryTable(stats: Record<string, number>) {
  // Filter out 0 values to make the table dynamic
  const activeRows = Object.entries(stats).filter(([_, val]) => val > 0);
  const totalCount = Object.values(stats).reduce((a, b) => a + b, 0);

  const rows = [
    ...activeRows.map(([label, val]) => new TableRow({
      children: [
        new TableCell({ 
          width: { size: 70, type: WidthType.PERCENTAGE },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          children: [new Paragraph({ children: [new TextRun({ text: label.toUpperCase(), size: 20 })] })] 
        }),
        new TableCell({ 
          width: { size: 30, type: WidthType.PERCENTAGE },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          children: [new Paragraph({ children: [new TextRun({ text: val.toString(), size: 20 })] })] 
        }),
      ]
    })),
    // Total Row
    new TableRow({
      children: [
        new TableCell({ 
          borders: { top: { style: BorderStyle.SINGLE, size: 1 }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          children: [new Paragraph({ children: [new TextRun({ text: "TOTAL", bold: true, size: 20 })] })] 
        }),
        new TableCell({ 
          borders: { top: { style: BorderStyle.SINGLE, size: 1 }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          children: [new Paragraph({ children: [new TextRun({ text: totalCount.toString(), bold: true, size: 20 })] })] 
        }),
      ]
    })
  ];

  return new Table({ width: { size: 60, type: WidthType.PERCENTAGE }, alignment: AlignmentType.CENTER, rows: rows });
}

function createOfferedUnitsTable(units: { code: string; name: string }[]) {
  if (!units || units.length === 0) return new Paragraph("No units recorded.");

  const cellMargin = { top: 50, bottom: 50, left: 100, right: 100 };
  const midPoint = Math.ceil(units.length / 2);
  const leftCol = units.slice(0, midPoint);
  const rightCol = units.slice(midPoint);

  const headerCell = (text: string) => new TableCell({
    margins: cellMargin,
    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text, bold: true, size: 22 })] })]
  });

  const headerRow = new TableRow({
    children: [
      headerCell("S/NO."), headerCell("CODE"), headerCell("NAME"),
      headerCell("S/NO."), headerCell("CODE"), headerCell("NAME"),
    ]
  });

  const dataRows = [];
  for (let i = 0; i < midPoint; i++) {
    const left = leftCol[i];
    const right = rightCol[i];

    dataRows.push(new TableRow({
      children: [
        // Left Side
        new TableCell({ margins: cellMargin, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (i + 1).toString(), size: 21 })] })] }),
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: left.code, size: 21 })] })] }),
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: left.name, size: 21 })] })] }),
        // Right Side
        new TableCell({ margins: cellMargin, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: right ? (midPoint + i + 1).toString() : "", size: 21 })] })] }),
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: right?.code || "", size: 21 })] })] }),
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: right?.name || "", size: 21 })] })] }),
      ]
    }));
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows]
  });
}

export const generateEligibleSummaryDoc = async (
  data: any,
): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, eligible, logoBuffer } = data;
  
  const candidateCountWords = numberToWords(eligible.length);
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);
  const nextYearOrdinal = getOrdinalYear(yearOfStudy + 1);

  const cellMargin = { top: 0, bottom: 0, left: 100, right: 100 };

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          // 1. LOGO
          ...(logoBuffer.length > 0 ? [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [ new ImageRun({ data: logoBuffer, transformation: { width: 120, height: 70 }, type: "png", }),],
                }),
              ] : []),

          // 2. Headers
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [ new TextRun({ text: config.instName.toUpperCase(), bold: true, size: 23 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.schoolName.toUpperCase(), bold: true, size: 23 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.departmentName.toUpperCase(), bold: true, size: 23 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: `PROGRAM: ${programName.toUpperCase()}`, bold: true, size: 23 })]  }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: "ORDINARY EXAMINATION RESULTS", bold: true, size: 23})] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [ new TextRun({ text: `${academicYear} ACADEMIC YEAR`, bold: true, size: 23 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:100, after: 100 }, children: [ new TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 23 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before:100, after: 300 }, children: [ new TextRun({ text: "PASS", bold: true, size: 23, underline: {} })] }),
          

          // 4. INTRODUCTORY TEXT
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 400, after: 300 },
            children: [
              new TextRun({ text: `The following `, size: 22 }),
              new TextRun({ text: `${candidateCountWords} (${eligible.length}) `, bold: true, size: 22  }),
              new TextRun({ text: `candidates satisfied the ${config.schoolName} Board of Examiners in the `, size: 22   }),
              new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
              new TextRun({ text: `Academic Year, `, size: 22 }),
              new TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22  }),
              new TextRun({ text: `Examinations for the `, size: 22 }),
              new TextRun({ text: `${programName}. `, bold: true, size: 22  }),
              new TextRun({ text: `The ${config.schoolName} Board of Examiners recommends that they proceed to their `, size: 22 }),         
              new TextRun({ text: `${nextYearOrdinal} Year `, bold: true, size: 22 }),
              new TextRun({ text: `of study.`, size: 22 }),
            ],
          }),

          // 5. THE PASS LIST TABLE
          createPassTable(eligible, cellMargin),

          // 6. SIGNATORIES (Aligned with the image layout)
          ...createDocFooter(),         
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

function createPassTable(students: any[], cellMargin: any) {
  const headerRow = new TableRow({
    children: [
      { text: "S/No.", width: 5 },
      { text: "REG. NO.", width: 30 }, { text: "NAME", width: 65 }
    ].map(
      (col) =>
        new TableCell({
          width: { size: col.width, type: WidthType.PERCENTAGE },
          margins: cellMargin,
          children: [
            new Paragraph({
              spacing: { before: 0, after: 0 }, // REMOVE PARAGRAPH SPACING
              children: [new TextRun({ text: col.text, bold: true, size: 18 })],
            }),
          ],
        }),
    ),
  });

  const dataRows = students.map(
    (s, index) =>
      new TableRow({
        children: [ 
          new TableCell({ margins: cellMargin, children: [ new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: (index + 1).toString(), size: 20 })] })] }),
          new TableCell({ margins: cellMargin, children: [ new Paragraph({ spacing: { before: 0, after: 0 },  children: [new TextRun({ text: s.regNo, size: 20 })], })] }),
          new TableCell({ margins: cellMargin, children: [ new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: formatStudentName(s.name), size: 20 })]})] })] }),
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
    },
    rows: [headerRow, ...dataRows],
  });
}

const filterSpecialsByGrounds = (students: any[], ground: "Financial" | "Compassionate") => {
  return students.filter((s) => {    
    const hasSpecialStatus = s.status?.includes("SPECIAL");
    const reasonsStr = s.reasons?.join(" ").toLowerCase() || "";
    
    const isSpecial = hasSpecialStatus || reasonsStr.includes("special");
    if (!isSpecial) return false;

    if (ground === "Financial") return reasonsStr.includes("financial") || !reasonsStr.includes("compassionate");
    if (ground === "Compassionate") return reasonsStr.includes("compassionate");
    return false;
  });
};

export const generateSpecialExamsDoc = async (
  data: PromotionData, groundType: "Financial" | "Compassionate"  = "Financial"
): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  
  // Filter only for Special Exam candidates
  const specialList = filterSpecialsByGrounds(blocked, groundType);
  
  const count = specialList.length;
  const candidateCountWords = numberToWords(specialList.length);
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, `SPECIAL (${groundType.toUpperCase()})`),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: `The following `, size: 22 }),
            new TextRun({ text: `${candidateCountWords} (${count}) `, bold: true, size: 22  }),
            new TextRun({ text: `candidate(s) have special examinations, on `, size: 22   }),
            new TextRun({ text: `${groundType} Grounds `, bold:true, size: 22   }),
            new TextRun({ text: `in the unit(s) indicated against their names during the `, size: 22   }),
            new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
            new TextRun({ text: `Academic Year, `, size: 22 }),
            new TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22  }),
            new TextRun({ text: `Examinations for the `, size: 22 }),
            new TextRun({ text: `${programName}. `, bold: true, size: 22  }),
            new TextRun({ text: `The ${config.schoolName} Board of Examiners upholds the decision of the Dean’s Committee. `, size: 22 }),         
          ],
        }),

        createAcademicTable(specialList, cellMargin, groundType),
        ...createDocFooter(),
      ],
    }],
  });

  return await Packer.toBuffer(doc);
};

export const generateSupplementaryExamsDoc = async ( data: PromotionData ): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const failureList = blocked.filter((s) => s.status === "SUPPLEMENTARY");
  const count = failureList.length;
  const candidateCountWords = numberToWords(failureList.length);
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "SUPPLEMENTARY"),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED, spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: `The following `, size: 22 }),
            new TextRun({ text: `${candidateCountWords} (${count}) `, bold: true, size: 22  }),
            new TextRun({ text: `candidate(s) failed to satisfy the ${config.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22   }),
            new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
            new TextRun({ text: `Academic Year, `, size: 22 }),
            new TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22  }),
            new TextRun({ text: `Examinations for the `, size: 22 }),
            new TextRun({ text: `${programName}. `, bold: true, size: 22  }),
            new TextRun({ text: `The ${config.schoolName} Board of Examiners recommends that they sit for the supplementary exams when next offered. `, size: 22 }),         
          ],
        }),

        createAcademicTable(failureList, cellMargin),
        ...createDocFooter(),
      ],
    }],
  });

  return await Packer.toBuffer(doc);
};

// function createAcademicTable(students: any[], cellMargin: any, groundType?: string) {
//   const headerRow = new TableRow({
//     children: [ { text: "S/No", width: 5 }, { text: "Reg No.", width: 20 }, { text: "Name", width: 35 }, { text: "Unit Code", width: 15 }, { text: "Unit Name", width: 25 } ].map((col) =>
//       new TableCell({
//         width: { size: col.width, type: WidthType.PERCENTAGE },
//         margins: cellMargin,
//         children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: col.text, bold: true, size: 18 })] })],
//       })
//     ),
//   });

//   const rows: TableRow[] = [headerRow];
//   let studentCounter = 1; // Increments per student, not per unit

//   students.forEach((s) => {
//     // Filter reasons to only include the ones relevant to this document
//     const relevantReasons = s.reasons?.filter((r: string) => {
//       if (!groundType) return !r.toLowerCase().includes("special"); // Supplementary logic
//       return r.toLowerCase().includes("special") && r.toLowerCase().includes(groundType.toLowerCase()); }) || [];

//     // Only process the student if they actually have relevant units
//     if (relevantReasons.length > 0) {
//       relevantReasons.forEach((rawReason: string, index: number) => {
//         const isFirstUnit = index === 0; // Check if this is the first unit for this specific student

//         let parts = rawReason.split(":").map((p: string) => p.trim());

//         // 1. Strip the "SPECIAL" or "FAILED" prefix
//         if (["failed", "special", "retake"].includes(parts[0].toLowerCase())) { parts.shift(); }

//         // 2. Extract Unit Code & Name
//         const uCode = parts[0] || "N/A";
//         let uName = parts.slice(1).join(": ") || "N/A";
        
//         // Clean Unit Name
//         uName = uName
//           .replace(/\s*-\s*SPECIAL\b/gi, "").replace(/\bSPECIAL\b/gi, "").replace(new RegExp(`\\s*[:-]?\\s*${groundType}\\s*Grounds?`, "gi"), "").trim();

//         rows.push(new TableRow({
//           children: [ isFirstUnit ? studentCounter.toString() : "", isFirstUnit ? s.regNo : "", isFirstUnit ? formatStudentName(s.name) : "", uCode, uName ].map(val =>
//             new TableCell({ margins: cellMargin, children: [ new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: val, size: 18 })] }) ]})
//           ),
//         }));
//       });

//       studentCounter++; // Move to next serial number after all units for this student are listed
//     }
//   });

//   return new Table({
//     width: { size: 100, type: WidthType.PERCENTAGE },
//     borders: {
//       top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
//       left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
//       insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
//     },
//     rows: rows,
//   });
// }

function createAcademicTable( students: any[], cellMargin: any, groundType?: "Financial" | "Compassionate" ) {
  const headerRow = new TableRow({
    children: ["S/No", "Reg No.", "Name", "Unit Code", "Unit Name"].map(
      (text, idx) => {
        const widths = [5, 20, 30, 15, 30];
        return new TableCell({
          width: { size: widths[idx], type: WidthType.PERCENTAGE },
          margins: cellMargin,
          children: [ new Paragraph({ children: [new TextRun({ text, bold: true, size: 18 })]}),
          ],
        });
      },
    ),
  });

  const rows: TableRow[] = [headerRow];
  let studentCounter = 1;

  students.forEach((s) => {
    // 1. FILTERING REASONS
    const relevantReasons =
      s.reasons?.filter((r: string) => {
        const lowerR = r.toLowerCase();
        if (!groundType) {
        return !lowerR.includes("special") && !lowerR.includes("incomplete");
      }
      // Specials list MUST match the ground
      return lowerR.includes("special") && lowerR.includes(groundType.toLowerCase());
    }) || [];

    if (relevantReasons.length > 0) {
      relevantReasons.forEach((rawReason: string, index: number) => {
        const isFirstUnit = index === 0;
        let uCode = "N/A";
        let uName = "N/A";
        const colonIndex = rawReason.indexOf(":");
        if (colonIndex !== -1) {
          uCode = rawReason.substring(0, colonIndex).trim();
          // Extract name and strip everything after the first dash or parenthesis
          let remainder = rawReason.substring(colonIndex + 1).trim();
          uName = remainder.split(/[(\-]/)[0].trim();
        }

        rows.push(
          new TableRow({
            children: [ isFirstUnit ? studentCounter.toString() : "", isFirstUnit ? s.regNo : "", isFirstUnit ? s.name : "", uCode, uName ].map(
              (val) =>
                new TableCell({ margins: cellMargin, children: [ new Paragraph({ children: [new TextRun({ text: val, size: 18 })]}) ]}),
            ),
          }),
        );
      });
      studentCounter++;
    }
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows,
  });
}


function createDocHeader(   logo: any,   program: string,   year: string,   ordinal: string,   type: string, ) {
  return [
    ...(logo && logo.length > 0 ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [ new ImageRun({ data: logo, transformation: { width: 120, height: 70 }, type: "png", }), ],
          }),
        ] : []),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [ new TextRun({ text: config.instName.toUpperCase(), bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.schoolName.toUpperCase(), bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: config.departmentName.toUpperCase(), bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: `PROGRAM: ${program.toUpperCase()}`, bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: "ORDINARY EXAMINATION RESULTS", bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: `${year} ACADEMIC YEAR`, bold: true, size: 23 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [ new TextRun({ text: `${ordinal} Year `, bold: true, size: 23 })] }),     
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 100 }, children: [ new TextRun({ text: type, bold: true, size: 23, underline: {} })] }),    
  ];
}

function createDocFooter() {
    return [
        new Paragraph({ spacing: { before: 900 }, children: [new TextRun({ text: `APPROVED BY THE BOARD OF EXAMINERS, ${config.schoolName.toUpperCase()}`, bold: true, size: 18 })] }),
        new Paragraph({ spacing: { before: 400 }, children: [new TextRun({ text: "SIGNED: __________________________\t\tDATE: _______________", bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: `\tDEAN, ${config.schoolName.toUpperCase()}`, size: 18 })] }),
    ];
}


export const generateStayoutExamsDoc = async ( data: PromotionData ): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  
  // Filter for students whose status is specifically "STAYOUT"
  const stayoutList = blocked.filter((s) => s.status === "STAYOUT");
  const count = stayoutList.length;
  const candidateCountWords = numberToWords(count);
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [
      {
        children: [
          ...createDocHeader(
            logoBuffer,
            programName,
            academicYear,
            currentYearOrdinal,
            "STAY OUT / RETAKE",
          ),

          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 400, after: 300 },
            children: [
              new TextRun({ text: `The following `, size: 22 }),
              new TextRun({
                text: `${candidateCountWords} (${count}) `,
                bold: true,
                size: 22,
              }),
              new TextRun({
                text: `candidate(s) failed more than one-third (1/3) but less than half (1/2) of the prescribed units in the `,
                size: 22,
              }),
              new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
              new TextRun({ text: `Academic Year, `, size: 22 }),
              new TextRun({
                text: `${currentYearOrdinal} Year `,
                bold: true,
                size: 22,
              }),
              new TextRun({ text: `Examinations for the `, size: 22 }),
              new TextRun({ text: `${programName}. `, bold: true, size: 22 }),
              new TextRun({
                text: `In accordance with ENG Rule 15 (h) “A candidate who fails more than a third and less than a half of the prescribed units in any year of study shall be required to retake examinations only in the failed units during the ordinary examination period when examinations for the individual units are offered. Such a candidate will not be allowed to retake examinations during the supplementary period immediately following the ordinary examinations period in which he/she failed the units”.`,
                size: 22, bold: true, italics: true,
              }),
            ],
          }),

          createFailureAnalysisTable(stayoutList, { top: 100, bottom: 100, left: 100, right: 100 }),
        ...createDocFooter(),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
};

export const generateRepeatYearDoc = async (data: PromotionData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const list = blocked.filter((s) => s.status === "REPEAT YEAR");
  const count = list.length;
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "REPEAT YEAR"),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: `The following `, size: 22 }),
            new TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
            new TextRun({ text: `candidate(s) failed fifty percent (50%) or more of the units or obtained a mean mark of less than 40% in the `, size: 22 }),
            new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
            new TextRun({ text: `Academic Year. In accordance with ENG Rule 16, they are required to `, size: 22 }),
            new TextRun({ text: `REPEAT THE YEAR `, bold: true, size: 22 }),
            new TextRun({ text: `and attend classes in all the failed units.`, size: 22 }),
          ],
        }),
        createFailureAnalysisTable(list, { top: 100, bottom: 100, left: 100, right: 100 }),
        ...createDocFooter(),
      ],
    }],
  });
  return await Packer.toBuffer(doc);
};

// B. For Repeat Year / Stayout (Unit Failure focused)
function createFailureAnalysisTable(students: any[], cellMargin: any) {
  const rows: TableRow[] = [
    new TableRow({
      children: ["S/No", "Reg No.", "Name", "Units Failed", "Mean Mark"].map(h => 
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18 })] })] })
      )
    })
  ];

  students.forEach((s, i) => {
    // Extract unit codes only for a compact view
    const unitCodes = s.reasons
      ?.filter((r: string) => !r.toLowerCase().includes("special"))
      .map((r: string) => r.split(":")[0])
      .join(", ");

    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.regNo, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.name, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: unitCodes || "N/A", size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.summary?.weightedMean || "N/A", size: 18 })] })] }),
      ]
    }));
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

// ---- Academic Leave and Deferment Block -----

// 
export const generateAcademicLeaveDoc = async (data: PromotionData, type: "ACADEMIC LEAVE" | "DEFERMENT"): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const list = blocked.filter((s) => s.status === type);

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, getOrdinalYear(yearOfStudy), type),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: `The following candidate(s) have been officially granted `, size: 22 }),
            new TextRun({ text: `${type} `, bold: true, size: 22 }),
            new TextRun({ text: `for the `, size: 22 }),
            new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
            new TextRun({ text: `Academic Year. They are expected to resume studies at the beginning of the next academic cycle.`, size: 22 }),
          ],
        }),
        createAdministrativeTable(list, { top: 100, bottom: 100, left: 100, right: 100 }),
        ...createDocFooter(),
      ],
    }],
  });
  return await Packer.toBuffer(doc);
};



//  For Deferment / Academic Leave (Status focused)
function createAdministrativeTable(students: any[], cellMargin: any) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: ["S/No", "Reg No.", "Name", "Effective Date", "Remarks"].map(h => 
          new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18 })] })] })
        )
      }),
      ...students.map((s, i) => new TableRow({
        children: [

          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.regNo, size: 18 })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.name, size: 18 })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.effectiveDate || "N/A", size: 18 })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.remarks || "Approved", size: 18 })] })] }),
        ]
      }))
    ]
  });
}
// ---- Academic Leave and Deferment Block end ----

//  ---- Carry Forward Block -----

export const generateCarryForwardDoc = async (data: PromotionData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, eligible, logoBuffer } = data;

  // Filter for students who are promoted but have carry-over units in their reasons
  const carryForwardList = eligible.filter((s) => 
    s.reasons?.length > 0 && s.status !== "ALREADY PROMOTED"
  );

  const count = carryForwardList.length;
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);
  const nextYearOrdinal = getOrdinalYear(yearOfStudy + 1);

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "CARRY FORWARD"),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: `The following `, size: 22 }),
            new TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
            new TextRun({ 
              text: `candidate(s) satisfied the Board of Examiners in at least two-thirds of the units. In accordance with `, 
              size: 22 
            }),
            new TextRun({ text: `ENG Rule 13 (e)`, bold: true, size: 22 }),
            new TextRun({ 
              text: `, they are allowed to proceed to `, 
              size: 22 
            }),
            new TextRun({ text: `${nextYearOrdinal} Year `, bold: true, size: 22 }),
            new TextRun({ 
              text: `but MUST carry forward the failed units indicated against their names to be taken when next offered.`, 
              size: 22 
            }),
          ],
        }),

        // HERE IS WHERE YOU USE THE TABLE FUNCTION
        createCarryForwardTable(carryForwardList, { top: 100, bottom: 100, left: 100, right: 100 }),

        ...createDocFooter(),
      ],
    }],
  });

  return await Packer.toBuffer(doc);
};

// C. For Carry Forward (Progressive tracking)
function createCarryForwardTable(students: any[], cellMargin: any) {
  const rows: TableRow[] = [
    new TableRow({
      children: ["S/No", "Reg No.", "Name", "Carry Over Units", "New Year"].map(h => 
        new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18 })] })] })
      )
    })
  ];

  students.forEach((s, i) => {
    const carryUnits = s.reasons?.map((r: string) => r.split(":")[0]).join(", ");
    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ text: (i + 1).toString() })] }),
        new TableCell({ children: [new Paragraph({ text: s.regNo })] }),
        new TableCell({ children: [new Paragraph({ text: s.name })] }),
        new TableCell({ children: [new Paragraph({ text: carryUnits })] }),
        new TableCell({ children: [new Paragraph({ text: (s.currentYearOfStudy + 1).toString() })] }),
      ]
    }));
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}
//  ---- Carry Forward Block end  -----

// ---- Discontinuation Block ----
export const generateDiscontinuationDoc = async (data: PromotionData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const list = blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED");
  const count = list.length;
  const currentYearOrdinal = getOrdinalYear(yearOfStudy);

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "DISCONTINUATION"),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: "The following ", size: 22 }),
            new TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
            new TextRun({ text: "candidate(s) failed to satisfy the Board of Examiners in the unit(s) indicated against their names on the maximum allowed attempts. In accordance with ", size: 22 }),
            new TextRun({ text: "ENG Rule 22", bold: true, size: 22 }),
            new TextRun({ text: ", the Board recommends that they be ", size: 22 }),
            new TextRun({ text: "DISCONTINUED ", bold: true, size: 22, color: "FF0000" }),
            new TextRun({ text: "from the program of study.", size: 22 }),
          ],
        }),
        createDiscontinuationTable(list, { top: 100, bottom: 100, left: 100, right: 100 }),
        ...createDocFooter(),
      ]
    }]
  });
  return await Packer.toBuffer(doc);
};

// D. For Discontinuation (Focus on Attempt History)
function createDiscontinuationTable(students: any[], cellMargin: any) {
  const rows: TableRow[] = [
    new TableRow({
      children: ["S/No", "Reg No.", "Name", "Unit(s) Failed on Max Attempts", "Total Attempts"].map(h => 
        new TableCell({ 
          margins: cellMargin, 
          shading: { fill: "F2F2F2" },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18 })] })] 
        })
      )
    })
  ];

  students.forEach((s, i) => {
    // Extract units where attempt count is high (Critical Failures)
    const criticalUnits = s.reasons
      ?.filter((r: string) => r.toLowerCase().includes("critical") || r.toLowerCase().includes("attempt: 3") || r.toLowerCase().includes("attempt: 4"))
      .map((r: string) => r.split(":")[0])
      .join(", ");

    rows.push(new TableRow({
      children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.regNo, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.name, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: criticalUnits || "Multiple Failures", size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Max Reached", size: 18, bold: true })] })] }),
      ]
    }));
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

//  ---- Discontinuation Block end ----

// ---- Deregistration Block ---

export const generateDeregistrationDoc = async (data: PromotionData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, blocked, logoBuffer } = data;
  const list = blocked.filter(s => s.status === "DEREGISTERED");
  const count = list.length;

  const doc = new Document({
    sections: [{
      children: [
        ...createDocHeader(logoBuffer, programName, academicYear, getOrdinalYear(yearOfStudy), "DEREGISTRATION"),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 400, after: 300 },
          children: [
            new TextRun({ text: "In accordance with ", size: 22 }),
            new TextRun({ text: "ENG Rule 23 (c)", bold: true, size: 22 }),
            new TextRun({ text: ", the following candidate(s) were absent from six (6) or more examinations in the ", size: 22 }),
            new TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
            new TextRun({ text: "Academic Year without official permission. They are therefore deemed to have deserted the program and are hereby ", size: 22 }),
            new TextRun({ text: "DEREGISTERED.", bold: true, size: 22 }),
          ],
        }),
        createDeregistrationTable(list, { top: 100, bottom: 100, left: 100, right: 100 }),
        ...createDocFooter(),
      ]
    }]
  });
  return await Packer.toBuffer(doc);
};

// E. For Deregistration (Focus on Absence/Desertion)
function createDeregistrationTable(students: any[], cellMargin: any) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: ["S/No", "Reg No.", "Name", "Missing Units Count", "Status"].map(h => 
          new TableCell({ 
            margins: cellMargin, 
            shading: { fill: "F2F2F2" },
            children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18 })] })] 
          })
        )
      }),
      ...students.map((s, i) => new TableRow({
        children: [
          new TableCell({ children: [new Paragraph ({ children: [new TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
          new TableCell({ children: [new Paragraph ({ children: [new TextRun({ text: s.regNo, size: 18 })] })] }),
          new TableCell({ children: [new Paragraph ({ children: [new TextRun({ text: s.name, size: 18 })] })] }),
          new TableCell({ children: [new Paragraph ({ children: [new TextRun({ text: s.summary?.missing?.toString() || "6+", size: 18 })] })] }),
          new TableCell({ children: [new Paragraph ({ children: [new TextRun({ text: "DEEMED DESERTED", size: 16, bold: true })] })] }),
        ]
      }))
    ]
  });
}

// ---- Deregistration Block end ----

// graduation-list
export const generateAwardListDoc = async ( data: PromotionData ): Promise<Buffer> => {
  const { programName, academicYear, eligible, logoBuffer } = data;

  const doc = new Document({
    sections: [
      {
        children: [
          ...createDocHeader( logoBuffer, programName, academicYear, "Final", "AWARD LIST" ),
          new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 400, after: 300 },
            children: [
              new TextRun({
                text: `The following candidates satisfied the Board of Examiners in all the prescribed units for the four/five years of study. The Board of Examiners recommends that they be `,
                size: 22,
              }),
              new TextRun({ text: `AWARDED THE DEGREE OF ${programName.toUpperCase()}.`, bold: true, size: 22 }),
            ],
          }),
          createAwardTable(eligible),
          ...createDocFooter(),
        ],
      },
    ],
  });
  return await Packer.toBuffer(doc);
};

function createAwardTable(students: any[]) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: ["S/No", "Reg No.", "Name", "Classification"].map(
          (h) =>
            new TableCell({ children: [ new Paragraph({ children: [new TextRun({ text: h, bold: true })] }) ] }),
        ),
      }),
      ...students.map((s, i) =>
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: (i + 1).toString() })] }),
              new TableCell({ children: [new Paragraph({ text: s.regNo })] }),
              new TableCell({ children: [new Paragraph({ text: s.name })] }),
              new TableCell({
                children: [new Paragraph({ text: s.classification || "PASS" })],
              }),
            ],
          }),
      ),
    ],
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
                      transformation: { width: 150, height: 80 },
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
                      transformation: { width: 150, height: 80 },
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
  const displayYear =
    data.yearToPromote ||
    data.yearOfStudy ||
    student.currentYearOfStudy ||
    "N/A";

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
                      transformation: { width: 150, height: 80 },
                      type: "png",
                    }),
                  ],
                }),
              ]
            : []),

          // 2. INSTITUTIONAL HEADERS
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 50, after: 100 },
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
            spacing: { after: 50 },
            children: [
              new TextRun({
                text: config.postalAddress,

                size: 18,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 50 },
            children: [
              new TextRun({
                text: "Cell Phone",

                size: 18,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 50 },
            children: [
              new TextRun({
                text: config.cellPhone,

                size: 18,
              }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
            children: [
              new TextRun({
                text: `Email: ${config.schoolEmail}`,

                size: 18,
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "UNDERGRADUATE ACADEMIC TRANSCRIPT",
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
              // top: { style: BorderStyle.NONE },
              top: {
                style: BorderStyle.SINGLE,
                size: 16, // 2pt thickness
                color: "000000", // Optional: black
              },
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
                            text: "NAME: ",
                            bold: true,
                            size: 20,
                          }),
                          new TextRun({
                            text: formatStudentName(student.name).toUpperCase(),
                            bold: false,
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
                            text: "REG NO: ",
                            bold: true,
                            size: 20,
                          }),
                          new TextRun({
                            text: student.regNo,
                            bold: false,
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
                text: "SCHOOL: ",
                bold: true,
                size: 18,
              }),
              new TextRun({
                text: config.schoolName.toUpperCase(),
                bold: false,
                size: 18,
              }),
            ],
          }),

          // 5. Programme name
          new Paragraph({
            spacing: { before: 100 },
            children: [
              new TextRun({
                text: "PROGRAM: ",
                bold: true,
                size: 18,
              }),
              new TextRun({
                text: `${programName.toUpperCase()}`,
                size: 18,
              }),
            ],
          }),

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
                    width: { size: 70, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        spacing: { before: 100, after: 300 },
                        children: [
                          new TextRun({
                            text: "ACADEMIC YEAR: ",
                            bold: true,
                            size: 20,
                          }),
                          new TextRun({
                            text: academicYear || "N/A",
                            bold: false,
                            size: 20,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                          new TextRun({
                            text: "YEAR OF STUDY: ",
                            bold: true,
                            size: 20,
                          }),
                          new TextRun({
                            text: `${displayYear}`,
                            bold: false,
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

          //  STATUS SUMMARY
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "RESULT:  PASS",
                bold: true,
                size: 20,
                underline: {},
              }),
            ],
          }),

          // 4. RESULTS TABLE (Units & Grades)
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            borders: {
              top: { style: BorderStyle.SINGLE, size: 2 },
              bottom: { style: BorderStyle.SINGLE, size: 2 },
              left: { style: BorderStyle.SINGLE, size: 2 },
              right: { style: BorderStyle.SINGLE, size: 2 },
              insideVertical: { style: BorderStyle.SINGLE, size: 2 },
              insideHorizontal: { style: BorderStyle.NIL },
            },
            rows: [
              // Header Row
              new TableRow({
                tableHeader: true,
                children: [
                  new TableCell({
                    width: { size: 15, type: WidthType.PERCENTAGE }, // Narrower Code
                    margins: cellMargin,
                    verticalAlign: VerticalAlign.CENTER,
                    borders: {
                      bottom: { style: BorderStyle.SINGLE, size: 2 },
                      top: { style: BorderStyle.SINGLE, size: 2 },
                    },
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: "CODE", bold: true, size: 18 }),
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    width: { size: 70, type: WidthType.PERCENTAGE }, // Larger Title
                    margins: cellMargin,
                    verticalAlign: VerticalAlign.CENTER,
                    borders: {
                      bottom: { style: BorderStyle.SINGLE, size: 2 },
                      top: { style: BorderStyle.SINGLE, size: 2 },
                    },
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
                    width: { size: 15, type: WidthType.PERCENTAGE }, // Narrower Grade
                    margins: cellMargin,
                    verticalAlign: VerticalAlign.CENTER,
                    borders: {
                      bottom: { style: BorderStyle.SINGLE, size: 2 },
                      top: { style: BorderStyle.SINGLE, size: 2 },
                    },
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
              ...results.map((r: any, index: number) => {
                const unitCode = String(r?.code ?? "N/A").toUpperCase();
                const unitName = String(
                  r?.name ?? "COURSE TITLE MISSING",
                ).toUpperCase();
                const unitGrade = String(r?.grade ?? "-");

                const isLastRow = index === results.length - 1;
                const bottomBorderStyle = isLastRow
                  ? BorderStyle.SINGLE
                  : BorderStyle.NIL;
                const bottomBorderSize = isLastRow ? 2 : 0;

                return new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 15, type: WidthType.PERCENTAGE },
                      margins: cellMargin,
                      verticalAlign: VerticalAlign.CENTER,
                      borders: {
                        bottom: {
                          style: bottomBorderStyle,
                          size: bottomBorderSize,
                        },
                      },
                      children: [
                        new Paragraph({
                          spacing: { after: 10 },
                          children: [new TextRun({ text: unitCode, size: 18 })],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 70, type: WidthType.PERCENTAGE },
                      margins: cellMargin,
                      verticalAlign: VerticalAlign.CENTER,
                      borders: {
                        bottom: {
                          style: bottomBorderStyle,
                          size: bottomBorderSize,
                        },
                      },
                      children: [
                        new Paragraph({
                          spacing: { after: 10 },
                          children: [new TextRun({ text: unitName, size: 18 })],
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 15, type: WidthType.PERCENTAGE },
                      margins: cellMargin,
                      verticalAlign: VerticalAlign.CENTER,
                      borders: {
                        bottom: {
                          style: bottomBorderStyle,
                          size: bottomBorderSize,
                        },
                      },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing: { after: 10 },
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

        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: {  before: 200 },
            children: [
            
            ],
          }),

          // 6. GRADING KEY (Formal Table Structure)
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
                  // LEFT CELL: The Grading Key Table
                  new TableCell({
                    width: { size: 40, type: WidthType.PERCENTAGE },
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                          top: { style: BorderStyle.SINGLE, size: 1 },
                          bottom: { style: BorderStyle.SINGLE, size: 1 },
                          left: { style: BorderStyle.SINGLE, size: 1 },
                          right: { style: BorderStyle.SINGLE, size: 1 },
                          insideVertical: {
                            style: BorderStyle.SINGLE,
                            size: 1,
                          },
                          insideHorizontal: { style: BorderStyle.NONE },
                        },
                        rows: [
                          // Header Row
                          new TableRow({
                            tableHeader: true,
                            children: [
                              ["GRADE", 20],
                              ["RANGE", 30],
                              ["DESCRIPTION", 50],
                            ].map(
                              ([text, width]) =>
                                new TableCell({
                                  width: {
                                    size: width as number,
                                    type: WidthType.PERCENTAGE,
                                  },
                                  shading: { fill: "F2F2F2" },
                                  children: [
                                    new Paragraph({
                                      alignment: AlignmentType.CENTER,
                                      children: [
                                        new TextRun({
                                          text: text as string,
                                          bold: true,
                                          size: 16,
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                            ),
                          }),
                          // Data Rows
                          ...[
                            { g: "A", r: "70 - 100%", d: "EXCELLENT" },
                            { g: "B", r: "60 - 69%", d: "GOOD" },
                            { g: "C", r: "50 - 59%", d: "SATISFACTORY" },
                            { g: "D", r: "40 - 49%", d: "PASS" },
                            { g: "E", r: "0 - 39%", d: "FAIL" },
                          ].map(
                            (item) =>
                              new TableRow({
                                children: [
                                  new TableCell({
                                    children: [
                                      new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                          new TextRun({
                                            text: item.g,
                                            size: 16,
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  new TableCell({
                                    children: [
                                      new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                          new TextRun({
                                            text: item.r,
                                            size: 16,
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  new TableCell({
                                    children: [
                                      new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        children: [
                                          new TextRun({
                                            text: item.d,
                                            size: 16,
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                          ),
                        ],
                      }),
                    ],
                  }),
                  // RIGHT CELL: Registration Number
                  new TableCell({
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    verticalAlign: VerticalAlign.BOTTOM, // Keeps it aligned with bottom of grading key
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                          new TextRun({ text: "NB: ", bold: true, size: 20 }),
                        ],
                      }),
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                          new TextRun({
                            text: "1 unit consists of 35 lecture hours or equivalent (3 Practical hours of two tutorial hours are equivalent to 0ne lecture hour ) ",
                            bold: false,
                            size: 16,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          // 7. FOOTER & SIGNATORIES
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              // top: { style: BorderStyle.NONE },
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
                    ],
                  }),
                  new TableCell({
                    children: [
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
                            text: `REGISTRAR, ${config.registrar.toUpperCase()}`,
                            bold: true,
                            size: 18,
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
