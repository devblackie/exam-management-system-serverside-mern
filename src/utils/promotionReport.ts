// serverside/src/utils/promotionReport.ts
import { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
  WidthType, AlignmentType, HeadingLevel, BorderStyle, ImageRun 
} from "docx";
import config from "../config/config";

interface PromotionData {
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

export const generatePromotionWordDoc = async (data: PromotionData): Promise<Buffer> => {
  const { programName, academicYear, yearOfStudy, eligible, blocked, logoBuffer } = data;

  // Formatting constant for reuse
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };

  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        // 1. LOGO (Using the correct ImageRun structure)
        ...(logoBuffer.length > 0 ? [
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
        ] : []),

        // 2. UNIVERSITY HEADERS
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200 },
          children: [
            new TextRun({ text: config.instName.toUpperCase(), bold: true, size: 28 }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "OFFICE OF THE REGISTRAR (ACADEMIC AFFAIRS)", bold: true, size: 20 }),
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
              underline: { type: BorderStyle.SINGLE } 
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
          children: [new TextRun({ text: "1.0 EXECUTIVE SUMMARY", bold: true })]
        }),
        
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: "Description", bold: true })] })] }),
                new TableCell({ margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: "Student Count", bold: true })] })] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ margins: cellMargin, children: [new Paragraph("Eligible for Promotion")] }),
                new TableCell({ margins: cellMargin, children: [new Paragraph(eligible.length.toString())] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ margins: cellMargin, children: [new Paragraph("Blocked / Action Required")] }),
                new TableCell({ margins: cellMargin, children: [new Paragraph(blocked.length.toString())] }),
              ],
            }),
          ],
        }),

        // 6. DETAILED LISTS
        new Paragraph({ 
          heading: HeadingLevel.HEADING_2, 
          spacing: { before: 400, after: 200 },
          children: [new TextRun({ text: "2.0 DETAILED PROMOTION LIST", bold: true })]
        }),
        createStudentTable([...eligible, ...blocked]),

        // 7. SIGNATORIES
        new Paragraph({
          spacing: { before: 1200 },
          children: [
            new TextRun({ text: "PREPARED BY: __________________________\t\tDATE: _______________", bold: true }),
          ],
        }),
        new Paragraph({ text: "FACULTY COORDINATOR" }),
        
        new Paragraph({
          spacing: { before: 800 },
          children: [
            new TextRun({ text: "APPROVED BY: __________________________\t\tDATE: _______________", bold: true }),
          ],
        }),
        new Paragraph({ text: "CHAIRMAN, ACADEMIC BOARD" }),
      ],
    }],
  });

  return await Packer.toBuffer(doc);
};

function createStudentTable(students: any[]) {
  const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
  
  const headerRow = new TableRow({
    children: [
      new TableCell({ shading: { fill: "E0E0E0" }, margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: "Reg No", bold: true })] })] }),
      new TableCell({ shading: { fill: "E0E0E0" }, margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: "Full Name", bold: true })] })] }),
      new TableCell({ shading: { fill: "E0E0E0" }, margins: cellMargin, children: [new Paragraph({ children: [new TextRun({ text: "Decision/Reasons", bold: true })] })] }),
    ],
  });

  const dataRows = students.map(s => new TableRow({
    children: [
      new TableCell({ margins: cellMargin, children: [new Paragraph(s.regNo)] }),
      // new TableCell({ margins: cellMargin, children: [new Paragraph(s.name)] }),
      new TableCell({ margins: cellMargin, children: [new Paragraph(formatStudentName(s.name))] }),
      new TableCell({ margins: cellMargin, children: [new Paragraph(s.status === "IN GOOD STANDING" ? "PROMOTED" : (s.reasons?.join(", ") || s.status))] }),
    ],
  }));

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow, ...dataRows],
  });
}