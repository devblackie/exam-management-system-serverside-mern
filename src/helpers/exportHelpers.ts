import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
import { IUnit } from "../models/Unit";
import { IMarks } from "../models/Mark";

export async function exportMarksToExcel(unit: IUnit, marks: IMarks[]) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Marks");

  sheet.columns = [
    { header: "Reg No", key: "regNo", width: 20 },
    { header: "Name", key: "name", width: 25 },
    { header: "CAT 1", key: "cat1", width: 10 },
    { header: "CAT 2", key: "cat2", width: 10 },
    { header: "CAT 3", key: "cat3", width: 10 },
    { header: "Assignment", key: "assignment", width: 12 },
    { header: "Practical", key: "practical", width: 12 },
    { header: "Exam", key: "exam", width: 10 },
    { header: "Total", key: "total", width: 10 },
  ];

  marks.forEach((m) => {
    const student: any = m.student;
    sheet.addRow({
      regNo: student.regNo,
      name: student.name,
      cat1: m.cat1 || "-",
      cat2: m.cat2 || "-",
      cat3: m.cat3 || "-",
      assignment: m.assignment || "-",
      practical: m.practical || "-",
      exam: m.exam || "-",
      total: m.total || "-",
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}

export async function exportMarksToPDF(unit: IUnit, marks: IMarks[]) {
  const doc = new PDFDocument({ margin: 30 });
  const chunks: any[] = [];

  doc.on("data", (chunk) => chunks.push(chunk));
  doc.on("end", () => {});

  doc.fontSize(16).text(`Marks Report for ${unit.name} (${unit.code})`, {
    align: "center",
  });
  doc.moveDown();

  doc.fontSize(10);
  marks.forEach((m) => {
    const student: any = m.student;
    doc.text(
      `${student.regNo} - ${student.name} | Total: ${
        m.total?.toFixed(2) || "-"
      }`
    );
  });

  doc.end();

  return new Promise<Buffer>((resolve) => {
    const result = Buffer.concat(chunks);
    resolve(result);
  });
}
