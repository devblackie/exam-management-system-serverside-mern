// // src/services/pdfGenerator.ts
// import PDFDocument from "pdfkit";
// import { Response } from "express";
// import FinalGrade from "../models/FinalGrade";
// import Student from "../models/Student";
// import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import Institution from "../models/Institution";

// const LOGO_PATH = "./assets/logo.png";
// const SIGNATURE_PATH = "./assets/signature.png";

// export async function generateStudentTranscript(regNo: string, res: Response,  filterYear?: string): Promise<void> {
//   const student = await Student.findOne({ regNo: regNo.toUpperCase() }).populate("program").lean();
//   if (!student) throw new Error("Student not found");

//     let query: any = { student: student._id };
//  if (filterYear) {
//       const yearDoc = await AcademicYear.findOne({ year: filterYear });
//       if (!yearDoc) throw new Error(`Academic year not found: ${filterYear}`);
//       query.academicYear = yearDoc._id;
//     }

//  const grades = await FinalGrade.find(query) 
//       .populate<{ unit: { code: string; name: string } }>("unit")
//       .populate<{ academicYear: { year: string } }>("academicYear")
//       .sort({ "academicYear.year": 1, "unit.code": 1 })
//       .lean();
      
//     if (grades.length === 0) throw new Error(`No grades found for student: ${regNo} (Year: ${filterYear || 'Full'})`);

//     const doc = new PDFDocument({ margin: 50, size: "A4" });

//     // FIX: Ensure filename matches client-side expectation:
//     const filename = filterYear
//       ? `Transcript_${regNo}_${filterYear.replace("/", "-")}.pdf`
//       : `Transcript_${regNo}_Full.pdf`;

//     res.setHeader("Content-Type", "application/pdf");
//     res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
//     doc.pipe(res);

//   // Header
//  const fs = require("fs"); 
    
//     if (fs.existsSync(LOGO_PATH)) doc.image(LOGO_PATH, 50, 30, { width: 80 });
//     doc.fontSize(20).text("OFFICIAL ACADEMIC TRANSCRIPT", { align: "center" });
//     doc.moveDown(2);

//   // Student Info
//   doc.fontSize(12).font("Helvetica");
//   doc.text(`Name: ${student.name.toUpperCase()}`);
//   doc.text(`Registration Number: ${student.regNo}`);
//   doc.text(`Program: ${(student.program as any)?.name || "N/A"}`);
//   doc.moveDown(2);

//   // Grades by Year → Semester
//   let currentYear = "";
//   let currentSemester = "";

//   grades.forEach((g) => {
//   const year = g.academicYear.year;
//     const sem = g.semester;

//     if (year !== currentYear) {
//       if (currentYear) doc.addPage();
//       currentYear = year;
//       currentSemester = "";
//       doc.fontSize(14).font("Helvetica-Bold").text(`Academic Year: ${year}`, { underline: true });
//       doc.moveDown(0.8);
//         // doc.fontSize(10).text("Code      Unit Name".padEnd(50) + "Grade   Points   Status");
//       // doc.moveDown(0.3);
//     }

//      if (sem !== currentSemester) {
//       currentSemester = sem;
//       doc.fontSize(12).font("Helvetica-Bold")
//         .text(`${sem}`, { indent: 20 })
//         .moveDown(0.5);
//     }

//     doc.fontSize(11).text(
//       `${g.unit.code.padEnd(12)} - ${g.unit.name.padEnd(50)} Grade: ${g.grade.padEnd(8)}   Status: ${g.status}`,
//       { indent: 40 }
//     );
//   });

//   // Footer
//   doc.moveDown(4);
//  doc.fontSize(10).text("This is a system-generated transcript.", { align: "center" });
//   doc.text("Verified electronically.", { align: "center" });

//   if (fs.existsSync(SIGNATURE_PATH)) {
//       doc.image(SIGNATURE_PATH, 400, doc.y + 20, { width: 120 });
//     }
//   doc.text("Registrar (Academics)", 400, doc.y + 80, { align: "left" });

//   doc.end();
// }

// export async function generatePassList(academicYearId: string, institutionId: string, res: Response): Promise<void> {
//   const doc = new PDFDocument({ margin: 50, size: "A4" });
//   res.setHeader("Content-Type", "application/pdf");
//   // res.setHeader("Content-Disposition", `attachment; filename="Pass_List_${Date.now()}.pdf"`);
//   doc.pipe(res);

//   const [institution, academicYear, grades] = await Promise.all([
//     Institution.findById(institutionId),
//     AcademicYear.findById(academicYearId),
//     FinalGrade.find({ academicYear: academicYearId, institution: institutionId, status: "PASS" })
//       .populate<{ student: { regNo: string; fullName: string } }>("student")
//       .populate<{ unit: { code: string; name: string } }>("unit")
//       .sort({ "student.regNo": 1 })
//       .lean(),
//   ]);

//   doc.fontSize(16).text(institution?.name || "UNIVERSITY", { align: "center" });
//   doc.fontSize(14).text("PASS LIST", { align: "center" });
//   doc.fontSize(12).text(`Academic Year: ${academicYear?.year}`, { align: "center" });
//   doc.moveDown(2);

//   const headers = ["S/No", "Reg No.", "Name", "Unit Code", "Unit Name", "Grade"];
//   const colWidths = [40, 90, 140, 80, 140, 60];
//   let y = doc.y;

//   doc.font("Helvetica-Bold");
//   headers.forEach((h, i) => {
//     const x = 50 + colWidths.slice(0, i).reduce((a, b) => a + b, 0);
//     doc.text(h, x, y, { width: colWidths[i] });
//   });
//   y += 20;
//   doc.moveTo(50, y).lineTo(550, y).stroke();

//   doc.font("Helvetica").fontSize(10);
//   grades.forEach((g, i) => {
//     y += 20;
//     if (y > 750) { doc.addPage(); y = 100; }
//     const row = [
//       (i + 1).toString(),
//       g.student.regNo,
//       g.student.fullName,
//       g.unit.code,
//       g.unit.name,
//       g.grade,
//     ];
//     row.forEach((cell, j) => {
//       const x = 50 + colWidths.slice(0, j).reduce((a, b) => a + b, 0);
//       doc.text(cell, x, y, { width: colWidths[j] });
//     });
//   });

//   doc.end();
// }

// export async function generateConsolidatedMarksheet(academicYearId: string, institutionId: string, res: Response): Promise<void> {
//   const doc = new PDFDocument({ margin: 40, size: "A4", layout: "landscape" });
//   res.setHeader("Content-Type", "application/pdf");
//   res.setHeader("Content-Disposition", `attachment; filename="Consolidated_Marksheet.pdf"`);
 
//   doc.pipe(res);

//   const grades = await FinalGrade.find({ academicYear: academicYearId, institution: institutionId })
//     .populate<{ student: { regNo: string; fullName: string } }>("student")
//     .populate<{ unit: { code: string } }>("unit")
//     .sort({ "student.regNo": 1, "unit.code": 1 })
//     .lean();

//   const students = [...new Map(grades.map(g => [g.student.regNo, g.student])).values()];
//   const units = [...new Map(grades.map(g => [g.unit.code, g.unit])).values()];

//   doc.fontSize(18).text("CONSOLIDATED MARKSHEET", { align: "center" });
//   doc.moveDown();

//   let y = 120;
//   const startX = 40;

//   doc.font("Helvetica-Bold").fontSize(9);
//   doc.text("Reg No.", startX, y);
//   doc.text("Name", startX + 80, y);
//   units.forEach((u: any, i) => doc.text(u.code, startX + 180 + i * 50, y, { width: 45, align: "center" }));
//   doc.text("Remark", startX + 180 + units.length * 50 + 50, y);

//   y += 20;
//   doc.moveTo(startX, y).lineTo(800, y).stroke();

//   doc.font("Helvetica").fontSize(8);
//   for (const student of students) {
//     y += 20;
//     if (y > 500) { doc.addPage(); y = 100; }

//     doc.text(student.regNo, startX, y);
//     doc.text(student.fullName.split(" ").slice(0, 3).join(" "), startX + 80, y, { width: 90 });

//     let passed = 0;
//     units.forEach((unit: any) => {
//       const grade = grades.find(g => g.student.regNo === student.regNo && g.unit.code === unit.code);
//       const mark = grade ? grade.grade : "-";
//       doc.text(mark, startX + 180 + units.indexOf(unit) * 50, y, { width: 45, align: "center" });
//       if (grade?.status === "PASS") passed++;
//     });

//     const remark = passed === units.length ? "PASS" : passed > 0 ? "SUPP" : "RETAKE";
//     doc.text(remark, startX + 180 + units.length * 50 + 50, y);
//   }

//   doc.end();
// }