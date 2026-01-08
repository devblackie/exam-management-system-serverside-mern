// // src/routes/programReports.ts

// import { requireAuth, requireRole } from "../middleware/auth";
// import Student from "../models/Student";
// import FinalGrade from "../models/FinalGrade";
// import AcademicYear from "../models/AcademicYear";
// import Program from "../models/Program";
// import asyncHandler from "express-async-handler";
// import { Response } from "express";

// const router = require("express").Router();

// // 1. PASS LIST — Students who passed ALL units in the academic year
// router.get("/passlist", requireAuth, requireRole("coordinator"), asyncHandler(async (req, res) => {
//   const { programId, academicYear: yearStr } = req.query;
//   if (!programId || !yearStr) return res.status(400).json({ error: "programId and academicYear required" });

//   const academicYear = await AcademicYear.findOne({ year: yearStr });
//   if (!academicYear) return res.status(404).json({ error: "Academic year not found" });

//   const students = await Student.find({ program: programId })
//     .select("regNo name")
//     .sort({ regNo: 1 });

//   const passList: any[] = [];

//   for (const student of students) {
//     const totalUnits = await FinalGrade.countDocuments({
//       student: student._id,
//       academicYear: academicYear._id,
//     });

//     const passedUnits = await FinalGrade.countDocuments({
//       student: student._id,
//       academicYear: academicYear._id,
//       status: "PASS",
//     });

//     const hasIncomplete = await FinalGrade.exists({
//       student: student._id,
//       academicYear: academicYear._id,
//       status: "INCOMPLETE",
//     });

//     if (totalUnits > 0 && passedUnits === totalUnits && !hasIncomplete) {
//       passList.push({
//         regNo: student.regNo,
//         name: student.name.toUpperCase(),
//         status: "PASS - PROCEED TO NEXT YEAR",
//       });
//     }
//   }

//   res.json({
//     academicYear: academicYear.year,
//     program: (await Program.findById(programId))?.name,
//     generatedAt: new Date().toLocaleString("en-KE"),
//     totalPassed: passList.length,
//     passList,
//   });
// }));

// // 2. CONSOLIDATED MARKSHEET — Full breakdown
// router.get("/consolidated", requireAuth, requireRole("coordinator"), asyncHandler(async (req, res) => {
//   const { programId, academicYear: yearStr } = req.query;
//   if (!programId || !yearStr) return res.status(400).json({ error: "programId and academicYear required" });

//   const academicYear = await AcademicYear.findOne({ year: yearStr });
//   if (!academicYear) return res.status(404).json({ error: "Academic year not found" });

//   const program = await Program.findById(programId);
//   if (!program) return res.status(404).json({ error: "Program not found" });

//   const students = await Student.find({ program: programId })
//     .select("regNo name")
//     .sort({ regNo: 1 });

//   const report: any[] = [];

//   for (const student of students) {
//     const grades = await FinalGrade.find({
//       student: student._id,
//       academicYear: academicYear._id,
//     })
//       .populate("unit", "code")
//       .sort({ "unit.code": 1 });

//     const summary = {
//       regNo: student.regNo,
//       name: student.name.toUpperCase(),
//       totalUnits: grades.length,
//       passed: grades.filter(g => g.status === "PASS").length,
//       supplementary: grades.filter(g => g.status === "SUPPLEMENTARY").length,
//       retake: grades.filter(g => g.status === "RETAKE").length,
//       incomplete: grades.filter(g => g.status === "INCOMPLETE").length,
//       finalStatus: "PENDING",
//       units: grades.map(g => ({
//         code: g.unit.code,
//         grade: g.grade,
//         status: g.status,
//       })),
//     };

//     if (summary.incomplete > 0) {
//       summary.finalStatus = "INCOMPLETE";
//     } else if (summary.retake > 0) {
//       summary.finalStatus = "RETAKE YEAR";
//     } else if (summary.supplementary > 0) {
//       summary.finalStatus = "SUPPLEMENTARY";
//     } else if (summary.passed === summary.totalUnits) {
//       summary.finalStatus = "PASS - PROCEED";
//     } else {
//       summary.finalStatus = "REPEAT FAILED UNITS";
//     }

//     report.push(summary);
//   }

//   res.json({
//     program: program.name,
//     academicYear: academicYear.year,
//     generatedAt: new Date().toLocaleString("en-KE"),
//     totalStudents: students.length,
//     report,
//   });
// }));