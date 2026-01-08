// // src/services/gradeCalculator.ts
// import mongoose from "mongoose";
// import Mark from "../models/Mark";
// import FinalGrade from "../models/FinalGrade";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { AuthenticatedRequest } from "../middleware/auth";

// interface ComputeOptions {
//   markId: mongoose.Types.ObjectId;
//     coordinatorReq?: AuthenticatedRequest;
//   session?: mongoose.ClientSession;
// }

// export async function computeFinalGrade({ markId, coordinatorReq, session }: ComputeOptions) {
//   const mark = await Mark.findById(markId)
//     .populate(["student", "academicYear", "programUnit"])
//     .session(session || null);

//   if (!mark) throw new Error("Mark not found");

//   const settings = await InstitutionSettings.findOne({
//     institution: mark.institution,
//   }).session(session || null);

//   if (!settings) throw new Error("Institution grading settings not configured");

//   // ——————————————————————————————————
//   // 1. Detect which components exist
//   // ——————————————————————————————————
//   const hasCat3 = mark.cat3 != null && settings.cat3Max > 0;
//   const hasAssignment = mark.assignment != null && settings.assignmentMax > 0;
//   const hasPractical = mark.practical != null && settings.practicalMax > 0;

//   const assignmentWeight = hasAssignment ? 5 : 0;
//   const practicalWeight = hasPractical ? 5 : 0;
//   const remainingCaWeight = 30 - assignmentWeight - practicalWeight;

//   // ——————————————————————————————————
//   // 2. Active CATs
//   // ——————————————————————————————————
//   const catComponents = [
//     { value: mark.cat1, max: settings.cat1Max },
//     { value: mark.cat2, max: settings.cat2Max },
//     ...(hasCat3 ? [{ value: mark.cat3!, max: settings.cat3Max }] : []),
//   ].filter(c => c.value != null && c.max > 0);

//   const catWeightEach = catComponents.length > 0 ? remainingCaWeight / catComponents.length : 0;

//   // ——————————————————————————————————
//   // 3. Calculate CA Score (out of 30)
//   // ——————————————————————————————————
//   let caScore = 0;

//   catComponents.forEach(cat => {
//     const percentage = (cat.value! / cat.max) * 100;
//     caScore += (percentage / 100) * catWeightEach;
//   });

//   if (hasAssignment && mark.assignment != null) {
//     const pct = (mark.assignment / settings.assignmentMax) * 100;
//     caScore += (pct / 100) * 5;
//   }

//   if (hasPractical && mark.practical != null) {
//     const pct = (mark.practical / settings.practicalMax) * 100;
//     caScore += (pct / 100) * 5;
//   }

//   // ——————————————————————————————————
//   // 4. Exam Score (out of 70)
//   // ——————————————————————————————————
//   const examScore = mark.exam != null ? ((mark.exam / 70) * 70) : 0;

//   // ——————————————————————————————————
//   // 5. Final Mark
//   // ——————————————————————————————————
//   let finalMark = Number((caScore + examScore).toFixed(2));

//   // Supplementary capping
//   if (mark.isSupplementary && finalMark > settings.supplementaryThreshold) {
//     finalMark = settings.supplementaryThreshold;
//   }

//   // ——————————————————————————————————
//   // 6. Determine Grade
//   // ——————————————————————————————————
//   let grade = "F";
//   let points = 0;

//   if (settings.gradingScale && settings.gradingScale.length > 0) {
//     const entry = settings.gradingScale
//       .sort((a, b) => b.min - a.min)
//       .find(s => finalMark >= s.min);
//     if (entry) {
//       grade = entry.grade;
//       points = entry.points || 0;
//     }
//   } else {
//     // Default scale
//     if (finalMark >= 70) { grade = "A"; points = 4.0; }
//     else if (finalMark >= 60) { grade = "B"; points = 3.0; }
//     else if (finalMark >= 50) { grade = "C"; points = 2.0; }
//     else if (finalMark >= 40) { grade = "D"; points = 1.0; }
//     else { grade = "F"; points = 0; }
//   }

//   // ——————————————————————————————————
//   // 7. Determine Status
//   // ——————————————————————————————————
//   let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" = "INCOMPLETE";

//   if (mark.exam == null) {
//     status = "INCOMPLETE";
//   } else if (finalMark >= settings.passMark) {
//     status = "PASS";
//   } else {
//     status = "SUPPLEMENTARY";

//     // Check for RETAKE
//     if (!mark.isSupplementary) {
//       const suppCount = await FinalGrade.countDocuments({
//         student: mark.student,
//         academicYear: mark.academicYear,
//         status: "SUPPLEMENTARY",
//         _id: { $ne: mark._id },
//       }).session(session || null);

//       if (suppCount + 1 >= settings.retakeThreshold) {
//         status = "RETAKE";
//       }
//     }
//   }

//   // ——————————————————————————————————
//   // 8. Save FinalGrade
//   // ——————————————————————————————————
//   await FinalGrade.findOneAndUpdate(
//     {
//       student: mark.student,
//       programUnit: mark.programUnit,
//       academicYear: mark.academicYear,
//     },
//     {
//       totalMark: finalMark,
//       grade,
//       points,
//       status,
//       cappedBecauseSupplementary: mark.isSupplementary && finalMark === settings.supplementaryThreshold,
//       computedFrom: {
//         cat1: mark.cat1,
//         cat2: mark.cat2,
//         cat3: mark.cat3,
//         assignment: mark.assignment,
//         practical: mark.practical,
//         exam: mark.exam,
//       },
//     },
//     { upsert: true, new: true, session: session || null }
//   );

//   return { finalMark, grade, points, status, caScore, examScore };
// }


// src/services/gradeCalculator.ts
import mongoose from "mongoose";
import Mark, { IMark } from "../models/Mark"; // Import IMark for strong typing
import FinalGrade from "../models/FinalGrade";
import InstitutionSettings from "../models/InstitutionSettings";
import { AuthenticatedRequest } from "../middleware/auth";

interface ComputeOptions {
 markId: mongoose.Types.ObjectId;
 coordinatorReq?: AuthenticatedRequest;
session?: mongoose.ClientSession;
}

export async function computeFinalGrade({ markId, coordinatorReq, session }: ComputeOptions) {
 // Explicitly cast the mark result to IMark to resolve type errors
 const mark = await Mark.findById(markId)
 .populate(["student", "academicYear", "programUnit"])
 .session(session || null) as (IMark & mongoose.Document) | null;

if (!mark) throw new Error("Mark not found");

 const settings = await InstitutionSettings.findOne({
 institution: mark.institution,
 }).session(session || null);

if (!settings) throw new Error("Institution grading settings not configured");

 // 1. Final Mark (Internal Mark /100) = CA Grand Total /30 + Total Exam /70

// Use the totals imported directly from the KU Scoresheet:
const caScore = mark.caTotal30; // Already out of 30
const examScore = mark.examTotal70; // Already out of 70
 // The final mark is simply the sum of the two approved totals.
 let finalMark = Number((caScore + examScore).toFixed(2));

 // NOTE: We can also use mark.agreedMark or mark.internalExaminerMark for the final score,
// but calculating it here ensures consistency between the raw data and the final grade record.

 // Supplementary capping
 if (mark.isSupplementary && finalMark > settings.supplementaryThreshold) {
 finalMark = settings.supplementaryThreshold;
 }

 // 2. Determine Grade
 let grade = "E";
let points = 0;

 if (settings.gradingScale && settings.gradingScale.length > 0) {
 const entry = settings.gradingScale
 .sort((a, b) => b.min - a.min)
 .find(s => finalMark >= s.min);
 if (entry) {
 grade = entry.grade;
 points = entry.points || 0;
 }
 } else {
 // Default scale
 if (finalMark >= 69.5) { grade = "A"; points = 4.0; }
 else if (finalMark >= 59.5) { grade = "B"; points = 3.0; }
 else if (finalMark >= 49.5) { grade = "C"; points = 2.0; }
 else if (finalMark >= 39.5) { grade = "D"; points = 1.0; }
 else { grade = "E"; points = 0; }
 }

 // 3. Determine Status
 let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" = "INCOMPLETE";

// Check if required components are present (CA total 30 and Exam total 70)
 if (mark.caTotal30 === 0 && mark.examTotal70 === 0) {
     status = "INCOMPLETE";
 } else if (finalMark >= settings.passMark) {
 status = "PASS";
 } else {
 status = "SUPPLEMENTARY";

 // Check for RETAKE
 if (!mark.isSupplementary) {
 const suppCount = await FinalGrade.countDocuments({
 student: mark.student,
 programUnit: mark.programUnit, // Use programUnit for accurate unit count
 academicYear: mark.academicYear,
 status: "SUPPLEMENTARY",
 unit: { $ne: mark.programUnit }, // Exclude the current unit's supplementary status
 }).session(session || null);

 if (suppCount + 1 >= settings.retakeThreshold) { // +1 for the current unit
 status = "RETAKE";
 }
 }
 }

 // 4. Save FinalGrade
 await FinalGrade.findOneAndUpdate(
 {
student: mark.student,
 programUnit: mark.programUnit,
 academicYear: mark.academicYear,
 },
 {
 totalMark: finalMark,
 grade,
 points,
status,
 cappedBecauseSupplementary: mark.isSupplementary && finalMark === settings.supplementaryThreshold,
 // Store the aggregated scores that led to the final mark for audit
 computedFrom: {
 caTotal30: mark.caTotal30,
 examTotal70: mark.examTotal70,
 // The original granular marks are available on the Mark document itself if needed
 },
 },
 { upsert: true, new: true, session: session || null }
 );

 return { finalMark, grade, points, status, caScore, examScore };
}