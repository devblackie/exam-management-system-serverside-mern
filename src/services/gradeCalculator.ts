// src/services/gradeCalculator.ts
import mongoose from "mongoose";
import Mark, { IMark } from "../models/Mark";
import FinalGrade from "../models/FinalGrade";
import InstitutionSettings from "../models/InstitutionSettings";
import { AuthenticatedRequest } from "../middleware/auth";

interface ComputeOptions {
  markId: mongoose.Types.ObjectId;
  coordinatorReq?: AuthenticatedRequest;
  session?: mongoose.ClientSession;
}

export async function computeFinalGrade({
  markId,
  coordinatorReq,
  session,
}: ComputeOptions) {
  const mark = (await Mark.findById(markId)
    .populate(["student", "academicYear", "programUnit"])
    .session(session || null)) as (IMark & mongoose.Document) | null;

  if (!mark) throw new Error("Mark not found");

  const settings = await InstitutionSettings.findOne({
    institution: mark.institution,
  }).session(session || null);

  const passMark = settings?.passMark || 40;
  const studentName = (mark.student as any)?.name || "Unknown";

  // 1. Calculate Final Mark
  const caScore = Number(mark.caTotal30) || 0;
  const examScore = Number(mark.examTotal70) || 0;
  const hasCA = caScore > 0;
  const hasExam = examScore > 0;
  let finalMark = Number((caScore + examScore).toFixed(0));


  console.log(`[GradeCalc] Processing ${studentName} (${mark.attempt}): CA=${caScore}, Exam=${examScore}, RawTotal=${finalMark}`);

  const isSpecial = mark.isSpecial || mark.attempt === "special";
  const isSupp = mark.attempt === "supplementary";
  // 2. APPLY SUPPLEMENTARY POLICY
  let isCapped = false;
  const isSuppAttempt = mark.attempt === "supplementary" || mark.isSupplementary;

  // if (isSuppAttempt) {
  //   if (finalMark >= passMark) {
  //     console.log(`[GradeCalc] Supp Passed. Capping ${finalMark} to ${passMark}`);
  //     finalMark = passMark;
  //     isCapped = true;
  //   } else {
  //     console.log(`[GradeCalc] Supp Failed. Final mark ${finalMark} remains below pass.`);
  //   }
  // }

  // 1. SPECIAL CASE HANDLING (Jakes/Tony)
  if (isSpecial) {
    if (!hasCA && hasExam) {
      // JAKES: No CATs, but sits Special Exam. 
      // Scale Exam to 100% (Option A: Pro-rated)
      finalMark = Number(((examScore / 70) * 100).toFixed(0));
    }
    // Tony Case: Has CA, sits Special later. 
    // Logic: Just use the finalMark (CA + Exam) without capping.
  } 
  // 2. SUPPLEMENTARY CAPPING
  else if (isSupp && finalMark >= passMark) {
    finalMark = passMark;
    isCapped = true;
  }

  // 3. Determine Grade
  let grade = "E";
  if (isSpecial && finalMark === 0) grade = "I"
  if (finalMark >= 70 && !isCapped) grade = "A";
  else if (finalMark >= 60 && !isCapped) grade = "B";
  else if (finalMark >= 50 && !isCapped) grade = "C";
  else if (finalMark >= passMark) grade = "D";
  else grade = "E";

  // 4. Determine Status & Progression
  let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" | "SPECIAL" = "PASS";
 
  if (isSpecial && !hasExam) {
    status = "SPECIAL"; // Tony is waiting for Special Exam
  } else if (!hasCA && !hasExam) {
    status = "RETAKE"; // Viotry never showed up
  } else if (!hasCA || !hasExam) {
    status = "INCOMPLETE"; // Jakes missing CA
  } else if (finalMark < 40) {
    status = isSupp ? "RETAKE" : "SUPPLEMENTARY";
  } else {
    status = "PASS";
  }

  // 5. Save FinalGrade
  await FinalGrade.findOneAndUpdate(
{ student: mark.student, programUnit: mark.programUnit, academicYear: mark.academicYear },
    {
      academicYear: mark.academicYear, 
      institution: mark.institution,
      semester: (mark.programUnit as any)?.requiredSemester || "SEMESTER 1",
      
      totalMark: finalMark,
      grade,
      status,
      attemptType: isSpecial ? "SPECIAL" : (isSupp ? "SUPPLEMENTARY" : "1ST_ATTEMPT"),
      cappedBecauseSupplementary: isCapped,
      remarks: mark.remarks
    },
    { upsert: true, new: true, session: session || null }
  );

  return { finalMark, grade, status, caScore, examScore };
}