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
  let finalMark = Number((caScore + examScore).toFixed(0));

  console.log(`[GradeCalc] Processing ${studentName} (${mark.attempt}): CA=${caScore}, Exam=${examScore}, RawTotal=${finalMark}`);

  // 2. APPLY SUPPLEMENTARY POLICY
  let isCapped = false;
  const isSuppAttempt = mark.attempt === "supplementary" || mark.isSupplementary;

  if (isSuppAttempt) {
    if (finalMark >= passMark) {
      console.log(`[GradeCalc] Supp Passed. Capping ${finalMark} to ${passMark}`);
      finalMark = passMark;
      isCapped = true;
    } else {
      console.log(`[GradeCalc] Supp Failed. Final mark ${finalMark} remains below pass.`);
    }
  }

  // 3. Determine Grade
  let grade = "E";
  if (finalMark >= 70 && !isCapped) grade = "A";
  else if (finalMark >= 60 && !isCapped) grade = "B";
  else if (finalMark >= 50 && !isCapped) grade = "C";
  else if (finalMark >= passMark) grade = "D";
  else grade = "E";

  // 4. Determine Status & Progression
  let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" = "INCOMPLETE";

  if (finalMark >= passMark) {
    status = "PASS";
  } else {
    status = isSuppAttempt ? "RETAKE" : "SUPPLEMENTARY";
    console.log(`[GradeCalc] Status set to ${status} for ${studentName}`);
  }

  // 5. Save FinalGrade
  await FinalGrade.findOneAndUpdate(
    { student: mark.student, programUnit: mark.programUnit },
    {
      institution: mark.institution,
      academicYear: mark.academicYear,
      semester: (mark.programUnit as any).requiredSemester === 1 ? "SEMESTER 1" : "SEMESTER 2",
      totalMark: finalMark,
      grade,
      status,
      attemptType: isSuppAttempt ? "SUPPLEMENTARY" : "1ST_ATTEMPT",
      attemptNumber: isSuppAttempt ? 2 : 1,
      cappedBecauseSupplementary: isCapped,
    },
    { upsert: true, new: true, session: session || null }
  );

  return { finalMark, grade, status, caScore, examScore };
}