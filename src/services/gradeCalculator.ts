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

export async function computeFinalGrade({ markId, session }: ComputeOptions) {
  const mark = await Mark.findById(markId)
    .populate(["student", "academicYear", "programUnit"])
    .session(session || null);
  if (!mark) throw new Error("Mark not found");

  const settings = await InstitutionSettings.findOne({
    institution: mark.institution,
  }).session(session || null);
  if (!settings) throw new Error("Institution configuration missing");

  // --- 1. CA CALCULATION (30%) ---
  const rawCats = [mark.cat1Raw, mark.cat2Raw, mark.cat3Raw].filter(
    (v) => v != null,
  );
  const catAvg =
    rawCats.length > 0
      ? rawCats.reduce((a, b) => a + b, 0) / rawCats.length
      : 0;

  const pWeight = settings.practicalMax > 0 ? 5 : 0;
  const aWeight = settings.assignmentMax > 0 ? 5 : 0;
  const cWeight = 30 - (pWeight + aWeight);

  const caFromCats = (catAvg / (settings.cat1Max || 20)) * cWeight;
  const caFromAss =
    aWeight > 0
      ? ((mark.assgnt1Raw || 0) / settings.assignmentMax) * aWeight
      : 0;
  const caFromPrac =
    pWeight > 0
      ? ((mark.practicalRaw || 0) / settings.practicalMax) * pWeight
      : 0;

  const caTotal30 = Number((caFromCats + caFromAss + caFromPrac).toFixed(2));

  // --- 2. EXAM CALCULATION (70%) ---
  // Apply "Best of" logic matching the Excel template
  const q1 = mark.examQ1Raw || 0;
  const others = [
    mark.examQ2Raw || 0,
    mark.examQ3Raw || 0,
    mark.examQ4Raw || 0,
    mark.examQ5Raw || 0,
  ].sort((a, b) => b - a); // Sort descending to get "Best"

  // Check if mandatory_q1 mode (you may need to pass this from the mark or programUnit)
  const isMandatoryQ1 = mark.examMode === "mandatory_q1";
  const takeCount = isMandatoryQ1 ? 2 : 3;

  const bestOthersSum = others.slice(0, takeCount).reduce((a, b) => a + b, 0);
  const examTotal70 = q1 + bestOthersSum;

  // Total out of 100
  // let finalMark = Math.round(caTotal30 + examTotal70);

  const totalRawMark = Math.round(caTotal30 + examTotal70);
  let finalMark = totalRawMark;
  let isCapped = false;

  // --- 3. CAPPING & GRADING ---  
  if (mark.isSpecial || mark.attempt === "special") {
    isCapped = false; 
    finalMark = totalRawMark; // Restore raw marks in case it was previously capped
  } else if (mark.attempt === "supplementary" && finalMark >= settings.passMark) {
    finalMark = settings.passMark;
    isCapped = true;
  }

  // --- ATTEMPT TYPE MAPPING for FinalGrade model ---
  let attemptType: "1ST_ATTEMPT" | "SPECIAL" | "SUPPLEMENTARY" | "RETAKE" | "RE_RETAKE" = "1ST_ATTEMPT";
  let attemptNumber = 1;

  if (mark.attempt === "special" || mark.isSpecial) {
    attemptType = "SPECIAL";
    attemptNumber = 1;
  } else if (mark.attempt === "supplementary") {
    attemptType = "SUPPLEMENTARY";
    attemptNumber = 1; // Or 2 depending on if you count Supp as a separate attempt number
  } else if (mark.attempt === "re-take") {
    attemptType = "RETAKE";
    attemptNumber = 2;
  } else if (mark.attempt === "re-retake") {
    attemptType = "RE_RETAKE";
    attemptNumber = 3;
  }

  // Determine Grade
  const sortedScale = [...(settings.gradingScale || [])].sort(
    (a, b) => b.min - a.min,
  );
  const matchedGrade = sortedScale.find((s) => finalMark >= s.min);
  const grade = matchedGrade ? matchedGrade.grade : "E";

  // Determine Status
  let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" | "SPECIAL" = "PASS";
  if (finalMark < settings.passMark) {
    // 1. If it's a Special Exam and the exam hasn't been sat (Exam = 0) or is pending
    if (mark.attempt === "special" || mark.isSpecial) {
      status = "SPECIAL";
    }
    // 2. Regular failure logic
    else {
      status = mark.attempt === "1st" ? "SUPPLEMENTARY" : "RETAKE";
    }
  }

  // --- 5. PERSIST ---
  await FinalGrade.findOneAndUpdate(
    { student: mark.student, programUnit: mark.programUnit, academicYear: mark.academicYear },
    {
      totalMark: finalMark, grade, status,
      attemptType, attemptNumber,
      cappedBecauseSupplementary: isCapped,
      caTotal30, examTotal70
    },
    { upsert: true, session },
  );

  // Also update the Mark document to keep totals in sync
  mark.caTotal30 = caTotal30;
  mark.examTotal70 = examTotal70;
  mark.agreedMark = finalMark;
  await mark.save({ session });

  return { finalMark, grade, status, caTotal30, examTotal70, isCapped };
}