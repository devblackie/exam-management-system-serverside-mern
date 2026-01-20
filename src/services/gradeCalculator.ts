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

export async function computeFinalGrade({
  markId,
  coordinatorReq,
  session,
}: ComputeOptions) {
  // Explicitly cast the mark result to IMark to resolve type errors
  const mark = (await Mark.findById(markId)
    .populate(["student", "academicYear", "programUnit"])
    .session(session || null)) as (IMark & mongoose.Document) | null;

  if (!mark) throw new Error("Mark not found");

  const settings = await InstitutionSettings.findOne({
    institution: mark.institution,
  }).session(session || null);

  if (!settings) throw new Error("Institution grading settings not configured");

  // 1. Final Mark (Internal Mark /100) = CA Grand Total /30 + Total Exam /70 Calculate Final Mark: CA (from 1st attempt) + New Exam Score

  // Use the totals imported directly from the Scoresheet:
  const caScore = Number(mark.caTotal30) || 0;
  const examScore = Number(mark.examTotal70) || 0;
  let finalMark = Number((caScore + examScore).toFixed(2));

 // 2. APPLY SUPPLEMENTARY POLICY
  // If this is a supplementary, the maximum grade allowed is 'D' (usually 40-49%)
  let isCapped = false;
  if (mark.attempt === "Supplementary" || mark.isSupplementary) {
     const gradeDMax = 49.4; // The upper limit for a 'D' grade
     if (finalMark >= settings.passMark) {
        // Student passed the Supp: Cap the mark at the minimum required for a 'D'
        // or keep it as is if it's already within the 'D' range.
        // Most institutions cap at exactly the pass mark (e.g., 40%)
        if (finalMark > settings.passMark) {
            finalMark = settings.passMark; 
            isCapped = true;
        }
     }
  }

// 3. Determine Grade based on the (potentially capped) mark
  let grade = "E";
  if (finalMark >= 69.5 && !isCapped) grade = "A";
  else if (finalMark >= 59.5 && !isCapped) grade = "B";
  else if (finalMark >= 49.5 && !isCapped) grade = "C";
  else if (finalMark >= settings.passMark) grade = "D";
  else grade = "E";

 // 4. Determine Status & Progression to Retake
  let status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE" = "INCOMPLETE";

  if (finalMark >= settings.passMark) {
    status = "PASS";
  } else {
    // If they were already sitting a Supplementary and failed again...
    if (mark.attempt === "Supplementary" || mark.isSupplementary) {
      status = "RETAKE"; // Second failure triggers a full Retake (Attempt 3)
    } else {
      status = "SUPPLEMENTARY"; // First failure triggers a Supplementary
    }
  }

  // 5. Save FinalGrade
  await FinalGrade.findOneAndUpdate(
    {
      student: mark.student,
      programUnit: mark.programUnit,
      // academicYear: mark.academicYear,
    },
    {
    totalMark: finalMark,
      grade,
      status,
      attemptType: mark.isSupplementary ? "SUPPLEMENTARY" : "1ST_ATTEMPT",
      attemptNumber: mark.isSupplementary ? 2 : 1,
      cappedBecauseSupplementary: isCapped
      
    },
    { upsert: true, new: true, session: session || null }
  );

  return { finalMark, grade,  status, caScore, examScore };
}
