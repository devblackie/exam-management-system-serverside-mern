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

  // 1. Final Mark (Internal Mark /100) = CA Grand Total /30 + Total Exam /70

  // Use the totals imported directly from the Scoresheet:
  const caScore = Number(mark.caTotal30) || 0;
  const examScore = Number(mark.examTotal70) || 0;
  // The final mark is simply the sum of the two approved totals.
  let finalMark = Number((caScore + examScore).toFixed(2));

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
      .find((s) => finalMark >= s.min);
    if (entry) {
      grade = entry.grade;
      points = entry.points || 0;
    }
  } else {
    // Default scale
    if (finalMark >= 69.5) {
      grade = "A";
      points = 4.0;
    } else if (finalMark >= 59.5) {
      grade = "B";
      points = 3.0;
    } else if (finalMark >= 49.5) {
      grade = "C";
      points = 2.0;
    } else if (finalMark >= 39.5) {
      grade = "D";
      points = 1.0;
    } else {
      grade = "E";
      points = 0;
    }
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

      if (suppCount + 1 >= settings.retakeThreshold) {
        // +1 for the current unit
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
      cappedBecauseSupplementary:
        mark.isSupplementary && finalMark === settings.supplementaryThreshold,
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
