// src/services/gradeCalculator.ts
import Mark from "../models/Mark";
import FinalGrade from "../models/FinalGrade";
import InstitutionSettings from "../models/InstitutionSettings";
import { calculateFinalResult } from "../utils/gradingCore";
import mongoose from "mongoose"; // Import mongoose
import { AuthenticatedRequest } from "../middleware/auth";

interface ComputeOptions { markId: mongoose.Types.ObjectId; coordinatorReq?: AuthenticatedRequest; session?: mongoose.ClientSession; }

export async function computeFinalGrade({ markId, session }: ComputeOptions) {
  const mark = await Mark.findById(markId)
    .populate(["student", "academicYear", "programUnit"])
    .session(session || null);
  if (!mark) throw new Error("Mark not found");


  const settings = await InstitutionSettings.findOne({ institution: mark.institution }).session(session || null);
  if (!settings) throw new Error("Settings not found");

  // Cast to any to bypass strict check on custom schema fields like unitType
  const markData = mark as any;

  const result = calculateFinalResult({
    cat1: markData.cat1Raw || 0, cat2: markData.cat2Raw || 0, cat3: markData.cat3Raw || 0,
    ass1: markData.assgnt1Raw || 0, ass2: markData.assgnt2Raw || 0, ass3: markData.assgnt3Raw || 0,
    practical: markData.practicalRaw || 0, examQ1: markData.examQ1Raw || 0, 
    examQ2: markData.examQ2Raw || 0, examQ3: markData.examQ3Raw || 0, examQ4: markData.examQ4Raw || 0, examQ5: markData.examQ5Raw || 0,
    unitType: (markData.unitType as "theory" | "lab" | "workshop") || "theory",
    examMode: (markData.examMode as "standard" | "mandatory_q1") || "standard",
    attempt: markData.attempt || "1st",
    settings: { catMax: settings.cat1Max || 30, assMax: settings.assignmentMax || 10, practicalMax: settings.practicalMax || 10, passMark: settings.passMark || 40 },
  });

  const grade = result.finalMark >= 70 ? "A" : result.finalMark >= 60 ? "B" : result.finalMark >= 50 ? "C" : result.finalMark >= 40 ? "D" : "E";
  const status = result.finalMark >= settings.passMark ? "PASS" : "SUPPLEMENTARY";
  

  // --- FIX 2: Atomic Update for Mark Document ---
  await Mark.updateOne(
    { _id: markId },
    { $set: { caTotal30: result.caTotal, examTotal70: result.examTotal, agreedMark: result.finalMark },},
  ).session(session || null);


  // --- FIX 3: Atomic Update for FinalGrade ---
  await FinalGrade.findOneAndUpdate(
    { student: mark.student, programUnit: mark.programUnit, academicYear: mark.academicYear },
    {
      $set: {
        totalMark: result.finalMark, grade, caTotal30: result.caTotal,
        examTotal70: result.examTotal, status: status,
        institution: mark.institution, semester: markData.semester || "SEMESTER 1",
      },
    },
    { upsert: true, session }, // Ensure session is passed here
  );
 
  return { ...result, grade, status };
}


