// src/services/gradeCalculator.ts
import Mark, { IMark } from "../models/Mark";
import MarkDirect, { IMarkDirect } from "../models/MarkDirect";
import FinalGrade from "../models/FinalGrade";
import InstitutionSettings from "../models/InstitutionSettings";
import { calculateFinalResult } from "../utils/gradingCore";
import mongoose from "mongoose";
import { AuthenticatedRequest } from "../middleware/auth";

interface ComputeOptions { 
  markId: mongoose.Types.ObjectId; 
  coordinatorReq?: AuthenticatedRequest; 
  session?: mongoose.ClientSession; 
}

export async function computeFinalGrade({ markId, session }: ComputeOptions) {
  let markDoc: any = await Mark.findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);

  let isDirect = false;
  if (!markDoc) {
    markDoc = await MarkDirect.findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);
    isDirect = true;
  }

  if (!markDoc) throw new Error("Mark not found");

  const settings = await InstitutionSettings.findOne({ institution: markDoc.institution }).session(session || null);
  if (!settings) throw new Error("Settings not found");

  // 1. Run the raw calculation
  const result = calculateFinalResult({
    cat1: markDoc.cat1Raw || 0, cat2: markDoc.cat2Raw || 0, cat3: markDoc.cat3Raw || 0,
    ass1: markDoc.assgnt1Raw || 0, ass2: markDoc.assgnt2Raw || 0, ass3: markDoc.assgnt3Raw || 0,
    practical: markDoc.practicalRaw || 0, 
    examQ1: markDoc.examQ1Raw || 0, examQ2: markDoc.examQ2Raw || 0, 
    examQ3: markDoc.examQ3Raw || 0, examQ4: markDoc.examQ4Raw || 0, examQ5: markDoc.examQ5Raw || 0,
    unitType: (markDoc.unitType as any) || "theory", examMode: (markDoc.examMode as any) || "standard", attempt: markDoc.attempt || "1st",
    settings: { catMax: settings.cat1Max || 30, assMax: settings.assignmentMax || 10, practicalMax: settings.practicalMax || 10, passMark: settings.passMark || 40 },
  });

  // 2. PRESERVATION STRATEGY
  let finalCA = (result.caTotal === 0 && markDoc.caTotal30 > 0) ? markDoc.caTotal30 : result.caTotal;
  let finalExam = (result.examTotal === 0 && markDoc.examTotal70 > 0) ? markDoc.examTotal70 : result.examTotal;
  let finalAgreed = finalCA + finalExam;

  // 3. SPECIAL EXAM OVERRIDE
  // Check if current state is Special (either just granted or already special)
  const isSpecial = markDoc.isSpecial || markDoc.attempt === "special";
  
  let grade: string;
  let status: string;

  if (isSpecial) {
    grade = "I"; // Incomplete
    status = "SPECIAL";
    finalExam = 0;
    finalAgreed = finalCA; 
  } else {
    grade = finalAgreed >= 70 ? "A" : finalAgreed >= 60 ? "B" : finalAgreed >= 50 ? "C" : finalAgreed >= 40 ? "D" : "E";
    status = finalAgreed >= settings.passMark ? "PASS" : "SUPPLEMENTARY";
  }

  // 4. Update the source document (Mark or MarkDirect)
  const updatePayload = { $set: { caTotal30: finalCA, examTotal70: finalExam, agreedMark: finalAgreed }};

  if (isDirect) await MarkDirect.updateOne({ _id: markId }, updatePayload).session(session || null);
  else await Mark.updateOne({ _id: markId }, updatePayload).session(session || null);
  

  // 5. Update the FinalGrade record for the transcripts
  await FinalGrade.findOneAndUpdate(
    { 
        student: markDoc.student._id || markDoc.student, 
        programUnit: markDoc.programUnit._id || markDoc.programUnit, 
        academicYear: markDoc.academicYear._id || markDoc.academicYear 
    },
    {
      $set: {
        totalMark: finalAgreed, 
        grade, 
        caTotal30: finalCA,
        examTotal70: finalExam, 
        status: status,
        institution: markDoc.institution, 
        semester: markDoc.semester || "SEMESTER 1",
      },
    },
    { upsert: true, session },
  );
 
  return { caTotal: finalCA, examTotal: finalExam, finalMark: finalAgreed, grade, status };
}






