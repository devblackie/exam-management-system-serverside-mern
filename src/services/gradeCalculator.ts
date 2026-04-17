// serverside/src/services/gradeCalculator.ts
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import FinalGrade from "../models/FinalGrade";
import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";
import { calculateFinalResult } from "../utils/gradingCore";
import mongoose from "mongoose";
import type { AuthenticatedRequest } from "../middleware/auth";

interface ComputeOptions {
  markId:          mongoose.Types.ObjectId;
  coordinatorReq?: AuthenticatedRequest;
  session?:        mongoose.ClientSession;
}

export async function computeFinalGrade({ markId, session }: ComputeOptions) {
  // ── 1. Resolve mark document ──────────────────────────────────────────────
  let markDoc: any = await Mark
    .findById(markId).populate(["student","academicYear","programUnit"]).session(session || null);

  let isDirect = false;
  if (!markDoc) {
    markDoc   = await MarkDirect
      .findById(markId).populate(["student","academicYear","programUnit"]).session(session || null);
    isDirect  = true;
  }
  if (!markDoc) throw new Error("Mark not found");

  const settings = await InstitutionSettings.findOne({ institution: markDoc.institution }).session(session || null);
  if (!settings) throw new Error("Institution settings not found");

  // ── 2. Raw calculation ────────────────────────────────────────────────────
  const result = calculateFinalResult({
    cat1: markDoc.cat1Raw || 0, cat2: markDoc.cat2Raw || 0, cat3: markDoc.cat3Raw || 0,
    ass1: markDoc.assgnt1Raw || 0, ass2: markDoc.assgnt2Raw || 0, ass3: markDoc.assgnt3Raw || 0,
    practical: markDoc.practicalRaw || 0,
    examQ1: markDoc.examQ1Raw || 0, examQ2: markDoc.examQ2Raw || 0,
    examQ3: markDoc.examQ3Raw || 0, examQ4: markDoc.examQ4Raw || 0, examQ5: markDoc.examQ5Raw || 0,
    unitType: (markDoc.unitType as any) || "theory",
    examMode: (markDoc.examMode as any) || "standard",
    attempt:  markDoc.attempt || "1st",
    settings: { catMax: settings.cat1Max || 30, assMax: settings.assignmentMax || 10, practicalMax: settings.practicalMax || 10, passMark: settings.passMark || 40 },
  });

  // ── 3. Preservation strategy ──────────────────────────────────────────────
  const finalCA    = result.caTotal   === 0 && markDoc.caTotal30   > 0 ? markDoc.caTotal30   : result.caTotal;
  const finalExam  = result.examTotal === 0 && markDoc.examTotal70 > 0 ? markDoc.examTotal70 : result.examTotal;
  const finalAgreed = finalCA + finalExam;

  // ── 4. Special exam override (ENG.18) ─────────────────────────────────────
  const isSpecial = markDoc.isSpecial === true || markDoc.attempt === "special";
  let grade:  string;
  let status: string;

  if (isSpecial) {
    grade = "I"; status = "SPECIAL";
  } else {
    const sortedScale = [...((settings as any).gradingScale || [])].sort((a: any, b: any) => b.min - a.min);
    grade  = sortedScale.find((s: any) => finalAgreed >= s.min)?.grade ?? "E";
    status = finalAgreed >= ((settings as any).passMark || 40) ? "PASS" : "SUPPLEMENTARY";
  }

  // ── 5. Update source mark ─────────────────────────────────────────────────
  const markUpdate = { $set: { caTotal30: finalCA, examTotal70: isSpecial ? 0 : finalExam, agreedMark: isSpecial ? finalCA : finalAgreed } };
  if (isDirect) await MarkDirect.updateOne({ _id: markId }, markUpdate).session(session || null);
  else          await Mark.updateOne({ _id: markId }, markUpdate).session(session || null);

  // ── 6. Upsert FinalGrade ──────────────────────────────────────────────────
  await FinalGrade.findOneAndUpdate(
    {
      student:      markDoc.student._id     || markDoc.student,
      programUnit:  markDoc.programUnit._id  || markDoc.programUnit,
      academicYear: markDoc.academicYear._id || markDoc.academicYear,
    },
    {
      $set: {
        totalMark:    isSpecial ? finalCA : finalAgreed,
        grade, caTotal30: finalCA, examTotal70: isSpecial ? 0 : finalExam,
        status, isSpecial,
        attemptType:  markDoc.attempt === "supplementary" ? "SUPPLEMENTARY" : markDoc.attempt === "re-take" ? "RETAKE" : "1ST_ATTEMPT",
        institution:  markDoc.institution,
        semester:     markDoc.semester || "SEMESTER 1",
      },
    },
    { upsert: true, session },
  );

  // ── 7. Carry-forward resolution (ENG.14) ──────────────────────────────────
  // Fire-and-forget: when a CF unit is passed, remove it from carryForwardUnits.
  if (status === "PASS") {
    _resolveCFUnit(
      (markDoc.student._id    || markDoc.student).toString(),
      (markDoc.programUnit._id || markDoc.programUnit).toString(),
    ).catch((err: Error) => console.error("[gradeCalculator] CF resolution:", err.message));
  }

  return { caTotal: finalCA, examTotal: isSpecial ? 0 : finalExam, finalMark: isSpecial ? finalCA : finalAgreed, grade, status };
}

// ─── Private: remove a CF unit once it is passed ─────────────────────────────

async function _resolveCFUnit(studentId: string, programUnitId: string): Promise<void> {
  const studentDoc = await Student.findById(studentId).select("carryForwardUnits qualifierSuffix").lean() as any;
  if (!studentDoc) return;

  const hasCF = ((studentDoc.carryForwardUnits || []) as any[]).some(
    (u: any) => u.programUnitId === programUnitId,
  );
  if (!hasCF) return;

  await Student.findByIdAndUpdate(studentId, { $pull: { carryForwardUnits: { programUnitId } } });

  const refreshed = await Student.findById(studentId).select("carryForwardUnits").lean() as any;
  if (((refreshed?.carryForwardUnits || []) as any[]).length === 0) {
    await Student.findByIdAndUpdate(studentId, { $set: { qualifierSuffix: "" } });
  }
}