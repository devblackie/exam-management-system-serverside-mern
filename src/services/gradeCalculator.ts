// src/services/gradeCalculator.ts
// import Mark, { IMark } from "../models/Mark";
// import MarkDirect, { IMarkDirect } from "../models/MarkDirect";
// import FinalGrade from "../models/FinalGrade";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { calculateFinalResult } from "../utils/gradingCore";
// import mongoose from "mongoose";
// import { AuthenticatedRequest } from "../middleware/auth";

// interface ComputeOptions { 
//   markId: mongoose.Types.ObjectId; 
//   coordinatorReq?: AuthenticatedRequest; 
//   session?: mongoose.ClientSession; 
// }

// export async function computeFinalGrade({ markId, session }: ComputeOptions) {
//   let markDoc: any = await Mark.findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);

//   let isDirect = false;
//   if (!markDoc) {
//     markDoc = await MarkDirect.findById(markId).populate(["student", "academicYear", "programUnit"]).session(session || null);
//     isDirect = true;
//   }

//   if (!markDoc) throw new Error("Mark not found");

//   const settings = await InstitutionSettings.findOne({ institution: markDoc.institution }).session(session || null);
//   if (!settings) throw new Error("Settings not found");

//   // 1. Run the raw calculation
//   const result = calculateFinalResult({
//     cat1: markDoc.cat1Raw || 0, cat2: markDoc.cat2Raw || 0, cat3: markDoc.cat3Raw || 0,
//     ass1: markDoc.assgnt1Raw || 0, ass2: markDoc.assgnt2Raw || 0, ass3: markDoc.assgnt3Raw || 0,
//     practical: markDoc.practicalRaw || 0, 
//     examQ1: markDoc.examQ1Raw || 0, examQ2: markDoc.examQ2Raw || 0, 
//     examQ3: markDoc.examQ3Raw || 0, examQ4: markDoc.examQ4Raw || 0, examQ5: markDoc.examQ5Raw || 0,
//     unitType: (markDoc.unitType as any) || "theory", examMode: (markDoc.examMode as any) || "standard", attempt: markDoc.attempt || "1st",
//     settings: { catMax: settings.cat1Max || 30, assMax: settings.assignmentMax || 10, practicalMax: settings.practicalMax || 10, passMark: settings.passMark || 40 },
//   });

//   // 2. PRESERVATION STRATEGY
//   let finalCA = (result.caTotal === 0 && markDoc.caTotal30 > 0) ? markDoc.caTotal30 : result.caTotal;
//   let finalExam = (result.examTotal === 0 && markDoc.examTotal70 > 0) ? markDoc.examTotal70 : result.examTotal;
//   let finalAgreed = finalCA + finalExam;

//   // 3. SPECIAL EXAM OVERRIDE
//   // Check if current state is Special (either just granted or already special)
//   const isSpecial = markDoc.isSpecial || markDoc.attempt === "special";
  
//   let grade: string;
//   let status: string;

//   if (isSpecial) {
//     grade = "I"; // Incomplete
//     status = "SPECIAL";
//     finalExam = 0;
//     finalAgreed = finalCA; 
//   } else {
//     grade = finalAgreed >= 70 ? "A" : finalAgreed >= 60 ? "B" : finalAgreed >= 50 ? "C" : finalAgreed >= 40 ? "D" : "E";
//     status = finalAgreed >= settings.passMark ? "PASS" : "SUPPLEMENTARY";
//   }

//   // 4. Update the source document (Mark or MarkDirect)
//   const updatePayload = { $set: { caTotal30: finalCA, examTotal70: finalExam, agreedMark: finalAgreed }};

//   if (isDirect) await MarkDirect.updateOne({ _id: markId }, updatePayload).session(session || null);
//   else await Mark.updateOne({ _id: markId }, updatePayload).session(session || null);
  

//   // 5. Update the FinalGrade record for the transcripts
//   await FinalGrade.findOneAndUpdate(
//     { 
//         student: markDoc.student._id || markDoc.student, 
//         programUnit: markDoc.programUnit._id || markDoc.programUnit, 
//         academicYear: markDoc.academicYear._id || markDoc.academicYear 
//     },
//     {
//       $set: {
//         totalMark: finalAgreed, 
//         grade, 
//         caTotal30: finalCA,
//         examTotal70: finalExam, 
//         status: status,
//         institution: markDoc.institution, 
//         semester: markDoc.semester || "SEMESTER 1",
//       },
//     },
//     { upsert: true, session },
//   );
 
//   return { caTotal: finalCA, examTotal: finalExam, finalMark: finalAgreed, grade, status };
// }


















// // serverside/src/services/gradeCalculator.ts
// // Complete file — integrates carry-forward resolution (ENG.14)

// import Mark from "../models/Mark";
// import MarkDirect from "../models/MarkDirect";
// import FinalGrade from "../models/FinalGrade";
// import Student from "../models/Student";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { calculateFinalResult } from "../utils/gradingCore";
// import mongoose from "mongoose";
// import { AuthenticatedRequest } from "../middleware/auth";

// interface ComputeOptions {
//   markId:          mongoose.Types.ObjectId;
//   coordinatorReq?: AuthenticatedRequest;
//   session?:        mongoose.ClientSession;
// }

// export async function computeFinalGrade({ markId, session }: ComputeOptions) {

//   // ── 1. Resolve mark document (Mark preferred, MarkDirect fallback) ─────────
//   let markDoc: any = await Mark
//     .findById(markId)
//     .populate(["student", "academicYear", "programUnit"])
//     .session(session || null);

//   let isDirect = false;
//   if (!markDoc) {
//     markDoc = await MarkDirect
//       .findById(markId)
//       .populate(["student", "academicYear", "programUnit"])
//       .session(session || null);
//     isDirect = true;
//   }

//   if (!markDoc) throw new Error("Mark not found");

//   const settings = await InstitutionSettings
//     .findOne({ institution: markDoc.institution })
//     .session(session || null);
//   if (!settings) throw new Error("Institution settings not found");

//   // ── 2. Raw calculation via gradingCore ────────────────────────────────────
//   const result = calculateFinalResult({
//     cat1:       markDoc.cat1Raw      || 0,
//     cat2:       markDoc.cat2Raw      || 0,
//     cat3:       markDoc.cat3Raw      || 0,
//     ass1:       markDoc.assgnt1Raw   || 0,
//     ass2:       markDoc.assgnt2Raw   || 0,
//     ass3:       markDoc.assgnt3Raw   || 0,
//     practical:  markDoc.practicalRaw || 0,
//     examQ1:     markDoc.examQ1Raw    || 0,
//     examQ2:     markDoc.examQ2Raw    || 0,
//     examQ3:     markDoc.examQ3Raw    || 0,
//     examQ4:     markDoc.examQ4Raw    || 0,
//     examQ5:     markDoc.examQ5Raw    || 0,
//     unitType:   (markDoc.unitType  as any) || "theory",
//     examMode:   (markDoc.examMode  as any) || "standard",
//     attempt:    markDoc.attempt       || "1st",
//     settings: {
//       catMax:       settings.cat1Max       || 30,
//       assMax:       settings.assignmentMax || 10,
//       practicalMax: settings.practicalMax  || 10,
//       passMark:     settings.passMark      || 40,
//     },
//   });

//   // ── 3. Preservation strategy ──────────────────────────────────────────────
//   // When raw calculation yields 0 but a pre-existing total exists on the mark
//   // document (e.g. manual upload, direct mark, legacy record), keep the
//   // existing value rather than overwriting with 0.
//   const finalCA    = result.caTotal   === 0 && markDoc.caTotal30   > 0
//     ? markDoc.caTotal30   : result.caTotal;
//   const finalExam  = result.examTotal === 0 && markDoc.examTotal70 > 0
//     ? markDoc.examTotal70 : result.examTotal;
//   const finalAgreed = finalCA + finalExam;

//   // ── 4. Special exam override (ENG.18) ─────────────────────────────────────
//   // Special exams are scored out of 100 including CA (ENG.18c).
//   // Grade remains "I" (Incomplete/Pending) until the special is sat.
//   // CA is locked to the value at time of special grant — do not recalculate.
//   const isSpecial = markDoc.isSpecial === true || markDoc.attempt === "special";

//   let grade:  string;
//   let status: string;

//   if (isSpecial) {
//     grade  = "I";
//     status = "SPECIAL";
//     // Exam total is 0 until the special exam is taken and marks re-uploaded
//   } else {
//     const sortedScale = [...(settings.gradingScale || [])]
//       .sort((a: any, b: any) => b.min - a.min);
//     grade  = sortedScale.find((s: any) => finalAgreed >= s.min)?.grade ?? "E";
//     status = finalAgreed >= (settings.passMark || 40) ? "PASS" : "SUPPLEMENTARY";
//   }

//   // ── 5. Update source mark document ────────────────────────────────────────
//   const markUpdate = {
//     $set: {
//       caTotal30:   finalCA,
//       examTotal70: isSpecial ? 0 : finalExam,
//       agreedMark:  isSpecial ? finalCA : finalAgreed,
//     },
//   };

//   if (isDirect) {
//     await MarkDirect.updateOne({ _id: markId }, markUpdate).session(session || null);
//   } else {
//     await Mark.updateOne({ _id: markId }, markUpdate).session(session || null);
//   }

//   // ── 6. Upsert FinalGrade (transcript record) ──────────────────────────────
//   await FinalGrade.findOneAndUpdate(
//     {
//       student:      markDoc.student._id     || markDoc.student,
//       programUnit:  markDoc.programUnit._id  || markDoc.programUnit,
//       academicYear: markDoc.academicYear._id || markDoc.academicYear,
//     },
//     {
//       $set: {
//         totalMark:    isSpecial ? finalCA : finalAgreed,
//         grade,
//         caTotal30:    finalCA,
//         examTotal70:  isSpecial ? 0 : finalExam,
//         status,
//         isSpecial,
//         attemptType:  markDoc.attempt === "supplementary" ? "SUPPLEMENTARY"
//                     : markDoc.attempt === "re-take"       ? "RETAKE"
//                     :                                       "1ST_ATTEMPT",
//         institution:  markDoc.institution,
//         semester:     markDoc.semester || "SEMESTER 1",
//       },
//     },
//     { upsert: true, session },
//   );

//   // ── 7. Carry-forward resolution (ENG.14) ──────────────────────────────────
//   // When a student passes a unit they were carrying forward, remove it from
//   // their carryForwardUnits array. If the array becomes empty, clear the
//   // qualifier suffix from their regNo.
//   //
//   // Runs AFTER the DB writes above (outside the transaction) so it cannot
//   // cause a transaction conflict. It is idempotent — safe to run multiple times.
//   if (status === "PASS") {
//     _resolveCFUnit(
//       (markDoc.student._id    || markDoc.student).toString(),
//       (markDoc.programUnit._id || markDoc.programUnit).toString(),
//     ).catch((err: any) =>
//       console.error("[gradeCalculator] CF resolution error:", err.message),
//     );
//   }

//   return {
//     caTotal:   finalCA,
//     examTotal: isSpecial ? 0 : finalExam,
//     finalMark: isSpecial ? finalCA : finalAgreed,
//     grade,
//     status,
//   };
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // PRIVATE: Carry-forward unit resolution
// // Extracted so it can be tested independently and called without awaiting
// // (fire-and-forget) from computeFinalGrade.
// // ─────────────────────────────────────────────────────────────────────────────

// async function _resolveCFUnit(
//   studentId:     string,
//   programUnitId: string,
// ): Promise<void> {
//   const studentDoc = await Student.findById(studentId)
//     .select("carryForwardUnits qualifierSuffix")
//     .lean() as any;

//   if (!studentDoc) return;

//   const hasCF = (studentDoc.carryForwardUnits || []).some(
//     (u: any) => u.programUnitId === programUnitId,
//   );
//   if (!hasCF) return;

//   // Remove this specific CF unit
//   await Student.findByIdAndUpdate(studentId, {
//     $pull: { carryForwardUnits: { programUnitId } },
//   });

//   // Re-fetch to check if array is now empty
//   const refreshed = await Student.findById(studentId)
//     .select("carryForwardUnits")
//     .lean() as any;

//   if ((refreshed?.carryForwardUnits || []).length === 0) {
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { qualifierSuffix: "" },
//     });
//   }
// }

















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