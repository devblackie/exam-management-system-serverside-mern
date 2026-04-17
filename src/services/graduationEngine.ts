


// // serverside/src/services/graduationEngine.ts
// // COMPLETE FIX
// //
// // ROOT CAUSE OF WAA = 0.00 FOR ALL STUDENTS:
// //   student.finalWeightedAverage was never stored for students promoted before
// //   that field was added. The fallback sum of weightedContribution also fails
// //   because those history entries have weightedContribution = 0 (legacy data).
// //
// // FIX: When both finalWeightedAverage and stored weightedContribution are
// //   missing/zero, recompute WAA directly from FinalGrade marks in the DB.
// //   This is the ground truth — the actual agreed marks for each unit.
// //
// // ROOT CAUSE OF YEAR 5 W:0%:
// //   getYearWeight(program, entryType, 5) returns 0 when program.schoolType
// //   is undefined and name-matching fails. Fixed in weightingRegistry.ts by
// //   treating durationYears=5 as ENG_5 unconditionally.
// //   ADDITIONALLY: the journey route calls calculateStudentStatus() for each
// //   history year even for graduated students — that function hits the terminal
// //   gate and returns weightedMean:"0.00". The stored annualMeanMark and
// //   weightedContribution from academicHistory are what the timeline should use.
// //   But if those are also 0 (legacy), we recompute from FinalGrade here and
// //   ALSO backfill them into the DB so future calls are fast.

// import Student from "../models/Student";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import { getYearWeight } from "../utils/weightingRegistry";

// export interface AwardListEntry {
//   studentId:      string;
//   regNo:          string;
//   name:           string;
//   waa:            number;
//   classification: string;
//   graduationYear: number;
// }

// export interface GraduationResult {
//   studentId:                string;
//   regNo:                    string;
//   name:                     string;
//   weightedAggregateAverage: number;
//   classification:           string;
//   isEligible:               boolean;
//   missingRequirements:      string[];
//   yearBreakdown:            Array<{
//     yearOfStudy:          number;
//     academicYear:         string;
//     annualMean:           number;
//     weight:               number;
//     weightedContribution: number;
//     isRepeat:             boolean;
//     failedUnits:          number;
//   }>;
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // Core WAA computer: given a student document, computes the WAA.
// //
// // Priority:
// //   1. student.finalWeightedAverage  — set at graduation, most accurate
// //   2. sum of academicHistory[].weightedContribution — set at each promotion
// //   3. Recompute from FinalGrade marks + weightingRegistry (legacy fallback)
// //      AND backfill the DB so we never need to recompute again.
// // ─────────────────────────────────────────────────────────────────────────────
// async function computeWAA(student: any): Promise<{
//   waa:          number;
//   yearBreakdown: GraduationResult["yearBreakdown"];
// }> {
//   const program   = student.program  as any;
//   const history   = (student.academicHistory || []) as any[];
//   const entryType = student.entryType || "Direct";

//   // ── Priority 1: finalWeightedAverage ─────────────────────────────────────
//   if (student.finalWeightedAverage != null) {
//     const stored = parseFloat(student.finalWeightedAverage);
//     if (stored > 0) {
//       const yearBreakdown = history.map((h: any) => ({
//         yearOfStudy:          h.yearOfStudy,
//         academicYear:         h.academicYear || "",
//         annualMean:           h.annualMeanMark || 0,
//         weight:               getYearWeight(program, entryType, h.yearOfStudy),
//         weightedContribution: h.weightedContribution || 0,
//         isRepeat:             h.isRepeatYear || false,
//         failedUnits:          h.failedUnitsCount || 0,
//       }));
//       return { waa: stored, yearBreakdown };
//     }
//   }

//   // ── Priority 2: sum stored weightedContribution ───────────────────────────
//   const totalStoredWC = history.reduce(
//     (sum: number, h: any) => sum + (h.weightedContribution || 0),
//     0,
//   );
//   if (totalStoredWC > 0) {
//     const yearBreakdown = history.map((h: any) => ({
//       yearOfStudy:          h.yearOfStudy,
//       academicYear:         h.academicYear || "",
//       annualMean:           h.annualMeanMark || 0,
//       weight:               getYearWeight(program, entryType, h.yearOfStudy),
//       weightedContribution: h.weightedContribution || 0,
//       isRepeat:             h.isRepeatYear || false,
//       failedUnits:          h.failedUnitsCount || 0,
//     }));
//     return { waa: totalStoredWC, yearBreakdown };
//   }

//   // ── Priority 3: Recompute from FinalGrade records ─────────────────────────
//   // This handles legacy students whose promotions predated the weightedContribution
//   // field. We look up their FinalGrade marks for each year of study, compute
//   // the annual mean, multiply by the year weight, and sum.
//   console.log(`[graduationEngine] Legacy recompute for ${student.regNo} — no stored WAA`);

//   const duration = program?.durationYears || 5;
//   let waa = 0;
//   const yearBreakdown: GraduationResult["yearBreakdown"] = [];
//   const historyUpdates: any[] = []; // collect DB backfill payloads

//   for (let yearOfStudy = 1; yearOfStudy <= duration; yearOfStudy++) {
//     const weight = getYearWeight(program, entryType, yearOfStudy);

//     // Find all ProgramUnits for this year
//     const programUnits = await ProgramUnit.find({
//       program:      student.program._id || student.program,
//       requiredYear: yearOfStudy,
//     }).lean() as any[];

//     const puIds = programUnits.map((pu: any) => pu._id);

//     // Get the best FinalGrade per unit (PASS > SUPPLEMENTARY > others)
//     const grades = await FinalGrade.find({
//       student:     student._id,
//       programUnit: { $in: puIds },
//     }).lean() as any[];

//     // Deduplicate — keep the latest grade per programUnit
//     const gradeMap = new Map<string, any>();
//     grades.forEach((g: any) => {
//       const key  = g.programUnit.toString();
//       const prev = gradeMap.get(key);
//       // Prefer PASS over SUPPLEMENTARY over others; within same status, latest wins
//       if (!prev) { gradeMap.set(key, g); return; }
//       const rank = (s: string) => s === "PASS" ? 2 : s === "SUPPLEMENTARY" ? 1 : 0;
//       if (rank(g.status) > rank(prev.status)) gradeMap.set(key, g);
//     });

//     const gradedList = Array.from(gradeMap.values());
//     const totalMark  = gradedList.reduce((sum: number, g: any) => sum + (g.totalMark || 0), 0);
//     const unitCount  = programUnits.length || gradedList.length || 1;
//     const annualMean = unitCount > 0 ? totalMark / unitCount : 0;
//     const wc         = annualMean * weight;

//     waa += wc;

//     yearBreakdown.push({
//       yearOfStudy,
//       academicYear:         history.find((h: any) => h.yearOfStudy === yearOfStudy)?.academicYear || "",
//       annualMean:           parseFloat(annualMean.toFixed(2)),
//       weight,
//       weightedContribution: parseFloat(wc.toFixed(4)),
//       isRepeat:             history.find((h: any) => h.yearOfStudy === yearOfStudy)?.isRepeatYear || false,
//       failedUnits:          history.find((h: any) => h.yearOfStudy === yearOfStudy)?.failedUnitsCount || 0,
//     });

//     historyUpdates.push({ yearOfStudy, annualMean, wc });
//   }

//   // ── Backfill the DB so we don't recompute next time ───────────────────────
//   // This is fire-and-forget — don't await, don't crash if it fails
//   _backfillStudentWAA(student._id, waa, historyUpdates).catch((err: Error) =>
//     console.warn(`[graduationEngine] Backfill failed for ${student.regNo}:`, err.message),
//   );

//   return { waa: parseFloat(waa.toFixed(2)), yearBreakdown };
// }

// // ── Backfill helper: writes computed WAA back to DB ──────────────────────────
// async function _backfillStudentWAA(
//   studentId:      any,
//   waa:            number,
//   yearUpdates:    Array<{ yearOfStudy: number; annualMean: number; wc: number }>,
// ): Promise<void> {
//   const student = await Student.findById(studentId).lean() as any;
//   if (!student) return;

//   // Build updated academicHistory array
//   const updatedHistory = (student.academicHistory || []).map((h: any) => {
//     const upd = yearUpdates.find(u => u.yearOfStudy === h.yearOfStudy);
//     if (!upd) return h;
//     return {
//       ...h,
//       annualMeanMark:       upd.annualMean,
//       weightedContribution: upd.wc,
//     };
//   });

//   // Classify for completeness
//   let classification = "PASS";
//   if      (waa >= 70) classification = "FIRST CLASS HONOURS";
//   else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//   else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

//   await Student.findByIdAndUpdate(studentId, {
//     $set: {
//       finalWeightedAverage: waa.toFixed(2),
//       classification,
//       academicHistory: updatedHistory,
//     },
//   });

//   console.log(`[graduationEngine] Backfilled WAA=${waa.toFixed(2)} for student ${studentId}`);
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // Single student graduation status
// // ─────────────────────────────────────────────────────────────────────────────
// export const calculateGraduationStatus = async (
//   studentId: string,
// ): Promise<GraduationResult> => {
//   const student = await Student.findById(studentId)
//     .populate("program")
//     .lean() as any;
//   if (!student) throw new Error(`Student ${studentId} not found`);

//   const program  = student.program as any;
//   const duration = program?.durationYears || 5;
//   const history  = (student.academicHistory || []) as any[];
//   const missingRequirements: string[] = [];

//   // ── Eligibility checks ────────────────────────────────────────────────────
//   const yearsPresent = new Set(history.map((h: any) => h.yearOfStudy));
//   for (let y = 1; y <= duration; y++) {
//     if (!yearsPresent.has(y)) {
//       missingRequirements.push(`Year ${y} has no academic history record`);
//     }
//   }

//   history.filter((h: any) => (h.failedUnitsCount || 0) > 0).forEach((h: any) => {
//     missingRequirements.push(
//       `Year ${h.yearOfStudy} has ${h.failedUnitsCount} uncleared failed unit(s)`,
//     );
//   });

//   const pendingCF = (student.carryForwardUnits || []).filter(
//     (u: any) => u.status === "pending" || !u.status,
//   );
//   if (pendingCF.length > 0) {
//     missingRequirements.push(
//       `${pendingCF.length} carry-forward unit(s) pending: ${pendingCF.map((u: any) => u.unitCode).join(", ")}`,
//     );
//   }

//   const isGraduand = ["graduand", "graduated"].includes(student.status);
//   if (!isGraduand && missingRequirements.length === 0) {
//     missingRequirements.push(
//       `Student status is "${student.status}" — must be "graduand" or "graduated"`,
//     );
//   }

//   // ── WAA ───────────────────────────────────────────────────────────────────
//   const { waa, yearBreakdown } = await computeWAA(student);

//   // ── Classification ────────────────────────────────────────────────────────
//   let classification = "FAIL";
//   if      (waa >= 70) classification = "FIRST CLASS HONOURS";
//   else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//   else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
//   else if (waa >= 40) classification = "PASS";

//   return {
//     studentId:                student._id.toString(),
//     regNo:                    student.regNo,
//     name:                     student.name,
//     weightedAggregateAverage: parseFloat(waa.toFixed(2)),
//     classification,
//     isEligible:               missingRequirements.length === 0,
//     missingRequirements,
//     yearBreakdown,
//   };
// };

// // ─────────────────────────────────────────────────────────────────────────────
// // Award list: all eligible graduates for a program
// // ─────────────────────────────────────────────────────────────────────────────
// export const generateAwardList = async (
//   programId:    string,
//   academicYear?: string,
// ): Promise<AwardListEntry[]> => {
//   const students = await Student.find({
//     program: programId,
//     status:  { $in: ["graduand", "graduated"] },
//   }).populate("program").lean() as any[];

//   const awardList: AwardListEntry[] = [];

//   for (const student of students) {
//     // Graduation year filter
//     if (academicYear) {
//       const lastHistory = (student.academicHistory || []).at(-1);
//       const lastYear    = lastHistory?.academicYear || "";
//       if (
//         !lastYear.includes(academicYear) &&
//         student.graduationYear !== parseInt(academicYear)
//       ) continue;
//     }

//     // Get WAA — recomputes from FinalGrade marks if needed and backfills DB
//     const { waa } = await computeWAA(student);

//     // Classification — prefer stored, recompute if "PASS" or missing
//     let classification = student.classification || "";
//     if (!classification || classification === "PASS" || waa !== parseFloat(student.finalWeightedAverage || "0")) {
//       if      (waa >= 70) classification = "FIRST CLASS HONOURS";
//       else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
//       else if (waa >= 40) classification = "PASS";
//       else                classification = "FAIL";
//     }

//     awardList.push({
//       studentId:      student._id.toString(),
//       regNo:          student.regNo,
//       name:           student.name,
//       waa:            parseFloat(waa.toFixed(2)),
//       classification,
//       graduationYear: student.graduationYear || new Date().getFullYear(),
//     });
//   }

//   // Sort: First Class → Upper → Lower → Pass → by WAA desc within class
//   const classOrder = [
//     "FIRST CLASS HONOURS",
//     "SECOND CLASS HONOURS (UPPER DIVISION)",
//     "SECOND CLASS HONOURS (LOWER DIVISION)",
//     "PASS",
//     "FAIL",
//   ];
//   awardList.sort((a, b) => {
//     const ai = classOrder.indexOf(a.classification);
//     const bi = classOrder.indexOf(b.classification);
//     if (ai !== bi) return ai - bi;
//     return b.waa - a.waa;
//   });

//   return awardList;
// };










// serverside/src/services/graduationEngine.ts
// COMPLETE FIX — v2
//
// KEY FIX: computeWAA() Priority 3 now reads BOTH FinalGrade AND MarkDirect
// when computing the annual mean per year. Previously it only read FinalGrade,
// but direct-entry marks (MarkDirect) have no FinalGrade records unless
// computeFinalGrade() was called — which it wasn't for direct uploads until
// directMarksImporter_FIXED.ts was deployed.
//
// For legacy data where FinalGrade is empty, we now resolve marks the same
// way journeyCMS.ts does: FinalGrade → MarkDirect → Mark (in that order).
// This ensures WAA is always computed from actual uploaded marks regardless
// of entry mode.

import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import MarkDirect from "../models/MarkDirect";
import Mark from "../models/Mark";
import ProgramUnit from "../models/ProgramUnit";
import { getYearWeight } from "../utils/weightingRegistry";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface AwardListEntry {
  studentId:      string;
  regNo:          string;
  name:           string;
  waa:            number;
  classification: string;
  graduationYear: number;
}

export interface GraduationResult {
  studentId:                string;
  regNo:                    string;
  name:                     string;
  weightedAggregateAverage: number;
  classification:           string;
  isEligible:               boolean;
  missingRequirements:      string[];
  yearBreakdown:            Array<{
    yearOfStudy:          number;
    academicYear:         string;
    annualMean:           number;
    weight:               number;
    weightedContribution: number;
    isRepeat:             boolean;
    failedUnits:          number;
  }>;
}

// ─────────────────────────────────────────────────────────────────────────────
// Resolve the best agreed mark for a student + programUnit across all sources.
// Priority: FinalGrade → MarkDirect → Mark
// This is the SINGLE source-of-truth resolver used by both WAA computation
// and the Journey CMS mark display.
// ─────────────────────────────────────────────────────────────────────────────
export async function resolveAgreedMark(
  studentId: string,
  puId:      string,
): Promise<number> {
  // 1. FinalGrade (set by computeFinalGrade — most authoritative)
  const fgs = await FinalGrade.find({
    student:     studentId,
    programUnit: puId,
  }).lean() as any[];

  if (fgs.length > 0) {
    const rank   = (s: string) => s === "PASS" ? 3 : s === "SPECIAL" ? 2 : s === "SUPPLEMENTARY" ? 1 : 0;
    const best   = fgs.sort((a: any, b: any) => rank(b.status) - rank(a.status))[0];
    const mark   = best.totalMark ?? 0;
    if (mark > 0) return mark;
  }

  // 2. MarkDirect (direct CA+Exam entry — no FinalGrade until directMarksImporter_FIXED runs)
  const md = await MarkDirect.findOne({ student: studentId, programUnit: puId }).lean() as any;
  if (md && (md.agreedMark ?? 0) > 0) return md.agreedMark;

  // 3. Mark (detailed breakdown with computeFinalGrade applied)
  const dm = await Mark.findOne({ student: studentId, programUnit: puId }).lean() as any;
  if (dm && (dm.agreedMark ?? 0) > 0) return dm.agreedMark;

  return 0; // no mark found
}

// ─────────────────────────────────────────────────────────────────────────────
// Core WAA computer
//
// Priority:
//   1. student.finalWeightedAverage  — stored at graduation (most accurate)
//   2. sum of academicHistory[].weightedContribution — stored at promotion
//   3. Recompute from actual marks (FinalGrade + MarkDirect + Mark fallback)
//      and backfill the DB so future calls use Priority 1.
// ─────────────────────────────────────────────────────────────────────────────
async function computeWAA(student: any): Promise<{
  waa:           number;
  yearBreakdown: GraduationResult["yearBreakdown"];
}> {
  const program   = student.program as any;
  const history   = (student.academicHistory || []) as any[];
  const entryType = student.entryType || "Direct";

  // ── Priority 1: finalWeightedAverage ─────────────────────────────────────
  if (student.finalWeightedAverage != null) {
    const stored = parseFloat(student.finalWeightedAverage);
    if (stored > 0) {
      const yearBreakdown = history.map((h: any) => ({
        yearOfStudy:          h.yearOfStudy,
        academicYear:         h.academicYear || "",
        annualMean:           h.annualMeanMark || 0,
        weight:               getYearWeight(program, entryType, h.yearOfStudy),
        weightedContribution: h.weightedContribution || 0,
        isRepeat:             h.isRepeatYear || false,
        failedUnits:          h.failedUnitsCount || 0,
      }));
      return { waa: stored, yearBreakdown };
    }
  }

  // ── Priority 2: stored weightedContribution ───────────────────────────────
  const totalStoredWC = history.reduce(
    (sum: number, h: any) => sum + (h.weightedContribution || 0),
    0,
  );
  if (totalStoredWC > 0) {
    const yearBreakdown = history.map((h: any) => ({
      yearOfStudy:          h.yearOfStudy,
      academicYear:         h.academicYear || "",
      annualMean:           h.annualMeanMark || 0,
      weight:               getYearWeight(program, entryType, h.yearOfStudy),
      weightedContribution: h.weightedContribution || 0,
      isRepeat:             h.isRepeatYear || false,
      failedUnits:          h.failedUnitsCount || 0,
    }));
    return { waa: totalStoredWC, yearBreakdown };
  }

  // ── Priority 3: Recompute from actual marks ───────────────────────────────
  // Reads FinalGrade first, then MarkDirect, then Mark as fallback.
  // Handles legacy students and students whose direct marks have no FinalGrade yet.
  // console.log(`[graduationEngine] Recomputing WAA from marks for ${student.regNo}`);

  const duration       = program?.durationYears || 5;
  let waa              = 0;
  const yearBreakdown: GraduationResult["yearBreakdown"] = [];
  const historyUpdates: Array<{ yearOfStudy: number; annualMean: number; wc: number }> = [];

  for (let yearOfStudy = 1; yearOfStudy <= duration; yearOfStudy++) {
    const weight = getYearWeight(program, entryType, yearOfStudy);

    const programUnits = await ProgramUnit.find({
      program:      student.program._id || student.program,
      requiredYear: yearOfStudy,
    }).lean() as any[];

    const unitCount = programUnits.length || 1;
    let   totalMark = 0;
    let   resolved  = 0; // how many units have a mark

    for (const pu of programUnits) {
      const puId = (pu as any)._id.toString();
      const mark = await resolveAgreedMark(student._id.toString(), puId);
      if (mark > 0) {
        totalMark += mark;
        resolved++;
      }
    }

    // If no marks at all for this year, skip (student may not have sat exams)
    const annualMean = unitCount > 0 ? totalMark / unitCount : 0;
    const wc         = annualMean * weight;

    waa += wc;

    yearBreakdown.push({
      yearOfStudy,
      academicYear:         history.find((h: any) => h.yearOfStudy === yearOfStudy)?.academicYear || "",
      annualMean:           parseFloat(annualMean.toFixed(2)),
      weight,
      weightedContribution: parseFloat(wc.toFixed(4)),
      isRepeat:             history.find((h: any) => h.yearOfStudy === yearOfStudy)?.isRepeatYear || false,
      failedUnits:          history.find((h: any) => h.yearOfStudy === yearOfStudy)?.failedUnitsCount || 0,
    });

    historyUpdates.push({ yearOfStudy, annualMean, wc });
  }

  // Backfill DB (fire-and-forget)
  _backfillStudentWAA(student._id, waa, historyUpdates).catch((err: Error) =>
    console.warn(`[graduationEngine] Backfill failed for ${student.regNo}:`, err.message),
  );

  return { waa: parseFloat(waa.toFixed(2)), yearBreakdown };
}

// ─── DB backfill ──────────────────────────────────────────────────────────────

async function _backfillStudentWAA(
  studentId:    any,
  waa:          number,
  yearUpdates:  Array<{ yearOfStudy: number; annualMean: number; wc: number }>,
): Promise<void> {
  const student = await Student.findById(studentId).lean() as any;
  if (!student) return;

  const updatedHistory = (student.academicHistory || []).map((h: any) => {
    const upd = yearUpdates.find(u => u.yearOfStudy === h.yearOfStudy);
    if (!upd) return h;
    return { ...h, annualMeanMark: upd.annualMean, weightedContribution: upd.wc };
  });

  let classification = "PASS";
  if      (waa >= 70) classification = "FIRST CLASS HONOURS";
  else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
  else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

  await Student.findByIdAndUpdate(studentId, {
    $set: {
      finalWeightedAverage: waa.toFixed(2),
      classification,
      academicHistory:      updatedHistory,
    },
  });

  // console.log(`[graduationEngine] Backfilled WAA=${waa.toFixed(2)}, class=${classification} for ${studentId}`);
}

// ─── Single student graduation status ─────────────────────────────────────────

export const calculateGraduationStatus = async (
  studentId: string,
): Promise<GraduationResult> => {
  const student = await Student.findById(studentId).populate("program").lean() as any;
  if (!student) throw new Error(`Student ${studentId} not found`);

  const program  = student.program as any;
  const duration = program?.durationYears || 5;
  const history  = (student.academicHistory || []) as any[];
  const missingRequirements: string[] = [];

  // Eligibility checks
  const yearsPresent = new Set(history.map((h: any) => h.yearOfStudy));
  for (let y = 1; y <= duration; y++) {
    if (!yearsPresent.has(y)) missingRequirements.push(`Year ${y} has no academic history record`);
  }
  history.filter((h: any) => (h.failedUnitsCount || 0) > 0).forEach((h: any) => {
    missingRequirements.push(`Year ${h.yearOfStudy} has ${h.failedUnitsCount} uncleared failed unit(s)`);
  });
  const pendingCF = (student.carryForwardUnits || []).filter(
    (u: any) => u.status === "pending" || !u.status,
  );
  if (pendingCF.length > 0) {
    missingRequirements.push(
      `${pendingCF.length} carry-forward unit(s) pending: ${pendingCF.map((u: any) => u.unitCode).join(", ")}`,
    );
  }
  if (!["graduand","graduated"].includes(student.status) && missingRequirements.length === 0) {
    missingRequirements.push(`Student status is "${student.status}" — must be "graduand" or "graduated"`);
  }

  const { waa, yearBreakdown } = await computeWAA(student);

  let classification = "FAIL";
  if      (waa >= 70) classification = "FIRST CLASS HONOURS";
  else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
  else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
  else if (waa >= 40) classification = "PASS";

  return {
    studentId:                student._id.toString(),
    regNo:                    student.regNo,
    name:                     student.name,
    weightedAggregateAverage: parseFloat(waa.toFixed(2)),
    classification,
    isEligible:               missingRequirements.length === 0,
    missingRequirements,
    yearBreakdown,
  };
};

// ─── Award list ───────────────────────────────────────────────────────────────

export const generateAwardList = async (
  programId:     string,
  academicYear?: string,
): Promise<AwardListEntry[]> => {
  const students = await Student.find({
    program: programId,
    status:  { $in: ["graduand","graduated"] },
  }).populate("program").lean() as any[];

  const awardList: AwardListEntry[] = [];

  for (const student of students) {
    if (academicYear) {
      const lastHistory = (student.academicHistory || []).at(-1);
      const lastYear    = lastHistory?.academicYear || "";
      if (!lastYear.includes(academicYear) && student.graduationYear !== parseInt(academicYear)) continue;
    }

    const { waa } = await computeWAA(student);

    let classification = student.classification || "";
    if (!classification || (classification === "PASS" && waa > 50)) {
      if      (waa >= 70) classification = "FIRST CLASS HONOURS";
      else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
      else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
      else if (waa >= 40) classification = "PASS";
      else                classification = "FAIL";
    }

    awardList.push({
      studentId:      student._id.toString(),
      regNo:          student.regNo,
      name:           student.name,
      waa:            parseFloat(waa.toFixed(2)),
      classification,
      graduationYear: student.graduationYear || new Date().getFullYear(),
    });
  }

  const classOrder = [
    "FIRST CLASS HONOURS",
    "SECOND CLASS HONOURS (UPPER DIVISION)",
    "SECOND CLASS HONOURS (LOWER DIVISION)",
    "PASS", "FAIL",
  ];
  awardList.sort((a, b) => {
    const ai = classOrder.indexOf(a.classification);
    const bi = classOrder.indexOf(b.classification);
    if (ai !== bi) return ai - bi;
    return b.waa - a.waa;
  });

  return awardList;
};