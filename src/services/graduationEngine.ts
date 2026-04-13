// // serverside/src/services/graduationEngine.ts

// import Student from "../models/Student";
// import InstitutionSettings from "../models/InstitutionSettings";

// export interface GraduationResult {
//   studentId: string;
//   regNo: string;
//   weightedAggregateAverage: number;
//   classification: string;
//   isEligible: boolean;
//   missingRequirements: string[];
// }

// export const calculateGraduationStatus = async (
//   studentId: string,
// ): Promise<GraduationResult> => {
//   const student = (await Student.findById(studentId)
//     .populate("program")
//     .lean()) as any;
//   const settings = await InstitutionSettings.findOne().lean();

//   const history = student.academicHistory || [];
//   const duration = student.program?.durationYears || 5;

//   // 1. Check for Pending Units (Incompletes, Fails, Specials) across ALL years
//   // Note: Your promoteStudent logic already pushes to history,
//   // but we must ensure Year 5 (Final Year) is also processed.

//   let totalWeightedScore = 0;
//   let totalWeightAccounted = 0;
//   const missingYears: number[] = [];

//   // 2. Sum up the weighted contributions from academic history
//   history.forEach((yearRecord: any) => {
//     totalWeightedScore += yearRecord.weightedContribution;
//     // We infer weight used by dividing contribution by mean
//     const weight =
//       yearRecord.annualMeanMark > 0
//         ? yearRecord.weightedContribution / yearRecord.annualMeanMark
//         : 0;
//     totalWeightAccounted += weight;
//   });

//   // 3. Determine Classification based on WAA
//   // Standard Engineering Scales (usually 70+, 60-69, 50-59, 40-49)
//   const waa = totalWeightedScore;
//   let classification = "FAIL";

//   if (waa >= 70) classification = "FIRST CLASS HONOURS";
//   else if (waa >= 60)
//     classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//   else if (waa >= 50)
//     classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
//   else if (waa >= 40) classification = "PASS";

//   // 4. Eligibility Check
//   const hasFailedUnits = history.some((h: any) => h.failedUnitsCount > 0);
//   const isFinalYearDone = history.some((h: any) => h.yearOfStudy === duration);

//   return {
//     studentId: student._id,
//     regNo: student.regNo,
//     weightedAggregateAverage: parseFloat(waa.toFixed(2)),
//     classification,
//     isEligible: !hasFailedUnits && isFinalYearDone,
//     missingRequirements: hasFailedUnits
//       ? ["Uncleared Failed Units in History"]
//       : [],
//   };
// };



// // serverside/src/services/graduationEngine.ts
// //
// // FIXES vs. original:
// //  1. WAA read from student.finalWeightedAverage when available (set at
// //     graduation time in promoteStudent) — most accurate single source of truth.
// //  2. Fallback: sum academicHistory[].weightedContribution directly.
// //     The original incorrectly re-derived weight from weightedContribution/annualMeanMark
// //     which loses precision (division by 0 when mean = 0).
// //  3. Eligibility check now requires:
// //     a) Student is "graduand" or "graduated" (not just "final year done")
// //     b) No pending carry-forward units
// //     c) No failed units in any year (all failedUnitsCount must be 0)
// //     d) All years of study are present in history (no missing year records)
// //  4. missingRequirements list is now specific — tells the coordinator exactly
// //     what's missing, not just a generic message.
// //  5. classification now mirrors the same logic in promoteStudent (ENG.25).

// import Student from "../models/Student";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { getYearWeight } from "../utils/weightingRegistry";

// export interface GraduationResult {
//   studentId:                string;
//   regNo:                    string;
//   name:                     string;
//   weightedAggregateAverage: number;
//   classification:           string;
//   isEligible:               boolean;
//   missingRequirements:      string[];
//   yearBreakdown:            Array<{
//     yearOfStudy:         number;
//     academicYear:        string;
//     annualMean:          number;
//     weight:              number;
//     weightedContribution: number;
//     isRepeat:            boolean;
//     failedUnits:         number;
//   }>;
// }

// export const calculateGraduationStatus = async (
//   studentId: string,
// ): Promise<GraduationResult> => {
//   const student  = await Student.findById(studentId).populate("program").lean() as any;
//   if (!student) throw new Error(`Student ${studentId} not found`);

//   const program  = student.program as any;
//   const duration = program?.durationYears || 5;
//   const history  = (student.academicHistory || []) as any[];
//   const entryType = student.entryType || "Direct";

//   const missingRequirements: string[] = [];

//   // ── 1. Check all years are in history ─────────────────────────────────────
//   const yearsPresent = new Set(history.map((h: any) => h.yearOfStudy));
//   for (let y = 1; y <= duration; y++) {
//     if (!yearsPresent.has(y)) {
//       missingRequirements.push(`Year ${y} has no academic history record`);
//     }
//   }

//   // ── 2. Check no failed units in any year ──────────────────────────────────
//   const yearsWithFails = history.filter((h: any) => (h.failedUnitsCount || 0) > 0);
//   yearsWithFails.forEach((h: any) => {
//     missingRequirements.push(
//       `Year ${h.yearOfStudy} has ${h.failedUnitsCount} uncleared failed unit(s)`,
//     );
//   });

//   // ── 3. Check no pending carry-forward units ────────────────────────────────
//   const pendingCF = (student.carryForwardUnits || []).filter(
//     (u: any) => u.status === "pending" || !u.status,
//   );
//   if (pendingCF.length > 0) {
//     missingRequirements.push(
//       `${pendingCF.length} carry-forward unit(s) not yet cleared: ${pendingCF.map((u: any) => u.unitCode).join(", ")}`,
//     );
//   }

//   // ── 4. Check student status ────────────────────────────────────────────────
//   const isGraduand = ["graduand", "graduated"].includes(student.status);
//   if (!isGraduand && missingRequirements.length === 0) {
//     missingRequirements.push(
//       `Student status is "${student.status}" — must be "graduand" or "graduated"`,
//     );
//   }

//   // ── 5. Compute WAA ────────────────────────────────────────────────────────
//   // Priority 1: use student.finalWeightedAverage (set precisely at graduation)
//   let waa: number;
//   let yearBreakdown: GraduationResult["yearBreakdown"] = [];

//   if (student.finalWeightedAverage != null && parseFloat(student.finalWeightedAverage) > 0) {
//     waa = parseFloat(student.finalWeightedAverage);

//     // Still build breakdown for display
//     yearBreakdown = history.map((h: any) => ({
//       yearOfStudy:          h.yearOfStudy,
//       academicYear:         h.academicYear || "",
//       annualMean:           h.annualMeanMark || 0,
//       weight:               getYearWeight(program, entryType, h.yearOfStudy),
//       weightedContribution: h.weightedContribution || 0,
//       isRepeat:             h.isRepeatYear || false,
//       failedUnits:          h.failedUnitsCount || 0,
//     }));
//   } else {
//     // Fallback: sum weightedContribution from history
//     // weightedContribution = annualMeanMark × yearWeight, stored at promotion time
//     let weightedSum = 0;

//     yearBreakdown = history.map((h: any) => {
//       const weight = getYearWeight(program, entryType, h.yearOfStudy);
//       const mean   = h.annualMeanMark || 0;
//       // Prefer stored weightedContribution; recompute only if missing
//       const wc     = (h.weightedContribution != null && h.weightedContribution > 0)
//         ? h.weightedContribution
//         : mean * weight;
//       weightedSum += wc;

//       return {
//         yearOfStudy:          h.yearOfStudy,
//         academicYear:         h.academicYear || "",
//         annualMean:           mean,
//         weight,
//         weightedContribution: wc,
//         isRepeat:             h.isRepeatYear || false,
//         failedUnits:          h.failedUnitsCount || 0,
//       };
//     });

//     waa = weightedSum;
//   }

//   // ── 6. Classify (ENG.25a) ─────────────────────────────────────────────────
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

// // ─── Award list: all eligible graduates for a program ────────────────────────

// export interface AwardListEntry {
//   studentId:      string;
//   regNo:          string;
//   name:           string;
//   waa:            number;
//   classification: string;
//   graduationYear: number;
// }

// export const generateAwardList = async (
//   programId:    string,
//   academicYear?: string,   // optional — filter by graduation year label
// ): Promise<AwardListEntry[]> => {
//   // Only students who have completed the program
//   const query: any = {
//     program: programId,
//     status:  { $in: ["graduand", "graduated"] },
//   };

//   const students = await Student.find(query)
//     .populate("program")
//     .lean() as any[];

//   const awardList: AwardListEntry[] = [];

//   for (const student of students) {
//     // Quick WAA resolve — prefer stored value
//     const waa = student.finalWeightedAverage != null
//       ? parseFloat(student.finalWeightedAverage)
//       : 0;

//     let classification = student.classification || "PASS";
//     if (!classification || classification === "PASS") {
//       // Re-derive classification from WAA
//       if      (waa >= 70) classification = "FIRST CLASS HONOURS";
//       else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
//       else if (waa >= 40) classification = "PASS";
//     }

//     // Filter by graduation year if provided
//     if (academicYear) {
//       const lastHistory = (student.academicHistory || []).at(-1);
//       const lastYear    = lastHistory?.academicYear || "";
//       if (!lastYear.includes(academicYear) && student.graduationYear !== parseInt(academicYear)) {
//         continue;
//       }
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

//   // Sort: First Class first, then by WAA descending
//   awardList.sort((a, b) => {
//     const order = [
//       "FIRST CLASS HONOURS",
//       "SECOND CLASS HONOURS (UPPER DIVISION)",
//       "SECOND CLASS HONOURS (LOWER DIVISION)",
//       "PASS",
//       "FAIL",
//     ];
//     const ai = order.indexOf(a.classification);
//     const bi = order.indexOf(b.classification);
//     if (ai !== bi) return ai - bi;
//     return b.waa - a.waa;
//   });

//   return awardList;
// };












// serverside/src/services/graduationEngine.ts
// COMPLETE FIX
//
// ROOT CAUSE OF WAA = 0.00 FOR ALL STUDENTS:
//   student.finalWeightedAverage was never stored for students promoted before
//   that field was added. The fallback sum of weightedContribution also fails
//   because those history entries have weightedContribution = 0 (legacy data).
//
// FIX: When both finalWeightedAverage and stored weightedContribution are
//   missing/zero, recompute WAA directly from FinalGrade marks in the DB.
//   This is the ground truth — the actual agreed marks for each unit.
//
// ROOT CAUSE OF YEAR 5 W:0%:
//   getYearWeight(program, entryType, 5) returns 0 when program.schoolType
//   is undefined and name-matching fails. Fixed in weightingRegistry.ts by
//   treating durationYears=5 as ENG_5 unconditionally.
//   ADDITIONALLY: the journey route calls calculateStudentStatus() for each
//   history year even for graduated students — that function hits the terminal
//   gate and returns weightedMean:"0.00". The stored annualMeanMark and
//   weightedContribution from academicHistory are what the timeline should use.
//   But if those are also 0 (legacy), we recompute from FinalGrade here and
//   ALSO backfill them into the DB so future calls are fast.

import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import { getYearWeight } from "../utils/weightingRegistry";

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
// Core WAA computer: given a student document, computes the WAA.
//
// Priority:
//   1. student.finalWeightedAverage  — set at graduation, most accurate
//   2. sum of academicHistory[].weightedContribution — set at each promotion
//   3. Recompute from FinalGrade marks + weightingRegistry (legacy fallback)
//      AND backfill the DB so we never need to recompute again.
// ─────────────────────────────────────────────────────────────────────────────
async function computeWAA(student: any): Promise<{
  waa:          number;
  yearBreakdown: GraduationResult["yearBreakdown"];
}> {
  const program   = student.program  as any;
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

  // ── Priority 2: sum stored weightedContribution ───────────────────────────
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

  // ── Priority 3: Recompute from FinalGrade records ─────────────────────────
  // This handles legacy students whose promotions predated the weightedContribution
  // field. We look up their FinalGrade marks for each year of study, compute
  // the annual mean, multiply by the year weight, and sum.
  console.log(`[graduationEngine] Legacy recompute for ${student.regNo} — no stored WAA`);

  const duration = program?.durationYears || 5;
  let waa = 0;
  const yearBreakdown: GraduationResult["yearBreakdown"] = [];
  const historyUpdates: any[] = []; // collect DB backfill payloads

  for (let yearOfStudy = 1; yearOfStudy <= duration; yearOfStudy++) {
    const weight = getYearWeight(program, entryType, yearOfStudy);

    // Find all ProgramUnits for this year
    const programUnits = await ProgramUnit.find({
      program:      student.program._id || student.program,
      requiredYear: yearOfStudy,
    }).lean() as any[];

    const puIds = programUnits.map((pu: any) => pu._id);

    // Get the best FinalGrade per unit (PASS > SUPPLEMENTARY > others)
    const grades = await FinalGrade.find({
      student:     student._id,
      programUnit: { $in: puIds },
    }).lean() as any[];

    // Deduplicate — keep the latest grade per programUnit
    const gradeMap = new Map<string, any>();
    grades.forEach((g: any) => {
      const key  = g.programUnit.toString();
      const prev = gradeMap.get(key);
      // Prefer PASS over SUPPLEMENTARY over others; within same status, latest wins
      if (!prev) { gradeMap.set(key, g); return; }
      const rank = (s: string) => s === "PASS" ? 2 : s === "SUPPLEMENTARY" ? 1 : 0;
      if (rank(g.status) > rank(prev.status)) gradeMap.set(key, g);
    });

    const gradedList = Array.from(gradeMap.values());
    const totalMark  = gradedList.reduce((sum: number, g: any) => sum + (g.totalMark || 0), 0);
    const unitCount  = programUnits.length || gradedList.length || 1;
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

  // ── Backfill the DB so we don't recompute next time ───────────────────────
  // This is fire-and-forget — don't await, don't crash if it fails
  _backfillStudentWAA(student._id, waa, historyUpdates).catch((err: Error) =>
    console.warn(`[graduationEngine] Backfill failed for ${student.regNo}:`, err.message),
  );

  return { waa: parseFloat(waa.toFixed(2)), yearBreakdown };
}

// ── Backfill helper: writes computed WAA back to DB ──────────────────────────
async function _backfillStudentWAA(
  studentId:      any,
  waa:            number,
  yearUpdates:    Array<{ yearOfStudy: number; annualMean: number; wc: number }>,
): Promise<void> {
  const student = await Student.findById(studentId).lean() as any;
  if (!student) return;

  // Build updated academicHistory array
  const updatedHistory = (student.academicHistory || []).map((h: any) => {
    const upd = yearUpdates.find(u => u.yearOfStudy === h.yearOfStudy);
    if (!upd) return h;
    return {
      ...h,
      annualMeanMark:       upd.annualMean,
      weightedContribution: upd.wc,
    };
  });

  // Classify for completeness
  let classification = "PASS";
  if      (waa >= 70) classification = "FIRST CLASS HONOURS";
  else if (waa >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
  else if (waa >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

  await Student.findByIdAndUpdate(studentId, {
    $set: {
      finalWeightedAverage: waa.toFixed(2),
      classification,
      academicHistory: updatedHistory,
    },
  });

  console.log(`[graduationEngine] Backfilled WAA=${waa.toFixed(2)} for student ${studentId}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Single student graduation status
// ─────────────────────────────────────────────────────────────────────────────
export const calculateGraduationStatus = async (
  studentId: string,
): Promise<GraduationResult> => {
  const student = await Student.findById(studentId)
    .populate("program")
    .lean() as any;
  if (!student) throw new Error(`Student ${studentId} not found`);

  const program  = student.program as any;
  const duration = program?.durationYears || 5;
  const history  = (student.academicHistory || []) as any[];
  const missingRequirements: string[] = [];

  // ── Eligibility checks ────────────────────────────────────────────────────
  const yearsPresent = new Set(history.map((h: any) => h.yearOfStudy));
  for (let y = 1; y <= duration; y++) {
    if (!yearsPresent.has(y)) {
      missingRequirements.push(`Year ${y} has no academic history record`);
    }
  }

  history.filter((h: any) => (h.failedUnitsCount || 0) > 0).forEach((h: any) => {
    missingRequirements.push(
      `Year ${h.yearOfStudy} has ${h.failedUnitsCount} uncleared failed unit(s)`,
    );
  });

  const pendingCF = (student.carryForwardUnits || []).filter(
    (u: any) => u.status === "pending" || !u.status,
  );
  if (pendingCF.length > 0) {
    missingRequirements.push(
      `${pendingCF.length} carry-forward unit(s) pending: ${pendingCF.map((u: any) => u.unitCode).join(", ")}`,
    );
  }

  const isGraduand = ["graduand", "graduated"].includes(student.status);
  if (!isGraduand && missingRequirements.length === 0) {
    missingRequirements.push(
      `Student status is "${student.status}" — must be "graduand" or "graduated"`,
    );
  }

  // ── WAA ───────────────────────────────────────────────────────────────────
  const { waa, yearBreakdown } = await computeWAA(student);

  // ── Classification ────────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────────────────────────────
// Award list: all eligible graduates for a program
// ─────────────────────────────────────────────────────────────────────────────
export const generateAwardList = async (
  programId:    string,
  academicYear?: string,
): Promise<AwardListEntry[]> => {
  const students = await Student.find({
    program: programId,
    status:  { $in: ["graduand", "graduated"] },
  }).populate("program").lean() as any[];

  const awardList: AwardListEntry[] = [];

  for (const student of students) {
    // Graduation year filter
    if (academicYear) {
      const lastHistory = (student.academicHistory || []).at(-1);
      const lastYear    = lastHistory?.academicYear || "";
      if (
        !lastYear.includes(academicYear) &&
        student.graduationYear !== parseInt(academicYear)
      ) continue;
    }

    // Get WAA — recomputes from FinalGrade marks if needed and backfills DB
    const { waa } = await computeWAA(student);

    // Classification — prefer stored, recompute if "PASS" or missing
    let classification = student.classification || "";
    if (!classification || classification === "PASS" || waa !== parseFloat(student.finalWeightedAverage || "0")) {
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

  // Sort: First Class → Upper → Lower → Pass → by WAA desc within class
  const classOrder = [
    "FIRST CLASS HONOURS",
    "SECOND CLASS HONOURS (UPPER DIVISION)",
    "SECOND CLASS HONOURS (LOWER DIVISION)",
    "PASS",
    "FAIL",
  ];
  awardList.sort((a, b) => {
    const ai = classOrder.indexOf(a.classification);
    const bi = classOrder.indexOf(b.classification);
    if (ai !== bi) return ai - bi;
    return b.waa - a.waa;
  });

  return awardList;
};