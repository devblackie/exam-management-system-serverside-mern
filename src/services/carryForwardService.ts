// // src/services/carryForwardService.ts
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import Mark from "../models/Mark";
// import MarkDirect from "../models/MarkDirect";
// import ProgramUnit from "../models/ProgramUnit";
// import InstitutionSettings from "../models/InstitutionSettings";
// import {
//   assessCarryForwardEligibility,
//   REG_QUALIFIERS,
// } from "../utils/academicRules";

// // ─── Types ────────────────────────────────────────────────────────────────────

// export interface CFUnit {
//   programUnitId: string;
//   unitCode: string;
//   unitName: string;
// }

// export interface CarryForwardResult {
//   eligible: boolean;
//   promoted: boolean;
//   units: CFUnit[];
//   qualifier: string;
//   reason: string;
// }

// // ─── assessCarryForward ───────────────────────────────────────────────────────
// // Determines if a student qualifies to carry forward failed units under ENG.14.
// // Called AFTER supplementary results are uploaded (session = SUPPLEMENTARY).
// //
// // Logic:
// //   - Collect all units in the prescribed curriculum for this year
// //   - Find which are failed at the supplementary stage
// //   - Exclude units failed because of missing CA (ENG.15a)
// //   - Allow max 2 carry-forward units (ENG.14a)
// //   - Block if this is the final year (ENG.14a — no CF to Year 5/4)

// export const assessCarryForward = async (
//   studentId: string,
//   programId: string,
//   academicYearName: string,
//   yearOfStudy: number,
// ): Promise<CarryForwardResult> => {
//   const student = await Student.findById(studentId).lean();
//   if (!student) throw new Error(`Student ${studentId} not found`);

//   const settings = await InstitutionSettings.findOne({
//     institution: (student as any).institution,
//   }).lean();
//   const passMark = settings?.passMark ?? 40;

//   // Check final year restriction
//   const programDoc = (await mongoose
//     .model("Program")
//     .findById(programId)
//     .lean()) as any;
//   const finalYear = programDoc?.durationYears || 5;
//   if (yearOfStudy >= finalYear) {
//     return {
//       eligible: false,
//       promoted: false,
//       units: [],
//       qualifier: "",
//       reason: `ENG.14: Carry-forward not permitted to final year (Year ${finalYear}).`,
//     };
//   }

//   const programUnits = await ProgramUnit.find({
//     program: programId,
//     requiredYear: yearOfStudy,
//   })
//     .populate("unit")
//     .lean();

//   const totalUnits = programUnits.length;
//   const puIds = programUnits.map((pu: any) => pu._id);

//   const [detailedMarks, directMarks] = await Promise.all([
//     Mark.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
//     MarkDirect.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
//   ]);

//   const markMap = new Map<string, any>();
//   [...detailedMarks, ...directMarks].forEach((m: any) => {
//     markMap.set(m.programUnit.toString(), m);
//   });

//   const failedUnitIds: string[] = [];
//   const noCAUnitIds: string[] = [];
//   const failedDetails: CFUnit[] = [];

//   for (const pu of programUnits) {
//     const puId = (pu as any)._id.toString();
//     const m = markMap.get(puId);
//     if (!m) continue;

//     const isSpecial = m.isSpecial || m.attempt === "special";
//     if (isSpecial) continue;

//     const mark = m.agreedMark ?? 0;
//     const hasCA = (m.caTotal30 ?? 0) > 0;

//     if (mark < passMark) {
//       failedUnitIds.push(puId);
//       if (!hasCA) noCAUnitIds.push(puId); // ENG.15a: missing CA → cannot CF
//       failedDetails.push({
//         programUnitId: puId,
//         unitCode: (pu as any).unit?.code || "N/A",
//         unitName: (pu as any).unit?.name || "N/A",
//       });
//     }
//   }

//   const eligibility = assessCarryForwardEligibility(
//     failedUnitIds,
//     noCAUnitIds,
//     totalUnits,
//   );

//   return {
//     eligible: eligibility.eligible,
//     promoted: eligibility.eligible,
//     units: eligibility.units.map(
//       (id) =>
//         failedDetails.find((d) => d.programUnitId === id) || {
//           programUnitId: id,
//           unitCode: "N/A",
//           unitName: "N/A",
//         },
//     ),
//     qualifier: eligibility.eligible ? REG_QUALIFIERS.carryForward(1) : "",
//     reason: eligibility.reason,
//   };
// };

// // ─── applyCarryForward ────────────────────────────────────────────────────────
// // Promotes the student to the next year and records the CF units on their record.

// export const applyCarryForward = async (
//   studentId: string,
//   programId: string,
//   academicYearName: string,
//   yearOfStudy: number,
//   cfUnits: CFUnit[],
// ): Promise<void> => {
//   const nextYear = yearOfStudy + 1;

//   const existing = (await Student.findById(studentId)
//     .select("qualifierSuffix")
//     .lean()) as any;
//   const priorQualifier = existing?.qualifierSuffix || "";

//   const priorCF = (priorQualifier.match(/RP(\d+)C/) ?? [])[1];
//   const cfCount = priorCF ? parseInt(priorCF) + 1 : 1;
//   const newQualifier = REG_QUALIFIERS.carryForward(cfCount);

//   const entries = cfUnits.map((u) => ({
//     programUnitId: new mongoose.Types.ObjectId(u.programUnitId),
//     unitCode: u.unitCode,
//     unitName: u.unitName,
//     fromYear: yearOfStudy,
//     fromAcademicYear: academicYearName,
//     attemptCount: cfCount,
//     status: "pending" as const,
//   }));

//   await Student.findByIdAndUpdate(studentId, {
//     $set: {
//       currentYearOfStudy: nextYear,
//       currentSemester: 1,
//       qualifierSuffix: newQualifier,
//     },
//     $push: {
//       carryForwardUnits: { $each: entries },
//       statusEvents: {
//         fromStatus: `year_${yearOfStudy}`,
//         toStatus: `year_${nextYear}_cf`,
//         date: new Date(),
//         academicYear: academicYearName,
//         reason: `ENG.14 Carry Forward to Year ${nextYear}. Units: ${cfUnits.map((u) => u.unitCode).join(", ")}. Qualifier: ${newQualifier}`,
//       },
//       statusHistory: {
//         status: "active",
//         previousStatus: "active",
//         date: new Date(),
//         reason: `Carry-forward promotion to Year ${nextYear} (${newQualifier})`,
//       },
//     },
//   });
// };

// // ─── resolveCarryForwardUnit ──────────────────────────────────────────────────
// // Updates one CF unit entry after results are processed.

// export const resolveCarryForwardUnit = async (
//   studentId: string,
//   programUnitId: string,
//   outcome: "passed" | "failed" | "escalated_to_rpu",
// ): Promise<void> => {
//   await Student.updateOne(
//     {
//       _id: studentId,
//       "carryForwardUnits.programUnitId": new mongoose.Types.ObjectId(
//         programUnitId,
//       ),
//     },
//     { $set: { "carryForwardUnits.$.status": outcome } },
//   );

//   if (outcome === "escalated_to_rpu") {
//     const s = (await Student.findById(studentId)
//       .select("qualifierSuffix")
//       .lean()) as any;
//     const prior = s?.qualifierSuffix || "";
//     const rpu = (prior.match(/RPU(\d+)/) ?? [])[1];
//     const rpuN = rpu ? parseInt(rpu) + 1 : 1;

//     await Student.findByIdAndUpdate(studentId, {
//       $set: { qualifierSuffix: `RPU${rpuN}` },
//       $push: {
//         statusEvents: {
//           fromStatus: "carry_forward",
//           toStatus: "repeat_unit",
//           date: new Date(),
//           academicYear: "CURRENT",
//           reason: `ENG.16b: Failed CF supplementary. Must repeat unit (RPU${rpuN}).`,
//         },
//       },
//     });
//   }

//   // If all pending CF units are resolved and all passed, clear the qualifier
//   const student = (await Student.findById(studentId).lean()) as any;
//   const cfUnits = student?.carryForwardUnits || [];
//   const allResolved = cfUnits.every((u: any) => u.status !== "pending");
//   const allPassed = cfUnits.every((u: any) => u.status === "passed");

//   if (allResolved && allPassed) {
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { qualifierSuffix: "" },
//     });
//   }
// };













// // serverside/src/services/carryForwardService.ts
// //
// // ENG.14 CARRY FORWARD — Complete Implementation
// //
// // WHAT CARRY FORWARD IS:
// //   After sitting supplementary exams, a student who has failed ≤ 2 units
// //   (and NOT because of missing coursework) may be allowed to proceed to
// //   the next year while those units are "carried forward".
// //
// // HOW IT WORKS IN THE SYSTEM:
// //
// //   1. GRANT CARRY FORWARD (called from promoteStudent when conditions met)
// //      - Student passed overall (mean ≥ 40%, failed ≤ 1/3 BUT also failed supp)
// //      - Failed ≤ 2 units at supplementary, none due to missing CA
// //      - We store the CF units in student.carryForwardUnits[]
// //      - We set student.qualifierSuffix = "RP1C" (or RP2C if already had CF before)
// //      - Student's currentYearOfStudy increments normally
// //
// //   2. APPEARING ON SCORESHEETS
// //      - In the ORDINARY session of the NEXT year, CF students appear on
// //        scoresheets for their carried units with attempt "RP1C"
// //      - The scoresheet generator queries student.carryForwardUnits to find
// //        which units to include them on
// //
// //   3. PASSING A CF UNIT
// //      - When marks are uploaded and the grade is PASS for a CF unit:
// //        clearCarryForwardUnit(studentId, programUnitId)
// //      - If all CF units cleared: qualifierSuffix reverts to "" (or next qualifier)
// //
// //   4. FAILING A CF UNIT AT ORDINARY
// //      - The student gets a supplementary for that unit (ENG.22(b) step 4)
// //      - If they fail supp: the unit goes to "Repeat Unit" (RPU1) — ENG.22(b) step 5
// //
// //   5. ENG.14 RESTRICTION: Cannot graduate with pending CF units
// //      - promoteStudent checks carryForwardUnits.length === 0 before final year graduation

// import mongoose from "mongoose";
// import Student from "../models/Student";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import { CarryForwardUnit } from "./carryForwardTypes";
// import { REG_QUALIFIERS, assessCarryForwardEligibility } from "../utils/academicRules";

// // ─────────────────────────────────────────────────────────────────────────────
// // ASSESS AND GRANT CARRY FORWARD
// // Called from the promotion flow after supplementary results are finalized.
// // ─────────────────────────────────────────────────────────────────────────────

// export interface CarryForwardResult {
//   granted:       boolean;
//   cfUnits:       CarryForwardUnit[];
//   qualifier:     string;
//   reason:        string;
// }

// export const assessAndGrantCarryForward = async (
//   studentId:        string,
//   programId:        string,
//   yearOfStudy:      number,
//   academicYearName: string,
// ): Promise<CarryForwardResult> => {

//   const student = await Student.findById(studentId).lean();
//   if (!student) throw new Error("Student not found");

//   // Get all final grades for this student in this year
//   const programUnitsThisYear = await ProgramUnit.find({
//     program:      programId,
//     requiredYear: yearOfStudy,
//   }).populate("unit").lean() as any[];

//   const grades = await FinalGrade.find({
//     student:     studentId,
//     programUnit: { $in: programUnitsThisYear.map(pu => pu._id) },
//   }).lean();

//   const totalUnits = programUnitsThisYear.length;

//   // Find failed units at supplementary (attempt 2)
//   const failedAtSupp = grades.filter(g =>
//     g.status !== "PASS" &&
//     g.status !== "SPECIAL" &&
//     (g.attemptType === "SUPPLEMENTARY" || g.attemptType === "1ST_ATTEMPT")
//   );

//   // Find units failed due to missing CA (isMissingCA flag on FinalGrade)
//   // These cannot be carried forward per ENG.15a
//   const failedDueToNoCA: string[] = [];
//   for (const g of failedAtSupp) {
//     const pu = programUnitsThisYear.find(pu =>
//       pu._id.toString() === (g.programUnit as any).toString()
//     );
//     if (pu) {
//       // Check if the mark record has isMissingCA
//       const Mark       = require("../models/Mark").default;
//       const MarkDirect = require("../models/MarkDirect").default;
//       const mark = await Mark.findOne({ student: studentId, programUnit: pu._id }).lean()
//              || await MarkDirect.findOne({ student: studentId, programUnit: pu._id }).lean();
//       if (mark?.isMissingCA) failedDueToNoCA.push(pu.unit.code);
//     }
//   }

//   const failedUnitCodes = failedAtSupp
//     .map(g => {
//       const pu = programUnitsThisYear.find(pu =>
//         pu._id.toString() === (g.programUnit as any).toString()
//       );
//       return pu?.unit.code || "";
//     })
//     .filter(Boolean);

//   // Assess eligibility
//   const eligibility = assessCarryForwardEligibility(
//     failedUnitCodes,
//     failedDueToNoCA,
//     totalUnits,
//   );

//   if (!eligibility.eligible) {
//     return { granted: false, cfUnits: [], qualifier: "", reason: eligibility.reason };
//   }

//   // Determine qualifier number (how many CF cycles has this student had before?)
//   const existingCFCount = (student.carryForwardUnits || []).length > 0
//     ? Math.max(...(student.carryForwardUnits || []).map((u: any) => {
//         const match = u.qualifier.match(/RP(\d+)C/);
//         return match ? parseInt(match[1]) : 0;
//       }))
//     : 0;
//   const cfNumber    = Math.min(existingCFCount + 1, 3);
//   const qualifier   = REG_QUALIFIERS.carryForward(cfNumber);

//   // Build CF unit records
//   const cfUnits: CarryForwardUnit[] = eligibility.units.map(unitCode => {
//     const pu = programUnitsThisYear.find(pu => pu.unit.code === unitCode)!;
//     return {
//       programUnitId:    pu._id.toString(),
//       unitCode:         pu.unit.code,
//       unitName:         pu.unit.name,
//       fromYear:         yearOfStudy,
//       fromAcademicYear: academicYearName,
//       attemptNumber:    cfNumber + 1,  // attempt 3 (1st=ordinary, 2nd=supp, 3rd=CF)
//       qualifier,
//       addedAt:          new Date(),
//     };
//   });

//   // Persist to student record
//   await Student.findByIdAndUpdate(studentId, {
//     $push:  { carryForwardUnits: { $each: cfUnits } },
//     $set:   { qualifierSuffix: qualifier },
//   });

//   return {
//     granted:   true,
//     cfUnits,
//     qualifier,
//     reason:    `ENG.14 carry forward granted — ${cfUnits.length} unit(s): ${cfUnits.map(u => u.unitCode).join(", ")}`,
//   };
// };

// // ─────────────────────────────────────────────────────────────────────────────
// // CLEAR A CARRY FORWARD UNIT (student passed it)
// // Called from computeFinalGrade when a CF unit is graded PASS.
// // ─────────────────────────────────────────────────────────────────────────────

// export const clearCarryForwardUnit = async (
//   studentId:     string,
//   programUnitId: string,
// ): Promise<void> => {
//   await Student.findByIdAndUpdate(studentId, {
//     $pull: { carryForwardUnits: { programUnitId } },
//   });

//   // After clearing, check if all CF units are resolved
//   const updated = await Student.findById(studentId).lean();
//   const remaining = (updated?.carryForwardUnits || []).length;

//   if (remaining === 0) {
//     // All CF units cleared — remove the CF qualifier from regNo
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { qualifierSuffix: "" },
//     });
//   }
// };

// // ─────────────────────────────────────────────────────────────────────────────
// // GET STUDENTS WITH CARRY FORWARD UNITS FOR A SPECIFIC PROGRAM UNIT
// // Used by the scoresheet generator to include CF students on ordinary scoresheets
// // ─────────────────────────────────────────────────────────────────────────────

// export const getCarryForwardStudentsForUnit = async (
//   programUnitId: string,
//   programId:     string,
// ): Promise<Array<{
//   student:         any;
//   cfUnit:          CarryForwardUnit;
//   attemptLabel:    string;
// }>> => {
//   const students = await Student.find({
//     program:              programId,
//     "carryForwardUnits.programUnitId": programUnitId.toString(),
//   }).lean() as any[];

//   return students.map(student => {
//     const cfUnit = student.carryForwardUnits.find(
//       (u: any) => u.programUnitId === programUnitId.toString()
//     );
//     return {
//       student,
//       cfUnit,
//       attemptLabel: cfUnit?.qualifier || "RP1C",
//     };
//   });
// };

// // ─────────────────────────────────────────────────────────────────────────────
// // GET STAYOUT STUDENTS FOR A SPECIFIC PROGRAM UNIT IN NEXT YEAR
// // Stayout students (ENG.15h) retake failed units in ordinary of NEXT year.
// // ─────────────────────────────────────────────────────────────────────────────

// export const getStayoutStudentsForUnit = async (
//   programUnitId: string,
//   programId:     string,
// ): Promise<Array<{ student: any; attemptLabel: string }>> => {
//   // Stayout students: status is "active" (they are in next year now)
//   // but their previous year had status STAYOUT and they have a failed grade
//   // for this specific programUnit
//   const failedGrades = await FinalGrade.find({
//     programUnit: programUnitId,
//     status:      { $ne: "PASS" },
//     attemptType: { $in: ["1ST_ATTEMPT", "SUPPLEMENTARY"] },
//   }).populate("student").lean() as any[];

//   const result: Array<{ student: any; attemptLabel: string }> = [];

//   for (const grade of failedGrades) {
//     const student = grade.student as any;
//     if (!student) continue;
//     if (student.program?.toString() !== programId) continue;
//     // Student should now be in the NEXT year (not the year this unit belongs to)
//     const pu = await ProgramUnit.findById(programUnitId).lean() as any;
//     if (!pu) continue;
//     if (student.currentYearOfStudy !== pu.requiredYear + 1) continue;
//     if (student.status !== "active") continue;

//     result.push({
//       student,
//       attemptLabel: "A/SO",
//     });
//   }

//   return result;
// };












// serverside/src/services/carryForwardService.ts
import mongoose from "mongoose";
import Student from "../models/Student";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import InstitutionSettings from "../models/InstitutionSettings";
import type { CarryForwardUnit } from "./carryForwardTypes";
import { REG_QUALIFIERS, assessCarryForwardEligibility } from "../utils/academicRules";

export interface CarryForwardResult {
  granted:   boolean;
  cfUnits:   CarryForwardUnit[];
  qualifier: string;
  reason:    string;
}

// ─── assessAndGrantCarryForward ───────────────────────────────────────────────
// Called from promoteStudent after supplementary results are finalized.
// Determines carry-forward eligibility per ENG.14 and persists to student record.

export const assessAndGrantCarryForward = async (
  studentId:        string,
  programId:        string,
  yearOfStudy:      number,
  academicYearName: string,
): Promise<CarryForwardResult> => {
  const student = await Student.findById(studentId).lean();
  if (!student) throw new Error("Student not found");

  const settings = await InstitutionSettings.findOne({ institution: (student as any).institution }).lean();
  const passMark = (settings as any)?.passMark ?? 40;

  // ENG.14a: No carry-forward to final year
  const programDoc = await mongoose.model("Program").findById(programId).lean() as any;
  const finalYear  = programDoc?.durationYears || 5;
  if (yearOfStudy >= finalYear) {
    return { granted: false, cfUnits: [], qualifier: "", reason: `ENG.14: No carry-forward to final year (Year ${finalYear}).` };
  }

  const programUnits = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy })
    .populate("unit").lean() as any[];

  const totalUnits = programUnits.length;
  const puIds      = programUnits.map((pu: any) => pu._id);

  const [detailedMarks, directMarks] = await Promise.all([
    Mark.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
    MarkDirect.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
  ]);

  const markMap = new Map<string, any>();
  [...detailedMarks, ...directMarks].forEach((m: any) => markMap.set(m.programUnit.toString(), m));

  const failedUnitCodes: string[]    = [];
  const noCAUnitCodes:   string[]    = [];
  const failedDetails:   Array<{ programUnitId: string; unitCode: string; unitName: string }> = [];

  for (const pu of programUnits) {
    const puId = (pu as any)._id.toString();
    const m    = markMap.get(puId);
    if (!m) continue;
    if ((m as any).isSpecial || (m as any).attempt === "special") continue;

    const mark  = (m as any).agreedMark ?? 0;
    const hasCA = ((m as any).caTotal30 ?? 0) > 0;

    if (mark < passMark) {
      const code = (pu as any).unit?.code || "N/A";
      failedUnitCodes.push(code);
      if (!hasCA) noCAUnitCodes.push(code); // ENG.15a: missing CA → cannot CF
      failedDetails.push({ programUnitId: puId, unitCode: code, unitName: (pu as any).unit?.name || "N/A" });
    }
  }

  const eligibility = assessCarryForwardEligibility(failedUnitCodes, noCAUnitCodes, totalUnits);
  if (!eligibility.eligible) return { granted: false, cfUnits: [], qualifier: "", reason: eligibility.reason };

  // Determine CF cycle number from existing qualifierSuffix
  const priorQualifier = (student as any).qualifierSuffix || "";
  const priorMatch     = priorQualifier.match(/RP(\d+)C/);
  const cfNumber       = priorMatch ? Math.min(parseInt(priorMatch[1]) + 1, 3) : 1;
  const qualifier      = REG_QUALIFIERS.carryForward(cfNumber);

  const cfUnits: CarryForwardUnit[] = eligibility.units.map((code: string) => {
    const detail = failedDetails.find((d) => d.unitCode === code);
    return {
      programUnitId:    detail?.programUnitId || "",
      unitCode:         code,
      unitName:         detail?.unitName || "N/A",
      fromYear:         yearOfStudy,
      fromAcademicYear: academicYearName,
      attemptNumber:    cfNumber + 2,
      qualifier,
      addedAt:          new Date(),
      status:           "pending" as const,
    };
  });

  await Student.findByIdAndUpdate(studentId, {
    $push: { carryForwardUnits: { $each: cfUnits } },
    $set:  { qualifierSuffix: qualifier },
  });

  return {
    granted:   true,
    cfUnits,
    qualifier,
    reason:    `ENG.14: Carry forward granted — ${cfUnits.length} unit(s): ${cfUnits.map((u) => u.unitCode).join(", ")}`,
  };
};

// ─── clearCarryForwardUnit ────────────────────────────────────────────────────
// Called from gradeCalculator when a CF unit is graded PASS.

export const clearCarryForwardUnit = async (
  studentId:     string,
  programUnitId: string,
): Promise<void> => {
  await Student.findByIdAndUpdate(studentId, {
    $pull: { carryForwardUnits: { programUnitId } },
  });

  const updated   = await Student.findById(studentId).select("carryForwardUnits").lean();
  const remaining = ((updated as any)?.carryForwardUnits || []).length;

  if (remaining === 0) {
    await Student.findByIdAndUpdate(studentId, { $set: { qualifierSuffix: "" } });
  }
};

// ─── getCarryForwardStudentsForUnit ──────────────────────────────────────────
// Used by scoresheetStudentList to include CF students on ORDINARY scoresheets.

export const getCarryForwardStudentsForUnit = async (
  programUnitId: string,
  programId:     string,
): Promise<Array<{ student: any; cfUnit: CarryForwardUnit; attemptLabel: string }>> => {
  const students = await Student.find({
    program:                              programId,
    "carryForwardUnits.programUnitId":    programUnitId,
    "carryForwardUnits.status":           "pending",
  }).lean() as any[];

  return students
    .map((student: any) => {
      const cfUnit = (student.carryForwardUnits as CarryForwardUnit[]).find(
        (u) => u.programUnitId === programUnitId && u.status === "pending",
      );
      if (!cfUnit) return null;
      return { student, cfUnit, attemptLabel: cfUnit.qualifier || "RP1C" };
    })
    .filter((item): item is NonNullable<typeof item> => item !== null);
};

// ─── getStayoutStudentsForUnit ────────────────────────────────────────────────
// ENG.15h: Stayout students retake in ORDINARY of NEXT year.

export const getStayoutStudentsForUnit = async (
  programUnitId: string,
  programId:     string,
): Promise<Array<{ student: any; attemptLabel: string }>> => {
  const pu = await ProgramUnit.findById(programUnitId).lean() as any;
  if (!pu) return [];

  const expectedYear = (pu.requiredYear || 1) + 1;

  const failedGrades = await FinalGrade.find({
    programUnit: programUnitId,
    status:      { $ne: "PASS" },
    attemptType: { $in: ["1ST_ATTEMPT", "SUPPLEMENTARY"] },
  }).populate("student").lean() as any[];

  const result: Array<{ student: any; attemptLabel: string }> = [];

  for (const grade of failedGrades) {
    const student = grade.student as any;
    if (!student)                                              continue;
    if (student.program?.toString() !== programId)             continue;
    if (student.currentYearOfStudy  !== expectedYear)          continue;
    if (student.status              !== "active")              continue;
    if ((student.qualifierSuffix || "").includes("C"))         continue; // CF students handled separately

    result.push({ student, attemptLabel: "A/SO" });
  }

  return result;
};