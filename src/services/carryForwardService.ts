// src/services/carryForwardService.ts
import mongoose from "mongoose";
import Student from "../models/Student";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import ProgramUnit from "../models/ProgramUnit";
import InstitutionSettings from "../models/InstitutionSettings";
import {
  assessCarryForwardEligibility,
  REG_QUALIFIERS,
} from "../utils/academicRules";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface CFUnit {
  programUnitId: string;
  unitCode: string;
  unitName: string;
}

export interface CarryForwardResult {
  eligible: boolean;
  promoted: boolean;
  units: CFUnit[];
  qualifier: string;
  reason: string;
}

// ─── assessCarryForward ───────────────────────────────────────────────────────
// Determines if a student qualifies to carry forward failed units under ENG.14.
// Called AFTER supplementary results are uploaded (session = SUPPLEMENTARY).
//
// Logic:
//   - Collect all units in the prescribed curriculum for this year
//   - Find which are failed at the supplementary stage
//   - Exclude units failed because of missing CA (ENG.15a)
//   - Allow max 2 carry-forward units (ENG.14a)
//   - Block if this is the final year (ENG.14a — no CF to Year 5/4)

export const assessCarryForward = async (
  studentId: string,
  programId: string,
  academicYearName: string,
  yearOfStudy: number,
): Promise<CarryForwardResult> => {
  const student = await Student.findById(studentId).lean();
  if (!student) throw new Error(`Student ${studentId} not found`);

  const settings = await InstitutionSettings.findOne({
    institution: (student as any).institution,
  }).lean();
  const passMark = settings?.passMark ?? 40;

  // Check final year restriction
  const programDoc = (await mongoose
    .model("Program")
    .findById(programId)
    .lean()) as any;
  const finalYear = programDoc?.durationYears || 5;
  if (yearOfStudy >= finalYear) {
    return {
      eligible: false,
      promoted: false,
      units: [],
      qualifier: "",
      reason: `ENG.14: Carry-forward not permitted to final year (Year ${finalYear}).`,
    };
  }

  const programUnits = await ProgramUnit.find({
    program: programId,
    requiredYear: yearOfStudy,
  })
    .populate("unit")
    .lean();

  const totalUnits = programUnits.length;
  const puIds = programUnits.map((pu: any) => pu._id);

  const [detailedMarks, directMarks] = await Promise.all([
    Mark.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
    MarkDirect.find({ student: studentId, programUnit: { $in: puIds } }).lean(),
  ]);

  const markMap = new Map<string, any>();
  [...detailedMarks, ...directMarks].forEach((m: any) => {
    markMap.set(m.programUnit.toString(), m);
  });

  const failedUnitIds: string[] = [];
  const noCAUnitIds: string[] = [];
  const failedDetails: CFUnit[] = [];

  for (const pu of programUnits) {
    const puId = (pu as any)._id.toString();
    const m = markMap.get(puId);
    if (!m) continue;

    const isSpecial = m.isSpecial || m.attempt === "special";
    if (isSpecial) continue;

    const mark = m.agreedMark ?? 0;
    const hasCA = (m.caTotal30 ?? 0) > 0;

    if (mark < passMark) {
      failedUnitIds.push(puId);
      if (!hasCA) noCAUnitIds.push(puId); // ENG.15a: missing CA → cannot CF
      failedDetails.push({
        programUnitId: puId,
        unitCode: (pu as any).unit?.code || "N/A",
        unitName: (pu as any).unit?.name || "N/A",
      });
    }
  }

  const eligibility = assessCarryForwardEligibility(
    failedUnitIds,
    noCAUnitIds,
    totalUnits,
  );

  return {
    eligible: eligibility.eligible,
    promoted: eligibility.eligible,
    units: eligibility.units.map(
      (id) =>
        failedDetails.find((d) => d.programUnitId === id) || {
          programUnitId: id,
          unitCode: "N/A",
          unitName: "N/A",
        },
    ),
    qualifier: eligibility.eligible ? REG_QUALIFIERS.carryForward(1) : "",
    reason: eligibility.reason,
  };
};

// ─── applyCarryForward ────────────────────────────────────────────────────────
// Promotes the student to the next year and records the CF units on their record.

export const applyCarryForward = async (
  studentId: string,
  programId: string,
  academicYearName: string,
  yearOfStudy: number,
  cfUnits: CFUnit[],
): Promise<void> => {
  const nextYear = yearOfStudy + 1;

  const existing = (await Student.findById(studentId)
    .select("qualifierSuffix")
    .lean()) as any;
  const priorQualifier = existing?.qualifierSuffix || "";

  const priorCF = (priorQualifier.match(/RP(\d+)C/) ?? [])[1];
  const cfCount = priorCF ? parseInt(priorCF) + 1 : 1;
  const newQualifier = REG_QUALIFIERS.carryForward(cfCount);

  const entries = cfUnits.map((u) => ({
    programUnitId: new mongoose.Types.ObjectId(u.programUnitId),
    unitCode: u.unitCode,
    unitName: u.unitName,
    fromYear: yearOfStudy,
    fromAcademicYear: academicYearName,
    attemptCount: cfCount,
    status: "pending" as const,
  }));

  await Student.findByIdAndUpdate(studentId, {
    $set: {
      currentYearOfStudy: nextYear,
      currentSemester: 1,
      qualifierSuffix: newQualifier,
    },
    $push: {
      carryForwardUnits: { $each: entries },
      statusEvents: {
        fromStatus: `year_${yearOfStudy}`,
        toStatus: `year_${nextYear}_cf`,
        date: new Date(),
        academicYear: academicYearName,
        reason: `ENG.14 Carry Forward to Year ${nextYear}. Units: ${cfUnits.map((u) => u.unitCode).join(", ")}. Qualifier: ${newQualifier}`,
      },
      statusHistory: {
        status: "active",
        previousStatus: "active",
        date: new Date(),
        reason: `Carry-forward promotion to Year ${nextYear} (${newQualifier})`,
      },
    },
  });
};

// ─── resolveCarryForwardUnit ──────────────────────────────────────────────────
// Updates one CF unit entry after results are processed.

export const resolveCarryForwardUnit = async (
  studentId: string,
  programUnitId: string,
  outcome: "passed" | "failed" | "escalated_to_rpu",
): Promise<void> => {
  await Student.updateOne(
    {
      _id: studentId,
      "carryForwardUnits.programUnitId": new mongoose.Types.ObjectId(
        programUnitId,
      ),
    },
    { $set: { "carryForwardUnits.$.status": outcome } },
  );

  if (outcome === "escalated_to_rpu") {
    const s = (await Student.findById(studentId)
      .select("qualifierSuffix")
      .lean()) as any;
    const prior = s?.qualifierSuffix || "";
    const rpu = (prior.match(/RPU(\d+)/) ?? [])[1];
    const rpuN = rpu ? parseInt(rpu) + 1 : 1;

    await Student.findByIdAndUpdate(studentId, {
      $set: { qualifierSuffix: `RPU${rpuN}` },
      $push: {
        statusEvents: {
          fromStatus: "carry_forward",
          toStatus: "repeat_unit",
          date: new Date(),
          academicYear: "CURRENT",
          reason: `ENG.16b: Failed CF supplementary. Must repeat unit (RPU${rpuN}).`,
        },
      },
    });
  }

  // If all pending CF units are resolved and all passed, clear the qualifier
  const student = (await Student.findById(studentId).lean()) as any;
  const cfUnits = student?.carryForwardUnits || [];
  const allResolved = cfUnits.every((u: any) => u.status !== "pending");
  const allPassed = cfUnits.every((u: any) => u.status === "passed");

  if (allResolved && allPassed) {
    await Student.findByIdAndUpdate(studentId, {
      $set: { qualifierSuffix: "" },
    });
  }
};
