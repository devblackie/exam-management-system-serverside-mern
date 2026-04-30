// serverside/src/services/academicAudit.ts
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import { calculateStudentStatus } from "./statusEngine";

/**
 * PRODUCTION-READY: ENG 22.b Automatic Discontinuation Logic
 * This handles both the 5th Attempt Rule and the Repeat Year Failure Rule.
 */
export interface AuditResult {
  discontinued: boolean;
  reason:       string;
}

// export const performAcademicAudit = async ( studentId: string, session?: mongoose.ClientSession ) => {
//   const student = await Student.findById(studentId).session(session || null);

export const performAcademicAudit = async (
  studentId: string,
  session?: mongoose.ClientSession,
): Promise<AuditResult> => {
  // ── Load student ───────────────────────────────────────────────────────────
  const query = Student.findById(studentId);
  if (session) query.session(session);
  const student = await query;
  if (!student) throw new Error("Student not found for audit.");

  // CHECK 1: ENG.22(b)(i) — 5th Attempt Rule
  //
  // A student is discontinued if they have a FinalGrade where:
  //   - attemptNumber >= 5  (this is their 5th attempt at this unit)
  //   - status !== "PASS"   (they failed it)
  //
  // The attemptNumber field is set by gradeCalculator.ts when computing
  // FinalGrade. It counts FinalGrade documents per (student, programUnit).
  //
  // Note: We look for >= 5 not === 5 because a corrupted import could
  // theoretically create 6 records. We catch anything at or beyond the limit.
  // ─────────────────────────────────────────────────────────────────────────
  const fgQuery = FinalGrade.findOne({
    student: studentId,
    attemptNumber: { $gte: 5 },
    status: { $ne: "PASS" },
  }).populate({ path: "programUnit", populate: { path: "unit" } });

  if (session) fgQuery.session(session);
  const fatalGrade = await fgQuery;

  if (fatalGrade) {
    const unitCode =
      (fatalGrade.programUnit as any)?.unit?.code || "Unknown Unit";
    const attemptN = fatalGrade.attemptNumber ?? 5;

    const reason = `ENG.22(b)(i): Failed unit ${unitCode} at attempt ${attemptN}. Discontinued.`;

    await Student.findByIdAndUpdate(
      studentId,
      {
        $set: { status: "discontinued", remarks: reason },
        $push: {
          statusHistory: {
            status: "discontinued",
            previousStatus: (student as any).status,
            date: new Date(),
            reason,
          },
        },
      },
      session ? { session } : {},
    );

    return { discontinued: true, reason };
  }

  // CHECK 2: ENG.22(b)(ii) — Repeat Year Failure Rule
  //
  // A student who was ALREADY in a repeat year (ENG.16) and fails again
  // (mean < 40% OR ≥ 50% units failed) is discontinued.
  //
  // FIX from original: status should be "repeat", not "active".
  // A student actively repeating a year has status === "repeat".
  // Checking status === "active" means this rule NEVER fired.
  //
  // We also check academicHistory to confirm they are genuinely in a
  // repeat year for the current year of study, not just that their
  // status field is "repeat" from a previous year.
  // ─────────────────────────────────────────────────────────────────────────
  const isInRepeatYear =
    (student as any).status === "repeat" &&
    ((student as any).academicHistory || []).some(
      (h: any) =>
        h.yearOfStudy === (student as any).currentYearOfStudy && h.isRepeatYear,
    );

  if (isInRepeatYear) {
    const performance = await calculateStudentStatus(
      studentId,
      (student as any).program.toString(),
      "CURRENT", // FIX: was "N/A" — "CURRENT" explicitly resolves to current year
      (student as any).currentYearOfStudy,
      { forPromotion: true },
    );

    if (performance.status === "REPEAT YEAR") {
      const reason =
        `ENG.22(b)(ii): Failed to clear repeat year requirements ` +
        `(Mean ${performance.weightedMean}% — ${performance.summary.failed}/${performance.summary.totalExpected} ` +
        `units failed). Discontinued.`;

      await Student.findByIdAndUpdate(
        studentId,
        {
          $set: { status: "discontinued", remarks: reason },
          $push: {
            statusHistory: {
              status: "discontinued",
              previousStatus: "repeat",
              date: new Date(),
              reason,
            },
          },
        },
        session ? { session } : {},
      );

      return { discontinued: true, reason };
    }
  }

  return { discontinued: false, reason: "" };
};
