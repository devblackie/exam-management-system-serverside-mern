import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import { calculateStudentStatus } from "./statusEngine";

/**
 * PRODUCTION-READY: ENG 22.b Automatic Discontinuation Logic
 * This handles both the 5th Attempt Rule and the Repeat Year Failure Rule.
 */
export const performAcademicAudit = async (
  studentId: string,
  session?: mongoose.ClientSession,
) => {
  const student = await Student.findById(studentId).session(session || null);
  if (!student) throw new Error("Student not found for audit.");

  // --- 1. THE 5th ATTEMPT RULE (ENG 22.b.i) ---
  // Find the specific grade that triggered this
  const fatalUnitFailure = await FinalGrade.findOne({
    student: studentId,
    attemptNumber: { $gte: 5 },
    status: { $ne: "PASS" },
  })
    .populate({ path: "programUnit", populate: { path: "unit" } })
    .session(session || null);

  if (fatalUnitFailure) {
    const unitInfo =
      (fatalUnitFailure.programUnit as any)?.unit?.code || "Unknown Unit";
    await Student.findByIdAndUpdate(
      studentId,
      {
        $set: {
          status: "discontinued",
          remarks: `ENG 22.b.i Violation: Failed unit ${unitInfo} on attempt #${fatalUnitFailure.attemptNumber}.`,
        },
      },
      { session },
    );
    return { discontinued: true, reason: "5th Attempt Failure" };
  }

  // --- 2. THE REPEAT YEAR FAILURE RULE (ENG 22.b.ii / 16.b) ---
  // If a student is already repeating, and fails the 50% rule again, they are out.
  if (
    student.status === "active" &&
    student.academicHistory?.some(
      (h) => h.yearOfStudy === student.currentYearOfStudy && h.isRepeatYear,
    )
  ) {
    // We call the status engine to get the current performance for the year
    const performance = await calculateStudentStatus(
      studentId.toString(),
      student.program.toString(),
      "N/A", // Use current session logic inside statusEngine
      student.currentYearOfStudy,
    );

    if (performance.status === "REPEAT YEAR") {
      await Student.findByIdAndUpdate(
        studentId,
        {
          $set: {
            status: "discontinued",
            remarks: `ENG 22.b.ii Violation: Failed to clear Repeat Year requirements (Mean < 40% or > 50% units failed).`,
          },
        },
        { session },
      );
      return { discontinued: true, reason: "Failed Repeat Year" };
    }
  }

  return { discontinued: false };
};
