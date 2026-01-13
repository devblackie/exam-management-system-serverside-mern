// serverside/src/services/statusEngine.ts
import mongoose from "mongoose";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import Student, { IStudent } from "../models/Student";

export const calculateStudentStatus = async (
  studentId: any,
  programId: any,
  academicYearName: string, // Changed to string for easier lookup
  yearOfStudy: number = 1
) => {
  // 1. Find the year document first to get its ID
  const yearDoc = await AcademicYear.findOne({ year: academicYearName });
  if (!yearDoc) return null;

  // 2. Get Curriculum
  const curriculum = await ProgramUnit.find({
    program: programId,
    requiredYear: yearOfStudy
  }).populate("unit").lean();

  // 3. Get Grades for this specific year
  const grades = await FinalGrade.find({
    student: studentId,
    academicYear: yearDoc._id
  }).lean();

// This handles cases where a student has a FAIL and a SUPP PASS for the same unit
  const unitResults = new Map<string, string>(); 
  grades.forEach(g => {

    if (!g.programUnit) {
    console.warn(`[StatusEngine] Skipping grade record ${g._id} - missing programUnit`);
    return;
  }
  
    const pUnitId = g.programUnit.toString();
    const existingStatus = unitResults.get(pUnitId);

    // If we already have a PASS, don't overwrite it with a FAIL/SUPP
    if (existingStatus === "PASS") return;
    
    unitResults.set(pUnitId, g.status);
  });

  let passed = 0;
  let failed = 0;
  let missing = 0;
  const failedUnits: string[] = [];

  curriculum.forEach((pUnit: any) => {
    const gradeStatus = unitResults.get(pUnit._id.toString());
    if (!gradeStatus) {
      missing++;
    } else if (gradeStatus === "PASS") {
      passed++;
    } else {
      failed++;
      failedUnits.push(pUnit.unit?.code || "Unknown");
    }
  });

  // 4. Determine Status String
  let status = "IN GOOD STANDING";
  let variant: "success" | "warning" | "error" | "info" = "success";

  if (failed > 0 && failed <= 3) {
    status = "SUPPLEMENTARY PENDING";
    variant = "warning";
  } else if (failed > 3) {
    status = "RETAKE YEAR";
    variant = "error";
  } else if (missing > 0) {
    status = "INCOMPLETE DATA";
    variant = "info";
  }

return {
  status,
  variant,
  details: missing > 0 
    ? `Student is missing marks for ${missing} units. Please upload all results.`
    : failedUnits.length > 0 
      ? `Student must sit for supplementaries in: ${failedUnits.join(", ")}` 
      : "Student has cleared all units for this academic year.",
  summary: { totalExpected: curriculum.length, passed, failed, missing }
};
};

export const promoteStudent = async (studentId: string) => {
  const student = await Student.findById(studentId);
  if (!student) throw new Error("Student not found");

  // Use the correct property: currentYearOfStudy
  const statusResult = await calculateStudentStatus(
    student._id,
    student.program,
    "2024/2025", 
    student.currentYearOfStudy || 1 
  );

  if (statusResult?.status === "IN GOOD STANDING") {
    const nextYear = (student.currentYearOfStudy || 1) + 1;
    
    await Student.findByIdAndUpdate(studentId, {
      $set: { 
        currentYearOfStudy: nextYear,
        // Optional: Reset semester to 1 upon promotion
        currentSemester: 1 
      }
    });

    return { success: true, message: `Promoted to Year ${nextYear}` };
  }

  return { 
    success: false, 
    message: `Cannot promote: Student is currently ${statusResult?.status}` 
  };
};

export const bulkPromoteClass = async (
  programId: string, 
  yearToPromote: number, 
  academicYearName: string
) => {
  // Use the IStudent interface to help TypeScript understand the results
  const students = await Student.find({ 
    program: programId, 
    currentYearOfStudy: yearToPromote,
    status: "active" 
  });

  const results = { promoted: 0, failed: 0, errors: [] as string[] };

  for (const student of students) {
    try {
      // Cast to 'any' or use (student as any)._id to bypass the 'unknown' check
      const studentId = (student._id as any).toString();
      
      const res = await promoteStudent(studentId);
      
      if (res.success) results.promoted++;
      else results.failed++;
    } catch (err: any) {
      results.errors.push(`${student.regNo}: ${err.message}`);
    }
  }

  return results;
};