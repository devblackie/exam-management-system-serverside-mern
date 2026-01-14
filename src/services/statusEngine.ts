// serverside/src/services/statusEngine.ts
import mongoose from "mongoose";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";

export const calculateStudentStatus = async (
  studentId: any,
  programId: any,
  academicYearName: string,
  yearOfStudy: number = 1
) => {
  const settings = await InstitutionSettings.findOne().lean();
  
  const rules = {
    passMark: settings?.passMark || 39.5,
    retakeLimit: settings?.retakeThreshold || 5 
  };
  
  const yearDoc = await AcademicYear.findOne({ year: academicYearName });
  if (!yearDoc) return null;

  // 1. Get Curriculum & explicitly type the populated unit
  const curriculum = await ProgramUnit.find({ 
    program: programId, 
    requiredYear: yearOfStudy 
  }).populate("unit").lean() as any[];

  // 2. Get Grades (All history to track Retake -> Re-Retake lifecycle)
  const grades = await FinalGrade.find({ student: studentId })
    .populate({ path: "programUnit", populate: { path: "unit" } })
    .sort({ createdAt: -1 }) 
    .lean() as any[];

  // 3. Map grades by UNIT CODE
  const unitResults = new Map<string, { status: string; attemptType: string; attemptNumber: number }>();

  grades.forEach((g) => {
        if (!g.programUnit || !g.programUnit.unit) {
      console.warn(
        `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`
      );
      return;
    }

    const unitCode = g.programUnit?.unit?.code?.toUpperCase();
    if (!unitCode) return;

    const existing = unitResults.get(unitCode);
    if (existing?.status === "PASS") return;

    unitResults.set(unitCode, { 
        status: g.status, 
        attemptType: g.attemptType, 
        attemptNumber: g.attemptNumber 
    });
  });

  // Trackers for the Coordinator View
  let passedCount = 0;
  const failedUnits: string[] = [];      // Standard Failed / Supp
  const retakeUnits: string[] = [];      // Retakes (Attempt 2)
  const reRetakeUnits: string[] = [];    // Re-Retakes (Attempt 3+)
  const missingUnits: string[] = [];     // Not attempted yet

  // 4. Compare Curriculum against results
  curriculum.forEach((pUnit: any) => {
    const unitCode = pUnit.unit?.code?.trim().toUpperCase();
    const unitName = pUnit.unit?.name;
    const displayName = `${unitCode}: ${unitName}`;

    const record = unitResults.get(unitCode);

    if (!record) {
      missingUnits.push(displayName);
    } else if (record.status === "PASS") {
      passedCount++;
    } else {
      // Categorize the failure by attempt type
      if (record.attemptType === "RE_RETAKE" || record.attemptNumber >= 3) {
        reRetakeUnits.push(displayName);
      } else if (record.attemptType === "RETAKE" || record.attemptNumber === 2) {
        retakeUnits.push(displayName);
      } else {
        failedUnits.push(displayName);
      }
    }
  });

  const totalFailed = failedUnits.length + retakeUnits.length + reRetakeUnits.length;
  const missingCount = missingUnits.length;

  // 5. Determine UI Status
  let status = "IN GOOD STANDING";
  let variant: "success" | "warning" | "error" | "info" = "success";

  if (missingCount > 0) {
    status = "INCOMPLETE DATA";
    variant = "info";
  } else if (reRetakeUnits.length > 0) {
    status = "RE-RETAKE FAILURE";
    variant = "error";
  } else if (totalFailed > rules.retakeLimit) {
    status = "RETAKE YEAR";
    variant = "error";
  } else if (totalFailed > 0) {
    status = "SUPPLEMENTARY PENDING";
    variant = "warning";
  }

  return {
    status,
    variant,
    details: missingCount > 0 
      ? `Missing marks for ${missingCount} units.`
      : reRetakeUnits.length > 0
      ? `Critical: Failed Re-Retake units.`
      : totalFailed > 0
      ? `Student has ${totalFailed} pending units.`
      : "All curriculum units cleared.",
    summary: { 
      totalExpected: curriculum.length, 
      passed: passedCount, 
      failed: totalFailed, 
      missing: missingCount 
    },
    // Detailed lists for Coordinator Search/View
    missingList: missingUnits,
    failedList: failedUnits,
    retakeList: retakeUnits,
    reRetakeList: reRetakeUnits
  };
};

export const promoteStudent = async (studentId: string) => {
  const student = await Student.findById(studentId);
  if (!student) throw new Error("Student not found");

  // Use the correct property: currentYearOfStudy
  const currentYear = await AcademicYear.findOne({ isActive: true });
  const statusResult = await calculateStudentStatus(
    student._id,
    student.program,
    currentYear?.year || "2024/2025",
    student.currentYearOfStudy || 1
  );

  if (statusResult?.status === "IN GOOD STANDING") {
    const nextYear = (student.currentYearOfStudy || 1) + 1;

    await Student.findByIdAndUpdate(studentId, {
      $set: {
        currentYearOfStudy: nextYear,
        // Optional: Reset semester to 1 upon promotion
        currentSemester: 1,
      },
    });

    return { success: true, message: `Promoted to Year ${nextYear}` };
  }

  return {
    success: false,
    message: `Cannot promote: Student is currently ${statusResult?.status}`,
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
    status: "active",
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
