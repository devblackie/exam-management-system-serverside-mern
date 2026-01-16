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

// serverside/src/services/statusEngine.ts

export const previewPromotion = async (
  programId: string,
  yearToPromote: number,
  academicYearName: string
) => {
  const students = await Student.find({
    program: programId,
    currentYearOfStudy: yearToPromote,
    status: "active",
  }).lean();

  const eligible: any[] = [];
  const blocked: any[] = [];

  for (const student of students) {
    const statusResult = await calculateStudentStatus(
      student._id,
      programId,
      academicYearName,
      yearToPromote
    );

    const report = {
      id: student._id,
      regNo: student.regNo,
      name: student.name,
      status: statusResult?.status,
      summary: statusResult?.summary,
      reasons: [] as string[]
    };

    if (statusResult?.status === "IN GOOD STANDING") {
      eligible.push(report);
    } else {
      // Add specific reasons for the block
      if (statusResult?.missingList?.length) report.reasons.push(`${statusResult.missingList.length} units missing`);
      if (statusResult?.failedList?.length) report.reasons.push(`${statusResult.failedList.length} units failed`);
      if (statusResult?.reRetakeList?.length) report.reasons.push("Critical Re-retake failure");
      
      blocked.push(report);
    }
  }

  return {
    totalProcessed: students.length,
    eligibleCount: eligible.length,
    blockedCount: blocked.length,
    eligible,
    blocked
  };
};

export const promoteStudent = async (studentId: string) => {
  const student = await Student.findById(studentId);
  if (!student) throw new Error("Student not found");

  // 1. Get the current active academic year for context
  const currentYear = await AcademicYear.findOne({ isActive: true });
  const yearName = currentYear?.year || "2024/2025";

  // 2. Calculate status based on the student's CURRENT year of study
  const statusResult = await calculateStudentStatus(
    student._id,
    student.program,
    yearName,
    student.currentYearOfStudy || 1
  );

  // 3. Promotion Guard: Only "IN GOOD STANDING" can move up
  if (statusResult?.status === "IN GOOD STANDING") {
    const nextYear = (student.currentYearOfStudy || 1) + 1;

    // Check if the student has reached the maximum years for their program (Optional)
    // if (nextYear > student.programDuration) return { success: false, message: "Student has completed final year." };

    await Student.findByIdAndUpdate(studentId, {
      $set: {
        currentYearOfStudy: nextYear,
        currentSemester: 1, // Always reset to Semester 1 on promotion
      },
      // Log the promotion in history if you have a history field
      $push: { promotionHistory: { from: student.currentYearOfStudy, to: nextYear, date: new Date() } }
    });

    return { 
      success: true, 
      message: `Successfully promoted to Year ${nextYear}` 
    };
  }

  // 4. Detailed Failure Messages
  let failureReason = `Cannot promote: ${statusResult?.status}.`;
  if (statusResult?.status === "INCOMPLETE DATA") {
    failureReason += ` Missing marks for: ${statusResult.missingList.join(", ")}`;
  } else if (statusResult?.status === "SUPPLEMENTARY PENDING") {
    failureReason += ` Student must clear failed units first.`;
  }

  return {
    success: false,
    message: failureReason,
    details: statusResult
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
