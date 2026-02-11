// serverside/src/services/statusEngine.ts
import mongoose from "mongoose";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";

const getLetterGrade = (mark: number, settings: any): string => {
  if (!settings || !settings.gradingScale) {
    // Fallback if settings are missing
    if (mark >= 69.5) return "A";
    if (mark >= 59.5) return "B";
    if (mark >= 49.5) return "C";
    if (mark >= 39.5) return "D";
    return "E";
  }

  // Sort scale descending (e.g., 70, 60, 50...) to find the first match
  const sortedScale = [...settings.gradingScale].sort((a, b) => b.min - a.min);
  const matched = sortedScale.find((s) => mark >= s.min);
  return matched ? matched.grade : settings.failingGrade || "E";
};

export const calculateStudentStatus = async (
  studentId: any,
  programId: any,
  academicYearName: string,
  yearOfStudy: number = 1
) => {
  const settings = await InstitutionSettings.findOne().lean();
  if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");

  let displayYearName = academicYearName;
 if (!displayYearName || displayYearName === "N/A") {
    const latestGrade = await FinalGrade.findOne({ student: studentId })
      .populate("academicYear")
      .sort({ createdAt: -1 });
    
    // Changed .year to .name
    displayYearName = (latestGrade?.academicYear as any)?.year || "N/A";
  }

  const rules = {
    passMark: settings.passMark,
    retakeLimit: settings.retakeThreshold || 5 
  };
  
  // const yearDoc = await AcademicYear.findOne({ year: academicYearName });
  // if (!yearDoc) return null;

  // 1. Get Curriculum & explicitly type the populated unit
  const curriculum = await ProgramUnit.find({ 
    program: programId, 
    requiredYear: yearOfStudy 
  }).populate("unit").lean() as any[];

  if (!curriculum.length) {
    return {
        status: "NO CURRICULUM",
        variant: "info",
        details: `No units defined for Year ${yearOfStudy} in this program.`,
        summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
        passedList: []
    };
  }


  // 2. Get Grades (All history to track Retake -> Re-Retake lifecycle)
  const grades = await FinalGrade.find({ student: studentId })
    .populate({ path: "programUnit", populate: { path: "unit" } })
    .populate({ 
        path: "academicYear", 
        model: "AcademicYear" 
    })
    .sort({ createdAt: -1 }) 
    .lean() as any[];

//     console.log(`DEBUG: Found ${grades.length} grades for student.`);
// if (grades.length > 0) {
//     console.log("Sample Grade AcademicYear Data:", JSON.stringify(grades[0].academicYear, null, 2));
// }

// console.log("RE-CHECK Sample Grade AcademicYear Data:", grades[0]?.academicYear);

  // 3. Map grades by UNIT CODE
//  const unitResults = new Map<string, any>();
const unitResults = new Map<string, { 
  status: string; 
  attemptType: string; 
  attemptNumber: number;
  totalMark: number; // Changed from finalMark to match Model
}>();

  grades.forEach((g) => {
        if (!g.programUnit || !g.programUnit.unit) {
      console.warn(
        `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`
      );
      return;
    }

    const unitCode = g.programUnit.unit.code.toUpperCase();
    // const unitCode = g.programUnit?.unit?.code?.toUpperCase();
    if (!unitCode) return;

    const existing = unitResults.get(unitCode);
    if (existing?.status === "PASS") return;

 unitResults.set(unitCode, { 
        status: g.status, 
        attemptType: g.attemptType, 
        attemptNumber: g.attemptNumber,
        totalMark: g.totalMark 
    });
  });

  // Trackers for the Coordinator View
  let passedCount = 0;
  const passedUnits: any[] = [];
  const failedUnits: string[] = [];      
  const retakeUnits: string[] = [];      
  const reRetakeUnits: string[] = [];    
  const missingUnits: string[] = [];     
  const specialExamUnits: string[] = []; // NEW: For Tony
  const incompleteUnits: string[] = [];

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

      const numericMark = record.totalMark || 0;
      // const letterGrade = getLetterGrade(numericMark);
      const letterGrade = getLetterGrade(numericMark, settings);
      
      passedUnits.push({
        code: unitCode,
        name: unitName,
        mark: numericMark,            
        grade: letterGrade            
      });
    } else if (record.status === "SPECIAL") {
      specialExamUnits.push(displayName);
    } else if (record.status === "INCOMPLETE") {
      incompleteUnits.push(displayName);
    } else {
      // Logic for failures based on attempts
      if (record.attemptNumber >= 3) reRetakeUnits.push(displayName);
      else if (record.attemptNumber === 2) retakeUnits.push(displayName);
      else failedUnits.push(displayName);
    }
  });

  const totalFailed = failedUnits.length + retakeUnits.length + reRetakeUnits.length;

  // 5. Determine UI Status
  let status = "IN GOOD STANDING";
  let variant: "success" | "warning" | "error" | "info" = "success";
  let details = `Year ${yearOfStudy} curriculum units cleared.`;

if (reRetakeUnits.length > 0) {
    status = "RE-RETAKE FAILURE";
    variant = "error";
    details = "Critical: Failed third attempt at units.";
  } else if (totalFailed > rules.retakeLimit) {
    status = "RETAKE YEAR";
    variant = "error";
    details = `Failed Year ${yearOfStudy} units (${totalFailed}) exceed limit.`;
  } else if (specialExamUnits.length > 0) {
    status = "SPECIAL EXAM PENDING";
    variant = "info";
    details = `Pending ${specialExamUnits.length} special exam(s) for Year ${yearOfStudy}.`;
  } else if (incompleteUnits.length > 0) {
    status = "INCOMPLETE MARKS";
    variant = "warning";
    details = `Incomplete components in Year ${yearOfStudy}.`;
  } else if (totalFailed > 0) {
    status = "SUPPLEMENTARY PENDING";
    variant = "warning";
    details = `Must sit supplementary exams for ${totalFailed} Year ${yearOfStudy} unit(s).`;
  } else if (missingUnits.length > 0) {
    status = "INCOMPLETE DATA";
    variant = "info";
    details = `No record found for ${missingUnits.length} units in Year ${yearOfStudy}.`;
  }


const sessionRecord = grades.find(g => g.programUnit?.requiredYear === yearOfStudy);

 let actualSessionName = "N/A";

if (academicYearName && academicYearName !== "N/A") {
    actualSessionName = academicYearName;
} else {
    // 1. Try to find a grade in the CURRENT year of study that has a valid year
    const yearSpecificGrade = grades.find(g => 
        g.programUnit?.requiredYear === yearOfStudy && 
        g.academicYear?.year
    );

    if (yearSpecificGrade?.academicYear?.year) {
        actualSessionName = yearSpecificGrade.academicYear.year;
    } else if (grades.length > 0) {
        // 2. If Year 2 is empty, but Year 1 exists, we "guess" the next year
        const previousYearGrade = grades.find(g => g.academicYear?.year);
        if (previousYearGrade?.academicYear?.year) {
            const baseYear = previousYearGrade.academicYear.year; // e.g., "2023/2024"
            // Simple logic to increment if we are looking at a higher year of study
            if (yearOfStudy > (previousYearGrade.programUnit?.requiredYear || 1)) {
                 // Optional: Add logic to increment "2023/2024" to "2024/2025"
                 actualSessionName = baseYear; // For now, keep as base to avoid wrong guesses
            } else {
                actualSessionName = baseYear;
            }
        }
    }
}
  

 return {
    status, variant, details,
    // academicYear: academicYearName, 
    academicYearName: actualSessionName,
    yearOfStudy: yearOfStudy,
    summary: { 
        totalExpected: curriculum.length, 
        passed: passedCount, 
        failed: totalFailed, 
        missing: missingUnits.length 
    },
    missingList: missingUnits,
    passedList: passedUnits,
    failedList: failedUnits,
    retakeList: retakeUnits,
    reRetakeList: reRetakeUnits,
    specialList: specialExamUnits,
    incompleteList: incompleteUnits
  };
};

export const previewPromotion = async (
  programId: string,
  yearToPromote: number,
  academicYearName: string
) => {
 // 1. Fetch students in the targeted year AND those already in the next year
  const nextYear = yearToPromote + 1;
  const students = await Student.find({
    program: programId,
    currentYearOfStudy: { $in: [yearToPromote, nextYear] },
    status: "active",
  }).lean();

  const eligible: any[] = [];
  const blocked: any[] = [];

  for (const student of students) {
    const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
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
      status: isAlreadyPromoted ? "ALREADY PROMOTED" : (statusResult?.status || "PENDING"),
      summary: statusResult?.summary,
      reasons: [] as string[],
      isAlreadyPromoted
    };

    if (isAlreadyPromoted || statusResult?.status === "IN GOOD STANDING") {
      eligible.push(report);
    } else {
      // Add specific reasons for the block
      if (statusResult?.specialList?.length) {
        report.reasons.push(...statusResult.specialList.map(u => `${u} - SPECIAL`));
      }
      if (statusResult?.incompleteList?.length) {
        report.reasons.push(...statusResult.incompleteList.map(u => `${u} - INCOMPLETE`));
      }
      if (statusResult?.missingList?.length) {
        report.reasons.push(...statusResult.missingList.map(u => `MISSING: ${u}`));
      }
      if (statusResult?.failedList?.length) {
        report.reasons.push(...statusResult.failedList.map(u => `FAILED: ${u}`));
      }
      if (statusResult?.retakeList?.length) {
        report.reasons.push(...statusResult.retakeList.map(u => `RETAKE FAILED: ${u}`));
      }
      if (statusResult?.reRetakeList?.length) {
        report.reasons.push(...statusResult.reRetakeList.map(u => `CRITICAL RE-RETAKE FAILED: ${u}`));
      }
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
  const actualCurrentYear = student.currentYearOfStudy || 1;

  // 2. Calculate status based on the student's CURRENT year of study
  const statusResult = await calculateStudentStatus(
    student._id,
    student.program,
    // "",
    "N/A",
    actualCurrentYear,
  );

  // 3. Promotion Guard: Only "IN GOOD STANDING" can move up
  if (statusResult?.status === "IN GOOD STANDING") {
const nextYear = actualCurrentYear + 1;

    // Check if the student has reached the maximum years for their program (Optional)
    // if (nextYear > student.programDuration) return { success: false, message: "Student has completed final year." };

    await Student.findByIdAndUpdate(studentId, {
      $set: {
        currentYearOfStudy: nextYear,
        currentSemester: 1, // Always reset to Semester 1 on promotion
      },
      // Log the promotion in history if you have a history field
      $push: { 
        promotionHistory: { 
          from: actualCurrentYear, 
          to: nextYear, 
          date: new Date() 
        } 
      }
    });

    return { 
      success: true, 
      message: `Successfully promoted to Year ${nextYear}` 
    };
  }

   return {
    success: false,
    message: `Promotion Blocked: Student is '${statusResult?.status}' for Year ${actualCurrentYear}.`,
    details: statusResult
  };
};

export const bulkPromoteClass = async (
  programId: string,
  yearToPromote: number,
  academicYearName: string
) => {
  // 1. Get everyone currently in the year AND those already promoted to the next year
  const nextYear = yearToPromote + 1;
  const students = await Student.find({
    program: programId,
    currentYearOfStudy: { $in: [yearToPromote, nextYear] },
    status: "active",
  });

  const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };

  for (const student of students) {
    try {
      // Cast to 'any' or use (student as any)._id to bypass the 'unknown' check
      const studentId = (student._id as any).toString();

      // Skip if already in the target year or higher
      if (student.currentYearOfStudy >= nextYear) {
        results.alreadyPromoted++;
        results.promoted++; // Still count them in the success total for the UI
        continue;
      }
      
      const res = await promoteStudent(studentId);

      if (res.success) results.promoted++;
      else results.failed++;
    } catch (err: any) {
      results.errors.push(`${student.regNo}: ${err.message}`);
    }
  }

  return results;
};

// add a "Transcript History" table so you can see a list of all previously generated PDFs for this student
