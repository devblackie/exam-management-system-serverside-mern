// serverside/src/services/statusEngine.ts
import mongoose from "mongoose";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";
import { getYearWeight } from "../utils/weightingRegistry";

const getLetterGrade = (mark: number, settings: any): string => {
  if (!settings || !settings.gradingScale) {
    // Fallback if settings are missing
    if (mark >= 69.5) return "A"; if (mark >= 59.5) return "B"; if (mark >= 49.5) return "C"; if (mark >= 39.5) return "D"; return "E";
  }

  // Sort scale descending (e.g., 70, 60, 50...) to find the first match
  const sortedScale = [...settings.gradingScale].sort((a, b) => b.min - a.min);
  const matched = sortedScale.find((s) => mark >= s.min);
  return matched ? matched.grade : settings.failingGrade || "E";
};

export interface StudentStatusResult {
  status: string;
  variant: "success" | "warning" | "error" | "info";
  details: string;
  weightedMean: string;
  summary: { totalExpected: number; passed: number; failed: number; missing: number; };
  passedList: { code: string; mark: number }[];
  failedList: { displayName: string; attempt: number }[];
  specialList: { displayName: string; grounds: string }[];
  missingList: string[];
  incompleteList: string[];
}

export const calculateStudentStatus = async (studentId: any, programId: any, academicYearName: string, yearOfStudy: number = 1): Promise<StudentStatusResult> => {
  const settings = await InstitutionSettings.findOne().lean();
  if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");
  const passMark = settings?.passMark || 40;

  const curriculum = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean() as any[];
    
  const grades = await FinalGrade.find({ student: studentId }).populate({ path: "programUnit", populate: { path: "unit" } }).lean() as any[];

  const unitResults = new Map();
  grades.sort((a, b) => new Date(a.updatedAt).getTime() - new Date(b.updatedAt).getTime());

  grades.forEach((g) => {
    if (!g.programUnit || !g.programUnit.unit) { console.warn( `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`, ); return; }
    const unitCode = g.programUnit?.unit?.code?.toUpperCase();
    if (!unitCode) return;
    const numericMark = g.agreedMark ?? g.totalMark ?? 0;
    
    unitResults.set(unitCode, {
      mark: Number(numericMark),
      status: g.status,
      attempt: parseInt(g.attempt) || g.attemptNumber || 1,
      isSpecial: g.isSpecial || g.status === "SPECIAL" || g.remarks?.toLowerCase().includes("financial") || g.remarks?.toLowerCase().includes("compassionate"),
      remarks: g.remarks || ""
    });
  });

  const lists = { 
    passed: [] as { code: string; mark: number }[], 
    failed: [] as { displayName: string; attempt: number }[], 
    special: [] as { displayName: string; grounds: string }[], 
    missing: [] as string[], incomplete: [] as string[] 
  };
  
  let totalFirstAttemptSum = 0;

  curriculum.forEach((pUnit) => {
    const code = pUnit.unit?.code?.toUpperCase();
    const displayName = `${code}: ${pUnit.unit?.name}`;
    const record = unitResults.get(code);

    if (!record) {
      lists.missing.push(displayName);
    } else if (record.isSpecial) {
      lists.special.push({ displayName, grounds: record.remarks || "Special Grounds" });
    } else if (record.mark === 0 || record.status === "INCOMPLETE") {
      lists.incomplete.push(displayName);
    } else if (record.mark >= passMark) {
      if (record.attempt === 1) totalFirstAttemptSum += record.mark;
      lists.passed.push({ code, mark: record.mark });
    } else {
      if (record.attempt === 1) totalFirstAttemptSum += record.mark;
      lists.failed.push({ displayName, attempt: record.attempt });
    }
  });

  const totalUnits = curriculum.length;
  const failCount = lists.failed.length;
  const meanMark = totalUnits > 0 ? totalFirstAttemptSum / totalUnits : 0;

  let status = "IN GOOD STANDING";
  let variant: "success" | "warning" | "error" | "info" = "success";
  let details = "Proceed to next year.";

  if (lists.missing.length >= 6) {
    status = "DEREGISTERED"; variant = "error"; details = "Absent from 6+ examinations (ENG 23c).";
  } else if (failCount >= totalUnits / 2 || meanMark < 40) {
    status = "REPEAT YEAR"; variant = "error"; details = "Failed >= 50% units or Mean < 40% (ENG 16).";
  } else if (failCount > totalUnits / 3) {
    status = "STAYOUT"; variant = "warning"; details = "Failed > 1/3 of units. Retake units next year (ENG 15h).";
  } else if (failCount > 0) {
    status = "SUPPLEMENTARY"; variant = "warning"; details = "Eligible for supplementaries (ENG 13a).";
  } else if (lists.special.length > 0) {
    status = "SPECIALS PENDING"; variant = "info"; details = "Awaiting special examinations.";
  }

  return {
    status, variant, details,
    weightedMean: meanMark.toFixed(2),
    summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: lists.missing.length },
    passedList: lists.passed, failedList: lists.failed, specialList: lists.special, missingList: lists.missing, incompleteList: lists.incomplete
  };
};

export const previewPromotion = async (programId: string, yearToPromote: number, academicYearName: string) => {
  const nextYear = yearToPromote + 1;
  const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active" }).lean();

  const eligible: any[] = [];
  const blocked: any[] = [];

  for (const student of students) {
    const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
    const statusResult = await calculateStudentStatus(student._id, programId, academicYearName, yearToPromote);

    const report = {
      id: student._id, regNo: student.regNo, name: student.name,
      status: isAlreadyPromoted ? "ALREADY PROMOTED" : statusResult.status,
      summary: statusResult.summary, reasons: [] as string[], isAlreadyPromoted
    };

    // Promotion criteria: Must be In Good Standing AND not already moved
    if (!isAlreadyPromoted && statusResult.status === "IN GOOD STANDING") {
      eligible.push(report);
    } else if (isAlreadyPromoted) {
      eligible.push(report); // Keep already promoted in eligible list but marked accordingly
    } else {
      // Mapping the new lists to reasons
      if (statusResult.specialList.length) report.reasons.push(...statusResult.specialList.map(s => `${s.displayName} (SPECIAL)`));
      if (statusResult.incompleteList.length) report.reasons.push(...statusResult.incompleteList.map(u => `${u} (INCOMPLETE)`));
      if (statusResult.missingList.length) report.reasons.push(...statusResult.missingList.map(u => `${u} (MISSING)`));
      if (statusResult.failedList.length) report.reasons.push(...statusResult.failedList.map(f => `${f.displayName} (FAIL ATTEMPT: ${f.attempt})`));
      
      blocked.push(report);
    }
  }

  return { totalProcessed: students.length, eligibleCount: eligible.length, blockedCount: blocked.length, eligible, blocked };
};

export const promoteStudent = async (studentId: string) => {
  const student = await Student.findById(studentId);
  if (!student) throw new Error("Student not found");

  const program = student.program as any;
  const actualCurrentYear = student.currentYearOfStudy || 1;
  const statusResult = await calculateStudentStatus( student._id, student.program, "N/A", actualCurrentYear );

  // REGULATION ENG 15 (b/c): Must have passed ALL units to register for the final year(s)
  // REGULATION ENG 18: Specials must be cleared before progression in most Engineering tracks
  
  if (statusResult?.status === "IN GOOD STANDING") {
    const nextYear = actualCurrentYear + 1;

    const yearWeight = getYearWeight(program.durationYears, student.entryType || "Direct", actualCurrentYear);

    const rawMean = parseFloat(statusResult.weightedMean);
    const weightedContribution = rawMean * yearWeight;

    await Student.findByIdAndUpdate(studentId, {
      $set: { currentYearOfStudy: nextYear, currentSemester: 1 },
      $push: { promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date()},
      academicHistory: {
        yearOfStudy: actualCurrentYear,
        annualMeanMark: rawMean,
        weightedContribution: weightedContribution,
        unitsTakenCount: statusResult.summary.totalExpected,
        failedUnitsCount: statusResult.summary.failed,
        isRepeatYear: false
      } }
    });

    return { success: true, message: `Successfully promoted to Year ${nextYear}` };
  }

  // Logic to protect students with Specials/Incompletes from being "Failed"
  let blockMessage = `Promotion Blocked: `;
  if (statusResult?.status === "SPECIALS PENDING") {
    blockMessage += "Student has pending Special Examinations. These must be sat and graded before promotion.";
  } else if (statusResult?.status === "REPEAT YEAR") {
    blockMessage += "Student is required to repeat the year based on academic performance.";
  } else {
    blockMessage += `Current status is '${statusResult?.status}'.`;
  }

  return { 
    success: false, 
    message: blockMessage, 
    details: statusResult 
  };
};

export const bulkPromoteClass = async ( programId: string, yearToPromote: number, academicYearName: string ) => {
  // 1. Get everyone currently in the year AND those already promoted to the next year
  const nextYear = yearToPromote + 1;
  const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active", });

  const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };

  for (const student of students) {
    try {
      // Cast to 'any' or use (student as any)._id to bypass the 'unknown' check
      const studentId = (student._id as any).toString();

      // Skip if already in the target year or higher
      if (student.currentYearOfStudy >= nextYear) { results.alreadyPromoted++; results.promoted++; continue; }
      
      const res = await promoteStudent(studentId);

      if (res.success) results.promoted++;
      else results.failed++;
    } catch (err: any) {
      results.errors.push(`${student.regNo}: ${err.message}`);
    }
  }

  return results;
};

// export const calculateStudentStatus = async ( studentId: any, programId: any, academicYearName: string, yearOfStudy: number = 1 ) => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");

//   // 1. Determine dispaly year for UI
//   let displayYearName = academicYearName;
//   if (!displayYearName || displayYearName === "N/A") {
//     const latestGrade = await FinalGrade.findOne({ student: studentId }).populate("academicYear").sort({ createdAt: -1 }); displayYearName = (latestGrade?.academicYear as any)?.year || "N/A"; }

//   // 1. Get Curriculum & explicitly type the populated unit
//   const curriculum = (await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean()) as any[];

//   if (!curriculum.length) {
//     return {
//       status: "NO CURRICULUM", variant: "info", details: `No units defined for Year ${yearOfStudy} in this program.`,
//       summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 }, passedList: [],
//     };
//   }

//   // 2. Get Grades (All history to track Retake -> Re-Retake lifecycle)
//   const grades = (await FinalGrade.find({ student: studentId }).populate({ path: "programUnit", populate: { path: "unit" } }).populate({ path: "academicYear", model: "AcademicYear" }).sort({ createdAt: -1 }).lean()) as any[];

//   //     console.log(`DEBUG: Found ${grades.length} grades for student.`);
//   // if (grades.length > 0) { console.log("Sample Grade AcademicYear Data:", JSON.stringify(grades[0].academicYear, null, 2)); }
  
//   // console.log("RE-CHECK Sample Grade AcademicYear Data:", grades[0]?.academicYear);

//   // 3. Map grades by UNIT CODE
//    const unitResults = new Map();
//   // const unitResults = new Map<string,  { status: string; attemptType: string; attemptNumber: number; totalMark: number; }>();

//   grades.forEach((g) => {
//     if (!g.programUnit || !g.programUnit.unit) { console.warn( `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`, ); return; }

//     const unitCode = g.programUnit.unit.code.toUpperCase();
//     // const unitCode = g.programUnit?.unit?.code?.toUpperCase();
//     if (!unitCode) return;
//     const existing = unitResults.get(unitCode);
//     if (existing?.status === "PASS") return;
//     unitResults.set(unitCode, { status: g.status, attemptType: g.attemptType, attemptNumber: g.attemptNumber, totalMark: g.totalMark, remarks: g.remarks || "" });
//   });

//   // Trackers for the Coordinator View
//   let totalMarksSum = 0;
//   const passedUnits: any[] = []; const failedList: string[] = []; const retakeUnits: string[] = []; const reRetakeUnits: string[] = [];
//   const missingUnits: string[] = [];  const incompleteUnits: string[] = []; const specialList: { displayName: string; grounds: string }[] = [];

//   // 4. Compare Curriculum against results
//   curriculum.forEach((pUnit: any) => {
//     const unitCode = pUnit.unit?.code?.trim().toUpperCase();
//     const unitName = pUnit.unit?.name;
//     const displayName = `${unitCode}: ${unitName}`;
//     const record = unitResults.get(unitCode);

//     if (!record) { missingUnits.push(displayName); totalMarksSum +=0;
//     } else if (record.status === "PASS") {
//       totalMarksSum += record.totalMark; const numericMark = record.totalMark || 0;
//       const letterGrade = getLetterGrade(numericMark, settings);
//       passedUnits.push({ code: unitCode,  name: unitName, mark: numericMark, grade: letterGrade });
//     } else if (record.status === "SPECIAL") {
//       let grounds = "Administrative"; 
//       if (record.remarks.toLowerCase().includes("financial")) grounds = "Financial";
//       if (record.remarks.toLowerCase().includes("compassionate")) grounds = "Compassionate";
      
//       specialList.push({ displayName, grounds });
//     } else if (record.status === "INCOMPLETE") {
//       incompleteUnits.push(displayName);
//     } else {
//       // Logic for failures based on attempts
//       if (record.attemptNumber >= 3) reRetakeUnits.push(displayName);
//       else if (record.attemptNumber === 2) retakeUnits.push(displayName);
//       else failedList.push(displayName);
//     }
//   });

//   const totalExpected = curriculum.length;
//   const totalFailed = failedList.length + retakeUnits.length + reRetakeUnits.length;
//   const weightedMean = totalMarksSum / totalExpected;
  

//   // 5. Determine UI Status
//   let status = "IN GOOD STANDING";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = `Year ${yearOfStudy} curriculum units cleared.`;

//   // ENG 23.c: Deregistration (Absent from 6+ exams)
//   if (missingUnits.length >= 6) {
//     // ENG 23.c
//     status = "DEREGISTERED"; variant = "error"; details = "Automatic deregistration: Absent from 6+ examinations.";
//   } else if (reRetakeUnits.length > 0) { status = "CRITICAL FAILURE";  variant = "error"; details = "Student failed a third attempt (Re-Retake).";
//   } else if (totalFailed >= totalExpected / 2 || weightedMean < 40) {
//     // ENG 16.a
//     status = "REPEAT YEAR"; variant = "error"; details = "Failed 50% or more units, or mean mark below 40%.";
//   } else if (totalFailed > totalExpected / 3) {
//     // ENG 15.h
//     status = "RETAKE REQUIRED"; variant = "warning"; details = "Failures exceed 1/3 of units. Must retake during ordinary sessions.";
//   } else if (totalFailed > 0) {
//     // ENG 13.a
//     status = "SUPPLEMENTARY PENDING"; variant = "warning"; details = `Eligible for supplementary exams in ${totalFailed} units.`;
//   } else if (specialList.length > 0) {
//     status = "SPECIAL EXAM PENDING"; variant = "info"; details = `Awaiting results for ${specialList.length} special exam(s).`;
//   } else if (missingUnits.length > 0) {
//     status = "INCOMPLETE DATA"; variant = "info"; details = "Some unit records are missing from the system.";
//   }

//   const sessionRecord = grades.find( (g) => g.programUnit?.requiredYear === yearOfStudy );

//   let actualSessionName = "N/A";

//   if (academicYearName && academicYearName !== "N/A") {
//     actualSessionName = academicYearName;
//   } else {
//     // 1. Try to find a grade in the CURRENT year of study that has a valid year
//     const yearSpecificGrade = grades.find((g) => g.programUnit?.requiredYear === yearOfStudy && g.academicYear?.year );

//     if (yearSpecificGrade?.academicYear?.year) {
//       actualSessionName = yearSpecificGrade.academicYear.year;
//     } else if (grades.length > 0) {
//       // 2. If Year 2 is empty, but Year 1 exists, we "guess" the next year
//       const previousYearGrade = grades.find((g) => g.academicYear?.year);
//       if (previousYearGrade?.academicYear?.year) {
//         const baseYear = previousYearGrade.academicYear.year; // e.g., "2023/2024"
//         // Simple logic to increment if we are looking at a higher year of study
//         if (yearOfStudy > (previousYearGrade.programUnit?.requiredYear || 1)) {
//           // Optional: Add logic to increment "2023/2024" to "2024/2025"
//           actualSessionName = baseYear; // For now, keep as base to avoid wrong guesses
//         } else {
//           actualSessionName = baseYear;
//         }
//       }
//     }
//   }

//   return {
//     status, variant, details,
//     academicYearName: actualSessionName, yearOfStudy: yearOfStudy, weightedMean: weightedMean.toFixed(2),
//     summary: { totalExpected: curriculum.length, passed: passedUnits.length, failed: totalFailed, missing: missingUnits.length },
//     missingList: missingUnits, passedList: passedUnits, failedList: failedList, retakeList: retakeUnits, reRetakeList: reRetakeUnits, specialList: specialList, incompleteList: incompleteUnits,
//   };
// };

// changes

// export const previewPromotion = async ( programId: string, yearToPromote: number, academicYearName: string ) => {
//  // 1. Fetch students in the targeted year AND those already in the next year
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active", }).lean();

//   const eligible: any[] = [];
//   const blocked: any[] = [];

//   for (const student of students) {
//     const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
//     const statusResult = await calculateStudentStatus( student._id, programId, academicYearName, yearToPromote );

//     const report = {
//       id: student._id, regNo: student.regNo, name: student.name,
//       status: isAlreadyPromoted ? "ALREADY PROMOTED" : (statusResult?.status || "PENDING"),
//       summary: statusResult?.summary, reasons: [] as string[], isAlreadyPromoted
//     };

//     if (isAlreadyPromoted || statusResult?.status === "IN GOOD STANDING") {
//       eligible.push(report);
//     } else {
//       // Add specific reasons for the block
//       if (statusResult?.specialList?.length) {
//         report.reasons.push( ...statusResult.specialList.map((item) => `${item.displayName} - SPECIAL: ${item.grounds} Grounds`));
//       }
//       if (statusResult?.incompleteList?.length) {
//         report.reasons.push(...statusResult.incompleteList.map(u => `${u} - INCOMPLETE`));
//       }
//       if (statusResult?.missingList?.length) {
//         report.reasons.push(...statusResult.missingList.map(u => `MISSING: ${u}`));
//       }
//       if (statusResult?.failedList?.length) {
//         report.reasons.push(...statusResult.failedList.map(u => `FAILED: ${u}`));
//       }
//       if (statusResult?.retakeList?.length) {
//         report.reasons.push(...statusResult.retakeList.map(u => `RETAKE FAILED: ${u}`));
//       }
//       if (statusResult?.reRetakeList?.length) {
//         report.reasons.push(...statusResult.reRetakeList.map(u => `CRITICAL RE-RETAKE FAILED: ${u}`));
//       }
//       blocked.push(report);
//     }
//   }

//   return { totalProcessed: students.length, eligibleCount: eligible.length, blockedCount: blocked.length, eligible, blocked };
// };

// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId);
//   if (!student) throw new Error("Student not found");

//   // 1. Get the current active academic year for context
//   const actualCurrentYear = student.currentYearOfStudy || 1;

//   // 2. Calculate status based on the student's CURRENT year of study
//   const statusResult = await calculateStudentStatus(
//     student._id, student.program,
//     // "",
//     "N/A", actualCurrentYear, );

//   // 3. Promotion Guard: Only "IN GOOD STANDING" can move up
//   if (statusResult?.status === "IN GOOD STANDING") {
// const nextYear = actualCurrentYear + 1;

//     // Check if the student has reached the maximum years for their program (Optional)
//     // if (nextYear > student.programDuration) return { success: false, message: "Student has completed final year." };

//     await Student.findByIdAndUpdate(studentId, {
//       $set: { currentYearOfStudy: nextYear, currentSemester: 1, },
//       // Log the promotion in history if you have a history field
//       $push: { promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date() }   }
//     });

//     return { success: true, message: `Successfully promoted to Year ${nextYear}` };
//   }

//    return { success: false, message: `Promotion Blocked: Student is '${statusResult?.status}' for Year ${actualCurrentYear}.`, details: statusResult };
// };

