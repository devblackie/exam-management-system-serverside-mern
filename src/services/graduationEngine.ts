// serverside/src/services/graduationEngine.ts

import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";

export interface GraduationResult {
  studentId: string;
  regNo: string;
  weightedAggregateAverage: number;
  classification: string;
  isEligible: boolean;
  missingRequirements: string[];
}

export const calculateGraduationStatus = async (
  studentId: string,
): Promise<GraduationResult> => {
  const student = (await Student.findById(studentId)
    .populate("program")
    .lean()) as any;
  const settings = await InstitutionSettings.findOne().lean();

  const history = student.academicHistory || [];
  const duration = student.program?.durationYears || 5;

  // 1. Check for Pending Units (Incompletes, Fails, Specials) across ALL years
  // Note: Your promoteStudent logic already pushes to history,
  // but we must ensure Year 5 (Final Year) is also processed.

  let totalWeightedScore = 0;
  let totalWeightAccounted = 0;
  const missingYears: number[] = [];

  // 2. Sum up the weighted contributions from academic history
  history.forEach((yearRecord: any) => {
    totalWeightedScore += yearRecord.weightedContribution;
    // We infer weight used by dividing contribution by mean
    const weight =
      yearRecord.annualMeanMark > 0
        ? yearRecord.weightedContribution / yearRecord.annualMeanMark
        : 0;
    totalWeightAccounted += weight;
  });

  // 3. Determine Classification based on WAA
  // Standard Engineering Scales (usually 70+, 60-69, 50-59, 40-49)
  const waa = totalWeightedScore;
  let classification = "FAIL";

  if (waa >= 70) classification = "FIRST CLASS HONOURS";
  else if (waa >= 60)
    classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
  else if (waa >= 50)
    classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
  else if (waa >= 40) classification = "PASS";

  // 4. Eligibility Check
  const hasFailedUnits = history.some((h: any) => h.failedUnitsCount > 0);
  const isFinalYearDone = history.some((h: any) => h.yearOfStudy === duration);

  return {
    studentId: student._id,
    regNo: student.regNo,
    weightedAggregateAverage: parseFloat(waa.toFixed(2)),
    classification,
    isEligible: !hasFailedUnits && isFinalYearDone,
    missingRequirements: hasFailedUnits
      ? ["Uncleared Failed Units in History"]
      : [],
  };
};
