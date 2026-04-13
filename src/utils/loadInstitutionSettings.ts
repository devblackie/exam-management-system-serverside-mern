// serverside/src/utils/loadInstitutionSettings.ts
import InstitutionSettings from "../models/InstitutionSettings";

export const loadInstitutionSettings = async (institutionId?: any) => {
  let settings: any = null;

  // 1. Try with institution filter first (most specific)
  if (institutionId) {
    settings = await InstitutionSettings.findOne({institution: institutionId}).lean();
    if (settings) {
      // console.log(`[directTemplate] Settings loaded for institution ${institutionId}`);
      return settings;
    }
    console.warn(`[directTemplate] No settings found for institution ${institutionId}, trying global fallback`);
  }

  // 2. Try without filter (single-institution deployments often omit the field)
  settings = await InstitutionSettings.findOne().lean();
  if (settings) {
    console.warn(`[directTemplate] Using global institution settings (no institution filter)`);
    return settings;
  }

  // 3. Return safe defaults so the template still generates
  console.warn(`[directTemplate] No InstitutionSettings document found — using hardcoded defaults`);
  return {
    passMark: 40,
    gradingScale: [ { min: 70, grade: "A" }, { min: 60, grade: "B" }, { min: 50, grade: "C" }, { min: 40, grade: "D" }],
    cat1Max: 20, cat2Max: 20, cat3Max: 20, assignmentMax: 10, practicalMax: 10, examMax: 70,
  };
}
