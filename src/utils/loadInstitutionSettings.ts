// // serverside/src/utils/loadInstitutionSettings.ts
// import InstitutionSettings from "../models/InstitutionSettings";

// export const loadInstitutionSettings = async (institutionId?: any) => {
//   let settings: any = null;

//   // 1. Try with institution filter first (most specific)
//   if (institutionId) {
//     settings = await InstitutionSettings.findOne({institution: institutionId}).lean();
//     if (settings) {
//       // console.log(`[directTemplate] Settings loaded for institution ${institutionId}`);
//       return settings;
//     }
//     console.warn(`[directTemplate] No settings found for institution ${institutionId}, trying global fallback`);
//   }

//   // 2. Try without filter (single-institution deployments often omit the field)
//   settings = await InstitutionSettings.findOne().lean();
//   if (settings) {
//     console.warn(`[directTemplate] Using global institution settings (no institution filter)`);
//     return settings;
//   }

//   // 3. Return safe defaults so the template still generates
//   console.warn(`[directTemplate] No InstitutionSettings document found — using hardcoded defaults`);
//   return {
//     passMark: 40,
//     gradingScale: [ { min: 70, grade: "A" }, { min: 60, grade: "B" }, { min: 50, grade: "C" }, { min: 40, grade: "D" }],
//     cat1Max: 20, cat2Max: 20, cat3Max: 20, assignmentMax: 10, practicalMax: 10, examMax: 70,
//   };
// }













// serverside/src/utils/loadInstitutionSettings.ts — COMPLETE FIXED VERSION

import InstitutionSettings, {
  DEFAULT_GRADING_SCALE,
  DEFAULT_WAA_CLASSIFICATION,
  DEKUT_DEFAULT_WEIGHTS,
  IRuleSet,
  IGradeEntry,
  IWAAClassification,
  ISemesterWeightMap,
  IDocumentMeta,
  IBrandingAssets,
} from "../models/InstitutionSettings";
import Program from "../models/Program";
import { cached, invalidateCache } from "./cache";

// What a program can override (only program-specific fields)
interface ProgramRuleOverride {
  maxDurationMultiplier?: number;
  passMark?:              number;
  suppMarkCap?:           number;
  maxCarryForwardUnits?:  number;
  caWeight?:              number;
  examWeight?:            number;
  semesterWeights?:       Array<{ year: number; weight: number }>;
}

export interface ResolvedSettings {
  rules: {
    supplementaryThreshold:  number;
    stayoutThreshold:        number;
    repeatYearMeanThreshold: number;
    passMark:                number;
    maxCarryForwardUnits:    number;
    carryForwardToFinalYear: boolean;
    maxDurationMultiplier:   number;
    maxStudyYears:           number;
    maxAttempts:             number;
    caWeight:                number;
    examWeight:              number;
    catMax:                  number;   // ← kept here
    assignmentMax:           number;
    practicalMax:            number;
    labMax:                  number;
    suppMarkCap:             number;
    hasLab:                  boolean;
    hasPractical:            boolean;
    hasWorkshop:             boolean;
    useSemesterWeighting:    boolean;
    minCourseworkAttendance: number;
    maxAbsentExams:          number;
    gradeAppealWindowDays:   number;
  };
  // ── Flat convenience accessors (used by uploadTemplate.ts and gradingCore.ts) ──
  // These mirror rules.* exactly — no separate computation needed
  passMark:       number;
  cat1Max:        number;   // ← catMax (single CAT max — all CATs share same max)
  cat2Max:        number;   // ← same as catMax
  cat3Max:        number;   // ← same as catMax (0 if hasWorkshop)
  assignmentMax:  number;
  practicalMax:   number;
  gradingScale:      IGradeEntry[];
  waaClassification: IWAAClassification[];
  semesterWeights:   ISemesterWeightMap[];
  docMeta:           IDocumentMeta;
  branding:          IBrandingAssets;
}

export async function loadInstitutionSettings(
  institutionId: string,
  programId?:    string,
): Promise<ResolvedSettings> {
  const cacheKey = `settings:${institutionId}:${programId ?? "base"}`;

  return cached(cacheKey, async () => {
    // Explicitly typed lean queries — eliminates all "does not exist on type" errors
    const settingsDoc = await InstitutionSettings
      .findOne({ institution: institutionId })
      .lean<{
        ruleSet?:          Partial<IRuleSet>;
        semesterWeights?:  ISemesterWeightMap[];
        gradingScale?:     IGradeEntry[];
        waaClassification?: IWAAClassification[];
        docMeta?:          IDocumentMeta;
        branding?:         IBrandingAssets;
      }>();

    const programDoc = programId
      ? await Program
          .findById(programId)
          .lean<{
            durationYears:  number;
            ruleOverrides?: ProgramRuleOverride;  // ← typed correctly, no IRuleSet mismatch
          }>()
      : null;

    // Institution-level rules — rs is a known partial shape
    const rs: Partial<IRuleSet> = settingsDoc?.ruleSet ?? {};

    // Program-level overrides — po is ProgramRuleOverride, not IRuleSet
    // Only contains fields that make sense to override per program
    const po: ProgramRuleOverride = programDoc?.ruleOverrides ?? {};

    const programDuration       = programDoc?.durationYears ?? 5;
    const maxDurationMultiplier = po.maxDurationMultiplier ?? rs.maxDurationMultiplier ?? 2.0;

    const rules = {
      // ── Institution-level (not overridable per program) ───────────────────
      // These are consistent across all programs in an institution
      supplementaryThreshold:  rs.supplementaryThreshold  ?? (1 / 3),
      stayoutThreshold:        rs.stayoutThreshold        ?? 0.5,
      repeatYearMeanThreshold: rs.repeatYearMeanThreshold ?? 40,
      maxAttempts:             rs.maxAttempts             ?? 5,
      catMax:                  rs.catMax                  ?? 20,
      assignmentMax:           rs.assignmentMax           ?? 10,
      practicalMax:            rs.practicalMax            ?? 10,
      labMax:                  rs.labMax                  ?? 30,
      hasLab:                  rs.hasLab                  ?? true,
      hasPractical:            rs.hasPractical            ?? true,
      hasWorkshop:             rs.hasWorkshop             ?? false,
      useSemesterWeighting:    rs.useSemesterWeighting    ?? true,
      minCourseworkAttendance: rs.minCourseworkAttendance ?? 0.75,
      maxAbsentExams:          rs.maxAbsentExams          ?? 6,
      gradeAppealWindowDays:   rs.gradeAppealWindowDays   ?? 28,
      carryForwardToFinalYear: rs.carryForwardToFinalYear ?? false,

      // ── Program-overridable (program wins, then institution, then default) ─
      passMark:             po.passMark            ?? rs.passMark            ?? 40,
      suppMarkCap:          po.suppMarkCap          ?? rs.suppMarkCap          ?? 40,
      maxCarryForwardUnits: po.maxCarryForwardUnits ?? rs.maxCarryForwardUnits ?? 2,
      maxDurationMultiplier,
      maxStudyYears:        Math.round(programDuration * maxDurationMultiplier),
      caWeight:             po.caWeight             ?? rs.caWeight             ?? 30,
      examWeight:           po.examWeight           ?? rs.examWeight           ?? 70,
    };

    const gradingScale =
      (settingsDoc?.gradingScale?.length ?? 0) > 0
        ? settingsDoc!.gradingScale!
        : DEFAULT_GRADING_SCALE;

    const waaClassification =
      (settingsDoc?.waaClassification?.length ?? 0) > 0
        ? settingsDoc!.waaClassification!
        : DEFAULT_WAA_CLASSIFICATION;

    const semesterWeights: ISemesterWeightMap[] =
      (po.semesterWeights?.length ?? 0) > 0
        ? po.semesterWeights!
        : (settingsDoc?.semesterWeights?.length ?? 0) > 0
          ? settingsDoc!.semesterWeights!
          : DEKUT_DEFAULT_WEIGHTS;

    const docMeta: IDocumentMeta = settingsDoc?.docMeta ?? {
      universityName: process.env.INST_NAME        ?? "University",
      universityAbbr: process.env.INST_ABBR         ?? "UNIV",
      schoolName:     process.env.SCHOOL_NAME       ?? "School",
      departmentName: process.env.DEPARTMENT_NAME   ?? "Department",
      registrar:      process.env.REGISTRAR         ?? "Academic Registrar",
      postalAddress:  process.env.POSTAL_ADDRESS    ?? "",
      telephone:      process.env.CELL_PHONE        ?? "",
      email:          process.env.SCHOOL_EMAIL      ?? "",
      website:        "",
      country:        "Kenya",
      city:           "",
    };

    const branding: IBrandingAssets = settingsDoc?.branding ?? {
      cmsHeaderColor:    "#1F4E79",
      cmsAccentColor:    "#D4AF37",
      wordDocFontFamily: "Times New Roman",
      wordDocFontSize:   12,
      useLetterhead:     true,
    };

    return {
      rules,
      gradingScale,
      waaClassification,
      semesterWeights,
      docMeta,
      branding,
      // Flat accessors — backward compat for uploadTemplate.ts and other callers
      passMark: rules.passMark,
      cat1Max: rules.catMax,
      cat2Max: rules.catMax,
      cat3Max: rules.hasWorkshop ? 0 : rules.catMax, // workshops have no CATs
      assignmentMax: rules.assignmentMax,
      practicalMax: rules.practicalMax,
    };
  }, 300);
}

export function invalidateSettingsCache(institutionId: string): void {
  invalidateCache(`settings:${institutionId}:`);
}