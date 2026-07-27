"use strict";
// serverside/src/utils/loadInstitutionSettings.ts — COMPLETE FIXED VERSION
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadInstitutionSettings = loadInstitutionSettings;
exports.invalidateSettingsCache = invalidateSettingsCache;
const InstitutionSettings_1 = __importStar(require("../models/InstitutionSettings"));
const Program_1 = __importDefault(require("../models/Program"));
const cache_1 = require("./cache");
async function loadInstitutionSettings(institutionId, programId) {
    const cacheKey = `settings:${institutionId}:${programId ?? "base"}`;
    return (0, cache_1.cached)(cacheKey, async () => {
        // Explicitly typed lean queries — eliminates all "does not exist on type" errors
        const settingsDoc = await InstitutionSettings_1.default
            .findOne({ institution: institutionId })
            .lean();
        const programDoc = programId
            ? await Program_1.default
                .findById(programId)
                .lean()
            : null;
        // Institution-level rules — rs is a known partial shape
        const rs = settingsDoc?.ruleSet ?? {};
        // Program-level overrides — po is ProgramRuleOverride, not IRuleSet
        // Only contains fields that make sense to override per program
        const po = programDoc?.ruleOverrides ?? {};
        const programDuration = programDoc?.durationYears ?? 5;
        const maxDurationMultiplier = po.maxDurationMultiplier ?? rs.maxDurationMultiplier ?? 2.0;
        const rules = {
            // ── Institution-level (not overridable per program) ───────────────────
            // These are consistent across all programs in an institution
            supplementaryThreshold: rs.supplementaryThreshold ?? (1 / 3),
            stayoutThreshold: rs.stayoutThreshold ?? 0.5,
            repeatYearMeanThreshold: rs.repeatYearMeanThreshold ?? 40,
            maxAttempts: rs.maxAttempts ?? 5,
            catMax: rs.catMax ?? 20,
            assignmentMax: rs.assignmentMax ?? 10,
            practicalMax: rs.practicalMax ?? 10,
            labMax: rs.labMax ?? 30,
            hasLab: rs.hasLab ?? true,
            hasPractical: rs.hasPractical ?? true,
            hasWorkshop: rs.hasWorkshop ?? false,
            useSemesterWeighting: rs.useSemesterWeighting ?? true,
            minCourseworkAttendance: rs.minCourseworkAttendance ?? 0.75,
            maxAbsentExams: rs.maxAbsentExams ?? 6,
            gradeAppealWindowDays: rs.gradeAppealWindowDays ?? 28,
            carryForwardToFinalYear: rs.carryForwardToFinalYear ?? false,
            // ── Program-overridable (program wins, then institution, then default) ─
            passMark: po.passMark ?? rs.passMark ?? 40,
            suppMarkCap: po.suppMarkCap ?? rs.suppMarkCap ?? 40,
            maxCarryForwardUnits: po.maxCarryForwardUnits ?? rs.maxCarryForwardUnits ?? 2,
            maxDurationMultiplier,
            maxStudyYears: Math.round(programDuration * maxDurationMultiplier),
            caWeight: po.caWeight ?? rs.caWeight ?? 30,
            examWeight: po.examWeight ?? rs.examWeight ?? 70,
        };
        const gradingScale = (settingsDoc?.gradingScale?.length ?? 0) > 0
            ? settingsDoc.gradingScale
            : InstitutionSettings_1.DEFAULT_GRADING_SCALE;
        const waaClassification = (settingsDoc?.waaClassification?.length ?? 0) > 0
            ? settingsDoc.waaClassification
            : InstitutionSettings_1.DEFAULT_WAA_CLASSIFICATION;
        const semesterWeights = (po.semesterWeights?.length ?? 0) > 0
            ? po.semesterWeights
            : (settingsDoc?.semesterWeights?.length ?? 0) > 0
                ? settingsDoc.semesterWeights
                : InstitutionSettings_1.DEKUT_DEFAULT_WEIGHTS;
        const docMeta = settingsDoc?.docMeta ?? {
            universityName: process.env.INST_NAME ?? "University",
            universityAbbr: process.env.INST_ABBR ?? "UNIV",
            schoolName: process.env.SCHOOL_NAME ?? "School",
            departmentName: process.env.DEPARTMENT_NAME ?? "Department",
            registrar: process.env.REGISTRAR ?? "Academic Registrar",
            postalAddress: process.env.POSTAL_ADDRESS ?? "",
            telephone: process.env.CELL_PHONE ?? "",
            email: process.env.SCHOOL_EMAIL ?? "",
            website: "",
            country: "Kenya",
            city: "",
        };
        const branding = settingsDoc?.branding ?? {
            cmsHeaderColor: "#1F4E79",
            cmsAccentColor: "#D4AF37",
            wordDocFontFamily: "Times New Roman",
            wordDocFontSize: 12,
            useLetterhead: true,
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
function invalidateSettingsCache(institutionId) {
    (0, cache_1.invalidateCache)(`settings:${institutionId}:`);
}
