"use strict";
// // src/models/InstitutionSettings.ts 
// import mongoose, { Schema, Document } from "mongoose";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEKUT_DEFAULT_WEIGHTS = exports.DEFAULT_WAA_CLASSIFICATION = exports.DEFAULT_GRADING_SCALE = void 0;
// export interface IInstitutionSettings extends Document {
//   institution: mongoose.Types.ObjectId;
//   cat1Max: number; cat2Max: number; cat3Max: number;
//   assignmentMax: number; practicalMax: number; workshopMax: number;
//   examMax: 70; passMark: number; unitType:string;
//   gradingScale?: Array<{ min: number; grade: string; points?: number }>;
// }
// const schema = new Schema<IInstitutionSettings>(
//   {
//     institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },
//     cat1Max: { type: Number, default: 20 },
//     cat2Max: { type: Number, default: 20 },
//     cat3Max: { type: Number, default: 0 },
//     assignmentMax: { type: Number, default: 10 },
//     practicalMax: { type: Number, default: 10 },
//     workshopMax: { type: Number, default: 100 },
//     examMax: { type: Number, default: 70, enum: [70] },
//     passMark: { type: Number, default: 40 },
//     unitType: { type: String, enum: ["theory", "lab", "workshop"], default: "theory" },
//     gradingScale: [ { min: { type: Number, required: true }, grade: { type: String, required: true }, points: { type: Number }}],
//   },
//   { timestamps: true },
// );
// schema.index({ institution: 1 }, { unique: true });
// export default mongoose.model<IInstitutionSettings>("InstitutionSettings", schema);
// serverside/src/models/InstitutionSettings.ts
const mongoose_1 = __importStar(require("mongoose"));
// ── Defaults ──────────────────────────────────────────────────────────────────
exports.DEFAULT_GRADING_SCALE = [
    { min: 70, max: 100, grade: "A", label: "Excellent" },
    { min: 60, max: 69, grade: "B", label: "Good" },
    { min: 50, max: 59, grade: "C", label: "Satisfactory" },
    { min: 40, max: 49, grade: "D", label: "Pass" },
    { min: 0, max: 39, grade: "E", label: "Fail" },
];
exports.DEFAULT_WAA_CLASSIFICATION = [
    { min: 70, max: 100, classification: "First Class Honours" },
    { min: 60, max: 69, classification: "Second Class Honours (Upper Division)" },
    { min: 50, max: 59, classification: "Second Class Honours (Lower Division)" },
    { min: 40, max: 49, classification: "Pass" },
    { min: 0, max: 39, classification: "Fail" },
];
exports.DEKUT_DEFAULT_WEIGHTS = [
    { year: 1, weight: 0.15 },
    { year: 2, weight: 0.15 },
    { year: 3, weight: 0.20 },
    { year: 4, weight: 0.25 },
    { year: 5, weight: 0.25 },
];
// ── Schemas ───────────────────────────────────────────────────────────────────
const GradeEntrySchema = new mongoose_1.Schema({
    min: { type: Number, required: true },
    max: { type: Number, required: true },
    grade: { type: String, required: true, enum: ["A", "B", "C", "D", "E"] },
    label: { type: String, required: true },
}, { _id: false });
const WAAClassificationSchema = new mongoose_1.Schema({
    min: { type: Number, required: true },
    max: { type: Number, required: true },
    classification: { type: String, required: true },
}, { _id: false });
const RegNoPatternSchema = new mongoose_1.Schema({
    prefix: { type: String, required: true },
    separator: { type: String, default: "" },
    yearDigits: { type: Number, default: 3 },
    manualRegex: { type: String },
    example: { type: String, required: true },
}, { _id: false });
const DepartmentSchema = new mongoose_1.Schema({
    name: { type: String, required: true },
    shortName: { type: String, required: true },
    code: { type: String, required: true, uppercase: true, trim: true },
    hod: { type: String },
    regNoPatterns: { type: [RegNoPatternSchema], default: [] },
});
const SchoolSchema = new mongoose_1.Schema({
    name: { type: String, required: true },
    shortName: { type: String, required: true },
    code: { type: String, required: true, uppercase: true, trim: true },
    dean: { type: String },
    departments: { type: [DepartmentSchema], default: [] },
});
const RuleSetSchema = new mongoose_1.Schema({
    supplementaryThreshold: { type: Number, default: 1 / 3, min: 0.1, max: 0.49 },
    stayoutThreshold: { type: Number, default: 0.5, min: 0.1, max: 0.9 },
    repeatYearMeanThreshold: { type: Number, default: 40, min: 0, max: 60 },
    passMark: { type: Number, default: 40, min: 0, max: 60 },
    maxCarryForwardUnits: { type: Number, default: 2, min: 0, max: 10 },
    carryForwardToFinalYear: { type: Boolean, default: false },
    maxDurationMultiplier: { type: Number, default: 2.0, min: 1, max: 5 },
    maxAttempts: { type: Number, default: 5, min: 1, max: 20 },
    caWeight: { type: Number, default: 30, min: 0, max: 100 },
    examWeight: { type: Number, default: 70, min: 0, max: 100 },
    catMax: { type: Number, default: 20, min: 1 },
    assignmentMax: { type: Number, default: 10, min: 1 },
    practicalMax: { type: Number, default: 10, min: 0 },
    labMax: { type: Number, default: 30, min: 0 },
    suppMarkCap: { type: Number, default: 40, min: 0, max: 100 },
    hasLab: { type: Boolean, default: true },
    hasPractical: { type: Boolean, default: true },
    hasWorkshop: { type: Boolean, default: false },
    useSemesterWeighting: { type: Boolean, default: true },
    minCourseworkAttendance: { type: Number, default: 0.75, min: 0, max: 1 },
    maxAbsentExams: { type: Number, default: 6, min: 1 },
    gradeAppealWindowDays: { type: Number, default: 28, min: 1 },
}, { _id: false });
const SemesterWeightSchema = new mongoose_1.Schema({
    year: { type: Number, required: true },
    weight: { type: Number, required: true, min: 0, max: 1 },
}, { _id: false });
const BrandingSchema = new mongoose_1.Schema({
    universityLogoPath: { type: String },
    reportHeaderText: { type: String },
    reportFooterText: { type: String },
    cmsHeaderColor: { type: String, default: "#1F4E79" },
    cmsAccentColor: { type: String, default: "#D4AF37" },
    wordDocFontFamily: { type: String, default: "Times New Roman" },
    wordDocFontSize: { type: Number, default: 12 },
    useLetterhead: { type: Boolean, default: true },
}, { _id: false });
const DocumentMetaSchema = new mongoose_1.Schema({
    universityName: { type: String, required: true },
    universityAbbr: { type: String, required: true },
    schoolName: { type: String, required: true },
    departmentName: { type: String, required: true },
    registrar: { type: String, default: "Academic Registrar" },
    postalAddress: { type: String, default: "" },
    telephone: { type: String, default: "" },
    email: { type: String, default: "" },
    website: { type: String, default: "" },
    country: { type: String, default: "Kenya" },
    city: { type: String, default: "" },
}, { _id: false });
const InstitutionSettingsSchema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },
    docMeta: { type: DocumentMetaSchema, required: true },
    schools: { type: [SchoolSchema], default: [] },
    ruleSet: { type: RuleSetSchema, default: () => ({}) },
    semesterWeights: { type: [SemesterWeightSchema], default: [] },
    gradingScale: { type: [GradeEntrySchema], default: [] },
    waaClassification: { type: [WAAClassificationSchema], default: [] },
    branding: { type: BrandingSchema, default: () => ({}) },
    supportedIntakes: {
        type: [String],
        enum: ["JAN", "MAY", "SEPT"],
        default: ["SEPT"],
    },
    enforceRegNoPattern: { type: Boolean, default: false },
}, { timestamps: true });
InstitutionSettingsSchema.index({ "schools.code": 1 });
InstitutionSettingsSchema.index({ "schools.departments.code": 1 });
exports.default = mongoose_1.default.model("InstitutionSettings", InstitutionSettingsSchema);
