// // src/models/Program.ts
// import mongoose, { Document, Schema } from "mongoose";
// import { normalizeProgramName } from "../services/programNormalizer";

// export interface IProgram extends Document {
//   institution: mongoose.Types.ObjectId;
//   name: string;
//   cleanName: string; // normalized version
//   code: string;
//   durationYears: number;
//   maxCompletionYears: number; // ENG 19.d/e & 22.f/g
//   schoolType: "ENGINEERING" | "IT" | "MEDICINE" | "GENERAL";
// }

// const schema = new Schema<IProgram>(
//   {
//     institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
//     name: { type: String, required: true, trim: true},
//     cleanName: { type: String, required: false, index: true },
//     code: { type: String, required: true, uppercase: true, trim: true },    
//     durationYears: {type: Number, default: 4, min: 1, max: 8 },
//     maxCompletionYears: { type: Number, default: 8 }, // Default for B.Ed/General
//     schoolType: { type: String, enum: ["ENGINEERING", "IT", "MEDICINE", "GENERAL"], default: "GENERAL" }
//   },
//   { timestamps: true }
// );

// schema.pre("save", function (next) {
//   if (this.isModified("name")) this.cleanName = normalizeProgramName(this.name);  
//   // Auto-set max years based on ENG 19/22
//     if (this.name.includes("Engineering")) this.maxCompletionYears = 10;
//   next();
// });

// // Before updateOne / findOneAndUpdate
// function applyNormalization(this: any, next: any) {
//   const update = this.getUpdate();
//   if (update?.name) update.cleanName = normalizeProgramName(update.name);  
//   if (update?.$set?.name) update.$set.cleanName = normalizeProgramName(update.$set.name);
//   next();
// }

// schema.pre("findOneAndUpdate", applyNormalization);
// schema.pre("updateOne", applyNormalization);

// schema.index({ institution: 1, code: 1 }, { unique: true });
// schema.index({ institution: 1, cleanName: 1 }, { unique: true }); 
// schema.index({ code: 1 });

// export default mongoose.model<IProgram>("Program", schema);









// serverside/src/models/Program.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IProgramRuleOverrides {
  maxDurationMultiplier?: number;
  passMark?:              number;
  suppMarkCap?:           number;
  maxCarryForwardUnits?:  number;
  caWeight?:              number;
  examWeight?:            number;
  semesterWeights?:       Array<{ year: number; weight: number }>;
}

export interface IProgram extends Document {
  institution:      mongoose.Types.ObjectId;
  schoolCode:       string;
  departmentCode:   string;
  name:             string;
  code:             string;
  description?:     string;

  // ── The authoritative program length ──────────────────────────────────────
  // ENG.19(d/e): max study years = durationYears × settings.ruleSet.maxDurationMultiplier
  // Financial Engineering (4yr) × 2 = 8yr max  ← CORRECT, no name-sniffing needed
  // BSc Civil Engineering (5yr) × 2 = 10yr max  ← CORRECT
  // BEd Computer Science (4yr) × 2 = 8yr max    ← CORRECT
  // The cron and status engine ALWAYS use this field — never infer from degreeType or name
  durationYears:    number;

  degreeType:       string;  // "BSc" | "BEd" | "BTech" | "BEng" | "BArch" | "MBBS" | "LLB" | "BPharm" | "Other"
  intakes:          ("JAN" | "MAY" | "SEPT")[];
  supportedEntryTypes: ("Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4")[];
  ruleOverrides:    IProgramRuleOverrides;
  isActive:         boolean;
  createdAt:        Date;
  updatedAt:        Date;
}

const ProgramRuleOverridesSchema = new Schema<IProgramRuleOverrides>({
  maxDurationMultiplier: { type: Number, min: 1,   max: 5 },
  passMark:              { type: Number, min: 0,   max: 100 },
  suppMarkCap:           { type: Number, min: 0,   max: 100 },
  maxCarryForwardUnits:  { type: Number, min: 0,   max: 10 },
  caWeight:              { type: Number, min: 0,   max: 100 },
  examWeight:            { type: Number, min: 0,   max: 100 },
  semesterWeights: [{
    year:   { type: Number, required: true },
    weight: { type: Number, required: true, min: 0, max: 1 },
  }],
}, { _id: false });

const ProgramSchema = new Schema<IProgram>(
  {
    institution:    { type: Schema.Types.ObjectId, ref: "Institution", required: true },
    schoolCode:     { type: String, required: true, uppercase: true, trim: true },
    departmentCode: { type: String, required: true, uppercase: true, trim: true },
    name:           { type: String, required: true, trim: true },
    code:           { type: String, required: true, uppercase: true, trim: true },
    description:    { type: String },

    // ── Program duration — the ONLY source of truth for max study years ──────
    // Do NOT use degreeType or name to infer duration in any service/cron.
    // Always: maxYears = program.durationYears × settings.ruleSet.maxDurationMultiplier
    durationYears: {
      type:     Number,
      required: true,
      min:      1,
      max:      10,
      default:  5,
    },

    degreeType: {
      type:     String,
      required: true,
      enum:     ["BSc","BEd","BTech","BEng","BArch","MBBS","LLB","BPharm","Other"],
      default:  "BSc",
    },
    intakes: {
      type:    [String],
      enum:    ["JAN","MAY","SEPT"],
      default: ["SEPT"],
    },
    supportedEntryTypes: {
      type:    [String],
      enum:    ["Direct","Mid-Entry-Y2","Mid-Entry-Y3","Mid-Entry-Y4"],
      default: ["Direct"],
    },
    ruleOverrides: { type: ProgramRuleOverridesSchema, default: () => ({}) },
    isActive:      { type: Boolean, default: true },
  },
  { timestamps: true },
);

ProgramSchema.index({ institution: 1, code: 1 }, { unique: true });
ProgramSchema.index({ institution: 1, schoolCode: 1 });
ProgramSchema.index({ institution: 1, departmentCode: 1 });
ProgramSchema.index({ institution: 1, isActive: 1 });
ProgramSchema.index(
  { institution: 1, schoolCode: 1, departmentCode: 1, name: 1 },
  { unique: true, collation: { locale: "en", strength: 2 } }, // case-insensitive
);

export default mongoose.model<IProgram>("Program", ProgramSchema);