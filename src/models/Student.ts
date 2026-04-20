// // src/models/Student.ts
// import mongoose, { Schema, Document } from "mongoose";

// export interface IStudent extends Document {
//   institution: mongoose.Types.ObjectId;
//   regNo: string;
//   name: string;
//   program: mongoose.Types.ObjectId;
//   programType: string; // "B.Sc", "B.Ed", "Diploma"
//   entryType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4";
//   currentYearOfStudy: number;
//   remarks?: string;
//   currentSemester: number;
//   status: "active" | "graduated" | "discontinued" | "deregistered" | "repeat" | "on_leave" | "deferred";
//   admissionAcademicYear: mongoose.Types.ObjectId;
//   intake: string;
//   qualifierSuffix:       string;
 
//   // ── CARRY FORWARD UNITS ───────────────────────────────────────────────────
//   // ENG.14: Up to 2 units carried from previous year.
//   // Populated when a student passes their year with 1–2 failed supps.
//   // Cleared when those units are passed (or escalated to RPU).
//   carryForwardUnits: Array<{ programUnitId: mongoose.Types.ObjectId; unitCode: string; unitName: string; fromYear: number; attemptCount: number; status: "pending" | "passed" | "failed" | "escalated_to_rpu";}>;
//   academicLeavePeriod?: { startDate: Date; endDate: Date; reason: string; type: "compassionate" | "financial" | "other"; };
//   totalTimeOutYears: number; // ENG 19.d/e: Total years allowed out
//   // ENG 22.b: Track attempts per unit (Limit: 5)
//   unitAttemptRegistry: { unitId: mongoose.Types.ObjectId; attempts: { attemptNumber: number; mark: number; passed: boolean; type: string; }[]; }[];

//   // ENG 25.b: Store snapshots of each year for final classification
//   academicHistory: { academicYear: string; yearOfStudy: number; annualMeanMark: number; weightedContribution: number; failedUnitsCount: number; isRepeatYear?: boolean; date?: Date; }[];
//   statusEvents: { fromStatus: string; toStatus: string; date: Date; academicYear: string; reason: string; }[];
// }

// const schema = new Schema<IStudent>({
//   institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
//   regNo: { type: String, required: true, uppercase: true },
//   name: { type: String, required: true },
//   program: { type: Schema.Types.ObjectId, ref: "Program", required: true },
//   programType: { type: String, required: true },
//   entryType: { type: String, default: "Direct" },
//   currentYearOfStudy: { type: Number, default: 1 },
//   currentSemester: { type: Number, default: 1 },
//   admissionAcademicYear: { type: Schema.Types.ObjectId, ref: "AcademicYear", required: true },
//   intake: { type: String, required: true, enum: ["JAN", "MAY", "SEPT"], uppercase: true, default: "SEPT" },
//   status: { type: String, default: "active" },
//   qualifierSuffix: { type: String, default: "" },
 
//     // ── CARRY FORWARD UNITS ───────────────────────────────────────────────
//     carryForwardUnits: [
//       {
//         programUnitId:    { type: Schema.Types.ObjectId, ref: "ProgramUnit" },
//         unitCode:         String,
//         unitName:         String,
//         fromYear:         Number,
//         fromAcademicYear: String,
//         attemptCount:     { type: Number, default: 1 },
//         status:           {
//           type: String,
//           enum: ["pending", "passed", "failed", "escalated_to_rpu"],
//           default: "pending",
//         },
//       },
//     ],
//   academicLeavePeriod: { startDate: Date, endDate: Date, reason: String, type: { type: String, enum: ["compassionate", "financial", "other"] }},
//   totalTimeOutYears: { type: Number, default: 0 },
//   unitAttemptRegistry: [{ unitId: { type: Schema.Types.ObjectId, ref: "Unit" }, attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }]}],
//   academicHistory: [{ academicYear: String, yearOfStudy: Number, annualMeanMark: Number, weightedContribution: Number, failedUnitsCount: Number, isRepeatYear: Boolean, date: Date }],
//   statusEvents: [{ fromStatus: String, toStatus: String, date: { type: Date, default: Date.now }, academicYear: String, reason: String }],
// }, { timestamps: true });

// schema.index({ institution: 1, regNo: 1 }, { unique: true });
// // schema.index({ regNo: 1 });
// schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
// schema.index({ institution: 1, intake: 1 });
// schema.index({ "statusEvents.toStatus": 1, "statusEvents.academicYear": 1 });
// schema.index({ institution: 1, status: 1, currentYearOfStudy: 1 });
// schema.index({ institution: 1, program: 1, currentYearOfStudy: 1, status: 1 });
// schema.index({ institution: 1, admissionAcademicYear: 1, intake: 1 });
// schema.index({ qualifierSuffix: 1 });
// export default mongoose.model<IStudent>("Student", schema);














// // src/models/Student.ts
// import mongoose, { Schema, Document } from "mongoose";

// export interface IStudent extends Document {
//   institution:           mongoose.Types.ObjectId;
//   regNo:                 string;
//   name:                  string;
//   program:               mongoose.Types.ObjectId;
//   programType:           string;
//   entryType:             "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4";
//   currentYearOfStudy:    number;
//   currentSemester:       number;
//   remarks?:              string;
//   status:
//     | "active"
//     | "graduated"
//     | "discontinued"
//     | "deregistered"
//     | "repeat"
//     | "on_leave"
//     | "deferred"
//     | "graduand";

//   admissionAcademicYear: mongoose.Types.ObjectId;
//   intake:                string;

//   // ── QUALIFIER SUFFIX ────────────────────────────────────────────────────
//   // Persists on the student record throughout their journey.
//   // Appended to regNo in all official documents (scoresheets, CMS, senate docs).
//   //   RP1, RP2       = Repeat Year 1st/2nd time
//   //   RP1C, RP2C     = Carry Forward 1st/2nd cycle
//   //   RPU1, RPU2     = Repeat Unit (ENG.16b — 4th attempt failed)
//   //   RA1, RA2       = Re-Admission
//   //   M2, M3         = Mid-Entry Year 2/3
//   //   TF2, TF3       = Transfer Year 2/3
//   // Empty string ("") = normal active student, clean pass
//   qualifierSuffix: string;

//   // ── CARRY FORWARD UNITS ─────────────────────────────────────────────────
//   // ENG.14: Up to 2 units carried from previous year to next.
//   // Populated by carryForwardService.applyCarryForward().
//   // Each entry tracks one specific failed unit being carried.
//   carryForwardUnits: Array<{
//     programUnitId:    mongoose.Types.ObjectId;
//     unitCode:         string;
//     unitName:         string;
//     fromYear:         number;
//     fromAcademicYear: string;
//     attemptCount:     number;
//     status:           "pending" | "passed" | "failed" | "escalated_to_rpu";
//   }>;

//   academicLeavePeriod?: {
//     startDate: Date;
//     endDate:   Date;
//     reason:    string;
//     type:      "compassionate" | "financial" | "other";
//   };

//   totalTimeOutYears: number;

//   unitAttemptRegistry: Array<{
//     unitId:   mongoose.Types.ObjectId;
//     attempts: Array<{
//       attemptNumber: number;
//       mark:          number;
//       passed:        boolean;
//       type:          string;
//     }>;
//   }>;

//   academicHistory: Array<{
//     academicYear:         string;
//     yearOfStudy:          number;
//     annualMeanMark:       number;
//     weightedContribution: number;
//     failedUnitsCount:     number;
//     isRepeatYear?:        boolean;
//     date?:                Date;
//   }>;

//   // statusHistory used by consolidatedMS REINSTATED detection
//   statusHistory: Array<{
//     status:         string;
//     previousStatus: string;
//     date:           Date;
//     reason:         string;
//   }>;

//   statusEvents: Array<{
//     fromStatus:   string;
//     toStatus:     string;
//     date:         Date;
//     academicYear: string;
//     reason:       string;
//   }>;
// }

// const schema = new Schema<IStudent>(
//   {
//     institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
//     regNo:       { type: String, required: true, uppercase: true },
//     name:        { type: String, required: true },
//     program:     { type: Schema.Types.ObjectId, ref: "Program", required: true },
//     programType: { type: String, required: true },
//     entryType:   { type: String, default: "Direct" },

//     currentYearOfStudy: { type: Number, default: 1 },
//     currentSemester:    { type: Number, default: 1 },
//     remarks:            { type: String },

//     admissionAcademicYear: { type: Schema.Types.ObjectId, ref: "AcademicYear", required: true },
//     intake: {
//       type: String, required: true,
//       enum: ["JAN", "MAY", "SEPT"], uppercase: true, default: "SEPT",
//     },
//     status: { type: String, default: "active" },

//     qualifierSuffix: { type: String, default: "" },

//     carryForwardUnits: [
//       {
//         programUnitId:    { type: Schema.Types.ObjectId, ref: "ProgramUnit" },
//         unitCode:         String,
//         unitName:         String,
//         fromYear:         Number,
//         fromAcademicYear: String,
//         attemptCount:     { type: Number, default: 1 },
//         status: {
//           type:    String,
//           enum:    ["pending", "passed", "failed", "escalated_to_rpu"],
//           default: "pending",
//         },
//       },
//     ],

//     academicLeavePeriod: {
//       startDate: Date,
//       endDate:   Date,
//       reason:    String,
//       type:      { type: String, enum: ["compassionate", "financial", "other"] },
//     },

//     totalTimeOutYears: { type: Number, default: 0 },

//     unitAttemptRegistry: [
//       {
//         unitId:   { type: Schema.Types.ObjectId, ref: "Unit" },
//         attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }],
//       },
//     ],

//     academicHistory: [
//       {
//         academicYear:         String,
//         yearOfStudy:          Number,
//         annualMeanMark:       Number,
//         weightedContribution: Number,
//         failedUnitsCount:     Number,
//         isRepeatYear:         Boolean,
//         date:                 Date,
//       },
//     ],

//     statusHistory: [
//       {
//         status:         String,
//         previousStatus: String,
//         date:           { type: Date, default: Date.now },
//         reason:         String,
//       },
//     ],

//     statusEvents: [
//       {
//         fromStatus:   String,
//         toStatus:     String,
//         date:         { type: Date, default: Date.now },
//         academicYear: String,
//         reason:       String,
//       },
//     ],
//   },
//   { timestamps: true },
// );

// schema.index({ institution: 1, regNo: 1 }, { unique: true });
// schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
// schema.index({ institution: 1, intake: 1 });
// schema.index({ "statusEvents.toStatus": 1, "statusEvents.academicYear": 1 });
// schema.index({ institution: 1, status: 1, currentYearOfStudy: 1 });
// schema.index({ institution: 1, program: 1, currentYearOfStudy: 1, status: 1 });
// schema.index({ institution: 1, admissionAcademicYear: 1, intake: 1 });
// schema.index({ qualifierSuffix: 1 });

// export default mongoose.model<IStudent>("Student", schema);
























// src/models/Student.ts
import mongoose, { Schema, Document } from "mongoose";
import type { CarryForwardUnit } from "../services/carryForwardTypes";

export interface IStudent extends Document {
  institution:           mongoose.Types.ObjectId;
  regNo:                 string;
  name:                  string;
  program:               mongoose.Types.ObjectId;
  programType:           string;
  entryType:             "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4";
  currentYearOfStudy:    number;
  currentSemester:       number;
  remarks?:              string;
  status:                "active" | "graduated" | "discontinued" | "deregistered" | "repeat" | "on_leave" | "deferred" | "graduand";
  admissionAcademicYear: mongoose.Types.ObjectId;
  intake:                string;
  qualifierSuffix:       string;
  carryForwardUnits:     CarryForwardUnit[];
  deferredSuppUnits: Array<{ programUnitId: string; unitCode: string; unitName: string; fromYear: number; fromAcademicYear: string; reason: "supp_deferred" | "special_deferred"; addedAt: Date; status: "pending" | "passed" | "failed"; }>;
  academicLeavePeriod?:  { startDate: Date; endDate: Date; reason: string; type: "compassionate" | "financial" | "other"; };
  totalTimeOutYears:     number;
  unitAttemptRegistry:   Array<{ unitId: mongoose.Types.ObjectId; attempts: Array<{ attemptNumber: number; mark: number; passed: boolean; type: string; }>; }>;
  academicHistory:       Array<{ academicYear: string; yearOfStudy: number; annualMeanMark: number; weightedContribution: number; failedUnitsCount: number; isRepeatYear?: boolean; date?: Date; }>;
  statusHistory:         Array<{ status: string; previousStatus: string; date: Date; reason: string; }>;
  statusEvents:          Array<{ fromStatus: string; toStatus: string; date: Date; academicYear: string; reason: string; }>;
}

const schema = new Schema<IStudent>({
  institution:           { type: Schema.Types.ObjectId, ref: "Institution", required: true },
  regNo:                 { type: String, required: true, uppercase: true },
  name:                  { type: String, required: true },
  program:               { type: Schema.Types.ObjectId, ref: "Program", required: true },
  programType:           { type: String, required: true },
  entryType:             { type: String, default: "Direct" },
  currentYearOfStudy:    { type: Number, default: 1 },
  currentSemester:       { type: Number, default: 1 },
  remarks:               { type: String },
  admissionAcademicYear: { type: Schema.Types.ObjectId, ref: "AcademicYear", required: true },
  intake:                { type: String, required: true, enum: ["JAN","MAY","SEPT"], uppercase: true, default: "SEPT" },
  status:                { type: String, default: "active" },
  qualifierSuffix:       { type: String, default: "" },

  carryForwardUnits: [{
    programUnitId:    { type: String, required: true },
    unitCode:         { type: String, required: true },
    unitName:         { type: String, required: true },
    fromYear:         { type: Number, required: true },
    fromAcademicYear: { type: String, required: true },
    attemptNumber:    { type: Number, default: 3 },
    qualifier:        { type: String, default: "RP1C" },
    addedAt:          { type: Date, default: Date.now },
    status:           { type: String, enum: ["pending","passed","failed","escalated_to_rpu"], default: "pending" },
  }],

  deferredSuppUnits: [{
    programUnitId:    { type: String, required: true },
    unitCode:         { type: String, required: true },
    unitName:         { type: String, required: true },
    fromYear:         { type: Number, required: true },
    fromAcademicYear: { type: String, required: true },
    // "supp_deferred"    = student skipped supp period, sitting in next ordinary
    // "special_deferred" = student's special exam deferred to next ordinary
    reason:  { type: String, enum: ["supp_deferred", "special_deferred"], default: "supp_deferred" },
    addedAt: { type: Date,   default: Date.now },
    status:  { type: String, enum: ["pending", "passed", "failed"], default: "pending" },
  }],

  academicLeavePeriod: { startDate: Date, endDate: Date, reason: String, type: { type: String, enum: ["compassionate","financial","other"] } },
  totalTimeOutYears:   { type: Number, default: 0 },

  unitAttemptRegistry: [{ unitId: { type: Schema.Types.ObjectId, ref: "Unit" }, attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }] }],
  academicHistory:     [{ academicYear: String, yearOfStudy: Number, annualMeanMark: Number, weightedContribution: Number, failedUnitsCount: Number, isRepeatYear: Boolean, date: Date }],
  statusHistory:       [{ status: String, previousStatus: String, date: { type: Date, default: Date.now }, reason: String }],
  statusEvents:        [{ fromStatus: String, toStatus: String, date: { type: Date, default: Date.now }, academicYear: String, reason: String }],
}, { timestamps: true });

// All indexes declared here only — no field-level index:true to avoid duplicate index warnings
schema.index({ institution: 1, regNo: 1 }, { unique: true });
schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
schema.index({ institution: 1, intake: 1 });
schema.index({ institution: 1, status: 1, currentYearOfStudy: 1 });
schema.index({ institution: 1, program: 1, currentYearOfStudy: 1, status: 1 });
schema.index({ institution: 1, admissionAcademicYear: 1, intake: 1 });
schema.index({ "statusEvents.toStatus": 1, "statusEvents.academicYear": 1 });
schema.index({ qualifierSuffix: 1 });
schema.index({ "carryForwardUnits.programUnitId": 1 });
schema.index({ "deferredSuppUnits.programUnitId": 1 });

export default mongoose.model<IStudent>("Student", schema);