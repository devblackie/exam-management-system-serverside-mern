// src/models/Student.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IStudent extends Document {
  institution: mongoose.Types.ObjectId;
  regNo: string;
  name: string;
  program: mongoose.Types.ObjectId;
  programType: string; // "B.Sc", "B.Ed", "Diploma"
  entryType: "Direct" | "Mid-Entry-Y2" | "Mid-Entry-Y3" | "Mid-Entry-Y4";
  currentYearOfStudy: number;
  remarks?: string;
  currentSemester: number;
  status: "active" | "graduated" | "discontinued" | "deregistered" | "repeat" | "on_leave" | "deferred";
  admissionAcademicYear: mongoose.Types.ObjectId;
  intake: string;
  academicLeavePeriod?: { startDate: Date; endDate: Date; reason: string; type: "compassionate" | "financial" | "other"; };
  totalTimeOutYears: number; // ENG 19.d/e: Total years allowed out
  // ENG 22.b: Track attempts per unit (Limit: 5)
  unitAttemptRegistry: { unitId: mongoose.Types.ObjectId; attempts: { attemptNumber: number; mark: number; passed: boolean; type: string; }[]; }[];

  // ENG 25.b: Store snapshots of each year for final classification
  academicHistory: { academicYear: string; yearOfStudy: number; annualMeanMark: number; weightedContribution: number; failedUnitsCount: number; isRepeatYear?: boolean; date?: Date; }[];
  statusEvents: { fromStatus: string; toStatus: string; date: Date; academicYear: string; reason: string; }[];
}

const schema = new Schema<IStudent>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
  regNo: { type: String, required: true, uppercase: true },
  name: { type: String, required: true },
  program: { type: Schema.Types.ObjectId, ref: "Program", required: true },
  programType: { type: String, required: true },
  entryType: { type: String, default: "Direct" },
  currentYearOfStudy: { type: Number, default: 1 },
  currentSemester: { type: Number, default: 1 },
  admissionAcademicYear: { type: Schema.Types.ObjectId, ref: "AcademicYear", required: true },
  intake: { type: String, required: true, enum: ["JAN", "MAY", "SEPT"], uppercase: true, default: "SEPT" },
  status: { type: String, default: "active" },
  academicLeavePeriod: { startDate: Date, endDate: Date, reason: String, type: { type: String, enum: ["compassionate", "financial", "other"] }},
  totalTimeOutYears: { type: Number, default: 0 },
  unitAttemptRegistry: [{ unitId: { type: Schema.Types.ObjectId, ref: "Unit" }, attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }]}],
  academicHistory: [{ academicYear: String, yearOfStudy: Number, annualMeanMark: Number, weightedContribution: Number, failedUnitsCount: Number, isRepeatYear: Boolean, date: Date }],
  statusEvents: [{ fromStatus: String, toStatus: String, date: { type: Date, default: Date.now }, academicYear: String, reason: String }],
}, { timestamps: true });

schema.index({ institution: 1, regNo: 1 }, { unique: true });
// schema.index({ regNo: 1 });
schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
schema.index({ institution: 1, intake: 1 });
export default mongoose.model<IStudent>("Student", schema);

// set up the Mongoose Indexes to optimize those multi-tenant queries?


