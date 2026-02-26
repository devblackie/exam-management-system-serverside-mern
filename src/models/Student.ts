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
  currentSemester: number;
  status: "active" | "graduated" | "discontinued" | "deregistered" | "repeat";
  admissionAcademicYear: mongoose.Types.ObjectId;

  // ENG 22.b: Track attempts per unit (Limit: 5)
  unitAttemptRegistry: {
    unitId: mongoose.Types.ObjectId;
    attempts: {
      attemptNumber: number;
      mark: number;
      passed: boolean;
      type: string;
    }[];
  }[];

  // ENG 25.b: Store snapshots of each year for final classification
  academicHistory: {
    yearOfStudy: number;
    annualMeanMark: number;
    weightedContribution: number;
    failedUnitsCount: number;
  }[];
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
  status: { type: String, default: "active" },
  unitAttemptRegistry: [{
    unitId: { type: Schema.Types.ObjectId, ref: "Unit" },
    attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }]
  }],
  academicHistory: [{
    yearOfStudy: Number,
    annualMeanMark: Number,
    weightedContribution: Number,
    failedUnitsCount: Number
  }]
}, { timestamps: true });

schema.index({ institution: 1, regNo: 1 }, { unique: true });
schema.index({ regNo: 1 });
schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
export default mongoose.model<IStudent>("Student", schema);

// set up the Mongoose Indexes to optimize those multi-tenant queries?


