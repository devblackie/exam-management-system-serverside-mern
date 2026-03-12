// src/models/MarkDirect.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IMarkDirect extends Document {
  institution: mongoose.Types.ObjectId;
  student: mongoose.Types.ObjectId;
  programUnit: mongoose.Types.ObjectId;
  academicYear: mongoose.Types.ObjectId;
  semester: string;
  caTotal30: number;    // Input directly: 0-30
  examTotal70: number;  // Input directly: 0-70
  externalTotal100?: number;  // Calculated: (examTotal70 / 70) * 100
  agreedMark: number;   // Calculated: CA + Exam
  attempt: string;

  isSupplementary: boolean; // Derived from attempt field
  isRetake: boolean; // Derived from attempt field
  isSpecial: boolean; // Special exam flag
  isMissingCA: boolean; // Missing CA flag
  remarks?: string; // Any remarks or notes
  
  uploadedBy: mongoose.Types.ObjectId;
  uploadedAt: Date; 
  deletedAt?: Date;
}

const schema = new Schema<IMarkDirect>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
  student: { type: Schema.Types.ObjectId, ref: "Student", required: true },
  programUnit: { type: Schema.Types.ObjectId, ref: "ProgramUnit", required: true },
  academicYear: { type: Schema.Types.ObjectId, ref: "AcademicYear", required: true },
  semester: { type: String, required: true },
  caTotal30: { type: Number, min: 0, max: 30, required: true },
  examTotal70: { type: Number, min: 0, max: 70, required: true },
  externalTotal100: { type: Number, min: 0, max: 100, default: null },
  agreedMark: { type: Number, min: 0, max: 100, required: true },
  attempt: { type: String, enum: ["1st", "re-take", "supplementary", "special"], default: "1st" },
  isSupplementary: { type: Boolean, default: false },
  isRetake: { type: Boolean, default: false },
  isSpecial: { type: Boolean, default: false }, // Explicit flag for Special Exams
  remarks: { type: String },
  uploadedBy: { type: Schema.Types.ObjectId, ref: "User", required: true },
  uploadedAt: { type: Date, default: Date.now },
  deletedAt: { type: Date, default: null },
}, { timestamps: true });

schema.index({ student: 1, programUnit: 1, academicYear: 1 }, { unique: true });
schema.index({ student: 1, programUnit: 1, academicYear: 1, deletedAt: 1 });

schema.pre(/^find/, function (this: mongoose.Query<any, any>, next) {
  if (!this.getQuery().deletedAt) {
    this.where({ deletedAt: null });
  }
  next();
});

export default mongoose.model<IMarkDirect>("MarkDirect", schema);