// src/models/Student
import mongoose, { Schema, Document } from "mongoose";

export interface IStudent extends Document {
  institution: mongoose.Types.ObjectId;
  regNo: string;                   // e.g. "SC/ICT/001/2023"
  name: string;
  email?: string;
  phone?: string;
  program: mongoose.Types.ObjectId;
  currentYearOfStudy: number;
  currentSemester: 1 | 2 | 3;
  // admissionAcademicYear: string;   // e.g. "2023/2024"
  admissionAcademicYear: mongoose.Types.ObjectId;
  status: "active" | "inactive" | "graduated" | "suspended" | "deferred";
}

const schema = new Schema<IStudent>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true ,index: true},
  regNo: { type: String, required: true, uppercase: true, trim: true },
  name: { type: String, required: true, trim: true },
  email: { type: String, lowercase: true, trim: true },
  phone: String,
  program: { type: Schema.Types.ObjectId, ref: "Program", required: true },
  currentYearOfStudy: { type: Number, min: 1, default: 1 },
  currentSemester: { type: Number, enum: [1, 2, 3], default: 1 },
  // admissionAcademicYear: { type: String, required: true }, // e.g. "2023/2024"
  admissionAcademicYear: { 
    type: Schema.Types.ObjectId, 
    ref: "AcademicYear", 
    required: true // Now it must point to a valid year
  },
  status: {
    type: String,
    enum: ["active", "inactive", "graduated", "suspended", "deferred"],
    default: "active"
  },
}, { timestamps: true });

// Critical: regNo must be unique per institution
schema.index({ institution: 1, regNo: 1 }, { unique: true });
// Fast search by regNo (most common lookup)
schema.index({ regNo: 1 });
// For student history reports
schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });

export default mongoose.model<IStudent>("Student", schema);