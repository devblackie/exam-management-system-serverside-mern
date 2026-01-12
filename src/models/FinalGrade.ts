// src/models/FinalGrade.ts
import mongoose, { Schema, Document, Model } from "mongoose";

export interface IFinalGrade extends Document {
  student: mongoose.Types.ObjectId;
  programUnit: mongoose.Types.ObjectId;
  academicYear: mongoose.Types.ObjectId;
  institution: mongoose.Types.ObjectId;
  semester: string;

  totalMark: number;
  grade: string;
  points?: number;
  status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "INCOMPLETE";
  cappedBecauseSupplementary: boolean;
}

const schema = new Schema<IFinalGrade>(
  {
    student: { type: Schema.Types.ObjectId, ref: "Student", required: true },
    programUnit: { type: Schema.Types.ObjectId, ref: "ProgramUnit", required: true },
    academicYear: {
      type: Schema.Types.ObjectId,
      ref: "AcademicYear",
      required: true,
    },
    institution: {
      type: Schema.Types.ObjectId,
      ref: "Institution",
      required: true,
    },
    semester: {
      type: Schema.Types.Mixed,
      enum: ["SEMESTER 1", "SEMESTER 2", "SEMESTER 3"], // or Number: 1, 2, 3
      required: true,
    },

    totalMark: { type: Number, required: true },
    grade: { type: String, required: true },
    points: Number,
    status: {
      type: String,
      enum: ["PASS", "SUPPLEMENTARY", "RETAKE", "INCOMPLETE"],
      required: true,
    },
    cappedBecauseSupplementary: { type: Boolean, default: false },
  },
  { timestamps: true }
);

schema.index({ student: 1, academicYear: 1 });
schema.index({ institution: 1, academicYear: 1, status: 1 });

export default mongoose.model<IFinalGrade>("FinalGrade", schema);
