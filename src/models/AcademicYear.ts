// src/models/AcademicYear.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IAcademicYear extends Document {
  institution: mongoose.Types.ObjectId;
  year: string;                   // e.g. "2024/2025"
  startDate: Date;
  endDate: Date;
  isCurrent: boolean;
}

const schema = new Schema<IAcademicYear>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
  year: { type: String, required: true },
  startDate: { type: Date, required: true },
  endDate: { type: Date, required: true },
  isCurrent: { type: Boolean, default: false }
});

schema.index({ institution: 1, year: 1 }, { unique: true });
schema.index({ institution: 1, isCurrent: 1 });

export default mongoose.model<IAcademicYear>("AcademicYear", schema);