// src/models/Institution.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IInstitution extends Document {
  name: string;
  shortName?: string;
  code: string;                   // e.g. "CE", for Dept. of Civil Engineering
  isActive: boolean;
  academicYearStartMonth?: number; // 1–12, e.g. 9 = September intake
}

const schema = new Schema<IInstitution>({
  name: { type: String, required: true },
  shortName: { type: String, default: "" },
  code: { type: String, required: true, unique: true, uppercase: true },
  isActive: { type: Boolean, default: true },
  academicYearStartMonth: { type: Number, min: 1, max: 12, default: 9 },
}, { timestamps: true });

export default mongoose.model<IInstitution>("Institution", schema);