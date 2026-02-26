// src/models/InstitutionSettings.ts 
import mongoose, { Schema, Document } from "mongoose";

export interface IInstitutionSettings extends Document {
  institution: mongoose.Types.ObjectId;
  cat1Max: number; cat2Max: number; cat3Max: number;
  assignmentMax: number; practicalMax: number; workshopMax: number;
  examMax: 70; passMark: number; unitType:string;
  gradingScale?: Array<{ min: number; grade: string; points?: number }>;
}
const schema = new Schema<IInstitutionSettings>(
  {
    institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },
    cat1Max: { type: Number, default: 20 },
    cat2Max: { type: Number, default: 20 },
    cat3Max: { type: Number, default: 0 },
    assignmentMax: { type: Number, default: 10 },
    practicalMax: { type: Number, default: 10 },
    workshopMax: { type: Number, default: 100 },
    examMax: { type: Number, default: 70, enum: [70] },
    passMark: { type: Number, default: 40 },
    unitType: { type: String, enum: ["theory", "lab", "workshop"], default: "theory" },
    gradingScale: [ { min: { type: Number, required: true }, grade: { type: String, required: true }, points: { type: Number }}],
  },
  { timestamps: true },
);
schema.index({ institution: 1 }, { unique: true });
export default mongoose.model<IInstitutionSettings>("InstitutionSettings", schema);