// src/models/InstitutionSettings.ts → NEW & FINAL VERSION
import mongoose, { Schema, Document } from "mongoose";

export interface IInstitutionSettings extends Document {
  institution: mongoose.Types.ObjectId;

  // MAX MARKS — This is what lecturers actually decide
  cat1Max: number;        // e.g. 30
  cat2Max: number;        // e.g. 40
  cat3Max: number;        // e.g. 20 or 0
  assignmentMax: number; // e.g. 20 or 0
  practicalMax: number;  // e.g. 10 or 0

  // These are FIXED by policy
  examMax: 70;  // Always 70

  // Grading rules
  passMark: number;
  supplementaryThreshold: number;
  retakeThreshold: number;

  // Optional grading scale
  gradingScale?: Array<{
    min: number;
    grade: string;
    points?: number;
  }>;
}

const schema = new Schema<IInstitutionSettings>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },

  // MAX MARKS — Real-world values
  cat1Max: { type: Number, default: 30, min: 20, max: 50 },
  cat2Max: { type: Number, default: 40, min: 20, max: 50 },
  cat3Max: { type: Number, default: 0, min: 0, max: 40 },
  assignmentMax: { type: Number, default: 0, min: 0, max: 30 },
  practicalMax: { type: Number, default: 0, min: 0, max: 20 },

  examMax: { type: Number, default: 70, enum: [70] }, // Locked forever

  passMark: { type: Number, default: 40 },
  supplementaryThreshold: { type: Number, default: 40 },
  retakeThreshold: { type: Number, default: 5 },

  gradingScale: [{
    min: { type: Number, required: true },
    grade: { type: String, required: true },
    points: { type: Number }
  }]
}, { timestamps: true });

schema.index({ institution: 1 }, { unique: true });

export default mongoose.model<IInstitutionSettings>("InstitutionSettings", schema);