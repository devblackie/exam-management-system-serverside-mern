// src/models/Program.ts
import mongoose, { Document, Schema } from "mongoose";
import { normalizeProgramName } from "../services/programNormalizer";

export interface IProgram extends Document {
  institution: mongoose.Types.ObjectId;
  name: string;
  cleanName: string; // normalized version
  code: string;
  durationYears: number;
  maxCompletionYears: number; // ENG 19.d/e & 22.f/g
  schoolType: "ENGINEERING" | "IT" | "MEDICINE" | "GENERAL";
}

const schema = new Schema<IProgram>(
  {
    institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
    name: { type: String, required: true, trim: true},
    cleanName: { type: String, required: false, index: true },
    code: { type: String, required: true, uppercase: true, trim: true },    
    durationYears: {type: Number, default: 4, min: 1, max: 8 },
    maxCompletionYears: { type: Number, default: 8 }, // Default for B.Ed/General
    schoolType: { type: String, enum: ["ENGINEERING", "IT", "MEDICINE", "GENERAL"], default: "GENERAL" }
  },
  { timestamps: true }
);

schema.pre("save", function (next) {
  if (this.isModified("name")) this.cleanName = normalizeProgramName(this.name);  
  // Auto-set max years based on ENG 19/22
    if (this.name.includes("Engineering")) this.maxCompletionYears = 10;
  next();
});

// Before updateOne / findOneAndUpdate
function applyNormalization(this: any, next: any) {
  const update = this.getUpdate();
  if (update?.name) update.cleanName = normalizeProgramName(update.name);  
  if (update?.$set?.name) update.$set.cleanName = normalizeProgramName(update.$set.name);
  next();
}

schema.pre("findOneAndUpdate", applyNormalization);
schema.pre("updateOne", applyNormalization);

schema.index({ institution: 1, code: 1 }, { unique: true });
schema.index({ institution: 1, cleanName: 1 }, { unique: true }); 
schema.index({ code: 1 });

export default mongoose.model<IProgram>("Program", schema);
