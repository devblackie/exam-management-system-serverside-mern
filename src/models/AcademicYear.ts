// serverside/src/models/AcademicYear.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IAcademicYear extends Document {
  institution: mongoose.Types.ObjectId;
  year: string;                   // e.g. "2024/2025"
  intakes: string[];              // e.g. ["JAN"]
  startDate: Date;
  endDate: Date;
  isCurrent: boolean;
  // Session control for ENG 13/18
  session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
  isRegistrationOpen: boolean;    // Unit registration window
}

const schema = new Schema<IAcademicYear>({
  institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
  year: { type: String, required: true },
  intakes: [{ type: String }],
  startDate: { type: Date, required: true },
  endDate: { type: Date, required: true },
  isCurrent: { type: Boolean, default: false },
  session: { 
    type: String, 
    enum: ["ORDINARY", "SUPPLEMENTARY", "CLOSED"], 
    default: "ORDINARY" 
  },
  isRegistrationOpen: { type: Boolean, default: true }
});

schema.index({ institution: 1, year: 1 }, { unique: true });

export default mongoose.model<IAcademicYear>("AcademicYear", schema);