// src/models/Unit.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IUnit extends Document {
  institution: mongoose.Types.ObjectId;
  code: string;
  name: string;
}

const schema = new Schema<IUnit>(
  {
    institution: {
      type: Schema.Types.ObjectId,
      ref: "Institution",
      required: true,
    },
    code: { type: String, required: true, uppercase: true, trim: true },
    name: { type: String, required: true, trim: true },
  },
  { timestamps: true }
);

// Indexes simplified to focus on institution/code uniqueness
schema.index({ institution: 1, code: 1 }, { unique: true });

export default mongoose.model<IUnit>("Unit", schema);
