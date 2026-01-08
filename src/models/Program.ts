// src/models/Program.ts
import mongoose, { Document, Schema } from "mongoose";
import { normalizeProgramName } from "../services/programNormalizer";

export interface IProgram extends Document {
  institution: mongoose.Types.ObjectId;
  name: string;
  cleanName: string; // normalized version
  code: string;
  description?: string;
  durationYears: number;
}

const schema = new Schema<IProgram>(
  {
    institution: {
      type: Schema.Types.ObjectId,
      ref: "Institution",
      required: true,
    },
    name: {
      type: String,
      required: true,
      trim: true,
    },

    // AUTO-NORMALIZED NAME (BSc → bachelor of science etc.)
    cleanName: {
      type: String,
      required: false,
      index: true,
    },

    code: {
      type: String,
      required: true,
      uppercase: true,
      trim: true,
    },

    description: String,

    durationYears: {
      type: Number,
      default: 4,
      min: 1,
      max: 8,
    },
  },
  { timestamps: true }
);

// -----------------------------
// 🔥 AUTO NORMALIZATION HOOKS
// -----------------------------

// Before save: normalize name → cleanName
schema.pre("save", function (next) {
  if (this.isModified("name")) {
    this.cleanName = normalizeProgramName(this.name);
  }
  next();
});

// Before updateOne / findOneAndUpdate
function applyNormalization(this: any, next: any) {
  const update = this.getUpdate();

  if (update?.name) {
    update.cleanName = normalizeProgramName(update.name);
  }
  if (update?.$set?.name) {
    update.$set.cleanName = normalizeProgramName(update.$set.name);
  }

  next();
}

schema.pre("findOneAndUpdate", applyNormalization);
schema.pre("updateOne", applyNormalization);

// -----------------------------
// Indexes
// -----------------------------
schema.index({ institution: 1, code: 1 }, { unique: true });
schema.index({ institution: 1, cleanName: 1 }, { unique: true }); // 🔥 match using normalized name
schema.index({ code: 1 });

export default mongoose.model<IProgram>("Program", schema);
