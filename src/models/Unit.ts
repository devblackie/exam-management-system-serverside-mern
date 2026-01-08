// // src/models/Unit.ts
// import mongoose, { Schema, Document } from "mongoose";

// export interface IUnit extends Document {
//   institution: mongoose.Types.ObjectId;
//   program: mongoose.Types.ObjectId;
//   code: string;
//   name: string;
//   year: number;           // Year 1, 2, 3, 4
//   semester: 1 | 2;
// }

// const schema = new Schema<IUnit>(
//   {
//     institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
//     program: { type: Schema.Types.ObjectId, ref: "Program", required: true },
//     code: { type: String, required: true, uppercase: true, trim: true },
//     name: { type: String, required: true, trim: true },
//     year: { type: Number, required: true, min: 1, max: 6 },
//     semester: { type: Number, required: true, enum: [1, 2] },
//   },
//   { timestamps: true }
// );

// schema.index({ institution: 1, code: 1 }, { unique: true });
// schema.index({ program: 1, year: 1, semester: 1 });        // ← FIXED: use 'year'
// schema.index({ program: 1, code: 1 });                    // ← Bonus: fast lookup by code

// export default mongoose.model<IUnit>("Unit", schema);

// src/models/Unit.ts (UPDATED)
import mongoose, { Schema, Document } from "mongoose";

export interface IUnit extends Document {
  institution: mongoose.Types.ObjectId;
  code: string;
  name: string;
  // ❌ REMOVED: program: mongoose.Types.ObjectId;
  // ❌ REMOVED: year: number;
  // ❌ REMOVED: semester: 1 | 2;
}

const schema = new Schema<IUnit>(
  {
    institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
    code: { type: String, required: true, uppercase: true, trim: true },
    name: { type: String, required: true, trim: true },
    // ❌ REMOVED fields here
  },
  { timestamps: true }
);

// Indexes simplified to focus on institution/code uniqueness
schema.index({ institution: 1, code: 1 }, { unique: true });

export default mongoose.model<IUnit>("Unit", schema);