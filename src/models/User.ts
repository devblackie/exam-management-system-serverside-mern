// serverside/src/models/User.ts

import mongoose, { Schema, Document, Model, Types } from "mongoose";

export interface IUser extends Document {
  _id: Types.ObjectId;
  name: string;
  email: string;
  password: string;
  role: "admin" | "coordinator" | "lecturer";
  status: "active" | "suspended";
  institution?: mongoose.Types.ObjectId;

  schoolCode?: string;
  departmentCode?: string;
  institutionWide: boolean;

  passwordResetToken?: string;
  passwordResetExpires?: Date;
  tokenVersion: number;
  twoFactorEnabled: boolean;
  twoFactorSecret?: string;
  twoFactorTempToken?: string;
  twoFactorTempExpires?: Date;
}

const userSchema = new Schema<IUser>(
  {
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true, lowercase: true, trim: true},
    password: { type: String, required: true, select: false },
    role: { type: String, enum: ["admin", "coordinator", "lecturer"], required: true },
    status: { type: String, enum: ["active", "suspended"], default: "active" },
    institution: { type: Schema.Types.ObjectId, ref: "Institution" },

    schoolCode: { type: String, uppercase: true, trim: true, default: null },
    departmentCode: { type: String, uppercase: true, trim: true, default: null},
    institutionWide: { type: Boolean, default: false },

    passwordResetToken: { type: String, select: false },
    passwordResetExpires: { type: Date, select: false },
    tokenVersion: { type: Number, default: 0, required: true },
    twoFactorEnabled: { type: Boolean, default: false },
    twoFactorSecret: { type: String, select: false },
    twoFactorTempToken: { type: String, select: false },
    twoFactorTempExpires: { type: Date },
  },
  { timestamps: true },
);

// ── Compound indexes only — single-field uniqueness is on the field definitions ──
userSchema.index({ institution: 1, role: 1 });
userSchema.index({ institution: 1, departmentCode: 1 });

const User: Model<IUser> = mongoose.model<IUser>("User", userSchema);
export default User;
