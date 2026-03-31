// serverside/src/models/User.ts
import mongoose, { Schema, Document, Model, Types } from "mongoose";

export interface IUser extends Document {
  _id: Types.ObjectId; // explicitly declare _id type
  name: string;
  email: string;
  password: string;
  role: "admin" | "coordinator" | "lecturer";
  status: "active" | "suspended";
  institution?: mongoose.Types.ObjectId;
  passwordResetToken?: string;
  passwordResetExpires?: Date;
  tokenVersion: number;
  twoFactorEnabled: boolean;
  twoFactorSecret: string;
  twoFactorTempToken: string;
  twoFactorTempExpires: Date;
}

const userSchema = new Schema<IUser>(
  {
    name: { type: String, required: true },
    email: { type: String, required: true, unique: true },
    password: { type: String, required: true },
    role: { type: String, enum: ["admin", "coordinator", "lecturer"], required: true },
    status: { type: String, enum: ["active", "suspended"], default: "active" },
    institution: { type: Schema.Types.ObjectId, ref: "Institution", required: false },
    passwordResetToken: { type: String, select: false },
    passwordResetExpires: { type: Date, select: false },
    tokenVersion: { type: Number, default: 0,required: true },

    // Add to existing User schema fields:
twoFactorEnabled:   { type: Boolean, default: false },
twoFactorSecret:    { type: String, select: false },   // TOTP secret (base32)
twoFactorTempToken: { type: String, select: false },   // short-lived token after password verify
twoFactorTempExpires: { type: Date },
  },

  { timestamps: true },
);

const User: Model<IUser> = mongoose.model<IUser>("User", userSchema);

export default User;
