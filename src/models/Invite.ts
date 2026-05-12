import mongoose, { Schema, model, Document } from "mongoose";

export interface IInvite extends Document {
  institution:           mongoose.Types.ObjectId;
  name: string;
  email: string;
  schoolCode: string;
  departmentCode: string;
  institutionWide: boolean;
  token: string;
  role: "coordinator" | "lecturer";
  used: boolean;
  expiresAt: Date;
  createdBy?: string; // admin id
}

const inviteSchema = new Schema<IInvite>(
  {
    institution: { type: Schema.Types.ObjectId, ref: "Institution", required: true },
    name: { type: String, required: true },
    email: { type: String, required: true },
    schoolCode: { type: String, default: null },
    departmentCode: { type: String, default: null },
    institutionWide: { type: Boolean, default: false },
    token: { type: String, required: true, unique: true },
    role: { type: String, enum: ["coordinator", "lecturer"], required: true },
    used: { type: Boolean, default: false },
    expiresAt: { type: Date, required: true },
    createdBy: { type: Schema.Types.ObjectId, ref: "User" },
  },
  { timestamps: true },
);

inviteSchema.index({ institution: 1, email: 1 });
inviteSchema.index({ token: 1 });
inviteSchema.index({ expiresAt: 1 });

export default model<IInvite>("Invite", inviteSchema);
