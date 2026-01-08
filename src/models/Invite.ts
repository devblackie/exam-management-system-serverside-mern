import { Schema, model, Document } from "mongoose";

export interface IInvite extends Document {
  name: string;
  email: string;
  token: string;
  role: "coordinator" | "lecturer";
  used: boolean;
  expiresAt: Date;
  createdBy?: string; // admin id
}

const inviteSchema = new Schema<IInvite>(
  {
    name: { type: String, required: true },
    email: { type: String, required: true },
    token: { type: String, required: true, unique: true },
    role: { type: String, enum: ["coordinator", "lecturer"], required: true },
    used: { type: Boolean, default: false },
    expiresAt: { type: Date, required: true },
    createdBy: { type: Schema.Types.ObjectId, ref: "User" },
  },
  { timestamps: true }
);

export default model<IInvite>("Invite", inviteSchema);
