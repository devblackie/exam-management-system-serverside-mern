// models/AuditLog.ts
import mongoose, { Schema, Document } from "mongoose";
import { IUser } from "./User"; // <-- import your User interface

export interface IAuditLog extends Document {
  action: string;
  targetUser?: mongoose.Types.ObjectId | IUser; // can be populated User
  actor?: mongoose.Types.ObjectId | IUser;       // can be populated User
  details?: Record<string, unknown>;
   ip?: string;
  userAgent?: string;
  createdAt: Date;
}

const AuditLogSchema = new Schema<IAuditLog>(
  {
    action: { type: String, required: true },
    targetUser: { type: Schema.Types.ObjectId, ref: "User" },
    actor: { type: Schema.Types.ObjectId, ref: "User" },
    details: { type: Schema.Types.Mixed, default: {} },
    ip: { type: String },
    userAgent: { type: String },
  },
  { timestamps: { createdAt: true, updatedAt: false } }
);

export default mongoose.model<IAuditLog>("AuditLog", AuditLogSchema);
