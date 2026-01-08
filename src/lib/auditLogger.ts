// lib/auditLogger.ts
import AuditLog from "../models/AuditLog";
import { Request } from "express";
import mongoose from "mongoose";

export async function logAudit(
  req: Request,
  {
    action,
    actor,
    targetUser,
    details = {},
  }: {
    action: string;
    actor?: mongoose.Types.ObjectId;        // optional
    targetUser?: mongoose.Types.ObjectId;   // optional
    details?: Record<string, unknown>;
  }
) {
  const ip =
    (req.headers["x-forwarded-for"] as string) ||
    req.socket.remoteAddress ||
    req.ip;
  const userAgent = req.headers["user-agent"];

  // Explicit partial type so actor/targetUser can be deleted safely
  const logEntry: Partial<{
    action: string;
    actor: mongoose.Types.ObjectId;
    targetUser: mongoose.Types.ObjectId;
    details: Record<string, unknown>;
    ip?: string;
    userAgent?: string;
  }> = {
    action,
    actor: actor || (req.user?._id as mongoose.Types.ObjectId),
    targetUser: targetUser || actor || (req.user?._id as mongoose.Types.ObjectId),
    details,
    ip,
    userAgent,
  };

  // Ensure empty fields don’t cause schema validation errors
  if (!logEntry.actor) delete logEntry.actor;
  if (!logEntry.targetUser) delete logEntry.targetUser;

  return Promise.resolve(
    AuditLog.create(logEntry).catch((err) =>
      console.error("Audit log failed:", err)
    )
  );
}
