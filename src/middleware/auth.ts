// src/middleware/auth.ts
import { Request, Response, NextFunction } from "express";
import { verifyToken } from "../lib/jwt";
import User from "../models/User";
import { logAudit } from "../lib/auditLogger";
import type { UserSafe } from "../types/express";
import mongoose from "mongoose";

export interface AuthenticatedRequest extends Request {
  user: UserSafe & { institution: mongoose.Types.ObjectId };
}

export async function requireAuth(
  req: Request,
  res: Response,
  next: NextFunction
): Promise<void> {
  const token = req.cookies?.token;

  if (!token) {
    await logAudit(req, { action: "unauthenticated_access", details: { path: req.originalUrl } });
    res.status(401).json({ message: "Not authenticated" });
    return;
  }

  try {
    const payload = verifyToken(token) as { id: string; role: string; institution?: string };

    if (!payload?.id) {
      res.status(401).json({ message: "Invalid token" });
      return;
    }

    const userDoc = await User.findById(payload.id).select("-password").lean();
    if (!userDoc) {
      res.status(401).json({ message: "User not found" });
      return;
    }

    if (userDoc.status === "suspended") {
      res.status(403).json({ message: "Account suspended" });
      return;
    }

    // CRITICAL: institution MUST come from JWT (not DB) — most reliable
    if (!payload.institution) {
      await logAudit(req, { action: "missing_institution_in_jwt", details: { userId: payload.id } });
      res.status(403).json({ message: "User not linked to institution" });
      return;
    }

    const safeUser: UserSafe & { institution: mongoose.Types.ObjectId } = {
      ...userDoc,
      _id: userDoc._id as mongoose.Types.ObjectId,
      institution: new mongoose.Types.ObjectId(payload.institution),
    };

    req.user = safeUser;
    (req as AuthenticatedRequest).user = safeUser; // Type assertion

    next();
  } catch (err: any) {
    await logAudit(req, { action: "token_verification_failed", details: { error: err.message } });
    res.status(401).json({ message: "Invalid token" });
  }
}

export function requireRole(...roles: string[]) {
  return (req: Request, res: Response, next: NextFunction) => {
    const user = (req as AuthenticatedRequest).user;

    if (!user) {
      res.status(401).json({ message: "Not authenticated" });
      return;
    }

    if (user.role === "admin") return next();

    if (!roles.includes(user.role)) {
      res.status(403).json({ message: "Forbidden: insufficient role" });
      return;
    }

    next();
  };
}