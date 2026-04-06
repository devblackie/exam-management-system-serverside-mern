// serverside/src/middleware/auth.ts
//
// UPDATED: Every user — including admins — must be linked to an institution.
// The previous version exempted admins from the institution check.
// Per the project requirement: "every user (admins included) should be linked
// to an institution."
//
// This means:
//   - Admin secret-register MUST supply an institutionId
//   - setAuthCookie MUST include institution in the JWT for all roles
//   - requireAuth blocks ANY user missing institution (no role exception)

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
  req:  Request,
  res:  Response,
  next: NextFunction,
): Promise<void> {
  const token = req.cookies?.token;

  if (!token) {
    await logAudit(req, {
      action:  "unauthenticated_access",
      details: { path: req.originalUrl },
    });
    res.status(401).json({ message: "Not authenticated" });
    return;
  }

  try {
    const payload = verifyToken(token) as {
      id:          string;
      role:        string;
      institution: string;
      version:     number;
    };

    if (!payload?.id) {
      res.status(401).json({ message: "Invalid token" });
      return;
    }

    const userDoc = await User.findById(payload.id)
      .select("-password")
      .lean();

    if (!userDoc) {
      res.clearCookie("token");
      res.status(401).json({ message: "User not found" });
      return;
    }

    if (userDoc.status === "suspended") {
      res.clearCookie("token");
      res.status(403).json({ message: "Account suspended. Contact your administrator." });
      return;
    }

    // Token version guard — invalidates sessions after password reset
    if (
      typeof payload.version === "number" &&
      payload.version !== (userDoc.tokenVersion ?? 0)
    ) {
      res.clearCookie("token");
      res.status(401).json({ message: "Session expired. Please log in again." });
      return;
    }

    // Institution guard — ALL users must have one
    if (!payload.institution) {
      await logAudit(req, {
        action:  "missing_institution_in_jwt",
        details: { userId: payload.id, role: userDoc.role },
      });
      res.status(403).json({
        message: "Account not linked to an institution. Contact a system administrator.",
      });
      return;
    }

    const safeUser: UserSafe & { institution: mongoose.Types.ObjectId } = {
      ...(userDoc as UserSafe),
      _id:         userDoc._id as mongoose.Types.ObjectId,
      institution: new mongoose.Types.ObjectId(payload.institution),
    };

    (req as AuthenticatedRequest).user = safeUser;
    next();

  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    await logAudit(req, {
      action:  "token_verification_failed",
      details: { error: message, path: req.originalUrl },
    });
    res.clearCookie("token");
    res.status(401).json({ message: "Session expired. Please log in again." });
  }
}

export function requireRole(...roles: string[]) {
  return (req: Request, res: Response, next: NextFunction): void => {
    const user = (req as AuthenticatedRequest).user;

    if (!user) {
      res.status(401).json({ message: "Not authenticated" });
      return;
    }

    // Admins bypass all role restrictions within their institution
    if (user.role === "admin") {
      next();
      return;
    }

    if (!roles.includes(user.role)) {
      res.status(403).json({ message: "Insufficient permissions for this action." });
      return;
    }

    next();
  };
}