// // serverside/src/middleware/auth.ts
// //
// // UPDATED: Every user — including admins — must be linked to an institution.
// // The previous version exempted admins from the institution check.
// // Per the project requirement: "every user (admins included) should be linked
// // to an institution."
// //
// // This means:
// //   - Admin secret-register MUST supply an institutionId
// //   - setAuthCookie MUST include institution in the JWT for all roles
// //   - requireAuth blocks ANY user missing institution (no role exception)

// import { Request, Response, NextFunction } from "express";
// import { verifyToken } from "../lib/jwt";
// import User from "../models/User";
// import { logAudit } from "../lib/auditLogger";
import type { UserSafe } from "../types/express";
// import mongoose from "mongoose";

// export interface AuthenticatedRequest extends Request {
//   user: UserSafe & { institution: mongoose.Types.ObjectId };
// }

// export async function requireAuth(
//   req:  Request,
//   res:  Response,
//   next: NextFunction,
// ): Promise<void> {
//   const token = req.cookies?.token;

//   if (!token) {
//     await logAudit(req, {
//       action:  "unauthenticated_access",
//       details: { path: req.originalUrl },
//     });
//     res.status(401).json({ message: "Not authenticated" });
//     return;
//   }

//   try {
//     const payload = verifyToken(token) as {
//       id:          string;
//       role:        string;
//       institution: string;
//       version:     number;
//     };

//     if (!payload?.id) {
//       res.status(401).json({ message: "Invalid token" });
//       return;
//     }

//     const userDoc = await User.findById(payload.id)
//       .select("-password")
//       .lean();

//     if (!userDoc) {
//       res.clearCookie("token");
//       res.status(401).json({ message: "User not found" });
//       return;
//     }

//     if (userDoc.status === "suspended") {
//       res.clearCookie("token");
//       res.status(403).json({ message: "Account suspended. Contact your administrator." });
//       return;
//     }

//     // Token version guard — invalidates sessions after password reset
//     if (
//       typeof payload.version === "number" &&
//       payload.version !== (userDoc.tokenVersion ?? 0)
//     ) {
//       res.clearCookie("token");
//       res.status(401).json({ message: "Session expired. Please log in again." });
//       return;
//     }

//     // Institution guard — ALL users must have one
//     if (!payload.institution) {
//       await logAudit(req, {
//         action:  "missing_institution_in_jwt",
//         details: { userId: payload.id, role: userDoc.role },
//       });
//       res.status(403).json({
//         message: "Account not linked to an institution. Contact a system administrator.",
//       });
//       return;
//     }

//     const safeUser: UserSafe & { institution: mongoose.Types.ObjectId } = {
//       ...(userDoc as UserSafe),
//       _id:         userDoc._id as mongoose.Types.ObjectId,
//       institution: new mongoose.Types.ObjectId(payload.institution),
//     };

//     (req as AuthenticatedRequest).user = safeUser;
//     next();

//   } catch (err: unknown) {
//     const message = err instanceof Error ? err.message : "Unknown error";
//     await logAudit(req, {
//       action:  "token_verification_failed",
//       details: { error: message, path: req.originalUrl },
//     });
//     res.clearCookie("token");
//     res.status(401).json({ message: "Session expired. Please log in again." });
//   }
// }

// export function requireRole(...roles: string[]) {
//   return (req: Request, res: Response, next: NextFunction): void => {
//     const user = (req as AuthenticatedRequest).user;

//     if (!user) {
//       res.status(401).json({ message: "Not authenticated" });
//       return;
//     }

//     // Admins bypass all role restrictions within their institution
//     if (user.role === "admin") {
//       next();
//       return;
//     }

//     if (!roles.includes(user.role)) {
//       res.status(403).json({ message: "Insufficient permissions for this action." });
//       return;
//     }

//     next();
//   };
// }












// // serverside/src/middleware/auth.ts
// import { Request, Response, NextFunction } from "express";
// import { verifyToken }   from "../lib/jwt";
// import User              from "../models/User";
// import Program           from "../models/Program";
// import mongoose          from "mongoose";
// import { logAudit }      from "../lib/auditLogger";

// export interface ScopedUser {
//   _id:             mongoose.Types.ObjectId;
//   name:            string;
//   email:           string;
//   role:            "admin" | "coordinator" | "lecturer";
//   status:          "active" | "suspended";
//   institution:     mongoose.Types.ObjectId;
//   schoolCode:      string | null;
//   departmentCode:  string | null;
//   institutionWide: boolean;
//   tokenVersion:    number;
// }

// export interface AuthenticatedRequest extends Request {
//   user: ScopedUser;
// }

// export async function requireAuth(
//   req:  Request,
//   res:  Response,
//   next: NextFunction,
// ): Promise<void> {
//   const token = (req as Request & { cookies?: Record<string, string> }).cookies?.token;

//   if (!token) {
//     await logAudit(req, { action: "unauthenticated_access", details: { path: req.originalUrl } });
//     res.status(401).json({ message: "Not authenticated" });
//     return;
//   }

//   try {
//     const payload = verifyToken(token) as {
//       id:          string;
//       role:        string;
//       institution: string;
//       version:     number;
//     };

//     if (!payload?.id) {
//       res.status(401).json({ message: "Invalid token" }); return;
//     }

//     const userDoc = await User.findById(payload.id)
//       .select("name email role status institution tokenVersion schoolCode departmentCode institutionWide")
//       .lean() as ScopedUser | null;

//     if (!userDoc) {
//       res.clearCookie("token");
//       res.status(401).json({ message: "User not found" }); return;
//     }

//     if (userDoc.status === "suspended") {
//       res.clearCookie("token");
//       res.status(403).json({ message: "Account suspended. Contact your administrator." }); return;
//     }

//     if (typeof payload.version === "number" && payload.version !== (userDoc.tokenVersion ?? 0)) {
//       res.clearCookie("token");
//       res.status(401).json({ message: "Session expired. Please log in again." }); return;
//     }

//     if (!payload.institution) {
//       await logAudit(req, {
//         action:  "missing_institution_in_jwt",
//         details: { userId: payload.id, role: userDoc.role },
//       });
//       res.status(403).json({
//         message: "Account not linked to an institution. Contact a system administrator.",
//       });
//       return;
//     }

//     (req as AuthenticatedRequest).user = {
//       ...userDoc,
//       _id:         userDoc._id as mongoose.Types.ObjectId,
//       institution: new mongoose.Types.ObjectId(payload.institution),
//     };

//     next();
//   } catch (err: unknown) {
//     const message = err instanceof Error ? err.message : "Unknown error";
//     await logAudit(req, {
//       action:  "token_verification_failed",
//       details: { error: message, path: req.originalUrl },
//     });
//     res.clearCookie("token");
//     res.status(401).json({ message: "Session expired. Please log in again." });
//   }
// }

// export function requireRole(...roles: string[]) {
//   return (req: Request, res: Response, next: NextFunction): void => {
//     const user = (req as AuthenticatedRequest).user;
//     if (!user) { res.status(401).json({ message: "Not authenticated" }); return; }
//     if (user.role === "admin") { next(); return; }  // admins bypass role checks
//     if (!roles.includes(user.role)) {
//       res.status(403).json({ message: "Insufficient permissions for this action." }); return;
//     }
//     next();
//   };
// }

// /**
//  * Returns the list of program ObjectId strings this user is allowed to see.
//  * - institutionWide admins/coordinators: all programs in institution
//  * - scoped coordinators: only programs in their departmentCode
//  */
// export async function getScopedProgramIds(req: AuthenticatedRequest): Promise<string[]> {
//   const filter: Record<string, unknown> = { institution: req.user.institution };

//   if (!req.user.institutionWide) {
//     if (req.user.departmentCode) filter.departmentCode = req.user.departmentCode;
//     if (req.user.schoolCode)     filter.schoolCode     = req.user.schoolCode;
//   }

//   const programs = await Program.find(filter).select("_id").lean() as Array<{ _id: mongoose.Types.ObjectId }>;
//   return programs.map(p => p._id.toString());
// }













// serverside/src/middleware/auth.ts
//
// CHANGES vs previous version
// ────────────────────────────
// 1. ScopedUser — lean POJO, never a Mongoose Document. Kept as-is; it is correct.
//
// 2. AuthenticatedRequest — exported as a plain interface extending express.Request.
//    It is NOT merged into the global Express namespace (that lives in express.d.ts).
//    Merging it there caused:
//      "ScopedUser is not assignable to UserSafe"
//      "ScopedUser is missing $assertPopulated, $clearModifiedPaths …"
//    because the global augmentation demanded the full Mongoose Document shape.
//
// 3. getScopedProgramIds — fixed the .lean() cast.
//    Mongoose .lean() returns FlattenMaps<T>, not T directly.
//    FlattenMaps<IProgram>._id is FlattenMaps<unknown>, NOT ObjectId.
//    Fix: cast through `unknown` first, then to the minimal shape we need,
//    which is safe because we only read ._id from each document.

import { Request, Response, NextFunction } from "express";
import { verifyToken }   from "../lib/jwt";
import User              from "../models/User";
import Program           from "../models/Program";
import mongoose          from "mongoose";
import { logAudit }      from "../lib/auditLogger";

// ── ScopedUser ────────────────────────────────────────────────────────────────
// A lean POJO — never a Mongoose Document. Populated from User.findById().lean().
export interface ScopedUser {
  _id:             mongoose.Types.ObjectId;
  name:            string;
  email:           string;
  role:            "admin" | "coordinator" | "lecturer";
  status:          "active" | "suspended";
  institution:     mongoose.Types.ObjectId;
  schoolCode:      string | null;
  departmentCode:  string | null;
  institutionWide: boolean;
  tokenVersion:    number;
}

// ── AuthenticatedRequest ──────────────────────────────────────────────────────
// Import this in route files — do NOT reference the global Express.Request user
// field after requireAuth because that is still typed as optional UserSafe.
export interface AuthenticatedRequest extends Request {
  user: ScopedUser;
}

// ── requireAuth ───────────────────────────────────────────────────────────────
export async function requireAuth(
  req:  Request,
  res:  Response,
  next: NextFunction,
): Promise<void> {
  const token = (req as Request & { cookies?: Record<string, string> }).cookies?.token;

  if (!token) {
    await logAudit(req, { action: "unauthenticated_access", details: { path: req.originalUrl } });
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

    // .lean() returns a plain object — select only the fields ScopedUser needs
    // so that TypeScript has a concrete shape to work with.
    const userDoc = await User.findById(payload.id)
      .select(
        "name email role status institution tokenVersion schoolCode departmentCode institutionWide",
      )
      .lean<{
        _id:             mongoose.Types.ObjectId;
        name:            string;
        email:           string;
        role:            "admin" | "coordinator" | "lecturer";
        status:          "active" | "suspended";
        institution?:    mongoose.Types.ObjectId;
        schoolCode?:     string | null;
        departmentCode?: string | null;
        institutionWide: boolean;
        tokenVersion?:   number;
      }>();

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

    if (
      typeof payload.version === "number" &&
      payload.version !== (userDoc.tokenVersion ?? 0)
    ) {
      res.clearCookie("token");
      res.status(401).json({ message: "Session expired. Please log in again." });
      return;
    }

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

    // Assemble the ScopedUser — all fields are now correctly typed
    const scopedUser: ScopedUser = {
      _id:             userDoc._id,
      name:            userDoc.name,
      email:           userDoc.email,
      role:            userDoc.role,
      status:          userDoc.status,
      institution:     new mongoose.Types.ObjectId(payload.institution),
      schoolCode:      userDoc.schoolCode ?? null,
      departmentCode:  userDoc.departmentCode ?? null,
      institutionWide: userDoc.institutionWide ?? false,
      tokenVersion:    userDoc.tokenVersion ?? 0,
    };

    (req as AuthenticatedRequest).user = scopedUser;
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

// ── requireRole ───────────────────────────────────────────────────────────────
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

// ── getScopedProgramIds ───────────────────────────────────────────────────────
// Returns the ObjectId strings of programs visible to this user.
//
// FIX: Mongoose .lean() returns FlattenMaps<T> where _id is FlattenMaps<unknown>.
// This is NOT assignable to ObjectId directly. The fix is to provide an explicit
// generic to .lean<MinimalShape>() so TypeScript knows the exact shape we get
// back, instead of letting it infer the problematic FlattenMaps<IProgram> type.
export async function getScopedProgramIds(req: AuthenticatedRequest): Promise<string[]> {
  const filter: Record<string, unknown> = { institution: req.user.institution };

  if (!req.user.institutionWide) {
    if (req.user.departmentCode) filter.departmentCode = req.user.departmentCode;
    if (req.user.schoolCode)     filter.schoolCode     = req.user.schoolCode;
  }

  // Provide the explicit shape to .lean() — avoids FlattenMaps<unknown> on _id
  const programs = await Program.find(filter)
    .select("_id")
    .lean<Array<{ _id: mongoose.Types.ObjectId }>>();

  return programs.map(p => p._id.toString());
}