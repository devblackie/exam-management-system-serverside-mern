// // serverside/src/lib/jwt.ts
// import jwt from "jsonwebtoken";
// import { Response } from "express";
// import config from "../config/config";

// // const JWT_SECRET = config.jwtSecret;
// const JWT_SECRET = config.jwtSecret + "_pending";

// export const signPendingToken = (userId: string): string => {
//   return jwt.sign({ id: userId, type: "pending" }, JWT_SECRET, {
//     expiresIn: "15m",
//   });
// };

// export const verifyPendingToken = (
//   token: string,
// ): { id: string; type: string } => {
//   return jwt.verify(token, JWT_SECRET) as { id: string; type: string };
// };

// // Create JWT and store in HttpOnly cookie
// export const setAuthCookie = (
//   res: Response,
//   userId: string,
//   role: string,
//   institution?: string | null,
//   version: number = 0,
// ) => {
//   const payload = {
//     id: userId,
//     role,
//     institution: institution || null, // ← Always include (even if null)
//  version,
//   };
//   // 1. JWT expires in 1 day
//   const token = jwt.sign(payload, JWT_SECRET, { expiresIn: "1d" });

//   // 2. Cookie expires in 1 day (24 hours * 60 mins * 60 secs * 1000 ms)
//   const ONE_DAY_MS = 24 * 60 * 60 * 1000;
//   res.cookie("token", token, {
//     httpOnly: true,
//     sameSite: "strict",
//     secure: process.env.NODE_ENV === "production",
//     maxAge: ONE_DAY_MS,
//   });
// };

// // Verify JWT
// export const verifyToken = (token: string) => {
//   return jwt.verify(token, JWT_SECRET);
// };

// serverside/src/lib/jwt.ts
import jwt from "jsonwebtoken";
import { Response } from "express";
import config from "../config/config";

// ─── Two separate secrets ──────────────────────────────────────────────────────
// Pending tokens (step cookies) use a different secret than real session tokens.
// This means a pending token can NEVER be accepted as a session token even if
// an attacker tampers with the "type" claim.
const SESSION_SECRET = config.jwtSecret;
const PENDING_SECRET = config.jwtSecret + "_pending";

// ─── Pending token (short-lived, step-binding only) ───────────────────────────
export const signPendingToken = (userId: string): string => {
  return jwt.sign({ id: userId, type: "pending" }, PENDING_SECRET, {
    expiresIn: "15m",
  });
};

export const verifyPendingToken = (
  token: string,
): { id: string; type: string } => {
  return jwt.verify(token, PENDING_SECRET) as { id: string; type: string };
};

// ─── Session cookie (issued only after full 3-step login) ─────────────────────
export const setAuthCookie = (
  res:          Response,
  userId:       string,
  role:         string,
  institution?: string | null,
  version:      number = 0,
): void => {
  const payload = {
    id:          userId,
    role,
    institution: institution || null,
    version,
  };

  // JWT expires in 1 day
  const token = jwt.sign(payload, SESSION_SECRET, { expiresIn: "1d" });

  // Cookie expires in 1 day
  res.cookie("token", token, {
    httpOnly: true,
    sameSite: "strict",
    secure:   process.env.NODE_ENV === "production",
    maxAge:   24 * 60 * 60 * 1000,
  });
};

// ─── Verify session token ─────────────────────────────────────────────────────
// Used by requireAuth middleware. Verifies against SESSION_SECRET only —
// pending tokens signed with PENDING_SECRET will be rejected here.
export const verifyToken = (token: string) => {
  return jwt.verify(token, SESSION_SECRET);
};
