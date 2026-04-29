// serverside/src/lib/jwt.ts

import jwt from "jsonwebtoken";
import { Response } from "express";
import config  from "../config/config";

// Single secret — used for all tokens.
// The Next.js middleware does NOT verify signatures, so this value
// only needs to exist in serverside/.env — not in clientside/.env.local.
const JWT_SECRET = config.jwtSecret;
if (!JWT_SECRET) {
  console.error("[JWT] FATAL: JWT_SECRET is not set. Exiting.");
  process.exit(1);
}

export interface JwtPayload {
  id:          string;
  role:        string;
  institution: string;   // ObjectId string — always required
  version:     number;
}

export const signToken = (
  id:          string,
  role:        string,
  institution: string,
  version:     number,
): string => {
  if (!institution) {
    throw new Error(
      `Cannot issue token for user ${id}: institution is required for all users.`
    );
  }
  return jwt.sign(
    { id, role, institution, version } satisfies JwtPayload,
    JWT_SECRET,
    // { expiresIn: "7d" },
    { expiresIn: "1d" },
  );
};

// export const verifyToken = (token: string): JwtPayload => {
//   return jwt.verify(token, JWT_SECRET) as JwtPayload;
// };

const JWT_SECRET_PREVIOUS = process.env.JWT_SECRET_PREVIOUS;

// REPLACE your existing verifyToken with this — identical interface, adds rotation
export const verifyToken = (token: string): JwtPayload => {
  try {
    return jwt.verify(token, JWT_SECRET) as JwtPayload;
  } catch (primaryErr) {
    // Only attempt previous secret during an active rotation window
    if (JWT_SECRET_PREVIOUS) {
      try {
        return jwt.verify(token, JWT_SECRET_PREVIOUS) as JwtPayload;
      } catch {
        // Previous secret also failed — throw the original error
      }
    }
    throw primaryErr;
  }
};

export const setAuthCookie = (
  res:         Response,
  id:          string,
  role:        string,
  institution: string | undefined | null,
  version:     number,
): void => {
  if (!institution) {
    // This should never happen if admin registration enforces institutionId.
    // If it does, it means a legacy user exists without an institution.
    throw new Error(`Cannot create session for user ${id}: no institution linked.`);
  }

  const token = signToken(id, role, institution, version);

  res.cookie("token", token, {
    httpOnly: true,
    // Lax in dev (allows cross-port navigation), Strict in production
    sameSite: process.env.NODE_ENV === "production" ? "strict" : "lax",
    secure:   process.env.NODE_ENV === "production",
    // maxAge:   7 * 24 * 60 * 60 * 1000,
    maxAge:   24 * 60 * 60 * 1000, // 1 day
    path:     "/",
  });
};

export const signPendingToken = (userId: string): string =>
  jwt.sign({ id: userId, type: "pending" }, JWT_SECRET, { expiresIn: "15m" });

export const verifyPendingToken = (token: string): { id: string; type: string } =>
  jwt.verify(token, JWT_SECRET) as { id: string; type: string };




