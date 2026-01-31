// serverside/src/lib/jwt.ts
import jwt from "jsonwebtoken";
import { Response } from "express";
import config from "../config/config";

const JWT_SECRET = config.jwtSecret || "supersecret";

// Create JWT and store in HttpOnly cookie
export const setAuthCookie = (res: Response, userId: string, role: string, institution?: string | null) => {

   const payload = {
    id: userId,
    role,
    institution: institution || null, // ← Always include (even if null)
  };
 // 1. JWT expires in 1 day
  const token = jwt.sign(payload, JWT_SECRET, { expiresIn: "1d" });

  // 2. Cookie expires in 1 day (24 hours * 60 mins * 60 secs * 1000 ms)
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  res.cookie("token", token, {
    httpOnly: true,
    sameSite: "strict",
    secure: process.env.NODE_ENV === "production",
    maxAge: ONE_DAY_MS,
  });
};

// Verify JWT
export const verifyToken = (token: string) => {
  return jwt.verify(token, JWT_SECRET);
};
