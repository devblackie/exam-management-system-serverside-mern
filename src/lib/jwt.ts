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
  const token = jwt.sign(payload, JWT_SECRET, { expiresIn: "1d" });
  res.cookie("token", token, {
    httpOnly: true,
    sameSite: "strict",
    secure: process.env.NODE_ENV === "production",
  });
};

// Verify JWT
export const verifyToken = (token: string) => {
  return jwt.verify(token, JWT_SECRET);
};
