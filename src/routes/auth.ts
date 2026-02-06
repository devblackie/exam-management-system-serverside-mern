// src/routes/auth.ts
import { Router, Request, Response } from "express";
import bcrypt from "bcryptjs";
import User from "../models/User";
import { setAuthCookie } from "../lib/jwt";
import { requireAuth } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { ApiError } from "../middleware/errorHandler";
import AuditLog from "../models/AuditLog";
import { logAudit } from "../lib/auditLogger";
import { loginRateLimiter, sanitizeInput } from "../middleware/security";
import crypto from "crypto";
import nodemailer from "nodemailer";
import { sendRecoveryEmail } from "../config/passwordResetEmail";

const router = Router();

// 🔑 Login
router.post(
  "/login",
  loginRateLimiter, 
  sanitizeInput, 
  asyncHandler(async (req: Request, res: Response) => {

    const email = String(req.body.email || "")
      .toLowerCase()
      .trim();
    const password = String(req.body.password || "");

    if (!email || !password) {
      throw { statusCode: 400, message: "Missing credentials" } as ApiError;
    }

    const user = await User.findOne({ email }).select(
      "+password +tokenVersion",
    );

    const dummyHash =
      "$2a$12$LRYuW9uB6S1EjSM0rE9Q9u3Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9";
    const userPassword = user?.password || dummyHash;
    const isPasswordValid = await bcrypt.compare(password, userPassword);

    // Now check if user actually exists and password matches
    if (!user || !isPasswordValid) {
      throw { statusCode: 401, message: "Invalid credentials" } as ApiError;
    }

   if (user.status === "suspended") {
     throw {
       statusCode: 403,
       message: "Account suspended. Contact administration.",
     } as ApiError;
   }

    // CRITICAL: Include institution in JWT
    setAuthCookie(
      res,
      user._id.toString(),
      user.role,
      user.institution?.toString(),
      user.tokenVersion || 0,
    );

    // Audit log (non-blocking)
    logAudit(req, {
      action: "login_success",
      actor: user._id,
      details: {
        email: user.email,
        role: user.role,
        institution: user.institution,
        userAgent: req.get("User-Agent"),
        ip: req.ip,
      },
    });

    // Send safe response
    res.json({
      message: "Login successful",
      user: {
        name: user.name,
        email: user.email,
        role: user.role,
        institution: user.institution, // ← Optional: send to frontend
      },
    });
  }),
);

// 👤 Current user
router.get(
  "/me",
  requireAuth,
  asyncHandler(async (req: Request & { user?: any }, res: Response) => {
    const user = req.user;
    if (!user) throw { statusCode: 401, message: "Not authenticated" } as ApiError;

    res.json({ role: user.role, email: user.email, name: user.name });
  })
);

// 🚪 Logout
router.post(
  "/logout",
  requireAuth, // ✅ make sure only logged-in users can logout
  asyncHandler(async (req: Request & { user?: any }, res: Response) => {
    const actorId = req.user?._id; // ✅ properly defined here

    res.clearCookie("token", {
      httpOnly: true,
      sameSite: "strict",
      secure: process.env.NODE_ENV === "production",
    });

    // ✅ Log logout event (non-blocking)
    if (actorId) {
      logAudit(req, {
        action: "logout",
        actor: actorId,
        targetUser: actorId,
        
      });
    }

    res.json({ message: "Logged out" });
  })
);

// 🔒 Forgot Password
router.post("/forgot-password", loginRateLimiter, asyncHandler(async (req: Request, res: Response) => {
  const email = String(req.body.email || "").toLowerCase().trim();
  const user = await User.findOne({ email });

  // Security Tip: Always return a success message even if the email doesn't exist
  // This prevents "Account Enumeration" attacks.
  if (!user) {
    return res.json({ message: "If an account exists, a recovery link has been sent." });
  }

  // 1. Generate Token
  const resetToken = crypto.randomBytes(32).toString("hex");
  const hashedToken = crypto.createHash("sha256").update(resetToken).digest("hex");

  // 2. Save hashed token to DB (valid for 1 hour)
  user.passwordResetToken = hashedToken;
  user.passwordResetExpires = new Date(Date.now() + 3600000); 
  await user.save();

  // 3. Configure Nodemailer
  try {
    await sendRecoveryEmail(user.email, resetToken, user.name);
  } catch (error) {
    console.error("Email Protocol Error:", error);
    // We don't throw an error to the user to avoid leaking info, 
    // but we log it internally.
  }

  res.json({ message: "If an account exists, a recovery link has been sent." });
}));

// 🔒 Reset Password
router.post("/reset-password/:token", asyncHandler(async (req: Request, res: Response) => {
  const { token } = req.params;
  const { password } = req.body;

  if (!password || password.length < 8) {
    throw {
      statusCode: 400,
      message: "Security Key must be at least 8 characters.",
    } as ApiError;
  }
  // 1. Hash the incoming URL token to match the DB version
  const hashedToken = crypto.createHash("sha256").update(token).digest("hex");

  // 2. Find user with valid token and check expiration
  const user = await User.findOne({
    passwordResetToken: hashedToken,
    passwordResetExpires: { $gt: new Date() },
  }).select("+password +tokenVersion");

  if (!user) {
    throw { statusCode: 400, message: "Link expired or invalid protocol." } as ApiError;
  }

  // 3. Update password and clear reset fields
  const salt = await bcrypt.genSalt(12);
  user.password = await bcrypt.hash(String(password), salt);
  user.tokenVersion = (user.tokenVersion || 0) + 1;
  user.passwordResetToken = undefined;
  user.passwordResetExpires = undefined;
  
  // Also unlock the account if it was suspended due to brute force
  user.status = "active"; 

  await user.save();

  res.clearCookie("token", {
    httpOnly: true,
    sameSite: "strict",
    });

  // 4. Log the security event
  logAudit(req, {
    action: "password_reset_success",
    actor: user._id,
    details: {
      email: user.email,
      ip: req.ip,
      message: "All sessions invalidated due to password update",
    },
  });

  res.json({ message: "Security Key Updated Successfully." });
}));

export default router;
