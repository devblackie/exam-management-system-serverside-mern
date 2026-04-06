// serverside/src/routes/auth.ts

import { Router, Request, Response } from "express";
import bcrypt from "bcryptjs";
import crypto from "crypto";
import mongoose from "mongoose";
import User from "../models/User";
import TempOTP from "../models/TempOTP";
import { setAuthCookie } from "../lib/jwt";
import { requireAuth } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { ApiError } from "../middleware/errorHandler";
import { logAudit } from "../lib/auditLogger";
import { sendRecoveryEmail } from "../config/passwordResetEmail";
import { sendOTPEmail } from "../services/twoFactorService";
import {
  emailCheckLimiter,
  passwordLimiter,
  otpLimiter,
  sanitizeInput,
  honeypotCheck,
  getRequestFingerprint,
  recordFailedAttempt,
  clearFailedAttempts,
  progressiveDelayMiddleware,
  checkAccountLockout,
  recordFailedPasswordAttempt,
  clearAccountLockout,
  loginRateLimiter,
} from "../middleware/security";

const router = Router();

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────────────────────

const COOKIE_OPTS = {
  httpOnly: true,
  sameSite: "strict" as const,
  secure:   process.env.NODE_ENV === "production",
};

// Must be a valid bcrypt hash with the same prefix ($2b$12$) as your real
// hashes so bcrypt.compare() doesn't short-circuit on format mismatch.
const DUMMY_HASH =
  "$2b$12$LRYuW9uB6S1EjSM0rE9Q9uLRYuW9uB6S1EjSM0rE9Q9uLRYuW9uBC";

// ─────────────────────────────────────────────────────────────────────────────
// STEP COOKIE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const setStepCookie = (
  res: Response, name: string, userId: string, fingerprint: string, maxAgeMs: number): void => {
  const value = Buffer.from(JSON.stringify({ userId, fingerprint })).toString("base64");
  res.cookie(name, value, { ...COOKIE_OPTS, maxAge: maxAgeMs });
};

const readStepCookie = (
  req: Request, name: string, fingerprint: string): { userId: string } | null => {
  const raw = req.cookies?.[name];
  if (!raw) return null;
  try {
    const parsed = JSON.parse(Buffer.from(raw, "base64").toString());
    if (parsed.fingerprint !== fingerprint) return null;
    return { userId: parsed.userId };
  } catch {
    return null;
  }
};

const clearStepCookies = (res: Response): void => {
  res.clearCookie("auth_step1", COOKIE_OPTS);
  res.clearCookie("auth_step2", COOKIE_OPTS);
};


router.post(
  "/check-email",
  emailCheckLimiter,
  sanitizeInput,
  honeypotCheck,
  asyncHandler(async (req: Request, res: Response) => {
    const email = String(req.body.email || "").toLowerCase().trim();

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      throw { statusCode: 400, message: "Invalid email format." } as ApiError;
    }

    const fingerprint = getRequestFingerprint(req);
    const user        = await User.findOne({ email })
      .select("_id name status")
      .lean();

    // Timing equalizer
    await bcrypt.hash(email + Date.now(), 4);

    if (!user || user.status === "suspended") {
      const dummyId = new mongoose.Types.ObjectId().toString();
      setStepCookie(res, "auth_step1", dummyId, fingerprint, 10 * 60 * 1000);
      return res.json({ nextStep: "password" });
    }

    setStepCookie(res, "auth_step1", user._id.toString(), fingerprint, 10 * 60 * 1000);

    logAudit(req, {
      action:  "login_step1_email_checked",
      details: { email },
    });

    res.json({
      nextStep:   "password",
      maskedName: user.name.split(" ")[0],
    });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// STEP 2 — POST /auth/verify-password
//
// THE BUG (now fixed):
// Mongoose silently drops `+password` when it appears alongside plain field
// names in a single .select() string on this version of Mongoose/driver.
// Confirmed by diagnostics:
//   .select("+password")                          → password returned ✅
//   .select("email name status ... +password")    → password undefined ❌
//
// FIX: Two separate .findById() calls on the same _id (both hit the PK index,
// negligible overhead). Query 1 fetches all normal fields. Query 2 fetches
// ONLY the password hash using an isolated .select("+password").
// ─────────────────────────────────────────────────────────────────────────────

router.post(
  "/verify-password",
  passwordLimiter,
  sanitizeInput,
  progressiveDelayMiddleware,
  asyncHandler(async (req: Request, res: Response) => {
    const password    = String(req.body.password || "");
    const fingerprint = getRequestFingerprint(req);
    const ip          = req.ip || req.socket.remoteAddress || "unknown";

    const step1 = readStepCookie(req, "auth_step1", fingerprint);
    if (!step1) {
      throw {
        statusCode: 401,
        message:    "Session expired or invalid. Please start again.",
      } as ApiError;
    }

    // Query 1 — all non-select:false fields
    const user = await User.findById(step1.userId)
      .select("email name status institution role tokenVersion")
      .lean();

    // Query 2 — ONLY the password hash, isolated to avoid projection conflict
    const userPw = await User.findById(step1.userId)
      .select("+password")
      .lean();

    const storedHash = userPw?.password ?? null;

    // Constant-time comparison — always runs bcrypt even when user is missing
    const hashToCompare = storedHash ?? DUMMY_HASH;
    const isValid       = await bcrypt.compare(password, hashToCompare);

    // Lockout check after bcrypt so timing stays consistent
    if (user) {
      const lockout = checkAccountLockout(user.email);
      if (lockout.isLocked) {
        const mins = Math.ceil(lockout.remainingMs / 60_000);
        throw {
          statusCode: 423,
          message:    `Account temporarily locked. Try again in ${mins} minute${mins !== 1 ? "s" : ""}.`,
        } as ApiError;
      }
    }

    if (!user || !storedHash || !isValid) {
      recordFailedAttempt(ip);
      if (user) recordFailedPasswordAttempt(user.email);
      logAudit(req, {
        action:  "login_step2_password_failed",
        details: { userId: step1.userId },
      });
      throw { statusCode: 401, message: "Invalid credentials." } as ApiError;
    }

    if (user.status === "suspended") {
      throw {
        statusCode: 403,
        message:    "Account suspended. Contact administration.",
      } as ApiError;
    }

    clearFailedAttempts(ip);
    clearAccountLockout(user.email);

    const otp     = crypto.randomInt(100000, 999999).toString();
    const otpHash = crypto.createHash("sha256").update(otp).digest("hex");

    await TempOTP.deleteMany({ userId: user._id });
    await TempOTP.create({
      userId:      user._id,
      otpHash,
      attempts:    0,
      expiresAt:   new Date(Date.now() + 10 * 60 * 1000),
      fingerprint,
    });

    res.clearCookie("auth_step1", COOKIE_OPTS);
    setStepCookie(res, "auth_step2", user._id.toString(), fingerprint, 12 * 60 * 1000);

    sendOTPEmail(user.email, user.name, otp, "login").catch((err: Error) => {
      console.error("[Auth] OTP email failed:", err.message);
    });

    logAudit(req, {
      action:  "login_step2_password_verified",
      actor:   user._id,
      details: { email: user.email },
    });

    const maskedEmail = user.email.replace(/(.{2})[^@]+(@.+)/, "$1***$2");
    res.json({ requiresOTP: true, maskedEmail });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// STEP 3 — POST /auth/verify-otp
// ─────────────────────────────────────────────────────────────────────────────

router.post(
  "/verify-otp",
  otpLimiter,
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const otp         = String(req.body.otp || "").trim().replace(/\s/g, "");
    const fingerprint = getRequestFingerprint(req);

    const step2 = readStepCookie(req, "auth_step2", fingerprint);
    if (!step2) {
      throw {
        statusCode: 401,
        message:    "Session expired or invalid. Please log in again.",
      } as ApiError;
    }

    if (!/^\d{6}$/.test(otp)) {
      throw { statusCode: 400, message: "Verification code must be 6 digits." } as ApiError;
    }

    const tempRecord = await TempOTP.findOne({
      userId:    step2.userId,
      expiresAt: { $gt: new Date() },
      fingerprint,
    });

    if (!tempRecord) {
      clearStepCookies(res);
      throw {
        statusCode: 401,
        message:    "Verification code expired. Please log in again.",
      } as ApiError;
    }

    const MAX_OTP_ATTEMPTS = 5;
    if (tempRecord.attempts >= MAX_OTP_ATTEMPTS) {
      await TempOTP.deleteOne({ _id: tempRecord._id });
      clearStepCookies(res);
      throw {
        statusCode: 401,
        message:    "Too many incorrect attempts. Please log in again.",
      } as ApiError;
    }

    const inputHash = crypto.createHash("sha256").update(otp).digest("hex");
    let isValidOTP  = false;
    try {
      isValidOTP = crypto.timingSafeEqual(
        Buffer.from(inputHash,          "hex"),
        Buffer.from(tempRecord.otpHash, "hex"),
      );
    } catch {
      isValidOTP = false;
    }

    if (!isValidOTP) {
      await TempOTP.updateOne({ _id: tempRecord._id }, { $inc: { attempts: 1 } });
      const remaining = MAX_OTP_ATTEMPTS - tempRecord.attempts - 1;
      throw {
        statusCode: 401,
        message:    `Invalid code. ${remaining} attempt${remaining !== 1 ? "s" : ""} remaining.`,
      } as ApiError;
    }

    await TempOTP.deleteOne({ _id: tempRecord._id });
    clearStepCookies(res);

    // tokenVersion is NOT select:false — plain select is safe here
    const user = await User.findById(step2.userId)
      .select("name email role institution tokenVersion")
      .lean();

    if (!user) {
      throw { statusCode: 401, message: "User not found." } as ApiError;
    }

    setAuthCookie(
      res,
      user._id.toString(),
      user.role,
      user.institution?.toString(),
      user.tokenVersion ?? 0,
    );

    logAudit(req, {
      action:  "login_success",
      actor:   user._id,
      details: { email: user.email, role: user.role },
    });

    res.json({
      message: "Login successful",
      user: {
        name:        user.name,
        email:       user.email,
        role:        user.role,
        institution: user.institution,
      },
    });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// GET /auth/me
// ─────────────────────────────────────────────────────────────────────────────

router.get(
  "/me",
  requireAuth,
  asyncHandler(async (req: Request & { user?: any }, res: Response) => {
    const user = req.user;
    if (!user) throw { statusCode: 401, message: "Not authenticated" } as ApiError;
    res.json({ role: user.role, email: user.email, name: user.name });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// POST /auth/logout
// ─────────────────────────────────────────────────────────────────────────────

router.post(
  "/logout",
  requireAuth,
  asyncHandler(async (req: Request & { user?: any }, res: Response) => {
    const actorId = req.user?._id;
    res.clearCookie("token",      COOKIE_OPTS);
    res.clearCookie("auth_step1", COOKIE_OPTS);
    res.clearCookie("auth_step2", COOKIE_OPTS);
    if (actorId) {
      logAudit(req, { action: "logout", actor: actorId, targetUser: actorId });
    }
    res.json({ message: "Logged out" });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// POST /auth/forgot-password
// ─────────────────────────────────────────────────────────────────────────────

router.post(
  "/forgot-password",
  loginRateLimiter,
  sanitizeInput,
  honeypotCheck,
  asyncHandler(async (req: Request, res: Response) => {
    const email   = String(req.body.email || "").toLowerCase().trim();
    const GENERIC = "If an account exists with that email, a recovery link has been sent.";

    if (!email) return res.json({ message: GENERIC });

    const user = await User.findOne({ email }).select("_id email name").lean();

    if (!user) {
      await bcrypt.hash(email, 4);
      return res.json({ message: GENERIC });
    }

    const resetToken  = crypto.randomBytes(32).toString("hex");
    const hashedToken = crypto.createHash("sha256").update(resetToken).digest("hex");

    await User.updateOne(
      { _id: user._id },
      {
        passwordResetToken:   hashedToken,
        passwordResetExpires: new Date(Date.now() + 3_600_000),
      }
    );

    sendRecoveryEmail(user.email, resetToken, user.name).catch((err: Error) => {
      console.error("[Auth] Recovery email failed:", err.message);
    });

    logAudit(req, { action: "password_reset_requested", details: { email } });
    res.json({ message: GENERIC });
  })
);

// ─────────────────────────────────────────────────────────────────────────────
// POST /auth/reset-password/:token
// ─────────────────────────────────────────────────────────────────────────────

router.post(
  "/reset-password/:token",
  loginRateLimiter,
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const { token }    = req.params;
    const { password } = req.body;

    if (!password || String(password).length < 8) {
      throw { statusCode: 400, message: "Password must be at least 8 characters." } as ApiError;
    }

    const hashedToken = crypto.createHash("sha256").update(token).digest("hex");

    // passwordResetToken + passwordResetExpires are select:false.
    // Use isolated .select() — same reason as the password fix above.
    const user = await User.findOne({
      passwordResetToken:   hashedToken,
      passwordResetExpires: { $gt: new Date() },
    }).select("+passwordResetToken +passwordResetExpires");

    if (!user) {
      throw { statusCode: 400, message: "Reset link is invalid or has expired." } as ApiError;
    }

    user.password             = await bcrypt.hash(String(password), 12);
    user.tokenVersion         = (user.tokenVersion ?? 0) + 1;
    user.passwordResetToken   = undefined;
    user.passwordResetExpires = undefined;
    user.status               = "active";
    await user.save();

    clearAccountLockout(user.email);

    res.clearCookie("token",      COOKIE_OPTS);
    res.clearCookie("auth_step1", COOKIE_OPTS);
    res.clearCookie("auth_step2", COOKIE_OPTS);

    logAudit(req, {
      action:  "password_reset_success",
      actor:   user._id,
      details: { email: user.email },
    });

    res.json({ message: "Password updated successfully. Please log in." });
  })
);

export default router;
