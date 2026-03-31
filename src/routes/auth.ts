// // src/routes/auth.ts
// import { Router, Request, Response } from "express";
// import bcrypt from "bcryptjs";
// import User from "../models/User";
// import { setAuthCookie } from "../lib/jwt";
// import { requireAuth } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { ApiError } from "../middleware/errorHandler";
// import AuditLog from "../models/AuditLog";
// import { logAudit } from "../lib/auditLogger";
// import { loginRateLimiter, sanitizeInput } from "../middleware/security";
// import crypto from "crypto";
// import nodemailer from "nodemailer";
// import { sendRecoveryEmail } from "../config/passwordResetEmail";
// import { sendOTPEmail } from "../services/twoFactorService";
// import TempOTP from "../models/TempOTP";

// const router = Router();

// // // 🔑 Login
// // router.post(
// //   "/login",
// //   loginRateLimiter, 
// //   sanitizeInput, 
// //   asyncHandler(async (req: Request, res: Response) => {
// //     const email = String(req.body.email || "")
// //       .toLowerCase()
// //       .trim();
// //     const password = String(req.body.password || "");

// //     if (!email || !password) {
// //       throw { statusCode: 400, message: "Missing credentials" } as ApiError;
// //     }

// //     // const user = await User.findOne({ email }).select(
// //     //   "+password +tokenVersion",
// //     // );
// //     const user = await User.findOne({ email }).select(
// //       "+password +tokenVersion +twoFactorTempToken +twoFactorTempExpires",
// //     );

// //     const dummyHash =
// //       "$2a$12$LRYuW9uB6S1EjSM0rE9Q9u3Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9";
// //     const userPassword = user?.password || dummyHash;
// //     const isPasswordValid = await bcrypt.compare(password, userPassword);

// //     // Now check if user actually exists and password matches
// //     if (!user || !isPasswordValid) {
// //       throw { statusCode: 401, message: "Invalid credentials" } as ApiError;
// //     }

// //     if (user.status === "suspended") {
// //       throw {
// //         statusCode: 403,
// //         message: "Account suspended. Contact administration.",
// //       } as ApiError;
// //     }

// //     // Generate OTP
// //     const otp = generateOTP();
// //     const hashedOTP = hashOTP(otp);
// //     const expires = new Date(Date.now() + 10 * 60 * 1000); // 10 minutes

// //     await User.findByIdAndUpdate(user._id, {
// //       twoFactorTempToken: hashedOTP,
// //       twoFactorTempExpires: expires,
// //     });

// //     // Send OTP email
// //     await sendOTPEmail(user.email, user.name, otp, "login");

// //     logAudit(req, {
// //       action: "login_otp_sent",
// //       actor: user._id,
// //       details: { email: user.email },
// //     });

// //     // Return a partial token so the frontend knows which user is pending
// //     // We sign a short-lived token that only unlocks the /verify-otp endpoint
// //     const pendingToken = require("../lib/jwt").signPendingToken(
// //       user._id.toString(),
// //     );

// //     res.json({
// //       requiresOTP: true,
// //       pendingToken,
// //       message: `Verification code sent to ${email.replace(/(.{2}).+(@.+)/, "$1***$2")}`,
// //     });

// //     // // CRITICAL: Include institution in JWT
// //     // setAuthCookie( res, user._id.toString(), user.role, user.institution?.toString(), user.tokenVersion || 0 );

// //     // // Audit log (non-blocking)
// //     // logAudit(req, {
// //     //   action: "login_success",
// //     //   actor: user._id,
// //     //   details: { email: user.email, role: user.role, institution: user.institution, userAgent: req.get("User-Agent"), ip: req.ip },
// //     // });

// //     // // Send safe response
// //     // res.json({
// //     //   message: "Login successful",
// //     //   user: { name: user.name, email: user.email, role: user.role, institution: user.institution },
// //     // });
// //   }),
// // );

// // // ── Step 2: Verify OTP, issue session cookie ──────────────────────────────────
// // router.post(
// //   "/verify-otp",
// //   loginRateLimiter,
// //   asyncHandler(async (req: Request, res: Response) => {
// //     const { pendingToken, otp } = req.body;

// //     if (!pendingToken || !otp) {
// //       throw { statusCode: 400, message: "Missing verification data" } as ApiError;
// //     }

// //     // Verify the pending token
// //     let payload: any;
// //     try {
// //       payload = require("../lib/jwt").verifyPendingToken(pendingToken);
// //     } catch {
// //       throw { statusCode: 401, message: "Verification session expired. Please log in again." } as ApiError;
// //     }

// //     const user = await User.findById(payload.id)
// //       .select("+twoFactorTempToken +twoFactorTempExpires +tokenVersion");

// //     if (!user) {
// //       throw { statusCode: 401, message: "User not found" } as ApiError;
// //     }

// //     // Check OTP expiry
// //     if (!user.twoFactorTempExpires || user.twoFactorTempExpires < new Date()) {
// //       throw { statusCode: 401, message: "Verification code expired. Please log in again." } as ApiError;
// //     }

// //     // Verify OTP hash
// //     const hashedInput = hashOTP(String(otp).trim());
// //     if (hashedInput !== user.twoFactorTempToken) {
// //       throw { statusCode: 401, message: "Invalid verification code." } as ApiError;
// //     }

// //     // Clear the temp OTP fields
// //     await User.findByIdAndUpdate(user._id, {
// //       twoFactorTempToken:   undefined,
// //       twoFactorTempExpires: undefined,
// //     });

// //     // Issue the real session cookie
// //     setAuthCookie(
// //       res,
// //       user._id.toString(),
// //       user.role,
// //       user.institution?.toString(),
// //       user.tokenVersion || 0,
// //     );

// //     logAudit(req, {
// //       action:  "login_success",
// //       actor:   user._id,
// //       details: { email: user.email, role: user.role },
// //     });

// //     res.json({
// //       message: "Login successful",
// //       user: {
// //         name:        user.name,
// //         email:       user.email,
// //         role:        user.role,
// //         institution: user.institution,
// //       },
// //     });
// //   }),
// // );

// router.post(
//   "/login",
//   loginRateLimiter,
//   sanitizeInput,
//   asyncHandler(async (req: Request, res: Response) => {
//     const email    = String(req.body.email    || "").toLowerCase().trim();
//     const password = String(req.body.password || "");
 
//     if (!email || !password) {
//       throw { statusCode: 400, message: "Missing credentials" } as ApiError;
//     }
 
//     const user = await User.findOne({ email }).select("+password +tokenVersion");
 
//     // Constant-time comparison even for missing user (prevents timing attacks)
//     const dummyHash =
//       "$2a$12$LRYuW9uB6S1EjSM0rE9Q9u3Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9Q9Z9";
//     const isValid = await bcrypt.compare(password, user?.password || dummyHash);
 
//     if (!user || !isValid) {
//       throw { statusCode: 401, message: "Invalid credentials" } as ApiError;
//     }
 
//     if (user.status === "suspended") {
//       throw {
//         statusCode: 403,
//         message:    "Account suspended. Contact administration.",
//       } as ApiError;
//     }
 
//     // Generate OTP
//     const otp     = crypto.randomInt(100000, 999999).toString();
//     const otpHash = crypto.createHash("sha256").update(otp).digest("hex");
 
//     // Delete any existing pending OTP for this user (prevents stale codes)
//     await TempOTP.deleteMany({ userId: user._id });
 
//     // Store hashed OTP (TTL handled by MongoDB index on expiresAt)
//     await TempOTP.create({
//       userId:    user._id,
//       otpHash,
//       expiresAt: new Date(Date.now() + 10 * 60 * 1000), // 10 minutes
//     });
 
//     // Set a short-lived httpOnly pending cookie
//     // This cookie is ONLY used to identify the user during /verify-otp
//     // It is NOT a session cookie — it has no auth privileges
//     res.cookie("pending_auth", user._id.toString(), {
//       httpOnly: true,
//       sameSite: "strict",
//       secure:   process.env.NODE_ENV === "production",
//       maxAge:   15 * 60 * 1000, // 15 minutes
//     });
 
//     // Send OTP email (non-blocking — don't await, log error if it fails)
//     sendOTPEmail(user.email, user.name, otp, "login").catch((err) => {
//       console.error("[Auth] OTP email failed:", err.message);
//     });
 
//     // Audit (non-blocking)
//     logAudit(req, {
//       action:  "login_otp_sent",
//       actor:   user._id,
//       details: { email: user.email },
//     });
 
//     // Return ONLY what the UI needs — no token, no userId
//     const maskedEmail = email.replace(/(.{2})[^@]+(@.+)/, "$1***$2");
//     res.json({
//       requiresOTP: true,
//       maskedEmail,
//     });
//   })
// );
 
// // ─────────────────────────────────────────────────────────────────────────────
// // STEP 2 — POST /auth/verify-otp
// //
// // Reads the pending_auth cookie (httpOnly, set in step 1) to get the userId.
// // Validates the OTP the user typed.
// // On success:
// //   1. Clears the pending_auth cookie
// //   2. Deletes the TempOTP record
// //   3. Issues the real session cookie (same as before)
// //   4. Returns the user object
// //
// // The frontend never sees or handles any token — it just POSTs the 6-digit code.
// // ─────────────────────────────────────────────────────────────────────────────
// router.post(
//   "/verify-otp",
//   loginRateLimiter,
//   asyncHandler(async (req: Request, res: Response) => {
//     const userId = req.cookies?.pending_auth;
//     const otp    = String(req.body.otp || "").trim();
 
//     if (!userId || !otp) {
//       throw {
//         statusCode: 400,
//         message:    "Session expired or missing verification code. Please log in again.",
//       } as ApiError;
//     }
 
//     // Find the pending OTP record
//     const tempRecord = await TempOTP.findOne({
//       userId,
//       expiresAt: { $gt: new Date() },
//     });
 
//     if (!tempRecord) {
//       // Clear the stale cookie
//       res.clearCookie("pending_auth");
//       throw {
//         statusCode: 401,
//         message:    "Verification code expired. Please log in again.",
//       } as ApiError;
//     }
 
//     // Verify the OTP
//     const inputHash = crypto.createHash("sha256").update(otp).digest("hex");
//     if (inputHash !== tempRecord.otpHash) {
//       throw { statusCode: 401, message: "Invalid verification code." } as ApiError;
//     }
 
//     // Clean up — OTP used, delete it
//     await TempOTP.deleteOne({ _id: tempRecord._id });
 
//     // Clear the pending cookie
//     res.clearCookie("pending_auth", {
//       httpOnly: true,
//       sameSite: "strict",
//       secure:   process.env.NODE_ENV === "production",
//     });
 
//     // Fetch full user
//     const user = await User.findById(userId).select("+tokenVersion").lean();
//     if (!user) {
//       throw { statusCode: 401, message: "User not found" } as ApiError;
//     }
 
//     // Issue the real session cookie (same logic as before)
//     setAuthCookie( res, user._id.toString(), user.role, user.institution?.toString(), user.tokenVersion || 0 );
 
//     // Audit
//     logAudit(req, { action: "login_success", actor: user._id, details: { email: user.email, role: user.role }});
 
//     res.json({
//       message: "Login successful",
//       user: { name: user.name, email: user.email, role: user.role, institution: user.institution },
//     });
//   })
// );

// // Current user
// router.get(
//   "/me",
//   requireAuth,
//   asyncHandler(async (req: Request & { user?: any }, res: Response) => {
//     const user = req.user;
//     if (!user) throw { statusCode: 401, message: "Not authenticated" } as ApiError;

//     res.json({ role: user.role, email: user.email, name: user.name });
//   })
// );

// // 🚪 Logout
// router.post(
//   "/logout",
//   requireAuth, // ✅ make sure only logged-in users can logout
//   asyncHandler(async (req: Request & { user?: any }, res: Response) => {
//     const actorId = req.user?._id; // ✅ properly defined here

//     res.clearCookie("token", {
//       httpOnly: true,
//       sameSite: "strict",
//       secure: process.env.NODE_ENV === "production",
//     });

//     // ✅ Log logout event (non-blocking)
//     if (actorId) {
//       logAudit(req, {
//         action: "logout",
//         actor: actorId,
//         targetUser: actorId,
        
//       });
//     }

//     res.json({ message: "Logged out" });
//   })
// );

// // 🔒 Forgot Password
// router.post("/forgot-password", loginRateLimiter, asyncHandler(async (req: Request, res: Response) => {
//   const email = String(req.body.email || "").toLowerCase().trim();
//   const user = await User.findOne({ email });

//   // Security Tip: Always return a success message even if the email doesn't exist
//   // This prevents "Account Enumeration" attacks.
//   if (!user) {
//     return res.json({ message: "If an account exists, a recovery link has been sent." });
//   }

//   // 1. Generate Token
//   const resetToken = crypto.randomBytes(32).toString("hex");
//   const hashedToken = crypto.createHash("sha256").update(resetToken).digest("hex");

//   // 2. Save hashed token to DB (valid for 1 hour)
//   user.passwordResetToken = hashedToken;
//   user.passwordResetExpires = new Date(Date.now() + 3600000); 
//   await user.save();

//   // 3. Configure Nodemailer
//   try {
//     await sendRecoveryEmail(user.email, resetToken, user.name);
//   } catch (error) {
//     console.error("Email Protocol Error:", error);
//     // We don't throw an error to the user to avoid leaking info, 
//     // but we log it internally.
//   }

//   res.json({ message: "If an account exists, a recovery link has been sent." });
// }));

// // 🔒 Reset Password
// router.post("/reset-password/:token", asyncHandler(async (req: Request, res: Response) => {
//   const { token } = req.params;
//   const { password } = req.body;

//   if (!password || password.length < 8) {
//     throw {
//       statusCode: 400,
//       message: "Security Key must be at least 8 characters.",
//     } as ApiError;
//   }
//   // 1. Hash the incoming URL token to match the DB version
//   const hashedToken = crypto.createHash("sha256").update(token).digest("hex");

//   // 2. Find user with valid token and check expiration
//   const user = await User.findOne({
//     passwordResetToken: hashedToken,
//     passwordResetExpires: { $gt: new Date() },
//   }).select("+password +tokenVersion");

//   if (!user) {
//     throw { statusCode: 400, message: "Link expired or invalid protocol." } as ApiError;
//   }

//   // 3. Update password and clear reset fields
//   const salt = await bcrypt.genSalt(12);
//   user.password = await bcrypt.hash(String(password), salt);
//   user.tokenVersion = (user.tokenVersion || 0) + 1;
//   user.passwordResetToken = undefined;
//   user.passwordResetExpires = undefined;
  
//   // Also unlock the account if it was suspended due to brute force
//   user.status = "active"; 

//   await user.save();

//   res.clearCookie("token", {
//     httpOnly: true,
//     sameSite: "strict",
//     });

//   // 4. Log the security event
//   logAudit(req, {
//     action: "password_reset_success",
//     actor: user._id,
//     details: {
//       email: user.email,
//       ip: req.ip,
//       message: "All sessions invalidated due to password update",
//     },
//   });

//   res.json({ message: "Security Key Updated Successfully." });
// }));

// export default router;
































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
  res:         Response,
  name:        string,
  userId:      string,
  fingerprint: string,
  maxAgeMs:    number,
): void => {
  const value = Buffer.from(
    JSON.stringify({ userId, fingerprint })
  ).toString("base64");
  res.cookie(name, value, { ...COOKIE_OPTS, maxAge: maxAgeMs });
};

const readStepCookie = (
  req:         Request,
  name:        string,
  fingerprint: string,
): { userId: string } | null => {
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

// ─────────────────────────────────────────────────────────────────────────────
// STEP 1 — POST /auth/check-email
// ─────────────────────────────────────────────────────────────────────────────

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
