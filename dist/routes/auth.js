"use strict";
// serverside/src/routes/auth.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = require("express");
const bcryptjs_1 = __importDefault(require("bcryptjs"));
const node_crypto_1 = __importDefault(require("node:crypto"));
const mongoose_1 = __importDefault(require("mongoose"));
const User_1 = __importDefault(require("../models/User"));
const TempOTP_1 = __importDefault(require("../models/TempOTP"));
const jwt_1 = require("../lib/jwt");
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const auditLogger_1 = require("../lib/auditLogger");
const passwordResetEmail_1 = require("../config/passwordResetEmail");
const twoFactorService_1 = require("../services/twoFactorService");
const security_1 = require("../middleware/security");
const validation_1 = require("../middleware/validation");
const router = (0, express_1.Router)();
const isProd = process.env.NODE_ENV === "production";
const COOKIE_OPTS = {
    httpOnly: true,
    sameSite: "lax", // ← was "strict" — fix to match jwt.ts
    secure: isProd,
    path: "/",
};
// Must be a valid bcrypt hash with the same prefix ($2b$12$) as your real
// hashes so bcrypt.compare() doesn't short-circuit on format mismatch.
const DUMMY_HASH = "$2b$12$LRYuW9uB6S1EjSM0rE9Q9uLRYuW9uB6S1EjSM0rE9Q9uLRYuW9uBC";
const setStepCookie = (res, name, userId, fingerprint, maxAgeMs) => {
    const value = Buffer.from(JSON.stringify({ userId, fingerprint })).toString("base64");
    res.cookie(name, value, { ...COOKIE_OPTS, maxAge: maxAgeMs });
};
const readStepCookie = (req, name, fingerprint) => {
    const raw = req.cookies?.[name];
    if (!raw)
        return null;
    try {
        const parsed = JSON.parse(Buffer.from(raw, "base64").toString());
        if (parsed.fingerprint !== fingerprint)
            return null;
        return { userId: parsed.userId };
    }
    catch {
        return null;
    }
};
const clearStepCookies = (res) => {
    res.clearCookie("auth_step1", COOKIE_OPTS);
    res.clearCookie("auth_step2", COOKIE_OPTS);
};
router.post("/check-email", security_1.emailCheckLimiter, security_1.sanitizeInput, security_1.honeypotCheck, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const email = String(req.body.email || "").toLowerCase().trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        throw { statusCode: 400, message: "Invalid email format." };
    }
    const fingerprint = (0, security_1.getRequestFingerprint)(req);
    const user = await User_1.default.findOne({ email }).select("_id name status").lean();
    // Timing equalizer
    await bcryptjs_1.default.hash(email + Date.now(), 4);
    if (!user || user.status === "suspended") {
        const dummyId = new mongoose_1.default.Types.ObjectId().toString();
        setStepCookie(res, "auth_step1", dummyId, fingerprint, 10 * 60 * 1000);
        return res.json({ nextStep: "password" });
    }
    setStepCookie(res, "auth_step1", user._id.toString(), fingerprint, 10 * 60 * 1000);
    (0, auditLogger_1.logAudit)(req, { action: "login_step1_email_checked", details: { email } });
    res.json({ nextStep: "password", maskedName: user.name.split(" ")[0] });
}));
// STEP 2 — POST /auth/verify-password
router.post("/verify-password", security_1.passwordLimiter, security_1.sanitizeInput, security_1.progressiveDelayMiddleware, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const password = String(req.body.password || "");
    const fingerprint = (0, security_1.getRequestFingerprint)(req);
    const ip = req.ip || req.socket.remoteAddress || "unknown";
    const step1 = readStepCookie(req, "auth_step1", fingerprint);
    if (!step1) {
        throw {
            statusCode: 401,
            message: "Session expired or invalid. Please start again.",
        };
    }
    // Query 1 — all non-select:false fields
    const user = await User_1.default.findById(step1.userId)
        .select("email name status institution role tokenVersion")
        .lean();
    // Query 2 — ONLY the password hash, isolated to avoid projection conflict
    const userPw = await User_1.default.findById(step1.userId).select("+password").lean();
    const storedHash = userPw?.password ?? null;
    // Constant-time comparison — always runs bcrypt even when user is missing
    const hashToCompare = storedHash ?? DUMMY_HASH;
    const isValid = await bcryptjs_1.default.compare(password, hashToCompare);
    // Lockout check after bcrypt so timing stays consistent
    if (user) {
        const lockout = (0, security_1.checkAccountLockout)(user.email);
        if (lockout.isLocked) {
            const mins = Math.ceil(lockout.remainingMs / 60000);
            throw {
                statusCode: 423,
                message: `Account temporarily locked. Try again in ${mins} minute${mins !== 1 ? "s" : ""}.`,
            };
        }
    }
    if (!user || !storedHash || !isValid) {
        (0, security_1.recordFailedAttempt)(ip);
        if (user)
            (0, security_1.recordFailedPasswordAttempt)(user.email);
        (0, auditLogger_1.logAudit)(req, { action: "login_step2_password_failed", details: { userId: step1.userId } });
        throw { statusCode: 401, message: "Invalid credentials." };
    }
    if (user.status === "suspended") {
        throw {
            statusCode: 403,
            message: "Account suspended. Contact administration.",
        };
    }
    (0, security_1.clearFailedAttempts)(ip);
    (0, security_1.clearAccountLockout)(user.email);
    const otp = node_crypto_1.default.randomInt(100000, 999999).toString();
    const otpHash = node_crypto_1.default.createHash("sha256").update(otp).digest("hex");
    await TempOTP_1.default.deleteMany({ userId: user._id });
    await TempOTP_1.default.create({
        userId: user._id,
        otpHash,
        attempts: 0,
        expiresAt: new Date(Date.now() + 10 * 60 * 1000),
        fingerprint,
    });
    res.clearCookie("auth_step1", COOKIE_OPTS);
    setStepCookie(res, "auth_step2", user._id.toString(), fingerprint, 12 * 60 * 1000);
    (0, twoFactorService_1.sendOTPEmail)(user.email, user.name, otp, "login").catch((err) => {
        console.error("[Auth] OTP email failed:", err.message);
    });
    (0, auditLogger_1.logAudit)(req, {
        action: "login_step2_password_verified",
        actor: user._id,
        details: { email: user.email },
    });
    const maskedEmail = user.email.replace(/(.{2})[^@]+(@.+)/, "$1***$2");
    res.json({ requiresOTP: true, maskedEmail });
}));
// STEP 3 — POST /auth/verify-otp
router.post("/verify-otp", security_1.otpLimiter, validation_1.otpValidation, validation_1.validateRequest, security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const otp = String(req.body.otp || "").trim().replace(/\s/g, "");
    const fingerprint = (0, security_1.getRequestFingerprint)(req);
    const step2 = readStepCookie(req, "auth_step2", fingerprint);
    if (!step2) {
        throw {
            statusCode: 401,
            message: "Session expired or invalid. Please log in again.",
        };
    }
    if (!/^\d{6}$/.test(otp)) {
        throw { statusCode: 400, message: "Verification code must be 6 digits." };
    }
    const tempRecord = await TempOTP_1.default.findOne({
        userId: step2.userId,
        expiresAt: { $gt: new Date() },
        fingerprint,
    });
    if (!tempRecord) {
        clearStepCookies(res);
        throw {
            statusCode: 401,
            message: "Verification code expired. Please log in again.",
        };
    }
    const MAX_OTP_ATTEMPTS = 5;
    if (tempRecord.attempts >= MAX_OTP_ATTEMPTS) {
        await TempOTP_1.default.deleteOne({ _id: tempRecord._id });
        clearStepCookies(res);
        throw {
            statusCode: 401,
            message: "Too many incorrect attempts. Please log in again.",
        };
    }
    const inputHash = node_crypto_1.default.createHash("sha256").update(otp).digest("hex");
    let isValidOTP = false;
    try {
        isValidOTP = node_crypto_1.default.timingSafeEqual(Buffer.from(inputHash, "hex"), Buffer.from(tempRecord.otpHash, "hex"));
    }
    catch {
        isValidOTP = false;
    }
    if (!isValidOTP) {
        await TempOTP_1.default.updateOne({ _id: tempRecord._id }, { $inc: { attempts: 1 } });
        const remaining = MAX_OTP_ATTEMPTS - tempRecord.attempts - 1;
        throw {
            statusCode: 401,
            message: `Invalid code. ${remaining} attempt${remaining !== 1 ? "s" : ""} remaining.`,
        };
    }
    await TempOTP_1.default.deleteOne({ _id: tempRecord._id });
    clearStepCookies(res);
    // tokenVersion is NOT select:false — plain select is safe here
    const user = await User_1.default.findById(step2.userId)
        .select("name email role institution tokenVersion")
        .lean();
    if (!user) {
        throw { statusCode: 401, message: "User not found." };
    }
    (0, jwt_1.setAuthCookie)(res, user._id.toString(), user.role, user.institution?.toString(), user.tokenVersion ?? 0);
    (0, auditLogger_1.logAudit)(req, {
        action: "login_success",
        actor: user._id,
        details: { email: user.email, role: user.role },
    });
    res.json({
        message: "Login successful",
        user: {
            name: user.name,
            email: user.email,
            role: user.role,
            institution: user.institution,
        },
    });
}));
// GET /auth/me
router.get("/me", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const user = req.user;
    if (!user)
        throw { statusCode: 401, message: "Not authenticated" };
    res.json({ role: user.role, email: user.email, name: user.name,
        schoolCode: user.schoolCode ?? null,
        departmentCode: user.departmentCode ?? null,
        institutionWide: user.institutionWide ?? false,
    });
}));
// POST /auth/logout
router.post("/logout", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const actorId = req.user?._id;
    res.clearCookie("token", COOKIE_OPTS);
    res.clearCookie("auth_step1", COOKIE_OPTS);
    res.clearCookie("auth_step2", COOKIE_OPTS);
    if (actorId) {
        (0, auditLogger_1.logAudit)(req, { action: "logout", actor: actorId, targetUser: actorId });
    }
    res.json({ message: "Logged out" });
}));
// POST /auth/forgot-password
router.post("/forgot-password", security_1.loginRateLimiter, security_1.sanitizeInput, security_1.honeypotCheck, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const email = String(req.body.email || "").toLowerCase().trim();
    const GENERIC = "If an account exists with that email, a recovery link has been sent.";
    if (!email)
        return res.json({ message: GENERIC });
    const user = await User_1.default.findOne({ email }).select("_id email name").lean();
    if (!user) {
        await bcryptjs_1.default.hash(email, 4);
        return res.json({ message: GENERIC });
    }
    const resetToken = node_crypto_1.default.randomBytes(32).toString("hex");
    const hashedToken = node_crypto_1.default.createHash("sha256").update(resetToken).digest("hex");
    await User_1.default.updateOne({ _id: user._id }, {
        passwordResetToken: hashedToken,
        passwordResetExpires: new Date(Date.now() + 3600000),
    });
    (0, passwordResetEmail_1.sendRecoveryEmail)(user.email, resetToken, user.name).catch((err) => {
        console.error("[Auth] Recovery email failed:", err.message);
    });
    (0, auditLogger_1.logAudit)(req, { action: "password_reset_requested", details: { email } });
    res.json({ message: GENERIC });
}));
// POST /auth/reset-password/:token
router.post("/reset-password/:token", security_1.loginRateLimiter, security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { token } = req.params;
    const { password } = req.body;
    if (!password || String(password).length < 8) {
        throw { statusCode: 400, message: "Password must be at least 8 characters." };
    }
    const hashedToken = node_crypto_1.default.createHash("sha256").update(token).digest("hex");
    // passwordResetToken + passwordResetExpires are select:false.
    // Use isolated .select() — same reason as the password fix above.
    const user = await User_1.default.findOne({
        passwordResetToken: hashedToken,
        passwordResetExpires: { $gt: new Date() },
    }).select("+passwordResetToken +passwordResetExpires");
    if (!user) {
        throw { statusCode: 400, message: "Reset link is invalid or has expired." };
    }
    user.password = await bcryptjs_1.default.hash(String(password), 12);
    user.tokenVersion = (user.tokenVersion ?? 0) + 1;
    user.passwordResetToken = undefined;
    user.passwordResetExpires = undefined;
    user.status = "active";
    await user.save();
    (0, security_1.clearAccountLockout)(user.email);
    res.clearCookie("token", COOKIE_OPTS);
    res.clearCookie("auth_step1", COOKIE_OPTS);
    res.clearCookie("auth_step2", COOKIE_OPTS);
    (0, auditLogger_1.logAudit)(req, { action: "password_reset_success", actor: user._id, details: { email: user.email } });
    res.json({ message: "Password updated successfully. Please log in." });
}));
exports.default = router;
