"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.clearAccountLockout = exports.recordFailedPasswordAttempt = exports.checkAccountLockout = exports.LOCKOUT_DURATION = exports.LOCKOUT_THRESHOLD = exports.progressiveDelayMiddleware = exports.clearFailedAttempts = exports.recordFailedAttempt = exports.getProgressiveDelay = exports.getRequestFingerprint = exports.blockSuspiciousRequests = exports.detectSuspiciousRequest = exports.additionalSecurityHeaders = exports.securityHeaders = exports.honeypotCheck = exports.sanitizeInput = exports.loginRateLimiter = exports.registrationLimiter = exports.apiLimiter = exports.otpLimiter = exports.passwordLimiter = exports.emailCheckLimiter = void 0;
// serverside/src/middleware/security.ts
const express_rate_limit_1 = __importDefault(require("express-rate-limit"));
const mongo_sanitize_1 = __importDefault(require("mongo-sanitize"));
const node_crypto_1 = __importDefault(require("node:crypto"));
const helmet_1 = __importDefault(require("helmet"));
// ─────────────────────────────────────────────────────────────────────────────
// RATE LIMITERS
// One limiter per endpoint — each tuned to the threat model of that endpoint.
// ─────────────────────────────────────────────────────────────────────────────
// Step 1: Email lookup — generous (no secret submitted yet)
exports.emailCheckLimiter = (0, express_rate_limit_1.default)({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 20, // 20 email checks per IP
    message: { message: "Too many requests. Please wait before trying again." },
    standardHeaders: true,
    legacyHeaders: false,
    skipSuccessfulRequests: false,
});
// Step 2: Password verify — tight (a secret is being submitted)
exports.passwordLimiter = (0, express_rate_limit_1.default)({
    windowMs: 15 * 60 * 1000,
    max: 5, // 5 attempts per IP per 15 min
    message: { message: "Too many login attempts. Access locked for 15 minutes." },
    standardHeaders: true,
    legacyHeaders: false,
});
// Step 3: OTP verify — very tight (brute-force window is 10 min)
exports.otpLimiter = (0, express_rate_limit_1.default)({
    windowMs: 10 * 60 * 1000, // matches OTP lifetime
    max: 5,
    message: { message: "Too many verification attempts. Please log in again." },
    standardHeaders: true,
    legacyHeaders: false,
});
// General API limiter (apply globally in app.ts if desired)
exports.apiLimiter = (0, express_rate_limit_1.default)({
    windowMs: 60 * 1000,
    max: 120,
    message: { message: "Too many requests." },
    standardHeaders: true,
    legacyHeaders: false,
});
// Registration — very slow to prevent mass account creation
exports.registrationLimiter = (0, express_rate_limit_1.default)({
    windowMs: 60 * 60 * 1000, // 1 hour
    max: 3,
    message: { message: "Too many registration attempts. Please try again later." },
    standardHeaders: true,
    legacyHeaders: false,
});
// ── Backward-compat alias ─────────────────────────────────────────────────────
// Your existing code imports loginRateLimiter from this file.
// Keep it pointing at the password limiter so nothing breaks.
exports.loginRateLimiter = exports.passwordLimiter;
// ─────────────────────────────────────────────────────────────────────────────
// INPUT SANITIZATION
// Strips MongoDB operators ($gt, $where, etc.) from every incoming request.
// Prevents NoSQL injection attacks.
// Your existing sanitizeInput is kept exactly as-is — just re-exported.
// ─────────────────────────────────────────────────────────────────────────────
const sanitizeInput = (req, res, next) => {
    if (req.body) {
        req.body = (0, mongo_sanitize_1.default)(req.body);
    }
    // Query and Params need key-by-key sanitization to avoid "only a getter" TypeError
    if (req.query) {
        Object.keys(req.query).forEach(key => {
            req.query[key] = (0, mongo_sanitize_1.default)(req.query[key]);
        });
    }
    if (req.params) {
        Object.keys(req.params).forEach(key => {
            req.params[key] = (0, mongo_sanitize_1.default)(req.params[key]);
        });
    }
    next();
};
exports.sanitizeInput = sanitizeInput;
// ─────────────────────────────────────────────────────────────────────────────
// HONEYPOT FIELD CHECK
// Every auth form includes a hidden <input name="website"> that real users
// never see or fill. Bots that auto-fill all fields are silently dropped.
// We return a success-looking response so bots don't know they were caught.
// ─────────────────────────────────────────────────────────────────────────────
const honeypotCheck = (req, res, next) => {
    if (req.body?.website || req.body?._gotcha) {
        // Silent drop — return a plausible response, log nothing useful to the bot
        res.json({ nextStep: "password" });
        return;
    }
    next();
};
exports.honeypotCheck = honeypotCheck;
exports.securityHeaders = (0, helmet_1.default)({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            scriptSrc: ["'self'"],
            styleSrc: ["'self'", "'unsafe-inline'"],
            imgSrc: ["'self'", "data:", "blob:"],
            connectSrc: ["'self'"],
            fontSrc: ["'self'"],
            objectSrc: ["'none'"],
            frameSrc: ["'none'"],
            upgradeInsecureRequests: [],
        },
    },
    hsts: {
        maxAge: 31536000, // 1 year
        includeSubDomains: true,
        preload: true,
    },
    noSniff: true,
    xssFilter: true,
    referrerPolicy: { policy: "strict-origin-when-cross-origin" },
    frameguard: { action: "deny" },
    permittedCrossDomainPolicies: false,
});
// Additional custom headers
const additionalSecurityHeaders = (_req, res, next) => {
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader("X-Frame-Options", "DENY");
    res.setHeader("X-XSS-Protection", "1; mode=block");
    res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    next();
};
exports.additionalSecurityHeaders = additionalSecurityHeaders;
// ─────────────────────────────────────────────────────────────────────────────
// SUSPICIOUS ACTIVITY DETECTOR
// Flags requests that look automated or suspicious.
// ─────────────────────────────────────────────────────────────────────────────
const detectSuspiciousRequest = (req) => {
    const ua = (req.headers["user-agent"] || "").toLowerCase();
    const suspiciousPatterns = [
        "curl", "wget", "python-requests", "go-http-client",
        "java/", "libwww-perl", "scrapy", "httpie",
        "axios/0", // raw axios (not a browser)
        "postman", // allow in dev, block in prod
    ];
    if (process.env.NODE_ENV === "production") {
        return suspiciousPatterns.some(p => ua.includes(p));
    }
    // In dev, only flag clearly automated tools
    return ["scrapy", "sqlmap", "nikto", "masscan"].some(p => ua.includes(p));
};
exports.detectSuspiciousRequest = detectSuspiciousRequest;
const blockSuspiciousRequests = (req, res, next) => {
    if ((0, exports.detectSuspiciousRequest)(req)) {
        res.status(403).json({ message: "Forbidden" });
        return;
    }
    next();
};
exports.blockSuspiciousRequests = blockSuspiciousRequests;
// ─────────────────────────────────────────────────────────────────────────────
// REQUEST FINGERPRINT
// Hashes stable browser/network characteristics into a 32-char hex string.
// Used to bind step cookies to the device that started the login flow.
// If someone steals the pending cookie and tries it from a different machine,
// the fingerprint won't match and the step is rejected.
//
// Fields chosen deliberately:
//   • IP address         — changes on different networks
//   • User-Agent         — identifies browser/OS combination
//   • Accept-Language    — locale setting, rarely spoofed
//   • Accept-Encoding    — compression preferences, browser-specific
//
// This is not a perfect fingerprint (it can collide on shared IPs or
// identical browsers), but it raises the bar significantly.
// ─────────────────────────────────────────────────────────────────────────────
const getRequestFingerprint = (req) => {
    const ua = req.headers["user-agent"] || "";
    const lang = req.headers["accept-language"] || "";
    const encoding = req.headers["accept-encoding"] || "";
    const ip = req.ip || req.socket.remoteAddress || "";
    return node_crypto_1.default
        .createHash("sha256")
        .update(`${ip}:${ua}:${lang}:${encoding}`)
        .digest("hex")
        .slice(0, 32);
};
exports.getRequestFingerprint = getRequestFingerprint;
// ─────────────────────────────────────────────────────────────────────────────
// PROGRESSIVE DELAY (per IP)
// After failed attempts, we make each subsequent attempt slower.
// Backoff schedule: 0s → 1s → 2s → 4s → 8s → capped at 30s
// This serializes automated attacks without blocking legitimate users much.
//
// NOTE: This uses in-process memory. In a multi-instance deployment,
// replace the Map with a Redis-backed store for consistency.
// ─────────────────────────────────────────────────────────────────────────────
const ipFailedAttempts = new Map();
const getProgressiveDelay = (ip) => {
    const record = ipFailedAttempts.get(ip);
    if (!record)
        return 0;
    // Forget failures older than 15 minutes
    if (Date.now() - record.lastAttempt > 15 * 60 * 1000) {
        ipFailedAttempts.delete(ip);
        return 0;
    }
    return Math.min(Math.pow(2, record.count - 1) * 1000, 30000);
};
exports.getProgressiveDelay = getProgressiveDelay;
const recordFailedAttempt = (ip) => {
    const record = ipFailedAttempts.get(ip) || { count: 0, lastAttempt: 0 };
    ipFailedAttempts.set(ip, { count: record.count + 1, lastAttempt: Date.now() });
};
exports.recordFailedAttempt = recordFailedAttempt;
const clearFailedAttempts = (ip) => {
    ipFailedAttempts.delete(ip);
};
exports.clearFailedAttempts = clearFailedAttempts;
// Middleware version — apply directly in the route chain
const progressiveDelayMiddleware = async (req, _res, next) => {
    const ip = req.ip || req.socket.remoteAddress || "unknown";
    const delay = (0, exports.getProgressiveDelay)(ip);
    if (delay > 0) {
        await new Promise(resolve => setTimeout(resolve, delay));
    }
    next();
};
exports.progressiveDelayMiddleware = progressiveDelayMiddleware;
const accountLockouts = new Map();
exports.LOCKOUT_THRESHOLD = 5;
exports.LOCKOUT_DURATION = 15 * 60 * 1000; // 15 minutes in ms
const checkAccountLockout = (email) => {
    const key = email.toLowerCase();
    const record = accountLockouts.get(key);
    if (!record)
        return { isLocked: false, remainingMs: 0 };
    if (record.lockedUntil && Date.now() < record.lockedUntil) {
        return { isLocked: true, remainingMs: record.lockedUntil - Date.now() };
    }
    // Lock has expired — clean it up
    if (record.lockedUntil && Date.now() >= record.lockedUntil) {
        accountLockouts.delete(key);
    }
    return { isLocked: false, remainingMs: 0 };
};
exports.checkAccountLockout = checkAccountLockout;
const recordFailedPasswordAttempt = (email) => {
    const key = email.toLowerCase();
    const existing = accountLockouts.get(key) || { failCount: 0, lockedUntil: null };
    const newCount = existing.failCount + 1;
    accountLockouts.set(key, {
        failCount: newCount,
        lockedUntil: newCount >= exports.LOCKOUT_THRESHOLD
            ? Date.now() + exports.LOCKOUT_DURATION
            : null,
    });
};
exports.recordFailedPasswordAttempt = recordFailedPasswordAttempt;
const clearAccountLockout = (email) => {
    accountLockouts.delete(email.toLowerCase());
};
exports.clearAccountLockout = clearAccountLockout;
