// import rateLimit from "express-rate-limit";
// import sanitize from "mongo-sanitize";
// import { Request, Response, NextFunction } from "express";
// // import { verifyToken } from "../lib/jwt";
// // import User from "../models/User";


// export const loginRateLimiter = rateLimit({
//   windowMs: 15 * 60 * 1000, // 15 minutes
//   max: 5, // 5 attempts per IP
//   message: {
//     message: "Too many attempts. Access locked for 15m.",
//   },
//   standardHeaders: true,
//   legacyHeaders: false,
// });


// export const sanitizeInput = (
//   req: Request,
//   res: Response,
//   next: NextFunction,
// ) => {
//   if (req.body) {
//     req.body = sanitize(req.body);
//   }

//   // For Query and Params, we clean the keys/values individually
//   // to avoid the "only a getter" TypeError
//   if (req.query) {
//     Object.keys(req.query).forEach((key) => {
//       req.query[key] = sanitize(req.query[key]);
//     });
//   }

//   if (req.params) {
//     Object.keys(req.params).forEach((key) => {
//       req.params[key] = sanitize(req.params[key]);
//     });
//   }

//   next();
// };


// export async function requireAuth(
//   req: Request,
//   res: Response,
//   next: NextFunction,
// ) {
//   const token = req.cookies?.token;
//   if (!token)
//     return res.status(401).json({ message: "Identity verification required" });

//   try {
//     const payload = verifyToken(token) as any;

//     // Session validation: Prevents suspended users from using "Zombie Tokens"
//     const userDoc = await User.findById(payload.id)
//       .select("status role institution tokenVersion")
//       .lean();

//     if (!userDoc || userDoc.status === "suspended") {
//       res.clearCookie("token");
//       return res
//         .status(403)
//         .json({ message: "Session revoked. Access denied." });
//     }

//     if (payload.version !== userDoc.tokenVersion) {
//       res.clearCookie("token");
//       return res
//         .status(401)
//         .json({ message: "Session expired due to security update." });
//     }
//     // Attach Context: This is your "Logical RLS"
//     // Every downstream query must use req.user.institution
//     req.user = {
//       ...userDoc,
//       _id: userDoc._id,
//       institution: userDoc.institution,
//     };

//     next();
//   } catch (err) {
//     res.status(401).json({ message: "Session expired or invalid" });
//   }
// }

// serverside/src/middleware/security.ts
//
// What lives here:
//   • Rate limiters (one per endpoint, tuned to its attack surface)
//   • Input sanitization (mongo-sanitize wrapper)
//   • Honeypot field check (bot detection)
//   • Request fingerprinting (device binding for step cookies)
//   • Progressive delay (exponential backoff after failed attempts)
//   • Account lockout (per-email, defeats distributed brute force)
//   • Security headers (helmet wrapper)
//
// What does NOT live here:
//   • requireAuth — lives in src/middleware/auth.ts (unchanged)
//   • requireRole — lives in src/middleware/auth.ts (unchanged)
//
// Your existing requireAuth and requireRole in auth.ts are KEPT AS-IS.
// This file is purely additive — it adds new security primitives that
// the auth routes import alongside the existing middleware.

import rateLimit       from "express-rate-limit";
import sanitize        from "mongo-sanitize";
import { Request, Response, NextFunction } from "express";
import crypto          from "crypto";
import helmet from "helmet";

// ─────────────────────────────────────────────────────────────────────────────
// RATE LIMITERS
// One limiter per endpoint — each tuned to the threat model of that endpoint.
// ─────────────────────────────────────────────────────────────────────────────

// Step 1: Email lookup — generous (no secret submitted yet)
export const emailCheckLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,  // 15 minutes
  max:      20,               // 20 email checks per IP
  message:  { message: "Too many requests. Please wait before trying again." },
  standardHeaders: true,
  legacyHeaders:   false,
  skipSuccessfulRequests: false,
});

// Step 2: Password verify — tight (a secret is being submitted)
export const passwordLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max:      5,                // 5 attempts per IP per 15 min
  message:  { message: "Too many login attempts. Access locked for 15 minutes." },
  standardHeaders: true,
  legacyHeaders:   false,
});

// Step 3: OTP verify — very tight (brute-force window is 10 min)
export const otpLimiter = rateLimit({
  windowMs: 10 * 60 * 1000,  // matches OTP lifetime
  max:      5,
  message:  { message: "Too many verification attempts. Please log in again." },
  standardHeaders: true,
  legacyHeaders:   false,
});

// General API limiter (apply globally in app.ts if desired)
export const apiLimiter = rateLimit({
  windowMs: 60 * 1000,
  max:      120,
  message:  { message: "Too many requests." },
  standardHeaders: true,
  legacyHeaders:   false,
});

// Registration — very slow to prevent mass account creation
export const registrationLimiter = rateLimit({
  windowMs: 60 * 60 * 1000,  // 1 hour
  max:      3,
  message:  { message: "Too many registration attempts. Please try again later." },
  standardHeaders: true,
  legacyHeaders:   false,
});

// ── Backward-compat alias ─────────────────────────────────────────────────────
// Your existing code imports loginRateLimiter from this file.
// Keep it pointing at the password limiter so nothing breaks.
export const loginRateLimiter = passwordLimiter;

// ─────────────────────────────────────────────────────────────────────────────
// INPUT SANITIZATION
// Strips MongoDB operators ($gt, $where, etc.) from every incoming request.
// Prevents NoSQL injection attacks.
// Your existing sanitizeInput is kept exactly as-is — just re-exported.
// ─────────────────────────────────────────────────────────────────────────────

export const sanitizeInput = (
  req:  Request,
  res:  Response,
  next: NextFunction,
): void => {
  if (req.body) {
    req.body = sanitize(req.body);
  }

  // Query and Params need key-by-key sanitization to avoid "only a getter" TypeError
  if (req.query) {
    Object.keys(req.query).forEach(key => {
      req.query[key] = sanitize(req.query[key]);
    });
  }

  if (req.params) {
    Object.keys(req.params).forEach(key => {
      req.params[key] = sanitize(req.params[key]);
    });
  }

  next();
};

// ─────────────────────────────────────────────────────────────────────────────
// HONEYPOT FIELD CHECK
// Every auth form includes a hidden <input name="website"> that real users
// never see or fill. Bots that auto-fill all fields are silently dropped.
// We return a success-looking response so bots don't know they were caught.
// ─────────────────────────────────────────────────────────────────────────────

export const honeypotCheck = (
  req:  Request,
  res:  Response,
  next: NextFunction,
): void => {
  if (req.body?.website || req.body?._gotcha) {
    // Silent drop — return a plausible response, log nothing useful to the bot
    res.json({ nextStep: "password" });
    return;
  }
  next();
};

export const securityHeaders = helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc:  ["'self'"],
      scriptSrc:   ["'self'"],
      styleSrc:    ["'self'", "'unsafe-inline'"],
      imgSrc:      ["'self'", "data:", "blob:"],
      connectSrc:  ["'self'"],
      fontSrc:     ["'self'"],
      objectSrc:   ["'none'"],
      frameSrc:    ["'none'"],
      upgradeInsecureRequests: [],
    },
  },
  hsts: {
    maxAge:            31536000,  // 1 year
    includeSubDomains: true,
    preload:           true,
  },
  noSniff:          true,
  xssFilter:        true,
  referrerPolicy:   { policy: "strict-origin-when-cross-origin" },
  frameguard:       { action: "deny" },
  permittedCrossDomainPolicies: false,
});
 
// Additional custom headers
export const additionalSecurityHeaders = (
  _req: Request,
  res: Response,
  next: NextFunction
): void => {
  res.setHeader("X-Content-Type-Options",       "nosniff");
  res.setHeader("X-Frame-Options",              "DENY");
  res.setHeader("X-XSS-Protection",             "1; mode=block");
  res.setHeader("Permissions-Policy",
    "camera=(), microphone=(), geolocation=(), payment=()");
  res.setHeader("Cache-Control",
    "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.setHeader("Pragma",                       "no-cache");
  res.setHeader("Expires",                      "0");
  next();
};
 
// ─────────────────────────────────────────────────────────────────────────────
// SUSPICIOUS ACTIVITY DETECTOR
// Flags requests that look automated or suspicious.
// ─────────────────────────────────────────────────────────────────────────────
 
export const detectSuspiciousRequest = (req: Request): boolean => {
  const ua = (req.headers["user-agent"] || "").toLowerCase();
 
  const suspiciousPatterns = [
    "curl", "wget", "python-requests", "go-http-client",
    "java/", "libwww-perl", "scrapy", "httpie",
    "axios/0",  // raw axios (not a browser)
    "postman",  // allow in dev, block in prod
  ];
 
  if (process.env.NODE_ENV === "production") {
    return suspiciousPatterns.some(p => ua.includes(p));
  }
 
  // In dev, only flag clearly automated tools
  return ["scrapy", "sqlmap", "nikto", "masscan"].some(p => ua.includes(p));
};
 
export const blockSuspiciousRequests = (
  req: Request,
  res: Response,
  next: NextFunction
): void => {
  if (detectSuspiciousRequest(req)) {
    res.status(403).json({ message: "Forbidden" });
    return;
  }
  next();
};

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

export const getRequestFingerprint = (req: Request): string => {
  const ua       = req.headers["user-agent"]       || "";
  const lang     = req.headers["accept-language"]  || "";
  const encoding = req.headers["accept-encoding"]  || "";
  const ip       = req.ip || req.socket.remoteAddress || "";

  return crypto
    .createHash("sha256")
    .update(`${ip}:${ua}:${lang}:${encoding}`)
    .digest("hex")
    .slice(0, 32);
};

// ─────────────────────────────────────────────────────────────────────────────
// PROGRESSIVE DELAY (per IP)
// After failed attempts, we make each subsequent attempt slower.
// Backoff schedule: 0s → 1s → 2s → 4s → 8s → capped at 30s
// This serializes automated attacks without blocking legitimate users much.
//
// NOTE: This uses in-process memory. In a multi-instance deployment,
// replace the Map with a Redis-backed store for consistency.
// ─────────────────────────────────────────────────────────────────────────────

const ipFailedAttempts = new Map<string, { count: number; lastAttempt: number }>();

export const getProgressiveDelay = (ip: string): number => {
  const record = ipFailedAttempts.get(ip);
  if (!record) return 0;

  // Forget failures older than 15 minutes
  if (Date.now() - record.lastAttempt > 15 * 60 * 1000) {
    ipFailedAttempts.delete(ip);
    return 0;
  }

  return Math.min(Math.pow(2, record.count - 1) * 1000, 30_000);
};

export const recordFailedAttempt = (ip: string): void => {
  const record = ipFailedAttempts.get(ip) || { count: 0, lastAttempt: 0 };
  ipFailedAttempts.set(ip, { count: record.count + 1, lastAttempt: Date.now() });
};

export const clearFailedAttempts = (ip: string): void => {
  ipFailedAttempts.delete(ip);
};

// Middleware version — apply directly in the route chain
export const progressiveDelayMiddleware = async (
  req:  Request,
  _res: Response,
  next: NextFunction,
): Promise<void> => {
  const ip    = req.ip || req.socket.remoteAddress || "unknown";
  const delay = getProgressiveDelay(ip);
  if (delay > 0) {
    await new Promise(resolve => setTimeout(resolve, delay));
  }
  next();
};

// ─────────────────────────────────────────────────────────────────────────────
// ACCOUNT LOCKOUT (per email address)
// After 5 wrong passwords for a specific account, lock it for 15 minutes.
// This defeats distributed attacks where many IPs each guess once —
// the IP-level rate limiter won't catch them, but the per-account lockout will.
//
// NOTE: Same in-process memory caveat as above. Use Redis for multi-instance.
// ─────────────────────────────────────────────────────────────────────────────

interface LockoutRecord {
  failCount:   number;
  lockedUntil: number | null;
}

const accountLockouts = new Map<string, LockoutRecord>();

export const LOCKOUT_THRESHOLD = 5;
export const LOCKOUT_DURATION  = 15 * 60 * 1000; // 15 minutes in ms

export const checkAccountLockout = (
  email: string,
): { isLocked: boolean; remainingMs: number } => {
  const key    = email.toLowerCase();
  const record = accountLockouts.get(key);

  if (!record) return { isLocked: false, remainingMs: 0 };

  if (record.lockedUntil && Date.now() < record.lockedUntil) {
    return { isLocked: true, remainingMs: record.lockedUntil - Date.now() };
  }

  // Lock has expired — clean it up
  if (record.lockedUntil && Date.now() >= record.lockedUntil) {
    accountLockouts.delete(key);
  }

  return { isLocked: false, remainingMs: 0 };
};

export const recordFailedPasswordAttempt = (email: string): void => {
  const key      = email.toLowerCase();
  const existing = accountLockouts.get(key) || { failCount: 0, lockedUntil: null };
  const newCount = existing.failCount + 1;

  accountLockouts.set(key, {
    failCount:   newCount,
    lockedUntil: newCount >= LOCKOUT_THRESHOLD
      ? Date.now() + LOCKOUT_DURATION
      : null,
  });
};

export const clearAccountLockout = (email: string): void => {
  accountLockouts.delete(email.toLowerCase());
};
