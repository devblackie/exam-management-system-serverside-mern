"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.csrfProtection = exports.attachCsrfToken = void 0;
const node_crypto_1 = __importDefault(require("node:crypto"));
/** Attach a CSRF token cookie on every GET (readable by JS, not HttpOnly). */
const attachCsrfToken = (_req, res, next) => {
    // Reuse existing token if already set this session
    if (!_req.cookies?.csrfToken) {
        const token = node_crypto_1.default.randomBytes(32).toString("hex");
        res.cookie("csrfToken", token, {
            httpOnly: false, // MUST be false — JS reads it to put in header
            sameSite: "strict",
            secure: process.env.NODE_ENV === "production",
            path: "/",
        });
    }
    next();
};
exports.attachCsrfToken = attachCsrfToken;
/** Verify the double-submit cookie on every state-changing request. */
const csrfProtection = (req, res, next) => {
    if (["GET", "HEAD", "OPTIONS"].includes(req.method)) {
        next();
        return;
    }
    const fromHeader = req.headers["x-csrf-token"];
    const fromCookie = req.cookies?.csrfToken;
    if (!fromHeader || !fromCookie) {
        res.status(403).json({ message: "CSRF token missing" });
        return;
    }
    try {
        const valid = node_crypto_1.default.timingSafeEqual(Buffer.from(fromHeader, "hex"), Buffer.from(fromCookie, "hex"));
        if (!valid)
            throw new Error("mismatch");
        next();
    }
    catch {
        res.status(403).json({ message: "CSRF token invalid" });
    }
};
exports.csrfProtection = csrfProtection;
