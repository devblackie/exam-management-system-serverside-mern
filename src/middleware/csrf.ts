// serverside/src/middleware/csrf.ts
import { Request, Response, NextFunction } from "express";
import crypto from "node:crypto";

/** Attach a CSRF token cookie on every GET (readable by JS, not HttpOnly). */
export const attachCsrfToken = (
  _req: Request,
  res:  Response,
  next: NextFunction,
): void => {
  // Reuse existing token if already set this session
  if (!_req.cookies?.csrfToken) {
    const token = crypto.randomBytes(32).toString("hex");
    res.cookie("csrfToken", token, {
      httpOnly: false,           // MUST be false — JS reads it to put in header
      sameSite: "strict",
      secure:   process.env.NODE_ENV === "production",
      path:     "/",
    });
  }
  next();
};

/** Verify the double-submit cookie on every state-changing request. */
export const csrfProtection = (
  req:  Request,
  res:  Response,
  next: NextFunction,
): void => {
  if (["GET", "HEAD", "OPTIONS"].includes(req.method)) {
    next();
    return;
  }

  const fromHeader = req.headers["x-csrf-token"] as string | undefined;
  const fromCookie = req.cookies?.csrfToken        as string | undefined;

  if (!fromHeader || !fromCookie) {
    res.status(403).json({ message: "CSRF token missing" });
    return;
  }

  try {
    const valid = crypto.timingSafeEqual(
      Buffer.from(fromHeader, "hex"),
      Buffer.from(fromCookie,  "hex"),
    );
    if (!valid) throw new Error("mismatch");
    next();
  } catch {
    res.status(403).json({ message: "CSRF token invalid" });
  }
};

