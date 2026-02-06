import rateLimit from "express-rate-limit";
import sanitize from "mongo-sanitize";
import { Request, Response, NextFunction } from "express";
import { verifyToken } from "../lib/jwt";
import User from "../models/User";


export const loginRateLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 5, // 5 attempts per IP
  message: {
    message: "Too many attempts. Access locked for 15m.",
  },
  standardHeaders: true,
  legacyHeaders: false,
});


export const sanitizeInput = (
  req: Request,
  res: Response,
  next: NextFunction,
) => {
  if (req.body) {
    req.body = sanitize(req.body);
  }

  // For Query and Params, we clean the keys/values individually
  // to avoid the "only a getter" TypeError
  if (req.query) {
    Object.keys(req.query).forEach((key) => {
      req.query[key] = sanitize(req.query[key]);
    });
  }

  if (req.params) {
    Object.keys(req.params).forEach((key) => {
      req.params[key] = sanitize(req.params[key]);
    });
  }

  next();
};


export async function requireAuth(
  req: Request,
  res: Response,
  next: NextFunction,
) {
  const token = req.cookies?.token;
  if (!token)
    return res.status(401).json({ message: "Identity verification required" });

  try {
    const payload = verifyToken(token) as any;

    // Session validation: Prevents suspended users from using "Zombie Tokens"
    const userDoc = await User.findById(payload.id)
      .select("status role institution tokenVersion")
      .lean();

    if (!userDoc || userDoc.status === "suspended") {
      res.clearCookie("token");
      return res
        .status(403)
        .json({ message: "Session revoked. Access denied." });
    }

    if (payload.version !== userDoc.tokenVersion) {
      res.clearCookie("token");
      return res
        .status(401)
        .json({ message: "Session expired due to security update." });
    }
    // Attach Context: This is your "Logical RLS"
    // Every downstream query must use req.user.institution
    req.user = {
      ...userDoc,
      _id: userDoc._id,
      institution: userDoc.institution,
    };

    next();
  } catch (err) {
    res.status(401).json({ message: "Session expired or invalid" });
  }
}
