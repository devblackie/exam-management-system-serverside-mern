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

const router = Router();

// 🔑 Login
router.post(
  "/login",
  asyncHandler(async (req: Request, res: Response) => {
    const { email, password } = req.body;

    // Find user and populate institution if exists
    const user = await User.findOne({ email: email.toLowerCase() })
      .select("+password") // include password (if select: false in schema)
      .lean();

    if (!user) {
      throw { statusCode: 401, message: "Invalid credentials" } as ApiError;
    }

    const valid = await bcrypt.compare(password, user.password);
    if (!valid) {
      throw { statusCode: 401, message: "Invalid credentials" } as ApiError;
    }

    if (user.status === "suspended") {
      throw { statusCode: 403, message: "Account suspended" } as ApiError;
    }

    // CRITICAL: Include institution in JWT
    setAuthCookie(
      res,
      user._id.toString(),
      user.role,
      user.institution?.toString() // ← This is the key!
    );

    // Audit log (non-blocking)
    logAudit(req, {
      action: "login_success",
      actor: user._id,
      details: {
        email: user.email,
        role: user.role,
        institution: user.institution,
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
  })
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


export default router;
