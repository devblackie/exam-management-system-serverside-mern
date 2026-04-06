// serverside/src/routes/admin.ts
//
// KEY CHANGES FROM YOUR CURRENT FILE:
//
// 1. INSTITUTION INHERITANCE
//    Every invited user inherits req.user.institution from the admin who sends
//    the invite. The Invite model stores institutionId. The /register/:token
//    route reads it back and assigns it to the new User document.
//    This means coordinators and lecturers automatically belong to the same
//    institution as the admin who invited them — no manual assignment needed.
//
// 2. ASYNC ERROR HANDLING
//    All routes use asyncHandler so unhandled promise rejections produce
//    proper { message } JSON responses instead of crashing Express or
//    returning an empty 500 that triggers "Unknown error occurred" in the UI.
//
// 3. NO `any` TYPES
//    All request objects are typed via AuthenticatedRequest. Lean query
//    results are typed explicitly.
//
// 4. DUPLICATE EMAIL CHECK ON INVITE
//    Prevents inviting an email that already has an account — returns a
//    clear 409 message the UI can display.
//
// 5. ADMIN SECRET REGISTER ALSO ASSIGNS INSTITUTION
//    The bootstrap /admin/secret-register route now requires institutionId
//    so the first admin is also linked to an institution from day one.

import { Router, Request, Response } from "express";
import bcrypt from "bcryptjs";
import crypto from "crypto";
import mongoose from "mongoose";
import User from "../models/User";
import Invite from "../models/Invite";
import AuditLog from "../models/AuditLog";
import Institution from "../models/Institution";
import { sendInviteEmail } from "../config/email";
import { requireAuth, requireRole, AuthenticatedRequest } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { sanitizeInput } from "../middleware/security";
import { logAudit } from "../lib/auditLogger";

const router = Router();

interface InviteDoc {
  _id:           mongoose.Types.ObjectId;
  name:          string;
  email:         string;
  token:         string;
  role:          "lecturer" | "coordinator";
  used:          boolean;
  expiresAt:     Date;
  createdBy:     mongoose.Types.ObjectId;
  institution:   mongoose.Types.ObjectId;
}

// POST /admin/secret-register
router.post(
  "/secret-register",
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const { secret, name, email, password, institutionId } = req.body as {
      secret: string; name: string; email: string;
      password: string; institutionId: string;
    };

    if (secret !== process.env.ADMIN_SECRET) {
      res.status(403).json({ message: "Invalid Code" });
      return;
    }

    if (!institutionId || !mongoose.isValidObjectId(institutionId)) {
      res.status(400).json({ message: "A valid institution ID is required." });
      return;
    }

    const institution = await Institution.findById(institutionId).lean();
    if (!institution) {
      res.status(400).json({ message: "Institution not found." });
      return;
    }

    const existing = await User.findOne({ email: email.toLowerCase() }).lean();
    if (existing) {
      await logAudit(req, { action: "coordinator_register_failed_duplicate", details: { email } });
      res.status(409).json({ message: "An account with this email already exists." });
      return;
    }

    const hashed = await bcrypt.hash(password, 12);
    const admin  = await User.create({
      name,
      email:       email.toLowerCase(),
      password:    hashed,
      role:        "admin",
      status:      "active",
      institution: new mongoose.Types.ObjectId(institutionId),
    });

    await logAudit(req, {
      action: "admin_registered",
      targetUser: admin._id,
      details: { email, name,institution: institutionId },
    });

    res.status(201).json({ message: "Admin registered successfully", id: admin._id });
  })
);

// GET /admin/invites
router.get(
  "/invites",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: Request, res: Response) => {
    const auth = req as AuthenticatedRequest;
    const invites = await Invite.find({
      institution: auth.user.institution,
    })
      .sort({ createdAt: -1 })
      .lean();
    res.json(invites);
  })
);

// POST /admin/invite
// Creates an invite. The invited user inherits the admin's institution.
router.post(
  "/invite",
  requireAuth,
  requireRole("admin"),
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const auth = req as AuthenticatedRequest;    

    const { email, role, name } = req.body as {
      email: string;
      role:  "lecturer" | "coordinator";
      name?: string;
    };

    if (!email || !role) {
      res.status(400).json({ message: "Email and role are required." });
      return;
    }

    if (!["lecturer", "coordinator"].includes(role)) {
      res.status(400).json({ message: "Invalid role." });
      return;
    }

    if (!auth.user.institution) {
      res.status(403).json({
        message: "Your account is not linked to an institution. Contact a system administrator.",
      });
      return;
    }

    // Prevent duplicate invites for the same email at this institution
    const existingUser = await User.findOne({
      email:       email.toLowerCase(),
      institution: auth.user.institution,
    }).lean();

    if (existingUser) {
      res.status(409).json({
        message: `An account for ${email} already exists in this institution.`,
      });
      return;
    }

    const existingInvite = await Invite.findOne({
      email: email.toLowerCase(),
      used:  false,
      institution: auth.user.institution,
      expiresAt: { $gt: new Date() },
    }).lean();

    if (existingInvite) {
      res.status(409).json({
        message: `An active invite for ${email} already exists. Revoke it first if you want to resend.`,
      });
      return;
    }

    // Derive a default name from the email if not provided
    const finalName = name?.trim() || email
      .split("@")[0]
      .split(".")
      .map((p: string) => p.charAt(0).toUpperCase() + p.slice(1))
      .join(" ");

    const token  = crypto.randomBytes(24).toString("hex");
    const invite = await Invite.create({
      name:        finalName,
      email:       email.toLowerCase(),
      token,
      role,
      used:        false,
      expiresAt:   new Date(Date.now() + 7 * 24 * 60 * 60 * 1000), // 7 days
      createdBy:   auth.user._id,
      institution: auth.user.institution, // ← inherited from admin
    });

    // Send invite email (non-blocking — email failure doesn't break the invite)
    sendInviteEmail(email, token, finalName).catch((err: Error) => {
      console.error("[Admin] Invite email failed:", err.message);
    });

    await AuditLog.create({
      action:     "invite_created",
      actor:      auth.user._id,
      targetUser: invite._id,
      details:    { email, role, institution: auth.user.institution },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));
    
    res.status(201).json({ message: `Invite sent to ${finalName}` });
  })
);

// POST /admin/register/:token
// Completes registration from an invite link.
// The new user inherits the institution stored on the Invite document.

router.post(
  "/register/:token",
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const { token }    = req.params;
    const { password } = req.body as { password: string };

    if (!password || password.length < 8) {
      res.status(400).json({ message: "Password must be at least 8 characters." });
      return;
    }

    const invite = await Invite.findOne({
      token,
      used:      false,
      expiresAt: { $gt: new Date() },
    }).lean() as InviteDoc | null;

    if (!invite) {
      res.status(400).json({ message: "Invite link is invalid or has expired." });
      return;
    }

    // Prevent double-registration
    const existingUser = await User.findOne({ email: invite.email }).lean();
    if (existingUser) {
      res.status(409).json({ message: "An account with this email already exists." });
      return;
    }

    const hashed = await bcrypt.hash(password, 12);

    const newUser = await User.create({
      name:        invite.name,
      email:       invite.email,
      password:    hashed,
      role:        invite.role,
      status:      "active",
      institution: invite.institution, // ← inherited from invite (= admin's institution)
    });

    await Invite.updateOne({ _id: invite._id }, { used: true });

    await AuditLog.create({
      action:     "invite_used",
      actor:      newUser._id,
      targetUser: newUser._id,
      details:    { email: invite.email, role: invite.role },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));

    res.status(201).json({ message: "Account created successfully. You can now log in." });
  })
);

// DELETE /admin/invites/:id
router.delete(
  "/invites/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: Request, res: Response) => {
    const auth   = req as AuthenticatedRequest;
    const invite = await Invite.findOne({
      _id:         req.params.id,
      institution: auth.user.institution,
    }).lean();

    if (!invite) {
      res.status(404).json({ message: "Invite not found." });
      return;
    }

    await Invite.deleteOne({ _id: invite._id });

    await AuditLog.create({
      action:     "invite_revoked",
      actor:      auth.user._id,
      targetUser: invite._id,
      details:    { email: invite.email, role: invite.role },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));

    res.json({ message: "Invite revoked." });
  })
);

// GET /admin/users
router.get(
  "/users",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: Request, res: Response) => {
    const auth  = req as AuthenticatedRequest;
    const users = await User.find({ institution: auth.user.institution })
      .select("-password")
      .lean();
    res.json(users);
  })
);

// PUT /admin/users/:id/role
router.put(
  "/users/:id/role",
  requireAuth,
  requireRole("admin"),
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const auth         = req as AuthenticatedRequest;
    const { role }     = req.body as { role: string };
    const { id }       = req.params;

    if (!["admin", "lecturer", "coordinator"].includes(role)) {
      res.status(400).json({ message: "Invalid role." });
      return;
    }

    if (auth.user._id.toString() === id) {
      res.status(403).json({ message: "You cannot change your own role." });
      return;
    }

    const user = await User.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
      res.status(404).json({ message: "User not found in your institution." });
      return;
    }

    if (user.role === "admin" && role !== "admin") {
      const adminCount = await User.countDocuments({
        institution: auth.user.institution,
        role:        "admin",
      });
      if (adminCount <= 1) {
        res.status(403).json({ message: "Cannot demote the last admin." });
        return;
      }
    }

    const oldRole = user.role;
    user.role     = role as "admin" | "lecturer" | "coordinator";
    await user.save();

    await AuditLog.create({
      action:     "role_changed",
      actor:      auth.user._id,
      targetUser: user._id,
      details:    { from: oldRole, to: role },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));

    res.json({ message: "Role updated.", user });
  })
);

// PUT /admin/users/:id/status
router.put(
  "/users/:id/status",
  requireAuth,
  requireRole("admin"),
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const auth         = req as AuthenticatedRequest;
    const { status }   = req.body as { status: string };
    const { id }       = req.params;

    if (!["active", "suspended"].includes(status)) {
      res.status(400).json({ message: "Status must be 'active' or 'suspended'." });
      return;
    }

    const user = await User.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
      res.status(404).json({ message: "User not found in your institution." });
      return;
    }

    const oldStatus = user.status;
    user.status     = status as "active" | "suspended";
    await user.save();

    await AuditLog.create({
      action:     "status_toggled",
      actor:      auth.user._id,
      targetUser: user._id,
      details:    { from: oldStatus, to: status },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));

    res.json({ message: "Status updated.", user });
  })
);

// DELETE /admin/users/:id
router.delete(
  "/users/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: Request, res: Response) => {
    const auth = req as AuthenticatedRequest;
    const { id } = req.params;

    if (auth.user._id.toString() === id) {
      res.status(403).json({ message: "You cannot delete your own account." });
      return;
    }

    const user = await User.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
      res.status(404).json({ message: "User not found in your institution." });
      return;
    }

    if (user.role === "admin") {
      const adminCount = await User.countDocuments({
        institution: auth.user.institution,
        role:        "admin",
      });
      if (adminCount <= 1) {
        res.status(403).json({ message: "Cannot delete the last admin." });
        return;
      }
    }

    await User.deleteOne({ _id: id });

    await AuditLog.create({
      action:     "user_deleted",
      actor:      auth.user._id,
      targetUser: user._id,
      details:    { email: user.email, role: user.role },
    }).catch((err: Error) => console.error("[AuditLog]", err.message));

    res.json({ message: "User deleted." });
  })
);

// GET /admin/lecturers
router.get(
  "/lecturers",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: Request, res: Response) => {
    const auth     = req as AuthenticatedRequest;
    const lecturers = await User.find({
      institution: auth.user.institution,
      role:        "lecturer",
    })
      .select("-password")
      .lean();
    res.json(lecturers);
  })
);

export default router;
