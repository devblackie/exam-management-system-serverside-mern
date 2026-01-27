// src/routes/admin.ts
import { Router, Request, Response } from "express";
import bcrypt from "bcryptjs";
import crypto from "crypto";
import User from "../models/User";
import Invite from "../models/Invite";
import AuditLog from "../models/AuditLog";
import { sendInviteEmail } from "../config/email";
import { requireAuth, requireRole } from "../middleware/auth";
import { cleanupOrphanedGrades } from "../scripts/cleanupGrades";

const router = Router();



// 📋 Get all invites
router.get("/invites", requireAuth, requireRole("admin"), async (_, res: Response) => {
  const invites = await Invite.find().sort({ createdAt: -1 });
  res.json(invites);
});

// 🔑 Admin secret registration
router.post("/secret-register", async (req: Request, res: Response) => {
  try {
    const { secret, name, email, password } = req.body;

    if (secret !== process.env.ADMIN_SECRET) {
      return res.status(403).json({ message: "Invalid secret" });
    }

    const hashed = await bcrypt.hash(password, 10);
    const admin = new User({ name, email, password: hashed, role: "admin" });
    await admin.save();

    res.json({ message: "Admin registered successfully" });
  } catch {
    res.status(500).json({ message: "Server error" });
  }
});

// 📧 Create invite
router.post("/invite", requireAuth, requireRole("admin"), async (req: Request, res: Response) => {
  const { email, role, name } = req.body;

  if (!["lecturer", "coordinator"].includes(role)) {
    return res.status(400).json({ message: "Invalid role" });
  }

  const defaultName = email
    .split("@")[0]
    .split(".")
    .map((p: string) => p.charAt(0).toUpperCase() + p.slice(1))
    .join(" ");

  const finalName = name?.trim() || defaultName;
  const token = crypto.randomBytes(20).toString("hex");

  const invite = new Invite({
    name: finalName,
    email,
    token,
    role,
    used: false,
    expiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
    createdBy: req.user?._id,
  });

  await invite.save();
  await sendInviteEmail(email, token, finalName);

  Promise.resolve(
    AuditLog.create({
      action: "invite_created",
      actor: req.user?._id,
      targetUser: invite._id,
      details: { email, role },
    })
  );

  res.json({ message: `Invite sent to ${finalName}` });
});

// 📝 Register with invite
router.post("/register/:token", async (req: Request, res: Response) => {
  const { token } = req.params;
  const { password } = req.body;

  const invite = await Invite.findOne({ token, used: false, expiresAt: { $gt: new Date() } });
  if (!invite) return res.status(400).json({ message: "Invalid or expired invite" });

  const hashed = await bcrypt.hash(password, 10);
  const user = new User({ name: invite.name, email: invite.email, password: hashed, role: invite.role });
  await user.save();

  invite.used = true;
  await invite.save();

  res.json({ message: "Account created successfully" });
});



// ❌ Revoke invite
router.delete("/invites/:id", requireAuth, requireRole("admin"), async (req: Request, res: Response) => {
  const invite = await Invite.findById(req.params.id);
  if (!invite) return res.status(404).json({ message: "Invite not found" });

  await invite.deleteOne();
  Promise.resolve(
    AuditLog.create({
      action: "invite_revoked",
      actor: req.user?._id,
      targetUser: invite._id,
      details: { email: invite.email, role: invite.role },
    })
  );
  res.json({ message: "Invite revoked successfully" });
});

// 👥 Get all users
router.get("/users", requireAuth, requireRole("admin"), async (_, res: Response) => {
  const users = await User.find().select("-password");
  res.json(users);
});

// 🔄 Update role
router.put("/users/:id/role", requireAuth, requireRole("admin"), async (req: Request, res: Response) => {
  const { id } = req.params;
  const { role } = req.body;

  if (!["admin", "lecturer", "coordinator"].includes(role)) {
    return res.status(400).json({ message: "Invalid role" });
  }

  if (req.user?._id.toString() === id) {
    return res.status(403).json({ message: "You cannot change your own role" });
  }

  const user = await User.findById(id);
  if (!user) return res.status(404).json({ message: "User not found" });

  if (user.role === "admin" && role !== "admin") {
    const adminCount = await User.countDocuments({ role: "admin" });
    if (adminCount <= 1) {
      return res.status(403).json({ message: "Cannot remove the last admin" });
    }
  }

  const oldRole = user.role;
  user.role = role;
  await user.save();

Promise.resolve(
    AuditLog.create({
      action: "role_changed",
      actor: req.user?._id,
      targetUser: user._id,
      details: { from: oldRole, to: role },
    })
  );

  res.json({ message: "Role updated", user });
});

// 🚦 Toggle status
router.put("/users/:id/status", requireAuth, requireRole("admin"), async (req: Request, res: Response) => {
  const { id } = req.params;
  const { status } = req.body;

  const user = await User.findById(id);
  if (!user) return res.status(404).json({ message: "User not found" });

  const oldStatus = user.status;
  user.status = status;
  await user.save();

  Promise.resolve(
    AuditLog.create({
      action: "status_toggled",
      actor: req.user?._id,
      targetUser: user._id,
      details: { from: oldStatus, to: status },
    })
  );

  res.json({ message: "Status updated", user });
});

// 🗑️ Delete user
router.delete("/users/:id", requireAuth, requireRole("admin"), async (req: Request, res: Response) => {
  if (req.user?._id.toString() === req.params.id) {
    return res.status(403).json({ message: "You cannot delete yourself" });
  }

  const user = await User.findById(req.params.id);
  if (!user) return res.status(404).json({ message: "User not found" });

  if (user.role === "admin") {
    const adminCount = await User.countDocuments({ role: "admin" });
    if (adminCount <= 1) {
      return res.status(403).json({ message: "Cannot delete the last admin" });
    }
  }

  await User.findByIdAndDelete(req.params.id);

    Promise.resolve(
    AuditLog.create({
      action: "user_deleted",
      actor: req.user?._id,
      targetUser: user._id,
      details: { email: user.email, role: user.role },
    })
  );

  res.json({ message: "User deleted" });
});

// 👨‍🏫 Get all lecturers
router.get("/lecturers", requireAuth, requireRole("admin","coordinator"), async (_, res: Response) => {
  try {
    const lecturers = await User.find({ role: "lecturer" }).select("-password");
    res.json(lecturers);
  } catch (err) {
    res.status(500).json({ message: "Server error" });
  }
});


export default router;
