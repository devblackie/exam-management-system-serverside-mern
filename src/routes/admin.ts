// serverside/src/routes/admin.ts

import { Router, Request, Response } from "express";
import bcrypt from "bcryptjs";
import crypto from "node:crypto";
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
import { ApiError } from "../middleware/errorHandler";

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
router.post("/secret-register", sanitizeInput,
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
router.get("/invites", requireAuth,
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

// serverside/src/routes/admin.ts — REPLACE the invite POST and register routes

router.post(
  "/invite",
  requireAuth,
  requireRole("admin"),
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response): Promise<void> => {
    const auth = req as AuthenticatedRequest;

    const {
      email,
      role,
      name,
      schoolCode,
      departmentCode,
      institutionWide,
    } = req.body as {
      email:            string;
      role:             "lecturer" | "coordinator";
      name?:            string;
      schoolCode?:      string;
      departmentCode?:  string;
      institutionWide?: boolean;
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
      res.status(403).json({ message: "Your account is not linked to an institution." });
      return;
    }

    // Duplicate checks
    const existingUser = await User.findOne({
      email:       email.toLowerCase(),
      institution: auth.user.institution,
    }).lean();
    if (existingUser) {
      res.status(409).json({ message: `An account for ${email} already exists.` });
      return;
    }

    const existingInvite = await Invite.findOne({
      email:       email.toLowerCase(),
      used:        false,
      institution: auth.user.institution,
      expiresAt:   { $gt: new Date() },
    }).lean();
    if (existingInvite) {
      res.status(409).json({
        message: `An active invite for ${email} already exists. Revoke it first.`,
      });
      return;
    }

    // Resolve human-readable school/department names for the email
    // Load institution settings to get the names
    const InstitutionSettings = (await import("../models/InstitutionSettings")).default;
    const Institution         = (await import("../models/Institution")).default;

    const [settingsDoc, institutionDoc] = await Promise.all([
      InstitutionSettings.findOne({ institution: auth.user.institution })
        .select("schools docMeta")
        .lean() as Promise<{
          schools?: Array<{
            code: string; name: string;
            departments?: Array<{ code: string; name: string }>;
          }>;
          docMeta?: { universityName?: string };
        } | null>,
      Institution.findById(auth.user.institution).select("name").lean() as Promise<{
        name: string;
      } | null>,
    ]);

    // Resolve names — prefer docMeta.universityName, fall back to Institution.name
    const universityName = settingsDoc?.docMeta?.universityName
      ?? institutionDoc?.name
      ?? "University";

    let resolvedSchoolName:      string | undefined;
    let resolvedDepartmentName:  string | undefined;

    if (schoolCode && settingsDoc?.schools) {
      const school = settingsDoc.schools.find(
        s => s.code === schoolCode.toUpperCase(),
      );
      resolvedSchoolName = school?.name;

      if (departmentCode && school?.departments) {
        const dept = school.departments.find(
          d => d.code === departmentCode.toUpperCase(),
        );
        resolvedDepartmentName = dept?.name;
      }
    }

    const finalName = name?.trim() || email
      .split("@")[0]
      .split(".")
      .map((p: string) => p.charAt(0).toUpperCase() + p.slice(1))
      .join(" ");

    const token     = crypto.randomBytes(24).toString("hex");
    const hasDepts  = !!(settingsDoc?.schools?.some(s => (s.departments?.length ?? 0) > 0));

    await Invite.create({
      name:            finalName,
      email:           email.toLowerCase(),
      token,
      role,
      used:            false,
      expiresAt:       new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
      createdBy:       auth.user._id,
      institution:     auth.user.institution,
      schoolCode:      schoolCode?.toUpperCase()    ?? null,
      departmentCode:  departmentCode?.toUpperCase() ?? null,
      institutionWide: institutionWide ?? (role === "lecturer" || !hasDepts),
    });

    // Send rich invitation email with university/school/department context
    const { sendInviteEmail } = await import("../config/email");
    sendInviteEmail({
      to:             email,
      token,
      name:           finalName,
      role,
      universityName,
      schoolName:     resolvedSchoolName,
      departmentName: resolvedDepartmentName,
      institutionWide: institutionWide ?? (role === "lecturer" || !hasDepts),
    }).catch((err: Error) => {
      console.error("[Admin] Invite email failed:", err.message);
    });

    await logAudit(auth, {
      action:  "invite_created",
      details: {
        email, role, schoolCode, departmentCode,
        institution: auth.user.institution.toString(),
      },
    });

    res.status(201).json({ message: `Invite sent to ${finalName}` });
  }),
);



// POST /admin/register/:token
router.post(
  "/register/:token",
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response): Promise<void> => {
    const { token }    = req.params;
    const { password } = req.body as { password: string };

    if (!password || password.length < 8) {
      res.status(400).json({ message: "Password must be at least 8 characters." });
      return;
    }

    interface InviteLean {
      _id:             mongoose.Types.ObjectId;
      name:            string;
      email:           string;
      token:           string;
      role:            "lecturer" | "coordinator";
      used:            boolean;
      expiresAt:       Date;
      institution:     mongoose.Types.ObjectId;
      schoolCode?:     string | null;
      departmentCode?: string | null;
      institutionWide?: boolean;
    }

    const invite = await Invite.findOne({
      token,
      used:      false,
      expiresAt: { $gt: new Date() },
    }).lean() as InviteLean | null;

    if (!invite) {
      res.status(400).json({ message: "Invite link is invalid or has expired." });
      return;
    }

    const existingUser = await User.findOne({ email: invite.email }).lean();
    if (existingUser) {
      res.status(409).json({ message: "An account with this email already exists." });
      return;
    }

    const hashed = await bcrypt.hash(password, 12);

    await User.create({
      name:            invite.name,
      email:           invite.email,
      password:        hashed,
      role:            invite.role,
      status:          "active",
      institution:     invite.institution,
      schoolCode:      invite.schoolCode     ?? null,
      departmentCode:  invite.departmentCode ?? null,
      institutionWide: invite.institutionWide ?? (invite.role === "lecturer"),
    });

    await Invite.updateOne({ _id: invite._id }, { used: true });

    await AuditLog.create({
      action:  "invite_used",
      details: { email: invite.email, role: invite.role },
    }).catch((e: Error) => console.error("[AuditLog]", e.message));

    res.status(201).json({ message: "Account created successfully. You can now log in." });
  }),
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

// PUT /admin/users/:id/details - Update coordinator details (school/dept)
router.put(
  "/users/:id/details",
  requireAuth,
  requireRole("admin"),
  sanitizeInput,
  asyncHandler(async (req: Request, res: Response) => {
    const auth = req as AuthenticatedRequest;
    const { id } = req.params;
    const { name, schoolCode, departmentCode, institutionWide } = req.body as {
      name?: string;
      schoolCode?: string;
      departmentCode?: string;
      institutionWide?: boolean;
    };

    // Prevent self-modification
    if (auth.user._id.toString() === id) {
      throw {
        statusCode: 403,
        message: "You cannot modify your own account details.",
      } as ApiError;
    }

    const user = await User.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
      throw { statusCode: 404, message: "User not found in your institution." } as ApiError;
    }

    // Only coordinators can have school/department assignments
    if (user.role !== "coordinator" && (schoolCode !== undefined || departmentCode !== undefined || institutionWide !== undefined)) {
      throw {
        statusCode: 400,
        message: "School/department assignment only applies to coordinators.",
      } as ApiError;
    }

    // For coordinators: validate school/department exist if provided
    if (user.role === "coordinator" && !institutionWide && schoolCode && departmentCode) {
      const InstitutionSettings = (await import("../models/InstitutionSettings")).default;
      const settings = await InstitutionSettings.findOne({ institution: auth.user.institution })
        .select("schools")
        .lean() as {
          schools?: Array<{
            code: string;
            departments?: Array<{ code: string }>;
          }>;
        } | null;

      const school = settings?.schools?.find(s => s.code === schoolCode.toUpperCase());
      if (!school) {
        throw { statusCode: 400, message: `School "${schoolCode}" not found.` } as ApiError;
      }

      const department = school.departments?.find(d => d.code === departmentCode.toUpperCase());
      if (!department) {
        throw { statusCode: 400, message: `Department "${departmentCode}" not found in school "${schoolCode}".` } as ApiError;
      }
    }

    // Apply updates
    if (name) user.name = name;
    
    // Handle schoolCode - convert empty string to undefined (Mongoose will ignore undefined)
    if (schoolCode !== undefined) {
      user.schoolCode = schoolCode && schoolCode.trim() !== "" ? schoolCode.toUpperCase() : undefined;
    }
    
    // Handle departmentCode - convert empty string to undefined
    if (departmentCode !== undefined) {
      user.departmentCode = departmentCode && departmentCode.trim() !== "" ? departmentCode.toUpperCase() : undefined;
    }
    
    if (institutionWide !== undefined) user.institutionWide = institutionWide;

    // If institution-wide is true, clear school/department assignments
    if (institutionWide === true) {
      user.schoolCode = undefined;
      user.departmentCode = undefined;
    }

    await user.save();

    await logAudit(auth, {
      action: "user_details_updated",
      targetUser: user._id,
      details: { 
        name: name || user.name, 
        schoolCode: user.schoolCode, 
        departmentCode: user.departmentCode, 
        institutionWide: user.institutionWide 
      },
    });

    const updatedUser = await User.findById(id).select("-password").lean();
    res.json({ message: "User details updated successfully", user: updatedUser });
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
