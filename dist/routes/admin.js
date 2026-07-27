"use strict";
// serverside/src/routes/admin.ts
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = require("express");
const bcryptjs_1 = __importDefault(require("bcryptjs"));
const node_crypto_1 = __importDefault(require("node:crypto"));
const mongoose_1 = __importDefault(require("mongoose"));
const User_1 = __importDefault(require("../models/User"));
const Invite_1 = __importDefault(require("../models/Invite"));
const AuditLog_1 = __importDefault(require("../models/AuditLog"));
const Institution_1 = __importDefault(require("../models/Institution"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const security_1 = require("../middleware/security");
const auditLogger_1 = require("../lib/auditLogger");
const router = (0, express_1.Router)();
// POST /admin/secret-register
router.post("/secret-register", security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { secret, name, email, password, institutionId } = req.body;
    if (secret !== process.env.ADMIN_SECRET) {
        res.status(403).json({ message: "Invalid Code" });
        return;
    }
    if (!institutionId || !mongoose_1.default.isValidObjectId(institutionId)) {
        res.status(400).json({ message: "A valid institution ID is required." });
        return;
    }
    const institution = await Institution_1.default.findById(institutionId).lean();
    if (!institution) {
        res.status(400).json({ message: "Institution not found." });
        return;
    }
    const existing = await User_1.default.findOne({ email: email.toLowerCase() }).lean();
    if (existing) {
        await (0, auditLogger_1.logAudit)(req, { action: "coordinator_register_failed_duplicate", details: { email } });
        res.status(409).json({ message: "An account with this email already exists." });
        return;
    }
    const hashed = await bcryptjs_1.default.hash(password, 12);
    const admin = await User_1.default.create({
        name,
        email: email.toLowerCase(),
        password: hashed,
        role: "admin",
        status: "active",
        institution: new mongoose_1.default.Types.ObjectId(institutionId),
    });
    await (0, auditLogger_1.logAudit)(req, {
        action: "admin_registered",
        targetUser: admin._id,
        details: { email, name, institution: institutionId },
    });
    res.status(201).json({ message: "Admin registered successfully", id: admin._id });
}));
// GET /admin/invites
router.get("/invites", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const invites = await Invite_1.default.find({
        institution: auth.user.institution,
    })
        .sort({ createdAt: -1 })
        .lean();
    res.json(invites);
}));
// POST /admin/invite
router.post("/invite", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const { email, role, name, schoolCode, departmentCode, institutionWide, } = req.body;
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
    const existingUser = await User_1.default.findOne({
        email: email.toLowerCase(),
        institution: auth.user.institution,
    }).lean();
    if (existingUser) {
        res.status(409).json({ message: `An account for ${email} already exists.` });
        return;
    }
    const existingInvite = await Invite_1.default.findOne({
        email: email.toLowerCase(),
        used: false,
        institution: auth.user.institution,
        expiresAt: { $gt: new Date() },
    }).lean();
    if (existingInvite) {
        res.status(409).json({
            message: `An active invite for ${email} already exists. Revoke it first.`,
        });
        return;
    }
    // Resolve human-readable school/department names for the email
    // Load institution settings to get the names
    const InstitutionSettings = (await Promise.resolve().then(() => __importStar(require("../models/InstitutionSettings")))).default;
    const Institution = (await Promise.resolve().then(() => __importStar(require("../models/Institution")))).default;
    const [settingsDoc, institutionDoc] = await Promise.all([
        InstitutionSettings.findOne({ institution: auth.user.institution })
            .select("schools docMeta")
            .lean(),
        Institution.findById(auth.user.institution).select("name").lean(),
    ]);
    // Resolve names — prefer docMeta.universityName, fall back to Institution.name
    const universityName = settingsDoc?.docMeta?.universityName
        ?? institutionDoc?.name
        ?? "University";
    let resolvedSchoolName;
    let resolvedDepartmentName;
    if (schoolCode && settingsDoc?.schools) {
        const school = settingsDoc.schools.find(s => s.code === schoolCode.toUpperCase());
        resolvedSchoolName = school?.name;
        if (departmentCode && school?.departments) {
            const dept = school.departments.find(d => d.code === departmentCode.toUpperCase());
            resolvedDepartmentName = dept?.name;
        }
    }
    const finalName = name?.trim() || email
        .split("@")[0]
        .split(".")
        .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
        .join(" ");
    const token = node_crypto_1.default.randomBytes(24).toString("hex");
    const hasDepts = !!(settingsDoc?.schools?.some(s => (s.departments?.length ?? 0) > 0));
    await Invite_1.default.create({
        name: finalName,
        email: email.toLowerCase(),
        token,
        role,
        used: false,
        expiresAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
        createdBy: auth.user._id,
        institution: auth.user.institution,
        schoolCode: schoolCode?.toUpperCase() ?? null,
        departmentCode: departmentCode?.toUpperCase() ?? null,
        institutionWide: institutionWide ?? (role === "lecturer" || !hasDepts),
    });
    // Send rich invitation email with university/school/department context
    const { sendInviteEmail } = await Promise.resolve().then(() => __importStar(require("../config/email")));
    sendInviteEmail({
        to: email,
        token,
        name: finalName,
        role,
        universityName,
        schoolName: resolvedSchoolName,
        departmentName: resolvedDepartmentName,
        institutionWide: institutionWide ?? (role === "lecturer" || !hasDepts),
    }).catch((err) => {
        console.error("[Admin] Invite email failed:", err.message);
    });
    await (0, auditLogger_1.logAudit)(auth, {
        action: "invite_created",
        details: {
            email, role, schoolCode, departmentCode,
            institution: auth.user.institution.toString(),
        },
    });
    res.status(201).json({ message: `Invite sent to ${finalName}` });
}));
// POST /admin/register/:token
router.post("/register/:token", security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { token } = req.params;
    const { password } = req.body;
    if (!password || password.length < 8) {
        res.status(400).json({ message: "Password must be at least 8 characters." });
        return;
    }
    const invite = await Invite_1.default.findOne({
        token,
        used: false,
        expiresAt: { $gt: new Date() },
    }).lean();
    if (!invite) {
        res.status(400).json({ message: "Invite link is invalid or has expired." });
        return;
    }
    const existingUser = await User_1.default.findOne({ email: invite.email }).lean();
    if (existingUser) {
        res.status(409).json({ message: "An account with this email already exists." });
        return;
    }
    const hashed = await bcryptjs_1.default.hash(password, 12);
    await User_1.default.create({
        name: invite.name,
        email: invite.email,
        password: hashed,
        role: invite.role,
        status: "active",
        institution: invite.institution,
        schoolCode: invite.schoolCode ?? null,
        departmentCode: invite.departmentCode ?? null,
        institutionWide: invite.institutionWide ?? (invite.role === "lecturer"),
    });
    await Invite_1.default.updateOne({ _id: invite._id }, { used: true });
    await AuditLog_1.default.create({
        action: "invite_used",
        details: { email: invite.email, role: invite.role },
    }).catch((e) => console.error("[AuditLog]", e.message));
    res.status(201).json({ message: "Account created successfully. You can now log in." });
}));
// DELETE /admin/invites/:id
router.delete("/invites/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const invite = await Invite_1.default.findOne({
        _id: req.params.id,
        institution: auth.user.institution,
    }).lean();
    if (!invite) {
        res.status(404).json({ message: "Invite not found." });
        return;
    }
    await Invite_1.default.deleteOne({ _id: invite._id });
    await AuditLog_1.default.create({
        action: "invite_revoked",
        actor: auth.user._id,
        targetUser: invite._id,
        details: { email: invite.email, role: invite.role },
    }).catch((err) => console.error("[AuditLog]", err.message));
    res.json({ message: "Invite revoked." });
}));
// GET /admin/users
router.get("/users", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const users = await User_1.default.find({ institution: auth.user.institution })
        .select("-password")
        .lean();
    res.json(users);
}));
// PUT /admin/users/:id/role
router.put("/users/:id/role", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const { role } = req.body;
    const { id } = req.params;
    if (!["admin", "lecturer", "coordinator"].includes(role)) {
        res.status(400).json({ message: "Invalid role." });
        return;
    }
    if (auth.user._id.toString() === id) {
        res.status(403).json({ message: "You cannot change your own role." });
        return;
    }
    const user = await User_1.default.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
        res.status(404).json({ message: "User not found in your institution." });
        return;
    }
    if (user.role === "admin" && role !== "admin") {
        const adminCount = await User_1.default.countDocuments({
            institution: auth.user.institution,
            role: "admin",
        });
        if (adminCount <= 1) {
            res.status(403).json({ message: "Cannot demote the last admin." });
            return;
        }
    }
    const oldRole = user.role;
    user.role = role;
    await user.save();
    await AuditLog_1.default.create({
        action: "role_changed",
        actor: auth.user._id,
        targetUser: user._id,
        details: { from: oldRole, to: role },
    }).catch((err) => console.error("[AuditLog]", err.message));
    res.json({ message: "Role updated.", user });
}));
// PUT /admin/users/:id/status
router.put("/users/:id/status", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const { status } = req.body;
    const { id } = req.params;
    if (!["active", "suspended"].includes(status)) {
        res.status(400).json({ message: "Status must be 'active' or 'suspended'." });
        return;
    }
    const user = await User_1.default.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
        res.status(404).json({ message: "User not found in your institution." });
        return;
    }
    const oldStatus = user.status;
    user.status = status;
    await user.save();
    await AuditLog_1.default.create({
        action: "status_toggled",
        actor: auth.user._id,
        targetUser: user._id,
        details: { from: oldStatus, to: status },
    }).catch((err) => console.error("[AuditLog]", err.message));
    res.json({ message: "Status updated.", user });
}));
// DELETE /admin/users/:id
router.delete("/users/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const { id } = req.params;
    if (auth.user._id.toString() === id) {
        res.status(403).json({ message: "You cannot delete your own account." });
        return;
    }
    const user = await User_1.default.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
        res.status(404).json({ message: "User not found in your institution." });
        return;
    }
    if (user.role === "admin") {
        const adminCount = await User_1.default.countDocuments({
            institution: auth.user.institution,
            role: "admin",
        });
        if (adminCount <= 1) {
            res.status(403).json({ message: "Cannot delete the last admin." });
            return;
        }
    }
    await User_1.default.deleteOne({ _id: id });
    await AuditLog_1.default.create({
        action: "user_deleted",
        actor: auth.user._id,
        targetUser: user._id,
        details: { email: user.email, role: user.role },
    }).catch((err) => console.error("[AuditLog]", err.message));
    res.json({ message: "User deleted." });
}));
// PUT /admin/users/:id/details - Update coordinator details (school/dept)
router.put("/users/:id/details", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), security_1.sanitizeInput, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const { id } = req.params;
    const { name, schoolCode, departmentCode, institutionWide } = req.body;
    // Prevent self-modification
    if (auth.user._id.toString() === id) {
        throw {
            statusCode: 403,
            message: "You cannot modify your own account details.",
        };
    }
    const user = await User_1.default.findOne({ _id: id, institution: auth.user.institution });
    if (!user) {
        throw { statusCode: 404, message: "User not found in your institution." };
    }
    // Only coordinators can have school/department assignments
    if (user.role !== "coordinator" && (schoolCode !== undefined || departmentCode !== undefined || institutionWide !== undefined)) {
        throw {
            statusCode: 400,
            message: "School/department assignment only applies to coordinators.",
        };
    }
    // For coordinators: validate school/department exist if provided
    if (user.role === "coordinator" && !institutionWide && schoolCode && departmentCode) {
        const InstitutionSettings = (await Promise.resolve().then(() => __importStar(require("../models/InstitutionSettings")))).default;
        const settings = await InstitutionSettings.findOne({ institution: auth.user.institution })
            .select("schools")
            .lean();
        const school = settings?.schools?.find(s => s.code === schoolCode.toUpperCase());
        if (!school) {
            throw { statusCode: 400, message: `School "${schoolCode}" not found.` };
        }
        const department = school.departments?.find(d => d.code === departmentCode.toUpperCase());
        if (!department) {
            throw { statusCode: 400, message: `Department "${departmentCode}" not found in school "${schoolCode}".` };
        }
    }
    // Apply updates
    if (name)
        user.name = name;
    // Handle schoolCode - convert empty string to undefined (Mongoose will ignore undefined)
    if (schoolCode !== undefined) {
        user.schoolCode = schoolCode && schoolCode.trim() !== "" ? schoolCode.toUpperCase() : undefined;
    }
    // Handle departmentCode - convert empty string to undefined
    if (departmentCode !== undefined) {
        user.departmentCode = departmentCode && departmentCode.trim() !== "" ? departmentCode.toUpperCase() : undefined;
    }
    if (institutionWide !== undefined)
        user.institutionWide = institutionWide;
    // If institution-wide is true, clear school/department assignments
    if (institutionWide === true) {
        user.schoolCode = undefined;
        user.departmentCode = undefined;
    }
    await user.save();
    await (0, auditLogger_1.logAudit)(auth, {
        action: "user_details_updated",
        targetUser: user._id,
        details: {
            name: name || user.name,
            schoolCode: user.schoolCode,
            departmentCode: user.departmentCode,
            institutionWide: user.institutionWide
        },
    });
    const updatedUser = await User_1.default.findById(id).select("-password").lean();
    res.json({ message: "User details updated successfully", user: updatedUser });
}));
// GET /admin/lecturers
router.get("/lecturers", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const auth = req;
    const lecturers = await User_1.default.find({
        institution: auth.user.institution,
        role: "lecturer",
    })
        .select("-password")
        .lean();
    res.json(lecturers);
}));
exports.default = router;
