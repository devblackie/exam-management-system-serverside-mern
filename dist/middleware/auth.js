"use strict";
// // serverside/src/middleware/auth.ts
// //
// // UPDATED: Every user — including admins — must be linked to an institution.
// // The previous version exempted admins from the institution check.
// // Per the project requirement: "every user (admins included) should be linked
// // to an institution."
// //
// // This means:
// //   - Admin secret-register MUST supply an institutionId
// //   - setAuthCookie MUST include institution in the JWT for all roles
// //   - requireAuth blocks ANY user missing institution (no role exception)
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.requireAuth = requireAuth;
exports.requireRole = requireRole;
exports.getScopedProgramIds = getScopedProgramIds;
const jwt_1 = require("../lib/jwt");
const User_1 = __importDefault(require("../models/User"));
const Program_1 = __importDefault(require("../models/Program"));
const mongoose_1 = __importDefault(require("mongoose"));
const auditLogger_1 = require("../lib/auditLogger");
// ── requireAuth ───────────────────────────────────────────────────────────────
async function requireAuth(req, res, next) {
    const token = req.cookies?.token;
    if (!token) {
        await (0, auditLogger_1.logAudit)(req, { action: "unauthenticated_access", details: { path: req.originalUrl } });
        res.status(401).json({ message: "Not authenticated" });
        return;
    }
    try {
        const payload = (0, jwt_1.verifyToken)(token);
        if (!payload?.id) {
            res.status(401).json({ message: "Invalid token" });
            return;
        }
        // .lean() returns a plain object — select only the fields ScopedUser needs
        // so that TypeScript has a concrete shape to work with.
        const userDoc = await User_1.default.findById(payload.id)
            .select("name email role status institution tokenVersion schoolCode departmentCode institutionWide")
            .lean();
        if (!userDoc) {
            res.clearCookie("token");
            res.status(401).json({ message: "User not found" });
            return;
        }
        if (userDoc.status === "suspended") {
            res.clearCookie("token");
            res.status(403).json({ message: "Account suspended. Contact your administrator." });
            return;
        }
        if (typeof payload.version === "number" &&
            payload.version !== (userDoc.tokenVersion ?? 0)) {
            res.clearCookie("token");
            res.status(401).json({ message: "Session expired. Please log in again." });
            return;
        }
        if (!payload.institution) {
            await (0, auditLogger_1.logAudit)(req, {
                action: "missing_institution_in_jwt",
                details: { userId: payload.id, role: userDoc.role },
            });
            res.status(403).json({
                message: "Account not linked to an institution. Contact a system administrator.",
            });
            return;
        }
        // Assemble the ScopedUser — all fields are now correctly typed
        const scopedUser = {
            _id: userDoc._id,
            name: userDoc.name,
            email: userDoc.email,
            role: userDoc.role,
            status: userDoc.status,
            institution: new mongoose_1.default.Types.ObjectId(payload.institution),
            schoolCode: userDoc.schoolCode ?? null,
            departmentCode: userDoc.departmentCode ?? null,
            institutionWide: userDoc.institutionWide ?? false,
            tokenVersion: userDoc.tokenVersion ?? 0,
        };
        req.user = scopedUser;
        next();
    }
    catch (err) {
        const message = err instanceof Error ? err.message : "Unknown error";
        await (0, auditLogger_1.logAudit)(req, {
            action: "token_verification_failed",
            details: { error: message, path: req.originalUrl },
        });
        res.clearCookie("token");
        res.status(401).json({ message: "Session expired. Please log in again." });
    }
}
// ── requireRole ───────────────────────────────────────────────────────────────
function requireRole(...roles) {
    return (req, res, next) => {
        const user = req.user;
        if (!user) {
            res.status(401).json({ message: "Not authenticated" });
            return;
        }
        // Admins bypass all role restrictions within their institution
        if (user.role === "admin") {
            next();
            return;
        }
        if (!roles.includes(user.role)) {
            res.status(403).json({ message: "Insufficient permissions for this action." });
            return;
        }
        next();
    };
}
// ── getScopedProgramIds ───────────────────────────────────────────────────────
// Returns the ObjectId strings of programs visible to this user.
//
// FIX: Mongoose .lean() returns FlattenMaps<T> where _id is FlattenMaps<unknown>.
// This is NOT assignable to ObjectId directly. The fix is to provide an explicit
// generic to .lean<MinimalShape>() so TypeScript knows the exact shape we get
// back, instead of letting it infer the problematic FlattenMaps<IProgram> type.
async function getScopedProgramIds(req) {
    const filter = { institution: req.user.institution };
    if (!req.user.institutionWide) {
        if (req.user.departmentCode)
            filter.departmentCode = req.user.departmentCode;
        if (req.user.schoolCode)
            filter.schoolCode = req.user.schoolCode;
    }
    // Provide the explicit shape to .lean() — avoids FlattenMaps<unknown> on _id
    const programs = await Program_1.default.find(filter)
        .select("_id")
        .lean();
    return programs.map(p => p._id.toString());
}
