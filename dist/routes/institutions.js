"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/routes/institutions.ts — COMPLETE
const express_1 = require("express");
const Institution_1 = __importDefault(require("../models/Institution"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const auditLogger_1 = require("../lib/auditLogger");
const cache_1 = require("../utils/cache");
const router = (0, express_1.Router)();
// ── GET /institutions/public — for the secret-register page dropdown ──────────
// No auth required — only returns name, code, _id
router.get("/public", (0, asyncHandler_1.asyncHandler)(async (_req, res) => {
    const institutions = await Institution_1.default.find({ isActive: true })
        .select("name code")
        .sort({ name: 1 })
        .lean();
    res.json(institutions);
}));
// ── GET /institutions/mine — returns the institution the admin belongs to ──────
router.get("/mine", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institution = await Institution_1.default.findById(req.user.institution).lean();
    if (!institution) {
        res.status(404).json({ message: "Institution not found" });
        return;
    }
    res.json(institution);
}));
// ── PATCH /institutions/mine — admin updates their institution's profile ───────
// This is how "Demo University" becomes "University of Nairobi"
router.patch("/mine", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { name, code, abbreviation, address, website, email, phone, city, country } = req.body;
    // Prevent renaming to a name already taken by another institution
    if (name) {
        const clash = await Institution_1.default.findOne({
            name,
            _id: { $ne: req.user.institution },
        }).lean();
        if (clash) {
            res.status(409).json({ message: `An institution named "${name}" already exists.` });
            return;
        }
    }
    if (code) {
        const clash = await Institution_1.default.findOne({
            code: code.toUpperCase(),
            _id: { $ne: req.user.institution },
        }).lean();
        if (clash) {
            res.status(409).json({ message: `Institution code "${code}" is already in use.` });
            return;
        }
    }
    const updated = await Institution_1.default.findByIdAndUpdate(req.user.institution, {
        $set: {
            ...(name ? { name } : {}),
            ...(code ? { code: code.toUpperCase() } : {}),
            ...(abbreviation ? { abbreviation } : {}),
            ...(address ? { address } : {}),
            ...(website ? { website } : {}),
            ...(email ? { email } : {}),
            ...(phone ? { phone } : {}),
            ...(city ? { city } : {}),
            ...(country ? { country } : {}),
        },
    }, { new: true, runValidators: true });
    // Bust settings cache since institution name may be in docMeta
    (0, cache_1.invalidateCache)(`settings:${req.user.institution}`);
    await (0, auditLogger_1.logAudit)(req, {
        action: "institution_profile_updated",
        details: { institutionId: req.user.institution?.toString(), changes: req.body },
    });
    res.json({ message: "Institution updated", institution: updated });
}));
// ── GET /institutions — list all (admin only, platform-level) ─────────────────
// Only needed for platform-level super-admin; not exposed to regular admins
router.get("/", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    // Regular admins only see their own institution
    const institution = await Institution_1.default.findById(req.user.institution).lean();
    res.json(institution ? [institution] : []);
}));
exports.default = router;
