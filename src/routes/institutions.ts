
// serverside/src/routes/institutions.ts — COMPLETE
import { Router, Request, Response } from "express";
import mongoose  from "mongoose";
import Institution from "../models/Institution";
import { requireAuth, requireRole, AuthenticatedRequest } from "../middleware/auth";
import { asyncHandler }  from "../middleware/asyncHandler";
import { logAudit }      from "../lib/auditLogger";
import { invalidateCache } from "../utils/cache";

const router = Router();

// ── GET /institutions/public — for the secret-register page dropdown ──────────
// No auth required — only returns name, code, _id
router.get(
  "/public",
  asyncHandler(async (_req: Request, res: Response): Promise<void> => {
    const institutions = await Institution.find({ isActive: true })
      .select("name code")
      .sort({ name: 1 })
      .lean();
    res.json(institutions);
  }),
);

// ── GET /institutions/mine — returns the institution the admin belongs to ──────
router.get(
  "/mine",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const institution = await Institution.findById(req.user.institution).lean();
    if (!institution) {
      res.status(404).json({ message: "Institution not found" });
      return;
    }
    res.json(institution);
  }),
);

// ── PATCH /institutions/mine — admin updates their institution's profile ───────
// This is how "Demo University" becomes "University of Nairobi"
router.patch(
  "/mine",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { name, code, abbreviation, address, website, email, phone, city, country } =
      req.body as Partial<{
        name:         string;
        code:         string;
        abbreviation: string;
        address:      string;
        website:      string;
        email:        string;
        phone:        string;
        city:         string;
        country:      string;
      }>;

    // Prevent renaming to a name already taken by another institution
    if (name) {
      const clash = await Institution.findOne({
        name,
        _id: { $ne: req.user.institution },
      }).lean();
      if (clash) {
        res.status(409).json({ message: `An institution named "${name}" already exists.` });
        return;
      }
    }

    if (code) {
      const clash = await Institution.findOne({
        code: code.toUpperCase(),
        _id: { $ne: req.user.institution },
      }).lean();
      if (clash) {
        res.status(409).json({ message: `Institution code "${code}" is already in use.` });
        return;
      }
    }

    const updated = await Institution.findByIdAndUpdate(
      req.user.institution,
      {
        $set: {
          ...(name         ? { name }                        : {}),
          ...(code         ? { code: code.toUpperCase() }    : {}),
          ...(abbreviation ? { abbreviation }                : {}),
          ...(address      ? { address }                     : {}),
          ...(website      ? { website }                     : {}),
          ...(email        ? { email }                       : {}),
          ...(phone        ? { phone }                       : {}),
          ...(city         ? { city }                        : {}),
          ...(country      ? { country }                     : {}),
        },
      },
      { new: true, runValidators: true },
    );

    // Bust settings cache since institution name may be in docMeta
    invalidateCache(`settings:${req.user.institution}`);

    await logAudit(req, {
      action:  "institution_profile_updated",
      details: { institutionId: req.user.institution?.toString(), changes: req.body },
    });

    res.json({ message: "Institution updated", institution: updated });
  }),
);

// ── GET /institutions — list all (admin only, platform-level) ─────────────────
// Only needed for platform-level super-admin; not exposed to regular admins
router.get(
  "/",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    // Regular admins only see their own institution
    const institution = await Institution.findById(req.user.institution).lean();
    res.json(institution ? [institution] : []);
  }),
);

export default router;