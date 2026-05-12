
// // src/routes/institutions.ts
// import { Request, Response, Router } from "express";
// import Institution from "../models/Institution";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { logAudit } from "../lib/auditLogger";

// const router = Router();

// // PUBLIC: Active institutions list (unauthenticated — login/register pages)
// router.get(
//   "/public",
//   asyncHandler(async (req: Request, res: Response) => {
//     const institutions = await Institution.find({ isActive: true })
//       .select("name code _id")
//       .lean();

//     // Fire-and-forget — no actor on unauthenticated route
//     logAudit(req, {
//       action: "institutions_public_listed",
//       details: {
//         count: institutions.length,
//         ip: req.ip,
//         userAgent: req.headers["user-agent"] ?? "unknown",
//       },
//     }).catch(() => {});

//     const response = institutions.map((inst) => ({
//       _id: inst._id.toString(),
//       name: inst.name,
//       code: inst.code,
//     }));

//     res.json(response);
//   })
// );

// // ADMIN: Full institution list
// router.get(
//   "/",
//   asyncHandler(async (req: Request, res: Response) => {
//     const institutions = await Institution.find().lean();

//     logAudit(req, {
//       action: "institutions_admin_listed",
//       details: {
//         count: institutions.length,
//         ip: req.ip,
//         userAgent: req.headers["user-agent"] ?? "unknown",
//       },
//     }).catch(() => {});

//     res.json(institutions);
//   })
// );

// export default router;








// // serverside/src/routes/institutions.ts
// import { Router, Response } from "express";
// import mongoose from "mongoose";
// import Institution from "../models/Institution";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { logAudit } from "../lib/auditLogger";

// const router = Router();

// // GET /institutions - List all institutions (admin only)
// router.get(
//   "/",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const institutions = await Institution.find({}).select("-__v").lean();
//     res.json(institutions);
//   })
// );

// // GET /institutions/public - Public list (no auth required)
// router.get(
//   "/public",
//   asyncHandler(async (_req: Request, res: Response) => {
//     const institutions = await Institution.find({ isActive: true })
//       .select("_id name code")
//       .lean();
//     res.json(institutions);
//   })
// );

// // POST /institutions - Create new institution (super admin only)
// router.post(
//   "/",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { name, code, isActive } = req.body;

//     if (!name || !code) {
//       res.status(400).json({ message: "Name and code are required." });
//       return;
//     }

//     // Check for duplicate code
//     const existing = await Institution.findOne({ code: code.toUpperCase() });
//     if (existing) {
//       res.status(409).json({ message: "Institution with this code already exists." });
//       return;
//     }

//     const institution = await Institution.create({
//       name,
//       code: code.toUpperCase(),
//       isActive: isActive !== false,
//     });

//     await logAudit(req, {
//       action: "institution_created",
//       actor: req.user._id,
//       details: { name, code, id: institution._id },
//     });

//     res.status(201).json(institution);
//   })
// );

// // PUT /institutions/:id - Update institution (super admin only)
// router.put(
//   "/:id",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { id } = req.params;
//     const { name, code, isActive } = req.body;

//     if (!mongoose.isValidObjectId(id)) {
//       res.status(400).json({ message: "Invalid institution ID." });
//       return;
//     }

//     const institution = await Institution.findById(id);
//     if (!institution) {
//       res.status(404).json({ message: "Institution not found." });
//       return;
//     }

//     if (code && code !== institution.code) {
//       const existing = await Institution.findOne({ code: code.toUpperCase(), _id: { $ne: id } });
//       if (existing) {
//         res.status(409).json({ message: "Institution with this code already exists." });
//         return;
//       }
//       institution.code = code.toUpperCase();
//     }

//     if (name) institution.name = name;
//     if (isActive !== undefined) institution.isActive = isActive;

//     await institution.save();

//     await logAudit(req, {
//       action: "institution_updated",
//       actor: req.user._id,
//       details: { id, name, code, isActive },
//     });

//     res.json(institution);
//   })
// );

// // DELETE /institutions/:id - Soft delete institution (super admin only)
// router.delete(
//   "/:id",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { id } = req.params;

//     if (!mongoose.isValidObjectId(id)) {
//       res.status(400).json({ message: "Invalid institution ID." });
//       return;
//     }

//     const institution = await Institution.findById(id);
//     if (!institution) {
//       res.status(404).json({ message: "Institution not found." });
//       return;
//     }

//     // Check if there are any users/admins left
//     // You may want to add logic to prevent deletion if there are active users

//     institution.isActive = false;
//     await institution.save();

//     await logAudit(req, {
//       action: "institution_deactivated",
//       actor: req.user._id,
//       details: { id, name: institution.name },
//     });

//     res.json({ message: "Institution deactivated successfully." });
//   })
// );

// export default router;







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