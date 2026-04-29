
// src/routes/institutions.ts
import { Request, Response, Router } from "express";
import Institution from "../models/Institution";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";

const router = Router();

// PUBLIC: Active institutions list (unauthenticated — login/register pages)
router.get(
  "/public",
  asyncHandler(async (req: Request, res: Response) => {
    const institutions = await Institution.find({ isActive: true })
      .select("name code _id")
      .lean();

    // Fire-and-forget — no actor on unauthenticated route
    logAudit(req, {
      action: "institutions_public_listed",
      details: {
        count: institutions.length,
        ip: req.ip,
        userAgent: req.headers["user-agent"] ?? "unknown",
      },
    }).catch(() => {});

    const response = institutions.map((inst) => ({
      _id: inst._id.toString(),
      name: inst.name,
      code: inst.code,
    }));

    res.json(response);
  })
);

// ADMIN: Full institution list
router.get(
  "/",
  asyncHandler(async (req: Request, res: Response) => {
    const institutions = await Institution.find().lean();

    logAudit(req, {
      action: "institutions_admin_listed",
      details: {
        count: institutions.length,
        ip: req.ip,
        userAgent: req.headers["user-agent"] ?? "unknown",
      },
    }).catch(() => {});

    res.json(institutions);
  })
);

export default router;