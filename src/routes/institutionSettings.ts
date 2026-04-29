
// src/routes/institutionSettings.ts
import express, { Response } from "express";
import InstitutionSettings from "../models/InstitutionSettings";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";
import Mark from "../models/Mark";
import { cached, invalidateCache } from "../utils/cache";

const router = express.Router();

// GET: Fetch current settings
router.get("/", requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    // const settings = await InstitutionSettings.findOne({institution: req.user.institution });

    const institutionId = req.user.institution;
    const settings = await cached(`settings:${institutionId}`, () => 
      InstitutionSettings.findOne({ institution: institutionId }).lean()
    );

    if (!settings) {
      await logAudit(req, { action: "institution_settings_fetch_failed", actor: req.user._id, details: { reason: "Settings not yet configured", institutionId: req.user.institution?.toString()}});
      return res.status(404).json({ message: "Settings not configured yet" });
    }

    await logAudit(req, { action: "institution_settings_viewed", actor: req.user._id, details: { institutionId: req.user.institution?.toString(), passMark: settings.passMark, gradingScaleCount: settings.gradingScale?.length ?? 0 }});
    res.json(settings);
  })
);

// POST: Save / Update settings
router.post(
  "/",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const data = req.body;

    const previous = await InstitutionSettings.findOne({
      institution: req.user.institution,
    }).lean();

    const existingMarksCount = await Mark.countDocuments({
      institution: req.user.institution,
    });

    const updated = await InstitutionSettings.findOneAndUpdate(
      { institution: req.user.institution },
      data,
      {
        new: true,
        upsert: true,
        setDefaultsOnInsert: true,
        runValidators: true,
      }
    );

    await logAudit(req, {
      action: previous
        ? "institution_settings_updated"
        : "institution_settings_created",
      actor: req.user._id,
      details: {
        settingsId: updated?._id?.toString(),
        institutionId: req.user.institution?.toString(),
        existingMarksAtTimeOfChange: existingMarksCount,
        previous: previous
          ? {
              passMark: previous.passMark,
              gradingScaleCount: previous.gradingScale?.length ?? 0,
            }
          : null,
        updated: {
          passMark: data.passMark,
          gradingScaleCount: data.gradingScale?.length ?? 0,
        },
        fullPayload: data,
      },
    });
    invalidateCache(`settings:${req.user.institution}`);
    res.json({
      message: "Institution settings saved successfully",
      settings: updated,
    });
  })
);

export default router;