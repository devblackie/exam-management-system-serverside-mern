// src/routes/institutionSettings.ts
import express, { Response } from "express";
import InstitutionSettings from "../models/InstitutionSettings";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";

const router = express.Router();

// GET: Fetch current settings
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res:Response) => {
    const settings = await InstitutionSettings.findOne({
      institution: req.user.institution,
    });

    if (!settings) {
      return res.status(404).json({ message: "Settings not configured yet" });
    }

    res.json(settings);
  })
);

// POST: Save / Update settings (coordinator only)
router.post(
  "/",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res:Response) => {
    const data = req.body;

    // Before updating, warn if there are existing marks for the current year
    const existingMarks = await Mark.countDocuments({ institution: req.user.institution });
    // -----
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

    res.json({
      message: "Institution settings saved successfully",
      settings: updated,
    });
  })
);

export default router;