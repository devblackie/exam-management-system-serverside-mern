// src/routes/programs.ts
import { Response, Router } from "express";
import Program from "../models/Program"; // ← Fixed: Capital P, correct path
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

/**
 * CREATE Program
 */
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { name, code, description, durationYears } = req.body;

    // Ensure code is unique per institution
    const exists = await Program.findOne({
      code: code?.toUpperCase(),
      institution: req.user.institution,
    });

    if (exists) {
      return res.status(400).json({
        message: "Program code already exists in your institution",
      });
    }

    const program = await Program.create({
      name,
      code: code?.toUpperCase(),
      description,
      durationYears,
      institution: req.user.institution, // ← Critical: multi-institution
    });

    res.status(201).json(program);
  })
);

/**
 * GET All Programs (for current institution)
 */
router.get(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const programs = await Program.find({
      institution: req.user.institution,
    }).sort({ code: 1 });

    res.json(programs);
  })
);

/**
 * UPDATE Program
 */
router.put(
  "/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const program = await Program.findOneAndUpdate(
      {
        _id: req.params.id,
        institution: req.user.institution, // ← Security: can't edit other inst
      },
      req.body,
      { new: true, runValidators: true }
    );

    if (!program) {
      return res.status(404).json({ message: "Program not found" });
    }

    res.json(program);
  })
);

/**
 * DELETE Program
 */
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const program = await Program.findOneAndDelete({
      _id: req.params.id,
      institution: req.user.institution,
    });

    if (!program) {
      return res.status(404).json({ message: "Program not found" });
    }

    res.json({ message: "Program deleted successfully" });
  })
);

export default router;