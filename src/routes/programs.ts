// src/routes/programs.ts
import { Response, Router } from "express";
import Program from "../models/Program";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

// CREATE
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { name, code, description, durationYears } = req.body;

    const exists = await Program.findOne({
      code: code?.toUpperCase(),
      institution: req.user.institution,
    });

    if (exists) {
      await logAudit(req, {
        action: "program_create_failed",
        actor: req.user._id,
        details: {
          reason: "Duplicate program code",
          attemptedCode: code?.toUpperCase(),
          attemptedName: name,
          institutionId: req.user.institution?.toString(),
        },
      });
      return res.status(400).json({
        message: "Program code already exists in your institution",
      });
    }

    const program = await Program.create({
      name,
      code: code?.toUpperCase(),
      description,
      durationYears,
      institution: req.user.institution,
    });

    await logAudit(req, {
      action: "program_created",
      actor: req.user._id,
      details: {
        name: program.name,
        code: program.code,
        durationYears: program.durationYears,
        institutionId: req.user.institution?.toString(),
      },
    });

    res.status(201).json(program);
  })
);

// GET ALL
router.get(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const programs = await Program.find({
      institution: req.user.institution,
    }).sort({ code: 1 });

    await logAudit(req, {
      action: "programs_listed",
      actor: req.user._id,
      details: {
        count: programs.length,
        institutionId: req.user.institution?.toString(),
      },
    });

    res.json(programs);
  })
);

// UPDATE
router.put(
  "/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const before = await Program.findOne({
      _id: req.params.id,
      institution: req.user.institution,
    }).lean();

    if (!before) {
      await logAudit(req, {
        action: "program_update_failed",
        actor: req.user._id,
        details: {
          programId: req.params.id,
          reason: "Not found or institution mismatch",
          attemptedChanges: req.body,
          institutionId: req.user.institution?.toString(),
        },
      });
      return res.status(404).json({ message: "Program not found" });
    }

    const program = await Program.findOneAndUpdate(
      { _id: req.params.id, institution: req.user.institution },
      req.body,
      { new: true, runValidators: true }
    );

    await logAudit(req, {
      action: "program_updated",
      actor: req.user._id,
      details: {
        programId: req.params.id,
        institutionId: req.user.institution?.toString(),
        before: {
          name: before.name,
          code: before.code,
          durationYears: before.durationYears,
        },
        after: req.body,
      },
    });

    res.json(program);
  })
);

// DELETE
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
      await logAudit(req, {
        action: "program_delete_failed",
        actor: req.user._id,
        details: {
          programId: req.params.id,
          reason: "Not found or institution mismatch",
          institutionId: req.user.institution?.toString(),
        },
      });
      return res.status(404).json({ message: "Program not found" });
    }

    await logAudit(req, {
      action: "program_deleted",
      actor: req.user._id,
      details: {
        programId: req.params.id,
        name: program.name,
        code: program.code,
        durationYears: program.durationYears,
        institutionId: req.user.institution?.toString(),
      },
    });

    res.json({ message: "Program deleted successfully" });
  })
);

export default router;
