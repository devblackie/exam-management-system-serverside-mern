// serverside/src/routes/promote.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { bulkPromoteClass } from "../services/statusEngine";

const router = Router();

router.post(
  "/bulk-promote",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    if (!programId || !yearToPromote || !academicYearName) {
      return res.status(400).json({ error: "Missing required promotion parameters" });
    }

    const results = await bulkPromoteClass(programId, yearToPromote, academicYearName);

    res.json({
      success: true,
      message: `Process completed: ${results.promoted} promoted, ${results.failed} failed.`,
      data: results
    });
  })
);

export default router;