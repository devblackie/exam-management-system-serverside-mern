// serverside/src/routes/academicYears.ts
import { Response, Router } from "express";
import AcademicYear from "../models/AcademicYear";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

// CREATE ACADEMIC YEAR
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res:Response) => {
    const { year, startDate, endDate } = req.body;

    if (!year || !startDate || !endDate) {
      return res.status(400).json({ message: "Year, startDate, and endDate are required" });
    }

    // Prevent duplicates
    const exists = await AcademicYear.findOne({
      institution: req.user.institution,
      year,
    });

    if (exists) {
      return res.status(400).json({ message: `Academic year ${year} already exists` });
    }

    const academicYear = await AcademicYear.create({
      institution: req.user.institution,
      year,
      startDate: new Date(startDate),
      endDate: new Date(endDate),
      isCurrent: false, // You can add logic later to set only one as current
    });

    res.status(201).json(academicYear);
  })
);

// GET ALL ACADEMIC YEARS
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res:Response) => {
    const years = await AcademicYear.find({
      institution: req.user.institution,
    })
      .sort({ startDate: -1 })
      .lean();

    res.json(years);
  })
);

export default router;