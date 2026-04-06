// // serverside/src/routes/academicYears.ts
// import { Response, Router } from "express";
// import AcademicYear from "../models/AcademicYear";
// import { requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import type { AuthenticatedRequest } from "../middleware/auth";

// const router = Router();

// // 1. CREATE
// router.post("/", requireAuth, requireRole("admin", "coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { year, startDate, endDate } = req.body;
//     const start = new Date(startDate);
//     const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
//     const intake = monthNames[start.getMonth()];

//     const academicYear = await AcademicYear.create({
//       institution: req.user.institution,
//       year,
//       intakes: [intake],
//       startDate: start,
//       endDate: new Date(endDate),
//       isCurrent: false,
//     });

    
//     res.status(201).json(academicYear);
//   })
// );

// // 2. UPDATE (PATCH) - Used for Session Toggles and Date Extensions
// router.patch("/:id", requireAuth, requireRole("admin", "coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { isCurrent } = req.body;
//     if (isCurrent) {
//       await AcademicYear.updateMany({ institution: req.user.institution }, { isCurrent: false });
//     }
//     const updated = await AcademicYear.findByIdAndUpdate(req.params.id, { $set: req.body }, { new: true });
//     res.json(updated);
//   })
// );

// // 3. GET ALL
// router.get("/", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const years = await AcademicYear.find({ institution: req.user.institution }).sort({ startDate: -1 }).lean();
//     res.json(years);
//   })
// );

// // 4. DELETE
// router.delete("/:id", requireAuth, requireRole("admin", "coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const year = await AcademicYear.findOne({ _id: req.params.id, institution: req.user.institution });
//     if (!year) return res.status(404).json({ message: "Year not found" });
//     if (year.isCurrent) return res.status(400).json({ message: "Cannot delete the active academic year" });

//     await AcademicYear.findByIdAndDelete(req.params.id);
//     res.json({ message: "Deleted successfully" });
//   })
// );

// export default router;

// serverside/src/routes/academicYears.ts
import { Response, Router } from "express";
import AcademicYear from "../models/AcademicYear";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

// 1. CREATE
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { year, startDate, endDate } = req.body;
    const start = new Date(startDate);
    const monthNames = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
    const intake = monthNames[start.getMonth()];

    const academicYear = await AcademicYear.create({
      institution: req.user.institution,
      year,
      intakes: [intake],
      startDate: start,
      endDate: new Date(endDate),
      isCurrent: false,
    });

    await logAudit(req, {
      action: "academic_year_created",
      actor: req.user._id,
      details: {
        year,
        intake,
        startDate: start.toISOString(),
        endDate: new Date(endDate).toISOString(),
        institutionId: req.user.institution?.toString(),
      },
    });

    res.status(201).json(academicYear);
  })
);

// 2. UPDATE (PATCH)
router.patch(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const before = await AcademicYear.findOne({
      _id: req.params.id,
      institution: req.user.institution,
    }).lean();

    if (!before) {
      await logAudit(req, {
        action: "academic_year_update_failed",
        actor: req.user._id,
        details: {
          academicYearId: req.params.id,
          reason: "Not found or institution mismatch",
          attemptedChanges: req.body,
        },
      });
      return res.status(404).json({ message: "Academic year not found" });
    }

    const { isCurrent } = req.body;
    if (isCurrent) {
      await AcademicYear.updateMany(
        { institution: req.user.institution },
        { isCurrent: false }
      );
    }

    const updated = await AcademicYear.findByIdAndUpdate(
      req.params.id,
      { $set: req.body },
      { new: true }
    );

    await logAudit(req, {
      action: "academic_year_updated",
      actor: req.user._id,
      details: {
        academicYearId: req.params.id,
        year: before.year,
        institutionId: req.user.institution?.toString(),
        before: {
          isCurrent: before.isCurrent,
          session: (before as any).session,
          startDate: before.startDate,
          endDate: before.endDate,
        },
        after: req.body,
        demotedOthers: !!isCurrent,
      },
    });

    res.json(updated);
  })
);

// 3. GET ALL
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const years = await AcademicYear.find({ institution: req.user.institution })
      .sort({ startDate: -1 })
      .lean();

    await logAudit(req, {
      action: "academic_years_listed",
      actor: req.user._id,
      details: {
        count: years.length,
        institutionId: req.user.institution?.toString(),
      },
    });

    res.json(years);
  })
);

// 4. DELETE
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const year = await AcademicYear.findOne({
      _id: req.params.id,
      institution: req.user.institution,
    });

    if (!year) {
      await logAudit(req, {
        action: "academic_year_delete_failed",
        actor: req.user._id,
        details: {
          academicYearId: req.params.id,
          reason: "Not found or institution mismatch",
          institutionId: req.user.institution?.toString(),
        },
      });
      return res.status(404).json({ message: "Year not found" });
    }

    if (year.isCurrent) {
      await logAudit(req, {
        action: "academic_year_delete_failed",
        actor: req.user._id,
        details: {
          academicYearId: req.params.id,
          year: year.year,
          reason: "Attempted deletion of active academic year",
          institutionId: req.user.institution?.toString(),
        },
      });
      return res.status(400).json({ message: "Cannot delete the active academic year" });
    }

    await AcademicYear.findByIdAndDelete(req.params.id);

    await logAudit(req, {
      action: "academic_year_deleted",
      actor: req.user._id,
      details: {
        academicYearId: req.params.id,
        year: year.year,
        startDate: year.startDate,
        endDate: year.endDate,
        institutionId: req.user.institution?.toString(),
      },
    });

    res.json({ message: "Deleted successfully" });
  })
);

export default router;