// // src/routes/maintenance.ts
// import express, { Response } from "express";
// import mongoose from "mongoose";
// import Student from "../models/Student";
// import FinalGrade from "../models/FinalGrade";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import Mark from "../models/Mark";
// import Unit from "../models/Unit";
// import AcademicYear from "../models/AcademicYear";
// import ProgramUnit from "../models/ProgramUnit";
// import { computeFinalGrade } from "../services/gradeCalculator";

// const router = express.Router();

// // 1. BULK SOFT DELETE
// router.post(
//   "/bulk-cleanup",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { unitCode, programId, academicYear } = req.body;
//     // console.log(`[CLEANUP] Initiated by ${req.user.email}. Params:`, req.body);

//     if (!academicYear) return res.status(400).json({ error: "Academic Year is required" });
    
//     // 1. Resolve IDs
//     const yearDoc = await AcademicYear.findOne({ year: academicYear });
//     const unitDoc = unitCode ? await Unit.findOne({ code: unitCode }) : null;

//     if (!yearDoc) return res.status(404).json({ error: "Academic Year not found" });
    

//     // 2. Build Query
//     let query: any = {
//       academicYear: yearDoc._id,      
//     };

//     if (programId) {
//       const studentIds = await Student.find({ program: programId }).distinct( "_id" );
//       query.student = { $in: studentIds };
//     }

//     if (unitDoc) {
//       const pUnits = await ProgramUnit.find({ unit: unitDoc._id }).distinct("_id" );
//       query.programUnit = { $in: pUnits };
//     }

//     const marksToTrash = await Mark.find({ ...query, deletedAt: null });  
//     const updateResult = await Mark.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() }});
//     const gradeResult = await FinalGrade.deleteMany(query);

//     res.json({
//       count: updateResult.modifiedCount,
//       message: `Successfully moved ${updateResult.modifiedCount} marks to trash and removed ${gradeResult.deletedCount} grades.`,
//     });
//   }),
// );

// // 2. GET TRASHED MARKS
// router.get(
//   "/trash-bin",
//   requireAuth,
//   requireRole("coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     // PASS THE CRITERIA DIRECTLY TO FIND()
//     const trashed = await Mark.find({ deletedAt: { $ne: null } }) 
//       .populate("student", "regNo name")
//       .populate({ path: "programUnit", populate: { path: "unit", select: "code" }})
//       .populate("academicYear", "year")
//       .sort({ deletedAt: -1 })
//       .limit(100);

//     // console.log("Trashed marks found:", trashed.length);
//     res.json(trashed);
//   }),
// );

// // 3. RESTORE OR PERMANENT DELETE
// router.post(
//   "/trash-action",
//   requireAuth,
//   requireRole("coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     let { markIds, action } = req.body;                

//     // Ensure markIds is an array
//     if (!Array.isArray(markIds)) markIds = [markIds];
    
    
//     // console.log(`[MAINTENANCE] Action: ${action}, IDs:`, markIds);
//     if (action === "restore") {
//       // console.log(`[MAINTENANCE] Restoring marks: ${markIds}`);
//       // 1. Restore the marks
//       const marks = await Mark.find({ _id: { $in: markIds }, deletedAt: { $ne: null }});
//       // console.log(`[MAINTENANCE] Found ${marks.length} marks to restore.`);


//       await Mark.updateMany({ _id: { $in: markIds } }, { $set: { deletedAt: null }});
//       // console.log(`[MAINTENANCE] Mark documents updated to null deletedAt.`);

//       // 2. RECALCULATE GRADES for those marks
//       for (const mark of marks) await computeFinalGrade({ markId: mark._id as any})
     
//       return res.json({ message: "Marks restored and grades recalculated" });
//     }

//     if (action === "purge") {
//       await Mark.deleteMany({ _id: { $in: markIds } });
//       return res.json({ message: "Marks permanently deleted from database" });
//     }

//     res.status(400).json({ error: "Invalid action" });
//   }),
// );



// export default router;








// src/routes/maintenance.ts
// Updated to include MarkDirect in all operations:
//   - bulk-cleanup: soft-deletes both Mark and MarkDirect records
//   - trash-bin: returns trashed records from both collections, tagged by source
//   - trash-action restore: restores from the correct collection and recalculates FinalGrade
//   - trash-action purge: permanently deletes from the correct collection

import express, { Response } from "express";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";

const router = express.Router();

// ─── 1. BULK SOFT DELETE ──────────────────────────────────────────────────────
// Soft-deletes both Mark (detailed) and MarkDirect records matching the criteria,
// then removes the corresponding FinalGrade documents.
router.post(
  "/bulk-cleanup",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { unitCode, programId, academicYear } = req.body;

    if (!academicYear) return res.status(400).json({ error: "Academic Year is required" });

    const yearDoc = await AcademicYear.findOne({ year: academicYear });
    const unitDoc = unitCode ? await Unit.findOne({ code: unitCode }) : null;

    if (!yearDoc) return res.status(404).json({ error: "Academic Year not found" });

    // Build shared query filter
    const query: any = { academicYear: yearDoc._id };

    if (programId) {
      const studentIds = await Student.find({ program: programId }).distinct("_id");
      query.student = { $in: studentIds };
    }

    if (unitDoc) {
      const pUnits = await ProgramUnit.find({ unit: unitDoc._id }).distinct("_id");
      query.programUnit = { $in: pUnits };
    }

    const [detailedResult, directResult, gradeResult] = await Promise.all([
      Mark.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() } }),
      MarkDirect.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() } }),
      FinalGrade.deleteMany(query),
    ]);

    const totalTrashed = detailedResult.modifiedCount + directResult.modifiedCount;

    console.log(`[bulk-cleanup] detailed=${detailedResult.modifiedCount}, direct=${directResult.modifiedCount}, grades=${gradeResult.deletedCount}`);

    res.json({
      count:   totalTrashed,
      message: `Moved ${detailedResult.modifiedCount} detailed + ${directResult.modifiedCount} direct marks to trash. Removed ${gradeResult.deletedCount} grades.`,
      detail: {detailed: detailedResult.modifiedCount, direct: directResult.modifiedCount, grades: gradeResult.deletedCount},
    });
  }),
);

// ─── 2. GET TRASHED MARKS ─────────────────────────────────────────────────────
// Returns trashed records from both Mark and MarkDirect, tagged with `source`.
router.get(
  "/trash-bin",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const populateOpts = [
      { path: "student",     select: "regNo name" },
      { path: "programUnit", populate: { path: "unit", select: "code name" } },
      { path: "academicYear",select: "year" },
    ];

    const [detailedTrashed, directTrashed] = await Promise.all([
      Mark.find({ deletedAt: { $ne: null } })
        .populate(populateOpts).sort({ deletedAt: -1 }).limit(100).lean(),

      MarkDirect.find({ deletedAt: { $ne: null } })
        .populate(populateOpts).sort({ deletedAt: -1 }).limit(100).lean(),
    ]);

    // Tag each record so the frontend knows which collection to target on restore/purge
    const combined = [
      ...detailedTrashed.map((m) => ({ ...m, source: "detailed" })),
      ...directTrashed.map((m)   => ({ ...m, source: "direct"   })),
    ].sort((a, b) => new Date(b.deletedAt!).getTime() - new Date(a.deletedAt!).getTime());

    res.json(combined);
  }),
);

// ─── 3. RESTORE OR PERMANENT DELETE ──────────────────────────────────────────
// Each entry in markIds must carry a `source` field ("detailed" | "direct") so
// we can route the action to the correct collection.
// Format: markIds = [{ id: "...", source: "detailed" | "direct" }, ...]
// For backwards compatibility, plain string IDs are treated as "detailed".
router.post(
  "/trash-action",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { markIds, action } = req.body;

    if (!Array.isArray(markIds)) markIds = [markIds];

    // Normalise: support both plain string IDs (legacy) and { id, source } objects
    const normalised: Array<{ id: string; source: "detailed" | "direct" }> = markIds.map(
      (entry: any) => {
        if (typeof entry === "string") return { id: entry, source: "detailed" as const };
        return { id: entry.id || entry._id, source: entry.source || "detailed" };
      },
    );

    const detailedIds = normalised.filter((e) => e.source === "detailed").map((e) => e.id);
    const directIds   = normalised.filter((e) => e.source === "direct")  .map((e) => e.id);

    // ── RESTORE ─────────────────────────────────────────────────────────────
    if (action === "restore") {
      // Restore detailed marks
      if (detailedIds.length > 0) {
        const marks = await Mark.find({ _id: { $in: detailedIds }, deletedAt: { $ne: null } });
        await Mark.updateMany({ _id: { $in: detailedIds } }, { $set: { deletedAt: null } });
        for (const mark of marks) {
          try {
            await computeFinalGrade({ markId: mark._id as any });
          } catch (e: any) {
            console.warn(`[maintenance] FinalGrade recalc failed for detailed ${mark._id}: ${e.message}`);
          }
        }
      }

      // Restore direct marks
      if (directIds.length > 0) {
        const directMarks = await MarkDirect.find({ _id: { $in: directIds }, deletedAt: { $ne: null } });
        await MarkDirect.updateMany({ _id: { $in: directIds } }, { $set: { deletedAt: null } });
        for (const mark of directMarks) {
          try {
            await computeFinalGrade({ markId: mark._id as any });
          } catch (e: any) {
            console.warn(`[maintenance] FinalGrade recalc failed for direct ${mark._id}: ${e.message}`);
          }
        }
      }

      console.log(`[maintenance] Restored: ${detailedIds.length} detailed, ${directIds.length} direct`);
      return res.json({
        message: `Restored ${detailedIds.length} detailed and ${directIds.length} direct marks. Grades recalculated.`,
      });
    }

    // ── PURGE ────────────────────────────────────────────────────────────────
    if (action === "purge") {
      const [detailedDel, directDel] = await Promise.all([
        detailedIds.length > 0 ? Mark.deleteMany({ _id: { $in: detailedIds } }) : Promise.resolve({ deletedCount: 0 }),
        directIds.length   > 0 ? MarkDirect.deleteMany({ _id: { $in: directIds } }) : Promise.resolve({ deletedCount: 0 }),
      ]);

      console.log(`[maintenance] Purged: ${detailedDel.deletedCount} detailed, ${directDel.deletedCount} direct`);
      return res.json({
        message: `Permanently deleted ${detailedDel.deletedCount} detailed and ${directDel.deletedCount} direct marks.`,
      });
    }

    res.status(400).json({ error: "Invalid action. Use 'restore' or 'purge'." });
  }),
);

export default router;

