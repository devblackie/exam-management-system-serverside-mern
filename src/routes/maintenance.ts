// src/routes/maintenance.ts
import express, { Response } from "express";
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import {
  AuthenticatedRequest,
  requireAuth,
  requireRole,
} from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";

const router = express.Router();

// 1. BULK SOFT DELETE
router.post(
  "/bulk-cleanup",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { unitCode, programId, academicYear } = req.body;
    // console.log(`[CLEANUP] Initiated by ${req.user.email}. Params:`, req.body);

    if (!academicYear) return res.status(400).json({ error: "Academic Year is required" });
    
    // 1. Resolve IDs
    const yearDoc = await AcademicYear.findOne({ year: academicYear });
    const unitDoc = unitCode ? await Unit.findOne({ code: unitCode }) : null;

    if (!yearDoc) return res.status(404).json({ error: "Academic Year not found" });
    

    // 2. Build Query
    let query: any = {
      academicYear: yearDoc._id,      
    };

    if (programId) {
      const studentIds = await Student.find({ program: programId }).distinct( "_id" );
      query.student = { $in: studentIds };
    }

    if (unitDoc) {
      const pUnits = await ProgramUnit.find({ unit: unitDoc._id }).distinct("_id" );
      query.programUnit = { $in: pUnits };
    }

    const marksToTrash = await Mark.find({ ...query, deletedAt: null });  
    const updateResult = await Mark.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() }});
    const gradeResult = await FinalGrade.deleteMany(query);

    res.json({
      count: updateResult.modifiedCount,
      message: `Successfully moved ${updateResult.modifiedCount} marks to trash and removed ${gradeResult.deletedCount} grades.`,
    });
  }),
);

// 2. GET TRASHED MARKS
router.get(
  "/trash-bin",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    // PASS THE CRITERIA DIRECTLY TO FIND()
    const trashed = await Mark.find({ deletedAt: { $ne: null } }) 
      .populate("student", "regNo name")
      .populate({ path: "programUnit", populate: { path: "unit", select: "code" }})
      .populate("academicYear", "year")
      .sort({ deletedAt: -1 })
      .limit(100);

    // console.log("Trashed marks found:", trashed.length);
    res.json(trashed);
  }),
);

// 3. RESTORE OR PERMANENT DELETE
router.post(
  "/trash-action",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { markIds, action } = req.body;                

    // Ensure markIds is an array
    if (!Array.isArray(markIds)) markIds = [markIds];
    
    
    // console.log(`[MAINTENANCE] Action: ${action}, IDs:`, markIds);
    if (action === "restore") {
      // console.log(`[MAINTENANCE] Restoring marks: ${markIds}`);
      // 1. Restore the marks
      const marks = await Mark.find({ _id: { $in: markIds }, deletedAt: { $ne: null }});
      // console.log(`[MAINTENANCE] Found ${marks.length} marks to restore.`);


      await Mark.updateMany({ _id: { $in: markIds } }, { $set: { deletedAt: null }});
      // console.log(`[MAINTENANCE] Mark documents updated to null deletedAt.`);

      // 2. RECALCULATE GRADES for those marks
      for (const mark of marks) await computeFinalGrade({ markId: mark._id as any})
     
      return res.json({ message: "Marks restored and grades recalculated" });
    }

    if (action === "purge") {
      await Mark.deleteMany({ _id: { $in: markIds } });
      return res.json({ message: "Marks permanently deleted from database" });
    }

    res.status(400).json({ error: "Invalid action" });
  }),
);



export default router;

