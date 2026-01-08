import { Response, Router } from "express";
import ProgramUnit from "../models/ProgramUnit";
import Mark from "../models/Mark";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

// --- POST /program-units: Link a Unit to a Program (Curriculum Definition) ---
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, unitId, requiredYear, requiredSemester, isElective } = req.body;

    if (!programId || !unitId || !requiredYear || !requiredSemester) {
      return res.status(400).json({ 
        message: "Program ID, Unit ID, Year, and Semester are required." 
      });
    }

    const institutionId = req.user.institution;
    
    // Check if the link already exists (enforced by index, but good for user feedback)
    const exists = await ProgramUnit.findOne({
      program: programId,
      unit: unitId,
      institution: institutionId,
    });

    if (exists) {
      return res.status(400).json({ 
        message: "This unit is already linked to this program. Use PUT to update." 
      });
    }

    const programUnit = await ProgramUnit.create({
      institution: institutionId,
      program: programId,
      unit: unitId,
      requiredYear: Number(requiredYear),
      requiredSemester: Number(requiredSemester),
      isElective: Boolean(isElective),
    });

    // Populate references for a rich response
    await programUnit.populate([
      { path: "program", select: "name code" },
      { path: "unit", select: "name code" }
    ]);

    res.status(201).json(programUnit);
  })
);

// --- GET /program-units?programId: Get all Units for a specific Program (Curriculum View) ---
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId } = req.query;

    if (!programId) {
      return res.status(400).json({ 
        message: "A programId query parameter is required to view the curriculum." 
      });
    }

    // 1. Find all ProgramUnit documents matching the programId
    const programUnits = await ProgramUnit.find({
      institution: req.user.institution,
      program: programId,
    })
    // 2. Populate the linked Program and Unit details
    .populate([
      { path: "program", select: "name code" },
      { path: "unit", select: "name code" }
    ])
    // 3. Sort by year, semester, and unit code for a readable curriculum list
    .sort({ requiredYear: 1, requiredSemester: 1, "unit.code": 1 }) 
    .lean();

    // 4. Clean the output for the frontend
    const formattedCurriculum = programUnits.map((pu: any) => ({
      _id: pu._id.toString(),
      requiredYear: pu.requiredYear,
      requiredSemester: pu.requiredSemester,
      isElective: pu.isElective,
      // Flatten the Unit details
      unit: {
        _id: pu.unit._id.toString(),
        code: pu.unit.code,
        name: pu.unit.name,
      },
      // Keep Program ID/Name context if needed
      program: {
          _id: pu.program._id.toString(),
          name: pu.program.name,
      }
    }));

    res.json(formattedCurriculum);
  })
);

// --- PUT /program-units/:id: Update a Curriculum Link (Year/Semester/Elective Status) ---
router.put(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    // We only allow changing the context fields, not the unit or program link itself.
    const { requiredYear, requiredSemester, isElective } = req.body;

    const updateData: any = {};
    if (requiredYear !== undefined) updateData.requiredYear = Number(requiredYear);
    if (requiredSemester !== undefined) updateData.requiredSemester = Number(requiredSemester);
    if (isElective !== undefined) updateData.isElective = Boolean(isElective);

    // If no valid fields are provided for update
    if (Object.keys(updateData).length === 0) {
        return res.status(400).json({ message: "No valid fields provided for update." });
    }

    const programUnit = await ProgramUnit.findOneAndUpdate(
      { _id: req.params.id, institution: req.user.institution },
      updateData,
      { new: true, runValidators: true }
    )
    .populate([
        { path: "program", select: "name code" },
        { path: "unit", select: "name code" }
    ]);

    if (!programUnit) {
      return res.status(404).json({ message: "Curriculum link not found." });
    }

    res.json(programUnit);
  })
);


// --- DELETE /program-units/:id: Remove a Curriculum Link ---
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const programUnitId = req.params.id;
    const institutionId = req.user.institution;
    
    // --- CONSTRAINT ENFORCEMENT: Check if student marks are recorded for this link ---
    const marksCount = await Mark.countDocuments({
      programUnit: programUnitId, // Assuming your Mark model links to ProgramUnit
      institution: institutionId
    });

    if (marksCount > 0) {
      return res.status(400).json({
        message: `Cannot remove this curriculum link. ${marksCount} student marks rely on this link for historical tracking. Please archive the link instead of deleting it.`
      });
    }
    // --- END CONSTRAINT CHECK ---

    const programUnit = await ProgramUnit.findOneAndDelete({
      _id: programUnitId,
      institution: institutionId,
    });

    if (!programUnit) {
      return res.status(404).json({ message: "Curriculum link not found." });
    }

    res.json({ message: "Unit successfully delinked from the program (Curriculum updated)." });
  })
);
export default router;