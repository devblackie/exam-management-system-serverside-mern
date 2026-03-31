// src/routes/units.ts 

import { Response, Router } from "express";
import Unit from "../models/Unit";
import ProgramUnit from "../models/ProgramUnit";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

// CREATE UNIT TEMPLATE (No longer linked to a program here)
router.post(
  "/",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { code, name } = req.body;
    const institutionId = req.user.institution;

    if (!code || !name) {
      return res.status(400).json({ message: "Unit code and name are required." });
    }

    const unitCode = code.toUpperCase();
    const unitName = name.trim();

    // 1. Check for duplicate unit code within the institution
    const duplicate = await Unit.findOne({
      code: unitCode,
      institution: institutionId,
    });

    if (duplicate) {
      return res.status(400).json({ message: `Unit code '${unitCode}' already exists.` });
    }

    // 2. Create the standalone unit template
    const newUnit = new Unit({
      code: unitCode,
      name: unitName,
      institution: institutionId,
      createdBy: req.user._id,
    });

    await newUnit.save();

    res.status(201).json(newUnit);
  })
);



router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const units = await Unit.find({ institution: req.user.institution }).sort({ code: 1 });
    res.json(units);
  })
);

// UPDATE UNIT TEMPLATE
router.put(
 "/:id",
 requireAuth,
 requireRole("admin", "coordinator"),
 asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
 const { code, name } = req.body;
 const unitId = req.params.id;
 const institutionId = req.user.institution;

 const updateData: any = {};
 const isCodeModified = code !== undefined;
 const isNameModified = name !== undefined;
 
 if (isCodeModified) updateData.code = code.toUpperCase();
 if (isNameModified) updateData.name = name.trim();

 if (Object.keys(updateData).length === 0) {
     return res.status(400).json({ message: "No fields provided for update." });
 }

 // --- 1. PRE-CHECK: FIND EXISTING UNIT ---
    const existingUnit = await Unit.findById(unitId);
    if (!existingUnit) {
      return res.status(404).json({ message: "Unit template not found" });
    }
    if (existingUnit.institution.toString() !== institutionId.toString()) {
        return res.status(403).json({ message: "Access denied." });
    }


 // --- 2. CONSTRAINT ENFORCEMENT: Check if linked to any Program ---
    // If EITHER code or name is being modified, we must check for links.
    // NOTE: For simplicity, we are preventing *any* modification if linked. 
    // Allowing name change but not code change requires more granular logic, but this is safer.
 if (isCodeModified || isNameModified) {
 const linkedCount = await ProgramUnit.countDocuments({ 
 unit: unitId,
 institution: institutionId 
});

 if (linkedCount > 0) {
 return res.status(400).json({ 
message: `Cannot modify this Unit Template. It is currently linked to ${linkedCount} program(s) in the curriculum. Please delink it first.` 
});
}
 }
 // --- END CONSTRAINT CHECK ---
    
    // --- 3. PRE-CHECK: Duplicate Code Check ---
    if (isCodeModified) {
        // Check if the NEW code already exists on another unit template
        const duplicate = await Unit.findOne({
            code: updateData.code,
            institution: institutionId,
            _id: { $ne: unitId }, // Exclude the current unit being updated
        });
        if (duplicate) {
            return res.status(400).json({ message: `Unit code '${updateData.code}' already exists on another template.` });
        }
    }
    // --- END Duplicate Check ---


// --- 4. PERFORM UPDATE ---
 const unit = await Unit.findOneAndUpdate(
{ _id: unitId, institution: institutionId },
 updateData,
 { new: true, runValidators: true }
 );

    // This check is now redundant since we found the unit earlier, but harmless.
if (!unit) {
return res.status(404).json({ message: "Unit template not found" });
}

res.json(unit);
 })
);

// DELETE UNIT TEMPLATE
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin","coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const unitId = req.params.id;
    const institutionId = req.user.institution;

    // --- CONSTRAINT ENFORCEMENT: Check if linked to any Program ---
    const linkedCount = await ProgramUnit.countDocuments({ 
      unit: unitId,
      institution: institutionId 
    });

    if (linkedCount > 0) {
      return res.status(400).json({ 
        message: `Cannot delete unit template. It is currently linked to ${linkedCount} program(s) in the curriculum. Please remove all curriculum links first.` 
      });
    }
    // --- END CONSTRAINT CHECK ---
    
    const unit = await Unit.findOneAndDelete({
      _id: unitId,
      institution: institutionId,
    });

    if (!unit) {
      return res.status(404).json({ message: "Unit template not found" });
    }

    res.json({ message: "Unit template deleted successfully" });
  })
);

export default router;