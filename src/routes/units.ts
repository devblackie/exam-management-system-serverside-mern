// src/routes/units.ts 

import { Response, Router } from "express";
import Unit from "../models/Unit";
import ProgramUnit from "../models/ProgramUnit";
import { getScopedProgramIds, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { cached, invalidateCache } from "../utils/cache";

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
    invalidateCache(`settings:${req.user.institution}`);
    res.status(201).json(newUnit);
  })
);



// router.get(
//   "/",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const units = await Unit.find({ institution: req.user.institution }).sort({ code: 1 });
//     res.json(units);
//   })
// );

// GET
// router.get("/", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const institutionId = req.user.institution;
//   const units = await cached(`units:${institutionId}`, () => 
//     Unit.find({ institution: institutionId }).sort({ code: 1 }).lean()
//   );
//   res.json(units);
// }));

router.get("/", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const institutionId = req.user.institution;
  
  // SCOPING: Coordinators can only see units linked to their programs
  if (req.user.role === "coordinator" && !req.user.institutionWide) {
    const scopedProgramIds = await getScopedProgramIds(req);
    const programUnits = await ProgramUnit.find({ 
      program: { $in: scopedProgramIds },
      institution: institutionId 
    }).populate("unit").lean();
    
    const uniqueUnits = new Map();
    for (const pu of programUnits) {
      if (pu.unit && !uniqueUnits.has(pu.unit._id.toString())) {
        uniqueUnits.set(pu.unit._id.toString(), pu.unit);
      }
    }
    res.json(Array.from(uniqueUnits.values()));
  } else {
    const units = await cached(`units:${institutionId}`, () => 
      Unit.find({ institution: institutionId }).sort({ code: 1 }).lean()
    );
    res.json(units);
  }
}));

// UPDATE UNIT TEMPLATE
// router.put(
//  "/:id",
//  requireAuth,
//  requireRole("admin", "coordinator"),
//  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//  const { code, name } = req.body;
//  const unitId = req.params.id;
//  const institutionId = req.user.institution;

//  const updateData: any = {};
//  const isCodeModified = code !== undefined;
//  const isNameModified = name !== undefined;
 
//  if (isCodeModified) updateData.code = code.toUpperCase();
//  if (isNameModified) updateData.name = name.trim();

//  if (Object.keys(updateData).length === 0) {
//      return res.status(400).json({ message: "No fields provided for update." });
//  }

//  // --- 1. PRE-CHECK: FIND EXISTING UNIT ---
//     const existingUnit = await Unit.findById(unitId);
//     if (!existingUnit) {
//       return res.status(404).json({ message: "Unit template not found" });
//     }
//     if (existingUnit.institution.toString() !== institutionId.toString()) {
//         return res.status(403).json({ message: "Access denied." });
//     }


//  // --- 2. CONSTRAINT ENFORCEMENT: Check if linked to any Program ---
//     // If EITHER code or name is being modified, we must check for links.
//     // NOTE: For simplicity, we are preventing *any* modification if linked. 
//     // Allowing name change but not code change requires more granular logic, but this is safer.
//  if (isCodeModified || isNameModified) {
//  const linkedCount = await ProgramUnit.countDocuments({ 
//  unit: unitId,
//  institution: institutionId 
// });

//  if (linkedCount > 0) {
//  return res.status(400).json({ 
// message: `Cannot modify this Unit Template. It is currently linked to ${linkedCount} program(s) in the curriculum. Please delink it first.` 
// });
// }
//  }
//  // --- END CONSTRAINT CHECK ---
    
//     // --- 3. PRE-CHECK: Duplicate Code Check ---
//     if (isCodeModified) {
//         // Check if the NEW code already exists on another unit template
//         const duplicate = await Unit.findOne({
//             code: updateData.code,
//             institution: institutionId,
//             _id: { $ne: unitId }, // Exclude the current unit being updated
//         });
//         if (duplicate) {
//             return res.status(400).json({ message: `Unit code '${updateData.code}' already exists on another template.` });
//         }
//     }
//     // --- END Duplicate Check ---


// // --- 4. PERFORM UPDATE ---
//  const unit = await Unit.findOneAndUpdate(
// { _id: unitId, institution: institutionId },
//  updateData,
//  { new: true, runValidators: true }
//  );

//     // This check is now redundant since we found the unit earlier, but harmless.
// if (!unit) {
// return res.status(404).json({ message: "Unit template not found" });
// }
// invalidateCache(`settings:${req.user.institution}`);
// res.json(unit);
//  })
// );

// UPDATE UNIT TEMPLATE - WITH SCOPING
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

    const existingUnit = await Unit.findById(unitId);
    if (!existingUnit) {
      return res.status(404).json({ message: "Unit template not found" });
    }
    if (existingUnit.institution.toString() !== institutionId.toString()) {
      return res.status(403).json({ message: "Access denied." });
    }

    // SCOPING: Coordinator can only edit units linked to their programs
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
      const scopedProgramIds = await getScopedProgramIds(req);
      const linkedPrograms = await ProgramUnit.find({ 
        unit: unitId, 
        institution: institutionId 
      }).distinct("program");
      
      const hasOwnProgram = linkedPrograms.some(pid => scopedProgramIds.includes(pid.toString()));
      if (!hasOwnProgram && linkedPrograms.length > 0) {
        return res.status(403).json({ 
          message: "You can only edit units that are linked to programs within your department." 
        });
      }
    }

    // Check if linked to any program when modifying code/name
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
    
    if (isCodeModified) {
      const duplicate = await Unit.findOne({
        code: updateData.code,
        institution: institutionId,
        _id: { $ne: unitId },
      });
      if (duplicate) {
        return res.status(400).json({ message: `Unit code '${updateData.code}' already exists on another template.` });
      }
    }

    const unit = await Unit.findOneAndUpdate(
      { _id: unitId, institution: institutionId },
      updateData,
      { new: true, runValidators: true }
    );

    if (!unit) {
      return res.status(404).json({ message: "Unit template not found" });
    }
    invalidateCache(`units:${req.user.institution}`);
    res.json(unit);
  })
);

// DELETE UNIT TEMPLATE
// router.delete(
//   "/:id",
//   requireAuth,
//   requireRole("admin","coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const unitId = req.params.id;
//     const institutionId = req.user.institution;

//     // --- CONSTRAINT ENFORCEMENT: Check if linked to any Program ---
//     const linkedCount = await ProgramUnit.countDocuments({ 
//       unit: unitId,
//       institution: institutionId 
//     });

//     if (linkedCount > 0) {
//       return res.status(400).json({ 
//         message: `Cannot delete unit template. It is currently linked to ${linkedCount} program(s) in the curriculum. Please remove all curriculum links first.` 
//       });
//     }
//     // --- END CONSTRAINT CHECK ---
    
//     const unit = await Unit.findOneAndDelete({
//       _id: unitId,
//       institution: institutionId,
//     });

//     if (!unit) {
//       return res.status(404).json({ message: "Unit template not found" });
//     }
//     invalidateCache(`settings:${req.user.institution}`);
//     res.json({ message: "Unit template deleted successfully" });
//   })
// );

router.delete(
  "/:id",
  requireAuth,
  requireRole("admin", "coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const unitId = req.params.id;
    const institutionId = req.user.institution;

    // SCOPING: Coordinator can only delete units linked to their programs
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
      const scopedProgramIds = await getScopedProgramIds(req);
      const linkedPrograms = await ProgramUnit.find({
        unit: unitId,
        institution: institutionId,
      }).distinct("program");

      const hasOwnProgram = linkedPrograms.some((pid) =>
        scopedProgramIds.includes(pid.toString()),
      );
      if (!hasOwnProgram && linkedPrograms.length > 0) {
        return res.status(403).json({
          message:
            "You can only delete units that are linked to programs within your department.",
        });
      }
    }

    const linkedCount = await ProgramUnit.countDocuments({
      unit: unitId,
      institution: institutionId,
    });

    if (linkedCount > 0) {
      return res.status(400).json({
        message: `Cannot delete unit template. It is currently linked to ${linkedCount} program(s) in the curriculum. Please remove all curriculum links first.`,
      });
    }

    const unit = await Unit.findOneAndDelete({
      _id: unitId,
      institution: institutionId,
    });

    if (!unit) {
      return res.status(404).json({ message: "Unit template not found" });
    }
    invalidateCache(`units:${req.user.institution}`);
    res.json({ message: "Unit template deleted successfully" });
  }),
);

export default router;