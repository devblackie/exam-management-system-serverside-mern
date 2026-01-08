// src/lib/assignLecturer.ts
import { Request, Response } from "express";
import UnitAssignment from "../models/UnitAssignment";
import Unit from "../models/Unit";
import User from "../models/User";

export const assignLecturer = async (req: Request, res: Response) => {
  try {
    if (!req.user) {
      return res.status(401).json({ message: "Unauthorized" });
    }
    const { lecturerId, unitIds } = req.body; // unitIds = array
    const adminId = req.user.id;

    // ✅ Ensure lecturer exists
    const lecturer = await User.findOne({ _id: lecturerId, role: "lecturer" });
    if (!lecturer) return res.status(404).json({ message: "Lecturer not found" });

    // ✅ Ensure units exist
    const units = await Unit.find({ _id: { $in: unitIds } });
    if (units.length !== unitIds.length) {
      return res.status(400).json({ message: "Some units not found" });
    }

    // ✅ Filter out duplicates
    const existingAssignments = await UnitAssignment.find({
      lecturer: lecturerId,
      unit: { $in: unitIds },
    }).select("unit");

    const alreadyAssignedIds = existingAssignments.map(a => a.unit.toString());
    const newUnitIds = unitIds.filter((id: string) => !alreadyAssignedIds.includes(id));

    // ✅ Insert new assignments in bulk
    const newAssignments = newUnitIds.map((unitId: string) => ({
      lecturer: lecturerId,
      unit: unitId,
      assignedBy: adminId,
    }));

    if (newAssignments.length > 0) {
      await UnitAssignment.insertMany(newAssignments);
    }

    res.status(201).json({
      message: "Lecturer assigned to selected units successfully",
      assigned: newAssignments.length,
      skipped: alreadyAssignedIds.length,
    });
  } catch (err) {
    res.status(500).json({ message: "Server error", error: err });
  }
};
