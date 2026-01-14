// src/routes/coordinator.ts
import { Router, Response, Request } from "express";
import bcrypt from "bcryptjs";
import User from "../models/User";
import Institution from "../models/Institution";
import { logAudit } from "../lib/auditLogger";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import FinalGrade from "../models/FinalGrade";
import Student from "../models/Student";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";

const router = Router();

// Coordinator secret registration
router.post(
  "/secret-register",
  asyncHandler(async (req: Request, res: Response) => {
    const { secret, name, email, password ,institutionId } = req.body;

    // Validate Secret Key
    if (secret !== process.env.COORDINATOR_SECRET) {
      await logAudit(req, { action: "coordinator_register_failed_invalid_secret" });
      return res.status(403).json({ message: "Invalid secret" });
    }

    // Check duplicate email
    const existing = await User.findOne({ email });
    if (existing) {
      await logAudit(req, { action: "coordinator_register_failed_duplicate", details: { email } });
      return res.status(400).json({ message: "Email already in use" });
    }

     // Validate institution exists (if provided)
    let institution = null;
    if (institutionId) {
      // console.log("Received institutionId:", institutionId);

      // institution = await Institution.findById(institutionId);

      // Validate it's a proper ObjectId string
      if (!institutionId.match(/^[0-9a-fA-F]{24}$/)) {
        // console.log("Invalid ObjectId format:", institutionId);
        return res.status(400).json({ message: "Invalid institution ID format" });
      }

       institution = await Institution.findById(institutionId);
      // console.log("Found institution:", institution); // ← DEBUG LOG

      if (!institution) {
        // console.log("Institution not found in DB for ID:", institutionId);
        return res.status(400).json({ message: "Invalid institution ID" });
      }

      if (!institution.isActive) {
        return res.status(400).json({ message: "Institution is not active" });
      }
      
      // if (!institution) return res.status(400).json({ message: "Invalid institution" });
    }

    // Create coordinator
    const hashed = await bcrypt.hash(password, 12);
    const coordinator = await User.create({
      name,
      email: email.toLowerCase(),
      password: hashed,
      role: "coordinator",
      institution: institution?._id || null, // optional: assign institution
    });

    await logAudit(req, {
      action: "coordinator_registered",
      targetUser: coordinator._id,
      details: { email, name,institution: institutionId },
    });

    res.status(201).json({ message: "Coordinator created successfully" });
  })
);

// Coordinator creates lecturer (no login needed)
router.post(
  "/lecturers",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { name, email } = req.body;

    if (!name || !email) {
      return res.status(400).json({ message: "Name and email required" });
    }

    const existing = await User.findOne({ email: email.toLowerCase() });
    if (existing) {
      return res.status(400).json({ message: "Email already exists" });
    }

    const password = Math.random().toString(36).slice(-8) + "A1!";
    const hashed = await bcrypt.hash(password, 12);

    const lecturer = await User.create({
      name,
      email: email.toLowerCase(),
      password: hashed,
      role: "lecturer",
      institution: req.user.institution,
      status: "active",
    });

    await logAudit(req, {
      action: "lecturer_created",
      actor: req.user._id,
      targetUser: lecturer._id,
      details: { email, name },
    });

    res.status(201).json({
      message: "Lecturer created",
      email,
      temporaryPassword: password,
    });
  })
);

// View student results (senate-style)
router.get(
  "/students/:regNo/results",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo } = req.params;
    const student = await Student.findOne({
      regNo: regNo.toUpperCase(),
      institution: req.user.institution,
    });

    if (!student) {
      return res.status(404).json({ message: "Student not found" });
    }

    const grades = await FinalGrade.find({ student: student._id })
      .populate({
        // ⬅️ FIX 1: Populate the ProgramUnit link
        path: "programUnit",
        // We need the unit (template) and the required year/semester from the link
        select: "requiredYear requiredSemester unit", 
        populate: {
            // ⬅️ Nested populate to get the Unit Template details
            path: "unit",
            model: "Unit", // Specify model for nested population
            select: "code name" // Only need code and name from the Unit Template
        }
      })
      .populate({
        path: "academicYear",
        select: "year",
      })
      .sort({ "academicYear.year": 1, "programUnit.requiredYear": 1, "programUnit.requiredSemester": 1 }) // ⬅️ Adjusted sort fields
      .lean();

    await logAudit(req, {
      action: "coordinator_viewed_student_results",
      actor: req.user._id,
      // targetUser: student._id,
      details: { regNo },
    });
    
    // Define the type for the populated grade to improve safety and readability
    interface PopulatedFinalGrade extends Omit<typeof grades[0], 'programUnit' | 'academicYear'> {
        programUnit: { 
            requiredYear: number; 
            requiredSemester: number;
            unit: { code: string; name: string };
        };
        academicYear: { year: string };
    }


    res.json({
      student: {
        name: student.name,
        regNo: student.regNo,
        program: student.program,
      },
      results: grades.map(g => {
          const grade = g as unknown as PopulatedFinalGrade;
          
          return {
            // ⬅️ FIX 2: Access unit and scheduling details through programUnit
            unitCode: grade.programUnit.unit.code,
            unitName: grade.programUnit.unit.name,
            year: grade.programUnit.requiredYear,
            semester: grade.programUnit.requiredSemester,
            
            academicYear: grade.academicYear.year,
            totalMark: grade.totalMark,
            grade: grade.grade,
            status: grade.status,
            capped: grade.cappedBecauseSupplementary,
          };
      }),
    });
  })
);

export default router;