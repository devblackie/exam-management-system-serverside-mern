// src/routes/coordinator.ts
import { Router, Response, Request } from "express";
import bcrypt from "bcryptjs";
import User from "../models/User";
import Institution from "../models/Institution";
import { logAudit } from "../lib/auditLogger";
import { getScopedProgramIds, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import FinalGrade from "../models/FinalGrade";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";
import { cleanupOrphanedGrades } from "../scripts/cleanupGrades";
import mongoose from "mongoose";
import Program from "../models/Program";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import DisciplinaryCase from "../models/DisciplinaryCase";

const router = Router();

// Coordinator secret registration
router.post("/secret-register", asyncHandler(async (req: Request, res: Response) => {
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

router.post("/maintain/cleanup-grades", requireAuth,  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  try {
    console.log("Admin initiated manual database cleanup...");
    await cleanupOrphanedGrades();
    
    res.json({ 
      success: true, 
      message: "Data integrity restored. Orphaned grades have been purged." 
    });
  } catch (error) {
    console.error("Cleanup Route Error:", error);
    res.status(500).json({ 
      success: false, 
      error: "Failed to perform database maintenance." 
    });
  }
}));

// ── GET /coordinator/dashboard-stats ─────────────────────────────────────────
router.get("/dashboard-stats", requireAuth, requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId     = req.user.institution;
    const allowedProgramIds = await getScopedProgramIds(req);

    // Resolve student IDs in scope first — needed for disciplinary sub-query
    const scopedStudentIds = await Student.find({
      institution: institutionId,
      program:     { $in: allowedProgramIds },
    })
      .select("_id")
      .lean()
      .then(ss => ss.map(s => s._id));

    // ── All queries in parallel ───────────────────────────────────────────────
    const [
      studentCounts,
      programs,
      currentYear,
      openCases,
      markBatchIds,
      directMarkBatchIds,
      lastMark,
      lastDirectMark,
    ] = await Promise.all([

      // Student status breakdown
      Student.aggregate([
        {
          $match: {
            institution: new mongoose.Types.ObjectId(institutionId.toString()),
            program:     { $in: allowedProgramIds },
          },
        },
        { $group: { _id: "$status", count: { $sum: 1 } } },
      ]) as Promise<Array<{ _id: string; count: number }>>,

      // Scoped program names
      Program.find({
        institution: institutionId,
        _id:         { $in: allowedProgramIds },
        isActive:    true,
      })
        .select("name")
        .lean<Array<{ name: string }>>(),

      // Current academic year
      AcademicYear.findOne({ institution: institutionId, isCurrent: true })
        .select("year session")
        .lean<{ year: string; session?: string } | null>(),

      // Open disciplinary cases for students in scope
      DisciplinaryCase.countDocuments({
        institution: institutionId,
        outcome:     "PENDING",
        student:     { $in: scopedStudentIds },
      }),

      // Distinct mark batch IDs
      Mark.distinct("batchId", {
        institution: institutionId,
        program:     { $in: allowedProgramIds },
      }) as Promise<string[]>,

      // Distinct direct mark batch IDs
      MarkDirect.distinct("batchId", {
        institution: institutionId,
        program:     { $in: allowedProgramIds },
      }) as Promise<string[]>,

      // Most recent detailed mark upload
      Mark.findOne({
        institution: institutionId,
        program:     { $in: allowedProgramIds },
      })
        .sort({ uploadedAt: -1 })
        .select("uploadedAt")
        .lean<{ uploadedAt?: Date } | null>(),

      // Most recent direct mark upload
      MarkDirect.findOne({
        institution: institutionId,
        program:     { $in: allowedProgramIds },
      })
        .sort({ uploadedAt: -1 })
        .select("uploadedAt")
        .lean<{ uploadedAt?: Date } | null>(),
    ]);

    // ── Aggregate student statuses ────────────────────────────────────────────
    const sm: Record<string, number> = {};
    for (const row of studentCounts) sm[row._id] = row.count;
    const total = Object.values(sm).reduce((a, b) => a + b, 0);

    // ── Resolve last upload date ──────────────────────────────────────────────
    const d1 = lastMark?.uploadedAt       ? new Date(lastMark.uploadedAt).getTime()       : 0;
    const d2 = lastDirectMark?.uploadedAt ? new Date(lastDirectMark.uploadedAt).getTime() : 0;
    const lastUploadDate =
      d1 === 0 && d2 === 0
        ? null
        : new Date(Math.max(d1, d2)).toISOString();

    res.json({
      students: {
        total,
        active:       sm["active"]                  ?? 0,
        repeat:       sm["repeat"]                  ?? 0,
        discontinued: sm["discontinued"]            ?? 0,
        graduated:    sm["graduated"]               ?? 0,
        suspended:    sm["disciplinary_suspension"] ?? 0,
      },
      marks: {
        totalUploads:  markBatchIds.length + directMarkBatchIds.length,
        pendingReview: 0,
        lastUploadDate,
      },
      disciplinary: {
        openCases:      openCases,
        pendingOutcome: openCases,
      },
      programs: {
        total: programs.length,
        names: programs.map(p => p.name),
      },
      promotion: {
        lastRunDate:   null,
        eligibleCount: 0,
      },
      academicYear: {
        current: currentYear?.year    ?? null,
        session: currentYear?.session ?? null,
      },
    });
  }),
);

// ── GET /coordinator/lecturers ────────────────────────────────────────────────
// Coordinators can list lecturers in their department
router.get("/lecturers", requireAuth, requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const User = (await import("../models/User")).default;
    const lecturers = await User.find({
      institution:    req.user.institution,
      role:           "lecturer",
      departmentCode: req.user.departmentCode,
    })
      .select("name email departmentCode schoolCode createdAt")
      .lean();
    res.json(lecturers);
  }),
);

// ── POST /coordinator/lecturers ───────────────────────────────────────────────
router.post("/lecturers", requireAuth, requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { name, email, password } = req.body as {
      name:      string;
      email:     string;
      password?: string;
    };

    if (!name?.trim() || !email?.trim()) {
      res.status(400).json({ message: "name and email are required." });
      return;
    }

    const User   = (await import("../models/User")).default;
    const bcrypt = await import("bcryptjs");

    const existing = await User.findOne({ email: email.toLowerCase() });
    if (existing) {
      res.status(409).json({ message: "A user with this email already exists." });
      return;
    }

    const hash = await bcrypt.hash(
      password?.trim() || Math.random().toString(36).slice(-10),
      10,
    );

    const lecturer = await User.create({
      name:           name.trim(),
      email:          email.toLowerCase().trim(),
      password:       hash,
      role:           "lecturer",
      institution:    req.user.institution,
      schoolCode:     req.user.schoolCode,
      departmentCode: req.user.departmentCode,
      isVerified:     true,
    });

    res.status(201).json({
      message: "Lecturer created.",
      lecturer: { _id: lecturer._id, name: lecturer.name, email: lecturer.email },
    });
  }),
);

// Coordinator creates lecturer (no login needed)
router.post("/lecturers", requireAuth, requireRole("coordinator", "admin"),
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
router.get("/students/:regNo/results", requireAuth,
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
        path: "programUnit",
        // We need the unit (template) and the required year/semester from the link
        select: "requiredYear requiredSemester unit", 
        populate: { path: "unit", model: "Unit", select: "code name" }
      })
      .populate({ path: "academicYear", select: "year"})
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
      student: { name: student.name, regNo: student.regNo, program: student.program },
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

