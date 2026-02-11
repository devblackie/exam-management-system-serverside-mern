// src/routes/studentSearch.ts
import express, { Response } from "express";
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import {
  AuthenticatedRequest,requireAuth, requireRole,
} from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";
import { calculateStudentStatus } from "../services/statusEngine";
import { scopeQuery } from "../lib/multiTenant";

const router = express.Router();

// SEARCH STUDENT BY REG NO
router.get(
  "/search",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { q } = req.query;

    // console.log("[STUDENT SEARCH] Query received:", q);

    if (!q || typeof q !== "string" || q.trim().length < 3) {
      return res.status(400).json({ error: "Enter at least 3 characters" });
    }

    const searchQuery = q.trim();

    const query = scopeQuery(req, {
      regNo: { $regex: `^${searchQuery}`, $options: "i" },
    });

    // console.log("Final Scoped Query:", JSON.stringify(query, null, 2));
    const students = await Student.find(query)
      .limit(10)
      .select("regNo name program admissionYear")
      .populate("program", "name");

    // console.log(
    //   `[STUDENT SEARCH] Found ${students.length} students for "${searchQuery}"`
    // );

    res.json(students);
  }),
);

// GET STUDENT FULL RESULTS + STATUS
router.get(
  "/record",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { regNo, yearOfStudy } = req.query;
    if (!regNo || typeof regNo !== "string")
      return res.status(400).json({ error: "regNo is required" });

    const targetYearOfStudy = parseInt(yearOfStudy as string) || 1;

    regNo = decodeURIComponent(regNo);
    const student = await Student.findOne(
      scopeQuery(req, {
        regNo: { $regex: `^${regNo}$`, $options: "i" },
      }),
    ).populate("program");

    if (!student) return res.status(404).json({ error: "Student not found" });

    // 1. Fetch grades with the correct nested path
    const grades = await FinalGrade.find({ student: student._id })
      .populate({
        path: "programUnit",
        populate: { path: "unit", select: "code name" },
      })
      .populate("academicYear", "year")
      .sort({ "academicYear.year": 1 })
      .lean();

    const academicYearName = (req.query.academicYear as string) || "2024/2025";

    // 2. Safe mapping to prevent "undefined" errors
    const processedGrades = grades
      .filter((g) => {
        const pUnit = g.programUnit as any;
        // Only include units meant for the Year of Study selected by the user
        return pUnit?.requiredYear === targetYearOfStudy;
      })
      .map((g) => {
        const pUnit = g.programUnit as any;
        return {
          ...g,
          unit: {
            code: pUnit?.unit?.code || "N/A",
            name: pUnit?.unit?.name || "Unknown Unit",
          },
          semester: pUnit?.requiredSemester || "N/A",
          academicYear: g.academicYear || { year: "N/A" },
        };
      });

    // 2. USE THE SERVICE for Status
    // We pass studentId, programId, the Year string, and Year of Study
    const academicStatus = await calculateStudentStatus(
      student._id,
      (student.program as any)._id,
      "", // academicYearName is left empty if status engine can derive it from yearOfStudy
      targetYearOfStudy,
    );

    res.json({
      student: {
        _id: student._id, name: student.name, regNo: student.regNo,
        programId: (student.program as any)?._id || student.program,
        programName: (student.program as any)?.name,
        currentYear: student.currentYearOfStudy || 1, // Student's actual level
        currentSemester: student.currentSemester || 1,
      },
      grades: processedGrades, // Only Year X grades
      viewingYearOfStudy: targetYearOfStudy,
      academicStatus: academicStatus,
      summary: academicStatus?.summary || {
        totalUnits: processedGrades.length,
        passed: processedGrades.filter((g) => g.status === "PASS").length,
        failed: processedGrades.filter((g) =>
          ["RETAKE", "SUPPLEMENTARY", "FAIL"].includes(g.status),
        ).length,
      },
    });
  }),
);

// POST /approve-special: One-click approval for Special Exams
router.post(
  "/approve-special",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { markId, reason } = req.body;

    if (!markId) {
      return res.status(400).json({ error: "Mark ID is required" });
    }

    // 1. Find the mark and update to Special status
    const mark = await Mark.findByIdAndUpdate(
      markId,
      {
        $set: {
          isSpecial: true,
          attempt: "special",
          remarks: reason || "Approved by Coordinator",
        },
      },
      { new: true },
    );

    if (!mark) {
      return res.status(404).json({ error: "Mark record not found" });
    }

    // 2. Trigger Grade Recalculation
    // This will update FinalGrade status to 'SPECIAL' and Grade to 'I'
    const gradeResult = await computeFinalGrade({
      markId: mark._id as any,
      coordinatorReq: req,
    });

    res.json({
      success: true,
      message: "Special exam approved successfully",
      newStatus: gradeResult.status,
    });
  }),
);

// GET RAW MARKS FOR A STUDENT (FOR COORDINATOR USE)
router.get(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, yearOfStudy } = req.query;

    if (!regNo) return res.status(400).json({ error: "regNo required" });

    const student = await Student.findOne({
      regNo: { $regex: `^${regNo}$`, $options: "i" },
    });

    if (!student) return res.status(404).json({ error: "Student not found" });

    let query: any = { student: student._id };

    // Fetch marks linked to ProgramUnits of a specific Year of Study
    const marks = await Mark.find(query)
      .populate({
        path: "programUnit",
        match: yearOfStudy
          ? { requiredYear: parseInt(yearOfStudy as string) }
          : {}, // FILTER BY YEAR
        populate: { path: "unit", select: "code name" },
      })
      .populate("academicYear", "year")
      .lean();

    // Filter out marks where programUnit didn't match the yearOfStudy
    const filteredMarks = marks.filter((m) => m.programUnit !== null);

    res.json(filteredMarks);
  }),
);

// POST /raw-marks: Upload or Update Raw Marks for a Student
router.post(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const {
      regNo, unitCode, academicYear,
      cat1, cat2, cat3,
      assignment1, practicalRaw,
      examQ1, examQ2, examQ3, examQ4, examQ5,
      attempt,
    } = req.body;

    if (!regNo || !unitCode || !academicYear) {
      return res.status(400).json({ error: "regNo, unitCode, and academicYear are required" });
    }

    const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" },});
    const unitDoc = await Unit.findOne({ code: unitCode });
    const yearDoc = await AcademicYear.findOne({ year: academicYear });

    if (!student || !unitDoc || !yearDoc) {
      return res.status(404).json({ error: "Student, Unit, or Academic Year not found" });
    }

    const programUnit = await ProgramUnit.findOne({
      program: student.program,
      unit: unitDoc._id,
    });
    if (!programUnit) {
      return res.status(400).json({ error: `Unit ${unitCode} is not linked to program.` });
    }

    const previousMarks = await Mark.find({
      student: student._id,
      programUnit: programUnit._id,
    }).sort({ createdAt: 1 });

    let detectedAttempt = "1st";
    let isSpecial = false;

    if (previousMarks.length > 0) {
      const lastMark = previousMarks[previousMarks.length - 1];

      // If the last mark was flagged as "Special" but never completed, keep it Special
      // Or if the coordinator explicitly requested a Special Exam
      if (req.body.isSpecial || lastMark.attempt === "special") {
        detectedAttempt = "special";
        isSpecial = true;
      }
      // If student previously had a 1st attempt and failed -> Supplementary
      else if (previousMarks.length === 1) {
        detectedAttempt = "supplementary";
      }
      // If student failed a Supp -> Retake
      else if (previousMarks.length === 2) {
        detectedAttempt = "re-take";
      }
      // If student failed a Retake -> Re-Retake
      else if (previousMarks.length >= 3) {
        detectedAttempt = "re-retake"; // Note: Ensure "re-retake" is in your Mark model enum
      }
    }

    const mark = await Mark.findOneAndUpdate(
      { student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id, },
      {
        cat1Raw: Number(cat1) || 0, cat2Raw: Number(cat2) || 0,
        cat3Raw: cat3 ? Number(cat3) : undefined,
        assgnt1Raw: Number(assignment1) || 0, practicalRaw: Number(practicalRaw) || 0, examQ1Raw: Number(examQ1) || 0,
        examQ2Raw: Number(examQ2) || 0, examQ3Raw: Number(examQ3) || 0,
        examQ4Raw: Number(examQ4) || 0, examQ5Raw: Number(examQ5) || 0,
        attempt: detectedAttempt,
        isSpecial: isSpecial,
        isSupplementary: detectedAttempt === "supplementary",
        isRetake: detectedAttempt.includes("re-take"),
        uploadedBy: req.user._id,
        uploadedAt: new Date(),
      },
      { upsert: true, new: true, setDefaultsOnInsert: true },
    ).populate({
      path: "programUnit",
      populate: { path: "unit", select: "code name" },
    });

    // Recalculate Final Grade
    const gradeResult = await computeFinalGrade({
      markId: mark!._id as any,
      coordinatorReq: req,
    });
    // console.log("[SUCCESS] Marks saved & grade recalculated:", {
    //   regNo, unit: unitCode,year: academicYear,finalGrade: gradeResult.grade,status: gradeResult.status, });
    //   });

    // res.json({ success: true, data: { caTotal: gradeResult.caTotal30, examTotal: gradeResult.examTotal70, finalMark: gradeResult.finalMark, grade: gradeResult.grade, status: gradeResult.status,},});

    res.json({ success: true, data: gradeResult });
  }),
);

export default router;
