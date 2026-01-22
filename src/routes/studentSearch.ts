// src/routes/studentSearch.ts
import express, { Response } from "express";
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import {
  AuthenticatedRequest,
  requireAuth,
  requireRole,
} from "../middleware/auth";
import { generateStudentTranscript } from "../services/pdfGenerator";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";
import { calculateStudentStatus } from "../services/statusEngine";

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

    const students = await Student.find({
      regNo: { $regex: `^${searchQuery}`, $options: "i" }, // Case-insensitive + starts with
    })
      .limit(10)
      .select("regNo name program admissionYear")
      .populate("program", "name");

    // console.log(
    //   `[STUDENT SEARCH] Found ${students.length} students for "${searchQuery}"`
    // );

    res.json(students);
  })
);

// GET STUDENT FULL RESULTS + STATUS
router.get(
  "/record",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { regNo } = req.query;
    if (!regNo || typeof regNo !== "string")
      return res.status(400).json({ error: "regNo is required" });

    regNo = decodeURIComponent(regNo);
    const student = await Student.findOne({
      regNo: { $regex: `^${regNo}$`, $options: "i" },
    }).populate("program");

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
    const processedGrades = grades.map((g) => {
      const pUnit = g.programUnit as any;

      // if (!pUnit) {
      //   console.log(`[DEBUG] Grade ${g._id} missing programUnit for student ${regNo}`);
      // }

      // 2. Format the semester for the UI (e.g., convert 1 to "1" or "SEMESTER 1")
      const semesterValue = pUnit?.requiredSemester || "N/A";

      return {
        ...g,
        // Provide a fallback unit object so frontend doesn't crash
        unit: {
          // If programUnit is missing, try to see if 'unit' was stored directly on 'g'
          code: pUnit?.unit?.code || "N/A",
          name: pUnit?.unit?.name || "Unknown Unit",
        },
        // Semester is usually stored directly on the Grade record in most setups
        semester: semesterValue,
        academicYear: g.academicYear || { year: "N/A" },
      };
    });

   // 2. USE THE SERVICE for Status
    // We pass studentId, programId, the Year string, and Year of Study
    const academicStatus = await calculateStudentStatus(
      student._id,
      (student.program as any)._id,
     academicYearName, // Use the dynamic variable
     student.currentYearOfStudy || 1
    );

    res.json({
      student: {
        name: student.name,
        regNo: student.regNo,
        program: (student.program as any)?.name,
        currentYear: student.currentYearOfStudy || 1, 
        currentSemester: student.currentSemester || 1,
      },
      grades: processedGrades, // Return the flattened version
      currentStatus: academicStatus?.status || "UNKNOWN",
      academicStatus: academicStatus,
      summary: academicStatus?.summary || {
        totalUnits: grades.length,
        passed: grades.filter((g) => g.status === "PASS").length,
        supplementary: grades.filter((g) => g.status === "SUPPLEMENTARY")
          .length,
        retake: grades.filter((g) => g.status === "RETAKE").length,
      },
    });
  })
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
      { new: true }
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
  })
);

// GET RAW MARKS FOR A STUDENT (FOR COORDINATOR USE)
router.get(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, academicYear, unitCode } = req.query;

    if (!regNo || typeof regNo !== "string") {
      return res.status(400).json({ error: "regNo required" });
    }

    let query: any = {
      student: (
        await Student.findOne({
          regNo: { $regex: `^${regNo}$`, $options: "i" },
        })
      )?._id,
    };

    if (academicYear && typeof academicYear === "string") {
      const year = await AcademicYear.findOne({ year: academicYear });
      if (year) query.academicYear = year._id;
    }

    if (unitCode && typeof unitCode === "string") {
      const unit = await Unit.findOne({ code: unitCode });
      if (unit) query.unit = unit._id;
    }

    const marks = await Mark.find(query)
      // .populate("unit", "code name")
      .populate({
        path: "programUnit",
        populate: {
          path: "unit",
          select: "code name",
        },
      })
      .populate("academicYear", "year")
      .select(
        "programUnit academicYear cat1Raw cat2Raw cat3Raw assgnt1Raw examQ1Raw examQ2Raw examQ3Raw examQ4Raw examQ5Raw caTotal30 examTotal70 agreedMark"
      )
      .lean();

    res.json(marks);
  })
);

// POST /raw-marks: Upload or Update Raw Marks for a Student
router.post(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const {
      regNo,
      unitCode,
      academicYear,
      cat1,
      cat2,
      cat3,
      assignment1,
      assignment2,
      assignment3,
      examQ1,
      examQ2,
      examQ3,
      examQ4,
      examQ5,
    } = req.body;

    if (!regNo || !unitCode || !academicYear) {
      return res
        .status(400)
        .json({ error: "regNo, unitCode, and academicYear are required" });
    }

    const student = await Student.findOne({
      regNo: { $regex: `^${regNo}$`, $options: "i" },
    });
    const unitDoc = await Unit.findOne({ code: unitCode });
    const yearDoc = await AcademicYear.findOne({ year: academicYear });

    if (!student || !unitDoc || !yearDoc) {
      return res
        .status(404)
        .json({ error: "Student, Unit, or Academic Year not found" });
    }

    const programUnit = await ProgramUnit.findOne({
      program: student.program,
      unit: unitDoc._id,
    });
    if (!programUnit) {
      return res
        .status(400)
        .json({ error: `Unit ${unitCode} is not linked to program.` });
    }

    // --- Math Calculation ---
    const cats = [cat1, cat2, cat3].filter(
      (v) => v !== undefined && v !== "" && v !== null
    );
    const catAvg =
      cats.length > 0
        ? cats.reduce((a, b) => Number(a) + Number(b), 0) / cats.length
        : 0;
    const assgns = [assignment1, assignment2, assignment3].filter(
      (v) => v !== undefined && v !== "" && v !== null
    );
    const assgnAvg =
      assgns.length > 0
        ? assgns.reduce((a, b) => Number(a) + Number(b), 0) / assgns.length
        : 0;

    const caTotal30 = Number((catAvg + assgnAvg).toFixed(2));
    const examTotal70 = [examQ1, examQ2, examQ3, examQ4, examQ5].reduce(
      (sum, q) => sum + (Number(q) || 0),
      0
    );
    const finalAgreed = Math.round(caTotal30 + examTotal70);

    // --- Update Database ---
    const mark = await Mark.findOneAndUpdate(
      {
        student: student._id,
        programUnit: programUnit._id,
        academicYear: yearDoc._id,
      },
      {
        cat1Raw: Number(cat1) || 0,
        cat2Raw: Number(cat2) || 0,
        cat3Raw: cat3 ? Number(cat3) : undefined,
        assgnt1Raw: Number(assignment1) || 0,
        examQ1Raw: Number(examQ1) || 0,
        examQ2Raw: Number(examQ2) || 0,
        examQ3Raw: Number(examQ3) || 0,
        examQ4Raw: Number(examQ4) || 0,
        examQ5Raw: Number(examQ5) || 0,
        caTotal30,
        examTotal70,
        agreedMark: finalAgreed,
        uploadedBy: req.user._id,
        uploadedAt: new Date(),
        attempt: "1st",
      },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    ).populate({
      path: "programUnit",
      populate: { path: "unit", select: "code name" },
    });

    // Recalculate Final Grade
    const gradeResult = await computeFinalGrade({
      markId: mark!._id as any,
      coordinatorReq: req,
    });
    console.log("[SUCCESS] Marks saved & grade recalculated:", {
      regNo,
      unit: unitCode,
      year: academicYear,
      finalGrade: gradeResult.grade,
      status: gradeResult.status,
    });

    res.json({
      success: true,
      data: {
        caTotal: caTotal30,
        examTotal: examTotal70,
        finalMark: finalAgreed,
        grade: gradeResult.grade,
      },
    });
  })
);



// Full Transcript Download
router.get(
  "/transcript",
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    // router.get("/transcript", requireAuth, asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
    let { regNo } = req.query;
    if (!regNo || typeof regNo !== "string") {
      return res.status(400).json({ error: "regNo required" });
    }
    regNo = decodeURIComponent(regNo);

    console.log("[TRANSCRIPT] Generating for:", regNo);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Pragma", "no-cache");

    await generateStudentTranscript(regNo, res);
  })
);

// Transcript for Specific Year
router.get(
  "/transcript/year",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { regNo, year } = req.query;
    if (
      !regNo ||
      !year ||
      typeof regNo !== "string" ||
      typeof year !== "string"
    )
      return res.status(400).json({ error: "regNo and year required" });

    regNo = decodeURIComponent(regNo);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Cache-Control", "no-cache");

    await generateStudentTranscript(regNo, res, year);
  })
);

export default router;
