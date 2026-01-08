// src/routes/studentSearch.ts
import express, { Response } from "express";
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
import { generateStudentTranscript } from "../services/pdfGenerator";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";

const router = express.Router();

// SEARCH STUDENT BY REG NO
router.get("/search", requireAuth, asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
const { q } = req.query;

  console.log("[STUDENT SEARCH] Query received:", q); // LOG

  if (!q || typeof q !== "string" || q.trim().length < 3) {
    return res.status(400).json({ error: "Enter at least 3 characters" });
  }

  const searchQuery = q.trim();

  const students = await Student.find({
    regNo: { $regex: `^${searchQuery}`, $options: "i" } // Case-insensitive + starts with
  })
    .limit(10)
    .select("regNo name program admissionYear")
    .populate("program", "name");

  console.log(`[STUDENT SEARCH] Found ${students.length} students for "${searchQuery}"`);

  res.json(students);
}));

// GET STUDENT FULL RESULTS + STATUS
router.get("/record", requireAuth, asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
  let { regNo } = req.query;
  console.log("[STUDENT RECORD] Requested regNo (raw):", regNo);

  if (!regNo || typeof regNo !== "string") {
    return res.status(400).json({ error: "regNo is required" });
  }

   regNo = decodeURIComponent(regNo);

  console.log("[STUDENT RECORD] Decoded regNo:", regNo);

   const student = await Student.findOne({
    regNo: { $regex: `^${regNo}$`, $options: "i" } // Exact match, case-insensitive
  }).populate("program");

  if (!student) {
    console.log("[STUDENT RECORD] Not found:", regNo);
    return res.status(404).json({ error: "Student not found" });
  }

  console.log("[STUDENT RECORD] Found:", student.name, student.regNo);

  const grades = await FinalGrade.find({ student: student._id })
    .populate<{
      unit: { code: string; name: string; creditHours?: number };
      academicYear: { year: string };  // ← THIS IS THE KEY
    }>("unit academicYear")
    .sort({ "academicYear.year": 1, "unit.code": 1 })
    .lean();

  // NOW THIS WORKS PERFECTLY — NO TYPE ERROR
  const currentAcademicYear = "2024/2025"; // Or get dynamically

  const failedThisYear = grades.filter(g =>
    (g.academicYear as any).year === currentAcademicYear 
    && (g.status === "SUPPLEMENTARY" || g.status === "RETAKE")
  );

  const hasRetake = failedThisYear.some(g => g.status === "RETAKE");

  const status = failedThisYear.length === 0
    ? "IN GOOD STANDING"
    : hasRetake
    ? "RETAKE YEAR"
    : "SUPPLEMENTARY PENDING";

  res.json({
    student: {
      name: student.name,
      regNo: student.regNo,
      program: (student.program as any)?.name,
    },
    grades,
    currentStatus: status,
    summary: {
      totalUnits: grades.length,
      passed: grades.filter(g => g.status === "PASS").length,
      supplementary: grades.filter(g => g.status === "SUPPLEMENTARY").length,
      retake: grades.filter(g => g.status === "RETAKE").length,
    }
  });
}));

// Full Transcript Download
router.get("/transcript", asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
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
}));

// Transcript for Specific Year
router.get("/transcript/year", requireAuth, asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
  let { regNo, year } = req.query;
  if (!regNo || !year || typeof regNo !== "string" || typeof year !== "string")
    return res.status(400).json({ error: "regNo and year required" });

  regNo = decodeURIComponent(regNo);

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Cache-Control", "no-cache");
  
  await generateStudentTranscript(regNo, res, year);
}));

// src/routes/studentSearch.ts → ADD THIS ROUTE

router.get("/raw-marks", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest , res:Response) => {
  const { regNo, academicYear, unitCode } = req.query;

  if (!regNo || typeof regNo !== "string") {
    return res.status(400).json({ error: "regNo required" });
  }

  let query: any = { 
    student: (await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } }))?._id 
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
    .populate("unit", "code name")
    .populate("academicYear", "year")
    .select("unit academicYear cat1 cat2 cat3 assignment practical exam isSupplementary")
    .lean();

  res.json(marks);
}));


router.post(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, unitCode, academicYear, cat1, cat2, cat3, assignment, practical, exam } = req.body;

    if (!regNo || !unitCode || !academicYear) {
      return res.status(400).json({ error: "regNo, unitCode, and academicYear required" });
    }

    // Find student, unit, year
    const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } });
    const unit = await Unit.findOne({ code: unitCode });
    const year = await AcademicYear.findOne({ year: academicYear });

    if (!student || !unit || !year) {
      return res.status(404).json({ error: "Student, Unit, or Academic Year not found" });
    }

    // Explicitly convert to ObjectId — THIS KILLS THE ERROR
    const studentId = student._id as unknown as mongoose.Types.ObjectId;
    const unitId = unit._id as unknown as mongoose.Types.ObjectId;
    const yearId = year._id as unknown as mongoose.Types.ObjectId;
    const uploadedById = req.user._id as unknown as mongoose.Types.ObjectId;

    // Save or update raw marks
    const mark = await Mark.findOneAndUpdate(
      {
        student: studentId,
        unit: unitId,
        academicYear: yearId,
      },
      {
        cat1,
        cat2,
        cat3,
        assignment,
        practical,
        exam,
        uploadedBy: uploadedById,
        uploadedAt: new Date(),
        isSupplementary: false,
      },
      { upsert: true, new: true, setDefaultsOnInsert: true }
    )
      .populate("unit", "code name")
      .populate("academicYear", "year");

    if (!mark) {
      return res.status(500).json({ error: "Failed to save marks" });
    }

    // RECALCULATE FINAL GRADE — NOW SAFE
    const result = await computeFinalGrade({
      markId: mark._id as unknown as mongoose.Types.ObjectId, // ← THIS IS THE KEY
      coordinatorReq: req,
    });

    console.log("[SUCCESS] Marks saved & grade recalculated:", {
      regNo,
      unit: unitCode,
      year: academicYear,
      finalGrade: result.grade,
      status: result.status,
    });

    res.json({
      success: true,
      message: "Marks saved and final grade recalculated",
      rawMark: mark,
      finalGrade: result,
    });
  })
);



export default router;