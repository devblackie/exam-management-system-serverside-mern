// src/routes/studentSearch.ts
import express, { Response } from "express";
import mongoose from "mongoose";
import Student from "../models/Student";
import FinalGrade, { IFinalGrade } from "../models/FinalGrade";
import { AuthenticatedRequest,requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import Mark from "../models/Mark";
import Unit from "../models/Unit";
import AcademicYear from "../models/AcademicYear";
import ProgramUnit from "../models/ProgramUnit";
import { computeFinalGrade } from "../services/gradeCalculator";
import { calculateStudentStatus } from "../services/statusEngine";
import { scopeQuery } from "../lib/multiTenant";
import { deferAdmission, grantAcademicLeave, readmitStudent, revertStatusToActive } from "../services/academicLeave";
import { getYearWeight } from "../utils/weightingRegistry";
import MarkDirect from "../models/MarkDirect";
import InstitutionSettings from "../models/InstitutionSettings";
import { resolveStudentStatus } from "../utils/studentStatusResolver";

const router = express.Router();

// SEARCH STUDENT BY REG NO
router.get(
  "/search",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { q } = req.query;

    // console.log("[STUDENT SEARCH] Query received:", q);

    if (!q || typeof q !== "string" || q.trim().length < 3) return res.status(400).json({ error: "Enter at least 3 characters" });
    
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
    if (!regNo || typeof regNo !== "string") return res.status(400).json({ error: "regNo is required" });

    const targetYearOfStudy = parseInt(yearOfStudy as string) || 1;
    regNo = decodeURIComponent(regNo);

    const [student, settings] = await Promise.all([
      Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" }}).populate("program"),
      InstitutionSettings.findOne().lean(),
    ]);

    if (!student) return res.status(404).json({ error: "Student not found" });
    if (!settings) return res.status(500).json({ error: "Institution settings missing" });

    const [grades, detailedMarks, directMarks] = await Promise.all([
      FinalGrade.find({ student: student._id })
        .populate({ path: "programUnit", populate: { path: "unit", select: "code name" }})
        .populate("academicYear", "year")
        .lean(),
      Mark.find({ student: student._id }).lean(),
      MarkDirect.find({ student: student._id })
        .populate({ path: "programUnit", populate: { path: "unit", select: "code name" }})
        .populate("academicYear", "year")
        .lean(),
    ]);

    const marksMap = new Map();
    detailedMarks.forEach((m) => marksMap.set(m.programUnit.toString(), m));
    const gradedUnitIds = new Set(
      grades.map((g) => (g.programUnit as any)?._id.toString()),
    );

    // 1. Process Official Grades
    const processedGrades: any[] = grades
      .filter((g: any) => g.programUnit?.requiredYear === targetYearOfStudy)
      .map((g: any) => {
        const pUnit = g.programUnit;
        const rawMarkRecord = marksMap.get(pUnit._id.toString());
        let finalStatus = g.status;
        if (rawMarkRecord?.isSpecial || g.status === "SPECIAL")
          finalStatus = "SPECIAL";

        return {
          ...g,
          unit: { code: pUnit?.unit?.code, name: pUnit?.unit?.name },
          semester: pUnit?.requiredSemester || "N/A",
          status: finalStatus,
          agreedMark: g.totalMark, // Backend alias for frontend table
        };
      });
    
    // 2. Inject Direct Marks
    directMarks.forEach((dm: any) => {
      const pUnit = dm.programUnit;
      if ( pUnit?.requiredYear === targetYearOfStudy && !gradedUnitIds.has(pUnit._id.toString())) {
        const mark = dm.agreedMark || 0;
        const matchedScale = [...(settings.gradingScale || [])].sort((a, b) => b.min - a.min).find((s) => mark >= s.min);

        // Normalize semester to string to prevent .localeCompare error
        const semesterStr = String( dm.semester || pUnit.requiredSemester || "N/A" );

        processedGrades.push({
          _id: dm._id,
          academicYear: dm.academicYear, // Keep as object for frontend display
          semester: semesterStr,
          unit: {
            code: pUnit.unit?.code || "N/A",
            name: pUnit.unit?.name || "N/A",
          },
          totalMark: mark,
          agreedMark: mark,
          grade: matchedScale ? matchedScale.grade : "E",
          // status: dm.attempt?.toUpperCase() === "SPECIAL" ? "SPECIAL" : mark >= settings.passMark ? "PASS" : "FAIL",
          status: (() => {
            if (dm.attempt === "special") return "SPECIAL";
            if (dm.attempt === "supplementary") return "SUPPLEMENTARY";
            if (dm.attempt === "re-take") return "RETAKE";
            return mark >= settings.passMark ? "PASS" : "SUPPLEMENTARY";
          })(),
        });
      }
    });

    // 3. Robust Sorting (Prevents localeCompare crash)
    const sortedGrades = processedGrades.sort((a, b) => {
      const semA = String(a.semester || "");
      const semB = String(b.semester || "");
      return semA.localeCompare(semB);
    });

    const LOCKED_STATUSES = ["on_leave", "deferred", "discontinued", "deregistered", "graduated"];

    if (LOCKED_STATUSES.includes(student.status)) {
        const resolved = resolveStudentStatus(student);

        const variantMap: Record<string, "info" | "warning" | "error" | "success"> = {
          on_leave:      "info", deferred:      "info", discontinued:  "error", deregistered:  "error", graduated:     "success",
        };

        return res.json({
          student: {
            _id: student._id, name: student.name, regNo: student.regNo, programId: (student.program as any)?._id, 
            currentYear: student.currentYearOfStudy, programName: (student.program as any)?.name, status: student.status,
        },
        grades: sortedGrades,
        academicStatus: {
          status: resolved.status, variant: variantMap[student.status] ?? "info",
          details: resolved.reason || `Student is currently ${resolved.status}.`,
          weightedMean: "0.00",
          summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0, isOnLeave: true },
          passedList:     [], failedList:     [], specialList:    [],
          missingList:    [], incompleteList: [], academicYearName: "",
        },
      });
    }

    const academicStatus = await calculateStudentStatus( student._id, (student.program as any)._id, "", targetYearOfStudy );     

    res.json({
      student: {
        _id: student._id, name: student.name, regNo: student.regNo, programId: (student.program as any)?._id,
        currentYear: student.currentYearOfStudy, programName: (student.program as any)?.name, status: student.status,
      },
      grades: sortedGrades, academicStatus,
    });
  })
);

// --- UPDATED HELPER ---
const formatChallenges = (analysis: any) => ({
  supplementary: analysis.failedList.filter((f: any) => f.attempt === "A/S").map((f: any) => f.displayName.split(":")[0]),
  retakes: analysis.failedList.filter((f: any) => f.attempt === "A/RA" || f.attempt === "A/RPU").map((f: any) => f.displayName.split(":")[0]),
  stayouts: analysis.failedList.filter((f: any) => f.attempt === "A/SO").map((f: any) => f.displayName.split(":")[0]),
  specials: analysis.specialList.map((s: any) => s.displayName.split(":")[0]),
  incomplete: [...analysis.incompleteList, ...analysis.missingList].map(i => i.split(":")[0])
});

// GET /journey: Fetch complete academic timeline with Cumulative Tracker
router.get("/journey", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { regNo } = req.query;
  if (!regNo) return res.status(400).json({ error: "regNo is required" });

  const student = await Student.findOne(scopeQuery(req, { regNo: regNo as string }))
    .populate("admissionAcademicYear", "year")
    .populate("program")
    .lean();

  if (!student) return res.status(404).json({ error: "Student not found" });

  const admissionYearString = (student.admissionAcademicYear as any)?.year || "N/A";
  const journey: any[] = [];
  const program = student.program as any;
  const entryType = student.entryType || "Direct";
  let cumulativeWeightApplied = 0;
  let weightedSum = 0;

  const incrementYear = (yearStr: string, offset: number) => {
    if (!yearStr || !yearStr.includes("/")) return yearStr;
    const [start] = yearStr.split("/").map(Number);
    const newStart = start + offset;
    return `${newStart}/${newStart + 1}`;
  };

  const addWeight = (year: number) => getYearWeight(program, entryType, year);
  // console.log(`[JourneyAudit] Student: ${student.regNo}`);
  // console.log(`[JourneyAudit] Admission Year: ${admissionYearString}`);
  // --- PART A: History ---
  for (const record of student.academicHistory || []) {
    // console.log(
    //   `[HistoryItem] YOS: ${record.yearOfStudy} | DB_Year: ${record.academicYear}`,
    // );
    const analysis = await calculateStudentStatus( student._id, student.program, record.academicYear, record.yearOfStudy );
    const weight = addWeight(record.yearOfStudy);

    // LOGIC: If DB has no year, or it's incorrectly repeating the admission year for Y2+
    let resolvedYear = record.academicYear;
    if (
      !resolvedYear ||
      (record.yearOfStudy > 1 && resolvedYear === admissionYearString)
    ) {
      resolvedYear = incrementYear(admissionYearString, record.yearOfStudy - 1);
    }

    let formalOutcome = analysis.status;
    if (analysis.status === "PASS") formalOutcome = "PASS (PROMOTED)";
    else if (analysis.status === "SUPP") formalOutcome = "PROCEEDED (WITH SUPPLEMENTARIES)";

    journey.push({
      type: "ACADEMIC",
      // academicYear: record.academicYear,
      academicYear: resolvedYear,
      yearOfStudy: record.yearOfStudy,
      // status: analysis.status,
      status: formalOutcome,
      // status: analysis.status === "PASS" ? "PASS (PROMOTED)" : analysis.status,
      totalUnits: analysis.summary.totalExpected,
      weight: Math.round(weight * 100),
      challenges: formatChallenges(analysis),
      date: record.date || new Date(0),
      isCurrent: false,
    });

    weightedSum += parseFloat(analysis.weightedMean) * weight;
    cumulativeWeightApplied += weight;
  }

  // --- PART B: Current Year ---
  const LOCKED_FROM_CURRENT_YEAR = ["on_leave", "deferred", "deregistered", "discontinued", "graduated"];
  
  const isLocked = LOCKED_FROM_CURRENT_YEAR.includes(student.status);
  if ( !isLocked && !student.academicHistory?.some(h => h.yearOfStudy === student.currentYearOfStudy)) {
    console.log(`[CurrentYear] YOS: ${student.currentYearOfStudy} | Looking for current AcademicYear doc...`);
    const currentYearDoc = await AcademicYear.findOne({ institution: student.institution, isCurrent: true }).lean();
    const live = await calculateStudentStatus(student._id, student.program, "CURRENT", student.currentYearOfStudy);
    const weight = addWeight(student.currentYearOfStudy);
   

    journey.push({
      type: "ACADEMIC",
      academicYear: currentYearDoc?.year || "2026/2027",
      yearOfStudy: student.currentYearOfStudy,
      status: live.status,
      totalUnits: live.summary.totalExpected,
      weight: Math.round(weight * 100),
      challenges: formatChallenges(live),
      date: new Date(),
      isCurrent: true,
    });

    weightedSum += (parseFloat(live.weightedMean) * weight);
    cumulativeWeightApplied += weight;
  }

  // --- PART C: Status Events ---
  student.statusEvents?.forEach(event => {
    journey.push({ type: "STATUS_CHANGE", academicYear: event.academicYear, toStatus: event.toStatus, reason: event.reason, date: event.date });
  });

  journey.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

  // Calculate Projected Mean (normalized to 100% of weight encountered so far)
  const projectedMean = cumulativeWeightApplied > 0 ? (weightedSum / cumulativeWeightApplied) : 0;

  res.json({
    admissionYear: (student.admissionAcademicYear as any)?.year,
    intake: student.intake || "SEPT",
    currentStatus: student.status.toUpperCase(),
    cumulativeMean: projectedMean.toFixed(2),
    timeline: journey,
  });
}));

// POST /approve-special: One-click approval for Special Exams
router.post(
  "/approve-special",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { markId, reason, undo = false } = req.body;
    
    if (!markId) return res.status(400).json({ error: "Mark ID is required" });

    const updateData = undo
      ? { $set: { isSpecial: false, attempt: "1st", remarks: "Special Exam Revoked" } }
      : { $set: { isSpecial: true, attempt: "special", remarks: `Special Granted: ${reason || "Administrative"}` } };

    // Try updating in Detailed Marks first
    let mark = await Mark.findByIdAndUpdate(markId, updateData, { new: true });

    // If not found, try updating in Direct Marks
    if (!mark) mark = await MarkDirect.findByIdAndUpdate(markId, updateData, { new: true });  

    if (!mark) return res.status(404).json({ error: "Mark record not found in DB" });    

    try {
      const gradeResult = await computeFinalGrade({ markId: mark._id as any, coordinatorReq: req });
      res.json({ success: true, message: undo ? "Special exam revoked" : "Special exam approved", newStatus: gradeResult.status });
    } catch (calcError: any) {
      console.error("[SpecialExam] Calculation Error:", calcError.message);
      res.status(500).json({ error: calcError.message || "Grade calculation failed." });
    }
  })
);

// POST /student/leave
router.post(
  "/leave",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId, startDate, endDate, reason, leaveType } = req.body;
    const result = await grantAcademicLeave( studentId, new Date(startDate), new Date(endDate), reason, leaveType );
    res.json({ success: true, student: result });
  })
);

// POST /student/defer
router.post(
  "/defer",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId, years } = req.body;
    const result = await deferAdmission(studentId, parseInt(years) as 1 | 2);
    res.json({ success: true, student: result });
  })
);

// POST /student/revert-active
router.post(
  "/revert-active",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId } = req.body;
    const result = await revertStatusToActive(studentId);
    res.json({ success: true, student: result });
  }),
);

// POST /student/readmit: Handle formal readmission for deregistered/discontinued students
router.post(
  "/readmit",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId, remarks } = req.body;
    
    if (!studentId) return res.status(400).json({ error: "Student ID required" });

    const result = await readmitStudent(studentId, remarks || "Approved by Senate/Department");
    
    // Log the administrative action
    console.log(`[READMISSION] ${result?.regNo} reinstated by ${req.user?.name}`);

    res.json({ success: true, message: "Student readmitted successfully", student: result });
  })
);


// GET RAW MARKS FOR A STUDENT
router.get(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, yearOfStudy } = req.query;
    if (!regNo) return res.status(400).json({ error: "regNo required" });

    const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" }});
    if (!student) return res.status(404).json({ error: "Student not found" });

    // FETCH FROM BOTH COLLECTIONS
    const [detailedMarks, directMarks] = await Promise.all([
      Mark.find({ student: student._id })
        .populate({
          path: "programUnit",
          match: yearOfStudy ? { requiredYear: parseInt(yearOfStudy as string) } : {},
          populate: { path: "unit", select: "code name" },
        })
        .populate("academicYear", "year").lean(),
      MarkDirect.find({ student: student._id })
        .populate({
          path: "programUnit",
          match: yearOfStudy ? { requiredYear: parseInt(yearOfStudy as string) } : {},
          populate: { path: "unit", select: "code name" },
        })
        .populate("academicYear", "year").lean()
    ]);

    // Merge and flag sources so the frontend knows how to display them
    const combined = [
      ...detailedMarks.filter(m => m.programUnit).map(m => ({ ...m, entryMode: 'detailed' })),
      ...directMarks.filter(m => m.programUnit).map(m => ({ ...m, entryMode: 'direct' }))
    ];

    res.json(combined);
  }),
);

// POST /raw-marks: Handles both Detailed and Direct mark entries

router.post(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, unitCode, academicYear, semester, cat1, cat2, cat3, assignment1, practicalRaw, examQ1, examQ2, examQ3, examQ4, examQ5, examMode, caDirect, examDirect, agreedMark, isSpecial, attempt } = req.body;
    if (!regNo || !unitCode || !academicYear) return res.status(400).json({ error: "regNo, unitCode, and academicYear are required" });

    // 1. Resolve Metadata
    const [student, unitDoc, yearDoc] = await Promise.all([
      Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } }),
      Unit.findOne({ code: unitCode.toUpperCase().trim() }),
      AcademicYear.findOne({ year: academicYear })
    ]);

    if (!student || !unitDoc || !yearDoc) return res.status(404).json({ error: "Required metadata (Student/Unit/Year) not found" });
    const programUnit = await ProgramUnit.findOne({ program: student.program, unit: unitDoc._id });
    if (!programUnit) return res.status(400).json({ error: `Unit ${unitCode} is not linked to student's program.` });
    const isDirectEntry = caDirect !== undefined || examDirect !== undefined;

    // 2. Logic Fork: Direct vs Detailed
    if (isDirectEntry) {
      // console.log(`[LOG] Processing DIRECT Mark for ${regNo} - ${unitCode}`);
      
      const directUpdate = {
        institution: student.institution, student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id,
        semester: semester || "SEMESTER 1", caTotal30: Number(caDirect) || 0, examTotal70: Number(examDirect) || 0,
        agreedMark: Number(agreedMark) || (Number(caDirect) + Number(examDirect)),
        attempt: isSpecial ? "special" : (attempt || "1st"),
        isSpecial: isSpecial === true, uploadedBy: req.user._id, uploadedAt: new Date(),
      };

      const updatedDirect = await MarkDirect.findOneAndUpdate(
        { student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id },
        directUpdate,
        { upsert: true, new: true, setDefaultsOnInsert: true }
      );

      return res.json({ success: true, message: "Direct mark saved successfully", data: updatedDirect });

    } else {
      // 3. Detailed Mark Logic
      // console.log(`[LOG] Processing DETAILED Mark for ${regNo} - ${unitCode}`);
      
      const detailedUpdate = {
        institution: student.institution, student: student._id, programUnit: programUnit._id,
        academicYear: yearDoc._id, semester: semester || "SEMESTER 1", cat1Raw: Number(cat1) || 0,
        cat2Raw: Number(cat2) || 0, cat3Raw: cat3 ? Number(cat3) : undefined, assgnt1Raw: Number(assignment1) || 0,
        practicalRaw: Number(practicalRaw) || 0, examQ1Raw: Number(examQ1) || 0, examQ2Raw: Number(examQ2) || 0, 
        examQ3Raw: Number(examQ3) || 0, examQ4Raw: Number(examQ4) || 0, examQ5Raw: Number(examQ5) || 0,
        examMode: examMode || "standard", isSpecial: isSpecial === true,
        attempt: isSpecial ? "special" : (attempt || "1st"), uploadedBy: req.user._id, uploadedAt: new Date(),
      };

      const updatedMark = await Mark.findOneAndUpdate(
        { student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id },
        detailedUpdate,
        { upsert: true, new: true, setDefaultsOnInsert: true }
      );

      // Trigger the complex grading engine for detailed marks
      const gradeResult = await computeFinalGrade({ markId: updatedMark._id as any, coordinatorReq: req });
      return res.json({ success: true, message: "Detailed marks processed", data: gradeResult });
    }
  }),
);

export default router;
