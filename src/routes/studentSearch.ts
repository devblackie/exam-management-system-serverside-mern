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
import { deferAdmission, grantAcademicLeave, revertStatusToActive } from "../services/academicLeave";
import { getYearWeight } from "../utils/weightingRegistry";

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

// // GET STUDENT FULL RESULTS + STATUS
// router.get(
//   "/record",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     let { regNo, yearOfStudy } = req.query;
//     if (!regNo || typeof regNo !== "string") return res.status(400).json({ error: "regNo is required" });

//     const targetYearOfStudy = parseInt(yearOfStudy as string) || 1;

//     regNo = decodeURIComponent(regNo);
//     const student = await Student.findOne( scopeQuery(req, { regNo: { $regex: `^${regNo}$`, $options: "i" }})).populate("program");

//     if (!student) return res.status(404).json({ error: "Student not found" });

//     // 1. Fetch grades with the correct nested path
//     const grades = await FinalGrade.find({ student: student._id })
//       .populate({ path: "programUnit", populate: { path: "unit", select: "code name" }})
//       .populate("academicYear", "year").sort({ "academicYear.year": 1 }).lean();

//     const academicYearName = (req.query.academicYear as string) || "2024/2025";

//     // 2. Safe mapping to prevent "undefined" errors
//     const processedGrades = grades
//       .filter((g) => {
//         const pUnit = g.programUnit as any;
//         // Only include units meant for the Year of Study selected by the user
//         return pUnit?.requiredYear === targetYearOfStudy;
//       })
//       .map((g) => {
//         const pUnit = g.programUnit as any;
//         return {
//           ...g,
//           unit: { code: pUnit?.unit?.code || "N/A", name: pUnit?.unit?.name || "Unknown Unit" },
//           semester: pUnit?.requiredSemester || "N/A",
//           academicYear: g.academicYear || { year: "N/A" },
//         };
//       });

//     // 2. USE THE SERVICE for Status
//     // We pass studentId, programId, the Year string, and Year of Study
//     const academicStatus = await calculateStudentStatus(
//       student._id,
//       (student.program as any)._id,
//       "", // academicYearName is left empty if status engine can derive it from yearOfStudy
//       targetYearOfStudy,
//     );

//     res.json({
//       student: {
//         _id: student._id, name: student.name, regNo: student.regNo,
//         programId: (student.program as any)?._id || student.program,
//         programName: (student.program as any)?.name,
//         currentYear: student.currentYearOfStudy || 1, // Student's actual level
//         currentSemester: student.currentSemester || 1, status: student.status,
//       },
//       grades: processedGrades, // Only Year X grades
//       viewingYearOfStudy: targetYearOfStudy,
//       academicStatus: academicStatus,
//       summary: academicStatus?.summary || {
//         totalUnits: processedGrades.length,
//         passed: processedGrades.filter((g) => g.status === "PASS").length,
//         failed: processedGrades.filter((g) =>
//           ["RETAKE", "SUPPLEMENTARY", "FAIL"].includes(g.status),
//         ).length,
//       },
//     });
//   }),
// );

// GET STUDENT FULL RESULTS + STATUS
router.get(
  "/record",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    let { regNo, yearOfStudy } = req.query;
    if (!regNo || typeof regNo !== "string") return res.status(400).json({ error: "regNo is required" });

    const targetYearOfStudy = parseInt(yearOfStudy as string) || 1;
    regNo = decodeURIComponent(regNo);
    const student = await Student.findOne( scopeQuery(req, { regNo: { $regex: `^${regNo}$`, $options: "i" }})).populate("program");

    if (!student) return res.status(404).json({ error: "Student not found" });

    // 1. Fetch grades
    const grades = await FinalGrade.find({ student: student._id })
      .populate({ path: "programUnit", populate: { path: "unit", select: "code name" }})
      .populate("academicYear", "year").sort({ "academicYear.year": 1 }).lean();

    const marks = await Mark.find({ student: student._id }).lean();
    const marksMap = new Map();
    marks.forEach(m => marksMap.set(m.programUnit.toString(), m));

    // 2. Safe mapping and status merging
    const processedGrades = grades
      .filter((g) => {
        const pUnit = g.programUnit as any;
        return pUnit?.requiredYear === targetYearOfStudy;
      })
      .map((g) => {
        const pUnit = g.programUnit as any;
        const programUnitId = pUnit._id.toString();

        // Merge special status from raw marks
        const rawMarkRecord = marksMap.get(programUnitId);

        // --- FIX: Logic to determine status ---
        let finalStatus = g.status;
        const isSpecial =
          rawMarkRecord?.isSpecial || g.isSpecial || g.status === "SPECIAL";

        if (isSpecial) {
          finalStatus = "SPECIAL";
        } else if (
          rawMarkRecord &&
          (!rawMarkRecord.caTotal30 || !rawMarkRecord.examTotal70)
        ) {
          // If raw marks exist but are incomplete (CAT/Exam missing)
          finalStatus = "INCOMPLETE";
        } else if (g.totalMark === 0 && g.status !== "PASS") {
          finalStatus = "INCOMPLETE";
        }
        // ----------------------------------------
        return {
          ...g,
          unit: {
            code: pUnit?.unit?.code || "N/A",
            name: pUnit?.unit?.name || "Unknown Unit",
          },
          semester: pUnit?.requiredSemester || "N/A",
          academicYear: g.academicYear || { year: "N/A" },
          // Override status if special
          status: finalStatus,
        };
      });

    // 3. Calculate Status
    const academicStatus = await calculateStudentStatus(
      student._id,
      (student.program as any)._id,
      "",
      targetYearOfStudy,
    );

    res.json({
      student: {
        _id: student._id, name: student.name, regNo: student.regNo,
        programId: (student.program as any)?._id || student.program,
        programName: (student.program as any)?.name,
        currentYear: student.currentYearOfStudy || 1,
        currentSemester: student.currentSemester || 1, status: student.status,
      },
      grades: processedGrades,
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

// Helper to keep code clean
const formatChallenges = (analysis: any) => ({
supplementary: analysis.failedList.map((f: any) => f.displayName.split(":")[0]),
retakes: analysis.status === "REPEAT YEAR" ? analysis.failedList.map((f: any) => f.displayName.split(":")[0]) : [],
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

  const journey: any[] = [];
  const program = student.program as any;
  const entryType = student.entryType || "Direct";
  let cumulativeWeightApplied = 0;
  let weightedSum = 0;

  const addWeight = (year: number) => getYearWeight(program, entryType, year);

  // --- PART A: History ---
  for (const record of student.academicHistory || []) {
    const analysis = await calculateStudentStatus(student._id, student.program, record.academicYear, record.yearOfStudy);
    const weight = addWeight(record.yearOfStudy);
    
    journey.push({
      type: "ACADEMIC",
      academicYear: record.academicYear,
      yearOfStudy: record.yearOfStudy,
      status: analysis.status,
      totalUnits: analysis.summary.totalExpected,
      weight: Math.round(weight * 100),
      challenges: formatChallenges(analysis),
      date: record.date || new Date(0),
    });

    weightedSum += (parseFloat(analysis.weightedMean) * weight);
    cumulativeWeightApplied += weight;
  }

  // --- PART B: Current Year ---
  if (!student.academicHistory?.some(h => h.yearOfStudy === student.currentYearOfStudy)) {
    const live = await calculateStudentStatus(student._id, student.program, "CURRENT", student.currentYearOfStudy);
    const weight = addWeight(student.currentYearOfStudy);

    journey.push({
      type: "ACADEMIC",
      academicYear: "2026/2027",
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
    currentStatus: student.status.toUpperCase(),
    cumulativeMean: projectedMean.toFixed(2),
    timeline: journey,
  });
}));

// router.get("/journey", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//   const { regNo } = req.query;
//   if (!regNo) return res.status(400).json({ error: "regNo is required" });

//   const student = await Student.findOne(
//     scopeQuery(req, { regNo: regNo as string }),
//   )
//     .populate("admissionAcademicYear", "year")
//     .populate("program")
//     .lean();

//   if (!student) return res.status(404).json({ error: "Student not found" });

//   const journey: any[] = [];
//   const program = student.program as any;
//   const entryType = student.entryType || "Direct";

//   // --- Helper to add Weighting ---
//   const addWeight = (year: number) => {
//     // getYearWeight returns 0.15, 0.20 etc. We convert to percentage
//     return Math.round(getYearWeight(program, entryType, year) * 100);
//   };

//   // --- PART A: Process Completed Years (The History) ---
//   const historicalYears = student.academicHistory || [];
//   for (const record of historicalYears) {
//     const analysis = await calculateStudentStatus(
//       student._id,
//       student.program,
//       record.academicYear,
//       record.yearOfStudy,
//     );
//     journey.push({
//       type: "ACADEMIC",
//       academicYear: record.academicYear,
//       yearOfStudy: record.yearOfStudy,
//       status: analysis.status,
//       totalUnits: analysis.summary.totalExpected,
//       weight: addWeight(record.yearOfStudy),
//       challenges: formatChallenges(analysis),
//       date: record.date || new Date(0),
//     });
//   }

//   // --- PART B: Process Current Year (The Live Snapshot) ---
//   // Check if the current year is already in history to avoid duplicates
//   const isCurrentYearInHistory = historicalYears.some(
//     (h) => h.yearOfStudy === student.currentYearOfStudy,
//   );

//   if (!isCurrentYearInHistory) {
//     const liveAnalysis = await calculateStudentStatus(
//       student._id,
//       student.program,
//       "CURRENT",
//       student.currentYearOfStudy,
//     );
//     journey.push({
//       type: "ACADEMIC",
//       academicYear: "2026/2027", // Or dynamically get current session
//       yearOfStudy: student.currentYearOfStudy,
//       status: liveAnalysis.status,
//       totalUnits: liveAnalysis.summary.totalExpected,
//       weight: addWeight(student.currentYearOfStudy),
//       challenges: formatChallenges(liveAnalysis),
//       date: new Date(), // Now
//       isCurrent: true,
//     });
//   }

//   // --- PART C: Administrative Events ---
//   student.statusEvents?.forEach((event) => {
//     journey.push({
//       type: "STATUS_CHANGE",
//       academicYear: event.academicYear,
//       toStatus: event.toStatus,
//       reason: event.reason,
//       date: event.date,
//     });
//   });

//   journey.sort(
//     (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime(),
//   );

//   res.json({
//     admissionYear: (student.admissionAcademicYear as any)?.year,
//     currentStatus: student.status.toUpperCase(),
//     timeline: journey,
//   });
// }));




// POST /approve-special: One-click approval for Special Exams
router.post(
  "/approve-special",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { markId, reason, undo = false } = req.body;

    if (!markId) return res.status(400).json({ error: "Mark ID is required" });
    

    let updateData;
    if (undo) {
      // Revert to normal state
      updateData = {
        $set: { isSpecial: false, attempt: "1st", remarks: "Special Exam Revoked" },
      };
    } else {
      // Set to Special state
      // Validate reason for granting
      const finalReason = ["Financial", "Compassionate"].includes(reason)
        ? reason
        : "Administrative";
        
      updateData = {
        $set: { isSpecial: true, attempt: "special", remarks: `Special Granted: ${finalReason}` },
      };
    }

    // 1. Find the mark and update to Special status
    const mark = await Mark.findByIdAndUpdate(
      markId, updateData, { new: true },
    );

    if (!mark) {
      return res.status(404).json({ error: "Mark record not found" });
    }

    // 2. Trigger Grade Recalculation
    // This will update FinalGrade status to 'SPECIAL' and Grade to 'I'
    const gradeResult = await computeFinalGrade({ markId: mark._id as any, coordinatorReq: req,   });

    res.json({ success: true, message: undo ? "Special exam revoked" : "Special exam approved", newStatus: gradeResult.status, });
  }),
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

// GET RAW MARKS FOR A STUDENT (FOR COORDINATOR USE)
router.get(
  "/raw-marks",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo, yearOfStudy } = req.query;

    if (!regNo) return res.status(400).json({ error: "regNo required" });

    const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" }});

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
    const { regNo, unitCode, academicYear, semester, cat1, cat2, cat3, assignment1, practicalRaw, examQ1, examQ2, examQ3, examQ4, examQ5, isSpecial, examMode } = req.body;

    if (!regNo || !unitCode || !academicYear) return res.status(400).json({ error: "regNo, unitCode, and academicYear are required" });    

    console.log(`[LOG] Incoming Update -> Student: ${regNo}, Unit: ${unitCode}, Year: ${academicYear}`);

    const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" },});
    const unitDoc = await Unit.findOne({ code: unitCode });
    const yearDoc = await AcademicYear.findOne({ year: academicYear });

    if (!student || !unitDoc || !yearDoc) return res.status(404).json({ error: "Student, Unit, or Academic Year not found" });

    if (!student || !unitDoc || !yearDoc) {
      console.error(`[ERROR] Missing Metadata: Student(${!!student}), Unit(${!!unitDoc}), Year(${!!yearDoc})`);
      return res.status(404).json({ error: "Required metadata (Student/Unit/Year) not found" });
    }
    


    const programUnit = await ProgramUnit.findOne({ program: student.program, unit: unitDoc._id });
    if (!programUnit) return res.status(400).json({ error: `Unit ${unitCode} is not linked to program.` });
    
    // let isSpecial = false;

    
    let existingMark = await Mark.findOne({ student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id });

    let detectedAttempt = "1st";
    if (!existingMark) {
      // Only check previous years if this is a BRAND NEW mark entry
      const totalPastAttempts = await Mark.countDocuments({ student: student._id, programUnit: programUnit._id });
      if (totalPastAttempts === 1) detectedAttempt = "supplementary";
      else if (totalPastAttempts >= 2) detectedAttempt = "re-take";
    } else {
      // If mark exists for this year, keep its current attempt status!
      detectedAttempt = existingMark.attempt;
    }

  const updateData = {
    institution: student.institution, student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id, semester: semester || "SEMESTER 1",
    cat1Raw: Number(cat1) || 0, cat2Raw: Number(cat2) || 0, cat3Raw: cat3 ? Number(cat3) : undefined, assgnt1Raw: Number(assignment1) || 0, practicalRaw: Number(practicalRaw) || 0,
    examQ1Raw: Number(examQ1) || 0, examQ2Raw: Number(examQ2) || 0, examQ3Raw: Number(examQ3) || 0, examQ4Raw: Number(examQ4) || 0, examQ5Raw: Number(examQ5) || 0,
    examMode: examMode || "standard", isSpecial: isSpecial === true || isSpecial === "true", attempt: isSpecial ? "special" : detectedAttempt,
    // isSupplementary: detectedAttempt === "supplementary",
    // isRetake: detectedAttempt.includes("re-take"),
    uploadedBy: req.user._id, uploadedAt: new Date(),
    caTotal30: existingMark?.caTotal30 || 0, examTotal70: existingMark?.examTotal70 || 0, agreedMark: existingMark?.agreedMark || 0,
  };

    // const mark = await Mark.findOneAndUpdate(
    const updatedMark = await Mark.findOneAndUpdate(
      {
        student: student._id,
        programUnit: programUnit._id,
        academicYear: yearDoc._id,
      },
     updateData,
      { upsert: true, new: true, setDefaultsOnInsert: true },
    ).populate({
      path: "programUnit",
      populate: { path: "unit", select: "code name" },
    });

    console.log(`[LOG] Mark Document ${updatedMark._id} updated. Triggering Grade Calc...`);

    // Recalculate Final Grade
    const gradeResult = await computeFinalGrade({ markId: updatedMark._id as any, coordinatorReq: req  });

    console.log(`[SUCCESS] Grade Processed: ${gradeResult.finalMark} (${gradeResult.grade}) - Status: ${gradeResult.status}`);
    res.json({ success: true, message: "Marks updated and grades recalculated", data: gradeResult });
  }),
);




export default router;
