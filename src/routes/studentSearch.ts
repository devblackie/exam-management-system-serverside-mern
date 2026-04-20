// serverside/src/routes/studentSearch.ts

import express, { Response } from "express";
import Student from "../models/Student";
import FinalGrade from "../models/FinalGrade";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
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

router.get("/search", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { q } = req.query;
  if (!q || typeof q !== "string" || q.trim().length < 3)
    return res.status(400).json({ error: "Enter at least 3 characters" });
  const students = await Student.find(scopeQuery(req, { regNo: { $regex: `^${q.trim()}`, $options: "i" } }))
    .limit(10).select("regNo name program admissionYear").populate("program", "name");
  res.json(students);
}));

router.get("/record", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  let { regNo, yearOfStudy } = req.query;
  if (!regNo || typeof regNo !== "string") return res.status(400).json({ error: "regNo is required" });
  const targetYearOfStudy = parseInt(yearOfStudy as string) || 1;
  regNo = decodeURIComponent(regNo);

  const [student, settings] = await Promise.all([
    Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } }).populate("program"),
    InstitutionSettings.findOne().lean(),
  ]);
  if (!student) return res.status(404).json({ error: "Student not found" });
  if (!settings) return res.status(500).json({ error: "Institution settings missing" });

  const [grades, detailedMarks, directMarks] = await Promise.all([
    FinalGrade.find({ student: student._id }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" } }).populate("academicYear", "year").lean(),
    Mark.find({ student: student._id }).lean(),
    MarkDirect.find({ student: student._id }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" } }).populate("academicYear", "year").lean(),
  ]);

  const marksMap = new Map();
  detailedMarks.forEach((m) => marksMap.set(m.programUnit.toString(), m));
  const gradedUnitIds = new Set(grades.map((g) => (g.programUnit as any)?._id.toString()));

  const processedGrades: any[] = grades
    .filter((g: any) => g.programUnit?.requiredYear === targetYearOfStudy)
    .map((g: any) => {
      const pUnit = g.programUnit;
      const raw = marksMap.get(pUnit._id.toString());
      return { ...g, unit: { code: pUnit?.unit?.code, name: pUnit?.unit?.name }, semester: pUnit?.requiredSemester || "N/A", status: (raw?.isSpecial || g.status === "SPECIAL") ? "SPECIAL" : g.status, agreedMark: g.totalMark };
    });

  directMarks.forEach((dm: any) => {
    const pUnit = dm.programUnit;
    if (pUnit?.requiredYear === targetYearOfStudy && !gradedUnitIds.has(pUnit._id.toString())) {
      const mark = dm.agreedMark || 0;
      const matchedScale = [...((settings as any).gradingScale || [])].sort((a, b) => b.min - a.min).find((s) => mark >= s.min);
      processedGrades.push({
        _id: dm._id, academicYear: dm.academicYear, semester: String(dm.semester || pUnit.requiredSemester || "N/A"),
        unit: { code: pUnit.unit?.code || "N/A", name: pUnit.unit?.name || "N/A" },
        totalMark: mark, agreedMark: mark, grade: matchedScale ? matchedScale.grade : "E",
        status: (() => {
          if (dm.attempt === "special") return "SPECIAL";
          if (dm.attempt === "supplementary") return "SUPPLEMENTARY";
          if (dm.attempt === "re-take") return "RETAKE";
          return mark >= (settings as any).passMark ? "PASS" : "SUPPLEMENTARY";
        })(),
      });
    }
  });

  const sortedGrades = processedGrades.sort((a, b) => String(a.semester || "").localeCompare(String(b.semester || "")));
  const LOCKED = ["on_leave","deferred","discontinued","deregistered","graduated","graduand"];

  if (LOCKED.includes(student.status)) {
    const resolved = resolveStudentStatus(student);
    const vm: Record<string, "info"|"warning"|"error"|"success"> = { on_leave:"info", deferred:"info", discontinued:"error", deregistered:"error", graduated:"success", graduand:"success" };
    return res.json({
      student: { _id: student._id, name: student.name, regNo: student.regNo, programId: (student.program as any)?._id, currentYear: student.currentYearOfStudy, programName: (student.program as any)?.name, status: student.status },
      grades: sortedGrades,
      academicStatus: { status: resolved.status, variant: vm[student.status] ?? "info", details: resolved.reason || `Student is currently ${resolved.status}.`, weightedMean: "0.00", summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0, isOnLeave: true }, passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [], academicYearName: "" },
    });
  }

  const academicStatus = await calculateStudentStatus(student._id, (student.program as any)._id, "", targetYearOfStudy);
  res.json({ student: { _id: student._id, name: student.name, regNo: student.regNo, programId: (student.program as any)?._id, currentYear: student.currentYearOfStudy, programName: (student.program as any)?.name, status: student.status }, grades: sortedGrades, academicStatus });
}));

const formatChallenges = (analysis: any) => ({
  supplementary: analysis.failedList.filter((f: any) => f.attempt === "A/S").map((f: any) => f.displayName.split(":")[0]),
  retakes: analysis.failedList.filter((f: any) => ["A/RA","A/RPU"].includes(f.attempt)).map((f: any) => f.displayName.split(":")[0]),
  stayouts: analysis.failedList.filter((f: any) => f.attempt === "A/SO").map((f: any) => f.displayName.split(":")[0]),
  specials: analysis.specialList.map((s: any) => s.displayName.split(":")[0]),
  incomplete: [...analysis.incompleteList, ...analysis.missingList].map((i) => i.split(":")[0]),
});

router.get(
  "/journey",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { regNo } = req.query;
    if (!regNo) return res.status(400).json({ error: "regNo is required" });

    const student = (await Student.findOne(
      scopeQuery(req, { regNo: regNo as string }),
    )
      .populate("admissionAcademicYear", "year")
      .populate("program")
      .lean()) as any;
    if (!student) return res.status(404).json({ error: "Student not found" });

    // ── Robust admission year resolution ────────────────────────────────────
    // Tries every field the Student schema might use before falling back.
    const admissionYearString = (() => {
      // A) Populated ObjectId ref → { year: "2016/2017" }
      const ref = student.admissionAcademicYear?.year;
      if (ref && typeof ref === "string" && ref.includes("/")) return ref;

      // B) Plain string field "admissionYear" (some schemas store it directly)
      const plain = student.admissionYear;
      if (plain && typeof plain === "string" && plain.trim()) {
        if (plain.includes("/")) return plain.trim();
        const yr = parseInt(plain.trim());
        if (!isNaN(yr) && yr > 1990) return `${yr}/${yr + 1}`;
      }

      // C) Legacy field name
      const legacy = student.admissionAcademicYearString;
      if (legacy && typeof legacy === "string" && legacy.trim())
        return legacy.trim();

      // D) Extract from regNo — e.g. "E024-01-1231/2016" → "2016/2017"
      const regMatch = (student.regNo || "").match(/\/(\d{4})$/);
      if (regMatch) {
        const yr = parseInt(regMatch[1]);
        if (yr > 1990 && yr < 2100) return `${yr}/${yr + 1}`;
      }

      // E) Earliest entry in academicHistory
      const sorted = [...(student.academicHistory || [])].sort(
        (a: any, b: any) => (a.yearOfStudy || 0) - (b.yearOfStudy || 0),
      );
      const earliest = sorted[0]?.academicYear;
      if (earliest && typeof earliest === "string" && earliest.includes("/"))
        return earliest;

      return "N/A";
    })();

    const journey: any[] = [];
    const program = student.program as any;
    const entryType = student.entryType || "Direct";
    const duration = program?.durationYears || 5;
    const isGraduated = ["graduand", "graduated"].includes(student.status);

    const incrementYear = (yr: string, offset: number): string => {
      if (!yr || !yr.includes("/")) return yr;
      const [start] = yr.split("/").map(Number);
      return `${start + offset}/${start + offset + 1}`;
    };

    // ── Part A: Academic history entries ──────────────────────────────────────
    // Display weight comes directly from the registry — NOT from stored history.
    // This guarantees correct percentages (15/15/20/25/25) in the Telemetry card.
    let activeWeightedSum = 0; // used only for non-graduated students
    let activeWeightTotal = 0;

    for (const record of (student.academicHistory || []) as any[]) {
      if ((record.yearOfStudy || 1) > duration) continue; // skip over-duration entries

      // Weight for display — always from registry, never from stored history
      const weight = getYearWeight(program, entryType, record.yearOfStudy);

      let resolvedYear = record.academicYear;
      if (
        !resolvedYear ||
        (record.yearOfStudy > 1 && resolvedYear === admissionYearString)
      ) {
        resolvedYear = incrementYear(
          admissionYearString,
          record.yearOfStudy - 1,
        );
      }

      // Challenge display: call engine for non-graduated students only
      let challenges = {
        supplementary: [] as string[],
        retakes: [] as string[],
        stayouts: [] as string[],
        specials: [] as string[],
        incomplete: [] as string[],
      };
      let totalUnits = record.unitsTakenCount || 0;
      let displayStatus = record.isRepeatYear
        ? "REPEAT YEAR"
        : "PASS (PROMOTED)";

      if (!isGraduated) {
        try {
          const analysis = await calculateStudentStatus(
            student._id,
            student.program,
            resolvedYear,
            record.yearOfStudy,
            { forPromotion: false },
          );
          challenges = formatChallenges(analysis);
          totalUnits = analysis.summary.totalExpected || totalUnits;
          const mean = record.annualMeanMark || 0;
          if (
            !["ACADEMIC LEAVE", "GRADUATED", "DEFERMENT"].includes(
              analysis.status,
            )
          ) {
            displayStatus = record.isRepeatYear
              ? "REPEAT YEAR"
              : mean > 0 || analysis.summary.passed > 0
                ? "PASS (PROMOTED)"
                : analysis.status;
          }
          // Accumulate for active student WAA
          activeWeightedSum += parseFloat(analysis.weightedMean) * weight;
          activeWeightTotal += weight;
        } catch {
          /* keep defaults */
        }
      } else {
        // Graduated: just get unit count for the Telemetry card
        try {
          const count = await ProgramUnit.countDocuments({
            program: student.program?._id || student.program,
            requiredYear: record.yearOfStudy,
          });
          totalUnits = count || totalUnits;
        } catch {
          /* keep stored value */
        }
      }

      journey.push({
        type: "ACADEMIC",
        academicYear: resolvedYear,
        yearOfStudy: record.yearOfStudy,
        status: displayStatus,
        totalUnits,
        weight: Math.round(weight * 100), // ← registry value, always correct
        challenges,
        date: record.date || new Date(0),
        isCurrent: false,
      });
    }

    // ── Part B: Current year (active, non-graduated only) ─────────────────────
    const LOCKED = [
      "on_leave",
      "deferred",
      "deregistered",
      "discontinued",
      "graduated",
      "graduand",
    ];
    const isLocked = LOCKED.includes(student.status);
    const currentInHistory = (student.academicHistory || []).some(
      (h: any) => h.yearOfStudy === student.currentYearOfStudy,
    );
    const withinDuration = (student.currentYearOfStudy || 1) <= duration;

    if (!isLocked && !currentInHistory && withinDuration) {
      const currentYearDoc = await AcademicYear.findOne({
        institution: student.institution,
        isCurrent: true,
      }).lean();
      const weight = getYearWeight(
        program,
        entryType,
        student.currentYearOfStudy,
      );

      try {
        const live = await calculateStudentStatus(
          student._id,
          student.program,
          "CURRENT",
          student.currentYearOfStudy,
        );

        journey.push({
          type: "ACADEMIC",
          academicYear: currentYearDoc?.year || "Current",
          yearOfStudy: student.currentYearOfStudy,
          status: live.status,
          totalUnits: live.summary.totalExpected,
          weight: Math.round(weight * 100),
          challenges: formatChallenges(live),
          date: new Date(),
          isCurrent: true,
        });

        activeWeightedSum += parseFloat(live.weightedMean) * weight;
        activeWeightTotal += weight;
      } catch (e: any) {
        console.warn("[journey] Current year engine error:", e.message);
      }
    }

    // ── Part C: Status events ──────────────────────────────────────────────────
    (student.statusEvents || []).forEach((ev: any) => {
      journey.push({
        type: "STATUS_CHANGE",
        academicYear: ev.academicYear || "",
        toStatus: ev.toStatus,
        reason: ev.reason,
        date: ev.date || new Date(0),
      });
    });

    journey.sort(
      (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime(),
    );

    // ── Cumulative mean ────────────────────────────────────────────────────────
    // GRADUATED/GRADUAND: always use finalWeightedAverage (authoritative).
    // This is the same value shown in the award list — no disparity.
    // ACTIVE: compute from engine means × registry weights.
    let projectedMean: number;

    if (isGraduated) {
      const stored = parseFloat(student.finalWeightedAverage || "0");
      if (stored > 0) {
        projectedMean = stored;
      } else {
        // finalWeightedAverage not yet set — trigger recompute via graduation engine
        // (this path should be rare after running /admin/recompute-graduation-waa)
        try {
          const { calculateGraduationStatus } =
            await import("../services/graduationEngine");
          const result = await calculateGraduationStatus(
            student._id.toString(),
          );
          projectedMean = result.weightedAggregateAverage;
        } catch {
          projectedMean = 0;
        }
      }
    } else {
      projectedMean =
        activeWeightTotal > 0 ? activeWeightedSum / activeWeightTotal : 0;
    }

    res.json({
      admissionYear: admissionYearString,
      intake: student.intake || "SEPT",
      currentStatus: student.status.toUpperCase(),
      cumulativeMean: projectedMean.toFixed(2),
      timeline: journey,
    });
  }),
);
router.post("/approve-special", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { markId, reason, undo = false } = req.body;
  if (!markId) return res.status(400).json({ error: "Mark ID is required" });
  const upd = undo
    ? { $set: { isSpecial: false, attempt: "1st", remarks: "Special Exam Revoked" } }
    : { $set: { isSpecial: true, attempt: "special", remarks: `Special Granted: ${reason || "Administrative"}` } };
  let mark = await Mark.findByIdAndUpdate(markId, upd, { new: true });
  if (!mark) mark = await MarkDirect.findByIdAndUpdate(markId, upd, { new: true });
  if (!mark) return res.status(404).json({ error: "Mark record not found" });
  try {
    const gr = await computeFinalGrade({ markId: mark._id as any, coordinatorReq: req });
    res.json({ success: true, message: undo ? "Special exam revoked" : "Special exam approved", newStatus: gr.status });
  } catch (e: any) { res.status(500).json({ error: e.message }); }
}));

router.post("/leave", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { studentId, startDate, endDate, reason, leaveType } = req.body;
  const result = await grantAcademicLeave(studentId, new Date(startDate), new Date(endDate), reason, leaveType);
  res.json({ success: true, student: result });
}));

router.post("/defer", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { studentId, years } = req.body;
  const result = await deferAdmission(studentId, parseInt(years) as 1 | 2);
  res.json({ success: true, student: result });
}));

router.post("/revert-active", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { studentId } = req.body;
  const result = await revertStatusToActive(studentId);
  res.json({ success: true, student: result });
}));

router.post("/readmit", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { studentId, remarks } = req.body;
  if (!studentId) return res.status(400).json({ error: "Student ID required" });
  const result = await readmitStudent(studentId, remarks || "Approved by Senate/Department");
  console.log(`[READMISSION] ${result?.regNo} reinstated by ${req.user?.name}`);
  res.json({ success: true, message: "Student readmitted successfully", student: result });
}));

router.get("/raw-marks", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { regNo, yearOfStudy } = req.query;
  if (!regNo) return res.status(400).json({ error: "regNo required" });
  const student = await Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } });
  if (!student) return res.status(404).json({ error: "Student not found" });
  const [dm, dmt] = await Promise.all([
    Mark.find({ student: student._id }).populate({ path: "programUnit", match: yearOfStudy ? { requiredYear: parseInt(yearOfStudy as string) } : {}, populate: { path: "unit", select: "code name" } }).populate("academicYear", "year").lean(),
    MarkDirect.find({ student: student._id }).populate({ path: "programUnit", match: yearOfStudy ? { requiredYear: parseInt(yearOfStudy as string) } : {}, populate: { path: "unit", select: "code name" } }).populate("academicYear", "year").lean(),
  ]);
  res.json([...dm.filter((m) => m.programUnit).map((m) => ({ ...m, entryMode: "detailed" })), ...dmt.filter((m) => m.programUnit).map((m) => ({ ...m, entryMode: "direct" }))]);
}));

router.post("/raw-marks", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  const { regNo, unitCode, academicYear, semester, cat1, cat2, cat3, assignment1, practicalRaw, examQ1, examQ2, examQ3, examQ4, examQ5, examMode, caDirect, examDirect, agreedMark, isSpecial, attempt } = req.body;
  if (!regNo || !unitCode || !academicYear) return res.status(400).json({ error: "regNo, unitCode, and academicYear are required" });
  const [student, unitDoc, yearDoc] = await Promise.all([Student.findOne({ regNo: { $regex: `^${regNo}$`, $options: "i" } }), Unit.findOne({ code: unitCode.toUpperCase().trim() }), AcademicYear.findOne({ year: academicYear })]);
  if (!student || !unitDoc || !yearDoc) return res.status(404).json({ error: "Required metadata not found" });
  const programUnit = await ProgramUnit.findOne({ program: student.program, unit: unitDoc._id });
  if (!programUnit) return res.status(400).json({ error: `Unit ${unitCode} not linked to student's program.` });
  const isDirectEntry = caDirect !== undefined || examDirect !== undefined;
  if (isDirectEntry) {
    const upd = { institution: student.institution, student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id, semester: semester || "SEMESTER 1", caTotal30: Number(caDirect)||0, examTotal70: Number(examDirect)||0, agreedMark: Number(agreedMark)||(Number(caDirect)+Number(examDirect)), attempt: isSpecial ? "special" : (attempt||"1st"), isSpecial: isSpecial===true, uploadedBy: req.user._id, uploadedAt: new Date() };
    const u = await MarkDirect.findOneAndUpdate({ student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id }, upd, { upsert: true, new: true, setDefaultsOnInsert: true });
    return res.json({ success: true, message: "Direct mark saved", data: u });
  }
  const upd = { institution: student.institution, student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id, semester: semester||"SEMESTER 1", cat1Raw: Number(cat1)||0, cat2Raw: Number(cat2)||0, cat3Raw: cat3?Number(cat3):undefined, assgnt1Raw: Number(assignment1)||0, practicalRaw: Number(practicalRaw)||0, examQ1Raw: Number(examQ1)||0, examQ2Raw: Number(examQ2)||0, examQ3Raw: Number(examQ3)||0, examQ4Raw: Number(examQ4)||0, examQ5Raw: Number(examQ5)||0, examMode: examMode||"standard", isSpecial: isSpecial===true, attempt: isSpecial?"special":(attempt||"1st"), uploadedBy: req.user._id, uploadedAt: new Date() };
  const u = await Mark.findOneAndUpdate({ student: student._id, programUnit: programUnit._id, academicYear: yearDoc._id }, upd, { upsert: true, new: true, setDefaultsOnInsert: true });
  const gr = await computeFinalGrade({ markId: u._id as any, coordinatorReq: req });
  return res.json({ success: true, message: "Detailed marks processed", data: gr });
}));

export default router;