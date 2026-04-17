// serverside/src/services/statusEngine.ts
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import Student from "../models/Student";
import Mark from "../models/Mark";
import InstitutionSettings from "../models/InstitutionSettings";
import MarkDirect from "../models/MarkDirect";
import AcademicYear from "../models/AcademicYear";
import { getYearWeight } from "../utils/weightingRegistry";
import { performAcademicAudit } from "./academicAudit";
import { getAttemptLabel, REG_QUALIFIERS } from "../utils/academicRules";

// ── Lazy-load carry-forward so a broken CF module never prevents promotions ──
// If carryForwardService throws on import, we fall back to a no-op.
let _assessAndGrant: ((studentId: string, programId: string, yearOfStudy: number, academicYearName: string) => Promise<{ granted: boolean; cfUnits: any[]; qualifier: string; reason: string }>) | null = null;

async function tryAssessAndGrantCarryForward(
  studentId:        string,
  programId:        string,
  yearOfStudy:      number,
  academicYearName: string,
): Promise<{ granted: boolean; cfUnits: any[]; qualifier: string; reason: string }> {
  try {
    if (!_assessAndGrant) {
      // Dynamic import so a bad carryForwardService never crashes this module
      const mod = await import("./carryForwardService");
      _assessAndGrant = mod.assessAndGrantCarryForward;
    }
    return await _assessAndGrant!(studentId, programId, yearOfStudy, academicYearName);
  } catch (err: any) {
    console.warn("[StatusEngine] carryForwardService unavailable, skipping CF check:", err.message);
    return { granted: false, cfUnits: [], qualifier: "", reason: err.message };
  }
}

// ─── Terminal DB sync ─────────────────────────────────────────────────────────

const syncTerminalStatusToDb = async (
  studentId:    string,
  engineStatus: string,
  details:      string,
  academicYear: string,
): Promise<void> => {
  interface Mapping { dbStatus: string; qualifierFn: (s: any) => string }
  const terminalMap: Record<string, Mapping> = {
    "DEREGISTERED": { dbStatus: "deregistered", qualifierFn: () => "" },
    "REPEAT YEAR":  {
      dbStatus: "repeat",
      qualifierFn: (s: any) => {
        const count = ((s.academicHistory || []) as any[]).filter((h: any) => h.isRepeatYear).length + 1;
        return REG_QUALIFIERS.repeatYear(count);
      },
    },
    "STAYOUT":      { dbStatus: "active",       qualifierFn: () => "" },
    "DISCONTINUED": { dbStatus: "discontinued", qualifierFn: () => "" },
  };

  const entry = terminalMap[engineStatus];
  if (!entry) return;

  const student = await Student.findById(studentId).lean();
  if (!student) return;
  if (entry.dbStatus !== "active" && (student as any).status === entry.dbStatus) return;

  const fromStatus   = (student as any).status;
  const newQualifier = entry.qualifierFn(student);

  const updatePayload: any = {
    $set:  { status: entry.dbStatus, remarks: details },
    $push: {
      statusEvents:  { fromStatus, toStatus: entry.dbStatus, date: new Date(), reason: `Auto-Sync: ${details}`, academicYear },
      statusHistory: { status: entry.dbStatus, previousStatus: fromStatus, date: new Date(), reason: details },
    },
  };
  if (newQualifier) updatePayload.$set.qualifierSuffix = newQualifier;

  await Student.findByIdAndUpdate(studentId, updatePayload);
  // console.log(`[StatusEngine] syncTerminalStatus: ${fromStatus} → ${entry.dbStatus} (${engineStatus}) for student ${studentId}`);
};

// ─── Result type ──────────────────────────────────────────────────────────────

export interface StudentStatusResult {
  status:        string;
  variant:       "success" | "warning" | "error" | "info";
  details:       string;
  weightedMean:  string;
  sessionState:  string;
  summary:       { totalExpected: number; passed: number; failed: number; missing: number; isOnLeave?: boolean };
  passedList:    { code: string; mark: number }[];
  failedList:    { displayName: string; attempt: string | number }[];
  specialList:   { displayName: string; grounds: string }[];
  missingList:   string[];
  incompleteList: string[];
  leaveDetails?: string;
}

// ─── calculateStudentStatus ───────────────────────────────────────────────────

export const calculateStudentStatus = async (
  studentId:        any,
  programId:        any,
  academicYearName: string,
  yearOfStudy:      number = 1,
  options:          { forPromotion?: boolean } = {},
): Promise<StudentStatusResult> => {
  const settings = await InstitutionSettings.findOne().lean();
  if (!settings) throw new Error("Institution settings not found. Please configure grading scales in System Settings.");

  const passMark = (settings as any).passMark || 40;

  // Single fetch — fixes TS2451 "duplicate const student" from prior versions
  const student = await Student.findById(studentId).lean();
  if (!student) throw new Error("Student not found");

  // ── Terminal status gate ──────────────────────────────────────────────────
  const TERMINAL: Record<string, { label: string; variant: "info" | "error" | "success" | "warning" }> = {
    on_leave:     { label: "ACADEMIC LEAVE", variant: "info"    },
    deferred:     { label: "DEFERMENT",      variant: "info"    },
    discontinued: { label: "DISCONTINUED",   variant: "error"   },
    deregistered: { label: "DEREGISTERED",   variant: "error"   },
    graduated:    { label: "GRADUATED",      variant: "success" },
    graduand:     { label: "GRADUATED",      variant: "success" },
  };

  const terminalEntry = TERMINAL[(student as any).status ?? ""];
  if (terminalEntry) {
    const leaveType = (student as any).academicLeavePeriod?.type || "";
    const rem       = ((student as any).remarks || "").toLowerCase();
    let grounds = "";
    if (leaveType === "financial" || rem.includes("financial"))                                grounds = "FINANCIAL";
    else if (leaveType === "compassionate" || rem.includes("compassionate") || rem.includes("medical")) grounds = "COMPASSIONATE";
    else if (leaveType) grounds = leaveType.toUpperCase();

    const count = await ProgramUnit.countDocuments({ program: programId, requiredYear: yearOfStudy });
    return {
      status: terminalEntry.label, variant: terminalEntry.variant,
      details: `Student is currently ${terminalEntry.label}.${grounds ? ` Grounds: ${grounds}.` : ""}`,
      weightedMean: "0.00", sessionState: "ORDINARY",
      summary: { totalExpected: count, passed: 0, failed: 0, missing: 0, isOnLeave: true },
      passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
      leaveDetails: grounds,
    };
  }

  // ── Academic year ─────────────────────────────────────────────────────────
  let targetYearDoc: any = null;
  if (!academicYearName || academicYearName === "CURRENT" || academicYearName === "undefined") {
    targetYearDoc = (await AcademicYear.findOne({ isCurrent: true }).lean())
      || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());
  } else {
    targetYearDoc = (await AcademicYear.findOne({ year: academicYearName }).lean())
      || (await AcademicYear.findOne({ year: { $regex: new RegExp(`^${academicYearName.replace("/", "\\/")}$`, "i") } }).lean());
    if (!targetYearDoc) console.warn(`[StatusEngine] AcademicYear "${academicYearName}" not found.`);
  }

  // ── Curriculum ────────────────────────────────────────────────────────────
  const curriculum = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy })
    .populate("unit").lean() as any[];

  if (!curriculum?.length) {
    return {
      status: "CURRICULUM NOT SET", variant: "info",
      details: `No units defined for Year ${yearOfStudy}. Contact admin.`,
      weightedMean: "0.00", sessionState: "ORDINARY",
      summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
      passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
    };
  }

  const programUnitIds = curriculum.map((pu: any) => pu._id);

  const [detailedMarks, directMarks, finalGrades] = await Promise.all([
    Mark.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
    MarkDirect.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
    FinalGrade.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
  ]);

  const marksMap = new Map<string, any>();
  finalGrades.forEach((fg: any) => {
    const key = fg.programUnit?.toString(); if (!key) return;
    marksMap.set(key, {
      agreedMark:  fg.totalMark ?? 0,
      caTotal30:   fg.caTotal30   != null ? fg.caTotal30   : (fg.totalMark > 0 ? 1 : 0),
      examTotal70: fg.examTotal70 != null ? fg.examTotal70 : (fg.totalMark > 0 ? 1 : 0),
      attempt:     fg.attemptType === "SUPPLEMENTARY" ? "supplementary" : fg.attemptType === "RETAKE" ? "re-take" : "1st",
      isSpecial:   fg.isSpecial === true || fg.status === "SPECIAL",
      source:      "finalGrade",
    });
  });
  directMarks.forEach((m: any)   => marksMap.set(m.programUnit.toString(), { ...m, source: "direct" }));
  detailedMarks.forEach((m: any) => marksMap.set(m.programUnit.toString(), { ...m, source: "detailed" }));

  const lists = {
    passed:     [] as { code: string; mark: number }[],
    failed:     [] as { displayName: string; attempt: string | number }[],
    special:    [] as { displayName: string; grounds: string }[],
    missing:    [] as string[],
    incomplete: [] as string[],
  };
  let totalFirstAttemptSum = 0;

  curriculum.forEach((pUnit: any) => {
    const code        = pUnit.unit?.code?.toUpperCase();
    const displayName = `${code}: ${pUnit.unit?.name}`;
    const rawMark     = marksMap.get(pUnit._id.toString());

    if (!rawMark) { lists.missing.push(displayName); return; }

    const hasCAT  = (rawMark.caTotal30  || 0) > 0;
    const hasExam = (rawMark.examTotal70 || 0) > 0;
    const markVal = rawMark.agreedMark || 0;
    const isSupp  = rawMark.attempt === "supplementary";
    const isSpc   = rawMark.attempt === "special" || rawMark.isSpecial;

    const notation = getAttemptLabel({
      markAttempt:      rawMark.attempt,
      studentStatus:    (student as any).status,
      studentQualifier: (student as any).qualifierSuffix,
    });

    if (isSpc) {
      lists.special.push({ displayName, grounds: rawMark.remarks || "Special" });
    } else if (!hasCAT && !hasExam) {
      lists.missing.push(`${displayName} (Absent)`);
    } else if (!hasCAT && hasExam) {
      if (isSupp) {
        if (markVal >= passMark) lists.passed.push({ code, mark: markVal });
        else                     lists.failed.push({ displayName, attempt: notation });
        totalFirstAttemptSum += markVal;
      } else {
        lists.incomplete.push(`${displayName} (No CAT)`);
      }
    } else if (!hasExam && hasCAT) {
      lists.missing.push(`${displayName} (Missing Exam)`);
    } else {
      if (markVal >= passMark) lists.passed.push({ code, mark: markVal });
      else                     lists.failed.push({ displayName, attempt: notation });
      totalFirstAttemptSum += markVal;
    }
  });

  const totalUnits   = curriculum.length;
  const failCount    = lists.failed.length;
  const missingCount = lists.missing.length;
  const specialCount = lists.special.length;
  const incCount     = lists.incomplete.length;
  const officialMean = totalUnits > 0 ? totalFirstAttemptSum / totalUnits : 0;
  const attemptedN   = totalUnits - (specialCount + missingCount + incCount);
  const perfMean     = attemptedN > 0 ? totalFirstAttemptSum / attemptedN : 0;

  const currentYearDoc = targetYearDoc?.isCurrent
    ? targetYearDoc
    : (await AcademicYear.findOne({ isCurrent: true }).lean()) || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

  const targetSession   = targetYearDoc?.session ?? "ORDINARY";
  const [tStart]        = (academicYearName || "0/0").split("/").map(Number);
  const [gStart]        = currentYearDoc?.year ? currentYearDoc.year.split("/").map(Number) : [0];
  const isPastYear      = targetYearDoc && gStart > 0 ? tStart < gStart : false;
  const isSessionClosed = targetSession === "CLOSED" || isPastYear;

  let status  = "PASS";
  let variant: "success" | "warning" | "error" | "info" = "success";
  let details = "Proceed to next year.";

  if (!options.forPromotion && targetSession === "ORDINARY" && !isPastYear) {
    status = "SESSION IN PROGRESS"; variant = "info"; details = "Marks are currently being entered.";
  } else if (missingCount >= 6 && isSessionClosed) {
    status = "DEREGISTERED"; variant = "error"; details = `Absent from 6+ (${missingCount}) examinations (ENG 23c).`;
  } else if (specialCount > 0 && failCount < totalUnits / 2) {
    const parts: string[] = [];
    if (failCount > 0)    parts.push(`SUPP ${failCount}`);
    parts.push(`SPEC ${specialCount}`);
    if (incCount > 0)     parts.push(`INC ${incCount}`);
    if (missingCount > 0) parts.push(`MISSING ${missingCount}`);
    status = parts.join("; "); variant = "info"; details = `Awaiting specials. Mean: ${perfMean.toFixed(2)}`;
  } else if (failCount >= totalUnits / 2 || officialMean < 40) {
    status = "REPEAT YEAR"; variant = "error";
    details = `Failed >= 50% (${failCount}/${totalUnits}) or Mean (${officialMean.toFixed(2)}) < 40% (ENG 16).`;
  } else if (failCount > totalUnits / 3) {
    status = "STAYOUT"; variant = "warning";
    details = `Failed > 1/3 (${failCount}/${totalUnits}). Retake in next ordinary period (ENG 15h).`;
  } else if (failCount > 0 || incCount > 0 || missingCount > 0) {
    const parts: string[] = [];
    if (failCount > 0)    parts.push(`SUPP ${failCount}`);
    if (incCount > 0)     parts.push(`INC ${incCount}`);
    if (missingCount > 0) parts.push(`INC ${missingCount}`);
    status = parts.join("; "); variant = "warning"; details = "Eligible for supplementary exams.";
  }

  return {
    status, variant, details,
    weightedMean:  officialMean.toFixed(2),
    sessionState:  targetSession,
    summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: missingCount },
    passedList:    lists.passed,
    failedList:    lists.failed,
    specialList:   lists.special,
    missingList:   lists.missing,
    incompleteList: lists.incomplete,
  };
};

// ─── promoteStudent ───────────────────────────────────────────────────────────

export const promoteStudent = async (studentId: string) => {
  // console.log(`[promoteStudent] Starting for studentId: ${studentId}`);

  const student = await Student.findById(studentId).populate("program");
  if (!student) {
    // console.error(`[promoteStudent] Student ${studentId} not found`);
    throw new Error("Student not found");
  }

  const st          = (student as any).status as string;
  const regNo       = (student as any).regNo;
  const currentYear = (student as any).currentYearOfStudy || 1;

  // console.log(`[promoteStudent] Student: ${regNo}, status: ${st}, yearOfStudy: ${currentYear}`);

  // ── Hard blocks ───────────────────────────────────────────────────────────
  if (["deregistered", "discontinued", "graduated", "graduand"].includes(st)) {
    // console.log(`[promoteStudent] Blocked — student is ${st}`);
    return { success: false, message: `Action blocked: Student is ${st}` };
  }

  if (st !== "active" && st !== "repeat") {
    // console.log(`[promoteStudent] Blocked — unexpected status: ${st}`);
    return { success: false, message: `Promotion blocked: Student status is ${st}` };
  }

  // ── Academic audit ────────────────────────────────────────────────────────
  const auditResult = await performAcademicAudit(studentId);
  if (auditResult.discontinued) {
    // console.log(`[promoteStudent] Discontinued by audit: ${auditResult.reason}`);
    return { success: false, message: `Discontinued: ${auditResult.reason}` };
  }

  const program      = student.program as any;
  const duration     = program.durationYears || 5;
  const currentSession = await AcademicYear.findOne({ isCurrent: true }).lean();
  const completedYear  = currentSession?.year || "N/A";

  // console.log(`[promoteStudent] Running status engine for ${regNo}, year: ${currentYear}, academicYear: ${completedYear}`);

  const statusResult = await calculateStudentStatus(student._id, student.program, completedYear, currentYear, { forPromotion: true });

  // console.log(`[promoteStudent] Status result for ${regNo}: status="${statusResult.status}", mean=${statusResult.weightedMean}, failed=${statusResult.summary.failed}/${statusResult.summary.totalExpected}`);

  // ── DEREGISTERED / DISCONTINUED ───────────────────────────────────────────
  if (["DEREGISTERED", "DISCONTINUED"].includes(statusResult.status)) {
    // console.log(`[promoteStudent] Syncing terminal status: ${statusResult.status}`);
    await syncTerminalStatusToDb(studentId, statusResult.status, statusResult.details, completedYear);
    return { success: false, message: `Promotion Blocked: ${statusResult.status}`, details: statusResult };
  }

  // ── REPEAT YEAR ───────────────────────────────────────────────────────────
  if (statusResult.status === "REPEAT YEAR") {
    const repeatCount = ((student as any).academicHistory || []).filter((h: any) => h.isRepeatYear && h.yearOfStudy === currentYear).length + 1;
    const qualifier = REG_QUALIFIERS.repeatYear(repeatCount);

    // console.log(`[promoteStudent] Repeat year for ${regNo}: qualifier=${qualifier}`);

    await Student.findByIdAndUpdate(studentId, {
      $set:  { status: "repeat", qualifierSuffix: qualifier, remarks: statusResult.details },
      $push: {
        statusEvents:  { fromStatus: st, toStatus: "repeat", date: new Date(), academicYear: completedYear, reason: `ENG.16: ${statusResult.details}` },
        statusHistory: { status: "repeat", previousStatus: st, date: new Date(), reason: statusResult.details },
        academicHistory: {
          academicYear: completedYear, yearOfStudy: currentYear,
          annualMeanMark: parseFloat(statusResult.weightedMean),
          weightedContribution: 0, failedUnitsCount: statusResult.summary.failed,
          isRepeatYear: true, date: new Date(),
        },
      },
    });
    return { success: false, message: `Repeat year required (${qualifier})`, details: statusResult };
  }

  // ── STAYOUT ───────────────────────────────────────────────────────────────
  if (statusResult.status === "STAYOUT") {
    // console.log(`[promoteStudent] Stayout for ${regNo}: ${statusResult.details}`);
    await Student.findByIdAndUpdate(studentId, {
      $set:  { remarks: `ENG.15h: ${statusResult.details}` },
      $push: { statusEvents: { fromStatus: st, toStatus: "active", date: new Date(), academicYear: completedYear, reason: `ENG.15h: ${statusResult.details}` } },
    });
    return { success: false, message: "Stay out required (ENG.15h)", details: statusResult };
  }

  // ── SPECIALS PENDING (no failures alongside) ──────────────────────────────
  if (statusResult.specialList.length > 0 && statusResult.failedList.length === 0) {
    // console.log(`[promoteStudent] Special examinations pending for ${regNo}`);
    return { success: false, message: "Special examinations pending", details: statusResult };
  }

  // ── PASS path — proceed to promotion ─────────────────────────────────────
  // Also handles "SUPP N" students if N ≤ 2 and carry-forward is eligible.
  // The status being PASS means all units are cleared. Supp status means
  // some units failed at supplementary. The carry-forward check below handles both.

  const rawMean    = parseFloat(statusResult.weightedMean);
  const yearWeight = getYearWeight(program, (student as any).entryType || "Direct", currentYear);

  const histRecord = {
    academicYear:         completedYear,
    yearOfStudy:          currentYear,
    annualMeanMark:       rawMean,
    weightedContribution: rawMean * yearWeight,
    failedUnitsCount:     statusResult.summary.failed,
    isRepeatYear:         false,
    date:                 new Date(),
  };

  const pendingCF = ((student as any).carryForwardUnits || []).filter((u: any) => u.status === "pending").length;

  // console.log(`[promoteStudent] Promoting ${regNo}: mean=${rawMean}, weight=${yearWeight}, pendingCF=${pendingCF}, duration=${duration}`);

  // ── GRADUATION ────────────────────────────────────────────────────────────
  if (currentYear === duration) {
    if (pendingCF > 0) {
      // console.log(`[promoteStudent] Cannot graduate ${regNo}: ${pendingCF} CF unit(s) pending`);
      return { success: false, message: `Cannot graduate: ${pendingCF} carry-forward unit(s) not yet cleared` };
    }

    const fullHistory = [...((student as any).academicHistory || []), histRecord];
    const finalWAA    = fullHistory.reduce((acc: number, h: any) => acc + (h.weightedContribution || 0), 0);

    let classification = "PASS";
    if (finalWAA >= 70)      classification = "FIRST CLASS HONOURS";
    else if (finalWAA >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
    else if (finalWAA >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

    // console.log(`[promoteStudent] Graduating ${regNo}: WAA=${finalWAA.toFixed(2)}, class=${classification}`);

    await Student.findByIdAndUpdate(studentId, {
      $set: {
        status: "graduand", qualifierSuffix: "",
        finalWeightedAverage: finalWAA.toFixed(2), classification,
        graduationYear: new Date().getFullYear(), currentYearOfStudy: currentYear + 1,
      },
      $push: { academicHistory: histRecord },
    });
    return { success: true, message: `Graduated: ${classification}`, isGraduation: true };
  }

  // ── NORMAL PROMOTION ──────────────────────────────────────────────────────
  const nextYear = currentYear + 1;

  // Only attempt carry-forward check if student has SUPP status (failed some units)
  // PASS students get a clean promotion — no CF check needed.
  let cfGranted  = false;
  let cfMessage  = "";
  let cfQualifier = "";

  if (statusResult.status !== "PASS") {
    // Student still has unresolved units — attempt carry-forward
    // console.log(`[promoteStudent] Checking carry-forward for ${regNo} (status="${statusResult.status}")`);
    const cfResult = await tryAssessAndGrantCarryForward( studentId, student.program.toString(), currentYear, completedYear );
    cfGranted   = cfResult.granted;
    cfQualifier = cfResult.qualifier;
    cfMessage   = cfGranted
      ? `Promoted to Year ${nextYear} with carry-forward (${cfResult.qualifier}): ${cfResult.cfUnits.map((u) => u.unitCode).join(", ")}`
      : "";
    // console.log(`[promoteStudent] CF result for ${regNo}: granted=${cfGranted}, reason="${cfResult.reason}"`);

    // If carry-forward was NOT granted and status is not PASS, block the promotion
    if (!cfGranted) {
      // console.log(`[promoteStudent] Cannot promote ${regNo}: ${cfResult.reason}`);
      return { success: false, message: `Promotion Blocked: ${cfResult.reason}`, details: statusResult };
    }
  }

  // Apply the promotion
  // console.log(`[promoteStudent] Applying promotion for ${regNo}: Year ${currentYear} → ${nextYear}${cfGranted ? ` (CF: ${cfQualifier})` : " (clean pass)"}`);

  await Student.findByIdAndUpdate(studentId, {
    $set: {
      currentYearOfStudy: nextYear,
      currentSemester:    1,
      // Clear qualifier on clean pass; CF qualifier was already set by assessAndGrantCarryForward
      ...((!cfGranted) ? { qualifierSuffix: "" } : {}),
    },
    $push: {
      promotionHistory: { from: currentYear, to: nextYear, date: new Date() },
      academicHistory:  histRecord,
      statusHistory:    { status: "active", previousStatus: st, date: new Date(), reason: `Promoted to Year ${nextYear}` },
    },
  });

  const successMessage = cfGranted
    ? cfMessage
    : `Successfully promoted to Year ${nextYear}`;

  // console.log(`[promoteStudent] ✓ ${regNo} → Year ${nextYear}. ${successMessage}`);

  return { success: true, message: successMessage };
};

// ─── previewPromotion ─────────────────────────────────────────────────────────

export const previewPromotion = async (
  programId:        string,
  yearToPromote:    number,
  academicYearName: string,
) => {
  const nextYear      = yearToPromote + 1;
  const targetYearDoc = await AcademicYear.findOne({ year: academicYearName }).lean();

  if (!targetYearDoc) {
    console.warn(`[previewPromotion] AcademicYear "${academicYearName}" not found.`);
    return { totalProcessed: 0, eligibleCount: 0, blockedCount: 0, eligible: [], blocked: [] };
  }

  const admissionStudents = await Student.find({
    program: programId, currentYearOfStudy: yearToPromote,
    admissionAcademicYear: (targetYearDoc as any)._id,
  }).lean();

  const [m1, m2] = await Promise.all([
    Mark.distinct("student",       { academicYear: (targetYearDoc as any)._id }),
    MarkDirect.distinct("student", { academicYear: (targetYearDoc as any)._id }),
  ]);

  const markedIds    = new Set<string>([...m1, ...m2].map((id: any) => id.toString()));
  const admissionIds = new Set(admissionStudents.map((s: any) => s._id.toString()));

  const returningStudents = await Student.find({
    program: programId, currentYearOfStudy: yearToPromote,
    _id: { $in: Array.from(markedIds), $nin: Array.from(admissionIds) },
  }).lean();

  const adminStudents = await Student.find({
    program: programId, currentYearOfStudy: yearToPromote,
    status: { $in: ["on_leave","deferred","deregistered","discontinued"] },
    $or: [
      { admissionAcademicYear: (targetYearDoc as any)._id },
      { "academicHistory.academicYear": academicYearName },
    ],
    _id: { $nin: [...Array.from(admissionIds), ...returningStudents.map((s: any) => s._id.toString())] },
  }).lean();

  const allStudents = [...admissionStudents, ...returningStudents, ...adminStudents];

  // console.log(`[previewPromotion] Program=${programId}, Year=${yearToPromote}, AcadYear=${academicYearName}. Total students in cohort: ${allStudents.length}`);

  const ADMIN_LABELS: Record<string, string> = {
    on_leave: "ACADEMIC LEAVE", deferred: "DEFERMENT",
    discontinued: "DISCONTINUED", deregistered: "DEREGISTERED", graduated: "GRADUATED",
  };

  const eligible: any[] = [];
  const blocked:  any[] = [];

  for (const student of allStudents) {
    const isAlreadyPromoted = (student as any).currentYearOfStudy === nextYear;
    const adminLabel        = ADMIN_LABELS[(student as any).status];

    if (isAlreadyPromoted) {
      eligible.push({
        id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
        status: "ALREADY PROMOTED", reasons: [], specialGrounds: "",
        summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 }, details: "",
      });
      continue;
    }

    if (adminLabel) {
      const leaveType    = (student as any).academicLeavePeriod?.type?.toUpperCase();
      const adminGrounds = [
        ((student as any).academicLeavePeriod?.type || "").toLowerCase(),
        ((student as any).remarks || "").toLowerCase(),
      ].join(" ").trim() || "other";

      blocked.push({
        id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
        status: adminLabel,
        reasons: [leaveType ? `${adminLabel} (${leaveType})` : adminLabel],
        specialGrounds: adminGrounds,
        summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
        academicLeavePeriod: (student as any).academicLeavePeriod,
        remarks: (student as any).remarks, details: "",
      });
      continue;
    }

    const sr = await calculateStudentStatus(
      (student as any)._id, programId, academicYearName, yearToPromote, { forPromotion: true },
    );

    const specialGrounds = [
      (sr.specialList || []).map((s: any) => (s.grounds || "").toLowerCase()).join(" "),
      ((student as any).remarks || "").toLowerCase(),
      ((student as any).academicLeavePeriod?.type || "").toLowerCase(),
    ].join(" ").trim() || "other";

    const report: any = {
      id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
      status: sr.status, summary: sr.summary, reasons: [],
      remarks: (student as any).remarks, academicLeavePeriod: (student as any).academicLeavePeriod,
      details: sr.details, specialGrounds,
      isEligibleForSupp: !["STAYOUT","REPEAT YEAR","DEREGISTERED"].includes(sr.status) &&
        (sr.failedList.length > 0 || sr.specialList.length > 0),
    };

    if (sr.status === "PASS") {
      eligible.push(report);
    } else {
      if (sr.status === "STAYOUT")      report.reasons.push("ENG 15h: > 1/3 units failed");
      if (sr.status === "REPEAT YEAR")  report.reasons.push("ENG 16: >= 1/2 units failed or mean < 40%");
      if (sr.status === "DEREGISTERED") report.reasons.push("ENG 23c: Absent from 6+ examinations");
      sr.specialList.forEach((s: any)     => report.reasons.push(`${s.displayName} (SPECIAL)`));
      sr.incompleteList.forEach((u: string) => report.reasons.push(`${u} (INCOMPLETE)`));
      sr.missingList.forEach((u: string)    => report.reasons.push(`${u} (MISSING)`));
      sr.failedList.forEach((f: any)        => report.reasons.push(`${f.displayName} (FAIL: ${f.attempt})`));
      blocked.push(report);
    }
  }

  // console.log(`[previewPromotion] Result: ${eligible.length} eligible, ${blocked.length} blocked`);

  return {
    totalProcessed: allStudents.length,
    eligibleCount:  eligible.length,
    blockedCount:   blocked.length,
    eligible,
    blocked,
  };
};

// ─── bulkPromoteClass ─────────────────────────────────────────────────────────

export const bulkPromoteClass = async (
  programId:        string,
  yearToPromote:    number,
  academicYearName: string,
) => {
  const nextYear = yearToPromote + 1;

  // console.log(`[bulkPromoteClass] Starting: program=${programId}, year=${yearToPromote}, academicYear=${academicYearName}`);

  const students = await Student.find({
    program:            programId,
    currentYearOfStudy: { $in: [yearToPromote, nextYear] },
    status:             { $in: ["active", "repeat"] },
  });

  // console.log(`[bulkPromoteClass] Found ${students.length} candidate students`);

  const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };

  for (const student of students) {
    const sid   = (student._id as any).toString();
    const rNo   = (student as any).regNo;
    const curYr = (student as any).currentYearOfStudy;

    try {
      if (curYr >= nextYear) {
        // console.log(`[bulkPromoteClass] ${rNo}: already at Year ${curYr}, skipping`);
        results.alreadyPromoted++;
        results.promoted++;
        continue;
      }

      const res = await promoteStudent(sid);

      if (res.success) {
        // console.log(`[bulkPromoteClass] ✓ ${rNo}: ${res.message}`);
        results.promoted++;
      } else {
        // console.log(`[bulkPromoteClass] ✗ ${rNo}: ${res.message}`);
        results.failed++;
      }
    } catch (err: any) {
      // console.error(`[bulkPromoteClass] ERROR for ${rNo}:`, err.message, err.stack);
      results.errors.push(`${rNo}: ${err.message}`);
    }
  }

  // console.log(`[bulkPromoteClass] Done: promoted=${results.promoted}, failed=${results.failed}, alreadyPromoted=${results.alreadyPromoted}, errors=${results.errors.length}`);

  return results;
};