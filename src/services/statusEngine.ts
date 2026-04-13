// // serverside/src/services/statusEngine.ts
// import mongoose from "mongoose";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import Student from "../models/Student";
// import Mark from "../models/Mark"; // Required for checking special flags
// import InstitutionSettings from "../models/InstitutionSettings";
// import { getYearWeight } from "../utils/weightingRegistry";
// import { performAcademicAudit } from "./academicAudit";
// import { computeFinalGrade } from "./gradeCalculator"; // Assume this is used in the route
// import { resolveStudentStatus } from "../utils/studentStatusResolver";
// import MarkDirect from "../models/MarkDirect";
// import AcademicYear from "../models/AcademicYear";
// import { getAttemptLabel } from "../utils/academicRules";

// const getLetterGrade = (mark: number, settings: any): string => {
//   if (!settings || !settings.gradingScale) {
//     if (mark >= 69.5) return "A";
//     if (mark >= 59.5) return "B";
//     if (mark >= 49.5) return "C";
//     if (mark >= 39.5) return "D";
//     return "E";
//   }
//   const sortedScale = [...settings.gradingScale].sort((a, b) => b.min - a.min);
//   const matched = sortedScale.find((s) => mark >= s.min);
//   return matched ? matched.grade : settings.failingGrade || "E";
// };

// // const syncTerminalStatusToDb = async ( studentId: string, engineStatus: string, details: string, academicYear: string ) => {
// //   const terminalMap: Record<string, string> = { DEREGISTERED: "deregistered", "REPEAT YEAR": "repeat", STAYOUT: "active", DISCONTINUED: "discontinued"  };
// //   // const terminalMap: Record<string, string> = { DEREGISTERED: "deregistered", "REPEAT YEAR": "repeat", STAYOUT: "active", DISCONTINUED: "discontinued"  };

// //   const dbStatus = terminalMap[engineStatus];
// //   if (!dbStatus) return;

// //   const student = await Student.findById(studentId);
// //   // Only update if there's a change to avoid infinite loops/redundant writes
// //   if (!student || student.status === dbStatus) return;

// //   const fromStatus = student.status;

// //   await Student.findByIdAndUpdate(studentId, {
// //     $set: { status: dbStatus as any, remarks: details },
// //     $push: { statusEvents: { fromStatus, toStatus: dbStatus, date: new Date(), reason: `Auto-Sync: ${details}`, academicYear }},
// //   });
// // };

// const syncTerminalStatusToDb = async (
//   studentId:    string,
//   engineStatus: string,
//   details:      string,
//   academicYear: string,
// ): Promise<void> => {
//   type Mapping = { dbStatus: string; qualifierFn: (s: any) => string };

//   const terminalMap: Record<string, Mapping> = {
//     "DEREGISTERED": {
//       dbStatus:    "deregistered",
//       qualifierFn: () => "",                    // no qualifier on reg — just DB status
//     },
//     "REPEAT YEAR": {
//       dbStatus:    "repeat",
//       qualifierFn: (student) => {               // RP1, RP2 etc.
//         const count = ((student.academicHistory || []) as any[]).filter(
//           (h: any) => h.isRepeatYear,
//         ).length + 1;
//         return REG_QUALIFIERS.repeatYear(count);
//       },
//     },
//     "STAYOUT": {
//       dbStatus:    "active",                    // stayout keeps DB status "active"
//       qualifierFn: () => "",                    // qualifier shown in scoresheet only
//     },
//     "DISCONTINUED": {
//       dbStatus:    "discontinued",
//       qualifierFn: () => "",
//     },
//   };

//   const entry = terminalMap[engineStatus];
//   if (!entry) return;

//   const student = await Student.findById(studentId).lean();
//   if (!student) return;

//   // Only sync if there is a status change (prevents re-triggering repeat year sync)
//   if (entry.dbStatus !== "active" && (student as any).status === entry.dbStatus) return;

//   const fromStatus  = (student as any).status;
//   const newQualifier = entry.qualifierFn(student);

//   const updatePayload: any = {
//     $set: {
//       status:  entry.dbStatus,
//       remarks: details,
//     },
//     $push: {
//       statusEvents: {
//         fromStatus,
//         toStatus:     entry.dbStatus,
//         date:         new Date(),
//         reason:       `Auto-Sync: ${details}`,
//         academicYear,
//       },
//       statusHistory: {
//         status:         entry.dbStatus,
//         previousStatus: fromStatus,
//         date:           new Date(),
//         reason:         details,
//       },
//     },
//   };

//   // Only update qualifier if it changes (don't wipe RP1C by syncing STAYOUT again)
//   if (newQualifier) {
//     updatePayload.$set.qualifierSuffix = newQualifier;
//   }

//   await Student.findByIdAndUpdate(studentId, updatePayload);
// };

// export interface StudentStatusResult {
//   status: string;
//   variant: "success" | "warning" | "error" | "info";
//   details: string;
//   weightedMean: string;
//   sessionState: string;
//   summary: { totalExpected: number; passed: number; failed: number; missing: number; isOnLeave?: boolean; };
//   passedList: { code: string; mark: number }[];
//   failedList: { displayName: string; attempt: string | number }[];
//   specialList: { displayName: string; grounds: string }[];
//   missingList: string[];
//   incompleteList: string[];
//   leaveDetails?: string;
// }

// export const calculateStudentStatus = async (
//   studentId: any,
//   programId: any,
//   academicYearName: string,
//   yearOfStudy: number = 1,
//   options: { forPromotion?: boolean } = {}, 
// ): Promise<StudentStatusResult> => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");

//   const passMark = settings?.passMark || 40;
//   const student = await Student.findById(studentId).lean();
//   if (!student) throw new Error("Student not found");

//   // ── Terminal status gate ──────────────────────────────────────────────────
//   // These students never reach the mark engine regardless of session or flags.
//   const TERMINAL_STATUSES: Record<string, { label: string; variant: "info" | "error" | "success" | "warning" }> = {
//     on_leave:     { label: "ACADEMIC LEAVE", variant: "info"    },
//     deferred:     { label: "DEFERMENT",      variant: "info"    },
//     discontinued: { label: "DISCONTINUED",   variant: "error"   },
//     deregistered: { label: "DEREGISTERED",   variant: "error"   },
//     graduated:    { label: "GRADUATED",      variant: "success" },
//     graduand:     { label: "GRADUATED",      variant: "success" },
//   };

//   const terminalEntry = TERMINAL_STATUSES[student.status ?? ""];
//   if (terminalEntry) {
//     // Resolve leave grounds from multiple sources
//     const leaveType = (student.academicLeavePeriod?.type || "").toLowerCase();
//     const rem       = (student.remarks || "").toLowerCase();
//     let grounds = "";
//     if      (leaveType === "financial"     || rem.includes("financial"))                               grounds = "FINANCIAL";
//     else if (leaveType === "compassionate" || rem.includes("compassionate") || rem.includes("medical")) grounds = "COMPASSIONATE";
//     else if (student.academicLeavePeriod?.type) grounds = student.academicLeavePeriod.type.toUpperCase();

//     // Fetch curriculum count only (cheap — no mark queries)
//     const curriculumCount = await ProgramUnit.countDocuments({ program: programId, requiredYear: yearOfStudy });

//     return {
//       status:       terminalEntry.label,
//       variant:      terminalEntry.variant,
//       details:      `Student is currently ${terminalEntry.label}.${grounds ? ` Grounds: ${grounds}.` : ""}`,
//       weightedMean: "0.00",
//       sessionState: "ORDINARY",
//       summary: { totalExpected: curriculumCount, passed: 0, failed: 0, missing: 0, isOnLeave: true },
//       passedList:    [],
//       failedList:    [],
//       specialList:   [],
//       missingList:   [],
//       incompleteList: [],
//       leaveDetails:  grounds,
//     };
//   }

//   // --------------

//   let targetYearDoc: any = null;

//   if ( !academicYearName || academicYearName === "CURRENT" || academicYearName === "undefined" ) {
//     targetYearDoc = (await AcademicYear.findOne({ isCurrent: true }).lean()) || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

//     if (!targetYearDoc) console.warn("[StatusEngine] No AcademicYear documents found. Proceeding without session context.");
//   } else {
//     targetYearDoc =
//       (await AcademicYear.findOne({ year: academicYearName }).lean()) ||
//       (await AcademicYear.findOne({ year: { $regex: new RegExp(`^${academicYearName.replace("/", "\\/")}$`, "i")}}).lean());

//     if (!targetYearDoc) console.warn(`[StatusEngine] AcademicYear "${academicYearName}" not found. isPastYear/session checks will be skipped.`);
//   }

//   const curriculum = (await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean()) as any[];

//   if (!curriculum || !curriculum.length || curriculum.length === 0) {
//     return {
//       status: "CURRICULUM NOT SET", variant: "info",
//       details: `No units defined for Year ${yearOfStudy} in this program. Please contact the Admin to set the curriculum.`,
//       weightedMean: "0.00", summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//       sessionState: "ORDINARY",
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
//     };
//   }

//   const programUnitIds = curriculum.map((pu) => pu._id);

//   const [detailedMarks, directMarks, finalGrades] = await Promise.all([
//     Mark.find({ student: studentId, programUnit: { $in: programUnitIds }}).lean(),
//     MarkDirect.find({ student: studentId, programUnit: { $in: programUnitIds }}).lean(),
//     FinalGrade.find({ student: studentId, programUnit: { $in: programUnitIds }}).lean(),
//   ]);

//   const marksMap = new Map();

//   // Priority 3 (lowest): FinalGrade — the computed official record.
//   // Used as fallback when no raw Mark/MarkDirect document exists,
//   // which happens for legacy imports or historical data.
//   finalGrades.forEach((fg: any) => {
//     const key = fg.programUnit?.toString();
//     if (!key) return;

//     // Build a mark-compatible object from FinalGrade fields
//     marksMap.set(key, {
//       agreedMark: fg.totalMark ?? 0,
//       // If caTotal30 is stored, use it. Otherwise infer from totalMark.
//       // A non-zero value prevents the "No CAT" incomplete trigger.
//       caTotal30: fg.caTotal30 != null ? fg.caTotal30 : fg.totalMark > 0 ? 1 : 0,
//       examTotal70: fg.examTotal70 != null ? fg.examTotal70 : fg.totalMark > 0 ? 1 : 0,
//       attempt: fg.attemptType === "SUPPLEMENTARY" ? "supplementary" : fg.attemptType === "RETAKE" ? "re-take" : "1st",
//       isSpecial: fg.isSpecial === true || fg.status === "SPECIAL",
//       isSupplementary: fg.status === "SUPPLEMENTARY",
//       isMissingCA: false, // FinalGrade exists = coursework was assessed
//       source: "finalGrade",
//     });
//   });

//   // Priority 2: MarkDirect — overrides FinalGrade when present
//   directMarks.forEach((m) => marksMap.set(m.programUnit.toString(), { ...m, source: "direct" }));

//   // Priority 1 (highest): Detailed Mark — most granular, always wins
//   detailedMarks.forEach((m) => marksMap.set(m.programUnit.toString(), { ...m, source: "detailed" }));
//   // LOG 2: Verify the Map contains the keys
//   // console.log(`[StatusEngine] Map Keys:`, Array.from(marksMap.keys()));

//   const lists = {
//     passed: [] as { code: string; mark: number }[],
//     failed: [] as { displayName: string; attempt: string | number }[],
//     special: [] as { displayName: string; grounds: string }[],
//     missing: [] as string[],
//     incomplete: [] as string[],
//   };

//   let totalFirstAttemptSum = 0;
//   let unitsContributingToMean = 0;

//   curriculum.forEach((pUnit) => {
//     const code = pUnit.unit?.code?.toUpperCase();
//     const displayName = `${code}: ${pUnit.unit?.name}`;
//     // const rawMarkRecord = marksMap.get(pUnit._id.toString());
//     const unitId = pUnit._id.toString();
//     const rawMarkRecord = marksMap.get(unitId);

//     if (!rawMarkRecord) {
//       lists.missing.push(displayName);
//       return;
//     }

//     // Direct marks already have totals; Detailed marks rely on agreedMark calculated previously
//     const hasCAT = (rawMarkRecord.caTotal30 || 0) > 0;
//     const hasExam = (rawMarkRecord.examTotal70 || 0) > 0;
//     const markValue = rawMarkRecord.agreedMark || 0;
//     const isSupplementary = rawMarkRecord.attempt === "supplementary";
//     const isSpecial = rawMarkRecord.attempt === "special" || (rawMarkRecord.source === "detailed" && rawMarkRecord.isSpecial);
//     // const notation = getAttemptLabel( rawMarkRecord.attemptNumber || 1, student.status, student.regNo );
//     const notation = getAttemptLabel({
//       markAttempt: rawMarkRecord.attempt,
//       studentStatus: student.status,
//       // regNo: student.regNo,
//       studentQualifier: (student as any).qualifierSuffix,

//     });
//     // Case B: Special Exams
//     if (isSpecial)
//       lists.special.push({ displayName, grounds: rawMarkRecord.remarks || "Special" });
//     // Case C: Absolute Zero (No CAT AND No Exam) -> ENG 23c
//     else if (!hasCAT && !hasExam) lists.missing.push(`${displayName} (Absent)`);
//     // Case D: Partial Data -> ENG 15a
//     // else if (!hasCAT && hasExam) lists.incomplete.push(`${displayName} (No CAT)`);
//     else if (!hasCAT && hasExam) {
//       // If it's a Supp, No CAT is allowed/expected.
//       // Move it to Passed/Failed instead of Incomplete.
//       if (isSupplementary) {
//         if (markValue >= passMark) lists.passed.push({ code, mark: markValue });
//         // else lists.failed.push({ displayName, attempt: 2 }); // attempt 2 for supp
//         else lists.failed.push({ displayName, attempt: notation });
//         totalFirstAttemptSum += markValue;
//         unitsContributingToMean++;
//       } else lists.incomplete.push(`${displayName} (No CAT)`);
//     } else if (!hasExam && hasCAT)
//       lists.missing.push(`${displayName} (Missing Exam)`);
//     // Case E: Numerical Result (Passed or Failed)
//     else {
//       if (markValue >= passMark) lists.passed.push({ code, mark: markValue });
//       // else lists.failed.push({ displayName, attempt: 1 });
//       else lists.failed.push({ displayName, attempt: notation });

//       // ENG 16c: Mean is based on first attempt marks
//       totalFirstAttemptSum += markValue;
//       unitsContributingToMean++;
//     }
//   });

//   const totalUnits = curriculum.length;
//   const failCount = lists.failed.length;
//   const missingCount = lists.missing.length;
//   const specialCount = lists.special.length;
//   const incCount = lists.incomplete.length;
  
//   const attemptedUnitsCount = totalUnits - (specialCount + missingCount + incCount);
//   const performanceMean = attemptedUnitsCount > 0 ? totalFirstAttemptSum / attemptedUnitsCount : 0;

//   // The official Mean for ENG 16 (where missing = 0)
//   const officialMean = totalFirstAttemptSum / totalUnits;

//   const currentYearDoc = targetYearDoc?.isCurrent
//     ? targetYearDoc
//     : (await AcademicYear.findOne({ isCurrent: true }).lean()) ||
//       (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

//   // The session of the TARGET year (not the global current year)
//   const targetSession = targetYearDoc?.session ?? "ORDINARY";

//   const [targetStart] = (academicYearName || "0/0").split("/").map(Number);
//   const [globalStart] = currentYearDoc?.year ? currentYearDoc.year.split("/").map(Number) : [0];
//   const isPastYear    = targetYearDoc && globalStart > 0 ? targetStart < globalStart : false;
//   const isSessionClosed = targetSession === "CLOSED" || isPastYear;

//   // ── Status decision tree ───────────────────────────────────────────────────
//   let status = "PASS";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = "Proceed to next year.";

//   // 1. SESSION IN PROGRESS — only during ORDINARY and not a promotion/CMS call
//   if (!options.forPromotion && targetSession === "ORDINARY" && !isPastYear) {
//     status = "SESSION IN PROGRESS";
//     variant = "info";
//     details = "Marks are currently being entered for this session.";
//   }

//   // 2. DEREGISTERED (ENG 23c) — ONLY when the year is fully CLOSED or past
//   //    Never fires during ORDINARY or SUPPLEMENTARY — students may have
//   //    pending specials approved, which appear as missing marks until sat.
//   else if (missingCount >= 6 && isSessionClosed) {
//     status = "DEREGISTERED";
//     variant = "error";
//     details = `Absent from 6+ (${missingCount}) examinations (ENG 23c).`;
//   }

//   // 3. SPECIALS PENDING — pre-empts failure classification if specials exist
//   //    and failures have not yet exceeded the repeat threshold
//   else if (specialCount > 0 && failCount < totalUnits / 2) {
//     const parts: string[] = [];
//     if (failCount > 0) parts.push(`SUPP ${failCount}`);
//     parts.push(`SPEC ${specialCount}`);
//     if (incCount > 0) parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`MISSING ${missingCount}`);
//     status = parts.join("; ");
//     variant = "info";
//     details = `Awaiting specials. Mean in sat units: ${performanceMean.toFixed(2)}`;
//   }

//   // 4. REPEAT YEAR (ENG 16a / 16c)
//   else if (failCount >= totalUnits / 2 || officialMean < 40) {
//     status = "REPEAT YEAR";
//     variant = "error";
//     details =
//       `Failed >= 50% (${failCount}/${totalUnits}) units or Mean ` +
//       `(${officialMean.toFixed(2)}) < 40% (ENG 16).`;
//   }

//   // 5. STAYOUT (ENG 15h) — more than 1/3 but less than 1/2 failed
//   else if (failCount > totalUnits / 3) {
//     status = "STAYOUT";
//     variant = "warning";
//     details =
//       `Failed > 1/3 of units (${failCount}/${totalUnits}). ` +
//       `Retake failed units next year (ENG 15h).`;
//   }

//   // 6. SUPPLEMENTARY — at least one failure or incomplete
//   else if (failCount > 0 || incCount > 0 || missingCount > 0) {
//     const parts: string[] = [];
//     if (failCount > 0) parts.push(`SUPP ${failCount}`);
//     if (incCount > 0) parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`INC ${missingCount}`);
//     status = parts.join("; ");
//     variant = "warning";
//     details = "Eligible for supplementary exams or pending incomplete marks.";
//   }

//   // 7. PASS — falls through to default

//   // ── Add sessionState to the return so the UI can show the right banner ─────
//   return {
//     status,
//     variant,
//     details,
//     weightedMean: officialMean.toFixed(2),
//     sessionState: targetSession,
//     summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: lists.missing.length },
//     passedList:    lists.passed,
//     failedList:    lists.failed,
//     specialList:   lists.special,
//     missingList:   lists.missing,
//     incompleteList: lists.incomplete,
//   };
 
  // ----- previewPromotion
// export const previewPromotion = async (
//   programId:        string,
//   yearToPromote:    number,
//   academicYearName: string,
// ) => {
//   const nextYear = yearToPromote + 1;

//   // ── Resolve the target academic year document ──────────────────────────────
//   const targetYearDoc = await AcademicYear.findOne({ year: academicYearName }).lean();

//   if (!targetYearDoc) {
//     console.warn(`[previewPromotion] AcademicYear "${academicYearName}" not found.`);
//     return { totalProcessed: 0, eligibleCount: 0, blockedCount: 0, eligible: [], blocked: [] };
//   }

//   // ── COHORT SCOPING ─────────────────────────────────────────────────────────
//   // Step 1: Students admitted THIS year (new intake for Year 1, or re-entrants)
//   const admissionStudents = await Student.find({
//     program:               programId,
//     currentYearOfStudy:    yearToPromote,
//     admissionAcademicYear: targetYearDoc._id,
//   }).lean();

//   // Step 2: Students from earlier cohorts who have marks in this specific year
//   // (repeaters, stayout students sitting this year's ordinary exams)
//   const [marksThisYear, directMarksThisYear] = await Promise.all([
//     Mark.distinct("student", { academicYear: targetYearDoc._id }),
//     MarkDirect.distinct("student", { academicYear: targetYearDoc._id }),
//   ]);

//   const markedStudentIds = new Set<string>([
//     ...marksThisYear.map((id: any) => id.toString()),
//     ...directMarksThisYear.map((id: any) => id.toString()),
//   ]);

//   const admissionIds = new Set(admissionStudents.map((s) => s._id.toString()));

//   const returningStudents = await Student.find({
//     program:            programId,
//     currentYearOfStudy: yearToPromote,
//     _id: {
//       $in:  Array.from(markedStudentIds),
//       $nin: Array.from(admissionIds),
//     },
//   }).lean();

//   // Step 3: Admin-status students (on_leave, deferred) belonging to this cohort
//   const adminStudents = await Student.find({
//     program:            programId,
//     currentYearOfStudy: yearToPromote,
//     status:             { $in: ["on_leave", "deferred"] },
//     $or: [
//       { admissionAcademicYear: targetYearDoc._id },
//       { "academicHistory.academicYear": academicYearName },
//     ],
//     _id: {
//       $nin: [
//         ...Array.from(admissionIds),
//         ...returningStudents.map((s) => s._id.toString()),
//       ],
//     },
//   }).lean();

//   const allStudents = [...admissionStudents, ...returningStudents, ...adminStudents];

//   // ── Admin status map ───────────────────────────────────────────────────────
//   const ADMIN_STATUSES: Record<string, string> = {
//     on_leave:     "ACADEMIC LEAVE",
//     deferred:     "DEFERMENT",
//     discontinued: "DISCONTINUED",
//     deregistered: "DEREGISTERED",
//     graduated:    "GRADUATED",
//   };

//   const eligible: any[] = [];
//   const blocked:  any[] = [];

//   for (const student of allStudents) {
//     const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
//     const adminLabel        = ADMIN_STATUSES[student.status];

//     // ── Already promoted ───────────────────────────────────────────────────
//     if (isAlreadyPromoted) {
//       eligible.push({
//         id:            student._id,
//         regNo:         student.regNo,
//         name:          student.name,
//         status:        "ALREADY PROMOTED",
//         reasons:       [],
//         specialGrounds: "",
//         summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//         remarks:       student.remarks,
//         academicLeavePeriod: student.academicLeavePeriod,
//         details:       "",
//       });
//       continue;
//     }

//     // ── Admin locked status — skip engine ──────────────────────────────────
//     if (adminLabel) {
//       const leaveType = student.academicLeavePeriod?.type?.toUpperCase();
//       const reason    = leaveType ? `${adminLabel} (${leaveType})` : adminLabel;

//       // Build specialGrounds from admin student too (they may have specials
//       // recorded from before they went on leave)
//       const adminGrounds = [
//         (student.academicLeavePeriod?.type || "").toLowerCase(),
//         (student.remarks || "").toLowerCase(),
//       ].join(" ").trim() || "other";

//       blocked.push({
//         id:            student._id,
//         regNo:         student.regNo,
//         name:          student.name,
//         status:        adminLabel,
//         reasons:       [reason],
//         specialGrounds: adminGrounds,
//         summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//         academicLeavePeriod: student.academicLeavePeriod,
//         remarks:       student.remarks,
//         details:       "",
//       });
//       continue;
//     }

//     // ── Active student — run engine ────────────────────────────────────────
//     const statusResult = await calculateStudentStatus(
//       student._id,
//       programId,
//       academicYearName,
//       yearToPromote,
//       { forPromotion: true },
//     );

//     // ── specialGrounds: collected from ALL available sources ───────────────
//     // This is the field that drives the special list generation in promote.ts.
//     // It MUST be non-empty for the special list filter to work.
//     const specialGrounds = (() => {
//       // Source 1: grounds recorded on each special list entry (from mark remarks)
//       const fromList = (statusResult.specialList || [])
//         .map((s: { grounds?: string }) => (s.grounds || "").toLowerCase())
//         .join(" ");

//       // Source 2: student.remarks (e.g. "Special Granted: Financial")
//       const fromRemarks = (student.remarks || "").toLowerCase();

//       // Source 3: academicLeavePeriod.type (financial/compassionate — often
//       // correlates with special exam ground for the same student)
//       const fromLeave = (student.academicLeavePeriod?.type || "").toLowerCase();

//       const combined = `${fromList} ${fromRemarks} ${fromLeave}`.trim();

//       // Fall back to "other" so the catch-all filter always picks up
//       // students whose ground was recorded as "Administrative" or left blank
//       return combined || "other";
//     })();

//     const report: any = {
//       id:      student._id,
//       regNo:   student.regNo,
//       name:    student.name,
//       status:  statusResult.status,
//       summary: statusResult.summary,
//       reasons: [] as string[],
//       remarks: student.remarks,
//       academicLeavePeriod: student.academicLeavePeriod,
//       details: statusResult.details,
//       specialGrounds,                       // ← THE CRITICAL FIELD
//       isEligibleForSupp:
//         !["STAYOUT", "REPEAT YEAR", "DEREGISTERED"].includes(statusResult.status) &&
//         (statusResult.failedList.length > 0 || statusResult.specialList.length > 0),
//     };

//     if (statusResult.status === "PASS") {
//       eligible.push(report);
//     } else {
//       // Build reasons list
//       if (statusResult.status === "STAYOUT")
//         report.reasons.push("ENG 15h: Failures > 1/3 (Must Stay Out)");
//       if (statusResult.status === "REPEAT YEAR")
//         report.reasons.push("ENG 16: Failures >= 1/2 (Must Repeat)");
//       if ((statusResult as any).leaveDetails)
//         report.reasons.push(`${statusResult.status}: ${(statusResult as any).leaveDetails}`);
//       if (statusResult.specialList.length > 0)
//         report.reasons.push(
//           ...statusResult.specialList.map((s: any) => `${s.displayName} (SPECIAL)`),
//         );
//       if (statusResult.incompleteList.length)
//         report.reasons.push(
//           ...statusResult.incompleteList.map((u: string) => `${u} (INCOMPLETE)`),
//         );
//       if (statusResult.missingList.length)
//         report.reasons.push(
//           ...statusResult.missingList.map((u: string) => `${u} (MISSING)`),
//         );
//       if (statusResult.failedList.length)
//         report.reasons.push(
//           ...statusResult.failedList.map(
//             (f: any) => `${f.displayName} (FAIL ATTEMPT: ${f.attempt})`,
//           ),
//         );

//       blocked.push(report);
//     }
//   }

//   return {
//     totalProcessed: allStudents.length,
//     eligibleCount:  eligible.length,
//     blockedCount:   blocked.length,
//     eligible,
//     blocked,
//   };
// };

// -----previewPromotion


// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId).populate("program");
//   if (!student) throw new Error("Student not found");

//   // If already non-active, don't re-process unless explicitly requested
//   if (student.status !== "active")
//     return {
//       success: false,
//       message: `Promotion blocked: Student status is ${student.status}`,
//     };

//   // Prevent promoting someone already out of the system unless they were readmitted
//   if (["deregistered", "discontinued", "graduated"].includes(student.status)) {
//     return {
//       success: false,
//       message: `Action blocked: Student is currently ${student.status}`,
//     };
//   }

//   const auditResult = await performAcademicAudit(studentId);
//   if (auditResult.discontinued)
//     return { success: false, message: `Discontinued: ${auditResult.reason}` };

//   const program = student.program as any;
//   const duration = program.durationYears || 4;
//   const actualCurrentYear = student.currentYearOfStudy || 1;
//   const currentSession = await AcademicYear.findOne({ isCurrent: true }).lean();
//   const completedYearLabel = currentSession?.year || "N/A";

//   // 1. Run the audit for the year they just finished
//   const statusResult = await calculateStudentStatus( student._id, student.program, completedYearLabel, actualCurrentYear, { forPromotion: true } );

//   // --- AUTO-SYNC TERMINAL STATUSES ---
//   const terminalStatuses = ["DEREGISTERED", "REPEAT YEAR", "DISCONTINUED"];
//   if (terminalStatuses.includes(statusResult.status)) {
//     await syncTerminalStatusToDb( studentId, statusResult.status, statusResult.details, completedYearLabel );

//     return {
//       success: false,
//       message: `Promotion Blocked: Student status updated to ${statusResult.status}`,
//       details: statusResult,
//     };
//   }

//   // --- EXISTING PROMOTION LOGIC ---
//   if (statusResult?.status === "PASS") {
//     const rawMean = parseFloat(statusResult.weightedMean);
//     const yearWeight = getYearWeight( program, student.entryType || "Direct", actualCurrentYear );
//     const weightedContribution = rawMean * yearWeight;

//     const historyRecord = {
//       academicYear: completedYearLabel,
//       yearOfStudy: actualCurrentYear,
//       annualMeanMark: rawMean,
//       weightedContribution: weightedContribution,
//       unitsTakenCount: statusResult.summary.totalExpected,
//       failedUnitsCount: statusResult.summary.failed,
//       isRepeatYear: false,
//       date: new Date(),
//     };

//     if (actualCurrentYear === duration) {
//       const fullHistory = [...(student.academicHistory || []), historyRecord];
//       const finalWAA = fullHistory.reduce(
//         (acc, curr) => acc + (curr.weightedContribution || 0),
//         0,
//       );

//       let classification = "PASS";
//       if (finalWAA >= 70) classification = "FIRST CLASS HONOURS";
//       else if (finalWAA >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (finalWAA >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

//       await Student.findByIdAndUpdate(studentId, {
//         $set: {
//           status: "graduand",
//           finalWeightedAverage: finalWAA.toFixed(2),
//           classification: classification,
//           graduationYear: new Date().getFullYear(),
//           currentYearOfStudy: actualCurrentYear + 1,
//           currentSemester: 1,
//         },
//         $push: { academicHistory: historyRecord },
//       });
//       return {
//         success: true,
//         message: `Student completed final year. Classified as ${classification}`,
//         isGraduation: true,
//       };
//     }

//     const nextYear = actualCurrentYear + 1;
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { currentYearOfStudy: nextYear, currentSemester: 1 },
//       $push: {
//         promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date() },
//         academicHistory: historyRecord,
//       },
//     });
//     return {
//       success: true,
//       message: `Successfully promoted to Year ${nextYear}`,
//     };
//   }

  

//   // Handle Repeat Year/Stay Out (Optional: also sync status to DB)
//   let blockMessage = `Promotion Blocked: `;
//   if (statusResult?.status === "REPEAT YEAR") {
//     // Sync to DB so they appear as 'repeat' status
//     await Student.findByIdAndUpdate(studentId, { $set: { status: "repeat" } });
//     blockMessage +=
//       "Student is required to repeat the year based on academic performance.";
//   } else if (statusResult?.status === "SPECIALS PENDING")
//     blockMessage += "Student has pending Special Examinations.";
//   else blockMessage += `Current status is '${statusResult?.status}'.`;

//   return { success: false, message: blockMessage, details: statusResult };
// };

// ---- promote 2 ----
// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId).populate("program");
//   if (!student) throw new Error("Student not found");
 
//   if (student.status !== "active") {
//     return {
//       success: false,
//       message: `Promotion blocked: Student status is ${student.status}`,
//     };
//   }
 
//   if (["deregistered", "discontinued", "graduated"].includes(student.status)) {
//     return {
//       success: false,
//       message: `Action blocked: Student is currently ${student.status}`,
//     };
//   }
 
//   const auditResult = await performAcademicAudit(studentId);
//   if (auditResult.discontinued) {
//     return { success: false, message: `Discontinued: ${auditResult.reason}` };
//   }
 
//   const program         = student.program as any;
//   const duration        = program.durationYears || 5;
//   const actualCurrentYear = student.currentYearOfStudy || 1;
//   const currentSession  = await AcademicYear.findOne({ isCurrent: true }).lean();
//   const completedYearLabel = currentSession?.year || "N/A";
 
//   const statusResult = await calculateStudentStatus(
//     student._id,
//     student.program,
//     completedYearLabel,
//     actualCurrentYear,
//     { forPromotion: true },
//   );
 
//   // ── Terminal status sync ─────────────────────────────────────────────────
//   const terminalStatuses = ["DEREGISTERED", "REPEAT YEAR", "DISCONTINUED"];
//   if (terminalStatuses.includes(statusResult.status)) {
//     await syncTerminalStatusToDb(
//       studentId,
//       statusResult.status,
//       statusResult.details,
//       completedYearLabel,
//     );
//     return {
//       success: false,
//       message: `Promotion Blocked: Student status updated to ${statusResult.status}`,
//       details: statusResult,
//     };
//   }
 
//   // ── PASS path ────────────────────────────────────────────────────────────
//   if (statusResult.status === "PASS") {
//     const rawMean          = parseFloat(statusResult.weightedMean);
//     const yearWeight       = getYearWeight(program, student.entryType || "Direct", actualCurrentYear);
//     const weightedContrib  = rawMean * yearWeight;
 
//     const historyRecord = {
//       academicYear:         completedYearLabel,
//       yearOfStudy:          actualCurrentYear,
//       annualMeanMark:       rawMean,
//       weightedContribution: weightedContrib,
//       unitsTakenCount:      statusResult.summary.totalExpected,
//       failedUnitsCount:     statusResult.summary.failed,
//       isRepeatYear:         false,
//       date:                 new Date(),
//     };
 
//     if (actualCurrentYear === duration) {
//       // GRADUATION
//       const fullHistory = [...(student.academicHistory || []), historyRecord];
//       const finalWAA    = fullHistory.reduce((acc, h) => acc + (h.weightedContribution || 0), 0);
 
//       let classification = "PASS";
//       if (finalWAA >= 70)      classification = "FIRST CLASS HONOURS";
//       else if (finalWAA >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (finalWAA >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
 
//       await Student.findByIdAndUpdate(studentId, {
//         $set: {
//           status:              "graduand",
//           finalWeightedAverage: finalWAA.toFixed(2),
//           classification,
//           graduationYear:      new Date().getFullYear(),
//           currentYearOfStudy:  actualCurrentYear + 1,
//           currentSemester:     1,
//           qualifierSuffix:     "", // clear on graduation
//         },
//         $push: { academicHistory: historyRecord },
//       });
 
//       return {
//         success:      true,
//         message:      `Student completed final year. Classified as ${classification}`,
//         isGraduation: true,
//       };
//     }
 
//     // Normal promotion — no failed units at all
//     const nextYear = actualCurrentYear + 1;
 
//     await Student.findByIdAndUpdate(studentId, {
//       $set: {
//         currentYearOfStudy: nextYear,
//         currentSemester:    1,
//         qualifierSuffix:    "", // clear qualifier — clean pass
//       },
//       $push: {
//         promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date() },
//         academicHistory:  historyRecord,
//         statusHistory: {
//           status:         "active",
//           previousStatus: "active",
//           date:           new Date(),
//           reason:         `Promoted to Year ${nextYear}`,
//         },
//       },
//     });
 
//     return { success: true, message: `Successfully promoted to Year ${nextYear}` };
//   }
 
//   // ── SUPP path — check carry-forward eligibility ──────────────────────────
//   //
//   // The status engine returns "SUPP N" for N failed units.
//   // We check if carry-forward applies (max 2 units, no missing-CA units).
//   //
//   // ENG.15d: "A candidate... who fails more than two units... at the
//   //   supplementary examinations period shall not be allowed to proceed
//   //   to the next year of study"  → STAYOUT (not carry-forward)
 
//   const suppMatch = statusResult.status.match(/SUPP\s+(\d+)/);
//   if (suppMatch) {
//     const failedCount = parseInt(suppMatch[1]);
 
//     if (failedCount > 2) {
//       // ENG.15d: STAYOUT — stays at same year
//       return {
//         success: false,
//         message: `STAYOUT: Failed ${failedCount} supplementary units. Max for carry-forward is 2 (ENG.15d).`,
//         details: statusResult,
//       };
//     }
 
//     // 1 or 2 failed at supp → check carry-forward eligibility
//     const cfResult = await assessCarryForward(
//       studentId,
//       student.program.toString(),
//       completedYearLabel,
//       actualCurrentYear,
//     );
 
//     if (!cfResult.eligible) {
//       return {
//         success: false,
//         message: `Cannot carry forward: ${cfResult.reason}`,
//         details: statusResult,
//       };
//     }
 
//     // Apply carry-forward promotion
//     await applyCarryForward(
//       studentId,
//       student.program.toString(),
//       completedYearLabel,
//       actualCurrentYear,
//       cfResult.units,
//     );
 
//     return {
//       success:  true,
//       message:  `Carry-forward applied. Promoted to Year ${actualCurrentYear + 1} with qualifier ${cfResult.qualifier}. Units to retake: ${cfResult.units.map((u) => u.unitCode).join(", ")}`,
//       isCarryForward: true,
//       cfUnits:  cfResult.units,
//       qualifier: cfResult.qualifier,
//     };
//   }
 
//   // ── STAYOUT explicit path ────────────────────────────────────────────────
//   if (statusResult.status === "STAYOUT") {
//     // ENG.15h: Student stays at same year, retakes ONLY in next ORDINARY period.
//     // DB status stays "active" — they are not deregistered or discontinued.
//     // No qualifier is set on the student here; the ATTEMPT column on next
//     // year's ordinary scoresheet will show A/SO based on context.
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { status: "active" as any },
//       $push: {
//         statusEvents: {
//           fromStatus:   "active",
//           toStatus:     "active",
//           date:         new Date(),
//           academicYear: completedYearLabel,
//           reason:       `ENG.15h STAYOUT: ${statusResult.details}`,
//         },
//       },
//     });
 
//     return {
//       success: false,
//       message: `STAYOUT: ${statusResult.details}`,
//       details: statusResult,
//     };
//   }
 
//   // Fallback
//   return {
//     success: false,
//     message: `Promotion Blocked: Current status is '${statusResult.status}'.`,
//     details: statusResult,
//   };
// };

// ----- promote 2 ----


// export const bulkPromoteClass = async ( programId: string, yearToPromote: number, academicYearName: string ) => {
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active"});
//   const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[]};
//   for (const student of students) {
//     try {
//       const studentId = (student._id as any).toString();
//       if (student.currentYearOfStudy >= nextYear) { results.alreadyPromoted++; results.promoted++; continue; }
//       const res = await promoteStudent(studentId);
//       if (res.success) results.promoted++;
//       else results.failed++;
//     } catch (err: any) {
//       results.errors.push(`${student.regNo}: ${err.message}`);
//     }
//   }
//   return results;
// };

//  two

// // serverside/src/services/statusEngine.ts
// import mongoose from "mongoose";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import AcademicYear from "../models/AcademicYear";
// import Student from "../models/Student";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { getYearWeight } from "../utils/weightingRegistry";
// import { performAcademicAudit } from "./academicAudit";

// const getLetterGrade = (mark: number, settings: any): string => {
//   if (!settings || !settings.gradingScale) {
//     // Fallback if settings are missing
//     if (mark >= 69.5) return "A"; if (mark >= 59.5) return "B"; if (mark >= 49.5) return "C"; if (mark >= 39.5) return "D"; return "E";
//   }

//   // Sort scale descending (e.g., 70, 60, 50...) to find the first match
//   const sortedScale = [...settings.gradingScale].sort((a, b) => b.min - a.min);
//   const matched = sortedScale.find((s) => mark >= s.min);
//   return matched ? matched.grade : settings.failingGrade || "E";
// };

// export interface StudentStatusResult {
//   status: string;
//   variant: "success" | "warning" | "error" | "info";
//   details: string;
//   weightedMean: string;
//   summary: { totalExpected: number; passed: number; failed: number; missing: number; };
//   passedList: { code: string; mark: number }[];
//   failedList: { displayName: string; attempt: number }[];
//   specialList: { displayName: string; grounds: string }[];
//   missingList: string[];
//   incompleteList: string[];
//   leaveDetails?: string;
// }

// export const calculateStudentStatus = async (studentId: any, programId: any, academicYearName: string, yearOfStudy: number = 1): Promise<StudentStatusResult> => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");
//   const passMark = settings?.passMark || 40;

//   const student = await Student.findById(studentId).lean();
//   if (!student) throw new Error("Student not found");

//   // If student is on leave/deferred, return immediately
//   if (["ACADEMIC LEAVE", "DEFERMENT"].includes(student.status)) {
//     return {
//       status: student.status,
//       variant: "info",
//       details: `Student is on ${student.status.toLowerCase()}.`,
//       weightedMean: "0.00",
//       summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
//       leaveDetails: student.remarks || "No reason provided."
//     };
//   }

//   const curriculum = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean() as any[];

//   const grades = await FinalGrade.find({ student: studentId }).populate({ path: "programUnit", populate: { path: "unit" } }).lean() as any[];

//   const unitResults = new Map();
//   grades.sort((a, b) => new Date(a.updatedAt).getTime() - new Date(b.updatedAt).getTime());

//   grades.forEach((g) => {
//     if (!g.programUnit || !g.programUnit.unit) { console.warn( `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`, ); return; }
//     const unitCode = g.programUnit?.unit?.code?.toUpperCase();
//     if (!unitCode) return;
//     const numericMark = g.agreedMark ?? g.totalMark ?? 0;

//     const isSpecial = g.isSpecial || g.status === "SPECIAL";
//     unitResults.set(unitCode, {
//       mark: Number(numericMark),
//       status: g.status,
//       attempt: parseInt(g.attempt) || g.attemptNumber || 1,
//       isSpecial: isSpecial,
//       isSpecial: g.isSpecial || g.status === "SPECIAL" || g.remarks?.toLowerCase().includes("financial") || g.remarks?.toLowerCase().includes("compassionate"),
//       remarks: g.remarks || ""
//     });
//   });

//   const lists = {
//     passed: [] as { code: string; mark: number }[],
//     failed: [] as { displayName: string; attempt: number }[],
//     special: [] as { displayName: string; grounds: string }[],
//     missing: [] as string[], incomplete: [] as string[]
//   };

//   let totalFirstAttemptSum = 0;

//   curriculum.forEach((pUnit) => {
//     const code = pUnit.unit?.code?.toUpperCase();
//     const displayName = `${code}: ${pUnit.unit?.name}`;
//     const record = unitResults.get(code);

//     if (!record) {
//       lists.missing.push(displayName);
//     } else if (record.isSpecial) {
//       lists.special.push({ displayName, grounds: record.remarks || "Special Grounds" });
//     } else if (record.mark === 0 || record.status === "INCOMPLETE") {
//       lists.incomplete.push(displayName);
//     } else if (record.mark >= passMark) {
//       if (record.attempt === 1) totalFirstAttemptSum += record.mark;
//       lists.passed.push({ code, mark: record.mark });
//     } else {
//       if (record.attempt === 1) totalFirstAttemptSum += record.mark;
//       lists.failed.push({ displayName, attempt: record.attempt });
//     }
//   });

//   const totalUnits = curriculum.length;
//   const failCount = lists.failed.length;
//   const meanMark = totalUnits > 0 ? totalFirstAttemptSum / totalUnits : 0;

//   let status = "IN GOOD STANDING";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = "Proceed to next year.";

//   if (lists.missing.length >= 6) {
//     status = "DEREGISTERED"; variant = "error"; details = "Absent from 6+ examinations (ENG 23c).";
//   } else if (failCount >= totalUnits / 2 || meanMark < 40) {
//     status = "REPEAT YEAR"; variant = "error"; details = "Failed >= 50% units or Mean < 40% (ENG 16).";
//   } else if (failCount > totalUnits / 3) {
//     status = "STAYOUT"; variant = "warning"; details = "Failed > 1/3 of units. Retake units next year (ENG 15h).";
//   } else if (failCount > 0) {
//     status = "SUPPLEMENTARY"; variant = "warning"; details = "Eligible for supplementaries (ENG 13a).";
//   } else if (lists.special.length > 0) {
//     status = "SPECIALS PENDING"; variant = "info"; details = "Awaiting special examinations.";
//   }

//   return {
//     status, variant, details,
//     weightedMean: meanMark.toFixed(2),
//     summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: lists.missing.length },
//     passedList: lists.passed, failedList: lists.failed, specialList: lists.special, missingList: lists.missing, incompleteList: lists.incomplete
//   };
// };

// export const previewPromotion = async (programId: string, yearToPromote: number, academicYearName: string) => {
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active" }).lean();
//   const allStudents = await Student.find({ program: programId, currentYearOfStudy: yearToPromote }).lean();

//   const eligible: any[] = [];
//   const blocked: any[] = [];

//   for (const student of allStudents) {
//     const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
//     const statusResult = await calculateStudentStatus(student._id, programId, academicYearName, yearToPromote);

//     const report = {
//       id: student._id, regNo: student.regNo, name: student.name,
//       status: isAlreadyPromoted ? "ALREADY PROMOTED" : statusResult.status,
//       summary: statusResult.summary, reasons: [] as string[], isAlreadyPromoted
//     };

//     // Promotion criteria: Must be In Good Standing AND not already moved
//     if (!isAlreadyPromoted && statusResult.status === "IN GOOD STANDING") {
//       eligible.push(report);
//     } else if (isAlreadyPromoted) {
//       eligible.push(report); // Keep already promoted in eligible list but marked accordingly
//     } else {
//       // Mapping the new lists to reasons
//       if (statusResult.leaveDetails) report.reasons.push(`${statusResult.status}: ${statusResult.leaveDetails}`);
//       if (statusResult.specialList.length) report.reasons.push(...statusResult.specialList.map(s => `${s.displayName} (SPECIAL)`));
//       if (statusResult.incompleteList.length) report.reasons.push(...statusResult.incompleteList.map(u => `${u} (INCOMPLETE)`));
//       if (statusResult.missingList.length) report.reasons.push(...statusResult.missingList.map(u => `${u} (MISSING)`));
//       if (statusResult.failedList.length) report.reasons.push(...statusResult.failedList.map(f => `${f.displayName} (FAIL ATTEMPT: ${f.attempt})`));

//       blocked.push(report);
//     }
//   }

//   return { totalProcessed: allStudents.length, eligibleCount: eligible.length, blockedCount: blocked.length, eligible, blocked };
// };

// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId).populate("program");
//   if (!student) throw new Error("Student not found");

//   if (student.status !== "active") {
//     return { success: false, message: `Promotion blocked: Student status is ${student.status}` };
//   }

//   const auditResult = await performAcademicAudit(studentId);
//   if (auditResult.discontinued) {
//     return { success: false, message: `Discontinued: ${auditResult.reason}` };
//   }

//   const program = student.program as any;
//   const duration = program.durationYears || 4;
//   const actualCurrentYear = student.currentYearOfStudy || 1;

//   const statusResult = await calculateStudentStatus( student._id, student.program, "N/A", actualCurrentYear);

//   if (statusResult?.status === "IN GOOD STANDING") {
//     const rawMean = parseFloat(statusResult.weightedMean);
//     const yearWeight = getYearWeight(program, student.entryType || "Direct", actualCurrentYear);
//     const weightedContribution = rawMean * yearWeight;

//     const historyRecord = {
//       yearOfStudy: actualCurrentYear,
//       annualMeanMark: rawMean,
//       weightedContribution: weightedContribution,
//       unitsTakenCount: statusResult.summary.totalExpected,
//       failedUnitsCount: statusResult.summary.failed,
//       isRepeatYear: false
//     };

//     // CASE A: FINAL YEAR STUDENT (Graduation Path)
//     if (actualCurrentYear === duration) {
//       // Calculate Final Classification using the existing history + current year
//       const fullHistory = [...(student.academicHistory || []), historyRecord];
//       const finalWAA = fullHistory.reduce((acc, curr) => acc + (curr.weightedContribution || 0), 0);

//       let classification = "PASS";
//       if (finalWAA >= 70) classification = "FIRST CLASS HONOURS";
//       else if (finalWAA >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (finalWAA >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

//       await Student.findByIdAndUpdate(studentId, {
//         $set: { status: "graduand", finalWeightedAverage: finalWAA.toFixed(2), classification: classification, graduationYear: new Date().getFullYear() },
//         $push: { academicHistory: historyRecord }
//       });

//       return { success: true, message: `Student completed final year. Classified as ${classification}`, isGraduation: true };
//     }

//     // CASE B: NORMAL PROMOTION (N -> N+1)
//     const nextYear = actualCurrentYear + 1;
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { currentYearOfStudy: nextYear, currentSemester: 1 },
//       $push: { promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date() }, academicHistory: historyRecord }
//     });

//     return { success: true, message: `Successfully promoted to Year ${nextYear}` };
//   }

//   // Logic to protect students with Specials/Incompletes from being "Failed"
//   let blockMessage = `Promotion Blocked: `;
//   if (statusResult?.status === "SPECIALS PENDING") {
//     blockMessage += "Student has pending Special Examinations. These must be sat and graded before promotion.";
//   } else if (statusResult?.status === "REPEAT YEAR") {
//     blockMessage += "Student is required to repeat the year based on academic performance.";
//   } else {
//     blockMessage += `Current status is '${statusResult?.status}'.`;
//   }

//   return { success: false, message: blockMessage, details: statusResult };
// };

// export const bulkPromoteClass = async ( programId: string, yearToPromote: number, academicYearName: string ) => {
//   // 1. Get everyone currently in the year AND those already promoted to the next year
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active", });

//   const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };

//   for (const student of students) {
//     try {
//       // Cast to 'any' or use (student as any)._id to bypass the 'unknown' check
//       const studentId = (student._id as any).toString();

//       // Skip if already in the target year or higher
//       if (student.currentYearOfStudy >= nextYear) { results.alreadyPromoted++; results.promoted++; continue; }

//       const res = await promoteStudent(studentId);

//       if (res.success) results.promoted++;
//       else results.failed++;
//     } catch (err: any) {
//       results.errors.push(`${student.regNo}: ${err.message}`);
//     }
//   }

//   return results;
// };

// one
// export const calculateStudentStatus = async ( studentId: any, programId: any, academicYearName: string, yearOfStudy: number = 1 ) => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found. Please configure grading scales.");

//   // 1. Determine dispaly year for UI
//   let displayYearName = academicYearName;
//   if (!displayYearName || displayYearName === "N/A") {
//     const latestGrade = await FinalGrade.findOne({ student: studentId }).populate("academicYear").sort({ createdAt: -1 }); displayYearName = (latestGrade?.academicYear as any)?.year || "N/A"; }

//   // 1. Get Curriculum & explicitly type the populated unit
//   const curriculum = (await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean()) as any[];

//   if (!curriculum.length) {
//     return {
//       status: "NO CURRICULUM", variant: "info", details: `No units defined for Year ${yearOfStudy} in this program.`,
//       summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 }, passedList: [],
//     };
//   }

//   // 2. Get Grades (All history to track Retake -> Re-Retake lifecycle)
//   const grades = (await FinalGrade.find({ student: studentId }).populate({ path: "programUnit", populate: { path: "unit" } }).populate({ path: "academicYear", model: "AcademicYear" }).sort({ createdAt: -1 }).lean()) as any[];

//   //     console.log(`DEBUG: Found ${grades.length} grades for student.`);
//   // if (grades.length > 0) { console.log("Sample Grade AcademicYear Data:", JSON.stringify(grades[0].academicYear, null, 2)); }

//   // console.log("RE-CHECK Sample Grade AcademicYear Data:", grades[0]?.academicYear);

//   // 3. Map grades by UNIT CODE
//    const unitResults = new Map();
//   // const unitResults = new Map<string,  { status: string; attemptType: string; attemptNumber: number; totalMark: number; }>();

//   grades.forEach((g) => {
//     if (!g.programUnit || !g.programUnit.unit) { console.warn( `[StatusEngine] Skipping grade record ${g._id} - missing programUnit or unit`, ); return; }

//     const unitCode = g.programUnit.unit.code.toUpperCase();
//     // const unitCode = g.programUnit?.unit?.code?.toUpperCase();
//     if (!unitCode) return;
//     const existing = unitResults.get(unitCode);
//     if (existing?.status === "PASS") return;
//     unitResults.set(unitCode, { status: g.status, attemptType: g.attemptType, attemptNumber: g.attemptNumber, totalMark: g.totalMark, remarks: g.remarks || "" });
//   });

//   // Trackers for the Coordinator View
//   let totalMarksSum = 0;
//   const passedUnits: any[] = []; const failedList: string[] = []; const retakeUnits: string[] = []; const reRetakeUnits: string[] = [];
//   const missingUnits: string[] = [];  const incompleteUnits: string[] = []; const specialList: { displayName: string; grounds: string }[] = [];

//   // 4. Compare Curriculum against results
//   curriculum.forEach((pUnit: any) => {
//     const unitCode = pUnit.unit?.code?.trim().toUpperCase();
//     const unitName = pUnit.unit?.name;
//     const displayName = `${unitCode}: ${unitName}`;
//     const record = unitResults.get(unitCode);

//     if (!record) { missingUnits.push(displayName); totalMarksSum +=0;
//     } else if (record.status === "PASS") {
//       totalMarksSum += record.totalMark; const numericMark = record.totalMark || 0;
//       const letterGrade = getLetterGrade(numericMark, settings);
//       passedUnits.push({ code: unitCode,  name: unitName, mark: numericMark, grade: letterGrade });
//     } else if (record.status === "SPECIAL") {
//       let grounds = "Administrative";
//       if (record.remarks.toLowerCase().includes("financial")) grounds = "Financial";
//       if (record.remarks.toLowerCase().includes("compassionate")) grounds = "Compassionate";

//       specialList.push({ displayName, grounds });
//     } else if (record.status === "INCOMPLETE") {
//       incompleteUnits.push(displayName);
//     } else {
//       // Logic for failures based on attempts
//       if (record.attemptNumber >= 3) reRetakeUnits.push(displayName);
//       else if (record.attemptNumber === 2) retakeUnits.push(displayName);
//       else failedList.push(displayName);
//     }
//   });

//   const totalExpected = curriculum.length;
//   const totalFailed = failedList.length + retakeUnits.length + reRetakeUnits.length;
//   const weightedMean = totalMarksSum / totalExpected;

//   // 5. Determine UI Status
//   let status = "IN GOOD STANDING";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = `Year ${yearOfStudy} curriculum units cleared.`;

//   // ENG 23.c: Deregistration (Absent from 6+ exams)
//   if (missingUnits.length >= 6) {
//     // ENG 23.c
//     status = "DEREGISTERED"; variant = "error"; details = "Automatic deregistration: Absent from 6+ examinations.";
//   } else if (reRetakeUnits.length > 0) { status = "CRITICAL FAILURE";  variant = "error"; details = "Student failed a third attempt (Re-Retake).";
//   } else if (totalFailed >= totalExpected / 2 || weightedMean < 40) {
//     // ENG 16.a
//     status = "REPEAT YEAR"; variant = "error"; details = "Failed 50% or more units, or mean mark below 40%.";
//   } else if (totalFailed > totalExpected / 3) {
//     // ENG 15.h
//     status = "RETAKE REQUIRED"; variant = "warning"; details = "Failures exceed 1/3 of units. Must retake during ordinary sessions.";
//   } else if (totalFailed > 0) {
//     // ENG 13.a
//     status = "SUPPLEMENTARY PENDING"; variant = "warning"; details = `Eligible for supplementary exams in ${totalFailed} units.`;
//   } else if (specialList.length > 0) {
//     status = "SPECIAL EXAM PENDING"; variant = "info"; details = `Awaiting results for ${specialList.length} special exam(s).`;
//   } else if (missingUnits.length > 0) {
//     status = "INCOMPLETE DATA"; variant = "info"; details = "Some unit records are missing from the system.";
//   }

//   const sessionRecord = grades.find( (g) => g.programUnit?.requiredYear === yearOfStudy );

//   let actualSessionName = "N/A";

//   if (academicYearName && academicYearName !== "N/A") {
//     actualSessionName = academicYearName;
//   } else {
//     // 1. Try to find a grade in the CURRENT year of study that has a valid year
//     const yearSpecificGrade = grades.find((g) => g.programUnit?.requiredYear === yearOfStudy && g.academicYear?.year );

//     if (yearSpecificGrade?.academicYear?.year) {
//       actualSessionName = yearSpecificGrade.academicYear.year;
//     } else if (grades.length > 0) {
//       // 2. If Year 2 is empty, but Year 1 exists, we "guess" the next year
//       const previousYearGrade = grades.find((g) => g.academicYear?.year);
//       if (previousYearGrade?.academicYear?.year) {
//         const baseYear = previousYearGrade.academicYear.year; // e.g., "2023/2024"
//         // Simple logic to increment if we are looking at a higher year of study
//         if (yearOfStudy > (previousYearGrade.programUnit?.requiredYear || 1)) {
//           // Optional: Add logic to increment "2023/2024" to "2024/2025"
//           actualSessionName = baseYear; // For now, keep as base to avoid wrong guesses
//         } else {
//           actualSessionName = baseYear;
//         }
//       }
//     }
//   }

//   return {
//     status, variant, details,
//     academicYearName: actualSessionName, yearOfStudy: yearOfStudy, weightedMean: weightedMean.toFixed(2),
//     summary: { totalExpected: curriculum.length, passed: passedUnits.length, failed: totalFailed, missing: missingUnits.length },
//     missingList: missingUnits, passedList: passedUnits, failedList: failedList, retakeList: retakeUnits, reRetakeList: reRetakeUnits, specialList: specialList, incompleteList: incompleteUnits,
//   };
// };

// changes

// export const previewPromotion = async ( programId: string, yearToPromote: number, academicYearName: string ) => {
//  // 1. Fetch students in the targeted year AND those already in the next year
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active", }).lean();

//   const eligible: any[] = [];
//   const blocked: any[] = [];

//   for (const student of students) {
//     const isAlreadyPromoted = student.currentYearOfStudy === nextYear;
//     const statusResult = await calculateStudentStatus( student._id, programId, academicYearName, yearToPromote );

//     const report = {
//       id: student._id, regNo: student.regNo, name: student.name,
//       status: isAlreadyPromoted ? "ALREADY PROMOTED" : (statusResult?.status || "PENDING"),
//       summary: statusResult?.summary, reasons: [] as string[], isAlreadyPromoted
//     };

//     if (isAlreadyPromoted || statusResult?.status === "IN GOOD STANDING") {
//       eligible.push(report);
//     } else {
//       // Add specific reasons for the block
//       if (statusResult?.specialList?.length) {
//         report.reasons.push( ...statusResult.specialList.map((item) => `${item.displayName} - SPECIAL: ${item.grounds} Grounds`));
//       }
//       if (statusResult?.incompleteList?.length) {
//         report.reasons.push(...statusResult.incompleteList.map(u => `${u} - INCOMPLETE`));
//       }
//       if (statusResult?.missingList?.length) {
//         report.reasons.push(...statusResult.missingList.map(u => `MISSING: ${u}`));
//       }
//       if (statusResult?.failedList?.length) {
//         report.reasons.push(...statusResult.failedList.map(u => `FAILED: ${u}`));
//       }
//       if (statusResult?.retakeList?.length) {
//         report.reasons.push(...statusResult.retakeList.map(u => `RETAKE FAILED: ${u}`));
//       }
//       if (statusResult?.reRetakeList?.length) {
//         report.reasons.push(...statusResult.reRetakeList.map(u => `CRITICAL RE-RETAKE FAILED: ${u}`));
//       }
//       blocked.push(report);
//     }
//   }

//   return { totalProcessed: students.length, eligibleCount: eligible.length, blockedCount: blocked.length, eligible, blocked };
// };

// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId);
//   if (!student) throw new Error("Student not found");

//   // 1. Get the current active academic year for context
//   const actualCurrentYear = student.currentYearOfStudy || 1;

//   // 2. Calculate status based on the student's CURRENT year of study
//   const statusResult = await calculateStudentStatus(
//     student._id, student.program,
//     // "",
//     "N/A", actualCurrentYear, );

//   // 3. Promotion Guard: Only "IN GOOD STANDING" can move up
//   if (statusResult?.status === "IN GOOD STANDING") {
// const nextYear = actualCurrentYear + 1;

//     // Check if the student has reached the maximum years for their program (Optional)
//     // if (nextYear > student.programDuration) return { success: false, message: "Student has completed final year." };

//     await Student.findByIdAndUpdate(studentId, {
//       $set: { currentYearOfStudy: nextYear, currentSemester: 1, },
//       // Log the promotion in history if you have a history field
//       $push: { promotionHistory: { from: actualCurrentYear, to: nextYear, date: new Date() }   }
//     });

//     return { success: true, message: `Successfully promoted to Year ${nextYear}` };
//   }

//    return { success: false, message: `Promotion Blocked: Student is '${statusResult?.status}' for Year ${actualCurrentYear}.`, details: statusResult };
// };






















// // serverside/src/services/statusEngine.ts
// import mongoose from "mongoose";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import Student from "../models/Student";
// import Mark from "../models/Mark";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { getYearWeight } from "../utils/weightingRegistry";
// import { performAcademicAudit } from "./academicAudit";
// import { resolveStudentStatus } from "../utils/studentStatusResolver";
// import MarkDirect from "../models/MarkDirect";
// import AcademicYear from "../models/AcademicYear";
// import { getAttemptLabel, REG_QUALIFIERS } from "../utils/academicRules";

// // ─── Grade helper ─────────────────────────────────────────────────────────────

// const getLetterGrade = (mark: number, settings: any): string => {
//   if (!settings?.gradingScale) {
//     if (mark >= 69.5) return "A";
//     if (mark >= 59.5) return "B";
//     if (mark >= 49.5) return "C";
//     if (mark >= 39.5) return "D";
//     return "E";
//   }
//   const sorted = [...settings.gradingScale].sort((a: any, b: any) => b.min - a.min);
//   return sorted.find((s: any) => mark >= s.min)?.grade ?? "E";
// };

// // ─── Terminal DB sync ─────────────────────────────────────────────────────────
// // Called after calculateStudentStatus returns a terminal result.
// // Sets DB status + qualifierSuffix + audit trail.

// const syncTerminalStatusToDb = async (
//   studentId:    string,
//   engineStatus: string,
//   details:      string,
//   academicYear: string,
// ): Promise<void> => {
//   type Mapping = { dbStatus: string; qualifierFn: (s: any) => string };

//   const terminalMap: Record<string, Mapping> = {
//     "DEREGISTERED": {
//       dbStatus:    "deregistered",
//       qualifierFn: () => "",                    // no qualifier on reg — just DB status
//     },
//     "REPEAT YEAR": {
//       dbStatus:    "repeat",
//       qualifierFn: (student) => {               // RP1, RP2 etc.
//         const count = ((student.academicHistory || []) as any[]).filter(
//           (h: any) => h.isRepeatYear,
//         ).length + 1;
//         return REG_QUALIFIERS.repeatYear(count);
//       },
//     },
//     "STAYOUT": {
//       dbStatus:    "active",                    // stayout keeps DB status "active"
//       qualifierFn: () => "",                    // qualifier shown in scoresheet only
//     },
//     "DISCONTINUED": {
//       dbStatus:    "discontinued",
//       qualifierFn: () => "",
//     },
//   };

//   const entry = terminalMap[engineStatus];
//   if (!entry) return;

//   const student = await Student.findById(studentId).lean();
//   if (!student) return;

//   // Only sync if there is a status change (prevents re-triggering repeat year sync)
//   if (entry.dbStatus !== "active" && (student as any).status === entry.dbStatus) return;

//   const fromStatus  = (student as any).status;
//   const newQualifier = entry.qualifierFn(student);

//   const updatePayload: any = {
//     $set: {
//       status:  entry.dbStatus,
//       remarks: details,
//     },
//     $push: {
//       statusEvents: {
//         fromStatus,
//         toStatus:     entry.dbStatus,
//         date:         new Date(),
//         reason:       `Auto-Sync: ${details}`,
//         academicYear,
//       },
//       statusHistory: {
//         status:         entry.dbStatus,
//         previousStatus: fromStatus,
//         date:           new Date(),
//         reason:         details,
//       },
//     },
//   };

//   // Only update qualifier if it changes (don't wipe RP1C by syncing STAYOUT again)
//   if (newQualifier) {
//     updatePayload.$set.qualifierSuffix = newQualifier;
//   }

//   await Student.findByIdAndUpdate(studentId, updatePayload);
// };

// // ─── Result type ──────────────────────────────────────────────────────────────

// export interface StudentStatusResult {
//   status:       string;
//   variant:      "success" | "warning" | "error" | "info";
//   details:      string;
//   weightedMean: string;
//   sessionState: string;
//   summary: {
//     totalExpected: number;
//     passed:        number;
//     failed:        number;
//     missing:       number;
//     isOnLeave?:    boolean;
//   };
//   passedList:    { code: string; mark: number }[];
//   failedList:    { displayName: string; attempt: string | number }[];
//   specialList:   { displayName: string; grounds: string }[];
//   missingList:   string[];
//   incompleteList: string[];
//   leaveDetails?: string;
// }

// // ─── calculateStudentStatus ───────────────────────────────────────────────────

// export const calculateStudentStatus = async (
//   studentId:        any,
//   programId:        any,
//   academicYearName: string,
//   yearOfStudy:      number = 1,
//   options:          { forPromotion?: boolean } = {},
// ): Promise<StudentStatusResult> => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found.");

//   const passMark = (settings as any).passMark || 40;
//   const student  = await Student.findById(studentId).lean();
//   if (!student) throw new Error("Student not found");

//   // ── Terminal status gate ──────────────────────────────────────────────────
//   // These students never reach the mark engine regardless of session or flags.
//   const TERMINAL_STATUSES: Record<string, { label: string; variant: "info" | "error" | "success" | "warning" }> = {
//     on_leave:     { label: "ACADEMIC LEAVE", variant: "info"    },
//     deferred:     { label: "DEFERMENT",      variant: "info"    },
//     discontinued: { label: "DISCONTINUED",   variant: "error"   },
//     deregistered: { label: "DEREGISTERED",   variant: "error"   },
//     graduated:    { label: "GRADUATED",      variant: "success" },
//     graduand:     { label: "GRADUATED",      variant: "success" },
//   };

//   const terminalEntry = TERMINAL_STATUSES[(student as any).status ?? ""];
//   if (terminalEntry) {
//     const leaveType = (student as any).academicLeavePeriod?.type || "";
//     const rem       = ((student as any).remarks || "").toLowerCase();
//     let grounds = "";
//     if (leaveType === "financial"     || rem.includes("financial"))                              grounds = "FINANCIAL";
//     else if (leaveType === "compassionate" || rem.includes("compassionate") || rem.includes("medical")) grounds = "COMPASSIONATE";
//     else if (leaveType)                                                                           grounds = leaveType.toUpperCase();

//     const curriculumCount = await ProgramUnit.countDocuments({ program: programId, requiredYear: yearOfStudy });

//     return {
//       status:       terminalEntry.label,
//       variant:      terminalEntry.variant,
//       details:      `Student is currently ${terminalEntry.label}.${grounds ? ` Grounds: ${grounds}.` : ""}`,
//       weightedMean: "0.00",
//       sessionState: "ORDINARY",
//       summary: { totalExpected: curriculumCount, passed: 0, failed: 0, missing: 0, isOnLeave: true },
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
//       leaveDetails: grounds,
//     };
//   }

//   // ── Academic year resolution ───────────────────────────────────────────────
//   let targetYearDoc: any = null;

//   if (!academicYearName || academicYearName === "CURRENT" || academicYearName === "undefined") {
//     targetYearDoc =
//       (await AcademicYear.findOne({ isCurrent: true }).lean()) ||
//       (await AcademicYear.findOne().sort({ startDate: -1 }).lean());
//   } else {
//     targetYearDoc =
//       (await AcademicYear.findOne({ year: academicYearName }).lean()) ||
//       (await AcademicYear.findOne({ year: { $regex: new RegExp(`^${academicYearName.replace("/", "\\/")}$`, "i") } }).lean());
//     if (!targetYearDoc)
//       console.warn(`[StatusEngine] AcademicYear "${academicYearName}" not found.`);
//   }

//   // ── Curriculum ────────────────────────────────────────────────────────────
//   const curriculum = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy })
//     .populate("unit").lean() as any[];

//   if (!curriculum?.length) {
//     return {
//       status: "CURRICULUM NOT SET", variant: "info",
//       details: `No units defined for Year ${yearOfStudy}. Contact admin.`,
//       weightedMean: "0.00", sessionState: "ORDINARY",
//       summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
//     };
//   }

//   const programUnitIds = curriculum.map((pu) => pu._id);

//   const student = await Student.findById(studentId).lean();
//   const cfUnitIds = new Set(
//     (student?.carryForwardUnits || []).map((u: any) => u.programUnitId),
//   );
//   const [detailedMarks, directMarks, finalGrades] = await Promise.all([
//     Mark.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//     MarkDirect.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//     FinalGrade.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//   ]);

//   const marksMap = new Map<string, any>();

//   // Priority 3 (lowest): FinalGrade
//   finalGrades.forEach((fg: any) => {
//     const key = fg.programUnit?.toString();
//     if (!key) return;
//     marksMap.set(key, {
//       agreedMark:      fg.totalMark ?? 0,
//       caTotal30:       fg.caTotal30   != null ? fg.caTotal30   : (fg.totalMark > 0 ? 1 : 0),
//       examTotal70:     fg.examTotal70 != null ? fg.examTotal70 : (fg.totalMark > 0 ? 1 : 0),
//       attempt:         fg.attemptType === "SUPPLEMENTARY" ? "supplementary" : fg.attemptType === "RETAKE" ? "re-take" : "1st",
//       isSpecial:       fg.isSpecial === true || fg.status === "SPECIAL",
//       isSupplementary: fg.status === "SUPPLEMENTARY",
//       source:          "finalGrade",
//     });
//   });

//   // Priority 2: MarkDirect
//   directMarks.forEach((m: any) => marksMap.set(m.programUnit.toString(), { ...m, source: "direct" }));

//   // Priority 1 (highest): Detailed Mark
//   detailedMarks.forEach((m: any) => marksMap.set(m.programUnit.toString(), { ...m, source: "detailed" }));

//   const lists = {
//     passed:     [] as { code: string; mark: number }[],
//     failed:     [] as { displayName: string; attempt: string | number }[],
//     special:    [] as { displayName: string; grounds: string }[],
//     missing:    [] as string[],
//     incomplete: [] as string[],
//   };

//   let totalFirstAttemptSum     = 0;
//   let unitsContributingToMean  = 0;

//   curriculum.forEach((pUnit) => {
//     const code         = pUnit.unit?.code?.toUpperCase();
//     const displayName  = `${code}: ${pUnit.unit?.name}`;
//     const rawMarkRecord = marksMap.get(pUnit._id.toString());

//     if (!rawMarkRecord) {
//       lists.missing.push(displayName);
//       return;
//     }

//     const hasCAT         = (rawMarkRecord.caTotal30  || 0) > 0;
//     const hasExam        = (rawMarkRecord.examTotal70 || 0) > 0;
//     const markValue      = rawMarkRecord.agreedMark  || 0;
//     const isSupplementary = rawMarkRecord.attempt === "supplementary";
//     const isSpecial      = rawMarkRecord.attempt === "special" || rawMarkRecord.isSpecial;

//     const notation = getAttemptLabel({
//       markAttempt:      rawMarkRecord.attempt,
//       studentStatus:    (student as any).status,
//       studentQualifier: (student as any).qualifierSuffix,
//     });

//     if (isSpecial) {
//       lists.special.push({ displayName, grounds: rawMarkRecord.remarks || "Special" });
//     } else if (!hasCAT && !hasExam) {
//       lists.missing.push(`${displayName} (Absent)`);
//     } else if (!hasCAT && hasExam) {
//       if (isSupplementary) {
//         if (markValue >= passMark) lists.passed.push({ code, mark: markValue });
//         else                       lists.failed.push({ displayName, attempt: notation });
//         totalFirstAttemptSum += markValue;
//         unitsContributingToMean++;
//       } else {
//         lists.incomplete.push(`${displayName} (No CAT)`);
//       }
//     } else if (!hasExam && hasCAT) {
//       lists.missing.push(`${displayName} (Missing Exam)`);
//     } else {
//       if (markValue >= passMark) lists.passed.push({ code, mark: markValue });
//       else                       lists.failed.push({ displayName, attempt: notation });
//       totalFirstAttemptSum += markValue;
//       unitsContributingToMean++;
//     }
//   });

//   const totalUnits    = curriculum.length;
//   const failCount     = lists.failed.length;
//   const missingCount  = lists.missing.length;
//   const specialCount  = lists.special.length;
//   const incCount      = lists.incomplete.length;
//   const officialMean  = totalFirstAttemptSum / totalUnits;

//   const attemptedCount    = totalUnits - (specialCount + missingCount + incCount);
//   const performanceMean   = attemptedCount > 0 ? totalFirstAttemptSum / attemptedCount : 0;

//   const currentYearDoc = targetYearDoc?.isCurrent
//     ? targetYearDoc
//     : (await AcademicYear.findOne({ isCurrent: true }).lean()) ||
//       (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

//   const targetSession = targetYearDoc?.session ?? "ORDINARY";

//   const [targetStart] = (academicYearName || "0/0").split("/").map(Number);
//   const [globalStart] = currentYearDoc?.year ? currentYearDoc.year.split("/").map(Number) : [0];
//   const isPastYear    = targetYearDoc && globalStart > 0 ? targetStart < globalStart : false;
//   const isSessionClosed = targetSession === "CLOSED" || isPastYear;

//   // ── Status decision tree ───────────────────────────────────────────────────

//   let status  = "PASS";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = "Proceed to next year.";

//   if (!options.forPromotion && targetSession === "ORDINARY" && !isPastYear) {
//     status  = "SESSION IN PROGRESS";
//     variant = "info";
//     details = "Marks are currently being entered for this session.";
//   } else if (missingCount >= 6 && isSessionClosed) {
//     // ENG.23c — ONLY when CLOSED or a past year
//     status  = "DEREGISTERED";
//     variant = "error";
//     details = `Absent from 6+ (${missingCount}) examinations (ENG 23c).`;
//   } else if (specialCount > 0 && failCount < totalUnits / 2) {
//     const parts: string[] = [];
//     if (failCount > 0)   parts.push(`SUPP ${failCount}`);
//     parts.push(`SPEC ${specialCount}`);
//     if (incCount > 0)    parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`MISSING ${missingCount}`);
//     status  = parts.join("; ");
//     variant = "info";
//     details = `Awaiting specials. Mean in sat units: ${performanceMean.toFixed(2)}`;
//   } else if (failCount >= totalUnits / 2 || officialMean < 40) {
//     status  = "REPEAT YEAR";
//     variant = "error";
//     details = `Failed >= 50% (${failCount}/${totalUnits}) units or Mean (${officialMean.toFixed(2)}) < 40% (ENG 16).`;
//   } else if (failCount > totalUnits / 3) {
//     status  = "STAYOUT";
//     variant = "warning";
//     details = `Failed > 1/3 of units (${failCount}/${totalUnits}). Retake failed units next ordinary period (ENG 15h).`;
//   } else if (failCount > 0 || incCount > 0 || missingCount > 0) {
//     const parts: string[] = [];
//     if (failCount > 0)   parts.push(`SUPP ${failCount}`);
//     if (incCount > 0)    parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`INC ${missingCount}`);
//     status  = parts.join("; ");
//     variant = "warning";
//     details = "Eligible for supplementary exams or pending incomplete marks.";
//   }

//   return {
//     status,
//     variant,
//     details,
//     weightedMean: officialMean.toFixed(2),
//     sessionState: targetSession,
//     summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: lists.missing.length },
//     passedList:    lists.passed,
//     failedList:    lists.failed,
//     specialList:   lists.special,
//     missingList:   lists.missing,
//     incompleteList: lists.incomplete,
//   };
// };

// // ─── promoteStudent ───────────────────────────────────────────────────────────

// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId).populate("program");
//   if (!student) throw new Error("Student not found");

//   if (
//     ["deregistered", "discontinued", "graduated", "graduand"].includes(
//       student.status,
//     )
//   ) {
//     return {
//       success: false,
//       message: `Action blocked: Student is ${student.status}`,
//     };
//   }

//   const auditResult = await performAcademicAudit(studentId);
//   if (auditResult.discontinued)
//     return { success: false, message: `Discontinued: ${auditResult.reason}` };

//   const program = student.program as any;
//   const duration = program.durationYears || 4;
//   const actualCurrentYear = student.currentYearOfStudy || 1;
//   const currentSession = await AcademicYear.findOne({ isCurrent: true }).lean();
//   const completedYearLabel = currentSession?.year || "N/A";

//   const statusResult = await calculateStudentStatus(
//     student._id,
//     student.program,
//     completedYearLabel,
//     actualCurrentYear,
//     { forPromotion: true },
//   );

//   // ── TERMINAL STATUSES ──────────────────────────────────────────────────
//   const terminalStatuses = ["DEREGISTERED", "DISCONTINUED"];
//   if (terminalStatuses.includes(statusResult.status)) {
//     await syncTerminalStatusToDb(
//       studentId,
//       statusResult.status,
//       statusResult.details,
//       completedYearLabel,
//     );
//     return {
//       success: false,
//       message: `Promotion Blocked: ${statusResult.status}`,
//       details: statusResult,
//     };
//   }

//   // ── REPEAT YEAR (ENG.16a/c) ────────────────────────────────────────────
//   if (statusResult.status === "REPEAT YEAR") {
//     const repeatCount =
//       (student.academicHistory || []).filter(
//         (h) => h.isRepeatYear && h.yearOfStudy === actualCurrentYear,
//       ).length + 1;
//     const qualifier =
//       require("../utils/academicRules").REG_QUALIFIERS.repeatYear(repeatCount);

//     await Student.findByIdAndUpdate(studentId, {
//       $set: {
//         status: "repeat",
//         qualifierSuffix: qualifier,
//         remarks: statusResult.details,
//       },
//       $push: {
//         statusEvents: {
//           fromStatus: student.status,
//           toStatus: "repeat",
//           date: new Date(),
//           academicYear: completedYearLabel,
//           reason: `ENG.16: ${statusResult.details}`,
//         },
//       },
//     });
//     return {
//       success: false,
//       message: `Repeat year required`,
//       details: statusResult,
//     };
//   }

//   // ── STAYOUT (ENG.15h) ─────────────────────────────────────────────────
//   // Stayout students stay in same year but are "active" — they retake
//   // failed units in the next ordinary period.
//   if (statusResult.status === "STAYOUT") {
//     // Do NOT promote — they stay in current year
//     await Student.findByIdAndUpdate(studentId, {
//       $set: { remarks: `ENG.15h: ${statusResult.details}` },
//     });
//     return {
//       success: false,
//       message: `Stay out required (ENG.15h)`,
//       details: statusResult,
//     };
//   }

//   // ── SPECIALS PENDING ───────────────────────────────────────────────────
//   if (
//     statusResult.specialList.length > 0 &&
//     statusResult.failedList.length === 0
//   ) {
//     return {
//       success: false,
//       message: `Special examinations pending`,
//       details: statusResult,
//     };
//   }

//   // ── PASS (promote normally) OR SUPP PASS (check carry-forward) ────────
//   if (statusResult.status === "PASS") {
//     const rawMean = parseFloat(statusResult.weightedMean);
//     const yearWeight = getYearWeight(
//       program,
//       student.entryType || "Direct",
//       actualCurrentYear,
//     );
//     const weightedContribution = rawMean * yearWeight;

//     // Check if there are any pending carry-forward units from a previous year
//     const pendingCF = (student.carryForwardUnits || []).length;
//     if (actualCurrentYear === duration && pendingCF > 0) {
//       return {
//         success: false,
//         message: `Cannot graduate: ${pendingCF} carry-forward unit(s) not yet cleared`,
//       };
//     }

//     const historyRecord = {
//       academicYear: completedYearLabel,
//       yearOfStudy: actualCurrentYear,
//       annualMeanMark: rawMean,
//       weightedContribution,
//       unitsTakenCount: statusResult.summary.totalExpected,
//       failedUnitsCount: statusResult.summary.failed,
//       isRepeatYear: false,
//       date: new Date(),
//     };

//     if (actualCurrentYear === duration) {
//       // GRADUATION PATH
//       const fullHistory = [...(student.academicHistory || []), historyRecord];
//       const finalWAA = fullHistory.reduce(
//         (acc, curr) => acc + (curr.weightedContribution || 0),
//         0,
//       );
//       let classification = "PASS";
//       if (finalWAA >= 70) classification = "FIRST CLASS HONOURS";
//       else if (finalWAA >= 60)
//         classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//       else if (finalWAA >= 50)
//         classification = "SECOND CLASS HONOURS (LOWER DIVISION)";

//       await Student.findByIdAndUpdate(studentId, {
//         $set: {
//           status: "graduand",
//           qualifierSuffix: "",
//           finalWeightedAverage: finalWAA.toFixed(2),
//           classification,
//           graduationYear: new Date().getFullYear(),
//           currentYearOfStudy: actualCurrentYear + 1,
//         },
//         $push: { academicHistory: historyRecord },
//       });
//       return {
//         success: true,
//         message: `Graduated: ${classification}`,
//         isGraduation: true,
//       };
//     }

//     // NORMAL PROMOTION — check carry forward
//     const {
//       assessAndGrantCarryForward,
//     } = require("../services/carryForwardService");
//     const cfResult = await assessAndGrantCarryForward(
//       studentId,
//       student.program.toString(),
//       actualCurrentYear,
//       completedYearLabel,
//     );

//     const nextYear = actualCurrentYear + 1;
//     await Student.findByIdAndUpdate(studentId, {
//       $set: {
//         currentYearOfStudy: nextYear,
//         currentSemester: 1,
//         // qualifierSuffix is set by assessAndGrantCarryForward if CF granted,
//         // otherwise cleared if student is now clean
//         ...(cfResult.granted ? {} : { qualifierSuffix: "" }),
//       },
//       $push: {
//         promotionHistory: {
//           from: actualCurrentYear,
//           to: nextYear,
//           date: new Date(),
//         },
//         academicHistory: historyRecord,
//       },
//     });

//     return {
//       success: true,
//       message: cfResult.granted
//         ? `Promoted to Year ${nextYear} with carry-forward: ${cfResult.cfUnits.map((u) => u.unitCode).join(", ")}`
//         : `Successfully promoted to Year ${nextYear}`,
//       carryForward: cfResult,
//     };
//   }

//   // ── SUPP STATUS — assess carry-forward possibility ────────────────────
//   // Student failed some units but may still proceed (ENG.14)
//   const {
//     assessAndGrantCarryForward,
//   } = require("../services/carryForwardService");
//   const cfResult = await assessAndGrantCarryForward(
//     studentId,
//     student.program.toString(),
//     actualCurrentYear,
//     completedYearLabel,
//   );

//   if (cfResult.granted) {
//     const nextYear = actualCurrentYear + 1;
//     const rawMean = parseFloat(statusResult.weightedMean);
//     const yearWeight = getYearWeight(
//       program,
//       student.entryType || "Direct",
//       actualCurrentYear,
//     );

//     await Student.findByIdAndUpdate(studentId, {
//       $set: {
//         currentYearOfStudy: nextYear,
//         currentSemester: 1,
//       },
//       $push: {
//         academicHistory: {
//           academicYear: completedYearLabel,
//           yearOfStudy: actualCurrentYear,
//           annualMeanMark: rawMean,
//           weightedContribution: rawMean * yearWeight,
//           failedUnitsCount: cfResult.cfUnits.length,
//           isRepeatYear: false,
//           date: new Date(),
//         },
//       },
//     });

//     return {
//       success: true,
//       message: `Promoted to Year ${nextYear} with ${cfResult.cfUnits.length} carry-forward unit(s)`,
//       carryForward: cfResult,
//     };
//   }

//   // No carry-forward possible — block promotion
//   let blockMessage = "Promotion Blocked: ";
//   if (statusResult.status.includes("SUPP"))
//     blockMessage += "Student has unresolved supplementary examinations.";
//   else blockMessage += `Current status is '${statusResult.status}'.`;

//   return { success: false, message: blockMessage, details: statusResult };
// };

// // ─── previewPromotion ─────────────────────────────────────────────────────────

// export const previewPromotion = async (
//   programId:        string,
//   yearToPromote:    number,
//   academicYearName: string,
// ) => {
//   const nextYear = yearToPromote + 1;

//   const targetYearDoc = await AcademicYear.findOne({ year: academicYearName }).lean();
//   if (!targetYearDoc) {
//     console.warn(`[previewPromotion] AcademicYear "${academicYearName}" not found.`);
//     return { totalProcessed: 0, eligibleCount: 0, blockedCount: 0, eligible: [], blocked: [] };
//   }

//   const admissionStudents = await Student.find({
//     program: programId, currentYearOfStudy: yearToPromote,
//     admissionAcademicYear: (targetYearDoc as any)._id,
//   }).lean();

//   const [marksThisYear, directMarksThisYear] = await Promise.all([
//     Mark.distinct("student",       { academicYear: (targetYearDoc as any)._id }),
//     MarkDirect.distinct("student", { academicYear: (targetYearDoc as any)._id }),
//   ]);

//   const markedIds   = new Set<string>([...marksThisYear, ...directMarksThisYear].map((id: any) => id.toString()));
//   const admissionIds = new Set(admissionStudents.map((s) => (s as any)._id.toString()));

//   const returningStudents = await Student.find({
//     program: programId, currentYearOfStudy: yearToPromote,
//     _id: { $in: Array.from(markedIds), $nin: Array.from(admissionIds) },
//   }).lean();

//   const adminStudents = await Student.find({
//     program: programId, currentYearOfStudy: yearToPromote,
//     status: { $in: ["on_leave", "deferred", "deregistered", "discontinued"] },
//     $or: [
//       { admissionAcademicYear: (targetYearDoc as any)._id },
//       { "academicHistory.academicYear": academicYearName },
//     ],
//     _id: { $nin: [...Array.from(admissionIds), ...returningStudents.map((s) => (s as any)._id.toString())] },
//   }).lean();

//   const allStudents = [...admissionStudents, ...returningStudents, ...adminStudents];

//   const ADMIN_STATUSES: Record<string, string> = {
//     on_leave:     "ACADEMIC LEAVE",
//     deferred:     "DEFERMENT",
//     discontinued: "DISCONTINUED",
//     deregistered: "DEREGISTERED",
//     graduated:    "GRADUATED",
//   };

//   const eligible: any[] = [];
//   const blocked:  any[] = [];

//   for (const student of allStudents) {
//     const isAlreadyPromoted = (student as any).currentYearOfStudy === nextYear;
//     const adminLabel        = ADMIN_STATUSES[(student as any).status];

//     if (isAlreadyPromoted) {
//       eligible.push({
//         id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
//         status: "ALREADY PROMOTED", reasons: [], specialGrounds: "",
//         summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//         remarks: (student as any).remarks, academicLeavePeriod: (student as any).academicLeavePeriod, details: "",
//       });
//       continue;
//     }

//     if (adminLabel) {
//       const leaveType  = (student as any).academicLeavePeriod?.type?.toUpperCase();
//       const reason     = leaveType ? `${adminLabel} (${leaveType})` : adminLabel;
//       const adminGrounds = [
//         ((student as any).academicLeavePeriod?.type || "").toLowerCase(),
//         ((student as any).remarks || "").toLowerCase(),
//       ].join(" ").trim() || "other";

//       blocked.push({
//         id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
//         status: adminLabel, reasons: [reason], specialGrounds: adminGrounds,
//         summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//         academicLeavePeriod: (student as any).academicLeavePeriod,
//         remarks: (student as any).remarks, details: "",
//       });
//       continue;
//     }

//     const statusResult = await calculateStudentStatus(
//       (student as any)._id, programId, academicYearName, yearToPromote,
//       { forPromotion: true },
//     );

//     const specialGrounds = (() => {
//       const fromList    = (statusResult.specialList || []).map((s) => (s.grounds || "").toLowerCase()).join(" ");
//       const fromRemarks = ((student as any).remarks || "").toLowerCase();
//       const fromLeave   = ((student as any).academicLeavePeriod?.type || "").toLowerCase();
//       return `${fromList} ${fromRemarks} ${fromLeave}`.trim() || "other";
//     })();

//     const report: any = {
//       id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name,
//       status: statusResult.status, summary: statusResult.summary, reasons: [],
//       remarks: (student as any).remarks, academicLeavePeriod: (student as any).academicLeavePeriod,
//       details: statusResult.details, specialGrounds,
//       isEligibleForSupp: !["STAYOUT","REPEAT YEAR","DEREGISTERED"].includes(statusResult.status) &&
//         (statusResult.failedList.length > 0 || statusResult.specialList.length > 0),
//     };

//     if (statusResult.status === "PASS") {
//       eligible.push(report);
//     } else {
//       if (statusResult.status === "STAYOUT")      report.reasons.push("ENG 15h: > 1/3 units failed");
//       if (statusResult.status === "REPEAT YEAR")  report.reasons.push("ENG 16: >= 1/2 units failed or mean < 40%");
//       if (statusResult.status === "DEREGISTERED") report.reasons.push("ENG 23c: Absent from 6+ examinations");
//       statusResult.specialList.forEach((s) => report.reasons.push(`${s.displayName} (SPECIAL)`));
//       statusResult.incompleteList.forEach((u)   => report.reasons.push(`${u} (INCOMPLETE)`));
//       statusResult.missingList.forEach((u)       => report.reasons.push(`${u} (MISSING)`));
//       statusResult.failedList.forEach((f: any)   => report.reasons.push(`${f.displayName} (FAIL ATTEMPT: ${f.attempt})`));
//       blocked.push(report);
//     }
//   }

//   return {
//     totalProcessed: allStudents.length,
//     eligibleCount:  eligible.length,
//     blockedCount:   blocked.length,
//     eligible,
//     blocked,
//   };
// };

// // ─── bulkPromoteClass ─────────────────────────────────────────────────────────

// export const bulkPromoteClass = async (
//   programId:        string,
//   yearToPromote:    number,
//   academicYearName: string,
// ) => {
//   const nextYear  = yearToPromote + 1;
//   const students  = await Student.find({
//     program: programId,
//     currentYearOfStudy: { $in: [yearToPromote, nextYear] },
//     status: "active",
//   });

//   const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };

//   for (const student of students) {
//     try {
//       const studentId = (student._id as any).toString();
//       if ((student as any).currentYearOfStudy >= nextYear) {
//         results.alreadyPromoted++;
//         results.promoted++;
//         continue;
//       }
//       const res = await promoteStudent(studentId);
//       if (res.success) results.promoted++;
//       else             results.failed++;
//     } catch (err: any) {
//       results.errors.push(`${(student as any).regNo}: ${err.message}`);
//     }
//   }

//   return results;
// };





















// // serverside/src/services/statusEngine.ts
// // FIXES: (1) duplicate const student removed from calculateStudentStatus
// //         (2) all lambda params explicitly typed — no TS7006
// //         (3) no require() calls — all imports at top level
// import mongoose from "mongoose";
// import FinalGrade from "../models/FinalGrade";
// import ProgramUnit from "../models/ProgramUnit";
// import Student from "../models/Student";
// import Mark from "../models/Mark";
// import InstitutionSettings from "../models/InstitutionSettings";
// import MarkDirect from "../models/MarkDirect";
// import AcademicYear from "../models/AcademicYear";
// import { getYearWeight } from "../utils/weightingRegistry";
// import { performAcademicAudit } from "./academicAudit";
// import { getAttemptLabel, REG_QUALIFIERS } from "../utils/academicRules";
// import { assessAndGrantCarryForward } from "./carryForwardService";

// const syncTerminalStatusToDb = async (
//   studentId: string, engineStatus: string, details: string, academicYear: string,
// ): Promise<void> => {
//   interface Mapping { dbStatus: string; qualifierFn: (s: any) => string }
//   const terminalMap: Record<string, Mapping> = {
//     "DEREGISTERED": { dbStatus: "deregistered", qualifierFn: () => "" },
//     "REPEAT YEAR":  { dbStatus: "repeat", qualifierFn: (s: any) => {
//         const count = ((s.academicHistory || []) as any[]).filter((h: any) => h.isRepeatYear).length + 1;
//         return REG_QUALIFIERS.repeatYear(count);
//     }},
//     "STAYOUT":      { dbStatus: "active",       qualifierFn: () => "" },
//     "DISCONTINUED": { dbStatus: "discontinued", qualifierFn: () => "" },
//   };

//   const entry = terminalMap[engineStatus];
//   if (!entry) return;

//   const student = await Student.findById(studentId).lean();
//   if (!student) return;
//   if (entry.dbStatus !== "active" && (student as any).status === entry.dbStatus) return;

//   const fromStatus   = (student as any).status;
//   const newQualifier = entry.qualifierFn(student);

//   const updatePayload: any = {
//     $set:  { status: entry.dbStatus, remarks: details },
//     $push: {
//       statusEvents:  { fromStatus, toStatus: entry.dbStatus, date: new Date(), reason: `Auto-Sync: ${details}`, academicYear },
//       statusHistory: { status: entry.dbStatus, previousStatus: fromStatus, date: new Date(), reason: details },
//     },
//   };
//   if (newQualifier) updatePayload.$set.qualifierSuffix = newQualifier;
//   await Student.findByIdAndUpdate(studentId, updatePayload);
// };

// export interface StudentStatusResult {
//   status: string; variant: "success" | "warning" | "error" | "info"; details: string;
//   weightedMean: string; sessionState: string;
//   summary: { totalExpected: number; passed: number; failed: number; missing: number; isOnLeave?: boolean };
//   passedList: { code: string; mark: number }[];
//   failedList: { displayName: string; attempt: string | number }[];
//   specialList: { displayName: string; grounds: string }[];
//   missingList: string[]; incompleteList: string[]; leaveDetails?: string;
// }

// export const calculateStudentStatus = async (
//   studentId: any, programId: any, academicYearName: string,
//   yearOfStudy: number = 1, options: { forPromotion?: boolean } = {},
// ): Promise<StudentStatusResult> => {
//   const settings = await InstitutionSettings.findOne().lean();
//   if (!settings) throw new Error("Institution settings not found.");
//   const passMark = (settings as any).passMark || 40;

//   // Declared exactly ONCE — fixes TS2451 "Cannot redeclare block-scoped variable 'student'"
//   const student = await Student.findById(studentId).lean();
//   if (!student) throw new Error("Student not found");

//   const TERMINAL: Record<string, { label: string; variant: "info" | "error" | "success" | "warning" }> = {
//     on_leave:     { label: "ACADEMIC LEAVE", variant: "info"    },
//     deferred:     { label: "DEFERMENT",      variant: "info"    },
//     discontinued: { label: "DISCONTINUED",   variant: "error"   },
//     deregistered: { label: "DEREGISTERED",   variant: "error"   },
//     graduated:    { label: "GRADUATED",      variant: "success" },
//     graduand:     { label: "GRADUATED",      variant: "success" },
//   };

//   const terminalEntry = TERMINAL[(student as any).status ?? ""];
//   if (terminalEntry) {
//     const leaveType = (student as any).academicLeavePeriod?.type || "";
//     const rem       = ((student as any).remarks || "").toLowerCase();
//     let grounds = "";
//     if (leaveType === "financial" || rem.includes("financial"))                                grounds = "FINANCIAL";
//     else if (leaveType === "compassionate" || rem.includes("compassionate") || rem.includes("medical")) grounds = "COMPASSIONATE";
//     else if (leaveType) grounds = leaveType.toUpperCase();
//     const count = await ProgramUnit.countDocuments({ program: programId, requiredYear: yearOfStudy });
//     return {
//       status: terminalEntry.label, variant: terminalEntry.variant,
//       details: `Student is currently ${terminalEntry.label}.${grounds ? ` Grounds: ${grounds}.` : ""}`,
//       weightedMean: "0.00", sessionState: "ORDINARY",
//       summary: { totalExpected: count, passed: 0, failed: 0, missing: 0, isOnLeave: true },
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [], leaveDetails: grounds,
//     };
//   }

//   let targetYearDoc: any = null;
//   if (!academicYearName || academicYearName === "CURRENT" || academicYearName === "undefined") {
//     targetYearDoc = (await AcademicYear.findOne({ isCurrent: true }).lean()) || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());
//   } else {
//     targetYearDoc = (await AcademicYear.findOne({ year: academicYearName }).lean())
//       || (await AcademicYear.findOne({ year: { $regex: new RegExp(`^${academicYearName.replace("/", "\\/")}$`, "i") } }).lean());
//   }

//   const curriculum = await ProgramUnit.find({ program: programId, requiredYear: yearOfStudy }).populate("unit").lean() as any[];
//   if (!curriculum?.length) {
//     return {
//       status: "CURRICULUM NOT SET", variant: "info", details: `No units defined for Year ${yearOfStudy}.`,
//       weightedMean: "0.00", sessionState: "ORDINARY",
//       summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 },
//       passedList: [], failedList: [], specialList: [], missingList: [], incompleteList: [],
//     };
//   }

//   const programUnitIds = curriculum.map((pu: any) => pu._id);

//   // Read carryForwardUnits from the already-fetched student (no second DB call)
//   const cfUnitIds = new Set<string>(
//     ((student as any).carryForwardUnits || []).map((u: any): string => u.programUnitId as string),
//   );

//   const [detailedMarks, directMarks, finalGrades] = await Promise.all([
//     Mark.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//     MarkDirect.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//     FinalGrade.find({ student: studentId, programUnit: { $in: programUnitIds } }).lean(),
//   ]);

//   const marksMap = new Map<string, any>();
//   finalGrades.forEach((fg: any) => {
//     const key = fg.programUnit?.toString(); if (!key) return;
//     marksMap.set(key, {
//       agreedMark: fg.totalMark ?? 0,
//       caTotal30:  fg.caTotal30  != null ? fg.caTotal30  : (fg.totalMark > 0 ? 1 : 0),
//       examTotal70: fg.examTotal70 != null ? fg.examTotal70 : (fg.totalMark > 0 ? 1 : 0),
//       attempt:    fg.attemptType === "SUPPLEMENTARY" ? "supplementary" : fg.attemptType === "RETAKE" ? "re-take" : "1st",
//       isSpecial:  fg.isSpecial === true || fg.status === "SPECIAL",
//       source:     "finalGrade",
//     });
//   });
//   directMarks.forEach((m: any)   => marksMap.set(m.programUnit.toString(), { ...m, source: "direct" }));
//   detailedMarks.forEach((m: any) => marksMap.set(m.programUnit.toString(), { ...m, source: "detailed" }));

//   const lists = {
//     passed:     [] as { code: string; mark: number }[],
//     failed:     [] as { displayName: string; attempt: string | number }[],
//     special:    [] as { displayName: string; grounds: string }[],
//     missing:    [] as string[],
//     incomplete: [] as string[],
//   };
//   let totalFirstAttemptSum = 0, unitsContributingToMean = 0;

//   curriculum.forEach((pUnit: any) => {
//     const code        = pUnit.unit?.code?.toUpperCase();
//     const displayName = `${code}: ${pUnit.unit?.name}`;
//     const rawMark     = marksMap.get(pUnit._id.toString());

//     if (!rawMark) { lists.missing.push(displayName); return; }

//     const hasCAT  = (rawMark.caTotal30  || 0) > 0;
//     const hasExam = (rawMark.examTotal70 || 0) > 0;
//     const markVal = rawMark.agreedMark || 0;
//     const isSupp  = rawMark.attempt === "supplementary";
//     const isSpc   = rawMark.attempt === "special" || rawMark.isSpecial;

//     const notation = getAttemptLabel({ markAttempt: rawMark.attempt, studentStatus: (student as any).status, studentQualifier: (student as any).qualifierSuffix });

//     if (isSpc) {
//       lists.special.push({ displayName, grounds: rawMark.remarks || "Special" });
//     } else if (!hasCAT && !hasExam) {
//       lists.missing.push(`${displayName} (Absent)`);
//     } else if (!hasCAT && hasExam) {
//       if (isSupp) {
//         if (markVal >= passMark) lists.passed.push({ code, mark: markVal });
//         else                     lists.failed.push({ displayName, attempt: notation });
//         totalFirstAttemptSum += markVal; unitsContributingToMean++;
//       } else { lists.incomplete.push(`${displayName} (No CAT)`); }
//     } else if (!hasExam && hasCAT) {
//       lists.missing.push(`${displayName} (Missing Exam)`);
//     } else {
//       if (markVal >= passMark) lists.passed.push({ code, mark: markVal });
//       else                     lists.failed.push({ displayName, attempt: notation });
//       totalFirstAttemptSum += markVal; unitsContributingToMean++;
//     }
//   });

//   const totalUnits    = curriculum.length;
//   const failCount     = lists.failed.length;
//   const missingCount  = lists.missing.length;
//   const specialCount  = lists.special.length;
//   const incCount      = lists.incomplete.length;
//   const officialMean  = totalUnits > 0 ? totalFirstAttemptSum / totalUnits : 0;
//   const attemptedN    = totalUnits - (specialCount + missingCount + incCount);
//   const perfMean      = attemptedN > 0 ? totalFirstAttemptSum / attemptedN : 0;

//   const currentYearDoc = targetYearDoc?.isCurrent
//     ? targetYearDoc
//     : (await AcademicYear.findOne({ isCurrent: true }).lean()) || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

//   const targetSession   = targetYearDoc?.session ?? "ORDINARY";
//   const [tStart]        = (academicYearName || "0/0").split("/").map(Number);
//   const [gStart]        = currentYearDoc?.year ? currentYearDoc.year.split("/").map(Number) : [0];
//   const isPastYear      = targetYearDoc && gStart > 0 ? tStart < gStart : false;
//   const isSessionClosed = targetSession === "CLOSED" || isPastYear;

//   let status  = "PASS";
//   let variant: "success" | "warning" | "error" | "info" = "success";
//   let details = "Proceed to next year.";

//   if (!options.forPromotion && targetSession === "ORDINARY" && !isPastYear) {
//     status = "SESSION IN PROGRESS"; variant = "info"; details = "Marks are currently being entered.";
//   } else if (missingCount >= 6 && isSessionClosed) {
//     status = "DEREGISTERED"; variant = "error"; details = `Absent from 6+ (${missingCount}) examinations (ENG 23c).`;
//   } else if (specialCount > 0 && failCount < totalUnits / 2) {
//     const parts: string[] = [];
//     if (failCount > 0)    parts.push(`SUPP ${failCount}`);
//     parts.push(`SPEC ${specialCount}`);
//     if (incCount > 0)     parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`MISSING ${missingCount}`);
//     status = parts.join("; "); variant = "info"; details = `Awaiting specials. Mean: ${perfMean.toFixed(2)}`;
//   } else if (failCount >= totalUnits / 2 || officialMean < 40) {
//     status = "REPEAT YEAR"; variant = "error";
//     details = `Failed >= 50% (${failCount}/${totalUnits}) or Mean (${officialMean.toFixed(2)}) < 40% (ENG 16).`;
//   } else if (failCount > totalUnits / 3) {
//     status = "STAYOUT"; variant = "warning";
//     details = `Failed > 1/3 (${failCount}/${totalUnits}). Retake in next ordinary period (ENG 15h).`;
//   } else if (failCount > 0 || incCount > 0 || missingCount > 0) {
//     const parts: string[] = [];
//     if (failCount > 0)    parts.push(`SUPP ${failCount}`);
//     if (incCount > 0)     parts.push(`INC ${incCount}`);
//     if (missingCount > 0) parts.push(`INC ${missingCount}`);
//     status = parts.join("; "); variant = "warning"; details = "Eligible for supplementary exams.";
//   }

//   return {
//     status, variant, details, weightedMean: officialMean.toFixed(2), sessionState: targetSession,
//     summary: { totalExpected: totalUnits, passed: lists.passed.length, failed: failCount, missing: missingCount },
//     passedList: lists.passed, failedList: lists.failed, specialList: lists.special,
//     missingList: lists.missing, incompleteList: lists.incomplete,
//   };
// };

// export const promoteStudent = async (studentId: string) => {
//   const student = await Student.findById(studentId).populate("program");
//   if (!student) throw new Error("Student not found");

//   const st = (student as any).status as string;
//   if (["deregistered", "discontinued", "graduated", "graduand"].includes(st))
//     return { success: false, message: `Action blocked: Student is ${st}` };

//   const auditResult = await performAcademicAudit(studentId);
//   if (auditResult.discontinued) return { success: false, message: `Discontinued: ${auditResult.reason}` };

//   const program       = student.program as any;
//   const duration      = program.durationYears || 5;
//   const currentYear   = (student as any).currentYearOfStudy || 1;
//   const currentSession = await AcademicYear.findOne({ isCurrent: true }).lean();
//   const completedYear = currentSession?.year || "N/A";

//   const statusResult = await calculateStudentStatus(student._id, student.program, completedYear, currentYear, { forPromotion: true });

//   if (["DEREGISTERED", "DISCONTINUED"].includes(statusResult.status)) {
//     await syncTerminalStatusToDb(studentId, statusResult.status, statusResult.details, completedYear);
//     return { success: false, message: `Promotion Blocked: ${statusResult.status}`, details: statusResult };
//   }

//   if (statusResult.status === "REPEAT YEAR") {
//     const repeatCount = ((student as any).academicHistory || []).filter((h: any) => h.isRepeatYear && h.yearOfStudy === currentYear).length + 1;
//     const qualifier   = REG_QUALIFIERS.repeatYear(repeatCount);
//     await Student.findByIdAndUpdate(studentId, {
//       $set:  { status: "repeat", qualifierSuffix: qualifier, remarks: statusResult.details },
//       $push: {
//         statusEvents:  { fromStatus: st, toStatus: "repeat", date: new Date(), academicYear: completedYear, reason: `ENG.16: ${statusResult.details}` },
//         statusHistory: { status: "repeat", previousStatus: st, date: new Date(), reason: statusResult.details },
//         academicHistory: { academicYear: completedYear, yearOfStudy: currentYear, annualMeanMark: parseFloat(statusResult.weightedMean), weightedContribution: 0, failedUnitsCount: statusResult.summary.failed, isRepeatYear: true, date: new Date() },
//       },
//     });
//     return { success: false, message: `Repeat year required (${qualifier})`, details: statusResult };
//   }

//   if (statusResult.status === "STAYOUT") {
//     await Student.findByIdAndUpdate(studentId, {
//       $set:  { remarks: `ENG.15h: ${statusResult.details}` },
//       $push: { statusEvents: { fromStatus: st, toStatus: "active", date: new Date(), academicYear: completedYear, reason: `ENG.15h: ${statusResult.details}` } },
//     });
//     return { success: false, message: "Stay out required (ENG.15h)", details: statusResult };
//   }

//   if (statusResult.specialList.length > 0 && statusResult.failedList.length === 0)
//     return { success: false, message: "Special examinations pending", details: statusResult };

//   const rawMean    = parseFloat(statusResult.weightedMean);
//   const yearWeight = getYearWeight(program, (student as any).entryType || "Direct", currentYear);
//   const histRecord = {
//     academicYear: completedYear, yearOfStudy: currentYear, annualMeanMark: rawMean,
//     weightedContribution: rawMean * yearWeight, failedUnitsCount: statusResult.summary.failed,
//     isRepeatYear: false, date: new Date(),
//   };

//   const pendingCF = ((student as any).carryForwardUnits || []).filter((u: any) => u.status === "pending").length;

//   if (currentYear === duration) {
//     if (pendingCF > 0) return { success: false, message: `Cannot graduate: ${pendingCF} carry-forward unit(s) not yet cleared` };
//     const fullHistory = [...((student as any).academicHistory || []), histRecord];
//     const finalWAA    = fullHistory.reduce((acc: number, h: any) => acc + (h.weightedContribution || 0), 0);
//     let classification = "PASS";
//     if (finalWAA >= 70)      classification = "FIRST CLASS HONOURS";
//     else if (finalWAA >= 60) classification = "SECOND CLASS HONOURS (UPPER DIVISION)";
//     else if (finalWAA >= 50) classification = "SECOND CLASS HONOURS (LOWER DIVISION)";
//     await Student.findByIdAndUpdate(studentId, {
//       $set:  { status: "graduand", qualifierSuffix: "", finalWeightedAverage: finalWAA.toFixed(2), classification, graduationYear: new Date().getFullYear(), currentYearOfStudy: currentYear + 1 },
//       $push: { academicHistory: histRecord },
//     });
//     return { success: true, message: `Graduated: ${classification}`, isGraduation: true };
//   }

//   const cfResult  = await assessAndGrantCarryForward(studentId, student.program.toString(), currentYear, completedYear);
//   const nextYear  = currentYear + 1;

//   await Student.findByIdAndUpdate(studentId, {
//     $set:  { currentYearOfStudy: nextYear, currentSemester: 1, ...(cfResult.granted ? {} : { qualifierSuffix: "" }) },
//     $push: { promotionHistory: { from: currentYear, to: nextYear, date: new Date() }, academicHistory: histRecord, statusHistory: { status: "active", previousStatus: st, date: new Date(), reason: `Promoted Year ${nextYear}` } },
//   });

//   return {
//     success: true,
//     message: cfResult.granted
//       ? `Promoted to Year ${nextYear} with carry-forward (${cfResult.qualifier}): ${cfResult.cfUnits.map((u) => u.unitCode).join(", ")}`
//       : `Successfully promoted to Year ${nextYear}`,
//     carryForward: cfResult.granted ? cfResult : undefined,
//   };
// };

// export const previewPromotion = async (programId: string, yearToPromote: number, academicYearName: string) => {
//   const nextYear      = yearToPromote + 1;
//   const targetYearDoc = await AcademicYear.findOne({ year: academicYearName }).lean();
//   if (!targetYearDoc) return { totalProcessed: 0, eligibleCount: 0, blockedCount: 0, eligible: [], blocked: [] };

//   const admissionStudents = await Student.find({ program: programId, currentYearOfStudy: yearToPromote, admissionAcademicYear: (targetYearDoc as any)._id }).lean();
//   const [m1, m2] = await Promise.all([Mark.distinct("student", { academicYear: (targetYearDoc as any)._id }), MarkDirect.distinct("student", { academicYear: (targetYearDoc as any)._id })]);
//   const markedIds    = new Set<string>([...m1, ...m2].map((id: any) => id.toString()));
//   const admissionIds = new Set(admissionStudents.map((s: any) => s._id.toString()));

//   const returningStudents = await Student.find({ program: programId, currentYearOfStudy: yearToPromote, _id: { $in: Array.from(markedIds), $nin: Array.from(admissionIds) } }).lean();
//   const adminStudents     = await Student.find({ program: programId, currentYearOfStudy: yearToPromote, status: { $in: ["on_leave","deferred","deregistered","discontinued"] }, $or: [{ admissionAcademicYear: (targetYearDoc as any)._id }, { "academicHistory.academicYear": academicYearName }], _id: { $nin: [...Array.from(admissionIds), ...returningStudents.map((s: any) => s._id.toString())] } }).lean();

//   const allStudents = [...admissionStudents, ...returningStudents, ...adminStudents];
//   const ADMIN_LABELS: Record<string, string> = { on_leave: "ACADEMIC LEAVE", deferred: "DEFERMENT", discontinued: "DISCONTINUED", deregistered: "DEREGISTERED", graduated: "GRADUATED" };
//   const eligible: any[] = [];
//   const blocked:  any[] = [];

//   for (const student of allStudents) {
//     const isAlreadyPromoted = (student as any).currentYearOfStudy === nextYear;
//     const adminLabel        = ADMIN_LABELS[(student as any).status];

//     if (isAlreadyPromoted) { eligible.push({ id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name, status: "ALREADY PROMOTED", reasons: [], specialGrounds: "", summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 }, details: "" }); continue; }
//     if (adminLabel) {
//       const leaveType = (student as any).academicLeavePeriod?.type?.toUpperCase();
//       blocked.push({ id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name, status: adminLabel, reasons: [leaveType ? `${adminLabel} (${leaveType})` : adminLabel], specialGrounds: ((student as any).remarks || "other").toLowerCase(), summary: { totalExpected: 0, passed: 0, failed: 0, missing: 0 }, details: "" });
//       continue;
//     }

//     const sr = await calculateStudentStatus((student as any)._id, programId, academicYearName, yearToPromote, { forPromotion: true });
//     const specialGrounds = [(sr.specialList || []).map((s: any) => (s.grounds || "").toLowerCase()).join(" "), ((student as any).remarks || "").toLowerCase(), ((student as any).academicLeavePeriod?.type || "").toLowerCase()].join(" ").trim() || "other";

//     const report: any = { id: (student as any)._id, regNo: (student as any).regNo, name: (student as any).name, status: sr.status, summary: sr.summary, reasons: [], details: sr.details, specialGrounds, isEligibleForSupp: !["STAYOUT","REPEAT YEAR","DEREGISTERED"].includes(sr.status) && (sr.failedList.length > 0 || sr.specialList.length > 0) };

//     if (sr.status === "PASS") { eligible.push(report); continue; }
//     if (sr.status === "STAYOUT")      report.reasons.push("ENG 15h: > 1/3 units failed");
//     if (sr.status === "REPEAT YEAR")  report.reasons.push("ENG 16: >= 1/2 units failed or mean < 40%");
//     if (sr.status === "DEREGISTERED") report.reasons.push("ENG 23c: Absent from 6+ examinations");
//     sr.specialList.forEach((s: any)     => report.reasons.push(`${s.displayName} (SPECIAL)`));
//     sr.incompleteList.forEach((u: string) => report.reasons.push(`${u} (INCOMPLETE)`));
//     sr.missingList.forEach((u: string)    => report.reasons.push(`${u} (MISSING)`));
//     sr.failedList.forEach((f: any)        => report.reasons.push(`${f.displayName} (FAIL: ${f.attempt})`));
//     blocked.push(report);
//   }

//   return { totalProcessed: allStudents.length, eligibleCount: eligible.length, blockedCount: blocked.length, eligible, blocked };
// };

// export const bulkPromoteClass = async (programId: string, yearToPromote: number, academicYearName: string) => {
//   const nextYear = yearToPromote + 1;
//   const students = await Student.find({ program: programId, currentYearOfStudy: { $in: [yearToPromote, nextYear] }, status: "active" });
//   const results = { promoted: 0, failed: 0, alreadyPromoted: 0, errors: [] as string[] };
//   for (const student of students) {
//     try {
//       const sid = (student._id as any).toString();
//       if ((student as any).currentYearOfStudy >= nextYear) { results.alreadyPromoted++; results.promoted++; continue; }
//       const res = await promoteStudent(sid);
//       if (res.success) results.promoted++; else results.failed++;
//     } catch (err: any) { results.errors.push(`${(student as any).regNo}: ${err.message}`); }
//   }
//   return results;
// };
















// serverside/src/services/statusEngine.ts
import mongoose from "mongoose";
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
  console.log(`[StatusEngine] syncTerminalStatus: ${fromStatus} → ${entry.dbStatus} (${engineStatus}) for student ${studentId}`);
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
    : (await AcademicYear.findOne({ isCurrent: true }).lean())
      || (await AcademicYear.findOne().sort({ startDate: -1 }).lean());

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

  const statusResult = await calculateStudentStatus(
    student._id, student.program, completedYear, currentYear, { forPromotion: true },
  );

  // console.log(`[promoteStudent] Status result for ${regNo}: status="${statusResult.status}", mean=${statusResult.weightedMean}, failed=${statusResult.summary.failed}/${statusResult.summary.totalExpected}`);

  // ── DEREGISTERED / DISCONTINUED ───────────────────────────────────────────
  if (["DEREGISTERED", "DISCONTINUED"].includes(statusResult.status)) {
    // console.log(`[promoteStudent] Syncing terminal status: ${statusResult.status}`);
    await syncTerminalStatusToDb(studentId, statusResult.status, statusResult.details, completedYear);
    return { success: false, message: `Promotion Blocked: ${statusResult.status}`, details: statusResult };
  }

  // ── REPEAT YEAR ───────────────────────────────────────────────────────────
  if (statusResult.status === "REPEAT YEAR") {
    const repeatCount = ((student as any).academicHistory || [])
      .filter((h: any) => h.isRepeatYear && h.yearOfStudy === currentYear).length + 1;
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

  const pendingCF = ((student as any).carryForwardUnits || [])
    .filter((u: any) => u.status === "pending").length;

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
    const cfResult = await tryAssessAndGrantCarryForward(
      studentId, student.program.toString(), currentYear, completedYear,
    );
    cfGranted   = cfResult.granted;
    cfQualifier = cfResult.qualifier;
    cfMessage   = cfGranted
      ? `Promoted to Year ${nextYear} with carry-forward (${cfResult.qualifier}): ${cfResult.cfUnits.map((u) => u.unitCode).join(", ")}`
      : "";
    // console.log(`[promoteStudent] CF result for ${regNo}: granted=${cfGranted}, reason="${cfResult.reason}"`);

    // If carry-forward was NOT granted and status is not PASS, block the promotion
    if (!cfGranted) {
      // console.log(`[promoteStudent] Cannot promote ${regNo}: ${cfResult.reason}`);
      return {
        success: false,
        message: `Promotion Blocked: ${cfResult.reason}`,
        details: statusResult,
      };
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
        console.log(`[bulkPromoteClass] ✗ ${rNo}: ${res.message}`);
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