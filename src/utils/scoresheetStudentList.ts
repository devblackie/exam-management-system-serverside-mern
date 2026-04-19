
// serverside/src/utils/scoresheetStudentList.ts

import mongoose from "mongoose";
import type * as ExcelJS from "exceljs";
import Student from "../models/Student";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import { buildDisplayRegNo } from "./academicRules";
import { getCarryForwardStudentsForUnit, getStayoutStudentsForUnit } from "../services/carryForwardService";

export interface ScoresheetStudent {
  regNo: string;
  displayRegNo: string;
  name: string;
  studentId: string;
  attemptLabel: string;
  isSupp: boolean;
  isSpecial: boolean;
  isCarriedSpecial: boolean;
  isCarryForward: boolean;
  isStayout: boolean;
  isRepeatYear: boolean;
  prevMark?: any;
  qualifierSuffix?: string;
}

const ATTEMPT = {FIRST: "1st", SUPPLEMENTARY: "Supp", RETAKE: "Retake", SPECIAL: "Special", REPEAT_YEAR: "Repeat", STAYOUT: "Retake", CF: "Retake"};

const EXCLUDED = ["deregistered", "discontinued", "graduated", "graduand", "deferred", "on_leave"];

// ─── Helper: resolve the best status for a student on a specific unit ─────────
// Returns the grade status ("PASS", "SUPPLEMENTARY", "SPECIAL", null)
// across ALL sources: FinalGrade → MarkDirect → Mark.
// null means no mark exists for this unit.
async function _resolveUnitStatus(
  studentId: string,
  puId: string,
  passMark: number,
): Promise<{ status: string | null; mark: any | null }> {
  // Priority 1: FinalGrade (most authoritative)
  const fg = (await FinalGrade.findOne({student: studentId, programUnit: puId}).sort({ createdAt: -1 }).lean()) as any;
  if (fg) return { status: fg.status, mark: fg };

  // Priority 2: MarkDirect
  const md = (await MarkDirect.findOne({student: studentId, programUnit: puId}).sort({ createdAt: -1 }).lean()) as any;
  if (md) {
    const isSpecial = md.isSpecial || md.attempt === "special";
    const status = isSpecial ? "SPECIAL" : (md.agreedMark ?? 0) >= passMark ? "PASS" : "SUPPLEMENTARY";
    return { status, mark: md };
  }

  // Priority 3: Mark (detailed breakdown)
  const dm = (await Mark.findOne({ student: studentId, programUnit: puId })
    .sort({ createdAt: -1 }).lean()) as any;
  if (dm) {
    const isSpecial = dm.isSpecial || dm.attempt === "special";
    const status = isSpecial ? "SPECIAL" : (dm.agreedMark ?? 0) >= passMark ? "PASS" : "SUPPLEMENTARY";
    return { status, mark: dm };
  }

  return { status: null, mark: null };
}

// ─── Helper: find a SPECIAL mark from ANY prior year for this unit ─────────────
async function _findPriorSpecialMark(studentId: string, puId: string): Promise<any | null> {
  const fg = (await FinalGrade.findOne({
    student: studentId, programUnit: puId, status: "SPECIAL",
  }).sort({ createdAt: -1 }).lean()) as any;
  if (fg && (fg.caTotal30 ?? 0) > 0) return { ...fg, _source: "finalGrade" };

  const dm = (await Mark.findOne({
    student: studentId, programUnit: puId, isSpecial: true,
  }).sort({ createdAt: -1 }).lean()) as any;
  if (dm && (dm.caTotal30 ?? 0) > 0) return { ...dm, _source: "detailed" };

  const md = (await MarkDirect.findOne({
    student: studentId, programUnit: puId, attempt: "special",
  }).sort({ createdAt: -1 }).lean()) as any;
  if (md && (md.caTotal30 ?? 0) > 0) return { ...md, _source: "direct" };

  return null;
}

// ─── Helper: derive attempt label + flags from a mark and student ─────────────
function _flagsFromMark(
  mark: any | null, studentStatus: string, priorSpecial: any | null,
): Partial<ScoresheetStudent> & { attemptLabel: string } {
  if (!mark) {
    const isRepeat = studentStatus === "repeat";
    return { attemptLabel: isRepeat ? ATTEMPT.REPEAT_YEAR : ATTEMPT.FIRST, isRepeatYear: isRepeat};
  }

  const attempt = (mark.attempt || mark.attemptType || "1st").toLowerCase();
  const isSpecial = mark.isSpecial === true || attempt === "special" || mark.status === "SPECIAL" || mark.attemptType === "SPECIAL";
  const isSupp = attempt === "supplementary" || mark.status === "SUPPLEMENTARY" || mark.attemptType === "SUPPLEMENTARY";
  const isRetake = attempt === "re-take" || attempt === "retake" || mark.attemptType === "RETAKE";

  if (isSpecial) {
    return { attemptLabel: ATTEMPT.SPECIAL, isSpecial: true, isCarriedSpecial: priorSpecial != null, prevMark: priorSpecial ?? mark};
  }
  if (isSupp) return { attemptLabel: ATTEMPT.SUPPLEMENTARY, isSupp: true };
  if (isRetake) return { attemptLabel: ATTEMPT.RETAKE };
  if (studentStatus === "repeat")
    return { attemptLabel: ATTEMPT.REPEAT_YEAR, isRepeatYear: true };

  return { attemptLabel: ATTEMPT.FIRST };
}

// ─────────────────────────────────────────────────────────────────────────────
// CORE: Build the student list for a scoresheet
// ─────────────────────────────────────────────────────────────────────────────

export const buildScoresheetStudentList = async (params: {
  programId: mongoose.Types.ObjectId;
  programUnitId: mongoose.Types.ObjectId;
  unitId: mongoose.Types.ObjectId;
  yearOfStudy: number;
  academicYearId: mongoose.Types.ObjectId;
  session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED";
  passMark: number;
}): Promise<ScoresheetStudent[]> => {
  const {programId, programUnitId, yearOfStudy, academicYearId, session, passMark} = params;

  const academicYear = (await AcademicYear.findById(academicYearId).lean()) as any;
  if (!academicYear) throw new Error("Academic year not found");

  const result: ScoresheetStudent[] = [];
  const addedIds: Set<string> = new Set();

  const add = (
    student: any,
    attemptLabel: string,
    flags: Partial<ScoresheetStudent> = {},
    prevMark?: any,
  ) => {
    const id = student._id.toString();
    if (addedIds.has(id)) return;
    addedIds.add(id);
    result.push({
      regNo: student.regNo,
      displayRegNo: buildDisplayRegNo(student.regNo, student.qualifierSuffix),
      name: student.name,
      studentId: id,
      attemptLabel,
      isSupp: false,
      isSpecial: false,
      isCarriedSpecial: false,
      isCarryForward: false,
      isStayout: false,
      isRepeatYear: false,
      qualifierSuffix: student.qualifierSuffix || "",
      prevMark,
      ...flags,
    });
  };

  // ── ORDINARY SESSION ───────────────────────────────────────────────────────
  if (session === "ORDINARY" || session === "CLOSED") {
    // ── Primary pool: students who have a mark for this unit + academic year ──
    // This covers:
    //   (a) Current-year first sitters who have already had marks entered
    //   (b) Repeat-year students re-sitting
    //   (c) Historical scoresheets (marks exist from past)
    const [detailedIds, directIds] = await Promise.all([
      Mark.find({programUnit: programUnitId, academicYear: academicYearId}).distinct("student"),
      MarkDirect.find({programUnit: programUnitId, academicYear: academicYearId}).distinct("student"),
    ]);

    const allMarkedIds = [
      ...new Set([
        ...detailedIds.map((id: any) => id.toString()),
        ...directIds.map((id: any) => id.toString()),
      ]),
    ];

    if (allMarkedIds.length > 0) {
      const studentsWithMarks = (await Student.find({
        _id: { $in: allMarkedIds }, status: { $nin: EXCLUDED },
      }).sort({ regNo: 1 }).lean()) as any[];

      for (const student of studentsWithMarks) {
        const { mark } = await _resolveUnitStatus(student._id.toString(), programUnitId.toString(), passMark);
        const priorSpecial = mark?.isSpecial
          ? await _findPriorSpecialMark(student._id.toString(), programUnitId.toString())
          : null;
        const flags = _flagsFromMark(mark, student.status, priorSpecial);
        add(student, flags.attemptLabel, flags, flags.prevMark ?? mark ?? undefined);
      }
    }

    // ── Current-year fallback: no marks yet for this unit this year ───────────
    // Only runs for the current academic year (isCurrent).
    // Covers first-sitters and deferred-special students.
    if (academicYear.isCurrent) {
      const currentStudents = (await Student.find({
        program: programId, currentYearOfStudy: yearOfStudy, status: { $nin: EXCLUDED }, _id: { $nin: allMarkedIds },
      }).sort({ regNo: 1 }).lean()) as any[];

      for (const student of currentStudents) {
        // ── KEY FIX: Check if this student already PASSED this unit in a
        //    prior academic year. If so, they do NOT need to sit it again.
        const { status: priorStatus } = await _resolveUnitStatus(
          student._id.toString(), programUnitId.toString(), passMark,
        );

        if (priorStatus === "PASS") {
          // Already passed this unit — exclude from this scoresheet
          continue;
        }

        // Check for a deferred special (ENG.18c)
        const priorSpecial = await _findPriorSpecialMark(
          student._id.toString(), programUnitId.toString(),
        );

        if (priorSpecial) {
          // Deferred special — CA must be pre-populated and locked
          add(
            student,
            ATTEMPT.SPECIAL,
            {
              isSpecial: true, isCarriedSpecial: true,
            },
            priorSpecial,
          );
        } else if (priorStatus === "SUPPLEMENTARY") {
          // They failed this unit in a prior year and need to retake
          add(student, ATTEMPT.RETAKE, { isStayout: true });
        } else {
          // Genuinely new first sitter (or repeat year re-sitting all units)
          const isRepeat = student.status === "repeat";
          add(student, isRepeat ? ATTEMPT.REPEAT_YEAR : ATTEMPT.FIRST, {
            isRepeatYear: isRepeat,
          });
        }
      }
    }

    // ── Carry-forward students (ENG.14) ───────────────────────────────────────
    // These are students in a HIGHER year who are carrying THIS specific unit.
    // getCarryForwardStudentsForUnit already filters by programUnitId.
    try {
      const cfStudents = await getCarryForwardStudentsForUnit(
        programUnitId.toString(), programId.toString(),
      );
      for (const { student } of cfStudents) {
        if (EXCLUDED.includes(student.status)) continue;
        add(student, ATTEMPT.CF, { isCarryForward: true });
      }
    } catch {
      /* skip if CF service not available */
    }

    // ── Stayout students (ENG.15h) ────────────────────────────────────────────
    // Students who failed >1/3 <1/2 units and must retake in next ordinary.
    // getStayoutStudentsForUnit ALREADY filters by programUnitId — it only
    // returns students who have a FAILING FinalGrade for this specific unit.
    // The bug was NOT here — it was in the primary pool above (a stayout student
    // who sat all units had marks for all units, so appeared on all scoresheets).
    // Now that the primary pool only adds students who NEED to sit this unit,
    // stayout students appear only on scoresheets for their failed units.
    try {
      const stayoutStudents = await getStayoutStudentsForUnit(
        programUnitId.toString(),
        programId.toString(),
      );
      for (const { student } of stayoutStudents) {
        if (EXCLUDED.includes(student.status)) continue;
        add(student, ATTEMPT.STAYOUT, { isStayout: true });
      }
    } catch {
      /* skip */
    }

    // ── SUPPLEMENTARY SESSION ──────────────────────────────────────────────────
  } else if (session === "SUPPLEMENTARY") {
    // Only students with a FAILING FinalGrade for THIS specific unit + THIS year
    const failingGrades = (await FinalGrade.find({
      programUnit: programUnitId,
      academicYear: academicYearId,
      status: { $in: ["SUPPLEMENTARY", "SPECIAL"] },
    }).lean()) as any[];

    // Also check MarkDirect without FinalGrade (pre-backfill)
    const directFailing = (await MarkDirect.find({
      programUnit: programUnitId,
      academicYear: academicYearId,
      agreedMark: { $lt: passMark },
      attempt: { $ne: "special" },
    }).lean()) as any[];

    // Special students for this unit in this year
    const specialMarksDirect = (await MarkDirect.find({
      programUnit: programUnitId,
      academicYear: academicYearId,
      attempt: "special",
    }).lean()) as any[];
    const specialMarksDetailed = (await Mark.find({
      programUnit: programUnitId,
      academicYear: academicYearId,
      $or: [{ isSpecial: true }, { attempt: "special" }],
    }).lean()) as any[];

    const candidateIds = new Set<string>([
      ...failingGrades.map((g: any) => g.student.toString()),
      ...directFailing.map((m: any) => m.student.toString()),
      ...specialMarksDirect.map((m: any) => m.student.toString()),
      ...specialMarksDetailed.map((m: any) => m.student.toString()),
    ]);

    if (candidateIds.size > 0) {
      const candidates = (await Student.find({
        _id: { $in: [...candidateIds] }, status: { $nin: EXCLUDED },
      })
        .sort({ regNo: 1 }).lean()) as any[];

      const totalUnits = await ProgramUnit.countDocuments({
        program: programId, requiredYear: yearOfStudy,
      });

      for (const student of candidates) {
        const sid = student._id.toString();

        // Determine if special for this specific unit
        const gradeForUnit = failingGrades.find(
          (g: any) => g.student.toString() === sid,
        );
        const isSpecialUnit =
          gradeForUnit?.status === "SPECIAL" ||
          gradeForUnit?.isSpecial ||
          specialMarksDirect.some((m: any) => m.student.toString() === sid) ||
          specialMarksDetailed.some((m: any) => m.student.toString() === sid);

        if (!isSpecialUnit) {
          // ENG.15h: >1/3 <1/2 failed overall → stayout, not supp
          const failedCount = await FinalGrade.countDocuments({
            student: student._id,
            academicYear: academicYearId,
            status: "SUPPLEMENTARY",
          });
          const failedFrac = totalUnits > 0 ? failedCount / totalUnits : 0;
          if (failedFrac > 1 / 3 && failedFrac < 1 / 2) continue;

          // ENG.16: ≥1/2 → repeat year
          if (failedFrac >= 1 / 2 || student.status === "repeat") continue;
        }

        let prevMark: any = null;
        if (isSpecialUnit) {
          prevMark = await _findPriorSpecialMark(sid, programUnitId.toString());
        }

        add(
          student,
          isSpecialUnit ? ATTEMPT.SPECIAL : ATTEMPT.SUPPLEMENTARY,
          {
            isSupp: !isSpecialUnit,
            isSpecial: isSpecialUnit,
          },
          prevMark ?? undefined,
        );
      }
    }
  }

  // Sort: primary first, then CF, then stayout, within each group by regNo
  return result.sort((a, b) => {
    const aPriority = a.isCarryForward ? 2 : a.isStayout ? 3 : 1;
    const bPriority = b.isCarryForward ? 2 : b.isStayout ? 3 : 1;
    if (aPriority !== bPriority) return aPriority - bPriority;
    return a.regNo.localeCompare(b.regNo);
  });
};

// ─────────────────────────────────────────────────────────────────────────────
// HELPER: Pre-existing marks for cell pre-population in templates
// Looks across ALL academic years so deferred-special CA marks are found.
// ─────────────────────────────────────────────────────────────────────────────

export const getExistingMarksForStudents = async (
  studentIds: string[],
  programUnitId: mongoose.Types.ObjectId,
): Promise<Map<string, any>> => {
  const [detailed, direct] = await Promise.all([
    Mark.find({ student: { $in: studentIds }, programUnit: programUnitId })
      .sort({ createdAt: -1 }).lean(),
    MarkDirect.find({ student: { $in: studentIds }, programUnit: programUnitId })
      .sort({ createdAt: -1 }).lean(),
  ]);

  const map = new Map<string, any>();
  for (const m of direct) {
    const key = m.student.toString();
    if (!map.has(key)) map.set(key, { ...m, source: "direct" });
  }
  for (const m of detailed) {
    const key = m.student.toString();
    const existing = map.get(key);
    // Prefer special marks or more recent detailed over direct
    if (!existing || (m as any).isSpecial || !existing.isSpecial) {
      map.set(key, { ...m, source: "detailed" });
    }
  }
  return map;
};

export function buildRichRegNo(
  regNo: string, qualifier: string | undefined | null, fontName = "Book Antiqua", baseSize = 8,
): ExcelJS.CellRichTextValue | string {
  if (!qualifier || qualifier.trim() === "") return regNo;

  return {
    richText: [
      { text: regNo, font: { name: fontName, size: baseSize, bold: false }},
      { text: qualifier.trim(), font: { name: fontName, size: Math.max(baseSize - 2, 5), vertAlign: "subscript",  color: { argb: "FF002B1B" }}},
      // { text: qualifier.trim(), font: { name: fontName, size: Math.max(baseSize - 2, 5), vertAlign: "subscript"}},
    ],
  };
}

/**
 * CMS variant — Arial font, slightly larger base size.
 */
export function buildRichRegNoCMS(
  regNo: string,
  qualifier: string | undefined | null
): ExcelJS.CellRichTextValue | string {
  return buildRichRegNo(regNo, qualifier, "Times New Roman", 8);
}