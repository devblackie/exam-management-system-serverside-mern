// serverside/src/utils/scoresheetStudentList.ts
// Single source of truth for which students appear on which scoresheet.
// Both directTemplate.ts and uploadTemplate.ts import from here.

import mongoose from "mongoose";
import Student from "../models/Student";
import Mark from "../models/Mark";
import MarkDirect from "../models/MarkDirect";
import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import AcademicYear from "../models/AcademicYear";
import { buildDisplayRegNo, ATTEMPT_LABELS } from "./academicRules";
import { getCarryForwardStudentsForUnit, getStayoutStudentsForUnit } from "../services/carryForwardService";
import { calculateStudentStatus } from "../services/statusEngine";

export interface ScoresheetStudent {
  regNo: string; displayRegNo: string; // regNo + qualifierSuffix (e.g. E024-01-1339/2016RP1)
  name: string; studentId: string; attemptLabel: string; // B/S, A/S, SPEC, A/SO, RP1C, RPU1 etc.
  isSupp: boolean; isSpecial: boolean; isCarryForward: boolean; isStayout: boolean; isRepeatYear: boolean;
}

// ─── buildScoresheetStudentList ───────────────────────────────────────────────

export const buildScoresheetStudentList = async (params: {
  programId: mongoose.Types.ObjectId; programUnitId: mongoose.Types.ObjectId;
  unitId: mongoose.Types.ObjectId; yearOfStudy: number; academicYearId: mongoose.Types.ObjectId;
  session: "ORDINARY" | "SUPPLEMENTARY" | "CLOSED"; passMark: number;
}): Promise<ScoresheetStudent[]> => {
  const { programId, programUnitId, yearOfStudy, academicYearId, session, passMark } = params;

  const academicYear = (await AcademicYear.findById(academicYearId).lean()) as any;
  if (!academicYear) throw new Error("Academic year not found");

  const EXCLUDED = ["deregistered", "discontinued", "graduated", "graduand", "deferred", "on_leave"];

  const result: ScoresheetStudent[] = [];
  const addedIds = new Set<string>();

  const addStudent = (student: any, attemptLabel: string, flags: Partial<ScoresheetStudent>) => {
    const id = student._id.toString();
    if (addedIds.has(id)) return;
    addedIds.add(id);
    result.push({
      regNo: student.regNo, displayRegNo: buildDisplayRegNo(student.regNo, student.qualifierSuffix),
      name: student.name, studentId: id, attemptLabel, isSupp: false,
      isSpecial: false, isCarryForward: false, isStayout: false, isRepeatYear: false, ...flags
    });
  };

  if (session === "ORDINARY") {
    // 1. Primary pool: active + repeat students in this yearOfStudy
    const primary = (await Student.find({
      program: programId, currentYearOfStudy: yearOfStudy, status: { $nin: EXCLUDED },
    }).sort({ regNo: 1 }).lean()) as any[];

    for (const s of primary) {
      const isRepeat = s.status === "repeat";
      addStudent(
        s,
        isRepeat ? ATTEMPT_LABELS.REPEAT_YEAR_1 : ATTEMPT_LABELS.ORDINARY,
        { isRepeatYear: isRepeat },
      );
    }

    // 2. Carry-forward students (ENG.14) in a higher year but carrying THIS unit
    const cfStudents = await getCarryForwardStudentsForUnit(
      programUnitId.toString(), programId.toString());
    for (const { student, attemptLabel } of cfStudents) {
      if (EXCLUDED.includes(student.status)) continue;
      addStudent(student, attemptLabel, { isCarryForward: true });
    }

    // 3. Stayout students (ENG.15h) retaking in next ordinary
    const stayoutStudents = await getStayoutStudentsForUnit(
      programUnitId.toString(), programId.toString());
    for (const { student, attemptLabel } of stayoutStudents) {
      if (EXCLUDED.includes(student.status)) continue;
      addStudent(student, attemptLabel, { isStayout: true });
    }
  } else {
    // SUPPLEMENTARY session
    // WHO qualifies:
    //   ✅ Failed THIS unit AND failFraction ≤ 1/3 overall (ENG.13a)
    //   ✅ Has an approved SPECIAL for THIS unit (ENG.18)
    // WHO is excluded:
    //   ❌ status:"repeat" (they join ORDINARY next year, B/S)
    //   ❌ failFraction > 1/3 STAYOUT (they retake in next ORDINARY, A/SO)
    //   ❌ All EXCLUDED_STATUSES

    const candidates = (await Student.find({
      program: programId,
      status: { $nin: EXCLUDED },
      $or: [
        { currentYearOfStudy: yearOfStudy },
        { "academicHistory.yearOfStudy": yearOfStudy },
      ],
    })
      .sort({ regNo: 1 })
      .lean()) as any[];

    for (const s of candidates) {
      if (s.status === "repeat") continue; // repeat year → ORDINARY next year

      // Check if this student has a failing grade for THIS specific unit
      const gradeForUnit = (await FinalGrade.findOne({
        student: s._id, programUnit: programUnitId,
        status: { $in: ["SUPPLEMENTARY", "SPECIAL", "INCOMPLETE"] },
      }).lean()) as any;

      if (!gradeForUnit) continue;

      const isSpecial = gradeForUnit.isSpecial === true || gradeForUnit.status === "SPECIAL";

      // Non-special students: check overall fail fraction
      if (!isSpecial) {
        const sr = await calculateStudentStatus(
          s._id, programId, academicYear.year, yearOfStudy, { forPromotion: false },
        );
        const totalUnits = sr.summary.totalExpected;
        const failFrac = totalUnits > 0 ? sr.summary.failed / totalUnits : 0;

        // STAYOUT (>1/3 <1/2): retake next ORDINARY, NOT supp
        if (failFrac > 1 / 3 && sr.status !== "REPEAT YEAR") continue;

        // REPEAT YEAR (≥1/2 or mean<40%): excluded from supp
        if (sr.status === "REPEAT YEAR") continue;
      }

      addStudent(
        s,
        isSpecial ? ATTEMPT_LABELS.SPECIAL : ATTEMPT_LABELS.SUPPLEMENTARY,
        { isSupp: !isSpecial, isSpecial },
      );
    }
  }

  // Sort: primary cohort first (not CF, not stayout), then by regNo
  return result.sort((a, b) => {
    const aOther = a.isCarryForward || a.isStayout ? 1 : 0;
    const bOther = b.isCarryForward || b.isStayout ? 1 : 0;
    if (aOther !== bOther) return aOther - bOther;
    return a.regNo.localeCompare(b.regNo);
  });
};

// ─── getExistingMarksForStudents ──────────────────────────────────────────────
// Used by template generators to pre-populate cells with existing marks.

export const getExistingMarksForStudents = async (
  studentIds: string[], programUnitId: mongoose.Types.ObjectId ): Promise<Map<string, any>> => {
  const [detailed, direct] = await Promise.all([
    Mark.find({ student: { $in: studentIds }, programUnit: programUnitId }).lean(),
    MarkDirect.find({ student: { $in: studentIds }, programUnit: programUnitId }).lean(),
  ]);

  const map = new Map<string, any>();
  direct.forEach((m: any) => map.set(m.student.toString(), { ...m, source: "direct" }));
  detailed.forEach((m: any) => map.set(m.student.toString(), { ...m, source: "detailed" })); // detailed wins
  return map;
};
