// serverside/src/services/historicalYearValidator.ts
//
// ════════════════════════════════════════════════════════════════════════════
// ENG.15(b) — FULL HISTORICAL VALIDATION BEFORE YEAR 5 ENTRY
// ════════════════════════════════════════════════════════════════════════════
//
// THE PROBLEM WITH THE EXISTING CODE
// ────────────────────────────────────────────────────────────────────────────
// promoteStudent() in statusEngine.ts checks whether a student can move from
// Year 4 → Year 5. It calls calculateStudentStatus() for the CURRENT year
// only. This is correct for Year 1→2, 2→3, 3→4 because those transitions
// only depend on the current year's performance.
//
// But ENG.15(b) says:
//   "A student shall not be allowed to sit for final year examinations unless
//    they have passed all the examinations of the preceding years."
//
// "Preceding years" means Year 1, Year 2, Year 3, AND Year 4 — every unit,
// in every year, must be conclusively passed before Year 5 is granted.
// The current code only checks Year 4 → Year 5 transition marks. A student
// who has an unresolved carry-forward from Year 2 that was never cleared
// would still be promoted to Year 5 under the old logic.
//
// WHAT "PASSED" MEANS HERE (not obvious)
// ────────────────────────────────────────────────────────────────────────────
// A unit is considered passed for ENG.15(b) purposes if:
//   1. FinalGrade.status = "PASS" exists for this student + programUnit, OR
//   2. The unit is in carryForwardUnits with status = "cleared", OR
//   3. The student received a "DISCONTINUED" verdict on the unit ladder
//      (in which case they can't be in Year 5 anyway — caught by other checks)
//
// A unit is BLOCKING if:
//   1. It has FinalGrade.status = "SUPPLEMENTARY" with no subsequent PASS, OR
//   2. It is in carryForwardUnits with status = "pending", OR
//   3. It has no FinalGrade record at all (genuinely INC — missing mark)
//
// Note on carry-forwards (ENG.14): CF units are explicitly allowed to proceed
// to the next year. So a Y2 CF unit that is pending in Y3 is NOT a block — it
// is expected to be examined in Y3's ordinary/supp. The block only applies
// when that unit reaches Y5 entry and STILL hasn't been cleared.
//
// HOW THIS INTEGRATES WITH promoteStudent()
// ────────────────────────────────────────────────────────────────────────────
// In promoteStudent() (statusEngine.ts), the Y4→Y5 block has this check:
//
//   if (currentYear === duration - 1) {   // currentYear = 4 in a 5-year programme
//     // EXISTING: only checks Y4 status
//     const statusResult = await calculateStudentStatus(...)
//
//   AFTER THIS CHANGE:
//     const histCheck = await validateHistoricalYears(studentId, programId, duration - 1)
//     if (!histCheck.canEnterFinalYear) {
//       return { success: false, message: histCheck.blockReason, blockDetails: histCheck.blockingUnits }
//     }
//     // then proceed with Y4 current-year check as normal
//   }
//
// This means the historical check runs BEFORE the current-year engine.
// If historical years are clean, the existing Y4 engine runs normally.
// If historical years have unresolved units, we block immediately with a
// detailed list of exactly which units from which year are the problem.

import FinalGrade from "../models/FinalGrade";
import ProgramUnit from "../models/ProgramUnit";
import Student from "../models/Student";

// ── Types ──────────────────────────────────────────────────────────────────────

export interface BlockingUnit {
  yearOfStudy: number;
  unitCode: string;
  unitName: string;
  reason:
    | "UNCLEARED_SUPPLEMENTARY"
    | "PENDING_CARRY_FORWARD"
    | "NO_MARK_RECORD";
  // The attempt history for this unit — useful for the UI to show context
  attemptCount: number;
}

export interface HistoricalYearValidation {
  canEnterFinalYear: boolean;
  blockReason: string;
  blockingUnits: BlockingUnit[];
  // Clean years list — shown as green checks in the UI
  cleanYears: number[];
  // Summary per year: { year: 1, passed: 8, total: 8, isClean: true }
  yearSummaries: Array<{
    yearOfStudy: number;
    totalUnits: number;
    passed: number;
    pending: number;
    isClean: boolean;
  }>;
}

// ── Main validation function ───────────────────────────────────────────────────

export async function validateHistoricalYears(
  studentId: string,
  programId: string,
  yearsToCheck: number, // For a 5-year programme: pass 4 (check Y1–Y4)
): Promise<HistoricalYearValidation> {
  const student = (await Student.findById(studentId).lean()) as any;
  if (!student) {
    return {
      canEnterFinalYear: false,
      blockReason: "Student record not found.",
      blockingUnits: [],
      cleanYears: [],
      yearSummaries: [],
    };
  }

  // Pending carry-forward units from the Student model
  // carryForwardUnits[].status = "pending" | "cleared" | "expired"
  const pendingCFCodes = new Set<string>(
    ((student.carryForwardUnits ?? []) as any[])
      .filter((u: any) => u.status === "pending")
      .map((u: any) => u.unitCode as string),
  );

  const blockingUnits: BlockingUnit[] = [];
  const cleanYears: number[] = [];
  const yearSummaries: HistoricalYearValidation["yearSummaries"] = [];

  // Check each year 1 → yearsToCheck inclusive
  for (let y = 1; y <= yearsToCheck; y++) {
    // All programUnits prescribed for this year
    const pus = (await ProgramUnit.find({
      program: programId,
      requiredYear: y,
    })
      .populate("unit")
      .lean()) as any[];

    if (pus.length === 0) {
      // No units defined for this year — treat as clean (data gap, not a rule failure)
      cleanYears.push(y);
      yearSummaries.push({
        yearOfStudy: y,
        totalUnits: 0,
        passed: 0,
        pending: 0,
        isClean: true,
      });
      continue;
    }

    let yearPending = 0;

    for (const pu of pus) {
      const puId = pu._id.toString();
      const unitCode = pu.unit?.code ?? "N/A";
      const unitName = pu.unit?.name ?? "Unknown Unit";

      // ── Check 1: Is there a PASS FinalGrade for this unit? ──────────────
      // We look across ALL academic years — a Y2 unit passed in Y3 (retake)
      // still counts as cleared for ENG.15(b).
      const passingGrade = await FinalGrade.findOne({
        student: studentId,
        programUnit: puId,
        status: "PASS",
      }).lean();

      if (passingGrade) {
        // Unit is conclusively passed — not a block regardless of carry-forward
        continue;
      }

      // ── Check 2: Is this unit a pending carry-forward? ───────────────────
      // Pending CF units are allowed to be examined later — they only become
      // a Y5 block if still pending AT THE TIME of Y4→Y5 promotion.
      if (pendingCFCodes.has(unitCode)) {
        blockingUnits.push({
          yearOfStudy: y,
          unitCode,
          unitName,
          reason: "PENDING_CARRY_FORWARD",
          attemptCount: await getAttemptCount(studentId, puId),
        });
        yearPending++;
        continue;
      }

      // ── Check 3: Is there an uncleared SUPPLEMENTARY grade? ─────────────
      const suppGrade = await FinalGrade.findOne({
        student: studentId,
        programUnit: puId,
        status: { $in: ["SUPPLEMENTARY", "RETAKE"] },
      })
        .sort({ createdAt: -1 })
        .lean();

      if (suppGrade) {
        blockingUnits.push({
          yearOfStudy: y,
          unitCode,
          unitName,
          reason: "UNCLEARED_SUPPLEMENTARY",
          attemptCount: await getAttemptCount(studentId, puId),
        });
        yearPending++;
        continue;
      }

      // ── Check 4: No grade record at all ─────────────────────────────────
      // This student was never entered for the unit. Could be a data entry gap
      // or a genuine missing mark. Either way, it blocks ENG.15(b).
      const anyGrade = await FinalGrade.findOne({
        student: studentId,
        programUnit: puId,
      }).lean();

      if (!anyGrade) {
        // Only flag as blocking if the student's history shows they WERE in this
        // year. If they haven't reached this year yet, don't flag it.
        const wasInYear = ((student.academicHistory ?? []) as any[]).some(
          (h: any) => h.yearOfStudy === y,
        );
        if (wasInYear) {
          blockingUnits.push({
            yearOfStudy: y,
            unitCode,
            unitName,
            reason: "NO_MARK_RECORD",
            attemptCount: 0,
          });
          yearPending++;
        }
      }
    }

    const isClean = yearPending === 0;
    if (isClean) cleanYears.push(y);

    yearSummaries.push({
      yearOfStudy: y,
      totalUnits: pus.length,
      passed: pus.length - yearPending,
      pending: yearPending,
      isClean,
    });
  }

  const canEnterFinalYear = blockingUnits.length === 0;

  // Construct a human-readable block reason for the UI and audit log
  const blockReason = canEnterFinalYear ? "" : buildBlockReason(blockingUnits);

  return {
    canEnterFinalYear,
    blockReason,
    blockingUnits,
    cleanYears,
    yearSummaries,
  };
}

// ── Helpers ───────────────────────────────────────────────────────────────────

async function getAttemptCount(
  studentId: string,
  puId: string,
): Promise<number> {
  return FinalGrade.countDocuments({ student: studentId, programUnit: puId });
}

function buildBlockReason(units: BlockingUnit[]): string {
  const byYear = new Map<number, BlockingUnit[]>();
  for (const u of units) {
    if (!byYear.has(u.yearOfStudy)) byYear.set(u.yearOfStudy, []);
    byYear.get(u.yearOfStudy)!.push(u);
  }

  const parts: string[] = [];
  byYear.forEach((us, yr) => {
    const codes = us.map((u) => u.unitCode).join(", ");
    parts.push(`Year ${yr}: ${codes}`);
  });

  return (
    `ENG.15(b) VIOLATION — Cannot enter final year. ` +
    `The following units from prior years are unresolved: ${parts.join(" | ")}. ` +
    `All preceding year examinations must be conclusively passed before Year 5 entry.`
  );
}

// ── Patch to apply in promoteStudent() ───────────────────────────────────────
//
// In statusEngine.ts, inside promoteStudent(), BEFORE the existing Y4 check:
//
// FIND this block (around line ~42900):
//   if (currentYear === duration) {   // final year graduation block
//
// ADD ABOVE IT (so it runs when currentYear = duration - 1, i.e. Y4 for 5yr):
//
//   // ENG.15(b): Historical year validation before final year entry
//   if (currentYear === duration - 1) {
//     const { validateHistoricalYears } = await import("./historicalYearValidator");
//     const histCheck = await validateHistoricalYears(
//       studentId.toString(),
//       student.program.toString(),
//       duration - 1,  // Check all years up to and including Y4
//     );
//     if (!histCheck.canEnterFinalYear) {
//       // Log the block for the audit trail
//       await Student.findByIdAndUpdate(studentId, {
//         $push: {
//           statusEvents: {
//             fromStatus:  st,
//             toStatus:    st,
//             date:        new Date(),
//             academicYear: completedYear,
//             reason:      histCheck.blockReason,
//           },
//         },
//       });
//       return {
//         success: false,
//         message: histCheck.blockReason,
//         eng15bBlock: true,
//         blockingUnits: histCheck.blockingUnits,
//         yearSummaries: histCheck.yearSummaries,
//       };
//     }
//   }
//
// ── Patch to apply in previewPromotion() ─────────────────────────────────────
//
// In previewPromotion(), inside the student loop AFTER the adminLabel check,
// add a similar early-exit before calculateStudentStatus():
//
//   if (yearToPromote === duration - 1) {
//     const { validateHistoricalYears } = await import("./historicalYearValidator");
//     const histCheck = await validateHistoricalYears(
//       (student as any)._id.toString(),
//       programId,
//       duration - 1,
//     );
//     if (!histCheck.canEnterFinalYear) {
//       blocked.push({
//         ...report,
//         status:  "ENG.15(b) BLOCK",
//         reasons: [histCheck.blockReason],
//         eng15bBlock: true,
//         blockingUnits: histCheck.blockingUnits,
//         yearSummaries: histCheck.yearSummaries,
//       });
//       continue;
//     }
//   }
