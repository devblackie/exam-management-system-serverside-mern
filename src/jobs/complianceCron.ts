// serverside/src/jobs/complianceCron.ts — COMPLETE, CORRECT VERSION
import cron from "node-cron";
import Student from "../models/Student";
import Program from "../models/Program";
import { loadInstitutionSettings } from "../utils/loadInstitutionSettings";

interface StudentWithProgram {
  _id: unknown;
  regNo: string;
  status: string;
  programType?: string;
  totalTimeOutYears?: number;
  institution?: unknown;
  admissionAcademicYear?: { year?: string };
  program?: {
    _id: unknown;
    institution: unknown;
    durationYears: number;
  };
}

export const startComplianceCronJobs = (): void => {
  // ── ENG.19(d/e): Auto-discontinue students exceeding max study duration ────
  // Runs daily at 02:00
  cron.schedule("0 2 * * *", async () => {
    console.log("[Cron] ENG.19 duration check starting…");
    const now = new Date();
    const nowYear = now.getFullYear();

    const activeStudents = await Student.find({
      status: { $in: ["active", "repeat", "on_leave", "deferred"] },
    })
      .populate<{ admissionAcademicYear: { year: string } }>(
        "admissionAcademicYear",
        "year",
      )
      .populate<{
        program: { _id: unknown; institution: unknown; durationYears: number };
      }>("program", "institution durationYears")
      .lean<StudentWithProgram[]>();

    let discontinued = 0;

    for (const student of activeStudents) {
      const yearStr = (student.admissionAcademicYear as any)?.year ?? "";
      const admissionYear = parseInt(yearStr.split("/")[0]);
      if (!admissionYear || isNaN(admissionYear)) continue;

      const institutionId =
        (student.program as any)?.institution?.toString() ??
        (student.institution as any)?.toString();
      if (!institutionId) continue;

      // ── THE CORRECT APPROACH ──────────────────────────────────────────────
      // Use Program.durationYears × settings.ruleSet.maxDurationMultiplier
      // NEVER infer from degree name ("Financial Engineering" is BSc but 4yr)
      const programDuration = (student.program as any)?.durationYears ?? 5;
      const settings = await loadInstitutionSettings(institutionId).catch(
        () => null,
      );
      const multiplier = settings?.rules.maxDurationMultiplier ?? 2.0;
      const maxYears = Math.round(programDuration * multiplier);

      // Subtract approved leave years
      const effectiveYears =
        nowYear - admissionYear - (student.totalTimeOutYears ?? 0);

      if (effectiveYears > maxYears) {
        const rule = `ENG.19(${programDuration > 4 ? "d" : "e"})`;
        await Student.findByIdAndUpdate(student._id, {
          $set: { status: "discontinued" },
          $push: {
            statusEvents: {
              fromStatus: student.status,
              toStatus: "discontinued",
              date: now,
              academicYear: yearStr,
              reason: `AUTO [${rule}]: ${effectiveYears} effective years exceeds ${maxYears}-year maximum (${programDuration}yr program × ${multiplier} multiplier)`,
            },
          },
        });
        discontinued++;
        console.log(
          `[Cron] ${rule} auto-discontinued: ${student.regNo} (${effectiveYears}y > ${maxYears}y)`,
        );
      }
    }

    console.log(
      `[Cron] ENG.19 check complete — ${discontinued} students discontinued`,
    );
  });

  console.log("[Cron] Compliance jobs registered");
};
