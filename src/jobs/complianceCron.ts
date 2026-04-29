// serverside/src/jobs/complianceCron.ts — REPLACE the duration check section
import cron from "node-cron";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";

// ENG.19(d/e): BSc Engineering = max 10 years; BEd = max 8 years
const MAX_YEARS_BY_TYPE: Record<string, number> = {
  BSc: 10,
  BEd: 8,
  BTech: 10,
  BEng: 10,
};

function getMaxYears(programType: string): number {
  // programType is stored on Student e.g. "BSc", "BEd", "Direct", "Mid-Entry-Y2"
  // Match against the degree prefix
  for (const [prefix, max] of Object.entries(MAX_YEARS_BY_TYPE)) {
    if (programType?.toUpperCase().startsWith(prefix.toUpperCase())) return max;
  }
  return 10; // safe default for unknown types
}

export const startComplianceCronJobs = () => {
  // ── ENG.19(d/e): Time-limit auto-discontinue ─────────────────────────────
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
      .lean();

    let discontinued = 0;

    for (const student of activeStudents) {
      const yearStr = student.admissionAcademicYear?.year ?? ""; // e.g. "2014/2015"
      const admissionYear = parseInt(yearStr.split("/")[0]);
      if (!admissionYear || isNaN(admissionYear)) continue;

      // Subtract any approved leave years (stored as totalTimeOutYears)
      const effectiveYears =
        nowYear - admissionYear - (student.totalTimeOutYears ?? 0);
      const maxYears = getMaxYears(student.programType ?? "BSc");

      if (effectiveYears > maxYears) {
        const rule = maxYears === 8 ? "ENG.19(e)" : "ENG.19(d)";
        await Student.findByIdAndUpdate(student._id, {
          $set: { status: "discontinued" },
          $push: {
            statusEvents: {
              fromStatus: student.status,
              toStatus: "discontinued",
              date: now,
              academicYear: yearStr,
              reason: `AUTO [${rule}]: ${effectiveYears} effective years exceeds ${maxYears}-year maximum`,
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
