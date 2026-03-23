// serverside/src/jobs/defermentCron.ts
import cron from "node-cron";
import Student from "../models/Student";
import AcademicYear from "../models/AcademicYear";

export const startStatusReversionJob = () => {
  cron.schedule("0 1 * * *", async () => {
    const now = new Date();

    try {
      const expiredStudents = await Student.find({ status: { $in: ["deferred", "on_leave"] }, "academicLeavePeriod.endDate": { $lte: now }});

      if (expiredStudents.length === 0) return;

      const updatePromises = expiredStudents.map(async (student) => {
        const previousStatus = student.status;

        // Use existing history to find the relevant academic year
        const lastHistoryEntry = student.academicHistory?.slice(-1)[0];

        const currentYearDoc = await AcademicYear.findOne({ isCurrent: true }).session(null).lean();

        const currentAY =
          currentYearDoc?.year || // real year from DB
          lastHistoryEntry?.academicYear || // last known year from student history
          new Date().getFullYear() + "/" + (new Date().getFullYear() + 1); // absolute last resort

        return Student.findByIdAndUpdate(student._id, {
          $set: {
            status: "active",
            remarks: `System: Auto-reinstated from ${previousStatus.toUpperCase()} on ${now.toDateString()}.`,
          },
          $push: {
            statusEvents: {
              fromStatus: previousStatus,
              toStatus: "active",
              date: now,
              academicYear: currentAY, // now a real year string
              reason: `SYSTEM AUTO-REINSTATEMENT: End of authorized ${previousStatus.replace("_", " ")} period.`,
            },
            academicHistory: {
              academicYear: currentAY, // now a real year string
              yearOfStudy: student.currentYearOfStudy,
              annualMeanMark: lastHistoryEntry?.annualMeanMark || 0,
              weightedContribution: lastHistoryEntry?.weightedContribution || 0,
              failedUnitsCount: lastHistoryEntry?.failedUnitsCount || 0,
              date: now,
            },
          },
          $unset: { academicLeavePeriod: 1 },
        });
      });

      await Promise.all(updatePromises);
      console.log(`[Cron] Reactivated ${expiredStudents.length} students. Updated academicHistory and statusEvents.`);
    } catch (error) {
      console.error("[Cron] Error in Status Reversion Job:", error);
    }
  });
};
