// import cron from "node-cron";
// import Student from "../models/Student";

// /**
//  * PRODUCTION-READY: Automatic Status Reversion Job
//  * Runs daily at 1:00 AM to check for expired leaves and deferments.
//  * ENG. 19 & 20 compliance.
//  */
// export const startStatusReversionJob = () => {
//   // Cron syntax: "minute hour day-of-month month day-of-week"
//   // "0 1 * * *" = 1:00 AM every day
//   cron.schedule("0 1 * * *", async () => {
//     console.log(
//       `[Cron] Running Status Reversion Job at ${new Date().toISOString()}`,
//     );

//     try {
//       const now = new Date();

//       // 1. Find students whose status is not 'active' and their leave/deferment date is due or past
//       // 2. Set status to 'active'
//       // 3. Remove the academicLeavePeriod object to clean up the data
//       const result = await Student.updateMany(
//         {
//           status: { $in: ["deferred", "on_leave"] },
//           "academicLeavePeriod.endDate": { $lte: now },
//         },
//         {
//           $set: {
//             status: "active",
//             remarks: `Automatically activated on ${now.toISOString().split("T")[0]}.`,
//           },
//           $unset: { academicLeavePeriod: 1 }, // Cleans up the dates
//         },
//       );

//       console.log(
//         `[Cron] Successfully updated ${result.modifiedCount} student(s) to active.`,
//       );
//     } catch (error) {
//       console.error("[Cron] Error updating student statuses:", error);
//     }
//   });

//   console.log("[Cron] Status Reversion Job Scheduled.");
// };

// import cron from "node-cron";
// import Student from "../models/Student";

// /**
//  * AUTOMATED STATUS REVERSION WITH AUDIT TRAIL
//  * Runs daily at 1:00 AM.
//  * Reverts 'deferred' or 'on_leave' students to 'active' once endDate is reached.
//  */
// export const startStatusReversionJob = () => {
//   cron.schedule("0 1 * * *", async () => {
//     const now = new Date();
//     console.log(
//       `[Cron] Checking for expired leaves/deferments: ${now.toISOString()}`,
//     );

//     try {
//       // 1. Find all students whose leave/deferment has ended
//       const expiredStudents = await Student.find({
//         status: { $in: ["deferred", "on_leave"] },
//         "academicLeavePeriod.endDate": { $lte: now },
//       });

//       if (expiredStudents.length === 0) {
//         console.log("[Cron] No expired statuses found today.");
//         return;
//       }

//       // 2. Process each student to preserve Journey Timeline history
//       const updatePromises = expiredStudents.map((student) => {
//         const previousStatus = student.status.toUpperCase();

//         return Student.findByIdAndUpdate(student._id, {
//           $set: {
//             status: "active",
//             remarks: `Automatically resumed studies after ${previousStatus} on ${now.toDateString()}.`,
//           },
//           $push: {
//             promotionHistory: {
//               from: student.currentYearOfStudy,
//               to: student.currentYearOfStudy,
//               date: now,
//               remarks: `SYSTEM: RETURNED FROM ${previousStatus}`,
//             },
//           },
//           $unset: { academicLeavePeriod: 1 },
//         });
//       });

//       await Promise.all(updatePromises);

//       console.log(
//         `[Cron] Successfully reactivated ${expiredStudents.length} student(s).`,
//       );
//     } catch (error) {
//       console.error("[Cron] Critical error in Status Reversion Job:", error);
//     }
//   });

//   console.log(
//     "[Cron] Status Reversion Job with Timeline Integration Scheduled.",
//   );
// };

import cron from "node-cron";
import Student from "../models/Student";

export const startStatusReversionJob = () => {
  cron.schedule("0 1 * * *", async () => {
    const now = new Date();

    try {
      const expiredStudents = await Student.find({
        status: { $in: ["deferred", "on_leave"] },
        "academicLeavePeriod.endDate": { $lte: now },
      });

      if (expiredStudents.length === 0) return;

      const updatePromises = expiredStudents.map((student) => {
        const previousStatus = student.status;

        // Use existing history to find the relevant academic year
        const lastHistoryEntry = student.academicHistory?.slice(-1)[0];
        const currentAY = lastHistoryEntry?.academicYear || "Current Cycle";

        return Student.findByIdAndUpdate(student._id, {
          $set: {
            status: "active",
            remarks: `System: Auto-reinstated from ${previousStatus.toUpperCase()} on ${now.toDateString()}.`,
          },
          $push: {
            // 1. Permanent Ledger for Journey UI (Reinstatement Event)
            statusEvents: {
              fromStatus: previousStatus,
              toStatus: "active",
              date: now,
              academicYear: currentAY,
              reason: `SYSTEM AUTO-REINSTATEMENT: End of authorized ${previousStatus.replace("_", " ")} period.`,
            },
            // 2. Aligning with your Model's academicHistory (ENG 25.b)
            // We record that the student returned during this specific Year of Study
            academicHistory: {
              academicYear: currentAY,
              yearOfStudy: student.currentYearOfStudy,
              annualMeanMark: lastHistoryEntry?.annualMeanMark || 0,
              weightedContribution: lastHistoryEntry?.weightedContribution || 0,
              failedUnitsCount: lastHistoryEntry?.failedUnitsCount || 0,
              date: now, // Critical for sorting the Journey Timeline
            },
          },
          $unset: { academicLeavePeriod: 1 },
        });
      });

      await Promise.all(updatePromises);
      console.log(
        `[Cron] Reactivated ${expiredStudents.length} students. Updated academicHistory and statusEvents.`,
      );
    } catch (error) {
      console.error("[Cron] Error in Status Reversion Job:", error);
    }
  });
};
