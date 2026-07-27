"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.startComplianceCronJobs = void 0;
// serverside/src/jobs/complianceCron.ts — COMPLETE, CORRECT VERSION
const node_cron_1 = __importDefault(require("node-cron"));
const Student_1 = __importDefault(require("../models/Student"));
const loadInstitutionSettings_1 = require("../utils/loadInstitutionSettings");
const startComplianceCronJobs = () => {
    // ── ENG.19(d/e): Auto-discontinue students exceeding max study duration ────
    // Runs daily at 02:00
    node_cron_1.default.schedule("0 2 * * *", async () => {
        console.log("[Cron] ENG.19 duration check starting…");
        const now = new Date();
        const nowYear = now.getFullYear();
        const activeStudents = await Student_1.default.find({
            status: { $in: ["active", "repeat", "on_leave", "deferred"] },
        })
            .populate("admissionAcademicYear", "year")
            .populate("program", "institution durationYears")
            .lean();
        let discontinued = 0;
        for (const student of activeStudents) {
            const yearStr = student.admissionAcademicYear?.year ?? "";
            const admissionYear = parseInt(yearStr.split("/")[0]);
            if (!admissionYear || isNaN(admissionYear))
                continue;
            const institutionId = student.program?.institution?.toString() ??
                student.institution?.toString();
            if (!institutionId)
                continue;
            // ── THE CORRECT APPROACH ──────────────────────────────────────────────
            // Use Program.durationYears × settings.ruleSet.maxDurationMultiplier
            // NEVER infer from degree name ("Financial Engineering" is BSc but 4yr)
            const programDuration = student.program?.durationYears ?? 5;
            const settings = await (0, loadInstitutionSettings_1.loadInstitutionSettings)(institutionId).catch(() => null);
            const multiplier = settings?.rules.maxDurationMultiplier ?? 2.0;
            const maxYears = Math.round(programDuration * multiplier);
            // Subtract approved leave years
            const effectiveYears = nowYear - admissionYear - (student.totalTimeOutYears ?? 0);
            if (effectiveYears > maxYears) {
                const rule = `ENG.19(${programDuration > 4 ? "d" : "e"})`;
                await Student_1.default.findByIdAndUpdate(student._id, {
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
                console.log(`[Cron] ${rule} auto-discontinued: ${student.regNo} (${effectiveYears}y > ${maxYears}y)`);
            }
        }
        console.log(`[Cron] ENG.19 check complete — ${discontinued} students discontinued`);
    });
    console.log("[Cron] Compliance jobs registered");
};
exports.startComplianceCronJobs = startComplianceCronJobs;
