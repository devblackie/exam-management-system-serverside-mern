"use strict";
// src/routes/maintenance.ts
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const Student_1 = __importDefault(require("../models/Student"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const Unit_1 = __importDefault(require("../models/Unit"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const gradeCalculator_1 = require("../services/gradeCalculator");
const router = express_1.default.Router();
// ─── 1. BULK SOFT DELETE ──────────────────────────────────────────────────────
// Soft-deletes both Mark (detailed) and MarkDirect records matching the criteria,
// then removes the corresponding FinalGrade documents.
router.post("/bulk-cleanup", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { unitCode, programId, academicYear } = req.body;
    if (!academicYear)
        return res.status(400).json({ error: "Academic Year is required" });
    const yearDoc = await AcademicYear_1.default.findOne({ year: academicYear });
    const unitDoc = unitCode ? await Unit_1.default.findOne({ code: unitCode }) : null;
    if (!yearDoc)
        return res.status(404).json({ error: "Academic Year not found" });
    // Build shared query filter
    const query = { academicYear: yearDoc._id };
    if (programId) {
        const studentIds = await Student_1.default.find({ program: programId }).distinct("_id");
        query.student = { $in: studentIds };
    }
    if (unitDoc) {
        const pUnits = await ProgramUnit_1.default.find({ unit: unitDoc._id }).distinct("_id");
        query.programUnit = { $in: pUnits };
    }
    const [detailedResult, directResult, gradeResult] = await Promise.all([
        Mark_1.default.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() } }),
        MarkDirect_1.default.updateMany({ ...query, deletedAt: null }, { $set: { deletedAt: new Date() } }),
        FinalGrade_1.default.deleteMany(query),
    ]);
    const totalTrashed = detailedResult.modifiedCount + directResult.modifiedCount;
    console.log(`[bulk-cleanup] detailed=${detailedResult.modifiedCount}, direct=${directResult.modifiedCount}, grades=${gradeResult.deletedCount}`);
    res.json({
        count: totalTrashed,
        message: `Moved ${detailedResult.modifiedCount} detailed + ${directResult.modifiedCount} direct marks to trash. Removed ${gradeResult.deletedCount} grades.`,
        detail: { detailed: detailedResult.modifiedCount, direct: directResult.modifiedCount, grades: gradeResult.deletedCount },
    });
}));
// ─── 2. GET TRASHED MARKS ─────────────────────────────────────────────────────
// Returns trashed records from both Mark and MarkDirect, tagged with `source`.
router.get("/trash-bin", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const populateOpts = [
        { path: "student", select: "regNo name" },
        { path: "programUnit", populate: { path: "unit", select: "code name" } },
        { path: "academicYear", select: "year" },
    ];
    const [detailedTrashed, directTrashed] = await Promise.all([
        Mark_1.default.find({ deletedAt: { $ne: null } })
            .populate(populateOpts).sort({ deletedAt: -1 }).limit(100).lean(),
        MarkDirect_1.default.find({ deletedAt: { $ne: null } })
            .populate(populateOpts).sort({ deletedAt: -1 }).limit(100).lean(),
    ]);
    // Tag each record so the frontend knows which collection to target on restore/purge
    const combined = [
        ...detailedTrashed.map((m) => ({ ...m, source: "detailed" })),
        ...directTrashed.map((m) => ({ ...m, source: "direct" })),
    ].sort((a, b) => new Date(b.deletedAt).getTime() - new Date(a.deletedAt).getTime());
    res.json(combined);
}));
// ─── 3. RESTORE OR PERMANENT DELETE ──────────────────────────────────────────
// Each entry in markIds must carry a `source` field ("detailed" | "direct") so
// we can route the action to the correct collection.
// Format: markIds = [{ id: "...", source: "detailed" | "direct" }, ...]
// For backwards compatibility, plain string IDs are treated as "detailed".
router.post("/trash-action", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    let { markIds, action } = req.body;
    // console.log(`[DEBUG] Received action: ${action} for ${markIds.length} marks`);
    if (!Array.isArray(markIds))
        markIds = [markIds];
    // Normalise: support both plain string IDs (legacy) and { id, source } objects
    const normalised = markIds.map((entry) => {
        if (typeof entry === "string")
            return { id: entry, source: "detailed" };
        if (!entry.source) {
            console.warn(`[DEBUG] Missing source for record: ${entry.id || entry._id}`);
        }
        return { id: entry.id || entry._id, source: entry.source || "detailed" };
    });
    const detailedIds = normalised.filter((e) => e.source === "detailed").map((e) => e.id);
    const directIds = normalised.filter((e) => e.source === "direct").map((e) => e.id);
    // ── RESTORE ─────────────────────────────────────────────────────────────
    if (action === "restore") {
        // Restore detailed marks
        if (detailedIds.length > 0) {
            const marks = await Mark_1.default.find({
                _id: { $in: detailedIds },
                deletedAt: { $ne: null },
            });
            // console.log(`[DEBUG] Found ${marks.length} detailed marks to restore`);
            const updateRes = await Mark_1.default.updateMany({ _id: { $in: detailedIds } }, { $set: { deletedAt: null } });
            // console.log(`[DEBUG] Detailed restore update result:`, updateRes);
            for (const mark of marks) {
                try {
                    await (0, gradeCalculator_1.computeFinalGrade)({ markId: mark._id });
                }
                catch (e) {
                    console.warn(`[maintenance] FinalGrade recalc failed for detailed ${mark._id}: ${e.message}`);
                }
            }
        }
        // Restore direct marks
        if (directIds.length > 0) {
            const directMarks = await MarkDirect_1.default.find({ _id: { $in: directIds }, deletedAt: { $ne: null } });
            // console.log(`[DEBUG] Found ${directMarks.length} direct marks to restore`);
            const updateRes = await MarkDirect_1.default.updateMany({ _id: { $in: directIds } }, { $set: { deletedAt: null } });
            // console.log(`[DEBUG] Direct restore update result:`, updateRes);
            for (const mark of directMarks) {
                try {
                    await (0, gradeCalculator_1.computeFinalGrade)({ markId: mark._id });
                }
                catch (e) {
                    console.warn(`[maintenance] FinalGrade recalc failed for direct ${mark._id}: ${e.message}`);
                }
            }
        }
        // console.log(`[maintenance] Restored: ${detailedIds.length} detailed, ${directIds.length} direct`);
        return res.json({
            message: `Restored ${detailedIds.length} detailed and ${directIds.length} direct marks. Grades recalculated.`
        });
    }
    // ── PURGE ────────────────────────────────────────────────────────────────
    if (action === "purge") {
        // console.log(`[DEBUG] Purging IDs - Detailed: ${detailedIds.length}, Direct: ${directIds.length}`);
        const [detailedDel, directDel] = await Promise.all([
            detailedIds.length > 0 ? Mark_1.default.deleteMany({ _id: { $in: detailedIds } }) : Promise.resolve({ deletedCount: 0 }),
            directIds.length > 0 ? MarkDirect_1.default.deleteMany({ _id: { $in: directIds } }) : Promise.resolve({ deletedCount: 0 }),
        ]);
        // console.log(`[DEBUG] Purge results - Detailed: ${detailedDel.deletedCount}, Direct: ${directDel.deletedCount}`);
        // console.log(`[maintenance] Purged: ${detailedDel.deletedCount} detailed, ${directDel.deletedCount} direct`);
        return res.json({ message: `Permanently deleted ${detailedDel.deletedCount} detailed and ${directDel.deletedCount} direct marks.` });
    }
    res.status(400).json({ error: "Invalid action. Use 'restore' or 'purge'." });
}));
exports.default = router;
