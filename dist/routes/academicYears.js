"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/routes/academicYears.ts
const express_1 = require("express");
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const auditLogger_1 = require("../lib/auditLogger");
const cache_1 = require("../utils/cache");
const router = (0, express_1.Router)();
// 1. CREATE
router.post("/", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { year, startDate, endDate } = req.body;
    const start = new Date(startDate);
    const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    const intake = monthNames[start.getMonth()];
    const academicYear = await AcademicYear_1.default.create({
        institution: req.user.institution,
        year,
        intakes: [intake],
        startDate: start,
        endDate: new Date(endDate),
        isCurrent: false,
    });
    await (0, auditLogger_1.logAudit)(req, {
        action: "academic_year_created",
        actor: req.user._id,
        details: {
            year,
            intake,
            startDate: start.toISOString(),
            endDate: new Date(endDate).toISOString(),
            institutionId: req.user.institution?.toString(),
        },
    });
    (0, cache_1.invalidateCache)(`academicYears:${req.user.institution}`);
    res.status(201).json(academicYear);
}));
// 2. UPDATE (PATCH)
router.patch("/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const before = await AcademicYear_1.default.findOne({
        _id: req.params.id,
        institution: req.user.institution,
    }).lean();
    if (!before) {
        await (0, auditLogger_1.logAudit)(req, {
            action: "academic_year_update_failed",
            actor: req.user._id,
            details: {
                academicYearId: req.params.id,
                reason: "Not found or institution mismatch",
                attemptedChanges: req.body,
            },
        });
        return res.status(404).json({ message: "Academic year not found" });
    }
    const { isCurrent } = req.body;
    if (isCurrent) {
        await AcademicYear_1.default.updateMany({ institution: req.user.institution }, { isCurrent: false });
    }
    const updated = await AcademicYear_1.default.findByIdAndUpdate(req.params.id, { $set: req.body }, { new: true });
    await (0, auditLogger_1.logAudit)(req, {
        action: "academic_year_updated",
        actor: req.user._id,
        details: {
            academicYearId: req.params.id,
            year: before.year,
            institutionId: req.user.institution?.toString(),
            before: { isCurrent: before.isCurrent, session: before.session, startDate: before.startDate, endDate: before.endDate },
            after: req.body, demotedOthers: !!isCurrent
        }
    });
    (0, cache_1.invalidateCache)(`academicYears:${req.user.institution}`);
    res.json(updated);
}));
// 3. GET ALL
router.get("/", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    // const years = await AcademicYear.find({ institution: req.user.institution }).sort({ startDate: -1 }).lean();
    const institutionId = req.user.institution;
    const years = await (0, cache_1.cached)(`academicYears:${institutionId}`, () => AcademicYear_1.default.find({ institution: institutionId }).sort({ startDate: -1 }).lean());
    await (0, auditLogger_1.logAudit)(req, { action: "academic_years_listed", actor: req.user._id, details: { count: years.length, institutionId: req.user.institution?.toString() } });
    res.json(years);
}));
// 4. DELETE
router.delete("/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin", "coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const year = await AcademicYear_1.default.findOne({ _id: req.params.id, institution: req.user.institution });
    if (!year) {
        await (0, auditLogger_1.logAudit)(req, { action: "academic_year_delete_failed", actor: req.user._id, details: { academicYearId: req.params.id, reason: "Not found or institution mismatch", institutionId: req.user.institution?.toString() } });
        return res.status(404).json({ message: "Year not found" });
    }
    if (year.isCurrent) {
        await (0, auditLogger_1.logAudit)(req, { action: "academic_year_delete_failed", actor: req.user._id, details: { academicYearId: req.params.id, year: year.year, reason: "Attempted deletion of active academic year", institutionId: req.user.institution?.toString() } });
        return res.status(400).json({ message: "Cannot delete the active academic year" });
    }
    await AcademicYear_1.default.findByIdAndDelete(req.params.id);
    await (0, auditLogger_1.logAudit)(req, { action: "academic_year_deleted", actor: req.user._id, details: { academicYearId: req.params.id, year: year.year, startDate: year.startDate, endDate: year.endDate, institutionId: req.user.institution?.toString() } });
    (0, cache_1.invalidateCache)(`academicYears:${req.user.institution}`);
    res.json({ message: "Deleted successfully" });
}));
exports.default = router;
