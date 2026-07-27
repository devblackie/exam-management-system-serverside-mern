"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// src/routes/coordinator.ts
const express_1 = require("express");
const bcryptjs_1 = __importDefault(require("bcryptjs"));
const User_1 = __importDefault(require("../models/User"));
const Institution_1 = __importDefault(require("../models/Institution"));
const auditLogger_1 = require("../lib/auditLogger");
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const Student_1 = __importDefault(require("../models/Student"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const cleanupGrades_1 = require("../scripts/cleanupGrades");
const mongoose_1 = __importDefault(require("mongoose"));
const Program_1 = __importDefault(require("../models/Program"));
const Mark_1 = __importDefault(require("../models/Mark"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const DisciplinaryCase_1 = __importDefault(require("../models/DisciplinaryCase"));
const router = (0, express_1.Router)();
// Coordinator secret registration
router.post("/secret-register", (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { secret, name, email, password, institutionId } = req.body;
    // Validate Secret Key
    if (secret !== process.env.COORDINATOR_SECRET) {
        await (0, auditLogger_1.logAudit)(req, { action: "coordinator_register_failed_invalid_secret" });
        return res.status(403).json({ message: "Invalid secret" });
    }
    // Check duplicate email
    const existing = await User_1.default.findOne({ email });
    if (existing) {
        await (0, auditLogger_1.logAudit)(req, { action: "coordinator_register_failed_duplicate", details: { email } });
        return res.status(400).json({ message: "Email already in use" });
    }
    // Validate institution exists (if provided)
    let institution = null;
    if (institutionId) {
        // console.log("Received institutionId:", institutionId);
        // institution = await Institution.findById(institutionId);
        // Validate it's a proper ObjectId string
        if (!institutionId.match(/^[0-9a-fA-F]{24}$/)) {
            // console.log("Invalid ObjectId format:", institutionId);
            return res.status(400).json({ message: "Invalid institution ID format" });
        }
        institution = await Institution_1.default.findById(institutionId);
        // console.log("Found institution:", institution); // ← DEBUG LOG
        if (!institution) {
            // console.log("Institution not found in DB for ID:", institutionId);
            return res.status(400).json({ message: "Invalid institution ID" });
        }
        if (!institution.isActive) {
            return res.status(400).json({ message: "Institution is not active" });
        }
        // if (!institution) return res.status(400).json({ message: "Invalid institution" });
    }
    // Create coordinator
    const hashed = await bcryptjs_1.default.hash(password, 12);
    const coordinator = await User_1.default.create({
        name,
        email: email.toLowerCase(),
        password: hashed,
        role: "coordinator",
        institution: institution?._id || null, // optional: assign institution
    });
    await (0, auditLogger_1.logAudit)(req, {
        action: "coordinator_registered",
        targetUser: coordinator._id,
        details: { email, name, institution: institutionId },
    });
    res.status(201).json({ message: "Coordinator created successfully" });
}));
router.post("/maintain/cleanup-grades", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    try {
        console.log("Admin initiated manual database cleanup...");
        await (0, cleanupGrades_1.cleanupOrphanedGrades)();
        res.json({
            success: true,
            message: "Data integrity restored. Orphaned grades have been purged."
        });
    }
    catch (error) {
        console.error("Cleanup Route Error:", error);
        res.status(500).json({
            success: false,
            error: "Failed to perform database maintenance."
        });
    }
}));
// ── GET /coordinator/dashboard-stats ─────────────────────────────────────────
router.get("/dashboard-stats", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institutionId = req.user.institution;
    const allowedProgramIds = await (0, auth_1.getScopedProgramIds)(req);
    // Resolve student IDs in scope first — needed for disciplinary sub-query
    const scopedStudentIds = await Student_1.default.find({
        institution: institutionId,
        program: { $in: allowedProgramIds },
    })
        .select("_id")
        .lean()
        .then(ss => ss.map(s => s._id));
    // ── All queries in parallel ───────────────────────────────────────────────
    const [studentCounts, programs, currentYear, openCases, markBatchIds, directMarkBatchIds, lastMark, lastDirectMark,] = await Promise.all([
        // Student status breakdown
        Student_1.default.aggregate([
            {
                $match: {
                    institution: new mongoose_1.default.Types.ObjectId(institutionId.toString()),
                    program: { $in: allowedProgramIds },
                },
            },
            { $group: { _id: "$status", count: { $sum: 1 } } },
        ]),
        // Scoped program names
        Program_1.default.find({
            institution: institutionId,
            _id: { $in: allowedProgramIds },
            isActive: true,
        })
            .select("name")
            .lean(),
        // Current academic year
        AcademicYear_1.default.findOne({ institution: institutionId, isCurrent: true })
            .select("year session")
            .lean(),
        // Open disciplinary cases for students in scope
        DisciplinaryCase_1.default.countDocuments({
            institution: institutionId,
            outcome: "PENDING",
            student: { $in: scopedStudentIds },
        }),
        // Distinct mark batch IDs
        Mark_1.default.distinct("batchId", {
            institution: institutionId,
            program: { $in: allowedProgramIds },
        }),
        // Distinct direct mark batch IDs
        MarkDirect_1.default.distinct("batchId", {
            institution: institutionId,
            program: { $in: allowedProgramIds },
        }),
        // Most recent detailed mark upload
        Mark_1.default.findOne({
            institution: institutionId,
            program: { $in: allowedProgramIds },
        })
            .sort({ uploadedAt: -1 })
            .select("uploadedAt")
            .lean(),
        // Most recent direct mark upload
        MarkDirect_1.default.findOne({
            institution: institutionId,
            program: { $in: allowedProgramIds },
        })
            .sort({ uploadedAt: -1 })
            .select("uploadedAt")
            .lean(),
    ]);
    // ── Aggregate student statuses ────────────────────────────────────────────
    const sm = {};
    for (const row of studentCounts)
        sm[row._id] = row.count;
    const total = Object.values(sm).reduce((a, b) => a + b, 0);
    // ── Resolve last upload date ──────────────────────────────────────────────
    const d1 = lastMark?.uploadedAt ? new Date(lastMark.uploadedAt).getTime() : 0;
    const d2 = lastDirectMark?.uploadedAt ? new Date(lastDirectMark.uploadedAt).getTime() : 0;
    const lastUploadDate = d1 === 0 && d2 === 0
        ? null
        : new Date(Math.max(d1, d2)).toISOString();
    res.json({
        students: {
            total,
            active: sm["active"] ?? 0,
            repeat: sm["repeat"] ?? 0,
            discontinued: sm["discontinued"] ?? 0,
            graduated: sm["graduated"] ?? 0,
            suspended: sm["disciplinary_suspension"] ?? 0,
        },
        marks: {
            totalUploads: markBatchIds.length + directMarkBatchIds.length,
            pendingReview: 0,
            lastUploadDate,
        },
        disciplinary: {
            openCases: openCases,
            pendingOutcome: openCases,
        },
        programs: {
            total: programs.length,
            names: programs.map(p => p.name),
        },
        promotion: {
            lastRunDate: null,
            eligibleCount: 0,
        },
        academicYear: {
            current: currentYear?.year ?? null,
            session: currentYear?.session ?? null,
        },
    });
}));
// ── GET /coordinator/lecturers ────────────────────────────────────────────────
// Coordinators can list lecturers in their department
router.get("/lecturers", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const User = (await Promise.resolve().then(() => __importStar(require("../models/User")))).default;
    const lecturers = await User.find({
        institution: req.user.institution,
        role: "lecturer",
        departmentCode: req.user.departmentCode,
    })
        .select("name email departmentCode schoolCode createdAt")
        .lean();
    res.json(lecturers);
}));
// ── POST /coordinator/lecturers ───────────────────────────────────────────────
router.post("/lecturers", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { name, email, password } = req.body;
    if (!name?.trim() || !email?.trim()) {
        res.status(400).json({ message: "name and email are required." });
        return;
    }
    const User = (await Promise.resolve().then(() => __importStar(require("../models/User")))).default;
    const bcrypt = await Promise.resolve().then(() => __importStar(require("bcryptjs")));
    const existing = await User.findOne({ email: email.toLowerCase() });
    if (existing) {
        res.status(409).json({ message: "A user with this email already exists." });
        return;
    }
    const hash = await bcrypt.hash(password?.trim() || Math.random().toString(36).slice(-10), 10);
    const lecturer = await User.create({
        name: name.trim(),
        email: email.toLowerCase().trim(),
        password: hash,
        role: "lecturer",
        institution: req.user.institution,
        schoolCode: req.user.schoolCode,
        departmentCode: req.user.departmentCode,
        isVerified: true,
    });
    res.status(201).json({
        message: "Lecturer created.",
        lecturer: { _id: lecturer._id, name: lecturer.name, email: lecturer.email },
    });
}));
// Coordinator creates lecturer (no login needed)
router.post("/lecturers", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { name, email } = req.body;
    if (!name || !email) {
        return res.status(400).json({ message: "Name and email required" });
    }
    const existing = await User_1.default.findOne({ email: email.toLowerCase() });
    if (existing) {
        return res.status(400).json({ message: "Email already exists" });
    }
    const password = Math.random().toString(36).slice(-8) + "A1!";
    const hashed = await bcryptjs_1.default.hash(password, 12);
    const lecturer = await User_1.default.create({
        name,
        email: email.toLowerCase(),
        password: hashed,
        role: "lecturer",
        institution: req.user.institution,
        status: "active",
    });
    await (0, auditLogger_1.logAudit)(req, {
        action: "lecturer_created",
        actor: req.user._id,
        targetUser: lecturer._id,
        details: { email, name },
    });
    res.status(201).json({
        message: "Lecturer created",
        email,
        temporaryPassword: password,
    });
}));
// View student results (senate-style)
router.get("/students/:regNo/results", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { regNo } = req.params;
    const student = await Student_1.default.findOne({
        regNo: regNo.toUpperCase(),
        institution: req.user.institution,
    });
    if (!student) {
        return res.status(404).json({ message: "Student not found" });
    }
    const grades = await FinalGrade_1.default.find({ student: student._id })
        .populate({
        path: "programUnit",
        // We need the unit (template) and the required year/semester from the link
        select: "requiredYear requiredSemester unit",
        populate: { path: "unit", model: "Unit", select: "code name" }
    })
        .populate({ path: "academicYear", select: "year" })
        .sort({ "academicYear.year": 1, "programUnit.requiredYear": 1, "programUnit.requiredSemester": 1 }) // ⬅️ Adjusted sort fields
        .lean();
    await (0, auditLogger_1.logAudit)(req, {
        action: "coordinator_viewed_student_results",
        actor: req.user._id,
        // targetUser: student._id,
        details: { regNo },
    });
    res.json({
        student: { name: student.name, regNo: student.regNo, program: student.program },
        results: grades.map(g => {
            const grade = g;
            return {
                // ⬅️ FIX 2: Access unit and scheduling details through programUnit
                unitCode: grade.programUnit.unit.code,
                unitName: grade.programUnit.unit.name,
                year: grade.programUnit.requiredYear,
                semester: grade.programUnit.requiredSemester,
                academicYear: grade.academicYear.year,
                totalMark: grade.totalMark,
                grade: grade.grade,
                status: grade.status,
                capped: grade.cappedBecauseSupplementary,
            };
        }),
    });
}));
exports.default = router;
