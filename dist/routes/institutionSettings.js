"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/routes/institutionSettings.ts — COMPLETE, FINAL VERSION
const express_1 = __importDefault(require("express"));
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const auth_1 = require("../middleware/auth");
const asyncHandler_1 = require("../middleware/asyncHandler");
const auditLogger_1 = require("../lib/auditLogger");
const cache_1 = require("../utils/cache");
const loadInstitutionSettings_1 = require("../utils/loadInstitutionSettings");
const multer_1 = __importDefault(require("multer"));
const path_1 = __importDefault(require("path"));
const fs_1 = __importDefault(require("fs"));
const router = express_1.default.Router();
// ── Logo upload (admin only) ──────────────────────────────────────────────────
const logoStorage = multer_1.default.diskStorage({
    destination: (_req, _file, cb) => {
        const dir = path_1.default.join(process.cwd(), "uploads", "logos");
        fs_1.default.mkdirSync(dir, { recursive: true });
        cb(null, dir);
    },
    filename: (_req, file, cb) => {
        const ext = path_1.default.extname(file.originalname).toLowerCase();
        cb(null, `logo-${Date.now()}${ext}`);
    },
});
const uploadLogo = (0, multer_1.default)({
    storage: logoStorage,
    limits: { fileSize: 2 * 1024 * 1024 },
    fileFilter: (_req, file, cb) => {
        const allowed = [".png", ".jpg", ".jpeg", ".svg"];
        if (allowed.includes(path_1.default.extname(file.originalname).toLowerCase())) {
            cb(null, true);
        }
        else {
            cb(new Error("Only PNG, JPG, JPEG, SVG files allowed"));
        }
    },
});
// ── GET /institution-settings — ALL authenticated users (read) ────────────────
router.get("/", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
        res.status(400).json({ message: "No institution linked to this account" });
        return;
    }
    const settings = await (0, cache_1.cached)(`settings:${institutionId}`, () => InstitutionSettings_1.default.findOne({ institution: institutionId }).lean(), 300);
    await (0, auditLogger_1.logAudit)(req, {
        action: "institution_settings_viewed",
        details: { institutionId },
    });
    res.json(settings ?? null);
}));
// ── POST /institution-settings — COORDINATOR + ADMIN ─────────────────────────
// Coordinator manages: ruleSet, gradingScale, waaClassification,
//                      semesterWeights, enforceRegNoPattern, supportedIntakes
// Admin manages:       docMeta, schools, departments (via separate PATCH routes)
//
// Both roles can POST to this endpoint, but each saves different fields.
// The server accepts the full payload and applies whichever fields are present.
// Field-level authorization is enforced inside the handler.
router.post("/", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
        res.status(400).json({ message: "No institution linked to this account" });
        return;
    }
    const data = req.body;
    const role = req.user.role;
    // ── Field-level authorization ─────────────────────────────────────────────
    // Coordinator can ONLY save academic/grading fields
    // Admin can ONLY save identity/structure fields via this endpoint
    // (admin uses PATCH /schools and PATCH /schools/:code/departments for structure)
    const COORDINATOR_FIELDS = new Set([
        "ruleSet",
        "gradingScale",
        "waaClassification",
        "semesterWeights",
        "enforceRegNoPattern",
        "supportedIntakes",
    ]);
    const ADMIN_FIELDS = new Set([
        "docMeta",
        // schools/departments are managed via PATCH endpoints
    ]);
    const allowedFields = role === "admin" ? ADMIN_FIELDS : COORDINATOR_FIELDS;
    // Strip fields the caller is not allowed to set
    const filteredData = {};
    for (const key of Object.keys(data)) {
        if (allowedFields.has(key)) {
            filteredData[key] = data[key];
        }
    }
    if (Object.keys(filteredData).length === 0) {
        res.status(422).json({
            message: `No writable fields provided for role "${role}". ` +
                (role === "coordinator"
                    ? "Coordinators can set: ruleSet, gradingScale, waaClassification, semesterWeights, enforceRegNoPattern, supportedIntakes"
                    : "Admins can set: docMeta (schools via PATCH /schools)"),
        });
        return;
    }
    // Validate CA + Exam weights if ruleSet provided
    if (filteredData.ruleSet) {
        const rs = filteredData.ruleSet;
        if (rs.caWeight !== undefined && rs.examWeight !== undefined) {
            const total = Number(rs.caWeight) + Number(rs.examWeight);
            if (Math.abs(total - 100) > 0.01) {
                res.status(422).json({
                    message: `CA weight (${rs.caWeight}) + Exam weight (${rs.examWeight}) must equal 100`,
                });
                return;
            }
        }
    }
    const previous = await InstitutionSettings_1.default.findOne({ institution: institutionId }).lean();
    const updated = await InstitutionSettings_1.default.findOneAndUpdate({ institution: institutionId }, { $set: { ...filteredData, institution: institutionId } }, { new: true, upsert: true, setDefaultsOnInsert: true, runValidators: true });
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    await (0, auditLogger_1.logAudit)(req, {
        action: previous ? "institution_settings_updated" : "institution_settings_created",
        details: { institutionId, role, fields: Object.keys(filteredData) },
    });
    res.json({ message: "Settings saved successfully", settings: updated });
}));
// ── POST /institution-settings/logo — ADMIN ONLY ──────────────────────────────
router.post("/logo", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), uploadLogo.single("logo"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    if (!req.file) {
        throw { statusCode: 400, message: "No file uploaded" };
    }
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
        throw { statusCode: 400, message: "No institution linked" };
    }
    // Delete old logo file
    const existing = await InstitutionSettings_1.default.findOne({ institution: institutionId })
        .select("branding.universityLogoPath")
        .lean();
    if (existing?.branding?.universityLogoPath) {
        const oldPath = path_1.default.join(process.cwd(), existing.branding.universityLogoPath);
        if (fs_1.default.existsSync(oldPath))
            fs_1.default.unlinkSync(oldPath);
    }
    const logoPath = `uploads/logos/${req.file.filename}`;
    await InstitutionSettings_1.default.findOneAndUpdate({ institution: institutionId }, { $set: { "branding.universityLogoPath": logoPath } }, { upsert: true });
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    await (0, auditLogger_1.logAudit)(req, {
        action: "institution_logo_uploaded",
        details: { institutionId, logoPath },
    });
    res.json({ message: "Logo uploaded successfully", path: logoPath });
}));
// ── GET /institution-settings/logo — serve logo (all authenticated users) ─────
router.get("/logo", auth_1.requireAuth, (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
        res.status(400).json({ message: "No institution linked" });
        return;
    }
    const settings = await InstitutionSettings_1.default.findOne({ institution: institutionId })
        .select("branding.universityLogoPath")
        .lean();
    const logoPath = settings?.branding?.universityLogoPath;
    if (!logoPath) {
        res.status(404).json({ message: "No logo uploaded" });
        return;
    }
    const fullPath = path_1.default.join(process.cwd(), logoPath);
    if (!fs_1.default.existsSync(fullPath)) {
        res.status(404).json({ message: "Logo file not found" });
        return;
    }
    res.sendFile(fullPath);
}));
// ── PATCH /institution-settings/schools — ADMIN ONLY ─────────────────────────
router.patch("/schools", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { school } = req.body;
    const institutionId = req.user.institution?.toString();
    if (!school?.code || !school?.name) {
        throw { statusCode: 400, message: "School code and name are required" };
    }
    // Uniqueness: same name, different code = clash
    const nameClash = await InstitutionSettings_1.default.findOne({
        institution: institutionId,
        "schools.name": { $regex: new RegExp(`^${String(school.name).trim()}$`, "i") },
        "schools.code": { $ne: String(school.code).toUpperCase() },
    });
    if (nameClash) {
        throw {
            statusCode: 409,
            message: `A school named "${school.name}" already exists.`,
        };
    }
    const existing = await InstitutionSettings_1.default.findOne({
        institution: institutionId,
        "schools.code": String(school.code).toUpperCase(),
    });
    const schoolDoc = { ...school, code: String(school.code).toUpperCase() };
    if (existing) {
        await InstitutionSettings_1.default.updateOne({ institution: institutionId, "schools.code": String(school.code).toUpperCase() }, { $set: { "schools.$": schoolDoc } });
    }
    else {
        await InstitutionSettings_1.default.updateOne({ institution: institutionId }, { $push: { schools: schoolDoc } }, { upsert: true });
    }
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    res.json({ message: `School ${school.code} saved` });
}));
// ── PATCH /institution-settings/schools/:schoolCode/departments — ADMIN ONLY ──
// Reg number patterns inside departments are COORDINATOR responsibility.
// Admin creates the department shell; coordinator fills in regNoPatterns.
router.patch("/schools/:schoolCode/departments", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { schoolCode } = req.params;
    const { department } = req.body;
    const institutionId = req.user.institution?.toString();
    if (!department?.code || !department?.name) {
        throw { statusCode: 400, message: "Department code and name are required" };
    }
    const nameClash = await InstitutionSettings_1.default.findOne({
        institution: institutionId,
        "schools.code": schoolCode.toUpperCase(),
        "schools.departments.name": {
            $regex: new RegExp(`^${String(department.name).trim()}$`, "i"),
        },
        "schools.departments.code": { $ne: String(department.code).toUpperCase() },
    });
    if (nameClash) {
        throw {
            statusCode: 409,
            message: `A department named "${department.name}" already exists in this school.`,
        };
    }
    const deptDoc = { ...department, code: String(department.code).toUpperCase() };
    const result = await InstitutionSettings_1.default.updateOne({
        institution: institutionId,
        "schools.code": schoolCode.toUpperCase(),
        "schools.departments.code": { $ne: String(department.code).toUpperCase() },
    }, { $push: { "schools.$.departments": deptDoc } });
    if (result.matchedCount === 0) {
        await InstitutionSettings_1.default.updateOne({
            institution: institutionId,
            "schools.code": schoolCode.toUpperCase(),
            "schools.departments.code": String(department.code).toUpperCase(),
        }, { $set: { "schools.$[s].departments.$[d]": deptDoc } }, {
            arrayFilters: [
                { "s.code": schoolCode.toUpperCase() },
                { "d.code": String(department.code).toUpperCase() },
            ],
        });
    }
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    res.json({ message: `Department ${department.code} saved in ${schoolCode}` });
}));
// ── PATCH /institution-settings/schools/:schoolCode/departments/:deptCode/reg-patterns
// COORDINATOR ONLY — set/update reg number patterns for a department they belong to
// router.patch(
//   "/schools/:schoolCode/departments/:deptCode/reg-patterns",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode, deptCode } = req.params;
//     const { regNoPatterns }        = req.body as {
//       regNoPatterns: Array<{
//         prefix: string; separator: string; yearDigits: number;
//         example: string; manualRegex?: string;
//       }>;
//     };
//     const institutionId = req.user.institution?.toString();
//     // Coordinators can only update patterns for their own department
//     if (req.user.role === "coordinator") {
//       const myDept  = req.user.departmentCode;
//       const mySchool= req.user.schoolCode;
//       if (
//         !req.user.institutionWide &&
//         (myDept !== deptCode.toUpperCase() || mySchool !== schoolCode.toUpperCase())
//       ) {
//         throw {
//           statusCode: 403,
//           message:    "You can only configure reg number patterns for your own department.",
//         } as ApiError;
//       }
//     }
//     if (!Array.isArray(regNoPatterns)) {
//       throw { statusCode: 400, message: "regNoPatterns must be an array" } as ApiError;
//     }
//     await InstitutionSettings.updateOne(
//       {
//         institution:    institutionId,
//         "schools.code": schoolCode.toUpperCase(),
//         "schools.departments.code": deptCode.toUpperCase(),
//       },
//       {
//         $set: {
//           "schools.$[s].departments.$[d].regNoPatterns": regNoPatterns,
//           "schools.$[s].departments.$[d].enforceRegNoPattern": true,
//         },
//       },
//       {
//         arrayFilters: [
//           { "s.code": schoolCode.toUpperCase() },
//           { "d.code": deptCode.toUpperCase() },
//         ],
//       },
//     );
//     // Also update the top-level enforceRegNoPattern flag
//     await InstitutionSettings.updateOne(
//       { institution: institutionId },
//       { $set: { enforceRegNoPattern: true } },
//     );
//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     await logAudit(req, {
//       action:  "reg_no_patterns_updated",
//       details: { institutionId, schoolCode, deptCode, patternCount: regNoPatterns.length },
//     });
//     res.json({ message: `Reg number patterns updated for ${deptCode}` });
//   }),
// );
// serverside/src/routes/institutionSettings.ts - Add this endpoint with ownership check
// ── PATCH /institution-settings/schools/:schoolCode/departments/:deptCode/reg-patterns
// COORDINATOR ONLY — can only edit their own department
router.patch("/schools/:schoolCode/departments/:deptCode/reg-patterns", auth_1.requireAuth, (0, auth_1.requireRole)("coordinator", "admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { schoolCode, deptCode } = req.params;
    const { regNoPatterns } = req.body;
    const institutionId = req.user.institution?.toString();
    // ENFORCE: coordinators can only edit their OWN department
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
        if (req.user.schoolCode?.toUpperCase() !== schoolCode.toUpperCase() ||
            req.user.departmentCode?.toUpperCase() !== deptCode.toUpperCase()) {
            throw {
                statusCode: 403,
                message: "You can only configure registration number patterns for your own department.",
            };
        }
    }
    if (!Array.isArray(regNoPatterns)) {
        throw { statusCode: 400, message: "regNoPatterns must be an array" };
    }
    await InstitutionSettings_1.default.updateOne({
        institution: institutionId,
        "schools.code": schoolCode.toUpperCase(),
        "schools.departments.code": deptCode.toUpperCase(),
    }, {
        $set: {
            "schools.$[s].departments.$[d].regNoPatterns": regNoPatterns,
            "schools.$[s].departments.$[d].enforceRegNoPattern": true,
        },
    }, {
        arrayFilters: [
            { "s.code": schoolCode.toUpperCase() },
            { "d.code": deptCode.toUpperCase() },
        ],
    });
    // Also update the top-level enforceRegNoPattern flag
    await InstitutionSettings_1.default.updateOne({ institution: institutionId }, { $set: { enforceRegNoPattern: true } });
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    await (0, auditLogger_1.logAudit)(req, {
        action: "reg_no_patterns_updated",
        details: { institutionId, schoolCode, deptCode, patternCount: regNoPatterns.length },
    });
    res.json({ message: `Registration number patterns updated for ${deptCode}` });
}));
// ── DELETE /institution-settings/schools/:schoolCode — ADMIN ONLY ─────────────
router.delete("/schools/:schoolCode", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { schoolCode } = req.params;
    const institutionId = req.user.institution?.toString();
    await InstitutionSettings_1.default.updateOne({ institution: institutionId }, { $pull: { schools: { code: schoolCode.toUpperCase() } } });
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    res.json({ message: `School ${schoolCode} removed` });
}));
// ── DELETE /institution-settings/schools/:schoolCode/departments/:deptCode ────
router.delete("/schools/:schoolCode/departments/:deptCode", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { schoolCode, deptCode } = req.params;
    const institutionId = req.user.institution?.toString();
    await InstitutionSettings_1.default.updateOne({ institution: institutionId, "schools.code": schoolCode.toUpperCase() }, { $pull: { "schools.$.departments": { code: deptCode.toUpperCase() } } });
    (0, cache_1.invalidateCache)(`settings:${institutionId}`);
    (0, loadInstitutionSettings_1.invalidateSettingsCache)(institutionId);
    res.json({ message: `Department ${deptCode} removed from ${schoolCode}` });
}));
exports.default = router;
