
// // src/routes/institutionSettings.ts
// import express, { Response } from "express";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { logAudit } from "../lib/auditLogger";
// import Mark from "../models/Mark";
// import { cached, invalidateCache } from "../utils/cache";

// const router = express.Router();

// // GET: Fetch current settings
// router.get("/", requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     // const settings = await InstitutionSettings.findOne({institution: req.user.institution });

//     const institutionId = req.user.institution;
//     const settings = await cached(`settings:${institutionId}`, () => 
//       InstitutionSettings.findOne({ institution: institutionId }).lean()
//     );

//     if (!settings) {
//       await logAudit(req, { action: "institution_settings_fetch_failed", actor: req.user._id, details: { reason: "Settings not yet configured", institutionId: req.user.institution?.toString()}});
//       return res.status(404).json({ message: "Settings not configured yet" });
//     }

//     await logAudit(req, { action: "institution_settings_viewed", actor: req.user._id, details: { institutionId: req.user.institution?.toString(), passMark: settings.passMark, gradingScaleCount: settings.gradingScale?.length ?? 0 }});
//     res.json(settings);
//   })
// );

// // POST: Save / Update settings
// router.post(
//   "/",
//   requireAuth,
//   requireRole("coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const data = req.body;

//     const previous = await InstitutionSettings.findOne({
//       institution: req.user.institution,
//     }).lean();

//     const existingMarksCount = await Mark.countDocuments({
//       institution: req.user.institution,
//     });

//     const updated = await InstitutionSettings.findOneAndUpdate(
//       { institution: req.user.institution },
//       data,
//       {
//         new: true,
//         upsert: true,
//         setDefaultsOnInsert: true,
//         runValidators: true,
//       }
//     );

//     await logAudit(req, {
//       action: previous
//         ? "institution_settings_updated"
//         : "institution_settings_created",
//       actor: req.user._id,
//       details: {
//         settingsId: updated?._id?.toString(),
//         institutionId: req.user.institution?.toString(),
//         existingMarksAtTimeOfChange: existingMarksCount,
//         previous: previous
//           ? {
//               passMark: previous.passMark,
//               gradingScaleCount: previous.gradingScale?.length ?? 0,
//             }
//           : null,
//         updated: {
//           passMark: data.passMark,
//           gradingScaleCount: data.gradingScale?.length ?? 0,
//         },
//         fullPayload: data,
//       },
//     });
//     invalidateCache(`settings:${req.user.institution}`);
//     res.json({
//       message: "Institution settings saved successfully",
//       settings: updated,
//     });
//   })
// );

// export default router;








// // serverside/src/routes/institutionSettings.ts — COMPLETE FIXED VERSION
// import express, { Response } from "express";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { logAudit } from "../lib/auditLogger";
// import { cached, invalidateCache } from "../utils/cache";
// import { invalidateSettingsCache } from "../utils/loadInstitutionSettings";
// import multer from "multer";
// import path   from "path";
// import fs     from "fs";
// import Program from "../models/Program";

// // ── ApiError type (local — matches errorHandler.ts shape) ────────────────────
// interface ApiError { statusCode: number; message: string }

// const router = express.Router();

// // ── Logo upload middleware ────────────────────────────────────────────────────
// const logoStorage = multer.diskStorage({
//   destination: (_req, _file, cb) => {
//     const dir = path.join(process.cwd(), "uploads", "logos");
//     fs.mkdirSync(dir, { recursive: true });
//     cb(null, dir);
//   },
//   filename: (_req, file, cb) => {
//     const ext = path.extname(file.originalname).toLowerCase();
//     cb(null, `logo-${Date.now()}${ext}`);
//   },
// });

// const uploadLogo = multer({
//   storage:  logoStorage,
//   limits:   { fileSize: 2 * 1024 * 1024 },
//   fileFilter: (_req, file, cb) => {
//     const allowed = [".png", ".jpg", ".jpeg", ".svg"];
//     const ext     = path.extname(file.originalname).toLowerCase();
//     if (allowed.includes(ext)) {
//       cb(null, true);
//     } else {
//       cb(new Error("Only PNG, JPG, JPEG, SVG files allowed"));
//     }
//   },
// });

// // ── GET /institution-settings ─────────────────────────────────────────────────
// router.get(
//   "/",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await cached(
//       `settings:${institutionId}`,
//       () => InstitutionSettings.findOne({ institution: institutionId }).lean(),
//       300,
//     );

//     if (!settings) {
//       res.status(404).json({ message: "Settings not configured yet" });
//       return;
//     }

//     await logAudit(req, {
//       action: "institution_settings_viewed",
//       actor:  req.user._id,
//       details: { institutionId },
//     });

//     res.json(settings);
//   }),
// );

// // ── POST /institution-settings ────────────────────────────────────────────────
// // Coordinator or admin can save settings
// router.post(
//   "/",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const data = req.body;

//     // Validate: caWeight + examWeight must equal 100 if both provided
//     if (
//       data.ruleSet?.caWeight !== undefined &&
//       data.ruleSet?.examWeight !== undefined
//     ) {
//       const total = Number(data.ruleSet.caWeight) + Number(data.ruleSet.examWeight);
//       if (Math.abs(total - 100) > 0.01) {
//         res.status(422).json({
//           message: `CA weight (${data.ruleSet.caWeight}) + Exam weight (${data.ruleSet.examWeight}) must equal 100`,
//         });
//         return;
//       }
//     }

//     const previous = await InstitutionSettings.findOne({
//       institution: institutionId,
//     }).lean();

//     const updated = await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { ...data, institution: institutionId } },
//       { new: true, upsert: true, setDefaultsOnInsert: true, runValidators: true },
//     );

//     // Bust ALL settings cache entries for this institution
//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action:  previous ? "institution_settings_updated" : "institution_settings_created",
//       actor:   req.user._id,
//       details: { institutionId, settingsId: updated?._id?.toString() },
//     });

//     res.json({ message: "Settings saved successfully", settings: updated });
//   }),
// );

// // ── POST /institution-settings/logo ──────────────────────────────────────────
// // ONE university logo — used in ALL senate documents and CMS exports
// router.post(
//   "/logo",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   uploadLogo.single("logo"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     if (!req.file) {
//       throw { statusCode: 400, message: "No file uploaded" } as ApiError;
//     }

//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       throw { statusCode: 400, message: "No institution linked to this account" } as ApiError;
//     }

//     // Delete previous logo file if it exists
//     const existing = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean() as any;

//     if (existing?.branding?.universityLogoPath) {
//       const oldPath = path.join(process.cwd(), existing.branding.universityLogoPath);
//       if (fs.existsSync(oldPath)) {
//         fs.unlinkSync(oldPath);
//       }
//     }

//     const logoPath = `uploads/logos/${req.file.filename}`;

//     await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { "branding.universityLogoPath": logoPath } },
//       { upsert: true },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action:  "institution_logo_uploaded",
//       actor:   req.user._id,
//       details: { institutionId, logoPath },
//     });

//     res.json({ message: "Logo uploaded successfully", path: logoPath });
//   }),
// );

// // ── GET /institution-settings/logo — serve logo for frontend preview ──────────
// router.get(
//   "/logo",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked" });
//       return;
//     }

//     const settings = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean() as any;

//     const logoPath = settings?.branding?.universityLogoPath;

//     if (!logoPath) {
//       res.status(404).json({ message: "No logo uploaded" });
//       return;
//     }

//     const fullPath = path.join(process.cwd(), logoPath);
//     if (!fs.existsSync(fullPath)) {
//       res.status(404).json({ message: "Logo file not found" });
//       return;
//     }

//     res.sendFile(fullPath);
//   }),
// );

// // ── PATCH /institution-settings/schools — add/update a school ─────────────────
// router.patch(
//   "/schools",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { school } = req.body;
//     const institutionId = req.user.institution?.toString();

//     if (!school?.code || !school?.name) {
//       throw { statusCode: 400, message: "School code and name are required" } as ApiError;
//     }

//     const nameClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.name": { $regex: new RegExp(`^${school.name.trim()}$`, "i") },
//       "schools.code": { $ne: school.code }, // different code, same name = clash
//     });
//     if (nameClash) {
//       throw {
//         statusCode: 409,
//         message: `A school named "${school.name}" already exists.`,
//       } as ApiError;
//     }

//     // Upsert school by code
//     const existing = await InstitutionSettings.findOne({
//       institution:  institutionId,
//       "schools.code": school.code,
//     });

//     if (existing) {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId, "schools.code": school.code },
//         { $set: { "schools.$": school } },
//       );
//     } else {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId },
//         { $push: { schools: school } },
//         { upsert: true },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `School ${school.code} saved` });
//   }),
// );

// // ── PATCH /institution-settings/schools/:schoolCode/departments ───────────────
// router.patch(
//   "/schools/:schoolCode/departments",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const { department } = req.body;
//     const institutionId  = req.user.institution?.toString();

//     if (!department?.code || !department?.name) {
//       throw { statusCode: 400, message: "Department code and name are required" } as ApiError;
//     }

//     const nameClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.code": schoolCode,
//       "schools.departments.name": { $regex: new RegExp(`^${department.name.trim()}$`, "i") },
//       "schools.departments.code": { $ne: department.code },
//     });
//     if (nameClash) {
//       throw { statusCode: 409, message: `A department named "${department.name}" already exists in this school.` } as ApiError;
//     }

//     const result = await InstitutionSettings.updateOne(
//       {
//         institution:    institutionId,
//         "schools.code": schoolCode,
//         "schools.departments.code": { $ne: department.code },
//       },
//       { $push: { "schools.$.departments": department } },
//     );

//     if (result.matchedCount === 0) {
//       // Department already exists — update it
//       await InstitutionSettings.updateOne(
//         {
//           institution:    institutionId,
//           "schools.code": schoolCode,
//           "schools.departments.code": department.code,
//         },
//         { $set: { "schools.$[s].departments.$[d]": department } },
//         { arrayFilters: [{ "s.code": schoolCode }, { "d.code": department.code }] },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `Department ${department.code} saved in ${schoolCode}` });
//   }),
// );

// // serverside/src/routes/institutionSettings.ts - ADD DELETE endpoints

// // Add these to the existing institutionSettings router:

// // DELETE /institution-settings/schools/:schoolCode
// router.delete(
//   "/schools/:schoolCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     // Check if school has any programs
//     // const Program = mongoose.model("Program");
//     const hasPrograms = await Program.exists({ institution: institutionId, schoolCode });
    
//     if (hasPrograms) {
//       res.status(409).json({ 
//         message: "Cannot delete school with existing programs. Reassign or delete programs first." 
//       });
//       return;
//     }

//     const result = await InstitutionSettings.updateOne(
//       { institution: institutionId },
//       { $pull: { schools: { code: schoolCode } } }
//     );

//     if (result.modifiedCount === 0) {
//       res.status(404).json({ message: "School not found" });
//       return;
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: "school_deleted",
//       actor: req.user._id,
//       details: { schoolCode, institutionId },
//     });

//     res.json({ message: `School ${schoolCode} deleted successfully` });
//   })
// );

// // DELETE /institution-settings/schools/:schoolCode/departments/:departmentCode
// router.delete(
//   "/schools/:schoolCode/departments/:departmentCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode, departmentCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     // Check if department has any programs
//     // const Program = mongoose.model("Program");
//     const hasPrograms = await Program.exists({ 
//       institution: institutionId, 
//       schoolCode, 
//       departmentCode 
//     });
    
//     if (hasPrograms) {
//       res.status(409).json({ 
//         message: "Cannot delete department with existing programs. Reassign or delete programs first." 
//       });
//       return;
//     }

//     const result = await InstitutionSettings.updateOne(
//       { 
//         institution: institutionId,
//         "schools.code": schoolCode 
//       },
//       { $pull: { "schools.$.departments": { code: departmentCode } } }
//     );

//     if (result.modifiedCount === 0) {
//       res.status(404).json({ message: "Department not found" });
//       return;
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: "department_deleted",
//       actor: req.user._id,
//       details: { schoolCode, departmentCode, institutionId },
//     });

//     res.json({ message: `Department ${departmentCode} deleted successfully` });
//   })
// );

// export default router;






























// // serverside/src/routes/institutionSettings.ts — COMPLETE FIXED VERSION
// import express, { Response } from "express";
// import mongoose from "mongoose";
// import InstitutionSettings from "../models/InstitutionSettings";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler } from "../middleware/asyncHandler";
// import { logAudit } from "../lib/auditLogger";
// import { cached, invalidateCache } from "../utils/cache";
// import { invalidateSettingsCache } from "../utils/loadInstitutionSettings";
// import multer from "multer";
// import path from "path";
// import fs from "fs";
// import Program from "../models/Program";
// import User from "../models/User";

// // ── ApiError type (local — matches errorHandler.ts shape) ────────────────────
// interface ApiError {
//   statusCode: number;
//   message: string;
// }

// const router = express.Router();

// // ── Logo upload middleware ────────────────────────────────────────────────────
// const logoStorage = multer.diskStorage({
//   destination: (_req, _file, cb) => {
//     const dir = path.join(process.cwd(), "uploads", "logos");
//     fs.mkdirSync(dir, { recursive: true });
//     cb(null, dir);
//   },
//   filename: (_req, file, cb) => {
//     const ext = path.extname(file.originalname).toLowerCase();
//     cb(null, `logo-${Date.now()}${ext}`);
//   },
// });

// const uploadLogo = multer({
//   storage: logoStorage,
//   limits: { fileSize: 2 * 1024 * 1024 },
//   fileFilter: (_req, file, cb) => {
//     const allowed = [".png", ".jpg", ".jpeg", ".svg"];
//     const ext = path.extname(file.originalname).toLowerCase();
//     if (allowed.includes(ext)) {
//       cb(null, true);
//     } else {
//       cb(new Error("Only PNG, JPG, JPEG, SVG files allowed"));
//     }
//   },
// });

// // ── GET /institution-settings ─────────────────────────────────────────────────
// router.get(
//   "/",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await cached(
//       `settings:${institutionId}`,
//       () => InstitutionSettings.findOne({ institution: institutionId }).lean(),
//       300,
//     );

//     if (!settings) {
//       res.status(404).json({ message: "Settings not configured yet" });
//       return;
//     }

//     await logAudit(req, {
//       action: "institution_settings_viewed",
//       actor: req.user._id,
//       details: { institutionId },
//     });

//     res.json(settings);
//   }),
// );

// // ── POST /institution-settings ────────────────────────────────────────────────
// router.post(
//   "/",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const data = req.body;

//     // Validate: caWeight + examWeight must equal 100 if both provided
//     if (data.ruleSet?.caWeight !== undefined && data.ruleSet?.examWeight !== undefined) {
//       const total = Number(data.ruleSet.caWeight) + Number(data.ruleSet.examWeight);
//       if (Math.abs(total - 100) > 0.01) {
//         res.status(422).json({
//           message: `CA weight (${data.ruleSet.caWeight}) + Exam weight (${data.ruleSet.examWeight}) must equal 100`,
//         });
//         return;
//       }
//     }

//     const previous = await InstitutionSettings.findOne({
//       institution: institutionId,
//     }).lean();

//     const updated = await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { ...data, institution: institutionId } },
//       { new: true, upsert: true, setDefaultsOnInsert: true, runValidators: true },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: previous ? "institution_settings_updated" : "institution_settings_created",
//       actor: req.user._id,
//       details: { institutionId, settingsId: updated?._id?.toString() },
//     });

//     res.json({ message: "Settings saved successfully", settings: updated });
//   }),
// );

// // ── POST /institution-settings/logo ──────────────────────────────────────────
// router.post(
//   "/logo",
//   requireAuth,
//   requireRole("admin", "coordinator"),
//   uploadLogo.single("logo"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     if (!req.file) {
//       throw { statusCode: 400, message: "No file uploaded" } as ApiError;
//     }

//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       throw { statusCode: 400, message: "No institution linked to this account" } as ApiError;
//     }

//     const existing = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean();

//     if (existing?.branding?.universityLogoPath) {
//       const oldPath = path.join(process.cwd(), existing.branding.universityLogoPath);
//       if (fs.existsSync(oldPath)) {
//         fs.unlinkSync(oldPath);
//       }
//     }

//     const logoPath = `uploads/logos/${req.file.filename}`;

//     await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { "branding.universityLogoPath": logoPath } },
//       { upsert: true },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: "institution_logo_uploaded",
//       actor: req.user._id,
//       details: { institutionId, logoPath },
//     });

//     res.json({ message: "Logo uploaded successfully", path: logoPath });
//   }),
// );

// // ── GET /institution-settings/logo ──────────────────────────────────────────
// router.get(
//   "/logo",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked" });
//       return;
//     }

//     const settings = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean();

//     const logoPath = settings?.branding?.universityLogoPath;

//     if (!logoPath) {
//       res.status(404).json({ message: "No logo uploaded" });
//       return;
//     }

//     const fullPath = path.join(process.cwd(), logoPath);
//     if (!fs.existsSync(fullPath)) {
//       res.status(404).json({ message: "Logo file not found" });
//       return;
//     }

//     res.sendFile(fullPath);
//   }),
// );

// // ── GET /institution-settings/schools ─────────────────────────────────────────
// router.get(
//   "/schools",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("schools")
//       .lean();

//     res.json(settings?.schools || []);
//   }),
// );

// // ── GET /institution-settings/schools/:schoolCode ─────────────────────────────
// router.get(
//   "/schools/:schoolCode",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await InstitutionSettings.findOne(
//       { institution: institutionId, "schools.code": schoolCode },
//       { "schools.$": 1 }
//     ).lean();

//     if (!settings || !settings.schools?.length) {
//       res.status(404).json({ message: "School not found" });
//       return;
//     }

//     res.json(settings.schools[0]);
//   }),
// );

// // ── PATCH /institution-settings/schools ───────────────────────────────────────
// router.patch(
//   "/schools",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { school } = req.body;
//     const institutionId = req.user.institution?.toString();

//     if (!school?.code || !school?.name) {
//       throw { statusCode: 400, message: "School code and name are required" } as ApiError;
//     }

//     // Check for duplicate name (case-insensitive)
//     const nameClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.name": { $regex: new RegExp(`^${school.name.trim()}$`, "i") },
//       "schools.code": { $ne: school.code },
//     });

//     if (nameClash) {
//       throw {
//         statusCode: 409,
//         message: `A school named "${school.name}" already exists.`,
//       } as ApiError;
//     }

//     // Check for duplicate code
//     const codeClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.code": school.code,
//       "schools.name": { $ne: school.name },
//     });

//     if (codeClash) {
//       throw {
//         statusCode: 409,
//         message: `A school with code "${school.code}" already exists.`,
//       } as ApiError;
//     }

//     const existing = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.code": school.code,
//     });

//     if (existing) {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId, "schools.code": school.code },
//         { $set: { "schools.$": school } },
//       );
//     } else {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId },
//         { $push: { schools: school } },
//         { upsert: true },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `School ${school.code} saved successfully` });
//   }),
// );

// // ── DELETE /institution-settings/schools/:schoolCode ──────────────────────────
// router.delete(
//   "/schools/:schoolCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     // Check if school has any programs
//     const hasPrograms = await Program.exists({ institution: institutionId, schoolCode });

//     if (hasPrograms) {
//       throw {
//         statusCode: 409,
//         message: "Cannot delete school with existing programs. Reassign or delete programs first.",
//       } as ApiError;
//     }

//     const result = await InstitutionSettings.updateOne(
//       { institution: institutionId },
//       { $pull: { schools: { code: schoolCode } } }
//     );

//     if (result.modifiedCount === 0) {
//       throw { statusCode: 404, message: "School not found" } as ApiError;
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: "school_deleted",
//       actor: req.user._id,
//       details: { schoolCode, institutionId },
//     });

//     res.json({ message: `School ${schoolCode} deleted successfully` });
//   }),
// );

// // ── GET /institution-settings/schools/:schoolCode/departments ─────────────────
// router.get(
//   "/schools/:schoolCode/departments",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await InstitutionSettings.findOne(
//       { institution: institutionId, "schools.code": schoolCode },
//       { "schools.$.departments": 1 }
//     ).lean();

//     const school = settings?.schools?.[0];
//     res.json(school?.departments || []);
//   }),
// );

// // ── PATCH /institution-settings/schools/:schoolCode/departments ───────────────
// router.patch(
//   "/schools/:schoolCode/departments",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const { department } = req.body;
//     const institutionId = req.user.institution?.toString();

//     if (!department?.code || !department?.name) {
//       throw { statusCode: 400, message: "Department code and name are required" } as ApiError;
//     }

//     if (!schoolCode) {
//       throw { statusCode: 400, message: "School code is required" } as ApiError;
//     }

//     // Check for duplicate name within the same school (case-insensitive)
//     const nameClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.code": schoolCode,
//       "schools.departments.name": { $regex: new RegExp(`^${department.name.trim()}$`, "i") },
//       "schools.departments.code": { $ne: department.code },
//     });

//     if (nameClash) {
//       throw {
//         statusCode: 409,
//         message: `A department named "${department.name}" already exists in this school.`,
//       } as ApiError;
//     }

//     // Check for duplicate code within the same school
//     const codeClash = await InstitutionSettings.findOne({
//       institution: institutionId,
//       "schools.code": schoolCode,
//       "schools.departments.code": department.code,
//       "schools.departments.name": { $ne: department.name },
//     });

//     if (codeClash) {
//       throw {
//         statusCode: 409,
//         message: `A department with code "${department.code}" already exists in this school.`,
//       } as ApiError;
//     }

//     const result = await InstitutionSettings.updateOne(
//       {
//         institution: institutionId,
//         "schools.code": schoolCode,
//         "schools.departments.code": { $ne: department.code },
//       },
//       { $push: { "schools.$.departments": department } },
//     );

//     if (result.matchedCount === 0) {
//       // Department already exists — update it
//       await InstitutionSettings.updateOne(
//         {
//           institution: institutionId,
//           "schools.code": schoolCode,
//           "schools.departments.code": department.code,
//         },
//         { $set: { "schools.$[s].departments.$[d]": department } },
//         { arrayFilters: [{ "s.code": schoolCode }, { "d.code": department.code }] },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);

//     await logAudit(req, {
//       action: result.matchedCount > 0 ? "department_created" : "department_updated",
//       actor: req.user._id,
//       details: { schoolCode, departmentCode: department.code, institutionId },
//     });

//     res.json({ message: `Department ${department.code} saved in ${schoolCode}` });
//   }),
// );

// // ── DELETE /institution-settings/schools/:schoolCode/departments/:departmentCode ──
// router.delete(
//   "/schools/:schoolCode/departments/:departmentCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode, departmentCode } = req.params;
//     const institutionId = req.user.institution?.toString();

//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     // Check if department has any programs
//     const hasPrograms = await Program.exists({
//       institution: institutionId,
//       schoolCode,
//       departmentCode,
//     });

//     if (hasPrograms) {
//       throw {
//         statusCode: 409,
//         message: "Cannot delete department with existing programs. Reassign or delete programs first.",
//       } as ApiError;
//     }

//     // Check if department has any users (coordinators/lecturers)
//     const hasUsers = await User.exists({
//       institution: institutionId,
//       departmentCode,
//     });

//     if (hasUsers) {
//       throw {
//         statusCode: 409,
//         message: "Cannot delete department with assigned users. Reassign or remove users first.",
//       } as ApiError;
//     }

//     const result = await InstitutionSettings.updateOne(
//       {
//         institution: institutionId,
//         "schools.code": schoolCode,
//       },
//       { $pull: { "schools.$.departments": { code: departmentCode } } }
//     );

//     if (result.modifiedCount === 0) {
//       throw { statusCode: 404, message: "Department not found" } as ApiError;
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action: "department_deleted",
//       actor: req.user._id,
//       details: { schoolCode, departmentCode, institutionId },
//     });

//     res.json({ message: `Department ${departmentCode} deleted successfully` });
//   }),
// );

// export default router;





















// // serverside/src/routes/institutionSettings.ts — COMPLETE
// import express, { Response } from "express";
// import InstitutionSettings   from "../models/InstitutionSettings";
// import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
// import { asyncHandler }         from "../middleware/asyncHandler";
// import { logAudit }             from "../lib/auditLogger";
// import { cached, invalidateCache } from "../utils/cache";
// import { invalidateSettingsCache } from "../utils/loadInstitutionSettings";
// import multer from "multer";
// import path   from "path";
// import fs     from "fs";

// interface ApiError { statusCode: number; message: string }

// const router = express.Router();

// // ── Logo upload ───────────────────────────────────────────────────────────────
// const logoStorage = multer.diskStorage({
//   destination: (_req, _file, cb) => {
//     const dir = path.join(process.cwd(), "uploads", "logos");
//     fs.mkdirSync(dir, { recursive: true });
//     cb(null, dir);
//   },
//   filename: (_req, file, cb) => {
//     const ext = path.extname(file.originalname).toLowerCase();
//     cb(null, `logo-${Date.now()}${ext}`);
//   },
// });

// const uploadLogo = multer({
//   storage: logoStorage,
//   limits:  { fileSize: 2 * 1024 * 1024 },
//   fileFilter: (_req, file, cb) => {
//     const allowed = [".png", ".jpg", ".jpeg", ".svg"];
//     if (allowed.includes(path.extname(file.originalname).toLowerCase())) {
//       cb(null, true);
//     } else {
//       cb(new Error("Only PNG, JPG, JPEG, SVG files allowed"));
//     }
//   },
// });

// // ── GET /institution-settings — admin AND coordinator (read) ──────────────────
// router.get(
//   "/",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const settings = await cached(
//       `settings:${institutionId}`,
//       () => InstitutionSettings.findOne({ institution: institutionId }).lean(),
//       300,
//     );

//     await logAudit(req, {
//       action:  "institution_settings_viewed",
//       details: { institutionId },
//     });

//     // Return settings or empty object — coordinator sees same data, just can't edit
//     res.json(settings ?? null);
//   }),
// );

// // ── POST /institution-settings — ADMIN ONLY ────────────────────────────────────
// router.post(
//   "/",
//   requireAuth,
//   requireRole("admin"),   // ← coordinator REMOVED — admin only
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       res.status(400).json({ message: "No institution linked to this account" });
//       return;
//     }

//     const data = req.body;

//     if (
//       data.ruleSet?.caWeight !== undefined &&
//       data.ruleSet?.examWeight !== undefined
//     ) {
//       const total = Number(data.ruleSet.caWeight) + Number(data.ruleSet.examWeight);
//       if (Math.abs(total - 100) > 0.01) {
//         res.status(422).json({
//           message: `CA weight (${data.ruleSet.caWeight}) + Exam weight (${data.ruleSet.examWeight}) must equal 100`,
//         });
//         return;
//       }
//     }

//     const previous = await InstitutionSettings.findOne({ institution: institutionId }).lean();

//     const updated = await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { ...data, institution: institutionId } },
//       { new: true, upsert: true, setDefaultsOnInsert: true, runValidators: true },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action:  previous ? "institution_settings_updated" : "institution_settings_created",
//       details: { institutionId, settingsId: updated?._id?.toString() },
//     });

//     res.json({ message: "Settings saved successfully", settings: updated });
//   }),
// );

// // ── POST /institution-settings/logo — ADMIN ONLY ──────────────────────────────
// router.post(
//   "/logo",
//   requireAuth,
//   requireRole("admin"),   // ← admin only
//   uploadLogo.single("logo"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     if (!req.file) {
//       throw { statusCode: 400, message: "No file uploaded" } as ApiError;
//     }

//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) {
//       throw { statusCode: 400, message: "No institution linked" } as ApiError;
//     }

//     // Delete old logo
//     const existing = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean() as { branding?: { universityLogoPath?: string } } | null;

//     if (existing?.branding?.universityLogoPath) {
//       const oldPath = path.join(process.cwd(), existing.branding.universityLogoPath);
//       if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
//     }

//     const logoPath = `uploads/logos/${req.file.filename}`;

//     await InstitutionSettings.findOneAndUpdate(
//       { institution: institutionId },
//       { $set: { "branding.universityLogoPath": logoPath } },
//       { upsert: true },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId);

//     await logAudit(req, {
//       action:  "institution_logo_uploaded",
//       details: { institutionId, logoPath },
//     });

//     res.json({ message: "Logo uploaded successfully", path: logoPath });
//   }),
// );

// // ── GET /institution-settings/logo — serve logo (all authenticated users) ─────
// router.get(
//   "/logo",
//   requireAuth,
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const institutionId = req.user.institution?.toString();
//     if (!institutionId) { res.status(400).json({ message: "No institution linked" }); return; }

//     const settings = await InstitutionSettings.findOne({ institution: institutionId })
//       .select("branding.universityLogoPath")
//       .lean() as { branding?: { universityLogoPath?: string } } | null;

//     const logoPath = settings?.branding?.universityLogoPath;
//     if (!logoPath) { res.status(404).json({ message: "No logo uploaded" }); return; }

//     const fullPath = path.join(process.cwd(), logoPath);
//     if (!fs.existsSync(fullPath)) { res.status(404).json({ message: "Logo file not found" }); return; }

//     res.sendFile(fullPath);
//   }),
// );

// // ── PATCH /institution-settings/schools — ADMIN ONLY ─────────────────────────
// router.patch(
//   "/schools",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { school }    = req.body;
//     const institutionId = req.user.institution?.toString();

//     if (!school?.code || !school?.name) {
//       throw { statusCode: 400, message: "School code and name are required" } as ApiError;
//     }

//     // Check name uniqueness (different code, same name = clash)
//     const nameClash = await InstitutionSettings.findOne({
//       institution:  institutionId,
//       "schools.name": { $regex: new RegExp(`^${school.name.trim()}$`, "i") },
//       "schools.code": { $ne: school.code.toUpperCase() },
//     });
//     if (nameClash) {
//       throw { statusCode: 409, message: `A school named "${school.name}" already exists.` } as ApiError;
//     }

//     const existing = await InstitutionSettings.findOne({
//       institution:    institutionId,
//       "schools.code": school.code.toUpperCase(),
//     });

//     if (existing) {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId, "schools.code": school.code.toUpperCase() },
//         { $set: { "schools.$": { ...school, code: school.code.toUpperCase() } } },
//       );
//     } else {
//       await InstitutionSettings.updateOne(
//         { institution: institutionId },
//         { $push: { schools: { ...school, code: school.code.toUpperCase() } } },
//         { upsert: true },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `School ${school.code} saved` });
//   }),
// );

// // ── PATCH /institution-settings/schools/:schoolCode/departments — ADMIN ONLY ──
// router.patch(
//   "/schools/:schoolCode/departments",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const { department } = req.body;
//     const institutionId  = req.user.institution?.toString();

//     if (!department?.code || !department?.name) {
//       throw { statusCode: 400, message: "Department code and name are required" } as ApiError;
//     }

//     // Check name uniqueness within the school
//     const nameClash = await InstitutionSettings.findOne({
//       institution:    institutionId,
//       "schools.code": schoolCode.toUpperCase(),
//       "schools.departments.name": {
//         $regex: new RegExp(`^${department.name.trim()}$`, "i"),
//       },
//       "schools.departments.code": { $ne: department.code.toUpperCase() },
//     });
//     if (nameClash) {
//       throw {
//         statusCode: 409,
//         message: `A department named "${department.name}" already exists in this school.`,
//       } as ApiError;
//     }

//     const result = await InstitutionSettings.updateOne(
//       {
//         institution:    institutionId,
//         "schools.code": schoolCode.toUpperCase(),
//         "schools.departments.code": { $ne: department.code.toUpperCase() },
//       },
//       { $push: { "schools.$.departments": { ...department, code: department.code.toUpperCase() } } },
//     );

//     if (result.matchedCount === 0) {
//       await InstitutionSettings.updateOne(
//         {
//           institution:    institutionId,
//           "schools.code": schoolCode.toUpperCase(),
//           "schools.departments.code": department.code.toUpperCase(),
//         },
//         {
//           $set: {
//             "schools.$[s].departments.$[d]": {
//               ...department,
//               code: department.code.toUpperCase(),
//             },
//           },
//         },
//         {
//           arrayFilters: [
//             { "s.code": schoolCode.toUpperCase() },
//             { "d.code": department.code.toUpperCase() },
//           ],
//         },
//       );
//     }

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `Department ${department.code} saved in ${schoolCode}` });
//   }),
// );

// // ── DELETE /institution-settings/schools/:schoolCode — ADMIN ONLY ─────────────
// router.delete(
//   "/schools/:schoolCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode } = req.params;
//     const institutionId  = req.user.institution?.toString();

//     await InstitutionSettings.updateOne(
//       { institution: institutionId },
//       { $pull: { schools: { code: schoolCode.toUpperCase() } } },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `School ${schoolCode} removed` });
//   }),
// );

// // ── DELETE /institution-settings/schools/:schoolCode/departments/:deptCode ────
// router.delete(
//   "/schools/:schoolCode/departments/:deptCode",
//   requireAuth,
//   requireRole("admin"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
//     const { schoolCode, deptCode } = req.params;
//     const institutionId            = req.user.institution?.toString();

//     await InstitutionSettings.updateOne(
//       { institution: institutionId, "schools.code": schoolCode.toUpperCase() },
//       {
//         $pull: {
//           "schools.$.departments": { code: deptCode.toUpperCase() },
//         },
//       },
//     );

//     invalidateCache(`settings:${institutionId}`);
//     invalidateSettingsCache(institutionId!);
//     res.json({ message: `Department ${deptCode} removed from ${schoolCode}` });
//   }),
// );

// export default router;




















































// serverside/src/routes/institutionSettings.ts — COMPLETE, FINAL VERSION
import express, { Response } from "express";
import InstitutionSettings   from "../models/InstitutionSettings";
import { AuthenticatedRequest, requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler }         from "../middleware/asyncHandler";
import { logAudit }             from "../lib/auditLogger";
import { cached, invalidateCache } from "../utils/cache";
import { invalidateSettingsCache } from "../utils/loadInstitutionSettings";
import multer from "multer";
import path   from "path";
import fs     from "fs";

interface ApiError { statusCode: number; message: string }

const router = express.Router();

// ── Logo upload (admin only) ──────────────────────────────────────────────────
const logoStorage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    const dir = path.join(process.cwd(), "uploads", "logos");
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `logo-${Date.now()}${ext}`);
  },
});

const uploadLogo = multer({
  storage: logoStorage,
  limits:  { fileSize: 2 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const allowed = [".png", ".jpg", ".jpeg", ".svg"];
    if (allowed.includes(path.extname(file.originalname).toLowerCase())) {
      cb(null, true);
    } else {
      cb(new Error("Only PNG, JPG, JPEG, SVG files allowed"));
    }
  },
});

// ── GET /institution-settings — ALL authenticated users (read) ────────────────
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
      res.status(400).json({ message: "No institution linked to this account" });
      return;
    }

    const settings = await cached(
      `settings:${institutionId}`,
      () => InstitutionSettings.findOne({ institution: institutionId }).lean(),
      300,
    );

    await logAudit(req, {
      action:  "institution_settings_viewed",
      details: { institutionId },
    });

    res.json(settings ?? null);
  }),
);

// ── POST /institution-settings — COORDINATOR + ADMIN ─────────────────────────
// Coordinator manages: ruleSet, gradingScale, waaClassification,
//                      semesterWeights, enforceRegNoPattern, supportedIntakes
// Admin manages:       docMeta, schools, departments (via separate PATCH routes)
//
// Both roles can POST to this endpoint, but each saves different fields.
// The server accepts the full payload and applies whichever fields are present.
// Field-level authorization is enforced inside the handler.
router.post(
  "/",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
      res.status(400).json({ message: "No institution linked to this account" });
      return;
    }

    const data = req.body as Record<string, unknown>;
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
    const filteredData: Record<string, unknown> = {};
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
      const rs = filteredData.ruleSet as Record<string, number>;
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

    const previous = await InstitutionSettings.findOne({ institution: institutionId }).lean();

    const updated = await InstitutionSettings.findOneAndUpdate(
      { institution: institutionId },
      { $set: { ...filteredData, institution: institutionId } },
      { new: true, upsert: true, setDefaultsOnInsert: true, runValidators: true },
    );

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId);

    await logAudit(req, {
      action:  previous ? "institution_settings_updated" : "institution_settings_created",
      details: { institutionId, role, fields: Object.keys(filteredData) },
    });

    res.json({ message: "Settings saved successfully", settings: updated });
  }),
);

// ── POST /institution-settings/logo — ADMIN ONLY ──────────────────────────────
router.post(
  "/logo",
  requireAuth,
  requireRole("admin"),
  uploadLogo.single("logo"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    if (!req.file) {
      throw { statusCode: 400, message: "No file uploaded" } as ApiError;
    }

    const institutionId = req.user.institution?.toString();
    if (!institutionId) {
      throw { statusCode: 400, message: "No institution linked" } as ApiError;
    }

    // Delete old logo file
    const existing = await InstitutionSettings.findOne({ institution: institutionId })
      .select("branding.universityLogoPath")
      .lean() as { branding?: { universityLogoPath?: string } } | null;

    if (existing?.branding?.universityLogoPath) {
      const oldPath = path.join(process.cwd(), existing.branding.universityLogoPath);
      if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
    }

    const logoPath = `uploads/logos/${req.file.filename}`;
    

    await InstitutionSettings.findOneAndUpdate(
      { institution: institutionId },
      { $set: { "branding.universityLogoPath": logoPath } },
      { upsert: true },
    );

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId);

    await logAudit(req, {
      action:  "institution_logo_uploaded",
      details: { institutionId, logoPath },
    });

    res.json({ message: "Logo uploaded successfully", path: logoPath });
  }),
);

// ── GET /institution-settings/logo — serve logo (all authenticated users) ─────
router.get(
  "/logo",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const institutionId = req.user.institution?.toString();
    if (!institutionId) { res.status(400).json({ message: "No institution linked" }); return; }

    const settings = await InstitutionSettings.findOne({ institution: institutionId })
      .select("branding.universityLogoPath")
      .lean() as { branding?: { universityLogoPath?: string } } | null;

    const logoPath = settings?.branding?.universityLogoPath;
    if (!logoPath) { res.status(404).json({ message: "No logo uploaded" }); return; }

    const fullPath = path.join(process.cwd(), logoPath);
    if (!fs.existsSync(fullPath)) { res.status(404).json({ message: "Logo file not found" }); return; }

    res.sendFile(fullPath);
  }),
);

// ── PATCH /institution-settings/schools — ADMIN ONLY ─────────────────────────
router.patch(
  "/schools",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { school }    = req.body as { school: Record<string, unknown> };
    const institutionId = req.user.institution?.toString();

    if (!school?.code || !school?.name) {
      throw { statusCode: 400, message: "School code and name are required" } as ApiError;
    }

    // Uniqueness: same name, different code = clash
    const nameClash = await InstitutionSettings.findOne({
      institution:  institutionId,
      "schools.name": { $regex: new RegExp(`^${String(school.name).trim()}$`, "i") },
      "schools.code": { $ne: String(school.code).toUpperCase() },
    });
    if (nameClash) {
      throw {
        statusCode: 409,
        message: `A school named "${school.name}" already exists.`,
      } as ApiError;
    }

    const existing = await InstitutionSettings.findOne({
      institution:    institutionId,
      "schools.code": String(school.code).toUpperCase(),
    });

    const schoolDoc = { ...school, code: String(school.code).toUpperCase() };

    if (existing) {
      await InstitutionSettings.updateOne(
        { institution: institutionId, "schools.code": String(school.code).toUpperCase() },
        { $set: { "schools.$": schoolDoc } },
      );
    } else {
      await InstitutionSettings.updateOne(
        { institution: institutionId },
        { $push: { schools: schoolDoc } },
        { upsert: true },
      );
    }

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId!);
    res.json({ message: `School ${school.code} saved` });
  }),
);

// ── PATCH /institution-settings/schools/:schoolCode/departments — ADMIN ONLY ──
// Reg number patterns inside departments are COORDINATOR responsibility.
// Admin creates the department shell; coordinator fills in regNoPatterns.
router.patch(
  "/schools/:schoolCode/departments",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { schoolCode } = req.params;
    const { department } = req.body as { department: Record<string, unknown> };
    const institutionId  = req.user.institution?.toString();

    if (!department?.code || !department?.name) {
      throw { statusCode: 400, message: "Department code and name are required" } as ApiError;
    }

    const nameClash = await InstitutionSettings.findOne({
      institution:    institutionId,
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
      } as ApiError;
    }

    const deptDoc = { ...department, code: String(department.code).toUpperCase() };

    const result = await InstitutionSettings.updateOne(
      {
        institution:    institutionId,
        "schools.code": schoolCode.toUpperCase(),
        "schools.departments.code": { $ne: String(department.code).toUpperCase() },
      },
      { $push: { "schools.$.departments": deptDoc } },
    );

    if (result.matchedCount === 0) {
      await InstitutionSettings.updateOne(
        {
          institution:    institutionId,
          "schools.code": schoolCode.toUpperCase(),
          "schools.departments.code": String(department.code).toUpperCase(),
        },
        { $set: { "schools.$[s].departments.$[d]": deptDoc } },
        {
          arrayFilters: [
            { "s.code": schoolCode.toUpperCase() },
            { "d.code": String(department.code).toUpperCase() },
          ],
        },
      );
    }

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId!);
    res.json({ message: `Department ${department.code} saved in ${schoolCode}` });
  }),
);

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
router.patch(
  "/schools/:schoolCode/departments/:deptCode/reg-patterns",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { schoolCode, deptCode } = req.params;
    const { regNoPatterns } = req.body as {
      regNoPatterns: Array<{
        prefix: string;
        separator: string;
        yearDigits: number;
        example: string;
        manualRegex?: string;
      }>;
    };
    const institutionId = req.user.institution?.toString();

    // ENFORCE: coordinators can only edit their OWN department
    if (req.user.role === "coordinator" && !req.user.institutionWide) {
      if (
        req.user.schoolCode?.toUpperCase() !== schoolCode.toUpperCase() ||
        req.user.departmentCode?.toUpperCase() !== deptCode.toUpperCase()
      ) {
        throw {
          statusCode: 403,
          message: "You can only configure registration number patterns for your own department.",
        } as ApiError;
      }
    }

    if (!Array.isArray(regNoPatterns)) {
      throw { statusCode: 400, message: "regNoPatterns must be an array" } as ApiError;
    }

    await InstitutionSettings.updateOne(
      {
        institution: institutionId,
        "schools.code": schoolCode.toUpperCase(),
        "schools.departments.code": deptCode.toUpperCase(),
      },
      {
        $set: {
          "schools.$[s].departments.$[d].regNoPatterns": regNoPatterns,
          "schools.$[s].departments.$[d].enforceRegNoPattern": true,
        },
      },
      {
        arrayFilters: [
          { "s.code": schoolCode.toUpperCase() },
          { "d.code": deptCode.toUpperCase() },
        ],
      }
    );

    // Also update the top-level enforceRegNoPattern flag
    await InstitutionSettings.updateOne(
      { institution: institutionId },
      { $set: { enforceRegNoPattern: true } }
    );

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId!);

    await logAudit(req, {
      action: "reg_no_patterns_updated",
      details: { institutionId, schoolCode, deptCode, patternCount: regNoPatterns.length },
    });

    res.json({ message: `Registration number patterns updated for ${deptCode}` });
  })
);

// ── DELETE /institution-settings/schools/:schoolCode — ADMIN ONLY ─────────────
router.delete(
  "/schools/:schoolCode",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { schoolCode } = req.params;
    const institutionId  = req.user.institution?.toString();

    await InstitutionSettings.updateOne(
      { institution: institutionId },
      { $pull: { schools: { code: schoolCode.toUpperCase() } } },
    );

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId!);
    res.json({ message: `School ${schoolCode} removed` });
  }),
);

// ── DELETE /institution-settings/schools/:schoolCode/departments/:deptCode ────
router.delete(
  "/schools/:schoolCode/departments/:deptCode",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response): Promise<void> => {
    const { schoolCode, deptCode } = req.params;
    const institutionId            = req.user.institution?.toString();

    await InstitutionSettings.updateOne(
      { institution: institutionId, "schools.code": schoolCode.toUpperCase() },
      { $pull: { "schools.$.departments": { code: deptCode.toUpperCase() } } },
    );

    invalidateCache(`settings:${institutionId}`);
    invalidateSettingsCache(institutionId!);
    res.json({ message: `Department ${deptCode} removed from ${schoolCode}` });
  }),
);

export default router;