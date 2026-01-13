// src/routes/marks.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { generateFullScoresheetTemplate } from "../utils/uploadTemplate";
import { uploadMarksFile } from "../middleware/upload";
import { importMarksFromBuffer } from "../services/marksImporter";
import { logAudit } from "../lib/auditLogger";
import mongoose from "mongoose";
import fs from "fs";
import path from "path";
import Unit from "../models/Unit";

const router = Router();

// --- TEMPORARY INDEX CLEANUP ---
// mongoose.connection.once("open", async () => {
//   try {
//     const collection = mongoose.connection.collection("marks");
//     // This command tells MongoDB to delete the specific index causing the error
//     await collection.dropIndex("student_1_unit_1_academicYear_1");
//     console.log("✅ SUCCESS: The ghost index has been deleted.");
//   } catch (err: any) {
//     if (err.codeName === "IndexNotFound") {
//       console.log("ℹ️ INFO: Ghost index not found, it might already be gone.");
//     } else {
//       console.error("❌ ERROR deleting index:", err.message);
//     }
//   }
// });
// -------------------------------

router.get(
  "/template",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, unitId, academicYearId, yearOfStudy, semester } =
      req.query;

    if (!programId || !unitId || !academicYearId || !yearOfStudy || !semester) {
      return res.status(400).json({
        error:
          "Missing required parameters: programId, unitId, academicYearId, yearOfStudy, semester",
      });
    }

    try {
      // Validation and Casting: Checks if the input string can be an ObjectId.
      if (
        !mongoose.Types.ObjectId.isValid(programId as string) ||
        !mongoose.Types.ObjectId.isValid(unitId as string) ||
        !mongoose.Types.ObjectId.isValid(academicYearId as string)
      ) {
        throw new Error(
          "One or more provided IDs (programId, unitId, academicYearId) are invalid."
        );
      }

      // Fetch Unit details for the filename
      const unit = await Unit.findById(unitId).lean();
      // 1. Trim invisible spaces first
      const rawCode = (unit?.code || "UNIT").trim();
      const rawName = (unit?.name || "TEMPLATE").trim();

      const cleanName =
        `${rawCode}_${rawName}`
          .replace(/[^a-zA-Z0-9]/g, "_") // Replace anything not a letter or number
          .replace(/_+/g, "_") // Collapse multiple underscores (___ -> _)
          .replace(/^_|_$/g, "") // Remove _ from start or end
          ?.toUpperCase() || "TEMPLATE";

      const fileName = `Scoresheet_${cleanName}.xlsx`;

      // --- 1. LOAD THE LOGO IMAGE ---
      const logoPath = path.join(
        __dirname,
        "../../public/institutionLogoExcel.png"
      );

      let logoBuffer: Buffer;
      if (fs.existsSync(logoPath)) {
        logoBuffer = fs.readFileSync(logoPath);
      } else {
        console.warn("Logo file not found at:", logoPath);
        logoBuffer = Buffer.alloc(0); // Fallback to empty if file is missing
      }

      // --- 2. GENERATE THE EXCEL BUFFER ---
      const excelBuffer = await generateFullScoresheetTemplate(
        new mongoose.Types.ObjectId(programId as string),
        new mongoose.Types.ObjectId(unitId as string),
        parseInt(yearOfStudy as string, 10),
        parseInt(semester as string, 10),
        new mongoose.Types.ObjectId(academicYearId as string),
        logoBuffer // Pass the actual image data here
      );

      res
        // .header("Content-Type", "text/csv")
        .header(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        .header("Access-Control-Expose-Headers", "Content-Disposition")
        .attachment(fileName)
        // .send(templateContent);
        .send(excelBuffer);
    } catch (error: any) {
      // Handle cases where IDs are invalid or DB lookup fails
      console.error("Error generating scoresheet template:", error.message);

      // Return a 400 for invalid ID format, otherwise a 500
      const status = error.message.includes("invalid") ? 400 : 500;

      return res.status(status).json({
        message: "Failed to generate scoresheet template. Check input IDs.",
        error: error.message,
      });
    }
  })
);

router.post(
  "/upload",
  requireAuth,
  requireRole("coordinator", "admin"),
  uploadMarksFile.single("file"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!req.file) {
      await logAudit(req, { action: "marks_upload_no_file" });
      return res.status(400).json({ message: "No file uploaded" });
    }

    try {
      const result = await importMarksFromBuffer(
        req.file.buffer,
        req.file.originalname,
        req
      );

      await logAudit(req, {
        action: "marks_upload_success",
        details: {
          filename: req.file.originalname,
          total: result.total,
          success: result.success,
          failed: result.errors.length,
        },
      });

      res.json({
        message: "Import completed",
        ...result,
      });
    } catch (err: any) {
      await logAudit(req, {
        action: "marks_upload_failed",
        details: { error: err.message, filename: req.file.originalname },
      });
      throw err; // asyncHandler will catch
    }
  })
);

export default router;
