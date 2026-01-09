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

const router = Router();

router.get(
  "/template",
  requireAuth,
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const {
      programId,
      unitId,
      academicYearId,
      yearOfStudy, 
      semester, 
    } = req.query;

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

     // --- 1. LOAD THE LOGO IMAGE ---
      const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png"); 
      
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

      const fileName = `Scoresheet_Template_${unitId}.xlsx`;

      res
        // .header("Content-Type", "text/csv")
        .header(
          "Content-Type", 
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
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
