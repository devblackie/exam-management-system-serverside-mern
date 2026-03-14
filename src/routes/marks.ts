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
import * as xlsx from "xlsx";
import path from "path";
import Unit from "../models/Unit";
import { generateDirectScoresheetTemplate } from "../utils/directTemplate";
import { importDirectMarksFromBuffer } from "../services/directMarksImporter";

const router = Router();

// Route: GET /api/marks/template?programId=&unitId=&academicYearId=&yearOfStudy=&semester=
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
      examMode,
      unitType,
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
        // console.warn("Logo file not found at:", logoPath);
        logoBuffer = Buffer.alloc(0); // Fallback to empty if file is missing
      }

      // --- 2. GENERATE THE EXCEL BUFFER ---
      const excelBuffer = await generateFullScoresheetTemplate(
        new mongoose.Types.ObjectId(programId as string),
        new mongoose.Types.ObjectId(unitId as string),
        parseInt(yearOfStudy as string, 10),
        parseInt(semester as string, 10),
        new mongoose.Types.ObjectId(academicYearId as string),
        logoBuffer, // Pass the actual image data here
        examMode as string,
        (unitType as any) || "theory"
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
      // console.error("Error generating scoresheet template:", error.message);

      // Return a 400 for invalid ID format, otherwise a 500
      const status = error.message.includes("invalid") ? 400 : 500;

      return res.status(status).json({
        message: "Failed to generate scoresheet template. Check input IDs.",
        error: error.message,
      });
    }
  })
);

router.get(
  "/direct-template",
  requireAuth, 
  requireRole("coordinator", "admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, unitId, academicYearId, yearOfStudy, semester } = req.query;

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

    const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
    const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

    const buffer = await generateDirectScoresheetTemplate(
      new mongoose.Types.ObjectId(programId as string),
        new mongoose.Types.ObjectId(unitId as string),
        parseInt(yearOfStudy as string, 10),
        parseInt(semester as string, 10),
        new mongoose.Types.ObjectId(academicYearId as string),
        logoBuffer
    );

    res
      .header(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )
      .header("Access-Control-Expose-Headers", "Content-Disposition")
      .attachment(fileName)
      .send(buffer);
  } catch (error: any) {
    // Handle cases where IDs are invalid or DB lookup fails
    // console.error("Error generating scoresheet template:", error.message);

    // Return a 400 for invalid ID format, otherwise a 500
    const status = error.message.includes("invalid") ? 400 : 500;

    return res.status(status).json({
      message: "Failed to generate scoresheet template. Check input IDs.",
      error: error.message,
    });
  }
})
);

// Route: POST marks/upload
router.post(
  "/upload",
  requireAuth,
  requireRole("coordinator", "admin"),
  uploadMarksFile.single("file"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!req.file) {
      return res.status(400).json({ message: "No file uploaded" });
    }

    try {
      // 1. Peek at the file to detect the template type
      const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Look at Row 15, Column E (Excel coordinate E15)
      // Headers are usually at index 14 in sheet_to_json or direct access
      const columnEHeader = sheet["E15"]?.v?.toString().toUpperCase() || "";

      let result;

      // 2. Logic Switch based on Column E content
      if (columnEHeader.includes("CA TOTAL")) {
        // This is the simplified "Direct" template
        console.log("[Upload] Detected Direct Entry Template");
        result = await importDirectMarksFromBuffer(req.file.buffer, req.file.originalname, req);
      } else {
        // Default to the original detailed logic
        console.log("[Upload] Detected Detailed Breakdown Template");
        result = await importMarksFromBuffer(req.file.buffer, req.file.originalname, req);
      }

      await logAudit(req, {
        action: "marks_upload_success",
        details: { 
          templateType: columnEHeader.includes("CA TOTAL") ? "direct" : "detailed",
          total: result.total, 
          success: result.success 
        },
      });

      res.json({ message: "Import completed", ...result });
    } catch (err: any) {
            await logAudit(req, {
              action: "marks_upload_failed",
              details: { error: err.message, filename: req.file.originalname },
            });
            throw err;
    }
  })
);

// Route: POST marks/upload
// router.post(
//   "/upload",
//   requireAuth,
//   requireRole("coordinator", "admin"),
//   uploadMarksFile.single("file"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     if (!req.file) {
//       await logAudit(req, { action: "marks_upload_no_file" });
//       return res.status(400).json({ message: "No file uploaded" });
//     }

//     try {
//       const result = await importMarksFromBuffer( req.file.buffer, req.file.originalname, req );

//       await logAudit(req, {
//         action: "marks_upload_success",
//         details: {
//           filename: req.file.originalname,
//           total: result.total,
//           success: result.success,
//           failed: result.errors.length,
//         },
//       });

//       res.json({
//         message: "Import completed",
//         ...result,
//       });
//     } catch (err: any) {
//       await logAudit(req, {
//         action: "marks_upload_failed",
//         details: { error: err.message, filename: req.file.originalname },
//       });
//       throw err; // asyncHandler will catch
//     }
//   })
// );

// Add to src/routes/marks.ts



export default router;

// Would you like me to help you refine the frontend downloadTemplate function to handle potential network timeouts for these large file generations?
