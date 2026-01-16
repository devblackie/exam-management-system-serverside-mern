// serverside/src/routes/promote.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { bulkPromoteClass, previewPromotion } from "../services/statusEngine";
import Program from "../models/Program";
import { generatePromotionWordDoc } from "../utils/promotionReport";
import fs from "fs";
import path from "path";

const router = Router();

router.post(
  "/preview-promotion",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    if (!programId || !yearToPromote || !academicYearName) {
      return res.status(400).json({ error: "Missing parameters" });
    }

    const previewData = await previewPromotion(programId, yearToPromote, academicYearName);
    res.json({ success: true, data: previewData });
  })
);

router.post(
  "/bulk-promote",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    if (!programId || !yearToPromote || !academicYearName) {
      return res.status(400).json({ error: "Missing required promotion parameters" });
    }

    const results = await bulkPromoteClass(programId, yearToPromote, academicYearName);

    res.json({
      success: true,
      message: `Process completed: ${results.promoted} promoted, ${results.failed} failed.`,
      data: results
    });
  })
);

// 2. DOWNLOAD PROMOTION REPORT (WORD DOC)
router.post(
  "/download-report",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    // Fetch necessary data
    const preview = await previewPromotion(programId, yearToPromote, academicYearName);
    const program = await Program.findById(programId).lean();

    // --- LOAD LOGO (Same pattern as marks.ts) ---
    const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
    let logoBuffer = Buffer.alloc(0);
    if (fs.existsSync(logoPath)) {
      logoBuffer = fs.readFileSync(logoPath);
    }

    // Generate Word Buffer
    const docBuffer = await generatePromotionWordDoc({
      programName: program?.name || "Unknown Program",
      academicYear: academicYearName,
      yearOfStudy: yearToPromote,
      eligible: preview.eligible,
      blocked: preview.blocked,
      logoBuffer
    });

    const fileName = `Promotion_Report_${program?.code}_Year${yearToPromote}.docx`;

    res
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .header("Content-Disposition", `attachment; filename=${fileName}`)
      .send(docBuffer);
  })
);

export default router;