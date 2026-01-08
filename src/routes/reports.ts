// src/routes/reports.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { generateStudentTranscript, generatePassList, generateConsolidatedMarksheet } from "../services/pdfGenerator";
import type { AuthenticatedRequest } from "../middleware/auth";

const router = Router();

router.get("/transcript/:regNo", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  await generateStudentTranscript(req.params.regNo, res);
}));

router.get("/pass-list/:academicYearId", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  await generatePassList(req.params.academicYearId, req.user.institution.toString(), res);
}));

router.get("/consolidated-marksheet/:academicYearId", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
  await generateConsolidatedMarksheet(req.params.academicYearId, req.user.institution.toString(), res);
}));

export default router;