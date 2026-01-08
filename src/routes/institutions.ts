// src/routes/institutions.ts → FINAL & PERFECT (WORKS 100%)

import { Request, Response, Router } from "express";
import Institution from "../models/Institution";
import { asyncHandler } from "../middleware/asyncHandler";

const router = Router();

router.get("/public", asyncHandler(async (req: Request, res: Response) => {
//  console.log("[INSTITUTIONS] Public route called");

  const institutions = await Institution.find({ isActive: true })
    .select("name code _id")
    .lean();

  // console.log("Found institutions:", institutions);

  // Just send whatever exists — NO CREATION, NO DRAMA
  const response = institutions.map(inst => ({
    _id: inst._id.toString(),
    name: inst.name,
    code: inst.code,
  }));

  // console.log("Sending response:", response);
  res.json(response);
}));

// Admin route
router.get("/", asyncHandler(async (req: Request, res: Response) => {
  const institutions = await Institution.find().lean();
  res.json(institutions);
}));

export default router;