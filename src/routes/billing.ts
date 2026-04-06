
// FILE 2: serverside/src/routes/billing.ts
//
// This is the backend route. Mount it in app.ts as:
//   import billingRouter from "./routes/billing";
//   app.use("/billing", billingRouter);

import { Router, Response } from "express";
import Student from "../models/Student";
import Billing from "../models/Billing";
import { requireAuth, requireRole, AuthenticatedRequest } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";

const router = Router();

// GET /billing/summary
// Returns the current plan, seat usage, and last 3 invoices for the institution.
router.get(
  "/summary",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution;

    // Live seat count from the Student collection
    const seatsUsed = await Student.countDocuments({
      institution: institutionId,
      status:      { $in: ["active", "repeat"] },
    });

    // Billing record stored in the Billing collection (one per institution)
    const record = await Billing.findOne({ institution: institutionId }).lean();

    if (!record) {
      res.status(404).json({
        message: "No billing record found for this institution. Contact support.",
      });
      return;
    }

    res.json({
      planName:          record.planName,
      billingCycle:      record.billingCycle,
      seatLimit:         record.seatLimit,
      seatsUsed,
      nextInvoiceAmount: record.nextInvoiceAmount,
      nextInvoiceDate:   record.nextInvoiceDate,
      invoices:          (record.invoices ?? []).slice(-6).reverse(),
    });
  })
);

export default router;
