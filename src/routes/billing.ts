// serverside/src/routes/billing.ts
//
// ENDPOINTS
// ──────────────────────────────────────────────────────────────────────────
// GET    /billing/summary           → dashboard overview (admin)
// GET    /billing/invoices          → paginated invoice list (admin)
// GET    /billing/invoices/:id      → single invoice detail (admin)
// POST   /billing/invoices/generate → manually trigger invoice (admin)
// PATCH  /billing/invoices/:id/pay  → record manual payment (admin)
// PATCH  /billing/invoices/:id/void → void invoice (admin)
// GET    /billing/usage             → usage history (admin)
// PATCH  /billing/plan              → change plan (admin)
// PATCH  /billing/contact           → update billing contact (admin)
// PATCH  /billing/cycle             → switch monthly/annual (admin)
// POST   /billing/cron/invoices     → trigger monthly generation (system)
// POST   /billing/cron/overdue      → mark overdue invoices (system)

import { Router, Response } from "express";
import {
  requireAuth,
  requireRole,
  AuthenticatedRequest,
} from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";
import Billing from "../models/Billing";
import Student from "../models/Student";
import {
  generateInvoice,
  recordPayment,
  voidInvoice,
  changePlan,
  markOverdueInvoices,
  generateMonthlyInvoices,
  takeUsageSnapshot,
  PLAN_CATALOGUE,
} from "../services/billingService";

const router = Router();
router.use(requireAuth);

// ── GET /billing/summary ──────────────────────────────────────────────────────
// Dashboard overview card. Shows plan, seat usage, next invoice, and last 3 invoices.
router.get(
  "/summary",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();

    const billing = await Billing.findOne({
      institution: institutionId,
    }).lean();
    if (!billing) {
      res
        .status(404)
        .json({
          message:
            "No billing record found. Contact support to set up billing.",
        });
      return;
    }

    const [activeStudents, totalStudents] = await Promise.all([
      Student.countDocuments({
        institution: institutionId,
        status: { $in: ["active", "repeat"] },
      }),
      Student.countDocuments({ institution: institutionId }),
    ]);

    const overage = Math.max(0, activeStudents - billing.seatLimit);
    const recentInvoices = [...(billing.invoices ?? [])]
      .sort(
        (a, b) =>
          new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
      )
      .slice(0, 3);

    const overdueCount = (billing.invoices ?? []).filter(
      (i) => i.status === "overdue",
    ).length;
    const unpaidTotal = (billing.invoices ?? [])
      .filter((i) => ["sent", "overdue"].includes(i.status))
      .reduce((sum, i) => sum + i.total, 0);

    res.json({
      plan: {
        name: billing.planName,
        cycle: billing.billingCycle,
        seatLimit: billing.seatLimit,
        basePrice: billing.basePrice,
        overageRate: billing.overageRate,
        currency: billing.currency,
        taxRate: billing.taxRate,
        isCustomPlan: billing.isCustomPlan,
      },
      usage: {
        activeStudents,
        totalStudents,
        seatLimit: billing.seatLimit,
        overage,
        usagePercent: Math.round((activeStudents / billing.seatLimit) * 100),
      },
      billing: {
        accountStatus: billing.accountStatus,
        nextInvoiceDate: billing.nextInvoiceDate,
        trialEndsAt: billing.trialEndsAt,
        suspendedAt: billing.suspendedAt,
        suspensionReason: billing.suspensionReason,
      },
      alerts: {
        overdueCount,
        unpaidTotal,
        overageWarning: overage > 0,
        nearLimit: activeStudents / billing.seatLimit > 0.9,
      },
      recentInvoices,
      billingContact: billing.billingContact,
      planCatalogue: PLAN_CATALOGUE,
    });
  }),
);

// ── GET /billing/invoices ─────────────────────────────────────────────────────
// Paginated invoice list with status filter.
router.get(
  "/invoices",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const {
      status,
      page = "1",
      limit = "12",
    } = req.query as Record<string, string>;

    const billing = await Billing.findOne({ institution: institutionId })
      .select("invoices currency planName")
      .lean();

    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    let invoices = [...(billing.invoices ?? [])].sort(
      (a, b) =>
        new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
    );

    if (status) invoices = invoices.filter((i) => i.status === status);

    const total = invoices.length;
    const pageNum = parseInt(page);
    const limitNum = parseInt(limit);
    const paginated = invoices.slice(
      (pageNum - 1) * limitNum,
      pageNum * limitNum,
    );

    res.json({
      total,
      page: pageNum,
      invoices: paginated,
      currency: billing.currency,
    });
  }),
);

// ── GET /billing/invoices/:id ─────────────────────────────────────────────────
router.get(
  "/invoices/:id",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { id } = req.params;

    const billing = await Billing.findOne({
      institution: institutionId,
    }).lean();
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    const invoice = (billing.invoices ?? []).find((i) => i.id === id);
    if (!invoice) {
      res.status(404).json({ message: "Invoice not found." });
      return;
    }

    res.json({
      invoice,
      currency: billing.currency,
      institution: billing.institution,
    });
  }),
);

// ── POST /billing/invoices/generate ──────────────────────────────────────────
// Admin manually triggers invoice generation (in addition to the monthly cron).
router.post(
  "/invoices/generate",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const invoice = await generateInvoice(institutionId);
    if (!invoice) {
      res
        .status(400)
        .json({
          message: "Invoice could not be generated. Check account status.",
        });
      return;
    }

    await logAudit(req, {
      action: "billing_invoice_manually_generated",
      actor: req.user._id,
      details: {
        invoiceId: invoice.id,
        invoiceNumber: invoice.invoiceNumber,
        total: invoice.total,
      },
    });

    res.status(201).json({ message: "Invoice generated.", invoice });
  }),
);

// ── PATCH /billing/invoices/:id/pay ──────────────────────────────────────────
// Record a manual payment — no gateway, just an admin marking it paid with ref.
router.patch(
  "/invoices/:id/pay",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { id } = req.params;
    const { paidAmount, paymentRef, paymentMethod, notes } = req.body;

    if (!paidAmount || !paymentRef || !paymentMethod) {
      res
        .status(400)
        .json({
          message: "paidAmount, paymentRef, and paymentMethod are required.",
        });
      return;
    }

    const result = await recordPayment({
      institutionId,
      invoiceId: id,
      paidAmount: Number(paidAmount),
      paymentRef,
      paymentMethod,
      notes,
    });

    if (!result.ok) {
      res.status(400).json({ message: result.message });
      return;
    }

    await logAudit(req, {
      action: "billing_payment_recorded",
      actor: req.user._id,
      details: { invoiceId: id, paidAmount, paymentRef, paymentMethod },
    });

    res.json({ message: result.message });
  }),
);

// ── PATCH /billing/invoices/:id/void ─────────────────────────────────────────
router.patch(
  "/invoices/:id/void",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { id } = req.params;
    const { reason } = req.body;

    if (!reason?.trim()) {
      res
        .status(400)
        .json({ message: "A reason is required to void an invoice." });
      return;
    }

    const result = await voidInvoice(institutionId, id, reason);
    if (!result.ok) {
      res.status(400).json({ message: result.message });
      return;
    }

    await logAudit(req, {
      action: "billing_invoice_voided",
      actor: req.user._id,
      details: { invoiceId: id, reason },
    });
    res.json({ message: result.message });
  }),
);

// ── GET /billing/usage ────────────────────────────────────────────────────────
// Usage history for charts — seat count over time.
router.get(
  "/usage",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const billing = await Billing.findOne({ institution: institutionId })
      .select("usageHistory seatLimit currency planName")
      .lean();

    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    // Latest snapshot on demand (not stored)
    const latest = await takeUsageSnapshot(institutionId, billing.seatLimit);

    res.json({
      latest,
      history: (billing.usageHistory ?? [])
        .sort(
          (a, b) =>
            new Date(b.snapshotDate).getTime() -
            new Date(a.snapshotDate).getTime(),
        )
        .slice(0, 24), // Last 24 months
    });
  }),
);

// ── PATCH /billing/plan ───────────────────────────────────────────────────────
router.patch(
  "/plan",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const {
      newPlanName,
      reason,
      customSeatLimit,
      customBasePrice,
      customOverageRate,
    } = req.body;

    if (!newPlanName?.trim()) {
      res.status(400).json({ message: "newPlanName is required." });
      return;
    }

    const result = await changePlan({
      institutionId,
      newPlanName,
      changedBy: req.user._id.toString(),
      reason,
      customSeatLimit: customSeatLimit ? Number(customSeatLimit) : undefined,
      customBasePrice: customBasePrice ? Number(customBasePrice) : undefined,
      customOverageRate: customOverageRate
        ? Number(customOverageRate)
        : undefined,
    });

    if (!result.ok) {
      res.status(400).json({ message: result.message });
      return;
    }

    await logAudit(req, {
      action: "billing_plan_changed",
      actor: req.user._id,
      details: { newPlanName, reason },
    });
    res.json({ message: result.message });
  }),
);

// ── PATCH /billing/contact ────────────────────────────────────────────────────
router.patch(
  "/contact",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { name, email, phone, address } = req.body;

    if (!name?.trim() || !email?.trim()) {
      res.status(400).json({ message: "name and email are required." });
      return;
    }

    await Billing.findOneAndUpdate(
      { institution: institutionId },
      { $set: { billingContact: { name, email, phone, address } } },
    );

    await logAudit(req, {
      action: "billing_contact_updated",
      actor: req.user._id,
      details: { email },
    });
    res.json({ message: "Billing contact updated." });
  }),
);

// ── PATCH /billing/cycle ──────────────────────────────────────────────────────
router.patch(
  "/cycle",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { cycle } = req.body;

    if (!["monthly", "annual"].includes(cycle)) {
      res.status(400).json({ message: "cycle must be 'monthly' or 'annual'." });
      return;
    }

    await Billing.findOneAndUpdate(
      { institution: institutionId },
      { $set: { billingCycle: cycle } },
    );
    await logAudit(req, {
      action: "billing_cycle_changed",
      actor: req.user._id,
      details: { cycle },
    });
    res.json({ message: `Billing cycle changed to ${cycle}.` });
  }),
);

// ── GET /billing/plan-history ─────────────────────────────────────────────────
router.get(
  "/plan-history",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const billing = await Billing.findOne({ institution: institutionId })
      .select("planHistory")
      .populate("planHistory.changedBy", "name email")
      .lean();
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }
    res.json({ planHistory: (billing as any).planHistory ?? [] });
  }),
);

// ── CRON endpoints — protected by a system secret, not admin role ─────────────
// These are called by a cron job or Kubernetes CronJob, not by the browser.
// They use the CRON_SECRET header for authentication.

const verifyCronSecret = (req: AuthenticatedRequest) =>
  req.headers["x-cron-secret"] === process.env.CRON_SECRET;

router.post(
  "/cron/invoices",
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!verifyCronSecret(req)) {
      res.status(403).json({ message: "Forbidden." });
      return;
    }
    const result = await generateMonthlyInvoices();
    res.json({ message: "Monthly invoice generation complete.", ...result });
  }),
);

router.post(
  "/cron/overdue",
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!verifyCronSecret(req)) {
      res.status(403).json({ message: "Forbidden." });
      return;
    }
    const count = await markOverdueInvoices();
    res.json({ message: `Marked ${count} invoice(s) as overdue.` });
  }),
);

export default router;
