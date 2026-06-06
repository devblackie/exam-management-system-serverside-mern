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
import mongoose from "mongoose";
import { requireAuth, requireRole, AuthenticatedRequest } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import { logAudit } from "../lib/auditLogger";
import { cached, invalidateCache } from "../utils/cache";
import Billing from "../models/Billing";
import Student from "../models/Student";
import InstitutionSettings from "../models/InstitutionSettings";
import EmailLog from "../models/EmailLog";
import {
  generateInvoice, recordPayment, voidInvoice, changePlan, markOverdueInvoices,
  generateMonthlyInvoices, takeUsageSnapshot, PLAN_CATALOGUE,
  sendInvoiceEmail, 
} from "../services/billingService";

const router = Router();
router.use(requireAuth);

// ── GET /billing/summary ──────────────────────────────────────────────────────
// Dashboard overview card. Shows plan, seat usage, next invoice, and last 3 invoices.
router.get("/summary", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();

    const billing = await Billing.findOne({
      institution: institutionId,
    }).lean();
    if (!billing) {
      res.status(404).json({
          message:
            "No billing record found. Contact support to set up billing.",
        });
      return;
    }

    const [activeStudents, totalStudents] = await Promise.all([
      Student.countDocuments({ institution: institutionId, status: { $in: ["active", "repeat"] } }),
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
          includedSeats: billing.seatLimit, // NEW – clear name
          basePrice: billing.basePrice,
          perSeatRate: billing.overageRate, // NEW – clear name
          currency: billing.currency,
          taxRate: billing.taxRate,
          isCustomPlan: billing.isCustomPlan,
        },
        usage: {
          activeStudents,
          totalStudents,
          includedSeats: billing.seatLimit,
          extraSeats: overage, // renamed from overage
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
        planCatalogue: PLAN_CATALOGUE.map((p) => ({
          name: p.name,
          includedSeats: p.includedSeats,
          monthlyKES: p.monthlyKES,
          perSeatRate: p.perSeatRate,
        })),
      });
  }),
);

// ── GET /billing/invoices ─────────────────────────────────────────────────────
// Paginated invoice list with status filter.
router.get("/invoices", requireRole("admin"),
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
router.get("/invoices/:id", requireRole("admin"),
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
router.post("/invoices/generate", requireRole("admin"),
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
router.patch("/invoices/:id/pay", requireRole("admin"),
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

// PATCH /billing/invoices/bulk-pay
router.patch(
  "/invoices/bulk-pay",
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { invoiceIds, paidAmount, paymentRef, paymentMethod, notes } = req.body;

    if (!invoiceIds?.length || !paidAmount || !paymentRef || !paymentMethod) {
      res.status(400).json({ message: "invoiceIds, paidAmount, paymentRef, paymentMethod required." });
      return;
    }

    const billing = await Billing.findOne({ institution: institutionId });
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    let updatedCount = 0;
    for (const invoiceId of invoiceIds) {
      const invoice = billing.invoices.find(i => i.id === invoiceId);
      if (!invoice || invoice.status === "void" || invoice.status === "paid") continue;
      invoice.paidAmount = Number(paidAmount);
      invoice.paidAt = new Date();
      invoice.paymentRef = paymentRef;
      invoice.paymentMethod = paymentMethod;
      invoice.status = Number(paidAmount) >= invoice.total ? "paid" : invoice.status;
      if (notes) invoice.notes = (invoice.notes ? invoice.notes + " | " : "") + notes;
      updatedCount++;
    }

    await billing.save();

    await logAudit(req, {
      action: "billing_bulk_payment_recorded",
      actor: req.user._id,
      details: { invoiceIds, paidAmount, paymentRef, paymentMethod, count: updatedCount },
    });

    res.json({ message: `${updatedCount} invoice(s) updated.` });
  }),
);

// ── PATCH /billing/invoices/:id/void ─────────────────────────────────────────
router.patch("/invoices/:id/void", requireRole("admin"),
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
router.get("/usage", requireRole("admin"),
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
router.patch("/plan", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const {
      newPlanName,
      reason,
      customSeatLimit,
      customBasePrice,
      customPerSeatRate,          // ← new name for the extra seat rate
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
      customPerSeatRate: customPerSeatRate   // pass through to the service
        ? Number(customPerSeatRate)
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
router.patch("/contact", requireRole("admin"),
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
router.patch("/cycle", requireRole("admin"),
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
router.get("/plan-history", requireRole("admin"),
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

router.post("/cron/invoices",
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!verifyCronSecret(req)) {
      res.status(403).json({ message: "Forbidden." });
      return;
    }
    const result = await generateMonthlyInvoices();
    res.json({ message: "Monthly invoice generation complete.", ...result });
  }),
);

router.post("/cron/overdue",
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    if (!verifyCronSecret(req)) {
      res.status(403).json({ message: "Forbidden." });
      return;
    }
    const count = await markOverdueInvoices();
    res.json({ message: `Marked ${count} invoice(s) as overdue.` });
  }),
);

// ── GET /billing/department-stats ────────────────────────────────────────
// Aggregates active student counts per department, with institution context.
router.get("/department-stats", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();

    // 1. Get institution billing info for seat limits & rate
    const billing = await Billing.findOne({ institution: institutionId })
      .select("seatLimit overageRate currency")
      .lean();
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    // 2. Aggregate active students per department using Student -> Program
    const deptStats = await Student.aggregate([
      {
        $match: {
          institution: new mongoose.Types.ObjectId(institutionId),
          status: { $in: ["active", "repeat"] },
        },
      },
      {
        $lookup: {
          from: "programs",                 // the Program collection name
          localField: "program",
          foreignField: "_id",
          as: "programDoc",
        },
      },
      { $unwind: "$programDoc" },
      {
        $group: {
          _id: "$programDoc.departmentCode",
          activeStudents: { $sum: 1 },
        },
      },
      {
        $lookup: {
          from: "institutionsettings",
          let: { deptCode: "$_id" },
          pipeline: [
            { $match: { institution: new mongoose.Types.ObjectId(institutionId) } },
            { $unwind: "$schools" },
            { $unwind: "$schools.departments" },
            {
              $match: {
                $expr: {
                  $eq: ["$schools.departments.code", "$$deptCode"],
                },
              },
            },
            {
              $project: {
                _id: 0,
                departmentName: "$schools.departments.name",
                schoolName: "$schools.name",
              },
            },
          ],
          as: "deptInfo",
        },
      },
      {
        $addFields: {
          name: {
            $ifNull: [
              { $arrayElemAt: ["$deptInfo.departmentName", 0] },
              "$_id",
            ],
          },
          schoolName: { $arrayElemAt: ["$deptInfo.schoolName", 0] },
        },
      },
      {
        $project: {
          _id: 0,
          departmentCode: "$_id",
          name: 1,
          schoolName: 1,
          activeStudents: 1,
        },
      },
      { $sort: { activeStudents: -1 } },
    ]);

    // 3. Compute overage per department (if institution seats are divided equally? 
    //    Better: just show count; overage is institution-level.
    //    We can include a simple comparison if the institution has per-department limits.
    //    Here we assume institution-wide limit only, so departments don't have individual limits.
    const departments = deptStats.map((d: any) => ({
      ...d,
      includedSeats: null,          // per-department limit not set
      overage: null,
      extraSeatCost: null,
    }));

    res.json({
      departments,
      institutionIncludedSeats: billing.seatLimit,
      institutionPerSeatRate: billing.overageRate,
    });
  }),
);

// GET /billing/hierarchy
// ═══════════════════════════════════════════════════════════════════════
// NEW: GET /billing/hierarchy (cached)
// ═══════════════════════════════════════════════════════════════════════
// GET /billing/hierarchy (cached)
router.get("/hierarchy", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const cacheKey = `billing:hierarchy:${institutionId}`;

    const result = await cached(
      cacheKey,
      async () => {
        const billing = await Billing.findOne({ institution: institutionId })
          .select("seatLimit overageRate currency departmentSeats")
          .lean();
        if (!billing) return null;

        const settings = await InstitutionSettings.findOne({ institution: institutionId })
          .select("schools")
          .lean() as any;

        const programStats = await Student.aggregate([
          {
            $match: {
              institution: new mongoose.Types.ObjectId(institutionId),
              status: { $in: ["active", "repeat"] },
            },
          },
          {
            $lookup: {
              from: "programs",
              localField: "program",
              foreignField: "_id",
              as: "prog",
            },
          },
          { $unwind: "$prog" },
          // Only keep programs that actually have school & department codes
          {
            $match: {
              "prog.departmentCode": { $exists: true, $nin: [null, ""] },
              "prog.schoolCode": { $exists: true, $nin: [null, ""] },
            },
          },
          {
            $group: {
              _id: {
                departmentCode: "$prog.departmentCode",
                schoolCode: "$prog.schoolCode",
                programId: "$prog._id",
                programName: "$prog.name",
              },
              count: { $sum: 1 },
            },
          },
        ]);

        // Department seat limit map
        const deptSeatMap = new Map<string, number>();
        (billing.departmentSeats ?? []).forEach(ds => {
          deptSeatMap.set(ds.departmentCode.toUpperCase(), ds.seatLimit);
        });

        // Resolve school/department names from InstitutionSettings
        const schoolNameMap = new Map<string, string>();
        const deptNameMap   = new Map<string, string>();
        settings?.schools?.forEach((school: any) => {
          schoolNameMap.set(school.code.toUpperCase(), school.name);
          school.departments?.forEach((dept: any) => {
            deptNameMap.set(
              `${school.code.toUpperCase()}|${dept.code.toUpperCase()}`,
              dept.name,
            );
          });
        });

        // Build hierarchy map
        const schoolMap = new Map<string, {
          schoolName: string; schoolCode: string; schoolTotal: number;
          departments: Map<string, {
            deptName: string; deptCode: string; deptTotal: number;
            deptSeatLimit: number | null;
            programs: Array<{ programName: string; programId: string; activeStudents: number }>;
          }>;
        }>();

        for (const item of programStats) {
          const { schoolCode, departmentCode, programId, programName } = item._id;
          const count = item.count;

          const schoolKey = (schoolCode ?? "").toUpperCase();
          const deptKey   = (departmentCode ?? "").toUpperCase();
          if (!schoolKey || !deptKey) continue;

          if (!schoolMap.has(schoolKey)) {
            schoolMap.set(schoolKey, {
              schoolName: schoolNameMap.get(schoolKey) ?? schoolKey,
              schoolCode: schoolKey,
              schoolTotal: 0,
              departments: new Map(),
            });
          }

          const school = schoolMap.get(schoolKey)!;
          school.schoolTotal += count;

          if (!school.departments.has(deptKey)) {
            school.departments.set(deptKey, {
              deptName: deptNameMap.get(`${schoolKey}|${deptKey}`) ?? deptKey,
              deptCode: deptKey,
              deptTotal: 0,
              deptSeatLimit: deptSeatMap.has(deptKey) ? deptSeatMap.get(deptKey)! : null,
              programs: [],
            });
          }

          const dept = school.departments.get(deptKey)!;
          dept.deptTotal += count;
          dept.programs.push({
            programName,
            programId: programId.toString(),
            activeStudents: count,
          });
        }

        // Convert to plain objects
        const schools = Array.from(schoolMap.values())
          .map(school => ({
            schoolName: school.schoolName,
            schoolCode: school.schoolCode,
            totalStudents: school.schoolTotal,
            departments: Array.from(school.departments.values())
              .map(dept => ({
                deptName: dept.deptName,
                deptCode: dept.deptCode,
                totalStudents: dept.deptTotal,
                seatLimit: dept.deptSeatLimit,
                overage: dept.deptSeatLimit
                  ? Math.max(0, dept.deptTotal - dept.deptSeatLimit)
                  : 0,
                programs: dept.programs.sort((a, b) =>
                  a.programName.localeCompare(b.programName),
                ),
              }))
              .sort((a, b) => a.deptName.localeCompare(b.deptName)),
          }))
          .sort((a, b) => a.schoolName.localeCompare(b.schoolName));

        return {
          schools,
          institutionCurrency: billing.currency,
          institutionOverageRate: billing.overageRate,
        };
      },
      300, // 5 min TTL
    );

    if (!result) {
      res.status(404).json({ message: "No billing record." });
      return;
    }
    res.json(result);
  }),
);

// PATCH /billing/department-seats
router.patch("/department-seats", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { departmentCode, seatLimit } = req.body;

    if (!departmentCode || seatLimit == null || seatLimit < 1) {
      res.status(400).json({ message: "Valid departmentCode and seatLimit required." });
      return;
    }

    const billing = await Billing.findOne({ institution: institutionId });
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    const existing = billing.departmentSeats.find(
      ds => ds.departmentCode === departmentCode.toUpperCase(),
    );
    if (existing) {
      existing.seatLimit = seatLimit;
    } else {
      billing.departmentSeats.push({
        departmentCode: departmentCode.toUpperCase(),
        seatLimit,
      });
    }

    await billing.save();
    invalidateCache(`billing:hierarchy:${institutionId}`);

    await logAudit(req, {
      action: "billing_department_seat_limit_set",
      actor: req.user._id,
      details: { departmentCode, seatLimit },
    });

    res.json({ message: `Seat limit for ${departmentCode} updated to ${seatLimit}.` });
  }),
);

// POST /billing/invoices/export – export current filtered invoices
router.post("/invoices/export", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const { status, invoiceIds } = req.body as { status?: string; invoiceIds?: string[] };

    const billing = await Billing.findOne({ institution: institutionId }).lean();
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }

    let invoices = billing.invoices ?? [];
    if (invoiceIds && invoiceIds.length > 0) {
      invoices = invoices.filter(inv => invoiceIds.includes(inv.id));
    } else if (status) {
      invoices = invoices.filter(inv => inv.status === status);
    }

    const ExcelJS = require("exceljs");
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Invoices");

    sheet.columns = [
      { header: "Invoice Number", key: "invoiceNumber", width: 20 },
      { header: "Label", key: "label", width: 30 },
      { header: "Period Start", key: "periodStart", width: 15 },
      { header: "Period End", key: "periodEnd", width: 15 },
      { header: "Total", key: "total", width: 12 },
      { header: "Currency", key: "currency", width: 10 },
      { header: "Status", key: "status", width: 12 },
      { header: "Due Date", key: "dueAt", width: 15 },
      { header: "Paid At", key: "paidAt", width: 15 },
      { header: "Payment Ref", key: "paymentRef", width: 20 },
    ];

    invoices.forEach(inv => {
      sheet.addRow({
        invoiceNumber: inv.invoiceNumber,
        label: inv.label,
        periodStart: new Date(inv.periodStart).toLocaleDateString("en-KE"),
        periodEnd: new Date(inv.periodEnd).toLocaleDateString("en-KE"),
        total: inv.total,
        currency: inv.currency,
        status: inv.status,
        dueAt: inv.dueAt ? new Date(inv.dueAt).toLocaleDateString("en-KE") : "",
        paidAt: inv.paidAt ? new Date(inv.paidAt).toLocaleDateString("en-KE") : "",
        paymentRef: inv.paymentRef ?? "",
      });
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=invoices.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  }),
);

// GET /billing/email-logs
router.get("/email-logs", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const logs = await EmailLog.find({ institution: institutionId })
      .sort({ timestamp: -1 })
      .limit(10)
      .lean();
    res.json(logs);
  }),
);

// POST /billing/invoices/:id/resend-email
router.post("/invoices/:id/resend-email", requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution.toString();
    const invoiceId = req.params.id;

    const billing = await Billing.findOne({ institution: institutionId }).lean();
    if (!billing) {
      res.status(404).json({ message: "No billing record." });
      return;
    }
    const invoice = (billing.invoices ?? []).find(inv => inv.id === invoiceId);
    if (!invoice) {
      res.status(404).json({ message: "Invoice not found." });
      return;
    }
    if (!billing.billingContact?.email) {
      res.status(400).json({ message: "No billing contact email set." });
      return;
    }

    // Reuse the send function
    await sendInvoiceEmail(billing, invoice);
    res.json({ message: "Invoice email resent." });
  }),
);



export default router;
