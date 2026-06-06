// serverside/src/services/billingService.ts

import crypto from "node:crypto";
import mongoose from "mongoose";
import Billing, {
  IBilling,
  IInvoice,
  IInvoiceLine,
  IUsageSnapshot,
} from "../models/Billing";
import Student from "../models/Student";
import Institution from "../models/Institution";
import EmailLog from "../models/EmailLog";
import { sendEmail } from "../lib/mailer"; // <-- this is the imported sendEmail

// ── Plan catalogue (band + per‑seat) ──────────────────────────────────────────
export const PLAN_CATALOGUE = [
  { name: "Starter", includedSeats: 300, monthlyKES: 12_500, perSeatRate: 25 },
  { name: "Growth", includedSeats: 800, monthlyKES: 22_500, perSeatRate: 25 },
  { name: "Pro", includedSeats: 1_500, monthlyKES: 35_000, perSeatRate: 20 },
  { name: "Enterprise", includedSeats: 99_999, monthlyKES: 0, perSeatRate: 0 },
] as const;

export type PlanName = (typeof PLAN_CATALOGUE)[number]["name"];
export const ANNUAL_DISCOUNT = 0.15;

export function suggestPlan(seats: number) {
  return (
    PLAN_CATALOGUE.find((p) => seats <= p.includedSeats) ??
    PLAN_CATALOGUE[PLAN_CATALOGUE.length - 1]
  );
}

function buildInvoiceNumber(counter: number): string {
  return `INV-${new Date().getFullYear()}-${String(counter).padStart(4, "0")}`;
}

export async function takeUsageSnapshot(
  institutionId: string,
  seatLimit: number,
): Promise<IUsageSnapshot> {
  const [activeStudents, totalStudents] = await Promise.all([
    Student.countDocuments({
      institution: institutionId,
      status: { $in: ["active", "repeat"] },
    }),
    Student.countDocuments({ institution: institutionId }),
  ]);
  return {
    snapshotDate: new Date(),
    activeStudents,
    totalStudents,
    seatLimit,
    overage: Math.max(0, activeStudents - seatLimit),
  };
}

export function buildInvoiceLines(params: {
  billing: IBilling;
  snapshot: IUsageSnapshot;
  periodStart: Date;
  periodEnd: Date;
}): { lines: IInvoiceLine[]; subtotal: number } {
  const { billing, snapshot } = params;
  const isAnnual = billing.billingCycle === "annual";
  const lines: IInvoiceLine[] = [];
  const includedSeats = billing.seatLimit;
  const perSeatRate = billing.overageRate;

  const baseMonthly = billing.basePrice;
  const baseQty = isAnnual ? 12 : 1;
  const baseUnit = isAnnual ? baseMonthly * (1 - ANNUAL_DISCOUNT) : baseMonthly;
  const baseTotal = Math.round(baseUnit * baseQty);

  lines.push({
    description: `${billing.planName} Plan — ${isAnnual ? "Annual" : "Monthly"} Subscription (${includedSeats} seats included)`,
    quantity: baseQty,
    unitPrice: Math.round(baseUnit),
    total: baseTotal,
  });

  const extraSeats = Math.max(0, snapshot.activeStudents - includedSeats);
  if (!isAnnual && extraSeats > 0 && perSeatRate > 0) {
    lines.push({
      description: `Additional seats (${extraSeats} × ${billing.currency} ${perSeatRate}/seat)`,
      quantity: extraSeats,
      unitPrice: perSeatRate,
      total: extraSeats * perSeatRate,
    });
  }
  const subtotal = lines.reduce((sum, l) => sum + l.total, 0);
  return { lines, subtotal };
}

export async function generateInvoice(
  institutionId: string,
): Promise<IInvoice | null> {
  const billing = await Billing.findOne({ institution: institutionId });
  if (!billing) return null;
  if (billing.accountStatus !== "active") {
    console.warn(
      `[Billing] Skipping invoice for ${institutionId}: accountStatus = ${billing.accountStatus}`,
    );
    return null;
  }

  const now = new Date();
  const periodStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const periodEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const dueAt = new Date(now.getFullYear(), now.getMonth() + 1, 5);

  const snapshot = await takeUsageSnapshot(institutionId, billing.seatLimit);
  const { lines, subtotal } = buildInvoiceLines({
    billing,
    snapshot,
    periodStart,
    periodEnd,
  });
  const tax = Math.round(subtotal * billing.taxRate);
  const total = subtotal + tax;

  const updatedBilling = await Billing.findByIdAndUpdate(
    billing._id,
    { $inc: { invoiceCounter: 1 } },
    { new: true, select: "invoiceCounter" },
  ).lean();
  const counter = updatedBilling?.invoiceCounter ?? billing.invoiceCounter + 1;

  const invoice: IInvoice = {
    id: crypto.randomUUID(),
    invoiceNumber: buildInvoiceNumber(counter),
    label: `${billing.planName} Plan — ${now.toLocaleString("en-KE", { month: "long", year: "numeric" })}`,
    periodStart,
    periodEnd,
    lines,
    subtotal,
    tax,
    total,
    currency: billing.currency,
    status: "sent",
    dueAt,
    createdAt: now,
  };

  await Billing.findByIdAndUpdate(billing._id, {
    $push: {
      invoices: invoice,
      usageHistory: snapshot,
    },
    $set: {
      nextInvoiceDate: new Date(now.getFullYear(), now.getMonth() + 1, 1),
    },
  });

  sendInvoiceEmail(billing, invoice).catch((err) =>
    console.error(
      `[Billing] Invoice email failed for ${institutionId}:`,
      err.message,
    ),
  );
  return invoice;
}

// ── Record a manual payment ───────────────────────────────────────────────────
export async function recordPayment(params: {
  institutionId: string; invoiceId: string; paidAmount: number;
  paymentRef: string; paymentMethod: string; notes?: string;
}): Promise<{ ok: boolean; message: string }> {
  const {
    institutionId, invoiceId, paidAmount,
    paymentRef, paymentMethod, notes,
  } = params;

  const billing = await Billing.findOne({ institution: institutionId });
  if (!billing) return { ok: false, message: "No billing record found." };

  const invoice = billing.invoices.find((i) => i.id === invoiceId);
  if (!invoice) return { ok: false, message: "Invoice not found." };
  if (invoice.status === "void")
    return { ok: false, message: "Cannot mark a voided invoice as paid." };

  invoice.paidAmount = paidAmount;
  invoice.paidAt = new Date();
  invoice.paymentRef = paymentRef;
  invoice.paymentMethod = paymentMethod;
  invoice.status = paidAmount >= invoice.total ? "paid" : invoice.status;
  if (notes)
    invoice.notes = (invoice.notes ? invoice.notes + " | " : "") + notes;

  await billing.save();
  return {
    ok: true,
    message: `Payment of ${billing.currency} ${paidAmount.toLocaleString()} recorded.`,
  };
}

// ── Void an invoice ───────────────────────────────────────────────────────────
export async function voidInvoice(
  institutionId: string,
  invoiceId: string,
  reason: string,
): Promise<{ ok: boolean; message: string }> {
  const billing = await Billing.findOne({ institution: institutionId });
  if (!billing) return { ok: false, message: "Billing record not found." };

  const invoice = billing.invoices.find((i) => i.id === invoiceId);
  if (!invoice) return { ok: false, message: "Invoice not found." };
  if (invoice.status === "paid")
    return { ok: false, message: "Cannot void a paid invoice." };

  invoice.status = "void";
  invoice.notes =
    `VOID: ${reason}` + (invoice.notes ? ` | ${invoice.notes}` : "");
  await billing.save();
  return { ok: true, message: "Invoice voided." };
}

// ── Change plan ───────────────────────────────────────────────────────────────
export async function changePlan(params: {
  institutionId: string;
  newPlanName: string;
  changedBy: string;
  reason?: string;
  customSeatLimit?: number; // custom included seats
  customBasePrice?: number;
  customPerSeatRate?: number; // custom per‑seat rate
}): Promise<{ ok: boolean; message: string }> {
  const {
    institutionId,
    newPlanName,
    changedBy,
    reason,
    customSeatLimit,
    customBasePrice,
    customPerSeatRate,
  } = params;

  const billing = await Billing.findOne({ institution: institutionId });
  if (!billing) return { ok: false, message: "Billing record not found." };

  const fromPlan = billing.planName;
  billing.planHistory.push({
    date: new Date(),
    fromPlan,
    toPlan: newPlanName,
    changedBy: new mongoose.Types.ObjectId(changedBy),
    reason,
  });

  const catalogue = PLAN_CATALOGUE.find((p) => p.name === newPlanName);
  if (catalogue) {
    billing.planName = catalogue.name;
    billing.seatLimit = catalogue.includedSeats;
    billing.basePrice = catalogue.monthlyKES;
    billing.overageRate = catalogue.perSeatRate;
    billing.isCustomPlan = false;
  } else {
    billing.planName = newPlanName;
    billing.isCustomPlan = true;
    if (customSeatLimit !== undefined) billing.seatLimit = customSeatLimit;
    if (customBasePrice !== undefined) billing.basePrice = customBasePrice;
    if (customPerSeatRate !== undefined)
      billing.overageRate = customPerSeatRate;
  }

  await billing.save();
  return {
    ok: true,
    message: `Plan changed from ${fromPlan} to ${newPlanName}.`,
  };
}

// ── Mark overdue invoices ─────────────────────────────────────────────────────
export async function markOverdueInvoices(): Promise<number> {
  const now = new Date();
  const billings = await Billing.find({ "invoices.status": "sent" });
  let count = 0;

  for (const billing of billings) {
    let changed = false;
    for (const inv of billing.invoices) {
      if (inv.status === "sent" && inv.dueAt < now) {
        inv.status = "overdue";
        changed = true;
        count++;
      }
    }
    if (changed) await billing.save();
  }

  return count;
}

// ── Monthly invoice generation for ALL institutions ───────────────────────────
export async function generateMonthlyInvoices(): Promise<{
  generated: number;
  skipped: number;
  errors: string[];
}> {
  const billings = await Billing.find({ accountStatus: "active" })
    .select("institution")
    .lean();
  let generated = 0;
  let skipped = 0;
  const errors: string[] = [];

  for (const b of billings) {
    try {
      const invoice = await generateInvoice(b.institution.toString());
      if (invoice) generated++;
      else skipped++;
    } catch (err: any) {
      errors.push(`${b.institution}: ${err.message}`);
      console.error(`[Billing] Failed for ${b.institution}:`, err.message);
    }
  }

  return { generated, skipped, errors };
}

// ── Invoice email (unchanged logic) ───────────────────────────────────────────
export async function sendInvoiceEmail(billing: IBilling, invoice: IInvoice): Promise<void> {
  const institution = (await Institution.findById(
    billing.institution,
  ).lean()) as any;
  const contact = billing.billingContact;
  if (!contact?.email) return;

  const dueDateStr = new Date(invoice.dueAt).toLocaleDateString("en-KE", {
    day: "2-digit", month: "long", year: "numeric",
  });

  const linesHtml = invoice.lines
    .map(
      (l) => `
    <tr>
      <td style="padding:8px 12px;font-size:12px;">${l.description}</td>
      <td style="padding:8px 12px;font-size:12px;text-align:center;">${l.quantity}</td>
      <td style="padding:8px 12px;font-size:12px;text-align:right;">${invoice.currency} ${l.unitPrice.toLocaleString()}</td>
      <td style="padding:8px 12px;font-size:12px;text-align:right;font-weight:bold;">${invoice.currency} ${l.total.toLocaleString()}</td>
    </tr>
  `,
    )
    .join("");

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:#002B1B;padding:24px;border-radius:8px 8px 0 0;">
        <h1 style="color:#EAB308;font-size:18px;margin:0;">Exam Management System</h1>
        <p style="color:rgba(255,255,255,0.5);font-size:11px;margin:4px 0 0;">${institution?.name ?? "Your Institution"}</p>
      </div>
      <div style="background:#F8F9FA;padding:32px;border-radius:0 0 8px 8px;border:1px solid #e2e8f0;border-top:none;">
        <p style="font-size:14px;color:#374151;">Dear ${contact.name},</p>
        <p style="font-size:13px;color:#374151;">
          Your invoice <strong>${invoice.invoiceNumber}</strong> for the period
          ${new Date(invoice.periodStart).toLocaleDateString("en-KE", { month: "short", year: "numeric" })} is ready.
        </p>
        <table style="width:100%;border-collapse:collapse;margin:20px 0;background:white;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">
          <thead>
            <tr style="background:#002B1B;color:white;">
              <th style="padding:10px 12px;font-size:11px;text-align:left;">Description</th>
              <th style="padding:10px 12px;font-size:11px;text-align:center;">Qty</th>
              <th style="padding:10px 12px;font-size:11px;text-align:right;">Unit Price</th>
              <th style="padding:10px 12px;font-size:11px;text-align:right;">Total</th>
            </tr>
          </thead>
          <tbody>${linesHtml}</tbody>
          <tfoot>
            <tr style="border-top:2px solid #002B1B;">
              <td colspan="3" style="padding:10px 12px;font-size:13px;font-weight:bold;text-align:right;">Total Due</td>
              <td style="padding:10px 12px;font-size:15px;font-weight:bold;color:#002B1B;text-align:right;">
                ${invoice.currency} ${invoice.total.toLocaleString()}
              </td>
            </tr>
          </tfoot>
        </table>
        <div style="background:#FEF9EE;border:1px solid #FDE68A;border-radius:8px;padding:16px;margin-bottom:20px;">
          <p style="font-size:12px;color:#92400E;margin:0;line-height:1.6;">
            <strong>Due date:</strong> ${dueDateStr}<br>
            <strong>Reference:</strong> ${invoice.invoiceNumber}<br>
            Please include the invoice number in your payment reference.
          </p>
        </div>
        <p style="font-size:11px;color:#9ca3af;">
          To view your full billing history, log in to the EMS admin dashboard and navigate to Billing.
        </p>
      </div>
    </div>`;

  try {
    await sendEmail({
      to: contact.email,
      subject: `${invoice.invoiceNumber} — ${invoice.currency} ${invoice.total.toLocaleString()} due ${dueDateStr}`,
      html,
    });
    await EmailLog.create({
      institution: billing.institution,
      invoiceNumber: invoice.invoiceNumber,
      recipient: contact.email,
      subject: `${invoice.invoiceNumber} — ${invoice.currency} ${invoice.total.toLocaleString()} due ${dueDateStr}`,
      status: "sent",
    });
  } catch (err: any) {
    console.error("[Billing] Invoice email failed:", err.message);
    await EmailLog.create({
      institution: billing.institution,
      invoiceNumber: invoice.invoiceNumber,
      recipient: contact.email,
      subject: `${invoice.invoiceNumber} — ...`,
      status: "failed",
      errorMessage: err.message,
    });
  }
}
