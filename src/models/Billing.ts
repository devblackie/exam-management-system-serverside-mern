// serverside/src/models/Billing.ts
//
// WHAT CHANGED FROM THE ORIGINAL
// ──────────────────────────────────────────────────────────────────────────
// The original Billing model was a stub — it stored a planName, a seatLimit,
// and an invoices array. No plan tiers, no overage tracking, no usage history,
// no plan-change log, no billing contact. It was a read-only document that
// nobody ever wrote to.
//
// This replaces it with a model that:
//   1. Stores configurable plan tiers (not hardcoded in code)
//   2. Tracks monthly active seat counts in usageHistory
//   3. Records plan changes with who made them and when
//   4. Supports both monthly and annual billing cycles
//   5. Supports per-institution custom pricing (Enterprise)
//   6. Stores billing contact separately from the admin user
//   7. Tracks invoice line items (base + overage as separate lines)
//   8. Records manual payments (bank transfer, cheque, cash) with ref numbers
//
// NOT INCLUDED (intentionally): Payment gateway fields. No Stripe customer ID,
// no M-Pesa merchant data. This is a clean billing record. Gateway data lives
// in its own collection or external system.

import mongoose, { Schema, Document } from "mongoose";

// ── Invoice line item ──────────────────────────────────────────────────────────
// Invoices now have line items so the PDF shows:
//   - Base plan fee
//   - Overage charge (students above limit)
//   - Any one-time charges (setup, training, etc.)
export interface IInvoiceLine {
  description: string;
  quantity: number;
  unitPrice: number;
  total: number;
}

// ── Invoice ────────────────────────────────────────────────────────────────────
export interface IInvoice {
  id: string;
  invoiceNumber: string; // Human-readable: INV-2024-001
  label: string;
  periodStart: Date; // First day of billing period
  periodEnd: Date; // Last day of billing period
  lines: IInvoiceLine[]; // Itemised breakdown
  subtotal: number;
  tax: number; // VAT/Tax amount (0 if exempt)
  total: number; // subtotal + tax
  currency: string; // "KES", "USD", "UGX", etc.
  status: "draft" | "sent" | "paid" | "overdue" | "void";
  dueAt: Date;
  paidAt?: Date;
  paidAmount?: number; // May differ from total if partial payment
  paymentRef?: string; // Manual payment reference (cheque no., bank ref)
  paymentMethod?: string; // "bank_transfer" | "cheque" | "cash" | "mpesa" | etc.
  notes?: string;
  createdAt: Date;
}

// ── Usage snapshot ─────────────────────────────────────────────────────────────
// Recorded automatically at invoice generation time — proves the seat count
// at the time of billing. Useful for disputes.
export interface IUsageSnapshot {
  snapshotDate: Date;
  activeStudents: number; // status = "active" | "repeat"
  totalStudents: number; // all students including on_leave, deferred, etc.
  seatLimit: number; // plan limit at snapshot time
  overage: number; // max(0, activeStudents - seatLimit)
}

// ── Plan change audit ──────────────────────────────────────────────────────────
export interface IPlanChange {
  date: Date;
  fromPlan: string;
  toPlan: string;
  changedBy: mongoose.Types.ObjectId; // User who made the change
  reason?: string;
}

// ── Billing contact (separate from admin user) ─────────────────────────────────
export interface IBillingContact {
  name: string;
  email: string;
  phone?: string;
  address?: string;
}

// ── Main Billing document ──────────────────────────────────────────────────────
export interface IBilling extends Document {
  institution: mongoose.Types.ObjectId;

  // Plan configuration
  planName: string; // "Starter" | "Growth" | "Pro" | "Enterprise" | "Custom"
  billingCycle: "monthly" | "annual";
  seatLimit: number;
  basePrice: number; // Base monthly price in currency
  overageRate: number; // Price per seat above seatLimit
  currency: string; // "KES" | "USD" | etc.
  taxRate: number; // 0–1, e.g. 0.16 for 16% VAT
  isCustomPlan: boolean; // True for enterprise with negotiated pricing
  customNotes?: string; // Private notes on custom pricing

  // Billing contact
  billingContact?: IBillingContact;

  // Invoice tracking
  invoices: IInvoice[];
  nextInvoiceDate: Date;
  invoiceCounter: number; // Monotonic counter for INV-YYYY-NNN generation

  // Usage history
  usageHistory: IUsageSnapshot[];

  // Plan change log
  planHistory: IPlanChange[];

  // Account status
  accountStatus: "active" | "suspended" | "cancelled" | "trial";
  trialEndsAt?: Date;
  suspendedAt?: Date;
  suspensionReason?: string;

  createdAt: Date;
  updatedAt: Date;
}

// ── Sub-schemas ────────────────────────────────────────────────────────────────

const invoiceLineSchema = new Schema<IInvoiceLine>(
  {
    description: { type: String, required: true },
    quantity: { type: Number, required: true, min: 0 },
    unitPrice: { type: Number, required: true, min: 0 },
    total: { type: Number, required: true, min: 0 },
  },
  { _id: false },
);

const invoiceSchema = new Schema<IInvoice>({
  id: { type: String, required: true },
  invoiceNumber: { type: String, required: true },
  label: { type: String, required: true },
  periodStart: { type: Date, required: true },
  periodEnd: { type: Date, required: true },
  lines: { type: [invoiceLineSchema], default: [] },
  subtotal: { type: Number, required: true, min: 0 },
  tax: { type: Number, default: 0, min: 0 },
  total: { type: Number, required: true, min: 0 },
  currency: { type: String, default: "KES" },
  status: {
    type: String,
    enum: ["draft", "sent", "paid", "overdue", "void"],
    default: "draft",
  },
  dueAt: { type: Date, required: true },
  paidAt: { type: Date },
  paidAmount: { type: Number, min: 0 },
  paymentRef: { type: String },
  paymentMethod: { type: String },
  notes: { type: String },
  createdAt: { type: Date, default: Date.now },
});

const usageSnapshotSchema = new Schema<IUsageSnapshot>(
  {
    snapshotDate: { type: Date, required: true },
    activeStudents: { type: Number, required: true },
    totalStudents: { type: Number, required: true },
    seatLimit: { type: Number, required: true },
    overage: { type: Number, required: true, min: 0 },
  },
  { _id: false },
);

const planChangeSchema = new Schema<IPlanChange>(
  {
    date: { type: Date, required: true, default: Date.now },
    fromPlan: { type: String, required: true },
    toPlan: { type: String, required: true },
    changedBy: { type: Schema.Types.ObjectId, ref: "User", required: true },
    reason: { type: String },
  },
  { _id: false },
);

const billingContactSchema = new Schema<IBillingContact>(
  {
    name: { type: String, required: true },
    email: { type: String, required: true },
    phone: { type: String },
    address: { type: String },
  },
  { _id: false },
);

// ── Main schema ────────────────────────────────────────────────────────────────

const billingSchema = new Schema<IBilling>(
  {
    institution: {
      type: Schema.Types.ObjectId,
      ref: "Institution",
      required: true,
      unique: true,
    },

    planName: { type: String, default: "Starter" },
    billingCycle: {
      type: String,
      enum: ["monthly", "annual"],
      default: "monthly",
    },
    seatLimit: { type: Number, default: 500, min: 1 },
    basePrice: { type: Number, default: 15000, min: 0 }, // KES 15,000 default
    overageRate: { type: Number, default: 25, min: 0 }, // KES 25 per overage seat
    currency: { type: String, default: "KES" },
    taxRate: { type: Number, default: 0, min: 0, max: 1 },
    isCustomPlan: { type: Boolean, default: false },
    customNotes: { type: String },

    billingContact: { type: billingContactSchema },

    invoices: { type: [invoiceSchema], default: [] },
    nextInvoiceDate: { type: Date, required: true },
    invoiceCounter: { type: Number, default: 0, min: 0 },

    usageHistory: { type: [usageSnapshotSchema], default: [] },
    planHistory: { type: [planChangeSchema], default: [] },

    accountStatus: {
      type: String,
      enum: ["active", "suspended", "cancelled", "trial"],
      default: "trial",
    },
    trialEndsAt: { type: Date },
    suspendedAt: { type: Date },
    suspensionReason: { type: String },
  },
  { timestamps: true },
);

// Indexes
billingSchema.index({ institution: 1 }, { unique: true });
billingSchema.index({ accountStatus: 1 });
billingSchema.index({ nextInvoiceDate: 1 });
billingSchema.index({ "invoices.status": 1 });
billingSchema.index({ "invoices.dueAt": 1 });

export default mongoose.model<IBilling>("Billing", billingSchema);
