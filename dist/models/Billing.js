"use strict";
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
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const mongoose_1 = __importStar(require("mongoose"));
// ── Sub-schemas ────────────────────────────────────────────────────────────────
const invoiceLineSchema = new mongoose_1.Schema({
    description: { type: String, required: true },
    quantity: { type: Number, required: true, min: 0 },
    unitPrice: { type: Number, required: true, min: 0 },
    total: { type: Number, required: true, min: 0 },
}, { _id: false });
const invoiceSchema = new mongoose_1.Schema({
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
    status: { type: String, enum: ["draft", "sent", "paid", "overdue", "void"], default: "draft" },
    dueAt: { type: Date, required: true },
    paidAt: { type: Date },
    paidAmount: { type: Number, min: 0 },
    paymentRef: { type: String },
    paymentMethod: { type: String },
    notes: { type: String },
    createdAt: { type: Date, default: Date.now },
});
const usageSnapshotSchema = new mongoose_1.Schema({
    snapshotDate: { type: Date, required: true },
    activeStudents: { type: Number, required: true },
    totalStudents: { type: Number, required: true },
    seatLimit: { type: Number, required: true },
    overage: { type: Number, required: true, min: 0 },
}, { _id: false });
const planChangeSchema = new mongoose_1.Schema({
    date: { type: Date, required: true, default: Date.now },
    fromPlan: { type: String, required: true },
    toPlan: { type: String, required: true },
    changedBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User", required: true },
    reason: { type: String },
}, { _id: false });
const billingContactSchema = new mongoose_1.Schema({
    name: { type: String, required: true },
    email: { type: String, required: true },
    phone: { type: String },
    address: { type: String },
}, { _id: false });
const departmentSeatSchema = new mongoose_1.Schema({
    departmentCode: { type: String, required: true, uppercase: true },
    seatLimit: { type: Number, required: true, min: 1 },
}, { _id: false });
// ── Main schema ────────────────────────────────────────────────────────────────
const billingSchema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },
    planName: { type: String, default: "Starter" },
    billingCycle: { type: String, enum: ["monthly", "annual"], default: "monthly" },
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
    accountStatus: { type: String, enum: ["active", "suspended", "cancelled", "trial"], default: "trial" },
    trialEndsAt: { type: Date },
    suspendedAt: { type: Date },
    suspensionReason: { type: String },
    departmentSeats: { type: [departmentSeatSchema], default: [] },
}, { timestamps: true });
// Indexes
billingSchema.index({ accountStatus: 1 });
billingSchema.index({ nextInvoiceDate: 1 });
billingSchema.index({ "invoices.status": 1 });
billingSchema.index({ "invoices.dueAt": 1 });
exports.default = mongoose_1.default.model("Billing", billingSchema);
