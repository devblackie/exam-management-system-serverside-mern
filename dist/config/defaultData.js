"use strict";
// // src/config/defaultData.ts
// import Institution from "../models/Institution";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ensureDefaultInstitution = exports.ensureBillingForAllInstitutions = exports.ensureBillingRecord = void 0;
const Institution_1 = __importDefault(require("../models/Institution"));
const Billing_1 = __importDefault(require("../models/Billing"));
// ── ensureBillingRecord ───────────────────────────────────────────────────────
const ensureBillingRecord = async (institutionId) => {
    const exists = await Billing_1.default.exists({ institution: institutionId });
    if (exists)
        return;
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    nextMonth.setDate(1);
    await Billing_1.default.create({
        institution: institutionId,
        planName: "Trial",
        billingCycle: "monthly",
        seatLimit: 500,
        basePrice: 0,
        overageRate: 25,
        currency: "KES",
        taxRate: 0,
        isCustomPlan: false,
        accountStatus: "trial",
        trialEndsAt: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
        nextInvoiceDate: nextMonth,
        invoiceCounter: 0,
        invoices: [],
        usageHistory: [],
        planHistory: [],
    });
    console.log(`[Billing] Created billing record for institution ${institutionId}`);
};
exports.ensureBillingRecord = ensureBillingRecord;
// ── ensureBillingForAllInstitutions ───────────────────────────────────────────
const ensureBillingForAllInstitutions = async () => {
    const institutions = await Institution_1.default.find({}).select("_id name").lean();
    let created = 0;
    for (const inst of institutions) {
        // FIX: cast _id — lean() returns plain object where _id is unknown in strict TS
        const id = inst._id.toString();
        const exists = await Billing_1.default.exists({ institution: id });
        if (!exists) {
            await (0, exports.ensureBillingRecord)(id);
            created++;
        }
    }
    console.log(`[Billing] Migration complete: created ${created} billing records for ${institutions.length} institutions.`);
};
exports.ensureBillingForAllInstitutions = ensureBillingForAllInstitutions;
// ── ensureDefaultInstitution ──────────────────────────────────────────────────
const ensureDefaultInstitution = async (retries = 3) => {
    for (let i = 0; i < retries; i++) {
        try {
            const count = await Institution_1.default.countDocuments();
            if (count === 0) {
                console.log("[Setup] No institution found. Creating placeholder...");
                const inst = await Institution_1.default.create({
                    name: "My University", // ← generic placeholder, not "Demo University"
                    code: "UNIV",
                    isActive: true,
                });
                const id = inst._id.toString();
                console.log("[Setup] Placeholder institution created:", id);
                console.log("[Setup] IMPORTANT: Log in as admin and go to Admin → Institution Profile to set your university name.");
                await (0, exports.ensureBillingRecord)(id);
                return;
            }
            const active = await Institution_1.default.findOne({ isActive: true });
            if (!active) {
                const created = await Institution_1.default.create({ name: "My University", code: "UNIV", isActive: true });
                await (0, exports.ensureBillingRecord)(created._id.toString());
            }
            else {
                console.log("[Setup] Active institution found:", active.name);
                await (0, exports.ensureBillingRecord)(active._id.toString());
            }
            return;
        }
        catch (err) {
            const msg = err instanceof Error ? err.message : "unknown";
            console.error(`[Setup] Attempt ${i + 1} failed:`, msg);
            if (i < retries - 1)
                await new Promise(r => setTimeout(r, 2000));
        }
    }
};
exports.ensureDefaultInstitution = ensureDefaultInstitution;
