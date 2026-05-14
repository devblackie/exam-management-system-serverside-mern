// // src/config/defaultData.ts
// import Institution from "../models/Institution";

// export const ensureDefaultInstitution = async (retries = 3) => {
//   for (let i = 0; i < retries; i++) {
//     try {
//       const count = await Institution.countDocuments();

//       if (count === 0) {
//         console.log("No institution found. Creating default...");
//         const defaultInst = await Institution.create({
//           name: "Demo University",
//           code: "DEMO",
//           isActive: true,
//         });
//         console.log("Default institution created →", defaultInst.name, defaultInst._id);
//         return;
//       }

//       const active = await Institution.findOne({ isActive: true });
//       if (!active) {
//         console.log("No active institution. Creating default...");
//         await Institution.create({
//           name: "Department of Civil Engineering",
//           code: "CE",
//           isActive: true,
//         });
//       } else {
//         console.log("Active institution found →", active.name);
//       }
//       return;

//     } catch (err: any) {
//       console.error(`Attempt ${i + 1} failed:`, err.message);
//       if (i === retries - 1) {
//         console.error("Failed to ensure default institution after retries");
//       } else {
//         await new Promise(resolve => setTimeout(resolve, 2000));
//       }
//     }
//   }
// };

// import Institution from "../models/Institution";
// import Billing from "../models/Billing";

// // ── ensureBillingRecord ────────────────────────────────────────────────────────
// // Creates a Billing document for the given institution if one doesn't exist.
// // Called both at startup and whenever a new institution is created.
// export const ensureBillingRecord = async (
//   institutionId: string,
// ): Promise<void> => {
//   const exists = await Billing.exists({ institution: institutionId });
//   if (exists) return; // Already has a billing record — nothing to do

//   const nextMonth = new Date();
//   nextMonth.setMonth(nextMonth.getMonth() + 1);
//   nextMonth.setDate(1); // First of next month

//   await Billing.create({
//     institution: institutionId,
//     planName: "Trial",
//     billingCycle: "monthly",
//     seatLimit: 500,
//     basePrice: 0, // No charge during trial
//     overageRate: 25,
//     currency: "KES",
//     taxRate: 0,
//     isCustomPlan: false,
//     accountStatus: "trial",
//     trialEndsAt: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), // 30 days
//     nextInvoiceDate: nextMonth,
//     invoiceCounter: 0,
//     invoices: [],
//     usageHistory: [],
//     planHistory: [],
//   });

//   console.log(
//     `[Billing] Created billing record for institution ${institutionId}`,
//   );
// };

// // ── ensureBillingForAllInstitutions ───────────────────────────────────────────
// // One-time migration helper. Run once against existing Atlas data to backfill
// // Billing documents for all institutions that don't have one.
// export const ensureBillingForAllInstitutions = async (): Promise<void> => {
//   const institutions = await Institution.find({}).select("_id name").lean();
//   let created = 0;

//   for (const inst of institutions) {
//     const exists = await Billing.exists({ institution: inst._id });
//     if (!exists) {
//       await ensureBillingRecord(inst._id.toString());
//       created++;
//     }
//   }

//   console.log(
//     `[Billing] Migration complete: created ${created} billing records for ${institutions.length} institutions.`,
//   );
// };

// // ── ensureDefaultInstitution (patched) ────────────────────────────────────────
// // Same as original but now calls ensureBillingRecord after creating/finding
// // an institution.
// export const ensureDefaultInstitution = async (retries = 3) => {
//   for (let i = 0; i < retries; i++) {
//     try {
//       const count = await Institution.countDocuments();

//       if (count === 0) {
//         console.log("No institution found. Creating default...");
//         const defaultInst = await Institution.create({
//           name: "Demo University",
//           code: "DEMO",
//           isActive: true,
//         });
//         console.log(
//           "Default institution created →",
//           defaultInst.name,
//           defaultInst._id,
//         );
//         // ← NEW: seed billing for the new institution
//         // await ensureBillingRecord(defaultInst._id.toString());
//         await ensureBillingRecord(defaultInst._id.toString());
//         return;
//       }

//       const active = await Institution.findOne({ isActive: true });
//       if (!active) {
//         console.log("No active institution. Creating default...");
//         const created = await Institution.create({
//           name: "Department of Civil Engineering",
//           code: "CE",
//           isActive: true,
//         });
//         // ← NEW
//         // await ensureBillingRecord(created._id.toString());
//         await ensureBillingRecord(created._id.toString());
//       } else {
//         console.log("Active institution found →", active.name);
//         // ← NEW: backfill if missing (handles existing deployments)
//         // await ensureBillingRecord(active._id.toString());
//         await ensureBillingRecord(active._id.toString());
//       }

//       return;
//     } catch (err: any) {
//       console.error(`Attempt ${i + 1} failed:`, err.message);
//       if (i === retries - 1) {
//         console.error("Failed to ensure default institution after retries");
//       } else {
//         await new Promise((resolve) => setTimeout(resolve, 2000));
//       }
//     }
//   }
// };





// src/config/defaultData.ts
import mongoose from "mongoose";
import Institution from "../models/Institution";
import Billing from "../models/Billing";

// ── ensureBillingRecord ───────────────────────────────────────────────────────
export const ensureBillingRecord = async (institutionId: string): Promise<void> => {
  const exists = await Billing.exists({ institution: institutionId });
  if (exists) return;

  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  nextMonth.setDate(1);

  await Billing.create({
    institution:    institutionId,
    planName:       "Trial",
    billingCycle:   "monthly",
    seatLimit:      500,
    basePrice:      0,
    overageRate:    25,
    currency:       "KES",
    taxRate:        0,
    isCustomPlan:   false,
    accountStatus:  "trial",
    trialEndsAt:    new Date(Date.now() + 30 * 24 * 60 * 60 * 1000),
    nextInvoiceDate: nextMonth,
    invoiceCounter: 0,
    invoices:       [],
    usageHistory:   [],
    planHistory:    [],
  });

  console.log(`[Billing] Created billing record for institution ${institutionId}`);
};

// ── ensureBillingForAllInstitutions ───────────────────────────────────────────
export const ensureBillingForAllInstitutions = async (): Promise<void> => {
  const institutions = await Institution.find({}).select("_id name").lean();
  let created = 0;

  for (const inst of institutions) {
    // FIX: cast _id — lean() returns plain object where _id is unknown in strict TS
    const id = (inst._id as mongoose.Types.ObjectId).toString();
    const exists = await Billing.exists({ institution: id });
    if (!exists) {
      await ensureBillingRecord(id);
      created++;
    }
  }

  console.log(`[Billing] Migration complete: created ${created} billing records for ${institutions.length} institutions.`);
};

// ── ensureDefaultInstitution ──────────────────────────────────────────────────
export const ensureDefaultInstitution = async (retries = 3): Promise<void> => {
  for (let i = 0; i < retries; i++) {
    try {
      const count = await Institution.countDocuments();

      if (count === 0) {
        console.log("[Setup] No institution found. Creating placeholder...");
        const inst = await Institution.create({
          name:     "My University",   // ← generic placeholder, not "Demo University"
          code:     "UNIV",
          isActive: true,
        });
        const id = (inst._id as mongoose.Types.ObjectId).toString();
        console.log("[Setup] Placeholder institution created:", id);
        console.log("[Setup] IMPORTANT: Log in as admin and go to Admin → Institution Profile to set your university name.");
        await ensureBillingRecord(id);
        return;
      }

      const active = await Institution.findOne({ isActive: true });
      if (!active) {
        const created = await Institution.create({ name: "My University", code: "UNIV", isActive: true });
        await ensureBillingRecord((created._id as mongoose.Types.ObjectId).toString());
      } else {
        console.log("[Setup] Active institution found:", active.name);
        await ensureBillingRecord((active._id as mongoose.Types.ObjectId).toString());
      }
      return;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : "unknown";
      console.error(`[Setup] Attempt ${i + 1} failed:`, msg);
      if (i < retries - 1) await new Promise(r => setTimeout(r, 2000));
    }
  }
};