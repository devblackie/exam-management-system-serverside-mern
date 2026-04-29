// serverside/src/utils/dateValidation.ts
//
// WHY DATE VALIDATION MATTERS
// ────────────────────────────────────────────────────────────────────────────
// When a date is set to "0000-00-00" or left as an empty string and cast to
// a JavaScript Date, it becomes `Invalid Date` or resolves to 1970-01-01
// (Unix epoch) depending on the parser. Here's what breaks:
//
//   1. AcademicYear.startDate = epoch → year sorting breaks, "2024/2025"
//      appears before "1970/1971" which doesn't exist but corrupts all
//      queries that use { startDate: { $gte: ... } }
//
//   2. student.academicLeavePeriod.endDate = epoch → leave is treated as
//      having ended 54 years ago. The leave logic thinks they're active
//      when they're not.
//
//   3. Billing.nextInvoiceDate = epoch → billing dashboard shows invoice
//      is perpetually overdue. If you add auto-payment triggers later,
//      this will fire immediately on every startup.
//
//   4. TempOTP.expiresAt = epoch → OTP query `{ expiresAt: { $gt: now } }`
//      finds nothing because epoch < now, so OTP verification always fails.
//
//   5. JourneyTimeline renders "Year 1 [0000/0001]" in the timeline header.
//      The frontend crashes or shows garbage.
//
// SOLUTION
// ────────────────────────────────────────────────────────────────────────────
// Two layers:
//   1. Mongoose validator (server) — rejects dates before 2000-01-01
//   2. Frontend date input constraint (client) — min="2000-01-01" + JS check
//
// Together these catch the problem at both layers.
 
import { Schema } from "mongoose";
 
// The minimum sensible date for this system — no student was at DeKUT in 1999
const MIN_DATE = new Date("2000-01-01T00:00:00.000Z");
const MAX_DATE = new Date("2100-01-01T00:00:00.000Z");
 
/**
 * Mongoose validator for required Date fields.
 * Returns false (fails validation) if date is:
 *   - Invalid Date (NaN)
 *   - Before 2000-01-01
 *   - After 2100-01-01 (typo guard)
 *   - Unix epoch exactly (the "00/00/0000" corruption marker)
 */
export function validateDate(value: Date | null | undefined): boolean {
  if (!value) return false;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return false;
  if (d.getTime() === 0) return false;          // Unix epoch — invalid input
  if (d < MIN_DATE) return false;               // Before year 2000
  if (d > MAX_DATE) return false;               // Implausible far future
  return true;
}
 
/**
 * Mongoose validator for optional Date fields.
 * Same rules but allows null/undefined (field is optional).
 */
export function validateOptionalDate(value: Date | null | undefined): boolean {
  if (value === null || value === undefined) return true;
  return validateDate(value);
}
 
/**
 * Reusable Mongoose schema field definition for a REQUIRED date.
 * Usage:
 *   startDate: requiredDateField("Start date"),
 */
export function requiredDateField(fieldName: string): object {
  return {
    type:     Date,
    required: true,
    validate: {
      validator: validateDate,
      message:   `${fieldName} cannot be before year 2000 or set to an invalid date. ` +
                 `Check for 00/00/0000 entries — these corrupt academic year logic.`,
    },
  };
}
 
/**
 * Reusable Mongoose schema field definition for an OPTIONAL date.
 */
export function optionalDateField(fieldName: string): object {
  return {
    type:     Date,
    validate: {
      validator: validateOptionalDate,
      message:   `${fieldName} cannot be before year 2000 or set to an invalid date.`,
    },
  };
}