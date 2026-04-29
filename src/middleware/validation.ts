// serverside/src/middleware/validation.ts — NEW FILE
import { body, query, validationResult, ValidationError } from "express-validator";
import { Request, Response, NextFunction } from "express";

export const validateRequest = (
  req:  Request,
  res:  Response,
  next: NextFunction,
): void => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) {
    res.status(422).json({
      message: "Validation failed",
      // ✅ Explicitly type the error — no more implicit 'any'
      errors: errors.array().map((e: ValidationError) => ({
        field:   e.type === "field" ? e.path : e.type,
        message: e.msg,
      })),
    });
    return;
  }
  next();
};
// All other validators (loginValidation, otpValidation, etc.) stay exactly as-is.

// ── Auth validations ──────────────────────────────────────────────────────────
export const loginValidation = [
  body("email")
    .isEmail()
    .withMessage("Valid email required")
    .normalizeEmail()
    .isLength({ max: 254 })
    .withMessage("Email too long"),
  body("password")
    .isLength({ min: 8, max: 128 })
    .withMessage("Password must be 8–128 characters"),
];

export const otpValidation = [
  body("otp")
    .isNumeric()
    .withMessage("OTP must be numeric")
    .isLength({ min: 6, max: 6 })
    .withMessage("OTP must be exactly 6 digits"),
];

// ── Student validations ───────────────────────────────────────────────────────
export const studentRegistrationValidation = [
  body("students")
    .isArray({ min: 1, max: 500 })
    .withMessage("Provide 1–500 students"),
  body("students.*.regNo")
    .matches(/^[A-Z0-9\-\/]+$/)
    .withMessage("Invalid reg number format")
    .isLength({ max: 30 })
    .withMessage("Reg number too long"),
  body("students.*.name")
    .isLength({ min: 2, max: 120 })
    .withMessage("Name must be 2–120 characters")
    .trim()
    .escape(),
];

export const studentUpdateValidation = [
  body("name")
    .optional()
    .isLength({ min: 2, max: 120 })
    .withMessage("Name must be 2–120 characters")
    .trim()
    .escape(),
  body("remarks")
    .optional()
    .isLength({ max: 500 })
    .withMessage("Remarks too long"),
];

// ── Marks validations ─────────────────────────────────────────────────────────
export const marksUploadValidation = [
  body("academicYearId")
    .isMongoId()
    .withMessage("Valid academic year ID required"),
  body("programId").isMongoId().withMessage("Valid program ID required"),
];

// ── Disciplinary validations ──────────────────────────────────────────────────
export const raiseCaseValidation = [
  body("studentId").isMongoId().withMessage("Valid student ID required"),
  body("grounds")
    .isIn([
      "exam_irregularity",
      "academic_misconduct",
      "misconduct",
      "financial",
      "other",
    ])
    .withMessage("Invalid grounds"),
  body("description")
    .isLength({ min: 10, max: 2000 })
    .withMessage("Description must be 10–2000 characters")
    .trim()
    .escape(),
  body("hearingDate")
    .optional()
    .isISO8601()
    .withMessage("Hearing date must be a valid ISO date"),
];

export const outcomeValidation = [
  body("outcome")
    .isIn(["WARNING", "SENT_HOME", "REINSTATED", "DISCONTINUED", "DISMISSED"])
    .withMessage("Invalid outcome"),
  body("outcomeNotes").optional().isLength({ max: 1000 }),
  body("suspensionStart")
    .optional()
    .isISO8601()
    .withMessage("Suspension start must be a valid ISO date"),
  body("suspensionEnd")
    .optional()
    .isISO8601()
    .withMessage("Suspension end must be a valid ISO date"),
];

// ── Pagination validations (reusable on any list route) ───────────────────────
export const paginationValidation = [
  query("page")
    .optional()
    .isInt({ min: 1, max: 10000 })
    .withMessage("Page must be a positive integer")
    .toInt(),
  query("limit")
    .optional()
    .isInt({ min: 1, max: 100 })
    .withMessage("Limit must be 1–100")
    .toInt(),
];
