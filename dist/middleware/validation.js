"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.paginationValidation = exports.outcomeValidation = exports.raiseCaseValidation = exports.marksUploadValidation = exports.studentUpdateValidation = exports.studentRegistrationValidation = exports.otpValidation = exports.loginValidation = exports.validateRequest = void 0;
// serverside/src/middleware/validation.ts — NEW FILE
const express_validator_1 = require("express-validator");
const validateRequest = (req, res, next) => {
    const errors = (0, express_validator_1.validationResult)(req);
    if (!errors.isEmpty()) {
        res.status(422).json({
            message: "Validation failed",
            // ✅ Explicitly type the error — no more implicit 'any'
            errors: errors.array().map((e) => ({
                field: e.type === "field" ? e.path : e.type,
                message: e.msg,
            })),
        });
        return;
    }
    next();
};
exports.validateRequest = validateRequest;
// All other validators (loginValidation, otpValidation, etc.) stay exactly as-is.
// ── Auth validations ──────────────────────────────────────────────────────────
exports.loginValidation = [
    (0, express_validator_1.body)("email")
        .isEmail()
        .withMessage("Valid email required")
        .normalizeEmail()
        .isLength({ max: 254 })
        .withMessage("Email too long"),
    (0, express_validator_1.body)("password")
        .isLength({ min: 8, max: 128 })
        .withMessage("Password must be 8–128 characters"),
];
exports.otpValidation = [
    (0, express_validator_1.body)("otp")
        .isNumeric()
        .withMessage("OTP must be numeric")
        .isLength({ min: 6, max: 6 })
        .withMessage("OTP must be exactly 6 digits"),
];
// ── Student validations ───────────────────────────────────────────────────────
exports.studentRegistrationValidation = [
    (0, express_validator_1.body)("students")
        .isArray({ min: 1, max: 500 })
        .withMessage("Provide 1–500 students"),
    (0, express_validator_1.body)("students.*.regNo")
        .matches(/^[A-Z0-9\-\/]+$/)
        .withMessage("Invalid reg number format")
        .isLength({ max: 30 })
        .withMessage("Reg number too long"),
    (0, express_validator_1.body)("students.*.name")
        .isLength({ min: 2, max: 120 })
        .withMessage("Name must be 2–120 characters")
        .trim()
        .escape(),
];
exports.studentUpdateValidation = [
    (0, express_validator_1.body)("name")
        .optional()
        .isLength({ min: 2, max: 120 })
        .withMessage("Name must be 2–120 characters")
        .trim()
        .escape(),
    (0, express_validator_1.body)("remarks")
        .optional()
        .isLength({ max: 500 })
        .withMessage("Remarks too long"),
];
// ── Marks validations ─────────────────────────────────────────────────────────
exports.marksUploadValidation = [
    (0, express_validator_1.body)("academicYearId")
        .isMongoId()
        .withMessage("Valid academic year ID required"),
    (0, express_validator_1.body)("programId").isMongoId().withMessage("Valid program ID required"),
];
// ── Disciplinary validations ──────────────────────────────────────────────────
exports.raiseCaseValidation = [
    (0, express_validator_1.body)("studentId").isMongoId().withMessage("Valid student ID required"),
    (0, express_validator_1.body)("grounds")
        .isIn([
        "exam_irregularity",
        "academic_misconduct",
        "misconduct",
        "financial",
        "other",
    ])
        .withMessage("Invalid grounds"),
    (0, express_validator_1.body)("description")
        .isLength({ min: 10, max: 2000 })
        .withMessage("Description must be 10–2000 characters")
        .trim()
        .escape(),
    (0, express_validator_1.body)("hearingDate")
        .optional()
        .isISO8601()
        .withMessage("Hearing date must be a valid ISO date"),
];
exports.outcomeValidation = [
    (0, express_validator_1.body)("outcome")
        .isIn(["WARNING", "SENT_HOME", "REINSTATED", "DISCONTINUED", "DISMISSED"])
        .withMessage("Invalid outcome"),
    (0, express_validator_1.body)("outcomeNotes").optional().isLength({ max: 1000 }),
    (0, express_validator_1.body)("suspensionStart")
        .optional()
        .isISO8601()
        .withMessage("Suspension start must be a valid ISO date"),
    (0, express_validator_1.body)("suspensionEnd")
        .optional()
        .isISO8601()
        .withMessage("Suspension end must be a valid ISO date"),
];
// ── Pagination validations (reusable on any list route) ───────────────────────
exports.paginationValidation = [
    (0, express_validator_1.query)("page")
        .optional()
        .isInt({ min: 1, max: 10000 })
        .withMessage("Page must be a positive integer")
        .toInt(),
    (0, express_validator_1.query)("limit")
        .optional()
        .isInt({ min: 1, max: 100 })
        .withMessage("Limit must be 1–100")
        .toInt(),
];
