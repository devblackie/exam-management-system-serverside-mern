"use strict";
// serverside/src/models/DisciplinaryCase.ts
//
// WHY THIS EXISTS
// ──────────────a─
// The previous codebase had `isDisciplinary` flags and `RP1D` qualifier logic
// in academicRules.ts but zero backend machinery to actually CREATE a disciplinary
// case, record a hearing outcome, or suspend a student. A coordinator had no way
// to formally "send a student home" in the system — they could only manually edit
// the student's status field with no audit trail.
//
// This model gives that process a proper home. Every suspension, every hearing,
// every appeal lives here. The Student model is updated as a side-effect.
//
// HOW IT LINKS TO THE REST OF THE SYSTEM
// ───────────────────────────────────────
// 1. DisciplinaryCase created → student.status set to "disciplinary_suspension"
// 2. statusEvents entry written to Student (picked up by JourneyTimeline)
// 3. requireAuth middleware blocks suspended students on every API call
// 4. If outcome = "SENT_HOME", qualifierSuffix is set to RP1D via deriveQualifierSuffix
// 5. If outcome = "REINSTATED", student.status reverts to prior status
// 6. AuditLog entry written for every state change
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
const schema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    student: { type: mongoose_1.Schema.Types.ObjectId, ref: "Student", required: true },
    raisedBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User", required: true },
    academicYear: { type: mongoose_1.Schema.Types.ObjectId, ref: "AcademicYear", required: true },
    yearOfStudy: { type: Number, required: true, min: 1, max: 8 },
    grounds: {
        type: String,
        enum: ["exam_irregularity", "academic_misconduct", "misconduct", "financial", "other"],
        required: true,
    },
    description: { type: String, required: true, minlength: 10 },
    hearingDate: { type: Date },
    outcome: {
        type: String,
        enum: ["PENDING", "WARNING", "SENT_HOME", "REINSTATED", "DISCONTINUED", "DISMISSED"],
        default: "PENDING",
    },
    outcomeNotes: { type: String },
    resolvedBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User" },
    resolvedAt: { type: Date },
    appealed: { type: Boolean, default: false },
    appealDate: { type: Date },
    appealOutcome: { type: String, enum: ["UPHELD", "DISMISSED"] },
    appealNotes: { type: String },
    suspensionStart: { type: Date },
    suspensionEnd: { type: Date },
    priorStudentStatus: { type: String, required: true, default: "active" },
}, { timestamps: true });
// Indexes
schema.index({ institution: 1, student: 1 });
schema.index({ institution: 1, outcome: 1 });
schema.index({ institution: 1, academicYear: 1 });
schema.index({ student: 1, createdAt: -1 });
exports.default = mongoose_1.default.model("DisciplinaryCase", schema);
