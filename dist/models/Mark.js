"use strict";
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
// src/models/Mark.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true, },
    student: { type: mongoose_1.Schema.Types.ObjectId, ref: "Student", required: true },
    programUnit: { type: mongoose_1.Schema.Types.ObjectId, ref: "ProgramUnit", required: true, }, // IMPORTANT
    academicYear: { type: mongoose_1.Schema.Types.ObjectId, ref: "AcademicYear", required: true, },
    semester: { type: String, enum: ["SEMESTER 1", "SEMESTER 2", "SEMESTER 3"], required: true, default: "SEMESTER 1" },
    batchId: { type: String, required: true, index: true },
    // RAW CA SCORES (The system will derive the final CA/30 from these)
    cat1Raw: { type: Number, min: 0, default: 0 },
    cat2Raw: { type: Number, min: 0, default: 0 },
    cat3Raw: { type: Number, min: 0 },
    assgnt1Raw: { type: Number, min: 0, max: 10, default: 0 },
    assgnt2Raw: { type: Number, min: 0, max: 10 },
    assgnt3Raw: { type: Number, min: 0, max: 10 },
    practicalRaw: { type: Number, min: 0, max: 100 }, // Assume flexible scale if not explicitly 20/10
    // RAW EXAM SCORES
    examQ1Raw: { type: Number, min: 0, default: 0 },
    examQ2Raw: { type: Number, min: 0, max: 20, default: 0 },
    examQ3Raw: { type: Number, min: 0, max: 20, default: 0 },
    examQ4Raw: { type: Number, min: 0, max: 20, default: 0 },
    examQ5Raw: { type: Number, min: 0, max: 20 },
    // FINAL AUDIT FIELDS (Filled with scoresheet data)
    caTotal30: { type: Number, min: 0, max: 30, required: true },
    examTotal70: { type: Number, min: 0, max: 70, required: true },
    internalExaminerMark: { type: Number, min: 0, max: 100, required: false },
    agreedMark: { type: Number, min: 0, max: 100, required: true },
    attempt: { type: String, enum: ["1st", "re-take", "supplementary", "special"], default: "1st", },
    examMode: { type: String, enum: ["standard", "mandatory_q1"], default: "standard" },
    // METADATA
    isSupplementary: { type: Boolean, default: false },
    isRetake: { type: Boolean, default: false },
    isSpecial: { type: Boolean, default: false }, // Explicit flag for Special Exams
    remarks: { type: String },
    uploadedBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User", required: true },
    uploadedAt: { type: Date, default: Date.now },
    deletedAt: { type: Date, default: null },
}, { timestamps: true });
// Unique index must be on the combination of Student, ProgramUnit, and AcademicYear
schema.index({ student: 1, programUnit: 1, academicYear: 1 }, { unique: true });
schema.pre(/^find/, function (next) {
    const query = this.getQuery();
    if (query.deletedAt === undefined) {
        this.where({ deletedAt: null });
    }
    next();
});
schema.index({ student: 1, academicYear: 1, programUnit: 1, deletedAt: 1 });
schema.index({ institution: 1, student: 1, deletedAt: 1 });
schema.index({ institution: 1, uploadedAt: -1 });
schema.index({ student: 1, programUnit: 1, deletedAt: 1 });
schema.index({ batchId: 1, institution: 1 });
schema.index({ batchId: 1, deletedAt: 1 });
exports.default = mongoose_1.default.model("Mark", schema);
