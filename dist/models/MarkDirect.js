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
// src/models/MarkDirect.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    student: { type: mongoose_1.Schema.Types.ObjectId, ref: "Student", required: true },
    programUnit: { type: mongoose_1.Schema.Types.ObjectId, ref: "ProgramUnit", required: true },
    academicYear: { type: mongoose_1.Schema.Types.ObjectId, ref: "AcademicYear", required: true },
    semester: { type: String, enum: ["SEMESTER 1", "SEMESTER 2", "SEMESTER 3"], required: true, default: "SEMESTER 1" },
    batchId: { type: String, required: true, index: true },
    caTotal30: { type: Number, min: 0, max: 30, required: true },
    examTotal70: { type: Number, min: 0, max: 70, required: true },
    externalTotal100: { type: Number, min: 0, max: 100, default: null },
    agreedMark: { type: Number, min: 0, max: 100, required: true },
    attempt: { type: String, enum: ["1st", "re-take", "supplementary", "special"], default: "1st" },
    isSupplementary: { type: Boolean, default: false },
    isRetake: { type: Boolean, default: false },
    isSpecial: { type: Boolean, default: false }, // Explicit flag for Special Exams
    remarks: { type: String },
    uploadedBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User", required: true },
    uploadedAt: { type: Date, default: Date.now },
    deletedAt: { type: Date, default: null },
}, { timestamps: true });
schema.index({ student: 1, programUnit: 1, academicYear: 1 }, { unique: true });
schema.index({ student: 1, programUnit: 1, academicYear: 1, deletedAt: 1 });
schema.pre(/^find/, function (next) {
    const query = this.getQuery();
    if (query.deletedAt === undefined) {
        this.where({ deletedAt: null });
    }
    next();
});
schema.index({ batchId: 1, institution: 1 });
schema.index({ batchId: 1, deletedAt: 1 });
exports.default = mongoose_1.default.model("MarkDirect", schema);
