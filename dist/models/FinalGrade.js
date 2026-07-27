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
// src/models/FinalGrade.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    student: { type: mongoose_1.Schema.Types.ObjectId, ref: "Student", required: true },
    programUnit: { type: mongoose_1.Schema.Types.ObjectId, ref: "ProgramUnit", required: true, },
    academicYear: { type: mongoose_1.Schema.Types.ObjectId, ref: "AcademicYear", required: true, },
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    semester: { type: mongoose_1.Schema.Types.Mixed, enum: ["SEMESTER 1", "SEMESTER 2", "SEMESTER 3"], required: true }, // or Number: 1, 2, 3   
    totalMark: { type: Number, required: true },
    grade: { type: String, required: true },
    remarks: { type: String },
    points: Number,
    status: { type: String, enum: ["PASS", "SUPPLEMENTARY", "RETAKE", "INCOMPLETE", "SPECIAL"], required: true },
    isSpecial: { type: Boolean, default: false },
    attemptType: { type: String, enum: ["1ST_ATTEMPT", "SPECIAL", "SUPPLEMENTARY", "RETAKE", "RE_RETAKE"], default: "1ST_ATTEMPT", required: true },
    attemptNumber: { type: Number, default: 1 },
    cappedBecauseSupplementary: { type: Boolean, default: false },
}, { timestamps: true });
schema.index({ student: 1, academicYear: 1 });
schema.index({ institution: 1, academicYear: 1, status: 1 });
schema.index({ student: 1, academicYear: 1, programUnit: 1 });
schema.index({ student: 1, createdAt: -1 });
schema.index({ institution: 1, academicYear: 1, status: 1, attemptType: 1 });
schema.index({ student: 1, programUnit: 1, status: 1 });
schema.index({ student: 1, academicYear: 1, status: 1 });
exports.default = mongoose_1.default.model("FinalGrade", schema);
