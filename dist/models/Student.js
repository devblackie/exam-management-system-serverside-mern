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
// src/models/Student.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    regNo: { type: String, required: true, uppercase: true },
    name: { type: String, required: true },
    program: { type: mongoose_1.Schema.Types.ObjectId, ref: "Program", required: true },
    programType: { type: String, required: true },
    entryType: { type: String, default: "Direct" },
    currentYearOfStudy: { type: Number, default: 1 },
    currentSemester: { type: Number, default: 1 },
    remarks: { type: String },
    admissionAcademicYear: { type: mongoose_1.Schema.Types.ObjectId, ref: "AcademicYear", required: true },
    intake: { type: String, required: true, enum: ["JAN", "MAY", "SEPT"], uppercase: true, default: "SEPT" },
    status: { type: String, default: "active" },
    qualifierSuffix: { type: String, default: "" },
    carryForwardUnits: [{
            programUnitId: { type: String, required: true },
            unitCode: { type: String, required: true },
            unitName: { type: String, required: true },
            fromYear: { type: Number, required: true },
            fromAcademicYear: { type: String, required: true },
            attemptNumber: { type: Number, default: 3 },
            qualifier: { type: String, default: "RP1C" },
            addedAt: { type: Date, default: Date.now },
            status: { type: String, enum: ["pending", "passed", "failed", "escalated_to_rpu"], default: "pending" },
        }],
    deferredSuppUnits: [{
            programUnitId: { type: String, required: true },
            unitCode: { type: String, required: true },
            unitName: { type: String, required: true },
            fromYear: { type: Number, required: true },
            fromAcademicYear: { type: String, required: true },
            // "supp_deferred"    = student skipped supp period, sitting in next ordinary
            // "special_deferred" = student's special exam deferred to next ordinary
            reason: { type: String, enum: ["supp_deferred", "special_deferred"], default: "supp_deferred" },
            addedAt: { type: Date, default: Date.now },
            status: { type: String, enum: ["pending", "passed", "failed"], default: "pending" },
        }],
    academicLeavePeriod: { startDate: Date, endDate: Date, reason: String, type: { type: String, enum: ["compassionate", "financial", "other"] } },
    totalTimeOutYears: { type: Number, default: 0 },
    unitAttemptRegistry: [{ unitId: { type: mongoose_1.Schema.Types.ObjectId, ref: "Unit" }, attempts: [{ attemptNumber: Number, mark: Number, passed: Boolean, type: String }] }],
    academicHistory: [{ academicYear: String, yearOfStudy: Number, annualMeanMark: Number, weightedContribution: Number, failedUnitsCount: Number, isRepeatYear: Boolean, date: Date }],
    statusHistory: [{ status: String, previousStatus: String, date: { type: Date, default: Date.now }, reason: String }],
    statusEvents: [{ fromStatus: String, toStatus: String, date: { type: Date, default: Date.now }, academicYear: String, reason: String }],
}, { timestamps: true });
// All indexes declared here only — no field-level index:true to avoid duplicate index warnings
schema.index({ institution: 1, regNo: 1 }, { unique: true });
schema.index({ institution: 1, program: 1, admissionAcademicYear: 1 });
schema.index({ institution: 1, intake: 1 });
schema.index({ institution: 1, status: 1, currentYearOfStudy: 1 });
schema.index({ institution: 1, program: 1, currentYearOfStudy: 1, status: 1 });
schema.index({ institution: 1, admissionAcademicYear: 1, intake: 1 });
schema.index({ "statusEvents.toStatus": 1, "statusEvents.academicYear": 1 });
schema.index({ qualifierSuffix: 1 });
schema.index({ "carryForwardUnits.programUnitId": 1 });
schema.index({ "deferredSuppUnits.programUnitId": 1 });
schema.index({ institution: 1, regNo: 1, status: 1 });
exports.default = mongoose_1.default.model("Student", schema);
