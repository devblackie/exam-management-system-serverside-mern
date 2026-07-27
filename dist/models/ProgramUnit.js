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
// src/models/ProgramUnit.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    program: { type: mongoose_1.Schema.Types.ObjectId, ref: "Program", required: true },
    unit: { type: mongoose_1.Schema.Types.ObjectId, ref: "Unit", required: true },
    academicYear: { type: String },
    requiredYear: { type: Number, required: true, min: 1, max: 6 },
    requiredSemester: { type: Number, required: true, enum: [1, 2] },
    isElective: { type: Boolean, default: false },
}, { timestamps: true });
// Index: A Unit should be defined once per Program/Year/Semester
schema.index({ program: 1, unit: 1, academicYear: 1 }, { unique: true });
schema.index({ program: 1, requiredYear: 1, requiredSemester: 1 });
schema.index({ program: 1, requiredYear: 1, requiredSemester: 1, institution: 1 });
exports.default = mongoose_1.default.model("ProgramUnit", schema);
