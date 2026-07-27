"use strict";
// // src/models/Unit.ts
// import mongoose, { Schema, Document } from "mongoose";
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
// export interface IUnit extends Document {
//   institution: mongoose.Types.ObjectId;
//   code: string;
//   name: string;
// }
// const schema = new Schema<IUnit>(
//   {
//     institution: {
//       type: Schema.Types.ObjectId,
//       ref: "Institution",
//       required: true,
//     },
//     code: { type: String, required: true, uppercase: true, trim: true },
//     name: { type: String, required: true, trim: true },
//   },
//   { timestamps: true }
// );
// // Indexes simplified to focus on institution/code uniqueness
// schema.index({ institution: 1, code: 1 }, { unique: true });
// export default mongoose.model<IUnit>("Unit", schema);
// serverside/src/models/Unit.ts
const mongoose_1 = __importStar(require("mongoose"));
const schema = new mongoose_1.Schema({
    institution: {
        type: mongoose_1.Schema.Types.ObjectId,
        ref: "Institution",
        required: true,
    },
    code: { type: String, required: true, uppercase: true, trim: true },
    name: { type: String, required: true, trim: true },
    schoolCode: { type: String, required: true, uppercase: true, trim: true },
    departmentCode: { type: String, required: true, uppercase: true, trim: true },
}, { timestamps: true });
// A unit code is unique within its department (not globally)
schema.index({ institution: 1, departmentCode: 1, code: 1 }, { unique: true });
// For querying all units in a department
// For querying all units in a school
schema.index({ institution: 1, schoolCode: 1 });
exports.default = mongoose_1.default.model("Unit", schema);
