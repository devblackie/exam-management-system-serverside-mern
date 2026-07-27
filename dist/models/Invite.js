"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const mongoose_1 = require("mongoose");
const inviteSchema = new mongoose_1.Schema({
    institution: { type: mongoose_1.Schema.Types.ObjectId, ref: "Institution", required: true },
    name: { type: String, required: true },
    email: { type: String, required: true },
    schoolCode: { type: String, default: null },
    departmentCode: { type: String, default: null },
    institutionWide: { type: Boolean, default: false },
    token: { type: String, required: true, unique: true },
    role: { type: String, enum: ["coordinator", "lecturer"], required: true },
    used: { type: Boolean, default: false },
    expiresAt: { type: Date, required: true },
    createdBy: { type: mongoose_1.Schema.Types.ObjectId, ref: "User" },
}, { timestamps: true });
inviteSchema.index({ institution: 1, email: 1 });
inviteSchema.index({ expiresAt: 1 });
exports.default = (0, mongoose_1.model)("Invite", inviteSchema);
