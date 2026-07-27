"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.logAudit = logAudit;
// lib/auditLogger.ts
const AuditLog_1 = __importDefault(require("../models/AuditLog"));
async function logAudit(req, { action, actor, targetUser, details = {}, }) {
    const ip = req.headers["x-forwarded-for"] ||
        req.socket.remoteAddress ||
        req.ip;
    const userAgent = req.headers["user-agent"];
    // Explicit partial type so actor/targetUser can be deleted safely
    const logEntry = {
        action,
        actor: actor || req.user?._id,
        targetUser: targetUser || actor || req.user?._id,
        details,
        ip,
        userAgent,
    };
    // Ensure empty fields don’t cause schema validation errors
    if (!logEntry.actor)
        delete logEntry.actor;
    if (!logEntry.targetUser)
        delete logEntry.targetUser;
    return Promise.resolve(AuditLog_1.default.create(logEntry).catch((err) => console.error("Audit log failed:", err)));
}
