"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/routes/auditLogs.ts
const express_1 = require("express");
const exceljs_1 = require("exceljs");
const AuditLog_1 = __importDefault(require("../models/AuditLog"));
const auth_1 = require("../middleware/auth");
const auditLogger_1 = require("../lib/auditLogger");
const asyncHandler_1 = require("../middleware/asyncHandler");
const router = (0, express_1.Router)();
function buildFilter(query) {
    const { action, actorId, fromDate, toDate } = query;
    const filter = {};
    if (action && typeof action === "string")
        filter.action = action;
    if (actorId && typeof actorId === "string")
        filter.actor = actorId;
    if (fromDate || toDate) {
        filter.createdAt = {};
        if (fromDate && typeof fromDate === "string")
            filter.createdAt.$gte = new Date(fromDate);
        if (toDate && typeof toDate === "string")
            filter.createdAt.$lte = new Date(toDate);
    }
    return filter;
}
// GET /audit-logs
// Paginated listing.
router.get("/", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { page = 1, limit = 10, sort = "desc" } = req.query;
    const filter = buildFilter(req.query);
    const skip = (Number(page) - 1) * Number(limit);
    const [logs, total] = await Promise.all([
        AuditLog_1.default.find(filter)
            .populate("actor", "name email")
            .populate("targetUser", "name email")
            .sort({ createdAt: sort === "asc" ? 1 : -1 })
            .skip(skip)
            .limit(Number(limit))
            .lean(),
        AuditLog_1.default.countDocuments(filter),
    ]);
    // Non-blocking audit of this view (fire and forget)
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_logs_viewed",
        actor: req.user._id,
        details: { page: Number(page), resultsReturned: logs.length, totalMatching: total },
    }).catch(console.error);
    res.json({
        data: logs,
        total,
        page: Number(page),
        pages: Math.ceil(total / Number(limit)),
    });
}));
// GET /audit-logs/export/csv
router.get("/export/csv", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { sort = "desc" } = req.query;
    const filter = buildFilter(req.query);
    const logs = await AuditLog_1.default.find(filter)
        .populate("actor", "name email")
        .populate("targetUser", "name email")
        .sort({ createdAt: sort === "asc" ? 1 : -1 })
        .lean();
    const header = "Actor Name,Actor Email,Target Name,Target Email,Action,Details,Created At\n";
    const rows = logs.map(log => {
        const actor = log.actor;
        const target = log.targetUser;
        return [
            `"${actor?.name ?? ""}"`,
            `"${actor?.email ?? ""}"`,
            `"${target?.name ?? ""}"`,
            `"${target?.email ?? ""}"`,
            `"${log.action}"`,
            `"${JSON.stringify(log.details ?? {}).replace(/"/g, '""')}"`,
            `"${log.createdAt.toISOString()}"`,
        ].join(",");
    });
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_logs_exported_csv",
        actor: req.user._id,
        details: { exportedCount: logs.length },
    }).catch(console.error);
    res.setHeader("Content-Type", "text/csv");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.csv");
    res.send(header + rows.join("\n"));
}));
// GET /audit-logs/export/excel
router.get("/export/excel", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { sort = "desc" } = req.query;
    const filter = buildFilter(req.query);
    const logs = await AuditLog_1.default.find(filter)
        .populate("actor", "name email")
        .populate("targetUser", "name email")
        .sort({ createdAt: sort === "asc" ? 1 : -1 })
        .lean();
    const workbook = new exceljs_1.Workbook();
    const worksheet = workbook.addWorksheet("Audit Logs");
    worksheet.columns = [
        { header: "Actor Name", key: "actorName", width: 20 },
        { header: "Actor Email", key: "actorEmail", width: 28 },
        { header: "Target Name", key: "targetName", width: 20 },
        { header: "Target Email", key: "targetEmail", width: 28 },
        { header: "Action", key: "action", width: 22 },
        { header: "Details", key: "details", width: 40 },
        { header: "IP", key: "ip", width: 16 },
        { header: "User Agent", key: "userAgent", width: 40 },
        { header: "Created At", key: "createdAt", width: 26 },
    ];
    logs.forEach(log => {
        const actor = log.actor;
        const target = log.targetUser;
        // const logDoc = log as IAuditLog & { createdAt: Date; ip?: string; userAgent?: string };
        worksheet.addRow({
            actorName: actor?.name,
            actorEmail: actor?.email,
            targetName: target?.name,
            targetEmail: target?.email,
            action: log.action,
            details: JSON.stringify(log.details ?? {}),
            ip: log.ip,
            userAgent: log.userAgent,
            createdAt: log.createdAt.toISOString(),
        });
    });
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_logs_exported_excel",
        actor: req.user._id,
        details: { exportedCount: logs.length },
    }).catch(console.error);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.xlsx");
    await workbook.xlsx.write(res);
    res.end();
}));
// DELETE /audit-logs/bulk
router.delete("/bulk", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
        res.status(400).json({ message: "Provide a non-empty array of log IDs." });
        return;
    }
    // Validate all elements are strings before hitting the DB
    const validIds = ids.filter((id) => typeof id === "string");
    if (validIds.length !== ids.length) {
        res.status(400).json({ message: "All IDs must be strings." });
        return;
    }
    // Snapshot for audit trail before deletion
    const targets = await AuditLog_1.default.find({ _id: { $in: validIds } })
        .select("action createdAt")
        .lean();
    const foundIds = targets.map(t => t._id.toString());
    const missingIds = validIds.filter(id => !foundIds.includes(id));
    const result = await AuditLog_1.default.deleteMany({ _id: { $in: validIds } });
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_logs_bulk_deleted",
        actor: req.user._id,
        details: { requested: validIds.length, deleted: result.deletedCount, notFound: missingIds },
    }).catch(console.error);
    res.json({
        message: `${result.deletedCount} log(s) deleted.`,
        deletedCount: result.deletedCount,
        notFound: missingIds,
    });
}));
// DELETE /audit-logs/purge/by-date
router.delete("/purge/by-date", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { before } = req.body;
    if (!before || typeof before !== "string") {
        res.status(400).json({ message: "'before' date string is required." });
        return;
    }
    const cutoff = new Date(before);
    if (isNaN(cutoff.getTime())) {
        res.status(400).json({ message: `Invalid date: "${before}". Use ISO 8601 format.` });
        return;
    }
    const countBefore = await AuditLog_1.default.countDocuments({ createdAt: { $lt: cutoff } });
    const result = await AuditLog_1.default.deleteMany({ createdAt: { $lt: cutoff } });
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_logs_purged_by_date",
        actor: req.user._id,
        details: { cutoff: cutoff.toISOString(), deleted: result.deletedCount },
    }).catch(console.error);
    res.json({
        message: `Purged ${result.deletedCount} log(s) older than ${cutoff.toISOString()}.`,
        deletedCount: result.deletedCount,
    });
}));
// DELETE /audit-logs/:id
router.delete("/:id", auth_1.requireAuth, (0, auth_1.requireRole)("admin"), (0, asyncHandler_1.asyncHandler)(async (req, res) => {
    const { id } = req.params;
    const target = await AuditLog_1.default.findById(id).lean();
    if (!target) {
        res.status(404).json({ message: "Audit log not found." });
        return;
    }
    await AuditLog_1.default.findByIdAndDelete(id);
    (0, auditLogger_1.logAudit)(req, {
        action: "audit_log_deleted",
        actor: req.user._id,
        details: { deletedLogId: id, originalAction: target.action },
    }).catch(console.error);
    res.json({ message: "Audit log deleted." });
}));
exports.default = router;
