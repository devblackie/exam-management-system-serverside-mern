// serverside/src/routes/auditLogs.ts
import { Router, Response } from "express";
import { Workbook } from "exceljs";
import AuditLog, { IAuditLog } from "../models/AuditLog";
import { requireAuth, requireRole, AuthenticatedRequest } from "../middleware/auth";
import { logAudit } from "../lib/auditLogger";
import { asyncHandler } from "../middleware/asyncHandler";

const router = Router();

interface LogFilter {
  action?:    string;
  actor?:     string;
  createdAt?: { $gte?: Date; $lte?: Date };
}

function buildFilter(query: Record<string, unknown>): LogFilter {
  const { action, actorId, fromDate, toDate } = query;
  const filter: LogFilter = {};
  if (action   && typeof action   === "string") filter.action = action;
  if (actorId  && typeof actorId  === "string") filter.actor  = actorId;
  if (fromDate || toDate) {
    filter.createdAt = {};
    if (fromDate && typeof fromDate === "string") filter.createdAt.$gte = new Date(fromDate);
    if (toDate   && typeof toDate   === "string") filter.createdAt.$lte = new Date(toDate);
  }
  return filter;
}

interface PopulatedActor {
  _id:   string;
  name:  string;
  email: string;
}

// GET /audit-logs
// Paginated listing.
router.get(
  "/",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { page = 1, limit = 10, sort = "desc" } = req.query;
    const filter = buildFilter(req.query);
    const skip   = (Number(page) - 1) * Number(limit);

    const [logs, total] = await Promise.all([
      AuditLog.find(filter)
        .populate<{ actor: PopulatedActor }>("actor", "name email")
        .populate<{ targetUser: PopulatedActor }>("targetUser", "name email")
        .sort({ createdAt: sort === "asc" ? 1 : -1 })
        .skip(skip)
        .limit(Number(limit))
        .lean(),
      AuditLog.countDocuments(filter),
    ]);

    // Non-blocking audit of this view (fire and forget)
    logAudit(req, {
      action:  "audit_logs_viewed",
      actor:   req.user._id,
      details: { page: Number(page), resultsReturned: logs.length, totalMatching: total },
    }).catch(console.error);

    res.json({
      data:  logs,
      total,
      page:  Number(page),
      pages: Math.ceil(total / Number(limit)),
    });
  })
);

// GET /audit-logs/export/csv
router.get(
  "/export/csv",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { sort = "desc" } = req.query;
    const filter = buildFilter(req.query);

    const logs = await AuditLog.find(filter)
      .populate<{ actor: PopulatedActor }>("actor", "name email")
      .populate<{ targetUser: PopulatedActor }>("targetUser", "name email")
      .sort({ createdAt: sort === "asc" ? 1 : -1 })
      .lean();

    const header = "Actor Name,Actor Email,Target Name,Target Email,Action,Details,Created At\n";
    const rows   = logs.map(log => {
      const actor  = log.actor  as unknown as PopulatedActor | undefined;
      const target = log.targetUser as unknown as PopulatedActor | undefined;
      return [
        `"${actor?.name  ?? ""}"`,
        `"${actor?.email ?? ""}"`,
        `"${target?.name  ?? ""}"`,
        `"${target?.email ?? ""}"`,
        `"${log.action}"`,
        `"${JSON.stringify(log.details ?? {}).replace(/"/g, '""')}"`,
        `"${log.createdAt.toISOString()}"`,
      ].join(",");
    });

    logAudit(req, {
      action:  "audit_logs_exported_csv",
      actor:   req.user._id,
      details: { exportedCount: logs.length },
    }).catch(console.error);

    res.setHeader("Content-Type",        "text/csv");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.csv");
    res.send(header + rows.join("\n"));
  })
);

// GET /audit-logs/export/excel
router.get(
  "/export/excel",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { sort = "desc" } = req.query;
    const filter = buildFilter(req.query);

    const logs = await AuditLog.find(filter)
      .populate<{ actor: PopulatedActor }>("actor", "name email")
      .populate<{ targetUser: PopulatedActor }>("targetUser", "name email")
      .sort({ createdAt: sort === "asc" ? 1 : -1 })
      .lean();

    const workbook  = new Workbook();
    const worksheet = workbook.addWorksheet("Audit Logs");

    worksheet.columns = [
      { header: "Actor Name",        key: "actorName",   width: 20 },
      { header: "Actor Email",       key: "actorEmail",  width: 28 },
      { header: "Target Name",       key: "targetName",  width: 20 },
      { header: "Target Email",      key: "targetEmail", width: 28 },
      { header: "Action",            key: "action",      width: 22 },
      { header: "Details",           key: "details",     width: 40 },
      { header: "IP",                key: "ip",          width: 16 },
      { header: "User Agent",        key: "userAgent",   width: 40 },
      { header: "Created At",        key: "createdAt",   width: 26 },
    ];

    logs.forEach(log => {
      const actor  = log.actor  as unknown as PopulatedActor | undefined;
      const target = log.targetUser as unknown as PopulatedActor | undefined;
      // const logDoc = log as IAuditLog & { createdAt: Date; ip?: string; userAgent?: string };
      worksheet.addRow({
        actorName:   actor?.name,
        actorEmail:  actor?.email,
        targetName:  target?.name,
        targetEmail: target?.email,
        action:      log.action,
        details:     JSON.stringify(log.details ?? {}),
        ip:          log.ip,
        userAgent:   log.userAgent,
        createdAt:   log.createdAt.toISOString(),
      });
    });

    logAudit(req, {
      action:  "audit_logs_exported_excel",
      actor:   req.user._id,
      details: { exportedCount: logs.length },
    }).catch(console.error);

    res.setHeader("Content-Type",        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.xlsx");
    await workbook.xlsx.write(res);
    res.end();
  })
);

// DELETE /audit-logs/bulk
router.delete(
  "/bulk",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { ids } = req.body as { ids?: unknown };

    if (!Array.isArray(ids) || ids.length === 0) {
      res.status(400).json({ message: "Provide a non-empty array of log IDs." });
      return;
    }

    // Validate all elements are strings before hitting the DB
    const validIds = ids.filter((id): id is string => typeof id === "string");
    if (validIds.length !== ids.length) {
      res.status(400).json({ message: "All IDs must be strings." });
      return;
    }

    // Snapshot for audit trail before deletion
    const targets = await AuditLog.find({ _id: { $in: validIds } })
      .select("action createdAt")
      .lean();

    const foundIds   = targets.map(t => t._id.toString());
    const missingIds = validIds.filter(id => !foundIds.includes(id));

    const result = await AuditLog.deleteMany({ _id: { $in: validIds } });

    logAudit(req, {
      action:  "audit_logs_bulk_deleted",
      actor:   req.user._id,
      details: {
        requested: validIds.length,
        deleted:   result.deletedCount,
        notFound:  missingIds,
      },
    }).catch(console.error);

    res.json({
      message:      `${result.deletedCount} log(s) deleted.`,
      deletedCount: result.deletedCount,
      notFound:     missingIds,
    });
  })
);

// DELETE /audit-logs/purge/by-date
router.delete(
  "/purge/by-date",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { before } = req.body as { before?: unknown };

    if (!before || typeof before !== "string") {
      res.status(400).json({ message: "'before' date string is required." });
      return;
    }

    const cutoff = new Date(before);
    if (isNaN(cutoff.getTime())) {
      res.status(400).json({ message: `Invalid date: "${before}". Use ISO 8601 format.` });
      return;
    }

    const countBefore = await AuditLog.countDocuments({ createdAt: { $lt: cutoff } });
    const result      = await AuditLog.deleteMany({ createdAt: { $lt: cutoff } });

    logAudit(req, {
      action:  "audit_logs_purged_by_date",
      actor:   req.user._id,
      details: { cutoff: cutoff.toISOString(), deleted: result.deletedCount },
    }).catch(console.error);

    res.json({
      message:      `Purged ${result.deletedCount} log(s) older than ${cutoff.toISOString()}.`,
      deletedCount: result.deletedCount,
    });
  })
);

// DELETE /audit-logs/:id
router.delete(
  "/:id",
  requireAuth,
  requireRole("admin"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { id } = req.params;

    const target = await AuditLog.findById(id).lean();
    if (!target) {
      res.status(404).json({ message: "Audit log not found." });
      return;
    }

    await AuditLog.findByIdAndDelete(id);

    logAudit(req, {
      action:  "audit_log_deleted",
      actor:   req.user._id,
      details: { deletedLogId: id, originalAction: target.action },
    }).catch(console.error);

    res.json({ message: "Audit log deleted." });
  })
);

export default router;