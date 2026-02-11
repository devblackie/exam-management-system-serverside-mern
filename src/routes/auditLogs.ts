// serverside/src/routes/auditLogs.ts
import { Router } from "express";
import { Workbook } from "exceljs";
import { IUser } from "../models/User";
import AuditLog, { IAuditLog } from "../models/AuditLog";
import { requireAuth, requireRole } from "../middleware/auth";

const router = Router();


router.get("/", requireAuth, requireRole("admin"), async (req, res) => {
  try {
    const { action, actorId, fromDate, toDate, page = 1, limit = 10, sort = "desc" } = req.query;

    const filter: Record<string, any> = {};
    if (action) filter.action = action;
    if (actorId) filter.actor = actorId;
    if (fromDate || toDate) {
      filter.createdAt = {};
      if (fromDate) filter.createdAt.$gte = new Date(fromDate as string);
      if (toDate) filter.createdAt.$lte = new Date(toDate as string);
    }

    const skip = (Number(page) - 1) * Number(limit);

    const logs = await AuditLog.find(filter)
      .populate<{ actor: IUser }>("actor", "name email")
      .populate<{ targetUser: IUser }>("targetUser", "name email")
      .sort({ createdAt: sort === "asc" ? 1 : -1 })
      .skip(skip)
      .limit(Number(limit));

    const total = await AuditLog.countDocuments(filter);

    res.json({
      data: logs,
      total,
      page: Number(page),
      pages: Math.ceil(total / Number(limit)),
    });
  } catch (err) {
    console.error("Error fetching audit logs:", err);
    res.status(500).json({ message: "Server error" });
  }
});


// EXPORT audit logs (CSV)
router.get("/export/csv", requireAuth, requireRole("admin"), async (req, res) => {
  try {
    const { action, actorId, fromDate, toDate, sort = "desc" } = req.query;

    const filter: Record<string, any> = {};
    if (action) filter.action = action;
    if (actorId) filter.actor = actorId;
    if (fromDate || toDate) {
      filter.createdAt = {};
      if (fromDate) filter.createdAt.$gte = new Date(fromDate as string);
      if (toDate) filter.createdAt.$lte = new Date(toDate as string);
    }

    const logs = await AuditLog.find(filter)
      .populate<{ actor: IUser }>("actor", "name email")
      .populate<{ targetUser: IUser }>("targetUser", "name email")
      .sort({ createdAt: sort === "asc" ? 1 : -1 });

    // Build CSV
    let csv = "Actor Name,Actor Email,Target Name,Target Email,Action,Details,Created At\n";
    logs.forEach((log) => {
      const actor = log.actor as IUser;
      const target = log.targetUser as IUser;
      csv += `"${actor?.name || ""}","${actor?.email || ""}","${target?.name || ""}","${target?.email || ""}","${log.action}","${JSON.stringify(log.details)}","${log.createdAt.toISOString()}"\n`;
    });

    res.setHeader("Content-Type", "text/csv");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.csv");
    res.send(csv);
  } catch (err) {
    console.error("Error exporting CSV:", err);
    res.status(500).json({ message: "Server error" });
  }
});


// EXPORT audit logs (Excel)
router.get("/export/excel", requireAuth, requireRole("admin"), async (req, res) => {
  try {
    const { action, actorId, fromDate, toDate, sort = "desc" } = req.query;

    const filter: Record<string, any> = {};
    if (action) filter.action = action;
    if (actorId) filter.actor = actorId;
    if (fromDate || toDate) {
      filter.createdAt = {};
      if (fromDate) filter.createdAt.$gte = new Date(fromDate as string);
      if (toDate) filter.createdAt.$lte = new Date(toDate as string);
    }

    const logs = await AuditLog.find(filter)
      .populate<{ actor: IUser }>("actor", "name email")
      .populate<{ targetUser: IUser }>("targetUser", "name email")
      .sort({ createdAt: sort === "asc" ? 1 : -1 });

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Audit Logs");

    worksheet.columns = [
      { header: "Actor Name", key: "actorName", width: 20 },
      { header: "Actor Email", key: "actorEmail", width: 25 },
      { header: "Target User Name", key: "targetName", width: 20 },
      { header: "Target User Email", key: "targetEmail", width: 25 },
      { header: "Action", key: "action", width: 20 },
      { header: "Details", key: "details", width: 30 },
      { header: "IP", key: "ip", width: 20 },
      { header: "UserAgent", key: "userAgent", width: 40 },
      { header: "Created At", key: "createdAt", width: 25 },
    ];

    logs.forEach((log) => {
      const actor = log.actor as IUser;
      const target = log.targetUser as IUser;

      worksheet.addRow({
        actorName: actor?.name,
        actorEmail: actor?.email,
        targetName: target?.name,
        targetEmail: target?.email,
        action: log.action,
        details: JSON.stringify(log.details),
          ip: log.ip,
        userAgent: log.userAgent,
        createdAt: log.createdAt.toISOString(),
      });
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=audit_logs.xlsx");

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Error exporting Excel:", err);
    res.status(500).json({ message: "Server error" });
  }
});

export default router;
