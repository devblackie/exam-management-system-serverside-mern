// src/routes/lecturer.ts
import { Router, Request, Response } from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import { Types } from "mongoose";
import Unit, { IUnit } from "../models/Unit";
import Submission from "../models/Submission";
import InstitutionSettings from "../models/InstitutionSettings";
import { requireAuth, requireRole } from "../middleware/auth";
import { logAudit } from "../lib/auditLogger";
import { toNodeBuffer } from "../lib/bufferUtils";

const router = Router();

// memory storage (file is available on req.file.buffer)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
});

// GET units assigned to currently logged-in lecturer
router.get(
  "/units",
  requireAuth,
  requireRole("lecturer"),
  async (req: Request & { user?: any }, res: Response) => {
    const lecturerId = req.user?._id;
    const units = await Unit.find({ assignedLecturer: lecturerId }).populate(
      "program",
      "code title"
    );
    return res.json(units);
  }
);

// Upload results (lecturer)
router.post(
  "/upload",
  requireAuth,
  requireRole("lecturer"),
  upload.single("file"),
  async (
    req: Request & { user?: any; file?: Express.Multer.File },
    res: Response
  ) => {
    try {
      const lecturer = req.user;
      const { unitId } = req.body;
      if (!unitId) return res.status(400).json({ message: "unitId required" });

      const unit = (await Unit.findById(unitId)) as
        | (IUnit & { _id: Types.ObjectId })
        | null;
      if (!unit) return res.status(404).json({ message: "Unit not found" });

      // check lecturer is assigned to unit (safer string compare)
      const assignedLecturerId = unit.assignedLecturer?.toString();
      if (!assignedLecturerId || assignedLecturerId !== String(lecturer._id)) {
        return res
          .status(403)
          .json({ message: "You are not assigned to this unit" });
      }

      // get settings with sane defaults
      const settings = (await InstitutionSettings.findOne({})) || {
        cat1Weight: 10,
        cat2Weight: 10,
        cat3Weight: 0,
        assignmentWeight: 5,
        practicalWeight: 5,
        examWeight: 70,
        supplementaryThreshold: 40,
        retakeThreshold: 5,
      };

      if (!req.file) return res.status(400).json({ message: "File required" });

      // === CONVERT the incoming buffer to a Node Buffer reliably ===
      const fileBuffer = toNodeBuffer((req.file as Express.Multer.File).buffer);

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(fileBuffer); // now type-safe and runtime-safe
      const worksheet = workbook.worksheets[0];

      // Read headers and map columns (case-insensitive)
      const headerRow = worksheet.getRow(1).values as any[];
      const headers = headerRow.slice(1).map((h: any) =>
        String(h || "")
          .trim()
          .toLowerCase()
      );

      const findCol = (name: string) => {
        const idx = headers.findIndex((h: string) => h.includes(name));
        return idx >= 0 ? idx + 1 : -1;
      };

      const colRegistration = findCol("registration");
      const colName = findCol("student");
      const colCat1 = findCol("cat1");
      const colCat2 = findCol("cat2");
      const colCat3 = findCol("cat3");
      const colAssignment = findCol("assign");
      const colPractical = findCol("pract");
      const colExam = findCol("exam");

      if (colRegistration === -1) {
        return res
          .status(400)
          .json({
            message:
              "Registration column not found. Include a 'RegistrationNo' column.",
          });
      }

      const rows: any[] = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const regNo = String(row.getCell(colRegistration).text || "").trim();
        if (!regNo) return;

        const sName =
          colName !== -1
            ? String(row.getCell(colName).text || "").trim()
            : undefined;
        const cat1 =
          colCat1 !== -1 ? Number(row.getCell(colCat1).value ?? null) : null;
        const cat2 =
          colCat2 !== -1 ? Number(row.getCell(colCat2).value ?? null) : null;
        const cat3 =
          colCat3 !== -1 ? Number(row.getCell(colCat3).value ?? null) : null;
        const assignment =
          colAssignment !== -1
            ? Number(row.getCell(colAssignment).value ?? null)
            : null;
        const practical =
          colPractical !== -1
            ? Number(row.getCell(colPractical).value ?? null)
            : null;
        const exam =
          colExam !== -1 ? Number(row.getCell(colExam).value ?? null) : null;

        let computedTotal = 0;
        let missing = false;

        const addWeighted = (mark: number | null, weight: number) => {
          if (mark === null || Number.isNaN(mark)) {
            missing = true;
            return 0;
          }
          return mark * (weight / 100);
        };

        computedTotal += addWeighted(cat1, settings.cat1Weight);
        computedTotal += addWeighted(cat2, settings.cat2Weight);
        if ((settings.cat3Weight || 0) > 0)
          computedTotal += addWeighted(cat3, settings.cat3Weight);
        if ((settings.assignmentWeight || 0) > 0)
          computedTotal += addWeighted(assignment, settings.assignmentWeight);
        if ((settings.practicalWeight || 0) > 0)
          computedTotal += addWeighted(practical, settings.practicalWeight);
        computedTotal += addWeighted(exam, settings.examWeight);

        rows.push({
          registrationNo: regNo,
          studentName: sName,
          cat1: Number.isNaN(cat1 ?? NaN) ? null : cat1,
          cat2: Number.isNaN(cat2 ?? NaN) ? null : cat2,
          cat3: Number.isNaN(cat3 ?? NaN) ? null : cat3,
          assignment: Number.isNaN(assignment ?? NaN) ? null : assignment,
          practical: Number.isNaN(practical ?? NaN) ? null : practical,
          exam: Number.isNaN(exam ?? NaN) ? null : exam,
          computedTotal: missing ? null : Number(computedTotal.toFixed(2)),
          status: missing ? "missing" : "ok",
        });
      });

      const submission = new Submission({
        unit: unit._id,
        lecturer: lecturer._id,
        fileName: req.file.originalname,
        rows,
        status: "pending",
      });
      await submission.save();

      // silent audit log
      logAudit(req, {
        action: "lecturer_upload",
        actor: lecturer._id,
        targetUser: lecturer._id,
        details: {
          unit: unit._id.toString(),
          filename: req.file.originalname,
          rowsCount: rows.length,
        },
      });

      return res.json({
        message: "File parsed and stored (pending coordinator review)",
        submissionId: submission._id,
        preview: rows.slice(0, 50),
      });
    } catch (err) {
      console.error("Upload error:", err);
      return res.status(500).json({ message: "Error processing file" });
    }
  }
);

export default router;
