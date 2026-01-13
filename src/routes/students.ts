// src/routes/students.ts
import { Router, Response } from "express";
import { normalizeProgramName } from "../services/programNormalizer";
import Student from "../models/Student";
import Program from "../models/Program";
import AcademicYear from "../models/AcademicYear";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import mongoose from "mongoose";

const router = Router();

// GET all students
router.get(
  "/",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const students = await Student.find({ institution: req.user.institution })
      .select("regNo name program admissionAcademicYear currentYearOfStudy")
      .populate("program", "name code")
      .lean();

    res.json(students);
  })
);

// GET student statistics for dashboard
router.get(
  "/stats",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const institutionId = req.user.institution;

    const stats = await Student.aggregate([
      // 1. Filter students by the coordinator's institution
      { $match: { institution: institutionId } },
      
      // 2. Group by status and count
      {
        $group: {
          _id: "$status",
          count: { $sum: 1 },
        },
      },
    ]);

    // Format the result into a clean object (e.g., { active: 100, inactive: 5, total: 105 })
    let active = 0;
    let inactive = 0;
    let total = 0;

    for (const stat of stats) {
      total += stat.count;
      if (stat._id === "active") {
        active = stat.count;
      }
      // Sum all non-active statuses into 'inactive' for the dashboard display
      if (stat._id !== "active") {
        inactive += stat.count;
      }
    }

    res.json({
      active,
      inactive, // This will include graduated, suspended, deferred, and true inactive
      total,
    });
  })
);

// BULK register students — NO DUPLICATES
router.post(
  "/bulk",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { students } = req.body;
    if (!Array.isArray(students) || students.length === 0) {
      return res.status(400).json({ message: "No students provided" });
    }

    const institutionId = req.user.institution;

    // CLEAN & NORMALIZE INPUT
    const incoming = students.map((s) => ({
      regNo: s.regNo?.trim().toUpperCase(),
      name: s.name?.trim(),
      rawProgram: s.program?.trim(),
      normalizedProgram: normalizeProgramName(s.program?.trim() || ""),
      yearOfStudy: Number(s.yearOfStudy) || 1,
      admissionAcademicYearString: s.admissionAcademicYear || "2024/2025", // ⬅️ Renamed field to avoid conflict
    }));

    // Validate required fields (Reg No, Name, Program) - unchanged
    const invalid = incoming.filter(
      (s) => !s.regNo || !s.name || !s.rawProgram
    );
    if (invalid.length > 0) {
      return res.status(400).json({ message: "Missing Reg No, Name, or Program" });
    }

    // --- STEP 1: LOOKUP PROGRAM IDs (Unchanged) ---
    // ... (Program lookup logic remains here) ...
    
    // Get unique normalized program names
    const normNames = [
      ...new Set(incoming.map((s) => s.normalizedProgram)),
    ];
    // Fetch all programs for this institution
    const programs = await Program.find({ institution: institutionId }).lean();
    // Build a normalized map
    const programMap = new Map<string, string>();
    for (const p of programs) {
      const norm = normalizeProgramName(p.name);
      programMap.set(norm, p._id.toString());
    }

    // Identify missing programs
    const missingPrograms: any[] = [];
    for (const n of normNames) {
      if (!programMap.has(n)) {
        missingPrograms.push(n);
      }
    }
    if (missingPrograms.length > 0) {
      return res.status(400).json({
        message: "Some programs not found",
        notFound: missingPrograms,
      });
    }


    // --- STEP 2: LOOKUP ACADEMIC YEAR IDs (NEW LOGIC) ---
    
    // 2a. Get unique academic year strings from incoming data
    const yearStrings = [
      ...new Set(incoming.map((s) => s.admissionAcademicYearString)),
    ];

    // 2b. Fetch corresponding AcademicYear documents
    const academicYears = await AcademicYear.find({
      institution: institutionId,
      year: { $in: yearStrings },
    }).lean();

    // 2c. Build a map of year string -> ObjectId
    const academicYearMap = new Map<string, mongoose.Types.ObjectId>();
 const existingYearStrings = new Set(academicYears.map(y => y.year));

    // Map existing IDs
 for (const year of academicYears) {
 academicYearMap.set(year.year, year._id as mongoose.Types.ObjectId);
 }

    // Identify years that need to be created
 const yearsToCreate = yearStrings.filter(yearStr => !existingYearStrings.has(yearStr));

    // 2d. CRITICAL: Create missing academic years with default dates.
    if (yearsToCreate.length > 0) {
        // Create an array of documents to insert/upsert
        const documentsToUpsert = yearsToCreate.map(yearStr => {
            const [startYear, endYear] = yearStr.split('/').map(Number);
            const defaultStartDate = new Date(`${startYear}-08-01`); // Default to Aug 1st
            const defaultEndDate = new Date(`${endYear}-07-31`);     // Default to July 31st

            return {
                year: yearStr,
                institution: institutionId,
                startDate: defaultStartDate,
                endDate: defaultEndDate,
                isCurrent: false,
            };
        });

        // Use insertMany (or bulkWrite) for efficiency
        try {
           
            const bulkOps = yearsToCreate.map(yearStr => {
                 const [startYear, endYear] = yearStr.split('/').map(Number);
                 return {
                    updateOne: {
                        filter: { year: yearStr, institution: institutionId },
                        update: { 
                            $setOnInsert: { 
                                year: yearStr,
                                institution: institutionId,
                                startDate: new Date(`${startYear}-08-01`),
                                endDate: new Date(`${endYear}-07-31`),
                                isCurrent: false,
                            }
                        },
                        upsert: true,
                    }
                };
            });

            await AcademicYear.bulkWrite(bulkOps);

            // Re-fetch all academic years, including the newly created ones, to populate the map with ObjectIds
            const allAcademicYears = await AcademicYear.find({
                institution: institutionId,
                year: { $in: yearStrings },
            }).lean();

            for (const year of allAcademicYears) {
                // Map the ID (guaranteed to exist now)
                academicYearMap.set(year.year, year._id as mongoose.Types.ObjectId);
            }

        } catch (error) {
            // Handle any database error during creation (e.g., race condition, validation)
            console.error("Error during bulk creation of academic years:", error);
            return res.status(500).json({ 
                message: "A database error occurred while defining missing academic years. Please check the logs." 
            });
        }
    }


    // --- STEP 3: DETECT EXISTING STUDENTS (Unchanged) ---
    const regNos = incoming.map((s) => s.regNo);
    const existing = await Student.find({
      regNo: { $in: regNos },
      institution: institutionId,
    });
    
    const existingRegNos = new Set(existing.map((s) => s.regNo));


    // --- STEP 4: BUILD FINAL PAYLOAD (UPDATED) ---
    const toCreate = incoming
      .filter((s) => !existingRegNos.has(s.regNo)) // Only include non-existing students
      .map((s) => ({
        regNo: s.regNo,
        name: s.name,
        program: programMap.get(s.normalizedProgram)!,
        currentYearOfStudy: s.yearOfStudy,
        
        // 🔥 CRUCIAL CHANGE: Use ObjectId from the map
        admissionAcademicYear: academicYearMap.get(s.admissionAcademicYearString)!,
        
        institution: institutionId,
        status: "active",
      }));


    // --- STEP 5: INSERT AND RESPOND (Unchanged) ---
    if (toCreate.length > 0) {
      await Student.insertMany(toCreate);

      return res.status(200).json({
        message: `${toCreate.length} students registered successfully.`,
        registered: toCreate.map((s) => s.regNo),
        alreadyRegistered: Array.from(existingRegNos),
      });
    }

    return res.status(200).json({
        message: "All students in the list are already registered.",
        alreadyRegistered: Array.from(existingRegNos),
    });
  })
);

export default router;