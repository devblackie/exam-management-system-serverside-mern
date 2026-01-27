// serverside/src/routes/promote.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { bulkPromoteClass, calculateStudentStatus, previewPromotion } from "../services/statusEngine";
import Program from "../models/Program";
import { generatePromotionWordDoc, generateEligibleSummaryDoc, generateIneligibleSummaryDoc, generateIneligibilityNotice, PromotionData, generateSpecialExamNotice, generateStudentTranscript } from "../utils/promotionReport";
import fs from "fs";
import path from "path";
import AdmZip from "adm-zip";
import Student from "../models/Student";
import { logAudit } from "../lib/auditLogger";

const router = Router();

// preview-promotion
router.post(
  "/preview-promotion",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    if (!programId || !yearToPromote || !academicYearName) {
      return res.status(400).json({ error: "Missing parameters" });
    }

    const previewData = await previewPromotion(programId, yearToPromote, academicYearName);
    res.json({ success: true, data: previewData });
  })
);
// bulk-promote
router.post(
  "/bulk-promote",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    if (!programId || !yearToPromote || !academicYearName) {
      return res.status(400).json({ error: "Missing required promotion parameters" });
    }

    const results = await bulkPromoteClass(programId, yearToPromote, academicYearName);

    res.json({
      success: true,
      message: `Process completed: ${results.promoted} promoted, ${results.failed} failed.`,
      data: results
    });
  })
);

// promote (individual student)
router.post(
  "/promote/:studentId",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId } = req.params;

    // 1. Find the student
    const student = await Student.findById(studentId);
    if (!student) {
      return res.status(404).json({ error: "Student not found" });
    }

    // 2. Safety Check: Verify status before promoting
    // We import calculateStudentStatus to ensure they actually passed
    const status = await calculateStudentStatus(
      student._id,
      student.program,
      "", // Academic year is optional if we are checking current curriculum
      student.currentYearOfStudy
    );

    if (status.variant !== "success") {
      return res.status(400).json({ 
        error: "Promotion Denied", 
        details: "Student has not met the requirements (Failed/Missing units)." 
      });
    }

    // 3. Increment the year
    const oldYear = student.currentYearOfStudy;
    student.currentYearOfStudy += 1;
    await student.save();

    // 4. Log the Audit
    await logAudit(req, {
      action: "individual_student_promoted",
      targetUser: student._id as any,
      details: { fromYear: oldYear, toYear: student.currentYearOfStudy },
    });

    res.json({
      success: true,
      message: `${student.name} promoted to Year ${student.currentYearOfStudy}`,
      newYear: student.currentYearOfStudy
    });
  })
);

// download-report
router.post(
  "/download-report",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    
    // 1. SET HEADERS FOR STREAMING
    // This tells the browser/proxy NOT to buffer the response
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename=Promotion_Package_Year_${yearToPromote}.zip`);
    res.setHeader('X-Content-Type-Options', 'nosniff');

    try {
      // 2. FETCH DATA
      const preview = await previewPromotion(programId, yearToPromote, academicYearName);
      const program = await Program.findById(programId).lean();
      
      // SEND HEARTBEAT (A single space keeps the socket open)
      res.write(" "); 

      const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
      let logoBuffer = Buffer.alloc(0);
      if (fs.existsSync(logoPath)) {
        logoBuffer = fs.readFileSync(logoPath);
      }

      const promotionData: PromotionData = {
        programName: program?.name || "Unknown Program",
        academicYear: academicYearName,
        yearOfStudy: yearToPromote,
        eligible: preview.eligible,
        blocked: preview.blocked,
        logoBuffer
      };

      // 3. GENERATE DOCUMENTS WITH PULSES
      const mainBuffer = await generatePromotionWordDoc(promotionData);
      res.write(" "); // Pulse

      const eligibleBuffer = await generateEligibleSummaryDoc(promotionData);
      res.write(" "); // Pulse

      const ineligibleBuffer = await generateIneligibleSummaryDoc(promotionData);
      res.write(" "); // Pulse

      // 4. CREATE ZIP
      const zip = new AdmZip();
      zip.addFile(`Promotion_Report_${program?.code}_Year${yearToPromote}.docx`, mainBuffer);
      zip.addFile(`Eligible_Students_${program?.code}_Year${yearToPromote}.docx`, eligibleBuffer);
      zip.addFile(`Ineligible_Students_${program?.code}_Year${yearToPromote}.docx`, ineligibleBuffer);

      const zipBuffer = zip.toBuffer();

      // 5. FINAL SEND
      // We use .end() because we used .write() earlier
      res.write(zipBuffer);
      res.end();
      
      console.log("[download-report] Streaming complete");
    } catch (err: any) {
      console.error("[download-report] CRASH:", err);
      // If we already sent headers/pulses, we can't send a JSON error
      if (!res.headersSent) {
        res.status(500).json({ error: "Failed to generate report" });
      } else {
        res.end();
      }
    }
  })
);

// download-report-progress
router.post(
  "/download-report-progress",
  requireAuth,
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

   const sendProgress = (percent: number, message: string, file?: string) => {
  const data = JSON.stringify({ percent, message, file });
  res.write(`data: ${data}\n\n`); // Must have two \n
};

    try {
      sendProgress(10, "Fetching student data...");
      const preview = await previewPromotion(programId, yearToPromote, academicYearName);
      const program = await Program.findById(programId).lean();

      sendProgress(30, "Generating Main Word Document...");
      const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
      let logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

      const promotionData = {
        programName: program?.name || "Program",
        academicYear: academicYearName,
        yearOfStudy: yearToPromote,
        eligible: preview.eligible,
        blocked: preview.blocked,
        logoBuffer
      };

      const mainBuffer = await generatePromotionWordDoc(promotionData);
      
      sendProgress(60, "Generating Eligible Summary...");
      const eligibleBuffer = await generateEligibleSummaryDoc(promotionData);

      sendProgress(80, "Generating Ineligible Summary...");
      const ineligibleBuffer = await generateIneligibleSummaryDoc(promotionData);

      sendProgress(95, "Creating ZIP Archive...");
      const zip = new AdmZip();
      zip.addFile(`Report_${program?.code}.docx`, mainBuffer);
      zip.addFile(`Eligible_${program?.code}.docx`, eligibleBuffer);
      zip.addFile(`Ineligible_${program?.code}.docx`, ineligibleBuffer);

      const zipBase64 = zip.toBuffer().toString('base64');
      
      // Final message with the file data
      res.write(`data: ${JSON.stringify({ percent: 100, message: "Complete!", file: zipBase64 })}\n\n`);
      res.end();
    } catch (err) {
      res.write(`data: ${JSON.stringify({ error: "Failed to generate" })}\n\n`);
      res.end();
    }
  })
);

// download-notices
router.post(
  "/download-notices",
  requireAuth,
  requireRole("coordinator"),
  (req, res, next) => {
    // Give more time — notices can still be slow if many students
    req.setTimeout(600000);  // 10 minutes
    res.setTimeout(600000);
    next();
  },
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    const preview = await previewPromotion(programId, yearToPromote, academicYearName);
    const program = await Program.findById(programId).lean();

    const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
    let logoBuffer = Buffer.alloc(0);
    if (fs.existsSync(logoPath)) {
      logoBuffer = fs.readFileSync(logoPath);
    }

    if (preview.blocked.length === 0) {
      return res.status(400).json({ error: "No ineligible students found" });
    }

    // Optional safety: limit to avoid server overload
    const MAX_NOTICES = 150;
    const blockedToProcess = preview.blocked.slice(0, MAX_NOTICES);

    const zip = new AdmZip();

    const noticeData = {
      programName: program?.name || "Unknown Program",
      academicYear: academicYearName,
      yearOfStudy: yearToPromote,
      logoBuffer
    };

    for (const student of blockedToProcess) {
      const noticeBuffer = await generateIneligibilityNotice(student, noticeData);
      zip.addFile(`Ineligibility_Notices/${student.regNo}_Notice.docx`, noticeBuffer);
    }

    if (preview.blocked.length > MAX_NOTICES) {
      // Optional: add a text file explaining the limit
      zip.addFile("README.txt", Buffer.from(
        `Only the first ${MAX_NOTICES} notices were generated.\n` +
        `Total ineligible students: ${preview.blocked.length}\n` +
        `Contact system admin for the remaining notices.`
      ));
    }

    const zipBuffer = zip.toBuffer();

    res
      .header("Content-Type", "application/zip")
      .header("Content-Disposition", `attachment; filename=Ineligibility_Notices_${program?.code}_Y${yearToPromote}.zip`)
      .send(zipBuffer);
  })
);


// download-notices-progress
router.post(
  "/download-notices-progress",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    const sendProgress = (percent: number, message: string, file?: string) => {
      res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
    };

    try {
      sendProgress(5, "Analyzing exam records...");
      const preview = await previewPromotion(programId, yearToPromote, academicYearName);
      const program = await Program.findById(programId).lean();

      const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
      const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

      const zip = new AdmZip();
      const students = preview.blocked;

      for (let i = 0; i < students.length; i++) {
        const student = students[i];

const statusText = (student.status || "").toUpperCase();
  const hasSpecialReason = student.reasons?.some((r: string) => 
    r.toUpperCase().includes("SPECIAL")
  );

  const isSpecialCase = statusText.includes("SPECIAL") || hasSpecialReason;

        let docBuffer: Buffer;
        let fileName: string;

        if (isSpecialCase) {
            docBuffer = await generateSpecialExamNotice(student, {
            programName: program?.name || "Program",
            academicYear: academicYearName,
            logoBuffer
          });
          fileName = `SPECIAL_NOTICE_${student.regNo}.docx`;
        } else {
          docBuffer = await generateIneligibilityNotice(student, {
            programName: program?.name || "Program",
            academicYear: academicYearName,
            yearOfStudy: yearToPromote,
            logoBuffer
          });
          fileName = `FAIL_NOTICE_${student.regNo}.docx`;
        }

        zip.addFile(fileName, docBuffer);

        if (i % 5 === 0 || i === students.length - 1) {
          const percent = Math.floor((i / students.length) * 85) + 10;
          sendProgress(percent, `Processing ${i + 1} of ${students.length}: ${student.regNo}`);
        }
      }

      sendProgress(95, "Packing ZIP archive...");
      const zipBase64 = zip.toBuffer().toString('base64');
      sendProgress(100, "Complete!", zipBase64);
      res.end();
    } catch (err: any) {
      res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
      res.end();
    }
  })
);

// download-transcripts-progress
router.post(
  "/download-transcripts-progress",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName, studentId } = req.body;

    console.log("--- ROUTE DEBUG START ---");
    console.log("INCOMING_BODY_YEAR_NAME:", academicYearName);
    console.log("INCOMING_BODY_STUDENT_ID:", studentId);

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    const sendProgress = (percent: number, message: string, file?: string) => {
      res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
    };

    try {
      let targetStudents: any[] = [];

      if (studentId) {
        // SINGLE STUDENT MODE
        sendProgress(10, "Fetching student record...");
        const student = await Student.findById(studentId).lean();
        if (!student) throw new Error("Student not found.");
        
        targetStudents = [{
          id: student._id,
          regNo: student.regNo,
          name: student.name,
          program: student.program
        }];
      } else {
        // BULK MODE (Existing Logic)
        sendProgress(5, "Filtering eligible students...");
        const preview = await previewPromotion(programId, yearToPromote, academicYearName);
        targetStudents = preview.eligible;
      }

      if (targetStudents.length === 0) {
        throw new Error("No eligible students found.");
      }
    

      const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
      const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

      const zip = new AdmZip();

      for (let i = 0; i < targetStudents.length; i++) {
        const student = targetStudents[i];
        
        // Use student's own program if programId wasn't provided (single mode)
        const activeProgramId = programId || student.program;

        const statusResult = await calculateStudentStatus(
            student.id, 
            activeProgramId, 
            academicYearName, 
            yearToPromote
        );

        if (!statusResult) continue;

        // DEBUG LOG FOR ENGINE RESULT
        console.log(`ENGINE_RESULT FOR ${student.regNo}:`, statusResult.academicYearName);

      const program = await Program.findById(activeProgramId).lean();
        
      // Use the resolved year name from statusResult if the body's version is missing
      const finalYear = (academicYearName && academicYearName !== "N/A") 
  ? academicYearName 
  : statusResult.academicYearName;
        console.log(`FINAL_SENDING_TO_GENERATOR:`, finalYear);

        // passedList now contains {code, name, grade} objects from our previous fix
        const transcriptBuffer = await generateStudentTranscript(student, statusResult.passedList, {
          programName: program?.name || "Program",
          academicYear: finalYear,
          yearToPromote: yearToPromote,
          logoBuffer
        });

        const safeRegNo = student.regNo.replace(/\//g, '_');
        zip.addFile(`TRANSCRIPT_${safeRegNo}.docx`, transcriptBuffer);

       sendProgress(Math.floor(((i + 1) / targetStudents.length) * 80) + 10, `Processing ${student.regNo}...`);
      
      }

      console.log("--- ROUTE DEBUG END ---");

      sendProgress(95, "Compiling ZIP file...");
      const zipBase64 = zip.toBuffer().toString('base64');
      sendProgress(100, "Complete!", zipBase64);
      res.end();
    } catch (err: any) {
      res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
      res.end();
    }
  })
);

export default router;