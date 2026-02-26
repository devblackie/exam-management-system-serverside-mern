// serverside/src/routes/promote.ts
import { Router, Response } from "express";
import { requireAuth, requireRole } from "../middleware/auth";
import { asyncHandler } from "../middleware/asyncHandler";
import type { AuthenticatedRequest } from "../middleware/auth";
import { bulkPromoteClass, calculateStudentStatus, previewPromotion, promoteStudent } from "../services/statusEngine";
import Program from "../models/Program";
import {
  generatePromotionWordDoc, generateEligibleSummaryDoc, generateIneligibilityNotice,   PromotionData,
  generateSpecialExamNotice, generateStudentTranscript, generateSupplementaryExamsDoc, generateSpecialExamsDoc,
  generateStayoutExamsDoc, generateAcademicLeaveDoc, generateDeregistrationDoc, generateDiscontinuationDoc, generateRepeatYearDoc,
} from "../utils/promotionReport";
import fs from "fs";
import path from "path";
import AdmZip from "adm-zip";
import Student from "../models/Student";
import { logAudit } from "../lib/auditLogger";
import ProgramUnit from "../models/ProgramUnit";
import Mark from "../models/Mark";
import { generateConsolidatedMarkSheet, ConsolidatedData } from "../utils/consolidatedMS";

const router = Router();

// preview-promotion
router.post( 
  "/preview-promotion", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    if (!programId || !yearToPromote || !academicYearName) return res.status(400).json({ error: "Missing parameters" });
    const previewData = await previewPromotion( programId, yearToPromote, academicYearName );
    res.json({ success: true, data: previewData });
  }),
);

// bulk-promote
router.post(
  "/bulk-promote", requireAuth, requireRole("coordinator"), asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;
    if (!programId || !yearToPromote || !academicYearName) return res.status(400).json({ error: "Missing required promotion parameters" }); 

    const results = await bulkPromoteClass( programId, yearToPromote, academicYearName );
    res.json({ success: true, message: `Process completed: ${results.promoted} promoted, ${results.failed} failed.`, data: results });
  }),
);

// download-report-progress
router.post( "/download-report-progress", requireAuth, asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    res.setHeader("Content-Type", "text/event-stream"); res.setHeader("Cache-Control", "no-cache"); res.setHeader("Connection", "keep-alive");
    const sendProgress = (percent: number, message: string, file?: string) => { const data = JSON.stringify({ percent, message, file }); res.write(`data: ${data}\n\n`); };

    try {
      sendProgress(10, "Fetching student data and raw marks...");
      // 1. Fetch Basic Data
      const preview = await previewPromotion( programId, yearToPromote, academicYearName );
      const program = await Program.findById(programId).lean();
      const logoPath = path.join( __dirname, "../../public/institutionLogoExcel.png", );
      const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

      // 3. Fetch Marks
      const allStudents = [...preview.eligible, ...preview.blocked];
      const studentIds = allStudents.map((s) => { const id = s._id || s.student?._id || s.id || s.student; return id?.toString(); }).filter((id) => id && id.length >= 24); 
      const rawMarks = await Mark.find({ student: { $in: studentIds } }).populate({ path: "programUnit", populate: { path: "unit", select: "code name" }}).lean();
      const filteredMarks = rawMarks.filter((m: any) => { return m.programUnit && Number(m.programUnit.requiredYear) === Number(yearToPromote); });
      const offeredUnitsRaw = await ProgramUnit.find({ program: programId, requiredYear: yearToPromote }).populate("unit").lean();
      const offeredUnits = offeredUnitsRaw.map((pu: any) => ({ code: pu.unit?.code || "N/A", name: pu.unit?.name || "N/A" }));

      // 5. Prepare Data Objects (Separated for type safety)
      const wordData: PromotionData = {
        programName: program?.name || "Program", academicYear: academicYearName, yearOfStudy: yearToPromote,
        eligible: preview.eligible, blocked: preview.blocked, offeredUnits, logoBuffer,
      };

      const excelData: ConsolidatedData = {
        programName: program?.name || "Program", academicYear: academicYearName, yearOfStudy: yearToPromote,
        students: [...preview.eligible, ...preview.blocked], marks: filteredMarks, offeredUnits, logoBuffer,
      };

      // 6. Generate and Zip reports
      const cleanAcadYear = academicYearName.replace(/\//g, "_");
      const progCode = program?.code || "PROG";
      const progName = program?.name || "Program";
      const yearPrefix = `Year_${yearToPromote}`;

      const getFileName = (reportType: string) => 
        `${reportType}_${progCode}_${progName}_${cleanAcadYear}_${yearPrefix}.docx`.replace(/\s+/g, "_");
      const zip = new AdmZip();

      // Helper to conditionally add documents
      const addDocIfNotEmpty = async (
        list: any[], fileName: string, generator: (data: any, ...args: any[]) => Promise<Buffer>, ...extraArgs: any[] ) => {
        if (list && list.length > 0) { 
          const buffer = await generator(wordData, ...extraArgs);
          zip.addFile(fileName, buffer);
          return true; }
        return false;
      };
     
      sendProgress(30, "Generating Main Summary & Marksheet...");
      zip.addFile(`Summary_Ordinary_Exams_${progCode}_${progName}_${yearPrefix}_${cleanAcadYear}.docx`, await generatePromotionWordDoc(wordData));
      zip.addFile(`${progName}__${progCode}_${cleanAcadYear}_${yearPrefix}_CMS.xlsx`, await generateConsolidatedMarkSheet(excelData));

      sendProgress(40, "Checking Pass List...");
      await addDocIfNotEmpty(wordData.eligible, getFileName("PASS_LIST"), generateEligibleSummaryDoc);

      sendProgress(50, "Checking Supplementary List...");
      const suppList = wordData.blocked.filter(s => s.status === "SUPPLEMENTARY");
      await addDocIfNotEmpty(suppList, getFileName("Supplementary_List"), generateSupplementaryExamsDoc);

      sendProgress(60, "Checking Special Exams...");
      const finSpecials = wordData.blocked.filter(s => s.reasons?.some((r: String) => r.toLowerCase().includes("special") && r.toLowerCase().includes("financial")));
      await addDocIfNotEmpty(finSpecials, getFileName("Special_Financial"), generateSpecialExamsDoc, "Financial");

      const compSpecials = wordData.blocked.filter(s => s.reasons?.some((r: String) => r.toLowerCase().includes("special") && r.toLowerCase().includes("compassionate")));
      await addDocIfNotEmpty(compSpecials, getFileName("SpecialExams_Compassionate"), generateSpecialExamsDoc, "Compassionate");

      sendProgress(70, "Checking Stayout & Repeat Year...");
      const stayoutList = wordData.blocked.filter(s => s.status === "STAYOUT");
      await addDocIfNotEmpty(stayoutList, getFileName("Stayout_Retake_List"), generateStayoutExamsDoc);

      const repeatList = wordData.blocked.filter(s => s.status === "REPEAT YEAR");
      await addDocIfNotEmpty(repeatList, getFileName("Repeat_Year_List"), generateRepeatYearDoc);

      sendProgress(80, "Checking Discontinuations & Deregistrations...");
      const discoList = wordData.blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED");
      await addDocIfNotEmpty(discoList, getFileName("Discontinuation_List"), generateDiscontinuationDoc);

      const deregList = wordData.blocked.filter(s => s.status === "DEREGISTERED");
      await addDocIfNotEmpty(deregList, getFileName("Deregistration_List"), generateDeregistrationDoc);

      sendProgress(90, "Checking Administrative Statuses...");
      const leaveList = wordData.blocked.filter(s => s.status === "ACADEMIC LEAVE" || s.status === "DEFERMENT");
      await addDocIfNotEmpty( leaveList, getFileName("Academic_Leave_Deferment"), generateAcademicLeaveDoc);

      // 7. Zip and Send
      sendProgress(95, "Creating ZIP Archive...");
      const zipBase64 = zip.toBuffer().toString("base64");
      res.write(`data: ${JSON.stringify({ percent: 100, message: "Complete!", file: zipBase64 })}\n\n`);
      res.end();

    } catch (err) {
      console.error("Report Generation Error:", err);
      res.write(`data: ${JSON.stringify({ error: "Failed to generate" })}\n\n`);
      res.end();
    }
  }),
);

// download-report
// router.post(
//   "/download-report",
//   requireAuth,
//   requireRole("coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { programId, yearToPromote, academicYearName } = req.body;

//     // 1. SET HEADERS FOR STREAMING
//     // This tells the browser/proxy NOT to buffer the response
//     res.setHeader('Content-Type', 'application/zip');
//     res.setHeader('Content-Disposition', `attachment; filename=Promotion_Package_Year_${yearToPromote}.zip`);
//     res.setHeader('X-Content-Type-Options', 'nosniff');

//     try {
//       // 2. FETCH DATA
//       const preview = await previewPromotion(programId, yearToPromote, academicYearName);
//       const program = await Program.findById(programId).lean();

//       // SEND HEARTBEAT (A single space keeps the socket open)
//       res.write(" ");

//       const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
//       let logoBuffer = Buffer.alloc(0);
//       if (fs.existsSync(logoPath)) {
//         logoBuffer = fs.readFileSync(logoPath);
//       }

//       const promotionData: PromotionData = {
//         programName: program?.name || "Unknown Program",
//         academicYear: academicYearName,
//         yearOfStudy: yearToPromote,
//         eligible: preview.eligible,
//         blocked: preview.blocked,
//         logoBuffer
//       };

//       // 3. GENERATE DOCUMENTS WITH PULSES
//       const mainBuffer = await generatePromotionWordDoc(promotionData);
//       res.write(" "); // Pulse

//       const eligibleBuffer = await generateEligibleSummaryDoc(promotionData);
//       res.write(" "); // Pulse

//       const ineligibleBuffer = await generateIneligibleSummaryDoc(promotionData);
//       res.write(" "); // Pulse

//       // 4. CREATE ZIP
//       const zip = new AdmZip();
//       zip.addFile(`Promotion_Report_${program?.code}_Year${yearToPromote}.docx`, mainBuffer);
//       zip.addFile(`Eligible_Students_${program?.code}_Year${yearToPromote}.docx`, eligibleBuffer);
//       zip.addFile(`Ineligible_Students_${program?.code}_Year${yearToPromote}.docx`, ineligibleBuffer);

//       const zipBuffer = zip.toBuffer();

//       // 5. FINAL SEND
//       // We use .end() because we used .write() earlier
//       res.write(zipBuffer);
//       res.end();

//       console.log("[download-report] Streaming complete");
//     } catch (err: any) {
//       console.error("[download-report] CRASH:", err);
//       // If we already sent headers/pulses, we can't send a JSON error
//       if (!res.headersSent) {
//         res.status(500).json({ error: "Failed to generate report" });
//       } else {
//         res.end();
//       }
//     }
//   })
// );

// download-notices
// router.post(
//   "/download-notices",
//   requireAuth,
//   requireRole("coordinator"),
//   (req, res, next) => {
//     // Give more time — notices can still be slow if many students
//     req.setTimeout(600000);  // 10 minutes
//     res.setTimeout(600000);
//     next();
//   },
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { programId, yearToPromote, academicYearName } = req.body;

//     const preview = await previewPromotion(programId, yearToPromote, academicYearName);
//     const program = await Program.findById(programId).lean();

//     const logoPath = path.join(__dirname, "../../public/institutionLogoExcel.png");
//     let logoBuffer = Buffer.alloc(0);
//     if (fs.existsSync(logoPath)) {
//       logoBuffer = fs.readFileSync(logoPath);
//     }

//     if (preview.blocked.length === 0) {
//       return res.status(400).json({ error: "No ineligible students found" });
//     }

//     // Optional safety: limit to avoid server overload
//     const MAX_NOTICES = 150;
//     const blockedToProcess = preview.blocked.slice(0, MAX_NOTICES);

//     const zip = new AdmZip();

//     const noticeData = {
//       programName: program?.name || "Unknown Program",
//       academicYear: academicYearName,
//       yearOfStudy: yearToPromote,
//       logoBuffer
//     };

//     for (const student of blockedToProcess) {
//       const noticeBuffer = await generateIneligibilityNotice(student, noticeData);
//       zip.addFile(`Ineligibility_Notices/${student.regNo}_Notice.docx`, noticeBuffer);
//     }

//     if (preview.blocked.length > MAX_NOTICES) {
//       // Optional: add a text file explaining the limit
//       zip.addFile("README.txt", Buffer.from(
//         `Only the first ${MAX_NOTICES} notices were generated.\n` +
//         `Total ineligible students: ${preview.blocked.length}\n` +
//         `Contact system admin for the remaining notices.`
//       ));
//     }

//     const zipBuffer = zip.toBuffer();

//     res
//       .header("Content-Type", "application/zip")
//       .header("Content-Disposition", `attachment; filename=Ineligibility_Notices_${program?.code}_Y${yearToPromote}.zip`)
//       .send(zipBuffer);
//   })
// );

// download-notices-progress
router.post(
  "/download-notices-progress",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { programId, yearToPromote, academicYearName } = req.body;

    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");

    const sendProgress = (percent: number, message: string, file?: string) => {
      res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
    };

    try {
      sendProgress(5, "Analyzing exam records...");
      const preview = await previewPromotion( programId, yearToPromote, academicYearName );
      const program = await Program.findById(programId).lean();

      const logoPath = path.join(  __dirname, "../../public/institutionLogoExcel.png", );
      const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : Buffer.alloc(0);

      const zip = new AdmZip();
      const students = preview.blocked;

      for (let i = 0; i < students.length; i++) {
        const student = students[i];

        const statusText = (student.status || "").toUpperCase();
        const hasSpecialReason = student.reasons?.some((r: string) => r.toUpperCase().includes("SPECIAL"));

        const isSpecialCase = statusText.includes("SPECIAL") || hasSpecialReason;

        let docBuffer: Buffer;
        let fileName: string;

        if (isSpecialCase) {
          docBuffer = await generateSpecialExamNotice(student, { programName: program?.name || "Program", academicYear: academicYearName, logoBuffer });
          fileName = `SPECIAL_NOTICE_${student.regNo}.docx`;
        } else {
          docBuffer = await generateIneligibilityNotice(student, { programName: program?.name || "Program", academicYear: academicYearName, yearOfStudy: yearToPromote, logoBuffer });
          fileName = `FAIL_NOTICE_${student.regNo}.docx`;
        }

        zip.addFile(fileName, docBuffer);

        if (i % 5 === 0 || i === students.length - 1) {
          const percent = Math.floor((i / students.length) * 85) + 10;
          sendProgress( percent, `Processing ${i + 1} of ${students.length}: ${student.regNo}`);
        }
      }

      sendProgress(95, "Packing ZIP archive...");
      const zipBase64 = zip.toBuffer().toString("base64");
      sendProgress(100, "Complete!", zipBase64);
      res.end();
    } catch (err: any) {
      res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
      res.end();
    }
  }),
);

// download-transcripts-progress
// router.post(
//   "/download-transcripts-progress",
//   requireAuth,
//   requireRole("coordinator"),
//   asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
//     const { programId, yearToPromote, academicYearName, studentId } = req.body;

//     // console.log("--- ROUTE DEBUG START ---");
//     // console.log("INCOMING_BODY_YEAR_NAME:", academicYearName);
//     // console.log("INCOMING_BODY_STUDENT_ID:", studentId);

//     res.setHeader("Content-Type", "text/event-stream");
//     res.setHeader("Cache-Control", "no-cache");
//     res.setHeader("Connection", "keep-alive");

//     const sendProgress = (percent: number, message: string, file?: string) => {
//       res.write(`data: ${JSON.stringify({ percent, message, file })}\n\n`);
//     };

//     try {
//       let targetStudents: any[] = [];

//       if (studentId) {
//         // SINGLE STUDENT MODE
//         sendProgress(10, "Fetching student record...");
//         const student = await Student.findById(studentId).lean();
//         if (!student) throw new Error("Student not found.");

//         targetStudents = [
//           {
//             id: student._id,
//             regNo: student.regNo,
//             name: student.name,
//             program: student.program,
//           },
//         ];
//       } else {
//         // BULK MODE (Existing Logic)
//         sendProgress(5, "Filtering eligible students...");
//         const preview = await previewPromotion(
//           programId,
//           yearToPromote,
//           academicYearName,
//         );
//         targetStudents = preview.eligible;
//       }

//       if (targetStudents.length === 0) {
//         throw new Error("No eligible students found.");
//       }

//       const logoPath = path.join(
//         __dirname,
//         "../../public/institutionLogoExcel.png",
//       );
//       const logoBuffer = fs.existsSync(logoPath)
//         ? fs.readFileSync(logoPath)
//         : Buffer.alloc(0);

//       const zip = new AdmZip();

//       for (let i = 0; i < targetStudents.length; i++) {
//         const student = targetStudents[i];

//         // Use student's own program if programId wasn't provided (single mode)
//         const activeProgramId = programId || student.program;

//         const statusResult = await calculateStudentStatus(
//           student.id,
//           activeProgramId,
//           academicYearName,
//           yearToPromote,
//         );

//         if (!statusResult) continue;

//         // DEBUG LOG FOR ENGINE RESULT
//         // console.log(`ENGINE_RESULT FOR ${student.regNo}:`, statusResult.academicYearName);

//         const program = await Program.findById(activeProgramId).lean();

//         // Use the resolved year name from statusResult if the body's version is missing
//         const finalYear =
//           academicYearName && academicYearName !== "N/A"
//             ? academicYearName
//             : statusResult.academicYearName;
//         // console.log(`FINAL_SENDING_TO_GENERATOR:`, finalYear);

//         // passedList now contains {code, name, grade} objects from our previous fix
//         const transcriptBuffer = await generateStudentTranscript(
//           student,
//           statusResult.passedList,
//           {
//             programName: program?.name || "Program",
//             academicYear: finalYear,
//             yearToPromote: yearToPromote,
//             logoBuffer,
//           },
//         );

//         const safeRegNo = student.regNo.replace(/\//g, "_");
//         zip.addFile(`TRANSCRIPT_${safeRegNo}.docx`, transcriptBuffer);

//         sendProgress(
//           Math.floor(((i + 1) / targetStudents.length) * 80) + 10,
//           `Processing ${student.regNo}...`,
//         );
//       }

//       // console.log("--- ROUTE DEBUG END ---");

//       sendProgress(95, "Compiling ZIP file...");
//       const zipBase64 = zip.toBuffer().toString("base64");
//       sendProgress(100, "Complete!", zipBase64);
//       res.end();
//     } catch (err: any) {
//       res.write(`data: ${JSON.stringify({ error: err.message })}\n\n`);
//       res.end();
//     }
//   }),
// );

// promote (individual student)
router.post(
  "/:studentId",
  requireAuth,
  requireRole("coordinator"),
  asyncHandler(async (req: AuthenticatedRequest, res: Response) => {
    const { studentId } = req.params;

    const result = await promoteStudent(studentId);

    if (!result.success) {
      return res.status(400).json({
        error: "Promotion Denied",
        message: result.message,
        details: result.details,
      });
    }

    await logAudit(req, {
      action: "individual_student_promoted",
      targetUser: studentId as any,
      details: { message: result.message },
    });

    res.json(result);
  }),
);

export default router;

			

																
