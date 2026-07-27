"use strict";
// serverside/src/services/marksImporter.ts — COMPLETE, PRODUCTION READY
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.importMarksFromBuffer = importMarksFromBuffer;
const node_crypto_1 = require("node:crypto");
const xlsx_1 = __importDefault(require("xlsx"));
const Student_1 = __importDefault(require("../models/Student"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const Mark_1 = __importDefault(require("../models/Mark"));
const gradeCalculator_1 = require("./gradeCalculator");
function stripQualifier(rawRegNo) {
    return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
}
async function importMarksFromBuffer(buffer, filename, req) {
    const institutionId = req.user.institution;
    if (!institutionId)
        throw new Error("Coordinator not linked to institution");
    const batchId = (0, node_crypto_1.randomUUID)();
    const result = { total: 0, success: 0, errors: [], warnings: [] };
    const workbook = xlsx_1.default.read(buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const modeIndicator = (sheet["E10"]?.v?.toString() ?? "").toUpperCase();
    let detectedUnitType = "theory";
    if (modeIndicator.includes("LAB"))
        detectedUnitType = "lab";
    if (modeIndicator.includes("WORKSHOP"))
        detectedUnitType = "workshop";
    const unitCode = sheet["H12"]?.v?.toString().trim().toUpperCase();
    const academicYearText = sheet["D8"]?.v?.toString() ?? "";
    const yearMatch = academicYearText.match(/\d{4}\/\d{4}/);
    const academicYearStr = yearMatch ? yearMatch[0] : null;
    if (!unitCode || !academicYearStr) {
        throw new Error(`Invalid Template: Missing Unit Code (H12) or Academic Year (D8). Found Unit: ${unitCode}, Year: ${academicYearStr}`);
    }
    const rawRows = xlsx_1.default.utils.sheet_to_json(sheet, { header: 1, range: 16 });
    const academicYearDoc = await AcademicYear_1.default.findOne({
        year: { $regex: new RegExp(`^${academicYearStr}$`, "i") },
        institution: institutionId,
    }).lean();
    if (!academicYearDoc) {
        throw new Error(`Academic Year '${academicYearStr}' not found.`);
    }
    const acadYearObj = academicYearDoc;
    const allProgramUnits = await ProgramUnit_1.default.find({ institution: institutionId })
        .populate("unit")
        .lean();
    const programUnitMap = new Map(allProgramUnits.map((pu) => {
        const puObj = pu;
        const unitObj = puObj.unit;
        return [
            `${String(puObj.program)}_${String(unitObj?.code ?? "").toUpperCase()}`,
            puObj,
        ];
    }));
    for (const [index, row] of rawRows.entries()) {
        const rowArr = row;
        const rawCell = String(rowArr[1] ?? "").trim().toUpperCase();
        const sn = rowArr[0];
        if (!rawCell || sn === "")
            continue;
        const regNo = stripQualifier(rawCell);
        result.total++;
        const rowNum = index + 17;
        try {
            const student = await Student_1.default.findOne({
                regNo,
                institution: institutionId,
            }).lean();
            if (!student) {
                result.errors.push(`Row ${rowNum} (${regNo}): Student not found.`);
                continue;
            }
            const studentObj = student;
            const programUnitKey = `${String(studentObj.program)}_${unitCode}`;
            const programUnit = programUnitMap.get(programUnitKey);
            if (!programUnit) {
                result.errors.push(`Row ${rowNum} (${regNo}): Unit "${unitCode}" not linked to student's program.`);
                continue;
            }
            const excelAttemptLabel = String(rowArr[3] ?? "").trim().toLowerCase();
            let finalAttempt = "1st";
            if (excelAttemptLabel.includes("supp")) {
                finalAttempt = "supplementary";
            }
            else if (excelAttemptLabel.includes("special")) {
                finalAttempt = "special";
            }
            else if (excelAttemptLabel.includes("retake")) {
                finalAttempt = "re-take";
            }
            const markData = {
                student: studentObj._id,
                programUnit: programUnit._id,
                academicYear: acadYearObj._id,
                institution: institutionId,
                uploadedBy: req.user._id,
                deletedAt: null,
                batchId,
                cat1Raw: Number(rowArr[4]) || 0,
                cat2Raw: Number(rowArr[5]) || 0,
                cat3Raw: Number(rowArr[6]) || 0,
                assgnt1Raw: Number(rowArr[8]) || 0,
                assgnt2Raw: Number(rowArr[9]) || 0,
                assgnt3Raw: Number(rowArr[10]) || 0,
                practicalRaw: Number(rowArr[12]) || 0,
                examQ1Raw: Number(rowArr[14]) || 0,
                examQ2Raw: Number(rowArr[15]) || 0,
                examQ3Raw: Number(rowArr[16]) || 0,
                examQ4Raw: Number(rowArr[17]) || 0,
                examQ5Raw: Number(rowArr[18]) || 0,
                caTotal30: Number(rowArr[13]) || 0,
                examTotal70: Number(rowArr[19]) || 0,
                internalExaminerMark: Number(rowArr[20]) || 0,
                agreedMark: Number(rowArr[22]) || 0,
                attempt: finalAttempt,
                isSpecial: finalAttempt === "special",
                isSupplementary: finalAttempt === "supplementary",
                isRetake: finalAttempt === "re-take",
                unitType: detectedUnitType,
                examMode: sheet["O16"]?.v === 30 ? "mandatory_q1" : "standard",
            };
            const mark = await Mark_1.default.findOneAndUpdate({
                student: studentObj._id,
                programUnit: programUnit._id,
                academicYear: acadYearObj._id,
            }, markData, { upsert: true, new: true, runValidators: true });
            await (0, gradeCalculator_1.computeFinalGrade)({ markId: mark._id });
            result.success++;
        }
        catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            console.error(`[marksImporter] Row ${rowNum} (${regNo}):`, msg);
            result.errors.push(`Row ${rowNum} (${regNo}): ${msg}`);
        }
    }
    return result;
}
