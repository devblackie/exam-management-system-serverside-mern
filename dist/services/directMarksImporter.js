"use strict";
// serverside/src/services/directMarksImporter.ts — COMPLETE, PRODUCTION READY
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.importDirectMarksFromBuffer = importDirectMarksFromBuffer;
const node_crypto_1 = require("node:crypto");
const xlsx_1 = __importDefault(require("xlsx"));
const Student_1 = __importDefault(require("../models/Student"));
const AcademicYear_1 = __importDefault(require("../models/AcademicYear"));
const ProgramUnit_1 = __importDefault(require("../models/ProgramUnit"));
const Unit_1 = __importDefault(require("../models/Unit"));
const MarkDirect_1 = __importDefault(require("../models/MarkDirect"));
const gradeCalculator_1 = require("./gradeCalculator");
function stripQualifier(rawRegNo) {
    return rawRegNo.replace(/(\/\d{4})[A-Z][A-Z0-9]*$/i, "$1");
}
function detectAttemptType(rawCell) {
    const raw = (String(rawCell ?? "")).toLowerCase().trim();
    if (!raw)
        return "1st";
    if (raw === "a/s" || raw.startsWith("supp"))
        return "supplementary";
    if (raw === "spec" || raw.includes("special"))
        return "special";
    if (/rp\d+c/i.test(raw) || raw === "a/cf")
        return "re-take";
    if (raw === "a/so" || raw === "a/sos" || raw.includes("stayout"))
        return "re-take";
    if (/rpu\d*/i.test(raw))
        return "re-take";
    if (raw === "b/s" || /a\/ra\d/i.test(raw) || /rp\d+(?!c)/i.test(raw))
        return "1st";
    return "1st";
}
async function importDirectMarksFromBuffer(buffer, filename, req) {
    const institutionId = req.user.institution;
    if (!institutionId)
        throw new Error("Coordinator not linked to institution");
    const batchId = (0, node_crypto_1.randomUUID)();
    const result = { total: 0, success: 0, errors: [], warnings: [] };
    const workbook = xlsx_1.default.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const unitCode = sheet["F12"]?.v?.toString().trim().toUpperCase();
    const yearText = sheet["C8"]?.v?.toString() ?? "";
    const yearMatch = yearText.match(/\d{4}\/\d{4}/);
    const academicYearStr = yearMatch ? yearMatch[0] : null;
    if (!unitCode || !academicYearStr) {
        throw new Error(`Metadata missing. Unit code at F12: "${unitCode}", Academic year at C8: "${academicYearStr}". Check cells F12 and C8.`);
    }
    const [unitDoc, academicYearDoc] = await Promise.all([
        Unit_1.default.findOne({ code: unitCode }).lean(),
        AcademicYear_1.default.findOne({
            year: { $regex: new RegExp(`^${academicYearStr.replace("/", "\\/")}$`, "i") },
            institution: institutionId,
        }).lean(),
    ]);
    if (!unitDoc)
        throw new Error(`Unit "${unitCode}" not found in the database.`);
    if (!academicYearDoc) {
        throw new Error(`Academic Year "${academicYearStr}" not found for this institution.`);
    }
    const rawRows = xlsx_1.default.utils.sheet_to_json(sheet, { header: 1, range: 15 });
    const allProgramUnitsForUnit = await ProgramUnit_1.default.find({
        unit: unitDoc._id,
    }).lean();
    for (const [index, row] of rawRows.entries()) {
        const rowArr = row;
        const rawCell = String(rowArr[1] ?? "").trim().toUpperCase();
        if (!rawCell || rawCell === "REG. NO." || rawCell === "REG NO")
            continue;
        const regNo = stripQualifier(rawCell);
        result.total++;
        const rowNum = index + 16;
        try {
            const student = await Student_1.default.findOne({
                regNo,
                institution: institutionId,
            }).lean();
            if (!student) {
                result.errors.push(`Row ${rowNum} (${regNo}): Student not found.`);
                continue;
            }
            const studentDoc = student;
            const unitDocObj = unitDoc;
            const acadYearObj = academicYearDoc;
            let programUnit = allProgramUnitsForUnit.find((pu) => String(pu.program) ===
                String(studentDoc.program));
            if (!programUnit) {
                const found = await ProgramUnit_1.default.findOne({
                    program: studentDoc.program,
                    unit: unitDocObj._id,
                }).lean();
                if (!found) {
                    result.errors.push(`Row ${rowNum} (${regNo}): Unit "${unitCode}" not linked to this student's curriculum.`);
                    continue;
                }
                programUnit = found;
            }
            const puObj = programUnit;
            const attempt = detectAttemptType(rowArr[3]);
            const isSpecial = attempt === "special";
            const isSupp = attempt === "supplementary";
            const isRetake = attempt === "re-take";
            const caTotal30 = Number(rowArr[4]) || 0;
            const examTotal70 = Number(rowArr[5]) || 0;
            const rawAgreed = rowArr[8];
            const agreedMark = rawAgreed !== undefined && rawAgreed !== null && rawAgreed !== ""
                ? Number(rawAgreed)
                : caTotal30 + examTotal70;
            const externalRaw = rowArr[7];
            const externalTotal100 = externalRaw !== undefined && externalRaw !== null && externalRaw !== ""
                ? Number(externalRaw)
                : null;
            const isMissingCA = caTotal30 === 0 && !isSupp && !isSpecial;
            const markData = {
                institution: institutionId,
                student: studentDoc._id,
                programUnit: puObj._id,
                academicYear: acadYearObj._id,
                batchId,
                caTotal30,
                examTotal70,
                externalTotal100,
                agreedMark,
                attempt,
                isSpecial,
                isSupplementary: isSupp,
                isRetake,
                isMissingCA,
                uploadedBy: req.user._id,
                uploadedAt: new Date(),
                deletedAt: null,
            };
            const saved = await MarkDirect_1.default.findOneAndUpdate({
                student: studentDoc._id,
                programUnit: puObj._id,
                academicYear: acadYearObj._id,
            }, { $set: markData }, { upsert: true, new: true });
            try {
                await (0, gradeCalculator_1.computeFinalGrade)({
                    markId: saved._id,
                });
            }
            catch (gradeErr) {
                const msg = gradeErr instanceof Error ? gradeErr.message : String(gradeErr);
                console.warn(`[directImporter] Row ${rowNum} (${regNo}): grade calc warning — ${msg}`);
                result.warnings.push(`Row ${rowNum} (${regNo}): grade calc — ${msg}`);
            }
            result.success++;
        }
        catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            console.error(`[directImporter] Row ${rowNum} (${regNo}):`, msg);
            result.errors.push(`Row ${rowNum} (${regNo}): ${msg}`);
        }
    }
    return result;
}
