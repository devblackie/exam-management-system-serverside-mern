"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateStudentTranscript = exports.generateSpecialExamNotice = exports.generateIneligibilityNotice = exports.generateSimpleAwardListDoc = exports.generateAcademicLeaveSuppDoc = exports.generateReadmissionSuppDoc = exports.generateRepeatSuppDoc = exports.generateStayoutSuppDoc = exports.generateCarryForwardSuppDoc = exports.generateSupplementaryExamsDoc = exports.generateDefermentDoc = exports.generateAwardListDoc = exports.generateDeregistrationDoc = exports.generateDiscontinuationDoc = exports.generateCarryForwardDoc = exports.generateAcademicLeaveDoc = exports.generateRepeatYearDoc = exports.generateStayoutExamsDoc = exports.generateIncompleteListDoc = exports.generateSpecialExamsDoc = exports.generateEligibleSummaryDoc = exports.generatePromotionWordDoc = void 0;
// serverside/src/utils/promotionReport.ts
const docx_1 = require("docx");
const config_1 = __importDefault(require("../config/config"));
const loadInstitutionSettings_1 = require("./loadInstitutionSettings");
const loadLogoBuffer_1 = require("./loadLogoBuffer");
const numberToWords = (num) => {
    const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
    const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
    if (num === 0)
        return "Zero";
    if (num < 20)
        return ones[num];
    const digit = num % 10;
    return tens[Math.floor(num / 10)] + (digit ? "-" + ones[digit] : "");
};
// Helper to convert Year 1 to "First Year", etc.
const getOrdinalYear = (year) => {
    const ordinals = ["", "First", "Second", "Third", "Fourth", "Fifth", "Sixth"];
    return ordinals[year] || `${year}th`;
};
const formatStudentName = (fullName) => {
    if (!fullName)
        return "";
    const parts = fullName.trim().split(/\s+/);
    // If only one name exists, just uppercase it
    if (parts.length <= 1)
        return fullName.toUpperCase();
    // Remove the last name, uppercase it, then join back
    const lastName = parts.pop()?.toUpperCase();
    return `${parts.join(" ")} ${lastName}`;
};
// Cache meta per function call (loaded once per report generation)
async function getDocContext(data) {
    const [settings, logoBuffer] = await Promise.all([
        (0, loadInstitutionSettings_1.loadInstitutionSettings)(data.institutionId),
        (0, loadLogoBuffer_1.loadLogoBuffer)(data.institutionId),
    ]);
    return { meta: settings.docMeta, logoBuffer };
}
function createDocHeader(logo, program, year, ordinal, listType, meta, examType = "ORDINARY") {
    return [
        ...(logo && logo.length > 0
            ? [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [
                        new docx_1.ImageRun({ data: logo, transformation: { width: 120, height: 70 }, type: "png" })
                    ] })]
            : []),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 100, after: 100 },
            children: [new docx_1.TextRun({ text: meta.universityName.toUpperCase(), bold: true, size: 23 })] }),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER,
            children: [new docx_1.TextRun({ text: meta.schoolName.toUpperCase(), bold: true, size: 23 })] }),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER,
            children: [new docx_1.TextRun({ text: meta.departmentName.toUpperCase(), bold: true, size: 23 })] }),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER,
            children: [new docx_1.TextRun({ text: `PROGRAM: ${program.toUpperCase()}`, bold: true, size: 23 })] }),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER,
            children: [new docx_1.TextRun({ text: `${year} ACADEMIC YEAR`, bold: true, size: 23 })] }),
        ...(examType !== undefined && !listType.toUpperCase().includes("AWARD LIST")
            ? [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 40, after: 40 },
                    children: [new docx_1.TextRun({
                            text: examType === "SUPPLEMENTARY"
                                ? "SUPPLEMENTARY AND SPECIAL EXAMINATION RESULTS"
                                : "ORDINARY EXAMINATION RESULTS",
                            bold: true, size: 24,
                        })] })]
            : []),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 100, after: 100 },
            children: [new docx_1.TextRun({ text: `${ordinal} Year`, bold: true, size: 23 })] }),
        new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 100, after: 100 },
            children: [new docx_1.TextRun({ text: listType, bold: true, size: 23, underline: {} })] }),
    ];
}
function createDocFooter(meta) {
    return [
        new docx_1.Paragraph({ spacing: { before: 900 }, children: [new docx_1.TextRun({ text: `APPROVED BY THE BOARD OF EXAMINERS, ${meta.schoolName.toUpperCase()}`, bold: true, size: 18 })] }),
        new docx_1.Paragraph({ spacing: { before: 400 }, children: [new docx_1.TextRun({ text: "SIGNED: __________________________\t\tDATE: _______________", bold: true })] }),
        new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: `\tDEAN, ${meta.schoolName.toUpperCase()}`, size: 18 })] }),
    ];
}
const generatePromotionWordDoc = async (data) => {
    const { meta, logoBuffer } = await getDocContext(data);
    const { programName, academicYear, yearOfStudy, eligible, blocked, offeredUnits = [] } = data;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const stats = {
        "PASS": eligible.length,
        "SUPPLEMENTARY (After Readmission)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r) => r.toLowerCase().includes("readmission"))).length,
        "SUPPLEMENTARY (After Stayout)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r) => r.toLowerCase().includes("stayout"))).length,
        "SUPPLEMENTARY (After Carry Forward)": blocked.filter(s => s.status === "SUPPLEMENTARY" && s.reasons?.some((r) => r.toLowerCase().includes("carry forward"))).length,
        "DEFERMENT": blocked.filter(s => s.status === "DEFERMENT" || s.status === "DEFERRED").length,
        "STAYOUT": blocked.filter(s => s.status === "STAYOUT").length,
        "DISCONTINUATION": blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED").length,
        "DEREGISTRATION": blocked.filter(s => s.status === "DEREGISTERED").length,
        "REPEAT YEAR": blocked.filter(s => s.status === "REPEAT YEAR").length,
        "INCOMPLETE": blocked.filter(s => s.status.includes("INC") && !s.status.includes("SUPP") && !s.status.includes("SPEC")).length,
        "SUPPLEMENTARY": blocked.filter(s => s.status.includes("SUPP") && !s.status.includes("SPEC")).length,
        "ACADEMIC LEAVE (FINANCIAL)": blocked.filter(s => (s.status === "ACADEMIC LEAVE" || s.status === "ON LEAVE") && (s.academicLeavePeriod?.type === "financial" || s.remarks?.toLowerCase().includes("financial"))).length,
        "ACADEMIC LEAVE (COMPASSIONATE)": blocked.filter(s => (s.status === "ACADEMIC LEAVE" || s.status === "ON LEAVE") && (s.academicLeavePeriod?.type === "compassionate" || s.remarks?.toLowerCase().includes("compassionate"))).length,
        "SPECIALS (FINANCIAL)": blocked.filter(s => s.status.includes("SPEC") && s.remarks?.toLowerCase().includes("financial")).length,
        "SPECIALS (COMPASSIONATE)": blocked.filter(s => s.status.includes("SPEC") && (s.remarks?.toLowerCase().includes("compassionate") || s.remarks?.toLowerCase().includes("medical"))).length,
        "SPECIALS (OTHER)": blocked.filter(s => s.status.includes("SPEC") && !s.remarks?.toLowerCase().match(/financial|compassionate|medical/)).length,
    };
    const doc = new docx_1.Document({
        sections: [
            {
                properties: {},
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "SUMMARY", meta, data.examType || "ORDINARY"),
                    createSummaryTable(stats),
                    new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 400, after: 100 }, children: [new docx_1.TextRun({ text: "UNITS OFFERED", bold: true, size: 22, underline: {} })] }),
                    createOfferedUnitsTable(offeredUnits),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generatePromotionWordDoc = generatePromotionWordDoc;
function createSummaryTable(stats) {
    const activeRows = Object.entries(stats).filter(([_, val]) => val > 0);
    const totalCount = Object.values(stats).reduce((a, b) => a + b, 0);
    const cellMargin = { top: 50, bottom: 50, left: 100, right: 100 };
    const rows = [
        ...activeRows.map(([label, val]) => new docx_1.TableRow({
            children: [
                new docx_1.TableCell({
                    margins: cellMargin,
                    width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                    borders: { top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE } },
                    children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: label.toUpperCase(), size: 20 })] })]
                }),
                new docx_1.TableCell({
                    margins: cellMargin,
                    width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                    borders: { top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE } },
                    children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: val.toString(), size: 20 })] })]
                }),
            ]
        })),
        new docx_1.TableRow({
            children: [
                new docx_1.TableCell({
                    margins: cellMargin,
                    borders: { top: { style: docx_1.BorderStyle.SINGLE, size: 1 }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE } },
                    children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: "TOTAL", bold: true, size: 20 })] })]
                }),
                new docx_1.TableCell({
                    margins: cellMargin,
                    borders: { top: { style: docx_1.BorderStyle.SINGLE, size: 1 }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE } },
                    children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: totalCount.toString(), bold: true, size: 20 })] })]
                }),
            ]
        })
    ];
    return new docx_1.Table({ width: { size: 60, type: docx_1.WidthType.PERCENTAGE }, alignment: docx_1.AlignmentType.CENTER, rows: rows });
}
function createOfferedUnitsTable(units) {
    if (!units || units.length === 0)
        return new docx_1.Paragraph("No units recorded.");
    const cellMargin = { top: 50, bottom: 50, left: 100, right: 100 };
    const midPoint = Math.ceil(units.length / 2);
    const leftCol = units.slice(0, midPoint);
    const rightCol = units.slice(midPoint);
    const headerCell = (text) => new docx_1.TableCell({
        margins: cellMargin,
        children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.JUSTIFIED, children: [new docx_1.TextRun({ text, bold: true, size: 20 })] })]
    });
    const headerRow = new docx_1.TableRow({
        children: [
            headerCell("S/NO."), headerCell("CODE"), headerCell("NAME"),
            headerCell("S/NO."), headerCell("CODE"), headerCell("NAME"),
        ]
    });
    const dataRows = [];
    for (let i = 0; i < midPoint; i++) {
        const left = leftCol[i];
        const right = rightCol[i];
        dataRows.push(new docx_1.TableRow({
            children: [
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: (i + 1).toString(), size: 21 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: left.code, size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: left.name, size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: right ? (midPoint + i + 1).toString() : "", size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: right?.code || "", size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: right?.name || "", size: 20 })] })] }),
            ]
        }));
    }
    return new docx_1.Table({
        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
        rows: [headerRow, ...dataRows]
    });
}
const generateEligibleSummaryDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, eligible } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const candidateCountWords = numberToWords(eligible.length);
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const nextYearOrdinal = getOrdinalYear(yearOfStudy + 1);
    const cellMargin = { top: 0, bottom: 0, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [
            {
                properties: {},
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "PASS", meta, "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${candidateCountWords} (${eligible.length}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidates satisfied the ${meta.schoolName} Board of Examiners in the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, `, size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they proceed to their `, size: 22 }),
                            new docx_1.TextRun({ text: `${nextYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `of study.`, size: 22 }),
                        ],
                    }),
                    createPassTable(eligible, cellMargin),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateEligibleSummaryDoc = generateEligibleSummaryDoc;
function createPassTable(students, cellMargin) {
    const headerRow = new docx_1.TableRow({
        children: [{ text: "S/No.", width: 5 }, { text: "REG. NO.", width: 30 }, { text: "NAME", width: 65 }].map((col) => new docx_1.TableCell({
            width: { size: col.width, type: docx_1.WidthType.PERCENTAGE },
            margins: cellMargin,
            children: [new docx_1.Paragraph({ spacing: { before: 0, after: 0 }, children: [new docx_1.TextRun({ text: col.text, bold: true, size: 18 })] })],
        })),
    });
    const dataRows = students.map((s, index) => new docx_1.TableRow({
        children: [
            new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ spacing: { before: 0, after: 0 }, children: [new docx_1.TextRun({ text: (index + 1).toString(), size: 20 })] })] }),
            new docx_1.TableCell({ margins: cellMargin, children: [
                    new docx_1.Paragraph({ spacing: { before: 0, after: 0 }, children: [
                            new docx_1.TextRun({ text: s.regNo, size: 20 }),
                            ...(s.qualifierSuffix ? [new docx_1.TextRun({ text: s.qualifierSuffix, size: 16, subScript: true, color: "000000" })] : []),
                        ] })
                ] }),
            new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ spacing: { before: 0, after: 0 }, children: [new docx_1.TextRun({ text: formatStudentName(s.name), size: 20 })] })] })
        ]
    }));
    return new docx_1.Table({
        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE },
            right: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
        },
        rows: [headerRow, ...dataRows],
    });
}
const generateSpecialExamsDoc = async (data, groundType) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const getGrounds = (s) => {
        const g = (s.specialGrounds || "").toLowerCase();
        const r = (s.remarks || "").toLowerCase();
        const lt = (s.academicLeavePeriod?.type || "").toLowerCase();
        const d = (s.details || "").toLowerCase();
        return `${g} ${r} ${lt} ${d}`;
    };
    const isSpecial = (s) => /spec/i.test(s.status);
    const list = blocked.filter((s) => {
        if (!isSpecial(s))
            return false;
        const g = getGrounds(s);
        if (groundType === "Financial")
            return g.includes("financial");
        if (groundType === "Compassionate")
            return /compassionate|medical|sick/.test(g);
        return !g.includes("financial") && !/compassionate|medical|sick/.test(g);
    });
    const count = list.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    console.log(`[generateSpecialExamsDoc] groundType="${groundType}" | found=${count}`, list.map((s) => ({
        regNo: s.regNo,
        status: s.status,
        specialGrounds: s.specialGrounds,
        remarks: s.remarks,
    })));
    const doc = new docx_1.Document({
        sections: [
            {
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, `SPECIAL EXAMINATIONS (${groundType.toUpperCase()} GROUNDS)`, meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "candidate(s) have special examinations, on ", size: 22 }),
                            new docx_1.TextRun({ text: `${groundType} Grounds `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "in the unit(s) indicated against their names during the ", size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "Academic Year, ", size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, size: 22 }),
                            new docx_1.TextRun({ text: "Examinations for the ", size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners upholds the decision of the Dean's Committee. `,
                                size: 22,
                            }),
                        ],
                    }),
                    createSpecialUnitDetailTable(list, cellMargin),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateSpecialExamsDoc = generateSpecialExamsDoc;
function createSpecialUnitDetailTable(students, cellMargin) {
    const headerRow = new docx_1.TableRow({
        children: [
            { text: "S/No", w: 5 },
            { text: "Reg No.", w: 20 },
            { text: "Name", w: 25 },
            { text: "Unit Code", w: 15 },
            { text: "Unit Name", w: 35 },
        ].map((h) => new docx_1.TableCell({
            width: { size: h.w, type: docx_1.WidthType.PERCENTAGE },
            margins: cellMargin,
            children: [
                new docx_1.Paragraph({
                    children: [new docx_1.TextRun({ text: h.text, bold: true, size: 18 })],
                }),
            ],
        })),
    });
    const rows = [headerRow];
    let counter = 1;
    for (const s of students) {
        const specialReasons = (s.reasons || []).filter((r) => r.toUpperCase().includes("SPECIAL"));
        if (specialReasons.length > 0) {
            specialReasons.forEach((reason, idx) => {
                const colonIdx = reason.indexOf(":");
                let uCode = "N/A";
                let uName = "N/A";
                if (colonIdx !== -1) {
                    uCode = reason.substring(0, colonIdx).trim();
                    uName = reason
                        .substring(colonIdx + 1)
                        .split("(")[0]
                        .trim();
                }
                rows.push(new docx_1.TableRow({
                    children: [
                        new docx_1.TableCell({
                            margins: cellMargin,
                            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: idx === 0 ? String(counter) : "", size: 18 })] })],
                        }),
                        new docx_1.TableCell({
                            margins: cellMargin,
                            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: idx === 0 ? s.regNo : "", size: 18 })] })],
                        }),
                        new docx_1.TableCell({
                            margins: cellMargin,
                            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: idx === 0 ? (s.name || "").toUpperCase() : "", size: 18 })] })],
                        }),
                        new docx_1.TableCell({
                            margins: cellMargin,
                            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: uCode, size: 18 })] })],
                        }),
                        new docx_1.TableCell({
                            margins: cellMargin,
                            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: uName, size: 18 })] })],
                        }),
                    ],
                }));
            });
            counter++;
        }
        else {
            console.warn(`[createSpecialUnitDetailTable] Student ${s.regNo} has SPEC status but no special reasons listed.`);
            rows.push(new docx_1.TableRow({
                children: [
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: String(counter), size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.regNo, size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: (s.name || "").toUpperCase(), size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, columnSpan: 2, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: "Special exam — unit details pending confirmation.", size: 18, italics: true })] })] }),
                ],
            }));
            counter++;
        }
    }
    return new docx_1.Table({
        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.NONE },
            bottom: { style: docx_1.BorderStyle.NONE },
            left: { style: docx_1.BorderStyle.NONE },
            right: { style: docx_1.BorderStyle.NONE },
            insideHorizontal: { style: docx_1.BorderStyle.NONE },
            insideVertical: { style: docx_1.BorderStyle.NONE },
        },
        rows,
    });
}
const generateIncompleteListDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const incompleteList = blocked.filter(s => s.status.includes("INC") && !s.status.includes("SPEC"));
    const count = incompleteList.length;
    const candidateCountWords = numberToWords(count);
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "INCOMPLETE", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${candidateCountWords} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) have incomplete results in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year. These results are pending due to missing CATs or Examination marks.`, size: 22 }),
                        ],
                    }),
                    createStandardUnitDetailTable(incompleteList, cellMargin, "INCOMPLETE"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateIncompleteListDoc = generateIncompleteListDoc;
const generateStayoutExamsDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const stayoutList = blocked.filter((s) => s.status === "STAYOUT");
    const count = stayoutList.length;
    const candidateCountWords = numberToWords(count);
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [
            {
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "STAY OUT / RETAKE", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${candidateCountWords} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) failed to satisfy the ${meta.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, `, size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they Stay Out according to `, size: 22 }),
                            new docx_1.TextRun({
                                text: `ENG Rule 15 (h) "A candidate who fails more than a third and less than a half of the prescribed units in any year of study shall be required to retake examinations only in the failed units during the ordinary examination period when examinations for the individual units are offered. Such a candidate will not be allowed to retake examinations during the supplementary period immediately following the ordinary examinations period in which he/she failed the units".`,
                                size: 20, bold: true, italics: true,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(stayoutList, cellMargin),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateStayoutExamsDoc = generateStayoutExamsDoc;
const generateRepeatYearDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status === "REPEAT YEAR");
    const count = list.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "REPEAT YEAR", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) failed to satisfy the ${meta.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, `, size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they Repeat according to `, size: 22 }),
                            new docx_1.TextRun({
                                text: `ENG Rule 16 (c) "A candidate, who attains an average mark of less than 40% in any year of study based on the marks obtained on the 1st attempt for each unit, shall be required to repeat the entire year. Such a candidate will enrol for all the units and sit for all CATs and assignment and the exams will be marked out of 100%. `,
                                size: 20, bold: true, italics: true,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateRepeatYearDoc = generateRepeatYearDoc;
const generateAcademicLeaveDoc = async (data, groundType, type) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const safeType = (type || "ACADEMIC LEAVE").toUpperCase();
    const safeGround = (groundType || "General").toUpperCase();
    const list = (blocked || []).filter((s) => {
        const statusStr = (s.status || "").toUpperCase();
        const isTargetStatus = statusStr.includes(safeType) || statusStr === "ON LEAVE";
        const targetGroundLower = safeGround.toLowerCase();
        const leaveTypeLower = (s.academicLeavePeriod?.type || "").toLowerCase();
        const remarksLower = (s.remarks || "").toLowerCase();
        return (isTargetStatus && (leaveTypeLower === targetGroundLower || remarksLower.includes(targetGroundLower)));
    });
    const formattedList = list.map((s) => ({
        regNo: s.regNo || "N/A",
        name: s.name || "N/A",
        effectiveDate: s.academicLeavePeriod?.startDate ? new Date(s.academicLeavePeriod.startDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "N/A",
        remarks: s.remarks?.includes(":") ? s.remarks.split(":")[1].trim() : s.remarks || "Approved",
    }));
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [
            {
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, `${safeType} (${safeGround} GROUNDS)`, meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following candidate(s) have been officially granted `, size: 22 }),
                            new docx_1.TextRun({ text: `${safeType} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `on `, size: 22 }),
                            new docx_1.TextRun({ text: `${safeGround.toLowerCase()} grounds `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year. They are expected to resume studies at the beginning of the next academic cycle.`, size: 22 }),
                        ],
                    }),
                    createAdministrativeTable(formattedList, cellMargin),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateAcademicLeaveDoc = generateAcademicLeaveDoc;
function createAdministrativeTable(students, cellMargin) {
    return new docx_1.Table({
        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE },
            right: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
        },
        rows: [
            new docx_1.TableRow({
                children: ["S/No", "Reg No.", "Name", "Effective Date", "Remarks"].map((h) => new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: h, bold: true, size: 18 })] })] })),
            }),
            ...students.map((s, i) => new docx_1.TableRow({
                children: [
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.regNo, size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.name, size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.effectiveDate || "N/A", size: 18 })] })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.remarks || "Approved", size: 18 })] })] }),
                ],
            })),
        ],
    });
}
const generateCarryForwardDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, eligible, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const carryForwardList = eligible.filter((s) => s.reasons?.length > 0 && s.status !== "ALREADY PROMOTED");
    const count = carryForwardList.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const nextYearOrdinal = getOrdinalYear(yearOfStudy + 1);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "CARRY FORWARD", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) satisfied the Board of Examiners in at least two-thirds of the units. In accordance with `, size: 22 }),
                            new docx_1.TextRun({ text: `ENG Rule 13 (e)`, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `, they are allowed to proceed to `, size: 22 }),
                            new docx_1.TextRun({ text: `${nextYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `but MUST carry forward the failed units indicated against their names to be taken when next offered.`, size: 22 }),
                        ],
                    }),
                    createStandardUnitDetailTable(carryForwardList, cellMargin),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateCarryForwardDoc = generateCarryForwardDoc;
const generateDiscontinuationDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter(s => s.status === "CRITICAL FAILURE" || s.status === "DISCONTINUED");
    const count = list.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "DISCONTINUATION", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) failed to satisfy the ${meta.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, `, size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they be `, size: 22 }),
                            new docx_1.TextRun({ text: `Discontinued  `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `according to `, size: 22 }),
                            new docx_1.TextRun({
                                text: `ENG Rule 23 (c) "A candidate who fails third but less than half units of a year of study after the first attempt and subsequently fails the same units after retaking the examinations shall be discontinued."  `,
                                size: 20, bold: true, italics: true,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin),
                    ...createDocFooter(meta),
                ]
            }]
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateDiscontinuationDoc = generateDiscontinuationDoc;
const generateDeregistrationDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter(s => s.status === "DEREGISTERED");
    const count = list.length;
    const candidateCountWords = numberToWords(count);
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, getOrdinalYear(yearOfStudy), "DEREGISTRATION", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${candidateCountWords} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) failed to satisfy the ${meta.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, `, size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they be `, size: 22 }),
                            new docx_1.TextRun({ text: `Deregistered `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `according to `, size: 22 }),
                            new docx_1.TextRun({
                                text: `ENG 23 (e) "A candidate who absents himself/herself from all the Special Examinations which he/she was required to sit, or fails to undertake all extra assignments for continuous assessment without good cause, shall be assumed to have deserted the degree course, and shall be deregistered forthwith.  `,
                                size: 20, bold: true, italics: true,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin),
                    ...createDocFooter(meta),
                ]
            }]
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateDeregistrationDoc = generateDeregistrationDoc;
function createStandardUnitDetailTable(students, cellMargin, filterKeyword) {
    const headerRow = new docx_1.TableRow({
        children: [
            { text: "S/No", w: 5 }, { text: "Reg No.", w: 20 }, { text: "Name", w: 25 },
            { text: "Unit Code", w: 15 }, { text: "Unit Name", w: 35 }
        ].map(h => new docx_1.TableCell({
            width: { size: h.w, type: docx_1.WidthType.PERCENTAGE },
            margins: cellMargin,
            children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: h.text, bold: true, size: 18 })] })]
        }))
    });
    const rows = [headerRow];
    let studentCounter = 1;
    students.forEach((s) => {
        const relevantReasons = s.reasons?.filter((r) => {
            const lowerR = r.toLowerCase();
            const isRuleTag = lowerR.startsWith("eng") || lowerR.includes("failures >") || lowerR.includes("failures >=") || lowerR.includes("mean <");
            if (isRuleTag)
                return false;
            if (filterKeyword)
                return lowerR.includes(filterKeyword.toLowerCase());
            return !lowerR.includes("special") && !lowerR.includes("leave");
        }) || [];
        if (relevantReasons.length > 0) {
            relevantReasons.forEach((rawReason, index) => {
                const isFirstUnit = index === 0;
                let uCode = "N/A";
                let uName = "N/A";
                const colonIndex = rawReason.indexOf(":");
                if (colonIndex !== -1) {
                    uCode = rawReason.substring(0, colonIndex).trim();
                    const afterColon = rawReason.substring(colonIndex + 1);
                    uName = afterColon.split("(")[0].trim();
                }
                rows.push(new docx_1.TableRow({
                    children: [
                        new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: isFirstUnit ? studentCounter.toString() : "", size: 18 })] })] }),
                        new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [
                                        new docx_1.TextRun({ text: isFirstUnit ? s.regNo : "", size: 18 }),
                                        ...(isFirstUnit && s.qualifierSuffix ? [new docx_1.TextRun({ text: s.qualifierSuffix, size: 14, subScript: true })] : []),
                                    ] })] }),
                        new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: isFirstUnit ? s.name : "", size: 18 })] })] }),
                        new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: uCode, size: 18 })] })] }),
                        new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: uName, size: 18 })] })] }),
                    ]
                }));
            });
            studentCounter++;
        }
        else {
            rows.push(new docx_1.TableRow({
                children: [
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ text: studentCounter.toString() })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ text: s.regNo })] }),
                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ text: s.name })] }),
                    new docx_1.TableCell({ margins: cellMargin, columnSpan: 2, children: [new docx_1.Paragraph({ text: "Refer to individual transcript for unit details." })] }),
                ]
            }));
            studentCounter++;
        }
    });
    return new docx_1.Table({ width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
        borders: {
            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE },
            left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE },
            insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
        }, rows });
}
const generateAwardListDoc = async (data) => {
    const { programName, academicYear, logoBuffer, awardList, institutionId } = data;
    const { meta } = await getDocContext({ institutionId });
    if (awardList.length === 0) {
        const doc = new docx_1.Document({
            sections: [
                {
                    children: [
                        new docx_1.Paragraph({
                            alignment: docx_1.AlignmentType.CENTER,
                            children: [
                                new docx_1.TextRun({
                                    text: `No eligible graduates found for ${programName} — ${academicYear}.`,
                                    bold: true,
                                    size: 24,
                                }),
                            ],
                        }),
                    ],
                },
            ],
        });
        return docx_1.Packer.toBuffer(doc);
    }
    const cellMargin = { top: 60, bottom: 60, left: 100, right: 100 };
    const classOrder = [
        "FIRST CLASS HONOURS",
        "SECOND CLASS HONOURS (UPPER DIVISION)",
        "SECOND CLASS HONOURS (LOWER DIVISION)",
        "PASS",
    ];
    const byClass = new Map();
    classOrder.forEach((c) => byClass.set(c, []));
    awardList.forEach((s) => {
        const key = classOrder.includes(s.classification)
            ? s.classification
            : "PASS";
        byClass.get(key).push(s);
    });
    const sections = [
        ...createDocHeader(logoBuffer, programName, academicYear, "Final Year", "AWARD LIST", meta, "ORDINARY"),
        new docx_1.Paragraph({
            alignment: docx_1.AlignmentType.JUSTIFIED,
            spacing: { before: 300, after: 400 },
            children: [
                new docx_1.TextRun({
                    text: `The following ${awardList.length} candidate(s) have satisfied the Board of Examiners in all prescribed units. The Board of Examiners recommends that they be `,
                    size: 22,
                }),
                new docx_1.TextRun({
                    text: `AWARDED THE DEGREE OF ${programName.toUpperCase()}.`,
                    bold: true,
                    size: 22,
                }),
            ],
        }),
    ];
    let globalCounter = 1;
    for (const cls of classOrder) {
        const group = byClass.get(cls) || [];
        if (group.length === 0)
            continue;
        sections.push(new docx_1.Paragraph({
            spacing: { before: 400, after: 150 },
            children: [
                new docx_1.TextRun({ text: cls, bold: true, size: 24, underline: {} }),
            ],
        }));
        const headerRow = new docx_1.TableRow({
            children: ["S/No.", "Reg. No.", "Name", "WAA (%)"].map((h) => new docx_1.TableCell({
                margins: cellMargin,
                children: [
                    new docx_1.Paragraph({
                        children: [new docx_1.TextRun({ text: h, bold: true, size: 20 })],
                    }),
                ],
            })),
        });
        const dataRows = group.map((s) => new docx_1.TableRow({
            children: [
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: String(globalCounter++), size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.regNo, size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: formatStudentName(s.name), size: 20 })] })] }),
                new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.waa.toFixed(2), size: 20 })] })] }),
            ],
        }));
        sections.push(new docx_1.Table({
            width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
            borders: {
                top: { style: docx_1.BorderStyle.NONE },
                bottom: { style: docx_1.BorderStyle.NONE },
                left: { style: docx_1.BorderStyle.NONE },
                right: { style: docx_1.BorderStyle.NONE },
                insideHorizontal: { style: docx_1.BorderStyle.NONE },
                insideVertical: { style: docx_1.BorderStyle.NONE },
            },
            rows: [headerRow, ...dataRows],
        }));
    }
    sections.push(new docx_1.Paragraph({
        spacing: { before: 500 },
        children: [
            new docx_1.TextRun({
                text: `TOTAL: ${awardList.length} CANDIDATES`,
                bold: true,
                size: 22,
            }),
        ],
    }), ...createDocFooter(meta));
    const doc = new docx_1.Document({ sections: [{ children: sections }] });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateAwardListDoc = generateAwardListDoc;
const generateDefermentDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status === "DEFERMENT");
    const count = list.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const formattedList = list.map((s) => ({
        regNo: s.regNo || "N/A", name: s.name || "N/A",
        effectiveDate: s.academicLeavePeriod?.startDate ? new Date(s.academicLeavePeriod.startDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "N/A",
        endDate: s.academicLeavePeriod?.endDate ? new Date(s.academicLeavePeriod.endDate).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "N/A",
        remarks: s.remarks?.includes(":") ? s.remarks.split(":")[1].trim() : s.remarks || "Approved",
    }));
    const doc = new docx_1.Document({
        sections: [
            {
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "DEFERMENT OF ADMISSION", meta, data.examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: `The following `, size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) have been granted deferment of admission in accordance with `, size: 22 }),
                            new docx_1.TextRun({ text: `ENG Rule 20 `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `and are expected to register at the commencement of the academic year following the end of their deferment period.`, size: 22 }),
                        ],
                    }),
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE },
                            right: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: ["S/No", "Reg No.", "Name", "From", "To", "Reason"].map((h) => new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: h, bold: true, size: 18 })] })] })),
                            }),
                            ...formattedList.map((s, i) => new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: (i + 1).toString(), size: 18 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.regNo, size: 18 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.name, size: 18 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.effectiveDate, size: 18 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.endDate, size: 18 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.remarks, size: 18 })] })] }),
                                ],
                            })),
                        ],
                    }),
                    ...createDocFooter(meta),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateDefermentDoc = generateDefermentDoc;
const generateSupplementaryExamsDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const suppCandidates = blocked.filter((s) => s.status.includes("SUPP") &&
        !s.reasons?.some((r) => /stayout|a\/so|repeat\s+year|a\/ra|readmission|readmit|carry\s*forward|a\/cf/i.test(r)));
    const count = suppCandidates.length;
    const currentYearOrdinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, currentYearOrdinal, "SUPPLEMENTARY", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `candidate(s) failed to satisfy the ${meta.schoolName} Board of Examiners in the unit(s) indicated against their names during the `, size: 22 }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "Academic Year, ", size: 22 }),
                            new docx_1.TextRun({ text: `${currentYearOrdinal} Year `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "Examinations for the ", size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `The ${meta.schoolName} Board of Examiners recommends that they sit for the supplementary exams when next offered. `, size: 22 }),
                        ],
                    }),
                    createStandardUnitDetailTable(suppCandidates, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateSupplementaryExamsDoc = generateSupplementaryExamsDoc;
const generateCarryForwardSuppDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status.includes("SUPP") &&
        s.reasons?.some((r) => /carry\s*forward|a\/cf/i.test(r)));
    const count = list.length;
    const ordinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, ordinal, "SUPPLEMENTARY (After Carry Forward)", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `candidate(s) failed their carry-forward supplementary examinations (A/CFS) in the unit(s) indicated against their names during the `,
                                size: 22,
                            }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, ${ordinal} Year Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners recommends that they sit for supplementary examinations when next offered.`,
                                size: 22,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateCarryForwardSuppDoc = generateCarryForwardSuppDoc;
const generateStayoutSuppDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status.includes("SUPP") &&
        s.reasons?.some((r) => /stayout|a\/so/i.test(r)));
    const count = list.length;
    const ordinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, ordinal, "SUPPLEMENTARY (After Stayout)", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `candidate(s) were required to stay out and retake examinations. ` +
                                    `Having now sat for the failed units, they have not fully satisfied the ${meta.schoolName} ` +
                                    `Board of Examiners in the unit(s) indicated against their names during the `,
                                size: 22,
                            }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, ${ordinal} Year Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners recommends that they sit for supplementary examinations (A/SOS) when next offered.`,
                                size: 22,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateStayoutSuppDoc = generateStayoutSuppDoc;
const generateRepeatSuppDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status.includes("SUPP") &&
        s.reasons?.some((r) => /repeat\s+year|a\/ra/i.test(r)));
    const count = list.length;
    const ordinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, ordinal, "SUPPLEMENTARY (After Repeat Year)", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `candidate(s) were required to repeat the academic year. ` +
                                    `After repeating, they have not fully satisfied the ${meta.schoolName} Board of Examiners ` +
                                    `in the unit(s) indicated against their names during the `,
                                size: 22,
                            }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, ${ordinal} Year Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners recommends that they sit for supplementary examinations when next offered.`,
                                size: 22,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateRepeatSuppDoc = generateRepeatSuppDoc;
const generateReadmissionSuppDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status.includes("SUPP") &&
        s.reasons?.some((r) => /readmission|readmit/i.test(r)));
    const count = list.length;
    const ordinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, ordinal, "SUPPLEMENTARY (After Readmission)", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `candidate(s) have been formally readmitted to the programme and have outstanding ` +
                                    `supplementary examinations in the unit(s) indicated against their names during the `,
                                size: 22,
                            }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, ${ordinal} Year Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners recommends that they sit for supplementary examinations when next offered.`,
                                size: 22,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateReadmissionSuppDoc = generateReadmissionSuppDoc;
const generateAcademicLeaveSuppDoc = async (data) => {
    const { programName, academicYear, yearOfStudy, blocked, institutionId, examType } = data;
    const { meta, logoBuffer } = await getDocContext(data);
    const list = blocked.filter((s) => s.status.includes("SUPP") &&
        s.reasons?.some((r) => /academic\s*leave|on\s*leave|a\/sp/i.test(r)));
    const count = list.length;
    const ordinal = getOrdinalYear(yearOfStudy);
    const cellMargin = { top: 100, bottom: 100, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, ordinal, "SUPPLEMENTARY (After Academic Leave)", meta, examType || "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(count)} (${count}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `candidate(s) have resumed studies after academic leave and have outstanding supplementary ` +
                                    `examinations in the unit(s) indicated against their names during the `,
                                size: 22,
                            }),
                            new docx_1.TextRun({ text: `${academicYear} `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: `Academic Year, ${ordinal} Year Examinations for the `, size: 22 }),
                            new docx_1.TextRun({ text: `${programName}. `, bold: true, size: 22 }),
                            new docx_1.TextRun({
                                text: `The ${meta.schoolName} Board of Examiners recommends that they sit for supplementary examinations when next offered.`,
                                size: 22,
                            }),
                        ],
                    }),
                    createStandardUnitDetailTable(list, cellMargin, "FAIL"),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateAcademicLeaveSuppDoc = generateAcademicLeaveSuppDoc;
const generateSimpleAwardListDoc = async (data) => {
    const { programName, academicYear, awardList, institutionId } = data;
    const { meta, logoBuffer } = await getDocContext({ institutionId });
    const cellMargin = { top: 50, bottom: 50, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [{
                children: [
                    ...createDocHeader(logoBuffer, programName, academicYear, "Final Year", "AWARD LIST", meta, "ORDINARY"),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.JUSTIFIED,
                        spacing: { before: 400, after: 300 },
                        children: [
                            new docx_1.TextRun({ text: "The following ", size: 22 }),
                            new docx_1.TextRun({ text: `${numberToWords(awardList.length)} (${awardList.length}) `, bold: true, size: 22 }),
                            new docx_1.TextRun({ text: "candidate(s) have satisfied the Board of Examiners in all prescribed units. The Board recommends they be ", size: 22 }),
                            new docx_1.TextRun({ text: `AWARDED THE DEGREE OF ${programName.toUpperCase()}.`, bold: true, size: 22 }),
                        ],
                    }),
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE },
                            left: { style: docx_1.BorderStyle.NONE }, right: { style: docx_1.BorderStyle.NONE },
                            insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({ width: { size: 8, type: docx_1.WidthType.PERCENTAGE }, margins: cellMargin,
                                        children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: "S/No.", bold: true, size: 20 })] })] }),
                                    new docx_1.TableCell({ width: { size: 27, type: docx_1.WidthType.PERCENTAGE }, margins: cellMargin,
                                        children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: "Reg. No.", bold: true, size: 20 })] })] }),
                                    new docx_1.TableCell({ width: { size: 65, type: docx_1.WidthType.PERCENTAGE }, margins: cellMargin,
                                        children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: "Name", bold: true, size: 20 })] })] }),
                                ],
                            }),
                            ...awardList.map((s, i) => new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: String(i + 1), size: 20 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: s.regNo, size: 20 })] })] }),
                                    new docx_1.TableCell({ margins: cellMargin, children: [new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: formatStudentName(s.name), size: 20 })] })] }),
                                ],
                            })),
                        ],
                    }),
                    ...createDocFooter(meta),
                ],
            }],
    });
    return docx_1.Packer.toBuffer(doc);
};
exports.generateSimpleAwardListDoc = generateSimpleAwardListDoc;
const generateIneligibilityNotice = async (student, data) => {
    const { programName, academicYear, yearOfStudy, logoBuffer } = data;
    const capitalizedStudentName = formatStudentName(student.name).toUpperCase();
    const doc = new docx_1.Document({
        sections: [
            {
                properties: {},
                children: [
                    // 1. LOGO
                    ...(logoBuffer.length > 0
                        ? [
                            new docx_1.Paragraph({
                                alignment: docx_1.AlignmentType.CENTER,
                                children: [
                                    new docx_1.ImageRun({
                                        data: logoBuffer,
                                        transformation: { width: 150, height: 80 },
                                        type: "png",
                                    }),
                                ],
                            }),
                        ]
                        : []),
                    // 2. HEADERS (Consistent with Pass List)
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { before: 100 },
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.instName.toUpperCase(),
                                bold: true,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.schoolName.toUpperCase(),
                                bold: true,
                                size: 20,
                            }),
                        ],
                    }),
                    // 3. INTERNAL MEMO STYLE ADDRESSING
                    new docx_1.Paragraph({
                        spacing: { before: 400 },
                        children: [
                            new docx_1.TextRun({ text: "TO: ", bold: true }),
                            new docx_1.TextRun({ text: capitalizedStudentName }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "REG NO: ", bold: true }),
                            new docx_1.TextRun({ text: student.regNo }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "PROGRAM: ", bold: true }),
                            new docx_1.TextRun({ text: programName.toUpperCase() }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "ACADEMIC YEAR: ", bold: true }),
                            new docx_1.TextRun({ text: academicYear }),
                        ],
                    }),
                    // 4. SUBJECT LINE
                    new docx_1.Paragraph({
                        spacing: { before: 400, after: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: `RE: INELIGIBILITY FOR PROMOTION TO YEAR ${yearOfStudy + 1}`,
                                bold: true,
                                size: 22,
                                underline: {},
                            }),
                        ],
                    }),
                    // 5. BODY TEXT
                    new docx_1.Paragraph({
                        spacing: { after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: "Following the School Board of Examiners meeting, it was noted that you did not satisfy the examiners in the following units during the current academic year:",
                                size: 20,
                            }),
                        ],
                    }),
                    // 6. UNIT LIST (Dynamically rendered from student.reasons)
                    ...student.reasons.map((unitString) => new docx_1.Paragraph({
                        bullet: { level: 0 },
                        spacing: { before: 150 },
                        children: [
                            new docx_1.TextRun({
                                text: unitString,
                                bold: true,
                                size: 19,
                            }),
                        ],
                    })),
                    new docx_1.Paragraph({
                        spacing: { before: 300, after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: "Consequently, you are not eligible for promotion. You are advised to prepare for Supplementary Examinations or register for Retakes as per the University Examination Regulations.",
                                size: 20,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { after: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: `Please contact the Office of the Dean, ${config_1.default.schoolName}, for the schedule of supplementary examinations.`,
                                size: 20,
                            }),
                        ],
                    }),
                    // 7. SIGNATORY
                    new docx_1.Paragraph({
                        spacing: { before: 600, after: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: `APPROVED BY THE BOARD OF EXAMINERS, ${config_1.default.schoolName.toUpperCase()}`,
                                bold: true,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { before: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: "SIGNED: __________________________\t\tDATE: _______________",
                                bold: true,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: `\tDEAN, ${config_1.default.schoolName.toUpperCase()}`,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { before: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: "Cc: Registrar (Academic Affairs)\n    ",
                                size: 16,
                            }),
                        ],
                    }),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateIneligibilityNotice = generateIneligibilityNotice;
// New function for Special Exam Notice
const generateSpecialExamNotice = async (student, data) => {
    const { programName, academicYear, logoBuffer } = data;
    const capitalizedStudentName = formatStudentName(student.name).toUpperCase();
    const specialUnits = student.reasons
        .filter((r) => r.toUpperCase().includes("SPECIAL"))
        .map((r) => r.replace("- SPECIAL", "").replace("SPECIAL", "").trim());
    const doc = new docx_1.Document({
        sections: [
            {
                properties: {},
                children: [
                    // 1. LOGO
                    ...(logoBuffer.length > 0
                        ? [
                            new docx_1.Paragraph({
                                alignment: docx_1.AlignmentType.CENTER,
                                children: [
                                    new docx_1.ImageRun({
                                        data: logoBuffer,
                                        transformation: { width: 150, height: 80 },
                                        type: "png",
                                    }),
                                ],
                            }),
                        ]
                        : []),
                    // 2. HEADERS
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.instName.toUpperCase(),
                                bold: true,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.schoolName.toUpperCase(),
                                bold: true,
                                size: 20,
                            }),
                        ],
                    }),
                    // 3. ADDRESSING
                    new docx_1.Paragraph({
                        spacing: { before: 400 },
                        children: [
                            new docx_1.TextRun({ text: "TO: ", bold: true }),
                            new docx_1.TextRun({ text: capitalizedStudentName }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "REG NO: ", bold: true }),
                            new docx_1.TextRun({ text: student.regNo }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "PROGRAM: ", bold: true }),
                            new docx_1.TextRun({ text: programName.toUpperCase() }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({ text: "ACADEMIC YEAR: ", bold: true }),
                            new docx_1.TextRun({ text: academicYear }),
                        ],
                    }),
                    // 4. SUBJECT
                    new docx_1.Paragraph({
                        spacing: { before: 400, after: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: "RE: APPROVAL TO SIT FOR SPECIAL EXAMINATIONS",
                                bold: true,
                                size: 22,
                                underline: { type: docx_1.BorderStyle.SINGLE },
                            }),
                        ],
                    }),
                    // 5. BODY
                    new docx_1.Paragraph({
                        spacing: { after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: `This is to inform you that the College Board of Examiners has approved your request to sit for Special Examinations in the following unit(s) during the ${academicYear} academic cycle:`,
                                size: 20,
                            }),
                        ],
                    }),
                    // DYNAMIC UNIT LIST
                    ...specialUnits.map((unitName) => new docx_1.Paragraph({
                        bullet: { level: 0 },
                        spacing: { before: 100 },
                        children: [
                            new docx_1.TextRun({ text: unitName, bold: true, size: 20 }),
                        ],
                    })),
                    new docx_1.Paragraph({
                        spacing: { before: 300, after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: "Please note that a Special Examination is treated as a first attempt. Failure to sit for these exams will result in the units being graded as 'Incomplete'.",
                                size: 20,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { after: 400 },
                        children: [
                            new docx_1.TextRun({
                                text: "Check the departmental notice board for the scheduled dates and venues.",
                                size: 20,
                            }),
                        ],
                    }),
                    // 7. SIGNATORY
                    new docx_1.Paragraph({
                        spacing: { before: 800 },
                        children: [
                            new docx_1.TextRun({
                                text: "SIGNED: __________________________\t\tDATE: _______________",
                                bold: true,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        children: [
                            new docx_1.TextRun({
                                text: `\tDEAN, ${config_1.default.schoolName.toUpperCase()}`,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { before: 400 },
                        children: [new docx_1.TextRun({ text: "Cc: Exam Coordinator", size: 16 })],
                    }),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateSpecialExamNotice = generateSpecialExamNotice;
const generateStudentTranscript = async (student, results, data) => {
    const { programName, academicYear, logoBuffer, status } = data;
    // Use the value passed from the controller, falling back to student record if needed
    const displayYear = data.yearToPromote ||
        data.yearOfStudy ||
        student.currentYearOfStudy ||
        "N/A";
    const cellMargin = { top: 80, bottom: 80, left: 100, right: 100 };
    const doc = new docx_1.Document({
        sections: [
            {
                properties: {},
                children: [
                    // 1. LOGO (Same 80x80 size as Summary)
                    ...(logoBuffer.length > 0
                        ? [
                            new docx_1.Paragraph({
                                alignment: docx_1.AlignmentType.CENTER,
                                children: [
                                    new docx_1.ImageRun({
                                        data: logoBuffer,
                                        transformation: { width: 150, height: 80 },
                                        type: "png",
                                    }),
                                ],
                            }),
                        ]
                        : []),
                    // 2. INSTITUTIONAL HEADERS
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { before: 50, after: 100 },
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.instName.toUpperCase(),
                                bold: true,
                                size: 24,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 50 },
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.postalAddress,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 50 },
                        children: [
                            new docx_1.TextRun({
                                text: "Cell Phone",
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 50 },
                        children: [
                            new docx_1.TextRun({
                                text: config_1.default.cellPhone,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 100 },
                        children: [
                            new docx_1.TextRun({
                                text: `Email: ${config_1.default.schoolEmail}`,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: "UNDERGRADUATE ACADEMIC TRANSCRIPT",
                                bold: true,
                                size: 20,
                                underline: {},
                            }),
                        ],
                    }),
                    // 3. STUDENT PROFILE BOX
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            // top: { style: BorderStyle.NONE },
                            top: {
                                style: docx_1.BorderStyle.SINGLE,
                                size: 16, // 2pt thickness
                                color: "000000", // Optional: black
                            },
                            bottom: { style: docx_1.BorderStyle.NONE },
                            left: { style: docx_1.BorderStyle.NONE },
                            right: { style: docx_1.BorderStyle.NONE },
                            insideHorizontal: { style: docx_1.BorderStyle.NONE },
                            insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({
                                        children: [
                                            new docx_1.Paragraph({
                                                children: [
                                                    new docx_1.TextRun({
                                                        text: "NAME: ",
                                                        bold: true,
                                                        size: 20,
                                                    }),
                                                    new docx_1.TextRun({
                                                        text: formatStudentName(student.name).toUpperCase(),
                                                        bold: false,
                                                        size: 20,
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new docx_1.TableCell({
                                        children: [
                                            new docx_1.Paragraph({
                                                alignment: docx_1.AlignmentType.RIGHT,
                                                children: [
                                                    new docx_1.TextRun({
                                                        text: "REG NO: ",
                                                        bold: true,
                                                        size: 20,
                                                    }),
                                                    new docx_1.TextRun({
                                                        text: student.regNo,
                                                        bold: false,
                                                        size: 20,
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({
                        spacing: { before: 100 },
                        children: [
                            new docx_1.TextRun({
                                text: "SCHOOL: ",
                                bold: true,
                                size: 18,
                            }),
                            new docx_1.TextRun({
                                text: config_1.default.schoolName.toUpperCase(),
                                bold: false,
                                size: 18,
                            }),
                        ],
                    }),
                    // 5. Programme name
                    new docx_1.Paragraph({
                        spacing: { before: 100 },
                        children: [
                            new docx_1.TextRun({
                                text: "PROGRAM: ",
                                bold: true,
                                size: 18,
                            }),
                            new docx_1.TextRun({
                                text: `${programName.toUpperCase()}`,
                                size: 18,
                            }),
                        ],
                    }),
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: docx_1.BorderStyle.NONE },
                            bottom: { style: docx_1.BorderStyle.NONE },
                            left: { style: docx_1.BorderStyle.NONE },
                            right: { style: docx_1.BorderStyle.NONE },
                            insideHorizontal: { style: docx_1.BorderStyle.NONE },
                            insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({
                                        width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                        children: [
                                            new docx_1.Paragraph({
                                                spacing: { before: 100, after: 300 },
                                                children: [
                                                    new docx_1.TextRun({
                                                        text: "ACADEMIC YEAR: ",
                                                        bold: true,
                                                        size: 20,
                                                    }),
                                                    new docx_1.TextRun({
                                                        text: academicYear || "N/A",
                                                        bold: false,
                                                        size: 20,
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new docx_1.TableCell({
                                        width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        children: [
                                            new docx_1.Paragraph({
                                                alignment: docx_1.AlignmentType.RIGHT,
                                                children: [
                                                    new docx_1.TextRun({
                                                        text: "YEAR OF STUDY: ",
                                                        bold: true,
                                                        size: 20,
                                                    }),
                                                    new docx_1.TextRun({
                                                        text: `${displayYear}`,
                                                        bold: false,
                                                        size: 20,
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    //  STATUS SUMMARY
                    new docx_1.Paragraph({
                        alignment: docx_1.AlignmentType.CENTER,
                        spacing: { after: 200 },
                        children: [
                            new docx_1.TextRun({
                                text: "RESULT:  PASS",
                                bold: true,
                                size: 20,
                                underline: {},
                            }),
                        ],
                    }),
                    // 4. RESULTS TABLE (Units & Grades)
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        layout: docx_1.TableLayoutType.FIXED,
                        borders: {
                            top: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                            left: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                            right: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                            insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                            insideHorizontal: { style: docx_1.BorderStyle.NIL },
                        },
                        rows: [
                            // Header Row
                            new docx_1.TableRow({
                                tableHeader: true,
                                children: [
                                    new docx_1.TableCell({
                                        width: { size: 15, type: docx_1.WidthType.PERCENTAGE }, // Narrower Code
                                        margins: cellMargin,
                                        verticalAlign: docx_1.VerticalAlign.CENTER,
                                        borders: {
                                            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                            top: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                        },
                                        children: [
                                            new docx_1.Paragraph({
                                                children: [
                                                    new docx_1.TextRun({ text: "CODE", bold: true, size: 18 }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new docx_1.TableCell({
                                        width: { size: 70, type: docx_1.WidthType.PERCENTAGE }, // Larger Title
                                        margins: cellMargin,
                                        verticalAlign: docx_1.VerticalAlign.CENTER,
                                        borders: {
                                            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                            top: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                        },
                                        children: [
                                            new docx_1.Paragraph({
                                                children: [
                                                    new docx_1.TextRun({
                                                        text: "COURSE UNIT TITLE",
                                                        bold: true,
                                                        size: 18,
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new docx_1.TableCell({
                                        width: { size: 15, type: docx_1.WidthType.PERCENTAGE }, // Narrower Grade
                                        margins: cellMargin,
                                        verticalAlign: docx_1.VerticalAlign.CENTER,
                                        borders: {
                                            bottom: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                            top: { style: docx_1.BorderStyle.SINGLE, size: 2 },
                                        },
                                        children: [
                                            new docx_1.Paragraph({
                                                alignment: docx_1.AlignmentType.CENTER,
                                                children: [
                                                    new docx_1.TextRun({ text: "GRADE", bold: true, size: 18 }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            // Data Rows
                            ...results.map((r, index) => {
                                const unitCode = String(r?.code ?? "N/A").toUpperCase();
                                const unitName = String(r?.name ?? "COURSE TITLE MISSING").toUpperCase();
                                const unitGrade = String(r?.grade ?? "-");
                                const isLastRow = index === results.length - 1;
                                const bottomBorderStyle = isLastRow
                                    ? docx_1.BorderStyle.SINGLE
                                    : docx_1.BorderStyle.NIL;
                                const bottomBorderSize = isLastRow ? 2 : 0;
                                return new docx_1.TableRow({
                                    children: [
                                        new docx_1.TableCell({
                                            width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                            margins: cellMargin,
                                            verticalAlign: docx_1.VerticalAlign.CENTER,
                                            borders: {
                                                bottom: {
                                                    style: bottomBorderStyle,
                                                    size: bottomBorderSize,
                                                },
                                            },
                                            children: [
                                                new docx_1.Paragraph({
                                                    spacing: { after: 10 },
                                                    children: [new docx_1.TextRun({ text: unitCode, size: 18 })],
                                                }),
                                            ],
                                        }),
                                        new docx_1.TableCell({
                                            width: { size: 70, type: docx_1.WidthType.PERCENTAGE },
                                            margins: cellMargin,
                                            verticalAlign: docx_1.VerticalAlign.CENTER,
                                            borders: {
                                                bottom: {
                                                    style: bottomBorderStyle,
                                                    size: bottomBorderSize,
                                                },
                                            },
                                            children: [
                                                new docx_1.Paragraph({
                                                    spacing: { after: 10 },
                                                    children: [new docx_1.TextRun({ text: unitName, size: 18 })],
                                                }),
                                            ],
                                        }),
                                        new docx_1.TableCell({
                                            width: { size: 15, type: docx_1.WidthType.PERCENTAGE },
                                            margins: cellMargin,
                                            verticalAlign: docx_1.VerticalAlign.CENTER,
                                            borders: {
                                                bottom: {
                                                    style: bottomBorderStyle,
                                                    size: bottomBorderSize,
                                                },
                                            },
                                            children: [
                                                new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { after: 10 }, children: [new docx_1.TextRun({ text: unitGrade, bold: true, size: 18 })] }),
                                            ],
                                        }),
                                    ],
                                });
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 200 }, children: [] }),
                    // 6. GRADING KEY (Formal Table Structure)
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE },
                            right: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: [
                                    // LEFT CELL: The Grading Key Table
                                    new docx_1.TableCell({
                                        width: { size: 40, type: docx_1.WidthType.PERCENTAGE },
                                        children: [
                                            new docx_1.Table({
                                                width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                                                borders: {
                                                    top: { style: docx_1.BorderStyle.SINGLE, size: 1 }, bottom: { style: docx_1.BorderStyle.SINGLE, size: 1 }, left: { style: docx_1.BorderStyle.SINGLE, size: 1 },
                                                    right: { style: docx_1.BorderStyle.SINGLE, size: 1 }, insideVertical: { style: docx_1.BorderStyle.SINGLE, size: 1 }, insideHorizontal: { style: docx_1.BorderStyle.NONE },
                                                },
                                                rows: [
                                                    // Header Row
                                                    new docx_1.TableRow({
                                                        tableHeader: true,
                                                        children: [["GRADE", 20], ["RANGE", 30], ["DESCRIPTION", 50]].map(([text, width]) => new docx_1.TableCell({
                                                            width: { size: width, type: docx_1.WidthType.PERCENTAGE },
                                                            shading: { fill: "F2F2F2" },
                                                            children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: text, bold: true, size: 16 })] })],
                                                        })),
                                                    }),
                                                    // Data Rows
                                                    ...[{ g: "A", r: "70 - 100%", d: "EXCELLENT" }, { g: "B", r: "60 - 69%", d: "GOOD" }, { g: "C", r: "50 - 59%", d: "SATISFACTORY" }, { g: "D", r: "40 - 49%", d: "PASS" }, { g: "E", r: "0 - 39%", d: "FAIL" }].map((item) => new docx_1.TableRow({
                                                        children: [
                                                            new docx_1.TableCell({ children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: item.g, size: 16 })] })] }),
                                                            new docx_1.TableCell({ children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: item.r, size: 16 })] })] }),
                                                            new docx_1.TableCell({ children: [new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, children: [new docx_1.TextRun({ text: item.d, size: 16 })] })] }),
                                                        ],
                                                    })),
                                                ],
                                            }),
                                        ],
                                    }),
                                    // RIGHT CELL: Registration Number
                                    new docx_1.TableCell({
                                        width: { size: 30, type: docx_1.WidthType.PERCENTAGE },
                                        verticalAlign: docx_1.VerticalAlign.BOTTOM,
                                        children: [
                                            new docx_1.Paragraph({ alignment: docx_1.AlignmentType.RIGHT, children: [new docx_1.TextRun({ text: "NB: ", bold: true, size: 20 })] }),
                                            new docx_1.Paragraph({ alignment: docx_1.AlignmentType.RIGHT, children: [new docx_1.TextRun({ text: "1 unit consists of 35 lecture hours or equivalent (3 Practical hours of two tutorial hours are equivalent to 0ne lecture hour ) ", bold: false, size: 16 })] }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    // 7. FOOTER & SIGNATORIES
                    new docx_1.Table({
                        width: { size: 100, type: docx_1.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: docx_1.BorderStyle.NONE }, bottom: { style: docx_1.BorderStyle.NONE }, left: { style: docx_1.BorderStyle.NONE },
                            right: { style: docx_1.BorderStyle.NONE }, insideHorizontal: { style: docx_1.BorderStyle.NONE }, insideVertical: { style: docx_1.BorderStyle.NONE },
                        },
                        rows: [
                            new docx_1.TableRow({
                                children: [
                                    new docx_1.TableCell({
                                        children: [
                                            new docx_1.Paragraph({ spacing: { before: 800 }, children: [new docx_1.TextRun({ text: "SIGNED: __________________________________________", bold: true })] }),
                                            new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: `DEAN, ${config_1.default.schoolName.toUpperCase()}`, bold: true, size: 18 })] }),
                                        ],
                                    }),
                                    new docx_1.TableCell({
                                        children: [
                                            new docx_1.Paragraph({ spacing: { before: 800 }, children: [new docx_1.TextRun({ text: "SIGNED: __________________________________________", bold: true })] }),
                                            new docx_1.Paragraph({ children: [new docx_1.TextRun({ text: `REGISTRAR, ${config_1.default.registrar.toUpperCase()}`, bold: true, size: 18 })] }),
                                        ]
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new docx_1.Paragraph({ spacing: { before: 100 }, children: [new docx_1.TextRun({ text: `DATE OF ISSUE: ${new Date().toLocaleDateString()}`, size: 14, italics: true })] }),
                    new docx_1.Paragraph({ alignment: docx_1.AlignmentType.CENTER, spacing: { before: 400 }, children: [new docx_1.TextRun({ text: "--- This result slip is issued without any erasures or alterations ---", italics: true, size: 14 })] }),
                ],
            },
        ],
    });
    return await docx_1.Packer.toBuffer(doc);
};
exports.generateStudentTranscript = generateStudentTranscript;
