"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.evaluateMark = evaluateMark;
exports.summarizeYearResults = summarizeYearResults;
function evaluateMark(total, hasMissing, passMark = 40) {
    if (hasMissing)
        return "MISSING";
    if (total === undefined || total < passMark)
        return "SUPPLEMENTARY";
    return "PASS";
}
function summarizeYearResults(results) {
    const supplementaries = results.filter((r) => r.status === "SUPPLEMENTARY");
    if (supplementaries.length >= 5) {
        return "RETAKE";
    }
    if (results.every((r) => r.status === "PASS")) {
        return "PASS";
    }
    return "SUPPLEMENTARY";
}
