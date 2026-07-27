"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateRegNo = validateRegNo;
// serverside/src/utils/validateRegNo.ts
const InstitutionSettings_1 = __importDefault(require("../models/InstitutionSettings"));
const mongoose_1 = __importDefault(require("mongoose"));
/**
 * Validates a student's reg number against the department's configured patterns.
 *
 * Returns { valid: true } in these situations — no error:
 *   1. enforceRegNoPattern is false on InstitutionSettings
 *   2. The department has NO patterns configured
 *   3. The reg no matches ANY configured pattern
 *
 * Returns { valid: false, reason } only when:
 *   - enforceRegNoPattern is true AND
 *   - The department has ≥ 1 pattern configured AND
 *   - The reg no matches NONE of them
 */
async function validateRegNo(regNo, institutionId, programId) {
    const VALID = { valid: true };
    if (!regNo?.trim())
        return VALID;
    // 1. Load settings — if missing, skip validation (fail open, not closed)
    const settings = await InstitutionSettings_1.default.findOne({
        institution: new mongoose_1.default.Types.ObjectId(institutionId),
    })
        .select("ruleSet.passMark enforceRegNoPattern schools")
        .lean();
    if (!settings)
        return VALID; // no settings → allow anything
    // 2. Check enforcement flag — if off, skip entirely
    if (!settings.enforceRegNoPattern)
        return VALID;
    // 3. Find which department this program belongs to
    //    We need the program's departmentCode to look up patterns
    //    Load program to get departmentCode
    let departmentCode = null;
    let schoolCode = null;
    try {
        const Program = (await Promise.resolve().then(() => __importStar(require("../models/Program")))).default;
        const prog = await Program.findById(programId)
            .select("departmentCode schoolCode")
            .lean();
        departmentCode = prog?.departmentCode?.toUpperCase() ?? null;
        schoolCode = prog?.schoolCode?.toUpperCase() ?? null;
    }
    catch {
        return VALID; // if program lookup fails, don't block registration
    }
    if (!departmentCode || !schoolCode)
        return VALID;
    // 4. Find the department's patterns
    const school = settings.schools?.find(s => s.code === schoolCode);
    const dept = school?.departments?.find(d => d.code === departmentCode);
    const patterns = dept?.regNoPatterns ?? [];
    // 5. No patterns configured → skip validation
    if (patterns.length === 0)
        return VALID;
    // 6. Try to match against each pattern
    const normalised = regNo.trim().toUpperCase();
    for (const pattern of patterns) {
        let regex;
        if (pattern.manualRegex?.trim()) {
            // Use the coordinator-supplied regex directly
            try {
                regex = new RegExp(pattern.manualRegex.trim(), "i");
            }
            catch {
                // If the regex is malformed, skip this pattern (don't block)
                continue;
            }
        }
        else {
            // Build from prefix + separator + yearDigits + sequence
            // Example: prefix="E", separator="-", yearDigits=3
            // Matches: E024-0001, E024-1234, E023-001, etc.
            const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            const sep = pattern.separator ? esc(pattern.separator) : "";
            const yr = `\\d{${pattern.yearDigits ?? 3}}`;
            // Sequence: 1 or more digits, optionally preceded by separator
            const seq = `\\d+`;
            // Full reg-no may also have additional segments (e.g. /2024 at end) — allow them
            // Pattern: PREFIX + optional_sep + YEAR + optional_sep + SEQUENCE + optional_suffix
            const pattern_str = `^${esc(pattern.prefix)}${sep}${yr}${sep ? sep : ""}${seq}`;
            regex = new RegExp(pattern_str, "i");
        }
        if (regex.test(normalised))
            return VALID;
    }
    // 7. None matched — invalid
    const examples = patterns
        .map(p => p.example)
        .filter(Boolean)
        .join(", ");
    return {
        valid: false,
        reason: examples
            ? `Reg no "${normalised}" does not match department pattern. Expected format: ${examples}`
            : `Reg no "${normalised}" does not match the configured pattern for department ${departmentCode}`,
    };
}
