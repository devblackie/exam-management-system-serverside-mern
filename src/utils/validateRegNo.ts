// serverside/src/utils/validateRegNo.ts — COMPLETE REWRITE
import InstitutionSettings from "../models/InstitutionSettings";
import mongoose from "mongoose";

export interface RegNoValidationResult {
  valid:   boolean;
  reason?: string;   // populated only when invalid
}

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
export async function validateRegNo(
  regNo:          string,
  institutionId:  string,
  programId:      string,
): Promise<RegNoValidationResult> {
  const VALID: RegNoValidationResult = { valid: true };

  if (!regNo?.trim()) return VALID;

  // 1. Load settings — if missing, skip validation (fail open, not closed)
  const settings = await InstitutionSettings.findOne({
    institution: new mongoose.Types.ObjectId(institutionId),
  })
    .select("ruleSet.passMark enforceRegNoPattern schools")
    .lean() as {
      enforceRegNoPattern?: boolean;
      schools?: Array<{
        code: string;
        departments?: Array<{
          code: string;
          regNoPatterns?: Array<{
            prefix:      string;
            separator:   string;
            yearDigits:  number;
            example:     string;
            manualRegex?: string;
          }>;
        }>;
      }>;
    } | null;

  if (!settings) return VALID;  // no settings → allow anything

  // 2. Check enforcement flag — if off, skip entirely
  if (!settings.enforceRegNoPattern) return VALID;

  // 3. Find which department this program belongs to
  //    We need the program's departmentCode to look up patterns
  //    Load program to get departmentCode
  let departmentCode: string | null = null;
  let schoolCode:     string | null = null;

  try {
    const Program = (await import("../models/Program")).default;
    const prog = await Program.findById(programId)
      .select("departmentCode schoolCode")
      .lean() as { departmentCode?: string; schoolCode?: string } | null;
    departmentCode = prog?.departmentCode?.toUpperCase() ?? null;
    schoolCode     = prog?.schoolCode?.toUpperCase()     ?? null;
  } catch {
    return VALID;  // if program lookup fails, don't block registration
  }

  if (!departmentCode || !schoolCode) return VALID;

  // 4. Find the department's patterns
  const school = settings.schools?.find(s => s.code === schoolCode);
  const dept   = school?.departments?.find(d => d.code === departmentCode);
  const patterns = dept?.regNoPatterns ?? [];

  // 5. No patterns configured → skip validation
  if (patterns.length === 0) return VALID;

  // 6. Try to match against each pattern
  const normalised = regNo.trim().toUpperCase();

  for (const pattern of patterns) {
    let regex: RegExp;

    if (pattern.manualRegex?.trim()) {
      // Use the coordinator-supplied regex directly
      try {
        regex = new RegExp(pattern.manualRegex.trim(), "i");
      } catch {
        // If the regex is malformed, skip this pattern (don't block)
        continue;
      }
    } else {
      // Build from prefix + separator + yearDigits + sequence
      // Example: prefix="E", separator="-", yearDigits=3
      // Matches: E024-0001, E024-1234, E023-001, etc.
      const esc = (s: string) =>
        s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

      const sep     = pattern.separator ? esc(pattern.separator) : "";
      const yr      = `\\d{${pattern.yearDigits ?? 3}}`;
      // Sequence: 1 or more digits, optionally preceded by separator
      const seq     = `\\d+`;
      // Full reg-no may also have additional segments (e.g. /2024 at end) — allow them
      // Pattern: PREFIX + optional_sep + YEAR + optional_sep + SEQUENCE + optional_suffix
      const pattern_str = `^${esc(pattern.prefix)}${sep}${yr}${sep ? sep : ""}${seq}`;
      regex = new RegExp(pattern_str, "i");
    }

    if (regex.test(normalised)) return VALID;
  }

  // 7. None matched — invalid
  const examples = patterns
    .map(p => p.example)
    .filter(Boolean)
    .join(", ");

  return {
    valid:  false,
    reason: examples
      ? `Reg no "${normalised}" does not match department pattern. Expected format: ${examples}`
      : `Reg no "${normalised}" does not match the configured pattern for department ${departmentCode}`,
  };
}
