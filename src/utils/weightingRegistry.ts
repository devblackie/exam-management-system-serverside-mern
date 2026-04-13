// // serverside/src/utils/weightingRegistry.ts

// export const WEIGHTING_REGISTRY: any = {
//   // Table 1: 5-Year (B.Sc Engineering)
//   ENG_5: {
//     Direct: { 1: 0.15, 2: 0.15, 3: 0.2, 4: 0.25, 5: 0.25 },
//     "Mid-Entry-Y2": { 2: 0.175, 3: 0.225, 4: 0.3, 5: 0.3 },
//     "Mid-Entry-Y3": { 3: 0.3, 4: 0.35, 5: 0.35 },
//   },
//   // Table 2: 4-Year (B.Ed Tech or IT)
//   ENG_4: {
//     Direct: { 1: 0.15, 2: 0.25, 3: 0.3, 4: 0.3 },
//     "Mid-Entry-Y2": { 2: 0.3, 3: 0.35, 4: 0.35 },
//     "Mid-Entry-Y3": { 3: 0.5, 4: 0.5 },
//   },
//   // Default for other schools (General 4-year)
//   GEN_4: {
//     Direct: { 1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25 },
//   },
// };


// export const getYearWeight = (prog: any, entry: string, year: number) => {
//   // 1. Determine the School Type Key
//   // Check if schoolType exists, otherwise infer from the name or duration
//   let schoolType = prog.schoolType;

//   if (!schoolType) {
//     if (prog.name?.toUpperCase().includes("ENGINEERING")) {
//       schoolType = "ENGINEERING";
//     } else {
//       schoolType = "GENERAL";
//     }
//   }

//   const key = schoolType === "ENGINEERING" ? `ENG_${prog.durationYears}` : `GEN_${prog.durationYears}`;

//   // 2. Retrieve the Scheme
//   const scheme = WEIGHTING_REGISTRY[key] || WEIGHTING_REGISTRY["GEN_4"];

//   // 3. Get weights for the entry type (Direct, Mid-Entry, etc.)
//   const entryWeights = scheme[entry] || scheme["Direct"];

//   // 4. Return the specific year weight
//   return entryWeights[year] || 0;
// };











// // serverside/src/utils/weightingRegistry.ts
// // COMPLETE — with robust school type detection
// //
// // FIX: Year 5 showing W:0% for Engineering students
// //   Root cause: getYearWeight fell through to GEN_4 when:
// //     (a) prog.schoolType was undefined/null, AND
// //     (b) prog.name did not contain "ENGINEERING" (e.g. "Bachelor of Science
// //         in Civil Engineering" — the check was case-sensitive and program
// //         names vary).
// //   GEN_4 has no Year 5 entry → returns 0 → display shows W:0%.
// //
// //   Fix: Use durationYears as the primary signal for ENG vs GEN lookup.
// //   5-year programs are engineering (ENG_5), 4-year can be either.
// //   Name matching is now more inclusive. Added ENG_6 as a safety net.

// export const WEIGHTING_REGISTRY: any = {
//   // Table 1: 5-Year (B.Sc Engineering)
//   ENG_5: {
//     Direct:          { 1: 0.15, 2: 0.15, 3: 0.20, 4: 0.25, 5: 0.25 },
//     "Mid-Entry-Y2":  { 2: 0.175, 3: 0.225, 4: 0.30, 5: 0.30 },
//     "Mid-Entry-Y3":  { 3: 0.30,  4: 0.35,  5: 0.35 },
//   },
//   // Table 2: 4-Year (B.Ed Tech or IT)
//   ENG_4: {
//     Direct:          { 1: 0.15, 2: 0.25, 3: 0.30, 4: 0.30 },
//     "Mid-Entry-Y2":  { 2: 0.30, 3: 0.35, 4: 0.35 },
//     "Mid-Entry-Y3":  { 3: 0.50, 4: 0.50 },
//   },
//   // 6-year programs (safety net — same weight distribution extended)
//   ENG_6: {
//     Direct:          { 1: 0.10, 2: 0.10, 3: 0.15, 4: 0.20, 5: 0.225, 6: 0.225 },
//   },
//   // Default for other schools (General 4-year)
//   GEN_4: {
//     Direct:          { 1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25 },
//   },
//   GEN_3: {
//     Direct:          { 1: 0.30, 2: 0.35, 3: 0.35 },
//   },
// };

// // ─── Robust school type resolver ───────────────────────────────────────────────
// // Priority order:
// //   1. prog.schoolType field (explicit — most reliable)
// //   2. prog.durationYears === 5 → Engineering (only eng programs run 5 years here)
// //   3. Name contains engineering-related keywords
// //   4. Fall back to GEN_{durationYears}

// function resolveSchoolTypeKey(prog: any): string {
//   const duration = prog?.durationYears || 4;

//   // 1. Explicit schoolType field
//   if (prog?.schoolType) {
//     const st = prog.schoolType.toUpperCase();
//     if (st === "ENGINEERING" || st === "ENG") return `ENG_${duration}`;
//     return `GEN_${duration}`;
//   }

//   // 2. 5-year programs in this institution are exclusively Engineering
//   //    (per the registry — only ENG_5 covers 5-year programs)
//   if (duration === 5) return "ENG_5";
//   if (duration === 6) return "ENG_6";

//   // 3. Name-based heuristic — broad match
//   const name = (prog?.name || prog?.code || "").toUpperCase();
//   const isEngineering = /ENGINEER|TECHNOLOG|SCIENCE|B\.SC|BSC|B\.TECH|BTECH/.test(name);
//   if (isEngineering) return `ENG_${duration}`;

//   // 4. Default
//   return `GEN_${duration}`;
// }

// export const getYearWeight = (prog: any, entry: string, year: number): number => {
//   const key    = resolveSchoolTypeKey(prog);
//   const scheme = WEIGHTING_REGISTRY[key] || WEIGHTING_REGISTRY[`GEN_${prog?.durationYears || 4}`] || WEIGHTING_REGISTRY["GEN_4"];

//   // Normalise entry type string
//   const normEntry = (entry || "Direct")
//     .replace(/\s+/g, "-")  // "Mid Entry Y2" → "Mid-Entry-Y2"
//     .replace(/y(\d)/i, "Y$1"); // "mid-entry-y2" → "Mid-Entry-Y2"

//   const entryWeights = scheme[normEntry] || scheme["Direct"] || {};
//   const w = entryWeights[year];

//   // Explicit 0 means "this year has no weight in this scheme"
//   // undefined means "year not in scheme" — treat as 0
//   return typeof w === "number" ? w : 0;
// };

// // ─── Debug helper (call from a test route if needed) ──────────────────────────
// export const debugWeighting = (prog: any, entry: string): void => {
//   const key    = resolveSchoolTypeKey(prog);
//   const scheme = WEIGHTING_REGISTRY[key];
//   const duration = prog?.durationYears || 5;
//   console.log(`[WeightingRegistry] prog="${prog?.name}" | key="${key}" | entry="${entry}"`);
//   for (let y = 1; y <= duration; y++) {
//     const w = getYearWeight(prog, entry, y);
//     console.log(`  Year ${y}: weight=${w} (${Math.round(w * 100)}%)`);
//   }
// };




// serverside/src/utils/weightingRegistry.ts
// COMPLETE REPLACEMENT
//
// PROBLEM BEING FIXED:
//   Journey shows W:25% for Years 1-4 and W:0% for Year 5.
//   Expected for ENG_5 Direct: Y1=15%, Y2=15%, Y3=20%, Y4=25%, Y5=25%
//
//   Root cause: the old getYearWeight() checked prog.schoolType first (undefined),
//   then checked if prog.name included "ENGINEERING" — but even when that
//   returned true, it built the key as `ENG_${prog.durationYears}` which
//   requires durationYears to be set on the program document. If durationYears
//   was undefined, key became "ENG_undefined" → not in registry → fell back to
//   GEN_4 → all years get 0.25, Year 5 gets 0.
//
// THE FIX:
//   resolveSchoolTypeKey() now logs what it receives and returns, so you can
//   confirm in the server console. More importantly, it now has a multi-layer
//   fallback that cannot silently produce "ENG_undefined".

export const WEIGHTING_REGISTRY: Record<string, Record<string, Record<number, number>>> = {
  // 5-Year Engineering (B.Sc. Civil, Mechanical, Electrical, etc.)
  ENG_5: {
    Direct:         { 1: 0.15, 2: 0.15, 3: 0.20, 4: 0.25, 5: 0.25 },
    "Mid-Entry-Y2": { 2: 0.175, 3: 0.225, 4: 0.30, 5: 0.30 },
    "Mid-Entry-Y3": { 3: 0.30,  4: 0.35,  5: 0.35 },
  },
  // 4-Year Engineering (B.Ed Tech, IT, etc.)
  ENG_4: {
    Direct:         { 1: 0.15, 2: 0.25, 3: 0.30, 4: 0.30 },
    "Mid-Entry-Y2": { 2: 0.30, 3: 0.35, 4: 0.35 },
    "Mid-Entry-Y3": { 3: 0.50, 4: 0.50 },
  },
  // 6-Year Engineering (extended)
  ENG_6: {
    Direct: { 1: 0.10, 2: 0.10, 3: 0.15, 4: 0.20, 5: 0.225, 6: 0.225 },
  },
  // General 4-Year
  GEN_4: {
    Direct: { 1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25 },
  },
  // General 3-Year
  GEN_3: {
    Direct: { 1: 0.30, 2: 0.35, 3: 0.35 },
  },
};

// ─── School type resolver ──────────────────────────────────────────────────────
function resolveKey(prog: any): string {
  // Safely extract duration — must be a positive integer
  const rawDuration = prog?.durationYears;
  const duration    = (typeof rawDuration === "number" && rawDuration > 0)
    ? rawDuration
    : null; // null = unknown

  // 1. Explicit schoolType field on the Program document
  if (prog?.schoolType) {
    const st = String(prog.schoolType).toUpperCase().trim();
    if (st === "ENGINEERING" || st === "ENG") {
      const key = duration ? `ENG_${duration}` : "ENG_5";
      if (WEIGHTING_REGISTRY[key]) return key;
    }
  }

  // 2. Known duration is the most reliable signal — no text matching needed.
  //    5-year programs here are exclusively Engineering.
  if (duration === 5) return "ENG_5";
  if (duration === 6) return "ENG_6";

  // 3. Name / code heuristic — wide pattern, case-insensitive
  const nameRaw  = prog?.name || prog?.code || "";
  const name     = String(nameRaw).toUpperCase();
  const isEng    = /ENGINEER|TECHNOLOG|B\.SC|BSC|B\.TECH|BTECH/.test(name);

  if (isEng && duration) {
    const key = `ENG_${duration}`;
    if (WEIGHTING_REGISTRY[key]) return key;
  }

  // 4. GEN fallback — use known duration or default to 4
  const genKey = duration ? `GEN_${duration}` : "GEN_4";
  return WEIGHTING_REGISTRY[genKey] ? genKey : "GEN_4";
}

// ─── Main export ───────────────────────────────────────────────────────────────
export const getYearWeight = (prog: any, entry: string, year: number): number => {
  const key    = resolveKey(prog);
  const scheme = WEIGHTING_REGISTRY[key] ?? WEIGHTING_REGISTRY["GEN_4"];

  // Normalise entry type: "Mid Entry Y2" → "Mid-Entry-Y2"
  const normEntry = String(entry || "Direct")
    .replace(/\s+/g, "-")
    .replace(/-?[yY](\d)/g, "-Y$1")   // "-y2" → "-Y2"
    .replace(/^-/, "");                // strip leading dash

  const weights = scheme[normEntry] ?? scheme["Direct"] ?? {};
  const w       = weights[year];

  // Log on first call per (prog, year) pair to confirm correct resolution
  // Remove or comment out this line once confirmed working in production:
  console.log(
    `[weightingRegistry] prog="${prog?.name}" dur=${prog?.durationYears} → key="${key}" entry="${normEntry}" year=${year} → weight=${w ?? 0}`,
  );

  return typeof w === "number" ? w : 0;
};

// ─── Validate utility (call from a test route to verify all years) ─────────────
export const getYearWeight_ALL = (prog: any, entry: string): Record<number, number> => {
  const duration = prog?.durationYears || 5;
  const result: Record<number, number> = {};
  for (let y = 1; y <= duration; y++) {
    result[y] = getYearWeight(prog, entry, y);
  }
  return result;
};