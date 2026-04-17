// serverside/src/utils/weightingRegistry.ts

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
  // console.log(`[weightingRegistry] prog="${prog?.name}" dur=${prog?.durationYears} → key="${key}" entry="${normEntry}" year=${year} → weight=${w ?? 0}`);

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