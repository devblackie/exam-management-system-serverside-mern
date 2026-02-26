// src/utils/weightingRegistry.ts

export const WEIGHTING_REGISTRY: any = {
  // Table 1: 5-Year (B.Sc Engineering)
  "ENG_5": {
    "Direct":       { 1: 0.15, 2: 0.15, 3: 0.20, 4: 0.25, 5: 0.25 },
    "Mid-Entry-Y2": { 2: 0.175, 3: 0.225, 4: 0.30, 5: 0.30 },
    "Mid-Entry-Y3": { 3: 0.30, 4: 0.35, 5: 0.35 }
  },
  // Table 2: 4-Year (B.Ed Tech or IT)
  "ENG_4": {
    "Direct":       { 1: 0.15, 2: 0.25, 3: 0.30, 4: 0.30 },
    "Mid-Entry-Y2": { 2: 0.30, 3: 0.35, 4: 0.35 },
    "Mid-Entry-Y3": { 3: 0.50, 4: 0.50 }
  },
  // Default for other schools (General 4-year)
  "GEN_4": {
    "Direct": { 1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25 }
  }
};

export const getYearWeight = (prog: any, entry: string, year: number) => {
  const key = prog.schoolType === "ENGINEERING" ? `ENG_${prog.durationYears}` : `GEN_${prog.durationYears}`;
  const scheme = WEIGHTING_REGISTRY[key] || WEIGHTING_REGISTRY["GEN_4"];
  const entryWeights = scheme[entry] || scheme["Direct"];
  return entryWeights[year] || 0;
};