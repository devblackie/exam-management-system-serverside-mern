// serverside/src/utils/weightingRegistry.ts

export const WEIGHTING_REGISTRY: any = {
  // Table 1: 5-Year (B.Sc Engineering)
  ENG_5: {
    Direct: { 1: 0.15, 2: 0.15, 3: 0.2, 4: 0.25, 5: 0.25 },
    "Mid-Entry-Y2": { 2: 0.175, 3: 0.225, 4: 0.3, 5: 0.3 },
    "Mid-Entry-Y3": { 3: 0.3, 4: 0.35, 5: 0.35 },
  },
  // Table 2: 4-Year (B.Ed Tech or IT)
  ENG_4: {
    Direct: { 1: 0.15, 2: 0.25, 3: 0.3, 4: 0.3 },
    "Mid-Entry-Y2": { 2: 0.3, 3: 0.35, 4: 0.35 },
    "Mid-Entry-Y3": { 3: 0.5, 4: 0.5 },
  },
  // Default for other schools (General 4-year)
  GEN_4: {
    Direct: { 1: 0.25, 2: 0.25, 3: 0.25, 4: 0.25 },
  },
};


export const getYearWeight = (prog: any, entry: string, year: number) => {
  // 1. Determine the School Type Key
  // Check if schoolType exists, otherwise infer from the name or duration
  let schoolType = prog.schoolType;

  if (!schoolType) {
    if (prog.name?.toUpperCase().includes("ENGINEERING")) {
      schoolType = "ENGINEERING";
    } else {
      schoolType = "GENERAL";
    }
  }

  const key = schoolType === "ENGINEERING" ? `ENG_${prog.durationYears}` : `GEN_${prog.durationYears}`;

  // 2. Retrieve the Scheme
  const scheme = WEIGHTING_REGISTRY[key] || WEIGHTING_REGISTRY["GEN_4"];

  // 3. Get weights for the entry type (Direct, Mid-Entry, etc.)
  const entryWeights = scheme[entry] || scheme["Direct"];

  // 4. Return the specific year weight
  return entryWeights[year] || 0;
};