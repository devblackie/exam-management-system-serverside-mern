// helpers/calc.ts
type Weights = {
  cat1: number; cat2: number; cat3?: number; assignment?: number; practical?: number; exam: number;
};

function calculateFinalScore(marks: Record<string, number | null>, weights: Weights) {
  // normalize missing -> 0 for computation if that's your policy, or treat missing as "missing mark"
  // compute component contributions
  const cat1 = (marks.cat1 ?? 0) * (weights.cat1 / 100);
  const cat2 = (marks.cat2 ?? 0) * (weights.cat2 / 100);
  const cat3 = weights.cat3 ? (marks.cat3 ?? 0) * (weights.cat3 / 100) : 0;
  const assignment = weights.assignment ? (marks.assignment ?? 0) * (weights.assignment / 100) : 0;
  const practical = weights.practical ? (marks.practical ?? 0) * (weights.practical / 100) : 0;
  const exam = (marks.exam ?? 0) * (weights.exam / 100);

  const final = cat1 + cat2 + cat3 + assignment + practical + exam;
  return Number(final.toFixed(2));
}

function determineStatus(finalScore: number, thresholds: { supplementaryPercent: number }) {
  if (finalScore < thresholds.supplementaryPercent) return "supplementary";
  return "pass";
}
