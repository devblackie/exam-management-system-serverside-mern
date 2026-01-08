export interface GradeResult {
  unitId: string;
  total: number;
  status: "PASS" | "SUPPLEMENTARY" | "RETAKE" | "MISSING";
}

export function evaluateMark(
  total: number | undefined,
  hasMissing: boolean,
  passMark: number = 40
): "PASS" | "SUPPLEMENTARY" | "MISSING" {
  if (hasMissing) return "MISSING";
  if (total === undefined || total < passMark) return "SUPPLEMENTARY";
  return "PASS";
}

export function summarizeYearResults(results: GradeResult[]) {
  const supplementaries = results.filter((r) => r.status === "SUPPLEMENTARY");
  if (supplementaries.length >= 5) {
    return "RETAKE";
  }
  if (results.every((r) => r.status === "PASS")) {
    return "PASS";
  }
  return "SUPPLEMENTARY";
}
