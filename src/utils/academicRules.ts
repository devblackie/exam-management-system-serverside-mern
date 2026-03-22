// serverside/src/utils/academicRules.ts
export const ATTEMPT_NOTATIONS: Record<string, string> = {
  ORDINARY: "B/S",
  SUPPLEMENTARY: "A/S",
  CARRY_FORWARD: "A/CF",
  REPEAT_YEAR: "A/RA",
  STAY_OUT: "A/SO",
  REPEAT_UNIT: "A/RPU",
};

/**
 * Returns the correct label for the 'Attempt' column in marksheets
 */

/**
 * Advanced Attempt Labeling for Engineering (ENG Rules)
 */
export const getAttemptLabel = (attemptCount: number, status: string, regNo?: string): string => {
  const normalizedStatus = status?.toLowerCase();
  const regStr = regNo?.toUpperCase() || "";

  // 1. Check for specific suffixes in Registration Number (Real-life behavior)
  if (regStr.includes("RA1")) return "A/RA1";
  if (regStr.includes("RA2")) return "A/RA2";
  if (regStr.includes("RP2")) return "A/RP2";
  if (regStr.includes("TF1")) return "B/TF1"; // Transfer student

  // 2. Handle Repeat Year Status
  if (normalizedStatus === "repeat year" || normalizedStatus === "repeat") {
     // If attemptCount > 1, they are repeating the repeat
     return `A/RA${attemptCount > 1 ? attemptCount : '1'}`;
  }

  // 3. Handle Leave/Stayout
  if (normalizedStatus === "on_leave" || normalizedStatus === "stayout" || normalizedStatus === "academic leave") {
    return "A/SO";
  }

  // 4. Handle Standard Attempts
  if (attemptCount === 2) return "A/S";   // Supplementary
  if (attemptCount === 3) return "A/CF";  // Carry Forward
  
  return "B/S"; // Default Ordinary (1st Attempt)
};
// export const getAttemptLabel = (attemptCount: number, status: string, regNo: string): string => {
//   const normalizedStatus = status?.toLowerCase();
  
//   if (normalizedStatus === "repeat") return `A/RA${attemptCount > 1 ? attemptCount : ''}`;
//   if (normalizedStatus === "on_leave" || normalizedStatus === "stayout") return "A/SO";
//   if (attemptCount === 2) return "A/S"; // Supplementary
//   if (attemptCount === 3) return "A/CF"; // Carry Forward
  
//   return "B/S"; // Default Ordinary
// };