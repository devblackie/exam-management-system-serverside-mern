// // serverside/src/services/programNormalizer.ts
// export function normalizeProgramName(name: string): string {
//   if (!name) return "";

//   return (
//     name
//       .replace(/["“”']/g, "")                      // remove quotes
//       .replace(/\r|\n/g, " ")                      // remove line breaks
//       .replace(/\s+/g, " ")                        // collapse whitespace
//       .replace(/\./g, "")  
//       .replace(/[\r\n"]/g, "")                        // remove dots (BSc. → BSc)
//       .trim()
//       .toLowerCase()

//       // Degrees
//       .replace(/\bbsc\b/g, "bachelor of science")
//       .replace(/\bbs\b/g, "bachelor of science")
//       .replace(/\bbtech\b/g, "bachelor of technology")
//       .replace(/\bmsc\b/g, "master of science")
//       .replace(/\bba\b/g, "bachelor of arts")
//       .replace(/\bma\b/g, "master of arts")
//        .replace(/\bbsc\.?\b/g, "bachelor of science")
//     .replace(/\bb\.sc\.?\b/g, "bachelor of science")
//     .replace(/\bbtech\.?\b/g, "bachelor of technology")
//     .replace(/\bmtech\.?\b/g, "master of technology")
//     .replace(/\bmsc\.?\b/g, "master of science")
//     .replace(/\bba\.?\b/g, "bachelor of arts")
//     .replace(/\bma\.?\b/g, "master of arts")
//     .replace(/\bbeng\.?\b/g, "bachelor of engineering")
//     .replace(/\bmeng\.?\b/g, "master of engineering")
//     .replace(/\bdip\.?\b/g, "diploma in")
//     .replace(/\bcert\.?\b/g, "certificate in")

//       // Common patterns
//       .replace(/\bin\b/g, " in ")                  // normalize spacing around "in"
//       .replace(/\s+/g, " ")                        // cleanup
//   );
// }


export function normalizeProgramName(name: string): string {
  if (!name) return "";

  let normalized = name.toLowerCase();

  // 1. Basic Cleaning: Remove quotes, newlines, and collapse spaces
  normalized = normalized
    .replace(/["“”']/g, "")      // Remove excel/csv quotes
    .replace(/[\r\n]+/g, " ")     // specific fix for the newlines in your input
    .replace(/\s+/g, " ")         // collapse multiple spaces
    .trim();

  // 2. Expand Abbreviations
  // We do this BEFORE removing dots to catch "B.Sc." vs "BSc"
  normalized = normalized
    // Science
    .replace(/\bb\.?sc\.?\b/g, "bachelor science") // expanded without "of" for easier matching
    .replace(/\bm\.?sc\.?\b/g, "master science")
    // Technology
    .replace(/\bb\.?tech\.?\b/g, "bachelor technology")
    .replace(/\bm\.?tech\.?\b/g, "master technology")
    // Arts
    .replace(/\bb\.?a\.?\b/g, "bachelor arts")
    .replace(/\bm\.?a\.?\b/g, "master arts")
    // Engineering
    .replace(/\bb\.?eng\.?\b/g, "bachelor engineering")
    .replace(/\bm\.?eng\.?\b/g, "master engineering")
    // Education
    .replace(/\bb\.?ed\.?\b/g, "bachelor education")
    // Other
    .replace(/\bdip\.?\b/g, "diploma")
    .replace(/\bcert\.?\b/g, "certificate");

  // 3. Remove Punctuation (dots, commas, hyphens)
  normalized = normalized.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "");

  // 4. THE FIX: Remove "Stop Words" (Connector words)
  // This ensures "Bachelor of Science in Marine Engineering" 
  // matches "Bachelor Science Marine Engineering"
  normalized = normalized
    .replace(/\b(in|of|and|the|with|honours|hons)\b/g, "")
    .replace(/\s+/g, " ") // Final cleanup of spaces
    .trim();

  return normalized;
}