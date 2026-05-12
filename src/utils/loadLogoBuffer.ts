// serverside/src/utils/loadLogoBuffer.ts
import fs from "fs";
import path from "path";
import { loadInstitutionSettings } from "./loadInstitutionSettings";
import { cached } from "./cache";

export async function loadLogoBuffer(institutionId: string): Promise<Buffer> {
  const settings = await loadInstitutionSettings(institutionId);
  const logoPath = settings.branding.universityLogoPath;
  if (!logoPath) return Buffer.alloc(0);

  const cacheKey = `logo:${institutionId}`;
  return cached(
    cacheKey,
    async () => {
      const fullPath = path.join(process.cwd(), logoPath);
      if (!fs.existsSync(fullPath)) {
        console.warn(`[Logo] File not found: ${fullPath}`);
        return Buffer.alloc(0);
      }
      return fs.readFileSync(fullPath);
    },
    600,
  );
}
