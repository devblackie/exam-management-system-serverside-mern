"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.loadLogoBuffer = loadLogoBuffer;
// serverside/src/utils/loadLogoBuffer.ts
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const loadInstitutionSettings_1 = require("./loadInstitutionSettings");
const cache_1 = require("./cache");
async function loadLogoBuffer(institutionId) {
    const settings = await (0, loadInstitutionSettings_1.loadInstitutionSettings)(institutionId);
    const logoPath = settings.branding.universityLogoPath;
    if (!logoPath)
        return Buffer.alloc(0);
    const cacheKey = `logo:${institutionId}`;
    return (0, cache_1.cached)(cacheKey, async () => {
        const fullPath = path_1.default.join(process.cwd(), logoPath);
        if (!fs_1.default.existsSync(fullPath)) {
            console.warn(`[Logo] File not found: ${fullPath}`);
            return Buffer.alloc(0);
        }
        return fs_1.default.readFileSync(fullPath);
    }, 600);
}
