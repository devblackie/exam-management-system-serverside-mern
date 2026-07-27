"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.cleanupOrphanedGrades = void 0;
// serverside/src/scripts/cleanupGrades.ts
const mongoose_1 = __importDefault(require("mongoose"));
const FinalGrade_1 = __importDefault(require("../models/FinalGrade"));
const cleanupOrphanedGrades = async () => {
    console.log("Starting data integrity cleanup...");
    // 1. Delete grades where programUnit field is missing/null
    const nullRefs = await FinalGrade_1.default.deleteMany({
        $or: [
            { programUnit: { $exists: false } },
            { programUnit: null }
        ]
    });
    console.log(`🗑️ Removed ${nullRefs.deletedCount} grades with null programUnit references.`);
    // 2. Find grades where the reference exists but the target document is gone
    const allGrades = await FinalGrade_1.default.find().populate("programUnit");
    const brokenGrades = allGrades.filter(g => !g.programUnit);
    if (brokenGrades.length > 0) {
        const brokenIds = brokenGrades.map(g => g._id);
        const orphanRefs = await FinalGrade_1.default.deleteMany({ _id: { $in: brokenIds } });
        console.log(`🗑️ Removed ${orphanRefs.deletedCount} orphaned grades (broken references).`);
    }
    console.log("✅ Cleanup complete. Data integrity restored.");
};
exports.cleanupOrphanedGrades = cleanupOrphanedGrades;
// Add this to the end of cleanupGrades.ts
if (require.main === module) {
    mongoose_1.default.connect(process.env.MONGODB_URI)
        .then(async () => {
        await (0, exports.cleanupOrphanedGrades)();
        process.exit(0);
    })
        .catch(err => {
        console.error(err);
        process.exit(1);
    });
}
