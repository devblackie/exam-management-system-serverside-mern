// serverside/src/scripts/cleanupGrades.ts
import mongoose from "mongoose";
import FinalGrade from "../models/FinalGrade";

export const cleanupOrphanedGrades = async () => {
  console.log("Starting data integrity cleanup...");

  // 1. Delete grades where programUnit field is missing/null
  const nullRefs = await FinalGrade.deleteMany({
    $or: [
      { programUnit: { $exists: false } },
      { programUnit: null }
    ]
  });
  console.log(`🗑️ Removed ${nullRefs.deletedCount} grades with null programUnit references.`);

  // 2. Find grades where the reference exists but the target document is gone
  const allGrades = await FinalGrade.find().populate("programUnit");
  const brokenGrades = allGrades.filter(g => !g.programUnit);
  
  if (brokenGrades.length > 0) {
    const brokenIds = brokenGrades.map(g => g._id);
    const orphanRefs = await FinalGrade.deleteMany({ _id: { $in: brokenIds } });
    console.log(`🗑️ Removed ${orphanRefs.deletedCount} orphaned grades (broken references).`);
  }

  console.log("✅ Cleanup complete. Data integrity restored.");
};

// Add this to the end of cleanupGrades.ts
if (require.main === module) {
  mongoose.connect(process.env.MONGODB_URI!)
    .then(async () => {
      await cleanupOrphanedGrades();
      process.exit(0);
    })
    .catch(err => {
      console.error(err);
      process.exit(1);
    });
}