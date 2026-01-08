// src/config/defaultData.ts
import Institution from "../models/Institution";

export const ensureDefaultInstitution = async (retries = 3) => {
  for (let i = 0; i < retries; i++) {
    try {
      const count = await Institution.countDocuments();
      
      if (count === 0) {
        console.log("No institution found. Creating default...");
        const defaultInst = await Institution.create({
          name: "Demo University",
          code: "DEMO",
          isActive: true,
        });
        console.log("Default institution created →", defaultInst.name, defaultInst._id);
        return;
      }

      const active = await Institution.findOne({ isActive: true });
      if (!active) {
        console.log("No active institution. Creating default...");
        await Institution.create({
          name: "Department of Civil Engineering",
          code: "CE",
          isActive: true,
        });
      } else {
        console.log("Active institution found →", active.name);
      }
      return;

    } catch (err: any) {
      console.error(`Attempt ${i + 1} failed:`, err.message);
      if (i === retries - 1) {
        console.error("Failed to ensure default institution after retries");
      } else {
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
    }
  }
};