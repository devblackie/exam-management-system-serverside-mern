// serverside/src/server.ts
import app from "./app";
import connectDB from "./config/db";
import config from "./config/config";
import { ensureDefaultInstitution } from "./config/defaultData"; // ← ADD THIS
import mongoose from "mongoose";

const PORT = config.port || 8000;

const startServer = async () => {
  try {

    // 1. Connect to MongoDB
    await connectDB();   
    
       // WAIT FOR MONGOOSE TO BE FULLY READY
    await mongoose.connection.once("connected", async () => {
      console.log("Mongoose fully initialized");
      await ensureDefaultInstitution();
      console.log("Default data initialized");
    });

    app.listen(PORT, () => {
      console.log(`Server running on http://localhost:${PORT}`);
      console.log(`Frontend: ${config.frontendUrl}`);
      console.log(`Environment: ${process.env.NODE_ENV || "development"}`);
    });
  } catch (error) {
    console.error("Failed to start server:", error);
    process.exit(1);
  }
};

startServer();





