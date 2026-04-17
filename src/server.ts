// serverside/src/server.ts
// Force reliable DNS resolver (fixes querySrv ETIMEOUT on Windows/Node 18+)
// import dns from "node:dns";
// dns.setServers(["8.8.8.8", "1.1.1.1", "8.8.4.4"]);

import app from "./app";
import connectDB from "./config/db";
import config from "./config/config";
import { ensureDefaultInstitution } from "./config/defaultData";
import mongoose from "mongoose";
import { cleanupOrphanedGrades } from "./scripts/cleanupGrades";
import { startStatusReversionJob } from "./jobs/defermentCron";
const PORT = config.port || 8000;

const startServer = async () => {
  try {
    // 1. Connect to MongoDB
    await connectDB();   
    
    // WAIT FOR MONGOOSE TO BE FULLY READY
    await mongoose.connection.once("connected", async () => {
      console.log("Mongoose fully initialized");
      await cleanupOrphanedGrades();
      await ensureDefaultInstitution();
      console.log("Default data initialized");
      startStatusReversionJob();
    });

   // 2. Capture the server instance
    const server = app.listen(PORT, () => {
      console.log(`Server running on http://localhost:${PORT}`);
      console.log(`Frontend: ${config.frontendUrl}`);
      console.log(`Environment: ${process.env.NODE_ENV || "devepment"}`);
    });



    // 3. SET GLOBAL TIMEOUTS
    // This prevents the underlying TCP socket from closing during long report generations
    server.timeout = 600000;      // 10 minutes
    server.keepAliveTimeout = 610000; // Slightly higher than timeout to prevent race conditions
    server.headersTimeout = 620000;   // Higher than keepAliveTimeout

  } catch (error) {
    console.error("Failed to start server:", error);
    process.exit(1);
  }
};

startServer();