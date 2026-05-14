// // serverside/src/server.ts

// import app from "./app";
// import connectDB from "./config/db";
// import config from "./config/config";
// import { ensureDefaultInstitution } from "./config/defaultData";
// import mongoose from "mongoose";
// import { cleanupOrphanedGrades } from "./scripts/cleanupGrades";
// import { startStatusReversionJob } from "./jobs/defermentCron";
// const PORT = config.port || 8000;

// const startServer = async () => {
//   try {
//     // 1. Connect to MongoDB
//     await connectDB();   
    
//     // WAIT FOR MONGOOSE TO BE FULLY READY
//     await mongoose.connection.once("connected", async () => {
//       console.log("Mongoose fully initialized");
//       await cleanupOrphanedGrades();
//       await ensureDefaultInstitution();
//       console.log("Default data initialized");
//       startStatusReversionJob();
//     });

//    // 2. Capture the server instance
//     const server = app.listen(PORT, () => {
//       console.log(`Server running on http://localhost:${PORT}`);
//       console.log(`Frontend: ${config.frontendUrl}`);
//       console.log(`Environment: ${process.env.NODE_ENV || "development"}`);
//     });



//     // 3. SET GLOBAL TIMEOUTS
//     // This prevents the underlying TCP socket from closing during long report generations
//     server.timeout = 600000;      // 10 minutes
//     server.keepAliveTimeout = 610000; // Slightly higher than timeout to prevent race conditions
//     server.headersTimeout = 620000;   // Higher than keepAliveTimeout

//   } catch (error) {
//     console.error("Failed to start server:", error);
//     process.exit(1);
//   }
// };

// startServer();







// serverside/src/server.ts
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
    console.log("Mongoose fully initialized");

    // 2. Run startup tasks — MUST complete before server listens
    //    This guarantees the institution exists before any request arrives
    await cleanupOrphanedGrades();
    await ensureDefaultInstitution();
    console.log("Default data initialized");

    // 3. Start background jobs (don't block server start)
    startStatusReversionJob();

    // 4. Start server only AFTER all setup is complete
    const server = app.listen(PORT, () => {
      console.log(`Server running on http://localhost:${PORT}`);
      console.log(`Frontend: ${config.frontendUrl}`);
      console.log(`Environment: ${process.env.NODE_ENV || "development"}`);
    });

    // 5. Global timeouts for long operations (reports, uploads)
    server.timeout = 600000;           // 10 minutes
    server.keepAliveTimeout = 610000;  // Slightly higher to prevent races
    server.headersTimeout = 620000;    // Higher than keepAliveTimeout

    // 6. Graceful shutdown
    const shutdown = async (signal: string) => {
      console.log(`${signal} received — shutting down gracefully`);
      server.close(async () => {
        await mongoose.connection.close();
        console.log("MongoDB connection closed");
        process.exit(0);
      });
      
      // Force exit if graceful shutdown hangs
      setTimeout(() => {
        console.error("Forced shutdown after timeout");
        process.exit(1);
      }, 10000);
    };

    process.on("SIGTERM", () => shutdown("SIGTERM"));
    process.on("SIGINT", () => shutdown("SIGINT"));

  } catch (err) {
    console.error("Failed to start server:", err);
    process.exit(1);
  }
};

startServer();