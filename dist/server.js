"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/server.ts
const app_1 = __importDefault(require("./app"));
const db_1 = __importDefault(require("./config/db"));
const config_1 = __importDefault(require("./config/config"));
const defaultData_1 = require("./config/defaultData");
const mongoose_1 = __importDefault(require("mongoose"));
const cleanupGrades_1 = require("./scripts/cleanupGrades");
const defermentCron_1 = require("./jobs/defermentCron");
const PORT = config_1.default.port || 8000;
const startServer = async () => {
    try {
        // 1. Connect to MongoDB
        await (0, db_1.default)();
        console.log("Mongoose fully initialized");
        // 2. Run startup tasks — MUST complete before server listens
        //    This guarantees the institution exists before any request arrives
        await (0, cleanupGrades_1.cleanupOrphanedGrades)();
        await (0, defaultData_1.ensureDefaultInstitution)();
        console.log("Default data initialized");
        // 3. Start background jobs (don't block server start)
        (0, defermentCron_1.startStatusReversionJob)();
        // 4. Start server only AFTER all setup is complete
        const server = app_1.default.listen(PORT, () => {
            console.log(`Server running on http://localhost:${PORT}`);
            console.log(`Frontend: ${config_1.default.frontendUrl}`);
            console.log(`Environment: ${process.env.NODE_ENV || "development"}`);
        });
        // 5. Global timeouts for long operations (reports, uploads)
        server.timeout = 600000; // 10 minutes
        server.keepAliveTimeout = 610000; // Slightly higher to prevent races
        server.headersTimeout = 620000; // Higher than keepAliveTimeout
        // 6. Graceful shutdown
        const shutdown = async (signal) => {
            console.log(`${signal} received — shutting down gracefully`);
            server.close(async () => {
                await mongoose_1.default.connection.close();
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
    }
    catch (err) {
        console.error("Failed to start server:", err);
        process.exit(1);
    }
};
startServer();
