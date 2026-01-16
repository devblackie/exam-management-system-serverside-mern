// serverside/src/app.ts
import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import dotenv from "dotenv";
import connectDB from "./config/db";
import config from "./config/config";
import { errorHandler } from "./middleware/errorHandler";
import { logAudit } from "./lib/auditLogger";

// Routes
import authRoutes from "./routes/auth";
import adminRoutes from "./routes/admin";
import auditLogsRoutes from "./routes/auditLogs";
import programsRoutes from "./routes/programs";
import unitsRoutes from "./routes/units";
import coordinatorRoutes from "./routes/coordinator";
import marksRoutes from "./routes/marks";
import reportsRoutes from "./routes/reports"; // ← UNCOMMENTED & ADDED
import institutionsRoutes from "./routes/institutions";
import studentsRoutes from "./routes/students";
import academicYearsRoutes from "./routes/academicYears";
import institutionSettingsRoutes from "./routes/institutionSettings";
import studentSearchRoutes from "./routes/studentSearch";
import programUnitsRouter from './routes/programUnits';
import promoteRoutes from "./routes/promote";

dotenv.config();

const app = express();
const PORT = config.port || 8000;

// Security & Performance Middleware
app.use(
  helmet({
    contentSecurityPolicy: false, // Allow if using React inline scripts
  })
);

app.use(
  cors({
    // origin: config.frontendUrl || "http://localhost:3000",
     origin: [
      config.frontendUrl || "http://localhost:3000",
      "http://127.0.0.1:3000",
      "http://192.168.1.9:3000",     // ← ADD YOUR IP
      "http://192.168.1.6:3000",     // ← ADD YOUR IP
      "http://10.41.19.124:3000",    // ← ADD THIS TOO (your other device)
    ],
  
    credentials: true,
  })
);

app.use(cookieParser());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

// Rate limiting (per IP)
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 1000,
  standardHeaders: true,
  legacyHeaders: false,
  message: { message: "Too many requests from this IP. Please try again later." },
});
app.use("/auth/", limiter); // Heavy on login attempts
app.use("/marks/upload", rateLimit({ windowMs: 60 * 60 * 1000, max: 50 })); // 50 uploads/hour

// Health check
app.get("/health", (req, res) => {
  res.status(200).json({
    status: "OK",
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
  });
});

// API Routes
app.use("/auth", authRoutes);
app.use("/admin", adminRoutes);
app.use("/audit-logs", auditLogsRoutes);
app.use("/programs", programsRoutes);
app.use("/units", unitsRoutes);
app.use("/coordinator", coordinatorRoutes);
app.use("/marks", marksRoutes);
app.use("/reports", reportsRoutes);
app.use("/institutions", institutionsRoutes);
app.use("/students", studentsRoutes);
app.use("/academic-years", academicYearsRoutes);
app.use("/institution-settings", institutionSettingsRoutes);
app.use("/student", studentSearchRoutes);
app.use('/program-units', programUnitsRouter);
app.use("/promote", promoteRoutes);

app.use((req, res) => {
  res.status(404).json({
    message: `Route ${req.originalUrl} not found`,
    method: req.method,
  });
});

// Global error handler
app.use(errorHandler);

// Connect to MongoDB
connectDB();

export default app;
