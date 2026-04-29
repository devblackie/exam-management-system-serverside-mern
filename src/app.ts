// serverside/src/app.ts
import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import compression from "compression";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import dotenv from "dotenv";
import config from "./config/config";
import { errorHandler } from "./middleware/errorHandler";
import { securityHeaders, additionalSecurityHeaders, sanitizeInput, apiLimiter } from "./middleware/security";
import { attachCsrfToken, csrfProtection } from "./middleware/csrf";

// Routes
import authRoutes from "./routes/auth";
import adminRoutes from "./routes/admin";
import auditLogsRoutes from "./routes/auditLogs";
import programsRoutes from "./routes/programs";
import unitsRoutes from "./routes/units";
import coordinatorRoutes from "./routes/coordinator";
import marksRoutes from "./routes/marks";
import institutionsRoutes from "./routes/institutions";
import studentsRoutes from "./routes/students";
import academicYearsRoutes from "./routes/academicYears";
import institutionSettingsRoutes from "./routes/institutionSettings";
import studentSearchRoutes from "./routes/studentSearch";
import programUnitsRouter from './routes/programUnits';
import promoteRoutes from "./routes/promote";
import maintenanceRoutes from "./routes/maintenance";
import billingRoutes from "./routes/billing";
import disciplinaryRoutes from "./routes/disciplinary";

dotenv.config();

const app = express();
app.use(compression());

app.use(securityHeaders); // configured helmet (CSP, HSTS, etc.)
app.use(additionalSecurityHeaders); // Permissions-Policy, Cache-Control, etc.
app.use(helmet());
app.disable("x-powered-by");

const allowedOrigins = [
  config.frontendUrl,
  "http://localhost:3000",
  "http://127.0.0.1:3000",
  // Add LAN IPs here for local dev — remove in production
  ...(process.env.NODE_ENV !== "production"
    ? [
        "http://192.168.1.10:3000",
        "http://10.105.149.124:3000",
        "http://192.168.17.124:3000",
      ]
    : []),
].filter(Boolean) as string[];


// const allowedOrigins = [
//   config.frontendUrl,
//   "http://localhost:8000",
//   "http://127.0.0.1:3000",
//   "http://192.168.1.10:3000",
//   "http://10.105.149.124:3000",
//   "http://192.168.17.124:3000",
// ];
app.use(
  cors({
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.includes(origin)) {
        callback(null, true);
      } else {
        callback(new Error("CORS Blocking: Unauthorized Origin"));
      }
    },

    credentials: true,
    methods: ["GET","POST","PUT","PATCH","DELETE","OPTIONS"],
    allowedHeaders: ["Content-Type","Authorization","X-CSRF-Token"],
  }),
);

app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ limit: '5mb', extended: true }));


app.use(cookieParser());
app.use(sanitizeInput);
app.use(attachCsrfToken);
// app.use(csrfProtection);

app.use((req, res, next) => {
  const uploadPaths = [
    "/marks/upload",
    "/students/bulk", // bulk student registration also uses multipart
    "/students/template",
  ];
  const isUpload = uploadPaths.some((p) => req.path.startsWith(p));
  if (isUpload) return next();
  return csrfProtection(req, res, next);
});

app.use(apiLimiter); // 120 req/min per IP globally

// Rate limiting (per IP)
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100,
  standardHeaders: true,
  legacyHeaders: false,
  message: { message: "Too many requests from this IP. Please try again later." },
});
app.use("/auth/", limiter); // Heavy on login attempts
app.use(
  "/marks/upload",
  rateLimit({
    windowMs: 60 * 60 * 1000,
    max: 50,
    message: { message: "Upload limit reached. Try again in an hour." },
    standardHeaders: true,
    legacyHeaders: false,
  }),
); // 50 uploads/hour

// Health check
app.get("/health", (req, res) => {
  res.status(200).json({ status: "OK", timestamp: new Date().toISOString(), uptime: process.uptime()});
});

// API Routes
app.use("/auth", authRoutes);
app.use("/admin", adminRoutes);
app.use("/audit-logs", auditLogsRoutes);
app.use("/programs", programsRoutes);
app.use("/units", unitsRoutes);
app.use("/coordinator", coordinatorRoutes);
app.use("/marks", marksRoutes);
app.use("/institutions", institutionsRoutes);
app.use("/students", studentsRoutes);
app.use("/academic-years", academicYearsRoutes);
app.use("/institution-settings", institutionSettingsRoutes);
app.use("/student", studentSearchRoutes);
app.use('/program-units', programUnitsRouter);
app.use("/promote", promoteRoutes);
app.use("/maintenance", maintenanceRoutes);
app.use("/billing", billingRoutes);
app.use("/disciplinary", disciplinaryRoutes);

app.use((req, res) => {
  res.status(404).json({ message: `Route ${req.originalUrl} not found`, method: req.method });
});

// Global error handler
app.use(errorHandler);

export default app;
