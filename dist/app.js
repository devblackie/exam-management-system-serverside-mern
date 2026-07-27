"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/app.ts
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const cookie_parser_1 = __importDefault(require("cookie-parser"));
const compression_1 = __importDefault(require("compression"));
const helmet_1 = __importDefault(require("helmet"));
const express_rate_limit_1 = __importDefault(require("express-rate-limit"));
const dotenv_1 = __importDefault(require("dotenv"));
const config_1 = __importDefault(require("./config/config"));
const errorHandler_1 = require("./middleware/errorHandler");
const security_1 = require("./middleware/security");
const csrf_1 = require("./middleware/csrf");
// Routes
const auth_1 = __importDefault(require("./routes/auth"));
const admin_1 = __importDefault(require("./routes/admin"));
const auditLogs_1 = __importDefault(require("./routes/auditLogs"));
const programs_1 = __importDefault(require("./routes/programs"));
const units_1 = __importDefault(require("./routes/units"));
const coordinator_1 = __importDefault(require("./routes/coordinator"));
const marks_1 = __importDefault(require("./routes/marks"));
const institutions_1 = __importDefault(require("./routes/institutions"));
const students_1 = __importDefault(require("./routes/students"));
const academicYears_1 = __importDefault(require("./routes/academicYears"));
const institutionSettings_1 = __importDefault(require("./routes/institutionSettings"));
const studentSearch_1 = __importDefault(require("./routes/studentSearch"));
const programUnits_1 = __importDefault(require("./routes/programUnits"));
const promote_1 = __importDefault(require("./routes/promote"));
const maintenance_1 = __importDefault(require("./routes/maintenance"));
const billing_1 = __importDefault(require("./routes/billing"));
const disciplinary_1 = __importDefault(require("./routes/disciplinary"));
dotenv_1.default.config();
const app = (0, express_1.default)();
// ── MUST be the very first app configuration ──────────────────────────────────
// Tells Express to trust the X-Forwarded-Proto header from Nginx.
// Without this:
//   - req.secure is false even on HTTPS
//   - res.cookie({ secure: true }) may not behave correctly
//   - Some security middleware may make wrong decisions
app.set("trust proxy", 1);
// ─────────────────────────────────────────────────────────────────────────────
app.use((0, compression_1.default)());
app.use(security_1.securityHeaders); // configured helmet (CSP, HSTS, etc.)
app.use(security_1.additionalSecurityHeaders); // Permissions-Policy, Cache-Control, etc.
app.use((0, helmet_1.default)());
app.disable("x-powered-by");
const allowedOrigins = [
    config_1.default.frontendUrl,
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
].filter(Boolean);
app.use((0, cors_1.default)({
    origin: (origin, callback) => {
        if (!origin || allowedOrigins.includes(origin)) {
            callback(null, true);
        }
        else {
            console.error(`[CORS] Blocked: ${origin}`);
            callback(new Error("CORS Blocking: Unauthorized Origin"));
        }
    },
    credentials: true,
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization", "X-CSRF-Token"],
    exposedHeaders: ["Content-Disposition"],
}));
app.use(express_1.default.json({ limit: '5mb' }));
app.use(express_1.default.urlencoded({ limit: '5mb', extended: true }));
app.use((0, cookie_parser_1.default)());
app.use(security_1.sanitizeInput);
app.use(csrf_1.attachCsrfToken);
// app.use(csrfProtection);
// app.use((req, res, next) => {
//   const uploadPaths = [
//     "/marks/upload",
//     "/students/bulk", // bulk student registration also uses multipart
//     "/students/template",
//   ];
//   const isUpload = uploadPaths.some((p) => req.path.startsWith(p));
//   if (isUpload) return next();
//   return csrfProtection(req, res, next);
// });
app.use((req, res, next) => {
    // Routes that must bypass CSRF:
    //   1. File uploads (multipart — token can't be sent in the body)
    //   2. SSE streaming endpoints — browser EventSource API cannot set custom headers,
    //      so the X-CSRF-Token header can never be sent by the client
    //   3. Public routes (no session, no token)
    const CSRF_BYPASS_PATHS = [
        // File uploads
        "/marks/upload",
        "/students/bulk",
        "/students/template",
        // SSE streaming report endpoints — EventSource cannot send headers
        "/promote/download-report-progress",
        "/promote/download-cms",
        "/promote/download-journey-cms",
        // Public endpoints
        "/institutions/public",
        "/auth/check-email",
        "/auth/verify-password",
        "/auth/verify-otp",
        "/admin/secret-register",
        "/admin/register",
        "/lead-capture",
        "/api/lead-capture",
    ];
    const isBypassed = CSRF_BYPASS_PATHS.some(p => req.path.startsWith(p));
    if (isBypassed)
        return next();
    return (0, csrf_1.csrfProtection)(req, res, next);
});
app.use(security_1.apiLimiter); // 120 req/min per IP globally
// Rate limiting (per IP)
const limiter = (0, express_rate_limit_1.default)({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100,
    standardHeaders: true,
    legacyHeaders: false,
    message: { message: "Too many requests from this IP. Please try again later." },
});
app.use("/api/auth/", limiter);
app.use("/api/marks/upload", (0, express_rate_limit_1.default)({
    windowMs: 60 * 60 * 1000,
    max: 50,
    message: { message: "Upload limit reached. Try again in an hour." },
    standardHeaders: true,
    legacyHeaders: false,
})); // 50 uploads/hour
// Health check - bypasses CORS
app.get("/health", (req, res) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.status(200).json({ status: "OK", timestamp: new Date().toISOString(), uptime: process.uptime() });
});
// ── API router ───────────────────
const apiRouter = express_1.default.Router();
apiRouter.use("/auth", auth_1.default);
apiRouter.use("/admin", admin_1.default);
apiRouter.use("/audit-logs", auditLogs_1.default);
apiRouter.use("/programs", programs_1.default);
apiRouter.use("/units", units_1.default);
apiRouter.use("/coordinator", coordinator_1.default);
apiRouter.use("/marks", marks_1.default);
apiRouter.use("/institutions", institutions_1.default);
apiRouter.use("/students", students_1.default);
apiRouter.use("/academic-years", academicYears_1.default);
apiRouter.use("/institution-settings", institutionSettings_1.default);
apiRouter.use("/student", studentSearch_1.default);
apiRouter.use("/program-units", programUnits_1.default);
apiRouter.use("/promote", promote_1.default);
apiRouter.use("/maintenance", maintenance_1.default);
apiRouter.use("/billing", billing_1.default);
apiRouter.use("/disciplinary", disciplinary_1.default);
app.use("/api", apiRouter);
// 404 handler
app.use((req, res) => {
    res.status(404).json({ message: `Route ${req.originalUrl} not found`, method: req.method });
});
// Global error handler
app.use(errorHandler_1.errorHandler);
exports.default = app;
