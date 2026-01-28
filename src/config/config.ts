// serverside/src/config/config.ts
import dotenv from "dotenv";

dotenv.config();

const config = Object.freeze({
  port: process.env.PORT || 3000,
  databaseURI: process.env.MONGODB_URI || "mongodb://localhost:2701",
  frontendUrl: process.env.FRONTEND_URL || "http://localhost:3000",
  // jwtSecret: process.env.JWT_SECRET || "please-change-me",
  jwtSecret: process.env.JWT_SECRET!,
  emailUser: process.env.EMAIL_USER || "",
  emailPass: process.env.EMAIL_PASS || "",
  appName: process.env.APP_NAME || "Exam System", 
  instName: process.env.INST_NAME || "My Institution",
  schoolName: process.env.SCHOOL_NAME || "My School",
  registrar: process.env.REGISTRAR || "Registrar Office",
  postalAddress: process.env.POSTAL_ADDRESS || "Postal Address",
  cellPhone: process.env.CELL_PHONE || "Cell Phone",
  schoolEmail: process.env.SCHOOL_EMAIL || "School Email",
});

export default config;
