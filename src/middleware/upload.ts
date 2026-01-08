// src/middleware/upload.ts
import multer from "multer";
import path from "path";

const storage = multer.memoryStorage();

export const uploadMarksFile = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (![".csv", ".xlsx", ".xls"].includes(ext)) {
      return cb(new Error("Only CSV and Excel files allowed"));
    }
    cb(null, true);
  },
});