// src/middleware/errorHandler.ts
import { Request, Response, NextFunction } from "express";

export interface ApiError extends Error {
  statusCode?: number;
  details?: any;
}

export function errorHandler(
  err: ApiError,
  req: Request,
  res: Response,
  next: NextFunction
) {
  console.error(`[ERROR] ${req.method} ${req.url}`, err);

  const status = err.statusCode || 500;

  res.status(status).json({
    success: false,
    message: err.message || "Internal Server Error",
    ...(err.details && { details: err.details }),
  });
}
