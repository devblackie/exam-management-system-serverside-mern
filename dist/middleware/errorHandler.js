"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.errorHandler = errorHandler;
function errorHandler(err, req, res, next) {
    console.error(`[ERROR] ${req.method} ${req.url}`, err);
    const status = err.statusCode || 500;
    res.status(status).json({
        success: false,
        message: err.message || "Internal Server Error",
        ...(err.details && { details: err.details }),
    });
}
