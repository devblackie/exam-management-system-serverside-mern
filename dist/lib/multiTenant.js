"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.scopeQuery = void 0;
// serverside/src/lib/multiTenant.ts
const mongoose_1 = require("mongoose");
const scopeQuery = (req, query = {}) => {
    const institutionId = req.user?.institution;
    if (req.user?.role === "admin" && !institutionId) {
        return query;
    }
    if (!institutionId) {
        throw new Error("MULTI_TENANT_VIOLATION: Institution context missing");
    }
    // FIX: Ensure we are using an ObjectId for the query
    try {
        const validId = typeof institutionId === 'string'
            ? new mongoose_1.Types.ObjectId(institutionId)
            : institutionId;
        return { ...query, institution: validId };
    }
    catch (error) {
        throw new Error("MULTI_TENANT_VIOLATION: Invalid Institution ID format");
    }
};
exports.scopeQuery = scopeQuery;
