// serverside/src/lib/multiTenant.ts
import { Types } from "mongoose";

export const scopeQuery = (req: any, query: any = {}) => {
  const institutionId = req.user?.institution;

  if (req.user?.role === "admin" && !institutionId) {
    return query;
  }

  

  if (!institutionId) {
    throw new Error("MULTI_TENANT_VIOLATION: Institution context missing");
  }

  // 🛡️ FIX: Ensure we are using an ObjectId for the query
  try {
    const validId = typeof institutionId === 'string' 
      ? new Types.ObjectId(institutionId) 
      : institutionId;

    return { ...query, institution: validId };
  } catch (error) {
    throw new Error("MULTI_TENANT_VIOLATION: Invalid Institution ID format");
  }
};
