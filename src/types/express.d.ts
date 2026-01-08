// src/types/express.d.ts
import { Types } from "mongoose";
import { IUser } from "../models/User";
import { Buffer } from "buffer";

// Safe user without password
export type UserSafe = Omit<IUser, "password"> & {
  _id: Types.ObjectId;
  // institution?: Types.ObjectId; 
   institution?: Types.ObjectId | null; 
};

// Augment Express Request to include user
declare global {
  namespace Express {
    // Base request — user may be undefined
    interface Request {
      user?: UserSafe;
    }

    // After requireAuth — user is guaranteed
    interface AuthenticatedRequest extends Request {
      user: UserSafe & { institution: Types.ObjectId }; // non-nullable
    }

    namespace Multer {
      interface File {
        buffer: Buffer;
      }
    }
  }
}

export {};