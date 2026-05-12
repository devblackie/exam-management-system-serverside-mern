// // src/types/express.d.ts
// import { Types } from "mongoose";
// import { IUser } from "../models/User";
// import { Buffer } from "buffer";

// // Safe user without password
// export type UserSafe = Omit<IUser, "password"> & {
//   _id: Types.ObjectId;
//   // institution?: Types.ObjectId; 
//    institution?: Types.ObjectId | null; 
// };

// // Augment Express Request to include user
// declare global {
//   namespace Express {
//     // Base request — user may be undefined
//     interface Request {
//       user?: UserSafe;
//     }

//     // After requireAuth — user is guaranteed
//     interface AuthenticatedRequest extends Request {
//       user: UserSafe & { institution: Types.ObjectId }; // non-nullable
//     }

//     namespace Multer {
//       interface File {
//         // buffer: Buffer;  // Original definition
//         buffer: Buffer<ArrayBufferLike>; // Updated definition
//       }
//     }
//   }
// }

// export {};








// src/types/express.d.ts
//
// ── WHY the global user type must NOT reference IUser ────────────────────────
//
// IUser extends mongoose.Document, which carries ~55 internal methods
// ($assertPopulated, $clearModifiedPaths, $clone, …).
//
// auth.ts declares:
//   interface AuthenticatedRequest extends Request { user: ScopedUser }
//
// TypeScript enforces that ScopedUser is assignable to whatever type
// Express.Request.user holds globally. If that global type is:
//   user?: UserSafe    and   UserSafe = Omit<IUser, …>
// then ScopedUser must still satisfy all those Mongoose Document methods —
// which it never will because it is a lean POJO.
//
// The only correct fix is to make the global user slot a PLAIN OBJECT type
// that both UserSafe and ScopedUser can satisfy. We define RequestUser as
// the minimal intersection of fields that any authenticated user always has.
// Both ScopedUser and UserSafe structurally satisfy RequestUser, so
// TypeScript accepts AuthenticatedRequest without complaint.
//
// UserSafe is still exported for use by code that needs the richer shape.
// It is NOT used as the global Express.Request.user type.

import { Types } from "mongoose";
import { IUser }  from "../models/User";

// ── RequestUser ───────────────────────────────────────────────────────────────
// The global slot type. A minimal plain-object interface.
// Both ScopedUser (auth.ts) and UserSafe (below) structurally satisfy this,
// so AuthenticatedRequest.user: ScopedUser is assignable to Request.user.
export interface RequestUser {
  _id:         Types.ObjectId;
  name:        string;
  email:       string;
  role:        "admin" | "coordinator" | "lecturer";
  status:      "active" | "suspended";
  institution?: Types.ObjectId;
}

// ── UserSafe ──────────────────────────────────────────────────────────────────
// Richer shape used by legacy helpers that need more fields.
// Import UserSafe directly where you need it — do NOT rely on req.user being
// this type after requireAuth; use AuthenticatedRequest from auth.ts instead.
export type UserSafe = Omit<
  IUser,
  | "password"
  | "passwordResetToken"
  | "passwordResetExpires"
  | "twoFactorSecret"
  | "twoFactorTempToken"
  | "twoFactorTempExpires"
> & {
  _id: Types.ObjectId;
};

// ── Global Express augmentation ───────────────────────────────────────────────
declare global {
  namespace Express {
    interface Request {
      // RequestUser is a plain object → ScopedUser satisfies it → no TS error
      user?: RequestUser;
    }

    namespace Multer {
      interface File {
        buffer: Buffer<ArrayBufferLike>;
      }
    }
  }
}

export {};