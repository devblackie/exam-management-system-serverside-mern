// // serverside/src/models/TempOTP.ts
// import mongoose, { Schema, Document } from "mongoose";

// export interface ITempOTP extends Document {
//   userId:    mongoose.Types.ObjectId;
//   otpHash:   string;
//   expiresAt: Date;
// }

// const schema = new Schema<ITempOTP>({
//   userId:    { type: Schema.Types.ObjectId, ref: "User", required: true, index: true },
//   otpHash:   { type: String, required: true },
//   expiresAt: { type: Date,   required: true },
// });

// // MongoDB TTL index — automatically deletes documents when expiresAt passes
// // No manual cleanup needed
// schema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });

// export default mongoose.model<ITempOTP>("TempOTP", schema);


// serverside/src/models/TempOTP.ts
import mongoose, { Schema, Document } from "mongoose";

export interface ITempOTP extends Document {
  userId:      mongoose.Types.ObjectId;
  otpHash:     string;
  attempts:    number;       // how many wrong guesses so far
  fingerprint: string;       // device fingerprint — OTP is device-bound
  expiresAt:   Date;
}

const schema = new Schema<ITempOTP>({
  userId: {
    type:     Schema.Types.ObjectId,
    ref:      "User",
    required: true,
    index:    true,
  },
  otpHash: {
    type:     String,
    required: true,
  },
  attempts: {
    type:    Number,
    default: 0,
  },
  fingerprint: {
    type:     String,
    required: true,
  },
  expiresAt: {
    type:     Date,
    required: true,
  },
});

// MongoDB TTL — auto-deletes expired OTP records
// No manual cleanup job needed
schema.index({ expiresAt: 1 }, { expireAfterSeconds: 0 });

export default mongoose.model<ITempOTP>("TempOTP", schema);