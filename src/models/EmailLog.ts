// serverside/src/models/EmailLog.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IEmailLog extends Document {
  institution: mongoose.Types.ObjectId;
  invoiceNumber?: string;
  recipient: string;
  subject: string;
  status: "sent" | "failed";
  errorMessage?: string;
  timestamp: Date;
}

const emailLogSchema = new Schema<IEmailLog>(
  {
    institution: {
      type: Schema.Types.ObjectId,
      ref: "Institution",
      required: true,
    },
    invoiceNumber: { type: String },
    recipient: { type: String, required: true },
    subject: { type: String, required: true },
    status: { type: String, enum: ["sent", "failed"], required: true },
    errorMessage: { type: String },
    timestamp: { type: Date, default: Date.now },
  },
  { timestamps: false },
);

export default mongoose.model<IEmailLog>("EmailLog", emailLogSchema);
