// serverside/src/models/Submission.ts
import mongoose, { Schema, Document } from "mongoose";

export type SubmissionStatus = "pending" | "accepted" | "rejected";

export interface ISubmissionRow {
  registrationNo: string;
  studentName?: string;
  cat1?: number | null;
  cat2?: number | null;
  cat3?: number | null;
  assignment?: number | null;
  practical?: number | null;
  exam?: number | null;
  computedTotal?: number | null; // computed preview
  status?: "ok" | "missing"; // simple flag
}

export interface ISubmission extends Document {
  unit: mongoose.Types.ObjectId;
  lecturer: mongoose.Types.ObjectId;
  fileName?: string;
  rows: ISubmissionRow[];
  status: SubmissionStatus;
  createdAt: Date;
}

const RowSchema = new Schema({
  registrationNo: String,
  studentName: String,
  cat1: Number,
  cat2: Number,
  cat3: Number,
  assignment: Number,
  practical: Number,
  exam: Number,
  computedTotal: Number,
  status: String,
}, { _id: false });

const SubmissionSchema = new Schema<ISubmission>({
  unit: { type: Schema.Types.ObjectId, ref: "Unit", required: true },
  lecturer: { type: Schema.Types.ObjectId, ref: "User", required: true },
  fileName: { type: String },
  rows: { type: [RowSchema], default: [] },
  status: { type: String, enum: ["pending", "accepted", "rejected"], default: "pending" }
}, { timestamps: true });

export default mongoose.model<ISubmission>("Submission", SubmissionSchema);
