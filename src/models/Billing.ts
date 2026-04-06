// serverside/src/models/Billing.ts

import mongoose, { Schema, Document } from "mongoose";

export interface IInvoice {
  id:      string;
  label:   string;
  amount:  number;
  paid:    boolean;
  paidAt?: Date;
  dueAt:   Date;
}

export interface IBilling extends Document {
  institution:        mongoose.Types.ObjectId;
  planName:           string;
  billingCycle:       "monthly" | "annual";
  seatLimit:          number;
  nextInvoiceAmount:  number;
  nextInvoiceDate:    Date;
  invoices:           IInvoice[];
  createdAt:          Date;
  updatedAt:          Date;
}

const invoiceSchema = new Schema<IInvoice>({
  id:      { type: String, required: true },
  label:   { type: String, required: true },
  amount:  { type: Number, required: true },
  paid:    { type: Boolean, default: false },
  paidAt:  { type: Date },
  dueAt:   { type: Date, required: true },
});

const billingSchema = new Schema<IBilling>(
  {
    institution:       { type: Schema.Types.ObjectId, ref: "Institution", required: true, unique: true },
    planName:          { type: String, default: "Institution Pro" },
    billingCycle:      { type: String, enum: ["monthly", "annual"], default: "monthly" },
    seatLimit:         { type: Number, default: 2000 },
    nextInvoiceAmount: { type: Number, required: true },
    nextInvoiceDate:   { type: Date,   required: true },
    invoices:          [invoiceSchema],
  },
  { timestamps: true }
);

export default mongoose.model<IBilling>("Billing", billingSchema);
