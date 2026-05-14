// serverside/src/models/Institution.ts
import mongoose, { Schema, Document } from "mongoose";

export interface IInstitution extends Document {
  name:          string;   // "University of Nairobi"
  code:          string;   // "UoN" — short code
  abbreviation?: string;   // "UoN"
  address?:      string;
  website?:      string;
  email?:        string;
  phone?:        string;
  country:       string;
  city?:         string;
  isActive:      boolean;
  createdAt:     Date;
  updatedAt:     Date;
}

const InstitutionSchema = new Schema<IInstitution>(
  {
    name:         { type: String, required: true, unique: true, trim: true },
    code:         { type: String, required: true, unique: true, uppercase: true, trim: true },
    abbreviation: { type: String, trim: true },
    address:      { type: String },
    website:      { type: String },
    email:        { type: String },
    phone:        { type: String },
    country:      { type: String, default: "Kenya" },
    city:         { type: String },
    isActive:     { type: Boolean, default: true },
  },
  { timestamps: true },
);

export default mongoose.model<IInstitution>("Institution", InstitutionSchema);