// // src/models/Institution.ts
// import mongoose, { Schema, Document } from "mongoose";

// export interface IInstitution extends Document {
//   name: string;
//   shortName?: string;
//   code: string;                   // e.g. "CE", for Dept. of Civil Engineering
//   isActive: boolean;
//   academicYearStartMonth?: number; // 1–12, e.g. 9 = September intake
// }

// const schema = new Schema<IInstitution>({
//   name: { type: String, required: true },
//   shortName: { type: String, default: "" },
//   code: { type: String, required: true, unique: true, uppercase: true },
//   isActive: { type: Boolean, default: true },
//   academicYearStartMonth: { type: Number, min: 1, max: 12, default: 9 },
// }, { timestamps: true });

// export default mongoose.model<IInstitution>("Institution", schema);



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

InstitutionSchema.index({ name: 1 }, { unique: true });
InstitutionSchema.index({ code: 1 }, { unique: true });

export default mongoose.model<IInstitution>("Institution", InstitutionSchema);