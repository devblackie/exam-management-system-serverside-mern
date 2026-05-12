// // serverside/src/models/User.ts
// import mongoose, { Schema, Document, Model, Types } from "mongoose";

// export interface IUser extends Document {
//   _id: Types.ObjectId; // explicitly declare _id type
//   name: string;
//   email: string;
//   password: string;
//   role: "admin" | "coordinator" | "lecturer";
//   status: "active" | "suspended";
//   institution?: mongoose.Types.ObjectId;
//   passwordResetToken?: string;
//   passwordResetExpires?: Date;
//   tokenVersion: number;
//   twoFactorEnabled: boolean;
//   twoFactorSecret: string;
//   twoFactorTempToken: string;
//   twoFactorTempExpires: Date;
// }

// const userSchema = new Schema<IUser>(
//   {
//     name: { type: String, required: true },
//     email: { type: String, required: true, unique: true },
//     password: { type: String, required: true },
//     role: { type: String, enum: ["admin", "coordinator", "lecturer"], required: true },
//     status: { type: String, enum: ["active", "suspended"], default: "active" },
//     institution: { type: Schema.Types.ObjectId, ref: "Institution", required: false },
//     passwordResetToken: { type: String, select: false },
//     passwordResetExpires: { type: Date, select: false },
//     tokenVersion: { type: Number, default: 0,required: true },

//     // Add to existing User schema fields:
// twoFactorEnabled:   { type: Boolean, default: false },
// twoFactorSecret:    { type: String, select: false },   // TOTP secret (base32)
// twoFactorTempToken: { type: String, select: false },   // short-lived token after password verify
// twoFactorTempExpires: { type: Date },
//   },

//   { timestamps: true },
// );

// const User: Model<IUser> = mongoose.model<IUser>("User", userSchema);

// export default User;






// serverside/src/models/User.ts
import mongoose, { Schema, Document, Model, Types } from "mongoose";

export interface IUser extends Document {
  _id:          Types.ObjectId;
  name:         string;
  email:        string;
  password:     string;
  role:         "admin" | "coordinator" | "lecturer";
  status:       "active" | "suspended";
  institution?: mongoose.Types.ObjectId;

  // ── Org scoping ─────────────────────────────────────────────────────────
  // These match codes in InstitutionSettings.schools[].code
  // and InstitutionSettings.schools[].departments[].code
  //
  // coordinator: must have schoolCode + departmentCode — sees only their dept's programs/students
  // admin:       institutionWide = true → sees everything, OR scoped to a school/dept
  // lecturer:    may be scoped to a department
  //
  // A null schoolCode/departmentCode means institution-wide access (admin use case)
  schoolCode?:      string;      // "ENG", "MED" — null = all schools
  departmentCode?:  string;      // "CE", "CS"   — null = all departments
  institutionWide:  boolean;     // if true, ignores schoolCode/departmentCode filters

  passwordResetToken?:    string;
  passwordResetExpires?:  Date;
  tokenVersion:           number;
  twoFactorEnabled:       boolean;
  twoFactorSecret?:       string;
  twoFactorTempToken?:    string;
  twoFactorTempExpires?:  Date;
 
}

const userSchema = new Schema<IUser>(
  {
    name:        { type: String, required: true },
    email:       { type: String, required: true, unique: true, lowercase: true, trim: true },
    password:    { type: String, required: true, select: false },
    role:        { type: String, enum: ["admin","coordinator","lecturer"], required: true },
    status:      { type: String, enum: ["active","suspended"], default: "active" },
    institution: { type: Schema.Types.ObjectId, ref: "Institution" },

    // Org scoping — null means no restriction at that level
    schoolCode:     { type: String, uppercase: true, trim: true, default: null },
    departmentCode: { type: String, uppercase: true, trim: true, default: null },
    institutionWide: { type: Boolean, default: false },

    passwordResetToken:   { type: String, select: false },
    passwordResetExpires: { type: Date,   select: false },
    tokenVersion:         { type: Number, default: 0, required: true },
    twoFactorEnabled:     { type: Boolean, default: false },
    twoFactorSecret:      { type: String,  select: false },
    twoFactorTempToken:   { type: String,  select: false },
    twoFactorTempExpires: { type: Date },
  },
  { timestamps: true },
);

userSchema.index({ institution: 1, role: 1 });
userSchema.index({ email: 1 }, { unique: true });
userSchema.index({ institution: 1, departmentCode: 1 });

const User: Model<IUser> = mongoose.model<IUser>("User", userSchema);
export default User;
