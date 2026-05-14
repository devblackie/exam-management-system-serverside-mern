// serverside/src/models/DisciplinaryCase.ts
//
// WHY THIS EXISTS
// ──────────────a─
// The previous codebase had `isDisciplinary` flags and `RP1D` qualifier logic
// in academicRules.ts but zero backend machinery to actually CREATE a disciplinary
// case, record a hearing outcome, or suspend a student. A coordinator had no way
// to formally "send a student home" in the system — they could only manually edit
// the student's status field with no audit trail.
//
// This model gives that process a proper home. Every suspension, every hearing,
// every appeal lives here. The Student model is updated as a side-effect.
//
// HOW IT LINKS TO THE REST OF THE SYSTEM
// ───────────────────────────────────────
// 1. DisciplinaryCase created → student.status set to "disciplinary_suspension"
// 2. statusEvents entry written to Student (picked up by JourneyTimeline)
// 3. requireAuth middleware blocks suspended students on every API call
// 4. If outcome = "SENT_HOME", qualifierSuffix is set to RP1D via deriveQualifierSuffix
// 5. If outcome = "REINSTATED", student.status reverts to prior status
// 6. AuditLog entry written for every state change

import mongoose, { Schema, Document } from "mongoose";

export type DisciplinaryGrounds =
  | "exam_irregularity"       // ENG.17 — cheating, impersonation
  | "academic_misconduct"     // plagiarism, falsifying records
  | "misconduct"              // general misconduct on campus
  | "financial"               // fees-related (rare but happens)
  | "other";

export type DisciplinaryOutcome =
  | "PENDING"                 // case filed, hearing not yet held
  | "WARNING"                 // verbal/written warning, no suspension
  | "SENT_HOME"               // suspended — blocked from campus & system
  | "REINSTATED"              // suspension lifted, returned to studies
  | "DISCONTINUED"            // case resulted in full discontinuation (ENG.22)
  | "DISMISSED";              // case found baseless, no action

export interface IDisciplinaryCase extends Document {
  institution:      mongoose.Types.ObjectId;
  student:          mongoose.Types.ObjectId;
  raisedBy:         mongoose.Types.ObjectId;   // coordinator/admin who filed the case
  academicYear:     mongoose.Types.ObjectId;
  yearOfStudy:      number;

  grounds:          DisciplinaryGrounds;
  description:      string;                    // narrative of the incident

  hearingDate?:     Date;
  outcome:          DisciplinaryOutcome;
  outcomeNotes?:    string;                    // what the disciplinary committee decided
  resolvedBy?:      mongoose.Types.ObjectId;   // admin who recorded the outcome
  resolvedAt?:      Date;

  // Appeal tracking
  appealed:         boolean;
  appealDate?:      Date;
  appealOutcome?:   "UPHELD" | "DISMISSED";    // upheld = student wins, dismissed = outcome stands
  appealNotes?:     string;

  // If SENT_HOME: when are they expected back?
  suspensionStart?: Date;
  suspensionEnd?:   Date;                      // null = indefinite until Senate decides

  // Snapshot of student status before this case changed it (for reverting)
  priorStudentStatus: string;

  createdAt:        Date;
  updatedAt:        Date;
}

const schema = new Schema<IDisciplinaryCase>(
  {
    institution:    { type: Schema.Types.ObjectId, ref: "Institution", required: true },
    student:        { type: Schema.Types.ObjectId, ref: "Student",     required: true },
    raisedBy:       { type: Schema.Types.ObjectId, ref: "User",        required: true },
    academicYear:   { type: Schema.Types.ObjectId, ref: "AcademicYear",required: true },
    yearOfStudy:    { type: Number, required: true, min: 1, max: 8 },

    grounds: {
      type: String,
      enum: ["exam_irregularity","academic_misconduct","misconduct","financial","other"],
      required: true,
    },
    description:    { type: String, required: true, minlength: 10 },

    hearingDate:    { type: Date },
    outcome: {
      type:    String,
      enum:    ["PENDING","WARNING","SENT_HOME","REINSTATED","DISCONTINUED","DISMISSED"],
      default: "PENDING",
    },
    outcomeNotes:   { type: String },
    resolvedBy:     { type: Schema.Types.ObjectId, ref: "User" },
    resolvedAt:     { type: Date },

    appealed:       { type: Boolean, default: false },
    appealDate:     { type: Date },
    appealOutcome:  { type: String, enum: ["UPHELD","DISMISSED"] },
    appealNotes:    { type: String },

    suspensionStart:{ type: Date },
    suspensionEnd:  { type: Date },

    priorStudentStatus: { type: String, required: true, default: "active" },
  },
  { timestamps: true }
);

// Indexes
schema.index({ institution: 1, student: 1 });
schema.index({ institution: 1, outcome: 1 });
schema.index({ institution: 1, academicYear: 1 });
schema.index({ student: 1, createdAt: -1 });

export default mongoose.model<IDisciplinaryCase>("DisciplinaryCase", schema);