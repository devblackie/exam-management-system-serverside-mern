import mongoose, { Schema, Document } from "mongoose";

interface IUnitAssignment extends Document {
  lecturer: mongoose.Types.ObjectId;
  unit: mongoose.Types.ObjectId;
  assignedBy: mongoose.Types.ObjectId;
  createdAt: Date;
}

const UnitAssignmentSchema = new Schema<IUnitAssignment>({
  lecturer: { type: Schema.Types.ObjectId, ref: "User", required: true },  // role = lecturer
  unit: { type: Schema.Types.ObjectId, ref: "Unit", required: true },
  assignedBy: { type: Schema.Types.ObjectId, ref: "User", required: true }, // role = admin
  createdAt: { type: Date, default: Date.now }
});

export default mongoose.model<IUnitAssignment>("UnitAssignment", UnitAssignmentSchema);
