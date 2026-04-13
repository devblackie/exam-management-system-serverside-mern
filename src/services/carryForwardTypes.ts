// serverside/src/services/carryForwardTypes.ts
export interface CarryForwardUnit {
  programUnitId: string;
  unitCode: string;
  unitName: string;
  fromYear: number;
  fromAcademicYear: string;
  attemptNumber: number; // which attempt this CF is (3rd, 5th etc.)
  qualifier: string; // "RP1C", "RP2C", "RP3C"
  addedAt: Date;
  status: "pending" | "passed" | "failed" | "escalated_to_rpu";
}
