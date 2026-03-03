import Student from "../models/Student";
import { differenceInYears, format, } from "date-fns";

// --- IMPLEMENTING ACADEMIC LEAVE (ENG. 19) ---
export const grantAcademicLeave = async ( 
  studentId: string, startDate: Date, endDate: Date, reason: string, leaveType: "compassionate" | "financial" | "other") => {
  const student = await Student.findById(studentId).populate("program");
  if (!student) throw new Error("Student not found");

  // Calculate duration to check max stay-out
  const yearsRequested = differenceInYears(endDate, startDate);

  // ENG 19.d/e: Max stay-out check (10 yrs Eng / 8 yrs Ed)
  const program = student.program as any;
  const maxStayOut = program.name.includes("Engineering") ? 10 : 8;

  if ((student.totalTimeOutYears || 0) + yearsRequested > maxStayOut) {
    throw new Error(`Leave denied: Exceeds max stay-out period of ${maxStayOut} years.`);
  }

  const dateRange = `${format(startDate, "MMM yyyy")} - ${format(endDate, "MMM yyyy")}`;

  // Update student status and track time
  return await Student.findByIdAndUpdate(studentId, {
      // $set: {
      //   status: "on_leave",
      //   academicLeavePeriod: { startDate, endDate, reason, type: leaveType },
      //   totalTimeOutYears: (student.totalTimeOutYears || 0) + yearsRequested,
      // },
      $set: {
        status: "on_leave",
        remarks: `Academic Leave: ${reason} (${dateRange})`,
        academicLeavePeriod: { startDate, endDate, reason, type: leaveType },
        totalTimeOutYears: (student.totalTimeOutYears || 0) + yearsRequested,
      },
      $push: {
        promotionHistory: {
          from: student.currentYearOfStudy,
          to: student.currentYearOfStudy, // Stay in same year
          date: new Date(),
          remarks: `GRANTED ACADEMIC LEAVE: ${reason}`
        }
      }
    },
    { new: true },
  );
};

// --- IMPLEMENTING DEFERMENT (ENG. 20) ---
export const deferAdmission = async (studentId: string, academicYearsToDefer: 1 | 2) => {
  const student = await Student.findById(studentId);
  if (!student || student.currentYearOfStudy !== 1) throw new Error("Only new students can defer admission.");  

  // ENG 20.a: Senate approval check (usually via UI workflow)
  const endDate = new Date();
  endDate.setFullYear(endDate.getFullYear() + academicYearsToDefer);

  return await Student.findByIdAndUpdate(
    studentId,
    {
      $set: {
        status: "deferred",
        remarks: `Deferred for ${academicYearsToDefer} year(s).`,
        // Set end date based on academic year
        academicLeavePeriod: { startDate: new Date(), endDate: endDate, reason: "Admission Deferral", type: "other" },
      },
      $push: {
        promotionHistory: {
          from: 0, // 0 indicates pre-admission/entry
          to: 1,
          date: new Date(),
          remarks: `ADMISSION DEFERRED: ${academicYearsToDefer} Year(s)`
        }
      }
    },
    { new: true },
  );
};

// --- UNDO/REVERT STATUS (For both Leave/Defer) ---
export const revertStatusToActive = async (studentId: string) => {
  const now = new Date();
  const student = await Student.findById(studentId);
  if (!student) throw new Error("Student not found");

  const previousStatus = student.status.toUpperCase();

  return await Student.findByIdAndUpdate(studentId, {
    $set: { 
      status: "active", 
      remarks: `Status manually reverted to active from ${previousStatus} on ${now.toDateString()}.` 
    },
    $push: {
      promotionHistory: {
        from: student.currentYearOfStudy,
        to: student.currentYearOfStudy,
        date: now,
        remarks: `MANUAL: RETURNED FROM ${previousStatus}`
      }
    },
    $unset: { academicLeavePeriod: 1 }
  }, { new: true });
};
