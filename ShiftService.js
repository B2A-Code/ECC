function getEmployeeShifts() {
  return {
    available: getAvailableShifts(),
    assigned: getMyShifts()
  };
}

/**
 * Get all shifts available to the current user (offered and unassigned)
 */
function getAvailableShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftSheet = ss.getSheetByName("Shifts");
  const data = shiftSheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf("Status");
  const assignedCol = headers.indexOf("AssignedUserID");

  const availableShifts = data.slice(1).filter(row =>
    row[statusCol] === "Offered" && !row[assignedCol]
  ).map(row => mapRowToObject(headers, row));

  return availableShifts;
}

/**
 * Get all shifts assigned to the current user
 */
function getMyShifts() {
  const email = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(email);
  if (!userId) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftSheet = ss.getSheetByName("Shifts");
  const data = shiftSheet.getDataRange().getValues();
  const headers = data[0];
  const assignedCol = headers.indexOf("AssignedUserID");

  const myShifts = data.slice(1).filter(row =>
    row[assignedCol] === userId
  ).map(row => mapRowToObject(headers, row));

  return myShifts;
}

/**
 * Accept a shift offered to the user
 */
function acceptShift(shiftId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("ShiftID");
  const assignedCol = headers.indexOf("AssignedUserID");
  const statusCol = headers.indexOf("Status");
  const acceptedTsCol = headers.indexOf("AcceptedTimestamp");
  const updatedAtCol = headers.indexOf("UpdatedAt");

  const email = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(email);
  if (!userId) return { success: false, error: "User not found." };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idCol] === shiftId && row[statusCol] === "Offered") {
      sheet.getRange(i + 1, assignedCol + 1).setValue(userId);
      sheet.getRange(i + 1, statusCol + 1).setValue("Accepted");
      sheet.getRange(i + 1, acceptedTsCol + 1).setValue(new Date().toISOString());
      sheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return { success: true };
    }
  }

  return { success: false, error: "Shift not found or already accepted." };
}

/**
 * Mark shift complete and record actual hours
 * Also flags that an invoice should be generated
 */
function markShiftComplete(userId, shiftId, actualHours) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const shifts = sheet.getDataRange().getValues();
  const header = shifts[0];
  const idIndex = header.indexOf("ShiftID");
  const userIndex = header.indexOf("AssignedUserID");
  const statusIndex = header.indexOf("Status");
  const hoursIndex = header.indexOf("ActualHoursWorked");
  const completedIndex = header.indexOf("CompletedTimestamp");

  for (let i = 1; i < shifts.length; i++) {
    if (shifts[i][idIndex] === shiftId && shifts[i][userIndex] === userId) {
      const row = i + 1;
      sheet.getRange(row, statusIndex + 1).setValue("Completed");
      sheet.getRange(row, hoursIndex + 1).setValue(actualHours);
      sheet.getRange(row, completedIndex + 1).setValue(new Date());
      
      // ✅ Flag invoice not yet generated (used for future batch runs too)
      const draftIndex = header.indexOf("IsInvoiceDraftGenerated");
      if (draftIndex !== -1) {
        sheet.getRange(row, draftIndex + 1).setValue(false);
      }

      // ⏭️ Trigger draft invoice logic
      return generateDraftInvoiceForShift(shiftId, userId);
    }
  }

  return { success: false, error: "Shift not found or not assigned to user" };
}


/**
 * Utility: Convert row to object
 */
function mapRowToObject(headers, row) {
  const obj = {};
  headers.forEach((key, i) => {
    obj[key] = row[i];
  });
  return obj;
}


