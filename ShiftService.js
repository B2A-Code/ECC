function getEmployeeShifts() {
  try {
    const email = Session.getActiveUser().getEmail();
    Logger.log("📥 [getEmployeeShifts] Called by: " + email);

    const available = getAvailableShifts();
    Logger.log("📤 [getEmployeeShifts] Available shifts retrieved: " + available.length);
    Logger.log("📦 [getEmployeeShifts] Available shifts preview: " + JSON.stringify(available.slice(0, 5)));

    const assigned = getMyShifts();
    Logger.log("📤 [getEmployeeShifts] Assigned shifts retrieved: " + assigned.length);
    Logger.log("📦 [getEmployeeShifts] Assigned shifts preview: " + JSON.stringify(assigned.slice(0, 5)));

    return {
      available,
      assigned
    };
  } catch (err) {
    Logger.log("❌ [getEmployeeShifts] Error occurred: " + err.message);
    throw new Error("Failed to fetch employee shifts: " + err.message);
  }
}

function getAvailableShifts() {
  const email = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(email);

  Logger.log("📥 [getAvailableShifts] Called by: " + email);
  Logger.log("🆔 [getAvailableShifts] Resolved UserID: " + userId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftSheet = ss.getSheetByName("Shifts");
  const data = shiftSheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log("🧭 [getAvailableShifts] Headers: " + JSON.stringify(headers));

  if (!headers || headers.length === 0) {
    throw new Error("⚠️ Shifts sheet header is empty or missing");
  }

  const statusCol = headers.indexOf("Status");
  const assignedCol = headers.indexOf("AssignedUserID");

  if (statusCol === -1 || assignedCol === -1) {
    throw new Error("⚠️ Required columns (Status or AssignedUserID) are missing");
  }

  const userMap = getAllUsersMap();
  Logger.log("📌 [getAvailableShifts] User map loaded with " + Object.keys(userMap).length + " users");

  const normalizeISO = val => {
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? '' : d.toISOString();
    } catch {
      return '';
    }
  };

  const result = data.slice(1).filter(row => {
    const status = row[statusCol];
    const assigned = row[assignedCol];
    return status === 'Offered' && (!assigned || assigned === userId);
  }).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);

    obj.ShiftDate = normalizeISO(obj.ShiftDate);
    obj.StartTime = obj.StartTime || '';
    obj.EndTime = obj.EndTime || '';
    obj.AcceptedTimestamp = normalizeISO(obj.AcceptedTimestamp);
    obj.CompletedTimestamp = normalizeISO(obj.CompletedTimestamp);
    obj.CreatedAt = normalizeISO(obj.CreatedAt);
    obj.UpdatedAt = normalizeISO(obj.UpdatedAt);
    obj.AssignedFullName = userMap[obj.AssignedUserID] || '';

    return obj;
  });

  Logger.log("📦 [getAvailableShifts] Filtered available: " + result.length + " out of " + (data.length - 1));
  Logger.log("🔍 [getAvailableShifts] Preview: " + JSON.stringify(result.slice(0, 3)));

  return result;
}

function getMyShifts() {
  const email = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(email);

  Logger.log("📥 [getMyShifts] Called by: " + email);
  Logger.log("🆔 [getMyShifts] Resolved UserID: " + userId);

  if (!userId) throw new Error("❌ User ID not found for " + email);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log("🧭 [getMyShifts] Headers: " + JSON.stringify(headers));

  const assignedCol = headers.indexOf("AssignedUserID");
  if (assignedCol === -1) {
    throw new Error("⚠️ 'AssignedUserID' column missing from sheet");
  }

  const userMap = getAllUsersMap();
  Logger.log("📌 [getMyShifts] User map loaded with " + Object.keys(userMap).length + " users");

  const normalizeISO = val => {
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? '' : d.toISOString();
    } catch {
      return '';
    }
  };

  const result = data.slice(1).filter(row => {
    return (row[assignedCol] || '').toString().trim() === userId;
  }).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);

    obj.ShiftDate = normalizeISO(obj.ShiftDate);
    obj.StartTime = obj.StartTime || '';
    obj.EndTime = obj.EndTime || '';
    obj.AcceptedTimestamp = normalizeISO(obj.AcceptedTimestamp);
    obj.CompletedTimestamp = normalizeISO(obj.CompletedTimestamp);
    obj.CreatedAt = normalizeISO(obj.CreatedAt);
    obj.UpdatedAt = normalizeISO(obj.UpdatedAt);
    obj.AssignedFullName = userMap[obj.AssignedUserID] || '';

    return obj;
  });

  Logger.log("📦 [getMyShifts] Assigned shifts: " + result.length + " out of " + (data.length - 1));
  Logger.log("🔍 [getMyShifts] Preview: " + JSON.stringify(result.slice(0, 3)));

  return result;
}

function getShiftsForManager() {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);

  Logger.log("📥 [getShiftsForManager] Called by: " + email);
  Logger.log("🧑‍💼 [getShiftsForManager] User object: " + JSON.stringify(user));

  if (!user || (user.Role || '').trim().toLowerCase() !== 'manager') {
    throw new Error('Only managers can view team shifts');
  }

  const dept = (user.Department || '').trim().toLowerCase();
  Logger.log("🏢 [getShiftsForManager] Department: " + dept);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  Logger.log("🧭 [getShiftsForManager] Headers: " + JSON.stringify(headers));

  const userMap = getAllUsersMap();
  Logger.log("📌 [getShiftsForManager] User map loaded with " + Object.keys(userMap).length + " users");

  const normalizeISO = val => {
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? '' : d.toISOString();
    } catch {
      return '';
    }
  };

  const shifts = data.slice(1).filter(row => {
    const shiftDept = (row[headers.indexOf("Department")] || '').trim().toLowerCase();
    return shiftDept === dept;
  }).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);

    obj.ShiftDate = normalizeISO(obj.ShiftDate);
    obj.StartTime = obj.StartTime || '';
    obj.EndTime = obj.EndTime || '';
    obj.AcceptedTimestamp = normalizeISO(obj.AcceptedTimestamp);
    obj.CompletedTimestamp = normalizeISO(obj.CompletedTimestamp);
    obj.CreatedAt = normalizeISO(obj.CreatedAt);
    obj.UpdatedAt = normalizeISO(obj.UpdatedAt);
    obj.AssignedFullName = userMap[obj.AssignedUserID] || '';

    return obj;
  });

  Logger.log(`✅ [getShiftsForManager] Found ${shifts.length} shifts for department: ${dept}`);
  Logger.log("🔍 [getShiftsForManager] Preview: " + JSON.stringify(shifts.slice(0, 3)));

  return shifts;
}

function getEligibleEmployeesForShift(department) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const roleCol = headers.indexOf("Role");
  const deptCol = headers.indexOf("Department");
  const userIdCol = headers.indexOf("UserID");
  const firstCol = headers.indexOf("FirstName");
  const lastCol = headers.indexOf("LastName");

  return data.slice(1)
    .filter(row => row[roleCol] === "Employee" && row[deptCol] === department)
    .map(row => ({
      UserID: row[userIdCol],
      FullName: `${row[firstCol]} ${row[lastCol]}`
    }));
}

function createShiftForEmployees(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const now = new Date().toISOString();
  const managerEmail = Session.getActiveUser().getEmail();
  const managerUser = getUserByEmail(managerEmail);
  const createdBy = managerUser?.UserID || "unknown";

  const newRows = payload.SelectedEmployees.map(userId => {
    const newRow = [];

    headers.forEach(header => {
      switch (header) {
        case "ShiftID":
          newRow.push(`shift-${Utilities.getUuid().slice(0, 8)}`);
          break;
        case "Department":
          newRow.push(payload.Department);
          break;
        case "ShiftDate":
          newRow.push(new Date(payload.ShiftDate));
          break;
        case "StartTime":
          newRow.push(formatTime12Hour(payload.StartTime));
          break;
        case "EndTime":
          newRow.push(formatTime12Hour(payload.EndTime));
          break;
        case "Description":
          newRow.push(payload.Description);
          break;
        case "Status":
          newRow.push("Offered");
          break;
        case "AssignedUserID":
          newRow.push(userId);
          break;
        case "CreatedByUserID":
          newRow.push(createdBy);
          break;
        case "CreatedAt":
        case "UpdatedAt":
          newRow.push(now);
          break;
        case "ShiftType":
          newRow.push(payload.ShiftType);
          break;
        default:
          newRow.push(""); // default empty for other fields
      }
    });

    return newRow;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);

  return { success: true, added: newRows.length };
}

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
      // ✅ Always override AssignedUserID when accepting
      sheet.getRange(i + 1, assignedCol + 1).setValue(userId);
      sheet.getRange(i + 1, statusCol + 1).setValue("Accepted");
      sheet.getRange(i + 1, acceptedTsCol + 1).setValue(new Date().toISOString());
      sheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return { success: true };
    }
  }

  return { success: false, error: "Shift not found or already accepted." };
}

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
      sheet.getRange(row, completedIndex + 1).setValue(new Date().toISOString());
      
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

function mapRowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, i) => {
    obj[header] = row[i];
  });

  // Normalize any fields that look like ISO dates
  obj.ShiftDate = normalizeISO(obj.ShiftDate);
  obj.StartTime = typeof obj.StartTime === 'string' ? obj.StartTime : formatTimeString(obj.StartTime);
  obj.EndTime   = typeof obj.EndTime   === 'string' ? obj.EndTime   : formatTimeString(obj.EndTime);
  obj.AcceptedTimestamp = normalizeISO(obj.AcceptedTimestamp);
  obj.CompletedTimestamp = normalizeISO(obj.CompletedTimestamp);
  obj.CreatedAt = normalizeISO(obj.CreatedAt);
  obj.UpdatedAt = normalizeISO(obj.UpdatedAt);

  return obj;
}

function normalizeISO(dateString) {
  if (!dateString) return '';
  try {
    const d = new Date(dateString);
    return isNaN(d.getTime()) ? '' : d.toISOString();
  } catch (e) {
    return '';
  }
}

function declineShift(shiftId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Shifts");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("ShiftID");
  const statusCol = headers.indexOf("Status");
  const declinedReasonCol = headers.indexOf("DeclinedReason");
  const updatedAtCol = headers.indexOf("UpdatedAt");
  const userIdCol = headers.indexOf("AssignedUserID");
  const deptCol = headers.indexOf("Department");

  const email = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(email);
  const user = getUserByEmail(email);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idCol] === shiftId && row[userIdCol] === userId) {
      sheet.getRange(i + 1, statusCol + 1).setValue("Declined");
      sheet.getRange(i + 1, declinedReasonCol + 1).setValue(reason);
      sheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());

      // 📧 Email manager of same department
      const manager = getDepartmentManager(row[deptCol]);
      if (manager?.Email) {
        MailApp.sendEmail({
          to: manager.Email,
          subject: `⚠️ Shift Declined by ${user.FirstName} ${user.LastName}`,
          htmlBody: `
            <p><strong>${user.FirstName} ${user.LastName}</strong> has declined a shift (ID: <strong>${shiftId}</strong>).</p>
            <p><strong>Reason:</strong> ${reason}</p>
            <p>Please follow up if necessary.</p>
          `
        });
      }

      return { success: true };
    }
  }

  return { success: false, error: "Shift not found or not assigned to user" };
}

function getDepartmentManager(dept) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const roleCol = headers.indexOf("Role");
  const deptCol = headers.indexOf("Department");

  for (let i = 1; i < data.length; i++) {
    if (data[i][roleCol] === "Manager" && data[i][deptCol] === dept) {
      return rowToObject(data[i], headers);
    }
  }

  return null;
}

function getAllUsersMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const userIdCol = headers.indexOf("UserID");
  const firstNameCol = headers.indexOf("FirstName");
  const lastNameCol = headers.indexOf("LastName");

  if ([userIdCol, firstNameCol, lastNameCol].includes(-1)) {
    throw new Error("Missing UserID, FirstName, or LastName columns in Users sheet");
  }

  const map = {};
  data.slice(1).forEach(row => {
    const userId = row[userIdCol];
    const fullName = `${row[firstNameCol]} ${row[lastNameCol]}`.trim();
    if (userId) map[userId] = fullName;
  });

  return map;
}

function formatTime12Hour(timeStr) {
  if (!timeStr) return '';
  const [hours, minutes] = timeStr.split(':').map(Number);
  const date = new Date();
  date.setHours(hours);
  date.setMinutes(minutes);
  date.setSeconds(0);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "hh:mm:ss a");
}

