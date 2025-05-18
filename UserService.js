const USER_SHEET = 'Users';

function jsonSuccess(data) {
  return { success: true, data };
}

function jsonError(message) {
  return { success: false, error: message };
}

function getUserSession() {
  try {
    const email = Session.getActiveUser().getEmail();
    Logger.log(`[getUserSession] Email: ${email}`);
    const user = getUserByEmail(email);
    Logger.log(`[getUserSession] Found user: ${JSON.stringify(user)}`);
    return { success: true, data: user };
  } catch (err) {
    Logger.log(`[getUserSession] Error: ${err.message}`);
    return { success: false, error: err.message };
  }
}

function getUserByEmail(email) {
  Logger.log(`[getUserByEmail] Called with email: ${email}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const emailIndex = headers.indexOf('Email');
  Logger.log(`[getUserByEmail] Email column index: ${emailIndex}`);

  const userIndex = rows.findIndex((row, i) => {
    const match = row[emailIndex] === email;
    if (match) Logger.log(`[getUserByEmail] Match found on row ${i + 1}`);
    return match;
  });

  if (userIndex < 1) {
    Logger.log(`[getUserByEmail] No matching user found for email: ${email}`);
    throw new Error('User not found');
  }

  const userObj = rowToObject(rows[userIndex], headers);
  Logger.log(`[getUserByEmail] Returning user: ${JSON.stringify(userObj)}`);
  return userObj;
}

function getAllUsers() {
  Logger.log('[getAllUsers] ⚙️ Starting user fetch operation');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('[getAllUsers] 📄 Active spreadsheet obtained: ' + ss.getName());

  const sheet = ss.getSheetByName(USER_SHEET);
  if (!sheet) {
    Logger.log(`[getAllUsers] ❌ Sheet "${USER_SHEET}" not found`);
    throw new Error(`Sheet "${USER_SHEET}" not found`);
  }
  Logger.log(`[getAllUsers] ✅ Sheet "${USER_SHEET}" loaded`);

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  Logger.log(`[getAllUsers] 📊 Retrieved ${data.length} rows from sheet`);

  if (data.length < 2) {
    Logger.log('[getAllUsers] ⚠️ Sheet contains headers only or is empty');
    return [];
  }

  const headers = data[0];
  Logger.log(`[getAllUsers] 🧭 Headers: ${JSON.stringify(headers)}`);

  const users = data.slice(1).map((row, i) => {
    const userObj = rowToObject(row, headers);
    Logger.log(`[getAllUsers] 🧍 Processed user at row ${i + 2}: ${JSON.stringify(userObj)}`);
    return userObj;
  });

  Logger.log(`[getAllUsers] ✅ Successfully processed ${users.length} users`);
  return users;
}

function createUser(userData) {
  try {
    Logger.log(`[createUser] Incoming data: ${JSON.stringify(userData)}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET);
    const headers = sheet.getDataRange().getValues()[0];

    const newRow = headers.map(header => {
      let val = userData[header];

      // Format ISO dates
      if (['DateOfBirth', 'CreatedAt', 'UpdatedAt'].includes(header)) {
        if (val) {
          const parsedDate = new Date(val);
          return isNaN(parsedDate) ? '' : parsedDate.toISOString();
        }
        return '';
      }

      // Ensure boolean
      if (header === 'BankDetailsConfirmation') {
        return val === true || val === 'TRUE';
      }

      // Ensure numeric fields
      if (['HourlyRate', 'HolidayEntitlementAccruedHours'].includes(header)) {
        return val ? Number(val) : 0;
      }

      // Everything else
      return val || '';
    });

    const newId = generateId('USR');
    const createdAt = new Date().toISOString();
    const updatedAt = new Date().toISOString();

    newRow[headers.indexOf('UserID')] = newId;
    newRow[headers.indexOf('CreatedAt')] = createdAt;
    newRow[headers.indexOf('UpdatedAt')] = updatedAt;

    sheet.appendRow(newRow);

    Logger.log(`[createUser] ✅ User created with ID: ${newId}`);
    return jsonSuccess({ message: `User created`, id: newId });
  } catch (err) {
    Logger.log(`[createUser] ❌ Error: ${err.message}`);
    return jsonError(err.message);
  }
}

function updateUser(userData) {
  try {
    Logger.log(`[updateUser] Incoming data: ${JSON.stringify(userData)}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USER_SHEET);
    const headers = sheet.getDataRange().getValues()[0];
    const data = sheet.getDataRange().getValues();

    const userId = userData.UserID;
    if (!userId) {
      return jsonError('Missing UserID for update.');
    }

    const rowIndex = data.findIndex(row => row[headers.indexOf('UserID')] === userId);
    if (rowIndex === -1) {
      return jsonError(`User with ID ${userId} not found.`);
    }

    // Format and sanitize the incoming data
    const updatedRow = headers.map(header => {
      let val = userData[header];

      // Handle ISO dates
      if (['DateOfBirth', 'CreatedAt', 'UpdatedAt'].includes(header)) {
        if (val) {
          const parsedDate = new Date(val);
          return isNaN(parsedDate) ? '' : parsedDate.toISOString();
        }
        return '';
      }

      // Handle boolean
      if (header === 'BankDetailsConfirmation') {
        return val === true || val === 'TRUE';
      }

      // Numeric fields
      if (['HourlyRate', 'HolidayEntitlementAccruedHours'].includes(header)) {
        return val ? Number(val) : 0;
      }

      return val || '';
    });

    // Update the UpdatedAt timestamp
    updatedRow[headers.indexOf('UpdatedAt')] = new Date().toISOString();

    // Write back to the sheet (rowIndex + 1 because Sheets is 1-based and first row is headers)
    sheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);

    Logger.log(`[updateUser] ✅ Updated user: ${userId}`);
    return jsonSuccess({ message: `User ${userId} updated.` });

  } catch (err) {
    Logger.log(`[updateUser] ❌ Error: ${err.message}`);
    return jsonError(err.message);
  }
}

function rowToObject(row, headers) {
  const obj = {};
  headers.forEach((h, i) => {
    const val = row[i];

    // Handle date columns
    if (['DateOfBirth', 'CreatedAt', 'UpdatedAt', 'LastLogin'].includes(h) && val instanceof Date) {
      obj[h] = val.toISOString(); // Standard ISO format
    }
    // Handle mobile numbers as strings
    else if (['MobileNumber', 'EmergencyContactPhone'].includes(h)) {
      obj[h] = val ? String(val).padStart(10, '0') : '';
    }
    // Ensure booleans
    else if (h === 'BankDetailsConfirmation') {
      obj[h] = val === true || val === 'TRUE';
    }
    // Default case
    else {
      obj[h] = val;
    }
  });
  return obj;
}

function deleteUserById(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('UserID');

  const rowIndex = data.findIndex((row, idx) => idx > 0 && row[idCol] === userId);
  if (rowIndex === -1) return jsonError("User not found");

  sheet.deleteRow(rowIndex + 1);
  return jsonSuccess({ message: "User deleted" });
}

function generateId(prefix) {
  return `${prefix}_${Utilities.getUuid().slice(0, 8)}`;
}


/**
 * Utility: Get user ID by email
 */
function getUserIdByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  const idCol = headers.indexOf("UserID");

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email) {
      return data[i][idCol];
    }
  }
  return null;
}

