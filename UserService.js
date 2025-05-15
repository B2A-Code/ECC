const USER_SHEET = 'Users';

function getUserSession() {
  try {
    const email = Session.getActiveUser().getEmail();
    const user = getUserByEmail(email); // already working
    return { success: true, data: user }; // ✅ just return this!
  } catch (err) {
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
  Logger.log('[getAllUsers] Fetching all users');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const allUsers = data.slice(1).map(row => rowToObject(row, headers));
  Logger.log(`[getAllUsers] Found ${allUsers.length} users`);
  return allUsers;
}

function createUser(userData) {
  Logger.log(`[createUser] Creating user with data: ${JSON.stringify(userData)}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  const headers = sheet.getDataRange().getValues()[0];

  const newRow = headers.map(h => userData[h] || '');
  const newId = generateId('USR');
  newRow[headers.indexOf('UserID')] = newId;
  newRow[headers.indexOf('CreatedAt')] = new Date();
  newRow[headers.indexOf('UpdatedAt')] = new Date();

  sheet.appendRow(newRow);
  Logger.log(`[createUser] User created with ID: ${newId}`);
  return `User created with ID: ${newId}`;
}

function updateUser(userId, updates) {
  Logger.log(`[updateUser] Updating user ${userId} with: ${JSON.stringify(updates)}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('UserID');

  const rowIndex = data.findIndex(row => row[idIndex] === userId);
  if (rowIndex < 1) {
    Logger.log(`[updateUser] User with ID ${userId} not found`);
    throw new Error('User not found');
  }

  const updatedRow = data[rowIndex].slice();
  for (let key in updates) {
    const colIndex = headers.indexOf(key);
    if (colIndex >= 0) {
      Logger.log(`[updateUser] Updating ${key} at column ${colIndex} to ${updates[key]}`);
      updatedRow[colIndex] = updates[key];
    }
  }

  updatedRow[headers.indexOf('UpdatedAt')] = new Date();
  sheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);

  Logger.log(`[updateUser] User ${userId} updated successfully`);
  return `User ${userId} updated successfully`;
}

function rowToObject(row, headers) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function generateId(prefix) {
  return `${prefix}_${Utilities.getUuid().slice(0, 8)}`;
}