function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  Logger.log(`[doPost] Request received`);
  Logger.log(`[doPost] Raw event: ${JSON.stringify(e)}`);

  const action = e && e.parameter && e.parameter.action ? e.parameter.action : null;
  Logger.log(`[doPost] Action: ${action}`);

  let payload;
  try {
    payload = JSON.parse(e.postData.contents || '{}');
    Logger.log(`[doPost] Payload: ${JSON.stringify(payload)}`);
  } catch (parseErr) {
    Logger.log(`[doPost] Failed to parse JSON payload: ${parseErr.message}`);
    return jsonError('Invalid JSON payload.');
  }

  const email = Session.getActiveUser().getEmail();
  Logger.log(`[doPost] Active User Email: ${email}`);

  try {
    switch (action) {
      case 'getUserByEmail':
        const user = getUserByEmail(email);
        Logger.log(`[doPost:getUserByEmail] Found user: ${JSON.stringify(user)}`);
        return jsonSuccess(user);

      case 'getAllUsers':
        const users = getAllUsers();
        Logger.log(`[doPost:getAllUsers] Found ${users.length} users`);
        return jsonSuccess(users);

      case 'createUser':
        const creationResult = createUser(payload);
        Logger.log(`[doPost:createUser] ${creationResult}`);
        return jsonSuccess(creationResult);

      case 'updateUser':
        const updateResult = updateUser(payload.userId, payload.updates);
        Logger.log(`[doPost:updateUser] ${updateResult}`);
        return jsonSuccess(updateResult);

      default:
        Logger.log(`[doPost] Unknown action: ${action}`);
        return jsonError(`Unknown action: ${action}`);
    }
  } catch (err) {
    Logger.log(`[doPost] Error in action "${action}": ${err.message}`);
    return jsonError(err.message || 'Unknown error');
  }
}

function jsonSuccess(data) {
  return ContentService.createTextOutput(JSON.stringify({ success: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonError(message) {
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function forceUserAuth() {
  Logger.log('🔐 [forceUserAuth] Function called');

  const email = Session.getActiveUser().getEmail();
  Logger.log(`📧 [forceUserAuth] Retrieved Email: "${email}"`);

  if (!email) {
    Logger.log('⚠️ [forceUserAuth] Email is empty — user likely has not authorized access');
    throw new Error(
      'You must authorize the app to continue. Please visit the login link directly to grant permission.'
    );
  }

  // 🔍 Check if ScriptVerified = TRUE in Users sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName("Users");
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf("Email");
  const verifiedCol = headers.indexOf("ScriptVerified");

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email) {
      if (data[i][verifiedCol] === true) {
        Logger.log("✅ [forceUserAuth] Script already verified");
        return email;
      }

      // ✅ Mark as verified
      usersSheet.getRange(i + 1, verifiedCol + 1).setValue(true);
      Logger.log("🆕 [forceUserAuth] Marked script as verified");
      return email;
    }
  }

  Logger.log("❌ [forceUserAuth] User not found in sheet");
  throw new Error("User not found. Please ensure you are registered.");
}
