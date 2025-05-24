function getAvatarByUserId(userId) {
  Logger.log(`🖼️ [getAvatarByUserId] Looking up avatar for: ${userId}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  if (!userSheet) {
    Logger.log("❌ Users sheet not found!");
    return '';
  }

  const data = userSheet.getDataRange().getValues();
  const headers = data.shift();
  const idIndex = headers.indexOf("UserID");
  const avatarIndex = headers.indexOf("Avatar");

  if (idIndex === -1 || avatarIndex === -1) {
    Logger.log("❌ Required columns not found (UserID or Avatar)");
    return '';
  }

  for (const row of data) {
    if (row[idIndex] === userId) {
      Logger.log(`✅ Match found: ${row[avatarIndex]}`);
      return row[avatarIndex] || '';
    }
  }

  Logger.log("⚠️ No matching user found");
  return '';
}