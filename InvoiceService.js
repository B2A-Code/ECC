function getMyInvoices() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Invoices");
  const userEmail = Session.getActiveUser().getEmail();
  const userId = getUserIdByEmail(userEmail);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const result = rows
    .filter(row => row[1] === userId)
    .map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      return obj;
    });

  return result;
}

function getInvoicesForUser() {
  const email = Session.getActiveUser().getEmail();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const invoicesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoices');

  const users = usersSheet.getDataRange().getValues();
  const header = users[0];
  const emailIdx = header.indexOf("Email");
  const userIdIdx = header.indexOf("UserID");

  const user = users.find(row => row[emailIdx] === email);
  if (!user) return { success: false, error: "User not found" };

  const userId = user[userIdIdx];

  const invoices = invoicesSheet.getDataRange().getValues();
  const headers = invoices[0];

  const result = invoices.slice(1).filter(row => row[headers.indexOf("UserID")] === userId)
    .map(row => {
      const obj = {};
      headers.forEach((key, i) => obj[key] = row[i]);
      return obj;
    });

  return { success: true, data: result };
}

function getInvoiceItemsByInvoiceId(invoiceId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("InvoiceItems");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIndex = headers.indexOf("InvoiceID");

  const result = data
    .slice(1)
    .filter(row => row[idIndex] === invoiceId)
    .map(row => {
      const item = {};
      headers.forEach((key, i) => (item[key] = row[i]));
      return item;
    });

  return result;
}

function getDepartmentInvoicesForApproval() {
  const email = Session.getActiveUser().getEmail();
  const manager = getUserByEmail(email);
  if (!manager || manager.Role !== 'Manager') {
    throw new Error("Not authorized");
  }

  const invoiceSheet = SpreadsheetApp.getActive().getSheetByName("Invoices");
  const userSheet = SpreadsheetApp.getActive().getSheetByName("Users");
  const invoiceData = invoiceSheet.getDataRange().getValues();
  const userData = userSheet.getDataRange().getValues();

  const headers = invoiceData[0];
  const userHeaders = userData[0];

  const departmentUsers = userData.slice(1).filter(r =>
    r[userHeaders.indexOf("Department")] === manager.Department
  );

  const userIdMap = {};
  departmentUsers.forEach(r => {
    userIdMap[r[userHeaders.indexOf("UserID")]] = `${r[userHeaders.indexOf("FirstName")]} ${r[userHeaders.indexOf("LastName")]}`;
  });

  const result = [];
  invoiceData.slice(1).forEach(r => {
    const status = r[headers.indexOf("Status")];
    const userId = r[headers.indexOf("UserID")];

    if (status === "Submitted" && userIdMap[userId]) {
      result.push({
        InvoiceID: r[headers.indexOf("InvoiceID")],
        InvoiceDate: r[headers.indexOf("InvoiceDate")],
        InvoiceType: r[headers.indexOf("InvoiceType")],
        TotalAmount: r[headers.indexOf("TotalAmount")],
        Status: status,
        EmployeeName: userIdMap[userId],
      });
    }
  });

  return result;
}

function submitInvoice(invoiceId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invoices');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf("InvoiceID");
  const statusIdx = headers.indexOf("Status");
  const updatedIdx = headers.indexOf("UpdatedAt");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idIdx] === invoiceId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue("Submitted");
      sheet.getRange(i + 1, updatedIdx + 1).setValue(new Date().toISOString());
      return { success: true };
    }
  }

  return { success: false, error: "Invoice not found" };
}

function createManualInvoice(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Invoices");
  if (!sheet) return { success: false, error: "Invoices sheet not found" };

  const userEmail = Session.getActiveUser().getEmail();
  const invoiceId = Utilities.getUuid();

  const row = [
    invoiceId,
    getUserIdByEmail(userEmail),
    data.InvoiceDate || '',
    data.InvoiceType || '',
    data.TotalAmount || '',
    data.Status || 'Draft',
    '', // RelatedShiftIDs
    data.DescriptionOrPurpose || '',
    '', // ReceiptFileID
    '', '', '', '', '', '', // approvals, carry over, etc
    data.CreatedAt,
    data.UpdatedAt
  ];

  sheet.appendRow(row);
  return { success: true, id: invoiceId };
}

function createManualInvoice(invoiceData) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const user = getUserByEmail(userEmail);
    if (!user) throw new Error('User not found');

    const sheet = SpreadsheetApp.getActive().getSheetByName('Invoices');
    const newId = Utilities.getUuid();
    const now = new Date();

    const newRow = [
      newId,                       // InvoiceID
      user.UserID,                // UserID
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'), // InvoiceDate
      invoiceData.InvoiceType,    // InvoiceType
      invoiceData.TotalAmount,    // TotalAmount
      "Draft",                    // Status
      "",                         // RelatedShiftIDs
      invoiceData.DescriptionOrPurpose, // Description
      "",                         // ExpenseReceiptFileID
      "", "", "", "", "", "",     // SubmittedToManagerID, ApprovalTimestamps, Payment, Rejection
      now.toISOString(), now.toISOString()
    ];

    sheet.appendRow(newRow);
    return { success: true, id: newId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function generateDraftInvoiceForShift(shiftId, userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftsSheet = ss.getSheetByName("Shifts");
  const usersSheet = ss.getSheetByName("Users");
  const invoicesSheet = ss.getSheetByName("Invoices");
  const invoiceItemsSheet = ss.getSheetByName("InvoiceItems");

  const now = new Date();

  // ---- Get Headers and Indexes
  const shiftData = shiftsSheet.getDataRange().getValues();
  const shiftHeaders = shiftData[0];
  const userData = usersSheet.getDataRange().getValues();
  const userHeaders = userData[0];

  const idIndex = shiftHeaders.indexOf("ShiftID");
  const assignedIndex = shiftHeaders.indexOf("AssignedUserID");
  const hoursIndex = shiftHeaders.indexOf("ActualHoursWorked");
  const rateIndex = userHeaders.indexOf("HourlyRate");
  const draftGeneratedIndex = shiftHeaders.indexOf("IsInvoiceDraftGenerated");
  const invoiceIdIndex = shiftHeaders.indexOf("GeneratedInvoiceID");
  const descIndex = shiftHeaders.indexOf("Description");

  const userIdIndex = userHeaders.indexOf("UserID");

  // ---- Find Shift
  const shiftRow = shiftData.find(row => row[idIndex] === shiftId && row[assignedIndex] === userId);
  if (!shiftRow) return { success: false, error: "Shift not found or not assigned to you." };

  if (shiftRow[draftGeneratedIndex]) {
    return { success: false, error: "Invoice already generated for this shift." };
  }

  const actualHours = parseFloat(shiftRow[hoursIndex]);
  if (isNaN(actualHours) || actualHours <= 0) {
    return { success: false, error: "Invalid Actual Hours Worked" };
  }

  // ---- Find User
  const userRow = userData.find(row => row[userIdIndex] === userId);
  if (!userRow) return { success: false, error: "User not found" };

  const hourlyRate = parseFloat(userRow[rateIndex]);
  if (isNaN(hourlyRate)) return { success: false, error: "User has no hourly rate defined." };

  const totalAmount = parseFloat((hourlyRate * actualHours).toFixed(2));
  const invoiceId = Utilities.getUuid();

  // ---- Create Invoice
  const invoiceRow = [
    invoiceId,
    userId,
    now,
    "ShiftWork",
    totalAmount,
    "Draft",
    shiftId, // RelatedShiftIDs
    "Auto-generated from completed shift",
    "", "", "", "", "", "", "", now, now
  ];
  invoicesSheet.appendRow(invoiceRow);

  // ---- Create InvoiceItem
  const itemId = Utilities.getUuid();
  const itemDesc = shiftRow[descIndex] || "Shift Work";
  const itemRow = [
    itemId,
    invoiceId,
    itemDesc,
    actualHours,
    hourlyRate,
    totalAmount
  ];
  invoiceItemsSheet.appendRow(itemRow);

  // ---- Update Shift row
  const shiftRowIndex = shiftData.findIndex(row => row[idIndex] === shiftId);
  if (shiftRowIndex >= 1) {
    const rowNum = shiftRowIndex + 1;
    shiftsSheet.getRange(rowNum, draftGeneratedIndex + 1).setValue(true);
    shiftsSheet.getRange(rowNum, invoiceIdIndex + 1).setValue(invoiceId);
  }

  return { success: true, invoiceId };
}

function approveOrRejectInvoice(invoiceId, action, reason) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Invoices");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idx = data.findIndex(row => row[headers.indexOf("InvoiceID")] === invoiceId);
  if (idx < 1) throw new Error("Invoice not found");

  const now = new Date();

  if (action === "approve") {
    sheet.getRange(idx + 1, headers.indexOf("Status") + 1).setValue("ManagerApproved");
    sheet.getRange(idx + 1, headers.indexOf("ManagerApprovalTimestamp") + 1).setValue(now);
  } else if (action === "reject") {
    sheet.getRange(idx + 1, headers.indexOf("Status") + 1).setValue("Rejected");
    sheet.getRange(idx + 1, headers.indexOf("RejectionReason") + 1).setValue(reason);
  } else {
    throw new Error("Invalid action");
  }

  return true;
}

function deleteInvoiceById(invoiceId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Invoices");
  const data = sheet.getDataRange().getValues();
  const idIndex = data[0].indexOf("InvoiceID");

  const rowIndex = data.findIndex((row, i) => i > 0 && row[idIndex] === invoiceId);
  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex + 1);
    return true;
  }
  return false;
}
