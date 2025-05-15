/**
 * @OnlyCurrentDoc
 *
 * This script sets up the necessary sheets and columns for the Employee Control Centre spreadsheet.
 * It creates sheets if they don't exist, adds headers, and applies basic data validation.
 */

/**
 * Helper function to create or get a sheet, set its headers, and apply basic formatting.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} sheetName The name of the sheet.
 * @param {Array<string>} headers An array of header strings.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The created or existing sheet.
 */
function createSheetWithHeaders(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created.`);
  } else {
    Logger.log(`Sheet "${sheetName}" already exists.`);
  }

  if (headers && headers.length > 0) {
    let headersMatch = false;
    if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
      const currentHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (currentHeaderValues && Array.isArray(currentHeaderValues)) {
        headersMatch = currentHeaderValues.length === headers.length &&
                       currentHeaderValues.every((value, index) => value === headers[index]);
      }
    }

    if (!headersMatch) {
      const colsToClear = sheet.getLastColumn() > 0 ? Math.max(sheet.getLastColumn(), headers.length) : headers.length;
      if (colsToClear > 0) {
        sheet.getRange(1, 1, 1, colsToClear).clearContent();
      }
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold")
        .setBackground("#f0f0f0");
      Logger.log(`Headers set/updated for sheet "${sheetName}".`);
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
    } else {
      Logger.log(`Headers for sheet "${sheetName}" are already correct.`);
    }
  }
  // Freeze header row
  if (sheet.getMaxRows() > 1) { // Check to ensure sheet is not minimal (e.g. just created and empty)
      try {
          sheet.setFrozenRows(1);
      } catch(e) {
          Logger.log(`Could not freeze rows for sheet ${sheet.getName()}: ${e.message}. This can happen if the sheet is completely empty or too small.`);
      }
  }
  return sheet;
}

/**
 * Main function to set up the entire spreadsheet structure.
 * Run this function once from the Google Apps Script editor.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Configuration for _ValidationLists sheet (processed first) ---
  const validationListSheetDef = {
    name: "_ValidationLists",
    headers: [
      "RoleOptions", "DepartmentOptions", "AccountStatusOptions", "ShiftStatusOptions",
      "InvoiceTypeOptions", "InvoiceStatusOptions", "HolidayStatusOptions", "DayOfWeekOptions"
    ],
    data: [
      ["Employee", "Manager", "CFO", "Administrator"],
      ["SIC", "Performance", "Operations", "Creative", "B2AFC", "N/A"],
      ["Active", "Disabled"],
      ["Offered", "Accepted", "Completed", "Cancelled"],
      ["ShiftWork", "CPD", "Expense", "Course"],
      ["Draft", "Submitted", "ManagerApproved", "CFOApproved", "Rejected", "Paid", "CarriedOver"],
      ["PendingManagerApproval", "PendingCFOApproval", "Approved", "Rejected", "Cancelled"],
      ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    ],
    namedRanges: [
      { name: "RoleList", column: 1 }, { name: "DepartmentList", column: 2 },
      { name: "AccountStatusList", column: 3 }, { name: "ShiftStatusList", column: 4 },
      { name: "InvoiceTypeList", column: 5 }, { name: "InvoiceStatusList", column: 6 },
      { name: "HolidayStatusList", column: 7 }, { name: "DayOfWeekList", column: 8 }
    ]
  };

  // --- Process _ValidationLists sheet FIRST ---
  const validationSheet = createSheetWithHeaders(ss, validationListSheetDef.name, validationListSheetDef.headers);
  if (validationListSheetDef.data) {
    const numRows = Math.max(...validationListSheetDef.data.map(col => col.length));
    const transposedData = [];
    for (let r = 0; r < numRows; r++) {
      transposedData[r] = [];
      for (let c = 0; c < validationListSheetDef.data.length; c++) {
        transposedData[r][c] = validationListSheetDef.data[c][r] || "";
      }
    }
    if (transposedData.length > 0) {
      validationSheet.getRange(2, 1, transposedData.length, transposedData[0].length).setValues(transposedData);
      Logger.log(`Data populated for sheet "${validationListSheetDef.name}".`);
    }
  }
  if (validationListSheetDef.namedRanges) {
    validationListSheetDef.namedRanges.forEach(nr => {
      const columnValues = validationSheet.getRange(1, nr.column, validationSheet.getMaxRows()).getValues();
      let lastDataRowInColumn = 0;
      for (let i = columnValues.length - 1; i >= 1; i--) {
        if (columnValues[i][0] !== "") {
          lastDataRowInColumn = i + 1;
          break;
        }
      }
      if (lastDataRowInColumn > 1) {
        const range = validationSheet.getRange(2, nr.column, lastDataRowInColumn - 1, 1);
        try {
          const existingNamedRange = ss.getRangeByName(nr.name);
          if (existingNamedRange) ss.removeNamedRange(nr.name);
          ss.setNamedRange(nr.name, range);
          Logger.log(`Named range "${nr.name}" created/updated for sheet "${validationListSheetDef.name}".`);
        } catch (e) {
          Logger.log(`Error creating/updating named range "${nr.name}": ${e.message}.`);
        }
      } else {
        Logger.log(`Skipping named range "${nr.name}" for sheet "${validationListSheetDef.name}" as no data in column ${nr.column}.`);
      }
    });
  }

  // --- Configuration for other Sheets ---
  const sheetsData = [
    {
      name: "Users",
      headers: [
        "UserID", "Email", "FirstName", "LastName", "DateOfBirth", "MobileNumber",
        "Role", "Department", "HourlyRate", "HolidayEntitlementAccruedHours",
        "EmergencyContactName", "EmergencyContactPhone", "HealthConditions", "Medication",
        "BankDetailsConfirmation", "DBSDetails", "FirstAidQualifications", "AccountStatus",
        "LastLogin", "CreatedAt", "UpdatedAt"
      ],
      dropdowns: [
        { column: 7, rangeName: "RoleList" }, { column: 8, rangeName: "DepartmentList" },
        { column: 18, rangeName: "AccountStatusList" }
      ]
    },
    { name: "Qualifications", headers: ["QualificationID", "UserID", "CourseName", "Provider", "DateCompleted", "ExpiryDate", "CertificateFileID", "CreatedAt", "UpdatedAt"] },
    {
      name: "Shifts",
      headers: ["ShiftID", "Department", "ShiftDate", "StartTime", "EndTime", "ActualHoursWorked", "Description", "Status", "AssignedUserID", "AcceptedTimestamp", "CompletedTimestamp", "IsInvoiceDraftGenerated", "CreatedByUserID", "CreatedAt", "UpdatedAt"],
      dropdowns: [{ column: 2, rangeName: "DepartmentList" }, { column: 8, rangeName: "ShiftStatusList" }]
    },
    {
      name: "Invoices",
      headers: ["InvoiceID", "UserID", "InvoiceDate", "InvoiceType", "TotalAmount", "Status", "RelatedShiftIDs", "DescriptionOrPurpose", "ExpenseReceiptFileID", "SubmittedToManagerID", "ManagerApprovalTimestamp", "CFOApprovalTimestamp", "PaymentDate", "PaymentMonthCarryOver", "RejectionReason", "CreatedAt", "UpdatedAt"],
      dropdowns: [{ column: 4, rangeName: "InvoiceTypeList" }, { column: 6, rangeName: "InvoiceStatusList" }]
    },
    { name: "InvoiceItems", headers: ["InvoiceItemID", "InvoiceID", "Description", "Quantity", "UnitPrice", "LineTotal"] },
    {
      name: "HolidayRequests",
      headers: ["HolidayRequestID", "UserID", "RequestDate", "StartDate", "EndDate", "NumberOfDays", "AccruedHoursUsed", "Status", "ManagerApprovalTimestamp", "CFOApprovalTimestamp", "RejectionReason", "CreatedAt", "UpdatedAt"],
      dropdowns: [{ column: 8, rangeName: "HolidayStatusList" }]
    },
    {
      name: "Availability",
      headers: ["AvailabilityID", "UserID", "DayOfWeek", "StartTime", "EndTime", "IsAvailable", "Notes"],
      dropdowns: [{ column: 3, rangeName: "DayOfWeekList" }]
    },
    { name: "SystemSettings", headers: ["SettingName", "SettingValue", "Description"] }
  ];

  // --- Process other sheets ---
  sheetsData.forEach(sheetDef => {
    const sheet = createSheetWithHeaders(ss, sheetDef.name, sheetDef.headers);

    if (sheetDef.dropdowns) {
      sheetDef.dropdowns.forEach(dd => {
        const namedRange = ss.getRangeByName(dd.rangeName);
        if (namedRange) {
          const rule = SpreadsheetApp.newDataValidation().requireValueInRange(namedRange, true).setAllowInvalid(false).build();
          sheet.getRange(2, dd.column, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
          Logger.log(`Data validation set for column ${dd.column} in sheet "${sheetDef.name}" using named range "${dd.rangeName}".`);
        } else {
          Logger.log(`Could not find named range "${dd.rangeName}" for data validation in sheet "${sheetDef.name}".`);
        }
      });
    }
  });

  // Optional: Move _ValidationLists to the end or hide it
  // if (validationSheet) {
  //   ss.setActiveSheet(validationSheet);
  //   ss.moveActiveSheet(ss.getNumSheets()); // Moves it to the end
  //   validationSheet.hideSheet();
  //   Logger.log(`Sheet "_ValidationLists" moved to end and hidden.`);
  // }

  SpreadsheetApp.flush();
  Logger.log("Spreadsheet setup complete.");
  SpreadsheetApp.getUi().alert("Spreadsheet Setup Complete!", "All sheets, headers, and basic data validations have been configured.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Helper function to clear all existing named ranges in the spreadsheet.
 */
function clearAllNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  namedRanges.forEach(namedRange => {
    namedRange.remove();
    Logger.log(`Removed named range: ${namedRange.getName()}`);
  });
  Logger.log('All named ranges cleared.');
}
