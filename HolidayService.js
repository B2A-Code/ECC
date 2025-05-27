/**
*
* Holiday Approvals
* 
*/

function getTeamHolidayCalendar() {
  Logger.log("🚀 getTeamHolidayCalendar started");

  // 1. User & auth
  const email = Session.getActiveUser().getEmail();
  Logger.log(`🔐 Session user email: ${email}`);

  const user = getUserByEmail(email);
  if (!user) {
    Logger.log("❌ No user found for email, exiting.");
    return [];
  }
  Logger.log(`👤 Authenticated user: ${user.FirstName} ${user.LastName} (Role=${user.Role}, Dept=${user.Department})`);

  if (user.Role !== "Manager") {
    Logger.log(`🚫 Access denied: not a Manager (Role=${user.Role})`);
    return [];
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName("HolidayRequests");
  const userSheet = ss.getSheetByName("Users");
  if (!sheet || !userSheet) {
    Logger.log(`❌ Missing sheet(s): HolidayRequests=${!!sheet}, Users=${!!userSheet}`);
    return [];
  }
  Logger.log("✅ Both sheets found");

  const requestData    = sheet.getDataRange().getValues();
  const requestHeaders = requestData.shift();
  Logger.log(`📄 HolidayRequests: fetched ${requestData.length} rows; headers=${JSON.stringify(requestHeaders)}`);

  const userData       = userSheet.getDataRange().getValues();
  const userHeaders    = userData.shift();
  Logger.log(`👥 Users: fetched ${userData.length} rows; headers=${JSON.stringify(userHeaders)}`);

  const userMap = userData.reduce((m, row, idx) => {
    const u = userHeaders.reduce((o, h, i) => {
      o[h] = row[i];
      return o;
    }, {});
    m[u.UserID] = u;
    Logger.log(`   🔹 userMap[${u.UserID}] = ${u.FirstName} ${u.LastName}, Dept=${u.Department}`);
    return m;
  }, {});
  Logger.log(`📌 Built userMap with ${Object.keys(userMap).length} entries`);

  const toISO = val => {
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? "" : d.toISOString();
    } catch (e) {
      Logger.log(`⚠️ toISO() failed for value=${val}: ${e}`);
      return "";
    }
  };

  const result = [];
  requestData.forEach((row, i) => {
    const rec = requestHeaders.reduce((o, h, j) => {
      o[h] = row[j];
      return o;
    }, {});
    Logger.log(`\n🔍 Processing row ${i+1}: HolidayRequestID=${rec.HolidayRequestID}, UserID=${rec.UserID}`);

    const reqUser = userMap[rec.UserID];
    if (!reqUser) {
      Logger.log(`   ❗ No matching user in map for UserID=${rec.UserID}`);
      return;
    }
    Logger.log(`   👤 Matches user ${reqUser.FirstName} ${reqUser.LastName} (Dept=${reqUser.Department})`);

    if (reqUser.Department !== user.Department) {
      Logger.log(`   🚫 Skipping due to department mismatch (${reqUser.Department} ≠ ${user.Department})`);
      return;
    }

    const norm = {
      holidayRequestId: rec.HolidayRequestID,
      avatar:           reqUser.Avatar      || "",
      fullName:         reqUser.FirstName + " " + reqUser.LastName,
      email:            reqUser.Email,
      start:            toISO(rec.StartDate),
      end:              toISO(rec.EndDate),
      numberOfDays:     rec.NumberOfDays,
      reason:           rec.Reason          || "",
      status:           rec.Status          || "",
      createdAt:        toISO(rec.CreatedAt),
      updatedAt:        toISO(rec.UpdatedAt),
      requestDate:      toISO(rec.RequestDate),
      managerApproved:  toISO(rec.ManagerApprovalTimestamp),
      cfoApproved:      toISO(rec.CFOApprovalTimestamp)
    };
    Logger.log(`   ✅ Normalized record: ${JSON.stringify(norm)}`);

    result.push(norm);
  });

  Logger.log(`🎯 Finished processing. Returning ${result.length} records`);
  return result;
}

function rejectHolidayEventInCalendar(eventId, rejectionReason) {
  Logger.log("📥 [rejectHolidayEventInCalendar] Function called");
  Logger.log(`🔑 Received eventId: ${eventId}`);
  Logger.log(`📄 Rejection reason: ${rejectionReason}`);

  try {
    if (!eventId || eventId.trim() === "") {
      throw new Error("❌ No event ID provided");
    }

    // Normalize event ID if it doesn't include domain
    if (!eventId.includes('@')) {
      Logger.log("⚠️ Event ID missing domain, appending '@google.com'...");
      eventId += '@google.com';
      Logger.log(`ℹ️ Event ID modified to: ${eventId}`);
    }

    const calendarId = 'your_shared_calendar_id@group.calendar.google.com';
    Logger.log(`📅 Using calendar ID: ${calendarId}`);

    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error(`❌ Calendar not found for ID: ${calendarId}`);
    }
    Logger.log("✅ Calendar retrieved successfully");

    const event = calendar.getEventById(eventId);
    if (!event) {
      throw new Error(`❌ Event not found for ID: ${eventId}`);
    }
    Logger.log(`✅ Event found. Title: "${event.getTitle()}"`);

    const oldDescription = event.getDescription();
    Logger.log(`📝 Old description:\n${oldDescription}`);

    const updatedDescription = `${oldDescription}\n\n[REJECTED] ${rejectionReason}`;
    event.setDescription(updatedDescription);
    Logger.log("✅ Event description updated with rejection note");

    return { success: true };

  } catch (err) {
    Logger.log(`❌ Error in rejectHolidayEventInCalendar: ${err.message}`);
    return { success: false, error: err.message };
  }
}

function getPendingHolidayRequestsForManager() {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || user.Role !== "Manager") {
    Logger.log("❌ Not a valid Manager user or user not found.");
    return [];
  }

  const requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");

  const requestRows = requestsSheet.getDataRange().getValues();
  const requestHeaders = requestRows.shift();
  const userRows = usersSheet.getDataRange().getValues();
  const userHeaders = userRows.shift();

  const userMap = {};
  for (const row of userRows) {
    const userObj = userHeaders.reduce((acc, key, i) => {
      acc[key] = row[i];
      return acc;
    }, {});
    userMap[userObj.UserID] = userObj;
  }

  const idIndex = requestHeaders.indexOf("HolidayRequestID");
  const userIdIndex = requestHeaders.indexOf("UserID");
  const startDateIndex = requestHeaders.indexOf("StartDate");
  const endDateIndex = requestHeaders.indexOf("EndDate");
  const statusIndex = requestHeaders.indexOf("Status");
  const calendarEventIdIndex = requestHeaders.indexOf("CalendarEventID");

  const missingColumns = [idIndex, userIdIndex, startDateIndex, endDateIndex, statusIndex, calendarEventIdIndex]
    .some(idx => idx === -1);

  if (missingColumns) {
    Logger.log("❌ One or more required columns are missing in the HolidayRequests sheet.");
    return [];
  }

  const pendingRequests = [];

  for (const row of requestRows) {
    const status = row[statusIndex];
    const userId = row[userIdIndex];
    const requestingUser = userMap[userId];

    if (!requestingUser) {
      Logger.log(`⚠️ No user found for UserID: ${userId}`);
      continue;
    }

    if (
      (status !== HOLIDAY_STATUSES.PENDING_MANAGER && status !== HOLIDAY_STATUSES.PENDING_CFO) ||
      requestingUser.Department !== user.Department
    ) {
      continue;
    }

    pendingRequests.push({
      holidayRequestId: row[idIndex],
      userId,
      fullName: `${requestingUser.FirstName} ${requestingUser.LastName}`,
      start: row[startDateIndex],
      end: row[endDateIndex],
      status,
      calendarEventId: row[calendarEventIdIndex],
      avatar: requestingUser.Avatar || null,
      email: requestingUser.Email || null
    });
  }

  Logger.log(`📦 Returning ${pendingRequests.length} pending requests for manager ${user.Email}`);
  return pendingRequests;
}

function getMyAvailability() {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user) return [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availability");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const result = [];

  const statusIndex = headers.indexOf("Status");
  const userIdIndex = headers.indexOf("UserID");

  for (const row of data) {
    const record = headers.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {});

    if (record.UserID === user.UserID && record.Status !== 'Cancelled' && record.Status !== 'Rejected') {
      result.push(record);
    }
  }

  return result;
}

function approveOrRejectHoliday(requestId, action, reason = "") {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || (user.Role !== "Manager" && user.Role !== "CFO")) {
    Logger.log("❌ Not authorized");
    return { success: false, error: "Not authorized" };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const now = new Date();
  const idIndex = headers.indexOf("HolidayRequestID");
  const statusIndex = headers.indexOf("Status");
  const managerTimestamp = headers.indexOf("ManagerApprovalTimestamp");
  const cfoTimestamp = headers.indexOf("CFOApprovalTimestamp");
  const rejectionReasonIndex = headers.indexOf("RejectionReason");
  const userIdIndex = headers.indexOf("UserID");
  const hoursUsedIndex = headers.indexOf("AccruedHoursUsed");
  const startDateIndex = headers.indexOf("StartDate");
  const endDateIndex = headers.indexOf("EndDate");
  const calendarEventIdIndex = headers.indexOf("CalendarEventID");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIndex] === requestId) {
      const currentStatus = row[statusIndex];
      Logger.log(`🔎 Found request: ${requestId} with current status: ${currentStatus}`);

      const requestUser = getUserById(row[userIdIndex]);
      if (!requestUser) {
        Logger.log("❌ User not found for this request");
        return { success: false, error: "User not found" };
      }

      Logger.log("📋 Row BEFORE update: " + JSON.stringify(row));
      Logger.log("🧾 Row with headers BEFORE update:\n" + JSON.stringify(headers.reduce((obj, key, idx) => {
        obj[key] = row[idx];
        return obj;
      }, {}), null, 2));

      if (action === "approve") {
        const nextStatus = getNextHolidayStatus(currentStatus, user.Role, requestUser.Role);
        Logger.log(`📌 Next status: ${nextStatus}`);

        if (!nextStatus) {
          Logger.log("⛔ No next status returned from getNextHolidayStatus");
          return { success: false, error: "Approval not allowed at this stage" };
        }

        row[statusIndex] = nextStatus;

        if (user.Role === "Manager") {
          row[managerTimestamp] = now.toISOString();

          if (nextStatus === "PendingCFOApproval") {
            const cfo = getAllUsers().find(u => u.Role === "CFO");
            if (cfo?.Email) {
              MailApp.sendEmail(
                cfo.Email,
                "🔔 Holiday Request Awaiting CFO Approval",
                `Request ${requestId} from ${requestUser.FirstName} ${requestUser.LastName} has been approved by the Manager and awaits your review.`
              );
            }
          }
        }

        if (user.Role === "CFO") {
          row[cfoTimestamp] = now.toISOString();
          const hoursUsed = parseFloat(row[hoursUsedIndex]);
          deductHolidayHours(row[userIdIndex], hoursUsed);

          // ✅ Use centralised calendar function
          const event = addHolidayToCalendar(
            requestUser,
            row[startDateIndex],
            row[endDateIndex],
            `Holiday: ${requestUser.FirstName} ${requestUser.LastName}`
          );

          if (event) {
            row[calendarEventIdIndex] = event.getId();
          }
        }

        MailApp.sendEmail(requestUser.Email, "✅ Holiday Approved", `Your holiday has been approved (${nextStatus}).`);
      }

      else if (action === "reject") {
        if (!canRejectHoliday(currentStatus, user.Role)) {
          return { success: false, error: "Rejection not allowed at this stage" };
        }

        row[statusIndex] = HOLIDAY_STATUSES.REJECTED;
        row[rejectionReasonIndex] = reason;

        if (user.Role === "Manager") row[managerTimestamp] = now.toISOString();
        if (user.Role === "CFO") row[cfoTimestamp] = now.toISOString();

        const eventId = row[calendarEventIdIndex];
        if (eventId) {
          try {
            const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
            const event = calendar.getEventById(eventId);
            if (event) event.deleteEvent();
            row[calendarEventIdIndex] = ""; // Clear the event ID
          } catch (e) {
            Logger.log("❌ Failed to delete calendar event: " + e.message);
          }
        }

        MailApp.sendEmail(requestUser.Email, "❌ Holiday Rejected", `Reason: ${reason}`);
      }

      Logger.log("📋 Row AFTER update: " + JSON.stringify(row));
      Logger.log("🧾 Row with headers AFTER update:\n" + JSON.stringify(headers.reduce((obj, key, idx) => {
        obj[key] = row[idx];
        return obj;
      }, {}), null, 2));

      Logger.log(`💾 Writing updated row back to sheet at row ${i + 1}`);
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }

  Logger.log("❌ No request found with ID: " + requestId);
  return { success: false, error: "Request not found" };
}

/**
*
* My Holidays
* 
*/

function getMyHolidayRequests() {
  Logger.log("🚀 getMyHolidayRequests started");

  const email = Session.getActiveUser().getEmail();
  Logger.log(`🔐 Session email: ${email}`);

  const user = getUserByEmail(email);
  if (!user) {
    Logger.log("❌ No user found");
    return {
      success: false,
      error: "User not found",
      data: [],
      daysTaken: 0,
      daysPending: 0,
      daysRemaining: 28
    };
  }

  Logger.log(`👤 Authenticated: ${user.FirstName} ${user.LastName} (UserID=${user.UserID})`);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  if (!sheet) {
    Logger.log("❌ HolidayRequests sheet missing");
    return {
      success: false,
      error: "HolidayRequests sheet not found",
      data: [],
      daysTaken: 0,
      daysPending: 0,
      daysRemaining: 28
    };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  Logger.log(`📄 Loaded ${data.length} rows`);

  const toISO = val => {
    try {
      const d = new Date(val);
      return isNaN(d.getTime()) ? "" : d.toISOString();
    } catch (e) {
      return "";
    }
  };

  const requests = [];
  let totalApprovedDays = 0;
  let totalPendingDays = 0; // 🆕

  data.forEach((row) => {
    const rec = headers.reduce((o, h, j) => {
      o[h] = row[j];
      return o;
    }, {});

    if (rec.UserID !== user.UserID) return;

    const normalized = {
      holidayRequestId: rec.HolidayRequestID,
      avatar:           user.Avatar || "",
      fullName:         `${user.FirstName} ${user.LastName}`,
      email:            user.Email,
      start:            toISO(rec.StartDate),
      end:              toISO(rec.EndDate),
      numberOfDays:     Number(rec.NumberOfDays || 0),
      reason:           rec.RejectionReason || "",
      status:           rec.Status || "",
      createdAt:        toISO(rec.CreatedAt),
      updatedAt:        toISO(rec.UpdatedAt),
      requestDate:      toISO(rec.RequestDate),
      managerApproved:  toISO(rec.ManagerApprovalTimestamp),
      cfoApproved:      toISO(rec.CFOApprovalTimestamp)
    };

    requests.push(normalized);

    const status = (normalized.status || "").toLowerCase();
    if (status === "approved") {
      totalApprovedDays += normalized.numberOfDays;
    } else if (
      status === "pendingmanagerapproval" ||
      status === "pendingcfoapproval"
    ) {
      totalPendingDays += normalized.numberOfDays;
    }
  });

  const HOLIDAY_ENTITLEMENT = Number(user.HolidayEntitlementAccruedHours || 28);
  const daysRemaining = HOLIDAY_ENTITLEMENT - totalApprovedDays;

  Logger.log(`✔️ Days Taken: ${totalApprovedDays}`);
  Logger.log(`🕒 Days Pending: ${totalPendingDays}`);
  Logger.log(`✔️ Days Remaining: ${daysRemaining}`);

  return {
    success: true,
    data: requests,
    daysTaken: totalApprovedDays,
    daysPending: totalPendingDays, // 🆕 return pending days
    daysRemaining
  };
}

function getHolidayEntitlementSummary() {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  if (!user || (!user.Permanent && user.Role !== "Manager")) return null;

  const totalEntitlement = 28; // UK default
  const usedHours = parseFloat(user.HolidayEntitlementAccruedHours) || 0;
  const daysTaken = usedHours / 7;
  const daysLeft = totalEntitlement - daysTaken;

  return {
    totalEntitlement,
    daysTaken: parseFloat(daysTaken.toFixed(1)),
    daysLeft: parseFloat(daysLeft.toFixed(1))
  };
}

function submitHolidayRequest(payload) {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  if (!user) {
    return { success: false, error: "User not found." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  if (!sheet) {
    return { success: false, error: "HolidayRequests sheet not found." };
  }

  const headers = sheet.getDataRange().getValues()[0];
  const nextId = generateNextId("HREQ_", sheet);

  const startDate = new Date(payload.StartDate);
  const endDate = new Date(payload.EndDate);
  const numberOfDays = calculateWorkingDays(startDate, endDate);

  const now = new Date();

  const row = headers.map(h => {
    switch (h) {
      case "HolidayRequestID": return nextId;
      case "UserID": return user.UserID;
      case "RequestDate": return now.toISOString();
      case "StartDate": return now.toISOString();
      case "EndDate": return now.toISOString();
      case "NumberOfDays": return numberOfDays;
      case "AccruedHoursUsed": return "";
      case "Status": return "Pending";
      case "ManagerApprovalTimestamp":
      case "CFOApprovalTimestamp":
      case "RejectionReason":
      case "UpdatedAt": return "";
      case "CreatedAt": return now.toISOString();
      default: return "";
    }
  });

  sheet.appendRow(row);

  return { success: true, message: "Holiday request submitted." };
}

function submitHolidayRequestViaGAS(payload) {
  try {
    Logger.log("📥 [submitHolidayRequestViaGAS] Called with payload: " + JSON.stringify(payload));

    const email = Session.getActiveUser().getEmail();
    Logger.log(`📧 Authenticated user: ${email}`);

    const user = getUserByEmail(email);
    if (!user || (!user.Permanent && user.Role !== "Manager")) {
      Logger.log("❌ Not authorized to submit request");
      return { success: false, error: "Not authorized" };
    }

    Logger.log(`👤 User found: ${user.FirstName} ${user.LastName} (${user.Role})`);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
    if (!sheet) {
      Logger.log("❌ HolidayRequests sheet not found");
      return { success: false, error: "Sheet not found" };
    }

    const now = new Date();
    const newId = "HR" + (sheet.getLastRow() + 1).toString().padStart(4, "0");
    Logger.log(`🆔 Generated HolidayRequestID: ${newId}`);

    const { StartDate, EndDate, StartType, EndType, Reason } = payload;
    const start = new Date(StartDate);
    const end = new Date(EndDate);
    const startISO = start.toISOString();
    const endISO = end.toISOString();

    let workingDays = calculateWorkingDays(start, end);
    if (StartType === "Half") workingDays -= 0.5;
    if (EndType === "Half") workingDays -= 0.5;

    Logger.log(`📅 Dates: ${startISO} → ${endISO}`);
    Logger.log(`📊 Working Days: ${workingDays}`);

    const totalEntitled = parseFloat(user.HolidayEntitlementAccruedHours || 0);
    const usedDays = getUsedHolidayDays(user.UserID);
    const availableDays = Math.max(0, totalEntitled - usedDays);

    Logger.log(`💼 Entitled: ${totalEntitled} | Used: ${usedDays} | Available: ${availableDays}`);

    if (workingDays > availableDays) {
      const error = `You only have ${availableDays.toFixed(1)} days left. Requested: ${workingDays.toFixed(1)}`;
      Logger.log("❌ " + error);
      return { success: false, error };
    }

    const initialStatus = user.Role === "Employee"
      ? HOLIDAY_STATUSES.PENDING_MANAGER
      : HOLIDAY_STATUSES.PENDING_CFO;

    Logger.log(`🧭 Initial status based on role (${user.Role}): ${initialStatus}`);

    const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
    if (!calendar) {
      Logger.log("❌ Holiday calendar not found");
      throw new Error("Holiday calendar not found");
    }

    Logger.log("📆 Creating calendar event...");
    const calendarEvent = calendar.createAllDayEvent(
      `Holiday: ${user.FirstName} ${user.LastName}`,
      start,
      end,
      {
        description: `Pending holiday for ${user.FirstName} ${user.LastName}\nReason: ${Reason || "Annual Leave"}`,
        extendedProperties: {
          private: {
            userID: user.UserID,
            userEmail: user.Email,
            type: "holiday",
            numberOfDays: workingDays.toString(),
            reason: Reason || "Annual Leave",
            status: initialStatus
          }
        }
      }
    );

    const eventId = calendarEvent.getId();
    Logger.log(`✅ Calendar event created. Event ID: ${eventId}`);

    const headers = sheet.getDataRange().getValues()[0];
    const row = headers.map(header => {
      switch (header) {
        case "HolidayRequestID": return newId;
        case "UserID": return user.UserID;
        case "RequestDate": return now.toISOString();
        case "StartDate": return startISO;
        case "EndDate": return endISO;
        case "NumberOfDays": return workingDays;
        case "AccruedHoursUsed": return workingDays * 7;
        case "Status": return initialStatus;
        case "ManagerApprovalTimestamp":
        case "CFOApprovalTimestamp":
        case "RejectionReason": return "";
        case "CreatedAt":
        case "UpdatedAt": return now.toISOString();
        case "CalendarEventID": return eventId;
        case "CalendarId": return STAFF_HOLIDAY_CALENDAR_ID;
        default: return "";
      }
    });

    Logger.log("📋 Row to append:\n" + JSON.stringify(row, null, 2));
    sheet.appendRow(row);
    Logger.log("💾 Row appended to sheet");

    const managerEmail = getManagerEmail(user.Department);
    Logger.log(`📨 Sending notification to manager: ${managerEmail}`);

    MailApp.sendEmail(
      managerEmail,
      "🆕 Holiday Request Submitted",
      `${user.FirstName} ${user.LastName} has submitted a holiday request.\n\nFrom: ${startISO}\nTo: ${endISO}\nDays: ${workingDays.toFixed(1)}\n\nPlease review in the system.`
    );

    Logger.log("✅ Notification sent. Request complete.");
    return { success: true, daysRequested: workingDays };

  } catch (err) {
    Logger.log("❌ Error submitting holiday request: " + err.message);
    return { success: false, error: err.message };
  }
}

function removeCalendarEventsForUser(userEmail, calendarId, startDate, endDate) {
  const cal = CalendarApp.getCalendarById(calendarId);
  const events = cal.getEvents(new Date(startDate), new Date(endDate));
  
  for (const ev of events) {
    if (ev.getTag("userEmail") === userEmail) {
      ev.deleteEvent();
    }
  }
}



/**
*
* My Avilability
* 
*/

function getGoogleHolidayEventsFromGAS() {
  try {
    const now = new Date().toISOString();
    const calendarId = STAFF_HOLIDAY_CALENDAR_ID;

    const events = Calendar.Events.list(calendarId, {
      timeMin: now,
      showDeleted: false,
      singleEvents: true,
      maxResults: 100,
      orderBy: 'startTime',
      q: 'Holiday'
    });

    return events;
  } catch (err) {
    return {
      error: true,
      message: err.message || "Unknown error"
    };
  }
}

function addAvailabilityToCalendar(user, startDate, endDate, reason) {
  const calendar = CalendarApp.getCalendarById(CONTRACTOR_AVAILABILITY_CALENDAR_ID);
  if (!calendar) throw new Error("Contractor availability calendar not found.");

  calendar.createAllDayEvent(reason || "Unavailable", new Date(startDate), new Date(endDate), {
    description: `Availability block for ${user.FirstName} ${user.LastName}`,
    extendedProperties: {
      private: {
        userID: user.UserID,
        userEmail: user.Email,
        type: "availability"
      }
    }
  });
}

function submitHolidayRequestViaGAS(payload) {
  try {
    const email = Session.getActiveUser().getEmail();
    const user = getUserByEmail(email);
    if (!user || (!user.Permanent && user.Role !== "Manager")) {
      return { success: false, error: "Not authorized" };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
    const now = new Date();
    const newId = "HR" + (sheet.getLastRow() + 1).toString().padStart(4, "0");

    const { StartDate, EndDate, StartType, EndType, Reason } = payload;
    const start = new Date(StartDate);
    const end = new Date(EndDate);
    const startISO = start.toISOString();
    const endISO = end.toISOString();

    let workingDays = calculateWorkingDays(start, end);
    if (StartType === "Half") workingDays -= 0.5;
    if (EndType === "Half") workingDays -= 0.5;

    const availableDays = Math.max(0, parseFloat(user.HolidayEntitlementAccruedHours || 0) / 7);
    if (workingDays > availableDays) {
      return {
        success: false,
        error: `You only have ${availableDays.toFixed(1)} days left. Requested: ${workingDays.toFixed(1)}`
      };
    }

    // Determine initial status dynamically from role
    const initialStatus = getInitialHolidayStatusByRole(user.Role);
    Logger.log(`🧭 Initial status based on role (${user.Role}): ${initialStatus}`);

    // 📅 Create placeholder event
    const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
    if (!calendar) throw new Error("Holiday calendar not found");

    const calendarEvent = calendar.createAllDayEvent(
      `Holiday: ${user.FirstName} ${user.LastName}`,
      start,
      end,
      {
        description: `Pending holiday for ${user.FirstName} ${user.LastName}\nReason: ${Reason || "Annual Leave"}`,
        extendedProperties: {
          private: {
            userID: user.UserID,
            userEmail: user.Email,
            type: "holiday",
            numberOfDays: workingDays.toString(),
            reason: Reason || "Annual Leave",
            status: initialStatus
          }
        }
      }
    );

    const eventId = calendarEvent.getId();

    // ✅ Append the request to the sheet
    sheet.appendRow([
      newId,
      user.UserID,
      now.toISOString(),
      startISO,
      endISO,
      workingDays,
      workingDays * 7,
      initialStatus,
      "",
      "",
      "", // Rejection reason
      now.toISOString(),
      now.toISOString(),
      eventId, // CalendarEventID
      STAFF_HOLIDAY_CALENDAR_ID // Calendar ID (N2:N)
    ]);

    // 📧 Notify Manager (always, even for CFO to keep record)
    const managerEmail = getManagerEmail(user.Department);
    if (managerEmail) {
      MailApp.sendEmail(
        managerEmail,
        "🆕 Holiday Request Submitted",
        `${user.FirstName} ${user.LastName} has submitted a holiday request.\n\nFrom: ${startISO}\nTo: ${endISO}\nDays: ${workingDays.toFixed(1)}\n\nPlease review in the system.`
      );
    }

    return { success: true, daysRequested: workingDays };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function cancelAvailabilityRequest(availabilityId) {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availability");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf("AvailabilityID");
  const userIndex = headers.indexOf("UserID");
  const statusIndex = headers.indexOf("Status");
  const calendarEventIdIndex = headers.indexOf("CalendarEventID");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIndex] === availabilityId && row[userIndex] === user.UserID) {
      row[statusIndex] = "Cancelled";

      const eventId = row[calendarEventIdIndex];
      if (eventId) {
        try {
          const calendar = CalendarApp.getCalendarById(CONTRACTOR_AVAILABILITY_CALENDAR_ID);
          const event = calendar.getEventById(eventId);
          if (event) event.deleteEvent();
        } catch (e) {
          Logger.log("Failed to delete availability event: " + e.message);
        }
      }

      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }

  return { success: false, error: "Not found or not authorized" };
}

function cancelHolidayRequest(requestId) {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  if (!user) return { success: false, error: "User not found" };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idIndex = headers.indexOf("HolidayRequestID");
  const userIndex = headers.indexOf("UserID");
  const statusIndex = headers.indexOf("Status");
  const calendarEventIdIndex = headers.indexOf("CalendarEventID");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIndex] === requestId && row[userIndex] === user.UserID) {
      const currentStatus = row[statusIndex];
      if (currentStatus === HOLIDAY_STATUSES.APPROVED || currentStatus === HOLIDAY_STATUSES.REJECTED) {
        return { success: false, error: "Cannot cancel after final decision" };
      }

      row[statusIndex] = HOLIDAY_STATUSES.CANCELLED;

      const eventId = row[calendarEventIdIndex];
      if (eventId) {
        try {
          const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
          const event = calendar.getEventById(eventId);
          if (event) event.deleteEvent();
        } catch (e) {
          Logger.log("Failed to delete calendar event: " + e.message);
        }
      }

      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      MailApp.sendEmail(user.Email, "⛔ Holiday Request Cancelled", `Your request ${requestId} has been cancelled.`);

      return { success: true };
    }
  }

  return { success: false, error: "Request not found or not owned by user" };
}

/**
*
* Utility
* 
*/

const STAFF_HOLIDAY_CALENDAR_ID = 'c_0f679b8ddbbc4c4fe84f4be0938b7a8170a6dd47667a7b2dd46675dc4a74523c@group.calendar.google.com';
const CONTRACTOR_AVAILABILITY_CALENDAR_ID = 'c_c4f71d0d3d33796a92c2dfe7ebc381e5bd6fb67bb57b06c00c433d4355f28b4e@group.calendar.google.com';

function getDepartmentByUserId(userId) {
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const usersData = usersSheet.getDataRange().getValues();
  const headers = usersData[0];
  const idIndex = headers.indexOf("UserID");
  const deptIndex = headers.indexOf("Department");

  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][idIndex] === userId) {
      return usersData[i][deptIndex];
    }
  }
  return null;
}

function addHolidayToCalendar(user, startDate, endDate, summary) {
  const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
  if (!calendar) throw new Error("Staff holiday calendar not found.");

  const department = getDepartmentByUserId(user.UserID);
  const colorMap = {
    "SIC": CalendarApp.EventColor.PALE_GREEN,
    "Performance": CalendarApp.EventColor.PALE_RED,
    "Operations": CalendarApp.EventColor.PALE_YELLOW,
    "Creative": CalendarApp.EventColor.PALE_BLUE,
    "B2AFC": CalendarApp.EventColor.PALE_ORANGE
  };

  const eventColor = colorMap[department] || CalendarApp.EventColor.GRAY;

  calendar.createAllDayEvent(summary, new Date(startDate), new Date(endDate), {
    description: `Holiday for ${user.FirstName} ${user.LastName} (${user.Email})`,
    color: eventColor,
    extendedProperties: {
      private: {
        userID: user.UserID,
        userEmail: user.Email,
        type: "holiday"
      }
    }
  });
}

function generateNextId(prefix, sheet) {
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const idIndex = header.indexOf("HolidayRequestID");
  let maxNum = 0;

  data.forEach(row => {
    const id = row[idIndex];
    if (typeof id === "string" && id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ""));
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  });

  return prefix + String(maxNum + 1).padStart(3, "0");
}

function calculateWorkingDays(startDate, endDate) {
  let count = 0;
  const cur = new Date(startDate);

  while (cur <= endDate) {
    const day = cur.getDay();
    if (day !== 0 && day !== 6) count++; // Mon-Fri only
    cur.setDate(cur.getDate() + 1);
  }

  return count;
}

function calculateWorkingDays(startDate, endDate) {
  let count = 0;
  const cur = new Date(startDate);

  while (cur <= endDate) {
    const day = cur.getDay();
    if (day !== 0 && day !== 6) count++; // Exclude weekends
    cur.setDate(cur.getDate() + 1);
  }

  return count;
}

function deductHolidayHours(userId, hoursToDeduct) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const header = data[0];

  const userIdx = data.findIndex(r => r[0] === userId);
  if (userIdx === -1) return;

  const hoursCol = header.indexOf("HolidayEntitlementAccruedHours");
  const current = parseFloat(data[userIdx][hoursCol] || 0);
  const updated = Math.max(0, current - hoursToDeduct);

  sheet.getRange(userIdx + 1, hoursCol + 1).setValue(updated);
}

function getManagerEmail(department) {
  const users = getAllUsers();
  const manager = users.find(u => u.Department === department && u.Role === "Manager");
  return manager?.Email || Session.getActiveUser().getEmail(); // fallback
}


const HOLIDAY_STATUSES = {
  PENDING_MANAGER: 'PendingManagerApproval',
  PENDING_CFO: 'PendingCFOApproval',
  APPROVED: 'Approved',
  REJECTED: 'Rejected',
  CANCELLED: 'Cancelled'
};

function getHolidayStatusLabel(status) {
  switch (status) {
    case HOLIDAY_STATUSES.PENDING_MANAGER: return 'Awaiting Manager Approval';
    case HOLIDAY_STATUSES.PENDING_CFO: return 'Awaiting CFO Approval';
    case HOLIDAY_STATUSES.APPROVED: return 'Approved';
    case HOLIDAY_STATUSES.REJECTED: return 'Rejected';
    case HOLIDAY_STATUSES.CANCELLED: return 'Cancelled by Employee';
    default: return 'Unknown';
  }
}

function getNextHolidayStatus(currentStatus, approverRole, requestUserRole) {
  if (currentStatus === "PendingManagerApproval") {
    if (approverRole === "Manager" && requestUserRole === "Employee") {
      return "Approved";
    }
    if (approverRole === "Manager" && requestUserRole === "Manager") {
      return "PendingCFOApproval";
    }
    if (approverRole === "CFO") {
      return "Approved";
    }
  }

  if (currentStatus === "PendingCFOApproval" && approverRole === "CFO") {
    return "Approved";
  }

  return null;
}

function canApproveHoliday(status, role) {
  return (
    (role === 'Manager' && status === HOLIDAY_STATUSES.PENDING_MANAGER) ||
    (role === 'CFO' && status === HOLIDAY_STATUSES.PENDING_CFO)
  );
}

function canRejectHoliday(status, role) {
  return canApproveHoliday(status, role);
}

function getHolidayStatusMeta(status) {
  const label = getHolidayStatusLabel(status);
  let color = "#999";

  switch (status) {
    case HOLIDAY_STATUSES.PENDING_MANAGER: color = "#f1c40f"; break;
    case HOLIDAY_STATUSES.PENDING_CFO: color = "#f39c12"; break;
    case HOLIDAY_STATUSES.APPROVED: color = "#2ecc71"; break;
    case HOLIDAY_STATUSES.REJECTED: color = "#e74c3c"; break;
    case HOLIDAY_STATUSES.CANCELLED: color = "#95a5a6"; break;
  }

  return { label, color };
}

function getDepartmentColor(department) {
  const pastelColors = {
    SIC: "#FFB3BA",         // light red
    Performance: "#BAE1FF", // light blue
    Operations: "#BFFCC6",  // light green
    Creative: "#FFFFBA",    // light yellow
    B2AFC: "#D5BAFF"        // light purple
  };
  return pastelColors[department] || "#E0E0E0"; // fallback grey
}

function getStatusMeta(status) {
  const map = {
    PendingManagerApproval: { label: "Awaiting Manager", color: "#f1c40f" },
    PendingCFOApproval: { label: "Awaiting CFO", color: "#f39c12" },
    Approved: { label: "Approved", color: "#2ecc71" },
    Rejected: { label: "Rejected", color: "#e74c3c" },
    Cancelled: { label: "Cancelled", color: "#95a5a6" }
  };
  return map[status] || { label: "Unknown", color: "#ccc" };
}

function getUserById(userId) {
  const users = getAllUsers();
  return users.find(u => u.UserID === userId);
}

function auditAvailabilityCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availability");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const calendarEventIdIndex = headers.indexOf("CalendarEventID");

  const calendar = CalendarApp.getCalendarById(CONTRACTOR_AVAILABILITY_CALENDAR_ID);
  const eventIdsInSheet = new Set();

  for (let i = 1; i < data.length; i++) {
    const eventId = data[i][calendarEventIdIndex];
    if (eventId) eventIdsInSheet.add(eventId);
  }

  const events = calendar.getEvents(new Date('2023-01-01'), new Date('2030-01-01'));

  for (const ev of events) {
    if (!eventIdsInSheet.has(ev.getId())) {
      Logger.log(`Deleting orphaned event: ${ev.getTitle()}`);
      ev.deleteEvent();
    }
  }
}

function auditHolidayCalendarDiscrepancies() {
  const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const calendarEventIdIndex = headers.indexOf("CalendarEventID");
  const holidayRequestIdIndex = headers.indexOf("HolidayRequestID");
  const userIdIndex = headers.indexOf("UserID");

  const eventIdsInSheet = new Set();
  const eventIdsInCalendar = new Map();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eventId = row[calendarEventIdIndex];
    if (eventId) eventIdsInSheet.add(eventId);
  }

  const calendarEvents = calendar.getEvents(new Date('2023-01-01'), new Date('2030-01-01'));
  for (const ev of calendarEvents) {
    eventIdsInCalendar.set(ev.getId(), ev);
  }

  // Find events in calendar that aren't tracked in the sheet
  for (const [calendarId, event] of eventIdsInCalendar.entries()) {
    if (!eventIdsInSheet.has(calendarId)) {
      Logger.log(`🕵️ Orphaned calendar event found: "${event.getTitle()}" (${calendarId})`);
      // Optionally delete or keep
      // event.deleteEvent();
    }
  }

  // Find missing calendar events from approved requests
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[headers.indexOf("Status")];
    const eventId = row[calendarEventIdIndex];

    if (status === HOLIDAY_STATUSES.APPROVED && !eventId) {
      Logger.log(`⚠️ Missing calendar event for approved request ${row[holidayRequestIdIndex]}`);
    }
  }
}

function nightlyCleanOldCalendarEvents() {
  const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
  const today = new Date();
  const cutoff = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());

  const events = calendar.getEvents(new Date(2000, 0, 1), cutoff);

  let count = 0;
  for (const ev of events) {
    try {
      ev.deleteEvent();
      count++;
    } catch (e) {
      Logger.log("🧹 Failed to delete old event: " + e.message);
    }
  }

  Logger.log(`🧹 Cleaned ${count} old events from calendar`);
}

function seedHolidayRequestsToCalendar() {
  const calendarId = 'c_0f679b8ddbbc4c4fe84f4be0938b7a8170a6dd47667a7b2dd46675dc4a74523c@group.calendar.google.com';
  Logger.log('🚀 Starting seedHolidayRequestsToCalendar...');
  Logger.log(`📅 Using Calendar ID: ${calendarId}`);

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log('❌ ERROR: Calendar not found. Check the calendar ID and sharing settings.');
    return;
  }

  const userMap = getUserMap(); // 👈 Builds UserID => FullName

  const requests = [
    {
      id: 'HREQ_001',
      userId: 'USR_002',
      startDate: new Date('2025-05-26'),
      endDate: new Date('2025-05-30'),
      status: 'Approved'
    },
    {
      id: 'HREQ_002',
      userId: 'USR_002',
      startDate: new Date('2025-06-09'),
      endDate: new Date('2025-06-13'),
      status: 'Pending'
    },
    {
      id: 'HREQ_003',
      userId: 'USR_004',
      startDate: new Date('2025-06-30'),
      endDate: new Date('2025-07-04'),
      status: 'Rejected'
    }
  ];

  Logger.log(`🧾 Seeding ${requests.length} holiday requests...`);
  let successCount = 0;
  let failCount = 0;

  requests.forEach((req, index) => {
    const fullName = userMap[req.userId] || 'Unknown User';
    const summary = `Annual Leave - ${fullName}`;

    try {
      Logger.log(`➡️ [${index + 1}] Creating event for: ${summary}`);
      Logger.log(`   🔹 Start: ${req.startDate}`);
      Logger.log(`   🔹 End: ${req.endDate}`);
      Logger.log(`   🔹 Status: ${req.status}`);

      const event = calendar.createEvent(summary, req.startDate, req.endDate, {
        description: `Status: ${req.status}\nRequest ID: ${req.id}\nUserID: ${req.userId}`,
        extendedProperties: {
          private: {
            status: req.status,
            requestId: req.id,
            userId: req.userId
          }
        }
      });

      Logger.log(`✅ Event created: ${event.getId()} (${summary})`);
      successCount++;
    } catch (error) {
      Logger.log(`❌ Failed to create event for request ID: ${req.id}`);
      Logger.log(`   🚨 Error: ${error.message}`);
      failCount++;
    }
  });

  Logger.log(`🎯 Done! ${successCount} events created, ${failCount} failed.`);
}

function getUserMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const uidIndex = header.indexOf('UserID');
  const fnameIndex = header.indexOf('FirstName');
  const lnameIndex = header.indexOf('LastName');

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[uidIndex];
    const fullName = `${row[fnameIndex]} ${row[lnameIndex]}`;
    map[id] = fullName;
  }
  return map;
}

function getInitialHolidayStatusByRole(role) {
  switch (role) {
    case "Employee":
      return HOLIDAY_STATUSES.PENDING_MANAGER;
    case "Manager":
      return HOLIDAY_STATUSES.PENDING_CFO;
    case "CFO":
      return HOLIDAY_STATUSES.PENDING_CFO;
    default:
      return HOLIDAY_STATUSES.PENDING_MANAGER; // Fallback
  }
}

function getUsedHolidayDays(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const userIdIndex = headers.indexOf("UserID");
  const daysIndex = headers.indexOf("NumberOfDays");
  const statusIndex = headers.indexOf("Status");

  let used = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusIndex];
    const user = row[userIdIndex];
    const days = parseFloat(row[daysIndex]);

    if (user === userId && days && ["PendingManagerApproval", "PendingCFOApproval", "Approved"].includes(status)) {
      used += days;
    }
  }
  return used;
}

function debugUserCalendars() {
  const email = Session.getActiveUser().getEmail();
  Logger.log("🧪 Running debugUserCalendars...");
  Logger.log(`🔐 Current user email: ${email}`);

  if (email === "darknastyuk@gmail.com") {
    Logger.log("🎯 Debugging for: darknastyuk@gmail.com (✅ matched)");
  } else {
    Logger.log("ℹ️ Not darknastyuk@gmail.com — this may affect access visibility.");
  }

  const calendars = CalendarApp.getAllCalendars();

  if (!calendars.length) {
    Logger.log("⚠️ No calendars accessible to this user.");
    return;
  }

  Logger.log(`📦 Total calendars found: ${calendars.length}`);

  calendars.forEach((cal, index) => {
    Logger.log(`📆 [${index + 1}] ${cal.getName()} → ${cal.getId()}`);
  });

  Logger.log("✅ Calendar debug complete.");
}


