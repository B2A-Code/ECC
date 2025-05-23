// Replace with your actual Calendar IDs
const STAFF_HOLIDAY_CALENDAR_ID = '0vequats8e4v0vr96l62t2eq0fh5023j@import.calendar.google.com';
const CONTRACTOR_AVAILABILITY_CALENDAR_ID = 'c_c4f71d0d3d33796a92c2dfe7ebc381e5bd6fb67bb57b06c00c433d4355f28b4e@group.calendar.google.com';

function addHolidayToCalendar(user, startDate, endDate, summary) {
  const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
  if (!calendar) throw new Error("Staff holiday calendar not found.");

  calendar.createAllDayEvent(summary, new Date(startDate), new Date(endDate), {
    description: `Holiday for ${user.FirstName} ${user.LastName} (${user.Email})`,
    extendedProperties: {
      private: {
        userID: user.UserID,
        userEmail: user.Email,
        type: "holiday"
      }
    }
  });
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

function getMyCalendarEvents(calendarType) {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user) return [];

  let calendarId;
  if (calendarType === "holiday") {
    if (!user.Permanent && user.Role !== "Manager") return [];
    calendarId = STAFF_HOLIDAY_CALENDAR_ID;
  } else if (calendarType === "availability") {
    if (user.Permanent) return [];
    calendarId = CONTRACTOR_AVAILABILITY_CALENDAR_ID;
  } else {
    return [];
  }

  const now = new Date();
  const events = CalendarApp.getCalendarById(calendarId).getEvents(
    now,
    new Date(now.getFullYear() + 1, 11, 31)
  );

  return events
    .filter(ev => ev.getTag("userEmail") === user.Email)
    .map(ev => ({
      id: ev.getId(),
      summary: ev.getTitle(),
      description: ev.getDescription(),
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString()
    }));
}

function getMyHolidayRequests() {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || !(user.Permanent === true || user.Permanent === "TRUE")) return [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const result = [];

  for (const row of data) {
    const record = headers.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {});
    if (record.UserID === user.UserID) {
      result.push(record);
    }
  }

  return result;
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

function getPendingHolidayRequestsForManager() {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || user.Role !== "Manager") return [];

  const requestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");

  const requests = requestsSheet.getDataRange().getValues();
  const headers = requests.shift();

  const users = usersSheet.getDataRange().getValues();
  const userHeaders = users.shift();
  const userMap = {};
  for (const row of users) {
    const u = userHeaders.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {});
    userMap[u.UserID] = u;
  }

  return requests
    .map(row => headers.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {}))
    .filter(req =>
      req.Status === HOLIDAY_STATUSES.PENDING_MANAGER &&
      userMap[req.UserID]?.Department === user.Department
    )
    .map(req => ({
      ...req,
      FullName: `${userMap[req.UserID]?.FirstName} ${userMap[req.UserID]?.LastName}`,
      Email: userMap[req.UserID]?.Email
    }));
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

function getTeamHolidayCalendar() {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || user.Role !== "Manager") return [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HolidayRequests");
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");

  const requestData = sheet.getDataRange().getValues();
  const requestHeaders = requestData.shift();

  const userData = userSheet.getDataRange().getValues();
  const userHeaders = userData.shift();
  const userMap = userData.reduce((map, row) => {
    const u = userHeaders.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {});
    map[u.UserID] = u;
    return map;
  }, {});

  const result = [];

  for (const row of requestData) {
    const record = requestHeaders.reduce((acc, h, i) => {
      acc[h] = row[i];
      return acc;
    }, {});

    const reqUser = userMap[record.UserID];
    if (reqUser && reqUser.Department === user.Department) {
      result.push({
        ...record,
        FullName: `${reqUser.FirstName} ${reqUser.LastName}`,
        Email: reqUser.Email
      });
    }
  }

  return result;
}

function submitAvailabilityViaGAS(payload) {
  try {
    const user = getUserByEmail(Session.getActiveUser().getEmail());
    if (!user || user.Permanent === true || user.Permanent === "TRUE") {
      return { success: false, error: "Only contract staff can submit availability." };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availability");
    if (!sheet) throw new Error("Missing 'Availability' sheet");

    const newId = "AV" + Utilities.getUuid().slice(0, 8);
    const now = new Date();

    const calendar = CalendarApp.getCalendarById(CONTRACTOR_AVAILABILITY_CALENDAR_ID);
    if (!calendar) throw new Error("Calendar not found");

    const event = calendar.createAllDayEvent(payload.Reason || "Unavailable", new Date(payload.StartDate), new Date(payload.EndDate), {
      description: `Availability block for ${user.FirstName} ${user.LastName}`,
      extendedProperties: {
        private: {
          userID: user.UserID,
          userEmail: user.Email,
          type: "availability"
        }
      }
    });

    const eventId = event.getId();

    sheet.appendRow([
      newId,
      user.UserID,
      new Date(payload.StartDate),
      new Date(payload.EndDate),
      payload.Reason || "Unavailable",
      payload.Repeat || "None",
      now,
      now,
      eventId // 🔐 Track calendar event ID
    ]);

    return { success: true };
  } catch (err) {
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
    let numDays = (end - start) / (1000 * 60 * 60 * 24) + 1;

    if (StartType === "Half") numDays -= 0.5;
    if (EndType === "Half") numDays -= 0.5;

    const availableDays = Math.max(0, parseFloat(user.HolidayEntitlementAccruedHours || 0) / 7);
    if (numDays > availableDays) {
      return {
        success: false,
        error: `You only have ${availableDays.toFixed(1)} days left. Requested: ${numDays.toFixed(1)}`
      };
    }

    // Create Google Calendar placeholder event
    const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
    if (!calendar) throw new Error("Holiday calendar not found");

    const calendarEvent = calendar.createAllDayEvent(`Holiday: ${user.FirstName} ${user.LastName}`, start, end, {
      description: `Pending holiday for ${user.FirstName} ${user.LastName}`,
      extendedProperties: {
        private: {
          userID: user.UserID,
          userEmail: user.Email,
          type: "holiday"
        }
      }
    });

    const eventId = calendarEvent.getId(); // 🔐 Track it!

    sheet.appendRow([
      newId,
      user.UserID,
      now.toISOString(),
      StartDate,
      EndDate,
      numDays,
      numDays * 7,
      HOLIDAY_STATUSES.PENDING_MANAGER,
      "",
      "",
      "", // Rejection reason
      now.toISOString(),
      now.toISOString(),
      eventId
    ]);

    const managerEmail = getManagerEmail(user.Department);
    MailApp.sendEmail(
      managerEmail,
      "🆕 Holiday Request Submitted",
      `${user.FirstName} ${user.LastName} has submitted a holiday request.\n\nFrom: ${StartDate}\nTo: ${EndDate}\nDays: ${numDays.toFixed(1)}\n\nPlease review in the system.`
    );

    return { success: true, daysRequested: numDays };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


function approveOrRejectHoliday(requestId, action, reason = "") {
  const user = getUserByEmail(Session.getActiveUser().getEmail());
  if (!user || (user.Role !== "Manager" && user.Role !== "CFO")) {
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
      const requestUser = getUserById(row[userIdIndex]);

      if (action === "approve") {
        const nextStatus = getNextHolidayStatus(currentStatus, user.Role);
        if (!nextStatus) {
          return { success: false, error: "Approval not allowed at this stage" };
        }

        row[statusIndex] = nextStatus;

        if (user.Role === "Manager") {
          row[managerTimestamp] = now.toISOString();

          const cfo = getAllUsers().find(u => u.Role === "CFO");
          if (cfo?.Email) {
            MailApp.sendEmail(
              cfo.Email,
              "🔔 Holiday Request Awaiting CFO Approval",
              `Request ${requestId} from ${requestUser.FirstName} ${requestUser.LastName} has been approved by the Manager and awaits your review.`
            );
          }
        }

        if (user.Role === "CFO") {
          row[cfoTimestamp] = now.toISOString();
          const hoursUsed = parseFloat(row[hoursUsedIndex]);
          deductHolidayHours(row[userIdIndex], hoursUsed);

          const calendar = CalendarApp.getCalendarById(STAFF_HOLIDAY_CALENDAR_ID);
          const event = calendar.createAllDayEvent(
            `Holiday: ${requestUser.FirstName} ${requestUser.LastName}`,
            new Date(row[startDateIndex]),
            new Date(row[endDateIndex]),
            {
              description: `Approved holiday for ${requestUser.Email}`,
              guests: requestUser.Email,
              extendedProperties: {
                private: {
                  userID: requestUser.UserID,
                  userEmail: requestUser.Email,
                  type: "holiday"
                }
              }
            }
          );

          row[calendarEventIdIndex] = event.getId();
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

      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
  }

  return { success: false, error: "Request not found" };
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

      // 🧹 Remove calendar event if it exists
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

function getNextHolidayStatus(currentStatus, role) {
  if (role === 'Manager' && currentStatus === HOLIDAY_STATUSES.PENDING_MANAGER) {
    return HOLIDAY_STATUSES.PENDING_CFO;
  }
  if (role === 'CFO' && currentStatus === HOLIDAY_STATUSES.PENDING_CFO) {
    return HOLIDAY_STATUSES.APPROVED;
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
