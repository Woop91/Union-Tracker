/**
 * 509 Dashboard - Meeting Check-In System
 *
 * Stewards can plan meetings in advance. Events appear in Google Calendar.
 * On the day of the event the check-in form becomes active.
 * A few hours after the event ends the check-in becomes inactive.
 * If multiple events are planned for the same day, members select which event.
 * Stewards can opt to receive an attendance report via email.
 *
 * Flow:
 * 1. Steward creates a meeting via "Setup Meeting" dialog (with date, time, duration)
 * 2. Meeting appears in Google Calendar
 * 3. On the day of the meeting, check-in becomes active
 * 4. Steward opens the check-in kiosk dialog (stays open)
 * 5. Members enter email + PIN -> click Check In
 * 6. After meeting duration + buffer, check-in deactivates and attendance is emailed
 *
 * @version 2.0.0
 * @requires 01_Core.gs (SHEETS, MEMBER_COLS, MEETING_CHECKIN_COLS, MEETING_STATUS, COLORS)
 * @requires 05_Integrations.gs (createMeetingCalendarEvent, emailMeetingAttendanceReport)
 * @requires 13_MemberSelfService.gs (authenticateMember, verifyPIN, hashPIN)
 */

// ============================================================================
// MEETING SETUP (Steward-facing)
// ============================================================================

/**
 * Show the Setup Meeting dialog for stewards
 */
function showSetupMeetingDialog() {
  var html = HtmlService.createHtmlOutput(getSetupMeetingHtml_())
    .setWidth(520)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Setup New Meeting');
}

/**
 * Create a new meeting in the Check-In Log sheet
 * Called from the Setup Meeting dialog
 * @param {Object} meetingData - { name, date, time, duration, type, notifyEmails }
 * @returns {Object} { success, meetingId, error }
 */
function createMeeting(meetingData) {
  if (!meetingData || !meetingData.name || !meetingData.date) {
    return { success: false, error: 'Meeting name and date are required' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  // Auto-create sheet if it doesn't exist
  if (!sheet) {
    createMeetingCheckInLogSheet(ss);
    sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
  }

  // Generate meeting ID: MTG-YYYYMMDD-NNN
  var dateStr = meetingData.date.replace(/-/g, '');
  var meetingId = generateMeetingId_(sheet, dateStr);

  var meetingName = String(meetingData.name).substring(0, 200);
  var meetingType = meetingData.type || 'In-Person';
  var meetingDate = new Date(meetingData.date + 'T00:00:00');
  var meetingTime = meetingData.time || '09:00';
  var meetingDuration = parseFloat(meetingData.duration) || 1;
  var notifyEmails = meetingData.notifyEmails ? String(meetingData.notifyEmails).trim() : '';
  var agendaStewards = meetingData.agendaStewards ? String(meetingData.agendaStewards).trim() : '';

  // Create Google Calendar event
  var calendarEventId = '';
  if (typeof createMeetingCalendarEvent === 'function') {
    calendarEventId = createMeetingCalendarEvent({
      meetingId: meetingId,
      name: meetingName,
      date: meetingData.date,
      time: meetingTime,
      duration: meetingDuration,
      type: meetingType
    });
  }

  // Determine initial status: Active if today, Scheduled if future
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var eventDay = new Date(meetingDate);
  eventDay.setHours(0, 0, 0, 0);
  var initialStatus = (eventDay.getTime() === today.getTime()) ? MEETING_STATUS.ACTIVE : MEETING_STATUS.SCHEDULED;

  // Create Meeting Notes and Agenda Google Docs
  var notesDocUrl = '';
  var agendaDocUrl = '';
  if (typeof createMeetingDocs === 'function') {
    var docs = createMeetingDocs({
      meetingId: meetingId,
      name: meetingName,
      date: meetingData.date
    });
    notesDocUrl = docs.notesUrl || '';
    agendaDocUrl = docs.agendaUrl || '';
  }

  // Add a placeholder row so the meeting exists in the sheet
  // Columns A-P: ID, Name, Date, Type, MemberID, MemberName, CheckInTime, Email, Time, Duration, Status, Notify, CalendarEventId, NotesUrl, AgendaUrl, AgendaStewards
  sheet.appendRow([
    meetingId,
    meetingName,
    meetingDate,
    meetingType,
    '',  // No member yet - this is the meeting header
    '',
    '',
    '',
    meetingTime,
    meetingDuration,
    initialStatus,
    notifyEmails,
    calendarEventId,
    notesDocUrl,
    agendaDocUrl,
    agendaStewards
  ]);

  if (typeof logAuditEvent === 'function') {
    logAuditEvent('MEETING_CREATED', {
      meetingId: meetingId,
      meetingName: meetingName,
      meetingDate: meetingData.date,
      meetingTime: meetingTime,
      meetingDuration: meetingDuration,
      meetingType: meetingType,
      calendarEvent: calendarEventId ? 'created' : 'skipped',
      notesDoc: notesDocUrl ? 'created' : 'skipped',
      agendaDoc: agendaDocUrl ? 'created' : 'skipped'
    });
  }

  return {
    success: true,
    meetingId: meetingId,
    message: 'Meeting "' + meetingName + '" created (ID: ' + meetingId + ').' +
             (calendarEventId ? ' Calendar event added.' : '') +
             (notesDocUrl ? ' Meeting Notes doc created.' : '') +
             (agendaDocUrl ? ' Meeting Agenda doc created.' : '')
  };
}

/**
 * Returns a list of stewards with name and email for the setup meeting dialog
 * Called from the HTML dialog to populate steward checkboxes
 * @returns {Array} Array of { name, email } objects
 */
function getStewardEmailsForMeetingSetup() {
  var result = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet || sheet.getLastRow() < 2) return result;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isTruthyValue(isSteward)) {
        var email = String(data[i][MEMBER_COLS.EMAIL - 1] || '').trim();
        if (email) {
          result.push({
            name: (data[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (data[i][MEMBER_COLS.LAST_NAME - 1] || ''),
            email: email
          });
        }
      }
    }
  } catch (e) {
    Logger.log('Error getting steward emails: ' + e.message);
  }
  return result;
}

/**
 * Generate a unique meeting ID
 * @private
 */
function generateMeetingId_(sheet, dateStr) {
  var data = sheet.getDataRange().getValues();
  var prefix = 'MTG-' + dateStr + '-';
  var maxNum = 0;

  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    if (id.indexOf(prefix) === 0) {
      var num = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return prefix + ('00' + (maxNum + 1)).slice(-3);
}

/**
 * Get all scheduled/active meetings (today or future) for the planning view
 * @returns {Object} { success, meetings: [{ id, name, date, type, time, status }] }
 */
function getActiveMeetings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, meetings: [] };
  }

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var meetingsMap = {};

  for (var i = 1; i < data.length; i++) {
    var meetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    if (!meetingId) continue;

    // Only include each meeting once
    if (meetingsMap[meetingId]) continue;

    var eventStatus = String(data[i][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] || '');
    // Skip completed meetings
    if (eventStatus === MEETING_STATUS.COMPLETED) continue;

    var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
    if (meetingDate instanceof Date) {
      var d = new Date(meetingDate);
      d.setHours(0, 0, 0, 0);
      // Include today and future meetings
      if (d >= today) {
        meetingsMap[meetingId] = {
          id: meetingId,
          name: String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || ''),
          date: meetingDate.toLocaleDateString(),
          type: String(data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || ''),
          time: String(data[i][MEETING_CHECKIN_COLS.MEETING_TIME - 1] || ''),
          status: eventStatus || MEETING_STATUS.SCHEDULED
        };
      }
    }
  }

  var meetings = [];
  for (var key in meetingsMap) {
    if (meetingsMap.hasOwnProperty(key)) {
      meetings.push(meetingsMap[key]);
    }
  }

  return { success: true, meetings: meetings };
}

/**
 * Get meetings eligible for check-in (today only, Active or Scheduled status,
 * within the active time window)
 * @returns {Object} { success, meetings: [{ id, name, date, type, time }] }
 */
function getCheckInEligibleMeetings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, meetings: [] };
  }

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var today = new Date(now);
  today.setHours(0, 0, 0, 0);

  var deactivateHours = 3;
  if (typeof CALENDAR_CONFIG !== 'undefined' && CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS) {
    deactivateHours = CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS;
  }

  var meetingsMap = {};

  for (var i = 1; i < data.length; i++) {
    var meetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    if (!meetingId) continue;
    if (meetingsMap[meetingId]) continue;

    var eventStatus = String(data[i][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] || '');
    // Only allow Scheduled or Active meetings for check-in
    if (eventStatus === MEETING_STATUS.COMPLETED) continue;

    var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
    if (!(meetingDate instanceof Date)) continue;

    var d = new Date(meetingDate);
    d.setHours(0, 0, 0, 0);

    // Only today's meetings
    if (d.getTime() !== today.getTime()) continue;

    // Check time window: from start of day to meeting end + deactivate buffer
    var meetingTime = String(data[i][MEETING_CHECKIN_COLS.MEETING_TIME - 1] || '09:00');
    var durationHours = parseFloat(data[i][MEETING_CHECKIN_COLS.MEETING_DURATION - 1]) || 1;
    var timeParts = meetingTime.split(':');
    var startHour = parseInt(timeParts[0], 10) || 9;
    var startMin = parseInt(timeParts[1], 10) || 0;

    var endTime = new Date(d);
    endTime.setHours(startHour + Math.floor(durationHours), startMin + (durationHours % 1) * 60, 0, 0);

    var deactivateTime = new Date(endTime.getTime() + deactivateHours * 60 * 60 * 1000);

    // Meeting is eligible if current time is before deactivation cutoff
    if (now < deactivateTime) {
      meetingsMap[meetingId] = {
        id: meetingId,
        name: String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || ''),
        date: meetingDate.toLocaleDateString(),
        type: String(data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || ''),
        time: meetingTime
      };
    }
  }

  var meetings = [];
  for (var key in meetingsMap) {
    if (meetingsMap.hasOwnProperty(key)) {
      meetings.push(meetingsMap[key]);
    }
  }

  return { success: true, meetings: meetings };
}

// ============================================================================
// MEETING CHECK-IN (Member-facing kiosk)
// ============================================================================

/**
 * Show the Meeting Check-In kiosk dialog
 * This stays open so multiple members can check in sequentially
 */
function showMeetingCheckInDialog() {
  var html = HtmlService.createHtmlOutput(getMeetingCheckInHtml_())
    .setWidth(600)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Meeting Check-In');
}

/**
 * Process a member check-in
 * Authenticates by email + PIN, then logs attendance
 * @param {string} meetingId - The meeting to check into
 * @param {string} email - Member's email address
 * @param {string} pin - Member's 6-digit PIN
 * @returns {Object} { success, memberName, error }
 */
function processMeetingCheckIn(meetingId, email, pin) {
  // Validate inputs
  if (!meetingId || !email || !pin) {
    return { success: false, error: 'Meeting, email, and PIN are required' };
  }

  email = String(email).trim().toLowerCase();
  pin = String(pin).trim();

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, error: 'Please enter a valid email address' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Look up member by email
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) {
    return { success: false, error: 'System error: Member directory not found' };
  }

  var memberData = memberSheet.getDataRange().getValues();
  var memberId = null;
  var memberName = '';
  var storedHash = '';
  var memberRow = -1;

  for (var i = 1; i < memberData.length; i++) {
    var rowEmail = String(memberData[i][MEMBER_COLS.EMAIL - 1] || '').trim().toLowerCase();
    if (rowEmail === email) {
      memberId = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '');
      memberName = String(memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                   String(memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      storedHash = String(memberData[i][MEMBER_PIN_COLS.PIN_HASH - 1] || '');
      memberRow = i;
      break;
    }
  }

  if (!memberId) {
    return { success: false, error: 'No member found with that email address' };
  }

  // Check for lockout
  if (typeof checkPINLockout === 'function') {
    var lockoutStatus = checkPINLockout(memberId);
    if (lockoutStatus.isLocked) {
      return {
        success: false,
        error: 'Account temporarily locked. Try again in ' + lockoutStatus.remainingMinutes + ' minutes.'
      };
    }
  }

  // Verify PIN
  if (!storedHash) {
    return {
      success: false,
      error: 'PIN not set. Please contact your steward to set up your PIN.'
    };
  }

  if (!verifyPIN(pin, memberId, storedHash)) {
    if (typeof recordFailedPINAttempt === 'function') {
      var attemptResult = recordFailedPINAttempt(memberId);
      if (attemptResult.isNowLocked) {
        return {
          success: false,
          error: 'Too many failed attempts. Account locked for ' + PIN_CONFIG.LOCKOUT_MINUTES + ' minutes.'
        };
      }
      return {
        success: false,
        error: 'Invalid PIN. ' + attemptResult.attemptsRemaining + ' attempts remaining.'
      };
    }
    return { success: false, error: 'Invalid PIN' };
  }

  // PIN correct - clear failed attempts
  if (typeof clearPINAttempts === 'function') {
    clearPINAttempts(memberId);
  }

  // Check if already checked in to this meeting
  var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
  if (!checkInSheet) {
    return { success: false, error: 'Meeting check-in sheet not found' };
  }

  var checkInData = checkInSheet.getDataRange().getValues();
  for (var j = 1; j < checkInData.length; j++) {
    var rowMeetingId = String(checkInData[j][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    var rowMemberId = String(checkInData[j][MEETING_CHECKIN_COLS.MEMBER_ID - 1] || '');
    if (rowMeetingId === meetingId && rowMemberId === memberId) {
      return {
        success: false,
        error: memberName.trim() + ' is already checked in to this meeting.'
      };
    }
  }

  // Find meeting details from the first row with this meeting ID
  var meetingName = '';
  var meetingDate = '';
  var meetingType = '';
  for (var k = 1; k < checkInData.length; k++) {
    if (String(checkInData[k][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
      meetingName = checkInData[k][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '';
      meetingDate = checkInData[k][MEETING_CHECKIN_COLS.MEETING_DATE - 1] || '';
      meetingType = checkInData[k][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || '';
      break;
    }
  }

  // Record the check-in (columns A-H only; I-M are meeting-level metadata)
  checkInSheet.appendRow([
    meetingId,
    meetingName,
    meetingDate,
    meetingType,
    memberId,
    memberName.trim(),
    new Date(),
    email
  ]);

  // If this is a Scheduled meeting getting its first check-in, mark as Active
  for (var m = 1; m < checkInData.length; m++) {
    if (String(checkInData[m][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
      var currentStatus = String(checkInData[m][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] || '');
      if (currentStatus === MEETING_STATUS.SCHEDULED) {
        checkInSheet.getRange(m + 1, MEETING_CHECKIN_COLS.EVENT_STATUS).setValue(MEETING_STATUS.ACTIVE);
      }
      break;
    }
  }

  // Audit log
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('MEETING_CHECKIN', {
      meetingId: meetingId,
      memberId: memberId
    });
  }

  return {
    success: true,
    memberName: memberName.trim(),
    message: memberName.trim() + ' checked in successfully!'
  };
}

/**
 * Get attendance list for a specific meeting
 * @param {string} meetingId - The meeting ID
 * @returns {Object} { success, attendees, count }
 */
function getMeetingAttendees(meetingId) {
  if (!meetingId) {
    return { success: false, error: 'Meeting ID is required' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, attendees: [], count: 0 };
  }

  var data = sheet.getDataRange().getValues();
  var attendees = [];

  for (var i = 1; i < data.length; i++) {
    var rowMeetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    var rowMemberId = String(data[i][MEETING_CHECKIN_COLS.MEMBER_ID - 1] || '');
    if (rowMeetingId === meetingId && rowMemberId) {
      attendees.push({
        memberId: rowMemberId,
        name: String(data[i][MEETING_CHECKIN_COLS.MEMBER_NAME - 1] || ''),
        time: data[i][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1]
          ? new Date(data[i][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1]).toLocaleTimeString()
          : ''
      });
    }
  }

  return { success: true, attendees: attendees, count: attendees.length };
}

// ============================================================================
// EVENT LIFECYCLE MANAGEMENT
// ============================================================================

/**
 * Activate today's scheduled meetings and deactivate expired ones.
 * Called from dailyTrigger or hourly trigger.
 * - Meetings dated today with status "Scheduled" become "Active"
 * - Meetings whose end time + buffer has passed become "Completed"
 * - Completed meetings with notify steward emails get attendance report emailed
 * @returns {Object} { activated, deactivated }
 */
function updateMeetingStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { activated: 0, deactivated: 0 };
  }

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var today = new Date(now);
  today.setHours(0, 0, 0, 0);

  var deactivateHours = 3;
  if (typeof CALENDAR_CONFIG !== 'undefined' && CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS) {
    deactivateHours = CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS;
  }

  var activated = 0;
  var deactivated = 0;
  var processedIds = {};

  for (var i = 1; i < data.length; i++) {
    var meetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
    if (!meetingId || processedIds[meetingId]) continue;
    processedIds[meetingId] = true;

    var eventStatus = String(data[i][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] || '');
    var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
    if (!(meetingDate instanceof Date)) continue;

    var d = new Date(meetingDate);
    d.setHours(0, 0, 0, 0);

    // Activate: Scheduled meetings dated today
    if (eventStatus === MEETING_STATUS.SCHEDULED && d.getTime() === today.getTime()) {
      sheet.getRange(i + 1, MEETING_CHECKIN_COLS.EVENT_STATUS).setValue(MEETING_STATUS.ACTIVE);
      activated++;
    }

    // Deactivate: Active (or Scheduled past-date) meetings that have expired
    if (eventStatus === MEETING_STATUS.ACTIVE || eventStatus === MEETING_STATUS.SCHEDULED) {
      var meetingTime = String(data[i][MEETING_CHECKIN_COLS.MEETING_TIME - 1] || '09:00');
      var durationHours = parseFloat(data[i][MEETING_CHECKIN_COLS.MEETING_DURATION - 1]) || 1;
      var timeParts = meetingTime.split(':');
      var startHour = parseInt(timeParts[0], 10) || 9;
      var startMin = parseInt(timeParts[1], 10) || 0;

      var endTime = new Date(d);
      endTime.setHours(startHour + Math.floor(durationHours), startMin + (durationHours % 1) * 60, 0, 0);
      var deactivateTime = new Date(endTime.getTime() + deactivateHours * 60 * 60 * 1000);

      if (now >= deactivateTime) {
        sheet.getRange(i + 1, MEETING_CHECKIN_COLS.EVENT_STATUS).setValue(MEETING_STATUS.COMPLETED);
        deactivated++;

        // Email attendance report if steward emails are configured
        var notifyEmails = String(data[i][MEETING_CHECKIN_COLS.NOTIFY_STEWARDS - 1] || '').trim();
        if (notifyEmails && typeof emailMeetingAttendanceReport === 'function') {
          try {
            emailMeetingAttendanceReport(meetingId, notifyEmails);
          } catch (emailError) {
            Logger.log('Error emailing attendance for ' + meetingId + ': ' + emailError.message);
          }
        }
      }
    }

    // Also deactivate past-date Scheduled meetings (event day has passed entirely)
    if (eventStatus === MEETING_STATUS.SCHEDULED && d < today) {
      sheet.getRange(i + 1, MEETING_CHECKIN_COLS.EVENT_STATUS).setValue(MEETING_STATUS.COMPLETED);
      deactivated++;
    }
  }

  if (activated > 0 || deactivated > 0) {
    Logger.log('Meeting status update: ' + activated + ' activated, ' + deactivated + ' deactivated');
  }

  return { activated: activated, deactivated: deactivated };
}

// ============================================================================
// CLEANUP
// ============================================================================

/**
 * Archive expired meeting records (meetings older than 90 days)
 * Called from dailyTrigger to keep the check-in log manageable.
 * Rows are deleted (not moved) since the data is attendance-only.
 * @returns {number} Number of rows removed
 */
function cleanupExpiredMeetings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);

  if (!sheet || sheet.getLastRow() < 2) return 0;

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 90);
  cutoff.setHours(0, 0, 0, 0);

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
    if (meetingDate instanceof Date && meetingDate < cutoff) {
      rowsToDelete.push(i + 1); // 1-indexed sheet row
    }
  }

  // Delete bottom-up to preserve row indices
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  if (rowsToDelete.length > 0 && typeof logAuditEvent === 'function') {
    logAuditEvent('MEETING_CLEANUP', {
      rowsRemoved: rowsToDelete.length,
      cutoffDate: cutoff.toISOString()
    });
  }

  return rowsToDelete.length;
}

// ============================================================================
// HTML TEMPLATES
// ============================================================================

/**
 * Setup Meeting dialog HTML (steward-facing)
 * Includes time, duration, and notification email fields
 * @private
 * @returns {string} HTML content
 */
function getSetupMeetingHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;padding:20px}' +
    '.card{background:white;border-radius:12px;padding:30px;box-shadow:0 2px 12px rgba(0,0,0,0.1)}' +
    '.card h2{color:#059669;margin-bottom:20px;text-align:center}' +
    '.field{margin-bottom:16px}' +
    '.field label{display:block;margin-bottom:6px;font-weight:600;color:#333;font-size:14px}' +
    '.field input,.field select,.field textarea{width:100%;padding:10px;border:2px solid #e0e0e0;border-radius:8px;font-size:15px}' +
    '.field input:focus,.field select:focus,.field textarea:focus{outline:none;border-color:#059669}' +
    '.field .hint{font-size:12px;color:#888;margin-top:4px}' +
    '.row{display:flex;gap:12px}' +
    '.row .field{flex:1}' +
    '.steward-list{max-height:150px;overflow-y:auto;border:2px solid #e0e0e0;border-radius:8px;padding:8px}' +
    '.steward-item{display:flex;align-items:center;gap:8px;padding:6px 4px;font-size:14px}' +
    '.steward-item input[type=checkbox]{width:18px;height:18px;cursor:pointer}' +
    '.steward-item label{cursor:pointer;flex:1}' +
    '.steward-item .email{color:#888;font-size:12px}' +
    '.select-actions{display:flex;gap:8px;margin-top:6px}' +
    '.select-actions button{padding:4px 10px;font-size:12px;border:1px solid #ddd;border-radius:4px;background:#f5f5f5;cursor:pointer}' +
    '.select-actions button:hover{background:#e0e0e0}' +
    '.btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;background:#059669;color:white;margin-top:8px}' +
    '.btn:hover{background:#047857}' +
    '.btn:disabled{background:#ccc;cursor:not-allowed}' +
    '.error{color:#dc2626;font-size:14px;margin-top:10px;text-align:center;padding:10px;background:#fee2e2;border-radius:8px;display:none}' +
    '.success{color:#059669;font-size:14px;margin-top:10px;text-align:center;padding:10px;background:#d1fae5;border-radius:8px;display:none}' +
    '.section-label{font-weight:700;color:#059669;font-size:13px;margin:20px 0 8px;padding-top:12px;border-top:1px solid #e0e0e0}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<h2>Schedule Meeting</h2>' +
    '<div class="field">' +
    '<label for="meetingName">Meeting Name</label>' +
    '<input type="text" id="meetingName" placeholder="e.g. Monthly General Meeting">' +
    '</div>' +
    '<div class="row">' +
    '<div class="field">' +
    '<label for="meetingDate">Date</label>' +
    '<input type="date" id="meetingDate">' +
    '</div>' +
    '<div class="field">' +
    '<label for="meetingTime">Start Time</label>' +
    '<input type="time" id="meetingTime" value="12:00">' +
    '</div>' +
    '</div>' +
    '<div class="row">' +
    '<div class="field">' +
    '<label for="meetingDuration">Duration (hours)</label>' +
    '<select id="meetingDuration">' +
    '<option value="0.5">30 min</option>' +
    '<option value="1" selected>1 hour</option>' +
    '<option value="1.5">1.5 hours</option>' +
    '<option value="2">2 hours</option>' +
    '<option value="3">3 hours</option>' +
    '<option value="4">4 hours</option>' +
    '</select>' +
    '</div>' +
    '<div class="field">' +
    '<label for="meetingType">Type</label>' +
    '<select id="meetingType">' +
    '<option value="In-Person">In-Person</option>' +
    '<option value="Virtual">Virtual</option>' +
    '<option value="Hybrid">Hybrid</option>' +
    '</select>' +
    '</div>' +
    '</div>' +
    '<div class="section-label">Steward Notifications</div>' +
    '<div class="field">' +
    '<label>Email Attendance Report To</label>' +
    '<input type="text" id="notifyEmails" placeholder="steward@email.com, steward2@email.com">' +
    '<div class="hint">Comma-separated emails to receive attendance report after the meeting</div>' +
    '</div>' +
    '<div class="field">' +
    '<label>Send Agenda Early To (3 days prior)</label>' +
    '<div id="stewardListContainer" class="steward-list"><em style="color:#888">Loading stewards...</em></div>' +
    '<div class="select-actions">' +
    '<button type="button" onclick="selectAllStewards()">Select All</button>' +
    '<button type="button" onclick="clearAllStewards()">Clear All</button>' +
    '</div>' +
    '<div class="hint">Selected stewards get the agenda link 3 days before. All stewards get it the day before regardless.</div>' +
    '</div>' +
    '<button class="btn" id="createBtn" onclick="createMeeting()">Schedule Meeting</button>' +
    '<div id="error" class="error"></div>' +
    '<div id="success" class="success"></div>' +
    '</div>' +
    '<script>' +
    'document.getElementById("meetingDate").valueAsDate=new Date();' +
    'var stewardData=[];' +
    // Load stewards on dialog open
    'google.script.run.withSuccessHandler(function(list){' +
    '  stewardData=list||[];' +
    '  var container=document.getElementById("stewardListContainer");' +
    '  if(stewardData.length===0){container.innerHTML="<em style=\\"color:#888\\">No stewards found</em>";return}' +
    '  var html="";' +
    '  stewardData.forEach(function(s,i){' +
    '    html+="<div class=\\"steward-item\\"><input type=\\"checkbox\\" id=\\"stew_"+i+"\\" value=\\""+s.email+"\\"><label for=\\"stew_"+i+"\\">"+s.name+" <span class=\\"email\\">"+s.email+"</span></label></div>"' +
    '  });' +
    '  container.innerHTML=html;' +
    '}).withFailureHandler(function(){' +
    '  document.getElementById("stewardListContainer").innerHTML="<em style=\\"color:#888\\">Could not load stewards</em>"' +
    '}).getStewardEmailsForMeetingSetup();' +
    'function selectAllStewards(){document.querySelectorAll("#stewardListContainer input[type=checkbox]").forEach(function(cb){cb.checked=true})}' +
    'function clearAllStewards(){document.querySelectorAll("#stewardListContainer input[type=checkbox]").forEach(function(cb){cb.checked=false})}' +
    'function getSelectedStewardEmails(){' +
    '  var emails=[];' +
    '  document.querySelectorAll("#stewardListContainer input[type=checkbox]:checked").forEach(function(cb){emails.push(cb.value)});' +
    '  return emails.join(", ");' +
    '}' +
    'function createMeeting(){' +
    '  var name=document.getElementById("meetingName").value.trim();' +
    '  var date=document.getElementById("meetingDate").value;' +
    '  var time=document.getElementById("meetingTime").value;' +
    '  var duration=document.getElementById("meetingDuration").value;' +
    '  var type=document.getElementById("meetingType").value;' +
    '  var notifyEmails=document.getElementById("notifyEmails").value.trim();' +
    '  var agendaStewards=getSelectedStewardEmails();' +
    '  if(!name){showError("Please enter a meeting name");return}' +
    '  if(!date){showError("Please select a date");return}' +
    '  document.getElementById("createBtn").disabled=true;' +
    '  document.getElementById("createBtn").textContent="Scheduling...";' +
    '  document.getElementById("error").style.display="none";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("createBtn").disabled=false;' +
    '      document.getElementById("createBtn").textContent="Schedule Meeting";' +
    '      if(r.success){' +
    '        document.getElementById("success").textContent=r.message;' +
    '        document.getElementById("success").style.display="block";' +
    '        setTimeout(function(){google.script.host.close()},2500);' +
    '      }else{showError(r.error)}' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("createBtn").disabled=false;' +
    '      document.getElementById("createBtn").textContent="Schedule Meeting";' +
    '      showError("Error: "+e.message);' +
    '    })' +
    '    .createMeeting({name:name,date:date,time:time,duration:duration,type:type,notifyEmails:notifyEmails,agendaStewards:agendaStewards});' +
    '}' +
    'function showError(msg){' +
    '  var el=document.getElementById("error");' +
    '  el.textContent=msg;el.style.display="block";' +
    '}' +
    '</script></body></html>';
}

/**
 * Meeting Check-In kiosk dialog HTML (member-facing)
 * Only shows today's eligible meetings. If multiple, user selects.
 * @private
 * @returns {string} HTML content
 */
function getMeetingCheckInHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header
    '.header{background:linear-gradient(135deg,#059669,#047857);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:22px;margin-bottom:5px}' +
    '.header .subtitle{font-size:13px;opacity:0.9}' +

    // Container
    '.container{max-width:550px;margin:0 auto;padding:20px}' +

    // Card
    '.card{background:white;border-radius:12px;padding:25px;box-shadow:0 2px 12px rgba(0,0,0,0.1);margin-bottom:15px}' +
    '.card h2{color:#059669;margin-bottom:15px;text-align:center;font-size:18px}' +

    // Form fields
    '.field{margin-bottom:18px}' +
    '.field label{display:block;margin-bottom:8px;font-weight:600;color:#333;font-size:14px}' +
    '.field input,.field select{width:100%;padding:12px;border:2px solid #e0e0e0;border-radius:8px;font-size:16px;transition:border-color 0.2s}' +
    '.field input:focus,.field select:focus{outline:none;border-color:#059669}' +

    // Button
    '.btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;transition:all 0.2s}' +
    '.btn-checkin{background:#059669;color:white;font-size:18px;padding:16px}' +
    '.btn-checkin:hover{background:#047857}' +
    '.btn-checkin:disabled{background:#ccc;cursor:not-allowed}' +

    // Messages
    '.error{color:#dc2626;font-size:14px;margin-top:12px;text-align:center;padding:12px;background:#fee2e2;border-radius:8px;display:none}' +
    '.success-banner{text-align:center;padding:25px;background:#d1fae5;border-radius:12px;margin-bottom:15px;display:none}' +
    '.success-banner .checkmark{font-size:48px;margin-bottom:10px}' +
    '.success-banner .name{font-size:20px;font-weight:700;color:#059669;margin-bottom:5px}' +
    '.success-banner .msg{color:#047857;font-size:14px}' +

    // Attendee count
    '.attendee-count{text-align:center;padding:12px;background:#f0fdf4;border-radius:8px;color:#059669;font-weight:600;margin-top:10px}' +

    // No meetings
    '.no-meetings{text-align:center;padding:30px;color:#666}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h1>Meeting Check-In</h1>' +
    '<div class="subtitle">Enter your email and PIN to check in</div>' +
    '</div>' +

    '<div class="container">' +

    // Success banner (shown briefly after check-in)
    '<div id="successBanner" class="success-banner">' +
    '<div class="checkmark">&#10003;</div>' +
    '<div class="name" id="successName"></div>' +
    '<div class="msg">Checked in successfully!</div>' +
    '</div>' +

    // Meeting selector
    '<div class="card" id="meetingCard">' +
    '<h2>Select Meeting</h2>' +
    '<div class="field">' +
    '<label for="meetingSelect">Active Meeting</label>' +
    '<select id="meetingSelect" onchange="meetingChanged()">' +
    '<option value="">Loading meetings...</option>' +
    '</select>' +
    '</div>' +
    '<div id="attendeeCount" class="attendee-count" style="display:none"></div>' +
    '</div>' +

    // Check-in form
    '<div class="card" id="checkinForm" style="display:none">' +
    '<h2>Check In</h2>' +
    '<div class="field">' +
    '<label for="email">Email Address</label>' +
    '<input type="email" id="email" placeholder="Enter your email" autocomplete="email">' +
    '</div>' +
    '<div class="field">' +
    '<label for="pin">PIN</label>' +
    '<input type="password" id="pin" placeholder="Enter your 6-digit PIN" maxlength="6" pattern="[0-9]*" inputmode="numeric" autocomplete="off">' +
    '</div>' +
    '<button class="btn btn-checkin" id="checkinBtn" onclick="checkIn()">Check In</button>' +
    '<div id="error" class="error"></div>' +
    '</div>' +

    '<div id="noMeetings" class="no-meetings" style="display:none">' +
    '<p style="font-size:18px;margin-bottom:10px">No active meetings for today</p>' +
    '<p style="font-size:14px">Check-in opens on the day of a scheduled event.</p>' +
    '</div>' +

    '</div>' +

    '<script>' +
    'function escapeHtml(t){if(t==null)return"";return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#x27;")}' +

    // Load only today's eligible meetings for check-in
    'google.script.run' +
    '  .withSuccessHandler(populateMeetings)' +
    '  .withFailureHandler(function(){document.getElementById("meetingSelect").innerHTML="<option value=\\"\\">Error loading meetings</option>"})' +
    '  .getCheckInEligibleMeetings();' +

    'function populateMeetings(result){' +
    '  var sel=document.getElementById("meetingSelect");' +
    '  if(!result.success||result.meetings.length===0){' +
    '    sel.innerHTML="<option value=\\"\\">No meetings available</option>";' +
    '    document.getElementById("noMeetings").style.display="block";' +
    '    document.getElementById("meetingCard").style.display="none";' +
    '    return;' +
    '  }' +
    '  var html="<option value=\\"\\">-- Select a meeting --</option>";' +
    '  result.meetings.forEach(function(m){' +
    '    var label=escapeHtml(m.name)+" ("+escapeHtml(m.type);' +
    '    if(m.time)label+=", "+escapeHtml(m.time);' +
    '    label+=")";' +
    '    html+="<option value=\\""+escapeHtml(m.id)+"\\">"+label+"</option>";' +
    '  });' +
    '  sel.innerHTML=html;' +
    '  if(result.meetings.length===1){' +
    '    sel.selectedIndex=1;' +
    '    meetingChanged();' +
    '  }' +
    '}' +

    'function meetingChanged(){' +
    '  var meetingId=document.getElementById("meetingSelect").value;' +
    '  if(meetingId){' +
    '    document.getElementById("checkinForm").style.display="block";' +
    '    document.getElementById("email").focus();' +
    '    refreshAttendeeCount(meetingId);' +
    '  }else{' +
    '    document.getElementById("checkinForm").style.display="none";' +
    '    document.getElementById("attendeeCount").style.display="none";' +
    '  }' +
    '}' +

    'function refreshAttendeeCount(meetingId){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      if(r.success){' +
    '        var el=document.getElementById("attendeeCount");' +
    '        el.textContent="Checked in: "+r.count+" member"+(r.count!==1?"s":"");' +
    '        el.style.display="block";' +
    '      }' +
    '    })' +
    '    .getMeetingAttendees(meetingId);' +
    '}' +

    // Check-in
    'function checkIn(){' +
    '  var meetingId=document.getElementById("meetingSelect").value;' +
    '  var email=document.getElementById("email").value.trim();' +
    '  var pin=document.getElementById("pin").value.trim();' +
    '  var errEl=document.getElementById("error");' +
    '  errEl.style.display="none";' +
    '  if(!meetingId){showError("Please select a meeting");return}' +
    '  if(!email){showError("Please enter your email");return}' +
    '  if(!pin||pin.length<4){showError("Please enter your PIN");return}' +
    '  document.getElementById("checkinBtn").disabled=true;' +
    '  document.getElementById("checkinBtn").textContent="Checking in...";' +
    '  google.script.run' +
    '    .withSuccessHandler(handleCheckInResult)' +
    '    .withFailureHandler(function(e){showError("Error: "+e.message);resetBtn()})' +
    '    .processMeetingCheckIn(meetingId,email,pin);' +
    '}' +

    'function handleCheckInResult(result){' +
    '  resetBtn();' +
    '  if(result.success){' +
    '    document.getElementById("successName").textContent=result.memberName;' +
    '    document.getElementById("successBanner").style.display="block";' +
    '    document.getElementById("email").value="";' +
    '    document.getElementById("pin").value="";' +
    '    document.getElementById("error").style.display="none";' +
    '    var meetingId=document.getElementById("meetingSelect").value;' +
    '    refreshAttendeeCount(meetingId);' +
    '    setTimeout(function(){' +
    '      document.getElementById("successBanner").style.display="none";' +
    '      document.getElementById("email").focus();' +
    '    },3000);' +
    '  }else{' +
    '    showError(result.error);' +
    '  }' +
    '}' +

    'function showError(msg){' +
    '  var el=document.getElementById("error");' +
    '  el.textContent=msg;el.style.display="block";' +
    '}' +

    'function resetBtn(){' +
    '  document.getElementById("checkinBtn").disabled=false;' +
    '  document.getElementById("checkinBtn").textContent="Check In";' +
    '}' +

    // Enter key handlers
    'document.getElementById("email").addEventListener("keypress",function(e){if(e.key==="Enter")document.getElementById("pin").focus()});' +
    'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter")checkIn()});' +

    '</script></body></html>';
}
