/**
 * 14_MeetingCheckIn.gs - Meeting Check-In System
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
 * @version 4.7.0
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
    return errorResponse('Meeting name and date are required');
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
    escapeForFormula(meetingName),
    meetingDate,
    meetingType,
    '',  // No member yet - this is the meeting header
    '',
    '',
    '',
    meetingTime,
    meetingDuration,
    initialStatus,
    escapeForFormula(notifyEmails),
    calendarEventId,
    notesDocUrl,
    agendaDocUrl,
    escapeForFormula(agendaStewards)
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
    return errorResponse('Meeting, email, and PIN are required');
  }

  email = String(email).trim().toLowerCase();
  pin = String(pin).trim();

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return errorResponse('Please enter a valid email address');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Look up member by email
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) {
    return errorResponse('System error: Member directory not found');
  }

  var memberData = memberSheet.getDataRange().getValues();
  var memberId = null;
  var memberName = '';
  var storedHash = '';
  var _memberRow = -1;

  for (var i = 1; i < memberData.length; i++) {
    var rowEmail = String(memberData[i][MEMBER_COLS.EMAIL - 1] || '').trim().toLowerCase();
    if (rowEmail === email) {
      memberId = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '');
      memberName = String(memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                   String(memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      storedHash = String(memberData[i][MEMBER_PIN_COLS.PIN_HASH - 1] || '');
      _memberRow = i;
      break;
    }
  }

  if (!memberId) {
    return errorResponse('No member found with that email address');
  }

  // Check for lockout
  if (typeof checkPINLockout === 'function') {
    var lockoutStatus = checkPINLockout(memberId);
    if (lockoutStatus.isLocked) {
      return errorResponse('Account temporarily locked. Try again in ' + lockoutStatus.remainingMinutes + ' minutes.');
    }
  }

  // Verify PIN
  if (!storedHash) {
    return errorResponse('PIN not set. Please contact your steward to set up your PIN.');
  }

  if (!verifyPIN(pin, memberId, storedHash)) {
    if (typeof recordFailedPINAttempt === 'function') {
      var attemptResult = recordFailedPINAttempt(memberId);
      if (attemptResult.isNowLocked) {
        return errorResponse('Too many failed attempts. Account locked for ' + PIN_CONFIG.LOCKOUT_MINUTES + ' minutes.');
      }
      return errorResponse('Invalid PIN. ' + attemptResult.attemptsRemaining + ' attempts remaining.');
    }
    return errorResponse('Invalid PIN');
  }

  // PIN correct - clear failed attempts
  if (typeof clearPINAttempts === 'function') {
    clearPINAttempts(memberId);
  }

  // H-5: Acquire lock to prevent TOCTOU race between duplicate check and appendRow
  var checkInLock = LockService.getScriptLock();
  if (!checkInLock.tryLock(10000)) {
    return errorResponse('Check-in temporarily unavailable — please try again.');
  }
  try {
    // Check if already checked in to this meeting
    var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!checkInSheet) {
      return errorResponse('Meeting check-in sheet not found');
    }

    var checkInData = checkInSheet.getDataRange().getValues();
    for (var j = 1; j < checkInData.length; j++) {
      var rowMeetingId = String(checkInData[j][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      var rowMemberId = String(checkInData[j][MEETING_CHECKIN_COLS.MEMBER_ID - 1] || '');
      if (rowMeetingId === meetingId && rowMemberId === memberId) {
        return errorResponse(memberName.trim() + ' is already checked in to this meeting.');
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
      escapeForFormula(memberName.trim()),
      new Date(),
      escapeForFormula(email)
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
  } finally {
    checkInLock.releaseLock();
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
    return errorResponse('Meeting ID is required');
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

        // Save attendance sheet to Event Check-In/ Drive folder
        if (typeof saveAttendanceToDriveFolder_ === 'function') {
          try {
            saveAttendanceToDriveFolder_(meetingId, data[i]);
          } catch (driveErr) {
            Logger.log('Error saving attendance to Drive for ' + meetingId + ': ' + driveErr.message);
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
    getClientSideEscapeHtml() +
    'document.getElementById("meetingDate").valueAsDate=new Date();' +
    'var stewardData=[];' +
    // Load stewards on dialog open
    'google.script.run.withSuccessHandler(function(list){' +
    '  stewardData=list||[];' +
    '  var container=document.getElementById("stewardListContainer");' +
    '  if(stewardData.length===0){container.textContent="";var em=document.createElement("em");em.style.color="#888";em.textContent="No stewards found";container.appendChild(em);return}' +
    '  var html="";' +
    '  stewardData.forEach(function(s,i){' +
    '    html+="<div class=\\"steward-item\\"><input type=\\"checkbox\\" id=\\"stew_"+i+"\\" value=\\""+escapeHtml(s.email)+"\\"><label for=\\"stew_"+i+"\\">"+escapeHtml(s.name)+" <span class=\\"email\\">"+escapeHtml(s.email)+"</span></label></div>"' +
    '  });' +
    '  container.innerHTML=html;' +
    '}).withFailureHandler(function(){' +
    '  var fc=document.getElementById("stewardListContainer");fc.textContent="";var em2=document.createElement("em");em2.style.color="#888";em2.textContent="Could not load stewards";fc.appendChild(em2)' +
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
    getClientSideEscapeHtml() +

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

// ============================================================================
// WEB ROUTE FUNCTIONS (standalone ?page=checkin kiosk)
// ============================================================================

/**
 * Validates email + PIN and records a meeting check-in.
 * Thin wrapper around processMeetingCheckIn for the web route.
 * @param {string} meetingId - The meeting to check into
 * @param {string} email - Member's email address
 * @param {string} pin - Member's PIN
 * @returns {Object} { success, memberName, message } or { success: false, error }
 */
function webCheckInMember(meetingId, email, pin) {
  return processMeetingCheckIn(meetingId, email, pin);
}

/**
 * Returns a mobile-optimized kiosk-mode HTML page for meeting check-in.
 * Accessible via ?page=checkin web route. Self-contained (no external CDN deps).
 * Large touch-friendly UI designed for shared tablets in landscape or phones.
 * Shows today's eligible meetings, email+PIN form, success feedback, auto-reset.
 * @returns {string} Complete HTML document
 */
function getCheckInPageHtml() {
  var baseUrl = '';
  try { baseUrl = ScriptApp.getService().getUrl(); } catch (_e) { /* ignore */ }

  // Pre-fetch today's eligible meetings server-side for faster initial render
  var meetingsResult = { success: true, meetings: [] };
  try {
    meetingsResult = getCheckInEligibleMeetings();
  } catch (_e) {
    meetingsResult = { success: false, meetings: [] };
  }
  var meetingsJson = JSON.stringify(meetingsResult.meetings || []);

  var html = '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Meeting Check-In | ' + escapeHtml(typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Union Dashboard') : 'Union Dashboard') + '</title>' +
    '<style>' +

    // Reset
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +

    // Body
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;' +
    'background:#fafaf9;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#065f46 0%,#047857 100%);color:white;' +
    'padding:28px 20px 36px;text-align:center;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;' +
    'height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-size:clamp(22px,5vw,30px);font-weight:700;margin-bottom:6px}' +
    '.header .subtitle{font-size:clamp(13px,2.5vw,15px);opacity:0.85}' +

    // Container
    '.container{max-width:520px;margin:0 auto;padding:20px 16px}' +

    // Cards
    '.card{background:white;border-radius:16px;padding:24px;' +
    'box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;margin-bottom:16px}' +
    '.card h2{font-size:18px;font-weight:700;color:#065f46;margin-bottom:16px;text-align:center}' +

    // Meeting option cards
    '.meeting-option{padding:16px;border:2px solid #e7e5e4;border-radius:14px;' +
    'margin-bottom:10px;cursor:pointer;transition:all 0.2s;min-height:56px}' +
    '.meeting-option:hover{border-color:#047857;background:#f0fdf4}' +
    '.meeting-option.selected{border-color:#047857;background:#ecfdf5}' +
    '.meeting-name{font-weight:700;color:#1c1917;font-size:16px;margin-bottom:4px}' +
    '.meeting-meta{font-size:13px;color:#78716c}' +
    '.meeting-badge{display:inline-block;padding:2px 10px;border-radius:20px;' +
    'font-size:11px;font-weight:600;background:#ecfdf5;color:#047857;margin-left:6px}' +

    // Form fields (large touch targets)
    '.field{margin-bottom:20px}' +
    '.field label{display:block;margin-bottom:8px;font-weight:600;color:#1c1917;font-size:15px}' +
    '.field input{width:100%;padding:16px;border:2px solid #e7e5e4;border-radius:12px;' +
    'font-size:18px;font-family:inherit;transition:border-color 0.2s}' +
    '.field input:focus{outline:none;border-color:#047857;box-shadow:0 0 0 3px rgba(4,120,87,0.1)}' +

    // Check-in button (large, thumb-friendly)
    '.btn-checkin{width:100%;padding:18px;border:none;border-radius:14px;font-size:19px;' +
    'font-weight:700;font-family:inherit;cursor:pointer;background:#047857;color:white;' +
    'transition:all 0.2s;min-height:60px}' +
    '.btn-checkin:hover{background:#065f46}' +
    '.btn-checkin:active{transform:scale(0.98)}' +
    '.btn-checkin:disabled{background:#d6d3d1;cursor:not-allowed;transform:none}' +

    // Error message
    '.error-msg{color:#dc2626;font-size:14px;text-align:center;padding:14px;' +
    'background:#fef2f2;border-radius:12px;margin-top:14px;display:none;font-weight:500}' +

    // Success overlay (fullscreen, auto-dismisses)
    '.success-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(4,120,87,0.95);' +
    'display:none;align-items:center;justify-content:center;z-index:1000;flex-direction:column}' +
    '.success-overlay.show{display:flex}' +
    '.success-check{width:80px;height:80px;border-radius:50%;border:4px solid white;' +
    'display:flex;align-items:center;justify-content:center;margin-bottom:20px;' +
    'font-size:40px;color:white;animation:popIn 0.3s ease}' +
    '.success-name{font-size:clamp(22px,5vw,28px);font-weight:700;color:white;margin-bottom:8px;text-align:center}' +
    '.success-msg{font-size:16px;color:rgba(255,255,255,0.85);text-align:center}' +
    '@keyframes popIn{from{transform:scale(0.5);opacity:0}to{transform:scale(1);opacity:1}}' +

    // Attendee count badge
    '.attendee-badge{text-align:center;padding:12px;background:#ecfdf5;border-radius:12px;' +
    'color:#047857;font-weight:600;margin-top:12px;font-size:14px}' +

    // No meetings state
    '.no-meetings{text-align:center;padding:40px 20px;color:#78716c}' +
    '.no-meetings-icon{font-size:48px;margin-bottom:12px;opacity:0.5}' +
    '.no-meetings p{margin-bottom:6px}' +

    // Loading spinner
    '.loading{text-align:center;padding:40px;color:#78716c}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:28px;height:28px;border:3px solid #e7e5e4;' +
    'border-top-color:#047857;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom navigation
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;' +
    'display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));' +
    'box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;' +
    'text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#065f46}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    // Landscape tablet optimizations
    '@media(min-width:768px){' +
    '.container{max-width:600px;padding:30px 24px}' +
    '.field input{font-size:20px;padding:18px}' +
    '.btn-checkin{font-size:22px;padding:22px}' +
    '.meeting-option{padding:20px}' +
    '.meeting-name{font-size:18px}' +
    '}' +

    '</style></head><body>' +

    // Success overlay
    '<div id="successOverlay" class="success-overlay">' +
    '<div class="success-check">&#10003;</div>' +
    '<div class="success-name" id="successName"></div>' +
    '<div class="success-msg">Checked in successfully!</div>' +
    '</div>' +

    // Header
    '<div class="header">' +
    '<h1>Meeting Check-In</h1>' +
    '<div class="subtitle">Enter your email and PIN to check in</div>' +
    '</div>' +

    '<div class="container">' +

    // Meeting selection card
    '<div class="card" id="meetingCard">' +
    '<h2>Select Meeting</h2>' +
    '<div id="meetingList"><div class="loading"><div class="spinner"></div>' +
    '<div style="margin-top:12px">Loading meetings...</div></div></div>' +
    '</div>' +

    // Check-in form card (hidden until meeting selected)
    '<div class="card" id="checkinForm" style="display:none">' +
    '<h2>Your Information</h2>' +
    '<div class="field">' +
    '<label for="email">Email Address</label>' +
    '<input type="email" id="email" placeholder="your.email@example.com" autocomplete="email">' +
    '</div>' +
    '<div class="field">' +
    '<label for="pin">PIN</label>' +
    '<input type="password" id="pin" inputmode="numeric" maxlength="8" placeholder="Enter your PIN" autocomplete="off">' +
    '</div>' +
    '<button class="btn-checkin" id="checkinBtn" onclick="doCheckIn()">Check In</button>' +
    '<div class="error-msg" id="errorMsg"></div>' +
    '<div class="attendee-badge" id="attendeeBadge" style="display:none"></div>' +
    '</div>' +

    '</div>' +

    // Bottom navigation
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '">' +
    '<span class="nav-icon">&#x1F4CA;</span>Home</a>' +
    '<a class="nav-item active" href="' + escapeHtml(baseUrl) + '?page=checkin">' +
    '<span class="nav-icon">&#x2705;</span>Check In</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=selfservice">' +
    '<span class="nav-icon">&#x1F464;</span>My Info</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=links">' +
    '<span class="nav-icon">&#x1F517;</span>Links</a>' +
    '</nav>' +

    '<script>' +
    // Client-side escapeHtml (canonical helper from 00_Security.gs)
    getClientSideEscapeHtml() +

    // State
    'var selectedMeetingId=null;' +

    // Render meetings from server-side pre-fetched data
    'var preloadedMeetings=' + meetingsJson + ';' +
    'renderMeetings(preloadedMeetings);' +

    'function renderMeetings(meetings){' +
    '  var el=document.getElementById("meetingList");' +
    '  if(!meetings||meetings.length===0){' +
    '    el.innerHTML="<div class=\\"no-meetings\\"><div class=\\"no-meetings-icon\\">&#x1F4C5;</div>' +
    '<p style=\\"font-size:18px\\">No active meetings right now</p>' +
    '<p style=\\"font-size:13px\\">Check-in opens on the day of a scheduled meeting</p></div>";' +
    '    return;' +
    '  }' +
    '  if(meetings.length===1){' +
    '    selectedMeetingId=meetings[0].id;' +
    '    el.innerHTML=buildMeetingCard(meetings[0],true);' +
    '    document.getElementById("checkinForm").style.display="block";' +
    '    refreshAttendeeCount(selectedMeetingId);' +
    '  }else{' +
    '    var h="<div style=\\"font-size:14px;color:#78716c;margin-bottom:12px\\">Tap your meeting:</div>";' +
    '    meetings.forEach(function(m){' +
    '      h+="<div class=\\"meeting-option\\" onclick=\\"selectMeeting(\\x27"+escapeHtml(m.id)+"\\x27,this)\\">"' +
    '        +buildMeetingCardInner(m)+"</div>";' +
    '    });' +
    '    el.innerHTML=h;' +
    '  }' +
    '}' +

    'function buildMeetingCard(m,sel){' +
    '  return "<div class=\\"meeting-option"+(sel?" selected":"")+"\\">"+buildMeetingCardInner(m)+"</div>";' +
    '}' +

    'function buildMeetingCardInner(m){' +
    '  var t=escapeHtml(m.type||"");' +
    '  return "<div class=\\"meeting-name\\">"+escapeHtml(m.name)+"</div>"' +
    '    +"<div class=\\"meeting-meta\\">"+escapeHtml(m.time||"")' +
    '    +(t?" <span class=\\"meeting-badge\\">"+t+"</span>":"")+"</div>";' +
    '}' +

    'function selectMeeting(id,el){' +
    '  selectedMeetingId=id;' +
    '  document.querySelectorAll(".meeting-option").forEach(function(o){o.classList.remove("selected")});' +
    '  el.classList.add("selected");' +
    '  document.getElementById("checkinForm").style.display="block";' +
    '  document.getElementById("email").focus();' +
    '  refreshAttendeeCount(id);' +
    '}' +

    'function refreshAttendeeCount(mid){' +
    '  google.script.run.withSuccessHandler(function(r){' +
    '    if(r&&r.success){' +
    '      var el=document.getElementById("attendeeBadge");' +
    '      el.textContent="Checked in: "+r.count+" member"+(r.count!==1?"s":"");' +
    '      el.style.display="block";' +
    '    }' +
    '  }).getMeetingAttendees(mid);' +
    '}' +

    // Check-in action
    'function doCheckIn(){' +
    '  var email=document.getElementById("email").value.trim();' +
    '  var pin=document.getElementById("pin").value.trim();' +
    '  var errEl=document.getElementById("errorMsg");' +
    '  errEl.style.display="none";' +
    '  if(!selectedMeetingId){showErr("Please select a meeting");return}' +
    '  if(!email){showErr("Please enter your email");return}' +
    '  if(!pin||pin.length<4){showErr("Please enter your PIN (4+ digits)");return}' +
    '  var btn=document.getElementById("checkinBtn");' +
    '  btn.disabled=true;btn.textContent="Checking in...";' +
    '  google.script.run' +
    '    .withSuccessHandler(handleResult)' +
    '    .withFailureHandler(function(e){showErr("Error: "+String(e||"Unknown"));resetBtn()})' +
    '    .webCheckInMember(selectedMeetingId,email,pin);' +
    '}' +

    'function handleResult(r){' +
    '  resetBtn();' +
    '  if(r&&r.success){' +
    '    document.getElementById("successName").textContent=r.memberName||"Member";' +
    '    var overlay=document.getElementById("successOverlay");' +
    '    overlay.classList.add("show");' +
    '    document.getElementById("email").value="";' +
    '    document.getElementById("pin").value="";' +
    '    if(selectedMeetingId)refreshAttendeeCount(selectedMeetingId);' +
    '    setTimeout(function(){' +
    '      overlay.classList.remove("show");' +
    '      document.getElementById("email").focus();' +
    '    },3500);' +
    '  }else{' +
    '    showErr((r&&r.error)||"Check-in failed. Verify your email and PIN.");' +
    '  }' +
    '}' +

    'function showErr(msg){var el=document.getElementById("errorMsg");el.textContent=msg;el.style.display="block"}' +
    'function resetBtn(){var b=document.getElementById("checkinBtn");b.disabled=false;b.textContent="Check In"}' +

    // Enter key navigation
    'document.getElementById("email").addEventListener("keypress",function(e){if(e.key==="Enter"){e.preventDefault();document.getElementById("pin").focus()}});' +
    'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter"){e.preventDefault();doCheckIn()}});' +

    '</script></body></html>';

  return html;
}

// ============================================================================
// DRIVE FOLDER INTEGRATION (v4.20.18)
// ============================================================================

/**
 * Saves a completed meeting attendance list as a Google Sheet in the
 * Event Check-In/ Drive subfolder.
 *
 * Called by updateMeetingStatuses() when a meeting is deactivated.
 * Reads the Event Check-In folder ID from Config sheet (MINUTES_FOLDER_ID)
 * and Script Properties (both set by setupDashboardDriveFolders).
 * Fails silently — Drive errors never block check-in functionality.
 *
 * @param {string} meetingId   - Meeting ID
 * @param {Array}  meetingRow  - One data row from MEETING_CHECKIN_LOG for metadata
 */
function saveAttendanceToDriveFolder_(meetingId, meetingRow) {
  if (!meetingId) return;

  // Get folder ID from Config or Script Properties
  var folderId = '';
  try {
    if (typeof getConfigValue_ === 'function' && typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.EVENT_CHECKIN_FOLDER_ID) {
      folderId = getConfigValue_(CONFIG_COLS.EVENT_CHECKIN_FOLDER_ID) || '';
    }
    if (!folderId) {
      folderId = PropertiesService.getScriptProperties().getProperty('EVENT_CHECKIN_FOLDER_ID') || '';
    }
  } catch (_e) {}

  if (!folderId) {
    Logger.log('saveAttendanceToDriveFolder_: Event Check-In folder ID not configured — skipping Drive save for ' + meetingId);
    return;
  }

  // Pull attendance rows for this meeting
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
  if (!checkInSheet || checkInSheet.getLastRow() < 2) return;

  var allData = checkInSheet.getDataRange().getValues();
  var attendees = [];
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
      attendees.push(allData[i]);
    }
  }
  if (attendees.length === 0) {
    Logger.log('saveAttendanceToDriveFolder_: no attendees for ' + meetingId + ', skipping Drive save');
    return;
  }

  // Build meeting metadata from the passed row (or first matching row)
  var metaRow = meetingRow || attendees[0];
  var meetingName = String(metaRow[MEETING_CHECKIN_COLS.MEETING_NAME - 1] || 'Meeting');
  var meetingDate = metaRow[MEETING_CHECKIN_COLS.MEETING_DATE - 1];
  var dateStr = meetingDate instanceof Date
    ? Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(meetingDate || new Date().toDateString());

  var docTitle = meetingName + ' Attendance — ' + dateStr;

  // Create a Google Doc with the attendance list
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();

  body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Meeting ID: ' + meetingId);
  body.appendParagraph('Total Attendees: ' + attendees.length);
  body.appendParagraph('Generated: ' + new Date().toString());
  body.appendParagraph('');
  body.appendParagraph('Attendees').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  attendees.forEach(function(row, idx) {
    var memberId   = String(row[MEETING_CHECKIN_COLS.MEMBER_ID   - 1] || '');
    var memberName = String(row[MEETING_CHECKIN_COLS.MEMBER_NAME - 1] || '');
    var checkInTime = row[MEETING_CHECKIN_COLS.CHECK_IN_TIME - 1];
    var timeStr = checkInTime instanceof Date
      ? Utilities.formatDate(checkInTime, Session.getScriptTimeZone(), 'h:mm a')
      : String(checkInTime || '');
    body.appendListItem((idx + 1) + '. ' + memberName + ' (' + memberId + ') — ' + timeStr);
  });

  doc.saveAndClose();

  // Move the doc into the Event Check-In/ folder
  try {
    var docFile = DriveApp.getFileById(doc.getId());
    var folder  = DriveApp.getFolderById(folderId);
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    Logger.log('saveAttendanceToDriveFolder_: saved "' + docTitle + '" to Event Check-In/ folder');
  } catch (moveErr) {
    Logger.log('saveAttendanceToDriveFolder_: could not move doc to Event Check-In/ folder: ' + moveErr.message);
  }
}
