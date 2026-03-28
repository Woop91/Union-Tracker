/**
 * ============================================================================
 * 14_MeetingCheckIn.gs - MEETING CHECK-IN SYSTEM
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Meeting management and check-in system. Full lifecycle: steward creates
 *   meeting -> event appears in Google Calendar -> check-in kiosk activates
 *   on meeting day -> members enter email+PIN -> attendance is recorded ->
 *   report is emailed after meeting ends. Supports multiple meetings per day
 *   (member selects which). Auto-deactivates check-in a few hours
 *   (CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS) after the event ends.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Physical check-in kiosks at union meetings use a dialog that stays open
 *   on a shared device. Members authenticate with email+PIN (same as
 *   self-service portal) so a single auth system is maintained. Google
 *   Calendar integration ensures meetings appear on stewards' calendars
 *   automatically. Attendance reports are emailed rather than stored in a
 *   visible sheet to protect member privacy.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Meetings can't be created. Check-in kiosk won't activate. Attendance
 *   isn't recorded. Post-meeting reports aren't emailed. Google Calendar
 *   events aren't created. If the deactivation timer breaks, check-in stays
 *   active indefinitely.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, MEMBER_COLS, MEETING_CHECKIN_COLS,
 *   MEETING_STATUS, COLORS), 05_Integrations.gs (Calendar),
 *   13_MemberSelfService.gs (PIN auth).
 *   Used by menu items in 03_.
 *
 * @version 4.43.1
 * @license Free for use by non-profit collective bargaining groups and unions
 * ============================================================================
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
  // Auth check: only stewards/admins may create meetings
  var callerRole = typeof getUserRole_ === 'function' ? getUserRole_(Session.getActiveUser().getEmail()) : null;
  if (callerRole !== 'steward' && callerRole !== 'admin' && callerRole !== 'both') {
    return errorResponse('Authorization required: steward or admin access needed', 'createMeeting');
  }

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

  // Generate QR check-in URL for in-person attendance
  var qrInfo = null;
  if (typeof getMeetingQRCode === 'function') {
    try {
      qrInfo = getMeetingQRCode(meetingId);
    } catch (_qrErr) {
      Logger.log('createMeeting: QR code generation skipped: ' + _qrErr.message);
    }
  }

  return {
    success: true,
    meetingId: meetingId,
    qrUrl: qrInfo && qrInfo.success ? qrInfo.qrUrl : '',
    checkInUrl: qrInfo && qrInfo.success ? qrInfo.checkInUrl : '',
    message: 'Meeting "' + meetingName + '" created (ID: ' + meetingId + ').' +
             (calendarEventId ? ' Calendar event added.' : '') +
             (notesDocUrl ? ' Meeting Notes doc created.' : '') +
             (agendaDocUrl ? ' Meeting Agenda doc created.' : '') +
             (qrInfo && qrInfo.success ? ' QR check-in code available.' : '')
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

  for (var i = 1; i < memberData.length; i++) {
    var rowEmail = String(memberData[i][MEMBER_COLS.EMAIL - 1] || '').trim().toLowerCase();
    if (rowEmail === email) {
      memberId = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '');
      memberName = String(memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                   String(memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      storedHash = String(memberData[i][MEMBER_PIN_COLS.PIN_HASH - 1] || '');
      break;
    }
  }

  if (!memberId) {
    return errorResponse('Invalid credentials. Please check your information and try again.');
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

// ============================================================================
// QR CODE ATTENDANCE SYSTEM (v4.43.0)
// ============================================================================

/**
 * Process a QR-based meeting check-in using phone number + PIN.
 * Called from the mobile ?page=qr-checkin web route.
 * Looks up member by phone number (normalized), verifies PIN, records attendance.
 *
 * @param {string} meetingId - The meeting to check into
 * @param {string} phone - Member's phone number (any format)
 * @param {string} pin - Member's PIN
 * @returns {Object} { success, memberName, message } or { success: false, error }
 */
function processQRCheckIn(meetingId, phone, pin) {
  if (!meetingId || !phone || !pin) {
    return errorResponse('Meeting, phone number, and PIN are required');
  }

  phone = String(phone).trim();
  pin = String(pin).trim();

  // Normalize phone to digits only for comparison
  var inputDigits = phone.replace(/\D/g, '');
  if (inputDigits.length === 11 && inputDigits.charAt(0) === '1') {
    inputDigits = inputDigits.substring(1); // Strip leading country code
  }
  if (inputDigits.length !== 10) {
    return errorResponse('Please enter a valid 10-digit phone number');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return errorResponse('System temporarily unavailable');

  // Look up member by phone number
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) {
    return errorResponse('System error: Member directory not found');
  }

  var memberData = memberSheet.getDataRange().getValues();
  var memberId = null;
  var memberName = '';
  var memberEmail = '';
  var storedHash = '';

  for (var i = 1; i < memberData.length; i++) {
    var rowPhone = String(memberData[i][MEMBER_COLS.PHONE - 1] || '').trim();
    var rowDigits = rowPhone.replace(/\D/g, '');
    if (rowDigits.length === 11 && rowDigits.charAt(0) === '1') {
      rowDigits = rowDigits.substring(1);
    }
    if (rowDigits.length === 10 && rowDigits === inputDigits) {
      memberId = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '');
      memberName = String(memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                   String(memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      memberEmail = String(memberData[i][MEMBER_COLS.EMAIL - 1] || '').trim().toLowerCase();
      storedHash = String(memberData[i][MEMBER_PIN_COLS.PIN_HASH - 1] || '');
      break;
    }
  }

  if (!memberId) {
    return errorResponse('Invalid credentials. Please check your information and try again.');
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

  // PIN correct — clear failed attempts
  if (typeof clearPINAttempts === 'function') {
    clearPINAttempts(memberId);
  }

  // Acquire lock to prevent TOCTOU race between duplicate check and appendRow
  var checkInLock = LockService.getScriptLock();
  if (!checkInLock.tryLock(10000)) {
    return errorResponse('Check-in temporarily unavailable — please try again.');
  }
  try {
    var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!checkInSheet) {
      return errorResponse('Meeting check-in sheet not found');
    }

    var checkInData = checkInSheet.getDataRange().getValues();

    // Check if already checked in
    for (var j = 1; j < checkInData.length; j++) {
      var rowMeetingId = String(checkInData[j][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      var rowMemberId = String(checkInData[j][MEETING_CHECKIN_COLS.MEMBER_ID - 1] || '');
      if (rowMeetingId === meetingId && rowMemberId === memberId) {
        return errorResponse(memberName.trim() + ' is already checked in to this meeting.');
      }
    }

    // Find meeting details
    var meetingName = '';
    var meetingDate = '';
    var meetingType = '';
    var meetingFound = false;
    for (var k = 1; k < checkInData.length; k++) {
      if (String(checkInData[k][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
        meetingName = checkInData[k][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '';
        meetingDate = checkInData[k][MEETING_CHECKIN_COLS.MEETING_DATE - 1] || '';
        meetingType = checkInData[k][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || '';
        meetingFound = true;
        break;
      }
    }

    if (!meetingFound) {
      return errorResponse('Meeting not found. The QR code may be invalid.');
    }

    // Record the check-in
    checkInSheet.appendRow([
      meetingId,
      meetingName,
      meetingDate,
      meetingType,
      memberId,
      escapeForFormula(memberName.trim()),
      new Date(),
      escapeForFormula(memberEmail)
    ]);

    // Auto-activate Scheduled meetings on first check-in
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
    logAuditEvent('MEETING_QR_CHECKIN', {
      meetingId: meetingId,
      memberId: memberId,
      method: 'qr_phone'
    });
  }

  // Badge refresh
  if (typeof _refreshNavBadges === 'function') _refreshNavBadges();

  return {
    success: true,
    memberName: memberName.trim(),
    message: memberName.trim() + ' checked in successfully!'
  };
}

/**
 * Get the QR check-in URL for a specific meeting.
 * Uses Google Charts API to generate a QR code image that encodes the
 * web app URL with the meeting ID pre-filled.
 *
 * @param {string} meetingId - The meeting ID (e.g., MTG-20260327-001)
 * @returns {Object} { success, qrUrl, checkInUrl } or { success: false, error }
 */
function getMeetingQRCode(meetingId) {
  if (!meetingId) return errorResponse('Meeting ID is required');

  var webAppUrl = '';
  try {
    webAppUrl = ScriptApp.getService().getUrl() || '';
  } catch (_e) {
    Logger.log('getMeetingQRCode: could not get web app URL: ' + _e.message);
  }

  if (!webAppUrl) {
    return errorResponse('Web app URL unavailable. Ensure the script is deployed as a web app.');
  }

  var checkInUrl = webAppUrl + '?page=qr-checkin&meeting=' + encodeURIComponent(meetingId);

  // Google Charts QR code API (free, no key needed, max 300x300)
  var qrImageUrl = 'https://chart.googleapis.com/chart?cht=qr&chs=300x300&chl=' +
    encodeURIComponent(checkInUrl) + '&choe=UTF-8';

  return {
    success: true,
    qrUrl: qrImageUrl,
    checkInUrl: checkInUrl,
    meetingId: meetingId
  };
}

/**
 * Show a dialog with the QR code for a meeting.
 * Stewards display or print this for members to scan at in-person meetings.
 */
function showMeetingQRCodeDialog() {
  var html = HtmlService.createHtmlOutput(getMeetingQRCodeHtml_())
    .setWidth(480)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Meeting QR Check-In Code');
}

/**
 * QR Code display dialog HTML (steward-facing).
 * Shows a dropdown of today's meetings + QR code image + print button.
 * @private
 * @returns {string} HTML content
 */
function getMeetingQRCodeHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;padding:20px}' +
    '.card{background:white;border-radius:12px;padding:25px;box-shadow:0 2px 12px rgba(0,0,0,0.1)}' +
    '.card h2{color:#059669;margin-bottom:16px;text-align:center;font-size:18px}' +
    '.field{margin-bottom:16px}' +
    '.field label{display:block;margin-bottom:6px;font-weight:600;color:#333;font-size:14px}' +
    '.field select{width:100%;padding:10px;border:2px solid #e0e0e0;border-radius:8px;font-size:15px}' +
    '.field select:focus{outline:none;border-color:#059669}' +
    '.qr-container{text-align:center;margin:20px 0;display:none}' +
    '.qr-container img{border:8px solid white;box-shadow:0 2px 12px rgba(0,0,0,0.15);border-radius:8px}' +
    '.qr-instructions{text-align:center;font-size:14px;color:#555;margin:12px 0;line-height:1.5}' +
    '.qr-meeting-name{text-align:center;font-weight:700;color:#059669;font-size:16px;margin:8px 0}' +
    '.btn-row{display:flex;gap:10px;margin-top:16px}' +
    '.btn{flex:1;padding:12px;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer}' +
    '.btn-print{background:#059669;color:white}' +
    '.btn-print:hover{background:#047857}' +
    '.btn-copy{background:#e0e0e0;color:#333}' +
    '.btn-copy:hover{background:#d0d0d0}' +
    '.error{color:#dc2626;font-size:14px;margin-top:10px;text-align:center;padding:10px;background:#fee2e2;border-radius:8px;display:none}' +
    '.hint{font-size:12px;color:#888;text-align:center;margin-top:12px}' +
    '@media print{body{background:white;padding:0}.card{box-shadow:none;border:none}.field,.btn-row,.hint,.error{display:none}.qr-container{display:block!important}}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<h2>Meeting QR Check-In</h2>' +
    '<div class="field">' +
    '<label for="meetingSelect">Select Meeting</label>' +
    '<select id="meetingSelect" onchange="meetingSelected()">' +
    '<option value="">Loading meetings...</option>' +
    '</select>' +
    '</div>' +
    '<div id="qrContainer" class="qr-container">' +
    '<div id="qrMeetingName" class="qr-meeting-name"></div>' +
    '<img id="qrImage" src="" alt="QR Code" width="250" height="250">' +
    '<div class="qr-instructions">Scan with your phone camera to check in.<br>Enter your phone number and PIN when prompted.</div>' +
    '</div>' +
    '<div class="btn-row" id="btnRow" style="display:none">' +
    '<button class="btn btn-print" onclick="window.print()">Print QR Code</button>' +
    '<button class="btn btn-copy" id="copyBtn" onclick="copyLink()">Copy Link</button>' +
    '</div>' +
    '<div id="error" class="error"></div>' +
    '<div class="hint">Members scan this QR code with their phone, then enter their phone number and PIN to check in.</div>' +
    '</div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'var currentCheckInUrl="";' +
    'google.script.run' +
    '  .withSuccessHandler(function(result){' +
    '    var sel=document.getElementById("meetingSelect");' +
    '    if(!result.success||result.meetings.length===0){' +
    '      sel.innerHTML="<option value=\\"\\">No meetings available today</option>";' +
    '      return;' +
    '    }' +
    '    var html="<option value=\\"\\">-- Select a meeting --</option>";' +
    '    result.meetings.forEach(function(m){' +
    '      html+="<option value=\\""+escapeHtml(m.id)+"\\">"+escapeHtml(m.name)+" ("+escapeHtml(m.time)+")</option>";' +
    '    });' +
    '    sel.innerHTML=html;' +
    '    if(result.meetings.length===1){sel.selectedIndex=1;meetingSelected()}' +
    '  })' +
    '  .withFailureHandler(function(){' +
    '    document.getElementById("meetingSelect").innerHTML="<option value=\\"\\">Error loading meetings</option>"' +
    '  })' +
    '  .getCheckInEligibleMeetings();' +
    'function meetingSelected(){' +
    '  var meetingId=document.getElementById("meetingSelect").value;' +
    '  if(!meetingId){' +
    '    document.getElementById("qrContainer").style.display="none";' +
    '    document.getElementById("btnRow").style.display="none";' +
    '    return;' +
    '  }' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      if(r.success){' +
    '        document.getElementById("qrImage").src=r.qrUrl;' +
    '        document.getElementById("qrMeetingName").textContent=document.getElementById("meetingSelect").selectedOptions[0].textContent;' +
    '        document.getElementById("qrContainer").style.display="block";' +
    '        document.getElementById("btnRow").style.display="flex";' +
    '        currentCheckInUrl=r.checkInUrl;' +
    '      }else{' +
    '        document.getElementById("error").textContent=r.error;' +
    '        document.getElementById("error").style.display="block";' +
    '      }' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("error").textContent="Error: "+e.message;' +
    '      document.getElementById("error").style.display="block";' +
    '    })' +
    '    .getMeetingQRCode(meetingId);' +
    '}' +
    'function copyLink(){' +
    '  if(!currentCheckInUrl)return;' +
    '  navigator.clipboard.writeText(currentCheckInUrl).then(function(){' +
    '    document.getElementById("copyBtn").textContent="Copied!";' +
    '    setTimeout(function(){document.getElementById("copyBtn").textContent="Copy Link"},2000);' +
    '  }).catch(function(){' +
    '    prompt("Copy this link:",currentCheckInUrl);' +
    '  });' +
    '}' +
    '</script></body></html>';
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
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
