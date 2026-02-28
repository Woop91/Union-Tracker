/**
 * ============================================================================
 * 05_Integrations.gs - External Service Integration
 * ============================================================================
 *
 * This module handles all interactions with external Google services:
 * - Google Drive folder management for grievance documents
 * - Google Calendar deadline synchronization
 * - Email notifications
 * - External API calls
 *
 * SEPARATION OF CONCERNS: Isolating external dependencies ensures that if
 * one service (e.g., Drive) has an outage, core spreadsheet functionality
 * remains responsive.
 *
 * @fileoverview External service integrations
 * @version 4.7.0
 * @requires 01_Core.gs
 */

// ============================================================================
// CALENDAR CONFIGURATION
// ============================================================================

/**
 * Calendar configuration for deadline tracking
 * @const {Object}
 */
var CALENDAR_CONFIG = {
  CALENDAR_NAME: 'Grievance Deadlines',
  MEETING_CALENDAR_NAME: 'Union Meetings',
  REMINDER_DAYS: [7, 3, 1],
  MEETING_DEACTIVATE_HOURS: 3  // Hours after meeting end to deactivate check-in
};

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

/**
 * Gets or creates the root folder for grievance files
 * @return {Folder} The root grievance folder
 */
function getOrCreateRootFolder() {
  // Check for stored folder ID first to avoid global name-search ambiguity
  var props = PropertiesService.getScriptProperties();
  var storedFolderId = props.getProperty('GRIEVANCE_ROOT_FOLDER_ID');
  if (storedFolderId) {
    try {
      return DriveApp.getFolderById(storedFolderId);
    } catch (_e) {
      // Stored ID invalid, fall through to name search
    }
  }

  const folderName = DRIVE_CONFIG.ROOT_FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    var existing = folders.next();
    props.setProperty('GRIEVANCE_ROOT_FOLDER_ID', existing.getId());
    return existing;
  }

  // Create the root folder
  const newFolder = DriveApp.createFolder(folderName);
  props.setProperty('GRIEVANCE_ROOT_FOLDER_ID', newFolder.getId());

  // Set folder color/description
  newFolder.setDescription('Union Grievance Documentation - Auto-managed by Dashboard');

  logAuditEvent(AUDIT_EVENTS.FOLDER_CREATED, {
    folderId: newFolder.getId(),
    folderName: folderName,
    type: 'ROOT',
    createdBy: Session.getActiveUser().getEmail()
  });

  return newFolder;
}

/**
 * Sets up a Drive folder for a specific grievance
 * Folder naming format: LastName, FirstName - YYYY-MM-DD
 * @param {string} grievanceId - The grievance ID
 * @return {Object} Result with folder URL or error
 */
function setupDriveFolderForGrievance(grievanceId) {
  try {
    // Get grievance data
    const grievance = getGrievanceById(grievanceId);
    if (!grievance) {
      return errorResponse('Grievance not found', 'setupDriveFolderForGrievance');
    }

    // Extract fields for folder naming
    const firstName = grievance['First Name'] || grievance.firstName || '';
    const lastName = grievance['Last Name'] || grievance.lastName || '';
    const dateFiled = grievance['Date Filed'] || grievance.dateFiled || new Date();

    // Format date as YYYY-MM-DD (full date the claim was initiated)
    const dateStr = Utilities.formatDate(
      new Date(dateFiled),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    // Create folder name from template
    // Simplified Format: LastName, FirstName - YYYY-MM-DD
    let folderName;
    if (firstName && lastName) {
      folderName = DRIVE_CONFIG.SUBFOLDER_TEMPLATE
        .replace('{date}', dateStr)
        .replace('{lastName}', sanitizeFolderName(lastName))
        .replace('{firstName}', sanitizeFolderName(firstName));
    } else {
      // Fallback if name not available: GrievanceID - Date
      folderName = DRIVE_CONFIG.SUBFOLDER_TEMPLATE_SIMPLE
        .replace('{date}', dateStr)
        .replace('{grievanceId}', grievanceId);
    }

    // Get root folder
    const rootFolder = getOrCreateRootFolder();

    // Check if folder already exists
    const existingFolders = rootFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      const existing = existingFolders.next();
      return {
        success: true,
        folderId: existing.getId(),
        folderUrl: existing.getUrl(),
        message: 'Folder already exists'
      };
    }

    // Create new folder
    const newFolder = rootFolder.createFolder(folderName);

    // Create standard subfolders
    newFolder.createFolder('Step 1 - Informal');
    newFolder.createFolder('Step 2 - Written');
    newFolder.createFolder('Step 3 - Review');
    newFolder.createFolder('Supporting Documents');

    // Share folder with assigned steward if available
    if (grievance && grievance.steward) {
      try {
        var stewardEmail = typeof grievance.stewardEmail === 'string' ? grievance.stewardEmail : '';
        if (stewardEmail) {
          newFolder.addEditor(stewardEmail);
        }
      } catch (shareError) {
        Logger.log('Could not share folder with steward: ' + shareError.message);
      }
    }

    // Update grievance record with folder link
    updateGrievanceFolderLink(grievanceId, newFolder.getUrl());

    logAuditEvent(AUDIT_EVENTS.FOLDER_CREATED, {
      grievanceId: grievanceId,
      folderId: newFolder.getId(),
      folderName: folderName,
      createdBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      folderId: newFolder.getId(),
      folderUrl: newFolder.getUrl(),
      message: 'Folder created successfully'
    };

  } catch (error) {
    console.error('Error creating Drive folder:', error);
    return errorResponse(error.message, 'setupDriveFolderForGrievance');
  }
}

/**
 * Creates Drive folders for multiple grievances in batches
 * @param {string[]} grievanceIds - Array of grievance IDs
 * @return {Object} Result with success count
 */
function batchCreateGrievanceFolders(grievanceIds) {
  const results = {
    success: 0,
    failed: 0,
    errors: []
  };

  const startTime = new Date().getTime();

  for (let i = 0; i < grievanceIds.length; i++) {
    // Check time limit
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
      results.errors.push('Time limit reached, some folders not created');
      break;
    }

    // Batch pause
    if (i > 0 && i % BATCH_LIMITS.MAX_API_CALLS_PER_BATCH === 0) {
      Utilities.sleep(BATCH_LIMITS.PAUSE_BETWEEN_BATCHES_MS);
    }

    const result = setupDriveFolderForGrievance(grievanceIds[i]);

    if (result.success) {
      results.success++;
    } else {
      results.failed++;
      results.errors.push(`${grievanceIds[i]}: ${result.error}`);
    }
  }

  return {
    success: true,
    created: results.success,
    failed: results.failed,
    errors: results.errors,
    message: `Created ${results.success} folders, ${results.failed} failed`
  };
}

/**
 * Menu wrapper: Setup Drive folder for the currently selected grievance
 * Gets the grievance ID from the active row in Grievance Log
 */
function setupFolderForSelectedGrievance() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Ensure sheet has enough columns for Drive folder columns
  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var range = ss.getActiveRange();
  var row = range.getRow();

  // Validate selection
  var validationError = null;
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    validationError = 'Please select a grievance row in the Grievance Log sheet first.';
  } else if (range.getRow() < 2) {
    validationError = 'Please select a data row (not the header).';
  } else if (!sheet.getRange(range.getRow(), GRIEVANCE_COLS.GRIEVANCE_ID).getValue()) {
    validationError = 'No Grievance ID found in the selected row.';
  }

  if (validationError) {
    ui.alert('📁 Setup Grievance Folder', validationError, ui.ButtonSet.OK);
    return;
  }

  var grievanceId = sheet.getRange(range.getRow(), GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  // Check if folder already exists
  var existingUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
  if (existingUrl) {
    var response = ui.alert('📁 Folder Already Exists',
      'A folder already exists for grievance ' + grievanceId + '.\n\n' +
      'Would you like to open it?',
      ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      var html = HtmlService.createHtmlOutput(
        '<script>window.open(' + JSON.stringify(existingUrl) + ', "_blank"); google.script.host.close();</script>'
      ).setWidth(1).setHeight(1);
      ui.showModalDialog(html, 'Opening folder...');
    }
    return;
  }

  // Create the folder
  ss.toast('Creating folder for ' + grievanceId + '...', '📁 Drive', 3);
  var result = setupDriveFolderForGrievance(grievanceId);

  if (result.success) {
    ui.alert('✅ Folder Created',
      'Folder created for grievance ' + grievanceId + '.\n\n' +
      'The folder URL has been saved to the Grievance Log.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Error',
      'Failed to create folder: ' + result.error,
      ui.ButtonSet.OK);
  }
}

/**
 * Menu wrapper: Batch create folders for all grievances missing folders
 */
function batchCreateAllMissingFolders() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    ui.alert('Error', 'Grievance Log not found.', ui.ButtonSet.OK);
    return;
  }

  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('📁 Batch Create Folders', 'No grievances found.', ui.ButtonSet.OK);
    return;
  }

  // Get grievance IDs and folder URLs
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(GRIEVANCE_COLS.GRIEVANCE_ID, GRIEVANCE_COLS.DRIVE_FOLDER_URL)).getValues();

  var missingFolders = [];
  for (var i = 0; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var folderUrl = data[i][GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1];

    if (grievanceId && !folderUrl) {
      missingFolders.push(grievanceId);
    }
  }

  if (missingFolders.length === 0) {
    ui.alert('📁 Batch Create Folders',
      'All grievances already have folders!',
      ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert('📁 Batch Create Folders',
    'Found ' + missingFolders.length + ' grievances without folders.\n\n' +
    'Create folders for all of them?\n\n' +
    'This may take a few minutes for large numbers.',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Creating ' + missingFolders.length + ' folders...', '📁 Drive', 10);

  var result = batchCreateGrievanceFolders(missingFolders);

  ui.alert('📁 Batch Create Complete',
    'Created: ' + result.created + ' folders\n' +
    'Failed: ' + result.failed + '\n\n' +
    (result.errors.length > 0 ? 'Errors:\n' + result.errors.slice(0, 5).join('\n') : ''),
    ui.ButtonSet.OK);
}

/**
 * Updates the Drive folder link in grievance record
 * @param {string} grievanceId - The grievance ID
 * @param {string} folderUrl - The folder URL
 */
function updateGrievanceFolderLink(grievanceId, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      sheet.getRange(i + 1, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folderUrl);
      break;
    }
  }
}

/**
 * Opens the Drive folder for selected grievance
 */
function openGrievanceFolder() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.GRIEVANCE_TRACKER) {
    showAlert('Please select a grievance in the Grievance Tracker', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) return;

  const folderUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();

  if (folderUrl) {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open(${JSON.stringify(folderUrl)}, '_blank'); google.script.host.close();</script>`
    ).setWidth(100).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening folder...');
  } else {
    if (showConfirmation('No folder exists. Create one now?', 'Create Folder')) {
      const grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
      const result = setupDriveFolderForGrievance(grievanceId);
      if (result.success) {
        const html = HtmlService.createHtmlOutput(
          `<script>window.open(${JSON.stringify(result.folderUrl)}, '_blank'); google.script.host.close();</script>`
        ).setWidth(100).setHeight(50);
        SpreadsheetApp.getUi().showModalDialog(html, 'Opening folder...');
      }
    }
  }
}

/**
 * Sanitizes a string for use as a folder name
 * @param {string} name - The name to sanitize
 * @return {string} Sanitized folder name
 */
function sanitizeFolderName(name) {
  if (!name) return 'Unknown';
  return name
    .trim()
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 50);
}

// ============================================================================
// MEETING CALENDAR INTEGRATION
// ============================================================================

/**
 * Gets or creates the union meetings calendar
 * @return {Calendar} The meetings calendar
 */
function getOrCreateMeetingsCalendar() {
  var calendarName = CALENDAR_CONFIG.MEETING_CALENDAR_NAME;
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  }
  var newCalendar = CalendarApp.createCalendar(calendarName, {
    summary: 'Union meeting events - Auto-managed by Union Dashboard',
    color: CalendarApp.Color.GREEN
  });
  return newCalendar;
}

/**
 * Creates a Google Calendar event for a meeting
 * @param {Object} meetingData - { name, date, time, duration, type }
 * @returns {string} Calendar event ID or empty string on failure
 */
function createMeetingCalendarEvent(meetingData) {
  try {
    var calendar = getOrCreateMeetingsCalendar();
    var meetingDate = new Date(meetingData.date + 'T00:00:00');
    // Note: Date parsing uses script timezone. For explicit control, consider
    // Utilities.formatDate() with Session.getScriptTimeZone().
    var startTime = meetingData.time || '09:00';
    var durationHours = parseFloat(meetingData.duration) || 1;
    var timeParts = startTime.split(':');
    var startHour = parseInt(timeParts[0], 10) || 9;
    var startMin = parseInt(timeParts[1], 10) || 0;

    var start = new Date(meetingDate);
    start.setHours(startHour, startMin, 0, 0);

    var end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

    var eventTitle = '[MTG] ' + meetingData.name + ' (' + (meetingData.type || 'In-Person') + ')';
    var event;
    try {
      event = calendar.createEvent(eventTitle, start, end, {
        description: 'Meeting ID: ' + (meetingData.meetingId || 'TBD') + '\n' +
                     'Type: ' + (meetingData.type || 'In-Person') + '\n' +
                     'Check-in opens on the day of the event.\n\n' +
                     'Auto-generated by Union Dashboard'
      });
    } catch (calErr) {
      // F31: Quota-aware error handling for calendar creation
      var msg = String(calErr.message || '');
      if (msg.indexOf('quota') !== -1 || msg.indexOf('limit') !== -1 || msg.indexOf('rate') !== -1) {
        Logger.log('Calendar quota exceeded, skipping event creation: ' + msg);
      } else {
        Logger.log('Error creating calendar event: ' + msg);
      }
      return '';
    }

    // Add email reminders
    event.removeAllReminders();
    event.addEmailReminder(24 * 60);  // 1 day before
    event.addEmailReminder(60);        // 1 hour before

    return event.getId();
  } catch (error) {
    Logger.log('Error creating meeting calendar event: ' + error.message);
    return '';
  }
}

/**
 * Deletes a meeting calendar event by ID
 * @param {string} eventId - Calendar event ID
 */
function deleteMeetingCalendarEvent(eventId) {
  if (!eventId) return;
  try {
    var calendar = getOrCreateMeetingsCalendar();
    var event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
    }
  } catch (error) {
    Logger.log('Error deleting meeting calendar event: ' + error.message);
  }
}

/**
 * Emails the attendance report for a meeting to specified stewards
 * @param {string} meetingId - Meeting ID to report on
 * @param {string} recipientEmails - Comma-separated steward emails
 * @returns {Object} { success, error }
 */
function emailMeetingAttendanceReport(meetingId, recipientEmails) {
  if (!meetingId || !recipientEmails) {
    return errorResponse('Meeting ID and recipient emails are required');
  }

  // Validate all recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',');
  for (var e = 0; e < emails.length; e++) {
    if (!emailRegex.test(emails[e].trim())) {
      return errorResponse('Invalid email address: ' + emails[e].trim());
    }
  }

  try {
    var result = getMeetingAttendees(meetingId);
    if (!result.success) {
      return errorResponse(result.error || 'Could not retrieve attendance data');
    }

    // Find meeting details from the check-in log
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() <= 1) {
      return errorResponse('No meeting check-in data found');
    }
    var data = sheet.getDataRange().getValues();
    var meetingName = '';
    var meetingDate = '';
    var meetingType = '';

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
        meetingName = data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '';
        meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
        meetingType = data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || '';
        break;
      }
    }

    var dateStr = meetingDate instanceof Date ? meetingDate.toLocaleDateString() : String(meetingDate);

    // Build email body
    var body = '<h2>Meeting Attendance Report</h2>' +
      '<table style="border-collapse:collapse;margin:10px 0">' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Meeting:</td><td style="padding:4px 12px">' + escapeHtml(meetingName) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Date:</td><td style="padding:4px 12px">' + escapeHtml(dateStr) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Type:</td><td style="padding:4px 12px">' + escapeHtml(meetingType) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Total Attendees:</td><td style="padding:4px 12px">' + escapeHtml(String(result.count)) + '</td></tr>' +
      '</table>';

    if (result.attendees.length > 0) {
      body += '<h3>Attendees</h3>' +
        '<table style="border-collapse:collapse;border:1px solid #ddd">' +
        '<tr style="background:#059669;color:white">' +
        '<th style="padding:8px;border:1px solid #ddd">#</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Member ID</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Name</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Check-In Time</th>' +
        '</tr>';

      for (var j = 0; j < result.attendees.length; j++) {
        var a = result.attendees[j];
        body += '<tr>' +
          '<td style="padding:6px;border:1px solid #ddd">' + (j + 1) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.memberId)) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.name)) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.time)) + '</td>' +
          '</tr>';
      }
      body += '</table>';
    } else {
      body += '<p><em>No members checked in to this meeting.</em></p>';
    }

    body += '<br><p style="font-size:12px;color:#666">Auto-generated by Union Dashboard</p>';

    var emailResult = safeSendEmail_({
      to: recipientEmails,
      subject: 'Meeting Attendance: ' + meetingName + ' (' + dateStr + ')',
      htmlBody: body,
      name: 'Union Dashboard'
    });
    if (!emailResult.success) return errorResponse(emailResult.error);

    return { success: true, message: 'Attendance report emailed to ' + recipientEmails };
  } catch (error) {
    Logger.log('Error emailing attendance report: ' + error.message);
    return errorResponse(error.message);
  }
}

// ============================================================================
// MEETING DOCS (Notes & Agenda) INTEGRATION
// ============================================================================

/**
 * Gets or creates the Meeting Notes folder under the root Drive folder
 * @return {Folder} The Meeting Notes folder
 */
function getOrCreateMeetingNotesFolder() {
  var rootFolder = getOrCreateRootFolder();
  var folderName = 'Meeting Notes';
  var folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var newFolder = rootFolder.createFolder(folderName);
  newFolder.setDescription('Meeting Notes - Auto-managed by Union Dashboard');
  return newFolder;
}

/**
 * Gets or creates the Meeting Agenda folder under the root Drive folder
 * @return {Folder} The Meeting Agenda folder
 */
function getOrCreateMeetingAgendaFolder() {
  var rootFolder = getOrCreateRootFolder();
  var folderName = 'Meeting Agenda';
  var folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var newFolder = rootFolder.createFolder(folderName);
  newFolder.setDescription('Meeting Agenda - Auto-managed by Union Dashboard (Steward access only)');
  return newFolder;
}

/**
 * Creates Meeting Notes and Meeting Agenda Google Docs for a meeting
 * Notes: shared view-only with members (day after meeting)
 * Agenda: shared with stewards only (3 days prior), NOT shared with members
 * @param {Object} meetingData - { meetingId, name, date }
 * @returns {Object} { notesUrl, agendaUrl }
 */
function createMeetingDocs(meetingData) {
  var result = { notesUrl: '', agendaUrl: '' };
  try {
    var dateStr = meetingData.date || '';
    var meetingName = meetingData.name || 'Meeting';

    // Create Meeting Notes doc
    var notesFolder = getOrCreateMeetingNotesFolder();
    var notesTitle = 'Meeting Notes - ' + meetingName + ' - ' + dateStr;
    var notesDoc = DocumentApp.create(notesTitle);
    var notesBody = notesDoc.getBody();
    notesBody.appendParagraph(notesTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    notesBody.appendParagraph('Meeting ID: ' + (meetingData.meetingId || ''));
    notesBody.appendParagraph('Date: ' + dateStr);
    notesBody.appendParagraph('');
    notesBody.appendParagraph('Notes:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    notesBody.appendParagraph('');
    notesDoc.saveAndClose();

    // Move to Meeting Notes folder
    var notesFile = DriveApp.getFileById(notesDoc.getId());
    notesFolder.addFile(notesFile);
    var parents = notesFile.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getId() !== notesFolder.getId()) {
        parent.removeFile(notesFile);
      }
    }
    result.notesUrl = notesDoc.getUrl();

    // Create Meeting Agenda doc
    var agendaFolder = getOrCreateMeetingAgendaFolder();
    var agendaTitle = 'Meeting Agenda - ' + meetingName + ' - ' + dateStr;
    var agendaDoc = DocumentApp.create(agendaTitle);
    var agendaBody = agendaDoc.getBody();
    agendaBody.appendParagraph(agendaTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    agendaBody.appendParagraph('Meeting ID: ' + (meetingData.meetingId || ''));
    agendaBody.appendParagraph('Date: ' + dateStr);
    agendaBody.appendParagraph('');
    agendaBody.appendParagraph('Agenda Items:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    agendaBody.appendParagraph('1. ');
    agendaBody.appendParagraph('2. ');
    agendaBody.appendParagraph('3. ');
    agendaBody.appendParagraph('');
    agendaBody.appendParagraph('Action Items:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    agendaBody.appendParagraph('');
    agendaDoc.saveAndClose();

    // Move to Meeting Agenda folder
    var agendaFile = DriveApp.getFileById(agendaDoc.getId());
    agendaFolder.addFile(agendaFile);
    var agendaParents = agendaFile.getParents();
    while (agendaParents.hasNext()) {
      var agendaParent = agendaParents.next();
      if (agendaParent.getId() !== agendaFolder.getId()) {
        agendaParent.removeFile(agendaFile);
      }
    }
    result.agendaUrl = agendaDoc.getUrl();

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MEETING_DOCS_CREATED', {
        meetingId: meetingData.meetingId,
        notesDocUrl: result.notesUrl,
        agendaDocUrl: result.agendaUrl
      });
    }
  } catch (error) {
    Logger.log('Error creating meeting docs: ' + error.message);
  }
  return result;
}

/**
 * Sets a Google Doc to view-only for anyone with the link
 * Used to make meeting notes viewable by members after the meeting
 * @param {string} docUrl - The Google Doc URL
 */
function setDocViewOnlyByLink(docUrl) {
  try {
    // Extract file ID from URL
    var match = docUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return;
    var fileId = match[1];
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    Logger.log('Error setting doc to view-only: ' + error.message);
  }
}

/**
 * Sends meeting document links to stewards via email
 * @param {string} meetingName - Name of the meeting
 * @param {string} meetingDate - Date of the meeting
 * @param {string} docUrl - The document URL to share
 * @param {string} docType - 'notes' or 'agenda'
 * @param {string} recipientEmails - Comma-separated steward emails
 */
function emailMeetingDocLink(meetingName, meetingDate, docUrl, docType, recipientEmails) {
  if (!recipientEmails || !docUrl) return;

  // Validate all recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',');
  for (var e = 0; e < emails.length; e++) {
    if (!emailRegex.test(emails[e].trim())) {
      Logger.log('Invalid email address in recipient list: ' + emails[e].trim());
      return;
    }
  }

  try {
    var typeLabel = docType === 'agenda' ? 'Meeting Agenda' : 'Meeting Notes';
    var safeDocUrl = /^https:\/\/docs\.google\.com\//.test(docUrl) ? docUrl : '';
    var body = '<h2>' + escapeHtml(typeLabel) + '</h2>' +
      '<p><strong>Meeting:</strong> ' + escapeHtml(meetingName) + '</p>' +
      '<p><strong>Date:</strong> ' + escapeHtml(meetingDate) + '</p>' +
      '<p>Click the link below to access the ' + escapeHtml(typeLabel.toLowerCase()) + ':</p>' +
      (safeDocUrl ? '<p><a href="' + escapeHtml(safeDocUrl) + '" style="background:#1a73e8;color:white;padding:10px 20px;border-radius:4px;text-decoration:none;display:inline-block">' +
      'Open ' + escapeHtml(typeLabel) + '</a></p>' : '<p><em>Document link unavailable.</em></p>') +
      '<br><p style="font-size:12px;color:#666">Auto-generated by Union Dashboard</p>';

    safeSendEmail_({
      to: recipientEmails,
      subject: typeLabel + ': ' + meetingName + ' (' + meetingDate + ')',
      htmlBody: body,
      name: 'Union Dashboard'
    });
  } catch (error) {
    Logger.log('Error emailing meeting doc link: ' + error.message);
  }
}

/**
 * Gets all steward emails from the Member Directory
 * @returns {string} Comma-separated steward emails
 */
function getAllStewardEmails_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet || sheet.getLastRow() < 2) return '';

    var data = sheet.getDataRange().getValues();
    var emails = [];
    for (var i = 1; i < data.length; i++) {
      var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isTruthyValue(isSteward)) {
        var email = String(data[i][MEMBER_COLS.EMAIL - 1] || '').trim();
        if (email) emails.push(email);
      }
    }
    return emails.join(', ');
  } catch (e) {
    Logger.log('Error getting all steward emails: ' + e.message);
    return '';
  }
}

/**
 * Sends scheduled meeting doc notifications from dailyTrigger
 * - Agenda link: 3 days before -> selected stewards (AGENDA_STEWARDS column)
 * - Agenda link: 1 day before -> ALL stewards (from Member Directory)
 * - Notes link: 1 day before -> attendance notification stewards (NOTIFY_STEWARDS column)
 * - Sets notes to view-only: 1 day after meeting (for members)
 * @returns {Object} { agendaSent, notesSent, notesPublished }
 */
function processMeetingDocNotifications() {
  var result = { agendaSent: 0, notesSent: 0, notesPublished: 0 };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() < 2) return result;

    var data = sheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var processed = {};  // Track processed meeting IDs
    var allStewardEmails = null;  // Lazy-loaded

    for (var i = 1; i < data.length; i++) {
      var meetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      if (!meetingId || processed[meetingId]) continue;
      processed[meetingId] = true;

      var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
      if (!(meetingDate instanceof Date)) continue;

      var meetingDay = new Date(meetingDate);
      meetingDay.setHours(0, 0, 0, 0);
      var diffDays = Math.round((meetingDay.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));

      var meetingName = String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '');
      var notifyEmails = String(data[i][MEETING_CHECKIN_COLS.NOTIFY_STEWARDS - 1] || '');
      var notesUrl = String(data[i][MEETING_CHECKIN_COLS.NOTES_DOC_URL - 1] || '');
      var agendaUrl = String(data[i][MEETING_CHECKIN_COLS.AGENDA_DOC_URL - 1] || '');
      var agendaStewards = String(data[i][MEETING_CHECKIN_COLS.AGENDA_STEWARDS - 1] || '');
      var dateStr = meetingDate.toLocaleDateString();

      // 3 days before: send agenda link to SELECTED stewards only
      if (diffDays === 3 && agendaUrl && agendaStewards) {
        emailMeetingDocLink(meetingName, dateStr, agendaUrl, 'agenda', agendaStewards);
        result.agendaSent++;
      }

      // 1 day before: send agenda link to ALL stewards, notes link to notify stewards
      if (diffDays === 1) {
        if (agendaUrl) {
          // Lazy-load all steward emails only when needed
          if (allStewardEmails === null) {
            allStewardEmails = getAllStewardEmails_();
          }
          if (allStewardEmails) {
            emailMeetingDocLink(meetingName, dateStr, agendaUrl, 'agenda', allStewardEmails);
            result.agendaSent++;
          }
        }
        if (notesUrl && notifyEmails) {
          emailMeetingDocLink(meetingName, dateStr, notesUrl, 'notes', notifyEmails);
          result.notesSent++;
        }
      }

      // 1 day after: set notes to view-only (available to members)
      if (diffDays === -1 && notesUrl) {
        setDocViewOnlyByLink(notesUrl);
        result.notesPublished++;
      }
    }
  } catch (error) {
    Logger.log('Error processing meeting doc notifications: ' + error.message);
  }
  return result;
}

// ============================================================================
// MEMBER DRIVE FOLDER
// ============================================================================

/**
 * Creates or retrieves a Google Drive folder for a member
 * If the member has an existing grievance with a folder, reuses that folder
 * Otherwise creates a new folder under the root folder
 * @param {string} memberId - The member ID
 * @returns {Object} { success, folderUrl, folderId, message, error }
 */
function setupDriveFolderForMember(memberId) {
  try {
    if (!memberId) {
      return errorResponse('Member ID is required');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet) return errorResponse('Member Directory not found');

    // Find member row
    var memberData = memberSheet.getDataRange().getValues();
    var memberRow = -1;
    var firstName = '';
    var lastName = '';
    for (var i = 1; i < memberData.length; i++) {
      if (String(memberData[i][MEMBER_COLS.MEMBER_ID - 1]) === String(memberId)) {
        memberRow = i + 1;
        firstName = memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '';
        lastName = memberData[i][MEMBER_COLS.LAST_NAME - 1] || '';
        break;
      }
    }
    if (memberRow === -1) {
      return errorResponse('Member not found: ' + memberId);
    }

    // Check if member has an existing grievance with a Drive folder
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      var gData = grievanceSheet.getDataRange().getValues();
      for (var g = 1; g < gData.length; g++) {
        if (String(gData[g][GRIEVANCE_COLS.MEMBER_ID - 1]) === String(memberId)) {
          var existingUrl = gData[g][GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1];
          if (existingUrl) {
            return {
              success: true,
              folderUrl: existingUrl,
              message: 'Using existing grievance folder for this member'
            };
          }
        }
      }
    }

    // No existing folder found - create a new one
    var rootFolder = getOrCreateRootFolder();
    var folderName = sanitizeFolderName(lastName) + ', ' + sanitizeFolderName(firstName) + ' - ' + memberId;

    // Check if folder already exists
    var existingFolders = rootFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      var existing = existingFolders.next();
      return {
        success: true,
        folderId: existing.getId(),
        folderUrl: existing.getUrl(),
        message: 'Folder already exists'
      };
    }

    var newFolder = rootFolder.createFolder(folderName);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MEMBER_FOLDER_CREATED', {
        memberId: memberId,
        folderId: newFolder.getId(),
        folderName: folderName
      });
    }

    return {
      success: true,
      folderId: newFolder.getId(),
      folderUrl: newFolder.getUrl(),
      message: 'Folder created successfully'
    };
  } catch (error) {
    Logger.log('Error creating member Drive folder: ' + error.message);
    return errorResponse(error.message);
  }
}

// ============================================================================
// GRIEVANCE CALENDAR INTEGRATION
// ============================================================================

/**
 * Gets or creates the grievance deadlines calendar
 * @return {Calendar} The deadlines calendar
 */
function getOrCreateDeadlinesCalendar() {
  const calendarName = CALENDAR_CONFIG.CALENDAR_NAME;

  // Check owned calendars
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  }

  // Create new calendar
  const newCalendar = CalendarApp.createCalendar(calendarName, {
    summary: 'Grievance deadline tracking - Auto-managed by Union Dashboard',
    color: CalendarApp.Color.RED
  });

  return newCalendar;
}

/**
 * Syncs all grievance deadlines to calendar
 * @return {Object} Result with sync count
 */
function syncDeadlinesToCalendar() {
  try {
    const calendar = getOrCreateDeadlinesCalendar();
    const openGrievances = getOpenGrievances();

    let synced = 0;
    let skipped = 0;
    const startTime = new Date().getTime();

    for (const grievance of openGrievances) {
      // Check time limit
      if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
        break;
      }

      const result = syncGrievanceDeadlinesToCalendar(
        grievance,
        calendar
      );

      if (result.synced) {
        synced++;
      } else {
        skipped++;
      }

      // Rate limiting pause
      if (synced % 20 === 0) {
        Utilities.sleep(200);
      }
    }

    logAuditEvent(AUDIT_EVENTS.CALENDAR_SYNCED, {
      synced: synced,
      skipped: skipped,
      syncedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      synced: synced,
      skipped: skipped,
      message: `Synced ${synced} grievances to calendar`
    };

  } catch (error) {
    console.error('Error syncing to calendar:', error);
    return errorResponse(error.message);
  }
}

/**
 * Syncs a single grievance's deadlines to calendar
 * @param {Object} grievance - Grievance data object
 * @param {Calendar} calendar - Target calendar
 * @return {Object} Sync result
 */
function syncGrievanceDeadlinesToCalendar(grievance, calendar) {
  const grievanceId = grievance['Grievance ID'];
  const memberName = grievance['Member Name'];
  const currentStep = grievance['Current Step'];

  // Get the deadline for current step
  let deadline;
  switch (currentStep) {
    case 'Step I':
    case 'Informal':
      deadline = grievance['Step 1 Due'];
      break;
    case 'Step II':
      deadline = grievance['Step 2 Due'];
      break;
    case 'Step III':
    case 'Arbitration':
      deadline = grievance['Step 3 Due'];
      break;
    default:
      return { synced: false, reason: 'No applicable deadline' };
  }

  if (!(deadline instanceof Date)) {
    return { synced: false, reason: 'Invalid deadline date' };
  }

  // Check if deadline is in the past
  if (deadline < new Date()) {
    return { synced: false, reason: 'Deadline already passed' };
  }

  // Create event title
  const eventTitle = `[GRV] ${grievanceId} - Step ${currentStep} Due (${memberName})`;

  // Check for existing event to avoid duplicates
  const existingEvents = calendar.getEventsForDay(deadline, {
    search: grievanceId
  });

  if (existingEvents.length > 0) {
    // Update existing event
    const event = existingEvents[0];
    event.setTitle(eventTitle);
    return { synced: true, updated: true };
  }

  // Create all-day event
  const event = calendar.createAllDayEvent(eventTitle, deadline, {
    description: `Grievance: ${grievanceId}\n` +
                 `Member: ${memberName}\n` +
                 `Step: ${currentStep}\n` +
                 `Action Required: Response deadline\n\n` +
                 `Auto-generated by Union Dashboard`
  });

  // Set reminders
  event.removeAllReminders();
  CALENDAR_CONFIG.REMINDER_DAYS.forEach(days => {
    event.addEmailReminder(days * 24 * 60); // Convert days to minutes
  });

  return { synced: true, created: true };
}

// Note: syncSingleGrievanceToCalendar() is defined in MobileQuickActions.gs

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Sends deadline reminder email for upcoming grievances
 * @param {number} daysAhead - Days to look ahead
 * @return {Object} Result object
 */
function sendDeadlineReminders(daysAhead) {
  try {
    const deadlines = getUpcomingDeadlines(daysAhead || 7);
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(userEmail)) {
      return errorResponse('Could not determine a valid email address for the current user');
    }

    if (deadlines.length === 0) {
      return { success: true, sent: false, message: 'No upcoming deadlines' };
    }

    // Build email body
    let body = `<h2>Upcoming Grievance Deadlines</h2>`;
    body += `<p>The following grievances have deadlines in the next ${daysAhead} days:</p>`;
    body += `<table style="border-collapse: collapse; width: 100%;">`;
    body += `<tr style="background: #f0f0f0;">
               <th style="padding: 10px; border: 1px solid #ddd;">Grievance</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Member</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Step</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Due Date</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Days Left</th>
             </tr>`;

    deadlines.forEach(d => {
      const urgent = d.daysLeft <= 3 ? 'style="background: #fee2e2;"' : '';
      body += `<tr ${urgent}>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.grievanceId)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.memberName)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.step)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.date)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(String(d.daysLeft))}</td>
               </tr>`;
    });

    body += `</table>`;
    body += `<p style="margin-top: 20px; color: #666; font-size: 12px;">
               This is an automated reminder from the Union Dashboard.
             </p>`;

    // Send email
    safeSendEmail_({
      to: userEmail,
      subject: `[Union Dashboard] ${deadlines.length} Upcoming Grievance Deadline${deadlines.length > 1 ? 's' : ''}`,
      htmlBody: body
    });

    return {
      success: true,
      sent: true,
      count: deadlines.length,
      message: `Sent reminder for ${deadlines.length} deadlines`
    };

  } catch (error) {
    console.error('Error sending reminders:', error);
    return errorResponse(error.message);
  }
}

/**
 * Sends email to a member
 * @param {string} memberId - The member ID
 * @param {string} subject - Email subject
 * @param {string} body - Email body (HTML)
 * @return {Object} Result object
 */
function sendEmailToMember(memberId, subject, body) {
  try {
    const member = getMemberById(memberId);
    if (!member) {
      return errorResponse('Member not found');
    }

    const email = member['Email'] || member.email;
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim())) {
      return errorResponse('Invalid email address');
    }

    safeSendEmail_({
      to: email,
      subject: subject,
      htmlBody: body
    });

    return {
      success: true,
      message: `Email sent to ${email}`
    };

  } catch (error) {
    console.error('Error sending email:', error);
    return errorResponse(error.message);
  }
}

// ============================================================================
// PDF SIGNATURE ENGINE (Strategic Command Center)
// ============================================================================

/**
 * Gets or creates an archive folder for a specific member
 * Used for storing grievance PDFs and documents
 * @param {string} name - Member's name
 * @param {string} id - Member's ID
 * @returns {Folder} The member's archive folder
 */
function getOrCreateMemberFolder(name, id) {
  // Get archive folder ID from Config or COMMAND_CONFIG
  var archiveFolderId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID) || COMMAND_CONFIG.ARCHIVE_FOLDER_ID;

  if (!archiveFolderId) {
    // Fall back to creating in root folder
    var rootFolder = getOrCreateRootFolder();
    var folderName = name + ' (' + id + ')';
    var folders = rootFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
  }

  try {
    var parentFolder = DriveApp.getFolderById(archiveFolderId);
    var folderName = name + ' (' + id + ')';
    var folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  } catch (e) {
    Logger.log('Archive folder not found, using root: ' + e.message);
    var rootFolder = getOrCreateRootFolder();
    var folderName = name + ' (' + id + ')';
    var folders = rootFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
  }
}

/**
 * Creates a signature-ready PDF from a grievance template
 * Merges data and adds signature blocks
 * @param {Folder} folder - Target folder for the PDF
 * @param {Object} data - Grievance data object
 * @returns {File} The created PDF file
 */
function createSignatureReadyPDF(folder, data) {
  // Get template ID from Config or COMMAND_CONFIG
  var templateId = getConfigValue_(CONFIG_COLS.TEMPLATE_ID) || COMMAND_CONFIG.TEMPLATE_ID;

  if (!templateId) {
    throw new Error('Document template ID not configured. Please set TEMPLATE_ID in Config sheet.');
  }

  try {
    // Copy template
    var temp = DriveApp.getFileById(templateId).makeCopy('SIGNATURE_REQUIRED_' + data.name, folder);
    var doc = DocumentApp.openById(temp.getId());
    var body = doc.getBody();

    // Replace placeholders with data
    body.replaceText('{{MemberName}}', data.name || 'Unknown');
    body.replaceText('{{MemberID}}', data.id || '000');
    body.replaceText('{{Date}}', new Date().toLocaleDateString());
    body.replaceText('{{Details}}', data.details || 'No details provided.');
    body.replaceText('{{GrievanceID}}', data.grievanceId || '');
    body.replaceText('{{Articles}}', data.articles || '');
    body.replaceText('{{Status}}', data.status || '');
    body.replaceText('{{Unit}}', data.unit || '');
    body.replaceText('{{Location}}', data.location || '');
    body.replaceText('{{Steward}}', data.steward || '');

    // Append legal signature block
    body.appendParagraph(COMMAND_CONFIG.PDF.SIGNATURE_BLOCK ||
      '\n\n__________________________\nMember Signature\n\n__________________________\nSteward Signature\n\n__________________________\nDate');

    doc.saveAndClose();

    // Convert to PDF
    var pdf = folder.createFile(temp.getAs(MimeType.PDF))
                    .setName('Grievance_UNSIGNED_' + data.name + '_' + new Date().toISOString().slice(0,10) + '.pdf');

    // Remove the temp document
    temp.setTrashed(true);

    return pdf;

  } catch (e) {
    Logger.log('Error creating signature PDF: ' + e.message);
    throw new Error('Failed to create PDF: ' + e.message);
  }
}

/**
 * Creates a PDF for the currently selected grievance
 * Saves to member's Drive folder and optionally emails to member
 * Accessible from the Command menu
 */
function createPDFForSelectedGrievance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('Please select a grievance row in the Grievance Log sheet');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a grievance row (not the header)');
    return;
  }

  // Get grievance data including member email
  var data = {
    grievanceId: sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue(),
    name: sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
          sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue(),
    id: sheet.getRange(row, GRIEVANCE_COLS.MEMBER_ID).getValue(),
    status: sheet.getRange(row, GRIEVANCE_COLS.STATUS).getValue(),
    articles: sheet.getRange(row, GRIEVANCE_COLS.ARTICLES).getValue(),
    details: sheet.getRange(row, GRIEVANCE_COLS.RESOLUTION).getValue() || 'Pending',
    location: sheet.getRange(row, GRIEVANCE_COLS.LOCATION).getValue(),
    steward: sheet.getRange(row, GRIEVANCE_COLS.STEWARD).getValue(),
    memberEmail: sheet.getRange(row, GRIEVANCE_COLS.MEMBER_EMAIL).getValue()
  };

  var response = ui.alert(
    'Create Signature PDF',
    'Create a signature-ready PDF for grievance ' + data.grievanceId + '?\n\n' +
    'Member: ' + data.name + '\n' +
    'Status: ' + data.status + '\n' +
    'Email: ' + (data.memberEmail || 'Not on file'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    ss.toast('Creating PDF...', COMMAND_CONFIG.SYSTEM_NAME, 5);

    // Get or create member folder
    var folder = getOrCreateMemberFolder(data.name, data.id);

    // Create the PDF
    var pdf = createSignatureReadyPDF(folder, data);

    // Update grievance record with PDF link and folder URL
    if (GRIEVANCE_COLS.DRIVE_FOLDER_URL) {
      sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folder.getUrl());
    }

    // Ask if user wants to email the PDF to member
    var emailResponse = ui.alert(
      'Email PDF to Member?',
      'PDF created successfully!\n\n' +
      'File: ' + pdf.getName() + '\n' +
      'Saved to: ' + folder.getName() + '\n\n' +
      (data.memberEmail
        ? 'Would you like to email this PDF to ' + data.memberEmail + '?'
        : 'No email on file for this member. Add email to column X to enable this feature.'),
      data.memberEmail ? ui.ButtonSet.YES_NO : ui.ButtonSet.OK
    );

    if (emailResponse === ui.Button.YES && data.memberEmail) {
      sendGrievancePdfEmail_(data, pdf);
      ss.toast('PDF emailed to ' + data.memberEmail, COMMAND_CONFIG.SYSTEM_NAME, 5);
    }

    // Open folder in new tab
    var html = HtmlService.createHtmlOutput(
      '<script>window.open(' + JSON.stringify(folder.getUrl()) + ', "_blank"); google.script.host.close();</script>'
    ).setWidth(100).setHeight(50);
    ui.showModalDialog(html, 'Opening folder...');

  } catch (e) {
    ui.alert('Error', 'Failed to create PDF: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Sends grievance PDF to member via email
 * @param {Object} data - Grievance data object with memberEmail
 * @param {File} pdf - The PDF file to attach
 * @private
 */
function sendGrievancePdfEmail_(data, pdf) {
  if (!data.memberEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.memberEmail).trim())) {
    throw new Error('Invalid or missing member email address');
  }

  var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Grievance Form - ' + data.grievanceId;

  var body = 'Dear ' + data.name + ',\n\n' +
    'Please find attached your grievance form for case ' + data.grievanceId + '.\n\n' +
    'GRIEVANCE DETAILS:\n' +
    '─────────────────────────────────\n' +
    'Grievance ID: ' + data.grievanceId + '\n' +
    'Status: ' + data.status + '\n' +
    'Articles: ' + (data.articles || 'N/A') + '\n' +
    'Unit: ' + (data.unit || 'N/A') + '\n' +
    'Location: ' + (data.location || 'N/A') + '\n' +
    'Assigned Steward: ' + (data.steward || 'N/A') + '\n' +
    '─────────────────────────────────\n\n' +
    'NEXT STEPS:\n' +
    '1. Review the attached form for accuracy\n' +
    '2. Sign where indicated\n' +
    '3. Return the signed form to your steward\n\n' +
    'If you have any questions, please contact your steward.\n' +
    COMMAND_CONFIG.EMAIL.FOOTER;

  safeSendEmail_({
    to: data.memberEmail,
    subject: subject,
    body: body,
    attachments: [pdf.getAs(MimeType.PDF)],
    name: COMMAND_CONFIG.SYSTEM_NAME || 'Union Grievance System'
  });

  // Use secureLog to mask PII in logs
  if (typeof secureLog === 'function') {
    secureLog('EmailPDF', 'Grievance PDF emailed', { recipientMasked: typeof maskEmail === 'function' ? maskEmail(data.memberEmail) : '[REDACTED]' });
  }
}

/**
 * Handles form submission for grievance intake forms
 * Creates PDF and links it back to the log
 * @param {Object} e - Form submission event object
 */
function onGrievanceFormSubmit(e) {
  try {
    var responses = e.namedValues;
    var data = {
      name: responses['Member Name'] ? responses['Member Name'][0] : 'Unknown',
      id: responses['Member ID'] ? responses['Member ID'][0] : '000',
      details: responses['Details'] ? responses['Details'][0] : 'No details provided.',
      grievanceId: responses['Grievance ID'] ? responses['Grievance ID'][0] : '',
      articles: responses['Articles'] ? responses['Articles'][0] : ''
    };

    // Create member folder and PDF
    var memberFolder = getOrCreateMemberFolder(data.name, data.id);
    var pdfFile = createSignatureReadyPDF(memberFolder, data);

    // Link PDF back to the grievance log
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (sheet) {
      // Use the event range to find the exact row the form submission added to
      var targetRow = e.range ? e.range.getRow() : sheet.getLastRow();
      if (GRIEVANCE_COLS.DRIVE_FOLDER_URL) {
        sheet.getRange(targetRow, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(pdfFile.getUrl());
      }
    }

    // Use secureLog to avoid logging PII
    if (typeof secureLog === 'function') {
      secureLog('FormSubmission', 'Form submission processed - PDF created', {});
    }

  } catch (err) {
    Logger.log('Error processing form submission: ' + err.message);
  }
}

// ============================================================================
// UI DIALOGS FOR INTEGRATIONS
// ============================================================================

/**
 * Shows calendar sync dialog
 */
function showCalendarSyncDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .sync-container { padding: 20px; }
        .sync-option { margin: 15px 0; padding: 15px; border: 1px solid #ddd; border-radius: 8px; }
        .sync-option h4 { margin: 0 0 8px 0; }
        .sync-option p { margin: 0; color: #666; font-size: 13px; }
        .action-buttons { margin-top: 20px; display: flex; gap: 10px; justify-content: flex-end; }
        .status { margin-top: 15px; padding: 10px; background: #f0f0f0; border-radius: 4px; display: none; }
      </style>
    </head>
    <body>
      <div class="sync-container">
        <div class="sync-option">
          <h4>Sync All Open Grievances</h4>
          <p>Creates calendar events for all deadlines of open grievances</p>
          <button class="btn btn-primary" onclick="syncAll()" style="margin-top: 10px;">
            Sync All Deadlines
          </button>
        </div>

        <div id="status" class="status"></div>

        <div class="action-buttons">
          <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
        </div>
      </div>

      <script>
        function showStatus(msg, isError) {
          const el = document.getElementById('status');
          el.style.display = 'block';
          el.style.background = isError ? '#fee2e2' : '#d1fae5';
          el.textContent = msg;
        }

        function syncAll() {
          showStatus('Syncing...', false);
          google.script.run
            .withSuccessHandler(function(r) {
              showStatus(r.success ? r.message : 'Error: ' + r.error, !r.success);
            })
            .syncDeadlinesToCalendar();
        }
      </script>
    </body>
    </html>
  `).setWidth(450).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Calendar Sync');
}

/**
 * Shows upcoming deadlines dialog
 */
function showUpcomingDeadlines() {
  const deadlines = getUpcomingDeadlines(14);

  let tableRows = '';
  if (deadlines.length === 0) {
    tableRows = '<tr><td colspan="4" style="text-align:center; padding:20px;">No upcoming deadlines</td></tr>';
  } else {
    deadlines.forEach(d => {
      const urgentStyle = d.daysLeft <= 3 ? 'background:#fee2e2;' : '';
      tableRows += `<tr style="${urgentStyle}">
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.grievanceId))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.memberName))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.step))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.date))} (${escapeHtml(String(d.daysLeft))} days)</td>
      </tr>`;
    });
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
    </head>
    <body style="padding: 20px;">
      <h3 style="margin-bottom: 15px;">Upcoming Deadlines (Next 14 Days)</h3>
      <table style="width:100%; border-collapse: collapse;">
        <tr style="background:#f0f0f0;">
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Grievance</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Member</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Step</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Due</th>
        </tr>
        ${tableRows}
      </table>
      <div style="margin-top:20px; text-align:right;">
        <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
      </div>
    </body>
    </html>
  `).setWidth(600).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Upcoming Deadlines');
}
/**
 * ============================================================================
 * WEB APP DEPLOYMENT FOR MOBILE ACCESS
 * ============================================================================
 * This file enables the dashboard to be deployed as a standalone web app
 * that can be accessed directly via URL on mobile devices.
 *
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Go to Extensions → Apps Script
 * 2. Click "Deploy" → "New deployment"
 * 3. Select "Web app" as the deployment type
 * 4. Set "Execute as" to your account
 * 5. Set "Who has access" to your organization or anyone
 * 6. Click "Deploy" and copy the URL
 * 7. Bookmark this URL on your mobile device for easy access
 */

/**
 * Web app entry point - serves the mobile dashboard and member portal
 * Consolidated to handle both mobile dashboard pages and member portal requests
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} The HTML page to display
 *
 * URL Parameters:
 * - id=<memberId> - Returns personalized member portal
 * - page=search|grievances|members|links|dashboard|portal - Returns specific page
 * - (no params) - Returns default dashboard
 */
function doGetLegacy(e) {
  // v4.5.0: Add access control and input validation
  // NOTE: Renamed from doGet → doGetLegacy. The web-dashboard WebApp.gs now owns doGet().

  // Step 1: Validate request parameters (prevents injection attacks)
  var validation = validateWebAppRequest(e);
  if (!validation.isValid) {
    secureLog('doGet', 'Invalid request parameters', { errors: validation.errors });
    if (typeof recordSecurityEvent === 'function') {
      recordSecurityEvent('INVALID_REQUEST', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.MEDIUM : 'MEDIUM',
        'Invalid web app request parameters detected',
        { errors: validation.errors });
    }
    return getAccessDeniedPage('Invalid request: ' + validation.errors.join(', '));
  }

  // Step 2: Check for unified dashboard mode parameter
  var mode = validation.params.mode || (e && e.parameter && e.parameter.mode) || '';
  if (mode === 'steward' || mode === 'member') {
    // v4.5.2: isPII is determined by auth result, NOT the URL parameter.
    // Default to false; only grant PII access after confirmed steward/admin authorization.
    var isPII = false;

    if (mode === 'steward') {
      // Steward mode requires authorization (contains PII)
      var authResult = checkWebAppAuthorization('steward');
      if (!authResult.isAuthorized) {
        secureLog('doGet', 'Unauthorized steward access attempt', {
          email: authResult.email,
          role: authResult.role
        });
        if (typeof recordSecurityEvent === 'function') {
          recordSecurityEvent('UNAUTHORIZED_ACCESS', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.HIGH : 'HIGH',
            'Unauthorized steward mode access attempt',
            { email: authResult.email, role: authResult.role, page: 'steward' });
        }
        return getAccessDeniedPage(authResult.message || 'Steward access required');
      }
      // PII access granted ONLY after successful authorization check
      isPII = true;
      secureLog('doGet', 'Steward access granted', { email: authResult.email });
    }

    var title = isPII ? 'STEWARD COMMAND CENTER' : 'MEMBER DASHBOARD';
    var html = getUnifiedDashboardHtml(isPII);
    return HtmlService.createHtmlOutput(html)
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // Step 3: (Removed in v4.5.2) The ?id= member portal route has been removed.
  // Members access their portal via ?page=selfservice using either Google account
  // verification or PIN authentication. This eliminates IDOR risk from URL-based member IDs.

  // Step 4: Standard page routing for mobile dashboard (legacy support)
  var page = validation.params.page || (e && e.parameter && e.parameter.page) || 'dashboard';

  // v4.5.2: Check if dashboard member authentication is required (toggled via Admin menu)
  // When enabled, redirect to selfservice portal for public dashboard pages
  if (typeof isDashboardMemberAuthRequired === 'function' && isDashboardMemberAuthRequired()) {
    var publicPages = ['dashboard', 'portal'];
    if (publicPages.indexOf(page) >= 0 || !page) {
      secureLog('doGet', 'Dashboard auth required - checking member auth', { requestedPage: page });

      // Try Google account verification first
      try {
        var dashAuthUser = Session.getEffectiveUser();
        var dashAuthEmail = dashAuthUser ? dashAuthUser.getEmail() : null;
        if (dashAuthEmail && typeof getMemberIdByEmail === 'function') {
          var dashLinkedId = getMemberIdByEmail(dashAuthEmail);
          if (dashLinkedId && typeof buildMemberPortal === 'function') {
            secureLog('doGet', 'Dashboard auth - member verified via Google account', { email: dashAuthEmail });
            return buildMemberPortal(dashLinkedId);
          }
        }
      } catch (_dashGoogleErr) {
        // Google auth not available - fall through to PIN login
      }

      // No linked Google account - show PIN login form
      if (typeof getMemberSelfServicePortalHtml === 'function') {
        var authHtml = getMemberSelfServicePortalHtml();
        return HtmlService.createHtmlOutput(authHtml)
          .setTitle('Member Login')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
      }
    }
  }

  // v4.5.1: Add authorization for sensitive pages (search, grievances, members)
  var sensitivePages = ['search', 'grievances', 'members'];
  if (sensitivePages.indexOf(page) >= 0) {
    var pageAuthResult = checkWebAppAuthorization('steward');
    if (!pageAuthResult.isAuthorized) {
      secureLog('doGet', 'Unauthorized access to ' + page + ' page', {
        email: pageAuthResult.email,
        role: pageAuthResult.role
      });
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('UNAUTHORIZED_ACCESS', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.HIGH : 'HIGH',
          'Unauthorized access to ' + page + ' page',
          { email: pageAuthResult.email, role: pageAuthResult.role, page: page });
      }
      return getAccessDeniedPage(pageAuthResult.message || 'Steward authorization required to view ' + page);
    }
    secureLog('doGet', 'Authorized access to ' + page + ' page', { email: pageAuthResult.email });
  }

  var html;
  switch (page) {
    case 'search':
      html = getWebAppSearchHtml();
      break;
    case 'grievances':
      html = getWebAppGrievanceListHtml();
      break;
    case 'members':
      html = getWebAppMemberListHtml();
      break;
    case 'links':
      html = getWebAppLinksHtml();
      break;
    case 'selfservice':
      // v4.5.2: Try Google account verification first, fall back to PIN
      try {
        var selfServiceUser = Session.getEffectiveUser();
        var selfServiceEmail = selfServiceUser ? selfServiceUser.getEmail() : null;
        if (selfServiceEmail && typeof getMemberIdByEmail === 'function') {
          var linkedMemberId = getMemberIdByEmail(selfServiceEmail);
          if (linkedMemberId && typeof buildMemberPortal === 'function') {
            secureLog('doGet', 'Member authenticated via Google account', { email: selfServiceEmail });
            return buildMemberPortal(linkedMemberId);
          }
        }
      } catch (_googleAuthErr) {
        // Google auth not available - fall through to PIN login
      }
      // No linked Google account found - show PIN login form
      if (typeof getMemberSelfServicePortalHtml === 'function') {
        html = getMemberSelfServicePortalHtml();
      } else {
        return getAccessDeniedPage('Member self-service portal not available');
      }
      break;
    case 'workload':
      // Workload Tracker portal — member PIN auth handled client-side
      if (typeof getWorkloadTrackerPortalHtml === 'function') {
        return HtmlService.createHtmlOutput(getWorkloadTrackerPortalHtml())
          .setTitle('Workload Tracker | SEIU 509')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
      }
      html = getUnifiedDashboardHtml(false);
      break;
    case 'checkin':
      // v4.11.0: Meeting check-in as standalone web page (mobile-friendly kiosk)
      html = getWebAppCheckInHtml();
      break;
    case 'resources':
      // v4.12.2: Route through SPA (educational hub with search, category pills)
      if (typeof doGetWebDashboard === 'function') {
        return doGetWebDashboard(e);
      }
      html = getWebAppResourcesHtml();
      break;
    case 'notifications':
      // v4.12.2: Route through SPA (hero headers, type badges, compose form)
      if (typeof doGetWebDashboard === 'function') {
        return doGetWebDashboard(e);
      }
      html = getWebAppNotificationsHtml();
      break;
    case 'portal':
      // Public portal without member ID
      if (typeof buildPublicPortal === 'function') {
        return buildPublicPortal();
      }
      // Fall through to SPA dashboard
    case 'dashboard':
    default:
      // v4.12.2: Default to SPA web dashboard (SSO + magic link auth)
      if (typeof doGetWebDashboard === 'function') {
        return doGetWebDashboard(e);
      }
      html = getUnifiedDashboardHtml(false);
      break;
  }

  if (!html) {
    return getAccessDeniedPage('Unable to load the requested page');
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle('Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Returns the main dashboard HTML for web app (enhanced with clickable stats, win rate, overdue preview)
 */
function getWebAppDashboardHtml() {
  var stats = getWebAppDashboardStats();
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">' +
    '<link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>📊</text></svg>">' +
    '<title>Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h1{font-size:clamp(20px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Stats grid - clickable cards
    '.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:15px 10px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;cursor:pointer;transition:transform 0.2s;text-decoration:none;display:block}' +
    '.stat-card:active{transform:scale(0.96)}' +
    '.stat-value{font-size:clamp(22px,6vw,32px);font-weight:bold;color:#7C3AED}' +
    '.stat-value.warning{color:#F97316}' +
    '.stat-value.danger{color:#DC2626}' +
    '.stat-value.success{color:#059669}' +
    '.stat-label{font-size:clamp(9px,2.2vw,11px);color:#666;text-transform:uppercase;margin-top:4px;letter-spacing:0.3px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Overdue preview
    '.overdue-section{background:#FEF2F2;border-left:4px solid #DC2626;border-radius:12px;padding:15px;margin-bottom:20px}' +
    '.overdue-title{font-size:14px;font-weight:600;color:#DC2626;margin-bottom:10px;display:flex;align-items:center;gap:6px}' +
    '.overdue-item{background:white;padding:12px;border-radius:10px;margin-bottom:8px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}' +
    '.overdue-item:last-child{margin-bottom:0}' +
    '.overdue-id{font-weight:600;color:#7C3AED;font-size:13px}' +
    '.overdue-name{font-size:14px;color:#333;margin-top:2px}' +
    '.overdue-detail{font-size:12px;color:#666;margin-top:2px}' +
    '.view-all-btn{background:#DC2626;color:white;border:none;padding:10px;border-radius:8px;font-size:13px;font-weight:500;width:100%;margin-top:10px;cursor:pointer}' +

    // Action buttons
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{background:white;border:none;padding:18px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:16px;cursor:pointer;' +
    'text-decoration:none;color:inherit;min-height:64px;transition:all 0.2s}' +
    '.action-btn:active{transform:scale(0.98);background:#f0f0f0}' +
    '.action-icon{font-size:28px;width:50px;height:50px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#EDE9FE,#DDD6FE);border-radius:14px;flex-shrink:0}' +
    '.action-label{font-weight:600;color:#333}' +
    '.action-desc{font-size:13px;color:#666;margin-top:3px}' +

    // Loading state
    '.loading{text-align:center;padding:20px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:20px;height:20px;border:2px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    // Refresh indicator
    '.refresh-btn{position:absolute;right:15px;top:50%;transform:translateY(-50%);background:rgba(255,255,255,0.2);border:none;color:white;width:40px;height:40px;border-radius:50%;font-size:20px;cursor:pointer}' +

    '</style></head><body>' +

    // Header
    '<div class="header">' +
    '<button class="refresh-btn" onclick="location.reload()">🔄</button>' +
    '<h1>📊 Dashboard</h1>' +
    '<div class="subtitle">Union Grievance Management</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section - 6 clickable stats in 3x2 grid
    '<div class="stats">' +
    '<a class="stat-card" href="' + baseUrl + '?page=members"><div class="stat-value">' + stats.totalMembers + '</div><div class="stat-label">Members</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Grievances</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=open"><div class="stat-value success">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=pending"><div class="stat-value warning">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=overdue"><div class="stat-value danger">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></a>' +
    '<div class="stat-card"><div class="stat-value success">' + stats.winRate + '</div><div class="stat-label">Win Rate</div></div>' +
    '</div>' +

    // Overdue preview section (loaded dynamically)
    '<div id="overdue-preview"></div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<a class="action-btn" href="' + baseUrl + '?page=search">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find members or grievances</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=grievances">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">All Grievances</div><div class="action-desc">Browse and filter grievances</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=members">' +
    '<div class="action-icon">👥</div>' +
    '<div><div class="action-label">Members</div><div class="action-desc">View member directory</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=links">' +
    '<div class="action-icon">🔗</div>' +
    '<div><div class="action-label">Links</div><div class="action-desc">Forms, resources, GitHub</div></div>' +
    '</a>' +

    '</div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item active" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    // Script to load overdue preview
    '<script>' +
    getClientSideEscapeHtml() +
    'var baseUrl=' + JSON.stringify(baseUrl) + ';' +
    'var retryCount=0;' +
    'function loadOverdue(){' +
    '  if(!navigator.onLine){document.getElementById("overdue-preview").innerHTML="<div style=\\"padding:15px;text-align:center;color:#666\\">📡 Offline</div>";return}' +
    '  document.getElementById("overdue-preview").innerHTML="<div style=\\"padding:15px;text-align:center;color:#666\\">Loading...</div>";' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    retryCount=0;' +
    '    if(!data||!Array.isArray(data)){document.getElementById("overdue-preview").innerHTML="";return}' +
    '    var overdue=data.filter(function(g){return g&&g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overdue-preview").innerHTML="<div style=\\"padding:15px;text-align:center;color:#10B981\\">✅ No overdue cases - great job!</div>";return}' +
    '    var html="<div class=\\"overdue-section\\"><div class=\\"overdue-title\\">⚠️ Overdue Cases ("+overdue.length+")</div>";' +
    '    overdue.slice(0,3).forEach(function(g){' +
    '      html+="<div class=\\"overdue-item\\"><div class=\\"overdue-id\\">"+escapeHtml(g.id||"")+"</div><div class=\\"overdue-name\\">"+escapeHtml(g.name||"")+"</div><div class=\\"overdue-detail\\">"+escapeHtml(g.category||"")+" • "+escapeHtml(g.step||"")+"</div></div>";' +
    '    });' +
    '    if(overdue.length>3)html+="<button class=\\"view-all-btn\\" onclick=\\"location.href=baseUrl+\'?page=grievances&filter=overdue\'\\">View All "+overdue.length+" Overdue Cases</button>";' +
    '    html+="</div>";' +
    '    document.getElementById("overdue-preview").innerHTML=html;' +
    '  }).withFailureHandler(function(err){' +
    '    console.error("Failed to load overdue:",err);' +
    '    if(retryCount<3){retryCount++;var delay=Math.pow(2,retryCount)*1000;setTimeout(loadOverdue,delay)}' +
    '    else{document.getElementById("overdue-preview").innerHTML="<div style=\\"padding:15px;text-align:center;color:#EF4444\\">Failed to load<br><button onclick=\\"retryCount=0;loadOverdue()\\" style=\\"margin-top:8px;padding:6px 12px;background:#7C3AED;color:white;border:none;border-radius:6px;cursor:pointer\\">Retry</button></div>"}' +
    '  }).getWebAppGrievanceList();' +
    '}' +
    'window.addEventListener("online",function(){retryCount=0;loadOverdue()});' +
    'loadOverdue();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the search page HTML for web app
 */
function getWebAppSearchHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Search - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);margin-bottom:12px;text-align:center}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:14px 14px 14px 45px;border:none;border-radius:12px;font-size:16px;background:white;-webkit-appearance:none}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#666}' +
    '.clear-btn{position:absolute;right:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#999;background:none;border:none;cursor:pointer;display:none}' +

    // Tabs
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0;position:sticky;top:76px;z-index:99}' +
    '.tab{flex:1;padding:14px;text-align:center;font-size:14px;font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:48px}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED}' +

    // Results
    '.results{padding:15px}' +
    '.result-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px}' +
    '.result-type{font-size:12px;color:#7C3AED;font-weight:600;text-transform:uppercase;margin-bottom:6px}' +
    '.result-title{font-size:17px;font-weight:600;color:#333;margin-bottom:4px}' +
    '.result-detail{font-size:14px;color:#666;margin-top:4px}' +
    '.result-badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;margin-top:8px}' +
    '.badge-open{background:#FEE2E2;color:#DC2626}' +
    '.badge-pending{background:#FEF3C7;color:#D97706}' +
    '.badge-resolved{background:#D1FAE5;color:#059669}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +
    '.empty-text{font-size:16px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔍 Search</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search members or grievances..." oninput="handleSearch(this.value)" autocomplete="off" autocapitalize="off">' +
    '<button class="clear-btn" id="clearBtn" onclick="clearSearch()">✕</button>' +
    '</div></div>' +

    '<div class="tabs">' +
    '<button class="tab active" data-tab="all" onclick="setTab(\'all\',this)">All</button>' +
    '<button class="tab" data-tab="members" onclick="setTab(\'members\',this)">Members</button>' +
    '<button class="tab" data-tab="grievances" onclick="setTab(\'grievances\',this)">Grievances</button>' +
    '</div>' +

    '<div class="results" id="results">' +
    '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">Type to search members or grievances</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    getClientSideEscapeHtml() +
    'var currentTab="all";' +
    'var searchTimeout=null;' +
    'var lastQuery="";' +

    'function setTab(tab,btn){' +
    '  currentTab=tab;' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  if(lastQuery.length>=2)performSearch(lastQuery);' +
    '}' +

    'function handleSearch(q){' +
    '  lastQuery=q;' +
    '  document.getElementById("clearBtn").style.display=q?"block":"none";' +
    '  if(searchTimeout)clearTimeout(searchTimeout);' +
    '  if(!q||q.length<2){' +
    '    showEmpty("Type at least 2 characters to search");' +
    '    return;' +
    '  }' +
    '  showLoading();' +
    '  searchTimeout=setTimeout(function(){performSearch(q)},300);' +
    '}' +

    'function performSearch(q){' +
    '  google.script.run.withSuccessHandler(renderResults).withFailureHandler(showError).getWebAppSearchResults(q,currentTab);' +
    '}' +

    'function clearSearch(){' +
    '  document.getElementById("searchInput").value="";' +
    '  document.getElementById("clearBtn").style.display="none";' +
    '  lastQuery="";' +
    '  showEmpty("Type to search members or grievances");' +
    '}' +

    'function showEmpty(msg){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">🔍</div><div class=\\"empty-text\\">"+msg+"</div></div>";' +
    '}' +

    'function showLoading(){' +
    '  document.getElementById("results").innerHTML="<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:15px\\">Searching...</div></div>";' +
    '}' +

    'function showError(err){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error: "+escapeHtml(err.message||"Unknown error")+"</div></div>";' +
    '}' +

    'function getBadgeClass(status){' +
    '  if(!status)return"";' +
    '  var s=status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"badge-open";' +
    '  if(s.indexOf("pending")>=0)return"badge-pending";' +
    '  if(s.indexOf("resolved")>=0||s.indexOf("closed")>=0||s.indexOf("withdrawn")>=0)return"badge-resolved";' +
    '  return"";' +
    '}' +

    'function renderResults(data){' +
    '  var c=document.getElementById("results");' +
    '  if(!data||data.length===0){' +
    '    showEmpty("No results found");' +
    '    return;' +
    '  }' +
    '  c.innerHTML=data.map(function(r){' +
    '    var badge=r.status?"<span class=\\"result-badge "+getBadgeClass(r.status)+"\\">"+escapeHtml(r.status)+"</span>":"";' +
    '    return"<div class=\\"result-card\\">"+"<div class=\\"result-type\\">"+(r.type==="member"?"👤 Member":"📋 Grievance")+"</div>"+"<div class=\\"result-title\\">"+escapeHtml(r.title)+"</div>"+"<div class=\\"result-detail\\">"+escapeHtml(r.subtitle)+"</div>"+(r.detail?"<div class=\\"result-detail\\">"+escapeHtml(r.detail)+"</div>":"")+badge+"</div>";' +
    '  }).join("");' +
    '}' +

    // Auto-focus search on load
    'document.getElementById("searchInput").focus();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the grievance list HTML for web app (enhanced with Overdue filter, expandable details)
 */
function getWebAppGrievanceListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Grievances - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px 15px 12px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +

    // Filter pills with Overdue
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 16px;border-radius:20px;font-size:13px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +
    '.filter-pill.danger{background:#DC2626;color:white}' +
    '.filter-pill.danger.active{background:#FEE2E2;color:#DC2626}' +

    // List with expandable cards
    '.grievance-list{padding:15px}' +
    '.grievance-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.grievance-card:active{transform:scale(0.99)}' +
    '.grievance-card.overdue{border-left:4px solid #DC2626}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}' +
    '.grievance-id{font-size:15px;font-weight:700;color:#7C3AED}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600}' +
    '.status-open{background:#FEE2E2;color:#DC2626}' +
    '.status-pending{background:#FEF3C7;color:#D97706}' +
    '.status-resolved{background:#D1FAE5;color:#059669}' +
    '.status-overdue{background:#DC2626;color:white;animation:pulse 2s infinite}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}' +
    '.grievance-name{font-size:16px;font-weight:500;color:#333;margin-bottom:4px}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:4px}' +
    '.grievance-step{display:inline-block;padding:3px 8px;background:#E0E7FF;color:#4F46E5;border-radius:6px;font-size:11px;font-weight:500;margin-top:8px}' +

    // Expandable details
    '.grievance-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.grievance-card.expanded .grievance-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#333;font-weight:500}' +
    '.detail-value.danger{color:#DC2626}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>📋 Grievances</h2>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="open" onclick="setFilter(\'open\',this)">Open</button>' +
    '<button class="filter-pill" data-filter="pending" onclick="setFilter(\'pending\',this)">Pending</button>' +
    '<button class="filter-pill danger" data-filter="overdue" onclick="setFilter(\'overdue\',this)">⚠️ Overdue</button>' +
    '<button class="filter-pill" data-filter="resolved" onclick="setFilter(\'resolved\',this)">Resolved</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="grievance-list" id="grievanceList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading grievances...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=25;' +
    'var currentPage=0;' +
    'var dataCache={data:null,timestamp:0};' +
    'var CACHE_TTL=300000;' +

    // Check URL for filter parameter
    'var urlParams=new URLSearchParams(window.location.search);' +
    'var initialFilter=urlParams.get("filter");' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  currentPage=0;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  if(btn)btn.classList.add("active");' +
    '  renderList();' +
    '}' +

    'function getStatusClass(g){' +
    '  if(g.isOverdue)return"status-overdue";' +
    '  if(!g.status)return"";' +
    '  var s=g.status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"status-open";' +
    '  if(s.indexOf("pending")>=0)return"status-pending";' +
    '  return"status-resolved";' +
    '}' +

    'function getStatusText(g){' +
    '  if(g.isOverdue)return"⚠️ OVERDUE";' +
    '  return g.status||"";' +
    '}' +

    'function matchesFilter(g){' +
    '  if(currentFilter==="all")return true;' +
    '  if(currentFilter==="overdue")return g.isOverdue;' +
    '  if(!g.status)return false;' +
    '  var s=g.status.toLowerCase();' +
    '  if(currentFilter==="open")return s.indexOf("open")>=0;' +
    '  if(currentFilter==="pending")return s.indexOf("pending")>=0;' +
    '  if(currentFilter==="resolved")return s.indexOf("resolved")>=0||s.indexOf("withdrawn")>=0||s.indexOf("closed")>=0;' +
    '  return true;' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function renderList(){' +
    '  var filtered=allData.filter(function(g){return matchesFilter(g)});' +
    '  var totalPages=Math.ceil(filtered.length/PAGE_SIZE);' +
    '  var start=currentPage*PAGE_SIZE;' +
    '  var paged=filtered.slice(start,start+PAGE_SIZE);' +
    '  document.getElementById("countBadge").textContent="Showing "+(start+1)+"-"+Math.min(start+PAGE_SIZE,filtered.length)+" of "+filtered.length;' +
    '  var c=document.getElementById("grievanceList");' +
    '  if(filtered.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div>No grievances found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=paged.map(function(g){' +
    '    var cardClass="grievance-card"+(g.isOverdue?" overdue":"");' +
    '    var daysInfo=g.isOverdue?"<span class=\\"detail-value danger\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?"<span class=\\"detail-value\\">"+g.daysToDeadline+" days</span>":"<span class=\\"detail-value\\">N/A</span>");' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"grievance-header\\">"+"<span class=\\"grievance-id\\">"+escapeHtml(g.id)+"</span>"+"<span class=\\"grievance-status "+getStatusClass(g)+"\\">"+getStatusText(g)+"</span>"+"</div>"+"<div class=\\"grievance-name\\">"+escapeHtml(g.name)+"</div>"+(g.category?"<div class=\\"grievance-detail\\">"+escapeHtml(g.category)+"</div>":"")+(g.step?"<span class=\\"grievance-step\\">"+escapeHtml(g.step)+"</span>":"")+"<div class=\\"grievance-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+escapeHtml(g.filedDate)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+escapeHtml(g.incidentDate)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span>"+daysInfo+"</div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+escapeHtml(g.daysOpen)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+escapeHtml(g.location)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+escapeHtml(g.articles)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+escapeHtml(g.steward)+"</span></div>"+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+escapeHtml(g.resolution)+"</span></div>":"")+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(totalPages>1){html+="<div style=\\"display:flex;justify-content:center;gap:10px;padding:15px\\"><button onclick=\\"prevPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage>0?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage===0?"disabled":"")+">← Prev</button><span style=\\"display:flex;align-items:center\\">Page "+(currentPage+1)+" of "+totalPages+"</span><button onclick=\\"nextPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage<totalPages-1?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage>=totalPages-1?"disabled":"")+">Next →</button></div>"}' +
    '  c.innerHTML=html;' +
    '}' +
    'function prevPage(){if(currentPage>0){currentPage--;renderList();window.scrollTo(0,0)}}' +
    'function nextPage(){var filtered=allData.filter(function(g){return matchesFilter(g)});if(currentPage<Math.ceil(filtered.length/PAGE_SIZE)-1){currentPage++;renderList();window.scrollTo(0,0)}}' +

    'function loadData(){' +
    '  if(!navigator.onLine){document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📡</div><div>You appear to be offline</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:10px 20px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";return}' +
    '  if(dataCache.data&&(Date.now()-dataCache.timestamp)<CACHE_TTL){console.log("Using cached data");allData=dataCache.data;if(initialFilter){currentFilter=initialFilter;var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}}renderList();return}' +
    '  console.log("Loading grievance data...");' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    console.log("Data received:",data?data.length:0,"items");' +
    '    allData=data||[];' +
    '    dataCache={data:allData,timestamp:Date.now()};' +
    '    if(initialFilter){' +
    '      currentFilter=initialFilter;' +
    '      var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");' +
    '      if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}' +
    '    }' +
    '    renderList();' +
    '  }).withFailureHandler(function(err){' +
    '    console.error("Failed to load data:",err);' +
    '    document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:10px 20px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button><div style=\\"font-size:11px;color:#999;margin-top:8px\\">"+String(err||"Unknown error")+"</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'window.addEventListener("online",loadData);' +
    'window.addEventListener("pagehide",function(){allData=null;dataCache={data:null,timestamp:0}});' +
    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the member list HTML for web app
 */
function getWebAppMemberListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Members - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:none;border-radius:12px;font-size:15px;background:white}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:18px;color:#666}' +

    // Filter pills
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:8px 0 2px;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 14px;border-radius:20px;font-size:12px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Member list
    '.member-list{padding:15px}' +
    '.member-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.member-card:active{transform:scale(0.99)}' +
    '.member-card.has-grievance{border-left:4px solid #F97316}' +
    '.member-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px}' +
    '.member-name{font-size:16px;font-weight:600;color:#333}' +
    '.member-id{font-size:12px;color:#7C3AED;font-weight:500}' +
    '.member-title{font-size:14px;color:#666;margin-bottom:4px}' +
    '.member-location{font-size:13px;color:#888}' +
    '.member-badges{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}' +
    '.badge{padding:3px 8px;border-radius:12px;font-size:11px;font-weight:500}' +
    '.badge-steward{background:#DDD6FE;color:#7C3AED}' +
    '.badge-grievance{background:#FEF3C7;color:#D97706}' +

    // Expandable details
    '.member-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.member-card.expanded .member-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:80px}' +
    '.detail-value{color:#333;font-weight:500}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>👥 Members</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search by name, ID, title..." oninput="debounceSearch()">' +
    '</div>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="steward" onclick="setFilter(\'steward\',this)">Stewards</button>' +
    '<button class="filter-pill" data-filter="grievance" onclick="setFilter(\'grievance\',this)">With Grievance</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="member-list" id="memberList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading members...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=25;' +
    'var currentPage=0;' +
    'var dataCache={data:null,timestamp:0};' +
    'var CACHE_TTL=300000;' +
    'var searchTimeout=null;' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  currentPage=0;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterMembers();' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function debounceSearch(){clearTimeout(searchTimeout);searchTimeout=setTimeout(function(){currentPage=0;filterMembers()},300)}' +

    'function filterMembers(){' +
    '  var query=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allData.filter(function(m){' +
    '    var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0||(m.title||"").toLowerCase().indexOf(query)>=0||(m.location||"").toLowerCase().indexOf(query)>=0;' +
    '    var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);' +
    '    return matchesQuery&&matchesFilter;' +
    '  });' +
    '  var totalPages=Math.ceil(filtered.length/PAGE_SIZE);' +
    '  var start=currentPage*PAGE_SIZE;' +
    '  var paged=filtered.slice(start,start+PAGE_SIZE);' +
    '  document.getElementById("countBadge").textContent="Showing "+(filtered.length>0?(start+1)+"-"+Math.min(start+PAGE_SIZE,filtered.length):"0")+" of "+filtered.length;' +
    '  renderList(paged,totalPages);' +
    '}' +

    'function renderList(data,totalPages){' +
    '  var c=document.getElementById("memberList");' +
    '  if(!data||data.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div>No members found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=data.map(function(m){' +
    '    var cardClass="member-card"+(m.hasOpenGrievance?" has-grievance":"");' +
    '    var badges="";' +
    '    if(m.isSteward)badges+="<span class=\\"badge badge-steward\\">🛡️ Steward</span>";' +
    '    if(m.hasOpenGrievance)badges+="<span class=\\"badge badge-grievance\\">⚠️ Open Grievance</span>";' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"member-header\\"><span class=\\"member-name\\">"+escapeHtml(m.name)+"</span><span class=\\"member-id\\">"+escapeHtml(m.id)+"</span></div>"+"<div class=\\"member-title\\">"+escapeHtml(m.title)+"</div>"+"<div class=\\"member-location\\">📍 "+escapeHtml(m.location)+"</div>"+(badges?"<div class=\\"member-badges\\">"+badges+"</div>":"")+"<div class=\\"member-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+escapeHtml(m.email||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📞 Phone:</span><span class=\\"detail-value\\">"+escapeHtml(m.phone||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+escapeHtml(m.unit)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">👔 Supervisor:</span><span class=\\"detail-value\\">"+escapeHtml(m.supervisor)+"</span></div>"+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(totalPages>1){html+="<div style=\\"display:flex;justify-content:center;gap:10px;padding:15px\\"><button onclick=\\"prevPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage>0?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage===0?"disabled":"")+">← Prev</button><span style=\\"display:flex;align-items:center\\">Page "+(currentPage+1)+" of "+totalPages+"</span><button onclick=\\"nextPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage<totalPages-1?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage>=totalPages-1?"disabled":"")+">Next →</button></div>"}' +
    '  c.innerHTML=html;' +
    '}' +
    'function prevPage(){if(currentPage>0){currentPage--;filterMembers();window.scrollTo(0,0)}}' +
    'function nextPage(){var query=(document.getElementById("searchInput").value||"").toLowerCase();var filtered=allData.filter(function(m){var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0;var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);return matchesQuery&&matchesFilter});if(currentPage<Math.ceil(filtered.length/PAGE_SIZE)-1){currentPage++;filterMembers();window.scrollTo(0,0)}}' +

    'function loadData(){' +
    '  if(!navigator.onLine){document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📡</div><div>You appear to be offline</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:8px 16px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";return}' +
    '  if(dataCache.data&&(Date.now()-dataCache.timestamp)<CACHE_TTL){allData=dataCache.data;filterMembers();return}' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allData=data||[];' +
    '    dataCache={data:allData,timestamp:Date.now()};' +
    '    filterMembers();' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:8px 16px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";' +
    '  }).getWebAppMemberList();' +
    '}' +

    'window.addEventListener("online",loadData);' +
    'window.addEventListener("pagehide",function(){allData=null;dataCache={data:null,timestamp:0}});' +
    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the links/resources page HTML for web app
 */
function getWebAppLinksHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Links - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px)}' +
    '.header .subtitle{font-size:13px;opacity:0.9;margin-top:5px}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Link cards
    '.link-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}' +
    '.link-card{background:white;padding:20px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-decoration:none;display:flex;flex-direction:column;align-items:center;text-align:center;gap:10px;transition:all 0.2s}' +
    '.link-card:active{transform:scale(0.96);background:#f8f4ff}' +
    '.link-icon{font-size:32px}' +
    '.link-label{font-weight:600;color:#333;font-size:14px}' +
    '.link-desc{font-size:12px;color:#666}' +

    // Full-width link
    '.link-card.full{grid-column:span 2;flex-direction:row;padding:16px;justify-content:flex-start;text-align:left}' +
    '.link-card.full .link-icon{font-size:28px}' +
    '.link-card.full .link-content{flex:1}' +

    // GitHub special styling
    '.link-card.github{background:linear-gradient(135deg,#24292e,#1a1e22);color:white}' +
    '.link-card.github .link-label{color:white}' +
    '.link-card.github .link-desc{color:rgba(255,255,255,0.7)}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔗 Links & Resources</h2>' +
    '<div class="subtitle">Quick access to forms and tools</div>' +
    '</div>' +

    '<div class="container" id="linksContent">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading links...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    getClientSideEscapeHtml() +
    'function safeUrl(u){if(!u)return"#";u=String(u);return/^https?:\\/\\//i.test(u)?u:"#";}' +
    'function loadLinks(){' +
    '  google.script.run.withSuccessHandler(function(links){' +
    '    renderLinks(links);' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("linksContent").innerHTML="<div class=\\"loading\\">⚠️ Error loading links</div>";' +
    '  }).getWebAppResourceLinks();' +
    '}' +

    'function renderLinks(links){' +
    '  var html="";' +

    '  // Forms section' +
    '  html+="<div class=\\"section-title\\">📝 Forms</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  if(links.grievanceForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.grievanceForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📋</span><span class=\\"link-label\\">Grievance Form</span><span class=\\"link-desc\\">File a grievance</span></a>";}' +
    '  if(links.contactForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.contactForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">✉️</span><span class=\\"link-label\\">Contact Form</span><span class=\\"link-desc\\">Send a message</span></a>";}' +
    '  if(links.satisfactionForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.satisfactionForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Satisfaction Survey</span><span class=\\"link-desc\\">Give feedback</span></a>";}' +
    '  if(!links.grievanceForm&&!links.contactForm&&!links.satisfactionForm){html+="<div class=\\"link-card full\\"><span class=\\"link-icon\\">ℹ️</span><div class=\\"link-content\\"><span class=\\"link-label\\">No Forms Configured</span><span class=\\"link-desc\\">Add form URLs to Config sheet</span></div></div>";}' +
    '  html+="</div>";' +

    '  // Resources section' +
    '  html+="<div class=\\"section-title\\">🔧 Resources</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.spreadsheetUrl))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Spreadsheet</span><span class=\\"link-desc\\">Open full dashboard</span></a>";' +
    '  html+="<a class=\\"link-card github\\" href=\\""+escapeHtml(safeUrl(links.githubRepo))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📦</span><span class=\\"link-label\\">GitHub Repo</span><span class=\\"link-desc\\">Source code</span></a>";' +
    '  html+="</div>";' +

    '  document.getElementById("linksContent").innerHTML=html;' +
    '}' +

    'loadLinks();' +
    '</script>' +

    '</body></html>';
}

/**
 * API function to get search results for web app
 * @param {string} query - Search query
 * @param {string} tab - Tab filter (all, members, grievances)
 * @returns {Array} Search results
 */
function getWebAppSearchResults(query, tab) {
  return getMobileSearchData(query, tab);
}

/**
 * API function to get grievance list for web app (full fields like Interactive Dashboard)
 * @returns {Array} Grievance data with all fields
 */
function getWebAppGrievanceList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppGrievanceList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) {
      Logger.log('getWebAppGrievanceList: Grievance Log sheet not found');
      return [];
    }

    ensureMinimumColumns(sheet, getGrievanceHeaders().length);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppGrievanceList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

      var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
      var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

      return {
        id: grievanceId,
        memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
        name: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
        status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
        step: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
        category: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
        articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
        filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
        incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
        nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
        daysToDeadline: daysToDeadline,
        isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
        daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
        location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
        steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
        resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
      };
    }).filter(function(g) { return g !== null; }).slice(0, 100);

    Logger.log('getWebAppGrievanceList: Returning ' + result.length + ' grievances');
    return result;
  } catch (e) {
    Logger.log('getWebAppGrievanceList error: ' + e.toString());
    throw new Error('Failed to load grievances: ' + e.message);
  }
}

/**
 * API function to get member list for web app
 * @returns {Array} Member data
 */
function getWebAppMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppMemberList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) {
      Logger.log('getWebAppMemberList: Member Directory sheet not found');
      return [];
    }

    ensureMinimumColumns(sheet, getMemberHeaders().length);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppMemberList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();

    var result = data.map(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return null;

      return {
        id: memberId,
        firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
        name: ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim(),
        title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
        location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
        unit: row[MEMBER_COLS.UNIT - 1] || 'N/A',
        email: row[MEMBER_COLS.EMAIL - 1] || '',
        phone: row[MEMBER_COLS.PHONE - 1] || '',
        isSteward: isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1]),
        supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
        hasOpenGrievance: isTruthyValue(row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1])
      };
    }).filter(function(m) { return m !== null; }).slice(0, 100);

    Logger.log('getWebAppMemberList: Returning ' + result.length + ' members');
    return result;
  } catch (e) {
    Logger.log('getWebAppMemberList error: ' + e.toString());
    throw new Error('Failed to load members: ' + e.message);
  }
}

/**
 * API function to get resource links for web app
 * @returns {Object} Resource links
 */
function getWebAppResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: '',
    githubRepo: ''  // Set via Config sheet ORG_WEBSITE or manually
  };

  // Get form URLs from Config sheet using CONFIG_COLS constants (data in row 3)
  if (configSheet && configSheet.getLastRow() >= 3) {
    try {
      var configRow = configSheet.getRange(3, 1, 1, CONFIG_COLS.SATISFACTION_FORM_URL).getValues()[0];
      links.grievanceForm = configRow[CONFIG_COLS.GRIEVANCE_FORM_URL - 1] || '';
      links.contactForm = configRow[CONFIG_COLS.CONTACT_FORM_URL - 1] || '';
      links.satisfactionForm = configRow[CONFIG_COLS.SATISFACTION_FORM_URL - 1] || '';
      links.orgWebsite = configRow[CONFIG_COLS.ORG_WEBSITE - 1] || '';
    } catch (_e) {
      // Ignore errors reading config
    }
  }

  return links;
}

/**
 * API function to get dashboard stats with win rate for web app
 * @returns {Object} Dashboard statistics
 */
function getWebAppDashboardStats() {
  var stats = getMobileDashboardStats();

  // Calculate win rate
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (sheet && sheet.getLastRow() > 1) {
    var resolutions = sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, sheet.getLastRow() - 1, 1).getValues();
    var won = 0, total = 0;
    resolutions.forEach(function(row) {
      var res = (row[0] || '').toString().toLowerCase();
      if (res) {
        total++;
        if (res.indexOf('won') >= 0 || res.indexOf('favorable') >= 0) {
          won++;
        }
      }
    });
    stats.winRate = total > 0 ? Math.round((won / total) * 100) + '%' : 'N/A';
  } else {
    stats.winRate = 'N/A';
  }

  // Also get total members
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
    var validMembers = memberIds.filter(function(row) {
      var id = row[0] || '';
      return id && id.toString().match(/^M/i);
    }).length;
    stats.totalMembers = validMembers;
  } else {
    stats.totalMembers = 0;
  }

  return stats;
}

/**
 * Menu function to show the deployed mobile dashboard URL
 */
function showWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var url = ScriptApp.getService().getUrl();

  if (url) {
    ui.alert(
      '📱 Mobile Dashboard URL',
      'Your mobile dashboard URL:\n\n' + url + '\n\n' +
      'Open this URL on your phone and add it to your home screen for quick access!',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '📱 Mobile Dashboard URL',
      'No web app deployment found.\n\n' +
      'To deploy:\n' +
      '1. Go to Extensions → Apps Script\n' +
      '2. Click "Deploy" → "New deployment"\n' +
      '3. Select type: "Web app"\n' +
      '4. Set "Who has access" to your preference\n' +
      '5. Click "Deploy" and copy the URL',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Add Mobile Dashboard link to Config sheet for easy mobile access
 * Creates a clickable hyperlink cell that works on mobile devices
 */
function addMobileDashboardLinkToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    ui.alert('Error', 'Config sheet not found', ui.ButtonSet.OK);
    return;
  }

  // Prompt user to enter the URL from Manage deployments
  var response = ui.prompt(
    '📱 Add Mobile Dashboard Link',
    'To get your web app URL:\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click "Deploy" → "Manage deployments"\n' +
    '3. Copy the Web app URL\n\n' +
    'Paste the URL below:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var url = response.getResponseText().trim();
  if (!url || !url.match(/^https:\/\/script\.google\.com/)) {
    ui.alert(
      'Invalid URL',
      'Please enter a valid Google Apps Script web app URL.\n\n' +
      'It should look like:\nhttps://script.google.com/macros/s/XXXXX/exec',
      ui.ButtonSet.OK
    );
    return;
  }

  // Use the canonical column constant — never hardcode column numbers.
  var targetCol = CONFIG_COLS.MOBILE_DASHBOARD_URL;

  // Write data starting at row 3 (first data row).
  // Rows 1-2 are section/column headers managed by createConfigSheet — don't touch them.
  var linkCell = configSheet.getRange(3, targetCol);
  linkCell.setFormula('=HYPERLINK(' + JSON.stringify(url) + ', "📱 Tap to Open Dashboard")');
  linkCell.setFontSize(14);
  linkCell.setFontWeight('bold');
  linkCell.setFontColor('#1a73e8');
  linkCell.setBackground('#e8f0fe');

  // Also add plain URL below for copying
  var urlCell = configSheet.getRange(4, targetCol);
  urlCell.setValue(url);
  urlCell.setFontSize(10);
  urlCell.setWrap(true);

  // Add instructions
  var instructionCell = configSheet.getRange(5, targetCol);
  instructionCell.setValue('Open Google Sheets on your phone, navigate to Config tab, and tap the blue link above to access the dashboard.');
  instructionCell.setFontSize(9);
  instructionCell.setFontColor('#666666');
  instructionCell.setWrap(true);

  // Set column width
  configSheet.setColumnWidth(targetCol, 300);

  SpreadsheetApp.getUi().alert(
    '📱 Mobile Dashboard Link Added!',
    'A clickable link has been added to the "📱 Mobile Dashboard URL" column of the Config sheet.\n\n' +
    'To access on mobile:\n' +
    '1. Open this spreadsheet in Google Sheets mobile app\n' +
    '2. Go to the Config tab\n' +
    '3. Scroll to the Mobile Dashboard section\n' +
    '4. Tap the blue "Tap to Open Dashboard" link\n\n' +
    'URL: ' + url,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// CONSTANT CONTACT V3 API INTEGRATION — Read-Only Engagement Metrics
// ============================================================================

/**
 * Constant Contact API configuration
 * @const {Object}
 */
var CC_CONFIG = {
  API_BASE: 'https://api.cc.email/v3',
  AUTH_URL: 'https://authz.constantcontact.com/oauth2/default/v1/authorize',
  TOKEN_URL: 'https://authz.constantcontact.com/oauth2/default/v1/token',
  CONTACTS_ENDPOINT: '/contacts',
  ACTIVITY_SUMMARY_ENDPOINT: '/reports/contact_reports/{contact_id}/activity_summary',
  RATE_LIMIT_PER_SECOND: 4,
  RATE_LIMIT_DELAY_MS: 300,
  PAGE_LIMIT: 500,
  ACTIVITY_LOOKBACK_DAYS: 365,
  PROP_API_KEY: 'CC_API_KEY',
  PROP_API_SECRET: 'CC_API_SECRET',
  PROP_ACCESS_TOKEN: 'CC_ACCESS_TOKEN',
  PROP_REFRESH_TOKEN: 'CC_REFRESH_TOKEN',
  PROP_TOKEN_EXPIRY: 'CC_TOKEN_EXPIRY'
};

/**
 * Shows dialog for entering Constant Contact API credentials.
 * Stores API key and secret in Script Properties (encrypted at rest).
 * One-time setup required before syncing engagement metrics.
 */
function showConstantContactSetup() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var existingKey = props.getProperty(CC_CONFIG.PROP_API_KEY);

  var statusMsg = existingKey
    ? 'Current status: API key is configured (ends in ...' + existingKey.slice(-4) + ')\n\n'
    : 'Current status: Not configured\n\n';

  var response = ui.prompt(
    '📧 Constant Contact Setup',
    statusMsg +
    'Enter your Constant Contact API key (client ID).\n\n' +
    'To get your API key:\n' +
    '1. Go to app.constantcontact.com/pages/dma/portal\n' +
    '2. Click "New Application"\n' +
    '3. Copy your API Key (Client ID)\n\n' +
    'API Key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText().trim()) {
    return;
  }

  var apiKey = response.getResponseText().trim();

  var secretResponse = ui.prompt(
    '📧 Constant Contact Setup (Step 2)',
    'Enter your Client Secret.\n\n' +
    'Click "Generate Client Secret" in your CC app settings.\n' +
    'Important: Copy it immediately — it only appears once.\n\n' +
    'Client Secret:',
    ui.ButtonSet.OK_CANCEL
  );

  if (secretResponse.getSelectedButton() !== ui.Button.OK || !secretResponse.getResponseText().trim()) {
    return;
  }

  var apiSecret = secretResponse.getResponseText().trim();

  props.setProperty(CC_CONFIG.PROP_API_KEY, apiKey);
  props.setProperty(CC_CONFIG.PROP_API_SECRET, apiSecret);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'API credentials saved. Next step: authorize your account.',
    '✅ Constant Contact Configured', 5
  );

  Logger.log('Constant Contact API credentials configured');
}

/**
 * Initiates the OAuth2 authorization flow for Constant Contact.
 * Opens a dialog with the authorization URL for the user to grant access.
 * After granting access, the user pastes back the authorization code.
 */
function authorizeConstantContact() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);

  if (!apiKey) {
    ui.alert('⚠️ Setup Required',
      'Please run "Setup API Credentials" first.',
      ui.ButtonSet.OK);
    return;
  }

  // Check if already authorized
  var existingToken = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  if (existingToken && expiry && new Date(expiry) > new Date()) {
    var reauth = ui.alert('Already Authorized',
      'You already have a valid access token.\nExpires: ' + new Date(expiry).toLocaleString() +
      '\n\nRe-authorize?',
      ui.ButtonSet.YES_NO);
    if (reauth !== ui.Button.YES) return;
  }

  // Build authorization URL — using server flow (authorization code grant)
  var redirectUri = 'https://localhost';
  var scope = 'contact_data offline_access';
  var state = Utilities.getUuid();

  var authUrl = CC_CONFIG.AUTH_URL +
    '?response_type=code' +
    '&client_id=' + encodeURIComponent(apiKey) +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=' + encodeURIComponent(scope) +
    '&state=' + encodeURIComponent(state);

  // Show the URL for the user to visit
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px;max-width:500px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.step{background:#f0f7ff;padding:15px;margin:12px 0;border-radius:8px;border-left:4px solid #1a73e8}' +
    '.step-num{font-weight:bold;color:#1a73e8}' +
    'a{color:#1a73e8;word-break:break-all}' +
    'input{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;margin:8px 0;font-size:14px}' +
    'button{padding:12px 24px;background:#1a73e8;color:white;border:none;border-radius:4px;cursor:pointer;font-size:14px}' +
    'button:hover{background:#1557b0}' +
    '.note{font-size:12px;color:#666;margin-top:12px}' +
    '</style>' +
    '<div class="container">' +
    '<h2>Authorize Constant Contact</h2>' +
    '<div class="step"><span class="step-num">Step 1:</span> Click the link below to authorize:<br><br>' +
    '<a href="' + authUrl + '" target="_blank">Open Constant Contact Authorization</a></div>' +
    '<div class="step"><span class="step-num">Step 2:</span> Log in and click "Allow"</div>' +
    '<div class="step"><span class="step-num">Step 3:</span> You\'ll be redirected to a URL like:<br>' +
    '<code>https://localhost?code=XXXX&state=...</code><br><br>' +
    'Copy the entire URL from your browser address bar and paste it below:<br>' +
    '<input type="text" id="callbackUrl" placeholder="Paste the full redirect URL here...">' +
    '<button onclick="submitCode()">Submit</button></div>' +
    '<div class="note">Your browser may show an error page — that\'s expected. ' +
    'Just copy the URL from the address bar.</div>' +
    '</div>' +
    '<script>' +
    'function submitCode(){' +
    '  var url=document.getElementById("callbackUrl").value.trim();' +
    '  if(!url){alert("Please paste the redirect URL");return;}' +
    '  var match=url.match(/[?&]code=([^&]+)/);' +
    '  if(!match){alert("Could not find authorization code in that URL. Make sure you copied the full URL.");return;}' +
    '  google.script.run' +
    '    .withSuccessHandler(function(msg){' +
    '      var c=document.querySelector(".container");c.innerHTML="";var h=document.createElement("h2");h.textContent="\\u2705 "+msg;var p=document.createElement("p");p.textContent="You can close this dialog.";c.appendChild(h);c.appendChild(p);' +
    '    })' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);})' +
    '    .exchangeConstantContactCode(match[1]);' +
    '}' +
    '</script>'
  ).setWidth(550).setHeight(520);

  ui.showModalDialog(html, '📧 Authorize Constant Contact');
}

/**
 * Exchanges an authorization code for access and refresh tokens.
 * Called from the authorization dialog after user grants access.
 * @param {string} code - The authorization code from the OAuth callback
 * @returns {string} Success message
 */
function exchangeConstantContactCode(code) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var apiSecret = props.getProperty(CC_CONFIG.PROP_API_SECRET);

  if (!apiKey || !apiSecret) {
    throw new Error('API credentials not configured. Run setup first.');
  }

  var payload = {
    grant_type: 'authorization_code',
    code: code,
    redirect_uri: 'https://localhost'
  };

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret)
    },
    payload: Object.keys(payload).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(payload[k]);
    }).join('&'),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CC_CONFIG.TOKEN_URL, options);
  var responseCode = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (responseCode !== 200) {
    var errorMsg = body.error_description || body.error || 'Unknown error';
    throw new Error('Token exchange failed (' + responseCode + '): ' + errorMsg);
  }

  // Store tokens
  storeConstantContactTokens_(body);

  Logger.log('Constant Contact authorized successfully');
  return 'Authorized Successfully!';
}

/**
 * Stores OAuth tokens from a token response.
 * @param {Object} tokenResponse - The token endpoint response
 * @private
 */
function storeConstantContactTokens_(tokenResponse) {
  var props = PropertiesService.getScriptProperties();

  props.setProperty(CC_CONFIG.PROP_ACCESS_TOKEN, tokenResponse.access_token);

  if (tokenResponse.refresh_token) {
    props.setProperty(CC_CONFIG.PROP_REFRESH_TOKEN, tokenResponse.refresh_token);
  }

  // Calculate expiry (tokens last ~2 hours; subtract 5 min buffer)
  var expiresIn = (tokenResponse.expires_in || 7200) - 300;
  var expiry = new Date(Date.now() + expiresIn * 1000).toISOString();
  props.setProperty(CC_CONFIG.PROP_TOKEN_EXPIRY, expiry);
}

/**
 * Gets a valid access token, refreshing if expired.
 * @returns {string|null} The access token, or null if not authorized
 * @private
 */
function getConstantContactToken_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  var refreshToken = props.getProperty(CC_CONFIG.PROP_REFRESH_TOKEN);

  if (!token) return null;

  // Check if token is still valid
  if (expiry && new Date(expiry) > new Date()) {
    return token;
  }

  // Token expired — try to refresh
  if (!refreshToken) return null;

  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var apiSecret = props.getProperty(CC_CONFIG.PROP_API_SECRET);

  if (!apiKey || !apiSecret) return null;

  var payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken
  };

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret)
    },
    payload: Object.keys(payload).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(payload[k]);
    }).join('&'),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(CC_CONFIG.TOKEN_URL, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('CC token refresh failed: ' + response.getContentText());
      return null;
    }

    var body = JSON.parse(response.getContentText());
    storeConstantContactTokens_(body);
    return body.access_token;
  } catch (e) {
    Logger.log('CC token refresh error: ' + e.message);
    return null;
  }
}

/**
 * Makes an authenticated GET request to the Constant Contact v3 API.
 * Handles token refresh and rate limiting.
 * @param {string} endpoint - The API endpoint path (e.g., '/contacts')
 * @param {Object} [params] - Optional query parameters
 * @returns {Object|null} Parsed JSON response, or null on failure
 * @private
 */
function ccApiGet_(endpoint, params) {
  var token = getConstantContactToken_();
  if (!token) return null;

  var url = CC_CONFIG.API_BASE + endpoint;

  if (params) {
    var queryParts = [];
    for (var key in params) {
      if (params.hasOwnProperty(key) && params[key] !== undefined && params[key] !== null) {
        queryParts.push(encodeURIComponent(key) + '=' + encodeURIComponent(params[key]));
      }
    }
    if (queryParts.length > 0) {
      url += '?' + queryParts.join('&');
    }
  }

  var options = {
    method: 'get',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 401) {
      // Token may have just expired — force refresh and retry once
      var props = PropertiesService.getScriptProperties();
      props.deleteProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
      token = getConstantContactToken_();
      if (!token) return null;

      options.headers['Authorization'] = 'Bearer ' + token;
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
    }

    if (code === 429) {
      // Rate limited — wait and retry once
      Utilities.sleep(1000);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
    }

    if (code !== 200) {
      Logger.log('CC API error (' + code + '): ' + response.getContentText());
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('CC API request error: ' + e.message);
    return null;
  }
}

/**
 * Fetches all contacts from Constant Contact, handling pagination.
 * Returns a map of lowercase email -> contact_id for matching.
 * @returns {Object} Map of email -> contact_id
 * @private
 */
function fetchCCContacts_() {
  var emailToContactId = {};
  var endpoint = CC_CONFIG.CONTACTS_ENDPOINT;
  var params = {
    limit: CC_CONFIG.PAGE_LIMIT,
    include: 'email_address'
  };

  var pageCount = 0;
  var maxPages = 20; // Safety limit

  while (pageCount < maxPages) {
    var data = ccApiGet_(endpoint, params);
    if (!data || !data.contacts) break;

    for (var i = 0; i < data.contacts.length; i++) {
      var contact = data.contacts[i];
      var contactId = contact.contact_id;
      var emailAddresses = contact.email_address;

      if (emailAddresses && emailAddresses.address) {
        emailToContactId[emailAddresses.address.toLowerCase()] = contactId;
      }
    }

    // Check for next page
    if (data._links && data._links.next && data._links.next.href) {
      // The next link is a full path — extract just the path + query
      var nextUrl = data._links.next.href;
      // Parse the path portion
      var pathMatch = nextUrl.match(/\/v3(\/contacts.*)/);
      if (pathMatch) {
        endpoint = pathMatch[1].split('?')[0];
        // Parse query params from the next URL
        var queryString = pathMatch[1].split('?')[1] || '';
        params = {};
        queryString.split('&').forEach(function(part) {
          var kv = part.split('=');
          if (kv.length === 2) {
            params[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1]);
          }
        });
      } else {
        break;
      }
    } else {
      break;
    }

    pageCount++;
    Utilities.sleep(CC_CONFIG.RATE_LIMIT_DELAY_MS);
  }

  return emailToContactId;
}

/**
 * Fetches engagement metrics (activity summary) for a single CC contact.
 * @param {string} contactId - The CC contact UUID
 * @returns {Object|null} Object with openRate and lastActivityDate, or null
 * @private
 */
function fetchCCContactEngagement_(contactId) {
  var now = new Date();
  var start = new Date(now.getTime() - CC_CONFIG.ACTIVITY_LOOKBACK_DAYS * 24 * 60 * 60 * 1000);

  var endpoint = CC_CONFIG.ACTIVITY_SUMMARY_ENDPOINT.replace('{contact_id}', contactId);
  var params = {
    start: start.toISOString(),
    end: now.toISOString()
  };

  var data = ccApiGet_(endpoint, params);
  if (!data || !data.campaign_activities) {
    return null;
  }

  var totalSends = 0;
  var totalOpens = 0;
  var lastActivityDate = null;

  for (var i = 0; i < data.campaign_activities.length; i++) {
    var activity = data.campaign_activities[i];

    totalSends += (activity.em_sends || 0);
    totalOpens += (activity.em_opens || 0);

    // Track the most recent activity date
    var dates = [activity.em_sends_date, activity.em_opens_date, activity.em_clicks_date]
      .filter(function(d) { return d; });

    for (var j = 0; j < dates.length; j++) {
      var d = new Date(dates[j]);
      if (!isNaN(d.getTime()) && (!lastActivityDate || d > lastActivityDate)) {
        lastActivityDate = d;
      }
    }
  }

  var openRate = totalSends > 0 ? Math.round((totalOpens / totalSends) * 100) : 0;

  return {
    openRate: openRate,
    lastActivityDate: lastActivityDate
  };
}

/**
 * Syncs Constant Contact email engagement metrics to the Member Directory.
 * Read-only: pulls open rates and last activity dates from CC, never writes to CC.
 *
 * Matches CC contacts to members by email address (case-insensitive).
 * Updates OPEN_RATE and RECENT_CONTACT_DATE columns.
 *
 * Call from menu: Admin > Data Sync > Sync CC Engagement
 */
function syncConstantContactEngagement() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    ui.alert('❌ Error', 'Member Directory sheet not found.', ui.ButtonSet.OK);
    return;
  }

  // Check authorization
  var token = getConstantContactToken_();
  if (!token) {
    ui.alert('⚠️ Not Authorized',
      'Constant Contact is not connected.\n\n' +
      'Go to Admin > Data Sync > Constant Contact Setup to configure your API credentials, ' +
      'then use "Authorize Constant Contact" to connect your account.',
      ui.ButtonSet.OK);
    return;
  }

  ss.toast('Fetching contacts from Constant Contact...', '📧 CC Sync', 10);

  // Step 1: Fetch all CC contacts (email -> contact_id map)
  var emailToContactId = fetchCCContacts_();
  var ccContactCount = Object.keys(emailToContactId).length;

  if (ccContactCount === 0) {
    ui.alert('⚠️ No Contacts Found',
      'No contacts were returned from Constant Contact.\n\n' +
      'Make sure your CC account has contacts and your API key has the correct permissions.',
      ui.ButtonSet.OK);
    return;
  }

  ss.toast('Found ' + ccContactCount + ' CC contacts. Matching to members...', '📧 CC Sync', 10);

  // Step 2: Read Member Directory emails
  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No members found in Member Directory.', '⚠️ CC Sync', 3);
    return;
  }

  var memberData = memberSheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.RECENT_CONTACT_DATE).getValues();

  // Step 3: Match members to CC contacts and fetch engagement
  var openRateUpdates = [];
  var contactDateUpdates = [];
  var matchCount = 0;
  var processedCount = 0;

  for (var i = 0; i < memberData.length; i++) {
    var memberEmail = (memberData[i][MEMBER_COLS.EMAIL - 1] || '').toString().toLowerCase().trim();
    var contactId = memberEmail ? emailToContactId[memberEmail] : null;

    if (contactId) {
      matchCount++;

      // Rate limit: pause between API calls
      if (processedCount > 0 && processedCount % CC_CONFIG.RATE_LIMIT_PER_SECOND === 0) {
        Utilities.sleep(1000);
      }

      var engagement = fetchCCContactEngagement_(contactId);
      processedCount++;

      if (engagement) {
        openRateUpdates.push([engagement.openRate]);
        contactDateUpdates.push([engagement.lastActivityDate || '']);
      } else {
        openRateUpdates.push([numericField_(memberData[i][MEMBER_COLS.OPEN_RATE - 1])]);
        contactDateUpdates.push([memberData[i][MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '']);
      }

      // Progress update every 25 contacts
      if (processedCount % 25 === 0) {
        ss.toast('Processing... ' + processedCount + '/' + matchCount + ' contacts', '📧 CC Sync', 5);
      }
    } else {
      // No CC match — preserve existing values
      openRateUpdates.push([numericField_(memberData[i][MEMBER_COLS.OPEN_RATE - 1])]);
      contactDateUpdates.push([memberData[i][MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '']);
    }
  }

  // Step 4: Write updates to Member Directory
  if (openRateUpdates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.OPEN_RATE, openRateUpdates.length, 1).setValues(openRateUpdates);
    memberSheet.getRange(2, MEMBER_COLS.RECENT_CONTACT_DATE, contactDateUpdates.length, 1).setValues(contactDateUpdates);
  }

  // Step 5: Report results
  var summary = 'Constant Contact Sync Complete!\n\n' +
    '• CC contacts found: ' + ccContactCount + '\n' +
    '• Members matched by email: ' + matchCount + '\n' +
    '• Engagement data updated: ' + processedCount + '\n' +
    '• Members without CC match: ' + (memberData.length - matchCount);

  ui.alert('✅ CC Sync Complete', summary, ui.ButtonSet.OK);
  Logger.log('CC sync: ' + matchCount + ' matches out of ' + memberData.length + ' members');
}

/**
 * Shows the current Constant Contact connection status.
 * Displays API key info, token status, and last sync details.
 */
function showConstantContactStatus() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var token = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  var refreshToken = props.getProperty(CC_CONFIG.PROP_REFRESH_TOKEN);

  var status = '📧 Constant Contact Connection Status\n\n';

  if (!apiKey) {
    status += '❌ API Key: Not configured\n';
    status += '❌ Authorization: Not connected\n\n';
    status += 'To get started:\n';
    status += '1. Admin > Data Sync > CC Setup: API Credentials\n';
    status += '2. Admin > Data Sync > CC Authorize Account';
  } else {
    status += '✅ API Key: Configured (ends in ...' + apiKey.slice(-4) + ')\n';

    if (token && expiry) {
      var expiryDate = new Date(expiry);
      if (expiryDate > new Date()) {
        status += '✅ Access Token: Valid (expires ' + expiryDate.toLocaleString() + ')\n';
      } else {
        status += '⚠️ Access Token: Expired\n';
      }
    } else {
      status += '❌ Access Token: Not authorized\n';
    }

    status += refreshToken ? '✅ Refresh Token: Available\n' : '❌ Refresh Token: Missing\n';
    status += '\nReady to sync: ' + (token && refreshToken ? 'Yes' : 'No — re-authorize');
  }

  ui.alert('Constant Contact Status', status, ui.ButtonSet.OK);
}

/**
 * Removes all stored Constant Contact credentials and tokens.
 * Use this to disconnect CC or before re-configuring with a different account.
 */
function disconnectConstantContact() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ Disconnect Constant Contact',
    'This will remove all stored API credentials and tokens.\n\n' +
    'You will need to re-run the setup to reconnect.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(CC_CONFIG.PROP_API_KEY);
  props.deleteProperty(CC_CONFIG.PROP_API_SECRET);
  props.deleteProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  props.deleteProperty(CC_CONFIG.PROP_REFRESH_TOKEN);
  props.deleteProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Constant Contact disconnected. All credentials removed.',
    '🔌 Disconnected', 5
  );
  Logger.log('Constant Contact credentials removed');
}

// ============================================================================
// v4.11.0: EDUCATIONAL RESOURCES HUB
// ============================================================================

/**
 * API function to get resources list for web app.
 * Reads from the 📚 Resources sheet, returns visible items sorted by Sort Order.
 * @param {string} [audience] - Filter by audience: 'All', 'Members', 'Stewards'. Defaults to 'All'.
 * @returns {Array} Resource objects
 */
function getWebAppResourcesList(audience) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);

    // Auto-create if missing
    if (!sheet) {
      if (typeof createResourcesSheet === 'function') {
        sheet = createResourcesSheet(ss);
      }
      if (!sheet) return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, RESOURCES_COLS.ADDED_BY).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var visible = String(row[RESOURCES_COLS.VISIBLE - 1] || '').toLowerCase();
      if (visible !== 'yes') return null;

      var rowAudience = row[RESOURCES_COLS.AUDIENCE - 1] || 'All';
      if (audience && audience !== 'All' && rowAudience !== 'All' && rowAudience !== audience) return null;

      var dateAdded = row[RESOURCES_COLS.DATE_ADDED - 1];

      return {
        id: row[RESOURCES_COLS.RESOURCE_ID - 1] || '',
        title: row[RESOURCES_COLS.TITLE - 1] || '',
        category: row[RESOURCES_COLS.CATEGORY - 1] || 'General',
        summary: row[RESOURCES_COLS.SUMMARY - 1] || '',
        content: row[RESOURCES_COLS.CONTENT - 1] || '',
        url: row[RESOURCES_COLS.URL - 1] || '',
        icon: row[RESOURCES_COLS.ICON - 1] || '📄',
        sortOrder: row[RESOURCES_COLS.SORT_ORDER - 1] || 999,
        audience: rowAudience,
        dateAdded: dateAdded instanceof Date ? Utilities.formatDate(dateAdded, tz, 'MMM d, yyyy') : (dateAdded || '')
      };
    }).filter(function(r) { return r !== null; });

    // Sort by sortOrder
    result.sort(function(a, b) { return (a.sortOrder || 999) - (b.sortOrder || 999); });

    return result;
  } catch (e) {
    Logger.log('getWebAppResourcesList error: ' + e.toString());
    return [];
  }
}

/**
 * Returns the educational resources hub HTML page.
 * Warm, welcoming design with categorized content cards.
 * Content is fully dynamic — reads from 📚 Resources sheet.
 */
function getWebAppResourcesHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  // Read org name from config
  var orgName = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && configSheet.getLastRow() >= 3) {
      orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue() || '';
    }
  } catch (_e) { /* ignore */ }

  return '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Know Your Rights</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:wght@600;700;800&display=swap" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:"DM Sans",sans-serif;background:#fafaf9;color:#1c1917;min-height:100vh;padding-bottom:80px}' +

    // Header — warm gradient, not cold purple
    '.header{background:linear-gradient(135deg,#1e3a5f 0%,#2d5a87 100%);color:white;padding:24px 20px 28px;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-family:"Fraunces",serif;font-size:clamp(22px,5vw,28px);font-weight:700;margin-bottom:6px;letter-spacing:-0.02em}' +
    '.header .subtitle{font-size:14px;opacity:0.85;line-height:1.4}' +

    // Search
    '.search-wrap{padding:0 16px;margin-top:-16px;position:relative;z-index:2}' +
    '.search-bar{width:100%;padding:14px 14px 14px 44px;border:2px solid #e7e5e4;border-radius:14px;font-size:15px;font-family:"DM Sans",sans-serif;background:white;box-shadow:0 2px 8px rgba(0,0,0,0.06);transition:border-color 0.2s}' +
    '.search-bar:focus{outline:none;border-color:#2d5a87;box-shadow:0 2px 12px rgba(45,90,135,0.15)}' +
    '.search-icon{position:absolute;left:30px;top:50%;transform:translateY(-50%);color:#a8a29e;font-size:20px}' +

    // Category pills
    '.cat-pills{display:flex;gap:8px;overflow-x:auto;padding:20px 16px 8px;-webkit-overflow-scrolling:touch}' +
    '.cat-pill{flex-shrink:0;padding:8px 16px;border-radius:24px;font-size:13px;font-weight:600;border:2px solid #e7e5e4;cursor:pointer;background:white;color:#57534e;transition:all 0.2s;white-space:nowrap}' +
    '.cat-pill.active{background:#1e3a5f;color:white;border-color:#1e3a5f}' +
    '.cat-pill:hover{border-color:#1e3a5f}' +

    // Resource cards
    '.resources-list{padding:12px 16px}' +
    '.res-card{background:white;border-radius:16px;padding:20px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;cursor:pointer;transition:all 0.25s}' +
    '.res-card:active{transform:scale(0.99)}' +
    '.res-card.expanded{box-shadow:0 4px 16px rgba(0,0,0,0.1)}' +
    '.res-icon{font-size:28px;margin-bottom:10px;display:block}' +
    '.res-title{font-family:"Fraunces",serif;font-size:17px;font-weight:700;color:#1c1917;margin-bottom:4px;line-height:1.3}' +
    '.res-category{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#2d5a87;margin-bottom:8px}' +
    '.res-summary{font-size:14px;color:#57534e;line-height:1.5}' +
    '.res-content{display:none;margin-top:16px;padding-top:16px;border-top:1px solid #f5f5f4;font-size:14px;color:#44403c;line-height:1.7;white-space:pre-line}' +
    '.res-card.expanded .res-content{display:block}' +
    '.res-link{display:inline-flex;align-items:center;gap:6px;margin-top:12px;padding:8px 16px;background:#eff6ff;color:#1e3a5f;border-radius:8px;font-size:13px;font-weight:600;text-decoration:none;transition:background 0.2s}' +
    '.res-link:hover{background:#dbeafe}' +
    '.res-meta{font-size:11px;color:#a8a29e;margin-top:8px}' +
    '.expand-hint{font-size:12px;color:#a8a29e;display:flex;align-items:center;gap:4px;margin-top:8px}' +
    '.expand-hint i{font-size:16px;transition:transform 0.2s}' +
    '.res-card.expanded .expand-hint i{transform:rotate(180deg)}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 24px}' +
    '.empty-icon{font-size:56px;margin-bottom:16px;opacity:0.6}' +
    '.empty-text{font-size:16px;color:#78716c;line-height:1.5}' +
    '.empty-sub{font-size:13px;color:#a8a29e;margin-top:8px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#78716c}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:28px;height:28px;border:3px solid #e7e5e4;border-top-color:#1e3a5f;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#1e3a5f}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h1>📚 Know Your Rights</h1>' +
    '<div class="subtitle">Your guide to the union contract, grievance process, and workplace protections' + (orgName ? ' at ' + escapeHtml(orgName) : '') + '</div>' +
    '</div>' +

    '<div class="search-wrap">' +
    '<i class="material-icons search-icon">search</i>' +
    '<input type="text" class="search-bar" id="searchInput" placeholder="Search articles, rights, FAQ..." oninput="filterResources()">' +
    '</div>' +

    '<div class="cat-pills" id="catPills"></div>' +

    '<div class="resources-list" id="resourcesList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:12px">Loading resources...</div></div>' +
    '</div>' +

    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=checkin">' +
    '<span class="nav-icon">✅</span>Check In</a>' +
    '<a class="nav-item active" href="' + escapeHtml(baseUrl) + '?page=resources">' +
    '<span class="nav-icon">📚</span>Learn</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=selfservice">' +
    '<span class="nav-icon">👤</span>My Info</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allResources=[];var currentCat="all";' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function filterResources(){' +
    '  var q=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allResources.filter(function(r){' +
    '    var matchesCat=currentCat==="all"||r.category===currentCat;' +
    '    var matchesQ=!q||q.length<2||r.title.toLowerCase().indexOf(q)>=0||r.summary.toLowerCase().indexOf(q)>=0||(r.content||"").toLowerCase().indexOf(q)>=0||r.category.toLowerCase().indexOf(q)>=0;' +
    '    return matchesCat&&matchesQ;' +
    '  });' +
    '  renderList(filtered);' +
    '}' +

    'function setCat(cat,btn){' +
    '  currentCat=cat;' +
    '  document.querySelectorAll(".cat-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterResources();' +
    '}' +

    'function renderList(items){' +
    '  var c=document.getElementById("resourcesList");' +
    '  if(!items||items.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📖</div><div class=\\"empty-text\\">No resources found</div><div class=\\"empty-sub\\">Try a different search or category</div></div>";' +
    '    return;' +
    '  }' +
    '  c.innerHTML=items.map(function(r){' +
    '    var hasContent=r.content&&r.content.trim().length>0;' +
    '    var hasUrl=r.url&&r.url.trim().length>0;' +
    '    return "<div class=\\"res-card\\"' + ('" onclick=\\"toggleCard(this)\\"') + '>"' +
    '      +"<span class=\\"res-icon\\">"+escapeHtml(r.icon)+"</span>"' +
    '      +"<div class=\\"res-category\\">"+escapeHtml(r.category)+"</div>"' +
    '      +"<div class=\\"res-title\\">"+escapeHtml(r.title)+"</div>"' +
    '      +"<div class=\\"res-summary\\">"+escapeHtml(r.summary)+"</div>"' +
    '      +(hasContent?"<div class=\\"expand-hint\\"><i class=\\"material-icons\\">expand_more</i>Tap to read more</div>":"")' +
    '      +(hasContent?"<div class=\\"res-content\\">"+escapeHtml(r.content).replace(/\\\\n/g,"\\n")+"</div>":"")' +
    '      +(hasUrl?"<a class=\\"res-link\\" href=\\""+escapeHtml(r.url)+"\\" target=\\"_blank\\" onclick=\\"event.stopPropagation()\\"><i class=\\"material-icons\\" style=\\"font-size:16px\\">open_in_new</i>Open Resource</a>":"")' +
    '      +(r.dateAdded?"<div class=\\"res-meta\\">Added "+escapeHtml(r.dateAdded)+"</div>":"")' +
    '      +"</div>";' +
    '  }).join("");' +
    '}' +

    'function buildCatPills(items){' +
    '  var cats={};items.forEach(function(r){cats[r.category]=(cats[r.category]||0)+1});' +
    '  var el=document.getElementById("catPills");' +
    '  var html="<button class=\\"cat-pill active\\" onclick=\\"setCat(\\x27all\\x27,this)\\">All ("+items.length+")</button>";' +
    '  Object.keys(cats).sort().forEach(function(c){' +
    '    html+="<button class=\\"cat-pill\\" onclick=\\"setCat(\\x27"+c+"\\x27,this)\\">"+escapeHtml(c)+" ("+cats[c]+")</button>";' +
    '  });' +
    '  el.innerHTML=html;' +
    '}' +

    'google.script.run.withSuccessHandler(function(data){' +
    '  allResources=data||[];' +
    '  buildCatPills(allResources);' +
    '  renderList(allResources);' +
    '}).withFailureHandler(function(err){' +
    '  document.getElementById("resourcesList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error loading resources</div><div class=\\"empty-sub\\">"+String(err||"")+"</div></div>";' +
    '}).getWebAppResourcesList();' +
    '</script>' +

    '</body></html>';
}

// ============================================================================
// v4.11.0: MEETING CHECK-IN WEB ROUTE
// ============================================================================

/**
 * Returns a standalone web page for meeting check-in.
 * Reuses the existing check-in logic from 14_MeetingCheckIn.gs.
 * Mobile-optimized for kiosk mode (shared device) or individual phone check-in.
 */
function getWebAppCheckInHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Meeting Check-In</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:wght@600;700&display=swap" rel="stylesheet">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:"DM Sans",sans-serif;background:#fafaf9;min-height:100vh;padding-bottom:80px}' +

    '.header{background:linear-gradient(135deg,#065f46 0%,#047857 100%);color:white;padding:24px 20px 28px;text-align:center;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-family:"Fraunces",serif;font-size:clamp(22px,5vw,28px);font-weight:700;margin-bottom:4px}' +
    '.header .subtitle{font-size:14px;opacity:0.85}' +

    '.container{max-width:480px;margin:0 auto;padding:20px 16px}' +

    '.card{background:white;border-radius:16px;padding:24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;margin-bottom:16px}' +
    '.card h2{font-family:"Fraunces",serif;font-size:18px;color:#065f46;margin-bottom:16px;text-align:center}' +

    '.field{margin-bottom:18px}' +
    '.field label{display:block;margin-bottom:8px;font-weight:600;color:#1c1917;font-size:14px}' +
    '.field input,.field select{width:100%;padding:14px;border:2px solid #e7e5e4;border-radius:12px;font-size:16px;font-family:"DM Sans",sans-serif;transition:border-color 0.2s}' +
    '.field input:focus,.field select:focus{outline:none;border-color:#047857;box-shadow:0 0 0 3px rgba(4,120,87,0.1)}' +

    '.btn{width:100%;padding:16px;border:none;border-radius:12px;font-size:17px;font-weight:700;font-family:"DM Sans",sans-serif;cursor:pointer;transition:all 0.2s}' +
    '.btn-checkin{background:#047857;color:white}' +
    '.btn-checkin:hover{background:#065f46}' +
    '.btn-checkin:disabled{background:#d6d3d1;cursor:not-allowed}' +

    '.error-msg{color:#dc2626;font-size:14px;text-align:center;padding:12px;background:#fef2f2;border-radius:10px;margin-top:12px;display:none}' +

    '.success-banner{text-align:center;padding:28px;background:#d1fae5;border-radius:16px;margin-bottom:16px;display:none}' +
    '.success-banner .checkmark{font-size:56px;margin-bottom:8px}' +
    '.success-banner .name{font-family:"Fraunces",serif;font-size:22px;font-weight:700;color:#065f46;margin-bottom:4px}' +
    '.success-banner .msg{color:#047857;font-size:14px}' +

    '.attendee-count{text-align:center;padding:14px;background:#ecfdf5;border-radius:12px;color:#047857;font-weight:600;margin-top:12px;font-size:14px}' +

    '.no-meetings{text-align:center;padding:40px 20px;color:#78716c}' +
    '.no-meetings-icon{font-size:48px;margin-bottom:12px;opacity:0.5}' +

    '.meeting-option{padding:14px;border:2px solid #e7e5e4;border-radius:12px;margin-bottom:10px;cursor:pointer;transition:all 0.2s}' +
    '.meeting-option:hover{border-color:#047857}' +
    '.meeting-option.selected{border-color:#047857;background:#ecfdf5}' +
    '.meeting-name{font-weight:600;color:#1c1917;margin-bottom:2px}' +
    '.meeting-time{font-size:13px;color:#78716c}' +

    '.loading{text-align:center;padding:40px;color:#78716c}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:28px;height:28px;border:3px solid #e7e5e4;border-top-color:#047857;border-radius:50%;animation:spin 0.8s linear infinite}' +

    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#065f46}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h1>✅ Meeting Check-In</h1>' +
    '<div class="subtitle">Enter your email and PIN to check in</div>' +
    '</div>' +

    '<div class="container">' +

    '<div id="successBanner" class="success-banner">' +
    '<div class="checkmark">✓</div>' +
    '<div class="name" id="successName"></div>' +
    '<div class="msg">Checked in successfully!</div>' +
    '</div>' +

    '<div class="card" id="meetingCard">' +
    '<h2>Today\'s Meetings</h2>' +
    '<div id="meetingList"><div class="loading"><div class="spinner"></div><div style="margin-top:12px">Loading meetings...</div></div></div>' +
    '</div>' +

    '<div class="card" id="checkinForm" style="display:none">' +
    '<h2>Check In</h2>' +
    '<div class="field"><label>Email Address</label>' +
    '<input type="email" id="email" placeholder="your.email@example.com" autocomplete="email"></div>' +
    '<div class="field"><label>PIN</label>' +
    '<input type="password" id="pin" inputmode="numeric" maxlength="8" placeholder="Enter your PIN"></div>' +
    '<button class="btn btn-checkin" id="checkinBtn" onclick="doCheckIn()">Check In</button>' +
    '<div class="error-msg" id="errorMsg"></div>' +
    '<div class="attendee-count" id="attendeeCount" style="display:none"></div>' +
    '</div>' +

    '</div>' +

    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item active" href="' + escapeHtml(baseUrl) + '?page=checkin">' +
    '<span class="nav-icon">✅</span>Check In</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=resources">' +
    '<span class="nav-icon">📚</span>Learn</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=selfservice">' +
    '<span class="nav-icon">👤</span>My Info</a>' +
    '<a class="nav-item" href="' + escapeHtml(baseUrl) + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var selectedMeetingId=null;' +

    'google.script.run.withSuccessHandler(function(meetings){' +
    '  var el=document.getElementById("meetingList");' +
    '  if(!meetings||meetings.length===0){' +
    '    el.innerHTML="<div class=\\"no-meetings\\"><div class=\\"no-meetings-icon\\">📅</div><div>No active meetings right now</div><div style=\\"font-size:13px;margin-top:8px\\">Check-in opens on the day of a scheduled meeting</div></div>";' +
    '    return;' +
    '  }' +
    '  if(meetings.length===1){' +
    '    selectedMeetingId=meetings[0].id;' +
    '    el.innerHTML="<div class=\\"meeting-option selected\\"><div class=\\"meeting-name\\">"+meetings[0].name+"</div><div class=\\"meeting-time\\">"+meetings[0].time+"</div></div>";' +
    '    document.getElementById("checkinForm").style.display="block";' +
    '  }else{' +
    '    var html="<div style=\\"font-size:14px;color:#78716c;margin-bottom:12px\\">Select your meeting:</div>";' +
    '    meetings.forEach(function(m){' +
    '      html+="<div class=\\"meeting-option\\" onclick=\\"selectMeeting(\\x27"+m.id+"\\x27,this)\\"><div class=\\"meeting-name\\">"+m.name+"</div><div class=\\"meeting-time\\">"+m.time+"</div></div>";' +
    '    });' +
    '    el.innerHTML=html;' +
    '  }' +
    '}).withFailureHandler(function(err){' +
    '  document.getElementById("meetingList").innerHTML="<div class=\\"no-meetings\\"><div class=\\"no-meetings-icon\\">⚠️</div><div>Error loading meetings</div></div>";' +
    '}).getCheckInEligibleMeetings();' +

    'function selectMeeting(id,el){' +
    '  selectedMeetingId=id;' +
    '  document.querySelectorAll(".meeting-option").forEach(function(o){o.classList.remove("selected")});' +
    '  el.classList.add("selected");' +
    '  document.getElementById("checkinForm").style.display="block";' +
    '  document.getElementById("email").focus();' +
    '}' +

    'function doCheckIn(){' +
    '  var email=document.getElementById("email").value.trim();' +
    '  var pin=document.getElementById("pin").value.trim();' +
    '  var errEl=document.getElementById("errorMsg");' +
    '  errEl.style.display="none";' +
    '  if(!selectedMeetingId){errEl.textContent="Please select a meeting";errEl.style.display="block";return}' +
    '  if(!email){errEl.textContent="Please enter your email";errEl.style.display="block";return}' +
    '  if(!pin||pin.length<4){errEl.textContent="Please enter your PIN (4-8 digits)";errEl.style.display="block";return}' +
    '  var btn=document.getElementById("checkinBtn");' +
    '  btn.disabled=true;btn.textContent="Checking in...";' +
    '  google.script.run.withSuccessHandler(function(result){' +
    '    btn.disabled=false;btn.textContent="Check In";' +
    '    if(result&&result.success){' +
    '      document.getElementById("successName").textContent=result.memberName||"Member";' +
    '      document.getElementById("successBanner").style.display="block";' +
    '      if(result.attendeeCount){var ac=document.getElementById("attendeeCount");ac.textContent="✓ "+result.attendeeCount+" checked in so far";ac.style.display="block"}' +
    '      document.getElementById("email").value="";document.getElementById("pin").value="";' +
    '      setTimeout(function(){document.getElementById("successBanner").style.display="none"},4000);' +
    '    }else{' +
    '      errEl.textContent=(result&&result.message)||"Check-in failed. Verify your email and PIN.";errEl.style.display="block";' +
    '    }' +
    '  }).withFailureHandler(function(err){' +
    '    btn.disabled=false;btn.textContent="Check In";' +
    '    errEl.textContent="Error: "+String(err||"Unknown");errEl.style.display="block";' +
    '  }).processMeetingCheckIn(selectedMeetingId,email,pin);' +
    '}' +

    'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter")doCheckIn()});' +
    '</script>' +

    '</body></html>';
}


// ============================================================================
// NOTIFICATIONS API (v4.12.0)
// ============================================================================
// Steward-composed, member-dismissable notifications.
// Data lives in 📢 Notifications sheet. Shown in member web view.
// Persist until Expires date set by steward OR dismissed by member.
// Steward composes via separate form accessible from steward dashboard.
// ============================================================================

/**
 * Get active notifications for a user.
 * Filters: Status=Active, not expired, not dismissed by this user.
 * Matches: direct email recipient, "All Members", "All Stewards", "Everyone"
 * @param {string} userEmail — the logged-in user's email
 * @param {string} [userRole] — "steward" or "member" for audience matching
 * @returns {Object[]} array of notification objects
 */
function getWebAppNotifications(userEmail, userRole) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var C = NOTIFICATIONS_COLS;
    var now = new Date();
    var results = [];
    userEmail = (userEmail || '').toLowerCase().trim();
    userRole = (userRole || 'member').toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Status must be Active
      var status = String(row[C.STATUS - 1] || '').trim();
      if (status !== 'Active') continue;

      // Check expiry
      var expires = row[C.EXPIRES_DATE - 1];
      if (expires && expires instanceof Date && expires < now) continue;
      if (expires && typeof expires === 'string' && expires.trim()) {
        var parsed = new Date(expires);
        if (!isNaN(parsed.getTime()) && parsed < now) continue;
      }

      // Check dismissed
      var dismissed = String(row[C.DISMISSED_BY - 1] || '');
      var dismissedList = dismissed.split(',').map(function(e) { return e.trim().toLowerCase(); });
      if (userEmail && dismissedList.indexOf(userEmail) !== -1) continue;

      // Check recipient match
      var recipient = String(row[C.RECIPIENT - 1] || '').trim().toLowerCase();
      var matches = false;
      if (recipient === 'everyone') matches = true;
      else if (recipient === 'all members') matches = true;
      else if (recipient === 'all stewards' && userRole === 'steward') matches = true;
      else if (recipient === userEmail) matches = true;
      if (!matches) continue;

      results.push({
        id: String(row[C.NOTIFICATION_ID - 1] || ''),
        recipient: String(row[C.RECIPIENT - 1] || ''),
        type: String(row[C.TYPE - 1] || 'System'),
        title: String(row[C.TITLE - 1] || ''),
        message: String(row[C.MESSAGE - 1] || ''),
        priority: String(row[C.PRIORITY - 1] || 'Normal'),
        sentBy: String(row[C.SENT_BY_NAME - 1] || ''),
        createdDate: row[C.CREATED_DATE - 1] ? Utilities.formatDate(
          row[C.CREATED_DATE - 1] instanceof Date ? row[C.CREATED_DATE - 1] : new Date(row[C.CREATED_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        expiresDate: row[C.EXPIRES_DATE - 1] ? Utilities.formatDate(
          row[C.EXPIRES_DATE - 1] instanceof Date ? row[C.EXPIRES_DATE - 1] : new Date(row[C.EXPIRES_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        rowIndex: i + 1
      });
    }

    // Sort: Urgent first, then newest first
    results.sort(function(a, b) {
      if (a.priority === 'Urgent' && b.priority !== 'Urgent') return -1;
      if (b.priority === 'Urgent' && a.priority !== 'Urgent') return 1;
      return 0; // preserve sheet order (newest at bottom, but we'll reverse)
    });
    results.reverse();

    return results;
  } catch (e) {
    logError_('getWebAppNotifications', e);
    return [];
  }
}


/**
 * Dismiss a notification for a specific user.
 * Appends user's email to the Dismissed_By column (comma-separated).
 * @param {string} notificationId — e.g. "NOTIF-001"
 * @param {string} userEmail — the user dismissing
 * @returns {Object} { success: boolean, message: string }
 */
function dismissWebAppNotification(notificationId, userEmail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return { success: false, message: 'Notifications sheet not found' };

    var data = sheet.getDataRange().getValues();
    var C = NOTIFICATIONS_COLS;
    userEmail = (userEmail || '').toLowerCase().trim();
    if (!userEmail) return { success: false, message: 'No email provided' };

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][C.NOTIFICATION_ID - 1] || '').trim();
      if (rowId === notificationId) {
        var existing = String(data[i][C.DISMISSED_BY - 1] || '').trim();
        var emails = existing ? existing.split(',').map(function(e) { return e.trim().toLowerCase(); }) : [];
        if (emails.indexOf(userEmail) === -1) {
          emails.push(userEmail);
          sheet.getRange(i + 1, C.DISMISSED_BY).setValue(emails.join(', '));
        }
        return { success: true, message: 'Notification dismissed' };
      }
    }

    return { success: false, message: 'Notification not found' };
  } catch (e) {
    logError_('dismissWebAppNotification', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Send a new notification (steward form submission).
 * Creates a new row in the Notifications sheet.
 * @param {Object} data — { recipient, type, title, message, priority, expiresDate }
 * @returns {Object} { success: boolean, notificationId: string, message: string }
 */
function sendWebAppNotification(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) {
      sheet = createNotificationsSheet(ss);
    }

    // Validate required fields
    if (!data.title || !data.message) {
      return { success: false, message: 'Title and message are required' };
    }
    if (!data.recipient) {
      return { success: false, message: 'Recipient is required' };
    }

    // Get steward info from session
    var stewardEmail = Session.getActiveUser().getEmail() || 'unknown';
    var stewardName = data.senderName || stewardEmail.split('@')[0];

    // Generate next ID
    var allData = sheet.getDataRange().getValues();
    var maxNum = 0;
    var C = NOTIFICATIONS_COLS;
    for (var i = 1; i < allData.length; i++) {
      var existId = String(allData[i][C.NOTIFICATION_ID - 1] || '');
      var match = existId.match(/NOTIF-(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    var nextId = 'NOTIF-' + String(maxNum + 1).padStart(3, '0');

    var tz = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Build row (must match NOTIFICATIONS_HEADER_MAP_ order)
    var newRow = [
      nextId,
      data.recipient || 'All Members',
      data.type || 'Steward Message',
      data.title,
      data.message,
      data.priority || 'Normal',
      stewardEmail,
      stewardName,
      today,
      data.expiresDate || '',   // blank = no expiry
      '',                        // Dismissed_By starts empty
      'Active'
    ];

    sheet.appendRow(newRow);

    return { success: true, notificationId: nextId, message: 'Notification sent' };
  } catch (e) {
    logError_('sendWebAppNotification', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Get member list for notification recipient picker (steward use).
 * Returns simplified list: name + email for dropdown.
 * @returns {Object[]} [{ name, email }]
 */
function getNotificationRecipientList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = MEMBER_COLS;
    var results = [];
    // Preset groups first
    results.push({ name: '📢 All Members', email: 'All Members' });
    results.push({ name: '🛡️ All Stewards', email: 'All Stewards' });
    results.push({ name: '🌐 Everyone', email: 'Everyone' });

    for (var i = 1; i < data.length; i++) {
      var firstName = String(data[i][C.FIRST_NAME - 1] || '').trim();
      var lastName = String(data[i][C.LAST_NAME - 1] || '').trim();
      var email = String(data[i][C.EMAIL - 1] || '').trim();
      if (!email) continue;
      var fullName = (firstName + ' ' + lastName).trim();
      results.push({ name: fullName || email, email: email });
    }

    return results;
  } catch (e) {
    logError_('getNotificationRecipientList', e);
    return [];
  }
}


/**
 * Get full member list with directory columns for filtering/sorting.
 * Used by steward notification compose form recipient picker.
 * Returns: name, email, location, department, jobTitle for each member.
 * @returns {Object[]}
 */
function getNotificationRecipientListFull() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = MEMBER_COLS;
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var firstName = String(data[i][C.FIRST_NAME - 1] || '').trim();
      var lastName = String(data[i][C.LAST_NAME - 1] || '').trim();
      var email = String(data[i][C.EMAIL - 1] || '').trim();
      if (!email) continue;

      var fullName = (firstName + ' ' + lastName).trim();
      var location = '';
      var department = '';
      var jobTitle = '';

      if (C.WORK_LOCATION) location = String(data[i][C.WORK_LOCATION - 1] || '').trim();
      if (C.DEPARTMENT) department = String(data[i][C.DEPARTMENT - 1] || '').trim();
      if (C.JOB_TITLE) jobTitle = String(data[i][C.JOB_TITLE - 1] || '').trim();

      results.push({
        name: fullName || email,
        email: email,
        location: location,
        department: department,
        jobTitle: jobTitle
      });
    }

    results.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return results;
  } catch (e) {
    logError_('getNotificationRecipientListFull', e);
    return [];
  }
}


// ============================================================================
// NOTIFICATIONS WEB PAGE (v4.12.0)
// ============================================================================
// ?page=notifications - dual-role page
// Members: view + dismiss notifications
// Stewards: inline compose form + recipient picker + view/manage
// Recipient picker: groups + individuals, filterable by location/dept/title
// ============================================================================

/**
 * Generate notifications page HTML.
 * Detects role via checkWebAppAuthorization - stewards get compose form.
 * @returns {string} full HTML page
 */
function getWebAppNotificationsHtml() {
  var isSteward = false;
  var userEmail = '';
  var userName = '';
  try {
    var authResult = checkWebAppAuthorization('steward');
    isSteward = authResult.isAuthorized;
    userEmail = authResult.email || Session.getActiveUser().getEmail() || '';
    userName = userEmail.split('@')[0] || '';
  } catch (authErr) {
    try { userEmail = Session.getActiveUser().getEmail() || ''; } catch (e2) { userEmail = ''; }
  }

  var userRole = isSteward ? 'steward' : 'member';
  var notifications = getWebAppNotifications(userEmail, userRole);
  var notifJson = JSON.stringify(notifications || []);
  var recipientJson = isSteward ? JSON.stringify(getNotificationRecipientListFull() || []) : '[]';

  var p = [];
  p.push('<!DOCTYPE html><html><head>');
  p.push('<meta charset="utf-8">');
  p.push('<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">');
  p.push('<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700;9..144,800&display=swap" rel="stylesheet">');
  p.push('<title>Notifications | MassAbility DDS</title>');

  // ── CSS ──
  p.push('<style>');
  p.push('*{box-sizing:border-box;margin:0;padding:0}');
  p.push('body{font-family:"DM Sans",sans-serif;background:#fafaf9;min-height:100vh;padding-bottom:20px;color:#1c1917}');
  p.push('.hero{background:linear-gradient(145deg,#92400e,#b45309);color:#fff;padding:26px 20px 34px;position:relative;overflow:hidden}');
  p.push('.hero h1{font-family:"Fraunces",serif;font-size:clamp(22px,5vw,28px);font-weight:700;letter-spacing:-0.02em;margin-bottom:5px}');
  p.push('.hero .sub{font-size:14px;opacity:0.85}');
  p.push('.hero .curve{position:absolute;bottom:-20px;left:-20px;right:-20px;height:40px;background:#fafaf9;border-radius:50% 50% 0 0}');
  p.push('.badge{display:inline-flex;align-items:center;gap:6px;padding:4px 12px;background:rgba(255,255,255,0.15);border-radius:20px;font-size:12px;font-weight:600;margin-top:8px}');
  p.push('.container{max-width:600px;margin:0 auto;padding:0 16px}');
  // Compose
  p.push('.compose{background:#fff;border-radius:16px;padding:20px;margin-top:-12px;position:relative;z-index:2;box-shadow:0 2px 12px rgba(0,0,0,0.08);border:1px solid #f5f5f4;margin-bottom:20px}');
  p.push('.compose h2{font-family:"Fraunces",serif;font-size:18px;color:#92400e;margin-bottom:16px}');
  p.push('.fg{margin-bottom:14px}');
  p.push('.fg label{display:block;font-weight:600;font-size:13px;margin-bottom:6px;color:#44403c}');
  p.push('.fg input,.fg select,.fg textarea{width:100%;padding:12px;border:2px solid #e7e5e4;border-radius:10px;font-size:14px;font-family:"DM Sans",sans-serif;outline:none}');
  p.push('.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:#b45309}');
  p.push('.fg textarea{min-height:80px;resize:vertical}');
  // Tabs
  p.push('.rtabs{display:flex;gap:6px;margin-bottom:10px}');
  p.push('.rtabs button{padding:6px 14px;border-radius:20px;border:2px solid #e7e5e4;background:#fff;font-size:12px;font-weight:600;cursor:pointer;font-family:"DM Sans",sans-serif}');
  p.push('.rtabs button.act{background:#92400e;color:#fff;border-color:#92400e}');
  // Groups
  p.push('.gbtn{padding:8px 14px;border-radius:10px;border:2px solid #e7e5e4;background:#fff;font-size:13px;font-weight:600;cursor:pointer;font-family:"DM Sans",sans-serif;margin:0 6px 6px 0}');
  p.push('.gbtn.sel{background:#fffbeb;border-color:#b45309;color:#92400e}');
  // Filter bar
  p.push('.fbar{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}');
  p.push('.fbar select{padding:8px 10px;border:2px solid #e7e5e4;border-radius:8px;font-size:12px;font-family:"DM Sans",sans-serif;background:#fff;outline:none;min-width:100px}');
  p.push('.fbar input{flex:1;min-width:120px;padding:8px 12px;border:2px solid #e7e5e4;border-radius:8px;font-size:12px;font-family:"DM Sans",sans-serif;outline:none}');
  // Member list
  p.push('.mlist{max-height:240px;overflow-y:auto;border:2px solid #e7e5e4;border-radius:10px;background:#fff}');
  p.push('.mrow{display:flex;align-items:center;gap:10px;padding:10px 12px;border-bottom:1px solid #f5f5f4;cursor:pointer}');
  p.push('.mrow:hover{background:#fffbeb}.mrow.sel{background:#fef3c7}');
  p.push('.mchk{width:18px;height:18px;border-radius:4px;border:2px solid #d6d3d1;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:11px}');
  p.push('.mchk.on{background:#92400e;border-color:#92400e;color:#fff}');
  p.push('.mname{font-weight:600;font-size:13px}.mdtl{font-size:11px;color:#78716c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}');
  p.push('.scnt{font-size:12px;color:#78716c;margin-top:6px}');
  // Send button
  p.push('.btn-send{width:100%;padding:14px;border:none;border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:"DM Sans",sans-serif;background:#92400e;color:#fff;margin-top:16px}');
  p.push('.btn-send:disabled{opacity:0.5;cursor:not-allowed}');
  // Notification cards
  p.push('.slabel{font-size:12px;font-weight:700;color:#a8a29e;text-transform:uppercase;letter-spacing:0.06em;margin:20px 0 10px}');
  p.push('.nc{background:#fff;border-radius:14px;padding:16px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;position:relative;animation:fi 0.3s ease both}');
  p.push('.nc.urg{border-left:4px solid #dc2626}');
  p.push('.nc h3{font-family:"Fraunces",serif;font-size:15px;font-weight:700;color:#1c1917;margin-bottom:4px;line-height:1.3}');
  p.push('.nc .msg{font-size:13px;color:#57534e;line-height:1.5}');
  p.push('.nc .meta{display:flex;gap:12px;margin-top:10px;font-size:11px;color:#a8a29e;flex-wrap:wrap}');
  p.push('.nc .dbtn{position:absolute;top:12px;right:12px;background:none;border:none;font-size:18px;color:#d6d3d1;cursor:pointer;padding:4px;line-height:1}');
  p.push('.nc .dbtn:hover{color:#78716c}');
  p.push('.tbdg{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;text-transform:uppercase;margin-right:4px}');
  p.push('.empty{text-align:center;padding:40px 20px;color:#a8a29e}');
  p.push('.toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);padding:12px 24px;border-radius:12px;font-size:14px;font-weight:600;z-index:100;box-shadow:0 4px 16px rgba(0,0,0,0.15);display:none}');
  p.push('.toast.ok{background:#065f46;color:#fff}.toast.err{background:#dc2626;color:#fff}');
  p.push('@keyframes fi{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}');
  p.push('</style></head><body>');

  // ── Hero ──
  p.push('<div class="hero"><div class="curve"></div>');
  p.push('<h1>&#128276; Notifications</h1>');
  var heroSub = isSteward ? 'Compose and manage member notifications' : 'Messages from your union steward';
  p.push('<div class="sub">' + heroSub + '</div>');
  if (isSteward) {
    p.push('<div class="badge">&#128737;&#65039; Steward &#183; Compose Access</div>');
  }
  p.push('</div>');
  p.push('<div class="container">');

  // ── Steward Compose Form ──
  if (isSteward) {
    p.push('<div class="compose" id="composeForm">');
    p.push('<h2>&#9997;&#65039; Send Notification</h2>');

    // Recipient section
    p.push('<div class="fg">');
    p.push('<label>To</label>');
    p.push('<div class="rtabs" id="rtabs">');
    p.push('<button class="act" onclick="switchTab(0)">Groups</button>');
    p.push('<button onclick="switchTab(1)">Individuals</button>');
    p.push('</div>');

    // Groups tab
    p.push('<div id="grpTab">');
    p.push('<button class="gbtn sel" onclick="selGrp(this,\'All Members\')">&#128226; All Members</button>');
    p.push('<button class="gbtn" onclick="selGrp(this,\'All Stewards\')">&#128737;&#65039; All Stewards</button>');
    p.push('<button class="gbtn" onclick="selGrp(this,\'Everyone\')">&#127760; Everyone</button>');
    p.push('</div>');

    // Individuals tab
    p.push('<div id="indTab" style="display:none">');
    p.push('<div class="fbar">');
    p.push('<input type="text" id="mSearch" placeholder="Search name..." oninput="filterM()">');
    p.push('<select id="fLoc" onchange="filterM()"><option value="">All Locations</option></select>');
    p.push('<select id="fDept" onchange="filterM()"><option value="">All Depts</option></select>');
    p.push('<select id="fTitle" onchange="filterM()"><option value="">All Titles</option></select>');
    p.push('</div>');
    p.push('<div class="mlist" id="mList"></div>');
    p.push('<div class="scnt" id="sCnt">0 selected</div>');
    p.push('</div>');
    p.push('</div>');

    // Type + Priority row
    p.push('<div style="display:flex;gap:10px">');
    p.push('<div class="fg" style="flex:1"><label>Type</label>');
    p.push('<select id="nType"><option>Steward Message</option><option>Announcement</option><option>Deadline</option><option>System</option></select></div>');
    p.push('<div class="fg" style="flex:1"><label>Priority</label>');
    p.push('<select id="nPri"><option>Normal</option><option>Urgent</option></select></div>');
    p.push('</div>');

    // Title, Message, Expires
    p.push('<div class="fg"><label>Title</label><input type="text" id="nTitle" placeholder="Short headline..."></div>');
    p.push('<div class="fg"><label>Message</label><textarea id="nMsg" placeholder="Notification body..."></textarea></div>');
    p.push('<div class="fg"><label>Expires (blank = manual archive only)</label><input type="date" id="nExp"></div>');
    p.push('<button class="btn-send" id="sendBtn" onclick="doSend()">Send Notification</button>');
    p.push('</div>'); // end compose
  }

  // ── Notification List ──
  p.push('<div class="slabel">Your Notifications</div>');
  p.push('<div id="nList"></div>');
  p.push('</div>'); // end container
  p.push('<div id="toast" class="toast"></div>');

  // ── JavaScript ──
  p.push('<script>');
  p.push('var UE="' + userEmail.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '";');
  p.push('var UR="' + userRole + '";');
  p.push('var UN="' + userName.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '";');
  p.push('var IS=' + String(isSteward) + ';');
  p.push('var notifs=' + notifJson + ';');
  p.push('var allM=' + recipientJson + ';');
  p.push('var rMode="groups";');
  p.push('var selGroup="All Members";');
  p.push('var selMems={};');

  // Render notifications
  p.push('function renderN(){');
  p.push('var el=document.getElementById("nList");');
  p.push('if(!notifs||!notifs.length){');
  p.push('el.innerHTML=\'<div class="empty"><div style="font-size:48px;margin-bottom:12px;opacity:0.4">\\ud83d\\udd14</div><div style="font-family:Fraunces,serif;font-size:18px;font-weight:700;color:#78716c;margin-bottom:4px">All caught up!</div><div>No new notifications</div></div>\';');
  p.push('return;}');
  p.push('var tc={"Steward Message":{bg:"#eff6ff",c:"#1e40af"},"Announcement":{bg:"#f0fdf4",c:"#065f46"},"Deadline":{bg:"#fef2f2",c:"#dc2626"},"System":{bg:"#faf5ff",c:"#7c3aed"}};');
  p.push('var h="";');
  p.push('for(var i=0;i<notifs.length;i++){');
  p.push('var n=notifs[i];var t=tc[n.type]||tc.System;var u=n.priority==="Urgent"?" urg":"";');
  p.push('h+=\'<div class="nc\'+u+\'" style="animation-delay:\'+i*0.06+\'s" id="n-\'+n.id+\'">\';');
  p.push('h+=\'<button class="dbtn" onclick="dismissN(\\x27\'+n.id+\'\\x27)">\\u2715</button>\';');
  p.push('h+=\'<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px">\';');
  p.push('h+=\'<div><span class="tbdg" style="background:\'+t.bg+\';color:\'+t.c+\'">\'+n.type+\'</span>\';');
  p.push('if(n.priority==="Urgent")h+=\'<span class="tbdg" style="background:#fef2f2;color:#dc2626">URGENT</span>\';');
  p.push('h+=\'</div><span style="font-size:11px;color:#a8a29e">\'+n.createdDate+\'</span></div>\';');
  p.push('h+=\'<h3>\'+n.title+\'</h3>\';');
  p.push('h+=\'<div class="msg">\'+n.message+\'</div>\';');
  p.push('h+=\'<div class="meta">\';');
  p.push('if(n.sentBy)h+=\'<span>From: \'+n.sentBy+\'</span>\';');
  p.push('if(n.expiresDate)h+=\'<span>Expires: \'+n.expiresDate+\'</span>\';');
  p.push('h+=\'</div></div>\';');
  p.push('}');
  p.push('el.innerHTML=h;}');

  // Dismiss
  p.push('function dismissN(id){');
  p.push('var c=document.getElementById("n-"+id);');
  p.push('if(c){c.style.opacity="0.3";c.style.pointerEvents="none";}');
  p.push('google.script.run.withSuccessHandler(function(r){');
  p.push('if(r&&r.success){if(c)c.remove();');
  p.push('notifs=notifs.filter(function(n){return n.id!==id;});');
  p.push('if(!notifs.length)renderN();showT("Dismissed","ok");}');
  p.push('else{if(c){c.style.opacity="1";c.style.pointerEvents="auto";}');
  p.push('showT((r&&r.message)||"Failed","err");}');
  p.push('}).withFailureHandler(function(err){');
  p.push('if(c){c.style.opacity="1";c.style.pointerEvents="auto";}');
  p.push('showT("Error: "+err,"err");');
  p.push('}).dismissWebAppNotification(id,UE);}');

  // Steward-only functions
  if (isSteward) {
    // Switch tabs
    p.push('function switchTab(idx){');
    p.push('rMode=idx===0?"groups":"individual";');
    p.push('document.getElementById("grpTab").style.display=idx===0?"block":"none";');
    p.push('document.getElementById("indTab").style.display=idx===1?"block":"none";');
    p.push('var btns=document.querySelectorAll("#rtabs button");');
    p.push('btns[0].className=idx===0?"act":"";btns[1].className=idx===1?"act":"";');
    p.push('if(idx===1&&!document.getElementById("mList").innerHTML)renderML();}');

    // Select group
    p.push('function selGrp(btn,grp){');
    p.push('var all=document.querySelectorAll(".gbtn");');
    p.push('for(var i=0;i<all.length;i++)all[i].className="gbtn";');
    p.push('btn.className="gbtn sel";selGroup=grp;}');

    // Build filter dropdowns
    p.push('function buildFilters(){');
    p.push('var locs={},depts={},titles={};');
    p.push('for(var i=0;i<allM.length;i++){');
    p.push('if(allM[i].location)locs[allM[i].location]=1;');
    p.push('if(allM[i].department)depts[allM[i].department]=1;');
    p.push('if(allM[i].jobTitle)titles[allM[i].jobTitle]=1;}');
    p.push('var addOpts=function(selId,obj){var s=document.getElementById(selId);var ks=Object.keys(obj).sort();');
    p.push('for(var j=0;j<ks.length;j++){var o=document.createElement("option");o.value=ks[j];o.textContent=ks[j];s.appendChild(o);}};');
    p.push('addOpts("fLoc",locs);addOpts("fDept",depts);addOpts("fTitle",titles);}');

    // Render member list
    p.push('function renderML(){');
    p.push('var q=(document.getElementById("mSearch").value||"").toLowerCase();');
    p.push('var loc=document.getElementById("fLoc").value;');
    p.push('var dept=document.getElementById("fDept").value;');
    p.push('var title=document.getElementById("fTitle").value;');
    p.push('var fl=[];');
    p.push('for(var i=0;i<allM.length;i++){');
    p.push('var m=allM[i];');
    p.push('if(q&&m.name.toLowerCase().indexOf(q)===-1&&m.email.toLowerCase().indexOf(q)===-1)continue;');
    p.push('if(loc&&m.location!==loc)continue;');
    p.push('if(dept&&m.department!==dept)continue;');
    p.push('if(title&&m.jobTitle!==title)continue;');
    p.push('fl.push(m);}');
    p.push('var el=document.getElementById("mList");');
    p.push('if(!fl.length){el.innerHTML=\'<div style="padding:20px;text-align:center;color:#a8a29e">No members match</div>\';return;}');
    p.push('var h="";');
    p.push('for(var j=0;j<fl.length;j++){');
    p.push('var m=fl[j];var s=!!selMems[m.email];');
    p.push('h+=\'<div class="mrow\'+(s?" sel":"")+\'" onclick="togM(\\x27\'+m.email+\'\\x27,\\x27\'+m.name.replace(/\\x27/g,"\\\\\\x27")+\'\\x27)">\';');
    p.push('h+=\'<div class="mchk\'+(s?" on":"")+"\">"+(s?"\\u2713":"")+"</div>";');
    p.push('h+=\'<div style="flex:1;min-width:0"><div class="mname">\'+m.name+\'</div>\';');
    p.push('h+=\'<div class="mdtl">\'+m.email;');
    p.push('if(m.location)h+=" \\u00b7 "+m.location;');
    p.push('if(m.department)h+=" \\u00b7 "+m.department;');
    p.push('h+=\'</div></div></div>\';}');
    p.push('el.innerHTML=h;updCnt();}');

    p.push('function filterM(){renderML();}');

    // Toggle member
    p.push('function togM(email,name){');
    p.push('if(selMems[email])delete selMems[email];');
    p.push('else selMems[email]=name;');
    p.push('renderML();}');

    p.push('function updCnt(){');
    p.push('var c=Object.keys(selMems).length;');
    p.push('document.getElementById("sCnt").textContent=c+" selected";}');

    // Send notification
    p.push('function doSend(){');
    p.push('var title=document.getElementById("nTitle").value.trim();');
    p.push('var msg=document.getElementById("nMsg").value.trim();');
    p.push('if(!title||!msg){showT("Title and message required","err");return;}');
    p.push('var btn=document.getElementById("sendBtn");');
    p.push('btn.disabled=true;btn.textContent="Sending...";');
    p.push('var recips=[];');
    p.push('if(rMode==="groups"){recips=[selGroup];}');
    p.push('else{recips=Object.keys(selMems);');
    p.push('if(!recips.length){showT("Select at least one recipient","err");btn.disabled=false;btn.textContent="Send Notification";return;}}');
    p.push('var sent=0;var total=recips.length;var errs=[];');
    p.push('for(var i=0;i<recips.length;i++){');
    p.push('(function(recip){');
    p.push('google.script.run.withSuccessHandler(function(r){');
    p.push('sent++;if(!r||!r.success)errs.push(recip);');
    p.push('if(sent===total){btn.disabled=false;btn.textContent="Send Notification";');
    p.push('if(errs.length){showT(errs.length+" failed","err");}');
    p.push('else{showT("Sent to "+total+" recipient"+(total>1?"s":""),"ok");');
    p.push('document.getElementById("nTitle").value="";');
    p.push('document.getElementById("nMsg").value="";');
    p.push('document.getElementById("nExp").value="";');
    p.push('selMems={};updCnt();renderML();refreshN();}}');
    p.push('}).withFailureHandler(function(err){');
    p.push('sent++;errs.push(recip);');
    p.push('if(sent===total){btn.disabled=false;btn.textContent="Send Notification";showT("Error: "+err,"err");}');
    p.push('}).sendWebAppNotification({');
    p.push('recipient:recip,');
    p.push('type:document.getElementById("nType").value,');
    p.push('title:title,message:msg,');
    p.push('priority:document.getElementById("nPri").value,');
    p.push('expiresDate:document.getElementById("nExp").value,');
    p.push('senderName:UN});');
    p.push('})(recips[i]);}');
    p.push('}');

    // Refresh after send
    p.push('function refreshN(){');
    p.push('google.script.run.withSuccessHandler(function(data){');
    p.push('notifs=data||[];renderN();');
    p.push('}).getWebAppNotifications(UE,UR);}');
  }

  // Toast
  p.push('function showT(msg,type){');
  p.push('var t=document.getElementById("toast");');
  p.push('t.className="toast "+(type||"ok");');
  p.push('t.textContent=msg;t.style.display="block";');
  p.push('setTimeout(function(){t.style.display="none";},3000);}');

  // Init
  p.push('renderN();');
  if (isSteward) {
    p.push('buildFilters();');
  }
  p.push('</script></body></html>');

  return p.join('\n');
}
