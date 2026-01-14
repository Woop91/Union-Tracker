/**
 * ============================================================================
 * Integrations.gs - External Service Integration
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
 * @version 2.0.0
 * @requires Constants.gs
 */

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

/**
 * Gets or creates the root folder for grievance files
 * @return {Folder} The root grievance folder
 */
function getOrCreateRootFolder() {
  const folderName = DRIVE_CONFIG.ROOT_FOLDER_NAME;
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  // Create the root folder
  const newFolder = DriveApp.createFolder(folderName);

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
 * @param {string} grievanceId - The grievance ID
 * @return {Object} Result with folder URL or error
 */
function setupDriveFolderForGrievance(grievanceId) {
  try {
    // Get grievance data
    const grievance = getGrievanceById(grievanceId);
    if (!grievance) {
      return { success: false, error: 'Grievance not found' };
    }

    const memberName = grievance['Member Name'] ||
                       grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.MEMBER_NAME]];

    // Create folder name from template
    const folderName = DRIVE_CONFIG.SUBFOLDER_TEMPLATE
      .replace('{grievanceId}', grievanceId)
      .replace('{memberName}', sanitizeFolderName(memberName));

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
    return { success: false, error: error.message };
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
 * Updates the Drive folder link in grievance record
 * @param {string} grievanceId - The grievance ID
 * @param {string} folderUrl - The folder URL
 */
function updateGrievanceFolderLink(grievanceId, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
      sheet.getRange(i + 1, GRIEVANCE_COLUMNS.DRIVE_FOLDER + 1).setValue(folderUrl);
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

  const folderUrl = sheet.getRange(row, GRIEVANCE_COLUMNS.DRIVE_FOLDER + 1).getValue();

  if (folderUrl) {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open('${folderUrl}', '_blank'); google.script.host.close();</script>`
    ).setWidth(100).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening folder...');
  } else {
    if (showConfirmation('No folder exists. Create one now?', 'Create Folder')) {
      const grievanceId = sheet.getRange(row, GRIEVANCE_COLUMNS.GRIEVANCE_ID + 1).getValue();
      const result = setupDriveFolderForGrievance(grievanceId);
      if (result.success) {
        const html = HtmlService.createHtmlOutput(
          `<script>window.open('${result.folderUrl}', '_blank'); google.script.host.close();</script>`
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
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 50);
}

// ============================================================================
// GOOGLE CALENDAR INTEGRATION
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
    return { success: false, error: error.message };
  }
}

/**
 * Syncs a single grievance's deadlines to calendar
 * @param {Object} grievance - Grievance data object
 * @param {Calendar} calendar - Target calendar
 * @return {Object} Sync result
 */
function syncGrievanceDeadlinesToCalendar(grievance, calendar) {
  const grievanceId = grievance['Grievance ID'] ||
                      grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.GRIEVANCE_ID]];
  const memberName = grievance['Member Name'] ||
                     grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.MEMBER_NAME]];
  const currentStep = grievance['Current Step'] ||
                      grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.CURRENT_STEP]];

  // Get the deadline for current step
  let deadline;
  switch (currentStep) {
    case 1:
      deadline = grievance['Step 1 Due'] ||
                 grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.STEP_1_DUE]];
      break;
    case 2:
      deadline = grievance['Step 2 Due'] ||
                 grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.STEP_2_DUE]];
      break;
    case 3:
      deadline = grievance['Step 3 Due'] ||
                 grievance[Object.keys(grievance)[GRIEVANCE_COLUMNS.STEP_3_DUE]];
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

/**
 * Clears all calendar events created by the dashboard
 * @return {Object} Result with count of deleted events
 */
function clearAllCalendarEvents() {
  try {
    const calendar = getOrCreateDeadlinesCalendar();

    // Get all events from now until 1 year from now
    const startDate = new Date();
    const endDate = new Date();
    endDate.setFullYear(endDate.getFullYear() + 1);

    const events = calendar.getEvents(startDate, endDate, {
      search: '[GRV]'
    });

    let deleted = 0;
    for (const event of events) {
      event.deleteEvent();
      deleted++;

      // Rate limiting
      if (deleted % 50 === 0) {
        Utilities.sleep(200);
      }
    }

    return {
      success: true,
      deleted: deleted,
      message: `Deleted ${deleted} calendar events`
    };

  } catch (error) {
    console.error('Error clearing calendar:', error);
    return { success: false, error: error.message };
  }
}

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
                 <td style="padding: 10px; border: 1px solid #ddd;">${d.grievanceId}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${d.memberName}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${d.step}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${d.date}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${d.daysLeft}</td>
               </tr>`;
    });

    body += `</table>`;
    body += `<p style="margin-top: 20px; color: #666; font-size: 12px;">
               This is an automated reminder from the Union Dashboard.
             </p>`;

    // Send email
    MailApp.sendEmail({
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
    return { success: false, error: error.message };
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
      return { success: false, error: 'Member not found' };
    }

    const email = member['Email'] || member[Object.keys(member)[MEMBER_COLUMNS.EMAIL]];
    if (!email || !VALIDATION_RULES.EMAIL_PATTERN.test(email)) {
      return { success: false, error: 'Invalid email address' };
    }

    MailApp.sendEmail({
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
    return { success: false, error: error.message };
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

        <div class="sync-option">
          <h4>Clear All Events</h4>
          <p>Removes all grievance-related events from the calendar</p>
          <button class="btn btn-danger" onclick="clearAll()" style="margin-top: 10px;">
            Clear Calendar
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

        function clearAll() {
          if (!confirm('Are you sure you want to clear all grievance events from the calendar?')) return;
          showStatus('Clearing...', false);
          google.script.run
            .withSuccessHandler(function(r) {
              showStatus(r.success ? r.message : 'Error: ' + r.error, !r.success);
            })
            .clearAllCalendarEvents();
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
        <td style="padding:8px; border-bottom:1px solid #ddd;">${d.grievanceId}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${d.memberName}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${d.step}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${d.date} (${d.daysLeft} days)</td>
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
 * Shows confirmation dialog for clearing calendar
 */
function showClearCalendarConfirm() {
  const result = showConfirmation(
    'This will delete ALL grievance-related events from your calendar. This cannot be undone. Continue?',
    'Clear Calendar Events'
  );

  if (result) {
    const clearResult = clearAllCalendarEvents();
    if (clearResult.success) {
      showToast(clearResult.message, 'Calendar Cleared');
    } else {
      showAlert('Error: ' + clearResult.error, 'Error');
    }
  }
}
