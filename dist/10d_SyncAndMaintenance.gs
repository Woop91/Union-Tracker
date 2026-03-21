/**
 * ============================================================================
 * 10d_SyncAndMaintenance.gs — Visual Formatting & Data Sync Utilities
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Visual formatting utilities and data synchronization functions.
 *   applyWinRateGradients() applies gradient heatmaps (red->yellow->green)
 *   to win rate columns in the dashboard sheet. Also contains formatting
 *   functions for steward performance sections and email snapshot generation
 *   for periodic reports.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Gradient heatmaps provide at-a-glance visual indicators for steward
 *   performance. Red=low win rate, green=high win rate. Applied via conditional
 *   formatting rules rather than cell backgrounds so they auto-update as data
 *   changes.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Dashboard loses visual formatting — numbers still display but without color
 *   coding. Stewards lose the at-a-glance visual indicators. Data sync functions
 *   failing means manual formatting is needed.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS)
 *   Used by:    Menu items in 03_UIComponents.gs, Admin menu
 */

// ============================================================================
// ENHANCED VISUAL FORMATTING - Gradient Heatmaps
// ============================================================================
/**
 * Applies gradient heatmap to Win Rate columns across all steward performance sections
 */
function applyWinRateGradients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboard) {
    ss.toast('Dashboard not found', '❌ Error', 3);
    return;
  }

  var existingRules = dashboard.getConditionalFormatRules();

  // Win Rate in Type Analysis (column E, rows 26-30)
  var typeWinRate = dashboard.getRange('E26:E30');
  var typeGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FEE2E2', SpreadsheetApp.InterpolationType.PERCENT, '0')
    .setGradientMidpointWithValue('#FEF3C7', SpreadsheetApp.InterpolationType.PERCENT, '50')
    .setGradientMaxpointWithValue('#D1FAE5', SpreadsheetApp.InterpolationType.PERCENT, '100')
    .setRanges([typeWinRate])
    .build();

  // Win Rate in Location Breakdown (column E, rows 35-39)
  var locWinRate = dashboard.getRange('E35:E39');
  var locGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FEE2E2', SpreadsheetApp.InterpolationType.PERCENT, '0')
    .setGradientMidpointWithValue('#FEF3C7', SpreadsheetApp.InterpolationType.PERCENT, '50')
    .setGradientMaxpointWithValue('#D1FAE5', SpreadsheetApp.InterpolationType.PERCENT, '100')
    .setRanges([locWinRate])
    .build();

  // Score in Top Performers (column C, rows 93-102) - higher is better
  var perfScore = dashboard.getRange('C93:C102');
  var perfGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FEE2E2', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#FEF3C7', SpreadsheetApp.InterpolationType.NUMBER, '50')
    .setGradientMaxpointWithValue('#D1FAE5', SpreadsheetApp.InterpolationType.NUMBER, '100')
    .setRanges([perfScore])
    .build();

  // Score in Needing Support (column C, rows 107-116) - lower scores highlighted
  var needScore = dashboard.getRange('C107:C116');
  var needGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#D1FAE5', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#FEF3C7', SpreadsheetApp.InterpolationType.NUMBER, '50')
    .setGradientMaxpointWithValue('#FEE2E2', SpreadsheetApp.InterpolationType.NUMBER, '100')
    .setRanges([needScore])
    .build();

  existingRules.push(typeGradient, locGradient, perfGradient, needGradient);
  dashboard.setConditionalFormatRules(existingRules);

  ss.toast('Win Rate & Score gradients applied to dashboard!', '🎨 Gradients Applied', 5);
}

/**
 * Syncs all dashboard data and refreshes visualizations
 * Called from Visual Control Panel
 */
function syncAllDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Syncing all dashboard data...', '🔄 Syncing', 2);

  // F44: Prevent concurrent syncs
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      ss.toast('Another sync is already running. Try again shortly.', '⏳ Busy', 5);
      return;
    }

    // Sync hidden calculation sheets first
    if (typeof syncGrievanceCalcSheet === 'function') syncGrievanceCalcSheet();
    if (typeof syncDashboardCalcValues === 'function') syncDashboardCalcValues();
    if (typeof syncStewardPerformanceValues === 'function') syncStewardPerformanceValues();

    // Sync visible dashboard values
    if (typeof syncDashboardValues === 'function') syncDashboardValues();
    if (typeof syncSatisfactionValues === 'function') syncSatisfactionValues();

    ss.toast('All dashboard data synced successfully!', '✅ Complete', 5);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error syncing: ' + e.message, '❌ Error', 5);
    Logger.log('syncAllDashboardData error: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// TESTING FUNCTIONS
// ============================================================================

/**
 * Activates the Test Results sheet, or alerts the user if none exists yet.
 * @returns {void}
 */
function viewTestResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TEST_RESULTS);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('No test results yet. Run tests first using 🧪 Testing menu.');
  }
}

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

/**
 * Get or create the root Dashboard folder in Drive
 */
function getOrCreateDashboardFolder_() {
  var folderName = getDriveRootFolderName_();
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

/**
 * Show files in the selected grievance's folder
 */
function showGrievanceFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('📁 View Files', 'Please go to the Grievance Log sheet and select a grievance row first.', ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('📁 View Files', 'Please select a grievance row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var folderId = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_ID).getValue();
  var folderUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
  var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!folderId) {
    var response = ui.alert('📁 No Folder',
      'No folder exists for ' + grievanceId + '.\n\nWould you like to create one?',
      ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      setupDriveFolderForGrievance();
    }
    return;
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var fileList = [];

    while (files.hasNext()) {
      var file = files.next();
      fileList.push('• ' + file.getName());
    }

    if (fileList.length === 0) {
      response = ui.alert('📁 ' + grievanceId + ' Files',
        'Folder is empty.\n\nWould you like to open the folder to add files?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        var html = HtmlService.createHtmlOutput(
          '<script>window.open(' + JSON.stringify(folderUrl) + ', "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    } else {
      response = ui.alert('📁 ' + grievanceId + ' Files (' + fileList.length + ')',
        fileList.join('\n') + '\n\nOpen folder in Drive?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        html = HtmlService.createHtmlOutput(
          '<script>window.open(' + JSON.stringify(folderUrl) + ', "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    }
  } catch (e) {
    ui.alert('❌ Error', 'Could not access folder: ' + e.message + '\n\nThe folder may have been deleted.', ui.ButtonSet.OK);
  }
}

// ============================================================================
// GOOGLE CALENDAR INTEGRATION
// ============================================================================

/**
 * Show upcoming deadlines from calendar with member names
 */
function showUpcomingDeadlinesFromCalendar() {
  var ui = SpreadsheetApp.getUi();

  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var today = new Date();
    var nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

    var events = calendar.getEvents(today, nextWeek, {search: 'Grievance'});

    if (events.length === 0) {
      ui.alert('📅 Upcoming Deadlines',
        'No grievance deadlines in the next 7 days!\n\n' +
        'Use "Sync Deadlines to Calendar" to add deadline events.',
        ui.ButtonSet.OK);
      return;
    }

    // Build a lookup of grievance IDs to member names
    var memberLookup = buildGrievanceMemberLookup();

    var eventList = events.map(function(e) {
      var date = Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'MM/dd');
      var title = e.getTitle();

      // Extract grievance ID from title (format: "Grievance GR-XXXX: Step X Due")
      var match = title.match(/Grievance\s+(GR-\d+)/i);
      var memberInfo = '';

      if (match && match[1] && memberLookup[match[1]]) {
        memberInfo = ' (' + memberLookup[match[1]] + ')';
      }

      return '• ' + date + ': ' + title + memberInfo;
    });

    ui.alert('📅 Upcoming Deadlines (Next 7 Days)',
      'Events with member names:\n\n' + eventList.join('\n'),
      ui.ButtonSet.OK);

  } catch (error) {
    if (error.message.indexOf('too many') !== -1 || error.message.indexOf('rate') !== -1) {
      ui.alert('⚠️ Calendar Rate Limit',
        'Google Calendar is temporarily limiting requests.\n\n' +
        'Please wait a few minutes and try again.\n\n' +
        'Tip: Avoid running calendar operations repeatedly in quick succession.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Calendar Error', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Build a lookup map of grievance IDs to member names
 * @return {Object} Map of grievanceId -> "First Last"
 */
function buildGrievanceMemberLookup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var lookup = {};

  if (!sheet || sheet.getLastRow() <= 1) return lookup;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.LAST_NAME).getValues();

  data.forEach(function(row) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
    var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';

    if (grievanceId) {
      lookup[grievanceId] = (firstName + ' ' + lastName).trim() || 'Unknown';
    }
  });

  return lookup;
}

// ============================================================================
// GRIEVANCE TOOLS - ADDITIONAL FUNCTIONS
// ============================================================================

/**
 * Fix existing "Overdue" text in Days to Deadline column
 * Converts text back to negative numbers for proper counting
 */
function fixOverdueTextToNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log not found.');
    return;
  }

  ss.toast('Fixing overdue data...', '🔧 Fix', 3);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var daysCol = GRIEVANCE_COLS.DAYS_TO_DEADLINE;
  var nextActionCol = GRIEVANCE_COLS.NEXT_ACTION_DUE;

  var daysData = sheet.getRange(2, daysCol, lastRow - 1, 1).getValues();
  var nextActionData = sheet.getRange(2, nextActionCol, lastRow - 1, 1).getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var updates = [];
  var fixCount = 0;

  for (var i = 0; i < daysData.length; i++) {
    var currentValue = daysData[i][0];
    var nextAction = nextActionData[i][0];

    if (currentValue === 'Overdue' && nextAction instanceof Date) {
      var days = Math.floor((nextAction - today) / (1000 * 60 * 60 * 24));
      updates.push([days]);
      fixCount++;
    } else {
      updates.push([currentValue]);
    }
  }

  if (fixCount > 0) {
    sheet.getRange(2, daysCol, updates.length, 1).setValues(updates);
    ss.toast('Fixed ' + fixCount + ' overdue entries!', '✅ Success', 3);
  } else {
    ss.toast('No "Overdue" text found to fix.', '✅ All Good', 3);
  }
}

// ============================================================================
// GRIEVANCE LOG SORTING
// ============================================================================

/**
 * Auto-sort the Grievance Log by status priority
 * Message Alert rows appear FIRST (highlighted),
 * then active cases (Open, Pending Info, In Arbitration, Appealed),
 * then resolved cases (Settled, Won, Denied, Withdrawn, Closed) appear last
 */
function sortGrievanceLogByStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  // Ensure sheet has enough columns for checkbox re-application after sort
  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // Need at least 2 data rows to sort

  // Get all data (excluding header row) - use actual column count to avoid truncation
  var lastCol = sheet.getLastColumn();
  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var data  = dataRange.getValues();
  var notes = dataRange.getNotes(); // capture per-cell notes so they move with their rows

  // Zip rows with their notes, sort, then unzip
  var rows = data.map(function(row, i) { return { d: row, n: notes[i] }; });

  // Sort with Message Alert first, then by status priority
  rows.sort(function(a, b) {
    // FIRST: Message Alert rows go to the very top
    var alertA = a.d[GRIEVANCE_COLS.MESSAGE_ALERT - 1] === true;
    var alertB = b.d[GRIEVANCE_COLS.MESSAGE_ALERT - 1] === true;

    if (alertA && !alertB) return -1; // A has alert, B doesn't - A goes first
    if (!alertA && alertB) return 1;  // B has alert, A doesn't - B goes first

    // SECOND: Sort by status priority
    var statusA = a.d[GRIEVANCE_COLS.STATUS - 1] || '';
    var statusB = b.d[GRIEVANCE_COLS.STATUS - 1] || '';

    var priorityA = GRIEVANCE_STATUS_PRIORITY[statusA] || 99;
    var priorityB = GRIEVANCE_STATUS_PRIORITY[statusB] || 99;

    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }

    // THIRD: Sort by Days to Deadline - most urgent first
    var daysA = a.d[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var daysB = b.d[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    // Handle non-numeric values: 'Overdue' → -1 (most urgent), blank/null → 9999 (least urgent)
    if (typeof daysA !== 'number' || isNaN(daysA)) daysA = (daysA === 'Overdue' ? -1 : 9999);
    if (typeof daysB !== 'number' || isNaN(daysB)) daysB = (daysB === 'Overdue' ? -1 : 9999);

    return daysA - daysB;
  });

  // Write sorted data and notes back
  dataRange.setValues(rows.map(function(r) { return r.d; }));
  dataRange.setNotes(rows.map(function(r) { return r.n; }));

  // Re-apply checkboxes - setValues overwrites them
  if (lastRow >= 2) {
    sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();
    sheet.getRange(2, GRIEVANCE_COLS.QUICK_ACTIONS, lastRow - 1, 1).insertCheckboxes();
  }

  // Apply highlighting to Message Alert rows
  applyMessageAlertHighlighting_(sheet, lastRow);

  Logger.log('Grievance Log sorted by status priority');
  ss.toast('Grievance Log sorted by status priority', 'Sorted', 2);
}

/**
 * Apply or remove highlighting for Message Alert rows
 * @private
 */
function applyMessageAlertHighlighting_(sheet, lastRow) {
  if (lastRow < 2) return;

  var alertCol = GRIEVANCE_COLS.MESSAGE_ALERT;
  var alertValues = sheet.getRange(2, alertCol, lastRow - 1, 1).getValues();
  var highlightColor = '#FFF2CC'; // Light yellow/orange
  var normalColor = null; // Remove background (white)
  var lastCol = sheet.getLastColumn();

  // Build backgrounds array and apply in a single batch
  var backgrounds = [];
  for (var i = 0; i < alertValues.length; i++) {
    var color = alertValues[i][0] === true ? highlightColor : normalColor;
    backgrounds.push(new Array(lastCol).fill(color));
  }
  if (backgrounds.length > 0) {
    sheet.getRange(2, 1, backgrounds.length, lastCol).setBackgrounds(backgrounds);
  }
}

/**
 * Automatically create Drive folders for grievances that don't have one
 * Called by onEditAutoSync when Grievance Log is edited
 * @private
 */
function autoCreateMissingGrievanceFolders_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Get all data at once for efficiency
  var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValues();
  var rootFolder = null;
  var created = 0;

  for (var i = 0; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var firstName = data[i][GRIEVANCE_COLS.FIRST_NAME - 1];
    var lastName = data[i][GRIEVANCE_COLS.LAST_NAME - 1];
    var dateFiled = data[i][GRIEVANCE_COLS.DATE_FILED - 1];
    var existingFolderId = data[i][GRIEVANCE_COLS.DRIVE_FOLDER_ID - 1];

    // Skip if no grievance ID or already has a folder
    if (!grievanceId || existingFolderId) continue;

    // Lazy-load root folder only when needed
    if (!rootFolder) {
      rootFolder = getOrCreateDashboardFolder_();
    }

    try {
      // Format date as YYYY-MM-DD (default to current date if not provided)
      var date = dateFiled ? new Date(dateFiled) : new Date();
      var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      // Create folder name: LastName, FirstName - YYYY-MM-DD
      var sanitizedFirst = sanitizeFolderName_(firstName || '');
      var sanitizedLast = sanitizeFolderName_(lastName || '');

      var folderName;
      if (sanitizedFirst && sanitizedLast) {
        folderName = sanitizedLast + ', ' + sanitizedFirst + ' - ' + dateStr;
      } else {
        folderName = grievanceId + ' - ' + dateStr;
      }

      // Create the folder
      var folder = rootFolder.createFolder(folderName);

      // Create subfolders for organization
      folder.createFolder('📄 Documents');
      folder.createFolder('📧 Correspondence');
      folder.createFolder('📝 Notes');

      // Update the sheet with folder info
      var row = i + 2; // Convert to 1-indexed row number
      sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_ID).setValue(folder.getId());
      sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folder.getUrl());

      created++;
      Logger.log('Auto-created folder for ' + grievanceId + ': ' + folder.getUrl());

    } catch (e) {
      Logger.log('Error auto-creating folder for ' + grievanceId + ': ' + e.message);
    }
  }

  if (created > 0) {
    ss.toast('Auto-created ' + created + ' folder(s) for new grievance(s)', '📁 Folders Created', 3);
  }
}

/**
 * Open the grievance form pre-populated with member data from a specific row
 * @param {Sheet} sheet - The Member Directory sheet
 * @param {number} row - The row number to get member data from
 * @private
 */
function openGrievanceFormForRow_(sheet, row) {
  var rowData = sheet.getRange(row, 1, 1, MEMBER_COLS.START_GRIEVANCE).getValues()[0];
  var memberId = rowData[MEMBER_COLS.MEMBER_ID - 1];

  if (!memberId) {
    SpreadsheetApp.getActiveSpreadsheet().toast('This row has no Member ID', '⚠️ Cannot Start Grievance', 3);
    return;
  }

  var memberData = {
    memberId: memberId,
    firstName: rowData[MEMBER_COLS.FIRST_NAME - 1] || '',
    lastName: rowData[MEMBER_COLS.LAST_NAME - 1] || '',
    jobTitle: rowData[MEMBER_COLS.JOB_TITLE - 1] || '',
    workLocation: rowData[MEMBER_COLS.WORK_LOCATION - 1] || '',
    unit: rowData[MEMBER_COLS.UNIT - 1] || '',
    email: rowData[MEMBER_COLS.EMAIL - 1] || '',
    manager: rowData[MEMBER_COLS.MANAGER - 1] || ''
  };

  // Get current user as steward
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stewardData = getCurrentStewardInfo_(ss);

  // Build pre-filled form URL
  var formUrl = buildGrievanceFormUrl_(memberData, stewardData);

  // Open form in new window
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open(' + JSON.stringify(formUrl) + ', "_blank");google.script.host.close();</script>'
  ).setWidth(200).setHeight(50);

  ui.showModalDialog(html, 'Opening Grievance Form...');

  ss.toast('Grievance form opened for ' + memberData.firstName + ' ' + memberData.lastName, '📋 Form Opened', 3);
}

// ============================================================================
// DATA QUALITY
// ============================================================================

/**
 * Check data quality and return list of issues
 * @return {Array} List of issue descriptions
 */
function checkDataQuality() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var issues = [];

  // Check Grievance Log
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return issues;

  var lastGRow = grievanceSheet.getLastRow();
  var lastMRow = memberSheet.getLastRow();

  if (lastGRow <= 1) return issues;

  // Get all member IDs for lookup
  var memberIds = {};
  if (lastMRow > 1) {
    var memberData = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastMRow - 1, 1).getValues();
    memberData.forEach(function(row) {
      if (row[0]) memberIds[row[0]] = true;
    });
  }

  // Check grievances for missing/invalid member IDs
  var grievanceData = grievanceSheet.getRange(2, 1, lastGRow - 1, GRIEVANCE_COLS.MEMBER_ID).getValues();
  var missingMemberIds = 0;
  var invalidMemberIds = 0;

  grievanceData.forEach(function(row) {
    var _grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];

    if (!memberId || memberId === '') {
      missingMemberIds++;
    } else if (!memberIds[memberId]) {
      invalidMemberIds++;
    }
  });

  if (missingMemberIds > 0) {
    issues.push('⚠️ ' + missingMemberIds + ' grievance(s) have no Member ID');
  }
  if (invalidMemberIds > 0) {
    issues.push('⚠️ ' + invalidMemberIds + ' grievance(s) have Member IDs not found in Member Directory');
  }

  return issues;
}

/**
 * Fix data quality issues with interactive dialog
 */
function fixDataQualityIssues() {
  var ui = SpreadsheetApp.getUi();

  var issues = checkDataQuality();

  if (issues.length === 0) {
    ui.alert('✅ No Data Issues',
      'All data passes quality checks!\n\n' +
      '• All grievances have valid Member IDs\n' +
      '• All Member IDs exist in Member Directory',
      ui.ButtonSet.OK);
    return;
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#DC2626;margin-top:0}' +
    '.issue{background:#fff5f5;padding:15px;margin:10px 0;border-radius:8px;border-left:4px solid #DC2626}' +
    '.issue-title{font-weight:bold;margin-bottom:5px}' +
    '.issue-desc{font-size:13px;color:#666}' +
    '.fix-option{background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;display:flex;align-items:center}' +
    '.fix-option input{margin-right:10px}' +
    'button{padding:12px 24px;border:none;border-radius:4px;cursor:pointer;margin:5px}' +
    '.primary{background:#1a73e8;color:white}' +
    '.secondary{background:#e0e0e0}' +
    '</style></head><body><div class="container">' +
    '<h2>⚠️ Data Quality Issues</h2>' +
    '<p>The following issues were found:</p>' +
    issues.map(function(i) { return '<div class="issue">' + escapeHtml(String(i)) + '</div>'; }).join('') +
    '<h3>How to Fix:</h3>' +
    '<div class="fix-option"><strong>Option 1:</strong> Manually update Member IDs in Grievance Log</div>' +
    '<div class="fix-option"><strong>Option 2:</strong> Add missing members to Member Directory first</div>' +
    '<p style="margin-top:20px"><button class="primary" onclick="google.script.run.showGrievancesWithMissingMemberIds();google.script.host.close()">📋 View Affected Rows</button>' +
    '<button class="secondary" onclick="google.script.host.close()">Close</button></p>' +
    '</div></body></html>'
  ).setWidth(500).setHeight(450);
  ui.showModalDialog(html, '⚠️ Data Quality Issues');
}

/**
 * Show grievances that have missing or invalid Member IDs
 */
function showGrievancesWithMissingMemberIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) {
    ui.alert('Member Directory not found');
    return;
  }

  if (!grievanceSheet) {
    ui.alert('Grievance Log not found');
    return;
  }

  var lastGRow = grievanceSheet.getLastRow();
  if (lastGRow <= 1) {
    ui.alert('No grievances found');
    return;
  }

  // Get all member IDs
  var memberIds = {};
  var lastMRow = memberSheet ? memberSheet.getLastRow() : 1;
  if (lastMRow > 1) {
    var memberData = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastMRow - 1, 1).getValues();
    memberData.forEach(function(row) {
      if (row[0]) memberIds[row[0]] = true;
    });
  }

  // Find problematic rows
  var grievanceData = grievanceSheet.getRange(2, 1, lastGRow - 1, GRIEVANCE_COLS.MEMBER_ID).getValues();
  var problemRows = [];

  grievanceData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var rowNum = index + 2;

    if (!memberId || memberId === '') {
      problemRows.push('Row ' + rowNum + ': ' + grievanceId + ' - NO MEMBER ID');
    } else if (!memberIds[memberId]) {
      problemRows.push('Row ' + rowNum + ': ' + grievanceId + ' - Invalid ID: "' + memberId + '"');
    }
  });

  if (problemRows.length === 0) {
    ui.alert('✅ All Good', 'All grievances have valid Member IDs!', ui.ButtonSet.OK);
    return;
  }

  // Show first 20 rows
  var displayRows = problemRows.slice(0, 20);
  var msg = displayRows.join('\n');
  if (problemRows.length > 20) {
    msg += '\n\n... and ' + (problemRows.length - 20) + ' more rows with issues';
  }

  ui.alert('📋 Grievances with Member ID Issues (' + problemRows.length + ' total)',
    msg + '\n\n' +
    'To fix: Open Grievance Log and update the Member ID column (B) for these rows.',
    ui.ButtonSet.OK);

  // Activate Grievance Log sheet
  ss.setActiveSheet(grievanceSheet);
}

/**
 * Repair checkboxes in Grievance Log (Message Alert column AC)
 * Call this after any bulk data operations that might overwrite checkboxes
 */
function repairGrievanceCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  // Ensure sheet has enough columns for checkbox columns
  ensureMinimumColumns(grievanceSheet, getGrievanceHeaders().length);

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return;

  // Re-apply checkboxes to Message Alert column (AC = column 29)
  grievanceSheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();

  Logger.log('Repaired checkboxes for ' + (lastRow - 1) + ' grievance rows');
}

/**
 * Repair checkboxes in Member Directory (Start Grievance column AE)
 */
function repairMemberCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return;

  // Re-apply checkboxes to Start Grievance column (AE = column 31)
  memberSheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();

  Logger.log('Repaired checkboxes for ' + (lastRow - 1) + ' member rows');
}
// ============================================================================
// ENGAGEMENT TRACKING SYNC FUNCTIONS
// ============================================================================

/** Reads headers from a sheet, returns {lowercaseName: 0-indexed col}. Falls back to row 2 if row 1 has < 3 populated cells. */
function findColumnsByHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  var row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0], pop = 0;
  for (var c = 0; c < row1.length; c++) { if (String(row1[c]).trim() !== '') pop++; }
  var hdr = (pop >= 3) ? row1 : (sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] : row1);
  var map = {};
  for (var i = 0; i < hdr.length; i++) { var n = String(hdr[i]).toLowerCase().trim(); if (n) map[n] = i; }
  return map;
}
/** Debounce helper — returns true if syncName was called within the last 30s. */
function isSyncDebounced_(syncName) {
  var cache = CacheService.getScriptCache(), key = 'SYNC_DEBOUNCE_' + syncName;
  if (cache.get(key)) return true;
  cache.put(key, 'running', 30);
  return false;
}
/** Builds {memberId: true} lookup from Member Directory data (skips header row).
 *  M-68: Explicitly starts at index 1 to skip the header row. Also validates
 *  that the value looks like a member ID (not a header label) as a safeguard. */
function buildMemberIdSet_(memberData) {
  var set = {};
  // Start at 1 to skip header row (index 0)
  for (var i = 1; i < memberData.length; i++) {
    var id = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim();
    // Skip empty values and the header label itself (safeguard against
    // data arrays that don't include a header at index 0)
    if (id !== '' && id !== 'Member ID') set[id] = true;
  }
  return set;
}
/** Resolves a header column: checks aliases in order, falls back to default index. */
function resolveCol_(headers, aliases, fallback) {
  for (var a = 0; a < aliases.length; a++) { if (headers[aliases[a]] !== undefined) return headers[aliases[a]]; }
  return fallback;
}
/** Parses and validates a date value. Returns Date or null. Rejects future dates (> tomorrow). */
function parseValidDate_(val, tomorrow) {
  var d;
  if (val instanceof Date) { d = val; } else if (val && typeof val === 'string') { d = new Date(val); } else { return null; }
  return (isNaN(d.getTime()) || d > tomorrow) ? null : d;
}

/**
 * Syncs volunteer hours from Volunteer Hours sheet to Member Directory.
 * Dynamic headers, data validation, member existence checks, toast notifications.
 */
function syncVolunteerHoursToMemberDirectory() {
  if (isSyncDebounced_('volunteerHours')) { Logger.log('VH sync debounced'); return; }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Syncing volunteer hours...', '🔄 Sync', 3);

    var volunteerSheet = ss.getSheetByName(SHEETS.VOLUNTEER_HOURS);
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!volunteerSheet) { ss.toast('Volunteer Hours sheet not found. Run CREATE_DASHBOARD to create it.', '⚠️ Sync Error', 5); return; }
    if (!memberSheet) { ss.toast('Member Directory not found. Run CREATE_DASHBOARD to create it.', '⚠️ Sync Error', 5); return; }

    // Dynamic header reading
    var vh = findColumnsByHeader_(volunteerSheet);
    var idCol = resolveCol_(vh, ['member id', 'memberid'], 1);
    var hrsCol = resolveCol_(vh, ['hours'], 5);
    var dtCol = resolveCol_(vh, ['date'], 3);

    var volunteerData = volunteerSheet.getDataRange().getValues();
    if (volunteerData.length < 3) { ss.toast('No volunteer hours data to sync.', 'ℹ️ Sync', 3); return; }

    var memberData = memberSheet.getDataRange().getValues();
    if (memberData.length < 2) { ss.toast('Member Directory is empty.', 'ℹ️ Sync', 3); return; }
    var validMembers = buildMemberIdSet_(memberData);

    var hoursLookup = {}, skipped = 0, badIds = {};
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(23, 59, 59, 999);

    for (var i = 2; i < volunteerData.length; i++) {
      var row = volunteerData[i];
      var rawId = String(row[idCol] || '').trim();
      if (rawId === '') { skipped++; continue; }
      var hours = parseFloat(row[hrsCol]);
      if (isNaN(hours) || hours <= 0) { skipped++; continue; }
      if (parseValidDate_(row[dtCol], tomorrow) === null && row[dtCol]) { skipped++; continue; }
      if (!validMembers[rawId]) { badIds[rawId] = true; skipped++; continue; }
      hoursLookup[rawId] = (hoursLookup[rawId] || 0) + hours;
    }

    var invalidIds = Object.keys(badIds);
    if (invalidIds.length > 0) Logger.log('syncVH: ' + invalidIds.length + ' bad ID(s): ' + invalidIds.slice(0, 10).join(', '));

    // M-25: Only update hours for members found in the volunteer sheet.
    // Members with no volunteer data retain their existing hours value.
    var updates = [], membersUpdated = 0;
    var existingHours = memberSheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, memberData.length - 1, 1).getValues();
    for (var j = 1; j < memberData.length; j++) {
      var mid = String(memberData[j][MEMBER_COLS.MEMBER_ID - 1] || '').trim();
      if (hoursLookup[mid] !== undefined) {
        membersUpdated++;
        updates.push([hoursLookup[mid]]);
      } else {
        // Keep existing value — don't zero out members without volunteer data
        updates.push([existingHours[j - 1][0]]);
      }
    }
    var expected = memberData.length - 1;
    if (updates.length !== expected) {
      ss.toast('Write aborted: data length mismatch. Check logs.', '⚠️ Sync Error', 5);
      Logger.log('syncVH: length mismatch updates=' + updates.length + ' expected=' + expected);
      return;
    }
    if (updates.length > 0) memberSheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, updates.length, 1).setValues(updates);

    var totalSum = 0;
    for (var k in hoursLookup) totalSum += hoursLookup[k];
    ss.toast('Synced ' + Math.round(totalSum * 10) / 10 + ' hrs for ' + membersUpdated + ' members, ' + skipped + ' skipped.', '✅ Volunteer Hours', 4);
    Logger.log('Synced VH: ' + membersUpdated + ' members, ' + skipped + ' skipped');
  } catch (e) {
    Logger.log('syncVH error: ' + e.message);
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Error syncing volunteer hours: ' + e.message, '❌ Sync Error', 5); } catch (_) { Logger.log('_: ' + (_.message || _)); }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Syncs meeting attendance from Meeting Attendance sheet to Member Directory.
 * Dynamic headers, case-insensitive type matching, data validation, toast notifications.
 */
function syncMeetingAttendanceToMemberDirectory() {
  if (isSyncDebounced_('meetingAttendance')) { Logger.log('MA sync debounced'); return; }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Syncing meeting attendance...', '🔄 Sync', 3);

    var attendanceSheet = ss.getSheetByName(SHEETS.MEETING_ATTENDANCE);
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!attendanceSheet) { ss.toast('Meeting Attendance sheet not found. Run CREATE_DASHBOARD to create it.', '⚠️ Sync Error', 5); return; }
    if (!memberSheet) { ss.toast('Member Directory not found. Run CREATE_DASHBOARD to create it.', '⚠️ Sync Error', 5); return; }

    // Dynamic header reading
    var ma = findColumnsByHeader_(attendanceSheet);
    var maDateCol = resolveCol_(ma, ['date'], 1);
    var maTypeCol = resolveCol_(ma, ['type', 'meeting type'], 2);
    var maMemberCol = resolveCol_(ma, ['member id', 'memberid'], 4);
    var maAttendedCol = resolveCol_(ma, ['attended'], 6);

    var attendanceData = attendanceSheet.getDataRange().getValues();
    if (attendanceData.length < 3) { ss.toast('No meeting attendance data to sync.', 'ℹ️ Sync', 3); return; }

    var memberData = memberSheet.getDataRange().getValues();
    if (memberData.length < 2) { ss.toast('Member Directory is empty.', 'ℹ️ Sync', 3); return; }
    var validMembers = buildMemberIdSet_(memberData);

    var lookup = {}, skipped = 0, badIds = {};
    var tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(23, 59, 59, 999);

    for (var i = 2; i < attendanceData.length; i++) {
      var row = attendanceData[i];
      var rawId = String(row[maMemberCol] || '').trim();
      if (rawId === '') { skipped++; continue; }
      if (!row[maAttendedCol]) { skipped++; continue; }

      var dateValue = parseValidDate_(row[maDateCol], tomorrow);
      if (!dateValue) { skipped++; continue; }

      if (!validMembers[rawId]) { badIds[rawId] = true; skipped++; continue; }
      if (!lookup[rawId]) lookup[rawId] = { lastVirtual: null, lastInPerson: null };

      // Case-insensitive meeting type matching
      var mType = String(row[maTypeCol] || '').toLowerCase().trim();
      if (mType === 'virtual') {
        if (!lookup[rawId].lastVirtual || dateValue > lookup[rawId].lastVirtual) lookup[rawId].lastVirtual = dateValue;
      } else if (mType === 'in-person' || mType === 'in person' || mType === 'inperson') {
        if (!lookup[rawId].lastInPerson || dateValue > lookup[rawId].lastInPerson) lookup[rawId].lastInPerson = dateValue;
      }
    }

    var invalidIds = Object.keys(badIds);
    if (invalidIds.length > 0) Logger.log('syncMA: ' + invalidIds.length + ' bad ID(s): ' + invalidIds.slice(0, 10).join(', '));

    // Write to Member Directory with array length validation
    var updates = [], membersUpdated = 0;
    for (var j = 1; j < memberData.length; j++) {
      var mid = String(memberData[j][MEMBER_COLS.MEMBER_ID - 1] || '').trim();
      var att = lookup[mid] || { lastVirtual: null, lastInPerson: null };
      if (att.lastVirtual || att.lastInPerson) membersUpdated++;
      updates.push([att.lastVirtual || '', att.lastInPerson || '']);
    }
    var expected = memberData.length - 1;
    if (updates.length !== expected) {
      ss.toast('Write aborted: data length mismatch. Check logs.', '⚠️ Sync Error', 5);
      Logger.log('syncMA: length mismatch updates=' + updates.length + ' expected=' + expected);
      return;
    }
    if (updates.length > 0) memberSheet.getRange(2, MEMBER_COLS.LAST_VIRTUAL_MTG, updates.length, 2).setValues(updates);

    ss.toast('Synced attendance for ' + membersUpdated + ' members, ' + skipped + ' skipped.', '✅ Meeting Attendance', 4);
    Logger.log('Synced MA: ' + membersUpdated + ' members, ' + skipped + ' skipped');
  } catch (e) {
    Logger.log('syncMA error: ' + e.message);
    try { SpreadsheetApp.getActiveSpreadsheet().toast('Error syncing attendance: ' + e.message, '❌ Sync Error', 5); } catch (_) { Logger.log('_: ' + (_.message || _)); }
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Unified sync: volunteer hours + meeting attendance to Member Directory. */
function syncEngagementToMemberDirectory() {
  syncVolunteerHoursToMemberDirectory();
  syncMeetingAttendanceToMemberDirectory();
  SpreadsheetApp.getActiveSpreadsheet().toast('Engagement data synced successfully!', '✅ Sync Complete', 3);
}

// ============================================================================
// DEPRECATED SHEET CLEANUP
// ============================================================================

// removeDeprecatedDashboard() removed v4.33.0 — merged into removeDeprecatedTabs() in 06_Maintenance.gs

