// ============================================================================
// ENHANCED VISUAL FORMATTING - Gradient Heatmaps
// ============================================================================

/**
 * Applies gradient heatmap conditional formatting to numeric columns
 * Creates smooth color transitions instead of solid fills
 * Applies to: Days Open, Days to Deadline columns in Grievance Log
 */
function applyGradientHeatmaps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No data to format', 'ℹ️ Info', 3);
    return;
  }

  // Get existing rules to preserve them
  var existingRules = sheet.getConditionalFormatRules();

  // Days Open column (column S = 19)
  var daysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var daysOpenRange = sheet.getRange(daysOpenCol + '2:' + daysOpenCol + lastRow);

  // Days to Deadline column (column U = 21)
  var daysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var deadlineRange = sheet.getRange(daysToDeadlineCol + '2:' + daysToDeadlineCol + lastRow);

  // Create gradient rule for Days Open (Green = low/good, Red = high/bad)
  // Lower days open is better
  var daysOpenGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(COLORS.GRADIENT_LOW, SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue(COLORS.GRADIENT_MID_LOW, SpreadsheetApp.InterpolationType.NUMBER, '30')
    .setGradientMaxpointWithValue(COLORS.GRADIENT_HIGH, SpreadsheetApp.InterpolationType.NUMBER, '90')
    .setRanges([daysOpenRange])
    .build();

  // Create gradient rule for Days to Deadline (Green = high/good, Red = low/urgent)
  // More days remaining is better
  var deadlineGradient = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(COLORS.GRADIENT_HIGH, SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue(COLORS.GRADIENT_MID, SpreadsheetApp.InterpolationType.NUMBER, '7')
    .setGradientMaxpointWithValue(COLORS.GRADIENT_LOW, SpreadsheetApp.InterpolationType.NUMBER, '14')
    .setRanges([deadlineRange])
    .build();

  // Add gradient rules to existing rules
  existingRules.push(daysOpenGradient, deadlineGradient);
  sheet.setConditionalFormatRules(existingRules);

  ss.toast('Gradient heatmaps applied to Days Open & Days to Deadline columns!', '🎨 Heatmaps Applied', 5);
}

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

  // Prevent concurrent sync operations which can cause data corruption
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ss.toast('Another sync is already running. Please wait.', '⏳ Sync Busy', 5);
    return;
  }

  try {
    ss.toast('Syncing all dashboard data...', '🔄 Syncing', 2);

    // Sync hidden calculation sheets first
    if (typeof syncGrievanceCalcSheet === 'function') syncGrievanceCalcSheet();
    if (typeof syncDashboardCalcValues === 'function') syncDashboardCalcValues();
    if (typeof syncStewardPerformanceValues === 'function') syncStewardPerformanceValues();

    // Sync visible dashboard values
    if (typeof syncDashboardValues === 'function') syncDashboardValues();
    if (typeof syncSatisfactionValues === 'function') syncSatisfactionValues();

    ss.toast('All dashboard data synced successfully!', '✅ Complete', 5);
  } catch (e) {
    ss.toast('Error syncing: ' + e.message, '❌ Error', 5);
    Logger.log('syncAllDashboardData error: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// TESTING FUNCTIONS
// ============================================================================

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
  var folderName = DRIVE_CONFIG.ROOT_FOLDER_NAME;
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
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

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
  var data = dataRange.getValues();

  // Sort with Message Alert first, then by status priority
  data.sort(function(a, b) {
    // FIRST: Message Alert rows go to the very top
    var alertA = a[GRIEVANCE_COLS.MESSAGE_ALERT - 1] === true;
    var alertB = b[GRIEVANCE_COLS.MESSAGE_ALERT - 1] === true;

    if (alertA && !alertB) return -1; // A has alert, B doesn't - A goes first
    if (!alertA && alertB) return 1;  // B has alert, A doesn't - B goes first

    // SECOND: Sort by status priority
    var statusA = a[GRIEVANCE_COLS.STATUS - 1] || '';
    var statusB = b[GRIEVANCE_COLS.STATUS - 1] || '';

    var priorityA = GRIEVANCE_STATUS_PRIORITY[statusA] || 99;
    var priorityB = GRIEVANCE_STATUS_PRIORITY[statusB] || 99;

    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }

    // THIRD: Sort by Days to Deadline - most urgent first
    var daysA = a[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var daysB = b[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    if (daysA === '' || daysA === null) daysA = 9999;
    if (daysB === '' || daysB === null) daysB = 9999;

    return daysA - daysB;
  });

  // Write sorted data back
  dataRange.setValues(data);

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

  for (var i = 0; i < alertValues.length; i++) {
    var row = i + 2;
    var rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());

    if (alertValues[i][0] === true) {
      // Highlight the entire row
      rowRange.setBackground(highlightColor);
    } else {
      // Remove highlighting (reset to white)
      rowRange.setBackground(normalColor);
    }
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
    var _issueCategory = data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'General';
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
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
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

/**
 * Repair all checkboxes in both sheets
 */
function repairAllCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Repairing checkboxes...', '🔧 Repair', 2);

  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  ss.toast('All checkboxes repaired!', '✅ Success', 3);
}

// ============================================================================
// ENGAGEMENT TRACKING SYNC FUNCTIONS
// ============================================================================

/**
 * Syncs volunteer hours from Volunteer Hours sheet to Member Directory
 * Calculates total hours for each member and updates column S (VOLUNTEER_HOURS)
 * @returns {void}
 */
function syncVolunteerHoursToMemberDirectory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var volunteerSheet = ss.getSheetByName(SHEETS.VOLUNTEER_HOURS);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!volunteerSheet || !memberSheet) {
    Logger.log('Required sheets not found for volunteer hours sync');
    return;
  }

  // Get volunteer hours data
  var volunteerData = volunteerSheet.getDataRange().getValues();
  if (volunteerData.length < 3) {
    // No data rows (only headers row 1-2)
    Logger.log('No volunteer hours data to sync');
    return;
  }

  // Build lookup map: memberId -> totalHours
  var hoursLookup = {};

  for (var i = 2; i < volunteerData.length; i++) {  // Start at row 3 (index 2)
    var row = volunteerData[i];
    var memberId = row[1];  // Column B = Member ID (Volunteer Hours: A=Title, B=MemberID, C=Name, D=Date, E=Activity, F=Hours)
    var hours = row[5];     // Column F = Hours (see column layout above)

    if (!memberId) continue;

    // Initialize if doesn't exist
    if (!hoursLookup[memberId]) {
      hoursLookup[memberId] = 0;
    }

    // Add hours (handle both number and string)
    var hoursNum = parseFloat(hours) || 0;
    hoursLookup[memberId] += hoursNum;
  }

  // Get member data
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) return;

  // Update VOLUNTEER_HOURS column (U)
  var updates = [];
  for (var j = 1; j < memberData.length; j++) {
    memberId = memberData[j][MEMBER_COLS.MEMBER_ID - 1];
    var totalHours = hoursLookup[memberId] || 0;
    updates.push([totalHours]);
  }

  if (updates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, updates.length, 1).setValues(updates);
  }

  Logger.log('Synced volunteer hours to ' + updates.length + ' members');
}

/**
 * Syncs meeting attendance from Meeting Attendance sheet to Member Directory
 * Updates Last Virtual Mtg (column R) and Last In-Person Mtg (column S)
 * @returns {void}
 */
function syncMeetingAttendanceToMemberDirectory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName(SHEETS.MEETING_ATTENDANCE);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!attendanceSheet || !memberSheet) {
    Logger.log('Required sheets not found for meeting attendance sync');
    return;
  }

  // Get attendance data
  var attendanceData = attendanceSheet.getDataRange().getValues();
  if (attendanceData.length < 3) {
    // No data rows (only headers row 1-2)
    Logger.log('No meeting attendance data to sync');
    return;
  }

  // Build lookup map: memberId -> {lastVirtual, lastInPerson}
  var attendanceLookup = {};

  for (var i = 2; i < attendanceData.length; i++) {  // Start at row 3 (index 2)
    var row = attendanceData[i];
    var meetingDate = row[1];     // Column B = Meeting Date (Meeting Attendance: A=MeetingID, B=Date, C=Type, D=Name, E=MemberID, F=Email, G=Attended)
    var meetingType = row[2];     // Column C = Meeting Type (see column layout above)
    var memberId = row[4];        // Column E = Member ID (see column layout above)
    var attended = row[6];        // Column G = Attended (see column layout above)

    if (!memberId || !attended || !meetingDate) continue;

    // Initialize if doesn't exist
    if (!attendanceLookup[memberId]) {
      attendanceLookup[memberId] = {
        lastVirtual: null,
        lastInPerson: null
      };
    }

    // Convert to date if it's a string
    var dateValue = meetingDate instanceof Date ? meetingDate : new Date(meetingDate);

    // Update last meeting date by type
    if (meetingType === 'Virtual' || meetingType === 'virtual') {
      if (!attendanceLookup[memberId].lastVirtual || dateValue > attendanceLookup[memberId].lastVirtual) {
        attendanceLookup[memberId].lastVirtual = dateValue;
      }
    } else if (meetingType === 'In-Person' || meetingType === 'in-person') {
      if (!attendanceLookup[memberId].lastInPerson || dateValue > attendanceLookup[memberId].lastInPerson) {
        attendanceLookup[memberId].lastInPerson = dateValue;
      }
    }
  }

  // Get member data
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) return;

  // Update columns P (LAST_VIRTUAL_MTG) and Q (LAST_INPERSON_MTG)
  var updates = [];
  for (var j = 1; j < memberData.length; j++) {
    memberId = memberData[j][MEMBER_COLS.MEMBER_ID - 1];
    var memberAttendance = attendanceLookup[memberId] || {lastVirtual: '', lastInPerson: ''};
    updates.push([
      memberAttendance.lastVirtual || '',
      memberAttendance.lastInPerson || ''
    ]);
  }

  if (updates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.LAST_VIRTUAL_MTG, updates.length, 2).setValues(updates);
  }

  Logger.log('Synced meeting attendance to ' + updates.length + ' members');
}

/**
 * Unified sync function for all engagement tracking
 * Syncs volunteer hours and meeting attendance to Member Directory
 * Call this function after edits to Volunteer Hours or Meeting Attendance sheets
 * @returns {void}
 */
function syncEngagementToMemberDirectory() {
  syncVolunteerHoursToMemberDirectory();
  syncMeetingAttendanceToMemberDirectory();
  SpreadsheetApp.getActiveSpreadsheet().toast('Engagement data synced successfully!', '✅ Sync Complete', 3);
}

// ============================================================================
// DEPRECATED SHEET CLEANUP
// ============================================================================

/**
 * Removes the deprecated Dashboard sheet and provides migration info
 * The Dashboard sheet was deprecated in v4.3.2 in favor of modal dashboards
 * @returns {void}
 */
function removeDeprecatedDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var dashSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashSheet) {
    ui.alert('✅ Clean',
      'No deprecated Dashboard sheet found.\n\n' +
      'Your spreadsheet is already using the modern modal dashboard system.',
      ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert(
    '🗑️ Remove Deprecated Dashboard Sheet',
    'The "💼 Dashboard" sheet is deprecated as of v4.3.2.\n\n' +
    'Modern dashboards are now modal (popup) based:\n' +
    '• Union Hub > Dashboards > Interactive Dashboard\n' +
    '• Union Hub > Dashboards > Executive Command\n' +
    '• Union Hub > Dashboards > Member Dashboard\n' +
    '• Union Hub > Dashboards > Steward Performance\n\n' +
    'Delete the deprecated sheet?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ss.deleteSheet(dashSheet);
    ss.toast('Deprecated Dashboard sheet removed', '✅ Cleanup Complete', 3);
    Logger.log('Removed deprecated Dashboard sheet');
  } else {
    ss.toast('Deprecated sheet retained', 'ℹ️ Info', 3);
  }
}

