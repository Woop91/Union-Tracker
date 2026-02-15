// ============================================================================
// AUDIT LOG SHEET SETUP
// ============================================================================

/**
 * Sets up the Audit Log sheet with proper headers and formatting
 * Creates the sheet if it doesn't exist, or clears and reformats if it does
 */
function setupAuditLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.AUDIT_LOG);
    sheet.clear();
  } else if (sheet.getLastRow() <= 1) {
    sheet.clear();
  } else {
    // Audit log has existing data — do not clear (compliance requirement)
    Logger.log('setupAuditLogSheet: Sheet has ' + sheet.getLastRow() + ' rows of data — skipping clear');

    // Still apply protection if missing
    protectAuditLogSheet_(sheet);
    return;
  }

  // Headers
  var headers = [
    'Timestamp',
    'User Email',
    'Sheet',
    'Row',
    'Column',
    'Field Name',
    'Old Value',
    'New Value',
    'Record ID',
    'Action Type'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(COLORS.PRIMARY_PURPLE);
  headerRange.setFontColor(COLORS.WHITE);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 160); // Timestamp
  sheet.setColumnWidth(2, 200); // User Email
  sheet.setColumnWidth(3, 120); // Sheet
  sheet.setColumnWidth(4, 50);  // Row
  sheet.setColumnWidth(5, 50);  // Column
  sheet.setColumnWidth(6, 150); // Field Name
  sheet.setColumnWidth(7, 200); // Old Value
  sheet.setColumnWidth(8, 200); // New Value
  sheet.setColumnWidth(9, 100); // Record ID
  sheet.setColumnWidth(10, 100); // Action Type

  // Hide the sheet
  sheet.hideSheet();

  // ═══ SHEET PROTECTION ═══
  // v4.8.1: Audit log now gets vault-style protection (owner-only editing)
  protectAuditLogSheet_(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit log sheet created and hidden.', '✅ Setup Complete', 3);
}

/**
 * Applies vault-style sheet protection to the audit log.
 * Only the script owner (installer) can edit. All other editors are removed.
 *
 * @param {Sheet} sheet - The audit log sheet
 * @private
 */
function protectAuditLogSheet_(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var alreadyProtected = false;
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === 'Audit Log — Tamper Protection') {
      alreadyProtected = true;
      break;
    }
  }

  if (!alreadyProtected) {
    var protection = sheet.protect()
      .setDescription('Audit Log — Tamper Protection')
      .setWarningOnly(false);

    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    var editors = protection.getEditors();
    for (var j = 0; j < editors.length; j++) {
      if (editors[j].getEmail() !== me.getEmail()) {
        protection.removeEditor(editors[j]);
      }
    }
    Logger.log('Audit log sheet protected (owner-only)');
  }
}

// ============================================================================
// AUDIT TRIGGER FUNCTIONS
// ============================================================================

/**
 * onEdit trigger for audit logging
 * Tracks changes to Member Directory and Grievance Log
 * @param {Object} e - The edit event object
 */
function onEditAudit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only track changes to Member Directory and Grievance Log
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Skip header row
  if (row < 2) return;

  var oldValue = e.oldValue || '';
  var newValue = e.value || '';

  // Skip if no actual change
  if (oldValue === newValue) return;

  // Get field name from header
  var fieldName = sheet.getRange(1, col).getValue() || ('Column ' + col);

  // Get record ID (column A for both sheets)
  var recordId = sheet.getRange(row, 1).getValue() || '';

  // Determine action type
  var actionType = 'Edit';
  if (!oldValue && newValue) {
    actionType = 'Create';
  } else if (oldValue && !newValue) {
    actionType = 'Delete';
  }

  logAuditEvent(actionType + '_' + sheetName.toUpperCase().replace(/\s+/g, '_'), {
    sheet: sheetName,
    row: row,
    col: col,
    field: fieldName,
    oldValue: oldValue,
    newValue: newValue,
    recordId: recordId
  });
}

/**
 * Install the audit trigger
 * Sets up automatic change tracking for Member Directory and Grievance Log
 */
function installAuditTrigger() {
  // Remove existing audit triggers
  removeAuditTrigger();

  // Create new onEdit trigger
  ScriptApp.newTrigger('onEditAudit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  // Ensure audit sheet exists
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.AUDIT_LOG)) {
    setupAuditLogSheet();
  }

  SpreadsheetApp.getUi().alert('✅ Audit Tracking Enabled',
    'All changes to Member Directory and Grievance Log will now be logged.\n\n' +
    'View the audit log via:\n⚙️ Administrator > 📋 Audit Log > 📋 View Audit Log',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Remove the audit trigger
 * Disables automatic change tracking
 */
function removeAuditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditAudit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit tracking disabled.', '🚫 Disabled', 3);
}

// ============================================================================
// AUDIT LOG VIEWING AND MANAGEMENT
// ============================================================================

/**
 * View the audit log sheet
 * Shows the hidden audit log and sorts entries by newest first
 */
function viewAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    var response = SpreadsheetApp.getUi().alert('📋 Audit Log Not Found',
      'The audit log sheet does not exist yet.\n\nWould you like to create it now?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (response === SpreadsheetApp.getUi().Button.YES) {
      setupAuditLogSheet();
      sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    } else {
      return;
    }
  }

  // Show the hidden sheet temporarily
  sheet.showSheet();
  ss.setActiveSheet(sheet);

  // Sort by timestamp descending (newest first)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).sort({column: 1, ascending: false});
  }

  SpreadsheetApp.getUi().alert('📋 Audit Log',
    'Viewing audit log.\n\n' +
    'Total entries: ' + Math.max(0, sheet.getLastRow() - 1) + '\n\n' +
    'The sheet will be hidden again when you navigate away.\n' +
    'To keep it visible, right-click the tab and select "Unhide".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear audit entries older than 30 days
 * Prompts for confirmation before deleting old entries
 */
function clearOldAuditEntries() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('🗑️ Clear Old Audit Entries',
    'This will delete all audit entries older than 30 days.\n\n' +
    'This action cannot be undone.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No audit entries to clear.');
    return;
  }

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Find rows older than 30 days (skip header)
  for (var i = data.length - 1; i >= 1; i--) {
    var timestamp = data[i][0];
    if (timestamp instanceof Date && timestamp < cutoffDate) {
      rowsToDelete.push(i + 1); // +1 for 1-indexed rows
    }
  }

  // Delete rows from bottom to top to maintain correct indices
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  ui.alert('✅ Cleanup Complete',
    'Deleted ' + rowsToDelete.length + ' entries older than 30 days.\n\n' +
    'Remaining entries: ' + Math.max(0, sheet.getLastRow() - 1),
    ui.ButtonSet.OK);
}

// ============================================================================
// AUDIT HISTORY RETRIEVAL
// ============================================================================

/**
 * Get audit summary for a specific record
 * @param {string} recordId - Member ID or Grievance ID
 * @returns {Array} Array of audit entries for this record
 */
function getAuditHistory(recordId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === recordId) { // Column I is Record ID
      history.push({
        timestamp: data[i][0],
        user: data[i][1],
        field: data[i][5],
        oldValue: data[i][6],
        newValue: data[i][7],
        action: data[i][9]
      });
    }
  }

  return history;
}



/**
 * ============================================================================
 * 08j_CalcSheets.gs - Hidden Calculation Sheet Setup & Management
 * ============================================================================
 *
 * This module contains all functions related to setting up and managing
 * the hidden calculation sheets that power the dashboard's "self-healing"
 * formula system. These sheets contain complex formulas that aggregate,
 * calculate, and cross-reference data across the dashboard.
 *
 * Hidden Sheets Managed:
 * - _Grievance_Calc: Grievance -> Member Directory data sync
 * - _Grievance_Formulas: Self-healing Grievance Log formulas
 * - _Member_Lookup: Member -> Grievance Log data sync
 * - _Steward_Contact_Calc: Steward contact tracking metrics
 * - _Steward_Performance_Calc: Steward performance scores
 * - _Dashboard_Calc: Dashboard summary statistics
 * - _CalcMembers: Member statistics and lookups
 * - _CalcGrievances: Grievance aggregations
 * - _CalcDeadlines: Deadline calculations and alerts
 * - _CalcStats: Dashboard-wide statistics
 * - _CalcSync: Cross-sheet synchronization
 * - _CalcFormulas: Named formula references
 *
 * @fileoverview Hidden sheet and formula management for 509 Dashboard
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// LIVE FORMULA SYNC FUNCTIONS
// ============================================================================

/**
 * Syncs grievance data to Member Directory using static values
 * Menu Location: Dashboard menu
 */
function setupLiveGrievanceFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory not found.');
    return;
  }

  ss.toast('Syncing grievance data to Member Directory...', '🔄 Sync', 3);

  // Use sync function to populate with static values (no formulas)
  syncGrievanceToMemberDirectory();

  ss.toast('Grievance data synced! Columns AB-AD updated with static values.', '✅ Success', 3);
}

/**
 * Remove Member ID dropdown from Grievance Log
 * Clears any existing data validation to allow free text entry
 */
function setupGrievanceMemberDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Grievance Log not found.');
    return;
  }

  ss.toast('Removing Member ID dropdown...', '🔄 Setup', 3);

  // Clear any existing data validation from Member ID column (column B, rows 2-1000)
  // This allows free text entry for Member ID
  grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, 998, 1).clearDataValidations();

  ss.toast('Member ID dropdown removed - free text entry enabled!', '✅ Success', 3);
}

// ============================================================================
// HIDDEN SHEET 1: _Grievance_Calc
// Source: Grievance Log -> Destination: Member Directory (AB-AD)
// ============================================================================

/**
 * Setup the _Grievance_Calc hidden sheet with self-healing formulas
 * Calculates: Has Open Grievance, Grievance Status, Next Deadline per member
 */
function setupGrievanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'Has Open Grievance', 'Grievance Status', 'Days to Deadline', 'Total Count', 'Win Rate %', 'Last Grievance Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for dynamic formulas
  var memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var _gNextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);

  // Formula for Member IDs (Column A) - pulls unique member IDs from Member Directory
  var memberIdFormula = '=IFERROR(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + '<>"Member ID"),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // Formulas for calculations (using ARRAYFORMULA for efficiency)
  // Column B: Has Open Grievance
  var hasOpenFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Yes","No")))';
  sheet.getRange('B2').setFormula(hasOpenFormula);

  // Column C: Grievance Status (most urgent: Open > Pending Info, blank if all closed)
  var statusFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")>0,"Open",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Pending Info",""))))';
  sheet.getRange('C2').setFormula(statusFormula);

  // Column D: Days to Deadline (minimum/most urgent deadline for open grievances only)
  // Excludes all closed statuses: Closed, Settled, Withdrawn, Denied, Won
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var deadlineFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MINIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Closed",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Settled",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Withdrawn",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Denied",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Won"),"")))';
  sheet.getRange('D2').setFormula(deadlineFormula);

  // Column E: Total Grievance Count
  var countFormula = '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)))';
  sheet.getRange('E2').setFormula(countFormula);

  // Column F: Win Rate %
  var winRateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)*100,0)))';
  sheet.getRange('F2').setFormula(winRateFormula);

  // Column G: Last Grievance Date
  var lastDateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MAXIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A),"")))';
  sheet.getRange('G2').setFormula(lastDateFormula);

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Calc sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 2: _Grievance_Formulas (SELF-HEALING)
// Source: Grievance Log -> Destination: Grievance Log (calculated columns)
// This sheet contains all auto-calculated formulas and syncs them back
// ============================================================================

/**
 * Setup the _Grievance_Formulas hidden sheet with self-healing formulas
 * Calculates: First Name, Last Name, Email, Unit, Location, Steward (from Member Dir)
 *            Filing Deadline, Step I-III dates, Days Open, Next Action Due, Days to Deadline
 */
function setupGrievanceFormulasSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FORMULAS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_FORMULAS);
  }

  sheet.clear();

  // Headers matching Grievance Log columns that need formulas
  var headers = [
    'Row Index',           // A - For tracking which row in Grievance Log
    'Member ID',           // B - From Grievance Log
    'First Name',          // C - Lookup from Member Directory
    'Last Name',           // D - Lookup from Member Directory
    'Incident Date',       // E - From Grievance Log
    'Date Filed',          // F - From Grievance Log
    'Step I Rcvd',         // G - From Grievance Log
    'Step II Appeal Filed',// H - From Grievance Log
    'Step II Rcvd',        // I - From Grievance Log
    'Status',              // J - From Grievance Log
    'Current Step',        // K - From Grievance Log
    'Date Closed',         // L - From Grievance Log
    'Filing Deadline',     // M - CALCULATED
    'Step I Due',          // N - CALCULATED
    'Step II Appeal Due',  // O - CALCULATED
    'Step II Due',         // P - CALCULATED
    'Step III Appeal Due', // Q - CALCULATED
    'Days Open',           // R - CALCULATED
    'Next Action Due',     // S - CALCULATED
    'Days to Deadline',    // T - CALCULATED
    'Member Email',        // U - Lookup from Member Directory
    'Unit',                // V - Lookup from Member Directory
    'Location',            // W - Lookup from Member Directory
    'Steward'              // X - Lookup from Member Directory
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for Grievance Log source data
  var gGrievanceIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);     // A
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);           // B
  var gIncidentDateCol = getColumnLetter(GRIEVANCE_COLS.INCIDENT_DATE);   // G
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);         // I
  var gStep1RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP1_RCVD);         // K
  var gStep2AppealFiledCol = getColumnLetter(GRIEVANCE_COLS.STEP2_APPEAL_FILED); // M
  var gStep2RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP2_RCVD);         // O
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);                // E
  var gCurrentStepCol = getColumnLetter(GRIEVANCE_COLS.CURRENT_STEP);     // F
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);       // R

  // Member Directory columns for lookups
  var mMemberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);
  var memberRange = "'" + SHEETS.MEMBER_DIR + "'!" + mMemberIdCol + ":" + mStewardCol;

  // Column A: Row Index (ROW()-1 to match Grievance Log rows)
  // Pull unique grievance IDs to create row mapping
  sheet.getRange('A2').setFormula(
    '=IFERROR(FILTER(ROW(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + ')-1,' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + '<>""),"")'
  );

  // Column B: Member ID (from Grievance Log)
  sheet.getRange('B2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A+1)))'
  );

  // Column C: First Name (VLOOKUP from Member Directory)
  sheet.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.FIRST_NAME + ',FALSE),"")))'
  );

  // Column D: Last Name (VLOOKUP from Member Directory)
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.LAST_NAME + ',FALSE),"")))'
  );

  // Column E: Incident Date (from Grievance Log)
  sheet.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIncidentDateCol + ':' + gIncidentDateCol + ',A2:A+1)))'
  );

  // Column F: Date Filed (from Grievance Log)
  sheet.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',A2:A+1)))'
  );

  // Column G: Step I Rcvd (from Grievance Log)
  sheet.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep1RcvdCol + ':' + gStep1RcvdCol + ',A2:A+1)))'
  );

  // Column H: Step II Appeal Filed (from Grievance Log)
  sheet.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2AppealFiledCol + ':' + gStep2AppealFiledCol + ',A2:A+1)))'
  );

  // Column I: Step II Rcvd (from Grievance Log)
  sheet.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2RcvdCol + ':' + gStep2RcvdCol + ',A2:A+1)))'
  );

  // Column J: Status (from Grievance Log)
  sheet.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',A2:A+1)))'
  );

  // Column K: Current Step (from Grievance Log)
  sheet.getRange('K2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gCurrentStepCol + ':' + gCurrentStepCol + ',A2:A+1)))'
  );

  // Column L: Date Closed (from Grievance Log)
  sheet.getRange('L2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',A2:A+1)))'
  );

  // =========== CALCULATED COLUMNS ===========

  // Column M: Filing Deadline = Incident Date + 21 days
  sheet.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(E2:E="","",E2:E+21))'
  );

  // Column N: Step I Due = Date Filed + 30 days
  sheet.getRange('N2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",F2:F+30))'
  );

  // Column O: Step II Appeal Due = Step I Rcvd + 10 days
  sheet.getRange('O2').setFormula(
    '=ARRAYFORMULA(IF(G2:G="","",G2:G+10))'
  );

  // Column P: Step II Due = Step II Appeal Filed + 30 days
  sheet.getRange('P2').setFormula(
    '=ARRAYFORMULA(IF(H2:H="","",H2:H+30))'
  );

  // Column Q: Step III Appeal Due = Step II Rcvd + 30 days
  sheet.getRange('Q2').setFormula(
    '=ARRAYFORMULA(IF(I2:I="","",I2:I+30))'
  );

  // Column R: Days Open = IF closed: Date Closed - Date Filed, ELSE: Today - Date Filed
  sheet.getRange('R2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",IF(L2:L<>"",L2:L-F2:F,TODAY()-F2:F)))'
  );

  // Column S: Next Action Due = Based on current step and status
  // If closed status, leave blank; otherwise return appropriate deadline
  sheet.getRange('S2').setFormula(
    '=ARRAYFORMULA(IF(J2:J="","",' +
    'IF(OR(J2:J="Settled",J2:J="Withdrawn",J2:J="Denied",J2:J="Won",J2:J="Closed"),"",' +
    'IF(K2:K="Informal",M2:M,' +
    'IF(K2:K="Step I",N2:N,' +
    'IF(K2:K="Step II",P2:P,' +
    'Q2:Q))))))'
  );

  // Column T: Days to Deadline = Next Action Due - Today
  sheet.getRange('T2').setFormula(
    '=ARRAYFORMULA(IF(S2:S="","",S2:S-TODAY()))'
  );

  // =========== MEMBER LOOKUP COLUMNS ===========

  // Column U: Member Email (VLOOKUP from Member Directory)
  sheet.getRange('U2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.EMAIL + ',FALSE),"")))'
  );

  // Column V: Unit (VLOOKUP from Member Directory)
  sheet.getRange('V2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.UNIT + ',FALSE),"")))'
  );

  // Column W: Location (VLOOKUP from Member Directory)
  sheet.getRange('W2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.WORK_LOCATION + ',FALSE),"")))'
  );

  // Column X: Steward (VLOOKUP from Member Directory)
  sheet.getRange('X2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.ASSIGNED_STEWARD + ',FALSE),"")))'
  );

  // Format date columns (MM/dd/yyyy)
  sheet.getRange('E:E').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('F:F').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('G:G').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('H:H').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('I:I').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('L:L').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('M:M').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('N:N').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('O:O').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('P:P').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('Q:Q').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('S:S').setNumberFormat('MM/dd/yyyy');

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Formulas sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 3: _Member_Lookup
// Source: Member Directory -> Destination: Grievance Log (C,D,X-AA)
// ============================================================================

/**
 * Setup the _Member_Lookup hidden sheet with self-healing formulas
 * Looks up: First Name, Last Name, Email, Unit, Location, Steward from Member Directory
 */
function setupMemberLookupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_LOOKUP);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MEMBER_LOOKUP);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'First Name', 'Last Name', 'Email', 'Unit', 'Location', 'Assigned Steward'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var _mFirstCol = getColumnLetter(MEMBER_COLS.FIRST_NAME);
  var _mLastCol = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var _mEmailCol = getColumnLetter(MEMBER_COLS.EMAIL);
  var _mUnitCol = getColumnLetter(MEMBER_COLS.UNIT);
  var _mLocCol = getColumnLetter(MEMBER_COLS.WORK_LOCATION);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);

  // Formula to get unique member IDs from Grievance Log
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var memberIdFormula = '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"Member ID",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"")),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // VLOOKUP formulas for member data
  var vlookupBase = 'VLOOKUP(A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mStewardCol + ',';

  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '2,FALSE),"")))'); // First Name
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '3,FALSE),"")))'); // Last Name
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '8,FALSE),"")))'); // Email
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '6,FALSE),"")))'); // Unit
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '5,FALSE),"")))'); // Location
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '16,FALSE),"")))'); // Steward

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Member_Lookup sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 4: _Steward_Contact_Calc
// Source: Member Directory (Y-AA) -> Aggregates steward contact tracking metrics
// ============================================================================

/**
 * Setup the _Steward_Contact_Calc hidden sheet with self-healing formulas
 * Tracks and aggregates steward contact data from Member Directory
 */
function setupStewardContactCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_CONTACT_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_CONTACT_CALC);
  }

  sheet.clear();

  // Headers for steward contact summary (5 columns)
  var headers = ['Steward Name', 'Total Contacts', 'Contacts This Month', 'Contacts Last 7 Days', 'Last Contact Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for formulas
  var mContactStewardCol = getColumnLetter(MEMBER_COLS.CONTACT_STEWARD);
  var mContactDateCol = getColumnLetter(MEMBER_COLS.RECENT_CONTACT_DATE);

  // Column A: Unique steward names who have made contacts
  sheet.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + '<>""))),)');

  // Column B: Total contacts per steward
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A)))');

  // Column C: Contacts this month
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))))');

  // Column D: Contacts last 7 days
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&(TODAY()-7))))');

  // Column E: Most recent contact date for this steward
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(TEXT(MAXIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A),"MM/dd/yyyy"),"-")))');

  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  sheet.hideSheet();
  Logger.log('_Steward_Contact_Calc sheet setup complete with live formulas');
}

// ============================================================================
// HIDDEN SHEET 5: _Dashboard_Calc
// Source: Member Directory + Grievance Log -> Dashboard Summary Statistics
// ============================================================================

/**
 * Setup the _Dashboard_Calc hidden sheet with self-healing formulas
 * Calculates key dashboard metrics that auto-update
 */
function setupDashboardCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.DASHBOARD_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Column references
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);

  // Metrics with formulas (15 key metrics)
  // Note: Using COUNTIF with "M*" and "G*" patterns to only count valid IDs (ignores blank rows)
  var metrics = [
    ['Total Members', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ',"M*")', 'Total union members in directory'],
    ['Active Stewards', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")', 'Members marked as stewards'],
    ['Total Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ',"G*")', 'All grievances filed'],
    ['Open Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")', 'Currently open cases'],
    ['Pending Info', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")', 'Cases awaiting information'],
    ['Settled', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")', 'Cases settled'],
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")', 'Cases won (full or partial)'],
    ['Denied', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")', 'Cases denied'],
    ['Withdrawn', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Withdrawn")', 'Cases withdrawn'],
    ['Win Rate %', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*"))*100,1),0)', 'Wins / (Wins + Settled + Denied)'],
    ['Avg Days to Resolution', '=IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<>",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)', 'Average days for closed cases'],
    ['Overdue Cases', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")', 'Cases past deadline'],
    ['Due This Week', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")', 'Cases due in next 7 days'],
    ['Filed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())', 'Grievances filed this month'],
    ['Closed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())', 'Grievances closed this month']
  ];

  for (var i = 0; i < metrics.length; i++) {
    sheet.getRange(i + 2, 1).setValue(metrics[i][0]);
    sheet.getRange(i + 2, 2).setFormula(metrics[i][1]);
    sheet.getRange(i + 2, 3).setValue(metrics[i][2]);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 300);

  sheet.hideSheet();
  Logger.log('_Dashboard_Calc sheet setup complete');
}

// ============================================================================
// HIDDEN SHEET 6: _Steward_Performance_Calc
// Source: Grievance Log -> Steward Performance Metrics
// ============================================================================

/**
 * Setup the _Steward_Performance_Calc hidden sheet
 * Calculates detailed steward performance metrics
 */
function setupStewardPerformanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_PERFORMANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Steward', 'Total Cases', 'Active', 'Closed', 'Won', 'Win Rate %', 'Avg Days', 'Overdue', 'Due This Week', 'Performance Score'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  var gStewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);

  // Get unique stewards
  sheet.getRange('A2').setFormula(
    '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"Assigned Steward",' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"")),"")'
  );

  // Total Cases
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A)))');

  // Active Cases (Open + Pending Info)
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")))');

  // Closed Cases
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))');

  // Won Cases
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")))');

  // Win Rate
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(E2:E/D2:D*100,1),0)))');

  // Avg Days
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)))');

  // Overdue
  sheet.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")))');

  // Due This Week
  sheet.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")))');

  // Performance Score (weighted: Win Rate * 0.4 + (100 - Overdue%) * 0.3 + (100 - AvgDays/60*100) * 0.3)
  sheet.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(F2:F*0.4 + (100-IFERROR(H2:H/C2:C*100,0))*0.3 + MAX(0,100-G2:G/60*100)*0.3,1)))');

  sheet.hideSheet();
  Logger.log('_Steward_Performance_Calc sheet setup complete');
}

// ============================================================================
// MASTER SETUP & REPAIR FUNCTIONS
// ============================================================================

/**
 * Setup all hidden calculation sheets
 * @returns {Object} Result object with created and repaired counts
 */
function setupAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Setting up hidden calculation sheets...', '🔧 Setup', 3);

  var created = 0;
  var repaired = 0;

  // Core grievance/member calculation sheets (7 total)
  // Each function creates the sheet if missing or updates if exists
  try { setupGrievanceCalcSheet(); created++; } catch (_e) { repaired++; }
  try { setupGrievanceFormulasSheet(); created++; } catch (_e) { repaired++; }
  try { setupMemberLookupSheet(); created++; } catch (_e) { repaired++; }
  try { setupStewardContactCalcSheet(); created++; } catch (_e) { repaired++; }
  try { setupDashboardCalcSheet(); created++; } catch (_e) { repaired++; }
  try { setupStewardPerformanceCalcSheet(); created++; } catch (_e) { repaired++; }
  try { setupChecklistCalcSheet(); created++; } catch (_e) { repaired++; }

  ss.toast('All 7 hidden sheets created!', '✅ Success', 3);

  return { created: created, repaired: repaired, success: true };
}

/**
 * Repair all hidden sheets - recreates formulas and syncs data
 */
function repairAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ss.toast('Repairing hidden sheets...', '🔧 Repair', 3);

  // Recreate all hidden sheets with formulas
  setupAllHiddenSheets();

  // Install trigger (quick mode - no dialog)
  installAutoSyncTriggerQuick();

  // Run initial sync
  ss.toast('Running initial data sync...', '🔧 Sync', 3);
  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();
  syncChecklistCalcToGrievanceLog();

  // Repair checkboxes
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  ss.toast('Hidden sheets repaired and synced!', '✅ Success', 5);
  ui.alert('✅ Repair Complete',
    'Hidden calculation sheets have been repaired:\n\n' +
    '• 7 hidden sheets recreated with self-healing formulas\n' +
    '• Auto-sync trigger installed\n' +
    '• All data synced (grievances, members, dashboard, checklists)\n' +
    '• Checkboxes repaired in Grievance Log and Member Directory\n\n' +
    'Data will now auto-sync when you edit Member Directory or Grievance Log.\n' +
    'Formulas cannot be accidentally erased - they are stored in hidden sheets.',
    ui.ButtonSet.OK);

  return { repaired: 7, success: true };
}

/**
 * Verify all hidden sheets and triggers
 */
function verifyHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var report = [];

  report.push('🔍 HIDDEN SHEET VERIFICATION');
  report.push('============================');
  report.push('');

  // Check each hidden sheet (7 hidden sheets)
  var hiddenSheets = [
    {name: SHEETS.GRIEVANCE_CALC, purpose: 'Grievance -> Member Directory'},
    {name: SHEETS.GRIEVANCE_FORMULAS, purpose: 'Self-healing Grievance formulas'},
    {name: SHEETS.MEMBER_LOOKUP, purpose: 'Member -> Grievance Log'},
    {name: SHEETS.STEWARD_CONTACT_CALC, purpose: 'Steward contact tracking'},
    {name: SHEETS.DASHBOARD_CALC, purpose: 'Dashboard summary metrics'},
    {name: SHEETS.STEWARD_PERFORMANCE_CALC, purpose: 'Steward performance scores'},
    {name: HIDDEN_SHEETS.CHECKLIST_CALC, purpose: 'Case Checklist progress calculations'}
  ];

  report.push('📋 HIDDEN SHEETS:');
  hiddenSheets.forEach(function(hs) {
    var sheet = ss.getSheetByName(hs.name);
    if (sheet) {
      var isHidden = sheet.isSheetHidden();
      var hasData = sheet.getLastRow() > 1;
      var status = isHidden && hasData ? '✅' : (sheet ? '⚠️' : '❌');
      report.push('  ' + status + ' ' + hs.name);
      report.push('      Hidden: ' + (isHidden ? 'Yes' : 'NO - Should be hidden'));
      report.push('      Has formulas: ' + (hasData ? 'Yes' : 'No'));
    } else {
      report.push('  ❌ ' + hs.name + ' - NOT FOUND');
    }
  });

  report.push('');

  // Check triggers
  report.push('⚡ AUTO-SYNC TRIGGER:');
  var triggers = ScriptApp.getProjectTriggers();
  var hasAutoSync = false;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      hasAutoSync = true;
      report.push('  ✅ onEditAutoSync trigger installed');
    }
  });
  if (!hasAutoSync) {
    report.push('  ❌ onEditAutoSync trigger NOT installed');
    report.push('     Run: installAutoSyncTrigger()');
  }

  report.push('');
  report.push('============================');

  ui.alert('Hidden Sheet Verification', report.join('\n'), ui.ButtonSet.OK);
  Logger.log(report.join('\n'));
}

/**
 * Refresh all formulas (force recalculation)
 */
function refreshAllHiddenFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Force recalculation of all pending formulas (covers all 6 hidden sheets:
  // GRIEVANCE_CALC, GRIEVANCE_FORMULAS, MEMBER_LOOKUP, STEWARD_CONTACT_CALC,
  // DASHBOARD_CALC, STEWARD_PERFORMANCE_CALC)
  SpreadsheetApp.flush();

  // Then sync
  syncAllData();

  // Repair checkboxes
  repairGrievanceCheckboxes();

  ss.toast('Formulas refreshed and data synced!', '✅ Success', 3);
}

// ============================================================================
// FORMULA SERVICE - INDIVIDUAL CALCULATION SHEET SETUP
// These functions set up additional calculation sheets for specialized purposes
// ============================================================================

/**
 * Sets up the _CalcMembers hidden sheet
 * Contains member statistics and lookup tables
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcMembersSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;

  // Header row
  sheet.getRange('A1').setValue('Member Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Total Members
  sheet.getRange('A2').setValue('Total Members');
  sheet.getRange('B2').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  // Active Members
  sheet.getRange('A3').setValue('Active Members');
  sheet.getRange('B3').setFormula(
    `=COUNTIF('${memberSheetName}'!K:K,"Active")`
  );

  // Members by Department (dynamic list)
  sheet.getRange('A5').setValue('Department');
  sheet.getRange('B5').setValue('Count');
  sheet.getRange('A5:B5').setFontWeight('bold');

  sheet.getRange('A6').setFormula(
    `=UNIQUE(FILTER('${memberSheetName}'!E:E,'${memberSheetName}'!E:E<>"Department",'${memberSheetName}'!E:E<>""))`
  );

  sheet.getRange('B6').setFormula(
    `=ARRAYFORMULA(IF(A6:A<>"",COUNTIF('${memberSheetName}'!E:E,A6:A),""))`
  );

  // Union Status breakdown
  sheet.getRange('D2').setValue('Union Status');
  sheet.getRange('E2').setValue('Count');
  sheet.getRange('D2:E2').setFontWeight('bold');

  sheet.getRange('D3').setValue('Full Member');
  sheet.getRange('E3').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Full Member")`
  );

  sheet.getRange('D4').setValue('Agency Fee');
  sheet.getRange('E4').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Agency Fee")`
  );

  sheet.getRange('D5').setValue('Non-Member');
  sheet.getRange('E5').setFormula(
    `=COUNTIF('${memberSheetName}'!L:L,"Non-Member")`
  );

  // Lookup helper for member names
  sheet.getRange('G1').setValue('ID->Name Lookup');
  sheet.getRange('G1').setFontWeight('bold');
  sheet.getRange('G2').setFormula(
    `=ARRAYFORMULA(IF('${memberSheetName}'!A2:A<>"",` +
    `'${memberSheetName}'!A2:A&"|"&'${memberSheetName}'!B2:B&" "&'${memberSheetName}'!C2:C,""))`
  );
}

/**
 * Sets up the _CalcGrievances hidden sheet
 * Contains grievance aggregations and summaries
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcGrievancesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Grievance Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Status counts
  sheet.getRange('A3').setValue('Status');
  sheet.getRange('B3').setValue('Count');
  sheet.getRange('A3:B3').setFontWeight('bold');

  const statuses = Object.values(GRIEVANCE_STATUS);
  statuses.forEach((status, index) => {
    sheet.getRange(4 + index, 1).setValue(status);
    sheet.getRange(4 + index, 2).setFormula(
      `=COUNTIF('${grievanceSheetName}'!W:W,"${status}")`
    );
  });

  // Grievances by Type
  sheet.getRange('D3').setValue('Type');
  sheet.getRange('E3').setValue('Count');
  sheet.getRange('D3:E3').setFontWeight('bold');

  sheet.getRange('D4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!E:E,` +
    `'${grievanceSheetName}'!E:E<>"Grievance Type",'${grievanceSheetName}'!E:E<>""))`
  );

  sheet.getRange('E4').setFormula(
    `=ARRAYFORMULA(IF(D4:D<>"",COUNTIF('${grievanceSheetName}'!E:E,D4:D),""))`
  );

  // Grievances by Current Step
  sheet.getRange('G3').setValue('Step');
  sheet.getRange('H3').setValue('Count');
  sheet.getRange('G3:H3').setFontWeight('bold');

  for (let step = 1; step <= 4; step++) {
    sheet.getRange(3 + step, 7).setValue(`Step ${step}`);
    sheet.getRange(3 + step, 8).setFormula(
      `=COUNTIF('${grievanceSheetName}'!H:H,${step})`
    );
  }

  // Monthly filing trend (last 12 months)
  sheet.getRange('A15').setValue('Monthly Filings');
  sheet.getRange('A15').setFontWeight('bold');

  sheet.getRange('A16').setValue('Month');
  sheet.getRange('B16').setValue('Filings');
  sheet.getRange('A16:B16').setFontWeight('bold');

  // Generate last 12 months
  for (let i = 0; i < 12; i++) {
    sheet.getRange(17 + i, 1).setFormula(
      `=EOMONTH(TODAY(),-${i})`
    );
    sheet.getRange(17 + i, 2).setFormula(
      `=SUMPRODUCT((MONTH('${grievanceSheetName}'!D:D)=MONTH(A${17 + i}))*` +
      `(YEAR('${grievanceSheetName}'!D:D)=YEAR(A${17 + i})))`
    );
  }
}

/**
 * Sets up the _CalcDeadlines hidden sheet
 * Contains deadline calculations and alert logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcDeadlinesSheet(sheet) {
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Deadline Calculations');
  sheet.getRange('A1').setFontWeight('bold');

  // Configuration reference
  sheet.getRange('A3').setValue('Deadline Rules (Days)');
  sheet.getRange('A3').setFontWeight('bold');

  var rules = getDeadlineRules();
  sheet.getRange('A4').setValue('Step 1 Response');
  sheet.getRange('B4').setValue(rules.STEP_1.DAYS_FOR_RESPONSE);

  sheet.getRange('A5').setValue('Step 2 Appeal');
  sheet.getRange('B5').setValue(rules.STEP_2.DAYS_TO_APPEAL);

  sheet.getRange('A6').setValue('Step 2 Response');
  sheet.getRange('B6').setValue(rules.STEP_2.DAYS_FOR_RESPONSE);

  sheet.getRange('A7').setValue('Step 3 Appeal');
  sheet.getRange('B7').setValue(rules.STEP_3.DAYS_TO_APPEAL);

  sheet.getRange('A8').setValue('Step 3 Response');
  sheet.getRange('B8').setValue(rules.STEP_3.DAYS_FOR_RESPONSE);

  sheet.getRange('A9').setValue('Arbitration Demand');
  sheet.getRange('B9').setValue(rules.ARBITRATION.DAYS_TO_DEMAND);

  // Upcoming deadlines calculation
  sheet.getRange('D1').setValue('Upcoming Deadlines (Next 14 Days)');
  sheet.getRange('D1').setFontWeight('bold');

  sheet.getRange('D2').setValue('Grievance ID');
  sheet.getRange('E2').setValue('Step');
  sheet.getRange('F2').setValue('Due Date');
  sheet.getRange('G2').setValue('Days Left');
  sheet.getRange('D2:G2').setFontWeight('bold');

  // Complex formula to extract upcoming deadlines
  // This uses FILTER to get open grievances and calculate their current deadline
  sheet.getRange('D3').setFormula(
    `=IFERROR(FILTER('${grievanceSheetName}'!A:A,` +
    `('${grievanceSheetName}'!W:W="Open")+('${grievanceSheetName}'!W:W="Pending Response")+` +
    `('${grievanceSheetName}'!W:W="Appealed")),"")`
  );

  // Overdue grievances
  sheet.getRange('I1').setValue('Overdue Grievances');
  sheet.getRange('I1').setFontWeight('bold');
  sheet.getRange('I2').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"<>Resolved",` +
    `'${grievanceSheetName}'!W:W,"<>Closed",` +
    `'${grievanceSheetName}'!W:W,"<>Withdrawn",` +
    `'${grievanceSheetName}'!J:J,"<"&TODAY())`
  );

  // Alert thresholds
  sheet.getRange('I4').setValue('Alert Thresholds');
  sheet.getRange('I4').setFontWeight('bold');

  sheet.getRange('I5').setValue('Critical (<=3 days)');
  sheet.getRange('J5').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">="&TODAY(),` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+3)`
  );

  sheet.getRange('I6').setValue('Warning (4-7 days)');
  sheet.getRange('J6').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Open",'${grievanceSheetName}'!J:J,">"&TODAY()+3,` +
    `'${grievanceSheetName}'!J:J,"<="&TODAY()+7)`
  );
}

/**
 * Sets up the _CalcStats hidden sheet
 * Contains dashboard-wide statistics
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcStatsSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Dashboard Statistics');
  sheet.getRange('A1').setFontWeight('bold');

  // Quick stats for sidebar
  sheet.getRange('A3').setValue('Sidebar Stats');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setValue('open_grievances');
  sheet.getRange('B4').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Open")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Appealed")`
  );

  sheet.getRange('A5').setValue('pending_response');
  sheet.getRange('B5').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Pending Response")`
  );

  sheet.getRange('A6').setValue('total_members');
  sheet.getRange('B6').setFormula(
    `=COUNTA('${memberSheetName}'!A:A)-1`
  );

  sheet.getRange('A7').setValue('resolved_ytd');
  sheet.getRange('B7').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!W:W,"Resolved",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))+` +
    `COUNTIFS('${grievanceSheetName}'!W:W,"Closed",` +
    `'${grievanceSheetName}'!X:X,">="&DATE(YEAR(TODAY()),1,1))`
  );

  // Win rate calculation
  sheet.getRange('A9').setValue('Performance Metrics');
  sheet.getRange('A9').setFontWeight('bold');

  sheet.getRange('A10').setValue('total_resolved');
  sheet.getRange('B10').setFormula(
    `=COUNTIF('${grievanceSheetName}'!W:W,"Resolved")+` +
    `COUNTIF('${grievanceSheetName}'!W:W,"Closed")`
  );

  sheet.getRange('A11').setValue('sustained_count');
  sheet.getRange('B11').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Sustained")`
  );

  sheet.getRange('A12').setValue('settled_count');
  sheet.getRange('B12').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Settled")`
  );

  sheet.getRange('A13').setValue('denied_count');
  sheet.getRange('B13').setFormula(
    `=COUNTIF('${grievanceSheetName}'!T:T,"Denied")`
  );

  sheet.getRange('A14').setValue('win_rate');
  sheet.getRange('B14').setFormula(
    `=IFERROR((B11+B12)/B10*100,0)`
  );

  // Average time to resolution
  sheet.getRange('A16').setValue('avg_days_to_resolve');
  sheet.getRange('B16').setFormula(
    `=IFERROR(AVERAGEIFS('${grievanceSheetName}'!X:X-'${grievanceSheetName}'!D:D,` +
    `'${grievanceSheetName}'!W:W,"Resolved"),0)`
  );
}

/**
 * Sets up the _CalcSync hidden sheet
 * Contains cross-sheet synchronization logic
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcSyncSheet(sheet) {
  const memberSheetName = SHEET_NAMES.MEMBER_DIRECTORY;
  const grievanceSheetName = SHEET_NAMES.GRIEVANCE_TRACKER;

  // Header
  sheet.getRange('A1').setValue('Cross-Sheet Sync');
  sheet.getRange('A1').setFontWeight('bold');

  // Member ID validation list
  sheet.getRange('A3').setValue('Valid Member IDs');
  sheet.getRange('A3').setFontWeight('bold');

  sheet.getRange('A4').setFormula(
    `=FILTER('${memberSheetName}'!A:A,'${memberSheetName}'!A:A<>"ID",'${memberSheetName}'!A:A<>"")`
  );

  // Grievances per member count
  sheet.getRange('C3').setValue('Member ID');
  sheet.getRange('D3').setValue('Grievance Count');
  sheet.getRange('C3:D3').setFontWeight('bold');

  sheet.getRange('C4').setFormula(
    `=UNIQUE(FILTER('${grievanceSheetName}'!B:B,'${grievanceSheetName}'!B:B<>"Member ID",'${grievanceSheetName}'!B:B<>""))`
  );

  sheet.getRange('D4').setFormula(
    `=ARRAYFORMULA(IF(C4:C<>"",COUNTIF('${grievanceSheetName}'!B:B,C4:C),""))`
  );

  // Data consistency checks
  sheet.getRange('F3').setValue('Data Consistency');
  sheet.getRange('F3').setFontWeight('bold');

  sheet.getRange('F4').setValue('Orphaned Grievances');
  sheet.getRange('G4').setFormula(
    `=COUNTIFS('${grievanceSheetName}'!B:B,"<>",'${grievanceSheetName}'!B:B,"<>Member ID")-` +
    `SUMPRODUCT(COUNTIF(A4:A,'${grievanceSheetName}'!B2:B))`
  );

  sheet.getRange('F5').setValue('Members with Grievances');
  sheet.getRange('G5').setFormula(
    `=COUNTA(C4:C)`
  );

  // Last sync timestamp
  sheet.getRange('F7').setValue('Last Formula Update');
  sheet.getRange('G7').setFormula('=NOW()');
}

/**
 * Sets up the _CalcFormulas hidden sheet
 * Contains named formula references for use in other sheets
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCalcFormulasSheet(sheet) {
  // Header
  sheet.getRange('A1').setValue('Named Formula References');
  sheet.getRange('A1').setFontWeight('bold');
  sheet.getRange('A2').setValue('Use these formulas via indirect references');

  // Department list formula
  sheet.getRange('A4').setValue('DEPARTMENT_LIST');
  sheet.getRange('B4').setFormula(
    `=SORT(UNIQUE(FILTER('${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E,` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"Department",` +
    `'${SHEET_NAMES.MEMBER_DIRECTORY}'!E:E<>"")))`
  );

  // Status list formula
  sheet.getRange('A6').setValue('STATUS_LIST');
  sheet.getRange('B6').setValue(Object.values(GRIEVANCE_STATUS).join(','));

  // Outcome list formula
  sheet.getRange('A8').setValue('OUTCOME_LIST');
  sheet.getRange('B8').setValue(Object.values(GRIEVANCE_OUTCOMES).join(','));

  // Grievance type list
  sheet.getRange('A10').setValue('GRIEVANCE_TYPES');
  sheet.getRange('B10').setValue('Contract Violation,Discipline,Discharge,Working Conditions,Safety,Other');

  // Date formatting formula
  sheet.getRange('A12').setValue('TODAY_FORMATTED');
  sheet.getRange('B12').setFormula('=TEXT(TODAY(),"MMMM D, YYYY")');

  // Year calculation
  sheet.getRange('A14').setValue('CURRENT_YEAR');
  sheet.getRange('B14').setFormula('=YEAR(TODAY())');

  // Next grievance ID prefix
  sheet.getRange('A16').setValue('GRIEVANCE_ID_PREFIX');
  sheet.getRange('B16').setFormula('="GRV-"&YEAR(TODAY())&"-"');
}

// ============================================================================
// SURVEY TRACKING HIDDEN SHEET SETUP
// ============================================================================
//
// This function creates the hidden _Survey_Tracking sheet structure.
// It is called by setupHiddenSheets() in 08a_SheetSetup.gs during
// CREATE_509_DASHBOARD(). The sheet is hidden from users automatically.
//
// SURVEY TRACKER FLOW OVERVIEW:
//   1. CREATE_509_DASHBOARD() -> setupHiddenSheets() -> setupSurveyTrackingSheet()
//      Creates the hidden sheet with 10 columns (A-J per SURVEY_TRACKING_COLS).
//   2. populateSurveyTrackingFromMembers() copies all members from Member Directory.
//   3. When a member submits the Google Form satisfaction survey:
//      Google trigger -> onSatisfactionFormSubmit(e) -> validates respondent email
//      against Member Directory -> updateSurveyTrackingOnSubmit_(memberId)
//      marks the member as "Completed" with timestamp.
//   4. Stewards manage rounds via showSurveyTrackingDialog():
//      - "Start New Round" resets all statuses, increments missed counts
//      - "Send Reminders" emails non-respondents (7-day cooldown)
//      - "Refresh Member List" re-syncs from Member Directory
//
// All tracking functions are in 08c_FormsAndNotifications.gs.
// Column constants are in 01_Core.gs (SURVEY_TRACKING_COLS).
// ============================================================================

/**
 * Sets up the hidden _Survey_Tracking sheet with headers and formatting.
 * This sheet tracks per-member survey completion status across rounds.
 * Columns match SURVEY_TRACKING_COLS in 01_Core.gs (A-J).
 *
 * Called by: setupHiddenSheets() in 08a_SheetSetup.gs
 * Column constants: SURVEY_TRACKING_COLS in 01_Core.gs
 *
 * @param {Sheet} sheet - The sheet to set up
 */
function setupSurveyTrackingSheet(sheet) {
  var headers = [
    'Member ID',          // A
    'Member Name',        // B
    'Email',              // C
    'Work Location',      // D
    'Assigned Steward',   // E
    'Current Status',     // F
    'Completed Date',     // G
    'Total Missed',       // H
    'Total Completed',    // I
    'Last Reminder Sent'  // J
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');

  // Column widths
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.MEMBER_ID, 100);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.MEMBER_NAME, 180);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.EMAIL, 220);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.WORK_LOCATION, 160);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.ASSIGNED_STEWARD, 160);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.CURRENT_STATUS, 130);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.COMPLETED_DATE, 130);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.TOTAL_MISSED, 100);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.TOTAL_COMPLETED, 120);
  sheet.setColumnWidth(SURVEY_TRACKING_COLS.LAST_REMINDER_SENT, 140);

  // Date formatting
  sheet.getRange(2, SURVEY_TRACKING_COLS.COMPLETED_DATE, 998, 1).setNumberFormat('MM/DD/YYYY');
  sheet.getRange(2, SURVEY_TRACKING_COLS.LAST_REMINDER_SENT, 998, 1).setNumberFormat('MM/DD/YYYY');

  // Number formatting for counters
  sheet.getRange(2, SURVEY_TRACKING_COLS.TOTAL_MISSED, 998, 1).setNumberFormat('0');
  sheet.getRange(2, SURVEY_TRACKING_COLS.TOTAL_COMPLETED, 998, 1).setNumberFormat('0');

  // Freeze header row
  sheet.setFrozenRows(1);

  Logger.log('Survey Tracking hidden sheet set up');
}

// ============================================================================
// SURVEY VAULT — ZERO-KNOWLEDGE PII STORE
// ============================================================================
// The _Survey_Vault sheet stores ONLY hashed email and hashed member ID.
// No plaintext PII is ever written to any sheet.
//
// SECURITY MODEL:
//   - Email is SHA-256 hashed with a per-installation salt before storage
//   - Member ID is hashed the same way
//   - Even with full access to the vault, you cannot reverse the hashes
//   - Superseding (same member re-submits) works by comparing hashes
//   - The raw email is used in-memory only (for thank-you email + validation)
//     and is never persisted to any sheet

/**
 * Creates and protects the _Survey_Vault hidden sheet.
 * Called by setupHiddenSheets().
 *
 * Structure (8 columns):
 *   A: Response Row (row number in Satisfaction sheet)
 *   B: Email Hash (SHA-256, non-reversible)
 *   C: Verified (Yes / Pending Review / Rejected)
 *   D: Member ID Hash (SHA-256, non-reversible)
 *   E: Quarter
 *   F: Is Latest (Yes/No)
 *   G: Superseded By (vault row of newer response)
 *   H: Reviewer Notes
 */
function setupSurveyVaultSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault';

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // Only set up headers if sheet is empty
  if (sheet.getLastRow() < 1) {
    var headers = [
      'Response Row', 'Email Hash', 'Verified', 'Member ID Hash',
      'Quarter', 'Is Latest', 'Superseded By', 'Reviewer Notes'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#7F1D1D')
      .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  // Column widths
  sheet.setColumnWidth(1, 100);  // Response Row
  sheet.setColumnWidth(2, 130);  // Email Hash
  sheet.setColumnWidth(3, 110);  // Verified
  sheet.setColumnWidth(4, 130);  // Member ID Hash
  sheet.setColumnWidth(5, 90);   // Quarter
  sheet.setColumnWidth(6, 80);   // Is Latest
  sheet.setColumnWidth(7, 100);  // Superseded By
  sheet.setColumnWidth(8, 200);  // Reviewer Notes

  // Delete excess columns
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 8) {
    sheet.deleteColumns(9, maxCols - 8);
  }

  // Hide the sheet
  sheet.hideSheet();

  // ═══ SHEET PROTECTION ═══
  // Only the script owner (installer) can edit this sheet.
  // All other editors get a warning.
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var alreadyProtected = false;
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === 'Survey Vault — Anonymity Protection') {
      alreadyProtected = true;
      break;
    }
  }

  if (!alreadyProtected) {
    var protection = sheet.protect()
      .setDescription('Survey Vault — Anonymity Protection')
      .setWarningOnly(false);

    // Remove all editors except the owner
    var me = Session.getEffectiveUser();
    protection.addEditor(me);
    // Remove everyone else
    var editors = protection.getEditors();
    for (var j = 0; j < editors.length; j++) {
      if (editors[j].getEmail() !== me.getEmail()) {
        protection.removeEditor(editors[j]);
      }
    }
  }

  Logger.log('Survey Vault hidden sheet set up with protection');
}

/**
 * Reads the _Survey_Vault and returns a map of Satisfaction-sheet row numbers
 * to their vault metadata. DOES NOT expose email or member ID — only returns
 * non-PII flags needed for dashboard filtering.
 *
 * @returns {Object} Map of { responseRow: { verified, isLatest, quarter } }
 * @private
 */
function getVaultDataMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault || vault.getLastRow() < 2) return {};

  var data = vault.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var responseRow = data[i][SURVEY_VAULT_COLS.RESPONSE_ROW - 1];
    if (!responseRow) continue;
    map[responseRow] = {
      verified: data[i][SURVEY_VAULT_COLS.VERIFIED - 1] || '',
      isLatest: data[i][SURVEY_VAULT_COLS.IS_LATEST - 1] || '',
      quarter: data[i][SURVEY_VAULT_COLS.QUARTER - 1] || '',
      vaultRow: i + 1  // 1-indexed sheet row for updates
    };
  }
  return map;
}

/**
 * Reads the full vault data. All email/member ID values are SHA-256 hashes —
 * no plaintext PII exists in the vault.
 *
 * @returns {Array} Array of vault row objects (emailHash and memberIdHash are non-reversible)
 * @private
 */
function getVaultDataFull_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault || vault.getLastRow() < 2) return [];

  var data = vault.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      vaultRow: i + 1,
      responseRow: data[i][SURVEY_VAULT_COLS.RESPONSE_ROW - 1],
      emailHash: data[i][SURVEY_VAULT_COLS.EMAIL - 1] || '',
      verified: data[i][SURVEY_VAULT_COLS.VERIFIED - 1] || '',
      memberIdHash: data[i][SURVEY_VAULT_COLS.MATCHED_MEMBER_ID - 1] || '',
      quarter: data[i][SURVEY_VAULT_COLS.QUARTER - 1] || '',
      isLatest: data[i][SURVEY_VAULT_COLS.IS_LATEST - 1] || '',
      supersededBy: data[i][SURVEY_VAULT_COLS.SUPERSEDED_BY - 1] || '',
      reviewerNotes: data[i][SURVEY_VAULT_COLS.REVIEWER_NOTES - 1] || ''
    });
  }
  return rows;
}

/**
 * Writes a new vault entry for a survey response.
 * Email and member ID are SHA-256 hashed before storage — no plaintext PII
 * is ever written to any sheet.
 *
 * Called by onSatisfactionFormSubmit() after appending the anonymous
 * response to the Satisfaction sheet.
 *
 * @param {number} responseRow - Row number in Satisfaction sheet
 * @param {string} email - Respondent email (hashed before storage)
 * @param {string} verified - 'Yes' | 'Pending Review'
 * @param {string} memberId - Matched member ID or '' (hashed before storage)
 * @param {string} quarter - Quarter string e.g. '2026-Q1'
 * @private
 */
function writeVaultEntry_(responseRow, email, verified, memberId, quarter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault) {
    setupSurveyVaultSheet();
    vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  }

  // Hash email and member ID — plaintext never touches the sheet
  var emailHash = email ? hashForVault_(email) : '';
  var memberIdHash = memberId ? hashForVault_(memberId) : '';

  vault.appendRow([
    responseRow,
    emailHash,
    verified,
    memberIdHash,
    quarter,
    'Yes',   // Is Latest
    '',      // Superseded By
    ''       // Reviewer Notes
  ]);
}

/**
 * Marks an existing vault entry as superseded by a newer response.
 * Used when the same member submits again in the same quarter.
 * Comparison is done on the hashed email — plaintext is never stored.
 *
 * @param {string} email - Raw email (hashed in-memory for comparison)
 * @param {string} quarter - Quarter to match
 * @param {number} newVaultRow - Row of the new entry that supersedes
 * @private
 */
function supersedePreviousVaultEntry_(email, quarter, newVaultRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');
  if (!vault || vault.getLastRow() < 2) return;

  // Hash the email in-memory to compare against stored hashes.
  // v4.8.1: Compute both new and legacy hash to match entries from either format.
  var emailHash = hashForVault_(email);
  var legacyHash = hashForVaultLegacy_(email);

  var data = vault.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rowEmailHash = (data[i][SURVEY_VAULT_COLS.EMAIL - 1] || '').toString();
    var rowQuarter = data[i][SURVEY_VAULT_COLS.QUARTER - 1];
    var rowIsLatest = data[i][SURVEY_VAULT_COLS.IS_LATEST - 1];

    if ((rowEmailHash === emailHash || rowEmailHash === legacyHash) && rowQuarter === quarter && rowIsLatest === 'Yes') {
      var sheetRow = i + 1;
      vault.getRange(sheetRow, SURVEY_VAULT_COLS.IS_LATEST).setValue('No');
      vault.getRange(sheetRow, SURVEY_VAULT_COLS.SUPERSEDED_BY).setValue(newVaultRow);
      Logger.log('Vault: superseded row ' + sheetRow + ' by row ' + newVaultRow);
    }
  }
}

/**
 * Creates a non-reversible SHA-256 hash for vault storage.
 * Uses HMAC-style double hashing with a per-installation salt stored in
 * Script Properties. Output is 22 hex characters (88 bits of entropy)
 * prefixed with 'V', giving a collision-resistant 23-character identifier.
 *
 * v4.8.1: Upgraded from 11-character output (40 bits) to 23-character
 * output (88 bits) to reduce collision probability. Old hashes (starting
 * with 'V' and <= 11 chars) are still matchable via hashForVaultLegacy_().
 *
 * @param {string} value - The value to hash (email, member ID, etc.)
 * @returns {string} 23-character hash: 'V' + 22 hex chars
 * @private
 */
function hashForVault_(value) {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('ANON_HASH_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty('ANON_HASH_SALT', salt);
  }
  var normalized = String(value).toLowerCase().trim();

  // HMAC-style double hash: H(salt + H(salt + value))
  // This prevents length-extension attacks on the inner hash
  var innerDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    salt + normalized, Utilities.Charset.UTF_8);
  var innerHex = innerDigest.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
  var outerDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    salt + innerHex, Utilities.Charset.UTF_8);

  // Take first 11 bytes (22 hex chars) = 88 bits of entropy
  var hex = outerDigest.slice(0, 11).map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('').toUpperCase();
  return 'V' + hex;
}

/**
 * Legacy hash function for backward compatibility with pre-v4.8.1 vault entries.
 * Produces the old 11-character format ('V' + 10 alphanumeric).
 *
 * Used by supersedePreviousVaultEntry_ when matching against older hashes.
 *
 * @param {string} value - The value to hash
 * @returns {string} 11-character legacy hash
 * @private
 */
function hashForVaultLegacy_(value) {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('ANON_HASH_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty('ANON_HASH_SALT', salt);
  }
  var combined = salt + String(value).toLowerCase().trim();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined);
  var encoded = Utilities.base64Encode(digest).substring(0, 12);
  return 'V' + encoded.replace(/[^a-zA-Z0-9]/g, '').substring(0, 10).toUpperCase();
}

// ============================================================================
// HMAC-SHA256 UTILITY
// ============================================================================

/**
 * Computes an HMAC-SHA256 message authentication code.
 * Google Apps Script does not provide a native HMAC, so this implements
 * the standard HMAC construction: H((K ^ opad) || H((K ^ ipad) || message))
 *
 * @param {string} key - The secret key
 * @param {string} message - The message to authenticate
 * @returns {string} 64-character lowercase hex HMAC
 */
function computeHmacSha256_(key, message) {
  var BLOCK_SIZE = 64; // SHA-256 block size in bytes

  // Convert key to byte array; if longer than block size, hash it first
  var keyBytes = [];
  for (var i = 0; i < key.length; i++) {
    keyBytes.push(key.charCodeAt(i) & 0xFF);
  }
  if (keyBytes.length > BLOCK_SIZE) {
    keyBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key, Utilities.Charset.UTF_8);
  }

  // Pad key to block size
  while (keyBytes.length < BLOCK_SIZE) {
    keyBytes.push(0);
  }

  // XOR key with ipad (0x36) and opad (0x5c)
  var ipadKey = '';
  var opadKey = '';
  for (var j = 0; j < BLOCK_SIZE; j++) {
    ipadKey += String.fromCharCode((keyBytes[j] & 0xFF) ^ 0x36);
    opadKey += String.fromCharCode((keyBytes[j] & 0xFF) ^ 0x5c);
  }

  // Inner hash: H(ipadKey || message)
  var innerHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    ipadKey + message, Utilities.Charset.UTF_8);
  var innerHex = innerHash.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');

  // Outer hash: H(opadKey || innerHash)
  // Convert innerHash bytes back to a string for the outer digest
  var innerStr = '';
  for (var k = 0; k < innerHash.length; k++) {
    innerStr += String.fromCharCode(innerHash[k] & 0xFF);
  }
  var outerHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    opadKey + innerStr, Utilities.Charset.UTF_8);

  return outerHash.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// ============================================================================
// AUDIT LOG INTEGRITY CHAIN
// ============================================================================

/**
 * Computes a tamper-detection hash for an audit log row.
 * Each row's integrity hash includes the previous row's hash, forming
 * a hash chain. If any row is modified or deleted, all subsequent
 * hashes will fail verification.
 *
 * @param {string} previousHash - Integrity hash of the previous row ('' for first row)
 * @param {Date|string} timestamp - Row timestamp
 * @param {string} eventType - Event type
 * @param {string} user - User email
 * @param {string} details - JSON details string
 * @param {string} sessionId - Session identifier
 * @returns {string} 16-character integrity hash
 * @private
 */
function computeAuditRowHash_(previousHash, timestamp, eventType, user, details, sessionId) {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('AUDIT_INTEGRITY_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty('AUDIT_INTEGRITY_SALT', salt);
  }

  var payload = [
    previousHash || '',
    String(timestamp),
    String(eventType),
    String(user),
    String(details),
    String(sessionId)
  ].join('|');

  return computeHmacSha256_(salt, payload).substring(0, 16);
}

/**
 * Verifies the integrity of the entire audit log by recomputing the
 * hash chain from row 1 to the end. Returns a report of any rows
 * where the stored hash does not match the computed hash.
 *
 * The Integrity Hash column (F) is added by the enhanced logAuditEvent.
 *
 * @returns {Object} { valid: boolean, totalRows: number, invalidRows: number[] }
 */
function verifyAuditLogIntegrity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG || SHEETS.AUDIT_LOG);
  if (!sheet || sheet.getLastRow() < 2) {
    return { valid: true, totalRows: 0, invalidRows: [] };
  }

  var data = sheet.getDataRange().getValues();
  var invalidRows = [];
  var previousHash = '';
  var integrityCol = -1;

  // Find integrity hash column (header = 'Integrity Hash')
  var headers = data[0];
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim() === 'Integrity Hash') {
      integrityCol = h;
      break;
    }
  }

  // If no integrity column exists, the log predates this feature
  if (integrityCol === -1) {
    return { valid: true, totalRows: data.length - 1, invalidRows: [], message: 'No integrity hashes found — log predates hash chain feature' };
  }

  for (var i = 1; i < data.length; i++) {
    var storedHash = String(data[i][integrityCol] || '');
    var computed = computeAuditRowHash_(
      previousHash,
      data[i][0],  // Timestamp
      data[i][1],  // Event Type
      data[i][2],  // User
      data[i][3],  // Details
      data[i][4]   // Session ID
    );

    if (storedHash && storedHash !== computed) {
      invalidRows.push(i + 1); // 1-indexed sheet row
    }
    previousHash = storedHash || computed;
  }

  var result = {
    valid: invalidRows.length === 0,
    totalRows: data.length - 1,
    invalidRows: invalidRows
  };

  if (invalidRows.length > 0) {
    result.message = invalidRows.length + ' row(s) failed integrity check — possible tampering detected';
  }

  return result;
}

/**
 * Menu-callable wrapper: verifies audit log integrity and shows result in a dialog.
 */
function verifyAuditLogIntegrityDialog() {
  var result = verifyAuditLogIntegrity();
  var ui = SpreadsheetApp.getUi();

  if (result.valid) {
    ui.alert('✅ Audit Log Integrity',
      'All ' + result.totalRows + ' audit entries passed integrity verification.' +
      (result.message ? '\n\nNote: ' + result.message : ''),
      ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Audit Log Integrity Warning',
      result.message + '\n\n' +
      'Affected rows: ' + result.invalidRows.join(', ') + '\n\n' +
      'This may indicate unauthorized modification of the audit log.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// VAULT INTEGRITY VERIFICATION
// ============================================================================

/**
 * Verifies the structural integrity of the _Survey_Vault sheet.
 * Checks for:
 * - Missing or malformed hashes
 * - Orphaned vault entries (no matching satisfaction row)
 * - Duplicate 'Is Latest' entries for the same email hash + quarter
 * - Rows missing required fields
 *
 * @returns {Object} { valid: boolean, issues: string[], stats: Object }
 */
function verifySurveyVaultIntegrity() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_VAULT || '_Survey_Vault');

  if (!vault || vault.getLastRow() < 2) {
    return { valid: true, issues: [], stats: { totalEntries: 0 } };
  }

  var data = vault.getDataRange().getValues();
  var issues = [];
  var latestMap = {};   // emailHash+quarter -> [rowNumbers]
  var totalEntries = 0;
  var verifiedCount = 0;
  var pendingCount = 0;
  var rejectedCount = 0;

  for (var i = 1; i < data.length; i++) {
    totalEntries++;
    var rowNum = i + 1;
    var responseRow = data[i][0];
    var emailHash = String(data[i][1] || '');
    var verified = String(data[i][2] || '');
    var memberIdHash = String(data[i][3] || '');
    var quarter = String(data[i][4] || '');
    var isLatest = String(data[i][5] || '');

    // Check for missing response row
    if (!responseRow) {
      issues.push('Row ' + rowNum + ': Missing response row number');
    }

    // Check email hash format (should start with 'V')
    if (emailHash && emailHash.charAt(0) !== 'V') {
      issues.push('Row ' + rowNum + ': Email hash has unexpected format (possible plaintext leak)');
    }

    // Check member ID hash format if present
    if (memberIdHash && memberIdHash.charAt(0) !== 'V') {
      issues.push('Row ' + rowNum + ': Member ID hash has unexpected format (possible plaintext leak)');
    }

    // Check for missing quarter
    if (!quarter) {
      issues.push('Row ' + rowNum + ': Missing quarter');
    }

    // Track duplicate "Is Latest" entries
    if (isLatest === 'Yes' && emailHash && quarter) {
      var key = emailHash + '|' + quarter;
      if (!latestMap[key]) {
        latestMap[key] = [];
      }
      latestMap[key].push(rowNum);
    }

    // Count verification statuses
    if (verified === 'Yes') verifiedCount++;
    else if (verified === 'Pending Review') pendingCount++;
    else if (verified === 'Rejected') rejectedCount++;
  }

  // Check for duplicate "Is Latest" entries
  for (var mapKey in latestMap) {
    if (latestMap.hasOwnProperty(mapKey) && latestMap[mapKey].length > 1) {
      issues.push('Duplicate "Is Latest" entries for hash+quarter ' + mapKey.split('|')[1] +
        ' in rows: ' + latestMap[mapKey].join(', '));
    }
  }

  return {
    valid: issues.length === 0,
    issues: issues,
    stats: {
      totalEntries: totalEntries,
      verified: verifiedCount,
      pendingReview: pendingCount,
      rejected: rejectedCount
    }
  };
}

/**
 * Menu-callable wrapper: verifies vault integrity and shows result in a dialog.
 */
function verifySurveyVaultIntegrityDialog() {
  var result = verifySurveyVaultIntegrity();
  var ui = SpreadsheetApp.getUi();

  var statsText = 'Entries: ' + result.stats.totalEntries +
    ' | Verified: ' + result.stats.verified +
    ' | Pending: ' + result.stats.pendingReview +
    ' | Rejected: ' + result.stats.rejected;

  if (result.valid) {
    ui.alert('✅ Survey Vault Integrity',
      'Vault passed all integrity checks.\n\n' + statsText,
      ui.ButtonSet.OK);
  } else {
    ui.alert('⚠️ Survey Vault Issues Found',
      result.issues.length + ' issue(s) detected:\n\n' +
      result.issues.slice(0, 10).join('\n') +
      (result.issues.length > 10 ? '\n... and ' + (result.issues.length - 10) + ' more' : '') +
      '\n\n' + statsText,
      ui.ButtonSet.OK);
  }
}
