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
  }

  sheet.clear();

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

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit log sheet created and hidden.', '✅ Setup Complete', 3);
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
  var gNextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);
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
  var mFirstCol = getColumnLetter(MEMBER_COLS.FIRST_NAME);
  var mLastCol = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var mEmailCol = getColumnLetter(MEMBER_COLS.EMAIL);
  var mUnitCol = getColumnLetter(MEMBER_COLS.UNIT);
  var mLocCol = getColumnLetter(MEMBER_COLS.WORK_LOCATION);
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
  try { setupGrievanceCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupGrievanceFormulasSheet(); created++; } catch (e) { repaired++; }
  try { setupMemberLookupSheet(); created++; } catch (e) { repaired++; }
  try { setupStewardContactCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupDashboardCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupStewardPerformanceCalcSheet(); created++; } catch (e) { repaired++; }
  try { setupChecklistCalcSheet(); created++; } catch (e) { repaired++; }

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

  // Touch each hidden sheet to force recalc (6 hidden sheets)
  var hiddenSheetNames = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_CONTACT_CALC,
    SHEETS.DASHBOARD_CALC,
    SHEETS.STEWARD_PERFORMANCE_CALC
  ];

  hiddenSheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      // Force recalc by getting values
      sheet.getDataRange().getValues();
    }
  });

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

  sheet.getRange('A4').setValue('Step 1 Response');
  sheet.getRange('B4').setValue(DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);

  sheet.getRange('A5').setValue('Step 2 Appeal');
  sheet.getRange('B5').setValue(DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL);

  sheet.getRange('A6').setValue('Step 2 Response');
  sheet.getRange('B6').setValue(DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE);

  sheet.getRange('A7').setValue('Step 3 Appeal');
  sheet.getRange('B7').setValue(DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL);

  sheet.getRange('A8').setValue('Step 3 Response');
  sheet.getRange('B8').setValue(DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE);

  sheet.getRange('A9').setValue('Arbitration Demand');
  sheet.getRange('B9').setValue(DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND);

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
