/**
 * ============================================================================
 * Diagnostics.gs - System Diagnostics and Repair
 * ============================================================================
 *
 * This module handles all diagnostic and repair functions including:
 * - System diagnostics (DIAGNOSE_SETUP)
 * - Dashboard repair (REPAIR_DASHBOARD)
 * - Modal diagnostics
 * - Sheet verification
 *
 * REFACTORED: Split from 06_Maintenance.gs for better maintainability
 *
 * @fileoverview System diagnostics and repair functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// SYSTEM DIAGNOSTICS
// ============================================================================

/**
 * Runs a complete diagnostic check on the dashboard setup
 * Checks all required sheets, columns, formulas, and configurations
 * @returns {Object} Diagnostic results with status, checks, warnings, and errors
 */
function DIAGNOSE_SETUP() {
  var results = {
    timestamp: new Date().toISOString(),
    status: 'OK',
    checks: [],
    warnings: [],
    errors: []
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check 1: Required sheets exist
  results.checks.push('Checking required sheets...');
  var existingSheets = ss.getSheets().map(function(s) { return s.getName(); });

  for (var key in SHEET_NAMES) {
    var sheetName = SHEET_NAMES[key];
    if (existingSheets.indexOf(sheetName) === -1) {
      results.errors.push('Missing required sheet: ' + sheetName);
      results.status = 'ERROR';
    }
  }

  // Check 2: Hidden sheets exist
  results.checks.push('Checking hidden calculation sheets...');
  for (var hiddenKey in HIDDEN_SHEETS) {
    var hiddenName = HIDDEN_SHEETS[hiddenKey];
    if (existingSheets.indexOf(hiddenName) === -1) {
      results.warnings.push('Missing hidden sheet: ' + hiddenName + ' (will be auto-created on repair)');
      if (results.status === 'OK') results.status = 'WARNING';
    }
  }

  // Check 3: Member Directory structure
  results.checks.push('Verifying Member Directory structure...');
  var memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
  if (memberSheet) {
    var headers = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    var requiredHeaders = [
      'ID', 'First Name', 'Last Name', 'Employee ID', 'Department',
      'Job Title', 'Hire Date', 'Email', 'Phone', 'Status'
    ];

    requiredHeaders.forEach(function(header) {
      var found = headers.some(function(h) {
        return h.toString().toLowerCase().indexOf(header.toLowerCase()) !== -1;
      });
      if (!found) {
        results.warnings.push('Member Directory may be missing column: ' + header);
      }
    });
  }

  // Check 4: Grievance Tracker structure
  results.checks.push('Verifying Grievance Tracker structure...');
  var grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  if (grievanceSheet) {
    var gHeaders = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues()[0];
    var requiredGHeaders = [
      'Grievance ID', 'Member', 'Filing Date', 'Type', 'Current Step',
      'Status', 'Last Updated'
    ];

    requiredGHeaders.forEach(function(header) {
      var found = gHeaders.some(function(h) {
        return h.toString().toLowerCase().indexOf(header.toLowerCase()) !== -1;
      });
      if (!found) {
        results.warnings.push('Grievance Tracker may be missing column: ' + header);
      }
    });
  }

  // Check 5: Data integrity
  results.checks.push('Checking data integrity...');
  if (grievanceSheet) {
    var data = grievanceSheet.getDataRange().getValues();
    var orphanedGrievances = 0;
    var invalidDates = 0;

    for (var i = 1; i < data.length; i++) {
      var memberId = data[i][GRIEVANCE_COLUMNS.MEMBER_ID];
      var filingDate = data[i][GRIEVANCE_COLUMNS.FILING_DATE];

      if (!memberId && data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID]) {
        orphanedGrievances++;
      }

      if (filingDate && !(filingDate instanceof Date)) {
        invalidDates++;
      }
    }

    if (orphanedGrievances > 0) {
      results.warnings.push('Found ' + orphanedGrievances + ' grievances without member ID');
    }
    if (invalidDates > 0) {
      results.warnings.push('Found ' + invalidDates + ' grievances with invalid filing dates');
    }
  }

  // Check 6: Calendar permissions
  results.checks.push('Checking Calendar permissions...');
  try {
    CalendarApp.getAllCalendars();
    results.checks.push('Calendar access: OK');
  } catch (e) {
    results.warnings.push('Cannot access Calendar - sync features may not work');
  }

  // Check 7: Drive permissions
  results.checks.push('Checking Drive permissions...');
  try {
    DriveApp.getRootFolder();
    results.checks.push('Drive access: OK');
  } catch (e) {
    results.warnings.push('Cannot access Drive - folder features may not work');
  }

  // Finalize status
  if (results.errors.length > 0) {
    results.status = 'ERROR';
  } else if (results.warnings.length > 0) {
    results.status = 'WARNING';
  }

  results.summary = 'Completed ' + results.checks.length + ' checks: ' +
                    results.errors.length + ' errors, ' + results.warnings.length + ' warnings';

  // Log the diagnostic run
  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'DIAGNOSTICS',
    status: results.status,
    errors: results.errors.length,
    warnings: results.warnings.length,
    runBy: Session.getActiveUser().getEmail()
  });

  return results;
}

// ============================================================================
// REPAIR FUNCTIONS
// ============================================================================

/**
 * Repairs the dashboard by recreating missing sheets and fixing issues
 * @returns {void}
 */
function REPAIR_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    '🔧 Repair Dashboard',
    'This will:\n' +
    '• Recreate missing hidden sheets\n' +
    '• Reapply formulas and validations\n' +
    '• Fix broken references\n' +
    '• Install required triggers\n\n' +
    'Your data will NOT be deleted.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Repair cancelled.');
    return;
  }

  ss.toast('Starting repair...', '🔧 Repair', 5);

  try {
    // Recreate hidden sheets
    ss.toast('Recreating hidden sheets...', '🔧 Progress', 3);
    setupHiddenSheets(ss);

    // Reapply data validations
    ss.toast('Reapplying data validations...', '🔧 Progress', 3);
    setupDataValidations();

    // Install triggers
    ss.toast('Installing triggers...', '🔧 Progress', 3);
    installAutoSyncTrigger();

    // Sync data
    ss.toast('Syncing data...', '🔧 Progress', 3);
    syncAllData();

    ss.toast('Repair complete!', '✅ Success', 5);
    ui.alert('✅ Repair Complete',
      'Dashboard has been repaired successfully!\n\n' +
      '• Hidden sheets recreated\n' +
      '• Validations reapplied\n' +
      '• Triggers installed\n' +
      '• Data synced',
      ui.ButtonSet.OK);

    // Log the repair
    logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
      action: 'REPAIR_DASHBOARD',
      runBy: Session.getActiveUser().getEmail()
    });

  } catch (error) {
    Logger.log('Error in REPAIR_DASHBOARD: ' + error.message);
    ui.alert('❌ Error', 'Repair failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Removes deprecated tabs from the spreadsheet
 * @returns {void}
 */
function removeDeprecatedTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deprecatedSheets = [
    'Dashboard_OLD',
    'Member List_OLD',
    'BACKUP_',
    'TEST_'
  ];

  var removed = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    deprecatedSheets.forEach(function(prefix) {
      if (name.indexOf(prefix) === 0 || name.indexOf(prefix) !== -1) {
        try {
          ss.deleteSheet(sheet);
          removed.push(name);
        } catch (e) {
          Logger.log('Could not delete sheet: ' + name);
        }
      }
    });
  });

  if (removed.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Removed ' + removed.length + ' deprecated sheets:\n' + removed.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('No deprecated sheets found.');
  }
}

/**
 * Sets up sheet structure for a given sheet type
 * @param {Sheet} sheet - The sheet to set up
 * @param {string} sheetType - Type of sheet ('member', 'grievance', 'config')
 * @returns {void}
 */
function setupSheetStructure(sheet, sheetType) {
  // This is a placeholder that delegates to specific setup functions
  switch (sheetType) {
    case 'member':
      // Setup member directory structure
      break;
    case 'grievance':
      // Setup grievance log structure
      break;
    case 'config':
      // Setup config structure
      break;
    default:
      Logger.log('Unknown sheet type: ' + sheetType);
  }
}

/**
 * Shows the repair dialog
 * @returns {void}
 */
function showRepairDialog() {
  var diagnostics = DIAGNOSE_SETUP();

  var html = HtmlService.createHtmlOutput(getRepairDialogHtml_(diagnostics))
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '🔧 Repair Dashboard');
}

/**
 * Generates HTML for repair dialog
 * @param {Object} diagnostics - Diagnostic results
 * @returns {string} HTML content
 * @private
 */
function getRepairDialogHtml_(diagnostics) {
  var statusColor = diagnostics.status === 'OK' ? '#10b981' :
                    diagnostics.status === 'WARNING' ? '#f59e0b' : '#ef4444';

  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;margin:0}' +
    '.status{padding:10px;border-radius:8px;text-align:center;margin-bottom:20px;color:white;background:' + statusColor + '}' +
    '.section{margin-bottom:15px}' +
    '.section-title{font-weight:600;margin-bottom:8px}' +
    '.item{padding:4px 0;font-size:13px}' +
    '.error{color:#ef4444}' +
    '.warning{color:#f59e0b}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin-right:10px}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151}' +
    '</style></head><body>' +
    '<div class="status">' + diagnostics.status + '</div>' +
    '<p>' + diagnostics.summary + '</p>' +
    (diagnostics.errors.length > 0 ?
      '<div class="section"><div class="section-title error">Errors</div>' +
      diagnostics.errors.map(function(e) { return '<div class="item error">' + e + '</div>'; }).join('') +
      '</div>' : '') +
    (diagnostics.warnings.length > 0 ?
      '<div class="section"><div class="section-title warning">Warnings</div>' +
      diagnostics.warnings.map(function(w) { return '<div class="item warning">' + w + '</div>'; }).join('') +
      '</div>' : '') +
    '<div style="margin-top:20px;text-align:right">' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="runRepair()">Run Repair</button>' +
    '</div>' +
    '<script>' +
    'function runRepair(){google.script.run.withSuccessHandler(function(){google.script.host.close()}).REPAIR_DASHBOARD()}' +
    '</script></body></html>';
}

// ============================================================================
// MODAL DIAGNOSTICS
// ============================================================================

/**
 * Diagnoses why modals might not be loading data
 * @returns {Object} Diagnostic results
 */
function diagnoseModalIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = {
    status: 'OK',
    sheetChecks: [],
    dataChecks: [],
    modalTests: [],
    errors: [],
    warnings: []
  };

  var actualSheets = ss.getSheets().map(function(s) { return s.getName(); });

  // Check required sheets
  var requiredSheets = {
    'Member Directory': SHEETS.MEMBER_DIR,
    'Grievance Log': SHEETS.GRIEVANCE_LOG,
    'Member Satisfaction': SHEETS.SATISFACTION,
    'Config': SHEETS.CONFIG
  };

  for (var displayName in requiredSheets) {
    var expectedName = requiredSheets[displayName];
    var found = actualSheets.indexOf(expectedName) !== -1;
    results.sheetChecks.push({
      name: expectedName,
      expected: expectedName,
      found: found,
      status: found ? 'OK' : 'MISSING'
    });
    if (!found) {
      results.errors.push('Sheet "' + expectedName + '" not found.');
      results.status = 'ERROR';
    }
  }

  // Check data structure
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet) {
    var memberCols = memberSheet.getLastColumn();
    var memberRows = memberSheet.getLastRow();
    results.dataChecks.push({
      sheet: 'Member Directory',
      columns: memberCols,
      rows: memberRows,
      dataRows: Math.max(0, memberRows - 1),
      status: memberCols >= 20 ? 'OK' : 'WARNING'
    });
  }

  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet) {
    var grievCols = grievanceSheet.getLastColumn();
    var grievRows = grievanceSheet.getLastRow();
    results.dataChecks.push({
      sheet: 'Grievance Log',
      columns: grievCols,
      rows: grievRows,
      dataRows: Math.max(0, grievRows - 1),
      status: grievCols >= 20 ? 'OK' : 'WARNING'
    });
  }

  results.summary = results.status === 'OK'
    ? 'All modal checks passed.'
    : results.status === 'WARNING'
    ? 'Some warnings detected.'
    : 'Errors detected.';

  return results;
}

/**
 * Shows modal diagnostic results in a dialog
 * @returns {void}
 */
function showModalDiagnostics() {
  var results = diagnoseModalIssues();

  var html = '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px}' +
    '.card{background:#f9fafb;border-radius:8px;padding:15px;margin-bottom:15px}' +
    '.title{font-size:18px;font-weight:600;margin-bottom:15px}' +
    '.status-ok{color:#059669;background:#d1fae5;padding:4px 12px;border-radius:20px}' +
    '.status-error{color:#dc2626;background:#fee2e2;padding:4px 12px;border-radius:20px}' +
    'table{width:100%;border-collapse:collapse;font-size:13px}' +
    'th,td{padding:8px;text-align:left;border-bottom:1px solid #e5e7eb}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;background:#7c3aed;color:white}' +
    '</style></head><body>' +
    '<div class="title">Modal Diagnostics</div>' +
    '<span class="status-' + (results.status === 'OK' ? 'ok' : 'error') + '">' + results.status + '</span>' +
    '<p>' + results.summary + '</p>' +
    '<div class="card"><strong>Sheet Checks</strong><table>' +
    '<tr><th>Sheet</th><th>Status</th></tr>' +
    results.sheetChecks.map(function(c) {
      return '<tr><td>' + c.name + '</td><td>' + (c.found ? '✅' : '❌') + '</td></tr>';
    }).join('') +
    '</table></div>' +
    '<button class="btn" onclick="google.script.host.close()">Close</button>' +
    '</body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, '🔍 Modal Diagnostics');
}

/**
 * Shows the diagnostics dialog
 * @returns {void}
 */
function showDiagnosticsDialog() {
  var results = DIAGNOSE_SETUP();

  var html = HtmlService.createHtmlOutput(getDiagnosticsDialogHtml_(results))
    .setWidth(550)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '🩺 System Diagnostics');
}

/**
 * Generates HTML for diagnostics dialog
 * @param {Object} results - Diagnostic results
 * @returns {string} HTML content
 * @private
 */
function getDiagnosticsDialogHtml_(results) {
  var statusClass = results.status === 'OK' ? 'ok' :
                    results.status === 'WARNING' ? 'warning' : 'error';

  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;margin:0}' +
    '.status{display:inline-block;padding:8px 16px;border-radius:20px;font-weight:600;margin-bottom:15px}' +
    '.ok{background:#d1fae5;color:#065f46}' +
    '.warning{background:#fef3c7;color:#92400e}' +
    '.error{background:#fee2e2;color:#991b1b}' +
    '.section{margin:15px 0}' +
    '.section-title{font-weight:600;margin-bottom:8px}' +
    '.list{background:#f9fafb;padding:10px;border-radius:6px;max-height:120px;overflow-y:auto}' +
    '.item{padding:3px 0;font-size:13px}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151;margin-right:10px}' +
    '</style></head><body>' +
    '<div class="status ' + statusClass + '">' + results.status + '</div>' +
    '<p>' + results.summary + '</p>' +
    '<div class="section">' +
    '<div class="section-title">Checks Performed (' + results.checks.length + ')</div>' +
    '<div class="list">' +
    results.checks.map(function(c) { return '<div class="item">✓ ' + c + '</div>'; }).join('') +
    '</div></div>' +
    (results.errors.length > 0 ?
      '<div class="section"><div class="section-title error">Errors (' + results.errors.length + ')</div>' +
      '<div class="list">' + results.errors.map(function(e) { return '<div class="item" style="color:#991b1b">❌ ' + e + '</div>'; }).join('') + '</div></div>' : '') +
    (results.warnings.length > 0 ?
      '<div class="section"><div class="section-title warning">Warnings (' + results.warnings.length + ')</div>' +
      '<div class="list">' + results.warnings.map(function(w) { return '<div class="item" style="color:#92400e">⚠ ' + w + '</div>'; }).join('') + '</div></div>' : '') +
    '<div style="margin-top:20px;text-align:right">' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '<button class="btn btn-primary" onclick="runRepair()">Run Repair</button>' +
    '</div>' +
    '<script>function runRepair(){google.script.run.REPAIR_DASHBOARD();google.script.host.close()}</script>' +
    '</body></html>';
}
