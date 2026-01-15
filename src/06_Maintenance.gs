/**
 * ============================================================================
 * Maintenance.gs - Admin & Diagnostic Tools
 * ============================================================================
 *
 * This module contains administrative and diagnostic functions including:
 * - System diagnostics (DIAGNOSE_SETUP)
 * - Dashboard repair (REPAIR_DASHBOARD)
 * - Hidden sheet verification
 * - Data quality fixes
 * - Audit logging
 *
 * IMPORTANT: These are "heavy" operations that should only be run by
 * administrators. Moving them to a separate file prevents regular stewards
 * from accidentally triggering high-intensity system scans.
 *
 * @fileoverview Administrative and maintenance utilities
 * @version 2.0.0
 * @requires Constants.gs
 */

// ============================================================================
// SYSTEM DIAGNOSTICS
// ============================================================================

/**
 * Runs a complete diagnostic check on the dashboard setup
 * Checks all required sheets, columns, formulas, and configurations
 * @return {Object} Diagnostic results
 */
function DIAGNOSE_SETUP() {
  const results = {
    timestamp: new Date().toISOString(),
    status: 'OK',
    checks: [],
    warnings: [],
    errors: []
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check 1: Required sheets exist
  results.checks.push('Checking required sheets...');
  const allSheetNames = getAllSheetNames();
  const existingSheets = ss.getSheets().map(s => s.getName());

  for (const sheetName of Object.values(SHEET_NAMES)) {
    if (!existingSheets.includes(sheetName)) {
      results.errors.push(`Missing required sheet: ${sheetName}`);
      results.status = 'ERROR';
    }
  }

  // Check 2: Hidden sheets exist
  results.checks.push('Checking hidden calculation sheets...');
  for (const hiddenName of Object.values(HIDDEN_SHEETS)) {
    if (!existingSheets.includes(hiddenName)) {
      results.warnings.push(`Missing hidden sheet: ${hiddenName} (will be auto-created on repair)`);
      if (results.status === 'OK') results.status = 'WARNING';
    }
  }

  // Check 3: Member Directory structure
  results.checks.push('Verifying Member Directory structure...');
  const memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
  if (memberSheet) {
    const headers = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    const requiredHeaders = [
      'ID', 'First Name', 'Last Name', 'Employee ID', 'Department',
      'Job Title', 'Hire Date', 'Email', 'Phone', 'Status'
    ];

    for (const header of requiredHeaders) {
      if (!headers.some(h => h.toString().toLowerCase().includes(header.toLowerCase()))) {
        results.warnings.push(`Member Directory may be missing column: ${header}`);
      }
    }
  }

  // Check 4: Grievance Tracker structure
  results.checks.push('Verifying Grievance Tracker structure...');
  const grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  if (grievanceSheet) {
    const headers = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues()[0];
    const requiredHeaders = [
      'Grievance ID', 'Member', 'Filing Date', 'Type', 'Current Step',
      'Status', 'Last Updated'
    ];

    for (const header of requiredHeaders) {
      if (!headers.some(h => h.toString().toLowerCase().includes(header.toLowerCase()))) {
        results.warnings.push(`Grievance Tracker may be missing column: ${header}`);
      }
    }
  }

  // Check 5: Data integrity
  results.checks.push('Checking data integrity...');
  if (grievanceSheet) {
    const data = grievanceSheet.getDataRange().getValues();
    let orphanedGrievances = 0;
    let invalidDates = 0;

    for (let i = 1; i < data.length; i++) {
      const memberId = data[i][GRIEVANCE_COLUMNS.MEMBER_ID];
      const filingDate = data[i][GRIEVANCE_COLUMNS.FILING_DATE];

      // Check for missing member reference
      if (!memberId && data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID]) {
        orphanedGrievances++;
      }

      // Check for invalid dates
      if (filingDate && !(filingDate instanceof Date)) {
        invalidDates++;
      }
    }

    if (orphanedGrievances > 0) {
      results.warnings.push(`Found ${orphanedGrievances} grievances without member ID`);
    }
    if (invalidDates > 0) {
      results.warnings.push(`Found ${invalidDates} grievances with invalid filing dates`);
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

  results.summary = `Completed ${results.checks.length} checks: ` +
                    `${results.errors.length} errors, ${results.warnings.length} warnings`;

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

/**
 * Shows diagnostic results in a dialog
 */
function showDiagnosticsDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .diag-container { padding: 20px; }
        .status-badge { padding: 8px 16px; border-radius: 20px; font-weight: bold; margin-bottom: 20px; display: inline-block; }
        .status-ok { background: #d1fae5; color: #065f46; }
        .status-warning { background: #fef3c7; color: #92400e; }
        .status-error { background: #fee2e2; color: #991b1b; }
        .section { margin: 20px 0; }
        .section-title { font-weight: 600; margin-bottom: 8px; }
        .item-list { max-height: 150px; overflow-y: auto; background: #f9fafb; padding: 10px; border-radius: 8px; }
        .item { padding: 4px 0; font-size: 13px; }
        .item-error { color: #991b1b; }
        .item-warning { color: #92400e; }
        .loading { text-align: center; padding: 40px; }
      </style>
    </head>
    <body>
      <div class="diag-container">
        <div id="loading" class="loading">Running diagnostics...</div>
        <div id="results" style="display:none;">
          <div id="statusBadge"></div>
          <div id="summary"></div>

          <div class="section" id="errorsSection" style="display:none;">
            <div class="section-title">Errors</div>
            <div class="item-list" id="errorsList"></div>
          </div>

          <div class="section" id="warningsSection" style="display:none;">
            <div class="section-title">Warnings</div>
            <div class="item-list" id="warningsList"></div>
          </div>

          <div class="section">
            <div class="section-title">Checks Performed</div>
            <div class="item-list" id="checksList"></div>
          </div>

          <div style="margin-top: 20px; text-align: right;">
            <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
            <button class="btn btn-primary" onclick="runRepair()">Run Repair</button>
          </div>
        </div>
      </div>

      <script>
        function runDiagnostics() {
          google.script.run
            .withSuccessHandler(displayResults)
            .withFailureHandler(function(e) {
              document.getElementById('loading').textContent = 'Error: ' + e.message;
            })
            .DIAGNOSE_SETUP();
        }

        function displayResults(results) {
          document.getElementById('loading').style.display = 'none';
          document.getElementById('results').style.display = 'block';

          // Status badge
          const statusClass = results.status === 'OK' ? 'status-ok' :
                             results.status === 'WARNING' ? 'status-warning' : 'status-error';
          document.getElementById('statusBadge').innerHTML =
            '<span class="status-badge ' + statusClass + '">' + results.status + '</span>';

          // Summary
          document.getElementById('summary').textContent = results.summary;

          // Errors
          if (results.errors.length > 0) {
            document.getElementById('errorsSection').style.display = 'block';
            document.getElementById('errorsList').innerHTML =
              results.errors.map(e => '<div class="item item-error">• ' + e + '</div>').join('');
          }

          // Warnings
          if (results.warnings.length > 0) {
            document.getElementById('warningsSection').style.display = 'block';
            document.getElementById('warningsList').innerHTML =
              results.warnings.map(w => '<div class="item item-warning">• ' + w + '</div>').join('');
          }

          // Checks
          document.getElementById('checksList').innerHTML =
            results.checks.map(c => '<div class="item">✓ ' + c + '</div>').join('');
        }

        function runRepair() {
          if (confirm('Run dashboard repair? This will attempt to fix detected issues.')) {
            google.script.run
              .withSuccessHandler(function(r) {
                alert(r.success ? r.message : 'Error: ' + r.error);
                runDiagnostics(); // Re-run diagnostics
              })
              .REPAIR_DASHBOARD();
          }
        }

        runDiagnostics();
      </script>
    </body>
    </html>
  `).setWidth(550).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'System Diagnostics');
}

// ============================================================================
// DASHBOARD REPAIR
// ============================================================================

/**
 * Repairs common dashboard issues
 * - Creates missing sheets
 * - Repairs hidden calculation sheets
 * - Fixes data quality issues
 * @return {Object} Repair results
 */
function REPAIR_DASHBOARD() {
  const results = {
    timestamp: new Date().toISOString(),
    repairs: [],
    skipped: [],
    errors: []
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Repair 1: Create missing visible sheets
    results.repairs.push('Checking visible sheets...');
    for (const [key, sheetName] of Object.entries(SHEET_NAMES)) {
      if (!ss.getSheetByName(sheetName)) {
        const newSheet = ss.insertSheet(sheetName);
        setupSheetStructure(newSheet, key);
        results.repairs.push(`Created sheet: ${sheetName}`);
      }
    }

    // Repair 2: Setup/repair hidden sheets
    results.repairs.push('Repairing hidden calculation sheets...');
    const hiddenResult = setupAllHiddenSheets();
    results.repairs.push(`Hidden sheets: ${hiddenResult.created} created, ${hiddenResult.repaired} repaired`);

    // Repair 3: Fix data quality issues
    results.repairs.push('Checking data quality...');
    const dataResult = fixDataQualityIssues();
    results.repairs.push(`Data fixes: ${dataResult.fixed} issues corrected`);

    // Repair 4: Verify and repair formulas
    results.repairs.push('Verifying formulas...');
    const formulaResult = repairAllHiddenSheets();
    results.repairs.push(`Formulas: ${formulaResult.repaired} sheets updated`);

    // Log the repair
    logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
      action: 'REPAIR',
      repairs: results.repairs.length,
      errors: results.errors.length,
      performedBy: Session.getActiveUser().getEmail()
    });

    results.success = true;
    results.message = `Repair completed: ${results.repairs.length} actions taken`;

  } catch (error) {
    console.error('Error during repair:', error);
    results.errors.push(error.message);
    results.success = false;
    results.error = error.message;
  }

  return results;
}

/**
 * Sets up basic structure for a sheet based on type
 * @param {Sheet} sheet - The sheet to set up
 * @param {string} sheetType - The type key from SHEET_NAMES
 */
function setupSheetStructure(sheet, sheetType) {
  switch (sheetType) {
    case 'MEMBER_DIRECTORY':
      sheet.appendRow([
        'ID', 'First Name', 'Last Name', 'Employee ID', 'Department',
        'Job Title', 'Hire Date', 'Seniority Date', 'Email', 'Phone',
        'Status', 'Union Status', 'Notes', 'Last Updated'
      ]);
      sheet.setFrozenRows(1);
      break;

    case 'GRIEVANCE_TRACKER':
      sheet.appendRow([
        'Grievance ID', 'Member ID', 'Member Name', 'Filing Date', 'Grievance Type',
        'Article Violated', 'Description', 'Current Step', 'Step 1 Date', 'Step 1 Due',
        'Step 1 Status', 'Step 2 Date', 'Step 2 Due', 'Step 2 Status', 'Step 3 Date',
        'Step 3 Due', 'Step 3 Status', 'Arbitration Date', 'Resolution', 'Outcome',
        'Drive Folder', 'Notes', 'Status', 'Last Updated'
      ]);
      sheet.setFrozenRows(1);
      break;

    case 'AUDIT_LOG':
      sheet.appendRow([
        'Timestamp', 'Event Type', 'User', 'Details', 'IP Address'
      ]);
      sheet.setFrozenRows(1);
      break;

    default:
      // Basic setup for other sheets
      sheet.setFrozenRows(1);
  }
}

/**
 * Shows the repair dialog
 */
function showRepairDialog() {
  const result = showConfirmation(
    'This will attempt to repair the dashboard by:\n\n' +
    '• Creating any missing sheets\n' +
    '• Repairing hidden calculation sheets\n' +
    '• Fixing data quality issues\n' +
    '• Rebuilding formulas\n\n' +
    'This is safe to run and won\'t delete any data. Continue?',
    'Repair Dashboard'
  );

  if (result) {
    const repairResult = REPAIR_DASHBOARD();
    if (repairResult.success) {
      showToast(repairResult.message, 'Repair Complete');
    } else {
      showAlert('Repair failed: ' + repairResult.error, 'Error');
    }
  }
}

// ============================================================================
// Note: verifyHiddenSheets() and fixDataQualityIssues() are defined in
// HiddenSheets.gs and DataIntegrity.gs which contain comprehensive implementations.
// ============================================================================

// ============================================================================
// AUDIT LOGGING
// ============================================================================

/**
 * Logs an audit event to the Audit Log sheet
 * @param {string} eventType - The type of event from AUDIT_EVENTS
 * @param {Object} details - Event details object
 */
function logAuditEvent(eventType, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

    // Create audit sheet if it doesn't exist
    if (!auditSheet) {
      auditSheet = ss.insertSheet(SHEET_NAMES.AUDIT_LOG);
      auditSheet.appendRow(['Timestamp', 'Event Type', 'User', 'Details', 'Session ID']);
      auditSheet.setFrozenRows(1);
    }

    // Build log entry
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    const detailsJson = JSON.stringify(details);
    const sessionId = Session.getTemporaryActiveUserKey() || '';

    auditSheet.appendRow([timestamp, eventType, user, detailsJson, sessionId]);

    // Trim old entries if sheet gets too large (keep last 10,000)
    const rowCount = auditSheet.getLastRow();
    if (rowCount > 10000) {
      auditSheet.deleteRows(2, rowCount - 10000);
    }

  } catch (error) {
    console.error('Error logging audit event:', error);
    // Don't throw - audit logging shouldn't break main functionality
  }
}

/**
 * Gets recent audit log entries
 * @param {number} count - Number of entries to retrieve
 * @return {Array} Array of audit log entries
 */
function getRecentAuditLogs(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

  if (!auditSheet) return [];

  const lastRow = auditSheet.getLastRow();
  const startRow = Math.max(2, lastRow - count + 1);
  const numRows = lastRow - startRow + 1;

  if (numRows <= 0) return [];

  const data = auditSheet.getRange(startRow, 1, numRows, 5).getValues();

  return data.map(row => ({
    timestamp: row[0],
    eventType: row[1],
    user: row[2],
    details: row[3],
    sessionId: row[4]
  })).reverse();
}

// ============================================================================
// NUCLEAR OPTIONS (ADMIN ONLY)
// ============================================================================

/**
 * WARNING: Completely resets all hidden sheets
 * This is a destructive operation that should only be used as a last resort
 * @return {Object} Reset results
 */
function NUCLEAR_RESET_HIDDEN_SHEETS() {
  // Require double confirmation
  const ui = SpreadsheetApp.getUi();
  const response1 = ui.alert(
    '⚠️ NUCLEAR RESET',
    'This will DELETE and REBUILD all hidden calculation sheets. ' +
    'This may temporarily break formulas. Are you absolutely sure?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    return { success: false, message: 'Cancelled by user' };
  }

  const response2 = ui.alert(
    '⚠️ FINAL WARNING',
    'Type "CONFIRM" to proceed with the nuclear reset.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response2 !== ui.Button.OK) {
    return { success: false, message: 'Cancelled by user' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete all hidden sheets
  for (const sheetName of Object.values(HIDDEN_SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  // Rebuild all hidden sheets
  const result = setupAllHiddenSheets();

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'NUCLEAR_RESET',
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    message: `Nuclear reset complete. Created ${result.created} sheets.`
  };
}

/**
 * WARNING: Wipes all grievance data
 * For testing/development use only
 */
function NUCLEAR_WIPE_GRIEVANCES() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ DANGER - DATA WIPE',
    'This will PERMANENTLY DELETE all grievance records. ' +
    'This action CANNOT be undone. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { success: false, message: 'Cancelled' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);

  if (!sheet) {
    return { success: false, error: 'Grievance Tracker not found' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'NUCLEAR_WIPE_GRIEVANCES',
    rowsDeleted: lastRow - 1,
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    message: `Deleted ${lastRow - 1} grievance records`
  };
}

// ============================================================================
// SETTINGS MANAGEMENT
// ============================================================================

/**
 * Shows the settings dialog
 */
function showSettingsDialog() {
  const currentSettings = getSettings();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .settings-container { padding: 20px; }
        .setting-group { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #eee; }
        .setting-title { font-weight: 600; margin-bottom: 8px; }
        .setting-desc { font-size: 12px; color: #666; margin-bottom: 10px; }
        .danger-zone { background: #fff5f5; padding: 15px; border-radius: 8px; border: 1px solid #feb2b2; margin-top: 20px; }
        .danger-title { color: #c53030; font-weight: 600; margin-bottom: 10px; }
      </style>
    </head>
    <body>
      <div class="settings-container">
        <div class="setting-group">
          <div class="setting-title">Calendar Integration</div>
          <div class="setting-desc">Automatically sync grievance deadlines to Google Calendar</div>
          <label>
            <input type="checkbox" id="autoSyncCalendar"
                   ${currentSettings.autoSyncCalendar ? 'checked' : ''}>
            Enable auto-sync on grievance changes
          </label>
        </div>

        <div class="setting-group">
          <div class="setting-title">Email Notifications</div>
          <div class="setting-desc">Send email reminders for upcoming deadlines</div>
          <label>
            <input type="checkbox" id="emailReminders"
                   ${currentSettings.emailReminders ? 'checked' : ''}>
            Enable deadline reminders
          </label>
          <div style="margin-top: 8px;">
            <label>Days before deadline:
              <select id="reminderDays">
                <option value="3" ${currentSettings.reminderDays === 3 ? 'selected' : ''}>3 days</option>
                <option value="5" ${currentSettings.reminderDays === 5 ? 'selected' : ''}>5 days</option>
                <option value="7" ${currentSettings.reminderDays === 7 ? 'selected' : ''}>7 days</option>
              </select>
            </label>
          </div>
        </div>

        <div class="setting-group">
          <div class="setting-title">Drive Integration</div>
          <div class="setting-desc">Automatically create Drive folders for new grievances</div>
          <label>
            <input type="checkbox" id="autoCreateFolders"
                   ${currentSettings.autoCreateFolders ? 'checked' : ''}>
            Auto-create folders
          </label>
        </div>

        <div class="danger-zone">
          <div class="danger-title">⚠️ Danger Zone</div>
          <p style="font-size: 13px; margin-bottom: 10px;">
            Administrative functions that should be used with caution
          </p>
          <button class="btn btn-secondary" onclick="runDiagnostics()">
            Run Diagnostics
          </button>
          <button class="btn btn-secondary" onclick="repairDashboard()">
            Repair Dashboard
          </button>
          <button class="btn btn-danger" onclick="nuclearReset()">
            Nuclear Reset
          </button>
        </div>

        <div style="margin-top: 20px; text-align: right;">
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
          <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
        </div>
      </div>

      <script>
        function saveSettings() {
          const settings = {
            autoSyncCalendar: document.getElementById('autoSyncCalendar').checked,
            emailReminders: document.getElementById('emailReminders').checked,
            reminderDays: parseInt(document.getElementById('reminderDays').value),
            autoCreateFolders: document.getElementById('autoCreateFolders').checked
          };

          google.script.run
            .withSuccessHandler(function() {
              alert('Settings saved!');
              google.script.host.close();
            })
            .saveSettings(settings);
        }

        function runDiagnostics() {
          google.script.run.showDiagnosticsDialog();
          google.script.host.close();
        }

        function repairDashboard() {
          google.script.run.showRepairDialog();
          google.script.host.close();
        }

        function nuclearReset() {
          if (confirm('This is an extreme action. Are you sure?')) {
            google.script.run.NUCLEAR_RESET_HIDDEN_SHEETS();
            google.script.host.close();
          }
        }
      </script>
    </body>
    </html>
  `).setWidth(500).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Settings');
}

/**
 * Gets current settings from document properties
 * @return {Object} Current settings
 */
function getSettings() {
  const props = PropertiesService.getDocumentProperties();
  return {
    autoSyncCalendar: props.getProperty('autoSyncCalendar') === 'true',
    emailReminders: props.getProperty('emailReminders') === 'true',
    reminderDays: parseInt(props.getProperty('reminderDays') || '7'),
    autoCreateFolders: props.getProperty('autoCreateFolders') === 'true'
  };
}

/**
 * Saves settings to document properties
 * @param {Object} settings - Settings to save
 */
function saveSettings(settings) {
  const props = PropertiesService.getDocumentProperties();

  props.setProperty('autoSyncCalendar', String(settings.autoSyncCalendar));
  props.setProperty('emailReminders', String(settings.emailReminders));
  props.setProperty('reminderDays', String(settings.reminderDays));
  props.setProperty('autoCreateFolders', String(settings.autoCreateFolders));

  logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
    settings: settings,
    changedBy: Session.getActiveUser().getEmail()
  });
}
/**
 * 509 Dashboard - Data Integrity and Performance Enhancements
 *
 * This module contains improvements for:
 * - Batch operations for optimized data writing
 * - Comprehensive error handling with retry logic
 * - Confirmation dialogs for destructive actions
 * - Dynamic validation ranges
 * - Duplicate ID validation
 * - Ghost validation for orphaned grievances
 * - Steward load balancing metrics
 * - Self-healing Config tool
 * - Enhanced audit logging
 * - Auto-archive for closed grievances
 *
 * @version 2.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// BATCH OPERATIONS UTILITIES
// ============================================================================

/**
 * Batch write utility - writes multiple values to a sheet in a single operation
 * Significantly faster than individual setValue() calls
 * @param {Sheet} sheet - Target sheet
 * @param {Array<Object>} updates - Array of {row, col, value} objects
 */
function batchSetValues(sheet, updates) {
  if (!updates || updates.length === 0) return;

  // Group updates by row for efficient writing
  var rowGroups = {};
  updates.forEach(function(update) {
    if (!rowGroups[update.row]) {
      rowGroups[update.row] = [];
    }
    rowGroups[update.row].push(update);
  });

  // Get current sheet data for merging
  var lastRow = Math.max.apply(null, updates.map(function(u) { return u.row; }));
  var lastCol = Math.max.apply(null, updates.map(function(u) { return u.col; }));

  // Read current data
  var currentData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Apply updates to the array
  updates.forEach(function(update) {
    currentData[update.row - 1][update.col - 1] = update.value;
  });

  // Write back in single operation
  sheet.getRange(1, 1, lastRow, lastCol).setValues(currentData);
}

/**
 * Batch write utility for a specific row - updates multiple columns in one operation
 * @param {Sheet} sheet - Target sheet
 * @param {number} row - Row number (1-indexed)
 * @param {Object} columnValues - Object mapping column numbers to values {col: value}
 */
function batchSetRowValues(sheet, row, columnValues) {
  var cols = Object.keys(columnValues);
  if (cols.length === 0) return;

  // Find range bounds
  var minCol = Math.min.apply(null, cols.map(Number));
  var maxCol = Math.max.apply(null, cols.map(Number));
  var numCols = maxCol - minCol + 1;

  // Read current row data
  var rowData = sheet.getRange(row, minCol, 1, numCols).getValues()[0];

  // Apply updates
  cols.forEach(function(col) {
    var colNum = parseInt(col);
    rowData[colNum - minCol] = columnValues[col];
  });

  // Write back in single operation
  sheet.getRange(row, minCol, 1, numCols).setValues([rowData]);
}

/**
 * Batch append rows - adds multiple rows in a single operation
 * Much faster than multiple appendRow() calls
 * @param {Sheet} sheet - Target sheet
 * @param {Array<Array>} rows - 2D array of row data
 */
function batchAppendRows(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var lastRow = sheet.getLastRow();
  var numCols = rows[0].length;

  // Write all rows in a single operation
  sheet.getRange(lastRow + 1, 1, rows.length, numCols).setValues(rows);
}

// ============================================================================
// ERROR HANDLING WITH RETRY LOGIC
// ============================================================================

/**
 * Execute a function with retry logic for transient failures
 * @param {Function} fn - Function to execute
 * @param {Object} options - Options: {maxRetries, baseDelay, onError}
 * @returns {*} Result of the function
 */
function executeWithRetry(fn, options) {
  options = options || {};
  var maxRetries = options.maxRetries || 3;
  var baseDelay = options.baseDelay || 1000; // 1 second
  var onError = options.onError || function() {};

  var lastError;

  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      // Log the error
      Logger.log('Attempt ' + (attempt + 1) + ' failed: ' + error.message);
      onError(error, attempt);

      // Check if we should retry
      if (attempt < maxRetries) {
        // Exponential backoff: 1s, 2s, 4s, 8s...
        var delay = baseDelay * Math.pow(2, attempt);
        Utilities.sleep(delay);
      }
    }
  }

  // All retries exhausted
  throw new Error('Operation failed after ' + (maxRetries + 1) + ' attempts: ' + lastError.message);
}

/**
 * Safe sheet operation wrapper with error handling
 * @param {Function} operation - Sheet operation to perform
 * @param {string} operationName - Name for logging
 * @returns {Object} {success: boolean, result: *, error: string}
 */
function safeSheetOperation(operation, operationName) {
  try {
    var result = executeWithRetry(operation, {
      maxRetries: 2,
      baseDelay: 500,
      onError: function(error, attempt) {
        Logger.log(operationName + ' - Retry ' + (attempt + 1) + ': ' + error.message);
      }
    });
    return { success: true, result: result, error: null };
  } catch (error) {
    Logger.log(operationName + ' FAILED: ' + error.message);
    return { success: false, result: null, error: error.message };
  }
}

// ============================================================================
// CONFIRMATION DIALOGS FOR DESTRUCTIVE ACTIONS
// ============================================================================

/**
 * Safe version of getOrCreateSheet that confirms before deleting existing data
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} name - Sheet name
 * @param {boolean} forceDelete - Skip confirmation if true
 * @returns {Sheet|null} The created sheet, or null if user cancelled
 */
function getOrCreateSheetSafe(ss, name, forceDelete) {
  var sheet = ss.getSheetByName(name);

  if (sheet) {
    // Check if sheet has data
    var hasData = sheet.getLastRow() > 1 || sheet.getLastColumn() > 1;

    if (hasData && !forceDelete) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        '⚠️ Confirm Data Deletion',
        'The sheet "' + name + '" already exists and contains data.\n\n' +
        'Deleting this sheet will permanently remove all data in it.\n\n' +
        'Are you sure you want to continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        Logger.log('User cancelled deletion of sheet: ' + name);
        return null;
      }
    }

    ss.deleteSheet(sheet);
    Logger.log('Deleted existing sheet: ' + name);
  }

  return ss.insertSheet(name);
}

/**
 * Confirm before performing any destructive operation
 * @param {string} actionName - Description of the action
 * @param {string} warningMessage - Detailed warning message
 * @returns {boolean} True if user confirmed, false otherwise
 */
function confirmDestructiveAction(actionName, warningMessage) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ ' + actionName,
    warningMessage + '\n\nThis action cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

// ============================================================================
// DYNAMIC VALIDATION RANGES
// ============================================================================

/**
 * Set dropdown validation using dynamic row count instead of fixed 100
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setDropdownValidationDynamic(targetSheet, targetCol, configSheet, sourceCol) {
  // Get the actual last row with data in the config column
  var configData = configSheet.getRange(3, sourceCol, configSheet.getLastRow() - 2, 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  // Use at least 10 rows, or actual count + buffer
  var rowCount = Math.max(10, actualRows + 10);

  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Set multi-select validation using dynamic row count
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setMultiSelectValidationDynamic(targetSheet, targetCol, configSheet, sourceCol) {
  // Get the actual last row with data in the config column
  var configData = configSheet.getRange(3, sourceCol, configSheet.getLastRow() - 2, 1).getValues();
  var actualRows = 0;

  for (var i = 0; i < configData.length; i++) {
    if (configData[i][0] !== '' && configData[i][0] !== null) {
      actualRows = i + 1;
    }
  }

  // Use at least 10 rows, or actual count + buffer
  var rowCount = Math.max(10, actualRows + 10);

  var sourceRange = configSheet.getRange(3, sourceCol, rowCount, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true) // Allow comma-separated values
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

// ============================================================================
// DUPLICATE MEMBER ID VALIDATION
// ============================================================================

/**
 * Check if a Member ID already exists in the Member Directory
 * @param {string} memberId - Member ID to check
 * @returns {Object} {exists: boolean, row: number|null}
 */
function checkDuplicateMemberId(memberId) {
  if (!memberId) return { exists: false, row: null };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return { exists: false, row: null };

  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();

  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0] === memberId) {
      return { exists: true, row: i + 2 }; // +2 for header and 0-index
    }
  }

  return { exists: false, row: null };
}

/**
 * Real-time duplicate Member ID validator for onEdit trigger
 * Call this from onEdit when Member ID column is modified
 * @param {Event} e - Edit event
 */
function validateMemberIdOnEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Only check Member Directory, Member ID column
  if (sheet.getName() !== SHEETS.MEMBER_DIR) return;
  if (range.getColumn() !== MEMBER_COLS.MEMBER_ID) return;

  var newValue = e.value;
  if (!newValue) return;

  var currentRow = range.getRow();
  var result = checkDuplicateMemberId(newValue);

  if (result.exists && result.row !== currentRow) {
    // Duplicate found in a different row
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Duplicate Member ID',
      'The Member ID "' + newValue + '" already exists in row ' + result.row + '.\n\n' +
      'Please use a unique Member ID.',
      ui.ButtonSet.OK
    );

    // Optionally highlight the cell
    range.setBackground('#FFCDD2'); // Light red

    // Revert to old value if available
    if (e.oldValue) {
      range.setValue(e.oldValue);
    }
  }
}

// ============================================================================
// GHOST VALIDATION FOR ORPHANED GRIEVANCES
// ============================================================================

/**
 * Find grievances with Member IDs that don't exist in Member Directory
 * @returns {Array<Object>} Array of orphaned grievance info
 */
function findOrphanedGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    return [];
  }

  // Build set of valid member IDs
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
  var validMemberIds = {};
  memberIds.forEach(function(row) {
    if (row[0]) validMemberIds[row[0]] = true;
  });

  // Check grievance member IDs
  var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 4).getValues();
  var orphaned = [];

  grievanceData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];

    if (memberId && !validMemberIds[memberId]) {
      orphaned.push({
        row: index + 2,
        grievanceId: grievanceId,
        memberId: memberId,
        memberName: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')
      });
    }
  });

  return orphaned;
}

/**
 * Highlight orphaned grievances in the Grievance Log
 * Adds red background to rows with invalid Member IDs
 */
function highlightOrphanedGrievances() {
  var orphaned = findOrphanedGrievances();

  if (orphaned.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Data Integrity Check',
      'All grievances have valid Member IDs. No orphaned records found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Highlight orphaned rows
  orphaned.forEach(function(item) {
    grievanceSheet.getRange(item.row, 1, 1, grievanceSheet.getLastColumn())
      .setBackground('#FFCDD2'); // Light red
  });

  // Report findings
  var message = 'Found ' + orphaned.length + ' orphaned grievance(s):\n\n';
  orphaned.slice(0, 10).forEach(function(item) {
    message += '• Row ' + item.row + ': ' + item.grievanceId + ' (Member: ' + item.memberId + ')\n';
  });

  if (orphaned.length > 10) {
    message += '\n...and ' + (orphaned.length - 10) + ' more.';
  }

  message += '\n\nThese rows have been highlighted in red.';

  SpreadsheetApp.getUi().alert('⚠️ Orphaned Grievances Found', message, SpreadsheetApp.getUi().ButtonSet.OK);

  // Log to audit
  logIntegrityEvent('GHOST_VALIDATION', 'Found ' + orphaned.length + ' orphaned grievances');
}

/**
 * Run ghost validation automatically (for scheduled trigger)
 * Sends email to admin if orphans are found
 */
function runScheduledGhostValidation() {
  var orphaned = findOrphanedGrievances();

  if (orphaned.length > 0) {
    // Get admin emails from Config
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    var adminEmail = configSheet.getRange(3, CONFIG_COLS.ADMIN_EMAILS, 1, 1).getValue();

    if (adminEmail) {
      var subject = '⚠️ 509 Dashboard: Orphaned Grievances Detected';
      var body = 'The scheduled data integrity check found ' + orphaned.length + ' grievances with invalid Member IDs.\n\n';

      orphaned.slice(0, 20).forEach(function(item) {
        body += '• Row ' + item.row + ': ' + item.grievanceId + ' - Member ID: ' + item.memberId + ' (' + item.memberName + ')\n';
      });

      if (orphaned.length > 20) {
        body += '\n...and ' + (orphaned.length - 20) + ' more.\n';
      }

      body += '\nPlease review and correct these records in the Grievance Log.';
      body += '\n\n--\n509 Dashboard Automated Alert';

      try {
        MailApp.sendEmail(adminEmail, subject, body);
        Logger.log('Orphan alert sent to: ' + adminEmail);
      } catch (e) {
        Logger.log('Failed to send orphan alert: ' + e.message);
      }
    }
  }
}

// ============================================================================
// STEWARD LOAD BALANCING METRICS
// ============================================================================

/**
 * Calculate steward workload metrics
 * @returns {Array<Object>} Array of steward workload data
 */
function calculateStewardWorkload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return [];

  var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();

  // Count active grievances per steward
  var stewardCounts = {};
  var activeStatuses = ['Open', 'Pending Info', 'In Arbitration', 'Appealed'];

  grievanceData.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var steward = row[GRIEVANCE_COLS.STEWARD - 1];

    if (steward && activeStatuses.indexOf(status) !== -1) {
      if (!stewardCounts[steward]) {
        stewardCounts[steward] = {
          name: steward,
          activeCount: 0,
          urgentCount: 0,
          totalAssigned: 0
        };
      }
      stewardCounts[steward].activeCount++;
      stewardCounts[steward].totalAssigned++;

      // Check if urgent (days to deadline <= 7)
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
      if (typeof daysToDeadline === 'number' && daysToDeadline <= 7) {
        stewardCounts[steward].urgentCount++;
      }
    }
  });

  // Convert to array and calculate load scores
  var stewards = Object.keys(stewardCounts).map(function(name) {
    var data = stewardCounts[name];
    // Load score: active cases + (urgent cases * 2)
    data.loadScore = data.activeCount + (data.urgentCount * 2);
    return data;
  });

  // Sort by load score descending
  stewards.sort(function(a, b) { return b.loadScore - a.loadScore; });

  return stewards;
}

/**
 * Show steward workload dashboard dialog
 */
function showStewardWorkloadDashboard() {
  var stewards = calculateStewardWorkload();

  if (stewards.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Steward Workload',
      'No stewards have active grievances assigned.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Calculate statistics
  var totalActive = stewards.reduce(function(sum, s) { return sum + s.activeCount; }, 0);
  var avgLoad = totalActive / stewards.length;

  var html = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; margin-bottom: 10px; }' +
    '.stats { background: #F3F4F6; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    '.stat { display: inline-block; margin-right: 30px; }' +
    '.stat-value { font-size: 24px; font-weight: bold; color: #7C3AED; }' +
    '.stat-label { font-size: 12px; color: #6B7280; }' +
    'table { width: 100%; border-collapse: collapse; }' +
    'th { background: #7C3AED; color: white; padding: 10px; text-align: left; }' +
    'td { padding: 8px; border-bottom: 1px solid #E5E7EB; }' +
    '.high { background: #FEE2E2; }' +
    '.medium { background: #FEF3C7; }' +
    '.low { background: #D1FAE5; }' +
    '.load-bar { height: 10px; background: #E5E7EB; border-radius: 5px; }' +
    '.load-fill { height: 100%; border-radius: 5px; }' +
    '</style>';

  html += '<h2>Steward Workload Dashboard</h2>';

  html += '<div class="stats">' +
    '<div class="stat"><div class="stat-value">' + stewards.length + '</div><div class="stat-label">Active Stewards</div></div>' +
    '<div class="stat"><div class="stat-value">' + totalActive + '</div><div class="stat-label">Total Active Cases</div></div>' +
    '<div class="stat"><div class="stat-value">' + avgLoad.toFixed(1) + '</div><div class="stat-label">Avg Cases/Steward</div></div>' +
    '</div>';

  html += '<table><tr><th>Steward</th><th>Active</th><th>Urgent</th><th>Load Score</th><th>Status</th></tr>';

  var maxLoad = stewards[0] ? stewards[0].loadScore : 1;

  stewards.forEach(function(s) {
    var statusClass = s.loadScore > avgLoad * 1.5 ? 'high' : (s.loadScore > avgLoad ? 'medium' : 'low');
    var statusText = s.loadScore > avgLoad * 1.5 ? 'Overloaded' : (s.loadScore > avgLoad ? 'Busy' : 'Normal');
    var loadPct = (s.loadScore / maxLoad * 100).toFixed(0);
    var loadColor = statusClass === 'high' ? '#DC2626' : (statusClass === 'medium' ? '#F59E0B' : '#059669');

    html += '<tr class="' + statusClass + '">' +
      '<td><strong>' + s.name + '</strong></td>' +
      '<td>' + s.activeCount + '</td>' +
      '<td>' + s.urgentCount + '</td>' +
      '<td>' +
        '<div class="load-bar"><div class="load-fill" style="width: ' + loadPct + '%; background: ' + loadColor + ';"></div></div>' +
        '<small>' + s.loadScore + '</small>' +
      '</td>' +
      '<td>' + statusText + '</td>' +
      '</tr>';
  });

  html += '</table>';
  html += '<p style="margin-top: 15px; color: #6B7280; font-size: 12px;">' +
    'Load Score = Active Cases + (Urgent Cases x 2). Urgent = deadline within 7 days.</p>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Steward Workload Analysis');
}

/**
 * Get steward with lowest workload for new case assignment
 * @returns {string|null} Name of steward with lowest load, or null if none
 */
function getStewardWithLowestWorkload() {
  var stewards = calculateStewardWorkload();

  // Also get all stewards from Config (some may have 0 cases)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var stewardList = configSheet.getRange(3, CONFIG_COLS.STEWARDS, 50, 1).getValues()
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) { return row[0]; });

  // Find stewards with 0 active cases
  var activeStewardNames = stewards.map(function(s) { return s.name; });
  var availableStewards = stewardList.filter(function(name) {
    return activeStewardNames.indexOf(name) === -1;
  });

  if (availableStewards.length > 0) {
    return availableStewards[0]; // Return first available steward with 0 cases
  }

  // Otherwise return steward with lowest load score
  if (stewards.length > 0) {
    return stewards[stewards.length - 1].name;
  }

  return null;
}

// ============================================================================
// SELF-HEALING CONFIG VALIDATION TOOL
// ============================================================================

/**
 * Scan all data sheets for values not in Config dropdowns
 * @returns {Object} Report of missing config values by column
 */
function findMissingConfigValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet || !memberSheet || !grievanceSheet) {
    return { error: 'Required sheets not found' };
  }

  var report = { missingValues: [], autoFixable: [] };

  // Define field mappings to check
  var fieldsToCheck = [
    { sheet: memberSheet, col: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES, name: 'Job Title' },
    { sheet: memberSheet, col: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS, name: 'Work Location' },
    { sheet: memberSheet, col: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS, name: 'Unit' },
    { sheet: memberSheet, col: MEMBER_COLS.SUPERVISOR, configCol: CONFIG_COLS.SUPERVISORS, name: 'Supervisor' },
    { sheet: memberSheet, col: MEMBER_COLS.MANAGER, configCol: CONFIG_COLS.MANAGERS, name: 'Manager' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.STATUS, configCol: CONFIG_COLS.GRIEVANCE_STATUS, name: 'Grievance Status' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.CURRENT_STEP, configCol: CONFIG_COLS.GRIEVANCE_STEP, name: 'Grievance Step' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.ISSUE_CATEGORY, configCol: CONFIG_COLS.ISSUE_CATEGORY, name: 'Issue Category' }
  ];

  fieldsToCheck.forEach(function(field) {
    // Get valid config values
    var configValues = configSheet.getRange(3, field.configCol, 100, 1).getValues()
      .filter(function(row) { return row[0] !== ''; })
      .map(function(row) { return row[0]; });

    var validSet = {};
    configValues.forEach(function(v) { validSet[v] = true; });

    // Get data values
    var lastRow = field.sheet.getLastRow();
    if (lastRow < 2) return;

    var dataValues = field.sheet.getRange(2, field.col, lastRow - 1, 1).getValues();

    // Find values not in config
    var missing = {};
    dataValues.forEach(function(row, index) {
      var value = row[0];
      if (value && !validSet[value] && !missing[value]) {
        missing[value] = {
          field: field.name,
          value: value,
          configCol: field.configCol,
          exampleRow: index + 2
        };
      }
    });

    Object.values(missing).forEach(function(item) {
      report.missingValues.push(item);
      report.autoFixable.push(item);
    });
  });

  return report;
}

/**
 * Show Config health check dialog with auto-fix option
 */
function showConfigHealthCheck() {
  var report = findMissingConfigValues();

  if (report.error) {
    SpreadsheetApp.getUi().alert('Error', report.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (report.missingValues.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Config Health Check',
      'All dropdown values in your data sheets exist in the Config sheet.\n\nNo issues found!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var html = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #DC2626; }' +
    '.warning { background: #FEF3C7; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    'table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }' +
    'th { background: #7C3AED; color: white; padding: 10px; text-align: left; }' +
    'td { padding: 8px; border-bottom: 1px solid #E5E7EB; }' +
    '.btn { background: #059669; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; }' +
    '.btn:hover { background: #047857; }' +
    '</style>';

  html += '<h2>⚠️ Missing Config Values Found</h2>';
  html += '<div class="warning">' +
    '<strong>' + report.missingValues.length + ' value(s)</strong> in your data sheets are not in the Config dropdowns. ' +
    'This can cause validation errors and data inconsistency.' +
    '</div>';

  html += '<table><tr><th>Field</th><th>Missing Value</th><th>Example Row</th></tr>';

  report.missingValues.slice(0, 20).forEach(function(item) {
    html += '<tr><td>' + item.field + '</td><td><strong>' + item.value + '</strong></td><td>' + item.exampleRow + '</td></tr>';
  });

  if (report.missingValues.length > 20) {
    html += '<tr><td colspan="3">...and ' + (report.missingValues.length - 20) + ' more</td></tr>';
  }

  html += '</table>';

  html += '<button class="btn" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).autoFixMissingConfigValues()">Auto-Add Missing Values to Config</button>';
  html += '<p style="margin-top: 10px; color: #6B7280; font-size: 12px;">This will add the missing values to the appropriate Config columns.</p>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Config Health Check');
}

/**
 * Auto-add missing values to Config sheet
 */
function autoFixMissingConfigValues() {
  var report = findMissingConfigValues();

  if (report.error || report.autoFixable.length === 0) {
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  // Group by config column
  var byColumn = {};
  report.autoFixable.forEach(function(item) {
    if (!byColumn[item.configCol]) {
      byColumn[item.configCol] = [];
    }
    byColumn[item.configCol].push(item.value);
  });

  // Add values to each column
  Object.keys(byColumn).forEach(function(colStr) {
    var col = parseInt(colStr);
    var values = byColumn[col];

    // Find next empty row in this column
    var colData = configSheet.getRange(3, col, 100, 1).getValues();
    var nextRow = 3;
    for (var i = 0; i < colData.length; i++) {
      if (colData[i][0] !== '') {
        nextRow = i + 4;
      }
    }

    // Add values
    values.forEach(function(value, index) {
      configSheet.getRange(nextRow + index, col).setValue(value);
    });
  });

  // Log the fix
  logIntegrityEvent('CONFIG_AUTO_FIX', 'Added ' + report.autoFixable.length + ' missing values to Config');

  SpreadsheetApp.getUi().alert(
    '✅ Config Updated',
    'Successfully added ' + report.autoFixable.length + ' missing values to the Config sheet.\n\n' +
    'Please run "Setup Data Validations" from the Settings menu to refresh dropdowns.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// ENHANCED AUDIT LOGGING
// ============================================================================

/**
 * Log a data integrity event to the audit log with timestamp and user info
 * Note: This is separate from the detailed edit audit in Code.gs logAuditEvent()
 * @param {string} eventType - Type of event (e.g., 'STATUS_CHANGE', 'DATA_EDIT')
 * @param {string} details - Description of what happened
 * @param {Object} additionalInfo - Optional additional information
 */
function logIntegrityEvent(eventType, details, additionalInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  // Create audit log if it doesn't exist
  if (!auditSheet) {
    auditSheet = ss.insertSheet(SHEETS.AUDIT_LOG);
    auditSheet.hideSheet();

    // Set up headers
    var headers = ['Timestamp', 'User Email', 'Event Type', 'Details', 'Additional Info', 'Spreadsheet ID'];
    auditSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold')
      .setBackground('#7C3AED')
      .setFontColor('#FFFFFF');
  }

  // Get user info
  var userEmail = Session.getActiveUser().getEmail() || 'Unknown';
  var timestamp = new Date();
  var spreadsheetId = ss.getId();

  // Append log entry
  var logEntry = [
    timestamp,
    userEmail,
    eventType,
    details,
    additionalInfo ? JSON.stringify(additionalInfo) : '',
    spreadsheetId
  ];

  auditSheet.appendRow(logEntry);

  // Keep audit log to reasonable size (last 5000 entries)
  var lastRow = auditSheet.getLastRow();
  if (lastRow > 5000) {
    auditSheet.deleteRows(2, lastRow - 5000);
  }
}

/**
 * Log grievance status change (call from onEdit trigger)
 * @param {string} grievanceId - The grievance ID
 * @param {string} oldStatus - Previous status
 * @param {string} newStatus - New status
 */
function logGrievanceStatusChange(grievanceId, oldStatus, newStatus) {
  logIntegrityEvent('STATUS_CHANGE',
    'Grievance ' + grievanceId + ' status changed from "' + oldStatus + '" to "' + newStatus + '"',
    { grievanceId: grievanceId, oldStatus: oldStatus, newStatus: newStatus }
  );
}

/**
 * Log steward assignment change
 * @param {string} grievanceId - The grievance ID
 * @param {string} oldSteward - Previous steward
 * @param {string} newSteward - New steward
 */
function logStewardAssignmentChange(grievanceId, oldSteward, newSteward) {
  logIntegrityEvent('STEWARD_CHANGE',
    'Grievance ' + grievanceId + ' reassigned from "' + (oldSteward || 'None') + '" to "' + newSteward + '"',
    { grievanceId: grievanceId, oldSteward: oldSteward, newSteward: newSteward }
  );
}

/**
 * Show audit log viewer dialog
 */
function showAuditLogViewer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!auditSheet || auditSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert(
      'Audit Log',
      'No audit log entries found. The audit log records important changes like status updates and steward assignments.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Get recent entries
  var lastRow = auditSheet.getLastRow();
  var startRow = Math.max(2, lastRow - 49); // Last 50 entries
  var data = auditSheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();

  var html = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; }' +
    'table { width: 100%; border-collapse: collapse; font-size: 12px; }' +
    'th { background: #7C3AED; color: white; padding: 8px; text-align: left; position: sticky; top: 0; }' +
    'td { padding: 6px; border-bottom: 1px solid #E5E7EB; }' +
    'tr:hover { background: #F3F4F6; }' +
    '.status { background: #DBEAFE; padding: 2px 6px; border-radius: 4px; }' +
    '.steward { background: #D1FAE5; padding: 2px 6px; border-radius: 4px; }' +
    '.config { background: #FEF3C7; padding: 2px 6px; border-radius: 4px; }' +
    '</style>';

  html += '<h2>Recent Audit Log Entries</h2>';
  html += '<p style="color: #6B7280;">Showing last ' + data.length + ' entries</p>';

  html += '<table><tr><th>Timestamp</th><th>User</th><th>Event</th><th>Details</th></tr>';

  // Reverse to show most recent first
  data.reverse().forEach(function(row) {
    var timestamp = row[0] instanceof Date ? row[0].toLocaleString() : row[0];
    var user = row[1] ? row[1].split('@')[0] : 'Unknown';
    var eventType = row[2] || '';
    var details = row[3] || '';

    var eventClass = eventType.indexOf('STATUS') !== -1 ? 'status' :
                     (eventType.indexOf('STEWARD') !== -1 ? 'steward' :
                     (eventType.indexOf('CONFIG') !== -1 ? 'config' : ''));

    html += '<tr>' +
      '<td>' + timestamp + '</td>' +
      '<td>' + user + '</td>' +
      '<td><span class="' + eventClass + '">' + eventType + '</span></td>' +
      '<td>' + details.substring(0, 100) + (details.length > 100 ? '...' : '') + '</td>' +
      '</tr>';
  });

  html += '</table>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Audit Log Viewer');
}

// ============================================================================
// AUTO-ARCHIVE FOR CLOSED GRIEVANCES
// ============================================================================

/**
 * Archive closed grievances older than specified days
 * @param {number} daysOld - Archive grievances closed more than this many days ago (default 90)
 */
function archiveClosedGrievances(daysOld) {
  daysOld = daysOld || 90;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    Logger.log('Grievance Log not found');
    return { archived: 0 };
  }

  // Get or create archive sheet
  var archiveSheetName = '_Archive_Grievances';
  var archiveSheet = ss.getSheetByName(archiveSheetName);

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveSheetName);
    archiveSheet.hideSheet();

    // Copy headers
    var headers = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues();
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
      .setFontWeight('bold')
      .setBackground('#6B7280')
      .setFontColor('#FFFFFF');
  }

  // Find rows to archive
  var closedStatuses = ['Closed', 'Won', 'Denied', 'Settled', 'Withdrawn'];
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysOld);

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return { archived: 0 };

  var data = grievanceSheet.getRange(2, 1, lastRow - 1, grievanceSheet.getLastColumn()).getValues();

  var rowsToArchive = [];
  var rowIndicesToDelete = [];

  data.forEach(function(row, index) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];

    if (closedStatuses.indexOf(status) !== -1 && dateClosed instanceof Date && dateClosed < cutoffDate) {
      rowsToArchive.push(row);
      rowIndicesToDelete.push(index + 2); // +2 for header and 0-index
    }
  });

  if (rowsToArchive.length === 0) {
    return { archived: 0 };
  }

  // Append to archive sheet
  var archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
    .setValues(rowsToArchive);

  // Delete from main sheet (in reverse order to maintain row indices)
  rowIndicesToDelete.reverse().forEach(function(rowIndex) {
    grievanceSheet.deleteRow(rowIndex);
  });

  // Log the archive operation
  logIntegrityEvent('AUTO_ARCHIVE',
    'Archived ' + rowsToArchive.length + ' closed grievances older than ' + daysOld + ' days',
    { count: rowsToArchive.length, daysOld: daysOld }
  );

  return { archived: rowsToArchive.length };
}

/**
 * Show archive dialog with options
 */
function showArchiveDialog() {
  var html = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; }' +
    '.info { background: #EFF6FF; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    'label { display: block; margin-bottom: 5px; font-weight: bold; }' +
    'input[type="number"] { width: 100px; padding: 8px; border: 1px solid #D1D5DB; border-radius: 4px; margin-bottom: 15px; }' +
    '.btn { background: #7C3AED; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }' +
    '.btn:hover { background: #6D28D9; }' +
    '.btn-secondary { background: #6B7280; margin-left: 10px; }' +
    '</style>';

  html += '<h2>Archive Closed Grievances</h2>';

  html += '<div class="info">' +
    '<strong>What this does:</strong><br>' +
    'Moves closed grievances (Won, Denied, Settled, Withdrawn, Closed) to a hidden archive sheet. ' +
    'This keeps your active Grievance Log fast and focused on current cases.' +
    '</div>';

  html += '<label>Archive grievances closed more than:</label>';
  html += '<input type="number" id="daysOld" value="90" min="30" max="365"> days ago';

  html += '<br><br>';
  html += '<button class="btn" onclick="runArchive()">Archive Now</button>';
  html += '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>';

  html += '<script>' +
    'function runArchive() {' +
    '  var days = document.getElementById("daysOld").value;' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    alert("Archived " + result.archived + " grievances.");' +
    '    google.script.host.close();' +
    '  }).archiveClosedGrievances(parseInt(days));' +
    '}' +
    '</script>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Archive Grievances');
}

/**
 * Restore grievances from archive
 * @param {Array<string>} grievanceIds - Array of grievance IDs to restore
 */
function restoreFromArchive(grievanceIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var archiveSheet = ss.getSheetByName('_Archive_Grievances');

  if (!archiveSheet || !grievanceSheet) {
    return { restored: 0, error: 'Required sheets not found' };
  }

  var archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archiveSheet.getLastColumn()).getValues();

  var rowsToRestore = [];
  var rowIndicesToDelete = [];

  archiveData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (grievanceIds.indexOf(grievanceId) !== -1) {
      rowsToRestore.push(row);
      rowIndicesToDelete.push(index + 2);
    }
  });

  if (rowsToRestore.length === 0) {
    return { restored: 0 };
  }

  // Add back to main sheet
  var lastRow = grievanceSheet.getLastRow();
  grievanceSheet.getRange(lastRow + 1, 1, rowsToRestore.length, rowsToRestore[0].length)
    .setValues(rowsToRestore);

  // Remove from archive
  rowIndicesToDelete.reverse().forEach(function(rowIndex) {
    archiveSheet.deleteRow(rowIndex);
  });

  logIntegrityEvent('ARCHIVE_RESTORE',
    'Restored ' + rowsToRestore.length + ' grievances from archive',
    { grievanceIds: grievanceIds }
  );

  return { restored: rowsToRestore.length };
}

// ============================================================================
// VISUAL DEADLINE HEATMAP WITH SPARKLINES
// ============================================================================

/**
 * Apply deadline heatmap conditional formatting to Grievance Log
 * Colors cells based on urgency level
 */
function applyDeadlineHeatmap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return;

  // Days to Deadline column
  var deadlineRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);

  // Clear existing conditional formatting for this range
  var rules = grievanceSheet.getConditionalFormatRules();
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(function(r) {
      return r.getColumn() === GRIEVANCE_COLS.DAYS_TO_DEADLINE;
    });
  });

  // Create new rules for heatmap
  // Overdue (negative or "Overdue" text)
  var overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#DC2626')  // Bright red
    .setFontColor('#FFFFFF')
    .setBold(true)
    .setRanges([deadlineRange])
    .build();

  // Critical (1-3 days)
  var criticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 3)
    .setBackground('#F87171')  // Light red
    .setFontColor('#7F1D1D')
    .setBold(true)
    .setRanges([deadlineRange])
    .build();

  // Warning (4-7 days)
  var warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 7)
    .setBackground('#FBBF24')  // Yellow/orange
    .setFontColor('#78350F')
    .setRanges([deadlineRange])
    .build();

  // Caution (8-14 days)
  var cautionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(8, 14)
    .setBackground('#FEF3C7')  // Light yellow
    .setFontColor('#92400E')
    .setRanges([deadlineRange])
    .build();

  // Safe (15+ days)
  var safeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(14)
    .setBackground('#D1FAE5')  // Light green
    .setFontColor('#065F46')
    .setRanges([deadlineRange])
    .build();

  newRules.push(overdueRule, criticalRule, warningRule, cautionRule, safeRule);
  grievanceSheet.setConditionalFormatRules(newRules);

  // Also apply to status column for closed cases
  var statusRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1);

  var wonRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Won')
    .setBackground('#059669')  // Green
    .setFontColor('#FFFFFF')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  var deniedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Denied')
    .setBackground('#DC2626')  // Red
    .setFontColor('#FFFFFF')
    .setRanges([statusRange])
    .build();

  var settledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Settled')
    .setBackground('#7C3AED')  // Purple
    .setFontColor('#FFFFFF')
    .setRanges([statusRange])
    .build();

  var existingRules = grievanceSheet.getConditionalFormatRules();
  existingRules.push(wonRule, deniedRule, settledRule);
  grievanceSheet.setConditionalFormatRules(existingRules);

  SpreadsheetApp.getUi().alert(
    '✅ Heatmap Applied',
    'Deadline heatmap has been applied to the Grievance Log:\n\n' +
    '🔴 Red: Overdue or 1-3 days\n' +
    '🟡 Yellow: 4-7 days\n' +
    '🟢 Green: 15+ days',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// MOBILE STEWARD PORTAL VIEW
// ============================================================================

/**
 * Create or update the mobile steward portal sheet
 * Shows only essential columns for mobile access
 */
function createMobileStewardPortal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var portalSheetName = '📱 Steward Portal';

  // Get or create portal sheet
  var portalSheet = ss.getSheetByName(portalSheetName);
  if (portalSheet) {
    portalSheet.clear();
  } else {
    portalSheet = ss.insertSheet(portalSheetName);
  }

  // Get grievance data
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet) {
    portalSheet.getRange('A1').setValue('Error: Grievance Log not found');
    return;
  }

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    portalSheet.getRange('A1').setValue('No grievances found');
    return;
  }

  // Title
  portalSheet.getRange('A1').setValue('📱 STEWARD PORTAL')
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor('#7C3AED');
  portalSheet.getRange('A1:E1').merge();

  portalSheet.getRange('A2').setValue('Quick access to your active cases. Updated: ' + new Date().toLocaleString())
    .setFontStyle('italic')
    .setFontColor('#6B7280');
  portalSheet.getRange('A2:E2').merge();

  // Section: Urgent Cases (deadline <= 7 days)
  portalSheet.getRange('A4').setValue('🚨 URGENT CASES')
    .setFontWeight('bold')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF');
  portalSheet.getRange('A4:E4').merge();

  // Headers for mobile view
  var headers = ['ID', 'Member', 'Status', 'Deadline', 'Steward'];
  portalSheet.getRange('A5:E5').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#F3F4F6');

  // Get grievance data and filter
  var grievanceData = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var activeStatuses = ['Open', 'Pending Info', 'In Arbitration', 'Appealed'];

  var urgentCases = [];
  var normalCases = [];

  grievanceData.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    if (activeStatuses.indexOf(status) === -1) return;

    var mobileRow = [
      row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
      (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '').charAt(0) + '.',
      row[GRIEVANCE_COLS.STATUS - 1],
      row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
      row[GRIEVANCE_COLS.STEWARD - 1]
    ];

    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    if (typeof daysToDeadline === 'number' && daysToDeadline <= 7) {
      urgentCases.push(mobileRow);
    } else {
      normalCases.push(mobileRow);
    }
  });

  // Sort urgent cases by deadline ascending
  urgentCases.sort(function(a, b) { return (a[3] || 999) - (b[3] || 999); });
  normalCases.sort(function(a, b) { return (a[3] || 999) - (b[3] || 999); });

  var currentRow = 6;

  // Write urgent cases
  if (urgentCases.length > 0) {
    portalSheet.getRange(currentRow, 1, urgentCases.length, 5).setValues(urgentCases)
      .setBackground('#FEE2E2');
    currentRow += urgentCases.length + 1;
  } else {
    portalSheet.getRange(currentRow, 1).setValue('No urgent cases!')
      .setFontColor('#059669');
    currentRow += 2;
  }

  // Section: All Active Cases
  portalSheet.getRange(currentRow, 1).setValue('📋 ALL ACTIVE CASES')
    .setFontWeight('bold')
    .setBackground('#7C3AED')
    .setFontColor('#FFFFFF');
  portalSheet.getRange(currentRow, 1, 1, 5).merge();
  currentRow++;

  // Headers
  portalSheet.getRange(currentRow, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#F3F4F6');
  currentRow++;

  // Write normal cases
  if (normalCases.length > 0) {
    portalSheet.getRange(currentRow, 1, normalCases.length, 5).setValues(normalCases);
  } else {
    portalSheet.getRange(currentRow, 1).setValue('No other active cases');
  }

  // Format columns for mobile
  portalSheet.setColumnWidth(1, 90);  // ID
  portalSheet.setColumnWidth(2, 120); // Member
  portalSheet.setColumnWidth(3, 80);  // Status
  portalSheet.setColumnWidth(4, 70);  // Deadline
  portalSheet.setColumnWidth(5, 100); // Steward

  // Freeze header
  portalSheet.setFrozenRows(3);

  // Move to front
  ss.setActiveSheet(portalSheet);
  ss.moveActiveSheet(2);

  SpreadsheetApp.getUi().alert(
    '📱 Steward Portal Created',
    'The mobile-friendly Steward Portal has been created/updated.\n\n' +
    'This view shows:\n' +
    '• Urgent cases (7 days or less)\n' +
    '• All active cases\n' +
    '• Narrow columns optimized for mobile',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// ENHANCED onEdit TRIGGER WITH AUDIT LOGGING
// ============================================================================

/**
 * Enhanced onEdit handler that logs important changes
 * Add this to your onEdit trigger
 * @param {Event} e - Edit event
 */
function onEditWithAuditLogging(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var oldValue = e.oldValue;
  var newValue = e.value;

  // Only process single cell edits
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  var sheetName = sheet.getName();
  var col = range.getColumn();
  var row = range.getRow();

  // Skip header row
  if (row === 1) return;

  // Grievance Log changes
  if (sheetName === SHEETS.GRIEVANCE_LOG) {
    var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

    // Status change
    if (col === GRIEVANCE_COLS.STATUS && oldValue !== newValue) {
      logGrievanceStatusChange(grievanceId, oldValue, newValue);
    }

    // Steward assignment change
    if (col === GRIEVANCE_COLS.STEWARD && oldValue !== newValue) {
      logStewardAssignmentChange(grievanceId, oldValue, newValue);
    }
  }

  // Member Directory - duplicate ID check
  if (sheetName === SHEETS.MEMBER_DIR && col === MEMBER_COLS.MEMBER_ID) {
    validateMemberIdOnEdit(e);
  }
}

// ============================================================================
// MENU ADDITIONS
// ============================================================================

/**
 * Add Data Integrity menu items to the existing menu structure
 * Call this from onOpen or add to existing menu creation
 */
function addDataIntegrityMenuItems() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('🛡️ Data Integrity')
    .addItem('🔍 Find Orphaned Grievances', 'highlightOrphanedGrievances')
    .addItem('🗂️ Clean Up Orphaned Folders', 'showOrphanedFolderCleanupDialog')
    .addItem('⚙️ Config Health Check', 'showConfigHealthCheck')
    .addSeparator()
    .addItem('📊 Steward Workload Dashboard', 'showStewardWorkloadDashboard')
    .addItem('📱 Create/Update Steward Portal', 'createMobileStewardPortal')
    .addSeparator()
    .addItem('🎨 Apply Deadline Heatmap', 'applyDeadlineHeatmap')
    .addItem('📦 Archive Closed Grievances', 'showArchiveDialog')
    .addSeparator()
    .addItem('📜 View Audit Log', 'showAuditLogViewer')
    .addToUi();
}
/**
 * ============================================================================
 * PERFORMANCE CACHING & UNDO/REDO SYSTEM
 * ============================================================================
 * Data caching for performance + action history
 */

// ==================== CACHE CONFIGURATION ====================

var CACHE_CONFIG = {
  MEMORY_TTL: 300,
  PROPS_TTL: 3600,
  ENABLE_LOGGING: false
};

var CACHE_KEYS = {
  ALL_GRIEVANCES: 'cache_grievances',
  ALL_MEMBERS: 'cache_members',
  ALL_STEWARDS: 'cache_stewards',
  DASHBOARD_METRICS: 'cache_metrics'
};

// ==================== CACHING FUNCTIONS ====================

function getCachedData(key, loader, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;
  try {
    var memCache = CacheService.getScriptCache();
    var cached = memCache.get(key);
    if (cached) { if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT] ' + key); return JSON.parse(cached); }
    var propsCache = PropertiesService.getScriptProperties();
    var propsCached = propsCache.getProperty(key);
    if (propsCached) {
      var obj = JSON.parse(propsCached);
      if (obj.timestamp && (Date.now() - obj.timestamp) < (ttl * 1000)) {
        memCache.put(key, JSON.stringify(obj.data), ttl);
        return obj.data;
      }
    }
    if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE MISS] ' + key);
    var data = loader();
    setCachedData(key, data, ttl);
    return data;
  } catch (e) { Logger.log('Cache error: ' + e.message); return loader(); }
}

function setCachedData(key, data, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;
  try {
    var str = JSON.stringify(data);
    var memCache = CacheService.getScriptCache();
    memCache.put(key, str, Math.min(ttl, 21600));
    if (str.length < 400000) {
      var propsCache = PropertiesService.getScriptProperties();
      propsCache.setProperty(key, JSON.stringify({ data: data, timestamp: Date.now() }));
    }
  } catch (e) { Logger.log('Set cache error: ' + e.message); }
}

function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
    PropertiesService.getScriptProperties().deleteProperty(key);
    Logger.log('Cache invalidated: ' + key);
  } catch (e) { Logger.log('Invalidate error: ' + e.message); }
}

function invalidateAllCaches() {
  try {
    var keys = Object.keys(CACHE_KEYS).map(function(k) { return CACHE_KEYS[k]; });
    CacheService.getScriptCache().removeAll(keys);
    var props = PropertiesService.getScriptProperties();
    keys.forEach(function(k) { props.deleteProperty(k); });
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ All caches cleared', 'Cache', 3);
  } catch (e) { Logger.log('Clear all error: ' + e.message); }
}

function warmUpCaches() {
  SpreadsheetApp.getActiveSpreadsheet().toast('🔥 Warming caches...', 'Cache', -1);
  try {
    getCachedGrievances();
    getCachedMembers();
    getCachedStewards();
    getCachedDashboardMetrics();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Caches warmed', 'Cache', 3);
  } catch (e) { Logger.log('Warmup error: ' + e.message); }
}

function getCachedGrievances() {
  return getCachedData(CACHE_KEYS.ALL_GRIEVANCES, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 300);
}

function getCachedMembers() {
  return getCachedData(CACHE_KEYS.ALL_MEMBERS, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 600);
}

function getCachedStewards() {
  return getCachedData(CACHE_KEYS.ALL_STEWARDS, function() {
    var members = getCachedMembers();
    return members.filter(function(row) { return row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes'; });
  }, 600);
}

function getCachedDashboardMetrics() {
  return getCachedData(CACHE_KEYS.DASHBOARD_METRICS, function() {
    var grievances = getCachedGrievances();
    var metrics = { total: grievances.length, open: 0, closed: 0, overdue: 0, byStatus: {}, byIssueType: {}, bySteward: {} };
    var today = new Date(); today.setHours(0, 0, 0, 0);
    grievances.forEach(function(row) {
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var issue = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
      var steward = row[GRIEVANCE_COLS.STEWARD - 1];
      var deadline = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysTo = deadline ? Math.floor((new Date(deadline) - today) / (1000 * 60 * 60 * 24)) : null;
      if (status === 'Open') metrics.open++;
      if (status === 'Closed' || status === 'Resolved') metrics.closed++;
      if (daysTo !== null && daysTo < 0) metrics.overdue++;
      metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;
      if (issue) metrics.byIssueType[issue] = (metrics.byIssueType[issue] || 0) + 1;
      if (steward) metrics.bySteward[steward] = (metrics.bySteward[steward] || 0) + 1;
    });
    return metrics;
  }, 180);
}

function showCacheStatusDashboard() {
  var memCache = CacheService.getScriptCache();
  var propsCache = PropertiesService.getScriptProperties();
  var rows = Object.keys(CACHE_KEYS).map(function(name) {
    var key = CACHE_KEYS[name];
    var inMem = memCache.get(key) !== null;
    var inProps = propsCache.getProperty(key) !== null;
    var age = 'N/A';
    if (inProps) {
      try { var obj = JSON.parse(propsCache.getProperty(key)); if (obj.timestamp) age = Math.floor((Date.now() - obj.timestamp) / 1000) + 's'; } catch (e) {}
    }
    return '<tr><td>' + name + '</td><td>' + (inMem ? '✅' : '❌') + '</td><td>' + (inProps ? '✅' : '❌') + '</td><td>' + age + '</td></tr>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}table{width:100%;border-collapse:collapse;margin:20px 0}th{background:#1a73e8;color:white;padding:12px;text-align:left}td{padding:10px;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:5px}button.danger{background:#dc3545}</style></head><body><div class="container"><h2>🗄️ Cache Status</h2><table><tr><th>Cache</th><th>Memory</th><th>Props</th><th>Age</th></tr>' + rows + '</table><button onclick="google.script.run.withSuccessHandler(function(){location.reload()}).warmUpCaches()">🔥 Warm Up</button><button class="danger" onclick="google.script.run.withSuccessHandler(function(){location.reload()}).invalidateAllCaches()">🗑️ Clear All</button></div></body></html>'
  ).setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '🗄️ Cache Status');
}

// ==================== UNDO/REDO SYSTEM ====================

var UNDO_CONFIG = { MAX_HISTORY: 50, STORAGE_KEY: 'undoRedoHistory' };

function getUndoHistory() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty(UNDO_CONFIG.STORAGE_KEY);
  if (json) return JSON.parse(json);
  return { actions: [], currentIndex: 0 };
}

function saveUndoHistory(history) {
  var props = PropertiesService.getScriptProperties();
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }
  props.setProperty(UNDO_CONFIG.STORAGE_KEY, JSON.stringify(history));
}

function recordAction(type, description, beforeState, afterState) {
  var history = getUndoHistory();
  if (history.currentIndex < history.actions.length) history.actions = history.actions.slice(0, history.currentIndex);
  history.actions.push({ type: type, description: description, timestamp: new Date().toISOString(), beforeState: beforeState, afterState: afterState });
  history.currentIndex = history.actions.length;
  saveUndoHistory(history);
}

function recordCellEdit(row, col, oldValue, newValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var colName = sheet.getRange(1, col).getValue();
  recordAction('EDIT_CELL', 'Edited ' + colName + ' in row ' + row, { row: row, col: col, value: oldValue, sheet: sheet.getName() }, { row: row, col: col, value: newValue, sheet: sheet.getName() });
}

function recordRowAddition(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  recordAction('ADD_ROW', 'Added row ' + row, null, { row: row, data: rowData, sheet: sheet.getName() });
}

function recordRowDeletion(row, rowData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  recordAction('DELETE_ROW', 'Deleted row ' + row, { row: row, data: rowData, sheet: sheet.getName() }, null);
}

function undoLastAction() {
  var history = getUndoHistory();
  if (history.currentIndex === 0) throw new Error('Nothing to undo');
  var action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);
  history.currentIndex--;
  saveUndoHistory(history);
  SpreadsheetApp.getActiveSpreadsheet().toast('↩️ Undone: ' + action.description, 'Undo', 3);
}

function redoLastAction() {
  var history = getUndoHistory();
  if (history.currentIndex >= history.actions.length) throw new Error('Nothing to redo');
  var action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);
  history.currentIndex++;
  saveUndoHistory(history);
  SpreadsheetApp.getActiveSpreadsheet().toast('↪️ Redone: ' + action.description, 'Redo', 3);
}

function undoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex > targetIndex) undoLastAction();
}

function redoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) redoLastAction();
}

function applyState(state, actionType) {
  if (!state) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(state.sheet);
  if (!sheet) throw new Error('Sheet ' + state.sheet + ' not found');
  switch (actionType) {
    case 'EDIT_CELL': sheet.getRange(state.row, state.col).setValue(state.value); break;
    case 'ADD_ROW': if (state.row) sheet.deleteRow(state.row); break;
    case 'DELETE_ROW': if (state.row && state.data) { sheet.insertRowAfter(state.row - 1); sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]); } break;
    case 'BATCH_UPDATE': if (state.changes) state.changes.forEach(function(c) { sheet.getRange(c.row, c.col).setValue(c.oldValue); }); break;
  }
}

function clearUndoHistory() {
  PropertiesService.getScriptProperties().deleteProperty(UNDO_CONFIG.STORAGE_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ History cleared', 'Undo/Redo', 3);
}

function exportUndoHistoryToSheet() {
  var history = getUndoHistory();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Undo_History_Export');
  if (sheet) sheet.clear();
  else sheet = ss.insertSheet('Undo_History_Export');
  var headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
  if (history.actions.length > 0) {
    var rows = history.actions.map(function(a, i) { return [i + 1, a.type, a.description, new Date(a.timestamp).toLocaleString(), i < history.currentIndex ? 'Applied' : 'Undone']; });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  for (var c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);
  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

function showUndoRedoPanel() {
  var history = getUndoHistory();
  var rows = history.actions.length === 0 ? '<tr><td colspan="5" style="text-align:center;padding:40px;color:#999">No actions recorded</td></tr>' :
    history.actions.slice().reverse().map(function(a, i) {
      var idx = history.actions.length - i;
      var time = new Date(a.timestamp).toLocaleString();
      var canUndo = i < history.actions.length - history.currentIndex;
      return '<tr><td>' + idx + '</td><td><span class="badge ' + a.type.toLowerCase() + '">' + a.type + '</span></td><td>' + a.description + '</td><td style="font-size:12px;color:#666">' + time + '</td><td>' + (canUndo ? '<button onclick="undo(' + (idx - 1) + ')">↩️</button>' : '') + '</td></tr>';
    }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin:20px 0}.stat{background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;border-left:4px solid #1a73e8}.num{font-size:32px;font-weight:bold;color:#1a73e8}table{width:100%;border-collapse:collapse;margin:20px 0}th{background:#1a73e8;color:white;padding:12px;text-align:left}td{padding:12px;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;margin:2px}.badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:11px;font-weight:bold}.edit_cell{background:#e3f2fd;color:#1976d2}.add_row{background:#e8f5e9;color:#388e3c}.delete_row{background:#ffebee;color:#d32f2f}.quick{margin:20px 0;padding:20px;background:#f8f9fa;border-radius:8px}</style></head><body><div class="container"><h2>↩️ Undo/Redo History</h2><div class="stats"><div class="stat"><div class="num">' + history.actions.length + '</div><div>Total Actions</div></div><div class="stat"><div class="num">' + history.currentIndex + '</div><div>Current Position</div></div><div class="stat"><div class="num">' + (history.actions.length - history.currentIndex) + '</div><div>Available</div></div></div><div class="quick"><button onclick="performUndo()">↩️ Undo Last</button><button onclick="performRedo()">↪️ Redo</button><button onclick="clear()" style="background:#d32f2f">🗑️ Clear</button><button onclick="exp()" style="background:#00796b">📥 Export</button></div><table><tr><th>#</th><th>Type</th><th>Description</th><th>Time</th><th></th></tr>' + rows + '</table></div><script>function performUndo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("❌ "+e.message)}).undoLastAction()}function performRedo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("❌ "+e.message)}).redoLastAction()}function undo(i){google.script.run.withSuccessHandler(function(){location.reload()}).undoToIndex(i)}function clear(){if(confirm("Clear all history?")){google.script.run.withSuccessHandler(function(){location.reload()}).clearUndoHistory()}}function exp(){google.script.run.withSuccessHandler(function(url){alert("✅ Exported!");window.open(url,"_blank")}).exportUndoHistoryToSheet()}</script></body></html>'
  ).setWidth(800).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '↩️ Undo/Redo History');
}

function createGrievanceSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  return { timestamp: new Date().toISOString(), data: sheet.getDataRange().getValues(), lastRow: sheet.getLastRow(), lastColumn: sheet.getLastColumn() };
}

function restoreFromSnapshot(snapshot) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  if (snapshot.data.length > 1) sheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length).setValues(snapshot.data.slice(1));
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Snapshot restored', 'Undo/Redo', 3);
}
