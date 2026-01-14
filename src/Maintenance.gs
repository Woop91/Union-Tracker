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
