/**
 * ============================================================================
 * Main.gs - Dashboard Entry Point & Triggers
 * ============================================================================
 *
 * This is the main entry point for the Union Steward Dashboard.
 * It contains trigger functions and coordinates between all modules.
 *
 * Module Architecture:
 * - Constants.gs     : Configuration and constants (single source of truth)
 * - UIService.gs     : Dialogs, sidebars, and UI components
 * - GrievanceManager.gs : Grievance lifecycle management
 * - Integrations.gs  : Drive, Calendar, and email services
 * - Maintenance.gs   : Admin tools and diagnostics
 * - FormulaService.gs: Hidden sheet and formula logic
 *
 * Build Instructions:
 * During development, keep files separate. Use build.js to merge all files
 * into ConsolidatedDashboard.gs for deployment:
 *   node build.js
 *
 * @fileoverview Main entry point and trigger functions
 * @version 2.0.0
 * @author Dashboard Team
 */

// ============================================================================
// TRIGGER FUNCTIONS
// ============================================================================

/**
 * Runs when the spreadsheet is opened
 * Sets up the custom menu and initializes the dashboard
 */
function onOpen() {
  try {
    // Create the dashboard menu
    createDashboardMenu();

    // Show welcome toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Dashboard loaded successfully',
      '🏛️ Union Dashboard',
      3
    );

  } catch (error) {
    console.error('Error in onOpen:', error);
    // Still try to create a basic menu
    SpreadsheetApp.getUi()
      .createMenu('Union Dashboard')
      .addItem('Initialize Dashboard', 'initializeDashboard')
      .addToUi();
  }
}

/**
 * Runs when a cell is edited
 * Handles auto-calculations, validations, security audit, and auto-styling
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();

    // Skip header rows
    if (row <= 1) return;

    // 1. Security Audit & Change Tracking (runs for all tracked sheets)
    handleSecurityAudit_(e);

    // 2. Handle edits in Grievance Log
    if (sheetName === SHEETS.GRIEVANCE_LOG) {
      handleGrievanceEdit(e);
      applyAutoStyleToRow_(sheet, row);  // Auto-styling
      handleStageGateWorkflow_(e);        // Escalation alerts
    }

    // 3. Handle edits in Member Directory
    if (sheetName === SHEETS.MEMBER_DIR) {
      handleMemberEdit(e);
      applyAutoStyleToRow_(sheet, row);  // Auto-styling
    }

    // 4. Handle edits in Case Checklist (auto-update progress)
    if (sheetName === 'Case Checklist' || (typeof CHECKLIST_SHEET_NAME !== 'undefined' && sheetName === CHECKLIST_SHEET_NAME)) {
      if (typeof handleChecklistEdit === 'function') {
        handleChecklistEdit(e);
      }
    }

  } catch (error) {
    console.error('Error in onEdit:', error);
    // Don't show error to user for automatic functions
  }
}

/**
 * Handles security audit logging for change tracking
 * Logs all edits to the Audit Log sheet for accountability
 * Includes sabotage protection for mass deletions (>15 cells)
 * @param {Object} e - The edit event object
 * @private
 */
function handleSecurityAudit_(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

    if (!auditSheet) {
      // Create audit log sheet if it doesn't exist
      auditSheet = ss.insertSheet(SHEETS.AUDIT_LOG);
      auditSheet.getRange('A1:F1').setValues([['Timestamp', 'User', 'Cell', 'Old Value', 'New Value', 'Alert']]);
      auditSheet.getRange('A1:F1').setFontWeight('bold').setBackground(COLORS.CARD_DARK_BG).setFontColor(COLORS.CARD_DARK_TEXT);
      auditSheet.hideSheet();
    }

    var userEmail = '';
    try {
      userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    } catch (authError) {
      userEmail = 'Auth Required';
    }

    var range = e.range;
    var numCells = range.getNumColumns() * range.getNumRows();
    var alertMessage = '';

    // SABOTAGE PROTECTION: Detect mass deletions (>15 cells cleared)
    if (e.oldValue && !e.value && numCells > 15) {
      alertMessage = 'MASS_DELETION_ALERT';

      // Send alert to Chief Steward
      var chiefEmail = '';
      try {
        chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
      } catch (configError) {
        // Config not available
      }

      if (chiefEmail) {
        try {
          MailApp.sendEmail({
            to: chiefEmail,
            subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' SABOTAGE ALERT',
            body: 'Mass deletion detected in ' + COMMAND_CONFIG.SYSTEM_NAME + '.\n\n' +
                  'User: ' + userEmail + '\n' +
                  'Sheet: ' + range.getSheet().getName() + '\n' +
                  'Range: ' + range.getA1Notation() + '\n' +
                  'Cells Affected: ' + numCells + '\n' +
                  'Time: ' + new Date().toLocaleString() + '\n\n' +
                  'Please review immediately.' +
                  COMMAND_CONFIG.EMAIL.FOOTER
          });
        } catch (emailError) {
          Logger.log('Failed to send sabotage alert: ' + emailError.message);
        }
      }

      // Log to console for visibility
      Logger.log('SABOTAGE ALERT: Mass deletion by ' + userEmail + ' in ' +
                 range.getSheet().getName() + ' (' + numCells + ' cells)');
    }

    // Detect large-scale changes (not necessarily malicious but notable)
    if (numCells > 50) {
      alertMessage = alertMessage || 'LARGE_CHANGE';
    }

    auditSheet.appendRow([
      new Date(),
      userEmail,
      range.getA1Notation() + ' (' + range.getSheet().getName() + ')',
      e.oldValue || '(empty)',
      e.value || '(deleted)',
      alertMessage
    ]);

  } catch (auditError) {
    // Silently fail - don't break user's edit for audit logging
    console.log('Audit log error: ' + auditError.message);
  }
}

/**
 * Applies automatic row styling based on theme settings
 * Includes zebra striping and status-based coloring
 * @param {Sheet} sheet - The sheet to style
 * @param {number} row - The row number to style
 * @private
 */
function applyAutoStyleToRow_(sheet, row) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    var rowRange = sheet.getRange(row, 1, 1, lastCol);

    // Apply theme font
    rowRange.setFontFamily(COMMAND_CONFIG.THEME.FONT)
            .setFontSize(COMMAND_CONFIG.THEME.FONT_SIZE)
            .setVerticalAlignment('middle');

    // Zebra striping
    if (row % 2 === 0) {
      rowRange.setBackground(COMMAND_CONFIG.THEME.ALT_ROW);
    }

    // Status-based coloring (Grievance Log only)
    if (sheet.getName() === SHEETS.GRIEVANCE_LOG) {
      var statusCell = sheet.getRange(row, GRIEVANCE_COLS.STATUS);
      var status = statusCell.getValue();

      if (COMMAND_CONFIG.STATUS_COLORS[status]) {
        var colors = COMMAND_CONFIG.STATUS_COLORS[status];
        statusCell.setBackground(colors.bg)
                  .setFontColor(colors.text)
                  .setFontWeight('bold');
      } else {
        // Reset to default if status not in mapping
        statusCell.setBackground(null)
                  .setFontColor(null)
                  .setFontWeight('normal');
      }
    }
  } catch (styleError) {
    console.log('Auto-style error: ' + styleError.message);
  }
}

/**
 * Handles stage-gate workflow and escalation alerts
 * Sends email to Chief Steward when case escalates
 * @param {Object} e - The edit event object
 * @private
 */
function handleStageGateWorkflow_(e) {
  try {
    var sheet = e.range.getSheet();
    var col = e.range.getColumn();
    var row = e.range.getRow();
    var newValue = e.value;

    // Check if status column was edited (using dynamic column reference)
    if (col === GRIEVANCE_COLS.STATUS) {
      // Only update Date Closed timestamp for closed statuses
      var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];
      if (closedStatuses.indexOf(newValue) !== -1) {
        sheet.getRange(row, GRIEVANCE_COLS.DATE_CLOSED).setValue(new Date());
      }

      // Check if this is an escalation status (reads from Config or falls back to default)
      var escalationStatuses = getEscalationStatuses_();
      if (escalationStatuses.indexOf(newValue) !== -1) {
        var memberName = sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
                        sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue();
        var caseID = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

        sendEscalationAlert_(memberName, caseID, newValue);
      }
    }

    // Also check if Current Step column was edited (use ESCALATION_STEPS for step values)
    if (col === GRIEVANCE_COLS.CURRENT_STEP) {
      var escalationSteps = getEscalationSteps_();
      if (escalationSteps.indexOf(newValue) !== -1) {
        var memberName2 = sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
                         sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue();
        var caseID2 = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

        sendEscalationAlert_(memberName2, caseID2, newValue);
      }
    }
  } catch (workflowError) {
    console.log('Workflow error: ' + workflowError.message);
  }
}

/**
 * Sends escalation alert email to Chief Steward
 * Reads email from Config sheet (column AS)
 * @param {string} memberName - Name of the member
 * @param {string} caseID - Grievance case ID
 * @param {string} status - New status/step
 * @private
 */
function sendEscalationAlert_(memberName, caseID, status) {
  // Get Chief Steward email from Config sheet
  var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);

  if (!chiefStewardEmail) {
    console.log('Chief Steward email not configured in Config sheet (column AS) - skipping escalation alert');
    return;
  }

  try {
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Case Escalation Alert';
    var body = 'ESCALATION NOTICE\n\n' +
               'Case ' + caseID + ' for ' + memberName + ' has been escalated to: ' + status + '\n\n' +
               'Immediate review is required.\n' +
               COMMAND_CONFIG.EMAIL.FOOTER;

    MailApp.sendEmail(chiefStewardEmail, subject, body);
    SpreadsheetApp.getActiveSpreadsheet().toast('Escalation alert sent to Chief Steward', 'Alert Sent', 3);
  } catch (emailError) {
    console.log('Escalation email error: ' + emailError.message);
  }
}

/**
 * Gets a value from the Config sheet by column number
 * Reads from row 3 (first data row after headers)
 * @param {number} columnNum - Column number (1-indexed)
 * @returns {string} The value from the Config sheet, or empty string if not found
 * @private
 */
function getConfigValue_(columnNum) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!configSheet) {
      console.log('Config sheet not found');
      return '';
    }

    // Config values are typically in row 3 (row 1 = section headers, row 2 = column headers)
    var value = configSheet.getRange(3, columnNum).getValue();
    return value ? String(value).trim() : '';
  } catch (e) {
    console.log('Error reading config value: ' + e.message);
    return '';
  }
}

/**
 * Gets escalation status values from Config sheet or falls back to defaults
 * Config format: comma-separated values (e.g., "In Arbitration,Appealed")
 * @returns {Array} Array of status values that trigger escalation alerts
 * @private
 */
function getEscalationStatuses_() {
  var configValue = getConfigValue_(CONFIG_COLS.ESCALATION_STATUSES);

  if (configValue) {
    return configValue.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }

  // Fall back to defaults from COMMAND_CONFIG
  return COMMAND_CONFIG.ESCALATION_STATUSES || ['In Arbitration', 'Appealed'];
}

/**
 * Gets escalation step values from Config sheet or falls back to defaults
 * Config format: comma-separated values (e.g., "Step II,Step III,Arbitration")
 * @returns {Array} Array of step values that trigger escalation alerts
 * @private
 */
function getEscalationSteps_() {
  var configValue = getConfigValue_(CONFIG_COLS.ESCALATION_STEPS);

  if (configValue) {
    return configValue.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }

  // Fall back to defaults from COMMAND_CONFIG
  return COMMAND_CONFIG.ESCALATION_STEPS || ['Step II', 'Step III', 'Arbitration'];
}

/**
 * Gets unit codes mapping from Config sheet or falls back to defaults
 * Config format: "Unit Name:CODE,Unit2:CODE2" (e.g., "Main Station:MS,Field Ops:FO")
 * @returns {Object} Object mapping unit names to code prefixes
 * @private
 */
function getUnitCodes_() {
  var configValue = getConfigValue_(CONFIG_COLS.UNIT_CODES);

  if (configValue) {
    var unitCodes = {};
    configValue.split(',').forEach(function(pair) {
      var parts = pair.split(':');
      if (parts.length === 2) {
        unitCodes[parts[0].trim()] = parts[1].trim();
      }
    });
    if (Object.keys(unitCodes).length > 0) {
      return unitCodes;
    }
  }

  // Fall back to default codes
  return {
    "Main Station": "MS",
    "Field Ops": "FO",
    "Health": "HC",
    "Admin": "AD",
    "Remote": "RM"
  };
}

/**
 * Handles edits to the Grievance Log sheet
 * Uses dynamic column references from GRIEVANCE_COLS
 * @param {Object} e - The edit event object
 */
function handleGrievanceEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Skip header row
  if (row <= 1) return;

  const sheet = e.range.getSheet();

  // Auto-update Next Action Due timestamp when any field changes
  if (col !== GRIEVANCE_COLS.NEXT_ACTION_DUE) {
    sheet.getRange(row, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue(new Date());
  }

  // If status changed, check for auto-actions
  if (col === GRIEVANCE_COLS.STATUS) {
    try {
      const settings = typeof getSettings === 'function' ? getSettings() : {};
      const grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

      // Auto-sync to calendar if enabled
      if (settings.autoSyncCalendar && grievanceId && typeof syncSingleGrievanceToCalendar === 'function') {
        syncSingleGrievanceToCalendar(grievanceId);
      }
    } catch (settingsError) {
      console.log('Settings error in handleGrievanceEdit: ' + settingsError.message);
    }
  }

  // If step dates changed, recalculate deadlines (using dynamic column references)
  const stepDateColumns = [
    GRIEVANCE_COLS.DATE_FILED,      // Step 1 filing
    GRIEVANCE_COLS.STEP2_APPEAL_FILED,  // Step 2 appeal
    GRIEVANCE_COLS.STEP3_APPEAL_FILED   // Step 3 appeal
  ];

  if (stepDateColumns.includes(col)) {
    const step = stepDateColumns.indexOf(col) + 1;
    const stepDate = e.value;

    if (stepDate && typeof calculateResponseDeadline === 'function') {
      const deadline = calculateResponseDeadline(step, new Date(stepDate));
      if (deadline) {
        // Update the corresponding due column
        const dueColumns = [GRIEVANCE_COLS.STEP1_DUE, GRIEVANCE_COLS.STEP2_DUE, GRIEVANCE_COLS.DATE_CLOSED];
        sheet.getRange(row, dueColumns[step - 1]).setValue(deadline);
      }
    }
  }
}

/**
 * Handles edits to the Member Directory sheet
 * Uses dynamic column references from MEMBER_COLS
 * @param {Object} e - The edit event object
 */
function handleMemberEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Skip header row
  if (row <= 1) return;

  const sheet = e.range.getSheet();

  // Auto-update Recent Contact Date when contact-related fields change
  var contactColumns = [MEMBER_COLS.EMAIL, MEMBER_COLS.PHONE, MEMBER_COLS.CONTACT_NOTES];
  if (contactColumns.includes(col)) {
    sheet.getRange(row, MEMBER_COLS.RECENT_CONTACT_DATE).setValue(new Date());
  }

  // Validate email format (using dynamic column reference)
  if (col === MEMBER_COLS.EMAIL) {
    const email = e.value;
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (email && !emailPattern.test(email)) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Warning: Email format may be invalid',
        'Validation',
        3
      );
    }
  }

  // Validate phone format (using dynamic column reference)
  if (col === MEMBER_COLS.PHONE) {
    const phone = e.value;
    var phonePattern = /^\d{3}-\d{3}-\d{4}$/;
    if (phone && !phonePattern.test(phone)) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Warning: Phone format may be invalid (expected: XXX-XXX-XXXX)',
        'Validation',
        3
      );
    }
  }
}

/**
 * Runs on a time-based trigger (daily)
 * Handles scheduled tasks like deadline reminders
 */
function dailyTrigger() {
  try {
    const settings = getSettings();

    // Send deadline reminders if enabled
    if (settings.emailReminders) {
      sendDeadlineReminders(settings.reminderDays);
    }

    // Log the trigger run
    logAuditEvent('DAILY_TRIGGER', {
      timestamp: new Date().toISOString(),
      remindersSent: settings.emailReminders
    });

  } catch (error) {
    console.error('Error in dailyTrigger:', error);
  }
}

// ============================================================================
// INITIALIZATION FUNCTIONS
// ============================================================================

/**
 * Initializes the dashboard for first-time setup
 * Creates all required sheets and configurations
 */
function initializeDashboard() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Initialize Dashboard',
    'This will set up the Union Steward Dashboard with all required sheets and configurations. ' +
    'Existing data will not be affected. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    showToast('Initializing dashboard...', 'Setup');

    // Create visible sheets if missing
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    for (const [key, sheetName] of Object.entries(SHEET_NAMES)) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        setupSheetStructure(sheet, key);
        showToast(`Created sheet: ${sheetName}`, 'Setup');
      }
    }

    // Setup hidden calculation sheets
    const hiddenResult = setupAllHiddenSheets();
    showToast(`Set up ${hiddenResult.created} hidden sheets`, 'Setup');

    // Setup triggers
    setupTriggers();

    // Create menu
    createDashboardMenu();

    // Log initialization
    logAuditEvent('DASHBOARD_INITIALIZED', {
      sheetsCreated: Object.keys(SHEET_NAMES).length,
      hiddenSheetsCreated: hiddenResult.created,
      initializedBy: Session.getActiveUser().getEmail()
    });

    ui.alert(
      'Setup Complete',
      'The Union Steward Dashboard has been initialized successfully!\n\n' +
      'Use the "Union Dashboard" menu to access all features.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Error initializing dashboard:', error);
    ui.alert('Error', 'Failed to initialize: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Sets up time-based triggers for the dashboard
 */
function setupTriggers() {
  // Remove existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create daily trigger at 8 AM
  ScriptApp.newTrigger('dailyTrigger')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

/**
 * Removes all triggers (for cleanup)
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
}

// ============================================================================
// HELP & DOCUMENTATION
// ============================================================================

/**
 * Shows the searchable Help Guide modal with menu breakdown and FAQ
 * v4.3.7 - Complete rewrite with search, menu reference, and FAQ
 */
function showHelpDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;600&display=swap" rel="stylesheet">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Roboto', sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #e2e8f0; min-height: 100vh; }
        .container { padding: 16px; max-height: 100vh; overflow-y: auto; }
        .search-box { position: sticky; top: 0; background: #1e293b; padding: 12px 0; z-index: 10; }
        .search-input { width: 100%; padding: 12px 16px 12px 44px; border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; background: rgba(255,255,255,0.05); color: #e2e8f0; font-size: 14px; }
        .search-input:focus { outline: none; border-color: #3b82f6; }
        .search-icon { position: absolute; left: 14px; top: 50%; transform: translateY(-50%); color: #64748b; }
        .tabs { display: flex; gap: 8px; margin: 16px 0; flex-wrap: wrap; }
        .tab { padding: 8px 16px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; cursor: pointer; font-size: 12px; color: #94a3b8; transition: all 0.2s; }
        .tab:hover { background: rgba(255,255,255,0.1); }
        .tab.active { background: #3b82f6; border-color: #3b82f6; color: white; }
        .section { margin-bottom: 20px; }
        .section-title { font-size: 14px; font-weight: 600; color: #60a5fa; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
        .section-title .material-icons { font-size: 18px; }
        .card { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; padding: 12px; margin-bottom: 8px; }
        .card-title { font-weight: 500; color: #e2e8f0; margin-bottom: 6px; font-size: 13px; }
        .card-desc { color: #94a3b8; font-size: 12px; line-height: 1.5; }
        .menu-item { display: flex; justify-content: space-between; align-items: flex-start; padding: 10px 12px; background: rgba(255,255,255,0.03); border-radius: 6px; margin-bottom: 6px; }
        .menu-item:hover { background: rgba(255,255,255,0.08); }
        .menu-path { color: #60a5fa; font-size: 11px; font-family: monospace; margin-bottom: 4px; }
        .menu-name { font-weight: 500; color: #e2e8f0; font-size: 13px; }
        .menu-desc { color: #94a3b8; font-size: 11px; margin-top: 4px; }
        .faq-item { border-left: 3px solid #3b82f6; padding-left: 12px; margin-bottom: 12px; }
        .faq-q { font-weight: 500; color: #e2e8f0; margin-bottom: 6px; font-size: 13px; }
        .faq-a { color: #94a3b8; font-size: 12px; line-height: 1.6; }
        .highlight { background: #fbbf24; color: #0f172a; padding: 0 2px; border-radius: 2px; }
        .hidden { display: none; }
        .repo-link { display: inline-flex; align-items: center; gap: 6px; color: #60a5fa; text-decoration: none; font-size: 12px; margin-top: 8px; }
        .repo-link:hover { text-decoration: underline; }
        .version { color: #64748b; font-size: 11px; text-align: center; margin-top: 20px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1); }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="search-box">
          <div style="position: relative;">
            <span class="material-icons search-icon">search</span>
            <input type="text" class="search-input" id="searchInput" placeholder="Search help topics, menus, or FAQ..." oninput="filterContent()">
          </div>
        </div>

        <div class="tabs">
          <div class="tab active" onclick="showTab('overview')">Overview</div>
          <div class="tab" onclick="showTab('menus')">Menu Reference</div>
          <div class="tab" onclick="showTab('faq')">FAQ</div>
          <div class="tab" onclick="showTab('shortcuts')">Quick Tips</div>
        </div>

        <div id="content">
          <!-- OVERVIEW TAB -->
          <div id="overview-tab" class="tab-content">
            <div class="section">
              <div class="section-title"><span class="material-icons">dashboard</span>Two-Dashboard Architecture</div>
              <div class="card">
                <div class="card-title">🛡️ Steward Dashboard (Internal)</div>
                <div class="card-desc">Comprehensive dashboard for stewards with 6 tabs: Overview, Workload, Analytics, Hot Spots, Bargaining, and Satisfaction. Contains member names and PII - for internal use only.</div>
              </div>
              <div class="card">
                <div class="card-title">👥 Member Dashboard (Public)</div>
                <div class="card-desc">PII-safe dashboard for sharing with members. Shows aggregate union stats, steward directory, and satisfaction scores without personal information.</div>
              </div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">star</span>Key Features</div>
              <div class="card">
                <div class="card-title">Grievance Tracking</div>
                <div class="card-desc">Track grievances through all steps with automatic deadline calculations based on Article 23A.</div>
              </div>
              <div class="card">
                <div class="card-title">Member Directory</div>
                <div class="card-desc">Maintain member records with contact info, union status, and steward assignments.</div>
              </div>
              <div class="card">
                <div class="card-title">Satisfaction Surveys</div>
                <div class="card-desc">8-section member satisfaction analysis covering steward ratings, leadership, communication, and more.</div>
              </div>
              <div class="card">
                <div class="card-title">Calendar & Drive Integration</div>
                <div class="card-desc">Sync deadlines to Google Calendar and auto-create Drive folders for documentation.</div>
              </div>
            </div>
          </div>

          <!-- MENU REFERENCE TAB -->
          <div id="menus-tab" class="tab-content hidden">
            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>🏛️ Union Hub Menu</div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Dashboards</div><div class="menu-name">👥 Member Dashboard</div><div class="menu-desc">Opens the public-safe member dashboard with aggregate stats</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Dashboards</div><div class="menu-name">🛡️ Steward Dashboard</div><div class="menu-desc">Opens the internal steward dashboard with full analytics</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Quick Search</div><div class="menu-name">🔍 Quick Search</div><div class="menu-desc">Fast search across members and grievances</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Multi-Select Panel</div><div class="menu-name">📋 Multi-Select Panel</div><div class="menu-desc">Select multiple grievances for bulk operations</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>🎯 Strategic Ops Menu</div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops</div><div class="menu-name">👥 Member Dashboard</div><div class="menu-desc">PII-safe aggregate statistics portal</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops</div><div class="menu-name">🛡️ Steward Dashboard</div><div class="menu-desc">Internal analytics with all 6 tabs</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > Cases</div><div class="menu-name">➕ New Case/Grievance</div><div class="menu-desc">Open the new grievance filing form</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > Cases</div><div class="menu-name">✏️ Edit Selected</div><div class="menu-desc">Edit the currently selected grievance</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > ID Engines</div><div class="menu-name">🆔 Generate Missing Member IDs</div><div class="menu-desc">Auto-generate IDs for members without one</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>🛠️ Admin Menu</div>
              <div class="menu-item"><div><div class="menu-path">Admin</div><div class="menu-name">🩺 System Diagnostics</div><div class="menu-desc">Run health check on all system components</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin</div><div class="menu-name">🔧 Repair Dashboard</div><div class="menu-desc">Fix common issues automatically</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Data Sync</div><div class="menu-name">🔄 Sync All Data Now</div><div class="menu-desc">Synchronize member and grievance data</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Validation</div><div class="menu-name">🔍 Run Bulk Validation</div><div class="menu-desc">Validate all data for consistency</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Setup</div><div class="menu-name">🔧 Setup All Hidden Sheets</div><div class="menu-desc">Initialize all calculation sheets</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>📊 509 Dashboard Menu</div>
              <div class="menu-item"><div><div class="menu-path">509 Dashboard</div><div class="menu-name">📊 Dashboard</div><div class="menu-desc">Legacy 5-tab dashboard modal</div></div></div>
              <div class="menu-item"><div><div class="menu-path">509 Dashboard > Field Access</div><div class="menu-name">📱 Pocket/Mobile View</div><div class="menu-desc">Optimize display for mobile devices</div></div></div>
              <div class="menu-item"><div><div class="menu-path">509 Dashboard > Personnel</div><div class="menu-name">➕ Add New Member</div><div class="menu-desc">Open the member registration form</div></div></div>
              <div class="menu-item"><div><div class="menu-path">509 Dashboard > Styling</div><div class="menu-name">🎨 Apply Global Styling</div><div class="menu-desc">Apply Roboto theme and zebra stripes</div></div></div>
            </div>
          </div>

          <!-- FAQ TAB -->
          <div id="faq-tab" class="tab-content hidden">
            <div class="section">
              <div class="section-title"><span class="material-icons">help</span>Frequently Asked Questions</div>

              <div class="faq-item">
                <div class="faq-q">How do I file a new grievance?</div>
                <div class="faq-a">Go to <strong>Strategic Ops > Cases & Grievances > ➕ New Case/Grievance</strong>. Fill in the member info, select the grievance type, and provide details. Deadlines are calculated automatically.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">What's the difference between the two dashboards?</div>
                <div class="faq-a">The <strong>Steward Dashboard</strong> is for internal use with full PII and 6 analytical tabs. The <strong>Member Dashboard</strong> is PII-safe for sharing with members, showing only aggregate statistics.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do deadline calculations work?</div>
                <div class="faq-a">Based on Article 23A: Step 1 has ${DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE} days for response, Step 2 has ${DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL} days to appeal, Step 3 has ${DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL} days, and Arbitration must be demanded within ${DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND} days of Step 3.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Why are some cells showing errors?</div>
                <div class="faq-a">Run <strong>Admin > 🩺 System Diagnostics</strong> to check for issues. Then try <strong>Admin > 🔧 Repair Dashboard</strong> to fix common problems automatically.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I search for a member or grievance?</div>
                <div class="faq-a">Use <strong>Union Hub > 🔍 Quick Search</strong> or <strong>Strategic Ops > 🔍 Desktop Search</strong> for advanced filtering options.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I sync deadlines to Google Calendar?</div>
                <div class="faq-a">Select a grievance, then use <strong>509 Dashboard > Calendar > 📅 Create Calendar Event</strong>. This creates reminder events for all deadline dates.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">What does the Satisfaction tab show?</div>
                <div class="faq-a">It analyzes member satisfaction surveys across 8 sections: Overall Satisfaction, Steward Ratings, Chapter Effectiveness, Local Leadership, Contract Enforcement, Communication Quality, Member Voice, and Value & Action.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I apply zebra stripes to all rows?</div>
                <div class="faq-a">Go to <strong>509 Dashboard > Styling > 🎨 Apply Global Styling</strong>. This applies Roboto font and zebra stripes to all rows in Member Directory and Grievance Log.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I start fresh with a new spreadsheet?</div>
                <div class="faq-a">Visit our GitHub repository to get a fresh copy with seed data and demo features. See the link below.</div>
              </div>
            </div>

            <a href="https://github.com/Woop91/MULTIPLE-SCRIPS-REPO" target="_blank" class="repo-link">
              <span class="material-icons" style="font-size: 16px;">open_in_new</span>
              Get a fresh copy from GitHub (with Seed & Nuke features)
            </a>
          </div>

          <!-- SHORTCUTS TAB -->
          <div id="shortcuts-tab" class="tab-content hidden">
            <div class="section">
              <div class="section-title"><span class="material-icons">bolt</span>Quick Tips</div>

              <div class="card">
                <div class="card-title">🔍 Quick Search</div>
                <div class="card-desc">Use Union Hub > Quick Search to find any member or grievance instantly. Supports partial name matching.</div>
              </div>

              <div class="card">
                <div class="card-title">📊 Dashboard Views</div>
                <div class="card-desc">Use Steward Dashboard for internal analysis, Member Dashboard for sharing with members (PII-safe).</div>
              </div>

              <div class="card">
                <div class="card-title">📱 Mobile Access</div>
                <div class="card-desc">Enable Pocket/Mobile View to hide non-essential columns when using on phones or tablets.</div>
              </div>

              <div class="card">
                <div class="card-title">🔄 Data Sync</div>
                <div class="card-desc">Run Admin > Data Sync > Sync All Data Now periodically to ensure member and grievance data stays linked.</div>
              </div>

              <div class="card">
                <div class="card-title">🎨 Visual Styling</div>
                <div class="card-desc">Apply Global Styling adds zebra stripes to ALL rows (not just data rows) for consistent appearance.</div>
              </div>

              <div class="card">
                <div class="card-title">⚡ Auto-Triggers</div>
                <div class="card-desc">Install auto-sync and midnight refresh triggers from Admin menu to keep data current automatically.</div>
              </div>
            </div>
          </div>
        </div>

        <div class="version">
          509 Dashboard v${VERSION_INFO.CURRENT} | ${VERSION_INFO.CODENAME}
        </div>
      </div>

      <script>
        function showTab(tabName) {
          // Hide all tabs
          document.querySelectorAll('.tab-content').forEach(t => t.classList.add('hidden'));
          document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));

          // Show selected tab
          document.getElementById(tabName + '-tab').classList.remove('hidden');
          event.target.classList.add('active');
        }

        function filterContent() {
          const query = document.getElementById('searchInput').value.toLowerCase();
          const items = document.querySelectorAll('.card, .menu-item, .faq-item');

          items.forEach(item => {
            const text = item.textContent.toLowerCase();
            if (query === '' || text.includes(query)) {
              item.classList.remove('hidden');
              // Highlight matches
              if (query !== '') {
                highlightText(item, query);
              } else {
                removeHighlight(item);
              }
            } else {
              item.classList.add('hidden');
            }
          });

          // Show all tabs when searching
          if (query !== '') {
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('hidden'));
          }
        }

        function highlightText(element, query) {
          // Simple highlight - just mark parent as searched
          element.style.borderLeft = '3px solid #fbbf24';
        }

        function removeHighlight(element) {
          element.style.borderLeft = '';
        }
      </script>
    </body>
    </html>
  `).setWidth(650).setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '📖 Help Guide - 509 Dashboard');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets the current version info
 * @return {Object} Version information
 */
function getVersionInfo() {
  return {
    version: '2.0.0',
    architecture: 'Modular Multi-File',
    modules: [
      'Constants.gs',
      'UIService.gs',
      'GrievanceManager.gs',
      'Integrations.gs',
      'Maintenance.gs',
      'FormulaService.gs',
      'Main.gs'
    ],
    lastUpdated: new Date().toISOString()
  };
}

/**
 * Updates a grievance with new data
 * @param {string} grievanceId - The grievance ID
 * @param {Object} updates - Fields to update
 * @return {Object} Result object
 */
function updateGrievance(grievanceId, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Grievance not found' };
    }

    // Update each provided field
    if (updates.description !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.DESCRIPTION + 1).setValue(updates.description);
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.NOTES + 1).setValue(updates.notes);
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.STATUS + 1).setValue(updates.status);
    }

    // Update timestamp
    sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.LAST_UPDATED + 1).setValue(new Date());

    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      updates: Object.keys(updates),
      updatedBy: Session.getActiveUser().getEmail()
    });

    return { success: true, message: 'Grievance updated successfully' };

  } catch (error) {
    console.error('Error updating grievance:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Handles bulk status selection from multi-select dialog
 * @param {string[]} selectedIds - Selected grievance IDs
 */
function handleBulkStatusSelection(selectedIds) {
  if (!selectedIds || selectedIds.length === 0) {
    showAlert('No grievances selected', 'Bulk Update');
    return;
  }

  // Show status selection dialog
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
    </head>
    <body style="padding: 20px;">
      <h3>Update ${selectedIds.length} Grievances</h3>
      <div class="form-group">
        <label class="form-label">New Status</label>
        <select class="form-select" id="newStatus">
          ${Object.values(GRIEVANCE_STATUS).map(s =>
            `<option value="${s}">${s}</option>`
          ).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Notes (optional)</label>
        <textarea class="form-textarea" id="notes" rows="2"></textarea>
      </div>
      <div style="margin-top: 20px; text-align: right;">
        <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        <button class="btn btn-primary" onclick="applyUpdate()">Update All</button>
      </div>
      <script>
        const ids = ${JSON.stringify(selectedIds)};
        function applyUpdate() {
          const status = document.getElementById('newStatus').value;
          const notes = document.getElementById('notes').value;
          google.script.run
            .withSuccessHandler(function(r) {
              alert(r.success ? r.message : 'Error: ' + r.error);
              google.script.host.close();
            })
            .bulkUpdateGrievanceStatus(ids, status, notes);
        }
      </script>
    </body>
    </html>
  `).setWidth(400).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Status Update');
}

// ============================================================================
// MEMBER MANAGEMENT
// ============================================================================

/**
 * Shows new member dialog
 */
function showNewMemberDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .form-container { padding: 20px; }
        .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }
      </style>
    </head>
    <body>
      <div class="form-container">
        <form id="memberForm">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">First Name *</label>
              <input type="text" class="form-input" id="firstName" required>
            </div>
            <div class="form-group">
              <label class="form-label">Last Name *</label>
              <input type="text" class="form-input" id="lastName" required>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Employee ID</label>
              <input type="text" class="form-input" id="employeeId" placeholder="XX000000">
            </div>
            <div class="form-group">
              <label class="form-label">Department *</label>
              <select class="form-select" id="department" required>
                <option value="">Select department...</option>
              </select>
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Job Title</label>
              <input type="text" class="form-input" id="jobTitle">
            </div>
            <div class="form-group">
              <label class="form-label">Hire Date</label>
              <input type="date" class="form-input" id="hireDate">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Email</label>
              <input type="email" class="form-input" id="email">
            </div>
            <div class="form-group">
              <label class="form-label">Phone</label>
              <input type="tel" class="form-input" id="phone" placeholder="XXX-XXX-XXXX">
            </div>
          </div>

          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Union Status</label>
              <select class="form-select" id="unionStatus">
                <option value="Full Member">Full Member</option>
                <option value="Agency Fee">Agency Fee</option>
                <option value="Non-Member">Non-Member</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Status</label>
              <select class="form-select" id="status">
                <option value="Active">Active</option>
                <option value="On Leave">On Leave</option>
                <option value="Terminated">Terminated</option>
              </select>
            </div>
          </div>

          <div style="margin-top: 20px; text-align: right;">
            <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button type="submit" class="btn btn-primary">Add Member</button>
          </div>
        </form>
      </div>

      <script>
        // Load departments
        google.script.run
          .withSuccessHandler(function(depts) {
            const select = document.getElementById('department');
            depts.forEach(function(d) {
              const opt = document.createElement('option');
              opt.value = d;
              opt.textContent = d;
              select.appendChild(opt);
            });
            // Add "Other" option
            const other = document.createElement('option');
            other.value = 'Other';
            other.textContent = 'Other';
            select.appendChild(other);
          })
          .getDepartmentList();

        document.getElementById('memberForm').addEventListener('submit', function(e) {
          e.preventDefault();
          const memberData = {
            firstName: document.getElementById('firstName').value,
            lastName: document.getElementById('lastName').value,
            employeeId: document.getElementById('employeeId').value,
            department: document.getElementById('department').value,
            jobTitle: document.getElementById('jobTitle').value,
            hireDate: document.getElementById('hireDate').value,
            email: document.getElementById('email').value,
            phone: document.getElementById('phone').value,
            unionStatus: document.getElementById('unionStatus').value,
            status: document.getElementById('status').value
          };

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                alert('Member added successfully!');
                google.script.host.close();
              } else {
                alert('Error: ' + result.error);
              }
            })
            .addNewMember(memberData);
        });
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Member');
}

/**
 * Adds a new member to the directory
 * @param {Object} memberData - Member information
 * @return {Object} Result object
 */
function addNewMember(memberData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

    if (!sheet) {
      return { success: false, error: 'Member Directory sheet not found' };
    }

    // Generate new member ID
    const lastRow = sheet.getLastRow();
    const newId = `MEM-${String(lastRow).padStart(5, '0')}`;

    // Prepare row data
    const rowData = [
      newId,
      memberData.firstName,
      memberData.lastName,
      memberData.employeeId,
      memberData.department,
      memberData.jobTitle,
      memberData.hireDate ? new Date(memberData.hireDate) : '',
      memberData.hireDate ? new Date(memberData.hireDate) : '', // Seniority same as hire
      memberData.email,
      memberData.phone,
      memberData.status || 'Active',
      memberData.unionStatus || 'Full Member',
      '',  // Notes
      new Date()  // Last Updated
    ];

    sheet.appendRow(rowData);

    logAuditEvent(AUDIT_EVENTS.MEMBER_ADDED, {
      memberId: newId,
      name: `${memberData.firstName} ${memberData.lastName}`,
      addedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      memberId: newId,
      message: 'Member added successfully'
    };

  } catch (error) {
    console.error('Error adding member:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Starts a grievance for a specific member (from quick action)
 */
function startGrievanceForMember() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.MEMBER_DIRECTORY) {
    showAlert('Please select a member in the Member Directory', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    showAlert('Please select a member row', 'No Selection');
    return;
  }

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const memberId = data[MEMBER_COLUMNS.ID];
  const memberName = `${data[MEMBER_COLUMNS.FIRST_NAME]} ${data[MEMBER_COLUMNS.LAST_NAME]}`;

  // Open new grievance dialog pre-populated with member info
  const html = HtmlService.createHtmlOutput(`
    <script>
      // Store member info for the form
      sessionStorage.setItem('prefillMemberId', '${memberId}');
      sessionStorage.setItem('prefillMemberName', '${memberName}');
      google.script.host.close();
      google.script.run.showNewGrievanceDialog();
    </script>
  `).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Loading...');
}
