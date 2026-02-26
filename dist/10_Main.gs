/**
 * ============================================================================
 * 10_Main.gs - Dashboard Entry Point & Triggers
 * ============================================================================
 *
 * This is the main entry point for the Union Steward Dashboard.
 * It contains trigger functions and coordinates between all modules.
 *
 * Module Architecture:
 * - 01_Core.gs       : Configuration and constants (single source of truth)
 * - 03_UIComponents.gs : Dialogs, sidebars, and UI components
 * - 04_GrievanceManager.gs : Grievance lifecycle management
 * - 05_Integrations.gs : Drive, Calendar, and email services
 * - 06_Maintenance.gs : Admin tools and diagnostics
 * - 08a_SheetSetup.gs : Hidden sheet and formula logic
 *
 * Build Instructions:
 * During development, keep files separate. Use build.js to merge all files
 * into ConsolidatedDashboard.gs for deployment:
 *   node build.js
 *
 * @fileoverview Main entry point and trigger functions
 * @version 4.7.0
 * @author Dashboard Team
 */

// ============================================================================
// TRIGGER FUNCTIONS
// ============================================================================

/**
 * Runs when the spreadsheet is opened
 * Sets up the custom menu, applies tab colors, and initializes the dashboard
 */
function onOpen() {
  try {
    // Create the dashboard menu immediately — this is the critical path (F47 perf fix)
    createDashboardMenu();

    // Defer non-critical initialization to a 1-second timed trigger so the
    // spreadsheet UI is not blocked by heavy I/O (column sync, tab colors,
    // hidden-sheet enforcement, etc.)
    try {
      ScriptApp.newTrigger('onOpenDeferred_')
        .timeBased()
        .after(1000)
        .create();
    } catch (triggerError) {
      // If trigger creation fails (quota, permissions), run inline as fallback
      console.log('Deferred trigger failed, running inline: ' + triggerError.message);
      onOpenDeferred_();
    }

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
 * Deferred onOpen work — runs via a one-shot timed trigger so the
 * spreadsheet opens fast. Handles column sync, sheet validation,
 * tab colors, and hidden-sheet enforcement. (F47)
 * @private
 */
function onOpenDeferred_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Sync column maps from actual sheet headers so that column constants
    // (MEMBER_COLS, GRIEVANCE_COLS, CONFIG_COLS, etc.) reflect the real
    // spreadsheet layout — even if columns were reordered manually.
    try {
      syncColumnMaps();
    } catch (syncError) {
      console.log('Column sync skipped: ' + syncError.message);
    }

    // Ensure all primary sheets have enough columns for current header maps.
    // This prevents "columns are out of bounds" errors when sheets were
    // created by an older version with fewer columns.
    try {
      ensureAllSheetColumns_();
    } catch (colError) {
      console.log('Column check skipped: ' + colError.message);
    }

    // Apply tab colors automatically on open
    try {
      if (typeof applyTabColors_ === 'function') {
        applyTabColors_(ss);
      }
    } catch (tabError) {
      console.log('Tab colors not applied: ' + tabError.message);
    }

    // Enforce hidden sheets on every open (prevents mobile visibility)
    try {
      enforceHiddenSheets();
    } catch (hideError) {
      console.log('Hidden sheet enforcement skipped: ' + hideError.message);
    }

    // Proactive Constant Contact token health check (non-blocking toast)
    try {
      checkConstantContactHealth();
    } catch (ccError) {
      console.log('CC health check skipped: ' + ccError.message);
    }

    // Show welcome toast
    ss.toast(
      'Dashboard loaded successfully',
      '🏛️ Union Dashboard',
      3
    );

  } catch (error) {
    console.error('Error in onOpenDeferred_:', error);
  } finally {
    // Clean up the one-shot trigger so it does not accumulate
    cleanUpOnOpenTrigger_();
  }
}

/**
 * Removes any onOpenDeferred_ triggers created by onOpen.
 * Prevents trigger accumulation across multiple spreadsheet opens.
 * @private
 */
function cleanUpOnOpenTrigger_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onOpenDeferred_') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    console.log('Trigger cleanup error: ' + e.message);
  }
}

/**
 * Runs when a cell is edited
 * This is the SINGLE entry point for all edit triggers.
 * Dispatches to specialized handlers based on sheet and context.
 *
 * Consolidated handlers (v4.5.0):
 * - handleSecurityAudit_: Security audit and change tracking
 * - handleGrievanceEdit: Grievance-specific edits
 * - handleMemberEdit: Member directory edits
 * - handleChecklistEdit: Case checklist updates
 * - onEditMultiSelect: Multi-select dropdown handling (08_SheetUtils.gs)
 * - onEditAutoSync: Auto-sync to external services (09_Dashboards.gs)
 *
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  try {
    // Restore column positions that syncColumnMaps() resolved in onOpen().
    // Without this, each onEdit execution would start with array-order defaults
    // which are wrong if a user manually reordered columns.
    try { loadCachedColumnMaps_(); } catch (_cacheErr) { /* use defaults */ }

    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();

    // Skip header rows and hidden/calculation sheets (start with '_')
    if (row <= 1 || sheetName.charAt(0) === '_') return;

    // Fast-exit for sheets that have no onEdit logic (e.g. dashboard output tabs)
    var EDITABLE_SHEETS_ = [
      SHEETS.GRIEVANCE_LOG, SHEETS.MEMBER_DIR, SHEETS.CONFIG,
      SHEETS.CASE_CHECKLIST, SHEETS.VOLUNTEER_HOURS, SHEETS.MEETING_ATTENDANCE
    ];
    if (typeof CHECKLIST_SHEET_NAME !== 'undefined') EDITABLE_SHEETS_.push(CHECKLIST_SHEET_NAME);
    if (EDITABLE_SHEETS_.indexOf(sheetName) === -1) return;

    // ========================================
    // 1. Security Audit & Change Tracking
    // ========================================
    handleSecurityAudit_(e);

    // ========================================
    // 2. Multi-Select Dropdown Handler
    // ========================================
    if (typeof onEditMultiSelect === 'function') {
      try {
        onEditMultiSelect(e);
      } catch (multiSelectError) {
        console.log('MultiSelect handler error: ' + multiSelectError.message);
      }
    }

    // ========================================
    // 3. Sheet-Specific Handlers (else-if: only one sheet matches)
    // ========================================

    if (sheetName === SHEETS.GRIEVANCE_LOG) {
      // Grievance Log edits
      handleGrievanceEdit(e);
      applyAutoStyleToRow_(sheet, row);
      handleStageGateWorkflow_(e);

      // Bidirectional sync: sync custom dropdown values back to Config
      try { syncDropdownToConfig_(e, sheetName); }
      catch (syncErr) { console.log('Config sync error: ' + syncErr.message); }

      if (typeof sortGrievanceLogByStatus === 'function') {
        try { sortGrievanceLogByStatus(); }
        catch (sortError) { Logger.log('Auto-sort error: ' + sortError.message); }
      }

      if (typeof onEditAutoSync === 'function') {
        try { onEditAutoSync(e); }
        catch (syncError) { console.log('AutoSync handler error: ' + syncError.message); }
      }

    } else if (sheetName === SHEETS.MEMBER_DIR) {
      // Member Directory edits
      handleMemberEdit(e);
      applyAutoStyleToRow_(sheet, row);

      // Bidirectional sync: sync custom dropdown values back to Config
      try { syncDropdownToConfig_(e, sheetName); }
      catch (syncErr) { console.log('Config sync error: ' + syncErr.message); }

    } else if (sheetName === SHEETS.CASE_CHECKLIST || (typeof CHECKLIST_SHEET_NAME !== 'undefined' && sheetName === CHECKLIST_SHEET_NAME)) {
      // Case Checklist edits
      if (typeof handleChecklistEdit === 'function') {
        handleChecklistEdit(e);
      }

    } else if (sheetName === SHEETS.VOLUNTEER_HOURS) {
      // Volunteer Hours edits
      if (typeof syncVolunteerHoursToMemberDirectory === 'function') {
        try { syncVolunteerHoursToMemberDirectory(); }
        catch (syncError) { Logger.log('Volunteer Hours sync error: ' + syncError.message); }
      }

    } else if (sheetName === SHEETS.MEETING_ATTENDANCE) {
      // Meeting Attendance edits
      if (typeof syncMeetingAttendanceToMemberDirectory === 'function') {
        try { syncMeetingAttendanceToMemberDirectory(); }
        catch (syncError) { Logger.log('Meeting Attendance sync error: ' + syncError.message); }
      }

    } else if (sheetName === SHEETS.CONFIG) {
      // Config sheet edits — sync STEWARDS column to Member Directory
      handleConfigStewardEdit_(e);
    }

    // ========================================
    // 4. Additional Audit Logging (high-value sheets only)
    // ========================================
    if (typeof onEditAudit === 'function' &&
        (sheetName === SHEETS.GRIEVANCE_LOG || sheetName === SHEETS.MEMBER_DIR)) {
      try {
        onEditAudit(e);
      } catch (auditError) {
        console.log('Audit handler error: ' + auditError.message);
      }
    }

  } catch (error) {
    console.error('Error in onEdit:', error);
  }
}

/**
 * Simple trigger: fires when the user changes cell selection.
 * Delegates to the multi-select auto-open handler when the
 * user has enabled it via Tools > Multi-Select > Enable Auto-Open.
 *
 * Note: onSelectionChange is only available as a simple trigger
 * in Google Apps Script — it cannot be installed via ScriptApp.newTrigger().
 *
 * @param {Object} e - The selection change event object
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;

  try {
    var autoOpen = PropertiesService.getUserProperties()
      .getProperty('multiSelectAutoOpen');
    if (autoOpen !== 'true') return;

    if (typeof onSelectionChangeMultiSelect === 'function') {
      onSelectionChangeMultiSelect(e);
    }
  } catch (_err) {
    // Silent — selection-change triggers must not surface errors
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
      auditSheet.appendRow(['Timestamp', 'Event Type', 'User', 'Details', 'Session ID']);
      auditSheet.setFrozenRows(1);
      setSheetVeryHidden_(auditSheet);
    }

    var userEmail = '';
    try {
      userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    } catch (_authError) {
      userEmail = 'Auth Required';
    }

    var range = e.range;
    var numCells = range.getNumColumns() * range.getNumRows();
    var alertMessage = '';

    // SABOTAGE PROTECTION: Detect mass deletions (>15 cells cleared)
    // Exclude multi-cell paste operations (e.value is undefined for pastes AND deletions,
    // but pastes set range values while deletions clear them)
    var rangeIsEmpty = range.isBlank();
    if (numCells > 15 && !e.value && rangeIsEmpty) {
      alertMessage = 'MASS_DELETION_ALERT';

      // Send alert to Chief Steward
      var chiefEmail = '';
      try {
        chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
      } catch (_configError) {
        // Config not available
      }

      if (chiefEmail && MailApp.getRemainingDailyQuota() > 0) {
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
          Logger.log('Sabotage alert email failed (simple onEdit limit): ' + emailError.message);
        }
      }

      // Log to console for visibility (PII-safe)
      secureLog('SabotageDetection', 'Mass deletion detected', {
        email: userEmail,
        sheet: range.getSheet().getName(),
        cellsAffected: numCells
      });

      // Record through centralized security event system
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('MASS_DELETION', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.CRITICAL : 'CRITICAL',
          'Mass deletion of ' + numCells + ' cells detected in ' + range.getSheet().getName(),
          { email: userEmail, sheet: range.getSheet().getName(), range: range.getA1Notation(), cellsAffected: numCells });
      }
    }

    // Detect large-scale changes (not necessarily malicious but notable)
    if (numCells > 50) {
      alertMessage = alertMessage || 'LARGE_CHANGE';
    }

    var details = JSON.stringify({
      cell: range.getA1Notation(),
      sheet: range.getSheet().getName(),
      oldValue: e.oldValue || '(empty)',
      newValue: e.value || '(deleted)',
      alert: alertMessage || ''
    });

    auditSheet.appendRow([
      new Date(),
      alertMessage || 'CELL_EDIT',
      userEmail,
      details,
      ''
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
 * Reads email from Config sheet (column AQ)
 * @param {string} memberName - Name of the member
 * @param {string} caseID - Grievance case ID
 * @param {string} status - New status/step
 * @private
 */
function sendEscalationAlert_(memberName, caseID, status) {
  // Get Chief Steward email from Config sheet
  var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);

  if (!chiefStewardEmail) {
    console.log('Chief Steward email not configured in Config sheet (column AQ) - skipping escalation alert');
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
 * Reads all non-empty rows from the ESCALATION_STATUSES config column.
 * @returns {Array} Array of status values that trigger escalation alerts
 * @private
 */
function getEscalationStatuses_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var values = getConfigValues(configSheet, CONFIG_COLS.ESCALATION_STATUSES);
      if (values.length > 0) return values;
    }
  } catch (e) {
    console.log('Error reading escalation statuses: ' + e.message);
  }

  // Fall back to defaults from COMMAND_CONFIG
  return COMMAND_CONFIG.ESCALATION_STATUSES || ['In Arbitration', 'Appealed'];
}

/**
 * Gets escalation step values from Config sheet or falls back to defaults
 * Reads all non-empty rows from the ESCALATION_STEPS config column.
 * @returns {Array} Array of step values that trigger escalation alerts
 * @private
 */
function getEscalationSteps_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var values = getConfigValues(configSheet, CONFIG_COLS.ESCALATION_STEPS);
      if (values.length > 0) return values;
    }
  } catch (e) {
    console.log('Error reading escalation steps: ' + e.message);
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
 *
 * Date Override: Stewards can overwrite auto-calculated date columns
 * (FILING_DEADLINE, STEP1_DUE, STEP2_APPEAL_DUE, STEP2_DUE, STEP3_APPEAL_DUE).
 * When overridden, a cell note is added and all downstream deadlines are
 * recalculated from the steward-provided date.
 *
 * @param {Object} e - The edit event object
 */
function handleGrievanceEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Skip header row
  if (row <= 1) return;

  const sheet = e.range.getSheet();

  // Only update Next Action Due when status or step date changes, not on every edit
  var statusAndDateCols = [
    GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.CURRENT_STEP,
    GRIEVANCE_COLS.DATE_FILED, GRIEVANCE_COLS.STEP1_RCVD,
    GRIEVANCE_COLS.STEP2_APPEAL_FILED, GRIEVANCE_COLS.STEP2_RCVD,
    GRIEVANCE_COLS.STEP3_APPEAL_FILED, GRIEVANCE_COLS.DATE_CLOSED
  ];
  if (statusAndDateCols.indexOf(col) !== -1) {
    // Compute the actual next deadline based on current step, matching the logic
    // in recalculateDownstreamDeadlines_. Only set to a real deadline date, not now().
    var currentStep = sheet.getRange(row, GRIEVANCE_COLS.CURRENT_STEP).getValue();
    var status = sheet.getRange(row, GRIEVANCE_COLS.STATUS).getValue();
    var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];

    if (closedStatuses.indexOf(status) === -1 && currentStep) {
      var nextActionDate = '';
      if (currentStep === 'Informal') {
        nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.FILING_DEADLINE).getValue();
      } else if (currentStep === 'Step I') {
        nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP1_DUE).getValue();
      } else if (currentStep === 'Step II') {
        nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP2_DUE).getValue();
      } else if (currentStep === 'Step III') {
        nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP3_APPEAL_DUE).getValue();
      }

      if (nextActionDate instanceof Date) {
        sheet.getRange(row, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue(nextActionDate);
        var today = new Date();
        today.setHours(0, 0, 0, 0);
        var daysTo = Math.floor((nextActionDate - today) / (1000 * 60 * 60 * 24));
        sheet.getRange(row, GRIEVANCE_COLS.DAYS_TO_DEADLINE).setValue(daysTo);
      }
    } else if (closedStatuses.indexOf(status) !== -1) {
      // Clear deadline for closed grievances
      sheet.getRange(row, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue('');
      sheet.getRange(row, GRIEVANCE_COLS.DAYS_TO_DEADLINE).setValue('');
    }
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

  // ─── STEWARD DATE OVERRIDE ───
  // When a steward edits an auto-calculated deadline column, mark it as
  // overridden and recalculate all downstream dates from the new value.
  var calculatedDateCols = [
    GRIEVANCE_COLS.FILING_DEADLINE,    // H - normally Incident Date + 21
    GRIEVANCE_COLS.STEP1_DUE,          // J - normally Date Filed + 30
    GRIEVANCE_COLS.STEP2_APPEAL_DUE,   // L - normally Step I Rcvd + 10
    GRIEVANCE_COLS.STEP2_DUE,          // N - normally Step II Appeal Filed + 30
    GRIEVANCE_COLS.STEP3_APPEAL_DUE    // P - normally Step II Rcvd + 30
  ];

  if (calculatedDateCols.indexOf(col) !== -1 && e.value) {
    var overrideDate = new Date(e.value);
    if (!isNaN(overrideDate.getTime())) {
      // Mark cell as steward-overridden so syncGrievanceFormulasToLog respects it
      sheet.getRange(row, col).setNote('Steward override');

      // Recalculate downstream deadlines from this override
      recalculateDownstreamDeadlines_(sheet, row, col, overrideDate);
    }
  }

  // If a SOURCE date changed (e.g., Incident Date, Date Filed, Step I Rcvd),
  // clear override notes on the dependent calculated column and recalculate
  var sourceToDependentMap = {};
  sourceToDependentMap[GRIEVANCE_COLS.INCIDENT_DATE] = GRIEVANCE_COLS.FILING_DEADLINE;
  sourceToDependentMap[GRIEVANCE_COLS.DATE_FILED] = GRIEVANCE_COLS.STEP1_DUE;
  sourceToDependentMap[GRIEVANCE_COLS.STEP1_RCVD] = GRIEVANCE_COLS.STEP2_APPEAL_DUE;
  sourceToDependentMap[GRIEVANCE_COLS.STEP2_APPEAL_FILED] = GRIEVANCE_COLS.STEP2_DUE;
  sourceToDependentMap[GRIEVANCE_COLS.STEP2_RCVD] = GRIEVANCE_COLS.STEP3_APPEAL_DUE;

  if (sourceToDependentMap[col] !== undefined) {
    var dependentCol = sourceToDependentMap[col];
    var dependentCell = sheet.getRange(row, dependentCol);
    var note = dependentCell.getNote();
    // Only clear override if the source changed - the formula recalc should take over
    if (note === 'Steward override') {
      dependentCell.setNote('');
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
        const dueColumns = [GRIEVANCE_COLS.STEP1_DUE, GRIEVANCE_COLS.STEP2_DUE, GRIEVANCE_COLS.STEP3_APPEAL_DUE];
        sheet.getRange(row, dueColumns[step - 1]).setValue(deadline);
      }
    }
  }
}

/**
 * Recalculate downstream deadline dates when a steward overrides a calculated date.
 * Uses the override value as the new basis for all subsequent deadlines in the chain.
 *
 * Chain: Filing Deadline -> Step1 Due -> Step2 Appeal Due -> Step2 Due -> Step3 Appeal Due
 * Each subsequent deadline uses the NEXT source date if available, or falls through.
 *
 * @param {Sheet} sheet - The Grievance Log sheet
 * @param {number} row - Data row (1-indexed)
 * @param {number} overriddenCol - The column that was overridden
 * @param {Date} overrideDate - The steward-provided date
 * @private
 */
function recalculateDownstreamDeadlines_(sheet, row, overriddenCol, overrideDate) {
  // The deadline chain maps: overridden calculated col -> downstream calculated cols
  // Each downstream deadline depends on a specific source date column.
  // When a calculated date is overridden, we don't change downstream dates that have
  // their OWN source dates already entered - those will recalculate from their own source.
  // We only cascade when downstream source dates are empty.

  var deadlineChain = [
    { calcCol: GRIEVANCE_COLS.FILING_DEADLINE, sourceCol: GRIEVANCE_COLS.INCIDENT_DATE, days: 21 },
    { calcCol: GRIEVANCE_COLS.STEP1_DUE, sourceCol: GRIEVANCE_COLS.DATE_FILED, days: 30 },
    { calcCol: GRIEVANCE_COLS.STEP2_APPEAL_DUE, sourceCol: GRIEVANCE_COLS.STEP1_RCVD, days: 10 },
    { calcCol: GRIEVANCE_COLS.STEP2_DUE, sourceCol: GRIEVANCE_COLS.STEP2_APPEAL_FILED, days: 30 },
    { calcCol: GRIEVANCE_COLS.STEP3_APPEAL_DUE, sourceCol: GRIEVANCE_COLS.STEP2_RCVD, days: 30 }
  ];

  // Find position of overridden column in the chain
  var startIdx = -1;
  for (var i = 0; i < deadlineChain.length; i++) {
    if (deadlineChain[i].calcCol === overriddenCol) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) return;

  // Recalculate Next Action Due and Days to Deadline based on current step
  var currentStep = sheet.getRange(row, GRIEVANCE_COLS.CURRENT_STEP).getValue();
  var status = sheet.getRange(row, GRIEVANCE_COLS.STATUS).getValue();
  var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];

  if (closedStatuses.indexOf(status) === -1 && currentStep) {
    var nextActionDate = '';
    if (currentStep === 'Informal') {
      nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.FILING_DEADLINE).getValue();
    } else if (currentStep === 'Step I') {
      nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP1_DUE).getValue();
    } else if (currentStep === 'Step II') {
      nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP2_DUE).getValue();
    } else if (currentStep === 'Step III') {
      nextActionDate = sheet.getRange(row, GRIEVANCE_COLS.STEP3_APPEAL_DUE).getValue();
    }

    if (nextActionDate instanceof Date) {
      sheet.getRange(row, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue(nextActionDate);
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      var daysTo = Math.floor((nextActionDate - today) / (1000 * 60 * 60 * 24));
      sheet.getRange(row, GRIEVANCE_COLS.DAYS_TO_DEADLINE).setValue(daysTo);
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

  // Sync IS_STEWARD changes to Config STEWARDS list in real time
  if (col === MEMBER_COLS.IS_STEWARD) {
    var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
    var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
    var fullName = (firstName || '').toString().trim() + ' ' + (lastName || '').toString().trim();

    if (fullName.trim()) {
      if (isTruthyValue(e.value)) {
        // Promoted to steward — add to Config if not already present
        addToConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);
      } else {
        // Demoted from steward — remove from Config
        removeFromConfigDropdown_(CONFIG_COLS.STEWARDS, fullName);
      }
    }
  }
}

/**
 * Bidirectional sync: when a user enters a custom value in a dropdown column,
 * add it to the corresponding Config sheet column so it appears for future use.
 * @param {Object} e - The edit event object
 * @param {string} sheetName - Name of the edited sheet
 * @private
 */
function syncDropdownToConfig_(e, sheetName) {
  // Accept e.value (simple edits) or fall back to reading the cell directly
  // (handles pastes, programmatic setValue, and multi-cell edits)
  var newValue = e.value;
  if (!newValue && e.range && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
    try { newValue = e.range.getValue(); } catch (_) { /* skip */ }
  }
  if (!newValue || typeof newValue !== 'string' || newValue.trim() === '') return;
  newValue = newValue.trim();

  var col = e.range.getColumn();

  // Look up the Config column from DROPDOWN_MAP and MULTI_SELECT_COLS.
  // Both maps are checked so that custom values typed into either single-select
  // or multi-select columns get synced back to Config.
  var ddEntries = (sheetName === SHEETS.MEMBER_DIR) ? DROPDOWN_MAP.MEMBER_DIR
                : (sheetName === SHEETS.GRIEVANCE_LOG) ? DROPDOWN_MAP.GRIEVANCE_LOG
                : [];
  var msEntries = (sheetName === SHEETS.MEMBER_DIR) ? MULTI_SELECT_COLS.MEMBER_DIR
                : (sheetName === SHEETS.GRIEVANCE_LOG) ? MULTI_SELECT_COLS.GRIEVANCE_LOG
                : [];
  if (ddEntries.length === 0 && msEntries.length === 0) return;

  var configCol = null;
  var isMultiSelect = false;
  for (var d = 0; d < ddEntries.length; d++) {
    if (ddEntries[d].col === col) { configCol = ddEntries[d].configCol; break; }
  }
  if (!configCol) {
    for (var ms = 0; ms < msEntries.length; ms++) {
      if (msEntries[ms].col === col) { configCol = msEntries[ms].configCol; isMultiSelect = true; break; }
    }
  }

  if (!configCol) return; // Not a synced dropdown or multi-select column

  // For multi-select columns, split comma-separated values and sync each individually
  var valuesToSync = isMultiSelect ? newValue.split(',') : [newValue];

  // Filter out pure-numeric values — they're data-entry errors (e.g. index numbers
  // typed instead of text labels like "Email", "Phone").  All dropdown/multi-select
  // Config columns expect text labels, never bare integers.
  var filteredValues = [];
  for (var fv = 0; fv < valuesToSync.length; fv++) {
    var candidate = valuesToSync[fv].trim();
    if (candidate && !/^\d+$/.test(candidate)) {
      filteredValues.push(candidate);
    }
  }
  valuesToSync = filteredValues;
  if (valuesToSync.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return;

  // Build set of existing Config values for this column
  var existingSet = {};
  var lastRow = configSheet.getLastRow();
  var configRows = lastRow >= 3 ? lastRow - 2 : 0;
  if (configRows > 0) {
    var existingValues = configSheet.getRange(3, configCol, configRows, 1).getValues();
    for (var i = 0; i < existingValues.length; i++) {
      var ev = (existingValues[i][0] || '').toString().trim();
      if (ev) existingSet[ev] = true;
    }
  }

  // Add each value that doesn't already exist in Config
  for (var v = 0; v < valuesToSync.length; v++) {
    var val = valuesToSync[v].trim();
    if (val && !existingSet[val]) {
      addToConfigDropdown_(configCol, val);
      existingSet[val] = true;
    }
  }
}

/**
 * Handles edits to the Config sheet's STEWARDS column.
 * When a steward name is added or removed in Config, updates the matching
 * member's IS_STEWARD field in the Member Directory.
 * @param {Object} e - The edit event object
 * @private
 */
function handleConfigStewardEdit_(e) {
  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Only handle edits to the STEWARDS column (H), data rows (row >= 3)
  if (col !== CONFIG_COLS.STEWARDS || row < 3) return;

  var newValue = (e.value || '').toString().trim();
  var oldValue = (e.oldValue || '').toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) return;

  var memberData = memberSheet.getDataRange().getValues();

  // If a name was removed (cell cleared or old value replaced), set that member to No
  if (oldValue && oldValue !== newValue) {
    for (var i = 1; i < memberData.length; i++) {
      var firstName = (memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '').toString().trim();
      var lastName = (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '').toString().trim();
      var fullName = firstName + ' ' + lastName;
      if (fullName === oldValue) {
        memberSheet.getRange(i + 1, MEMBER_COLS.IS_STEWARD).setValue('No');
        break;
      }
    }
  }

  // If a new name was added, set that member to Yes
  if (newValue) {
    for (var j = 1; j < memberData.length; j++) {
      var fn = (memberData[j][MEMBER_COLS.FIRST_NAME - 1] || '').toString().trim();
      var ln = (memberData[j][MEMBER_COLS.LAST_NAME - 1] || '').toString().trim();
      var full = fn + ' ' + ln;
      if (full === newValue) {
        var currentStatus = memberData[j][MEMBER_COLS.IS_STEWARD - 1];
        if (!isTruthyValue(currentStatus)) {
          memberSheet.getRange(j + 1, MEMBER_COLS.IS_STEWARD).setValue('Yes');
        }
        break;
      }
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

    // Update meeting statuses: activate today's meetings, deactivate expired ones
    var meetingStatusResult = { activated: 0, deactivated: 0 };
    try {
      if (typeof updateMeetingStatuses === 'function') {
        meetingStatusResult = updateMeetingStatuses();
      }
    } catch (e) {
      console.error('Meeting status update error:', e);
    }

    // Process meeting doc notifications (agenda 3 days before, notes 1 day before, publish 1 day after)
    var meetingDocResult = { agendaSent: 0, notesSent: 0, notesPublished: 0 };
    try {
      if (typeof processMeetingDocNotifications === 'function') {
        meetingDocResult = processMeetingDocNotifications();
      }
    } catch (e) {
      console.error('Meeting doc notification error:', e);
    }

    // Cleanup expired meeting check-in records (>90 days old)
    var meetingRowsCleaned = 0;
    try {
      meetingRowsCleaned = cleanupExpiredMeetings();
    } catch (e) {
      console.error('Meeting cleanup error:', e);
    }

    // Log the trigger run
    logAuditEvent('DAILY_TRIGGER', {
      timestamp: new Date().toISOString(),
      remindersSent: settings.emailReminders,
      meetingsActivated: meetingStatusResult.activated,
      meetingsDeactivated: meetingStatusResult.deactivated,
      meetingDocAgendaSent: meetingDocResult.agendaSent,
      meetingDocNotesSent: meetingDocResult.notesSent,
      meetingDocNotesPublished: meetingDocResult.notesPublished,
      meetingRowsCleaned: meetingRowsCleaned
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
  // Delegate to CREATE_DASHBOARD which properly creates all sheets
  // with headers, formatting, validations, and hidden calculation sheets
  CREATE_DASHBOARD();
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
 * Removes dashboard-specific triggers (dailyTrigger only)
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
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
      ${getMobileOptimizedHead()}
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
        .tabs { display: flex; gap: 6px; margin: 16px 0; flex-wrap: wrap; }
        .tab { padding: 8px 12px; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; cursor: pointer; font-size: 11px; color: #94a3b8; transition: all 0.2s; }
        .tab:hover { background: rgba(255,255,255,0.1); }
        .tab.active { background: #3b82f6; border-color: #3b82f6; color: white; }
        .section { margin-bottom: 20px; }
        .section-title { font-size: 14px; font-weight: 600; color: #60a5fa; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }
        .section-title .material-icons { font-size: 18px; }
        .card { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 8px; padding: 12px; margin-bottom: 8px; }
        .card-title { font-weight: 500; color: #e2e8f0; margin-bottom: 6px; font-size: 13px; }
        .card-desc { color: #94a3b8; font-size: 12px; line-height: 1.5; }
        .card-path { color: #60a5fa; font-size: 10px; font-family: monospace; margin-top: 6px; }
        .menu-item { display: flex; justify-content: space-between; align-items: flex-start; padding: 10px 12px; background: rgba(255,255,255,0.03); border-radius: 6px; margin-bottom: 6px; }
        .menu-item:hover { background: rgba(255,255,255,0.08); }
        .menu-path { color: #60a5fa; font-size: 11px; font-family: monospace; margin-bottom: 4px; }
        .menu-name { font-weight: 500; color: #e2e8f0; font-size: 13px; }
        .menu-desc { color: #94a3b8; font-size: 11px; margin-top: 4px; }
        .faq-item { border-left: 3px solid #3b82f6; padding-left: 12px; margin-bottom: 12px; }
        .faq-q { font-weight: 500; color: #e2e8f0; margin-bottom: 6px; font-size: 13px; }
        .faq-a { color: #94a3b8; font-size: 12px; line-height: 1.6; }
        .faq-category { font-size: 12px; font-weight: 600; color: #10b981; margin: 16px 0 8px 0; padding: 6px 10px; background: rgba(16,185,129,0.1); border-radius: 6px; display: inline-block; }
        .feature-row { display: grid; grid-template-columns: 1fr 2fr; gap: 8px; padding: 8px 10px; background: rgba(255,255,255,0.03); border-radius: 6px; margin-bottom: 6px; font-size: 12px; }
        .feature-row:hover { background: rgba(255,255,255,0.08); }
        .feature-name { font-weight: 500; color: #e2e8f0; }
        .feature-desc { color: #94a3b8; }
        .feature-category { font-size: 11px; font-weight: 600; color: #f59e0b; margin: 14px 0 8px 0; padding: 4px 8px; background: rgba(245,158,11,0.1); border-radius: 4px; display: inline-block; }
        .highlight { background: #fbbf24; color: #0f172a; padding: 0 2px; border-radius: 2px; }
        .hidden { display: none; }
        .repo-link { display: inline-flex; align-items: center; gap: 6px; color: #60a5fa; text-decoration: none; font-size: 12px; margin-top: 8px; }
        .repo-link:hover { text-decoration: underline; }
        .version { color: #64748b; font-size: 11px; text-align: center; margin-top: 20px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1); }
        .result-count { font-size: 11px; color: #64748b; margin: 8px 0; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="search-box">
          <div style="position: relative;">
            <span class="material-icons search-icon">search</span>
            <input type="text" class="search-input" id="searchInput" placeholder="Search features, help topics, menus, or FAQ..." oninput="filterContent()">
          </div>
          <div class="result-count" id="resultCount"></div>
        </div>

        <div class="tabs">
          <div class="tab active" onclick="showTab('overview')">Overview</div>
          <div class="tab" onclick="showTab('features')">Features</div>
          <div class="tab" onclick="showTab('menus')">Menus</div>
          <div class="tab" onclick="showTab('faq')">FAQ</div>
          <div class="tab" onclick="showTab('shortcuts')">Tips</div>
        </div>

        <div id="content">
          <!-- OVERVIEW TAB -->
          <div id="overview-tab" class="tab-content">
            <div class="section">
              <div class="section-title"><span class="material-icons">dashboard</span>Two-Dashboard Architecture</div>
              <div class="card">
                <div class="card-title">Steward Dashboard (Internal)</div>
                <div class="card-desc">Comprehensive dashboard with 6 tabs: Overview, Workload, Analytics, Hot Spots, Bargaining, Satisfaction. Contains PII - for internal use only.</div>
                <div class="card-path">Strategic Ops > Command Center > Steward Dashboard</div>
              </div>
              <div class="card">
                <div class="card-title">Member Dashboard (Public)</div>
                <div class="card-desc">PII-safe dashboard for sharing with members. Shows aggregate union stats, steward directory, and satisfaction scores without personal information.</div>
                <div class="card-path">Strategic Ops > Command Center > Member Dashboard</div>
              </div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">star</span>Key Capabilities</div>
              <div class="card">
                <div class="card-title">Grievance Tracking</div>
                <div class="card-desc">Full lifecycle management from filing to resolution with automatic Article 23A deadline calculations.</div>
              </div>
              <div class="card">
                <div class="card-title">Member Directory</div>
                <div class="card-desc">31-column member records with contact info, union status, steward assignments, and auto-calculated grievance status.</div>
              </div>
              <div class="card">
                <div class="card-title">Strategic Intelligence</div>
                <div class="card-desc">Unit Hot Zones, Rising Stars, Management Hostility Reports, and Bargaining Cheat Sheets for contract negotiations.</div>
              </div>
              <div class="card">
                <div class="card-title">Integrations</div>
                <div class="card-desc">Google Calendar sync, Drive folder auto-creation, email notifications, PDF generation, and Web App deployment.</div>
              </div>
            </div>
          </div>

          <!-- FEATURES TAB (lazy-loaded) -->
          <div id="features-tab" class="tab-content hidden">
            <div class="lazy-placeholder" style="text-align:center;padding:40px 0;color:#64748b;">Loading features...</div>
          </div>
          <template id="tmpl-features">
            <div class="section">
              <div class="section-title"><span class="material-icons">apps</span>Complete Features Reference</div>

              <div class="feature-category">Dashboard & Analytics</div>
              <div class="feature-row"><div class="feature-name">Steward Dashboard</div><div class="feature-desc">Internal 6-tab dashboard with PII (Strategic Ops > Command Center)</div></div>
              <div class="feature-row"><div class="feature-name">Member Dashboard</div><div class="feature-desc">PII-safe public dashboard with aggregate stats</div></div>
              <div class="feature-row"><div class="feature-name">Steward Performance</div><div class="feature-desc">View win rates, active cases, total cases for all stewards</div></div>
              <div class="feature-row"><div class="feature-name">Satisfaction Analysis</div><div class="feature-desc">8-section survey analysis with trends and breakdowns</div></div>

              <div class="feature-category">Search & Discovery</div>
              <div class="feature-row"><div class="feature-name">Desktop Search</div><div class="feature-desc">Advanced search with filters by status/department/date</div></div>
              <div class="feature-row"><div class="feature-name">Quick Search</div><div class="feature-desc">Fast minimal interface with partial name matching</div></div>
              <div class="feature-row"><div class="feature-name">Mobile Search</div><div class="feature-desc">Touch-optimized search for field use</div></div>
              <div class="feature-row"><div class="feature-name">Live Steward Search</div><div class="feature-desc">Real-time client-side filtering in Member Dashboard</div></div>

              <div class="feature-category">Grievance Management</div>
              <div class="feature-row"><div class="feature-name">New Case/Grievance</div><div class="feature-desc">Pre-filled form with auto-calculated deadlines</div></div>
              <div class="feature-row"><div class="feature-name">Advance Step</div><div class="feature-desc">Move grievance through Step I → II → III → Arbitration</div></div>
              <div class="feature-row"><div class="feature-name">Auto Deadlines</div><div class="feature-desc">Article 23A calculations (7/14/10/21/30 day rules)</div></div>
              <div class="feature-row"><div class="feature-name">Message Alert Flag</div><div class="feature-desc">Highlight urgent cases in yellow, move to top</div></div>
              <div class="feature-row"><div class="feature-name">Bulk Status Update</div><div class="feature-desc">Update multiple grievances at once</div></div>

              <div class="feature-category">Member Management</div>
              <div class="feature-row"><div class="feature-name">Add/Find Members</div><div class="feature-desc">Registration form and search functionality</div></div>
              <div class="feature-row"><div class="feature-name">Generate Member IDs</div><div class="feature-desc">Auto-create IDs (M + First2 + Last2 + 3 digits)</div></div>
              <div class="feature-row"><div class="feature-name">Check Duplicates</div><div class="feature-desc">Find and highlight duplicate Member IDs</div></div>
              <div class="feature-row"><div class="feature-name">Import/Export</div><div class="feature-desc">Bulk data import and CSV export</div></div>

              <div class="feature-category">Steward Tools</div>
              <div class="feature-row"><div class="feature-name">Promote/Demote</div><div class="feature-desc">One-click steward status changes with toolkit emails</div></div>
              <div class="feature-row"><div class="feature-name">Workload Report</div><div class="feature-desc">Capacity analysis, overload detection (8+ cases)</div></div>
              <div class="feature-row"><div class="feature-name">Rising Stars</div><div class="feature-desc">Top performers by score and win rate</div></div>
              <div class="feature-row"><div class="feature-name">Steward Directory</div><div class="feature-desc">Contact info for all stewards</div></div>

              <div class="feature-category">Calendar & Drive</div>
              <div class="feature-row"><div class="feature-name">Sync Deadlines</div><div class="feature-desc">Create Google Calendar events for all deadlines</div></div>
              <div class="feature-row"><div class="feature-name">Drive Folders</div><div class="feature-desc">Auto-create folders with step subfolders</div></div>
              <div class="feature-row"><div class="feature-name">Batch Create</div><div class="feature-desc">Create folders for multiple grievances</div></div>

              <div class="feature-category">Strategic Intelligence</div>
              <div class="feature-row"><div class="feature-name">Unit Hot Zones</div><div class="feature-desc">Locations with 3+ active grievances</div></div>
              <div class="feature-row"><div class="feature-name">Hostility Report</div><div class="feature-desc">Management denial rate analysis</div></div>
              <div class="feature-row"><div class="feature-name">Bargaining Sheet</div><div class="feature-desc">Strategic contract negotiation data</div></div>
              <div class="feature-row"><div class="feature-name">Treemap</div><div class="feature-desc">Visual heat map of grievance activity</div></div>
              <div class="feature-row"><div class="feature-name">Sentiment Trends</div><div class="feature-desc">Union morale tracking over time</div></div>

              <div class="feature-category">Accessibility</div>
              <div class="feature-row"><div class="feature-name">Focus Mode</div><div class="feature-desc">Distraction-free view, hides non-essential sheets</div></div>
              <div class="feature-row"><div class="feature-name">Zebra Stripes</div><div class="feature-desc">Alternating row colors for readability</div></div>
              <div class="feature-row"><div class="feature-name">Dark Mode</div><div class="feature-desc">Dark gradient backgrounds on all modals</div></div>
              <div class="feature-row"><div class="feature-name">High Contrast</div><div class="feature-desc">Enhanced visibility for accessibility</div></div>

              <div class="feature-category">Security & Audit</div>
              <div class="feature-row"><div class="feature-name">Audit Logging</div><div class="feature-desc">Track all changes with timestamps and users</div></div>
              <div class="feature-row"><div class="feature-name">Sabotage Protection</div><div class="feature-desc">Mass deletion detection (>15 cells) with alerts</div></div>
              <div class="feature-row"><div class="feature-name">PII Scrubbing</div><div class="feature-desc">Auto-redact phone/SSN from public dashboards</div></div>
              <div class="feature-row"><div class="feature-name">Weingarten Rights</div><div class="feature-desc">Emergency rights statement utility</div></div>

              <div class="feature-category">Administration</div>
              <div class="feature-row"><div class="feature-name">System Diagnostics</div><div class="feature-desc">Comprehensive health check on all components</div></div>
              <div class="feature-row"><div class="feature-name">Repair Dashboard</div><div class="feature-desc">Auto-fix missing sheets, broken formulas</div></div>
              <div class="feature-row"><div class="feature-name">Midnight Trigger</div><div class="feature-desc">Daily 12AM refresh and overdue alerts</div></div>
              <div class="feature-row"><div class="feature-name">Hidden Sheets</div><div class="feature-desc">6 auto-calculating sheets for formulas</div></div>

              <div class="feature-category">Mobile & Web</div>
              <div class="feature-row"><div class="feature-name">Pocket View</div><div class="feature-desc">Hide columns for phone access</div></div>
              <div class="feature-row"><div class="feature-name">Web App</div><div class="feature-desc">Standalone deployment with URL routing</div></div>
              <div class="feature-row"><div class="feature-name">Member Portal</div><div class="feature-desc">Personalized view via ?member=ID URL</div></div>
              <div class="feature-row"><div class="feature-name">Email Links</div><div class="feature-desc">Send portal URLs to members</div></div>

              <div class="feature-category">Documents</div>
              <div class="feature-row"><div class="feature-name">PDF Generation</div><div class="feature-desc">Signature-ready PDFs with legal blocks</div></div>
              <div class="feature-row"><div class="feature-name">Email PDFs</div><div class="feature-desc">Send generated documents via email</div></div>
            </div>

            <div style="margin-top: 12px; padding: 10px; background: rgba(59,130,246,0.1); border-radius: 8px; font-size: 11px; color: #94a3b8;">
              <strong style="color: #60a5fa;">Tip:</strong> For a printable reference, go to Admin > Setup > Create Features Reference Sheet
            </div>
          </template>

          <!-- MENU REFERENCE TAB (lazy-loaded) -->
          <div id="menus-tab" class="tab-content hidden">
            <div class="lazy-placeholder" style="text-align:center;padding:40px 0;color:#64748b;">Loading menus...</div>
          </div>
          <template id="tmpl-menus">
            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>Union Hub Menu</div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Dashboards</div><div class="menu-name">Member Dashboard</div><div class="menu-desc">Public-safe dashboard with aggregate stats</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Dashboards</div><div class="menu-name">Steward Dashboard</div><div class="menu-desc">Internal dashboard with full analytics</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Search</div><div class="menu-name">Quick/Desktop/Advanced Search</div><div class="menu-desc">Multiple search interfaces</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Cases</div><div class="menu-name">New Case, Edit, Checklist</div><div class="menu-desc">Grievance management</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Members</div><div class="menu-name">Add, Find, Import, Export</div><div class="menu-desc">Member directory operations</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Calendar</div><div class="menu-name">Sync, View, Clear</div><div class="menu-desc">Google Calendar integration</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Union Hub > Drive</div><div class="menu-name">Setup, View, Batch Create</div><div class="menu-desc">Google Drive folder management</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>Strategic Ops Menu</div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > Command Center</div><div class="menu-name">Dashboards & Performance</div><div class="menu-desc">Both dashboards and steward performance</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > Cases</div><div class="menu-name">New, Edit, Checklist</div><div class="menu-desc">Grievance operations</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > ID Engines</div><div class="menu-name">Generate IDs, Check Duplicates, PDF</div><div class="menu-desc">ID and document generation</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Strategic Ops > Steward Mgmt</div><div class="menu-name">Promote, Demote, Forms, Surveys</div><div class="menu-desc">Steward management tools</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>Admin Menu</div>
              <div class="menu-item"><div><div class="menu-path">Admin</div><div class="menu-name">Diagnostics, Repair, Settings</div><div class="menu-desc">System health and fixes</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Automation</div><div class="menu-name">Refresh, Triggers, Email</div><div class="menu-desc">Automated tasks</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Data Sync</div><div class="menu-name">Sync All, Triggers</div><div class="menu-desc">Data synchronization</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Validation</div><div class="menu-name">Run, Settings, Indicators</div><div class="menu-desc">Data validation</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Setup</div><div class="menu-name">Hidden Sheets, Validations, Features</div><div class="menu-desc">System setup including Features Reference Sheet</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Admin > Demo Data</div><div class="menu-name">Seed, NUKE</div><div class="menu-desc">Test data management (dev only)</div></div></div>
            </div>

            <div class="section">
              <div class="section-title"><span class="material-icons">menu</span>Field Portal Menu</div>
              <div class="menu-item"><div><div class="menu-path">Field Portal > Accessibility</div><div class="menu-name">Mobile View, Get URL</div><div class="menu-desc">Mobile optimization</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Field Portal > Analytics</div><div class="menu-name">Unit Health, Trends, Precedents</div><div class="menu-desc">Field analytics</div></div></div>
              <div class="menu-item"><div><div class="menu-path">Field Portal > Web App</div><div class="menu-name">Deploy, Portals, Email Links</div><div class="menu-desc">Web app management</div></div></div>
            </div>
          </template>

          <!-- FAQ TAB (lazy-loaded) -->
          <div id="faq-tab" class="tab-content hidden">
            <div class="lazy-placeholder" style="text-align:center;padding:40px 0;color:#64748b;">Loading FAQ...</div>
          </div>
          <template id="tmpl-faq">
            <div class="section">
              <div class="section-title"><span class="material-icons">help</span>Frequently Asked Questions</div>

              <div class="faq-category">Getting Started</div>

              <div class="faq-item">
                <div class="faq-q">How do I set up the dashboard for the first time?</div>
                <div class="faq-a">Go to <strong>Admin > System Diagnostics</strong> to check your system, then customize the <strong>Config</strong> tab with your organization's dropdown values (job titles, locations, stewards, etc.).</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Can I use this with existing member data?</div>
                <div class="faq-a">Yes! You can paste member data into the <strong>Member Directory</strong> tab. Just make sure the columns match and Member IDs follow the format (MJOHN123).</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I test the system without real data?</div>
                <div class="faq-a">Use <strong>Admin > Demo Data > Seed All Sample Data</strong> to generate 1,000 test members and 300 grievances. Use <strong>NUKE SEEDED DATA</strong> when done testing.</div>
              </div>

              <div class="faq-category">Member Directory</div>

              <div class="faq-item">
                <div class="faq-q">What format should Member IDs use?</div>
                <div class="faq-a">Format is <strong>M + first 2 letters of first name + first 2 letters of last name + 3 random digits</strong>. Example: John Smith = MJOSM123</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Why are some columns not editable (AB-AD)?</div>
                <div class="faq-a">Columns AB-AD are <strong>auto-calculated</strong> from the Grievance Log: Has Open Grievance, Grievance Status, and Days to Deadline update automatically when you edit grievances.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I assign a steward to multiple members?</div>
                <div class="faq-a">Use the <strong>Assigned Steward</strong> dropdown in column P. You can use the multi-select editor from <strong>Union Hub > Multi-Select</strong>.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">What does the "Start Grievance" checkbox do?</div>
                <div class="faq-a">Checking this opens a <strong>pre-filled grievance form</strong> for that member. The checkbox auto-resets after use.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How does multi-select work in the Member Directory?</div>
                <div class="faq-a">Five columns support multi-select: <strong>Office Days, Preferred Communication, Best Time to Contact, Committees, and Assigned Steward(s)</strong>. Click a cell in any of these columns and use <strong>Tools &gt; Multi-Select &gt; Open Editor</strong> to open a checkbox dialog. Select multiple values and click Save. Values are stored as comma-separated text.</div>
              </div>

              <div class="faq-category">Grievances</div>

              <div class="faq-item">
                <div class="faq-q">How do I file a new grievance?</div>
                <div class="faq-a">Go to <strong>Strategic Ops > Cases > New Case/Grievance</strong>. Fill in the member info, select the grievance type, and provide details. Deadlines are calculated automatically.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How are deadlines calculated?</div>
                <div class="faq-a">Based on <strong>Article 23A</strong>: Filing = Incident + 21 days, Step I = Filed + 30 days, Step II Appeal = Step I Decision + 10 days, Step II Decision = Appeal + 30 days. Arbitration within 30 days of Step 3.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">What does "Message Alert" do?</div>
                <div class="faq-a">When checked, the row is <strong>highlighted yellow</strong> and moves to the top of the list when sorted. Use it to flag urgent cases.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Why does Days to Deadline show "Overdue"?</div>
                <div class="faq-a">This means the next deadline has <strong>passed</strong>. Check the Next Action Due column to see which deadline is overdue and take action.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I create a folder for grievance documents?</div>
                <div class="faq-a">Select the grievance row, then go to <strong>Union Hub > Drive > Setup Folder</strong>. This creates a Google Drive folder with subfolders for each step.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I select multiple articles violated or issue categories?</div>
                <div class="faq-a">The <strong>Articles Violated</strong> (column V) and <strong>Issue Category</strong> (column W) columns support multi-select. Click the cell and use <strong>Tools &gt; Multi-Select &gt; Open Editor</strong> to pick multiple values from a checkbox dialog. Options come from the Config sheet and values are stored as comma-separated text.</div>
              </div>

              <div class="faq-category">Troubleshooting</div>

              <div class="faq-item">
                <div class="faq-q">Dropdowns are empty or not working</div>
                <div class="faq-a">Check the <strong>Config</strong> tab - the corresponding column may be empty. Run <strong>Admin > Setup > Setup Data Validations</strong> to reapply dropdowns.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Data isn't syncing between sheets</div>
                <div class="faq-a">Run <strong>Admin > Data Sync > Install Auto-Sync Trigger</strong>. Also try <strong>Admin > Data Sync > Sync All Data Now</strong> for immediate sync.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">The dashboard shows wrong numbers</div>
                <div class="faq-a">Try <strong>Admin > Repair Dashboard</strong>. If issues persist, run <strong>Admin > Setup > Repair All Hidden Sheets</strong> to rebuild calculation sheets.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">I accidentally deleted data - can I undo?</div>
                <div class="faq-a">Use <strong>Ctrl+Z</strong> (or Cmd+Z on Mac) immediately. For older changes, go to <strong>File > Version history > See version history</strong>.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Menus are not appearing</div>
                <div class="faq-a">Close and reopen the spreadsheet. If still missing, go to <strong>Extensions > Apps Script</strong> and run the <code>onOpen</code> function manually.</div>
              </div>

              <div class="faq-category">Advanced</div>

              <div class="faq-item">
                <div class="faq-q">What's the difference between the two dashboards?</div>
                <div class="faq-a">The <strong>Steward Dashboard</strong> is for internal use with full PII and 6 analytical tabs. The <strong>Member Dashboard</strong> is PII-safe for sharing with members, showing only aggregate statistics.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">Can multiple people use this at the same time?</div>
                <div class="faq-a">Yes! Google Sheets supports <strong>real-time collaboration</strong>. Changes sync automatically between users.</div>
              </div>

              <div class="faq-item">
                <div class="faq-q">How do I customize deadline days?</div>
                <div class="faq-a">The default deadlines (21, 30, 10 days) are configured in the <strong>Config</strong> tab columns AA-AD. You can modify these values for your contract.</div>
              </div>
            </div>

            <a href="' + escapeHtml(getConfigValue_(CONFIG_COLS.ORG_WEBSITE) || '#') + '" target="_blank" class="repo-link">
              <span class="material-icons" style="font-size: 16px;">open_in_new</span>
              Organization Website
            </a>
          </template>

          <!-- SHORTCUTS TAB (lazy-loaded) -->
          <div id="shortcuts-tab" class="tab-content hidden">
            <div class="lazy-placeholder" style="text-align:center;padding:40px 0;color:#64748b;">Loading tips...</div>
          </div>
          <template id="tmpl-shortcuts">
            <div class="section">
              <div class="section-title"><span class="material-icons">bolt</span>Quick Tips</div>

              <div class="card">
                <div class="card-title">Quick Search</div>
                <div class="card-desc">Use <strong>Union Hub > Quick Search</strong> to find any member or grievance instantly. Supports partial name matching.</div>
              </div>

              <div class="card">
                <div class="card-title">Dashboard Views</div>
                <div class="card-desc">Use <strong>Steward Dashboard</strong> for internal analysis, <strong>Member Dashboard</strong> for sharing with members (PII-safe).</div>
              </div>

              <div class="card">
                <div class="card-title">Mobile Access</div>
                <div class="card-desc">Enable <strong>Pocket/Mobile View</strong> to hide non-essential columns when using on phones or tablets.</div>
              </div>

              <div class="card">
                <div class="card-title">Data Sync</div>
                <div class="card-desc">Run <strong>Admin > Data Sync > Sync All Data Now</strong> periodically to ensure member and grievance data stays linked.</div>
              </div>

              <div class="card">
                <div class="card-title">Visual Styling</div>
                <div class="card-desc"><strong>Visual Control Panel</strong> (Union Hub) lets you toggle dark mode, zebra stripes, and focus mode.</div>
              </div>

              <div class="card">
                <div class="card-title">Auto-Triggers</div>
                <div class="card-desc">Install <strong>auto-sync</strong> and <strong>midnight refresh</strong> triggers from Admin menu to keep data current automatically.</div>
              </div>

              <div class="card">
                <div class="card-title">Features Reference Sheet</div>
                <div class="card-desc">Create a printable features sheet via <strong>Admin > Setup > Create Features Reference Sheet</strong>.</div>
              </div>

              <div class="card">
                <div class="card-title">Keyboard Shortcut</div>
                <div class="card-desc">Use <strong>Ctrl+F</strong> (Cmd+F on Mac) in any sheet to search within the spreadsheet itself.</div>
              </div>
            </div>
          </template>
        </div>

        <div class="version">
          Dashboard v${VERSION_INFO.CURRENT} (${VERSION_INFO.BUILD_DATE}) | ${VERSION_INFO.CODENAME}
        </div>
      </div>

      <script>
        // Lazy-load: track which tabs have been rendered
        var loadedTabs = { overview: true }; // overview loads eagerly

        function showTab(tabName) {
          _ensureTabLoaded(tabName);
          document.querySelectorAll('.tab-content').forEach(t => t.classList.add('hidden'));
          document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
          document.getElementById(tabName + '-tab').classList.remove('hidden');
          event.target.classList.add('active');
          document.getElementById('resultCount').textContent = '';
        }

        function _ensureTabLoaded(tabName) {
          if (loadedTabs[tabName]) return;
          var container = document.getElementById(tabName + '-tab');
          if (!container) return;
          var placeholder = container.querySelector('.lazy-placeholder');
          if (placeholder) {
            var tmpl = document.getElementById('tmpl-' + tabName);
            if (tmpl) {
              container.innerHTML = tmpl.innerHTML;
              loadedTabs[tabName] = true;
            }
          }
        }

        function _ensureAllTabsLoaded() {
          ['features', 'menus', 'faq', 'shortcuts'].forEach(_ensureTabLoaded);
        }

        function filterContent() {
          const query = document.getElementById('searchInput').value.toLowerCase().trim();

          if (query === '') {
            const allItems = document.querySelectorAll('.card, .menu-item, .faq-item, .feature-row');
            allItems.forEach(item => {
              item.classList.remove('hidden');
              item.style.borderLeft = '';
            });
            document.getElementById('resultCount').textContent = '';
            return;
          }

          // Load all tabs before searching across them
          _ensureAllTabsLoaded();

          // Show all tabs when searching
          document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('hidden'));
          document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));

          const allItems = document.querySelectorAll('.card, .menu-item, .faq-item, .feature-row');
          let visibleCount = 0;
          allItems.forEach(item => {
            const text = item.textContent.toLowerCase();
            if (text.includes(query)) {
              item.classList.remove('hidden');
              item.style.borderLeft = '3px solid #fbbf24';
              visibleCount++;
            } else {
              item.classList.add('hidden');
              item.style.borderLeft = '';
            }
          });

          document.getElementById('resultCount').textContent = visibleCount + ' results found';
        }
      </script>
    </body>
    </html>
  `).setWidth(700).setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, '📖 Help & Features Guide - Dashboard v' + VERSION_INFO.CURRENT + ' (' + VERSION_INFO.BUILD_DATE + ')');
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
    version: VERSION_INFO.CURRENT,
    date: VERSION_INFO.BUILD_DATE,
    codename: VERSION_INFO.CODENAME,
    architecture: 'Modular Multi-File',
    history: VERSION_HISTORY,
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
      if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return errorResponse('Grievance not found');
    }

    // Update each provided field (use canonical GRIEVANCE_COLS, 1-indexed)
    if (updates.description !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLS.ISSUE_CATEGORY).setValue(escapeForFormula(updates.description));
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLS.RESOLUTION).setValue(escapeForFormula(updates.notes));
    }
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLS.STATUS).setValue(escapeForFormula(updates.status));
    }

    // Update timestamp
    sheet.getRange(rowIndex, GRIEVANCE_COLS.LAST_UPDATED).setValue(new Date());

    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_UPDATED, {
      grievanceId: grievanceId,
      updates: Object.keys(updates),
      updatedBy: Session.getActiveUser().getEmail()
    });

    return { success: true, message: 'Grievance updated successfully' };

  } catch (error) {
    console.error('Error updating grievance:', error);
    return errorResponse(error.message);
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
      return errorResponse('Member Directory sheet not found');
    }

    // Generate unique member ID using timestamp to avoid collisions from row deletions
    const timestamp = Date.now().toString(36).toUpperCase();
    const newId = `MEM-${timestamp}`;

    // Prepare row data using MEMBER_COLS constants (1-indexed)
    var rowData = new Array(MEMBER_COLS.STATE).fill('');
    rowData[MEMBER_COLS.MEMBER_ID - 1] = newId;
    rowData[MEMBER_COLS.FIRST_NAME - 1] = memberData.firstName || '';
    rowData[MEMBER_COLS.LAST_NAME - 1] = memberData.lastName || '';
    rowData[MEMBER_COLS.JOB_TITLE - 1] = memberData.jobTitle || '';
    rowData[MEMBER_COLS.WORK_LOCATION - 1] = memberData.workLocation || '';
    rowData[MEMBER_COLS.UNIT - 1] = memberData.unit || '';
    rowData[MEMBER_COLS.EMAIL - 1] = memberData.email || '';
    rowData[MEMBER_COLS.PHONE - 1] = memberData.phone || '';
    rowData[MEMBER_COLS.IS_STEWARD - 1] = 'No';
    rowData[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = new Date();
    rowData[MEMBER_COLS.EMPLOYEE_ID - 1] = memberData.employeeId || '';
    rowData[MEMBER_COLS.DEPARTMENT - 1] = memberData.department || '';
    rowData[MEMBER_COLS.HIRE_DATE - 1] = memberData.hireDate ? new Date(memberData.hireDate) : '';

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
    return errorResponse(error.message);
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
  const memberId = data[MEMBER_COLS.MEMBER_ID - 1];
  const memberName = `${data[MEMBER_COLS.FIRST_NAME - 1]} ${data[MEMBER_COLS.LAST_NAME - 1]}`;

  // Open new grievance dialog pre-populated with member info
  // Sanitize values before embedding in script context
  const html = HtmlService.createHtmlOutput(
    '<script>' +
      'sessionStorage.setItem("prefillMemberId", ' + JSON.stringify(String(memberId || '')) + ');' +
      'sessionStorage.setItem("prefillMemberName", ' + JSON.stringify(String(memberName || '')) + ');' +
      'google.script.host.close();' +
      'google.script.run.showNewGrievanceDialog();' +
    '</script>'
  ).setWidth(100).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(html, 'Loading...');
}

// ============================================================================
// IMPORT/EXPORT DIALOGS (Added to fix missing menu functions)
// ============================================================================

/**
 * Shows import members dialog
 * Allows importing members from CSV/Excel format
 */
function showImportDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      ${getMobileOptimizedHead()}
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 25px; border-radius: 8px; max-width: 500px; margin: 0 auto; }
        h2 { color: #1a73e8; margin-top: 0; }
        .info { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .format-section { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0; }
        .format-section h4 { margin-top: 0; color: #333; }
        .format-section code { background: #e0e0e0; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
        textarea { width: 100%; height: 150px; margin: 10px 0; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-family: monospace; font-size: 12px; }
        button { padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin: 5px; }
        .primary { background: #1a73e8; color: white; }
        .secondary { background: #e0e0e0; color: #333; }
        .result { margin-top: 15px; padding: 10px; border-radius: 4px; display: none; }
        .success { background: #d4edda; color: #155724; }
        .error { background: #f8d7da; color: #721c24; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📥 Import Members</h2>
        <div class="info">
          💡 Paste member data in CSV format. Each line represents one member.
        </div>
        <div class="format-section">
          <h4>Expected Format (comma-separated):</h4>
          <code>First Name, Last Name, Employee ID, Department, Job Title, Email, Phone</code>
        </div>
        <textarea id="importData" placeholder="John,Doe,EMP001,Engineering,Developer,john@example.com,555-1234
Jane,Smith,EMP002,HR,Manager,jane@example.com,555-5678"></textarea>
        <div>
          <button class="primary" onclick="importMembers()">📥 Import</button>
          <button class="secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
        <div id="result" class="result"></div>
      </div>
      <script>
        function importMembers() {
          var data = document.getElementById('importData').value.trim();
          if (!data) {
            showResult('Please paste some data to import', 'error');
            return;
          }
          google.script.run
            .withSuccessHandler(function(result) {
              showResult('Successfully imported ' + result.count + ' members', 'success');
            })
            .withFailureHandler(function(e) {
              showResult('Error: ' + e.message, 'error');
            })
            .importMembersFromText(data);
        }
        function showResult(msg, type) {
          var el = document.getElementById('result');
          el.textContent = msg;
          el.className = 'result ' + type;
          el.style.display = 'block';
        }
      </script>
    </body>
    </html>
  `).setWidth(550).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '📥 Import Members');
}

/**
 * Imports members from text (CSV format)
 * @param {string} text - CSV formatted text
 * @returns {Object} Result with count of imported members
 */
function importMembersFromText(text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  const lines = text.split('\n').filter(line => line.trim());
  let imported = 0;

  for (const line of lines) {
    const parts = line.split(',').map(p => p.trim());
    if (parts.length >= 2) {
      const newRow = [
        '',                    // ID (will be auto-generated)
        parts[0] || '',        // First Name
        parts[1] || '',        // Last Name
        parts[2] || '',        // Employee ID
        parts[3] || '',        // Department
        parts[4] || '',        // Job Title
        '',                    // Hire Date
        parts[5] || '',        // Email
        parts[6] || ''         // Phone
      ];
      sheet.appendRow(newRow);
      imported++;
    }
  }

  // Generate IDs for imported members
  if (typeof generateMissingMemberIDs === 'function') {
    generateMissingMemberIDs();
  }

  return { success: true, count: imported };
}

/**
 * Shows export directory dialog
 * Allows exporting member directory to various formats
 */
function showExportDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      ${getMobileOptimizedHead()}
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 25px; border-radius: 8px; max-width: 450px; margin: 0 auto; }
        h2 { color: #1a73e8; margin-top: 0; }
        .info { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .option { display: flex; align-items: center; padding: 15px; margin: 10px 0; background: #f8f9fa; border-radius: 8px; cursor: pointer; border: 2px solid transparent; }
        .option:hover { background: #e8f0fe; }
        .option.selected { border-color: #1a73e8; background: #e8f0fe; }
        .option-icon { font-size: 32px; margin-right: 15px; }
        .option-text h4 { margin: 0 0 5px 0; }
        .option-text p { margin: 0; font-size: 12px; color: #666; }
        button { padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin: 5px; }
        .primary { background: #1a73e8; color: white; }
        .secondary { background: #e0e0e0; color: #333; }
        .button-row { margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📤 Export Directory</h2>
        <div class="info">
          💡 Export member directory data to your preferred format.
        </div>
        <div class="option selected" onclick="selectOption(this, 'csv')">
          <div class="option-icon">📄</div>
          <div class="option-text">
            <h4>CSV File</h4>
            <p>Comma-separated values for Excel/Sheets</p>
          </div>
        </div>
        <div class="option" onclick="selectOption(this, 'email')">
          <div class="option-icon">📧</div>
          <div class="option-text">
            <h4>Email Report</h4>
            <p>Send formatted report to your email</p>
          </div>
        </div>
        <div class="option" onclick="selectOption(this, 'print')">
          <div class="option-icon">🖨️</div>
          <div class="option-text">
            <h4>Print View</h4>
            <p>Open print-friendly view in browser</p>
          </div>
        </div>
        <div class="button-row">
          <button class="primary" onclick="doExport()">📤 Export</button>
          <button class="secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
      </div>
      <script>
        var selectedFormat = 'csv';
        function selectOption(el, format) {
          document.querySelectorAll('.option').forEach(function(o) { o.classList.remove('selected'); });
          el.classList.add('selected');
          selectedFormat = format;
        }
        function doExport() {
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.url) {
                window.open(result.url, '_blank');
              }
              alert(result.message || 'Export complete!');
              google.script.host.close();
            })
            .withFailureHandler(function(e) {
              alert('Error: ' + e.message);
            })
            .exportMemberDirectory(selectedFormat);
        }
      </script>
    </body>
    </html>
  `).setWidth(500).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '📤 Export Directory');
}

/**
 * Exports member directory to specified format
 * @param {string} format - Export format (csv, email, print)
 * @returns {Object} Result with success message
 */
function exportMemberDirectory(format) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  // Exclude PII and sensitive columns (PIN_HASH, STREET_ADDRESS, CITY, STATE)
  var allData = sheet.getDataRange().getValues();
  var excludeCols = PII_MEMBER_COLS.concat([MEMBER_COLS.PIN_HASH]);
  const data = allData.map(function(row) {
    return row.filter(function(_, colIdx) {
      return excludeCols.indexOf(colIdx + 1) === -1;
    });
  });

  switch (format) {
    case 'csv':
      // Create CSV content
      const csv = data.map(row => row.map(cell => {
        var val = String(cell === null || cell === undefined ? '' : cell);
        if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
          return '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      }).join(',')).join('\n');
      const blob = Utilities.newBlob(csv, 'text/csv', 'MemberDirectory.csv');
      const file = DriveApp.createFile(blob);
      return {
        success: true,
        message: 'CSV file created! Opening...',
        url: file.getUrl()
      };

    case 'email':
      // Send email with summary
      const email = Session.getActiveUser().getEmail();
      const subject = 'Member Directory Export - ' + new Date().toLocaleDateString();
      let body = 'Member Directory Export\n\n';
      body += 'Total Members: ' + (data.length - 1) + '\n\n';
      body += 'Spreadsheet: ' + ss.getUrl();
      GmailApp.sendEmail(email, subject, body);
      return { success: true, message: 'Report sent to ' + email };

    case 'print':
      // Return URL to spreadsheet for printing
      return {
        success: true,
        message: 'Opening print view...',
        url: ss.getUrl() + '#gid=' + sheet.getSheetId()
      };

    default:
      throw new Error('Unknown export format: ' + format);
  }
}

/**
 * Shows a search dialog to find existing members
 * Menu wrapper - calls the backend findExistingMember() with user input
 */
function showFindMemberDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      ${getMobileOptimizedHead()}
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { background: white; padding: 25px; border-radius: 8px; max-width: 450px; margin: 0 auto; }
        h2 { color: #1a73e8; margin-top: 0; }
        .info { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 15px; font-size: 13px; }
        .field { margin: 15px 0; }
        .field label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
        .field input { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
        button { padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin: 5px; }
        .primary { background: #1a73e8; color: white; }
        .secondary { background: #e0e0e0; color: #333; }
        .button-row { margin-top: 20px; }
        .results { margin-top: 15px; max-height: 200px; overflow-y: auto; }
        .result-item { padding: 10px; background: #f8f9fa; margin: 5px 0; border-radius: 4px; cursor: pointer; }
        .result-item:hover { background: #e8f0fe; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🔍 Find Member</h2>
        <div class="info">
          💡 Search by name, email, or member ID to find existing members.
        </div>
        <div class="field">
          <label>Search Term</label>
          <input type="text" id="searchTerm" placeholder="Enter name, email, or member ID..." autofocus>
        </div>
        <div class="button-row">
          <button class="primary" onclick="searchMembers()">🔍 Search</button>
          <button class="secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
        <div id="results" class="results"></div>
      </div>
      <script>
        ${getClientSideEscapeHtml()}
        document.getElementById('searchTerm').addEventListener('keypress', function(e) {
          if (e.key === 'Enter') searchMembers();
        });
        function searchMembers() {
          var term = document.getElementById('searchTerm').value.trim();
          if (!term) { alert('Please enter a search term'); return; }
          google.script.run
            .withSuccessHandler(function(results) {
              var html = '';
              if (results.length === 0) {
                html = '<div class="result-item">No members found</div>';
              } else {
                results.forEach(function(m) {
                  html += '<div class="result-item" onclick="goToMember(' + m.row + ')">' +
                    '<strong>' + escapeHtml(m.name) + '</strong><br>' +
                    '<small>' + escapeHtml(m.id || 'No ID') + ' • ' + escapeHtml(m.email || 'No email') + '</small>' +
                    '</div>';
                });
              }
              document.getElementById('results').innerHTML = html;
            })
            .withFailureHandler(function(e) {
              document.getElementById('results').innerHTML = '<div class="result-item">Error: ' + escapeHtml(e.message) + '</div>';
            })
            .searchMembersForDialog(term);
        }
        function goToMember(row) {
          google.script.run.navigateToMemberRow(row);
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `).setWidth(500).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Find Member');
}

/**
 * Searches members for the find dialog
 * @param {string} term - Search term
 * @returns {Array} Array of matching members
 */
function searchMembersForDialog(term) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (!sheet) {
    throw new Error('Member Directory sheet not found');
  }

  const data = sheet.getDataRange().getValues();
  const results = [];
  const searchLower = term.toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const firstName = (row[MEMBER_COLS.FIRST_NAME - 1] || '').toString().toLowerCase();
    const lastName = (row[MEMBER_COLS.LAST_NAME - 1] || '').toString().toLowerCase();
    const email = (row[MEMBER_COLS.EMAIL - 1] || '').toString().toLowerCase();
    const memberId = (row[MEMBER_COLS.MEMBER_ID - 1] || '').toString().toLowerCase();

    if (firstName.includes(searchLower) ||
        lastName.includes(searchLower) ||
        email.includes(searchLower) ||
        memberId.includes(searchLower)) {
      results.push({
        row: i + 1,
        name: row[MEMBER_COLS.FIRST_NAME - 1] + ' ' + row[MEMBER_COLS.LAST_NAME - 1],
        email: row[MEMBER_COLS.EMAIL - 1],
        id: row[MEMBER_COLS.MEMBER_ID - 1]
      });

      if (results.length >= 10) break; // Limit results
    }
  }

  return results;
}

/**
 * Navigates to a specific member row
 * @param {number} row - Row number to navigate to
 */
function navigateToMemberRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);

  if (sheet) {
    sheet.activate();
    sheet.getRange(row, 1).activate();
    SpreadsheetApp.flush();
  }
}
