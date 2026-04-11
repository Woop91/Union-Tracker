/**
 * ============================================================================
 * 10_Main.gs — THE Entry Point & Triggers
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   THE entry point for the entire application. Contains the three main GAS
 *   trigger functions:
 *     1. onOpen()       — Simple trigger that creates the menu bar when the
 *                         spreadsheet opens. Clears caches so fresh config is loaded.
 *     2. onEdit(e)      — Simple trigger that fires on cell edits. Dispatches
 *                         to EventBus for decoupled handling.
 *     3. dailyTrigger() — Installable trigger that runs nightly maintenance
 *                         (deadline alerts, data sync, security digest).
 *   Also contains handleSecurityAudit_() for edit-time audit logging.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   onOpen is a GAS "simple trigger" which runs with restricted authorization —
 *   it CANNOT call ScriptApp, MailApp, GmailApp, or any OAuth service (violations
 *   cause silent failure). That's why onOpen ONLY creates menus and clears caches.
 *   All deferred work runs via installable triggers. The prior onOpen approach used
 *   a self-deleting deferred trigger which had a race condition (the finally block
 *   deleted the trigger before it fired). Fix: deferred work is now in standalone
 *   installable triggers set up via Admin menu.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   If onOpen fails: no menus appear, users can't access any features through
 *   the UI. If onEdit fails: real-time updates (status changes, auto-formatting)
 *   stop. If dailyTrigger fails: deadline alerts, email notifications, and
 *   security digests stop sending.
 *
 * DEPENDENCIES:
 *   Depends on: 03_UIComponents.gs (createDashboardMenu), 15_EventBus.gs
 *               (EventBus.emit), DevMenu.gs (buildDevMenu, optional)
 *   Used by:    GAS runtime — called automatically by Google Apps Script
 *
 * @fileoverview Main entry point and trigger functions
 */

// ============================================================================
// TRIGGER FUNCTIONS
// ============================================================================

/**
 * Runs when the spreadsheet is opened
 * Sets up the custom menu, applies tab colors, and initializes the dashboard
 */
function onOpen() {
  // FIX v4.25.7: onOpen is a GAS simple trigger — ScriptApp.getProjectTriggers()
  // requires authorization not available in simple triggers and throws silently.
  // The prior approach also had a race condition: the finally block called
  // cleanUpOnOpenTrigger_() synchronously, deleting the deferred trigger before
  // the 1000ms elapsed, so onOpenDeferred_() never ran.
  //
  // Fix: onOpen does ONLY menu creation. onOpenDeferred_ is installed once as a
  // standalone installable trigger via Admin → Triggers → Install All Survey Triggers
  // (menuInstallSurveyTriggers), which is the correct GAS pattern.
  try {
    // Clear memoized caches so fresh config values are picked up
    if (typeof _systemNameCache_ !== 'undefined') _systemNameCache_ = null;
    createDashboardMenu();
    if (typeof buildDevMenu === 'function') buildDevMenu();
  } catch (error) {
    log_('onOpen', 'Error: ' + (error.message || error));
  }
}


/**
 * Deferred onOpen tasks — runs heavy init after the UI is responsive.
 * Installed as a persistent installable onOpen trigger via setupOpenDeferredTrigger().
 * Has full authorization (can show modals, access ScriptApp, etc.).
 * @private
 */
function onOpenDeferred_() {
  // FIX v4.50.7: Removed self-deletion code that was left over from the old
  // one-shot timed trigger approach. This function is now installed as a
  // PERSISTENT installable onOpen trigger via setupOpenDeferredTrigger().
  // Deleting the trigger here meant it only fired on the first open after
  // install — all subsequent opens skipped deferred init entirely.

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    log_('onOpenDeferred_', 'getActiveSpreadsheet() returned null');
    return;
  }

  // REL-04: Wrap entire deferred init in try/catch so trigger failures are logged
  // and don't silently disappear (GAS only logs trigger errors to Stackdriver).
  try {
    var _columnsChanged = false;

    try {
      var syncResult = syncColumnMaps();
      if (syncResult && syncResult.synced && syncResult.synced.length > 0) {
        _columnsChanged = true;
        log_('onOpenDeferred_', 'columns shifted: ' + syncResult.synced.join(', '));
      }
    } catch (syncError) {
      log_('onOpenDeferred_', 'Column sync skipped: ' + syncError.message);
    }

    try {
      ensureAllSheetColumns_();
    } catch (colError) {
      log_('onOpenDeferred_', 'Column check skipped: ' + colError.message);
    }

    // When columns were backfilled or shifted, re-apply ALL validations so
    // dropdowns/checkboxes land on the correct (freshly-resolved) columns.
    // Without this, stale validations persist on old column positions until
    // the user manually runs a repair — the gap that caused phantom dropdowns
    // on Member ID, Contact Notes, PIN Hash, etc.
    if (_columnsChanged) {
      try {
        setupDataValidations();
        log_('onOpenDeferred_', 'validations re-applied after column shift');
      } catch (valError) {
        log_('onOpenDeferred_', 'Validation re-apply skipped: ' + valError.message);
      }
    }

    try {
      loadUnitCodes_();
    } catch (unitCodeErr) {
      log_('onOpenDeferred_', 'Unit codes load skipped: ' + unitCodeErr.message);
    }

    try {
      if (typeof migrateSheetTabTitles_ === 'function') {
        migrateSheetTabTitles_(ss);
      }
    } catch (migrateError) {
      log_('Sheet tab title migration skipped', migrateError.message);
    }

    try {
      if (typeof applyTabColors_ === 'function') {
        applyTabColors_(ss);
      }
    } catch (tabError) {
      log_('Tab colors not applied', tabError.message);
    }

    try {
      enforceHiddenSheets();
    } catch (hideError) {
      log_('Hidden sheet enforcement skipped', hideError.message);
    }

    // EventBus subscribers are NOT registered here — GAS runs each trigger
    // invocation in an isolated context, so onEdit() re-registers on every call.
    // Registering here wastes ~200ms on spreadsheet open for no benefit.

    ss.toast('Dashboard loaded successfully', '\uD83C\uDFDB\uFE0F Union Dashboard', 3);

    // FIX v4.50.7: Auto-show tab modal for the active sheet on spreadsheet open.
    // This works because onOpenDeferred_ is an INSTALLABLE trigger with full
    // authorization — unlike onSelectionChange (simple trigger) which cannot call
    // showModalDialog(). Respects ENABLE_TAB_MODALS config and per-user dismissals.
    try {
      if (typeof showCurrentTabModal === 'function' &&
          (typeof isTabModalsEnabled_ !== 'function' || isTabModalsEnabled_())) {
        var activeSheet = ss.getActiveSheet().getName();
        // Check per-user dismissal before showing
        var sheetKey = '';
        if (typeof SHEETS !== 'undefined') {
          for (var sk in SHEETS) {
            if (SHEETS[sk] === activeSheet) { sheetKey = sk; break; }
          }
        }
        if (!sheetKey || typeof isTabModalDismissed_ !== 'function' || !isTabModalDismissed_(sheetKey)) {
          showCurrentTabModal();
        }
      }
    } catch (modalErr) {
      log_('Auto-open tab modal skipped', modalErr.message);
    }
  } catch (deferredErr) {
    log_('onOpenDeferred_ failed', deferredErr.message + '\n' + deferredErr.stack);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('DEFERRED_INIT_FAILED', 'onOpenDeferred_ error: ' + deferredErr.message);
    }
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
    try { loadCachedColumnMaps_(); } catch (_cacheErr) { log_('onEdit', 'Error: ' + (_cacheErr.message || _cacheErr)); }

    // Ensure EventBus subscribers are registered for this execution context.
    // GAS runs each trigger invocation in an isolated context — global state
    // from onOpenDeferred_ doesn't persist into onEdit executions.
    try {
      if (typeof registerEventBusSubscribers === 'function') registerEventBusSubscribers();
    } catch (_busErr) { log_('onEdit', 'EventBus registration error: ' + (_busErr.message || _busErr)); }

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
        log_('onEdit', 'MultiSelect handler error: ' + multiSelectError.message);
      }
    }

    // ========================================
    // 3. Sheet-Specific Handlers via EventBus
    // ========================================
    // All sheet-specific handlers (grievance edit, member edit, checklist,
    // volunteer hours, meeting attendance, config, auto-sync, audit) are
    // registered as EventBus subscribers in registerEventBusSubscribers()
    // (15_EventBus.gs). Priority ordering ensures correct execution sequence.
    if (typeof emitEditEvent === 'function') {
      emitEditEvent(e);
    }

  } catch (error) {
    log_('onEdit', 'Error: ' + (error.message || error));
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
    var props = PropertiesService.getUserProperties();

    // ── Tab Modal auto-open (v4.48.0) ──
    // Detect sheet tab switches and show contextual modals.
    var currentSheet = e.range.getSheet().getName();
    var lastSheet = props.getProperty('tabModal_lastSheet');
    if (currentSheet !== lastSheet) {
      props.setProperty('tabModal_lastSheet', currentSheet);
      // Only fire on actual tab change (not initial load or same-tab clicks)
      if (lastSheet !== null) {
        onTabSwitch_(currentSheet);
      }
    }

    // ── Multi-select auto-open ──
    // Cache the preference to avoid reading PropertiesService on every selection change
    var userCache = CacheService.getUserCache();
    var autoOpen = userCache ? userCache.get('autoOpenPref') : null;
    if (autoOpen === null) {
      autoOpen = props.getProperty('multiSelectAutoOpen');
      if (userCache) try { userCache.put('autoOpenPref', autoOpen || '', 300); } catch(_) {}
    }
    // Default is ON — only skip if user has explicitly disabled it ('false').
    // Absence of the property (fresh install) means enabled.
    if (autoOpen === 'false') return;

    if (typeof onSelectionChangeMultiSelect === 'function') {
      onSelectionChangeMultiSelect(e);
    }
  } catch (_err) { log_('onSelectionChange', 'Error: ' + (_err.message || _err)); }
}

/**
 * Called when the user switches to a different sheet tab.
 * Shows a toast hint about available quick actions for this tab.
 *
 * DESIGN NOTE (v4.50.6): GAS simple triggers (including onSelectionChange)
 * CANNOT call showModalDialog() — the call fails silently. The prior approach
 * (v4.48.0) tried to open modals directly from onSelectionChange, which never
 * worked. Fix: show a toast hint and let the user open the modal via menu
 * (🔧 Tools > 📑 Tab Quick Actions) or (🔧 Tools > 📑 Tab Modals > ...).
 *
 * @param {string} sheetName - Name of the newly-activated sheet
 * @private
 */
function onTabSwitch_(sheetName) {
  try {
    if (typeof TAB_MODAL_REGISTRY === 'undefined') return;
    if (typeof isTabModalsEnabled_ === 'function' && !isTabModalsEnabled_()) return;

    for (var i = 0; i < TAB_MODAL_REGISTRY.length; i++) {
      var entry = TAB_MODAL_REGISTRY[i];
      if (entry.sheet === sheetName) {
        // Check per-user dismissal
        var sheetKey = '';
        for (var k in SHEETS) {
          if (SHEETS[k] === sheetName) { sheetKey = k; break; }
        }
        if (sheetKey && typeof isTabModalDismissed_ === 'function' && isTabModalDismissed_(sheetKey)) return;

        // Show a toast hint — toasts work from simple triggers, modals don't.
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Quick actions available for ' + entry.title + '. Use 🔧 Tools > 📑 Tab Quick Actions to open.',
          '📑 ' + entry.title, 4);
        return;
      }
    }
  } catch (_e) { log_('onTabSwitch_ error', (_e.message || _e)); }
}

/**
 * Shows the tab modal for the currently active sheet.
 * Auto-detects which sheet the user is on and opens the corresponding
 * contextual modal with quick actions, tips, and links.
 * Accessible via menu: 🔧 Tools > 📑 Tab Quick Actions
 */
function showCurrentTabModal() {
  try {
    // T5-1: Respect ENABLE_TAB_MODALS toggle for manual menu path too
    if (typeof isTabModalsEnabled_ === 'function' && !isTabModalsEnabled_()) {
      SpreadsheetApp.getUi().alert('Tab modals are currently disabled. Enable them in Admin Settings.');
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheetName = ss.getActiveSheet().getName();

    if (typeof TAB_MODAL_REGISTRY === 'undefined') {
      SpreadsheetApp.getUi().alert('Tab modal registry not available.');
      return;
    }

    var fnMap = {
      showTabModalConfig: typeof showTabModalConfig === 'function' ? showTabModalConfig : null,
      showTabModalMemberDirectory: typeof showTabModalMemberDirectory === 'function' ? showTabModalMemberDirectory : null,
      showTabModalGrievanceLog: typeof showTabModalGrievanceLog === 'function' ? showTabModalGrievanceLog : null,
      showTabModalCaseChecklist: typeof showTabModalCaseChecklist === 'function' ? showTabModalCaseChecklist : null,
      showTabModalVolunteerHours: typeof showTabModalVolunteerHours === 'function' ? showTabModalVolunteerHours : null,
      showTabModalMeetingAttendance: typeof showTabModalMeetingAttendance === 'function' ? showTabModalMeetingAttendance : null,
      showTabModalMeetingCheckIn: typeof showTabModalMeetingCheckIn === 'function' ? showTabModalMeetingCheckIn : null,
      showTabModalResources: typeof showTabModalResources === 'function' ? showTabModalResources : null
    };

    for (var i = 0; i < TAB_MODAL_REGISTRY.length; i++) {
      var entry = TAB_MODAL_REGISTRY[i];
      if (entry.sheet === sheetName && fnMap[entry.fn]) {
        fnMap[entry.fn]();
        return;
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'No quick actions modal available for the "' + sheetName + '" tab.',
      'Tab Quick Actions', 3);
  } catch (e) {
    log_('showCurrentTabModal error', e.message);
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
      // Delegate to canonical schema setup in 08d_AuditAndFormulas.gs
      if (typeof setupAuditLogSheet === 'function') {
        setupAuditLogSheet();
        auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
      }
      if (!auditSheet) {
        // Fallback: create with correct headers
        auditSheet = ss.insertSheet(SHEETS.AUDIT_LOG);
        auditSheet.appendRow(['Timestamp', 'Event Type', 'User', 'Details', 'Session ID']);
        auditSheet.setFrozenRows(1);
        setSheetVeryHidden_(auditSheet);
      }
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
    // Skip isBlank check for very large edits — too expensive for simple trigger context
    var rangeIsEmpty = false;
    if (numCells <= 500) {
      rangeIsEmpty = range.isBlank();
    } else if (!e.value) {
      // For large ranges, sample the first cell as a heuristic
      try { rangeIsEmpty = range.getCell(1, 1).isBlank(); } catch(_) { rangeIsEmpty = false; }
    }
    if (numCells > 15 && !e.value && rangeIsEmpty) {
      alertMessage = 'MASS_DELETION_ALERT';

      // L-41: MailApp.sendEmail() is unavailable in simple trigger context (onEdit).
      // Immediate email send is impossible here. Fix (v4.33.1): explicitly queue the
      // event for the daily security digest so admins receive email within 24 hours.
      // recordSecurityEvent(CRITICAL) calls sendSecurityAlertEmail_() which silently
      // fails in this context — so we ALSO queue directly here.
      var deletionDetails = {
        email: userEmail,
        sheet: range.getSheet().getName(),
        range: range.getA1Notation(),
        cellsAffected: numCells
      };

      // Log to console for visibility (PII-safe)
      secureLog('SabotageDetection', 'Mass deletion detected', deletionDetails);

      // Queue for 24h digest via PropertiesService (works in simple trigger context)
      if (typeof queueSecurityDigestEvent_ === 'function') {
        queueSecurityDigestEvent_('MASS_DELETION', 'Mass deletion of ' + numCells + ' cells in ' + range.getSheet().getName(), deletionDetails);
      }

      // Record through centralized security event system (audit log + attempt immediate email)
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('MASS_DELETION', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.CRITICAL : 'CRITICAL',
          'Mass deletion of ' + numCells + ' cells detected in ' + range.getSheet().getName(),
          deletionDetails);
      }
    }

    // Detect large-scale changes (not necessarily malicious but notable)
    if (numCells > 50) {
      alertMessage = alertMessage || 'LARGE_CHANGE';
    }

    // v4.55.1 D10-BUG-01: route through logAuditEvent() so entries are protected by the
    // HMAC chain. Previously the direct 5-column appendRow bypassed the chain and left
    // entries without integrity hashes. Falls back to direct append only if logAuditEvent
    // is unavailable at load time (shouldn't happen in prod since 08d_ loads before 10_).
    var auditDetails = {
      cell: range.getA1Notation(),
      sheet: range.getSheet().getName(),
      oldValue: e.oldValue || '(empty)',
      newValue: e.value || '(deleted)',
      alert: alertMessage || '',
      user: userEmail
    };
    if (typeof logAuditEvent === 'function') {
      logAuditEvent(alertMessage || 'CELL_EDIT', auditDetails);
    } else {
      var auditRow = [
        new Date(),
        alertMessage || 'CELL_EDIT',
        userEmail,
        JSON.stringify(auditDetails),
        ''
      ];
      auditSheet.appendRow(auditRow);
    }

  } catch (auditError) {
    // Silently fail - don't break user's edit for audit logging
    log_('Audit log error', auditError.message);
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
    log_('Auto-style error', styleError.message);
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
        // CR-12: Only set DATE_CLOSED if the cell is currently empty,
        // preserving any manually entered close date.
        var existingCloseDate = sheet.getRange(row, GRIEVANCE_COLS.DATE_CLOSED).getValue();
        if (!existingCloseDate) {
          sheet.getRange(row, GRIEVANCE_COLS.DATE_CLOSED).setValue(new Date());
        }
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
    log_('Workflow error', workflowError.message);
  }
}

/**
 * Sends escalation alert email to Chief Steward.
 * Reads email dynamically from CONFIG_COLS.CHIEF_STEWARD_EMAIL — the actual
 * column letter depends on the active header order and is not fixed.
 * @param {string} memberName - Name of the member
 * @param {string} caseID - Grievance case ID
 * @param {string} status - New status/step
 * @private
 */
function sendEscalationAlert_(memberName, caseID, status) {
  var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);

  if (!chiefStewardEmail) {
    log_('sendEscalationAlert_', 'Chief Steward email not configured (CONFIG_COLS.CHIEF_STEWARD_EMAIL) - skipping escalation alert');
    return;
  }

  try {
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Case Escalation Alert';
    var body = 'ESCALATION NOTICE\n\n' +
               'Case ' + caseID + ' for ' + memberName + ' has been escalated to: ' + status + '\n\n' +
               'Immediate review is required.\n' +
               COMMAND_CONFIG.EMAIL.FOOTER;

    // L-41: MailApp is unavailable in simple trigger context (onEdit).
    // Try direct send first; if it fails, queue for daily trigger flush.
    try {
      safeSendEmail_({ to: chiefStewardEmail, subject: subject, body: body });
      SpreadsheetApp.getActiveSpreadsheet().toast('Escalation alert sent to Chief Steward', 'Alert Sent', 3);
    } catch (_sendErr) {
      _queueEscalationEmail(chiefStewardEmail, subject, body);
      log_('sendEscalationAlert_', 'Escalation email queued for daily flush (simple trigger context)');
    }
  } catch (emailError) {
    log_('Escalation email error', emailError.message);
  }
}

/**
 * Queues an escalation email for later flush by a daily time-driven trigger.
 * Used when MailApp is unavailable (e.g., simple trigger context).
 * @param {string} to - Recipient email
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @private
 */
function _queueEscalationEmail(to, subject, body) {
  try {
    var props = PropertiesService.getScriptProperties();
    var queue = JSON.parse(props.getProperty('ESCALATION_EMAIL_QUEUE') || '[]');
    queue.push({ to: to, subject: subject, body: body, queued: new Date().toISOString() });
    props.setProperty('ESCALATION_EMAIL_QUEUE', JSON.stringify(queue));
  } catch(_) {}
}

/**
 * Flushes queued escalation emails. Call from a daily time-driven trigger.
 * @returns {number} Number of emails sent
 */
function flushEscalationEmailQueue() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('ESCALATION_EMAIL_QUEUE');
  if (!raw) return 0;
  var queue = [];
  try { queue = JSON.parse(raw); } catch(_) { return 0; }
  if (!queue.length) return 0;

  var sent = 0;
  var failed = [];
  for (var i = 0; i < queue.length; i++) {
    try {
      safeSendEmail_({ to: queue[i].to, subject: queue[i].subject, body: queue[i].body });
      sent++;
    } catch(e) {
      log_('flushEscalationEmailQueue', 'Send failed: ' + e.message);
      failed.push(queue[i]);
    }
  }
  if (failed.length > 0) {
    props.setProperty('ESCALATION_EMAIL_QUEUE', JSON.stringify(failed));
  } else {
    props.deleteProperty('ESCALATION_EMAIL_QUEUE');
  }
  return sent;
}

/**
 * Gets a value from the Config sheet by column number
 * Reads from row 3 (first data row after headers)
 * @param {number} columnNum - Column number (1-indexed)
 * @returns {string} The value from the Config sheet, or empty string if not found
 * @private
 */
function getConfigValue_(columnNum, fallback) {
  var fb = arguments.length >= 2 ? fallback : '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!configSheet) {
      log_('getConfigValue_', 'Config sheet not found');
      return fb;
    }

    // Config values are typically in row 3 (row 1 = section headers, row 2 = column headers)
    var value = configSheet.getRange(3, columnNum).getValue();
    return value != null && value !== '' ? String(value).trim() : fb;
  } catch (e) {
    log_('getConfigValue_', 'Error reading config value col ' + columnNum + ': ' + e.message);
    return fb;
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
    log_('getEscalationStatuses_', 'Error reading escalation statuses: ' + e.message);
  }

  // Fall back to defaults from COMMAND_CONFIG
  return COMMAND_CONFIG.ESCALATION_STATUSES || ['In Arbitration', 'Appealed'];
}

/**
 * Loads unit codes from Config sheet into COMMAND_CONFIG.UNIT_CODES.
 * Expected format per cell: "Unit Name:CODE" (e.g., "Main Station:MS").
 * Falls back silently if no data or parsing fails.
 * @private
 */
function loadUnitCodes_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return;

    var values = getConfigValues(configSheet, CONFIG_COLS.UNIT_CODES);
    if (values.length === 0) return;

    var codes = {};
    for (var i = 0; i < values.length; i++) {
      var parts = String(values[i]).split(':');
      if (parts.length >= 2) {
        var name = parts[0].trim();
        var code = parts.slice(1).join(':').trim();
        if (name && code) codes[name] = code;
      }
    }
    if (Object.keys(codes).length > 0) {
      COMMAND_CONFIG.UNIT_CODES = codes;
    }
  } catch (e) {
    log_('loadUnitCodes_', 'Error loading unit codes: ' + e.message);
  }
}

/**
 * Validates Config sheet edits and warns the user via toast if the value
 * doesn't match the column's expected data type.
 * Non-blocking — warns only, does not reject the edit.
 * @param {Object} e - The edit event object
 * @private
 */
function warnInvalidConfigValue_(e) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 3) return;

  for (var i = 0; i < CONFIG_HEADER_MAP_.length; i++) {
    if (CONFIG_COLS[CONFIG_HEADER_MAP_[i].key] === col) {
      var entry = CONFIG_HEADER_MAP_[i];
      var value = e.range.getValue();
      if (!value || String(value).trim() === '') return;
      var result = validateConfigValue_(entry.key, value);
      if (!result.valid) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          '"' + entry.header + '" expects ' + result.type + ' — got: ' + String(value).substring(0, 30),
          '\u26A0\uFE0F Config Warning', 8
        );
      }
      return;
    }
  }
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
    log_('getEscalationSteps_', 'Error reading escalation steps: ' + e.message);
  }

  // Fall back to defaults from COMMAND_CONFIG
  return COMMAND_CONFIG.ESCALATION_STEPS || ['Step II', 'Step III', 'Arbitration'];
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

  // Batch-read the entire row once to avoid multiple getValue() calls
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

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
    var currentStep = col_(rowData, GRIEVANCE_COLS.CURRENT_STEP);
    var status = col_(rowData, GRIEVANCE_COLS.STATUS);
    var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];

    if (closedStatuses.indexOf(status) === -1 && currentStep) {
      var nextActionDate = '';
      if (currentStep === 'Informal') {
        nextActionDate = col_(rowData, GRIEVANCE_COLS.FILING_DEADLINE);
      } else if (currentStep === 'Step I') {
        nextActionDate = col_(rowData, GRIEVANCE_COLS.STEP1_DUE);
      } else if (currentStep === 'Step II') {
        nextActionDate = col_(rowData, GRIEVANCE_COLS.STEP2_DUE);
      } else if (currentStep === 'Step III') {
        nextActionDate = col_(rowData, GRIEVANCE_COLS.STEP3_APPEAL_DUE);
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
      const grievanceId = col_(rowData, GRIEVANCE_COLS.GRIEVANCE_ID);

      // Auto-sync to calendar if enabled
      if (settings.autoSyncCalendar && grievanceId && typeof syncSingleGrievanceToCalendar === 'function') {
        syncSingleGrievanceToCalendar(grievanceId);
      }
    } catch (settingsError) {
      log_('Settings error in handleGrievanceEdit', settingsError.message);
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
  // H-36: Skip downstream recalculation if the edited column IS itself a deadline
  // column — that would overwrite the steward's manual edit.
  const stepDateColumns = [
    GRIEVANCE_COLS.DATE_FILED,      // Step 1 filing
    GRIEVANCE_COLS.STEP2_APPEAL_FILED,  // Step 2 appeal
    GRIEVANCE_COLS.STEP3_APPEAL_FILED   // Step 3 appeal
  ];

  if (stepDateColumns.includes(col) && calculatedDateCols.indexOf(col) === -1) {
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

  // FIX-MAIN-01: v4.25.8 — Read deadline days from getDeadlineRules() (Config sheet).
  // v4.51.2: Uses addCalendarDays for response/filing deadlines (Art. 23 calendar days),
  // addBusinessDays only for appeal deadlines (Art. 23 business days).
  var rules = getDeadlineRules();

  var deadlineChain = [
    { calcCol: GRIEVANCE_COLS.FILING_DEADLINE,  sourceCol: GRIEVANCE_COLS.INCIDENT_DATE,       days: rules.FILING_DAYS,                 business: false },
    { calcCol: GRIEVANCE_COLS.STEP1_DUE,        sourceCol: GRIEVANCE_COLS.DATE_FILED,           days: rules.STEP_1.DAYS_FOR_RESPONSE,    business: false },
    { calcCol: GRIEVANCE_COLS.STEP2_APPEAL_DUE, sourceCol: GRIEVANCE_COLS.STEP1_RCVD,          days: rules.STEP_2.DAYS_TO_APPEAL,       business: true },
    { calcCol: GRIEVANCE_COLS.STEP2_DUE,        sourceCol: GRIEVANCE_COLS.STEP2_APPEAL_FILED,  days: rules.STEP_2.DAYS_FOR_RESPONSE,    business: false },
    { calcCol: GRIEVANCE_COLS.STEP3_APPEAL_DUE, sourceCol: GRIEVANCE_COLS.STEP2_RCVD,          days: rules.STEP_3.DAYS_TO_APPEAL,       business: true }
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

  // M-32: Actually cascade downstream — for each deadline after the overridden one,
  // check if its source date is empty. If so, use the previous calculated date
  // as the basis and write the new downstream deadline.
  var previousDate = overrideDate;
  for (var d = startIdx + 1; d < deadlineChain.length; d++) {
    var downstream = deadlineChain[d];
    var sourceValue = sheet.getRange(row, downstream.sourceCol).getValue();

    // If the downstream step has its own source date, it will be calculated from that
    // source by syncGrievanceFormulasToLog — skip it.
    if (sourceValue instanceof Date) {
      var _addDays = downstream.business ? addBusinessDays : addCalendarDays;
      previousDate = _addDays(sourceValue, downstream.days);
      continue;
    }

    // No source date entered — cascade from the previous calculated date
    var downstreamNote = sheet.getRange(row, downstream.calcCol).getNote();
    // Don't overwrite a steward override on a downstream column
    if (downstreamNote === 'Steward override') {
      previousDate = sheet.getRange(row, downstream.calcCol).getValue();
      if (previousDate instanceof Date) continue;
    }

    var _addDaysFn = downstream.business ? addBusinessDays : addCalendarDays;
    var newDeadline = _addDaysFn(previousDate, downstream.days);
    sheet.getRange(row, downstream.calcCol).setValue(newDeadline);
    previousDate = newDeadline;
  }

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
    try { newValue = e.range.getValue(); } catch (_) { log_('syncDropdownToConfig_', 'Error: ' + (_.message || _)); }
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
  var addedNew = false;
  for (var v = 0; v < valuesToSync.length; v++) {
    var val = valuesToSync[v].trim();
    if (val && !existingSet[val]) {
      addToConfigDropdown_(configCol, val);
      existingSet[val] = true;
      addedNew = true;
    }
  }

  // Refresh data validation on the source column so the new Config value
  // appears in the dropdown immediately — requireValueInList embeds values,
  // not a live reference, so stale rules won't show newly-added options.
  if (addedNew && configCol && configSheet) {
    try {
      if (isMultiSelect && typeof setMultiSelectValidation === 'function') {
        setMultiSelectValidation(e.range.getSheet(), col, configSheet, configCol);
      } else {
        setDropdownValidation(e.range.getSheet(), col, configSheet, configCol);
      }
    } catch (_refreshErr) {
      log_('syncDropdownToConfig_ validation refresh', _refreshErr.message);
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
      var firstName = (col_(memberData[i], MEMBER_COLS.FIRST_NAME) || '').toString().trim();
      var lastName = (col_(memberData[i], MEMBER_COLS.LAST_NAME) || '').toString().trim();
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
      var fn = (col_(memberData[j], MEMBER_COLS.FIRST_NAME) || '').toString().trim();
      var ln = (col_(memberData[j], MEMBER_COLS.LAST_NAME) || '').toString().trim();
      var full = fn + ' ' + ln;
      if (full === newValue) {
        var currentStatus = col_(memberData[j], MEMBER_COLS.IS_STEWARD);
        if (!isTruthyValue(currentStatus)) {
          memberSheet.getRange(j + 1, MEMBER_COLS.IS_STEWARD).setValue('Yes');
        }
        break;
      }
    }
  }
}

/**
 * Handles edits to Config dropdown columns — refreshes the corresponding
 * data validation in Member Directory and/or Grievance Log so dropdowns
 * stay in sync automatically (Config → Sheet direction).
 * @param {Object} e - The edit event object
 * @private
 */
function syncConfigToSheetValidation_(e) {
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 3) return; // header rows, not data

  // Check if this Config column is a dropdown source for any sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = e.range.getSheet();
  var refreshed = false;

  // Wrap validation refresh in try/catch — if a Config column exceeds 500 items,
  // the validation call must not crash the onEdit handler.
  try {
    // Member Directory dropdowns
    var memberDD = DROPDOWN_MAP.MEMBER_DIR;
    for (var m = 0; m < memberDD.length; m++) {
      if (memberDD[m].configCol === col) {
        var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
        if (memberSheet) setDropdownValidation(memberSheet, memberDD[m].col, configSheet, col);
        refreshed = true;
        break;
      }
    }

    // Grievance Log dropdowns
    if (!refreshed) {
      var grievDD = DROPDOWN_MAP.GRIEVANCE_LOG;
      for (var g = 0; g < grievDD.length; g++) {
        if (grievDD[g].configCol === col) {
          var grievSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
          if (grievSheet) setDropdownValidation(grievSheet, grievDD[g].col, configSheet, col);
          break;
        }
      }
    }

    // Multi-select columns (Member Directory)
    var memberMS = MULTI_SELECT_COLS.MEMBER_DIR;
    for (var mm = 0; mm < memberMS.length; mm++) {
      if (memberMS[mm].configCol === col) {
        var msSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
        if (msSheet && typeof setMultiSelectValidation === 'function') {
          setMultiSelectValidation(msSheet, memberMS[mm].col, configSheet, col);
        }
        break;
      }
    }

    // Multi-select columns (Grievance Log)
    var grievMS = MULTI_SELECT_COLS.GRIEVANCE_LOG;
    for (var gm = 0; gm < grievMS.length; gm++) {
      if (grievMS[gm].configCol === col) {
        var gmsSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
        if (gmsSheet && typeof setMultiSelectValidation === 'function') {
          setMultiSelectValidation(gmsSheet, grievMS[gm].col, configSheet, col);
        }
        break;
      }
    }
  } catch (_validationErr) {
    log_('syncConfigToSheetValidation_', 'validation refresh failed: ' + (_validationErr.message || _validationErr));
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
      log_('Meeting status update error', e);
    }

    // Process meeting doc notifications (agenda 3 days before, notes 1 day before, publish 1 day after)
    var meetingDocResult = { agendaSent: 0, notesSent: 0, notesPublished: 0 };
    try {
      if (typeof processMeetingDocNotifications === 'function') {
        meetingDocResult = processMeetingDocNotifications();
      }
    } catch (e) {
      log_('Meeting doc notification error', e);
    }

    // Cleanup expired meeting check-in records (>90 days old)
    var meetingRowsCleaned = 0;
    try {
      meetingRowsCleaned = cleanupExpiredMeetings();
    } catch (e) {
      log_('Meeting cleanup error', e);
    }

    // v4.36.0 — Run trend alert detection
    try {
      if (typeof triggerDailyTrendDetection === 'function') {
        triggerDailyTrendDetection();
      }
    } catch (e) {
      log_('Trend detection error', e.message);
    }

    // ── v4.33.1 DAILY MAINTENANCE: previously orphaned functions now wired ──

    // Read retention thresholds from Config tab — defaults to 90 days if blank/invalid.
    // Admins can change these in the Config tab (Grievance Archive Days / Audit Log Archive Days).
    var grievanceArchiveDays = 90;
    var auditArchiveDays = 90;
    try {
      if (typeof getConfigValue_ === 'function' && typeof CONFIG_COLS !== 'undefined') {
        var gDays = parseInt(getConfigValue_(CONFIG_COLS.GRIEVANCE_ARCHIVE_DAYS), 10);
        var aDays = parseInt(getConfigValue_(CONFIG_COLS.AUDIT_ARCHIVE_DAYS), 10);
        if (!isNaN(gDays) && gDays > 0) grievanceArchiveDays = gDays;
        if (!isNaN(aDays) && aDays > 0) auditArchiveDays = aDays;
      }
    } catch (e) {
      log_('dailyTrigger', 'Could not read archive threshold from Config, using defaults: ' + e.message);
    }

    // Auto-archive closed grievances older than configured days
    // archiveClosedGrievances() existed since v4.30.0 but was never called from a trigger.
    var archiveResult = { archived: 0 };
    try {
      if (typeof archiveClosedGrievances === 'function') {
        archiveResult = archiveClosedGrievances(grievanceArchiveDays) || archiveResult;
        if (archiveResult.archived > 0) {
          log_('dailyTrigger', 'auto-archived ' + archiveResult.archived + ' closed grievances');
        }
      }
    } catch (e) {
      log_('dailyTrigger', 'Auto-archive grievances error: ' + e.message);
    }

    // Archive old audit log entries to Drive CSV
    // dailyAuditArchive() existed since v4.30.0 but was never called from a trigger.
    var auditArchiveResult = { archived: 0 };
    try {
      if (typeof archiveOldAuditLogs_ === 'function') {
        auditArchiveResult = archiveOldAuditLogs_(auditArchiveDays) || auditArchiveResult;
      }
    } catch (e) {
      log_('dailyTrigger', 'Audit log archive error: ' + e.message);
    }

    // Send daily security digest (batched HIGH/CRITICAL events from the last 24h)
    // sendDailySecurityDigest() existed since v4.8.1 but was never wired to a trigger.
    // NOTE: CRITICAL events logged by sabotage detection (onEdit simple trigger) cannot
    // send email immediately (MailApp unavailable in simple triggers). They are queued
    // here and emailed via the digest so admins are notified within 24h.
    try {
      if (typeof sendDailySecurityDigest === 'function') {
        sendDailySecurityDigest();
      }
    } catch (e) {
      log_('dailyTrigger', 'Security digest error: ' + e.message);
    }

    // Cleanup expired auth session tokens (removes tokens older than TTL)
    // authCleanupExpiredTokens() existed since v4.22.9 but was never wired.
    try {
      if (typeof authCleanupExpiredTokens === 'function') {
        authCleanupExpiredTokens();
      }
    } catch (e) {
      log_('dailyTrigger', 'Auth token cleanup error: ' + e.message);
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
      meetingRowsCleaned: meetingRowsCleaned,
      grievancesArchived: archiveResult.archived,
      auditEntriesArchived: auditArchiveResult.archived || 0
    });

  } catch (error) {
    log_('dailyTrigger', 'Error: ' + (error.message || error));
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

// ============================================================================
// HELP & DOCUMENTATION
// ============================================================================

/**
 * Shows the searchable Help Guide modal with menu breakdown and FAQ
 * v4.3.7 - Complete rewrite with search, menu reference, and FAQ
 */
function showHelpDialog() {
  const htmlContent = `
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

            <a href="${escapeHtml(getConfigValue_(CONFIG_COLS.ORG_WEBSITE) || '#')}" target="_blank" class="repo-link">
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
          Dashboard v${VERSION_INFO.CURRENT} (${VERSION_INFO.BUILD_DATE}) | ${VERSION_INFO.codename}
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
  `;
  showDialog_(htmlContent, '📖 Help & Features Guide - Dashboard v' + VERSION_INFO.CURRENT + ' (' + VERSION_INFO.BUILD_DATE + ')', 700, 750);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================
/**
 * Updates a grievance with new data
 * @param {string} grievanceId - The grievance ID
 * @param {Object} updates - Fields to update
 * @return {Object} Result object
 */
function updateGrievance(grievanceId, updates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return errorResponse('Grievance Log sheet not found');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID) === grievanceId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return errorResponse('Grievance not found');
    }

    // Update each provided field (use canonical GRIEVANCE_COLS, 1-indexed)
    if (updates.description !== undefined) {
      // Preserve newlines — grievance descriptions are multi-paragraph.
      sheet.getRange(rowIndex, GRIEVANCE_COLS.DESCRIPTION).setValue(escapeForFormulaPreserveNewlines(updates.description));
    }
    if (updates.notes !== undefined) {
      sheet.getRange(rowIndex, GRIEVANCE_COLS.RESOLUTION).setValue(escapeForFormulaPreserveNewlines(updates.notes));
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
    log_('updateGrievance', 'Error: ' + (error.message || error));
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
  const htmlContent = `
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
            `<option value="${escapeHtml(String(s))}">${escapeHtml(String(s))}</option>`
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
            .withFailureHandler(function(e) { alert('Error: ' + e.message); })
            .bulkUpdateGrievanceStatus(ids, status, notes);
        }
      </script>
    </body>
    </html>
  `;
  showDialog_(htmlContent, 'Bulk Status Update', 400, 300);
}

// ============================================================================
// MEMBER MANAGEMENT
// ============================================================================

/**
 * Shows new member dialog
 */
function showNewMemberDialog() {
  // Read Dues Status options from Config sheet (same pattern as _configList)
  var duesStatusOptions = [];
  try {
    var cfgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
    if (cfgSheet && CONFIG_COLS.DUES_STATUSES > 0) {
      var lr = cfgSheet.getLastRow();
      if (lr >= 3) {
        duesStatusOptions = cfgSheet.getRange(3, CONFIG_COLS.DUES_STATUSES, lr - 2, 1)
          .getValues().map(function(r) { return String(r[0]).trim(); })
          .filter(function(v) { return v !== ''; });
      }
    }
  } catch (_e) { /* fall through to hardcoded defaults */ }
  if (!duesStatusOptions.length) {
    duesStatusOptions = ['Current', 'Past Due', 'Inactive', 'Non Member'];
  }
  var duesStatusOptionsHtml = duesStatusOptions.map(function(v) {
    return '<option value="' + v.replace(/"/g, '&quot;') + '">' + v.replace(/</g, '&lt;') + '</option>';
  }).join('\n                ');
  const htmlContent = `
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
              <label class="form-label">Dues Status</label>
              <select class="form-select" id="duesStatus">
                ${duesStatusOptionsHtml}
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
          .withFailureHandler(function(e) { console.error('Failed to load departments:', e.message); })
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
            duesStatus: document.getElementById('duesStatus').value,
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
  `;
  showDialog_(htmlContent, 'Add New Member', 600, 500);
}

/**
 * Adds a new member to the directory
 * @param {Object} memberData - Member information
 * @return {Object} Result object
 */
function addNewMember(memberData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      return errorResponse('Member Directory sheet not found');
    }

    // Generate unique member ID using timestamp to avoid collisions from row deletions
    const timestamp = Date.now().toString(36).toUpperCase();
    const newId = `MEM-${timestamp}`;

    // M-33: Size row by the maximum column index in MEMBER_COLS (not just STATE,
    // which may not be the last column).
    var maxMemberCol = 0;
    for (var colKey in MEMBER_COLS) {
      if (Object.prototype.hasOwnProperty.call(MEMBER_COLS, colKey) && MEMBER_COLS[colKey] > maxMemberCol) {
        maxMemberCol = MEMBER_COLS[colKey];
      }
    }
    // Prepare row data using MEMBER_COLS constants (1-indexed)
    var rowData = new Array(maxMemberCol).fill('');
    setCol_(rowData, MEMBER_COLS.MEMBER_ID, newId);
    setCol_(rowData, MEMBER_COLS.FIRST_NAME, escapeForFormula(memberData.firstName || ''));
    setCol_(rowData, MEMBER_COLS.LAST_NAME, escapeForFormula(memberData.lastName || ''));
    setCol_(rowData, MEMBER_COLS.JOB_TITLE, escapeForFormula(memberData.jobTitle || ''));
    setCol_(rowData, MEMBER_COLS.WORK_LOCATION, escapeForFormula(memberData.workLocation || ''));
    setCol_(rowData, MEMBER_COLS.UNIT, escapeForFormula(memberData.unit || ''));
    setCol_(rowData, MEMBER_COLS.EMAIL, escapeForFormula(memberData.email || ''));
    setCol_(rowData, MEMBER_COLS.PHONE, escapeForFormula(memberData.phone || ''));
    setCol_(rowData, MEMBER_COLS.IS_STEWARD, 'No');
    setCol_(rowData, MEMBER_COLS.RECENT_CONTACT_DATE, new Date());
    setCol_(rowData, MEMBER_COLS.EMPLOYEE_ID, escapeForFormula(memberData.employeeId || ''));
    setCol_(rowData, MEMBER_COLS.DEPARTMENT, escapeForFormula(memberData.department || ''));
    setCol_(rowData, MEMBER_COLS.HIRE_DATE, memberData.hireDate ? new Date(memberData.hireDate) : '');
    setCol_(rowData, MEMBER_COLS.DUES_STATUS, escapeForFormula(memberData.duesStatus || ''));

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
    log_('addNewMember', 'Error: ' + (error.message || error));
    return errorResponse(error.message);
  }
}
// ============================================================================
// IMPORT/EXPORT DIALOGS (Added to fix missing menu functions)
// ============================================================================
/**
 * Exports member directory to specified format
 * @param {string} format - Export format (csv, email, print)
 * @returns {Object} Result with success message
 */
function exportMemberDirectory(format) {
  // F20-21: PII export requires steward authorization
  var authResult = checkWebAppAuthorization('steward');
  if (!authResult.isAuthorized) {
    return errorResponse(authResult.message || 'Unauthorized: steward access required for export', 'exportMemberDirectory');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

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
      // M-60: Export files can be auto-cleaned by running setupWeeklyExportCleanupTrigger()
      // from the Tools menu. See cleanupOldExportFiles() in 06_Maintenance.gs.
      return {
        success: true,
        message: 'CSV file created! Opening... (Note: remember to delete from Drive when done)',
        url: file.getUrl()
      };

    case 'email':
      // Send email with summary
      const email = Session.getActiveUser().getEmail();
      const subject = 'Member Directory Export - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
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

// Dead code removed (v4.55.2): searchMembersForDialog() and navigateToMemberRow()
// were server-side callbacks for a find-member HtmlService dialog that was
// removed in an earlier refactor. Zero callers in src/ or test/.

