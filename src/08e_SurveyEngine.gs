/**
 * ============================================================================
 * 08e_SurveyEngine.gs - Survey Period Management and Aggregation
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Native survey engine replacing Google Form integration (v4.21.0).
 *   initSurveyEngine() is a one-shot setup function that creates survey
 *   sheets, seeds config, installs triggers (quarterly auto-open, weekly
 *   reminders), and activates the current period. Manages survey periods,
 *   tracking, and aggregation.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Google Forms required external access and couldn't enforce the union's
 *   specific survey sections and branching logic. The native engine uses a
 *   schema-driven approach where survey questions are defined in the Survey
 *   Questions sheet and can be edited by the spreadsheet owner.
 *   initSurveyEngine() is idempotent — safe to re-run.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Surveys cannot be opened/closed. Survey responses are not collected.
 *   Quarterly auto-triggers don't fire. Member satisfaction data goes stale.
 *   The satisfaction dashboard shows outdated scores.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, SURVEY_QUESTIONS_COLS),
 *   10b_SurveyDocSheets.gs (createSurveyQuestionsSheet). Used by DevMenu.gs
 *   (init), daily/quarterly triggers, and the satisfaction dashboard in 09_.
 *
 * @fileoverview Survey engine setup, period management, and aggregation
 * @requires 01_Core.gs, 10b_SurveyDocSheets.gs
 */

// ── One-shot init (run once after deployment) ─────────────────────────────────

/**
 * RUN THIS ONCE after deployment.
 * Does everything needed to activate the survey engine:
 *   1. Creates _Survey_Periods hidden sheet
 *   2. Seeds Survey Priority Options into Config tab (if empty)
 *   3. Installs quarterly auto-trigger (Jan/Apr/Jul/Oct day 1 at 6 AM)
 *   4. Installs weekly reminder trigger (Tuesdays at 9 AM)
 *   5. Opens the current quarter's survey period immediately
 *   6. Resets _Survey_Tracking statuses to 'Not Completed' for new period
 *   7. Pushes 'Survey Now Open' notification to all members
 *
 * Safe to re-run — each step is idempotent.
 */
function initSurveyEngine() {
  // Auth check: only admins/stewards may initialize the survey engine
  var callerRole = typeof getUserRole_ === 'function' ? getUserRole_(Session.getActiveUser().getEmail()) : null;
  if (callerRole !== 'steward' && callerRole !== 'admin' && callerRole !== 'both') {
    SpreadsheetApp.getUi().alert('Authorization required: steward or admin access needed.');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var log = [];

  try {
    // 0. Survey Questions sheet (seed if not exists — owner edits question text here)
    createSurveyQuestionsSheet(SpreadsheetApp.getActiveSpreadsheet());
    log.push('✅ 📋 Survey Questions sheet ready');
  } catch(e) { log.push('❌ Survey Questions sheet: ' + e.message); }

  try {
    // 1. Hidden sheet
    setupSurveyPeriodsSheet();
    log.push('✅ _Survey_Periods sheet ready');
  } catch(e) { log.push('❌ _Survey_Periods: ' + e.message); }

  try {
    // 2. Seed Config Survey Priority Options (non-destructive)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && CONFIG_COLS.SURVEY_PRIORITY_OPTIONS) {
      var existing = configSheet.getLastRow() > 2
        ? configSheet.getRange(3, CONFIG_COLS.SURVEY_PRIORITY_OPTIONS, configSheet.getLastRow() - 2, 1).getValues().flat().filter(String)
        : [];
      if (existing.length === 0) {
        var defaults = [
          'Contract Enforcement','Workload & Staffing','Scheduling & Office Days',
          'Pay & Benefits','Health & Safety','Training & Development',
          'Equity & Inclusion','Communication','Steward Support & Access',
          'Member Organizing','Other'
        ];
        configSheet.getRange(3, CONFIG_COLS.SURVEY_PRIORITY_OPTIONS, defaults.length, 1)
          .setValues(defaults.map(function(v) { return [v]; }));
        log.push('✅ Survey Priority Options seeded (' + defaults.length + ' options)');
      } else {
        log.push('✅ Survey Priority Options already set (' + existing.length + ' options) — not overwritten');
      }
    } else {
      log.push('⚠️  Config sheet or SURVEY_PRIORITY_OPTIONS column not found');
    }
  } catch(e) { log.push('❌ Config seed: ' + e.message); }

  try {
    // 3. Quarterly trigger
    setupQuarterlyTrigger();
    log.push('✅ Quarterly auto-trigger installed');
  } catch(e) { log.push('❌ Quarterly trigger: ' + e.message); }

  try {
    // 4. Weekly reminder trigger
    setupWeeklyReminderTrigger();
    log.push('✅ Weekly reminder trigger installed');
  } catch(e) { log.push('❌ Weekly trigger: ' + e.message); }

  try {
    // 5–7. Open the current period (also resets tracking + sends notification)
    var activePeriod = getSurveyPeriod();
    if (activePeriod) {
      log.push('✅ Survey period already active: ' + activePeriod.periodId + ' — not reopened');
    } else {
      var result = openNewSurveyPeriod('initSurveyEngine()');
      if (result.success) {
        log.push('✅ Survey period opened: ' + result.periodId);
        log.push('✅ _Survey_Tracking reset to Not Completed');
        log.push('✅ Survey-open notification sent to members');
      } else {
        log.push('❌ Open period: ' + result.message);
      }
    }
  } catch(e) { log.push('❌ Open period: ' + e.message); }

  try {
    // 8. Populate _Survey_Tracking from Member Directory (non-destructive if responses exist)
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var trackSheet = ss2.getSheetByName(SHEETS.SURVEY_TRACKING);
    var memberSheet = ss2.getSheetByName(SHEETS.MEMBER_DIR);
    if (!trackSheet) { log.push('⚠️  _Survey_Tracking sheet not found — skipping.'); throw new Error('_Survey_Tracking sheet not found'); }
    var memberCount = memberSheet ? Math.max(0, memberSheet.getLastRow() - 1) : 0;
    if (memberCount === 0) {
      log.push('⚠️  Member Directory empty — tracking not populated. Add members first, then re-run.');
    } else {
      populateSurveyTrackingFromMembers();
      var newCount = trackSheet ? Math.max(0, trackSheet.getLastRow() - 1) : 0;
      log.push('✅ _Survey_Tracking populated: ' + newCount + ' members');
    }
  } catch(e) { log.push('❌ Tracking populate: ' + e.message); }

  var summary = log.join('\n');
  log_('initSurveyEngine', 'initSurveyEngine:\n' + summary);
  ui.alert('Survey Engine Initialized', summary, ui.ButtonSet.OK);
}
//
// RESPONSIBILITIES:
//   - Quarterly period lifecycle (open, track, close, archive)
//   - Auto-trigger quarterly period creation via time-driven trigger
//   - Archive closed periods to Drive → Past Survey Questions/[Period Name]/
//   - Pending member list for steward dashboard
//   - Section-level satisfaction summary (plain values, no Sheets formulas)
//   - Push survey-open notification to all active members
//
// DEPENDENCIES:
//   01_Core.gs    — HIDDEN_SHEETS, SURVEY_PERIODS_COLS, SATISFACTION_COLS,
//                   SURVEY_TRACKING_COLS, CONFIG_COLS, SHEETS, MEMBER_COLS
//   00_Security.gs — hashForVault_(), withScriptLock_()
//   15_EventBus.gs — EventBus.publish() for notifications (optional, guarded)
//
// ANONYMITY: This file never reads from _Survey_Vault or individual response
//   rows in a way that links any response to a person. getSatisfactionSummary()
//   reads only numeric columns and aggregates them. getPendingSurveyMembers()
//   reads only _Survey_Tracking status column — no survey answers.
// ============================================================================

// ── Hidden sheet setup ───────────────────────────────────────────────────────

/**
 * Creates or verifies the _Survey_Periods hidden sheet.
 * Safe to re-run — skips creation if sheet already exists with data.
 * Called by setupHiddenSheets() in 08a_SheetSetup.gs.
 */
function setupSurveyPeriodsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
  if (existing) {
    // Already exists — verify headers present
    if (existing.getLastRow() >= 1 &&
        String(existing.getRange(1, 1).getValue()).indexOf('Period ID') !== -1) {
      return; // Already set up
    }
  }

  var sheet = existing || ss.insertSheet(HIDDEN_SHEETS.SURVEY_PERIODS);
  sheet.clear();

  // Headers
  var headers = [
    'Period ID', 'Period Name', 'Start Date', 'End Date',
    'Status', 'Archive Folder URL', 'Created By', 'Response Count'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(SHEET_COLORS.BG_LIGHT_BLUE);

  sheet.setFrozenRows(1);

  // Hide the sheet at API level
  try { setSheetVeryHidden_(sheet); } catch(_e) { sheet.hideSheet(); }

  log_('setupSurveyPeriodsSheet', '_Survey_Periods created.');
}

// ── Period lifecycle ─────────────────────────────────────────────────────────

/**
 * Returns the currently Active survey period, or null if none open.
 * Reads from _Survey_Periods hidden sheet.
 *
 * @returns {Object|null} { periodId, name, startDate, endDate, status, responseCount }
 */
function getSurveyPeriod() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    var C = SURVEY_PERIODS_COLS;

    for (var i = data.length - 1; i >= 0; i--) {
      var status = String(data[i][C.STATUS - 1] || '').trim();
      if (status === 'Active') {
        return {
          periodId:      String(data[i][C.PERIOD_ID      - 1] || ''),
          name:          String(data[i][C.PERIOD_NAME    - 1] || ''),
          startDate:     data[i][C.START_DATE   - 1] || null,
          endDate:       data[i][C.END_DATE     - 1] || null,
          status:        'Active',
          archiveUrl:    String(data[i][C.ARCHIVE_URL   - 1] || ''),
          responseCount: parseInt(data[i][C.RESPONSE_COUNT - 1], 10) || 0,
          rowIndex:      i + 2  // 1-indexed sheet row
        };
      }
    }
    return null; // No active period
  } catch(e) {
    log_('getSurveyPeriod error', e.message);
    return null;
  }
}

/**
 * Opens a new quarterly survey period.
 * If a period is currently Active, archives it first.
 * Pushes "Survey Open" notification to all active members.
 *
 * @param {string} [callerEmail] - Email of steward opening the period (optional)
 * @returns {Object} { success, periodId, message }
 */
function openNewSurveyPeriod(callerEmail) {
  return withScriptLock_(function() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // Archive any currently active period
      var activePeriod = getSurveyPeriod();
      if (activePeriod) {
        archiveSurveyPeriod_(activePeriod.periodId);
      }

      // Generate new period ID and name
      var now       = new Date();
      var year      = now.getFullYear();
      var quarter   = Math.ceil((now.getMonth() + 1) / 3);
      var periodId  = 'Q' + quarter + '-' + year;
      var periodName= 'Q' + quarter + ' ' + year + ' Member Satisfaction Survey';

      // Quarter date bounds
      // Day 0 of month N gives the last day of month N-1, so:
      //   Q1 → new Date(y, 3, 0) = Mar 31
      //   Q2 → new Date(y, 6, 0) = Jun 30
      //   Q3 → new Date(y, 9, 0) = Sep 30
      //   Q4 → new Date(y, 12,0) = Dec 31
      var qStart = new Date(year, (quarter - 1) * 3, 1);
      var qEnd   = new Date(year,  quarter * 3,      0);

      // Write to _Survey_Periods
      var periodsSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
      if (!periodsSheet) {
        setupSurveyPeriodsSheet();
        periodsSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
      }

      var C = SURVEY_PERIODS_COLS;
      var newRow = new Array(8).fill('');
      newRow[C.PERIOD_ID      - 1] = periodId;
      newRow[C.PERIOD_NAME    - 1] = periodName;
      newRow[C.START_DATE     - 1] = qStart;
      newRow[C.END_DATE       - 1] = qEnd;
      newRow[C.STATUS         - 1] = 'Active';
      newRow[C.ARCHIVE_URL    - 1] = '';
      newRow[C.CREATED_BY     - 1] = callerEmail || 'Auto (quarterly trigger)';
      newRow[C.RESPONSE_COUNT - 1] = 0;
      periodsSheet.appendRow(newRow);

      // Reset _Survey_Tracking statuses to 'Not Completed' for new period
      startNewSurveyRound();

      // Push survey-open notification to all active members
      pushSurveyOpenNotification_(periodName);

      log_('openNewSurveyPeriod', 'Opened ' + periodId);
      return { success: true, periodId: periodId, message: periodName + ' is now open.' };

    } catch(e) {
      log_('openNewSurveyPeriod error', e.message);
      return { success: false, periodId: null, message: 'Error opening period: ' + e.message };
    }
  }, 30);
}

/**
 * Archives a survey period: exports question snapshot + all responses to Drive,
 * marks period as Closed, updates the archive URL in _Survey_Periods.
 *
 * Drive path: [Past Surveys Folder] / [Period Name] /
 *   - questions.json  — snapshot of question definitions at time of archive
 *   - responses.csv   — all response rows for this period (anonymous)
 *
 * @param {string} periodId - e.g. 'Q1-2026'
 */
function archiveSurveyPeriod_(periodId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Find period row ────────────────────────────────────────────────
    var periodsSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
    if (!periodsSheet || periodsSheet.getLastRow() < 2) return;

    var periodData = periodsSheet.getRange(2, 1, periodsSheet.getLastRow() - 1, 8).getValues();
    var C = SURVEY_PERIODS_COLS;
    var periodRow = -1;
    var periodName = periodId;
    for (var i = 0; i < periodData.length; i++) {
      if (String(periodData[i][C.PERIOD_ID - 1]) === periodId) {
        periodRow = i + 2; // 1-indexed
        periodName = String(periodData[i][C.PERIOD_NAME - 1] || periodId);
        break;
      }
    }
    if (periodRow < 0) return; // period not found

    // ── Get or create Past Surveys drive folder ────────────────────────
    var pastFolderUrl = '';
    try {
      var pastFolderId = getConfigValue_(CONFIG_COLS.PAST_SURVEYS_FOLDER_ID);
      var parentFolder;
      if (pastFolderId) {
        parentFolder = DriveApp.getFolderById(pastFolderId);
      } else {
        // Fall back to root drive folder from Config
        var rootFolderId = getConfigValue_(CONFIG_COLS.DASHBOARD_ROOT_FOLDER_ID);
        var root = rootFolderId ? DriveApp.getFolderById(rootFolderId) : DriveApp.getRootFolder();
        // Create "Past Survey Questions" folder under root
        var pastFolderIter = root.getFoldersByName('Past Survey Questions');
        parentFolder = pastFolderIter.hasNext()
          ? pastFolderIter.next()
          : root.createFolder('Past Survey Questions');
      }

      // Create period subfolder
      var periodFolderIter = parentFolder.getFoldersByName(periodName);
      var periodFolder = periodFolderIter.hasNext()
        ? periodFolderIter.next()
        : parentFolder.createFolder(periodName);
      pastFolderUrl = periodFolder.getUrl();

      // Pin the parent folder ID back to Config so future archives use it directly
      // (only writes if PAST_SURVEYS_FOLDER_ID is currently empty)
      try {
        var existingPinId = getConfigValue_(CONFIG_COLS.PAST_SURVEYS_FOLDER_ID);
        if (!existingPinId) {
          var configSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
          if (configSh && CONFIG_COLS.PAST_SURVEYS_FOLDER_ID) {
            configSh.getRange(3, CONFIG_COLS.PAST_SURVEYS_FOLDER_ID).setValue(parentFolder.getId());
            log_('archiveSurveyPeriod_', 'Pinned Past Surveys Folder ID to Config: ' + parentFolder.getId());
          }
        }
      } catch(pe) {
        log_('archiveSurveyPeriod_', 'Could not pin folder ID to Config: ' + pe.message);
      }

      // ── Export question snapshot (JSON) ──────────────────────────────
      try {
        var questionData = getSurveyQuestions();
        var questionsJson = JSON.stringify({
          exportedAt: new Date().toISOString(),
          periodId: periodId,
          periodName: periodName,
          questions: questionData.questions,
          sections: questionData.sections
        }, null, 2);
        var existingQ = periodFolder.getFilesByName('questions.json');
        if (existingQ.hasNext()) existingQ.next().setTrashed(true);
        periodFolder.createFile('questions.json', questionsJson, 'application/json');
      } catch(qe) {
        log_('archiveSurveyPeriod_', 'question export error: ' + qe.message);
      }

      // ── Export responses CSV ──────────────────────────────────────────
      try {
        var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
        if (satSheet && satSheet.getLastRow() > 1) {
          // Build CSV from header row + all data rows
          var headers = satSheet.getRange(1, 1, 1, satSheet.getLastColumn()).getValues()[0];
          var allRows = satSheet.getRange(2, 1, satSheet.getLastRow() - 1, satSheet.getLastColumn()).getValues();

          var csvLines = [headers.map(function(h) {
            return '"' + String(h).replace(/"/g, '""') + '"';
          }).join(',')];

          for (var row = 0; row < allRows.length; row++) {
            csvLines.push(allRows[row].map(function(cell) {
              return '"' + String(cell instanceof Date ? cell.toISOString() : (cell || '')).replace(/"/g, '""') + '"';
            }).join(','));
          }

          var existingR = periodFolder.getFilesByName('responses.csv');
          if (existingR.hasNext()) existingR.next().setTrashed(true);
          periodFolder.createFile('responses.csv', csvLines.join('\n'), 'text/csv');
        }
      } catch(re) {
        log_('archiveSurveyPeriod_', 'response export error: ' + re.message);
      }

    } catch(de) {
      log_('archiveSurveyPeriod_', 'Drive error: ' + de.message);
    }

    // ── Mark period as Closed in _Survey_Periods ──────────────────────
    var rowData = periodsSheet.getRange(periodRow, 1, 1, 8).getValues()[0];
    rowData[C.STATUS      - 1] = 'Closed';
    rowData[C.ARCHIVE_URL - 1] = pastFolderUrl;
    periodsSheet.getRange(periodRow, 1, 1, 8).setValues([rowData]);

    log_('archiveSurveyPeriod_', 'Archived ' + periodId + ' → ' + (pastFolderUrl || '(no Drive URL)'));

  } catch(e) {
    log_('archiveSurveyPeriod_ error', e.message);
  }
}

/**
 * Increments the response count for a period in _Survey_Periods.
 * @param {string} periodId
 */
function incrementPeriodResponseCount_(periodId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    var C = SURVEY_PERIODS_COLS;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][C.PERIOD_ID - 1]) === periodId) {
        var countCell = sheet.getRange(i + 2, C.RESPONSE_COUNT);
        countCell.setValue((parseInt(countCell.getValue(), 10) || 0) + 1);
        return;
      }
    }
  } catch(e) {
    log_('incrementPeriodResponseCount_ error', e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Quarterly auto-trigger ───────────────────────────────────────────────────

/**
 * Time-driven trigger entry point. Runs on the 1st of Jan, Apr, Jul, Oct.
 * Opens a new survey period for the current quarter.
 * Install via setupQuarterlyTrigger() once, then GAS runs it automatically.
 */
function autoTriggerQuarterlyPeriod() {
  var now     = new Date();
  var month   = now.getMonth(); // 0-indexed
  var day     = now.getDate();

  // Only open on the 1st day of quarter-start months: Jan(0), Apr(3), Jul(6), Oct(9)
  var quarterStarts = [0, 3, 6, 9];
  if (quarterStarts.indexOf(month) === -1 || day !== 1) {
    log_('autoTriggerQuarterlyPeriod', 'not a quarter start, skipping.');
    return;
  }

  var result = openNewSurveyPeriod('Auto (quarterly trigger)');
  log_('autoTriggerQuarterlyPeriod', JSON.stringify(result));
}

/**
 * Installs the monthly time-driven trigger for autoTriggerQuarterlyPeriod.
 * Safe to run multiple times — deletes existing trigger before creating new one.
 * Run once manually from the Sheets menu after deployment.
 */
function setupQuarterlyTrigger() {
  // Remove existing trigger with same function name to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoTriggerQuarterlyPeriod') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Monthly trigger — GAS will fire on the 1st. autoTriggerQuarterlyPeriod()
  // then checks whether it's a quarter-start month before acting.
  ScriptApp.newTrigger('autoTriggerQuarterlyPeriod')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)  // 6 AM
    .create();

  log_('setupQuarterlyTrigger', 'Monthly trigger installed for autoTriggerQuarterlyPeriod on day 1 at 6 AM.');
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Quarterly survey trigger installed. New periods will open automatically on Jan 1, Apr 1, Jul 1, Oct 1.', 'Survey Engine', 5);
}

// ── Steward dashboard data ───────────────────────────────────────────────────

/**
 * Returns the list of active members who have NOT yet completed
 * the current survey period. Used by steward dashboard pending widget.
 *
 * Does NOT read survey answers. Reads only _Survey_Tracking status column.
 *
 * @returns {Object} { periodId, periodName, total, pending, completed, rate, members[] }
 *   members[]: { memberId, name, email, workLocation, assignedSteward }
 */
function getPendingSurveyMembers() {
  try {
    var period = getSurveyPeriod();
    if (!period) return { periodId: null, periodName: null, total: 0, pending: 0, completed: 0, rate: 0, members: [] };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trackSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
    if (!trackSheet || trackSheet.getLastRow() < 2) {
      return { periodId: period.periodId, periodName: period.name, total: 0, pending: 0, completed: 0, rate: 0, members: [] };
    }

    var data = trackSheet.getRange(2, 1, trackSheet.getLastRow() - 1, trackSheet.getLastColumn()).getValues();
    var C = SURVEY_TRACKING_COLS;
    var pendingMembers = [];
    var completedCount = 0;

    // First pass: identify pending vs completed (no enrichment yet)
    var pendingRows = [];
    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][C.CURRENT_STATUS - 1] || '').trim().toLowerCase();
      if (status === 'completed') {
        completedCount++;
      } else {
        pendingRows.push(data[i]);
      }
    }

    // Early exit: if no pending members, skip the expensive member directory read
    if (pendingRows.length === 0) {
      return {
        periodId:    period.periodId,
        periodName:  period.name,
        total:       completedCount,
        pending:     0,
        completed:   completedCount,
        rate:        completedCount > 0 ? 100 : 0,
        members:     []
      };
    }

    // Build email → member record map for enrichment (cubicle, officeDays, hireDate)
    // Only loaded when there are pending members to enrich
    var memberMap = {};
    try {
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberSheet && memberSheet.getLastRow() > 1) {
        var mData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
        var mHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
        var mColMap = {};
        for (var h = 0; h < mHeaders.length; h++) {
          var hName = String(mHeaders[h]).trim().toLowerCase();
          if (hName) mColMap[hName] = h;
        }
        var mEmailCol = mColMap['email'] !== undefined ? mColMap['email'] : -1;
        var mCubicleCol = mColMap['cubicle'] !== undefined ? mColMap['cubicle'] : -1;
        var mOfficeDaysCol = mColMap['office days'] !== undefined ? mColMap['office days'] : -1;
        var mHireDateCol = mColMap['hire date'] !== undefined ? mColMap['hire date'] : -1;
        if (mEmailCol !== -1) {
          for (var mi = 0; mi < mData.length; mi++) {
            var mEmail = String(mData[mi][mEmailCol]).trim().toLowerCase();
            if (mEmail) {
              memberMap[mEmail] = {
                cubicle: mCubicleCol !== -1 ? String(mData[mi][mCubicleCol] || '').trim() : '',
                officeDays: mOfficeDaysCol !== -1 ? String(mData[mi][mOfficeDaysCol] || '').trim() : '',
                hireDate: mHireDateCol !== -1 ? (mData[mi][mHireDateCol] instanceof Date ? Utilities.formatDate(mData[mi][mHireDateCol], Session.getScriptTimeZone(), 'MM/dd/yyyy') : String(mData[mi][mHireDateCol] || '').trim()) : '',
              };
            }
          }
        }
      }
    } catch(enrichErr) {
      log_('getPendingSurveyMembers enrichment error', enrichErr.message);
    }

    // Second pass: build enriched pending member objects
    for (var pi = 0; pi < pendingRows.length; pi++) {
      var pRow = pendingRows[pi];
      var email = String(pRow[C.EMAIL - 1] || '').trim().toLowerCase();
      var enriched = memberMap[email] || {};
      var totalMissed = parseInt(pRow[C.TOTAL_MISSED - 1], 10) || 0;
      var totalCompleted = parseInt(pRow[C.TOTAL_COMPLETED - 1], 10) || 0;
      var totalSurveys = totalMissed + totalCompleted;
      var participationRate = totalSurveys > 0 ? Math.round((totalCompleted / totalSurveys) * 100) : null;

      pendingMembers.push({
        memberId:        String(pRow[C.MEMBER_ID        - 1] || ''),
        name:            String(pRow[C.MEMBER_NAME      - 1] || ''),
        email:           String(pRow[C.EMAIL            - 1] || ''),
        workLocation:    String(pRow[C.WORK_LOCATION    - 1] || ''),
        assignedSteward: String(pRow[C.ASSIGNED_STEWARD - 1] || ''),
        cubicle:         enriched.cubicle || '',
        officeDays:      enriched.officeDays || '',
        hireDate:        enriched.hireDate || '',
        participationRate: participationRate,
        totalCompleted:  totalCompleted,
        totalMissed:     totalMissed,
      });
    }

    var total = pendingMembers.length + completedCount;
    return {
      periodId:    period.periodId,
      periodName:  period.name,
      total:       total,
      pending:     pendingMembers.length,
      completed:   completedCount,
      rate:        total > 0 ? Math.round((completedCount / total) * 100) : 0,
      members:     pendingMembers
    };

  } catch(e) {
    log_('getPendingSurveyMembers error', e.message);
    return { periodId: null, periodName: null, total: 0, pending: 0, completed: 0, rate: 0, members: [] };
  }
}

/**
 * Returns aggregate section averages for the current period.
 * Reads from the anonymous Satisfaction sheet — no PII involved.
 * Values returned as plain numbers (not Sheets formulas).
 *
 * @returns {Object} {
 *   periodId, responseCount,
 *   sections: { OVERALL_SAT: {avg, count}, STEWARD_3A: {...}, ... }
 * }
 */
/**
 * Returns aggregate section averages for the current survey period.
 * v4.23.0: Fully dynamic — sections and questions read from 📋 Survey Questions sheet.
 * No hardcoded section definitions. Only slider-10 questions contribute to averages.
 * Reads only numeric columns from anonymous Satisfaction sheet — no PII involved.
 * Plain values (not Sheets formulas). Cached 10 minutes per period.
 *
 * @returns {{ periodId, periodName, responseCount, sections: Object }}
 */
function getSatisfactionSummary() {
  try {
    var period   = getSurveyPeriod();
    var periodId = period ? period.periodId : 'unknown';

    var cacheKey = 'satisfactionSummary_' + periodId;
    try {
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (_ce) { log_('_ce', (_ce.message || _ce)); }

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!satSheet || satSheet.getLastRow() < 2) {
      return { periodId: periodId, periodName: period ? period.name : '', responseCount: 0, sections: {} };
    }

    // ── Get col map and all data ──────────────────────────────────────────
    var colMap      = getSatisfactionColMap_();
    var lastRow     = satSheet.getLastRow();
    var lastCol     = satSheet.getLastColumn();
    var dataRowCount = lastRow - 1;
    var allData     = satSheet.getRange(2, 1, dataRowCount, lastCol).getValues();

    // ── Get slider questions grouped by section from Survey Questions sheet ──
    var questionsData = getSurveyQuestions();
    var allQs = questionsData.questions || [];

    // Build section map: sectionKey → { name, questionIds[] }
    // Only include slider-10 questions that have a column in the sheet
    var sectionMap = {};
    var sectionOrder = [];
    allQs.forEach(function(q) {
      if (q.type !== 'slider-10') return;
      var col = colMap[q.id];
      if (!col) return; // question not yet in satisfaction sheet
      if (!sectionMap[q.sectionKey]) {
        sectionMap[q.sectionKey] = { name: q.sectionTitle, qIds: [] };
        sectionOrder.push(q.sectionKey);
      }
      sectionMap[q.sectionKey].qIds.push(q.id);
    });

    // ── Compute section averages ─────────────────────────────────────────
    var sections = {};
    sectionOrder.forEach(function(key) {
      var def   = sectionMap[key];
      var total = 0;
      var count = 0;
      allData.forEach(function(row) {
        def.qIds.forEach(function(qId) {
          var col = colMap[qId];
          if (!col) return;
          var val = parseFloat(row[col - 1]);
          if (!isNaN(val) && val >= 1 && val <= 10) {
            total += val;
            count++;
          }
        });
      });
      sections[key] = {
        name:  def.name,
        avg:   count > 0 ? Math.round((total / count) * 10) / 10 : null,
        count: count
      };
    });

    var result = {
      periodId:      periodId,
      periodName:    period ? period.name : '',
      responseCount: dataRowCount,
      sections:      sections
    };

    try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 600); } catch (_ce) { log_('_ce', (_ce.message || _ce)); }
    return result;

  } catch(e) {
    log_('getSatisfactionSummary error', e.message);
    return { periodId: null, periodName: '', responseCount: 0, sections: {} };
  }
}

// ── Notifications ────────────────────────────────────────────────────────────

/**
 * Pushes a "Survey is now open" notification to all active members
 * using the existing notification system (Notifications sheet).
 * Guarded against missing notification infrastructure.
 *
 * @param {string} periodName - e.g. 'Q1 2026 Member Satisfaction Survey'
 */
function pushSurveyOpenNotification_(periodName) {
  try {
    // Use existing notification creation function if available
    if (typeof createNotification === 'function') {
      createNotification({
        type:      'Survey',
        title:     '📋 Survey Now Open',
        message:   'The ' + periodName + ' is now open. Your feedback is anonymous and helps shape your union. Complete it in the Member Portal.',
        priority:  'Normal',
        sentBy:    'system',
        sentByName:'Union Dashboard',
        // No expiry — stays until member dismisses or period closes
        expiresDate: '',
        recipient: 'All'
      });
      log_('pushSurveyOpenNotification_', 'Notification created for ' + periodName);
    } else {
      log_('pushSurveyOpenNotification_', 'createNotification not available, skipping.');
    }
  } catch(e) {
    log_('pushSurveyOpenNotification_ error', e.message);
  }
}

// ── Repeat reminder trigger ───────────────────────────────────────────────────

/**
 * Weekly reminder to members who haven't completed the survey yet.
 * Delegates to existing sendSurveyCompletionReminders() in 08c.
 * Runs on the weekly trigger installed alongside quarterly trigger.
 */
function autoSurveyReminderWeekly() {
  var period = getSurveyPeriod();
  if (!period) return; // No active period — nothing to do

  try {
    if (typeof sendSurveyCompletionReminders === 'function') {
      sendSurveyCompletionReminders();
    }
  } catch(e) {
    log_('autoSurveyReminderWeekly error', e.message);
  }
}

/**
 * Installs weekly reminder trigger. Call once after setupQuarterlyTrigger().
 */
function setupWeeklyReminderTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSurveyReminderWeekly') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('autoSurveyReminderWeekly')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(9)
    .create();

  log_('setupWeeklyReminderTrigger', 'Weekly reminder trigger installed (Tuesdays at 9 AM).');
}

// ── Global wrappers (callable from webapp via google.script.run) ─────────────

/**
 * Webapp-callable: returns current period + survey questions.
 * Delegates to getSurveyQuestions() in 08c_FormsAndNotifications.gs.
 */
// dataGetSurveyQuestions() is already wired in 21_WebDashDataService.gs

/**
 * Webapp-callable: returns pending survey members for steward dashboard.
 */
function dataGetPendingSurveyMembers(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { total: 0, pending: 0, completed: 0, rate: 0, members: [] };
  return getPendingSurveyMembers();
}

/**
 * Webapp-callable: returns satisfaction section averages.
 * Accessible to any authenticated user (member OR steward) — data is fully
 * anonymised aggregate stats with no PII. Members see results in Survey Results
 * page; stewards see them in the Insights panel.
 * @param {string} sessionToken
 */
function dataGetSatisfactionSummary(sessionToken) {
  // _resolveCallerEmail accepts steward OR member session tokens
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { periodId: null, periodName: '', responseCount: 0, sections: {} };
  return getSatisfactionSummary();
}

/**
 * Webapp-callable: steward opens a new survey period manually.
 */
function dataOpenNewSurveyPeriod(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Not authorized.' };
  return openNewSurveyPeriod(s);
}

// ── Menu-callable wrappers ───────────────────────────────────────────────────

/**
 * Menu: Tools → Survey Engine → Open New Survey Period
 * Confirms with user before opening (will archive any active period first).
 */
function menuOpenNewSurveyPeriod() {
  var ui = SpreadsheetApp.getUi();
  var active = getSurveyPeriod();
  var msg = active
    ? 'An active period (' + active.periodId + ') will be archived first. Open the new quarter period now?'
    : 'Open a new survey period for the current quarter now?';
  var resp = ui.alert('Open New Survey Period', msg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var caller = Session.getActiveUser().getEmail();
  var result = openNewSurveyPeriod(caller);
  ui.alert(result.success ? '✅ Opened' : '❌ Error', result.message, ui.ButtonSet.OK);
}

/**
 * Menu: Tools → Survey Engine → View Current Period Status
 */
function menuShowSurveyPeriodStatus() {
  var ui = SpreadsheetApp.getUi();
  var period = getSurveyPeriod();
  if (!period) {
    ui.alert('Survey Status', 'No active survey period. Use "Open New Survey Period" to start one.', ui.ButtonSet.OK);
    return;
  }
  var pending = getPendingSurveyMembers();
  var msg = [
    'Period: ' + period.name,
    'Status: ' + period.status,
    'Responses: ' + period.responseCount,
    '',
    'Completion: ' + pending.completed + '/' + pending.total + ' members (' + pending.rate + '%)',
    'Pending: ' + pending.pending + ' members'
  ].join('\n');
  ui.alert('📋 Current Survey Period', msg, ui.ButtonSet.OK);
}

// ── Combined trigger installer ────────────────────────────────────────────────

/**
 * v4.22.0 — Installs BOTH survey triggers in one call.
 * Run once after deployment:
 *   1. Open the Sheet → Extensions → Apps Script → Run menuInstallSurveyTriggers
 *   OR
 *   2. Union Hub menu → Survey Engine → Install Survey Triggers
 *
 * Safe to re-run: removes existing handlers before creating new ones.
 */
function menuInstallSurveyTriggers() {
  setupQuarterlyTrigger();
  setupWeeklyReminderTrigger();
  setupOpenDeferredTrigger();

  var msg = [
    '✅ All triggers installed:',
    '',
    '1. Quarterly auto-open — fires on the 1st of each month at 6 AM;',
    '   auto-opens a new period on Jan 1, Apr 1, Jul 1, Oct 1.',
    '',
    '2. Weekly member reminders — fires every Tuesday at 9 AM;',
    '   emails members who have not yet completed the active survey period.',
    '',
    '3. onOpen deferred init — fires on spreadsheet open (installable trigger);',
    '   runs syncColumnMaps, enforceHiddenSheets, tab colors, and toast.',
  ].join('\n');

  try {
    SpreadsheetApp.getUi().alert('Triggers Installed', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_uiErr) {
    log_('menuInstallSurveyTriggers', msg);
  }
}

// ── RTO Section Toggle ──────────────────────────────────────────────────────

/**
 * Menu: Surveys & Polls → Toggle Return-to-Office Questions
 *
 * Enables or disables the RTO_CHANGE section (q68–q76) on the
 * Survey Questions sheet by flipping the Active column (Y↔N).
 *
 * LOCK RULE: The toggle is locked (deactivation blocked) until at least
 * one survey period that included these questions has been Closed.
 * This ensures the RTO section runs for a full survey cycle before
 * it can be turned off.
 */
function menuToggleRTOSection() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Read Survey Questions sheet ──────────────────────────────────────
  var qSheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
  if (!qSheet || qSheet.getLastRow() < 2) {
    ui.alert('Error', 'Survey Questions sheet not found. Run Initialize Survey Engine first.', ui.ButtonSet.OK);
    return;
  }

  var QC = SURVEY_QUESTIONS_COLS;
  var data = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, 16).getValues();

  // Find all RTO_CHANGE rows
  var rtoRows = []; // { rowIndex (1-indexed sheet row), active }
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][QC.SECTION_KEY - 1]).trim() === 'RTO_CHANGE') {
      rtoRows.push({
        rowIndex: i + 2,
        active: String(data[i][QC.ACTIVE - 1]).trim().toUpperCase() === 'Y'
      });
    }
  }

  if (rtoRows.length === 0) {
    ui.alert('Not Found', 'No RTO questions (Section Key = RTO_CHANGE) found on the Survey Questions sheet.', ui.ButtonSet.OK);
    return;
  }

  var currentlyActive = rtoRows[0].active;

  // ── If turning OFF: check the period-end lock ────────────────────────
  if (currentlyActive) {
    var locked = isRTOToggleLocked_(ss);
    if (locked) {
      ui.alert(
        '🔒 Toggle Locked',
        'The Return-to-Office questions cannot be deactivated yet.\n\n' +
        'At least one full survey period that includes these questions must be completed (Closed) ' +
        'before the toggle becomes available.\n\n' +
        'This ensures every member gets a chance to respond to the RTO questions at least once.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Confirm deactivation
    var resp = ui.alert(
      'Deactivate RTO Questions?',
      'This will hide the Return-to-Office section (' + rtoRows.length + ' questions) from future surveys.\n\n' +
      'Existing responses are preserved. You can re-activate them later from this same menu.',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  // ── Toggle Active column ─────────────────────────────────────────────
  var newValue = currentlyActive ? 'N' : 'Y';
  for (var j = 0; j < rtoRows.length; j++) {
    qSheet.getRange(rtoRows[j].rowIndex, QC.ACTIVE, 1, 1).setValue(newValue);
  }

  // Clear cache so changes take effect immediately
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('surveyQuestions_v1');
    cache.remove('satisfactionColMap_v1');
  } catch (_c) { log_('_c', (_c.message || _c)); }

  var action = currentlyActive ? 'deactivated (hidden)' : 'activated (visible)';
  ui.alert(
    '✅ RTO Section ' + (currentlyActive ? 'Deactivated' : 'Activated'),
    rtoRows.length + ' Return-to-Office questions have been ' + action + '.\n\n' +
    'Changes take effect immediately on the next survey load.',
    ui.ButtonSet.OK
  );
}

/**
 * Checks whether the RTO toggle is still locked.
 * The lock is released once at least one Closed survey period exists
 * whose start date is on or after the date the RTO questions were added
 * (i.e., the period must have included RTO questions).
 *
 * Heuristic: if ANY Closed period exists whose start date >= the earliest
 * Timestamp row in Satisfaction that contains an RTO question answer, the
 * lock is released. As a simpler fallback: if any Closed period exists
 * and the RTO questions are present in the Satisfaction sheet headers,
 * the lock is released.
 *
 * @param {Spreadsheet} ss
 * @returns {boolean} true = locked (cannot deactivate)
 * @private
 */
function isRTOToggleLocked_(ss) {
  try {
    var periodsSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_PERIODS);
    if (!periodsSheet || periodsSheet.getLastRow() < 2) return true; // No periods at all

    var pData = periodsSheet.getRange(2, 1, periodsSheet.getLastRow() - 1, 8).getValues();
    var C = SURVEY_PERIODS_COLS;

    // Check if RTO question columns exist in Satisfaction sheet headers
    var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
    var rtoInHeaders = false;
    if (satSheet && satSheet.getLastRow() >= 1) {
      var headers = satSheet.getRange(1, 1, 1, satSheet.getLastColumn()).getValues()[0];
      rtoInHeaders = headers.some(function(h) {
        var id = String(h).trim();
        return id.indexOf('q6') === 0 && parseInt(id.replace('q', ''), 10) >= 68 &&
               parseInt(id.replace('q', ''), 10) <= 76;
      });
    }

    // If RTO columns aren't even in the Satisfaction sheet yet, no period has included them
    if (!rtoInHeaders) return true;

    // Look for any Closed period
    for (var i = 0; i < pData.length; i++) {
      var status = String(pData[i][C.STATUS - 1]).trim();
      if (status === 'Closed') return false; // At least one period completed — unlock
    }

    return true; // No closed periods yet
  } catch(e) {
    log_('isRTOToggleLocked_ error', e.message);
    return true; // Default to locked on error
  }
}

/**
 * Installs onOpenDeferred_ as an installable onOpen trigger.
 * FIX v4.25.7: onOpen (simple trigger) cannot call ScriptApp — this installable
 * trigger replaces the broken spawn-from-onOpen approach.
 * Safe to re-run — removes existing before creating new.
 */
function setupOpenDeferredTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpenDeferred_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('onOpenDeferred_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  log_('setupOpenDeferredTrigger', 'onOpenDeferred_ installed as installable onOpen trigger.');
}
