// ============================================================================
// 08e_SurveyEngine.gs — Survey Period Management & Aggregation
// v4.21.0 — Replaces Google Form integration with fully native webapp survey
// ============================================================================

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
  var ui = SpreadsheetApp.getUi();
  var log = [];

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
  Logger.log('initSurveyEngine:\n' + summary);
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
    .setBackground('#E0E7FF');

  sheet.setFrozenRows(1);

  // Hide the sheet at API level
  try { setSheetVeryHidden_(sheet); } catch(e) { sheet.hideSheet(); }

  Logger.log('setupSurveyPeriodsSheet: _Survey_Periods created.');
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
    Logger.log('getSurveyPeriod error: ' + e.message);
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
      var qStart = new Date(year, (quarter - 1) * 3, 1);
      var qEnd   = new Date(year,  quarter * 3,      0);  // last day of quarter

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

      Logger.log('openNewSurveyPeriod: Opened ' + periodId);
      return { success: true, periodId: periodId, message: periodName + ' is now open.' };

    } catch(e) {
      Logger.log('openNewSurveyPeriod error: ' + e.message);
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
            Logger.log('archiveSurveyPeriod_: Pinned Past Surveys Folder ID to Config: ' + parentFolder.getId());
          }
        }
      } catch(pe) {
        Logger.log('archiveSurveyPeriod_: Could not pin folder ID to Config: ' + pe.message);
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
        Logger.log('archiveSurveyPeriod_: question export error: ' + qe.message);
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
        Logger.log('archiveSurveyPeriod_: response export error: ' + re.message);
      }

    } catch(de) {
      Logger.log('archiveSurveyPeriod_: Drive error: ' + de.message);
    }

    // ── Mark period as Closed in _Survey_Periods ──────────────────────
    var rowData = periodsSheet.getRange(periodRow, 1, 1, 8).getValues()[0];
    rowData[C.STATUS      - 1] = 'Closed';
    rowData[C.ARCHIVE_URL - 1] = pastFolderUrl;
    periodsSheet.getRange(periodRow, 1, 1, 8).setValues([rowData]);

    Logger.log('archiveSurveyPeriod_: Archived ' + periodId + ' → ' + (pastFolderUrl || '(no Drive URL)'));

  } catch(e) {
    Logger.log('archiveSurveyPeriod_ error: ' + e.message);
  }
}

/**
 * Increments the response count for a period in _Survey_Periods.
 * @param {string} periodId
 */
function incrementPeriodResponseCount_(periodId) {
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
    Logger.log('incrementPeriodResponseCount_ error: ' + e.message);
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
    Logger.log('autoTriggerQuarterlyPeriod: not a quarter start, skipping.');
    return;
  }

  var result = openNewSurveyPeriod('Auto (quarterly trigger)');
  Logger.log('autoTriggerQuarterlyPeriod: ' + JSON.stringify(result));
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

  Logger.log('setupQuarterlyTrigger: Monthly trigger installed for autoTriggerQuarterlyPeriod on day 1 at 6 AM.');
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

    for (var i = 0; i < data.length; i++) {
      var status = String(data[i][C.CURRENT_STATUS - 1] || '').trim();
      if (status === 'Completed') {
        completedCount++;
      } else {
        pendingMembers.push({
          memberId:        String(data[i][C.MEMBER_ID        - 1] || ''),
          name:            String(data[i][C.MEMBER_NAME      - 1] || ''),
          email:           String(data[i][C.EMAIL            - 1] || ''),
          workLocation:    String(data[i][C.WORK_LOCATION    - 1] || ''),
          assignedSteward: String(data[i][C.ASSIGNED_STEWARD - 1] || '')
        });
      }
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
    Logger.log('getPendingSurveyMembers error: ' + e.message);
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
function getSatisfactionSummary() {
  try {
    var period = getSurveyPeriod();
    var periodId = period ? period.periodId : 'unknown';

    // Check cache first
    var cacheKey = 'satisfactionSummary_' + periodId;
    try {
      var cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch(ce) {}

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!satSheet || satSheet.getLastRow() < 2) {
      return { periodId: periodId, responseCount: 0, sections: {} };
    }

    var lastRow = satSheet.getLastRow();
    var dataRowCount = lastRow - 1;

    // Read all numeric columns in one batch (cols 1 through Q67 col = 68 cols)
    var allData = satSheet.getRange(2, 1, dataRowCount, 68).getValues();

    // Column groups to average (0-indexed into allData row arrays)
    // SATISFACTION_COLS values are 1-indexed, so subtract 1 for array access
    var SC = SATISFACTION_COLS;
    var sectionDefs = [
      { key: 'OVERALL_SAT',    name: 'Overall Satisfaction',      cols: [SC.Q6_SATISFIED_REP-1, SC.Q7_TRUST_UNION-1, SC.Q8_FEEL_PROTECTED-1, SC.Q9_RECOMMEND-1] },
      { key: 'STEWARD_3A',     name: 'Steward Ratings',           cols: [SC.Q10_TIMELY_RESPONSE-1, SC.Q11_TREATED_RESPECT-1, SC.Q12_EXPLAINED_OPTIONS-1, SC.Q13_FOLLOWED_THROUGH-1, SC.Q14_ADVOCATED-1, SC.Q15_SAFE_CONCERNS-1, SC.Q16_CONFIDENTIALITY-1] },
      { key: 'STEWARD_3B',     name: 'Steward Access',            cols: [SC.Q18_KNOW_CONTACT-1, SC.Q19_CONFIDENT_HELP-1, SC.Q20_EASY_FIND-1] },
      { key: 'CHAPTER',        name: 'Chapter Effectiveness',     cols: [SC.Q21_UNDERSTAND_ISSUES-1, SC.Q22_CHAPTER_COMM-1, SC.Q23_ORGANIZES-1, SC.Q24_REACH_CHAPTER-1, SC.Q25_FAIR_REP-1] },
      { key: 'LEADERSHIP',     name: 'Local Leadership',          cols: [SC.Q26_DECISIONS_CLEAR-1, SC.Q27_UNDERSTAND_PROCESS-1, SC.Q28_TRANSPARENT_FINANCE-1, SC.Q29_ACCOUNTABLE-1, SC.Q30_FAIR_PROCESSES-1, SC.Q31_WELCOMES_OPINIONS-1] },
      { key: 'CONTRACT',       name: 'Contract Enforcement',      cols: [SC.Q32_ENFORCES_CONTRACT-1, SC.Q33_REALISTIC_TIMELINES-1, SC.Q34_CLEAR_UPDATES-1, SC.Q35_FRONTLINE_PRIORITY-1] },
      { key: 'REPRESENTATION', name: 'Representation Process',    cols: [SC.Q37_UNDERSTOOD_STEPS-1, SC.Q38_FELT_SUPPORTED-1, SC.Q39_UPDATES_OFTEN-1, SC.Q40_OUTCOME_JUSTIFIED-1] },
      { key: 'COMMUNICATION',  name: 'Communication Quality',     cols: [SC.Q41_CLEAR_ACTIONABLE-1, SC.Q42_ENOUGH_INFO-1, SC.Q43_FIND_EASILY-1, SC.Q44_ALL_SHIFTS-1, SC.Q45_MEETINGS_WORTH-1] },
      { key: 'MEMBER_VOICE',   name: 'Member Voice & Culture',    cols: [SC.Q46_VOICE_MATTERS-1, SC.Q47_SEEKS_INPUT-1, SC.Q48_DIGNITY-1, SC.Q49_NEWER_SUPPORTED-1, SC.Q50_CONFLICT_RESPECT-1] },
      { key: 'VALUE_ACTION',   name: 'Value & Collective Action', cols: [SC.Q51_GOOD_VALUE-1, SC.Q52_PRIORITIES_NEEDS-1, SC.Q53_PREPARED_MOBILIZE-1, SC.Q54_HOW_INVOLVED-1, SC.Q55_WIN_TOGETHER-1] },
      { key: 'SCHEDULING',     name: 'Scheduling & Office Days',  cols: [SC.Q56_UNDERSTAND_CHANGES-1, SC.Q57_ADEQUATELY_INFORMED-1, SC.Q58_CLEAR_CRITERIA-1, SC.Q59_WORK_EXPECTATIONS-1, SC.Q60_EFFECTIVE_OUTCOMES-1, SC.Q61_SUPPORTS_WELLBEING-1, SC.Q62_CONCERNS_SERIOUS-1] }
    ];

    var sections = {};
    for (var s = 0; s < sectionDefs.length; s++) {
      var def   = sectionDefs[s];
      var total = 0;
      var count = 0;
      for (var row = 0; row < allData.length; row++) {
        for (var c = 0; c < def.cols.length; c++) {
          var val = parseFloat(allData[row][def.cols[c]]);
          if (!isNaN(val) && val >= 1 && val <= 10) {
            total += val;
            count++;
          }
        }
      }
      sections[def.key] = {
        name:  def.name,
        avg:   count > 0 ? Math.round((total / count) * 10) / 10 : null,
        count: count
      };
    }

    var result = {
      periodId:      periodId,
      periodName:    period ? period.name : '',
      responseCount: dataRowCount,
      sections:      sections
    };

    // Cache for 10 minutes
    try {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 600);
    } catch(ce) {}

    return result;

  } catch(e) {
    Logger.log('getSatisfactionSummary error: ' + e.message);
    return { periodId: null, responseCount: 0, sections: {} };
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
      Logger.log('pushSurveyOpenNotification_: Notification created for ' + periodName);
    } else {
      Logger.log('pushSurveyOpenNotification_: createNotification not available, skipping.');
    }
  } catch(e) {
    Logger.log('pushSurveyOpenNotification_ error: ' + e.message);
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
    Logger.log('autoSurveyReminderWeekly error: ' + e.message);
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

  Logger.log('setupWeeklyReminderTrigger: Weekly reminder trigger installed (Tuesdays at 9 AM).');
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
function dataGetPendingSurveyMembers() {
  var s = _requireStewardAuth();
  if (!s) return { total: 0, pending: 0, completed: 0, rate: 0, members: [] };
  return getPendingSurveyMembers();
}

/**
 * Webapp-callable: returns satisfaction section averages.
 */
function dataGetSatisfactionSummary() {
  return getSatisfactionSummary();
}

/**
 * Webapp-callable: steward opens a new survey period manually.
 */
function dataOpenNewSurveyPeriod() {
  var s = _requireStewardAuth();
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
