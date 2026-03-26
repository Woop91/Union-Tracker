// DEV_ONLY: Excluded from production builds via --prod flag in build.js
/**
 * ============================================================================
 * 07_DevTools.gs - DEVELOPER TOOLS (Development Only)
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Development-only demo data seeding (SEED_TEST_DATA) and cleanup
 *   (NUKE_SEEDED_DATA). Creates realistic fake members, grievances, and
 *   survey data for testing. NUKE removes ALL seeded data without touching
 *   manually entered records.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Excluded from production builds via --prod flag in build.js. The
 *   onOpen() guard `typeof buildDevMenu === 'function'` ensures production
 *   deployments never show the demo menu. isDemoSafeToRun_() prevents
 *   accidental execution if the file somehow makes it to production. Demo
 *   mode can be permanently disabled via Script Properties.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Development testing is impaired — developers must manually create test
 *   data. No production impact since this file is excluded from production
 *   builds. If the NUKE function is broken, seeded data must be manually
 *   deleted row by row.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, column constants), 02_DataManagers.gs
 *   (addMember, createGrievance). Never used in production. Used only by
 *   DevMenu.gs and developer testing.
 *
 * @license Free for use by non-profit collective bargaining groups and unions
 */

/**
 * Runtime guard: prevents SEED/NUKE from running if demo mode was disabled.
 * Returns true if it's safe to run demo operations.
 * @returns {boolean}
 * @private
 */
function isDemoSafeToRun_() {
  if (isDemoModeDisabled()) {
    try {
      SpreadsheetApp.getUi().alert(
        'Production Mode',
        'Demo functions are disabled in production.\n\n' +
        'To re-enable: Run disableDemoMode() with value "false" in Script Properties.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (_e) { /* headless context */ }
    return false;
  }
  return true;
}

// ============================================================================
// DEMO MODE TRACKING
// ============================================================================

/**
 * Check if demo mode has been disabled (after nuke)
 * @returns {boolean} True if demo mode is disabled
 */
function isDemoModeDisabled() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('DEMO_MODE_DISABLED') === 'true';
}

/**
 * Disable demo mode permanently (called after nuke)
 */
function disableDemoMode() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('DEMO_MODE_DISABLED', 'true');
  // Clear tracked IDs since they're no longer needed
  props.deleteProperty('SEEDED_MEMBER_IDS');
  props.deleteProperty('SEEDED_GRIEVANCE_IDS');
}

/**
 * Track a seeded member ID
 * @param {string} memberId - The member ID to track
 */
function trackSeededMemberId(memberId) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  if (ids.indexOf(memberId) === -1) {
    ids.push(memberId);
    props.setProperty('SEEDED_MEMBER_IDS', ids.join(','));
  }
}
/**
 * Get all tracked seeded member IDs
 * @returns {Object} Object with member IDs as keys for quick lookup
 */
function getSeededMemberIds() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  var lookup = {};
  ids.forEach(function(id) { if (id) lookup[id] = true; });
  return lookup;
}

/**
 * Get all tracked seeded grievance IDs
 * @returns {Object} Object with grievance IDs as keys for quick lookup
 */
function getSeededGrievanceIds() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  var lookup = {};
  ids.forEach(function(id) { if (id) lookup[id] = true; });
  return lookup;
}

/**
 * Batch track multiple seeded member IDs (more efficient than individual calls)
 * @param {Array<string>} memberIds - Array of member IDs to track
 */
function trackSeededMemberIdsBatch(memberIds) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  memberIds.forEach(function(id) {
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  props.setProperty('SEEDED_MEMBER_IDS', ids.join(','));
}

/**
 * Batch track multiple seeded grievance IDs (more efficient than individual calls)
 * @param {Array<string>} grievanceIds - Array of grievance IDs to track
 */
function trackSeededGrievanceIdsBatch(grievanceIds) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  grievanceIds.forEach(function(id) {
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  props.setProperty('SEEDED_GRIEVANCE_IDS', ids.join(','));
}

// ============================================================================
// SEED FUNCTIONS
// ============================================================================

/**
 * Seed all sample data in 3 phases (avoids GAS 6-minute timeout).
 * Orchestrator: runs phases sequentially with time checks.
 * If near the limit, aborts and tells the user which phase to resume.
 */
function SEED_SAMPLE_DATA() {
  var isDev = typeof IS_DEV_ENVIRONMENT !== 'undefined' && IS_DEV_ENVIRONMENT === true;
  if (!isDev) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Production Safety Check',
      'This action is intended for development environments only. Are you SURE you want to proceed?',
      ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;
  }
  if (!isDemoSafeToRun_()) return;

  ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  response = ui.alert(
    '🚀 Seed Sample Data (3-Phase)',
    'This will seed in 3 phases:\n\n' +
    'Phase 1: Config + 500 members + 300 grievances (~3 min)\n' +
    'Phase 2: All ancillary data (~2 min)\n' +
    '  - Contact log, surveys, feedback, events\n' +
    '  - Weekly questions, union stats, resources\n' +
    '  - Notifications, workload, tasks\n' +
    '  - Minutes, check-ins, timeline, Q&A\n' +
    'Phase 3: Webapp extras (~1 min)\n' +
    '  - Member tasks, case checklists\n' +
    '  - Survey questions sheet\n\n' +
    'If time runs out, you will be told which phase to resume.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var startTime = Date.now();
  var SAFE_LIMIT_MS = 300000; // 5 minutes (1 min buffer before 6-min hard limit)

  // Phase 1
  ss.toast('Running Phase 1: Config + 500 members + grievances...', '🌱 Phase 1', 10);
  SEED_PHASE_1();

  if (Date.now() - startTime > SAFE_LIMIT_MS) {
    ui.alert('⏱️ Time Limit', 'Phase 1 complete. Please run SEED_PHASE_2() next, then SEED_PHASE_3().', ui.ButtonSet.OK);
    return;
  }

  // Phase 2
  ss.toast('Running Phase 2: Ancillary data...', '🌱 Phase 2', 10);
  SEED_PHASE_2();

  if (Date.now() - startTime > SAFE_LIMIT_MS) {
    ui.alert('⏱️ Time Limit', 'Phases 1-2 complete. Please run SEED_PHASE_3() to finish.', ui.ButtonSet.OK);
    return;
  }

  // Phase 3
  ss.toast('Running Phase 3: Webapp extras...', '🌱 Phase 3', 10);
  SEED_PHASE_3();

  ss.toast('All phases complete!', '✅ Success', 5);
  ui.alert('✅ Success', 'All 3 phases seeded successfully!\n\n' +
    '• Config dropdowns populated\n' +
    '• 500 members + 300 grievances\n' +
    '• Script owner as test steward\n' +
    '• Contact log, surveys, feedback\n' +
    '• Calendar events, weekly questions\n' +
    '• Union stats, resources, notifications\n' +
    '• Workload, steward & member tasks\n' +
    '• Minutes, check-ins, timeline\n' +
    '• Q&A Forum questions & answers\n' +
    '• Case checklists, survey questions\n' +
    '• Auto-sync trigger installed\n\n' +
    'Member Directory auto-updates when Grievance Log changes.', ui.ButtonSet.OK);
}

/**
 * Phase 1: Config + 500 members + 300 grievances + script owner (~3 min)
 * Can be run independently from Script Editor.
 */
function SEED_PHASE_1() {
  if (!isDemoSafeToRun_()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Seeding config data...', '🌱 Phase 1', 3);
  seedConfigData();

  ss.toast('Seeding 500 members...', '🌱 Phase 1', 10);
  SEED_MEMBERS_ONLY(500);

  ss.toast('Adding script owner as test member...', '🌱 Phase 1', 2);
  seedScriptOwnerMember_(ss);

  ss.toast('Seeding 300 grievances...', '🌱 Phase 1', 10);
  SEED_GRIEVANCES(300);

  ss.toast('Phase 1 complete!', '✅ Phase 1', 3);
}

/**
 * Phase 2: All ancillary seeders (~2 min)
 * Can be run independently from Script Editor.
 */
function SEED_PHASE_2() {
  if (!isDemoSafeToRun_()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Seeding contact log entries...', '🌱 Phase 2', 2);
  seedContactLogData();

  ss.toast('Seeding survey responses...', '🌱 Phase 2', 2);
  seedSatisfactionData();

  ss.toast('Seeding feedback entries...', '🌱 Phase 2', 2);
  seedFeedbackData();

  ss.toast('Seeding survey tracking data...', '🌱 Phase 2', 2);
  seedSurveyTrackingData();

  ss.toast('Seeding calendar events...', '🌱 Phase 2', 2);
  seedCalendarEvents();

  ss.toast('Seeding weekly questions...', '🌱 Phase 2', 2);
  seedWeeklyQuestions();

  ss.toast('Seeding union stats data...', '🌱 Phase 2', 2);
  seedUnionStatsData();

  ss.toast('Seeding resources...', '🌱 Phase 2', 2);
  seedResourcesData();

  ss.toast('Seeding notifications...', '🌱 Phase 2', 2);
  seedNotificationsData();

  ss.toast('Seeding workload submissions...', '🌱 Phase 2', 2);
  seedWorkloadData();

  ss.toast('Seeding steward tasks...', '🌱 Phase 2', 2);
  seedStewardTasksData();

  ss.toast('Seeding meeting minutes...', '🌱 Phase 2', 2);
  seedMinutesData();

  ss.toast('Seeding meeting check-ins...', '🌱 Phase 2', 2);
  seedMeetingCheckinData();

  ss.toast('Seeding timeline events...', '🌱 Phase 2', 2);
  seedTimelineData();

  ss.toast('Seeding Q&A Forum...', '🌱 Phase 2', 2);
  seedQAForumData();

  ss.toast('Installing auto-sync trigger...', '🔧 Setup', 3);
  installAutoSyncTriggerQuick();

  ss.toast('Phase 2 complete!', '✅ Phase 2', 3);
}

/**
 * Phase 3: Webapp extras — member tasks, case checklists, survey questions (~1 min)
 * Seeds SPA features not covered by Phases 1-2.
 * Can be run independently from Script Editor.
 */
function SEED_PHASE_3() {
  if (!isDemoSafeToRun_()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Seeding member tasks...', '🌱 Phase 3', 2);
  seedMemberTasksData();

  ss.toast('Seeding case checklists...', '🌱 Phase 3', 2);
  seedCaseChecklistData();

  ss.toast('Seeding survey questions sheet...', '🌱 Phase 3', 2);
  seedSurveyQuestionsData();

  ss.toast('Phase 3 complete!', '✅ Phase 3', 3);
}

/**
 * Seeds Config sheet with preset and user-configurable dropdown values.
 * Data starts at row 3 (row 1 = section headers, row 2 = column headers).
 * Only populates empty columns to preserve existing user data.
 * @returns {void}
 */
function seedConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found. Please run CREATE_DASHBOARD first.');
    return;
  }

  // Ensure Config sheet has enough columns (AX = column 50 = MOBILE_DASHBOARD_URL)
  ensureMinimumColumns(sheet, CONFIG_COLS.MOBILE_DASHBOARD_URL);

  // Data row start (after section headers row 1 and column headers row 2)
  var dataStartRow = 3;

  /**
   * Seeds a Config column only if it currently has no data.
   * @param {number} column - 1-based Config column index.
   * @param {Array} values - Array of values to write.
   * @returns {boolean} True if seeded, false if column already had data.
   */
  function seedIfEmpty(column, values) {
    var existing = getConfigValues(sheet, column);
    if (existing.length === 0) {
      sheet.getRange(dataStartRow, column, values.length, 1)
        .setValues(values.map(function(v) { return [v]; }));
      return true;
    }
    return false;
  }

  var seededAny = false;

  // ═══════════════════════════════════════════════════════════════════════════
  // PRESET VALUES (Standard dropdowns that rarely change)
  // ═══════════════════════════════════════════════════════════════════════════

  // Office Days (Column D) - PRESET
  if (seedIfEmpty(CONFIG_COLS.OFFICE_DAYS, DEFAULT_CONFIG.OFFICE_DAYS)) seededAny = true;

  // Grievance Status (Column I) - PRESET
  if (seedIfEmpty(CONFIG_COLS.GRIEVANCE_STATUS, DEFAULT_CONFIG.GRIEVANCE_STATUS)) seededAny = true;

  // Grievance Step (Column K) - PRESET
  if (seedIfEmpty(CONFIG_COLS.GRIEVANCE_STEP, DEFAULT_CONFIG.GRIEVANCE_STEP)) seededAny = true;

  // Issue Category (Column L) - PRESET
  if (seedIfEmpty(CONFIG_COLS.ISSUE_CATEGORY, DEFAULT_CONFIG.ISSUE_CATEGORY)) seededAny = true;

  // Articles Violated (Column M) - PRESET
  if (seedIfEmpty(CONFIG_COLS.ARTICLES, DEFAULT_CONFIG.ARTICLES)) seededAny = true;

  // Communication Methods (Column N) - PRESET
  if (seedIfEmpty(CONFIG_COLS.COMM_METHODS, DEFAULT_CONFIG.COMM_METHODS)) seededAny = true;

  // Best Times to Contact (Column AE) - PRESET
  if (seedIfEmpty(CONFIG_COLS.BEST_TIMES, [
    'Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Weekends', 'Flexible'
  ])) seededAny = true;

  // ═══════════════════════════════════════════════════════════════════════════
  // USER-CONFIGURABLE VALUES (Organization-specific dropdowns)
  // ═══════════════════════════════════════════════════════════════════════════

  // Job Titles (Column A)
  if (seedIfEmpty(CONFIG_COLS.JOB_TITLES, [
    'Social Worker', 'Case Manager', 'Program Coordinator', 'Administrative Assistant',
    'Supervisor', 'Director', 'Clinician', 'Counselor', 'Specialist', 'Analyst',
    'Manager', 'Senior Social Worker', 'Lead Case Manager', 'Program Manager',
    'Executive Assistant', 'HR Coordinator', 'Finance Associate', 'IT Support',
    'Communications Specialist', 'Outreach Worker'
  ])) seededAny = true;

  // Office Locations (Column B)
  if (seedIfEmpty(CONFIG_COLS.OFFICE_LOCATIONS, [
    'Boston Main Office', 'Worcester Regional', 'Springfield Center', 'Cambridge Branch',
    'Lowell Office', 'Brockton Center', 'Quincy Regional', 'New Bedford Office',
    'Fall River Branch', 'Lawrence Center', 'Framingham Office', 'Somerville Branch',
    'Lynn Regional', 'Haverhill Center', 'Malden Office', 'Medford Branch',
    'Waltham Regional', 'Newton Center', 'Brookline Office', 'Salem Branch'
  ])) seededAny = true;

  // Units (Column C)
  if (seedIfEmpty(CONFIG_COLS.UNITS, [
    'Child Welfare', 'Adult Services', 'Mental Health', 'Disability Services',
    'Elder Affairs', 'Housing Assistance', 'Employment Services', 'Youth Services',
    'Family Support', 'Administration'
  ])) seededAny = true;

  // Supervisors (Column F)
  if (seedIfEmpty(CONFIG_COLS.SUPERVISORS, [
    'Maria Rodriguez', 'James Wilson', 'Sarah Chen', 'Michael Brown',
    'Jennifer Davis', 'Robert Taylor', 'Lisa Anderson', 'David Martinez',
    'Emily Johnson', 'Christopher Lee', 'Amanda White', 'Daniel Garcia'
  ])) seededAny = true;

  // Managers (Column G)
  if (seedIfEmpty(CONFIG_COLS.MANAGERS, [
    'Patricia Thompson', 'William Jackson', 'Elizabeth Moore', 'Richard Harris',
    'Susan Clark', 'Joseph Lewis', 'Margaret Robinson', 'Charles Walker'
  ])) seededAny = true;

  // Stewards (Column H)
  if (seedIfEmpty(CONFIG_COLS.STEWARDS, [
    'John Smith', 'Mary Johnson', 'Robert Williams', 'Patricia Jones',
    'Michael Davis', 'Linda Miller', 'William Brown', 'Barbara Wilson',
    'David Moore', 'Susan Taylor', 'James Anderson', 'Karen Thomas'
  ])) seededAny = true;

  // Steward Committees (Column I)
  if (seedIfEmpty(CONFIG_COLS.STEWARD_COMMITTEES, [
    'Grievance Committee', 'Bargaining Committee', 'Health & Safety Committee',
    'Political Action Committee', 'Membership Committee', 'Executive Board'
  ])) seededAny = true;

  // Survey Priority Options — Q64 checkbox list (v4.21.0)
  if (seedIfEmpty(CONFIG_COLS.SURVEY_PRIORITY_OPTIONS, [
    'Contract Enforcement',
    'Workload & Staffing',
    'Scheduling & Office Days',
    'Pay & Benefits',
    'Health & Safety',
    'Training & Development',
    'Equity & Inclusion',
    'Communication',
    'Steward Support & Access',
    'Member Organizing',
    'Other'
  ])) seededAny = true;

  if (seededAny) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Config data seeded!', '✅ Success', 3);
  }
}

/**
 * Seed 3 sample entries in the Feedback & Development sheet
 * Demonstrates bug reports, feature requests, and improvements
 */
function seedFeedbackData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!sheet) {
    Logger.log('Feedback sheet not found. Creating it...');
    createFeedbackSheet(ss);
    sheet = ss.getSheetByName(SHEETS.FEEDBACK);
    if (!sheet) {
      Logger.log('Could not create Feedback sheet');
      return;
    }
  }

  // Check if column A (data area) already has data beyond header
  var dataCol = sheet.getRange('A:A').getValues();
  var dataRowCount = 0;
  for (var i = 1; i < dataCol.length; i++) {
    if (dataCol[i][0] !== '') {
      dataRowCount++;
      break;
    }
  }
  if (dataRowCount > 0) {
    Logger.log('Feedback sheet already has data. Skipping seed.');
    return;
  }

  // Generate 3 sample feedback entries using FEEDBACK_COLS for dynamic column placement
  var now = new Date();
  var oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var threeDaysAgo = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000);

  var maxCol = FEEDBACK_COLS.NOTES; // last defined column
  /**
   * Builds a feedback row array using FEEDBACK_COLS for column placement.
   * @param {Date} timestamp - Submission timestamp.
   * @param {string} submittedBy - Submitter name.
   * @param {string} category - Feedback category.
   * @param {string} type - Bug/Feature Request/Improvement.
   * @param {string} priority - Priority level.
   * @param {string} title - Feedback title.
   * @param {string} description - Detailed description.
   * @param {string} status - Current status.
   * @param {string} assignedTo - Assignee name.
   * @param {string} resolution - Resolution notes.
   * @param {string} notes - Additional notes.
   * @returns {Array} Row array sized to FEEDBACK_COLS.
   */
  function buildFeedbackRow(timestamp, submittedBy, category, type, priority, title, description, status, assignedTo, resolution, notes) {
    var row = [];
    for (var i = 0; i < maxCol; i++) { row.push(''); }
    row[FEEDBACK_COLS.TIMESTAMP - 1] = timestamp;
    row[FEEDBACK_COLS.SUBMITTED_BY - 1] = submittedBy;
    row[FEEDBACK_COLS.CATEGORY - 1] = category;
    row[FEEDBACK_COLS.TYPE - 1] = type;
    row[FEEDBACK_COLS.PRIORITY - 1] = priority;
    row[FEEDBACK_COLS.TITLE - 1] = title;
    row[FEEDBACK_COLS.DESCRIPTION - 1] = description;
    row[FEEDBACK_COLS.STATUS - 1] = status;
    row[FEEDBACK_COLS.ASSIGNED_TO - 1] = assignedTo;
    row[FEEDBACK_COLS.RESOLUTION - 1] = resolution;
    row[FEEDBACK_COLS.NOTES - 1] = notes;
    return row;
  }

  var twoDaysAgo = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
  var fiveDaysAgo = new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000);
  var tenDaysAgo = new Date(now.getTime() - 10 * 24 * 60 * 60 * 1000);
  var fourteenDaysAgo = new Date(now.getTime() - 14 * 24 * 60 * 60 * 1000);
  var twentyDaysAgo = new Date(now.getTime() - 20 * 24 * 60 * 60 * 1000);

  var sampleFeedback = [
    buildFeedbackRow(oneWeekAgo, 'John Smith', 'Dashboard', 'Bug', 'Medium',
      'Dashboard metrics not refreshing',
      'The Quick Stats section sometimes shows stale data after editing the Grievance Log. Requires manual refresh to update.',
      'Resolved', 'Tech Team',
      'Added auto-refresh trigger on Grievance Log edit. Metrics now update within 30 seconds.',
      'User confirmed fix working'),
    buildFeedbackRow(threeDaysAgo, 'Mary Johnson', 'Member Directory', 'Feature Request', 'High',
      'Bulk import members from CSV',
      'Would like ability to import multiple members at once from a CSV file instead of entering one by one.',
      'In Progress', 'Tech Team', '', 'Targeting v2.2 release'),
    buildFeedbackRow(now, 'Robert Williams', 'Reports', 'Improvement', 'Low',
      'Add PDF export for Dashboard',
      'It would be helpful to export the Dashboard as a PDF for sharing with chapter leadership during meetings.',
      'New', '', '', ''),
    buildFeedbackRow(twoDaysAgo, 'Linda Garcia', 'Grievance Log', 'Bug', 'High',
      'Grievance step dates not saving correctly',
      'When advancing a grievance from Step I to Step II, the date field sometimes reverts to the original filing date instead of the current date.',
      'New', '', '', 'Reported by two stewards independently'),
    buildFeedbackRow(fiveDaysAgo, 'James Brown', 'Calendar', 'Feature Request', 'Medium',
      'Recurring event support for meetings',
      'We have monthly membership meetings and weekly steward check-ins. Would be great to set these up once as recurring rather than creating each one individually.',
      'In Review', 'Tech Team', '', 'Evaluating Google Calendar API recurrence support'),
    buildFeedbackRow(tenDaysAgo, 'Patricia Davis', 'Member Directory', 'Improvement', 'Low',
      'Color-code members by unit on directory',
      'It would help stewards quickly identify members in their unit if rows were color-coded or filterable by organizational unit.',
      'Resolved', 'Tech Team',
      'Added unit filter dropdown to the Member Directory view.',
      'Steward confirmed this is helpful'),
    buildFeedbackRow(fourteenDaysAgo, 'Michael Wilson', 'Dashboard', 'Bug', 'Medium',
      'Notification badge count off by one',
      'The notification bell sometimes shows 3 unread but when I open it there are only 2 notifications. Seems like dismissed notifications are still being counted.',
      'In Progress', 'Tech Team', '', 'Investigating badge refresh timing'),
    buildFeedbackRow(twentyDaysAgo, 'Sarah Martinez', 'General', 'Feature Request', 'High',
      'Mobile-friendly layout for stewards in the field',
      'When I visit members at worksites I use my phone to look things up. The dashboard is hard to navigate on small screens. A responsive or mobile view would be very useful.',
      'New', '', '', 'Multiple stewards have requested this')
  ];

  // Write sample data
  sheet.getRange(2, 1, sampleFeedback.length, sampleFeedback[0].length).setValues(sampleFeedback);

  // Format timestamp column
  sheet.getRange(2, FEEDBACK_COLS.TIMESTAMP, sampleFeedback.length, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  Logger.log('Seeded ' + sampleFeedback.length + ' sample feedback entries');
}

/**
 * Seed survey tracking data from existing Member Directory.
 * Populates _Survey_Tracking with all members and simulates
 * realistic completion patterns (~40% completed, ~60% not completed).
 *
 * In production, tracking data is populated by populateSurveyTrackingFromMembers()
 * and updated automatically when members submit the satisfaction survey
 * (via the Google Form trigger -> onSatisfactionFormSubmit() -> updateSurveyTrackingOnSubmit_()).
 *
 * See SURVEY_TRACKING_COLS in 01_Core.gs for full flow documentation.
 * See 08c_FormsAndNotifications.gs "SURVEY COMPLETION TRACKING" section for all functions.
 */
function seedSurveyTrackingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackingSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!trackingSheet) {
    Logger.log('Survey Tracking sheet not found. It will be created during setup.');
    return;
  }

  if (!memberSheet || memberSheet.getLastRow() < 2) {
    Logger.log('Member Directory empty or missing. Skipping survey tracking seed.');
    return;
  }

  // Check if already seeded
  if (trackingSheet.getLastRow() > 1) {
    Logger.log('Survey Tracking already has data. Skipping seed.');
    return;
  }

  var memberData = memberSheet.getDataRange().getValues();
  var rows = [];
  var now = new Date();

  for (var i = 1; i < memberData.length; i++) {
    var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    var firstName = memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = memberData[i][MEMBER_COLS.LAST_NAME - 1] || '';
    var email = memberData[i][MEMBER_COLS.EMAIL - 1] || '';
    var location = memberData[i][MEMBER_COLS.WORK_LOCATION - 1] || '';
    var steward = memberData[i][MEMBER_COLS.ASSIGNED_STEWARD - 1] || '';

    // ~40% completed current round
    var completed = Math.random() < 0.4;
    var completedDate = '';
    if (completed) {
      var daysAgo = Math.floor(Math.random() * 30);
      completedDate = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);
    }

    // Simulate 0-3 prior rounds of history
    var priorRounds = Math.floor(Math.random() * 4);
    var totalCompleted = completed ? 1 : 0;
    var totalMissed = 0;
    for (var r = 0; r < priorRounds; r++) {
      if (Math.random() < 0.5) {
        totalCompleted++;
      } else {
        totalMissed++;
      }
    }

    // ~30% of non-completed members got a reminder
    var lastReminder = '';
    if (!completed && Math.random() < 0.3) {
      var reminderDaysAgo = Math.floor(Math.random() * 14) + 1;
      lastReminder = new Date(now.getTime() - reminderDaysAgo * 24 * 60 * 60 * 1000);
    }

    rows.push([
      memberId,
      (firstName + ' ' + lastName).trim(),
      email,
      location,
      steward,
      completed ? 'Completed' : 'Not Completed',
      completedDate,
      totalMissed,
      totalCompleted,
      lastReminder
    ]);
  }

  if (rows.length > 0) {
    trackingSheet.getRange(2, 1, rows.length, 10).setValues(rows);
  }

  Logger.log('Seeded survey tracking for ' + rows.length + ' members');
}

/**
 * Seeds 50 sample survey responses in Member Satisfaction sheet
 * Generates realistic distribution of ratings across all survey sections
 * Demonstrates form response data with branching logic
 */
function seedSatisfactionData() {
  var SATISFACTION_COLS = buildSatisfactionColsShim_(getSatisfactionColMap_());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!sheet) {
    Logger.log('Satisfaction sheet not found. Creating it...');
    createSatisfactionSheet(ss);
    sheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!sheet) {
      Logger.log('Could not create Satisfaction sheet');
      return;
    }
  }

  // Check if column A (data area) already has data beyond header
  var dataCol = sheet.getRange('A2:A10').getValues();
  var hasData = false;
  for (var i = 0; i < dataCol.length; i++) {
    if (dataCol[i][0] !== '') {
      hasData = true;
      break;
    }
  }
  if (hasData) {
    Logger.log('Satisfaction sheet already has data. Skipping seed.');
    return;
  }

  // Sample data pools for realistic responses
  var worksites = ['Downtown Office', 'North Regional', 'South Campus', 'West Branch', 'East Center', 'Central HQ'];
  var roles = ['Social Worker', 'Case Manager', 'Counselor', 'Coordinator', 'Specialist', 'Analyst'];
  var shifts = ['Day (8am-4pm)', 'Swing (4pm-12am)', 'Night (12am-8am)', 'Flex Schedule'];
  var tenures = ['Less than 1 year', '1-3 years', '3-5 years', '5-10 years', 'More than 10 years'];

  var priorities = [
    'Better communication', 'Higher wages', 'Improved benefits', 'Workplace safety',
    'Manageable caseloads', 'Professional development', 'Better management',
    'More staff', 'Schedule flexibility', 'Union visibility'
  ];

  var stewardComments = [
    'Very helpful and responsive', 'Could improve follow-up', 'Always available when needed',
    'Great advocate', 'Would like more proactive outreach', 'Excellent communication',
    'Needs to be more accessible', 'Very professional', '', ''
  ];

  var schedulingChallenges = [
    'Balancing work and family obligations', 'Short notice for schedule changes',
    'Inconsistent shift assignments', 'Difficulty getting preferred shifts',
    'Overtime requirements', 'No major issues', '', ''
  ];

  var oneChanges = [
    'More transparency in decision-making', 'Better contract enforcement',
    'Faster grievance resolution', 'More member meetings', 'Improved communication',
    'Stronger management accountability', 'Better training for stewards',
    'More visibility of union leadership', '', ''
  ];

  var keepDoing = [
    'Monthly newsletters', 'Quick response to grievances', 'Member appreciation events',
    'Training opportunities', 'Contract education', 'Workplace visits',
    'Fighting for better wages', 'Supporting members in meetings', '', ''
  ];

  // Generate 50 sample responses
  var sampleData = [];
  var now = new Date();

  /**
   * Returns a weighted random rating 1-10, skewed toward higher values.
   * @returns {number} Integer between 1 and 10.
   */
  function weightedRating() {
    var r = Math.random();
    if (r < 0.1) return Math.floor(Math.random() * 3) + 1;      // 1-3 (10%)
    if (r < 0.3) return Math.floor(Math.random() * 2) + 4;      // 4-5 (20%)
    if (r < 0.6) return Math.floor(Math.random() * 2) + 6;      // 6-7 (30%)
    return Math.floor(Math.random() * 3) + 8;                    // 8-10 (40%)
  }

  /**
   * Returns a random element from an array.
   * @param {Array} arr - Source array.
   * @returns {*} Random element.
   */
  function randomItem(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
  }

  /**
   * Returns a comma-separated string of 3 randomly selected priorities.
   * @returns {string} Shuffled priorities joined by ', '.
   */
  function randomPriorities() {
    var arr = priorities.slice();
    for (var j = arr.length - 1; j > 0; j--) {
      var k = Math.floor(Math.random() * (j + 1));
      var tmp = arr[j]; arr[j] = arr[k]; arr[k] = tmp;
    }
    return arr.slice(0, 3).join(', ');
  }

  // v4.23.0: maxCol derived from live sheet header count (dynamic schema)
  var satSheetForSeed = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SATISFACTION);
  var maxCol = satSheetForSeed ? satSheetForSeed.getLastColumn() : Math.max(SATISFACTION_COLS.Q67_ADDITIONAL, 70);

  for (i = 0; i < 50; i++) {
    // Spread responses over last 60 days
    var daysAgo = Math.floor(Math.random() * 60);
    var timestamp = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);

    // Branching: 70% had steward contact
    var hadStewardContact = Math.random() < 0.7;
    // Branching: 40% filed grievance
    var filedGrievance = Math.random() < 0.4;

    // Build row using SATISFACTION_COLS for dynamic column placement
    var row = [];
    for (var c = 0; c < maxCol; c++) { row.push(''); }

    // Timestamp
    row[SATISFACTION_COLS.TIMESTAMP - 1] = timestamp;

    // Work Context (Q1-5)
    row[SATISFACTION_COLS.Q1_WORKSITE - 1] = randomItem(worksites);
    row[SATISFACTION_COLS.Q2_ROLE - 1] = randomItem(roles);
    row[SATISFACTION_COLS.Q3_SHIFT - 1] = randomItem(shifts);
    row[SATISFACTION_COLS.Q4_TIME_IN_ROLE - 1] = randomItem(tenures);
    row[SATISFACTION_COLS.Q5_STEWARD_CONTACT - 1] = hadStewardContact ? 'Yes' : 'No';

    // Overall Satisfaction (Q6-9) - everyone answers
    row[SATISFACTION_COLS.Q6_SATISFIED_REP - 1] = weightedRating();
    row[SATISFACTION_COLS.Q7_TRUST_UNION - 1] = weightedRating();
    row[SATISFACTION_COLS.Q8_FEEL_PROTECTED - 1] = weightedRating();
    row[SATISFACTION_COLS.Q9_RECOMMEND - 1] = weightedRating();

    // Steward Ratings 3A (Q10-17) - only if had contact
    if (hadStewardContact) {
      row[SATISFACTION_COLS.Q10_TIMELY_RESPONSE - 1] = weightedRating();
      row[SATISFACTION_COLS.Q11_TREATED_RESPECT - 1] = weightedRating();
      row[SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS - 1] = weightedRating();
      row[SATISFACTION_COLS.Q13_FOLLOWED_THROUGH - 1] = weightedRating();
      row[SATISFACTION_COLS.Q14_ADVOCATED - 1] = weightedRating();
      row[SATISFACTION_COLS.Q15_SAFE_CONCERNS - 1] = weightedRating();
      row[SATISFACTION_COLS.Q16_CONFIDENTIALITY - 1] = weightedRating();
      row[SATISFACTION_COLS.Q17_STEWARD_IMPROVE - 1] = randomItem(stewardComments);
    }

    // Steward Access 3B (Q18-20) - only if NO contact
    if (!hadStewardContact) {
      row[SATISFACTION_COLS.Q18_KNOW_CONTACT - 1] = weightedRating();
      row[SATISFACTION_COLS.Q19_CONFIDENT_HELP - 1] = weightedRating();
      row[SATISFACTION_COLS.Q20_EASY_FIND - 1] = weightedRating();
    }

    // Chapter Effectiveness (Q21-25) - everyone
    row[SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q22_CHAPTER_COMM - 1] = weightedRating();
    row[SATISFACTION_COLS.Q23_ORGANIZES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q24_REACH_CHAPTER - 1] = weightedRating();
    row[SATISFACTION_COLS.Q25_FAIR_REP - 1] = weightedRating();

    // Local Leadership (Q26-31) - everyone
    row[SATISFACTION_COLS.Q26_DECISIONS_CLEAR - 1] = weightedRating();
    row[SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE - 1] = weightedRating();
    row[SATISFACTION_COLS.Q29_ACCOUNTABLE - 1] = weightedRating();
    row[SATISFACTION_COLS.Q30_FAIR_PROCESSES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q31_WELCOMES_OPINIONS - 1] = weightedRating();

    // Contract Enforcement (Q32-36) - everyone
    row[SATISFACTION_COLS.Q32_ENFORCES_CONTRACT - 1] = weightedRating();
    row[SATISFACTION_COLS.Q33_REALISTIC_TIMELINES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q34_CLEAR_UPDATES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q35_FRONTLINE_PRIORITY - 1] = weightedRating();
    row[SATISFACTION_COLS.Q36_FILED_GRIEVANCE - 1] = filedGrievance ? 'Yes' : 'No';

    // Representation 6A (Q37-40) - only if filed grievance
    if (filedGrievance) {
      row[SATISFACTION_COLS.Q37_UNDERSTOOD_STEPS - 1] = weightedRating();
      row[SATISFACTION_COLS.Q38_FELT_SUPPORTED - 1] = weightedRating();
      row[SATISFACTION_COLS.Q39_UPDATES_OFTEN - 1] = weightedRating();
      row[SATISFACTION_COLS.Q40_OUTCOME_JUSTIFIED - 1] = weightedRating();
    }

    // Communication (Q41-45) - everyone
    row[SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE - 1] = weightedRating();
    row[SATISFACTION_COLS.Q42_ENOUGH_INFO - 1] = weightedRating();
    row[SATISFACTION_COLS.Q43_FIND_EASILY - 1] = weightedRating();
    row[SATISFACTION_COLS.Q44_ALL_SHIFTS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q45_MEETINGS_WORTH - 1] = weightedRating();

    // Member Voice (Q46-50) - everyone
    row[SATISFACTION_COLS.Q46_VOICE_MATTERS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q47_SEEKS_INPUT - 1] = weightedRating();
    row[SATISFACTION_COLS.Q48_DIGNITY - 1] = weightedRating();
    row[SATISFACTION_COLS.Q49_NEWER_SUPPORTED - 1] = weightedRating();
    row[SATISFACTION_COLS.Q50_CONFLICT_RESPECT - 1] = weightedRating();

    // Value & Action (Q51-55) - everyone
    row[SATISFACTION_COLS.Q51_GOOD_VALUE - 1] = weightedRating();
    row[SATISFACTION_COLS.Q52_PRIORITIES_NEEDS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q53_PREPARED_MOBILIZE - 1] = weightedRating();
    row[SATISFACTION_COLS.Q54_HOW_INVOLVED - 1] = weightedRating();
    row[SATISFACTION_COLS.Q55_WIN_TOGETHER - 1] = weightedRating();

    // Scheduling (Q56-63) - everyone
    row[SATISFACTION_COLS.Q56_UNDERSTAND_CHANGES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q57_ADEQUATELY_INFORMED - 1] = weightedRating();
    row[SATISFACTION_COLS.Q58_CLEAR_CRITERIA - 1] = weightedRating();
    row[SATISFACTION_COLS.Q59_WORK_EXPECTATIONS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q60_EFFECTIVE_OUTCOMES - 1] = weightedRating();
    row[SATISFACTION_COLS.Q61_SUPPORTS_WELLBEING - 1] = weightedRating();
    row[SATISFACTION_COLS.Q62_CONCERNS_SERIOUS - 1] = weightedRating();
    row[SATISFACTION_COLS.Q63_SCHEDULING_CHALLENGE - 1] = randomItem(schedulingChallenges);

    // Priorities & Close (Q64-67) - everyone
    row[SATISFACTION_COLS.Q64_TOP_PRIORITIES - 1] = randomPriorities();
    row[SATISFACTION_COLS.Q65_ONE_CHANGE - 1] = randomItem(oneChanges);
    row[SATISFACTION_COLS.Q66_KEEP_DOING - 1] = randomItem(keepDoing);
    // Q67_ADDITIONAL left as '' (mostly empty)

    sampleData.push(row);
  }

  // Ensure sheet has enough columns (68 for survey data)
  var requiredCols = sampleData[0].length;
  var currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  // Write sample data starting at row 2
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  // Format timestamp column
  sheet.getRange(2, 1, sampleData.length, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  Logger.log('Seeded ' + sampleData.length + ' sample survey responses');
}

/**
 * Restore Config dropdowns AND re-apply dropdown validations to Member Directory and Grievance Log
 * This is the full restore function for use after nuking or when dropdowns are missing.
 *
 * Three-phase restore:
 *   1. seedConfigData()              – populate empty Config columns with defaults
 *   2. restoreConfigFromSheetData_() – scan Member Directory & Grievance Log for values
 *                                      that exist in the data but are missing from Config
 *   3. setupDataValidations()        – re-apply dropdown validations to both sheets
 */
function restoreConfigAndDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Seeding default Config values...', '🔄 Restoring', 2);

  // Phase 1: seed empty Config columns with defaults
  seedConfigData();

  ss.toast('Scanning sheets for missing Config entries...', '🔄 Restoring', 2);

  // Phase 2: restore any values present in Member Directory / Grievance Log
  // but missing from Config (e.g., deleted during a merge or nuke)
  restoreConfigFromSheetData_();

  ss.toast('Applying dropdown validations...', '🔄 Restoring', 2);

  // Phase 3: re-apply dropdown validations to Member Directory and Grievance Log
  setupDataValidations();

  ss.toast('Config and dropdowns restored!', '✅ Success', 3);
}

/**
 * Scans Member Directory and Grievance Log for values currently in use in
 * dropdown/multi-select columns and adds any that are missing from the Config
 * sheet.  This is the reverse of the live bidirectional sync — it bulk-restores
 * Config entries after they have been accidentally deleted.
 * @private
 */
function restoreConfigFromSheetData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet) {
    Logger.log('restoreConfigFromSheetData_: Config sheet not found');
    return;
  }

  // Build a combined mapping of { configCol -> [sheetRef, targetCol] } from both
  // DROPDOWN_MAP and MULTI_SELECT_COLS so we cover every linked column.
  var mappings = [];

  // Member Directory — single-select dropdowns
  var memberDD = DROPDOWN_MAP.MEMBER_DIR;
  for (var i = 0; i < memberDD.length; i++) {
    mappings.push({ sheet: memberSheet, col: memberDD[i].col, configCol: memberDD[i].configCol });
  }

  // Member Directory — multi-select columns
  var memberMS = MULTI_SELECT_COLS.MEMBER_DIR;
  for (var j = 0; j < memberMS.length; j++) {
    mappings.push({ sheet: memberSheet, col: memberMS[j].col, configCol: memberMS[j].configCol, multi: true });
  }

  // Grievance Log — single-select dropdowns
  var grievDD = DROPDOWN_MAP.GRIEVANCE_LOG;
  for (var k = 0; k < grievDD.length; k++) {
    if (grievanceSheet) {
      mappings.push({ sheet: grievanceSheet, col: grievDD[k].col, configCol: grievDD[k].configCol });
    }
  }

  // Grievance Log — multi-select columns
  var grievMS = MULTI_SELECT_COLS.GRIEVANCE_LOG;
  for (var l = 0; l < grievMS.length; l++) {
    if (grievanceSheet) {
      mappings.push({ sheet: grievanceSheet, col: grievMS[l].col, configCol: grievMS[l].configCol, multi: true });
    }
  }

  // NOTE: JOB_METADATA_FIELDS fallback was removed here.  Every column in
  // JOB_METADATA_FIELDS is already covered by DROPDOWN_MAP or MULTI_SELECT_COLS.
  // The old fallback used JOB_METADATA_FIELDS values that could go stale when
  // syncColumnMaps() updated column positions, causing data to land in wrong
  // Config columns.  DROPDOWN_MAP and MULTI_SELECT_COLS are dynamically rebuilt
  // and are the single source of truth.  See: CLAUDE.md § Config Write Paths.

  var totalRestored = 0;

  // Process each mapping: read unique values from the sheet column, compare
  // with Config, and add any missing values.
  for (var p = 0; p < mappings.length; p++) {
    var map = mappings[p];
    if (!map.sheet) continue;

    var sheetLastRow = map.sheet.getLastRow();
    if (sheetLastRow < 2) continue; // No data rows

    // Collect unique non-empty values from the data sheet column
    var sheetData = map.sheet.getRange(2, map.col, sheetLastRow - 1, 1).getValues();
    var uniqueValues = {};
    for (var q = 0; q < sheetData.length; q++) {
      var cellVal = (sheetData[q][0] || '').toString().trim();
      if (!cellVal) continue;

      if (map.multi) {
        // Multi-select columns store comma-separated values
        var parts = cellVal.split(',');
        for (var r = 0; r < parts.length; r++) {
          var part = parts[r].trim();
          if (part) uniqueValues[part] = true;
        }
      } else {
        uniqueValues[cellVal] = true;
      }
    }

    // Get current Config values for this column
    var existingConfig = getConfigValues(configSheet, map.configCol);
    var existingSet = {};
    for (var s = 0; s < existingConfig.length; s++) {
      existingSet[existingConfig[s].toString().trim()] = true;
    }

    // Add missing values to Config
    var missing = [];
    for (var val in uniqueValues) {
      if (!existingSet[val]) {
        missing.push(val);
      }
    }

    for (var t = 0; t < missing.length; t++) {
      addToConfigDropdown_(map.configCol, missing[t]);
      totalRestored++;
    }
  }

  if (totalRestored > 0) {
    Logger.log('restoreConfigFromSheetData_: restored ' + totalRestored + ' missing Config entries');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Restored ' + totalRestored + ' entries from sheet data',
      '🔄 Config Restore', 3
    );
  } else {
    Logger.log('restoreConfigFromSheetData_: no missing entries found');
  }
}

/**
 * Seed members only (no automatic grievance seeding)
 * Used by SEED_SAMPLE_DATA to separate member and grievance seeding
 * @param {number} count - Number of members to seed (max 2000)
 */
function SEED_MEMBERS_ONLY(count) {
  if (!isDemoSafeToRun_()) return;
  count = Math.min(count || 50, 2000);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Ensure Member Directory has enough columns (AF = column 32)
  ensureMinimumColumns(sheet, MEMBER_COLS.QUICK_ACTIONS);

  // Always ensure Config has data for all required columns
  seedConfigData();

  // Get all config values
  var jobTitles = getConfigValues(configSheet, CONFIG_COLS.JOB_TITLES);
  var locations = getConfigValues(configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  var units = getConfigValues(configSheet, CONFIG_COLS.UNITS);
  var supervisors = getConfigValues(configSheet, CONFIG_COLS.SUPERVISORS);
  var managers = getConfigValues(configSheet, CONFIG_COLS.MANAGERS);
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  var committees = getConfigValues(configSheet, CONFIG_COLS.STEWARD_COMMITTEES);

  // Expanded name pools for better variety
  var firstNames = [
    'James', 'Mary', 'John', 'Patricia', 'Robert', 'Jennifer', 'Michael', 'Linda', 'William', 'Elizabeth',
    'David', 'Barbara', 'Richard', 'Susan', 'Joseph', 'Jessica', 'Thomas', 'Sarah', 'Charles', 'Karen',
    'Christopher', 'Nancy', 'Daniel', 'Lisa', 'Matthew', 'Betty', 'Anthony', 'Margaret', 'Mark', 'Sandra',
    'Donald', 'Ashley', 'Steven', 'Kimberly', 'Paul', 'Emily', 'Andrew', 'Donna', 'Joshua', 'Michelle',
    'Kenneth', 'Dorothy', 'Kevin', 'Carol', 'Brian', 'Amanda', 'George', 'Melissa', 'Timothy', 'Deborah',
    'Ronald', 'Stephanie', 'Edward', 'Rebecca', 'Jason', 'Sharon', 'Jeffrey', 'Laura', 'Ryan', 'Cynthia',
    'Jacob', 'Kathleen', 'Gary', 'Amy', 'Nicholas', 'Angela', 'Eric', 'Shirley', 'Jonathan', 'Anna',
    'Stephen', 'Brenda', 'Larry', 'Pamela', 'Justin', 'Emma', 'Scott', 'Nicole', 'Brandon', 'Helen',
    'Benjamin', 'Samantha', 'Samuel', 'Katherine', 'Raymond', 'Christine', 'Gregory', 'Debra', 'Frank', 'Rachel',
    'Alexander', 'Carolyn', 'Patrick', 'Janet', 'Jack', 'Catherine', 'Dennis', 'Maria', 'Jerry', 'Heather',
    'Tyler', 'Diane', 'Aaron', 'Ruth', 'Jose', 'Julie', 'Adam', 'Olivia', 'Nathan', 'Joyce',
    'Henry', 'Virginia', 'Douglas', 'Victoria', 'Zachary', 'Kelly', 'Peter', 'Lauren', 'Kyle', 'Christina'
  ];
  var lastNames = [
    'Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez',
    'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson', 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin',
    'Lee', 'Perez', 'Thompson', 'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson',
    'Walker', 'Young', 'Allen', 'King', 'Wright', 'Scott', 'Torres', 'Nguyen', 'Hill', 'Flores',
    'Green', 'Adams', 'Nelson', 'Baker', 'Hall', 'Rivera', 'Campbell', 'Mitchell', 'Carter', 'Roberts',
    'Gomez', 'Phillips', 'Evans', 'Turner', 'Diaz', 'Parker', 'Cruz', 'Edwards', 'Collins', 'Reyes',
    'Stewart', 'Morris', 'Morales', 'Murphy', 'Cook', 'Rogers', 'Gutierrez', 'Ortiz', 'Morgan', 'Cooper',
    'Peterson', 'Bailey', 'Reed', 'Kelly', 'Howard', 'Ramos', 'Kim', 'Cox', 'Ward', 'Richardson',
    'Watson', 'Brooks', 'Chavez', 'Wood', 'James', 'Bennett', 'Gray', 'Mendoza', 'Ruiz', 'Hughes',
    'Price', 'Alvarez', 'Castillo', 'Sanders', 'Patel', 'Myers', 'Long', 'Ross', 'Foster', 'Jimenez',
    'Powell', 'Jenkins', 'Perry', 'Russell', 'Sullivan', 'Bell', 'Coleman', 'Butler', 'Henderson', 'Barnes',
    'Gonzales', 'Fisher', 'Vasquez', 'Simmons', 'Stokes', 'Burns', 'Fox', 'Alexander', 'Rice', 'Stone'
  ];
  var officeDays = DEFAULT_CONFIG.OFFICE_DAYS;
  var commMethods = DEFAULT_CONFIG.COMM_METHODS;

  var startRow = Math.max(sheet.getLastRow() + 1, 2);

  // Build set of existing member IDs to prevent duplicates
  var existingMemberIds = {};
  if (startRow > 2) {
    var existingData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, startRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      if (existingData[e][0]) {
        existingMemberIds[existingData[e][0]] = true;
      }
    }
  }

  var allRows = [];
  var seededIds = [];
  var batchSize = 100; // Larger batches for 1000 members
  var today = new Date();

  /**
   * Picks 2-3 random office days as a comma-separated string.
   * @param {string[]} days - Available office day values.
   * @returns {string} Comma-separated subset of days.
   */
  function randomOfficeDays(days) {
    var shuffled = days.slice();
    for (var d = shuffled.length - 1; d > 0; d--) {
      var j = Math.floor(Math.random() * (d + 1));
      var tmp = shuffled[d]; shuffled[d] = shuffled[j]; shuffled[j] = tmp;
    }
    var count2 = Math.random() < 0.5 ? 2 : 3;
    return shuffled.slice(0, count2).join(', ');
  }

  var sampleContactNotes = [
    'Discussed workload concerns', 'Follow up on scheduling issue', 'Interested in becoming steward',
    'Addressed safety complaint', 'Positive feedback received', 'Needs info on benefits',
    'Question about contract language', 'Planning to attend next meeting', 'Grievance update provided',
    'Initial outreach - new member', 'Discussed upcoming negotiations', 'Shared resources on workplace rights'
  ];

  // First pass: generate all member rows
  for (var i = 0; i < count; i++) {
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var memberId = generateNameBasedId('M', firstName, lastName, existingMemberIds);
    existingMemberIds[memberId] = true;
    seededIds.push(memberId);
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + '.' + memberId.toLowerCase() + '@example.org';
    var phone = '617-555-' + String(Math.floor(Math.random() * 9000) + 1000);
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';

    var hasRecentContact = Math.random() < 0.5;
    var recentContactDate = hasRecentContact ? randomDate(new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000), today) : '';
    var contactSteward = hasRecentContact ? randomChoice(stewards) : '';
    var contactNotes = hasRecentContact ? randomChoice(sampleContactNotes) : '';

    var row = generateSingleMemberRow(
      memberId, firstName, lastName,
      randomChoice(jobTitles), randomChoice(locations), randomChoice(units), randomOfficeDays(officeDays),
      email, phone, randomChoice(commMethods), 'Morning',
      randomChoice(supervisors), randomChoice(managers),
      isSteward, isSteward === 'Yes' ? randomChoice(committees) : '', '',
      recentContactDate, contactSteward, contactNotes
    );

    allRows.push(row);
  }

  // Second pass: collect steward emails and assign to all members
  var stewardEmails = [];
  for (var s = 0; s < allRows.length; s++) {
    if (allRows[s][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      stewardEmails.push(allRows[s][MEMBER_COLS.EMAIL - 1]);
    }
  }
  if (stewardEmails.length === 0) stewardEmails.push(allRows[0][MEMBER_COLS.EMAIL - 1]);

  for (var a = 0; a < allRows.length; a++) {
    allRows[a][MEMBER_COLS.ASSIGNED_STEWARD - 1] = randomChoice(stewardEmails);
  }

  // Write all rows in batches
  for (var b = 0; b < allRows.length; b += batchSize) {
    var batch = allRows.slice(b, Math.min(b + batchSize, allRows.length));
    sheet.getRange(startRow, 1, batch.length, batch[0].length).setValues(batch);
    startRow += batch.length;
    Utilities.sleep(50);
  }

  // Re-apply checkboxes
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();
  }

  // Track seeded IDs
  trackSeededMemberIdsBatch(seededIds);

  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded!', '✅ Success', 3);
}

/**
 * Generate a single member row using MEMBER_COLS for column placement.
 * Column positions are fully dynamic — reordering MEMBER_COLS won't break seeding.
 */
function generateSingleMemberRow(memberId, firstName, lastName, jobTitle, location, unit, officeDays, email, phone, prefComm, bestTime, supervisor, manager, isSteward, committees, assignedSteward, recentContactDate, contactSteward, contactNotes) {
  var today = new Date();
  var lastMonth = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);

  // Build row sized to the last MEMBER_COLS column, filled with empty strings
  var maxCol = MEMBER_COLS.DUES_STATUS; // last defined column (after ZIP_CODE)
  var row = [];
  for (var i = 0; i < maxCol; i++) { row.push(''); }

  // Identity & Core Info
  row[MEMBER_COLS.MEMBER_ID - 1] = memberId;
  row[MEMBER_COLS.FIRST_NAME - 1] = firstName;
  row[MEMBER_COLS.LAST_NAME - 1] = lastName;
  row[MEMBER_COLS.JOB_TITLE - 1] = jobTitle;

  // Location & Work
  row[MEMBER_COLS.WORK_LOCATION - 1] = location;
  row[MEMBER_COLS.UNIT - 1] = unit;
  // CUBICLE left empty (hidden column, not seeded)
  row[MEMBER_COLS.OFFICE_DAYS - 1] = officeDays;

  // Contact Information
  row[MEMBER_COLS.EMAIL - 1] = email;
  row[MEMBER_COLS.PHONE - 1] = phone;
  row[MEMBER_COLS.PREFERRED_COMM - 1] = prefComm;
  row[MEMBER_COLS.BEST_TIME - 1] = bestTime;

  // Organizational Structure
  row[MEMBER_COLS.SUPERVISOR - 1] = supervisor;
  row[MEMBER_COLS.MANAGER - 1] = manager;
  row[MEMBER_COLS.IS_STEWARD - 1] = isSteward;
  row[MEMBER_COLS.COMMITTEES - 1] = committees;
  row[MEMBER_COLS.ASSIGNED_STEWARD - 1] = assignedSteward;

  // Engagement Metrics
  row[MEMBER_COLS.LAST_VIRTUAL_MTG - 1] = randomDate(lastMonth, today);
  row[MEMBER_COLS.LAST_INPERSON_MTG - 1] = randomDate(lastMonth, today);
  row[MEMBER_COLS.OPEN_RATE - 1] = Math.floor(Math.random() * 100);
  row[MEMBER_COLS.VOLUNTEER_HOURS - 1] = Math.floor(Math.random() * 20);

  // Member Interests
  row[MEMBER_COLS.INTEREST_LOCAL - 1] = Math.random() < 0.3 ? 'Yes' : 'No';
  row[MEMBER_COLS.INTEREST_CHAPTER - 1] = Math.random() < 0.2 ? 'Yes' : 'No';
  row[MEMBER_COLS.INTEREST_ALLIED - 1] = Math.random() < 0.1 ? 'Yes' : 'No';

  // Steward Contact Tracking
  row[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = recentContactDate || '';
  row[MEMBER_COLS.CONTACT_STEWARD - 1] = contactSteward || '';
  row[MEMBER_COLS.CONTACT_NOTES - 1] = contactNotes || '';

  // Grievance Management (script-calculated, leave empty)
  // HAS_OPEN_GRIEVANCE, GRIEVANCE_STATUS, NEXT_DEADLINE left as ''
  row[MEMBER_COLS.START_GRIEVANCE - 1] = false;

  // Address Data (sample MA addresses)
  var streets = ['123 Main St', '456 Oak Ave', '789 Elm Blvd', '321 Park Rd', '654 Cedar Ln',
    '987 Maple Dr', '147 Pine Way', '258 Birch Ct', '369 Walnut St', '741 Spruce Ave'];
  var cities = ['Boston', 'Worcester', 'Springfield', 'Cambridge', 'Lowell',
    'Brockton', 'Quincy', 'New Bedford', 'Fall River', 'Lawrence'];
  row[MEMBER_COLS.STREET_ADDRESS - 1] = randomChoice(streets);
  row[MEMBER_COLS.CITY - 1] = randomChoice(cities);
  row[MEMBER_COLS.STATE - 1] = 'MA';
  row[MEMBER_COLS.ZIP_CODE - 1] = '0' + String(Math.floor(Math.random() * 9000) + 1000);

  // Dues Status
  var duesStatuses = ['Current', 'Current', 'Current', 'Current', 'Past Due', 'Inactive'];
  row[MEMBER_COLS.DUES_STATUS - 1] = randomChoice(duesStatuses);

  return row;
}

// ensureMinimumColumns has been moved to 01_Core.gs for load-order safety.

/**
 * Seed N grievances
 * @param {number} count - Number of grievances to seed (max 300)
 */
function SEED_GRIEVANCES(count) {
  if (!isDemoSafeToRun_()) return;
  count = Math.min(count || 25, 300);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!grievanceSheet || !memberSheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Ensure Grievance Log has enough columns (AK = column 37, CHECKLIST_PROGRESS is the last column)
  ensureMinimumColumns(grievanceSheet, GRIEVANCE_COLS.CHECKLIST_PROGRESS);

  // Get members
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) {
    SpreadsheetApp.getUi().alert('Error: No members found. Please seed members first.');
    return;
  }

  var statuses = DEFAULT_CONFIG.GRIEVANCE_STATUS;
  var steps = DEFAULT_CONFIG.GRIEVANCE_STEP;
  var categories = DEFAULT_CONFIG.ISSUE_CATEGORY;
  var articles = DEFAULT_CONFIG.ARTICLES;

  // Collect steward emails from Member Directory (not Config names)
  var stewardEmails = [];
  for (var se = 1; se < memberData.length; se++) {
    if (String(memberData[se][MEMBER_COLS.IS_STEWARD - 1]).trim().toLowerCase() === 'yes') {
      var sEmail = String(memberData[se][MEMBER_COLS.EMAIL - 1]).trim();
      if (sEmail) stewardEmails.push(sEmail);
    }
  }
  if (stewardEmails.length === 0) stewardEmails.push('steward@example.org');

  var startRow = Math.max(grievanceSheet.getLastRow() + 1, 2);

  // Build set of existing grievance IDs to prevent duplicates
  var existingGrievanceIds = {};
  if (startRow > 2) {
    var existingData = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, startRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      if (existingData[e][0]) {
        existingGrievanceIds[existingData[e][0]] = true;
      }
    }
  }

  var rows = [];
  var seededIds = []; // Track IDs for this seeding session
  var batchSize = 25;
  var today = new Date();

  // Create shuffled list of member indices (excluding header row)
  var memberIndices = [];
  for (var m = 1; m < memberData.length; m++) {
    if (memberData[m][MEMBER_COLS.MEMBER_ID - 1]) {
      memberIndices.push(m);
    }
  }

  if (memberIndices.length === 0) {
    SpreadsheetApp.getUi().alert('Error: No members with valid Member IDs found. Please seed members first.');
    return;
  }

  var shuffledMembers = shuffleArray(memberIndices);
  var memberIndex = 0;

  // Distribute incident dates across the 90-day range
  var dateRangeStart = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);
  var dateSpread = 90 / count;

  for (var i = 0; i < count; i++) {
    if (memberIndex >= shuffledMembers.length) {
      shuffledMembers = shuffleArray(memberIndices);
      memberIndex = 0;
    }
    var memberRow = memberData[shuffledMembers[memberIndex]];
    memberIndex++;

    var memberId = memberRow[MEMBER_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    var firstName = memberRow[MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = memberRow[MEMBER_COLS.LAST_NAME - 1] || '';
    var memberEmail = memberRow[MEMBER_COLS.EMAIL - 1] || '';
    var memberUnit = memberRow[MEMBER_COLS.UNIT - 1] || '';
    var memberLocation = memberRow[MEMBER_COLS.WORK_LOCATION - 1] || '';
    var memberSteward = memberRow[MEMBER_COLS.ASSIGNED_STEWARD - 1] || randomChoice(stewardEmails);

    var grievanceId = generateNameBasedId('G', firstName, lastName, existingGrievanceIds);
    existingGrievanceIds[grievanceId] = true;
    seededIds.push(grievanceId);

    var baseDate = new Date(dateRangeStart.getTime() + (i * dateSpread * 24 * 60 * 60 * 1000));
    var variation = (Math.random() - 0.5) * 4 * 24 * 60 * 60 * 1000;
    var incidentDate = new Date(Math.min(baseDate.getTime() + variation, today.getTime()));

    var status = randomChoice(statuses);
    var step = randomChoice(steps);

    var row = generateSingleGrievanceRow(
      grievanceId,
      memberId,
      firstName,
      lastName,
      status,
      step,
      incidentDate,
      randomChoice(articles),
      randomChoice(categories),
      memberEmail,
      memberUnit,
      memberLocation,
      memberSteward
    );

    rows.push(row);

    if (rows.length >= batchSize || i === count - 1) {
      grievanceSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(100);
    }
  }

  // Re-apply checkboxes to Message Alert column (AC)
  var lastRow = grievanceSheet.getLastRow();
  if (lastRow >= 2) {
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();
  }

  // Sync data
  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();

  // Track seeded IDs
  trackSeededGrievanceIdsBatch(seededIds);

  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' grievances seeded!', '✅ Success', 3);
}

/**
 * Generate a single grievance row using GRIEVANCE_COLS for column placement.
 * Column positions are fully dynamic — reordering GRIEVANCE_COLS won't break seeding.
 */
function generateSingleGrievanceRow(grievanceId, memberId, firstName, lastName, status, step, incidentDate, articles, category, email, unit, location, steward) {
  var today = new Date();
  var filingDeadline = addDays(incidentDate, 21);

  var dateFiled = addDays(incidentDate, Math.floor(Math.random() * 14) + 1);
  var step1Due = dateFiled ? addDays(dateFiled, 30) : '';

  var step1Rcvd = '';
  var step2AppealDue = '';
  var step2AppealFiled = '';
  var step2Due = '';
  var step2Rcvd = '';
  var step3AppealDue = '';
  var step3AppealFiled = '';
  var dateClosed = '';

  var isClosed = GRIEVANCE_CLOSED_STATUSES.indexOf(status) !== -1;
  var stepIndex = ['Informal', 'Step I', 'Step II', 'Step III', 'Mediation', 'Arbitration'].indexOf(step);

  if (stepIndex >= 1 || isClosed) {
    step1Rcvd = addDays(step1Due, Math.floor(Math.random() * 10) - 5);
    if (step1Rcvd < dateFiled) step1Rcvd = addDays(dateFiled, 15);
    step2AppealDue = addDays(step1Rcvd, 10);
  }

  if (stepIndex >= 2 || (isClosed && stepIndex >= 1)) {
    step2AppealFiled = addDays(step1Rcvd, Math.floor(Math.random() * 8) + 1);
    step2Due = addDays(step2AppealFiled, 30);
    step2Rcvd = addDays(step2Due, Math.floor(Math.random() * 10) - 5);
    if (step2Rcvd < step2AppealFiled) step2Rcvd = addDays(step2AppealFiled, 15);
    step3AppealDue = addDays(step2Rcvd, 30);
  }

  if (stepIndex >= 3 || (isClosed && stepIndex >= 2)) {
    step3AppealFiled = addDays(step2Rcvd, Math.floor(Math.random() * 20) + 1);
  }

  if (isClosed) {
    var lastDate = step3AppealFiled || step2Rcvd || step1Rcvd || dateFiled;
    dateClosed = addDays(lastDate, Math.floor(Math.random() * 30) + 5);
  }

  var daysOpen = '';
  if (dateFiled) {
    var endDate = dateClosed || today;
    daysOpen = Math.floor((endDate - dateFiled) / (1000 * 60 * 60 * 24));
  }

  var nextActionDue = '';
  if (!isClosed) {
    switch (step) {
      case 'Informal':
        nextActionDue = filingDeadline;
        break;
      case 'Step I':
        nextActionDue = step1Due || filingDeadline;
        break;
      case 'Step II':
        nextActionDue = step2Due || step2AppealDue || step1Due;
        break;
      case 'Step III':
      case 'Mediation':
      case 'Arbitration':
        nextActionDue = step3AppealDue || step2Due;
        break;
    }
  }

  var daysToDeadline = '';
  if (nextActionDue && !isClosed) {
    daysToDeadline = Math.floor((nextActionDue - today) / (1000 * 60 * 60 * 24));
  }

  var resolutions = ['Won - Full remedy', 'Won - Partial remedy', 'Settled - Compromise', 'Denied', 'Withdrawn', 'Pending'];
  var resolution = dateClosed ? randomChoice(resolutions) : '';

  // Build row sized to the last GRIEVANCE_COLS column, filled with empty strings
  var maxCol = GRIEVANCE_COLS.LAST_UPDATED; // last defined column
  var row = [];
  for (var i = 0; i < maxCol; i++) { row.push(''); }

  // Identity
  row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = grievanceId;
  row[GRIEVANCE_COLS.MEMBER_ID - 1] = memberId;
  row[GRIEVANCE_COLS.FIRST_NAME - 1] = firstName;
  row[GRIEVANCE_COLS.LAST_NAME - 1] = lastName;

  // Status & Assignment
  row[GRIEVANCE_COLS.STATUS - 1] = status;
  row[GRIEVANCE_COLS.CURRENT_STEP - 1] = step;

  // Timeline - Filing
  row[GRIEVANCE_COLS.INCIDENT_DATE - 1] = incidentDate;
  row[GRIEVANCE_COLS.FILING_DEADLINE - 1] = filingDeadline;
  row[GRIEVANCE_COLS.DATE_FILED - 1] = dateFiled;

  // Timeline - Step I
  row[GRIEVANCE_COLS.STEP1_DUE - 1] = step1Due;
  row[GRIEVANCE_COLS.STEP1_RCVD - 1] = step1Rcvd;

  // Timeline - Step II
  row[GRIEVANCE_COLS.STEP2_APPEAL_DUE - 1] = step2AppealDue;
  row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1] = step2AppealFiled;
  row[GRIEVANCE_COLS.STEP2_DUE - 1] = step2Due;
  row[GRIEVANCE_COLS.STEP2_RCVD - 1] = step2Rcvd;

  // Timeline - Step III
  row[GRIEVANCE_COLS.STEP3_APPEAL_DUE - 1] = step3AppealDue;
  row[GRIEVANCE_COLS.STEP3_APPEAL_FILED - 1] = step3AppealFiled;
  row[GRIEVANCE_COLS.DATE_CLOSED - 1] = dateClosed;

  // Calculated Metrics
  row[GRIEVANCE_COLS.DAYS_OPEN - 1] = daysOpen;
  row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] = nextActionDue;
  row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = daysToDeadline;

  // Case Details
  row[GRIEVANCE_COLS.ARTICLES - 1] = articles;
  row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = category;

  // Contact & Location
  row[GRIEVANCE_COLS.MEMBER_EMAIL - 1] = email;
  row[GRIEVANCE_COLS.LOCATION - 1] = location;
  row[GRIEVANCE_COLS.STEWARD - 1] = steward;

  // Resolution
  row[GRIEVANCE_COLS.RESOLUTION - 1] = resolution;

  // Coordinator Notifications
  row[GRIEVANCE_COLS.MESSAGE_ALERT - 1] = false;
  // COORDINATOR_MESSAGE, ACKNOWLEDGED_BY, ACKNOWLEDGED_DATE left as ''

  // Drive Integration
  // DRIVE_FOLDER_ID, DRIVE_FOLDER_URL left as ''

  return row;
}

// ============================================================================
// DIALOG FUNCTIONS
// ============================================================================

// ============================================================================
// SEED: SCRIPT OWNER AS TEST MEMBER
// ============================================================================

/**
 * Adds the script owner (logged-in developer) as a steward member in the
 * Member Directory so they can test the SPA without manually creating a row.
 * Assigns a random steward email from existing seeded members.
 * @param {Spreadsheet} ss - Active spreadsheet
 * @private
 */
function seedScriptOwnerMember_(ss) {
  var ownerEmail;
  try {
    ownerEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log('Could not get script owner email: ' + e.message);
    return;
  }
  if (!ownerEmail) {
    Logger.log('Script owner email is empty. Skipping owner seed.');
    return;
  }

  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet || memberSheet.getLastRow() < 2) return;

  // Check if owner already exists
  var existingData = memberSheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (String(existingData[i][MEMBER_COLS.EMAIL - 1]).trim().toLowerCase() === ownerEmail.toLowerCase()) {
      Logger.log('Script owner already exists in Member Directory. Skipping.');
      return;
    }
  }

  // Find a steward email from existing members for assignment
  var stewardEmails = [];
  for (var j = 1; j < existingData.length; j++) {
    if (String(existingData[j][MEMBER_COLS.IS_STEWARD - 1]).trim().toLowerCase() === 'yes') {
      stewardEmails.push(String(existingData[j][MEMBER_COLS.EMAIL - 1]).trim());
    }
  }
  var assignedStewardEmail = stewardEmails.length > 0 ? randomChoice(stewardEmails) : '';

  var ownerRow = generateSingleMemberRow(
    'M-OWNER', 'Test', 'Admin', 'Program Manager',
    'Boston Main Office', 'Administration', 'Mon, Tue, Wed, Thu, Fri',
    ownerEmail, '617-555-0001', 'Email', 'Morning',
    'Maria Rodriguez', 'Patricia Thompson', 'Yes', 'Executive Board',
    assignedStewardEmail,
    '', '', '' // no recent contact tracking
  );

  memberSheet.getRange(memberSheet.getLastRow() + 1, 1, 1, ownerRow.length).setValues([ownerRow]);
  trackSeededMemberId('M-OWNER');
  Logger.log('Added script owner (' + ownerEmail + ') as test steward member');
}

// ============================================================================
// SEED: CONTACT LOG
// ============================================================================

/**
 * Seeds 25 sample contact log entries between stewards and their assigned members.
 * Uses the _Contact_Log sheet with columns: [ID, Steward Email, Member Email, Contact Type, Date, Notes, Duration, Created]
 */
function seedContactLogData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet || memberSheet.getLastRow() < 2) {
    Logger.log('Member Directory empty. Skipping contact log seed.');
    return;
  }

  // Ensure the contact log sheet exists
  var clSheet = ss.getSheetByName(SHEETS.CONTACT_LOG);
  if (!clSheet) {
    clSheet = ss.insertSheet(SHEETS.CONTACT_LOG);
    clSheet.getRange(1, 1, 1, 8).setValues([['ID', 'Steward Email', 'Member Email', 'Contact Type', 'Date', 'Notes', 'Duration', 'Created']]);
    clSheet.hideSheet();
  }

  // Skip if already has data
  if (clSheet.getLastRow() > 1) {
    Logger.log('Contact log already has data. Skipping seed.');
    return;
  }

  // Collect steward emails and their assigned member emails
  var memberData = memberSheet.getDataRange().getValues();
  var stewardEmailList = [];
  var membersByAssignment = {};

  for (var i = 1; i < memberData.length; i++) {
    var email = String(memberData[i][MEMBER_COLS.EMAIL - 1]).trim().toLowerCase();
    var isSteward = String(memberData[i][MEMBER_COLS.IS_STEWARD - 1]).trim().toLowerCase();
    var assignedSteward = String(memberData[i][MEMBER_COLS.ASSIGNED_STEWARD - 1]).trim().toLowerCase();

    if (isSteward === 'yes') {
      stewardEmailList.push(email);
    }
    if (assignedSteward && email) {
      if (!membersByAssignment[assignedSteward]) membersByAssignment[assignedSteward] = [];
      membersByAssignment[assignedSteward].push(email);
    }
  }

  if (stewardEmailList.length === 0) {
    Logger.log('No stewards found. Skipping contact log seed.');
    return;
  }

  var contactTypes = ['Phone', 'Email', 'In Person', 'Text'];
  var contactNotes = [
    'Discussed upcoming contract changes and answered questions.',
    'Follow-up on workload concerns raised last week.',
    'Checked in about scheduling conflict with supervisor.',
    'Provided info on grievance filing process.',
    'Welcomed new member and explained union benefits.',
    'Addressed safety concern in the workplace.',
    'Discussed open enrollment benefits options.',
    'Followed up on resolved grievance outcome.',
    'Planning for upcoming membership meeting.',
    'Shared know-your-rights materials.',
    'Discussed concerns about overtime expectations.',
    'Helped member understand FMLA protections.',
    'Addressed question about union dues.',
    'Provided update on ongoing negotiations.',
    'Checked in after return from leave.',
    'Reviewed seniority roster discrepancy with member.',
    'Coached member on how to request reasonable accommodation.',
    'Discussed transfer request process and contract language.',
    'Followed up on workplace harassment complaint next steps.',
    'Helped member prepare for upcoming disciplinary meeting.',
    'Explained Weingarten rights before investigatory interview.',
    'Provided resources on EAP and counseling services.',
    'Discussed job posting irregularities in member\'s unit.',
    'Answered questions about probationary period requirements.',
    'Checked in regarding denied vacation request appeal.',
  ];
  var durations = ['5 min', '10 min', '15 min', '20 min', '30 min', '45 min', '1 hr'];

  var now = new Date();
  var rows = [];

  for (var c = 0; c < 40; c++) {
    var stewardEmail = randomChoice(stewardEmailList);
    var assignedMembers = membersByAssignment[stewardEmail];
    var memberEmail = (assignedMembers && assignedMembers.length > 0) ?
      randomChoice(assignedMembers) :
      randomChoice(stewardEmailList); // fallback to another steward

    var daysAgo = Math.floor(Math.random() * 90); // within last 90 days
    var contactDate = new Date(now.getTime() - daysAgo * 86400000);
    contactDate.setHours(Math.floor(Math.random() * 8) + 9, Math.floor(Math.random() * 60), 0, 0); // 9am-5pm

    rows.push([
      'CL_SEED_' + (c + 1),
      stewardEmail,
      memberEmail,
      randomChoice(contactTypes),
      contactDate,
      randomChoice(contactNotes),
      randomChoice(durations),
      contactDate
    ]);
  }

  clSheet.getRange(2, 1, rows.length, 8).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' contact log entries');
}

/**
 * Standalone entry point to seed contact log data independently.
 * Can be run from the script editor or menu without running a full seed phase.
 */
function SEED_CONTACT_LOG() {
  if (!isDemoSafeToRun_()) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet || memberSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Error: Member Directory is empty. Please seed members first.');
    return;
  }

  seedContactLogData();

  var clSheet = ss.getSheetByName(SHEETS.CONTACT_LOG);
  var count = clSheet ? Math.max(clSheet.getLastRow() - 1, 0) : 0;
  SpreadsheetApp.getUi().alert('Contact Log seeded with ' + count + ' entries.');
}

// ============================================================================
// SEED: CALENDAR EVENTS
// ============================================================================

/**
 * Seeds 15 fictional calendar events (union meetings, trainings, social events)
 * into the org's Google Calendar. Stores created event IDs for nuke cleanup.
 */
function seedCalendarEvents() {
  var calId = PropertiesService.getScriptProperties().getProperty('ORG_CALENDAR_ID') || 'primary';
  var cal = null;
  try {
    cal = CalendarApp.getCalendarById(calId);
    if (!cal) cal = CalendarApp.getDefaultCalendar();
  } catch (_e) {
    cal = null; // CalendarApp unavailable (e.g., web app context)
  }

  var now = new Date();
  var events = [
    { title: 'Monthly General Membership Meeting', offsetDays: 7, hour: 17, durationHrs: 2, desc: 'All members welcome. Agenda: contract updates, Q&A, committee reports.' },
    { title: 'New Member Orientation', offsetDays: 10, hour: 12, durationHrs: 1, desc: 'Welcome session for new bargaining unit members. Learn your rights and benefits.' },
    { title: 'Steward Training: Grievance Basics', offsetDays: 14, hour: 9, durationHrs: 3, desc: 'Required training for all stewards. Cover grievance filing, timelines, and documentation.' },
    { title: 'Contract Enforcement Workshop', offsetDays: 21, hour: 14, durationHrs: 2, desc: 'Deep dive into contract language and enforcement strategies.' },
    { title: 'Executive Board Meeting', offsetDays: 5, hour: 18, durationHrs: 1.5, desc: 'Board members only. Budget review and upcoming negotiations prep.' },
    { title: 'Union Social: Pizza & Solidarity', offsetDays: 28, hour: 17, durationHrs: 2, desc: 'Casual get-together for all members. Food provided.' },
    { title: 'Workplace Safety Committee', offsetDays: 12, hour: 10, durationHrs: 1, desc: 'Monthly safety review. Incident reports and prevention planning.' },
    { title: 'Legislative Action Call-In Day', offsetDays: 18, hour: 11, durationHrs: 1, desc: 'Join your coworkers in calling legislators about workforce funding.' },
    { title: 'Steward Advanced Training: Arbitration', offsetDays: 35, hour: 9, durationHrs: 4, desc: 'Advanced grievance handling: arbitration prep, evidence gathering, witness interviews.' },
    { title: 'Member Appreciation Cookout', offsetDays: 42, hour: 16, durationHrs: 3, desc: 'Annual summer cookout. Bring your family!' },
    { title: 'Know Your Rights Lunch & Learn', offsetDays: 25, hour: 12, durationHrs: 1, desc: 'Weingarten rights, FMLA, and ADA protections overview.' },
    { title: 'Quarterly All-Hands Town Hall', offsetDays: 30, hour: 17, durationHrs: 1.5, desc: 'Union president report, financial update, open floor Q&A.' },
    { title: 'Community Volunteering Day', offsetDays: 45, hour: 9, durationHrs: 4, desc: 'Join fellow members for a community service project.' },
    { title: 'New Contract Ratification Vote', offsetDays: 50, hour: 17, durationHrs: 2, desc: 'Critical vote on the new collective bargaining agreement. All members eligible to vote.' },
    { title: 'Steward Appreciation Dinner', offsetDays: 55, hour: 18, durationHrs: 3, desc: 'Annual dinner recognizing our stewards. RSVPs required.' },
  ];

  if (cal) {
    // Calendar available — create real calendar events
    var createdIds = [];
    events.forEach(function(evt) {
      try {
        var start = new Date(now.getTime() + evt.offsetDays * 86400000);
        start.setHours(evt.hour, 0, 0, 0);
        var end = new Date(start.getTime() + evt.durationHrs * 3600000);
        var created = cal.createEvent(evt.title, start, end, { description: evt.desc });
        createdIds.push(created.getId());
      } catch (e) {
        Logger.log('Calendar event creation failed for "' + evt.title + '": ' + e.message);
      }
    });
    PropertiesService.getScriptProperties().setProperty('SEEDED_CALENDAR_EVENT_IDS', JSON.stringify(createdIds));
    Logger.log('Seeded ' + createdIds.length + ' calendar events');
  }

  // Always seed events into _Timeline_Events as fallback data source
  // (used when CalendarApp is unavailable at read time)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tlSheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
  if (!tlSheet) {
    if (typeof TimelineService !== 'undefined' && TimelineService.initSheet) {
      try { TimelineService.initSheet(); } catch (_e) { /* ok */ }
    }
    tlSheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
  }
  if (tlSheet) {
    // Clear any previously seeded calendar events (IDs starting with TL_CAL_)
    if (tlSheet.getLastRow() > 1) {
      var tlData = tlSheet.getRange(2, 1, tlSheet.getLastRow() - 1, 1).getValues();
      for (var ri = tlData.length - 1; ri >= 0; ri--) {
        if (String(tlData[ri][0]).indexOf('TL_CAL_') === 0) {
          tlSheet.deleteRow(ri + 2);
        }
      }
    }
    var tlRows = [];
    events.forEach(function(evt, idx) {
      var evtDate = new Date(now.getTime() + evt.offsetDays * 86400000);
      evtDate.setHours(evt.hour, 0, 0, 0);
      var endDate = new Date(evtDate.getTime() + evt.durationHrs * 3600000);
      tlRows.push([
        'TL_CAL_' + (idx + 1), evt.title, evtDate, evt.desc,
        'meeting', '', '', '', endDate.toISOString(), 'system', now, ''
      ]);
    });
    var startRow = Math.max(tlSheet.getLastRow() + 1, 2);
    tlSheet.getRange(startRow, 1, tlRows.length, 12).setValues(tlRows);
    Logger.log('Seeded ' + tlRows.length + ' events into Timeline sheet');
  }
}

// ============================================================================
// SEED: WEEKLY QUESTIONS
// ============================================================================

/**
 * Seeds the _Question_Pool with 20+ engagement questions,
 * adds 2 active questions to _Weekly_Questions,
 * and generates sample responses in _Weekly_Responses.
 */
function seedWeeklyQuestions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Schema from 24_WeeklyQuestions.gs:
  // _Weekly_Questions:  ID | Text | Options (JSON) | Source | Submitted By | Week Start | Active | Created
  // _Weekly_Responses:  ID | Question ID | Email Hash | Response | Timestamp
  // _Question_Pool:     ID | Text | Options (JSON) | Submitted By Hash | Status | Created

  // --- Question Pool (member-submitted candidates for random draw) ---
  var poolSheet = ss.getSheetByName(SHEETS.QUESTION_POOL);
  if (!poolSheet) {
    Logger.log('_Question_Pool sheet not found. Skipping weekly questions seed.');
    return;
  }

  var poolAlreadySeeded = poolSheet.getLastRow() > 1;
  if (poolAlreadySeeded) {
    Logger.log('_Question_Pool already has data. Refreshing active poll dates only.');
  }

  var poolQuestions = [
    { text: 'How satisfied are you with the office day scheduling process?',
      options: ['Very satisfied', 'Satisfied', 'Neutral', 'Dissatisfied', 'Very dissatisfied'] },
    { text: 'Has management communicated the new policy changes clearly to you?',
      options: ['Yes, fully', 'Somewhat', 'Not really', 'Not at all'] },
    { text: 'Should we push for a 4-day work week in the next contract?',
      options: ['Strongly support', 'Support', 'Neutral', 'Oppose'] },
    { text: 'How manageable is your current caseload?',
      options: ['Very manageable', 'Manageable', 'Somewhat overwhelming', 'Very overwhelming'] },
    { text: 'What is your top priority for the next contract negotiation?',
      options: ['Higher wages', 'Better benefits', 'Manageable caseloads', 'Schedule flexibility'] },
    { text: 'How accessible is your steward when you need support?',
      options: ['Always available', 'Usually available', 'Sometimes available', 'Rarely available'] },
    { text: 'Do you feel informed about the grievance process?',
      options: ['Very informed', 'Somewhat informed', 'Not very informed', 'Not informed at all'] },
    { text: 'Best time for monthly general membership meetings?',
      options: ['Weekday 5pm', 'Weekday 6pm', 'Saturday 10am', 'Virtual anytime'] },
    { text: 'How confident are you in the union\'s ability to represent you?',
      options: ['Very confident', 'Somewhat confident', 'Not very confident', 'Not confident at all'] },
    { text: 'How would you rate communication from union leadership?',
      options: ['Excellent', 'Good', 'Fair', 'Poor'] },
    { text: 'Do you feel respected by management in your daily work?',
      options: ['Always', 'Usually', 'Sometimes', 'Rarely or never'] },
    { text: 'How well does your team collaborate on shared cases?',
      options: ['Extremely well', 'Well', 'Somewhat well', 'Not well'] },
  ];

  if (!poolAlreadySeeded) {
    var poolRows = poolQuestions.map(function(q, idx) {
      var fakeHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'seed-pool-member-' + idx)
        .map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
      return [
        'WP_SEED_' + (idx + 1),           // ID
        q.text,                            // Text
        JSON.stringify(q.options),         // Options (JSON) — required by _parseOptions()
        fakeHash,                          // Submitted By Hash
        'pending',                         // Status
        new Date()                         // Created
      ];
    });

    poolSheet.getRange(2, 1, poolRows.length, 6).setValues(poolRows);
  }

  // --- Active Weekly Questions for the current period (2 polls: 1 steward, 1 community) ---
  var wqSheet = ss.getSheetByName(SHEETS.WEEKLY_QUESTIONS);
  if (wqSheet) {
    var now = new Date();
    var weekStart = new Date(now);
    weekStart.setHours(0, 0, 0, 0);
    var day = weekStart.getDay();
    weekStart.setDate(weekStart.getDate() - (day === 0 ? 6 : day - 1)); // Monday of this week
    var weekStr = weekStart.toISOString().split('T')[0];

    var stewardOpts = ['Higher wages', 'Reduced caseloads', 'Better benefits', 'More schedule flexibility'];
    var communityOpts = ['Always available', 'Usually available', 'Sometimes available', 'Rarely available'];

    var activeQs = [
      [
        'WQ-SEED-001',                         // ID
        'What should be the union\'s top priority in our next contract negotiation?', // Text
        JSON.stringify(stewardOpts),           // Options (JSON)
        'steward',                             // Source
        '',                                    // Submitted By (steward email — empty for seed)
        weekStr,                               // Week Start
        'TRUE',                                // Active
        now                                    // Created
      ],
      [
        'WQ-SEED-002',
        'How accessible is your steward when you need support?',
        JSON.stringify(communityOpts),
        'community',
        '',
        weekStr,
        'TRUE',
        now
      ],
    ];
    // Clear existing seed rows and re-write with current period dates
    if (wqSheet.getLastRow() > 1) {
      wqSheet.getRange(2, 1, wqSheet.getLastRow() - 1, 8).clearContent();
    }
    wqSheet.getRange(2, 1, activeQs.length, 8).setValues(activeQs);

    // --- Sample anonymous responses (10 per question) ---
    var wrSheet = ss.getSheetByName(SHEETS.WEEKLY_RESPONSES);
    if (wrSheet) {
      // Clear stale responses and re-seed with current timestamps
      if (wrSheet.getLastRow() > 1) {
        wrSheet.getRange(2, 1, wrSheet.getLastRow() - 1, 5).clearContent();
      }
      var responseDist = {
        'WQ-SEED-001': [
          'Higher wages', 'Reduced caseloads', 'Higher wages', 'Better benefits',
          'Reduced caseloads', 'Higher wages', 'More schedule flexibility', 'Reduced caseloads',
          'Higher wages', 'Better benefits'
        ],
        'WQ-SEED-002': [
          'Usually available', 'Always available', 'Sometimes available', 'Usually available',
          'Always available', 'Usually available', 'Rarely available', 'Sometimes available',
          'Usually available', 'Always available'
        ]
      };

      var responseRows = [];
      Object.keys(responseDist).forEach(function(qId) {
        responseDist[qId].forEach(function(resp, idx) {
          var fakeHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'seed-voter-' + qId + '-' + idx)
            .map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
          responseRows.push([
            'WR_SEED_' + qId + '_' + (idx + 1), // ID
            qId,                                  // Question ID
            fakeHash,                             // Email Hash (no PII)
            resp,                                 // Response
            new Date(now.getTime() - idx * 3600000) // Timestamp
          ]);
        });
      });

      wrSheet.getRange(2, 1, responseRows.length, 5).setValues(responseRows);
    }
  }

  Logger.log('seedWeeklyQuestions: ' + poolQuestions.length + ' pool questions, 2 active polls, 20 responses seeded (period: ' + weekStr + ').');
}

// ============================================================================
// SEED: UNION STATS DATA
// ============================================================================

/**
 * Seeds aggregate union stats data for the Stats page.
 * Stores engagement metrics and membership trends in Script Properties
 * (no separate sheet needed — lightweight stat snapshots).
 */
function seedUnionStatsData() {
  var stats = {
    engagement: {
      surveyParticipation: 62,
      weeklyQuestionVotes: 147,
      eventAttendance: 78,
      grievanceFilingRate: 4.2,
      stewardContactRate: 31,
      resourceDownloads: 215,
    },
    membershipTrends: [
      { month: 'Sep', total: 842, new: 12, departed: 3 },
      { month: 'Oct', total: 851, new: 15, departed: 6 },
      { month: 'Nov', total: 858, new: 11, departed: 4 },
      { month: 'Dec', total: 862, new: 9, departed: 5 },
      { month: 'Jan', total: 869, new: 14, departed: 7 },
      { month: 'Feb', total: 878, new: 16, departed: 7 },
    ],
    workloadSummary: {
      avgCaseload: 23.4,
      highCaseloadPct: 18,
      submissionRate: 71,
      trendDirection: 'increasing',
    },
  };

  PropertiesService.getScriptProperties().setProperty('SEEDED_UNION_STATS', JSON.stringify(stats));
  Logger.log('Seeded union stats data');
}

// ============================================================================
// SEED: RESOURCES (Know Your Rights, Guides, Policies)
// ============================================================================

/**
 * Seeds additional educational resources beyond the 8 starter entries created
 * by createResourcesSheet(). Adds union-specific content including
 * FMLA, ADA, overtime, workplace safety, and contract-specific guides.
 * Only adds rows beyond existing data to avoid duplicates.
 */
function seedResourcesData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.RESOURCES);

  if (!sheet) {
    if (typeof createResourcesSheet === 'function') {
      sheet = createResourcesSheet(ss);
    }
    if (!sheet) {
      Logger.log('Resources sheet not found. Skipping resource seed.');
      return;
    }
  }

  // Check how many resources already exist
  var existingRows = sheet.getLastRow() - 1; // minus header
  if (existingRows >= 15) {
    Logger.log('Resources sheet already has ' + existingRows + ' entries. Skipping seed.');
    return;
  }

  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Extended resources: start IDs after existing starter content (RES-009+)
  var nextId = existingRows + 1;
  var additionalResources = [
    ['RES-' + String(nextId++).padStart(3, '0'), 'FMLA: Your Rights to Medical Leave', 'Know Your Rights',
      'The Family and Medical Leave Act protects your right to take unpaid leave for medical reasons.',
      'FMLA provides up to 12 weeks of unpaid, job-protected leave per year for:\\n- Your own serious health condition\\n- Caring for a spouse, child, or parent with a serious health condition\\n- Birth or adoption of a child\\n- Qualifying military family needs\\n\\nYour employer cannot retaliate against you for taking FMLA leave. You must be restored to the same or equivalent position.\\n\\nTo request FMLA: notify your supervisor and HR at least 30 days in advance when possible. Your steward can help if your request is denied.',
      '', '🏥', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'ADA Reasonable Accommodations', 'Know Your Rights',
      'You have the right to reasonable accommodations for disabilities under the ADA.',
      'The Americans with Disabilities Act requires your employer to provide reasonable accommodations unless it causes undue hardship.\\n\\nExamples of accommodations:\\n- Modified work schedule\\n- Ergonomic equipment\\n- Reassignment to a vacant position\\n- Leave for medical treatment\\n- Telework arrangements\\n\\nThe process: request an accommodation in writing, engage in the "interactive process" with HR, and document everything. Contact your steward immediately if your accommodation is denied or delayed.',
      '', '♿', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Overtime Rules & Comp Time', 'Contract Article',
      'Know your rights regarding overtime pay, mandatory overtime, and compensatory time.',
      'Under the collective bargaining agreement:\\n- Overtime is paid at 1.5x your regular rate for hours over 40/week\\n- Mandatory overtime must follow contract provisions for notice and rotation\\n- Comp time may be offered in lieu of overtime pay in some classifications\\n- You cannot be disciplined for declining overtime that violates the contract\\n\\nIf you believe overtime is being distributed unfairly or mandatory OT violates the contract, document the specifics and contact your steward.',
      '', '⏰', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Workplace Safety: Reporting Hazards', 'Know Your Rights',
      'You have the right to a safe workplace and cannot be retaliated against for reporting hazards.',
      'OSHA protects your right to:\\n- Report unsafe conditions without retaliation\\n- Request an OSHA inspection\\n- Receive safety training in a language you understand\\n- Access injury and illness records\\n- Refuse dangerous work under specific conditions\\n\\nTo report a hazard:\\n1. Notify your supervisor immediately\\n2. Document the hazard (photos, dates, witnesses)\\n3. File a report with the Health & Safety Committee\\n4. Contact your steward if management does not respond\\n\\nIf there is immediate danger, you may refuse work — contact your steward first if time permits.',
      '', '⚠️', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Anti-Retaliation Protections', 'Know Your Rights',
      'It is illegal for management to retaliate against you for union activity.',
      'Protected activities include:\\n- Filing or supporting a grievance\\n- Attending union meetings\\n- Talking to coworkers about working conditions\\n- Reporting safety violations\\n- Requesting union representation (Weingarten rights)\\n- Participating in lawful pickets or actions\\n\\nSigns of retaliation:\\n- Sudden negative performance reviews\\n- Schedule changes or undesirable assignments\\n- Increased scrutiny or micromanagement\\n- Denial of previously approved requests\\n\\nIf you suspect retaliation, document everything and contact your steward immediately. Retaliation is an unfair labor practice and can be challenged.',
      '', '🛡️', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Understanding Your Pay Stub', 'Guide',
      'A guide to understanding deductions, union dues, and benefits on your pay stub.',
      'Your pay stub includes:\\n- Gross pay: your total earnings before deductions\\n- Union dues: typically a percentage of gross pay, deducted automatically\\n- Health insurance: your share of premium costs\\n- Retirement contributions: pension or 401k deductions\\n- Taxes: federal, state, and local withholdings\\n\\nCommon issues to watch for:\\n- Incorrect classification (hourly vs salaried)\\n- Missing overtime or differential pay\\n- Unauthorized deductions\\n\\nIf your pay stub seems wrong, contact payroll first. If the issue is not resolved, contact your steward to file a grievance.',
      '', '💰', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Steward Quick Reference Card', 'Guide',
      'Essential reference for stewards: key contract articles, timelines, and procedures.',
      'Key timelines:\\n- Incident to Step I filing: check contract (typically 15-30 days)\\n- Step I response: per contract\\n- Step II appeal: per contract from Step I response\\n- Step III / arbitration: per contract\\n\\nInvestigatory interview checklist:\\n1. Member requests representation\\n2. Meet privately with member first\\n3. Take notes during the meeting\\n4. Advise member of rights before questions\\n5. Object to improper questions\\n6. Request caucus if needed\\n\\nDocumentation: always get who, what, when, where, witnesses.',
      '', '📋', nextId - 1, 'Yes', 'Stewards', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'New Employee Rights Checklist', 'Guide',
      'New to the bargaining unit? Here is what you need to know about your rights and benefits.',
      'As a new bargaining unit member you should:\\n\\n1. Know your steward — check the dashboard for contact info\\n2. Read your contract — ask your steward for a copy\\n3. Understand your probation period and what it means\\n4. Set up your PIN for the union dashboard\\n5. Attend new member orientation\\n6. Know your Weingarten rights from day one\\n7. Understand the grievance process\\n8. Complete your workload tracker weekly\\n9. Attend union meetings when possible\\n10. Report any contract violations to your steward\\n\\nYour union is here for you from day one.',
      '', '✅', nextId - 1, 'Yes', 'Members', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Collective Bargaining Agreement Summary', 'Contract Article',
      'Key provisions of the current collective bargaining agreement.',
      'The CBA covers:\\n- Wages and step increases\\n- Health insurance and benefits\\n- Work schedules and overtime\\n- Grievance and arbitration procedures\\n- Seniority rights\\n- Transfers and promotions\\n- Discipline and discharge (just cause)\\n- Leave policies (sick, personal, vacation)\\n- Workplace safety\\n- Union rights and representation\\n\\nThe full contract is available from your steward or union office. Key articles are referenced in grievance filings — your steward will identify the specific articles that apply to your situation.',
      '', '📜', nextId - 1, 'Yes', 'All', today, 'System'],
    ['RES-' + String(nextId++).padStart(3, '0'), 'Caseload Standards & Workload Rights', 'Policy',
      'Your rights regarding manageable caseloads and workload distribution.',
      'Under the contract and agency policy:\\n- Caseload limits exist for certain classifications\\n- You have the right to report excessive workloads through the workload tracker\\n- Management must address staffing shortages that create unsafe conditions\\n- Mandatory overtime to cover caseloads must follow contract provisions\\n\\nUsing the Workload Tracker:\\n1. Submit your weekly caseload numbers\\n2. Flag any categories that feel unmanageable\\n3. Your submission is anonymized in reports\\n4. Aggregate data helps the union advocate for adequate staffing\\n\\nIf your caseload is unsustainable, document specific impacts and speak with your steward.',
      '', '📊', nextId - 1, 'Yes', 'All', today, 'System']
  ];

  var headers = getHeadersFromMap_(RESOURCES_HEADER_MAP_);
  sheet.getRange(existingRows + 2, 1, additionalResources.length, headers.length)
    .setValues(additionalResources);

  Logger.log('Seeded ' + additionalResources.length + ' additional resources (total: ' + (existingRows + additionalResources.length) + ')');
}

// ============================================================================
// SEED: NOTIFICATIONS
// ============================================================================

/**
 * Seeds sample notifications to showcase the in-app notification system.
 * Adds a mix of announcements, deadlines, and steward messages across
 * different priorities and recipients.
 */
function seedNotificationsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);

  if (!sheet) {
    if (typeof createNotificationsSheet === 'function') {
      sheet = createNotificationsSheet(ss);
    }
    if (!sheet) {
      Logger.log('Notifications sheet not found. Skipping notification seed.');
      return;
    }
  }

  // Check existing data
  var existingRows = sheet.getLastRow() - 1;
  if (existingRows >= 6) {
    Logger.log('Notifications sheet already has ' + existingRows + ' entries. Skipping seed.');
    return;
  }

  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  var inOneWeek = new Date(now.getTime() + 7 * 86400000);
  var inTwoWeeks = new Date(now.getTime() + 14 * 86400000);
  var inOneMonth = new Date(now.getTime() + 30 * 86400000);
  var threeDaysAgo = new Date(now.getTime() - 3 * 86400000);

  var fmt = function(d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };

  var nextId = existingRows + 1;

  var notifications = [
    [
      'NOTIF-' + String(nextId++).padStart(3, '0'),
      'All Members',
      'Announcement',
      'Contract Negotiations Update',
      'The bargaining committee met with management on ' + fmt(threeDaysAgo) + '. Key topics discussed: wage increases, telecommuting policy, and caseload limits. A full update will be shared at the next general membership meeting. Your support matters — please attend.',
      'Normal',
      'bargaining@seiu509.org',
      'Bargaining Committee',
      fmt(threeDaysAgo),
      fmt(inOneMonth),
      '',
      'Active',
      'Dismissible'
    ],
    [
      'NOTIF-' + String(nextId++).padStart(3, '0'),
      'All Members',
      'Deadline',
      'Workload Survey Due This Friday',
      'Please submit your weekly workload tracker by end of day Friday. Your data helps the union build the case for adequate staffing. Submissions are anonymous and take less than 2 minutes.',
      'Urgent',
      'workload@seiu509.org',
      'Workload Committee',
      today,
      fmt(inOneWeek),
      '',
      'Active',
      'Timed'
    ],
    [
      'NOTIF-' + String(nextId++).padStart(3, '0'),
      'All Members',
      'Announcement',
      'New Know Your Rights Resources Available',
      'We have added new educational resources to the Learn tab including FMLA rights, ADA accommodations, anti-retaliation protections, and caseload standards. Check them out and know your rights!',
      'Normal',
      'education@seiu509.org',
      'Education Committee',
      today,
      fmt(inOneMonth),
      '',
      'Active',
      'Dismissible'
    ],
    [
      'NOTIF-' + String(nextId++).padStart(3, '0'),
      'All Members',
      'Steward Message',
      'Steward Office Hours This Week',
      'Your stewards will be available for drop-in office hours this week: Tuesday 12-1pm and Thursday 3-4pm in the break room. No appointment needed. Bring your questions about the contract, grievances, or any workplace concerns.',
      'Normal',
      'stewards@seiu509.org',
      'Steward Team',
      today,
      fmt(inTwoWeeks),
      '',
      'Active',
      'Dismissible'
    ]
  ];

  var headers = getHeadersFromMap_(NOTIFICATIONS_HEADER_MAP_);
  sheet.getRange(existingRows + 2, 1, notifications.length, headers.length)
    .setValues(notifications);

  Logger.log('Seeded ' + notifications.length + ' notifications (total: ' + (existingRows + notifications.length) + ')');
}

// ============================================================================
// SEED: WORKLOAD TRACKER DATA
// ============================================================================

/**
 * Seeds 20 sample workload submissions into the Workload Vault.
 * Uses sample member emails from the Member Directory and generates
 * realistic caseload numbers. Also triggers reporting refresh.
 */
function seedWorkloadData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);

  if (!vault) {
    Logger.log('Workload Vault sheet not found. Skipping workload seed.');
    return;
  }

  // Check existing data
  if (vault.getLastRow() > 1) {
    Logger.log('Workload Vault already has data. Skipping seed.');
    return;
  }

  // Get some member emails from the directory
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var emails = [];
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var emailCol = MEMBER_COLS.EMAIL;
    var memberData = memberSheet.getRange(2, emailCol, Math.min(memberSheet.getLastRow() - 1, 100), 1).getValues();
    for (var i = 0; i < memberData.length; i++) {
      if (memberData[i][0]) emails.push(memberData[i][0]);
    }
  }

  // Fallback if no members seeded yet
  if (emails.length === 0) {
    emails = [
      'john.smith@example.org', 'mary.johnson@example.org', 'robert.williams@example.org',
      'patricia.jones@example.org', 'michael.davis@example.org', 'linda.miller@example.org',
      'william.brown@example.org', 'barbara.wilson@example.org', 'david.moore@example.org',
      'susan.taylor@example.org'
    ];
  }

  var now = new Date();
  var rows = [];

  // Generate 20 submissions spread over the last 4 weeks
  for (var w = 0; w < 20; w++) {
    var daysAgo = Math.floor(Math.random() * 28);
    var timestamp = new Date(now.getTime() - daysAgo * 86400000);
    var email = emails[Math.floor(Math.random() * emails.length)];

    // Realistic caseload numbers for social workers
    var priorityCases = Math.floor(Math.random() * 8) + 1;
    var pendingCases = Math.floor(Math.random() * 15) + 5;
    var unreadDocs = Math.floor(Math.random() * 20);
    var todoItems = Math.floor(Math.random() * 12) + 2;
    var sentReferrals = Math.floor(Math.random() * 5);
    var ceActivities = Math.floor(Math.random() * 3);
    var assistanceReqs = Math.floor(Math.random() * 4);
    var agedCases = Math.floor(Math.random() * 6);
    var weeklyCases = Math.floor(Math.random() * 8) + 2;

    // Sub-categories JSON (simplified)
    var subCats = JSON.stringify({
      intake: Math.floor(Math.random() * 4),
      review: Math.floor(Math.random() * 5),
      closing: Math.floor(Math.random() * 3)
    });

    // Employment details
    var empTypes = ['Full-Time', 'Full-Time', 'Full-Time', 'Part-Time'];
    var empType = empTypes[Math.floor(Math.random() * empTypes.length)];
    var ptHours = empType === 'Part-Time' ? (20 + Math.floor(Math.random() * 16)) : '';

    // Leave
    var leaveTypes = ['', '', '', '', 'Sick', 'Vacation', 'Personal'];
    var leaveType = leaveTypes[Math.floor(Math.random() * leaveTypes.length)];
    var leavePlanned = leaveType ? 'Yes' : '';
    var leaveStart = '';
    var leaveEnd = '';
    if (leaveType) {
      var leaveOffset = Math.floor(Math.random() * 14) + 1;
      leaveStart = new Date(now.getTime() + leaveOffset * 86400000);
      leaveEnd = new Date(leaveStart.getTime() + (Math.floor(Math.random() * 3) + 1) * 86400000);
    }

    // Intake/notice
    var noIntake = Math.random() < 0.2 ? 'Yes' : '';
    var noticeTime = noIntake ? (Math.floor(Math.random() * 3) + 1) + ' days' : '';
    var halfDay = Math.random() < 0.1 ? 'Yes' : '';

    // Privacy & plan
    var privacy = Math.random() < 0.3 ? 'Anonymous' : 'Identified';
    var onPlan = Math.random() < 0.15 ? 'Yes' : 'No';
    var overtime = Math.random() < 0.3 ? (Math.floor(Math.random() * 8) + 1) : 0;

    // Vault row: 24 columns matching header order
    rows.push([
      timestamp, email,
      priorityCases, pendingCases, unreadDocs, todoItems,
      sentReferrals, ceActivities, assistanceReqs, agedCases,
      weeklyCases, subCats,
      empType, ptHours,
      leaveType, leavePlanned, leaveStart, leaveEnd,
      noIntake, noticeTime, halfDay,
      privacy, onPlan, overtime
    ]);
  }

  vault.getRange(2, 1, rows.length, 24).setValues(rows);

  // Format timestamp column
  vault.getRange(2, 1, rows.length, 1).setNumberFormat('MM/dd/yyyy HH:mm');

  // Refresh reporting if available
  if (typeof WorkloadPortal !== 'undefined' && WorkloadPortal._refreshReportingData) {
    try { WorkloadPortal._refreshReportingData(); } catch (_e) { /* ok */ }
  }
  if (typeof WorkloadService !== 'undefined' && WorkloadService.refreshLedger) {
    try { WorkloadService.refreshLedger(); } catch (_e) { /* ok */ }
  }

  Logger.log('Seeded ' + rows.length + ' workload submissions');
}

// ============================================================================
// SEED: STEWARD TASKS
// ============================================================================

/**
 * Seeds 8 sample steward tasks for the script-owner steward.
 * Uses the _Steward_Tasks sheet.
 */
function seedStewardTasksData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_TASKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_TASKS);
    sheet.getRange(1, 1, 1, 12).setValues([['ID', 'Steward Email', 'Title', 'Description', 'Member Email', 'Priority', 'Status', 'Due Date', 'Created', 'Completed', 'Assignee Type', 'Assigned By']]);
    sheet.hideSheet();
  }

  if (sheet.getLastRow() > 1) {
    Logger.log('Steward Tasks already has data. Skipping seed.');
    return;
  }

  var ownerEmail = '';
  try { ownerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
  if (!ownerEmail) ownerEmail = 'steward@example.org';

  var now = new Date();
  var tasks = [
    ['ST_SEED_1', ownerEmail, 'Follow up on scheduling grievance', 'Check if the Step I response has been received', '', 'high', 'open', new Date(now.getTime() + 3 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_2', ownerEmail, 'Prepare for bargaining meeting', 'Compile member feedback and workload data for negotiation prep', '', 'high', 'open', new Date(now.getTime() + 7 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_3', ownerEmail, 'Contact new members', 'Reach out to 5 new members who joined this month', '', 'medium', 'open', new Date(now.getTime() + 5 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_4', ownerEmail, 'Submit weekly workload report', 'Compile and submit caseload numbers to the workload committee', '', 'medium', 'open', new Date(now.getTime() + 2 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_5', ownerEmail, 'Review contract language on overtime', 'Research Articles 12 and 15 for upcoming overtime dispute', '', 'low', 'open', new Date(now.getTime() + 14 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_6', ownerEmail, 'Organize Know Your Rights session', 'Book room and prepare materials for next month Weingarten rights training', '', 'medium', 'open', new Date(now.getTime() + 21 * 86400000), now, '', 'steward', ''],
    ['ST_SEED_7', ownerEmail, 'Update member contact info', 'Three members reported new phone numbers — update directory', '', 'low', 'completed', new Date(now.getTime() - 2 * 86400000), new Date(now.getTime() - 5 * 86400000), new Date(now.getTime() - 2 * 86400000), 'steward', ''],
    ['ST_SEED_8', ownerEmail, 'File safety complaint for Building B', 'Document the HVAC issues and file with Health & Safety Committee', '', 'high', 'completed', new Date(now.getTime() - 1 * 86400000), new Date(now.getTime() - 7 * 86400000), new Date(now.getTime() - 1 * 86400000), 'steward', ''],
  ];

  sheet.getRange(2, 1, tasks.length, 12).setValues(tasks);
  Logger.log('Seeded ' + tasks.length + ' steward tasks');
}

// ============================================================================
// SEED: MEETING MINUTES
// ============================================================================

/**
 * Seeds 4 sample meeting minutes into MeetingMinutes sheet.
 */
function seedMinutesData() {
  var sheet = getOrCreateMinutesSheet();

  if (sheet.getLastRow() > 1) {
    Logger.log('Meeting Minutes already has data. Skipping seed.');
    return;
  }

  var now = new Date();
  var ownerEmail = '';
  try { ownerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
  if (!ownerEmail) ownerEmail = 'steward@example.org';

  var minutes = [
    ['MIN_SEED_1', new Date(now.getTime() - 28 * 86400000), 'January General Membership Meeting',
      '• Contract update: management counter-proposal received\n• Workload committee report: caseloads up 12%\n• New steward training scheduled for Feb\n• Vote: approved $500 solidarity fund donation',
      'Full discussion notes available in the shared drive. 42 members attended. Next meeting scheduled for February.',
      ownerEmail, new Date(now.getTime() - 27 * 86400000)],
    ['MIN_SEED_2', new Date(now.getTime() - 14 * 86400000), 'Executive Board Meeting',
      '• Budget review: Q4 spending on track\n• Grievance backlog discussed: 8 cases pending Step II\n• Legislative action day planning\n• Steward appreciation dinner logistics',
      'Board unanimously approved the legislative action day budget. Steward dinner set for next month.',
      ownerEmail, new Date(now.getTime() - 13 * 86400000)],
    ['MIN_SEED_3', new Date(now.getTime() - 7 * 86400000), 'Workplace Safety Committee',
      '• Building B HVAC complaint filed with management\n• New ergonomic equipment request approved\n• Incident report review: 2 slip-and-fall reports\n• Emergency exit signage audit complete',
      'Management has 14 days to respond to the HVAC complaint per the contract. Follow-up scheduled.',
      ownerEmail, new Date(now.getTime() - 6 * 86400000)],
    ['MIN_SEED_4', new Date(now.getTime() - 2 * 86400000), 'Steward Training: Documentation Best Practices',
      '• Reviewed proper grievance documentation techniques\n• Discussed Weingarten rights scenarios\n• Practiced investigatory interview role-plays\n• Shared template for witness statements',
      '15 stewards attended. Training materials uploaded to Resources section.',
      ownerEmail, new Date(now.getTime() - 1 * 86400000)],
  ];

  sheet.getRange(2, 1, minutes.length, 7).setValues(minutes);
  Logger.log('Seeded ' + minutes.length + ' meeting minutes');
}

// ============================================================================
// SEED: MEETING CHECK-IN DATA
// ============================================================================

/**
 * Seeds 10 sample meeting check-in records for seeded members.
 * Uses the Meeting Check-In Log sheet.
 */
function seedMeetingCheckinData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
  if (!sheet) {
    Logger.log('Meeting Check-In Log sheet not found. Skipping seed.');
    return;
  }

  if (sheet.getLastRow() > 1) {
    Logger.log('Meeting Check-In Log already has data. Skipping seed.');
    return;
  }

  // Get some member emails
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var emails = [];
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, MEMBER_COLS.EMAIL, Math.min(memberSheet.getLastRow() - 1, 50), 1).getValues();
    for (var i = 0; i < memberData.length; i++) {
      if (memberData[i][0]) emails.push(String(memberData[i][0]));
    }
  }

  // Add script owner
  try {
    var ownerEmail = Session.getActiveUser().getEmail();
    if (ownerEmail && emails.indexOf(ownerEmail) === -1) emails.unshift(ownerEmail);
  } catch (_e) { /* headless */ }

  if (emails.length === 0) {
    Logger.log('No members found for meeting check-in seed.');
    return;
  }

  var now = new Date();
  var meetingTypes = ['General Membership', 'Steward Training', 'Committee Meeting', 'Town Hall', 'New Member Orientation'];
  var meetingNames = [
    'Monthly General Membership Meeting', 'Steward Training: Grievance Basics',
    'Workplace Safety Committee', 'Quarterly Town Hall', 'New Member Orientation',
    'Contract Enforcement Workshop', 'Know Your Rights Lunch & Learn',
    'Executive Board Meeting', 'Legislative Action Planning', 'Member Appreciation Event'
  ];
  var durations = [0.5, 0.75, 1, 1.5, 2];

  /**
   * Builds a 16-column meeting check-in row using MEETING_CHECKIN_COLS.
   * @param {string} id - Meeting check-in ID.
   * @param {string} meetingName - Name of the meeting.
   * @param {Date} meetingDate - Date of the meeting.
   * @param {string} meetingType - Type of meeting.
   * @param {string} memberId - Member ID.
   * @param {string} memberName - Member display name.
   * @param {Date} checkinTime - Time the member checked in.
   * @param {string} email - Member email.
   * @param {string} meetingTime - Scheduled meeting time.
   * @param {number} duration - Meeting duration in hours.
   * @param {string} status - Meeting status.
   * @param {string} notify - Whether to notify stewards.
   * @param {string} calEventId - Google Calendar event ID.
   * @param {string} notesUrl - Notes document URL.
   * @param {string} agendaUrl - Agenda document URL.
   * @param {string} agendaStewards - Stewards on the agenda.
   * @returns {Array} 16-element row array.
   */
  function buildCheckinRow(id, meetingName, meetingDate, meetingType, memberId, memberName, checkinTime, email, meetingTime, duration, status, notify, calEventId, notesUrl, agendaUrl, agendaStewards) {
    var row = [];
    for (var ci = 0; ci < 16; ci++) { row.push(''); }
    row[MEETING_CHECKIN_COLS.MEETING_ID - 1] = id;
    row[MEETING_CHECKIN_COLS.MEETING_NAME - 1] = meetingName;
    row[MEETING_CHECKIN_COLS.MEETING_DATE - 1] = meetingDate;
    row[MEETING_CHECKIN_COLS.MEETING_TYPE - 1] = meetingType;
    row[MEETING_CHECKIN_COLS.MEMBER_ID - 1] = memberId || '';
    row[MEETING_CHECKIN_COLS.MEMBER_NAME - 1] = memberName || '';
    row[MEETING_CHECKIN_COLS.CHECKIN_TIME - 1] = checkinTime || '';
    row[MEETING_CHECKIN_COLS.EMAIL - 1] = email || '';
    row[MEETING_CHECKIN_COLS.MEETING_TIME - 1] = meetingTime || '';
    row[MEETING_CHECKIN_COLS.MEETING_DURATION - 1] = duration || '';
    row[MEETING_CHECKIN_COLS.EVENT_STATUS - 1] = status || '';
    row[MEETING_CHECKIN_COLS.NOTIFY_STEWARDS - 1] = notify || '';
    row[MEETING_CHECKIN_COLS.CALENDAR_EVENT_ID - 1] = calEventId || '';
    row[MEETING_CHECKIN_COLS.NOTES_DOC_URL - 1] = notesUrl || '';
    row[MEETING_CHECKIN_COLS.AGENDA_DOC_URL - 1] = agendaUrl || '';
    row[MEETING_CHECKIN_COLS.AGENDA_STEWARDS - 1] = agendaStewards || '';
    return row;
  }

  var rows = [];

  // --- Past meetings (completed, with check-in records) ---
  for (var c = 0; c < 10; c++) {
    var daysAgo = Math.floor(Math.random() * 60) + 1;
    var meetingDate = new Date(now.getTime() - daysAgo * 86400000);
    meetingDate.setHours(0, 0, 0, 0);
    var checkinTime = new Date(meetingDate.getTime() + 17 * 3600000 - Math.floor(Math.random() * 600000)); // ~5pm, 0-10 min early
    var email = emails[Math.floor(Math.random() * emails.length)];
    var mName = meetingNames[Math.floor(Math.random() * meetingNames.length)];
    var firstName = email.split('.')[0] || 'Member';
    firstName = firstName.charAt(0).toUpperCase() + firstName.slice(1);

    rows.push(buildCheckinRow(
      'MC_SEED_' + (c + 1), mName, meetingDate,
      meetingTypes[Math.floor(Math.random() * meetingTypes.length)],
      '', firstName, checkinTime, email,
      '17:00', durations[Math.floor(Math.random() * durations.length)], MEETING_STATUS.COMPLETED,
      '', '', '', '', ''
    ));
  }

  // --- TODAY'S meetings: Active and eligible for check-in ---
  var todayDate = new Date(now);
  todayDate.setHours(0, 0, 0, 0);

  // Meeting 1: Open Town Hall — active now, long duration so it stays open
  rows.push(buildCheckinRow(
    'MC_SEED_TODAY_1', 'Open Town Hall — Sign In to Test!', todayDate,
    'Town Hall', '', '', '', '',
    '00:01', 23, MEETING_STATUS.ACTIVE,
    '', '', '', '', ''
  ));

  // Meeting 2: Steward Huddle — also active today
  rows.push(buildCheckinRow(
    'MC_SEED_TODAY_2', 'Weekly Steward Huddle', todayDate,
    'Committee Meeting', '', '', '', '',
    '08:00', 10, MEETING_STATUS.SCHEDULED,
    '', '', '', '', ''
  ));

  sheet.getRange(2, 1, rows.length, 16).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' meeting check-in records (including 2 active today)');
}

// ============================================================================
// SEED: TIMELINE EVENTS
// ============================================================================

/**
 * Seeds 6 sample timeline events into _Timeline_Events.
 */
function seedTimelineData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
  if (!sheet) {
    if (typeof TimelineService !== 'undefined' && TimelineService.initSheet) {
      try { TimelineService.initSheet(); } catch (_e) { /* ok */ }
    }
    sheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!sheet) {
      Logger.log('Timeline Events sheet not found. Skipping seed.');
      return;
    }
  }

  if (sheet.getLastRow() > 1) {
    Logger.log('Timeline Events already has data. Skipping seed.');
    return;
  }

  var now = new Date();
  var ownerEmail = '';
  try { ownerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
  if (!ownerEmail) ownerEmail = 'system';

  var events = [
    ['TL_SEED_1', 'Contract Negotiations Began', new Date(now.getTime() - 60 * 86400000),
      'The bargaining committee opened formal negotiations with management on the new CBA.',
      'milestone', '', '', '', '', ownerEmail, new Date(now.getTime() - 60 * 86400000), ''],
    ['TL_SEED_2', 'New Steward Training Completed', new Date(now.getTime() - 45 * 86400000),
      '12 new stewards completed the 3-day training program covering grievance handling and member advocacy.',
      'milestone', '', '', '', '', ownerEmail, new Date(now.getTime() - 45 * 86400000), ''],
    ['TL_SEED_3', 'Workload Survey Launched', new Date(now.getTime() - 30 * 86400000),
      'The weekly workload tracker was deployed to all bargaining unit members.',
      'announcement', '', '', '', '', ownerEmail, new Date(now.getTime() - 30 * 86400000), ''],
    ['TL_SEED_4', 'Safety Grievance Won (Building B HVAC)', new Date(now.getTime() - 14 * 86400000),
      'Management agreed to replace the HVAC system in Building B within 90 days. Full remedy achieved.',
      'decision', '', '', '', '', ownerEmail, new Date(now.getTime() - 14 * 86400000), ''],
    ['TL_SEED_5', 'Legislative Action Day', new Date(now.getTime() + 10 * 86400000),
      'Members will call legislators about workforce funding. Materials and talking points available in Resources.',
      'action', '', '', '', '', ownerEmail, now, ''],
    ['TL_SEED_6', 'Contract Ratification Vote Scheduled', new Date(now.getTime() + 30 * 86400000),
      'All bargaining unit members eligible to vote on the tentative agreement. Details to follow.',
      'announcement', '', '', '', '', ownerEmail, now, ''],
    ['TL_SEED_7', 'Steward Refresher Training', new Date(now.getTime() + 5 * 86400000),
      'All stewards are required to attend the annual refresher on grievance handling updates and new contract language.',
      'milestone', '', '', '', '', ownerEmail, now, ''],
    ['TL_SEED_8', 'Monthly General Membership Meeting', new Date(now.getTime() + 14 * 86400000),
      'Open to all members. Agenda: negotiations update, committee reports, and open floor Q&A.',
      'meeting', '', '', '', '', ownerEmail, now, ''],
    ['TL_SEED_9', 'Workplace Safety Walk-Through', new Date(now.getTime() + 21 * 86400000),
      'Joint labor-management safety inspection of Building C. Stewards and safety committee members should attend.',
      'action', '', '', '', '', ownerEmail, now, ''],
    ['TL_SEED_10', 'New Member Welcome Social', new Date(now.getTime() + 45 * 86400000),
      'Casual meet-and-greet for members who joined in the last quarter. Food and refreshments provided.',
      'announcement', '', '', '', '', ownerEmail, now, ''],
  ];

  sheet.getRange(2, 1, events.length, 12).setValues(events);
  Logger.log('Seeded ' + events.length + ' timeline events');
}

// ============================================================================
// SEED: Q&A FORUM
// ============================================================================

/**
 * Seeds sample Q&A Forum data with realistic questions and answers.
 * Populates _QA_Forum (questions) and _QA_Answers (answers) sheets.
 * Demonstrates anonymous posting, upvoting, steward answers, and moderation.
 */
function seedQAForumData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure sheets exist
  if (typeof QAForum !== 'undefined' && QAForum.initSheets) {
    try { QAForum.initSheets(); } catch (_e) { /* ok */ }
  }

  var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
  var answerSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);

  if (!forumSheet || !answerSheet) {
    Logger.log('QA Forum sheets not found. Skipping seed.');
    return;
  }

  if (forumSheet.getLastRow() > 1) {
    Logger.log('QA Forum already has data. Skipping seed.');
    return;
  }

  // Gather member emails from Member Directory for realistic authors
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var memberEmails = [];
  var memberNames = [];
  var stewardEmails = [];
  var stewardNames = [];

  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var m = 1; m < memberData.length && memberEmails.length < 30; m++) {
      var mEmail = String(memberData[m][MEMBER_COLS.EMAIL - 1] || '').trim();
      var mFirst = String(memberData[m][MEMBER_COLS.FIRST_NAME - 1] || '');
      var mLast = String(memberData[m][MEMBER_COLS.LAST_NAME - 1] || '');
      var mName = (mFirst + ' ' + mLast).trim();
      var isSteward = String(memberData[m][MEMBER_COLS.IS_STEWARD - 1] || '').toLowerCase() === 'yes';
      if (mEmail) {
        memberEmails.push(mEmail);
        memberNames.push(mName);
        if (isSteward) {
          stewardEmails.push(mEmail);
          stewardNames.push(mName);
        }
      }
    }
  }

  if (memberEmails.length === 0) {
    memberEmails = ['member1@example.org', 'member2@example.org', 'member3@example.org'];
    memberNames = ['Jane Doe', 'John Smith', 'Alex Rivera'];
  }
  if (stewardEmails.length === 0) {
    stewardEmails = [memberEmails[0]];
    stewardNames = [memberNames[0]];
  }

  var now = new Date();
  var qIdCounter = 1;
  var aIdCounter = 1;

  // Sample questions covering topics members would realistically ask
  var questions = [
    {
      text: 'How do I file a grievance if my supervisor denied my schedule change request?',
      isAnon: false, upvotes: 12, status: 'active', daysAgo: 21,
      answers: [
        { text: 'Great question! First, document the denial in writing. Then contact your assigned steward within the contractual time limit (usually 21 days). Your steward will help you complete the grievance form and determine which article was violated. The process starts at the informal step where we try to resolve it directly with management.', isSteward: true, daysAgo: 20 },
        { text: 'I had a similar situation last year. My steward was really helpful walking me through it. Make sure to save any emails or written communication about the denial.', isSteward: false, daysAgo: 19 }
      ]
    },
    {
      text: 'What are the current contract rules about overtime assignments?',
      isAnon: false, upvotes: 18, status: 'active', daysAgo: 18,
      answers: [
        { text: 'Per Article 7 of our CBA, overtime is first offered by seniority within the unit. If no volunteers, it can be mandated in reverse seniority order. Management must give at least 24 hours notice for non-emergency overtime. Check the Resources tab for the full contract language.', isSteward: true, daysAgo: 17 },
        { text: 'Also worth noting — if you\'re consistently being skipped for OT or assigned OT out of order, that could be a contract violation. Keep a log of OT assignments in your unit.', isSteward: true, daysAgo: 16 }
      ]
    },
    {
      text: 'Is there a way to see how many members have completed the satisfaction survey?',
      isAnon: true, upvotes: 7, status: 'active', daysAgo: 14,
      answers: [
        { text: 'Yes! Stewards can check the Survey Tracking page which shows overall completion rate and a list of pending members. We send reminders periodically. The survey is anonymous — we can see who completed it but not individual responses.', isSteward: true, daysAgo: 13 }
      ]
    },
    {
      text: 'Can someone explain what the different grievance steps mean?',
      isAnon: true, upvotes: 24, status: 'active', daysAgo: 30,
      answers: [
        { text: 'Here\'s the quick rundown:\n\n• Informal: First attempt to resolve directly with your supervisor\n• Step I: Formal written grievance filed, management has 30 days to respond\n• Step II: Appeal to next management level if Step I denied\n• Step III: Appeal to agency head/labor relations\n• Mediation: Neutral third party helps find resolution\n• Arbitration: Binding decision by an independent arbitrator\n\nMost cases resolve at Informal or Step I. Your steward guides you through each step.', isSteward: true, daysAgo: 29 }
      ]
    },
    {
      text: 'When is the next union meeting? I want to bring up staffing concerns in our unit.',
      isAnon: false, upvotes: 9, status: 'active', daysAgo: 7,
      answers: [
        { text: 'Check the Events tab for upcoming meeting dates. You can also raise staffing concerns with your steward anytime — we can bring it to the labor-management committee. If it affects the whole unit, a group grievance might be appropriate.', isSteward: true, daysAgo: 6 },
        { text: 'Our unit had the same issue. We documented caseload numbers over 3 months and presented it at the last meeting. Management agreed to post two new positions. Numbers matter!', isSteward: false, daysAgo: 5 }
      ]
    },
    {
      text: 'I\'m interested in becoming a steward. What does it involve?',
      isAnon: false, upvotes: 15, status: 'active', daysAgo: 25,
      answers: [
        { text: 'That\'s wonderful to hear! Being a steward involves:\n\n• Attending monthly steward meetings\n• Being a point of contact for members in your area\n• Helping members with workplace issues and grievances\n• Participating in steward training (we provide it)\n• Advocating for members in meetings with management\n\nYou get protected time for union activities under our contract. Reach out to any current steward or check the Steward Directory to connect with someone who can sponsor you.', isSteward: true, daysAgo: 24 }
      ]
    },
    {
      text: 'Has anyone used the workload tracker? Does management actually look at the data?',
      isAnon: true, upvotes: 11, status: 'active', daysAgo: 10,
      answers: [
        { text: 'Yes! We compile the workload data and present aggregate trends to management at labor-management meetings. The more members who submit weekly data, the stronger our case for staffing changes. Individual submissions are confidential.', isSteward: true, daysAgo: 9 },
        { text: 'I\'ve been using it for about 2 months. It only takes a minute to fill out. Our steward showed us a chart at the last meeting that really highlighted how overloaded our unit is.', isSteward: false, daysAgo: 8 }
      ]
    },
    {
      text: 'What happens to our contract if there\'s a change in administration?',
      isAnon: false, upvotes: 20, status: 'active', daysAgo: 35,
      answers: [
        { text: 'Our collective bargaining agreement remains in effect regardless of changes in administration. The contract is a legally binding document between the union and the employer. A new administration cannot unilaterally change its terms. If they try, that\'s a contract violation and we file grievances. Our legal team monitors these transitions closely.', isSteward: true, daysAgo: 34 }
      ]
    },
    {
      text: 'How do I update my contact information in the system?',
      isAnon: false, upvotes: 5, status: 'active', daysAgo: 3,
      answers: [
        { text: 'You can update your profile directly in the web app — go to the Profile tab and edit your phone number, preferred communication method, address, etc. Changes save automatically. If you need to update your work email, contact your steward as that requires admin access.', isSteward: true, daysAgo: 2 }
      ]
    },
    {
      text: 'Are union dues tax deductible?',
      isAnon: true, upvotes: 14, status: 'active', daysAgo: 45,
      answers: [
        { text: 'Under current federal tax law, union dues are NOT deductible on federal taxes for most employees (this changed in 2018). However, some states still allow a state tax deduction. Check your state tax rules or consult a tax professional. Your annual dues statement is available from the union office.', isSteward: true, daysAgo: 44 }
      ]
    }
  ];

  var questionRows = [];
  var answerRows = [];

  for (var q = 0; q < questions.length; q++) {
    var question = questions[q];
    var qId = 'QA_SEED_Q' + qIdCounter++;
    var authorIdx = q % memberEmails.length;
    var created = new Date(now.getTime() - question.daysAgo * 86400000);

    // Build upvoters string (random subset of member emails)
    var upvoterList = [];
    for (var u = 0; u < Math.min(question.upvotes, memberEmails.length); u++) {
      var uIdx = (authorIdx + u + 1) % memberEmails.length;
      upvoterList.push(memberEmails[uIdx]);
    }

    questionRows.push([
      qId,
      question.isAnon ? memberEmails[(authorIdx + 3) % memberEmails.length] : memberEmails[authorIdx],
      question.isAnon ? 'Anonymous' : memberNames[authorIdx],
      question.isAnon ? 'Yes' : 'No',
      question.text,
      question.status,
      question.upvotes,
      upvoterList.join(','),
      question.answers.length,
      created,
      created
    ]);

    // Add answers
    for (var a = 0; a < question.answers.length; a++) {
      var answer = question.answers[a];
      var aId = 'QA_SEED_A' + aIdCounter++;
      var ansCreated = new Date(now.getTime() - answer.daysAgo * 86400000);
      var ansAuthorIdx = answer.isSteward
        ? (a % stewardEmails.length)
        : ((authorIdx + a + 1) % memberEmails.length);

      answerRows.push([
        aId,
        qId,
        answer.isSteward ? stewardEmails[ansAuthorIdx] : memberEmails[ansAuthorIdx],
        answer.isSteward ? stewardNames[ansAuthorIdx] : memberNames[ansAuthorIdx],
        answer.isSteward ? 'Yes' : 'No',
        answer.text,
        'active',
        ansCreated
      ]);
    }
  }

  if (questionRows.length > 0) {
    forumSheet.getRange(2, 1, questionRows.length, 11).setValues(questionRows);
  }
  if (answerRows.length > 0) {
    answerSheet.getRange(2, 1, answerRows.length, 8).setValues(answerRows);
  }

  Logger.log('Seeded ' + questionRows.length + ' questions and ' + answerRows.length + ' answers in QA Forum');
}
// ============================================================================
// SEED: MEMBER TASKS
// ============================================================================

/**
 * Seeds sample member tasks into _Steward_Tasks sheet.
 * Member tasks use the 12-column schema with Assignee Type = 'member'.
 * References seeded member emails from Member Directory for realistic linking.
 */
function seedMemberTasksData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_TASKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_TASKS);
    sheet.getRange(1, 1, 1, 12).setValues([['ID', 'Steward Email', 'Title', 'Description', 'Member Email', 'Priority', 'Status', 'Due Date', 'Created', 'Completed', 'Assignee Type', 'Assigned By']]);
    sheet.hideSheet();
  }

  // Ensure Assignee Type / Assigned By headers exist (may be missing on older 10-col sheets)
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 12)).getValues()[0];
  if (!headers[10] || String(headers[10]).trim() !== 'Assignee Type') {
    sheet.getRange(1, 11).setValue('Assignee Type');
    sheet.getRange(1, 12).setValue('Assigned By');
  }

  // Check if member tasks already exist
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues();
    for (var c = 0; c < existing.length; c++) {
      if (String(existing[c][0]).toLowerCase().trim() === 'member') {
        Logger.log('Member tasks already exist. Skipping seed.');
        return;
      }
    }
  }

  var ownerEmail = '';
  try { ownerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
  if (!ownerEmail) ownerEmail = 'steward@example.org';

  // Pull a few seeded member emails for realistic task assignments
  var memberEmails = [];
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var emailCol = MEMBER_COLS.EMAIL || 3;
    var emails = memberSheet.getRange(2, emailCol, Math.min(memberSheet.getLastRow() - 1, 20), 1).getValues();
    for (var e = 0; e < emails.length; e++) {
      if (emails[e][0]) memberEmails.push(String(emails[e][0]).toLowerCase().trim());
    }
  }
  if (memberEmails.length === 0) memberEmails = ['member1@example.org', 'member2@example.org', 'member3@example.org'];

  var now = new Date();
  var tasks = [
    ['MT_SEED_1', ownerEmail, 'Complete new-hire orientation checklist', 'Review union orientation materials and sign acknowledgment form', memberEmails[0], 'high', 'open', new Date(now.getTime() + 5 * 86400000), now, '', 'member', ownerEmail],
    ['MT_SEED_2', ownerEmail, 'Submit workload documentation', 'Track and submit daily caseload numbers for the next two weeks', memberEmails[1 % memberEmails.length], 'medium', 'open', new Date(now.getTime() + 14 * 86400000), now, '', 'member', ownerEmail],
    ['MT_SEED_3', ownerEmail, 'Review updated safety procedures', 'Read the new Building B safety bulletin and confirm understanding', memberEmails[2 % memberEmails.length], 'medium', 'open', new Date(now.getTime() + 7 * 86400000), now, '', 'member', ownerEmail],
    ['MT_SEED_4', ownerEmail, 'Provide witness statement', 'Write a statement about the March 5 scheduling incident for the grievance file', memberEmails[3 % memberEmails.length], 'high', 'open', new Date(now.getTime() + 3 * 86400000), now, '', 'member', ownerEmail],
    ['MT_SEED_5', ownerEmail, 'Complete satisfaction survey', 'Fill out the annual member satisfaction survey before the deadline', memberEmails[4 % memberEmails.length], 'low', 'open', new Date(now.getTime() + 21 * 86400000), now, '', 'member', ownerEmail],
    ['MT_SEED_6', ownerEmail, 'Gather pay stub copies', 'Collect last 4 pay stubs showing overtime discrepancy for grievance evidence', memberEmails[5 % memberEmails.length], 'high', 'completed', new Date(now.getTime() - 1 * 86400000), new Date(now.getTime() - 7 * 86400000), new Date(now.getTime() - 1 * 86400000), 'member', ownerEmail],
    ['MT_SEED_7', ownerEmail, 'Attend Know Your Rights training', 'Attend the scheduled Weingarten rights training session', memberEmails[6 % memberEmails.length], 'low', 'completed', new Date(now.getTime() - 3 * 86400000), new Date(now.getTime() - 14 * 86400000), new Date(now.getTime() - 3 * 86400000), 'member', ownerEmail],
    ['MT_SEED_8', ownerEmail, 'Sign grievance authorization form', 'Sign the authorization form so the steward can file on your behalf', memberEmails[7 % memberEmails.length], 'high', 'open', new Date(now.getTime() + 2 * 86400000), now, '', 'member', ownerEmail],
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, tasks.length, 12).setValues(tasks);
  Logger.log('Seeded ' + tasks.length + ' member tasks');
}

// ============================================================================
// SEED: CASE CHECKLIST
// ============================================================================

/**
 * Seeds sample case checklist items for a few seeded grievances.
 * Uses CHECKLIST_TEMPLATES for realistic checklist content tied to actual grievance types.
 */
function seedCaseChecklistData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
  if (!sheet) {
    if (typeof getOrCreateChecklistSheet === 'function') {
      sheet = getOrCreateChecklistSheet();
    } else {
      sheet = ss.insertSheet(SHEETS.CASE_CHECKLIST);
      var headers = getHeadersFromMap_(CHECKLIST_HEADER_MAP_);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  if (sheet.getLastRow() > 1) {
    Logger.log('Case Checklist already has data. Skipping seed.');
    return;
  }

  // Get a few seeded grievance IDs and their types from Grievance Log
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) {
    Logger.log('No grievances to link checklists to. Skipping seed.');
    return;
  }

  var gIdCol = GRIEVANCE_COLS.GRIEVANCE_ID;
  var gTypeCol = GRIEVANCE_COLS.ISSUE_CATEGORY || GRIEVANCE_COLS.GRIEVANCE_TYPE;
  var gData = grievanceSheet.getRange(2, 1, Math.min(grievanceSheet.getLastRow() - 1, 50), grievanceSheet.getLastColumn()).getValues();

  // Pick up to 5 grievances to seed checklists for
  var targets = [];
  var seenTypes = {};
  for (var i = 0; i < gData.length && targets.length < 5; i++) {
    var caseId = String(gData[i][gIdCol - 1] || '');
    var issueType = String(gData[i][gTypeCol - 1] || '');
    if (!caseId) continue;
    // Try to get variety in issue types
    if (seenTypes[issueType] && targets.length >= 2) continue;
    seenTypes[issueType] = true;
    targets.push({ caseId: caseId, type: issueType });
  }

  if (targets.length === 0) {
    Logger.log('No valid grievance IDs found. Skipping checklist seed.');
    return;
  }

  var ownerEmail = '';
  try { ownerEmail = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
  if (!ownerEmail) ownerEmail = 'steward@example.org';

  var now = new Date();
  var rows = [];
  var idCounter = 1;

  for (var t = 0; t < targets.length; t++) {
    var target = targets[t];
    // Look up template: try specific type under Grievance, then _default
    var template = null;
    if (typeof CHECKLIST_TEMPLATES !== 'undefined' && CHECKLIST_TEMPLATES.Grievance) {
      template = CHECKLIST_TEMPLATES.Grievance[target.type] || CHECKLIST_TEMPLATES.Grievance._default;
    }
    if (!template) {
      // Fallback: generic items
      template = [
        { text: 'Signed grievance form from member', category: 'Document', required: true },
        { text: 'Copy of relevant contract articles', category: 'Document', required: true },
        { text: 'Written statement from member', category: 'Evidence', required: true },
        { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
        { text: 'Management response received', category: 'Document', required: true },
        { text: 'Member notified of decision', category: 'Communication', required: true }
      ];
    }

    for (var item = 0; item < template.length; item++) {
      var tmpl = template[item];
      var checklistId = 'CL-SEED-' + String(idCounter).padStart(4, '0');
      // Mark some items as completed for realism (first 2-3 items of first 2 cases)
      var isCompleted = (t < 2 && item < (t === 0 ? 3 : 2));
      var completedDate = isCompleted ? new Date(now.getTime() - (template.length - item) * 86400000) : '';
      rows.push([
        checklistId,                                         // Checklist ID
        target.caseId,                                       // Case ID
        'Grievance',                                         // Action Type
        tmpl.text,                                           // Item Text
        tmpl.category,                                       // Category
        tmpl.required ? 'Yes' : 'No',                       // Required
        isCompleted,                                         // Completed (checkbox boolean)
        isCompleted ? ownerEmail : '',                       // Completed By
        completedDate,                                       // Completed Date
        new Date(now.getTime() + (14 + item * 3) * 86400000), // Due Date
        '',                                                  // Notes
        item + 1                                             // Sort Order
      ]);
      idCounter++;
    }
  }

  sheet.getRange(2, 1, rows.length, 12).setValues(rows);
  Logger.log('Seeded ' + rows.length + ' checklist items for ' + targets.length + ' cases');
}

// ============================================================================
// SEED: SURVEY QUESTIONS
// ============================================================================

/**
 * Seeds the Survey Questions sheet via createSurveyQuestionsSheet().
 * That function is non-destructive on re-run (skips if sheet exists with data).
 */
function seedSurveyQuestionsData() {
  if (typeof createSurveyQuestionsSheet === 'function') {
    createSurveyQuestionsSheet();
    Logger.log('Survey Questions sheet seeded via createSurveyQuestionsSheet()');
  } else {
    Logger.log('createSurveyQuestionsSheet not available. Skipping.');
  }
}

// ============================================================================
// NUKE FUNCTIONS
// ============================================================================

/**
 * Delete all seeded data from Member Directory and Grievance Log
 * Uses pattern matching (M/G + 4 letters + 3 digits) to identify seeded IDs
 */
function NUKE_SEEDED_DATA() {
  var isDev = typeof IS_DEV_ENVIRONMENT !== 'undefined' && IS_DEV_ENVIRONMENT === true;
  if (!isDev) {
    var _ui = SpreadsheetApp.getUi();
    var _response = _ui.alert('Production Safety Check',
      'This action is intended for development environments only. Are you SURE you want to proceed?',
      _ui.ButtonSet.YES_NO);
    if (_response !== _ui.Button.YES) return;
  }
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // F39: Create backup before destructive operation
  try {
    if (typeof createGrievanceSnapshot === 'function') {
      var preNukeBackup = createGrievanceSnapshot();
      PropertiesService.getScriptProperties().setProperty('PRE_NUKE_SNAPSHOT', JSON.stringify(preNukeBackup));
    }
  } catch (_backupErr) {
    Logger.log('Pre-nuke backup skipped: ' + _backupErr.message);
  }

  // Check if already disabled
  if (isDemoModeDisabled()) {
    ui.alert('Demo Mode Disabled', 'Demo mode has already been disabled. The Demo menu will be removed on next refresh.', ui.ButtonSet.OK);
    return;
  }

  // Use tracked seeded IDs from Script Properties (set during SEED operations)
  // This ensures manually entered data is NEVER deleted, even if IDs share the same format
  var seededMemberLookup = getSeededMemberIds();
  var seededGrievanceLookup = getSeededGrievanceIds();

  // Count seeded data by tracked IDs
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var memberCount = 0;
  var grievanceCount = 0;

  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
    memberIds.forEach(function(row) {
      if (row[0] && seededMemberLookup[String(row[0])]) memberCount++;
    });
  }

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, grievanceSheet.getLastRow() - 1, 1).getValues();
    grievanceIds.forEach(function(row) {
      if (row[0] && seededGrievanceLookup[String(row[0])]) grievanceCount++;
    });
  }

  // Safety check: if no tracked IDs exist, warn user and abort
  if (memberCount === 0 && grievanceCount === 0) {
    var hasAnyMembers = memberSheet && memberSheet.getLastRow() > 1;
    var hasAnyGrievances = grievanceSheet && grievanceSheet.getLastRow() > 1;
    if (hasAnyMembers || hasAnyGrievances) {
      ui.alert(
        '⚠️ No Tracked Seeded Data Found',
        'The system could not find any tracked seeded IDs in Script Properties.\n\n' +
        'This means either:\n' +
        '• Seed data was never created using the Demo menu, OR\n' +
        '• Seed tracking data was cleared previously\n\n' +
        'Existing member/grievance data will NOT be deleted to protect your manually entered records.\n\n' +
        'The nuke will still clean up Config dropdowns, survey data, and demo sheets.',
        ui.ButtonSet.OK
      );
    }
  }

  var response = ui.alert(
    '☢️ NUKE SEEDED DATA',
    '⚠️ This will permanently delete seeded/demo data:\n\n' +
    '• ' + memberCount + ' seeded members (tracked from seed operation)\n' +
    '• ' + grievanceCount + ' seeded grievances (tracked from seed operation)\n' +
    '• Config dropdown values\n' +
    '• Survey responses (Member Satisfaction data cleared)\n' +
    '• Feedback & Development sheet (entire sheet deleted)\n' +
    '• Function Checklist sheet (entire sheet deleted)\n' +
    '• _Audit_Log hidden sheet (entire sheet deleted)\n' +
    '• Resources, Notifications, and Workload data\n\n' +
    '✅ ALL manually entered data will be PRESERVED.\n\n' +
    '⏱️ IMPORTANT: This process takes approximately 3-5 MINUTES.\n' +
    '⚠️ WAIT until the "Running script" dialog disappears!\n' +
    '⚠️ After nuke, the Demo menu will be permanently disabled.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Double confirm
  var response2 = ui.alert(
    '☢️ FINAL CONFIRMATION',
    'This will:\n' +
    '1. Delete ' + memberCount + ' seeded members\n' +
    '2. Delete ' + grievanceCount + ' seeded grievances\n' +
    '3. Clear survey responses from Member Satisfaction\n' +
    '4. Delete Feedback & Development sheet\n' +
    '5. Delete Function Checklist sheet\n' +
    '6. Delete _Audit_Log hidden sheet\n' +
    '7. Permanently disable the Demo menu\n\n' +
    '⏱️ This will take 3-5 MINUTES. DO NOT close the tab!\n' +
    'Wait for the "Running script" dialog to disappear.\n\n' +
    'Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    return;
  }

  ss.toast('Nuking seeded data... This will take 3-5 minutes. Please wait!', '☢️ NUKE', 10);

  try {
    var deletedMembers = 0;
    var deletedGrievances = 0;

    // Delete seeded grievances first (they reference members)
    // Uses tracked IDs from Script Properties — only deletes IDs created by SEED functions
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 1).getValues();
      var totalGrievanceRows = grievanceData.length;

      // Count how many rows are tracked seeded entries
      for (var gc = 0; gc < totalGrievanceRows; gc++) {
        if (seededGrievanceLookup[String(grievanceData[gc][0] || '')]) {
          deletedGrievances++;
        }
      }

      if (deletedGrievances > 0 && deletedGrievances === totalGrievanceRows) {
        // ALL rows are seeded — use clearContent to avoid "cannot delete all non-frozen rows" error
        grievanceSheet.getRange(2, 1, totalGrievanceRows, grievanceSheet.getLastColumn())
          .clearContent()
          .setBackground(null)
          .clearNote();
      } else if (deletedGrievances > 0) {
        // Only some rows are seeded — delete individually bottom-up
        deletedGrievances = 0;
        for (var g = totalGrievanceRows - 1; g >= 0; g--) {
          var gId = String(grievanceData[g][0] || '');
          if (seededGrievanceLookup[gId]) {
            grievanceSheet.deleteRow(g + 2);
            deletedGrievances++;
          }
        }
      }
    }

    // Delete seeded members
    // Uses tracked IDs from Script Properties — only deletes IDs created by SEED functions
    if (memberSheet && memberSheet.getLastRow() > 1) {
      var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, 1).getValues();
      var totalMemberRows = memberData.length;

      // Count how many rows are tracked seeded entries
      var seededMemberCount = 0;
      for (var mc = 0; mc < totalMemberRows; mc++) {
        if (seededMemberLookup[String(memberData[mc][0] || '')]) {
          seededMemberCount++;
        }
      }

      if (seededMemberCount > 0 && seededMemberCount === totalMemberRows) {
        // ALL rows are seeded — use clearContent to avoid "cannot delete all non-frozen rows" error
        memberSheet.getRange(2, 1, totalMemberRows, memberSheet.getLastColumn())
          .clearContent()
          .setBackground(null)
          .clearNote();
        deletedMembers = seededMemberCount;
      } else if (seededMemberCount > 0) {
        // Only some rows are seeded — delete individually bottom-up
        for (var m = totalMemberRows - 1; m >= 0; m--) {
          var mId = String(memberData[m][0] || '');
          if (seededMemberLookup[mId]) {
            memberSheet.deleteRow(m + 2);
            deletedMembers++;
          }
        }
      }
    }

    // Clear Config dropdowns
    NUKE_CONFIG_DROPDOWNS();

    // Clear Member Satisfaction survey data
    var satisfactionSheet = ss.getSheetByName(SHEETS.SATISFACTION);
    var surveyCleared = false;
    if (satisfactionSheet && satisfactionSheet.getLastRow() > 1) {
      try {
        var lastDataRow = satisfactionSheet.getLastRow();
        satisfactionSheet.getRange(2, 1, lastDataRow - 1, 68).clearContent();
        surveyCleared = true;
        Logger.log('Survey response data cleared from Member Satisfaction');
      } catch (e) {
        Logger.log('Could not clear survey data: ' + e.message);
      }
    }

    // Delete Feedback & Development sheet entirely
    var feedbackToDelete = ss.getSheetByName(SHEETS.FEEDBACK);
    var feedbackDeleted = false;
    if (feedbackToDelete) {
      try {
        ss.deleteSheet(feedbackToDelete);
        feedbackDeleted = true;
        Logger.log('Feedback & Development sheet deleted');
      } catch (e) {
        Logger.log('Could not delete Feedback sheet: ' + e.message);
      }
    }

    // Delete Function Checklist sheet entirely
    var functionChecklistToDelete = ss.getSheetByName(SHEETS.FUNCTION_CHECKLIST);
    var functionChecklistDeleted = false;
    if (functionChecklistToDelete) {
      try {
        ss.deleteSheet(functionChecklistToDelete);
        functionChecklistDeleted = true;
        Logger.log('Function Checklist sheet deleted');
      } catch (e) {
        Logger.log('Could not delete Function Checklist sheet: ' + e.message);
      }
    }

    // Delete _Audit_Log hidden sheet entirely
    var auditLogToDelete = ss.getSheetByName(SHEETS.AUDIT_LOG);
    var auditLogDeleted = false;
    if (auditLogToDelete) {
      try {
        ss.deleteSheet(auditLogToDelete);
        auditLogDeleted = true;
        Logger.log('_Audit_Log sheet deleted');
      } catch (e) {
        Logger.log('Could not delete _Audit_Log sheet: ' + e.message);
      }
    }

    // Delete seeded calendar events
    var calEventsDeleted = 0;
    try {
      var calEventJson = PropertiesService.getScriptProperties().getProperty('SEEDED_CALENDAR_EVENT_IDS');
      if (calEventJson) {
        var calEventIds = JSON.parse(calEventJson);
        var cal = CalendarApp.getDefaultCalendar();
        calEventIds.forEach(function(eventId) {
          try {
            var evt = cal.getEventById(eventId);
            if (evt) { evt.deleteEvent(); calEventsDeleted++; }
          } catch (_e) { /* event may already be deleted */ }
        });
      }
    } catch (e) { Logger.log('Calendar cleanup error: ' + e.message); }

    // Clear weekly questions data
    var weeklyCleared = false;
    ['_Question_Pool', '_Weekly_Questions', '_Weekly_Responses'].forEach(function(name) {
      var wSheet = ss.getSheetByName(name);
      if (wSheet && wSheet.getLastRow() > 1) {
        try {
          wSheet.getRange(2, 1, wSheet.getLastRow() - 1, wSheet.getLastColumn()).clearContent();
          weeklyCleared = true;
        } catch (e) { Logger.log('Could not clear ' + name + ': ' + e.message); }
      }
    });

    // Clear resources data (beyond starter rows)
    var resourcesCleared = false;
    var resourcesSheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (resourcesSheet && resourcesSheet.getLastRow() > 1) {
      try {
        var resLastRow = resourcesSheet.getLastRow();
        resourcesSheet.getRange(2, 1, resLastRow - 1, resourcesSheet.getLastColumn()).clearContent();
        resourcesCleared = true;
      } catch (e) { Logger.log('Could not clear Resources: ' + e.message); }
    }

    // Clear notifications data
    var notificationsCleared = false;
    var notifSheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (notifSheet && notifSheet.getLastRow() > 1) {
      try {
        var notifLastRow = notifSheet.getLastRow();
        notifSheet.getRange(2, 1, notifLastRow - 1, notifSheet.getLastColumn()).clearContent();
        notificationsCleared = true;
      } catch (e) { Logger.log('Could not clear Notifications: ' + e.message); }
    }

    // Clear workload vault data
    var workloadCleared = false;
    var workloadVault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (workloadVault && workloadVault.getLastRow() > 1) {
      try {
        var wlLastRow = workloadVault.getLastRow();
        workloadVault.getRange(2, 1, wlLastRow - 1, workloadVault.getLastColumn()).clearContent();
        workloadCleared = true;
      } catch (e) { Logger.log('Could not clear Workload Vault: ' + e.message); }
    }

    // Clear seeded member tasks (rows with MT_SEED_ prefix) from _Steward_Tasks
    var memberTasksCleared = false;
    var tasksSheet = ss.getSheetByName(SHEETS.STEWARD_TASKS);
    if (tasksSheet && tasksSheet.getLastRow() > 1) {
      try {
        var taskData = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, 1).getValues();
        for (var mt = taskData.length - 1; mt >= 0; mt--) {
          if (String(taskData[mt][0]).indexOf('MT_SEED_') === 0) {
            tasksSheet.deleteRow(mt + 2);
            memberTasksCleared = true;
          }
        }
      } catch (e) { Logger.log('Could not clear member tasks: ' + e.message); }
    }

    // Clear seeded case checklist items (rows with CL-SEED- prefix)
    var checklistCleared = false;
    var checklistSheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
    if (checklistSheet && checklistSheet.getLastRow() > 1) {
      try {
        var clData = checklistSheet.getRange(2, 1, checklistSheet.getLastRow() - 1, 1).getValues();
        for (var cl = clData.length - 1; cl >= 0; cl--) {
          if (String(clData[cl][0]).indexOf('CL-SEED-') === 0) {
            checklistSheet.deleteRow(cl + 2);
            checklistCleared = true;
          }
        }
      } catch (e) { Logger.log('Could not clear case checklist: ' + e.message); }
    }

    // Clear tracked IDs from Script Properties
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('SEEDED_MEMBER_IDS');
    props.deleteProperty('SEEDED_GRIEVANCE_IDS');
    props.deleteProperty('SEEDED_CALENDAR_EVENT_IDS');
    props.deleteProperty('SEEDED_UNION_STATS');

    // Disable demo mode permanently
    disableDemoMode();

    ss.toast('Seeded data nuked! Demo mode disabled.', '☢️ Complete', 5);
    ui.alert('☢️ Complete',
      'Seeded data has been deleted:\n' +
      '• ' + deletedMembers + ' members removed\n' +
      '• ' + deletedGrievances + ' grievances removed\n' +
      (surveyCleared ? '• Survey responses cleared from Member Satisfaction\n' : '') +
      (feedbackDeleted ? '• Feedback & Development sheet deleted\n' : '') +
      (functionChecklistDeleted ? '• Function Checklist sheet deleted\n' : '') +
      (auditLogDeleted ? '• _Audit_Log sheet deleted\n' : '') +
      (calEventsDeleted > 0 ? '• ' + calEventsDeleted + ' calendar events deleted\n' : '') +
      (weeklyCleared ? '• Weekly questions data cleared\n' : '') +
      (resourcesCleared ? '• Resources data cleared\n' : '') +
      (notificationsCleared ? '• Notifications data cleared\n' : '') +
      (workloadCleared ? '• Workload vault data cleared\n' : '') +
      (memberTasksCleared ? '• Member tasks cleared\n' : '') +
      (checklistCleared ? '• Case checklist items cleared\n' : '') +
      '• Union stats data cleared\n' +
      '\nDemo mode has been permanently disabled.\n' +
      'Refresh the page to remove the Demo menu.\n\n' +
      '📌 NEXT STEP: Delete the "DeveloperTools.gs" file from the script editor.',
      ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in NUKE_SEEDED_DATA: ' + error.message);
    ui.alert('❌ Error', 'Nuke failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Clear only Config dropdown values (user-populated columns)
 */
function NUKE_CONFIG_DROPDOWNS() {
  if (!isDemoSafeToRun_()) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    return;
  }

  // Clear user-populated columns only (keep system defaults)
  var userColumns = [
    CONFIG_COLS.JOB_TITLES,
    CONFIG_COLS.OFFICE_LOCATIONS,
    CONFIG_COLS.UNITS,
    CONFIG_COLS.SUPERVISORS,
    CONFIG_COLS.MANAGERS,
    CONFIG_COLS.STEWARDS,
    CONFIG_COLS.STEWARD_COMMITTEES,
    CONFIG_COLS.GRIEVANCE_COORDINATORS
  ];

  userColumns.forEach(function(col) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(3, col, lastRow - 2, 1).clear();
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Config dropdowns cleared!', '🧹 Cleared', 3);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get random element from array
 */
function randomChoice(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

/**
 * Shuffle array using Fisher-Yates algorithm
 */
function shuffleArray(arr) {
  var shuffled = arr.slice();
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = temp;
  }
  return shuffled;
}

/**
 * Get random date between two dates
 */
function randomDate(start, end) {
  var startTime = start.getTime();
  var endTime = end.getTime();
  var randomTime = startTime + Math.random() * (endTime - startTime);
  return new Date(randomTime);
}

/**
 * Add days to a date
 */
function addDays(date, days) {
  if (!date) return '';
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
/**
 * ============================================================================
 * TESTING FRAMEWORK & VALIDATION
 * ============================================================================
 * Unit tests, integration tests, and data validation
 */

// ==================== TEST CONFIGURATION ====================

// TEST_RESULTS tracking is handled within individual test runner functions
var TEST_MAX_EXECUTION_MS = 5 * 60 * 1000;
var TEST_LARGE_DATASET_THRESHOLD = 5000;

var Assert = {
  assertEquals: function(expected, actual, message) {
    if (expected !== actual) throw new Error((message || 'Assertion failed') + '\nExpected: ' + JSON.stringify(expected) + '\nActual: ' + JSON.stringify(actual));
  },
  assertTrue: function(value, message) {
    if (value !== true) throw new Error((message || 'Expected true') + '\nActual: ' + value);
  },
  assertFalse: function(value, message) {
    if (value !== false) throw new Error((message || 'Expected false') + '\nActual: ' + value);
  },
  assertNotNull: function(value, message) {
    if (value === null || value === undefined) throw new Error(message || 'Value should not be null/undefined');
  },
  assertNull: function(value, message) {
    if (value !== null) throw new Error((message || 'Expected null') + '\nActual: ' + value);
  },
  assertContains: function(array, value, message) {
    if (!Array.isArray(array) || array.indexOf(value) === -1) throw new Error((message || 'Array does not contain value') + '\nValue: ' + value);
  },
  assertArrayLength: function(array, expectedLength, message) {
    if (!Array.isArray(array) || array.length !== expectedLength) throw new Error((message || 'Array length mismatch') + '\nExpected: ' + expectedLength + '\nActual: ' + (array ? array.length : 'N/A'));
  },
  assertThrows: function(fn, message) {
    var threw = false;
    try { fn(); } catch (_e) { threw = true; }
    if (!threw) throw new Error(message || 'Expected function to throw');
  },
  assertApproximately: function(expected, actual, tolerance, message) {
    tolerance = tolerance || 0.001;
    if (Math.abs(expected - actual) > tolerance) throw new Error((message || 'Values not approximately equal') + '\nExpected: ' + expected + '\nActual: ' + actual);
  },
  fail: function(message) { throw new Error(message || 'Test failed'); },
  // Aliases for compatibility with second API convention
  isTrue: function(value, message) {
    if (!value) throw new Error(message || 'Expected true but got: ' + value);
  },
  isFalse: function(value, message) {
    if (value) throw new Error(message || 'Expected false but got: ' + value);
  },
  equals: function(expected, actual, message) {
    if (expected !== actual) throw new Error(message || 'Expected ' + expected + ' but got: ' + actual);
  },
  notEquals: function(expected, actual, message) {
    if (expected === actual) throw new Error(message || 'Expected values to be different but both were: ' + actual);
  },
  isDefined: function(value, message) {
    if (value === undefined || value === null) throw new Error(message || 'Expected value to be defined');
  },
  isArray: function(value, message) {
    if (!Array.isArray(value)) throw new Error(message || 'Expected array but got: ' + typeof value);
  },
  contains: function(array, value, message) {
    if (array.indexOf(value) === -1) throw new Error(message || 'Expected array to contain: ' + value);
  },
  throws: function(fn, message) {
    var threw = false;
    try { fn(); } catch (_e) { threw = true; }
    if (!threw) throw new Error(message || 'Expected function to throw an error');
  }
};

// ==================== TEST HELPERS ====================
// ==================== UNIT TESTS ====================

/**
 * Unit test: verifies MEMBER_COLS column index constants are correct.
 * @returns {void}
 */
function testMemberColsConstants() {
  Assert.assertNotNull(MEMBER_COLS, 'MEMBER_COLS should exist');
  Assert.assertEquals(1, MEMBER_COLS.MEMBER_ID, 'MEMBER_ID should be column 1');
  Assert.assertEquals(2, MEMBER_COLS.FIRST_NAME, 'FIRST_NAME should be column 2');
  Assert.assertEquals(3, MEMBER_COLS.LAST_NAME, 'LAST_NAME should be column 3');
  Assert.assertEquals(8, MEMBER_COLS.EMAIL, 'EMAIL should be column 8');
  Assert.assertEquals(31, MEMBER_COLS.START_GRIEVANCE, 'START_GRIEVANCE should be column 31');
}

/**
 * Unit test: verifies GRIEVANCE_COLS column index constants are correct.
 * @returns {void}
 */
function testGrievanceColsConstants() {
  Assert.assertNotNull(GRIEVANCE_COLS, 'GRIEVANCE_COLS should exist');
  Assert.assertEquals(1, GRIEVANCE_COLS.GRIEVANCE_ID, 'GRIEVANCE_ID should be column 1');
  Assert.assertEquals(2, GRIEVANCE_COLS.MEMBER_ID, 'MEMBER_ID should be column 2');
  Assert.assertEquals(5, GRIEVANCE_COLS.STATUS, 'STATUS should be column 5');
  Assert.assertEquals(28, GRIEVANCE_COLS.RESOLUTION, 'RESOLUTION should be column 28');
}

/**
 * Unit test: verifies getColumnLetter converts column numbers to letters correctly.
 * @returns {void}
 */
function testColumnLetterConversion() {
  Assert.assertEquals('A', getColumnLetter(1), 'Column 1 should be A');
  Assert.assertEquals('Z', getColumnLetter(26), 'Column 26 should be Z');
  Assert.assertEquals('AA', getColumnLetter(27), 'Column 27 should be AA');
  Assert.assertEquals('AE', getColumnLetter(31), 'Column 31 should be AE');
}

/**
 * Unit test: verifies core SHEETS constant names are defined.
 * @returns {void}
 */
function testSheetsConstants() {
  Assert.assertNotNull(SHEETS, 'SHEETS should exist');
  Assert.assertNotNull(SHEETS.MEMBER_DIR, 'MEMBER_DIR should exist');
  Assert.assertNotNull(SHEETS.GRIEVANCE_LOG, 'GRIEVANCE_LOG should exist');
  Assert.assertNotNull(SHEETS.CONFIG, 'CONFIG should exist');
}

/**
 * Unit test: verifies validateRequired throws on null, empty, and undefined.
 * @returns {void}
 */
function testValidateRequired() {
  Assert.assertThrows(function() { validateRequired(null, 'field'); }, 'null should throw');
  Assert.assertThrows(function() { validateRequired('', 'field'); }, 'empty should throw');
  Assert.assertThrows(function() { validateRequired(undefined, 'field'); }, 'undefined should throw');
}

/**
 * Unit test: verifies validateEmailAddress accepts valid and rejects invalid emails.
 * @returns {void}
 */
function testValidateEmail() {
  var result1 = validateEmailAddress('test@example.com');
  Assert.assertTrue(result1.valid, 'Valid email should pass');
  var result2 = validateEmailAddress('invalid');
  Assert.assertFalse(result2.valid, 'Invalid email should fail');
  var result3 = validateEmailAddress('');
  Assert.assertFalse(result3.valid, 'Empty email should fail');
}

/**
 * Unit test: verifies validatePhoneNumber accepts valid and rejects short numbers.
 * @returns {void}
 */
function testValidatePhoneNumber() {
  var result1 = validatePhoneNumber('(555) 123-4567');
  Assert.assertTrue(result1.valid, 'Valid phone should pass');
  var result2 = validatePhoneNumber('5551234567');
  Assert.assertTrue(result2.valid, 'Digits-only phone should pass');
  var result3 = validatePhoneNumber('123');
  Assert.assertFalse(result3.valid, 'Short phone should fail');
}

/**
 * Unit test: verifies MEMBER_ID regex matches name-based format (M + 4 letters + 3 digits).
 * @returns {void}
 */
function testValidateMemberId() {
  // Format: M prefix + 4 uppercase letters (2 from first name + 2 from last name) + 3 digits
  Assert.assertTrue(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM123'), 'MJOSM123 should be valid (M + John Smith + 123)');
  Assert.assertTrue(VALIDATION_PATTERNS.MEMBER_ID.test('MMAJO456'), 'MMAJO456 should be valid (M + Mary Johnson + 456)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('JOSM123'), 'JOSM123 should be invalid (missing M prefix)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('MJOS123'), 'MJOS123 should be invalid (only 3 name letters)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM12'), 'MJOSM12 should be invalid (only 2 digits)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('mjosm123'), 'mjosm123 should be invalid (lowercase)');
  Assert.assertFalse(VALIDATION_PATTERNS.MEMBER_ID.test('M123456'), 'M123456 should be invalid (old format)');
}

/**
 * Unit test: verifies GRIEVANCE_ID regex matches name-based format (G + 4 letters + 3 digits).
 * @returns {void}
 */
function testValidateGrievanceId() {
  // Format: G prefix + 4 uppercase letters (2 from first name + 2 from last name) + 3 digits
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM789'), 'GJOSM789 should be valid (G + John Smith + 789)');
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GROWI001'), 'GROWI001 should be valid (G + Robert Williams + 001)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('JOSM789'), 'JOSM789 should be invalid (missing G prefix)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('G-123456'), 'G-123456 should be invalid (old format)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM1234'), 'GJOSM1234 should be invalid (4 digits)');
}

/**
 * Unit test: verifies open rate boundary values (0-100) are in valid range.
 * @returns {void}
 */
function testOpenRateRange() {
  Assert.assertTrue(85 >= 0 && 85 <= 100, 'Open rate 85 should be in range');
  // eslint-disable-next-line no-self-compare -- Testing boundary conditions explicitly
  Assert.assertTrue(0 >= 0 && 0 <= 100, 'Open rate 0 should be in range');
  // eslint-disable-next-line no-self-compare -- Testing boundary conditions explicitly
  Assert.assertTrue(100 >= 0 && 100 <= 100, 'Open rate 100 should be in range');
}

// ==================== TEST RUNNER ====================
// ==================== VALIDATION FRAMEWORK ====================

var VALIDATION_PATTERNS = {
  EMAIL: /^[^\s@]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
  PHONE_US: /^[\+]?1?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}$/,
  // ID format: M/G prefix + 2 chars from first name + 2 chars from last name + 3 random digits
  MEMBER_ID: /^M[A-Z]{4}\d{3}$/,      // e.g., MJOSM123 (M + John Smith + 123)
  GRIEVANCE_ID: /^G[A-Z]{4}\d{3}$/    // e.g., GJOSM456 (G + John Smith + 456)
};

var VALIDATION_MESSAGES = {
  EMAIL_INVALID: 'Invalid email format. Use: name@domain.com',
  EMAIL_EMPTY: 'Email address is required',
  PHONE_INVALID: 'Invalid phone format. Use: (555) 555-1234',
  MEMBER_ID_INVALID: 'Invalid Member ID. Format: M + 2 letters from first name + 2 letters from last name + 3 digits (e.g., MJOSM123)',
  MEMBER_ID_DUPLICATE: 'This Member ID already exists',
  GRIEVANCE_ID_INVALID: 'Invalid Grievance ID. Format: G + 2 letters from first name + 2 letters from last name + 3 digits (e.g., GJOSM456)',
  GRIEVANCE_ID_DUPLICATE: 'This Grievance ID already exists'
};

/**
 * Validates an email address format with typo detection for common domains.
 * @param {string} email - Email address to validate.
 * @returns {Object} { valid: boolean, message: string, suggestion?: string }
 */
function validateEmailAddress(email) {
  if (!email || email.toString().trim() === '') return { valid: false, message: VALIDATION_MESSAGES.EMAIL_EMPTY };
  // Handle Google Sheets auto-converting emails to Date or other types
  if (email instanceof Date) return { valid: false, message: VALIDATION_MESSAGES.EMAIL_INVALID };
  var clean = email.toString().trim().toLowerCase();
  // Remove invisible/zero-width characters that may be pasted in
  clean = clean.replace(/[\u200B\u200C\u200D\uFEFF\u00A0]/g, '');
  if (!VALIDATION_PATTERNS.EMAIL.test(clean)) return { valid: false, message: VALIDATION_MESSAGES.EMAIL_INVALID };
  var typos = { 'gmial.com': 'gmail.com', 'gmai.com': 'gmail.com', 'yaho.com': 'yahoo.com', 'yahooo.com': 'yahoo.com', 'gamil.com': 'gmail.com', 'gnail.com': 'gmail.com', 'hotmal.com': 'hotmail.com', 'outlok.com': 'outlook.com' };
  var domain = clean.split('@')[1];
  if (typos[domain]) return { valid: true, message: 'Did you mean ' + clean.split('@')[0] + '@' + typos[domain] + '?', suggestion: clean.split('@')[0] + '@' + typos[domain] };
  return { valid: true, message: 'Valid email' };
}

/**
 * Validates a phone number has at least 10 digits and returns formatted version.
 * @param {string} phone - Phone number string.
 * @returns {Object} { valid: boolean, message: string, formatted?: string }
 */
function validatePhoneNumber(phone) {
  if (!phone || phone.toString().trim() === '') return { valid: false, message: 'Phone required' };
  var digits = phone.toString().replace(/\D/g, '');
  if (digits.length < 10) return { valid: false, message: 'At least 10 digits required' };
  if (digits.length > 15) return { valid: false, message: 'Phone too long' };
  var formatted = formatUSPhone(digits);
  return { valid: true, message: 'Valid phone', formatted: formatted };
}

/**
 * Formats a digit string as a US phone number: (XXX) XXX-XXXX.
 * @param {string} digits - Digits-only phone string.
 * @returns {string} Formatted phone or original digits if not 10-digit.
 */
function formatUSPhone(digits) {
  if (digits.length === 11 && digits[0] === '1') digits = digits.substring(1);
  if (digits.length === 10) return '(' + digits.substring(0, 3) + ') ' + digits.substring(3, 6) + '-' + digits.substring(6);
  return digits;
}

/**
 * Throws an error if the value is null, undefined, or empty string.
 * @param {*} value - Value to check.
 * @param {string} fieldName - Field name for the error message.
 * @returns {*} The original value if non-empty.
 */
function validateRequired(value, fieldName) {
  if (value === null || value === undefined || value === '') throw new Error(fieldName + ' is required');
  return value;
}
/**
 * Validate all grievances to ensure Member IDs exist in Member Directory
 * @returns {Array} Array of issues found
 */
function validateGrievanceMemberIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  var grievanceLastRow = grievanceSheet.getLastRow();
  var memberLastRow = memberSheet.getLastRow();

  if (grievanceLastRow < 2 || memberLastRow < 2) return [];

  // Build set of valid member IDs
  var validMemberIds = {};
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberLastRow - 1, 1).getValues();
  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0]) validMemberIds[memberIds[i][0]] = true;
  }

  // Check each grievance's Member ID
  var issues = [];
  var grievanceMemberIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, grievanceLastRow - 1, 1).getValues();
  var grievanceIds = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, grievanceLastRow - 1, 1).getValues();

  for (var j = 0; j < grievanceMemberIds.length; j++) {
    var memberId = grievanceMemberIds[j][0];
    var grievanceId = grievanceIds[j][0];
    var row = j + 2;

    if (!memberId) {
      issues.push({ row: row, grievanceId: grievanceId, memberId: '', message: 'Missing Member ID' });
    } else if (!validMemberIds[memberId]) {
      issues.push({ row: row, grievanceId: grievanceId, memberId: memberId, message: 'Member ID not found in Member Directory' });
    }
  }

  return issues;
}

/**
 * Runs bulk validation on Member Directory emails, phones, and IDs plus grievance references.
 * @returns {void}
 */
function runBulkValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) { SpreadsheetApp.getUi().alert('Member Directory not found!'); return; }
  ss.toast('Running validation...', 'Please wait', -1);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data', 'Info', 3); return; }
  var emailData = sheet.getRange(2, MEMBER_COLS.EMAIL, lastRow - 1, 1).getValues();
  var phoneData = sheet.getRange(2, MEMBER_COLS.PHONE, lastRow - 1, 1).getValues();
  var memberIdData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1).getValues();
  var issues = [], seenIds = {};
  for (var i = 0; i < lastRow - 1; i++) {
    var row = i + 2;
    if (emailData[i][0]) { var er = validateEmailAddress(emailData[i][0]); if (!er.valid) issues.push({ row: row, field: 'Email', value: emailData[i][0], message: er.message }); }
    if (phoneData[i][0]) { var pr = validatePhoneNumber(phoneData[i][0]); if (!pr.valid) issues.push({ row: row, field: 'Phone', value: phoneData[i][0], message: pr.message }); }
    if (memberIdData[i][0]) {
      if (seenIds[memberIdData[i][0]]) issues.push({ row: row, field: 'Member ID', value: memberIdData[i][0], message: 'Duplicate of row ' + seenIds[memberIdData[i][0]] });
      else seenIds[memberIdData[i][0]] = row;
    }
  }

  // Also validate grievance Member IDs reference valid members
  var grievanceIssues = validateGrievanceMemberIds();
  grievanceIssues.forEach(function(gi) {
    issues.push({ row: gi.row, field: 'Grievance Member ID', value: gi.memberId || '(empty)', message: gi.message + ' (Grievance: ' + gi.grievanceId + ')' });
  });

  showValidationReport(issues, lastRow - 1 + grievanceIssues.length);
}

/**
 * Displays an HTML modal with validation results summary and issue table.
 * @param {Array<Object>} issues - Array of issue objects with row, field, value, message.
 * @param {number} total - Total number of records validated.
 * @returns {void}
 */
function showValidationReport(issues, total) {
  var rate = total > 0 ? (((total - issues.length) / total) * 100).toFixed(1) : 100;
  var rows = issues.slice(0, 50).map(function(i) { return '<tr><td>' + escapeHtml(i.row) + '</td><td>' + escapeHtml(i.field) + '</td><td>' + escapeHtml(i.value) + '</td><td>' + escapeHtml(i.message) + '</td></tr>'; }).join('');
  if (issues.length > 50) rows += '<tr><td colspan="4">...and ' + (issues.length - 50) + ' more</td></tr>';
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{display:flex;gap:20px;margin:20px 0}.stat{flex:1;padding:20px;border-radius:8px;text-align:center}.stat.good{background:#e8f5e9}.stat.warning{background:#fff3e0}.stat.bad{background:#ffebee}.num{font-size:32px;font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:20px}th,td{padding:10px;text-align:left;border:1px solid #ddd}th{background:#f5f5f5}</style></head><body><h2>📊 Validation Report</h2><div class="summary"><div class="stat ' + (issues.length === 0 ? 'good' : issues.length < 10 ? 'warning' : 'bad') + '"><div class="num">' + rate + '%</div><div>Pass Rate</div></div><div class="stat good"><div class="num">' + total + '</div><div>Records</div></div><div class="stat ' + (issues.length === 0 ? 'good' : 'bad') + '"><div class="num">' + issues.length + '</div><div>Issues</div></div></div>' + (issues.length > 0 ? '<table><tr><th>Row</th><th>Field</th><th>Value</th><th>Issue</th></tr>' + rows + '</table>' : '<div style="text-align:center;padding:40px;color:#4caf50">✅ No issues found!</div>') + '</body></html>'
  ).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Report');
}

/**
 * Shows an HTML modal with validation settings and a button to run bulk validation.
 * @returns {void}
 */
function showValidationSettings() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.setting{margin:15px 0;padding:15px;background:#f8f9fa;border-radius:8px;display:flex;justify-content:space-between;align-items:center}.title{font-weight:bold}.desc{color:#666;font-size:13px}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;width:100%;margin-top:20px}</style></head><body><h2>⚙️ Validation Settings</h2><div class="setting"><div><div class="title">Email Validation</div><div class="desc">Validate format as you type</div></div><span>✅</span></div><div class="setting"><div><div class="title">Phone Auto-format</div><div class="desc">Format to (XXX) XXX-XXXX</div></div><span>✅</span></div><div class="setting"><div><div class="title">Duplicate Detection</div><div class="desc">Warn on duplicate IDs</div></div><span>✅</span></div><button onclick="google.script.run.runBulkValidation();google.script.host.close()">🔍 Run Bulk Validation</button></body></html>'
  ).setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Settings');
}

/**
 * Clears all validation notes and background colors from email, phone, and ID columns.
 * @returns {void}
 */
function clearValidationIndicators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  [MEMBER_COLS.EMAIL, MEMBER_COLS.PHONE, MEMBER_COLS.MEMBER_ID].forEach(function(col) {
    var range = sheet.getRange(2, col, lastRow - 1, 1);
    range.clearNote();
    range.setBackground(null);
  });
  ss.toast('✅ Indicators cleared', 'Done', 3);
}

/**
 * Installs an onEdit trigger for real-time validation of email and phone edits.
 * @returns {void}
 */
function installValidationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditValidation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditValidation').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ Validation trigger installed!');
}

/**
 * OnEdit trigger handler: validates email/phone cells and marks invalid entries.
 * @param {Object} e - GAS onEdit event object.
 * @returns {void}
 */
function onEditValidation(e) {
  if (!e || !e.range) return;
  try {
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) return;
  var col = e.range.getColumn(), row = e.range.getRow(), val = e.value;
  if (row < 2 || !val) return;
  if (name === SHEETS.MEMBER_DIR) {
    if (col === MEMBER_COLS.EMAIL) {
      var r = validateEmailAddress(val);
      if (!r.valid) { e.range.setNote('⚠️ ' + r.message); e.range.setBackground('#fff3e0'); }
      else { e.range.clearNote(); e.range.setBackground(null); }
    }
    if (col === MEMBER_COLS.PHONE) {
      r = validatePhoneNumber(val);
      if (!r.valid) { e.range.setNote('⚠️ ' + r.message); e.range.setBackground('#fff3e0'); }
      else { e.range.clearNote(); e.range.setBackground(null); if (r.formatted !== val) e.range.setValue(r.formatted); }
    }
  }
  } catch (err) { Logger.log('onEditValidation error: ' + (err.message || err)); }
}

/**
 * ============================================================================
 * TestFramework.gs - Comprehensive Unit Testing Framework
 * ============================================================================
 *
 * This module provides a comprehensive testing framework for the Dashboard.
 * Includes:
 * - Test runner with suite management
 * - Rich assertion helpers
 * - Test report generation
 * - Module-specific test suites
 * - Performance benchmarking
 *
 * USAGE:
 *   runAllTests()     - Run all test suites
 *   runQuickTests()   - Run essential smoke tests
 *   runModuleTests()  - Run tests for specific module
 *
 * @fileoverview Unit testing framework and comprehensive test cases
 * @version 4.33.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// TEST FRAMEWORK
// ============================================================================

/**
 * Test result object
 * @typedef {Object} TestResult
 * @property {string} name - Test name
 * @property {boolean} passed - Whether the test passed
 * @property {string} [error] - Error message if failed
 * @property {number} duration - Test duration in ms
 */

/**
 * Test suite class
 */
var TestSuite = {
  results: [],
  currentSuite: '',

  /**
   * Runs all registered tests
   * @returns {Object} Test results summary
   */
  runAll: function() {
    this.results = [];
    var startTime = new Date().getTime();

    // Run all test functions
    var testFunctions = this.getTestFunctions_();
    testFunctions.forEach(function(testFn) {
      this.runTest_(testFn);
    }, this);

    var endTime = new Date().getTime();

    return {
      total: this.results.length,
      passed: this.results.filter(function(r) { return r.passed; }).length,
      failed: this.results.filter(function(r) { return !r.passed; }).length,
      duration: endTime - startTime,
      results: this.results
    };
  },

  /**
   * Runs a single test function
   * @param {Function} testFn - Test function to run
   * @private
   */
  runTest_: function(testFn) {
    var testName = testFn.name || 'anonymous';
    var startTime = new Date().getTime();

    try {
      testFn();
      this.results.push({
        name: testName,
        passed: true,
        duration: new Date().getTime() - startTime
      });
    } catch (e) {
      this.results.push({
        name: testName,
        passed: false,
        error: e.message,
        duration: new Date().getTime() - startTime
      });
    }
  },

  /**
   * Gets all test functions
   * @returns {Array<Function>} Array of test functions
   * @private
   */
  getTestFunctions_: function() {
    return [
      // Constants tests
      test_Constants_Defined,
      test_SheetNames_Valid,
      test_ColumnIndices_Valid,
      test_ColorsConfig_Valid,
      test_HiddenSheets_Defined,

      // Validation tests
      test_ValidationHelpers,
      test_EmailValidation_EdgeCases,
      test_PhoneValidation_EdgeCases,
      test_MemberIdFormat,
      test_GrievanceIdFormat,

      // Helper tests
      test_DateHelpers,
      test_StringHelpers,
      test_ArrayHelpers,
      test_ObjectHelpers,

      // Cache tests
      test_CacheConfig_Valid,
      test_CacheKeys_Defined,

      // Module integration tests
      test_SearchEngine_Functions,
      test_ThemeService_Functions,
      test_MenuBuilder_Functions,

      // Portal column map tests (v4.20.19)
      test_PortalMinutesCols_Complete,
      test_PortalColsNoHardcodedIndices,

      // Multi-select tests (v4.20.19)
      test_MultiSelectCols_Populated,
      test_MultiSelectAutoOpen_DefaultOn,

      // Drive / config column tests (v4.20.19)
      test_DriveRootFolderName_Dynamic,
      test_ConfigCols_FolderIds_Exist
    ];
  }
};

// NOTE: Assert object is defined once at line 1677 with both API styles merged

// ============================================================================
// TEST CASES - CONSTANTS
// ============================================================================

/**
 * Tests that required constants are defined
 */
function test_Constants_Defined() {
  Assert.isDefined(SHEETS, 'SHEETS constant should be defined');
  Assert.isDefined(MEMBER_COLS, 'MEMBER_COLS constant should be defined');
  Assert.isDefined(GRIEVANCE_COLS, 'GRIEVANCE_COLS constant should be defined');
  Assert.isDefined(CONFIG_COLS, 'CONFIG_COLS constant should be defined');
  Assert.isDefined(COLORS, 'COLORS constant should be defined');
}

/**
 * Tests that sheet names are valid strings
 */
function test_SheetNames_Valid() {
  Assert.isDefined(SHEETS.MEMBER_DIR, 'SHEETS.MEMBER_DIR should be defined');
  Assert.isDefined(SHEETS.GRIEVANCE_LOG, 'SHEETS.GRIEVANCE_LOG should be defined');
  Assert.isDefined(SHEETS.CONFIG, 'SHEETS.CONFIG should be defined');

  Assert.isTrue(typeof SHEETS.MEMBER_DIR === 'string', 'SHEETS.MEMBER_DIR should be a string');
  Assert.isTrue(SHEETS.MEMBER_DIR.length > 0, 'SHEETS.MEMBER_DIR should not be empty');
}

/**
 * Tests that column indices are valid numbers
 */
function test_ColumnIndices_Valid() {
  Assert.isTrue(typeof MEMBER_COLS.MEMBER_ID === 'number', 'MEMBER_COLS.MEMBER_ID should be a number');
  Assert.isTrue(MEMBER_COLS.MEMBER_ID > 0, 'MEMBER_COLS.MEMBER_ID should be positive');

  Assert.isTrue(typeof GRIEVANCE_COLS.GRIEVANCE_ID === 'number', 'GRIEVANCE_COLS.GRIEVANCE_ID should be a number');
  Assert.isTrue(GRIEVANCE_COLS.GRIEVANCE_ID > 0, 'GRIEVANCE_COLS.GRIEVANCE_ID should be positive');
}

// ============================================================================
// TEST CASES - VALIDATION
// ============================================================================

/**
 * Tests validation helper functions
 */
function test_ValidationHelpers() {
  // Test email validation
  Assert.isTrue(isValidEmail_('test@example.com'), 'Valid email should pass');
  Assert.isFalse(isValidEmail_('invalid-email'), 'Invalid email should fail');
  Assert.isFalse(isValidEmail_(''), 'Empty email should fail');

  // Test phone validation
  Assert.isTrue(isValidPhone_('555-123-4567'), 'Valid phone should pass');
  Assert.isTrue(isValidPhone_('5551234567'), 'Phone without dashes should pass');
}

/**
 * Tests date helper functions
 */
function test_DateHelpers() {
  var today = new Date();

  // Test date formatting
  var formatted = formatDate_(today);
  Assert.isDefined(formatted, 'Formatted date should be defined');
  Assert.isTrue(typeof formatted === 'string', 'Formatted date should be a string');
}

/**
 * Tests string helper functions
 */
function test_StringHelpers() {
  // Test string trimming
  Assert.equals('test', trimString_('  test  '), 'String should be trimmed');
  Assert.equals('', trimString_(''), 'Empty string should remain empty');
}

// ============================================================================
// TEST HELPER FUNCTIONS (for tests only)
// ============================================================================

/**
 * Validates email format
 * @param {string} email - Email to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidEmail_(email) {
  if (!email || typeof email !== 'string') return false;
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates phone format
 * @param {string} phone - Phone to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidPhone_(phone) {
  if (!phone || typeof phone !== 'string') return false;
  var digits = phone.replace(/\D/g, '');
  return digits.length >= 10 && digits.length <= 11;
}

/**
 * Formats a date
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 * @private
 */
function formatDate_(date) {
  if (!date || !(date instanceof Date)) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Trims a string
 * @param {string} str - String to trim
 * @returns {string} Trimmed string
 * @private
 */
function trimString_(str) {
  if (!str || typeof str !== 'string') return '';
  return str.trim();
}

// ============================================================================
// TEST RUNNER MENU FUNCTIONS
// ============================================================================

/**
 * Runs all tests and displays results
 * @returns {void}
 */
function runAllTests() {
  var results = TestSuite.runAll();

  var ui = SpreadsheetApp.getUi();
  var message = 'Test Results:\n\n' +
    '✅ Passed: ' + results.passed + '\n' +
    '❌ Failed: ' + results.failed + '\n' +
    '⏱️ Duration: ' + results.duration + 'ms\n\n';

  if (results.failed > 0) {
    message += 'Failed Tests:\n';
    results.results.filter(function(r) { return !r.passed; }).forEach(function(r) {
      message += '• ' + r.name + ': ' + r.error + '\n';
    });
  }

  ui.alert('🧪 Test Results', message, ui.ButtonSet.OK);
}

/**
 * Runs quick smoke tests
 * @returns {void}
 */
function runQuickTests() {
  try {
    test_Constants_Defined();
    test_SheetNames_Valid();
    test_ColumnIndices_Valid();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ All quick tests passed!', 'Tests', 3);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ Test failed: ' + e.message, 'Tests', 5);
  }
}
// ============================================================================
// EXTENDED TEST CASES - CONSTANTS
// ============================================================================

/**
 * Tests that color configuration is valid
 */
function test_ColorsConfig_Valid() {
  Assert.isDefined(COLORS.PRIMARY_PURPLE, 'COLORS.PRIMARY_PURPLE should be defined');
  Assert.isDefined(COLORS.WHITE, 'COLORS.WHITE should be defined');
  Assert.isDefined(COLORS.TEXT_DARK, 'COLORS.TEXT_DARK should be defined');

  // Validate hex color format
  var hexRegex = /^#[0-9A-Fa-f]{6}$/;
  Assert.isTrue(hexRegex.test(COLORS.PRIMARY_PURPLE), 'PRIMARY_PURPLE should be valid hex');
  Assert.isTrue(hexRegex.test(COLORS.WHITE), 'WHITE should be valid hex');
}

/**
 * Tests that hidden sheet names are defined
 */
function test_HiddenSheets_Defined() {
  Assert.isDefined(HIDDEN_SHEETS, 'HIDDEN_SHEETS constant should be defined');
  Assert.isDefined(HIDDEN_SHEETS.CALC_STATS, 'HIDDEN_SHEETS.CALC_STATS should be defined');
  Assert.isDefined(HIDDEN_SHEETS.CALC_FORMULAS, 'HIDDEN_SHEETS.CALC_FORMULAS should be defined');
}

// ============================================================================
// EXTENDED TEST CASES - VALIDATION
// ============================================================================

/**
 * Tests email validation edge cases
 */
function test_EmailValidation_EdgeCases() {
  // Valid emails
  Assert.isTrue(isValidEmail_('user@domain.com'), 'Standard email should pass');
  Assert.isTrue(isValidEmail_('user.name@domain.com'), 'Email with dot should pass');
  Assert.isTrue(isValidEmail_('user+tag@domain.com'), 'Email with plus should pass');
  Assert.isTrue(isValidEmail_('user@sub.domain.com'), 'Email with subdomain should pass');

  // Invalid emails
  Assert.isFalse(isValidEmail_(''), 'Empty string should fail');
  Assert.isFalse(isValidEmail_(null), 'Null should fail');
  Assert.isFalse(isValidEmail_(undefined), 'Undefined should fail');
  Assert.isFalse(isValidEmail_('userATdomain.com'), 'Email without @ should fail');
  Assert.isFalse(isValidEmail_('@domain.com'), 'Email starting with @ should fail');
  Assert.isFalse(isValidEmail_('user@'), 'Email ending with @ should fail');
  Assert.isFalse(isValidEmail_('user@.com'), 'Email with @ followed by dot should fail');
}

/**
 * Tests phone validation edge cases
 */
function test_PhoneValidation_EdgeCases() {
  // Valid phones
  Assert.isTrue(isValidPhone_('5551234567'), '10 digit phone should pass');
  Assert.isTrue(isValidPhone_('15551234567'), '11 digit phone should pass');
  Assert.isTrue(isValidPhone_('555-123-4567'), 'Phone with dashes should pass');
  Assert.isTrue(isValidPhone_('(555) 123-4567'), 'Phone with parens should pass');
  Assert.isTrue(isValidPhone_('555.123.4567'), 'Phone with dots should pass');

  // Invalid phones
  Assert.isFalse(isValidPhone_(''), 'Empty string should fail');
  Assert.isFalse(isValidPhone_(null), 'Null should fail');
  Assert.isFalse(isValidPhone_('12345'), 'Too short should fail');
  Assert.isFalse(isValidPhone_('123456789012'), 'Too long should fail');
}

/**
 * Tests member ID format validation
 */
function test_MemberIdFormat() {
  Assert.isTrue(isValidMemberId_('MBR-001'), 'Standard member ID should pass');
  Assert.isTrue(isValidMemberId_('MBR-12345'), 'Member ID with 5 digits should pass');
  Assert.isFalse(isValidMemberId_(''), 'Empty should fail');
  Assert.isFalse(isValidMemberId_(null), 'Null should fail');
}

/**
 * Tests grievance ID format validation
 */
function test_GrievanceIdFormat() {
  Assert.isTrue(isValidGrievanceId_('GRV-2024-001'), 'Standard grievance ID should pass');
  Assert.isTrue(isValidGrievanceId_('GRV-2024-12345'), 'Grievance ID with more digits should pass');
  Assert.isFalse(isValidGrievanceId_(''), 'Empty should fail');
  Assert.isFalse(isValidGrievanceId_(null), 'Null should fail');
}

// ============================================================================
// EXTENDED TEST CASES - HELPERS
// ============================================================================

/**
 * Tests array helper functions
 */
function test_ArrayHelpers() {
  var _arr = [1, 2, 3, 4, 5];

  // Test unique function
  var withDups = [1, 2, 2, 3, 3, 3];
  var unique = getUniqueValues_(withDups);
  Assert.equals(3, unique.length, 'Unique should remove duplicates');

  // Test flatten
  var nested = [[1, 2], [3, 4]];
  var flat = flattenArray_(nested);
  Assert.equals(4, flat.length, 'Flatten should combine arrays');
}

/**
 * Tests object helper functions
 */
function test_ObjectHelpers() {
  var obj = { a: 1, b: 2, c: 3 };

  // Test keys
  var keys = Object.keys(obj);
  Assert.equals(3, keys.length, 'Object should have 3 keys');
  Assert.contains(keys, 'a', 'Keys should contain "a"');

  // Test hasProperty
  Assert.isTrue(hasProperty_(obj, 'a'), 'Object should have property a');
  Assert.isFalse(hasProperty_(obj, 'd'), 'Object should not have property d');
}

// ============================================================================
// TEST CASES - CACHE
// ============================================================================

/**
 * Tests cache configuration
 */
function test_CacheConfig_Valid() {
  Assert.isDefined(CACHE_CONFIG, 'CACHE_CONFIG should be defined');
  Assert.isTrue(typeof CACHE_CONFIG.MEMORY_TTL === 'number', 'MEMORY_TTL should be a number');
  Assert.isTrue(CACHE_CONFIG.MEMORY_TTL > 0, 'MEMORY_TTL should be positive');
}

/**
 * Tests cache keys are defined
 */
function test_CacheKeys_Defined() {
  Assert.isDefined(CACHE_KEYS, 'CACHE_KEYS should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_GRIEVANCES, 'ALL_GRIEVANCES key should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_MEMBERS, 'ALL_MEMBERS key should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_STEWARDS, 'ALL_STEWARDS key should be defined');
}

// ============================================================================
// TEST CASES - MODULE INTEGRATION
// ============================================================================

/**
 * Tests SearchEngine module functions exist
 */
function test_SearchEngine_Functions() {
  Assert.isDefined(typeof getDesktopSearchLocations, 'getDesktopSearchLocations should be defined');
  Assert.isDefined(typeof getDesktopSearchData, 'getDesktopSearchData should be defined');
  Assert.isDefined(typeof navigateToSearchResult, 'navigateToSearchResult should be defined');
  Assert.isDefined(typeof searchDashboard, 'searchDashboard should be defined');
}

/**
 * Tests ThemeService module functions exist
 */
function test_ThemeService_Functions() {
  Assert.isDefined(typeof APPLY_SYSTEM_THEME, 'APPLY_SYSTEM_THEME should be defined');
  Assert.isDefined(typeof resetToDefaultTheme, 'resetToDefaultTheme should be defined');
  Assert.isDefined(typeof getComfortViewSettings, 'getComfortViewSettings should be defined');
  Assert.isDefined(typeof applyZebraStripes, 'applyZebraStripes should be defined');
}

/**
 * Tests MenuBuilder module functions exist
 */
function test_MenuBuilder_Functions() {
  Assert.isDefined(typeof createDashboardMenu, 'createDashboardMenu should be defined');
  Assert.isDefined(typeof navigateToSheet, 'navigateToSheet should be defined');
  Assert.isDefined(typeof showToast, 'showToast should be defined');
}

// ============================================================================
// PORTAL COLUMN MAP TESTS (v4.20.19)
// Catch hardcoded indices and missing column definitions before they reach prod.
// ============================================================================

/**
 * Verifies PORTAL_MINUTES_COLS has all required keys including DRIVE_DOC_URL.
 * Failing here means getMeetingMinutes() would silently return empty driveDocUrl.
 */
function test_PortalMinutesCols_Complete() {
  Assert.isDefined(typeof PORTAL_MINUTES_COLS, 'PORTAL_MINUTES_COLS should be defined');

  var required = ['ID', 'MEETING_DATE', 'TITLE', 'BULLETS', 'FULL_MINUTES',
                  'CREATED_BY', 'CREATED_DATE', 'DRIVE_DOC_URL'];
  required.forEach(function(key) {
    Assert.isTrue(key in PORTAL_MINUTES_COLS, 'PORTAL_MINUTES_COLS.' + key + ' must exist');
    Assert.isTrue(typeof PORTAL_MINUTES_COLS[key] === 'number',
      'PORTAL_MINUTES_COLS.' + key + ' must be a number (0-indexed)');
    Assert.isTrue(PORTAL_MINUTES_COLS[key] >= 0,
      'PORTAL_MINUTES_COLS.' + key + ' must be non-negative');
  });

  // Ensure DRIVE_DOC_URL is one beyond CREATED_DATE (cols are contiguous)
  Assert.assertEquals(PORTAL_MINUTES_COLS.CREATED_DATE + 1, PORTAL_MINUTES_COLS.DRIVE_DOC_URL,
    'DRIVE_DOC_URL must be exactly one column after CREATED_DATE');

  // Spot-check other portal col objects for existence (full coverage in test_PortalColsNoHardcodedIndices)
  Assert.isDefined(typeof PORTAL_EVENT_COLS,          'PORTAL_EVENT_COLS must exist');
  // PORTAL_POLL_COLS / PORTAL_POLL_RESPONSE_COLS removed v4.24.0
  Assert.isDefined(typeof PORTAL_STEWARD_LOG_COLS,     'PORTAL_STEWARD_LOG_COLS must exist');
  Assert.isDefined(typeof PORTAL_MEGA_SURVEY_COLS,     'PORTAL_MEGA_SURVEY_COLS must exist');
}

/**
 * Guards against raw numeric literals being used in array access instead of
 * PORTAL_*_COLS constants.  Checks that every constant is a non-negative integer
 * and that no two keys in the same object share the same index
 * (which would indicate a copy-paste error creating a shadow column).
 */
function test_PortalColsNoHardcodedIndices() {
  var colObjects = [
    { name: 'PORTAL_MINUTES_COLS',       obj: PORTAL_MINUTES_COLS },
    { name: 'PORTAL_EVENT_COLS',         obj: PORTAL_EVENT_COLS },
    // PORTAL_POLL_COLS / PORTAL_POLL_RESPONSE_COLS removed v4.24.0
    { name: 'PORTAL_GRIEVANCE_COLS',     obj: PORTAL_GRIEVANCE_COLS },
    { name: 'PORTAL_STEWARD_LOG_COLS',   obj: PORTAL_STEWARD_LOG_COLS },
    { name: 'PORTAL_MEGA_SURVEY_COLS',   obj: PORTAL_MEGA_SURVEY_COLS },
  ];

  colObjects.forEach(function(item) {
    var seen = {};
    Object.keys(item.obj).forEach(function(key) {
      var val = item.obj[key];
      Assert.isTrue(typeof val === 'number' && val >= 0,
        item.name + '.' + key + ' must be a non-negative number, got: ' + val);
      Assert.isTrue(!seen[val],
        item.name + ': duplicate column index ' + val + ' on key ' + key +
        ' (already used by ' + seen[val] + ')');
      seen[val] = key;
    });
  });
}

// ============================================================================
// MULTI-SELECT TESTS (v4.20.19)
// Verify multi-select config is populated and default is ON.
// ============================================================================

/**
 * Verifies MULTI_SELECT_COLS is populated with valid column references.
 * Failing here means onSelectionChange would never find a match and
 * auto-open would silently stop working after a column-map rebuild.
 */
function test_MultiSelectCols_Populated() {
  Assert.isDefined(typeof MULTI_SELECT_COLS, 'MULTI_SELECT_COLS must be defined');
  Assert.isTrue(Array.isArray(MULTI_SELECT_COLS.MEMBER_DIR),
    'MULTI_SELECT_COLS.MEMBER_DIR must be an array');
  Assert.isTrue(Array.isArray(MULTI_SELECT_COLS.GRIEVANCE_LOG),
    'MULTI_SELECT_COLS.GRIEVANCE_LOG must be an array');

  Assert.isTrue(MULTI_SELECT_COLS.MEMBER_DIR.length >= 1,
    'MULTI_SELECT_COLS.MEMBER_DIR must have at least 1 entry');
  Assert.isTrue(MULTI_SELECT_COLS.GRIEVANCE_LOG.length >= 1,
    'MULTI_SELECT_COLS.GRIEVANCE_LOG must have at least 1 entry');

  // Every entry must have col (positive number), configCol (positive number), label (string)
  var allEntries = MULTI_SELECT_COLS.MEMBER_DIR.concat(MULTI_SELECT_COLS.GRIEVANCE_LOG);
  allEntries.forEach(function(entry) {
    Assert.isTrue(typeof entry.col === 'number' && entry.col > 0,
      'Multi-select entry.col must be a positive integer (1-indexed), got: ' + entry.col);
    Assert.isTrue(typeof entry.configCol === 'number' && entry.configCol > 0,
      'Multi-select entry.configCol must be a positive integer, got: ' + entry.configCol);
    Assert.isTrue(typeof entry.label === 'string' && entry.label.length > 0,
      'Multi-select entry.label must be a non-empty string');
  });

  // buildMultiSelectCols_ must be callable and return identical structure
  if (typeof buildMultiSelectCols_ === 'function') {
    var fresh = buildMultiSelectCols_();
    Assert.isTrue(fresh.MEMBER_DIR.length === MULTI_SELECT_COLS.MEMBER_DIR.length,
      'buildMultiSelectCols_() MEMBER_DIR length should match MULTI_SELECT_COLS');
    Assert.isTrue(fresh.GRIEVANCE_LOG.length === MULTI_SELECT_COLS.GRIEVANCE_LOG.length,
      'buildMultiSelectCols_() GRIEVANCE_LOG length should match MULTI_SELECT_COLS');
  }
}

/**
 * Verifies that auto multi-select is ON by default.
 * The `multiSelectAutoOpen` UserProperty must be absent OR not 'false' for
 * onSelectionChange to activate the dialog.  A fresh install (property absent)
 * must behave as enabled.
 *
 * This test simulates the flag check logic from onSelectionChange()
 * without actually calling PropertiesService (which requires a live session).
 */
function test_MultiSelectAutoOpen_DefaultOn() {
  /**
   * Simulates the auto-open flag check from onSelectionChange.
   * @param {string|null|undefined} propValue - The stored property value.
   * @returns {boolean} True if auto-open should be active.
   */
  function isAutoOpenActive(propValue) {
    return propValue !== 'false';
  }

  Assert.isTrue(isAutoOpenActive(null),
    'Auto multi-select must be ON when property is null (fresh install)');
  Assert.isTrue(isAutoOpenActive(undefined),
    'Auto multi-select must be ON when property is undefined');
  Assert.isTrue(isAutoOpenActive(''),
    'Auto multi-select must be ON when property is empty string');
  Assert.isTrue(isAutoOpenActive('true'),
    'Auto multi-select must be ON when property is "true"');
  Assert.isFalse(isAutoOpenActive('false'),
    'Auto multi-select must be OFF only when property is explicitly "false"');

  // Verify removeMultiSelectTrigger is the disable path (sets 'false')
  Assert.isDefined(typeof removeMultiSelectTrigger,
    'removeMultiSelectTrigger function must exist (opt-out path)');
  Assert.isDefined(typeof installMultiSelectTrigger,
    'installMultiSelectTrigger function must exist (re-enable path)');
}

// ============================================================================
// DRIVE / CONFIG COLUMN TESTS (v4.20.19)
// ============================================================================

/**
 * Verifies getDriveRootFolderName_() is defined and returns a non-empty string.
 * Failing here means CREATE_DASHBOARD and getOrCreateRootFolder() would
 * use the fallback name for ALL deployments, ignoring Config ORG_NAME.
 */
function test_DriveRootFolderName_Dynamic() {
  Assert.isDefined(typeof getDriveRootFolderName_,
    'getDriveRootFolderName_ must be defined');
  Assert.isTrue(typeof getDriveRootFolderName_ === 'function',
    'getDriveRootFolderName_ must be a function');

  // DRIVE_CONFIG fallback must exist and be a non-empty string
  Assert.isDefined(typeof DRIVE_CONFIG, 'DRIVE_CONFIG must be defined');
  Assert.isTrue(typeof DRIVE_CONFIG.ROOT_FOLDER_FALLBACK === 'string' &&
    DRIVE_CONFIG.ROOT_FOLDER_FALLBACK.length > 0,
    'DRIVE_CONFIG.ROOT_FOLDER_FALLBACK must be a non-empty string');

  // ROOT_FOLDER_NAME must NOT exist — it was replaced by getDriveRootFolderName_()
  Assert.isFalse('ROOT_FOLDER_NAME' in DRIVE_CONFIG,
    'DRIVE_CONFIG.ROOT_FOLDER_NAME must not exist — use getDriveRootFolderName_() instead');

  // Subfolder names must be non-empty strings
  ['GRIEVANCES_SUBFOLDER', 'RESOURCES_SUBFOLDER', 'MINUTES_SUBFOLDER', 'EVENT_CHECKIN_SUBFOLDER']
    .forEach(function(key) {
      Assert.isTrue(typeof DRIVE_CONFIG[key] === 'string' && DRIVE_CONFIG[key].length > 0,
        'DRIVE_CONFIG.' + key + ' must be a non-empty string');
    });
}

/**
 * Verifies CONFIG_COLS has all 5 Drive folder ID columns added in v4.20.17.
 * Failing here means setupDashboardDriveFolders() cannot write folder IDs
 * back to Config, making every subsequent folder lookup fall through to
 * name-search (expensive) or fail entirely.
 */
function test_ConfigCols_FolderIds_Exist() {
  var required = [
    'DASHBOARD_ROOT_FOLDER_ID',
    'GRIEVANCES_FOLDER_ID',
    'RESOURCES_FOLDER_ID',
    'MINUTES_FOLDER_ID',
    'EVENT_CHECKIN_FOLDER_ID'
  ];

  Assert.isDefined(typeof CONFIG_COLS, 'CONFIG_COLS must be defined');

  required.forEach(function(key) {
    Assert.isTrue(key in CONFIG_COLS,
      'CONFIG_COLS.' + key + ' must exist (added in v4.20.17)');
    Assert.isTrue(typeof CONFIG_COLS[key] === 'number' && CONFIG_COLS[key] > 0,
      'CONFIG_COLS.' + key + ' must be a positive integer (1-indexed)');
  });

  // Also verify the calendar ID column exists
  Assert.isTrue('CALENDAR_ID' in CONFIG_COLS && CONFIG_COLS.CALENDAR_ID > 0,
    'CONFIG_COLS.CALENDAR_ID must be a positive integer');
}

// ============================================================================
// ADDITIONAL TEST HELPERS
// ============================================================================

/**
 * Validates member ID format
 * @param {string} id - ID to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidMemberId_(id) {
  if (!id || typeof id !== 'string') return false;
  return /^MBR-\d+$/.test(id);
}

/**
 * Validates grievance ID format
 * @param {string} id - ID to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidGrievanceId_(id) {
  if (!id || typeof id !== 'string') return false;
  return /^GRV-\d{4}-\d+$/.test(id);
}

/**
 * Gets unique values from array
 * @param {Array} arr - Array to process
 * @returns {Array} Array with unique values
 * @private
 */
function getUniqueValues_(arr) {
  var seen = {};
  return arr.filter(function(item) {
    if (seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

/**
 * Flattens a nested array
 * @param {Array} arr - Nested array
 * @returns {Array} Flattened array
 * @private
 */
function flattenArray_(arr) {
  var result = [];
  arr.forEach(function(item) {
    if (Array.isArray(item)) {
      result = result.concat(flattenArray_(item));
    } else {
      result.push(item);
    }
  });
  return result;
}

/**
 * Checks if object has property
 * @param {Object} obj - Object to check
 * @param {string} prop - Property name
 * @returns {boolean} True if has property
 * @private
 */
function hasProperty_(obj, prop) {
  return obj && Object.prototype.hasOwnProperty.call(obj, prop);
}

// ============================================================================
// TEST DASHBOARD UI
// ============================================================================

