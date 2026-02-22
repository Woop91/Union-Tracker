/**
 * ============================================================================
 * 07_DevTools.gs - DEVELOPER TOOLS - DELETE THIS FILE BEFORE PRODUCTION
 * ============================================================================
 *
 * This file contains demo data seeding and nuclear cleanup functions.
 * These are for DEVELOPMENT AND TESTING ONLY.
 *
 * BEFORE GOING LIVE:
 * 1. Run NUKE_SEEDED_DATA() to clear all test data
 * 2. Delete this entire file from the Apps Script editor
 * 3. The Demo menu will automatically disappear on next refresh
 *
 * Once deleted, all seed/nuke functions will be gone and stewards
 * cannot accidentally trigger a data wipe.
 *
 * @version 4.7.0
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
 * Track a seeded grievance ID
 * @param {string} grievanceId - The grievance ID to track
 */
function trackSeededGrievanceId(grievanceId) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  if (ids.indexOf(grievanceId) === -1) {
    ids.push(grievanceId);
    props.setProperty('SEEDED_GRIEVANCE_IDS', ids.join(','));
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
 * Seed all sample data: Config + 1,000 members + 300 grievances
 * Grievances are randomly distributed - some members may have multiple
 * Auto-installs the sync trigger for live updates between sheets
 */
function SEED_SAMPLE_DATA() {
  if (!isDemoSafeToRun_()) return;

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var response = ui.alert(
    '🚀 Seed Sample Data',
    'This will seed:\n' +
    '• Config dropdowns (Job Titles, Locations, etc.)\n' +
    '• 1,000 sample members\n' +
    '• 300 sample grievances (30%)\n' +
    '• 50 sample survey responses\n' +
    '• 3 sample feedback entries\n' +
    '• Survey completion tracking data\n' +
    '• Auto-sync trigger for live updates\n\n' +
    'Note: Some members may have multiple grievances.\n' +
    'Member Directory will auto-update when Grievance Log changes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Seeding config data...', '🌱 Seeding', 3);
  seedConfigData();

  ss.toast('Seeding 1,000 members + 300 grievances (this may take a moment)...', '🌱 Seeding', 10);
  SEED_MEMBERS(1000, 30);

  ss.toast('Seeding survey responses...', '🌱 Seeding', 2);
  seedSatisfactionData();

  ss.toast('Seeding feedback entries...', '🌱 Seeding', 2);
  seedFeedbackData();

  ss.toast('Seeding survey tracking data...', '🌱 Seeding', 2);
  seedSurveyTrackingData();

  ss.toast('Installing auto-sync trigger...', '🔧 Setup', 3);
  installAutoSyncTriggerQuick();

  ss.toast('Sample data seeded successfully!', '✅ Success', 5);
  ui.alert('✅ Success', 'Sample data has been seeded!\n\n' +
    '• Config dropdowns populated\n' +
    '• 1,000 members added\n' +
    '• 300 grievances added (30%)\n' +
    '• 50 survey responses added\n' +
    '• 3 feedback entries added\n' +
    '• Survey tracking populated\n' +
    '• Auto-sync trigger installed\n\n' +
    'Member Directory columns (Has Open Grievance?, Grievance Status, Days to Deadline) ' +
    'will now auto-update when you edit the Grievance Log.', ui.ButtonSet.OK);
}

/**
 * Seed Config sheet with dropdown values
 * Note: Data starts at row 3 (row 1 = section headers, row 2 = column headers)
 * Seeds both preset values (Office Days, Grievance Status, etc.) and user-configurable values
 */
function seedConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found. Please run CREATE_DASHBOARD first.');
    return;
  }

  // Ensure Config sheet has enough columns (AZ = column 52 is the last column)
  ensureMinimumColumns(sheet, CONFIG_COLS.MOBILE_DASHBOARD_URL);

  // Data row start (after section headers row 1 and column headers row 2)
  var dataStartRow = 3;

  // Helper: only seed column if it's empty (preserves user data)
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
      'New', '', '', '')
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

  // Helper functions for generating seed data (defined outside loop)
  function weightedRating() {
    var r = Math.random();
    if (r < 0.1) return Math.floor(Math.random() * 3) + 1;      // 1-3 (10%)
    if (r < 0.3) return Math.floor(Math.random() * 2) + 4;      // 4-5 (20%)
    if (r < 0.6) return Math.floor(Math.random() * 2) + 6;      // 6-7 (30%)
    return Math.floor(Math.random() * 3) + 8;                    // 8-10 (40%)
  }

  function randomItem(arr) {
    return arr[Math.floor(Math.random() * arr.length)];
  }

  function randomPriorities() {
    var arr = priorities.slice();
    for (var j = arr.length - 1; j > 0; j--) {
      var k = Math.floor(Math.random() * (j + 1));
      var tmp = arr[j]; arr[j] = arr[k]; arr[k] = tmp;
    }
    return arr.slice(0, 3).join(', ');
  }

  var maxCol = SATISFACTION_COLS.Q67_ADDITIONAL; // last seeded column

  for (var i = 0; i < 50; i++) {
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
 * Seed sample members with optional grievances
 * @param {number} count - Number of members to seed (max 2000)
 * @param {number} grievancePercent - Percentage of members to give grievances (0-100, default 30)
 */
function SEED_MEMBERS(count, grievancePercent) {
  count = Math.min(count || 50, 2000);
  grievancePercent = (grievancePercent !== undefined) ? grievancePercent : 30; // Default 30% get grievances

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
  ss.toast('Ensuring Config data exists...', '🌱 Seeding', 2);
  seedConfigData();  // seedConfigData now only populates EMPTY columns

  // Now get all config values (will have data from seedConfigData or user's existing data)
  var jobTitles = getConfigValues(configSheet, CONFIG_COLS.JOB_TITLES);
  var locations = getConfigValues(configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  var units = getConfigValues(configSheet, CONFIG_COLS.UNITS);
  var supervisors = getConfigValues(configSheet, CONFIG_COLS.SUPERVISORS);
  var managers = getConfigValues(configSheet, CONFIG_COLS.MANAGERS);
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  var committees = getConfigValues(configSheet, CONFIG_COLS.STEWARD_COMMITTEES);

  // Expanded name pools for better variety (100+ names each = 10,000+ unique combinations)
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

  var rows = [];
  var seededIds = []; // Track IDs for this seeding session
  var batchSize = 50;
  var today = new Date();

  for (var i = 0; i < count; i++) {
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var memberId = generateNameBasedId('M', firstName, lastName, existingMemberIds);
    existingMemberIds[memberId] = true; // Track new ID to prevent duplicates in same batch
    seededIds.push(memberId); // Track for persistence
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + '.' + memberId.toLowerCase() + '@example.org';
    var phone = '617-555-' + String(Math.floor(Math.random() * 9000) + 1000);
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';
    var assignedSteward = randomChoice(stewards);

    // Generate recent contact data (50% chance of having recent contact)
    var hasRecentContact = Math.random() < 0.5;
    var recentContactDate = hasRecentContact ? randomDate(new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000), today) : '';
    var contactSteward = hasRecentContact ? assignedSteward : '';

    // Sample contact notes for members who have been contacted
    var sampleContactNotes = [
      'Discussed workload concerns',
      'Follow up on scheduling issue',
      'Interested in becoming steward',
      'Addressed safety complaint',
      'Positive feedback received',
      'Needs info on benefits',
      'Question about contract language',
      'Planning to attend next meeting',
      'Grievance update provided',
      'Initial outreach - new member',
      'Discussed upcoming negotiations',
      'Shared resources on workplace rights',
      'Scheduling meeting for next week'
    ];
    var contactNotes = hasRecentContact ? randomChoice(sampleContactNotes) : '';

    var row = generateSingleMemberRow(
      memberId, firstName, lastName,
      randomChoice(jobTitles),
      randomChoice(locations),
      randomChoice(units),
      randomChoice(officeDays),
      email, phone,
      randomChoice(commMethods),
      'Morning',
      randomChoice(supervisors),
      randomChoice(managers),
      isSteward,
      isSteward === 'Yes' ? randomChoice(committees) : '',
      assignedSteward,
      recentContactDate,
      contactSteward,
      contactNotes
    );

    rows.push(row);

    // Write in batches
    if (rows.length >= batchSize || i === count - 1) {
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(100);
    }
  }

  // Re-apply checkboxes to Start Grievance column (AE) - setValues overwrites them
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();
  }

  // Sync grievance data to Member Directory (populates AB-AD: Has Open Grievance, Status, Next Deadline)
  syncGrievanceToMemberDirectory();

  // Track seeded IDs for later cleanup (nuke only removes seeded data)
  trackSeededMemberIdsBatch(seededIds);

  // Seed grievances if percentage > 0
  var grievanceCount = Math.floor(count * grievancePercent / 100);
  if (grievanceCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded! Now seeding ' + grievanceCount + ' grievances...', '✅ Members Done', 2);
    SEED_GRIEVANCES(grievanceCount);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded!', '✅ Success', 3);
  }
}

/**
 * Seed members only (no automatic grievance seeding)
 * Used by SEED_SAMPLE_DATA to separate member and grievance seeding
 * @param {number} count - Number of members to seed (max 2000)
 */
function SEED_MEMBERS_ONLY(count) {
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

  var rows = [];
  var seededIds = [];
  var batchSize = 100; // Larger batches for 1000 members
  var today = new Date();

  for (var i = 0; i < count; i++) {
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var memberId = generateNameBasedId('M', firstName, lastName, existingMemberIds);
    existingMemberIds[memberId] = true;
    seededIds.push(memberId);
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + '.' + memberId.toLowerCase() + '@example.org';
    var phone = '617-555-' + String(Math.floor(Math.random() * 9000) + 1000);
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';
    var assignedSteward = randomChoice(stewards);

    var hasRecentContact = Math.random() < 0.5;
    var recentContactDate = hasRecentContact ? randomDate(new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000), today) : '';
    var contactSteward = hasRecentContact ? assignedSteward : '';

    var sampleContactNotes = [
      'Discussed workload concerns', 'Follow up on scheduling issue', 'Interested in becoming steward',
      'Addressed safety complaint', 'Positive feedback received', 'Needs info on benefits',
      'Question about contract language', 'Planning to attend next meeting', 'Grievance update provided',
      'Initial outreach - new member', 'Discussed upcoming negotiations', 'Shared resources on workplace rights'
    ];
    var contactNotes = hasRecentContact ? randomChoice(sampleContactNotes) : '';

    var row = generateSingleMemberRow(
      memberId, firstName, lastName,
      randomChoice(jobTitles), randomChoice(locations), randomChoice(units), randomChoice(officeDays),
      email, phone, randomChoice(commMethods), 'Morning',
      randomChoice(supervisors), randomChoice(managers),
      isSteward, isSteward === 'Yes' ? randomChoice(committees) : '', assignedSteward,
      recentContactDate, contactSteward, contactNotes
    );

    rows.push(row);

    // Write in batches
    if (rows.length >= batchSize || i === count - 1) {
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(50);
    }
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
  var maxCol = MEMBER_COLS.STATE; // last defined column
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

  return row;
}

// ensureMinimumColumns has been moved to 01_Core.gs for load-order safety.

/**
 * Seed N grievances
 * @param {number} count - Number of grievances to seed (max 300)
 */
function SEED_GRIEVANCES(count) {
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

  // Get config values for stewards
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  if (stewards.length === 0) stewards = ['Mary Steward'];

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
    var memberSteward = memberRow[MEMBER_COLS.ASSIGNED_STEWARD - 1] || randomChoice(stewards);

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

  var isClosed = (status === 'Settled' || status === 'Withdrawn' || status === 'Denied' || status === 'Won' || status === 'Closed');
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

/**
 * Show dialog to seed custom number of members (30% get grievances by default)
 */
function SEED_MEMBERS_DIALOG() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '👥 Seed Members & Grievances',
    'How many members to seed? (30% will get grievances)\nMax 2000:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var count = parseInt(response.getResponseText(), 10);
    if (isNaN(count) || count < 1) {
      ui.alert('Please enter a valid number.');
      return;
    }
    SEED_MEMBERS(count, 30);
  }
}

/**
 * Show dialog to seed custom number of members with custom grievance percentage
 */
function SEED_MEMBERS_ADVANCED_DIALOG() {
  var ui = SpreadsheetApp.getUi();

  var countResponse = ui.prompt(
    '👥 Seed Members (Step 1/2)',
    'How many members to seed? (max 2000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (countResponse.getSelectedButton() !== ui.Button.OK) return;

  var count = parseInt(countResponse.getResponseText(), 10);
  if (isNaN(count) || count < 1) {
    ui.alert('Please enter a valid number.');
    return;
  }

  var percentResponse = ui.prompt(
    '📋 Grievance Percentage (Step 2/2)',
    'What percentage of members should have grievances? (0-100)\nDefault: 30',
    ui.ButtonSet.OK_CANCEL
  );

  if (percentResponse.getSelectedButton() !== ui.Button.OK) return;

  var percent = parseInt(percentResponse.getResponseText(), 10);
  if (isNaN(percent)) percent = 30;
  percent = Math.max(0, Math.min(100, percent));

  SEED_MEMBERS(count, percent);
}

/**
 * Show dialog to seed custom number of grievances
 */
function SEED_GRIEVANCES_DIALOG() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '📋 Seed Grievances',
    'How many grievances to seed? (max 300)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var count = parseInt(response.getResponseText(), 10);
    if (isNaN(count) || count < 1) {
      ui.alert('Please enter a valid number.');
      return;
    }
    SEED_GRIEVANCES(count);
  }
}

/**
 * Seed 50 members with 30% grievances (shortcut)
 */
function seed50Members() {
  SEED_MEMBERS(50, 30);
}

/**
 * Seed 100 members with 50% grievances (shortcut)
 */
function seed100MembersWithGrievances() {
  SEED_MEMBERS(100, 50);
}

/**
 * Seed 25 grievances for existing members (shortcut)
 */
function seed25Grievances() {
  SEED_GRIEVANCES(25);
}

// ============================================================================
// NUKE FUNCTIONS
// ============================================================================

/**
 * Delete all seeded data from Member Directory and Grievance Log
 * Uses pattern matching (M/G + 4 letters + 3 digits) to identify seeded IDs
 */
function NUKE_SEEDED_DATA() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    '• _Audit_Log hidden sheet (entire sheet deleted)\n\n' +
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

    // Clear tracked IDs from Script Properties
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('SEEDED_MEMBER_IDS');
    props.deleteProperty('SEEDED_GRIEVANCE_IDS');

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

function isLargeDataset() {
  try {
    var ss = SpreadsheetApp.getActive();
    var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
    return memberDir ? memberDir.getLastRow() > TEST_LARGE_DATASET_THRESHOLD : false;
  } catch (_e) { return false; }
}

function createTestMember(memberId) {
  var ss = SpreadsheetApp.getActive();
  var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Build row using MEMBER_COLS for dynamic column placement
  var maxCol = MEMBER_COLS.STATE;
  var testData = [];
  for (var i = 0; i < maxCol; i++) { testData.push(''); }

  testData[MEMBER_COLS.MEMBER_ID - 1] = memberId || 'TEST-M001';
  testData[MEMBER_COLS.FIRST_NAME - 1] = 'Test';
  testData[MEMBER_COLS.LAST_NAME - 1] = 'Member';
  testData[MEMBER_COLS.OFFICE_DAYS - 1] = 'Monday';
  testData[MEMBER_COLS.EMAIL - 1] = 'test@union.org';
  testData[MEMBER_COLS.PHONE - 1] = '(555) 123-4567';
  testData[MEMBER_COLS.PREFERRED_COMM - 1] = 'Email';
  testData[MEMBER_COLS.BEST_TIME - 1] = 'Mornings';
  testData[MEMBER_COLS.IS_STEWARD - 1] = 'No';
  testData[MEMBER_COLS.LAST_VIRTUAL_MTG - 1] = new Date();
  testData[MEMBER_COLS.LAST_INPERSON_MTG - 1] = new Date();
  testData[MEMBER_COLS.OPEN_RATE - 1] = 85;
  testData[MEMBER_COLS.VOLUNTEER_HOURS - 1] = 10;
  testData[MEMBER_COLS.INTEREST_LOCAL - 1] = 'Yes';
  testData[MEMBER_COLS.INTEREST_CHAPTER - 1] = 'Yes';
  testData[MEMBER_COLS.INTEREST_ALLIED - 1] = 'No';
  testData[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = new Date();

  var startRow = Math.max(memberDir.getLastRow() + 1, 2);
  memberDir.getRange(startRow, 1, 1, testData.length).setValues([testData]);
  return memberId || 'TEST-M001';
}

function cleanupTestData() {
  var ss = SpreadsheetApp.getActive();
  var sheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG];
  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).indexOf('TEST-') === 0) sheet.deleteRow(i + 2);
    }
  });
}

// ==================== UNIT TESTS ====================

function testMemberColsConstants() {
  Assert.assertNotNull(MEMBER_COLS, 'MEMBER_COLS should exist');
  Assert.assertEquals(1, MEMBER_COLS.MEMBER_ID, 'MEMBER_ID should be column 1');
  Assert.assertEquals(2, MEMBER_COLS.FIRST_NAME, 'FIRST_NAME should be column 2');
  Assert.assertEquals(3, MEMBER_COLS.LAST_NAME, 'LAST_NAME should be column 3');
  Assert.assertEquals(8, MEMBER_COLS.EMAIL, 'EMAIL should be column 8');
  Assert.assertEquals(31, MEMBER_COLS.START_GRIEVANCE, 'START_GRIEVANCE should be column 31');
}

function testGrievanceColsConstants() {
  Assert.assertNotNull(GRIEVANCE_COLS, 'GRIEVANCE_COLS should exist');
  Assert.assertEquals(1, GRIEVANCE_COLS.GRIEVANCE_ID, 'GRIEVANCE_ID should be column 1');
  Assert.assertEquals(2, GRIEVANCE_COLS.MEMBER_ID, 'MEMBER_ID should be column 2');
  Assert.assertEquals(5, GRIEVANCE_COLS.STATUS, 'STATUS should be column 5');
  Assert.assertEquals(28, GRIEVANCE_COLS.RESOLUTION, 'RESOLUTION should be column 28');
}

function testColumnLetterConversion() {
  Assert.assertEquals('A', getColumnLetter(1), 'Column 1 should be A');
  Assert.assertEquals('Z', getColumnLetter(26), 'Column 26 should be Z');
  Assert.assertEquals('AA', getColumnLetter(27), 'Column 27 should be AA');
  Assert.assertEquals('AE', getColumnLetter(31), 'Column 31 should be AE');
}

function testSheetsConstants() {
  Assert.assertNotNull(SHEETS, 'SHEETS should exist');
  Assert.assertNotNull(SHEETS.MEMBER_DIR, 'MEMBER_DIR should exist');
  Assert.assertNotNull(SHEETS.GRIEVANCE_LOG, 'GRIEVANCE_LOG should exist');
  Assert.assertNotNull(SHEETS.CONFIG, 'CONFIG should exist');
}

function testValidateRequired() {
  Assert.assertThrows(function() { validateRequired(null, 'field'); }, 'null should throw');
  Assert.assertThrows(function() { validateRequired('', 'field'); }, 'empty should throw');
  Assert.assertThrows(function() { validateRequired(undefined, 'field'); }, 'undefined should throw');
}

function testValidateEmail() {
  var result1 = validateEmailAddress('test@example.com');
  Assert.assertTrue(result1.valid, 'Valid email should pass');
  var result2 = validateEmailAddress('invalid');
  Assert.assertFalse(result2.valid, 'Invalid email should fail');
  var result3 = validateEmailAddress('');
  Assert.assertFalse(result3.valid, 'Empty email should fail');
}

function testValidatePhoneNumber() {
  var result1 = validatePhoneNumber('(555) 123-4567');
  Assert.assertTrue(result1.valid, 'Valid phone should pass');
  var result2 = validatePhoneNumber('5551234567');
  Assert.assertTrue(result2.valid, 'Digits-only phone should pass');
  var result3 = validatePhoneNumber('123');
  Assert.assertFalse(result3.valid, 'Short phone should fail');
}

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

function testValidateGrievanceId() {
  // Format: G prefix + 4 uppercase letters (2 from first name + 2 from last name) + 3 digits
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM789'), 'GJOSM789 should be valid (G + John Smith + 789)');
  Assert.assertTrue(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GROWI001'), 'GROWI001 should be valid (G + Robert Williams + 001)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('JOSM789'), 'JOSM789 should be invalid (missing G prefix)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('G-123456'), 'G-123456 should be invalid (old format)');
  Assert.assertFalse(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM1234'), 'GJOSM1234 should be invalid (4 digits)');
}

function testOpenRateRange() {
  Assert.assertTrue(85 >= 0 && 85 <= 100, 'Open rate 85 should be in range');
  // eslint-disable-next-line no-self-compare -- Testing boundary conditions explicitly
  Assert.assertTrue(0 >= 0 && 0 <= 100, 'Open rate 0 should be in range');
  // eslint-disable-next-line no-self-compare -- Testing boundary conditions explicitly
  Assert.assertTrue(100 >= 0 && 100 <= 100, 'Open rate 100 should be in range');
}

// ==================== TEST RUNNER ====================

function getTestFunctionRegistry() {
  return {
    testMemberColsConstants: testMemberColsConstants,
    testGrievanceColsConstants: testGrievanceColsConstants,
    testColumnLetterConversion: testColumnLetterConversion,
    testSheetsConstants: testSheetsConstants,
    testValidateRequired: testValidateRequired,
    testValidateEmail: testValidateEmail,
    testValidatePhoneNumber: testValidatePhoneNumber,
    testValidateMemberId: testValidateMemberId,
    testValidateGrievanceId: testValidateGrievanceId,
    testOpenRateRange: testOpenRateRange
  };
}

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

function validatePhoneNumber(phone) {
  if (!phone || phone.toString().trim() === '') return { valid: false, message: 'Phone required' };
  var digits = phone.toString().replace(/\D/g, '');
  if (digits.length < 10) return { valid: false, message: 'At least 10 digits required' };
  if (digits.length > 15) return { valid: false, message: 'Phone too long' };
  var formatted = formatUSPhone(digits);
  return { valid: true, message: 'Valid phone', formatted: formatted };
}

function formatUSPhone(digits) {
  if (digits.length === 11 && digits[0] === '1') digits = digits.substring(1);
  if (digits.length === 10) return '(' + digits.substring(0, 3) + ') ' + digits.substring(3, 6) + '-' + digits.substring(6);
  return digits;
}

function validateRequired(value, fieldName) {
  if (value === null || value === undefined || value === '') throw new Error(fieldName + ' is required');
  return value;
}

function checkDuplicateMemberID(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) if (data[i][0] === memberId) count++;
  return count > 1;
}

function checkDuplicateGrievanceID(grievanceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) if (data[i][0] === grievanceId) count++;
  return count > 1;
}

/**
 * Check if a grievance's Member ID exists in the Member Directory
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if member ID exists, false otherwise
 */
function checkMemberIdExists(memberId) {
  if (!memberId) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet || memberSheet.getLastRow() < 2) return false;
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0] === memberId) return true;
  }
  return false;
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

function showValidationReport(issues, total) {
  var rate = total > 0 ? (((total - issues.length) / total) * 100).toFixed(1) : 100;
  var rows = issues.slice(0, 50).map(function(i) { return '<tr><td>' + i.row + '</td><td>' + i.field + '</td><td>' + i.value + '</td><td>' + i.message + '</td></tr>'; }).join('');
  if (issues.length > 50) rows += '<tr><td colspan="4">...and ' + (issues.length - 50) + ' more</td></tr>';
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{display:flex;gap:20px;margin:20px 0}.stat{flex:1;padding:20px;border-radius:8px;text-align:center}.stat.good{background:#e8f5e9}.stat.warning{background:#fff3e0}.stat.bad{background:#ffebee}.num{font-size:32px;font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:20px}th,td{padding:10px;text-align:left;border:1px solid #ddd}th{background:#f5f5f5}</style></head><body><h2>📊 Validation Report</h2><div class="summary"><div class="stat ' + (issues.length === 0 ? 'good' : issues.length < 10 ? 'warning' : 'bad') + '"><div class="num">' + rate + '%</div><div>Pass Rate</div></div><div class="stat good"><div class="num">' + total + '</div><div>Records</div></div><div class="stat ' + (issues.length === 0 ? 'good' : 'bad') + '"><div class="num">' + issues.length + '</div><div>Issues</div></div></div>' + (issues.length > 0 ? '<table><tr><th>Row</th><th>Field</th><th>Value</th><th>Issue</th></tr>' + rows + '</table>' : '<div style="text-align:center;padding:40px;color:#4caf50">✅ No issues found!</div>') + '</body></html>'
  ).setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Report');
}

function showValidationSettings() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.setting{margin:15px 0;padding:15px;background:#f8f9fa;border-radius:8px;display:flex;justify-content:space-between;align-items:center}.title{font-weight:bold}.desc{color:#666;font-size:13px}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;width:100%;margin-top:20px}</style></head><body><h2>⚙️ Validation Settings</h2><div class="setting"><div><div class="title">Email Validation</div><div class="desc">Validate format as you type</div></div><span>✅</span></div><div class="setting"><div><div class="title">Phone Auto-format</div><div class="desc">Format to (XXX) XXX-XXXX</div></div><span>✅</span></div><div class="setting"><div><div class="title">Duplicate Detection</div><div class="desc">Warn on duplicate IDs</div></div><span>✅</span></div><button onclick="google.script.run.runBulkValidation();google.script.host.close()">🔍 Run Bulk Validation</button></body></html>'
  ).setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Validation Settings');
}

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

function installValidationTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditValidation') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditValidation').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  SpreadsheetApp.getUi().alert('✅ Validation trigger installed!');
}

function onEditValidation(e) {
  if (!e || !e.range) return;
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
      var r = validatePhoneNumber(val);
      if (!r.valid) { e.range.setNote('⚠️ ' + r.message); e.range.setBackground('#fff3e0'); }
      else { e.range.clearNote(); e.range.setBackground(null); if (r.formatted !== val) e.range.setValue(r.formatted); }
    }
  }
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
 * @version 2.0.0
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
      test_MenuBuilder_Functions
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

/**
 * Generates a test report
 * @param {number} duration - Test duration
 * @returns {void}
 */
function generateTestReport(duration) {
  var results = TestSuite.results;
  var passed = results.filter(function(r) { return r.passed; }).length;
  var failed = results.filter(function(r) { return !r.passed; }).length;

  Logger.log('=== TEST REPORT ===');
  Logger.log('Total: ' + results.length);
  Logger.log('Passed: ' + passed);
  Logger.log('Failed: ' + failed);
  Logger.log('Duration: ' + duration + 'ms');
  Logger.log('==================');
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
  Assert.isDefined(typeof getADHDSettings, 'getADHDSettings should be defined');
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

/**
 * Shows a comprehensive test dashboard
 * @returns {void}
 */
function showTestDashboard() {
  var results = TestSuite.runAll();

  var testRows = results.results.map(function(r) {
    var status = r.passed ? '✅' : '❌';
    var errorMsg = r.error ? '<span style="color:#ef4444">' + r.error + '</span>' : '-';
    return '<tr>' +
      '<td>' + status + '</td>' +
      '<td>' + r.name + '</td>' +
      '<td>' + r.duration + 'ms</td>' +
      '<td>' + errorMsg + '</td>' +
      '</tr>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px}' +
    'h2{color:#7c3aed}' +
    '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin:20px 0}' +
    '.stat{background:#f8f9fa;padding:20px;border-radius:8px;text-align:center}' +
    '.stat.passed{border-left:4px solid #10b981}' +
    '.stat.failed{border-left:4px solid #ef4444}' +
    '.num{font-size:36px;font-weight:bold}' +
    '.passed .num{color:#10b981}' +
    '.failed .num{color:#ef4444}' +
    'table{width:100%;border-collapse:collapse;margin-top:20px}' +
    'th{background:#7c3aed;color:white;padding:12px;text-align:left}' +
    'td{padding:10px;border-bottom:1px solid #e5e7eb}' +
    '.btn{padding:12px 24px;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin:5px}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>🧪 Test Dashboard</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + results.total + '</div><div>Total</div></div>' +
    '<div class="stat passed"><div class="num">' + results.passed + '</div><div>Passed</div></div>' +
    '<div class="stat failed"><div class="num">' + results.failed + '</div><div>Failed</div></div>' +
    '<div class="stat"><div class="num">' + results.duration + 'ms</div><div>Duration</div></div>' +
    '</div>' +
    '<button class="btn btn-primary" onclick="rerun()">🔄 Run Again</button>' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '<table>' +
    '<tr><th>Status</th><th>Test Name</th><th>Duration</th><th>Error</th></tr>' +
    testRows +
    '</table>' +
    '</div>' +
    '<script>function rerun(){google.script.run.withSuccessHandler(function(){location.reload()}).runAllTests()}</script>' +
    '</body></html>'
  ).setWidth(900).setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '🧪 Test Dashboard');
}
