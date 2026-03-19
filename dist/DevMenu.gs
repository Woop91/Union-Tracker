// @dev-only
/**
 * ============================================================================
 * DevMenu.gs — Dev-Only Quick Deploy Menu
 * ============================================================================
 *
 * Consolidated "⚡ Dev Tools" menu providing one-click access to all
 * initialization, refresh, trigger installation, and setup functions.
 *
 * BUILD GATING: Excluded from production builds via PROD_EXCLUDE in build.js.
 * The onOpen() guard `typeof buildDevMenu === 'function'` ensures production
 * deployments (where this file is absent) never error.
 *
 * ARCHITECTURE:
 * - buildDevMenu()        → creates the menu (called from onOpen)
 * - devRunAll_*()          → master "run all" buttons (4 groups)
 * - runDevGroup_()         → shared helper for master buttons
 * - devWrap_*()            → individual item wrappers (try/catch + alert)
 *
 * @dev-only Excluded from production builds
 * @fileoverview Dev-only quick deploy menu
 */

// ============================================================================
// MENU BUILDER
// ============================================================================

/**
 * Builds the "⚡ Dev Tools" menu. Called from onOpen() with typeof guard.
 * Only available in dev builds (DevMenu.gs excluded from --prod).
 */
function buildDevMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('\u26A1 Dev Tools')

    // SECTION A — Master Group Actions
    .addItem('\u2699\uFE0F Initialize All', 'devRunAll_Initialize')
    .addItem('\uD83D\uDD04 Refresh All', 'devRunAll_Refresh')
    .addItem('\uD83D\uDCE6 Install All Triggers', 'devRunAll_Triggers')
    .addItem('\uD83D\uDCC5 Setup All Scheduled', 'devRunAll_Scheduled')
    .addSeparator()

    // SECTION B — Group 1: Initialize
    .addItem('Initialize All Trigger Functions', 'devWrap_InitAllTriggers')
    .addItem('Initialize All Sync Functions', 'devWrap_InitAllSync')
    .addItem('Initialize All Install Functions', 'devWrap_InitAllInstall')
    .addItem('Initialize All Refresh Functions', 'devWrap_InitAllRefresh')
    .addItem('Initialize Survey Engine', 'devWrap_InitSurveyEngine')
    .addItem('Initialize Poll Sheets', 'devWrap_InitPollSheets')
    .addItem('Workload: Initialize Sheets', 'devWrap_InitWorkloadSheets')
    .addSeparator()

    // SECTION B — Group 2: Refresh & Data Ops
    .addItem('Sync All Data Now', 'devWrap_SyncAllData')
    .addItem('Run Bulk Validation', 'devWrap_RunBulkValidation')
    .addItem('Force Global Refresh', 'devWrap_ForceGlobalRefresh')
    .addItem('Refresh All Formulas', 'devWrap_RefreshAllFormulas')
    .addItem('Refresh All Member Data', 'devWrap_RefreshAllMemberData')
    .addItem('Refresh View', 'devWrap_RefreshView')
    .addItem('Backfill Minutes', 'devWrap_BackfillMinutes')
    .addItem('Warm Up All Caches', 'devWrap_WarmUpCaches')
    .addSeparator()

    // SECTION B — Group 3: Trigger Installation
    .addItem('Install All Survey Triggers', 'devWrap_InstallAllSurveyTriggers')
    .addItem('Install Quarterly Trigger', 'devWrap_InstallQuarterlyTrigger')
    .addItem('Install Weekly Reminder Trigger', 'devWrap_InstallWeeklyReminderTrigger')
    .addItem('Install Community Poll Draw Trigger', 'devWrap_InstallCommunityPollTrigger')
    .addItem('Workload: Setup Reminders', 'devWrap_WorkloadSetupReminders')
    .addItem('Install onOpen Deferred Trigger', 'devWrap_InstallOnOpenDeferred')
    .addSeparator()

    // SECTION B — Group 4: Scheduled Tasks & Feature Setup
    .addItem('Enable Midnight Auto-Refresh', 'devWrap_EnableMidnightRefresh')
    .addItem('Enable 1 AM Dashboard Refresh', 'devWrap_Enable1AMRefresh')
    .addItem('Setup Weekly Backup', 'devWrap_SetupWeeklyBackup')
    .addItem('Setup Union Events Calendar', 'devWrap_SetupCalendar')
    .addItem('Setup / Repair Drive', 'devWrap_SetupDrive')
    .addItem('Create Meeting Check-In', 'devWrap_CreateMeetingCheckIn')

    .addToUi();
}

// ============================================================================
// SHARED HELPER — Master Group Runner
// ============================================================================

/**
 * Runs a group of dev actions sequentially, collecting results.
 * Each sub-function runs in its own try/catch so one failure doesn't block the rest.
 * Shows a single summary alert after all complete.
 *
 * @param {string} groupName - Display name for the group (e.g., "Initialize All")
 * @param {Array<{name: string, fn: Function}>} items - Array of {name, fn} pairs
 * @private
 */
function runDevGroup_(groupName, items) {
  var passed = [];
  var failed = [];
  for (var i = 0; i < items.length; i++) {
    try {
      items[i].fn();
      passed.push(items[i].name);
    } catch (e) {
      failed.push(items[i].name + ': ' + e.message);
    }
  }
  var msg = groupName + ' complete: ' + passed.length + '/' + items.length + ' succeeded.\n' +
    '\u2705 Passed: ' + (passed.length ? passed.join(', ') : 'none') + '\n' +
    '\u274C Failed: ' + (failed.length ? '\n  ' + failed.join('\n  ') : 'none');
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================================
// MASTER GROUP ACTIONS
// ============================================================================

/** Master: Run all Group 1 (Initialize) items */
function devRunAll_Initialize() {
  runDevGroup_('\u2699\uFE0F Initialize All', [
    { name: 'Initialize All Trigger Functions', fn: menuInstallSurveyTriggers },
    { name: 'Initialize All Sync Functions', fn: syncAllData },
    { name: 'Initialize All Install Functions', fn: initializeDashboard },
    { name: 'Initialize All Refresh Functions', fn: refreshAllFormulas },
    { name: 'Initialize Survey Engine', fn: initSurveyEngine },
    { name: 'Initialize Poll Sheets', fn: wqInitSheets },
    { name: 'Workload: Initialize Sheets', fn: initWorkloadTrackerSheets }
  ]);
}

/** Master: Run all Group 2 (Refresh & Data Ops) items */
function devRunAll_Refresh() {
  runDevGroup_('\uD83D\uDD04 Refresh All', [
    { name: 'Sync All Data Now', fn: syncAllData },
    { name: 'Run Bulk Validation', fn: runBulkValidation },
    { name: 'Force Global Refresh', fn: refreshAllFormulas },
    { name: 'Refresh All Formulas', fn: refreshAllFormulas },
    { name: 'Refresh All Member Data', fn: refreshMemberDirectoryFormulas },
    { name: 'Refresh View', fn: refreshMemberView },
    { name: 'Backfill Minutes', fn: BACKFILL_MINUTES_DRIVE_DOCS },
    { name: 'Warm Up All Caches', fn: warmUpCaches }
  ]);
}

/** Master: Run all Group 3 (Trigger Installation) items */
function devRunAll_Triggers() {
  runDevGroup_('\uD83D\uDCE6 Install All Triggers', [
    { name: 'Install All Survey Triggers', fn: menuInstallSurveyTriggers },
    { name: 'Install Quarterly Trigger', fn: setupQuarterlyTrigger },
    { name: 'Install Weekly Reminder Trigger', fn: setupWeeklyReminderTrigger },
    { name: 'Install Community Poll Draw Trigger', fn: setupCommunityPollTrigger },
    { name: 'Workload: Setup Reminders', fn: setupWorkloadReminderSystem },
    { name: 'Install onOpen Deferred Trigger', fn: setupOpenDeferredTrigger }
  ]);
}

/** Master: Run all Group 4 (Scheduled Tasks & Feature Setup) items */
function devRunAll_Scheduled() {
  runDevGroup_('\uD83D\uDCC5 Setup All Scheduled', [
    { name: 'Enable Midnight Auto-Refresh', fn: setupMidnightTrigger },
    { name: 'Enable 1 AM Dashboard Refresh', fn: createAutomationTriggers },
    { name: 'Setup Weekly Backup', fn: setupWeeklySnapshotTrigger },
    { name: 'Setup Union Events Calendar', fn: SETUP_CALENDAR },
    { name: 'Setup / Repair Drive', fn: SETUP_DRIVE_FOLDERS },
    { name: 'Create Meeting Check-In', fn: setupMeetingCheckInSheet }
  ]);
}

// ============================================================================
// INDIVIDUAL WRAPPER HELPER
// ============================================================================

/**
 * Wraps a single function call with try/catch and ui.alert feedback.
 * @param {string} label - Display name for the action
 * @param {Function} fn - The function to execute
 * @private
 */
function devWrap_(label, fn) {
  var ui = SpreadsheetApp.getUi();
  try {
    fn();
    ui.alert('\u2705 ' + label + ' completed successfully.');
  } catch (e) {
    ui.alert('\u274C ' + label + ' failed:\n\n' + e.message);
  }
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 1: Initialize
// ============================================================================

function devWrap_InitAllTriggers() {
  devWrap_('Initialize All Trigger Functions', menuInstallSurveyTriggers);
}

function devWrap_InitAllSync() {
  devWrap_('Initialize All Sync Functions', syncAllData);
}

function devWrap_InitAllInstall() {
  devWrap_('Initialize All Install Functions', initializeDashboard);
}

function devWrap_InitAllRefresh() {
  devWrap_('Initialize All Refresh Functions', refreshAllFormulas);
}

function devWrap_InitSurveyEngine() {
  devWrap_('Initialize Survey Engine', initSurveyEngine);
}

function devWrap_InitPollSheets() {
  devWrap_('Initialize Poll Sheets', wqInitSheets);
}

function devWrap_InitWorkloadSheets() {
  devWrap_('Workload: Initialize Sheets', initWorkloadTrackerSheets);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 2: Refresh & Data Ops
// ============================================================================

function devWrap_SyncAllData() {
  devWrap_('Sync All Data Now', syncAllData);
}

function devWrap_RunBulkValidation() {
  devWrap_('Run Bulk Validation', runBulkValidation);
}

function devWrap_ForceGlobalRefresh() {
  devWrap_('Force Global Refresh', refreshAllFormulas);
}

function devWrap_RefreshAllFormulas() {
  devWrap_('Refresh All Formulas', refreshAllFormulas);
}

function devWrap_RefreshAllMemberData() {
  devWrap_('Refresh All Member Data', refreshMemberDirectoryFormulas);
}

function devWrap_RefreshView() {
  devWrap_('Refresh View', refreshMemberView);
}

function devWrap_BackfillMinutes() {
  devWrap_('Backfill Minutes', BACKFILL_MINUTES_DRIVE_DOCS);
}

function devWrap_WarmUpCaches() {
  devWrap_('Warm Up All Caches', warmUpCaches);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 3: Trigger Installation
// ============================================================================

function devWrap_InstallAllSurveyTriggers() {
  devWrap_('Install All Survey Triggers', menuInstallSurveyTriggers);
}

function devWrap_InstallQuarterlyTrigger() {
  devWrap_('Install Quarterly Trigger', setupQuarterlyTrigger);
}

function devWrap_InstallWeeklyReminderTrigger() {
  devWrap_('Install Weekly Reminder Trigger', setupWeeklyReminderTrigger);
}

function devWrap_InstallCommunityPollTrigger() {
  devWrap_('Install Community Poll Draw Trigger', setupCommunityPollTrigger);
}

function devWrap_WorkloadSetupReminders() {
  devWrap_('Workload: Setup Reminders', setupWorkloadReminderSystem);
}

function devWrap_InstallOnOpenDeferred() {
  devWrap_('Install onOpen Deferred Trigger', setupOpenDeferredTrigger);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 4: Scheduled Tasks & Feature Setup
// ============================================================================

function devWrap_EnableMidnightRefresh() {
  devWrap_('Enable Midnight Auto-Refresh', setupMidnightTrigger);
}

function devWrap_Enable1AMRefresh() {
  devWrap_('Enable 1 AM Dashboard Refresh', createAutomationTriggers);
}

function devWrap_SetupWeeklyBackup() {
  devWrap_('Setup Weekly Backup', setupWeeklySnapshotTrigger);
}

function devWrap_SetupCalendar() {
  devWrap_('Setup Union Events Calendar', SETUP_CALENDAR);
}

function devWrap_SetupDrive() {
  devWrap_('Setup / Repair Drive', SETUP_DRIVE_FOLDERS);
}

function devWrap_CreateMeetingCheckIn() {
  devWrap_('Create Meeting Check-In', setupMeetingCheckInSheet);
}
