// @dev-only
/**
 * ============================================================================
 * DevMenu.gs — Dev-Only Quick Deploy Menu
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Development-only quick deploy menu providing one-click access to all
 *   initialization, refresh, trigger installation, and setup functions.
 *   Creates the "Dev Tools" top-level menu with 4 master group actions
 *   (Initialize All, Refresh All, Install All Triggers, Setup All Scheduled)
 *   and individual wrappers for every init/setup function across all modules.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Excluded from production builds via PROD_EXCLUDE in build.js. The
 *   onOpen() guard `typeof buildDevMenu === 'function'` ensures production
 *   deployments never error when this file is absent. devWrap_*() wrappers
 *   add try/catch + alert() around each function call so developers see
 *   immediate feedback. runDevGroup_() runs multiple functions in sequence
 *   with a progress toast between each.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Developers lose the quick deploy menu and must call init functions
 *   individually from the script editor. No production impact — this file
 *   is excluded from production builds. If the PROD_EXCLUDE filter fails
 *   and this file reaches production, the dev menu appears for all users
 *   (cosmetic issue but confusing).
 *
 * DEPENDENCIES:
 *   Depends on every module's init/setup functions (initSurveyEngine,
 *   initWeeklyQuestionSheets, initWorkloadSheets, etc.). Used only in
 *   development builds.
 *
 * @dev-only Excluded from production builds
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

/** Dev menu: initializes all trigger functions. */
function devWrap_InitAllTriggers() {
  devWrap_('Initialize All Trigger Functions', menuInstallSurveyTriggers);
}

/** Dev menu: initializes all sync functions. */
function devWrap_InitAllSync() {
  devWrap_('Initialize All Sync Functions', syncAllData);
}

/** Dev menu: runs full dashboard initialization. */
function devWrap_InitAllInstall() {
  devWrap_('Initialize All Install Functions', initializeDashboard);
}

/** Dev menu: refreshes all formulas. */
function devWrap_InitAllRefresh() {
  devWrap_('Initialize All Refresh Functions', refreshAllFormulas);
}

/** Dev menu: initializes the survey engine. */
function devWrap_InitSurveyEngine() {
  devWrap_('Initialize Survey Engine', initSurveyEngine);
}

/** Dev menu: initializes weekly question sheets. */
function devWrap_InitPollSheets() {
  devWrap_('Initialize Poll Sheets', wqInitSheets);
}

/** Dev menu: initializes workload tracker sheets. */
function devWrap_InitWorkloadSheets() {
  devWrap_('Workload: Initialize Sheets', initWorkloadTrackerSheets);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 2: Refresh & Data Ops
// ============================================================================

/** Dev menu: synchronizes all data now. */
function devWrap_SyncAllData() {
  devWrap_('Sync All Data Now', syncAllData);
}

/** Dev menu: runs bulk data validation. */
function devWrap_RunBulkValidation() {
  devWrap_('Run Bulk Validation', runBulkValidation);
}

/** Dev menu: forces a global formula refresh. */
function devWrap_ForceGlobalRefresh() {
  devWrap_('Force Global Refresh', refreshAllFormulas);
}

/** Dev menu: refreshes all formulas. */
function devWrap_RefreshAllFormulas() {
  devWrap_('Refresh All Formulas', refreshAllFormulas);
}

/** Dev menu: refreshes all member directory data. */
function devWrap_RefreshAllMemberData() {
  devWrap_('Refresh All Member Data', refreshMemberDirectoryFormulas);
}

/** Dev menu: refreshes the current member view. */
function devWrap_RefreshView() {
  devWrap_('Refresh View', refreshMemberView);
}

/** Dev menu: backfills meeting minutes Drive docs. */
function devWrap_BackfillMinutes() {
  devWrap_('Backfill Minutes', BACKFILL_MINUTES_DRIVE_DOCS);
}

/** Dev menu: warms up all caches. */
function devWrap_WarmUpCaches() {
  devWrap_('Warm Up All Caches', warmUpCaches);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 3: Trigger Installation
// ============================================================================

/** Dev menu: installs all survey triggers. */
function devWrap_InstallAllSurveyTriggers() {
  devWrap_('Install All Survey Triggers', menuInstallSurveyTriggers);
}

/** Dev menu: installs the quarterly survey trigger. */
function devWrap_InstallQuarterlyTrigger() {
  devWrap_('Install Quarterly Trigger', setupQuarterlyTrigger);
}

/** Dev menu: installs the weekly reminder trigger. */
function devWrap_InstallWeeklyReminderTrigger() {
  devWrap_('Install Weekly Reminder Trigger', setupWeeklyReminderTrigger);
}

/** Dev menu: installs the community poll draw trigger. */
function devWrap_InstallCommunityPollTrigger() {
  devWrap_('Install Community Poll Draw Trigger', setupCommunityPollTrigger);
}

/** Dev menu: sets up workload reminder system. */
function devWrap_WorkloadSetupReminders() {
  devWrap_('Workload: Setup Reminders', setupWorkloadReminderSystem);
}

/** Dev menu: installs the onOpen deferred trigger. */
function devWrap_InstallOnOpenDeferred() {
  devWrap_('Install onOpen Deferred Trigger', setupOpenDeferredTrigger);
}

// ============================================================================
// INDIVIDUAL WRAPPERS — Group 4: Scheduled Tasks & Feature Setup
// ============================================================================

/** Dev menu: enables the midnight auto-refresh trigger. */
function devWrap_EnableMidnightRefresh() {
  devWrap_('Enable Midnight Auto-Refresh', setupMidnightTrigger);
}

/** Dev menu: enables the 1 AM dashboard refresh trigger. */
function devWrap_Enable1AMRefresh() {
  devWrap_('Enable 1 AM Dashboard Refresh', createAutomationTriggers);
}

/** Dev menu: sets up the weekly backup trigger. */
function devWrap_SetupWeeklyBackup() {
  devWrap_('Setup Weekly Backup', setupWeeklySnapshotTrigger);
}

/** Dev menu: sets up the union events calendar. */
function devWrap_SetupCalendar() {
  devWrap_('Setup Union Events Calendar', SETUP_CALENDAR);
}

/** Dev menu: sets up or repairs Drive folder structure. */
function devWrap_SetupDrive() {
  devWrap_('Setup / Repair Drive', SETUP_DRIVE_FOLDERS);
}

/** Dev menu: creates the meeting check-in sheet. */
function devWrap_CreateMeetingCheckIn() {
  devWrap_('Create Meeting Check-In', setupMeetingCheckInSheet);
}
