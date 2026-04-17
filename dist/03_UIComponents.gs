/**
 * ============================================================================
 * 03_UIComponents.gs - Menu System and Navigation
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Menu system creation — the main UI entry point for all spreadsheet features.
 *   createDashboardMenu() builds 3 top-level menus (Union Hub, Tools, Admin)
 *   with 40+ items covering search, members, grievances, analytics, calendar,
 *   drive, notifications, admin, and more. Uses emoji prefixes for visual
 *   scanning in the menu bar.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Called by onOpen() in 10_Main.gs. Centralizes all menu items in one place
 *   so features are discoverable without scattering menu logic across modules.
 *   40+ items is acceptable for the power-user audience (union stewards who
 *   need fast access to many features). Refactored out of 04_UIService.gs to
 *   keep menu definition separate from presentation/dialog logic.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Users see NO menus when they open the spreadsheet. ALL spreadsheet-based
 *   features become inaccessible — users must manually call functions from the
 *   Apps Script editor, which is not viable for non-technical stewards. This
 *   is a critical file for usability.
 *
 * DEPENDENCIES:
 *   - Depends on: 01_Core.gs (SHEETS, getLocalNumberFromConfig_)
 *   - Used by:    10_Main.gs (onOpen calls createDashboardMenu)
 *
 * @fileoverview Menu creation and navigation functions
 * @requires 01_Core.gs
 */

// ============================================================================
// DIALOG UTILITIES
// ============================================================================

/**
 * Shows a modal dialog with standard boilerplate.
 * Replaces repeated HtmlService.createHtmlOutput().setWidth().setHeight() pattern.
 * @param {string} htmlContent - The HTML string to display
 * @param {string} title - Dialog title
 * @param {number} [width=600] - Dialog width in pixels
 * @param {number} [height=500] - Dialog height in pixels
 */
function showDialog_(htmlContent, title, width, height) {
  var output = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(width || 600)
    .setHeight(height || 500);
  SpreadsheetApp.getUi().showModalDialog(output, title);
}

// ============================================================================
// MENU CREATION
// ============================================================================

/**
 * Creates the custom menu when the spreadsheet opens
 * Consolidated menu structure (v4.4.0):
 * - 📊 Union Hub: Primary operations (dashboards, search, members, cases)
 * - 🔧 Tools: Supporting features (calendar, drive, notifications, etc.)
 * - 🛠️ Admin: System administration
 * @returns {void}
 */
function createDashboardMenu() {
  var ui = SpreadsheetApp.getUi();
  var localNumber = getLocalNumberFromConfig_();

  // ============================================================================
  // MENU 1: Union Hub — Daily Operations
  // ============================================================================
  var ver = (typeof VERSION_INFO !== 'undefined' && VERSION_INFO.version) ? ' v' + VERSION_INFO.version : '';
  ui.createMenu('📊 ' + localNumber + ' Union Hub' + ver)
    .addSubMenu(ui.createMenu('🔍 Search')
      .addItem('🔍 Desktop Search', 'showDesktopSearch')
      .addItem('⚡ Quick Search', 'showQuickSearch')
      .addItem('🔎 Advanced Search', 'showAdvancedSearch'))

    .addSubMenu(ui.createMenu('👥 Members')
      .addItem('➕ Add New Member', 'showNewMemberDialog')
      .addItem('🔍 Find Member', 'showSearchDialog')
      .addItem('🛡️ Steward Directory', 'showStewardDirectory')
      .addSeparator()
      .addItem('📥 Import Members', 'showImportMembersDialog')
      .addItem('📤 Export Directory', 'showExportMembersDialog')
      .addSeparator()
      .addItem('🆔 Generate Missing IDs', 'generateMissingMemberIDs')
      .addItem('🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs')
      .addSeparator()
      .addItem('⬆️ Promote to Steward', 'promoteSelectedMemberToSteward')
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward')
      .addItem('🔄 Sync Steward Status', 'syncStewardStatus'))

    .addSubMenu(ui.createMenu('📋 Cases & Grievances')
      .addItem('➕ New Case/Grievance', 'showNewGrievanceDialog')
      .addItem('✏️ Edit Selected', 'showEditGrievanceDialog')
      .addItem('✅ View Checklist', 'showChecklistDialog')
      .addSeparator()
      .addItem('🔄 Bulk Update Status', 'showBulkStatusUpdate')
      .addItem('📄 Create Signature PDF', 'createPDFForSelectedGrievance')
      .addItem('📄 Create Template PDF', 'createTemplatePDFForSelectedGrievance')
      .addSeparator()
      .addSubMenu(ui.createMenu('📋 Bulk Actions')
        .addItem('📋 Select All Open Cases', 'selectAllOpenCases')
        .addItem('🔄 Clear Selection', 'clearAllSelections')
        .addSeparator()
        .addItem('⚡ Bulk Actions on Selected...', 'showBulkActionsDialog'))
      .addSeparator()
      .addItem('🚦 Apply Traffic Lights', 'applyTrafficLightIndicators')
      .addItem('🔄 Clear Traffic Lights', 'clearTrafficLightIndicators'))

    .addSubMenu(ui.createMenu('📈 Analytics & Reports')
      .addItem('📊 Grievance Trends', 'showGrievanceTrends')
      .addItem('🏥 Unit Health Report', 'showUnitHealthReport')
      .addItem('📚 Search Precedents', 'showSearchPrecedents'))

    .addSeparator()
    .addItem('🤝 Log Volunteer Hours', 'showLogVolunteerHoursModal')
    .addItem('👥 Add Non-Member Contact', 'showAddContactModal')
    .addItem('✅ Case Checklist Progress', 'showCaseProgressModal')
    .addSeparator()
    .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
    .addItem('📖 Help & Documentation', 'showHelpDialog')
    .addToUi();

  // ============================================================================
  // MENU 2: Tools — Features Used While Working
  // ============================================================================
  ui.createMenu('🔧 Tools')
    .addSubMenu(ui.createMenu('📧 Email & Notifications')
      .addItem('📧 Send Portal Email', 'sendPortalEmailToSelectedMember')
      .addSeparator()
      .addItem('⚙️ Notification Settings', 'showNotificationSettings')
      .addItem('🧪 Test Notifications', 'testDeadlineNotifications'))

    .addSubMenu(ui.createMenu('📅 Calendar & Meetings')
      .addItem('📝 Setup New Meeting', 'showSetupMeetingDialog')
      .addItem('✅ Open Meeting Check-In', 'showMeetingCheckInDialog')
      .addItem('📱 QR Code Check-In', 'showMeetingQRCodeDialog')
      .addItem('📅 Add New Event', 'showAddEventModal')
      .addItem('📝 Add Meeting Minutes', 'showAddMinutesModal')
      .addItem('📋 Take Attendance', 'showTakeAttendanceModal')
      .addSeparator()
      .addItem('🔗 Sync Deadlines', 'syncDeadlinesToCalendar')
      .addItem('👁️ View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar')
      .addSeparator()
      .addItem('📅 Open Google Calendar', 'openGoogleCalendar'))

    .addSubMenu(ui.createMenu('📁 Google Drive')
      .addItem('📁 Setup Folder for Grievance', 'setupFolderForSelectedGrievance')
      .addItem('📁 View Grievance Files', 'showGrievanceFiles')
      .addItem('📁 Batch Create Folders', 'batchCreateAllMissingFolders')
      .addSeparator()
      .addItem('📁 Open Google Drive', 'openGoogleDrive'))

    .addSubMenu(ui.createMenu('📊 Workload Tracker')
      .addItem('🔄 Refresh Ledger', 'refreshWorkloadLedger')
      .addItem('💾 Create Backup', 'createWorkloadBackup')
      .addItem('🗄️ Archive Old Data', 'wtArchiveOldData')
      .addItem('🧹 Clean Vault Dedup', 'wtCleanVault')
      .addSeparator()
      .addItem('🩺 Health Status', 'showWorkloadHealthStatus'))

    .addSubMenu(ui.createMenu('📋 Surveys & Polls')
      .addItem('📂 Open New Survey Period', 'menuOpenNewSurveyPeriod')
      .addItem('📊 View Current Period Status', 'menuShowSurveyPeriodStatus')
      .addItem('📊 View Survey Responses', 'showSurveyResponseViewer')
      .addItem('📋 Edit Survey Question', 'showQuestionEditorModal')
      .addItem('🔔 Send Survey Reminders Now', 'sendSurveyCompletionReminders')
      .addSeparator()
      .addItem('⚡ Create Weekly Question', 'showCreateWeeklyQuestionModal')
      .addItem('🏢 Toggle Return-to-Office Questions', 'menuToggleRTOSection')
      .addSeparator()
      .addItem('🔄 Draw Community Poll Now', 'autoSelectCommunityPoll'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🎛️ Sheet Tools')
      .addItem('🎛️ Visual Control Panel', 'showVisualControlPanel')
      .addItem('📱 Toggle Pocket View', 'toggleMobileView')
      .addItem('📱 Pocket View (Active Sheet)', 'navToMobile')
      .addItem('🖥️ Restore All Columns', 'showAllColumns')
      .addSeparator()
      .addItem('📝 Multi-Select Editor', 'openCellMultiSelectEditor')
      .addItem('⚡ Enable Multi-Select Auto-Open', 'installMultiSelectTrigger')
      .addItem('🚫 Disable Multi-Select Auto-Open', 'removeMultiSelectTrigger')
      .addSeparator()
      .addItem('↩️ Undo Last Action', 'undoLastAction')
      .addItem('↪️ Redo Last Action', 'redoLastAction')
      .addItem('📋 Export Undo History', 'exportUndoHistoryToSheet')
      .addSeparator()
      .addItem('🔧 Repair Dynamic Formulas', 'repairDynamicFormulas'))

    .addItem('📑 Tab Quick Actions', 'showCurrentTabModal')
    .addSubMenu(ui.createMenu('📑 Tab Modals')
      .addItem('⚙️ Config', 'showTabModalConfig')
      .addItem('👥 Member Directory', 'showTabModalMemberDirectory')
      .addItem('⚖️ Grievance Log', 'showTabModalGrievanceLog')
      .addItem('✅ Case Checklist', 'showTabModalCaseChecklist')
      .addItem('🤝 Volunteer Hours', 'showTabModalVolunteerHours')
      .addItem('📅 Meeting Attendance', 'showTabModalMeetingAttendance')
      .addItem('📝 Meeting Check-In', 'showTabModalMeetingCheckIn')
      .addItem('📚 Resources', 'showTabModalResources')
      .addSeparator()
      .addItem('🔄 Reset "Don\'t Show Again"', 'resetTabModalDismissals'))

    .addSubMenu(ui.createMenu('🎨 Themes')
      .addItem('🎨 Apply Theme', 'APPLY_SYSTEM_THEME')
      .addItem('🏛️ Apply Union Brand Theme (All Tabs)', 'applyUnionThemeToAllTabs')
      .addItem('🎨 Theme Presets', 'showThemePresetPicker')
      .addItem('🔄 Reset to Default', 'resetToDefaultTheme')
      .addItem('✨ Refresh All Visuals', 'refreshAllVisuals'))

    .addSeparator()
    .addItem('📱 Get Mobile App URL', 'showWebAppUrl')
    .addToUi();

  // ============================================================================
  // MENU 3: Admin — Configuration & System Management
  // ============================================================================
  var adminMenu = ui.createMenu('🛠️ Admin')
    .addItem('🏗️ Welcome / Setup Wizard', 'showWelcomeWizardModal')
    .addItem('❓ Searchable Help Guide', 'showSearchableHelpModal')
    .addSeparator()
    .addItem('🩺 System Diagnostics', 'showDiagnosticsDialog')
    .addItem('🔍 Modal Diagnostics', 'showModalDiagnostics')
    .addItem('🔧 Repair Dashboard', 'showRepairDialog')
    .addItem('🔄 Update All Sheets', 'UPDATE_ALL_SHEETS')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addItem('🛠️ Admin Settings', 'showAdminSettingsSidebar')
    .addSeparator()

    .addSubMenu(ui.createMenu('🔄 Data Sync')
      .addItem('🔄 Sync All Data Now', 'syncAllData')
      .addItem('🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory')
      .addItem('🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog')
      .addItem('🤝 Sync Volunteer Hours → Members', 'syncVolunteerHoursToMemberDirectory')
      .addItem('📅 Sync Meeting Attendance → Members', 'syncMeetingAttendanceToMemberDirectory')
      .addItem('📊 Sync All Engagement Data', 'syncEngagementToMemberDirectory')
      .addSeparator()
      .addItem('📧 Sync CC Engagement → Members', 'syncConstantContactEngagement')
      .addItem('📧 CC Setup: API Credentials', 'showConstantContactSetup')
      .addItem('📧 CC Authorize Account', 'authorizeConstantContact')
      .addItem('📧 CC Connection Status', 'showConstantContactStatus')
      .addItem('🔌 CC Disconnect', 'disconnectConstantContact')
      .addSeparator()
      .addItem('⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger')
      .addItem('🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger'));

  // Validation submenu — functions are in 07_DevTools.gs (dev-only)
  if (!isProductionMode()) {
    adminMenu.addSubMenu(ui.createMenu('✅ Validation')
      .addItem('🔍 Run Bulk Validation', 'runBulkValidation')
      .addItem('⚙️ Validation Settings', 'showValidationSettings')
      .addItem('🧹 Clear Indicators', 'clearValidationIndicators')
      .addItem('⚡ Install Validation Trigger', 'installValidationTrigger'));
  }

  adminMenu
    .addSubMenu(ui.createMenu('⚙️ Automation')
      .addItem('🌙 Enable Midnight Auto-Refresh', 'setupMidnightTrigger')
      .addItem('❌ Disable Midnight Auto-Refresh', 'removeMidnightTrigger')
      .addItem('🔔 Enable 1AM Dashboard Refresh', 'createAutomationTriggers')
      .addItem('📑 Email Weekly PDF Snapshot', 'emailExecutivePDF'))

    .addSubMenu(ui.createMenu('🗄️ Cache & Performance')
      .addItem('🗄️ Cache Status', 'showCacheStatusDashboard')
      .addItem('🔥 Warm Up Caches', 'warmUpCaches')
      .addItem('🗑️ Clear All Caches', 'invalidateAllCaches')
      .addSeparator()
      .addItem('⏱️ Enable 1m Web App Keep-Warm', 'installWebAppCacheWarmupTrigger')
      .addItem('🚫 Disable Web App Auto-Warmup', 'removeWebAppCacheWarmupTrigger'))

    .addSubMenu(ui.createMenu('🛡️ Security & Backup')
      .addItem('📸 Create Manual Snapshot', 'createWeeklySnapshot')
      .addItem('📅 Setup Weekly Backup', 'setupWeeklySnapshotTrigger')
      .addItem('📜 View Audit Log', 'navigateToAuditLog')
      .addItem('📊 v4.0 Status Report', 'showV4StatusReport')
      .addSeparator()
      .addItem('🔐 Dashboard Auth Status', 'showDashboardAuthStatus')
      .addItem('✅ Enable Dashboard Member Auth', 'enableDashboardMemberAuth')
      .addItem('❌ Disable Dashboard Member Auth', 'disableDashboardMemberAuth'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🚀 Quick Setup (All Init/Sync)')
      .addItem('🚀 Initialize Dashboard', 'initializeDashboard')
      .addSeparator()
      .addItem('--- INITIALIZE ALL ---', 'initializeDashboard')
      .addItem('🚀 Initialize Survey Engine', 'initSurveyEngine')
      .addItem('🏗️ Initialize Poll Sheets', 'wqInitSheets')
      .addItem('⚙️ Workload: Initialize Sheets', 'initWorkloadTrackerSheets')
      .addItem('📝 Create Meeting Check-In Sheet', 'setupMeetingCheckInSheet')
      .addItem('🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets')
      .addSeparator()
      .addItem('--- INSTALL ALL TRIGGERS ---', 'setupOpenDeferredTrigger')
      .addItem('✅ Install ALL Survey Triggers', 'menuInstallSurveyTriggers')
      .addItem('⏱️ Install Quarterly Trigger', 'setupQuarterlyTrigger')
      .addItem('⏱️ Install Weekly Reminder Trigger', 'setupWeeklyReminderTrigger')
      .addItem('⏱️ Install Community Poll Draw Trigger', 'setupCommunityPollTrigger')
      .addItem('🔔 Workload: Setup Reminders', 'setupWorkloadReminderSystem')
      .addItem('🔓 Install onOpen Deferred Trigger', 'setupOpenDeferredTrigger')
      .addSeparator()
      .addItem('--- SYNC & REFRESH ---', 'syncAllData')
      .addItem('🔄 Sync All Data Now', 'syncAllData')
      .addItem('🔥 Force Global Refresh', 'refreshAllFormulas')
      .addItem('🌙 Enable Midnight Auto-Refresh', 'setupMidnightTrigger')
      .addItem('🔔 Enable 1AM Dashboard Refresh', 'createAutomationTriggers')
      .addItem('🔥 Warm Up All Caches', 'warmUpCaches')
      .addSeparator()
      .addItem('--- SETUP SERVICES ---', 'SETUP_CALENDAR')
      .addItem('🏗️ Setup Union Events Calendar', 'SETUP_CALENDAR')
      .addItem('🏗️ Setup / Repair Drive', 'SETUP_DRIVE_FOLDERS')
      .addItem('📅 Setup Weekly Backup', 'setupWeeklySnapshotTrigger')
      .addSeparator()
      .addItem('--- UPDATE & MAINTENANCE ---', 'UPDATE_ALL_SHEETS')
      .addItem('🔄 Update All Sheets', 'UPDATE_ALL_SHEETS')
      .addItem('🔄 Refresh All Formulas', 'refreshAllFormulas')
      .addItem('🔄 Refresh All Member Data', 'refreshMemberDirectoryFormulas')
      .addItem('🔄 Refresh View', 'refreshMemberView')
      .addItem('📄 Backfill Minutes', 'BACKFILL_MINUTES_DRIVE_DOCS'))

    .addSubMenu(ui.createMenu('🏗️ Initial Setup')
      .addItem('🚀 Initialize Dashboard', 'initializeDashboard')
      .addSeparator()
      .addItem('🏗️ Setup Union Events Calendar', 'SETUP_CALENDAR')
      .addItem('🏗️ Setup / Repair Drive Folder Structure', 'SETUP_DRIVE_FOLDERS')
      .addSeparator()
      .addItem('🚀 Initialize Survey Engine', 'initSurveyEngine')
      .addItem('🏗️ Initialize Poll Sheets', 'wqInitSheets')
      .addItem('⚙️ Workload: Initialize Sheets', 'initWorkloadTrackerSheets')
      .addItem('📝 Create Meeting Check-In Sheet', 'setupMeetingCheckInSheet')
      .addSeparator()
      .addItem('🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets')
      .addItem('🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets')
      .addItem('🔍 Verify Hidden Sheets', 'verifyHiddenSheets')
      .addItem('🔒 Enforce Hidden Sheets (Mobile Fix)', 'enforceHiddenSheets')
      .addItem('⚙️ Setup Data Validations', 'setupDataValidations')
      .addItem('👥 Setup Member Leader Role', 'setupMemberLeaderRole')
      .addItem('🎨 Setup Comfort View Defaults', 'setupComfortViewDefaults'))

    .addSubMenu(ui.createMenu('⏱️ Triggers')
      .addItem('✅ Install ALL Survey Triggers', 'menuInstallSurveyTriggers')
      .addItem('⏱️ Install Quarterly Trigger', 'setupQuarterlyTrigger')
      .addItem('⏱️ Install Weekly Reminder Trigger', 'setupWeeklyReminderTrigger')
      .addItem('⏱️ Install Community Poll Draw Trigger', 'setupCommunityPollTrigger')
      .addItem('🔓 Install onOpen Deferred Trigger', 'setupOpenDeferredTrigger')
      .addSeparator()
      .addItem('🔔 Workload: Setup Reminders', 'setupWorkloadReminderSystem'))

    .addSubMenu(ui.createMenu('🔄 Maintenance')
      .addItem('🔄 Refresh All Formulas', 'refreshAllFormulas')
      .addItem('🔄 Refresh Member Data', 'refreshMemberDirectoryFormulas')
      .addItem('🔄 Refresh View', 'refreshMemberView')
      .addSeparator()
      .addItem('📄 Backfill Minutes → Drive Docs', 'BACKFILL_MINUTES_DRIVE_DOCS')
      .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
      .addItem('📥 Populate Config from Sheet Data', 'populateConfigFromSheetData')
      .addItem('🔧 Repair Config Data Alignment', 'repairConfigData')
      .addItem('🔧 Repair Member Directory Columns', 'repairMemberDirectoryColumns')
      .addItem('📱 Add Mobile Dashboard Link', 'addMobileDashboardLinkToConfig')
      .addItem('🔓 Unlock Checklist Sheet', 'unlockChecklistSheet')
      .addSeparator()
      .addItem('📋 Create Features Reference Sheet', 'createFeaturesReferenceSheet')
      .addItem('❓ Create FAQ Sheet', 'createFAQSheet')
      .addItem('🗑️ Remove Deprecated Tabs', 'removeDeprecatedTabs'))

    .addSubMenu(ui.createMenu('🎨 Styling')
      .addItem('🎨 Apply Config Sheet Styling', 'applyConfigStyling')
      .addItem('🎨 Apply Tab Colors', 'applyTabColors')
      .addItem('📑 Apply Tab Titles', 'migrateSheetTabTitles')
      .addItem('🖌️ Setup Theme Columns', 'setupThemeColumns'))

    .addSubMenu(ui.createMenu('🌐 Web App & Portal')
      .addItem('📱 Get Mobile App URL', 'showWebAppUrl')
      .addItem('👤 Build Member Portal', 'buildPortalForSelectedMember')
      .addItem('📊 Build Public Portal', 'buildPublicPortal')
      .addSeparator()
      .addItem('🔑 Generate Member PIN', 'showGeneratePINDialog')
      .addItem('🔄 Reset Member PIN', 'showResetPINDialog')
      .addItem('📋 Bulk Generate PINs', 'showBulkGeneratePINDialog')
      .addSeparator()
      .addItem('📝 OCR Transcribe Form', 'showOCRDialog')
      .addItem('🔧 OCR Setup', 'setupOCRApiKey'));

  // Only show Demo Data menu if NOT in production mode
  if (!isProductionMode()) {
    adminMenu.addSeparator()
      .addSubMenu(ui.createMenu('🎭 Demo Data')
        .addItem('🚀 Seed All Sample Data', 'SEED_SAMPLE_DATA')
        .addItem('📇 Seed Contact Log Only', 'SEED_CONTACT_LOG')
        .addSeparator()
        .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
        .addItem('📥 Populate Config from Sheet Data', 'populateConfigFromSheetData')
        .addSeparator()
        .addItem('☢️ NUKE SEEDED DATA', 'NUKE_SEEDED_DATA'));
  }

  // Only show Test Runner menu in dev mode (30_TestRunner.gs excluded from prod builds)
  if (!isProductionMode()) {
    adminMenu.addSeparator()
      .addSubMenu(ui.createMenu('🧪 Test Runner')
        .addItem('▶ Run All Tests', 'runTestsFromMenu')
        .addSeparator()
        .addItem('🕐 Enable Daily Trigger', 'setupTestTriggerFromMenu')
        .addItem('🚫 Disable Trigger', 'removeTestTriggerFromMenu'));
  }

  adminMenu.addToUi();
}

// ============================================================================
// NAVIGATION FUNCTIONS
// ============================================================================
/**
 * Navigates to a specific sheet by name
 * @param {string} sheetName - Name of the sheet to navigate to
 * @returns {void}
 */
function navigateToSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Sheet "' + sheetName + '" not found.');
  }
}

/**
 * Shows a toast notification
 * @param {string} message - The message to display
 * @param {string} title - The toast title
 * @returns {void}
 */
function showToast(message, title) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  ss.toast(message, title || 'Info', 3);
}

// Dead code removed: showConfirmation() — zero callers in src

/**
 * Opens Google Calendar in a new browser tab
 */
function openGoogleCalendar() {
  showDialog_(
    '<script>window.open("https://calendar.google.com", "_blank");google.script.host.close();</script>',
    'Opening Calendar...', 1, 1);
}

/**
 * Opens Google Drive in a new browser tab
 */
function openGoogleDrive() {
  showDialog_(
    '<script>window.open("https://drive.google.com", "_blank");google.script.host.close();</script>',
    'Opening Drive...', 1, 1);
}

/**
 * Sets up theme-related columns and applies theme to all sheets.
 * This is the handler for the "Setup Theme Columns" menu item.
 */
function setupThemeColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Applying theme to all sheets...', 'Theme Setup', 2);

  try {
    // Apply the system theme (header styling + alternating rows)
    APPLY_SYSTEM_THEME();

    // Apply tab colors
    if (typeof applyTabColors_ === 'function') {
      applyTabColors_(ss);
    }

    ss.toast('Theme columns and styling applied to all sheets!', 'Theme Setup Complete', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Theme setup error: ' + e.message);
  }
}

/**
 * Displays a simple alert dialog in the spreadsheet UI.
 * @param {string} message - Alert body text
 * @param {string} [title] - Dialog title; defaults to 'Alert'
 */
function showAlert(message, title) {
  SpreadsheetApp.getUi().alert(title || 'Alert', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Navigates to a specific record
 * @param {string} id - The record ID
 * @param {string} type - 'member' or 'grievance'
 * @returns {void}
 */
function navigateToRecord(id, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName, idCol;

  if (type === 'member') {
    sheetName = SHEETS.MEMBER_DIR;
    idCol = MEMBER_COLS.MEMBER_ID;
  } else if (type === 'grievance') {
    sheetName = SHEETS.GRIEVANCE_LOG;
    idCol = GRIEVANCE_COLS.GRIEVANCE_ID;
  } else {
    return;
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No records found in ' + sheetName + '.');
    return;
  }

  var data = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      ss.setActiveSheet(sheet);
      sheet.setActiveRange(sheet.getRange(i + 2, 1));
      break;
    }
  }
}

/**
 * ============================================================================
 * ThemeService.gs - Theme Management and Visual Settings
 * ============================================================================
 *
 * This module handles all theme-related functions including:
 * - Theme application (dark mode, light mode, etc.)
 * - Visual settings persistence
 * - Comfort view / Comfort View-friendly settings
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Theme management and visual settings functions
 * @version 4.51.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// THEME APPLICATION
// ============================================================================

/**
 * Applies the system theme to all sheets
 * @returns {void}
 */
function APPLY_SYSTEM_THEME() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    // Skip hidden calculation sheets
    if (sheetName.indexOf('_Calc') === 0) return;

    applyThemeToSheet_(sheet);
  });

  ss.toast('Theme applied to all sheets!', 'Theme', 3);
}

/**
 * @private Clamps a font size value to the allowed range (8-24).
 * @param {number} size - Raw font size
 * @returns {number} Clamped font size
 */
function clampFontSize_(size) {
  var n = Number(size) || 10;
  return Math.max(8, Math.min(24, n));
}

/**
 * @private Applies theme styling to a single sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to style
 * @param {string} [themeKey] - Theme preset key; defaults to active theme
 * @returns {void}
 */
function applyThemeToSheet_(sheet, themeKey) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastCol < 1 || lastRow < 1) return;

  // Get theme by key or fall back to active theme preset
  var theme = (themeKey && THEME_PRESETS[themeKey]) ? THEME_PRESETS[themeKey] : getActiveThemePreset_();

  // Apply header styling (row 1)
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground(theme.headerBg)
    .setFontColor(theme.headerText)
    .setFontWeight('bold')
    .setFontFamily(theme.font)
    .setFontSize(clampFontSize_(theme.headerSize))
    .setHorizontalAlignment('center');

  // Apply data row styling
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange
      .setFontFamily(theme.font)
      .setFontSize(clampFontSize_(theme.fontSize));

    // F23b: Apply alternating row colors in single setBackgrounds call
    var backgrounds = [];
    for (var row = 2; row <= lastRow; row++) {
      var color = (row % 2 === 0) ? theme.altRow : SHEET_COLORS.BG_WHITE;
      backgrounds.push(new Array(lastCol).fill(color));
    }
    dataRange.setBackgrounds(backgrounds);
  }
}

// ============================================================================
// THEME PRESETS
// ============================================================================

/**
 * Available theme presets that apply to all tabs
 */
var THEME_PRESETS = {
  'default': {
    name: 'Default (Dark Slate)',
    headerBg: SHEET_COLORS.HEADER_SLATE,
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: SHEET_COLORS.BG_OFF_WHITE,
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 215,
    emoji: '🌑'
  },
  'union-blue': {
    name: 'Union Blue',
    headerBg: SHEET_COLORS.HEADER_DARK_BLUE_ALT,
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: SHEET_COLORS.BG_EXTRA_PALE_BLUE,
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 225,
    emoji: '💙'
  },
  'forest': {
    name: 'Forest Green',
    headerBg: SHEET_COLORS.TEXT_GREEN_ALT,
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: '#f0fdf4',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 148,
    emoji: '💚'
  },
  'charcoal': {
    name: 'Charcoal',
    headerBg: '#374151',
    headerText: SHEET_COLORS.BG_LIGHT_GRAY,
    altRow: SHEET_COLORS.BG_VERY_LIGHT_GRAY,
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 218,
    emoji: '⬛'
  },
  'midnight': {
    name: 'Midnight Purple',
    headerBg: '#581c87',
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: SHEET_COLORS.BG_LIGHT_PURPLE,
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 271,
    emoji: '💜'
  },
  'crimson': {
    name: 'Crimson',
    headerBg: '#991b1b',
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: '#fef2f2',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 0,
    emoji: '❤️'
  },
  'steel': {
    name: 'Steel Gray',
    headerBg: '#475569',
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: SHEET_COLORS.BG_SLATE_LIGHT,
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 215,
    emoji: '🩶'
  },
  'ocean': {
    name: 'Ocean Teal',
    headerBg: '#115e59',
    headerText: SHEET_COLORS.BG_WHITE,
    altRow: '#f0fdfa',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11,
    accentHue: 175,
    emoji: '🌊'
  }
};

/**
 * Gets the active theme preset based on saved settings
 * @returns {Object} Theme preset with colors and font settings
 * @private
 */
function getActiveThemePreset_() {
  var themeKey = getCurrentTheme();
  return THEME_PRESETS[themeKey] || THEME_PRESETS['default'];
}

/**
 * Shows theme preset picker dialog
 */
function showThemePresetPicker() {
  var currentTheme = getCurrentTheme();
  var presetNames = Object.keys(THEME_PRESETS);
  var html = '<html><head><style>' +
    'body { font-family: Roboto, sans-serif; padding: 16px; }' +
    '.preset { display: flex; align-items: center; padding: 10px; margin: 6px 0; border-radius: 6px; cursor: pointer; border: 2px solid #e5e7eb; }' +
    '.preset:hover { border-color: #3b82f6; }' +
    '.preset.active { border-color: #2563eb; background: #eff6ff; }' +
    '.swatch { width: 32px; height: 32px; border-radius: 4px; margin-right: 12px; display: flex; align-items: center; justify-content: center; color: white; font-weight: bold; font-size: 11px; }' +
    '.name { font-weight: 500; }' +
    '.btn { padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; margin-top: 12px; }' +
    '.btn-primary { background: #2563eb; color: white; }' +
    '.btn-primary:hover { background: #1d4ed8; }' +
    '</style></head><body>';

  html += '<h3 style="margin-top:0">Choose Theme Preset</h3>';

  for (var i = 0; i < presetNames.length; i++) {
    var key = presetNames[i];
    var preset = THEME_PRESETS[key];
    var activeClass = key === currentTheme ? ' active' : '';
    html += '<div class="preset' + activeClass + '" onclick="selectTheme(\'' + key + '\')">' +
      '<div class="swatch" style="background:' + preset.headerBg + '">Aa</div>' +
      '<div class="name">' + preset.name + '</div></div>';
  }

  html += '<script>' +
    'function selectTheme(key) {' +
    '  var items = document.querySelectorAll(".preset");' +
    '  items.forEach(function(el) { el.style.pointerEvents = "none"; el.style.opacity = "0.6"; });' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { google.script.host.close(); })' +
    '    .withFailureHandler(function(err) {' +
    '      items.forEach(function(el) { el.style.pointerEvents = ""; el.style.opacity = ""; });' +
    '      alert("Failed to apply theme: " + err.message);' +
    '    })' +
    '    .applyThemePreset(key);' +
    '}' +
    '</script></body></html>';

  showDialog_(html, 'Theme Presets', 340, 480);
}

/**
 * Applies a theme preset by key to all sheets
 * @param {string} presetKey - Theme preset key from THEME_PRESETS
 */
function applyThemePreset(presetKey) {
  if (!THEME_PRESETS[presetKey]) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Unknown theme preset: ' + presetKey, 'Error', 3);
    return;
  }

  saveVisualSetting('theme', presetKey);
  // Persist accentHue so the webapp picks it up on next load
  var preset = THEME_PRESETS[presetKey];
  if (preset.accentHue !== undefined) {
    saveVisualSetting('accentHue', preset.accentHue);
  }
  APPLY_SYSTEM_THEME();
  SpreadsheetApp.getActiveSpreadsheet().toast('Applied "' + preset.name + '" theme to all tabs!', 'Theme', 3);
}

/**
 * Returns the user's saved color theme key and accentHue.
 * Used by the webapp to sync sheet theme ↔ webapp accent.
 * @returns {{ themeKey: string, accentHue: number }}
 */
function getUserColorTheme() {
  var themeKey = getCurrentTheme();
  var preset = THEME_PRESETS[themeKey] || THEME_PRESETS['default'];
  return { themeKey: themeKey, accentHue: preset.accentHue || 250 };
}

/**
 * Returns THEME_PRESETS metadata for the webapp color theme picker.
 * @returns {Array<{ key: string, name: string, emoji: string, headerBg: string, accentHue: number }>}
 */
function getColorThemeList() {
  var keys = Object.keys(THEME_PRESETS);
  return keys.map(function(key) {
    var p = THEME_PRESETS[key];
    return { key: key, name: p.name, emoji: p.emoji || '', headerBg: p.headerBg, accentHue: p.accentHue };
  });
}

/**
 * Resets all sheets to default theme
 * @returns {void}
 */
function resetToDefaultTheme() {
  saveVisualSetting('theme', 'default');
  APPLY_SYSTEM_THEME();
  SpreadsheetApp.getActiveSpreadsheet().toast('Theme reset to default!', 'Theme', 3);
}

/**
 * Saves a visual setting
 * @param {string} setting - The setting name
 * @param {*} value - The setting value
 * @returns {void}
 */
function saveVisualSetting(setting, value) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('visual_' + setting, JSON.stringify(value));
}

/**
 * Applies a dashboard theme
 * @param {string} theme - Theme name ('default', 'dark', 'light', 'contrast')
 * @returns {void}
 */
function applyDashboardTheme(theme) {
  saveVisualSetting('theme', theme);
  APPLY_SYSTEM_THEME();
  SpreadsheetApp.getActiveSpreadsheet().toast('Theme changed to ' + theme, 'Theme', 2);
}

/**
 * Toggles dark mode
 * @returns {void}
 */
// Dead code removed (v4.55.2): toggleDarkMode() — the Visual Control Panel in
// 04a_UIMenus.gs uses saveVisualSetting('darkMode', ...) directly and never
// called this function. Left the panel as the single source of truth.
// ============================================================================
// COMFORT VIEW SETTINGS
// ============================================================================

/**
 * Gets Comfort View settings
 * @returns {Object} Settings object
 */
function getComfortViewSettings() {
  var props = PropertiesService.getUserProperties();
  var settings = props.getProperty('comfortViewSettings');
  if (!settings) return getDefaultComfortViewSettings_();
  try { return JSON.parse(settings); } catch (_) { return getDefaultComfortViewSettings_(); }
}

/**
 * Gets default Comfort View settings
 * @returns {Object} Default settings
 * @private
 */
function getDefaultComfortViewSettings_() {
  return {
    zebraStripes: true,
    reducedMotion: false,
    focusMode: false,
    hideGridlines: false
  };
}

/**
 * Saves Comfort View settings
 * @param {Object} settings - The settings to save
 * @returns {void}
 */
function saveComfortViewSettings(settings) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('comfortViewSettings', JSON.stringify(settings));
}

/**
 * Applies Comfort View settings
 * @param {Object} settings - The settings to apply
 * @returns {void}
 */
function applyComfortViewSettings(settings) {
  if (settings.zebraStripes) {
    applyZebraStripesToAllSheets_();
  }
  if (settings.hideGridlines) {
    hideAllGridlines();
  }
  saveComfortViewSettings(settings);
}

// Dead code removed (v4.55.2): resetComfortViewSettings() — no callers.

/**
 * Hides gridlines on all sheets
 * @returns {void}
 */
function hideAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    sheet.setHiddenGridlines(true);
  });
}
// Dead code removed (v4.55.2): toggleGridlinesComfortView() — no callers.

/**
 * Applies zebra stripes to a sheet
 * @param {Sheet} sheet - The sheet to style
 * @returns {void}
 */
function applyZebraStripes(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return;

  // F23b: Build backgrounds array and apply in single setBackgrounds call
  var backgrounds = [];
  for (var row = 2; row <= lastRow; row++) {
    var color = (row % 2 === 0) ? SHEET_COLORS.BG_SLATE_LIGHT : SHEET_COLORS.BG_WHITE;
    var rowColors = new Array(lastCol).fill(color);
    backgrounds.push(rowColors);
  }
  sheet.getRange(2, 1, lastRow - 1, lastCol).setBackgrounds(backgrounds);
}

/**
 * Applies zebra stripes to all sheets
 * @returns {void}
 * @private
 */
function applyZebraStripesToAllSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getName().indexOf('_Calc') !== 0) {
      applyZebraStripes(sheet);
    }
  });
}

/**
 * Removes zebra stripes from a sheet
 * @param {Sheet} sheet - The sheet to clear
 * @returns {void}
 */
function removeZebraStripes(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return;

  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.setBackground(SHEET_COLORS.BG_WHITE);
}

// Dead code removed (v4.55.2): toggleZebraStripes / activateFocusMode /
// deactivateFocusMode — the Visual Control Panel in 04a_UIMenus.gs is the
// only active entry point for these settings. The "Press Esc to exit" toast
// in activateFocusMode was also a UI lie — GAS can't intercept Esc, so that
// instruction would never have worked anyway.

/**
 * Gets the current theme
 * @returns {string} Current theme name
 */
function getCurrentTheme() {
  var props = PropertiesService.getUserProperties();
  var stored = props.getProperty('visual_theme');
  if (!stored) return 'default';
  try {
    return JSON.parse(stored);
  } catch (_e) {
    // Backward compatibility: old theme values were plain strings, not JSON
    return stored;
  }
}

/**
 * Applies a theme
 * @param {string} themeKey - Theme key
 * @param {string} scope - 'all' or 'current'
 * @returns {void}
 */
function applyTheme(themeKey, scope) {
  saveVisualSetting('theme', themeKey);

  if (scope === 'all') {
    APPLY_SYSTEM_THEME();
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    applyThemeToSheet_(sheet);
  }
}

/**
 * Sets up Comfort View defaults
 * @returns {void}
 */
function setupComfortViewDefaults() {
  var settings = getDefaultComfortViewSettings_();
  settings.zebraStripes = true;
  applyComfortViewSettings(settings);
  SpreadsheetApp.getActiveSpreadsheet().toast('Comfort View defaults applied', 'Settings', 3);
}
/**
 * Refreshes all visual elements
 * @returns {void}
 */
function refreshAllVisuals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var startTime = Date.now();
  var MAX_RUNTIME_MS = 120000; // 2 minute safety limit

  ss.toast('Refreshing visuals...', 'Refresh', 2);

  try {
    // Reapply theme with timeout protection
    var sheets = ss.getSheets();
    var processed = 0;
    for (var i = 0; i < sheets.length; i++) {
      // Check timeout
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        ss.toast('Refresh stopped after processing ' + processed + '/' + sheets.length + ' sheets (timeout).', 'Refresh', 5);
        return;
      }

      var sheetName = sheets[i].getName();
      if (sheetName.indexOf('_Calc') === 0 || sheetName.indexOf('_Audit') === 0) continue;

      try {
        applyThemeToSheet_(sheets[i]);
        processed++;
      } catch (_e) {
        log_('refreshAllVisuals', 'Skipped theme for sheet ' + sheetName + ': ' + _e.message);
      }
    }

    // Reapply Comfort View settings if enabled
    var settings = getComfortViewSettings();
    if (settings.zebraStripes) {
      applyZebraStripesToAllSheets_();
    }

    ss.toast('All visuals refreshed! (' + processed + ' sheets)', 'Refresh', 3);
  } catch (e) {
    ss.toast('Refresh error: ' + e.message, 'Error', 5);
    log_('refreshAllVisuals', 'Error: ' + e.toString());
  }
}

/**
 * ============================================================================
 * MobileInterface.gs - Mobile UI Components and Dashboard
 * ============================================================================
 *
 * This module handles all mobile-related functions including:
 * - Mobile context detection
 * - Mobile-optimized dashboard views
 * - Mobile grievance list and search interfaces
 * - Touch-friendly UI components
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Mobile interface components and dashboard functions
 * @version 4.51.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MOBILE GRIEVANCE DATA
// ============================================================================

/**
 * Gets recent grievances formatted for mobile display
 * @param {number} [limit=5] Maximum number of grievances to return
 * @returns {Array<Object>} Array of grievance objects with mobile-friendly properties
 */
// Dead code removed (v4.55.2): getRecentGrievancesForMobile() — zero callers.

/**
 * Gets search results for mobile search interface
 * @param {string} query Search query string
 * @param {string} tab Active tab ('all', 'members', or 'grievances')
 * @returns {Array<Object>} Array of search result objects
 */
function getMobileSearchData(query, tab) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  query = query.toLowerCase();
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
      mData.forEach(function(row) {
        var id = col_(row, MEMBER_COLS.MEMBER_ID) || '';
        var name = (col_(row, MEMBER_COLS.FIRST_NAME) || '') + ' ' + (col_(row, MEMBER_COLS.LAST_NAME) || '');
        var email = col_(row, MEMBER_COLS.EMAIL) || '';
        if (id.toLowerCase().indexOf(query) >= 0 || name.toLowerCase().indexOf(query) >= 0 || email.toLowerCase().indexOf(query) >= 0) {
          results.push({ type: 'member', title: name, subtitle: id, detail: email });
        }
      });
    }
  }
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, GRIEVANCE_COLS.STATUS).getValues();
      gData.forEach(function(row) {
        var id = col_(row, GRIEVANCE_COLS.GRIEVANCE_ID) || '';
        var name = (col_(row, GRIEVANCE_COLS.FIRST_NAME) || '') + ' ' + (col_(row, GRIEVANCE_COLS.LAST_NAME) || '');
        var status = col_(row, GRIEVANCE_COLS.STATUS) || '';
        if (id.toLowerCase().indexOf(query) >= 0 || name.toLowerCase().indexOf(query) >= 0) {
          results.push({ type: 'grievance', title: id, subtitle: name, detail: status });
        }
      });
    }
  }
  return results.slice(0, 20);
}

// ============================================================================
// MY ASSIGNED GRIEVANCES
// ============================================================================

/**
 * Shows grievances assigned to the current user
 * Displays a quick summary dialog of assigned cases
 * @returns {void}
 */
// Dead code removed (v4.55.2): showMyAssignedGrievances() — zero callers. The
// steward "My Cases" view in steward_view.html replaces this.

/**
 * ============================================================================
 * QuickActions.gs - Quick Actions and Email Functions
 * ============================================================================
 *
 * This module handles all quick action and email-related functions including:
 * - Quick Actions menu for Member Directory and Grievance Log
 * - Member email composition and sending
 * - Survey, contact form, and dashboard link emails
 * - Grievance status email updates
 * - Member grievance history display
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Quick actions menu and member email functions
 * @version 4.51.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// QUICK ACTIONS MENU - Main Entry Points
// ============================================================================

/**
 * Shows the Quick Actions menu based on current sheet and selection
 * Routes to appropriate quick actions dialog for Member Directory or Grievance Log
 */
function showQuickActionsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  var name = sheet.getName();
  var selection = sheet.getActiveRange();

  if (!selection) {
    ui.alert('⚡ Quick Actions - How to Use',
      'Quick Actions provides contextual shortcuts for the selected row.\n\n' +
      'TO USE:\n' +
      '1. Go to Member Directory or Grievance Log\n' +
      '2. Click on a data row (not the header)\n' +
      '3. Run this menu item again\n\n' +
      'MEMBER DIRECTORY ACTIONS:\n' +
      '• Start new grievance for member\n' +
      '• Send email to member\n' +
      '• View grievance history\n' +
      '• Copy Member ID\n\n' +
      'GRIEVANCE LOG ACTIONS:\n' +
      '• Sync deadlines to calendar\n' +
      '• Setup Drive folder\n' +
      '• Quick status update\n' +
      '• Copy Grievance ID',
      ui.ButtonSet.OK);
    return;
  }

  var row = selection.getRow();
  if (row < 2) {
    ui.alert('Quick Actions',
      'Please select a data row, not the header row.\n\n' +
      'Click on row 2 or below to use Quick Actions.',
      ui.ButtonSet.OK);
    return;
  }

  if (name === SHEETS.MEMBER_DIR) {
    showMemberQuickActions(row);
  } else if (name === SHEETS.GRIEVANCE_LOG) {
    showGrievanceQuickActions(row);
  } else {
    ui.alert('⚡ Quick Actions',
      'Quick Actions is available for:\n\n' +
      '• Member Directory - actions for members\n' +
      '• Grievance Log - actions for grievances\n\n' +
      'Current sheet: ' + name + '\n\n' +
      'Navigate to one of the supported sheets and select a row.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// MEMBER QUICK ACTIONS
// ============================================================================

/**
 * Shows quick actions dialog for a member row in Member Directory
 * @param {number} row - The selected row number
 */
function showMemberQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) { SpreadsheetApp.getUi().alert('Member Directory sheet not found.'); return; }
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var memberId = col_(data, MEMBER_COLS.MEMBER_ID);
  var name = col_(data, MEMBER_COLS.FIRST_NAME) + ' ' + col_(data, MEMBER_COLS.LAST_NAME);
  var email = col_(data, MEMBER_COLS.EMAIL);
  var hasOpen = col_(data, MEMBER_COLS.HAS_OPEN_GRIEVANCE);

  // Build email section buttons (only if email exists)
  var emailButtons = '';
  if (email) {
    // CR-XSS-6: Use JSON.stringify for memberId in onclick JS string context.
    // escapeHtml() is wrong here — HTML entities get decoded by the parser before
    // JS executes, so &apos; → ' which breaks single-quoted string delimiters.
    // JSON.stringify() produces a double-quoted JS string with correct escape sequences.
    var memberIdJs = JSON.stringify(memberId);
    emailButtons = `<div class="section-header">📨 Email Options</div>
      <button class="action-btn" onclick="google.script.run.composeEmailForMember(${memberIdJs});google.script.host.close()"><span class="icon">📧</span><span><div class="title">Send Custom Email</div><div class="desc">Compose email to ${escapeHtml(email)}</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(${memberIdJs});google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(${memberIdJs});google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailDashboardLinkToMember(${memberIdJs});google.script.host.close()"><span class="icon">🔗</span><span><div class="title">Send Dashboard Link</div><div class="desc">Share dashboard access with member</div></span></button>`;
  }

  var memberIdJson = JSON.stringify(memberId);
  showDialog_(`<!DOCTYPE html><html><head><base target="_top">
    ${getMobileOptimizedHead()}
    <style>
    body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}
    .container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}
    h2{color:${SHEET_COLORS.DIALOG_ACCENT};display:flex;align-items:center;gap:10px;font-size:clamp(16px,4.5vw,20px)}
    .info{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px}
    .name{font-size:clamp(15px,4vw,18px);font-weight:bold}
    .id{color:#666;font-size:clamp(12px,3vw,14px)}
    .status{margin-top:10px}
    .badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:clamp(10px,2.8vw,12px);font-weight:bold}
    .open{background:#ffebee;color:#c62828}
    .none{background:#e8f5e9;color:#2e7d32}
    .actions{display:flex;flex-direction:column;gap:10px}
    .action-btn{display:flex;align-items:center;gap:12px;padding:clamp(12px,3vw,15px);border:none;border-radius:8px;cursor:pointer;font-size:clamp(13px,3.5vw,14px);text-align:left;background:#f8f9fa;min-height:44px}
    .action-btn:hover{background:#e8f4fd}
    .icon{font-size:clamp(20px,5vw,24px)}
    .title{font-weight:bold}
    .desc{font-size:clamp(10px,2.8vw,12px);color:#666;margin-top:2px}
    .section-header{font-weight:bold;color:${SHEET_COLORS.DIALOG_ACCENT};margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0;font-size:clamp(12px,3vw,14px)}
    .close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer;min-height:44px;font-size:clamp(13px,3.5vw,14px)}
    </style></head><body><div class="container">
    <h2>⚡ Quick Actions</h2>
    <div class="info">
      <div class="name">${escapeHtml(name)}</div>
      <div class="id">${escapeHtml(memberId)} | ${escapeHtml(email || 'No email')}</div>
      <div class="status">${hasOpen === 'Yes' ? '<span class="badge open">🔴 Has Open Grievance</span>' : '<span class="badge none">🟢 No Open Grievances</span>'}</div>
    </div>
    <div class="actions">
      <div class="section-header">📋 Member Actions</div>
      <button class="action-btn" onclick="google.script.run.showMemberGrievanceHistory(${memberIdJson});google.script.host.close()"><span class="icon">📁</span><span><div class="title">View Grievance History</div><div class="desc">See all grievances for this member</div></span></button>
      <button class="action-btn" onclick="navigator.clipboard.writeText(${memberIdJson});alert('Copied!')"><span class="icon">📋</span><span><div class="title">Copy Member ID</div><div class="desc">${escapeHtml(memberId)}</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(r){if(r.success){alert(r.message+'\\n'+r.folderUrl)}else{alert('Error: '+r.error)}}).withFailureHandler(function(e){alert(e.message)}).setupDriveFolderForMember(${memberIdJson})"><span class="icon">📁</span><span><div class="title">Create Member Folder</div><div class="desc">Setup Google Drive folder for this member</div></span></button>
      ${emailButtons}
    </div>
    <button class="close" onclick="google.script.host.close()">Close</button>
    </div></body></html>`,
    'Member Quick Actions', 400, email ? 650 : 400);
}

// ============================================================================
// GRIEVANCE QUICK ACTIONS
// ============================================================================

/**
 * Shows quick actions dialog for a grievance row in Grievance Log
 * @param {number} row - The selected row number
 */
function showGrievanceQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) { SpreadsheetApp.getUi().alert('Grievance Log sheet not found.'); return; }
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var grievanceId = col_(data, GRIEVANCE_COLS.GRIEVANCE_ID);
  var memberId = col_(data, GRIEVANCE_COLS.MEMBER_ID);
  var name = col_(data, GRIEVANCE_COLS.FIRST_NAME) + ' ' + col_(data, GRIEVANCE_COLS.LAST_NAME);
  var status = col_(data, GRIEVANCE_COLS.STATUS);
  var step = col_(data, GRIEVANCE_COLS.CURRENT_STEP);
  var daysTo = col_(data, GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var memberEmail = col_(data, GRIEVANCE_COLS.MEMBER_EMAIL);
  var isOpen = status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING || status === GRIEVANCE_STATUS.IN_ARBITRATION || status === GRIEVANCE_STATUS.APPEALED;

  // Build email button (only if member has email)
  var emailStatusBtn = '';
  if (memberEmail) {
    var gIdJson = JSON.stringify(grievanceId);
    var mIdJson = JSON.stringify(memberId);
    emailStatusBtn = `<div class="section-header">📨 Communication</div>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailGrievanceStatusToMember(${gIdJson});google.script.host.close()"><span class="icon">📧</span><span><div class="title">Email Status to Member</div><div class="desc">Send grievance status update to ${escapeHtml(String(memberEmail))}</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(${mIdJson});google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>
      <button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(${mIdJson});google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>`;
  }

  var grievanceIdJson = JSON.stringify(grievanceId);
  var daysOverdue = daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0);
  var daysBadge = daysTo !== null && daysTo !== ''
    ? '<span class="badge" style="background:' + (daysOverdue ? '#ffebee;color:#c62828' : '#e3f2fd;color:#1565c0') + '">' + (daysOverdue ? '⚠️ Overdue' : '📅 ' + escapeHtml(String(daysTo)) + ' days') + '</span>'
    : '';
  var statusSection = isOpen
    ? `<div class="status-section"><h4>Quick Status Update</h4>
       <select id="statusSelect"><option value="">-- Select --</option><option value="Open">Open</option><option value="Pending Info">Pending Info</option><option value="Settled">Settled</option><option value="Withdrawn">Withdrawn</option><option value="Won">Won</option><option value="Denied">Denied</option><option value="Closed">Closed</option></select>
       <button class="action-btn" style="margin-top:10px" onclick="var s=document.getElementById('statusSelect').value;if(!s){alert('Select status');return}google.script.run.withSuccessHandler(function(){alert('Updated!');google.script.host.close()}).quickUpdateGrievanceStatus(${row},s)"><span class="icon">✓</span><span><div class="title">Update Status</div></span></button></div>`
    : '';

  showDialog_(`<!DOCTYPE html><html><head><base target="_top">
    ${getMobileOptimizedHead()}
    <style>
    body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}
    .container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}
    h2{color:#DC2626;font-size:clamp(16px,4.5vw,20px)}
    .info{background:#fff5f5;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;border-left:4px solid #DC2626}
    .gid{font-size:clamp(15px,4vw,18px);font-weight:bold}
    .gmem{color:#666;font-size:clamp(12px,3vw,14px)}
    .gstatus{margin-top:10px;display:flex;gap:8px;flex-wrap:wrap}
    .badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:clamp(10px,2.8vw,12px);font-weight:bold}
    .actions{display:flex;flex-direction:column;gap:10px}
    .action-btn{display:flex;align-items:center;gap:12px;padding:clamp(12px,3vw,15px);border:none;border-radius:8px;cursor:pointer;font-size:clamp(13px,3.5vw,14px);text-align:left;background:#f8f9fa;min-height:44px}
    .action-btn:hover{background:#fff5f5}
    .icon{font-size:clamp(20px,5vw,24px)}
    .title{font-weight:bold}
    .desc{font-size:clamp(10px,2.8vw,12px);color:#666;margin-top:2px}
    .section-header{font-weight:bold;color:#DC2626;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0;font-size:clamp(12px,3vw,14px)}
    .status-section{margin-top:15px;padding:clamp(10px,3vw,15px);background:#f8f9fa;border-radius:8px}
    .status-section h4{margin:0 0 10px;font-size:clamp(13px,3.5vw,15px)}
    select{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:16px;min-height:44px}
    .close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer;min-height:44px;font-size:clamp(13px,3.5vw,14px)}
    </style></head><body><div class="container">
    <h2>⚡ Grievance Actions</h2>
    <div class="info">
      <div class="gid">${escapeHtml(grievanceId)}</div>
      <div class="gmem">${escapeHtml(name)} (${escapeHtml(memberId)})${memberEmail ? ' - ' + escapeHtml(memberEmail) : ''}</div>
      <div class="gstatus">
        <span class="badge">${escapeHtml(status)}</span>
        <span class="badge">${escapeHtml(step)}</span>
        ${daysBadge}
      </div>
    </div>
    <div class="actions">
      <div class="section-header">📋 Case Management</div>
      <button class="action-btn" onclick="google.script.run.syncSingleGrievanceToCalendar(${grievanceIdJson});google.script.host.close()"><span class="icon">📅</span><span><div class="title">Sync to Calendar</div><div class="desc">Add deadlines to Google Calendar</div></span></button>
      <button class="action-btn" onclick="google.script.run.setupDriveFolderForGrievance(${grievanceIdJson});google.script.host.close()"><span class="icon">📁</span><span><div class="title">Setup Drive Folder</div><div class="desc">Create document folder</div></span></button>
      <button class="action-btn" onclick="navigator.clipboard.writeText(${grievanceIdJson});alert('Copied!')"><span class="icon">📋</span><span><div class="title">Copy Grievance ID</div><div class="desc">${escapeHtml(grievanceId)}</div></span></button>
      ${emailStatusBtn}
    </div>
    ${statusSection}
    <button class="close" onclick="google.script.host.close()">Close</button>
    </div></body></html>`,
    'Grievance Quick Actions', 400, memberEmail ? 750 : 550);
}

/**
 * Quick update grievance status from the quick actions dialog
 * @param {number} row - The grievance row number
 * @param {string} newStatus - The new status value
 */
function quickUpdateGrievanceStatus(row, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');

  // Validate row bounds
  if (row < 2 || row > sheet.getLastRow()) {
    throw new Error('Invalid row number: ' + row);
  }

  // Validate status against allowlist
  var validStatuses = ['Open', 'Pending Info', 'Settled', 'Withdrawn', 'Won', 'Denied', 'Closed', 'In Arbitration', 'Appealed'];
  if (validStatuses.indexOf(newStatus) === -1) {
    throw new Error('Invalid status: ' + newStatus);
  }

  sheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newStatus);
  if (['Closed', 'Settled', 'Withdrawn'].indexOf(newStatus) >= 0) {
    var closeCol = GRIEVANCE_COLS.DATE_CLOSED;
    if (!sheet.getRange(row, closeCol).getValue()) sheet.getRange(row, closeCol).setValue(new Date());
  }
  ss.toast('Grievance status updated to: ' + newStatus, 'Status Updated', 3);
}

// ============================================================================
// EMAIL COMPOSITION AND SENDING
// ============================================================================

/**
 * Opens email composition dialog for a member
 * @param {string} memberId - The member ID to compose email for
 */
function composeEmailForMember(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() <= 1) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], MEMBER_COLS.MEMBER_ID) === memberId) {
      var email = col_(data[i], MEMBER_COLS.EMAIL);
      var name = col_(data[i], MEMBER_COLS.FIRST_NAME) + ' ' + col_(data[i], MEMBER_COLS.LAST_NAME);
      if (!email) { SpreadsheetApp.getUi().alert('No email on file.'); return; }
      showDialog_(
        '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}.container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}h2{color:' + SHEET_COLORS.DIALOG_ACCENT + ';font-size:clamp(16px,4.5vw,20px)}.info{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;font-size:clamp(12px,3vw,14px)}.form-group{margin:15px 0}label{display:block;font-weight:bold;margin-bottom:5px;font-size:clamp(12px,3vw,14px)}input,textarea{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:16px;box-sizing:border-box;min-height:44px}textarea{min-height:180px}input:focus,textarea:focus{outline:none;border-color:' + SHEET_COLORS.DIALOG_ACCENT + '}.buttons{display:flex;gap:10px;margin-top:20px}button{padding:12px 24px;font-size:clamp(13px,3.5vw,14px);border:none;border-radius:4px;cursor:pointer;flex:1;min-height:44px}.primary{background:' + SHEET_COLORS.DIALOG_ACCENT + ';color:white}.secondary{background:#6c757d;color:white}@media(max-width:480px){.buttons{flex-direction:column}}</style></head><body><div class="container"><h2>📧 Email to Member</h2><div class="info"><strong>' + escapeHtml(name) + '</strong> (' + escapeHtml(memberId) + ')<br>' + escapeHtml(email) + '</div><div class="form-group"><label>Subject:</label><input type="text" id="subject" placeholder="Email subject"></div><div class="form-group"><label>Message:</label><textarea id="message" placeholder="Type your message..."></textarea></div><div class="buttons"><button class="primary" onclick="send()">📤 Send</button><button class="secondary" onclick="google.script.host.close()">Cancel</button></div></div><script>function send(){var s=document.getElementById("subject").value.trim();var m=document.getElementById("message").value.trim();if(!s||!m){alert("Fill in subject and message");return}google.script.run.withSuccessHandler(function(){alert("Email sent!");google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).sendQuickEmail(' + JSON.stringify(email) + ',s,m,' + JSON.stringify(memberId) + ')}</script></body></html>',
        '📧 Compose Email', 600, 500);
      return;
    }
  }
}

/**
 * Sends a quick email to a member
 * @param {string} to - Recipient email address
 * @param {string} subject - Email subject
 * @param {string} body - Email body text
 * @param {string} memberId - Member ID for logging purposes
 * @returns {Object} Success status object
 */
function sendQuickEmail(to, subject, body, memberId) {
  try {
    // Validate email format
    var EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    var sanitizedTo = String(to || '').trim();
    if (!EMAIL_REGEX.test(sanitizedTo)) {
      throw new Error('Invalid email format');
    }

    // Verify the email matches the member record (prevent address tampering)
    if (memberId) {
      var memberData = getMemberDataById_(memberId);
      if (!memberData || !memberData.email) {
        throw new Error('Member not found or has no email on file');
      }
      if (String(memberData.email).trim().toLowerCase() !== sanitizedTo.toLowerCase()) {
        throw new Error('Email address does not match the member record');
      }
    }

    // Rate limiting: max 10 emails per minute per user
    var cache = CacheService.getScriptCache();
    var callerEmail = '';
    try { callerEmail = Session.getActiveUser().getEmail(); } catch (_e) { log_('sendQuickEmail', 'Error: ' + (_e.message || _e)); }
    var rateLimitKey = 'email_rate_' + (callerEmail || 'unknown');
    var emailCount = parseInt(cache.get(rateLimitKey) || '0', 10);
    if (emailCount >= 10) {
      throw new Error('Rate limit exceeded. Please wait before sending more emails.');
    }
    cache.put(rateLimitKey, String(emailCount + 1), 60); // 60-second TTL

    // Validate subject and body are non-empty strings
    var safeSubject = String(subject || '').trim();
    var safeBody = String(body || '').trim();
    if (!safeSubject || !safeBody) {
      throw new Error('Subject and message body are required');
    }

    MailApp.sendEmail({ to: sanitizedTo, subject: safeSubject, body: safeBody, name: getOrgNameFromConfig_() + ' Dashboard' });
    return { success: true };
  } catch (e) { throw new Error('Failed to send: ' + e.message); }
}

// ============================================================================
// QUICK ACTION EMAIL FUNCTIONS - Send Forms, Surveys, and Status Updates
// ============================================================================

/**
 * Email the satisfaction survey link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailSurveyToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var surveyUrl = getFormUrlFromConfig('satisfaction');
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Member Satisfaction Survey';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'We value your feedback! Please take a few minutes to complete our Member Satisfaction Survey.\n\n' +
    'Survey Link: ' + surveyUrl + '\n\n' +
    'Your responses help us improve our representation and services.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Survey sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send survey: ' + e.message);
  }
}

/**
 * Email the contact info update form link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailContactFormToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var formUrl = getFormUrlFromConfig('contact');
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Update Your Contact Information';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'Please take a moment to verify and update your contact information. ' +
    'Keeping your information current helps us serve you better.\n\n' +
    'Update Form: ' + formUrl + '\n\n' +
    'This only takes a minute and helps ensure you receive important updates.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Contact form sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send contact form: ' + e.message);
  }
}

/**
 * Email the member dashboard/portal link to a member
 * @param {string} memberId - Member ID to look up email
 */
function emailDashboardLinkToMember(memberId) {
  var memberData = getMemberDataById_(memberId);
  if (!memberData || !memberData.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  // Use the deployed web app URL with personalized member ID parameter
  var webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert('Web app is not deployed. Please deploy the web app first via Extensions > Apps Script > Deploy.');
    return;
  }
  var portalUrl = webAppUrl + '?id=' + encodeURIComponent(memberId);
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Member Dashboard Access';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'You can access your personalized union member dashboard at:\n\n' +
    'Dashboard Link: ' + portalUrl + '\n\n' +
    'From the dashboard you can:\n' +
    '- View your member profile\n' +
    '- Track grievance status (if applicable)\n' +
    '- Stay updated on union activities\n\n' +
    'Keep this link private - it is personalized for you.\n\n' +
    'If you have any questions, please contact your steward.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: memberData.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard link sent to ' + memberData.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send dashboard link: ' + e.message);
  }
}

/**
 * Email grievance status update to the member
 * @param {string} grievanceId - Grievance ID to look up details
 */
function emailGrievanceStatusToMember(grievanceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('Grievance Log not found.');
    return;
  }

  // Find the grievance
  var data = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var grievance = null;

  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID) === grievanceId) {
      grievance = {
        id: col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID),
        memberId: col_(data[i], GRIEVANCE_COLS.MEMBER_ID),
        firstName: col_(data[i], GRIEVANCE_COLS.FIRST_NAME),
        lastName: col_(data[i], GRIEVANCE_COLS.LAST_NAME),
        status: col_(data[i], GRIEVANCE_COLS.STATUS),
        step: col_(data[i], GRIEVANCE_COLS.CURRENT_STEP),
        dateFiled: col_(data[i], GRIEVANCE_COLS.DATE_FILED),
        nextAction: col_(data[i], GRIEVANCE_COLS.NEXT_ACTION_DUE),
        daysToDeadline: col_(data[i], GRIEVANCE_COLS.DAYS_TO_DEADLINE),
        issueCategory: col_(data[i], GRIEVANCE_COLS.ISSUE_CATEGORY),
        steward: col_(data[i], GRIEVANCE_COLS.STEWARD),
        email: col_(data[i], GRIEVANCE_COLS.MEMBER_EMAIL)
      };
      break;
    }
  }

  if (!grievance) {
    SpreadsheetApp.getUi().alert('Grievance not found: ' + grievanceId);
    return;
  }

  if (!grievance.email) {
    SpreadsheetApp.getUi().alert('No email address on file for this member.');
    return;
  }

  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Grievance Status Update (' + grievance.id + ')';
  var body = 'Dear ' + grievance.firstName + ',\n\n' +
    'Here is the current status of your grievance:\n\n' +
    '================================\n' +
    'GRIEVANCE STATUS UPDATE\n' +
    '================================\n\n' +
    'Grievance ID: ' + grievance.id + '\n' +
    'Issue Category: ' + (grievance.issueCategory || 'Not specified') + '\n' +
    'Current Status: ' + grievance.status + '\n' +
    'Current Step: ' + grievance.step + '\n' +
    'Date Filed: ' + (grievance.dateFiled ? Utilities.formatDate(new Date(grievance.dateFiled), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'N/A') + '\n';

  if (grievance.daysToDeadline !== null && grievance.daysToDeadline !== '') {
    if (grievance.daysToDeadline === 'Overdue' || (typeof grievance.daysToDeadline === 'number' && grievance.daysToDeadline < 0)) {
      body += 'Next Deadline: OVERDUE\n';
    } else {
      body += 'Days Until Next Deadline: ' + grievance.daysToDeadline + '\n';
    }
  }

  if (grievance.steward) {
    body += 'Assigned Steward: ' + grievance.steward + '\n';
  }

  body += '\n================================\n\n' +
    'If you have any questions about your grievance, please contact your steward.\n\n' +
    'In Solidarity,\n' +
    orgName;

  try {
    MailApp.sendEmail({
      to: grievance.email,
      subject: subject,
      body: body,
      name: orgName + ' Dashboard'
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Status update sent to ' + grievance.email, 'Email Sent', 3);
    return { success: true };
  } catch (e) {
    throw new Error('Failed to send status update: ' + e.message);
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Helper: Get member data by Member ID
 * @private
 * @param {string} memberId - The member ID to look up
 * @returns {Object|null} Member data object or null if not found
 */
function getMemberDataById_(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (col_(data[i], MEMBER_COLS.MEMBER_ID) === memberId) {
      return {
        memberId: memberId,
        firstName: col_(data[i], MEMBER_COLS.FIRST_NAME),
        lastName: col_(data[i], MEMBER_COLS.LAST_NAME),
        email: col_(data[i], MEMBER_COLS.EMAIL)
      };
    }
  }
  return null;
}

// getOrgNameFromConfig_() is defined in 01_Core.gs as the single source of truth

// ============================================================================
// MEMBER GRIEVANCE HISTORY AND ACTIONS
// ============================================================================

/**
 * Shows grievance history dialog for a specific member
 * @param {string} memberId - The member ID to show history for
 */
function showMemberGrievanceHistory(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) { SpreadsheetApp.getUi().alert('No grievances found.'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.DATE_CLOSED).getValues();
  var mine = [];
  data.forEach(function(row) {
    if (col_(row, GRIEVANCE_COLS.MEMBER_ID) === memberId) {
      mine.push({ id: col_(row, GRIEVANCE_COLS.GRIEVANCE_ID), status: col_(row, GRIEVANCE_COLS.STATUS), step: col_(row, GRIEVANCE_COLS.CURRENT_STEP), issue: col_(row, GRIEVANCE_COLS.ISSUE_CATEGORY), filed: col_(row, GRIEVANCE_COLS.DATE_FILED), closed: col_(row, GRIEVANCE_COLS.DATE_CLOSED) });
    }
  });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances for this member.'); return; }
  var list = mine.map(function(g) {
    return '<div style="background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;border-left:4px solid ' + (g.status === GRIEVANCE_STATUS.OPEN ? '#f44336' : '#4caf50') + '"><strong>' + escapeHtml(g.id) + '</strong><br><span style="color:#666">Status: ' + escapeHtml(g.status) + ' | Step: ' + escapeHtml(g.step) + '</span><br><span style="color:#888;font-size:12px">' + escapeHtml(g.issue) + ' | Filed: ' + (g.filed ? Utilities.formatDate(new Date(g.filed), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'N/A') + '</span></div>';
  }).join('');
  showDialog_(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px)}h2{color:' + SHEET_COLORS.DIALOG_ACCENT + ';font-size:clamp(16px,4.5vw,20px)}.summary{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;font-size:clamp(12px,3vw,14px)}</style></head><body><h2>📁 Grievance History</h2><div class="summary"><strong>Member ID:</strong> ' + escapeHtml(memberId) + '<br><strong>Total:</strong> ' + mine.length + '<br><strong>Open:</strong> ' + mine.filter(function(g) { return g.status === GRIEVANCE_STATUS.OPEN; }).length + '<br><strong>Closed:</strong> ' + mine.filter(function(g) { return g.status !== GRIEVANCE_STATUS.OPEN; }).length + '</div>' + list + '</body></html>',
    'Grievance History - ' + memberId, 500, 500);
}

// openGrievanceFormForMember was a placeholder that showed an alert telling
// users to add a row manually. The quick-actions button that called it has
// been removed — use the Grievance Log directly to start a new grievance.

/**
 * Syncs a single grievance to the calendar. Delegates to
 * syncDeadlinesToCalendar(grievanceIdFilter) so the sync is actually scoped
 * to the requested id (the previous version fired a full org-wide sync).
 * @param {string} grievanceId - The grievance ID to sync
 */
function syncSingleGrievanceToCalendar(grievanceId) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing ' + grievanceId + '...', 'Calendar', 3);
  if (typeof syncDeadlinesToCalendar === 'function') syncDeadlinesToCalendar(grievanceId);
}

/**
 * ============================================================================
 * SearchDialogs.gs - Search Dialog Components
 * ============================================================================
 *
 * This module handles all search-related dialog interfaces including:
 * - Desktop search dialog
 * - Quick search dialog
 * - Advanced search dialog with filtering
 *
 * Extracted from UIService.gs for better modularity and maintainability.
 *
 * DEPENDENCIES:
 * - Constants.gs (DIALOG_SIZES, UI_THEME)
 * - UIService.gs (getCommonStyles)
 *
 * @fileoverview Search dialog UI components
 * @version 4.51.0
 */

// ============================================================================
// SEARCH DIALOG LAUNCHERS
// ============================================================================

/**
 * Shows desktop search dialog with full interface
 */
function showDesktopSearch() {
  showDialog_(getDesktopSearchHtml(), '🔍 Dashboard Search', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

/**
 * Shows quick search dialog (minimal interface)
 */
function showQuickSearch() {
  showDialog_(getQuickSearchHtml(), 'Quick Search', DIALOG_SIZES.SMALL.width, DIALOG_SIZES.SMALL.height);
}

/**
 * Shows advanced search with complex filtering options
 */
function showAdvancedSearch() {
  showDialog_(getAdvancedSearchHtml(), '🔍 Advanced Search', DIALOG_SIZES.FULLSCREEN.width, DIALOG_SIZES.FULLSCREEN.height);
}

// ============================================================================
// SEARCH DIALOG HTML GENERATORS
// ============================================================================

/**
 * Generates HTML for desktop search dialog
 * @return {string} HTML content
 */
function getDesktopSearchHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .search-container {
          padding: 20px;
        }
        .search-input-wrapper {
          display: flex;
          gap: 10px;
          margin-bottom: 20px;
        }
        .search-input {
          flex: 1;
          padding: 12px 16px;
          font-size: 16px;
          border: 2px solid ${UI_THEME.BORDER_COLOR};
          border-radius: 8px;
          outline: none;
          transition: border-color 0.2s;
        }
        .search-input:focus {
          border-color: ${UI_THEME.PRIMARY_COLOR};
        }
        .search-tabs {
          display: flex;
          gap: 4px;
          margin-bottom: 20px;
          border-bottom: 2px solid ${UI_THEME.BORDER_COLOR};
        }
        .search-tab {
          padding: 10px 20px;
          background: none;
          border: none;
          font-size: 14px;
          color: ${UI_THEME.TEXT_SECONDARY};
          cursor: pointer;
          border-bottom: 2px solid transparent;
          margin-bottom: -2px;
        }
        .search-tab.active {
          color: ${UI_THEME.PRIMARY_COLOR};
          border-bottom-color: ${UI_THEME.PRIMARY_COLOR};
        }
        .results-container {
          max-height: 350px;
          overflow-y: auto;
        }
        .result-item {
          padding: 12px;
          border: 1px solid ${UI_THEME.BORDER_COLOR};
          border-radius: 8px;
          margin-bottom: 8px;
          cursor: pointer;
          transition: background 0.2s;
        }
        .result-item:hover {
          background: #f8f9fa;
        }
        .result-title {
          font-weight: 600;
          color: ${UI_THEME.TEXT_PRIMARY};
        }
        .result-subtitle {
          font-size: 13px;
          color: ${UI_THEME.TEXT_SECONDARY};
          margin-top: 4px;
        }
        .no-results {
          text-align: center;
          padding: 40px;
          color: ${UI_THEME.TEXT_SECONDARY};
        }
        .filter-row {
          display: flex;
          gap: 10px;
          margin-bottom: 15px;
        }
        .filter-select {
          padding: 8px 12px;
          border: 1px solid ${UI_THEME.BORDER_COLOR};
          border-radius: 6px;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <div class="search-container">
        <div class="search-input-wrapper">
          <input type="text" class="search-input" id="searchInput"
                 placeholder="Search members, grievances, or keywords..."
                 autofocus onkeyup="handleSearch(event)">
          <button class="btn btn-primary" onclick="performSearch()">Search</button>
        </div>

        <div class="search-tabs">
          <button class="search-tab active" data-tab="all" onclick="switchTab('all')">All</button>
          <button class="search-tab" data-tab="members" onclick="switchTab('members')">Members</button>
          <button class="search-tab" data-tab="grievances" onclick="switchTab('grievances')">Grievances</button>
        </div>

        <div class="filter-row" id="filterRow">
          <select class="filter-select" id="statusFilter">
            <option value="">All Statuses</option>
            <option value="open">Open</option>
            <option value="pending">Pending</option>
            <option value="resolved">Resolved</option>
            <option value="closed">Closed</option>
          </select>
          <select class="filter-select" id="departmentFilter">
            <option value="">All Departments</option>
          </select>
          <select class="filter-select" id="dateFilter">
            <option value="">Any Time</option>
            <option value="7">Last 7 Days</option>
            <option value="30">Last 30 Days</option>
            <option value="90">Last 90 Days</option>
            <option value="365">Last Year</option>
          </select>
        </div>

        <div class="results-container" id="resultsContainer">
          <div class="no-results">
            <p>Enter a search term to find members or grievances</p>
          </div>
        </div>
      </div>

      <script>
        let currentTab = 'all';

        function switchTab(tab) {
          currentTab = tab;
          document.querySelectorAll('.search-tab').forEach(t => t.classList.remove('active'));
          document.querySelector('[data-tab="' + tab + '"]').classList.add('active');
          performSearch();
        }

        function handleSearch(event) {
          if (event.key === 'Enter') {
            performSearch();
          }
        }

        function performSearch() {
          const query = document.getElementById('searchInput').value.trim();
          const status = document.getElementById('statusFilter').value;
          const department = document.getElementById('departmentFilter').value;
          const dateRange = document.getElementById('dateFilter').value;

          if (query.length < 2) {
            document.getElementById('resultsContainer').innerHTML =
              '<div class="no-results"><p>Enter at least 2 characters to search</p></div>';
            return;
          }

          document.getElementById('resultsContainer').innerHTML =
            '<div class="no-results"><p>Searching...</p></div>';

          google.script.run
            .withSuccessHandler(displayResults)
            .withFailureHandler(handleError)
            .searchDashboard(query, currentTab, { status, department, dateRange });
        }

        function displayResults(results) {
          const container = document.getElementById('resultsContainer');

          if (!results || results.length === 0) {
            container.innerHTML = '<div class="no-results"><p>No results found</p></div>';
            return;
          }

          let html = '';
          results.forEach(function(item) {
            html += '<div class="result-item" data-id="' + escapeHtml(item.id) + '" data-type="' + escapeHtml(item.type) + '" onclick="selectResult(this.dataset.id, this.dataset.type)">';
            html += '<div class="result-title">' + escapeHtml(item.title) + '</div>';
            html += '<div class="result-subtitle">' + escapeHtml(item.subtitle) + '</div>';
            html += '</div>';
          });

          container.innerHTML = html;
        }

        function selectResult(id, type) {
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .navigateToRecord(id, type);
        }

        function handleError(error) {
          document.getElementById('resultsContainer').innerHTML =
            '<div class="no-results"><p>Error: ' + escapeHtml(error.message) + '</p></div>';
        }

        ${getClientSideEscapeHtml()}

        // Load departments on init
        google.script.run
          .withSuccessHandler(function(depts) {
            const select = document.getElementById('departmentFilter');
            depts.forEach(function(d) {
              const opt = document.createElement('option');
              opt.value = d;
              opt.textContent = d;
              select.appendChild(opt);
            });
          })
          .getDepartmentList();
      </script>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for quick search dialog
 * @return {string} HTML content
 */
function getQuickSearchHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .quick-search {
          padding: 20px;
        }
        .quick-input {
          width: 100%;
          padding: 14px;
          font-size: 18px;
          border: 2px solid ${UI_THEME.PRIMARY_COLOR};
          border-radius: 8px;
          outline: none;
          box-sizing: border-box;
        }
        .quick-results {
          margin-top: 15px;
          max-height: 180px;
          overflow-y: auto;
        }
        .quick-item {
          padding: 10px;
          border-radius: 4px;
          cursor: pointer;
        }
        .quick-item:hover {
          background: #e8f0fe;
        }
        .quick-hint {
          font-size: 12px;
          color: ${UI_THEME.TEXT_SECONDARY};
          margin-top: 10px;
        }
      </style>
    </head>
    <body>
      <div class="quick-search">
        <input type="text" class="quick-input" id="quickInput"
               placeholder="Type to search..." autofocus
               oninput="quickSearch(this.value)">
        <div class="quick-results" id="quickResults"></div>
        <div class="quick-hint">Press Enter to select first result, Esc to close</div>
      </div>
      <script>
        ${getClientSideEscapeHtml()}
        let debounceTimer;
        let results = [];

        function quickSearch(query) {
          clearTimeout(debounceTimer);
          if (query.length < 2) {
            document.getElementById('quickResults').innerHTML = '';
            return;
          }
          debounceTimer = setTimeout(function() {
            google.script.run
              .withSuccessHandler(showQuickResults)
              .quickSearchDashboard(query);
          }, 200);
        }

        function showQuickResults(data) {
          results = data || [];
          const container = document.getElementById('quickResults');
          if (results.length === 0) {
            container.innerHTML = '<div class="quick-item">No results</div>';
            return;
          }
          container.innerHTML = results.map(function(r, i) {
            return '<div class="quick-item" onclick="selectQuick(' + i + ')">' +
                   escapeHtml(r.title) + ' <small style="color:#666">(' + escapeHtml(r.type) + ')</small></div>';
          }).join('');
        }

        function selectQuick(index) {
          if (results[index]) {
            google.script.run.navigateToRecord(results[index].id, results[index].type);
            google.script.host.close();
          }
        }

        document.getElementById('quickInput').addEventListener('keydown', function(e) {
          if (e.key === 'Enter' && results.length > 0) {
            selectQuick(0);
          } else if (e.key === 'Escape') {
            google.script.host.close();
          }
        });
      </script>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for advanced search
 * @return {string} HTML content
 */
function getAdvancedSearchHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .advanced-container {
          display: grid;
          grid-template-columns: 280px 1fr;
          height: 100%;
        }
        .filter-panel {
          background: #f8f9fa;
          padding: 20px;
          border-right: 1px solid ${UI_THEME.BORDER_COLOR};
          overflow-y: auto;
        }
        .results-panel {
          padding: 20px;
          overflow-y: auto;
        }
        .filter-section {
          margin-bottom: 20px;
        }
        .filter-title {
          font-weight: 600;
          margin-bottom: 10px;
          color: ${UI_THEME.TEXT_PRIMARY};
        }
        .filter-group {
          margin-bottom: 12px;
        }
        .filter-label {
          display: block;
          font-size: 13px;
          color: ${UI_THEME.TEXT_SECONDARY};
          margin-bottom: 4px;
        }
        .filter-input {
          width: 100%;
          padding: 8px;
          border: 1px solid ${UI_THEME.BORDER_COLOR};
          border-radius: 4px;
          font-size: 14px;
          box-sizing: border-box;
        }
        .checkbox-group {
          display: flex;
          flex-direction: column;
          gap: 8px;
        }
        .checkbox-item {
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .results-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 15px;
        }
        .results-count {
          color: ${UI_THEME.TEXT_SECONDARY};
        }
        .results-table {
          width: 100%;
          border-collapse: collapse;
        }
        .results-table th,
        .results-table td {
          padding: 10px;
          text-align: left;
          border-bottom: 1px solid ${UI_THEME.BORDER_COLOR};
        }
        .results-table th {
          background: #f8f9fa;
          font-weight: 600;
        }
        .results-table tr:hover td {
          background: #f8f9fa;
        }
        .action-buttons {
          margin-top: 20px;
          display: flex;
          gap: 10px;
        }
      </style>
    </head>
    <body>
      <div class="advanced-container">
        <div class="filter-panel">
          <div class="filter-section">
            <div class="filter-title">Search Type</div>
            <div class="checkbox-group">
              <label class="checkbox-item">
                <input type="checkbox" id="searchMembers" checked> Members
              </label>
              <label class="checkbox-item">
                <input type="checkbox" id="searchGrievances" checked> Grievances
              </label>
            </div>
          </div>

          <div class="filter-section">
            <div class="filter-title">Keywords</div>
            <div class="filter-group">
              <input type="text" class="filter-input" id="keywords"
                     placeholder="Enter search terms...">
            </div>
          </div>

          <div class="filter-section">
            <div class="filter-title">Date Range</div>
            <div class="filter-group">
              <label class="filter-label">From</label>
              <input type="date" class="filter-input" id="dateFrom">
            </div>
            <div class="filter-group">
              <label class="filter-label">To</label>
              <input type="date" class="filter-input" id="dateTo">
            </div>
          </div>

          <div class="filter-section">
            <div class="filter-title">Grievance Status</div>
            <div class="checkbox-group" id="statusFilters">
              <label class="checkbox-item">
                <input type="checkbox" value="Open" checked> Open
              </label>
              <label class="checkbox-item">
                <input type="checkbox" value="Pending" checked> Pending
              </label>
              <label class="checkbox-item">
                <input type="checkbox" value="Resolved"> Resolved
              </label>
              <label class="checkbox-item">
                <input type="checkbox" value="Closed"> Closed
              </label>
            </div>
          </div>

          <div class="filter-section">
            <div class="filter-title">Department</div>
            <div class="filter-group">
              <select class="filter-input" id="departmentFilter">
                <option value="">All Departments</option>
              </select>
            </div>
          </div>

          <div class="action-buttons">
            <button class="btn btn-primary" onclick="runAdvancedSearch()">Search</button>
            <button class="btn btn-secondary" onclick="resetFilters()">Reset</button>
          </div>
        </div>

        <div class="results-panel">
          <div class="results-header">
            <h3>Results</h3>
            <span class="results-count" id="resultsCount">0 results</span>
          </div>
          <table class="results-table">
            <thead>
              <tr>
                <th>Type</th>
                <th>Name/ID</th>
                <th>Details</th>
                <th>Status</th>
                <th>Date</th>
              </tr>
            </thead>
            <tbody id="resultsBody">
              <tr>
                <td colspan="5" style="text-align: center; color: #666; padding: 40px;">
                  Configure filters and click Search
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
      <script>
        ${getClientSideEscapeHtml()}
        function runAdvancedSearch() {
          const filters = {
            includeMembers: document.getElementById('searchMembers').checked,
            includeGrievances: document.getElementById('searchGrievances').checked,
            keywords: document.getElementById('keywords').value,
            dateFrom: document.getElementById('dateFrom').value,
            dateTo: document.getElementById('dateTo').value,
            statuses: Array.from(document.querySelectorAll('#statusFilters input:checked'))
                          .map(cb => cb.value),
            department: document.getElementById('departmentFilter').value
          };

          document.getElementById('resultsBody').innerHTML =
            '<tr><td colspan="5" style="text-align:center">Searching...</td></tr>';

          google.script.run
            .withSuccessHandler(displayAdvancedResults)
            .withFailureHandler(function(e) {
              document.getElementById('resultsBody').innerHTML =
                '<tr><td colspan="5" style="text-align:center;color:red">Error: ' + escapeHtml(e.message) + '</td></tr>';
            })
            .advancedSearch(filters);
        }

        function displayAdvancedResults(results) {
          document.getElementById('resultsCount').textContent = results.length + ' results';
          const tbody = document.getElementById('resultsBody');

          if (results.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center">No results found</td></tr>';
            return;
          }

          tbody.innerHTML = results.map(function(r) {
            return '<tr data-id="' + escapeHtml(r.id) + '" data-type="' + escapeHtml(r.type) + '" onclick="selectResult(this.dataset.id, this.dataset.type)" style="cursor:pointer">' +
                   '<td>' + escapeHtml(r.type) + '</td>' +
                   '<td>' + escapeHtml(r.name) + '</td>' +
                   '<td>' + escapeHtml(r.details) + '</td>' +
                   '<td>' + escapeHtml(r.status) + '</td>' +
                   '<td>' + escapeHtml(r.date) + '</td>' +
                   '</tr>';
          }).join('');
        }

        function selectResult(id, type) {
          google.script.run.navigateToRecord(id, type);
          google.script.host.close();
        }

        function resetFilters() {
          document.getElementById('keywords').value = '';
          document.getElementById('dateFrom').value = '';
          document.getElementById('dateTo').value = '';
          document.getElementById('departmentFilter').value = '';
          document.querySelectorAll('#statusFilters input').forEach(function(cb) {
            cb.checked = cb.value === 'Open' || cb.value === 'Pending';
          });
        }

        // Load departments
        google.script.run
          .withSuccessHandler(function(depts) {
            const select = document.getElementById('departmentFilter');
            depts.forEach(function(d) {
              const opt = document.createElement('option');
              opt.value = d;
              opt.textContent = d;
              select.appendChild(opt);
            });
          })
          .getDepartmentList();
      </script>
    </body>
    </html>
  `;
}

// ============================================================================
// BULK ACTIONS DIALOG (PHASE2 Feature 4)
// ============================================================================

/**
 * Shows the Bulk Actions dialog for selected grievance rows.
 * Reads QUICK_ACTIONS checkboxes and presents action options.
 */
function showBulkActionsDialog() {
  var selected = getSelectedGrievanceRows();
  var count = selected.length;
  showDialog_(getBulkActionsDialogHtml_(count, selected), 'Bulk Actions — ' + count + ' Selected', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.MEDIUM.height);
}

/**
 * Generates HTML for the Bulk Actions dialog.
 * @param {number} selectedCount - Number of currently selected rows
 * @param {number[]} selectedRows - Array of selected row numbers
 * @returns {string} HTML content
 * @private
 */
function getBulkActionsDialogHtml_(selectedCount, selectedRows) {
  var countStr = escapeHtml(String(selectedCount));
  var rowsJson = JSON.stringify(selectedRows || []);

  var actionCards = selectedCount === 0
    ? `<div class="no-selection"><div style="font-size:48px">&#x1F4CB;</div>
       <p>No rows are currently selected.<br>Use the checkbox column in the Grievance Log,<br>or use the menu: Select All Open Cases.</p></div>`
    : `<div id="actionCards">
       <div class="action-card" onclick="doFlag()">
         <div class="icon">&#x1F6A9;</div><div class="info"><div class="title">Flag Selected</div>
         <div class="desc">Set Message Alert to TRUE for all selected grievances</div></div></div>
       <div class="action-card" onclick="showEmailForm()">
         <div class="icon">&#x1F4E7;</div><div class="info"><div class="title">Email Selected Members</div>
         <div class="desc">Send an email to members linked to selected grievances</div></div></div>
       <div class="action-card" onclick="doExport()">
         <div class="icon">&#x1F4E4;</div><div class="info"><div class="title">Export Selected as CSV</div>
         <div class="desc">Export selected rows and email CSV to yourself</div></div></div>
       </div>
       <div class="email-form" id="emailForm">
         <label for="emailSubject">Subject</label>
         <input type="text" id="emailSubject" placeholder="Enter email subject...">
         <label for="emailBody">Message</label>
         <textarea id="emailBody" placeholder="Enter email body..."></textarea>
         <div class="btn-row">
           <button class="btn btn-secondary" onclick="hideEmailForm()">Cancel</button>
           <button class="btn btn-primary" onclick="doEmail()">Send Emails</button>
         </div>
       </div>`;

  return `<!DOCTYPE html><html><head><base target="_top">
    ${getMobileOptimizedHead()}
    <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; background: #f9fafb; }
    .container { padding: 24px; max-width: 560px; margin: 0 auto; }
    .count-badge { display: inline-flex; align-items: center; gap: 8px; background: linear-gradient(135deg, #7C3AED, #5B21B6);
      color: white; padding: 10px 20px; border-radius: 50px; font-size: 15px; font-weight: 600; margin-bottom: 20px; }
    .count-badge .num { font-size: 22px; }
    .no-selection { text-align: center; padding: 40px 20px; color: #6B7280; }
    .no-selection p { margin-top: 12px; font-size: 14px; }
    .action-card { display: flex; align-items: center; gap: 16px; background: white; border: 1px solid #E5E7EB;
      border-radius: 12px; padding: 16px 20px; margin-bottom: 12px; cursor: pointer; transition: all 0.2s; }
    .action-card:hover { border-color: #7C3AED; box-shadow: 0 2px 8px rgba(124,58,237,0.15); transform: translateY(-1px); }
    .action-card .icon { font-size: 28px; flex-shrink: 0; }
    .action-card .info { flex: 1; }
    .action-card .title { font-size: 15px; font-weight: 600; color: #1F2937; }
    .action-card .desc { font-size: 13px; color: #6B7280; margin-top: 2px; }
    .spinner-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(255,255,255,0.85); z-index: 999; justify-content: center; align-items: center; flex-direction: column; }
    .spinner { width: 40px; height: 40px; border: 4px solid #E5E7EB; border-top: 4px solid #7C3AED;
      border-radius: 50%; animation: spin 0.8s linear infinite; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .spinner-text { margin-top: 12px; color: #6B7280; font-size: 14px; }
    .result-banner { padding: 12px 16px; border-radius: 8px; margin-top: 16px; font-size: 14px; display: none; }
    .result-success { background: #ecfdf5; color: #065f46; border: 1px solid #a7f3d0; }
    .result-error { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
    .email-form { display: none; background: white; border: 1px solid #E5E7EB; border-radius: 12px; padding: 20px; margin-top: 16px; }
    .email-form label { display: block; font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 4px; }
    .email-form input, .email-form textarea { width: 100%; padding: 10px 12px; border: 1px solid #E5E7EB;
      border-radius: 8px; font-size: 14px; font-family: inherit; margin-bottom: 12px; }
    .email-form textarea { min-height: 100px; resize: vertical; }
    .btn-row { display: flex; gap: 10px; justify-content: flex-end; }
    .btn { padding: 10px 20px; border: none; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; }
    .btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }
    .btn-secondary { background: #f3f4f6; color: #374151; }
    .btn:hover { opacity: 0.9; }
    .footer-bar { display: flex; justify-content: flex-end; margin-top: 20px; }
    </style></head><body>
    <div class="container">
    <div class="count-badge"><span class="num">${countStr}</span> row(s) selected</div>
    ${actionCards}
    <div class="result-banner" id="resultBanner"></div>
    <div class="spinner-overlay" id="spinnerOverlay">
      <div class="spinner"></div><div class="spinner-text" id="spinnerText">Processing...</div>
    </div>
    <div class="footer-bar"><button class="btn btn-secondary" onclick="google.script.host.close()">Close</button></div>
    </div>
    <script>
    var selectedRows = ${rowsJson};
    var selectedCount = ${countStr};
    function showSpinner(msg) {
      document.getElementById("spinnerText").textContent = msg || "Processing...";
      document.getElementById("spinnerOverlay").style.display = "flex";
    }
    function hideSpinner() { document.getElementById("spinnerOverlay").style.display = "none"; }
    function showResult(msg, isError) {
      var b = document.getElementById("resultBanner");
      b.textContent = msg;
      b.className = "result-banner " + (isError ? "result-error" : "result-success");
      b.style.display = "block";
    }
    function showEmailForm() { document.getElementById("emailForm").style.display = "block"; }
    function hideEmailForm() { document.getElementById("emailForm").style.display = "none"; }
    function doFlag() {
      if (!confirm("Flag " + selectedCount + " selected grievance(s) with Message Alert?")) return;
      showSpinner("Flagging selected grievances...");
      google.script.run.withSuccessHandler(function(r) {
        hideSpinner();
        if (r && r.success) { showResult(r.message); } else { showResult((r && r.error) || "Unknown error", true); }
      }).withFailureHandler(function(e) { hideSpinner(); showResult(e.message, true); }).bulkFlagGrievances(selectedRows);
    }
    function doEmail() {
      var subj = document.getElementById("emailSubject").value.trim();
      var body = document.getElementById("emailBody").value.trim();
      if (!subj || !body) { alert("Please enter both a subject and message body."); return; }
      if (!confirm("Send email to members of " + selectedCount + " selected grievance(s)?")) return;
      showSpinner("Sending emails...");
      google.script.run.withSuccessHandler(function(r) {
        hideSpinner(); hideEmailForm();
        if (r && r.success) { showResult(r.message); } else { showResult((r && r.error) || "Unknown error", true); }
      }).withFailureHandler(function(e) { hideSpinner(); showResult(e.message, true); }).bulkEmailGrievanceMembers(selectedRows, subj, body);
    }
    function doExport() {
      if (!confirm("Export " + selectedCount + " selected grievance(s) as CSV and email to you?")) return;
      showSpinner("Exporting CSV...");
      google.script.run.withSuccessHandler(function(r) {
        hideSpinner();
        if (r && r.success) { showResult(r.message); } else { showResult((r && r.error) || "Unknown error", true); }
      }).withFailureHandler(function(e) { hideSpinner(); showResult(e.message, true); }).bulkExportGrievancesToCsv(selectedRows);
    }
    </script>
    </body></html>`;
}

// ============================================================================
// MODAL DIALOGS (merged from 09b_TabModals.gs)
// ============================================================================

// ============================================================================
// SHARED MODAL STYLES
// ============================================================================

/**
 * Returns shared CSS for all tab modals.
 * @returns {string} CSS block
 * @private
 */
function getModalStyles_() {
  return '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; color: #1F2937; background: #F9FAFB; }' +
    'h2 { color: #1A2A4A; margin-bottom: 16px; font-size: 18px; }' +
    '.form-group { margin-bottom: 14px; }' +
    'label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 13px; color: #374151; }' +
    'input, select, textarea { width: 100%; padding: 8px 12px; border: 1px solid #D1D5DB; border-radius: 6px; font-size: 14px; font-family: inherit; }' +
    'input:focus, select:focus, textarea:focus { outline: none; border-color: #2C5282; box-shadow: 0 0 0 3px rgba(44,82,130,0.1); }' +
    'textarea { resize: vertical; min-height: 80px; }' +
    '.btn { display: inline-block; padding: 10px 20px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; margin-right: 8px; }' +
    '.btn-primary { background: #1A2A4A; color: white; }' +
    '.btn-primary:hover { background: #2C5282; }' +
    '.btn-secondary { background: #E5E7EB; color: #374151; }' +
    '.btn-secondary:hover { background: #D1D5DB; }' +
    '.btn-success { background: #276749; color: white; }' +
    '.btn-success:hover { background: #059669; }' +
    '.actions { margin-top: 20px; text-align: right; }' +
    '.msg { padding: 10px 14px; border-radius: 6px; margin-bottom: 12px; font-size: 13px; }' +
    '.msg-success { background: #D1FAE5; color: #065F46; }' +
    '.msg-error { background: #FEE2E2; color: #991B1B; }' +
    '.msg-info { background: #DBEAFE; color: #1E40AF; }' +
    '.row { display: flex; gap: 12px; }' +
    '.row > .form-group { flex: 1; }' +
    '.checklist-item { display: flex; align-items: center; padding: 8px 12px; border-bottom: 1px solid #E5E7EB; }' +
    '.checklist-item label { margin: 0; font-weight: normal; cursor: pointer; flex: 1; }' +
    '.checklist-item input[type=checkbox] { width: auto; margin-right: 10px; transform: scale(1.2); }' +
    '.progress-bar { height: 8px; background: #E5E7EB; border-radius: 4px; overflow: hidden; margin: 8px 0; }' +
    '.progress-fill { height: 100%; background: #276749; border-radius: 4px; transition: width 0.3s; }' +
    '.card { background: white; border-radius: 8px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }' +
    '.search-box { position: relative; margin-bottom: 16px; }' +
    '.search-box input { padding-left: 36px; }' +
    '.search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #9CA3AF; }' +
    '.tabs { display: flex; border-bottom: 2px solid #E5E7EB; margin-bottom: 16px; }' +
    '.tab { padding: 8px 16px; cursor: pointer; font-weight: 600; color: #6B7280; border-bottom: 2px solid transparent; margin-bottom: -2px; }' +
    '.tab.active { color: #1A2A4A; border-bottom-color: #D4A017; }' +
    '.tab-content { display: none; }' +
    '.tab-content.active { display: block; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }' +
    '.spinner { display: none; }' +
    '.spinner.show { display: inline-block; width: 16px; height: 16px; border: 2px solid #E5E7EB; border-top-color: #1A2A4A; border-radius: 50%; animation: spin 0.6s linear infinite; margin-right: 8px; }' +
    '@keyframes spin { to { transform: rotate(360deg); } }' +
    '</style>';
}

/**
 * Returns the base target and meta tags for modal HTML.
 * @returns {string}
 * @private
 */
function getModalHead_() {
  return '<base target="_top"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">';
}

// ============================================================================
// MODAL 1: SUBMIT FEEDBACK
// ============================================================================

// ============================================================================
// MODAL 2: ADD EVENT
// ============================================================================

/**
 * Shows the Add Event modal dialog.
 * Menu: Tools > Calendar & Meetings > Add New Event
 */
function showAddEventModal() {
  showDialog_(getAddEventHtml_(), '📅 Add New Event', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.LARGE.height);
}

function getAddEventHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Create a New Event</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Title</label><input id="title" placeholder="Event title"></div>' +
    '<div class="row"><div class="form-group"><label>Type</label>' +
    '<select id="type"><option>Meeting</option><option>Negotiation</option><option>Training</option><option>Social</option><option>Community</option></select></div>' +
    '<div class="form-group"><label>Date & Time</label><input type="datetime-local" id="dateTime"></div></div>' +
    '<div class="row"><div class="form-group"><label>End Time (optional)</label><input type="datetime-local" id="endTime"></div>' +
    '<div class="form-group"><label>Location</label><input id="location" placeholder="Room, address, or virtual"></div></div>' +
    '<div class="form-group"><label>Description</label><textarea id="description" rows="3" placeholder="Event details..."></textarea></div>' +
    '<div class="form-group"><label>Zoom/Video Link (optional)</label><input id="zoomLink" placeholder="https://..."></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitEvent()">Create Event</button></div>' +
    '<script>' +
    'function submitEvent() {' +
    '  var title = document.getElementById("title").value;' +
    '  var dt = document.getElementById("dateTime").value;' +
    '  if (!title || !dt) { showMsg("Title and Date/Time are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  document.getElementById("submitBtn").textContent = "Creating...";' +
    '  var data = { title: title, type: document.getElementById("type").value, dateTime: dt,' +
    '    endTime: document.getElementById("endTime").value, location: document.getElementById("location").value,' +
    '    description: document.getElementById("description").value, zoomLink: document.getElementById("zoomLink").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Event created! ID: " + r.id, "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Create Event"; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Create Event"; })' +
    '  .modalAddEvent(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Server-side handler for event creation.
 * @param {Object} data - Event data
 * @returns {{ success: boolean, id?: string, error?: string }}
 */
function modalAddEvent(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.PORTAL_EVENTS) || ss.getSheetByName('Events');
    if (!sheet) return { success: false, error: 'Events sheet not found' };

    var id = 'EVT-' + Date.now();
    var user = Session.getActiveUser().getEmail() || 'Unknown';

    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(data.title),
      escapeForFormula(data.type),
      escapeForFormula(data.dateTime),
      escapeForFormula(data.endTime || ''),
      escapeForFormula(data.location || ''),
      escapeForFormula(data.description || ''),
      escapeForFormula(data.zoomLink || ''),
      escapeForFormula(user),
      new Date()
    ]);

    return { success: true, id: id };
  } catch (e) {
    log_('modalAddEvent', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 3: ADD MEETING MINUTES
// ============================================================================

function showAddMinutesModal() {
  showDialog_(getAddMinutesHtml_(), '📝 Add Meeting Minutes', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

function getAddMinutesHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Add Meeting Minutes</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>Meeting Date</label><input type="date" id="meetingDate"></div>' +
    '<div class="form-group"><label>Title</label><input id="title" placeholder="Monthly General Meeting"></div></div>' +
    '<div class="form-group"><label>Bullet Points (one per line)</label><textarea id="bullets" rows="5" placeholder="Motion to approve budget passed unanimously\nCommittee reports presented\nNext meeting scheduled for..."></textarea></div>' +
    '<div class="form-group"><label>Full Minutes</label><textarea id="fullMinutes" rows="6" placeholder="Detailed meeting minutes..."></textarea></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitMinutes()">Save Minutes</button></div>' +
    '<script>' +
    'function submitMinutes() {' +
    '  var date = document.getElementById("meetingDate").value;' +
    '  var title = document.getElementById("title").value;' +
    '  if (!date || !title) { showMsg("Date and Title are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { meetingDate: date, title: title,' +
    '    bullets: document.getElementById("bullets").value, fullMinutes: document.getElementById("fullMinutes").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Minutes saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalAddMinutes(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalAddMinutes(data) {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch (_e) { /* headless */ }
    if (!email) return { success: false, error: 'Could not determine user email.' };
    return DataService.addMeetingMinutes(email, {
      title: data.title,
      meetingDate: data.meetingDate || data.date,
      bullets: data.bullets,
      fullMinutes: data.fullMinutes
    });
  } catch (e) {
    log_('modalAddMinutes', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 4: LOG VOLUNTEER HOURS
// ============================================================================

function showLogVolunteerHoursModal() {
  showDialog_(getLogVolunteerHoursHtml_(), '🤝 Log Volunteer Hours', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.MEDIUM.height);
}

function getLogVolunteerHoursHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Log Volunteer Hours</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>Member ID</label><input id="memberId" placeholder="MBR-001"></div>' +
    '<div class="form-group"><label>Member Name</label><input id="memberName" readonly placeholder="Auto-populated"></div></div>' +
    '<div class="row"><div class="form-group"><label>Activity Date</label><input type="date" id="activityDate"></div>' +
    '<div class="form-group"><label>Activity Type</label>' +
    '<select id="activityType"><option>Meeting</option><option>Outreach</option><option>Event</option><option>Training</option><option>Admin</option><option>Other</option></select></div></div>' +
    '<div class="row"><div class="form-group"><label>Hours</label><input type="number" id="hours" min="0.25" step="0.25" placeholder="2"></div>' +
    '<div class="form-group"><label>Verified By (optional)</label><input id="verifiedBy" placeholder="Steward name"></div></div>' +
    '<div class="form-group"><label>Description</label><textarea id="description" rows="3" placeholder="What was done..."></textarea></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitHours()">Log Hours</button></div>' +
    '<script>' +
    'document.getElementById("memberId").addEventListener("blur", function() {' +
    '  var mid = this.value.trim();' +
    '  if (mid) { google.script.run.withSuccessHandler(function(name) {' +
    '    document.getElementById("memberName").value = name || "Not found";' +
    '  }).modalLookupMemberName(mid); }' +
    '});' +
    'function submitHours() {' +
    '  var mid = document.getElementById("memberId").value;' +
    '  var date = document.getElementById("activityDate").value;' +
    '  var hours = document.getElementById("hours").value;' +
    '  if (!mid || !date || !hours || parseFloat(hours) <= 0) { showMsg("Member ID, Date, and Hours are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { memberId: mid, memberName: document.getElementById("memberName").value,' +
    '    activityDate: date, activityType: document.getElementById("activityType").value,' +
    '    hours: parseFloat(hours), description: document.getElementById("description").value,' +
    '    verifiedBy: document.getElementById("verifiedBy").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Hours logged!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalLogVolunteerHours(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Looks up a member name by ID from the Member Directory.
 * @param {string} memberId
 * @returns {string} Member name or empty string
 */
function modalLookupMemberName(memberId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return '';
    var data = sheet.getDataRange().getValues();
    var idCol    = MEMBER_COLS.MEMBER_ID - 1;
    var fnCol    = MEMBER_COLS.FIRST_NAME - 1;
    var lnCol    = MEMBER_COLS.LAST_NAME - 1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idCol]).trim() === String(memberId).trim()) {
        return String(data[r][fnCol] || '') + ' ' + String(data[r][lnCol] || '');
      }
    }
    return '';
  } catch (_e) {
    return '';
  }
}

function modalLogVolunteerHours(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.VOLUNTEER_HOURS);
    if (!sheet) return { success: false, error: 'Volunteer Hours sheet not found' };

    var id = 'VH-' + Date.now();
    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(data.memberId),
      escapeForFormula(data.memberName || ''),
      escapeForFormula(data.activityDate),
      escapeForFormula(data.activityType),
      data.hours,
      escapeForFormula(data.description || ''),
      escapeForFormula(data.verifiedBy || ''),
      ''
    ]);

    return { success: true, id: id };
  } catch (e) {
    log_('modalLogVolunteerHours', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 5: ADD NON-MEMBER CONTACT
// ============================================================================

function showAddContactModal() {
  showDialog_(getAddContactHtml_(), '👥 Add Non-Member Contact', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.MEDIUM.height);
}

function getAddContactHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Add Non-Member Contact</h2>' +
    '<div id="msg"></div>' +
    '<div class="row"><div class="form-group"><label>First Name</label><input id="firstName"></div>' +
    '<div class="form-group"><label>Last Name</label><input id="lastName"></div></div>' +
    '<div class="row"><div class="form-group"><label>Job Title</label><input id="jobTitle" placeholder="Supervisor, Manager..."></div>' +
    '<div class="form-group"><label>Work Location</label><input id="workLocation"></div></div>' +
    '<div class="row"><div class="form-group"><label>Unit</label><input id="unit"></div>' +
    '<div class="form-group"><label>Union Name</label><input id="unionName"></div></div>' +
    '<div class="row"><div class="form-group"><label>Shirt Size</label><select id="shirtSize"><option value="">—</option><option>XS</option><option>S</option><option>M</option><option>L</option><option>XL</option><option>2XL</option><option>3XL</option><option>4XL</option></select></div>' +
    '<div class="form-group"><label>Steward</label><select id="isSteward"><option value="No">No</option><option value="Yes">Yes</option></select></div></div>' +
    '<div class="row"><div class="form-group"><label>Email</label><input type="email" id="email"></div>' +
    '<div class="form-group"><label>Phone</label><input id="phone" placeholder="555-123-4567"></div></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitContact()">Add Contact</button></div>' +
    '<script>' +
    'function submitContact() {' +
    '  var fn = document.getElementById("firstName").value;' +
    '  var ln = document.getElementById("lastName").value;' +
    '  if (!fn || !ln) { showMsg("First and Last Name are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  var data = { firstName: fn, lastName: ln, jobTitle: document.getElementById("jobTitle").value,' +
    '    workLocation: document.getElementById("workLocation").value, unit: document.getElementById("unit").value,' +
    '    unionName: document.getElementById("unionName").value, shirtSize: document.getElementById("shirtSize").value,' +
    '    isSteward: document.getElementById("isSteward").value,' +
    '    email: document.getElementById("email").value, phone: document.getElementById("phone").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Contact added!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalAddContact(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalAddContact(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS) || ss.getSheetByName('Non member contacts');
    if (!sheet) return { success: false, error: 'Non-Member Contacts sheet not found' };

    // Duplicate check: same first+last name at same location. Use NMC_COLS so
    // a header reorder can't silently compare the wrong columns.
    var existing = sheet.getDataRange().getValues();
    for (var r = 1; r < existing.length; r++) {
      if (String(col_(existing[r], NMC_COLS.FIRST_NAME) || '').trim().toLowerCase() === String(data.firstName || '').trim().toLowerCase() &&
          String(col_(existing[r], NMC_COLS.LAST_NAME) || '').trim().toLowerCase() === String(data.lastName || '').trim().toLowerCase() &&
          String(col_(existing[r], NMC_COLS.WORK_LOCATION) || '').trim().toLowerCase() === String(data.workLocation || '').trim().toLowerCase()) {
        return { success: false, error: 'Duplicate: ' + data.firstName + ' ' + data.lastName + ' at ' + data.workLocation + ' already exists.' };
      }
    }

    // Delegate to DataService.addNonMemberContact which generates the CONTACT_ID.
    // The prior hand-rolled append omitted CONTACT_ID entirely, creating orphan
    // rows that update/delete APIs could never find (they look up by ID).
    if (typeof DataService !== 'undefined' && typeof DataService.addNonMemberContact === 'function') {
      return DataService.addNonMemberContact(data);
    }

    return { success: false, error: 'Contact service unavailable.' };
  } catch (e) {
    log_('modalAddContact', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 6: TAKE ATTENDANCE
// ============================================================================

function showTakeAttendanceModal() {
  showDialog_(getTakeAttendanceHtml_(), '📅 Take Meeting Attendance', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

function getTakeAttendanceHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.member-list { max-height: 340px; overflow-y: auto; border: 1px solid #E5E7EB; border-radius: 6px; } .count { font-size: 14px; color: #6B7280; margin-top: 8px; }</style>' +
    '</head><body>' +
    '<h2>Take Meeting Attendance</h2>' +
    '<div id="msg"></div>' +
    '<div class="card">' +
    '<div class="row"><div class="form-group"><label>Meeting Date</label><input type="date" id="meetingDate"></div>' +
    '<div class="form-group"><label>Meeting Type</label>' +
    '<select id="meetingType"><option>Regular</option><option>Special</option><option>Committee</option><option>Emergency</option></select></div></div>' +
    '<div class="form-group"><label>Meeting Name</label><input id="meetingName" placeholder="March General Meeting"></div>' +
    '</div>' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin:12px 0"><h3>Members</h3>' +
    '<label style="font-weight:normal"><input type="checkbox" id="selectAll" onchange="toggleAll(this.checked)"> Select All</label></div>' +
    '<div class="member-list" id="memberList"><div class="msg msg-info">Loading members...</div></div>' +
    '<div class="count" id="count"></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-success" id="submitBtn" onclick="saveAttendance()">Save Attendance</button></div>' +
    '<script>' +
    'var members = [];' +
    'google.script.run.withSuccessHandler(function(list) {' +
    '  members = list || [];' +
    '  var html = "";' +
    '  for (var i = 0; i < members.length; i++) {' +
    '    html += "<div class=\\"checklist-item\\"><input type=\\"checkbox\\" id=\\"m" + i + "\\" onchange=\\"updateCount()\\"><label for=\\"m" + i + "\\">" + members[i].name + " <span style=\\"color:#9CA3AF;font-size:12px\\">(" + members[i].id + ")</span></label></div>";' +
    '  }' +
    '  document.getElementById("memberList").innerHTML = html || "<div class=\\"msg msg-info\\">No members found in directory.</div>";' +
    '  updateCount();' +
    '}).modalGetMemberList();' +
    'function toggleAll(checked) { for (var i = 0; i < members.length; i++) { document.getElementById("m"+i).checked = checked; } updateCount(); }' +
    'function updateCount() { var c = 0; for (var i = 0; i < members.length; i++) { if (document.getElementById("m"+i).checked) c++; }' +
    '  document.getElementById("count").textContent = c + " / " + members.length + " members present"; }' +
    'function saveAttendance() {' +
    '  var date = document.getElementById("meetingDate").value;' +
    '  var name = document.getElementById("meetingName").value;' +
    '  if (!date || !name) { showMsg("Date and Meeting Name required.", "error"); return; }' +
    '  var attended = [];' +
    '  for (var i = 0; i < members.length; i++) { attended.push(document.getElementById("m"+i).checked); }' +
    '  document.getElementById("submitBtn").disabled = true; document.getElementById("submitBtn").textContent = "Saving...";' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg(r.count + " attendance records saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Save Attendance"; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; document.getElementById("submitBtn").textContent = "Save Attendance"; })' +
    '  .modalSaveAttendance({ date: date, name: name, type: document.getElementById("meetingType").value, attended: attended, members: members });' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

/**
 * Returns a list of active members for the attendance checklist.
 * @returns {Array<{id: string, name: string, email: string}>}
 */
function modalGetMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var idCol    = MEMBER_COLS.MEMBER_ID - 1;
    var fnCol    = MEMBER_COLS.FIRST_NAME - 1;
    var lnCol    = MEMBER_COLS.LAST_NAME - 1;
    var emCol    = MEMBER_COLS.EMAIL - 1;
    var members = [];
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][idCol] || '').trim();
      var name = String(data[r][fnCol] || '').trim();
      if (data[r].length > lnCol) name += ' ' + String(data[r][lnCol] || '').trim();
      if (id && name.trim()) {
        members.push({ id: escapeHtml(id), name: escapeHtml(name.trim()), email: String(data[r][emCol] || '') });
      }
    }
    return members;
  } catch (e) {
    log_('modalGetMemberList', 'Error: ' + e.message);
    return [];
  }
}

/**
 * Saves attendance records — one row per member.
 * @param {Object} data - { date, name, type, attended: boolean[], members: [{id,name}] }
 * @returns {{ success: boolean, count?: number, error?: string }}
 */
function modalSaveAttendance(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_ATTENDANCE);
    if (!sheet) return { success: false, error: 'Meeting Attendance sheet not found' };

    var rows = [];
    var count = 0;

    for (var i = 0; i < data.members.length; i++) {
      var entryId = 'ATT-' + Date.now() + '-' + i;
      rows.push([
        escapeForFormula(entryId),
        escapeForFormula(data.date),
        escapeForFormula(data.type),
        escapeForFormula(data.name),
        escapeForFormula(data.members[i].id),
        escapeForFormula(data.members[i].name),
        data.attended[i] ? true : false,
        ''
      ]);
      if (data.attended[i]) count++;
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: count };
  } catch (e) {
    log_('modalSaveAttendance', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 7: CASE PROGRESS VIEWER
// ============================================================================

function showCaseProgressModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var caseId = '';

  // Try to get Case ID from selected cell
  if (activeSheet.getName() === SHEETS.CASE_CHECKLIST || activeSheet.getName() === SHEETS.GRIEVANCE_LOG) {
    caseId = String(activeSheet.getActiveCell().getValue() || '');
  }

  showDialog_(getCaseProgressHtml_(caseId), '✅ Case Progress', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.LARGE.height);
}

function getCaseProgressHtml_(caseId) {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.item-list { max-height: 380px; overflow-y: auto; } .item-done { text-decoration: line-through; color: #9CA3AF; } .item-required { color: #9B2335; font-weight: 600; }</style>' +
    '</head><body>' +
    '<h2>Case Checklist Progress</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Case ID</label><div class="row"><input id="caseId" value="' + escapeHtml(caseId) + '" placeholder="Enter Case ID">' +
    '<button class="btn btn-primary" style="white-space:nowrap" onclick="loadCase()">Load</button></div></div>' +
    '<div id="progressArea" style="display:none"><div class="card"><strong id="progressText"></strong>' +
    '<div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div></div>' +
    '<div class="item-list" id="items"></div>' +
    '<div class="actions"><button class="btn btn-success" onclick="saveChecks()">Save Changes</button></div></div>' +
    '<script>' +
    'var items = [];' +
    'function loadCase() {' +
    '  var cid = document.getElementById("caseId").value.trim();' +
    '  if (!cid) { showMsg("Enter a Case ID", "error"); return; }' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    items = result || [];' +
    '    if (items.length === 0) { showMsg("No checklist items found for " + cid, "info"); return; }' +
    '    renderItems();' +
    '    document.getElementById("progressArea").style.display = "block";' +
    '  }).modalGetCaseChecklist(cid);' +
    '}' +
    'function renderItems() {' +
    '  var done = 0, html = "";' +
    '  for (var i = 0; i < items.length; i++) {' +
    '    var it = items[i]; if (it.completed) done++;' +
    '    var cls = it.completed ? "item-done" : (it.required ? "item-required" : "");' +
    '    html += "<div class=\\"checklist-item\\"><input type=\\"checkbox\\" id=\\"c"+i+"\\" "+(it.completed?"checked":"")+' +
    '    " onchange=\\"updateProgress()\\"><label for=\\"c"+i+"\\" class=\\""+cls+"\\">"+it.text+' +
    '    (it.required && !it.completed ? " ⚠️" : "")+(it.actionType ? " <span class=\\"badge\\" style=\\"background:#DBEAFE;color:#1E40AF\\">"+it.actionType+"</span>" : "")+' +
    '    "</label></div>";' +
    '  }' +
    '  document.getElementById("items").innerHTML = html;' +
    '  updateProgress();' +
    '}' +
    'function updateProgress() {' +
    '  var done = 0;' +
    '  for (var i = 0; i < items.length; i++) { if (document.getElementById("c"+i).checked) done++; }' +
    '  var pct = Math.round(done / items.length * 100);' +
    '  document.getElementById("progressText").textContent = done + "/" + items.length + " complete (" + pct + "%)";' +
    '  document.getElementById("progressFill").style.width = pct + "%";' +
    '}' +
    'function saveChecks() {' +
    '  var updates = [];' +
    '  for (var i = 0; i < items.length; i++) { updates.push({ row: items[i].row, completed: document.getElementById("c"+i).checked }); }' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) showMsg("Saved!", "success"); else showMsg(r && r.error || "Failed", "error");' +
    '  }).modalSaveCaseChecklist(updates);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    (caseId ? 'loadCase();' : '') +
    '</script></body></html>';
}

function modalGetCaseChecklist(caseId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var items = [];
    for (var r = 1; r < data.length; r++) {
      if (String(col_(data[r], CHECKLIST_COLS.CASE_ID)).trim() === String(caseId).trim()) {
        // Required is stored as "Yes"/"No" by generateChecklistItems (12_Features.gs).
        // The previous check String(...).toUpperCase() === 'Y' was always false for
        // "YES", silently flipping every item to required:false and corrupting writebacks.
        var _requiredRaw = String(col_(data[r], CHECKLIST_COLS.REQUIRED) || '').trim().toUpperCase();
        var _completedRaw = col_(data[r], CHECKLIST_COLS.COMPLETED);
        items.push({
          row: r + 1,
          checklistId: col_(data[r], CHECKLIST_COLS.CHECKLIST_ID),
          caseId: col_(data[r], CHECKLIST_COLS.CASE_ID),
          actionType: escapeHtml(String(col_(data[r], CHECKLIST_COLS.ACTION_TYPE) || '')),
          text: escapeHtml(String(col_(data[r], CHECKLIST_COLS.ITEM_TEXT) || '')),
          category: String(col_(data[r], CHECKLIST_COLS.CATEGORY) || ''),
          required: _requiredRaw === 'YES' || _requiredRaw === 'TRUE' || _requiredRaw === 'Y' || _requiredRaw === '1',
          completed: _completedRaw === true || String(_completedRaw).trim().toUpperCase() === 'TRUE',
          completedBy: col_(data[r], CHECKLIST_COLS.COMPLETED_BY) || '',
          completedDate: col_(data[r], CHECKLIST_COLS.COMPLETED_DATE) || ''
        });
      }
    }
    return items;
  } catch (e) {
    log_('modalGetCaseChecklist', 'Error: ' + e.message);
    return [];
  }
}

function modalSaveCaseChecklist(updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
    if (!sheet) return { success: false, error: 'Case Checklist sheet not found' };

    var user = Session.getActiveUser().getEmail() || 'Unknown';
    var now = new Date();

    // CHECKLIST_COLS is 1-indexed. CHECKLIST_COLS.COMPLETED/COMPLETED_BY/COMPLETED_DATE
    // are consecutive in the canonical header map; fall back to 7/8/9 if the
    // dynamic map hasn't been resolved in this execution context.
    var _completedCol = (typeof CHECKLIST_COLS === 'object' && CHECKLIST_COLS.COMPLETED) || 7;
    for (var i = 0; i < updates.length; i++) {
      var u = updates[i];
      var rowValues = [u.completed];
      if (u.completed) {
        rowValues.push(escapeForFormula(user));  // Completed By
        rowValues.push(now);                     // Completed Date
      } else {
        rowValues.push('');  // Clear Completed By
        rowValues.push('');  // Clear Completed Date
      }
      sheet.getRange(u.row, _completedCol, 1, rowValues.length).setValues([rowValues]);
    }

    return { success: true };
  } catch (e) {
    log_('modalSaveCaseChecklist', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 8: SURVEY RESPONSE VIEWER
// ============================================================================

function showSurveyResponseViewer() {
  showDialog_(getSurveyResponseViewerHtml_(), '📊 Survey Response Viewer', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

function getSurveyResponseViewerHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.response-card { padding: 16px; } .q-row { display:flex; align-items:center; margin-bottom:8px; } ' +
    '.q-label { flex:1; font-size:13px; color:#374151; } .q-value { font-weight:600; font-size:14px; min-width:40px; text-align:right; }' +
    '.bar-bg { flex:2; height:12px; background:#E5E7EB; border-radius:6px; margin:0 12px; overflow:hidden; }' +
    '.bar-fill { height:100%; border-radius:6px; transition:width 0.3s; }' +
    '.nav-btns { display:flex; justify-content:space-between; margin-top:16px; }' +
    '.demo-info { display:flex; gap:16px; flex-wrap:wrap; margin-bottom:16px; padding:12px; background:#EBF4FF; border-radius:8px; }' +
    '.demo-item { font-size:13px; } .demo-item strong { color:#1A2A4A; }</style>' +
    '</head><body>' +
    '<h2>Survey Response Viewer</h2>' +
    '<div id="msg"></div>' +
    '<div id="loading" class="msg msg-info">Loading responses...</div>' +
    '<div id="viewer" style="display:none">' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">' +
    '<span id="counter" style="font-size:13px;color:#6B7280"></span>' +
    '</div>' +
    '<div class="card response-card" id="responseCard"></div>' +
    '<div class="nav-btns"><button class="btn btn-secondary" onclick="prev()">← Previous</button>' +
    '<button class="btn btn-secondary" onclick="next()">Next →</button></div></div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'var responses = [], headers = [], idx = 0;' +
    'google.script.run.withSuccessHandler(function(data) {' +
    '  document.getElementById("loading").style.display = "none";' +
    '  if (!data || data.responses.length === 0) { showMsg("No survey responses found.", "info"); return; }' +
    '  responses = data.responses; headers = data.headers;' +
    '  document.getElementById("viewer").style.display = "block";' +
    '  renderResponse();' +
    '}).modalGetSurveyResponses();' +
    'function renderResponse() {' +
    '  var r = responses[idx]; document.getElementById("counter").textContent = "Response " + (idx+1) + " of " + responses.length;' +
    '  var html = "<div class=\\"demo-info\\">";' +
    '  for (var i = 0; i < Math.min(5, r.length); i++) { html += "<div class=\\"demo-item\\"><strong>" + escapeHtml(headers[i]) + ":</strong> " + escapeHtml(r[i] || "—") + "</div>"; }' +
    '  html += "</div>";' +
    '  for (var j = 5; j < r.length; j++) {' +
    '    var val = parseFloat(r[j]); var label = headers[j] || "Q" + (j+1);' +
    '    if (!isNaN(val) && val >= 1 && val <= 10) {' +
    '      var pct = val * 10; var color = val >= 7 ? "#6EE7B7" : (val >= 4 ? "#FDE68A" : "#FCA5A5");' +
    '      html += "<div class=\\"q-row\\"><span class=\\"q-label\\">" + escapeHtml(label) + "</span><div class=\\"bar-bg\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+color+"\\"></div></div><span class=\\"q-value\\">"+val+"</span></div>";' +
    '    } else if (r[j]) {' +
    '      html += "<div style=\\"margin-bottom:8px\\"><strong style=\\"font-size:13px\\">" + escapeHtml(label) + ":</strong> <span style=\\"font-size:13px\\">" + escapeHtml(r[j]) + "</span></div>";' +
    '    }' +
    '  }' +
    '  document.getElementById("responseCard").innerHTML = html;' +
    '}' +
    'function next() { if (idx < responses.length - 1) { idx++; renderResponse(); } }' +
    'function prev() { if (idx > 0) { idx--; renderResponse(); } }' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalGetSurveyResponses() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!sheet) return { headers: [], responses: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], responses: [] };

    var headers = data[0].map(function(h) {
      // Shorten long question headers to first 40 chars, then escape for safe HTML rendering
      var s = String(h).trim();
      s = s.length > 40 ? s.substring(0, 37) + '...' : s;
      return escapeHtml(s);
    });

    // Escape all response cell values for safe HTML rendering
    var responses = data.slice(1).map(function(row) {
      return row.map(function(cell) { return escapeHtml(String(cell || '')); });
    });

    return { headers: headers, responses: responses };
  } catch (e) {
    log_('modalGetSurveyResponses', 'Error: ' + e.message);
    return { headers: [], responses: [] };
  }
}

// ============================================================================
// MODAL 9: WELCOME / SETUP WIZARD
// ============================================================================

function showWelcomeWizardModal() {
  showDialog_(getWelcomeWizardHtml_(), '🏗️ Welcome — Setup Wizard', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

function getWelcomeWizardHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>' +
    '.wizard-steps { display:flex; gap:4px; margin-bottom:20px; } .step-dot { flex:1; height:6px; border-radius:3px; background:#E5E7EB; } .step-dot.active { background:#D4A017; } .step-dot.done { background:#276749; }' +
    '.step-content { min-height: 250px; } .step-title { font-size: 16px; font-weight: 700; color: #1A2A4A; margin-bottom: 12px; }' +
    '.step-desc { font-size: 14px; color: #4B5563; line-height: 1.6; margin-bottom: 16px; }' +
    '.check-item { display:flex; align-items:center; padding:10px; background:white; border-radius:6px; margin-bottom:8px; border:1px solid #E5E7EB; }' +
    '.check-item input { width:auto; margin-right:12px; transform:scale(1.3); }' +
    '</style></head><body>' +
    '<div class="wizard-steps" id="stepDots"></div>' +
    '<div class="step-content" id="stepContent"></div>' +
    '<div class="actions">' +
    '<button class="btn btn-secondary" id="prevBtn" onclick="prevStep()" style="display:none">← Back</button>' +
    '<button class="btn btn-secondary" id="skipBtn" onclick="google.script.host.close()">Skip for now</button>' +
    '<button class="btn btn-primary" id="nextBtn" onclick="nextStep()">Next →</button>' +
    '</div>' +
    '<script>' +
    'var step = 0, totalSteps = 4;' +
    'var steps = [' +
    '  { title: "Welcome to Your Union Dashboard!", desc: "This wizard will guide you through the essential setup steps. You can always come back to this wizard from the Admin menu.<br><br><strong>What you will configure:</strong><br>1. Organization details<br>2. Steward setup<br>3. Key features<br>4. Final checks" },' +
    '  { title: "Step 1: Organization Setup", desc: "Open the <strong>Config</strong> tab and fill in:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Organization Name</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Local Number</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Time Zone</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Contact Email</div>" },' +
    '  { title: "Step 2: Add Your First Members", desc: "Open <strong>Member Directory</strong> and add at least one member:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Add yourself as the first steward</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Import or manually add members</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Assign steward roles</div>" },' +
    '  { title: "Step 3: Explore Key Features", desc: "Try these essential features:<br><br><div class=\\"check-item\\"><input type=\\"checkbox\\"> Open Union Hub menu — explore Search & Members</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Create a test grievance case</div><div class=\\"check-item\\"><input type=\\"checkbox\\"> Check the Getting Started tab for more guidance</div>" },' +
    '  { title: "Setup Complete! 🎉", desc: "You are ready to start using your Union Dashboard!<br><br><strong>Quick links:</strong><br>• 📊 Union Hub — Daily operations<br>• 🔧 Tools — Calendar, drive, notifications<br>• 🛠️ Admin — System management<br>• ❓ FAQ — Common questions<br><br>You can re-run this wizard anytime from Admin menu." }' +
    '];' +
    'totalSteps = steps.length;' +
    'function render() {' +
    '  var dots = "";' +
    '  for (var i = 0; i < totalSteps; i++) { dots += "<div class=\\"step-dot " + (i < step ? "done" : (i === step ? "active" : "")) + "\\"></div>"; }' +
    '  document.getElementById("stepDots").innerHTML = dots;' +
    '  document.getElementById("stepContent").innerHTML = "<div class=\\"step-title\\">" + steps[step].title + "</div><div class=\\"step-desc\\">" + steps[step].desc + "</div>";' +
    '  document.getElementById("prevBtn").style.display = step > 0 ? "inline-block" : "none";' +
    '  document.getElementById("nextBtn").textContent = step === totalSteps - 1 ? "Finish ✅" : "Next →";' +
    '  document.getElementById("skipBtn").style.display = step === totalSteps - 1 ? "none" : "inline-block";' +
    '}' +
    'function nextStep() { if (step < totalSteps - 1) { step++; render(); } else { google.script.host.close(); } }' +
    'function prevStep() { if (step > 0) { step--; render(); } }' +
    'render();' +
    '</script></body></html>';
}

// ============================================================================
// MODAL 10: SEARCHABLE HELP GUIDE
// ============================================================================

function showSearchableHelpModal() {
  showDialog_(getSearchableHelpHtml_(), '❓ Help & Documentation', DIALOG_SIZES.LARGE.width, DIALOG_SIZES.LARGE.height);
}

function getSearchableHelpHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.result { padding:12px; border-bottom:1px solid #E5E7EB; cursor:pointer; } .result:hover { background:#EBF4FF; } ' +
    '.result-source { font-size:11px; color:#6B7280; } .result-title { font-weight:600; } .result-text { font-size:13px; color:#4B5563; margin-top:4px; } ' +
    '.results-container { max-height:450px; overflow-y:auto; border:1px solid #E5E7EB; border-radius:6px; }</style>' +
    '</head><body>' +
    '<div class="search-box"><span class="search-icon">🔍</span><input id="searchInput" placeholder="Search help articles, FAQ, features..." oninput="doSearch(this.value)"></div>' +
    '<div class="tabs" id="tabs">' +
    '<div class="tab active" onclick="switchTab(\'all\')">All</div>' +
    '<div class="tab" onclick="switchTab(\'faq\')">FAQ</div>' +
    '<div class="tab" onclick="switchTab(\'features\')">Features</div>' +
    '<div class="tab" onclick="switchTab(\'tips\')">Quick Tips</div></div>' +
    '<div class="results-container" id="results"><div class="msg msg-info">Loading help content...</div></div>' +
    '<script>' +
    getClientSideEscapeHtml() +
    'var allItems = [], activeTab = "all";' +
    'google.script.run.withSuccessHandler(function(data) {' +
    '  allItems = data || [];' +
    '  doSearch("");' +
    '}).modalGetHelpContent();' +
    'function switchTab(tab) { activeTab = tab;' +
    '  var tabs = document.querySelectorAll(".tab"); tabs.forEach(function(t) { t.className = "tab"; });' +
    '  event.target.className = "tab active"; doSearch(document.getElementById("searchInput").value); }' +
    'function doSearch(query) {' +
    '  var q = query.toLowerCase().trim();' +
    '  var filtered = allItems.filter(function(item) {' +
    '    if (activeTab !== "all" && item.source !== activeTab) return false;' +
    '    if (!q) return true;' +
    '    return (item.title + " " + item.text).toLowerCase().indexOf(q) !== -1;' +
    '  });' +
    '  var html = "";' +
    '  if (filtered.length === 0) { html = "<div class=\\"msg msg-info\\">No results found.</div>"; }' +
    '  for (var i = 0; i < Math.min(filtered.length, 50); i++) {' +
    '    var it = filtered[i];' +
    '    var badge = it.source === "faq" ? "❓ FAQ" : (it.source === "features" ? "📋 Feature" : "💡 Tip");' +
    '    html += "<div class=\\"result\\"><span class=\\"result-source\\">" + badge + "</span><div class=\\"result-title\\">" + escapeHtml(it.title) + "</div><div class=\\"result-text\\">" + escapeHtml(it.text.length > 150 ? it.text.substring(0, 147) + "..." : it.text) + "</div></div>";' +
    '  }' +
    '  document.getElementById("results").innerHTML = html;' +
    '}' +
    '</script></body></html>';
}

function modalGetHelpContent() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var items = [];

    // Pull from FAQ sheet
    var faqSheet = ss.getSheetByName(SHEETS.FAQ);
    if (faqSheet) {
      var faqData = faqSheet.getDataRange().getValues();
      for (var r = 1; r < faqData.length; r++) {
        var text = String(faqData[r][0] || '').trim();
        if (text && text.indexOf('?') !== -1) {
          var answer = (r + 1 < faqData.length) ? String(faqData[r + 1][0] || '').trim() : '';
          items.push({ source: 'faq', title: escapeHtml(text), text: escapeHtml(answer) });
        }
      }
    }

    // Pull from Features Reference sheet
    var featSheet = ss.getSheetByName(SHEETS.FEATURES_REFERENCE);
    if (featSheet) {
      var featData = featSheet.getDataRange().getValues();
      for (var f = 1; f < featData.length; f++) {
        var name = String(featData[f][1] || '').trim();
        var desc = String(featData[f][2] || '').trim();
        if (name && desc) {
          items.push({ source: 'features', title: escapeHtml(name), text: escapeHtml(desc) });
        }
      }
    }

    // Add quick tips
    var tips = [
      { title: 'Keyboard Shortcut: Quick Search', text: 'Press Ctrl+F (Cmd+F on Mac) to search any sheet.' },
      { title: 'Traffic Light Colors', text: 'Red = overdue, Orange = due soon (1-3 days), Green = on track.' },
      { title: 'Bulk Actions', text: 'Select multiple grievance rows, then use Union Hub > Grievances > Bulk Update.' },
      { title: 'Mobile Access', text: 'The web dashboard works on phones and tablets — share the URL with members.' },
      { title: 'Data Backup', text: 'The system automatically creates audit logs. Check Admin > Security > Audit Log.' }
    ];
    for (var t = 0; t < tips.length; t++) {
      items.push({ source: 'tips', title: tips[t].title, text: tips[t].text });
    }

    return items;
  } catch (e) {
    log_('modalGetHelpContent', 'Error: ' + e.message);
    return [];
  }
}

// ============================================================================
// MODAL 11: QUESTION EDITOR (Survey Questions)
// ============================================================================

function showQuestionEditorModal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  var questionRow = row;

  // Only works from Survey Questions sheet
  if (sheet.getName() !== SHEETS.SURVEY_QUESTIONS) {
    SpreadsheetApp.getUi().alert('Please select a question row in the Survey Questions sheet first.');
    return;
  }

  showDialog_(getQuestionEditorHtml_(questionRow), '📋 Edit Survey Question', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.LARGE.height);
}

function getQuestionEditorHtml_(row) {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() + '</head><body>' +
    '<h2>Edit Survey Question</h2>' +
    '<div id="msg"></div>' +
    '<div id="loading" class="msg msg-info">Loading question...</div>' +
    '<div id="form" style="display:none">' +
    '<div class="row"><div class="form-group"><label>Question ID</label><input id="qId" readonly></div>' +
    '<div class="form-group"><label>Section</label><input id="section"></div></div>' +
    '<div class="form-group"><label>Section Title</label><input id="sectionTitle"></div>' +
    '<div class="form-group"><label>Question Text</label><textarea id="questionText" rows="3"></textarea></div>' +
    '<div class="row"><div class="form-group"><label>Type</label>' +
    '<select id="qType"><option>dropdown</option><option>slider-10</option><option>radio</option><option>paragraph</option></select></div>' +
    '<div class="form-group"><label>Required</label><select id="required"><option>Y</option><option>N</option></select></div>' +
    '<div class="form-group"><label>Active</label><select id="active"><option>Y</option><option>N</option></select></div></div>' +
    '<div class="form-group"><label>Options (pipe-separated: Option A|Option B|Option C)</label><textarea id="options" rows="2"></textarea></div>' +
    '<div class="form-group"><label>Branch Parent (optional)</label><input id="branchParent"></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="saveQuestion()">Save Changes</button></div></div>' +
    '<script>' +
    'var ROW = ' + row + ';' +
    'google.script.run.withSuccessHandler(function(q) {' +
    '  document.getElementById("loading").style.display = "none";' +
    '  if (!q) { showMsg("No question data found at row " + ROW, "error"); return; }' +
    '  document.getElementById("form").style.display = "block";' +
    '  document.getElementById("qId").value = q.id || "";' +
    '  document.getElementById("section").value = q.section || "";' +
    '  document.getElementById("sectionTitle").value = q.sectionTitle || "";' +
    '  document.getElementById("questionText").value = q.questionText || "";' +
    '  document.getElementById("qType").value = q.type || "dropdown";' +
    '  document.getElementById("required").value = q.required || "Y";' +
    '  document.getElementById("active").value = q.active || "Y";' +
    '  document.getElementById("options").value = q.options || "";' +
    '  document.getElementById("branchParent").value = q.branchParent || "";' +
    '}).modalGetQuestion(ROW);' +
    'function saveQuestion() {' +
    '  var data = { row: ROW, section: document.getElementById("section").value,' +
    '    sectionTitle: document.getElementById("sectionTitle").value,' +
    '    questionText: document.getElementById("questionText").value,' +
    '    type: document.getElementById("qType").value,' +
    '    required: document.getElementById("required").value,' +
    '    active: document.getElementById("active").value,' +
    '    options: document.getElementById("options").value,' +
    '    branchParent: document.getElementById("branchParent").value };' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Question saved!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else showMsg(r && r.error || "Failed", "error");' +
    '  }).modalSaveQuestion(data);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalGetQuestion(row) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    if (!sheet || row < 2) return null;

    var data = sheet.getRange(row, 1, 1, 10).getValues()[0];
    return {
      id: data[0], section: data[1], sectionKey: data[2], sectionTitle: data[3],
      questionText: data[4], type: data[5], required: data[6], active: data[7],
      options: data[8], branchParent: data[9]
    };
  } catch (_e) {
    return null;
  }
}

function modalSaveQuestion(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SURVEY_QUESTIONS);
    if (!sheet) return { success: false, error: 'Survey Questions sheet not found' };

    // Write back edited fields (skip ID/col 1 and sectionKey/col 3 — auto-generated)
    // Batch: col 2 (section), then cols 4-10 (sectionTitle through branchParent)
    sheet.getRange(data.row, 2).setValue(escapeForFormula(data.section));
    sheet.getRange(data.row, 4, 1, 7).setValues([[
      escapeForFormula(data.sectionTitle),
      escapeForFormula(data.questionText),
      escapeForFormula(data.type),
      escapeForFormula(data.required),
      escapeForFormula(data.active),
      escapeForFormula(data.options),
      escapeForFormula(data.branchParent)
    ]]);

    return { success: true };
  } catch (e) {
    log_('modalSaveQuestion', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// MODAL 12: CREATE WEEKLY QUESTION (replaces Flash Polls)
// ============================================================================

function showCreateWeeklyQuestionModal() {
  showDialog_(getCreateWeeklyQuestionHtml_(), '⚡ Create Weekly Question', DIALOG_SIZES.MEDIUM.width, DIALOG_SIZES.MEDIUM.height);
}

function getCreateWeeklyQuestionHtml_() {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>.option-row { display:flex; gap:8px; margin-bottom:8px; align-items:center; } .option-row input { flex:1; } .remove-btn { background:#FEE2E2; color:#991B1B; border:none; border-radius:50%; width:28px; height:28px; cursor:pointer; font-size:16px; }</style>' +
    '</head><body>' +
    '<h2>Create Weekly Question</h2>' +
    '<div id="msg"></div>' +
    '<div class="form-group"><label>Question</label><textarea id="question" rows="3" placeholder="What is your biggest workplace concern this week?"></textarea></div>' +
    '<div class="form-group"><label>Answer Options</label><div id="options"></div>' +
    '<button class="btn btn-secondary" style="margin-top:8px;font-size:13px" onclick="addOption()">+ Add Option</button></div>' +
    '<div class="row"><div class="form-group"><label>Active</label><select id="active"><option value="Y">Yes — show to members</option><option value="N">No — save as draft</option></select></div></div>' +
    '<div class="actions"><button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="submitBtn" onclick="submitQuestion()">Create Question</button></div>' +
    '<script>' +
    'var optCount = 0;' +
    'function addOption() {' +
    '  optCount++; var div = document.createElement("div"); div.className = "option-row"; div.id = "opt" + optCount;' +
    '  div.innerHTML = "<input placeholder=\\"Option " + optCount + "\\" id=\\"optVal" + optCount + "\\"><button class=\\"remove-btn\\" onclick=\\"this.parentElement.remove()\\">×</button>";' +
    '  document.getElementById("options").appendChild(div);' +
    '}' +
    'addOption(); addOption(); addOption();' +
    'function submitQuestion() {' +
    '  var q = document.getElementById("question").value.trim();' +
    '  if (!q) { showMsg("Question text is required.", "error"); return; }' +
    '  var opts = []; var optEls = document.querySelectorAll("[id^=optVal]");' +
    '  for (var i = 0; i < optEls.length; i++) { var v = optEls[i].value.trim(); if (v) opts.push(v); }' +
    '  if (opts.length < 2) { showMsg("At least 2 options are required.", "error"); return; }' +
    '  document.getElementById("submitBtn").disabled = true;' +
    '  google.script.run.withSuccessHandler(function(r) {' +
    '    if (r && r.success) { showMsg("Weekly question created!", "success"); setTimeout(function() { google.script.host.close(); }, 1500); }' +
    '    else { showMsg(r && r.error || "Failed", "error"); document.getElementById("submitBtn").disabled = false; }' +
    '  }).withFailureHandler(function(e) { showMsg("Error: " + e.message, "error"); document.getElementById("submitBtn").disabled = false; })' +
    '  .modalCreateWeeklyQuestion(q, opts, document.getElementById("active").value);' +
    '}' +
    'function showMsg(t, type) { var el = document.getElementById("msg"); el.className = "msg msg-" + type; el.textContent = t; el.style.display = "block"; }' +
    '</script></body></html>';
}

function modalCreateWeeklyQuestion(question, options, active) {
  try {
    // Delegate to WeeklyQuestions module if available
    if (typeof WeeklyQuestions !== 'undefined' && typeof WeeklyQuestions.addStewardQuestion === 'function') {
      WeeklyQuestions.addStewardQuestion(question, options.join('|'));
      return { success: true };
    }

    // Fallback: write directly to _Weekly_Questions sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.WEEKLY_QUESTIONS);
    if (!sheet) return { success: false, error: 'Weekly Questions sheet not found. Run Admin > Initial Setup first.' };

    var id = 'WQ-' + Date.now();
    var user = Session.getActiveUser().getEmail() || 'Unknown';

    sheet.appendRow([
      escapeForFormula(id),
      escapeForFormula(question),
      escapeForFormula(options.join('|')),
      escapeForFormula(active),
      escapeForFormula(user),
      new Date()
    ]);

    return { success: true };
  } catch (e) {
    log_('modalCreateWeeklyQuestion', 'Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ============================================================================
// TAB-SPECIFIC MODALS (v4.48.0)
// Auto-open contextual modals when navigating to key sheet tabs.
// Each modal provides quick actions, tips, and links relevant to that tab.
// Controlled by Config: ENABLE_TAB_MODALS (default: 'yes').
// Users can dismiss per-tab via "Don't show again" checkbox.
// ============================================================================

/**
 * Returns true if tab modals are enabled in Config.
 * Defaults to true (enabled) when the config value is missing or blank.
 * @returns {boolean}
 */
function isTabModalsEnabled_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return true;
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return true;
    var val = configSheet.getRange(3, CONFIG_COLS.ENABLE_TAB_MODALS).getValue();
    return String(val).trim().toLowerCase() !== 'no';
  } catch (_e) { return true; }
}

/**
 * Returns true if the user has dismissed the tab modal for the given sheet.
 * @param {string} sheetKey - Key like 'CONFIG', 'MEMBER_DIR', etc.
 * @returns {boolean}
 */
function isTabModalDismissed_(sheetKey) {
  try {
    var val = PropertiesService.getUserProperties().getProperty('tabModal_dismiss_' + sheetKey);
    return val === 'true';
  } catch (_e) { return false; }
}

/**
 * Marks a tab modal as dismissed for the current user.
 * Called from the modal's "Don't show again" checkbox.
 * @param {string} sheetKey - Key like 'CONFIG', 'MEMBER_DIR', etc.
 */
function dismissTabModal(sheetKey) {
  try {
    PropertiesService.getUserProperties().setProperty('tabModal_dismiss_' + sheetKey, 'true');
  } catch (_e) { /* silent */ }
}

/**
 * Resets all tab modal dismissals for the current user.
 * Accessible via Admin > Settings or Tools menu.
 */
function resetTabModalDismissals() {
  try {
    var props = PropertiesService.getUserProperties();
    var all = props.getProperties();
    for (var key in all) {
      if (key.indexOf('tabModal_dismiss_') === 0) {
        props.deleteProperty(key);
      }
    }
    SpreadsheetApp.getUi().alert('Tab modal preferences have been reset. Modals will show again on tab navigation.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error resetting tab modals: ' + e.message);
  }
}

/**
 * Returns shared HTML wrapper for all tab modals.
 * Includes dismiss checkbox, close button, and shared styles.
 * @param {string} sheetKey - Key for dismiss tracking
 * @param {string} title - Modal title
 * @param {string} bodyHtml - Inner HTML content
 * @returns {string} Complete HTML string
 * @private
 */
function getTabModalHtml_(sheetKey, title, bodyHtml) {
  return '<!DOCTYPE html><html><head>' + getModalHead_() + getModalStyles_() +
    '<style>' +
    '.tab-modal-header { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 2px solid #E5E7EB; }' +
    '.tab-modal-header h2 { flex: 1; margin: 0; }' +
    '.quick-action { display: flex; align-items: center; gap: 10px; padding: 10px 14px; background: white; border: 1px solid #E5E7EB; border-radius: 8px; margin-bottom: 8px; cursor: pointer; transition: all 0.15s; }' +
    '.quick-action:hover { background: #F0F9FF; border-color: #2C5282; transform: translateX(2px); }' +
    '.quick-action .qa-icon { font-size: 20px; width: 28px; text-align: center; }' +
    '.quick-action .qa-text { flex: 1; }' +
    '.quick-action .qa-text strong { display: block; font-size: 13px; color: #1F2937; }' +
    '.quick-action .qa-text small { color: #6B7280; font-size: 11px; }' +
    '.tip-box { background: #FEF3C7; border: 1px solid #F59E0B; border-radius: 8px; padding: 10px 14px; margin-top: 12px; font-size: 12px; color: #92400E; }' +
    '.dismiss-row { display: flex; align-items: center; justify-content: space-between; margin-top: 16px; padding-top: 12px; border-top: 1px solid #E5E7EB; }' +
    '.dismiss-row label { font-size: 12px; color: #6B7280; cursor: pointer; display: flex; align-items: center; gap: 6px; }' +
    '.dismiss-row input[type=checkbox] { width: auto; }' +
    '.section-label { font-size: 11px; font-weight: 700; color: #9CA3AF; text-transform: uppercase; letter-spacing: 0.5px; margin: 14px 0 6px; }' +
    '</style></head><body>' +
    '<div class="tab-modal-header"><h2>' + title + '</h2></div>' +
    bodyHtml +
    '<div class="dismiss-row">' +
    '<label><input type="checkbox" id="dismissChk" onchange="handleDismiss()"> Don\'t show this again for this tab</label>' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<script>' +
    'function handleDismiss() { if (document.getElementById("dismissChk").checked) { google.script.run.dismissTabModal("' + sheetKey + '"); } }' +
    'function runAction(fn) { google.script.host.close(); google.script.run[fn](); }' +
    '</script></body></html>';
}

/**
 * Helper to show a tab modal dialog.
 * @param {string} html - Full HTML string
 * @param {string} title - Dialog title
 * @private
 */
function showTabModal_(html, title) {
  showDialog_(html, title, DIALOG_SIZES.MEDIUM.width, 420);
}

// ── Config Tab Modal ──

function showTabModalConfig() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showSettingsDialog\')"><div class="qa-icon">⚙️</div><div class="qa-text"><strong>Open Settings</strong><small>Calendar, email, drive integration toggles</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showRepairDialog\')"><div class="qa-icon">🔧</div><div class="qa-text"><strong>Repair Dashboard</strong><small>Fix missing sheets, formulas, and validations</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'restoreConfigAndDropdowns\')"><div class="qa-icon">🔄</div><div class="qa-text"><strong>Restore Dropdowns</strong><small>Re-seed default dropdown values in Config columns</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'applyConfigStyling\')"><div class="qa-icon">🎨</div><div class="qa-text"><strong>Apply Config Styling</strong><small>Re-apply colors, fonts, and section headers</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Rows 1–2 are headers (do not edit). Data starts at Row 3. Each column is a dropdown source or setting. Use the section headers in Row 1 to find what you need.</div>';
  showTabModal_(getTabModalHtml_('CONFIG', '⚙️ Config Tab', body), '⚙️ Config');
}

// ── Member Directory Tab Modal ──

function showTabModalMemberDirectory() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showNewMemberDialog\')"><div class="qa-icon">➕</div><div class="qa-text"><strong>Add New Member</strong><small>Create a new member record</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showSearchDialog\')"><div class="qa-icon">🔍</div><div class="qa-text"><strong>Find Member</strong><small>Search by name, email, or member ID</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showImportMembersDialog\')"><div class="qa-icon">📥</div><div class="qa-text"><strong>Import Members</strong><small>Bulk import from CSV or another sheet</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'generateMissingMemberIDs\')"><div class="qa-icon">🆔</div><div class="qa-text"><strong>Generate Missing IDs</strong><small>Auto-assign IDs to members without one</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'syncAllData\')"><div class="qa-icon">🔄</div><div class="qa-text"><strong>Sync All Data</strong><small>Sync grievances, attendance, and engagement</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Select a member row and use Union Hub > Members for actions like promote/demote steward, export, or duplicate checks.</div>';
  showTabModal_(getTabModalHtml_('MEMBER_DIR', '👥 Member Directory', body), '👥 Member Directory');
}

// ── Grievance Log Tab Modal ──

function showTabModalGrievanceLog() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showNewGrievanceDialog\')"><div class="qa-icon">📝</div><div class="qa-text"><strong>New Case / Grievance</strong><small>File a new grievance or case</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showEditGrievanceDialog\')"><div class="qa-icon">✏️</div><div class="qa-text"><strong>Edit Selected Grievance</strong><small>Modify the currently selected case</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showChecklistDialog\')"><div class="qa-icon">✅</div><div class="qa-text"><strong>View Case Checklist</strong><small>Track progress steps for the selected case</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'createPDFForSelectedGrievance\')"><div class="qa-icon">📄</div><div class="qa-text"><strong>Create Signature PDF</strong><small>Generate a PDF for the selected grievance</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'applyTrafficLightIndicators\')"><div class="qa-icon">🚦</div><div class="qa-text"><strong>Apply Traffic Lights</strong><small>Color-code rows by deadline urgency</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Select a grievance row first, then use quick actions. Traffic lights turn rows green/yellow/red based on deadline proximity.</div>';
  showTabModal_(getTabModalHtml_('GRIEVANCE_LOG', '⚖️ Grievance Log', body), '⚖️ Grievance Log');
}

// ── Case Checklist Tab Modal ──

function showTabModalCaseChecklist() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showCaseProgressModal\')"><div class="qa-icon">📊</div><div class="qa-text"><strong>Case Progress Overview</strong><small>View completion stats across all cases</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showChecklistDialog\')"><div class="qa-icon">✅</div><div class="qa-text"><strong>View Checklist for Selected</strong><small>Open checklist for the selected case</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'unlockChecklistSheet\')"><div class="qa-icon">🔓</div><div class="qa-text"><strong>Unlock Sheet</strong><small>Remove protection if locked</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Each row tracks a grievance case\'s checklist. Checkboxes auto-calculate completion percentage in the progress column.</div>';
  showTabModal_(getTabModalHtml_('CASE_CHECKLIST', '✅ Case Checklist', body), '✅ Case Checklist');
}

// ── Volunteer Hours Tab Modal ──

function showTabModalVolunteerHours() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showLogVolunteerHoursModal\')"><div class="qa-icon">🤝</div><div class="qa-text"><strong>Log Volunteer Hours</strong><small>Record hours for a member</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'syncVolunteerHoursToMemberDirectory\')"><div class="qa-icon">🔄</div><div class="qa-text"><strong>Sync to Member Directory</strong><small>Update total hours in member records</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Hours logged here automatically roll up to each member\'s total in the Member Directory when synced.</div>';
  showTabModal_(getTabModalHtml_('VOLUNTEER_HOURS', '🤝 Volunteer Hours', body), '🤝 Volunteer Hours');
}

// ── Meeting Attendance Tab Modal ──

function showTabModalMeetingAttendance() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showTakeAttendanceModal\')"><div class="qa-icon">📋</div><div class="qa-text"><strong>Take Attendance</strong><small>Record who attended a meeting</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'syncMeetingAttendanceToMemberDirectory\')"><div class="qa-icon">🔄</div><div class="qa-text"><strong>Sync to Member Directory</strong><small>Update attendance counts in member records</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showAddEventModal\')"><div class="qa-icon">📅</div><div class="qa-text"><strong>Add New Event</strong><small>Schedule a new union event</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Members can also check in via the web app or QR code. Use Tools > Calendar & Meetings for QR and meeting setup.</div>';
  showTabModal_(getTabModalHtml_('MEETING_ATTENDANCE', '📅 Meeting Attendance', body), '📅 Meeting Attendance');
}

// ── Meeting Check-In Log Tab Modal ──

function showTabModalMeetingCheckIn() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'showSetupMeetingDialog\')"><div class="qa-icon">📝</div><div class="qa-text"><strong>Setup New Meeting</strong><small>Create a new meeting session for check-ins</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showMeetingCheckInDialog\')"><div class="qa-icon">✅</div><div class="qa-text"><strong>Open Meeting Check-In</strong><small>Start checking in members</small></div></div>' +
    '<div class="quick-action" onclick="runAction(\'showMeetingQRCodeDialog\')"><div class="qa-icon">📱</div><div class="qa-text"><strong>QR Code Check-In</strong><small>Generate QR code for mobile check-in</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> This log records all check-ins (manual, web app, and QR). Each row is a single check-in event with timestamp and method.</div>';
  showTabModal_(getTabModalHtml_('MEETING_CHECKIN_LOG', '📝 Meeting Check-In Log', body), '📝 Meeting Check-In Log');
}

// ── Resources Tab Modal ──

function showTabModalResources() {
  var body =
    '<div class="section-label">Quick Actions</div>' +
    '<div class="quick-action" onclick="runAction(\'openGoogleDrive\')"><div class="qa-icon">📁</div><div class="qa-text"><strong>Open Google Drive</strong><small>Browse shared resource files</small></div></div>' +
    '<div class="tip-box">💡 <strong>Tip:</strong> Resources are organized by category. Members can browse and download them from the web app Resources tab. Manage categories in the Resource Config sheet.</div>';
  showTabModal_(getTabModalHtml_('RESOURCES', '📚 Resources', body), '📚 Resources');
}
