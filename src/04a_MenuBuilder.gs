/**
 * ============================================================================
 * MenuBuilder.gs - Menu System and Navigation
 * ============================================================================
 *
 * This module handles all menu-related functions including:
 * - Main menu creation (createDashboardMenu)
 * - Sheet navigation
 * - Quick action menus
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Menu creation and navigation functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MENU CREATION
// ============================================================================

/**
 * Creates the custom menu when the spreadsheet opens
 * Consolidated menu structure (v4.4.0):
 * - 📊 509 Union Hub: Primary operations (dashboards, search, members, cases)
 * - 🔧 Tools: Supporting features (calendar, drive, notifications, etc.)
 * - 🛠️ Admin: System administration
 * @returns {void}
 */
function createDashboardMenu() {
  var ui = SpreadsheetApp.getUi();

  // ============================================================================
  // MENU 1: 509 Union Hub - Primary Operations
  // ============================================================================
  ui.createMenu('📊 509 Union Hub')
    .addItem('👥 Member Dashboard', 'showPublicMemberDashboard')
    .addItem('🛡️ Steward Dashboard', 'showStewardDashboard')
    .addSeparator()

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
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward'))

    .addSubMenu(ui.createMenu('📋 Cases & Grievances')
      .addItem('➕ New Case/Grievance', 'showNewGrievanceDialog')
      .addItem('✏️ Edit Selected', 'showEditGrievanceDialog')
      .addItem('✅ View Checklist', 'showChecklistDialog')
      .addSeparator()
      .addItem('🔄 Bulk Update Status', 'showBulkStatusUpdate')
      .addItem('📄 Create PDF', 'createPDFForSelectedGrievance')
      .addSeparator()
      .addItem('🚦 Apply Traffic Lights', 'applyTrafficLightIndicators')
      .addItem('🔄 Clear Traffic Lights', 'clearTrafficLightIndicators'))

    .addSubMenu(ui.createMenu('📈 Analytics & Reports')
      .addItem('📈 Case Analytics', 'showInteractiveDashboardTab')
      .addItem('📊 Grievance Trends', 'showGrievanceTrends')
      .addItem('🏥 Unit Health Report', 'showUnitHealthReport')
      .addItem('📚 Search Precedents', 'showSearchPrecedents'))

    .addSeparator()
    .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
    .addItem('📖 Help & Documentation', 'showHelpDialog')
    .addToUi();

  // ============================================================================
  // MENU 2: Tools - Supporting Features
  // ============================================================================
  ui.createMenu('🔧 Tools')
    .addSubMenu(ui.createMenu('📧 Communication')
      .addItem('📧 Send Contact Form', 'sendContactInfoForm')
      .addItem('📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink')
      .addItem('📧 Send Portal Email', 'sendPortalEmailToSelectedMember'))

    .addSubMenu(ui.createMenu('📅 Calendar')
      .addItem('🔗 Sync Deadlines', 'syncDeadlinesToCalendar')
      .addItem('👁️ View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar')
      .addItem('🗑️ Clear Calendar Events', 'clearAllCalendarEvents'))

    .addSubMenu(ui.createMenu('📁 Google Drive')
      .addItem('📁 Setup Folder for Grievance', 'setupFolderForSelectedGrievance')
      .addItem('📁 View Grievance Files', 'showGrievanceFiles')
      .addItem('📁 Batch Create Folders', 'batchCreateAllMissingFolders'))

    .addSubMenu(ui.createMenu('🔔 Notifications')
      .addItem('⚙️ Notification Settings', 'showNotificationSettings')
      .addItem('🧪 Test Notifications', 'testDeadlineNotifications'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🎛️ View & Display')
      .addItem('🎛️ Visual Control Panel', 'showVisualControlPanel')
      .addItem('📱 Toggle Mobile View', 'toggleMobileView')
      .addItem('📱 Pocket View', 'navToMobile')
      .addItem('🖥️ Restore Desktop View', 'showAllMemberColumns')
      .addSeparator()
      .addItem('🎨 Apply Theme', 'APPLY_SYSTEM_THEME')
      .addItem('🔄 Reset to Default', 'resetToDefaultTheme')
      .addItem('✨ Refresh All Visuals', 'refreshAllVisuals'))

    .addSubMenu(ui.createMenu('📝 Multi-Select')
      .addItem('📝 Open Editor', 'openCellMultiSelectEditor')
      .addItem('⚡ Enable Auto-Open', 'installMultiSelectTrigger')
      .addItem('🚫 Disable Auto-Open', 'removeMultiSelectTrigger'))

    .addSubMenu(ui.createMenu('📝 OCR Tools')
      .addItem('📝 OCR Transcribe Form', 'showOCRDialog')
      .addItem('🔧 OCR Setup', 'setupOCRApiKey'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🌐 Web App & Portal')
      .addItem('📱 Get Mobile App URL', 'showWebAppUrl')
      .addItem('👤 Build Member Portal', 'buildPortalForSelectedMember')
      .addItem('📊 Build Public Portal', 'buildPublicPortal'))

    .addSubMenu(ui.createMenu('🔄 Maintenance')
      .addItem('🔄 Refresh All Formulas', 'refreshAllFormulas')
      .addItem('🔄 Refresh Member Data', 'refreshMemberDirectoryFormulas')
      .addItem('🔄 Refresh View', 'refreshMemberView'))

    .addToUi();

  // ============================================================================
  // MENU 3: Admin - System Administration
  // ============================================================================
  var adminMenu = ui.createMenu('🛠️ Admin')
    .addItem('🩺 System Diagnostics', 'showDiagnosticsDialog')
    .addItem('🔍 Modal Diagnostics', 'showModalDiagnostics')
    .addItem('🔧 Repair Dashboard', 'showRepairDialog')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addSeparator()

    .addSubMenu(ui.createMenu('🔄 Data Sync')
      .addItem('🔄 Sync All Data Now', 'syncAllData')
      .addItem('🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory')
      .addItem('🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog')
      .addSeparator()
      .addItem('⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger')
      .addItem('🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger'))

    .addSubMenu(ui.createMenu('✅ Validation')
      .addItem('🔍 Run Bulk Validation', 'runBulkValidation')
      .addItem('⚙️ Validation Settings', 'showValidationSettings')
      .addItem('🧹 Clear Indicators', 'clearValidationIndicators')
      .addItem('⚡ Install Validation Trigger', 'installValidationTrigger'))

    .addSubMenu(ui.createMenu('⚙️ Automation')
      .addItem('🔄 Force Global Refresh', 'refreshAllVisuals')
      .addItem('🌙 Enable Midnight Auto-Refresh', 'setupMidnightTrigger')
      .addItem('❌ Disable Midnight Auto-Refresh', 'removeMidnightTrigger')
      .addItem('🔔 Enable 1AM Dashboard Refresh', 'createAutomationTriggers')
      .addItem('📑 Email Weekly PDF Snapshot', 'emailExecutivePDF'))

    .addSubMenu(ui.createMenu('🗄️ Cache & Performance')
      .addItem('🗄️ Cache Status', 'showCacheStatusDashboard')
      .addItem('🔥 Warm Up Caches', 'warmUpCaches')
      .addItem('🗑️ Clear All Caches', 'invalidateAllCaches'))

    .addSubMenu(ui.createMenu('🛡️ Security & Backup')
      .addItem('📸 Create Manual Snapshot', 'createWeeklySnapshot')
      .addItem('📅 Setup Weekly Backup', 'setupWeeklySnapshotTrigger')
      .addItem('📜 View Audit Log', 'navigateToAuditLog')
      .addItem('📊 v4.0 Status Report', 'showV4StatusReport'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🎨 Styling')
      .addItem('🎨 Apply Config Sheet Styling', 'applyConfigStyling')
      .addItem('🎨 Apply Tab Colors', 'applyTabColors')
      .addItem('🖌️ Setup Theme Columns', 'setupThemeColumns'))

    .addSubMenu(ui.createMenu('🏗️ Setup')
      .addItem('🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets')
      .addItem('🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets')
      .addItem('🔍 Verify Hidden Sheets', 'verifyHiddenSheets')
      .addItem('⚙️ Setup Data Validations', 'setupDataValidations')
      .addItem('🎨 Setup Comfort View Defaults', 'setupADHDDefaults')
      .addSeparator()
      .addItem('📋 Create Features Reference Sheet', 'createFeaturesReferenceSheet')
      .addItem('❓ Create FAQ Sheet', 'createFAQSheet')
      .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
      .addItem('📱 Add Mobile Dashboard Link', 'addMobileDashboardLinkToConfig')
      .addItem('🔓 Unlock Checklist Sheet', 'unlockChecklistSheet'));

  // Only show Demo Data menu if NOT in production mode
  if (!isProductionMode()) {
    adminMenu.addSeparator()
      .addSubMenu(ui.createMenu('🎭 Demo Data')
        .addItem('🚀 Seed All Sample Data', 'SEED_SAMPLE_DATA')
        .addItem('👥 Seed Members Only...', 'SEED_MEMBERS_DIALOG')
        .addItem('📋 Seed Grievances Only...', 'SEED_GRIEVANCES_DIALOG')
        .addSeparator()
        .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
        .addSeparator()
        .addItem('☢️ NUKE SEEDED DATA', 'NUKE_SEEDED_DATA'));
  }

  adminMenu.addToUi();
}

// ============================================================================
// NAVIGATION FUNCTIONS
// ============================================================================

/**
 * Navigates to the Dashboard sheet
 * @returns {void}
 */
function navToDash() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * Navigates to mobile view
 * @returns {void}
 * @private
 */
function navToMobile_UIService_() {
  showSmartDashboard();
}

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
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Info', 3);
}

/**
 * Shows a confirmation dialog
 * @param {string} message - The message to display
 * @param {string} title - The dialog title
 * @returns {boolean} True if user clicked Yes
 */
function showConfirmation(message, title) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(title || 'Confirm', message, ui.ButtonSet.YES_NO);
  return result === ui.Button.YES;
}

/**
 * Shows an alert dialog
 * @param {string} message - The message to display
 * @param {string} title - The dialog title
 * @returns {void}
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
 * Navigates to a member in the sheet
 * @param {string} memberId - The member ID
 * @returns {void}
 */
function navigateToMemberInSheet(memberId) {
  navigateToRecord(memberId, 'member');
}

/**
 * Navigates to a grievance in the sheet
 * @param {string} grievanceId - The grievance ID
 * @returns {void}
 */
function navigateToGrievanceInSheet(grievanceId) {
  navigateToRecord(grievanceId, 'grievance');
}

/**
 * Shows the Member Directory sheet
 * @returns {void}
 */
function showMemberDirectory() {
  navigateToSheet(SHEETS.MEMBER_DIR);
}

/**
 * Shows the Grievance Log sheet
 * @returns {void}
 */
function showGrievanceLog() {
  navigateToSheet(SHEETS.GRIEVANCE_LOG);
}

/**
 * Shows the Config sheet
 * @returns {void}
 */
function showConfigSheet() {
  navigateToSheet(SHEETS.CONFIG);
}
