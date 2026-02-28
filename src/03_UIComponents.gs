/**
 * ============================================================================
 * 03_UIComponents.gs - Menu System and Navigation
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
 * @version 4.7.0
 * @requires 01_Core.gs
 */

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
  // MENU 1: Union Hub - Primary Operations
  // ============================================================================
  ui.createMenu('📊 ' + localNumber + ' Union Hub')
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
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward')
      .addItem('🔄 Sync Steward Status', 'syncStewardStatus'))

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
    .addItem('📅 Open Google Calendar', 'openGoogleCalendar')
    .addItem('📁 Open Google Drive', 'openGoogleDrive')
    .addItem('📖 Help & Documentation', 'showHelpDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Workload Tracker')
      .addItem('📋 Open Workload Portal', 'showWorkloadPortalUrl')
      .addItem('🔗 Save Portal Link', 'shareWorkloadPortalLink')
      .addSeparator()
      .addItem('🔄 Refresh Ledger', 'refreshWorkloadLedger')
      .addItem('💾 Create Backup', 'createWorkloadBackup')
      .addItem('🗄️ Archive Old Data', 'wtArchiveOldData_')
      .addSeparator()
      .addItem('🩺 Health Status', 'showWorkloadHealthStatus')
      .addItem('🔔 Setup Reminders', 'setupWorkloadReminderSystem'))
    .addToUi();

  // ============================================================================
  // MENU 2: Tools - Supporting Features
  // ============================================================================
  ui.createMenu('🔧 Tools')
    .addSubMenu(ui.createMenu('📧 Communication')
      .addItem('📧 Send Contact Form', 'sendContactInfoForm')
      .addItem('📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink')
      .addItem('📧 Send Portal Email', 'sendPortalEmailToSelectedMember'))

    .addSubMenu(ui.createMenu('📅 Calendar & Meetings')
      .addItem('📝 Setup New Meeting', 'showSetupMeetingDialog')
      .addItem('✅ Open Meeting Check-In', 'showMeetingCheckInDialog')
      .addSeparator()
      .addItem('🔗 Sync Deadlines', 'syncDeadlinesToCalendar')
      .addItem('👁️ View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar'))

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
      .addItem('📱 Toggle Pocket View', 'toggleMobileView')
      .addItem('📱 Pocket View (Active Sheet)', 'navToMobile')
      .addItem('🖥️ Restore All Columns', 'showAllColumns')
      .addSeparator()
      .addItem('🎨 Apply Theme', 'APPLY_SYSTEM_THEME')
      .addItem('🎨 Theme Presets', 'showThemePresetPicker')
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
      .addItem('📊 Build Public Portal', 'buildPublicPortal')
      .addSeparator()
      .addItem('🔑 Generate Member PIN', 'showGeneratePINDialog')
      .addItem('🔄 Reset Member PIN', 'showResetPINDialog')
      .addItem('📋 Bulk Generate PINs', 'showBulkGeneratePINDialog'))

    .addSubMenu(ui.createMenu('📊 Workload Tracker')
      .addItem('📊 Refresh Ledger', 'refreshWorkloadLedger')
      .addItem('💾 Create Backup', 'createWorkloadBackup')
      .addItem('🗄️ Archive Old Data', 'archiveWorkloadData')
      .addItem('🩺 Health Status', 'showWorkloadHealthStatus')
      .addSeparator()
      .addItem('🔔 Setup Reminders', 'setupWorkloadReminderSystem')
      .addItem('🔗 Show Portal URL', 'showWorkloadPortalUrl')
      .addItem('📋 Save Portal Link', 'shareWorkloadPortalLink')
      .addSeparator()
      .addItem('⚙️ Initialize Sheets', 'initWorkloadTrackerSheets'))

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
      .addItem('📊 v4.0 Status Report', 'showV4StatusReport')
      .addSeparator()
      .addItem('🔐 Dashboard Auth Status', 'showDashboardAuthStatus')
      .addItem('✅ Enable Dashboard Member Auth', 'enableDashboardMemberAuth')
      .addItem('❌ Disable Dashboard Member Auth', 'disableDashboardMemberAuth'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🎨 Styling')
      .addItem('🎨 Apply Config Sheet Styling', 'applyConfigStyling')
      .addItem('🎨 Apply Tab Colors', 'applyTabColors')
      .addItem('🖌️ Setup Theme Columns', 'setupThemeColumns'))

    .addSubMenu(ui.createMenu('🏗️ Setup')
      .addItem('🚀 Initialize Dashboard', 'initializeDashboard')
      .addSeparator()
      .addItem('🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets')
      .addItem('🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets')
      .addItem('🔍 Verify Hidden Sheets', 'verifyHiddenSheets')
      .addItem('🔒 Enforce Hidden Sheets (Mobile Fix)', 'enforceHiddenSheets')
      .addItem('⚙️ Setup Data Validations', 'setupDataValidations')
      .addItem('🎨 Setup Comfort View Defaults', 'setupADHDDefaults')
      .addSeparator()
      .addItem('📝 Create Meeting Check-In Sheet', 'setupMeetingCheckInSheet')
      .addItem('📋 Create Features Reference Sheet', 'createFeaturesReferenceSheet')
      .addItem('❓ Create FAQ Sheet', 'createFAQSheet')
      .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
      .addItem('📥 Populate Config from Sheet Data', 'populateConfigFromSheetData')
      .addItem('📱 Add Mobile Dashboard Link', 'addMobileDashboardLinkToConfig')
      .addItem('🔓 Unlock Checklist Sheet', 'unlockChecklistSheet')
      .addSeparator()
      .addItem('🗑️ Remove Deprecated Dashboard', 'removeDeprecatedDashboard'));

  // Only show Demo Data menu if NOT in production mode
  if (!isProductionMode()) {
    adminMenu.addSeparator()
      .addSubMenu(ui.createMenu('🎭 Demo Data')
        .addItem('🚀 Seed All Sample Data', 'SEED_SAMPLE_DATA')
        .addItem('👥 Seed Members Only...', 'SEED_MEMBERS_DIALOG')
        .addItem('📋 Seed Grievances Only...', 'SEED_GRIEVANCES_DIALOG')
        .addSeparator()
        .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns')
        .addItem('📥 Populate Config from Sheet Data', 'populateConfigFromSheetData')
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
/**
 * Opens Google Calendar in a new browser tab
 */
function openGoogleCalendar() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("https://calendar.google.com", "_blank");google.script.host.close();</script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Calendar...');
}

/**
 * Opens Google Drive in a new browser tab
 */
function openGoogleDrive() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("https://drive.google.com", "_blank");google.script.host.close();</script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Drive...');
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



/**
 * ============================================================================
 * ThemeService.gs - Theme Management and Visual Settings
 * ============================================================================
 *
 * This module handles all theme-related functions including:
 * - Theme application (dark mode, light mode, etc.)
 * - Visual settings persistence
 * - Comfort view / ADHD-friendly settings
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Theme management and visual settings functions
 * @version 1.0.0
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
 * Applies theme styling to a single sheet
 * @param {Sheet} sheet - The sheet to style
 * @returns {void}
 * @private
 */
function clampFontSize_(size) {
  var n = Number(size) || 10;
  return Math.max(8, Math.min(24, n));
}

function applyThemeToSheet_(sheet) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  if (lastCol < 1 || lastRow < 1) return;

  // Get the active theme preset
  var theme = getActiveThemePreset_();

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
      var color = (row % 2 === 0) ? theme.altRow : '#ffffff';
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
    headerBg: '#1e293b',
    headerText: '#ffffff',
    altRow: '#f8fafc',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'union-blue': {
    name: 'Union Blue',
    headerBg: '#1e40af',
    headerText: '#ffffff',
    altRow: '#eff6ff',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'forest': {
    name: 'Forest Green',
    headerBg: '#166534',
    headerText: '#ffffff',
    altRow: '#f0fdf4',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'charcoal': {
    name: 'Charcoal',
    headerBg: '#374151',
    headerText: '#f9fafb',
    altRow: '#f3f4f6',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'midnight': {
    name: 'Midnight Purple',
    headerBg: '#581c87',
    headerText: '#ffffff',
    altRow: '#faf5ff',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'crimson': {
    name: 'Crimson',
    headerBg: '#991b1b',
    headerText: '#ffffff',
    altRow: '#fef2f2',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'steel': {
    name: 'Steel Gray',
    headerBg: '#475569',
    headerText: '#ffffff',
    altRow: '#f1f5f9',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
  },
  'ocean': {
    name: 'Ocean Teal',
    headerBg: '#115e59',
    headerText: '#ffffff',
    altRow: '#f0fdfa',
    font: 'Roboto',
    fontSize: 10,
    headerSize: 11
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

  var output = HtmlService.createHtmlOutput(html).setWidth(340).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(output, 'Theme Presets');
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
  APPLY_SYSTEM_THEME();
  SpreadsheetApp.getActiveSpreadsheet().toast('Applied "' + THEME_PRESETS[presetKey].name + '" theme to all tabs!', 'Theme', 3);
}

/**
 * Applies global styling to the spreadsheet
 * @returns {void}
 */
function applyGlobalStyling() {
  APPLY_SYSTEM_THEME();
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
  SpreadsheetApp.getActiveSpreadsheet().toast('Theme changed to ' + theme, 'Theme', 2);
}

/**
 * Toggles dark mode
 * @returns {void}
 */
function toggleDarkMode() {
  var props = PropertiesService.getUserProperties();
  var isDark = props.getProperty('visual_darkMode') === 'true';
  props.setProperty('visual_darkMode', (!isDark).toString());
  SpreadsheetApp.getActiveSpreadsheet().toast(
    isDark ? 'Dark mode disabled' : 'Dark mode enabled',
    'Theme', 2
  );
}

/**
 * Shows the theme settings dialog
 * @returns {void}
 */
function showThemeSettings() {
  showVisualControlPanel();
}

// ============================================================================
// COMFORT VIEW / ADHD-FRIENDLY SETTINGS
// ============================================================================

/**
 * Gets ADHD/Comfort View settings
 * @returns {Object} Settings object
 */
function getADHDSettings() {
  var props = PropertiesService.getUserProperties();
  var settings = props.getProperty('adhdSettings');
  return settings ? JSON.parse(settings) : getDefaultADHDSettings_();
}

/**
 * Gets default ADHD settings
 * @returns {Object} Default settings
 * @private
 */
function getDefaultADHDSettings_() {
  return {
    zebraStripes: true,
    reducedMotion: false,
    focusMode: false,
    highContrast: false,
    largeText: false,
    hideGridlines: false
  };
}

/**
 * Saves ADHD/Comfort View settings
 * @param {Object} settings - The settings to save
 * @returns {void}
 */
function saveADHDSettings(settings) {
  var props = PropertiesService.getUserProperties();
  props.setProperty('adhdSettings', JSON.stringify(settings));
}

/**
 * Applies ADHD/Comfort View settings
 * @param {Object} settings - The settings to apply
 * @returns {void}
 */
function applyADHDSettings(settings) {
  if (settings.zebraStripes) {
    applyZebraStripesToAllSheets_();
  }
  if (settings.hideGridlines) {
    hideAllGridlines();
  }
  saveADHDSettings(settings);
}

/**
 * Resets ADHD/Comfort View settings to defaults
 * @returns {void}
 */
function resetADHDSettings() {
  saveADHDSettings(getDefaultADHDSettings_());
  SpreadsheetApp.getActiveSpreadsheet().toast('Comfort View settings reset', 'Settings', 2);
}

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

/**
 * Shows gridlines on all sheets
 * @returns {void}
 */
function showAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    sheet.setHiddenGridlines(false);
  });
}

/**
 * Toggles gridlines visibility
 * @returns {void}
 */
function toggleGridlinesADHD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  var hidden = currentSheet.hasHiddenGridlines();
  currentSheet.setHiddenGridlines(!hidden);
}

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
    var color = (row % 2 === 0) ? '#f1f5f9' : '#ffffff';
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
  dataRange.setBackground('#ffffff');
}

/**
 * Toggles zebra stripes on active sheet
 * @returns {void}
 */
function toggleZebraStripes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if zebra stripes are applied (row 2 has different color than row 3)
  var row2Color = sheet.getRange(2, 1).getBackground();
  var row3Color = sheet.getRange(3, 1).getBackground();

  if (row2Color !== row3Color) {
    removeZebraStripes(sheet);
  } else {
    applyZebraStripes(sheet);
  }
}

/**
 * Toggles reduced motion setting
 * @returns {void}
 */
function toggleReducedMotion() {
  var settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    settings.reducedMotion ? 'Reduced motion enabled' : 'Reduced motion disabled',
    'Accessibility', 2
  );
}

/**
 * Activates focus mode
 * @returns {void}
 */
function activateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selection = sheet.getSelection();

  if (!selection) {
    ss.toast('Please select a cell first', 'Focus Mode', 2);
    return;
  }

  var settings = getADHDSettings();
  settings.focusMode = true;
  saveADHDSettings(settings);

  ss.toast('Focus mode activated. Press Esc to exit.', 'Focus Mode', 3);
}

/**
 * Deactivates focus mode
 * @returns {void}
 */
function deactivateFocusMode() {
  var settings = getADHDSettings();
  settings.focusMode = false;
  saveADHDSettings(settings);

  SpreadsheetApp.getActiveSpreadsheet().toast('Focus mode deactivated', 'Focus Mode', 2);
}

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
  } catch(_e) {
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
 * Applies theme to a specific sheet
 * @param {Sheet} sheet - The sheet
 * @param {Object} theme - Theme object
 * @returns {void}
 */
function applyThemeToSheet(sheet, theme) {
  applyThemeToSheet_(sheet);
}

/**
 * Previews a theme
 * @param {string} themeKey - Theme key to preview
 * @returns {void}
 */
function previewTheme(themeKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  applyThemeToSheet_(sheet, themeKey);
  ss.toast('Previewing ' + themeKey + ' theme', 'Theme Preview', 2);
}

/**
 * Resets to default theme via system
 * @returns {void}
 * @private
 */
function resetToDefaultThemeViaSystem_() {
  resetToDefaultTheme();
}

/**
 * Quick toggle for dark mode
 * @returns {void}
 */
function quickToggleDarkMode() {
  toggleDarkMode();
}

/**
 * Sets up ADHD/Comfort View defaults
 * @returns {void}
 */
function setupADHDDefaults() {
  var settings = getDefaultADHDSettings_();
  settings.zebraStripes = true;
  applyADHDSettings(settings);
  SpreadsheetApp.getActiveSpreadsheet().toast('Comfort View defaults applied', 'Settings', 3);
}

/**
 * Applies ADHD defaults with options
 * @param {Object} options - Options to apply
 * @returns {void}
 */
function applyADHDDefaultsWithOptions(options) {
  var settings = getADHDSettings();
  Object.assign(settings, options);
  applyADHDSettings(settings);
}

/**
 * Undoes ADHD/Comfort View defaults
 * @returns {void}
 */
function undoADHDDefaults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.getName().indexOf('_Calc') !== 0) {
      removeZebraStripes(sheet);
      sheet.setHiddenGridlines(false);
    }
  });

  resetADHDSettings();
  ss.toast('Comfort View settings removed', 'Settings', 3);
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
        Logger.log('Skipped theme for sheet: ' + sheetName + ': ' + _e.message);
      }
    }

    // Reapply ADHD settings if enabled
    var settings = getADHDSettings();
    if (settings.zebraStripes) {
      applyZebraStripesToAllSheets_();
    }

    ss.toast('All visuals refreshed! (' + processed + ' sheets)', 'Refresh', 3);
  } catch (e) {
    ss.toast('Refresh error: ' + e.message, 'Error', 5);
    Logger.log('refreshAllVisuals error: ' + e.toString());
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
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MOBILE CONTEXT DETECTION
// ============================================================================
// Note: MOBILE_CONFIG is now defined in 01_Core.gs as a shared constant

/**
 * Checks if the current context is a mobile device
 * Server-side detection is limited; this function exists for potential
 * future use with session properties or client-side communication
 * @returns {boolean} Always returns false on server-side
 */
function isMobileContext() {
  // Server-side we can't reliably detect mobile
  // This function exists for potential future use with session properties
  return false;
}

// ============================================================================
// MOBILE DASHBOARD
// ============================================================================

/**
 * Shows the mobile-optimized dashboard interface
 * Features:
 * - Touch-friendly stat cards
 * - Quick action buttons
 * - Responsive layout
 * @returns {void}
 */
function showMobileDashboard() {
  var stats = getMobileDashboardStats();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;padding:0;margin:0;background:#f5f5f5}' +
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:clamp(14px,4vw,20px);padding-top:calc(clamp(14px,4vw,20px) + env(safe-area-inset-top,0px))}' +
    '.header h1{margin:0;font-size:clamp(20px,5vw,24px)}.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +
    '.container{padding:clamp(10px,3vw,15px);padding-bottom:80px}' +
    '.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:clamp(8px,2vw,12px);margin-bottom:20px}' +
    '.stat-card{background:white;padding:clamp(14px,3.5vw,20px);border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center}' +
    '.stat-value{font-size:clamp(24px,7vw,32px);font-weight:bold;color:#1a73e8}' +
    '.stat-label{font-size:clamp(10px,2.8vw,13px);color:#666;text-transform:uppercase;margin-top:4px}' +
    '.section-title{font-size:clamp(14px,3.8vw,16px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +
    '.action-btn{background:white;border:none;padding:clamp(12px,3.5vw,16px);margin-bottom:10px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);width:100%;text-align:left;display:flex;align-items:center;gap:clamp(10px,3vw,15px);font-size:clamp(13px,3.5vw,15px);cursor:pointer;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.action-btn:active{transform:scale(0.98)}' +
    '.action-icon{font-size:clamp(20px,5vw,24px);width:clamp(36px,9vw,40px);height:clamp(36px,9vw,40px);display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px;flex-shrink:0}' +
    '.action-label{font-weight:500}.action-desc{font-size:clamp(10px,2.8vw,12px);color:#666;margin-top:2px}' +
    '.fab{position:fixed;bottom:calc(20px + env(safe-area-inset-bottom,0px));right:20px;width:56px;height:56px;background:#1a73e8;color:white;border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer;z-index:100}' +
    '@media(max-width:360px){.stats{grid-template-columns:1fr 1fr;gap:6px}.stat-card{padding:10px}.container{padding:8px;padding-bottom:80px}}' +
    '</style></head><body>' +
    '<div class="header"><h1>📱 Dashboard</h1><div class="subtitle">Mobile View</div></div>' +
    '<div class="container"><div class="stats"><div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div><div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div><div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div><div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div></div><div class="section-title">⚡ Quick Actions</div><button class="action-btn" onclick="google.script.run.showMobileGrievanceList()"><div class="action-icon">📋</div><div><div class="action-label">View Grievances</div><div class="action-desc">Browse all grievances</div></div></button><button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()"><div class="action-icon">🔍</div><div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div></button><button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()"><div class="action-icon">👤</div><div><div class="action-label">My Cases</div><div class="action-desc">View assigned grievances</div></div></button></div><button class="fab" onclick="location.reload()">🔄</button></body></html>'
  ).setWidth(400).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Mobile Dashboard');
}

/**
 * Gets dashboard statistics for mobile display
 * @returns {Object} Statistics object with totalGrievances, activeGrievances, pendingGrievances, overdueGrievances
 */
function getMobileDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return { totalGrievances: 0, activeGrievances: 0, pendingGrievances: 0, overdueGrievances: 0 };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.DAYS_TO_DEADLINE).getValues();
  var stats = { totalGrievances: data.length, activeGrievances: 0, pendingGrievances: 0, overdueGrievances: 0 };
  var today = new Date(); today.setHours(0, 0, 0, 0);
  data.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var daysTo = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    if (status && status !== 'Resolved' && status !== 'Withdrawn') stats.activeGrievances++;
    if (status === 'Pending Info') stats.pendingGrievances++;
    if ((daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0)) && status === 'Open') stats.overdueGrievances++;
  });
  return stats;
}

// ============================================================================
// MOBILE GRIEVANCE DATA
// ============================================================================

/**
 * Gets recent grievances formatted for mobile display
 * @param {number} [limit=5] Maximum number of grievances to return
 * @returns {Array<Object>} Array of grievance objects with mobile-friendly properties
 */
function getRecentGrievancesForMobile(limit) {
  limit = limit || 5;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
  return data.map(function(row, idx) {
    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var deadline = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    return {
      id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
      memberName: row[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + row[GRIEVANCE_COLS.LAST_NAME - 1],
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
      status: row[GRIEVANCE_COLS.STATUS - 1],
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, Session.getScriptTimeZone(), 'MM/dd/yyyy') : filed,
      deadline: deadline instanceof Date ? Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'MM/dd/yyyy') : null,
      filedDateObj: filed
    };
  }).sort(function(a, b) {
    var da = a.filedDateObj instanceof Date ? a.filedDateObj : new Date(0);
    var db = b.filedDateObj instanceof Date ? b.filedDateObj : new Date(0);
    return db - da;
  }).slice(0, limit);
}

// ============================================================================
// MOBILE GRIEVANCE LIST
// ============================================================================

/**
 * Shows mobile-optimized grievance list interface
 * Features:
 * - Card-based layout
 * - Status filters
 * - Search functionality
 * - Responsive grid
 * @returns {void}
 */
function showMobileGrievanceList() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    getMobileOptimizedHead() +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;margin:0;padding:0;background:#f5f5f5}' +
    '.header{background:#1a73e8;color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{margin:0;font-size:clamp(18px,4vw,24px)}' +
    '.search{width:100%;padding:clamp(10px,2.5vw,14px);border:none;border-radius:8px;font-size:clamp(14px,3vw,16px);margin-top:10px}' +
    '.filters{display:flex;overflow-x:auto;padding:10px;background:white;gap:8px;-webkit-overflow-scrolling:touch}' +
    '.filter{padding:clamp(6px,1.5vw,10px) clamp(12px,3vw,18px);border-radius:20px;background:#f0f0f0;white-space:nowrap;cursor:pointer;font-size:clamp(12px,2.5vw,14px);border:none;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';display:flex;align-items:center}' +
    '.filter.active{background:#1a73e8;color:white}' +
    '.list{padding:10px;display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}' +
    '.card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 4px rgba(0,0,0,0.08)}' +
    '.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px}' +
    '.card-id{font-weight:bold;color:#1a73e8;font-size:clamp(14px,3vw,16px)}' +
    '.card-status{padding:4px 10px;border-radius:12px;font-size:clamp(10px,2vw,12px);font-weight:bold;background:#e8f0fe}' +
    '.card-row{font-size:clamp(12px,2.5vw,14px);margin:5px 0;color:#666}' +
    '@media (min-width:768px){.list{grid-template-columns:repeat(2,1fr)}}' +
    '@media (min-width:1024px){.list{grid-template-columns:repeat(3,1fr)}}' +
    '</style></head><body>' +
    '<div class="header"><h2>📋 Grievances</h2><input type="text" class="search" placeholder="Search..." oninput="filter(this.value)"></div>' +
    '<div class="filters"><button class="filter active" onclick="filterStatus(\'all\',this)">All</button><button class="filter" onclick="filterStatus(\'Open\',this)">Open</button><button class="filter" onclick="filterStatus(\'Pending Info\',this)">Pending</button><button class="filter" onclick="filterStatus(\'Resolved\',this)">Resolved</button></div>' +
    '<div class="list" id="list"><div style="text-align:center;padding:40px;color:#666;grid-column:1/-1">Loading...</div></div>' +
    '<script>' + getClientSideEscapeHtml() + 'var all=[];google.script.run.withSuccessHandler(function(data){all=data;render(data)}).getRecentGrievancesForMobile(100);function render(data){var c=document.getElementById("list");if(!data||data.length===0){c.innerHTML="<div style=\'text-align:center;padding:40px;color:#999;grid-column:1/-1\'>No grievances</div>";return}c.innerHTML=data.map(function(g){return"<div class=\'card\'><div class=\'card-header\'><div class=\'card-id\'>#"+escapeHtml(g.id)+"</div><div class=\'card-status\'>"+escapeHtml(g.status||"Filed")+"</div></div><div class=\'card-row\'><strong>Member:</strong> "+escapeHtml(g.memberName)+"</div><div class=\'card-row\'><strong>Issue:</strong> "+escapeHtml(g.issueType||"N/A")+"</div><div class=\'card-row\'><strong>Filed:</strong> "+escapeHtml(g.filedDate)+"</div></div>"}).join("")}function filterStatus(s,btn){document.querySelectorAll(".filter").forEach(function(f){f.classList.remove("active")});btn.classList.add("active");render(s==="all"?all:all.filter(function(g){return g.status===s}))}function filter(q){render(all.filter(function(g){q=q.toLowerCase();return g.id.toLowerCase().indexOf(q)>=0||g.memberName.toLowerCase().indexOf(q)>=0||(g.issueType||"").toLowerCase().indexOf(q)>=0}))}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Grievance List');
}

// ============================================================================
// MOBILE UNIFIED SEARCH
// ============================================================================

/**
 * Shows mobile-optimized unified search interface
 * Features:
 * - Tabbed search (All, Members, Grievances)
 * - Real-time search results
 * - Touch-friendly result cards
 * @returns {void}
 */
function showMobileUnifiedSearch() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    getMobileOptimizedHead() +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;margin:0;padding:0;background:#f5f5f5}' +
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:15px}' +
    '.header h2{margin:0 0 12px 0;font-size:clamp(18px,4vw,22px)}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:clamp(12px,3vw,16px) clamp(12px,3vw,16px) clamp(12px,3vw,16px) 45px;border:none;border-radius:10px;font-size:clamp(14px,3vw,16px);background:white}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:18px}' +
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0}' +
    '.tab{flex:1;padding:clamp(10px,2.5vw,14px);text-align:center;font-size:clamp(12px,2.5vw,14px);font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.tab.active{color:#1a73e8;border-bottom-color:#1a73e8}' +
    '.results{padding:10px;display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:10px}' +
    '.result-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 4px rgba(0,0,0,0.08)}' +
    '.result-title{font-weight:bold;color:#1a73e8;margin-bottom:5px;font-size:clamp(14px,3vw,16px)}' +
    '.result-detail{font-size:clamp(11px,2.5vw,13px);color:#666;margin:3px 0}' +
    '.empty-state{text-align:center;padding:60px;color:#999;grid-column:1/-1}' +
    '@media (min-width:768px){.results{grid-template-columns:repeat(2,1fr)}}' +
    '@media (min-width:1024px){.results{grid-template-columns:repeat(3,1fr)}}' +
    '</style></head><body>' +
    '<div class="header"><h2>🔍 Search</h2><div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="q" placeholder="Search members or grievances..." oninput="search(this.value)"></div></div>' +
    '<div class="tabs"><button class="tab active" onclick="setTab(\'all\',this)">All</button><button class="tab" onclick="setTab(\'members\',this)">Members</button><button class="tab" onclick="setTab(\'grievances\',this)">Grievances</button></div>' +
    '<div class="results" id="results"><div class="empty-state">Type to search...</div></div>' +
    '<script>' + getClientSideEscapeHtml() + 'var tab="all";function setTab(t,btn){tab=t;document.querySelectorAll(".tab").forEach(function(tb){tb.classList.remove("active")});btn.classList.add("active");search(document.getElementById("q").value)}function search(q){if(!q||q.length<2){document.getElementById("results").innerHTML="<div class=\'empty-state\'>Type to search...</div>";return}google.script.run.withSuccessHandler(function(data){render(data)}).getMobileSearchData(q,tab)}function render(data){var c=document.getElementById("results");if(!data||data.length===0){c.innerHTML="<div class=\'empty-state\'>No results</div>";return}c.innerHTML=data.map(function(r){return"<div class=\'result-card\'><div class=\'result-title\'>"+(r.type==="member"?"👤 ":"📋 ")+escapeHtml(r.title)+"</div><div class=\'result-detail\'>"+escapeHtml(r.subtitle)+"</div>"+(r.detail?"<div class=\'result-detail\'>"+escapeHtml(r.detail)+"</div>":"")+"</div>"}).join("")}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search');
}

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
        var id = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var name = (row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '');
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
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
        var id = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var name = (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '');
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
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
function showMyAssignedGrievances() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) { SpreadsheetApp.getUi().alert('No grievances found'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var mine = data.filter(function(row) { var steward = row[GRIEVANCE_COLS.STEWARD - 1]; return steward && steward.indexOf(email) >= 0; });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances assigned to you'); return; }
  var msg = 'You have ' + mine.length + ' assigned grievance(s):\n\n';
  mine.slice(0, 10).forEach(function(row) { msg += '#' + row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] + ' - ' + row[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + row[GRIEVANCE_COLS.LAST_NAME - 1] + ' (' + row[GRIEVANCE_COLS.STATUS - 1] + ')\n'; });
  if (mine.length > 10) msg += '\n... and ' + (mine.length - 10) + ' more';
  SpreadsheetApp.getUi().alert('My Cases', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}



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
 * @version 1.0.0
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
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var memberId = data[MEMBER_COLS.MEMBER_ID - 1];
  var name = data[MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[MEMBER_COLS.LAST_NAME - 1];
  var email = data[MEMBER_COLS.EMAIL - 1];
  var hasOpen = data[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1];

  // Build email section buttons (only if email exists)
  var emailButtons = '';
  if (email) {
    emailButtons =
      '<div class="section-header">📨 Email Options</div>' +
      '<button class="action-btn" onclick="google.script.run.composeEmailForMember(\'' + escapeHtml(memberId) + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Send Custom Email</div><div class="desc">Compose email to ' + escapeHtml(email) + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(\'' + escapeHtml(memberId) + '\');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(\'' + escapeHtml(memberId) + '\');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailDashboardLinkToMember(\'' + escapeHtml(memberId) + '\');google.script.host.close()"><span class="icon">🔗</span><span><div class="title">Send Dashboard Link</div><div class="desc">Share dashboard access with member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}' +
    '.container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}' +
    'h2{color:#1a73e8;display:flex;align-items:center;gap:10px;font-size:clamp(16px,4.5vw,20px)}' +
    '.info{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px}' +
    '.name{font-size:clamp(15px,4vw,18px);font-weight:bold}' +
    '.id{color:#666;font-size:clamp(12px,3vw,14px)}' +
    '.status{margin-top:10px}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:clamp(10px,2.8vw,12px);font-weight:bold}' +
    '.open{background:#ffebee;color:#c62828}' +
    '.none{background:#e8f5e9;color:#2e7d32}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:clamp(12px,3vw,15px);border:none;border-radius:8px;cursor:pointer;font-size:clamp(13px,3.5vw,14px);text-align:left;background:#f8f9fa;min-height:44px}' +
    '.action-btn:hover{background:#e8f4fd}' +
    '.icon{font-size:clamp(20px,5vw,24px)}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:clamp(10px,2.8vw,12px);color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#1a73e8;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0;font-size:clamp(12px,3vw,14px)}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer;min-height:44px;font-size:clamp(13px,3.5vw,14px)}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Quick Actions</h2>' +
    '<div class="info">' +
    '<div class="name">' + escapeHtml(name) + '</div>' +
    '<div class="id">' + escapeHtml(memberId) + ' | ' + escapeHtml(email || 'No email') + '</div>' +
    '<div class="status">' + (hasOpen === 'Yes' ? '<span class="badge open">🔴 Has Open Grievance</span>' : '<span class="badge none">🟢 No Open Grievances</span>') + '</div>' +
    '</div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Member Actions</div>' +
    '<button class="action-btn" onclick="google.script.run.openGrievanceFormForMember(' + row + ');google.script.host.close()"><span class="icon">📋</span><span><div class="title">Start New Grievance</div><div class="desc">Create a grievance for this member</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.showMemberGrievanceHistory(\'' + escapeHtml(memberId) + '\');google.script.host.close()"><span class="icon">📁</span><span><div class="title">View Grievance History</div><div class="desc">See all grievances for this member</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(\'' + escapeHtml(memberId) + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Member ID</div><div class="desc">' + escapeHtml(memberId) + '</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(r){if(r.success){alert(r.message+\'\\n\'+r.folderUrl)}else{alert(\'Error: \'+r.error)}}).withFailureHandler(function(e){alert(e.message)}).setupDriveFolderForMember(\'' + escapeHtml(memberId) + '\')"><span class="icon">📁</span><span><div class="title">Create Member Folder</div><div class="desc">Setup Google Drive folder for this member</div></span></button>' +
    emailButtons +
    '</div>' +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(email ? 650 : 400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Quick Actions');
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
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var grievanceId = data[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
  var memberId = data[GRIEVANCE_COLS.MEMBER_ID - 1];
  var name = data[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + data[GRIEVANCE_COLS.LAST_NAME - 1];
  var status = data[GRIEVANCE_COLS.STATUS - 1];
  var step = data[GRIEVANCE_COLS.CURRENT_STEP - 1];
  var daysTo = data[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
  var memberEmail = data[GRIEVANCE_COLS.MEMBER_EMAIL - 1];
  var isOpen = status === 'Open' || status === 'Pending Info' || status === 'In Arbitration' || status === 'Appealed';

  // Build email button (only if member has email)
  var emailStatusBtn = '';
  if (memberEmail) {
    emailStatusBtn =
      '<div class="section-header">📨 Communication</div>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailGrievanceStatusToMember(' + JSON.stringify(grievanceId) + ');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Email Status to Member</div><div class="desc">Send grievance status update to ' + escapeHtml(String(memberEmail)) + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(' + JSON.stringify(memberId) + ');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(' + JSON.stringify(memberId) + ');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}' +
    '.container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}' +
    'h2{color:#DC2626;font-size:clamp(16px,4.5vw,20px)}' +
    '.info{background:#fff5f5;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;border-left:4px solid #DC2626}' +
    '.gid{font-size:clamp(15px,4vw,18px);font-weight:bold}' +
    '.gmem{color:#666;font-size:clamp(12px,3vw,14px)}' +
    '.gstatus{margin-top:10px;display:flex;gap:8px;flex-wrap:wrap}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:clamp(10px,2.8vw,12px);font-weight:bold}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:clamp(12px,3vw,15px);border:none;border-radius:8px;cursor:pointer;font-size:clamp(13px,3.5vw,14px);text-align:left;background:#f8f9fa;min-height:44px}' +
    '.action-btn:hover{background:#fff5f5}' +
    '.icon{font-size:clamp(20px,5vw,24px)}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:clamp(10px,2.8vw,12px);color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#DC2626;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0;font-size:clamp(12px,3vw,14px)}' +
    '.divider{height:1px;background:#e0e0e0;margin:10px 0}' +
    '.status-section{margin-top:15px;padding:clamp(10px,3vw,15px);background:#f8f9fa;border-radius:8px}' +
    '.status-section h4{margin:0 0 10px;font-size:clamp(13px,3.5vw,15px)}' +
    'select{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:16px;min-height:44px}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer;min-height:44px;font-size:clamp(13px,3.5vw,14px)}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Grievance Actions</h2>' +
    '<div class="info">' +
    '<div class="gid">' + escapeHtml(grievanceId) + '</div>' +
    '<div class="gmem">' + escapeHtml(name) + ' (' + escapeHtml(memberId) + ')' + (memberEmail ? ' - ' + escapeHtml(memberEmail) : '') + '</div>' +
    '<div class="gstatus">' +
    '<span class="badge">' + escapeHtml(status) + '</span>' +
    '<span class="badge">' + escapeHtml(step) + '</span>' +
    (daysTo !== null && daysTo !== '' ? '<span class="badge" style="background:' + (daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0) ? '#ffebee;color:#c62828' : '#e3f2fd;color:#1565c0') + '">' + (daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0) ? '⚠️ Overdue' : '📅 ' + daysTo + ' days') + '</span>' : '') +
    '</div></div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Case Management</div>' +
    '<button class="action-btn" onclick="google.script.run.syncSingleGrievanceToCalendar(' + JSON.stringify(grievanceId) + ');google.script.host.close()"><span class="icon">📅</span><span><div class="title">Sync to Calendar</div><div class="desc">Add deadlines to Google Calendar</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.setupDriveFolderForGrievance(' + JSON.stringify(grievanceId) + ');google.script.host.close()"><span class="icon">📁</span><span><div class="title">Setup Drive Folder</div><div class="desc">Create document folder</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(' + JSON.stringify(grievanceId) + ');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Grievance ID</div><div class="desc">' + escapeHtml(grievanceId) + '</div></span></button>' +
    emailStatusBtn +
    '</div>' +
    (isOpen ? '<div class="status-section"><h4>Quick Status Update</h4><select id="statusSelect"><option value="">-- Select --</option><option value="Open">Open</option><option value="Pending Info">Pending Info</option><option value="Settled">Settled</option><option value="Withdrawn">Withdrawn</option><option value="Won">Won</option><option value="Denied">Denied</option><option value="Closed">Closed</option></select><button class="action-btn" style="margin-top:10px" onclick="var s=document.getElementById(\'statusSelect\').value;if(!s){alert(\'Select status\');return}google.script.run.withSuccessHandler(function(){alert(\'Updated!\');google.script.host.close()}).quickUpdateGrievanceStatus(' + row + ',s)"><span class="icon">✓</span><span><div class="title">Update Status</div></span></button></div>' : '') +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(memberEmail ? 750 : 550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance Quick Actions');
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
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      var email = data[i][MEMBER_COLS.EMAIL - 1];
      var name = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      if (!email) { SpreadsheetApp.getUi().alert('No email on file.'); return; }
      var html = HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px);background:#f5f5f5}.container{background:white;padding:clamp(15px,4vw,25px);border-radius:8px}h2{color:#1a73e8;font-size:clamp(16px,4.5vw,20px)}.info{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;font-size:clamp(12px,3vw,14px)}.form-group{margin:15px 0}label{display:block;font-weight:bold;margin-bottom:5px;font-size:clamp(12px,3vw,14px)}input,textarea{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:16px;box-sizing:border-box;min-height:44px}textarea{min-height:180px}input:focus,textarea:focus{outline:none;border-color:#1a73e8}.buttons{display:flex;gap:10px;margin-top:20px}button{padding:12px 24px;font-size:clamp(13px,3.5vw,14px);border:none;border-radius:4px;cursor:pointer;flex:1;min-height:44px}.primary{background:#1a73e8;color:white}.secondary{background:#6c757d;color:white}@media(max-width:480px){.buttons{flex-direction:column}}</style></head><body><div class="container"><h2>📧 Email to Member</h2><div class="info"><strong>' + escapeHtml(name) + '</strong> (' + escapeHtml(memberId) + ')<br>' + escapeHtml(email) + '</div><div class="form-group"><label>Subject:</label><input type="text" id="subject" placeholder="Email subject"></div><div class="form-group"><label>Message:</label><textarea id="message" placeholder="Type your message..."></textarea></div><div class="buttons"><button class="primary" onclick="send()">📤 Send</button><button class="secondary" onclick="google.script.host.close()">Cancel</button></div></div><script>function send(){var s=document.getElementById("subject").value.trim();var m=document.getElementById("message").value.trim();if(!s||!m){alert("Fill in subject and message");return}google.script.run.withSuccessHandler(function(){alert("Email sent!");google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).sendQuickEmail(' + JSON.stringify(email) + ',s,m,' + JSON.stringify(memberId) + ')}</script></body></html>'
      ).setWidth(600).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(html, '📧 Compose Email');
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
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: getOrgNameFromConfig_() + ' Dashboard' });
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
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      grievance = {
        id: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        memberId: data[i][GRIEVANCE_COLS.MEMBER_ID - 1],
        firstName: data[i][GRIEVANCE_COLS.FIRST_NAME - 1],
        lastName: data[i][GRIEVANCE_COLS.LAST_NAME - 1],
        status: data[i][GRIEVANCE_COLS.STATUS - 1],
        step: data[i][GRIEVANCE_COLS.CURRENT_STEP - 1],
        dateFiled: data[i][GRIEVANCE_COLS.DATE_FILED - 1],
        nextAction: data[i][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1],
        daysToDeadline: data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
        issueCategory: data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
        steward: data[i][GRIEVANCE_COLS.STEWARD - 1],
        email: data[i][GRIEVANCE_COLS.MEMBER_EMAIL - 1]
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
    'Date Filed: ' + (grievance.dateFiled ? new Date(grievance.dateFiled).toLocaleDateString() : 'N/A') + '\n';

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
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      return {
        memberId: memberId,
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1],
        lastName: data[i][MEMBER_COLS.LAST_NAME - 1],
        email: data[i][MEMBER_COLS.EMAIL - 1]
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
    if (row[GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
      mine.push({ id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1], status: row[GRIEVANCE_COLS.STATUS - 1], step: row[GRIEVANCE_COLS.CURRENT_STEP - 1], issue: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1], filed: row[GRIEVANCE_COLS.DATE_FILED - 1], closed: row[GRIEVANCE_COLS.DATE_CLOSED - 1] });
    }
  });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances for this member.'); return; }
  var list = mine.map(function(g) {
    return '<div style="background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;border-left:4px solid ' + (g.status === 'Open' ? '#f44336' : '#4caf50') + '"><strong>' + escapeHtml(g.id) + '</strong><br><span style="color:#666">Status: ' + escapeHtml(g.status) + ' | Step: ' + escapeHtml(g.step) + '</span><br><span style="color:#888;font-size:12px">' + escapeHtml(g.issue) + ' | Filed: ' + (g.filed ? new Date(g.filed).toLocaleDateString() : 'N/A') + '</span></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:clamp(12px,3vw,20px)}h2{color:#1a73e8;font-size:clamp(16px,4.5vw,20px)}.summary{background:#e8f4fd;padding:clamp(10px,3vw,15px);border-radius:8px;margin-bottom:20px;font-size:clamp(12px,3vw,14px)}</style></head><body><h2>📁 Grievance History</h2><div class="summary"><strong>Member ID:</strong> ' + escapeHtml(memberId) + '<br><strong>Total:</strong> ' + mine.length + '<br><strong>Open:</strong> ' + mine.filter(function(g) { return g.status === 'Open'; }).length + '<br><strong>Closed:</strong> ' + mine.filter(function(g) { return g.status !== 'Open'; }).length + '</div>' + list + '</body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance History - ' + memberId);
}

/**
 * Opens grievance form for a member (placeholder - directs to Grievance Log)
 * @param {number} row - The member row number
 */
function openGrievanceFormForMember(row) {
  SpreadsheetApp.getUi().alert('ℹ️ New Grievance', 'To start a new grievance for this member, navigate to the Grievance Log sheet and add a new row with their Member ID.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Syncs a single grievance to the calendar
 * @param {string} grievanceId - The grievance ID to sync
 */
function syncSingleGrievanceToCalendar(grievanceId) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing ' + grievanceId + '...', 'Calendar', 3);
  if (typeof syncDeadlinesToCalendar === 'function') syncDeadlinesToCalendar(grievanceId);
}

// ============================================================================
// DASHBOARD LINK AND STEWARD TOOLKIT EMAILS
// ============================================================================

/**
 * Prompts for email and sends member dashboard link
 * Opens a prompt dialog to enter email address and sends the dashboard URL
 */
function sendMemberDashboardLink() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = ss.getUrl();

  var response = ui.prompt('Send Report', 'Enter Member Email:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    var email = response.getResponseText();

    if (!email || !email.includes('@')) {
      ui.alert('Please enter a valid email address.');
      return;
    }

    var body = "Hello,\n\n" +
      "Access your Union Member Dashboard:\n" + url + "\n\n" +
      "HOW TO ACCESS:\n" +
      "1. Open the spreadsheet using the link above\n" +
      "2. Go to: Command > Command Center > Member Dashboard (No PII)\n" +
      "3. The dashboard shows aggregate union metrics (no personal info visible)\n\n" +
      "WHAT YOU'LL SEE:\n" +
      "- Morale & Trust Scores\n" +
      "- Leadership Pipeline\n" +
      "- Grievance Statistics\n" +
      "- Steward Contact Search\n" +
      "- Emergency Weingarten Rights\n\n" +
      "In Solidarity,\n" +
      COMMAND_CONFIG.SYSTEM_NAME;

    try {
      MailApp.sendEmail(email, COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + " Your Union Dashboard Access", body);
      ui.alert('Dashboard access link sent to ' + email);
    } catch (e) {
      ui.alert('Error sending email: ' + e.message);
    }
  }
}

/**
 * Sends the Member Dashboard URL to the selected member from Member Directory.
 * Uses the currently selected row to get member email and name.
 * This is a PII-protected view link.
 * NOTE: Duplicate exists in 11_SecureMemberDashboard.gs - this version kept for compatibility
 * @deprecated Use emailDashboardLink() in 11_SecureMemberDashboard.gs
 */
function emailDashboardLink_UIService_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveRange().getRow();

  // Validate we're on Member Directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR || row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a member row in the Member Directory first.');
    return;
  }

  // Get member email and name from the selected row
  var email = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();

  if (!email || !email.toString().includes('@')) {
    SpreadsheetApp.getUi().alert('No valid email found for this member.');
    return;
  }

  // Get organization name from config if available
  var orgName = '';
  try {
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var configOrgName = configSheet.getRange(2, CONFIG_COLS.ORG_NAME).getValue();
      if (configOrgName) orgName = configOrgName;
    }
  } catch (_e) {
    // Use default org name
  }

  // Build email body
  var dashboardUrl = ss.getUrl();
  var body = 'Hi ' + firstName + ',\n\n' +
    'You can view current union stats and representation here:\n' +
    dashboardUrl + '\n\n' +
    'This is a PII-protected view showing only aggregate union statistics.\n\n' +
    'From the dashboard you can:\n' +
    '- View active grievance counts and outcomes\n' +
    '- See member satisfaction trends\n' +
    '- Find your steward contact information\n' +
    '- Track union coverage and goals\n\n' +
    'If you have questions about your specific case or concerns, ' +
    'please contact your assigned steward directly.\n\n' +
    'In Solidarity,\n' +
    orgName + ' Union Leadership';

  try {
    MailApp.sendEmail({
      to: email,
      subject: orgName + ' - Your Member Dashboard Access',
      body: body,
      name: orgName + ' Union'
    });
    ss.toast('Sent dashboard link to ' + email, 'Success', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + e.message);
  }
}

/**
 * Sends the steward toolkit welcome email to a newly promoted steward
 * @private
 * @param {string} email - The steward's email address
 * @param {string} name - The steward's name
 */
function sendStewardToolkit_(email, name) {
  try {
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Welcome to the Steward Team!';
    var body = 'Congratulations ' + name + '!\n\n' +
      'You have been promoted to Union Steward. Welcome to the leadership team!\n\n' +
      'STEWARD TOOLKIT:\n' +
      '1. Access the Grievance Log to track and manage cases\n' +
      '2. Use the Executive Command dashboard for your analytics\n' +
      '3. Review the Member Directory for your assigned members\n' +
      '4. Contact your Chief Steward for mentorship and guidance\n\n' +
      'KEY RESPONSIBILITIES:\n' +
      '- Represent members in grievance proceedings\n' +
      '- Document workplace issues and contract violations\n' +
      '- Maintain confidentiality of member information\n' +
      '- Attend steward training and meetings\n\n' +
      'We are stronger together!' +
      COMMAND_CONFIG.EMAIL.FOOTER;

    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('Error sending steward toolkit: ' + e.message);
  }
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
 * @version 1.0.0
 */

// ============================================================================
// SEARCH DIALOG LAUNCHERS
// ============================================================================

/**
 * Shows desktop search dialog with full interface
 */
function showDesktopSearch() {
  const html = HtmlService.createHtmlOutput(getDesktopSearchHtml())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Dashboard Search');
}

/**
 * Shows quick search dialog (minimal interface)
 */
function showQuickSearch() {
  const html = HtmlService.createHtmlOutput(getQuickSearchHtml())
    .setWidth(DIALOG_SIZES.SMALL.width)
    .setHeight(DIALOG_SIZES.SMALL.height);

  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Search');
}

/**
 * Shows advanced search with complex filtering options
 */
function showAdvancedSearch() {
  const html = HtmlService.createHtmlOutput(getAdvancedSearchHtml())
    .setWidth(DIALOG_SIZES.FULLSCREEN.width)
    .setHeight(DIALOG_SIZES.FULLSCREEN.height);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Advanced Search');
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
