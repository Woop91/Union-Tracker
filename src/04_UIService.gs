/**
 * ============================================================================
 * UIService.gs - Unified Interface Logic
 * ============================================================================
 *
 * This module handles all user interface components including:
 * - Search dialogs (desktop and mobile)
 * - Multi-select editors
 * - Quick action menus
 * - Sidebars and modal dialogs
 * - Theme management
 * - HTML generation utilities
 *
 * SEPARATION OF CONCERNS: This file contains ONLY UI/presentation logic.
 * All business logic should remain in their respective modules.
 *
 * @fileoverview UI components and dialog management
 * @version 2.0.0
 * @requires Constants.gs
 */

// ============================================================================
// MENU CREATION
// ============================================================================

/**
 * Creates the custom menu when the spreadsheet opens
 * Uses consistent iconography: 📊 for views, ➕ for actions
 * Called automatically by onOpen trigger
 */
function createDashboardMenu() {
  const ui = SpreadsheetApp.getUi();

  // Main Menu: Union Hub
  ui.createMenu('📊 Union Hub')
    .addItem('📊 Dashboard Home', 'showDashboardSidebar')
    .addItem('📱 Toggle Mobile View', 'toggleMobileView')
    .addItem('🎛️ Visual Control Panel', 'showVisualControlPanel')
    .addSeparator()

    // Search submenu - Action icons
    .addSubMenu(ui.createMenu('🔍 Search')
      .addItem('🔍 Desktop Search', 'showDesktopSearch')
      .addItem('⚡ Quick Search', 'showQuickSearch')
      .addItem('🔎 Advanced Search', 'showAdvancedSearch'))

    // Grievances submenu - Mixed icons
    .addSubMenu(ui.createMenu('📋 Cases & Grievances')
      .addItem('➕ New Case/Grievance', 'showNewGrievanceDialog')
      .addItem('✏️ Edit Selected', 'showEditGrievanceDialog')
      .addItem('✅ View Checklist', 'showChecklistDialog')
      .addSeparator()
      .addItem('🔄 Bulk Update Status', 'showBulkStatusUpdate')
      .addItem('📈 Case Analytics', 'showInteractiveDashboardTab'))

    // Members submenu - Mixed icons
    .addSubMenu(ui.createMenu('👥 Members')
      .addItem('➕ Add New Member', 'showNewMemberDialog')
      .addItem('🔍 Find Existing Member', 'showSearchDialog')
      .addItem('📥 Import Members', 'showImportMembersDialog')
      .addItem('📤 Export Directory', 'showExportMembersDialog')
      .addSeparator()
      .addItem('🛡️ Steward Directory', 'showStewardDirectory')
      .addItem('🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas')
      .addSeparator()
      .addItem('📧 Send Contact Form', 'sendContactInfoForm')
      .addItem('📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink')
      .addSeparator()
      .addSubMenu(ui.createMenu('🆔 ID & Data Management')
        .addItem('🆔 Generate Missing Member IDs', 'generateMissingMemberIDs')
        .addItem('🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs')))

    // Calendar submenu - Action icons
    .addSubMenu(ui.createMenu('📅 Calendar')
      .addItem('🔗 Sync Deadlines to Calendar', 'syncDeadlinesToCalendar')
      .addItem('👁️ View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar')
      .addItem('🗑️ Clear All Calendar Events', 'clearAllCalendarEvents'))

    // Drive submenu - Folder icons
    .addSubMenu(ui.createMenu('📁 Google Drive')
      .addItem('📁 Setup Folder for Grievance', 'setupFolderForSelectedGrievance')
      .addItem('📁 View Grievance Files', 'showGrievanceFiles')
      .addItem('📁 Batch Create Folders', 'batchCreateAllMissingFolders'))

    // Notifications submenu
    .addSubMenu(ui.createMenu('🔔 Notifications')
      .addItem('⚙️ Notification Settings', 'showNotificationSettings')
      .addItem('🧪 Test Notifications', 'testDeadlineNotifications'))

    // Dashboards - Two unified dashboards (v4.3.2)
    .addSubMenu(ui.createMenu('📊 Dashboards')
      .addItem('👥 Member Dashboard (No PII)', 'showPublicMemberDashboard')
      .addItem('🛡️ Steward Dashboard (PII)', 'showStewardDashboard')
      .addSeparator()
      .addItem('📱 Mobile Dashboard', 'showMobileDashboard'))
    .addSubMenu(ui.createMenu('🔄 Maintenance')
      .addItem('🔄 Refresh All Formulas', 'refreshAllFormulas')
      .addItem('📱 Get Mobile App URL', 'showWebAppUrl'))

    // Appearance & Comfort View consolidated into Visual Control Panel
    .addItem('🎛️ Visual Control Panel', 'showVisualControlPanel')

    // Multi-Select Tools submenu
    .addSubMenu(ui.createMenu('📝 Multi-Select')
      .addItem('📝 Open Editor', 'openCellMultiSelectEditor')
      .addItem('⚡ Enable Auto-Open', 'installMultiSelectTrigger')
      .addItem('🚫 Disable Auto-Open', 'removeMultiSelectTrigger'))

    .addSeparator()
    .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
    .addItem('📖 Help & Documentation', 'showHelpDialog')
    .addToUi();

  // Admin Menu (Separate top-level menu)
  var adminMenu = ui.createMenu('🛠️ Admin')
    .addItem('🩺 System Diagnostics', 'showDiagnosticsDialog')
    .addItem('🔍 Modal Diagnostics', 'showModalDiagnostics')
    .addItem('🔧 Repair Dashboard', 'showRepairDialog')
    .addItem('⚙️ Settings', 'showSettingsDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Automation')
      .addItem('🔄 Force Global Refresh', 'refreshAllVisuals')
      .addItem('🌙 Enable Midnight Auto-Refresh', 'setupMidnightTrigger')
      .addItem('❌ Disable Midnight Auto-Refresh', 'removeMidnightTrigger')
      .addItem('🔔 Enable 1AM Dashboard Refresh', 'createAutomationTriggers')
      .addItem('📑 Email Weekly PDF Snapshot', 'emailExecutivePDF'))
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
    .addSubMenu(ui.createMenu('🗄️ Cache & Performance')
      .addItem('🗄️ Cache Status', 'showCacheStatusDashboard')
      .addItem('🔥 Warm Up Caches', 'warmUpCaches')
      .addItem('🗑️ Clear All Caches', 'invalidateAllCaches'))
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
      .addItem('🎨 Apply Tab Colors', 'applyTabColors')
      .addItem('📱 Add Mobile Dashboard Link to Config', 'addMobileDashboardLinkToConfig')
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

  // Strategic Operations Menu (Simplified v4.3.2)
  ui.createMenu('🎯 Strategic Ops')
    .addItem('👥 Member Dashboard', 'showPublicMemberDashboard')
    .addItem('🛡️ Steward Dashboard', 'showStewardDashboard')
    .addSeparator()
    .addItem('🔍 Desktop Search', 'showDesktopSearch')
    .addSubMenu(ui.createMenu('📋 Cases & Grievances')
      .addItem('➕ New Case/Grievance', 'showNewGrievanceDialog')
      .addItem('✏️ Edit Selected', 'showEditGrievanceDialog')
      .addItem('✅ View Checklist', 'showChecklistDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🆔 ID & Data Engines')
      .addItem('🆔 Generate Missing Member IDs', 'generateMissingMemberIDs')
      .addItem('🔍 Check Duplicate IDs', 'checkDuplicateMemberIDs')
      .addItem('📄 Create PDF for Selected Grievance', 'createPDFForSelectedGrievance'))
    .addSubMenu(ui.createMenu('👤 Steward Management')
      .addItem('⬆️ Promote to Steward', 'promoteSelectedMemberToSteward')
      .addItem('⬇️ Demote Steward', 'demoteSelectedSteward')
      .addItem('📧 Send Contact Form', 'sendContactInfoForm')
      .addItem('📊 Send Satisfaction Survey', 'getSatisfactionSurveyLink'))
    .addToUi();

  // Create the Field Portal menu (from CommandCenter.gs)
  createCommandCenterMenu();
}

/**
 * Shows the Visual Control Panel sidebar
 * Allows users to toggle visual settings like Dark Mode, Focus Mode, Zebra Stripes
 */
function showVisualControlPanel() {
  const html = HtmlService.createHtmlOutput(getVisualControlPanelHtml())
    .setTitle('🎛️ Visual Control Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Generates HTML for the Visual Control Panel sidebar
 * @return {string} HTML content
 */
function getVisualControlPanelHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
          font-family: 'Google Sans', -apple-system, BlinkMacSystemFont, sans-serif;
          background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
          min-height: 100vh;
          color: #F8FAFC;
        }
        .panel {
          padding: 20px;
        }
        .panel-header {
          text-align: center;
          margin-bottom: 25px;
        }
        .panel-header h1 {
          font-size: 20px;
          font-weight: 600;
          margin-bottom: 5px;
        }
        .panel-header p {
          font-size: 12px;
          color: #94A3B8;
        }
        .section {
          background: rgba(255,255,255,0.05);
          border-radius: 12px;
          padding: 15px;
          margin-bottom: 15px;
        }
        .section-title {
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 1px;
          color: #94A3B8;
          margin-bottom: 12px;
        }
        .toggle-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 10px 0;
          border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        .toggle-row:last-child { border-bottom: none; }
        .toggle-label {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .toggle-label span { font-size: 18px; }
        .toggle-label div { font-size: 13px; }
        .toggle-switch {
          position: relative;
          width: 50px;
          height: 26px;
        }
        .toggle-switch input {
          opacity: 0;
          width: 0;
          height: 0;
        }
        .slider {
          position: absolute;
          cursor: pointer;
          top: 0; left: 0; right: 0; bottom: 0;
          background: #475569;
          border-radius: 26px;
          transition: 0.3s;
        }
        .slider:before {
          position: absolute;
          content: "";
          height: 20px;
          width: 20px;
          left: 3px;
          bottom: 3px;
          background: white;
          border-radius: 50%;
          transition: 0.3s;
        }
        input:checked + .slider {
          background: #7C3AED;
        }
        input:checked + .slider:before {
          transform: translateX(24px);
        }
        .theme-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 10px;
          margin-top: 10px;
        }
        .theme-btn {
          padding: 12px;
          border-radius: 8px;
          border: 2px solid transparent;
          cursor: pointer;
          text-align: center;
          font-size: 12px;
          transition: all 0.2s;
        }
        .theme-btn:hover { transform: scale(1.02); }
        .theme-btn.active { border-color: #7C3AED; }
        .theme-default { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }
        .theme-dark { background: #1E293B; color: white; border: 1px solid #334155; }
        .theme-light { background: #F8FAFC; color: #1E293B; }
        .theme-contrast { background: #000; color: #FFD700; }
        .action-btn {
          width: 100%;
          padding: 12px;
          border: none;
          border-radius: 8px;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
          margin-top: 10px;
          transition: all 0.2s;
        }
        .btn-primary {
          background: linear-gradient(135deg, #7C3AED, #5B21B6);
          color: white;
        }
        .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(124,58,237,0.4); }
        .btn-secondary {
          background: rgba(255,255,255,0.1);
          color: white;
        }
        .btn-secondary:hover { background: rgba(255,255,255,0.15); }
        .status-bar {
          margin-top: 20px;
          padding: 10px;
          background: rgba(16,185,129,0.2);
          border-radius: 8px;
          text-align: center;
          font-size: 12px;
          color: #10B981;
        }
      </style>
    </head>
    <body>
      <div class="panel">
        <div class="panel-header">
          <h1>🎛️ Visual Control Panel</h1>
          <p>Customize your dashboard experience</p>
        </div>

        <div class="section">
          <div class="section-title">Display Settings</div>
          <div class="toggle-row">
            <div class="toggle-label">
              <span>🌙</span>
              <div>Dark Mode</div>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="darkMode" onchange="toggleSetting('darkMode')">
              <span class="slider"></span>
            </label>
          </div>
          <div class="toggle-row">
            <div class="toggle-label">
              <span>🎯</span>
              <div>Focus Mode</div>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="focusMode" onchange="toggleSetting('focusMode')">
              <span class="slider"></span>
            </label>
          </div>
          <div class="toggle-row">
            <div class="toggle-label">
              <span>🦓</span>
              <div>Zebra Stripes</div>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="zebraStripes" checked onchange="toggleSetting('zebraStripes')">
              <span class="slider"></span>
            </label>
          </div>
          <div class="toggle-row">
            <div class="toggle-label">
              <span>📊</span>
              <div>Show Charts</div>
            </div>
            <label class="toggle-switch">
              <input type="checkbox" id="showCharts" checked onchange="toggleSetting('showCharts')">
              <span class="slider"></span>
            </label>
          </div>
        </div>

        <div class="section">
          <div class="section-title">Theme Selection</div>
          <div class="theme-grid">
            <div class="theme-btn theme-default active" onclick="selectTheme('default')">
              🎨 Default
            </div>
            <div class="theme-btn theme-dark" onclick="selectTheme('dark')">
              🌙 Dark
            </div>
            <div class="theme-btn theme-light" onclick="selectTheme('light')">
              ☀️ Light
            </div>
            <div class="theme-btn theme-contrast" onclick="selectTheme('contrast')">
              👁️ High Contrast
            </div>
          </div>
        </div>

        <div class="section">
          <div class="section-title">Quick Actions</div>
          <button class="action-btn btn-primary" onclick="refreshDashboard()">
            🔄 Refresh Dashboard
          </button>
          <button class="action-btn btn-secondary" onclick="goToSheet('💼 Dashboard')">
            📊 Open Executive Dashboard
          </button>
          <button class="action-btn btn-secondary" onclick="goToSheet('🎯 Custom View')">
            🎯 Open Custom View
          </button>
        </div>

        <div class="status-bar" id="statusBar">
          ✅ Settings saved automatically
        </div>
      </div>

      <script>
        function toggleSetting(setting) {
          const isChecked = document.getElementById(setting).checked;
          google.script.run
            .withSuccessHandler(showStatus)
            .saveVisualSetting(setting, isChecked);
        }

        function selectTheme(theme) {
          document.querySelectorAll('.theme-btn').forEach(btn => btn.classList.remove('active'));
          event.target.classList.add('active');
          google.script.run
            .withSuccessHandler(showStatus)
            .applyDashboardTheme(theme);
        }

        function refreshDashboard() {
          showStatus('Refreshing...');
          google.script.run
            .withSuccessHandler(function() { showStatus('Dashboard refreshed!'); })
            .syncAllDashboardData();
        }

        function goToSheet(sheetName) {
          google.script.run.navigateToSheet(sheetName);
        }

        function showStatus(msg) {
          const bar = document.getElementById('statusBar');
          bar.textContent = '✅ ' + (msg || 'Settings applied');
          bar.style.background = 'rgba(16,185,129,0.2)';
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * Saves a visual setting to user properties
 * @param {string} setting - Setting name
 * @param {boolean} value - Setting value
 */
function saveVisualSetting(setting, value) {
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty('visual_' + setting, value.toString());
  return 'Setting saved';
}

/**
 * Applies a dashboard theme
 * @param {string} theme - Theme name (default, dark, light, contrast)
 */
function applyDashboardTheme(theme) {
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty('dashboard_theme', theme);
  // Theme application would be done on next dashboard refresh
  return 'Theme set to ' + theme;
}

/**
 * Helper functions for navigation
 * Note: Dashboards are now modal-based for better UX
 */
function showExecutiveDashboard() {
  // Redirects to Interactive Dashboard modal (was: sheet navigation)
  showInteractiveDashboardTab();
}

/**
 * @deprecated v4.3.8 - Sheet is now hidden. Modal version in 08_Code.gs is used instead.
 * Keeping for backward compatibility - redirects to modal.
 */
function showSatisfactionDashboard_DEPRECATED() {
  // The modal version showSatisfactionDashboard() in 08_Code.gs handles this now
  // This function is deprecated - the sheet is hidden
}

function showStewardDirectory() {
  // Navigate to Member Directory filtered by stewards
  navigateToSheet(SHEETS.MEMBER_DIR);
  SpreadsheetApp.getActive().toast('Filter by "Is Steward = Yes" to see steward directory', 'Steward Directory', 5);
}

function toggleDarkMode() {
  SpreadsheetApp.getActive().toast('Dark Mode toggle - use Visual Control Panel for full options', 'Theme', 3);
  showVisualControlPanel();
}

function showThemeSettings() {
  showVisualControlPanel();
}

// ============================================================================
// NAVIGATION HELPERS (Strategic Command Center)
// ============================================================================

/**
 * Navigate to Executive Dashboard (now launches modal)
 */
function navToDash() {
  // Launches Interactive Dashboard modal (was: sheet navigation)
  showInteractiveDashboardTab();
}

/**
 * Navigate to Mobile View (if exists)
 */
function navToMobile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mobileSheet = ss.getSheetByName('📱 Mobile View');

  if (mobileSheet) {
    mobileSheet.activate();
  } else {
    // Fall back to showing the mobile-optimized sidebar
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Mobile View sheet not found. Use the web app for best mobile experience.',
      COMMAND_CONFIG.SYSTEM_NAME,
      5
    );
  }
}

/**
 * Navigate to a specific sheet by name
 * @param {string} sheetName - The sheet name to navigate to
 */
function navigateToSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.activate();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Sheet "' + sheetName + '" not found',
      'Navigation Error',
      3
    );
  }
}

// ============================================================================
// GLOBAL STYLING (Strategic Command Center)
// ============================================================================

/**
 * Applies the system theme to all visible sheets
 * Includes header styling, zebra striping, and font standardization
 */
function APPLY_SYSTEM_THEME() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetsStyled = 0;

  ss.toast('Applying theme to all sheets...', COMMAND_CONFIG.SYSTEM_NAME, 10);

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();

    // Skip hidden sheets
    if (sheet.isSheetHidden()) return;

    // Skip sheets starting with underscore (calculation sheets)
    if (sheetName.startsWith('_')) return;

    try {
      applyThemeToSheet_(sheet);
      sheetsStyled++;
    } catch (e) {
      Logger.log('Failed to style sheet ' + sheetName + ': ' + e.message);
    }
  });

  ss.toast('Theme applied to ' + sheetsStyled + ' sheets', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Applies theme styling to a single sheet
 * For data sheets (Member Directory, Grievance Log), applies to ALL rows in the sheet
 * @param {Sheet} sheet - The sheet to style
 * @private
 */
function applyThemeToSheet_(sheet) {
  var sheetName = sheet.getName();
  var lastCol = sheet.getLastColumn();

  if (lastCol < 1) lastCol = 26; // Default columns for empty sheets

  // For data sheets, style ALL rows using getMaxRows() for consistent appearance
  var isDataSheet = (sheetName === SHEETS.MEMBER_DIR || sheetName === SHEETS.GRIEVANCE_LOG);
  var rowsToStyle = isDataSheet ? sheet.getMaxRows() : sheet.getLastRow();

  if (rowsToStyle < 1) return;

  // Style header row
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
             .setFontColor(COMMAND_CONFIG.THEME.HEADER_TEXT)
             .setFontWeight('bold')
             .setFontFamily(COMMAND_CONFIG.THEME.FONT)
             .setFontSize(11)
             .setVerticalAlignment('middle');

  // Apply font to all data rows
  if (rowsToStyle > 1) {
    var dataRange = sheet.getRange(2, 1, rowsToStyle - 1, lastCol);
    dataRange.setFontFamily(COMMAND_CONFIG.THEME.FONT)
             .setFontSize(COMMAND_CONFIG.THEME.FONT_SIZE)
             .setVerticalAlignment('middle');

    // Use banding for efficient zebra striping on data sheets
    if (isDataSheet) {
      var bandings = dataRange.getBandings();
      if (bandings.length > 0) bandings[0].remove();
      dataRange.applyRowBanding()
        .setHeaderRowColor(null)
        .setFirstRowColor('#ffffff')
        .setSecondRowColor(COMMAND_CONFIG.THEME.ALT_ROW || '#f8f9fa');
    } else {
      // Row-by-row for other sheets (smaller data sets)
      var actualLastRow = sheet.getLastRow();
      for (var row = 2; row <= actualLastRow; row++) {
        var rowRange = sheet.getRange(row, 1, 1, lastCol);
        if (row % 2 === 0) {
          rowRange.setBackground(COMMAND_CONFIG.THEME.ALT_ROW);
        } else {
          rowRange.setBackground(null);
        }
      }
    }
  }

  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Applies global styling to all visible sheets (alias for APPLY_SYSTEM_THEME)
 */
function applyGlobalStyling() {
  APPLY_SYSTEM_THEME();
}

/**
 * Resets all visible sheets to default styling
 */
function resetToDefaultTheme() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    'Reset Theme',
    'This will reset all sheets to default styling (white background, black text).\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.isSheetHidden()) return;
    if (sheet.getName().startsWith('_')) return;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) return;

    try {
      var allRange = sheet.getRange(1, 1, lastRow, lastCol);
      allRange.setBackground(null)
              .setFontColor(null)
              .setFontWeight('normal')
              .setFontFamily('Arial')
              .setFontSize(10);

      // Keep headers bold
      sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');
    } catch (e) {
      Logger.log('Failed to reset sheet ' + sheet.getName() + ': ' + e.message);
    }
  });

  ss.toast('Theme reset to defaults', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

/**
 * Refreshes visual elements - simple version
 * NOTE: Renamed to avoid duplicate. Use refreshAllVisuals() defined later in this file for full refresh.
 * @deprecated Use refreshAllVisuals() at line ~6399
 */
function refreshVisualsSimple_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Refreshing visuals and data...', COMMAND_CONFIG.SYSTEM_NAME, 10);

  // Refresh hidden calculation sheets if available
  try {
    if (typeof rebuildAllHiddenSheets === 'function') {
      rebuildAllHiddenSheets();
    }
  } catch (e) {
    Logger.log('Hidden sheets refresh error: ' + e.message);
  }

  // Apply traffic light indicators
  try {
    applyTrafficLightIndicators();
  } catch (e) {
    Logger.log('Traffic lights error: ' + e.message);
  }

  // Apply global theme
  try {
    APPLY_SYSTEM_THEME();
  } catch (e) {
    Logger.log('Theme error: ' + e.message);
  }

  // Check for alerts
  try {
    checkDashboardAlerts();
  } catch (e) {
    Logger.log('Alert check error: ' + e.message);
  }

  ss.toast('Refresh complete', COMMAND_CONFIG.SYSTEM_NAME, 5);
}

// ============================================================================
// SEARCH DIALOGS
// ============================================================================

/**
 * Shows the desktop search dialog with full functionality
 * Includes member search, grievance search, and filters
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
            html += '<div class="result-item" onclick="selectResult(\\'' + item.id + '\\', \\'' + item.type + '\\')">';
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
            '<div class="no-results"><p>Error: ' + error.message + '</p></div>';
        }

        function escapeHtml(text) {
          const div = document.createElement('div');
          div.textContent = text;
          return div.innerHTML;
        }

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
                   r.title + ' <small style="color:#666">(' + r.type + ')</small></div>';
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
                '<tr><td colspan="5" style="text-align:center;color:red">Error: ' + e.message + '</td></tr>';
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
            return '<tr onclick="selectResult(\\'' + r.id + '\\', \\'' + r.type + '\\')" style="cursor:pointer">' +
                   '<td>' + r.type + '</td>' +
                   '<td>' + r.name + '</td>' +
                   '<td>' + r.details + '</td>' +
                   '<td>' + r.status + '</td>' +
                   '<td>' + r.date + '</td>' +
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
// MULTI-SELECT DIALOGS
// ============================================================================

/**
 * Opens the multi-select editor for the currently selected cell
 * Called from menu: Tools > Multi-Select > Open Editor
 * Validates the cell is in Member Directory and is a multi-select column
 */
function openCellMultiSelectEditor() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();

  // Validate we're in Member Directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Multi-Select Editor',
      'Please select a cell in the Member Directory sheet.\n\n' +
      'Multi-select columns include:\n' +
      '• Office Days\n' +
      '• Preferred Communication\n' +
      '• Best Time to Contact\n' +
      '• Committees\n' +
      '• Assigned Steward(s)',
      ui.ButtonSet.OK);
    return;
  }

  // Validate single cell selection
  if (!range || range.getNumRows() > 1 || range.getNumColumns() > 1) {
    ui.alert('Multi-Select Editor',
      'Please select a single cell in a multi-select column.',
      ui.ButtonSet.OK);
    return;
  }

  var row = range.getRow();
  var col = range.getColumn();

  // Skip header row
  if (row < 2) {
    ui.alert('Multi-Select Editor',
      'Please select a data cell (not the header row).',
      ui.ButtonSet.OK);
    return;
  }

  // Check if this is a multi-select column
  var config = getMultiSelectConfig(col);
  if (!config) {
    ui.alert('Multi-Select Editor',
      'This column does not support multi-select.\n\n' +
      'Multi-select columns include:\n' +
      '• Office Days\n' +
      '• Preferred Communication\n' +
      '• Best Time to Contact\n' +
      '• Committees\n' +
      '• Assigned Steward(s)',
      ui.ButtonSet.OK);
    return;
  }

  // Store target cell coordinates for the callback
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('multiSelectRow', row.toString());
  props.setProperty('multiSelectCol', col.toString());

  // Get current cell value to pre-select items
  var currentValue = range.getValue().toString();
  var currentSelections = currentValue ? currentValue.split(',').map(function(s) { return s.trim(); }) : [];

  // Load available options from Config sheet
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) {
    ui.alert('Error', 'Config sheet not found.', ui.ButtonSet.OK);
    return;
  }

  var options = getConfigValues(configSheet, config.configCol);
  if (options.length === 0) {
    ui.alert('Multi-Select Editor',
      'No options found in the Config sheet for ' + config.label + '.\n\n' +
      'Please add options to the Config sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  // Build items array with selection state
  var items = options.map(function(option) {
    return {
      id: option,
      label: option,
      selected: currentSelections.indexOf(option) !== -1
    };
  });

  // Show the dialog
  showMultiSelectDialog('Select ' + config.label, items, 'applyMultiSelectValue');
}

/**
 * Shows multi-select dialog for bulk operations
 * @param {string} title - Dialog title
 * @param {Array} items - Items to display for selection
 * @param {string} callback - Name of function to call with selections
 */
function showMultiSelectDialog(title, items, callback) {
  const html = HtmlService.createHtmlOutput(getMultiSelectHtml(items, callback))
    .setWidth(DIALOG_SIZES.MEDIUM.width)
    .setHeight(DIALOG_SIZES.MEDIUM.height);

  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/**
 * Generates HTML for multi-select dialog
 * @param {Array} items - Items with id and label properties
 * @param {string} callback - Server callback function name
 * @return {string} HTML content
 */
function getMultiSelectHtml(items, callback) {
  const itemsJson = JSON.stringify(items);

  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .multi-select-container {
          padding: 20px;
        }
        .select-controls {
          display: flex;
          gap: 10px;
          margin-bottom: 15px;
        }
        .items-list {
          max-height: 300px;
          overflow-y: auto;
          border: 1px solid ${UI_THEME.BORDER_COLOR};
          border-radius: 8px;
        }
        .item-row {
          display: flex;
          align-items: center;
          padding: 10px 15px;
          border-bottom: 1px solid ${UI_THEME.BORDER_COLOR};
        }
        .item-row:last-child {
          border-bottom: none;
        }
        .item-row:hover {
          background: #f8f9fa;
        }
        .item-checkbox {
          margin-right: 12px;
        }
        .item-label {
          flex: 1;
        }
        .action-bar {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-top: 20px;
        }
        .selection-count {
          color: ${UI_THEME.TEXT_SECONDARY};
        }
      </style>
    </head>
    <body>
      <div class="multi-select-container">
        <div class="select-controls">
          <button class="btn btn-secondary" onclick="selectAll()">Select All</button>
          <button class="btn btn-secondary" onclick="selectNone()">Select None</button>
          <input type="text" placeholder="Filter items..." class="filter-input"
                 style="flex:1; margin-left:10px;" oninput="filterItems(this.value)">
        </div>

        <div class="items-list" id="itemsList">
        </div>

        <div class="action-bar">
          <span class="selection-count" id="selectionCount">0 selected</span>
          <div>
            <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button class="btn btn-primary" onclick="submitSelection()">Apply</button>
          </div>
        </div>
      </div>

      <script>
        const items = ${itemsJson};
        const callbackFn = '${callback}';
        let filteredItems = [...items];

        function renderItems() {
          const container = document.getElementById('itemsList');
          container.innerHTML = filteredItems.map(function(item) {
            return '<div class="item-row">' +
                   '<input type="checkbox" class="item-checkbox" data-id="' + item.id + '"' +
                   (item.selected ? ' checked' : '') + ' onchange="updateCount()">' +
                   '<span class="item-label">' + item.label + '</span>' +
                   '</div>';
          }).join('');
          updateCount();
        }

        function selectAll() {
          document.querySelectorAll('.item-checkbox').forEach(cb => cb.checked = true);
          updateCount();
        }

        function selectNone() {
          document.querySelectorAll('.item-checkbox').forEach(cb => cb.checked = false);
          updateCount();
        }

        function filterItems(query) {
          const q = query.toLowerCase();
          filteredItems = items.filter(item => item.label.toLowerCase().includes(q));
          renderItems();
        }

        function updateCount() {
          const count = document.querySelectorAll('.item-checkbox:checked').length;
          document.getElementById('selectionCount').textContent = count + ' selected';
        }

        function submitSelection() {
          const selected = Array.from(document.querySelectorAll('.item-checkbox:checked'))
                               .map(cb => cb.dataset.id);
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) { alert('Error: ' + e.message); })
            [callbackFn](selected);
        }

        renderItems();
      </script>
    </body>
    </html>
  `;
}

// ============================================================================
// Note: showQuickActionsMenu() is defined in MobileQuickActions.gs which
// contains the comprehensive quick actions implementation.
// ============================================================================

// ============================================================================
// SIDEBAR
// ============================================================================

/**
 * Shows the main dashboard sidebar
 */
function showDashboardSidebar() {
  const html = HtmlService.createHtmlOutput(getDashboardSidebarHtml())
    .setTitle(COMMAND_CONFIG.SYSTEM_NAME);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Generates HTML for dashboard sidebar
 * @return {string} HTML content
 */
function getDashboardSidebarHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .sidebar {
          padding: 15px;
          font-family: 'Google Sans', Roboto, sans-serif;
        }
        .stats-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 10px;
          margin-bottom: 20px;
        }
        .stat-card {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 8px;
          text-align: center;
        }
        .stat-number {
          font-size: 28px;
          font-weight: bold;
          color: ${UI_THEME.PRIMARY_COLOR};
        }
        .stat-label {
          font-size: 12px;
          color: ${UI_THEME.TEXT_SECONDARY};
          margin-top: 4px;
        }
        .section-title {
          font-size: 14px;
          font-weight: 600;
          color: ${UI_THEME.TEXT_PRIMARY};
          margin: 20px 0 10px 0;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .deadline-list {
          max-height: 200px;
          overflow-y: auto;
        }
        .deadline-item {
          display: flex;
          justify-content: space-between;
          padding: 10px;
          border-left: 3px solid ${UI_THEME.WARNING_COLOR};
          background: #fffbeb;
          margin-bottom: 8px;
          border-radius: 0 4px 4px 0;
        }
        .deadline-item.urgent {
          border-left-color: ${UI_THEME.DANGER_COLOR};
          background: #fef2f2;
        }
        .quick-link {
          display: block;
          padding: 10px;
          color: ${UI_THEME.PRIMARY_COLOR};
          text-decoration: none;
          border-radius: 4px;
        }
        .quick-link:hover {
          background: #e8f0fe;
        }
        .refresh-btn {
          width: 100%;
          margin-top: 15px;
        }
      </style>
    </head>
    <body>
      <div class="sidebar">
        <div class="stats-grid" id="statsGrid">
          <div class="stat-card">
            <div class="stat-number" id="openCount">-</div>
            <div class="stat-label">Open Grievances</div>
          </div>
          <div class="stat-card">
            <div class="stat-number" id="pendingCount">-</div>
            <div class="stat-label">Pending Response</div>
          </div>
          <div class="stat-card">
            <div class="stat-number" id="memberCount">-</div>
            <div class="stat-label">Total Members</div>
          </div>
          <div class="stat-card">
            <div class="stat-number" id="resolvedCount">-</div>
            <div class="stat-label">Resolved (YTD)</div>
          </div>
        </div>

        <div class="section-title">Upcoming Deadlines</div>
        <div class="deadline-list" id="deadlineList">
          <div style="text-align:center; color:#666; padding:20px;">Loading...</div>
        </div>

        <div class="section-title">Quick Links</div>
        <a href="#" class="quick-link" onclick="goToSheet('${SHEET_NAMES.GRIEVANCE_TRACKER}')">
          📋 Grievance Tracker
        </a>
        <a href="#" class="quick-link" onclick="goToSheet('${SHEET_NAMES.MEMBER_DIRECTORY}')">
          👥 Member Directory
        </a>
        <a href="#" class="quick-link" onclick="goToSheet('${SHEET_NAMES.REPORTS}')">
          📊 Reports
        </a>
        <a href="#" class="quick-link" onclick="showSearch()">
          🔍 Search Dashboard
        </a>

        <button class="btn btn-secondary refresh-btn" onclick="refreshData()">
          🔄 Refresh Data
        </button>
      </div>

      <script>
        function loadStats() {
          google.script.run
            .withSuccessHandler(function(stats) {
              document.getElementById('openCount').textContent = stats.open;
              document.getElementById('pendingCount').textContent = stats.pending;
              document.getElementById('memberCount').textContent = stats.members;
              document.getElementById('resolvedCount').textContent = stats.resolved;
            })
            .getDashboardStats();
        }

        function loadDeadlines() {
          google.script.run
            .withSuccessHandler(function(deadlines) {
              const container = document.getElementById('deadlineList');
              if (deadlines.length === 0) {
                container.innerHTML = '<div style="text-align:center;color:#666;padding:20px;">No upcoming deadlines</div>';
                return;
              }
              container.innerHTML = deadlines.map(function(d) {
                const urgentClass = d.daysLeft <= 3 ? 'urgent' : '';
                return '<div class="deadline-item ' + urgentClass + '">' +
                       '<div>' + d.grievanceId + '<br><small>' + d.step + '</small></div>' +
                       '<div style="text-align:right">' + d.daysLeft + ' days<br><small>' + d.date + '</small></div>' +
                       '</div>';
              }).join('');
            })
            .getUpcomingDeadlines(7);
        }

        function goToSheet(sheetName) {
          google.script.run.navigateToSheet(sheetName);
        }

        function showSearch() {
          google.script.run.showDesktopSearch();
        }

        function refreshData() {
          loadStats();
          loadDeadlines();
        }

        // Initial load
        loadStats();
        loadDeadlines();
      </script>
    </body>
    </html>
  `;
}

// ============================================================================
// COMMON UI UTILITIES
// ============================================================================

/**
 * Gets common CSS styles used across all dialogs
 * @return {string} CSS styles
 */
function getCommonStyles() {
  return `
    <style>
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      body {
        font-family: 'Google Sans', Roboto, Arial, sans-serif;
        font-size: 14px;
        color: ${UI_THEME.TEXT_PRIMARY};
        line-height: 1.5;
      }
      .btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 10px 20px;
        border: none;
        border-radius: 6px;
        font-size: 14px;
        font-weight: 500;
        cursor: pointer;
        transition: background 0.2s, transform 0.1s;
      }
      .btn:hover {
        transform: translateY(-1px);
      }
      .btn:active {
        transform: translateY(0);
      }
      .btn-primary {
        background: ${UI_THEME.PRIMARY_COLOR};
        color: white;
      }
      .btn-primary:hover {
        background: #1557b0;
      }
      .btn-secondary {
        background: #f1f3f4;
        color: ${UI_THEME.TEXT_PRIMARY};
      }
      .btn-secondary:hover {
        background: #e8eaed;
      }
      .btn-danger {
        background: ${UI_THEME.DANGER_COLOR};
        color: white;
      }
      .btn-danger:hover {
        background: #c5221f;
      }
      .btn-success {
        background: ${UI_THEME.SECONDARY_COLOR};
        color: white;
      }
      .form-group {
        margin-bottom: 15px;
      }
      .form-label {
        display: block;
        font-weight: 500;
        margin-bottom: 5px;
        color: ${UI_THEME.TEXT_PRIMARY};
      }
      .form-input {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        transition: border-color 0.2s;
      }
      .form-input:focus {
        outline: none;
        border-color: ${UI_THEME.PRIMARY_COLOR};
      }
      .form-select {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        background: white;
      }
      .form-textarea {
        width: 100%;
        padding: 10px 12px;
        border: 1px solid ${UI_THEME.BORDER_COLOR};
        border-radius: 6px;
        font-size: 14px;
        min-height: 100px;
        resize: vertical;
      }
      .alert {
        padding: 12px 16px;
        border-radius: 6px;
        margin-bottom: 15px;
      }
      .alert-info {
        background: #e8f0fe;
        color: #1967d2;
      }
      .alert-warning {
        background: #fef7e0;
        color: #ea8600;
      }
      .alert-error {
        background: #fce8e6;
        color: #c5221f;
      }
      .alert-success {
        background: #e6f4ea;
        color: #137333;
      }
    </style>
  `;
}

/**
 * Shows a toast notification
 * @param {string} message - Message to display
 * @param {string} title - Optional title
 */
function showToast(message, title) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Dashboard', 5);
}

/**
 * Shows a confirmation dialog
 * @param {string} message - Confirmation message
 * @param {string} title - Dialog title
 * @return {boolean} True if user confirmed
 */
function showConfirmation(message, title) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(title || 'Confirm', message, ui.ButtonSet.YES_NO);
  return response === ui.Button.YES;
}

/**
 * Shows an alert dialog
 * @param {string} message - Alert message
 * @param {string} title - Dialog title
 */
function showAlert(message, title) {
  SpreadsheetApp.getUi().alert(title || 'Alert', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// navigateToSheet() - REMOVED DUPLICATE - see line 565 for main definition

/**
 * Navigates to a specific record (row) in the appropriate sheet
 * @param {string} id - Record ID
 * @param {string} type - Record type ('member' or 'grievance')
 */
function navigateToRecord(id, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet, idColumn;

  if (type === 'member') {
    sheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
    idColumn = MEMBER_COLUMNS.ID + 1;
  } else if (type === 'grievance') {
    sheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
    idColumn = GRIEVANCE_COLUMNS.GRIEVANCE_ID + 1;
  }

  if (!sheet) return;

  ss.setActiveSheet(sheet);

  // Find the row with matching ID
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColumn - 1] === id) {
      sheet.setActiveRange(sheet.getRange(i + 1, 1));
      break;
    }
  }
}
/**
 * ============================================================================
 * COMFORT VIEW ACCESSIBILITY & THEMING
 * ============================================================================
 * Features for neurodivergent users + theme customization
 */

// Comfort View Configuration
var COMFORT_VIEW_CONFIG = {
  FOCUS_MODE_COLORS: { background: '#f5f5f5', header: '#4a4a4a', accent: '#6b9bd1' },
  HIGH_CONTRAST: { background: '#ffffff', header: '#000000', accent: '#0066cc' },
  PASTEL: { background: '#fef9e7', header: '#85929e', accent: '#7fb3d5' }
};

// Theme Configuration
var THEME_CONFIG = {
  THEMES: {
    LIGHT: { name: 'Light', icon: '☀️', background: '#ffffff', headerBackground: '#1a73e8', headerText: '#ffffff', evenRow: '#f8f9fa', oddRow: '#ffffff', text: '#202124', accent: '#1a73e8' },
    DARK: { name: 'Dark', icon: '🌙', background: '#202124', headerBackground: '#35363a', headerText: '#e8eaed', evenRow: '#292a2d', oddRow: '#202124', text: '#e8eaed', accent: '#8ab4f8' },
    PURPLE: { name: '509 Purple', icon: '💜', background: '#ffffff', headerBackground: '#5B4B9E', headerText: '#ffffff', evenRow: '#E8E3F3', oddRow: '#ffffff', text: '#1F2937', accent: '#6B5CA5' },
    GREEN: { name: 'Union Green', icon: '💚', background: '#ffffff', headerBackground: '#059669', headerText: '#ffffff', evenRow: '#D1FAE5', oddRow: '#ffffff', text: '#1F2937', accent: '#10B981' }
  },
  DEFAULT_THEME: 'LIGHT'
};

// ==================== COMFORT VIEW SETTINGS ====================

function getADHDSettings() {
  var props = PropertiesService.getUserProperties();
  var settingsJSON = props.getProperty('adhdSettings');
  if (settingsJSON) return JSON.parse(settingsJSON);
  return { theme: 'default', fontSize: 10, zebraStripes: false, gridlines: true, reducedMotion: false, breakInterval: 0 };
}

function saveADHDSettings(settings) {
  var props = PropertiesService.getUserProperties();
  var current = getADHDSettings();
  var newSettings = Object.assign({}, current, settings);
  props.setProperty('adhdSettings', JSON.stringify(newSettings));
  applyADHDSettings(newSettings);
}

function applyADHDSettings(settings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (settings.fontSize) sheet.getDataRange().setFontSize(parseInt(settings.fontSize));
    if (settings.zebraStripes) applyZebraStripes(sheet);
    if (settings.gridlines !== undefined) sheet.setHiddenGridlines(!settings.gridlines);
  });
}

function resetADHDSettings() {
  PropertiesService.getUserProperties().deleteProperty('adhdSettings');
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Comfort View settings reset', 'Settings', 3);
}

// ==================== VISUAL HELPERS ====================

function hideAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name !== SHEETS.CONFIG && name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) {
      sheet.setHiddenGridlines(true);
    }
  });
  SpreadsheetApp.getUi().alert('✅ Gridlines hidden on dashboards!');
}

function showAllGridlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) { sheet.showGridlines(); });
  SpreadsheetApp.getUi().alert('✅ Gridlines shown on all sheets.');
}

function toggleGridlinesADHD() {
  var settings = getADHDSettings();
  settings.gridlines = !settings.gridlines;
  saveADHDSettings(settings);
}

function applyZebraStripes(sheet) {
  var sheetName = sheet.getName();

  // For data sheets (Member Directory, Grievance Log), apply to ALL rows in the sheet
  // Uses getMaxRows() to cover all existing and future rows
  var isDataSheet = (sheetName === SHEETS.MEMBER_DIR || sheetName === SHEETS.GRIEVANCE_LOG);
  var totalRows = isDataSheet ? sheet.getMaxRows() : sheet.getLastRow();

  if (totalRows < 2) return;

  var lastCol = sheet.getLastColumn() || 26; // Default to 26 columns if sheet is empty
  var range = sheet.getRange(2, 1, totalRows - 1, lastCol);

  // Remove existing bandings first
  var bandings = range.getBandings();
  if (bandings.length > 0) bandings[0].remove();

  // Apply fresh banding
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function removeZebraStripes(sheet) {
  sheet.getBandings().forEach(function(b) { b.remove(); });
}

function toggleZebraStripes() {
  var settings = getADHDSettings();
  settings.zebraStripes = !settings.zebraStripes;
  saveADHDSettings(settings);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    if (settings.zebraStripes) applyZebraStripes(sheet);
    else removeZebraStripes(sheet);
  });
  ss.toast(settings.zebraStripes ? '✅ Zebra stripes enabled' : '🔕 Zebra stripes disabled', 'Visual', 3);
}

function toggleReducedMotion() {
  var settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
}

// ==================== FOCUS MODE ====================

/**
 * Activate Focus Mode - distraction-free view for focused work
 *
 * WHAT IT DOES:
 * - Hides all sheets except the one you're currently viewing
 * - Removes gridlines to reduce visual clutter
 * - Creates a clean, focused work environment
 *
 * HOW TO EXIT:
 * - Use menu: Comfort View > Exit Focus Mode
 * - Or run: deactivateFocusMode()
 *
 * BEST FOR:
 * - Deep work on a single task
 * - Reducing cognitive load
 * - Preventing tab-switching distractions
 */
function activateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var active = ss.getActiveSheet();

  // Count visible sheets to warn user
  var visibleSheets = ss.getSheets().filter(function(s) { return !s.isSheetHidden(); });

  if (visibleSheets.length === 1) {
    ui.alert('🎯 Focus Mode',
      'Focus mode is already active (only one sheet visible).\n\n' +
      'To exit focus mode, use:\n' +
      'Comfort View Menu > Exit Focus Mode',
      ui.ButtonSet.OK);
    return;
  }

  // Hide all sheets except active
  ss.getSheets().forEach(function(sheet) {
    if (sheet.getName() !== active.getName()) sheet.hideSheet();
  });
  active.setHiddenGridlines(true);

  ui.alert('🎯 Focus Mode Activated',
    'You are now in Focus Mode on: "' + active.getName() + '"\n\n' +
    'WHAT THIS DOES:\n' +
    '• Hides all other sheets to reduce distractions\n' +
    '• Removes gridlines for a cleaner view\n\n' +
    'TO EXIT:\n' +
    '• Use menu: ♿ Comfort View > Exit Focus Mode\n' +
    '• Or run: deactivateFocusMode()\n\n' +
    '💡 Tip: Focus mode helps with deep work and reduces cognitive load.',
    ui.ButtonSet.OK);
}

function deactivateFocusMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    if (sheet.isSheetHidden()) sheet.showSheet();
  });
  var settings = getADHDSettings();
  ss.getActiveSheet().setHiddenGridlines(!settings.gridlines);
  ss.toast('✅ Focus mode deactivated', 'Focus Mode', 3);
}

// ==================== QUICK CAPTURE NOTEPAD ====================

/**
 * Quick Capture Notepad - Fast note-taking without losing focus
 * Notes are stored per-user in Script Properties
 */

/**
 * Gets the current user's quick capture notes
 * @returns {string} The saved notes or empty string
 */
function getQuickCaptureNotes() {
  var userProps = PropertiesService.getUserProperties();
  return userProps.getProperty('quickCaptureNotes') || '';
}

/**
 * Saves quick capture notes for the current user
 * @param {string} notes - The notes to save
 * @returns {Object} Result object with success status
 */
function saveQuickCaptureNotes(notes) {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty('quickCaptureNotes', notes || '');
    userProps.setProperty('quickCaptureLastSaved', new Date().toISOString());
    return { success: true, message: 'Notes saved' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Clears the quick capture notes
 * @returns {Object} Result object with success status
 */
function clearQuickCaptureNotes() {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty('quickCaptureNotes');
    userProps.deleteProperty('quickCaptureLastSaved');
    return { success: true, message: 'Notes cleared' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Gets metadata about the quick capture notes
 * @returns {Object} Object with lastSaved timestamp and character count
 */
function getQuickCaptureMetadata() {
  var userProps = PropertiesService.getUserProperties();
  var notes = userProps.getProperty('quickCaptureNotes') || '';
  var lastSaved = userProps.getProperty('quickCaptureLastSaved') || null;
  return {
    charCount: notes.length,
    wordCount: notes.trim() ? notes.trim().split(/\s+/).length : 0,
    lastSaved: lastSaved
  };
}

/**
 * Shows the Quick Capture Notepad dialog
 * A fast notepad for capturing thoughts without losing focus
 */
function showQuickCaptureNotepad() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", Roboto, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; padding: 16px; color: #F8FAFC; }' +
    '.header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }' +
    'h3 { font-size: 18px; display: flex; align-items: center; gap: 8px; }' +
    '.meta { font-size: 12px; color: #64748B; }' +
    'textarea { width: 100%; height: 280px; padding: 12px; border: 2px solid #334155; border-radius: 8px; background: #1E293B; color: #F8FAFC; font-size: 14px; font-family: inherit; resize: none; outline: none; }' +
    'textarea:focus { border-color: #7C3AED; }' +
    'textarea::placeholder { color: #64748B; }' +
    '.btn-row { display: flex; gap: 8px; margin-top: 12px; }' +
    'button { flex: 1; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s; }' +
    '.btn-primary { background: linear-gradient(135deg, #7C3AED, #5B21B6); color: white; }' +
    '.btn-primary:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(124,58,237,0.3); }' +
    '.btn-secondary { background: #334155; color: #F8FAFC; }' +
    '.btn-danger { background: #DC2626; color: white; }' +
    '.status { margin-top: 8px; font-size: 12px; color: #10B981; text-align: center; opacity: 0; transition: opacity 0.3s; }' +
    '.status.show { opacity: 1; }' +
    '</style>' +
    '<div class="header">' +
    '  <h3>📝 Quick Capture</h3>' +
    '  <span class="meta" id="meta"></span>' +
    '</div>' +
    '<textarea id="notes" placeholder="Capture your thoughts quickly...\\n\\nUse this notepad to jot down ideas, reminders, or notes without losing focus on your current task.\\n\\nYour notes are auto-saved when you click Save."></textarea>' +
    '<div class="btn-row">' +
    '  <button class="btn-primary" onclick="saveNotes()">💾 Save</button>' +
    '  <button class="btn-secondary" onclick="copyNotes()">📋 Copy</button>' +
    '  <button class="btn-danger" onclick="clearNotes()">🗑️ Clear</button>' +
    '  <button class="btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '</div>' +
    '<div class="status" id="status"></div>' +
    '<script>' +
    'var notesEl = document.getElementById("notes");' +
    'var statusEl = document.getElementById("status");' +
    'var metaEl = document.getElementById("meta");' +
    '' +
    'google.script.run.withSuccessHandler(function(notes) {' +
    '  notesEl.value = notes || "";' +
    '  updateMeta();' +
    '}).getQuickCaptureNotes();' +
    '' +
    'notesEl.addEventListener("input", updateMeta);' +
    '' +
    'function updateMeta() {' +
    '  var text = notesEl.value;' +
    '  var words = text.trim() ? text.trim().split(/\\s+/).length : 0;' +
    '  metaEl.textContent = words + " words, " + text.length + " chars";' +
    '}' +
    '' +
    'function showStatus(msg, isError) {' +
    '  statusEl.textContent = msg;' +
    '  statusEl.style.color = isError ? "#EF4444" : "#10B981";' +
    '  statusEl.classList.add("show");' +
    '  setTimeout(function() { statusEl.classList.remove("show"); }, 2000);' +
    '}' +
    '' +
    'function saveNotes() {' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    showStatus(result.success ? "✅ Notes saved!" : "❌ " + result.message, !result.success);' +
    '  }).saveQuickCaptureNotes(notesEl.value);' +
    '}' +
    '' +
    'function copyNotes() {' +
    '  navigator.clipboard.writeText(notesEl.value).then(function() {' +
    '    showStatus("📋 Copied to clipboard!");' +
    '  }).catch(function() {' +
    '    showStatus("❌ Failed to copy", true);' +
    '  });' +
    '}' +
    '' +
    'function clearNotes() {' +
    '  if (confirm("Clear all notes? This cannot be undone.")) {' +
    '    notesEl.value = "";' +
    '    google.script.run.withSuccessHandler(function(result) {' +
    '      showStatus(result.success ? "🗑️ Notes cleared" : "❌ " + result.message, !result.success);' +
    '      updateMeta();' +
    '    }).clearQuickCaptureNotes();' +
    '  }' +
    '}' +
    '' +
    'notesEl.addEventListener("keydown", function(e) {' +
    '  if ((e.ctrlKey || e.metaKey) && e.key === "s") {' +
    '    e.preventDefault();' +
    '    saveNotes();' +
    '  }' +
    '});' +
    '</script>'
  )
  .setWidth(500)
  .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '📝 Quick Capture Notepad');
}

/**
 * Shows import dialog for bulk member import from CSV
 * Provides paste area for CSV data with preview and validation
 */
function showImportDialog() {
  var html = HtmlService.createHtmlOutput(getImportDialogHtml_())
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Import Members from CSV');
}

/**
 * Generates HTML for import dialog
 * @private
 */
function getImportDialogHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Roboto", Arial, sans-serif; padding: 20px; background: #f5f5f5; }' +
    '.container { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }' +
    'h2 { color: #1a73e8; margin-bottom: 15px; font-size: 18px; }' +
    '.section { margin-bottom: 20px; }' +
    '.section-title { font-weight: bold; color: #333; margin-bottom: 8px; font-size: 14px; }' +
    'textarea { width: 100%; height: 200px; border: 2px solid #e0e0e0; border-radius: 6px; padding: 12px; font-family: monospace; font-size: 12px; resize: vertical; }' +
    'textarea:focus { border-color: #1a73e8; outline: none; }' +
    '.format-hint { background: #e8f0fe; padding: 12px; border-radius: 6px; font-size: 12px; color: #1967d2; margin-bottom: 15px; }' +
    '.format-hint code { background: #fff; padding: 2px 6px; border-radius: 3px; }' +
    '.preview { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 12px; max-height: 150px; overflow-y: auto; font-size: 12px; display: none; }' +
    '.preview-row { padding: 4px 0; border-bottom: 1px solid #eee; }' +
    '.preview-row:last-child { border-bottom: none; }' +
    '.btn-row { display: flex; gap: 10px; margin-top: 15px; }' +
    'button { padding: 10px 20px; border-radius: 6px; font-size: 14px; cursor: pointer; border: none; font-weight: 500; }' +
    '.btn-primary { background: #1a73e8; color: white; }' +
    '.btn-primary:hover { background: #1557b0; }' +
    '.btn-secondary { background: #f1f3f4; color: #5f6368; }' +
    '.btn-secondary:hover { background: #e8eaed; }' +
    '.status { padding: 10px; border-radius: 6px; margin-top: 10px; display: none; }' +
    '.status.success { background: #e6f4ea; color: #137333; }' +
    '.status.error { background: #fce8e6; color: #c5221f; }' +
    '.count { font-weight: bold; color: #1a73e8; }' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>📥 Import Members from CSV</h2>' +
    '<div class="format-hint">' +
    '<strong>Required columns:</strong> First Name, Last Name<br>' +
    '<strong>Optional:</strong> Email, Phone, Job Title, Unit, Work Location, Supervisor, Is Steward, Dues Paying<br>' +
    '<strong>Format:</strong> Paste CSV with headers in first row. Use comma separator.' +
    '</div>' +
    '<div class="section">' +
    '<div class="section-title">Paste CSV Data:</div>' +
    '<textarea id="csvData" placeholder="First Name,Last Name,Email,Phone,Job Title,Unit&#10;John,Doe,john@example.com,555-1234,Clerk,Admin&#10;Jane,Smith,jane@example.com,555-5678,Manager,Operations"></textarea>' +
    '</div>' +
    '<div class="section">' +
    '<div class="section-title">Preview (<span id="rowCount" class="count">0</span> rows):</div>' +
    '<div id="preview" class="preview"></div>' +
    '</div>' +
    '<div id="status" class="status"></div>' +
    '<div class="btn-row">' +
    '<button class="btn-secondary" onclick="previewData()">👁️ Preview</button>' +
    '<button class="btn-primary" onclick="importData()">📥 Import</button>' +
    '<button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '</div></div>' +
    '<script>' +
    'function previewData() {' +
    '  var csv = document.getElementById("csvData").value.trim();' +
    '  if (!csv) { showStatus("Please paste CSV data first", "error"); return; }' +
    '  var rows = parseCSV(csv);' +
    '  if (rows.length < 2) { showStatus("Need at least a header row and one data row", "error"); return; }' +
    '  document.getElementById("rowCount").textContent = rows.length - 1;' +
    '  var previewHtml = "<div class=\\"preview-row\\"><strong>" + rows[0].join(" | ") + "</strong></div>";' +
    '  for (var i = 1; i < Math.min(rows.length, 6); i++) {' +
    '    previewHtml += "<div class=\\"preview-row\\">" + rows[i].join(" | ") + "</div>";' +
    '  }' +
    '  if (rows.length > 6) previewHtml += "<div class=\\"preview-row\\">... and " + (rows.length - 6) + " more rows</div>";' +
    '  document.getElementById("preview").innerHTML = previewHtml;' +
    '  document.getElementById("preview").style.display = "block";' +
    '  showStatus("Preview ready. Click Import to add " + (rows.length - 1) + " members.", "success");' +
    '}' +
    'function parseCSV(csv) {' +
    '  var lines = csv.split(/\\r?\\n/);' +
    '  return lines.filter(function(line) { return line.trim(); }).map(function(line) {' +
    '    var result = []; var cell = ""; var inQuotes = false;' +
    '    for (var i = 0; i < line.length; i++) {' +
    '      var c = line[i];' +
    '      if (c === "\\"") { inQuotes = !inQuotes; }' +
    '      else if (c === "," && !inQuotes) { result.push(cell.trim()); cell = ""; }' +
    '      else { cell += c; }' +
    '    }' +
    '    result.push(cell.trim());' +
    '    return result;' +
    '  });' +
    '}' +
    'function importData() {' +
    '  var csv = document.getElementById("csvData").value.trim();' +
    '  if (!csv) { showStatus("Please paste CSV data first", "error"); return; }' +
    '  showStatus("Importing...", "success");' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    if (result.success) {' +
    '      showStatus("✅ Successfully imported " + result.count + " members!", "success");' +
    '      setTimeout(function() { google.script.host.close(); }, 2000);' +
    '    } else {' +
    '      showStatus("❌ " + result.error, "error");' +
    '    }' +
    '  }).withFailureHandler(function(err) {' +
    '    showStatus("❌ Error: " + err.message, "error");' +
    '  }).processMemberImport(csv);' +
    '}' +
    'function showStatus(msg, type) {' +
    '  var el = document.getElementById("status");' +
    '  el.textContent = msg;' +
    '  el.className = "status " + type;' +
    '  el.style.display = "block";' +
    '}' +
    '</script></body></html>';
}

/**
 * Processes member import from CSV data
 * @param {string} csvData - Raw CSV data
 * @returns {Object} Result with success status and count/error
 */
function processMemberImport(csvData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      return { success: false, error: 'Member Directory sheet not found' };
    }

    // Parse CSV
    var lines = csvData.split(/\r?\n/).filter(function(l) { return l.trim(); });
    if (lines.length < 2) {
      return { success: false, error: 'Need at least header and one data row' };
    }

    // Parse header and map columns
    var headers = parseCSVLine_(lines[0]);
    var columnMap = mapImportColumns_(headers);

    if (!columnMap.firstName || !columnMap.lastName) {
      return { success: false, error: 'Required columns missing: First Name, Last Name' };
    }

    // Get existing data to find last row
    var lastRow = sheet.getLastRow();

    // Process data rows
    var importedCount = 0;
    for (var i = 1; i < lines.length; i++) {
      var values = parseCSVLine_(lines[i]);
      if (!values || values.length === 0) continue;

      // Build row data
      var rowData = [];
      rowData[MEMBER_COLS.FIRST_NAME - 1] = values[columnMap.firstName] || '';
      rowData[MEMBER_COLS.LAST_NAME - 1] = values[columnMap.lastName] || '';
      rowData[MEMBER_COLS.EMAIL - 1] = columnMap.email !== undefined ? values[columnMap.email] || '' : '';
      rowData[MEMBER_COLS.PHONE - 1] = columnMap.phone !== undefined ? values[columnMap.phone] || '' : '';
      rowData[MEMBER_COLS.JOB_TITLE - 1] = columnMap.jobTitle !== undefined ? values[columnMap.jobTitle] || '' : '';
      rowData[MEMBER_COLS.UNIT - 1] = columnMap.unit !== undefined ? values[columnMap.unit] || '' : '';
      rowData[MEMBER_COLS.WORK_LOCATION - 1] = columnMap.workLocation !== undefined ? values[columnMap.workLocation] || '' : '';
      rowData[MEMBER_COLS.SUPERVISOR - 1] = columnMap.supervisor !== undefined ? values[columnMap.supervisor] || '' : '';
      rowData[MEMBER_COLS.IS_STEWARD - 1] = columnMap.isSteward !== undefined ? (values[columnMap.isSteward] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'No';
      rowData[MEMBER_COLS.DUES_PAYING - 1] = columnMap.duesPaying !== undefined ? (values[columnMap.duesPaying] || '').toLowerCase() === 'yes' ? 'Yes' : 'No' : 'Yes';

      // Fill empty cells
      while (rowData.length < sheet.getLastColumn()) {
        rowData.push('');
      }

      // Append row
      sheet.appendRow(rowData);
      importedCount++;
    }

    // Generate Member IDs for imported rows
    if (typeof generateMissingMemberIDs === 'function') {
      generateMissingMemberIDs();
    }

    return { success: true, count: importedCount };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Parses a single CSV line handling quoted values
 * @private
 */
function parseCSVLine_(line) {
  var result = [];
  var cell = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if (c === ',' && !inQuotes) {
      result.push(cell.trim());
      cell = '';
    } else {
      cell += c;
    }
  }
  result.push(cell.trim());
  return result;
}

/**
 * Maps CSV headers to member column indices
 * @private
 */
function mapImportColumns_(headers) {
  var map = {};
  var headerLower = headers.map(function(h) { return (h || '').toLowerCase().replace(/[^a-z]/g, ''); });

  for (var i = 0; i < headerLower.length; i++) {
    var h = headerLower[i];
    if (h === 'firstname' || h === 'first') map.firstName = i;
    else if (h === 'lastname' || h === 'last') map.lastName = i;
    else if (h === 'email' || h === 'emailaddress') map.email = i;
    else if (h === 'phone' || h === 'phonenumber') map.phone = i;
    else if (h === 'jobtitle' || h === 'title' || h === 'position') map.jobTitle = i;
    else if (h === 'unit' || h === 'department' || h === 'dept') map.unit = i;
    else if (h === 'worklocation' || h === 'location' || h === 'worksite') map.workLocation = i;
    else if (h === 'supervisor' || h === 'manager') map.supervisor = i;
    else if (h === 'issteward' || h === 'steward') map.isSteward = i;
    else if (h === 'duespaying' || h === 'dues') map.duesPaying = i;
  }

  return map;
}

/**
 * Shows export dialog for member directory export
 * Creates downloadable CSV file
 */
function showExportDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory sheet not found.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No member data to export.');
    return;
  }

  // Get data
  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

  // Convert to CSV
  var csv = data.map(function(row) {
    return row.map(function(cell) {
      var str = String(cell === null || cell === undefined ? '' : cell);
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',');
  }).join('\n');

  // Create temporary file in Drive
  var fileName = 'MemberDirectory_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.csv';
  var blob = Utilities.newBlob(csv, 'text/csv', fileName);
  var file = DriveApp.createFile(blob);

  // Make file accessible
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Show download link
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }' +
    '.success { color: #137333; font-size: 48px; margin-bottom: 10px; }' +
    'h2 { color: #333; margin-bottom: 20px; }' +
    '.info { background: #e8f0fe; padding: 15px; border-radius: 8px; margin: 20px 0; }' +
    '.btn { display: inline-block; background: #1a73e8; color: white; padding: 12px 24px; ' +
    'text-decoration: none; border-radius: 6px; font-weight: bold; margin: 10px; }' +
    '.btn:hover { background: #1557b0; }' +
    '.btn-secondary { background: #5f6368; }' +
    '.note { color: #666; font-size: 12px; margin-top: 20px; }' +
    '</style>' +
    '<div class="success">✅</div>' +
    '<h2>Export Ready!</h2>' +
    '<div class="info">' +
    '<strong>' + (lastRow - 1) + ' members</strong> exported to CSV<br>' +
    'File: ' + fileName +
    '</div>' +
    '<a href="' + file.getDownloadUrl() + '" target="_blank" class="btn">📥 Download CSV</a>' +
    '<a href="' + file.getUrl() + '" target="_blank" class="btn btn-secondary">📂 Open in Drive</a>' +
    '<p class="note">File will be available in your Google Drive.<br>Link expires when you close this dialog.</p>' +
    '<script>setTimeout(function() { google.script.host.setHeight(350); }, 100);</script>'
  ).setWidth(450).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, '📤 Export Complete');
}

function setBreakReminders(minutes) {
  var settings = getADHDSettings();
  settings.breakInterval = minutes;
  saveADHDSettings(settings);
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'showBreakReminder') ScriptApp.deleteTrigger(t);
  });
  if (minutes > 0) {
    ScriptApp.newTrigger('showBreakReminder').timeBased().everyMinutes(minutes).create();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Break reminders: every ' + minutes + ' min', 'Comfort View', 3);
  }
}

function showBreakReminder() {
  SpreadsheetApp.getActiveSpreadsheet().toast('💆 Time for a break! Stretch and rest your eyes.', 'Break Reminder', 10);
}

// ==================== THEME MANAGEMENT ====================

function getCurrentTheme() {
  return PropertiesService.getUserProperties().getProperty('currentTheme') || THEME_CONFIG.DEFAULT_THEME;
}

function applyTheme(themeKey, scope) {
  scope = scope || 'all';
  var theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) throw new Error('Invalid theme');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = scope === 'all' ? ss.getSheets() : [ss.getActiveSheet()];
  sheets.forEach(function(sheet) { applyThemeToSheet(sheet, theme); });
  PropertiesService.getUserProperties().setProperty('currentTheme', themeKey);
  ss.toast(theme.icon + ' ' + theme.name + ' theme applied!', 'Theme', 3);
}

function applyThemeToSheet(sheet, theme) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(theme.headerBackground).setFontColor(theme.headerText).setFontWeight('bold');
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    sheet.getBandings().forEach(function(b) { b.remove(); });
    var banding = dataRange.applyRowBanding();
    banding.setFirstRowColor(theme.oddRow).setSecondRowColor(theme.evenRow);
    dataRange.setFontColor(theme.text);
  }
  sheet.setTabColor(theme.accent);
}

function previewTheme(themeKey) {
  var theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) throw new Error('Invalid theme');
  applyThemeToSheet(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), theme);
  SpreadsheetApp.getActiveSpreadsheet().toast('👁️ Previewing ' + theme.name, 'Preview', 5);
}

/**
 * Resets to default theme using theme system
 * NOTE: Renamed to avoid duplicate. Use resetToDefaultTheme() for hard reset to defaults.
 * This version uses the theme system; resetToDefaultTheme() clears all styling.
 */
function resetToDefaultThemeViaSystem_() {
  applyTheme(THEME_CONFIG.DEFAULT_THEME, 'all');
}

function quickToggleDarkMode() {
  var current = getCurrentTheme();
  applyTheme(current === 'LIGHT' ? 'DARK' : 'LIGHT', 'all');
}

// ==================== COMFORT VIEW CONTROL PANEL ====================

function showADHDControlPanel() {
  var settings = getADHDSettings();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8;border-bottom:3px solid #1a73e8;padding-bottom:10px}.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px;border-left:4px solid #1a73e8}.row{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid #e0e0e0}button{background:#1a73e8;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;margin:5px}button:hover{background:#1557b0}button.sec{background:#6c757d}</style></head><body><div class="container"><h2>♿ Comfort View Panel</h2><div class="section"><div class="row"><span>Zebra Stripes</span><button onclick="google.script.run.toggleZebraStripes();setTimeout(function(){location.reload()},1000)">' + (settings.zebraStripes ? '✅ On' : 'Off') + '</button></div><div class="row"><span>Gridlines</span><button onclick="google.script.run.toggleGridlinesADHD();setTimeout(function(){location.reload()},1000)">' + (settings.gridlines ? '✅ Visible' : 'Hidden') + '</button></div><div class="row"><span>Focus Mode</span><button onclick="google.script.run.activateFocusMode();google.script.host.close()">🎯 Activate</button></div></div><div class="section"><div class="row"><span>Quick Capture</span><button onclick="google.script.run.showQuickCaptureNotepad()">📝 Open</button></div><div class="row"><span>Pomodoro Timer</span><button onclick="google.script.run.startPomodoroTimer();google.script.host.close()">⏱️ Start</button></div></div><button class="sec" onclick="google.script.run.resetADHDSettings();google.script.host.close()">🔄 Reset</button><button class="sec" onclick="google.script.host.close()">Close</button></div></body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '♿ Comfort View Panel');
}

function showThemeManager() {
  var current = getCurrentTheme();
  var themeCards = Object.keys(THEME_CONFIG.THEMES).map(function(key) {
    var t = THEME_CONFIG.THEMES[key];
    return '<div style="background:#f8f9fa;padding:15px;border-radius:8px;cursor:pointer;border:3px solid ' + (current === key ? '#1a73e8' : 'transparent') + '" onclick="select(\'' + key + '\')">' +
      '<div style="font-size:32px;text-align:center">' + t.icon + '</div>' +
      '<div style="text-align:center;font-weight:bold">' + t.name + '</div>' +
      '<div style="height:20px;background:' + t.headerBackground + ';border-radius:4px;margin-top:10px"></div></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:5px}button.sec{background:#6c757d}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:15px;margin:20px 0}</style></head><body><div class="container"><h2>🎨 Theme Manager</h2><div class="grid">' + themeCards + '</div><button onclick="apply()">✅ Apply Theme</button><button class="sec" onclick="google.script.host.close()">Close</button></div><script>var sel="' + current + '";function select(k){sel=k;document.querySelectorAll(".grid>div").forEach(function(d){d.style.border="3px solid transparent"});event.currentTarget.style.border="3px solid #1a73e8"}function apply(){google.script.run.withSuccessHandler(function(){alert("Theme applied!");google.script.host.close()}).applyTheme(sel,"all")}</script></body></html>'
  ).setWidth(450).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🎨 Theme Manager');
}

// ==================== SETUP DEFAULTS ====================

/**
 * Setup Comfort View defaults with options dialog
 * User can choose which settings to apply and settings can be undone
 */
function setupADHDDefaults() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px;max-width:500px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.option{display:flex;align-items:center;padding:12px;margin:8px 0;background:#f8f9fa;border-radius:8px;cursor:pointer}' +
    '.option:hover{background:#e8f0fe}' +
    '.option input{margin-right:12px;width:18px;height:18px}' +
    '.option-text{flex:1}' +
    '.option-label{font-weight:bold;font-size:14px}' +
    '.option-desc{font-size:12px;color:#666;margin-top:2px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 24px;border:none;border-radius:4px;cursor:pointer;font-size:14px;flex:1}' +
    '.primary{background:#1a73e8;color:white}' +
    '.secondary{background:#e0e0e0;color:#333}' +
    '.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:15px;font-size:13px}' +
    '</style></head><body><div class="container">' +
    '<h2>🎨 Comfort View Setup</h2>' +
    '<div class="info">💡 These settings can be undone anytime via the Comfort View Panel or by running "Undo Comfort View"</div>' +
    '<div class="option" onclick="toggle(\'gridlines\')"><input type="checkbox" id="gridlines" checked><div class="option-text"><div class="option-label">Hide Gridlines</div><div class="option-desc">Reduce visual clutter by hiding sheet gridlines</div></div></div>' +
    '<div class="option" onclick="toggle(\'zebra\')"><input type="checkbox" id="zebra"><div class="option-text"><div class="option-label">Zebra Stripes</div><div class="option-desc">Alternating row colors for easier reading</div></div></div>' +
    '<div class="option" onclick="toggle(\'fontSize\')"><input type="checkbox" id="fontSize"><div class="option-text"><div class="option-label">Larger Font (12pt)</div><div class="option-desc">Increase default font size for better readability</div></div></div>' +
    '<div class="option" onclick="toggle(\'focus\')"><input type="checkbox" id="focus"><div class="option-text"><div class="option-label">Focus Mode</div><div class="option-desc">Hide all sheets except the active one</div></div></div>' +
    '<div class="buttons">' +
    '<button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="primary" onclick="apply()">Apply Settings</button>' +
    '</div></div>' +
    '<script>' +
    'function toggle(id){var cb=document.getElementById(id);cb.checked=!cb.checked}' +
    'function apply(){' +
    'var opts={gridlines:document.getElementById("gridlines").checked,zebra:document.getElementById("zebra").checked,fontSize:document.getElementById("fontSize").checked,focus:document.getElementById("focus").checked};' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()}).applyADHDDefaultsWithOptions(opts)}' +
    '</script></body></html>'
  ).setWidth(500).setHeight(450);
  ui.showModalDialog(html, '🎨 Comfort View Setup');
}

/**
 * Apply Comfort View defaults with selected options
 * @param {Object} options - Selected options
 */
function applyADHDDefaultsWithOptions(options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var applied = [];

  try {
    // Store previous settings for undo
    var prevSettings = {
      gridlinesWereHidden: [],
      zebraWasApplied: false,
      previousFontSize: 10
    };

    var sheets = ss.getSheets();

    if (options.gridlines) {
      sheets.forEach(function(sheet) {
        var name = sheet.getName();
        if (name !== SHEETS.CONFIG && name !== SHEETS.MEMBER_DIR && name !== SHEETS.GRIEVANCE_LOG) {
          sheet.setHiddenGridlines(true);
        }
      });
      applied.push('✅ Gridlines hidden on dashboard sheets');
    }

    if (options.zebra) {
      sheets.forEach(function(sheet) {
        applyZebraStripes(sheet);
      });
      saveADHDSettings({zebraStripes: true});
      applied.push('✅ Zebra stripes applied');
    }

    if (options.fontSize) {
      sheets.forEach(function(sheet) {
        if (sheet.getLastRow() > 0) {
          sheet.getDataRange().setFontSize(12);
        }
      });
      saveADHDSettings({fontSize: 12});
      applied.push('✅ Font size increased to 12pt');
    }

    if (options.focus) {
      activateFocusMode();
      applied.push('✅ Focus mode activated');
    }

    ss.toast(applied.join('\n'), '🎨 Setup Complete', 5);

  } catch (e) {
    SpreadsheetApp.getUi().alert('⚠️ Error: ' + e.message);
  }
}

/**
 * Undo Comfort View defaults - restore original settings
 */
function undoADHDDefaults() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('↩️ Undo Comfort View',
    'This will:\n\n' +
    '• Show all gridlines\n' +
    '• Remove zebra stripes\n' +
    '• Reset font size to 10pt\n' +
    '• Exit focus mode (show all sheets)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Show all gridlines
    ss.getSheets().forEach(function(sheet) {
      sheet.setHiddenGridlines(false);
    });

    // Remove zebra stripes
    ss.getSheets().forEach(function(sheet) {
      removeZebraStripes(sheet);
    });

    // Reset font size
    ss.getSheets().forEach(function(sheet) {
      if (sheet.getLastRow() > 0) {
        sheet.getDataRange().setFontSize(10);
      }
    });

    // Exit focus mode
    deactivateFocusMode();

    // Reset stored settings
    resetADHDSettings();

    ui.alert('↩️ Undo Complete',
      'Comfort View defaults have been reset:\n\n' +
      '✅ Gridlines restored\n' +
      '✅ Zebra stripes removed\n' +
      '✅ Font size reset to 10pt\n' +
      '✅ Focus mode deactivated',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('⚠️ Error: ' + e.message);
  }
}
/**
 * ============================================================================
 * MOBILE INTERFACE & QUICK ACTIONS
 * ============================================================================
 * Mobile-optimized views and context-aware quick actions
 * Includes automatic device detection for responsive experience
 */

// ==================== MOBILE CONFIGURATION ====================

var MOBILE_CONFIG = {
  MAX_COLUMNS_MOBILE: 8,
  CARD_LAYOUT_ENABLED: true,
  TOUCH_TARGET_SIZE: '44px',
  MOBILE_BREAKPOINT: 768,  // Width in pixels below which is considered mobile
  TABLET_BREAKPOINT: 1024  // Width in pixels below which is considered tablet
};

// ==================== DEVICE DETECTION ====================

/**
 * Shows a smart dashboard that automatically detects the device type
 * and displays the appropriate interface (mobile or desktop)
 */
function showSmartDashboard() {
  var html = HtmlService.createHtmlOutput(getSmartDashboardHtml())
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Dashboard Pend');
}

/**
 * Returns the HTML for the smart dashboard with device detection
 */
function getSmartDashboardHtml() {
  var stats = getMobileDashboardStats();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Responsive container
    '.container{padding:15px;max-width:1200px;margin:0 auto}' +

    // Header - responsive
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +
    '.device-badge{display:inline-block;padding:4px 12px;background:rgba(255,255,255,0.2);border-radius:20px;font-size:11px;margin-top:8px}' +

    // Stats grid - responsive
    '.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,6vw,36px);font-weight:bold;color:#1a73e8}' +
    '.stat-label{font-size:clamp(11px,2.5vw,13px);color:#666;text-transform:uppercase;margin-top:5px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Action buttons - responsive grid
    '.actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}' +
    '.action-btn{background:white;border:none;padding:16px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;' +
    'min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';transition:all 0.2s}' +
    '.action-btn:hover{background:#e8f0fe;transform:translateX(4px)}' +
    '.action-btn:active{transform:scale(0.98)}' +
    '.action-icon{font-size:24px;width:44px;height:44px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px;flex-shrink:0}' +
    '.action-label{font-weight:500}' +
    '.action-desc{font-size:12px;color:#666;margin-top:2px}' +

    // FAB (Floating Action Button)
    '.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;' +
    'border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer;z-index:1000}' +
    '.fab:hover{background:#1557b0}' +

    // Desktop-only elements
    '.desktop-only{display:none}' +

    // Mobile-specific adjustments
    '@media (max-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:1fr}' +
    '  .container{padding:10px}' +
    '  .header{padding:15px}' +
    '}' +

    // Tablet adjustments
    '@media (min-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px) and (max-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '}' +

    // Desktop view
    '@media (min-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(4,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '  .desktop-only{display:block}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header with dynamic device badge
    '<div class="header">' +
    '<h1>📋 Dashboard Pend</h1>' +
    '<div class="subtitle">Pending Actions & Quick Overview</div>' +
    '<div class="device-badge" id="deviceBadge">Detecting device...</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section
    '<div class="stats">' +
    '<div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div>' +
    '</div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<button class="action-btn" onclick="google.script.run.showMobileGrievanceList()">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">View Grievances</div><div class="action-desc">Browse and filter all grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()">' +
    '<div class="action-icon">👤</div>' +
    '<div><div class="action-label">My Cases</div><div class="action-desc">View your assigned grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showQuickActionsMenu()">' +
    '<div class="action-icon">⚡</div>' +
    '<div><div class="action-label">Row Actions</div><div class="action-desc">Quick actions for selected row</div></div>' +
    '</button>' +

    '</div>' +

    // Desktop-only additional info
    '<div class="desktop-only">' +
    '<div class="section-title">ℹ️ Dashboard Info</div>' +
    '<p style="color:#666;font-size:14px;padding:15px;background:white;border-radius:8px;">' +
    'This responsive dashboard automatically adjusts to your screen size. ' +
    'On mobile devices, you\'ll see a touch-optimized interface with larger buttons. ' +
    'Use the menu items above to manage grievances and member information.' +
    '</p>' +
    '</div>' +

    '</div>' +

    // FAB for refresh
    '<button class="fab" onclick="location.reload()" title="Refresh">🔄</button>' +

    // Device detection script
    '<script>' +
    'function detectDevice(){' +
    '  var w=window.innerWidth;' +
    '  var badge=document.getElementById("deviceBadge");' +
    '  var isTouchDevice="ontouchstart" in window||navigator.maxTouchPoints>0;' +
    '  var userAgent=navigator.userAgent.toLowerCase();' +
    '  var isMobileUA=/android|webos|iphone|ipad|ipod|blackberry|iemobile|opera mini/i.test(userAgent);' +
    '  ' +
    '  if(w<' + MOBILE_CONFIG.MOBILE_BREAKPOINT + '||isMobileUA){' +
    '    badge.textContent="📱 Mobile View";' +
    '    badge.style.background="rgba(76,175,80,0.3)";' +
    '  }else if(w<' + MOBILE_CONFIG.TABLET_BREAKPOINT + '){' +
    '    badge.textContent="📱 Tablet View";' +
    '    badge.style.background="rgba(255,152,0,0.3)";' +
    '  }else{' +
    '    badge.textContent="🖥️ Desktop View";' +
    '    badge.style.background="rgba(33,150,243,0.3)";' +
    '  }' +
    '  ' +
    '  if(isTouchDevice){' +
    '    document.body.classList.add("touch-device");' +
    '  }' +
    '}' +
    'detectDevice();' +
    'window.addEventListener("resize",detectDevice);' +
    '</script>' +

    '</body></html>';
}

/**
 * Check if the current context appears to be mobile
 * Note: This is a server-side heuristic based on available info
 * Real detection happens client-side in the HTML
 */
function isMobileContext() {
  // Server-side we can't reliably detect mobile
  // This function exists for potential future use with session properties
  return false;
}

// ==================== MOBILE DASHBOARD ====================

function showMobileDashboard() {
  var stats = getMobileDashboardStats();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"><style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;padding:0;margin:0;background:#f5f5f5}.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px}.header h1{margin:0;font-size:24px}.container{padding:15px}.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:20px}.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center}.stat-value{font-size:32px;font-weight:bold;color:#1a73e8}.stat-label{font-size:13px;color:#666;text-transform:uppercase}.section-title{font-size:16px;font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}.action-btn{background:white;border:none;padding:16px;margin-bottom:10px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}.action-btn:active{transform:scale(0.98)}.action-icon{font-size:24px;width:40px;height:40px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px}.action-label{font-weight:500}.action-desc{font-size:12px;color:#666}.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer}</style></head><body><div class="header"><h1>📱 509 Dashboard</h1><div style="font-size:14px;opacity:0.9">Mobile View</div></div><div class="container"><div class="stats"><div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div><div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div><div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div><div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div></div><div class="section-title">⚡ Quick Actions</div><button class="action-btn" onclick="google.script.run.showMobileGrievanceList()"><div class="action-icon">📋</div><div><div class="action-label">View Grievances</div><div class="action-desc">Browse all grievances</div></div></button><button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()"><div class="action-icon">🔍</div><div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div></button><button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()"><div class="action-icon">👤</div><div><div class="action-label">My Cases</div><div class="action-desc">View assigned grievances</div></div></button></div><button class="fab" onclick="location.reload()">🔄</button></body></html>'
  ).setWidth(400).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Mobile Dashboard');
}

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
    if (daysTo !== null && daysTo !== '' && daysTo < 0 && status === 'Open') stats.overdueGrievances++;
  });
  return stats;
}

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

function showMobileGrievanceList() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
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
    '<script>var all=[];google.script.run.withSuccessHandler(function(data){all=data;render(data)}).getRecentGrievancesForMobile(100);function render(data){var c=document.getElementById("list");if(!data||data.length===0){c.innerHTML="<div style=\'text-align:center;padding:40px;color:#999;grid-column:1/-1\'>No grievances</div>";return}c.innerHTML=data.map(function(g){return"<div class=\'card\'><div class=\'card-header\'><div class=\'card-id\'>#"+g.id+"</div><div class=\'card-status\'>"+(g.status||"Filed")+"</div></div><div class=\'card-row\'><strong>Member:</strong> "+g.memberName+"</div><div class=\'card-row\'><strong>Issue:</strong> "+(g.issueType||"N/A")+"</div><div class=\'card-row\'><strong>Filed:</strong> "+g.filedDate+"</div></div>"}).join("")}function filterStatus(s,btn){document.querySelectorAll(".filter").forEach(function(f){f.classList.remove("active")});btn.classList.add("active");render(s==="all"?all:all.filter(function(g){return g.status===s}))}function filter(q){render(all.filter(function(g){q=q.toLowerCase();return g.id.toLowerCase().indexOf(q)>=0||g.memberName.toLowerCase().indexOf(q)>=0||(g.issueType||"").toLowerCase().indexOf(q)>=0}))}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Grievance List');
}

function showMobileUnifiedSearch() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
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
    '<script>var tab="all";function setTab(t,btn){tab=t;document.querySelectorAll(".tab").forEach(function(tb){tb.classList.remove("active")});btn.classList.add("active");search(document.getElementById("q").value)}function search(q){if(!q||q.length<2){document.getElementById("results").innerHTML="<div class=\'empty-state\'>Type to search...</div>";return}google.script.run.withSuccessHandler(function(data){render(data)}).getMobileSearchData(q,tab)}function render(data){var c=document.getElementById("results");if(!data||data.length===0){c.innerHTML="<div class=\'empty-state\'>No results</div>";return}c.innerHTML=data.map(function(r){return"<div class=\'result-card\'><div class=\'result-title\'>"+(r.type==="member"?"👤 ":"📋 ")+r.title+"</div><div class=\'result-detail\'>"+r.subtitle+"</div>"+(r.detail?"<div class=\'result-detail\'>"+r.detail+"</div>":"")+"</div>"}).join("")}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search');
}

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

// ==================== QUICK ACTIONS ====================

/**
 * Show context-aware Quick Actions menu
 *
 * HOW IT WORKS:
 * Quick Actions provides contextual shortcuts based on your current selection.
 *
 * AVAILABLE ON:
 * - Member Directory: Start new grievance, send email, view history, copy ID
 * - Grievance Log: Sync to calendar, setup folder, update status, copy ID
 *
 * HOW TO USE:
 * 1. Navigate to Member Directory or Grievance Log
 * 2. Click on any data row (not the header)
 * 3. Run Quick Actions from the menu
 * 4. A popup will show relevant actions for that row
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
      '<button class="action-btn" onclick="google.script.run.composeEmailForMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Send Custom Email</div><div class="desc">Compose email to ' + email + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailDashboardLinkToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">🔗</span><span><div class="title">Send Dashboard Link</div><div class="desc">Share dashboard access with member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#1a73e8;display:flex;align-items:center;gap:10px}' +
    '.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}' +
    '.name{font-size:18px;font-weight:bold}' +
    '.id{color:#666;font-size:14px}' +
    '.status{margin-top:10px}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}' +
    '.open{background:#ffebee;color:#c62828}' +
    '.none{background:#e8f5e9;color:#2e7d32}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}' +
    '.action-btn:hover{background:#e8f4fd}' +
    '.icon{font-size:24px}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:12px;color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#1a73e8;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Quick Actions</h2>' +
    '<div class="info">' +
    '<div class="name">' + name + '</div>' +
    '<div class="id">' + memberId + ' | ' + (email || 'No email') + '</div>' +
    '<div class="status">' + (hasOpen === 'Yes' ? '<span class="badge open">🔴 Has Open Grievance</span>' : '<span class="badge none">🟢 No Open Grievances</span>') + '</div>' +
    '</div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Member Actions</div>' +
    '<button class="action-btn" onclick="google.script.run.openGrievanceFormForMember(' + row + ');google.script.host.close()"><span class="icon">📋</span><span><div class="title">Start New Grievance</div><div class="desc">Create a grievance for this member</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.showMemberGrievanceHistory(\'' + memberId + '\');google.script.host.close()"><span class="icon">📁</span><span><div class="title">View Grievance History</div><div class="desc">See all grievances for this member</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(\'' + memberId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Member ID</div><div class="desc">' + memberId + '</div></span></button>' +
    emailButtons +
    '</div>' +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(email ? 650 : 400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Quick Actions');
}

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
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailGrievanceStatusToMember(\'' + grievanceId + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Email Status to Member</div><div class="desc">Send grievance status update to ' + memberEmail + '</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailSurveyToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📊</span><span><div class="title">Send Satisfaction Survey</div><div class="desc">Email survey link to member</div></span></button>' +
      '<button class="action-btn" onclick="google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(e){alert(e.message)}).emailContactFormToMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📝</span><span><div class="title">Send Contact Update Form</div><div class="desc">Request info update from member</div></span></button>';
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#DC2626}' +
    '.info{background:#fff5f5;padding:15px;border-radius:8px;margin-bottom:20px;border-left:4px solid #DC2626}' +
    '.gid{font-size:18px;font-weight:bold}' +
    '.gmem{color:#666;font-size:14px}' +
    '.gstatus{margin-top:10px;display:flex;gap:10px;flex-wrap:wrap}' +
    '.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}' +
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}' +
    '.action-btn:hover{background:#fff5f5}' +
    '.icon{font-size:24px}' +
    '.title{font-weight:bold}' +
    '.desc{font-size:12px;color:#666;margin-top:2px}' +
    '.section-header{font-weight:bold;color:#DC2626;margin:15px 0 10px;padding-top:10px;border-top:1px solid #e0e0e0}' +
    '.divider{height:1px;background:#e0e0e0;margin:10px 0}' +
    '.status-section{margin-top:15px;padding:15px;background:#f8f9fa;border-radius:8px}' +
    '.status-section h4{margin:0 0 10px}' +
    'select{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px}' +
    '.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Grievance Actions</h2>' +
    '<div class="info">' +
    '<div class="gid">' + grievanceId + '</div>' +
    '<div class="gmem">' + name + ' (' + memberId + ')' + (memberEmail ? ' - ' + memberEmail : '') + '</div>' +
    '<div class="gstatus">' +
    '<span class="badge">' + status + '</span>' +
    '<span class="badge">' + step + '</span>' +
    (daysTo !== null && daysTo !== '' ? '<span class="badge" style="background:' + (daysTo < 0 ? '#ffebee;color:#c62828' : '#e3f2fd;color:#1565c0') + '">' + (daysTo < 0 ? '⚠️ Overdue' : '📅 ' + daysTo + ' days') + '</span>' : '') +
    '</div></div>' +
    '<div class="actions">' +
    '<div class="section-header">📋 Case Management</div>' +
    '<button class="action-btn" onclick="google.script.run.syncSingleGrievanceToCalendar(\'' + grievanceId + '\');google.script.host.close()"><span class="icon">📅</span><span><div class="title">Sync to Calendar</div><div class="desc">Add deadlines to Google Calendar</div></span></button>' +
    '<button class="action-btn" onclick="google.script.run.setupDriveFolderForGrievance();google.script.host.close()"><span class="icon">📁</span><span><div class="title">Setup Drive Folder</div><div class="desc">Create document folder</div></span></button>' +
    '<button class="action-btn" onclick="navigator.clipboard.writeText(\'' + grievanceId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Grievance ID</div><div class="desc">' + grievanceId + '</div></span></button>' +
    emailStatusBtn +
    '</div>' +
    (isOpen ? '<div class="status-section"><h4>Quick Status Update</h4><select id="statusSelect"><option value="">-- Select --</option><option value="Open">Open</option><option value="Pending Info">Pending Info</option><option value="Settled">Settled</option><option value="Withdrawn">Withdrawn</option><option value="Won">Won</option><option value="Denied">Denied</option><option value="Closed">Closed</option></select><button class="action-btn" style="margin-top:10px" onclick="var s=document.getElementById(\'statusSelect\').value;if(!s){alert(\'Select status\');return}google.script.run.withSuccessHandler(function(){alert(\'Updated!\');google.script.host.close()}).quickUpdateGrievanceStatus(' + row + ',s)"><span class="icon">✓</span><span><div class="title">Update Status</div></span></button></div>' : '') +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '</div></body></html>'
  ).setWidth(400).setHeight(memberEmail ? 750 : 550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance Quick Actions');
}

function quickUpdateGrievanceStatus(row, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  sheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newStatus);
  if (['Closed', 'Settled', 'Withdrawn'].indexOf(newStatus) >= 0) {
    var closeCol = GRIEVANCE_COLS.DATE_CLOSED;
    if (!sheet.getRange(row, closeCol).getValue()) sheet.getRange(row, closeCol).setValue(new Date());
  }
  ss.toast('Grievance status updated to: ' + newStatus, 'Status Updated', 3);
}

function composeEmailForMember(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      var email = data[i][MEMBER_COLS.EMAIL - 1];
      var name = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      if (!email) { SpreadsheetApp.getUi().alert('No email on file.'); return; }
      var html = HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}.form-group{margin:15px 0}label{display:block;font-weight:bold;margin-bottom:5px}input,textarea{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box}textarea{min-height:200px}input:focus,textarea:focus{outline:none;border-color:#1a73e8}.buttons{display:flex;gap:10px;margin-top:20px}button{padding:12px 24px;font-size:14px;border:none;border-radius:4px;cursor:pointer;flex:1}.primary{background:#1a73e8;color:white}.secondary{background:#6c757d;color:white}</style></head><body><div class="container"><h2>📧 Email to Member</h2><div class="info"><strong>' + name + '</strong> (' + memberId + ')<br>' + email + '</div><div class="form-group"><label>Subject:</label><input type="text" id="subject" placeholder="Email subject"></div><div class="form-group"><label>Message:</label><textarea id="message" placeholder="Type your message..."></textarea></div><div class="buttons"><button class="primary" onclick="send()">📤 Send</button><button class="secondary" onclick="google.script.host.close()">Cancel</button></div></div><script>function send(){var s=document.getElementById("subject").value.trim();var m=document.getElementById("message").value.trim();if(!s||!m){alert("Fill in subject and message");return}google.script.run.withSuccessHandler(function(){alert("Email sent!");google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).sendQuickEmail("' + email + '",s,m,"' + memberId + '")}</script></body></html>'
      ).setWidth(600).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(html, '📧 Compose Email');
      return;
    }
  }
}

function sendQuickEmail(to, subject, body, memberId) {
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'SEIU Local 509 Dashboard' });
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardUrl = ss.getUrl();
  var orgName = getOrgNameFromConfig_();

  var subject = orgName + ' - Member Dashboard Access';
  var body = 'Dear ' + memberData.firstName + ',\n\n' +
    'You can access your union member dashboard and track your information at:\n\n' +
    'Dashboard Link: ' + dashboardUrl + '\n\n' +
    'From the dashboard you can:\n' +
    '- View your member profile\n' +
    '- Track grievance status (if applicable)\n' +
    '- Stay updated on union activities\n\n' +
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

  if (!grievanceSheet) {
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
    if (grievance.daysToDeadline < 0) {
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

/**
 * Helper: Get member data by Member ID
 * @private
 */
function getMemberDataById_(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return null;

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

/**
 * Helper: Get organization name from Config
 * @private
 */
function getOrgNameFromConfig_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (configSheet) {
    var orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue();
    if (orgName) return orgName;
  }
  return 'SEIU Local 509';
}

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
    return '<div style="background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;border-left:4px solid ' + (g.status === 'Open' ? '#f44336' : '#4caf50') + '"><strong>' + g.id + '</strong><br><span style="color:#666">Status: ' + g.status + ' | Step: ' + g.step + '</span><br><span style="color:#888;font-size:12px">' + g.issue + ' | Filed: ' + (g.filed ? new Date(g.filed).toLocaleDateString() : 'N/A') + '</span></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}</style></head><body><h2>📁 Grievance History</h2><div class="summary"><strong>Member ID:</strong> ' + memberId + '<br><strong>Total:</strong> ' + mine.length + '<br><strong>Open:</strong> ' + mine.filter(function(g) { return g.status === 'Open'; }).length + '<br><strong>Closed:</strong> ' + mine.filter(function(g) { return g.status !== 'Open'; }).length + '</div>' + list + '</body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance History - ' + memberId);
}

function openGrievanceFormForMember(row) {
  SpreadsheetApp.getUi().alert('ℹ️ New Grievance', 'To start a new grievance for this member, navigate to the Grievance Log sheet and add a new row with their Member ID.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncSingleGrievanceToCalendar(grievanceId) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing ' + grievanceId + '...', 'Calendar', 3);
  if (typeof syncDeadlinesToCalendar === 'function') syncDeadlinesToCalendar();
}

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║   ██████╗ ██████╗  ██████╗ ████████╗███████╗ ██████╗████████╗███████╗██████╗  ║
// ║   ██╔══██╗██╔══██╗██╔═══██╗╚══██╔══╝██╔════╝██╔════╝╚══██╔══╝██╔════╝██╔══██╗ ║
// ║   ██████╔╝██████╔╝██║   ██║   ██║   █████╗  ██║        ██║   █████╗  ██║  ██║ ║
// ║   ██╔═══╝ ██╔══██╗██║   ██║   ██║   ██╔══╝  ██║        ██║   ██╔══╝  ██║  ██║ ║
// ║   ██║     ██║  ██║╚██████╔╝   ██║   ███████╗╚██████╗   ██║   ███████╗██████╔╝ ║
// ║   ╚═╝     ╚═╝  ╚═╝ ╚═════╝    ╚═╝   ╚══════╝ ╚═════╝   ╚═╝   ╚══════╝╚═════╝  ║
// ║                                                                           ║
// ║         ⚠️  DO NOT MODIFY THIS SECTION - PROTECTED CODE  ⚠️              ║
// ║                                                                           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║  INTERACTIVE DASHBOARD TAB - Modal Popup with Tabbed Interface           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║                                                                           ║
// ║  This code block is PROTECTED and should NOT be modified or removed.     ║
// ║                                                                           ║
// ║  Protected Functions:                                                     ║
// ║  • showInteractiveDashboardTab() - Opens the modal dialog                 ║
// ║  • getInteractiveDashboardHtml() - Returns the HTML/CSS/JS for the UI     ║
// ║  • getInteractiveOverviewData()  - Fetches overview statistics            ║
// ║  • getInteractiveMemberData()    - Fetches member list data               ║
// ║  • getInteractiveGrievanceData() - Fetches grievance list data            ║
// ║  • getInteractiveAnalyticsData() - Fetches analytics/charts data          ║
// ║                                                                           ║
// ║  Features:                                                                ║
// ║  • 4 Tabs: Overview, Members, Grievances, Analytics                       ║
// ║  • Live search and status filtering                                       ║
// ║  • Mobile-responsive design with touch targets                            ║
// ║  • Bar charts for status distribution and categories                      ║
// ║                                                                           ║
// ║  Menu Location: 👤 Dashboard > 🎯 Custom View                  ║
// ║                                                                           ║
// ║  Added: December 29, 2025 (commit c75c1cc)                                ║
// ║  Status: USER APPROVED - DO NOT CHANGE                                    ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Interactive dashboard is now consolidated into Steward Dashboard.
 */
function showInteractiveDashboardTab() {
  showStewardDashboard();
}

/** @deprecated - Legacy function kept for reference */
function showInteractiveDashboardTab_LEGACY() {
  var html = HtmlService.createHtmlOutput(getInteractiveDashboardHtml())
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Dashboard');
}

/**
 * Returns the HTML for the interactive dashboard with tabs
 */
function getInteractiveDashboardHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(11px,2.5vw,13px);opacity:0.9}' +

    // Tab navigation
    '.tabs{display:flex;background:white;border-bottom:2px solid #e0e0e0;position:sticky;top:0;z-index:100}' +
    '.tab{flex:1;padding:clamp(10px,2.5vw,14px);text-align:center;font-size:clamp(11px,2.5vw,13px);font-weight:600;color:#666;' +
    'border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.tab:hover{background:#f8f9fa;color:#7C3AED}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED;background:#f8f4ff}' +
    '.tab-icon{display:block;font-size:16px;margin-bottom:2px}' +

    // Tab content
    '.tab-content{display:none;padding:15px;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0}to{opacity:1}}' +

    // Stats grid
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:10px;margin-bottom:15px}' +
    '.stat-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s;cursor:pointer}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(20px,4vw,28px);font-weight:bold;color:#7C3AED}' +
    '.stat-label{font-size:clamp(9px,1.8vw,11px);color:#666;text-transform:uppercase;margin-top:4px}' +
    '.stat-card.green .stat-value{color:#059669}' +
    '.stat-card.red .stat-value{color:#DC2626}' +
    '.stat-card.orange .stat-value{color:#F97316}' +

    // Data table
    '.data-table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)}' +
    '.data-table th{background:#7C3AED;color:white;padding:12px;text-align:left;font-size:13px}' +
    '.data-table td{padding:12px;border-bottom:1px solid #eee;font-size:13px}' +
    '.data-table tr:hover{background:#f8f4ff}' +
    '.data-table tr:last-child td{border-bottom:none}' +

    // Status badges
    '.badge{display:inline-block;padding:4px 10px;border-radius:12px;font-size:11px;font-weight:bold}' +
    '.badge-open{background:#fee2e2;color:#dc2626}' +
    '.badge-pending{background:#fef3c7;color:#d97706}' +
    '.badge-closed{background:#d1fae5;color:#059669}' +
    '.badge-overdue{background:#7f1d1d;color:#fecaca}' +
    '.badge-steward{background:#ddd6fe;color:#7c3aed}' +

    // Action buttons
    '.action-btn{display:inline-flex;align-items:center;gap:6px;padding:8px 14px;border:none;border-radius:8px;' +
    'cursor:pointer;font-size:12px;font-weight:500;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.action-btn-primary{background:#7C3AED;color:white}' +
    '.action-btn-primary:hover{background:#5B21B6}' +
    '.action-btn-secondary{background:#f3f4f6;color:#374151}' +
    '.action-btn-secondary:hover{background:#e5e7eb}' +
    '.action-btn-danger{background:#dc2626;color:white}' +
    '.action-btn-danger:hover{background:#b91c1c}' +
    '.action-btn.active{background:#7C3AED;color:white}' +

    // List items - clickable
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:white;padding:15px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.06);cursor:pointer;transition:all 0.2s}' +
    '.list-item:hover{box-shadow:0 4px 12px rgba(0,0,0,0.12);transform:translateY(-1px)}' +
    '.list-item-header{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap}' +
    '.list-item-main{flex:1;min-width:180px}' +
    '.list-item-title{font-weight:600;color:#1f2937;margin-bottom:3px}' +
    '.list-item-subtitle{font-size:12px;color:#666}' +
    '.list-item-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:12px;color:#374151}' +
    '.list-item.expanded .list-item-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#1f2937;font-weight:500}' +
    '.detail-actions{margin-top:10px;display:flex;gap:8px;flex-wrap:wrap}' +

    // Filter dropdowns
    '.filter-bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}' +
    '.filter-select{padding:8px 12px;border:2px solid #e5e7eb;border-radius:6px;font-size:12px;background:white;cursor:pointer}' +
    '.filter-select:focus{outline:none;border-color:#7C3AED}' +

    // Search input
    '.search-container{position:relative;margin-bottom:12px}' +
    '.search-input{width:100%;padding:10px 10px 10px 36px;border:2px solid #e5e7eb;border-radius:8px;font-size:13px;transition:border-color 0.2s}' +
    '.search-input:focus{outline:none;border-color:#7C3AED}' +
    '.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:14px;color:#9ca3af}' +

    // Resource links
    '.resource-links{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-top:15px}' +
    '.resource-links h3{font-size:14px;color:#1f2937;margin-bottom:12px}' +
    '.link-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}' +
    '.resource-link{display:flex;align-items:center;gap:8px;padding:10px;background:#f8f4ff;border-radius:8px;text-decoration:none;color:#7C3AED;font-size:12px;font-weight:500;transition:all 0.2s}' +
    '.resource-link:hover{background:#7C3AED;color:white}' +

    // Charts section
    '.chart-container{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:12px}' +
    '.chart-title{font-weight:600;color:#1f2937;margin-bottom:12px;font-size:13px}' +
    '.bar-chart{display:flex;flex-direction:column;gap:8px}' +
    '.bar-row{display:flex;align-items:center;gap:8px}' +
    '.bar-label{width:90px;font-size:11px;color:#666;text-align:right}' +
    '.bar-container{flex:1;background:#e5e7eb;border-radius:4px;height:20px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.bar-value{width:40px;font-size:11px;font-weight:600;color:#374151}' +

    // Empty state
    '.empty-state{text-align:center;padding:30px;color:#9ca3af}' +
    '.empty-state-icon{font-size:40px;margin-bottom:8px}' +

    // Loading
    '.loading{text-align:center;padding:30px;color:#666}' +
    '.spinner{display:inline-block;width:20px;height:20px;border:3px solid #e5e7eb;border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Error state
    '.error-state{text-align:center;padding:30px;color:#dc2626;background:#fef2f2;border-radius:8px;margin:10px;border:1px solid #fecaca}' +
    '.error-state::before{content:"⚠️ ";font-size:20px}' +
    '.loading-state{text-align:center;padding:40px;color:#6b7280}' +
    '.loading-spinner{display:inline-block;width:24px;height:24px;border:3px solid #e5e7eb;border-top-color:#7c3aed;border-radius:50%;animation:spin 1s linear infinite;margin-bottom:10px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.debug-info{font-size:10px;color:#9ca3af;margin-top:5px;font-family:monospace}' +

    // Sankey Diagram
    '.sankey-container{position:relative;padding:15px 0}' +
    '.sankey-nodes{display:flex;justify-content:space-between;position:relative;z-index:2}' +
    '.sankey-column{display:flex;flex-direction:column;gap:6px;align-items:center}' +
    '.sankey-node{padding:8px 12px;border-radius:6px;color:white;font-weight:600;font-size:11px;text-align:center;min-width:70px;box-shadow:0 2px 4px rgba(0,0,0,0.2)}' +
    '.sankey-node.source{background:linear-gradient(135deg,#7C3AED,#9333EA)}' +
    '.sankey-node.status-open{background:linear-gradient(135deg,#dc2626,#ef4444)}' +
    '.sankey-node.status-pending{background:linear-gradient(135deg,#f97316,#fb923c)}' +
    '.sankey-node.status-closed{background:linear-gradient(135deg,#059669,#10b981)}' +
    '.sankey-node.resolution{background:linear-gradient(135deg,#1d4ed8,#3b82f6)}' +
    '.sankey-flows{position:absolute;top:0;left:0;right:0;bottom:0;z-index:1}' +
    '.sankey-flow{position:absolute;height:4px;border-radius:2px;opacity:0.6;transition:opacity 0.2s}' +
    '.sankey-flow:hover{opacity:1}' +
    '.sankey-label{font-size:10px;color:#666;margin-top:4px}' +
    '.sankey-legend{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;margin-top:12px;padding-top:12px;border-top:1px solid #e5e7eb}' +
    '.sankey-legend-item{display:flex;align-items:center;gap:4px;font-size:10px;color:#666}' +
    '.sankey-legend-color{width:10px;height:10px;border-radius:2px}' +

    // Responsive
    '@media (max-width:600px){' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .list-item-header{flex-direction:column;align-items:flex-start}' +
    '  .tab-icon{font-size:14px}' +
    '  .filter-bar{flex-direction:column}' +
    '  .filter-select{width:100%}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header
    '<div class="header">' +
    '<h1>📊 Dashboard</h1>' +
    '<div class="subtitle">Real-time union data at your fingertips</div>' +
    '</div>' +

    // Tab Navigation (6 tabs now - including My Cases)
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\',this)" id="tab-overview"><span class="tab-icon">📊</span>Overview</button>' +
    '<button class="tab" onclick="switchTab(\'mycases\',this)" id="tab-mycases"><span class="tab-icon">👤</span>My Cases</button>' +
    '<button class="tab" onclick="switchTab(\'members\',this)" id="tab-members"><span class="tab-icon">👥</span>Members</button>' +
    '<button class="tab" onclick="switchTab(\'grievances\',this)" id="tab-grievances"><span class="tab-icon">📋</span>Grievances</button>' +
    '<button class="tab" onclick="switchTab(\'analytics\',this)" id="tab-analytics"><span class="tab-icon">📈</span>Analytics</button>' +
    '<button class="tab" onclick="switchTab(\'resources\',this)" id="tab-resources"><span class="tab-icon">🔗</span>Links</button>' +
    '</div>' +

    // Overview Tab
    '<div class="tab-content active" id="content-overview">' +
    '<div class="stats-grid" id="overview-stats"><div class="loading"><div class="spinner"></div><p>Loading stats...</p></div></div>' +
    '<div id="overview-actions" style="margin-top:12px;display:flex;flex-wrap:wrap;gap:8px"></div>' +
    '<div id="overview-overdue" style="margin-top:15px"></div>' +
    '</div>' +

    // My Cases Tab - Shows steward's assigned grievances
    '<div class="tab-content" id="content-mycases">' +
    '<div class="section-card" style="background:linear-gradient(135deg,#f0f4ff,#e8f0fe);border-left:4px solid #7C3AED;margin-bottom:15px">' +
    '<div style="display:flex;align-items:center;gap:10px"><span style="font-size:24px">👤</span><div><strong>My Assigned Cases</strong><div style="font-size:12px;color:#666">Grievances where you are the assigned steward</div></div></div>' +
    '</div>' +
    '<div class="filter-bar" id="mycases-filter-bar">' +
    '<button class="action-btn action-btn-primary active" data-filter="all" onclick="filterMyCasesStatus(\'all\',this)">All My Cases</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Open" onclick="filterMyCasesStatus(\'Open\',this)">Open</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Pending Info" onclick="filterMyCasesStatus(\'Pending Info\',this)">Pending</button>' +
    '<button class="action-btn action-btn-danger" data-filter="Overdue" onclick="filterMyCasesStatus(\'Overdue\',this)">⚠️ Overdue</button>' +
    '</div>' +
    '<div id="mycases-stats" style="margin-bottom:15px"></div>' +
    '<div class="list-container" id="mycases-list"><div class="loading"><div class="spinner"></div><p>Loading your cases...</p></div></div>' +
    '</div>' +

    // Members Tab
    '<div class="tab-content" id="content-members">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="member-search" placeholder="Search by name, ID, title, location..." oninput="filterMembers()"></div>' +
    '<div class="filter-bar" id="member-filters"></div>' +
    '<div style="margin-bottom:12px"><button class="action-btn action-btn-primary" onclick="showAddMemberForm()">➕ Add New Member</button></div>' +
    '<div class="list-container" id="members-list"><div class="loading"><div class="spinner"></div><p>Loading members...</p></div></div>' +
    // Add Member Form Modal (hidden initially)
    '<div id="member-form-modal" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000;overflow-y:auto;padding:20px">' +
    '<div style="background:white;max-width:500px;margin:20px auto;border-radius:12px;padding:20px;box-shadow:0 10px 40px rgba(0,0,0,0.2)">' +
    '<h3 id="member-form-title" style="margin:0 0 15px;color:#7C3AED">➕ Add New Member</h3>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">First Name *</label><input type="text" id="form-firstName" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter first name"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Last Name *</label><input type="text" id="form-lastName" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter last name"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Job Title</label><input type="text" id="form-jobTitle" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter job title"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Email</label><input type="email" id="form-email" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter email address"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Phone</label><input type="tel" id="form-phone" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter phone number"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Work Location</label><select id="form-location" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select location...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Unit</label><select id="form-unit" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="">Select unit...</option></select></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Office Days</label><select id="form-officeDays" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" multiple size="3"><option value="Monday">Monday</option><option value="Tuesday">Tuesday</option><option value="Wednesday">Wednesday</option><option value="Thursday">Thursday</option><option value="Friday">Friday</option></select><small style="color:#999;font-size:10px">Hold Ctrl/Cmd to select multiple days</small></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Supervisor</label><input type="text" id="form-supervisor" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px" placeholder="Enter supervisor name"></div>' +
    '<div class="form-group" style="margin-bottom:12px"><label style="display:block;font-size:12px;color:#666;margin-bottom:4px">Is Steward?</label><select id="form-isSteward" style="width:100%;padding:10px;border:2px solid #e5e7eb;border-radius:6px;font-size:14px"><option value="No">No</option><option value="Yes">Yes</option></select></div>' +
    '<input type="hidden" id="form-memberId" value="">' +
    '<input type="hidden" id="form-mode" value="add">' +
    '<div style="display:flex;gap:10px;margin-top:20px">' +
    '<button class="action-btn action-btn-primary" style="flex:1" onclick="saveMemberForm()">💾 Save Member</button>' +
    '<button class="action-btn action-btn-secondary" style="flex:1" onclick="closeMemberForm()">Cancel</button>' +
    '</div>' +
    '</div></div>' +
    '</div>' +

    // Grievances Tab
    '<div class="tab-content" id="content-grievances">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="grievance-search" placeholder="Search by ID, member name, issue..." oninput="filterGrievances()"></div>' +
    '<div class="filter-bar" id="grievance-filter-bar">' +
    '<button class="action-btn action-btn-primary active" data-filter="all" onclick="filterGrievanceStatus(\'all\',this)">All</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Open" onclick="filterGrievanceStatus(\'Open\',this)">Open</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Pending Info" onclick="filterGrievanceStatus(\'Pending Info\',this)">Pending</button>' +
    '<button class="action-btn action-btn-danger" data-filter="Overdue" onclick="filterGrievanceStatus(\'Overdue\',this)">⚠️ Overdue</button>' +
    '<button class="action-btn action-btn-secondary" data-filter="Closed" onclick="filterGrievanceStatus(\'Closed\',this)">Closed</button>' +
    '</div>' +
    '<div class="list-container" id="grievances-list"><div class="loading"><div class="spinner"></div><p>Loading grievances...</p></div></div>' +
    '</div>' +

    // Analytics Tab
    '<div class="tab-content" id="content-analytics">' +
    '<div id="analytics-charts"><div class="loading"><div class="spinner"></div><p>Loading analytics...</p></div></div>' +
    '</div>' +

    // Resources Tab
    '<div class="tab-content" id="content-resources">' +
    '<div id="resources-content"><div class="loading"><div class="spinner"></div><p>Loading links...</p></div></div>' +
    '</div>' +

    // JavaScript
    '<script>' +
    'var allMembers=[];var allGrievances=[];var myCases=[];var currentGrievanceFilter="all";var currentMyCasesFilter="all";var memberFilters={location:"all",unit:"all",officeDays:"all"};var resourceLinks={};' +

    // Debug mode and error handler wrapper
    'var DEBUG_MODE=true;' +
    'function log(msg,data){if(DEBUG_MODE){console.log("[Dashboard] "+msg,data||"")}}' +
    'function logError(msg,e){console.error("[Dashboard Error] "+msg,e);if(DEBUG_MODE)alert("Debug: "+msg+"\\n"+e.message)}' +
    'function safeRun(fn,fallback){try{fn()}catch(e){console.error("[Dashboard]",e);if(fallback)fallback(e)}}' +
    'function showLoading(elementId,msg){var el=document.getElementById(elementId);if(el)el.innerHTML="<div class=\\"loading-state\\"><div class=\\"loading-spinner\\"></div><div>"+(msg||"Loading...")+"</div></div>"}' +

    // Tab switching with error handling
    'function switchTab(tabName,btn){' +
    '  safeRun(function(){' +
    '    document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '    document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '    btn.classList.add("active");' +
    '    document.getElementById("content-"+tabName).classList.add("active");' +
    '    if(tabName==="mycases"&&myCases.length===0)loadMyCases();' +
    '    if(tabName==="members"&&allMembers.length===0)loadMembers();' +
    '    if(tabName==="grievances"&&allGrievances.length===0)loadGrievances();' +
    '    if(tabName==="analytics")loadAnalytics();' +
    '    if(tabName==="resources")loadResources();' +
    '  });' +
    '}' +

    // Load overview data with error handling
    'function loadOverview(){' +
    '  log("Loading overview data...");' +
    '  showLoading("overview-stats","Loading dashboard stats...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){' +
    '      log("Overview data received:",data);' +
    '      safeRun(function(){renderOverview(data)},function(e){' +
    '        document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Error rendering stats<div class=\\"debug-info\\">"+e.message+"</div></div>"' +
    '      })' +
    '    })'  +
    '    .withFailureHandler(function(e){' +
    '      logError("Failed to load overview",e);' +
    '      document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Failed to load stats<div class=\\"debug-info\\">"+e.message+"<br>Check: Admin → Modal Diagnostics</div></div>"' +
    '    })' +
    '    .getInteractiveOverviewData();' +
    '}' +

    // Render overview with overdue section and location breakdown
    'function renderOverview(data){' +
    '  var html="";' +
    '  var colors=["#7C3AED","#059669","#1a73e8","#F97316","#DC2626","#8B5CF6","#10B981","#3B82F6"];' +
    '  html+="<div class=\\"stat-card\\" onclick=\\"switchTab(\'members\',document.getElementById(\'tab-members\'))\\"><div class=\\"stat-value\\">"+data.totalMembers+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.activeStewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\" onclick=\\"switchTab(\'grievances\',document.getElementById(\'tab-grievances\'))\\"><div class=\\"stat-value\\">"+data.totalGrievances+"</div><div class=\\"stat-label\\">Total Grievances</div></div>";' +
    '  html+="<div class=\\"stat-card red\\" onclick=\\"showOpenCases()\\"><div class=\\"stat-value\\">"+data.openGrievances+"</div><div class=\\"stat-label\\">Open Cases</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.pendingInfo+"</div><div class=\\"stat-label\\">Pending Info</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.winRate+"</div><div class=\\"stat-label\\">Win Rate</div></div>";' +
    '  document.getElementById("overview-stats").innerHTML=html;' +
    '  var actions="";' +
    '  actions+="<button class=\\"action-btn action-btn-primary\\" onclick=\\"google.script.run.showMobileUnifiedSearch()\\">🔍 Search</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMobileGrievanceList()\\">📋 All Grievances</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMyAssignedGrievances()\\">👤 My Cases</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"location.reload()\\">🔄 Refresh</button>";' +
    '  document.getElementById("overview-actions").innerHTML=actions;' +
    // Location breakdown with bubble chart
    '  if(data.byLocation&&data.byLocation.length>0){' +
    '    var locHtml="<div class=\\"chart-container\\" style=\\"margin-top:15px\\"><div class=\\"chart-title\\">📍 Members by Location</div>";' +
    '    locHtml+="<div style=\\"display:flex;flex-wrap:wrap;gap:10px;justify-content:center;padding:10px\\">";' +
    '    var maxLoc=Math.max.apply(null,data.byLocation.map(function(l){return l.count}))||1;' +
    '    var totalM=data.totalMembers||1;' +
    '    data.byLocation.forEach(function(loc,idx){' +
    '      var pct=Math.round(loc.count/totalM*100);' +
    '      var size=Math.max(55,Math.min(100,50+(loc.count/maxLoc*50)));' +
    '      var clr=colors[idx%colors.length];' +
    '      locHtml+="<div style=\\"text-align:center;cursor:pointer\\" onclick=\\"switchTab(\'analytics\',document.getElementById(\'tab-analytics\'))\\"><div style=\\"width:"+size+"px;height:"+size+"px;border-radius:50%;background:"+clr+";display:flex;flex-direction:column;align-items:center;justify-content:center;color:white;font-weight:bold;margin:0 auto;box-shadow:0 3px 10px rgba(0,0,0,0.15);transition:transform 0.2s\\" onmouseover=\\"this.style.transform=\'scale(1.05)\'\\" onmouseout=\\"this.style.transform=\'scale(1)\'\\"><span style=\\"font-size:"+(size/3)+"px\\">"+loc.count+"</span><span style=\\"font-size:9px;opacity:0.9\\">"+pct+"%</span></div><div style=\\"font-size:10px;color:#666;margin-top:5px;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap\\">"+loc.name+"</div></div>";' +
    '    });' +
    '    locHtml+="</div></div>";' +
    '    document.getElementById("overview-actions").insertAdjacentHTML("afterend",locHtml);' +
    '  }' +
    '  loadOverduePreview();' +
    '}' +

    // Show open cases - switch to grievances tab with Open filter
    'function showOpenCases(){switchTab("grievances",document.getElementById("tab-grievances"));setTimeout(function(){filterGrievanceStatus("Open",document.querySelector("[data-filter=\\"Open\\"]"))},300)}' +

    // Load overdue preview on overview
    'function loadOverduePreview(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    var overdue=data.filter(function(g){return g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overview-overdue").innerHTML="";return}' +
    '    var html="<div class=\\"chart-container\\" style=\\"border-left:4px solid #dc2626\\"><div class=\\"chart-title\\">⚠️ Overdue Cases ("+overdue.length+")</div>";' +
    '    html+="<div class=\\"list-container\\">";' +
    '    overdue.slice(0,3).forEach(function(g){html+="<div class=\\"list-item\\" onclick=\\"showGrievanceDetail(\'"+g.id+"\')\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+g.id+" - "+g.memberName+"</div><div class=\\"list-item-subtitle\\">"+g.issueType+" • "+g.currentStep+"</div></div><span class=\\"badge badge-overdue\\">Overdue</span></div>"});' +
    '    if(overdue.length>3)html+="<button class=\\"action-btn action-btn-danger\\" style=\\"width:100%;margin-top:8px\\" onclick=\\"switchTab(\'grievances\',document.getElementById(\'tab-grievances\'));setTimeout(function(){filterGrievanceStatus(\'Overdue\',document.querySelector(\'[data-filter=Overdue]\'))},300)\\">View All "+overdue.length+" Overdue Cases</button>";' +
    '    html+="</div></div>";' +
    '    document.getElementById("overview-overdue").innerHTML=html;' +
    '  }).getInteractiveGrievanceData();' +
    '}' +

    // Load my cases (steward's assigned grievances)
    'function loadMyCases(){' +
    '  log("Loading my cases...");' +
    '  showLoading("mycases-list","Loading your assigned cases...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("My cases received:",data?data.length:0);myCases=data||[];renderMyCases(myCases);renderMyCasesStats(data)})'  +
    '    .withFailureHandler(function(e){logError("Failed to load my cases",e);document.getElementById("mycases-list").innerHTML="<div class=\\"error-state\\">Failed to load your cases<div class=\\"debug-info\\">"+e.message+"</div></div>"})' +
    '    .getMyStewardCases();' +
    '}' +

    // Render my cases stats
    'function renderMyCasesStats(data){' +
    '  var total=data.length;' +
    '  var open=data.filter(function(g){return g.status==="Open"}).length;' +
    '  var pending=data.filter(function(g){return g.status==="Pending Info"}).length;' +
    '  var overdue=data.filter(function(g){return g.isOverdue}).length;' +
    '  var html="<div class=\\"stats-grid\\">";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+total+"</div><div class=\\"stat-label\\">Total Assigned</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+open+"</div><div class=\\"stat-label\\">Open</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+pending+"</div><div class=\\"stat-label\\">Pending Info</div></div>";' +
    '  if(overdue>0)html+="<div class=\\"stat-card\\" style=\\"border:2px solid #dc2626\\"><div class=\\"stat-value\\" style=\\"color:#dc2626\\">"+overdue+"</div><div class=\\"stat-label\\">⚠️ Overdue</div></div>";' +
    '  html+="</div>";' +
    '  document.getElementById("mycases-stats").innerHTML=html;' +
    '}' +

    // Render my cases list
    'function renderMyCases(data){' +
    '  var c=document.getElementById("mycases-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">👤</div><p>No cases assigned to you</p><p style=\\"font-size:12px;color:#999;margin-top:8px\\">Cases where you are listed as the steward will appear here</p></div>";return}' +
    '  c.innerHTML=data.map(function(g,i){' +
    '    var badgeClass=g.isOverdue?"badge-overdue":(g.status==="Open"?"badge-open":(g.status==="Pending Info"?"badge-pending":"badge-closed"));' +
    '    var statusText=g.isOverdue?"Overdue":g.status;' +
    '    var priorityBorder=g.isOverdue?"border-left:4px solid #dc2626;":"";' +
    '    return "<div class=\\"list-item\\" style=\\""+priorityBorder+"\\" onclick=\\"toggleMyCaseDetail(this)\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+g.id+" - "+g.memberName+"</div><div class=\\"list-item-subtitle\\">"+g.issueType+" • "+g.currentStep+"</div></div><div><span class=\\"badge "+badgeClass+"\\">"+statusText+"</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+g.filedDate+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span><span class=\\"detail-value\\">"+g.nextActionDue+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+g.daysOpen+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+g.location+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+g.articles+"</span></div>' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" onclick=\\"event.stopPropagation();google.script.run.showGrievanceQuickActions(\'"+g.id+"\')\\">⚡ Quick Actions</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" onclick=\\"event.stopPropagation();google.script.run.navigateToGrievanceInSheet(\'"+g.id+"\')\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '}' +

    // Toggle my case detail
    'function toggleMyCaseDetail(el){el.classList.toggle("expanded")}' +

    // Filter my cases by status
    'function filterMyCasesStatus(status,btn){' +
    '  currentMyCasesFilter=status;' +
    '  document.querySelectorAll("#mycases-filter-bar .action-btn").forEach(function(b){' +
    '    b.classList.remove("active","action-btn-primary");' +
    '    if(b.dataset.filter!=="Overdue")b.classList.add("action-btn-secondary");' +
    '  });' +
    '  if(btn){btn.classList.add("active");if(status!=="Overdue")btn.classList.add("action-btn-primary");btn.classList.remove("action-btn-secondary")}' +
    '  var filtered=myCases;' +
    '  if(status==="Overdue"){filtered=myCases.filter(function(g){return g.isOverdue})}' +
    '  else if(status!=="all"){filtered=myCases.filter(function(g){return g.status===status})}' +
    '  renderMyCases(filtered);' +
    '}' +

    // Load members with filters
    'function loadMembers(){' +
    '  log("Loading members...");' +
    '  showLoading("members-list","Loading member directory...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("Members received:",data?data.length:0);allMembers=data||[];renderMembers(allMembers);loadMemberFilters()})'  +
    '    .withFailureHandler(function(e){logError("Failed to load members",e);document.getElementById("members-list").innerHTML="<div class=\\"error-state\\">Failed to load members<div class=\\"debug-info\\">"+e.message+"</div></div>"})' +
    '    .getInteractiveMemberData();' +
    '}' +

    // Load member filter dropdowns
    'function loadMemberFilters(){' +
    '  var locations={};var units={};var officeDays={};' +
    '  allMembers.forEach(function(m){' +
    '    if(m.location&&m.location!=="N/A")locations[m.location]=1;' +
    '    if(m.unit&&m.unit!=="N/A")units[m.unit]=1;' +
    '    if(m.officeDays&&m.officeDays!=="N/A"){' +
    '      m.officeDays.split(",").forEach(function(d){var day=d.trim();if(day)officeDays[day]=1});' +
    '    }' +
    '  });' +
    '  var html="<select class=\\"filter-select\\" id=\\"filter-location\\" onchange=\\"memberFilters.location=this.value;filterMembers()\\"><option value=\\"all\\">All Locations</option>";' +
    '  Object.keys(locations).sort().forEach(function(l){html+="<option value=\\""+l+"\\">"+l+"</option>"});' +
    '  html+="</select><select class=\\"filter-select\\" id=\\"filter-unit\\" onchange=\\"memberFilters.unit=this.value;filterMembers()\\"><option value=\\"all\\">All Units</option>";' +
    '  Object.keys(units).sort().forEach(function(u){html+="<option value=\\""+u+"\\">"+u+"</option>"});' +
    '  html+="</select><select class=\\"filter-select\\" id=\\"filter-officeDays\\" onchange=\\"memberFilters.officeDays=this.value;filterMembers()\\"><option value=\\"all\\">All Office Days</option>";' +
    '  Object.keys(officeDays).sort(function(a,b){var days=[\"Monday\",\"Tuesday\",\"Wednesday\",\"Thursday\",\"Friday\",\"Saturday\",\"Sunday\"];return days.indexOf(a)-days.indexOf(b)}).forEach(function(d){html+="<option value=\\""+d+"\\">"+d+"</option>"});' +
    '  html+="</select><button class=\\"action-btn action-btn-secondary\\" onclick=\\"resetMemberFilters()\\">Reset</button>";' +
    '  document.getElementById("member-filters").innerHTML=html;' +
    '  populateFormDropdowns(locations,units);' +
    '}' +

    // Reset member filters
    'function resetMemberFilters(){memberFilters={location:"all",unit:"all",officeDays:"all"};document.getElementById("member-search").value="";document.getElementById("filter-location").value="all";document.getElementById("filter-unit").value="all";document.getElementById("filter-officeDays").value="all";renderMembers(allMembers)}' +

    // Render members with clickable details
    'function renderMembers(data){' +
    '  var c=document.getElementById("members-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">👥</div><p>No members found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(m,i){' +
    '    var badge=m.isSteward?"<span class=\\"badge badge-steward\\">Steward</span>":"";' +
    '    if(m.hasOpenGrievance)badge+="<span class=\\"badge badge-open\\" style=\\"margin-left:4px\\">Has Case</span>";' +
    '    return "<div class=\\"list-item\\" onclick=\\"toggleMemberDetail(this,"+i+")\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+m.name+"</div><div class=\\"list-item-subtitle\\">"+m.id+" • "+m.title+"</div></div><div>"+badge+"</div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+m.location+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+m.unit+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+(m.email||"N/A")+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📱 Phone:</span><span class=\\"detail-value\\">"+(m.phone||"N/A")+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Office Days:</span><span class=\\"detail-value\\">"+m.officeDays+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">👤 Supervisor:</span><span class=\\"detail-value\\">"+m.supervisor+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+m.assignedSteward+"</span></div>' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" onclick=\\"event.stopPropagation();showEditMemberForm("+i+")\\">✏️ Edit Member</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" onclick=\\"event.stopPropagation();google.script.run.navigateToMemberInSheet(\'"+m.id+"\')\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" members. Use search to find specific members.</p></div>";' +
    '}' +

    // Toggle member detail
    'function toggleMemberDetail(el,idx){el.classList.toggle("expanded")}' +

    // Filter members with all criteria
    'function filterMembers(){' +
    '  var query=(document.getElementById("member-search").value||"").toLowerCase();' +
    '  var filtered=allMembers.filter(function(m){' +
    '    if(memberFilters.location!=="all"&&m.location!==memberFilters.location)return false;' +
    '    if(memberFilters.unit!=="all"&&m.unit!==memberFilters.unit)return false;' +
    '    if(memberFilters.officeDays!=="all"&&m.officeDays&&m.officeDays.indexOf(memberFilters.officeDays)<0)return false;' +
    '    if(query&&query.length>=2){' +
    '      return m.name.toLowerCase().indexOf(query)>=0||' +
    '             m.id.toLowerCase().indexOf(query)>=0||' +
    '             m.title.toLowerCase().indexOf(query)>=0||' +
    '             m.location.toLowerCase().indexOf(query)>=0||' +
    '             (m.email||"").toLowerCase().indexOf(query)>=0;' +
    '    }' +
    '    return true;' +
    '  });' +
    '  renderMembers(filtered);' +
    '}' +

    // Populate form dropdowns with location/unit options
    'function populateFormDropdowns(locations,units){' +
    '  var locSelect=document.getElementById("form-location");' +
    '  var unitSelect=document.getElementById("form-unit");' +
    '  locSelect.innerHTML="<option value=\\"\\">Select location...</option>";' +
    '  unitSelect.innerHTML="<option value=\\"\\">Select unit...</option>";' +
    '  Object.keys(locations).sort().forEach(function(l){locSelect.innerHTML+="<option value=\\""+l+"\\">"+l+"</option>"});' +
    '  Object.keys(units).sort().forEach(function(u){unitSelect.innerHTML+="<option value=\\""+u+"\\">"+u+"</option>"});' +
    '}' +

    // Show add member form
    'function showAddMemberForm(){' +
    '  document.getElementById("member-form-title").innerHTML="➕ Add New Member";' +
    '  document.getElementById("form-mode").value="add";' +
    '  document.getElementById("form-memberId").value="";' +
    '  document.getElementById("form-firstName").value="";' +
    '  document.getElementById("form-lastName").value="";' +
    '  document.getElementById("form-jobTitle").value="";' +
    '  document.getElementById("form-email").value="";' +
    '  document.getElementById("form-phone").value="";' +
    '  document.getElementById("form-location").value="";' +
    '  document.getElementById("form-unit").value="";' +
    '  document.getElementById("form-supervisor").value="";' +
    '  document.getElementById("form-isSteward").value="No";' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  for(var i=0;i<daysSelect.options.length;i++)daysSelect.options[i].selected=false;' +
    '  document.getElementById("member-form-modal").style.display="block";' +
    '}' +

    // Show edit member form with existing data
    'function showEditMemberForm(idx){' +
    '  var m=allMembers[idx];' +
    '  if(!m)return;' +
    '  document.getElementById("member-form-title").innerHTML="✏️ Edit Member: "+m.name;' +
    '  document.getElementById("form-mode").value="edit";' +
    '  document.getElementById("form-memberId").value=m.id;' +
    '  document.getElementById("form-firstName").value=m.firstName||"";' +
    '  document.getElementById("form-lastName").value=m.lastName||"";' +
    '  document.getElementById("form-jobTitle").value=m.title!=="N/A"?m.title:"";' +
    '  document.getElementById("form-email").value=m.email||"";' +
    '  document.getElementById("form-phone").value=m.phone||"";' +
    '  document.getElementById("form-location").value=m.location!=="N/A"?m.location:"";' +
    '  document.getElementById("form-unit").value=m.unit!=="N/A"?m.unit:"";' +
    '  document.getElementById("form-supervisor").value=m.supervisor!=="N/A"?m.supervisor:"";' +
    '  document.getElementById("form-isSteward").value=m.isSteward?"Yes":"No";' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  var memberDays=m.officeDays&&m.officeDays!=="N/A"?m.officeDays.split(",").map(function(d){return d.trim()}):[];' +
    '  for(var i=0;i<daysSelect.options.length;i++){daysSelect.options[i].selected=memberDays.indexOf(daysSelect.options[i].value)>=0}' +
    '  document.getElementById("member-form-modal").style.display="block";' +
    '}' +

    // Close member form modal
    'function closeMemberForm(){' +
    '  document.getElementById("member-form-modal").style.display="none";' +
    '}' +

    // Save member (add or edit)
    'function saveMemberForm(){' +
    '  var mode=document.getElementById("form-mode").value;' +
    '  var firstName=document.getElementById("form-firstName").value.trim();' +
    '  var lastName=document.getElementById("form-lastName").value.trim();' +
    '  if(!firstName||!lastName){alert("First name and last name are required");return}' +
    '  var daysSelect=document.getElementById("form-officeDays");' +
    '  var selectedDays=[];' +
    '  for(var i=0;i<daysSelect.options.length;i++){if(daysSelect.options[i].selected)selectedDays.push(daysSelect.options[i].value)}' +
    '  var memberData={' +
    '    memberId:document.getElementById("form-memberId").value,' +
    '    firstName:firstName,' +
    '    lastName:lastName,' +
    '    jobTitle:document.getElementById("form-jobTitle").value.trim(),' +
    '    email:document.getElementById("form-email").value.trim(),' +
    '    phone:document.getElementById("form-phone").value.trim(),' +
    '    location:document.getElementById("form-location").value,' +
    '    unit:document.getElementById("form-unit").value,' +
    '    officeDays:selectedDays.join(", "),' +
    '    supervisor:document.getElementById("form-supervisor").value.trim(),' +
    '    isSteward:document.getElementById("form-isSteward").value' +
    '  };' +
    '  var btn=document.querySelector("#member-form-modal .action-btn-primary");' +
    '  btn.disabled=true;btn.innerHTML="⏳ Saving...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result){' +
    '      btn.disabled=false;btn.innerHTML="💾 Save Member";' +
    '      closeMemberForm();' +
    '      alert(mode==="add"?"Member added successfully!":"Member updated successfully!");' +
    '      allMembers=[];loadMembers();' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      btn.disabled=false;btn.innerHTML="💾 Save Member";' +
    '      alert("Error saving member: "+e.message);' +
    '    })' +
    '    .saveInteractiveMember(memberData,mode);' +
    '}' +

    // Load grievances
    'function loadGrievances(){' +
    '  log("Loading grievances...");' +
    '  showLoading("grievances-list","Loading grievance log...");' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){log("Grievances received:",data?data.length:0);allGrievances=data||[];renderGrievances(allGrievances)})'  +
    '    .withFailureHandler(function(e){logError("Failed to load grievances",e);document.getElementById("grievances-list").innerHTML="<div class=\\"error-state\\">Failed to load grievances<div class=\\"debug-info\\">"+e.message+"</div></div>"})' +
    '    .getInteractiveGrievanceData();' +
    '}' +

    // Render grievances with clickable details
    'function renderGrievances(data){' +
    '  var c=document.getElementById("grievances-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">📋</div><p>No grievances found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(g,i){' +
    '    var badgeClass=g.isOverdue?"badge-overdue":(g.status==="Open"?"badge-open":(g.status==="Pending Info"?"badge-pending":"badge-closed"));' +
    '    var statusText=g.isOverdue?"Overdue":g.status;' +
    '    var daysInfo=g.isOverdue?"<span style=\\"color:#dc2626;font-weight:bold\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?""+g.daysToDeadline+" days left":"");' +
    '    return "<div class=\\"list-item\\" onclick=\\"toggleGrievanceDetail(this,"+i+")\\">' +
    '      <div class=\\"list-item-header\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+g.id+" - "+g.memberName+"</div><div class=\\"list-item-subtitle\\">"+g.issueType+" • "+g.currentStep+"</div></div><div><span class=\\"badge "+badgeClass+"\\">"+statusText+"</span></div></div>' +
    '      <div class=\\"list-item-details\\">' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+g.filedDate+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+g.incidentDate+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span><span class=\\"detail-value\\">"+g.nextActionDue+" "+daysInfo+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+g.daysOpen+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+g.location+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+g.articles+"</span></div>' +
    '        <div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+g.steward+"</span></div>' +
    '        "+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+g.resolution+"</span></div>":"")+"' +
    '        <div class=\\"detail-actions\\">' +
    '          <button class=\\"action-btn action-btn-primary\\" onclick=\\"event.stopPropagation();google.script.run.showGrievanceQuickActions(\'"+g.id+"\')\\">⚡ Quick Actions</button>' +
    '          <button class=\\"action-btn action-btn-secondary\\" onclick=\\"event.stopPropagation();google.script.run.navigateToGrievanceInSheet(\'"+g.id+"\')\\">📄 View in Sheet</button>' +
    '        </div>' +
    '      </div>' +
    '    </div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" grievances. Use search/filters to find specific cases.</p></div>";' +
    '}' +

    // Toggle grievance detail
    'function toggleGrievanceDetail(el,idx){el.classList.toggle("expanded")}' +

    // Show specific grievance detail
    'function showGrievanceDetail(id){' +
    '  var g=allGrievances.find(function(x){return x.id===id});' +
    '  if(g){switchTab("grievances",document.getElementById("tab-grievances"));setTimeout(function(){' +
    '    document.getElementById("grievance-search").value=id;filterGrievances();' +
    '    var items=document.querySelectorAll("#grievances-list .list-item");if(items[0])items[0].classList.add("expanded");' +
    '  },300)}' +
    '}' +

    // Filter grievances
    'function filterGrievances(){' +
    '  var query=(document.getElementById("grievance-search").value||"").toLowerCase();' +
    '  var filtered=allGrievances;' +
    '  if(currentGrievanceFilter==="Overdue"){filtered=filtered.filter(function(g){return g.isOverdue})}' +
    '  else if(currentGrievanceFilter!=="all"){filtered=filtered.filter(function(g){return g.status===currentGrievanceFilter})}' +
    '  if(query&&query.length>=2){' +
    '    filtered=filtered.filter(function(g){' +
    '      return g.id.toLowerCase().indexOf(query)>=0||' +
    '             g.memberName.toLowerCase().indexOf(query)>=0||' +
    '             (g.issueType||"").toLowerCase().indexOf(query)>=0||' +
    '             (g.steward||"").toLowerCase().indexOf(query)>=0;' +
    '    });' +
    '  }' +
    '  renderGrievances(filtered);' +
    '}' +

    // Filter by status with button highlighting
    'function filterGrievanceStatus(status,btn){' +
    '  currentGrievanceFilter=status;' +
    '  document.querySelectorAll("#grievance-filter-bar .action-btn").forEach(function(b){' +
    '    b.classList.remove("active","action-btn-primary");' +
    '    if(b.dataset.filter!=="Overdue")b.classList.add("action-btn-secondary");' +
    '  });' +
    '  if(btn){btn.classList.add("active");if(status!=="Overdue")btn.classList.add("action-btn-primary");btn.classList.remove("action-btn-secondary")}' +
    '  filterGrievances();' +
    '}' +

    // Load analytics
    'function loadAnalytics(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){safeRun(function(){renderAnalytics(data)})})'  +
    '    .withFailureHandler(function(e){document.getElementById("analytics-charts").innerHTML="<div class=\\"error-state\\">Failed to load analytics</div>"})' +
    '    .getInteractiveAnalyticsData();' +
    '}' +

    // Load resources/links
    'function loadResources(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){resourceLinks=data||{};renderResources(data)})'  +
    '    .withFailureHandler(function(e){document.getElementById("resources-content").innerHTML="<div class=\\"error-state\\">Failed to load links</div>"})' +
    '    .getInteractiveResourceLinks();' +
    '}' +

    // Render analytics - ENHANCED VERSION
    'function renderAnalytics(data){' +
    '  var c=document.getElementById("analytics-charts");' +
    '  var html="";' +
    '  var colors=["#7C3AED","#059669","#1a73e8","#F97316","#DC2626","#8B5CF6","#10B981","#3B82F6","#F59E0B","#EF4444"];' +

    // ========== GRIEVANCE STATS SECTION ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #7C3AED\\"><div class=\\"chart-title\\">📊 Grievance Statistics</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '  var totalG=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+totalG+"</div><div class=\\"stat-label\\">Total Cases</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.statusCounts.open+"</div><div class=\\"stat-label\\">Open</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.statusCounts.pending+"</div><div class=\\"stat-label\\">Pending</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+(data.grievanceStats?data.grievanceStats.avgDaysToResolve:0)+"</div><div class=\\"stat-label\\">Avg Days to Resolve</div></div>";' +
    '  html+="</div></div>";' +

    // ========== GRIEVANCES BY STATUS (with colors and proportional bars) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📈 Grievances by Status</div><div class=\\"bar-chart\\">";' +
    '  if(data.grievanceStats&&data.grievanceStats.byStatus&&data.grievanceStats.byStatus.length>0){' +
    '    var maxStatus=Math.max.apply(null,data.grievanceStats.byStatus.map(function(s){return s.count}))||1;' +
    '    data.grievanceStats.byStatus.forEach(function(status){' +
    '      var pct=(status.count/maxStatus*100);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">"+status.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+status.color+"\\"></div></div><div class=\\"bar-value\\">"+status.count+"</div></div>";' +
    '    });' +
    '  }else if(totalG>0){' +
    '    var maxS=Math.max(data.statusCounts.open,data.statusCounts.pending,data.statusCounts.closed)||1;' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Open</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.open/maxS*100)+"%;background:#DC2626\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.open+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Pending</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.pending/maxS*100)+"%;background:#F97316\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.pending+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:100px\\">Closed</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.closed/maxS*100)+"%;background:#059669\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.closed+"</div></div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No grievances</div>"}' +
    '  html+="</div></div>";' +

    // ========== GRIEVANCES BY TYPE (Issue Categories) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📋 Grievances by Type (Issue Categories)</div><div class=\\"bar-chart\\">";' +
    '  var catData=data.grievanceStats&&data.grievanceStats.byType?data.grievanceStats.byType:data.topCategories;' +
    '  if(catData&&catData.length>0){' +
    '    var maxCat=Math.max.apply(null,catData.map(function(c){return c.count}))||1;' +
    '    catData.forEach(function(cat,idx){' +
    '      var pct=(cat.count/maxCat*100);' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+cat.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div></div><div class=\\"bar-value\\">"+cat.count+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No issue data</div>"}' +
    '  html+="</div></div>";' +

    // ========== LOCATION BREAKDOWN ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Grievances by Location</div><div class=\\"bar-chart\\">";' +
    '  if(data.grievanceStats&&data.grievanceStats.byLocation&&data.grievanceStats.byLocation.length>0){' +
    '    var maxLoc=Math.max.apply(null,data.grievanceStats.byLocation.map(function(l){return l.total}))||1;' +
    '    data.grievanceStats.byLocation.forEach(function(loc,idx){' +
    '      var pct=(loc.total/maxLoc*100);' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+loc.name+"</div><div class=\\"bar-container\\" style=\\"position:relative\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div>"+(loc.open>0?"<div style=\\"position:absolute;right:8px;top:2px;font-size:9px;color:#dc2626\\">"+loc.open+" open</div>":"")+"</div><div class=\\"bar-value\\">"+loc.total+"</div></div>";' +
    '    });' +
    '  }else{html+="<div class=\\"empty-state\\">No location data</div>"}' +
    '  html+="</div></div>";' +

    // ========== MONTH OVER MONTH TRENDS ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📅 Month Over Month Trends</div>";' +
    '  if(data.grievanceStats&&data.grievanceStats.monthlyTrends&&data.grievanceStats.monthlyTrends.length>0){' +
    '    html+="<div style=\\"display:flex;gap:5px;justify-content:space-around;margin:15px 0\\">";' +
    '    var maxMo=Math.max.apply(null,data.grievanceStats.monthlyTrends.map(function(m){return Math.max(m.filed,m.resolved)}))||1;' +
    '    data.grievanceStats.monthlyTrends.forEach(function(mo){' +
    '      var filedH=Math.max(mo.filed/maxMo*80,5);' +
    '      var resolvedH=Math.max(mo.resolved/maxMo*80,5);' +
    '      html+="<div style=\\"text-align:center;flex:1\\"><div style=\\"display:flex;gap:2px;justify-content:center;align-items:flex-end;height:90px\\">";' +
    '      html+="<div style=\\"width:16px;background:#DC2626;height:"+filedH+"px;border-radius:3px 3px 0 0\\" title=\\"Filed: "+mo.filed+"\\"></div>";' +
    '      html+="<div style=\\"width:16px;background:#059669;height:"+resolvedH+"px;border-radius:3px 3px 0 0\\" title=\\"Resolved: "+mo.resolved+"\\"></div>";' +
    '      html+="</div><div style=\\"font-size:10px;color:#666;margin-top:4px\\">"+mo.month.split("-")[1]+"/"+mo.month.split("-")[0].slice(2)+"</div></div>";' +
    '    });' +
    '    html+="</div><div style=\\"display:flex;justify-content:center;gap:15px;font-size:11px;color:#666\\"><span><span style=\\"display:inline-block;width:10px;height:10px;background:#DC2626;border-radius:2px\\"></span> Filed</span><span><span style=\\"display:inline-block;width:10px;height:10px;background:#059669;border-radius:2px\\"></span> Resolved</span></div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No trend data available</div>"}' +
    '  html+="</div>";' +

    // ========== TOP 10 PERFORMERS BY SCORE ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #059669\\"><div class=\\"chart-title\\">🏆 Top 10 Performers by Score</div>";' +
    '  if(data.stewardPerformance&&data.stewardPerformance.topPerformers&&data.stewardPerformance.topPerformers.length>0){' +
    '    html+="<table style=\\"width:100%;border-collapse:collapse;font-size:12px\\"><tr style=\\"background:#f3f4f6\\"><th style=\\"padding:8px;text-align:left\\">Rank</th><th style=\\"padding:8px;text-align:left\\">Steward</th><th style=\\"padding:8px;text-align:center\\">Score</th><th style=\\"padding:8px;text-align:center\\">Win Rate</th><th style=\\"padding:8px;text-align:center\\">Avg Days</th></tr>";' +
    '    data.stewardPerformance.topPerformers.forEach(function(p,i){' +
    '      var medal=i===0?"🥇":i===1?"🥈":i===2?"🥉":"";' +
    '      var scoreColor=p.score>=70?"#059669":p.score>=50?"#F97316":"#DC2626";' +
    '      html+="<tr style=\\"border-bottom:1px solid #e5e7eb\\"><td style=\\"padding:8px\\">"+medal+(i+1)+"</td><td style=\\"padding:8px\\">"+p.name+"</td><td style=\\"padding:8px;text-align:center;font-weight:bold;color:"+scoreColor+"\\">"+Math.round(p.score)+"</td><td style=\\"padding:8px;text-align:center\\">"+(p.winRate||0)+"%</td><td style=\\"padding:8px;text-align:center\\">"+(p.avgDays||0)+"</td></tr>";' +
    '    });' +
    '    html+="</table>";' +
    '  }else{html+="<div class=\\"empty-state\\">No performance data available.<br><small>Run Data Integrity Check to generate scores.</small></div>"}' +
    '  html+="</div>";' +

    // ========== TOP 10 BUSIEST STEWARDS ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #F97316\\"><div class=\\"chart-title\\">📊 Top 10 Busiest Stewards (Case Load)</div>";' +
    '  if(data.stewardPerformance&&data.stewardPerformance.busiestStewards&&data.stewardPerformance.busiestStewards.length>0){' +
    '    var maxCases=Math.max.apply(null,data.stewardPerformance.busiestStewards.map(function(s){return s.total}))||1;' +
    '    html+="<div class=\\"bar-chart\\">";' +
    '    data.stewardPerformance.busiestStewards.forEach(function(s,idx){' +
    '      var pct=(s.total/maxCases*100);' +
    '      var openPct=(s.open/s.total*100);' +
    '      html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:110px;font-size:11px\\">"+s.name+"</div><div class=\\"bar-container\\" style=\\"position:relative\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:linear-gradient(90deg,#F97316 "+openPct+"%,#059669 "+openPct+"%)\\"></div></div><div class=\\"bar-value\\">"+s.total+" <small style=\\"color:#F97316\\">("+s.open+" open)</small></div></div>";' +
    '    });' +
    '    html+="</div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No steward case data</div>"}' +
    '  html+="</div>";' +

    // ========== SURVEY RESULTS SECTION ==========
    '  html+="<div class=\\"chart-container\\" style=\\"border-left:4px solid #1a73e8\\"><div class=\\"chart-title\\">📊 Survey Results</div>";' +
    '  if(data.surveyResults){' +
    '    html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '    html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.surveyResults.totalResponses+"</div><div class=\\"stat-label\\">Total Responses</div></div>";' +
    '    html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+(data.surveyResults.avgSatisfaction||"-")+"</div><div class=\\"stat-label\\">Avg Satisfaction (1-10)</div></div>";' +
    '    html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+(data.surveyResults.responseRate||0)+"%</div><div class=\\"stat-label\\">Response Rate</div></div>";' +
    '    html+="</div>";' +
    '    if(data.surveyResults.bySection&&data.surveyResults.bySection.length>0){' +
    '      html+="<div class=\\"chart-title\\" style=\\"font-size:12px;margin:10px 0\\">Satisfaction by Section</div><div class=\\"bar-chart\\">";' +
    '      data.surveyResults.bySection.forEach(function(sec,idx){' +
    '        var pct=(sec.avg/10*100);' +
    '        var clr=sec.avg>=7?"#059669":sec.avg>=5?"#F97316":"#DC2626";' +
    '        html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:130px;font-size:11px\\">"+sec.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+pct+"%;background:"+clr+"\\"></div></div><div class=\\"bar-value\\">"+sec.avg+"</div></div>";' +
    '      });' +
    '      html+="</div>";' +
    '    }else{html+="<div style=\\"color:#999;font-size:12px;text-align:center;padding:10px\\">No section data. Complete surveys to see breakdown.</div>"}' +
    '  }else{html+="<div class=\\"empty-state\\">No survey data. Link Google Form to collect responses.</div>"}' +
    '  html+="</div>";' +

    // ========== RESOLUTION SUMMARY ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🏆 Resolution Summary</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin:0\\">";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.resolutions.won+"</div><div class=\\"stat-label\\">Won</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.resolutions.settled+"</div><div class=\\"stat-label\\">Settled</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.resolutions.withdrawn+"</div><div class=\\"stat-label\\">Withdrawn</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.resolutions.denied+"</div><div class=\\"stat-label\\">Denied</div></div>";' +
    '  html+="</div></div>";' +

    // ========== MEMBER DIRECTORY STATS ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👥 Member Directory Statistics</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.total+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.memberStats.stewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.withOpenGrievance+"</div><div class=\\"stat-label\\">With Open Case</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.stewardRatio+"</div><div class=\\"stat-label\\">Member:Steward</div></div>";' +
    '  html+="</div></div>";' +

    // ========== MEMBERS BY LOCATION (improved visualization) ==========
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Members by Location</div>";' +
    '  if(data.memberStats.byLocation&&data.memberStats.byLocation.length>0){' +
    '    var maxLoc=Math.max.apply(null,data.memberStats.byLocation.map(function(l){return l.count}))||1;' +
    '    var totalMembers=data.memberStats.total||1;' +
    '    html+="<div style=\\"display:flex;flex-wrap:wrap;gap:8px;justify-content:center\\">";' +
    '    data.memberStats.byLocation.forEach(function(loc,idx){' +
    '      var pct=Math.round(loc.count/totalMembers*100);' +
    '      var size=Math.max(60,Math.min(120,60+(loc.count/maxLoc*60)));' +
    '      var clr=colors[idx%colors.length];' +
    '      html+="<div style=\\"text-align:center;padding:10px\\"><div style=\\"width:"+size+"px;height:"+size+"px;border-radius:50%;background:"+clr+";display:flex;flex-direction:column;align-items:center;justify-content:center;color:white;font-weight:bold;margin:0 auto\\"><span style=\\"font-size:"+(size/3)+"px\\">"+loc.count+"</span><span style=\\"font-size:10px\\">"+pct+"%</span></div><div style=\\"font-size:11px;color:#666;margin-top:6px;max-width:100px;overflow:hidden;text-overflow:ellipsis\\">"+loc.name+"</div></div>";' +
    '    });' +
    '    html+="</div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No location data</div>"}' +
    '  html+="</div>";' +

    // ========== SANKEY DIAGRAM ==========
    '  var totalGrievances=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  if(totalGrievances>0){' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🔀 Grievance Flow</div>";' +
    '  html+="<div class=\\"sankey-container\\">";' +
    '  html+="<div class=\\"sankey-nodes\\">";' +
    '  html+="<div class=\\"sankey-column\\"><div class=\\"sankey-node source\\">Filed<br/>"+totalGrievances+"</div><div class=\\"sankey-label\\">Total Filed</div></div>";' +
    '  html+="<div class=\\"sankey-column\\">";' +
    '  if(data.statusCounts.open>0)html+="<div class=\\"sankey-node status-open\\">Open<br/>"+data.statusCounts.open+"</div>";' +
    '  if(data.statusCounts.pending>0)html+="<div class=\\"sankey-node status-pending\\">Pending<br/>"+data.statusCounts.pending+"</div>";' +
    '  if(data.statusCounts.closed>0)html+="<div class=\\"sankey-node status-closed\\">Closed<br/>"+data.statusCounts.closed+"</div>";' +
    '  html+="<div class=\\"sankey-label\\">Current Status</div></div>";' +
    '  html+="<div class=\\"sankey-column\\">";' +
    '  var totalResolved=data.resolutions.won+data.resolutions.settled+data.resolutions.withdrawn+data.resolutions.denied;' +
    '  if(data.resolutions.won>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#059669,#10b981)\\">Won "+data.resolutions.won+"</div>";' +
    '  if(data.resolutions.settled>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#f97316,#fb923c)\\">Settled "+data.resolutions.settled+"</div>";' +
    '  if(data.resolutions.withdrawn>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#6b7280,#9ca3af)\\">Withdrawn "+data.resolutions.withdrawn+"</div>";' +
    '  if(data.resolutions.denied>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#dc2626,#ef4444)\\">Denied "+data.resolutions.denied+"</div>";' +
    '  if(totalResolved===0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:#ccc\\">Pending</div>";' +
    '  html+="<div class=\\"sankey-label\\">Outcome</div></div>";' +
    '  html+="</div></div></div>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Render resources/links tab
    'function renderResources(data){' +
    '  var c=document.getElementById("resources-content");' +
    '  var html="";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📝 Forms & Submissions</div><div class=\\"link-grid\\">";' +
    '  if(data.grievanceForm)html+="<a href=\\""+data.grievanceForm+"\\" target=\\"_blank\\" class=\\"resource-link\\">📋 Grievance Form</a>";' +
    '  if(data.contactForm)html+="<a href=\\""+data.contactForm+"\\" target=\\"_blank\\" class=\\"resource-link\\">✉️ Contact Form</a>";' +
    '  if(data.satisfactionForm)html+="<a href=\\""+data.satisfactionForm+"\\" target=\\"_blank\\" class=\\"resource-link\\">📊 Satisfaction Survey</a>";' +
    '  if(!data.grievanceForm&&!data.contactForm&&!data.satisfactionForm)html+="<div class=\\"empty-state\\">No forms configured. Add URLs in Config sheet.</div>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📂 Data & Documents</div><div class=\\"link-grid\\">";' +
    '  html+="<a href=\\""+data.spreadsheetUrl+"\\" target=\\"_blank\\" class=\\"resource-link\\">📊 Open Full Spreadsheet</a>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMemberDirectory()\\">👥 Member Directory</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showGrievanceLog()\\">📋 Grievance Log</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showConfigSheet()\\">⚙️ Configuration</button>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🌐 External Links</div><div class=\\"link-grid\\">";' +
    '  if(data.orgWebsite)html+="<a href=\\""+data.orgWebsite+"\\" target=\\"_blank\\" class=\\"resource-link\\">🏛️ Organization Website</a>";' +
    '  html+="<a href=\\"https://github.com/Woop91/509-dashboard-second\\" target=\\"_blank\\" class=\\"resource-link\\">📦 GitHub Repository</a>";' +
    '  html+="</div></div>";' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">⚡ Quick Actions</div><div class=\\"link-grid\\">";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMobileUnifiedSearch()\\">🔍 Search All</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMobileGrievanceForm()\\">➕ New Grievance</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMyAssignedGrievances()\\">👤 My Cases</button>";' +
    '  html+="<button class=\\"resource-link\\" onclick=\\"google.script.run.showMemberSatisfactionDashboard()\\">📈 Satisfaction Dashboard</button>";' +
    '  html+="</div></div>";' +
    '  c.innerHTML=html;' +
    '}' +

    // Initialize
    'loadOverview();' +
    '</script>' +

    '</body></html>';
}

/**
 * Get overview data for interactive dashboard - ENHANCED with location data
 */
function getInteractiveOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {
    totalMembers: 0,
    activeStewards: 0,
    totalGrievances: 0,
    openGrievances: 0,
    pendingInfo: 0,
    winRate: '0%',
    byLocation: []  // NEW: Location breakdown for overview
  };

  var locationMap = {};

  // Get member stats - only count rows with valid member IDs (starting with M)
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
    memberData.forEach(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return;

      data.totalMembers++;
      if (row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes') data.activeStewards++;

      // Count by location
      var location = row[MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      if (!locationMap[location]) locationMap[location] = 0;
      locationMap[location]++;
    });

    // Convert location map to sorted array (top 8)
    data.byLocation = Object.keys(locationMap).map(function(key) {
      return { name: key, count: locationMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 8);
  }

  // Get grievance stats - only count rows with valid grievance IDs (starting with G)
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
    var wonCount = 0;
    var closedCount = 0;
    grievanceData.forEach(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return;

      data.totalGrievances++;
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var resolution = row[GRIEVANCE_COLS.RESOLUTION - 1] || '';
      if (status === 'Open') data.openGrievances++;
      if (status === 'Pending Info') data.pendingInfo++;
      if (status !== 'Open' && status !== 'Pending Info') closedCount++;
      if (resolution.toLowerCase().indexOf('won') >= 0 || resolution.toLowerCase().indexOf('favorable') >= 0) wonCount++;
    });
    if (closedCount > 0) {
      data.winRate = Math.round(wonCount / closedCount * 100) + '%';
    }
  }

  return data;
}

/**
 * Get member data for interactive dashboard (expanded with more details)
 */
function getInteractiveMemberData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();
  return data.map(function(row) {
    var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
    // Skip blank rows - must have a valid member ID starting with M
    if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return null;

    return {
      id: memberId,
      firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
      lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
      name: ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim(),
      title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
      location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
      unit: row[MEMBER_COLS.UNIT - 1] || 'N/A',
      officeDays: row[MEMBER_COLS.OFFICE_DAYS - 1] || 'N/A',
      email: row[MEMBER_COLS.EMAIL - 1] || '',
      phone: row[MEMBER_COLS.PHONE - 1] || '',
      preferredComm: row[MEMBER_COLS.PREFERRED_COMM - 1] || 'N/A',
      supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
      isSteward: row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
      assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1] || 'N/A',
      hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] === 'Yes',
      grievanceStatus: row[MEMBER_COLS.GRIEVANCE_STATUS - 1] || ''
    };
  }).filter(function(m) { return m !== null; });
}

/**
 * Get grievance data for interactive dashboard (expanded with more details)
 */
function getInteractiveGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
  var tz = Session.getScriptTimeZone();

  return data.map(function(row, idx) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
    // Skip blank rows - must have a valid grievance ID starting with G
    if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    return {
      id: grievanceId,
      rowNum: idx + 2, // For navigation back to sheet
      memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
      memberName: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
      status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
      currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
      articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
      incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
      nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
      daysToDeadline: daysToDeadline,
      isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
      daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
      location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
      unit: row[GRIEVANCE_COLS.UNIT - 1] || 'N/A',
      steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
      resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
    };
  }).filter(function(g) { return g !== null; });
}

/**
 * Get steward's assigned grievances for My Cases tab
 * Returns grievances where current user is the assigned steward
 */
function getMyStewardCases() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
  var tz = Session.getScriptTimeZone();

  // Also check Member Directory to get steward name for matching
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var userStewardName = '';
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
    for (var i = 0; i < memberData.length; i++) {
      var memberEmail = memberData[i][MEMBER_COLS.EMAIL - 1] || '';
      if (memberEmail.toLowerCase() === email.toLowerCase() && memberData[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
        userStewardName = ((memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '')).trim();
        break;
      }
    }
  }

  return data.map(function(row, idx) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
    // Skip blank rows
    if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

    // Check if current user is the steward for this grievance
    var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
    var isMyCase = false;

    // Match by email
    if (steward && steward.toLowerCase().indexOf(email.toLowerCase()) >= 0) {
      isMyCase = true;
    }
    // Match by name if we found the user's steward name
    if (!isMyCase && userStewardName && steward && steward.toLowerCase().indexOf(userStewardName.toLowerCase()) >= 0) {
      isMyCase = true;
    }

    if (!isMyCase) return null;

    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    return {
      id: grievanceId,
      rowNum: idx + 2,
      memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
      memberName: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
      status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
      currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
      articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
      nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
      daysToDeadline: daysToDeadline,
      isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
      daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
      location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A'
    };
  }).filter(function(g) { return g !== null; });
}

/**
 * Get analytics data for interactive dashboard - ENHANCED VERSION
 * Now includes: grievance stats, steward performance, survey results, trends
 */
function getInteractiveAnalyticsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var data = {
    memberStats: {
      total: 0,
      stewards: 0,
      withOpenGrievance: 0,
      stewardRatio: '0:0',
      byLocation: [],
      byUnit: []
    },
    statusCounts: { open: 0, pending: 0, closed: 0 },
    topCategories: [],
    resolutions: { won: 0, settled: 0, withdrawn: 0, denied: 0 },
    // NEW: Enhanced grievance stats
    grievanceStats: {
      avgDaysToResolve: 0,
      totalResolved: 0,
      byType: [],
      byStatus: [],
      byLocation: [],
      monthlyTrends: []
    },
    // NEW: Top performers and busiest stewards
    stewardPerformance: {
      topPerformers: [],
      busiestStewards: []
    },
    // NEW: Survey results
    surveyResults: {
      totalResponses: 0,
      avgSatisfaction: 0,
      responseRate: 0,
      bySection: []
    }
  };

  // Get Member Directory statistics
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var locationMap = {};
  var unitMap = {};

  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.HAS_OPEN_GRIEVANCE).getValues();

    memberData.forEach(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return;

      data.memberStats.total++;
      if (row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes') data.memberStats.stewards++;
      if (row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] === 'Yes') data.memberStats.withOpenGrievance++;

      var location = row[MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      if (!locationMap[location]) locationMap[location] = { members: 0, grievances: 0, open: 0 };
      locationMap[location].members++;

      var unit = row[MEMBER_COLS.UNIT - 1] || 'Unknown';
      if (!unitMap[unit]) unitMap[unit] = 0;
      unitMap[unit]++;
    });

    if (data.memberStats.stewards > 0) {
      var ratio = Math.round(data.memberStats.total / data.memberStats.stewards);
      data.memberStats.stewardRatio = ratio + ':1';
    } else {
      data.memberStats.stewardRatio = 'N/A';
    }

    data.memberStats.byLocation = Object.keys(locationMap).map(function(key) {
      return { name: key, count: locationMap[key].members };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

    data.memberStats.byUnit = Object.keys(unitMap).map(function(key) {
      return { name: key, count: unitMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  }

  // Get Grievance Log statistics - ENHANCED
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var rows = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
    var categoryMap = {};
    var statusMap = {};
    var stewardCaseCount = {};
    var grievanceLocationMap = {};
    var monthlyMap = {};
    var totalDaysToResolve = 0;
    var resolvedCount = 0;

    rows.forEach(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return;

      var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
      var category = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';
      var resolution = (row[GRIEVANCE_COLS.RESOLUTION - 1] || '').toLowerCase();
      var steward = row[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned';
      var location = row[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
      var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];
      var daysOpen = row[GRIEVANCE_COLS.DAYS_OPEN - 1];

      // Status counts
      if (status === 'Open') {
        data.statusCounts.open++;
        if (!statusMap['Open']) statusMap['Open'] = { count: 0, color: '#DC2626' };
        statusMap['Open'].count++;
      } else if (status === 'Pending Info') {
        data.statusCounts.pending++;
        if (!statusMap['Pending Info']) statusMap['Pending Info'] = { count: 0, color: '#F97316' };
        statusMap['Pending Info'].count++;
      } else if (status === 'Resolved' || status === 'Closed' || status === 'Withdrawn') {
        data.statusCounts.closed++;
        if (!statusMap[status]) statusMap[status] = { count: 0, color: '#059669' };
        statusMap[status].count++;

        // Calculate days to resolve
        if (dateFiled instanceof Date && dateClosed instanceof Date) {
          var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
          if (days >= 0) {
            totalDaysToResolve += days;
            resolvedCount++;
          }
        } else if (typeof daysOpen === 'number' && daysOpen > 0) {
          totalDaysToResolve += daysOpen;
          resolvedCount++;
        }
      } else if (status) {
        if (!statusMap[status]) statusMap[status] = { count: 0, color: '#6B7280' };
        statusMap[status].count++;
      }

      // Category counts
      if (!categoryMap[category]) categoryMap[category] = 0;
      categoryMap[category]++;

      // Steward case counts
      if (steward && steward !== 'Unassigned') {
        if (!stewardCaseCount[steward]) stewardCaseCount[steward] = { total: 0, open: 0 };
        stewardCaseCount[steward].total++;
        if (status === 'Open' || status === 'Pending Info') stewardCaseCount[steward].open++;
      }

      // Location grievance counts
      if (!grievanceLocationMap[location]) grievanceLocationMap[location] = { total: 0, open: 0 };
      grievanceLocationMap[location].total++;
      if (status === 'Open' || status === 'Pending Info') grievanceLocationMap[location].open++;

      // Monthly trends
      if (dateFiled instanceof Date) {
        var monthKey = Utilities.formatDate(dateFiled, Session.getScriptTimeZone(), 'yyyy-MM');
        if (!monthlyMap[monthKey]) monthlyMap[monthKey] = { filed: 0, resolved: 0 };
        monthlyMap[monthKey].filed++;
      }
      if (dateClosed instanceof Date) {
        var closeMonthKey = Utilities.formatDate(dateClosed, Session.getScriptTimeZone(), 'yyyy-MM');
        if (!monthlyMap[closeMonthKey]) monthlyMap[closeMonthKey] = { filed: 0, resolved: 0 };
        monthlyMap[closeMonthKey].resolved++;
      }

      // Resolution counts
      if (resolution.indexOf('won') >= 0 || resolution.indexOf('favorable') >= 0) data.resolutions.won++;
      else if (resolution.indexOf('settled') >= 0) data.resolutions.settled++;
      else if (resolution.indexOf('withdrawn') >= 0) data.resolutions.withdrawn++;
      else if (resolution.indexOf('denied') >= 0 || resolution.indexOf('lost') >= 0) data.resolutions.denied++;
    });

    // Average days to resolve
    data.grievanceStats.avgDaysToResolve = resolvedCount > 0 ? Math.round(totalDaysToResolve / resolvedCount) : 0;
    data.grievanceStats.totalResolved = resolvedCount;

    // Top categories (issue types) - ALL of them, properly formatted
    data.topCategories = Object.keys(categoryMap).map(function(key) {
      return { name: key, count: categoryMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 10);

    // By type (same as categories but for the new section)
    data.grievanceStats.byType = data.topCategories;

    // By status with colors
    data.grievanceStats.byStatus = Object.keys(statusMap).map(function(key) {
      return { name: key, count: statusMap[key].count, color: statusMap[key].color };
    }).sort(function(a, b) { return b.count - a.count; });

    // Location breakdown for grievances
    data.grievanceStats.byLocation = Object.keys(grievanceLocationMap).map(function(key) {
      return { name: key, total: grievanceLocationMap[key].total, open: grievanceLocationMap[key].open };
    }).sort(function(a, b) { return b.total - a.total; }).slice(0, 10);

    // Monthly trends (last 6 months)
    var sortedMonths = Object.keys(monthlyMap).sort().slice(-6);
    data.grievanceStats.monthlyTrends = sortedMonths.map(function(key) {
      return { month: key, filed: monthlyMap[key].filed, resolved: monthlyMap[key].resolved };
    });

    // Top 10 busiest stewards
    data.stewardPerformance.busiestStewards = Object.keys(stewardCaseCount).map(function(key) {
      return { name: key, total: stewardCaseCount[key].total, open: stewardCaseCount[key].open };
    }).sort(function(a, b) { return b.total - a.total; }).slice(0, 10);
  }

  // Get Steward Performance data from hidden sheet
  var perfSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
  if (perfSheet && perfSheet.getLastRow() > 1) {
    try {
      var perfData = perfSheet.getRange(2, 1, Math.min(perfSheet.getLastRow() - 1, 20), 10).getValues();
      data.stewardPerformance.topPerformers = perfData
        .filter(function(row) { return row[0] && row[9]; }) // Has name and score
        .map(function(row) {
          return {
            name: row[0],
            totalCases: row[1] || 0,
            active: row[2] || 0,
            closed: row[3] || 0,
            won: row[4] || 0,
            winRate: row[5] || 0,
            avgDays: row[6] || 0,
            score: row[9] || 0
          };
        })
        .sort(function(a, b) { return b.score - a.score; })
        .slice(0, 10);
    } catch (e) {
      Logger.log('Error reading steward performance: ' + e.message);
    }
  }

  // Get Survey Results from Member Satisfaction sheet
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet && satSheet.getLastRow() > 1) {
    try {
      var satLastRow = satSheet.getLastRow();
      var numResponses = satLastRow - 1;
      data.surveyResults.totalResponses = numResponses;

      // Calculate response rate (responses / total members)
      if (data.memberStats.total > 0) {
        data.surveyResults.responseRate = Math.round((numResponses / data.memberStats.total) * 100);
      }

      if (numResponses > 0) {
        // Get satisfaction scores (Q6-Q9: columns G-J, indices 6-9)
        var satScores = satSheet.getRange(2, SATISFACTION_COLS.Q6_SATISFIED_REP, numResponses, 4).getValues();
        var totalSat = 0;
        var validSatCount = 0;
        satScores.forEach(function(row) {
          for (var i = 0; i < 4; i++) {
            var val = parseFloat(row[i]);
            if (!isNaN(val) && val > 0) {
              totalSat += val;
              validSatCount++;
            }
          }
        });
        data.surveyResults.avgSatisfaction = validSatCount > 0 ? (totalSat / validSatCount).toFixed(1) : 0;

        // Get section averages
        var sections = [
          { name: 'Overall Satisfaction', startCol: SATISFACTION_COLS.Q6_SATISFIED_REP, numCols: 4 },
          { name: 'Steward Ratings', startCol: SATISFACTION_COLS.Q10_TIMELY_RESPONSE, numCols: 7 },
          { name: 'Chapter Effectiveness', startCol: SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, numCols: 5 },
          { name: 'Local Leadership', startCol: SATISFACTION_COLS.Q26_DECISIONS_CLEAR, numCols: 6 },
          { name: 'Contract Enforcement', startCol: SATISFACTION_COLS.Q32_ENFORCES_CONTRACT, numCols: 4 },
          { name: 'Communication', startCol: SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, numCols: 5 },
          { name: 'Member Voice', startCol: SATISFACTION_COLS.Q46_VOICE_MATTERS, numCols: 5 },
          { name: 'Value & Action', startCol: SATISFACTION_COLS.Q51_GOOD_VALUE, numCols: 5 }
        ];

        sections.forEach(function(section) {
          try {
            var sectionData = satSheet.getRange(2, section.startCol, numResponses, section.numCols).getValues();
            var sectionTotal = 0;
            var sectionValidCount = 0;
            sectionData.forEach(function(row) {
              for (var i = 0; i < section.numCols; i++) {
                var val = parseFloat(row[i]);
                if (!isNaN(val) && val > 0) {
                  sectionTotal += val;
                  sectionValidCount++;
                }
              }
            });
            var sectionAvg = sectionValidCount > 0 ? (sectionTotal / sectionValidCount).toFixed(1) : 0;
            data.surveyResults.bySection.push({ name: section.name, avg: parseFloat(sectionAvg) });
          } catch (e) {
            data.surveyResults.bySection.push({ name: section.name, avg: 0 });
          }
        });
      }
    } catch (e) {
      Logger.log('Error reading satisfaction data: ' + e.message);
    }
  }

  return data;
}

/**
 * Get resource links from Config sheet for the dashboard
 */
function getInteractiveResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: ''
  };

  if (configSheet && configSheet.getLastRow() > 1) {
    try {
      // Get URLs from Config sheet row 2
      var row = configSheet.getRange(2, 1, 1, CONFIG_COLS.SATISFACTION_FORM_URL).getValues()[0];
      links.grievanceForm = row[CONFIG_COLS.GRIEVANCE_FORM_URL - 1] || '';
      links.contactForm = row[CONFIG_COLS.CONTACT_FORM_URL - 1] || '';
      links.satisfactionForm = row[CONFIG_COLS.SATISFACTION_FORM_URL - 1] || '';
      links.orgWebsite = row[CONFIG_COLS.ORG_WEBSITE - 1] || '';
    } catch (e) {
      Logger.log('Error getting resource links: ' + e.message);
    }
  }

  return links;
}

/**
 * Get unique filter options for members (locations, units, office days)
 */
function getInteractiveMemberFilters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var filters = {
    locations: [],
    units: [],
    officeDays: []
  };

  if (configSheet && configSheet.getLastRow() > 1) {
    try {
      var lastRow = configSheet.getLastRow();
      var data = configSheet.getRange(2, 1, lastRow - 1, CONFIG_COLS.OFFICE_DAYS).getValues();

      // Get unique values from config
      data.forEach(function(row) {
        var loc = row[CONFIG_COLS.OFFICE_LOCATIONS - 1];
        var unit = row[CONFIG_COLS.UNITS - 1];
        var days = row[CONFIG_COLS.OFFICE_DAYS - 1];

        if (loc && filters.locations.indexOf(loc) === -1) filters.locations.push(loc);
        if (unit && filters.units.indexOf(unit) === -1) filters.units.push(unit);
        if (days && filters.officeDays.indexOf(days) === -1) filters.officeDays.push(days);
      });
    } catch (e) {
      Logger.log('Error getting filter options: ' + e.message);
    }
  }

  return filters;
}

/**
 * Navigate to a specific member in the Member Directory sheet
 */
function navigateToMemberInSheet(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;

  // Find the member row
  var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === memberId) {
      sheet.activate();
      var row = i + 2; // Row 1 is header
      sheet.setActiveRange(sheet.getRange(row, 1));
      ss.toast('Navigated to ' + memberId, 'Member Found', 3);
      return;
    }
  }
  ss.toast('Member not found: ' + memberId, 'Not Found', 3);
}

/**
 * Navigate to a specific grievance in the Grievance Log sheet
 */
function navigateToGrievanceInSheet(grievanceId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return;

  // Find the grievance row
  var data = sheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      sheet.activate();
      var row = i + 2; // Row 1 is header
      sheet.setActiveRange(sheet.getRange(row, 1));
      ss.toast('Navigated to ' + grievanceId, 'Grievance Found', 3);
      return;
    }
  }
  ss.toast('Grievance not found: ' + grievanceId, 'Not Found', 3);
}

/**
 * Show the Member Directory sheet
 */
function showMemberDirectory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (sheet) {
    sheet.activate();
  }
}

/**
 * Show the Grievance Log sheet
 */
function showGrievanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    sheet.activate();
  }
}

/**
 * Show the Config sheet
 */
function showConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (sheet) {
    sheet.activate();
  }
}

/**
 * Save a member from the interactive dashboard (add or edit)
 * @param {Object} memberData - Member data from the form
 * @param {string} mode - 'add' or 'edit'
 * @returns {Object} Result with success status
 */
function saveInteractiveMember(memberData, mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) throw new Error('Member Directory sheet not found');

  if (mode === 'add') {
    // Generate a new member ID
    var existingIds = {};
    var idData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    idData.forEach(function(row) {
      if (row[0]) existingIds[row[0]] = true;
    });

    var newId = generateNameBasedId('M', memberData.firstName, memberData.lastName, existingIds);

    // Create new row array
    var newRow = [];
    for (var i = 0; i < MEMBER_COLS.QUICK_ACTIONS; i++) newRow.push('');

    newRow[MEMBER_COLS.MEMBER_ID - 1] = newId;
    newRow[MEMBER_COLS.FIRST_NAME - 1] = memberData.firstName;
    newRow[MEMBER_COLS.LAST_NAME - 1] = memberData.lastName;
    newRow[MEMBER_COLS.JOB_TITLE - 1] = memberData.jobTitle || '';
    newRow[MEMBER_COLS.WORK_LOCATION - 1] = memberData.location || '';
    newRow[MEMBER_COLS.UNIT - 1] = memberData.unit || '';
    newRow[MEMBER_COLS.OFFICE_DAYS - 1] = memberData.officeDays || '';
    newRow[MEMBER_COLS.EMAIL - 1] = memberData.email || '';
    newRow[MEMBER_COLS.PHONE - 1] = memberData.phone || '';
    newRow[MEMBER_COLS.SUPERVISOR - 1] = memberData.supervisor || '';
    newRow[MEMBER_COLS.IS_STEWARD - 1] = memberData.isSteward || 'No';

    // Append the new row
    sheet.appendRow(newRow);
    ss.toast('New member added: ' + memberData.firstName + ' ' + memberData.lastName + ' (' + newId + ')', 'Member Added', 5);

    return { success: true, memberId: newId, mode: 'add' };

  } else if (mode === 'edit') {
    // Find the member row by ID
    var memberId = memberData.memberId;
    if (!memberId) throw new Error('Member ID is required for editing');

    var data = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1).getValues();
    var rowIndex = -1;
    for (var j = 0; j < data.length; j++) {
      if (data[j][0] === memberId) {
        rowIndex = j + 2; // Row 1 is header
        break;
      }
    }

    if (rowIndex === -1) throw new Error('Member not found: ' + memberId);

    // Update the member data
    sheet.getRange(rowIndex, MEMBER_COLS.FIRST_NAME).setValue(memberData.firstName);
    sheet.getRange(rowIndex, MEMBER_COLS.LAST_NAME).setValue(memberData.lastName);
    sheet.getRange(rowIndex, MEMBER_COLS.JOB_TITLE).setValue(memberData.jobTitle || '');
    sheet.getRange(rowIndex, MEMBER_COLS.WORK_LOCATION).setValue(memberData.location || '');
    sheet.getRange(rowIndex, MEMBER_COLS.UNIT).setValue(memberData.unit || '');
    sheet.getRange(rowIndex, MEMBER_COLS.OFFICE_DAYS).setValue(memberData.officeDays || '');
    sheet.getRange(rowIndex, MEMBER_COLS.EMAIL).setValue(memberData.email || '');
    sheet.getRange(rowIndex, MEMBER_COLS.PHONE).setValue(memberData.phone || '');
    sheet.getRange(rowIndex, MEMBER_COLS.SUPERVISOR).setValue(memberData.supervisor || '');
    sheet.getRange(rowIndex, MEMBER_COLS.IS_STEWARD).setValue(memberData.isSteward || 'No');

    ss.toast('Member updated: ' + memberData.firstName + ' ' + memberData.lastName, 'Member Updated', 5);

    return { success: true, memberId: memberId, mode: 'edit' };
  }

  throw new Error('Invalid mode: ' + mode);
}

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║         ⚠️  END OF PROTECTED SECTION - INTERACTIVE DASHBOARD  ⚠️         ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝
/**
 * ============================================================================
 * 509 DASHBOARD - STRATEGIC COMMAND CENTER MASTER ENGINE (V 3.6.0)
 * ============================================================================
 * CORE FEATURES:
 * 1. Dual-Dashboard Architecture:
 *    - Executive View (Internal, shows PII, Case Management)
 *    - Member Analytics (Expansive, No PII, Strategic Reporting)
 * 2. Visual KPIs: Success Gauges, Sentiment Radars, Participation Funnels.
 * 3. Strategic Intelligence: Hot Zone Heatmaps & Rising Star Identification.
 * 4. Automation: Midnight Auto-Refresh & Critical Performance Email Alerts.
 * 5. Peak Performance: Array-based batch processing for speed and reliability.
 * 6. Auto-ID Generator & Duplicate Prevention
 * 7. Legal Stage-Gate Workflow (Intake to Arbitration)
 * 8. Automatic Chief Steward Escalation Alerts
 * 9. Digital Signature Block PDF Generation
 * 10. Automated Global UI Styling (Roboto Theme & Status Colors)
 * ============================================================================
 *
 * OFFICIAL BRANDING PATHS:
 * - All automated emails use: "[509 Strategic Command Center] Status Update"
 * - All PDF metadata lists "509 Strategic Command Center" as the Author
 * ============================================================================
 */

// ============================================================================
// 1. NAVIGATION HELPERS
// ============================================================================

// navigateToSheet() - REMOVED DUPLICATE - see line 565 for main definition

// ============================================================================
// 2. EXECUTIVE COMMAND MODAL (SPA Architecture - Bridge Pattern)
// ============================================================================

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Executive metrics are now in the unified Steward Dashboard.
 */
function rebuildExecutiveDashboard() {
  showStewardDashboard();
}

/**
 * Launches the Executive Command Center Modal with professional UI
 */
function launchExecutiveDashboard() {
  var html = HtmlService.createHtmlOutput(getExecutiveDashboardHtml_())
    .setWidth(1000)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, '509 STRATEGIC COMMAND CENTER');
}

/**
 * High-Performance KPI Aggregator for the Modal (Bridge Pattern)
 * Returns JSON data for client-side rendering
 * @returns {string} JSON string with dashboard statistics
 */
function getDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var stats = {
    totalGrievances: 0,
    activeGrievances: 0,
    activeSteps: { step1: 0, step2: 0, arbitration: 0 },
    outcomes: { wins: 0, losses: 0, settled: 0, withdrawn: 0 },
    winRate: 0,
    overdueCount: 0,
    totalMembers: 0,
    stewardCount: 0,
    moraleScore: 5, // Placeholder for SurveyMonkey integration
    unitBreakdown: {},
    stewardWorkload: []
  };

  // Process Grievance Log
  if (logSheet && logSheet.getLastRow() > 1) {
    var logData = logSheet.getDataRange().getValues();
    logData.shift(); // Remove headers

    stats.totalGrievances = logData.length;

    logData.forEach(function(row) {
      var status = (row[GRIEVANCE_COLS.STATUS - 1] || '').toString().toLowerCase();
      var currentStep = (row[GRIEVANCE_COLS.CURRENT_STEP - 1] || '').toString().toLowerCase();
      var unit = row[GRIEVANCE_COLS.UNIT - 1] || 'Unknown';
      var steward = row[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned';

      // Count active vs closed
      if (status === 'open' || status === 'pending info' || status === 'appealed') {
        stats.activeGrievances++;
      }

      // Count by step
      if (currentStep.indexOf('step 1') !== -1 || currentStep === '1') stats.activeSteps.step1++;
      if (currentStep.indexOf('step 2') !== -1 || currentStep === '2') stats.activeSteps.step2++;
      if (currentStep.indexOf('arbitration') !== -1 || currentStep.indexOf('step 3') !== -1) stats.activeSteps.arbitration++;

      // Count outcomes
      if (status === 'won' || status === 'sustained') stats.outcomes.wins++;
      if (status === 'denied' || status === 'lost') stats.outcomes.losses++;
      if (status === 'settled') stats.outcomes.settled++;
      if (status === 'withdrawn') stats.outcomes.withdrawn++;

      // Unit breakdown
      if (!stats.unitBreakdown[unit]) stats.unitBreakdown[unit] = 0;
      stats.unitBreakdown[unit]++;

      // Steward workload (only active cases)
      if (status === 'open' || status === 'pending info') {
        var existingSteward = stats.stewardWorkload.find(function(s) { return s.name === steward; });
        if (existingSteward) {
          existingSteward.count++;
        } else {
          stats.stewardWorkload.push({ name: steward, count: 1 });
        }
      }

      // Check for overdue
      var step1Due = row[GRIEVANCE_COLS.STEP1_DUE - 1];
      if (step1Due && new Date(step1Due) < new Date() && (status === 'open' || status === 'pending info')) {
        stats.overdueCount++;
      }
    });

    // Calculate win rate
    var totalClosed = stats.outcomes.wins + stats.outcomes.losses + stats.outcomes.settled;
    if (totalClosed > 0) {
      stats.winRate = Math.round((stats.outcomes.wins / totalClosed) * 100);
    }

    // Sort steward workload
    stats.stewardWorkload.sort(function(a, b) { return b.count - a.count; });
  }

  // Process Member Directory
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var m = 1; m < memberData.length; m++) {
      if (memberData[m][MEMBER_COLS.MEMBER_ID - 1]) {
        stats.totalMembers++;
        if (memberData[m][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') stats.stewardCount++;
      }
    }
  }

  // Get morale score from satisfaction data
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet && satSheet.getLastRow() > 1) {
    var satData = satSheet.getDataRange().getValues();
    var totalScore = 0;
    var scoreCount = 0;
    for (var s = 1; s < satData.length; s++) {
      var avgScore = parseFloat(satData[s][SATISFACTION_COLS.AVG_OVERALL_SAT - 1]);
      if (!isNaN(avgScore) && avgScore > 0) {
        totalScore += avgScore;
        scoreCount++;
      }
    }
    if (scoreCount > 0) {
      stats.moraleScore = Math.round((totalScore / scoreCount) * 10) / 10;
    }
  }

  return JSON.stringify(stats);
}

/**
 * Generates the Executive Dashboard HTML with Chart.js
 * @returns {string} Complete HTML for the modal
 * @private
 */
function getExecutiveDashboardHtml_() {
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <base target="_top">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap" rel="stylesheet">' +
    '  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #f8fafc; min-height: 100vh; padding: 24px; }' +
    '    .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; }' +
    '    .header h1 { font-size: 24px; font-weight: 900; letter-spacing: -0.5px; color: #60a5fa; }' +
    '    .status-badge { background: rgba(16, 185, 129, 0.2); color: #34d399; padding: 6px 12px; border-radius: 20px; font-size: 11px; font-weight: 700; text-transform: uppercase; border: 1px solid rgba(16, 185, 129, 0.3); }' +
    '    .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }' +
    '    .kpi-card { background: rgba(30, 41, 59, 0.8); backdrop-filter: blur(12px); border: 1px solid rgba(255,255,255,0.1); border-radius: 12px; padding: 20px; text-align: center; }' +
    '    .kpi-card.alert { border-color: #ef4444; box-shadow: 0 0 20px rgba(239,68,68,0.2); }' +
    '    .kpi-label { font-size: 11px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }' +
    '    .kpi-value { font-size: 36px; font-weight: 900; }' +
    '    .kpi-value.green { color: #34d399; }' +
    '    .kpi-value.red { color: #f87171; }' +
    '    .kpi-value.blue { color: #60a5fa; }' +
    '    .kpi-value.yellow { color: #fbbf24; }' +
    '    .charts-grid { display: grid; grid-template-columns: 2fr 1fr; gap: 20px; margin-bottom: 24px; }' +
    '    .chart-card { background: rgba(30, 41, 59, 0.8); backdrop-filter: blur(12px); border: 1px solid rgba(255,255,255,0.1); border-radius: 12px; padding: 20px; }' +
    '    .chart-title { font-size: 14px; font-weight: 600; color: #e2e8f0; margin-bottom: 16px; text-transform: uppercase; letter-spacing: 0.5px; }' +
    '    .steward-list { max-height: 200px; overflow-y: auto; }' +
    '    .steward-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.05); }' +
    '    .steward-name { font-size: 13px; color: #cbd5e1; }' +
    '    .steward-count { background: #3b82f6; color: white; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }' +
    '    .footer { display: flex; justify-content: space-between; align-items: center; margin-top: 20px; padding-top: 16px; border-top: 1px solid rgba(255,255,255,0.1); }' +
    '    .btn { padding: 10px 20px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s; border: none; }' +
    '    .btn-primary { background: #3b82f6; color: white; }' +
    '    .btn-primary:hover { background: #2563eb; }' +
    '    .btn-secondary { background: rgba(255,255,255,0.1); color: #cbd5e1; }' +
    '    .btn-secondary:hover { background: rgba(255,255,255,0.15); }' +
    '    .pii-warning { background: rgba(220, 38, 38, 0.1); border: 1px solid rgba(220, 38, 38, 0.3); color: #fca5a5; padding: 8px 16px; border-radius: 6px; font-size: 11px; font-weight: 600; }' +
    '    .loading { text-align: center; padding: 60px; color: #94a3b8; }' +
    '    canvas { max-height: 250px !important; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <h1>509 EXECUTIVE COMMAND CENTER</h1>' +
    '    <div class="status-badge" id="statusBadge">Loading...</div>' +
    '  </div>' +
    '  <div id="content"><div class="loading">Loading dashboard data...</div></div>' +
    '  <div class="footer">' +
    '    <div class="pii-warning">⚠️ CONFIDENTIAL: Contains Member PII</div>' +
    '    <div>' +
    '      <button class="btn btn-secondary" onclick="google.script.run.emailExecutivePDF()">📧 Email PDF</button>' +
    '      <button class="btn btn-primary" onclick="google.script.host.close()">Close</button>' +
    '    </div>' +
    '  </div>' +
    '  <script>' +
    '    window.onload = function() {' +
    '      google.script.run.withSuccessHandler(renderDashboard).withFailureHandler(showError).getDashboardStats();' +
    '    };' +
    '    function showError(err) {' +
    '      document.getElementById("content").innerHTML = "<div class=\\"loading\\">Error loading data: " + err.message + "</div>";' +
    '    }' +
    '    function renderDashboard(jsonStats) {' +
    '      var stats = JSON.parse(jsonStats);' +
    '      document.getElementById("statusBadge").innerText = "SYSTEM LIVE";' +
    '      var html = "";' +
    '      html += "<div class=\\"kpi-grid\\">";' +
    '      html += "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Grievances</div><div class=\\"kpi-value blue\\">" + stats.totalGrievances + "</div></div>";' +
    '      html += "<div class=\\"kpi-card" + (stats.activeGrievances > 10 ? " alert" : "") + "\\"><div class=\\"kpi-label\\">Active Cases</div><div class=\\"kpi-value red\\">" + stats.activeGrievances + "</div></div>";' +
    '      html += "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Win Rate</div><div class=\\"kpi-value green\\">" + stats.winRate + "%</div></div>";' +
    '      html += "<div class=\\"kpi-card" + (stats.overdueCount > 0 ? " alert" : "") + "\\"><div class=\\"kpi-label\\">Overdue Steps</div><div class=\\"kpi-value yellow\\">" + stats.overdueCount + "</div></div>";' +
    '      html += "</div>";' +
    '      html += "<div class=\\"charts-grid\\">";' +
    '      html += "<div class=\\"chart-card\\"><div class=\\"chart-title\\">Case Outcomes</div><canvas id=\\"outcomeChart\\"></canvas></div>";' +
    '      html += "<div class=\\"chart-card\\"><div class=\\"chart-title\\">Cases by Step</div><canvas id=\\"stepChart\\"></canvas></div>";' +
    '      html += "</div>";' +
    '      html += "<div class=\\"charts-grid\\">";' +
    '      html += "<div class=\\"chart-card\\"><div class=\\"chart-title\\">Steward Workload</div><div class=\\"steward-list\\" id=\\"stewardList\\"></div></div>";' +
    '      html += "<div class=\\"chart-card\\"><div class=\\"chart-title\\">Membership</div><div style=\\"text-align:center;padding:20px;\\"><div class=\\"kpi-value blue\\">" + stats.totalMembers + "</div><div class=\\"kpi-label\\">Total Members</div><br><div class=\\"kpi-value green\\">" + stats.stewardCount + "</div><div class=\\"kpi-label\\">Active Stewards</div></div></div>";' +
    '      html += "</div>";' +
    '      document.getElementById("content").innerHTML = html;' +
    '      renderCharts(stats);' +
    '      renderStewardList(stats.stewardWorkload);' +
    '    }' +
    '    function renderCharts(stats) {' +
    '      new Chart(document.getElementById("outcomeChart"), {' +
    '        type: "bar",' +
    '        data: { labels: ["Wins", "Losses", "Settled", "Withdrawn"], datasets: [{ label: "Cases", data: [stats.outcomes.wins, stats.outcomes.losses, stats.outcomes.settled, stats.outcomes.withdrawn], backgroundColor: ["#10b981", "#ef4444", "#f59e0b", "#6b7280"] }] },' +
    '        options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { color: "#94a3b8" } }, x: { ticks: { color: "#94a3b8" } } } }' +
    '      });' +
    '      new Chart(document.getElementById("stepChart"), {' +
    '        type: "doughnut",' +
    '        data: { labels: ["Step 1", "Step 2", "Arbitration"], datasets: [{ data: [stats.activeSteps.step1, stats.activeSteps.step2, stats.activeSteps.arbitration], backgroundColor: ["#3b82f6", "#f59e0b", "#ef4444"] }] },' +
    '        options: { responsive: true, plugins: { legend: { position: "bottom", labels: { color: "#cbd5e1" } } } }' +
    '      });' +
    '    }' +
    '    function renderStewardList(workload) {' +
    '      var html = "";' +
    '      if (!workload || workload.length === 0) { html = "<div style=\\"color:#94a3b8;text-align:center;padding:20px;\\">No active cases assigned</div>"; }' +
    '      else { workload.slice(0, 10).forEach(function(s) { html += "<div class=\\"steward-item\\"><span class=\\"steward-name\\">" + s.name + "</span><span class=\\"steward-count\\">" + s.count + "</span></div>"; }); }' +
    '      document.getElementById("stewardList").innerHTML = html;' +
    '    }' +
    '  </script>' +
    '</body>' +
    '</html>';
}

/**
 * Gets executive metrics from dashboard calculations
 * @returns {Object} Metrics object with activeGrievances, winRate, overdueSteps
 * @private
 */
function getExecutiveMetrics_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC) || ss.getSheetByName("_Dashboard_Calc");

  var metrics = {
    activeGrievances: 0,
    activeTrend: "Stable",
    winRate: 0,
    winRateTrend: "Stable",
    overdueSteps: 0,
    overdueTrend: "Stable"
  };

  if (calcSheet) {
    try {
      // Pull values from calculation sheet
      var data = calcSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === 'Active Grievances') metrics.activeGrievances = data[i][1] || 0;
        if (data[i][0] === 'Win Rate') metrics.winRate = Math.round((data[i][1] || 0) * 100);
        if (data[i][0] === 'Overdue') metrics.overdueSteps = data[i][1] || 0;
      }
    } catch (e) {
      Logger.log('Error getting executive metrics: ' + e.message);
    }
  }

  // Try to get from grievance sheet if calc sheet not available
  if (metrics.activeGrievances === 0) {
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      var grievanceData = grievanceSheet.getDataRange().getValues();
      var openCount = 0;
      var wonCount = 0;
      var closedCount = 0;
      var overdueCount = 0;

      for (var g = 1; g < grievanceData.length; g++) {
        var status = grievanceData[g][GRIEVANCE_COLS.STATUS - 1];
        if (status === 'Open' || status === 'Pending Info') openCount++;
        if (status === 'Won') wonCount++;
        if (status === 'Won' || status === 'Denied' || status === 'Settled' || status === 'Withdrawn') closedCount++;

        var daysToDeadline = grievanceData[g][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
        if (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0)) {
          overdueCount++;
        }
      }

      metrics.activeGrievances = openCount;
      metrics.winRate = closedCount > 0 ? Math.round((wonCount / closedCount) * 100) : 0;
      metrics.overdueSteps = overdueCount;
    }
  }

  return metrics;
}

// ============================================================================
// 3. UNIFIED STEWARD DASHBOARD (v4.3.2)
// ============================================================================
// Consolidates all analytics, charts, and reports into a single tabbed interface.
// This replaces the individual chart modals for a unified experience.
// ============================================================================

/**
 * Shows the unified Steward Dashboard with all analytics
 * Contains all charts and statistical data for stewards
 */
function showStewardDashboard() {
  var html = HtmlService.createHtmlOutput(getStewardDashboardHtml_())
    .setWidth(DIALOG_SIZES.FULLSCREEN.width)
    .setHeight(DIALOG_SIZES.FULLSCREEN.height);
  SpreadsheetApp.getUi().showModalDialog(html, '509 STEWARD COMMAND CENTER');
}

/**
 * Gets comprehensive steward analytics data (Bridge Pattern)
 * @returns {string} JSON with all dashboard data
 */
function getStewardDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var data = {
    // KPIs
    totalMembers: 0,
    stewardCount: 0,
    totalGrievances: 0,
    openGrievances: 0,
    wins: 0,
    losses: 0,
    settled: 0,
    winRate: 0,
    overdueCount: 0,
    moraleScore: 7.5,

    // Breakdowns
    unitBreakdown: {},
    locationBreakdown: {},
    stewardWorkload: [],
    hotZones: [],
    risingStars: [],

    // Bargaining Data
    step1DenialRate: 0,
    avgSettlementDays: 0,
    topViolatedArticle: 'N/A',
    articleViolations: {},

    // Chart Data
    statusDistribution: { open: 0, pending: 0, won: 0, denied: 0, settled: 0 },
    monthlyTrend: [],
    sentimentTrend: [],

    // Satisfaction Survey Data (Section Averages)
    satisfactionData: {
      responseCount: 0,
      sections: [
        { name: 'Overall Satisfaction', key: 'overall', score: 0, questions: ['Satisfied with Rep', 'Trust Union', 'Feel Protected', 'Recommend'] },
        { name: 'Steward Ratings', key: 'steward', score: 0, questions: ['Timely Response', 'Treated Respect', 'Explained Options', 'Followed Through', 'Advocated', 'Safe Concerns', 'Confidentiality'] },
        { name: 'Chapter Effectiveness', key: 'chapter', score: 0, questions: ['Understand Issues', 'Chapter Comm', 'Organizes', 'Reach Chapter', 'Fair Rep'] },
        { name: 'Local Leadership', key: 'leadership', score: 0, questions: ['Decisions Clear', 'Understand Process', 'Transparent Finance', 'Accountable', 'Fair Processes', 'Welcomes Opinions'] },
        { name: 'Contract Enforcement', key: 'contract', score: 0, questions: ['Enforces Contract', 'Realistic Timelines', 'Clear Updates', 'Frontline Priority'] },
        { name: 'Communication Quality', key: 'communication', score: 0, questions: ['Clear Actionable', 'Enough Info', 'Find Easily', 'All Shifts', 'Meetings Worth'] },
        { name: 'Member Voice', key: 'voice', score: 0, questions: ['Voice Matters', 'Seeks Input', 'Dignity', 'Newer Supported', 'Conflict Respect'] },
        { name: 'Value & Action', key: 'value', score: 0, questions: ['Good Value', 'Priorities Needs', 'Prepared Mobilize', 'Win Together'] }
      ]
    }
  };

  // Process Members
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var m = 1; m < memberData.length; m++) {
      if (memberData[m][MEMBER_COLS.MEMBER_ID - 1]) {
        data.totalMembers++;
        if (memberData[m][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') data.stewardCount++;

        var location = memberData[m][MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
        var unit = memberData[m][MEMBER_COLS.UNIT - 1] || 'Unknown';
        if (!data.locationBreakdown[location]) data.locationBreakdown[location] = 0;
        data.locationBreakdown[location]++;
        if (!data.unitBreakdown[unit]) data.unitBreakdown[unit] = 0;
        data.unitBreakdown[unit]++;
      }
    }
  }

  // Process Grievances
  var stewardCases = {};
  var locationCases = {};
  var step1Total = 0, step1Denials = 0;
  var settlementDays = [];

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    for (var g = 1; g < grievanceData.length; g++) {
      var status = (grievanceData[g][GRIEVANCE_COLS.STATUS - 1] || '').toString();
      var steward = grievanceData[g][GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned';
      var location = grievanceData[g][GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
      var article = grievanceData[g][GRIEVANCE_COLS.ARTICLES - 1];

      if (!grievanceData[g][GRIEVANCE_COLS.GRIEVANCE_ID - 1]) continue;
      data.totalGrievances++;

      // Status distribution
      switch(status.toLowerCase()) {
        case 'open': data.statusDistribution.open++; data.openGrievances++; break;
        case 'pending info': data.statusDistribution.pending++; data.openGrievances++; break;
        case 'won': case 'sustained': data.statusDistribution.won++; data.wins++; break;
        case 'denied': case 'lost': data.statusDistribution.denied++; data.losses++; break;
        case 'settled': data.statusDistribution.settled++; data.settled++; break;
      }

      // Steward workload
      if (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info') {
        if (!stewardCases[steward]) stewardCases[steward] = 0;
        stewardCases[steward]++;
      }

      // Hot zones (locations with active cases)
      if (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info') {
        if (!locationCases[location]) locationCases[location] = 0;
        locationCases[location]++;
      }

      // Article violations
      if (article) {
        if (!data.articleViolations[article]) data.articleViolations[article] = 0;
        data.articleViolations[article]++;
      }

      // Step 1 denial rate
      if (grievanceData[g][GRIEVANCE_COLS.STEP_1_DATE - 1]) {
        step1Total++;
        if (status !== 'Won' && grievanceData[g][GRIEVANCE_COLS.STEP_2_DATE - 1]) step1Denials++;
      }

      // Settlement time
      var dateFiled = grievanceData[g][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = grievanceData[g][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateFiled instanceof Date && dateClosed instanceof Date) {
        var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
        if (days > 0) settlementDays.push(days);
      }

      // Check overdue
      var step1Due = grievanceData[g][GRIEVANCE_COLS.STEP1_DUE - 1];
      if (step1Due && new Date(step1Due) < new Date() && (status.toLowerCase() === 'open' || status.toLowerCase() === 'pending info')) {
        data.overdueCount++;
      }
    }
  }

  // Calculate derived metrics
  var totalClosed = data.wins + data.losses + data.settled;
  data.winRate = totalClosed > 0 ? Math.round((data.wins / totalClosed) * 100) : 0;
  data.step1DenialRate = step1Total > 0 ? Math.round((step1Denials / step1Total) * 100) : 0;
  data.avgSettlementDays = settlementDays.length > 0 ? Math.round(settlementDays.reduce(function(a,b){return a+b;},0) / settlementDays.length) : 0;

  // Build steward workload array
  for (var s in stewardCases) {
    var count = stewardCases[s];
    var statusLabel = count > 8 ? 'OVERLOAD' : count > 5 ? 'Heavy' : 'Available';
    var color = count > 8 ? '#ef4444' : count > 5 ? '#f59e0b' : '#22c55e';
    data.stewardWorkload.push({ name: s, count: count, status: statusLabel, color: color });
  }
  data.stewardWorkload.sort(function(a,b){return b.count - a.count;});

  // Build hot zones (locations with 3+ active cases)
  for (var loc in locationCases) {
    if (locationCases[loc] >= 3) {
      data.hotZones.push({ location: loc, count: locationCases[loc] });
    }
  }
  data.hotZones.sort(function(a,b){return b.count - a.count;});

  // Top violated article
  var maxViolations = 0;
  for (var art in data.articleViolations) {
    if (data.articleViolations[art] > maxViolations) {
      maxViolations = data.articleViolations[art];
      data.topViolatedArticle = art;
    }
  }

  // Process Satisfaction Survey for morale, sentiment, and section scores
  if (satSheet && satSheet.getLastRow() > 1) {
    var satData = satSheet.getDataRange().getValues();
    var trustScores = [];
    var monthlyTrust = {};

    // Section score accumulators (using summary columns BT-CD if available, else calculate)
    var sectionScores = {
      overall: [], steward: [], chapter: [], leadership: [],
      contract: [], communication: [], voice: [], value: []
    };

    for (var i = 1; i < satData.length; i++) {
      if (!satData[i][0]) continue; // Skip empty rows
      data.satisfactionData.responseCount++;

      var trustVal = parseFloat(satData[i][7]); // Q7_TRUST_UNION (col H, index 7)
      var timestamp = satData[i][0];

      if (!isNaN(trustVal) && trustVal >= 1 && trustVal <= 10) {
        trustScores.push(trustVal);
        if (timestamp) {
          var date = new Date(timestamp);
          var monthKey = date.toLocaleString('default', { month: 'short' });
          if (!monthlyTrust[monthKey]) monthlyTrust[monthKey] = { sum: 0, count: 0 };
          monthlyTrust[monthKey].sum += trustVal;
          monthlyTrust[monthKey].count++;
        }
      }

      // Calculate section averages from summary columns (BT onwards = index 71+)
      // Or calculate from raw question columns if summaries not available
      var avgOverall = parseFloat(satData[i][71]) || 0; // BT
      var avgSteward = parseFloat(satData[i][72]) || 0; // BU
      var avgChapter = parseFloat(satData[i][74]) || 0; // BW
      var avgLeadership = parseFloat(satData[i][75]) || 0; // BX
      var avgContract = parseFloat(satData[i][76]) || 0; // BY
      var avgComm = parseFloat(satData[i][78]) || 0; // CA
      var avgVoice = parseFloat(satData[i][79]) || 0; // CB
      var avgValue = parseFloat(satData[i][80]) || 0; // CC

      // If summary columns empty, calculate from raw questions
      if (avgOverall === 0) {
        var q6 = parseFloat(satData[i][6]) || 0, q7 = parseFloat(satData[i][7]) || 0;
        var q8 = parseFloat(satData[i][8]) || 0, q9 = parseFloat(satData[i][9]) || 0;
        avgOverall = (q6 + q7 + q8 + q9) / 4;
      }

      if (avgOverall > 0) sectionScores.overall.push(avgOverall);
      if (avgSteward > 0) sectionScores.steward.push(avgSteward);
      if (avgChapter > 0) sectionScores.chapter.push(avgChapter);
      if (avgLeadership > 0) sectionScores.leadership.push(avgLeadership);
      if (avgContract > 0) sectionScores.contract.push(avgContract);
      if (avgComm > 0) sectionScores.communication.push(avgComm);
      if (avgVoice > 0) sectionScores.voice.push(avgVoice);
      if (avgValue > 0) sectionScores.value.push(avgValue);
    }

    // Calculate final section averages
    function avg(arr) { return arr.length > 0 ? Math.round((arr.reduce(function(a,b){return a+b;},0) / arr.length) * 10) / 10 : 0; }
    data.satisfactionData.sections[0].score = avg(sectionScores.overall);
    data.satisfactionData.sections[1].score = avg(sectionScores.steward);
    data.satisfactionData.sections[2].score = avg(sectionScores.chapter);
    data.satisfactionData.sections[3].score = avg(sectionScores.leadership);
    data.satisfactionData.sections[4].score = avg(sectionScores.contract);
    data.satisfactionData.sections[5].score = avg(sectionScores.communication);
    data.satisfactionData.sections[6].score = avg(sectionScores.voice);
    data.satisfactionData.sections[7].score = avg(sectionScores.value);

    if (trustScores.length > 0) {
      data.moraleScore = Math.round((trustScores.reduce(function(a,b){return a+b;},0) / trustScores.length) * 10) / 10;
    }

    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    months.forEach(function(month) {
      if (monthlyTrust[month]) {
        data.sentimentTrend.push({ month: month, score: Math.round((monthlyTrust[month].sum / monthlyTrust[month].count) * 10) / 10 });
      }
    });
  }

  return JSON.stringify(data);
}

/**
 * Generates the Steward Dashboard HTML with tabbed interface
 * @returns {string} Complete HTML for the modal
 * @private
 */
function getStewardDashboardHtml_() {
  return '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #f8fafc; min-height: 100vh; }' +
    '.header { display: flex; justify-content: space-between; align-items: center; padding: 16px 24px; border-bottom: 1px solid rgba(255,255,255,0.1); }' +
    '.header h1 { font-size: 20px; font-weight: 700; color: #60a5fa; display: flex; align-items: center; gap: 10px; }' +
    '.header .material-icons { font-size: 28px; }' +
    '.pii-badge { background: rgba(239,68,68,0.2); color: #fca5a5; padding: 6px 12px; border-radius: 20px; font-size: 10px; font-weight: 700; text-transform: uppercase; }' +
    '.tabs { display: flex; gap: 4px; padding: 0 24px; background: rgba(0,0,0,0.2); }' +
    '.tab { padding: 12px 20px; cursor: pointer; font-size: 12px; font-weight: 500; color: #94a3b8; border-bottom: 2px solid transparent; transition: all 0.2s; }' +
    '.tab:hover { color: #e2e8f0; }' +
    '.tab.active { color: #60a5fa; border-bottom-color: #60a5fa; }' +
    '.content { padding: 20px 24px; overflow-y: auto; max-height: calc(100vh - 140px); }' +
    '.tab-content { display: none; }' +
    '.tab-content.active { display: block; }' +
    '.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }' +
    '.kpi-card { background: rgba(30,41,59,0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; padding: 16px; text-align: center; }' +
    '.kpi-card.alert { border-color: #ef4444; }' +
    '.kpi-label { font-size: 10px; color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }' +
    '.kpi-value { font-size: 28px; font-weight: 900; }' +
    '.kpi-value.green { color: #34d399; } .kpi-value.red { color: #f87171; } .kpi-value.blue { color: #60a5fa; } .kpi-value.yellow { color: #fbbf24; } .kpi-value.purple { color: #a78bfa; }' +
    '.charts-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }' +
    '.chart-card { background: rgba(30,41,59,0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; padding: 16px; }' +
    '.chart-title { font-size: 12px; font-weight: 600; color: #e2e8f0; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 8px; }' +
    '.chart-title .material-icons { font-size: 18px; color: #60a5fa; }' +
    'canvas { max-height: 200px !important; }' +
    '.list-container { max-height: 200px; overflow-y: auto; }' +
    '.list-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.05); }' +
    '.list-item:last-child { border-bottom: none; }' +
    '.badge { padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; }' +
    '.bargain-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }' +
    '.bargain-card { background: rgba(251,191,36,0.1); border: 1px solid rgba(251,191,36,0.3); border-radius: 10px; padding: 16px; text-align: center; }' +
    '.bargain-label { font-size: 10px; color: #fbbf24; text-transform: uppercase; letter-spacing: 1px; }' +
    '.bargain-value { font-size: 24px; font-weight: 700; color: #fcd34d; margin-top: 6px; }' +
    '.bargain-status { font-size: 10px; color: #94a3b8; margin-top: 4px; }' +
    '.hot-zone { display: flex; justify-content: space-between; padding: 12px; background: rgba(239,68,68,0.1); border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #ef4444; }' +
    '.footer { display: flex; justify-content: space-between; padding: 12px 24px; border-top: 1px solid rgba(255,255,255,0.1); }' +
    '.btn { padding: 10px 20px; border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer; border: none; }' +
    '.btn-primary { background: #3b82f6; color: white; } .btn-primary:hover { background: #2563eb; }' +
    '.loading { text-align: center; padding: 60px; color: #94a3b8; }' +
    '.sat-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }' +
    '.sat-response-count { background: rgba(96,165,250,0.2); color: #60a5fa; padding: 8px 16px; border-radius: 20px; font-size: 12px; font-weight: 600; }' +
    '.sat-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; }' +
    '.sat-section { background: rgba(30,41,59,0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 12px; padding: 16px; }' +
    '.sat-section-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }' +
    '.sat-section-name { font-size: 13px; font-weight: 600; color: #e2e8f0; }' +
    '.sat-section-score { font-size: 20px; font-weight: 900; }' +
    '.sat-score-bar { height: 8px; background: rgba(255,255,255,0.1); border-radius: 4px; overflow: hidden; margin-bottom: 12px; }' +
    '.sat-score-fill { height: 100%; border-radius: 4px; transition: width 0.3s; }' +
    '.sat-questions { font-size: 11px; color: #94a3b8; line-height: 1.6; }' +
    '.sat-question-item { padding: 4px 0; border-bottom: 1px solid rgba(255,255,255,0.05); }' +
    '.sat-question-item:last-child { border-bottom: none; }' +
    '.sat-overall-chart { margin-top: 20px; }' +
    '</style></head><body>' +
    '<div class="header"><h1><i class="material-icons">analytics</i>STEWARD COMMAND CENTER</h1><span class="pii-badge">INTERNAL USE ONLY</span></div>' +
    '<div class="tabs">' +
    '<div class="tab active" onclick="showTab(\'overview\')">Overview</div>' +
    '<div class="tab" onclick="showTab(\'workload\')">Workload</div>' +
    '<div class="tab" onclick="showTab(\'analytics\')">Analytics</div>' +
    '<div class="tab" onclick="showTab(\'hotspots\')">Hot Spots</div>' +
    '<div class="tab" onclick="showTab(\'bargaining\')">Bargaining</div>' +
    '<div class="tab" onclick="showTab(\'satisfaction\')">Satisfaction</div>' +
    '</div>' +
    '<div class="content"><div id="main-content"><div class="loading">Loading dashboard data...</div></div></div>' +
    '<div class="footer"><span style="font-size:11px;color:#64748b">Data refreshes on open</span><button class="btn btn-primary" onclick="google.script.host.close()">Close</button></div>' +
    '<script>' +
    'var dashData = null;' +
    'window.onload = function() { google.script.run.withSuccessHandler(render).withFailureHandler(showError).getStewardDashboardData(); };' +
    'function showError(e) { document.getElementById("main-content").innerHTML = "<div class=\\"loading\\">Error: " + e.message + "</div>"; }' +
    'function showTab(tab) { document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")}); document.querySelector(".tab[onclick*=\\\""+tab+"\\\"]").classList.add("active"); renderTab(tab); }' +
    'function render(json) { dashData = JSON.parse(json); renderTab("overview"); }' +
    'function renderTab(tab) {' +
    '  var d = dashData; var html = "";' +
    '  if (tab === "overview") {' +
    '    html = "<div class=\\"kpi-grid\\">" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Members</div><div class=\\"kpi-value blue\\">" + d.totalMembers + "</div></div>" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Stewards</div><div class=\\"kpi-value purple\\">" + d.stewardCount + "</div></div>" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Open Cases</div><div class=\\"kpi-value yellow\\">" + d.openGrievances + "</div></div>" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Win Rate</div><div class=\\"kpi-value green\\">" + d.winRate + "%</div></div>" +' +
    '      "<div class=\\"kpi-card " + (d.overdueCount > 0 ? "alert" : "") + "\\"><div class=\\"kpi-label\\">Overdue</div><div class=\\"kpi-value red\\">" + d.overdueCount + "</div></div>" +' +
    '    "</div>" +' +
    '    "<div class=\\"charts-row\\">" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">pie_chart</i>Case Status Distribution</div><canvas id=\\"statusChart\\"></canvas></div>" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">trending_up</i>Morale Trend</div><canvas id=\\"trendChart\\"></canvas></div>" +' +
    '    "</div>" +' +
    '    "<div class=\\"charts-row\\">" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">location_on</i>Cases by Location</div><canvas id=\\"locationChart\\"></canvas></div>" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">gavel</i>Article Violations</div><canvas id=\\"articleChart\\"></canvas></div>" +' +
    '    "</div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '    renderOverviewCharts();' +
    '  } else if (tab === "workload") {' +
    '    var totalCases = d.stewardWorkload.reduce(function(s,w){return s+w.count;},0);' +
    '    var avgCases = d.stewardWorkload.length > 0 ? (totalCases/d.stewardWorkload.length).toFixed(1) : 0;' +
    '    var overloaded = d.stewardWorkload.filter(function(w){return w.status==="OVERLOAD";}).length;' +
    '    html = "<div class=\\"kpi-grid\\" style=\\"grid-template-columns:repeat(4,1fr)\\">" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Total Stewards</div><div class=\\"kpi-value blue\\">" + d.stewardCount + "</div></div>" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Active Cases</div><div class=\\"kpi-value yellow\\">" + totalCases + "</div></div>" +' +
    '      "<div class=\\"kpi-card\\"><div class=\\"kpi-label\\">Avg per Steward</div><div class=\\"kpi-value green\\">" + avgCases + "</div></div>" +' +
    '      "<div class=\\"kpi-card " + (overloaded>0?"alert":"") + "\\"><div class=\\"kpi-label\\">Overloaded</div><div class=\\"kpi-value red\\">" + overloaded + "</div></div>" +' +
    '    "</div>" +' +
    '    "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">assignment_ind</i>Steward Caseload</div><div class=\\"list-container\\">";' +
    '    d.stewardWorkload.forEach(function(w) {' +
    '      html += "<div class=\\"list-item\\"><span>" + w.name + "</span><span class=\\"badge\\" style=\\"background:" + w.color + ";color:white\\">" + w.count + " cases - " + w.status + "</span></div>";' +
    '    });' +
    '    html += "</div></div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '  } else if (tab === "analytics") {' +
    '    html = "<div class=\\"charts-row\\">" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">donut_large</i>Unit Distribution</div><canvas id=\\"unitChart\\"></canvas></div>" +' +
    '      "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">bar_chart</i>Outcomes</div><canvas id=\\"outcomeChart\\"></canvas></div>" +' +
    '    "</div>" +' +
    '    "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">mood</i>Member Satisfaction Score: " + d.moraleScore + "/10</div>" +' +
    '    "<div style=\\"height:20px;background:rgba(255,255,255,0.1);border-radius:10px;overflow:hidden;margin-top:12px\\"><div style=\\"height:100%;width:" + (d.moraleScore*10) + "%;background:linear-gradient(90deg,#22c55e,#3b82f6)\\"></div></div></div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '    renderAnalyticsCharts();' +
    '  } else if (tab === "hotspots") {' +
    '    html = "<div class=\\"chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">local_fire_department</i>Hot Zones (3+ Active Cases)</div>";' +
    '    if (d.hotZones.length === 0) { html += "<div style=\\"text-align:center;padding:40px;color:#94a3b8\\">No hot zones detected - All clear!</div>"; }' +
    '    else { d.hotZones.forEach(function(h) { html += "<div class=\\"hot-zone\\"><span>" + h.location + "</span><span class=\\"badge\\" style=\\"background:#ef4444;color:white\\">" + h.count + " cases</span></div>"; }); }' +
    '    html += "</div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '  } else if (tab === "bargaining") {' +
    '    html = "<div class=\\"bargain-grid\\">" +' +
    '      "<div class=\\"bargain-card\\"><div class=\\"bargain-label\\">Step 1 Denial Rate</div><div class=\\"bargain-value\\">" + d.step1DenialRate + "%</div><div class=\\"bargain-status\\">" + (d.step1DenialRate > 60 ? "High Hostility" : "Normal Range") + "</div></div>" +' +
    '      "<div class=\\"bargain-card\\"><div class=\\"bargain-label\\">Avg Settlement Time</div><div class=\\"bargain-value\\">" + d.avgSettlementDays + " Days</div><div class=\\"bargain-status\\">" + (d.avgSettlementDays > 45 ? "Slower than normal" : "Within range") + "</div></div>" +' +
    '      "<div class=\\"bargain-card\\"><div class=\\"bargain-label\\">Most Violated Article</div><div class=\\"bargain-value\\">" + d.topViolatedArticle + "</div><div class=\\"bargain-status\\">Focus area</div></div>" +' +
    '    "</div>" +' +
    '    "<div class=\\"chart-card\\" style=\\"margin-top:16px\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">article</i>Violations by Contract Article</div><canvas id=\\"bargainChart\\"></canvas></div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '    renderBargainChart();' +
    '  } else if (tab === "satisfaction") {' +
    '    var sat = d.satisfactionData;' +
    '    html = "<div class=\\"sat-header\\"><h2 style=\\"color:#e2e8f0;font-size:16px\\"><i class=\\"material-icons\\" style=\\"vertical-align:middle;margin-right:8px;color:#22c55e\\">sentiment_satisfied</i>Member Satisfaction Survey Analysis</h2><span class=\\"sat-response-count\\">" + sat.responseCount + " Responses</span></div>";' +
    '    html += "<div class=\\"sat-grid\\">";' +
    '    sat.sections.forEach(function(section) {' +
    '      var scoreColor = section.score >= 7 ? "#22c55e" : section.score >= 5 ? "#f59e0b" : "#ef4444";' +
    '      var pct = (section.score / 10) * 100;' +
    '      html += "<div class=\\"sat-section\\">";' +
    '      html += "<div class=\\"sat-section-header\\"><span class=\\"sat-section-name\\">" + section.name + "</span><span class=\\"sat-section-score\\" style=\\"color:" + scoreColor + "\\">" + section.score + "/10</span></div>";' +
    '      html += "<div class=\\"sat-score-bar\\"><div class=\\"sat-score-fill\\" style=\\"width:" + pct + "%;background:" + scoreColor + "\\"></div></div>";' +
    '      html += "<div class=\\"sat-questions\\">";' +
    '      section.questions.forEach(function(q) { html += "<div class=\\"sat-question-item\\">" + q + "</div>"; });' +
    '      html += "</div></div>";' +
    '    });' +
    '    html += "</div>";' +
    '    html += "<div class=\\"sat-overall-chart chart-card\\"><div class=\\"chart-title\\"><i class=\\"material-icons\\">bar_chart</i>Section Score Comparison</div><canvas id=\\"satChart\\"></canvas></div>";' +
    '    document.getElementById("main-content").innerHTML = html;' +
    '    renderSatisfactionChart();' +
    '  }' +
    '}' +
    'function renderOverviewCharts() {' +
    '  var d = dashData;' +
    '  new Chart(document.getElementById("statusChart"),{type:"doughnut",data:{labels:["Open","Pending","Won","Denied","Settled"],datasets:[{data:[d.statusDistribution.open,d.statusDistribution.pending,d.statusDistribution.won,d.statusDistribution.denied,d.statusDistribution.settled],backgroundColor:["#3b82f6","#f59e0b","#22c55e","#ef4444","#8b5cf6"]}]},options:{responsive:true,plugins:{legend:{position:"right",labels:{color:"#cbd5e1",font:{size:10}}}}}});' +
    '  var trendLabels = d.sentimentTrend.length > 0 ? d.sentimentTrend.map(function(t){return t.month;}) : ["Jan","Feb","Mar","Apr","May","Jun"];' +
    '  var trendData = d.sentimentTrend.length > 0 ? d.sentimentTrend.map(function(t){return t.score;}) : [7.2,7.4,7.5,7.6,7.8,7.9];' +
    '  new Chart(document.getElementById("trendChart"),{type:"line",data:{labels:trendLabels,datasets:[{label:"Trust Score",data:trendData,borderColor:"#a78bfa",backgroundColor:"rgba(167,139,250,0.2)",fill:true,tension:0.4}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{min:0,max:10,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    '  var locLabels = Object.keys(d.locationBreakdown).slice(0,6);' +
    '  var locData = locLabels.map(function(l){return d.locationBreakdown[l];});' +
    '  new Chart(document.getElementById("locationChart"),{type:"bar",data:{labels:locLabels,datasets:[{label:"Members",data:locData,backgroundColor:"#3b82f6"}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1"}}}}});' +
    '  var artLabels = Object.keys(d.articleViolations).slice(0,6);' +
    '  var artData = artLabels.map(function(a){return d.articleViolations[a];});' +
    '  new Chart(document.getElementById("articleChart"),{type:"bar",data:{labels:artLabels.length>0?artLabels:["Art 5","Art 7","Art 12"],datasets:[{label:"Cases",data:artData.length>0?artData:[3,5,2],backgroundColor:"#f59e0b"}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    '}' +
    'function renderAnalyticsCharts() {' +
    '  var d = dashData;' +
    '  var unitLabels = Object.keys(d.unitBreakdown).slice(0,8);' +
    '  var unitData = unitLabels.map(function(u){return d.unitBreakdown[u];});' +
    '  new Chart(document.getElementById("unitChart"),{type:"doughnut",data:{labels:unitLabels,datasets:[{data:unitData,backgroundColor:["#3b82f6","#22c55e","#f59e0b","#ef4444","#8b5cf6","#ec4899","#06b6d4","#84cc16"]}]},options:{responsive:true,plugins:{legend:{position:"right",labels:{color:"#cbd5e1",font:{size:10}}}}}});' +
    '  new Chart(document.getElementById("outcomeChart"),{type:"bar",data:{labels:["Won","Denied","Settled"],datasets:[{label:"Cases",data:[d.wins,d.losses,d.settled],backgroundColor:["#22c55e","#ef4444","#8b5cf6"]}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    '}' +
    'function renderBargainChart() {' +
    '  var d = dashData;' +
    '  var artLabels = Object.keys(d.articleViolations).slice(0,8);' +
    '  var artData = artLabels.map(function(a){return d.articleViolations[a];});' +
    '  new Chart(document.getElementById("bargainChart"),{type:"bar",data:{labels:artLabels.length>0?artLabels:["Art 5","Art 7","Art 12"],datasets:[{label:"Violations",data:artData.length>0?artData:[3,5,2],backgroundColor:"#fbbf24"}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:"#94a3b8"}},x:{ticks:{color:"#94a3b8"}}}}});' +
    '}' +
    'function renderSatisfactionChart() {' +
    '  var d = dashData;' +
    '  var labels = d.satisfactionData.sections.map(function(s){return s.name;});' +
    '  var scores = d.satisfactionData.sections.map(function(s){return s.score;});' +
    '  var colors = scores.map(function(s){return s>=7?"#22c55e":s>=5?"#f59e0b":"#ef4444";});' +
    '  new Chart(document.getElementById("satChart"),{type:"bar",data:{labels:labels,datasets:[{label:"Score",data:scores,backgroundColor:colors}]},options:{responsive:true,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{min:0,max:10,ticks:{color:"#94a3b8"}},y:{ticks:{color:"#cbd5e1",font:{size:10}}}}}});' +
    '}' +
    '</script></body></html>';
}

// ============================================================================
// 4. STRATEGIC PRO MOVES & ALERTS
// ============================================================================

/**
 * Checks dashboard metrics and sends email alerts if thresholds exceeded
 */
function checkDashboardAlerts() {
  var metrics = getExecutiveMetrics_();
  var winRate = metrics.winRate;

  var alerts = [];
  if (winRate < 50) alerts.push("CRITICAL WIN RATE: " + winRate + "%");
  if (metrics.overdueSteps > 10) alerts.push("HIGH OVERDUE COUNT: " + metrics.overdueSteps + " cases");

  if (alerts.length > 0) {
    try {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        "509 DASHBOARD ALERT",
        "The following alerts require attention:\n\n" + alerts.join("\n")
      );
      SpreadsheetApp.getActiveSpreadsheet().toast("Alert email sent", "Notification");
    } catch (e) {
      Logger.log('Error sending alert email: ' + e.message);
    }
  }
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * All analytics are now in the unified Steward Dashboard (Bargaining tab).
 */
function renderBargainingCheatSheet() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function renderBargainingCheatSheet_LEGACY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var step1Denials = 0;
  var totalStep1 = 0;
  var settlementDays = [];
  var articleViolations = {};

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getDataRange().getValues();

    for (var g = 1; g < grievanceData.length; g++) {
      if (grievanceData[g][GRIEVANCE_COLS.STEP_1_DATE - 1]) {
        totalStep1++;
        if (grievanceData[g][GRIEVANCE_COLS.STATUS - 1] !== 'Won' &&
            grievanceData[g][GRIEVANCE_COLS.STEP_2_DATE - 1]) {
          step1Denials++;
        }
      }

      var dateFiled = grievanceData[g][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = grievanceData[g][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateFiled instanceof Date && dateClosed instanceof Date) {
        var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
        if (days > 0) settlementDays.push(days);
      }

      var article = grievanceData[g][GRIEVANCE_COLS.ARTICLES - 1];
      if (article) {
        articleViolations[article] = (articleViolations[article] || 0) + 1;
      }
    }
  }

  var denialRate = totalStep1 > 0 ? Math.round((step1Denials / totalStep1) * 100) : 0;
  var avgSettlement = settlementDays.length > 0 ?
    Math.round(settlementDays.reduce(function(a, b) { return a + b; }, 0) / settlementDays.length) : 0;

  var topArticle = "N/A";
  var topCount = 0;
  for (var art in articleViolations) {
    if (articleViolations[art] > topCount) {
      topArticle = art;
      topCount = articleViolations[art];
    }
  }

  // Build modal HTML
  var html = '<!DOCTYPE html><html><head>' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:"Roboto",sans-serif;background:#0f172a;color:#f8fafc;padding:20px;margin:0}' +
    '.header{color:#fbbf24;font-size:18px;font-weight:700;margin-bottom:16px}' +
    '.card{background:rgba(251,191,36,0.1);border:1px solid rgba(251,191,36,0.3);border-radius:8px;padding:16px;margin-bottom:12px}' +
    '.metric-row{display:flex;justify-content:space-between;align-items:center}' +
    '.label{font-size:13px;color:#fcd34d}.value{font-size:24px;font-weight:700;color:#fbbf24}' +
    '.status{font-size:11px;padding:4px 8px;border-radius:4px;margin-top:8px;display:inline-block}' +
    '.status.danger{background:rgba(239,68,68,0.2);color:#f87171}' +
    '.status.warning{background:rgba(251,191,36,0.2);color:#fcd34d}' +
    '.status.ok{background:rgba(16,185,129,0.2);color:#34d399}</style></head><body>' +
    '<div class="header">📊 STRATEGIC BARGAINING DATA</div>' +
    '<div class="card"><div class="metric-row"><div class="label">Step 1 Denial Rate</div><div class="value">' + denialRate + '%</div></div>' +
    '<div class="status ' + (denialRate > 60 ? 'danger' : 'ok') + '">' + (denialRate > 60 ? 'High Hostility - Management aggressive' : 'Normal Range') + '</div></div>' +
    '<div class="card"><div class="metric-row"><div class="label">Average Settlement Time</div><div class="value">' + avgSettlement + ' Days</div></div>' +
    '<div class="status ' + (avgSettlement > 45 ? 'warning' : 'ok') + '">' + (avgSettlement > 45 ? 'Slower than normal' : 'Within expected range') + '</div></div>' +
    '<div class="card"><div class="metric-row"><div class="label">Most Violated Article</div><div class="value">' + topArticle + '</div></div>' +
    '<div class="status ' + (topCount > 5 ? 'danger' : 'warning') + '">' + topCount + ' violations - ' + (topCount > 5 ? 'High Risk Area' : 'Monitor closely') + '</div></div>' +
    '</body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(output, 'Bargaining Intelligence');
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Hot Zones are now in the unified Steward Dashboard (Hot Spots tab).
 */
function renderHotZones() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function renderHotZones_LEGACY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('No Grievance Log found.');
    return;
  }

  var grievanceData = grievanceSheet.getDataRange().getValues();

  // Count grievances by location
  var locationIssues = {};
  for (var g = 1; g < grievanceData.length; g++) {
    var location = grievanceData[g][GRIEVANCE_COLS.LOCATION - 1];
    var status = grievanceData[g][GRIEVANCE_COLS.STATUS - 1];

    if (location && (status === 'Open' || status === 'Pending Info')) {
      locationIssues[location] = (locationIssues[location] || 0) + 1;
    }
  }

  // Sort and identify hot zones (locations with 3+ active issues)
  var hotZones = Object.entries(locationIssues)
    .filter(function(entry) { return entry[1] >= 3; })
    .sort(function(a, b) { return b[1] - a[1]; });

  // Build modal HTML
  var html = '<!DOCTYPE html><html><head>' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:"Roboto",sans-serif;background:#0f172a;color:#f8fafc;padding:20px;margin:0}' +
    '.header{color:#ef4444;font-size:18px;font-weight:700;margin-bottom:16px;display:flex;align-items:center;gap:8px}' +
    '.zone{display:flex;justify-content:space-between;padding:12px;background:rgba(239,68,68,0.1);border-radius:8px;margin-bottom:8px;border-left:4px solid #ef4444}' +
    '.location{font-weight:500}.count{background:#ef4444;color:white;padding:4px 12px;border-radius:12px;font-weight:700}' +
    '.empty{color:#94a3b8;font-style:italic;text-align:center;padding:40px}</style></head><body>' +
    '<div class="header">🔥 HOT ZONES (3+ Active Issues)</div>';

  if (hotZones.length === 0) {
    html += '<div class="empty">No hot zones detected - All clear!</div>';
  } else {
    hotZones.forEach(function(zone) {
      html += '<div class="zone"><span class="location">' + zone[0] + '</span><span class="count">' + zone[1] + ' cases</span></div>';
    });
  }

  html += '</body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(output, 'Hot Zones Analysis');
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Steward metrics are now in the unified Steward Dashboard (Workload tab).
 */
function identifyRisingStars() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function identifyRisingStars_LEGACY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var perfSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);

  if (!perfSheet || perfSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Steward performance data not available. Run system diagnostics first.");
    return;
  }

  var perfData = perfSheet.getDataRange().getValues();

  // Find top performers by score
  var performers = [];
  for (var p = 1; p < perfData.length; p++) {
    if (perfData[p][0] && perfData[p][9]) {  // Name and Score
      performers.push({
        name: perfData[p][0],
        score: perfData[p][9],
        winRate: perfData[p][5] || 0,
        avgDays: perfData[p][6] || 0
      });
    }
  }

  performers.sort(function(a, b) { return b.score - a.score; });
  var risingStars = performers.slice(0, 5);

  // Build modal HTML
  var html = '<!DOCTYPE html><html><head>' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:"Roboto",sans-serif;background:#0f172a;color:#f8fafc;padding:20px;margin:0}' +
    '.header{color:#10b981;font-size:18px;font-weight:700;margin-bottom:16px}' +
    '.star{display:grid;grid-template-columns:2fr 1fr 1fr 1fr;padding:12px;background:rgba(16,185,129,0.1);border-radius:8px;margin-bottom:8px;align-items:center}' +
    '.star.top{border-left:4px solid #fbbf24;background:rgba(251,191,36,0.1)}' +
    '.name{font-weight:600}.metric{text-align:center;font-size:13px}' +
    '.label{font-size:10px;color:#94a3b8;text-transform:uppercase}' +
    '.value{font-weight:700;color:#34d399}' +
    '.empty{color:#94a3b8;font-style:italic;text-align:center;padding:40px}</style></head><body>' +
    '<div class="header">⭐ RISING STARS (Top Performers)</div>';

  if (risingStars.length === 0) {
    html += '<div class="empty">No performance data available</div>';
  } else {
    risingStars.forEach(function(star, index) {
      var isTop = index === 0;
      html += '<div class="star' + (isTop ? ' top' : '') + '">' +
        '<div class="name">' + (isTop ? '🏆 ' : '') + star.name + '</div>' +
        '<div class="metric"><div class="label">Score</div><div class="value">' + star.score + '</div></div>' +
        '<div class="metric"><div class="label">Win Rate</div><div class="value">' + (typeof star.winRate === 'number' ? star.winRate + '%' : star.winRate) + '</div></div>' +
        '<div class="metric"><div class="label">Avg Days</div><div class="value">' + star.avgDays + '</div></div>' +
        '</div>';
    });
  }

  html += '</body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(output, 'Rising Stars');
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Hostility metrics are now in the unified Steward Dashboard (Bargaining tab).
 */
function renderHostilityFunnel() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function renderHostilityFunnel_LEGACY() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('No Grievance Log found.');
    return;
  }

  var grievanceData = grievanceSheet.getDataRange().getValues();

  // Analyze management response patterns
  var step1Count = 0, step1Denied = 0;
  var step2Count = 0, step2Denied = 0;
  var step3Count = 0;
  var arbitrationCount = 0;

  for (var g = 1; g < grievanceData.length; g++) {
    if (grievanceData[g][GRIEVANCE_COLS.STEP_1_DATE - 1]) step1Count++;
    if (grievanceData[g][GRIEVANCE_COLS.STEP_2_DATE - 1]) {
      step2Count++;
      step1Denied++;
    }
    if (grievanceData[g][GRIEVANCE_COLS.STEP_3_DATE - 1]) {
      step3Count++;
      step2Denied++;
    }
    var status = grievanceData[g][GRIEVANCE_COLS.STATUS - 1];
    if (status === 'In Arbitration') arbitrationCount++;
  }

  var step1Rate = 100;
  var step2Rate = step1Count > 0 ? Math.round((step1Denied/step1Count)*100) : 0;
  var step3Rate = step2Count > 0 ? Math.round((step2Denied/step2Count)*100) : 0;
  var arbRate = step3Count > 0 ? Math.round((arbitrationCount/step3Count)*100) : 0;

  // Helper to get color class
  function getRateClass(rate) {
    if (rate >= 60) return 'danger';
    if (rate >= 40) return 'warning';
    return 'ok';
  }

  // Build modal HTML
  var html = '<!DOCTYPE html><html><head>' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:"Roboto",sans-serif;background:#0f172a;color:#f8fafc;padding:20px;margin:0}' +
    '.header{color:#ef4444;font-size:18px;font-weight:700;margin-bottom:16px}' +
    '.funnel{margin-bottom:8px}.step{display:flex;align-items:center;padding:12px;border-radius:8px;margin-bottom:6px}' +
    '.step-label{flex:2;font-weight:500}.step-count{flex:1;text-align:center;font-size:20px;font-weight:700}' +
    '.step-rate{flex:1;text-align:right;padding:4px 12px;border-radius:16px;font-weight:600;font-size:13px}' +
    '.danger{background:rgba(239,68,68,0.2);color:#f87171}' +
    '.warning{background:rgba(251,191,36,0.2);color:#fcd34d}' +
    '.ok{background:rgba(16,185,129,0.2);color:#34d399}' +
    '.step1{background:rgba(59,130,246,0.15);border-left:4px solid #3b82f6}' +
    '.step2{background:rgba(251,191,36,0.15);border-left:4px solid #f59e0b}' +
    '.step3{background:rgba(239,68,68,0.15);border-left:4px solid #ef4444}' +
    '.arb{background:rgba(139,92,246,0.15);border-left:4px solid #8b5cf6}' +
    '.insight{margin-top:16px;padding:12px;background:rgba(255,255,255,0.05);border-radius:8px;font-size:12px;color:#94a3b8}</style></head><body>' +
    '<div class="header">⚡ MANAGEMENT HOSTILITY FUNNEL</div>' +
    '<div class="funnel">' +
    '<div class="step step1"><span class="step-label">Step 1 Filed</span><span class="step-count">' + step1Count + '</span><span class="step-rate ok">' + step1Rate + '%</span></div>' +
    '<div class="step step2"><span class="step-label">Denied → Step 2</span><span class="step-count">' + step1Denied + '</span><span class="step-rate ' + getRateClass(step2Rate) + '">' + step2Rate + '%</span></div>' +
    '<div class="step step3"><span class="step-label">Denied → Step 3</span><span class="step-count">' + step2Denied + '</span><span class="step-rate ' + getRateClass(step3Rate) + '">' + step3Rate + '%</span></div>' +
    '<div class="step arb"><span class="step-label">To Arbitration</span><span class="step-count">' + arbitrationCount + '</span><span class="step-rate ' + getRateClass(arbRate) + '">' + arbRate + '%</span></div>' +
    '</div>' +
    '<div class="insight">' +
    (step2Rate > 60 ? '⚠️ High Step 1 denial rate indicates aggressive management posture. Document everything carefully.' :
     step2Rate > 40 ? '📊 Moderate denial rate. Some management pushback but within normal range.' :
     '✅ Lower denial rate suggests more cooperative management environment.') +
    '</div></body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(output, 'Hostility Funnel');
}

// ============================================================================
// 5. AUTOMATION & COMMUNICATION
// ============================================================================

/**
 * Creates automation triggers for nightly refresh
 */
function createAutomationTriggers() {
  // Check for existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshAllVisuals') {
      SpreadsheetApp.getUi().alert('Automation trigger already exists. No action needed.');
      return;
    }
  }

  // Midnight Refresh Trigger
  ScriptApp.newTrigger('refreshAllVisuals')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  SpreadsheetApp.getUi().alert('Success: Dashboard will now auto-refresh every night at 1:00 AM.');
}

/**
 * Sets up the midnight auto-refresh trigger
 * Runs midnightAutoRefresh daily at midnight (12:00 AM)
 */
function setupMidnightTrigger() {
  // Check for existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  var hasExisting = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'midnightAutoRefresh') {
      hasExisting = true;
      break;
    }
  }

  if (hasExisting) {
    SpreadsheetApp.getUi().alert('Midnight auto-refresh trigger already exists. No action needed.');
    return;
  }

  // Create midnight trigger (12:00 AM)
  ScriptApp.newTrigger('midnightAutoRefresh')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert('Success: Midnight Auto-Refresh is now active.\n\nThe system will automatically refresh all dashboards and check alerts at 12:00 AM daily.');
}

/**
 * Removes the midnight auto-refresh trigger
 */
function removeMidnightTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = false;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'midnightAutoRefresh') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed = true;
    }
  }

  if (removed) {
    SpreadsheetApp.getUi().alert('Midnight auto-refresh trigger has been removed.');
  } else {
    SpreadsheetApp.getUi().alert('No midnight trigger found to remove.');
  }
}

/**
 * Midnight Auto-Refresh function
 * Called automatically by time-based trigger at midnight
 * Refreshes dashboards, hidden sheets, and checks for critical alerts
 */
function midnightAutoRefresh() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var startTime = new Date();

    Logger.log('Midnight Auto-Refresh started at ' + startTime.toISOString());

    // 1. Refresh hidden calculation sheets if function exists
    if (typeof rebuildAllHiddenSheets === 'function') {
      rebuildAllHiddenSheets();
      Logger.log('Hidden calculation sheets refreshed');
    }

    // 2. Check for critical dashboard alerts
    checkDashboardAlerts();
    Logger.log('Dashboard alerts checked');

    // 3. Check for overdue grievances and send reminders
    checkOverdueGrievances_();

    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;

    Logger.log('Midnight Auto-Refresh completed in ' + duration + ' seconds');

    // Note: Executive Command and Member Analytics are now modal-based
    // and don't require midnight refresh - data is fetched on-demand

  } catch (e) {
    Logger.log('Midnight Auto-Refresh error: ' + e.message);
    // Optionally send error notification
    try {
      var adminEmail = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
      if (adminEmail) {
        MailApp.sendEmail(adminEmail,
          COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Auto-Refresh Error',
          'The midnight auto-refresh encountered an error:\n\n' + e.message + '\n\nPlease check the script logs for details.');
      }
    } catch (emailErr) {
      Logger.log('Could not send error notification: ' + emailErr.message);
    }
  }
}

/**
 * Checks for overdue grievances and sends reminder notifications
 * @private
 */
function checkOverdueGrievances_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet || grievanceSheet.getLastRow() < 2) return;

    var data = grievanceSheet.getDataRange().getValues();
    var overdueList = [];

    for (var i = 1; i < data.length; i++) {
      var status = data[i][GRIEVANCE_COLS.STATUS - 1];
      var daysToDeadline = data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

      // Check for active cases that are overdue
      if ((status === 'Open' || status === 'Pending Info') &&
          (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0))) {
        overdueList.push({
          id: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1],
          name: data[i][GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + data[i][GRIEVANCE_COLS.LAST_NAME - 1],
          steward: data[i][GRIEVANCE_COLS.STEWARD - 1],
          days: daysToDeadline
        });
      }
    }

    if (overdueList.length > 0) {
      var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
      if (chiefStewardEmail) {
        var body = 'DAILY OVERDUE GRIEVANCE REPORT\n\n' +
                   'The following ' + overdueList.length + ' grievance(s) have passed their deadline:\n\n';

        overdueList.forEach(function(g) {
          body += '- ' + g.id + ': ' + g.name + ' (Steward: ' + (g.steward || 'Unassigned') + ')\n';
        });

        body += '\nPlease take immediate action to address these cases.' + COMMAND_CONFIG.EMAIL.FOOTER;

        MailApp.sendEmail(chiefStewardEmail,
          COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Daily Overdue Report - ' + overdueList.length + ' Case(s)',
          body);

        Logger.log('Sent overdue report with ' + overdueList.length + ' cases');
      }
    }
  } catch (e) {
    Logger.log('Error checking overdue grievances: ' + e.message);
  }
}

/**
 * @deprecated v4.3.9 - REMOVED DUPLICATE
 * Use refreshAllVisuals() at line 729 instead.
 * This duplicate focused on data refresh while the main one focuses on visuals.
 * The main function now handles both.
 */
function refreshAllVisuals_DataRefresh_DEPRECATED() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Refresh hidden calculation sheets if available
    if (typeof rebuildAllHiddenSheets === 'function') {
      rebuildAllHiddenSheets();
    }

    // Check for alerts
    checkDashboardAlerts();

    ss.toast("Data Refresh Complete - Open dashboards from 509 Command menu", COMMAND_CONFIG.SYSTEM_NAME);
  } catch (e) {
    Logger.log("Refresh Error: " + e.message);
  }
}

/**
 * Sends Member Analytics dashboard access link to specified email
 * Note: Member Analytics is now a modal dashboard launched from the menu
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
      "2. Go to: 509 Command > Command Center > Member Dashboard (No PII)\n" +
      "3. The dashboard shows aggregate union metrics (no personal info visible)\n\n" +
      "WHAT YOU'LL SEE:\n" +
      "- Morale & Trust Scores\n" +
      "- Leadership Pipeline\n" +
      "- Grievance Statistics\n" +
      "- Steward Contact Search\n" +
      "- Emergency Weingarten Rights\n\n" +
      "In Solidarity,\n" +
      "509 Strategic Command Center";

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
 */
function emailDashboardLink() {
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
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();

  if (!email || !email.toString().includes('@')) {
    SpreadsheetApp.getUi().alert('No valid email found for this member.');
    return;
  }

  // Get organization name from config if available
  var orgName = '509';
  try {
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var configOrgName = configSheet.getRange(2, CONFIG_COLS.ORG_NAME).getValue();
      if (configOrgName) orgName = configOrgName;
    }
  } catch (e) {
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
 * Shows Steward Performance Modal
 * Displays performance metrics for all stewards including:
 * - Active cases, total cases, and win rates
 * - Response time averages
 * - Member satisfaction scores (if available)
 */
function showStewardPerformanceModal() {
  var stewardData = getStewardWorkload();

  if (!stewardData || stewardData.length === 0) {
    SpreadsheetApp.getUi().alert('No steward data available. Ensure stewards are marked in the Member Directory.');
    return;
  }

  var html = '<!DOCTYPE html><html><head>' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Segoe UI", Roboto, sans-serif; background: #f0f4f8; padding: 15px; }' +
    '.header { display: flex; align-items: center; color: #059669; margin-bottom: 15px; }' +
    '.header h2 { font-size: 18px; font-weight: 600; margin-left: 8px; }' +
    '.stats-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 15px; }' +
    '.stat-box { background: white; border-radius: 10px; padding: 12px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }' +
    '.stat-box .value { font-size: 24px; font-weight: 700; color: #059669; }' +
    '.stat-box .label { font-size: 10px; text-transform: uppercase; color: #64748b; }' +
    '.steward-table { width: 100%; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }' +
    '.steward-table th { background: #059669; color: white; padding: 10px; font-size: 11px; text-transform: uppercase; text-align: left; }' +
    '.steward-table td { padding: 10px; font-size: 12px; border-bottom: 1px solid #f1f5f9; }' +
    '.steward-table tr:hover { background: #f8fafc; }' +
    '.win-rate { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }' +
    '.win-high { background: #dcfce7; color: #166534; }' +
    '.win-med { background: #fef3c7; color: #92400e; }' +
    '.win-low { background: #fee2e2; color: #991b1b; }' +
    '</style></head><body>' +
    '<div class="header"><i class="material-icons">shield</i><h2>Steward Performance</h2></div>';

  // Calculate totals
  var totalActive = 0, totalCases = 0, totalWon = 0;
  stewardData.forEach(function(s) {
    totalActive += s.activeCases || 0;
    totalCases += s.totalCases || 0;
    totalWon += s.wonCases || 0;
  });
  var overallWinRate = totalCases > 0 ? Math.round((totalWon / totalCases) * 100) : 0;

  html += '<div class="stats-row">' +
    '<div class="stat-box"><div class="value">' + stewardData.length + '</div><div class="label">Active Stewards</div></div>' +
    '<div class="stat-box"><div class="value">' + totalActive + '</div><div class="label">Active Cases</div></div>' +
    '<div class="stat-box"><div class="value">' + overallWinRate + '%</div><div class="label">Overall Win Rate</div></div>' +
    '</div>';

  html += '<table class="steward-table"><thead><tr>' +
    '<th>Steward</th><th>Unit</th><th>Active</th><th>Total</th><th>Win Rate</th>' +
    '</tr></thead><tbody>';

  stewardData.forEach(function(s) {
    var firstName = s['First Name'] || '';
    var lastName = s['Last Name'] || '';
    var unit = s['Unit'] || 'General';
    var winClass = s.winRate >= 70 ? 'win-high' : (s.winRate >= 40 ? 'win-med' : 'win-low');

    html += '<tr>' +
      '<td><strong>' + firstName + ' ' + lastName + '</strong></td>' +
      '<td>' + unit + '</td>' +
      '<td>' + (s.activeCases || 0) + '</td>' +
      '<td>' + (s.totalCases || 0) + '</td>' +
      '<td><span class="win-rate ' + winClass + '">' + (s.winRate || 0) + '%</span></td>' +
      '</tr>';
  });

  html += '</tbody></table></body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(output, 'Steward Performance Dashboard');
}

/**
 * Emails the Executive Dashboard as a PDF snapshot
 */
function emailExecutivePDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var date = new Date();
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var blob = ss.getAs('application/pdf').setName("509_Health_Report_" + dateStr + ".pdf");

    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + " Weekly Executive Summary - " + dateStr,
      body: "Attached is the latest strategic briefing." + COMMAND_CONFIG.EMAIL.FOOTER,
      attachments: [blob]
    });

    SpreadsheetApp.getUi().alert('PDF report sent to your email.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error generating PDF: ' + e.message);
  }
}

// ============================================================================
// 6. AUTO-ID GENERATOR ENGINE
// ============================================================================

/**
 * Generates missing Member IDs - UI Service version (Legacy)
 * NOTE: Renamed to avoid duplicate. Use generateMissingMemberIDs() from 02_MemberManager.gs
 * @deprecated Use generateMissingMemberIDs() from 02_MemberManager.gs
 */
function generateMissingMemberIDs_UIService_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory sheet not found.');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var countAdded = 0;

  // Get unit codes from Config sheet (falls back to defaults if not configured)
  var unitCodes = getUnitCodes_();

  for (var i = 1; i < data.length; i++) {
    var memberId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    var unit = data[i][MEMBER_COLS.UNIT - 1];

    // ID is empty but Unit is present
    if (!memberId && unit) {
      var prefix = unitCodes[unit] || "GEN";
      var nextNum = getNextMemberSequence_(prefix);
      var newId = prefix + "-" + nextNum + "-H";

      sheet.getRange(i + 1, MEMBER_COLS.MEMBER_ID).setValue(newId);
      countAdded++;
    }
  }

  ss.toast(countAdded + ' IDs generated for the 509 Command Center.', 'ID Engine', 5);
}

/**
 * Gets the next sequence number for a given prefix
 * @param {string} prefix - The unit code prefix (e.g., "MS")
 * @returns {number} The next sequence number
 * @private
 */
function getNextMemberSequence_(prefix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) return 100;

  var ids = sheet.getRange(1, MEMBER_COLS.MEMBER_ID, sheet.getLastRow(), 1).getValues().flat();
  var max = 100;

  ids.forEach(function(id) {
    if (typeof id === 'string' && id.startsWith(prefix + '-')) {
      var parts = id.split('-');
      if (parts.length >= 2) {
        var n = parseInt(parts[1]);
        if (!isNaN(n) && n > max) max = n;
      }
    }
  });

  return max + 1;
}

/**
 * Checks for duplicate Member IDs - UI Service version (Legacy)
 * NOTE: Renamed to avoid duplicate. Use checkDuplicateMemberIDs() from 02_MemberManager.gs
 * @deprecated Use checkDuplicateMemberIDs() from 02_MemberManager.gs
 */
function checkDuplicateMemberIDs_UIService_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory sheet not found.');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var idCounts = {};
  var duplicates = [];

  // Count occurrences of each ID
  for (var i = 1; i < data.length; i++) {
    var id = data[i][MEMBER_COLS.MEMBER_ID - 1];
    if (id) {
      idCounts[id] = (idCounts[id] || 0) + 1;
      if (idCounts[id] === 2) {
        duplicates.push(id);
      }
    }
  }

  if (duplicates.length === 0) {
    ss.toast('No duplicate Member IDs found.', 'Duplicate Check', 3);
  } else {
    // Highlight duplicates
    for (var j = 1; j < data.length; j++) {
      var memberId = data[j][MEMBER_COLS.MEMBER_ID - 1];
      if (duplicates.indexOf(memberId) !== -1) {
        sheet.getRange(j + 1, MEMBER_COLS.MEMBER_ID).setBackground('#FCA5A5');
      }
    }
    SpreadsheetApp.getUi().alert('Found ' + duplicates.length + ' duplicate IDs. They have been highlighted in red.');
  }
}

// ============================================================================
// 7. SIGNATURE PDF ENGINE
// ============================================================================

/**
 * Creates a grievance PDF document with signature blocks
 * Uses Google Docs template if configured, otherwise creates from scratch
 * @param {Object} data - Grievance data object with name, details, etc.
 * @returns {File} The created PDF file
 */
function createGrievancePDF(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create archive folder
  var folder;
  if (COMMAND_CONFIG.ARCHIVE_FOLDER_ID) {
    try {
      folder = DriveApp.getFolderById(COMMAND_CONFIG.ARCHIVE_FOLDER_ID);
    } catch (e) {
      folder = DriveApp.createFolder('509 Grievance Archive');
    }
  } else {
    // Create folder in root if not configured
    var folders = DriveApp.getFoldersByName('509 Grievance Archive');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('509 Grievance Archive');
  }

  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Check if template is configured
  if (COMMAND_CONFIG.TEMPLATE_ID) {
    try {
      var tempFile = DriveApp.getFileById(COMMAND_CONFIG.TEMPLATE_ID)
        .makeCopy('SIGNATURE_REQUIRED_' + data.name + '_' + dateStr, folder);
      var doc = DocumentApp.openById(tempFile.getId());
      var body = doc.getBody();

      // Replace placeholders
      body.replaceText('{{MemberName}}', data.name || '');
      body.replaceText('{{Date}}', dateStr);
      body.replaceText('{{Details}}', data.details || '');
      body.replaceText('{{GrievanceID}}', data.grievanceId || '');
      body.replaceText('{{Status}}', data.status || '');

      // Add signature blocks
      body.appendParagraph(COMMAND_CONFIG.PDF.SIGNATURE_BLOCK);

      doc.saveAndClose();

      // Convert to PDF
      var pdf = folder.createFile(tempFile.getAs(MimeType.PDF))
        .setName('Grievance_UNSIGNED_' + data.name + '_' + dateStr + '.pdf');
      tempFile.setTrashed(true);

      return pdf;
    } catch (e) {
      Logger.log('Template error, creating from scratch: ' + e.message);
    }
  }

  // Create document from scratch if no template
  var doc = DocumentApp.create('SIGNATURE_REQUIRED_' + data.name + '_' + dateStr);
  var body = doc.getBody();

  // Header
  body.appendParagraph(COMMAND_CONFIG.SYSTEM_NAME)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('GRIEVANCE DOCUMENT - REQUIRES SIGNATURE')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // Grievance details
  body.appendParagraph('Grievance ID: ' + (data.grievanceId || 'Pending'));
  body.appendParagraph('Member Name: ' + (data.name || 'Not specified'));
  body.appendParagraph('Date: ' + dateStr);
  body.appendParagraph('Status: ' + (data.status || 'Open'));

  body.appendParagraph('');
  body.appendParagraph('Details:').setBold(true);
  body.appendParagraph(data.details || 'No details provided.');

  // Signature blocks
  body.appendParagraph(COMMAND_CONFIG.PDF.SIGNATURE_BLOCK);

  doc.saveAndClose();

  // Move to archive folder and convert to PDF
  var docFile = DriveApp.getFileById(doc.getId());
  var pdf = folder.createFile(docFile.getAs(MimeType.PDF))
    .setName('Grievance_UNSIGNED_' + data.name + '_' + dateStr + '.pdf');

  docFile.setTrashed(true);

  SpreadsheetApp.getActiveSpreadsheet().toast('PDF created: ' + pdf.getName(), 'PDF Engine', 5);
  return pdf;
}

/**
 * Creates PDF for the currently selected grievance row
 */
function createPDFForSelectedGrievance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a grievance row (not the header).');
    return;
  }

  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var grievanceData = {
    grievanceId: data[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
    name: data[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + data[GRIEVANCE_COLS.LAST_NAME - 1],
    status: data[GRIEVANCE_COLS.STATUS - 1],
    details: 'Category: ' + (data[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A') +
             '\nArticles: ' + (data[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A') +
             '\nLocation: ' + (data[GRIEVANCE_COLS.LOCATION - 1] || 'N/A') +
             '\nSteward: ' + (data[GRIEVANCE_COLS.STEWARD - 1] || 'N/A') +
             '\nResolution: ' + (data[GRIEVANCE_COLS.RESOLUTION - 1] || 'Pending')
  };

  var pdf = createGrievancePDF(grievanceData);

  SpreadsheetApp.getUi().alert('PDF created successfully!\n\nFile: ' + pdf.getName() +
    '\n\nYou can find it in your 509 Grievance Archive folder.');
}

// ============================================================================
// 8. STEWARD PROMOTION ENGINE
// ============================================================================

/**
 * Promotes the selected member to Steward status
 * Sends steward toolkit email to the promoted member
 * Requires two confirmation dialogs for safety
 * @deprecated v4.3.9 - DUPLICATE: Primary function is in 02_MemberManager.gs:559
 */
function promoteSelectedMemberToSteward_UIService_DEPRECATED() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Member Directory sheet not found.');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Please select a member row (not the header).');
    return;
  }

  // Get member details early for confirmation dialogs
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
  var name = firstName + ' ' + lastName;

  // Check if already a steward
  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();
  if (currentStatus === 'Yes') {
    ui.alert('This member is already a Steward.');
    return;
  }

  // WARNING 1: Initial confirmation
  var response1 = ui.alert(
    '⬆️ Promote to Steward - Step 1 of 2',
    'You are about to promote ' + name + ' to Steward status.\n\n' +
    'This will:\n' +
    '• Set "Is Steward" to Yes\n' +
    '• Send steward toolkit email (if email on file)\n' +
    '• Grant access to steward-level functions\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('Promotion cancelled.');
    return;
  }

  // WARNING 2: Final confirmation
  var response2 = ui.alert(
    '⚠️ Final Confirmation - Step 2 of 2',
    'PLEASE CONFIRM: You are promoting ' + name + ' to Steward.\n\n' +
    'This action grants significant responsibilities including:\n' +
    '• Representing members in grievances\n' +
    '• Access to sensitive member information\n' +
    '• Authority to act on behalf of the union\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    ui.alert('Promotion cancelled.');
    return;
  }

  // Promote to steward
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('Yes');

  // Get email for toolkit
  var email = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();

  // Send toolkit email if email is available
  if (email) {
    sendStewardToolkit_(email, name);
  }

  ui.alert('✅ ' + name + ' has been promoted to Steward!' +
    (email ? '\n\nToolkit email sent to: ' + email : '\n\nNo email on file - please send toolkit manually.'));
}

/**
 * Sends steward toolkit email to newly promoted steward
 * @param {string} email - Steward's email address
 * @param {string} name - Steward's name
 * @private
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
 * Demotes the selected steward back to regular member status
 * Requires two confirmation dialogs for safety
 * @deprecated v4.3.9 - DUPLICATE: Primary function is in 02_MemberManager.gs:632
 */
function demoteSelectedSteward_UIService_DEPRECATED() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Member Directory sheet not found.');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Please select a member row (not the header).');
    return;
  }

  // Get member details early for confirmation dialogs
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
  var name = firstName + ' ' + lastName;

  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();
  if (currentStatus !== 'Yes') {
    ui.alert('This member is not currently a Steward.');
    return;
  }

  // WARNING 1: Initial confirmation
  var response1 = ui.alert(
    '⬇️ Demote Steward - Step 1 of 2',
    'You are about to remove Steward status from ' + name + '.\n\n' +
    'This will:\n' +
    '• Set "Is Steward" to No\n' +
    '• Clear committee assignments\n' +
    '• Remove steward-level access\n\n' +
    'Do you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    ui.alert('Demotion cancelled.');
    return;
  }

  // WARNING 2: Final confirmation
  var response2 = ui.alert(
    '⚠️ Final Confirmation - Step 2 of 2',
    'PLEASE CONFIRM: You are removing Steward status from ' + name + '.\n\n' +
    'This is a significant action that:\n' +
    '• Removes their authority to represent members\n' +
    '• Should be documented appropriately\n' +
    '• May require notification to the member\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    ui.alert('Demotion cancelled.');
    return;
  }

  // Demote steward
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('No');
  sheet.getRange(row, MEMBER_COLS.COMMITTEES).setValue('');  // Clear committee assignments

  ui.alert('✅ ' + name + ' has been removed from Steward status.\n\n' +
    'Committee assignments have been cleared.');
}

// ============================================================================
// 9. GLOBAL UI STYLING ENGINE (Status Colors)
// ============================================================================

/**
 * Applies status-based coloring to the Grievance Log
 * Called separately to update colors when statuses change
 */
function applyStatusColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  var lastDataRow = sheet.getLastRow();
  if (lastDataRow < 2) return;

  for (var row = 2; row <= lastDataRow; row++) {
    var statusCell = sheet.getRange(row, GRIEVANCE_COLS.STATUS);
    var status = statusCell.getValue();

    if (status && COMMAND_CONFIG.STATUS_COLORS[status]) {
      var colors = COMMAND_CONFIG.STATUS_COLORS[status];
      statusCell.setBackground(colors.bg)
                .setFontColor(colors.text)
                .setFontWeight('bold');
    }
  }

  ss.toast('Status colors applied to Grievance Log.', 'Style Engine', 3);
}
