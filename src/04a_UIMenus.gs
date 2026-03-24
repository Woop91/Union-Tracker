/**
 * ============================================================================
 * 04a_UIMenus.gs - Menu System & Navigation
 * ============================================================================
 *
 * This module handles menu creation and navigation including:
 * - Menu bar creation and organization
 * - Navigation sidebar
 * - Quick action menus
 * - Visual control panel dialogs
 *
 * SEPARATION OF CONCERNS: This file contains ONLY UI/presentation logic.
 * All business logic should remain in their respective modules.
 *
 * @fileoverview Menu system, navigation, and dialog management
 * @version 4.36.0
 * @requires 01_Core.gs
 */

// ACCEPTABLE: 40+ menu items match feature breadth for power-user audience
// ============================================================================
// MENU CREATION
// ============================================================================

/**
 * Creates the custom menu when the spreadsheet opens
 * Consolidated menu structure (v4.4.0):
 * - 📊 Union Hub: Primary operations (dashboards, search, members, cases)
 * - 🔧 Tools: Supporting features (calendar, drive, notifications, etc.)
 * - 🛠️ Admin: System administration
 * Called automatically by onOpen trigger
 */

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
      ${getMobileOptimizedHead()}
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
            .withFailureHandler(function(e){showStatus("Error: "+e.message)})
            .saveVisualSetting(setting, isChecked);
        }

        function selectTheme(theme) {
          document.querySelectorAll('.theme-btn').forEach(btn => btn.classList.remove('active'));
          event.target.classList.add('active');
          google.script.run
            .withSuccessHandler(showStatus)
            .withFailureHandler(function(e){showStatus("Error: "+e.message)})
            .applyDashboardTheme(theme);
        }

        function refreshDashboard() {
          showStatus('Refreshing...');
          google.script.run
            .withSuccessHandler(function() { showStatus('Dashboard refreshed!'); })
            .withFailureHandler(function(e){showStatus("Error: "+e.message)})
            .syncAllDashboardData();
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
 * Navigates to the Member Directory sheet and prompts the user to filter by stewards.
 * @returns {void}
 */
function showStewardDirectory() {
  // Navigate to Member Directory filtered by stewards
  navigateToSheet(SHEETS.MEMBER_DIR);
  SpreadsheetApp.getActive().toast('Filter by "Is Steward = Yes" to see steward directory', 'Steward Directory', 5);
}

// ============================================================================
// NAVIGATION HELPERS (Strategic Command Center)
// ============================================================================

// ============================================================================
// MULTI-SELECT DIALOGS
// ============================================================================

/**
 * Opens the multi-select editor for the currently selected cell
 * Called from menu: Tools > Multi-Select > Open Editor
 * Supports Member Directory and Grievance Log multi-select columns
 */
function openCellMultiSelectEditor() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = ss.getActiveRange();
  var sheetName = sheet.getName();

  // Validate we're in a supported sheet
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('Multi-Select Editor',
      'Please select a cell in the Member Directory or Grievance Log sheet.\n\n' +
      'Member Directory columns:\n' +
      '• Office Days\n' +
      '• Preferred Communication\n' +
      '• Best Time to Contact\n' +
      '• Committees\n' +
      '• Assigned Steward(s)\n\n' +
      'Grievance Log columns:\n' +
      '• Articles Violated\n' +
      '• Issue Category',
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

  // Check if this is a multi-select column for the active sheet
  var config = getMultiSelectConfig(col, sheetName);
  if (!config) {
    var hint = (sheetName === SHEETS.GRIEVANCE_LOG)
      ? 'Grievance Log multi-select columns:\n• Articles Violated\n• Issue Category'
      : 'Member Directory multi-select columns:\n• Office Days\n• Preferred Communication\n• Best Time to Contact\n• Committees\n• Assigned Steward(s)';
    ui.alert('Multi-Select Editor',
      'This column does not support multi-select.\n\n' + hint,
      ui.ButtonSet.OK);
    return;
  }

  // Store target cell coordinates and sheet for the callback
  var props = PropertiesService.getUserProperties();
  props.setProperty('multiSelectRow', row.toString());
  props.setProperty('multiSelectCol', col.toString());
  props.setProperty('multiSelectSheet', sheetName);

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
  var allowedCallbacks = ['applyMultiSelectValue', 'handleBulkStatusSelection'];
  if (allowedCallbacks.indexOf(callback) === -1) {
    throw new Error('Invalid callback function name: ' + callback);
  }
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
            <button class="btn btn-secondary" onclick="cancelSelection()">Cancel</button>
            <button class="btn btn-primary" onclick="submitSelection()">Apply</button>
          </div>
        </div>
      </div>

      <script>
        ${getClientSideEscapeHtml()}
        const items = ${itemsJson};
        const callbackFn = '${callback}';
        let filteredItems = [...items];

        function renderItems() {
          const container = document.getElementById('itemsList');
          container.innerHTML = filteredItems.map(function(item) {
            return '<div class="item-row">' +
                   '<input type="checkbox" class="item-checkbox" data-id="' + escapeHtml(item.id) + '"' +
                   (item.selected ? ' checked' : '') + ' onchange="updateCount()">' +
                   '<span class="item-label">' + escapeHtml(item.label) + '</span>' +
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

        function cancelSelection() {
          google.script.run.clearMultiSelectState();
          google.script.host.close();
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
// MODAL HUB — Centralized Modal Launcher (v4.36.0)
// ============================================================================

/**
 * Registry of all user-facing modals/dialogs organized by category.
 * Each entry: { label, fn (function name string), icon, desc }
 * @const {Object}
 */
var MODAL_REGISTRY = {
  'Members': [
    { label: 'Add New Member',       fn: 'showNewMemberDialog',       icon: '\u2795', desc: 'Add a new member to the directory' },
    { label: 'Find Member',          fn: 'showSearchDialog',          icon: '\uD83D\uDD0D', desc: 'Search for a member by name or ID' },
    { label: 'Import Members',       fn: 'showImportMembersDialog',   icon: '\uD83D\uDCE5', desc: 'Import members from CSV' },
    { label: 'Export Directory',     fn: 'showExportMembersDialog',   icon: '\uD83D\uDCE4', desc: 'Export member directory' },
    { label: 'Add Non-Member Contact', fn: 'showAddContactModal',    icon: '\uD83D\uDC65', desc: 'Add a non-member contact' }
  ],
  'Cases & Grievances': [
    { label: 'New Case/Grievance',   fn: 'showNewGrievanceDialog',    icon: '\u2795', desc: 'File a new case or grievance' },
    { label: 'Edit Selected Case',   fn: 'showEditGrievanceDialog',   icon: '\u270F\uFE0F', desc: 'Edit the currently selected case' },
    { label: 'View Checklist',       fn: 'showChecklistDialog',       icon: '\u2705', desc: 'View case checklist steps' },
    { label: 'Bulk Actions',         fn: 'showBulkActionsDialog',     icon: '\uD83D\uDCCB', desc: 'Perform bulk actions on selected cases' },
    { label: 'Bulk Update Status',   fn: 'showBulkStatusUpdate',      icon: '\uD83D\uDD04', desc: 'Update status on multiple cases' },
    { label: 'Case Checklist Progress', fn: 'showCaseProgressModal',  icon: '\u2705', desc: 'View case checklist progress summary' }
  ],
  'Search & Analytics': [
    { label: 'Desktop Search',       fn: 'showDesktopSearch',         icon: '\uD83D\uDD0D', desc: 'Full desktop search interface' },
    { label: 'Search Precedents',    fn: 'showSearchPrecedents',      icon: '\uD83D\uDCDA', desc: 'Search past case precedents' },
    { label: 'Grievance Trends',     fn: 'showGrievanceTrends',       icon: '\uD83D\uDCCA', desc: 'View grievance trend analytics' },
    { label: 'Unit Health Report',   fn: 'showUnitHealthReport',      icon: '\uD83C\uDFE5', desc: 'View unit health report' }
  ],
  'Calendar & Meetings': [
    { label: 'Setup New Meeting',    fn: 'showSetupMeetingDialog',    icon: '\uD83D\uDCDD', desc: 'Create a new meeting' },
    { label: 'Meeting Check-In',     fn: 'showMeetingCheckInDialog',  icon: '\u2705', desc: 'Open meeting check-in attendance' },
    { label: 'Add New Event',        fn: 'showAddEventModal',         icon: '\uD83D\uDCC5', desc: 'Add a calendar event' },
    { label: 'Add Meeting Minutes',  fn: 'showAddMinutesModal',       icon: '\uD83D\uDCDD', desc: 'Record meeting minutes' },
    { label: 'Take Attendance',      fn: 'showTakeAttendanceModal',   icon: '\uD83D\uDCCB', desc: 'Take meeting attendance' }
  ],
  'Surveys & Polls': [
    { label: 'Edit Survey Question', fn: 'showQuestionEditorModal',   icon: '\uD83D\uDCCB', desc: 'Edit a survey question' },
    { label: 'Create Weekly Question', fn: 'showCreateWeeklyQuestionModal', icon: '\u26A1', desc: 'Create a weekly poll question' },
    { label: 'View Survey Responses', fn: 'showSurveyResponseViewer', icon: '\uD83D\uDCCA', desc: 'View survey response data' }
  ],
  'Tools & Productivity': [
    { label: 'Visual Control Panel', fn: 'showVisualControlPanel',    icon: '\uD83C\uDF9B\uFE0F', desc: 'Toggle visual settings (dark mode, themes)' },
    { label: 'Multi-Select Editor',  fn: 'openCellMultiSelectEditor', icon: '\uD83D\uDCDD', desc: 'Multi-select editor for cells' },
    { label: 'Submit Feedback',      fn: 'showSubmitFeedbackModal',   icon: '\uD83D\uDCA1', desc: 'Submit feedback or suggestions' },
    { label: 'Log Volunteer Hours',  fn: 'showLogVolunteerHoursModal', icon: '\uD83E\uDD1D', desc: 'Log volunteer service hours' },
    { label: 'OCR Transcribe Form',  fn: 'showOCRDialog',            icon: '\uD83D\uDCDD', desc: 'OCR transcribe a physical form' }
  ],
  'Web App & Portal': [
    { label: 'Generate Member PIN',  fn: 'showGeneratePINDialog',     icon: '\uD83D\uDD11', desc: 'Generate a PIN for a member' },
    { label: 'Reset Member PIN',     fn: 'showResetPINDialog',        icon: '\uD83D\uDD04', desc: 'Reset a member PIN' },
    { label: 'Bulk Generate PINs',   fn: 'showBulkGeneratePINDialog', icon: '\uD83D\uDCCB', desc: 'Generate PINs for all members' }
  ],
  'Admin & System': [
    { label: 'Settings',             fn: 'showSettingsDialog',        icon: '\u2699\uFE0F', desc: 'System settings' },
    { label: 'Admin Settings',       fn: 'showAdminSettingsSidebar',  icon: '\uD83D\uDEE0\uFE0F', desc: 'Admin settings sidebar' },
    { label: 'System Diagnostics',   fn: 'showDiagnosticsDialog',     icon: '\uD83E\uDE7A', desc: 'Run system diagnostics' },
    { label: 'Modal Diagnostics',    fn: 'showModalDiagnostics',      icon: '\uD83D\uDD0D', desc: 'Diagnose modal issues' },
    { label: 'Repair Dashboard',     fn: 'showRepairDialog',          icon: '\uD83D\uDD27', desc: 'Repair dashboard components' },
    { label: 'Cache Status',         fn: 'showCacheStatusDashboard',  icon: '\uD83D\uDDC4\uFE0F', desc: 'View cache status dashboard' },
    { label: 'Welcome / Setup Wizard', fn: 'showWelcomeWizardModal',  icon: '\uD83C\uDFD7\uFE0F', desc: 'Run the welcome setup wizard' },
    { label: 'Help Guide',           fn: 'showSearchableHelpModal',   icon: '\u2753', desc: 'Searchable help guide' }
  ]
};

/**
 * Returns true if the Modal Hub feature is enabled via Config.
 * Defaults to true if the config value is missing or blank.
 * @returns {boolean}
 */
function _isModalHubEnabled() {
  try {
    var val = getConfigValue_(CONFIG_COLS.ENABLE_MODAL_HUB);
    if (!val || String(val).trim() === '') return true;
    return isTruthyValue(val);
  } catch (_e) { return true; }
}

/**
 * Shows the centralized Modal Hub dialog.
 * All modals in one place, searchable, with a master enable/disable toggle.
 */
function showModalHub() {
  var html = HtmlService.createHtmlOutput(getModalHubHtml_())
    .setWidth(DIALOG_SIZES.LARGE.width)
    .setHeight(DIALOG_SIZES.LARGE.height);
  SpreadsheetApp.getUi().showModalDialog(html, 'Modal Hub');
}

/**
 * Toggles the Modal Hub enable state (persisted to Config sheet).
 * @param {boolean} enabled - Whether modals should be enabled
 * @returns {{success: boolean}}
 */
function setModalHubEnabled(enabled) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) return errorResponse('Config sheet not found');
    configSheet.getRange(3, CONFIG_COLS.ENABLE_MODAL_HUB).setValue(enabled ? 'Yes' : 'No');
    return successResponse(null, 'Modal Hub ' + (enabled ? 'enabled' : 'disabled'));
  } catch (e) {
    return errorResponse(e.message, 'setModalHubEnabled');
  }
}

/**
 * Server callback: launches a modal by function name from the Modal Hub.
 * Validates the function name exists in the registry before calling.
 * @param {string} fnName - Global function name to call
 * @returns {{success: boolean}}
 */
function launchModalFromHub(fnName) {
  try {
    // Validate fnName is in the registry (security: prevent arbitrary function calls)
    var valid = false;
    var categories = Object.keys(MODAL_REGISTRY);
    for (var c = 0; c < categories.length; c++) {
      var modals = MODAL_REGISTRY[categories[c]];
      for (var m = 0; m < modals.length; m++) {
        if (modals[m].fn === fnName) { valid = true; break; }
      }
      if (valid) break;
    }
    if (!valid) return errorResponse('Invalid modal function: ' + fnName);

    // Close the hub first, then launch the target modal
    var targetFn = this[fnName] || globalThis[fnName];
    if (typeof targetFn !== 'function') return errorResponse('Function not found: ' + fnName);
    targetFn();
    return successResponse();
  } catch (e) {
    return errorResponse(e.message, 'launchModalFromHub');
  }
}

/**
 * Generates the HTML for the Modal Hub dialog.
 * @returns {string} Complete HTML document
 * @private
 */
function getModalHubHtml_() {
  var enabled = _isModalHubEnabled();
  // Build the registry JSON for client-side rendering
  var registryJson = JSON.stringify(MODAL_REGISTRY);

  return '<!DOCTYPE html>' +
  '<html><head><base target="_top">' +
  getMobileOptimizedHead() +
  '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Google Sans", -apple-system, BlinkMacSystemFont, sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; padding: 20px; }' +
    '.hub-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }' +
    '.hub-header h1 { font-size: 22px; font-weight: 700; }' +
    '.hub-header .subtitle { font-size: 12px; color: #94A3B8; margin-top: 2px; }' +
    '.master-toggle { display: flex; align-items: center; gap: 10px; background: rgba(255,255,255,0.08); padding: 8px 16px; border-radius: 10px; }' +
    '.master-toggle label { font-size: 13px; color: #CBD5E1; cursor: pointer; }' +
    '.toggle-switch { position: relative; width: 44px; height: 24px; }' +
    '.toggle-switch input { opacity: 0; width: 0; height: 0; }' +
    '.toggle-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background: #475569; border-radius: 24px; transition: 0.3s; }' +
    '.toggle-slider:before { content: ""; position: absolute; height: 18px; width: 18px; left: 3px; bottom: 3px; background: white; border-radius: 50%; transition: 0.3s; }' +
    '.toggle-switch input:checked + .toggle-slider { background: #7C3AED; }' +
    '.toggle-switch input:checked + .toggle-slider:before { transform: translateX(20px); }' +
    '.search-bar { width: 100%; padding: 10px 16px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.15); background: rgba(255,255,255,0.06); color: #F8FAFC; font-size: 14px; margin-bottom: 16px; outline: none; }' +
    '.search-bar:focus { border-color: #7C3AED; box-shadow: 0 0 0 2px rgba(124,58,237,0.3); }' +
    '.search-bar::placeholder { color: #64748B; }' +
    '.category { margin-bottom: 14px; }' +
    '.category-title { font-size: 12px; text-transform: uppercase; letter-spacing: 1.2px; color: #94A3B8; margin-bottom: 8px; padding-left: 4px; }' +
    '.modal-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px; }' +
    '.modal-card { display: flex; align-items: center; gap: 10px; padding: 10px 14px; background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.08); border-radius: 10px; cursor: pointer; transition: all 0.2s; }' +
    '.modal-card:hover { background: rgba(124,58,237,0.15); border-color: rgba(124,58,237,0.4); transform: translateY(-1px); }' +
    '.modal-card.disabled { opacity: 0.35; pointer-events: none; }' +
    '.modal-icon { font-size: 20px; flex-shrink: 0; width: 32px; text-align: center; }' +
    '.modal-info { min-width: 0; }' +
    '.modal-label { font-size: 13px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }' +
    '.modal-desc { font-size: 11px; color: #94A3B8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }' +
    '.no-results { text-align: center; color: #64748B; padding: 40px 0; font-size: 14px; }' +
    '.status-bar { display: flex; justify-content: space-between; align-items: center; margin-top: 16px; padding-top: 12px; border-top: 1px solid rgba(255,255,255,0.08); }' +
    '.status-bar span { font-size: 11px; color: #64748B; }' +
    '.btn-close { padding: 6px 18px; background: rgba(255,255,255,0.1); color: #CBD5E1; border: none; border-radius: 8px; cursor: pointer; font-size: 13px; }' +
    '.btn-close:hover { background: rgba(255,255,255,0.18); }' +
    '@media (max-width: 500px) { .modal-grid { grid-template-columns: 1fr; } .hub-header { flex-direction: column; gap: 10px; align-items: flex-start; } }' +
  '</style></head><body>' +
  '<div class="hub-header">' +
    '<div><h1>Modal Hub</h1><div class="subtitle">All dialogs in one place</div></div>' +
    '<div class="master-toggle">' +
      '<label for="masterToggle">Enable All</label>' +
      '<div class="toggle-switch">' +
        '<input type="checkbox" id="masterToggle" ' + (enabled ? 'checked' : '') + ' onchange="toggleMaster(this.checked)">' +
        '<span class="toggle-slider"></span>' +
      '</div>' +
    '</div>' +
  '</div>' +
  '<input type="text" class="search-bar" id="searchBar" placeholder="Search modals..." oninput="filterModals(this.value)">' +
  '<div id="modalList"></div>' +
  '<div class="status-bar">' +
    '<span id="countLabel"></span>' +
    '<button class="btn-close" onclick="google.script.host.close()">Close</button>' +
  '</div>' +
  '<script>' +
    'var registry = ' + registryJson + ';' +
    'var enabled = ' + (enabled ? 'true' : 'false') + ';' +
    'var totalCount = 0;' +
    'Object.keys(registry).forEach(function(k) { totalCount += registry[k].length; });' +
    'function render(filter) {' +
      'var q = (filter || "").toLowerCase();' +
      'var container = document.getElementById("modalList");' +
      'var html = "";' +
      'var shown = 0;' +
      'Object.keys(registry).forEach(function(cat) {' +
        'var items = registry[cat].filter(function(m) {' +
          'return !q || m.label.toLowerCase().indexOf(q) !== -1 || m.desc.toLowerCase().indexOf(q) !== -1 || cat.toLowerCase().indexOf(q) !== -1;' +
        '});' +
        'if (items.length === 0) return;' +
        'html += \'<div class="category"><div class="category-title">\' + cat + \' (\' + items.length + \')</div><div class="modal-grid">\';' +
        'items.forEach(function(m) {' +
          'var cls = enabled ? "modal-card" : "modal-card disabled";' +
          'html += \'<div class="\' + cls + \'" onclick="launch(\\\'\' + m.fn + \'\\\')" title="\' + m.desc + \'">\' +' +
            '\'<div class="modal-icon">\' + m.icon + \'</div>\' +' +
            '\'<div class="modal-info"><div class="modal-label">\' + m.label + \'</div><div class="modal-desc">\' + m.desc + \'</div></div>\' +' +
          '\'</div>\';' +
          'shown++;' +
        '});' +
        'html += \'</div></div>\';' +
      '});' +
      'if (shown === 0) html = \'<div class="no-results">No modals match your search</div>\';' +
      'container.innerHTML = html;' +
      'document.getElementById("countLabel").textContent = shown + " of " + totalCount + " modals" + (enabled ? "" : " (disabled)");' +
    '}' +
    'function filterModals(val) { render(val); }' +
    'function toggleMaster(checked) {' +
      'enabled = checked;' +
      'google.script.run.withFailureHandler(function(e) { alert("Error: " + e.message); }).setModalHubEnabled(checked);' +
      'render(document.getElementById("searchBar").value);' +
    '}' +
    'function launch(fnName) {' +
      'if (!enabled) return;' +
      'google.script.host.close();' +
      'google.script.run.withFailureHandler(function(e) { alert("Error launching modal: " + e.message); }).launchModalFromHub(fnName);' +
    '}' +
    'render("");' +
  '</script></body></html>';
}

// ============================================================================
// SIDEBAR
// ============================================================================
