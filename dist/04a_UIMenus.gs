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
 * @version 4.7.0
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

        function goToSheet(sheetName) {
          google.script.run.withFailureHandler(function(){}).navigateToSheet(sheetName);
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

/**
 * Applies a dashboard theme
 * @param {string} theme - Theme name (default, dark, light, contrast)
 */

/**
 * Helper functions for navigation
 * Note: Dashboards are now modal-based for better UX
 */
/** @deprecated Use showStewardDashboard() directly. Kept for menu backward compat. */
function showExecutiveDashboard() {
  showStewardDashboard();
}

function showStewardDirectory() {
  // Navigate to Member Directory filtered by stewards
  navigateToSheet(SHEETS.MEMBER_DIR);
  SpreadsheetApp.getActive().toast('Filter by "Is Steward = Yes" to see steward directory', 'Steward Directory', 5);
}

// ============================================================================
// NAVIGATION HELPERS (Strategic Command Center)
// ============================================================================

/**
 * Navigate to Executive Dashboard (now launches modal)
 */

/**
 * Navigate to Mobile View (if exists)
 * NOTE: Duplicate exists in 10_CommandCenter.gs - this version kept for compatibility
 */

/**
 * Navigate to a specific sheet by name
 * @param {string} sheetName - The sheet name to navigate to
 */

// ============================================================================
// GLOBAL STYLING (Strategic Command Center)
// ============================================================================

/**
 * Applies the system theme to all visible sheets
 * Includes header styling, zebra striping, and font standardization
 */

/**
 * Applies theme styling to a single sheet
 * For data sheets (Member Directory, Grievance Log), applies to ALL rows in the sheet
 * @param {Sheet} sheet - The sheet to style
 * @private
 */

/**
 * Applies global styling to all visible sheets (alias for APPLY_SYSTEM_THEME)
 */

/**
 * Resets all visible sheets to default styling
 */

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

/**
 * Refreshes all visual elements, data calculations, and alerts
 * Main entry point for Force Global Refresh menu item
 */

// ============================================================================
// SEARCH DIALOGS
// ============================================================================

/**
 * Shows the desktop search dialog with full functionality
 * Includes member search, grievance search, and filters
 */

/**
 * Shows quick search dialog (minimal interface)
 */

/**
 * Shows advanced search with complex filtering options
 */

/**
 * Generates HTML for desktop search dialog
 * @return {string} HTML content
 */

/**
 * Generates HTML for quick search dialog
 * @return {string} HTML content
 */

/**
 * Generates HTML for advanced search
 * @return {string} HTML content
 */

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
        <a href="#" class="quick-link" onclick="goToSheet('${SHEETS.GRIEVANCE_LOG}')">
          📋 Grievance Log
        </a>
        <a href="#" class="quick-link" onclick="goToSheet('${SHEETS.MEMBER_DIR}')">
          👥 Member Directory
        </a>
        <a href="#" class="quick-link" onclick="goToSheet('${SHEETS.DASHBOARD}')">
          📊 Dashboard
        </a>
        <a href="#" class="quick-link" onclick="showSearch()">
          🔍 Search Dashboard
        </a>

        <button class="btn btn-secondary refresh-btn" onclick="refreshData()">
          🔄 Refresh Data
        </button>
      </div>

      <script>
        ${getClientSideEscapeHtml()}
        function loadStats() {
          google.script.run
            .withSuccessHandler(function(stats) {
              if (typeof stats === 'string') { try { stats = JSON.parse(stats); } catch(e) { return; } }
              document.getElementById('openCount').textContent = stats.open;
              document.getElementById('pendingCount').textContent = stats.pending;
              document.getElementById('memberCount').textContent = stats.members;
              document.getElementById('resolvedCount').textContent = stats.resolved;
            })
            .withFailureHandler(function(){document.getElementById("openCount").textContent="--";document.getElementById("pendingCount").textContent="--";document.getElementById("memberCount").textContent="--";document.getElementById("resolvedCount").textContent="--";})
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
                       '<div>' + escapeHtml(d.grievanceId) + '<br><small>' + escapeHtml(d.step) + '</small></div>' +
                       '<div style="text-align:right">' + d.daysLeft + ' days<br><small>' + escapeHtml(d.date) + '</small></div>' +
                       '</div>';
              }).join('');
            })
            .withFailureHandler(function(){document.getElementById("deadlineList").innerHTML='<div style="text-align:center;color:#666;padding:20px;">Could not load deadlines</div>';})
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

