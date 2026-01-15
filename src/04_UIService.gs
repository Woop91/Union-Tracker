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
 * Called automatically by onOpen trigger
 */
function createDashboardMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🏛️ Union Dashboard')
    .addItem('📋 Dashboard Home', 'showDashboardSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 Search')
      .addItem('Desktop Search', 'showDesktopSearch')
      .addItem('Quick Search (Ctrl+Shift+F)', 'showQuickSearch')
      .addItem('Advanced Search', 'showAdvancedSearch'))
    .addSubMenu(ui.createMenu('📁 Grievances')
      .addItem('New Grievance', 'showNewGrievanceDialog')
      .addItem('Edit Selected', 'showEditGrievanceDialog')
      .addItem('Bulk Update Status', 'showBulkStatusUpdate'))
    .addSubMenu(ui.createMenu('👥 Members')
      .addItem('Add New Member', 'showNewMemberDialog')
      .addItem('Import Members', 'showImportDialog')
      .addItem('Export Directory', 'showExportDialog'))
    .addSubMenu(ui.createMenu('📅 Calendar')
      .addItem('Sync Deadlines', 'showCalendarSyncDialog')
      .addItem('View Upcoming', 'showUpcomingDeadlines')
      .addItem('Clear Calendar Events', 'showClearCalendarConfirm'))
    .addSubMenu(ui.createMenu('⚙️ Admin Tools')
      .addItem('System Diagnostics', 'showDiagnosticsDialog')
      .addItem('Repair Dashboard', 'showRepairDialog')
      .addItem('Settings', 'showSettingsDialog'))
    .addSeparator()
    .addItem('❓ Help & Documentation', 'showHelpDialog')
    .addToUi();
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
    .setTitle(SIDEBAR_CONFIG.TITLE);

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

/**
 * Navigates to a specific sheet
 * @param {string} sheetName - Name of sheet to navigate to
 */
function navigateToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

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
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var bandings = range.getBandings();
  if (bandings.length > 0) bandings[0].remove();
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

// ==================== QUICK CAPTURE & TIMER ====================

function getQuickCaptureNotes() {
  return PropertiesService.getUserProperties().getProperty('quickCaptureNotes') || '';
}

function saveQuickCaptureNotes(notes) {
  PropertiesService.getUserProperties().setProperty('quickCaptureNotes', notes);
}

function showQuickCaptureNotepad() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}textarea{width:100%;height:300px;padding:10px;border:2px solid #ddd;border-radius:4px}button{background:#1a73e8;color:white;border:none;padding:12px 24px;border-radius:4px;cursor:pointer;margin:10px 5px 0 0}</style></head><body><h3>📝 Quick Capture</h3><textarea id="notes" placeholder="Type your thoughts..."></textarea><br><button onclick="save()">💾 Save</button><button onclick="google.script.host.close()">Close</button><script>google.script.run.withSuccessHandler(function(n){document.getElementById("notes").value=n||""}).getQuickCaptureNotes();function save(){google.script.run.saveQuickCaptureNotes(document.getElementById("notes").value);alert("Saved!")}</script></body></html>'
  ).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '📝 Quick Capture');
}

function startPomodoroTimer() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:40px;text-align:center;background:#1a73e8;color:white}.timer{font-size:72px;font-weight:bold;margin:40px 0;font-family:monospace}button{background:white;color:#1a73e8;border:none;padding:15px 30px;font-size:16px;border-radius:8px;cursor:pointer;margin:10px}</style></head><body><h2>🍅 Pomodoro Timer</h2><div id="status">Focus Session</div><div class="timer" id="timer">25:00</div><button onclick="toggle()">▶️ Start</button><button onclick="google.script.host.close()">Close</button><script>var left=25*60,running=false,iv;function toggle(){if(running){clearInterval(iv);running=false}else{running=true;iv=setInterval(function(){if(left>0){left--;var m=Math.floor(left/60),s=left%60;document.getElementById("timer").textContent=(m<10?"0":"")+m+":"+(s<10?"0":"")+s}else{clearInterval(iv);alert("Session complete!")}},1000)}}</script></body></html>'
  ).setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModelessDialog(html, '🍅 Pomodoro Timer');
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

function resetToDefaultTheme() {
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
    var isActive = status && status !== 'Resolved' && status !== 'Withdrawn' && status !== 'Closed';
    if (isActive) stats.activeGrievances++;
    if (status === 'Pending Info') stats.pendingGrievances++;
    // Count overdue for ANY active case (Open, Pending Info, etc.), not just 'Open'
    if (daysTo !== null && daysTo !== '' && daysTo < 0 && isActive) stats.overdueGrievances++;
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
 * Shows the Custom View with tabbed interface
 * Features: Overview, Members, Grievances, and Analytics tabs
 *
 * ⚠️ PROTECTED FUNCTION - DO NOT MODIFY ⚠️
 */
function showInteractiveDashboardTab() {
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
    '.error-state{text-align:center;padding:30px;color:#dc2626;background:#fef2f2;border-radius:8px;margin:10px}' +

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

    // Error handler wrapper
    'function safeRun(fn,fallback){try{fn()}catch(e){console.error(e);if(fallback)fallback(e)}}' +

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
    '  google.script.run' +
    '    .withSuccessHandler(function(data){safeRun(function(){renderOverview(data)},function(){document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Error loading stats</div>"})})'  +
    '    .withFailureHandler(function(e){document.getElementById("overview-stats").innerHTML="<div class=\\"error-state\\">Failed to load: "+e.message+"</div>"})' +
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
    '    var overdue=(data||[]).filter(function(g){return g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overview-overdue").innerHTML="<div style=\\"text-align:center;padding:15px;color:#10B981\\">✅ All cases on track!</div>";return}' +
    '    var html="<div class=\\"chart-container\\" style=\\"border-left:4px solid #dc2626\\"><div class=\\"chart-title\\">⚠️ Overdue Cases ("+overdue.length+")</div>";' +
    '    html+="<div class=\\"list-container\\">";' +
    '    overdue.slice(0,3).forEach(function(g){html+="<div class=\\"list-item\\" onclick=\\"showGrievanceDetail(\'"+g.id+"\')\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+g.id+" - "+g.memberName+"</div><div class=\\"list-item-subtitle\\">"+g.issueType+" • "+g.currentStep+"</div></div><span class=\\"badge badge-overdue\\">Overdue</span></div>"});' +
    '    if(overdue.length>3)html+="<button class=\\"action-btn action-btn-danger\\" style=\\"width:100%;margin-top:8px\\" onclick=\\"switchTab(\'grievances\',document.getElementById(\'tab-grievances\'));setTimeout(function(){filterGrievanceStatus(\'Overdue\',document.querySelector(\'[data-filter=Overdue]\'))},300)\\">View All "+overdue.length+" Overdue Cases</button>";' +
    '    html+="</div></div>";' +
    '    document.getElementById("overview-overdue").innerHTML=html;' +
    '  }).withFailureHandler(function(e){document.getElementById("overview-overdue").innerHTML="<div class=\\"error-state\\">⚠️ Could not load overdue cases</div>"}).getInteractiveGrievanceData();' +
    '}' +

    // Load my cases (steward's assigned grievances)
    'function loadMyCases(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data){myCases=data||[];renderMyCases(myCases);renderMyCasesStats(data)})'  +
    '    .withFailureHandler(function(e){document.getElementById("mycases-list").innerHTML="<div class=\\"error-state\\">Failed to load your cases: "+e.message+"</div>"})' +
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
    '  google.script.run' +
    '    .withSuccessHandler(function(data){allMembers=data||[];renderMembers(allMembers);loadMemberFilters()})'  +
    '    .withFailureHandler(function(e){document.getElementById("members-list").innerHTML="<div class=\\"error-state\\">Failed to load members</div>"})' +
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
    '  google.script.run' +
    '    .withSuccessHandler(function(data){allGrievances=data||[];renderGrievances(allGrievances)})'  +
    '    .withFailureHandler(function(e){document.getElementById("grievances-list").innerHTML="<div class=\\"error-state\\">Failed to load grievances</div>"})' +
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
