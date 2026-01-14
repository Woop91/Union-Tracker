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
