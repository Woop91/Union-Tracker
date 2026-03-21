/**
 * ============================================================================
 * 08b_SearchAndCharts.gs - Search Engine and Chart Generation
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Search engine for the spreadsheet interface. Provides desktop search
 *   (multi-field: name, email, job title, location, grievance ID, issue type)
 *   and chart generation for analytics dashboards. getDesktopSearchData()
 *   searches both Member Directory and Grievance Log with filters.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Search is the most-used feature for stewards finding members or cases
 *   quickly. Desktop search searches more fields than mobile search (job
 *   title, location, issue type) because desktop users have wider screens to
 *   show results. Results are ranked by relevance.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The Desktop Search dialog returns no results. Stewards must manually scan
 *   sheets to find members/cases. Chart generation for analytics dashboards
 *   fails.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, MEMBER_COLS, GRIEVANCE_COLS). Used by
 *   menu items in 03_, the Steward Dashboard, and search dialogs.
 *
 * @fileoverview Search and chart generation functions
 * @requires 01_Core.gs
 */

// ============================================================================
// SEARCH FUNCTIONS
// ============================================================================

/**
 * Gets locations for desktop search filter dropdown
 * @returns {Array<string>} Array of unique locations sorted alphabetically
 */
function getDesktopSearchLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = [];

  var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (mSheet && mSheet.getLastRow() > 1) {
    var mData = mSheet.getRange(2, MEMBER_COLS.WORK_LOCATION, mSheet.getLastRow() - 1, 1).getValues();
    mData.forEach(function(row) {
      var loc = row[0];
      if (loc && locations.indexOf(loc) === -1) {
        locations.push(loc);
      }
    });
  }

  return locations.sort();
}

/**
 * Gets search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array<Object>} Array of search results
 */
function getDesktopSearchData(query, tab, filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = (query || '').toLowerCase();
  filters = filters || {};

  // Search Members
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var lastCol = Math.max(MEMBER_COLS.IS_STEWARD, MEMBER_COLS.WORK_LOCATION, MEMBER_COLS.JOB_TITLE, MEMBER_COLS.EMAIL);
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, lastCol).getValues();

      mData.forEach(function(row, index) {
        var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var firstName = row[MEMBER_COLS.FIRST_NAME - 1] || '';
        var lastName = row[MEMBER_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
        var jobTitle = row[MEMBER_COLS.JOB_TITLE - 1] || '';
        var location = row[MEMBER_COLS.WORK_LOCATION - 1] || '';
        var isSteward = row[MEMBER_COLS.IS_STEWARD - 1] || '';

        // Apply filters
        if (filters.location && location !== filters.location) return;
        if (filters.isSteward && isSteward !== filters.isSteward) return;

        // Search across fields
        var searchable = (memberId + ' ' + fullName + ' ' + email + ' ' + jobTitle + ' ' + location).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.location && !filters.isSteward) return;

        results.push({
          type: 'member',
          id: memberId,
          title: fullName.trim() || 'Unnamed Member',
          email: email,
          jobTitle: jobTitle,
          location: location,
          isSteward: isSteward,
          row: index + 2
        });
      });
    }
  }

  // Search Grievances
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var lastGCol = Math.max(GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.ISSUE_CATEGORY, GRIEVANCE_COLS.LOCATION, GRIEVANCE_COLS.STEWARD, GRIEVANCE_COLS.DATE_FILED);
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, lastGCol).getValues();

      gData.forEach(function(row, index) {
        var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
        var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
        var issueType = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '';
        var location = row[GRIEVANCE_COLS.LOCATION - 1] || '';
        var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
        var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1] || '';

        // Apply filters
        if (filters.status && status !== filters.status) return;
        if (filters.location && location !== filters.location) return;

        // Search across fields
        var searchable = (grievanceId + ' ' + fullName + ' ' + status + ' ' + issueType + ' ' + location + ' ' + steward).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.status && !filters.location) return;

        // Format date
        var filedDateStr = '';
        if (dateFiled) {
          try {
            filedDateStr = Utilities.formatDate(new Date(dateFiled), Session.getScriptTimeZone(), 'MM/dd/yyyy');
          } catch(_e) {
            filedDateStr = dateFiled.toString();
          }
        }

        results.push({
          type: 'grievance',
          id: grievanceId,
          title: fullName.trim() || 'Unknown Member',
          status: status,
          issueType: issueType,
          location: location,
          steward: steward,
          filedDate: filedDateStr,
          row: index + 2
        });
      });
    }
  }

  return results.slice(0, 50);
}

/**
 * Navigates to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
 * @returns {void}
 */
function navigateToSearchResult(type, id, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = type === 'member' ? SHEETS.MEMBER_DIR : SHEETS.GRIEVANCE_LOG;
  var sheet = ss.getSheetByName(sheetName);

  if (sheet && row) {
    ss.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(row, 1));
    SpreadsheetApp.flush();
  }
}

/**
 * Navigates to active grievances view
 * @returns {void}
 */
function viewActiveGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

/**
 * Searches the dashboard for matching records
 * @param {string} query - Search query
 * @param {string} searchType - 'all', 'members', or 'grievances'
 * @param {Object} filters - Additional filters
 * @returns {Array<Object>} Search results
 */
function searchDashboard(query, searchType, filters) {
  var results = [];
  var queryLower = query.toLowerCase();

  // Search members
  if (searchType === 'all' || searchType === 'members') {
    var members = getMemberList();
    members.forEach(function(m) {
      if (m.name.toLowerCase().indexOf(queryLower) !== -1 ||
          m.id.toLowerCase().indexOf(queryLower) !== -1 ||
          m.department.toLowerCase().indexOf(queryLower) !== -1) {
        results.push({
          id: m.id,
          type: 'member',
          title: m.name,
          subtitle: m.department + ' - ID: ' + m.id
        });
      }
    });
  }

  // Search grievances
  if (searchType === 'all' || searchType === 'grievances') {
    var grievances = getOpenGrievances();
    grievances.forEach(function(g) {
      var grievanceId = g['Grievance ID'] || '';
      var memberName = g['Member Name'] || '';
      var description = g['Description'] || '';
      var status = g['Status'] || '';

      if (grievanceId.toLowerCase().indexOf(queryLower) !== -1 ||
          memberName.toLowerCase().indexOf(queryLower) !== -1 ||
          description.toLowerCase().indexOf(queryLower) !== -1) {

        // Apply status filter
        if (filters.status && status.toLowerCase() !== filters.status.toLowerCase()) {
          return;
        }

        results.push({
          id: grievanceId,
          type: 'grievance',
          title: grievanceId + ' - ' + memberName,
          subtitle: status + ' - Step ' + (g['Current Step'] || 1)
        });
      }
    });
  }

  return results.slice(0, 50);
}

/**
 * Quick search for instant results
 * @param {string} query - Search query
 * @returns {Array<Object>} Quick search results (max 10)
 */
function quickSearchDashboard(query) {
  if (!query || query.length < 2) return [];

  var results = searchDashboard(query, 'all', {});
  return results.slice(0, 10);
}

/**
 * Advanced search with complex filters
 * @param {Object} filters - Search filters
 * @returns {Array<Object>} Search results
 */
function advancedSearch(filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search members if included
  if (filters.includeMembers) {
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIRECTORY);
    if (memberSheet) {
      var data = memberSheet.getDataRange().getValues();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var matches = true;

        // Apply department filter
        if (filters.department && row[MEMBER_COLS.JOB_TITLE - 1] !== filters.department) {
          matches = false;
        }

        // Apply name filter
        if (filters.name && matches) {
          var fullName = (row[MEMBER_COLS.FIRST_NAME - 1] + ' ' + row[MEMBER_COLS.LAST_NAME - 1]).toLowerCase();
          if (fullName.indexOf(filters.name.toLowerCase()) === -1) {
            matches = false;
          }
        }

        // Apply steward filter
        if (filters.stewardOnly && matches) {
          if (!isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1])) {
            matches = false;
          }
        }

        if (matches && row[MEMBER_COLS.MEMBER_ID - 1]) {
          results.push({
            id: row[MEMBER_COLS.MEMBER_ID - 1],
            type: 'member',
            title: row[MEMBER_COLS.FIRST_NAME - 1] + ' ' + row[MEMBER_COLS.LAST_NAME - 1],
            subtitle: row[MEMBER_COLS.JOB_TITLE - 1],
            row: i + 1
          });
        }
      }
    }
  }

  // Search grievances if included
  if (filters.includeGrievances) {
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet) {
      var gData = grievanceSheet.getDataRange().getValues();

      for (var j = 1; j < gData.length; j++) {
        var gRow = gData[j];
        var gMatches = true;

        // Apply status filter
        if (filters.status && gRow[GRIEVANCE_COLS.STATUS - 1] !== filters.status) {
          gMatches = false;
        }

        // Apply date range filter
        if (filters.startDate && gMatches) {
          var filedDate = gRow[GRIEVANCE_COLS.DATE_FILED - 1];
          if (filedDate && new Date(filedDate) < new Date(filters.startDate)) {
            gMatches = false;
          }
        }

        if (filters.endDate && gMatches) {
          var endFiledDate = gRow[GRIEVANCE_COLS.DATE_FILED - 1];
          if (endFiledDate && new Date(endFiledDate) > new Date(filters.endDate)) {
            gMatches = false;
          }
        }

        if (gMatches && gRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1]) {
          results.push({
            id: gRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
            type: 'grievance',
            title: gRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
            subtitle: gRow[GRIEVANCE_COLS.STATUS - 1] + ' - ' + gRow[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
            row: j + 1
          });
        }
      }
    }
  }

  return results.slice(0, 100);
}

/**
 * Gets department list from calc sheet or direct query
 * @returns {Array<string>} Array of department names
 */
function getDepartmentList() {
  var formulaSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HIDDEN_SHEETS.CALC_FORMULAS);

  if (!formulaSheet) {
    var memberSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(SHEETS.MEMBER_DIRECTORY);

    if (!memberSheet || memberSheet.getLastRow() <= 1) return [];

    var data = memberSheet.getRange(2, MEMBER_COLS.JOB_TITLE,
      memberSheet.getLastRow() - 1, 1).getValues();

    var depts = {};
    data.forEach(function(row) {
      if (row[0]) depts[row[0]] = true;
    });

    return Object.keys(depts).sort();
  }

  var deptData = formulaSheet.getRange('B4:B').getValues();
  return deptData.filter(function(row) { return row[0]; }).map(function(row) { return row[0]; });
}

/**
 * Gets member list for dropdowns
 * @returns {Array<Object>} Array of member objects with id, name, department
 */
function getMemberList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIRECTORY);

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1]) {
      members.push({
        id: data[i][MEMBER_COLS.MEMBER_ID - 1],
        name: data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1],
        department: data[i][MEMBER_COLS.JOB_TITLE - 1]
      });
    }
  }

  return members;
}
/**
 * ============================================================================
 * ChartBuilder.gs - Dashboard Chart Generation
 * ============================================================================
 *
 * This module handles all chart-related functions including:
 * - Chart options display
 * - Chart generation (bar, pie, line, gauge, etc.)
 * - Chart styling and customization
 *
 * REFACTORED: Split from 08_Code.gs for better maintainability
 *
 * @fileoverview Chart generation and display functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// CHART GENERATION
// ============================================================================

/**
 * Generates the selected chart based on user's choice in cell G120
 * Call this from the menu after user selects a chart number
 */
function generateSelectedChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!sheet) {
    ss.toast('Dashboard sheet not found', '❌ Error', 3);
    return;
  }

  var L = DASHBOARD_LAYOUT;
  var chartNum = sheet.getRange(L.CHART_INPUT_CELL).getValue();
  if (!chartNum || chartNum < 1 || chartNum > 15) {
    ss.toast('Please enter a valid chart number (1-15) in cell ' + L.CHART_INPUT_CELL, '⚠️ Invalid Selection', 5);
    return;
  }

  ss.toast('Generating chart #' + chartNum + '...', '📊 Creating Chart', 2);

  // Remove existing charts in the display area
  var existingCharts = sheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    sheet.removeChart(existingCharts[i]);
  }

  var chart;
  var chartBuilder;

  switch (parseInt(chartNum)) {
    case 1: // Gauge Chart - Win Rate
      createGaugeStyleChart_(sheet);
      break;

    case 2: // Bar Chart - Status Distribution
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange(L.GRIEVANCE_METRICS_ROW - 1, 1, 1, L.DATA_COLS))
        .addRange(sheet.getRange(L.GRIEVANCE_METRICS_ROW, 1, 1, L.DATA_COLS))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Grievance Status Distribution')
        .setOption('legend', {position: 'bottom'})
        .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.ACCENT_ORANGE, COLORS.STATUS_BLUE, COLORS.UNION_GREEN, '#9CA3AF', '#6B7280'])
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 3: // Pie Chart - Issue Categories
      var catRows = L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1;
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange(L.CATEGORY_START_ROW, 1, catRows, 2))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Grievances by Issue Category')
        .setOption('pieHole', 0)
        .setOption('legend', {position: 'right'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 4: // Line Chart - Trends
      createTrendLineChart_(sheet);
      break;

    case 5: // Column Chart - Location
      var locRows = L.LOCATION_END_ROW - L.LOCATION_START_ROW + 1;
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(sheet.getRange(L.LOCATION_START_ROW, 1, locRows, 2))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Members by Location')
        .setOption('legend', {position: 'none'})
        .setOption('colors', [COLORS.CHART_CYAN])
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 6: // Scorecard
      createScorecardChart_(sheet);
      break;

    case 7: // Heatmap Table
      ss.toast('Heatmap applied! Check "Apply Gradient Heatmaps" in Admin menu for color scales.', '🔥 Heatmap', 5);
      if (typeof applyWinRateGradients === 'function') {
        applyWinRateGradients();
      }
      break;

    case 8: // Stacked Bar
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(sheet.getRange(L.CATEGORY_START_ROW, 1, L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1, 4))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Cases by Issue Category (Open vs Resolved)')
        .setOption('isStacked', true)
        .setOption('legend', {position: 'bottom'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 9: // Donut Chart
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange(L.CATEGORY_START_ROW, 1, L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1, 2))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Issue Category Distribution')
        .setOption('pieHole', 0.4)
        .setOption('legend', {position: 'right'})
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 10: // Area Chart
      createAreaChart_(sheet);
      break;

    case 11: // Combo Chart
      createComboChart_(sheet);
      break;

    case 12: // Scatter Plot
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(sheet.getRange(L.LOCATION_START_ROW, 1, L.LOCATION_END_ROW - L.LOCATION_START_ROW + 1, 3))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Response Time vs Outcome Analysis')
        .setOption('legend', {position: 'bottom'})
        .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.UNION_GREEN])
        .setOption('width', 600)
        .setOption('height', 300)
        .setOption('pointSize', 8);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 13: // Histogram
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.HISTOGRAM)
        .addRange(sheet.getRange(L.CATEGORY_START_ROW, 2, L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1, 1))
        .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
        .setOption('title', 'Case Duration Distribution')
        .setOption('legend', {position: 'none'})
        .setOption('colors', [COLORS.CHART_CYAN])
        .setOption('width', 600)
        .setOption('height', 300)
        .setOption('histogram', {bucketSize: 5});
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 14: // Summary Table
      createSummaryTableChart_(sheet);
      break;

    case 15: // Steward Leaderboard
      createStewardLeaderboardChart_(sheet);
      break;

    default:
      ss.toast('Enter 1-15 in cell ' + L.CHART_INPUT_CELL + ' to select a chart type. See options table above.', 'ℹ️ Chart Help', 5);
  }

  ss.toast('Chart generated! Scroll down to "Chart Display Area" to view.', '✅ Done', 5);
}

// ============================================================================
// CHART HELPER FUNCTIONS
// ============================================================================

/**
 * Creates a gauge-style display (Google Sheets doesn't have native gauge)
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createGaugeStyleChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var winRate = sheet.getRange(L.QUICK_STATS_ROW, 4).getValue() || '0%';
  var gaugeText = '═══════════════════════════════════════\n' +
                  '           🎯 WIN RATE GAUGE\n' +
                  '═══════════════════════════════════════\n\n' +
                  '                 ' + winRate + '\n\n' +
                  '    ◀━━━━━━━━━━━━━━━━━━━━━━━━━━━▶\n' +
                  '    0%        50%        100%\n\n' +
                  '═══════════════════════════════════════';

  sheet.getRange(L.CHART_DISPLAY_ROW, 1).setValue(gaugeText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F0FDF4');
  sheet.getRange(L.CHART_DISPLAY_RANGE).merge();
}

/**
 * Creates a scorecard-style display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createScorecardChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var openCases = sheet.getRange(L.GRIEVANCE_METRICS_ROW, 1).getValue() || 0;
  var scorecardText = '┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓\n' +
                      '┃         📊 OPEN GRIEVANCES          ┃\n' +
                      '┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫\n' +
                      '┃                                     ┃\n' +
                      '┃              ' + openCases + '                    ┃\n' +
                      '┃                                     ┃\n' +
                      '┃           ▲ Active Cases            ┃\n' +
                      '┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛';

  sheet.getRange(L.CHART_DISPLAY_ROW, 1).setValue(scorecardText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FEF3C7');
  sheet.getRange(L.CHART_DISPLAY_RANGE).merge();
}

/**
 * Creates a trend line chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createTrendLineChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var trendRows = L.TREND_END_ROW - L.TREND_START_ROW + 1;
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(L.TREND_START_ROW, 1, trendRows, 3))
    .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
    .setOption('title', 'Month-Over-Month Trends')
    .setOption('legend', {position: 'bottom'})
    .setOption('curveType', 'function')
    .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.CHART_BLUE])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates an area chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createAreaChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var trendRows = L.TREND_END_ROW - L.TREND_START_ROW + 1;
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(sheet.getRange(L.TREND_START_ROW, 1, trendRows, 3))
    .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
    .setOption('title', 'Cumulative Trend Analysis')
    .setOption('legend', {position: 'bottom'})
    .setOption('colors', [COLORS.CHART_PURPLE, COLORS.CHART_BLUE])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates a combo chart (bars with line overlay)
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createComboChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange(L.CATEGORY_START_ROW, 1, L.CATEGORY_END_ROW - L.CATEGORY_START_ROW + 1, 3))
    .setPosition(L.CHART_DISPLAY_ROW, 1, 0, 0)
    .setOption('title', 'Cases by Category with Trend Line')
    .setOption('legend', {position: 'bottom'})
    .setOption('seriesType', 'bars')
    .setOption('series', {1: {type: 'line'}})
    .setOption('colors', [COLORS.CHART_BLUE, COLORS.SOLIDARITY_RED])
    .setOption('width', 600)
    .setOption('height', 300);

  var chart = chartBuilder.build();
  sheet.insertChart(chart);
}

/**
 * Creates a summary table display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createSummaryTableChart_(sheet) {
  var L = DASHBOARD_LAYOUT;
  var openCases = sheet.getRange(L.GRIEVANCE_METRICS_ROW, 1).getValue() || 0;
  var resolvedCases = sheet.getRange(L.GRIEVANCE_METRICS_ROW, 4).getValue() || 0;
  var winRate = sheet.getRange(L.QUICK_STATS_ROW, 4).getValue() || '0%';
  var avgDays = sheet.getRange(L.QUICK_STATS_ROW, 5).getValue() || 'N/A';

  var tableText = '╔═══════════════════════════════════════════════════════╗\n' +
                  '║            📋 KPI SUMMARY TABLE                       ║\n' +
                  '╠═══════════════════════════════════════════════════════╣\n' +
                  '║  Metric                    │  Value                   ║\n' +
                  '╠═══════════════════════════════════════════════════════╣\n' +
                  '║  Open Cases                │  ' + padRight(String(openCases), 23) + '║\n' +
                  '║  Resolved Cases            │  ' + padRight(String(resolvedCases), 23) + '║\n' +
                  '║  Win Rate                  │  ' + padRight(String(winRate), 23) + '║\n' +
                  '║  Avg Resolution Time       │  ' + padRight(String(avgDays), 23) + '║\n' +
                  '╚═══════════════════════════════════════════════════════╝';

  sheet.getRange(L.CHART_DISPLAY_ROW, 1).setValue(tableText)
    .setFontFamily('Courier New')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F3F4F6');
  sheet.getRange(L.CHART_DISPLAY_RANGE).merge();
}

/**
 * Creates steward leaderboard chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createStewardLeaderboardChart_(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var L = DASHBOARD_LAYOUT;
  if (!memberSheet) {
    sheet.getRange(L.CHART_DISPLAY_ROW, 1).setValue('Member Directory not found');
    return;
  }

  var data = memberSheet.getDataRange().getValues();

  // Build case counts from Grievance Log since MEMBER_COLS doesn't have TOTAL_CASES/WINS
  var caseCounts = {};
  var winCounts = {};
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var gData = grievanceSheet.getDataRange().getValues();
    for (var g = 1; g < gData.length; g++) {
      var stewardName = gData[g][GRIEVANCE_COLS.STEWARD - 1] || '';
      if (stewardName) {
        caseCounts[stewardName] = (caseCounts[stewardName] || 0) + 1;
        var gStatus = gData[g][GRIEVANCE_COLS.STATUS - 1] || '';
        if (gStatus === GRIEVANCE_STATUS.WON || gStatus === GRIEVANCE_STATUS.SETTLED) {
          winCounts[stewardName] = (winCounts[stewardName] || 0) + 1;
        }
      }
    }
  }

  var stewards = [];

  // Find stewards and look up their case counts
  for (var i = 1; i < data.length; i++) {
    if (isTruthyValue(data[i][MEMBER_COLS.IS_STEWARD - 1])) {
      var fullName = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      stewards.push({
        name: fullName,
        cases: caseCounts[fullName] || 0,
        wins: winCounts[fullName] || 0
      });
    }
  }

  // Sort by cases descending
  stewards.sort(function(a, b) { return b.cases - a.cases; });

  var leaderboardText = '╔═══════════════════════════════════════════════════════╗\n' +
                        '║           🏆 STEWARD LEADERBOARD                      ║\n' +
                        '╠═══════════════════════════════════════════════════════╣\n' +
                        '║  Rank │ Name                      │ Cases │ Wins     ║\n' +
                        '╠═══════════════════════════════════════════════════════╣\n';

  var medals = ['🥇', '🥈', '🥉'];
  for (var j = 0; j < Math.min(5, stewards.length); j++) {
    var rank = j < 3 ? medals[j] : (j + 1) + '.';
    leaderboardText += '║  ' + padRight(rank, 4) + ' │ ' +
                       padRight(stewards[j].name, 25) + ' │ ' +
                       padRight(String(stewards[j].cases), 5) + ' │ ' +
                       padRight(String(stewards[j].wins), 8) + '║\n';
  }

  leaderboardText += '╚═══════════════════════════════════════════════════════╝';

  sheet.getRange(L.CHART_DISPLAY_ROW, 1).setValue(leaderboardText)
    .setFontFamily('Courier New')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#EFF6FF');
  sheet.getRange(L.CHART_DISPLAY_RANGE).merge();
}

/**
 * Pads a string to a specified length
 * @param {string} str - String to pad
 * @param {number} len - Target length
 * @returns {string} Padded string
 */
function padRight(str, len) {
  str = String(str);
  while (str.length < len) {
    str += ' ';
  }
  return str.substring(0, len);
}

/**
 * ============================================================================
 * 08e_FormHandlers.gs - Form Handlers Module
 * ============================================================================
 *
 * This module contains all form-related functions for the Steward Dashboard:
 * - Form submission handlers (Contact, Satisfaction, Grievance)
 * - Form trigger setup functions
 * - Form URL handling and configuration
 * - Form value parsing utilities
 *
 * Dependencies:
 * - SHEETS, CONFIG_COLS constants from 00_Constants.gs
 * - GRIEVANCE_FORM_CONFIG, CONTACT_FORM_CONFIG, SATISFACTION_FORM_CONFIG
 * - MEMBER_COLS, SATISFACTION_COLS from column configuration
 * - Helper functions: findExistingMember, generateNameBasedId, validateMemberEmail
 * - Helper functions: getCurrentQuarter, computeSatisfactionRowAverages, syncSatisfactionValues
 *
 * Note: onGrievanceFormSubmit() is defined in 05_Integrations.gs
 *
 * @author SEIU Local Development Team
 * @version 1.0.0
 */
