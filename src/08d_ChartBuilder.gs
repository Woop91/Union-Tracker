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

  var chartNum = sheet.getRange('G120').getValue();
  if (!chartNum || chartNum < 1 || chartNum > 15) {
    ss.toast('Please enter a valid chart number (1-15) in cell G120', '⚠️ Invalid Selection', 5);
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
        .addRange(sheet.getRange('A15:A16'))
        .addRange(sheet.getRange('A16:F16'))
        .setPosition(135, 1, 0, 0)
        .setOption('title', 'Grievance Status Distribution')
        .setOption('legend', {position: 'bottom'})
        .setOption('colors', [COLORS.SOLIDARITY_RED, COLORS.ACCENT_ORANGE, COLORS.STATUS_BLUE, COLORS.UNION_GREEN, '#9CA3AF', '#6B7280'])
        .setOption('width', 600)
        .setOption('height', 300);
      chart = chartBuilder.build();
      sheet.insertChart(chart);
      break;

    case 3: // Pie Chart - Issue Categories
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(sheet.getRange('A26:B30'))
        .setPosition(135, 1, 0, 0)
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
      chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(sheet.getRange('A35:B39'))
        .setPosition(135, 1, 0, 0)
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
        .addRange(sheet.getRange('A26:D30'))
        .setPosition(135, 1, 0, 0)
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
        .addRange(sheet.getRange('A26:B30'))
        .setPosition(135, 1, 0, 0)
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
        .addRange(sheet.getRange('A35:C39'))
        .setPosition(135, 1, 0, 0)
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
        .addRange(sheet.getRange('B26:B30'))
        .setPosition(135, 1, 0, 0)
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
      ss.toast('Enter 1-15 in cell G120 to select a chart type. See options table above.', 'ℹ️ Chart Help', 5);
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
  var winRate = sheet.getRange('D6').getValue() || '0%';
  var gaugeText = '═══════════════════════════════════════\n' +
                  '           🎯 WIN RATE GAUGE\n' +
                  '═══════════════════════════════════════\n\n' +
                  '                 ' + winRate + '\n\n' +
                  '    ◀━━━━━━━━━━━━━━━━━━━━━━━━━━━▶\n' +
                  '    0%        50%        100%\n\n' +
                  '═══════════════════════════════════════';

  sheet.getRange('A135').setValue(gaugeText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F0FDF4');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates a scorecard-style display
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createScorecardChart_(sheet) {
  var openCases = sheet.getRange('A16').getValue() || 0;
  var scorecardText = '┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓\n' +
                      '┃         📊 OPEN GRIEVANCES          ┃\n' +
                      '┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫\n' +
                      '┃                                     ┃\n' +
                      '┃              ' + openCases + '                    ┃\n' +
                      '┃                                     ┃\n' +
                      '┃           ▲ Active Cases            ┃\n' +
                      '┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛';

  sheet.getRange('A135').setValue(scorecardText)
    .setFontFamily('Courier New')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#FEF3C7');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates a trend line chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createTrendLineChart_(sheet) {
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('A44:C46'))
    .setPosition(135, 1, 0, 0)
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
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(sheet.getRange('A44:C46'))
    .setPosition(135, 1, 0, 0)
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
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(sheet.getRange('A26:C30'))
    .setPosition(135, 1, 0, 0)
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
  var openCases = sheet.getRange('A16').getValue() || 0;
  var resolvedCases = sheet.getRange('D16').getValue() || 0;
  var winRate = sheet.getRange('D6').getValue() || '0%';
  var avgDays = sheet.getRange('E6').getValue() || 'N/A';

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

  sheet.getRange('A135').setValue(tableText)
    .setFontFamily('Courier New')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#F3F4F6');
  sheet.getRange('A135:G145').merge();
}

/**
 * Creates steward leaderboard chart
 * @param {Sheet} sheet - The dashboard sheet
 * @private
 */
function createStewardLeaderboardChart_(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    sheet.getRange('A135').setValue('Member Directory not found');
    return;
  }

  var data = memberSheet.getDataRange().getValues();
  var stewards = [];

  // Find stewards and their case counts
  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      stewards.push({
        name: data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1],
        cases: data[i][MEMBER_COLS.TOTAL_CASES - 1] || 0,
        wins: data[i][MEMBER_COLS.WINS - 1] || 0
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

  sheet.getRange('A135').setValue(leaderboardText)
    .setFontFamily('Courier New')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#EFF6FF');
  sheet.getRange('A135:G145').merge();
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
