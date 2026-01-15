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

/**
 * Jump to a specific sheet via code
 * @param {string} sheetName - Name of the sheet to navigate to
 */
function navigateToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.setActiveSheet(sheet);
}

// ============================================================================
// 2. EXECUTIVE DASHBOARD (INTERNAL - PII ENABLED)
// ============================================================================

/**
 * Rebuilds the Executive Command Dashboard with PII data
 * Shows sensitive internal metrics for leadership review
 */
function rebuildExecutiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName("Executive Command") || ss.insertSheet("Executive Command");

  dash.clear();

  // Header / Branding
  dash.getRange("B1:L1").merge().setValue("INTERNAL EXECUTIVE COMMAND CENTER")
      .setBackground("#111827").setFontColor("white").setFontSize(16).setHorizontalAlignment("center");

  // Peak Performance: Build KPIs in memory
  renderKPISection(dash);

  // Steward Load Section (PII Included)
  renderStewardList(dash, "B12");

  // Side-by-side Insight Panel (Grievance Breakdown)
  createStewardInsightPanel(dash, "F12");

  dash.getRange("B11").setValue("CONFIDENTIAL: CONTAINS MEMBER PII").setFontColor("#DC2626").setFontWeight("bold");
  ss.toast("Executive Command Loaded", "Success");
  navigateToSheet("Executive Command");
}

/**
 * Renders the KPI section with visual cards
 * @param {Sheet} sheet - The dashboard sheet
 */
function renderKPISection(sheet) {
  sheet.setRowHeight(2, 120);

  // Get live data from grievance calculations
  var metrics = getExecutiveMetrics_();

  // Setup KPI Cards (Active, Win Rate, Overdue)
  setupKPICard(sheet, "B2:D2", "ACTIVE GRIEVANCES", String(metrics.activeGrievances), metrics.activeTrend, "#EF4444");
  setupKPICard(sheet, "F2:H2", "UNION WIN RATE", metrics.winRate + "%", metrics.winRateTrend, "#10B981");
  setupKPICard(sheet, "J2:L2", "OVERDUE STEPS", String(metrics.overdueSteps), metrics.overdueTrend, "#F59E0B");
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

/**
 * Sets up a styled KPI card with rich text formatting
 * @param {Sheet} sheet - The dashboard sheet
 * @param {string} range - Cell range for the card (e.g., "B2:D2")
 * @param {string} label - Card label text
 * @param {string} value - Main value to display
 * @param {string} trend - Trend indicator text
 * @param {string} trendColor - Hex color for trend text
 */
function setupKPICard(sheet, range, label, value, trend, trendColor) {
  const r = sheet.getRange(range);
  r.merge().setBackground("#1F2937").setBorder(true, true, true, true, null, null, "#374151", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const rt = SpreadsheetApp.newRichTextValue()
    .setText(label + "\n" + value + "\n" + trend)
    .setTextStyle(0, label.length, SpreadsheetApp.newTextStyle().setForegroundColor("#9CA3AF").setFontSize(10).setBold(true).build())
    .setTextStyle(label.length + 1, label.length + 1 + value.length, SpreadsheetApp.newTextStyle().setForegroundColor("#FFFFFF").setFontSize(28).setBold(true).build())
    .setTextStyle(label.length + value.length + 2, (label.length + value.length + trend.length + 2), SpreadsheetApp.newTextStyle().setForegroundColor(trendColor).setFontSize(10).build())
    .build();

  r.setRichTextValue(rt).setVerticalAlignment("middle").setHorizontalAlignment("center");
}

/**
 * Renders the steward workload list with PII
 * @param {Sheet} sheet - The dashboard sheet
 * @param {string} startCell - Starting cell for the list
 */
function renderStewardList(sheet, startCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return;

  // Get steward data
  var memberData = memberSheet.getDataRange().getValues();
  var grievanceData = grievanceSheet.getDataRange().getValues();

  // Count active grievances per steward
  var stewardCases = {};
  for (var g = 1; g < grievanceData.length; g++) {
    var steward = grievanceData[g][GRIEVANCE_COLS.STEWARD - 1];
    var status = grievanceData[g][GRIEVANCE_COLS.STATUS - 1];
    if (steward && (status === 'Open' || status === 'Pending Info')) {
      stewardCases[steward] = (stewardCases[steward] || 0) + 1;
    }
  }

  // Build steward list
  var stewards = [];
  for (var m = 1; m < memberData.length; m++) {
    if (memberData[m][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      var name = memberData[m][MEMBER_COLS.FIRST_NAME - 1] + ' ' + memberData[m][MEMBER_COLS.LAST_NAME - 1];
      var email = memberData[m][MEMBER_COLS.EMAIL - 1] || '';
      var location = memberData[m][MEMBER_COLS.WORK_LOCATION - 1] || '';
      var cases = stewardCases[name] || 0;
      stewards.push([name, email, location, cases]);
    }
  }

  // Sort by case count (descending)
  stewards.sort(function(a, b) { return b[3] - a[3]; });

  // Write to sheet
  var startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
  var startCol = startCell.replace(/[0-9]/g, '');

  // Header
  sheet.getRange(startCell).setValue("STEWARD WORKLOAD (PII)")
    .setFontWeight("bold").setBackground("#1E293B").setFontColor("white");
  sheet.getRange(startRow, getColNumber_(startCol), 1, 4).merge();

  // Column headers
  var headerRow = startRow + 1;
  sheet.getRange(headerRow, getColNumber_(startCol), 1, 4)
    .setValues([["Name", "Email", "Location", "Active Cases"]])
    .setFontWeight("bold").setBackground("#374151").setFontColor("white");

  // Data rows (top 10)
  var displayStewards = stewards.slice(0, 10);
  if (displayStewards.length > 0) {
    sheet.getRange(headerRow + 1, getColNumber_(startCol), displayStewards.length, 4)
      .setValues(displayStewards)
      .setBackground("#1F2937").setFontColor("#E5E7EB");
  }
}

/**
 * Creates the insight panel showing grievance breakdown
 * @param {Sheet} sheet - The dashboard sheet
 * @param {string} startCell - Starting cell for the panel
 */
function createStewardInsightPanel(sheet, startCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  var grievanceData = grievanceSheet.getDataRange().getValues();

  // Count by status
  var statusCounts = {};
  var categoryCounts = {};

  for (var g = 1; g < grievanceData.length; g++) {
    var status = grievanceData[g][GRIEVANCE_COLS.STATUS - 1];
    var category = grievanceData[g][GRIEVANCE_COLS.ISSUE_CATEGORY - 1];

    if (status) statusCounts[status] = (statusCounts[status] || 0) + 1;
    if (category) categoryCounts[category] = (categoryCounts[category] || 0) + 1;
  }

  var startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
  var startCol = startCell.replace(/[0-9]/g, '');
  var colNum = getColNumber_(startCol);

  // Panel header
  sheet.getRange(startCell).setValue("GRIEVANCE INSIGHTS")
    .setFontWeight("bold").setBackground("#1E293B").setFontColor("white");
  sheet.getRange(startRow, colNum, 1, 3).merge();

  // Status breakdown
  var row = startRow + 2;
  sheet.getRange(row, colNum).setValue("By Status:").setFontWeight("bold").setFontColor("#9CA3AF");
  row++;

  var statusColors = {
    'Open': '#EF4444',
    'Pending Info': '#F59E0B',
    'Won': '#10B981',
    'Denied': '#DC2626',
    'Settled': '#3B82F6',
    'Withdrawn': '#6B7280'
  };

  for (var status in statusCounts) {
    sheet.getRange(row, colNum).setValue(status);
    sheet.getRange(row, colNum + 1).setValue(statusCounts[status]);
    sheet.getRange(row, colNum + 2).setBackground(statusColors[status] || '#374151');
    row++;
  }

  // Category breakdown
  row++;
  sheet.getRange(row, colNum).setValue("By Category:").setFontWeight("bold").setFontColor("#9CA3AF");
  row++;

  var sortedCategories = Object.entries(categoryCounts).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 5);
  for (var i = 0; i < sortedCategories.length; i++) {
    sheet.getRange(row, colNum).setValue(sortedCategories[i][0]);
    sheet.getRange(row, colNum + 1).setValue(sortedCategories[i][1]);
    row++;
  }
}

/**
 * Helper to convert column letter to number
 * @param {string} col - Column letter (e.g., "B")
 * @returns {number} Column number
 * @private
 */
function getColNumber_(col) {
  var result = 0;
  for (var i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result;
}

// ============================================================================
// 3. MEMBER ANALYTICS DASHBOARD (PUBLIC - NO PII)
// ============================================================================

/**
 * Rebuilds the Member Analytics Dashboard (PII-safe)
 * Shows aggregate metrics without personal information
 */
function rebuildMemberAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName("Member Analytics") || ss.insertSheet("Member Analytics");
  dash.clear();

  // Header
  dash.getRange("B1:L1").merge().setValue("UNION STRENGTH & SENTIMENT REPORT (PII PROTECTED)")
      .setBackground("#1E3A8A").setFontColor("white").setFontSize(16).setHorizontalAlignment("center");

  // Vitals Row (No Names)
  addMemberHappinessGauge(dash, 2, 2); // Top Left
  addParticipationFunnel(dash, 2, 7);   // Top Right

  // Strategic Geography
  renderMembershipHeatmap(dash, "B20");

  // Trend Analysis
  addSentimentTrendChart(dash, "B35");

  ss.toast("Member Analytics Loaded (PII Safe)", "Success");
  navigateToSheet("Member Analytics");
}

/**
 * Adds a morale/happiness gauge chart
 * @param {Sheet} sheet - The dashboard sheet
 * @param {number} row - Row position
 * @param {number} col - Column position
 */
function addMemberHappinessGauge(sheet, row, col) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC) || ss.getSheetByName("_Dashboard_Calc");

  // Create gauge data range if calc sheet exists
  if (calcSheet) {
    try {
      var gauge = sheet.newChart()
        .setChartType(Charts.ChartType.GAUGE)
        .addRange(calcSheet.getRange("AA1:AA2"))
        .setPosition(row, col, 0, 0)
        .setOption('title', 'Morale Score')
        .setOption('max', 5)
        .setOption('greenFrom', 3.75)
        .setOption('greenTo', 5)
        .setOption('yellowFrom', 2.5)
        .setOption('yellowTo', 3.75)
        .setOption('redFrom', 0)
        .setOption('redTo', 2.5)
        .build();
      sheet.insertChart(gauge);
    } catch (e) {
      // Fallback: Show text-based gauge
      sheet.getRange(row, col).setValue("Morale Score").setFontWeight("bold");
      sheet.getRange(row + 1, col).setValue("Data not available").setFontStyle("italic");
    }
  } else {
    // No calc sheet - show placeholder
    sheet.getRange(row, col).setValue("MORALE GAUGE")
      .setFontWeight("bold").setBackground("#10B981").setFontColor("white");
    sheet.getRange(row, col, 1, 3).merge();
    sheet.getRange(row + 1, col).setValue("Configure _Dashboard_Calc sheet to enable")
      .setFontStyle("italic").setFontColor("#6B7280");
    sheet.getRange(row + 1, col, 1, 3).merge();
  }
}

/**
 * Adds a leadership pipeline funnel chart
 * @param {Sheet} sheet - The dashboard sheet
 * @param {number} row - Row position
 * @param {number} col - Column position
 */
function addParticipationFunnel(sheet, row, col) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (memberSheet) {
    var memberData = memberSheet.getDataRange().getValues();

    var totalMembers = 0;
    var stewards = 0;
    var activeEngagement = 0;
    var meetingAttendees = 0;

    for (var m = 1; m < memberData.length; m++) {
      if (memberData[m][MEMBER_COLS.MEMBER_ID - 1]) {
        totalMembers++;
        if (memberData[m][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') stewards++;
        if (memberData[m][MEMBER_COLS.VOLUNTEER_HOURS - 1] > 0) activeEngagement++;
        if (memberData[m][MEMBER_COLS.DUES_PAYING - 1] === 'Yes') meetingAttendees++;
      }
    }

    // Create funnel data
    var funnelData = [
      ["Stage", "Count"],
      ["Total Members", totalMembers],
      ["Dues Paying", meetingAttendees],
      ["Active Engaged", activeEngagement],
      ["Stewards", stewards]
    ];

    // Write funnel data to a temp location
    sheet.getRange(row, col, funnelData.length, 2).setValues(funnelData);

    // Create bar chart as funnel visualization
    var funnel = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(row, col, funnelData.length, 2))
      .setPosition(row + 6, col, 0, 0)
      .setOption('title', 'Leadership Pipeline')
      .setOption('legend', {position: 'none'})
      .setOption('colors', ['#3B82F6'])
      .setOption('width', 350)
      .setOption('height', 200)
      .build();
    sheet.insertChart(funnel);

    // Style the header
    sheet.getRange(row, col).setFontWeight("bold").setBackground("#1E3A8A").setFontColor("white");
  }
}

/**
 * Renders a membership heatmap by location
 * @param {Sheet} sheet - The dashboard sheet
 * @param {string} startCell - Starting cell for the heatmap
 */
function renderMembershipHeatmap(sheet, startCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  var memberData = memberSheet.getDataRange().getValues();

  // Count members by location
  var locationCounts = {};
  for (var m = 1; m < memberData.length; m++) {
    var location = memberData[m][MEMBER_COLS.WORK_LOCATION - 1];
    if (location) {
      locationCounts[location] = (locationCounts[location] || 0) + 1;
    }
  }

  // Sort by count and get top locations
  var sortedLocations = Object.entries(locationCounts)
    .sort(function(a, b) { return b[1] - a[1]; })
    .slice(0, 10);

  var startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
  var startCol = startCell.replace(/[0-9]/g, '');
  var colNum = getColNumber_(startCol);

  // Header
  sheet.getRange(startCell).setValue("MEMBERSHIP BY LOCATION")
    .setFontWeight("bold").setBackground("#1E3A8A").setFontColor("white");
  sheet.getRange(startRow, colNum, 1, 3).merge();

  // Column headers
  sheet.getRange(startRow + 1, colNum, 1, 2)
    .setValues([["Location", "Members"]])
    .setFontWeight("bold").setBackground("#374151").setFontColor("white");

  // Find max for scaling
  var maxCount = sortedLocations.length > 0 ? sortedLocations[0][1] : 1;

  // Data with heatmap coloring
  for (var i = 0; i < sortedLocations.length; i++) {
    var row = startRow + 2 + i;
    sheet.getRange(row, colNum).setValue(sortedLocations[i][0]);
    sheet.getRange(row, colNum + 1).setValue(sortedLocations[i][1]);

    // Calculate color intensity
    var intensity = sortedLocations[i][1] / maxCount;
    var heatColor = intensity > 0.7 ? '#DC2626' : (intensity > 0.4 ? '#F59E0B' : '#10B981');
    sheet.getRange(row, colNum + 2).setBackground(heatColor);
  }
}

/**
 * Adds a sentiment trend chart
 * @param {Sheet} sheet - The dashboard sheet
 * @param {string} startCell - Starting cell for the chart
 */
function addSentimentTrendChart(sheet, startCell) {
  var startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
  var startCol = startCell.replace(/[0-9]/g, '');
  var colNum = getColNumber_(startCol);

  // Header
  sheet.getRange(startCell).setValue("SATISFACTION TRENDS")
    .setFontWeight("bold").setBackground("#1E3A8A").setFontColor("white");
  sheet.getRange(startRow, colNum, 1, 4).merge();

  // Placeholder trend data (would pull from satisfaction survey in production)
  var trendData = [
    ["Month", "Satisfaction", "Engagement", "Confidence"],
    ["Jan", 3.8, 3.5, 3.6],
    ["Feb", 3.9, 3.6, 3.7],
    ["Mar", 4.0, 3.7, 3.8],
    ["Apr", 3.9, 3.8, 3.9],
    ["May", 4.1, 3.9, 4.0],
    ["Jun", 4.2, 4.0, 4.1]
  ];

  sheet.getRange(startRow + 1, colNum, trendData.length, 4).setValues(trendData);

  // Create line chart
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(startRow + 1, colNum, trendData.length, 4))
    .setPosition(startRow + 8, colNum, 0, 0)
    .setOption('title', 'Member Sentiment Over Time')
    .setOption('legend', {position: 'bottom'})
    .setOption('colors', ['#10B981', '#3B82F6', '#F59E0B'])
    .setOption('width', 500)
    .setOption('height', 250)
    .setOption('vAxis', {minValue: 1, maxValue: 5})
    .build();
  sheet.insertChart(chart);
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
 * Renders strategic bargaining data cheat sheet
 */
function renderBargainingCheatSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Member Analytics");

  if (!dash) {
    rebuildMemberAnalytics();
    dash = ss.getSheetByName("Member Analytics");
  }

  // Get grievance data for analysis
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var step1Denials = 0;
  var totalStep1 = 0;
  var settlementDays = [];
  var articleViolations = {};

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getDataRange().getValues();

    for (var g = 1; g < grievanceData.length; g++) {
      // Count Step 1 denials
      if (grievanceData[g][GRIEVANCE_COLS.STEP_1_DATE - 1]) {
        totalStep1++;
        if (grievanceData[g][GRIEVANCE_COLS.STATUS - 1] !== 'Won' &&
            grievanceData[g][GRIEVANCE_COLS.STEP_2_DATE - 1]) {
          step1Denials++;
        }
      }

      // Track settlement time
      var dateFiled = grievanceData[g][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = grievanceData[g][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateFiled instanceof Date && dateClosed instanceof Date) {
        var days = Math.round((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
        settlementDays.push(days);
      }

      // Track contract articles
      var article = grievanceData[g][GRIEVANCE_COLS.CONTRACT_ARTICLE - 1];
      if (article) {
        articleViolations[article] = (articleViolations[article] || 0) + 1;
      }
    }
  }

  // Calculate metrics
  var denialRate = totalStep1 > 0 ? Math.round((step1Denials / totalStep1) * 100) : 0;
  var avgSettlement = settlementDays.length > 0 ?
    Math.round(settlementDays.reduce(function(a, b) { return a + b; }, 0) / settlementDays.length) : 0;

  // Find most violated article
  var topArticle = "N/A";
  var topCount = 0;
  for (var art in articleViolations) {
    if (articleViolations[art] > topCount) {
      topArticle = art;
      topCount = articleViolations[art];
    }
  }

  var data = [
    ["STRATEGIC BARGAINING DATA", "VALUE", "STATUS"],
    ["Step 1 Denials", denialRate + "%", denialRate > 60 ? "High Hostility" : "Normal"],
    ["Avg Settlement Time", avgSettlement + " Days", avgSettlement > 45 ? "Slower" : "Normal"],
    ["Contract Violation Spike", topArticle, topCount > 5 ? "High Risk" : "Monitor"]
  ];

  dash.getRange("H20:J23").setValues(data)
    .setBackground("#FFFBEB")
    .setBorder(true, true, true, true, null, null, "#D97706", SpreadsheetApp.BorderStyle.SOLID);

  // Style header row
  dash.getRange("H20:J20").setFontWeight("bold").setBackground("#FEF3C7");

  navigateToSheet("Member Analytics");
  SpreadsheetApp.getActiveSpreadsheet().toast("Bargaining cheat sheet generated", "Success");
}

/**
 * Generates unit hot zones report showing problem areas
 */
function renderHotZones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Executive Command");

  if (!dash) {
    rebuildExecutiveDashboard();
    dash = ss.getSheetByName("Executive Command");
  }

  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet) return;

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

  // Display hot zones
  var startRow = 25;
  dash.getRange("B25").setValue("HOT ZONES (3+ Active Issues)")
    .setFontWeight("bold").setBackground("#DC2626").setFontColor("white");
  dash.getRange("B25:E25").merge();

  if (hotZones.length === 0) {
    dash.getRange("B26").setValue("No hot zones detected").setFontStyle("italic");
  } else {
    dash.getRange("B26:C26").setValues([["Location", "Active Issues"]])
      .setFontWeight("bold").setBackground("#FEE2E2");

    for (var i = 0; i < hotZones.length; i++) {
      dash.getRange("B" + (27 + i)).setValue(hotZones[i][0]);
      dash.getRange("C" + (27 + i)).setValue(hotZones[i][1])
        .setBackground("#FCA5A5");
    }
  }

  navigateToSheet("Executive Command");
  ss.toast("Hot zones analysis complete", "Success");
}

/**
 * Identifies rising star stewards based on performance
 */
function identifyRisingStars() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Executive Command");

  if (!dash) {
    rebuildExecutiveDashboard();
    dash = ss.getSheetByName("Executive Command");
  }

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

  // Display rising stars
  var startRow = 35;
  dash.getRange("B35").setValue("RISING STARS (Top Performers)")
    .setFontWeight("bold").setBackground("#10B981").setFontColor("white");
  dash.getRange("B35:E35").merge();

  dash.getRange("B36:E36").setValues([["Name", "Score", "Win Rate", "Avg Days"]])
    .setFontWeight("bold").setBackground("#D1FAE5");

  for (var i = 0; i < risingStars.length; i++) {
    var star = risingStars[i];
    dash.getRange("B" + (37 + i)).setValue(star.name);
    dash.getRange("C" + (37 + i)).setValue(star.score);
    dash.getRange("D" + (37 + i)).setValue(typeof star.winRate === 'number' ? star.winRate + "%" : star.winRate);
    dash.getRange("E" + (37 + i)).setValue(star.avgDays);
  }

  // Add star emoji to top performer
  if (risingStars.length > 0) {
    dash.getRange("B37").setFontWeight("bold").setBackground("#A7F3D0");
  }

  navigateToSheet("Executive Command");
  ss.toast("Rising stars identified", "Success");
}

/**
 * Generates management hostility funnel report
 */
function renderHostilityFunnel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName("Member Analytics");

  if (!dash) {
    rebuildMemberAnalytics();
    dash = ss.getSheetByName("Member Analytics");
  }

  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet) return;

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

  var funnelData = [
    ["MANAGEMENT HOSTILITY FUNNEL", "Count", "Rate"],
    ["Step 1 Filed", step1Count, "100%"],
    ["Denied to Step 2", step1Denied, step1Count > 0 ? Math.round((step1Denied/step1Count)*100) + "%" : "0%"],
    ["Denied to Step 3", step2Denied, step2Count > 0 ? Math.round((step2Denied/step2Count)*100) + "%" : "0%"],
    ["To Arbitration", arbitrationCount, step3Count > 0 ? Math.round((arbitrationCount/step3Count)*100) + "%" : "0%"]
  ];

  dash.getRange("H30:J34").setValues(funnelData)
    .setBorder(true, true, true, true, null, null, "#DC2626", SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("H30:J30").setFontWeight("bold").setBackground("#FEE2E2");
  dash.getRange("H31:H34").setFontWeight("bold");

  // Color code rates
  for (var r = 31; r <= 34; r++) {
    var rateVal = dash.getRange("J" + r).getValue();
    var numRate = parseInt(rateVal);
    if (numRate >= 60) {
      dash.getRange("J" + r).setBackground("#FCA5A5");
    } else if (numRate >= 40) {
      dash.getRange("J" + r).setBackground("#FEF3C7");
    } else {
      dash.getRange("J" + r).setBackground("#D1FAE5");
    }
  }

  navigateToSheet("Member Analytics");
  ss.toast("Hostility funnel generated", "Success");
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

    // 1. Refresh Executive Dashboard if it exists
    if (ss.getSheetByName("Executive Command")) {
      rebuildExecutiveDashboard();
      Logger.log('Executive Command dashboard refreshed');
    }

    // 2. Refresh Member Analytics if it exists
    if (ss.getSheetByName("Member Analytics")) {
      rebuildMemberAnalytics();
      Logger.log('Member Analytics dashboard refreshed');
    }

    // 3. Refresh hidden calculation sheets if function exists
    if (typeof rebuildAllHiddenSheets === 'function') {
      rebuildAllHiddenSheets();
      Logger.log('Hidden calculation sheets refreshed');
    }

    // 4. Check for critical dashboard alerts
    checkDashboardAlerts();
    Logger.log('Dashboard alerts checked');

    // 5. Check for overdue grievances and send reminders
    checkOverdueGrievances_();

    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;

    Logger.log('Midnight Auto-Refresh completed in ' + duration + ' seconds');

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
 * Refreshes all dashboard visuals and checks alerts
 */
function refreshAllVisuals() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Refresh Executive Dashboard if it exists
    if (ss.getSheetByName("Executive Command")) {
      rebuildExecutiveDashboard();
    }

    // Refresh Member Analytics if it exists
    if (ss.getSheetByName("Member Analytics")) {
      rebuildMemberAnalytics();
    }

    // Check for alerts
    checkDashboardAlerts();

    ss.toast("Global Refresh Complete", "System");
  } catch (e) {
    Logger.log("Refresh Error: " + e.message);
  }
}

/**
 * Sends Member Analytics dashboard link to specified email
 */
function sendMemberDashboardLink() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure the sheet exists
  var sheet = ss.getSheetByName("Member Analytics");
  if (!sheet) {
    rebuildMemberAnalytics();
    sheet = ss.getSheetByName("Member Analytics");
  }

  var url = ss.getUrl() + "#gid=" + sheet.getSheetId();
  var response = ui.prompt('Send Report', 'Enter Member Email:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    var email = response.getResponseText();

    if (!email || !email.includes('@')) {
      ui.alert('Please enter a valid email address.');
      return;
    }

    var body = "Hello,\n\n" +
      "View your Union strength report: " + url + "\n\n" +
      "HOW TO READ:\n" +
      "1. Morale Gauge: Morale over 3.75 is green (healthy).\n" +
      "2. Pipeline: Tracks our movement growth.\n" +
      "3. PII Safe: No personal info is shown here.\n\n" +
      "509 Strategic Command.";

    try {
      MailApp.sendEmail(email, "Your Union Strength Report", body);
      ui.alert('Report link sent to ' + email);
    } catch (e) {
      ui.alert('Error sending email: ' + e.message);
    }
  }
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
 * Generates missing Member IDs based on Unit Code
 * Uses dynamic column references from MEMBER_COLS
 * Unit codes are read from Config sheet (column AT)
 * Format: [UNIT_PREFIX]-[SEQUENCE]-H (e.g., MS-101-H)
 */
function generateMissingMemberIDs() {
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
 * Checks for duplicate Member IDs and highlights them
 */
function checkDuplicateMemberIDs() {
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
 */
function promoteSelectedMemberToSteward() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory sheet not found.');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a member row (not the header).');
    return;
  }

  // Check if already a steward
  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();
  if (currentStatus === 'Yes') {
    SpreadsheetApp.getUi().alert('This member is already a Steward.');
    return;
  }

  // Promote to steward
  sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('Yes');

  // Get member details
  var email = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();
  var lastName = sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();
  var name = firstName + ' ' + lastName;

  // Send toolkit email if email is available
  if (email) {
    sendStewardToolkit_(email, name);
  }

  SpreadsheetApp.getUi().alert(name + ' has been promoted to Steward!' +
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
 */
function demoteSelectedSteward() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Member Directory sheet not found.');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a member row (not the header).');
    return;
  }

  var currentStatus = sheet.getRange(row, MEMBER_COLS.IS_STEWARD).getValue();
  if (currentStatus !== 'Yes') {
    SpreadsheetApp.getUi().alert('This member is not currently a Steward.');
    return;
  }

  var name = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue() + ' ' +
             sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm Demotion',
    'Are you sure you want to remove Steward status from ' + name + '?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    sheet.getRange(row, MEMBER_COLS.IS_STEWARD).setValue('No');
    sheet.getRange(row, MEMBER_COLS.COMMITTEES).setValue('');  // Clear committee assignments
    ui.alert(name + ' has been removed from Steward status.');
  }
}

// ============================================================================
// 9. GLOBAL UI STYLING ENGINE
// ============================================================================

/**
 * Applies global Roboto theme styling to all data sheets
 * Includes zebra striping and status-based coloring
 */
function applyGlobalStyling() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToStyle = [SHEETS.GRIEVANCE_LOG, SHEETS.MEMBER_DIR];

  sheetsToStyle.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) return;

    // Apply header styling
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontFamily(COMMAND_CONFIG.THEME.FONT)
               .setFontSize(11)
               .setFontWeight('bold')
               .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
               .setFontColor(COMMAND_CONFIG.THEME.HEADER_TEXT)
               .setHorizontalAlignment('center');

    // Apply data styling with zebra stripes
    for (var row = 2; row <= lastRow; row++) {
      var rowRange = sheet.getRange(row, 1, 1, lastCol);
      rowRange.setFontFamily(COMMAND_CONFIG.THEME.FONT)
              .setFontSize(COMMAND_CONFIG.THEME.FONT_SIZE)
              .setVerticalAlignment('middle');

      // Zebra striping
      if (row % 2 === 0) {
        rowRange.setBackground(COMMAND_CONFIG.THEME.ALT_ROW);
      } else {
        rowRange.setBackground(null);
      }

      // Status coloring for Grievance Log
      if (sheetName === SHEETS.GRIEVANCE_LOG) {
        var statusCell = sheet.getRange(row, GRIEVANCE_COLS.STATUS);
        var status = statusCell.getValue();

        if (COMMAND_CONFIG.STATUS_COLORS[status]) {
          var colors = COMMAND_CONFIG.STATUS_COLORS[status];
          statusCell.setBackground(colors.bg)
                    .setFontColor(colors.text)
                    .setFontWeight('bold');
        }
      }
    }
  });

  ss.toast('Global styling applied to all data sheets.', 'Style Engine', 3);
}

/**
 * Resets all custom formatting and applies default theme
 */
function resetToDefaultTheme() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Reset Theme',
    'This will reset all custom formatting. Continue?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    applyGlobalStyling();
    ui.alert('Theme has been reset to defaults.');
  }
}
