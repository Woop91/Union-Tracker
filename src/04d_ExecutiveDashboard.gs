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
      var unit = row[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
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
        if (isTruthyValue(memberData[m][MEMBER_COLS.IS_STEWARD - 1])) stats.stewardCount++;
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
    '  <base target="_top">' + getMobileOptimizedHead() +
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
    ' + getClientSideEscapeHtml() + ' +
    '    window.onload = function() {' +
    '      google.script.run.withSuccessHandler(renderDashboard).withFailureHandler(showError).getDashboardStats();' +
    '    };' +
    '    function showError(err) {' +
    '      document.getElementById("content").innerHTML = "<div class=\\"loading\\">Error loading data: " + escapeHtml(err.message) + "</div>";' +
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
    '      else { workload.slice(0, 10).forEach(function(s) { html += "<div class=\\"steward-item\\"><span class=\\"steward-name\\">" + escapeHtml(s.name) + "</span><span class=\\"steward-count\\">" + s.count + "</span></div>"; }); }' +
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
 * Shows the unified Steward Dashboard (web app URL)
 * Opens the unified dashboard in steward mode (with PII)
 * v4.4.0: Now opens as web app instead of modal for better experience
 */
function showStewardDashboard() {
  var url = ScriptApp.getService().getUrl() + '?mode=steward';
  var html = HtmlService.createHtmlOutput(
    '<html><head>' + getMobileOptimizedHead() + '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:100vh;background:#0f172a;color:#f8fafc;margin:0;padding:16px}' +
    '.icon{font-size:clamp(36px,10vw,48px);margin-bottom:16px}h1{margin:0 0 8px;font-size:clamp(18px,5vw,24px);text-align:center}p{color:#94a3b8;margin:0 0 24px;text-align:center;max-width:400px;line-height:1.5;font-size:clamp(13px,3.5vw,15px);padding:0 8px}' +
    'a.open-link{background:#3b82f6;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:600;display:inline-block;min-height:44px;line-height:20px;text-align:center}a.open-link:hover{background:#2563eb}' +
    '.copy-btn{background:#475569;cursor:pointer;border:none;padding:10px 16px;border-radius:8px;color:white;font-size:clamp(11px,3vw,13px);min-height:44px}' +
    '.url{background:#1e293b;padding:12px;border-radius:8px;font-family:monospace;font-size:clamp(10px,2.5vw,12px);word-break:break-all;max-width:90%;margin-bottom:16px;border:1px solid #334155;width:100%}' +
    '.warning{background:rgba(239,68,68,0.2);color:#fca5a5;padding:8px 16px;border-radius:8px;font-size:clamp(10px,2.5vw,12px);margin-bottom:16px}' +
    '.btn-row{display:flex;gap:8px;flex-wrap:wrap;justify-content:center}' +
    '@media(max-width:480px){.btn-row{flex-direction:column;width:100%}a.open-link,.copy-btn{width:100%;text-align:center}}' +
    '</style></head><body><div class="icon">🛡️</div><h1>Steward Command Center</h1>' +
    '<div class="warning">INTERNAL USE ONLY - Contains PII</div>' +
    '<p>Open the Steward Dashboard web app. This version includes full member details and sensitive information.</p>' +
    '<div class="url" id="url">' + url + '</div>' +
    '<div class="btn-row"><a class="open-link" href="' + url + '" target="_blank">Open Dashboard</a>' +
    '<button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'url\').textContent);this.textContent=\'Copied!\';setTimeout(function(){document.querySelector(\'.copy-btn\').textContent=\'Copy URL\'},2000)">Copy URL</button></div>' +
    '</body></html>'
  ).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Steward Command Center');
}

/**
 * @deprecated v4.4.0 - Replaced by getUnifiedDashboardData() and getUnifiedDashboardDataAPI()
 * Legacy getStewardDashboardData function has been removed.
 * Use the unified dashboard system via doGet() web app with ?mode=steward or ?mode=member
 */

/**
 * @deprecated v4.4.0 - Replaced by getUnifiedDashboardHtml()
 * Legacy getStewardDashboardHtml_ function has been removed.
 * Use the unified dashboard system via doGet() web app with ?mode=steward or ?mode=member
 */

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

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Hot Zones are now in the unified Steward Dashboard (Hot Spots tab).
 */
function renderHotZones() {
  showStewardDashboard();
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Steward metrics are now in the unified Steward Dashboard (Workload tab).
 */
function identifyRisingStars() {
  showStewardDashboard();
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Hostility metrics are now in the unified Steward Dashboard (Bargaining tab).
 */
function renderHostilityFunnel() {
  showStewardDashboard();
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
    var _ss = SpreadsheetApp.getActiveSpreadsheet();
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

// REMOVED: refreshAllVisuals_DataRefresh_DEPRECATED - Use refreshAllVisuals() instead

/**
 * Sends Member Analytics dashboard access link to specified email
 * Note: Member Analytics is now a modal dashboard launched from the menu
 */

/**
 * Sends the Member Dashboard URL to the selected member from Member Directory.
 * Uses the currently selected row to get member email and name.
 * This is a PII-protected view link.
 * NOTE: Duplicate exists in 11_SecureMemberDashboard.gs - this version kept for compatibility
 * @deprecated Use emailDashboardLink() in 11_SecureMemberDashboard.gs
 */

/**
 * Shows Steward Performance Modal
 * Displays performance metrics for all stewards including:
 * - Active cases, total cases, and win rates
 * - Response time averages
 * - Member satisfaction scores (if available)
 * NOTE: Duplicate exists in 11_SecureMemberDashboard.gs - this version kept for compatibility
 * @deprecated Use showStewardPerformanceModal() in 11_SecureMemberDashboard.gs
 */
function showStewardPerformanceModal_UIService_() {
  var stewardData = getStewardWorkload();

  if (!stewardData || stewardData.length === 0) {
    SpreadsheetApp.getUi().alert('No steward data available. Ensure stewards are marked in the Member Directory.');
    return;
  }

  var html = '<!DOCTYPE html><html><head>' + getMobileOptimizedHead() +
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
 * @deprecated Use generateMissingMemberIDs() from 02_DataManagers.gs
 */
function generateMissingMemberIDs_UIService_() {
  generateMissingMemberIDs();
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
 * NOTE: createGrievancePDF(folder, data) exists in 10_CommandCenter.gs with different signature
 * @param {Object} data - Grievance data object with name, details, etc.
 * @returns {File} The created PDF file
 */
function createGrievancePDF_UIService_(data) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create archive folder
  var folder;
  if (COMMAND_CONFIG.ARCHIVE_FOLDER_ID) {
    try {
      folder = DriveApp.getFolderById(COMMAND_CONFIG.ARCHIVE_FOLDER_ID);
    } catch (_e) {
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
 * NOTE: Duplicate exists in 05_Integrations.gs - this version is kept for backwards compatibility
 * @deprecated Use createPDFForSelectedGrievance() in 05_Integrations.gs
 */
function createPDFForSelectedGrievance_UIService_() {
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

  // Get or create the grievance archive folder
  var archiveFolderName = COMMAND_CONFIG.ARCHIVE_FOLDER_NAME || '509 Grievance Archive';
  var folders = DriveApp.getFoldersByName(archiveFolderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(archiveFolderName);

  var pdf = createGrievancePDF(folder, grievanceData);

  SpreadsheetApp.getUi().alert('PDF created successfully!\n\nFile: ' + pdf.getName() +
    '\n\nYou can find it in your 509 Grievance Archive folder.');
}

// ============================================================================
// 8. STEWARD PROMOTION ENGINE (Helper Functions)
// ============================================================================
// Note: Main promote/demote functions moved to 02_MemberManager.gs

/**
 * Sends steward toolkit email to newly promoted steward
 * @param {string} email - Steward's email address
 * @param {string} name - Steward's name
 * @private
 */

// REMOVED: demoteSelectedSteward_UIService_DEPRECATED - Use demoteSelectedSteward() in 02_MemberManager.gs instead

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

