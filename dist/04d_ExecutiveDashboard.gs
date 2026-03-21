// ============================================================================
// 1. NAVIGATION HELPERS
// ============================================================================

// navigateToSheet() - REMOVED DUPLICATE - see line 565 for main definition

// ============================================================================
// 2. EXECUTIVE COMMAND MODAL (SPA Architecture - Bridge Pattern)
// ============================================================================

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

    // Filter to only rows with a valid grievance ID (starts with "G")
    logData = logData.filter(function(row) {
      var gid = (row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '').toString();
      return isGrievanceId_(gid);
    });
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

      // M7: Count by step — use word-boundary matching to prevent 'step 10' matching 'step 1'
      if (/\bstep\s*1\b/.test(currentStep) || currentStep === '1' || currentStep === 'step i') stats.activeSteps.step1++;
      if (/\bstep\s*2\b/.test(currentStep) || currentStep === '2' || currentStep === 'step ii') stats.activeSteps.step2++;
      if (/\barbitration\b/.test(currentStep) || /\bstep\s*3\b/.test(currentStep) || currentStep === 'step iii') stats.activeSteps.arbitration++;

      // Count outcomes using the Resolution column (not Status)
      var resolution = (row[GRIEVANCE_COLS.RESOLUTION - 1] || '').toString().toLowerCase();
      if (resolution === 'won' || resolution === 'sustained') stats.outcomes.wins++;
      if (resolution === 'denied' || resolution === 'lost') stats.outcomes.losses++;
      if (resolution === 'settled') stats.outcomes.settled++;
      if (resolution === 'withdrawn') stats.outcomes.withdrawn++;

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

      // Check for overdue — use deadline matching the current step
      if (status === 'open' || status === 'pending info') {
        var dueDate = null;
        if (/\bstep\s*2\b/.test(currentStep) || currentStep === '2' || currentStep === 'step ii') {
          dueDate = row[GRIEVANCE_COLS.STEP2_DUE - 1];
        } else if (/\barbitration\b/.test(currentStep) || /\bstep\s*3\b/.test(currentStep) || currentStep === 'step iii') {
          dueDate = row[GRIEVANCE_COLS.STEP3_APPEAL_DUE - 1];
        } else {
          dueDate = row[GRIEVANCE_COLS.STEP1_DUE - 1];
        }
        if (dueDate && new Date(dueDate) < new Date()) {
          stats.overdueCount++;
        }
      }
    });

    // Calculate win rate — only resolved outcomes in denominator
    var totalResolved = stats.outcomes.wins + stats.outcomes.losses + stats.outcomes.settled + stats.outcomes.withdrawn;
    if (totalResolved > 0) {
      stats.winRate = Math.round((stats.outcomes.wins / totalResolved) * 100);
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

  // Get morale score from satisfaction data (v4.23.0: dynamic via getSatisfactionSummary)
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  if (satSheet && satSheet.getLastRow() > 1) {
    try {
      var summary = getSatisfactionSummary();
      if (summary && summary.sections && summary.sections['OVERALL_SAT']) {
        var overallAvg = summary.sections['OVERALL_SAT'].avg;
        if (overallAvg !== null && !isNaN(overallAvg)) {
          stats.moraleScore = Math.round(overallAvg * 10) / 10;
        }
      }
    } catch(e) {
      Logger.log('Error reading satisfaction summary for morale score: ' + e.message);
    }
  }

  return JSON.stringify(stats);
}
/**
 * Gets executive metrics from dashboard calculations
 * @returns {Object} Metrics object with activeGrievances, winRate, overdueSteps
 * @private
 */
function getExecutiveMetrics_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC);

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
        if (status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING) openCount++;
        if (status === GRIEVANCE_STATUS.WON) wonCount++;
        if (GRIEVANCE_CLOSED_STATUSES.indexOf(status) !== -1) closedCount++;

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
    '<div class="url" id="url">' + escapeHtml(url) + '</div>' +
    '<div class="btn-row"><a class="open-link" href="' + escapeHtml(url) + '" target="_blank">Open Dashboard</a>' +
    '<button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById(\'url\').textContent);this.textContent=\'Copied!\';setTimeout(function(){document.querySelector(\'.copy-btn\').textContent=\'Copy URL\'},2000)">Copy URL</button></div>' +
    '</body></html>'
  ).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Steward Command Center');
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
        "DASHBOARD ALERT",
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
      if ((status === GRIEVANCE_STATUS.OPEN || status === GRIEVANCE_STATUS.PENDING) &&
          (daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0))) {
        overdueList.push({
          id: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1],
          name: maskName(data[i][GRIEVANCE_COLS.FIRST_NAME - 1], data[i][GRIEVANCE_COLS.LAST_NAME - 1]),
          steward: data[i][GRIEVANCE_COLS.STEWARD - 1],
          days: daysToDeadline
        });
      }
    }

    if (overdueList.length > 0) {
      var chiefStewardEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
      if (chiefStewardEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(chiefStewardEmail)) {
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
 * Emails the Executive Dashboard as a PDF snapshot
 */
function emailExecutivePDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var date = new Date();
    var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Export only the active sheet as PDF (not the entire spreadsheet)
    var activeSheet = ss.getActiveSheet();
    var sheetGid = activeSheet.getSheetId();
    var url = ss.getUrl().replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf' +
      '&gid=' + sheetGid +
      '&size=letter&portrait=true&fitw=1' +
      '&gridlines=false&printtitle=false&sheetnames=false';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token }
    });
    var blob = response.getBlob().setName("Health_Report_" + dateStr + ".pdf");

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

// ============================================================================
// 7. SIGNATURE PDF ENGINE
// ============================================================================

// ============================================================================
// 8. STEWARD PROMOTION ENGINE (Helper Functions)
// ============================================================================
// Note: Main promote/demote functions moved to 02_MemberManager.gs

// demoteSelectedSteward_UIService_DEPRECATED removed — see demoteSelectedSteward() in 02_MemberManager.gs

// ============================================================================
// 9. GLOBAL UI STYLING ENGINE (Status Colors)
// ============================================================================
