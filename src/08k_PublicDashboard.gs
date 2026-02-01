/**
 * ============================================================================
 * 08k_PublicDashboard.gs - Public Dashboard & Flagged Submissions Module
 * ============================================================================
 *
 * This module contains functions for:
 * - Secure member dashboard HTML generation
 * - Public data retrieval (no PII exposed)
 * - Flagged survey submission review and management
 *
 * All functions in this module are designed to provide aggregate statistics
 * and public-facing data without exposing personally identifiable information.
 *
 * Dependencies:
 * - SHEETS constant (sheet names)
 * - MEMBER_COLS, GRIEVANCE_COLS, SATISFACTION_COLS (column mappings)
 * - syncSatisfactionValues() for updating dashboard after approval/rejection
 *
 * @fileoverview Public dashboard data functions and flagged submissions management
 * @version 1.0.0
 */

// ============================================================================
// FLAGGED SUBMISSIONS REVIEW
// ============================================================================

/**
 * Show the flagged submissions review modal dialog
 */
function showFlaggedSubmissionsReview() {
  var html = HtmlService.createHtmlOutput(getFlaggedSubmissionsHtml())
    .setWidth(700)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Flagged Survey Submissions Review');
}

/**
 * Get HTML for flagged submissions review interface
 * @returns {string} HTML content
 */
function getFlaggedSubmissionsHtml() {
  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    ':root{--purple:#5B4B9E;--green:#059669;--red:#DC2626;--orange:#F97316}' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f5f5f5;padding:20px}' +
    '.container{max-width:650px;margin:0 auto}' +
    '.stats-row{display:flex;gap:15px;margin-bottom:20px}' +
    '.stat-card{flex:1;background:white;padding:20px;border-radius:12px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.stat-card.pending{border-left:4px solid var(--orange)}' +
    '.stat-card.verified{border-left:4px solid var(--green)}' +
    '.stat-value{font-size:32px;font-weight:bold;color:#333}' +
    '.stat-label{font-size:13px;color:#666;margin-top:5px}' +
    '.section{background:white;border-radius:12px;padding:20px;margin-bottom:15px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    '.section-title{font-size:16px;font-weight:600;color:#333;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #eee}' +
    '.email-list{max-height:250px;overflow-y:auto}' +
    '.email-item{display:flex;align-items:center;justify-content:space-between;padding:12px;background:#f8f9fa;border-radius:8px;margin-bottom:8px}' +
    '.email-info{display:flex;align-items:center;gap:10px}' +
    '.email-text{font-size:14px;color:#333}' +
    '.email-date{font-size:12px;color:#666}' +
    '.actions{display:flex;gap:8px}' +
    '.btn{padding:6px 12px;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:500}' +
    '.btn-approve{background:#059669;color:white}' +
    '.btn-reject{background:#DC2626;color:white}' +
    '.empty-state{text-align:center;padding:40px;color:#666}' +
    '.info-box{background:#E8F4FD;padding:15px;border-radius:8px;margin-bottom:15px;font-size:13px;color:#1E40AF}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div id="content"><div class="empty-state">Loading...</div></div>' +
    '</div>' +
    '<script>' +
    'function load(){google.script.run.withSuccessHandler(render).getFlaggedSubmissionsData()}' +
    'function render(d){' +
    '  var h="<div class=\\"stats-row\\">";' +
    '  h+="<div class=\\"stat-card pending\\"><div class=\\"stat-value\\">"+d.pendingCount+"</div><div class=\\"stat-label\\">Pending Review</div></div>";' +
    '  h+="<div class=\\"stat-card verified\\"><div class=\\"stat-value\\">"+d.verifiedCount+"</div><div class=\\"stat-label\\">Verified Responses</div></div>";' +
    '  h+="</div>";' +
    '  h+="<div class=\\"info-box\\">⚠️ These submissions could not be matched to a member email. Survey answers are protected and not shown here.</div>";' +
    '  h+="<div class=\\"section\\"><div class=\\"section-title\\">📧 Pending Review Emails ("+d.pendingCount+")</div>";' +
    '  if(d.pendingEmails.length===0){' +
    '    h+="<div class=\\"empty-state\\">✅ No submissions pending review</div>";' +
    '  }else{' +
    '    h+="<div class=\\"email-list\\">";' +
    '    d.pendingEmails.forEach(function(e){' +
    '      h+="<div class=\\"email-item\\"><div class=\\"email-info\\">";' +
    '      h+="<span class=\\"email-text\\">"+e.email+"</span>";' +
    '      h+="<span class=\\"email-date\\">"+e.date+" | "+e.quarter+"</span></div>";' +
    '      h+="<div class=\\"actions\\">";' +
    '      h+="<button class=\\"btn btn-approve\\" onclick=\\"approve("+e.row+\")\\">✓ Approve</button>";' +
    '      h+="<button class=\\"btn btn-reject\\" onclick=\\"reject("+e.row+\")\\">✗ Reject</button>";' +
    '      h+="</div></div>";' +
    '    });' +
    '    h+="</div>";' +
    '  }' +
    '  h+="</div>";' +
    '  document.getElementById("content").innerHTML=h;' +
    '}' +
    'function approve(row){' +
    '  if(confirm("Mark this submission as verified? This will include it in statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).approveFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'function reject(row){' +
    '  if(confirm("Reject this submission? It will be excluded from all statistics.")){' +
    '    google.script.run.withSuccessHandler(function(){load()}).rejectFlaggedSubmission(row);' +
    '  }' +
    '}' +
    'load();' +
    '</script></body></html>';
}

/**
 * Get data for flagged submissions review
 * @returns {Object} Pending submissions data (email, date, row number - NO survey answers)
 */
function getFlaggedSubmissionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    pendingCount: 0,
    verifiedCount: 0,
    pendingEmails: []
  };

  if (!satSheet) return result;

  var data = satSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var verified = data[i][SATISFACTION_COLS.VERIFIED - 1];
    var email = data[i][SATISFACTION_COLS.EMAIL - 1] || '(no email provided)';
    var timestamp = data[i][SATISFACTION_COLS.TIMESTAMP - 1];
    var quarter = data[i][SATISFACTION_COLS.QUARTER - 1] || '';

    if (verified === 'Yes') {
      result.verifiedCount++;
    } else if (verified === 'Pending Review') {
      result.pendingCount++;
      result.pendingEmails.push({
        email: email.toString(),
        date: timestamp ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'MMM d, yyyy') : 'Unknown',
        quarter: quarter,
        row: i + 1  // 1-indexed row number for editing
      });
    }
    // Rejected submissions are counted but not shown
  }

  // Sort by most recent first
  result.pendingEmails.sort(function(a, b) { return b.row - a.row; });

  return result;
}

/**
 * Approve a flagged submission - mark as Verified
 * @param {number} rowNum - Row number (1-indexed)
 */
function approveFlaggedSubmission(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet || rowNum < 2) return;

  satSheet.getRange(rowNum, SATISFACTION_COLS.VERIFIED).setValue('Yes');
  satSheet.getRange(rowNum, SATISFACTION_COLS.REVIEWER_NOTES).setValue('Manually approved on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));

  // Update dashboard
  syncSatisfactionValues();
}

/**
 * Reject a flagged submission - mark as Rejected
 * @param {number} rowNum - Row number (1-indexed)
 */
function rejectFlaggedSubmission(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);

  if (!satSheet || rowNum < 2) return;

  satSheet.getRange(rowNum, SATISFACTION_COLS.VERIFIED).setValue('Rejected');
  satSheet.getRange(rowNum, SATISFACTION_COLS.IS_LATEST).setValue('No');
  satSheet.getRange(rowNum, SATISFACTION_COLS.REVIEWER_NOTES).setValue('Rejected on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));

  // Update dashboard
  syncSatisfactionValues();
}

// ============================================================================
// SECURE MEMBER DASHBOARD HTML
// ============================================================================

/**
 * Generates HTML for the secure member dashboard
 * @param {Object} stats - Grievance statistics
 * @param {Array} stewards - Array of steward objects
 * @param {Object} satisfaction - Satisfaction statistics
 * @param {Object} coverage - Steward coverage statistics
 * @returns {string} HTML content
 */
function getSecureMemberDashboardHtml(stats, stewards, satisfaction, coverage) {
  // Prepare steward data for display (sanitize for JSON)
  var stewardList = stewards.slice(0, 12).map(function(s) {
    return {
      firstName: (s['First Name'] || '').toString().replace(/"/g, '\\"'),
      lastName: (s['Last Name'] || '').toString().replace(/"/g, '\\"'),
      unit: (s['Unit'] || 'General').toString().replace(/"/g, '\\"'),
      location: (s['Work Location'] || '').toString().replace(/"/g, '\\"'),
      email: (s['Email'] || '').toString().replace(/"/g, '\\"')
    };
  });

  // Build trend data for area chart
  var trendChartData = [['Quarter', 'Trust Score']];
  if (satisfaction.trendData && satisfaction.trendData.length > 0) {
    satisfaction.trendData.forEach(function(item) {
      trendChartData.push(item);
    });
  } else {
    trendChartData.push(['Current', satisfaction.avgTrust || 7]);
  }

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Roboto", "Segoe UI", sans-serif; background: #f0f4f8; color: #1e293b; padding: 15px; }' +
    '.header { display: flex; align-items: center; color: #4338ca; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e0e7ff; }' +
    '.header h2 { font-size: 18px; font-weight: 700; margin-left: 8px; }' +
    '.card { background: white; border-radius: 12px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 12px; }' +
    '.stat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }' +
    '.stat-card { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; padding: 15px; border-radius: 10px; text-align: center; }' +
    '.stat-card.green { background: linear-gradient(135deg, #059669 0%, #047857 100%); }' +
    '.stat-card.blue { background: linear-gradient(135deg, #3B82F6 0%, #2563EB 100%); }' +
    '.stat-val { font-size: 28px; font-weight: 800; display: block; }' +
    '.stat-label { font-size: 11px; text-transform: uppercase; opacity: 0.9; font-weight: 500; }' +
    '.section-title { font-size: 12px; text-transform: uppercase; color: #64748b; font-weight: 600; margin-bottom: 10px; display: flex; align-items: center; gap: 6px; }' +
    '.chart-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }' +
    '.chart-container { height: 140px; width: 100%; }' +
    '.progress-section { margin-bottom: 8px; }' +
    '.progress-header { display: flex; justify-content: space-between; font-size: 12px; color: #475569; margin-bottom: 4px; }' +
    '.progress-bg { background: #e2e8f0; border-radius: 10px; height: 10px; width: 100%; }' +
    '.progress-fill { background: linear-gradient(90deg, #7C3AED, #a78bfa); height: 100%; border-radius: 10px; transition: width 0.5s ease; }' +
    '.search-box { width: 100%; padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 13px; margin-bottom: 10px; }' +
    '.search-box:focus { outline: none; border-color: #7C3AED; box-shadow: 0 0 0 2px rgba(124,58,237,0.1); }' +
    '.steward-list { max-height: 180px; overflow-y: auto; }' +
    '.steward-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid #f1f5f9; }' +
    '.steward-row:last-child { border-bottom: none; }' +
    '.steward-name { font-size: 13px; font-weight: 500; }' +
    '.steward-unit { background: #eff6ff; color: #1e40af; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; }' +
    '.steward-actions { display: flex; gap: 6px; }' +
    '.steward-actions a { color: #64748b; text-decoration: none; transition: color 0.2s; }' +
    '.steward-actions a:hover { color: #7C3AED; }' +
    '.footer { font-size: 9px; color: #94a3b8; text-align: center; margin-top: 12px; padding-top: 10px; border-top: 1px solid #e2e8f0; }' +
    '.trend-area { margin-top: 10px; }' +
    '</style>' +
    '<script type="text/javascript">' +
    'google.charts.load("current", {"packages":["corechart", "gauge"]});' +
    'google.charts.setOnLoadCallback(drawCharts);' +
    'function drawCharts() {' +
    // Issue Mix Pie Chart
    '  var issueData = google.visualization.arrayToDataTable(' + JSON.stringify(stats.categoryData || [['Category', 'Count'], ['No Data', 1]]) + ');' +
    '  var issueOptions = { pieHole: 0.4, chartArea: {width:"90%",height:"90%"}, legend: "none", colors: ["#7C3AED", "#059669", "#F97316", "#3B82F6", "#EC4899", "#6366F1"], pieSliceText: "label", fontSize: 10 };' +
    '  new google.visualization.PieChart(document.getElementById("issue_chart")).draw(issueData, issueOptions);' +
    // Trust Gauge
    '  var gaugeData = google.visualization.arrayToDataTable([["Label", "Value"],["Trust", ' + (satisfaction.avgTrust || 0) + ']]);' +
    '  var gaugeOptions = { width: 130, height: 130, greenFrom: 7, greenTo: 10, yellowFrom: 5, yellowTo: 7, redFrom: 0, redTo: 5, max: 10, minorTicks: 5 };' +
    '  new google.visualization.Gauge(document.getElementById("gauge_div")).draw(gaugeData, gaugeOptions);' +
    // Trend Area Chart
    '  var trendData = google.visualization.arrayToDataTable(' + JSON.stringify(trendChartData) + ');' +
    '  var trendOptions = { legend: "none", chartArea: {width:"85%",height:"70%"}, colors: ["#7C3AED"], areaOpacity: 0.3, hAxis: {textStyle:{fontSize:9}}, vAxis: {minValue:0,maxValue:10,textStyle:{fontSize:9}}, lineWidth: 2 };' +
    '  new google.visualization.AreaChart(document.getElementById("trend_chart")).draw(trendData, trendOptions);' +
    '}' +
    // Steward search filter
    'var stewards = ' + JSON.stringify(stewardList) + ';' +
    'function filterStewards(query) {' +
    '  var q = query.toLowerCase();' +
    '  var list = document.getElementById("steward-list");' +
    '  var html = "";' +
    '  stewards.forEach(function(s) {' +
    '    var fullName = s.firstName + " " + s.lastName;' +
    '    var searchText = (fullName + " " + s.unit + " " + s.location).toLowerCase();' +
    '    if (!q || searchText.indexOf(q) !== -1) {' +
    '      html += "<div class=\\"steward-row\\">";' +
    '      html += "<div><span class=\\"steward-name\\">" + s.firstName + " " + s.lastName + "</span></div>";' +
    '      html += "<div class=\\"steward-actions\\">";' +
    '      html += "<span class=\\"steward-unit\\">" + s.unit + "</span>";' +
    '      if (s.email) { html += " <a href=\\"mailto:" + s.email + "\\" title=\\"Email\\"><i class=\\"material-icons\\" style=\\"font-size:16px\\">email</i></a>"; }' +
    '      html += "</div></div>";' +
    '    }' +
    '  });' +
    '  if (!html) { html = "<div style=\\"text-align:center;padding:15px;color:#94a3b8\\">No stewards found</div>"; }' +
    '  list.innerHTML = html;' +
    '}' +
    'document.addEventListener("DOMContentLoaded", function() { filterStewards(""); });' +
    '</script>' +
    '</head>' +
    '<body>' +
    '<div class="header"><i class="material-icons">verified_user</i><h2>509 MEMBER PORTAL</h2></div>' +
    // Stats Grid
    '<div class="stat-grid">' +
    '<div class="stat-card"><span class="stat-val">' + (stats.open || 0) + '</span><span class="stat-label">Active Cases</span></div>' +
    '<div class="stat-card green"><span class="stat-val">' + (stats.resolved || 0) + '</span><span class="stat-label">Resolved (YTD)</span></div>' +
    '</div>' +
    // Charts Row
    '<div class="chart-row">' +
    '<div class="card"><div class="section-title"><i class="material-icons" style="font-size:14px">pie_chart</i> Issue Mix</div><div id="issue_chart" class="chart-container"></div></div>' +
    '<div class="card" style="text-align:center;"><div class="section-title"><i class="material-icons" style="font-size:14px">speed</i> Member Trust</div><div id="gauge_div" style="display:inline-block;margin-top:5px"></div></div>' +
    '</div>' +
    // Progress Bars
    '<div class="card">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">trending_up</i> Union Goals</div>' +
    '<div class="progress-section">' +
    '<div class="progress-header"><span>Steward Coverage</span><span>' + (coverage.coveragePercent || 0) + '%</span></div>' +
    '<div class="progress-bg"><div class="progress-fill" style="width:' + Math.min(100, coverage.coveragePercent || 0) + '%"></div></div>' +
    '</div>' +
    '<div class="progress-section">' +
    '<div class="progress-header"><span>Survey Participation</span><span>' + (satisfaction.responseCount > 0 ? Math.min(100, Math.round((satisfaction.responseCount / coverage.memberCount) * 100)) : 0) + '%</span></div>' +
    '<div class="progress-bg"><div class="progress-fill" style="width:' + (satisfaction.responseCount > 0 ? Math.min(100, Math.round((satisfaction.responseCount / Math.max(1, coverage.memberCount)) * 100)) : 0) + '%"></div></div>' +
    '</div>' +
    '</div>' +
    // Trust Trend
    '<div class="card trend-area">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">show_chart</i> Trust Score Trend</div>' +
    '<div id="trend_chart" style="height:80px;width:100%"></div>' +
    '</div>' +
    // Steward Directory with Search
    '<div class="card">' +
    '<div class="section-title"><i class="material-icons" style="font-size:14px">groups</i> Find Your Steward</div>' +
    '<input type="text" class="search-box" placeholder="Search by name, unit, or location..." oninput="filterStewards(this.value)">' +
    '<div id="steward-list" class="steward-list"></div>' +
    '</div>' +
    // Footer
    '<div class="footer"><i class="material-icons" style="font-size:11px;vertical-align:middle">lock</i> Protected View - No Private PII Displayed</div>' +
    '</body>' +
    '</html>';
}

// ============================================================================
// PUBLIC DATA RETRIEVAL (NO PII)
// ============================================================================

/**
 * Get public overview data (no PII)
 * @returns {Object} Overview statistics
 */
function getPublicOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    totalMembers: 0,
    totalStewards: 0,
    totalGrievances: 0,
    winRate: 0,
    locationBreakdown: []
  };

  // Count members and stewards
  if (memberSheet) {
    var memberData = memberSheet.getDataRange().getValues();
    var locationCounts = {};
    var stewardCount = 0;

    for (var i = 1; i < memberData.length; i++) {
      var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
      if (!memberId || !memberId.toString().match(/^M/i)) continue;

      result.totalMembers++;

      // Count by location
      var location = memberData[i][MEMBER_COLS.LOCATION - 1] || 'Unknown';
      locationCounts[location] = (locationCounts[location] || 0) + 1;

      // Count stewards
      var isSteward = memberData[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isSteward === true || isSteward === 'Yes' || isSteward === 'TRUE') {
        stewardCount++;
      }
    }

    result.totalStewards = stewardCount;

    // Convert location counts to array and sort
    Object.keys(locationCounts).forEach(function(loc) {
      result.locationBreakdown.push({ location: loc, count: locationCounts[loc] });
    });
    result.locationBreakdown.sort(function(a, b) { return b.count - a.count; });
    result.locationBreakdown = result.locationBreakdown.slice(0, 10); // Top 10
  }

  // Count grievances and win rate
  if (grievanceSheet) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    var won = 0, total = 0;

    for (var j = 1; j < grievanceData.length; j++) {
      var grievanceId = grievanceData[j][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      if (!grievanceId) continue;

      total++;
      var resolution = (grievanceData[j][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString().toLowerCase();
      if (resolution.includes('won') || resolution.includes('favor')) {
        won++;
      }
    }

    result.totalGrievances = total;
    result.winRate = total > 0 ? Math.round(won / total * 100) : 0;
  }

  return result;
}

/**
 * Get public survey data (anonymized)
 * Filters to only include Verified='Yes' and optionally IS_LATEST='Yes' responses
 * @param {boolean} includeHistory - If true, include superseded responses; if false, only latest per member
 * @returns {Object} Survey statistics
 */
function getPublicSurveyData(includeHistory) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = {
    totalResponses: 0,
    verifiedResponses: 0,
    avgSatisfaction: 0,
    responseRate: 0,
    sectionScores: [],
    includesHistory: includeHistory || false
  };

  if (!satSheet) return result;

  var data = satSheet.getDataRange().getValues();
  if (data.length < 2) return result;

  // Filter rows to only include verified responses
  // If includeHistory is false (default), also filter to IS_LATEST='Yes'
  var validRows = [];
  for (var i = 1; i < data.length; i++) {
    var verified = data[i][SATISFACTION_COLS.VERIFIED - 1];
    var isLatest = data[i][SATISFACTION_COLS.IS_LATEST - 1];

    // Only include Verified='Yes' responses
    if (verified !== 'Yes') continue;

    // If not including history, only include IS_LATEST='Yes'
    if (!includeHistory && isLatest !== 'Yes') continue;

    validRows.push(data[i]);
  }

  result.totalResponses = data.length - 1; // Total submissions (all)
  result.verifiedResponses = validRows.length; // Verified responses used in calculations

  // Calculate average satisfaction (Q6 - Satisfied with representation)
  var satSum = 0, satCount = 0;
  for (var j = 0; j < validRows.length; j++) {
    var sat = parseFloat(validRows[j][SATISFACTION_COLS.Q6_SATISFIED_REP - 1]);
    if (!isNaN(sat)) {
      satSum += sat;
      satCount++;
    }
  }
  result.avgSatisfaction = satCount > 0 ? satSum / satCount : 0;

  // Response rate (unique verified members / total members)
  if (memberSheet) {
    var memberCount = memberSheet.getLastRow() - 1;
    // Count unique verified member IDs
    var uniqueMembers = {};
    for (var k = 0; k < validRows.length; k++) {
      var memberId = validRows[k][SATISFACTION_COLS.MATCHED_MEMBER_ID - 1];
      if (memberId) uniqueMembers[memberId] = true;
    }
    var uniqueCount = Object.keys(uniqueMembers).length;
    result.responseRate = memberCount > 0 ? Math.round(uniqueCount / memberCount * 100) : 0;
  }

  // Section scores using only verified responses
  var sections = [
    { name: 'Overall Satisfaction', cols: [SATISFACTION_COLS.Q6_SATISFIED_REP, SATISFACTION_COLS.Q7_TRUST_UNION, SATISFACTION_COLS.Q8_FEEL_PROTECTED] },
    { name: 'Steward Ratings', cols: [SATISFACTION_COLS.Q10_TIMELY_RESPONSE, SATISFACTION_COLS.Q11_TREATED_RESPECT, SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS] },
    { name: 'Chapter Effectiveness', cols: [SATISFACTION_COLS.Q21_UNDERSTAND_ISSUES, SATISFACTION_COLS.Q22_CHAPTER_COMM, SATISFACTION_COLS.Q23_ORGANIZES] },
    { name: 'Local Leadership', cols: [SATISFACTION_COLS.Q26_DECISIONS_CLEAR, SATISFACTION_COLS.Q27_UNDERSTAND_PROCESS, SATISFACTION_COLS.Q28_TRANSPARENT_FINANCE] },
    { name: 'Communication', cols: [SATISFACTION_COLS.Q41_CLEAR_ACTIONABLE, SATISFACTION_COLS.Q42_ENOUGH_INFO] }
  ];

  sections.forEach(function(section) {
    var sum = 0, count = 0;
    for (var m = 0; m < validRows.length; m++) {
      section.cols.forEach(function(col) {
        if (col) {
          var val = parseFloat(validRows[m][col - 1]);
          if (!isNaN(val)) {
            sum += val;
            count++;
          }
        }
      });
    }
    result.sectionScores.push({
      section: section.name,
      score: count > 0 ? sum / count : 0
    });
  });

  result.sectionScores.sort(function(a, b) { return b.score - a.score; });

  return result;
}

/**
 * Get public grievance data (no PII)
 * @returns {Object} Grievance statistics
 */
function getPublicGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var result = {
    total: 0,
    open: 0,
    won: 0,
    settled: 0,
    avgDaysToResolve: 0,
    byType: [],
    byStatus: []
  };

  if (!grievanceSheet) return result;

  var data = grievanceSheet.getDataRange().getValues();
  var typeCounts = {};
  var statusCounts = {};
  var daysToResolve = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (!grievanceId) continue;

    result.total++;

    var status = data[i][GRIEVANCE_COLS.STATUS - 1] || 'Unknown';
    var resolution = (data[i][GRIEVANCE_COLS.RESOLUTION - 1] || '').toString();
    var gType = data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';

    // Count by status
    statusCounts[status] = (statusCounts[status] || 0) + 1;

    // Count by type
    typeCounts[gType] = (typeCounts[gType] || 0) + 1;

    // Track open/won/settled
    if (status === 'Open' || status === 'Pending Info') {
      result.open++;
    }
    if (resolution.toLowerCase().includes('won') || resolution.toLowerCase().includes('favor')) {
      result.won++;
    }
    if (resolution.toLowerCase().includes('settled')) {
      result.settled++;
    }

    // Calculate days to resolve for closed grievances
    if (status === 'Closed' || status === 'Resolved') {
      var dateOpened = data[i][GRIEVANCE_COLS.DATE_FILED - 1];
      var dateClosed = data[i][GRIEVANCE_COLS.DATE_CLOSED - 1];
      if (dateOpened && dateClosed) {
        var days = Math.round((new Date(dateClosed) - new Date(dateOpened)) / (1000 * 60 * 60 * 24));
        if (days > 0) daysToResolve.push(days);
      }
    }
  }

  // Average days to resolve
  if (daysToResolve.length > 0) {
    result.avgDaysToResolve = Math.round(daysToResolve.reduce(function(a, b) { return a + b; }, 0) / daysToResolve.length);
  }

  // Convert to arrays
  Object.keys(typeCounts).forEach(function(t) {
    result.byType.push({ type: t, count: typeCounts[t] });
  });
  result.byType.sort(function(a, b) { return b.count - a.count; });
  result.byType = result.byType.slice(0, 8);

  Object.keys(statusCounts).forEach(function(s) {
    result.byStatus.push({ status: s, count: statusCounts[s] });
  });
  result.byStatus.sort(function(a, b) { return b.count - a.count; });

  return result;
}

/**
 * Get public steward data (contact info only)
 * @returns {Object} Steward directory
 */
function getPublicStewardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  var result = { stewards: [] };

  if (!memberSheet) return result;

  var data = memberSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isSteward !== true && isSteward !== 'Yes' && isSteward !== 'TRUE') continue;

    var firstName = data[i][MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = data[i][MEMBER_COLS.LAST_NAME - 1] || '';

    result.stewards.push({
      name: firstName + ' ' + lastName,
      location: data[i][MEMBER_COLS.LOCATION - 1] || 'Not specified',
      officeDays: data[i][MEMBER_COLS.OFFICE_DAYS - 1] || 'Contact for availability',
      email: data[i][MEMBER_COLS.EMAIL - 1] || 'Contact union office'
    });
  }

  // Sort by name
  result.stewards.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return result;
}

// ============================================================================
// STEWARD COVERAGE STATISTICS
// ============================================================================

/**
 * Gets steward coverage ratio for progress tracking
 * @returns {Object} Coverage statistics
 */
function getStewardCoverageStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || memberSheet.getLastRow() < 2) {
    return { ratio: 0, stewardCount: 0, memberCount: 0, targetRatio: 15 };
  }

  var data = memberSheet.getDataRange().getValues();
  var stewardCount = 0;
  var memberCount = 0;

  for (var i = 1; i < data.length; i++) {
    var memberId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    memberCount++;
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isSteward === true || isSteward === 'Yes' || isSteward === 'TRUE') {
      stewardCount++;
    }
  }

  // Calculate ratio as members per steward (lower is better coverage)
  var ratio = stewardCount > 0 ? Math.round(memberCount / stewardCount) : 0;
  var targetRatio = 15; // Target: 1 steward per 15 members
  var coveragePercent = ratio > 0 ? Math.min(100, Math.round((targetRatio / ratio) * 100)) : 0;

  return {
    ratio: ratio,
    stewardCount: stewardCount,
    memberCount: memberCount,
    targetRatio: targetRatio,
    coveragePercent: coveragePercent
  };
}
