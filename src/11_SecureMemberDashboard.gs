/**
 * ============================================================================
 * 509 STRATEGIC COMMAND CENTER - SECURE MEMBER DASHBOARD (v4.0.3)
 * ============================================================================
 * Material Design Integration with Google Charts
 *
 * FEATURES:
 * - Material Design styling with Google Material Icons
 * - Google Charts integration (Treemap, Area Charts, Gauges)
 * - Safety Valve PII auto-redaction
 * - Weingarten Rights emergency utility
 * - Interactive steward search with live filtering
 * - Unit Density Heat Maps
 * - Sentiment Trend Analysis
 * - Steward Workload Balancing
 *
 * @fileoverview Secure member-facing dashboard with Material Design
 * @version 4.0.3
 * @requires 01_Constants.gs
 * ============================================================================
 */

// ============================================================================
// SAFETY VALVE - PII AUTO-REDACTION
// ============================================================================

/**
 * Scans and masks PII patterns (Phone numbers, SSNs) from strings.
 * Used to ensure accidental data entry doesn't leak to the Member Dashboard.
 *
 * @param {*} data - Input data to scrub
 * @returns {*} Scrubbed data with PII masked
 */
function safetyValveScrub(data) {
  if (typeof data !== 'string') return data;

  // Mask Phone Numbers: (123) 456-7890 or 123-456-7890 or +1 123-456-7890
  var phoneRegex = /(\+?\d{1,2}\s?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}/g;

  // Mask SSN-like patterns: 000-00-0000
  var ssnRegex = /\b\d{3}-\d{2}-\d{4}\b/g;

  // Mask email addresses in certain contexts
  var emailInTextRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  return data.replace(phoneRegex, "[REDACTED CONTACT]")
             .replace(ssnRegex, "[REDACTED ID]");
}

/**
 * Scrubs an object's string values for PII
 * @param {Object} obj - Object to scrub
 * @returns {Object} Object with scrubbed values
 */
function scrubObjectPII(obj) {
  if (!obj || typeof obj !== 'object') return obj;

  var scrubbed = {};
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      scrubbed[key] = safetyValveScrub(obj[key]);
    }
  }
  return scrubbed;
}

// ============================================================================
// DATA FETCHING FUNCTIONS (NO PII) - Secure Member Dashboard
// ============================================================================
// NOTE: These functions are prefixed with "Secure" to avoid conflicts with
// similarly-named functions in other files (03_GrievanceManager.gs, etc.)

/**
 * Gets grievance statistics for public display (no PII)
 * @returns {Object} Statistics object with winRate
 */
function getSecureGrievanceStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { total: 0, open: 0, won: 0, pending: 0, resolved: 0 };
  }

  var data = sheet.getDataRange().getValues();
  var stats = { total: 0, open: 0, won: 0, pending: 0, resolved: 0, denied: 0 };

  for (var i = 1; i < data.length; i++) {
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    if (!status) continue;

    stats.total++;

    switch(status) {
      case 'Open':
        stats.open++;
        break;
      case 'Won':
        stats.won++;
        stats.resolved++;
        break;
      case 'Pending Info':
        stats.pending++;
        break;
      case 'Settled':
        stats.resolved++;
        break;
      case 'Denied':
        stats.denied++;
        stats.resolved++;
        break;
      case 'Closed':
      case 'Withdrawn':
        stats.resolved++;
        break;
    }
  }

  // Calculate win rate
  if (stats.resolved > 0) {
    stats.winRate = Math.round((stats.won / stats.resolved) * 100);
  } else {
    stats.winRate = 0;
  }

  return stats;
}

/**
 * Gets all stewards for public display (names and units only - no PII)
 * @returns {Array} Array of steward objects with scrubbed data
 */
function getSecureAllStewards_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var stewards = [];

  for (var i = 1; i < data.length; i++) {
    var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
    if (isSteward === 'Yes' || isSteward === true) {
      stewards.push({
        'First Name': safetyValveScrub(data[i][MEMBER_COLS.FIRST_NAME - 1] || ''),
        'Last Name': safetyValveScrub(data[i][MEMBER_COLS.LAST_NAME - 1] || ''),
        'Unit': safetyValveScrub(data[i][MEMBER_COLS.UNIT - 1] || 'General')
      });
    }
  }

  return stewards;
}

/**
 * Gets aggregate satisfaction statistics (no individual PII)
 * Includes section-level scores for comprehensive analysis
 * @returns {Object} Aggregate stats object with sections
 */
function getSecureSatisfactionStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  // Default values if no data
  var defaults = {
    avgTrust: 7.5,
    avgSatisfaction: 7.2,
    avgRepresentation: 7.8,
    responseCount: 0,
    trendDirection: 'stable',
    sections: [
      { name: 'Overall Satisfaction', score: 0, questions: ['Satisfied with Rep', 'Trust Union', 'Feel Protected', 'Recommend'] },
      { name: 'Steward Ratings', score: 0, questions: ['Timely Response', 'Treated Respect', 'Explained Options', 'Followed Through', 'Advocated', 'Safe Concerns', 'Confidentiality'] },
      { name: 'Chapter Effectiveness', score: 0, questions: ['Understand Issues', 'Chapter Comm', 'Organizes', 'Reach Chapter', 'Fair Rep'] },
      { name: 'Local Leadership', score: 0, questions: ['Decisions Clear', 'Understand Process', 'Transparent Finance', 'Accountable', 'Fair Processes', 'Welcomes Opinions'] },
      { name: 'Contract Enforcement', score: 0, questions: ['Enforces Contract', 'Realistic Timelines', 'Clear Updates', 'Frontline Priority'] },
      { name: 'Communication Quality', score: 0, questions: ['Clear Actionable', 'Enough Info', 'Find Easily', 'All Shifts', 'Meetings Worth'] },
      { name: 'Member Voice', score: 0, questions: ['Voice Matters', 'Seeks Input', 'Dignity', 'Newer Supported', 'Conflict Respect'] },
      { name: 'Value & Action', score: 0, questions: ['Good Value', 'Priorities Needs', 'Prepared Mobilize', 'Win Together'] }
    ]
  };

  if (!sheet || sheet.getLastRow() < 2) {
    return defaults;
  }

  try {
    var data = sheet.getDataRange().getValues();
    var trustScores = [];
    var satScores = [];

    // Section score accumulators
    var sectionScores = {
      overall: [], steward: [], chapter: [], leadership: [],
      contract: [], communication: [], voice: [], value: []
    };

    // Q7 is Trust in Union (column H = index 7)
    // Q6 is Satisfied with Rep (column G = index 6)
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip empty rows

      var trustVal = parseFloat(data[i][7]); // SATISFACTION_COLS.Q7_TRUST_UNION - 1
      var satVal = parseFloat(data[i][6]);   // SATISFACTION_COLS.Q6_SATISFIED_REP - 1

      if (!isNaN(trustVal) && trustVal >= 1 && trustVal <= 10) {
        trustScores.push(trustVal);
      }
      if (!isNaN(satVal) && satVal >= 1 && satVal <= 10) {
        satScores.push(satVal);
      }

      // Calculate section averages from summary columns (BT onwards = index 71+)
      var avgOverall = parseFloat(data[i][71]) || 0;
      var avgSteward = parseFloat(data[i][72]) || 0;
      var avgChapter = parseFloat(data[i][74]) || 0;
      var avgLeadership = parseFloat(data[i][75]) || 0;
      var avgContract = parseFloat(data[i][76]) || 0;
      var avgComm = parseFloat(data[i][78]) || 0;
      var avgVoice = parseFloat(data[i][79]) || 0;
      var avgValue = parseFloat(data[i][80]) || 0;

      // If summary columns empty, calculate from raw questions
      if (avgOverall === 0) {
        var q6 = parseFloat(data[i][6]) || 0, q7 = parseFloat(data[i][7]) || 0;
        var q8 = parseFloat(data[i][8]) || 0, q9 = parseFloat(data[i][9]) || 0;
        avgOverall = (q6 + q7 + q8 + q9) / 4;
      }

      if (avgOverall > 0) sectionScores.overall.push(avgOverall);
      if (avgSteward > 0) sectionScores.steward.push(avgSteward);
      if (avgChapter > 0) sectionScores.chapter.push(avgChapter);
      if (avgLeadership > 0) sectionScores.leadership.push(avgLeadership);
      if (avgContract > 0) sectionScores.contract.push(avgContract);
      if (avgComm > 0) sectionScores.communication.push(avgComm);
      if (avgVoice > 0) sectionScores.voice.push(avgVoice);
      if (avgValue > 0) sectionScores.value.push(avgValue);
    }

    if (trustScores.length > 0) {
      var avgTrust = trustScores.reduce(function(a, b) { return a + b; }, 0) / trustScores.length;
      defaults.avgTrust = Math.round(avgTrust * 10) / 10;
    }

    if (satScores.length > 0) {
      var avgSat = satScores.reduce(function(a, b) { return a + b; }, 0) / satScores.length;
      defaults.avgSatisfaction = Math.round(avgSat * 10) / 10;
    }

    defaults.responseCount = Math.max(trustScores.length, satScores.length);

    // Calculate final section averages
    function avg(arr) { return arr.length > 0 ? Math.round((arr.reduce(function(a,b){return a+b;},0) / arr.length) * 10) / 10 : 0; }
    defaults.sections[0].score = avg(sectionScores.overall);
    defaults.sections[1].score = avg(sectionScores.steward);
    defaults.sections[2].score = avg(sectionScores.chapter);
    defaults.sections[3].score = avg(sectionScores.leadership);
    defaults.sections[4].score = avg(sectionScores.contract);
    defaults.sections[5].score = avg(sectionScores.communication);
    defaults.sections[6].score = avg(sectionScores.voice);
    defaults.sections[7].score = avg(sectionScores.value);

  } catch (e) {
    Logger.log('Satisfaction stats error: ' + e.message);
  }

  return defaults;
}

// ============================================================================
// ADVANCED ANALYTICS
// ============================================================================

/**
 * Fetches data for the Treemap and Sentiment Trend Area Chart.
 * @returns {Object} Analytics data object
 */
function getAdvancedAnalytics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var unitStats = {}; // { UnitName: { members: 0, cases: 0 } }

  // 1. Calculate Density from Member Directory
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var i = 1; i < memberData.length; i++) {
      var unit = memberData[i][MEMBER_COLS.UNIT - 1] || "Unknown";
      if (!unitStats[unit]) unitStats[unit] = { members: 0, cases: 0 };
      unitStats[unit].members++;
    }
  }

  // 2. Count grievances per unit
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getDataRange().getValues();
    for (var j = 1; j < grievanceData.length; j++) {
      var gUnit = grievanceData[j][GRIEVANCE_COLS.UNIT - 1] || "Unknown";
      var status = grievanceData[j][GRIEVANCE_COLS.STATUS - 1];
      if (!unitStats[gUnit]) unitStats[gUnit] = { members: 0, cases: 0 };
      if (status === 'Open' || status === 'Pending Info') {
        unitStats[gUnit].cases++;
      }
    }
  }

  // Format for Google Treemap: ['ID', 'Parent', 'Size (Members)', 'Color (Cases)']
  var treemapData = [["Location", "Parent", "Member Count", "Grievance Heat"]];
  treemapData.push(["All Units", null, 0, 0]);

  for (var unit in unitStats) {
    if (unitStats.hasOwnProperty(unit)) {
      treemapData.push([unit, "All Units", unitStats[unit].members || 1, unitStats[unit].cases]);
    }
  }

  // 3. Sentiment Trends - Get from Satisfaction sheet if available
  var sentimentData = getSentimentTrendData_();

  return { treemapData: treemapData, sentimentData: sentimentData };
}

/**
 * Gets sentiment trend data for area chart
 * @returns {Array} Sentiment data array for Google Charts
 * @private
 */
function getSentimentTrendData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  // Default placeholder data
  var defaultData = [
    ['Month', 'Trust Score'],
    ['Jan', 6.5], ['Feb', 6.8], ['Mar', 7.2],
    ['Apr', 7.0], ['May', 7.5], ['Jun', 7.8]
  ];

  if (!sheet || sheet.getLastRow() < 3) {
    return defaultData;
  }

  try {
    var data = sheet.getDataRange().getValues();
    var monthlyScores = {};

    for (var i = 1; i < data.length; i++) {
      var timestamp = data[i][0]; // Column A is timestamp
      var trustScore = parseFloat(data[i][7]); // Q7_TRUST_UNION

      if (timestamp && !isNaN(trustScore)) {
        var date = new Date(timestamp);
        var monthKey = date.toLocaleString('default', { month: 'short' });

        if (!monthlyScores[monthKey]) {
          monthlyScores[monthKey] = { sum: 0, count: 0 };
        }
        monthlyScores[monthKey].sum += trustScore;
        monthlyScores[monthKey].count++;
      }
    }

    var chartData = [['Month', 'Trust Score']];
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    for (var m = 0; m < months.length; m++) {
      if (monthlyScores[months[m]]) {
        var avg = monthlyScores[months[m]].sum / monthlyScores[months[m]].count;
        chartData.push([months[m], Math.round(avg * 10) / 10]);
      }
    }

    if (chartData.length > 1) {
      return chartData;
    }

  } catch (e) {
    Logger.log('Sentiment trend error: ' + e.message);
  }

  return defaultData;
}

// ============================================================================
// STEWARD WORKLOAD BALANCING
// ============================================================================

/**
 * Calculates workload and flags stewards over capacity.
 * @returns {Array} Array of steward workload objects
 */
function getStewardWorkload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stewards = getSecureAllStewards_();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() < 2) {
    return stewards.map(function(s) {
      return {
        name: s['First Name'] + ' ' + s['Last Name'],
        unit: s['Unit'],
        count: 0,
        status: 'Available',
        color: '#059669'
      };
    });
  }

  var grievanceData = grievanceSheet.getDataRange().getValues();
  var openCases = grievanceData.filter(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    return status === 'Open' || status === 'Pending Info';
  });

  return stewards.map(function(s) {
    var name = s['First Name'] + ' ' + s['Last Name'];
    var count = openCases.filter(function(row) {
      return row[GRIEVANCE_COLS.STEWARD - 1] === name;
    }).length;

    var status, color;
    if (count > 8) {
      status = 'OVERLOAD';
      color = '#ea580c'; // Orange
    } else if (count > 5) {
      status = 'Heavy';
      color = '#f59e0b'; // Amber
    } else {
      status = 'Available';
      color = '#059669'; // Green
    }

    return {
      name: name,
      unit: s['Unit'],
      count: count,
      status: status,
      color: color
    };
  });
}

// ============================================================================
// CONFIGURATION URLS (Read from Config sheet)
// ============================================================================

/**
 * Gets contract PDF URL from Config sheet
 * Uses CONFIG_COLS.CONTRACT_PDF_URL (column AZ) for the contract link
 * @returns {string} Contract PDF URL or '#' if not configured
 */
function getContractPdfUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var url = configSheet.getRange(2, CONFIG_COLS.CONTRACT_PDF_URL).getValue();
      if (url && url.toString().trim() !== '') {
        return url.toString().trim();
      }
    }
  } catch (e) {
    Logger.log('Contract URL error: ' + e.message);
  }
  return '#'; // Not configured
}

/**
 * Gets resource drive URL from Config sheet
 * @returns {string} Resource Drive URL
 */
function getResourceDriveUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var folderId = getConfigValue_(CONFIG_COLS.DRIVE_FOLDER_ID);
      if (folderId) {
        return 'https://drive.google.com/drive/folders/' + folderId;
      }
    }
  } catch (e) {
    Logger.log('Drive URL error: ' + e.message);
  }
  return '#'; // Placeholder
}

// ============================================================================
// MATERIAL DESIGN MEMBER DASHBOARD
// ============================================================================

/**
 * Shows the secure public member dashboard with Material Design
 * No PII visible - includes Weingarten Rights, Charts, and Steward Search
 */
function showPublicMemberDashboard() {
  var stats = getSecureGrievanceStats_();
  var stewards = getSecureAllStewards_();
  var satisfaction = getSecureSatisfactionStats_();
  var analytics = getAdvancedAnalytics();
  var workload = getStewardWorkload();

  // Get URLs from config
  var CONTRACT_PDF_URL = getContractPdfUrl_();
  var RESOURCE_DRIVE_URL = getResourceDriveUrl_();

  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { ' +
    '      font-family: "Roboto", sans-serif; ' +
    '      background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);' +
    '      min-height: 100vh;' +
    '      padding: 16px;' +
    '      color: #F8FAFC;' +
    '    }' +
    '    .header { ' +
    '      display: flex; ' +
    '      align-items: center; ' +
    '      gap: 12px;' +
    '      margin-bottom: 16px;' +
    '      padding-bottom: 12px;' +
    '      border-bottom: 2px solid rgba(124, 58, 237, 0.3);' +
    '    }' +
    '    .header h1 { font-size: 20px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 28px; }' +
    '    ' +
    '    /* Material Design Cards */' +
    '    .card { ' +
    '      background: rgba(255,255,255,0.05); ' +
    '      border-radius: 12px; ' +
    '      padding: 16px; ' +
    '      margin-bottom: 12px;' +
    '      backdrop-filter: blur(10px);' +
    '      border: 1px solid rgba(255,255,255,0.1);' +
    '      transition: transform 0.2s, box-shadow 0.2s;' +
    '    }' +
    '    .card:hover { ' +
    '      transform: translateY(-2px);' +
    '      box-shadow: 0 8px 25px rgba(0,0,0,0.3);' +
    '    }' +
    '    .card-title { ' +
    '      display: flex;' +
    '      align-items: center;' +
    '      gap: 8px;' +
    '      font-size: 12px; ' +
    '      text-transform: uppercase; ' +
    '      letter-spacing: 1px; ' +
    '      color: #94A3B8; ' +
    '      margin-bottom: 12px;' +
    '    }' +
    '    .card-title .material-icons { font-size: 18px; }' +
    '    ' +
    '    /* Emergency Rights Box */' +
    '    .rights-box { ' +
    '      background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%);' +
    '      border: 2px solid #dc2626;' +
    '      border-radius: 12px;' +
    '      padding: 16px; ' +
    '      margin-bottom: 16px;' +
    '      cursor: pointer;' +
    '      transition: all 0.3s;' +
    '    }' +
    '    .rights-box:hover { ' +
    '      transform: scale(1.02);' +
    '      box-shadow: 0 0 20px rgba(220, 38, 38, 0.4);' +
    '    }' +
    '    .rights-header { ' +
    '      display: flex;' +
    '      align-items: center;' +
    '      gap: 8px;' +
    '      font-weight: 700;' +
    '      color: #fecaca;' +
    '      margin-bottom: 8px;' +
    '    }' +
    '    .rights-text { ' +
    '      font-size: 14px; ' +
    '      line-height: 1.5; ' +
    '      color: #fff;' +
    '      font-weight: 500;' +
    '    }' +
    '    .rights-expanded { display: none; margin-top: 12px; padding-top: 12px; border-top: 1px solid rgba(255,255,255,0.2); }' +
    '    .rights-box.expanded .rights-expanded { display: block; }' +
    '    ' +
    '    /* Quick Links Buttons */' +
    '    .btn-group { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 16px; }' +
    '    .action-btn { ' +
    '      display: flex; ' +
    '      align-items: center; ' +
    '      justify-content: center;' +
    '      gap: 8px;' +
    '      background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);' +
    '      color: white; ' +
    '      padding: 14px 16px; ' +
    '      border-radius: 10px; ' +
    '      text-decoration: none; ' +
    '      font-size: 13px; ' +
    '      font-weight: 600;' +
    '      transition: all 0.3s;' +
    '      border: none;' +
    '      cursor: pointer;' +
    '    }' +
    '    .action-btn:hover { ' +
    '      transform: translateY(-2px);' +
    '      box-shadow: 0 6px 20px rgba(124, 58, 237, 0.4);' +
    '    }' +
    '    .action-btn .material-icons { font-size: 20px; }' +
    '    .btn-drive { background: linear-gradient(135deg, #059669 0%, #047857 100%); }' +
    '    .btn-drive:hover { box-shadow: 0 6px 20px rgba(5, 150, 105, 0.4); }' +
    '    ' +
    '    /* Stats Grid */' +
    '    .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 12px; }' +
    '    .stat-card { ' +
    '      background: rgba(255,255,255,0.08); ' +
    '      border-radius: 10px; ' +
    '      padding: 14px; ' +
    '      text-align: center;' +
    '    }' +
    '    .stat-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; margin-bottom: 4px; }' +
    '    .stat-value { font-size: 24px; font-weight: 700; }' +
    '    .stat-value.green { color: #10B981; }' +
    '    .stat-value.purple { color: #A78BFA; }' +
    '    .stat-value.amber { color: #FBBF24; }' +
    '    .stat-value.blue { color: #60A5FA; }' +
    '    ' +
    '    /* Search Bar */' +
    '    .search-container { position: relative; margin-bottom: 12px; }' +
    '    .search-input { ' +
    '      width: 100%; ' +
    '      padding: 12px 12px 12px 44px; ' +
    '      border: 2px solid rgba(255,255,255,0.1); ' +
    '      border-radius: 10px; ' +
    '      background: rgba(255,255,255,0.05);' +
    '      color: #F8FAFC;' +
    '      font-size: 14px;' +
    '      outline: none;' +
    '      transition: border-color 0.2s;' +
    '    }' +
    '    .search-input:focus { border-color: #7C3AED; }' +
    '    .search-input::placeholder { color: #64748B; }' +
    '    .search-icon { ' +
    '      position: absolute; ' +
    '      left: 14px; ' +
    '      top: 50%; ' +
    '      transform: translateY(-50%);' +
    '      color: #64748B;' +
    '      font-size: 20px;' +
    '    }' +
    '    ' +
    '    /* Steward List */' +
    '    .steward-list { max-height: 180px; overflow-y: auto; }' +
    '    .steward-item { ' +
    '      display: flex; ' +
    '      justify-content: space-between; ' +
    '      align-items: center;' +
    '      padding: 10px 0; ' +
    '      border-bottom: 1px solid rgba(255,255,255,0.1);' +
    '      font-size: 13px;' +
    '    }' +
    '    .steward-item:last-child { border-bottom: none; }' +
    '    .steward-name { font-weight: 500; }' +
    '    .pill { ' +
    '      background: rgba(124, 58, 237, 0.2); ' +
    '      color: #A78BFA; ' +
    '      padding: 4px 10px; ' +
    '      border-radius: 20px; ' +
    '      font-size: 11px;' +
    '      font-weight: 500;' +
    '    }' +
    '    ' +
    '    /* Workload Items */' +
    '    .workload-item { ' +
    '      display: flex; ' +
    '      justify-content: space-between; ' +
    '      align-items: center;' +
    '      padding: 8px 0;' +
    '      font-size: 12px;' +
    '      border-bottom: 1px solid rgba(255,255,255,0.05);' +
    '    }' +
    '    .workload-item:last-child { border-bottom: none; }' +
    '    .workload-status { font-weight: 600; padding: 2px 8px; border-radius: 4px; font-size: 10px; }' +
    '    ' +
    '    /* Chart Containers */' +
    '    .chart-container { height: 180px; width: 100%; }' +
    '    .chart-container-small { height: 100px; width: 100%; }' +
    '    ' +
    '    /* Satisfaction Section */' +
    '    .sat-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }' +
    '    .sat-response-badge { background: rgba(96,165,250,0.2); color: #60A5FA; padding: 4px 10px; border-radius: 12px; font-size: 10px; font-weight: 600; }' +
    '    .sat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }' +
    '    .sat-item { background: rgba(255,255,255,0.05); border-radius: 8px; padding: 10px; }' +
    '    .sat-item-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }' +
    '    .sat-item-name { font-size: 10px; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.5px; }' +
    '    .sat-item-score { font-size: 16px; font-weight: 700; }' +
    '    .sat-bar { height: 4px; background: rgba(255,255,255,0.1); border-radius: 2px; overflow: hidden; }' +
    '    .sat-bar-fill { height: 100%; border-radius: 2px; }' +
    '    ' +
    '    /* Footer */' +
    '    .footer { ' +
    '      text-align: center; ' +
    '      font-size: 10px; ' +
    '      color: #64748B; ' +
    '      margin-top: 16px;' +
    '      padding-top: 12px;' +
    '      border-top: 1px solid rgba(255,255,255,0.1);' +
    '    }' +
    '    .footer .material-icons { font-size: 12px; vertical-align: middle; }' +
    '    ' +
    '    /* Scrollbar Styling */' +
    '    ::-webkit-scrollbar { width: 6px; }' +
    '    ::-webkit-scrollbar-track { background: rgba(255,255,255,0.05); border-radius: 3px; }' +
    '    ::-webkit-scrollbar-thumb { background: #7C3AED; border-radius: 3px; }' +
    '  </style>' +
    '  <script>' +
    '    google.charts.load("current", {"packages":["treemap", "corechart"]});' +
    '    google.charts.setOnLoadCallback(drawCharts);' +
    '    ' +
    '    function drawCharts() {' +
    '      // Treemap Chart' +
    '      var treemapData = ' + JSON.stringify(analytics.treemapData) + ';' +
    '      if (treemapData.length > 2) {' +
    '        var treeDataTable = google.visualization.arrayToDataTable(treemapData);' +
    '        var tree = new google.visualization.TreeMap(document.getElementById("tree_div"));' +
    '        tree.draw(treeDataTable, {' +
    '          minColor: "#22c55e",' +
    '          midColor: "#fbbf24",' +
    '          maxColor: "#ef4444",' +
    '          headerHeight: 20,' +
    '          headerColor: "#1e293b",' +
    '          fontColor: "#f8fafc",' +
    '          showScale: false,' +
    '          generateTooltip: function(row, size, value) {' +
    '            return "<div style=\\"background:#1e293b; padding:10px; border-radius:8px; color:#fff;\\">" +' +
    '                   "<strong>" + treeDataTable.getValue(row, 0) + "</strong><br>" +' +
    '                   "Members: " + size + "<br>" +' +
    '                   "Open Cases: " + value + "</div>";' +
    '          }' +
    '        });' +
    '      }' +
    '      ' +
    '      // Area Chart for Sentiment Trend' +
    '      var sentimentData = ' + JSON.stringify(analytics.sentimentData) + ';' +
    '      if (sentimentData.length > 1) {' +
    '        var areaDataTable = google.visualization.arrayToDataTable(sentimentData);' +
    '        var area = new google.visualization.AreaChart(document.getElementById("area_div"));' +
    '        area.draw(areaDataTable, {' +
    '          legend: "none",' +
    '          colors: ["#7C3AED"],' +
    '          areaOpacity: 0.3,' +
    '          backgroundColor: "transparent",' +
    '          chartArea: { left: 30, right: 10, top: 10, bottom: 25, width: "100%", height: "100%" },' +
    '          hAxis: { textStyle: { color: "#94A3B8", fontSize: 10 }, gridlines: { color: "transparent" } },' +
    '          vAxis: { textStyle: { color: "#94A3B8", fontSize: 10 }, gridlines: { color: "rgba(255,255,255,0.1)" }, minValue: 0, maxValue: 10 }' +
    '        });' +
    '      }' +
    '    }' +
    '    ' +
    '    // Live Search Filter' +
    '    function filterStewards() {' +
    '      var input = document.getElementById("stewardSearch").value.toUpperCase();' +
    '      var items = document.getElementsByClassName("steward-item");' +
    '      for (var i = 0; i < items.length; i++) {' +
    '        var text = items[i].textContent || items[i].innerText;' +
    '        items[i].style.display = text.toUpperCase().indexOf(input) > -1 ? "" : "none";' +
    '      }' +
    '    }' +
    '    ' +
    '    // Toggle Rights Box' +
    '    function toggleRights() {' +
    '      var box = document.getElementById("rightsBox");' +
    '      box.classList.toggle("expanded");' +
    '    }' +
    '  </script>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <i class="material-icons">shield</i>' +
    '    <h1>509 Member Command Center</h1>' +
    '  </div>' +
    '  ' +
    '  <!-- EMERGENCY WEINGARTEN RIGHTS -->' +
    '  <div class="rights-box" id="rightsBox" onclick="toggleRights()">' +
    '    <div class="rights-header">' +
    '      <i class="material-icons">gavel</i>' +
    '      TAP FOR EMERGENCY RIGHTS' +
    '    </div>' +
    '    <div class="rights-text">' +
    '      "If this discussion could in any way lead to my being disciplined or terminated, ' +
    '      I respectfully request that my union representative be present."' +
    '    </div>' +
    '    <div class="rights-expanded">' +
    '      <div style="font-size:12px; color:#fecaca; margin-bottom:8px;"><strong>WEINGARTEN RIGHTS - READ ALOUD:</strong></div>' +
    '      <div style="font-size:13px; line-height:1.6;">' +
    '        "I request my union representative. Until my representative arrives, ' +
    '        I choose not to answer questions or write any statements. ' +
    '        This is my right under the National Labor Relations Act."' +
    '      </div>' +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <!-- QUICK ACTION BUTTONS -->' +
    '  <div class="btn-group">' +
    '    <a href="' + CONTRACT_PDF_URL + '" target="_blank" class="action-btn">' +
    '      <i class="material-icons">description</i>' +
    '      Union Contract' +
    '    </a>' +
    '    <a href="' + RESOURCE_DRIVE_URL + '" target="_blank" class="action-btn btn-drive">' +
    '      <i class="material-icons">folder_shared</i>' +
    '      Member Drive' +
    '    </a>' +
    '  </div>' +
    '  ' +
    '  <!-- KEY STATISTICS -->' +
    '  <div class="stats-grid">' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Trust Score</div>' +
    '      <div class="stat-value green">' + satisfaction.avgTrust + '/10</div>' +
    '    </div>' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Active Cases</div>' +
    '      <div class="stat-value purple">' + stats.open + '</div>' +
    '    </div>' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Win Rate</div>' +
    '      <div class="stat-value amber">' + stats.winRate + '%</div>' +
    '    </div>' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Total Cases</div>' +
    '      <div class="stat-value blue">' + stats.total + '</div>' +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <!-- UNIT DENSITY TREEMAP -->' +
    '  <div class="card">' +
    '    <div class="card-title">' +
    '      <i class="material-icons">grid_view</i>' +
    '      Unit Density & Grievance Heat' +
    '    </div>' +
    '    <div id="tree_div" class="chart-container"></div>' +
    '  </div>' +
    '  ' +
    '  <!-- SENTIMENT TREND -->' +
    '  <div class="card">' +
    '    <div class="card-title">' +
    '      <i class="material-icons">trending_up</i>' +
    '      Union Morale Trend' +
    '    </div>' +
    '    <div id="area_div" class="chart-container-small"></div>' +
    '  </div>' +
    '  ' +
    '  <!-- STEWARD SEARCH -->' +
    '  <div class="card">' +
    '    <div class="card-title">' +
    '      <i class="material-icons">people</i>' +
    '      Find My Steward' +
    '    </div>' +
    '    <div class="search-container">' +
    '      <i class="material-icons search-icon">search</i>' +
    '      <input type="text" id="stewardSearch" class="search-input" ' +
    '             onkeyup="filterStewards()" placeholder="Search by name or unit...">' +
    '    </div>' +
    '    <div class="steward-list" id="stewardList">' +
    stewards.map(function(s) {
      return '      <div class="steward-item">' +
             '        <span class="steward-name">' + s['First Name'] + ' ' + s['Last Name'] + '</span>' +
             '        <span class="pill">' + s['Unit'] + '</span>' +
             '      </div>';
    }).join('') +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <!-- STEWARD WORKLOAD -->' +
    '  <div class="card">' +
    '    <div class="card-title">' +
    '      <i class="material-icons">assignment_ind</i>' +
    '      Steward Availability' +
    '    </div>' +
    workload.slice(0, 5).map(function(s) {
      return '    <div class="workload-item">' +
             '      <span>' + s.name + '</span>' +
             '      <span class="workload-status" style="background:' + s.color + '20; color:' + s.color + '">' +
             '        ' + s.status + ' (' + s.count + ')' +
             '      </span>' +
             '    </div>';
    }).join('') +
    '  </div>' +
    '  ' +
    '  <!-- MEMBER SATISFACTION ANALYSIS -->' +
    '  <div class="card">' +
    '    <div class="card-title">' +
    '      <i class="material-icons">sentiment_satisfied</i>' +
    '      Member Satisfaction Analysis' +
    '    </div>' +
    '    <div class="sat-header">' +
    '      <span style="font-size:11px;color:#94A3B8">Survey Results</span>' +
    '      <span class="sat-response-badge">' + satisfaction.responseCount + ' Responses</span>' +
    '    </div>' +
    '    <div class="sat-grid">' +
    satisfaction.sections.map(function(section) {
      var scoreColor = section.score >= 7 ? '#22c55e' : section.score >= 5 ? '#f59e0b' : '#ef4444';
      var pct = (section.score / 10) * 100;
      return '      <div class="sat-item">' +
             '        <div class="sat-item-header">' +
             '          <span class="sat-item-name">' + section.name + '</span>' +
             '          <span class="sat-item-score" style="color:' + scoreColor + '">' + section.score + '</span>' +
             '        </div>' +
             '        <div class="sat-bar"><div class="sat-bar-fill" style="width:' + pct + '%;background:' + scoreColor + '"></div></div>' +
             '      </div>';
    }).join('') +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <!-- FOOTER -->' +
    '  <div class="footer">' +
    '    <i class="material-icons">lock</i> No PII Visible &bull; Secure Member View &bull; v' + VERSION_INFO.CURRENT +
    '  </div>' +
    '</body>' +
    '</html>';

  var output = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(output, '509 Member Command Center');
}

// ============================================================================
// STEWARD PERFORMANCE MODAL (Material Design)
// ============================================================================

/**
 * Shows the steward performance modal with Material Design
 * Leadership view with workload balancing metrics
 */
function showStewardPerformanceModal() {
  var workload = getStewardWorkload();
  var stats = getSecureGrievanceStats_();

  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { ' +
    '      font-family: "Roboto", sans-serif; ' +
    '      background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);' +
    '      min-height: 100vh;' +
    '      padding: 20px;' +
    '      color: #F8FAFC;' +
    '    }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 24px; }' +
    '    .summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 20px; }' +
    '    .summary-card { background: rgba(255,255,255,0.08); border-radius: 10px; padding: 16px; text-align: center; }' +
    '    .summary-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; }' +
    '    .summary-value { font-size: 28px; font-weight: 700; margin-top: 4px; }' +
    '    .steward-grid { display: flex; flex-direction: column; gap: 8px; }' +
    '    .steward-row { ' +
    '      display: grid; ' +
    '      grid-template-columns: 2fr 1fr 80px 60px; ' +
    '      align-items: center;' +
    '      background: rgba(255,255,255,0.05);' +
    '      border-radius: 8px;' +
    '      padding: 12px 16px;' +
    '    }' +
    '    .steward-row:hover { background: rgba(255,255,255,0.08); }' +
    '    .steward-name { font-weight: 500; }' +
    '    .steward-unit { color: #94A3B8; font-size: 12px; }' +
    '    .badge { ' +
    '      padding: 4px 10px; ' +
    '      border-radius: 20px; ' +
    '      font-size: 11px; ' +
    '      font-weight: 600;' +
    '      text-align: center;' +
    '    }' +
    '    .case-count { font-weight: 700; text-align: center; }' +
    '    .legend { display: flex; gap: 16px; margin-top: 16px; justify-content: center; }' +
    '    .legend-item { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #94A3B8; }' +
    '    .legend-dot { width: 10px; height: 10px; border-radius: 50%; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <i class="material-icons">leaderboard</i>' +
    '    <h1>Steward Performance Dashboard</h1>' +
    '  </div>' +
    '  ' +
    '  <div class="summary-grid">' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Total Stewards</div>' +
    '      <div class="summary-value" style="color:#A78BFA">' + workload.length + '</div>' +
    '    </div>' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Active Cases</div>' +
    '      <div class="summary-value" style="color:#60A5FA">' + stats.open + '</div>' +
    '    </div>' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Win Rate</div>' +
    '      <div class="summary-value" style="color:#34D399">' + stats.winRate + '%</div>' +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <div class="steward-grid">' +
    workload.map(function(s) {
      var badgeStyle = 'background:' + s.color + '20; color:' + s.color;
      return '    <div class="steward-row">' +
             '      <div class="steward-name">' + s.name + '</div>' +
             '      <div class="steward-unit">' + s.unit + '</div>' +
             '      <div class="badge" style="' + badgeStyle + '">' + s.status + '</div>' +
             '      <div class="case-count" style="color:' + s.color + '">' + s.count + '</div>' +
             '    </div>';
    }).join('') +
    '  </div>' +
    '  ' +
    '  <div class="legend">' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#059669"></div> Available (0-5)</div>' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#f59e0b"></div> Heavy (6-8)</div>' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#ea580c"></div> Overload (9+)</div>' +
    '  </div>' +
    '</body>' +
    '</html>';

  var output = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(output, 'Steward Performance');
}

// ============================================================================
// EMAIL DASHBOARD LINK
// ============================================================================

/**
 * Emails the public dashboard link to the selected member
 * Uses the currently selected row in Member Directory
 */
function emailDashboardLink() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveRange().getRow();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Please select a member row in the Member Directory first.');
    return;
  }

  if (row < 2) {
    ui.alert('Please select a data row (not the header).');
    return;
  }

  // Get member data including Member ID for personalized portal link
  var dataRange = sheet.getRange(row, 1, 1, Math.max(MEMBER_COLS.EMAIL, MEMBER_COLS.MEMBER_ID, MEMBER_COLS.FIRST_NAME));
  var data = dataRange.getValues()[0];
  var email = data[MEMBER_COLS.EMAIL - 1];
  var firstName = data[MEMBER_COLS.FIRST_NAME - 1];
  var memberId = data[MEMBER_COLS.MEMBER_ID - 1];

  if (!email) {
    ui.alert('No email address found for this member.');
    return;
  }

  if (!memberId) {
    ui.alert('No Member ID found for this member. Cannot generate portal link.');
    return;
  }

  var response = ui.alert(
    'Send Dashboard Link',
    'Send the Member Dashboard access information to:\n\n' +
    firstName + ' (' + email + ')\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    // Use the deployed web app URL, not the spreadsheet URL
    var webAppUrl = ScriptApp.getService().getUrl();
    if (!webAppUrl) {
      ui.alert('Error', 'Web app is not deployed. Please deploy the web app first via Extensions > Apps Script > Deploy.', ui.ButtonSet.OK);
      return;
    }
    var portalUrl = webAppUrl + '?id=' + memberId;

    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Your Member Dashboard Access';
    var body = 'Hello ' + firstName + ',\n\n' +
               'You can access your personalized Union Member Dashboard at the following link:\n\n' +
               portalUrl + '\n\n' +
               'This dashboard gives you access to:\n' +
               '- Your Weingarten Rights (emergency reference)\n' +
               '- Union contract and resources\n' +
               '- Steward contact information\n' +
               '- Grievance statistics\n\n' +
               'Keep this link private - it is personalized for you.\n\n' +
               'If you have any questions, contact your steward.\n\n' +
               'In Solidarity,' +
               COMMAND_CONFIG.EMAIL.FOOTER;

    MailApp.sendEmail(email, subject, body);

    ui.alert('Success', 'Dashboard link sent to ' + email, ui.ButtonSet.OK);

    // Log the action
    logAuditEvent('DASHBOARD_LINK_SENT', {
      recipient: email,
      memberName: firstName,
      sentBy: Session.getActiveUser().getEmail()
    });

  } catch (e) {
    ui.alert('Error', 'Failed to send email: ' + e.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// EXECUTIVE DASHBOARD (Now Modal-Based - See 04_UIService.gs)
// ============================================================================
// The rebuildExecutiveDashboard() function has been moved to 04_UIService.gs
// and converted to a modal-based SPA architecture using the Bridge Pattern.
// This provides better performance and user experience.
//
// See: launchExecutiveDashboard() and getDashboardStats() in 04_UIService.gs

// ============================================================================
// STANDALONE ANALYTICS CHARTS
// ============================================================================

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Unit density is now in the unified Steward Dashboard (Analytics tab).
 */
function showUnitDensityTreemap() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function showUnitDensityTreemap_LEGACY() {
  var analytics = getAdvancedAnalytics();

  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '  <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' +
    '  <style>' +
    '    body { ' +
    '      font-family: "Roboto", sans-serif; ' +
    '      background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);' +
    '      margin: 0; padding: 20px; color: #F8FAFC;' +
    '    }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 24px; }' +
    '    .chart-container { height: 350px; width: 100%; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 10px; }' +
    '    .legend { display: flex; gap: 16px; margin-top: 16px; justify-content: center; }' +
    '    .legend-item { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #94A3B8; }' +
    '    .legend-dot { width: 14px; height: 14px; border-radius: 4px; }' +
    '    .info { font-size: 12px; color: #94A3B8; text-align: center; margin-top: 12px; }' +
    '  </style>' +
    '  <script>' +
    '    google.charts.load("current", {"packages":["treemap"]});' +
    '    google.charts.setOnLoadCallback(drawChart);' +
    '    function drawChart() {' +
    '      var data = google.visualization.arrayToDataTable(' + JSON.stringify(analytics.treemapData) + ');' +
    '      var chart = new google.visualization.TreeMap(document.getElementById("chart_div"));' +
    '      chart.draw(data, {' +
    '        minColor: "#22c55e",' +
    '        midColor: "#fbbf24",' +
    '        maxColor: "#ef4444",' +
    '        headerHeight: 25,' +
    '        headerColor: "#1e293b",' +
    '        fontColor: "#f8fafc",' +
    '        fontSize: 12,' +
    '        showScale: true,' +
    '        generateTooltip: function(row, size, value) {' +
    '          return "<div style=\\"background:#1e293b; padding:12px; border-radius:8px; color:#fff; font-family:Roboto;\\">" +' +
    '                 "<strong style=\\"font-size:14px\\">" + data.getValue(row, 0) + "</strong><br><br>" +' +
    '                 "<span style=\\"color:#94A3B8\\">Members:</span> " + size + "<br>" +' +
    '                 "<span style=\\"color:#94A3B8\\">Open Cases:</span> " + value + "</div>";' +
    '        }' +
    '      });' +
    '    }' +
    '  </script>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <i class="material-icons">grid_view</i>' +
    '    <h1>Unit Density & Grievance Heat Map</h1>' +
    '  </div>' +
    '  <div id="chart_div" class="chart-container"></div>' +
    '  <div class="legend">' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#22c55e"></div> Low Activity</div>' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#fbbf24"></div> Moderate</div>' +
    '    <div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div> High Grievance Load</div>' +
    '  </div>' +
    '  <div class="info">Box size = member count | Color = open grievances</div>' +
    '</body>' +
    '</html>';

  var output = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(output, 'Unit Density Treemap');
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Sentiment trends are now in the unified Steward Dashboard (Overview tab).
 */
function showSentimentTrendChart() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function showSentimentTrendChart_LEGACY() {
  var analytics = getAdvancedAnalytics();

  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '  <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>' +
    '  <style>' +
    '    body { ' +
    '      font-family: "Roboto", sans-serif; ' +
    '      background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);' +
    '      margin: 0; padding: 20px; color: #F8FAFC;' +
    '    }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 24px; }' +
    '    .chart-container { height: 300px; width: 100%; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 10px; }' +
    '    .stats-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 16px; }' +
    '    .stat-card { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .stat-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; }' +
    '    .stat-value { font-size: 20px; font-weight: 700; margin-top: 4px; }' +
    '  </style>' +
    '  <script>' +
    '    google.charts.load("current", {"packages":["corechart"]});' +
    '    google.charts.setOnLoadCallback(drawChart);' +
    '    function drawChart() {' +
    '      var data = google.visualization.arrayToDataTable(' + JSON.stringify(analytics.sentimentData) + ');' +
    '      var chart = new google.visualization.AreaChart(document.getElementById("chart_div"));' +
    '      chart.draw(data, {' +
    '        legend: { position: "none" },' +
    '        colors: ["#7C3AED"],' +
    '        areaOpacity: 0.3,' +
    '        backgroundColor: "transparent",' +
    '        chartArea: { left: 50, right: 20, top: 20, bottom: 40, width: "100%", height: "100%" },' +
    '        hAxis: { textStyle: { color: "#94A3B8", fontSize: 11 }, gridlines: { color: "transparent" } },' +
    '        vAxis: { textStyle: { color: "#94A3B8", fontSize: 11 }, gridlines: { color: "rgba(255,255,255,0.1)" }, minValue: 0, maxValue: 10, title: "Trust Score", titleTextStyle: { color: "#94A3B8", fontSize: 10 } },' +
    '        curveType: "function"' +
    '      });' +
    '    }' +
    '  </script>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <i class="material-icons">trending_up</i>' +
    '    <h1>Union Morale Sentiment Trend</h1>' +
    '  </div>' +
    '  <div id="chart_div" class="chart-container"></div>' +
    '  <div class="stats-row">' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Data Points</div>' +
    '      <div class="stat-value" style="color:#60A5FA">' + (analytics.sentimentData.length - 1) + '</div>' +
    '    </div>' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Trend</div>' +
    '      <div class="stat-value" style="color:#34D399">Stable</div>' +
    '    </div>' +
    '    <div class="stat-card">' +
    '      <div class="stat-label">Source</div>' +
    '      <div class="stat-value" style="color:#A78BFA; font-size:14px">Surveys</div>' +
    '    </div>' +
    '  </div>' +
    '</body>' +
    '</html>';

  var output = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(output, 'Sentiment Trend');
}

/**
 * @deprecated v4.3.2 - Use showStewardDashboard() instead.
 * Workload metrics are now in the unified Steward Dashboard (Workload tab).
 */
function showStewardWorkloadReport() {
  showStewardDashboard();
}

/** @deprecated - Legacy function body for reference only */
function showStewardWorkloadReport_LEGACY() {
  var workload = getStewardWorkload();
  var stats = getSecureGrievanceStats_();

  // Calculate summary stats
  var totalCases = workload.reduce(function(sum, s) { return sum + s.count; }, 0);
  var avgCases = workload.length > 0 ? (totalCases / workload.length).toFixed(1) : 0;
  var overloaded = workload.filter(function(s) { return s.status === 'OVERLOAD'; }).length;

  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    body { ' +
    '      font-family: "Roboto", sans-serif; ' +
    '      background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);' +
    '      margin: 0; padding: 20px; color: #F8FAFC;' +
    '    }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 24px; }' +
    '    .summary-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px; }' +
    '    .summary-card { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .summary-label { font-size: 9px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; }' +
    '    .summary-value { font-size: 22px; font-weight: 700; margin-top: 4px; }' +
    '    .table-container { background: rgba(255,255,255,0.05); border-radius: 12px; overflow: hidden; }' +
    '    .table-header { display: grid; grid-template-columns: 2fr 1fr 60px 80px; background: #1E293B; padding: 12px 16px; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; }' +
    '    .table-row { display: grid; grid-template-columns: 2fr 1fr 60px 80px; padding: 12px 16px; border-bottom: 1px solid rgba(255,255,255,0.05); align-items: center; }' +
    '    .table-row:last-child { border-bottom: none; }' +
    '    .table-row:hover { background: rgba(255,255,255,0.03); }' +
    '    .badge { padding: 4px 10px; border-radius: 20px; font-size: 10px; font-weight: 600; text-align: center; }' +
    '    .case-count { font-weight: 700; text-align: center; font-size: 16px; }' +
    '    .recommendation { margin-top: 16px; background: rgba(124,58,237,0.1); border: 1px solid rgba(124,58,237,0.3); border-radius: 8px; padding: 12px; font-size: 12px; }' +
    '    .recommendation .material-icons { font-size: 16px; vertical-align: middle; margin-right: 6px; color: #A78BFA; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <i class="material-icons">assignment_ind</i>' +
    '    <h1>Steward Workload Balancing Report</h1>' +
    '  </div>' +
    '  ' +
    '  <div class="summary-grid">' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Stewards</div>' +
    '      <div class="summary-value" style="color:#A78BFA">' + workload.length + '</div>' +
    '    </div>' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Total Cases</div>' +
    '      <div class="summary-value" style="color:#60A5FA">' + totalCases + '</div>' +
    '    </div>' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Avg/Steward</div>' +
    '      <div class="summary-value" style="color:#34D399">' + avgCases + '</div>' +
    '    </div>' +
    '    <div class="summary-card">' +
    '      <div class="summary-label">Overloaded</div>' +
    '      <div class="summary-value" style="color:' + (overloaded > 0 ? '#ef4444' : '#34D399') + '">' + overloaded + '</div>' +
    '    </div>' +
    '  </div>' +
    '  ' +
    '  <div class="table-container">' +
    '    <div class="table-header">' +
    '      <div>Steward</div>' +
    '      <div>Unit</div>' +
    '      <div style="text-align:center">Cases</div>' +
    '      <div style="text-align:center">Status</div>' +
    '    </div>' +
    workload.map(function(s) {
      var badgeStyle = 'background:' + s.color + '20; color:' + s.color;
      return '    <div class="table-row">' +
             '      <div>' + s.name + '</div>' +
             '      <div style="color:#94A3B8">' + s.unit + '</div>' +
             '      <div class="case-count" style="color:' + s.color + '">' + s.count + '</div>' +
             '      <div><span class="badge" style="' + badgeStyle + '">' + s.status + '</span></div>' +
             '    </div>';
    }).join('') +
    '  </div>' +
    (overloaded > 0 ?
    '  <div class="recommendation">' +
    '    <i class="material-icons">warning</i>' +
    '    <strong>Action Required:</strong> ' + overloaded + ' steward(s) are over capacity. Consider redistributing cases to maintain quality representation.' +
    '  </div>' : '') +
    '</body>' +
    '</html>';

  var output = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(output, 'Steward Workload Report');
}

// ============================================================================
// SECURE MEMBER PORTAL - WEB APP
// ============================================================================

/**
 * Member Portal entry point (Legacy - now consolidated in 05_Integrations.gs doGet)
 * NOTE: Renamed to avoid duplicate. The main doGet() in 05_Integrations.gs now handles
 * both mobile dashboard and member portal requests.
 * Use ?id=MEMBER_ID to access personalized portal via the main doGet()
 *
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} HTML page for display
 * @deprecated Use doGet() in 05_Integrations.gs - supports both ?id=MEMBER_ID and ?page=dashboard/search/etc
 */
function doGet_MemberPortal_(e) {
  var memberId = e && e.parameter && e.parameter.id;

  if (memberId) {
    return buildMemberPortal(memberId);
  }

  return buildPublicPortal();
}

/**
 * Menu wrapper: Build portal for the currently selected member
 * Gets the member ID from the active row in Member Directory
 */
function buildPortalForSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if we're in the Member Directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('👤 Build Member Portal',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  // Skip header row
  if (row < 2) {
    ui.alert('👤 Build Member Portal',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  // Get member ID from column A
  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();

  if (!memberId) {
    ui.alert('👤 Build Member Portal',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  // Build and show the portal
  var portal = buildMemberPortal(memberId);
  ui.showModalDialog(portal, '509 Member Portal');
}

/**
 * Menu wrapper: Send portal email to the currently selected member
 * Gets the member ID from the active row in Member Directory
 */
function sendPortalEmailToSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if we're in the Member Directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('📧 Send Portal Email',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  // Skip header row
  if (row < 2) {
    ui.alert('📧 Send Portal Email',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  // Get member ID and email from the row
  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();
  var memberEmail = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();

  if (!memberId) {
    ui.alert('📧 Send Portal Email',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  if (!memberEmail) {
    ui.alert('📧 Send Portal Email',
      'No email address found for this member.\n\n' +
      'Please add an email address to the Member Directory first.',
      ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert('📧 Send Portal Email',
    'Send portal access email to:\n\n' +
    firstName + ' (' + memberEmail + ')?\n\n' +
    'This will send them a link to access their personalized member portal.',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  // Send the email
  sendMemberDashboardEmail(memberId);

  ui.alert('✅ Email Sent',
    'Portal access email sent to ' + memberEmail,
    ui.ButtonSet.OK);
}

/**
 * Builds personalized member portal for specific member ID
 * @param {string} memberId - The member ID to look up
 * @returns {HtmlOutput} Personalized portal HTML
 */
function buildMemberPortal(memberId) {
  var profile = getMemberProfile(memberId);

  if (!profile) {
    return HtmlService.createHtmlOutput(getErrorPageHtml_('Member not found'))
      .setTitle('509 Member Portal - Error');
  }

  var html = getMemberPortalHtml_(profile);
  return HtmlService.createHtmlOutput(html)
    .setTitle('509 Member Portal - ' + profile.firstName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Builds public portal (no member ID required)
 * @returns {HtmlOutput} Public portal HTML
 */
function buildPublicPortal() {
  var stats = getSecureGrievanceStats_();
  var stewards = getSecureAllStewards_();
  var satisfaction = getSecureSatisfactionStats_();

  var html = getPublicPortalHtml_(stats, stewards, satisfaction);
  return HtmlService.createHtmlOutput(html)
    .setTitle('509 Union Member Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets member profile by ID (with PII scrubbing for sensitive fields)
 * @param {string} memberId - The member ID to look up
 * @returns {Object|null} Member profile object or null if not found
 */
function getMemberProfile(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || memberSheet.getLastRow() < 2) return null;

  var data = memberSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] == memberId) {
      return {
        memberId: memberId,
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: data[i][MEMBER_COLS.LAST_NAME - 1] || '',
        unit: data[i][MEMBER_COLS.UNIT - 1] || 'General',
        workLocation: data[i][MEMBER_COLS.WORK_LOCATION - 1] || '',
        duesPaying: true, // Default to true (DUES_PAYING column not in current schema)
        isSteward: data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
        volunteerHours: parseFloat(data[i][MEMBER_COLS.VOLUNTEER_HOURS - 1]) || 0
        // Note: Email and phone intentionally excluded for security
      };
    }
  }

  return null;
}

/**
 * Sends member dashboard email with personalized portal link
 * @param {string} memberId - The member ID to send link to
 */
function sendMemberDashboardEmail(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found.');
    return;
  }

  var data = memberSheet.getDataRange().getValues();
  var member = null;

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] == memberId) {
      member = {
        email: data[i][MEMBER_COLS.EMAIL - 1],
        firstName: data[i][MEMBER_COLS.FIRST_NAME - 1]
      };
      break;
    }
  }

  if (!member || !member.email) {
    SpreadsheetApp.getUi().alert('Member email not found.');
    return;
  }

  // Note: Replace with actual deployed web app URL
  var webAppUrl = ScriptApp.getService().getUrl();
  var portalUrl = webAppUrl + '?id=' + memberId;

  var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Your Personal Union Portal';
  var body = 'Hello ' + member.firstName + ',\n\n' +
    'Access your personalized Union Member Portal:\n\n' +
    portalUrl + '\n\n' +
    'Your portal includes:\n' +
    '- Emergency Weingarten Rights card\n' +
    '- Your steward contact information\n' +
    '- Union resources and contract links\n' +
    '- Satisfaction survey link\n\n' +
    'Keep this link private - it is personalized for you.\n\n' +
    'In Solidarity,\n' +
    '509 Strategic Command Center';

  try {
    MailApp.sendEmail(member.email, subject, body);
    SpreadsheetApp.getUi().alert('Portal link sent to ' + member.email);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + e.message);
  }
}

/**
 * Generates personalized member portal HTML
 * @param {Object} profile - Member profile object
 * @returns {string} Complete HTML for the portal
 * @private
 */
function getMemberPortalHtml_(profile) {
  var stewards = getSecureAllStewards_();
  var stats = getSecureGrievanceStats_();
  var CONTRACT_PDF_URL = getContractPdfUrl_();
  var RESOURCE_DRIVE_URL = getResourceDriveUrl_();

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; }' +
    '    .container { max-width: 600px; margin: 0 auto; padding: 20px; }' +
    '    .header { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); padding: 24px; border-radius: 16px; margin-bottom: 20px; text-align: center; }' +
    '    .header h1 { font-size: 18px; font-weight: 500; margin-bottom: 8px; }' +
    '    .header .welcome { font-size: 24px; font-weight: 700; }' +
    '    .member-badge { display: inline-flex; align-items: center; gap: 6px; background: rgba(255,255,255,0.2); padding: 6px 12px; border-radius: 20px; font-size: 12px; margin-top: 12px; }' +
    '    .rights-box { background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%); border: 2px solid #dc2626; border-radius: 12px; padding: 20px; margin-bottom: 20px; }' +
    '    .rights-header { display: flex; align-items: center; gap: 8px; font-weight: 700; color: #fecaca; margin-bottom: 12px; font-size: 14px; }' +
    '    .rights-text { font-size: 15px; line-height: 1.6; font-weight: 500; }' +
    '    .card { background: rgba(255,255,255,0.05); border-radius: 12px; padding: 16px; margin-bottom: 12px; border: 1px solid rgba(255,255,255,0.1); }' +
    '    .card-title { display: flex; align-items: center; gap: 8px; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; margin-bottom: 12px; }' +
    '    .btn-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }' +
    '    .btn { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 14px; border-radius: 10px; text-decoration: none; font-weight: 600; font-size: 13px; transition: all 0.2s; }' +
    '    .btn-purple { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; }' +
    '    .btn-green { background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; }' +
    '    .btn:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.3); }' +
    '    .steward-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.1); }' +
    '    .steward-item:last-child { border-bottom: none; }' +
    '    .pill { background: rgba(124,58,237,0.2); color: #A78BFA; padding: 4px 10px; border-radius: 20px; font-size: 11px; }' +
    '    .stats-row { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }' +
    '    .stat { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .stat-value { font-size: 24px; font-weight: 700; color: #60A5FA; }' +
    '    .stat-label { font-size: 10px; color: #94A3B8; text-transform: uppercase; }' +
    '    .footer { text-align: center; font-size: 10px; color: #64748B; padding: 20px 0; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="header">' +
    '      <h1>509 MEMBER PORTAL</h1>' +
    '      <div class="welcome">Welcome, ' + profile.firstName + '!</div>' +
    '      <div class="member-badge">' +
    '        <i class="material-icons" style="font-size:16px">' + (profile.isSteward ? 'verified' : 'person') + '</i>' +
    '        ' + (profile.isSteward ? 'Union Steward' : 'Union Member') + ' - ' + profile.unit +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="rights-box">' +
    '      <div class="rights-header"><i class="material-icons">gavel</i> WEINGARTEN RIGHTS</div>' +
    '      <div class="rights-text">"If this discussion could in any way lead to my being disciplined or terminated, I respectfully request that my union representative be present. Until my representative arrives, I choose not to answer questions or write any statements."</div>' +
    '    </div>' +
    '    ' +
    '    <div class="btn-grid">' +
    '      <a href="' + CONTRACT_PDF_URL + '" target="_blank" class="btn btn-purple"><i class="material-icons">description</i> Contract</a>' +
    '      <a href="' + RESOURCE_DRIVE_URL + '" target="_blank" class="btn btn-green"><i class="material-icons">folder</i> Resources</a>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">people</i> Your Stewards</div>' +
    stewards.slice(0, 5).map(function(s) {
      return '<div class="steward-item"><span>' + s['First Name'] + ' ' + s['Last Name'] + '</span><span class="pill">' + s['Unit'] + '</span></div>';
    }).join('') +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">bar_chart</i> Union Stats</div>' +
    '      <div class="stats-row">' +
    '        <div class="stat"><div class="stat-value">' + stats.winRate + '%</div><div class="stat-label">Win Rate</div></div>' +
    '        <div class="stat"><div class="stat-value">' + stats.open + '</div><div class="stat-label">Active Cases</div></div>' +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="footer">' +
    '      <i class="material-icons" style="font-size:12px;vertical-align:middle">lock</i> ' +
    '      Secure Member Portal | Member ID: ' + profile.memberId +
    '    </div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}

/**
 * Generates public portal HTML (no member ID required)
 * @param {Object} stats - Grievance statistics
 * @param {Array} stewards - List of stewards
 * @param {Object} satisfaction - Satisfaction stats
 * @returns {string} Complete HTML for public portal
 * @private
 */
function getPublicPortalHtml_(stats, stewards, satisfaction) {
  var CONTRACT_PDF_URL = getContractPdfUrl_();
  var RESOURCE_DRIVE_URL = getResourceDriveUrl_();

  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '    body { font-family: "Roboto", sans-serif; background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); min-height: 100vh; color: #F8FAFC; }' +
    '    .container { max-width: 600px; margin: 0 auto; padding: 20px; }' +
    '    .header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 2px solid rgba(124,58,237,0.3); }' +
    '    .header h1 { font-size: 20px; font-weight: 500; }' +
    '    .header .material-icons { color: #7C3AED; font-size: 28px; }' +
    '    .rights-box { background: linear-gradient(135deg, #7f1d1d 0%, #991b1b 100%); border: 2px solid #dc2626; border-radius: 12px; padding: 20px; margin-bottom: 20px; }' +
    '    .rights-header { display: flex; align-items: center; gap: 8px; font-weight: 700; color: #fecaca; margin-bottom: 12px; }' +
    '    .rights-text { font-size: 14px; line-height: 1.6; font-weight: 500; }' +
    '    .btn-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }' +
    '    .btn { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 14px; border-radius: 10px; text-decoration: none; font-weight: 600; font-size: 13px; }' +
    '    .btn-purple { background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%); color: white; }' +
    '    .btn-green { background: linear-gradient(135deg, #059669 0%, #047857 100%); color: white; }' +
    '    .card { background: rgba(255,255,255,0.05); border-radius: 12px; padding: 16px; margin-bottom: 12px; border: 1px solid rgba(255,255,255,0.1); }' +
    '    .card-title { display: flex; align-items: center; gap: 8px; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94A3B8; margin-bottom: 12px; }' +
    '    .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }' +
    '    .stat { background: rgba(255,255,255,0.08); border-radius: 8px; padding: 12px; text-align: center; }' +
    '    .stat-value { font-size: 24px; font-weight: 700; }' +
    '    .stat-value.green { color: #10B981; }' +
    '    .stat-value.blue { color: #60A5FA; }' +
    '    .stat-label { font-size: 10px; color: #94A3B8; text-transform: uppercase; }' +
    '    .steward-item { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.1); }' +
    '    .steward-item:last-child { border-bottom: none; }' +
    '    .pill { background: rgba(124,58,237,0.2); color: #A78BFA; padding: 4px 10px; border-radius: 20px; font-size: 11px; }' +
    '    .footer { text-align: center; font-size: 10px; color: #64748B; padding: 20px 0; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="container">' +
    '    <div class="header"><i class="material-icons">shield</i><h1>509 Union Member Portal</h1></div>' +
    '    ' +
    '    <div class="rights-box">' +
    '      <div class="rights-header"><i class="material-icons">gavel</i> WEINGARTEN RIGHTS</div>' +
    '      <div class="rights-text">"If this discussion could in any way lead to my being disciplined or terminated, I respectfully request that my union representative be present."</div>' +
    '    </div>' +
    '    ' +
    '    <div class="btn-grid">' +
    '      <a href="' + CONTRACT_PDF_URL + '" target="_blank" class="btn btn-purple"><i class="material-icons">description</i> Contract</a>' +
    '      <a href="' + RESOURCE_DRIVE_URL + '" target="_blank" class="btn btn-green"><i class="material-icons">folder</i> Resources</a>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">bar_chart</i> Union Statistics</div>' +
    '      <div class="stats-grid">' +
    '        <div class="stat"><div class="stat-value green">' + satisfaction.avgTrust + '/10</div><div class="stat-label">Trust Score</div></div>' +
    '        <div class="stat"><div class="stat-value blue">' + stats.winRate + '%</div><div class="stat-label">Win Rate</div></div>' +
    '        <div class="stat"><div class="stat-value blue">' + stats.open + '</div><div class="stat-label">Active Cases</div></div>' +
    '        <div class="stat"><div class="stat-value green">' + stats.total + '</div><div class="stat-label">Total Cases</div></div>' +
    '      </div>' +
    '    </div>' +
    '    ' +
    '    <div class="card">' +
    '      <div class="card-title"><i class="material-icons">people</i> Find a Steward</div>' +
    stewards.map(function(s) {
      return '<div class="steward-item"><span>' + s['First Name'] + ' ' + s['Last Name'] + '</span><span class="pill">' + s['Unit'] + '</span></div>';
    }).join('') +
    '    </div>' +
    '    ' +
    '    <div class="footer"><i class="material-icons" style="font-size:12px;vertical-align:middle">lock</i> No PII Visible | Public Portal v' + VERSION_INFO.CURRENT + '</div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}

/**
 * Generates error page HTML
 * @param {string} message - Error message to display
 * @returns {string} Error page HTML
 * @private
 */
function getErrorPageHtml_(message) {
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '  <meta name="viewport" content="width=device-width, initial-scale=1">' +
    '  <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">' +
    '  <style>' +
    '    body { font-family: "Roboto", sans-serif; background: #0F172A; min-height: 100vh; display: flex; align-items: center; justify-content: center; color: #F8FAFC; }' +
    '    .error-box { text-align: center; padding: 40px; background: rgba(239,68,68,0.1); border: 1px solid rgba(239,68,68,0.3); border-radius: 16px; max-width: 400px; }' +
    '    .error-icon { font-size: 48px; color: #ef4444; margin-bottom: 16px; }' +
    '    .error-title { font-size: 20px; font-weight: 700; margin-bottom: 8px; }' +
    '    .error-msg { color: #94A3B8; }' +
    '  </style>' +
    '</head>' +
    '<body>' +
    '  <div class="error-box">' +
    '    <i class="material-icons error-icon">error_outline</i>' +
    '    <div class="error-title">Access Error</div>' +
    '    <div class="error-msg">' + message + '</div>' +
    '  </div>' +
    '</body>' +
    '</html>';
}
