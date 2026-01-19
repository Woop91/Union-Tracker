/**
 * ============================================================================
 * 509 STRATEGIC COMMAND CENTER - MEMBER PORTAL SERVICE (v4.5.0)
 * ============================================================================
 * Personalized member portal with Weingarten Rights and steward access.
 *
 * NOTE: The main dashboard functionality has been consolidated into
 * 04_UIService.gs (build509UnifiedDashboard). This file now focuses on:
 * - Personalized member portals (accessed via ?id=MEMBER_ID)
 * - Public portal without member ID
 * - PII safety utilities
 *
 * @fileoverview Member portal service with PII protection
 * @version 4.5.0
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
// DATA FETCHING FUNCTIONS - Portal Support
// ============================================================================

/**
 * Gets grievance statistics for portal display (no PII)
 * @returns {Object} Statistics object with winRate
 */
function getSecureGrievanceStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { total: 0, open: 0, won: 0, pending: 0, resolved: 0, winRate: 0 };
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
      case 'Denied':
        stats.denied++;
        stats.resolved++;
        break;
      case 'Settled':
        stats.won++;
        stats.resolved++;
        break;
      case 'Withdrawn':
        stats.resolved++;
        break;
      default:
        if (status.indexOf('Pending') >= 0) {
          stats.pending++;
        }
    }
  }

  stats.winRate = stats.resolved > 0 ? Math.round((stats.won / stats.resolved) * 100) : 0;
  return stats;
}

/**
 * Gets all stewards for portal display (public info only)
 * @returns {Array} Array of steward objects
 */
function getSecureAllStewards_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var stewards = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      stewards.push({
        'First Name': data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        'Last Name': data[i][MEMBER_COLS.LAST_NAME - 1] || '',
        'Unit': data[i][MEMBER_COLS.UNIT - 1] || 'General',
        'Work Location': data[i][MEMBER_COLS.WORK_LOCATION - 1] || ''
      });
    }
  }

  return stewards;
}

/**
 * Gets satisfaction statistics for portal display
 * @returns {Object} Satisfaction stats with avgTrust
 */
function getSecureSatisfactionStats_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.SATISFACTION);

  var result = {
    avgTrust: 0,
    avgSatisfaction: 0,
    responseCount: 0,
    recentTrend: 'stable'
  };

  if (!sheet || sheet.getLastRow() < 2) return result;

  try {
    var data = sheet.getDataRange().getValues();
    var trustScores = [];
    var satScores = [];

    for (var i = 1; i < data.length; i++) {
      var trustVal = parseFloat(data[i][7]); // SATISFACTION_COLS.Q7_TRUST_UNION - 1
      var satVal = parseFloat(data[i][6]);   // SATISFACTION_COLS.Q6_SATISFIED_REP - 1

      if (!isNaN(trustVal) && trustVal >= 1 && trustVal <= 10) {
        trustScores.push(trustVal);
      }
      if (!isNaN(satVal) && satVal >= 1 && satVal <= 10) {
        satScores.push(satVal);
      }
    }

    if (trustScores.length > 0) {
      result.avgTrust = Math.round((trustScores.reduce(function(a,b){return a+b;}, 0) / trustScores.length) * 10) / 10;
    }
    if (satScores.length > 0) {
      result.avgSatisfaction = Math.round((satScores.reduce(function(a,b){return a+b;}, 0) / satScores.length) * 10) / 10;
    }
    result.responseCount = Math.max(trustScores.length, satScores.length);

    // Trend analysis (compare last 10 vs previous 10)
    if (trustScores.length >= 20) {
      var recent = trustScores.slice(-10).reduce(function(a,b){return a+b;}, 0) / 10;
      var previous = trustScores.slice(-20, -10).reduce(function(a,b){return a+b;}, 0) / 10;
      result.recentTrend = recent > previous + 0.2 ? 'improving' : (recent < previous - 0.2 ? 'declining' : 'stable');
    }
  } catch (e) {
    Logger.log('Error in getSecureSatisfactionStats_: ' + e.message);
  }

  return result;
}

/**
 * Gets steward workload data
 * @returns {Array} Array of workload objects
 */
function getStewardWorkload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  var workload = {};

  // Get all stewards first
  var memberData = memberSheet.getDataRange().getValues();
  for (var i = 1; i < memberData.length; i++) {
    if (memberData[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes') {
      var name = (memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                 (memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      workload[name.trim()] = { name: name.trim(), openCases: 0, totalCases: 0 };
    }
  }

  // Count grievances per steward
  var gData = grievanceSheet.getDataRange().getValues();
  for (var g = 1; g < gData.length; g++) {
    var steward = gData[g][GRIEVANCE_COLS.STEWARD - 1];
    var status = gData[g][GRIEVANCE_COLS.STATUS - 1];

    if (steward && workload[steward]) {
      workload[steward].totalCases++;
      if (status === 'Open' || (status && status.indexOf('Pending') >= 0)) {
        workload[steward].openCases++;
      }
    }
  }

  return Object.keys(workload).map(function(key) { return workload[key]; });
}

/**
 * Gets contract PDF URL from config
 * @returns {string} Contract PDF URL
 */
function getContractPdfUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var url = configSheet.getRange(3, CONFIG_COLS.CONTRACT_URL).getValue();
      if (url) return url;
    }
  } catch (e) {}
  return '#';
}

/**
 * Gets resource Drive folder URL from config
 * @returns {string} Resource folder URL
 */
function getResourceDriveUrl_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var folderId = configSheet.getRange(3, CONFIG_COLS.ARCHIVE_FOLDER_ID).getValue();
      if (folderId) return 'https://drive.google.com/drive/folders/' + folderId;
    }
  } catch (e) {}
  return '#';
}

// ============================================================================
// PORTAL MENU WRAPPERS
// ============================================================================

/**
 * Menu wrapper: Build portal for the currently selected member
 * Gets the member ID from the active row in Member Directory
 */
function buildPortalForSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Build Member Portal',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Build Member Portal',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();

  if (!memberId) {
    ui.alert('Build Member Portal',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  var portal = buildMemberPortal(memberId);
  ui.showModalDialog(portal, '509 Member Portal');
}

/**
 * Menu wrapper: Send portal email to the currently selected member
 */
function sendPortalEmailToSelectedMember() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Send Portal Email',
      'Please select a member row in the Member Directory sheet first.',
      ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Send Portal Email',
      'Please select a data row (not the header).',
      ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();
  var memberEmail = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
  var firstName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue();

  if (!memberId) {
    ui.alert('Send Portal Email',
      'No Member ID found in the selected row.',
      ui.ButtonSet.OK);
    return;
  }

  if (!memberEmail) {
    ui.alert('Send Portal Email',
      'No email address found for this member.',
      ui.ButtonSet.OK);
    return;
  }

  sendMemberDashboardEmail(memberId);
}

// ============================================================================
// PORTAL BUILDERS
// ============================================================================

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
        duesPaying: true,
        isSteward: data[i][MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
        volunteerHours: parseFloat(data[i][MEMBER_COLS.VOLUNTEER_HOURS - 1]) || 0
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

// ============================================================================
// PORTAL HTML GENERATORS
// ============================================================================

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
