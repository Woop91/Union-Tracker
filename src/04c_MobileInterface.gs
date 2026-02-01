/**
 * ============================================================================
 * MobileInterface.gs - Mobile UI Components and Dashboard
 * ============================================================================
 *
 * This module handles all mobile-related functions including:
 * - Mobile context detection
 * - Mobile-optimized dashboard views
 * - Mobile grievance list and search interfaces
 * - Touch-friendly UI components
 *
 * REFACTORED: Split from 04_UIService.gs for better maintainability
 *
 * @fileoverview Mobile interface components and dashboard functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// MOBILE CONFIGURATION
// ============================================================================

/**
 * Mobile interface configuration settings
 * Note: This duplicates the config from 04_UIService.gs for module independence
 * TODO: Consider moving to a shared constants file after full extraction
 */
var MOBILE_CONFIG = {
  MAX_COLUMNS_MOBILE: 8,
  CARD_LAYOUT_ENABLED: true,
  TOUCH_TARGET_SIZE: '44px',
  MOBILE_BREAKPOINT: 768,  // Width in pixels below which is considered mobile
  TABLET_BREAKPOINT: 1024  // Width in pixels below which is considered tablet
};

// ============================================================================
// MOBILE CONTEXT DETECTION
// ============================================================================

/**
 * Checks if the current context is a mobile device
 * Server-side detection is limited; this function exists for potential
 * future use with session properties or client-side communication
 * @returns {boolean} Always returns false on server-side
 */
function isMobileContext() {
  // Server-side we can't reliably detect mobile
  // This function exists for potential future use with session properties
  return false;
}

// ============================================================================
// MOBILE DASHBOARD
// ============================================================================

/**
 * Shows the mobile-optimized dashboard interface
 * Features:
 * - Touch-friendly stat cards
 * - Quick action buttons
 * - Responsive layout
 * @returns {void}
 */
function showMobileDashboard() {
  var stats = getMobileDashboardStats();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"><style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;padding:0;margin:0;background:#f5f5f5}.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px}.header h1{margin:0;font-size:24px}.container{padding:15px}.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:20px}.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center}.stat-value{font-size:32px;font-weight:bold;color:#1a73e8}.stat-label{font-size:13px;color:#666;text-transform:uppercase}.section-title{font-size:16px;font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}.action-btn{background:white;border:none;padding:16px;margin-bottom:10px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}.action-btn:active{transform:scale(0.98)}.action-icon{font-size:24px;width:40px;height:40px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px}.action-label{font-weight:500}.action-desc{font-size:12px;color:#666}.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer}</style></head><body><div class="header"><h1>📱 509 Dashboard</h1><div style="font-size:14px;opacity:0.9">Mobile View</div></div><div class="container"><div class="stats"><div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div><div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div><div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div><div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div></div><div class="section-title">⚡ Quick Actions</div><button class="action-btn" onclick="google.script.run.showMobileGrievanceList()"><div class="action-icon">📋</div><div><div class="action-label">View Grievances</div><div class="action-desc">Browse all grievances</div></div></button><button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()"><div class="action-icon">🔍</div><div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div></button><button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()"><div class="action-icon">👤</div><div><div class="action-label">My Cases</div><div class="action-desc">View assigned grievances</div></div></button></div><button class="fab" onclick="location.reload()">🔄</button></body></html>'
  ).setWidth(400).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Mobile Dashboard');
}

/**
 * Gets dashboard statistics for mobile display
 * @returns {Object} Statistics object with totalGrievances, activeGrievances, pendingGrievances, overdueGrievances
 */
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
    if (status && status !== 'Resolved' && status !== 'Withdrawn') stats.activeGrievances++;
    if (status === 'Pending Info') stats.pendingGrievances++;
    if ((daysTo === 'Overdue' || (typeof daysTo === 'number' && daysTo < 0)) && status === 'Open') stats.overdueGrievances++;
  });
  return stats;
}

// ============================================================================
// MOBILE GRIEVANCE DATA
// ============================================================================

/**
 * Gets recent grievances formatted for mobile display
 * @param {number} [limit=5] Maximum number of grievances to return
 * @returns {Array<Object>} Array of grievance objects with mobile-friendly properties
 */
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

// ============================================================================
// MOBILE GRIEVANCE LIST
// ============================================================================

/**
 * Shows mobile-optimized grievance list interface
 * Features:
 * - Card-based layout
 * - Status filters
 * - Search functionality
 * - Responsive grid
 * @returns {void}
 */
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

// ============================================================================
// MOBILE UNIFIED SEARCH
// ============================================================================

/**
 * Shows mobile-optimized unified search interface
 * Features:
 * - Tabbed search (All, Members, Grievances)
 * - Real-time search results
 * - Touch-friendly result cards
 * @returns {void}
 */
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

/**
 * Gets search results for mobile search interface
 * @param {string} query Search query string
 * @param {string} tab Active tab ('all', 'members', or 'grievances')
 * @returns {Array<Object>} Array of search result objects
 */
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

// ============================================================================
// MY ASSIGNED GRIEVANCES
// ============================================================================

/**
 * Shows grievances assigned to the current user
 * Displays a quick summary dialog of assigned cases
 * @returns {void}
 */
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
