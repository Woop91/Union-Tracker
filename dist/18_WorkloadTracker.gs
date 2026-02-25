/**
 * ============================================================================
 * 18_WorkloadTracker.gs - Member Workload Tracking Module
 * ============================================================================
 *
 * Simplified workload tracking -- 8 categories, sub-categories, checkboxes.
 * No authentication required. Submissions are anonymous.
 *
 * Sheet Structure:
 *   "Workload Vault"     (hidden)  -- raw submission data
 *   "Workload Reporting" (visible) -- anonymized ledger (identity -> REDACTED)
 *
 * @version 4.10.0
 * @requires 01_Core.gs (SHEETS)
 * @requires 06_Maintenance.gs (logAuditEvent)
 */

// ============================================================================
// WORKLOAD TRACKER CONFIGURATION
// ============================================================================

var WT_APP_CONFIG = {
  version: '4.11.0',
  name: 'Workload Tracker',
  maxStringLength: 255,
  lockTimeoutMs: 30000
};

// ============================================================================
// COLUMN CONSTANTS (0-indexed for array access)
// ============================================================================

var WT_VAULT_COLS = {
  TIMESTAMP:           0,
  PRIORITY_CASES:      1,
  PENDING_CASES:       2,
  UNREAD_DOCS:         3,
  TODO_ITEMS:          4,
  SENT_REFERRALS:      5,
  CE_ACTIVITIES:       6,
  ASSISTANCE_REQUESTS: 7,
  AGED_CASES:          8,
  SUB_CATEGORIES:      9,
  ON_PLAN:             10,
  OVERTIME_HOURS:      11
};

// ============================================================================
// SUB-CATEGORY DEFINITIONS
// ============================================================================

var WT_SUB_CATEGORIES = {
  priority:   ['QDD', 'CAL', 'TERI', 'Aged Case', 'Congressional', 'Dire Need', 'Homeless',
               'Presumptive Disability', 'COBRA', 'DDS Aged', 'SOAR Involvement',
               'MC/WW', '100% P&T', 'Public Inquiry', 'DDS Important (High)', 'DDS Important (Medium)'],
  pending:    ['New Cases', 'Immediate Action', 'No Activity 15+ Days',
               'No Activity 30+ Days', 'Internal QA Returns', 'Federal QA Returns', 'Approval Returns'],
  unread:     ['All Unread Documents', 'MER', 'CE Reports', 'Trailer Mail'],
  todo:       ['Follow-Ups Due', 'Unread Case Notes', 'Delivery Failures', 'Updates'],
  referrals:  ['MC/PC', 'General'],
  ce:         ['All CE Activities', 'Scheduled', 'Not Kept', 'Under Review'],
  assistance: ['Outbound'],
  aged:       ['60-89 Days', '90-119 Days', '120-179 Days', '180+ Days']
};

// ============================================================================
// INPUT SANITIZATION
// ============================================================================

function sanitizeString(input, maxLength) {
  maxLength = maxLength || WT_APP_CONFIG.maxStringLength;
  if (!input || typeof input !== 'string') return '';
  var sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
  return sanitized.length > maxLength ? sanitized.substring(0, maxLength) : sanitized;
}

// ============================================================================
// LOCK SERVICE
// ============================================================================

function withLock(fn, lockName) {
  lockName = lockName || 'operation';
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(WT_APP_CONFIG.lockTimeoutMs)) {
      throw new Error('System busy. Please try again in a moment.');
    }
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

function initWorkloadTrackerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault) {
    vault = ss.insertSheet(SHEETS.WORKLOAD_VAULT);
    vault.appendRow([
      'Timestamp',
      'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
      'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'Sub-Categories (JSON)', 'On Plan', 'Overtime Hours'
    ]);
    vault.getRange(1, 1, 1, 12)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    vault.setFrozenRows(1);
  }
  vault.hideSheet();

  var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
  if (!report) {
    report = ss.insertSheet(SHEETS.WORKLOAD_REPORTING);
    report.appendRow([
      'Date', 'Identity',
      'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
      'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'On Plan', 'Overtime Hours'
    ]);
    report.getRange(1, 1, 1, 12)
      .setFontWeight('bold')
      .setBackground('#0d47a1')
      .setFontColor('#ffffff');
    report.setFrozenRows(1);
  }
}
// ============================================================================
// FORM PROCESSING
// ============================================================================

function processWorkloadForm(formObj) {
  try {
    var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
    for (var n = 0; n < numFields.length; n++) {
      var val = Number(formObj[numFields[n]]) || 0;
      if (val < 0 || val > 999 || !Number.isInteger(val)) {
        return 'Error: Workload values must be whole numbers between 0 and 999.';
      }
    }

    return withLock(function() {
      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault) {
        return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
      }

      var sanitizeNum = function(v) {
        return Math.min(999, Math.max(0, Math.floor(Number(v) || 0)));
      };

      var catKeys = Object.keys(WT_SUB_CATEGORIES);
      var subCatData = {};
      for (var ci = 0; ci < catKeys.length; ci++) {
        subCatData[catKeys[ci]] = {};
        var subNames = WT_SUB_CATEGORIES[catKeys[ci]];
        for (var si = 0; si < subNames.length; si++) {
          var fieldName = 'sub_' + (ci + 1) + '_' + si;
          var subVal = sanitizeNum(formObj[fieldName]);
          if (subVal > 0) { subCatData[catKeys[ci]][subNames[si]] = subVal; }
        }
      }

      vault.appendRow([
        new Date(),
        sanitizeNum(formObj.t1),
        sanitizeNum(formObj.t2),
        sanitizeNum(formObj.t3),
        sanitizeNum(formObj.t4),
        sanitizeNum(formObj.t5),
        sanitizeNum(formObj.t6),
        sanitizeNum(formObj.t7),
        sanitizeNum(formObj.t8),
        JSON.stringify(subCatData),
        formObj.on_plan ? 'Yes' : 'No',
        formObj.overtime_enabled ? (Number(formObj.overtime_hours) || 0) : ''
      ]);

      if (typeof logAuditEvent === 'function') {
        logAuditEvent('WORKLOAD_SUBMIT', 'Anonymous workload submission recorded.');
      }

      refreshWorkloadReportingData();
      return 'Success! Your workload data has been recorded.';
    }, 'processWorkloadForm');

  } catch (err) {
    console.error('processWorkloadForm error:', err);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WORKLOAD_SUBMIT_ERROR', 'Error: ' + err.message);
    }
    return 'Error: Unable to process submission. Please try again.';
  }
}

// ============================================================================
// REPORTING REFRESH
// ============================================================================

function refreshWorkloadReportingData() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var vault  = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
  if (!vault || !report) return;

  var vaultData = vault.getDataRange().getValues();
  if (vaultData.length <= 1) return;

  var rows = [];
  for (var j = 1; j < vaultData.length; j++) {
    rows.push([
      vaultData[j][WT_VAULT_COLS.TIMESTAMP],
      'REDACTED',
      vaultData[j][WT_VAULT_COLS.PRIORITY_CASES],
      vaultData[j][WT_VAULT_COLS.PENDING_CASES],
      vaultData[j][WT_VAULT_COLS.UNREAD_DOCS],
      vaultData[j][WT_VAULT_COLS.TODO_ITEMS],
      vaultData[j][WT_VAULT_COLS.SENT_REFERRALS],
      vaultData[j][WT_VAULT_COLS.CE_ACTIVITIES],
      vaultData[j][WT_VAULT_COLS.ASSISTANCE_REQUESTS],
      vaultData[j][WT_VAULT_COLS.AGED_CASES],
      vaultData[j][WT_VAULT_COLS.ON_PLAN],
      vaultData[j][WT_VAULT_COLS.OVERTIME_HOURS]
    ]);
  }

  report.clearContents();
  report.appendRow(['Date', 'Identity', 'Priority Cases', 'Pending Cases', 'Unread Documents',
    'To-Do Items', 'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
    'On Plan', 'Overtime Hours']);
  report.getRange(1, 1, 1, 12)
    .setFontWeight('bold')
    .setBackground('#0d47a1')
    .setFontColor('#ffffff');
  if (rows.length > 0) {
    report.getRange(2, 1, rows.length, 12).setValues(rows);
  }
}

function refreshWorkloadLedger() {
  refreshWorkloadReportingData();
  SpreadsheetApp.getUi().alert('Workload ledger refreshed successfully.');
}

// ============================================================================
// DASHBOARD DATA (for web portal analytics)
// ============================================================================

/**
 * Returns collective workload stats after PIN authentication.
 * @param {string} email
 * @param {string} pin
 * @returns {Object} { success, data, message }
 */
function getWorkloadDashboardData(email, pin) {
  if (!email || !pin) {
    return { success: false, data: null, message: 'Credentials required.' };
  }
  var auth = authenticateWorkloadMember_(email, pin);
  if (!auth.success) {
    return { success: false, data: null, message: auth.message };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault || vault.getLastRow() <= 1) {
    return { success: true, data: { totalSubmissions: 0, members: 0, categories: {} }, message: '' };
  }

  var data = vault.getDataRange().getValues();
  var userSharingStart = wtGetUserSharingStartDate_(email.toLowerCase().trim());

  // Aggregate stats from non-private rows within sharing window
  var totals = { t1: 0, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0 };
  var members = new Set ? new Set() : {};
  var submissionCount = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_VAULT_COLS.PRIVACY] === 'Private') continue;
    var ts = new Date(data[i][WT_VAULT_COLS.TIMESTAMP]);
    if (userSharingStart && ts < userSharingStart) continue;

    submissionCount++;
    var rowEmail = data[i][WT_VAULT_COLS.EMAIL];
    if (rowEmail) {
      if (typeof Set !== 'undefined') { members.add(rowEmail.toString().toLowerCase()); }
      else { members[rowEmail.toString().toLowerCase()] = true; }
    }
    totals.t1 += Number(data[i][WT_VAULT_COLS.PRIORITY_CASES]) || 0;
    totals.t2 += Number(data[i][WT_VAULT_COLS.PENDING_CASES])  || 0;
    totals.t3 += Number(data[i][WT_VAULT_COLS.UNREAD_DOCS])    || 0;
    totals.t4 += Number(data[i][WT_VAULT_COLS.TODO_ITEMS])     || 0;
    totals.t5 += Number(data[i][WT_VAULT_COLS.SENT_REFERRALS]) || 0;
    totals.t6 += Number(data[i][WT_VAULT_COLS.CE_ACTIVITIES])  || 0;
    totals.t7 += Number(data[i][WT_VAULT_COLS.ASSISTANCE_REQUESTS]) || 0;
    totals.t8 += Number(data[i][WT_VAULT_COLS.AGED_CASES])     || 0;
  }

  var memberCount = typeof Set !== 'undefined' ? members.size : Object.keys(members).length;

  return {
    success: true,
    message: '',
    data: {
      totalSubmissions: submissionCount,
      members: memberCount,
      categories: {
        'Priority Cases':      totals.t1,
        'Pending Cases':       totals.t2,
        'Unread Documents':    totals.t3,
        'To-Do Items':         totals.t4,
        'Sent Referrals':      totals.t5,
        'CE Activities':       totals.t6,
        'Assistance Requests': totals.t7,
        'Aged Cases':          totals.t8
      },
      averages: submissionCount > 0 ? {
        'Priority Cases':      (totals.t1 / submissionCount).toFixed(1),
        'Pending Cases':       (totals.t2 / submissionCount).toFixed(1),
        'Unread Documents':    (totals.t3 / submissionCount).toFixed(1),
        'To-Do Items':         (totals.t4 / submissionCount).toFixed(1),
        'Sent Referrals':      (totals.t5 / submissionCount).toFixed(1),
        'CE Activities':       (totals.t6 / submissionCount).toFixed(1),
        'Assistance Requests': (totals.t7 / submissionCount).toFixed(1),
        'Aged Cases':          (totals.t8 / submissionCount).toFixed(1)
      } : {}
    }
  };
}

// ============================================================================
// USER HISTORY (for web portal)
// ============================================================================

/**
 * Returns the authenticated member's personal submission history.
 * @param {string} email
 * @param {string} pin
 * @returns {Object} { success, history:Array, message }
 */
function getWorkloadUserHistory(email, pin) {
  if (!email || !pin) {
    return { success: false, history: [], message: 'Credentials required.' };
  }
  var auth = authenticateWorkloadMember_(email, pin);
  if (!auth.success) {
    return { success: false, history: [], message: auth.message };
  }

  var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault || vault.getLastRow() <= 1) {
    return { success: true, history: [], message: '' };
  }

  var emailLower = email.toLowerCase().trim();
  var data = vault.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][WT_VAULT_COLS.EMAIL];
    if (!rowEmail || rowEmail.toString().toLowerCase().trim() !== emailLower) continue;
    history.push({
      date:       data[i][WT_VAULT_COLS.TIMESTAMP],
      t1:         data[i][WT_VAULT_COLS.PRIORITY_CASES],
      t2:         data[i][WT_VAULT_COLS.PENDING_CASES],
      t3:         data[i][WT_VAULT_COLS.UNREAD_DOCS],
      t4:         data[i][WT_VAULT_COLS.TODO_ITEMS],
      t5:         data[i][WT_VAULT_COLS.SENT_REFERRALS],
      t6:         data[i][WT_VAULT_COLS.CE_ACTIVITIES],
      t7:         data[i][WT_VAULT_COLS.ASSISTANCE_REQUESTS],
      t8:         data[i][WT_VAULT_COLS.AGED_CASES],
      weeklyCases: data[i][WT_VAULT_COLS.WEEKLY_CASES],
      employment: data[i][WT_VAULT_COLS.EMPLOYMENT_TYPE],
      privacy:    data[i][WT_VAULT_COLS.PRIVACY],
      onPlan:     data[i][WT_VAULT_COLS.ON_PLAN],
      overtime:   data[i][WT_VAULT_COLS.OVERTIME_HOURS]
    });
  }

  // Newest first
  history.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  return { success: true, history: history, message: '' };
}

// ============================================================================
// WEB PORTAL HTML
// ============================================================================

/**
 * Returns the Workload Tracker portal HTML for serving via doGet (?page=workload).
 * The HTML is embedded by build.js from src/WorkloadTracker.html.
 * @returns {string} Full HTML string
 */
function getWorkloadTrackerPortalHtml() {
  try {
    // After build, getWorkloadTracker_HTML() is auto-generated by build.js from WorkloadTracker.html
    if (typeof getWorkloadTracker_HTML === 'function') {
      return getWorkloadTracker_HTML();
    }
    // Fallback for development (before build)
    return HtmlService.createHtmlOutputFromFile('WorkloadTracker').getContent();
  } catch (err) {
    return '<html><body style="font-family:sans-serif;padding:2rem;">' +
      '<h2>Workload Tracker</h2>' +
      '<p>Portal not yet configured. Please run <strong>Initialize Dashboard</strong> first.</p>' +
      '</body></html>';
  }
}

// ============================================================================
// BACKUP & ARCHIVE
// ============================================================================

/**
 * Creates a CSV backup of the Workload Vault to Google Drive.
 * Menu item: 📊 Workload Tracker > 💾 Create Backup
 */
function createWorkloadBackup() {
  var ui = SpreadsheetApp.getUi();
  try {
    var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) { ui.alert('Workload Vault sheet not found.'); return; }

    var data = vault.getDataRange().getValues();
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
    var fileName = 'WorkloadVault_Backup_' + timestamp + '.csv';

    var csv = data.map(function(row) {
      return row.map(function(cell) {
        var str = String(cell);
        return (str.indexOf(',') >= 0 || str.indexOf('"') >= 0 || str.indexOf('\n') >= 0)
          ? '"' + str.replace(/"/g, '""') + '"' : str;
      }).join(',');
    }).join('\n');

    var props = PropertiesService.getScriptProperties();
    var backupFolderId = props.getProperty('WT_BACKUP_FOLDER_ID');
    var folder;
    if (backupFolderId) {
      try { folder = DriveApp.getFolderById(backupFolderId); }
      catch (e) {
        folder = DriveApp.createFolder('WorkloadVault_Backups');
        props.setProperty('WT_BACKUP_FOLDER_ID', folder.getId());
      }
    } else {
      folder = DriveApp.createFolder('WorkloadVault_Backups');
      props.setProperty('WT_BACKUP_FOLDER_ID', folder.getId());
    }

    folder.createFile(fileName, csv, MimeType.CSV);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WORKLOAD_BACKUP', 'Backup created: ' + fileName);
    }
    ui.alert('Backup Created', 'Saved to Google Drive:\n' + fileName, ui.ButtonSet.OK);
  } catch (err) {
    console.error('createWorkloadBackup error:', err);
    ui.alert('Backup Failed', 'Error: ' + err.message, ui.ButtonSet.OK);
  }
}

/**
 * Archives Workload Vault data older than the retention period.
 * Menu item: 📊 Workload Tracker > 🗄️ Archive Old Data
 */
function wtArchiveOldData_() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Archive Old Workload Data',
    'Move workload data older than ' + WT_APP_CONFIG.dataRetentionMonths +
    ' months to the Archive sheet?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault) { ui.alert('Workload Vault not found.'); return; }

  var data = vault.getDataRange().getValues();
  if (data.length <= 1) { ui.alert('No data to archive.'); return; }

  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - WT_APP_CONFIG.dataRetentionMonths);

  var header = data[0];
  var current = [header];
  var archive = [];
  for (var i = 1; i < data.length; i++) {
    if (new Date(data[i][WT_VAULT_COLS.TIMESTAMP]) < cutoff) { archive.push(data[i]); }
    else { current.push(data[i]); }
  }
  if (archive.length === 0) { ui.alert('No data old enough to archive.'); return; }

  var archSheet = ss.getSheetByName(SHEETS.WORKLOAD_ARCHIVE);
  if (!archSheet) {
    archSheet = ss.insertSheet(SHEETS.WORKLOAD_ARCHIVE);
    archSheet.appendRow(header);
    archSheet.setFrozenRows(1);
    archSheet.hideSheet();
  }
  archSheet.getRange(archSheet.getLastRow() + 1, 1, archive.length, header.length).setValues(archive);

  vault.clear();
  vault.getRange(1, 1, current.length, header.length).setValues(current);
  refreshWorkloadReportingData();

  if (typeof logAuditEvent === 'function') {
    logAuditEvent('WORKLOAD_ARCHIVE', 'Archived ' + archive.length + ' workload records.');
  }
  ui.alert('Archived ' + archive.length + ' records.', '', ui.ButtonSet.OK);
}

// ============================================================================
// HEALTH STATUS
// ============================================================================

/**
 * Shows Workload Tracker system health in a dialog.
 * Menu item: 📊 Workload Tracker > 🩺 Health Status
 */
function showWorkloadHealthStatus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checks = {};

  var sheetNames = [
    SHEETS.WORKLOAD_VAULT, SHEETS.WORKLOAD_REPORTING,
    SHEETS.WORKLOAD_REMINDERS, SHEETS.WORKLOAD_USERMETA
  ];
  sheetNames.forEach(function(name) {
    var s = ss.getSheetByName(name);
    checks[name] = { exists: !!s, rows: s ? Math.max(0, s.getLastRow() - 1) : 0 };
  });

  var lines = ['Workload Tracker v' + WT_APP_CONFIG.version + '\n'];
  sheetNames.forEach(function(name) {
    var c = checks[name];
    lines.push((c.exists ? '✓ ' : '✗ ') + name + (c.exists ? ' (' + c.rows + ' records)' : ' — missing'));
  });
  lines.push('\nSetup: Admin > Setup > Initialize Dashboard to create missing sheets.');
  ui.alert('Workload Tracker Health', lines.join('\n'), ui.ButtonSet.OK);
}

// ============================================================================
// REMINDER SYSTEM SETUP
// ============================================================================

/**
 * Sets up the email reminder trigger for workload submissions.
 * Menu item: 📊 Workload Tracker > 🔔 Setup Reminders
 */
function setupWorkloadReminderSystem() {
  // Remove existing workload reminder triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'processWorkloadReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Daily trigger at 8 AM
  ScriptApp.newTrigger('processWorkloadReminders')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('Workload reminder system activated (daily at 8 AM).');
}

/**
 * Processes email reminders for members who opted in.
 * Called by daily trigger from setupWorkloadReminderSystem().
 */
function processWorkloadReminders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
  if (!sheet || sheet.getLastRow() <= 1) return;

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  var todayName = dayNames[now.getDay()];

  var portalUrl = PropertiesService.getScriptProperties().getProperty('WT_PORTAL_URL') || '';

  for (var i = 1; i < data.length; i++) {
    try {
      var email   = data[i][WT_REMINDERS_COLS.EMAIL];
      var enabled = data[i][WT_REMINDERS_COLS.ENABLED];
      var freq    = data[i][WT_REMINDERS_COLS.FREQUENCY] || 'weekly';
      var day     = (data[i][WT_REMINDERS_COLS.DAY] || 'monday').toLowerCase();

      if (!email || enabled !== 'Yes') continue;
      if (freq === 'weekly' && day !== todayName) continue;

      var subject = 'Reminder: Submit Your Workload Data';
      var body = 'Hello,\n\nThis is your scheduled reminder to submit your workload data for the week.\n\n' +
        (portalUrl ? 'Access the portal here:\n' + portalUrl + '?page=workload\n\n' : '') +
        'Thank you,\nSEIU 509 DDS Dashboard';

      MailApp.sendEmail({ to: email, subject: subject, body: body });
      sheet.getRange(i + 1, WT_REMINDERS_COLS.LAST_SENT + 1).setValue(now);
    } catch (err) {
      console.error('Reminder error for row ' + i + ': ' + err.message);
    }
  }
}

// ============================================================================
// PORTAL URL HELPERS (menu items)
// ============================================================================

/** Shows the Workload Portal URL. */
function showWorkloadPortalUrl() {
  var baseUrl = PropertiesService.getScriptProperties().getProperty('WT_PORTAL_URL') || '';
  var ui = SpreadsheetApp.getUi();
  if (baseUrl) {
    ui.alert('Workload Portal URL', baseUrl + '?page=workload\n\nShare this link with members to submit workload data.', ui.ButtonSet.OK);
  } else {
    ui.alert('Not configured', 'Deploy the web app, then use "Share Portal Link" to save its URL.', ui.ButtonSet.OK);
  }
}

/** Saves the web app URL for use in reminders and portal sharing. */
function shareWorkloadPortalLink() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Save Workload Portal URL',
    'Paste the deployed Web App URL (the ?page=workload suffix will be added automatically):',
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    var url = response.getResponseText().trim().replace(/[?&]page=workload.*$/, '');
    if (url) {
      PropertiesService.getScriptProperties().setProperty('WT_PORTAL_URL', url);
      ui.alert('Portal URL saved. Members can now access workload tracking at:\n' + url + '?page=workload');
    }
  }
}

// ============================================================================
// SSO WRAPPERS (v4.11.0) — called from web dashboard SPA
// Skip PIN authentication; rely on existing SSO session identity.
// ============================================================================

/**
 * Processes a workload form submission using SSO session identity (no PIN).
 * Called from the embedded workload tracker in the member dashboard.
 * @param {string} email — verified by SSO session
 * @param {Object} formData — { t1..t8, employment_type, part_time_hours, ... }
 * @returns {string} Success/error message
 */
function processWorkloadFormSSO(email, formData) {
  if (!email || !formData) return 'Error: Missing data.';
  var emailLower = email.toLowerCase().trim();

  // Submission rate limit
  var submitLimit = wtCheckSubmissionRateLimit_(emailLower);
  if (!submitLimit.allowed) {
    return 'Error: Submission limit reached. Please wait before submitting again.';
  }

  // Validate workload fields — accept range strings ("0", "1-5", "6-10", "11-20", "21+") or numeric
  var validRanges = ['0', '1-5', '6-10', '11-20', '21+'];
  var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
  for (var n = 0; n < numFields.length; n++) {
    var val = String(formData[numFields[n]] || '0').trim();
    // Accept range strings OR legacy numeric values
    if (validRanges.indexOf(val) === -1) {
      var numVal = Number(val);
      if (isNaN(numVal) || numVal < 0 || numVal > 999) {
        return 'Error: Invalid workload value for ' + numFields[n] + '.';
      }
    }
  }

  return withLock(function() {
    var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) {
      return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
    }

    // Accept both range strings and legacy numeric values
    var sanitizeVal = function(v) {
      var s = String(v || '0').trim();
      if (['0', '1-5', '6-10', '11-20', '21+'].indexOf(s) !== -1) return s;
      return Math.min(999, Math.max(0, Math.floor(Number(v) || 0)));
    };

    var row = [
      new Date(),                               // Timestamp
      emailLower,                               // Email
      sanitizeVal(formData.t1),                 // Priority Cases
      sanitizeVal(formData.t2),                 // Pending Cases
      sanitizeVal(formData.t3),                 // Unread Docs
      sanitizeVal(formData.t4),                 // To-Do Items
      sanitizeVal(formData.t5),                 // Sent Referrals
      sanitizeVal(formData.t6),                 // CE Activities
      sanitizeVal(formData.t7),                 // Assistance Requests
      sanitizeVal(formData.t8),                 // Aged Cases
      sanitizeNum(formData.weekly_cases),        // Weekly Cases
      formData.sub_categories || '',             // Sub-categories JSON
      formData.employment_type || 'Full-time',   // Employment type
      formData.part_time_hours || '',            // PT hours
      formData.leave_type || '',                 // Leave type
      formData.leave_planned || '',              // Leave planned
      formData.leave_start || '',                // Leave start
      formData.leave_end || '',                  // Leave end
      formData.no_intake_choice || '',           // No-intake choice
      formData.notice_time || '',                // Notice time
      formData.half_day || '',                   // Half day
      formData.privacy || 'Shared',              // Privacy
      formData.on_plan || '',                    // On plan
      formData.overtime_hours || ''              // Overtime hours
    ];

    vault.appendRow(row);
    wtRecordSubmission_(emailLower);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WT_SSO_SUBMIT', { email: emailLower });
    }

    return 'Workload data submitted successfully!';
  }, 'workload_sso_submit');
}

/**
 * Returns the authenticated member's submission history using SSO (no PIN).
 * @param {string} email — verified by SSO session
 * @returns {Object} { success, history:Array, message }
 */
function getWorkloadHistorySSO(email) {
  if (!email) return { success: false, history: [], message: 'Email required.' };

  var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault || vault.getLastRow() <= 1) {
    return { success: true, history: [], message: '' };
  }

  var emailLower = email.toLowerCase().trim();
  var data = vault.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][WT_VAULT_COLS.EMAIL];
    if (!rowEmail || rowEmail.toString().toLowerCase().trim() !== emailLower) continue;
    history.push({
      date:       data[i][WT_VAULT_COLS.TIMESTAMP],
      t1:         data[i][WT_VAULT_COLS.PRIORITY_CASES],
      t2:         data[i][WT_VAULT_COLS.PENDING_CASES],
      t3:         data[i][WT_VAULT_COLS.UNREAD_DOCS],
      t4:         data[i][WT_VAULT_COLS.TODO_ITEMS],
      t5:         data[i][WT_VAULT_COLS.SENT_REFERRALS],
      t6:         data[i][WT_VAULT_COLS.CE_ACTIVITIES],
      t7:         data[i][WT_VAULT_COLS.ASSISTANCE_REQUESTS],
      t8:         data[i][WT_VAULT_COLS.AGED_CASES],
      weeklyCases: data[i][WT_VAULT_COLS.WEEKLY_CASES],
      employment: data[i][WT_VAULT_COLS.EMPLOYMENT_TYPE],
      privacy:    data[i][WT_VAULT_COLS.PRIVACY],
    });
  }

  history.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  return { success: true, history: history, message: '' };
}

/**
 * Returns collective workload stats using SSO (no PIN).
 * Respects privacy settings and sharing start dates.
 * @param {string} email — verified by SSO session
 * @returns {Object} { success, data, message }
 */
function getWorkloadDashboardDataSSO(email) {
  if (!email) return { success: false, data: null, message: 'Email required.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  if (!vault || vault.getLastRow() <= 1) {
    return { success: true, data: { totalSubmissions: 0, members: 0, categories: {}, averages: {} }, message: '' };
  }

  var data = vault.getDataRange().getValues();
  var userSharingStart = typeof wtGetUserSharingStartDate_ === 'function'
    ? wtGetUserSharingStartDate_(email.toLowerCase().trim())
    : null;

  var totals = { t1: 0, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0 };
  var memberSet = {};
  var submissionCount = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_VAULT_COLS.PRIVACY] === 'Private') continue;
    var ts = new Date(data[i][WT_VAULT_COLS.TIMESTAMP]);
    if (userSharingStart && ts < userSharingStart) continue;

    submissionCount++;
    var rowEmail = data[i][WT_VAULT_COLS.EMAIL];
    if (rowEmail) memberSet[rowEmail.toString().toLowerCase()] = true;

    totals.t1 += Number(data[i][WT_VAULT_COLS.PRIORITY_CASES]) || 0;
    totals.t2 += Number(data[i][WT_VAULT_COLS.PENDING_CASES])  || 0;
    totals.t3 += Number(data[i][WT_VAULT_COLS.UNREAD_DOCS])    || 0;
    totals.t4 += Number(data[i][WT_VAULT_COLS.TODO_ITEMS])     || 0;
    totals.t5 += Number(data[i][WT_VAULT_COLS.SENT_REFERRALS]) || 0;
    totals.t6 += Number(data[i][WT_VAULT_COLS.CE_ACTIVITIES])  || 0;
    totals.t7 += Number(data[i][WT_VAULT_COLS.ASSISTANCE_REQUESTS]) || 0;
    totals.t8 += Number(data[i][WT_VAULT_COLS.AGED_CASES])     || 0;
  }

  var memberCount = Object.keys(memberSet).length;

  return {
    success: true,
    message: '',
    data: {
      totalSubmissions: submissionCount,
      members: memberCount,
      categories: {
        'Priority Cases':      totals.t1,
        'Pending Cases':       totals.t2,
        'Unread Documents':    totals.t3,
        'To-Do Items':         totals.t4,
        'Sent Referrals':      totals.t5,
        'CE Activities':       totals.t6,
        'Assistance Requests': totals.t7,
        'Aged Cases':          totals.t8
      },
      averages: submissionCount > 0 ? {
        'Priority Cases':      (totals.t1 / submissionCount).toFixed(1),
        'Pending Cases':       (totals.t2 / submissionCount).toFixed(1),
        'Unread Documents':    (totals.t3 / submissionCount).toFixed(1),
        'To-Do Items':         (totals.t4 / submissionCount).toFixed(1),
        'Sent Referrals':      (totals.t5 / submissionCount).toFixed(1),
        'CE Activities':       (totals.t6 / submissionCount).toFixed(1),
        'Assistance Requests': (totals.t7 / submissionCount).toFixed(1),
        'Aged Cases':          (totals.t8 / submissionCount).toFixed(1)
      } : {}
    }
  };
}


// ============================================================================
// SSO: REMINDER PREFERENCES (Web Dashboard)
// ============================================================================

/**
 * Gets the current user's reminder preferences for the web dashboard.
 * @param {string} email
 * @returns {Object} { enabled: boolean, day: string }
 */
function getWorkloadReminderSSO(email) {
  if (!email) return { enabled: false, day: 'monday' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
  if (!sheet || sheet.getLastRow() <= 1) return { enabled: false, day: 'monday' };

  var emailLower = email.toLowerCase().trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_REMINDERS_COLS.EMAIL] &&
        data[i][WT_REMINDERS_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
      return {
        enabled: data[i][WT_REMINDERS_COLS.ENABLED] === 'Yes',
        day: data[i][WT_REMINDERS_COLS.DAY] || 'monday'
      };
    }
  }
  return { enabled: false, day: 'monday' };
}

/**
 * Saves reminder preferences from the web dashboard.
 * @param {string} email
 * @param {boolean} enabled
 * @param {string} day — day of week (monday, tuesday, etc.)
 * @returns {Object} { success: boolean }
 */
function setWorkloadReminderSSO(email, enabled, day) {
  if (!email) return { success: false };
  wtSaveReminderPreference_({
    email: email,
    enabled: enabled,
    frequency: 'weekly',
    day: day || 'monday',
    time: '08:00'
  });
  return { success: true };
}
