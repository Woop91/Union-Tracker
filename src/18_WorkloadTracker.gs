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
  version: '4.10.0',
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
