/**
 * ============================================================================
 * 25_WorkloadService.gs - Workload Tracker Backend (SPA-only, SSO auth)
 * ============================================================================
 *
 * Full-featured workload tracking for the web dashboard SPA.
 * Uses SSO session identity (no PIN) via 19_WebDashAuth.gs.
 *
 * Replaces 18_WorkloadTracker.gs with correct column alignment,
 * full sub-category support, privacy reciprocity, overtime,
 * leave tracking, and enhanced analytics.
 *
 * Sheet Structure:
 *   "Workload Vault"       (hidden)  -- raw 24-column submissions
 *   "Workload Reporting"   (visible) -- anonymized ledger
 *   "Workload Reminders"   (hidden)  -- email reminder prefs
 *   "Workload UserMeta"    (hidden)  -- sharing start dates (reciprocity)
 *   "Workload Archive"     (hidden)  -- data older than 24 months
 *
 * @version 4.13.0
 * @requires 01_Core.gs   (SHEETS constants)
 * @requires 06_Maintenance.gs (logAuditEvent)
 */

var WorkloadService = (function() {

  // ── Configuration ─────────────────────────────────────────────────────────

  var CONFIG = {
    version: '4.13.0',
    maxStringLength: 255,
    dataRetentionMonths: 24,
    lockTimeoutMs: 30000,
    maxSubmissionsPerHour: 10,
    maxHistoryRequestsPerHour: 20
  };

  // ── Column Constants (0-indexed, matches vault appendRow layout) ──────────

  var VAULT_COLS = {
    TIMESTAMP:           0,
    EMAIL:               1,
    PRIORITY_CASES:      2,
    PENDING_CASES:       3,
    UNREAD_DOCS:         4,
    TODO_ITEMS:          5,
    SENT_REFERRALS:      6,
    CE_ACTIVITIES:       7,
    ASSISTANCE_REQUESTS: 8,
    AGED_CASES:          9,
    WEEKLY_CASES:        10,
    SUB_CATEGORIES:      11,
    EMPLOYMENT_TYPE:     12,
    PT_HOURS:            13,
    LEAVE_TYPE:          14,
    LEAVE_PLANNED:       15,
    LEAVE_START:         16,
    LEAVE_END:           17,
    NO_INTAKE_CHOICE:    18,
    NOTICE_TIME:         19,
    HALF_DAY:            20,
    PRIVACY:             21,
    ON_PLAN:             22,
    OVERTIME_HOURS:      23
  };

  var VAULT_COL_COUNT = 24;

  var USERMETA_COLS = { EMAIL: 0, SHARING_START_DATE: 1, CREATED_DATE: 2 };

  var REMINDERS_COLS = { EMAIL: 0, ENABLED: 1, FREQUENCY: 2, DAY: 3, TIME: 4, LAST_SENT: 5 };

  // ── Sub-Category Definitions ──────────────────────────────────────────────

  var SUB_CATEGORIES = {
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

  var CATEGORY_LABELS = {
    t1: 'Priority Cases', t2: 'Pending Cases', t3: 'Unread Documents', t4: 'To-Do Items',
    t5: 'Sent Referrals', t6: 'CE Activities', t7: 'Assistance Requests', t8: 'Aged Cases'
  };

  var CATEGORY_KEYS = ['priority', 'pending', 'unread', 'todo', 'referrals', 'ce', 'assistance', 'aged'];

  // ── Utility Functions (private) ───────────────────────────────────────────

  function _sanitizeString(input, maxLength) {
    maxLength = maxLength || CONFIG.maxStringLength;
    if (!input || typeof input !== 'string') return '';
    var sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
    return sanitized.length > maxLength ? sanitized.substring(0, maxLength) : sanitized;
  }

  function _sanitizeNum(val) {
    return Math.min(999, Math.max(0, Math.floor(Number(val) || 0)));
  }

  function _withLock(fn, lockName) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(CONFIG.lockTimeoutMs)) {
        throw new Error('System busy. Please try again in a moment.');
      }
      return fn();
    } finally {
      lock.releaseLock();
    }
  }

  function _getTimezone() {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  }

  // ── Rate Limiting (private) ───────────────────────────────────────────────

  function _checkRateLimit(key, maxAttempts, windowMinutes) {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'WT_RATE_' + key;
    var cached = cache.get(cacheKey);
    var attempts = cached ? (parseInt(cached, 10) || 0) : 0;
    if (attempts >= maxAttempts) {
      return { allowed: false, attemptsRemaining: 0, waitMinutes: windowMinutes };
    }
    return { allowed: true, attemptsRemaining: maxAttempts - attempts, waitMinutes: 0 };
  }

  function _recordRateLimitAttempt(key, windowMinutes) {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'WT_RATE_' + key;
    var cached = cache.get(cacheKey);
    var attempts = (cached ? (parseInt(cached, 10) || 0) : 0) + 1;
    cache.put(cacheKey, String(attempts), windowMinutes * 60);
  }

  function _checkSubmissionRateLimit(email) {
    return _checkRateLimit('SUBMIT_' + email.toLowerCase().trim(), CONFIG.maxSubmissionsPerHour, 60);
  }

  function _recordSubmission(email) {
    _recordRateLimitAttempt('SUBMIT_' + email.toLowerCase().trim(), 60);
  }

  // ── UserMeta / Sharing Start Date (private) ──────────────────────────────

  function _getUserSharingStartDate(email) {
    if (!email) return null;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_USERMETA);
    if (!sheet) return null;
    var emailLower = email.toLowerCase().trim();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][USERMETA_COLS.EMAIL] &&
          data[i][USERMETA_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
        return data[i][USERMETA_COLS.SHARING_START_DATE]
          ? new Date(data[i][USERMETA_COLS.SHARING_START_DATE]) : null;
      }
    }
    return null;
  }

  function _setUserSharingStartDate(email, startDate) {
    if (!email) return;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_USERMETA);
    if (!sheet) return;
    var emailLower = email.toLowerCase().trim();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][USERMETA_COLS.EMAIL] &&
          data[i][USERMETA_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
        sheet.getRange(i + 1, USERMETA_COLS.SHARING_START_DATE + 1).setValue(startDate);
        return;
      }
    }
    sheet.appendRow([emailLower, startDate, new Date()]);
  }

  // ── Reporting Refresh (private) ───────────────────────────────────────────

  function _refreshReportingData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
    if (!vault || !report) return;

    var vaultData = vault.getDataRange().getValues();
    if (vaultData.length <= 1) return;

    var tz = _getTimezone();

    // Determine latest privacy per email
    var latestPrivacy = {};
    for (var i = 1; i < vaultData.length; i++) {
      var rowEmail = vaultData[i][VAULT_COLS.EMAIL];
      if (rowEmail) {
        latestPrivacy[rowEmail.toString().toLowerCase().trim()] = vaultData[i][VAULT_COLS.PRIVACY];
      }
    }

    var header = [
      'Date', 'Identity',
      'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
      'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'Weekly Cases', 'Sub-Categories',
      'Employment Type', 'Leave Dates',
      'Anon Level', 'Workflow'
    ];
    var reportRows = [header];

    for (var j = 1; j < vaultData.length; j++) {
      var rEmail = vaultData[j][VAULT_COLS.EMAIL];
      if (!rEmail) continue;
      var _emailKey = rEmail.toString().toLowerCase().trim();
      var rowPrivacy = vaultData[j][VAULT_COLS.PRIVACY] || 'Unit';

      // Format sub-category summary
      var subCatSummary = '';
      var subCatJson = vaultData[j][VAULT_COLS.SUB_CATEGORIES];
      if (subCatJson && subCatJson !== '{}') {
        try {
          var subCats = typeof subCatJson === 'string' ? JSON.parse(subCatJson) : subCatJson;
          var parts = [];
          for (var ck in subCats) {
            if (subCats[ck] && typeof subCats[ck] === 'object') {
              var catTotal = 0;
              for (var sk in subCats[ck]) { catTotal += Number(subCats[ck][sk]) || 0; }
              if (catTotal > 0) parts.push(ck.charAt(0).toUpperCase() + ':' + catTotal);
            }
          }
          subCatSummary = parts.join(', ');
        } catch (_e) { /* skip */ }
      }

      // Format leave dates
      var leaveDates = '';
      var ls = vaultData[j][VAULT_COLS.LEAVE_START];
      var le = vaultData[j][VAULT_COLS.LEAVE_END];
      if (ls && le) {
        var lsStr = ls instanceof Date ? Utilities.formatDate(ls, tz, 'MM/dd/yyyy') : String(ls);
        var leStr = le instanceof Date ? Utilities.formatDate(le, tz, 'MM/dd/yyyy') : String(le);
        leaveDates = lsStr + ' - ' + leStr;
      } else if (ls) {
        leaveDates = ls instanceof Date ? Utilities.formatDate(ls, tz, 'MM/dd/yyyy') : String(ls);
      }

      var ts = vaultData[j][VAULT_COLS.TIMESTAMP];
      reportRows.push([
        ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yy') : String(ts),
        'REDACTED (' + rowPrivacy + ')',
        vaultData[j][VAULT_COLS.PRIORITY_CASES],
        vaultData[j][VAULT_COLS.PENDING_CASES],
        vaultData[j][VAULT_COLS.UNREAD_DOCS],
        vaultData[j][VAULT_COLS.TODO_ITEMS],
        vaultData[j][VAULT_COLS.SENT_REFERRALS],
        vaultData[j][VAULT_COLS.CE_ACTIVITIES],
        vaultData[j][VAULT_COLS.ASSISTANCE_REQUESTS],
        vaultData[j][VAULT_COLS.AGED_CASES],
        vaultData[j][VAULT_COLS.WEEKLY_CASES] || 15,
        subCatSummary,
        vaultData[j][VAULT_COLS.EMPLOYMENT_TYPE] || 'Full-time',
        leaveDates,
        rowPrivacy.toUpperCase(),
        vaultData[j][VAULT_COLS.ON_PLAN] === 'Yes' ? 'YES' : 'NO'
      ]);
    }

    // Write to reporting sheet
    report.clearContents();
    if (reportRows.length > 0) {
      report.getRange(1, 1, reportRows.length, header.length).setValues(reportRows);
      report.getRange(1, 1, 1, header.length)
        .setFontWeight('bold')
        .setBackground('#0d47a1')
        .setFontColor('#ffffff');
      report.setFrozenRows(1);

      // Pastel color coding
      if (reportRows.length > 1) {
        var numData = reportRows.length - 1;
        report.getRange(2, 1, numData, 2).setBackground('#F1F5F9');
        report.getRange(2, 3, numData, 8).setBackground('#ECFDF5');
        report.getRange(2, 11, numData, 2).setBackground('#FEF3C7');
        report.getRange(2, 13, numData, 2).setBackground('#FAF5FF');
        report.getRange(2, 15, numData, 2).setBackground('#FEE2E2');
      }
    }
  }

  // ── Reminder Helpers (private) ────────────────────────────────────────────

  function _saveReminderPreference(prefs) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!sheet) return;
    var emailLower = prefs.email.toLowerCase().trim();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][REMINDERS_COLS.EMAIL] &&
          data[i][REMINDERS_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          emailLower, prefs.enabled ? 'Yes' : 'No',
          prefs.frequency || 'weekly', prefs.day || 'monday', prefs.time || '08:00'
        ]]);
        return;
      }
    }
    sheet.appendRow([emailLower, prefs.enabled ? 'Yes' : 'No',
      prefs.frequency || 'weekly', prefs.day || 'monday', prefs.time || '08:00', '']);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // PUBLIC API
  // ══════════════════════════════════════════════════════════════════════════

  // ── Sheet Initialization ──────────────────────────────────────────────────

  function initSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Workload Vault (hidden raw data, 24 columns)
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) {
      vault = ss.insertSheet(SHEETS.WORKLOAD_VAULT);
      vault.appendRow([
        'Timestamp', 'Email',
        'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
        'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
        'Weekly Cases', 'Sub-Categories (JSON)',
        'Employment Type', 'PT Hours',
        'Leave Type', 'Leave Planned', 'Leave Start', 'Leave End',
        'No Intake Choice', 'Notice Time', 'Half Day',
        'Privacy', 'On Plan', 'Overtime Hours'
      ]);
      vault.getRange(1, 1, 1, VAULT_COL_COUNT)
        .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      vault.setFrozenRows(1);
    } else {
      // Existing sheet — expand headers if needed (handles migration from 12-col)
      var lastCol = vault.getLastColumn();
      if (lastCol < VAULT_COL_COUNT) {
        var expectedHeaders = [
          'Timestamp', 'Email',
          'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
          'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
          'Weekly Cases', 'Sub-Categories (JSON)',
          'Employment Type', 'PT Hours',
          'Leave Type', 'Leave Planned', 'Leave Start', 'Leave End',
          'No Intake Choice', 'Notice Time', 'Half Day',
          'Privacy', 'On Plan', 'Overtime Hours'
        ];
        for (var h = lastCol; h < VAULT_COL_COUNT; h++) {
          vault.getRange(1, h + 1).setValue(expectedHeaders[h]);
        }
      }
    }
    vault.hideSheet();

    // Workload Reporting (anonymized, visible)
    var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
    if (!report) {
      report = ss.insertSheet(SHEETS.WORKLOAD_REPORTING);
      report.appendRow([
        'Date', 'Identity',
        'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
        'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
        'Weekly Cases', 'Sub-Categories',
        'Employment Type', 'Leave Dates',
        'Anon Level', 'Workflow'
      ]);
      report.getRange(1, 1, 1, 16)
        .setFontWeight('bold').setBackground('#0d47a1').setFontColor('#ffffff');
      report.setFrozenRows(1);
    }

    // Workload Reminders (hidden)
    var reminders = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!reminders) {
      reminders = ss.insertSheet(SHEETS.WORKLOAD_REMINDERS);
      reminders.appendRow(['Email', 'Enabled', 'Frequency', 'Day', 'Time', 'Last Sent']);
      reminders.getRange(1, 1, 1, 6)
        .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      reminders.setFrozenRows(1);
      reminders.hideSheet();
    }

    // Workload UserMeta (hidden)
    var userMeta = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
    if (!userMeta) {
      userMeta = ss.insertSheet(SHEETS.WORKLOAD_USERMETA);
      userMeta.appendRow(['Email', 'Sharing Start Date', 'Created Date']);
      userMeta.getRange(1, 1, 1, 3)
        .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      userMeta.setFrozenRows(1);
      userMeta.hideSheet();
    }

    Logger.log('WorkloadService: sheets initialized (v' + CONFIG.version + ').');
  }

  // ── Form Submission (SSO) ─────────────────────────────────────────────────

  function processFormSSO(email, formData) {
    if (!email || !formData) return 'Error: Missing data.';
    var emailLower = email.toLowerCase().trim();

    // Rate limit
    var submitLimit = _checkSubmissionRateLimit(emailLower);
    if (!submitLimit.allowed) {
      return 'Error: Submission limit reached. Please wait before submitting again.';
    }

    // Validate numeric fields (t1-t8, 0-999)
    var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
    for (var n = 0; n < numFields.length; n++) {
      var val = Number(formData[numFields[n]]) || 0;
      if (val < 0 || val > 999 || val !== Math.floor(val)) {
        return 'Error: Workload values must be whole numbers between 0 and 999.';
      }
    }

    // Validate part-time hours if applicable
    if (formData.employment_type === 'Part-time') {
      var ptHours = Number(formData.part_time_hours);
      if (isNaN(ptHours) || ptHours < 1 || ptHours > 40 || ptHours !== Math.floor(ptHours)) {
        return 'Error: Part-time hours must be a whole number between 1 and 40.';
      }
    }

    return _withLock(function() {
      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault) {
        return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
      }

      // Build sub-category JSON
      var subCatData = {};
      for (var ci = 0; ci < CATEGORY_KEYS.length; ci++) {
        var catKey = CATEGORY_KEYS[ci];
        subCatData[catKey] = {};
        var subNames = SUB_CATEGORIES[catKey];
        for (var si = 0; si < subNames.length; si++) {
          var fieldName = 'sub_' + (ci + 1) + '_' + si;
          var subVal = _sanitizeNum(formData[fieldName]);
          if (subVal > 0) subCatData[catKey][subNames[si]] = subVal;
        }
      }

      // Weekly cases
      var weeklyCases = _sanitizeNum(formData.weekly_cases || 15);
      if (formData.weekly_cases_option === 'manual' && formData.weekly_cases_manual) {
        weeklyCases = _sanitizeNum(formData.weekly_cases_manual);
      }

      // Leave dates
      var tz = _getTimezone();
      var leaveStart = '', leaveEnd = '';
      if (formData.leave_start) {
        try { leaveStart = Utilities.formatDate(new Date(formData.leave_start), tz, 'MM/dd/yyyy'); } catch(_e) {}
      }
      if (formData.leave_end) {
        try { leaveEnd = Utilities.formatDate(new Date(formData.leave_end), tz, 'MM/dd/yyyy'); } catch(_e) {}
      }

      // Build 24-column row
      var row = [
        new Date(),                                                        // 0  Timestamp
        emailLower,                                                        // 1  Email
        _sanitizeNum(formData.t1),                                         // 2  Priority Cases
        _sanitizeNum(formData.t2),                                         // 3  Pending Cases
        _sanitizeNum(formData.t3),                                         // 4  Unread Docs
        _sanitizeNum(formData.t4),                                         // 5  To-Do Items
        _sanitizeNum(formData.t5),                                         // 6  Sent Referrals
        _sanitizeNum(formData.t6),                                         // 7  CE Activities
        _sanitizeNum(formData.t7),                                         // 8  Assistance Requests
        _sanitizeNum(formData.t8),                                         // 9  Aged Cases
        weeklyCases,                                                       // 10 Weekly Cases
        JSON.stringify(subCatData),                                        // 11 Sub-Categories JSON
        _sanitizeString(formData.employment_type || 'Full-time', 20),     // 12 Employment Type
        formData.employment_type === 'Part-time' ? (Number(formData.part_time_hours) || '') : '', // 13 PT Hours
        _sanitizeString(formData.leave_type || '', 20),                   // 14 Leave Type
        _sanitizeString(formData.leave_planned || '', 20),                // 15 Leave Planned
        leaveStart,                                                        // 16 Leave Start
        leaveEnd,                                                          // 17 Leave End
        _sanitizeString(formData.no_intake_choice || '', 20),             // 18 No Intake Choice
        _sanitizeString(formData.notice_time || '', 20),                  // 19 Notice Time
        formData.half_day ? 'Yes' : '',                                    // 20 Half Day
        _sanitizeString(formData.privacy || 'Unit', 20),                  // 21 Privacy
        formData.on_plan ? 'Yes' : 'No',                                   // 22 On Plan
        formData.overtime_enabled ? (Number(formData.overtime_hours) || 0) : '' // 23 Overtime
      ];

      vault.appendRow(row);
      _recordSubmission(emailLower);

      // Audit
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('WT_SSO_SUBMIT', 'Privacy: ' + (formData.privacy || 'Unit'));
      }

      // Reciprocity: set sharing start date
      var _privacySetting = _sanitizeString(formData.privacy || 'Unit', 20);
      var existingStartDate = _getUserSharingStartDate(emailLower);
      if (!existingStartDate) {
        _setUserSharingStartDate(emailLower, new Date());
      }

      // Save reminder preferences if provided
      if (formData.reminder_enabled !== undefined) {
        _saveReminderPreference({
          email: emailLower,
          enabled: formData.reminder_enabled === 'on' || formData.reminder_enabled === true || formData.reminder_enabled === 'true',
          frequency: formData.reminder_frequency || 'weekly',
          day: formData.reminder_day || 'monday',
          time: formData.reminder_time || '08:00'
        });
      }

      // Refresh anonymized ledger
      _refreshReportingData();

      return 'Success! Your workload data has been securely recorded.';
    }, 'processFormSSO');
  }

  // ── User History (SSO) ────────────────────────────────────────────────────

  function getHistorySSO(email) {
    if (!email) return { success: false, history: [], message: 'Email required.' };

    var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault || vault.getLastRow() <= 1) {
      return { success: true, history: [], message: '' };
    }

    var emailLower = email.toLowerCase().trim();
    var data = vault.getDataRange().getValues();
    var tz = _getTimezone();
    var history = [];

    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][VAULT_COLS.EMAIL];
      if (!rowEmail || rowEmail.toString().toLowerCase().trim() !== emailLower) continue;

      var ts = data[i][VAULT_COLS.TIMESTAMP];
      var entry = {
        date:        ts,
        dateDisplay: ts instanceof Date ? Utilities.formatDate(ts, tz, 'MM/dd/yyyy') : String(ts),
        t1:          Number(data[i][VAULT_COLS.PRIORITY_CASES]) || 0,
        t2:          Number(data[i][VAULT_COLS.PENDING_CASES]) || 0,
        t3:          Number(data[i][VAULT_COLS.UNREAD_DOCS]) || 0,
        t4:          Number(data[i][VAULT_COLS.TODO_ITEMS]) || 0,
        t5:          Number(data[i][VAULT_COLS.SENT_REFERRALS]) || 0,
        t6:          Number(data[i][VAULT_COLS.CE_ACTIVITIES]) || 0,
        t7:          Number(data[i][VAULT_COLS.ASSISTANCE_REQUESTS]) || 0,
        t8:          Number(data[i][VAULT_COLS.AGED_CASES]) || 0,
        weeklyCases: Number(data[i][VAULT_COLS.WEEKLY_CASES]) || 15,
        employment:  data[i][VAULT_COLS.EMPLOYMENT_TYPE] || 'Full-time',
        ptHours:     data[i][VAULT_COLS.PT_HOURS] || '',
        leaveType:   data[i][VAULT_COLS.LEAVE_TYPE] || '',
        leaveStart:  data[i][VAULT_COLS.LEAVE_START] || '',
        leaveEnd:    data[i][VAULT_COLS.LEAVE_END] || '',
        privacy:     data[i][VAULT_COLS.PRIVACY] || 'Unit',
        onPlan:      data[i][VAULT_COLS.ON_PLAN] === 'Yes' ? 'Yes' : 'No',
        overtime:    data[i][VAULT_COLS.OVERTIME_HOURS]
      };

      // Parse sub-categories
      var subJson = data[i][VAULT_COLS.SUB_CATEGORIES];
      if (subJson && subJson !== '{}') {
        try { entry.subCategories = typeof subJson === 'string' ? JSON.parse(subJson) : subJson; }
        catch (_e) { entry.subCategories = {}; }
      } else {
        entry.subCategories = {};
      }

      history.push(entry);
    }

    // Newest first
    history.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

    // Summary stats
    var totals = { t1: 0, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0 };
    var totalOT = 0, otCount = 0;
    for (var h = 0; h < history.length; h++) {
      for (var k = 1; k <= 8; k++) totals['t' + k] += history[h]['t' + k];
      var ot = history[h].overtime;
      if (ot !== '' && ot !== null && ot !== undefined && Number(ot) > 0) {
        totalOT += Number(ot); otCount++;
      }
    }
    var averages = {};
    for (var a = 1; a <= 8; a++) {
      averages['t' + a] = history.length > 0 ? Math.round((totals['t' + a] / history.length) * 10) / 10 : 0;
    }

    return {
      success: true,
      history: history,
      summary: {
        totalSubmissions: history.length,
        firstSubmission: history.length > 0 ? history[history.length - 1].dateDisplay : null,
        lastSubmission: history.length > 0 ? history[0].dateDisplay : null,
        averages: averages,
        overtime: {
          totalHours: totalOT,
          submissionsWithOvertime: otCount,
          averagePerSubmission: otCount > 0 ? Math.round((totalOT / otCount) * 10) / 10 : 0
        }
      },
      message: ''
    };
  }

  // ── Dashboard Data / Analytics (SSO) ──────────────────────────────────────

  function getDashboardDataSSO(email) {
    if (!email) return { success: false, data: null, message: 'Email required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault || vault.getLastRow() <= 1) {
      return { success: true, data: { totalSubmissions: 0, members: 0, categories: {}, averages: {} }, message: '' };
    }

    var emailLower = email.toLowerCase().trim();

    // Determine user's privacy setting (most recent submission)
    var data = vault.getDataRange().getValues();
    var _userPrivacy = null;
    for (var p = data.length - 1; p >= 1; p--) {
      var pEmail = data[p][VAULT_COLS.EMAIL];
      if (pEmail && pEmail.toString().toLowerCase().trim() === emailLower) {
        _userPrivacy = data[p][VAULT_COLS.PRIVACY];
        break;
      }
    }

    // Time-based reciprocity: only see data from sharing start date
    var userSharingStart = _getUserSharingStartDate(emailLower);
    if (!userSharingStart) {
      return {
        success: false,
        data: null,
        message: 'You must submit data with sharing enabled before viewing collective statistics.',
        reciprocityBlocked: true,
        neverShared: true
      };
    }

    var tz = _getTimezone();
    var totals = { t1: 0, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0 };
    var memberSet = {};
    var sharedCount = 0;
    var empBreakdown = { 'Full-time': 0, 'Part-time': 0 };
    var planBreakdown = { Yes: 0, No: 0 };
    var privacyBreakdown = { Unit: 0, Agency: 0 };
    var subCatTotals = {};
    var totalOT = 0, otSubmissions = 0;

    for (var i = 1; i < data.length; i++) {
      var privacy = data[i][VAULT_COLS.PRIVACY] || 'Unit';
      if (privacyBreakdown.hasOwnProperty(privacy)) privacyBreakdown[privacy]++;

      // Time-based reciprocity filter
      var ts = new Date(data[i][VAULT_COLS.TIMESTAMP]);
      if (userSharingStart && ts < userSharingStart) continue;

      sharedCount++;
      var rEmail = data[i][VAULT_COLS.EMAIL];
      if (rEmail) memberSet[rEmail.toString().toLowerCase()] = true;

      // Category totals
      totals.t1 += Number(data[i][VAULT_COLS.PRIORITY_CASES]) || 0;
      totals.t2 += Number(data[i][VAULT_COLS.PENDING_CASES]) || 0;
      totals.t3 += Number(data[i][VAULT_COLS.UNREAD_DOCS]) || 0;
      totals.t4 += Number(data[i][VAULT_COLS.TODO_ITEMS]) || 0;
      totals.t5 += Number(data[i][VAULT_COLS.SENT_REFERRALS]) || 0;
      totals.t6 += Number(data[i][VAULT_COLS.CE_ACTIVITIES]) || 0;
      totals.t7 += Number(data[i][VAULT_COLS.ASSISTANCE_REQUESTS]) || 0;
      totals.t8 += Number(data[i][VAULT_COLS.AGED_CASES]) || 0;

      // Employment breakdown
      var empType = data[i][VAULT_COLS.EMPLOYMENT_TYPE] || 'Full-time';
      if (empBreakdown.hasOwnProperty(empType)) empBreakdown[empType]++;

      // Workflow plan breakdown
      var onPlan = data[i][VAULT_COLS.ON_PLAN] === 'Yes' ? 'Yes' : 'No';
      planBreakdown[onPlan]++;

      // Overtime
      var otHours = Number(data[i][VAULT_COLS.OVERTIME_HOURS]) || 0;
      if (otHours > 0) { totalOT += otHours; otSubmissions++; }

      // Sub-category aggregation
      var scJson = data[i][VAULT_COLS.SUB_CATEGORIES];
      if (scJson && scJson !== '{}') {
        try {
          var sc = typeof scJson === 'string' ? JSON.parse(scJson) : scJson;
          for (var cKey in sc) {
            if (!subCatTotals[cKey]) subCatTotals[cKey] = {};
            for (var sKey in sc[cKey]) {
              var sVal = Number(sc[cKey][sKey]) || 0;
              if (!subCatTotals[cKey][sKey]) subCatTotals[cKey][sKey] = 0;
              subCatTotals[cKey][sKey] += sVal;
            }
          }
        } catch (_e) { /* skip */ }
      }
    }

    var memberCount = Object.keys(memberSet).length;

    // Averages
    var averages = {};
    if (sharedCount > 0) {
      for (var t = 1; t <= 8; t++) {
        var label = CATEGORY_LABELS['t' + t];
        averages[label] = (totals['t' + t] / sharedCount).toFixed(1);
      }
    }

    // Category totals with labels
    var catTotals = {};
    for (var c = 1; c <= 8; c++) catTotals[CATEGORY_LABELS['t' + c]] = totals['t' + c];

    return {
      success: true,
      message: '',
      userSharingStartDate: Utilities.formatDate(userSharingStart, tz, 'MM/dd/yyyy'),
      data: {
        totalSubmissions: sharedCount,
        members: memberCount,
        categories: catTotals,
        averages: averages,
        employment: empBreakdown,
        workflowPlan: planBreakdown,
        privacy: privacyBreakdown,
        subCategoryTotals: subCatTotals,
        overtime: {
          totalHours: totalOT,
          submissionsWithOvertime: otSubmissions,
          averagePerSubmission: otSubmissions > 0 ? Math.round((totalOT / otSubmissions) * 10) / 10 : 0,
          percentReporting: sharedCount > 0 ? Math.round((otSubmissions / sharedCount) * 100) : 0
        }
      }
    };
  }

  // ── Reminder Preferences (SSO) ────────────────────────────────────────────

  function getReminderSSO(email) {
    if (!email) return { enabled: false, frequency: 'weekly', day: 'monday', time: '08:00' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { enabled: false, frequency: 'weekly', day: 'monday', time: '08:00' };
    }
    var emailLower = email.toLowerCase().trim();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][REMINDERS_COLS.EMAIL] &&
          data[i][REMINDERS_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
        return {
          enabled: data[i][REMINDERS_COLS.ENABLED] === 'Yes',
          frequency: data[i][REMINDERS_COLS.FREQUENCY] || 'weekly',
          day: data[i][REMINDERS_COLS.DAY] || 'monday',
          time: data[i][REMINDERS_COLS.TIME] || '08:00'
        };
      }
    }
    return { enabled: false, frequency: 'weekly', day: 'monday', time: '08:00' };
  }

  function setReminderSSO(email, prefs) {
    if (!email) return { success: false };
    _saveReminderPreference({
      email: email,
      enabled: prefs.enabled,
      frequency: prefs.frequency || 'weekly',
      day: prefs.day || 'monday',
      time: prefs.time || '08:00'
    });
    return { success: true };
  }

  // ── CSV Export ─────────────────────────────────────────────────────────────

  function exportHistoryCSV(email) {
    var result = getHistorySSO(email);
    if (!result.success || !result.history.length) return '';

    var headers = ['Date', 'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
      'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'Weekly Cases', 'Employment Type', 'PT Hours', 'Leave Type', 'Leave Start', 'Leave End',
      'Privacy', 'On Plan', 'Overtime Hours'];

    var csvRows = [headers.join(',')];
    for (var i = 0; i < result.history.length; i++) {
      var e = result.history[i];
      csvRows.push([
        e.dateDisplay, e.t1, e.t2, e.t3, e.t4, e.t5, e.t6, e.t7, e.t8,
        e.weeklyCases, e.employment, e.ptHours, e.leaveType, e.leaveStart, e.leaveEnd,
        e.privacy, e.onPlan, e.overtime || ''
      ].map(function(cell) {
        var s = String(cell);
        return (s.indexOf(',') >= 0 || s.indexOf('"') >= 0) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(','));
    }
    return csvRows.join('\n');
  }

  // ── Sub-Categories Getter ─────────────────────────────────────────────────

  function getSubCategories() {
    return SUB_CATEGORIES;
  }

  // ── Process Reminders (daily trigger) ─────────────────────────────────────

  function processReminders() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!sheet || sheet.getLastRow() <= 1) return;

    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var currentHour = now.getHours();
    var dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    var currentDay = dayNames[now.getDay()];

    var portalUrl = PropertiesService.getScriptProperties().getProperty('WT_PORTAL_URL') || '';

    for (var i = 1; i < data.length; i++) {
      try {
        var rowEmail = data[i][REMINDERS_COLS.EMAIL];
        var enabled  = data[i][REMINDERS_COLS.ENABLED];
        var freq     = data[i][REMINDERS_COLS.FREQUENCY] || 'weekly';
        var day      = (data[i][REMINDERS_COLS.DAY] || 'monday').toLowerCase();
        var time     = data[i][REMINDERS_COLS.TIME] || '08:00';
        var lastSent = data[i][REMINDERS_COLS.LAST_SENT] ? new Date(data[i][REMINDERS_COLS.LAST_SENT]) : null;

        if (!rowEmail || enabled !== 'Yes') continue;

        // Check time
        var timeParts = time.toString().split(':');
        var reminderHour = parseInt(timeParts[0], 10);
        if (isNaN(reminderHour) || currentHour !== reminderHour) continue;

        // Check frequency
        var shouldSend = false;
        if (freq === 'daily') {
          shouldSend = true;
        } else if (freq === 'weekly' && currentDay === day) {
          shouldSend = true;
        } else if (freq === 'biweekly' && currentDay === day) {
          if (!lastSent || (now - lastSent) >= 14 * 24 * 60 * 60 * 1000) shouldSend = true;
        } else if (freq === 'monthly' && now.getDate() === 1) {
          shouldSend = true;
        } else if (freq === 'quarterly' && now.getDate() === 1) {
          if ([0, 3, 6, 9].indexOf(now.getMonth()) !== -1) shouldSend = true;
        }

        // Already sent today?
        if (shouldSend && lastSent && lastSent.toDateString() === today.toDateString()) {
          shouldSend = false;
        }

        if (shouldSend) {
          var subject = 'Reminder: Submit Your Workload Data';
          var body = 'Hello,\n\nThis is your scheduled reminder to submit your workload data.\n\n' +
            (portalUrl ? 'Access the dashboard here:\n' + portalUrl + '\n\n' : '') +
            'Thank you,\nSEIU 509 DDS Dashboard';
          MailApp.sendEmail({ to: rowEmail, subject: subject, body: body });
          sheet.getRange(i + 1, REMINDERS_COLS.LAST_SENT + 1).setValue(now);
        }
      } catch (err) {
        console.error('Reminder error for row ' + i + ': ' + err.message);
      }
    }
  }

  // ── Backup & Archive ──────────────────────────────────────────────────────

  function createBackup() {
    var ui = SpreadsheetApp.getUi();
    try {
      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault) { ui.alert('Workload Vault sheet not found.'); return; }

      var data = vault.getDataRange().getValues();
      var tz = _getTimezone();
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
      var folderId = props.getProperty('WT_BACKUP_FOLDER_ID');
      var folder;
      if (folderId) {
        try { folder = DriveApp.getFolderById(folderId); }
        catch (_e) {
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

  function archiveOldData() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Archive Old Workload Data',
      'Move workload data older than ' + CONFIG.dataRetentionMonths +
      ' months to the Archive sheet?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) { ui.alert('Workload Vault not found.'); return; }

    var data = vault.getDataRange().getValues();
    if (data.length <= 1) { ui.alert('No data to archive.'); return; }

    var cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - CONFIG.dataRetentionMonths);

    var header = data[0];
    var current = [header];
    var archive = [];
    for (var i = 1; i < data.length; i++) {
      if (new Date(data[i][VAULT_COLS.TIMESTAMP]) < cutoff) { archive.push(data[i]); }
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
    _refreshReportingData();

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WORKLOAD_ARCHIVE', 'Archived ' + archive.length + ' workload records.');
    }
    ui.alert('Archived ' + archive.length + ' records.', '', ui.ButtonSet.OK);
  }

  // ── Vault Cleaning ────────────────────────────────────────────────────────

  function cleanVault() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) return;

    var data = vault.getDataRange().getValues();
    if (data.length <= 1) return;

    var header = data[0];
    var dataRows = data.slice(1);
    var seen = {};
    var uniqueData = [];
    var tz = _getTimezone();

    // Sort newest first
    dataRows.sort(function(a, b) {
      return new Date(b[VAULT_COLS.TIMESTAMP]) - new Date(a[VAULT_COLS.TIMESTAMP]);
    });

    // Keep only first occurrence of each email+date combo
    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      var ts = row[VAULT_COLS.TIMESTAMP];
      var em = row[VAULT_COLS.EMAIL];
      if (!ts || !em) continue;
      try {
        var dateStr = Utilities.formatDate(new Date(ts), tz, 'yyyy-MM-dd');
        var id = em + '_' + dateStr;
        if (!seen[id]) { uniqueData.push(row); seen[id] = true; }
      } catch (_e) { uniqueData.push(row); }
    }

    uniqueData.reverse();
    var finalData = [header].concat(uniqueData);

    vault.clear();
    vault.getRange(1, 1, finalData.length, header.length).setValues(finalData);
    _refreshReportingData();
  }

  // ── Health Status ─────────────────────────────────────────────────────────

  function getHealthStatus() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = [
      SHEETS.WORKLOAD_VAULT, SHEETS.WORKLOAD_REPORTING,
      SHEETS.WORKLOAD_REMINDERS, SHEETS.WORKLOAD_USERMETA
    ];
    var lines = ['WorkloadService v' + CONFIG.version + '\n'];
    for (var i = 0; i < sheetNames.length; i++) {
      var name = sheetNames[i];
      var s = ss.getSheetByName(name);
      var exists = !!s;
      var rows = exists ? Math.max(0, s.getLastRow() - 1) : 0;
      lines.push((exists ? '\u2713 ' : '\u2717 ') + name + (exists ? ' (' + rows + ' records)' : ' \u2014 missing'));
    }
    lines.push('\nRun Setup > Initialize Dashboard to create missing sheets.');
    ui.alert('Workload Tracker Health', lines.join('\n'), ui.ButtonSet.OK);
  }

  // ── Ledger Refresh (menu) ─────────────────────────────────────────────────

  function refreshLedger() {
    _refreshReportingData();
    SpreadsheetApp.getUi().alert('Workload ledger refreshed successfully.');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // RETURN PUBLIC API
  // ══════════════════════════════════════════════════════════════════════════

  return {
    initSheets:          initSheets,
    processFormSSO:      processFormSSO,
    getHistorySSO:       getHistorySSO,
    getDashboardDataSSO: getDashboardDataSSO,
    getReminderSSO:      getReminderSSO,
    setReminderSSO:      setReminderSSO,
    exportHistoryCSV:    exportHistoryCSV,
    getSubCategories:    getSubCategories,
    processReminders:    processReminders,
    createBackup:        createBackup,
    archiveOldData:      archiveOldData,
    cleanVault:          cleanVault,
    getHealthStatus:     getHealthStatus,
    refreshLedger:       refreshLedger,
    SUB_CATEGORIES:      SUB_CATEGORIES,
    CATEGORY_LABELS:     CATEGORY_LABELS
  };

})();

// ============================================================================
// GLOBAL WRAPPERS (google.script.run calls from SPA)
// ============================================================================

function processWorkloadFormSSO(email, formData) { return WorkloadService.processFormSSO(email, formData); }
function getWorkloadHistorySSO(email) { return WorkloadService.getHistorySSO(email); }
function getWorkloadDashboardDataSSO(email) { return WorkloadService.getDashboardDataSSO(email); }
function getWorkloadReminderSSO(email) { return WorkloadService.getReminderSSO(email); }
function setWorkloadReminderSSO(email, prefs) { return WorkloadService.setReminderSSO(email, prefs); }
function exportWorkloadHistoryCSV(email) { return WorkloadService.exportHistoryCSV(email); }
function getWorkloadSubCategories() { return WorkloadService.getSubCategories(); }

// ============================================================================
// GLOBAL WRAPPERS (sheet init, triggers, menu items)
// ============================================================================

function initWorkloadTrackerSheets() { return WorkloadService.initSheets(); }
function processWorkloadReminders() { return WorkloadService.processReminders(); }
function refreshWorkloadLedger() { WorkloadService.refreshLedger(); }
function createWorkloadBackup() { WorkloadService.createBackup(); }
function wtArchiveOldData() { WorkloadService.archiveOldData(); }
function wtCleanVault() { WorkloadService.cleanVault(); }
function showWorkloadHealthStatus() { WorkloadService.getHealthStatus(); }

function setupWorkloadReminderSystem() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processWorkloadReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processWorkloadReminders')
    .timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('Workload reminder system activated (daily at 8 AM).');
}
