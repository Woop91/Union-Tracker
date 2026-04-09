/**
 * ============================================================================
 * 25_WorkloadService.gs - Workload Tracker Backend (SPA-only, SSO auth)
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Full-featured workload tracking backend for the web dashboard SPA. Tracks
 *   8 workload categories (priority cases, pending cases, unread docs, todo
 *   items, sent referrals, CE activities, assistance requests, aged cases)
 *   plus overtime hours, leave tracking, and sub-categories. Uses SSO identity
 *   (no PIN) via 19_WebDashAuth.gs. 5 hidden sheets:
 *     "Workload Vault"       — raw submissions
 *     "Workload Reporting"   — anonymized ledger
 *     "Workload Reminders"   — email reminder prefs
 *     "Workload UserMeta"    — sharing start dates (reciprocity)
 *     "Workload Archive"     — data older than 24 months
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Privacy reciprocity model — stewards only see each other's workload data
 *   from the date they started sharing their own. This prevents "lurking"
 *   (viewing others without contributing). 24-month data retention with
 *   automatic archiving keeps the vault sheet manageable. Rate limiting
 *   (10 submissions/hour, 20 history requests/hour) prevents abuse.
 *   0-indexed VAULT_COLS match the portal convention for getValues() arrays.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Stewards can't log or view workload data. The SPA workload tab shows
 *   nothing. If archiving breaks, the vault sheet grows unbounded (performance
 *   degrades). If rate limiting breaks, a script could spam submissions.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS), 19_WebDashAuth.gs (SSO identity),
 *   LockService (concurrency). Used by SPA workload views and email
 *   reminder triggers.
 *
 * @version 4.51.0
 */

var WorkloadService = (function() {

  // ── Configuration ─────────────────────────────────────────────────────────

  var CONFIG = {
    version: '4.43.1',
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

  /**
   * Strips control characters and truncates a string to the given max length.
   * @param {string} input - The raw string to sanitize.
   * @param {number} [maxLength] - Maximum allowed length (defaults to CONFIG.maxStringLength).
   * @returns {string} The sanitized, trimmed, and truncated string.
   */
  function _sanitizeString(input, maxLength) {
    maxLength = maxLength || CONFIG.maxStringLength;
    if (!input || typeof input !== 'string') return '';
    var sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
    return sanitized.length > maxLength ? sanitized.substring(0, maxLength) : sanitized;
  }

  /**
   * Clamps a value to an integer between 0 and 999.
   * @param {*} val - The value to sanitize.
   * @returns {number} An integer in the range [0, 999].
   */
  function _sanitizeNum(val) {
    return Math.min(999, Math.max(0, Math.floor(Number(val) || 0)));
  }

  /**
   * Executes a function while holding a script-level lock for concurrency safety.
   * @param {Function} fn - The callback to execute under lock.
   * @param {string} lockName - Descriptive name for logging (unused at runtime).
   * @returns {*} The return value of fn().
   */
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

  /**
   * Returns the spreadsheet's timezone, falling back to America/New_York.
   * @returns {string} IANA timezone identifier.
   */
  function _getTimezone() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss ? ss.getSpreadsheetTimeZone() : 'America/New_York';
  }

  // ── Rate Limiting (private) ───────────────────────────────────────────────

  /**
   * Atomic rate-limit check-and-increment. Reads, checks, and increments the
   * counter in one call so concurrent requests can't both pass the check before
   * either records.
   */
  function _checkAndRecordRateLimit(key, maxAttempts, windowMinutes) {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'WT_RATE_' + key;
    var cached = cache.get(cacheKey);
    var attempts = cached ? (parseInt(cached, 10) || 0) : 0;
    if (attempts >= maxAttempts) {
      return { allowed: false, attemptsRemaining: 0, waitMinutes: windowMinutes };
    }
    // Immediately increment so the next concurrent caller sees updated count
    cache.put(cacheKey, String(attempts + 1), windowMinutes * 60);
    return { allowed: true, attemptsRemaining: maxAttempts - attempts - 1, waitMinutes: 0 };
  }

  /**
   * Checks whether the given email has exceeded the submission rate limit.
   * @param {string} email - The submitter's email address.
   * @returns {{allowed: boolean, attemptsRemaining: number, waitMinutes: number}} Rate limit status.
   */
  function _checkSubmissionRateLimit(email) {
    return _checkAndRecordRateLimit('SUBMIT_' + email.toLowerCase().trim(), CONFIG.maxSubmissionsPerHour, 60);
  }

  // ── UserMeta / Sharing Start Date (private) ──────────────────────────────

  /**
   * Looks up a user's sharing-start date from the UserMeta sheet for reciprocity gating.
   * @param {string} email - The user's email address.
   * @returns {Date|null} The date the user began sharing, or null if not found.
   */
  function _getUserSharingStartDate(email) {
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
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

  /**
   * Sets or updates a user's sharing-start date in the UserMeta sheet.
   * @param {string} email - The user's email address.
   * @param {Date} startDate - The date to record as sharing start.
   * @returns {void}
   */
  function _setUserSharingStartDate(email, startDate) {
    if (!email) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
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

  /**
   * Rebuilds the anonymized Workload Reporting sheet from raw vault data.
   * @returns {void}
   */
  function _refreshReportingData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
    if (!vault || !report) return;

    var vaultData = vault.getDataRange().getValues();
    if (vaultData.length <= 1) return;

    var tz = _getTimezone();

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
        } catch (_e) { log_('_refreshReportingData', 'Error parsing sub-categories: ' + (_e.message || _e)); }
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
        (vaultData[j][VAULT_COLS.WEEKLY_CASES] != null ? vaultData[j][VAULT_COLS.WEEKLY_CASES] : 15),
        subCatSummary,
        vaultData[j][VAULT_COLS.EMPLOYMENT_TYPE] || 'Full-time',
        leaveDates,
        rowPrivacy.toUpperCase(),
        vaultData[j][VAULT_COLS.ON_PLAN] === 'Yes' ? 'YES' : 'NO'
      ]);
    }

    // Write to reporting sheet — write new data first, then clear stale rows to prevent data loss on timeout
    try {
      if (reportRows.length > 0) {
        // Write new data over existing content (overwrites, no clear needed for written range)
        report.getRange(1, 1, reportRows.length, header.length).setValues(reportRows);
        report.getRange(1, 1, 1, header.length)
          .setFontWeight('bold')
          .setBackground(SHEET_COLORS.HEADER_BLUE)
          .setFontColor(SHEET_COLORS.BG_WHITE);
        report.setFrozenRows(1);

        // Clear leftover rows beyond new data (if sheet previously had more rows)
        var lastRow = report.getLastRow();
        if (lastRow > reportRows.length) {
          report.getRange(reportRows.length + 1, 1, lastRow - reportRows.length, header.length).clearContent();
        }

        // Pastel color coding
        if (reportRows.length > 1) {
          var numData = reportRows.length - 1;
          report.getRange(2, 1, numData, 2).setBackground(SHEET_COLORS.BG_SLATE_LIGHT);
          report.getRange(2, 3, numData, 8).setBackground(SHEET_COLORS.BG_MINT);
          report.getRange(2, 11, numData, 2).setBackground(SHEET_COLORS.BG_LIGHT_YELLOW);
          report.getRange(2, 13, numData, 2).setBackground(SHEET_COLORS.BG_LIGHT_PURPLE);
          report.getRange(2, 15, numData, 2).setBackground(SHEET_COLORS.BG_LIGHT_RED);
        }
      }
    } catch (writeErr) {
      log_('_refreshReportingData', 'report write failed: ' + writeErr.message + '\n' + (writeErr.stack || ''));
    }
  }

  // ── Reminder Helpers (private) ────────────────────────────────────────────

  /**
   * Persists a user's email reminder preferences to the Reminders sheet.
   * @param {Object} prefs - Reminder settings.
   * @param {string} prefs.email - The user's email address.
   * @param {boolean} prefs.enabled - Whether reminders are enabled.
   * @param {string} [prefs.frequency] - Frequency (daily/weekly/biweekly/monthly/quarterly).
   * @param {string} [prefs.day] - Day of week for weekly/biweekly reminders.
   * @param {string} [prefs.time] - Time of day in HH:MM format.
   * @returns {void}
   */
  function _saveReminderPreference(prefs) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
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

  /**
   * Creates or migrates the four workload sheets (Vault, Reporting, Reminders, UserMeta).
   * @returns {void}
   */
  function initSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

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
        .setFontWeight('bold').setBackground(SHEET_COLORS.HEADER_NAVY).setFontColor(SHEET_COLORS.BG_WHITE);
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
    setSheetVeryHidden_(vault); // H-12: API-level hide — survives mobile Sheets

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
        .setFontWeight('bold').setBackground(SHEET_COLORS.HEADER_BLUE).setFontColor(SHEET_COLORS.BG_WHITE);
      report.setFrozenRows(1);
    }

    // Workload Reminders (hidden)
    var reminders = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!reminders) {
      reminders = ss.insertSheet(SHEETS.WORKLOAD_REMINDERS);
      reminders.appendRow(['Email', 'Enabled', 'Frequency', 'Day', 'Time', 'Last Sent']);
      reminders.getRange(1, 1, 1, 6)
        .setFontWeight('bold').setBackground(SHEET_COLORS.HEADER_NAVY).setFontColor(SHEET_COLORS.BG_WHITE);
      reminders.setFrozenRows(1);
      setSheetVeryHidden_(reminders); // H-12: API-level hide
    }

    // Workload UserMeta (hidden)
    var userMeta = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
    if (!userMeta) {
      userMeta = ss.insertSheet(SHEETS.WORKLOAD_USERMETA);
      userMeta.appendRow(['Email', 'Sharing Start Date', 'Created Date']);
      userMeta.getRange(1, 1, 1, 3)
        .setFontWeight('bold').setBackground(SHEET_COLORS.HEADER_NAVY).setFontColor(SHEET_COLORS.BG_WHITE);
      userMeta.setFrozenRows(1);
      setSheetVeryHidden_(userMeta); // H-12: API-level hide
    }

    log_('WorkloadService', 'sheets initialized (v' + CONFIG.version + ').');
  }

  // ── Form Submission (SSO) ─────────────────────────────────────────────────

  /**
   * Validates and records a workload submission from the SPA, with rate limiting and reciprocity.
   * @param {string} email - The authenticated user's email.
   * @param {Object} formData - Form fields (t1-t8, employment_type, leave_*, privacy, etc.).
   * @returns {string} Success or error message string.
   */
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
      var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
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
        try { leaveStart = Utilities.formatDate(new Date(formData.leave_start), tz, 'MM/dd/yyyy'); } catch (_e) { log_('processFormSSO', 'Error formatting leave start: ' + (_e.message || _e)); }
      }
      if (formData.leave_end) {
        try { leaveEnd = Utilities.formatDate(new Date(formData.leave_end), tz, 'MM/dd/yyyy'); } catch (_e) { log_('processFormSSO', 'Error formatting leave end: ' + (_e.message || _e)); }
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
      // Rate limit already recorded atomically in _checkSubmissionRateLimit

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

  /**
   * Returns the authenticated user's full workload submission history with summary stats.
   * @param {string} email - The authenticated user's email.
   * @returns {{success: boolean, history: Array<Object>, summary: Object, message: string}} History payload.
   */
  function getHistorySSO(email) {
    if (!email) return { success: false, history: [], message: 'Email required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, history: [], message: '' };
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
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

  /**
   * Returns aggregate workload analytics gated by the user's sharing-start date (reciprocity).
   * @param {string} email - The authenticated user's email.
   * @returns {{success: boolean, data: Object|null, message: string}} Dashboard analytics payload.
   */
  function getDashboardDataSSO(email) {
    if (!email) return { success: false, data: null, message: 'Email required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, data: null, message: 'Spreadsheet unavailable.' };
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
      if (Object.prototype.hasOwnProperty.call(privacyBreakdown, privacy)) privacyBreakdown[privacy]++;

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
      if (Object.prototype.hasOwnProperty.call(empBreakdown, empType)) empBreakdown[empType]++;

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
        } catch (_e) { log_('getDashboardDataSSO', 'Error parsing sub-categories: ' + (_e.message || _e)); }
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

  /**
   * Retrieves the user's email reminder preferences from the Reminders sheet.
   * @param {string} email - The user's email address.
   * @returns {{enabled: boolean, frequency: string, day: string, time: string}} Reminder settings.
   */
  function getReminderSSO(email) {
    if (!email) return { enabled: false, frequency: 'weekly', day: 'monday', time: '08:00' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { enabled: false, frequency: 'weekly', day: 'monday', time: '08:00' };
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
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

  /**
   * Saves updated email reminder preferences for the authenticated user.
   * @param {string} email - The user's email address.
   * @param {Object} prefs - Reminder settings (enabled, frequency, day, time).
   * @returns {{success: boolean}} Result indicator.
   */
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

  /**
   * Exports the user's workload history as a CSV-formatted string.
   * @param {string} email - The user's email address.
   * @returns {string} CSV content, or empty string if no data.
   */
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

  /**
   * Returns the sub-category definitions for all eight workload categories.
   * @returns {Object<string, string[]>} Map of category keys to arrays of sub-category names.
   */
  function getSubCategories() {
    return SUB_CATEGORIES;
  }

  // ── Process Reminders (daily trigger) ─────────────────────────────────────

  /**
   * Sends scheduled workload reminder emails based on each user's frequency/day/time preferences.
   * @returns {void}
   */
  function processReminders() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
    if (!sheet || sheet.getLastRow() <= 1) return;

    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var currentHour = now.getHours();
    var dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    var currentDay = dayNames[now.getDay()];

    var portalUrl = '';
    try { portalUrl = ScriptApp.getService().getUrl() || ''; } catch (_e) { log_('processReminders', 'Error getting portal URL: ' + (_e.message || _e)); }

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
            'Thank you,\n' + ((typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.SYSTEM_NAME) ? COMMAND_CONFIG.SYSTEM_NAME : 'Dashboard');
          MailApp.sendEmail({ to: rowEmail, subject: subject, body: body });
          sheet.getRange(i + 1, REMINDERS_COLS.LAST_SENT + 1).setValue(now);
        }
      } catch (err) {
        log_('processReminders', 'Reminder error for row ' + i + ': ' + err.message);
      }
    }
  }

  // ── Backup & Archive ──────────────────────────────────────────────────────

  /**
   * Exports the Workload Vault as a timestamped CSV backup to Google Drive.
   * @returns {void}
   */
  function createBackup() {
    var ui = SpreadsheetApp.getUi();
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) { ui.alert('Spreadsheet unavailable.'); return; }
      var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
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
        try {
          folder = DriveApp.getFolderById(folderId);
          try { folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (_e) {}
        }
        catch (_e) {
          folder = DriveApp.createFolder('WorkloadVault_Backups');
          folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
          folder.setDescription('Automated workload backups — contains member PII. Do not share.');
          props.setProperty('WT_BACKUP_FOLDER_ID', folder.getId());
        }
      } else {
        folder = DriveApp.createFolder('WorkloadVault_Backups');
        folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        folder.setDescription('Automated workload backups — contains member PII. Do not share.');
        props.setProperty('WT_BACKUP_FOLDER_ID', folder.getId());
      }

      folder.createFile(fileName, csv, MimeType.CSV);
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('WORKLOAD_BACKUP', 'Backup created: ' + fileName);
      }
      ui.alert('Backup Created', 'Saved to Google Drive:\n' + fileName, ui.ButtonSet.OK);
    } catch (err) {
      log_('createWorkloadBackup error', err);
      ui.alert('Backup Failed', 'Error: ' + err.message, ui.ButtonSet.OK);
    }
  }

  /**
   * Moves vault rows older than the retention period to the Archive sheet (with UI confirmation).
   * @returns {void}
   */
  function archiveOldData() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Archive Old Workload Data',
      'Move workload data older than ' + CONFIG.dataRetentionMonths +
      ' months to the Archive sheet?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) { ui.alert('Spreadsheet unavailable.'); return; }
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
      setSheetVeryHidden_(archSheet); // H-12: API-level hide
    }
    // M-37: Write archive data first, then rewrite vault — prevents data loss if write fails
    archSheet.getRange(archSheet.getLastRow() + 1, 1, archive.length, header.length).setValues(archive);

    // Only clear and rewrite vault after archive write succeeds
    vault.clearContents();
    vault.getRange(1, 1, current.length, header.length).setValues(current);
    _refreshReportingData();

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WORKLOAD_ARCHIVE', 'Archived ' + archive.length + ' workload records.');
    }
    ui.alert('Archived ' + archive.length + ' records.', '', ui.ButtonSet.OK);
  }

  // ── Vault Cleaning ────────────────────────────────────────────────────────

  /**
   * Deduplicates the vault by keeping the newest entry per email+date, using copy-on-write for safety.
   * @returns {void}
   */
  function cleanVault() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
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

    // M-38: Write new data before clearing to prevent data loss if write fails.
    // Use a lock to ensure atomicity of the read-clean-write cycle.
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(CONFIG.lockTimeoutMs);
    } catch (_lockErr) {
      log_('cleanVault', 'Could not acquire lock');
      return;
    }
    try {
      // Copy-on-write: validate finalData by writing to a staging range below current data
      var stagingStartRow = vault.getLastRow() + 2;
      vault.getRange(stagingStartRow, 1, finalData.length, header.length).setValues(finalData);
      // Staging write succeeded — now clear and swap
      try {
        vault.clearContents();
        vault.getRange(1, 1, finalData.length, header.length).setValues(finalData);
      } catch (swapErr) {
        log_('cleanVault', 'swap failed after staging — attempting restore: ' + swapErr.message);
        try {
          vault.clearContents();
          vault.getRange(1, 1, data.length, header.length).setValues(data);
        } catch (recoveryErr) {
          log_('cleanVault', 'recovery also failed: ' + recoveryErr.message);
          handleError(recoveryErr, 'cleanVault: data loss — swap and recovery both failed', ERROR_LEVEL.CRITICAL);
        }
      }
    } catch (stagingErr) {
      log_('cleanVault', 'staging write failed — original data untouched: ' + stagingErr.message);
    } finally {
      lock.releaseLock();
    }
    _refreshReportingData();
  }

  // ── Health Status ─────────────────────────────────────────────────────────

  /**
   * Displays a UI alert showing row counts and existence status for all workload sheets.
   * @returns {void}
   */
  function getHealthStatus() {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) { ui.alert('Spreadsheet unavailable.'); return; }
    var sheetNames = [
      SHEETS.WORKLOAD_VAULT, SHEETS.WORKLOAD_REPORTING,
      SHEETS.WORKLOAD_REMINDERS, SHEETS.WORKLOAD_USERMETA
    ];
    var lines = ['WorkloadService v' + CONFIG.version + '\n'];
    for (var i = 0; i < sheetNames.length; i++) {
      var name = sheetNames[i];
      var s = ss.getSheetByName(name);  // null if sheet doesn't exist — handled below
      var exists = !!s;
      var rows = exists ? Math.max(0, s.getLastRow() - 1) : 0;
      lines.push((exists ? '\u2713 ' : '\u2717 ') + name + (exists ? ' (' + rows + ' records)' : ' \u2014 missing'));
    }
    lines.push('\nRun Setup > Initialize Dashboard to create missing sheets.');
    ui.alert('Workload Tracker Health', lines.join('\n'), ui.ButtonSet.OK);
  }

  // ── Ledger Refresh (menu) ─────────────────────────────────────────────────

  /**
   * Manually triggers a reporting-sheet refresh and shows a confirmation alert.
   * @returns {void}
   */
  function refreshLedger() {
    _refreshReportingData();
    SpreadsheetApp.getUi().alert('Workload ledger refreshed successfully.');
  }

  // ── Predictive Workload Balancing ─────────────────────────────────────────

  /**
   * Predicts workload trend for a steward based on historical vault submissions.
   * Uses the last 4-8 entries to compute a simple linear trend.
   * @param {string} stewardEmail - Email of the steward to predict for.
   * @returns {{current: number, predicted: number, trend: string, history: number[]}}
   */
  function predictWorkloadTrend(stewardEmail) {
    var fallback = { current: 0, predicted: 0, trend: 'stable' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return fallback;
    var sheet = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!sheet) return fallback;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return fallback;

    // Use VAULT_COLS for column indices (0-indexed)
    var emailCol = VAULT_COLS.EMAIL;
    var dateCol = VAULT_COLS.TIMESTAMP;

    // Sum all 8 workload categories for total caseload per entry
    var loadCols = [
      VAULT_COLS.PRIORITY_CASES, VAULT_COLS.PENDING_CASES, VAULT_COLS.UNREAD_DOCS,
      VAULT_COLS.TODO_ITEMS, VAULT_COLS.SENT_REFERRALS, VAULT_COLS.CE_ACTIVITIES,
      VAULT_COLS.ASSISTANCE_REQUESTS, VAULT_COLS.AGED_CASES
    ];

    var target = String(stewardEmail).trim().toLowerCase();
    var entries = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol] || '').trim().toLowerCase() === target) {
        var load = 0;
        for (var lc = 0; lc < loadCols.length; lc++) {
          load += parseInt(data[i][loadCols[lc]], 10) || 0;
        }
        var date = data[i][dateCol] ? new Date(data[i][dateCol]) : null;
        entries.push({ load: load, date: date });
      }
    }

    if (entries.length === 0) return fallback;

    // Sort by date ascending
    entries.sort(function(a, b) { return (a.date || 0) - (b.date || 0); });

    var current = entries[entries.length - 1].load;
    var recent = entries.slice(-4);
    var avgRecent = recent.reduce(function(s, e) { return s + e.load; }, 0) / recent.length;

    // Simple linear prediction based on recent vs older averages
    var predicted = Math.round(avgRecent);
    var trend = 'stable';
    if (entries.length >= 3) {
      var older = entries.slice(-6, -3);
      if (older.length > 0) {
        var avgOlder = older.reduce(function(s, e) { return s + e.load; }, 0) / older.length;
        if (avgRecent > avgOlder * 1.15) {
          trend = 'increasing';
          predicted = Math.round(avgRecent * 1.1);
        } else if (avgRecent < avgOlder * 0.85) {
          trend = 'decreasing';
          predicted = Math.round(avgRecent * 0.9);
        }
      }
    }

    return {
      current: current,
      predicted: predicted,
      trend: trend,
      history: entries.slice(-8).map(function(e) { return e.load; })
    };
  }

  /**
   * Suggests rebalancing across all stewards based on active grievance case distribution.
   * Returns suggestions for overloaded (>150% avg) and under-capacity (<50% avg) stewards.
   * @returns {Array<{stewardEmail: string, currentLoad: number, averageLoad: number, suggestion: string}>}
   */
  function suggestRebalancing() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // Use GRIEVANCE_COLS for column indices (1-indexed, so subtract 1 for array access)
    var stewardCol = GRIEVANCE_COLS.STEWARD - 1;
    var statusCol = GRIEVANCE_COLS.STATUS - 1;

    // Count active cases per steward (exclude terminal statuses)
    var terminalStatuses = { resolved: true, closed: true, denied: true, settled: true, won: true, withdrawn: true };
    var loads = {};
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][statusCol] || '').toLowerCase().trim();
      if (terminalStatuses[status]) continue;
      var steward = String(data[i][stewardCol] || '').trim().toLowerCase();
      if (!steward) continue;
      loads[steward] = (loads[steward] || 0) + 1;
    }

    var entries = Object.keys(loads).map(function(email) { return { email: email, cases: loads[email] }; });
    if (entries.length < 2) return [];

    var avg = entries.reduce(function(s, e) { return s + e.cases; }, 0) / entries.length;
    var suggestions = [];

    entries.forEach(function(e) {
      if (e.cases > avg * 1.5) {
        suggestions.push({
          stewardEmail: e.email,
          currentLoad: e.cases,
          averageLoad: Math.round(avg),
          suggestion: 'Overloaded \u2014 consider reassigning ' + Math.round(e.cases - avg) + ' case(s)'
        });
      } else if (e.cases < avg * 0.5 && avg > 2) {
        suggestions.push({
          stewardEmail: e.email,
          currentLoad: e.cases,
          averageLoad: Math.round(avg),
          suggestion: 'Under capacity \u2014 can take ' + Math.round(avg - e.cases) + ' more case(s)'
        });
      }
    });

    return suggestions;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // RETURN PUBLIC API
  // ══════════════════════════════════════════════════════════════════════════

  return {
    initSheets:            initSheets,
    processFormSSO:        processFormSSO,
    getHistorySSO:         getHistorySSO,
    getDashboardDataSSO:   getDashboardDataSSO,
    getReminderSSO:        getReminderSSO,
    setReminderSSO:        setReminderSSO,
    exportHistoryCSV:      exportHistoryCSV,
    getSubCategories:      getSubCategories,
    processReminders:      processReminders,
    createBackup:          createBackup,
    archiveOldData:        archiveOldData,
    cleanVault:            cleanVault,
    getHealthStatus:       getHealthStatus,
    refreshLedger:         refreshLedger,
    predictWorkloadTrend:  predictWorkloadTrend,
    suggestRebalancing:    suggestRebalancing,
    SUB_CATEGORIES:        SUB_CATEGORIES,
    CATEGORY_LABELS:       CATEGORY_LABELS
  };

})();

// ============================================================================
// GLOBAL WRAPPERS (google.script.run calls from SPA)
// ============================================================================

/**
 * SPA wrapper: submits workload form data after SSO authentication.
 * @param {string} sessionToken - The caller's session token.
 * @param {Object} formData - Workload form fields.
 * @returns {string|Object} Success/error message or auth failure object.
 */
function processWorkloadFormSSO(sessionToken, formData) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.processFormSSO(e, formData) : { success: false, message: 'Not authenticated.' }; }
/**
 * SPA wrapper: retrieves the caller's workload submission history.
 * @param {string} sessionToken - The caller's session token.
 * @returns {{success: boolean, history: Array, message: string}} History payload or auth failure.
 */
function getWorkloadHistorySSO(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.getHistorySSO(e) : { success: false, history: [], message: 'Not authenticated.' }; }
/**
 * SPA wrapper: retrieves aggregate workload dashboard analytics.
 * @param {string} sessionToken - The caller's session token.
 * @returns {{success: boolean, data: Object|null, message: string}} Dashboard data or auth failure.
 */
function getWorkloadDashboardDataSSO(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.getDashboardDataSSO(e) : { success: false, message: 'Not authenticated.' }; }
/**
 * SPA wrapper: retrieves the caller's reminder preferences.
 * @param {string} sessionToken - The caller's session token.
 * @returns {Object} Reminder settings or auth failure object.
 */
function getWorkloadReminderSSO(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.getReminderSSO(e) : { success: false, reminder: null, message: 'Not authenticated.' }; }
/**
 * SPA wrapper: saves updated reminder preferences for the caller.
 * @param {string} sessionToken - The caller's session token.
 * @param {Object} prefs - Reminder settings (enabled, frequency, day, time).
 * @returns {{success: boolean}} Result indicator or auth failure.
 */
function setWorkloadReminderSSO(sessionToken, prefs) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.setReminderSSO(e, prefs) : { success: false, message: 'Not authenticated.' }; }
/**
 * SPA wrapper: exports the caller's workload history as CSV.
 * @param {string} sessionToken - The caller's session token.
 * @returns {string|Object} CSV string or auth failure object.
 */
function exportWorkloadHistoryCSV(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? WorkloadService.exportHistoryCSV(e) : { success: false, message: 'Not authenticated.' }; }
/**
 * SPA wrapper: returns sub-category definitions (no auth required).
 * @returns {Object<string, string[]>} Map of category keys to sub-category name arrays.
 */
function getWorkloadSubCategories() { return WorkloadService.getSubCategories(); }
/**
 * SPA wrapper: predicts workload trend for a steward (steward-only).
 * @param {string} sessionToken - The caller's session token.
 * @param {string} [stewardEmail] - Optional target email; defaults to caller.
 * @returns {{current: number, predicted: number, trend: string, history: number[]}}
 */
function dataGetWorkloadPrediction(sessionToken, stewardEmail) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { current: 0, predicted: 0, trend: 'stable' };
  return WorkloadService.predictWorkloadTrend(stewardEmail || s);
}
/**
 * SPA wrapper: suggests rebalancing across all stewards (steward-only).
 * @param {string} sessionToken - The caller's session token.
 * @returns {Array<{stewardEmail: string, currentLoad: number, averageLoad: number, suggestion: string}>}
 */
function dataGetRebalancingSuggestions(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return [];
  return WorkloadService.suggestRebalancing();
}

// ============================================================================
// GLOBAL WRAPPERS (sheet init, triggers, menu items)
// ============================================================================

/**
 * Global wrapper: creates or migrates all workload tracker sheets.
 * @returns {void}
 */
function initWorkloadTrackerSheets() { return WorkloadService.initSheets(); }
/**
 * Global wrapper: processes scheduled workload reminder emails (daily trigger target).
 * @returns {void}
 */
function processWorkloadReminders() { return WorkloadService.processReminders(); }
/**
 * Global wrapper: refreshes the anonymized workload reporting ledger (menu item).
 * @returns {void}
 */
function refreshWorkloadLedger() { WorkloadService.refreshLedger(); }
/**
 * Global wrapper: creates a CSV backup of the Workload Vault in Google Drive.
 * @returns {void}
 */
function createWorkloadBackup() { WorkloadService.createBackup(); }
/**
 * Global wrapper: archives vault data older than the retention period.
 * @returns {void}
 */
function wtArchiveOldData() { WorkloadService.archiveOldData(); }
/**
 * Global wrapper: deduplicates the Workload Vault sheet.
 * @returns {void}
 */
function wtCleanVault() { WorkloadService.cleanVault(); }
/**
 * Global wrapper: displays the workload tracker health status dialog.
 * @returns {void}
 */
function showWorkloadHealthStatus() { WorkloadService.getHealthStatus(); }

/**
 * Creates a daily time-based trigger for workload reminder emails (replaces any existing one).
 * @returns {void}
 */
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
