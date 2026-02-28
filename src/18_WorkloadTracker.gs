/**
 * ============================================================================
 * 18_WorkloadTracker.gs - Member Workload Tracking Module (IIFE)
 * ============================================================================
 *
 * Secure, anonymized workload tracking integrated with the DDS Dashboard.
 * Members submit weekly caseload data via the standalone web portal
 * (?page=workload). Authentication uses the DDS member PIN system.
 *
 * Sheet Structure:
 *   "Workload Vault"     (hidden)  — raw 24-column data with member email
 *   "Workload Reporting" (visible) — anonymized ledger (identity → REDACTED)
 *   "Workload Reminders"           — email reminder preferences
 *   "Workload UserMeta"  (hidden)  — sharing start dates for reciprocity
 *   "Workload Archive"   (hidden)  — data older than retention period
 *
 * Auth: email + DDS member PIN → verifyPIN(pin, memberId, storedHash)
 *
 * @version 4.13.0
 * @requires 01_Core.gs (SHEETS, MEMBER_COLS)
 * @requires 13_MemberSelfService.gs (verifyPIN, hashPIN, getPINSalt_)
 * @requires 06_Maintenance.gs (logAuditEvent)
 * @requires 00_Security.gs (escapeHtml)
 */

// ============================================================================
// WORKLOAD SERVICE — IIFE MODULE
// ============================================================================

var WorkloadPortal = (function () {

  // ── Private Constants ────────────────────────────────────────────────────

  var APP_CONFIG = {
    version: '4.13.0',
    name: 'Workload Tracker',
    maxStringLength: 255,
    dataRetentionMonths: 24,
    lockTimeoutMs: 30000
  };

  var SECURITY_CONFIG = {
    maxPinAttempts: 5,
    pinLockoutMinutes: 15,
    maxSubmissionsPerHour: 10,
    maxHistoryRequestsPerHour: 20,
    maxDashboardRequestsPerHour: 30
  };

  /** 24-column Vault layout (0-indexed for array access) */
  var VAULT_COLS = {
    TIMESTAMP: 0,
    EMAIL: 1,
    PRIORITY_CASES: 2,
    PENDING_CASES: 3,
    UNREAD_DOCS: 4,
    TODO_ITEMS: 5,
    SENT_REFERRALS: 6,
    CE_ACTIVITIES: 7,
    ASSISTANCE_REQUESTS: 8,
    AGED_CASES: 9,
    WEEKLY_CASES: 10,
    SUB_CATEGORIES: 11,
    EMPLOYMENT_TYPE: 12,
    PT_HOURS: 13,
    LEAVE_TYPE: 14,
    LEAVE_PLANNED: 15,
    LEAVE_START: 16,
    LEAVE_END: 17,
    NO_INTAKE_CHOICE: 18,
    NOTICE_TIME: 19,
    HALF_DAY: 20,
    PRIVACY: 21,
    ON_PLAN: 22,
    OVERTIME_HOURS: 23
  };

  var USERMETA_COLS = { EMAIL: 0, SHARING_START_DATE: 1, CREATED_DATE: 2 };
  var REMINDERS_COLS = { EMAIL: 0, ENABLED: 1, FREQUENCY: 2, DAY: 3, TIME: 4, LAST_SENT: 5 };

  var SUB_CATEGORIES = {
    priority: ['QDD', 'CAL', 'TERI', 'Aged Case', 'Congressional', 'Dire Need', 'Homeless',
               'Presumptive Disability', 'COBRA', 'DDS Aged', 'SOAR Involvement',
               'MC/WW', '100% P&T', 'Public Inquiry', 'DDS Important (High)', 'DDS Important (Medium)'],
    pending:  ['New Cases', 'Immediate Action', 'No Activity 15+ Days',
               'No Activity 30+ Days', 'Internal QA Returns', 'Federal QA Returns', 'Approval Returns'],
    unread:   ['All Unread Documents', 'MER', 'CE Reports', 'Trailer Mail'],
    todo:     ['Follow-Ups Due', 'Unread Case Notes', 'Delivery Failures', 'Updates'],
    referrals: ['MC/PC', 'General'],
    ce:       ['All CE Activities', 'Scheduled', 'Not Kept', 'Under Review'],
    assistance: ['Outbound'],
    aged:     ['60-89 Days', '90-119 Days', '120-179 Days', '180+ Days']
  };

  var CATEGORY_NAMES = [
    'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
    'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases'
  ];

  // ── Private Helpers ──────────────────────────────────────────────────────

  function _sanitizeString(input, maxLength) {
    maxLength = maxLength || APP_CONFIG.maxStringLength;
    if (!input || typeof input !== 'string') return '';
    var sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
    return sanitized.length > maxLength ? sanitized.substring(0, maxLength) : sanitized;
  }

  function _sanitizeNum(v) {
    return Math.min(999, Math.max(0, Math.floor(Number(v) || 0)));
  }

  function _withLock(fn) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(APP_CONFIG.lockTimeoutMs)) {
        throw new Error('System busy. Please try again in a moment.');
      }
      return fn();
    } finally {
      lock.releaseLock();
    }
  }

  // ── Rate Limiting ────────────────────────────────────────────────────────

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

  function _clearRateLimit(key) {
    CacheService.getScriptCache().remove('WT_RATE_' + key);
  }

  function _checkPinRateLimit(email) {
    return _checkRateLimit('PIN_' + email.toLowerCase().trim(),
      SECURITY_CONFIG.maxPinAttempts, SECURITY_CONFIG.pinLockoutMinutes);
  }

  function _recordFailedPinAttempt(email) {
    _recordRateLimitAttempt('PIN_' + email.toLowerCase().trim(), SECURITY_CONFIG.pinLockoutMinutes);
  }

  function _clearPinRateLimit(email) {
    _clearRateLimit('PIN_' + email.toLowerCase().trim());
  }

  function _checkSubmissionRateLimit(email) {
    return _checkRateLimit('SUBMIT_' + email.toLowerCase().trim(),
      SECURITY_CONFIG.maxSubmissionsPerHour, 60);
  }

  function _recordSubmission(email) {
    _recordRateLimitAttempt('SUBMIT_' + email.toLowerCase().trim(), 60);
  }

  // ── Reciprocity (Sharing Start Date) ─────────────────────────────────────

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

  // ── Reminder Persistence ─────────────────────────────────────────────────

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
          prefs.frequency || 'weekly', prefs.day || 'monday', prefs.time || '09:00'
        ]]);
        return;
      }
    }
    sheet.appendRow([emailLower, prefs.enabled ? 'Yes' : 'No',
      prefs.frequency || 'weekly', prefs.day || 'monday', prefs.time || '09:00', '']);
  }

  // ── PIN Authentication ───────────────────────────────────────────────────

  function _authenticateMember(email, pin) {
    if (!email || !pin) {
      return { success: false, memberId: null, message: 'Email and PIN are required.' };
    }
    var emailLower = email.toLowerCase().trim();

    var rateLimit = _checkPinRateLimit(emailLower);
    if (!rateLimit.allowed) {
      return { success: false, memberId: null,
        message: 'Too many failed attempts. Please wait ' + rateLimit.waitMinutes + ' minutes.' };
    }

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (!memberSheet) {
        return { success: false, memberId: null, message: 'System error: member directory unavailable.' };
      }

      var data = memberSheet.getDataRange().getValues();
      var emailCol = (MEMBER_COLS.EMAIL || 9) - 1;
      var idCol    = (MEMBER_COLS.MEMBER_ID || 1) - 1;
      var pinCol   = (MEMBER_COLS.PIN_HASH || 33) - 1;

      for (var i = 1; i < data.length; i++) {
        var rowEmail = data[i][emailCol];
        if (!rowEmail) continue;
        if (rowEmail.toString().toLowerCase().trim() !== emailLower) continue;

        var memberId   = data[i][idCol] ? data[i][idCol].toString().trim() : '';
        var storedHash = data[i][pinCol] ? data[i][pinCol].toString().trim() : '';

        if (!storedHash) {
          return { success: false, memberId: null,
            message: 'No PIN is set for this account. Please ask your steward to generate a PIN for you.' };
        }
        if (!memberId) {
          return { success: false, memberId: null, message: 'Member ID not found. Contact your steward.' };
        }
        if (verifyPIN(pin, memberId, storedHash)) {
          _clearPinRateLimit(emailLower);
          return { success: true, memberId: memberId, message: '' };
        } else {
          _recordFailedPinAttempt(emailLower);
          return { success: false, memberId: null, message: 'Invalid PIN. Please try again.' };
        }
      }

      _recordFailedPinAttempt(emailLower);
      return { success: false, memberId: null, message: 'Invalid credentials. Please try again.' };
    } catch (err) {
      console.error('_authenticateMember error:', err);
      return { success: false, memberId: null, message: 'Authentication error. Please try again.' };
    }
  }

  // ── Reporting Refresh ────────────────────────────────────────────────────

  function _refreshReportingData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault  = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
    if (!vault || !report) return;

    var vaultData = vault.getDataRange().getValues();
    if (vaultData.length <= 1) return;

    // Determine each user's latest privacy setting
    var latestPrivacy = {};
    for (var i = 1; i < vaultData.length; i++) {
      var rowEmail = vaultData[i][VAULT_COLS.EMAIL];
      if (rowEmail) {
        latestPrivacy[rowEmail.toString().toLowerCase().trim()] = vaultData[i][VAULT_COLS.PRIVACY];
      }
    }

    var rows = [];
    for (var j = 1; j < vaultData.length; j++) {
      var rowE = vaultData[j][VAULT_COLS.EMAIL];
      if (!rowE) continue;
      var _emailKey = rowE.toString().toLowerCase().trim();
      var rowPrivacy = vaultData[j][VAULT_COLS.PRIVACY];

      var identity = 'REDACTED';
      if (rowPrivacy === 'Unit') identity = 'REDACTED (Unit)';
      else if (rowPrivacy === 'Agency') identity = 'REDACTED (Agency)';

      rows.push([
        vaultData[j][VAULT_COLS.TIMESTAMP], identity,
        vaultData[j][VAULT_COLS.PRIORITY_CASES], vaultData[j][VAULT_COLS.PENDING_CASES],
        vaultData[j][VAULT_COLS.UNREAD_DOCS], vaultData[j][VAULT_COLS.TODO_ITEMS],
        vaultData[j][VAULT_COLS.SENT_REFERRALS], vaultData[j][VAULT_COLS.CE_ACTIVITIES],
        vaultData[j][VAULT_COLS.ASSISTANCE_REQUESTS], vaultData[j][VAULT_COLS.AGED_CASES],
        vaultData[j][VAULT_COLS.WEEKLY_CASES], vaultData[j][VAULT_COLS.EMPLOYMENT_TYPE],
        rowPrivacy, vaultData[j][VAULT_COLS.ON_PLAN], vaultData[j][VAULT_COLS.OVERTIME_HOURS]
      ]);
    }

    report.clearContents();
    report.appendRow(['Date', 'Identity', 'Priority Cases', 'Pending Cases', 'Unread Documents',
      'To-Do Items', 'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'Weekly Cases', 'Employment Type', 'Privacy', 'On Plan', 'Overtime Hours']);
    report.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#0d47a1').setFontColor('#ffffff');
    if (rows.length > 0) {
      report.getRange(2, 1, rows.length, 15).setValues(rows);
    }
  }

  // ── Vault Cleaning (deduplication) ───────────────────────────────────────

  function _cleanVaultData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault || vault.getLastRow() <= 1) return { removed: 0 };

    var data = vault.getDataRange().getValues();
    var header = data[0];
    var seen = {};
    var clean = [header];
    var removed = 0;

    for (var i = 1; i < data.length; i++) {
      var ts = data[i][VAULT_COLS.TIMESTAMP];
      var em = data[i][VAULT_COLS.EMAIL];
      var key = String(ts) + '|' + String(em).toLowerCase().trim();
      if (seen[key]) {
        removed++;
      } else {
        seen[key] = true;
        clean.push(data[i]);
      }
    }

    if (removed > 0) {
      vault.clear();
      vault.getRange(1, 1, clean.length, header.length).setValues(clean);
      vault.getRange(1, 1, 1, header.length).setFontWeight('bold')
        .setBackground('#1a1a2e').setFontColor('#ffffff');
      _refreshReportingData();
    }

    return { removed: removed };
  }

  // ── Frequency Helper ─────────────────────────────────────────────────────

  function _isReminderDue(freq, day, lastSent) {
    var now = new Date();
    var dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    var todayName = dayNames[now.getDay()];

    if (freq === 'daily') return true;

    if (freq === 'weekly') {
      return (day || 'monday').toLowerCase() === todayName;
    }

    if (freq === 'biweekly') {
      if ((day || 'monday').toLowerCase() !== todayName) return false;
      if (!lastSent) return true;
      var daysSince = Math.floor((now - new Date(lastSent)) / 86400000);
      return daysSince >= 13;
    }

    if (freq === 'monthly') {
      if (!lastSent) return now.getDate() <= 7;
      var monthsSince = (now.getFullYear() - new Date(lastSent).getFullYear()) * 12 +
        now.getMonth() - new Date(lastSent).getMonth();
      return monthsSince >= 1;
    }

    if (freq === 'quarterly') {
      if (!lastSent) return true;
      var qMonths = (now.getFullYear() - new Date(lastSent).getFullYear()) * 12 +
        now.getMonth() - new Date(lastSent).getMonth();
      return qMonths >= 3;
    }

    return false;
  }

  // ========================================================================
  // PUBLIC API
  // ========================================================================

  return {

    // ── Sheet Initialization ─────────────────────────────────────────────

    initSheets: function () {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // Workload Vault (hidden, 24 columns)
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
        vault.getRange(1, 1, 1, 24).setFontWeight('bold')
          .setBackground('#1a1a2e').setFontColor('#ffffff');
        vault.setFrozenRows(1);
      }
      vault.hideSheet();

      // Workload Reporting (visible, anonymized)
      var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
      if (!report) {
        report = ss.insertSheet(SHEETS.WORKLOAD_REPORTING);
        report.appendRow([
          'Date', 'Identity',
          'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
          'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
          'Weekly Cases', 'Employment Type', 'Privacy', 'On Plan', 'Overtime Hours'
        ]);
        report.getRange(1, 1, 1, 15).setFontWeight('bold')
          .setBackground('#0d47a1').setFontColor('#ffffff');
        report.setFrozenRows(1);
        report.setHiddenGridlines(true);
      }

      // Workload Reminders
      var reminders = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
      if (!reminders) {
        reminders = ss.insertSheet(SHEETS.WORKLOAD_REMINDERS);
        reminders.appendRow(['Email', 'Enabled', 'Frequency', 'Day', 'Time', 'Last Sent']);
        reminders.getRange(1, 1, 1, 6).setFontWeight('bold')
          .setBackground('#1a1a2e').setFontColor('#ffffff');
        reminders.setFrozenRows(1);
        reminders.hideSheet();
      }

      // Workload UserMeta (hidden)
      var userMeta = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
      if (!userMeta) {
        userMeta = ss.insertSheet(SHEETS.WORKLOAD_USERMETA);
        userMeta.appendRow(['Email', 'Sharing Start Date', 'Created Date']);
        userMeta.getRange(1, 1, 1, 3).setFontWeight('bold')
          .setBackground('#1a1a2e').setFontColor('#ffffff');
        userMeta.setFrozenRows(1);
        userMeta.hideSheet();
      }

      Logger.log('Workload Tracker sheets initialized.');
    },

    // ── Form Submission (PIN auth) ───────────────────────────────────────

    processForm: function (formObj) {
      var email = formObj.email ? _sanitizeString(formObj.email.trim().toLowerCase(), 100) : '';

      try {
        var pin = formObj.pin ? formObj.pin.toString().trim() : '';
        if (!email || !pin) return 'Error: Please provide your email and PIN.';

        var emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRe.test(email)) return 'Error: Please enter a valid email address.';

        var submitLimit = _checkSubmissionRateLimit(email);
        if (!submitLimit.allowed) return 'Error: Submission limit reached. Please wait before submitting again.';

        var auth = _authenticateMember(email, pin);
        if (!auth.success) return 'Error: ' + auth.message;

        // Validate numeric t1-t8
        var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
        for (var n = 0; n < numFields.length; n++) {
          var val = Number(formObj[numFields[n]]) || 0;
          if (val < 0 || val > 999 || !Number.isInteger(val)) {
            return 'Error: Workload values must be whole numbers between 0 and 999.';
          }
        }

        if (formObj.employment_type === 'Part-time') {
          var ptH = Number(formObj.part_time_hours);
          if (isNaN(ptH) || ptH < 1 || ptH > 40 || !Number.isInteger(ptH)) {
            return 'Error: Part-time hours must be a whole number between 1 and 40.';
          }
        }

        return _withLock(function () {
          var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
          if (!vault) return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';

          // Build sub-category JSON
          var subCatData = {};
          var catKeys = ['priority', 'pending', 'unread', 'todo', 'referrals', 'ce', 'assistance', 'aged'];
          for (var ci = 0; ci < catKeys.length; ci++) {
            subCatData[catKeys[ci]] = {};
            var subNames = SUB_CATEGORIES[catKeys[ci]];
            for (var si = 0; si < subNames.length; si++) {
              var fieldName = 'sub_' + (ci + 1) + '_' + si;
              var subVal = _sanitizeNum(formObj[fieldName]);
              if (subVal > 0) subCatData[catKeys[ci]][subNames[si]] = subVal;
            }
          }

          // Weekly cases
          var weeklyCases = _sanitizeNum(formObj.weekly_cases || 15);
          if (formObj.weekly_cases_option === 'manual' && formObj.weekly_cases_manual) {
            weeklyCases = _sanitizeNum(formObj.weekly_cases_manual);
          }

          // Leave dates
          var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
          var leaveStart = '', leaveEnd = '';
          if (formObj.leave_start) {
            leaveStart = Utilities.formatDate(new Date(formObj.leave_start), tz, 'MM/dd/yyyy');
          }
          if (formObj.leave_end) {
            leaveEnd = Utilities.formatDate(new Date(formObj.leave_end), tz, 'MM/dd/yyyy');
          }

          vault.appendRow([
            new Date(), email,
            _sanitizeNum(formObj.t1), _sanitizeNum(formObj.t2),
            _sanitizeNum(formObj.t3), _sanitizeNum(formObj.t4),
            _sanitizeNum(formObj.t5), _sanitizeNum(formObj.t6),
            _sanitizeNum(formObj.t7), _sanitizeNum(formObj.t8),
            weeklyCases, JSON.stringify(subCatData),
            _sanitizeString(formObj.employment_type || 'Full-time', 20),
            formObj.employment_type === 'Part-time' ? (Number(formObj.part_time_hours) || '') : '',
            _sanitizeString(formObj.leave_type || '', 20),
            _sanitizeString(formObj.leave_planned || '', 20),
            leaveStart, leaveEnd,
            _sanitizeString(formObj.no_intake_choice || '', 20),
            _sanitizeString(formObj.notice_time || '', 20),
            formObj.half_day_leave ? 'Yes' : '',
            _sanitizeString(formObj.privacy || 'Unit', 20),
            formObj.on_plan ? 'Yes' : 'No',
            formObj.overtime_enabled ? (Number(formObj.overtime_hours) || 0) : ''
          ]);

          if (typeof logAuditEvent === 'function') {
            logAuditEvent('WORKLOAD_SUBMIT',
              'Member ' + auth.memberId + ' submitted workload data. Privacy: ' +
              (formObj.privacy || 'Unit'));
          }

          // Reciprocity: set sharing start date
          var _privacySetting = _sanitizeString(formObj.privacy || 'Unit', 20);
          if (!_getUserSharingStartDate(email)) {
            _setUserSharingStartDate(email, new Date());
          }

          _recordSubmission(email);

          if (formObj.reminder_enabled !== undefined) {
            _saveReminderPreference({
              email: email,
              enabled: formObj.reminder_enabled === 'on' || formObj.reminder_enabled === true,
              frequency: formObj.reminder_frequency || 'weekly',
              day: formObj.reminder_day || 'monday',
              time: formObj.reminder_time || '09:00'
            });
          }

          _refreshReportingData();
          return 'Success! Your workload data has been securely recorded.';
        });

      } catch (err) {
        console.error('processForm error:', err);
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('WORKLOAD_SUBMIT_ERROR', 'Error for ' + email + ': ' + err.message);
        }
        return 'Error: Unable to process submission. Please try again.';
      }
    },

    // ── Dashboard Data (collective stats, PIN auth) ──────────────────────

    getDashboardData: function (email, pin) {
      if (!email || !pin) return { success: false, data: null, message: 'Credentials required.' };

      var auth = _authenticateMember(email, pin);
      if (!auth.success) return { success: false, data: null, message: auth.message };

      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault || vault.getLastRow() <= 1) {
        return { success: true, data: { totalSubmissions: 0, members: 0, categories: {}, averages: {} }, message: '' };
      }

      var data = vault.getDataRange().getValues();
      var emailLower = email.toLowerCase().trim();

      var userSharingStart = _getUserSharingStartDate(emailLower);
      var totals = [0, 0, 0, 0, 0, 0, 0, 0]; // t1-t8
      var members = {};
      var submissionCount = 0;
      var empBreakdown = {};
      var planCount = 0;
      var otTotal = 0, otCount = 0;
      var subCatTotals = {};

      for (var i = 1; i < data.length; i++) {
        var ts = new Date(data[i][VAULT_COLS.TIMESTAMP]);
        if (userSharingStart && ts < userSharingStart) continue;

        submissionCount++;
        var rowEmail = data[i][VAULT_COLS.EMAIL];
        if (rowEmail) members[rowEmail.toString().toLowerCase()] = true;

        totals[0] += Number(data[i][VAULT_COLS.PRIORITY_CASES]) || 0;
        totals[1] += Number(data[i][VAULT_COLS.PENDING_CASES]) || 0;
        totals[2] += Number(data[i][VAULT_COLS.UNREAD_DOCS]) || 0;
        totals[3] += Number(data[i][VAULT_COLS.TODO_ITEMS]) || 0;
        totals[4] += Number(data[i][VAULT_COLS.SENT_REFERRALS]) || 0;
        totals[5] += Number(data[i][VAULT_COLS.CE_ACTIVITIES]) || 0;
        totals[6] += Number(data[i][VAULT_COLS.ASSISTANCE_REQUESTS]) || 0;
        totals[7] += Number(data[i][VAULT_COLS.AGED_CASES]) || 0;

        // Employment breakdown
        var emp = data[i][VAULT_COLS.EMPLOYMENT_TYPE] || 'Full-time';
        empBreakdown[emp] = (empBreakdown[emp] || 0) + 1;

        // On-plan count
        if (data[i][VAULT_COLS.ON_PLAN] === 'Yes') planCount++;

        // Overtime
        var otHrs = Number(data[i][VAULT_COLS.OVERTIME_HOURS]);
        if (otHrs > 0) { otTotal += otHrs; otCount++; }

        // Sub-category aggregation
        var scJson = data[i][VAULT_COLS.SUB_CATEGORIES];
        if (scJson) {
          try {
            var sc = JSON.parse(scJson);
            for (var catKey in sc) {
              if (!subCatTotals[catKey]) subCatTotals[catKey] = {};
              for (var subName in sc[catKey]) {
                subCatTotals[catKey][subName] = (subCatTotals[catKey][subName] || 0) + (Number(sc[catKey][subName]) || 0);
              }
            }
          } catch (_e) { /* skip malformed JSON */ }
        }
      }

      var memberCount = Object.keys(members).length;
      var categories = {};
      var averages = {};
      for (var c = 0; c < CATEGORY_NAMES.length; c++) {
        categories[CATEGORY_NAMES[c]] = totals[c];
        averages[CATEGORY_NAMES[c]] = submissionCount > 0 ? (totals[c] / submissionCount).toFixed(1) : '0';
      }

      return {
        success: true, message: '',
        data: {
          totalSubmissions: submissionCount,
          members: memberCount,
          categories: categories,
          averages: averages,
          employmentBreakdown: empBreakdown,
          onPlanCount: planCount,
          onPlanPct: submissionCount > 0 ? Math.round(planCount / submissionCount * 100) : 0,
          overtime: { totalHours: otTotal, submissions: otCount,
            avgPerSubmission: otCount > 0 ? (otTotal / otCount).toFixed(1) : '0',
            pctReporting: submissionCount > 0 ? Math.round(otCount / submissionCount * 100) : 0 },
          subCategoryTotals: subCatTotals,
          sharingStartDate: userSharingStart ? userSharingStart.toISOString() : null
        }
      };
    },

    // ── User History (PIN auth) ──────────────────────────────────────────

    getUserHistory: function (email, pin) {
      if (!email || !pin) return { success: false, history: [], message: 'Credentials required.' };

      var auth = _authenticateMember(email, pin);
      if (!auth.success) return { success: false, history: [], message: auth.message };

      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault || vault.getLastRow() <= 1) return { success: true, history: [], message: '' };

      var emailLower = email.toLowerCase().trim();
      var data = vault.getDataRange().getValues();
      var history = [];

      for (var i = 1; i < data.length; i++) {
        var rowEmail = data[i][VAULT_COLS.EMAIL];
        if (!rowEmail || rowEmail.toString().toLowerCase().trim() !== emailLower) continue;

        var entry = {
          date:        data[i][VAULT_COLS.TIMESTAMP],
          t1:          data[i][VAULT_COLS.PRIORITY_CASES],
          t2:          data[i][VAULT_COLS.PENDING_CASES],
          t3:          data[i][VAULT_COLS.UNREAD_DOCS],
          t4:          data[i][VAULT_COLS.TODO_ITEMS],
          t5:          data[i][VAULT_COLS.SENT_REFERRALS],
          t6:          data[i][VAULT_COLS.CE_ACTIVITIES],
          t7:          data[i][VAULT_COLS.ASSISTANCE_REQUESTS],
          t8:          data[i][VAULT_COLS.AGED_CASES],
          weeklyCases: data[i][VAULT_COLS.WEEKLY_CASES],
          employment:  data[i][VAULT_COLS.EMPLOYMENT_TYPE],
          ptHours:     data[i][VAULT_COLS.PT_HOURS],
          leaveType:   data[i][VAULT_COLS.LEAVE_TYPE],
          leavePlanned: data[i][VAULT_COLS.LEAVE_PLANNED],
          leaveStart:  data[i][VAULT_COLS.LEAVE_START],
          leaveEnd:    data[i][VAULT_COLS.LEAVE_END],
          halfDay:     data[i][VAULT_COLS.HALF_DAY],
          privacy:     data[i][VAULT_COLS.PRIVACY],
          onPlan:      data[i][VAULT_COLS.ON_PLAN],
          overtime:    data[i][VAULT_COLS.OVERTIME_HOURS]
        };

        // Parse sub-categories
        var scJson = data[i][VAULT_COLS.SUB_CATEGORIES];
        if (scJson) {
          try { entry.subCategories = JSON.parse(scJson); } catch (_e) { entry.subCategories = {}; }
        } else {
          entry.subCategories = {};
        }

        history.push(entry);
      }

      history.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
      return { success: true, history: history, message: '' };
    },

    // ── CSV Export (PIN auth) ────────────────────────────────────────────

    exportHistoryCSV: function (email, pin) {
      if (!email || !pin) return { success: false, csv: '', message: 'Credentials required.' };

      var auth = _authenticateMember(email, pin);
      if (!auth.success) return { success: false, csv: '', message: auth.message };

      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault || vault.getLastRow() <= 1) return { success: true, csv: '', message: 'No data.' };

      var emailLower = email.toLowerCase().trim();
      var data = vault.getDataRange().getValues();
      var header = 'Date,Priority Cases,Pending Cases,Unread Docs,To-Do Items,Sent Referrals,' +
        'CE Activities,Assistance Requests,Aged Cases,Weekly Cases,Employment,Privacy,On Plan,Overtime Hours';
      var csvRows = [header];

      for (var i = 1; i < data.length; i++) {
        var re = data[i][VAULT_COLS.EMAIL];
        if (!re || re.toString().toLowerCase().trim() !== emailLower) continue;
        var ts = data[i][VAULT_COLS.TIMESTAMP];
        var dateStr = ts instanceof Date ? Utilities.formatDate(ts,
          SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : String(ts);
        csvRows.push([
          dateStr, data[i][VAULT_COLS.PRIORITY_CASES], data[i][VAULT_COLS.PENDING_CASES],
          data[i][VAULT_COLS.UNREAD_DOCS], data[i][VAULT_COLS.TODO_ITEMS],
          data[i][VAULT_COLS.SENT_REFERRALS], data[i][VAULT_COLS.CE_ACTIVITIES],
          data[i][VAULT_COLS.ASSISTANCE_REQUESTS], data[i][VAULT_COLS.AGED_CASES],
          data[i][VAULT_COLS.WEEKLY_CASES], data[i][VAULT_COLS.EMPLOYMENT_TYPE],
          data[i][VAULT_COLS.PRIVACY], data[i][VAULT_COLS.ON_PLAN],
          data[i][VAULT_COLS.OVERTIME_HOURS]
        ].join(','));
      }

      return { success: true, csv: csvRows.join('\n'), message: '' };
    },

    // ── Sub-category Definitions ─────────────────────────────────────────

    getSubCategories: function () {
      return SUB_CATEGORIES;
    },

    // ── Ledger Refresh ───────────────────────────────────────────────────

    refreshLedger: function () {
      _refreshReportingData();
    },

    // ── Backup ───────────────────────────────────────────────────────────

    createBackup: function () {
      var ui = SpreadsheetApp.getUi();
      try {
        var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
        if (!vault) { ui.alert('Workload Vault sheet not found.'); return; }

        var data = vault.getDataRange().getValues();
        var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
        var timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
        var fileName = 'WorkloadVault_Backup_' + timestamp + '.csv';

        var csv = data.map(function (row) {
          return row.map(function (cell) {
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
          catch (_e) { folder = DriveApp.createFolder('WorkloadVault_Backups'); props.setProperty('WT_BACKUP_FOLDER_ID', folder.getId()); }
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
        console.error('createBackup error:', err);
        ui.alert('Backup Failed', 'Error: ' + err.message, ui.ButtonSet.OK);
      }
    },

    // ── Archive ──────────────────────────────────────────────────────────

    archiveOldData: function () {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Archive Old Workload Data',
        'Move workload data older than ' + APP_CONFIG.dataRetentionMonths +
        ' months to the Archive sheet?', ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) return;

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault) { ui.alert('Workload Vault not found.'); return; }

      var data = vault.getDataRange().getValues();
      if (data.length <= 1) { ui.alert('No data to archive.'); return; }

      var cutoff = new Date();
      cutoff.setMonth(cutoff.getMonth() - APP_CONFIG.dataRetentionMonths);

      var header = data[0];
      var current = [header];
      var archive = [];
      for (var i = 1; i < data.length; i++) {
        if (new Date(data[i][VAULT_COLS.TIMESTAMP]) < cutoff) archive.push(data[i]);
        else current.push(data[i]);
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
    },

    // ── Vault Cleaning ───────────────────────────────────────────────────

    cleanVault: function () {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Clean Workload Vault',
        'Remove duplicate entries from the Workload Vault?', ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) return;

      var result = _cleanVaultData();
      ui.alert('Vault cleaned: ' + result.removed + ' duplicate(s) removed.', '', ui.ButtonSet.OK);
    },

    // ── Health Status ────────────────────────────────────────────────────

    getHealthStatus: function () {
      var ui = SpreadsheetApp.getUi();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetNames = [
        SHEETS.WORKLOAD_VAULT, SHEETS.WORKLOAD_REPORTING,
        SHEETS.WORKLOAD_REMINDERS, SHEETS.WORKLOAD_USERMETA
      ];
      var lines = ['Workload Tracker v' + APP_CONFIG.version + '\n'];
      sheetNames.forEach(function (name) {
        var s = ss.getSheetByName(name);
        var rows = s ? Math.max(0, s.getLastRow() - 1) : 0;
        lines.push((s ? '+ ' : '- ') + name + (s ? ' (' + rows + ' records)' : ' -- missing'));
      });
      lines.push('\nSetup: Admin > Setup > Initialize Dashboard to create missing sheets.');
      ui.alert('Workload Tracker Health', lines.join('\n'), ui.ButtonSet.OK);
    },

    // ── Reminder Setup ───────────────────────────────────────────────────

    setupReminders: function () {
      var triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(function (t) {
        if (t.getHandlerFunction() === 'processWorkloadReminders') ScriptApp.deleteTrigger(t);
      });
      ScriptApp.newTrigger('processWorkloadReminders')
        .timeBased().everyDays(1).atHour(8).create();
      SpreadsheetApp.getUi().alert('Workload reminder system activated (daily at 8 AM).');
    },

    // ── Reminder Processing ──────────────────────────────────────────────

    processReminders: function () {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
      if (!sheet || sheet.getLastRow() <= 1) return;

      var data = sheet.getDataRange().getValues();
      var now = new Date();
      var portalUrl = PropertiesService.getScriptProperties().getProperty('WT_PORTAL_URL') || '';

      for (var i = 1; i < data.length; i++) {
        try {
          var rEmail   = data[i][REMINDERS_COLS.EMAIL];
          var rEnabled = data[i][REMINDERS_COLS.ENABLED];
          var rFreq    = (data[i][REMINDERS_COLS.FREQUENCY] || 'weekly').toLowerCase();
          var rDay     = (data[i][REMINDERS_COLS.DAY] || 'monday').toLowerCase();
          var rLastSent = data[i][REMINDERS_COLS.LAST_SENT];

          if (!rEmail || rEnabled !== 'Yes') continue;
          if (!_isReminderDue(rFreq, rDay, rLastSent)) continue;

          var subject = 'Reminder: Submit Your Workload Data';
          var body = 'Hello,\n\nThis is your scheduled reminder to submit your workload data.\n\n' +
            (portalUrl ? 'Access the portal here:\n' + portalUrl + '?page=workload\n\n' : '') +
            'Thank you,\nSEIU 509 DDS Dashboard';

          MailApp.sendEmail({ to: rEmail, subject: subject, body: body });
          sheet.getRange(i + 1, REMINDERS_COLS.LAST_SENT + 1).setValue(now);
        } catch (err) {
          console.error('Reminder error for row ' + i + ': ' + err.message);
        }
      }
    },

    // ── Portal HTML ──────────────────────────────────────────────────────

    getPortalHtml: function () {
      try {
        return HtmlService.createHtmlOutputFromFile('WorkloadTracker').getContent();
      } catch (_err) {
        return '<html><body style="font-family:sans-serif;padding:2rem;">' +
          '<h2>Workload Tracker</h2>' +
          '<p>Portal not yet configured. Please run <strong>Initialize Dashboard</strong> first.</p>' +
          '</body></html>';
      }
    },

    // ── Portal URL Menu Helpers ──────────────────────────────────────────

    showPortalUrl: function () {
      var baseUrl = PropertiesService.getScriptProperties().getProperty('WT_PORTAL_URL') || '';
      var ui = SpreadsheetApp.getUi();
      if (baseUrl) {
        ui.alert('Workload Portal URL', baseUrl + '?page=workload\n\nShare this link with members.', ui.ButtonSet.OK);
      } else {
        ui.alert('Not configured', 'Deploy the web app, then use "Share Portal Link" to save its URL.', ui.ButtonSet.OK);
      }
    },

    sharePortalLink: function () {
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

  };  // end return

})();

// ============================================================================
// GLOBAL WRAPPERS — Portal-specific (PIN auth)
// Shared admin wrappers (initSheets, refreshLedger, backup, etc.) are in
// 25_WorkloadService.gs to avoid duplicate function names in GAS.
// ============================================================================

/** Process workload form submission (standalone portal — PIN auth). */
function processWorkloadForm(formObj) { return WorkloadPortal.processForm(formObj); }

/** Get collective dashboard data (standalone portal — PIN auth). */
function getWorkloadDashboardData(email, pin) { return WorkloadPortal.getDashboardData(email, pin); }

/** Get personal submission history (standalone portal — PIN auth). */
function getWorkloadUserHistory(email, pin) { return WorkloadPortal.getUserHistory(email, pin); }

/** Export personal history as CSV — portal version with PIN auth. */
function exportWorkloadHistoryPortal(email, pin) { return WorkloadPortal.exportHistoryCSV(email, pin); }

/** Refresh anonymized reporting sheet (internal use — portal module). */
function refreshWorkloadReportingData() { return WorkloadPortal.refreshLedger(); }

/** Get portal HTML for doGet routing. */
function getWorkloadTrackerPortalHtml() { return WorkloadPortal.getPortalHtml(); }

/** Show portal URL dialog (menu item). */
function showWorkloadPortalUrl() { return WorkloadPortal.showPortalUrl(); }

/** Save portal URL from user prompt (menu item). */
function shareWorkloadPortalLink() { return WorkloadPortal.sharePortalLink(); }
