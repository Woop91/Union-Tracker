/**
 * ============================================================================
 * 18_WorkloadTracker.gs - Member Workload Tracking Module
 * ============================================================================
 *
 * Secure, anonymized workload tracking integrated with the DDS Dashboard.
 * Members submit weekly caseload data via the web portal (?page=workload).
 * Authentication uses the existing DDS member PIN system (13_MemberSelfService.gs).
 *
 * Sheet Structure:
 *   "Workload Vault"     (hidden)  — raw data with member email + PIN hash
 *   "Workload Reporting" (visible) — anonymized ledger (identity → REDACTED)
 *   "Workload Reminders"           — email reminder preferences
 *   "Workload UserMeta"  (hidden)  — sharing start dates for reciprocity
 *   "Workload Archive"   (hidden)  — data older than retention period
 *
 * Auth: email + DDS member PIN → verifyPIN(pin, memberId, storedHash)
 *
 * @version 4.10.0
 * @requires 01_Core.gs (SHEETS, MEMBER_COLS)
 * @requires 13_MemberSelfService.gs (verifyPIN, hashPIN, getPINSalt_)
 * @requires 06_Maintenance.gs (logAuditEvent)
 * @requires 00_Security.gs (escapeHtml)
 */

// ============================================================================
// WORKLOAD TRACKER CONFIGURATION
// ============================================================================

/** @const {Object} Workload tracker app configuration */
var WT_APP_CONFIG = {
  version: '4.11.0',
  name: 'Workload Tracker',
  maxStringLength: 255,
  dataRetentionMonths: 24,
  lockTimeoutMs: 30000
};

/** @const {Object} Workload tracker security / rate-limit settings */
var WT_SECURITY_CONFIG = {
  maxPinAttempts: 5,
  pinLockoutMinutes: 15,
  maxSubmissionsPerHour: 10,
  maxHistoryRequestsPerHour: 20,
  maxDashboardRequestsPerHour: 30
};

// ============================================================================
// COLUMN CONSTANTS (0-indexed for array access)
// ============================================================================

/** @const {Object} Workload Vault column indices (0-indexed) */
var WT_VAULT_COLS = {
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

/** @const {Object} Workload UserMeta column indices (0-indexed) */
var WT_USERMETA_COLS = {
  EMAIL: 0,
  SHARING_START_DATE: 1,
  CREATED_DATE: 2
};

/** @const {Object} Workload Reminders column indices (0-indexed) */
var WT_REMINDERS_COLS = {
  EMAIL: 0,
  ENABLED: 1,
  FREQUENCY: 2,
  DAY: 3,
  TIME: 4,
  LAST_SENT: 5
};

// ============================================================================
// SUB-CATEGORY DEFINITIONS
// ============================================================================

/** @const {Object} Sub-category names for each workload category */
var WT_SUB_CATEGORIES = {
  priority: ['QDD', 'CAL', 'TERI', 'Aged Case', 'Congressional', 'Dire Need', 'Homeless',
             'Presumptive Disability', 'COBRA', 'DDS Aged', 'SOAR Involvement',
             'MC/WW', '100% P&T', 'Public Inquiry', 'DDS Important (High)', 'DDS Important (Medium)'],
  pending:  ['New Cases', 'Immediate Action', 'No Activity 15+ Days',
             'No Activity 30+ Days', 'Internal QA Returns', 'Federal QA Returns', 'Approval Returns'],
  unread:   ['All Unread Documents', 'MER', 'CE Reports', 'Trailer Mail'],
  todo:     ['Follow-Ups Due', 'Unread Case Notes', 'Delivery Failures', 'Updates'],
  referrals:['MC/PC', 'General'],
  ce:       ['All CE Activities', 'Scheduled', 'Not Kept', 'Under Review'],
  assistance:['Outbound'],
  aged:     ['60-89 Days', '90-119 Days', '120-179 Days', '180+ Days']
};

// ============================================================================
// INPUT SANITIZATION
// ============================================================================

/**
 * Sanitizes a string input: trims, limits length, strips control characters.
 * @param {string} input
 * @param {number} [maxLength]
 * @returns {string}
 */
function sanitizeString(input, maxLength) {
  maxLength = maxLength || WT_APP_CONFIG.maxStringLength;
  if (!input || typeof input !== 'string') return '';
  var sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').trim();
  return sanitized.length > maxLength ? sanitized.substring(0, maxLength) : sanitized;
}

// ============================================================================
// RATE LIMITING
// ============================================================================

/**
 * Checks if an action is rate-limited using CacheService.
 * @param {string} key
 * @param {number} maxAttempts
 * @param {number} windowMinutes
 * @returns {{allowed:boolean, attemptsRemaining:number, waitMinutes:number}}
 */
function wtCheckRateLimit_(key, maxAttempts, windowMinutes) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WT_RATE_' + key;
  var cached = cache.get(cacheKey);
  var attempts = cached ? (parseInt(cached, 10) || 0) : 0;
  if (attempts >= maxAttempts) {
    return { allowed: false, attemptsRemaining: 0, waitMinutes: windowMinutes };
  }
  return { allowed: true, attemptsRemaining: maxAttempts - attempts, waitMinutes: 0 };
}

/**
 * Records an attempt for rate limiting.
 * @param {string} key
 * @param {number} windowMinutes
 */
function wtRecordRateLimitAttempt_(key, windowMinutes) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'WT_RATE_' + key;
  var cached = cache.get(cacheKey);
  var attempts = (cached ? (parseInt(cached, 10) || 0) : 0) + 1;
  cache.put(cacheKey, String(attempts), windowMinutes * 60);
}

/** Clears rate limit for a key. @param {string} key */
function wtClearRateLimit_(key) {
  CacheService.getScriptCache().remove('WT_RATE_' + key);
}

/** @param {string} email @returns {object} */
function wtCheckPinRateLimit_(email) {
  return wtCheckRateLimit_('PIN_' + email.toLowerCase().trim(),
    WT_SECURITY_CONFIG.maxPinAttempts, WT_SECURITY_CONFIG.pinLockoutMinutes);
}

/** @param {string} email */
function wtRecordFailedPinAttempt_(email) {
  wtRecordRateLimitAttempt_('PIN_' + email.toLowerCase().trim(), WT_SECURITY_CONFIG.pinLockoutMinutes);
}

/** @param {string} email */
function wtClearPinRateLimit_(email) {
  wtClearRateLimit_('PIN_' + email.toLowerCase().trim());
}

/** @param {string} email @returns {object} */
function wtCheckSubmissionRateLimit_(email) {
  return wtCheckRateLimit_('SUBMIT_' + email.toLowerCase().trim(),
    WT_SECURITY_CONFIG.maxSubmissionsPerHour, 60);
}

/** @param {string} email */
function wtRecordSubmission_(email) {
  wtRecordRateLimitAttempt_('SUBMIT_' + email.toLowerCase().trim(), 60);
}

// ============================================================================
// LOCK SERVICE
// ============================================================================

/**
 * Executes a function with a script lock to prevent concurrent writes.
 * @param {Function} fn
 * @param {string} [lockName]
 * @returns {*}
 */
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
// AUTHENTICATION (uses DDS member PIN system)
// ============================================================================

/**
 * Authenticates a member for workload submission using their DDS member PIN.
 * Looks up the member by email in the Member Directory, then verifies the PIN
 * using verifyPIN() from 13_MemberSelfService.gs.
 *
 * @param {string} email - Member's email address
 * @param {string} pin   - Member's 6-digit DDS PIN
 * @returns {{success:boolean, memberId:string|null, message:string}}
 */
function authenticateWorkloadMember_(email, pin) {
  if (!email || !pin) {
    return { success: false, memberId: null, message: 'Email and PIN are required.' };
  }

  var emailLower = email.toLowerCase().trim();

  // PIN rate-limit check
  var rateLimit = wtCheckPinRateLimit_(emailLower);
  if (!rateLimit.allowed) {
    return {
      success: false,
      memberId: null,
      message: 'Too many failed attempts. Please wait ' + rateLimit.waitMinutes + ' minutes.'
    };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet) {
      return { success: false, memberId: null, message: 'System error: member directory unavailable.' };
    }

    var data = memberSheet.getDataRange().getValues();
    var emailCol = (MEMBER_COLS.EMAIL || 9) - 1;      // 0-indexed
    var idCol    = (MEMBER_COLS.MEMBER_ID || 1) - 1;  // 0-indexed
    var pinCol   = (MEMBER_COLS.PIN_HASH || 33) - 1;  // 0-indexed (column AG)

    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][emailCol];
      if (!rowEmail) continue;
      if (rowEmail.toString().toLowerCase().trim() !== emailLower) continue;

      var memberId   = data[i][idCol] ? data[i][idCol].toString().trim() : '';
      var storedHash = data[i][pinCol] ? data[i][pinCol].toString().trim() : '';

      if (!storedHash) {
        return {
          success: false,
          memberId: null,
          message: 'No PIN is set for this account. Please ask your steward to generate a PIN for you.'
        };
      }

      if (!memberId) {
        return { success: false, memberId: null, message: 'Member ID not found. Contact your steward.' };
      }

      // Use DDS's constant-time PIN verification
      if (verifyPIN(pin, memberId, storedHash)) {
        wtClearPinRateLimit_(emailLower);
        return { success: true, memberId: memberId, message: '' };
      } else {
        wtRecordFailedPinAttempt_(emailLower);
        return { success: false, memberId: null, message: 'Invalid PIN. Please try again.' };
      }
    }

    // Email not in directory (use generic message to prevent enumeration)
    wtRecordFailedPinAttempt_(emailLower);
    return { success: false, memberId: null, message: 'Invalid credentials. Please try again.' };

  } catch (err) {
    console.error('authenticateWorkloadMember_ error:', err);
    return { success: false, memberId: null, message: 'Authentication error. Please try again.' };
  }
}

// ============================================================================
// SHEET INITIALIZATION
// ============================================================================

/**
 * Creates all Workload Tracker sheets if they don't already exist.
 * Called from CREATE_DASHBOARD() in 08a_SheetSetup.gs.
 */
function initWorkloadTrackerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Workload Vault (hidden raw data) ---
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
    vault.getRange(1, 1, 1, 24)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    vault.setFrozenRows(1);
  }
  vault.hideSheet();

  // --- Workload Reporting (anonymized, visible) ---
  var report = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
  if (!report) {
    report = ss.insertSheet(SHEETS.WORKLOAD_REPORTING);
    report.appendRow([
      'Date', 'Identity',
      'Priority Cases', 'Pending Cases', 'Unread Documents', 'To-Do Items',
      'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
      'Weekly Cases', 'Employment Type', 'Privacy', 'On Plan', 'Overtime Hours'
    ]);
    report.getRange(1, 1, 1, 15)
      .setFontWeight('bold')
      .setBackground('#0d47a1')
      .setFontColor('#ffffff');
    report.setFrozenRows(1);
    report.setHiddenGridlines(true);
  }

  // --- Workload Reminders ---
  var reminders = ss.getSheetByName(SHEETS.WORKLOAD_REMINDERS);
  if (!reminders) {
    reminders = ss.insertSheet(SHEETS.WORKLOAD_REMINDERS);
    reminders.appendRow(['Email', 'Enabled', 'Frequency', 'Day', 'Time', 'Last Sent']);
    reminders.getRange(1, 1, 1, 6)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    reminders.setFrozenRows(1);
    reminders.hideSheet();
  }

  // --- Workload UserMeta (hidden) ---
  var userMeta = ss.getSheetByName(SHEETS.WORKLOAD_USERMETA);
  if (!userMeta) {
    userMeta = ss.insertSheet(SHEETS.WORKLOAD_USERMETA);
    userMeta.appendRow(['Email', 'Sharing Start Date', 'Created Date']);
    userMeta.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    userMeta.setFrozenRows(1);
    userMeta.hideSheet();
  }

  Logger.log('Workload Tracker sheets initialized.');
}

// ============================================================================
// SHARING START DATE (reciprocity)
// ============================================================================

/**
 * @param {string} email
 * @returns {Date|null}
 */
function wtGetUserSharingStartDate_(email) {
  if (!email) return null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_USERMETA);
  if (!sheet) return null;
  var emailLower = email.toLowerCase().trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_USERMETA_COLS.EMAIL] &&
        data[i][WT_USERMETA_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
      return data[i][WT_USERMETA_COLS.SHARING_START_DATE] ?
        new Date(data[i][WT_USERMETA_COLS.SHARING_START_DATE]) : null;
    }
  }
  return null;
}

/**
 * @param {string} email
 * @param {Date} startDate
 */
function wtSetUserSharingStartDate_(email, startDate) {
  if (!email) return;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_USERMETA);
  if (!sheet) return;
  var emailLower = email.toLowerCase().trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_USERMETA_COLS.EMAIL] &&
        data[i][WT_USERMETA_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
      sheet.getRange(i + 1, WT_USERMETA_COLS.SHARING_START_DATE + 1).setValue(startDate);
      return;
    }
  }
  sheet.appendRow([emailLower, startDate, new Date()]);
}

// ============================================================================
// FORM SUBMISSION
// ============================================================================

/**
 * Processes a workload form submission from the web portal.
 * Called client-side via google.script.run.processWorkloadForm(formObj).
 *
 * @param {Object} formObj - Form fields (email, pin, t1-t8, privacy, etc.)
 * @returns {string} "Success: ..." or "Error: ..."
 */
function processWorkloadForm(formObj) {
  var email = formObj.email ? sanitizeString(formObj.email.trim().toLowerCase(), 100) : '';

  try {
    var pin = formObj.pin ? formObj.pin.toString().trim() : '';

    if (!email || !pin) {
      return 'Error: Please provide your email and PIN.';
    }

    // Validate email format
    var emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRe.test(email)) {
      return 'Error: Please enter a valid email address.';
    }

    // Submission rate limit
    var submitLimit = wtCheckSubmissionRateLimit_(email);
    if (!submitLimit.allowed) {
      return 'Error: Submission limit reached. Please wait before submitting again.';
    }

    // Authenticate using DDS member PIN
    var auth = authenticateWorkloadMember_(email, pin);
    if (!auth.success) {
      return 'Error: ' + auth.message;
    }

    // Validate numeric fields (t1–t8, 0–999)
    var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
    for (var n = 0; n < numFields.length; n++) {
      var val = Number(formObj[numFields[n]]) || 0;
      if (val < 0 || val > 999 || !Number.isInteger(val)) {
        return 'Error: Workload values must be whole numbers between 0 and 999.';
      }
    }

    // Validate part-time hours if applicable
    if (formObj.employment_type === 'Part-time') {
      var ptHours = Number(formObj.part_time_hours);
      if (isNaN(ptHours) || ptHours < 1 || ptHours > 40 || !Number.isInteger(ptHours)) {
        return 'Error: Part-time hours must be a whole number between 1 and 40.';
      }
    }

    // Use lock service for thread-safe write
    return withLock(function() {
      var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
      if (!vault) {
        return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
      }

      var sanitizeNum = function(v) { return Math.min(999, Math.max(0, Math.floor(Number(v) || 0))); };

      // Build sub-category JSON
      var subCatData = { priority: {}, pending: {}, unread: {}, todo: {},
                         referrals: {}, ce: {}, assistance: {}, aged: {} };
      var catKeys = ['priority', 'pending', 'unread', 'todo', 'referrals', 'ce', 'assistance', 'aged'];
      for (var ci = 0; ci < catKeys.length; ci++) {
        var catKey = catKeys[ci];
        var subNames = WT_SUB_CATEGORIES[catKey];
        for (var si = 0; si < subNames.length; si++) {
          var fieldName = 'sub_' + (ci + 1) + '_' + si;
          var subVal = sanitizeNum(formObj[fieldName]);
          if (subVal > 0) { subCatData[catKey][subNames[si]] = subVal; }
        }
      }

      // Weekly cases
      var weeklyCases = sanitizeNum(formObj.weekly_cases || 15);
      if (formObj.weekly_cases_option === 'manual' && formObj.weekly_cases_manual) {
        weeklyCases = sanitizeNum(formObj.weekly_cases_manual);
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
        new Date(),                                                      // Timestamp
        email,                                                           // Email
        sanitizeNum(formObj.t1),                                         // Priority Cases
        sanitizeNum(formObj.t2),                                         // Pending Cases
        sanitizeNum(formObj.t3),                                         // Unread Documents
        sanitizeNum(formObj.t4),                                         // To-Do Items
        sanitizeNum(formObj.t5),                                         // Sent Referrals
        sanitizeNum(formObj.t6),                                         // CE Activities
        sanitizeNum(formObj.t7),                                         // Assistance Requests
        sanitizeNum(formObj.t8),                                         // Aged Cases
        weeklyCases,                                                     // Weekly Cases
        JSON.stringify(subCatData),                                      // Sub-Categories JSON
        sanitizeString(formObj.employment_type || 'Full-time', 20),     // Employment Type
        formObj.employment_type === 'Part-time' ? (Number(formObj.part_time_hours) || '') : '',
        sanitizeString(formObj.leave_type || '', 20),                   // Leave Type
        sanitizeString(formObj.leave_planned || '', 20),                // Leave Planned
        leaveStart,                                                      // Leave Start
        leaveEnd,                                                        // Leave End
        sanitizeString(formObj.no_intake_choice || '', 20),             // No Intake Choice
        sanitizeString(formObj.notice_time || '', 20),                  // Notice Time
        formObj.half_day_leave ? 'Yes' : '',                            // Half Day
        sanitizeString(formObj.privacy || 'Unit', 20),                  // Privacy
        formObj.on_plan ? 'Yes' : 'No',                                 // On Plan
        formObj.overtime_enabled ? (Number(formObj.overtime_hours) || 0) : ''  // Overtime Hours
      ]);

      // Audit log via DDS's system
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('WORKLOAD_SUBMIT',
          'Member ' + auth.memberId + ' submitted workload data. Privacy: ' +
          (formObj.privacy || 'Unit'));
      }

      // Handle sharing start date for reciprocity
      var privacySetting = sanitizeString(formObj.privacy || 'Unit', 20);
      if (privacySetting !== 'Private') {
        var existingStartDate = wtGetUserSharingStartDate_(email);
        if (!existingStartDate) {
          wtSetUserSharingStartDate_(email, new Date());
        }
      }

      // Record submission for rate limiting
      wtRecordSubmission_(email);

      // Save reminder preferences
      if (formObj.reminder_enabled !== undefined) {
        wtSaveReminderPreference_({
          email: email,
          enabled: formObj.reminder_enabled === 'on' || formObj.reminder_enabled === true,
          frequency: formObj.reminder_frequency || 'weekly',
          day: formObj.reminder_day || 'monday',
          time: formObj.reminder_time || '09:00'
        });
      }

      // Refresh anonymized reporting sheet
      refreshWorkloadReportingData();

      return 'Success! Your workload data has been securely recorded.';
    }, 'processWorkloadForm');

  } catch (err) {
    console.error('processWorkloadForm error:', err);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WORKLOAD_SUBMIT_ERROR', 'Error for ' + email + ': ' + err.message);
    }
    return 'Error: Unable to process submission. Please try again.';
  }
}

// ============================================================================
// REMINDER PREFERENCES
// ============================================================================

/**
 * @param {{email,enabled,frequency,day,time}} prefs
 */
function wtSaveReminderPreference_(prefs) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_REMINDERS);
  if (!sheet) return;
  var emailLower = prefs.email.toLowerCase().trim();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][WT_REMINDERS_COLS.EMAIL] &&
        data[i][WT_REMINDERS_COLS.EMAIL].toString().toLowerCase().trim() === emailLower) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        emailLower, prefs.enabled ? 'Yes' : 'No',
        prefs.frequency, prefs.day, prefs.time
      ]]);
      return;
    }
  }
  sheet.appendRow([emailLower, prefs.enabled ? 'Yes' : 'No', prefs.frequency, prefs.day, prefs.time, '']);
}

// ============================================================================
// REPORTING REFRESH
// ============================================================================

/**
 * Rebuilds the Workload Reporting sheet with anonymized data.
 * Private users and their submissions are excluded. Identities → REDACTED.
 */
function refreshWorkloadReportingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vault   = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
  var report  = ss.getSheetByName(SHEETS.WORKLOAD_REPORTING);
  if (!vault || !report) return;

  var vaultData = vault.getDataRange().getValues();
  if (vaultData.length <= 1) return;

  // Build set of emails that have ever chosen "Private" as their MOST RECENT privacy
  var latestPrivacy = {};
  for (var i = 1; i < vaultData.length; i++) {
    var rowEmail = vaultData[i][WT_VAULT_COLS.EMAIL];
    if (rowEmail) {
      latestPrivacy[rowEmail.toString().toLowerCase().trim()] = vaultData[i][WT_VAULT_COLS.PRIVACY];
    }
  }

  // Rebuild reporting rows
  var rows = [];
  for (var j = 1; j < vaultData.length; j++) {
    var rowE = vaultData[j][WT_VAULT_COLS.EMAIL];
    if (!rowE) continue;
    var emailKey = rowE.toString().toLowerCase().trim();
    var rowPrivacy = vaultData[j][WT_VAULT_COLS.PRIVACY];
    if (rowPrivacy === 'Private') continue;              // skip private rows
    if (latestPrivacy[emailKey] === 'Private') continue; // skip if user is now private

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
      vaultData[j][WT_VAULT_COLS.WEEKLY_CASES],
      vaultData[j][WT_VAULT_COLS.EMPLOYMENT_TYPE],
      rowPrivacy,
      vaultData[j][WT_VAULT_COLS.ON_PLAN],
      vaultData[j][WT_VAULT_COLS.OVERTIME_HOURS]
    ]);
  }

  // Write back
  report.clearContents();
  report.appendRow(['Date', 'Identity', 'Priority Cases', 'Pending Cases', 'Unread Documents',
    'To-Do Items', 'Sent Referrals', 'CE Activities', 'Assistance Requests', 'Aged Cases',
    'Weekly Cases', 'Employment Type', 'Privacy', 'On Plan', 'Overtime Hours']);
  report.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#0d47a1').setFontColor('#ffffff');
  if (rows.length > 0) {
    report.getRange(2, 1, rows.length, 15).setValues(rows);
  }
}

/** Public alias for menu item */
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

  // Validate numeric fields
  var numFields = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
  for (var n = 0; n < numFields.length; n++) {
    var val = Number(formData[numFields[n]]) || 0;
    if (val < 0 || val > 999 || !Number.isInteger(val)) {
      return 'Error: Workload values must be whole numbers between 0 and 999.';
    }
  }

  return withLock(function() {
    var vault = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault) {
      return 'Error: Workload sheets not initialized. Run Setup > Initialize Dashboard first.';
    }

    var sanitizeNum = function(v) { return Math.min(999, Math.max(0, Math.floor(Number(v) || 0))); };

    var row = [
      new Date(),                               // Timestamp
      emailLower,                               // Email
      sanitizeNum(formData.t1),                 // Priority Cases
      sanitizeNum(formData.t2),                 // Pending Cases
      sanitizeNum(formData.t3),                 // Unread Docs
      sanitizeNum(formData.t4),                 // To-Do Items
      sanitizeNum(formData.t5),                 // Sent Referrals
      sanitizeNum(formData.t6),                 // CE Activities
      sanitizeNum(formData.t7),                 // Assistance Requests
      sanitizeNum(formData.t8),                 // Aged Cases
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
    return { success: true, data: { totalSubmissions: 0, members: 0, categories: {} }, message: '' };
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
