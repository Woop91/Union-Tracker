/**
 * ============================================================================
 * 00_Security.gs - Security Utilities and Access Control
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Centralizes ALL security functions for the entire application:
 *   1. escapeHtml(input) — XSS prevention by escaping HTML special characters.
 *      Called everywhere user data is rendered into HTML dialogs or the SPA.
 *   2. escapeForFormula(input) — formula injection prevention for sheet writes.
 *      Prefixes dangerous characters (=, +, -, @) that Sheets interprets as formulas.
 *   3. checkWebAppAuthorization(role) — role-based access control for the web app.
 *      Resolves the caller's role (admin/steward/member/anonymous) from the Member Directory.
 *   4. maskEmail/maskPhone/maskName — PII masking for safe logging.
 *   5. secureLog(context, message, data) — structured logging that auto-masks PII fields.
 *   6. isValidSafeString(input) — input validation that rejects <script>, javascript:, etc.
 *   7. recordSecurityEvent() — centralized security alerting (CRITICAL=immediate email,
 *      HIGH=daily digest, MEDIUM/LOW=log only).
 *   8. safeSendEmail_() — quota-aware email wrapper that prevents quota exhaustion.
 *   9. Dashboard auth toggle — admin can require/disable PIN login for dashboards.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   - Loaded first (00_ prefix alongside 00_DataAccess.gs) because security
 *     functions like escapeHtml() must be available before ANY module renders HTML.
 *   - escapeHtml() is defined as a global function (not in a namespace) because
 *     GAS HTML templates call it via scriptlets (<?= escapeHtml(data) ?>).
 *     CLAUDE.md rule: "All HTML must use escapeHtml(). No exceptions."
 *   - ACCESS_CONTROL.ENABLED is fail-secure: when set to false (e.g., admin
 *     troubleshooting), access is DENIED rather than granted. This prevents
 *     accidental data exposure if an admin disables AC to debug something.
 *   - getUserRole_() checks the spreadsheet owner first (always admin), then
 *     scans the Member Directory for steward/member status. This means the
 *     script owner always has admin access even if not in the directory.
 *   - Security events use a tiered severity system (CRITICAL/HIGH/MEDIUM/LOW)
 *     to avoid alert fatigue. Only CRITICAL sends immediate email.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   - If escapeHtml() is missing: EVERY HTML dialog and SPA view becomes
 *     vulnerable to XSS attacks. User-supplied names, notes, and grievance
 *     descriptions could execute arbitrary JavaScript.
 *   - If escapeForFormula() is missing: sheet writes could inject formulas.
 *     A member name like "=IMPORTRANGE(...)" could exfiltrate data.
 *   - If checkWebAppAuthorization() is missing: the web app (22_WebDashApp.gs)
 *     cannot verify user roles. All dashboard access would fail.
 *   - If secureLog() is missing: logging falls back to raw Logger.log() which
 *     would expose PII (emails, phones, names) in Stackdriver logs.
 *   - If safeSendEmail_() is missing: security alerts and daily digests fail
 *     to send. Admins won't be notified of security incidents.
 *
 * DEPENDENCIES:
 *   Depends on:  PropertiesService, MailApp, Session, SpreadsheetApp (all GAS built-ins)
 *                SHEETS, MEMBER_COLS, CONFIG_COLS (01_Core.gs — for role lookup)
 *                COMMAND_CONFIG (01_Core.gs — for email subjects, optional)
 *   Used by:     EVERY module that renders HTML or writes to sheets.
 *                Key callers: 02_DataManagers.gs, 04e_PublicDashboard.gs,
 *                21_WebDashDataService.gs, 22_WebDashApp.gs, all HTML templates.
 *
 * @fileoverview Security utilities — XSS, formula injection, RBAC, PII masking
 * @version 4.43.1
 */

// ============================================================================
// ACCESS CONTROL CONFIGURATION
// ============================================================================

/**
 * Access control configuration
 * @const {Object}
 */
var ACCESS_CONTROL = {
  /** Whether to enforce access control (set to true in production) */
  ENABLED: true,

  /** Allowed modes for web app */
  ALLOWED_MODES: ['steward', 'member', 'dashboard', 'search', 'grievances', 'members', 'links', 'portal', 'selfservice'],

  /** Allowed page values */
  ALLOWED_PAGES: ['dashboard', 'search', 'grievances', 'members', 'links', 'portal', 'selfservice'],

  /** Cache duration for authorization check (5 minutes) */
  AUTH_CACHE_DURATION: 300,

  /** Property key for dashboard member auth toggle */
  DASHBOARD_AUTH_PROPERTY: 'REQUIRE_MEMBER_AUTH_FOR_DASHBOARDS'
};

// ============================================================================
// DASHBOARD ACCESS CONTROL TOGGLE
// ============================================================================

/**
 * Check if member authentication is required for dashboard access
 * Default: ON (all dashboards require authentication)
 * Only returns false if explicitly set to 'false' by admin
 * @returns {boolean} True if member auth is required
 */
function isDashboardMemberAuthRequired() {
  var props = PropertiesService.getScriptProperties();
  var setting = props.getProperty(ACCESS_CONTROL.DASHBOARD_AUTH_PROPERTY);
  return setting !== 'false';
}

/**
 * Enable member authentication requirement for dashboard access
 * When enabled, all dashboard pages require member PIN login
 */
function enableDashboardMemberAuth() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Enable Dashboard Authentication',
    'This will require ALL members to log in with a PIN to access dashboards.\n\n' +
    'Make sure all members have PINs generated first. Continue?',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  var props = PropertiesService.getScriptProperties();
  props.setProperty(ACCESS_CONTROL.DASHBOARD_AUTH_PROPERTY, 'true');
  Logger.log('Dashboard member authentication ENABLED');

  if (typeof secureLog === 'function') {
    secureLog('DashboardAuthEnabled', 'Member auth required for dashboards', {});
  }

  SpreadsheetApp.getUi().alert(
    'Dashboard Authentication Enabled',
    'Members will now need to log in with their PIN to access dashboards.\n\n' +
    'Make sure all members have PINs generated before enabling this.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Disable member authentication requirement for dashboard access
 * When disabled, dashboards are accessible without member login (default)
 */
function disableDashboardMemberAuth() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(ACCESS_CONTROL.DASHBOARD_AUTH_PROPERTY, 'false');
  Logger.log('Dashboard member authentication DISABLED');

  if (typeof secureLog === 'function') {
    secureLog('DashboardAuthDisabled', 'Member auth not required for dashboards', {});
  }

  SpreadsheetApp.getUi().alert(
    'Dashboard Authentication Disabled',
    'Dashboards are now accessible without member PIN login.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Show current dashboard auth status
 */
function showDashboardAuthStatus() {
  var isEnabled = isDashboardMemberAuthRequired();
  SpreadsheetApp.getUi().alert(
    'Dashboard Authentication Status',
    'Member authentication for dashboards is currently: ' + (isEnabled ? 'ENABLED' : 'DISABLED') + '\n\n' +
    (isEnabled
      ? 'Members must log in with their PIN to access dashboards.'
      : 'Dashboards are accessible without member login.'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// XSS PREVENTION - HTML SANITIZATION
// ============================================================================

/**
 * Sanitizes a string for safe insertion into HTML content.
 * Prevents XSS attacks by escaping HTML special characters.
 *
 * @param {*} input - The input to sanitize (will be converted to string)
 * @returns {string} HTML-safe string
 *
 * @example
 * // Returns: "&lt;script&gt;alert(&#x27;xss&#x27;)&lt;/script&gt;"
 * // Note: / and = are NOT escaped (not XSS vectors in HTML text/attribute contexts)
 * escapeHtml("<script>alert('xss')</script>");
 */
function escapeHtml(input) {
  if (input === null || input === undefined) {
    return '';
  }
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/`/g, '&#x60;');
}

// ============================================================================
// FORMULA INJECTION PREVENTION
// ============================================================================

/**
 * Sanitizes input for use in Google Sheets formulas.
 * Prevents formula injection attacks.
 *
 * @param {*} input - The input to sanitize
 * @returns {string} Formula-safe string
 *
 * @example
 * // Returns: "'Member Directory'"
 * escapeForFormula("Member Directory");
 */
function escapeForFormula(input) {
  if (input === null || input === undefined) {
    return '';
  }
  var str = String(input);

  // Remove or escape characters that could be used in formula injection.
  // NOTE: GAS only interprets formulas starting with =, +, -, @, or tab at the
  // beginning of a cell value. We also guard against formula chars appearing after
  // leading whitespace, since some spreadsheet engines trim before evaluating.
  return str
    .replace(/'/g, "''")       // Escape single quotes
    .replace(/"/g, '""')       // Escape double quotes
    .replace(/\\/g, '\\\\')    // Escape backslashes
    .replace(/[\r\n]/g, ' ')   // Replace newlines with spaces
    .replace(/\t/g, ' ')       // Replace tab characters (tab can trigger formula interpretation)
    .replace(/^(\s*)[=+\-@]/, function(match, leadingSpace) {
      // Prefix formula-starting characters with a single quote, even after leading whitespace
      return "'" + match;
    });
}

/**
 * Safely creates a sheet name reference for use in formulas
 * @param {string} sheetName - The sheet name to escape
 * @returns {string} Safe sheet name for formula use
 */
function safeSheetNameForFormula(sheetName) {
  if (!sheetName) return '';

  var escaped = String(sheetName)
    .replace(/'/g, "''");  // Escape single quotes

  // Wrap in quotes if contains special characters
  if (/[^a-zA-Z0-9_]/.test(escaped)) {
    return "'" + escaped + "'";
  }
  return escaped;
}
// ============================================================================
// ACCESS CONTROL FOR WEB APP
// ============================================================================

/**
 * Checks if the caller is authorized for the given role.
 *
 * In "Execute as: Me" web apps, Session.getActiveUser() returns empty for
 * magic link / session token users. Pass sessionToken to cover that path.
 *
 * @param {string} requiredRole - 'steward', 'admin', or any role
 * @param {string=} sessionToken - Optional session token to verify identity when SSO unavailable
 * @returns {Object} { isAuthorized, user, email, role, message }
 */
function checkWebAppAuthorization(requiredRole, sessionToken) {
  var result = {
    isAuthorized: false,
    user: null,
    email: null,
    role: 'anonymous',
    message: ''
  };

  try {
    // Use getActiveUser() — NOT getEffectiveUser(). In "Execute as me" web apps,
    // getEffectiveUser() returns the script owner, not the actual accessing user.
    // getActiveUser() returns the real user who is making the request.
    var user = Session.getActiveUser();
    var email = user ? user.getEmail() : null;

    // Fallback for magic link / session token users (getActiveUser() is empty in Execute-as-Me)
    if (!email && sessionToken && typeof Auth !== 'undefined' && typeof Auth.resolveEmailFromToken === 'function') {
      email = Auth.resolveEmailFromToken(sessionToken) || null;
      user = null; // no GAS user object available for token-based auth
    }

    if (!email) {
      result.message = 'Authentication required';
      return result;
    }

    result.user = user;
    result.email = email;

    // Check if access control is enabled
    // v4.25.5 SEC-01: Fail-secure — when access control is disabled, DENY by default.
    // Previously this auto-authorized non-privileged users, exposing member data
    // if an admin disabled AC to troubleshoot. Now disabled = locked down.
    if (!ACCESS_CONTROL.ENABLED) {
      Logger.log('ACCESS_CONTROL is disabled — denying access (fail-secure). Re-enable to restore normal operation.');
      result.message = 'Access control is currently disabled. Contact your administrator to re-enable it.';
      return result;
    }

    // Determine user role
    var role = getUserRole_(email);
    result.role = role;

    // Check role requirement
    if (requiredRole) {
      if (requiredRole === 'admin' && role !== 'admin') {
        result.message = 'Administrator access required';
        return result;
      }
      if (requiredRole === 'steward' && role !== 'steward' && role !== 'admin' && role !== 'both') {
        result.message = 'Steward access required';
        return result;
      }
    }

    result.isAuthorized = true;
    return result;

  } catch (e) {
    result.message = 'Authorization check failed: ' + e.message;
    Logger.log('Authorization error: ' + e.message);
    return result;
  }
}

/**
 * Gets the role of a user based on their email
 * @param {string} email - User email
 * @returns {string} Role ('admin', 'steward', 'member', 'anonymous')
 * @private
 */
function getUserRole_(email) {
  if (!email) return 'anonymous';

  // Check cache first to avoid reading entire Member Directory on every call
  var cacheKey = 'user_role_' + email.toLowerCase();
  var cached = CacheService.getScriptCache().get(cacheKey);
  if (cached) return cached;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return 'anonymous'; // Web app context — spreadsheet binding unavailable

    // Check if user is the spreadsheet owner (admin)
    var role;
    var owner = ss.getOwner();
    if (owner && owner.getEmail().toLowerCase() === email.toLowerCase()) {
      role = 'admin';
      try { CacheService.getScriptCache().put(cacheKey, role, ACCESS_CONTROL.AUTH_CACHE_DURATION || 300); } catch(_) {}
      return role;
    }

    // Check if user is in the stewards list
    if (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) {
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberSheet) {
        var data = memberSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          var memberEmail = data[i][MEMBER_COLS.EMAIL - 1] || '';
          if (memberEmail.toLowerCase() === email.toLowerCase()) {
            var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
            var roleCol = MEMBER_COLS.ROLE ? data[i][MEMBER_COLS.ROLE - 1] : '';
            var roleRaw = String(roleCol || '').trim().toLowerCase();
            var isBoth = roleRaw === 'both' || roleRaw === 'steward/member';
            if (isBoth) {
              role = 'both';
            } else if (isTruthyValue(isSteward)) {
              role = 'steward';
            } else {
              role = 'member';
            }
            try { CacheService.getScriptCache().put(cacheKey, role, ACCESS_CONTROL.AUTH_CACHE_DURATION || 300); } catch(_) {}
            return role;
          }
        }
      }
    }

    return 'anonymous';
  } catch (e) {
    Logger.log('Error getting user role: ' + e.message);
    return 'anonymous';
  }
}

// ============================================================================
// PII MASKING FOR LOGS
// ============================================================================

/**
 * Masks an email address for logging
 * @param {string} email - Email to mask
 * @returns {string} Masked email (e.g., "j***@example.com")
 */
function maskEmail(email) {
  if (!email || typeof email !== 'string') return '[no email]';

  var atIndex = email.indexOf('@');
  if (atIndex <= 1) return '***@***';

  var localPart = email.substring(0, atIndex);
  var domain = email.substring(atIndex);

  if (localPart.length <= 2) {
    return localPart.charAt(0) + '***' + domain;
  }

  return localPart.charAt(0) + '***' + localPart.charAt(localPart.length - 1) + domain;
}

/**
 * Masks a phone number for logging
 * @param {string} phone - Phone number to mask
 * @returns {string} Masked phone (e.g., "***-***-1234")
 */
function maskPhone(phone) {
  if (!phone || typeof phone !== 'string') return '[no phone]';

  var digits = phone.replace(/\D/g, '');
  if (digits.length < 4) return '***';

  return '***-***-' + digits.slice(-4);
}

/**
 * Masks a name for logging
 * @param {string} firstName - First name
 * @param {string} lastName - Last name
 * @returns {string} Masked name (e.g., "J. S.")
 */
function maskName(firstName, lastName) {
  var first = firstName ? String(firstName).charAt(0) + '.' : '';
  var last = lastName ? String(lastName).charAt(0) + '.' : '';
  return (first + ' ' + last).trim() || '[anonymous]';
}

// ============================================================================
// SECURE LOGGING
// ============================================================================

/**
 * Masks PII fields in an arbitrary key-value object.
 * Shared utility used by secureLog, sendSecurityAlertEmail_, and queueSecurityDigestEvent_.
 * @param {Object} obj - Object with potential PII fields
 * @param {boolean} [skipInternal=false] - If true, skip keys starting with '_'
 * @returns {Object} New object with PII fields masked
 */
function maskObjectPII_(obj, skipInternal) {
  if (!obj || typeof obj !== 'object') return {};
  var masked = {};
  for (var key in obj) {
    if (!obj.hasOwnProperty(key)) continue;
    if (skipInternal && key.charAt(0) === '_') continue;
    var val = obj[key];
    var keyLower = key.toLowerCase();
    if (keyLower.indexOf('email') !== -1 && typeof val === 'string') {
      masked[key] = maskEmail(val);
    } else if (keyLower.indexOf('phone') !== -1 && typeof val === 'string') {
      masked[key] = maskPhone(val);
    } else if ((key === 'firstName' || key === 'lastName' || key === 'name') && (typeof val === 'string' || !val)) {
      masked[key] = val ? String(val).charAt(0) + '.' : '';
    } else if ((keyLower.indexOf('address') !== -1 || keyLower.indexOf('ssn') !== -1 ||
                keyLower.indexOf('social') !== -1 || keyLower.indexOf('dob') !== -1 ||
                keyLower.indexOf('birthdate') !== -1) && typeof val === 'string') {
      masked[key] = '[REDACTED]';
    } else {
      masked[key] = val;
    }
  }
  return masked;
}

/**
 * Logs a message without exposing PII
 * @param {string} context - The context/function name
 * @param {string} message - The log message
 * @param {Object} [data] - Optional data (will be masked)
 */
function secureLog(context, message, data) {
  var logMessage = '[' + context + '] ' + message;

  if (data) {
    logMessage += ' | Data: ' + JSON.stringify(maskObjectPII_(data));
  }

  Logger.log(logMessage);
}

// ============================================================================
// INPUT VALIDATION
// ============================================================================

/**
 * Validates that a value is a safe string (no script tags, etc.)
 * @param {*} input - Input to validate
 * @param {number} [maxLength=1000] - Maximum allowed length
 * @returns {boolean} True if valid
 */
function isValidSafeString(input, maxLength) {
  // null/undefined are not valid safe strings — callers should handle them explicitly
  if (input === null || input === undefined) return false;
  if (typeof input !== 'string') return false;

  maxLength = maxLength || 1000;
  if (input.length > maxLength) return false;

  // Check for dangerous patterns
  var dangerousPatterns = [
    /<script/i,
    /javascript:/i,
    /on\w+\s*=/i,  // onclick=, onerror=, etc.
    /^\s*data:/i,
    /vbscript:/i
  ];

  for (var i = 0; i < dangerousPatterns.length; i++) {
    if (dangerousPatterns[i].test(input)) {
      return false;
    }
  }

  return true;
}
/**
 * Validates a grievance ID format
 * @param {string} grievanceId - Grievance ID to validate
 * @returns {boolean} True if valid format
 */
function isValidGrievanceId(grievanceId) {
  if (!grievanceId) return false;
  // Allow alphanumeric, hyphen, underscore, max 50 chars
  return /^[a-zA-Z0-9\-_]{1,50}$/.test(String(grievanceId));
}

// ============================================================================
// CLIENT-SIDE SECURITY HELPERS
// ============================================================================

/**
 * Returns JavaScript code for client-side HTML escaping.
 * Include this in your HTML templates inside a <script> tag.
 *
 * @returns {string} JavaScript code defining escapeHtml function
 *
 * @example
 * var html = '<script>' + getClientSideEscapeHtml() + '</script>';
 */
function getClientSideEscapeHtml() {
  return 'function escapeHtml(t){' +
    'if(t==null)return"";' +
    'return String(t)' +
    '.replace(/&/g,"&amp;")' +
    '.replace(/</g,"&lt;")' +
    '.replace(/>/g,"&gt;")' +
    '.replace(/"/g,"&quot;")' +
    '.replace(/\'/g,"&#x27;")' +
    '.replace(/`/g,"&#x60;");' +
    '}' +
    'function safeText(t){return escapeHtml(t);}';
}
// ============================================================================
// SECURITY EVENT ALERTING SYSTEM
// ============================================================================

/**
 * Security event severity levels.
 * CRITICAL events trigger immediate email alerts.
 * HIGH events are emailed in the daily digest.
 * MEDIUM/LOW events are logged only.
 * @const {Object}
 */
var SECURITY_SEVERITY = {
  CRITICAL: 'CRITICAL',
  HIGH: 'HIGH',
  MEDIUM: 'MEDIUM',
  LOW: 'LOW'
};

/**
 * Records a security event and triggers appropriate notifications.
 *
 * This is the centralized entry point for all security-relevant events.
 * Events are always written to the audit log. CRITICAL events additionally
 * send an immediate email alert to the Chief Steward. HIGH events are
 * batched and included in the next daily security digest.
 *
 * @param {string} eventType - Short event identifier (e.g. 'UNAUTHORIZED_ACCESS')
 * @param {string} severity - One of SECURITY_SEVERITY values
 * @param {string} description - Human-readable description
 * @param {Object} [details] - Additional context (will be PII-masked in logs)
 */
function recordSecurityEvent(eventType, severity, description, details) {
  details = details || {};
  details._severity = severity;
  details._description = description;

  // Always log to audit
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('SECURITY_' + eventType, details);
  }
  secureLog('Security', '[' + severity + '] ' + eventType + ': ' + description, details);

  // CRITICAL events get immediate email to Chief Steward and Admin
  if (severity === SECURITY_SEVERITY.CRITICAL) {
    sendSecurityAlertEmail_(eventType, description, details);
  }

  // HIGH events are batched for the daily digest
  if (severity === SECURITY_SEVERITY.HIGH) {
    queueSecurityDigestEvent_(eventType, description, details);
  }
}

/**
 * Sends an immediate security alert email to the Chief Steward and Admin.
 * Used for CRITICAL severity events only.
 *
 * @param {string} eventType - Event identifier
 * @param {string} description - Human-readable description
 * @param {Object} details - Event context
 * @private
 */
function sendSecurityAlertEmail_(eventType, description, details) {
  // L5: Known limitation — config reads (getConfigValue_ for CHIEF_STEWARD_EMAIL /
  // ADMIN_EMAILS) below are not protected by a LockService lock. In theory a concurrent
  // config write could cause a stale or partial read. This is accepted because:
  //   1. Config changes are extremely rare (admin-only, manual).
  //   2. Adding a lock here risks delaying time-sensitive security alerts.
  //   3. Worst case is an alert sent to a slightly stale recipient list.
  try {
    // Gather recipient emails from Config
    var recipients = [];

    if (typeof getConfigValue_ === 'function') {
      var chiefEmail = '';
      var adminEmails = '';
      try {
        chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
        adminEmails = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

      if (chiefEmail) recipients.push(chiefEmail);
      if (adminEmails) {
        adminEmails.split(',').forEach(function(e) {
          var trimmed = e.trim();
          if (trimmed && recipients.indexOf(trimmed) === -1) recipients.push(trimmed);
        });
      }
    }

    // Fall back to script owner
    if (recipients.length === 0) {
      try {
        var ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
        if (ownerEmail) recipients.push(ownerEmail);
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }

    // FIX-SEC-01 (cont): Quota check removed — safeSendEmail_() handles this internally.
    if (recipients.length === 0) return;

    // Mask PII in details before including in email
    var safeDetails = maskObjectPII_(details, true);

    var systemName = 'Union Dashboard';
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.SYSTEM_NAME) {
      systemName = COMMAND_CONFIG.SYSTEM_NAME;
    }

    var subjectPrefix = '[' + systemName + ']';
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.EMAIL && COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX) {
      subjectPrefix = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX;
    }

    var footer = '\n\n---\nGenerated by ' + systemName;
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.EMAIL && COMMAND_CONFIG.EMAIL.FOOTER) {
      footer = COMMAND_CONFIG.EMAIL.FOOTER;
    }

    var body = 'SECURITY ALERT: ' + eventType + '\n' +
      'Severity: ' + (details._severity || 'CRITICAL') + '\n' +
      'Time: ' + new Date().toLocaleString() + '\n\n' +
      'Description:\n' + description + '\n\n' +
      'Details:\n' + JSON.stringify(safeDetails, null, 2) + '\n\n' +
      'Action Required: Please review this event immediately.' +
      footer;

    // FIX-SEC-01: v4.25.8 — Use safeSendEmail_() for quota guard + format validation.
    // Removes redundant getRemainingDailyQuota() check above (handled by safeSendEmail_).
    safeSendEmail_({
      to: recipients.join(','),
      subject: subjectPrefix + ' SECURITY ALERT: ' + eventType,
      body: body
    });

  } catch (e) {
    Logger.log('Failed to send security alert email: ' + e.message);
  }
}

/**
 * Queues a HIGH-severity event for the daily security digest email.
 * Events are stored in Script Properties as a JSON array.
 *
 * @param {string} eventType - Event identifier
 * @param {string} description - Description
 * @param {Object} details - Context
 * @private
 */
function queueSecurityDigestEvent_(eventType, description, details) {
  try {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return;
    try {
      var props = PropertiesService.getScriptProperties();
      var existing = props.getProperty('SECURITY_DIGEST_QUEUE') || '[]';
      var queue = JSON.parse(existing);

      // Mask PII before storing
      var safeDetails = maskObjectPII_(details, true);

      queue.push({
        time: new Date().toISOString(),
        event: eventType,
        description: description,
        details: safeDetails
      });

      // Cap at 30 events to prevent property size overflow (9KB limit)
      if (queue.length > 30) {
        queue = queue.slice(-30);
      }

      // Size guard: ensure serialized queue stays under 9KB property limit
      var jsonStr = JSON.stringify(queue);
      if (jsonStr.length > 8500) {
        queue = queue.slice(-15);
        jsonStr = JSON.stringify(queue);
      }

      props.setProperty('SECURITY_DIGEST_QUEUE', jsonStr);
    } finally { lock.releaseLock(); }
  } catch (e) {
    Logger.log('Failed to queue security digest event: ' + e.message);
  }
}

/**
 * Sends the daily security digest email and clears the queue.
 * Should be called by a daily time-driven trigger.
 * If no events are queued, no email is sent.
 */
function sendDailySecurityDigest() {
  try {
    var props = PropertiesService.getScriptProperties();
    var existing = props.getProperty('SECURITY_DIGEST_QUEUE') || '[]';
    var queue = JSON.parse(existing);

    if (queue.length === 0) return; // Nothing to report

    // Also run integrity checks
    var auditIntegrity = null;
    var vaultIntegrity = null;
    if (typeof verifyAuditLogIntegrity === 'function') {
      try { auditIntegrity = verifyAuditLogIntegrity(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }
    if (typeof verifySurveyVaultIntegrity === 'function') {
      try { vaultIntegrity = verifySurveyVaultIntegrity(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }

    // Gather recipients
    var recipients = [];
    if (typeof getConfigValue_ === 'function' && typeof CONFIG_COLS !== 'undefined') {
      try {
        var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
        var adminEmails = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
        if (chiefEmail) recipients.push(chiefEmail);
        if (adminEmails) {
          adminEmails.split(',').forEach(function(e) {
            var trimmed = e.trim();
            if (trimmed && recipients.indexOf(trimmed) === -1) recipients.push(trimmed);
          });
        }
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }
    if (recipients.length === 0) {
      try {
        var ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
        if (ownerEmail) recipients.push(ownerEmail);
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }

    if (recipients.length === 0) return;

    var systemName = 'Union Dashboard';
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.SYSTEM_NAME) {
      systemName = COMMAND_CONFIG.SYSTEM_NAME;
    }

    var subjectPrefix = '[' + systemName + ']';
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.EMAIL && COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX) {
      subjectPrefix = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX;
    }

    // Build digest body
    var body = 'DAILY SECURITY DIGEST\n' +
      'System: ' + systemName + '\n' +
      'Period: Last 24 hours\n' +
      'Date: ' + new Date().toLocaleDateString() + '\n' +
      'Events: ' + queue.length + '\n\n';

    body += '═══════════════════════════════════════\n';
    body += 'SECURITY EVENTS (' + queue.length + ')\n';
    body += '═══════════════════════════════════════\n\n';

    for (var i = 0; i < queue.length; i++) {
      var evt = queue[i];
      body += (i + 1) + '. [' + evt.event + '] ' + evt.description + '\n';
      body += '   Time: ' + evt.time + '\n';
      if (evt.details && Object.keys(evt.details).length > 0) {
        body += '   Details: ' + JSON.stringify(evt.details) + '\n';
      }
      body += '\n';
    }

    // Integrity check results
    body += '═══════════════════════════════════════\n';
    body += 'INTEGRITY STATUS\n';
    body += '═══════════════════════════════════════\n\n';

    if (auditIntegrity) {
      body += 'Audit Log: ' + (auditIntegrity.valid ? 'PASS' : 'FAIL') +
        ' (' + auditIntegrity.totalRows + ' entries)';
      if (!auditIntegrity.valid) {
        body += ' — ' + auditIntegrity.invalidRows.length + ' tampered rows detected!';
      }
      if (auditIntegrity.message) {
        body += '\n   Note: ' + auditIntegrity.message;
      }
      body += '\n';
    }

    if (vaultIntegrity) {
      body += 'Survey Vault: ' + (vaultIntegrity.valid ? 'PASS' : 'FAIL') +
        ' (' + vaultIntegrity.stats.totalEntries + ' entries)';
      if (!vaultIntegrity.valid) {
        body += ' — ' + vaultIntegrity.issues.length + ' issue(s) found';
      }
      body += '\n';
    }

    var footer = '\n---\nGenerated by ' + systemName;
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.EMAIL && COMMAND_CONFIG.EMAIL.FOOTER) {
      footer = COMMAND_CONFIG.EMAIL.FOOTER;
    }
    body += footer;

    safeSendEmail_({
      to: recipients.join(','),
      subject: subjectPrefix + ' Daily Security Digest — ' + queue.length + ' event(s)',
      body: body
    });

    // Clear the queue only after successful email send
    props.setProperty('SECURITY_DIGEST_QUEUE', '[]');

    Logger.log('Security digest sent with ' + queue.length + ' events');

  } catch (e) {
    // Events remain in the queue so the next digest run can retry
    Logger.log('Failed to send security digest: ' + e.message);
  }
}

// ============================================================================
// SAFE EMAIL WRAPPER
// ============================================================================

/**
 * Sends an email with quota check and basic validation.
 * Drop-in replacement for MailApp.sendEmail() that prevents quota exhaustion.
 *
 * @param {Object} options - MailApp.sendEmail() options (to, subject, body, etc.)
 * @returns {{ success: boolean, error?: string }}
 * @private
 */
function safeSendEmail_(options) {
  if (!options || !options.to || !options.subject) {
    return { success: false, error: 'Missing required email fields (to, subject)' };
  }

  // Validate email format
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var recipients = String(options.to).split(',');
  for (var i = 0; i < recipients.length; i++) {
    if (!emailRegex.test(recipients[i].trim())) {
      return { success: false, error: 'Invalid email address: ' + maskEmail(recipients[i].trim()) };
    }
  }

  // Check quota before sending
  var remaining = MailApp.getRemainingDailyQuota();
  if (remaining < 1) {
    secureLog('safeSendEmail_', 'Email quota exhausted, skipping send', { to: maskEmail(String(options.to)) });
    return { success: false, error: 'Daily email quota exhausted' };
  }

  try {
    MailApp.sendEmail(options);
    return { success: true };
  } catch (e) {
    secureLog('safeSendEmail_', 'Email send failed: ' + e.message, { to: maskEmail(String(options.to)) });
    return { success: false, error: e.message };
  }
}
