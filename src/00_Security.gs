/**
 * ============================================================================
 * 00_Security.gs - Security Utilities and Access Control
 * ============================================================================
 *
 * This module provides centralized security functions including:
 * - XSS prevention (HTML sanitization)
 * - Formula injection prevention
 * - Access control for web apps
 * - Input validation and sanitization
 * - PII masking for logs
 *
 * MUST be loaded first in build order (00_ prefix).
 *
 * @fileoverview Security utilities for the dashboard
 * @version 1.0.0
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
 * Default: OFF (dashboards accessible without member login)
 * @returns {boolean} True if member auth is required
 */
function isDashboardMemberAuthRequired() {
  var props = PropertiesService.getScriptProperties();
  var setting = props.getProperty(ACCESS_CONTROL.DASHBOARD_AUTH_PROPERTY);
  return setting === 'true';
}

/**
 * Enable member authentication requirement for dashboard access
 * When enabled, all dashboard pages require member PIN login
 */
function enableDashboardMemberAuth() {
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
    .replace(/\//g, '&#x2F;')
    .replace(/`/g, '&#x60;')
    .replace(/=/g, '&#x3D;');
}

/**
 * Alias for escapeHtml - use in innerHTML contexts
 * @param {*} input - The input to sanitize
 * @returns {string} HTML-safe string
 */
function sanitizeForHtml(input) {
  return escapeHtml(input);
}

/**
 * Sanitizes an object's string values for HTML output
 * @param {Object} obj - Object with string values to sanitize
 * @returns {Object} New object with sanitized values
 */
function sanitizeObjectForHtml(obj) {
  if (!obj || typeof obj !== 'object') {
    return obj;
  }

  var result = {};
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      var value = obj[key];
      if (typeof value === 'string') {
        result[key] = escapeHtml(value);
      } else if (typeof value === 'object' && value !== null) {
        result[key] = sanitizeObjectForHtml(value);
      } else {
        result[key] = value;
      }
    }
  }
  return result;
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

  // Remove or escape characters that could be used in formula injection
  return str
    .replace(/'/g, "''")       // Escape single quotes
    .replace(/"/g, '""')       // Escape double quotes
    .replace(/\\/g, '\\\\')    // Escape backslashes
    .replace(/[\r\n]/g, ' ')   // Replace newlines with spaces
    .replace(/^[=+\-@]/, function(match) {
      // Prefix formula-starting characters with a single quote only at start of string
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

/**
 * Safely creates a QUERY formula with sanitized parameters
 * @param {string} sheetName - The sheet name
 * @param {string} query - The query string
 * @param {number} headers - Number of header rows
 * @returns {string} Safe QUERY formula
 */
function buildSafeQuery(sheetName, query, headers) {
  var safeSheet = safeSheetNameForFormula(sheetName);
  var safeHeaders = parseInt(headers, 10) || 1;

  // Sanitize the query - remove potentially dangerous characters
  var safeQuery = String(query)
    .replace(/'/g, "''")
    .replace(/"/g, '\\"');

  return '=QUERY(' + safeSheet + '!A:Z, "' + safeQuery + '", ' + safeHeaders + ')';
}

// ============================================================================
// CENTRALIZED USER SESSION
// ============================================================================

/**
 * Gets the current user's email with caching and error handling.
 * Use this instead of calling Session.getActiveUser().getEmail() directly.
 * @returns {string} The user's email or 'Unknown' if unavailable
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || 'Unknown';
  } catch (_e) {
    return 'Unknown';
  }
}

// ============================================================================
// ACCESS CONTROL FOR WEB APP
// ============================================================================

/**
 * Checks if a user is authorized to access the web app
 * @param {string} [requiredRole] - Optional role requirement ('steward', 'admin')
 * @returns {Object} Authorization result with isAuthorized and user info
 */
function checkWebAppAuthorization(requiredRole) {
  var result = {
    isAuthorized: false,
    user: null,
    email: null,
    role: 'anonymous',
    message: ''
  };

  try {
    // Get the effective user (the user accessing the web app)
    var user = Session.getEffectiveUser();
    var email = user ? user.getEmail() : null;

    if (!email) {
      result.message = 'Authentication required';
      return result;
    }

    result.user = user;
    result.email = email;

    // Check if access control is enabled
    // v4.5.2: Even when access control is disabled, ALWAYS enforce role checks
    // for privileged roles (steward/admin) to protect PII access.
    if (!ACCESS_CONTROL.ENABLED && requiredRole !== 'steward' && requiredRole !== 'admin') {
      result.isAuthorized = true;
      result.role = 'user';
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
      if (requiredRole === 'steward' && role !== 'steward' && role !== 'admin') {
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

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check if user is the spreadsheet owner (admin)
    var owner = ss.getOwner();
    if (owner && owner.getEmail().toLowerCase() === email.toLowerCase()) {
      return 'admin';
    }

    // Check if user is in the stewards list
    if (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) {
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberSheet) {
        var data = memberSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          var memberEmail = data[i][MEMBER_COLUMNS.EMAIL] || '';
          if (memberEmail.toLowerCase() === email.toLowerCase()) {
            var isSteward = data[i][MEMBER_COLUMNS.IS_STEWARD];
            if (isTruthyValue(isSteward)) {
              return 'steward';
            }
            return 'member';
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

/**
 * Validates web app request parameters
 * @param {Object} e - The event object from doGet
 * @returns {Object} Validation result with isValid and sanitized parameters
 */
function validateWebAppRequest(e) {
  var result = {
    isValid: true,
    params: {},
    errors: []
  };

  if (!e || !e.parameter) {
    return result;  // No parameters is valid
  }

  // Validate and sanitize 'mode' parameter
  if (e.parameter.mode) {
    var mode = String(e.parameter.mode).toLowerCase();
    if (ACCESS_CONTROL.ALLOWED_MODES.indexOf(mode) === -1) {
      result.isValid = false;
      result.errors.push('Invalid mode parameter');
    } else {
      result.params.mode = mode;
    }
  }

  // Validate and sanitize 'page' parameter
  if (e.parameter.page) {
    var page = String(e.parameter.page).toLowerCase();
    if (ACCESS_CONTROL.ALLOWED_PAGES.indexOf(page) === -1) {
      result.isValid = false;
      result.errors.push('Invalid page parameter');
    } else {
      result.params.page = page;
    }
  }

  // Validate and sanitize 'id' parameter (member ID)
  if (e.parameter.id) {
    var id = String(e.parameter.id);
    // Only allow alphanumeric and hyphen for IDs
    if (!/^[a-zA-Z0-9\-_]+$/.test(id)) {
      result.isValid = false;
      result.errors.push('Invalid ID format');
    } else {
      result.params.id = id;
    }
  }

  // Validate and sanitize 'filter' parameter
  if (e.parameter.filter) {
    var filter = String(e.parameter.filter).toLowerCase();
    var allowedFilters = ['open', 'closed', 'overdue', 'pending', 'all'];
    if (allowedFilters.indexOf(filter) === -1) {
      result.isValid = false;
      result.errors.push('Invalid filter parameter');
    } else {
      result.params.filter = filter;
    }
  }

  return result;
}

/**
 * Returns an access denied HTML page
 * @param {string} [message] - Custom message to display
 * @returns {HtmlOutput} Access denied page
 */
function getAccessDeniedPage(message) {
  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>Access Denied</title>' +
    '<style>' +
    'body{font-family:-apple-system,sans-serif;background:#f5f5f5;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}' +
    '.container{background:white;padding:40px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,0.1);text-align:center;max-width:400px}' +
    '.icon{font-size:64px;margin-bottom:20px}' +
    'h1{color:#DC2626;margin:0 0 15px}' +
    'p{color:#666;margin:0}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<div class="icon">🔒</div>' +
    '<h1>Access Denied</h1>' +
    '<p>' + escapeHtml(message || 'You do not have permission to access this resource.') + '</p>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Access Denied')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
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

/**
 * Creates a log-safe version of a member object (masks PII)
 * @param {Object} member - Member object with PII
 * @returns {Object} Member object with masked PII
 */
function maskMemberForLog(member) {
  if (!member) return null;

  return {
    id: member.memberId || member.id || '[no id]',
    name: maskName(member.firstName, member.lastName),
    email: maskEmail(member.email),
    phone: maskPhone(member.phone),
    unit: member.unit || '[no unit]',
    location: member.location || member.workLocation || '[no location]'
  };
}

/**
 * Creates a log-safe version of a grievance object (masks PII)
 * @param {Object} grievance - Grievance object with PII
 * @returns {Object} Grievance object with masked PII
 */
function maskGrievanceForLog(grievance) {
  if (!grievance) return null;

  return {
    id: grievance.grievanceId || grievance.id || '[no id]',
    memberName: maskName(grievance.firstName, grievance.lastName),
    status: grievance.status || '[no status]',
    step: grievance.currentStep || grievance.step || '[no step]',
    category: grievance.issueCategory || grievance.category || '[no category]'
  };
}

// ============================================================================
// SECURE LOGGING
// ============================================================================

/**
 * Logs a message without exposing PII
 * @param {string} context - The context/function name
 * @param {string} message - The log message
 * @param {Object} [data] - Optional data (will be masked)
 */
function secureLog(context, message, data) {
  var logMessage = '[' + context + '] ' + message;

  if (data) {
    // Create a masked version of data for logging
    var maskedData = {};
    for (var key in data) {
      if (data.hasOwnProperty(key)) {
        var value = data[key];
        // Check for PII fields and mask them
        if (key.toLowerCase().indexOf('email') !== -1) {
          maskedData[key] = maskEmail(value);
        } else if (key.toLowerCase().indexOf('phone') !== -1) {
          maskedData[key] = maskPhone(value);
        } else if (key === 'firstName' || key === 'lastName' || key === 'name') {
          maskedData[key] = value ? String(value).charAt(0) + '.' : '';
        } else {
          maskedData[key] = value;
        }
      }
    }
    logMessage += ' | Data: ' + JSON.stringify(maskedData);
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
  if (input === null || input === undefined) return true;
  if (typeof input !== 'string') return false;

  maxLength = maxLength || 1000;
  if (input.length > maxLength) return false;

  // Check for dangerous patterns
  var dangerousPatterns = [
    /<script/i,
    /javascript:/i,
    /on\w+\s*=/i,  // onclick=, onerror=, etc.
    /data:/i,
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
 * Validates a member ID format
 * @param {string} memberId - Member ID to validate
 * @returns {boolean} True if valid format
 */
function isValidMemberId(memberId) {
  if (!memberId) return false;
  // Allow alphanumeric, hyphen, underscore, max 50 chars
  return /^[a-zA-Z0-9\-_]{1,50}$/.test(String(memberId));
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
    '.replace(/\\//g,"&#x2F;")' +
    '.replace(/`/g,"&#x60;")' +
    '.replace(/=/g,"&#x3D;");' +
    '}' +
    'function safeText(t){return escapeHtml(t);}';
}

/**
 * Returns the full client-side security script as a <script> tag.
 * Include this at the start of your HTML body.
 *
 * @returns {string} Full script tag with security functions
 */
function getClientSecurityScript() {
  return '<script>' +
    getClientSideEscapeHtml() +
    "function safeAttr(t){return escapeHtml(t);}" +
    '</script>';
}

/**
 * Sanitizes data for embedding in JSON within HTML.
 * Use this when passing server data to client-side JavaScript.
 *
 * @param {Object} data - Data to sanitize
 * @returns {string} JSON string safe for HTML embedding
 */
function safeJsonForHtml(data) {
  if (!data) return '{}';

  // Convert to JSON and escape HTML entities in strings
  var json = JSON.stringify(data, function(key, value) {
    if (typeof value === 'string') {
      return escapeHtml(value);
    }
    return value;
  });

  // Escape </script> tags that could break out of script context
  return json.replace(/<\/script>/gi, '<\\/script>');
}

/**
 * Pre-sanitizes an array of objects for safe client-side rendering.
 * Call this on the server before passing data to client.
 *
 * @param {Array<Object>} dataArray - Array of data objects
 * @param {Array<string>} fieldsToSanitize - Field names to sanitize
 * @returns {Array<Object>} Array with sanitized string fields
 */
function sanitizeDataForClient(dataArray, fieldsToSanitize) {
  if (!Array.isArray(dataArray)) return [];
  if (!Array.isArray(fieldsToSanitize) || fieldsToSanitize.length === 0) {
    // Sanitize all string fields
    return dataArray.map(function(item) {
      return sanitizeObjectForHtml(item);
    });
  }

  return dataArray.map(function(item) {
    var sanitized = {};
    for (var key in item) {
      if (item.hasOwnProperty(key)) {
        if (fieldsToSanitize.indexOf(key) !== -1 && typeof item[key] === 'string') {
          sanitized[key] = escapeHtml(item[key]);
        } else {
          sanitized[key] = item[key];
        }
      }
    }
    return sanitized;
  });
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
  try {
    // Gather recipient emails from Config
    var recipients = [];

    if (typeof getConfigValue_ === 'function') {
      var chiefEmail = '';
      var adminEmails = '';
      try {
        chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);
        adminEmails = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
      } catch (_e) { /* Config not available */ }

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
      } catch (_e) { /* Can't get owner */ }
    }

    if (recipients.length === 0 || MailApp.getRemainingDailyQuota() < 1) return;

    // Mask PII in details before including in email
    var safeDetails = {};
    for (var key in details) {
      if (details.hasOwnProperty(key) && key.charAt(0) !== '_') {
        var val = details[key];
        if (key.toLowerCase().indexOf('email') !== -1 && typeof val === 'string') {
          safeDetails[key] = maskEmail(val);
        } else {
          safeDetails[key] = val;
        }
      }
    }

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

    MailApp.sendEmail({
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
    var props = PropertiesService.getScriptProperties();
    var existing = props.getProperty('SECURITY_DIGEST_QUEUE') || '[]';
    var queue = JSON.parse(existing);

    // Mask PII before storing
    var safeDetails = {};
    for (var key in details) {
      if (details.hasOwnProperty(key) && key.charAt(0) !== '_') {
        var val = details[key];
        if (key.toLowerCase().indexOf('email') !== -1 && typeof val === 'string') {
          safeDetails[key] = maskEmail(val);
        } else {
          safeDetails[key] = val;
        }
      }
    }

    queue.push({
      time: new Date().toISOString(),
      event: eventType,
      description: description,
      details: safeDetails
    });

    // Cap at 100 events to prevent property size overflow
    if (queue.length > 100) {
      queue = queue.slice(-100);
    }

    props.setProperty('SECURITY_DIGEST_QUEUE', JSON.stringify(queue));
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
      try { auditIntegrity = verifyAuditLogIntegrity(); } catch (_e) { /* skip */ }
    }
    if (typeof verifySurveyVaultIntegrity === 'function') {
      try { vaultIntegrity = verifySurveyVaultIntegrity(); } catch (_e) { /* skip */ }
    }

    // Gather recipients
    var recipients = [];
    if (typeof getConfigValue_ === 'function') {
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
      } catch (_e) { /* skip */ }
    }
    if (recipients.length === 0) {
      try {
        var ownerEmail = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
        if (ownerEmail) recipients.push(ownerEmail);
      } catch (_e) { /* skip */ }
    }

    if (recipients.length === 0 || MailApp.getRemainingDailyQuota() < 1) return;

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

    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subjectPrefix + ' Daily Security Digest — ' + queue.length + ' event(s)',
      body: body
    });

    // Clear the queue
    props.setProperty('SECURITY_DIGEST_QUEUE', '[]');

    Logger.log('Security digest sent with ' + queue.length + ' events');

  } catch (e) {
    Logger.log('Failed to send security digest: ' + e.message);
  }
}

/**
 * Installs the daily security digest trigger.
 * Runs at 7 AM in the script's timezone.
 */
function installSecurityDigestTrigger() {
  // Remove existing security digest triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailySecurityDigest') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create daily trigger at 7 AM
  ScriptApp.newTrigger('sendDailySecurityDigest')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ Security Digest Enabled',
    'A daily security digest email will be sent at 7 AM to the Chief Steward and Admin.\n\n' +
    'The digest includes:\n' +
    '• All security events from the past 24 hours\n' +
    '• Audit log integrity check results\n' +
    '• Survey vault integrity check results',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Shows a security status overview dialog for admins.
 * Displays current security posture at a glance.
 */
function showSecurityStatusDialog() {
  var ui = SpreadsheetApp.getUi();

  // Check pending digest events
  var pendingEvents = 0;
  try {
    var props = PropertiesService.getScriptProperties();
    var queue = JSON.parse(props.getProperty('SECURITY_DIGEST_QUEUE') || '[]');
    pendingEvents = queue.length;
  } catch (_e) { /* skip */ }

  // Check audit integrity
  var auditStatus = 'Not checked';
  if (typeof verifyAuditLogIntegrity === 'function') {
    try {
      var auditResult = verifyAuditLogIntegrity();
      auditStatus = auditResult.valid ? 'PASS (' + auditResult.totalRows + ' entries)' :
        'FAIL — ' + auditResult.invalidRows.length + ' tampered row(s)';
      if (auditResult.message) auditStatus += '\n   ' + auditResult.message;
    } catch (_e) {
      auditStatus = 'Error running check';
    }
  }

  // Check vault integrity
  var vaultStatus = 'Not checked';
  if (typeof verifySurveyVaultIntegrity === 'function') {
    try {
      var vaultResult = verifySurveyVaultIntegrity();
      vaultStatus = vaultResult.valid ? 'PASS (' + vaultResult.stats.totalEntries + ' entries)' :
        'ISSUES — ' + vaultResult.issues.length + ' problem(s)';
    } catch (_e) {
      vaultStatus = 'Error running check';
    }
  }

  // Check dashboard auth
  var dashAuthStatus = 'Disabled (open access)';
  if (typeof isDashboardMemberAuthRequired === 'function') {
    dashAuthStatus = isDashboardMemberAuthRequired() ? 'ENABLED (PIN required)' : 'Disabled (open access)';
  }

  // Check for digest trigger
  var digestTriggerInstalled = false;
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'sendDailySecurityDigest') {
        digestTriggerInstalled = true;
        break;
      }
    }
  } catch (_e) { /* skip */ }

  var message =
    'SECURITY POSTURE OVERVIEW\n\n' +
    'Audit Log Integrity: ' + auditStatus + '\n' +
    'Survey Vault Integrity: ' + vaultStatus + '\n' +
    'Dashboard Auth: ' + dashAuthStatus + '\n' +
    'Daily Digest Trigger: ' + (digestTriggerInstalled ? 'Installed' : 'NOT installed') + '\n' +
    'Pending Security Events: ' + pendingEvents + '\n\n' +
    'To enable daily alerts:\n' +
    '  Run: installSecurityDigestTrigger()';

  ui.alert('🛡️ Security Status', message, ui.ButtonSet.OK);
}
