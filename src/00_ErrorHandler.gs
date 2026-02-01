/**
 * ============================================================================
 * 00_ErrorHandler.gs - Centralized Error Handling
 * ============================================================================
 *
 * This module provides centralized error handling, logging, and user
 * notification utilities for the entire dashboard application.
 *
 * Features:
 * - Consistent error logging with context
 * - User-friendly error notifications
 * - Error tracking and reporting
 * - Performance monitoring
 * - Input sanitization
 *
 * @fileoverview Centralized error handling utilities
 * @version 1.0.0
 */

// ============================================================================
// ERROR HANDLING CONFIGURATION
// ============================================================================

/**
 * Error severity levels
 * @enum {string}
 */
var ERROR_LEVEL = {
  DEBUG: 'DEBUG',
  INFO: 'INFO',
  WARNING: 'WARNING',
  ERROR: 'ERROR',
  CRITICAL: 'CRITICAL'
};

/**
 * Error handler configuration
 */
var ERROR_CONFIG = {
  LOG_TO_SHEET: true,
  LOG_SHEET_NAME: '_ErrorLog',
  MAX_LOG_ROWS: 1000,
  SHOW_STACK_TRACE: false,
  NOTIFY_ON_CRITICAL: true
};

// ============================================================================
// CORE ERROR HANDLING
// ============================================================================

/**
 * Wraps a function with error handling
 * @param {Function} fn - Function to wrap
 * @param {string} context - Context description for error messages
 * @returns {Function} Wrapped function
 */
function withErrorHandling(fn, context) {
  return function() {
    try {
      return fn.apply(this, arguments);
    } catch (error) {
      handleError(error, context);
      return null;
    }
  };
}

/**
 * Central error handler
 * @param {Error} error - The error object
 * @param {string} context - Where the error occurred
 * @param {string} [level] - Error severity level
 */
function handleError(error, context, level) {
  level = level || ERROR_LEVEL.ERROR;

  var errorInfo = {
    timestamp: new Date().toISOString(),
    level: level,
    context: context || 'Unknown',
    message: error.message || String(error),
    stack: error.stack || '',
    user: Session.getActiveUser().getEmail() || 'Unknown'
  };

  // Log to console
  Logger.log('[' + level + '] ' + context + ': ' + errorInfo.message);

  // Log to sheet if enabled
  if (ERROR_CONFIG.LOG_TO_SHEET) {
    logErrorToSheet_(errorInfo);
  }

  // Show user notification for errors
  if (level === ERROR_LEVEL.ERROR || level === ERROR_LEVEL.CRITICAL) {
    showErrorNotification_(errorInfo);
  }

  // Send notification for critical errors
  if (level === ERROR_LEVEL.CRITICAL && ERROR_CONFIG.NOTIFY_ON_CRITICAL) {
    sendCriticalErrorNotification_(errorInfo);
  }

  return errorInfo;
}

/**
 * Log error to a dedicated sheet
 * @param {Object} errorInfo - Error information object
 * @private
 */
function logErrorToSheet_(errorInfo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(ERROR_CONFIG.LOG_SHEET_NAME);
      sheet.hideSheet();
      sheet.appendRow(['Timestamp', 'Level', 'Context', 'Message', 'User', 'Stack']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    // Add error row
    sheet.appendRow([
      errorInfo.timestamp,
      errorInfo.level,
      errorInfo.context,
      errorInfo.message,
      errorInfo.user,
      ERROR_CONFIG.SHOW_STACK_TRACE ? errorInfo.stack : ''
    ]);

    // Trim old entries if over limit
    var lastRow = sheet.getLastRow();
    if (lastRow > ERROR_CONFIG.MAX_LOG_ROWS) {
      sheet.deleteRows(2, lastRow - ERROR_CONFIG.MAX_LOG_ROWS);
    }
  } catch (e) {
    Logger.log('Failed to log error to sheet: ' + e.message);
  }
}

/**
 * Show error notification to user
 * @param {Object} errorInfo - Error information object
 * @private
 */
function showErrorNotification_(errorInfo) {
  try {
    var ui = SpreadsheetApp.getUi();
    var message = 'An error occurred in ' + errorInfo.context + ':\n\n' + errorInfo.message;

    if (errorInfo.level === ERROR_LEVEL.CRITICAL) {
      ui.alert('Critical Error', message, ui.ButtonSet.OK);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        errorInfo.message,
        'Error: ' + errorInfo.context,
        10
      );
    }
  } catch (e) {
    Logger.log('Failed to show error notification: ' + e.message);
  }
}

/**
 * Send notification for critical errors
 * @param {Object} errorInfo - Error information object
 * @private
 */
function sendCriticalErrorNotification_(errorInfo) {
  try {
    var adminEmail = Session.getEffectiveUser().getEmail();
    var subject = '[509 Dashboard] Critical Error: ' + errorInfo.context;
    var body = 'A critical error occurred in the 509 Dashboard:\n\n' +
               'Time: ' + errorInfo.timestamp + '\n' +
               'Context: ' + errorInfo.context + '\n' +
               'Message: ' + errorInfo.message + '\n' +
               'User: ' + errorInfo.user + '\n\n' +
               'Stack Trace:\n' + errorInfo.stack;

    MailApp.sendEmail(adminEmail, subject, body);
  } catch (e) {
    Logger.log('Failed to send critical error notification: ' + e.message);
  }
}

// ============================================================================
// INPUT SANITIZATION
// ============================================================================

/**
 * Sanitize HTML to prevent XSS attacks
 * @param {string} input - Input string to sanitize
 * @returns {string} Sanitized string
 */
function sanitizeHtml(input) {
  if (!input) return '';
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;');
}

/**
 * Sanitize input for use in SQL-like queries
 * @param {string} input - Input string to sanitize
 * @returns {string} Sanitized string
 */
function sanitizeForQuery(input) {
  if (!input) return '';
  return String(input)
    .replace(/'/g, "''")
    .replace(/\\/g, '\\\\');
}

/**
 * Validate and sanitize email address
 * @param {string} email - Email to validate
 * @returns {string|null} Sanitized email or null if invalid
 */
function sanitizeEmail(email) {
  if (!email) return null;
  var trimmed = String(email).trim().toLowerCase();
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(trimmed) ? trimmed : null;
}

/**
 * Sanitize phone number (US format)
 * @param {string} phone - Phone number to sanitize
 * @returns {string} Sanitized phone number (digits only)
 */
function sanitizePhone(phone) {
  if (!phone) return '';
  return String(phone).replace(/\D/g, '').slice(0, 10);
}

/**
 * Sanitize sheet name for Google Sheets
 * @param {string} name - Sheet name to sanitize
 * @returns {string} Sanitized sheet name
 */
function sanitizeSheetName(name) {
  if (!name) return 'Sheet';
  return String(name)
    .replace(/[\\\/\?\*\[\]\']/g, '')
    .substring(0, 100);
}

// ============================================================================
// PERFORMANCE MONITORING
// ============================================================================

/**
 * Performance timer for tracking operation duration
 */
var PerformanceTimer = {
  timers: {},

  /**
   * Start a timer
   * @param {string} name - Timer name
   */
  start: function(name) {
    this.timers[name] = {
      start: new Date().getTime(),
      end: null
    };
  },

  /**
   * Stop a timer and return duration
   * @param {string} name - Timer name
   * @returns {number} Duration in milliseconds
   */
  stop: function(name) {
    if (!this.timers[name]) {
      Logger.log('Timer not found: ' + name);
      return 0;
    }
    this.timers[name].end = new Date().getTime();
    var duration = this.timers[name].end - this.timers[name].start;
    Logger.log('[PERF] ' + name + ': ' + duration + 'ms');
    return duration;
  },

  /**
   * Log performance if over threshold
   * @param {string} name - Timer name
   * @param {number} threshold - Threshold in ms
   */
  logIfSlow: function(name, threshold) {
    var duration = this.stop(name);
    if (duration > threshold) {
      handleError(
        new Error('Slow operation: ' + duration + 'ms'),
        name,
        ERROR_LEVEL.WARNING
      );
    }
    return duration;
  }
};

// ============================================================================
// CONSTANTS VALIDATION
// ============================================================================

/**
 * Validate that required constants are defined
 * @returns {Object} Validation result with missing items
 */
function validateConstants() {
  var missing = [];
  var warnings = [];

  // Check SHEETS constant
  if (typeof SHEETS === 'undefined') {
    missing.push('SHEETS constant not defined');
  } else {
    var requiredSheets = ['MEMBER_DIR', 'GRIEVANCE_LOG', 'CONFIG', 'DASHBOARD'];
    requiredSheets.forEach(function(sheet) {
      if (!SHEETS[sheet]) {
        warnings.push('SHEETS.' + sheet + ' not defined');
      }
    });
  }

  // Check column constants
  if (typeof MEMBER_COLS === 'undefined') {
    missing.push('MEMBER_COLS constant not defined');
  }

  if (typeof GRIEVANCE_COLS === 'undefined') {
    missing.push('GRIEVANCE_COLS constant not defined');
  }

  return {
    valid: missing.length === 0,
    missing: missing,
    warnings: warnings
  };
}

/**
 * Validate required sheets exist
 * @returns {Object} Validation result
 */
function validateRequiredSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var missing = [];

  if (typeof SHEETS !== 'undefined') {
    var criticalSheets = [SHEETS.MEMBER_DIR, SHEETS.GRIEVANCE_LOG, SHEETS.CONFIG];
    criticalSheets.forEach(function(sheetName) {
      if (sheetName && !ss.getSheetByName(sheetName)) {
        missing.push(sheetName);
      }
    });
  }

  return {
    valid: missing.length === 0,
    missing: missing
  };
}

/**
 * Run all startup validations
 * @returns {boolean} True if all validations pass
 */
function runStartupValidation() {
  var constantsResult = validateConstants();
  var sheetsResult = validateRequiredSheets();

  if (!constantsResult.valid) {
    handleError(
      new Error('Missing constants: ' + constantsResult.missing.join(', ')),
      'Startup Validation',
      ERROR_LEVEL.CRITICAL
    );
    return false;
  }

  if (constantsResult.warnings.length > 0) {
    Logger.log('[WARNING] Constants warnings: ' + constantsResult.warnings.join(', '));
  }

  if (!sheetsResult.valid) {
    handleError(
      new Error('Missing sheets: ' + sheetsResult.missing.join(', ')),
      'Startup Validation',
      ERROR_LEVEL.WARNING
    );
  }

  Logger.log('[INFO] Startup validation completed');
  return true;
}

// ============================================================================
// API VERSIONING
// ============================================================================

/**
 * API version information
 */
var API_VERSION = {
  major: 4,
  minor: 5,
  patch: 0,
  toString: function() {
    return this.major + '.' + this.minor + '.' + this.patch;
  }
};

/**
 * Get current API version
 * @returns {string} Version string
 */
function getApiVersion() {
  return API_VERSION.toString();
}

/**
 * Check if client version is compatible
 * @param {string} clientVersion - Client version string (e.g., "4.2.0")
 * @returns {boolean} True if compatible
 */
function isVersionCompatible(clientVersion) {
  if (!clientVersion) return false;
  var parts = clientVersion.split('.');
  var clientMajor = parseInt(parts[0], 10);
  return clientMajor === API_VERSION.major;
}

// ============================================================================
// ERROR LOG VIEWER
// ============================================================================

/**
 * Show error log viewer dialog
 */
function showErrorLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('No error log found.');
    return;
  }

  sheet.showSheet();
  ss.setActiveSheet(sheet);
}

/**
 * Clear error log
 */
function clearErrorLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Error log cleared', 'Success');
  }
}
