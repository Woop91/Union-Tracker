/**
 * ============================================================================
 * 01_Core.gs - Core Constants, Error Handling & Configuration
 * ============================================================================
 *
 * This module provides centralized error handling, logging, constants,
 * configuration, and column definitions for the entire dashboard application.
 *
 * Features:
 * - Consistent error logging with context
 * - User-friendly error notifications
 * - Error tracking and reporting
 * - Performance monitoring
 * - Input sanitization
 * - Sheet and column constants
 * - Version management
 *
 * @fileoverview Core constants, error handling, and configuration
 * @version 4.6.0
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
 * Creates a standardized success response object
 * @param {*} [data] - Optional response data
 * @param {string} [message] - Optional success message
 * @returns {{success: true, data: *, message: string}}
 */
function successResponse(data, message) {
  return { success: true, data: data || null, message: message || '' };
}

/**
 * Creates a standardized error response object
 * @param {string} error - Error message
 * @param {string} [context] - Context where error occurred
 * @returns {{success: false, error: string, context: string}}
 */
function errorResponse(error, context) {
  return { success: false, error: error || 'An unexpected error occurred', context: context || '' };
}

/**
 * Normalizes boolean-like values from Google Sheets to a consistent boolean.
 * Handles: true, 'Yes', 'TRUE', 'true', 'yes', 1, '1'
 * @param {*} value - The value to check
 * @returns {boolean} True if the value represents a truthy/yes value
 */
function isTruthyValue(value) {
  if (value === true || value === 1) return true;
  if (typeof value === 'string') {
    var lower = value.trim().toLowerCase();
    return lower === 'yes' || lower === 'true' || lower === '1';
  }
  return false;
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
    user: (function() { try { return Session.getActiveUser().getEmail() || 'Unknown'; } catch (_e) { return 'Unknown'; } })()
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
    var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Critical Error: ' + errorInfo.context;
    var body = 'A critical error occurred in the ' + COMMAND_CONFIG.SYSTEM_NAME + ':\n\n' +
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
  // Delegate to escapeHtml in 00_Security.gs for consistent escaping
  return escapeHtml(input);
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
  minor: 6,
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



/**
 * 509 Dashboard - Constants and Configuration
 *
 * Single source of truth for all configuration constants.
 * This file must be loaded first in the build order.
 *
 * @version 4.6.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// STRATEGIC COMMAND CENTER CONFIG (v4.0)
// ============================================================================

/**
 * Strategic Command Center configuration
 * Controls system-wide settings for the 509 Dashboard
 *
 * v4.0 UNIFIED MASTER ENGINE FEATURES:
 * - Security: Audit Log & Sabotage Protection (>15 cells)
 * - Performance: Batch Array Processing (No-Lag Architecture)
 * - Workflow: Stage-Gate Case Tracking & Auto-PDF Generation
 * - Production: Nuke/Seed Isolation & UI Self-Hiding
 * - Accessibility: Mobile/Pocket View & Search Engine
 *
 * @const {Object}
 */
var COMMAND_CONFIG = {
  // System Identity — reads from Config sheet at runtime, falls back to defaults
  get SYSTEM_NAME() { return getSystemName_(); },
  VERSION: "4.6.0",

  // Document Templates (configure these with your Drive IDs)
  TEMPLATE_ID: '',  // Google Doc template ID for grievance PDFs
  ARCHIVE_FOLDER_ID: '',  // Drive folder ID for archived documents

  // Alert Configuration
  CHIEF_STEWARD_EMAIL: '',  // Read from Config sheet (CONFIG_COLS.CHIEF_STEWARD_EMAIL)
  // Escalation triggers - for STATUS column use status values, for CURRENT_STEP use step values
  ESCALATION_STATUSES: ['In Arbitration', 'Appealed'],  // Status values that trigger alerts
  ESCALATION_STEPS: ['Step II', 'Step III', 'Arbitration'],  // Step values that trigger alerts

  // Unit Code Prefixes - Now read from Config sheet (CONFIG_COLS.UNIT_CODES)
  // Format in Config: "Main Station:MS,Field Ops:FO,Health:HC,Admin:AD,Remote:RM"
  UNIT_CODES: {},  // Populated dynamically from Config sheet

  // Theme Settings (Roboto-based)
  THEME: {
    HEADER_BG: '#1e293b',
    HEADER_TEXT: '#ffffff',
    ALT_ROW: '#f8fafc',
    FONT: 'Roboto',
    FONT_SIZE: 10,
    HEADER_SIZE: 11
  },

  // Status Color Mapping for Auto-Styling (matches DEFAULT_CONFIG.GRIEVANCE_STATUS)
  STATUS_COLORS: {
    "Open": { bg: "#fef3c7", text: "#92400e" },           // Yellow/Orange - Active case
    "Pending Info": { bg: "#e0e7ff", text: "#3730a3" },   // Purple - Waiting on info
    "Settled": { bg: "#dbeafe", text: "#1e40af" },        // Blue - Negotiated resolution
    "Withdrawn": { bg: "#f1f5f9", text: "#475569" },      // Gray - Case withdrawn
    "Denied": { bg: "#fee2e2", text: "#991b1b" },         // Red - Loss
    "Won": { bg: "#dcfce7", text: "#166534" },            // Green - Victory
    "Appealed": { bg: "#fef3c7", text: "#92400e" },       // Yellow - Escalated
    "In Arbitration": { bg: "#fee2e2", text: "#991b1b" }, // Red - High stakes
    "Closed": { bg: "#f1f5f9", text: "#475569" }          // Gray - Complete
  },

  // PDF Generation Settings
  PDF: {
    get AUTHOR() { return getSystemName_(); },
    SIGNATURE_BLOCK: "\n\n__________________________\nMember Signature\n\n__________________________\nSteward Signature\n\n__________________________\nDate"
  },

  // Email Branding
  EMAIL: {
    get SUBJECT_PREFIX() { return '[' + getSystemName_() + ']'; },
    get FOOTER() { return '\n\n---\nGenerated by ' + getSystemName_(); }
  }
};

/**
 * Drive folder configuration for grievance document management
 * Simplified folder naming format: LastName, FirstName - YYYY-MM-DD
 * Example: "Smith, John - 2026-01-15"
 *
 * @const {Object}
 */
var DRIVE_CONFIG = {
  ROOT_FOLDER_NAME: '509 Dashboard - Grievance Files',
  // Simplified template: Member Name and Date Filed
  // Template uses placeholders: {date}, {lastName}, {firstName}
  SUBFOLDER_TEMPLATE: '{lastName}, {firstName} - {date}',
  // Fallback if member name not available
  SUBFOLDER_TEMPLATE_SIMPLE: '{grievanceId} - {date}'
};

/**
 * Get organization name from Config sheet, falling back to default.
 * Single source of truth for the org name used across the system.
 * @private
 * @returns {string} Organization name (e.g., "SEIU Local 509")
 */
function getOrgNameFromConfig_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue();
      if (orgName) return orgName;
    }
  } catch (_e) {
    // Fallback silently during initialization or when spreadsheet is unavailable
  }
  return 'SEIU Local 509';
}

/**
 * Get system name from Config sheet org name, falling back to default.
 * Derives the tool/system name from the org name for branding consistency.
 * Used by COMMAND_CONFIG.SYSTEM_NAME, EMAIL, and PDF getters.
 * @private
 * @returns {string} System name (e.g., "509 Strategic Command Center")
 */
function getSystemName_() {
  return getLocalNumberFromConfig_() + ' Strategic Command Center';
}

/**
 * Get local number from Config sheet, falling back to default.
 * Used for UI elements like menu names.
 * @private
 * @returns {string} Local number (e.g., "509")
 */
function getLocalNumberFromConfig_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var localNumber = configSheet.getRange(3, CONFIG_COLS.LOCAL_NUMBER).getValue();
      if (localNumber) return String(localNumber);
    }
  } catch (_e) {
    // Fallback silently during initialization or when spreadsheet is unavailable
  }
  return '509';
}

// ============================================================================
// VERSION INFO
// ============================================================================

/**
 * Version information for build system and display
 * @const {Object}
 */
var VERSION_INFO = {
  MAJOR: 4,
  MINOR: 6,
  PATCH: 0,
  BUILD: 'v4.6.0',
  CURRENT: '4.6.0',
  BUILD_DATE: '2026-02-12',
  CODENAME: 'Meeting Intelligence & Document Automation'
};

/**
 * Complete version history with release dates and codenames.
 * Ordered newest-first. Every version that has ever shipped is listed here
 * so that UI, audit, and diagnostic code can look up any past release date.
 * @const {Array<Object>}
 */
var VERSION_HISTORY = [
  { version: '4.6.0', date: '2026-02-12', codename: 'Meeting Intelligence & Document Automation', changes: 'Meeting Notes & Agenda doc automation, two-tier steward agenda sharing, Meeting Notes dashboard tab, member Drive folders, meeting event scheduling. Added Employee ID, Department, Hire Date columns to Member Directory. Added PII mailing address columns (Street, City, State) hidden by default. Added Last Updated to Grievance Log. Fixed diagnostics checks. Removed deprecated Dashboard/Satisfaction from sheet ordering. Added Export (seiu509.org only) and Lockdown future feature roadmap items.' },
  { version: '4.5.1', date: '2026-02-11', codename: 'Engagement Fixes',                          changes: 'Engagement tracking fixes, 950 Jest tests, GRIEVANCE_OUTCOMES/generateGrievanceId fixes' },
  { version: '4.5.0', date: '2026-02-01', codename: 'Security & Testing',                         changes: 'Security module, Data Access Layer, Member Self-Service, consolidated to 16 source files' },
  { version: '4.4.1', date: '2026-01-31', codename: 'Build System',                               changes: 'Initial build system with Node.js, source file concatenation' },
  { version: '4.4.0', date: '2026-01-30', codename: 'Menu Consolidation',                         changes: 'Unified 3-menu system (Union Hub, Tools, Admin), web app dashboards, deprecated command center menu' },
  { version: '4.3.8', date: '2026-01-28', codename: 'Features Reference',                         changes: 'Satisfaction modal dashboard, Features Reference sheet, hidden satisfaction sheet' },
  { version: '4.3.7', date: '2026-01-25', codename: 'Help System',                                changes: 'Complete rewrite of help guide with real-time search, menu reference, and FAQ tabs' },
  { version: '4.3.2', date: '2026-01-20', codename: 'Modal Dashboards',                           changes: 'Deprecated visible Dashboard sheet, switched to SPA-style modal dashboards' },
  { version: '4.3.0', date: '2026-01-15', codename: 'Checklist & Looker',                         changes: 'Case Checklist system, Looker data integration, dynamic field expansion engine' },
  { version: '4.1.0', date: '2026-01-10', codename: 'Strategic Config',                           changes: 'Strategic Command Center config, status color mapping, PDF/email branding' },
  { version: '4.0.0', date: '2026-01-05', codename: 'Strategic Command Center',                   changes: 'Unified master engine, audit logging, sabotage protection, batch processing, mobile views' },
  { version: '3.6.0', date: '2025-12-20', codename: 'Data Managers',                              changes: 'Member and Grievance data manager refactor, improved validation' },
  { version: '2.0.0', date: '2025-11-15', codename: 'Modular Architecture',                       changes: 'Split monolith into modular source files, build system, UI/business logic separation' }
];

/**
 * Look up release info for a specific version.
 * @param {string} ver - Version string, e.g. "4.5.0"
 * @returns {Object|null} The matching VERSION_HISTORY entry, or null.
 */
function getVersionDate(ver) {
  for (var i = 0; i < VERSION_HISTORY.length; i++) {
    if (VERSION_HISTORY[i].version === ver) {
      return VERSION_HISTORY[i];
    }
  }
  return null;
}

// ============================================================================
// SHEET NAMES
// ============================================================================

/**
 * Sheet name constants - use these instead of hardcoded strings
 * @const {Object}
 */
var SHEETS = {
  CONFIG: 'Config',
  MEMBER_DIR: 'Member Directory',
  GRIEVANCE_LOG: 'Grievance Log',
  // Dashboard sheets
  // @deprecated v4.3.2 - Dashboard sheet is deprecated. Use modal dashboards instead.
  // Access via: Union Hub > Dashboards menu. Use removeDeprecatedTabs() to remove.
  DASHBOARD: '💼 Dashboard',
  // Hidden calculation sheets (self-healing formulas)
  GRIEVANCE_CALC: '_Grievance_Calc',
  GRIEVANCE_FORMULAS: '_Grievance_Formulas',
  MEMBER_LOOKUP: '_Member_Lookup',
  STEWARD_CONTACT_CALC: '_Steward_Contact_Calc',
  DASHBOARD_CALC: '_Dashboard_Calc',
  STEWARD_PERFORMANCE_CALC: '_Steward_Performance_Calc',
  // Optional source sheets
  MEETING_ATTENDANCE: '📅 Meeting Attendance',
  MEETING_CHECKIN_LOG: '📝 Meeting Check-In Log',
  VOLUNTEER_HOURS: '🤝 Volunteer Hours',
  // Test Results
  TEST_RESULTS: 'Test Results',
  // Function Checklist
  FUNCTION_CHECKLIST: 'Function Checklist',
  // Audit Log (hidden)
  AUDIT_LOG: '_Audit_Log',
  // Case Checklist
  CASE_CHECKLIST: 'Case Checklist',
  // Satisfaction & Feedback sheets
  // @deprecated v4.3.8 - Satisfaction sheet is now hidden. Use showSatisfactionDashboard() modal instead.
  // Data is preserved for modal access. Use removeDeprecatedTabs() to hide.
  SATISFACTION: '📊 Member Satisfaction',
  FEEDBACK: '💡 Feedback & Development',
  // Help & Documentation sheets
  GETTING_STARTED: '📚 Getting Started',
  FAQ: '❓ FAQ',
  CONFIG_GUIDE: '📖 Config Guide',
  FEATURES_REFERENCE: '📋 Features Reference',
  // Aliases for backward compatibility (some code uses these alternate names)
  GRIEVANCE_TRACKER: 'Grievance Log',
  MEMBER_DIRECTORY: 'Member Directory',
  REPORTS: '💼 Dashboard'
};

// SHEET_NAMES alias for backward compatibility
// Some code references SHEET_NAMES instead of SHEETS
var SHEET_NAMES = SHEETS;

/**
 * Hidden sheets used for calculations (prefixed with underscore)
 * These are auto-created and hidden from users
 * @const {Object}
 */
var HIDDEN_SHEETS = {
  CALC_STATS: '_Dashboard_Calc',
  CALC_FORMULAS: '_Grievance_Formulas',
  GRIEVANCE_CALC: '_Grievance_Calc',
  MEMBER_LOOKUP: '_Member_Lookup',
  STEWARD_CONTACT_CALC: '_Steward_Contact_Calc',
  STEWARD_PERFORMANCE_CALC: '_Steward_Performance_Calc',
  AUDIT_LOG: '_Audit_Log',
  CHECKLIST_CALC: '_Checklist_Calc'
};

// ============================================================================
// COLOR SCHEME - Enhanced Visual Theme System
// ============================================================================

/**
 * Color scheme constants for consistent branding
 * @const {Object}
 */
var COLORS = {
  // Primary Brand Colors
  PRIMARY_PURPLE: '#7C3AED',    // Main brand purple
  UNION_GREEN: '#059669',       // Union/success green
  SOLIDARITY_RED: '#DC2626',    // Alert/urgent red
  PRIMARY_BLUE: '#7EC8E3',      // Light blue
  ACCENT_ORANGE: '#F97316',     // Warnings/attention
  LIGHT_GRAY: '#F3F4F6',        // Backgrounds
  TEXT_DARK: '#1F2937',         // Primary text
  WHITE: '#FFFFFF',             // White
  HEADER_BG: '#7C3AED',         // Header background (same as primary)
  HEADER_TEXT: '#FFFFFF',       // Header text color

  // Card Theme Colors (Dark Mode Headers)
  CARD_DARK_BG: '#1E293B',      // Dark slate for card headers
  CARD_DARK_TEXT: '#F8FAFC',    // Light text on dark backgrounds
  CARD_GRADIENT_START: '#4C1D95', // Purple gradient start
  CARD_GRADIENT_MID: '#5B21B6',   // Purple gradient middle
  CARD_ACCENT_GLOW: '#A78BFA',    // Purple accent glow

  // Status Colors (Enhanced)
  STATUS_GREEN: '#10B981',      // Success/Won
  STATUS_YELLOW: '#FBBF24',     // Warning/4-7 days
  STATUS_ORANGE: '#F97316',     // Caution/1-3 days
  STATUS_RED: '#EF4444',        // Danger/Overdue
  STATUS_BLUE: '#3B82F6',       // Info/Pending

  // Gradient Scale Colors (for heatmaps)
  GRADIENT_LOW: '#D1FAE5',      // Light green (good)
  GRADIENT_MID_LOW: '#FEF3C7',  // Light yellow
  GRADIENT_MID: '#FED7AA',      // Light orange
  GRADIENT_MID_HIGH: '#FECACA', // Light red
  GRADIENT_HIGH: '#FCA5A5',     // Red (attention needed)

  // Chart Colors
  CHART_PURPLE: '#8B5CF6',
  CHART_BLUE: '#3B82F6',
  CHART_GREEN: '#10B981',
  CHART_YELLOW: '#F59E0B',
  CHART_RED: '#EF4444',
  CHART_PINK: '#EC4899',
  CHART_INDIGO: '#6366F1',
  CHART_CYAN: '#06B6D4',

  // Section Colors (Cohesive Theme)
  SECTION_STATS: '#059669',     // Green - Quick Stats
  SECTION_MEMBERS: '#3B82F6',   // Blue - Member Metrics
  SECTION_GRIEVANCE: '#F97316', // Orange - Grievance Metrics
  SECTION_TIMELINE: '#7C3AED',  // Purple - Timeline
  SECTION_ANALYSIS: '#6366F1',  // Indigo - Analysis
  SECTION_LOCATION: '#0891B2',  // Cyan - Location
  SECTION_TRENDS: '#DC2626',    // Red - Trends
  SECTION_PERFORMANCE: '#8B5CF6', // Violet - Performance

  // Alternate Row Colors (Zebra Stripes)
  ROW_ALT_LIGHT: '#F9FAFB',     // Very light gray
  ROW_ALT_GREEN: '#ECFDF5',     // Light green tint
  ROW_ALT_RED: '#FEF2F2',       // Light red tint
  ROW_ALT_BLUE: '#EFF6FF'       // Light blue tint
};

/**
 * Mobile interface configuration settings
 * Shared constants for responsive design across all UI modules
 * @const {Object}
 */
var MOBILE_CONFIG = {
  MAX_COLUMNS_MOBILE: 8,
  CARD_LAYOUT_ENABLED: true,
  TOUCH_TARGET_SIZE: '44px',
  MOBILE_BREAKPOINT: 768,   // Width in pixels below which is considered mobile
  TABLET_BREAKPOINT: 1024   // Width in pixels below which is considered tablet
};

/**
 * UI Theme constants for dialogs and sidebars
 * Used in HTML templates for consistent styling
 * @const {Object}
 */
var UI_THEME = {
  // Primary Colors
  PRIMARY_COLOR: '#7C3AED',       // Main brand purple
  SECONDARY_COLOR: '#64748B',     // Secondary gray

  // Text Colors
  TEXT_PRIMARY: '#1F2937',        // Dark text
  TEXT_SECONDARY: '#6B7280',      // Gray text

  // UI Elements
  BORDER_COLOR: '#E5E7EB',        // Light gray borders
  BACKGROUND: '#F9FAFB',          // Light background
  CARD_BG: '#FFFFFF',             // White card background

  // Status Colors
  SUCCESS_COLOR: '#10B981',       // Green
  WARNING_COLOR: '#F59E0B',       // Amber
  DANGER_COLOR: '#EF4444',        // Red
  INFO_COLOR: '#3B82F6',          // Blue

  // Dark Mode (for dark headers)
  DARK_BG: '#1E293B',             // Dark slate
  DARK_TEXT: '#F8FAFC',           // Light text on dark

  // Gradients
  GRADIENT_START: '#7C3AED',
  GRADIENT_END: '#5B21B6'
};

/**
 * Standard dialog sizes for modals and sidebars
 * @const {Object}
 */
var DIALOG_SIZES = {
  SMALL: { width: 400, height: 300 },
  MEDIUM: { width: 600, height: 500 },
  LARGE: { width: 800, height: 650 },
  FULLSCREEN: { width: 1000, height: 750 },
  SIDEBAR: { width: 300 }
};

/**
 * Menu Icon Mapping - Consistent Iconography
 * Action icons for tools, View icons for dashboards
 * @const {Object}
 */
var MENU_ICONS = {
  // Dashboard/View Icons (📊 style)
  DASHBOARD: '📊',
  MEMBERS: '👥',
  GRIEVANCES: '📋',
  CALENDAR: '📅',
  REPORTS: '📈',
  SATISFACTION: '⭐',

  // Action Icons (➕ style)
  ADD: '➕',
  EDIT: '✏️',
  DELETE: '🗑️',
  SEARCH: '🔍',
  REFRESH: '🔄',
  SYNC: '🔗',
  EXPORT: '📤',
  IMPORT: '📥',

  // Status Icons
  SUCCESS: '✅',
  WARNING: '⚠️',
  ERROR: '❌',
  INFO: 'ℹ️',
  CLOCK: '⏰',

  // Admin/Settings Icons
  SETTINGS: '⚙️',
  TOOLS: '🛠️',
  ADMIN: '🔧',
  HELP: '❓',
  DOCS: '📖'
};

// ============================================================================
// MEMBER DIRECTORY COLUMNS (40 columns total: A-AN)
// ============================================================================

/**
 * Member Directory column positions (1-indexed)
 * CRITICAL: ALL column references must use these constants
 * @const {Object}
 */
var MEMBER_COLS = {
  // Section 1: Identity & Core Info (A-D)
  MEMBER_ID: 1,                    // A
  FIRST_NAME: 2,                   // B
  LAST_NAME: 3,                    // C
  JOB_TITLE: 4,                    // D

  // Section 2: Location & Work (E-H)
  WORK_LOCATION: 5,                // E
  UNIT: 6,                         // F
  CUBICLE: 7,                      // G - Cubicle / workspace ID (hidden by default)
  OFFICE_DAYS: 8,                  // H - Multi-select: days member works in office

  // Section 3: Contact Information (I-L)
  EMAIL: 9,                        // I
  PHONE: 10,                       // J
  PREFERRED_COMM: 11,              // K - Multi-select: preferred communication methods
  BEST_TIME: 12,                   // L - Multi-select: best times to reach member

  // Section 4: Organizational Structure (M-Q)
  SUPERVISOR: 13,                  // M
  MANAGER: 14,                     // N
  IS_STEWARD: 15,                  // O
  COMMITTEES: 16,                  // P - Multi-select: which committees steward is in
  ASSIGNED_STEWARD: 17,            // Q - Multi-select: assigned steward(s)

  // Section 5: Engagement Metrics (R-U) - Hidden by default
  LAST_VIRTUAL_MTG: 18,            // R
  LAST_INPERSON_MTG: 19,           // S
  OPEN_RATE: 20,                   // T
  VOLUNTEER_HOURS: 21,             // U

  // Section 6: Member Interests (V-X) - Hidden by default
  INTEREST_LOCAL: 22,              // V
  INTEREST_CHAPTER: 23,            // W
  INTEREST_ALLIED: 24,             // X

  // Section 7: Steward Contact Tracking (Y-AA)
  RECENT_CONTACT_DATE: 25,         // Y
  CONTACT_STEWARD: 26,             // Z
  CONTACT_NOTES: 27,               // AA

  // Section 8: Grievance Management (AB-AE)
  HAS_OPEN_GRIEVANCE: 28,          // AB - Script-calculated (static value)
  GRIEVANCE_STATUS: 29,            // AC - Script-calculated (static value)
  NEXT_DEADLINE: 30,               // AD - Script-calculated (static value)
  START_GRIEVANCE: 31,             // AE - Checkbox to start grievance

  // Section 9: Quick Actions (AF)
  QUICK_ACTIONS: 32,               // AF - Checkbox to open Quick Actions dialog

  // Section 10: Member Authentication (AG)
  PIN_HASH: 33,                    // AG - Hashed PIN for member self-service portal

  // Section 11: Employment Details (AH-AJ) - Added for Add Member form parity
  EMPLOYEE_ID: 34,                 // AH - Employee ID (e.g., XX000000)
  DEPARTMENT: 35,                  // AI - Department / work unit category
  HIRE_DATE: 36,                   // AJ - Hire date (date format)

  // Section 12: Mailing Address / PII (AK-AM) - Hidden by default, PII
  STREET_ADDRESS: 37,              // AK - Street address (PII)
  CITY: 38,                        // AL - City (PII)
  STATE: 39,                       // AM - State (PII)

  // ALIASES - For backward compatibility
  LOCATION: 5,                     // Alias for WORK_LOCATION
  DAYS_TO_DEADLINE: 30             // Alias for NEXT_DEADLINE
};

/**
 * Member Directory columns considered PII (Personally Identifiable Information)
 * These columns are hidden by default and MUST NOT appear in any modals or exports
 * unless explicitly authorized. Includes mailing address fields.
 * @const {Array<number>}
 */
var PII_MEMBER_COLS = [37, 38, 39]; // STREET_ADDRESS, CITY, STATE

// ============================================================================
// MEETING CHECK-IN LOG COLUMNS (16 columns: A-P)
// ============================================================================

/**
 * Meeting Check-In Log column positions (1-indexed)
 * Columns A-H are original; I-M added for event scheduling
 * @const {Object}
 */
var MEETING_CHECKIN_COLS = {
  MEETING_ID: 1,         // A - Meeting ID (steward-created)
  MEETING_NAME: 2,       // B - Meeting name/topic
  MEETING_DATE: 3,       // C - Date of meeting
  MEETING_TYPE: 4,       // D - Virtual or In-Person
  MEMBER_ID: 5,          // E - Checked-in member ID
  MEMBER_NAME: 6,        // F - Member first + last name
  CHECKIN_TIME: 7,       // G - Timestamp of check-in
  EMAIL: 8,              // H - Member email (for lookup)
  MEETING_TIME: 9,       // I - Start time (HH:mm)
  MEETING_DURATION: 10,  // J - Duration in hours
  EVENT_STATUS: 11,      // K - Scheduled / Active / Completed
  NOTIFY_STEWARDS: 12,   // L - Steward email(s) for attendance report
  CALENDAR_EVENT_ID: 13, // M - Google Calendar event ID
  NOTES_DOC_URL: 14,     // N - Meeting Notes Google Doc URL
  AGENDA_DOC_URL: 15,    // O - Meeting Agenda Google Doc URL
  AGENDA_STEWARDS: 16    // P - Steward emails for early agenda sharing (3 days prior)
};

/**
 * Meeting event statuses
 * @const {Object}
 */
var MEETING_STATUS = {
  SCHEDULED: 'Scheduled',
  ACTIVE: 'Active',
  COMPLETED: 'Completed'
};

// ============================================================================
// GRIEVANCE LOG COLUMNS (41 columns total: A-AO)
// ============================================================================

/**
 * Grievance Log column positions (1-indexed)
 * CRITICAL: ALL column references must use these constants
 * @const {Object}
 */
var GRIEVANCE_COLS = {
  // Section 1: Identity (A-D)
  GRIEVANCE_ID: 1,        // A - Grievance ID
  MEMBER_ID: 2,           // B - Member ID
  FIRST_NAME: 3,          // C - First Name
  LAST_NAME: 4,           // D - Last Name

  // Section 2: Status & Assignment (E-F)
  STATUS: 5,              // E - Status
  CURRENT_STEP: 6,        // F - Current Step

  // Section 3: Timeline - Filing (G-I)
  INCIDENT_DATE: 7,       // G - Incident Date
  FILING_DEADLINE: 8,     // H - Filing Deadline (21d) (auto-calc)
  DATE_FILED: 9,          // I - Date Filed (Step I)

  // Section 4: Timeline - Step I (J-K)
  STEP1_DUE: 10,          // J - Step I Decision Due (30d) (auto-calc)
  STEP1_RCVD: 11,         // K - Step I Decision Rcvd

  // Section 5: Timeline - Step II (L-O)
  STEP2_APPEAL_DUE: 12,   // L - Step II Appeal Due (10d) (auto-calc)
  STEP2_APPEAL_FILED: 13, // M - Step II Appeal Filed
  STEP2_DUE: 14,          // N - Step II Decision Due (30d) (auto-calc)
  STEP2_RCVD: 15,         // O - Step II Decision Rcvd

  // Section 6: Timeline - Step III (P-R)
  STEP3_APPEAL_DUE: 16,   // P - Step III Appeal Due (30d) (auto-calc)
  STEP3_APPEAL_FILED: 17, // Q - Step III Appeal Filed
  DATE_CLOSED: 18,        // R - Date Closed

  // Section 7: Calculated Metrics (S-U)
  DAYS_OPEN: 19,          // S - Days Open (auto-calc)
  NEXT_ACTION_DUE: 20,    // T - Next Action Due (auto-calc)
  DAYS_TO_DEADLINE: 21,   // U - Days to Deadline (auto-calc)

  // Section 8: Case Details (V-W)
  ARTICLES: 22,           // V - Articles Violated
  ISSUE_CATEGORY: 23,     // W - Issue Category

  // Section 9: Contact & Location (X-Z)
  MEMBER_EMAIL: 24,       // X - Member Email
  LOCATION: 25,           // Y - Work Location (Site)
  STEWARD: 26,            // Z - Assigned Steward (Name)

  // Section 10: Resolution (AA)
  RESOLUTION: 27,         // AA - Resolution Summary

  // Section 11: Coordinator Notifications (AB-AE)
  MESSAGE_ALERT: 28,      // AB - Message Alert checkbox
  COORDINATOR_MESSAGE: 29,// AC - Coordinator's message text
  ACKNOWLEDGED_BY: 30,    // AD - Steward who acknowledged
  ACKNOWLEDGED_DATE: 31,  // AE - When steward acknowledged

  // Section 12: Drive Integration (AF-AG)
  DRIVE_FOLDER_ID: 32,    // AF - Google Drive folder ID
  DRIVE_FOLDER_URL: 33,   // AG - Google Drive folder URL

  // Section 13: Quick Actions (AH)
  QUICK_ACTIONS: 34,      // AH - Checkbox to open Quick Actions dialog

  // Section 14: Action Type & Checklist (AI-AJ)
  ACTION_TYPE: 35,        // AI - Action Type (Grievance, Records Request, etc.)
  CHECKLIST_PROGRESS: 36, // AJ - Checklist Progress (e.g., "5/8" or "62%")

  // Section 15: Reminders (AK-AN) - For scheduling meetings/follow-ups
  REMINDER_1_DATE: 37,    // AK - First reminder date
  REMINDER_1_NOTE: 38,    // AL - First reminder note (e.g., "Schedule Step II meeting")
  REMINDER_2_DATE: 39,    // AM - Second reminder date
  REMINDER_2_NOTE: 40,    // AN - Second reminder note

  // Section 16: Record Tracking (AO)
  LAST_UPDATED: 41        // AO - Last Updated timestamp (auto-set on edit)
};

// ============================================================================
// BACKWARD COMPATIBILITY ALIASES (0-indexed for array access)
// ============================================================================

/**
 * GRIEVANCE_COLUMNS - 0-indexed alias for legacy code
 * Legacy code uses these as array indices (0-indexed)
 * Modern GRIEVANCE_COLS is 1-indexed for getRange() calls
 * @const {Object}
 */
var GRIEVANCE_COLUMNS = {
  // Core fields (0-indexed)
  GRIEVANCE_ID: 0,         // A - Grievance ID
  MEMBER_ID: 1,            // B - Member ID
  FIRST_NAME: 2,           // C - First Name
  LAST_NAME: 3,            // D - Last Name
  MEMBER_NAME: 2,          // Alias - uses FIRST_NAME column
  STATUS: 4,               // E - Status
  CURRENT_STEP: 5,         // F - Current Step

  // Dates and deadlines (0-indexed)
  INCIDENT_DATE: 6,        // G - Incident Date
  FILING_DEADLINE: 7,      // H - Filing Deadline
  FILING_DATE: 8,          // I - Date Filed (alias for DATE_FILED)
  DATE_FILED: 8,           // I - Date Filed

  // Step I (0-indexed)
  STEP1_DUE: 9,            // J - Step I Decision Due
  STEP_1_DUE: 9,           // Alias with underscore
  STEP1_RCVD: 10,          // K - Step I Decision Received
  STEP_1_DATE: 10,         // Alias
  STEP_1_STATUS: 4,        // Uses STATUS column for step status

  // Step II (0-indexed)
  STEP2_APPEAL_DUE: 11,    // L - Step II Appeal Due
  STEP2_APPEAL_FILED: 12,  // M - Step II Appeal Filed
  STEP_2_DATE: 12,         // Alias
  STEP2_DUE: 13,           // N - Step II Decision Due
  STEP_2_DUE: 13,          // Alias with underscore
  STEP2_RCVD: 14,          // O - Step II Decision Received
  STEP_2_STATUS: 4,        // Uses STATUS column

  // Step III (0-indexed)
  STEP3_APPEAL_DUE: 15,    // P - Step III Appeal Due
  STEP3_APPEAL_FILED: 16,  // Q - Step III Appeal Filed
  STEP_3_DATE: 16,         // Alias
  STEP3_DUE: 15,           // Same as appeal due
  STEP_3_DUE: 15,          // Alias with underscore
  STEP_3_STATUS: 4,        // Uses STATUS column
  DATE_CLOSED: 17,         // R - Date Closed
  ARBITRATION_DATE: 17,    // Alias for DATE_CLOSED

  // Calculated metrics (0-indexed)
  DAYS_OPEN: 18,           // S - Days Open
  NEXT_ACTION_DUE: 19,     // T - Next Action Due
  LAST_UPDATED: 40,        // Alias for RECORD_LAST_UPDATED (AO)
  DAYS_TO_DEADLINE: 20,    // U - Days to Deadline

  // Case details (0-indexed)
  ARTICLES: 21,            // V - Articles Violated
  ISSUE_CATEGORY: 22,      // W - Issue Category
  GRIEVANCE_TYPE: 22,      // Alias for ISSUE_CATEGORY
  DESCRIPTION: 22,         // Alias - uses ISSUE_CATEGORY

  // Contact & Location (0-indexed)
  MEMBER_EMAIL: 23,        // X - Member Email
  LOCATION: 24,            // Y - Work Location
  STEWARD: 25,             // Z - Assigned Steward

  // Resolution (0-indexed)
  RESOLUTION: 26,          // AA - Resolution Summary
  OUTCOME: 26,             // Alias for RESOLUTION
  NOTES: 26,               // Alias for RESOLUTION (used for notes)

  // Coordinator (0-indexed)
  MESSAGE_ALERT: 27,       // AB - Message Alert
  COORDINATOR_MESSAGE: 28, // AC - Coordinator Message
  ACKNOWLEDGED_BY: 29,     // AD - Acknowledged By
  ACKNOWLEDGED_DATE: 30,   // AE - Acknowledged Date

  // Drive (0-indexed)
  DRIVE_FOLDER_ID: 31,     // AF - Drive Folder ID
  DRIVE_FOLDER_URL: 32,    // AG - Drive Folder URL
  DRIVE_FOLDER: 32,        // Alias for DRIVE_FOLDER_URL

  // Quick Actions (0-indexed)
  QUICK_ACTIONS: 33,       // AH - Quick Actions

  // Action Type & Checklist (0-indexed)
  ACTION_TYPE: 34,         // AI - Action Type
  CHECKLIST_PROGRESS: 35,  // AJ - Checklist Progress

  // Reminders (0-indexed)
  REMINDER_1_DATE: 36,     // AK - First reminder date
  REMINDER_1_NOTE: 37,     // AL - First reminder note
  REMINDER_2_DATE: 38,     // AM - Second reminder date
  REMINDER_2_NOTE: 39,     // AN - Second reminder note

  // Record Tracking (0-indexed)
  RECORD_LAST_UPDATED: 40  // AO - Last Updated timestamp
};

/**
 * MEMBER_COLUMNS - 0-indexed alias for legacy code
 * Legacy code uses these as array indices (0-indexed)
 * Modern MEMBER_COLS is 1-indexed for getRange() calls
 * @const {Object}
 */
var MEMBER_COLUMNS = {
  // Core fields (0-indexed)
  ID: 0,                   // A - Member ID
  MEMBER_ID: 0,            // Alias
  FIRST_NAME: 1,           // B - First Name
  LAST_NAME: 2,            // C - Last Name
  JOB_TITLE: 3,            // D - Job Title
  JOB_DEPT: 3,             // Legacy alias for JOB_TITLE (DEPARTMENT now at index 34)

  // Location & Work (0-indexed)
  WORK_LOCATION: 4,        // E - Work Location
  LOCATION: 4,             // Alias
  UNIT: 5,                 // F - Unit
  CUBICLE: 6,              // G - Cubicle (hidden)
  OFFICE_DAYS: 7,          // H - Office Days

  // Contact (0-indexed)
  EMAIL: 8,                // I - Email
  PHONE: 9,                // J - Phone
  PREFERRED_COMM: 10,      // K - Preferred Communication
  BEST_TIME: 11,           // L - Best Time to Contact

  // Organization (0-indexed)
  SUPERVISOR: 12,          // M - Supervisor
  MANAGER: 13,             // N - Manager
  IS_STEWARD: 14,          // O - Is Steward
  COMMITTEES: 15,          // P - Committees
  ASSIGNED_STEWARD: 16,    // Q - Assigned Steward

  // Engagement (0-indexed)
  LAST_VIRTUAL_MTG: 17,    // R - Last Virtual Meeting
  LAST_INPERSON_MTG: 18,   // S - Last In-Person Meeting
  OPEN_RATE: 19,           // T - Open Rate
  VOLUNTEER_HOURS: 20,     // U - Volunteer Hours

  // Interests (0-indexed)
  INTEREST_LOCAL: 21,      // V - Interest in Local
  INTEREST_CHAPTER: 22,    // W - Interest in Chapter
  INTEREST_ALLIED: 23,     // X - Interest in Allied

  // Contact Tracking (0-indexed)
  RECENT_CONTACT_DATE: 24, // Y - Recent Contact Date
  LAST_UPDATED: 24,        // Alias for tracking last update
  CONTACT_STEWARD: 25,     // Z - Contact Steward
  CONTACT_NOTES: 26,       // AA - Contact Notes

  // Grievance fields (0-indexed)
  HAS_OPEN_GRIEVANCE: 27,  // AB - Has Open Grievance
  GRIEVANCE_STATUS: 28,    // AC - Grievance Status
  STATUS: 28,              // Alias for member status
  NEXT_DEADLINE: 29,       // AD - Next Deadline
  START_GRIEVANCE: 30,     // AE - Start Grievance

  // Quick Actions (0-indexed)
  QUICK_ACTIONS: 31,       // AF - Quick Actions

  // Member Authentication (0-indexed)
  PIN_HASH: 32,            // AG - PIN Hash

  // Employment Details (0-indexed)
  EMPLOYEE_ID: 33,         // AH - Employee ID
  DEPARTMENT: 34,          // AI - Department
  HIRE_DATE: 35,           // AJ - Hire Date

  // Mailing Address / PII (0-indexed) - Hidden by default
  STREET_ADDRESS: 36,      // AK - Street Address (PII)
  CITY: 37,                // AL - City (PII)
  STATE: 38                // AM - State (PII)
};

// ============================================================================
// CONFIG COLUMN MAPPING
// ============================================================================

/**
 * Config sheet column positions for dropdown sources
 * @const {Object}
 */
var CONFIG_COLS = {
  // ── EMPLOYMENT INFO ── (A-E)
  JOB_TITLES: 1,              // A
  OFFICE_LOCATIONS: 2,        // B
  UNITS: 3,                   // C
  OFFICE_DAYS: 4,             // D
  YES_NO: 5,                  // E

  // ── SUPERVISION ── (F-G)
  SUPERVISORS: 6,             // F
  MANAGERS: 7,                // G

  // ── STEWARD INFO ── (H-I)
  STEWARDS: 8,                // H
  STEWARD_COMMITTEES: 9,      // I

  // ── GRIEVANCE SETTINGS ── (J-M)
  GRIEVANCE_STATUS: 10,       // J
  GRIEVANCE_STEP: 11,         // K
  ISSUE_CATEGORY: 12,         // L
  ARTICLES: 13,               // M

  // ── LINKS & COORDINATORS ── (N-Q)
  COMM_METHODS: 14,           // N
  GRIEVANCE_COORDINATORS: 15, // O
  GRIEVANCE_FORM_URL: 16,     // P
  CONTACT_FORM_URL: 17,       // Q

  // ── NOTIFICATIONS ── (R-S)
  ADMIN_EMAILS: 18,           // R
  ALERT_DAYS: 19,             // S
  NOTIFICATION_RECIPIENTS: 20, // T

  // ── ORGANIZATION ── (U-X)
  ORG_NAME: 21,               // U
  LOCAL_NUMBER: 22,           // V
  MAIN_ADDRESS: 23,           // W
  MAIN_PHONE: 24,             // X

  // ── INTEGRATION ── (Y-Z)
  DRIVE_FOLDER_ID: 25,        // Y
  CALENDAR_ID: 26,            // Z

  // ── DEADLINES ── (AA-AD)
  FILING_DEADLINE_DAYS: 27,   // AA
  STEP1_RESPONSE_DAYS: 28,    // AB
  STEP2_APPEAL_DAYS: 29,      // AC
  STEP2_RESPONSE_DAYS: 30,    // AD

  // ── MULTI-SELECT OPTIONS ── (AE-AF)
  BEST_TIMES: 31,             // AE
  HOME_TOWNS: 32,             // AF

  // ── CONTRACT & LEGAL ── (AG-AJ)
  CONTRACT_GRIEVANCE: 33,     // AG
  CONTRACT_DISCIPLINE: 34,    // AH
  CONTRACT_WORKLOAD: 35,      // AI
  CONTRACT_NAME: 36,          // AJ

  // ── ORG IDENTITY ── (AK-AM)
  UNION_PARENT: 37,           // AK
  STATE_REGION: 38,           // AL
  ORG_WEBSITE: 39,            // AM

  // ── EXTENDED CONTACT ── (AN-AQ)
  OFFICE_ADDRESSES: 40,       // AN
  MAIN_FAX: 41,               // AO
  MAIN_CONTACT_NAME: 42,      // AP
  MAIN_CONTACT_EMAIL: 43,     // AQ

  // ── FORM LINKS ── (AR)
  SATISFACTION_FORM_URL: 44,  // AR - Member Satisfaction Survey form URL

  // ── STRATEGIC COMMAND CENTER ── (AS-AY)
  CHIEF_STEWARD_EMAIL: 45,    // AS - Email for escalation alerts
  UNIT_CODES: 46,             // AT - Unit code prefixes (format: "Unit Name:CODE,Unit2:CODE2")
  ARCHIVE_FOLDER_ID: 47,      // AU - Drive folder ID for archives
  ESCALATION_STATUSES: 48,    // AV - Status values that trigger alerts (comma-separated)
  ESCALATION_STEPS: 49,       // AW - Step values that trigger alerts (comma-separated)
  TEMPLATE_ID: 50,            // AX - Google Doc template ID for grievance PDFs
  PDF_FOLDER_ID: 51,          // AY - Drive folder for generated PDFs (optional, uses ARCHIVE_FOLDER_ID if not set)

  // ── MOBILE DASHBOARD ── (AZ)
  MOBILE_DASHBOARD_URL: 52    // AZ - Mobile Dashboard URL (auto-generated by addMobileDashboardLinkToConfig) - LAST COLUMN
};

// ============================================================================
// SATISFACTION SURVEY COLUMNS (Google Form Response + Summary)
// ============================================================================

/**
 * Member Satisfaction Survey column positions (1-indexed)
 *
 * FORM RESPONSE AREA (A-BQ): Auto-populated by Google Form link
 * - Column A: Timestamp (auto-generated by Google Forms)
 * - Columns B-BQ: 68 question responses
 *
 * SUMMARY/CHART DATA AREA (Column BT onwards): Aggregated metrics for charts
 *
 * @const {Object}
 */
var SATISFACTION_COLS = {
  // ── FORM RESPONSE COLUMNS (Auto-created by Google Form) ──
  TIMESTAMP: 1,                   // A - Auto-generated by Google Forms

  // Work Context (Q1-5)
  Q1_WORKSITE: 2,                 // B
  Q2_ROLE: 3,                     // C
  Q3_SHIFT: 4,                    // D
  Q4_TIME_IN_ROLE: 5,             // E
  Q5_STEWARD_CONTACT: 6,          // F (branching: Yes → 3A, No → 3B)

  // Overall Satisfaction (Q6-9) - Scale 1-10
  Q6_SATISFIED_REP: 7,            // G
  Q7_TRUST_UNION: 8,              // H
  Q8_FEEL_PROTECTED: 9,           // I
  Q9_RECOMMEND: 10,               // J

  // Steward Ratings 3A (Q10-17) - For those with steward contact
  Q10_TIMELY_RESPONSE: 11,        // K
  Q11_TREATED_RESPECT: 12,        // L
  Q12_EXPLAINED_OPTIONS: 13,      // M
  Q13_FOLLOWED_THROUGH: 14,       // N
  Q14_ADVOCATED: 15,              // O
  Q15_SAFE_CONCERNS: 16,          // P
  Q16_CONFIDENTIALITY: 17,        // Q
  Q17_STEWARD_IMPROVE: 18,        // R (paragraph)

  // Steward Access 3B (Q18-20) - For those without steward contact
  Q18_KNOW_CONTACT: 19,           // S
  Q19_CONFIDENT_HELP: 20,         // T
  Q20_EASY_FIND: 21,              // U

  // Chapter Effectiveness (Q21-25)
  Q21_UNDERSTAND_ISSUES: 22,      // V
  Q22_CHAPTER_COMM: 23,           // W
  Q23_ORGANIZES: 24,              // X
  Q24_REACH_CHAPTER: 25,          // Y
  Q25_FAIR_REP: 26,               // Z

  // Local Leadership (Q26-31)
  Q26_DECISIONS_CLEAR: 27,        // AA
  Q27_UNDERSTAND_PROCESS: 28,     // AB
  Q28_TRANSPARENT_FINANCE: 29,    // AC
  Q29_ACCOUNTABLE: 30,            // AD
  Q30_FAIR_PROCESSES: 31,         // AE
  Q31_WELCOMES_OPINIONS: 32,      // AF

  // Contract Enforcement (Q32-36)
  Q32_ENFORCES_CONTRACT: 33,      // AG
  Q33_REALISTIC_TIMELINES: 34,    // AH
  Q34_CLEAR_UPDATES: 35,          // AI
  Q35_FRONTLINE_PRIORITY: 36,     // AJ
  Q36_FILED_GRIEVANCE: 37,        // AK (branching: Yes → 6A, No → 7)

  // Representation Process 6A (Q37-40) - For those who filed grievance
  Q37_UNDERSTOOD_STEPS: 38,       // AL
  Q38_FELT_SUPPORTED: 39,         // AM
  Q39_UPDATES_OFTEN: 40,          // AN
  Q40_OUTCOME_JUSTIFIED: 41,      // AO

  // Communication Quality (Q41-45)
  Q41_CLEAR_ACTIONABLE: 42,       // AP
  Q42_ENOUGH_INFO: 43,            // AQ
  Q43_FIND_EASILY: 44,            // AR
  Q44_ALL_SHIFTS: 45,             // AS
  Q45_MEETINGS_WORTH: 46,         // AT

  // Member Voice & Culture (Q46-50)
  Q46_VOICE_MATTERS: 47,          // AU
  Q47_SEEKS_INPUT: 48,            // AV
  Q48_DIGNITY: 49,                // AW
  Q49_NEWER_SUPPORTED: 50,        // AX
  Q50_CONFLICT_RESPECT: 51,       // AY

  // Value & Collective Action (Q51-55)
  Q51_GOOD_VALUE: 52,             // AZ
  Q52_PRIORITIES_NEEDS: 53,       // BA
  Q53_PREPARED_MOBILIZE: 54,      // BB
  Q54_HOW_INVOLVED: 55,           // BC
  Q55_WIN_TOGETHER: 56,           // BD

  // Scheduling/Office Days (Q56-63)
  Q56_UNDERSTAND_CHANGES: 57,     // BE
  Q57_ADEQUATELY_INFORMED: 58,    // BF
  Q58_CLEAR_CRITERIA: 59,         // BG
  Q59_WORK_EXPECTATIONS: 60,      // BH
  Q60_EFFECTIVE_OUTCOMES: 61,     // BI
  Q61_SUPPORTS_WELLBEING: 62,     // BJ
  Q62_CONCERNS_SERIOUS: 63,       // BK
  Q63_SCHEDULING_CHALLENGE: 64,   // BL (paragraph)

  // Priorities & Close (Q64-68)
  Q64_TOP_PRIORITIES: 65,         // BM (checkboxes - comma separated)
  Q65_ONE_CHANGE: 66,             // BN (paragraph)
  Q66_KEEP_DOING: 67,             // BO (paragraph)
  Q67_ADDITIONAL: 68,             // BP (paragraph)
  Q68_SUBMIT: 69,                 // BQ (if present)

  // ── SUMMARY/CHART DATA AREA (Column BT onwards) ──
  SUMMARY_START: 72,              // BT - Start of summary section

  // Section averages for charts
  AVG_OVERALL_SAT: 72,            // BT - Avg of Q6-Q9
  AVG_STEWARD_RATING: 73,         // BU - Avg of Q10-Q16
  AVG_STEWARD_ACCESS: 74,         // BV - Avg of Q18-Q20
  AVG_CHAPTER: 75,                // BW - Avg of Q21-Q25
  AVG_LEADERSHIP: 76,             // BX - Avg of Q26-Q31
  AVG_CONTRACT: 77,               // BY - Avg of Q32-Q35
  AVG_REPRESENTATION: 78,         // BZ - Avg of Q37-Q40
  AVG_COMMUNICATION: 79,          // CA - Avg of Q41-Q45
  AVG_MEMBER_VOICE: 80,           // CB - Avg of Q46-Q50
  AVG_VALUE_ACTION: 81,           // CC - Avg of Q51-Q55
  AVG_SCHEDULING: 82,             // CD - Avg of Q56-Q62

  // ── VERIFICATION & TRACKING COLUMNS (CE onwards) ──
  EMAIL: 83,                      // CE - Email address from form submission
  VERIFIED: 84,                   // CF - Yes / Pending Review / Rejected
  MATCHED_MEMBER_ID: 85,          // CG - Member ID if email matched
  QUARTER: 86,                    // CH - Quarter string (e.g., "2026-Q1")
  IS_LATEST: 87,                  // CI - Yes/No - Is this the latest for this member this quarter?
  SUPERSEDED_BY: 88,              // CJ - Row number of newer response (if superseded)
  REVIEWER_NOTES: 89              // CK - Notes from reviewer
};

/**
 * Survey section definitions for grouping and analysis
 * @const {Object}
 */
var SATISFACTION_SECTIONS = {
  WORK_CONTEXT: { name: 'Work Context', questions: [2,3,4,5,6], scale: false },
  OVERALL_SAT: { name: 'Overall Satisfaction', questions: [7,8,9,10], scale: true },
  STEWARD_3A: { name: 'Steward Ratings', questions: [11,12,13,14,15,16,17], scale: true },
  STEWARD_3B: { name: 'Steward Access', questions: [19,20,21], scale: true },
  CHAPTER: { name: 'Chapter Effectiveness', questions: [22,23,24,25,26], scale: true },
  LEADERSHIP: { name: 'Local Leadership', questions: [27,28,29,30,31,32], scale: true },
  CONTRACT: { name: 'Contract Enforcement', questions: [33,34,35,36], scale: true },
  REPRESENTATION: { name: 'Representation Process', questions: [38,39,40,41], scale: true },
  COMMUNICATION: { name: 'Communication Quality', questions: [42,43,44,45,46], scale: true },
  MEMBER_VOICE: { name: 'Member Voice & Culture', questions: [47,48,49,50,51], scale: true },
  VALUE_ACTION: { name: 'Value & Collective Action', questions: [52,53,54,55,56], scale: true },
  SCHEDULING: { name: 'Scheduling/Office Days', questions: [57,58,59,60,61,62,63], scale: true },
  PRIORITIES: { name: 'Priorities & Close', questions: [65,66,67,68], scale: false }
};

// ============================================================================
// FEEDBACK & DEVELOPMENT COLUMNS (11 columns: A-K)
// ============================================================================

/**
 * Feedback & Development column positions (1-indexed)
 * @const {Object}
 */
var FEEDBACK_COLS = {
  TIMESTAMP: 1,                // A - Auto-generated timestamp
  SUBMITTED_BY: 2,             // B - Who submitted the feedback
  CATEGORY: 3,                 // C - Area of the system
  TYPE: 4,                     // D - Bug, Feature Request, Improvement
  PRIORITY: 5,                 // E - Low, Medium, High, Critical
  TITLE: 6,                    // F - Short title
  DESCRIPTION: 7,              // G - Detailed description
  STATUS: 8,                   // H - New, In Progress, Resolved, Won't Fix
  ASSIGNED_TO: 9,              // I - Who is working on it
  RESOLUTION: 10,              // J - How it was resolved
  NOTES: 11                    // K - Additional notes
};

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Convert column number to letter notation (e.g., 1 -> A, 27 -> AA)
 * @param {number} columnNumber - Column number (1-indexed)
 * @returns {string} Column letter(s)
 */
function getColumnLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * Convert column letter to number (e.g., A -> 1, AA -> 27)
 * @param {string} columnLetter - Column letter(s)
 * @returns {number} Column number (1-indexed)
 */
function getColumnNumber(columnLetter) {
  var result = 0;
  for (var i = 0; i < columnLetter.length; i++) {
    result = result * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Map a Member Directory row array to a structured object
 * @param {Array} row - Row data array from Member Directory
 * @returns {Object} Structured member object
 */
function mapMemberRow(row) {
  return {
    memberId: row[MEMBER_COLS.MEMBER_ID - 1] || '',
    firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
    lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
    fullName: (row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || ''),
    jobTitle: row[MEMBER_COLS.JOB_TITLE - 1] || '',
    workLocation: row[MEMBER_COLS.WORK_LOCATION - 1] || '',
    unit: row[MEMBER_COLS.UNIT - 1] || '',
    officeDays: row[MEMBER_COLS.OFFICE_DAYS - 1] || '',
    email: row[MEMBER_COLS.EMAIL - 1] || '',
    phone: row[MEMBER_COLS.PHONE - 1] || '',
    preferredComm: row[MEMBER_COLS.PREFERRED_COMM - 1] || '',
    bestTime: row[MEMBER_COLS.BEST_TIME - 1] || '',
    supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || '',
    manager: row[MEMBER_COLS.MANAGER - 1] || '',
    isSteward: row[MEMBER_COLS.IS_STEWARD - 1] || '',
    committees: row[MEMBER_COLS.COMMITTEES - 1] || '',
    assignedSteward: row[MEMBER_COLS.ASSIGNED_STEWARD - 1] || '',
    lastVirtualMtg: row[MEMBER_COLS.LAST_VIRTUAL_MTG - 1] || '',
    lastInPersonMtg: row[MEMBER_COLS.LAST_INPERSON_MTG - 1] || '',
    openRate: row[MEMBER_COLS.OPEN_RATE - 1] || '',
    volunteerHours: row[MEMBER_COLS.VOLUNTEER_HOURS - 1] || '',
    interestLocal: row[MEMBER_COLS.INTEREST_LOCAL - 1] || '',
    interestChapter: row[MEMBER_COLS.INTEREST_CHAPTER - 1] || '',
    interestAllied: row[MEMBER_COLS.INTEREST_ALLIED - 1] || '',
    homeTown: row[MEMBER_COLS.HOME_TOWN - 1] || '',
    recentContactDate: row[MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '',
    contactSteward: row[MEMBER_COLS.CONTACT_STEWARD - 1] || '',
    contactNotes: row[MEMBER_COLS.CONTACT_NOTES - 1] || '',
    hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] || '',
    grievanceStatus: row[MEMBER_COLS.GRIEVANCE_STATUS - 1] || '',
    nextDeadline: row[MEMBER_COLS.NEXT_DEADLINE - 1] || '',
    startGrievance: row[MEMBER_COLS.START_GRIEVANCE - 1] || false
  };
}

/**
 * Map a Grievance Log row array to a structured object
 * @param {Array} row - Row data array from Grievance Log
 * @returns {Object} Structured grievance object
 */
function mapGrievanceRow(row) {
  return {
    grievanceId: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '',
    memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
    firstName: row[GRIEVANCE_COLS.FIRST_NAME - 1] || '',
    lastName: row[GRIEVANCE_COLS.LAST_NAME - 1] || '',
    fullName: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || ''),
    status: row[GRIEVANCE_COLS.STATUS - 1] || '',
    currentStep: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || '',
    incidentDate: row[GRIEVANCE_COLS.INCIDENT_DATE - 1] || '',
    filingDeadline: row[GRIEVANCE_COLS.FILING_DEADLINE - 1] || '',
    dateFiled: row[GRIEVANCE_COLS.DATE_FILED - 1] || '',
    step1Due: row[GRIEVANCE_COLS.STEP1_DUE - 1] || '',
    step1Rcvd: row[GRIEVANCE_COLS.STEP1_RCVD - 1] || '',
    step2AppealDue: row[GRIEVANCE_COLS.STEP2_APPEAL_DUE - 1] || '',
    step2AppealFiled: row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1] || '',
    step2Due: row[GRIEVANCE_COLS.STEP2_DUE - 1] || '',
    step2Rcvd: row[GRIEVANCE_COLS.STEP2_RCVD - 1] || '',
    step3AppealDue: row[GRIEVANCE_COLS.STEP3_APPEAL_DUE - 1] || '',
    step3AppealFiled: row[GRIEVANCE_COLS.STEP3_APPEAL_FILED - 1] || '',
    dateClosed: row[GRIEVANCE_COLS.DATE_CLOSED - 1] || '',
    daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || '',
    nextActionDue: row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] || '',
    daysToDeadline: row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] || '',
    articles: row[GRIEVANCE_COLS.ARTICLES - 1] || '',
    issueCategory: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '',
    memberEmail: row[GRIEVANCE_COLS.MEMBER_EMAIL - 1] || '',
    unit: row[GRIEVANCE_COLS.UNIT - 1] || '',
    location: row[GRIEVANCE_COLS.LOCATION - 1] || '',
    steward: row[GRIEVANCE_COLS.STEWARD - 1] || '',
    resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || '',
    messageAlert: row[GRIEVANCE_COLS.MESSAGE_ALERT - 1] || false,
    coordinatorMessage: row[GRIEVANCE_COLS.COORDINATOR_MESSAGE - 1] || '',
    acknowledgedBy: row[GRIEVANCE_COLS.ACKNOWLEDGED_BY - 1] || '',
    acknowledgedDate: row[GRIEVANCE_COLS.ACKNOWLEDGED_DATE - 1] || '',
    driveFolderId: row[GRIEVANCE_COLS.DRIVE_FOLDER_ID - 1] || '',
    driveFolderUrl: row[GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1] || '',
    actionType: row[GRIEVANCE_COLS.ACTION_TYPE - 1] || 'Grievance',
    checklistProgress: row[GRIEVANCE_COLS.CHECKLIST_PROGRESS - 1] || '',
    reminder1Date: row[GRIEVANCE_COLS.REMINDER_1_DATE - 1] || '',
    reminder1Note: row[GRIEVANCE_COLS.REMINDER_1_NOTE - 1] || '',
    reminder2Date: row[GRIEVANCE_COLS.REMINDER_2_DATE - 1] || '',
    reminder2Note: row[GRIEVANCE_COLS.REMINDER_2_NOTE - 1] || ''
  };
}

/**
 * Get all member header labels in order
 * @returns {Array} Array of header labels for Member Directory
 */
function getMemberHeaders() {
  return [
    'Member ID', 'First Name', 'Last Name', 'Job Title',
    'Work Location', 'Unit', 'Cubicle', 'Office Days',
    'Email', 'Phone', 'Preferred Communication', 'Best Time to Contact',
    'Supervisor', 'Manager', 'Is Steward', 'Committees', 'Assigned Steward',
    'Last Virtual Mtg', 'Last In-Person Mtg', 'Open Rate %', 'Volunteer Hours',
    'Interest: Local', 'Interest: Chapter', 'Interest: Allied',
    'Recent Contact Date', 'Contact Steward', 'Contact Notes',
    'Has Open Grievance?', 'Grievance Status', 'Days to Deadline', 'Start Grievance',
    '⚡ Actions',
    'PIN Hash',
    'Employee ID', 'Department', 'Hire Date',
    'Street Address', 'City', 'State'
  ];
}

/**
 * Get all grievance header labels in order
 * @returns {Array} Array of header labels for Grievance Log
 */
function getGrievanceHeaders() {
  return [
    'Grievance ID', 'Member ID', 'First Name', 'Last Name',
    'Status', 'Current Step',
    'Incident Date', 'Filing Deadline', 'Date Filed',
    'Step I Due', 'Step I Rcvd',
    'Step II Appeal Due', 'Step II Appeal Filed', 'Step II Due', 'Step II Rcvd',
    'Step III Appeal Due', 'Step III Appeal Filed', 'Date Closed',
    'Days Open', 'Next Action Due', 'Days to Deadline',
    'Articles Violated', 'Issue Category',
    'Member Email', 'Work Location', 'Assigned Steward',
    'Resolution',
    'Message Alert', 'Coordinator Message', 'Acknowledged By', 'Acknowledged Date',
    'Drive Folder ID', 'Drive Folder URL',
    '⚡ Actions',
    'Action Type', 'Checklist Progress',
    'Reminder 1 Date', 'Reminder 1 Note', 'Reminder 2 Date', 'Reminder 2 Note',
    'Last Updated'
  ];
}

// ============================================================================
// VALIDATION VALUES
// ============================================================================

/**
 * Default values for Config sheet dropdowns
 */
var DEFAULT_CONFIG = {
  OFFICE_DAYS: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
  YES_NO: ['Yes', 'No'],
  // Status includes both workflow states (Open, Pending, In Arbitration) AND outcomes (Won, Denied, Settled, Withdrawn)
  // This single-column design allows Dashboard metrics to count outcomes directly from STATUS column
  GRIEVANCE_STATUS: ['Open', 'Pending Info', 'Settled', 'Withdrawn', 'Denied', 'Won', 'Appealed', 'In Arbitration', 'Closed'],
  GRIEVANCE_STEP: ['Informal', 'Step I', 'Step II', 'Step III', 'Mediation', 'Arbitration'],
  ISSUE_CATEGORY: ['Discipline', 'Workload', 'Scheduling', 'Pay', 'Benefits', 'Safety', 'Harassment', 'Discrimination', 'Contract Violation', 'Other'],
  ARTICLES: [
    'Art. 1 - Recognition',
    'Art. 2 - Management Rights',
    'Art. 3 - Union Rights',
    'Art. 4 - Dues Deduction',
    'Art. 5 - Non-Discrimination',
    'Art. 6 - Hours of Work',
    'Art. 7 - Overtime',
    'Art. 8 - Compensation',
    'Art. 9 - Benefits',
    'Art. 10 - Leave',
    'Art. 11 - Holidays',
    'Art. 12 - Seniority',
    'Art. 13 - Discipline',
    'Art. 14 - Safety',
    'Art. 15 - Training',
    'Art. 16 - Evaluations',
    'Art. 17 - Layoff',
    'Art. 18 - Vacancies',
    'Art. 19 - Transfers',
    'Art. 20 - Subcontracting',
    'Art. 21 - Personnel Files',
    'Art. 22 - Uniforms',
    'Art. 23 - Grievance Procedure',
    'Art. 24 - Arbitration',
    'Art. 25 - No Strike',
    'Art. 26 - Duration'
  ],
  COMM_METHODS: ['Email', 'Phone', 'Text', 'In Person']
};

/**
 * Grievance status priority order for auto-sorting
 * Lower number = higher priority (appears first in sorted list)
 * Active cases appear first, resolved cases last
 */
var GRIEVANCE_STATUS_PRIORITY = {
  'Open': 1,
  'Pending Info': 2,
  'In Arbitration': 3,
  'Appealed': 4,
  'Settled': 5,
  'Won': 6,
  'Denied': 7,
  'Withdrawn': 8,
  'Closed': 9
};

/**
 * Grievance status constants for programmatic access
 * Use these constants instead of hardcoded strings
 * @const {Object}
 */
var GRIEVANCE_STATUS = {
  OPEN: 'Open',
  PENDING: 'Pending Info',
  PENDING_INFO: 'Pending Info',
  SETTLED: 'Settled',
  WITHDRAWN: 'Withdrawn',
  DENIED: 'Denied',
  WON: 'Won',
  APPEALED: 'Appealed',
  IN_ARBITRATION: 'In Arbitration',
  AT_ARBITRATION: 'In Arbitration',
  CLOSED: 'Closed',
  RESOLVED: 'Settled'  // Alias for backward compatibility
};

/**
 * Grievance outcome constants for programmatic access
 * Use these constants instead of hardcoded strings
 * @const {Object}
 */
var GRIEVANCE_OUTCOMES = {
  PENDING: 'Pending',
  WON: 'Won',
  DENIED: 'Denied',
  SETTLED: 'Settled',
  WITHDRAWN: 'Withdrawn',
  CLOSED: 'Closed'
};

/**
 * Default deadline rules for grievance step calculations (fallback values).
 * Actual values are loaded from Config sheet at runtime via getDeadlineRules().
 * @const {Object}
 */
var DEADLINE_DEFAULTS = {
  FILING_DAYS: 21,
  STEP_1_RESPONSE: 7,
  STEP_2_APPEAL: 7,
  STEP_2_RESPONSE: 14,
  STEP_3_APPEAL: 10,
  STEP_3_RESPONSE: 21,
  ARBITRATION_DEMAND: 30,
  WARNING_THRESHOLD: 5,
  CRITICAL_THRESHOLD: 2,
  REMINDER_FIRST: 7,
  REMINDER_SECOND: 3,
  REMINDER_FINAL: 1
};

/**
 * Reads deadline rules from Config sheet (cols AA-AD), falling back to DEADLINE_DEFAULTS.
 * This is the single source of truth for all deadline calculations.
 * @returns {Object} DEADLINE_RULES-compatible object
 */
function getDeadlineRules() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var filing = Number(configSheet.getRange(3, CONFIG_COLS.FILING_DEADLINE_DAYS).getValue());
      var s1Resp = Number(configSheet.getRange(3, CONFIG_COLS.STEP1_RESPONSE_DAYS).getValue());
      var s2Appeal = Number(configSheet.getRange(3, CONFIG_COLS.STEP2_APPEAL_DAYS).getValue());
      var s2Resp = Number(configSheet.getRange(3, CONFIG_COLS.STEP2_RESPONSE_DAYS).getValue());
      return {
        FILING_DAYS: filing || DEADLINE_DEFAULTS.FILING_DAYS,
        STEP_1: { DAYS_FOR_RESPONSE: s1Resp || DEADLINE_DEFAULTS.STEP_1_RESPONSE },
        STEP_2: { DAYS_TO_APPEAL: s2Appeal || DEADLINE_DEFAULTS.STEP_2_APPEAL, DAYS_FOR_RESPONSE: s2Resp || DEADLINE_DEFAULTS.STEP_2_RESPONSE },
        STEP_3: { DAYS_TO_APPEAL: DEADLINE_DEFAULTS.STEP_3_APPEAL, DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_3_RESPONSE },
        ARBITRATION: { DAYS_TO_DEMAND: DEADLINE_DEFAULTS.ARBITRATION_DEMAND }
      };
    }
  } catch (e) {
    console.log('Error reading deadline config: ' + e.message);
  }
  // Fallback to defaults
  return {
    FILING_DAYS: DEADLINE_DEFAULTS.FILING_DAYS,
    STEP_1: { DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_1_RESPONSE },
    STEP_2: { DAYS_TO_APPEAL: DEADLINE_DEFAULTS.STEP_2_APPEAL, DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_2_RESPONSE },
    STEP_3: { DAYS_TO_APPEAL: DEADLINE_DEFAULTS.STEP_3_APPEAL, DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_3_RESPONSE },
    ARBITRATION: { DAYS_TO_DEMAND: DEADLINE_DEFAULTS.ARBITRATION_DEMAND }
  };
}

/** @deprecated Use getDeadlineRules() instead. Kept for backward compatibility. */
var DEADLINE_RULES = {
  STEP_1: { DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_1_RESPONSE },
  STEP_2: { DAYS_TO_APPEAL: DEADLINE_DEFAULTS.STEP_2_APPEAL, DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_2_RESPONSE },
  STEP_3: { DAYS_TO_APPEAL: DEADLINE_DEFAULTS.STEP_3_APPEAL, DAYS_FOR_RESPONSE: DEADLINE_DEFAULTS.STEP_3_RESPONSE },
  ARBITRATION: { DAYS_TO_DEMAND: DEADLINE_DEFAULTS.ARBITRATION_DEMAND }
};

/**
 * Dashboard sheet layout configuration.
 * All cell references for the Dashboard sheet are centralized here.
 * If the dashboard layout changes, update these values in one place.
 * @const {Object}
 */
var DASHBOARD_LAYOUT = {
  // ── DATA ROWS ──
  QUICK_STATS_ROW: 6,           // A6:F6 - Total Members, Stewards, Grievances, Win Rate, Overdue, Due
  MEMBER_METRICS_ROW: 11,       // A11:D11 - Members, Stewards, Avg Open Rate, Vol Hours
  GRIEVANCE_METRICS_ROW: 16,    // A16:F16 - Open, Pending, Settled, Won, Denied, Withdrawn
  TIMELINE_METRICS_ROW: 21,     // A21:D21 - Avg Days Open, Filed, Closed, Avg Resolution
  CATEGORY_START_ROW: 26,       // A26:F30 - Issue category breakdown (5 rows)
  CATEGORY_END_ROW: 30,
  LOCATION_START_ROW: 35,       // A35:F39 - Location breakdown (5 rows)
  LOCATION_END_ROW: 39,
  TREND_START_ROW: 44,          // A44:F46 - Month-over-month trends (3 rows)
  TREND_END_ROW: 46,
  STEWARD_SUMMARY_ROW: 54,      // A54:F54 - Steward summary
  BUSIEST_START_ROW: 59,        // A59:F88 - Top 30 busiest stewards
  BUSIEST_END_ROW: 88,
  TOP_PERFORMERS_START_ROW: 93, // A93:F102 - Top 10 performers
  TOP_PERFORMERS_END_ROW: 102,
  NEEDING_SUPPORT_START_ROW: 107, // A107:F116 - Needing support
  NEEDING_SUPPORT_END_ROW: 116,

  // ── CHART AREA ──
  CHART_INPUT_CELL: 'G120',     // Chart number selection input
  CHART_DISPLAY_ROW: 135,       // Row where chart text displays start
  CHART_DISPLAY_RANGE: 'A135:G145', // Merge range for chart text displays

  // ── COLUMN COUNTS ──
  DATA_COLS: 6,                 // Columns A-F for most data sections
  SPARKLINE_COL: 7              // Column G for sparkline formulas
};

/**
 * Audit event types for logging
 * @const {Object}
 */
var AUDIT_EVENTS = {
  // Grievance events
  GRIEVANCE_CREATED: 'GRIEVANCE_CREATED',
  GRIEVANCE_UPDATED: 'GRIEVANCE_UPDATED',
  GRIEVANCE_STEP_ADVANCED: 'GRIEVANCE_STEP_ADVANCED',
  GRIEVANCE_RESOLVED: 'GRIEVANCE_RESOLVED',
  GRIEVANCE_CLOSED: 'GRIEVANCE_CLOSED',

  // Member events
  MEMBER_ADDED: 'MEMBER_ADDED',
  MEMBER_UPDATED: 'MEMBER_UPDATED',
  MEMBER_DELETED: 'MEMBER_DELETED',

  // System events
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  SYSTEM_INITIALIZED: 'SYSTEM_INITIALIZED',
  DASHBOARD_INITIALIZED: 'DASHBOARD_INITIALIZED',

  // Integration events
  FOLDER_CREATED: 'FOLDER_CREATED',
  CALENDAR_SYNCED: 'CALENDAR_SYNCED',
  EMAIL_SENT: 'EMAIL_SENT',
  PDF_GENERATED: 'PDF_GENERATED',

  // Settings events
  SETTINGS_CHANGED: 'SETTINGS_CHANGED',
  TRIGGER_INSTALLED: 'TRIGGER_INSTALLED',
  TRIGGER_REMOVED: 'TRIGGER_REMOVED'
};

/**
 * Batch processing limits for performance optimization
 * @const {Object}
 */
var BATCH_LIMITS = {
  MAX_ROWS_PER_BATCH: 100,           // Max rows to process in one batch
  MAX_EXECUTION_TIME_MS: 300000,      // 5 minutes max execution time
  PAUSE_BETWEEN_BATCHES_MS: 100,      // Pause between batches to avoid quota limits
  MAX_PARALLEL_OPERATIONS: 10,        // Max concurrent operations
  MAX_API_CALLS_PER_BATCH: 50,        // Max API calls before pausing (Drive, Calendar, etc.)
  CACHE_EXPIRATION_SECONDS: 21600     // 6 hours cache expiration
};

// ============================================================================
// JOB METADATA FIELDS - Maps Member Directory fields to Config dropdown sources
// ============================================================================

/**
 * Job metadata field configuration
 * Maps each Member Directory field to its corresponding Config sheet column
 * @const {Array<Object>}
 */
var JOB_METADATA_FIELDS = [
  { label: 'Job Title', memberCol: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES, configName: 'Job Titles' },
  { label: 'Work Location', memberCol: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS, configName: 'Office Locations' },
  { label: 'Unit', memberCol: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS, configName: 'Units' },
  { label: 'Supervisor', memberCol: MEMBER_COLS.SUPERVISOR, configCol: CONFIG_COLS.SUPERVISORS, configName: 'Supervisors' },
  { label: 'Manager', memberCol: MEMBER_COLS.MANAGER, configCol: CONFIG_COLS.MANAGERS, configName: 'Managers' },
  { label: 'Assigned Steward', memberCol: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, configName: 'Stewards' },
  { label: 'Committees', memberCol: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, configName: 'Steward Committees' },
  { label: 'Home Town', memberCol: MEMBER_COLS.HOME_TOWN, configCol: CONFIG_COLS.HOME_TOWNS, configName: 'Home Towns' }
];

/**
 * Get job metadata field config by label
 * @param {string} label - The field label (e.g., 'Job Title')
 * @returns {Object|null} Field config if found, null otherwise
 */
function getJobMetadataField(label) {
  for (var i = 0; i < JOB_METADATA_FIELDS.length; i++) {
    if (JOB_METADATA_FIELDS[i].label === label) {
      return JOB_METADATA_FIELDS[i];
    }
  }
  return null;
}

/**
 * Get job metadata field config by member column number
 * @param {number} memberCol - The member column number
 * @returns {Object|null} Field config if found, null otherwise
 */
function getJobMetadataByMemberCol(memberCol) {
  for (var i = 0; i < JOB_METADATA_FIELDS.length; i++) {
    if (JOB_METADATA_FIELDS[i].memberCol === memberCol) {
      return JOB_METADATA_FIELDS[i];
    }
  }
  return null;
}

// ============================================================================
// MULTI-SELECT COLUMN CONFIGURATION
// ============================================================================

/**
 * Columns that support multiple selections (comma-separated values)
 * Maps column number to config source column for options
 */
var MULTI_SELECT_COLS = {
  // Member Directory multi-select columns
  MEMBER_DIR: [
    { col: MEMBER_COLS.OFFICE_DAYS, configCol: CONFIG_COLS.OFFICE_DAYS, label: 'Office Days' },
    { col: MEMBER_COLS.PREFERRED_COMM, configCol: CONFIG_COLS.COMM_METHODS, label: 'Preferred Communication' },
    { col: MEMBER_COLS.BEST_TIME, configCol: CONFIG_COLS.BEST_TIMES, label: 'Best Time to Contact' },
    { col: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, label: 'Committees' },
    { col: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, label: 'Assigned Steward(s)' }
  ]
};

/**
 * Check if a column in Member Directory is a multi-select column
 * @param {number} col - Column number (1-indexed)
 * @returns {Object|null} Multi-select config if found, null otherwise
 */
function getMultiSelectConfig(col) {
  for (var i = 0; i < MULTI_SELECT_COLS.MEMBER_DIR.length; i++) {
    if (MULTI_SELECT_COLS.MEMBER_DIR[i].col === col) {
      return MULTI_SELECT_COLS.MEMBER_DIR[i];
    }
  }
  return null;
}

// ============================================================================
// ID GENERATION
// ============================================================================

/**
 * Generate a name-based ID with prefix and 3 random digits
 * Format: Prefix + First 2 chars of firstName + First 2 chars of lastName + 3 random digits
 * Example: M + John Smith → MJOSM123, G + John Smith → GJOSM456
 * @param {string} prefix - ID prefix ('M' for members, 'G' for grievances)
 * @param {string} firstName - First name
 * @param {string} lastName - Last name
 * @param {Object} existingIds - Object with existing IDs as keys (for collision detection)
 * @returns {string} Generated ID (uppercase)
 */
function generateNameBasedId(prefix, firstName, lastName, existingIds) {
  var firstPart = (firstName || 'XX').substring(0, 2).toUpperCase();
  var lastPart = (lastName || 'XX').substring(0, 2).toUpperCase();
  var namePrefix = (prefix || '') + firstPart + lastPart;

  var maxAttempts = 100;
  for (var attempt = 0; attempt < maxAttempts; attempt++) {
    var randomDigits = String(Math.floor(Math.random() * 1000)).padStart(3, '0');
    var newId = namePrefix + randomDigits;

    if (!existingIds || !existingIds[newId]) {
      return newId;
    }
  }

  // Fallback: use UUID-style ID for guaranteed uniqueness in high-volume environments
  // Format: PREFIX + 8 hex chars from UUID
  var uuid = generateUUID_();
  return namePrefix + uuid.substring(0, 8).toUpperCase();
}

/**
 * Generate a UUID v4 for guaranteed uniqueness
 * Used as fallback when name-based ID collisions occur
 * @returns {string} UUID string
 * @private
 */
function generateUUID_() {
  var chars = '0123456789abcdef';
  var uuid = '';
  for (var i = 0; i < 36; i++) {
    if (i === 8 || i === 13 || i === 18 || i === 23) {
      uuid += '-';
    } else if (i === 14) {
      uuid += '4'; // UUID version 4
    } else if (i === 19) {
      uuid += chars.charAt((Math.random() * 4) | 8); // Variant bits
    } else {
      uuid += chars.charAt(Math.floor(Math.random() * 16));
    }
  }
  return uuid;
}

// ============================================================================
// ACTION TYPE CONFIGURATION (Grievances + Other Actions)
// ============================================================================

/**
 * Action types for tracking different case types
 * Allows the system to handle grievances AND other member advocacy actions
 * @const {Object}
 */
var ACTION_TYPES = {
  GRIEVANCE: 'Grievance',
  RECORDS_REQUEST: 'Records Request',
  INFO_REQUEST: 'Information Request',
  WEINGARTEN: 'Weingarten',
  ULP: 'ULP Filing',
  EEOC_MCAD: 'EEOC/MCAD',
  ACCOMMODATION: 'Accommodation',
  OTHER_ADMIN: 'Other Admin'
};

/**
 * Action type display names and configuration
 * @const {Array}
 */
var ACTION_TYPE_CONFIG = [
  { value: 'Grievance', label: 'Grievance', icon: '📋', usesGrievanceSteps: true, color: '#F97316' },
  { value: 'Records Request', label: 'Records Request (Art. 21)', icon: '📁', usesGrievanceSteps: false, color: '#3B82F6' },
  { value: 'Information Request', label: 'Union Information Request', icon: '📄', usesGrievanceSteps: false, color: '#8B5CF6' },
  { value: 'Weingarten', label: 'Weingarten Documentation', icon: '🛡️', usesGrievanceSteps: false, color: '#10B981' },
  { value: 'ULP Filing', label: 'Unfair Labor Practice (DLR)', icon: '⚖️', usesGrievanceSteps: false, color: '#EF4444' },
  { value: 'EEOC/MCAD', label: 'EEOC/MCAD Complaint', icon: '🏛️', usesGrievanceSteps: false, color: '#EC4899' },
  { value: 'Accommodation', label: 'ADA/Reasonable Accommodation', icon: '♿', usesGrievanceSteps: false, color: '#06B6D4' },
  { value: 'Other Admin', label: 'Other Administrative Action', icon: '📝', usesGrievanceSteps: false, color: '#64748B' }
];

/**
 * Get action type configuration by value
 * @param {string} actionType - The action type value
 * @returns {Object|null} Action type config if found
 */
function getActionTypeConfig(actionType) {
  for (var i = 0; i < ACTION_TYPE_CONFIG.length; i++) {
    if (ACTION_TYPE_CONFIG[i].value === actionType) {
      return ACTION_TYPE_CONFIG[i];
    }
  }
  return null;
}

// ============================================================================
// CHECKLIST CONFIGURATION
// ============================================================================

/**
 * Checklist sheet name
 * @const {string}
 */
var CHECKLIST_SHEET_NAME = 'Case Checklist';

/**
 * Checklist column positions (1-indexed)
 * @const {Object}
 */
var CHECKLIST_COLS = {
  CHECKLIST_ID: 1,       // A - Unique checklist item ID
  CASE_ID: 2,            // B - Links to Grievance ID or Action ID
  ACTION_TYPE: 3,        // C - Type of case (Grievance, Records Request, etc.)
  ITEM_TEXT: 4,          // D - Description of checklist item
  CATEGORY: 5,           // E - Document, Meeting, Deadline, Evidence, Communication
  REQUIRED: 6,           // F - Yes/No - Is this item required?
  COMPLETED: 7,          // G - Checkbox - Has item been completed?
  COMPLETED_BY: 8,       // H - Who completed this item
  COMPLETED_DATE: 9,     // I - When item was completed
  DUE_DATE: 10,          // J - Optional due date for time-sensitive items
  NOTES: 11,             // K - Additional notes
  SORT_ORDER: 12         // L - For custom ordering of items
};

/**
 * Checklist item categories
 * @const {Array}
 */
var CHECKLIST_CATEGORIES = [
  'Document',
  'Meeting',
  'Deadline',
  'Evidence',
  'Communication',
  'Follow-up',
  'Other'
];

/**
 * Default checklist templates by action type and issue category
 * Each template contains items that are auto-populated when a new case is created
 * @const {Object}
 */
var CHECKLIST_TEMPLATES = {
  // Standard Grievance Templates
  'Grievance': {
    '_default': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles', category: 'Document', required: true },
      { text: 'Written statement from member', category: 'Evidence', required: true },
      { text: 'Witness statements (if applicable)', category: 'Evidence', required: false },
      { text: 'Relevant emails/communications', category: 'Evidence', required: false },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Discipline': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of discipline letter/notice', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles (Art. 13)', category: 'Document', required: true },
      { text: 'Member\'s written response to discipline', category: 'Document', required: true },
      { text: 'Prior discipline history obtained', category: 'Evidence', required: true },
      { text: 'Comparator cases researched', category: 'Evidence', required: false },
      { text: 'Weingarten documentation (if applicable)', category: 'Document', required: false },
      { text: 'Witness statements', category: 'Evidence', required: false },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Workload': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles', category: 'Document', required: true },
      { text: 'Workload documentation (caseload reports, etc.)', category: 'Evidence', required: true },
      { text: 'Written statement describing workload issues', category: 'Evidence', required: true },
      { text: 'Time/task analysis if available', category: 'Evidence', required: false },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Scheduling': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles (Art. 6)', category: 'Document', required: true },
      { text: 'Copy of schedule showing violation', category: 'Evidence', required: true },
      { text: 'Written statement from member', category: 'Evidence', required: true },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Pay': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles (Art. 8)', category: 'Document', required: true },
      { text: 'Pay stubs showing discrepancy', category: 'Evidence', required: true },
      { text: 'Written statement from member', category: 'Evidence', required: true },
      { text: 'Calculation of amount owed', category: 'Document', required: true },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Harassment': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles', category: 'Document', required: true },
      { text: 'Detailed written statement from member', category: 'Evidence', required: true },
      { text: 'Timeline of incidents', category: 'Evidence', required: true },
      { text: 'Witness statements', category: 'Evidence', required: false },
      { text: 'Copies of any written communications', category: 'Evidence', required: false },
      { text: 'Prior complaints documented', category: 'Evidence', required: false },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ],
    'Discrimination': [
      { text: 'Signed grievance form from member', category: 'Document', required: true },
      { text: 'Copy of relevant contract articles (Art. 5)', category: 'Document', required: true },
      { text: 'Detailed written statement from member', category: 'Evidence', required: true },
      { text: 'Comparator evidence (how others treated)', category: 'Evidence', required: true },
      { text: 'Timeline of incidents', category: 'Evidence', required: true },
      { text: 'Witness statements', category: 'Evidence', required: false },
      { text: 'EEOC/MCAD filing discussed with member', category: 'Communication', required: false },
      { text: 'Step I meeting scheduled', category: 'Meeting', required: true },
      { text: 'Management response received', category: 'Document', required: true },
      { text: 'Member notified of decision', category: 'Communication', required: true }
    ]
  },

  // Records Request Templates
  'Records Request': {
    '_default': [
      { text: 'Written request submitted to HR', category: 'Document', required: true },
      { text: 'Member signed authorization form', category: 'Document', required: true },
      { text: 'Request receipt confirmation', category: 'Document', required: false },
      { text: 'Copy of records received', category: 'Document', required: true },
      { text: 'Records reviewed with member', category: 'Communication', required: true },
      { text: 'Member signed acknowledgment of receipt', category: 'Document', required: false }
    ]
  },

  // Information Request Templates
  'Information Request': {
    '_default': [
      { text: 'Written information request drafted', category: 'Document', required: true },
      { text: 'Request submitted to employer', category: 'Document', required: true },
      { text: 'Submission receipt confirmation', category: 'Document', required: false },
      { text: 'Response received from employer', category: 'Document', required: true },
      { text: 'Information reviewed and analyzed', category: 'Follow-up', required: true }
    ]
  },

  // Weingarten Templates
  'Weingarten': {
    '_default': [
      { text: 'Member requested representation', category: 'Document', required: true },
      { text: 'Date/time of meeting documented', category: 'Document', required: true },
      { text: 'Meeting attendees documented', category: 'Document', required: true },
      { text: 'Notes from meeting', category: 'Document', required: true },
      { text: 'Member debriefed after meeting', category: 'Communication', required: true },
      { text: 'Follow-up actions identified', category: 'Follow-up', required: false }
    ]
  },

  // ULP Filing Templates
  'ULP Filing': {
    '_default': [
      { text: 'ULP charge form completed', category: 'Document', required: true },
      { text: 'Supporting documentation gathered', category: 'Evidence', required: true },
      { text: 'Charge filed with DLR/NLRB', category: 'Document', required: true },
      { text: 'Filing confirmation received', category: 'Document', required: true },
      { text: 'Member notified of filing', category: 'Communication', required: true },
      { text: 'Case number assigned', category: 'Document', required: true },
      { text: 'Investigation meeting scheduled', category: 'Meeting', required: false },
      { text: 'Settlement discussions (if any)', category: 'Meeting', required: false }
    ]
  },

  // EEOC/MCAD Templates
  'EEOC/MCAD': {
    '_default': [
      { text: 'Intake questionnaire completed', category: 'Document', required: true },
      { text: 'Supporting documentation gathered', category: 'Evidence', required: true },
      { text: 'Charge filed with EEOC/MCAD', category: 'Document', required: true },
      { text: 'Right to Sue letter requested (if applicable)', category: 'Document', required: false },
      { text: 'Filing confirmation received', category: 'Document', required: true },
      { text: 'Member notified of filing', category: 'Communication', required: true },
      { text: 'Investigation meeting scheduled', category: 'Meeting', required: false },
      { text: 'Mediation scheduled (if offered)', category: 'Meeting', required: false }
    ]
  },

  // Accommodation Templates
  'Accommodation': {
    '_default': [
      { text: 'Accommodation request form completed', category: 'Document', required: true },
      { text: 'Medical documentation obtained', category: 'Document', required: true },
      { text: 'Request submitted to employer', category: 'Document', required: true },
      { text: 'Interactive process meeting scheduled', category: 'Meeting', required: true },
      { text: 'Meeting notes documented', category: 'Document', required: true },
      { text: 'Employer response received', category: 'Document', required: true },
      { text: 'Accommodation implemented/denied', category: 'Follow-up', required: true },
      { text: 'Member notified of outcome', category: 'Communication', required: true }
    ]
  },

  // Other Admin Templates
  'Other Admin': {
    '_default': [
      { text: 'Issue documented', category: 'Document', required: true },
      { text: 'Action request submitted', category: 'Document', required: true },
      { text: 'Response received', category: 'Document', required: false },
      { text: 'Member notified of outcome', category: 'Communication', required: true }
    ]
  }
};

/**
 * Get checklist template for an action type and issue category
 * @param {string} actionType - The action type (e.g., 'Grievance', 'Records Request')
 * @param {string} issueCategory - The issue category (e.g., 'Discipline', 'Pay') - only for grievances
 * @returns {Array} Array of checklist item templates
 */
function getChecklistTemplate(actionType, issueCategory) {
  var templates = CHECKLIST_TEMPLATES[actionType];
  if (!templates) {
    return CHECKLIST_TEMPLATES['Other Admin']['_default'];
  }

  // For grievances, check for category-specific template
  if (actionType === 'Grievance' && issueCategory && templates[issueCategory]) {
    return templates[issueCategory];
  }

  // Return default template for this action type
  return templates['_default'] || [];
}

/**
 * Generate a sequential ID for high-volume environments
 * Alternative to name-based IDs when strict uniqueness is required
 * @param {string} prefix - ID prefix ('M' for members, 'G' for grievances)
 * @param {Sheet} sheet - Sheet to check for existing IDs
 * @param {number} idColumn - Column number containing IDs (1-indexed)
 * @returns {string} Sequential ID (e.g., M00001, G00042)
 */
function generateSequentialId(prefix, sheet, idColumn) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return prefix + '00001';
  }

  var ids = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
  var maxNum = 0;

  for (var i = 0; i < ids.length; i++) {
    var id = ids[i][0];
    if (id && typeof id === 'string' && id.indexOf(prefix) === 0) {
      var numPart = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  }

  return prefix + String(maxNum + 1).padStart(5, '0');
}

// ============================================================================
// MOBILE OPTIMIZATION UTILITIES
// ============================================================================

/**
 * Returns the viewport meta tag for mobile-optimized modals.
 * Should be included in the <head> of every modal dialog HTML.
 * @returns {string} HTML meta tag string
 */
function getMobileMetaTag() {
  return '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">';
}

/**
 * Returns responsive CSS that should be injected into every modal dialog.
 * Handles common mobile patterns: touch targets, responsive grids,
 * fluid typography, overflow scrolling, and safe-area insets.
 * @returns {string} CSS style block string
 */
function getMobileResponsiveStyles() {
  return '<style id="mobile-responsive-styles">' +
    '/* Mobile-first responsive base */' +
    'html{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}' +
    'body{overflow-x:hidden;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;' +
    'padding-left:env(safe-area-inset-left,0px);padding-right:env(safe-area-inset-right,0px);}' +

    '/* Touch-friendly interactive elements */' +
    'button,select,.btn,.action-btn,.tab,.filter,.quick-item,.option-item,.result-item,.search-tab,.filter-select{' +
      'min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';' +
      'touch-action:manipulation;' +
      '-webkit-tap-highlight-color:transparent;' +
    '}' +
    'input[type="text"],input[type="search"],input[type="email"],input[type="number"],input[type="date"],input[type="password"],textarea,select{' +
      'font-size:16px;' +  /* Prevents iOS zoom on focus */
      'min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';' +
      'touch-action:manipulation;' +
    '}' +
    'input[type="checkbox"],input[type="radio"]{' +
      'width:22px;height:22px;' +
      'touch-action:manipulation;' +
    '}' +

    '/* Smooth scrolling for overflow containers */' +
    '[style*="overflow-y:auto"],[style*="overflow-y: auto"],.options-container,.results-container,.quick-results,.items-list,.preview,.steward-list,.filter-panel,.results-panel{' +
      '-webkit-overflow-scrolling:touch;' +
      'overscroll-behavior:contain;' +
      'scroll-behavior:smooth;' +
    '}' +

    '/* Active state feedback for touch */' +
    'button:active,.btn:active,.action-btn:active,.tab:active,.option-item:active{' +
      'transform:scale(0.97);' +
      'opacity:0.85;' +
      'transition:transform 0.1s,opacity 0.1s;' +
    '}' +

    '/* Mobile breakpoint: phones */' +
    '@media(max-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px){' +
      'body{padding:12px !important;font-size:clamp(13px,3.5vw,15px) !important;}' +
      'h1,h2,.modal-title{font-size:clamp(18px,5vw,24px) !important;}' +
      'h3{font-size:clamp(15px,4vw,18px) !important;}' +
      '.btn,.action-btn,button{' +
        'width:100%;margin-bottom:8px;padding:12px 16px !important;' +
        'font-size:clamp(13px,3.5vw,15px) !important;' +
      '}' +
      '.btn-row,.button-row{' +
        'flex-direction:column !important;gap:8px !important;' +
      '}' +
      'table,.data-table{display:block;overflow-x:auto;-webkit-overflow-scrolling:touch;white-space:nowrap;font-size:12px;}' +
      'th,td{padding:8px 6px !important;font-size:clamp(11px,2.8vw,13px) !important;}' +
      '.stats-row,[class*="stat"]{' +
        'display:grid !important;grid-template-columns:repeat(2,1fr) !important;gap:8px !important;' +
      '}' +
      '[class*="kpi-grid"],[class*="charts-grid"]{' +
        'grid-template-columns:1fr !important;' +
      '}' +
      '.filter-panel{display:none;}' +
      '.filter-toggle{display:block !important;}' +
      '.filter-panel.show{display:block !important;position:fixed;top:0;left:0;right:0;bottom:0;z-index:200;background:#fff;padding:16px;overflow-y:auto;}' +
      'select,input[type="text"],input[type="search"],input[type="email"],textarea{width:100% !important;}' +
      '.filter-group{flex-direction:column !important;}' +
      '.filter-group .action-btn,.filter-group button{width:100% !important;}' +
      '.chart-container{padding:10px !important;}' +
      '.bar-label{width:80px !important;font-size:11px !important;}' +
      '.bar-value{width:40px !important;font-size:11px !important;}' +
      '.gauge-container{flex-direction:column !important;align-items:center !important;}' +
      '.detail-grid{grid-template-columns:repeat(2,1fr) !important;}' +
      '.heatmap-grid{grid-template-columns:repeat(auto-fit,minmax(60px,1fr)) !important;gap:6px !important;}' +
      '.fab{bottom:calc(20px + env(safe-area-inset-bottom,0px)) !important;right:calc(20px + env(safe-area-inset-right,0px)) !important;}' +
    '}' +

    '/* Tablet breakpoint */' +
    '@media(min-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px) and (max-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
      '.stats-row,[class*="stat"]{' +
        'grid-template-columns:repeat(2,1fr) !important;' +
      '}' +
      '[class*="kpi-grid"]{grid-template-columns:repeat(2,1fr) !important;}' +
      'table{font-size:13px;}' +
    '}' +

    '/* Desktop: restore inline button layout */' +
    '@media(min-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px){' +
      '.btn-row,.button-row{' +
        'flex-direction:row !important;' +
      '}' +
      '.btn,.action-btn,button{' +
        'width:auto;margin-bottom:0;' +
      '}' +
    '}' +
  '</style>';
}

/**
 * Returns JavaScript for Android haptic feedback using the Vibration API.
 * Detects Android devices and provides haptic feedback on user interactions.
 * Automatically attaches to buttons, inputs, toggles, and interactive elements.
 * @returns {string} Script block string
 */
function getHapticFeedbackScript() {
  return '<script id="haptic-feedback">' +
    '(function(){' +
      '/* Detect Android device */' +
      'var isAndroid=/android/i.test(navigator.userAgent||"");' +
      'var canVibrate=!!navigator.vibrate;' +
      'if(!isAndroid||!canVibrate)return;' +

      '/* Haptic feedback patterns (duration in ms) */' +
      'var HapticPatterns={' +
        'light:10,' +      /* Light tap - toggles, checkboxes, minor selections */
        'medium:18,' +     /* Medium tap - button presses, tab switches */
        'heavy:30,' +      /* Heavy tap - important actions (submit, delete) */
        'success:[10,60,10],' +  /* Double pulse - success confirmation */
        'error:[40,30,40,30,40],' +  /* Triple pulse - error feedback */
        'warning:[20,40,20]' +  /* Double pulse - warning actions */
      '};' +

      '/* Main haptic trigger function */' +
      'window.hapticFeedback=function(pattern){' +
        'if(!canVibrate)return;' +
        'try{' +
          'var p=HapticPatterns[pattern]||HapticPatterns.light;' +
          'navigator.vibrate(p);' +
        '}catch(e){}' +
      '};' +

      '/* Auto-attach haptic feedback to interactive elements */' +
      'function attachHaptics(){' +
        '/* Buttons get medium feedback */' +
        'var buttons=document.querySelectorAll("button,.btn,.action-btn,.fab");' +
        'for(var i=0;i<buttons.length;i++){' +
          'if(!buttons[i].dataset.haptic){' +
            'buttons[i].dataset.haptic="1";' +
            '(function(btn){' +
              'var origClick=btn.onclick;' +
              'btn.addEventListener("touchstart",function(){' +
                '/* Determine feedback type from button context */' +
                'var cls=btn.className||"";' +
                'var txt=(btn.textContent||"").toLowerCase();' +
                'if(cls.indexOf("danger")>-1||cls.indexOf("delete")>-1||txt.indexOf("delete")>-1||txt.indexOf("remove")>-1){' +
                  'hapticFeedback("warning");' +
                '}else if(cls.indexOf("success")>-1||txt.indexOf("save")>-1||txt.indexOf("submit")>-1||txt.indexOf("confirm")>-1){' +
                  'hapticFeedback("medium");' +
                '}else if(cls.indexOf("primary")>-1){' +
                  'hapticFeedback("medium");' +
                '}else{' +
                  'hapticFeedback("light");' +
                '}' +
              '},{passive:true});' +
            '})(buttons[i]);' +
          '}' +
        '}' +

        '/* Checkboxes and radio buttons get light feedback */' +
        'var checks=document.querySelectorAll("input[type=checkbox],input[type=radio],.option-item");' +
        'for(var j=0;j<checks.length;j++){' +
          'if(!checks[j].dataset.haptic){' +
            'checks[j].dataset.haptic="1";' +
            'checks[j].addEventListener("touchstart",function(){hapticFeedback("light");},{passive:true});' +
          '}' +
        '}' +

        '/* Tabs and filters get light feedback */' +
        'var tabs=document.querySelectorAll(".tab,.filter,.search-tab,.nav-tab,.toggle-btn");' +
        'for(var k=0;k<tabs.length;k++){' +
          'if(!tabs[k].dataset.haptic){' +
            'tabs[k].dataset.haptic="1";' +
            'tabs[k].addEventListener("touchstart",function(){hapticFeedback("light");},{passive:true});' +
          '}' +
        '}' +

        '/* Select dropdowns get light feedback */' +
        'var selects=document.querySelectorAll("select");' +
        'for(var m=0;m<selects.length;m++){' +
          'if(!selects[m].dataset.haptic){' +
            'selects[m].dataset.haptic="1";' +
            'selects[m].addEventListener("touchstart",function(){hapticFeedback("light");},{passive:true});' +
          '}' +
        '}' +
      '}' +

      '/* Run on DOM ready and observe for dynamic content */' +
      'if(document.readyState==="loading"){' +
        'document.addEventListener("DOMContentLoaded",attachHaptics);' +
      '}else{' +
        'attachHaptics();' +
      '}' +

      '/* MutationObserver to handle dynamically added elements */' +
      'if(typeof MutationObserver!=="undefined"){' +
        'var observer=new MutationObserver(function(mutations){' +
          'var shouldReattach=false;' +
          'for(var i=0;i<mutations.length;i++){' +
            'if(mutations[i].addedNodes.length>0){shouldReattach=true;break;}' +
          '}' +
          'if(shouldReattach)attachHaptics();' +
        '});' +
        'observer.observe(document.body,{childList:true,subtree:true});' +
      '}' +

      '/* Provide success/error feedback for google.script.run callbacks */' +
      'var origRun=google.script&&google.script.run;' +
      'if(origRun){' +
        'var origSuccess=origRun.withSuccessHandler;' +
        'var origFailure=origRun.withFailureHandler;' +
        'if(origSuccess){' +
          'google.script.run.withSuccessHandler=function(fn){' +
            'return origSuccess.call(this,function(){' +
              'hapticFeedback("success");' +
              'if(fn)fn.apply(this,arguments);' +
            '});' +
          '};' +
        '}' +
        'if(origFailure){' +
          'google.script.run.withFailureHandler=function(fn){' +
            'return origFailure.call(this,function(){' +
              'hapticFeedback("error");' +
              'if(fn)fn.apply(this,arguments);' +
            '});' +
          '};' +
        '}' +
      '}' +
    '})();' +
  '</script>';
}

/**
 * Returns a complete mobile-optimized <head> section for modal dialogs.
 * Combines viewport meta tag, responsive styles, and haptic feedback.
 * Usage: Include the return value right after <head> in any modal HTML.
 * @returns {string} Combined HTML string for mobile optimization
 */
function getMobileOptimizedHead() {
  return getMobileMetaTag() + getMobileResponsiveStyles() + getHapticFeedbackScript();
}
