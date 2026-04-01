// GAS_LIMITATION: All .gs files share a single global namespace. Function names must be globally unique.
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
 * @version 4.48.0
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
 * Structured error codes for client-consumable error classification
 * @enum {string}
 */
var ERROR_CODES = {
  UNKNOWN: 'UNKNOWN',
  AUTH_EXPIRED: 'AUTH_EXPIRED',
  AUTH_REQUIRED: 'AUTH_REQUIRED',
  NOT_FOUND: 'NOT_FOUND',
  LOCK_TIMEOUT: 'LOCK_TIMEOUT',
  QUOTA_EXCEEDED: 'QUOTA_EXCEEDED',
  VALIDATION: 'VALIDATION',
  SHEET_UNAVAILABLE: 'SHEET_UNAVAILABLE',
  DUPLICATE: 'DUPLICATE'
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
 * @param {string} [errorCode] - Structured error code from ERROR_CODES enum
 * @returns {{success: false, error: string, context: string, errorCode: string}}
 */
function errorResponse(error, context, errorCode) {
  return { success: false, error: error || 'An unexpected error occurred', context: context || '', errorCode: errorCode || ERROR_CODES.UNKNOWN };
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
// STANDARD: All new code should use handleError() for consistent error handling.
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
    if (!ss) return; // Web app context — cannot log to sheet
    var sheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(ERROR_CONFIG.LOG_SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Level', 'Context', 'Message', 'User', 'Stack']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      setSheetVeryHidden_(sheet);
    }

    // Add error row (formula-protect user-influenced fields)
    var safeContext = typeof escapeForFormula === 'function' ? escapeForFormula(errorInfo.context) : errorInfo.context;
    var safeMessage = typeof escapeForFormula === 'function' ? escapeForFormula(errorInfo.message) : errorInfo.message;
    sheet.appendRow([
      errorInfo.timestamp,
      errorInfo.level,
      safeContext,
      safeMessage,
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
    // Gather admin recipients from Config (same pattern as sendSecurityAlertEmail_)
    var adminEmail = '';
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        var configSheet = ss.getSheetByName(typeof SHEETS !== 'undefined' ? SHEETS.CONFIG : 'Config');
        if (configSheet && typeof CONFIG_COLS !== 'undefined') {
          adminEmail = String(configSheet.getRange(3, CONFIG_COLS.CHIEF_STEWARD_EMAIL).getValue() || '').trim();
        }
      }
    } catch (_) {}
    // Fallback to effective user (script owner in web app context)
    if (!adminEmail) {
      try { adminEmail = Session.getEffectiveUser().getEmail(); } catch (_) {}
    }
    if (!adminEmail) return; // No recipient available

    var subject;
    try {
      subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Critical Error: ' + errorInfo.context;
    } catch (_e) {
      subject = 'Critical Error: ' + (errorInfo.context || 'Unknown');
    }
    var systemName;
    try {
      systemName = COMMAND_CONFIG.SYSTEM_NAME;
    } catch (_e) {
      systemName = 'SolidBase';
    }
    // Mask PII in the email body
    var maskedUser = typeof maskEmail === 'function' ? maskEmail(errorInfo.user) : '[REDACTED]';
    var body = 'A critical error occurred in the ' + systemName + ':\n\n' +
               'Time: ' + errorInfo.timestamp + '\n' +
               'Context: ' + errorInfo.context + '\n' +
               'Message: ' + errorInfo.message + '\n' +
               'User: ' + maskedUser + '\n\n' +
               'Stack Trace:\n' + errorInfo.stack;

    if (typeof safeSendEmail_ === 'function') {
      safeSendEmail_({ to: adminEmail, subject: subject, body: body });
    } else {
      MailApp.sendEmail(adminEmail, subject, body);
    }
  } catch (e) {
    Logger.log('Failed to send critical error notification: ' + e.message);
  }
}

// ============================================================================
// STRATEGIC COMMAND CENTER CONFIG (v4.0)
// ============================================================================

/**
 * Strategic Command Center configuration
 * Controls system-wide settings for the Dashboard
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
  VERSION: "4.50.4",

  // Document Templates (configure these with your Drive IDs)
  TEMPLATE_ID: '',  // Google Doc template ID for grievance PDFs
  ARCHIVE_FOLDER_ID: '',  // Drive folder ID for archived documents

  // Alert Configuration
  CHIEF_STEWARD_EMAIL: '',  // Read at runtime via getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL)
  // Escalation triggers - for STATUS column use status values, for CURRENT_STEP use step values
  ESCALATION_STATUSES: ['In Arbitration', 'Appealed'],  // Status values that trigger alerts
  ESCALATION_STEPS: ['Step II', 'Step III', 'Arbitration'],  // Step values that trigger alerts

  // Unit Code Prefixes
  // Format in Config: "Main Station:MS,Field Ops:FO,Health:HC,Admin:AD,Remote:RM"
  UNIT_CODES: {},  // Read at runtime from Config sheet — not populated here

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
  // Root folder name — derived dynamically from ORG_NAME in Config sheet at runtime.
  // Use getDriveRootFolderName_() everywhere you need the folder name.
  // Fallback used only if Config sheet is not yet set up (first-time CREATE_DASHBOARD run).
  ROOT_FOLDER_FALLBACK: 'Dashboard Files',
  // Subfolder names within the root — these are fixed (renaming would break stored IDs)
  GRIEVANCES_SUBFOLDER:      'Grievances',
  RESOURCES_SUBFOLDER:       'Resources',
  MINUTES_SUBFOLDER:         'Minutes',
  EVENT_CHECKIN_SUBFOLDER:   'Event Check-In',
  MEMBERS_SUBFOLDER: 'Members',  // per-member master admin folders (steward-only hub)
  // Grievance case folder naming templates (inside Grievances/)
  // Template uses placeholders: {date}, {lastName}, {firstName}
  SUBFOLDER_TEMPLATE: '{lastName}, {firstName} - {date}',
  // Fallback if member name not available
  SUBFOLDER_TEMPLATE_SIMPLE: '{grievanceId} - {date}'
};

/**
 * Returns the root Drive folder name for this deployment.
 * Derived from the ORG_NAME value in the Config sheet (row 3).
 * Example: "Union Local" → "Union Local Dashboard"
 * Falls back to DRIVE_CONFIG.ROOT_FOLDER_FALLBACK if Config is not yet set up.
 * Memoized per script execution (same pattern as getSystemName_).
 * @returns {string} Root folder name
 */
var _cachedDriveRootName_ = null;
function getDriveRootFolderName_() {
  if (_cachedDriveRootName_ !== null) return _cachedDriveRootName_;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && CONFIG_COLS.ORG_NAME) {
      var orgName = String(configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue() || '').trim();
      if (orgName) {
        _cachedDriveRootName_ = orgName + ' Dashboard';
        return _cachedDriveRootName_;
      }
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  _cachedDriveRootName_ = DRIVE_CONFIG.ROOT_FOLDER_FALLBACK;
  return _cachedDriveRootName_;
}

/**
 * Get organization name from Config sheet, falling back to default.
 * Single source of truth for the org name used across the system.
 * @private
 * @returns {string} Organization name (e.g., "Union Local")
 */
// M-43: _cachedOrgName / _cachedSystemName / _cachedLocalNumber are intentionally
// never invalidated. In Google Apps Script, each script execution is short-lived
// (max 6 minutes) and runs in an isolated context, so module-level caches are
// automatically cleared when the execution ends. No manual invalidation is needed.
var _cachedOrgName = null;
function getOrgNameFromConfig_() {
  if (_cachedOrgName !== null) return _cachedOrgName;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue();
      if (orgName) {
        _cachedOrgName = orgName;
        return _cachedOrgName;
      }
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  _cachedOrgName = 'SolidBase';
  return _cachedOrgName;
}

/**
 * Get system name from Config sheet org name, falling back to default.
 * Derives the tool/system name from the org name for branding consistency.
 * Used by COMMAND_CONFIG.SYSTEM_NAME, EMAIL, and PDF getters.
 * @private
 * @returns {string} System name (e.g., "Strategic Command Center")
 */
var _systemNameCache_ = null;
function getSystemName_() {
  if (_systemNameCache_ !== null) return _systemNameCache_;
  var localNum = getLocalNumberFromConfig_();
  _systemNameCache_ = (localNum ? localNum + ' ' : '') + 'Strategic Command Center';
  return _systemNameCache_;
}

/**
 * Get local number from Config sheet, falling back to default.
 * Used for UI elements like menu names.
 * Memoized: reads from Config sheet once per execution, then returns cached value.
 * @private
 * @returns {string} Local number
 */
var _cachedLocalNumber = null;
function getLocalNumberFromConfig_() {
  if (_cachedLocalNumber !== null) return _cachedLocalNumber;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var localNumber = configSheet.getRange(3, CONFIG_COLS.LOCAL_NUMBER).getValue();
      if (localNumber) {
        _cachedLocalNumber = String(localNumber);
        return _cachedLocalNumber;
      }
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  _cachedLocalNumber = '';
  return _cachedLocalNumber;
}

// ============================================================================
// VERSION INFO
// ============================================================================

/**
 * Version information for build system and display.
 * M-51: Derives version string from COMMAND_CONFIG.VERSION to avoid duplication.
 * @const {Object}
 */
var VERSION_INFO = (function() {
  var ver = (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.VERSION) ? COMMAND_CONFIG.VERSION : '4.50.0';
  var parts = ver.split('.');
  return {
    version: ver,
    MAJOR: parts.length > 0 ? parseInt(parts[0], 10) : 4,
    MINOR: parts.length > 1 ? parseInt(parts[1], 10) : 50,
    PATCH: parts.length > 2 ? parseInt(parts[2], 10) : 0,
    BUILD: 'v' + ver,
    CURRENT: ver,
    BUILD_DATE: '2026-04-01',
    CODENAME: 'Config Data Migration',
    codename: 'Config Data Migration'
  };
})();

/**
 * Complete version history with release dates and codenames.
 * Ordered newest-first. Every version that has ever shipped is listed here
 * so that UI, audit, and diagnostic code can look up any past release date.
 * @const {Array<Object>}
 */
var VERSION_HISTORY = [
  { version: '4.50.4', date: '2026-04-01', codename: 'Config Column Safety', changes: 'Fixed _migrateOrphanedColumns to use name-based lookup instead of sequential two-pointer walk — old algorithm would incorrectly delete valid columns when CONFIG_HEADER_MAP_ was reordered. Global-scope column map init now falls back to syncColumnMaps() on cold cache instead of silently using wrong array-order defaults, protecting all execution contexts (menu handlers, data functions, triggers).' },
  { version: '4.50.3', date: '2026-04-01', codename: 'Config Data Migration', changes: 'Fixed Config sheet data misalignment after column reorder (v4.50.0). createConfigSheet() now detects when existing headers are in a different order than CONFIG_HEADER_MAP_ and migrates row 3+ data to match new positions before overwriting headers. Prevents org name reading "yes"/"no" from boolean toggle columns. Also hardened _applyYesNoValidation to not overwrite non-boolean values — logs warning instead of silently converting misaligned data to "yes".' },
  { version: '4.50.2', date: '2026-03-31', codename: 'Config Cache Fix', changes: 'Fixed stale ConfigReader cache after cold column-map sync. When RESOLVED_COL_MAPS cache expires and syncColumnMaps() re-resolves column positions, ConfigReader now force-refreshes instead of serving cached values built with wrong (array-order) positions. syncColumnMaps() also invalidates ORG_CONFIG_v2 cache key when columns change. Fixes org name showing as "yes" and steward label showing URL after cache expiry.' },
  { version: '4.50.0', date: '2026-03-31', codename: 'Non-Member Fields', changes: 'Added Union Name, Shirt Size, and Steward (Yes/No) columns to Non-Member Contacts tab. New fields in NMC_HEADER_MAP_, sheet creation validations (Shirt Size dropdown, Steward Yes/No dropdown), full CRUD support in DataService, steward_view.html modal and contact cards, and fallback modal in 03_UIComponents.gs.' },
  { version: '4.48.0', date: '2026-03-30', codename: 'Tab Modals', changes: 'Auto-opening contextual modals for 9 key sheet tabs: Config, Member Directory, Grievance Log, Case Checklist, Feedback & Development, Volunteer Hours, Meeting Attendance, Meeting Check-In Log, Resources. Each modal shows tab-specific quick actions and tips. Triggered on tab switch via onSelectionChange sheet-change detection. Config toggle ENABLE_TAB_MODALS (default: yes) for system-wide control. Per-user "Don\'t show again" dismissal via UserProperties. New Tools > Tab Modals submenu for manual access. TAB_MODAL_REGISTRY in 01_Core.gs maps sheets to modal functions.' },
  { version: '4.45.0', date: '2026-03-28', codename: 'Phone PIN', changes: 'Phone-number-derived PINs: generateMemberPIN() now uses first 6 digits of member phone number when available, falling back to random. Email notifications hint at phone-based PIN without exposing it. PIN sessions blocked from poll submissions (wqSubmitPoolQuestion). Theme cleanup: removed 5 novelty themes (Comic, Brutalist, Retro OS, Liquid Pour, Blob Lava). Removed Lao Tzu duplicate quote from Knowledge Engine.' },
  { version: '4.44.0', date: '2026-03-28', codename: 'Idle Prefetch', changes: 'Performance optimization suite. (1) Dynamic tab prefetch: _recordTabFrequency() tracks per-role tab visit counts in localStorage (dds_tabFreq); _getTopTabs() determines user top 3; _PREFETCH_REGISTRY warms DataCache via requestIdleCallback. (2) Request deduplication: _dedupMap in _throttledServerCall coalesces identical in-flight server calls — if dataGetAllMembers is pending and another caller requests it, piggybacks on existing call. (3) Cases list converted to LazyList with IntersectionObserver auto-scroll (replaces manual Show More button). (4) Visibility-based refresh: visibilitychange listener silently re-fetches batch data when user returns after 30+ min away. (5) Extended caching: DataCache default 2min→15min, stable 5min→30min, new STATS_TTL 60min, SWR 30min→2hrs. Notifications tab now cacheable.' },
  { version: '4.43.1', date: '2026-03-27', codename: 'One-Tap Check-In', changes: 'In-app one-tap meeting check-in banner. When a steward activates a meeting, logged-in members and stewards see a green check-in banner at the top of their dashboard. One tap checks them in — no PIN re-entry needed (session-authenticated). Banner shows confirmed state after check-in or if already checked in. Active meeting data included in batch payloads for zero-latency display. dataWebAppCheckIn() endpoint with TOCTOU lock and audit logging (MEETING_WEBAPP_CHECKIN). New .checkin-banner CSS with slide-in animation.' },
  { version: '4.43.0', date: '2026-03-27', codename: 'QR Attendance', changes: 'QR code mobile attendance system. Stewards generate QR codes for meetings (Calendar & Meetings > QR Code Check-In). Members scan with phone, enter phone number + PIN to check in. New ?page=qr-checkin web route with dark/light mode mobile-optimized UI. processQRCheckIn() authenticates by phone number instead of email. getMeetingQRCode() generates QR via Google Charts API. createMeeting() now includes QR URL in response. Full security: PIN lockout, rate limiting, TOCTOU lock, audit logging (MEETING_QR_CHECKIN). Device key authentication enables biometric sign-in for all login methods. Remember-me defaults to ON. PIN auto-submits on 6th digit.' },
  { version: '4.42.0', date: '2026-03-27', codename: 'Biometric Login', changes: 'Password manager and biometric sign-in support. PIN and email forms wrapped in <form> elements with proper autocomplete attributes (username, current-password, email) for password manager detection. Credential Management API integration saves PIN credentials after successful login (Chrome/Edge). iOS/Safari: form semantics trigger native Keychain save with Face ID/Touch ID autofill. Biometric sign-in button on auth choose screen when saved credentials exist. Platform-aware biometric labels (Face ID, Touch ID, Fingerprint, Windows Hello). Auto-submit flow for Credential API retrieval. Forget saved credentials option on login screen.' },
  { version: '4.40.0', date: '2026-03-25', codename: 'PIN Login GA', changes: 'PIN login enabled in all environments (dev + production). Removed IS_DEV_MODE gates from auth_view.html PIN button, steward_view.html Manage PIN button, and 13_MemberSelfService.gs comments. Added isDevMode to auth page PAGE_DATA. Login page restyled: black background, glass card with backdrop blur, brighter quotes with text glow. Default org name fallback changed to SolidBase.' },
  { version: '4.39.0', date: '2026-03-24', codename: 'DEV PIN Login', changes: 'DEV ONLY (IS_DEV_MODE gate): (1) devAuthLoginByPIN(pin) — PIN-only login scan, no email, global rate limit 10/15min, audit logged. (2) showAuthPIN() on login screen — only when IS_DEV_MODE=true, redirects with sessionToken. (3) devStewardManageMemberPIN(sessionToken, email) — steward generates/resets PIN from webapp, returns plaintext once. (4) Manage PIN button in steward member detail panel — one-time display with copy button. All dev-only functions are IS_DEV_MODE gated.' },
  { version: '4.38.0', date: '2026-03-24', codename: 'Reliability', changes: 'Default view preference for dual-role users (sidebar selector, dataSetDefaultView endpoint, ScriptProperties storage). Loading fence on _loadMemberViewThen prevents double-tap concurrent fetches. Try/catch on mobile header role-switch buttons. Magic-link auth fix: removed Session.getActiveUser() checks from getMemberViewHtml/getOrgChartHtml/getPOMSReferenceHtml (returns empty for Execute-as-Me). Test runner included in prod builds (tab gated by IS_DEV_MODE). initMemberView verification after script injection with error UI.' },
  { version: '4.37.1', date: '2026-03-24', codename: 'Lazy Member View', changes: 'Lazy-load member_view.html for dual-role users to stay under GAS ~820KB HtmlOutput limit. Conditional template inclusion: steward/both get steward_view only, member-only get member_view inline. getMemberViewHtml() server endpoint for on-demand fetch.' },
  { version: '4.36.0', date: '2026-03-24', codename: 'Centralized Modal Hub', changes: 'New Modal Hub: centralized launcher for all 35+ dialogs organized by category. Searchable with real-time filtering. Master enable/disable toggle (ENABLE_MODAL_HUB). MODAL_REGISTRY in 04a_UIMenus.gs catalogs all modals. Tab review XSS/error handling fixes: DataService try/catch, AUDIT_LOG_COLS constants, poms_reference pesc() XSS prevention, grievance_form failureHandler + email injection guard.' },
  { version: '4.35.0', date: '2026-03-24', codename: 'Welcome Guide', changes: 'New Welcome Guide tab for both steward and member roles with expandable sections covering navigation, features, tips, and quick-launch buttons.' },
  { version: '4.34.5', date: '2026-03-24', codename: 'Android Login Fix', changes: 'Auth: magic link auth now always creates a session token (24h short-lived without remember-me, full cookieDuration with it) so users are not locked out after consuming the one-use token. Remember-me toggle defaults to ON for email link flow. SSO failure message updated to explain mobile/Android cookie restrictions and direct users to email link option.' },
  { version: '4.34.4', date: '2026-03-24', codename: 'Survey Period Cache Fix', changes: 'Fix survey visibility mismatch: steward view showed active survey but member view showed "No Survey Open". Root cause: getSurveyQuestions() cached the period status for 5 minutes, and openNewSurveyPeriod()/archiveSurveyPeriod_() did not invalidate the cache. Members hitting the cache received stale period=null. Fix: period is now always fetched fresh via getSurveyPeriod() on every call to getSurveyQuestions() (both cache-hit and cache-miss paths). Questions/sections remain cached (5 min). Belt-and-suspenders: openNewSurveyPeriod() and archiveSurveyPeriod_() now explicitly remove surveyQuestions_v1 cache key on period state change.' },
  { version: '4.34.3', date: '2026-03-23', codename: 'Mentorship Hardening', changes: 'Mentorship tab review fixes. createPairing: email validation, duplicate pairing guard, escapeForFormula on emails, self-pairing prevention. closePairing: MENTORSHIP_CLOSED audit event logged. All write ops call _refreshNavBadges(). suggestPairings: load-balanced assignment accounts for existing active pairings per mentor instead of naive round-robin. Frontend: manual pairing creation form with inline validation, Save Notes shows Saved!/error text feedback with auto-reset, accepting a suggestion refreshes the Active Pairings list.' },
  { version: '4.34.2', date: '2026-03-22', codename: 'Enable Defaults', changes: 'Enable all safety/security/maintenance features by default. BROADCAST_SCOPE_ALL default changed from no to yes. ENABLE_CORRELATION now seeded as yes with yes/no dropdown validation. All security features (ACCESS_CONTROL, DASHBOARD_MEMBER_AUTH, ERROR_LOGGING, NOTIFY_ON_CRITICAL) were already enabled by default.' },
  { version: '4.34.1', date: '2026-03-21', codename: 'Theme Unification', changes: 'Theme sync: org_chart.html and poms_reference.html now follow AppState dark/light toggle (class toggle on .madds-embed and .poms-root). Live toggle via UnifiedTheme.apply() also updates embedded components. esign.html: 7 hardcoded colors replaced with CSS vars (--accentHover, --accentMuted, --badge-*, --canvas-bg) with dark media query overrides. OS theme detection: first-visit default now respects prefers-color-scheme instead of always-dark. Modal overlay uses theme-aware --overlay-bg (light=0.35, dark=0.55). Offline banner uses var(--danger) instead of hardcoded #ef4444. Dead code removed: unused last7Keys variable, orphaned resourceDownloads local computation. Build fix: minifyHtml() now normalizes \\r\\n→\\n before regex processing, fixing dist-parity CI failure on Windows builds.' },
  { version: '4.31.1', date: '2026-03-20', codename: 'Public Dashboard Removal', changes: 'Remove orphaned 04e_PublicDashboard.gs (~3K lines). File was never routed via doGet and confirmed no public dashboard will exist. Deleted: src/04e_PublicDashboard.gs, dist/04e_PublicDashboard.gs, test/04e_PublicDashboard.test.js. Removed from build.js file list and architecture/auth-denial/UIService test file lists. XSS threshold in architecture.test.js lowered from 130 to 10 (04e contributed ~122 false positives). Updated all documentation references across AI_REFERENCE.md, CODE_REVIEW.md, DEVELOPER_GUIDE.md, FEATURES.md, README.md, CONTRIBUTING.md, QUICK_DEPLOY.md, presentation.html. Changelog entries in VERSION_HISTORY preserved as historical record.' },
  { version: '4.32.0', date: '2026-03-19', codename: 'Workforce Mobility Survey', changes: 'New survey Section 13 (WORKFORCE_RETENTION) and Section 13A (WORKFORCE_LEAVING) added to quarterly member survey. Section 13 (always shown, 4 questions): q80 likelihood to stay (radio: Very Likely–Very Unlikely), q81 exploring outside organization (radio-branch — Yes triggers 13A), q84 union addressing retention factors (slider-10), q86 optional open text. Section 13A (conditional on q81=Yes, 2 questions): q82 types of opportunities outside organization (checkbox, max 2: state transfer | leaving state service | private | non-profit/education | not sure), q83 reasons for leaving (checkbox, max 3: Pay & Benefits | Workload | Management | Limited Advancement or Transfer Opportunities | Work-Life Balance | RTO Policy | Burnout | Culture | Other — transfer awareness from former q85 folded into this option). Standalone q85 removed. Looker: Workforce Retention Avg, Likelihood to Stay, Exploring Outside Org columns added to both _Looker_Satisfaction and _Looker_Anon_Satisfaction headers and refresh functions. getSatisfactionSummary() workforce section added. Section colors: WORKFORCE_RETENTION #e8f4f8, WORKFORCE_LEAVING #fdecea. Always active (not toggled). Visible to all dashboard roles.' },
  { version: '4.31.0', date: '2026-03-17', codename: 'Security Hardening', changes: 'Security: magic token immediate-delete (TOCTOU fix), session token error handling, bootstrap admin audit logging, resource IDs removed from client config. Reliability: email index cache invalidation, Drive sharing fatal errors, preload race guard, layout render generation counter, timer/observer cleanup. Features: 15min auto-logout with 2min warning modal, Chart.js SRI integrity hash. Perf: glow animation transform/opacity, keyboard handler dedup. Removed ~1385 lines deprecated Interactive Dashboard code.' },
  { version: '4.30.2', date: '2026-03-17', codename: 'Contact Log Name Matching', changes: 'Contact Log stores member name (col 9). Autocomplete triggers at 1 char. By Member tab gets autocomplete. Recent contacts display member name.' },
  { version: '4.28.7', date: '2026-03-15', codename: 'Gmail Scope Test + Auth Sweep', changes: 'Fix gmailAppAccessible test (gmail.send has no side-effect-free probe). Fix testRunnerEndpointsGated false failure (SSO bypasses null-token rejection). 4 new authsweep tests for wq/qa/tl/fs endpoints. Full auth sweep: 100+ endpoints verified, 10 scopes confirmed, no auth gaps.' },
  { version: '4.28.2', date: '2026-03-14', codename: 'Union Stats + Tab Reviews', changes: 'Tab review merge: Polls — null sessionToken fixed in 5 client calls, server-side dues gate, rate limiting, confirmation dialog, 30+ new tests. POMS — steward routing fix, description correction, auth check on getPOMSReferenceHtml. Union Stats — member-safe grievance/hot-spots endpoints, membership sub-tab enrichment, client-side caching, engagement KPI fix, redundant sheet read fix, showLoading consistency, dataAssignSteward auth fix, 8 structural guard tests. Workload — crash-safe refresh, atomic rate limiting, dynamic bar chart scaling, WT_CAT_KEY_LABELS, auto-save draft, last-submitted indicator, 44 new tests.' },
  { version: '4.26.0', date: '2026-03-13', codename: 'Diagnostics + Quick Setup Menu + Contrast Fix', changes: 'Fix #5: TestRunner controls card contrast upgraded (accent border, card background, stronger shadow). Fix #7: Case Analytics deprecated menu item removed (showInteractiveDashboardTab). Fix #14: New Quick Setup (All Init/Sync) consolidated admin menu with all initialize, trigger install, sync, refresh, and setup functions in one place. Fix #15/#16: Workload Archive removed from DIAGNOSE_SETUP skipKeys so diagnostics detects it and repair creates it via setupHiddenSheets. Fix #17: TestRunner.fail already exists; MEMBER_COLS.ROLE (col 44) needs sheet column addition; remaining test failures are deployment-related — run clasp push.' },
  { version: '4.25.15', date: '2026-03-13', codename: 'Survey UX + Insights Enrichment + Menu Consolidation', changes: 'TestRunner dark-mode contrast fix. Survey default scope changed to location. Quick Setup & Sync duplicate menu removed (35 lines). Proximity badge pills on nearby survey members. Prev/Next pagination replaces Show All buttons. Enhanced participation stats (new member completion fraction, declining avg rate). withdrawnCount added to getGrievanceStats. newMembersLast90 + byHireMonth added to getMembershipStats. New insight cards: OVERDUE CASES, NEW MEMBERS, overtime averages, employment mix. Hash-based detail navigation with browser back-button support.' },
  { version: '4.25.12', date: '2026-03-12', codename: 'Function Cohesion Phase 2', changes: 'Inline lock patterns replaced with withScriptLock_() in 6 functions (addMember, updateMember, startNewGrievance, advanceGrievanceStep, bulkUpdateGrievanceStatus, resolveGrievance). Centralized ID validators isGrievanceId_()/isMemberId_() replace 10+ inline regex checks across 04c, 04d, 04e, 05, 09. Hardcoded grievance status strings replaced with GRIEVANCE_STATUS constants and GRIEVANCE_CLOSED_STATUSES across 03, 04e, 06, 07, 08b, 09, 11, 12, 17. onEdit() if/else chain replaced with EventBus dispatch via emitEditEvent() — all sheet-specific handlers now route through priority-ordered EventBus subscribers. ThemeEngine consolidated in auth_view.html and error_view.html. Dashboard entry points deprecated (showExecutiveDashboard → showStewardDashboard, showInteractiveDashboardTab → showStewardDashboard).' },
  { version: '4.25.11', date: '2026-03-12', codename: 'Function Cohesion & Process Consistency', changes: 'Comprehensive cohesion review and refactoring. UX fixes: badge refresh added to qaSubmitQuestion (member_view.html), missing withFailureHandler added to 3 steward task creation paths (steward_view.html). Auth fixes: QA moderation wrappers (qaModerateQuestion, qaModerateAnswer, qaGetFlaggedContent, qaResolveQuestion) now return consistent {success:false} objects instead of null; sessionToken passed to checkWebAppAuthorization in dataUpdateProfile and qaResolveQuestion; deprecated poll stubs (dataGetActivePolls/dataSubmitPollVote/dataAddPoll) now enforce auth gates. Cohesion: extracted maskObjectPII_() shared helper to replace 3 identical PII masking loops in 00_Security.gs (secureLog, sendSecurityAlertEmail_, queueSecurityDigestEvent_); sendDailySecurityDigest now routes through safeSendEmail_() instead of direct MailApp.sendEmail; shared fmtDateShort_() and hashEmail_() helpers in 01_Core.gs replace duplicate private definitions in QAForum, TimelineService, FailsafeService.' },
  { version: '4.30.0', date: '2026-03-16', codename: 'Review Findings Implementation', changes: 'All 3 code review findings implemented (61 total, 41 addressed). Phase 1: removed 213 dead-code stub comments, 22 console.error→Logger.log fixes, removed redundant EventBus init from onOpenDeferred_ (~200ms saved), 3 lock-failure sites now logged, 5 auth-error return types standardized. Phase 2: 130 empty catch blocks now log errors, addToConfigDropdown_ lock-protected, TimelineService 2-min in-memory cache, CorrelationEngine config-gated (Enable Correlation Engine), member profile batch setValue (11→2 API calls), QA Forum notification outside lock scope. Phase 3: Grievance archive split (_Archive_Grievances sheet, stats/hotspots/history read both active+archive, 10-min archive cache TTL), 6 calc sheets deprecated, audit log auto-archival (archiveOldAuditLogs_ + dailyAuditArchive trigger wrapper), lazy-load verified. Phase 4: getConfigValue_ fallback parameter, EventBus sheetKeyMap cached at module level, audit logging throttled (1/sec/user), column map cache TTL 6h→2h, _getSheetSafe global helper.' },
  { version: '4.25.11', date: '2026-03-12', codename: 'Web App Test Suites', changes: '9 new GAS-native test suites in 31_WebAppTests.gs covering all web app modules: webapp (doGet routing, templates, diagnoseWebApp), configrd (ConfigReader completeness, validation, JSON output), portal (PortalSheets 0-indexed column constants, sheet setup), weeklyq (WeeklyQuestions API, poll frequency, pool count), workload (WorkloadService categories, health status, sub-categories), qaforum (QAForum API, pagination, flagged content), timeline (TimelineService events, categories), failsafe (FailsafeService digest config, diagnostics), endpoints (comprehensive data/wq/qa/tl/fs wrapper existence + write endpoint null-token rejection). Total: 20 suites, ~170 tests. All read-only.' },
  { version: '4.25.11', date: '2026-03-12', codename: 'Seed Data Expansion & Survey UX', changes: 'New seedQAForumData() seeds 10 realistic Q&A questions with 15 answers (mix of steward and member responses, anonymous posts, upvotes). Added to SEED_PHASE_3. Member seed data now assigns 2-3 random office days (comma-separated multi-select) instead of 1. Survey Tracking: individual member participation progress bars (snapped to 5% chunks) showing lifetime completion rate with color coding (green >= 80%, yellow >= 50%, red < 50%). getPendingSurveyMembers() now returns allMembers array with totalCompleted/totalMissed stats. Loading indicator replaced: spinner removed, skeleton placeholder UI (pulsing cards/rows) for subtler loading state.' },
  { version: '4.25.10', date: '2026-03-11', codename: 'Web App Diagnostics & Parallel Render', changes: 'Parallel view rendering in index.html (steward + member views load concurrently). Nav tab reorder: Feedback moved to Admin section, stewarddirectory removed from member sidebar. diagnoseWebApp() 14-step diagnostic function in 22_WebDashApp.gs for debugging app loading issues. memberId added to profile data in 21_WebDashDataService.gs. Workload sheet diagnostics in 28_FailsafeService.gs. Enhanced error reporting: null-guard on fatalErr.stack, actual error message shown in bootstrap screen instead of generic failure.' },
  { version: '4.25.7', date: '2026-03-10', codename: 'onOpen Simple Trigger Fix', changes: 'Fixed critical bug: onOpen() called ScriptApp.getProjectTriggers() (not permitted in simple triggers — throws silently) and finally block deleted the deferred trigger before 1000ms elapsed. Result: onOpenDeferred_ never ran on sheet open. Fix: onOpen now only does cache clear + createDashboardMenu(). setupOpenDeferredTrigger() installs onOpenDeferred_ as an installable onOpen trigger. menuInstallSurveyTriggers() updated to include this. New menu item in Admin > Triggers.' },
  { version: '4.25.6', date: '2026-03-09', codename: 'Test Runner Timeout Fix', changes: 'Fixed GAS 6-minute execution timeout crash when running all 82 tests. Root cause: 5-min soft guard too tight + 16 redundant SpreadsheetApp.getActiveSpreadsheet() network calls + authsweep endpoints each doing sheet reads via getUserRole_(). Fixes: (1) timeout lowered 5min→3.5min for 2.5min safety margin, (2) _getCachedSS() replaces 12 individual SpreadsheetApp calls in test functions, (3) cache reset at start of runAll(), (4) SPA UI now shows timeout warning banner, skipped test count card, suite skipped counts, and failure handler detects execution time errors with helpful guidance to use suite filter.' },
  { version: '4.25.5', date: '2026-03-09', codename: 'Member View: Remove Directory Tab', changes: 'Removed "Directory" tab (id: contact) from member SPA sidebar in index.html. Members should not have a general Directory — steward contacts are accessed via the separate Steward Directory utility link (id: stewarddirectory). Removed case contact routing entry and contact color mapping from TAB_COLORS. Deep-links to #contact for members now fall through to Home. Updated README.md (expanded member tab listing), FEATURES.md (added Member SPA Sidebar Tabs section, bumped SPA version), INTERACTIVE_DASHBOARD_GUIDE.md (member tab count 10→9, removed Directory from tab list), CHANGELOG.md, AI_REFERENCE.md.' },
  { version: '4.25.4', date: '2026-03-09', codename: 'Stability: Timeout Guard + Trigger Audit', changes: 'Fixed webapp instability caused by GAS-native test runner exceeding 6-minute execution limit. TestRunner.runAll() now has global timeout (MAX_RUNTIME_MS=300000, 5 min) and per-test slow detection (PER_TEST_MAX_MS=30000, 30s). Remaining tests counted as skipped when timeout hit, status set to timeout. New trigger audit utilities in 06_Maintenance.gs: auditAllTriggers() (read-only, logs all installed triggers with duplicate detection), cleanupDuplicateTriggers(dryRun) (removes duplicate triggers, safe dryRun default), dataAuditTriggers(sessionToken) SPA endpoint. Audit Triggers menu item added to Data Integrity menu. Documented: 3.15 MB total .gs code, 558 KB HTML payload (all views inlined via include()), 20+ installable trigger sources across codebase.' },
  { version: '4.25.3', date: '2026-03-09', codename: 'Deadline Config Completeness', changes: 'Added 3 missing deadline Config columns: STEP3_APPEAL_DAYS (header "Step III Appeal Days", default 10), STEP3_RESPONSE_DAYS (header "Step III Response Days", default 21), ARBITRATION_DEMAND_DAYS (header "Arbitration Demand Days", default 30). CONFIG_HEADER_MAP_ expanded (3 entries after STEP2_RESPONSE_DAYS). DEADLINES section header widened from 4 to 7 cols in createConfigSheet. seedConfigDefault_ calls added for all 3. getDeadlineRules() now reads Step III and Arbitration values from Config instead of hardcoded DEADLINE_DEFAULTS; still falls back to defaults if empty/NaN. COMMAND_CONFIG.VERSION fixed from stale "4.24.4" to match actual version.' },
  { version: '4.25.2', date: '2026-03-09', codename: 'Test Failure Notifications', changes: 'Email notifications on scheduled test failures. New Config column: TEST_NOTIFY_EMAIL (Config tab header "Test Runner Notify Email"). When daily trigger runs and tests fail, an email is sent to the configured address with: pass/fail summary, failed test names grouped by suite, error messages, duration, timestamp. Email has both plain-text and styled HTML body (dark theme matching dashboard). Quota guard: skips email if MailApp remaining quota < 5. No email sent on success (failure-only). Manual runs still use toast only. Functions: _getTestNotifyEmail() reads from Config via CONFIG_COLS.TEST_NOTIFY_EMAIL, _sendTestFailureEmail(results) builds and sends the email. New CONFIG_HEADER_MAP_ entry: TEST_NOTIFY_EMAIL after BROADCAST_SCOPE_ALL.' },
  { version: '4.25.1', date: '2026-03-09', codename: 'Test Runner Expansion', changes: '4 new GAS-native test suites added to 30_TestRunner.gs. (1) dataservice suite (10 tests): DataService module existence, all public API methods callable, findUserByEmail shape validation, invalid email returns null, complete public API coverage check. (2) authsweep suite (6 tests): all data* functions exist, steward endpoints reject null token, member endpoints reject null token, dataGetBatchData no data leak on null, poll stubs return safe values, test runner endpoints auth-gated. (3) configlive suite (8 tests): live sheet header existence, CONFIG_COLS/MEMBER_COLS/GRIEVANCE_COLS don\'t exceed actual sheet width, syncColumnMaps callable, Config row 3 has values. (4) survey suite (10 tests): HIDDEN_SHEETS survey constants, SURVEY_PERIODS_COLS/SURVEY_QUESTIONS_COLS defined, getSurveyQuestions returns valid array with correct question shape (id/text/type), getSurveyPeriod callable, submitSurveyResponse exists, SATISFACTION_COLS backward compat, tracking sheet existence. Total: 10 suites, 82 tests. SPA dropdown updated with 4 new filter options.' },
  { version: '4.25.0', date: '2026-03-09', codename: 'GAS-Native Test Runner', changes: 'New GAS-native integration test framework (30_TestRunner.gs). Runs inside the Apps Script runtime against real Sheets/Config/Auth — not mocked. 6 test suites (config, colmap, auth, grievance, security, system) with 48 tests covering: Config tab reads, ConfigReader shape, column mapping integrity (GRIEVANCE_COLS/MEMBER_COLS/CONFIG_COLS), Auth module existence, _resolveCallerEmail/_requireStewardAuth/_checkWebAppAuth, grievance status constants, deadline rules, escapeHtml/escapeForFormula XSS prevention, VERSION_INFO format. SPA dashboard panel (steward-only testrunner tab) with run-all, per-suite filter, pass/fail cards, expandable error details, auto-expand failed suites. Sheets menu item under 🛠️ Admin → 🧪 Test Runner. Daily trigger support (6 AM). Server endpoints: dataRunTests, dataGetTestResults, dataManageTestTrigger — all steward-auth gated. Results stored in ScriptProperties. All tests read-only — never write to sheets. Files: src/30_TestRunner.gs (new), steward_view.html (renderTestRunnerPage), index.html (sidebar + routing), 03_UIComponents.gs (menu). Tests: architecture.test.js + spa-integrity.test.js updated. 2404/2404 Jest tests pass.' },
  { version: '4.24.8', date: '2026-03-08', codename: 'POMS Smart Search Integration', changes: 'Added POMS Reference as a new shared tab (both steward and member roles). 78 POMS sections covering DI/DIB/SSI sequential evaluation, SGA, medical evaluation, med-voc grid rules, CE lifecycle, FTC procedures, CDR (8-step MIRS), child claims, DAA, listings 1.00-14.00, SSI eligibility/income/resources/deeming/work incentives. 17 interactive flowcharts. Features: keyword search with fuzzy scoring, star ratings on explanations and flowcharts, bookmarks, personal notes, quick lookups. Mobile-optimized with 44px touch targets. Files: src/poms_reference.html (CSS-scoped under .poms-root), getPOMSReferenceHtml() in 22_WebDashApp.gs, renderPOMSReference() in index.html. Tab icon: 📘. Lazy-loaded same pattern as org chart.' },
  { version: '4.24.7', date: '2026-03-07', codename: 'Auth Sweep Final — Role Derivation + Residual Fixes', changes: 'FINAL: Three remaining issues closed. (1) getWebAppNotifications: userRole param was still client-trusted — replaced with server-derived role from checkWebAppAuthorization(null,sessionToken).auth.role; treats both/admin as steward for notification targeting. userRole param kept on getWebAppNotificationCount for API compat but ignored server-side. (2) addWebAppResource + sendWebAppNotification: had secondary Session.getActiveUser() call inside function body for audit log (addedBy/stewardEmail) — replaced with auth.email already resolved at function top. (3) getWebAppNotifications: for-loop header (for var i=1) was accidentally stripped during earlier str_replace; restored. restoreFromSnapshot() confirmed sheets-side only (uses ss.toast, never called from client views) — no fix needed.' },
  { version: '4.24.6', date: '2026-03-07', codename: 'Survey Post-Review Fix Batch', changes: '12-issue fix batch from survey system code review. (1) getSatisfactionTrends() fully rewritten — delegates to getSatisfactionSummary(), returns {overall, responseCount, categories[{name,avg}]} matching frontend shape; Insights satisfaction section renders for first time ever. (2) dataGetSatisfactionSummary(sessionToken) — added param, switched _requireStewardAuth→_resolveCallerEmail; both member and steward views can read aggregate anonymous stats; both HTML callers updated to pass SESSION_TOKEN. (3) Survey Version: periodId+"-v"+count replaces bare question count. (4) getActiveSectionKeys() branch rules fully dynamic from questions array via _sectionBranchRules{}. (5) localStorage→window._surveyDraft in-memory store (GAS iframe blocks localStorage). (6) buildSatisfactionColsShim_() in 08c — all deprecated SATISFACTION_COLS key names mapped to dynamic positions; injected into 20 functions across 5 files. (7) Summary write-back guarded if (summaryStart>0). (8) AVG_* per-row reads replaced with inline q-id computations. (9) DevTools maxCol: satSheet.getLastColumn(). (10) 11_CommandHub AVG fallback collapsed. (11) 28 double-paren syntax errors in 21_WebDashDataService.gs — all functions now parse correctly. (12) 5 blocks of getRange+SATISFACTION_COLS col replaced with single full-row reads+slice; stewardRange/leadershipRange ternary guards added.' },
  { version: '4.23.2', date: '2026-03-07', codename: 'Complete Session Token Auth Sweep', changes: 'SWEEP: All remaining client-callable server functions still trusting client-supplied email or using bare getActiveUser() now fixed. 05_Integrations.gs: getWebAppNotifications(sessionToken, role), getWebAppNotificationCount(sessionToken, role), dismissWebAppNotification(sessionToken, id), getAllWebAppNotifications(sessionToken), archiveWebAppNotification(sessionToken, id), addWebAppResource(sessionToken, data), updateWebAppResource(sessionToken, id, data), deleteWebAppResource(sessionToken, id), restoreWebAppResource(sessionToken, id), sendWebAppNotification(sessionToken, data) — all now use _resolveCallerEmail(sessionToken) or checkWebAppAuthorization(role, sessionToken). 25_WorkloadService.gs: processWorkloadFormSSO, getWorkloadHistorySSO, getWorkloadDashboardDataSSO, getWorkloadReminderSSO, setWorkloadReminderSSO, exportWorkloadHistoryCSV — all now use _resolveCallerEmail(sessionToken). FIX: tlGetTimelineEvents client call in member_view.html was missing SESSION_TOKEN (first arg silently treated as page number). CLIENT: 25 additional call sites in steward_view.html and member_view.html updated to pass SESSION_TOKEN.' },
  { version: '4.23.1', date: '2026-03-07', codename: 'System-Wide Session Token Auth Fix', changes: 'SYSTEMIC FIX: All 42+ server wrapper functions that call _requireStewardAuth() or _resolveCallerEmail() now accept sessionToken as first parameter. Root cause: getActiveUser() returns empty in Execute-as-Me webapp for magic link / session token users — every steward operation was silently broken for non-SSO auth. Server: sessionToken param added to all wrappers across 21_WebDashDataService.gs (36 functions), 08e_SurveyEngine.gs (2), 24_WeeklyQuestions.gs (8), 26_QAForum.gs (4), 27_TimelineService.gs (5). Direct Session.getActiveUser() calls in wq wrappers replaced with _resolveCallerEmail(sessionToken). Client: steward_view.html + member_view.html updated to pass SESSION_TOKEN instead of CURRENT_USER.email to all server calls. Q1 fix: server echoes back session token in pageData for method=session users (22_WebDashApp.gs). CLIENT: SESSION_TOKEN now reads PAGE_DATA.sessionToken || localStorage fallback. dataToggleChecklistItem double-paren syntax error fixed.' },
  { version: '4.23.0', date: '2026-03-07', codename: 'Dynamic Survey Schema', changes: 'Option B fully dynamic survey schema. New sheet: 📋 Survey Questions (16 cols: Question ID, Section, Section Key, Section Title, Question Text, Type, Required, Active, Options, Branch Parent, Branch Value, Branch Target, Max Selections, Slider Min/Max, Notes). Owner edits Question Text, Active, Options, Slider Labels, Notes directly — no deployment needed. Adding a new row to Survey Questions auto-creates a new column in 📊 Member Satisfaction on next submission. Setting Active=N deactivates a question. Satisfaction sheet rebuilt with dynamic headers: Timestamp | Period ID | Survey Version | q1 | q2 … qN. New functions: createSurveyQuestionsSheet() (seeds 67 questions, color-coded by section, non-destructive on re-run), getSatisfactionColMap_() (runtime header→col lookup, 5-min cache), syncSatisfactionSheetColumns_() (auto-appends missing question columns), clearSurveyQuestionsCache() (menu-callable). Rewrites: getSurveyQuestions() reads from sheet with 5-min cache; submitSurveyResponse() builds row via col map not hardcoded positions; getSatisfactionSummary() groups questions dynamically by Section Key. 04c and 04d updated to use getSatisfactionColMap_(). SATISFACTION_COLS kept as deprecated reference. New constants: SHEETS.SURVEY_QUESTIONS, SURVEY_QUESTIONS_COLS, SATISFACTION_PREFIX. Wired into CREATE_DASHBOARD setup and initSurveyEngine().' },
  { version: '4.22.7', date: '2026-03-07', codename: 'Survey Form URL Deprecation Cleanup', changes: 'Full removal of SATISFACTION_FORM_URL / satisfactionFormUrl / surveyFormUrl across 8 files. Deprecated since v4.21.0 (Google Form integration replaced by native webapp survey), now fully removed. 01_Core.gs: constant replaced with comment. 08c: satisfaction case removed from getFormUrlFromConfig(); saveFormUrlsToConfig_silent() no longer writes/formats satisfaction URL; sendSurveyCompletionReminders() now reads MOBILE_DASHBOARD_URL (member portal) instead of form URL. 04c: satisfactionForm removed from resource links; range now stops at ORG_WEBSITE. 04e: surveyUrl now reads MOBILE_DASHBOARD_URL. 05_Integrations: satisfactionForm removed from resource links, double-semicolon typo fixed. 11_CommandHub: Form Links color section removed. 20_WebDashConfigReader: satisfactionFormUrl line removed. 21_WebDashDataService: surveyFormUrl removed from resource links response. 22_WebDashApp: surveyFormUrl removed from sanitized config. member_view.html: dead else-if fallback branch removed from survey banner onClick handler.' },
  { version: '4.22.6', date: '2026-03-06', codename: 'Org Chart Default', changes: 'Replace org_chart.html with embedded org chart. Full org chart is now default in SolidBase. Conversion: CSS scoped to .madds-embed wrapper, :root vars moved into .madds-embed scope, body/body.light → .madds-embed/.madds-embed.light, #mode-toggle renamed #madds-mode-toggle, toggleMode() renamed maddstoggleMode() to avoid SPA collisions, Google Fonts loaded dynamically. Serves same HTML fragment via existing getOrgChartHtml() in 22_WebDashApp.gs — no server-side changes required.' },
  { version: '4.22.6', date: '2026-03-06', codename: 'Events Sentinel Propagation Fix', changes: 'Bug fix: home widget events section crashed when getUpcomingEvents returned a sentinel object ({_notConfigured} or {_calNotFound}) — events.length on a plain object returns undefined, rendering "undefined" in KPI counter and crashing forEach. Added Array.isArray guard before rendering and before DataCache.set. DataCache.set now only caches actual arrays; sentinel objects are dropped, preventing the bad value from poisoning the client-side cache on subsequent home re-renders.' },
  { version: '4.22.4', date: '2026-03-06', codename: 'Events Access & Calendar Targeting', changes: 'Events tab dues gate removed — any authenticated member can view events regardless of dues status. More menu Events item lock icon removed. Create Event button URL now includes &src=calendarId param so new events land on the union calendar, not the steward personal calendar. No structural changes to backend auth — dataGetUpcomingEvents still requires valid session.' },
  { version: '4.22.3', date: '2026-03-06', codename: 'Events Tab Hardening', changes: 'Bug fix: ISO date formatter in Add-to-Calendar URLs changed from .replace(\".000Z\",\"Z\") to .replace(/\\.\\d+Z$/,\"Z\") — handles any millisecond value (was silently broken for non-.000 ms values). Bug fix: CalendarApp.getCalendarById() returning null now returns {_calNotFound:true} sentinel instead of [] — distinguishes typo/permission error from genuinely empty calendar. Frontend handles _calNotFound with diagnostic message. Bug fix: Add-to-Calendar URL now includes &details= param (ev.description) in both home widget and Events page — was silently omitted. Feature: Steward Events page now shows \"Manage in Google Calendar\" and \"Create Event\" action buttons when a calendarId is configured. _sanitizeConfig in 22_WebDashApp.gs now exposes calendarId (non-sensitive) so frontend can conditionally show management links.' },
  { version: '4.22.2', date: '2026-03-06', codename: 'Notification Cleanup Pass', changes: 'Bug fix: steward inbox dismiss was passing stale second arg (CURRENT_USER.email) to dismissWebAppNotification() — function only accepts notificationId, gets identity from Session server-side. Arg removed. Dead empty2 block removed from member_view.html renderMemberNotifications — visible=notifications made the unreachable second empty-state check redundant. 04e_PublicDashboard.gs: inline JS notification calls updated from deleted getUserNotifications/markNotificationRead to getWebAppNotifications. Dead DataService.getMemberNotifications guard replaced with [] constant and comment. All four issues are in 04e which is confirmed orphaned (not routed in doGet).' },
  { version: '4.22.9', date: '2026-03-07', codename: 'FailsafeService Session Token Auth Fix', changes: 'FIX: Auth broken for magic link / session token users in Execute-as-Me webapp. Root cause: Session.getActiveUser() returns empty in this deployment mode for non-SSO users. Fix: Auth.resolveEmailFromToken() exposed in Auth public API. _resolveCallerEmail(sessionToken) now tries SSO first then verifies session token server-side. checkWebAppAuthorization(requiredRole, sessionToken) accepts optional token fallback. _requireStewardAuth(sessionToken) passes token through. All 8 fs* global wrappers updated to accept sessionToken param. Client calls in steward_view.html (3 trigger/backup buttons) and member_view.html (2 digest calls) updated to pass SESSION_TOKEN instead of CURRENT_USER.email. Security preserved: raw email never trusted, only server-validated token resolves identity.' },
  { version: '4.22.8', date: '2026-03-07', codename: 'FailsafeService Security & Reliability Fixes', changes: 'SEC: fsGetDigestConfig/fsUpdateDigestConfig no longer accept client-supplied email — server-resolved identity only, no fallback to client param. SEC: fsBackupCriticalSheets, fsSetupTriggers, fsRemoveTriggers, fsInitSheets now require _requireStewardAuth(). RELIABILITY: processScheduledDigests wraps send+lastSent-write in per-member ScriptLock with double-check re-read to prevent duplicate digests from concurrent trigger executions. STORAGE: backupCriticalSheets prunes old CSVs via _pruneOldBackups() keeping MAX_BACKUP_FILES=52 per sheet using file.setTrashed(). HEADER: stale @version 4.17.0 updated to 4.22.8.' },
  { version: '4.22.1', date: '2026-03-06', codename: 'Notification Manage Hardening', changes: 'Migration function MIGRATE_ADD_DISMISS_MODE_COLUMN() added to 05_Integrations.gs — safe one-time runner that appends Dismiss_Mode header to existing Notifications sheet and backfills all rows with "Dismissible". New archiveWebAppNotification(notificationId) in 05_Integrations.gs — steward-auth-gated, sets Status=Archived, non-destructive. Archive button added to each Active row in steward Manage tab (steward_view.html) — updates status pill in-place, hides button after success. Expired/Archived rows remain visible in Manage tab for auditing.' },
  { version: '4.22.0', date: '2026-03-06', codename: 'Notification System Overhaul', changes: 'Bug fix: sort in getWebAppNotifications() now reverses first then sorts by priority — previously Urgent notifications landed at the bottom. New DISMISS_MODE column (Dismissible | Timed) added to NOTIFICATIONS_HEADER_MAP_ (col 13). Dismissible: member can permanently dismiss (writes to Dismissed_By column). Timed: auto-expires on Expires_Date, dismiss button hidden, "Auto-expires" badge shown to member. Compose form gets dismiss mode toggle with Timed validation requiring an expiry date. Member dismiss changed from 1-hour localStorage TTL to permanent backend write via dismissWebAppNotification(). Bell badge DOM re-renders correctly on dismiss. Manage tab fixed: now calls getAllWebAppNotifications() (new function, steward-auth-gated) returning all rows with dismissedCount + status — was incorrectly calling getWebAppNotifications() which filtered to the steward\'s own inbox. Dead code removed: getUserNotifications(), markNotificationRead(), broadcastStewardNotification() (all ScriptProperties-based, orphaned since v4.13.0 Notifications sheet). getWebAppNotificationsHtml() removed — 345 lines of standalone ?page=notifications HTML never routed in doGet(). pushNotification() rerouted from ScriptProperties to Notifications sheet (still called by saveSharedView()).' },
  { version: '4.21.0', date: '2026-03-05', codename: 'Native Survey Engine', changes: 'Deprecate Google Form integration entirely. Full webapp-native satisfaction survey: getSurveyQuestions() returns all 67 questions with types (slider-10, dropdown, radio, checkbox, paragraph), branching rules (Q5→3A/3B, Q36→6A), and slider labels. submitSurveyResponse() maps all 67 SATISFACTION_COLS, period-aware vault dedup per SURVEY_PERIODS sheet. New: getSurveyPeriod(), autoTriggerQuarterlyPeriod(), archiveSurveyPeriod_() (Drive export to Past Survey Questions/), getPendingSurveyMembers(), getSatisfactionSummary() (section averages as plain values). New hidden sheet _Survey_Periods (SURVEY_PERIODS_COLS, 8 cols). New Config cols: Survey Priority Options (Q64 dynamic), Past Surveys Folder ID. Quarterly time trigger installed via setupQuarterlyTrigger(). Survey-open notifications pushed to all active members on period start. Slider label: 1=Strongly Disagree / 10=Strongly Agree (universal across all 52 scale questions). Anonymity architecture preserved: three-layer separation (Satisfaction sheet=anonymous scores, _Survey_Vault=hashed PII only, _Survey_Tracking=completion status only).' },
  { version: '4.20.15', date: '2026-03-05', codename: 'FULL_CODE_REVIEW Final Fixes', changes: 'C-XSS-18: Fix el() boolean attribute handling — el() now uses property assignment (elem[key]=value) for boolean attrs (selected, disabled, checked) instead of setAttribute which would set attr="false" (truthy in DOM). C-XSS-6: Replace escapeHtml() with JSON.stringify() for memberId in onclick JS string context in 03_UIComponents.gs — HTML entities decoded by parser before JS executes, JSON.stringify produces correct escape sequences. LOW: Remove 6 unused _-prefixed variables (_lastRow, _pdfFile, _headers×2, _stepDays, _mgmtResponseDays, _mode, _ss×2) from 04b/04d/04e/05_Integrations.' },
  { version: '4.20.14', date: '2026-03-05', codename: 'Trigger Null Guards', changes: 'onOpenDeferred_ (10_Main.gs): add null guard after getActiveSpreadsheet() — returns null in web app context, would crash ss.toast(). onEditWithAuditLogging (06_Maintenance.gs): add !e || !e.range early return guard and wrap body in try/catch — trigger functions must not throw or GAS silently drops all subsequent edits.' },
  { version: '4.20.13', date: '2026-03-04', codename: 'Accessibility & Config Hardening', changes: 'Accessibility (WCAG 2.1 SC 1.4.4): Replace user-scalable=no with user-scalable=yes,maximum-scale=5.0 in all 12 viewport meta tags across 5 files (index.html, 04c, 04e, 05_Integrations x8, 14_MeetingCheckIn). Config hardening: Replace 4 hardcoded org names with getConfigValue_(CONFIG_COLS.ORG_NAME). Replace 3 hardcoded sheet name strings in 21_WebDashDataService.gs (2x _Survey_Tracking → HIDDEN_SHEETS.SURVEY_TRACKING, Config → SHEETS.CONFIG). Remove redundant || ss.getSheetByName("_Dashboard_Calc") fallback in 04d_ExecutiveDashboard.gs. Modernize document.execCommand("copy") in 08c_FormsAndNotifications.gs (2 locations) with navigator.clipboard.writeText() + execCommand fallback for older environments.' },
  { version: '4.20.12', date: '2026-03-04', codename: 'Dead Code Removal — Testing Stubs', changes: 'Remove duplicate testing framework section from 06_Maintenance.gs (~84 lines): TEST_RESULTS var (zero callers), TEST_MAX_EXECUTION_MS/TEST_LARGE_DATASET_THRESHOLD (duplicates of 07_DevTools.gs), Assert object (duplicate of 07_DevTools.gs), 4 empty section headers, VALIDATION_PATTERNS/VALIDATION_MESSAGES (duplicates of 07_DevTools.gs), 2 orphaned JSDoc stubs with no function bodies. Active versions of all these remain in 07_DevTools.gs. 16 corresponding test assertions removed from 06_Maintenance.test.js.' },
  { version: '4.20.11', date: '2026-03-04', codename: 'Security & Logic Fixes', changes: 'H-12: Replace vault/reminders/userMeta/archive hideSheet() with setSheetVeryHidden_() in 25_WorkloadService.gs — API-level hide that persists on Google Sheets mobile (previously only UI-layer hide). H-16: Fix contact log sort in getMemberContactHistory and getStewardContactLog (21_WebDashDataService.gs) — was sorting by formatted date string (wrong alphabetical order); now stores _ts timestamp field, sorts numerically, deletes _ts before returning. H-20: Fix win rate denominator in getStewardWorkloadDetailed (02_DataManagers.gs) — was dividing wonCases by totalCases (includes active cases); now uses resolvedCases = Won+Denied+Settled+Withdrawn, matching the formula in getDashboardStats.' },
  { version: '4.20.10', date: '2026-03-04', codename: 'Performance & Data Integrity Fixes', changes: 'H-7: Batch 4 individual configSheet.getValue() calls in getDeadlineRules() into a single getValues() range read — 4 API calls → 1. H-2: sortGrievanceLogByStatus now captures cell notes via getNotes() before sort and restores them via setNotes() after setValues() — preserves user-added notes that would otherwise be silently discarded by the sort.' },
  { version: '4.20.9', date: '2026-03-04', codename: 'Hardcoded Index Fix & Dead Code', changes: 'C-DATA-6: Replace positional array indices (sections[6], sections[7], questions[0]–[4]) in satisfaction survey score assignment with key-based lookup maps (sectionByKey[key], s._qByKey[key]) — safe against section reordering. Dead code: emailDashboardLink_UIService_ (03_UIComponents.gs, deprecated, zero callers), getOrCreateMemberFolder_Legacy_ (11_CommandHub.gs, deprecated, zero callers). Note: DataCache.cachedCall deduplicates surveyStatus/events server calls by cache key — those are not real duplicates.' },
  { version: '4.20.8', date: '2026-03-04', codename: 'Dead Code Removal & Performance', changes: 'H-13: setupHiddenSheets skips sheet.clear() for SURVEY_TRACKING when data rows exist — prevents accidental wipe of survey history. Dead code removal: DataAccess namespace (~530 lines), validateInputLength_, getCurrentUserEmail, safeJsonForHtml, sanitizeDataForClient (00_Security.gs), getWebAppDashboardHtml (05_Integrations.gs), isMobileContext (03_UIComponents.gs), statStdDev_ (17_CorrelationEngine.gs). Performance: updateChecklistItem batches per-field setValue calls into single setValues row write. updateMemberProfile batches per-field setValue into single setValues row write. applyConfigSheetStyling replaces per-row setBackground loop with single setBackgrounds() call. 00_DataAccess.gs retains TIME_CONSTANTS and withScriptLock_ (both used in production).' },
  { version: '4.20.7', date: '2026-03-04', codename: 'Security & Data Integrity Fixes', changes: 'C-AUTH-4: dataGetMemberGrievanceHistoryPortal now resolves identity server-side (was IDOR). C-XSS-5: JSON.stringify+&quot; on category onclick param in 05_Integrations.gs. C-FORMULA-7: escapeForFormula on weekly question text in 24_WeeklyQuestions.gs (3 locations). C-DATA-1: vault dedup reads EMAIL column (col 2) not RESPONSE_ROW (col 1) — dedup was silently broken. C-DATA-5: weekly_cases manual fallback removed hardcoded \'15\' in member_view.html. H-5: LockService added to meeting check-in to prevent TOCTOU duplicate entries (14_MeetingCheckIn.gs). H-17: email regex validation before overdue report send (04d_ExecutiveDashboard.gs). H-18: LockService added to archiveClosedGrievances for atomic read-copy-delete (06_Maintenance.gs). H-3: applyState ADD_ROW bounds check (row > 1 and <= lastRow) before deleteRow. H-4: batchSetValues guards against header row (row <= 1) updates.' },
  { version: '4.20.6', date: '2026-03-04', codename: 'Critical Security Fixes', changes: 'Fix 9 confirmed vulnerabilities from FULL_CODE_REVIEW: C-AUTH-5 (getEffectiveUser→getActiveUser in 3 locations), C-AUTH-7 (PII guard || → && logic fix), C-XSS-7/8 (escapeHtml on custom field value+label in 12_Features.gs), C-XSS-9 (JSON.stringify+&quot; in onclick builder in 11_CommandHub.gs), C-XSS-14 (escapeHtml on renderGauge label in 09_Dashboards.gs), C-XSS-16 (escapeHtml in PublicDashboard HTML table; RFC 4180 double-quote CSV escaping), C-XSS-17 (email regex validation in 03_UIComponents.gs), C-OTHER-1 (delete dead buildSafeQuery function from 00_Security.gs). Add regression tests for C-AUTH-7 (+2) and C-XSS-17 (+9).' },
  { version: '4.20.5', date: '2026-03-04', codename: 'XSS Fix OCR Dialog', changes: 'Wrap escapeHtml() around API key suffix in setupOCRApiKey() dialog (11_CommandHub.gs:2441) — raw key material was injected into HTML string without escaping.' },
  { version: '4.20.4', date: '2026-03-04', codename: 'Regression Test Hardening', changes: 'Add 138 regression tests targeting known failure modes: N+1 sheet reads (getStewardSurveyTracking spy test), boolean normalization via isTruthyValue() (anonymous flag QAForum ×20, getAllStewards IS_STEWARD ×15, vault VERIFIED/IS_LATEST ×11), formula injection protection (approveFlaggedSubmission, rejectFlaggedSubmission, addToConfigDropdown_ ×8 total), sendEmailToMember auth/safeSubject/safeBody (×6). Architecture tests A16 (lock→finally contract across 8 files), A17 (lock-acquiring mutations log audit events), A18 (dataXxx wrappers call DataService — 56 wrappers tested). Fix formula injection bugs in approveFlaggedSubmission, rejectFlaggedSubmission, and addToConfigDropdown_. Total: 2083/2083 tests pass (+138 from 1945).' },
  { version: '4.20.3', date: '2026-03-04', codename: 'Code Review Fixes', changes: 'Fix N+1 sheet reads in getStewardSurveyTracking — pre-load _Survey_Tracking once, build email map, O(1) lookup per member. Add escapeForFormula() to profile update setValue() (formula injection fix). Replace all google.script.run with serverCall() in member_view.html and steward_view.html (~52 calls) — all server calls now have default failure handler. Normalize QAForum anonymous checks to use isTruthyValue() for consistent Sheets boolean handling.' },
  { version: '4.20.2', date: '2026-03-04', codename: 'Web App Error Resilience', changes: 'Fix 14 missing withFailureHandler() on google.script.run calls in member_view.html (11) and steward_view.html (2) — prevents infinite loading spinners on server errors. Add null guards on getActiveSpreadsheet() in 26_QAForum.gs (10), 27_TimelineService.gs (7), 28_FailsafeService.gs (7) — prevents "Cannot call method of null" crashes in web app context.' },
  { version: '4.20.1', date: '2026-03-03', codename: 'Test Suite 100% Pass', changes: 'Fix all 40 pre-existing test failures: null guards on getActiveSpreadsheet() in 21_WebDashDataService.gs (17), 25_WorkloadService.gs (17), 24_WeeklyQuestions.gs (1+_ensureSheet early-return); PropertiesService singleton mock fix (16_DashboardEnhancements); Session/CacheService mock fixes (19_WebDashAuth); EventBus SHEETS reverse-map fix (15_EventBus); DataService API alignment in 21_WebDashDataService tests; A12 threshold updated to 130; A13 7 failure handlers added to HTML views. 1945/1945 tests pass.' },
  { version: '4.20.0', date: '2026-03-03', codename: 'WorkloadTracker SPA Integration', changes: 'Remove standalone WorkloadTracker portal (18_WorkloadTracker.gs, WorkloadTracker.html). Workload tracker fully integrated into SPA via 25_WorkloadService.gs and member_view.html. Route ?page=workload deep-links to SPA workload tab after SSO auth. Merge v4.19.2-v4.19.5 error resilience hardening: fatal error guard in doGet(), null guards on getActiveSpreadsheet(), trigger try/catch, serverCall() client wrapper, 535 new unit tests (1146→1681).' },
  { version: '4.19.5', date: '2026-03-03', codename: 'Full Coverage Expansion', changes: '535 new unit tests across 14 test files covering all previously untested source modules: EventBus (15), DashboardEnhancements (16), CorrelationEngine (17), WorkloadTracker (18), WebDashAuth (19), WebDashConfigReader (20), WebDashDataService (21), WebDashApp (22), PortalSheets (23), WeeklyQuestions (24), WorkloadService (25), QAForum (26), TimelineService (27), FailsafeService (28). Test count increased from 1146 to 1681 (+46.6%). Gas-mock.js enhanced with createTemplateFromFile, MimeType, ScriptApp.WeekDay, MailApp.getRemainingDailyQuota, and corrected XFrameOptionsMode enum.' },
  { version: '4.19.4', date: '2026-03-03', codename: 'Regression Test Suite', changes: 'Architecture tests A9-A15 covering UI tab wiring, formula injection, auth enforcement, XSS prevention, google.script.run failure handler tracking, GAS enum validation, and error handler cascade detection. Fixed invalid XFrameOptions DENY enum usage in 00_Security.gs (DENY is not a valid GAS enum value — changed to DEFAULT).' },
  { version: '4.19.3', date: '2026-03-03', codename: 'Error Resilience Hardening', changes: 'Null guards on getActiveSpreadsheet() in all web app chain files (ConfigReader, DataService, PortalSheets, WeeklyQuestions, WorkloadService). Try/catch on onEditMultiSelect and onSelectionChangeMultiSelect trigger handlers. Client-side serverCall() wrapper with default withFailureHandler. Architecture tests A6-A8 enforce null safety, trigger try/catch, and failure handlers. CLAUDE.md error handling rules.' },
  { version: '4.19.2', date: '2026-03-03', codename: 'Fatal Error Guard', changes: 'Top-level try/catch in doGet() prevents generic Google "unable to open file" page. Safe config fallback in doGetWebDashboard error handler avoids error cascade when ConfigReader throws. Minimal _serveFatalError() page with zero external dependencies.' },
  { version: '4.19.1', date: '2026-03-02', codename: 'Org Chart Wiring', changes: 'Implement missing renderOrgChart() function — Org. Chart tab was throwing JS error on click. Renamed tab label to "Org. Chart" in both steward and member sidebars. Script re-execution for org_chart.html interactive toggles.' },
  { version: '4.19.0', date: '2026-03-02', codename: 'QA Bug Fixes & Resilience', changes: 'Server-side error handling for all DataService methods (Issues 1-7). Sign-out fix returns to login page (Issue 10). Member detail panel with expand/collapse and Full Profile loading (Issue 8). By Location chart falls back to all members when none assigned (Issue 9). Contact log autocomplete failure handler (Issue 11). Auto-initialize QA Forum and Timeline sheets on first access (Issue 12). Empty state messages and failure handlers for Events and Weekly Questions.' },
  { version: '4.18.0', date: '2026-02-26', codename: 'SPA Fixes, Seed Phasing & View Enhancements', changes: 'Split SEED_SAMPLE_DATA into 3 phased runners to avoid GAS 6-min timeout. 5 new seed functions (tasks, polls, minutes, check-ins, timeline events). Steward view: org-wide KPI fallback, all-contacts members tab, comma formatting, contact log autocomplete, survey tracking scope toggles, 6 new More menu items. Member view: Know Your Rights card, Contact-Directory nav, 1hr localStorage notification dismiss, meetings+minutes merge, 7 new More menu items. Backend globals: getAllMembers, startGrievanceDraft, createGrievanceDriveFolder. Broadcast uses all contacts. build.js BUILD_ORDER updated for 26_QAForum.gs, 27_TimelineService.gs, 28_FailsafeService.gs.' },
  { version: '4.17.0', date: '2026-02-26', codename: 'Q&A Forum, Timeline & Failsafe Services', changes: 'Q&A Forum (26_QAForum.gs, 389 lines) with _QA_Forum and _QA_Answers hidden sheets. Timeline Service (27_TimelineService.gs, 317 lines) with _Timeline_Events hidden sheet. Failsafe Service (28_FailsafeService.gs, 425 lines) with _Failsafe_Config hidden sheet. 08a_SheetSetup.gs updated for Q&A, Timeline, and Failsafe auto-creation. DataService methods for Q&A, Timeline, and Failsafe in 21_WebDashDataService.gs.' },
  { version: '4.16.0', date: '2026-02-26', codename: 'Wire 7 Unwired Sheets to SPA', changes: '15 new DataService methods (541 lines) in 21_WebDashDataService.gs wiring 7 previously unwired sheets to SPA. 15 global wrapper functions + 3 batch data fields. New SPA pages: Meetings, Polls, Minutes, Feedback. Insights page with Performance KPIs + Satisfaction Trends. Case detail views with checklist support. Per-question text scores with color-coding. questionTexts arrays for all 11 SATISFACTION_SECTIONS. Expansion test suite (332 lines). Removed Since N/A text, Dues Status charts. Fixed 122 test failures (1,363 tests passing across 23 suites).' },
  { version: '4.15.0', date: '2026-02-25', codename: 'Phase 7: Login, Surveys, Steward Management & Seed Enhancements', changes: 'Infrastructure: batch fetch, Drive cleanup trigger, calendar dedup, CC health check, lazy-load help dialog, search pagination, expansion test suite. Login UX: SSO loading state, sso_failed fallback, magic link clarification, resend cooldown. In-app survey wizard: multi-step mobile-optimized form with localStorage progress, 1-10 scale buttons, anonymous SHA-256 submission. Steward: chief steward task assignment, agency-wide grievance stats fallback, Insights tab (Quick Insights + Filed vs Resolved chart), Steward Directory with vCard download. Member dashboard: actionable KPI strip, conditional grievance card, engagement/workload stats tabs. Broadcast: checkbox pill filters with recipient preview. Workload: removed Private option. Seed data: calendar events, weekly questions, union stats.' },
  { version: '4.14.0', date: '2026-02-25', codename: 'Technical Debt Resolution & PHASE2 Features', changes: '130 code review findings resolved (15 CRITICAL XSS, 26 HIGH security, 50 MEDIUM, 39 LOW). 5 new features: Grievance History, Meeting Check-In Kiosk, Welcome Experience, Bulk Actions, Deadline Calendar View. Engagement sync overhaul with dynamic headers and validation. withScriptLock_() concurrency helper. safeSendEmail() quota wrapper. Version derived from single COMMAND_CONFIG.VERSION source.' },
  { version: '4.13.0', date: '2026-02-24', codename: 'Full Workload Tracker Migration', changes: 'Refactored 18_WorkloadTracker.gs to IIFE module (WorkloadService), enhanced getDashboardData with employment/plan/overtime breakdowns and sub-category aggregation, enhanced getUserHistory with all 24 columns, CSV export, vault deduplication, reciprocity blocking for Private users, multi-frequency reminders (daily/weekly/biweekly/monthly/quarterly), full leave tracking in portal, Weekly Cases dropdown, Clear All/Restore.' },
  { version: '4.12.0', date: '2026-02-24', codename: 'Version Alignment', changes: 'API_VERSION, VERSION_INFO, VERSION_HISTORY normalized. README and CODE_REVIEW updated with correct file counts and version scope.' },
  { version: '4.11.0', date: '2026-02-24', codename: 'Web Dashboard SPA Enhancement', changes: 'Responsive sidebar/bottom-nav layout, member and steward views with full render functions, WeeklyQuestions IIFE module, 12 new DataService functions, SSO workload wrappers, ConfigReader expansion.' },
  { version: '4.10.0', date: '2026-02-24', codename: 'Workload Tracker Integration', changes: 'Workload Tracker module (18_WorkloadTracker.gs + WorkloadTracker.html), 8 workload categories with sub-breakdowns, privacy controls, reciprocity enforcement, email reminders, 24-month data retention, CSV backup, Workload Tracker submenu in Union Hub, ?page=workload web route, 5 new hidden sheets.' },
  { version: '4.9.1', date: '2026-02-23', codename: 'Security Vulnerability Fix Pass', changes: 'Fix 15 broken getClientSideEscapeHtml() includes, escape member data in grievance form HTML templates, URL scheme validation on Config URLs, escape steward contact data in Public Dashboard, replace unsafe onclick injection, add email format validation, formula injection protection, server-side input validation.' },
  { version: '4.9.0', date: '2026-02-17', codename: 'Constant Contact Integration', changes: 'Constant Contact v3 API integration with OAuth2, multi-select dropdown support for Grievance Log, auto-discovery column system, 151 column system tests, dynamic CONFIG_COLS and MEMBER_COLS constants.' },
  { version: '4.8.2', date: '2026-02-16', codename: 'State Field', changes: 'State field added to member contact update across all surfaces.' },
  { version: '4.8.1', date: '2026-02-15', codename: 'Contact Fields Expansion', changes: '5 new contact form fields (Hire Date, Employee ID, Street Address, City, Zip Code), unified name-based Member ID system.' },
  { version: '4.8.0', date: '2026-02-15', codename: 'Security Event Alerting', changes: 'Security event alerting system, zero-knowledge survey vault with SHA-256 hashes, event bus architecture, survey completion tracker.' },
  { version: '4.7.0', date: '2026-02-14', codename: 'Security Hardening', changes: '40+ code review fixes across security, correctness, performance, and test quality. XSS hardening, onEdit optimization, deduplicated escapeHtml, 1090 tests.' },
  { version: '4.6.0', date: '2026-02-12', codename: 'Meeting Intelligence & Document Automation', changes: 'Meeting Notes & Agenda doc automation, two-tier steward agenda sharing, Meeting Notes dashboard tab, member Drive folders, meeting event scheduling. Added Employee ID, Department, Hire Date columns to Member Directory. Added PII mailing address columns (Street, City, State) hidden by default. Added Last Updated to Grievance Log. Fixed diagnostics checks. Removed deprecated Dashboard/Satisfaction from sheet ordering. Added Export (org email only) and Lockdown future feature roadmap items.' },
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
// ============================================================================
// SHEET NAMES
// ============================================================================

/**
 * Sheet name constants - use these instead of hardcoded strings.
 * L-05: Some sheet names contain emoji (e.g., "📅 Meeting Attendance").
 * Emoji in sheet names may cause issues on some platforms/locales but is fully
 * supported by Google Sheets. If cross-platform compatibility is needed, the
 * emoji prefix can be removed without affecting functionality.
 * @const {Object}
 */
var SHEETS = {
  CONFIG: '⚙️ Config',
  MEMBER_DIR: '👥 Member Directory',
  GRIEVANCE_LOG: '⚖️ Grievance Log',
  GRIEVANCE_ARCHIVE: '_Archive_Grievances', // v4.30.0 — closed/resolved cases, very hidden
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
  TEST_RESULTS: '🧪 Test Results',
  // Function Checklist
  FUNCTION_CHECKLIST: '🔧 Function Checklist',
  // Audit Log (hidden)
  AUDIT_LOG: '_Audit_Log',
  // Case Checklist
  CASE_CHECKLIST: '✅ Case Checklist',
  // Survey Completion Tracking (hidden) - tracks per-member survey participation.
  // See SURVEY_TRACKING_COLS below for column layout.
  // Detection flow: Google Form submit -> onSatisfactionFormSubmit() -> validateMemberEmail()
  //   -> updateSurveyTrackingOnSubmit_() marks member "Completed" in this sheet.
  // Management: showSurveyTrackingDialog() in 08c_FormsAndNotifications.gs
  SURVEY_TRACKING: '_Survey_Tracking',
  // Survey Vault (hidden + protected) — stores SHA-256 hashed email/member ID
  // for survey responses. No plaintext PII. Separated from Satisfaction sheet for anonymity.
  SURVEY_VAULT: '_Survey_Vault',
  // Satisfaction & Feedback sheets
  // @deprecated v4.3.8 - Satisfaction sheet is now hidden. Use showSatisfactionDashboard() modal instead.
  // Data is preserved for modal access. Use removeDeprecatedTabs() to hide.
  SATISFACTION: '📊 Member Satisfaction',
  SURVEY_QUESTIONS: '📋 Survey Questions',
  FEEDBACK: '💡 Feedback & Development',
  // Help & Documentation sheets
  GETTING_STARTED: '📚 Getting Started',
  FAQ: '❓ FAQ',
  CONFIG_GUIDE: '📖 Config Guide',
  // Aliases for backward compatibility (some code uses these alternate names)
  // NOTE: GRIEVANCE_TRACKER alias removed in v4.25.9 (FIX-CORE-02) — all 22
  //   callers migrated to canonical SHEETS.GRIEVANCE_LOG. Sheet is unchanged.
  MEMBER_DIRECTORY: '👥 Member Directory',
  // Workload Tracker sheets (18_WorkloadTracker.gs)
  WORKLOAD_VAULT:     'Workload Vault',      // hidden — raw submissions with email
  WORKLOAD_REPORTING: '📊 Workload Reporting',  // visible — anonymized ledger
  WORKLOAD_REMINDERS: 'Workload Reminders',  // hidden — email reminder prefs
  WORKLOAD_USERMETA:  'Workload UserMeta',   // hidden — sharing start dates
  WORKLOAD_ARCHIVE:   'Workload Archive',     // hidden — data older than 24 months
  // Resources & Knowledge (v4.11.0 — content management for educational hub)
  RESOURCES:          '📚 Resources',         // steward-managed educational content
  RESOURCE_CONFIG:    '📚 Resource Config',   // categories & settings for Resources tab (v4.22.x)
  KNOWLEDGE_ENGINE:   '📚 Knowledge Engine',  // v4.41.0 — centralized quotes, tips, concepts, mini-lessons
  // v4.32.1 — Hidden sheet for resource click/view analytics.
  // Schema: [Timestamp, User Email, Resource ID, Resource Title] (4 columns).
  // Auto-created by dataLogResourceClick() on first use; no manual setup needed.
  // Read by dataGetEngagementStats() to compute the "Resource Views" KPI.
  // Prefixed with underscore = hidden sheet convention (see also _Contact_Log, _Audit_Log).
  // IF MISSING: dataLogResourceClick() creates it automatically. dataGetEngagementStats()
  // returns resourceDownloads=0 (the _rows() helper returns [] for missing sheets).
  RESOURCE_CLICK_LOG: '_Resource_Click_Log',
  // Weekly Questions (24_WeeklyQuestions.gs) — anonymous pulse surveys
  WEEKLY_QUESTIONS:   '_Weekly_Questions',    // hidden — active/scheduled questions
  WEEKLY_RESPONSES:   '_Weekly_Responses',    // hidden — SHA-256 hashed anonymous responses
  QUESTION_POOL:      '_Question_Pool',       // hidden — reusable question bank
  // Notifications (v4.13.0 — SPA in-app notification system)
  NOTIFICATIONS:      '📢 Notifications',     // steward-to-member in-app messages
  // Contact Log & Steward Tasks (v4.12.0) — steward activity tracking
  CONTACT_LOG:        '_Contact_Log',         // hidden — steward-member contact history
  STEWARD_TASKS:      '_Steward_Tasks',       // hidden — task assignments for stewards
  // Q&A Forum (v4.17.0 — member-steward question/answer system)
  QA_FORUM:           '_QA_Forum',            // hidden — member questions
  QA_ANSWERS:         '_QA_Answers',          // hidden — steward/member answers
  // Timeline of Events (v4.17.0 — chronological event records)
  TIMELINE_EVENTS:    '_Timeline_Events',     // hidden — event timeline entries
  TIMELINE_CATEGORIES: '_Timeline_Categories', // hidden — steward-managed event categories
  // Data Failsafe (v4.17.0 — member digest preferences)
  FAILSAFE_CONFIG:    '_Failsafe_Config',     // hidden — digest/backup preferences
  // Portal-visible sheets (23_PortalSheets.gs)
  PORTAL_EVENTS:       '🗓️ Events',             // visible — union events calendar entries
  PORTAL_MINUTES:      '🗒️ Meeting Minutes',   // visible — meeting minutes archive
  // Reference & admin sheets (may be manually created)
  FEATURES_REFERENCE:  '📋 Features Reference', // visible — searchable feature catalog
  SETTINGS_OVERVIEW:   '🔍 Settings Overview',     // visible — live config snapshot (if present)
  NON_MEMBER_CONTACTS: '📇 Non-Member Contacts',    // visible — management contacts (if present)
  // New Feature Services (v4.35.0 — 33_NewFeatureServices.gs)
  HANDOFF_NOTES:      '_Handoff_Notes',          // hidden — steward shift handoff notes per case
  MENTORSHIP:         '_Mentorship',             // hidden — steward mentorship pairings
  COMMUNICATION_LOG:  '_Communication_Log',      // hidden — member outreach tracking
  KNOWLEDGE_BASE:     '_Knowledge_Base',         // hidden — contract article reference
  NOTIFICATION_PREFS: '_Notification_Prefs',      // hidden — digest/notification preferences
  // SMS & RSVP (v4.36.0 — 33_NewFeatureServices.gs)
  SMS_LOG:            '_SMS_Log',                 // hidden — outbound SMS log (Twilio)
  RSVP_LOG:           '_RSVP_Log',                // hidden — meeting RSVP tracking
  // Trend Alerts & Engagement (v4.36.0)
  TREND_ALERTS:       '_Trend_Alerts',            // hidden — auto-detected trend alerts
  ENGAGEMENT_SCORES:  '_Engagement_Scores',       // hidden — per-member engagement scores
  SHIRT_SIZE_LOG:     '👕 Shirt Size Log'          // visible — member shirt sizes with name, location, unit, steward
};

/**
 * Legacy sheet names — maps old (pre-emoji) tab names to their new SHEETS constant values.
 * Used by migrateSheetTabTitles_() and getOrCreateSheet() to auto-rename existing sheets.
 * @const {Object.<string, string>}
 */
var SHEET_LEGACY_NAMES_ = {
  'Config':               SHEETS.CONFIG,
  'Member Directory':     SHEETS.MEMBER_DIR,
  'Grievance Log':        SHEETS.GRIEVANCE_LOG,
  'Case Checklist':       SHEETS.CASE_CHECKLIST,
  'Test Results':         SHEETS.TEST_RESULTS,
  'Function Checklist':   SHEETS.FUNCTION_CHECKLIST,
  'Workload Reporting':   SHEETS.WORKLOAD_REPORTING,
  'Events':               SHEETS.PORTAL_EVENTS,
  'MeetingMinutes':       SHEETS.PORTAL_MINUTES,
  'Settings Overview':    SHEETS.SETTINGS_OVERVIEW,
  'Non member contacts':  SHEETS.NON_MEMBER_CONTACTS
};

/**
 * Hidden sheets used for calculations (prefixed with underscore)
 * These are auto-created and hidden from users
 * @const {Object}
 */
var HIDDEN_SHEETS = {
  CALC_STATS: '_Dashboard_Calc',
  CALC_FORMULAS: '_Grievance_Formulas',
  CALC_MEMBERS: '_Members_Calc',
  CALC_GRIEVANCES: '_Grievances_Calc',
  CALC_DEADLINES: '_Deadlines_Calc',
  CALC_SYNC: '_Sync_Calc',
  GRIEVANCE_CALC: '_Grievance_Calc',
  MEMBER_LOOKUP: '_Member_Lookup',
  STEWARD_CONTACT_CALC: '_Steward_Contact_Calc',
  STEWARD_PERFORMANCE_CALC: '_Steward_Performance_Calc',
  AUDIT_LOG: '_Audit_Log',
  CHECKLIST_CALC: '_Checklist_Calc',
  SURVEY_TRACKING: '_Survey_Tracking',
  SURVEY_VAULT: '_Survey_Vault',
  SURVEY_PERIODS: '_Survey_Periods',
  ARCHIVE_GRIEVANCES: '_Archive_Grievances',
  USAGE_LOG: '_Usage_Log'              // v4.40.0: production usage tracking (session time, tabs, perf, errors)
};

// ============================================================================
// HIDDEN SHEET ENFORCEMENT - Mobile-safe hiding via Sheets API
// ============================================================================
// Google Sheets mobile can display sheets hidden with hideSheet().
// These helpers use the Sheets Advanced Service (API v4) to set hidden=true
// at the API level AND apply sheet protection so the sheets stay invisible
// and uneditable even on mobile devices.
// Requires: Sheets Advanced Service enabled in appsscript.json

/**
 * Protection description used to identify auto-protections on hidden sheets.
 * @const {string}
 * @private
 */
var HIDDEN_SHEET_PROTECTION_DESC_ = 'Hidden Sheet — Auto Protected';

/**
 * Hides a sheet using the Sheets Advanced Service and applies protection.
 * This is more robust than sheet.hideSheet() alone because:
 *   1. The API-level hide is better enforced on Google Sheets mobile.
 *   2. Sheet protection prevents editing even if a mobile user discovers the tab.
 *
 * Falls back to sheet.hideSheet() if the Advanced Service is unavailable.
 *
 * @param {Sheet} sheet - The sheet to hide
 * @private
 */
function setSheetVeryHidden_(sheet) {
  if (!sheet) return;

  // 1. Hide via Sheets Advanced Service (API v4) for stronger mobile enforcement
  try {
    var ssId = sheet.getParent().getId();
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        updateSheetProperties: {
          properties: {
            sheetId: sheet.getSheetId(),
            hidden: true
          },
          fields: 'hidden'
        }
      }]
    }, ssId);
  } catch (_apiError) {
    // Fallback: use standard hideSheet() if Advanced Service is not available
    try {
      sheet.hideSheet();
    } catch (_hideError) {
      Logger.log('Could not hide sheet "' + sheet.getName() + '": ' + _hideError.message);
    }
  }

  // 2. Apply protection so the sheet cannot be edited even if visible on mobile
  protectHiddenSheet_(sheet);
}

/**
 * Shows a previously very-hidden sheet (for admin viewing).
 * Removes the auto-protection so the admin can inspect the sheet.
 *
 * @param {Sheet} sheet - The sheet to show
 * @private
 */
function setSheetVisible_(sheet) {
  if (!sheet) return;

  // 1. Show via Sheets Advanced Service
  try {
    var ssId = sheet.getParent().getId();
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        updateSheetProperties: {
          properties: {
            sheetId: sheet.getSheetId(),
            hidden: false
          },
          fields: 'hidden'
        }
      }]
    }, ssId);
  } catch (_apiError) {
    // Fallback
    try {
      sheet.showSheet();
    } catch (_showError) {
      Logger.log('Could not show sheet "' + sheet.getName() + '": ' + _showError.message);
    }
  }

  // 2. Remove the auto-protection so admin can inspect the data
  removeHiddenSheetProtection_(sheet);
}

/**
 * Applies sheet protection to a hidden sheet.
 * Only the spreadsheet owner / script installer can edit.
 * Skips if a protection with the same description already exists.
 *
 * @param {Sheet} sheet - The sheet to protect
 * @private
 */
function protectHiddenSheet_(sheet) {
  if (!sheet) return;

  // Check for existing protection
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === HIDDEN_SHEET_PROTECTION_DESC_) {
      return; // already protected
    }
  }

  try {
    var protection = sheet.protect().setDescription(HIDDEN_SHEET_PROTECTION_DESC_);
    protection.setWarningOnly(false);

    // Remove all editors except the owner (script installer)
    var editors = protection.getEditors();
    if (editors.length > 0) {
      protection.removeEditors(editors);
    }
    // The owner is always retained by Google Sheets automatically
  } catch (_e) {
    Logger.log('Could not protect sheet "' + sheet.getName() + '": ' + _e.message);
  }
}

/**
 * Removes the auto-protection placed by setSheetVeryHidden_().
 *
 * @param {Sheet} sheet - The sheet to unprotect
 * @private
 */
function removeHiddenSheetProtection_(sheet) {
  if (!sheet) return;

  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === HIDDEN_SHEET_PROTECTION_DESC_) {
      protections[i].remove();
    }
  }
}

/**
 * Enforces hidden state on ALL sheets that should be hidden.
 * Checks every sheet whose name starts with '_' (the project convention for hidden sheets).
 * Safe to call from onOpen() and from a time-driven trigger.
 *
 * @returns {Object} Result with counts of sheets checked and enforced
 */
function enforceHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var enforced = 0;
  var checked = 0;

  // Collect all known hidden sheet names from both HIDDEN_SHEETS and SHEETS constants
  var knownHiddenNames = {};
  for (var key in HIDDEN_SHEETS) {
    knownHiddenNames[HIDDEN_SHEETS[key]] = true;
  }
  // Also include the error log sheet
  knownHiddenNames[ERROR_CONFIG.LOG_SHEET_NAME] = true;

  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    var name = sheet.getName();

    // A sheet should be hidden if its name starts with '_' or is in the known list
    if (name.charAt(0) === '_' || knownHiddenNames[name]) {
      checked++;

      if (!sheet.isSheetHidden()) {
        // Sheet should be hidden but isn't — enforce
        setSheetVeryHidden_(sheet);
        enforced++;
      } else {
        // Already hidden — ensure protection is in place
        protectHiddenSheet_(sheet);
      }
    }
  }

  if (enforced > 0) {
    Logger.log('enforceHiddenSheets: re-hid ' + enforced + ' of ' + checked + ' hidden sheets');
  }

  return { checked: checked, enforced: enforced };
}

/**
 * Installs a time-driven trigger that enforces hidden sheets every hour.
 * Prevents mobile users from accessing hidden sheets between spreadsheet opens.
 */
function installHiddenSheetEnforcerTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);

  // Check if trigger already exists
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enforceHiddenSheets') {
      Logger.log('Hidden sheet enforcer trigger already installed');
      return;
    }
  }

  ScriptApp.newTrigger('enforceHiddenSheets')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Installed hourly hidden-sheet enforcer trigger');
}
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
  ROW_ALT_YELLOW: '#FEFCE8',    // Light yellow tint
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
 * Registry mapping visible sheet names to their tab modal show functions.
 * Used by onSelectionChange to auto-open contextual modals on tab switch.
 * Each entry: { sheet: SHEETS constant, fn: function name string, title: display name }.
 * @const {Array<Object>}
 */
var TAB_MODAL_REGISTRY = [
  { sheet: SHEETS.CONFIG,              fn: 'showTabModalConfig',            title: 'Config' },
  { sheet: SHEETS.MEMBER_DIR,          fn: 'showTabModalMemberDirectory',   title: 'Member Directory' },
  { sheet: SHEETS.GRIEVANCE_LOG,       fn: 'showTabModalGrievanceLog',      title: 'Grievance Log' },
  { sheet: SHEETS.CASE_CHECKLIST,      fn: 'showTabModalCaseChecklist',     title: 'Case Checklist' },
  { sheet: SHEETS.FEEDBACK,            fn: 'showTabModalFeedback',          title: 'Feedback & Development' },
  { sheet: SHEETS.VOLUNTEER_HOURS,     fn: 'showTabModalVolunteerHours',    title: 'Volunteer Hours' },
  { sheet: SHEETS.MEETING_ATTENDANCE,  fn: 'showTabModalMeetingAttendance', title: 'Meeting Attendance' },
  { sheet: SHEETS.MEETING_CHECKIN_LOG, fn: 'showTabModalMeetingCheckIn',    title: 'Meeting Check-In Log' },
  { sheet: SHEETS.RESOURCES,           fn: 'showTabModalResources',         title: 'Resources' }
];

// ============================================================================
// MEMBER DIRECTORY COLUMNS — Auto-derived from header map
// To add/remove/reorder columns, edit this array. Everything else follows.
// ============================================================================

var MEMBER_HEADER_MAP_ = [
  { key: 'MEMBER_ID',          header: 'Member ID' },
  { key: 'FIRST_NAME',         header: 'First Name' },
  { key: 'LAST_NAME',          header: 'Last Name' },
  { key: 'JOB_TITLE',          header: 'Job Title' },
  { key: 'WORK_LOCATION',      header: 'Work Location' },
  { key: 'UNIT',               header: 'Unit' },
  { key: 'CUBICLE',            header: 'Cubicle' },
  { key: 'OFFICE_DAYS',        header: 'Office Days' },
  { key: 'EMAIL',              header: 'Email' },
  { key: 'PHONE',              header: 'Phone' },
  { key: 'PREFERRED_COMM',     header: 'Preferred Communication' },
  { key: 'BEST_TIME',          header: 'Best Time to Contact' },
  { key: 'SUPERVISOR',         header: 'Supervisor' },
  { key: 'MANAGER',            header: 'Manager' },
  { key: 'IS_STEWARD',         header: 'Is Steward' },
  { key: 'SHARE_PHONE',        header: 'Share Phone' },       // Opt-in: steward allows members to see their phone number (Yes/No)
  { key: 'COMMITTEES',         header: 'Committees' },
  { key: 'ASSIGNED_STEWARD',   header: 'Assigned Steward' },
  { key: 'LAST_VIRTUAL_MTG',   header: 'Last Virtual Mtg' },
  { key: 'LAST_INPERSON_MTG',  header: 'Last In-Person Mtg' },
  { key: 'OPEN_RATE',          header: 'Open Rate %' },
  { key: 'VOLUNTEER_HOURS',    header: 'Volunteer Hours' },
  { key: 'INTEREST_LOCAL',     header: 'Interest: Local' },
  { key: 'INTEREST_CHAPTER',   header: 'Interest: Chapter' },
  { key: 'INTEREST_ALLIED',    header: 'Interest: Allied' },
  { key: 'RECENT_CONTACT_DATE', header: 'Recent Contact Date' },
  { key: 'CONTACT_STEWARD',    header: 'Contact Steward' },
  { key: 'CONTACT_NOTES',      header: 'Contact Notes' },
  { key: 'HAS_OPEN_GRIEVANCE', header: 'Has Open Grievance?' },
  { key: 'GRIEVANCE_STATUS',   header: 'Grievance Status' },
  { key: 'NEXT_DEADLINE',      header: 'Days to Deadline' },
  { key: 'START_GRIEVANCE',    header: 'Start Grievance' },
  { key: 'QUICK_ACTIONS',      header: '\u26A1 Actions' },
  { key: 'PIN_HASH',           header: 'PIN Hash' },
  { key: 'EMPLOYEE_ID',        header: 'Employee ID' },
  { key: 'DEPARTMENT',         header: 'Department' },
  { key: 'HIRE_DATE',          header: 'Hire Date' },
  { key: 'STREET_ADDRESS',     header: 'Street Address' },
  { key: 'CITY',               header: 'City' },
  { key: 'STATE',              header: 'State' },
  { key: 'ZIP_CODE',           header: 'Zip Code' },
  { key: 'SHIRT_SIZE',         header: 'Shirt Size' },
  { key: 'DUES_STATUS',             header: 'Dues Status' },
  { key: 'MEMBER_ADMIN_FOLDER_URL',  header: 'Member Admin Folder URL' },   // Drive master folder URL — auto-set on first contact or grievance; steward-visible only
  { key: 'ROLE',                     header: 'Role' }                        // Member/Steward/Both — used by auth to support dual-role users
];

// CONVENTION: Column constants are 1-indexed (Range API). Use COL - 1 for 0-indexed array access.
var MEMBER_COLS = buildColsFromMap_(MEMBER_HEADER_MAP_, {
  LOCATION: 'WORK_LOCATION',
  DAYS_TO_DEADLINE: 'NEXT_DEADLINE'
});

/** PII columns — auto-derived from MEMBER_COLS */
var PII_MEMBER_COLS = [MEMBER_COLS.STREET_ADDRESS, MEMBER_COLS.CITY, MEMBER_COLS.STATE, MEMBER_COLS.ZIP_CODE];

// ============================================================================
// NON-MEMBER CONTACTS COLUMNS — Auto-derived from header map (v4.48.0)
// External contacts: management, legal, HR, union reps, allies.
// ============================================================================

var NMC_HEADER_MAP_ = [
  { key: 'CONTACT_ID',    header: 'Contact ID' },
  { key: 'FIRST_NAME',    header: 'First Name' },
  { key: 'LAST_NAME',     header: 'Last Name' },
  { key: 'JOB_TITLE',     header: 'Job Title' },
  { key: 'WORK_LOCATION', header: 'Work Location' },
  { key: 'UNIT',          header: 'Unit' },
  { key: 'UNION_NAME',    header: 'Union Name' },
  { key: 'SHIRT_SIZE',    header: 'Shirt Size' },
  { key: 'IS_STEWARD',    header: 'Steward' },
  { key: 'EMAIL',         header: 'Email' },
  { key: 'PHONE',         header: 'Phone' },
  { key: 'CATEGORY',      header: 'Category' },
  { key: 'NOTES',         header: 'Notes' }
];

var NMC_COLS = buildColsFromMap_(NMC_HEADER_MAP_);

/**
 * Returns ordered header labels for Non-Member Contacts sheet.
 * @returns {Array<string>}
 */
function getNMCHeaders() {
  return NMC_HEADER_MAP_.map(function(e) { return e.header; });
}

// ============================================================================
// MEETING CHECK-IN LOG COLUMNS — Auto-derived from header map
// ============================================================================

var MEETING_CHECKIN_HEADER_MAP_ = [
  { key: 'MEETING_ID',       header: 'Meeting ID' },
  { key: 'MEETING_NAME',     header: 'Meeting Name' },
  { key: 'MEETING_DATE',     header: 'Meeting Date' },
  { key: 'MEETING_TYPE',     header: 'Meeting Type' },
  { key: 'MEMBER_ID',        header: 'Member ID' },
  { key: 'MEMBER_NAME',      header: 'Member Name' },
  { key: 'CHECKIN_TIME',     header: 'Check-In Time' },
  { key: 'EMAIL',            header: 'Email' },
  { key: 'MEETING_TIME',     header: 'Meeting Time' },
  { key: 'MEETING_DURATION', header: 'Duration' },
  { key: 'EVENT_STATUS',     header: 'Event Status' },
  { key: 'NOTIFY_STEWARDS',  header: 'Notify Stewards' },
  { key: 'CALENDAR_EVENT_ID', header: 'Calendar Event ID' },
  { key: 'NOTES_DOC_URL',    header: 'Notes Doc URL' },
  { key: 'AGENDA_DOC_URL',   header: 'Agenda Doc URL' },
  { key: 'AGENDA_STEWARDS',  header: 'Agenda Stewards' }
];

var MEETING_CHECKIN_COLS = buildColsFromMap_(MEETING_CHECKIN_HEADER_MAP_);

/**
 * Meeting event statuses
 * @const {Object}
 */
var MEETING_STATUS = {
  SCHEDULED: 'Scheduled',
  ACTIVE: 'Active',
  COMPLETED: 'Completed'
};

/**
 * RSVP statuses for meeting invitation tracking (v4.36.0)
 * @const {Object}
 */
var RSVP_STATUS = {
  INVITED:     'Invited',
  ACCEPTED:    'Accepted',
  DECLINED:    'Declined',
  NO_RESPONSE: 'No Response',
  ATTENDED:    'Attended'
};

// ============================================================================
// GRIEVANCE LOG COLUMNS — Auto-derived from header map
// To add/remove/reorder columns, edit this array. Everything else follows.
// ============================================================================

var GRIEVANCE_HEADER_MAP_ = [
  { key: 'GRIEVANCE_ID',       header: 'Grievance ID' },
  { key: 'MEMBER_ID',          header: 'Member ID' },
  { key: 'FIRST_NAME',         header: 'First Name' },
  { key: 'LAST_NAME',          header: 'Last Name' },
  { key: 'STATUS',             header: 'Status' },
  { key: 'CURRENT_STEP',       header: 'Current Step' },
  { key: 'INCIDENT_DATE',      header: 'Incident Date' },
  { key: 'FILING_DEADLINE',    header: 'Filing Deadline' },
  { key: 'DATE_FILED',         header: 'Date Filed' },
  { key: 'STEP1_DUE',          header: 'Step I Due' },
  { key: 'STEP1_RCVD',         header: 'Step I Rcvd' },
  { key: 'STEP2_APPEAL_DUE',   header: 'Step II Appeal Due' },
  { key: 'STEP2_APPEAL_FILED', header: 'Step II Appeal Filed' },
  { key: 'STEP2_DUE',          header: 'Step II Due' },
  { key: 'STEP2_RCVD',         header: 'Step II Rcvd' },
  { key: 'STEP3_APPEAL_DUE',   header: 'Step III Appeal Due' },
  { key: 'STEP3_APPEAL_FILED', header: 'Step III Appeal Filed' },
  { key: 'DATE_CLOSED',        header: 'Date Closed' },
  { key: 'DAYS_OPEN',          header: 'Days Open' },
  { key: 'NEXT_ACTION_DUE',    header: 'Next Action Due' },
  { key: 'DAYS_TO_DEADLINE',   header: 'Days to Deadline' },
  { key: 'ARTICLES',           header: 'Articles Violated' },
  { key: 'ISSUE_CATEGORY',     header: 'Issue Category' },
  { key: 'MEMBER_EMAIL',       header: 'Member Email' },
  { key: 'LOCATION',           header: 'Work Location' },
  { key: 'STEWARD',            header: 'Assigned Steward' },
  { key: 'RESOLUTION',         header: 'Resolution' },
  { key: 'MESSAGE_ALERT',      header: 'Message Alert' },
  { key: 'COORDINATOR_MESSAGE', header: 'Coordinator Message' },
  { key: 'ACKNOWLEDGED_BY',    header: 'Acknowledged By' },
  { key: 'ACKNOWLEDGED_DATE',  header: 'Acknowledged Date' },
  { key: 'DRIVE_FOLDER_ID',    header: 'Drive Folder ID' },
  { key: 'DRIVE_FOLDER_URL',   header: 'Drive Folder URL' },
  { key: 'QUICK_ACTIONS',      header: '\u26A1 Actions' },
  { key: 'ACTION_TYPE',        header: 'Action Type' },
  { key: 'CHECKLIST_PROGRESS', header: 'Checklist Progress' },
  { key: 'REMINDER_1_DATE',    header: 'Reminder 1 Date' },
  { key: 'REMINDER_1_NOTE',    header: 'Reminder 1 Note' },
  { key: 'REMINDER_2_DATE',    header: 'Reminder 2 Date' },
  { key: 'REMINDER_2_NOTE',    header: 'Reminder 2 Note' },
  { key: 'LAST_UPDATED',       header: 'Last Updated' }
];

var GRIEVANCE_COLS = buildColsFromMap_(GRIEVANCE_HEADER_MAP_, {
  GRIEVANCE_STATUS: 'STATUS',
  GRIEVANCE_STEP: 'CURRENT_STEP'
});

// ============================================================================
// CONFIG COLUMN MAPPING — Auto-derived from header map
// Config sheet uses row 2 for headers (row 1 is section headers).
// ============================================================================

// Column types: list = multi-row dropdown options, text = free text, number = numeric,
// url = URL, id = Google API ID (15+ chars), email = email address, days = positive integer,
// boolean = yes/no toggle, label = short UI text (≤50 chars), csv = comma-separated values
//
// Columns are grouped by logical section. syncColumnMaps() resolves columns by header
// name at runtime, so reordering here is safe for existing deployments.
var CONFIG_HEADER_MAP_ = [
  // ── ORGANIZATION (A–J) ── identity, contract info
  { key: 'ORG_NAME',              header: 'Organization Name',                 type: 'text' },
  { key: 'ORG_ABBREV',            header: 'Organization Abbreviation',         type: 'label' },
  { key: 'LOCAL_NUMBER',          header: 'Local Number',                      type: 'text' },
  { key: 'UNION_PARENT',          header: 'Union Parent',                      type: 'text' },
  { key: 'STATE_REGION',          header: 'State/Region',                      type: 'text' },
  { key: 'ORG_WEBSITE',           header: 'Organization Website',              type: 'url' },
  { key: 'CONTRACT_NAME',         header: 'Contract Name',                     type: 'text' },
  { key: 'CONTRACT_GRIEVANCE',    header: 'Contract Article (Grievance)',      type: 'text' },
  { key: 'CONTRACT_DISCIPLINE',   header: 'Contract Article (Discipline)',     type: 'text' },
  { key: 'CONTRACT_WORKLOAD',     header: 'Contract Article (Workload)',       type: 'text' },

  // ── CONTACT INFO (K–P) ── addresses, phones, emails
  { key: 'MAIN_ADDRESS',          header: 'Main Office Address',               type: 'text' },
  { key: 'MAIN_PHONE',            header: 'Main Phone',                        type: 'text' },
  { key: 'MAIN_FAX',              header: 'Main Fax',                          type: 'text' },
  { key: 'MAIN_CONTACT_NAME',     header: 'Main Contact Name',                type: 'text' },
  { key: 'MAIN_CONTACT_EMAIL',    header: 'Main Contact Email',               type: 'email' },
  { key: 'CHIEF_STEWARD_EMAIL',   header: 'Chief Steward Email',              type: 'email' },

  // ── EMPLOYMENT (Q–W) ── workplace structure, job info
  { key: 'JOB_TITLES',            header: 'Job Titles',                        type: 'list' },
  { key: 'OFFICE_LOCATIONS',      header: 'Office Locations',                  type: 'list' },
  { key: 'OFFICE_ADDRESSES',      header: 'Office Addresses',                  type: 'text' },
  { key: 'OFFICE_DAYS',           header: 'Office Days',                       type: 'list' },
  { key: 'UNITS',                 header: 'Units',                             type: 'list' },
  { key: 'UNIT_CODES',            header: 'Unit Codes',                        type: 'list' },
  { key: 'DUES_STATUSES',         header: 'Dues Statuses',                     type: 'list' },

  // ── PEOPLE (X–AC) ���─ supervisors, stewards, contact prefs
  { key: 'SUPERVISORS',           header: 'Supervisors',                       type: 'list' },
  { key: 'MANAGERS',              header: 'Managers',                          type: 'list' },
  { key: 'STEWARDS',              header: 'Stewards',                          type: 'list' },
  { key: 'STEWARD_COMMITTEES',    header: 'Steward Committees',                type: 'list' },
  { key: 'COMM_METHODS',          header: 'Communication Methods',             type: 'list' },
  { key: 'BEST_TIMES',            header: 'Best Times to Contact',             type: 'list' },

  // ── GRIEVANCE SETTINGS (AD–AJ) ── statuses, steps, categories, escalation
  { key: 'GRIEVANCE_STATUS',      header: 'Grievance Status',                  type: 'list' },
  { key: 'GRIEVANCE_STEP',        header: 'Grievance Step',                    type: 'list' },
  { key: 'ISSUE_CATEGORY',        header: 'Issue Category',                    type: 'list' },
  { key: 'ARTICLES',              header: 'Articles Violated',                 type: 'list' },
  { key: 'GRIEVANCE_COORDINATORS', header: 'Grievance Coordinators',           type: 'list' },
  { key: 'ESCALATION_STATUSES',   header: 'Escalation Statuses',              type: 'list' },
  { key: 'ESCALATION_STEPS',      header: 'Escalation Steps',                 type: 'list' },

  // ── DEADLINES (AK–AR) ── grievance timeline, alerts
  { key: 'FILING_DEADLINE_DAYS',  header: 'Filing Deadline Days',              type: 'days' },
  { key: 'STEP1_RESPONSE_DAYS',   header: 'Step I Response Days',              type: 'days' },
  { key: 'STEP2_APPEAL_DAYS',     header: 'Step II Appeal Days',               type: 'days' },
  { key: 'STEP2_RESPONSE_DAYS',   header: 'Step II Response Days',             type: 'days' },
  { key: 'STEP3_APPEAL_DAYS',     header: 'Step III Appeal Days',              type: 'days' },
  { key: 'STEP3_RESPONSE_DAYS',   header: 'Step III Response Days',            type: 'days' },
  { key: 'ARBITRATION_DEMAND_DAYS', header: 'Arbitration Demand Days',         type: 'days' },
  { key: 'ALERT_DAYS',            header: 'Alert Days Before Deadline',        type: 'csv' },

  // ── NOTIFICATIONS (AS–AU) ── admin & system email lists
  { key: 'ADMIN_EMAILS',          header: 'Admin Emails',                      type: 'list' },
  { key: 'NOTIFICATION_RECIPIENTS', header: 'Notification Recipients',         type: 'list' },
  { key: 'TEST_NOTIFY_EMAIL',     header: 'Test Runner Notify Email',          type: 'email' },

  // ── LINKS (AV–BB) ── external URLs, custom sidebar links
  { key: 'GRIEVANCE_FORM_URL',    header: 'Grievance Form URL',               type: 'url' },
  { key: 'CONTACT_FORM_URL',      header: 'Contact Form URL',                 type: 'url' },
  { key: 'MOBILE_DASHBOARD_URL',  header: '\uD83D\uDCF1 Mobile Dashboard URL', type: 'url' },
  { key: 'CUSTOM_LINK_1_NAME',    header: 'Custom Link 1 Name',               type: 'label' },
  { key: 'CUSTOM_LINK_1_URL',     header: 'Custom Link 1 URL',                type: 'url' },
  { key: 'CUSTOM_LINK_2_NAME',    header: 'Custom Link 2 Name',               type: 'label' },
  { key: 'CUSTOM_LINK_2_URL',     header: 'Custom Link 2 URL',                type: 'url' },

  // ── DRIVE & CALENDAR (BC–BM) ── Google Drive folder IDs, Calendar ID
  { key: 'DRIVE_FOLDER_ID',       header: 'Google Drive Folder ID',            type: 'id' },
  { key: 'CALENDAR_ID',           header: 'Google Calendar ID',                type: 'id' },
  { key: 'ARCHIVE_FOLDER_ID',     header: 'Archive Folder ID',                type: 'id' },
  { key: 'TEMPLATE_ID',           header: 'Template ID',                       type: 'id' },
  { key: 'PDF_FOLDER_ID',         header: 'PDF Folder ID',                    type: 'id' },
  { key: 'DASHBOARD_ROOT_FOLDER_ID',   header: 'Dashboard Root Folder ID',    type: 'id' },
  { key: 'GRIEVANCES_FOLDER_ID',       header: 'Grievances Folder ID',        type: 'id' },
  { key: 'RESOURCES_FOLDER_ID',        header: 'Resources Folder ID',         type: 'id' },
  { key: 'MINUTES_FOLDER_ID',          header: 'Minutes Folder ID',           type: 'id' },
  { key: 'EVENT_CHECKIN_FOLDER_ID',    header: 'Event Check-In Folder ID',    type: 'id' },
  { key: 'PAST_SURVEYS_FOLDER_ID',     header: 'Past Surveys Folder ID',      type: 'id' },

  // ── SURVEY (BN–BP) ── priority options, response logs
  { key: 'SURVEY_PRIORITY_OPTIONS',  header: 'Survey Priority Options',       type: 'list' },
  { key: 'SURVEY_LOG_IDS',        header: 'Survey Log (Member IDs)',           type: 'list' },
  { key: 'SURVEY_LOG_DATES',      header: 'Survey Log (Dates)',                type: 'list' },

  // ── BRANDING & UX (BQ–BV) ── visual theme, role labels, session config
  { key: 'ACCENT_HUE',            header: 'Accent Hue',                       type: 'number' },
  { key: 'LOGO_INITIALS',         header: 'Logo Initials',                    type: 'label' },
  { key: 'STEWARD_LABEL',         header: 'Steward Label',                    type: 'label' },
  { key: 'MEMBER_LABEL',          header: 'Member Label',                     type: 'label' },
  { key: 'MAGIC_LINK_EXPIRY_DAYS', header: 'Magic Link Expiry Days',          type: 'days' },
  { key: 'COOKIE_DURATION_DAYS',  header: 'Cookie Duration Days',             type: 'days' },

  // ── FEATURE TOGGLES (BW–CA) ── on/off switches, cache TTL
  { key: 'SHOW_GRIEVANCES',       header: 'Show Grievances',                   type: 'boolean' },
  { key: 'BROADCAST_SCOPE_ALL',   header: 'Broadcast: Allow All Members Scope', type: 'boolean' },
  { key: 'ENABLE_CORRELATION',    header: 'Enable Correlation Engine',         type: 'boolean' },
  { key: 'ENABLE_TAB_MODALS',     header: 'Enable Tab Modals',                type: 'boolean' },
  { key: 'INSIGHTS_CACHE_TTL_MIN', header: 'Insights Cache TTL (Minutes)',    type: 'number' },

  // ── RETENTION (CB–CC) ── auto-archival thresholds
  { key: 'GRIEVANCE_ARCHIVE_DAYS', header: 'Grievance Archive Days',           type: 'days' },
  { key: 'AUDIT_ARCHIVE_DAYS',    header: 'Audit Log Archive Days',            type: 'days' }
];

var CONFIG_COLS = buildColsFromMap_(CONFIG_HEADER_MAP_);

/**
 * Data type validators for Config columns.
 * Each validator returns true if the value is plausible for that column type.
 * Used by validateConfigValue_() to detect misaligned data at runtime.
 */
var CONFIG_DATA_VALIDATORS = {
  list:    function() { return true; },
  text:    function(v) { return typeof v === 'string'; },
  number:  function(v) { return /^\d+(\.\d+)?$/.test(String(v).trim()); },
  url:     function(v) { return /^https?:\/\//.test(String(v).trim()); },
  id:      function(v) { return /^[a-zA-Z0-9_-]{15,}$/.test(String(v).trim()); },
  email:   function(v) { return /@/.test(String(v)); },
  days:    function(v) { return /^\d+$/.test(String(v).trim()) && parseInt(v, 10) > 0; },
  boolean: function(v) { return /^(yes|no|true|false)$/i.test(String(v).trim()); },
  label:   function(v) { return typeof v === 'string' && v.length > 0 && v.length <= 50; },
  csv:     function(v) { return typeof v === 'string' && /\d/.test(v); }
};

/**
 * Validates a value against its Config column's declared type.
 * Empty/blank values always pass (columns can be empty).
 * @param {string} key - CONFIG_HEADER_MAP_ key (e.g. 'STEWARD_LABEL')
 * @param {*} value - The value to validate
 * @returns {{ valid: boolean, type: string, reason: string }}
 */
function validateConfigValue_(key, value) {
  for (var i = 0; i < CONFIG_HEADER_MAP_.length; i++) {
    if (CONFIG_HEADER_MAP_[i].key === key) {
      var entry = CONFIG_HEADER_MAP_[i];
      if (!entry.type) return { valid: true, type: '', reason: 'no type defined' };
      if (!value || String(value).trim() === '') return { valid: true, type: entry.type, reason: 'empty' };
      var validator = CONFIG_DATA_VALIDATORS[entry.type];
      if (!validator) return { valid: true, type: entry.type, reason: 'unknown type' };
      var ok = validator(value);
      return { valid: ok, type: entry.type, reason: ok ? 'ok' : 'expected ' + entry.type + ', got: ' + String(value).substring(0, 50) };
    }
  }
  return { valid: false, type: '', reason: 'unknown key: ' + key };
}

/**
 * Returns true if grievance features are enabled (Config: SHOW_GRIEVANCES ≠ 'no').
 * Defaults to true (enabled) if the config value is missing or blank.
 * @returns {boolean}
 */
function _isGrievancesEnabled() {
  try {
    var val = getConfigValue_(CONFIG_COLS.SHOW_GRIEVANCES);
    if (!val || String(val).trim() === '') return true;
    return String(val).toLowerCase().trim() !== 'no';
  } catch (_e) { return true; }
}

// ============================================================================
// SURVEY PERIODS COLUMNS — Hidden sheet: _Survey_Periods (v4.21.0)
// ============================================================================
/**
 * Tracks quarterly survey periods.
 * One row per period. Status: 'Active' | 'Closed'.
 * archiveSurveyPeriod_() writes Drive archive URL here when closing.
 */
var SURVEY_PERIODS_HEADER_MAP_ = [
  { key: 'PERIOD_ID',      header: 'Period ID' },
  { key: 'PERIOD_NAME',    header: 'Period Name' },
  { key: 'START_DATE',     header: 'Start Date' },
  { key: 'END_DATE',       header: 'End Date' },
  { key: 'STATUS',         header: 'Status' },
  { key: 'ARCHIVE_URL',    header: 'Archive Folder URL' },
  { key: 'CREATED_BY',     header: 'Created By' },
  { key: 'RESPONSE_COUNT', header: 'Response Count' }
];
var SURVEY_PERIODS_COLS = buildColsFromMap_(SURVEY_PERIODS_HEADER_MAP_);

// ============================================================================
// SURVEY QUESTIONS SHEET COLUMNS — v4.23.0 Dynamic Schema
// Sheet: '📋 Survey Questions'
// Owner edits: Question Text (col 5), Active (col 8), Options (col 9),
//              Slider Min/Max (cols 14-15), Notes (col 16)
// Structural (code-managed): cols 1-4, 6-7, 10-13
// ============================================================================

var SURVEY_QUESTIONS_HEADER_MAP_ = [
  { key: 'QUESTION_ID',    header: 'Question ID' },
  { key: 'SECTION_NUM',    header: 'Section' },
  { key: 'SECTION_KEY',    header: 'Section Key' },
  { key: 'SECTION_TITLE',  header: 'Section Title' },
  { key: 'QUESTION_TEXT',  header: 'Question Text' },
  { key: 'TYPE',           header: 'Type' },
  { key: 'REQUIRED',       header: 'Required' },
  { key: 'ACTIVE',         header: 'Active' },
  { key: 'OPTIONS',        header: 'Options' },
  { key: 'BRANCH_PARENT',  header: 'Branch Parent' },
  { key: 'BRANCH_VALUE',   header: 'Branch Value' },
  { key: 'BRANCH_TARGET',  header: 'Branch Target' },
  { key: 'MAX_SELECTIONS', header: 'Max Selections' },
  { key: 'SLIDER_MIN',     header: 'Slider Min Label' },
  { key: 'SLIDER_MAX',     header: 'Slider Max Label' },
  { key: 'NOTES',          header: 'Notes' }
];
var SURVEY_QUESTIONS_COLS = buildColsFromMap_(SURVEY_QUESTIONS_HEADER_MAP_);

// Satisfaction sheet fixed prefix columns (v4.23.0).
// Question data starts at DATA_START. Question IDs (q1, q2…) are used as headers.
var SATISFACTION_PREFIX = {
  TIMESTAMP:      1,
  PERIOD_ID:      2,
  SURVEY_VERSION: 3,
  DATA_START:     4
};
// ============================================================================

var STEWARD_PERF_HEADER_MAP_ = [
  { key: 'STEWARD',           header: 'Steward' },
  { key: 'TOTAL_CASES',       header: 'Total Cases' },
  { key: 'ACTIVE',            header: 'Active' },
  { key: 'CLOSED',            header: 'Closed' },
  { key: 'WON',               header: 'Won' },
  { key: 'WIN_RATE',          header: 'Win Rate %' },
  { key: 'AVG_DAYS',          header: 'Avg Days' },
  { key: 'OVERDUE',           header: 'Overdue' },
  { key: 'DUE_THIS_WEEK',     header: 'Due This Week' },
  { key: 'PERFORMANCE_SCORE', header: 'Performance Score' }
];

var STEWARD_PERF_COLS = buildColsFromMap_(STEWARD_PERF_HEADER_MAP_);

// ============================================================================
// AUDIT LOG COLUMNS — Auto-derived from header map
// ============================================================================

var AUDIT_LOG_HEADER_MAP_ = [
  { key: 'TIMESTAMP',   header: 'Timestamp' },
  { key: 'USER_EMAIL',  header: 'User Email' },
  { key: 'SHEET',       header: 'Sheet' },
  { key: 'ROW',         header: 'Row' },
  { key: 'COLUMN',      header: 'Column' },
  { key: 'FIELD_NAME',  header: 'Field Name' },
  { key: 'OLD_VALUE',   header: 'Old Value' },
  { key: 'NEW_VALUE',   header: 'New Value' },
  { key: 'RECORD_ID',   header: 'Record ID' },
  { key: 'ACTION_TYPE', header: 'Action Type' }
];

var AUDIT_LOG_COLS = buildColsFromMap_(AUDIT_LOG_HEADER_MAP_);

// ============================================================================
// EVENT AUDIT LOG COLUMNS — used by logAuditEvent() in 06_Maintenance.gs
// Same sheet, different schema from the edit-level AUDIT_LOG_COLS.
// ============================================================================

var EVENT_AUDIT_HEADER_MAP_ = [
  { key: 'TIMESTAMP',      header: 'Timestamp' },
  { key: 'EVENT_TYPE',     header: 'Event Type' },
  { key: 'USER',           header: 'User' },
  { key: 'DETAILS',        header: 'Details' },
  { key: 'SESSION_ID',     header: 'Session ID' },
  { key: 'INTEGRITY_HASH', header: 'Integrity Hash' }
];

var EVENT_AUDIT_COLS = buildColsFromMap_(EVENT_AUDIT_HEADER_MAP_);

// ============================================================================
// SATISFACTION SURVEY COLUMNS
// @deprecated v4.23.0 — column positions are now dynamic.
// Use getSatisfactionColMap_() for runtime column lookups.
// Kept as reference/fallback for legacy direct-sheet reads in 04c/04d.
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

  // ── VERIFICATION COLUMNS — MOVED TO _Survey_Vault (v4.8) ──
  // These constants are DEPRECATED. All PII is now stored in the
  // _Survey_Vault hidden sheet (SURVEY_VAULT_COLS) to ensure survey
  // anonymity. The Satisfaction sheet no longer contains any data that
  // can link a response to a specific member.
  // @deprecated v4.8 — use SURVEY_VAULT_COLS instead
  EMAIL: -1,
  VERIFIED: -1,
  MATCHED_MEMBER_ID: -1,
  QUARTER: -1,
  IS_LATEST: -1,
  SUPERSEDED_BY: -1,
  REVIEWER_NOTES: -1
};

// ============================================================================
// SURVEY VAULT COLUMNS (8 columns: A-H) — Zero-knowledge hash store
// ============================================================================

/**
 * Survey Vault column positions (1-indexed)
 * Hidden + protected sheet: _Survey_Vault
 *
 * PURPOSE:
 *   Provides verified/latest/quarter metadata for survey responses without
 *   storing any reversible PII. Email and member ID are SHA-256 hashed with
 *   a per-installation salt — even with full vault access, it is
 *   cryptographically impossible to determine who submitted a response.
 *
 * SECURITY MODEL:
 *   - Email and Member ID are stored as salted SHA-256 hashes only
 *   - Hashes are non-reversible — no one can recover the original values
 *   - Raw email exists in memory only (during form submit) and is never persisted
 *   - Sheet is hidden (prefixed with _) and sheet-protected
 *   - Dashboard code reads only {verified, isLatest, quarter} via getVaultDataMap_()
 *
 * DATA FLOW:
 *   1. onSatisfactionFormSubmit() writes survey answers to Satisfaction sheet
 *      (anonymous — no email, no member ID, no hashes)
 *   2. Same function hashes email + member ID in-memory, writes hashes to vault
 *   3. Raw email is used only to send thank-you email, then discarded
 *   4. Superseding logic compares hashes, never plaintext
 *
 * @const {Object}
 */
var SURVEY_VAULT_HEADER_MAP_ = [
  { key: 'RESPONSE_ROW',     header: 'Response Row' },
  { key: 'EMAIL',            header: 'Email Hash' },
  { key: 'VERIFIED',         header: 'Verified' },
  { key: 'MATCHED_MEMBER_ID', header: 'Member ID Hash' },
  { key: 'QUARTER',          header: 'Quarter' },
  { key: 'IS_LATEST',        header: 'Is Latest' },
  { key: 'SUPERSEDED_BY',    header: 'Superseded By' },
  { key: 'REVIEWER_NOTES',   header: 'Reviewer Notes' }
];

var SURVEY_VAULT_COLS = buildColsFromMap_(SURVEY_VAULT_HEADER_MAP_);

/**
 * Survey section definitions for grouping and analysis
 * @const {Object}
 */
var SATISFACTION_SECTIONS = {
  WORK_CONTEXT: { name: 'Work Context', questions: [2,3,4,5,6], scale: false },
  OVERALL_SAT: { name: 'Overall Satisfaction', questions: [7,8,9,10], scale: true,
    questionTexts: ['Satisfied with representation', 'Trust union advocacy', 'Feel protected by union', 'Would recommend joining'] },
  STEWARD_3A: { name: 'Steward Ratings', questions: [11,12,13,14,15,16,17], scale: true,
    questionTexts: ['Timely response', 'Treated with respect', 'Explained options clearly', 'Followed through', 'Advocated effectively', 'Safe raising concerns', 'Maintained confidentiality'] },
  STEWARD_3B: { name: 'Steward Access', questions: [19,20,21], scale: true,
    questionTexts: ['Know how to contact', 'Confident steward would help', 'Easy to find steward'] },
  CHAPTER: { name: 'Chapter Effectiveness', questions: [22,23,24,25,26], scale: true,
    questionTexts: ['Understands workplace issues', 'Effective communication', 'Organizes well', 'Easy to reach chapter', 'Fair representation'] },
  LEADERSHIP: { name: 'Local Leadership', questions: [27,28,29,30,31,32], scale: true,
    questionTexts: ['Decisions are clear', 'Understand grievance process', 'Transparent finances', 'Accountable leadership', 'Fair processes', 'Welcomes opinions'] },
  CONTRACT: { name: 'Contract Enforcement', questions: [33,34,35,36], scale: true,
    questionTexts: ['Enforces contract', 'Realistic timelines', 'Clear updates', 'Frontline priority'] },
  REPRESENTATION: { name: 'Representation Process', questions: [38,39,40,41], scale: true,
    questionTexts: ['Understood the steps', 'Felt supported', 'Updated often enough', 'Outcome was justified'] },
  COMMUNICATION: { name: 'Communication Quality', questions: [42,43,44,45,46], scale: true,
    questionTexts: ['Clear and actionable', 'Enough information', 'Easy to find info', 'Reaches all shifts', 'Meetings worth attending'] },
  MEMBER_VOICE: { name: 'Member Voice & Culture', questions: [47,48,49,50,51], scale: true,
    questionTexts: ['Voice matters', 'Seeks member input', 'Treated with dignity', 'Newer members supported', 'Conflicts handled respectfully'] },
  VALUE_ACTION: { name: 'Value & Collective Action', questions: [52,53,54,55,56], scale: true,
    questionTexts: ['Good value for dues', 'Priorities match needs', 'Prepared to mobilize', 'Know how to get involved', 'Win together'] },
  SCHEDULING: { name: 'Scheduling/Office Days', questions: [57,58,59,60,61,62,63], scale: true,
    questionTexts: ['Understand changes', 'Adequately informed', 'Clear criteria', 'Reasonable expectations', 'Effective outcomes', 'Supports wellbeing', 'Concerns taken seriously'] },
  PRIORITIES: { name: 'Priorities & Close', questions: [65,66,67,68], scale: false }
};

// ============================================================================
// SURVEY TRACKING COLUMNS (10 columns: A-J)
// ============================================================================

/**
 * Survey Completion Tracking column positions (1-indexed)
 * Hidden sheet: _Survey_Tracking
 *
 * PURPOSE:
 *   Tracks per-member survey completion status across multiple survey rounds.
 *   Lets stewards see who has/hasn't completed the satisfaction survey and
 *   send targeted reminders to non-respondents.
 *
 * HOW COMPLETION IS DETECTED:
 *   1. A Google Forms trigger calls onSatisfactionFormSubmit(e)
 *      when a member submits the satisfaction survey.
 *      (Trigger installed via setupSatisfactionFormTrigger() in 08c_FormsAndNotifications.gs)
 *   2. The respondent's email is extracted from the form response
 *      (tries field names "Email Address", "Email", or Google's respondent email).
 *   3. validateMemberEmail(email) in 08c_FormsAndNotifications.gs scans the
 *      Member Directory (MEMBER_COLS.EMAIL, column I) for a case-insensitive match.
 *   4. If a match is found, updateSurveyTrackingOnSubmit_(memberId) looks up the
 *      member in this _Survey_Tracking sheet and sets:
 *        - CURRENT_STATUS  = "Completed"
 *        - COMPLETED_DATE  = now
 *        - TOTAL_COMPLETED = previous + 1
 *   5. If no email match is found, the satisfaction response is still recorded
 *      but tracking status remains "Not Completed" for that member.
 *
 * RELATED FUNCTIONS (all in 08c_FormsAndNotifications.gs):
 *   - populateSurveyTrackingFromMembers() : Syncs member list from Member Directory
 *   - updateSurveyTrackingOnSubmit_(id)   : Auto-called on form submit (the detection hook)
 *   - startNewSurveyRound()               : Resets statuses, increments missed counts
 *   - sendSurveyCompletionReminders()     : Emails non-respondents (7-day cooldown)
 *   - getSurveyCompletionStats()          : Returns { total, completed, notCompleted, rate }
 *   - showSurveyTrackingDialog()          : Management UI modal
 *
 * SHEET SETUP:
 *   - setupSurveyTrackingSheet() in 08d_AuditAndFormulas.gs
 *   - Registered in setupHiddenSheets() in 08a_SheetSetup.gs
 *   - Seed data: seedSurveyTrackingData() in 07_DevTools.gs
 *
 * @const {Object}
 */
var SURVEY_TRACKING_HEADER_MAP_ = [
  { key: 'MEMBER_ID',        header: 'Member ID' },
  { key: 'MEMBER_NAME',      header: 'Member Name' },
  { key: 'EMAIL',            header: 'Email' },
  { key: 'WORK_LOCATION',    header: 'Work Location' },
  { key: 'ASSIGNED_STEWARD', header: 'Assigned Steward' },
  { key: 'CURRENT_STATUS',   header: 'Current Status' },
  { key: 'COMPLETED_DATE',   header: 'Completed Date' },
  { key: 'TOTAL_MISSED',     header: 'Total Missed' },
  { key: 'TOTAL_COMPLETED',  header: 'Total Completed' },
  { key: 'LAST_REMINDER_SENT', header: 'Last Reminder Sent' }
];

var SURVEY_TRACKING_COLS = buildColsFromMap_(SURVEY_TRACKING_HEADER_MAP_);

// ============================================================================
// FEEDBACK & DEVELOPMENT COLUMNS (11 columns: A-K)
// ============================================================================

/**
 * Feedback & Development column positions (1-indexed)
 * @const {Object}
 */
var FEEDBACK_HEADER_MAP_ = [
  { key: 'TIMESTAMP',    header: 'Timestamp' },
  { key: 'SUBMITTED_BY', header: 'Submitted By' },
  { key: 'CATEGORY',     header: 'Category' },
  { key: 'PRIORITY',     header: 'Priority' },
  { key: 'TITLE',        header: 'Title' },
  { key: 'DESCRIPTION',  header: 'Description' },
  { key: 'STATUS',       header: 'Status' },
  { key: 'ASSIGNED_TO',  header: 'Assigned To' },
  { key: 'RESOLUTION',   header: 'Resolution' },
  { key: 'NOTES',        header: 'Notes' }
]; // v4.24.1: Removed TYPE — was never populated by form (Category covers same ground)

var FEEDBACK_COLS = buildColsFromMap_(FEEDBACK_HEADER_MAP_);

// Resources sheet — educational content management (v4.11.0)
var RESOURCES_HEADER_MAP_ = [
  { key: 'RESOURCE_ID',  header: 'Resource ID' },
  { key: 'TITLE',        header: 'Title' },
  { key: 'CATEGORY',     header: 'Category' },
  { key: 'SUMMARY',      header: 'Summary' },
  { key: 'CONTENT',      header: 'Content' },
  { key: 'URL',          header: 'URL' },
  { key: 'ICON',         header: 'Icon' },
  { key: 'SORT_ORDER',   header: 'Sort Order' },
  { key: 'VISIBLE',      header: 'Visible' },
  { key: 'AUDIENCE',     header: 'Audience' },
  { key: 'DATE_ADDED',   header: 'Date Added' },
  { key: 'ADDED_BY',     header: 'Added By' }
];

var RESOURCES_COLS = buildColsFromMap_(RESOURCES_HEADER_MAP_);

// Resource Config sheet — categories & settings for the Resources tab (v4.22.x)
var RESOURCE_CONFIG_HEADER_MAP_ = [
  { key: 'SETTING',      header: 'Setting' },    // e.g. 'Category'
  { key: 'VALUE',        header: 'Value' },       // the category name
  { key: 'SORT_ORDER',   header: 'Sort Order' },  // display order (numeric)
  { key: 'ACTIVE',       header: 'Active' },      // Yes / No
  { key: 'NOTES',        header: 'Notes' }        // optional steward notes
];

var RESOURCE_CONFIG_COLS = buildColsFromMap_(RESOURCE_CONFIG_HEADER_MAP_);

// Knowledge Engine sheet — centralized educational content management (v4.41.0)
// Manages quotes, tips, concepts, mini-lessons, and manifesto phrases in one place.
// Replaces hardcoded content arrays in negotiation_knowledge.html, auth_manifesto.html,
// and member_hub_view.html. Stewards can manage all educational content from the webapp.
var KNOWLEDGE_HEADER_MAP_ = [
  { key: 'CONTENT_ID',    header: 'Content ID' },      // KE-001, auto-generated
  { key: 'TYPE',           header: 'Type' },             // Quote, Tip, Concept, Mini-Lesson, Manifesto, Negotiation Set
  { key: 'CATEGORY',       header: 'Category' },         // Negotiation, Rights, Grievance, Safety, Labor, General
  { key: 'TITLE',          header: 'Title' },             // Short title (e.g. "What Is Negotiation?")
  { key: 'CONTENT',        header: 'Content' },           // Main text / quote text / body
  { key: 'ATTRIBUTION',    header: 'Attribution' },       // Quote source (e.g. "John F. Kennedy")
  { key: 'BULLETS',        header: 'Bullets' },           // Pipe-delimited supporting points
  { key: 'AUDIENCE',       header: 'Audience' },          // All, Members, Stewards
  { key: 'PLACEMENT',      header: 'Placement' },        // Auth Screen, Home Widget, Sidebar, Hub Section, Inline Quote
  { key: 'ACTIVE',         header: 'Active' },            // Yes / No
  { key: 'PRIORITY',       header: 'Priority' },          // Sort/weight for rotation (numeric)
  { key: 'START_DATE',     header: 'Start Date' },        // Optional scheduling
  { key: 'END_DATE',       header: 'End Date' },          // Optional scheduling
  { key: 'DATE_ADDED',     header: 'Date Added' },        // Auto-populated
  { key: 'ADDED_BY',       header: 'Added By' }           // Auto-populated
];

var KNOWLEDGE_COLS = buildColsFromMap_(KNOWLEDGE_HEADER_MAP_);
// v4.22.0: Added DISMISS_MODE column — 'Dismissible' (user can permanently dismiss)
//          or 'Timed' (auto-expires on Expires_Date; dismiss button hidden from members).
var NOTIFICATIONS_HEADER_MAP_ = [
  { key: 'NOTIFICATION_ID', header: 'Notification ID' },
  { key: 'RECIPIENT',       header: 'Recipient' },
  { key: 'TYPE',             header: 'Type' },
  { key: 'TITLE',            header: 'Title' },
  { key: 'MESSAGE',          header: 'Message' },
  { key: 'PRIORITY',         header: 'Priority' },
  { key: 'SENT_BY',          header: 'Sent_By' },
  { key: 'SENT_BY_NAME',     header: 'Sent_By_Name' },
  { key: 'CREATED_DATE',     header: 'Created_Date' },
  { key: 'EXPIRES_DATE',     header: 'Expires_Date' },
  { key: 'DISMISSED_BY',     header: 'Dismissed_By' },
  { key: 'STATUS',           header: 'Status' },
  { key: 'DISMISS_MODE',     header: 'Dismiss_Mode' }  // v4.22.0: 'Dismissible' | 'Timed'
];

var NOTIFICATIONS_COLS = buildColsFromMap_(NOTIFICATIONS_HEADER_MAP_);

// ============================================================================
// COLUMN AUTO-DISCOVERY SYSTEM
// ============================================================================
// Header maps are the single source of truth for each sheet's column layout.
// Column constants (MEMBER_COLS, etc.) and header arrays (getMemberHeaders())
// are both auto-derived from these maps.
//
// To add/remove/reorder columns: edit the header map array.
// Column numbers, header arrays, and legacy compat objects all update automatically.
//
// For runtime adaptation (someone manually reorders columns in the spreadsheet),
// call syncColumnMaps() — it reads actual headers and auto-updates all constants.
// ============================================================================

/**
 * Build column constant object from a header map array.
 * Position is determined by array order (1-indexed) unless entry has explicit 'pos'.
 * @param {Array<{key: string, header: string, pos?: number}>} headerMap
 * @param {Object} [aliases] - { aliasKey: 'existingKey' } for backward compat
 * @returns {Object} Column constants object (1-indexed)
 */
function buildColsFromMap_(headerMap, aliases) {
  var cols = {};
  for (var i = 0; i < headerMap.length; i++) {
    cols[headerMap[i].key] = headerMap[i].pos !== undefined ? headerMap[i].pos : (i + 1);
  }
  if (aliases) {
    for (var alias in aliases) {
      if (aliases.hasOwnProperty(alias)) {
        cols[alias] = cols[aliases[alias]];
      }
    }
  }
  return cols;
}

/**
 * Extract ordered header text array from a header map.
 * Used for sheet creation (writing row 1 headers).
 * @param {Array<{key: string, header: string}>} headerMap
 * @returns {Array<string>} Ordered header labels
 */
function getHeadersFromMap_(headerMap) {
  var headers = [];
  for (var i = 0; i < headerMap.length; i++) {
    headers.push(headerMap[i].header);
  }
  return headers;
}

/**
 * Resolve column positions by reading actual sheet headers at runtime.
 * @param {string} sheetName - Sheet name to read
 * @param {Array<{key: string, header: string}>} headerMap - Expected headers
 * @param {Object} [options] - { headerRow: number } (default: 1)
 * @returns {Object|null} Resolved column map, or null if sheet doesn't exist
 */
function resolveColumnsFromSheet_(sheetName, headerMap, options) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastColumn() < 1) return null;

    var headerRow = (options && options.headerRow) || 1;
    var actualHeaders = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    var headerToKey = {};
    for (var i = 0; i < headerMap.length; i++) {
      headerToKey[headerMap[i].header] = headerMap[i].key;
    }

    var cols = {};
    for (var j = 0; j < actualHeaders.length; j++) {
      var text = String(actualHeaders[j]).trim();
      if (headerToKey[text]) {
        cols[headerToKey[text]] = j + 1;
      }
    }

    return cols;
  } catch (_e) {
    Logger.log('detectColumnLayout_ error: ' + _e.message);
    return null;
  }
}

/**
 * Sync all column maps with actual sheet headers at runtime.
 * Auto-updates global *_COLS objects if columns have been moved.
 * Call during onOpen or after sheet restructuring.
 * @returns {Object} { warnings: string[], synced: string[] }
 */
function syncColumnMaps() {
  var result = { warnings: [], synced: [] };

  var maps = [
    { name: 'MEMBER_COLS', sheet: SHEETS.MEMBER_DIR, map: MEMBER_HEADER_MAP_, target: MEMBER_COLS,
      aliases: { LOCATION: 'WORK_LOCATION', DAYS_TO_DEADLINE: 'NEXT_DEADLINE' } },
    { name: 'GRIEVANCE_COLS', sheet: SHEETS.GRIEVANCE_LOG, map: GRIEVANCE_HEADER_MAP_, target: GRIEVANCE_COLS },
    { name: 'CONFIG_COLS', sheet: SHEETS.CONFIG, map: CONFIG_HEADER_MAP_, target: CONFIG_COLS,
      opts: { headerRow: 2 } },
    { name: 'MEETING_CHECKIN_COLS', sheet: SHEETS.MEETING_CHECKIN_LOG, map: MEETING_CHECKIN_HEADER_MAP_, target: MEETING_CHECKIN_COLS },
    { name: 'STEWARD_PERF_COLS', sheet: SHEETS.STEWARD_PERFORMANCE_CALC, map: STEWARD_PERF_HEADER_MAP_, target: STEWARD_PERF_COLS },
    { name: 'AUDIT_LOG_COLS', sheet: SHEETS.AUDIT_LOG, map: AUDIT_LOG_HEADER_MAP_, target: AUDIT_LOG_COLS },
    { name: 'SURVEY_VAULT_COLS', sheet: SHEETS.SURVEY_VAULT, map: SURVEY_VAULT_HEADER_MAP_, target: SURVEY_VAULT_COLS },
    { name: 'SURVEY_TRACKING_COLS', sheet: SHEETS.SURVEY_TRACKING, map: SURVEY_TRACKING_HEADER_MAP_, target: SURVEY_TRACKING_COLS },
    { name: 'FEEDBACK_COLS', sheet: SHEETS.FEEDBACK, map: FEEDBACK_HEADER_MAP_, target: FEEDBACK_COLS },
    { name: 'CHECKLIST_COLS', sheet: SHEETS.CASE_CHECKLIST, map: CHECKLIST_HEADER_MAP_, target: CHECKLIST_COLS },
    { name: 'RESOURCES_COLS', sheet: SHEETS.RESOURCES, map: RESOURCES_HEADER_MAP_, target: RESOURCES_COLS },
    { name: 'RESOURCE_CONFIG_COLS', sheet: SHEETS.RESOURCE_CONFIG, map: RESOURCE_CONFIG_HEADER_MAP_, target: RESOURCE_CONFIG_COLS }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var m = 0; m < maps.length; m++) {
    var entry = maps[m];
    var sheet = ss ? ss.getSheetByName(entry.sheet) : null;

    // Backfill missing columns — older deployments may lack columns added in
    // later versions. Detect unresolved headers and append them so all sheets
    // converge to the expected column count automatically on next onOpen.
    if (sheet && sheet.getLastColumn() > 0) {
      try {
        var hRow = (entry.opts && entry.opts.headerRow) || 1;
        var sheetLastCol = sheet.getLastColumn();
        var sheetHeaders = sheet.getRange(hRow, 1, 1, sheetLastCol).getValues()[0];
        var headerSet = {};
        for (var sh = 0; sh < sheetHeaders.length; sh++) {
          var hText = String(sheetHeaders[sh]).trim();
          if (hText) headerSet[hText] = true;
        }

        var missingCols = [];
        for (var hm = 0; hm < entry.map.length; hm++) {
          if (!headerSet[entry.map[hm].header]) missingCols.push(entry.map[hm]);
        }

        if (missingCols.length > 0) {
          var appendAt = sheetLastCol + 1;
          ensureMinimumColumns(sheet, sheetLastCol + missingCols.length);
          for (var mc = 0; mc < missingCols.length; mc++) {
            sheet.getRange(hRow, appendAt + mc).setValue(missingCols[mc].header)
              .setBackground(COMMAND_CONFIG.THEME.HEADER_BG)
              .setFontColor('#ffffff')
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
          }
          result.warnings.push(entry.name + ': backfilled ' + missingCols.length + ' missing columns (' +
            missingCols.map(function(c) { return c.key; }).join(', ') + ')');
          result.synced.push(entry.name + '_BACKFILL');
          Logger.log('syncColumnMaps: Added ' + missingCols.length + ' missing columns to ' + entry.sheet + ': ' +
            missingCols.map(function(c) { return c.header; }).join(', '));
        }
      } catch (_bfErr) {
        Logger.log('syncColumnMaps backfill ' + entry.sheet + ': ' + _bfErr.message);
      }
    }

    var resolved = resolveColumnsFromSheet_(entry.sheet, entry.map, entry.opts);
    if (!resolved) continue;

    var moved = false;
    for (var key in resolved) {
      if (resolved.hasOwnProperty(key) && entry.target[key] !== resolved[key]) {
        result.warnings.push(entry.name + '.' + key + ': expected col ' + entry.target[key] + ', found col ' + resolved[key]);
        entry.target[key] = resolved[key];
        moved = true;
      }
    }

    // Re-apply aliases after sync
    if (entry.aliases) {
      for (var alias in entry.aliases) {
        if (entry.aliases.hasOwnProperty(alias)) {
          entry.target[alias] = entry.target[entry.aliases[alias]];
        }
      }
    }

    if (moved) result.synced.push(entry.name);
  }

  // Rebuild derived column configs so dropdown dialogs and bidirectional
  // sync use the up-to-date column positions after any columns shifted.
  if (result.synced.length > 0) {
    var freshMulti = buildMultiSelectCols_();
    MULTI_SELECT_COLS.MEMBER_DIR = freshMulti.MEMBER_DIR;
    MULTI_SELECT_COLS.GRIEVANCE_LOG = freshMulti.GRIEVANCE_LOG;

    var freshDD = buildDropdownMap_();
    DROPDOWN_MAP.MEMBER_DIR = freshDD.MEMBER_DIR;
    DROPDOWN_MAP.GRIEVANCE_LOG = freshDD.GRIEVANCE_LOG;

    // Rebuild JOB_METADATA_FIELDS so it reflects the resolved column positions.
    // Without this, functions using getJobMetadataByMemberCol() would use stale
    // column numbers captured at script load time.
    rebuildJobMetadataFields_();
  }

  if (result.synced.length > 0) {
    Logger.log('syncColumnMaps: Updated ' + result.synced.join(', '));
    if (result.warnings.length > 0) {
      Logger.log('Column changes detected:\n  ' + result.warnings.join('\n  '));
    }
    // Column positions changed — invalidate ConfigReader's cache so it re-reads
    // the sheet with correct column positions instead of serving stale values.
    try { CacheService.getScriptCache().remove('ORG_CONFIG_v2'); } catch (_) {}
  }

  // Persist resolved positions so other execution contexts (onEdit, etc.)
  // pick them up without re-reading sheet headers every time.
  persistColumnMaps_();

  return result;
}

// ============================================================================
// COLUMN MAP PERSISTENCE
// ============================================================================
// In Google Apps Script, each trigger execution (onEdit, onOpen, menu click)
// starts a brand-new V8 isolate.  Global column constants re-initialize to
// their default (array-order) positions.  If a user manually reordered
// columns, onEdit would silently use wrong positions until the next onOpen.
//
// Solution: syncColumnMaps() persists resolved positions to CacheService
// (2-hour TTL).  loadCachedColumnMaps_() restores them cheaply in onEdit.
// ============================================================================

/** Cache key for persisted column maps */
var COL_MAPS_CACHE_KEY_ = 'RESOLVED_COL_MAPS';

/**
 * Persist current column constant values to CacheService.
 * Called at the end of syncColumnMaps().
 * @private
 */
function persistColumnMaps_() {
  try {
    var data = {
      MEMBER_COLS: MEMBER_COLS,
      GRIEVANCE_COLS: GRIEVANCE_COLS,
      CONFIG_COLS: CONFIG_COLS,
      MEETING_CHECKIN_COLS: MEETING_CHECKIN_COLS,
      AUDIT_LOG_COLS: AUDIT_LOG_COLS,
      SURVEY_VAULT_COLS: SURVEY_VAULT_COLS,
      SURVEY_TRACKING_COLS: SURVEY_TRACKING_COLS,
      STEWARD_PERF_COLS: STEWARD_PERF_COLS,
      FEEDBACK_COLS: FEEDBACK_COLS,
      CHECKLIST_COLS: CHECKLIST_COLS
    };
    CacheService.getScriptCache().put(COL_MAPS_CACHE_KEY_, JSON.stringify(data), 21600); // 6 hours (max)
  } catch (_e) {
    // CacheService unavailable — degrade silently; defaults are still valid
    // for sheets created by this code.
  }
}

/**
 * Load persisted column positions from CacheService and apply to globals.
 * Returns true if cache was found and applied, false otherwise.
 * Call at the top of onEdit() to cheaply pick up any positions that
 * syncColumnMaps() resolved in a prior execution.
 * @returns {boolean} Whether cached positions were applied
 * @private
 */
function loadCachedColumnMaps_() {
  try {
    var json = CacheService.getScriptCache().get(COL_MAPS_CACHE_KEY_);
    if (!json) return false;
    var data = JSON.parse(json);

    // Apply to primary column constants
    var targets = {
      MEMBER_COLS: MEMBER_COLS,
      GRIEVANCE_COLS: GRIEVANCE_COLS,
      CONFIG_COLS: CONFIG_COLS,
      MEETING_CHECKIN_COLS: MEETING_CHECKIN_COLS,
      AUDIT_LOG_COLS: AUDIT_LOG_COLS,
      SURVEY_VAULT_COLS: SURVEY_VAULT_COLS,
      SURVEY_TRACKING_COLS: SURVEY_TRACKING_COLS,
      STEWARD_PERF_COLS: STEWARD_PERF_COLS,
      FEEDBACK_COLS: FEEDBACK_COLS,
      CHECKLIST_COLS: CHECKLIST_COLS
    };

    var changed = false;
    for (var name in targets) {
      if (data[name]) {
        for (var key in data[name]) {
          if (data[name].hasOwnProperty(key) && targets[name].hasOwnProperty(key)
              && targets[name][key] !== data[name][key]) {
            targets[name][key] = data[name][key];
            changed = true;
          }
        }
      }
    }

    // Rebuild derived objects (dropdown map, multi-select, job metadata)
    // only if something actually changed.
    if (changed) {
      var freshMulti = buildMultiSelectCols_();
      MULTI_SELECT_COLS.MEMBER_DIR = freshMulti.MEMBER_DIR;
      MULTI_SELECT_COLS.GRIEVANCE_LOG = freshMulti.GRIEVANCE_LOG;

      var freshDD = buildDropdownMap_();
      DROPDOWN_MAP.MEMBER_DIR = freshDD.MEMBER_DIR;
      DROPDOWN_MAP.GRIEVANCE_LOG = freshDD.GRIEVANCE_LOG;

      rebuildJobMetadataFields_();
    }

    return true;
  } catch (_e) {
    return false; // Cache unavailable — use defaults
  }
}

// ============================================================================
// SHEET COLUMN GUARD
// ============================================================================

/**
 * Ensure a sheet has at least the minimum required columns.
 * Moved here (from 07_DevTools) so it is available to every source file
 * regardless of load order.
 */
function ensureMinimumColumns(sheet, requiredColumns) {
  var currentColumns = sheet.getMaxColumns();
  if (currentColumns < requiredColumns) {
    var columnsToAdd = requiredColumns - currentColumns;
    sheet.insertColumnsAfter(currentColumns, columnsToAdd);
    Logger.log('Added ' + columnsToAdd + ' columns to ' + sheet.getName() + ' (now has ' + requiredColumns + ' columns)');
  }
}

/**
 * Ensures the three primary sheets (Member Directory, Grievance Log, Config)
 * have enough columns for all defined headers. Call from onOpen() or before
 * any operation that accesses high column numbers.
 *
 * Safe to call multiple times — ensureMinimumColumns is a no-op when the
 * sheet already has sufficient columns.
 */
function ensureAllSheetColumns_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checks = [
    { name: SHEETS.MEMBER_DIR,    count: getMemberHeaders().length },
    { name: SHEETS.GRIEVANCE_LOG, count: getGrievanceHeaders().length },
    { name: SHEETS.CONFIG,        count: getHeadersFromMap_(CONFIG_HEADER_MAP_).length }
  ];
  for (var i = 0; i < checks.length; i++) {
    var sheet = ss.getSheetByName(checks[i].name);
    if (sheet) {
      ensureMinimumColumns(sheet, checks[i].count);
    }
  }
}

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
 * Safe accessor for numeric fields — returns the value as-is when it's a
 * number (including 0), and '' for null/undefined/empty-string.
 * Avoids the `value || ''` pitfall where 0 is treated as falsy.
 * @param {*} value - Cell value
 * @returns {number|string}
 * @private
 */
function numericField_(value) {
  if (value === null || value === undefined || value === '') return '';
  return value;
}

/**
 * Get all member header labels in order — auto-derived from MEMBER_HEADER_MAP_
 * @returns {Array<string>} Ordered header labels for Member Directory
 */
function getMemberHeaders() {
  return getHeadersFromMap_(MEMBER_HEADER_MAP_);
}

/**
 * Get all grievance header labels in order — auto-derived from GRIEVANCE_HEADER_MAP_
 * @returns {Array<string>} Ordered header labels for Grievance Log
 */
function getGrievanceHeaders() {
  return getHeadersFromMap_(GRIEVANCE_HEADER_MAP_);
}

// ============================================================================
// VALIDATION VALUES
// ============================================================================

/**
 * Default values for Config sheet dropdowns
 */
var DEFAULT_CONFIG = {
  OFFICE_DAYS: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
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
 * Statuses that indicate a grievance is no longer active.
 * Used to filter open/closed grievances across dashboards and web views.
 * Values are Title Case for sheet comparisons; lowercase with .map(s => s.toLowerCase()) when needed.
 * @const {string[]}
 */
var GRIEVANCE_CLOSED_STATUSES = [
  GRIEVANCE_STATUS.CLOSED,
  GRIEVANCE_STATUS.SETTLED,
  GRIEVANCE_STATUS.WITHDRAWN,
  GRIEVANCE_STATUS.DENIED,
  GRIEVANCE_STATUS.WON
];

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
 *
 * SOURCE: Sample CBA grievance deadlines — adjust per your local contract.
 * Policy 270.03 for Bargaining Units 8 & 10 (July 2025 edition).
 *
 * CONTRACT SUMMARY FOR UNIT 8 (only 2 steps + arbitration):
 *   Filing:        21 calendar days from incident
 *   Step I resp:   30 calendar days (21 if meeting held)
 *   Step II appeal: 10 BUSINESS days from Step I decision  ← note: business days
 *   Step II resp:  30 calendar days (21 if conference held)
 *   Arbitration:   30 calendar days from Step II
 *   NOTE: Unit 8 CBA has NO Step III — goes Step I → Step II → Arbitration.
 *         STEP_3_* fields below are retained for configurability but are not
 *         required by this CBA.  If your local contract differs, override via
 *         the Config sheet.
 *
 * FIX-CORE-01 (v4.25.9): Corrected STEP_1_RESPONSE (7→30), STEP_2_APPEAL
 *   (7→10), STEP_2_RESPONSE (14→30) to match verified contract language.
 *   Old hardcoded values were wrong; Config sheet must be re-seeded.
 * @const {Object}
 */
var DEADLINE_DEFAULTS = {
  FILING_DAYS: 21,          // Art. 23: 21 calendar days to file
  STEP_1_RESPONSE: 30,      // Art. 23: 30 calendar days for Step I response (21 if meeting)
  STEP_2_APPEAL: 10,        // Art. 23: 10 BUSINESS days to appeal to Step II
  STEP_2_RESPONSE: 30,      // Art. 23: 30 calendar days for Step II response (21 if conference)
  STEP_3_APPEAL: 10,        // Not in Unit 8 CBA — retained for configurable deployments
  STEP_3_RESPONSE: 21,      // Not in Unit 8 CBA — retained for configurable deployments
  ARBITRATION_DEMAND: 30,   // Art. 23: 30 calendar days from Step II to demand arbitration
  WARNING_THRESHOLD: 5,
  CRITICAL_THRESHOLD: 2,
  REMINDER_FIRST: 7,
  REMINDER_SECOND: 3,
  REMINDER_FINAL: 1
};

/**
 * Reads deadline rules from Config sheet, falling back to DEADLINE_DEFAULTS.
 * This is the single source of truth for all deadline calculations.
 * Covers: Filing, Step I, Step II, Step III, and Arbitration.
 * @returns {Object} DEADLINE_RULES-compatible object
 */
function getDeadlineRules() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      var colStart = CONFIG_COLS.FILING_DEADLINE_DAYS;
      var colEnd   = CONFIG_COLS.ARBITRATION_DEMAND_DAYS;
      var dVals    = configSheet.getRange(3, colStart, 1, colEnd - colStart + 1).getValues()[0];
      var filing     = Number(dVals[0]);
      var s1Resp     = Number(dVals[CONFIG_COLS.STEP1_RESPONSE_DAYS - colStart]);
      var s2Appeal   = Number(dVals[CONFIG_COLS.STEP2_APPEAL_DAYS   - colStart]);
      var s2Resp     = Number(dVals[CONFIG_COLS.STEP2_RESPONSE_DAYS - colStart]);
      var s3Appeal   = Number(dVals[CONFIG_COLS.STEP3_APPEAL_DAYS   - colStart]);
      var s3Resp     = Number(dVals[CONFIG_COLS.STEP3_RESPONSE_DAYS - colStart]);
      var arbDemand  = Number(dVals[CONFIG_COLS.ARBITRATION_DEMAND_DAYS - colStart]);
      return {
        FILING_DAYS: isNaN(filing) ? DEADLINE_DEFAULTS.FILING_DAYS : filing,
        STEP_1: { DAYS_FOR_RESPONSE: isNaN(s1Resp) ? DEADLINE_DEFAULTS.STEP_1_RESPONSE : s1Resp },
        STEP_2: { DAYS_TO_APPEAL: isNaN(s2Appeal) ? DEADLINE_DEFAULTS.STEP_2_APPEAL : s2Appeal, DAYS_FOR_RESPONSE: isNaN(s2Resp) ? DEADLINE_DEFAULTS.STEP_2_RESPONSE : s2Resp },
        STEP_3: { DAYS_TO_APPEAL: isNaN(s3Appeal) ? DEADLINE_DEFAULTS.STEP_3_APPEAL : s3Appeal, DAYS_FOR_RESPONSE: isNaN(s3Resp) ? DEADLINE_DEFAULTS.STEP_3_RESPONSE : s3Resp },
        ARBITRATION: { DAYS_TO_DEMAND: isNaN(arbDemand) ? DEADLINE_DEFAULTS.ARBITRATION_DEMAND : arbDemand }
      };
    }
  } catch (e) {
    Logger.log('Error reading deadline config: ' + e.message);
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
  TRIGGER_REMOVED: 'TRIGGER_REMOVED',

  // Resource events (v4.32.1)
  // Currently used as a constant key for future audit integration.
  // dataLogResourceClick() writes directly to _Resource_Click_Log (not _Audit_Log)
  // for performance — click events are high-frequency and don't need integrity hashing.
  RESOURCE_CLICKED: 'RESOURCE_CLICKED'
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

/**
 * Scale thresholds — triggers adaptive behavior as sheet row counts grow.
 * Health monitor in _getCachedSheetData logs warnings at WARN and switches
 * to paginated mode at THROTTLE. Frontend adapts via dataGetSheetHealth.
 * @const {Object}
 */
var SCALE_THRESHOLDS = {
  WARN_ROWS:       5000,   // Log warning — sheet approaching performance limits
  THROTTLE_ROWS:   7000,   // Switch to paginated mode automatically
  CRITICAL_ROWS:   8000    // Log critical alert — manual intervention recommended
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
  { label: 'Committees', memberCol: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, configName: 'Steward Committees' }
];
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

/**
 * Rebuild JOB_METADATA_FIELDS from the current *_COLS values.
 * Called by syncColumnMaps() and loadCachedColumnMaps_() so that
 * JOB_METADATA_FIELDS stays in sync when columns shift at runtime.
 * @private
 */
function rebuildJobMetadataFields_() {
  JOB_METADATA_FIELDS.length = 0; // clear in-place to preserve the reference
  JOB_METADATA_FIELDS.push(
    { label: 'Job Title', memberCol: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES, configName: 'Job Titles' },
    { label: 'Work Location', memberCol: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS, configName: 'Office Locations' },
    { label: 'Unit', memberCol: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS, configName: 'Units' },
    { label: 'Supervisor', memberCol: MEMBER_COLS.SUPERVISOR, configCol: CONFIG_COLS.SUPERVISORS, configName: 'Supervisors' },
    { label: 'Manager', memberCol: MEMBER_COLS.MANAGER, configCol: CONFIG_COLS.MANAGERS, configName: 'Managers' },
    { label: 'Assigned Steward', memberCol: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, configName: 'Stewards' },
    { label: 'Committees', memberCol: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, configName: 'Steward Committees' }
  );
}

// ============================================================================
// MULTI-SELECT COLUMN CONFIGURATION
// ============================================================================

/**
 * Build multi-select column config from current *_COLS values.
 * Returns fresh references every call so that syncColumnMaps() changes
 * are reflected without a script reload.
 * @returns {Object} { MEMBER_DIR: [...], GRIEVANCE_LOG: [...] }
 */
function buildMultiSelectCols_() {
  return {
    MEMBER_DIR: [
      { col: MEMBER_COLS.OFFICE_DAYS, configCol: CONFIG_COLS.OFFICE_DAYS, label: 'Office Days' },
      { col: MEMBER_COLS.PREFERRED_COMM, configCol: CONFIG_COLS.COMM_METHODS, label: 'Preferred Communication' },
      { col: MEMBER_COLS.BEST_TIME, configCol: CONFIG_COLS.BEST_TIMES, label: 'Best Time to Contact' },
      { col: MEMBER_COLS.COMMITTEES, configCol: CONFIG_COLS.STEWARD_COMMITTEES, label: 'Committees' },
      { col: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS, label: 'Assigned Steward(s)' }
    ],
    GRIEVANCE_LOG: [
      { col: GRIEVANCE_COLS.ARTICLES, configCol: CONFIG_COLS.ARTICLES, label: 'Articles Violated' },
      { col: GRIEVANCE_COLS.ISSUE_CATEGORY, configCol: CONFIG_COLS.ISSUE_CATEGORY, label: 'Issue Category' }
    ]
  };
}

/**
 * Columns that support multiple selections (comma-separated values).
 * Initialized at load time; refreshed by syncColumnMaps().
 */
var MULTI_SELECT_COLS = buildMultiSelectCols_();

/**
 * Check if a column is a multi-select column for the given sheet.
 * Reads from the live MULTI_SELECT_COLS object (updated by syncColumnMaps).
 * @param {number} col - Column number (1-indexed)
 * @param {string} sheetName - Sheet name (defaults to Member Directory)
 * @returns {Object|null} Multi-select config if found, null otherwise
 */
function getMultiSelectConfig(col, sheetName) {
  var configs;
  if (sheetName === SHEETS.GRIEVANCE_LOG) {
    configs = MULTI_SELECT_COLS.GRIEVANCE_LOG;
  } else {
    configs = MULTI_SELECT_COLS.MEMBER_DIR;
  }
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].col === col) {
      return configs[i];
    }
  }
  return null;
}

// ============================================================================
// DROPDOWN COLUMN MAP — Single source of truth
// ============================================================================
// Both setupDataValidations() and syncDropdownToConfig_() derive their
// column-to-config mappings from these arrays.  When you add a new dropdown
// column, add it HERE and everything else follows.
//
// Each entry:  { col: <target sheet col>, configCol: <Config sheet col> }
//   • 'multi' entries use multi-select validation (comma-separated values)
//   • 'single' entries use normal dropdown validation
// ============================================================================

/**
 * Build the single-select dropdown map from current *_COLS values.
 * Returns fresh references so syncColumnMaps() changes are reflected.
 * @returns {Object} { MEMBER_DIR: [...], GRIEVANCE_LOG: [...] }
 */
function buildDropdownMap_() {
  return {
    MEMBER_DIR: [
      { col: MEMBER_COLS.JOB_TITLE,        configCol: CONFIG_COLS.JOB_TITLES },
      { col: MEMBER_COLS.WORK_LOCATION,     configCol: CONFIG_COLS.OFFICE_LOCATIONS },
      { col: MEMBER_COLS.UNIT,              configCol: CONFIG_COLS.UNITS },
      // IS_STEWARD and INTEREST_* columns deliberately excluded — they use hardcoded
      // validation ('Yes'/'No'), not a Config column.  The YES_NO Config column was
      // removed to eliminate contamination risk.  Steward status sync is handled by
      // handleMemberEdit() and syncStewardStatus(), which write to CONFIG_COLS.STEWARDS.
      { col: MEMBER_COLS.SUPERVISOR,        configCol: CONFIG_COLS.SUPERVISORS },
      { col: MEMBER_COLS.MANAGER,           configCol: CONFIG_COLS.MANAGERS },
      { col: MEMBER_COLS.CONTACT_STEWARD,   configCol: CONFIG_COLS.STEWARDS },
      { col: MEMBER_COLS.DUES_STATUS,        configCol: CONFIG_COLS.DUES_STATUSES }
    ],
    // ISSUE_CATEGORY and ARTICLES are multi-select columns (comma-separated values).
    // They live in MULTI_SELECT_COLS.GRIEVANCE_LOG, NOT here.
    // Keeping them here caused setupDataValidations() to apply single-select first,
    // then multi-select would overwrite — wasting a sheet API call and risking the
    // wrong validation type if execution order ever changed.
    GRIEVANCE_LOG: [
      { col: GRIEVANCE_COLS.STATUS,         configCol: CONFIG_COLS.GRIEVANCE_STATUS },
      { col: GRIEVANCE_COLS.CURRENT_STEP,   configCol: CONFIG_COLS.GRIEVANCE_STEP }
    ]
  };
}

/**
 * Single-select dropdown map.  Initialized at load; refreshed by syncColumnMaps().
 */
var DROPDOWN_MAP = buildDropdownMap_();

// ── Eagerly restore cached column positions ──
// Every GAS execution starts with array-order defaults from buildColsFromMap_.
// After CONFIG_HEADER_MAP_ was reordered (v4.50.0), these defaults no longer
// match the actual sheet layout.  syncColumnMaps() resolves positions at runtime
// but only runs in onOpen().  This call restores cached positions so that ALL
// execution contexts (doGet, data* web functions, onEdit) use correct columns.
try {
  if (!loadCachedColumnMaps_()) {
    // Cache is cold — column defaults may be wrong after CONFIG_HEADER_MAP_ reorder.
    // Run a full sync to resolve positions from actual sheet headers.
    // This covers ALL execution contexts (doGet, data* web functions, menu handlers,
    // onEdit) not just doGetWebDashboard and onOpen.
    syncColumnMaps();
  }
} catch (_initCache) { /* CacheService or sheet unavailable — defaults used */ }

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
  // CR-OTHER-1: Use Utilities.getUuid() for cryptographically secure UUIDs
  // instead of Math.random() which has insufficient entropy
  return Utilities.getUuid();
}

// ============================================================================
// ACTION TYPE CONFIGURATION (Grievances + Other Actions)
// ============================================================================

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
// ============================================================================
// CHECKLIST CONFIGURATION
// ============================================================================

/**
 * Checklist sheet name
 * @const {string}
 */
var CHECKLIST_SHEET_NAME = SHEETS.CASE_CHECKLIST;

/**
 * Checklist column positions (1-indexed)
 * @const {Object}
 */
var CHECKLIST_HEADER_MAP_ = [
  { key: 'CHECKLIST_ID',   header: 'Checklist ID' },
  { key: 'CASE_ID',        header: 'Case ID' },
  { key: 'ACTION_TYPE',    header: 'Action Type' },
  { key: 'ITEM_TEXT',      header: 'Item Text' },
  { key: 'CATEGORY',       header: 'Category' },
  { key: 'REQUIRED',       header: 'Required' },
  { key: 'COMPLETED',      header: 'Completed' },
  { key: 'COMPLETED_BY',   header: 'Completed By' },
  { key: 'COMPLETED_DATE', header: 'Completed Date' },
  { key: 'DUE_DATE',       header: 'Due Date' },
  { key: 'NOTES',          header: 'Notes' },
  { key: 'SORT_ORDER',     header: 'Sort Order' }
];

var CHECKLIST_COLS = buildColsFromMap_(CHECKLIST_HEADER_MAP_);

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

      '/* H-4 fix: Haptic feedback for google.script.run — preserves fluent chaining */' +
      'var origRun=google.script&&google.script.run;' +
      'if(origRun){' +
        'var origSuccess=origRun.withSuccessHandler;' +
        'var origFailure=origRun.withFailureHandler;' +
        'if(origSuccess){' +
          'google.script.run.withSuccessHandler=function(fn){' +
            'var result=origSuccess.call(this,function(){' +
              'hapticFeedback("success");' +
              'if(fn)fn.apply(this,arguments);' +
            '});' +
            'return result||this;' +
          '};' +
        '}' +
        'if(origFailure){' +
          'google.script.run.withFailureHandler=function(fn){' +
            'var result=origFailure.call(this,function(){' +
              'hapticFeedback("error");' +
              'if(fn)fn.apply(this,arguments);' +
            '});' +
            'return result||this;' +
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

// ============================================================================
// ID PATTERN VALIDATORS + SHARED HELPERS
// ============================================================================
// Centralized ID format checks. All modules must use these instead of inline
// regex like match(/^G/i) or match(/^M/i) which scatter business rules.

/** @const {RegExp} Grievance IDs start with 'G' (case-insensitive) */
var GRIEVANCE_ID_PATTERN = /^G/i;
/** @const {RegExp} Member IDs start with 'M' (case-insensitive) */
var MEMBER_ID_PATTERN = /^M/i;

/**
 * Returns true if the value looks like a valid Grievance ID.
 * @param {*} id
 * @returns {boolean}
 */
function isGrievanceId_(id) {
  return !!id && GRIEVANCE_ID_PATTERN.test(String(id));
}

/**
 * Returns true if the value looks like a valid Member ID.
 * @param {*} id
 * @returns {boolean}
 */
function isMemberId_(id) {
  return !!id && MEMBER_ID_PATTERN.test(String(id));
}

// ============================================================================
// SHARED DATE/HASH HELPERS
// ============================================================================

/**
 * Formats a Date object as "Mon DD, YYYY" (e.g., "Mar 12, 2026").
 * @param {Date} date
 * @returns {string}
 */
function fmtDateShort_(date) {
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
}

/**
 * Returns a truncated SHA-256 hash of an email (12 hex chars).
 * Used for anonymous attribution without storing PII.
 * @param {string} email
 * @returns {string}
 */
function hashEmail_(email) {
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email.toLowerCase().trim());
  return hash.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 12);
}

// ============================================================================
// SHEET STYLING COLORS
// ============================================================================

/**
 * Centralized sheet styling colors — single source of truth for all setBackground/setFontColor calls.
 * Google Sheets API doesn't support CSS variables, but centralizing here makes branding changes a single-file edit.
 */
var SHEET_COLORS = {
  // Header backgrounds (dark)
  HEADER_NAVY: '#1a1a2e',
  HEADER_BLUE: '#0d47a1',
  HEADER_SLATE: '#1e293b',
  HEADER_SLATE_MED: '#334155',
  HEADER_DARK_BLUE: '#1a365d',
  HEADER_DARK_RED: '#7F1D1D',
  HEADER_DARK_PURPLE: '#5B21B6',
  HEADER_DARK_BLUE_ALT: '#1E40AF',

  // Status indicators
  STATUS_SUCCESS: '#059669',
  STATUS_WARNING: '#FBBF24',
  STATUS_ERROR: '#DC2626',
  STATUS_ERROR_DARK: '#C62828',
  STATUS_INFO: '#3B82F6',
  STATUS_DISABLED: '#6B7280',
  STATUS_PURPLE: '#7C3AED',

  // Light backgrounds (pastel/tinted)
  BG_WHITE: '#ffffff',
  BG_LIGHT_GRAY: '#F9FAFB',
  BG_VERY_LIGHT_GRAY: '#F3F4F6',
  BG_OFF_WHITE: '#f8fafc',
  BG_LIGHT_BLUE: '#E0E7FF',
  BG_LIGHT_BLUE_ALT: '#e3f2fd',
  BG_PALE_BLUE: '#F0F9FF',
  BG_EXTRA_PALE_BLUE: '#EFF6FF',
  BG_LINK_BLUE: '#e8f0fe',
  BG_LIGHT_GREEN: '#E8F5E9',
  BG_PALE_GREEN: '#D1FAE5',
  BG_GREEN_ALT: '#C8E6C9',
  BG_MINT: '#ECFDF5',
  BG_LIGHT_RED: '#FEE2E2',
  BG_LIGHT_RED_ALT: '#FFCDD2',
  BG_PINK: '#FFE2E2',
  BG_LIGHT_YELLOW: '#FEF3C7',
  BG_PALE_YELLOW: '#FFF9C4',
  BG_EXTRA_PALE_YELLOW: '#fffde7',
  BG_CREAM: '#fff8e1',
  BG_LIGHT_ORANGE: '#FFE0B2',
  BG_WARM: '#FFF3E0',
  BG_LIGHT_PURPLE: '#FAF5FF',
  BG_SLATE_LIGHT: '#F1F5F9',

  // Text colors
  TEXT_WHITE: '#ffffff',
  TEXT_GRAY: '#6B7280',
  TEXT_LIGHT_GRAY: '#9CA3AF',
  TEXT_DARK_GREEN: '#065F46',
  TEXT_GREEN: '#2e7d32',
  TEXT_GREEN_ALT: '#166534',
  TEXT_GREEN_DARK: '#1B5E20',
  TEXT_DARK_RED: '#B71C1C',
  TEXT_RED: '#c62828',
  TEXT_DARK_ORANGE: '#92400E',
  TEXT_ORANGE: '#E65100',
  TEXT_YELLOW_DARK: '#F57F17',
  TEXT_BROWN: '#78350F',
  TEXT_MED_GRAY: '#666666',

  // Link colors
  LINK_PRIMARY: '#1a73e8',
  LINK_SECONDARY: '#1155cc',

  // Dialog accent — matches webapp purple fallback; use for dialog buttons/headers
  DIALOG_ACCENT: '#7C3AED',
  DIALOG_ACCENT_DARK: '#5B21B6',

  // Union Brand Theme (v4.35.0 — global tab formatting)
  THEME_NAVY: '#1A2A4A',            // Primary header bg
  THEME_GOLD: '#D4A017',            // Accent / highlight
  THEME_STEEL_BLUE: '#2C5282',      // Section header bg
  THEME_LIGHT_SKY: '#EBF4FF',       // Alternate row bg
  THEME_GREEN: '#276749',           // Completed / positive
  THEME_AMBER: '#C05621',           // Warning / overdue
  THEME_RED: '#9B2335',             // Critical / urgent

  // Tab bar color groups (by function)
  TAB_BLUE: '#1565C0',              // Data entry tabs
  TAB_GREEN: '#2E7D32',             // Engagement / survey tabs
  TAB_GOLD: '#F9A825',              // Documentation / guide tabs
  TAB_RED_ORANGE: '#C62828'         // Admin / technical tabs
};
