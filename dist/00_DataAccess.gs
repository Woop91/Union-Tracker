// GAS_LIMITATION: All .gs files share a single global namespace. Function names must be globally unique.
/**
 * ============================================================================
 * 00_DataAccess.gs - Shared Time Constants and Lock Helper
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Provides three foundational utilities used across the entire codebase:
 *   1. TIME_CONSTANTS — millisecond conversion factors and grievance deadline
 *      day counts (filing windows, appeal periods, reminder intervals).
 *   2. withScriptLock_(fn) — executes a function under a GAS script-level lock
 *      to prevent concurrent writes from corrupting sheet data.
 *   3. _getSheetSafe(name) — a null-safe sheet accessor that logs warnings
 *      instead of silently returning null when a sheet is missing.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   - Loaded first (00_ prefix) so every other module can depend on these
 *     constants without worrying about load order. GAS loads .gs files
 *     alphabetically by filename, so the 00_ prefix guarantees this file
 *     is evaluated before anything in 01_ through 31_.
 *   - TIME_CONSTANTS.DEADLINE_DAYS uses JavaScript getters that reference
 *     DEADLINE_DEFAULTS from 01_Core.gs. This creates a single source of
 *     truth in 01_Core.gs while allowing 00_DataAccess.gs to provide safe
 *     fallback values (e.g., 21 days) if 01_Core.gs hasn't loaded yet.
 *   - withScriptLock_ uses tryLock() instead of waitLock() because tryLock
 *     is non-blocking with a max-wait timeout, preventing indefinite hangs.
 *   - Previously housed a full DataAccess namespace (sheet read/write layer).
 *     That was removed when it had zero callers — modules now use direct
 *     SpreadsheetApp calls via SHEETS constants from 01_Core.gs.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   - If TIME_CONSTANTS is missing: deadline calculations in 08c, 05_, and
 *     12_ will throw ReferenceErrors. All grievance deadline tracking stops.
 *   - If withScriptLock_ is missing: 02_DataManagers.gs (addMember,
 *     updateMember, createGrievance) and 06_Maintenance.gs will throw
 *     ReferenceErrors on any write operation. No data can be modified.
 *   - If _getSheetSafe is missing: callers fall back to raw getSheetByName()
 *     which returns null silently — harder to debug missing-sheet issues.
 *
 * DEPENDENCIES:
 *   Depends on:  LockService (GAS built-in), SpreadsheetApp (GAS built-in)
 *                DEADLINE_DEFAULTS (01_Core.gs, optional — has fallback values)
 *   Used by:     02_DataManagers.gs, 06_Maintenance.gs, 08c_FormsAndNotifications.gs,
 *                05_Integrations.gs, 12_Features.gs, test suites
 *
 * @fileoverview Shared constants and lock helper — loaded first in build order
 * @version 4.51.0
 */

// ============================================================================
// TIME CONSTANTS
// ============================================================================

/**
 * Time-related constants for deadline calculations
 * @const {Object}
 */
var TIME_CONSTANTS = {
  /** Milliseconds per second */
  MS_PER_SECOND: 1000,

  /** Milliseconds per minute */
  MS_PER_MINUTE: 60 * 1000,

  /** Milliseconds per hour */
  MS_PER_HOUR: 60 * 60 * 1000,

  /** Milliseconds per day */
  MS_PER_DAY: 24 * 60 * 60 * 1000,

  /** Milliseconds per week */
  MS_PER_WEEK: 7 * 24 * 60 * 60 * 1000,

  /** Deadline days configuration — references DEADLINE_DEFAULTS (01_Core.gs) as single source of truth */
  DEADLINE_DAYS: {
    /** Days to file grievance after incident */
    get FILING() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.FILING_DAYS : 21; },

    /** Days for Step I response */
    get STEP1_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_1_RESPONSE : 30; },

    /** Days to appeal to Step II */
    get STEP2_APPEAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_APPEAL : 10; },

    /** Days for Step II response */
    get STEP2_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_RESPONSE : 30; },

    /** Days to appeal to Step III */
    get STEP3_APPEAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_3_APPEAL : 10; },

    /** Days before deadline to show warning */
    get WARNING_THRESHOLD() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.WARNING_THRESHOLD : 5; },

    /** Days before deadline to show critical alert */
    get CRITICAL_THRESHOLD() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.CRITICAL_THRESHOLD : 2; }
  },

  /** Reminder intervals — references DEADLINE_DEFAULTS (01_Core.gs) as single source of truth */
  REMINDER_DAYS: {
    /** First reminder before deadline */
    get FIRST() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_FIRST : 7; },

    /** Second reminder before deadline */
    get SECOND() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_SECOND : 3; },

    /** Final reminder before deadline */
    get FINAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.REMINDER_FINAL : 1; }
  }
};

// ============================================================================
// LOCKSERVICE HELPER
// ============================================================================

/**
 * Executes a function while holding a script-level lock.
 * Prevents concurrent execution of critical mutation operations.
 * @param {Function} fn - The function to execute under lock
 * @param {number} [timeoutMs=10000] - Lock acquisition timeout in milliseconds
 * @returns {*} The return value of fn
 * @throws {Error} If lock cannot be acquired or fn throws
 */
function withScriptLock_(fn, timeoutMs) {
  var lock = LockService.getScriptLock();
  // GAS-01: Standardize on tryLock() to match 02_DataManagers.gs pattern.
  // tryLock returns false immediately if lock unavailable (non-blocking);
  // the timeout is the max wait, not a guaranteed sleep.
  if (!lock.tryLock(timeoutMs || 10000)) {
    throw new Error('Could not acquire lock. Another operation is in progress. Please try again.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SAFE SHEET ACCESSOR (v4.30.0)
// ============================================================================

/**
 * Global safe sheet accessor — logs a warning when a sheet is not found
 * instead of returning null silently. Use this for critical sheet lookups
 * where a missing sheet indicates a configuration problem.
 *
 * @param {string} name - Sheet name (use SHEETS.* constants)
 * @param {Spreadsheet} [ss] - Optional spreadsheet reference
 * @returns {Sheet|null} The sheet, or null if not found
 */
function _getSheetSafe(name, ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    log_('_getSheetSafe', 'getActiveSpreadsheet() returned null for "' + name + '"');
    return null;
  }
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    log_('_getSheetSafe', 'Sheet "' + name + '" not found');
  }
  return sheet;
}
