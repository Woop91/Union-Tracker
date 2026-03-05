// GAS_LIMITATION: All .gs files share a single global namespace. Function names must be globally unique.
/**
 * ============================================================================
 * 00_DataAccess.gs - Shared Time Constants and Lock Helper
 * ============================================================================
 *
 * Previously housed the DataAccess namespace (removed — zero callers).
 * Now provides TIME_CONSTANTS (used by deadline logic and tests) and
 * withScriptLock_ (used by DataManagers and Maintenance).
 *
 * @fileoverview Shared constants and lock helper
 * @version 1.0.0
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
    get STEP1_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_1_RESPONSE : 7; },

    /** Days to appeal to Step II */
    get STEP2_APPEAL() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_APPEAL : 7; },

    /** Days for Step II response */
    get STEP2_RESPONSE() { return (typeof DEADLINE_DEFAULTS !== 'undefined') ? DEADLINE_DEFAULTS.STEP_2_RESPONSE : 14; },

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
  try {
    lock.waitLock(timeoutMs || 10000);
  } catch (_e) {
    throw new Error('Could not acquire lock. Another operation is in progress. Please try again.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
