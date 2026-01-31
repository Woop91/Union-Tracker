/**
 * ============================================================================
 * TestFramework.gs - Unit Testing Framework
 * ============================================================================
 *
 * This module provides a lightweight testing framework for the 509 Dashboard.
 * Includes:
 * - Test runner
 * - Assertion helpers
 * - Test report generation
 *
 * NEW: Added as part of refactoring to improve code quality
 *
 * @fileoverview Unit testing framework and test cases
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// TEST FRAMEWORK
// ============================================================================

/**
 * Test result object
 * @typedef {Object} TestResult
 * @property {string} name - Test name
 * @property {boolean} passed - Whether the test passed
 * @property {string} [error] - Error message if failed
 * @property {number} duration - Test duration in ms
 */

/**
 * Test suite class
 */
var TestSuite = {
  results: [],
  currentSuite: '',

  /**
   * Runs all registered tests
   * @returns {Object} Test results summary
   */
  runAll: function() {
    this.results = [];
    var startTime = new Date().getTime();

    // Run all test functions
    var testFunctions = this.getTestFunctions_();
    testFunctions.forEach(function(testFn) {
      this.runTest_(testFn);
    }, this);

    var endTime = new Date().getTime();

    return {
      total: this.results.length,
      passed: this.results.filter(function(r) { return r.passed; }).length,
      failed: this.results.filter(function(r) { return !r.passed; }).length,
      duration: endTime - startTime,
      results: this.results
    };
  },

  /**
   * Runs a single test function
   * @param {Function} testFn - Test function to run
   * @private
   */
  runTest_: function(testFn) {
    var testName = testFn.name || 'anonymous';
    var startTime = new Date().getTime();

    try {
      testFn();
      this.results.push({
        name: testName,
        passed: true,
        duration: new Date().getTime() - startTime
      });
    } catch (e) {
      this.results.push({
        name: testName,
        passed: false,
        error: e.message,
        duration: new Date().getTime() - startTime
      });
    }
  },

  /**
   * Gets all test functions
   * @returns {Array<Function>} Array of test functions
   * @private
   */
  getTestFunctions_: function() {
    return [
      test_Constants_Defined,
      test_SheetNames_Valid,
      test_ColumnIndices_Valid,
      test_ValidationHelpers,
      test_DateHelpers,
      test_StringHelpers
    ];
  }
};

// ============================================================================
// ASSERTION HELPERS
// ============================================================================

/**
 * Assertion helper
 */
var Assert = {
  /**
   * Asserts that a value is truthy
   * @param {*} value - Value to check
   * @param {string} [message] - Error message
   */
  isTrue: function(value, message) {
    if (!value) {
      throw new Error(message || 'Expected true but got: ' + value);
    }
  },

  /**
   * Asserts that a value is falsy
   * @param {*} value - Value to check
   * @param {string} [message] - Error message
   */
  isFalse: function(value, message) {
    if (value) {
      throw new Error(message || 'Expected false but got: ' + value);
    }
  },

  /**
   * Asserts that two values are equal
   * @param {*} expected - Expected value
   * @param {*} actual - Actual value
   * @param {string} [message] - Error message
   */
  equals: function(expected, actual, message) {
    if (expected !== actual) {
      throw new Error(message || 'Expected ' + expected + ' but got: ' + actual);
    }
  },

  /**
   * Asserts that two values are not equal
   * @param {*} expected - Expected value
   * @param {*} actual - Actual value
   * @param {string} [message] - Error message
   */
  notEquals: function(expected, actual, message) {
    if (expected === actual) {
      throw new Error(message || 'Expected values to be different but both were: ' + actual);
    }
  },

  /**
   * Asserts that a value is defined (not undefined or null)
   * @param {*} value - Value to check
   * @param {string} [message] - Error message
   */
  isDefined: function(value, message) {
    if (value === undefined || value === null) {
      throw new Error(message || 'Expected value to be defined');
    }
  },

  /**
   * Asserts that a value is an array
   * @param {*} value - Value to check
   * @param {string} [message] - Error message
   */
  isArray: function(value, message) {
    if (!Array.isArray(value)) {
      throw new Error(message || 'Expected array but got: ' + typeof value);
    }
  },

  /**
   * Asserts that an array contains a value
   * @param {Array} array - Array to check
   * @param {*} value - Value to find
   * @param {string} [message] - Error message
   */
  contains: function(array, value, message) {
    if (array.indexOf(value) === -1) {
      throw new Error(message || 'Expected array to contain: ' + value);
    }
  },

  /**
   * Asserts that a function throws an error
   * @param {Function} fn - Function to execute
   * @param {string} [message] - Error message
   */
  throws: function(fn, message) {
    var threw = false;
    try {
      fn();
    } catch (e) {
      threw = true;
    }
    if (!threw) {
      throw new Error(message || 'Expected function to throw an error');
    }
  }
};

// ============================================================================
// TEST CASES - CONSTANTS
// ============================================================================

/**
 * Tests that required constants are defined
 */
function test_Constants_Defined() {
  Assert.isDefined(SHEETS, 'SHEETS constant should be defined');
  Assert.isDefined(MEMBER_COLS, 'MEMBER_COLS constant should be defined');
  Assert.isDefined(GRIEVANCE_COLS, 'GRIEVANCE_COLS constant should be defined');
  Assert.isDefined(CONFIG_COLS, 'CONFIG_COLS constant should be defined');
  Assert.isDefined(COLORS, 'COLORS constant should be defined');
}

/**
 * Tests that sheet names are valid strings
 */
function test_SheetNames_Valid() {
  Assert.isDefined(SHEETS.MEMBER_DIR, 'SHEETS.MEMBER_DIR should be defined');
  Assert.isDefined(SHEETS.GRIEVANCE_LOG, 'SHEETS.GRIEVANCE_LOG should be defined');
  Assert.isDefined(SHEETS.CONFIG, 'SHEETS.CONFIG should be defined');

  Assert.isTrue(typeof SHEETS.MEMBER_DIR === 'string', 'SHEETS.MEMBER_DIR should be a string');
  Assert.isTrue(SHEETS.MEMBER_DIR.length > 0, 'SHEETS.MEMBER_DIR should not be empty');
}

/**
 * Tests that column indices are valid numbers
 */
function test_ColumnIndices_Valid() {
  Assert.isTrue(typeof MEMBER_COLS.MEMBER_ID === 'number', 'MEMBER_COLS.MEMBER_ID should be a number');
  Assert.isTrue(MEMBER_COLS.MEMBER_ID > 0, 'MEMBER_COLS.MEMBER_ID should be positive');

  Assert.isTrue(typeof GRIEVANCE_COLS.GRIEVANCE_ID === 'number', 'GRIEVANCE_COLS.GRIEVANCE_ID should be a number');
  Assert.isTrue(GRIEVANCE_COLS.GRIEVANCE_ID > 0, 'GRIEVANCE_COLS.GRIEVANCE_ID should be positive');
}

// ============================================================================
// TEST CASES - VALIDATION
// ============================================================================

/**
 * Tests validation helper functions
 */
function test_ValidationHelpers() {
  // Test email validation
  Assert.isTrue(isValidEmail_('test@example.com'), 'Valid email should pass');
  Assert.isFalse(isValidEmail_('invalid-email'), 'Invalid email should fail');
  Assert.isFalse(isValidEmail_(''), 'Empty email should fail');

  // Test phone validation
  Assert.isTrue(isValidPhone_('555-123-4567'), 'Valid phone should pass');
  Assert.isTrue(isValidPhone_('5551234567'), 'Phone without dashes should pass');
}

/**
 * Tests date helper functions
 */
function test_DateHelpers() {
  var today = new Date();

  // Test date formatting
  var formatted = formatDate_(today);
  Assert.isDefined(formatted, 'Formatted date should be defined');
  Assert.isTrue(typeof formatted === 'string', 'Formatted date should be a string');
}

/**
 * Tests string helper functions
 */
function test_StringHelpers() {
  // Test string trimming
  Assert.equals('test', trimString_('  test  '), 'String should be trimmed');
  Assert.equals('', trimString_(''), 'Empty string should remain empty');
}

// ============================================================================
// TEST HELPER FUNCTIONS (for tests only)
// ============================================================================

/**
 * Validates email format
 * @param {string} email - Email to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidEmail_(email) {
  if (!email || typeof email !== 'string') return false;
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates phone format
 * @param {string} phone - Phone to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidPhone_(phone) {
  if (!phone || typeof phone !== 'string') return false;
  var digits = phone.replace(/\D/g, '');
  return digits.length >= 10 && digits.length <= 11;
}

/**
 * Formats a date
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string
 * @private
 */
function formatDate_(date) {
  if (!date || !(date instanceof Date)) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Trims a string
 * @param {string} str - String to trim
 * @returns {string} Trimmed string
 * @private
 */
function trimString_(str) {
  if (!str || typeof str !== 'string') return '';
  return str.trim();
}

// ============================================================================
// TEST RUNNER MENU FUNCTIONS
// ============================================================================

/**
 * Runs all tests and displays results
 * @returns {void}
 */
function runAllTests() {
  var results = TestSuite.runAll();

  var ui = SpreadsheetApp.getUi();
  var message = 'Test Results:\n\n' +
    '✅ Passed: ' + results.passed + '\n' +
    '❌ Failed: ' + results.failed + '\n' +
    '⏱️ Duration: ' + results.duration + 'ms\n\n';

  if (results.failed > 0) {
    message += 'Failed Tests:\n';
    results.results.filter(function(r) { return !r.passed; }).forEach(function(r) {
      message += '• ' + r.name + ': ' + r.error + '\n';
    });
  }

  ui.alert('🧪 Test Results', message, ui.ButtonSet.OK);
}

/**
 * Runs quick smoke tests
 * @returns {void}
 */
function runQuickTests() {
  try {
    test_Constants_Defined();
    test_SheetNames_Valid();
    test_ColumnIndices_Valid();
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ All quick tests passed!', 'Tests', 3);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ Test failed: ' + e.message, 'Tests', 5);
  }
}

/**
 * Generates a test report
 * @param {number} duration - Test duration
 * @returns {void}
 */
function generateTestReport(duration) {
  var results = TestSuite.results;
  var passed = results.filter(function(r) { return r.passed; }).length;
  var failed = results.filter(function(r) { return !r.passed; }).length;

  Logger.log('=== TEST REPORT ===');
  Logger.log('Total: ' + results.length);
  Logger.log('Passed: ' + passed);
  Logger.log('Failed: ' + failed);
  Logger.log('Duration: ' + duration + 'ms');
  Logger.log('==================');
}
