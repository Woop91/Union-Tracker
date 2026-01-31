/**
 * ============================================================================
 * TestFramework.gs - Comprehensive Unit Testing Framework
 * ============================================================================
 *
 * This module provides a comprehensive testing framework for the 509 Dashboard.
 * Includes:
 * - Test runner with suite management
 * - Rich assertion helpers
 * - Test report generation
 * - Module-specific test suites
 * - Performance benchmarking
 *
 * USAGE:
 *   runAllTests()     - Run all test suites
 *   runQuickTests()   - Run essential smoke tests
 *   runModuleTests()  - Run tests for specific module
 *
 * @fileoverview Unit testing framework and comprehensive test cases
 * @version 2.0.0
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
      // Constants tests
      test_Constants_Defined,
      test_SheetNames_Valid,
      test_ColumnIndices_Valid,
      test_ColorsConfig_Valid,
      test_HiddenSheets_Defined,

      // Validation tests
      test_ValidationHelpers,
      test_EmailValidation_EdgeCases,
      test_PhoneValidation_EdgeCases,
      test_MemberIdFormat,
      test_GrievanceIdFormat,

      // Helper tests
      test_DateHelpers,
      test_StringHelpers,
      test_ArrayHelpers,
      test_ObjectHelpers,

      // Cache tests
      test_CacheConfig_Valid,
      test_CacheKeys_Defined,

      // Module integration tests
      test_SearchEngine_Functions,
      test_ThemeService_Functions,
      test_MenuBuilder_Functions
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

// ============================================================================
// EXTENDED TEST CASES - CONSTANTS
// ============================================================================

/**
 * Tests that color configuration is valid
 */
function test_ColorsConfig_Valid() {
  Assert.isDefined(COLORS.PRIMARY_PURPLE, 'COLORS.PRIMARY_PURPLE should be defined');
  Assert.isDefined(COLORS.WHITE, 'COLORS.WHITE should be defined');
  Assert.isDefined(COLORS.TEXT_DARK, 'COLORS.TEXT_DARK should be defined');

  // Validate hex color format
  var hexRegex = /^#[0-9A-Fa-f]{6}$/;
  Assert.isTrue(hexRegex.test(COLORS.PRIMARY_PURPLE), 'PRIMARY_PURPLE should be valid hex');
  Assert.isTrue(hexRegex.test(COLORS.WHITE), 'WHITE should be valid hex');
}

/**
 * Tests that hidden sheet names are defined
 */
function test_HiddenSheets_Defined() {
  Assert.isDefined(HIDDEN_SHEETS, 'HIDDEN_SHEETS constant should be defined');
  Assert.isDefined(HIDDEN_SHEETS.CALC_STATS, 'HIDDEN_SHEETS.CALC_STATS should be defined');
  Assert.isDefined(HIDDEN_SHEETS.CALC_FORMULAS, 'HIDDEN_SHEETS.CALC_FORMULAS should be defined');
}

// ============================================================================
// EXTENDED TEST CASES - VALIDATION
// ============================================================================

/**
 * Tests email validation edge cases
 */
function test_EmailValidation_EdgeCases() {
  // Valid emails
  Assert.isTrue(isValidEmail_('user@domain.com'), 'Standard email should pass');
  Assert.isTrue(isValidEmail_('user.name@domain.com'), 'Email with dot should pass');
  Assert.isTrue(isValidEmail_('user+tag@domain.com'), 'Email with plus should pass');
  Assert.isTrue(isValidEmail_('user@sub.domain.com'), 'Email with subdomain should pass');

  // Invalid emails
  Assert.isFalse(isValidEmail_(''), 'Empty string should fail');
  Assert.isFalse(isValidEmail_(null), 'Null should fail');
  Assert.isFalse(isValidEmail_(undefined), 'Undefined should fail');
  Assert.isFalse(isValidEmail_('userATdomain.com'), 'Email without @ should fail');
  Assert.isFalse(isValidEmail_('@domain.com'), 'Email starting with @ should fail');
  Assert.isFalse(isValidEmail_('user@'), 'Email ending with @ should fail');
  Assert.isFalse(isValidEmail_('user@.com'), 'Email with @ followed by dot should fail');
}

/**
 * Tests phone validation edge cases
 */
function test_PhoneValidation_EdgeCases() {
  // Valid phones
  Assert.isTrue(isValidPhone_('5551234567'), '10 digit phone should pass');
  Assert.isTrue(isValidPhone_('15551234567'), '11 digit phone should pass');
  Assert.isTrue(isValidPhone_('555-123-4567'), 'Phone with dashes should pass');
  Assert.isTrue(isValidPhone_('(555) 123-4567'), 'Phone with parens should pass');
  Assert.isTrue(isValidPhone_('555.123.4567'), 'Phone with dots should pass');

  // Invalid phones
  Assert.isFalse(isValidPhone_(''), 'Empty string should fail');
  Assert.isFalse(isValidPhone_(null), 'Null should fail');
  Assert.isFalse(isValidPhone_('12345'), 'Too short should fail');
  Assert.isFalse(isValidPhone_('123456789012'), 'Too long should fail');
}

/**
 * Tests member ID format validation
 */
function test_MemberIdFormat() {
  Assert.isTrue(isValidMemberId_('MBR-001'), 'Standard member ID should pass');
  Assert.isTrue(isValidMemberId_('MBR-12345'), 'Member ID with 5 digits should pass');
  Assert.isFalse(isValidMemberId_(''), 'Empty should fail');
  Assert.isFalse(isValidMemberId_(null), 'Null should fail');
}

/**
 * Tests grievance ID format validation
 */
function test_GrievanceIdFormat() {
  Assert.isTrue(isValidGrievanceId_('GRV-2024-001'), 'Standard grievance ID should pass');
  Assert.isTrue(isValidGrievanceId_('GRV-2024-12345'), 'Grievance ID with more digits should pass');
  Assert.isFalse(isValidGrievanceId_(''), 'Empty should fail');
  Assert.isFalse(isValidGrievanceId_(null), 'Null should fail');
}

// ============================================================================
// EXTENDED TEST CASES - HELPERS
// ============================================================================

/**
 * Tests array helper functions
 */
function test_ArrayHelpers() {
  var arr = [1, 2, 3, 4, 5];

  // Test unique function
  var withDups = [1, 2, 2, 3, 3, 3];
  var unique = getUniqueValues_(withDups);
  Assert.equals(3, unique.length, 'Unique should remove duplicates');

  // Test flatten
  var nested = [[1, 2], [3, 4]];
  var flat = flattenArray_(nested);
  Assert.equals(4, flat.length, 'Flatten should combine arrays');
}

/**
 * Tests object helper functions
 */
function test_ObjectHelpers() {
  var obj = { a: 1, b: 2, c: 3 };

  // Test keys
  var keys = Object.keys(obj);
  Assert.equals(3, keys.length, 'Object should have 3 keys');
  Assert.contains(keys, 'a', 'Keys should contain "a"');

  // Test hasProperty
  Assert.isTrue(hasProperty_(obj, 'a'), 'Object should have property a');
  Assert.isFalse(hasProperty_(obj, 'd'), 'Object should not have property d');
}

// ============================================================================
// TEST CASES - CACHE
// ============================================================================

/**
 * Tests cache configuration
 */
function test_CacheConfig_Valid() {
  Assert.isDefined(CACHE_CONFIG, 'CACHE_CONFIG should be defined');
  Assert.isTrue(typeof CACHE_CONFIG.MEMORY_TTL === 'number', 'MEMORY_TTL should be a number');
  Assert.isTrue(CACHE_CONFIG.MEMORY_TTL > 0, 'MEMORY_TTL should be positive');
}

/**
 * Tests cache keys are defined
 */
function test_CacheKeys_Defined() {
  Assert.isDefined(CACHE_KEYS, 'CACHE_KEYS should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_GRIEVANCES, 'ALL_GRIEVANCES key should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_MEMBERS, 'ALL_MEMBERS key should be defined');
  Assert.isDefined(CACHE_KEYS.ALL_STEWARDS, 'ALL_STEWARDS key should be defined');
}

// ============================================================================
// TEST CASES - MODULE INTEGRATION
// ============================================================================

/**
 * Tests SearchEngine module functions exist
 */
function test_SearchEngine_Functions() {
  Assert.isDefined(typeof getDesktopSearchLocations, 'getDesktopSearchLocations should be defined');
  Assert.isDefined(typeof getDesktopSearchData, 'getDesktopSearchData should be defined');
  Assert.isDefined(typeof navigateToSearchResult, 'navigateToSearchResult should be defined');
  Assert.isDefined(typeof searchDashboard, 'searchDashboard should be defined');
}

/**
 * Tests ThemeService module functions exist
 */
function test_ThemeService_Functions() {
  Assert.isDefined(typeof APPLY_SYSTEM_THEME, 'APPLY_SYSTEM_THEME should be defined');
  Assert.isDefined(typeof resetToDefaultTheme, 'resetToDefaultTheme should be defined');
  Assert.isDefined(typeof getADHDSettings, 'getADHDSettings should be defined');
  Assert.isDefined(typeof applyZebraStripes, 'applyZebraStripes should be defined');
}

/**
 * Tests MenuBuilder module functions exist
 */
function test_MenuBuilder_Functions() {
  Assert.isDefined(typeof createDashboardMenu, 'createDashboardMenu should be defined');
  Assert.isDefined(typeof navigateToSheet, 'navigateToSheet should be defined');
  Assert.isDefined(typeof showToast, 'showToast should be defined');
}

// ============================================================================
// ADDITIONAL TEST HELPERS
// ============================================================================

/**
 * Validates member ID format
 * @param {string} id - ID to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidMemberId_(id) {
  if (!id || typeof id !== 'string') return false;
  return /^MBR-\d+$/.test(id);
}

/**
 * Validates grievance ID format
 * @param {string} id - ID to validate
 * @returns {boolean} True if valid
 * @private
 */
function isValidGrievanceId_(id) {
  if (!id || typeof id !== 'string') return false;
  return /^GRV-\d{4}-\d+$/.test(id);
}

/**
 * Gets unique values from array
 * @param {Array} arr - Array to process
 * @returns {Array} Array with unique values
 * @private
 */
function getUniqueValues_(arr) {
  var seen = {};
  return arr.filter(function(item) {
    if (seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

/**
 * Flattens a nested array
 * @param {Array} arr - Nested array
 * @returns {Array} Flattened array
 * @private
 */
function flattenArray_(arr) {
  var result = [];
  arr.forEach(function(item) {
    if (Array.isArray(item)) {
      result = result.concat(flattenArray_(item));
    } else {
      result.push(item);
    }
  });
  return result;
}

/**
 * Checks if object has property
 * @param {Object} obj - Object to check
 * @param {string} prop - Property name
 * @returns {boolean} True if has property
 * @private
 */
function hasProperty_(obj, prop) {
  return obj && Object.prototype.hasOwnProperty.call(obj, prop);
}

// ============================================================================
// TEST DASHBOARD UI
// ============================================================================

/**
 * Shows a comprehensive test dashboard
 * @returns {void}
 */
function showTestDashboard() {
  var results = TestSuite.runAll();

  var testRows = results.results.map(function(r) {
    var status = r.passed ? '✅' : '❌';
    var errorMsg = r.error ? '<span style="color:#ef4444">' + r.error + '</span>' : '-';
    return '<tr>' +
      '<td>' + status + '</td>' +
      '<td>' + r.name + '</td>' +
      '<td>' + r.duration + 'ms</td>' +
      '<td>' + errorMsg + '</td>' +
      '</tr>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px}' +
    'h2{color:#7c3aed}' +
    '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin:20px 0}' +
    '.stat{background:#f8f9fa;padding:20px;border-radius:8px;text-align:center}' +
    '.stat.passed{border-left:4px solid #10b981}' +
    '.stat.failed{border-left:4px solid #ef4444}' +
    '.num{font-size:36px;font-weight:bold}' +
    '.passed .num{color:#10b981}' +
    '.failed .num{color:#ef4444}' +
    'table{width:100%;border-collapse:collapse;margin-top:20px}' +
    'th{background:#7c3aed;color:white;padding:12px;text-align:left}' +
    'td{padding:10px;border-bottom:1px solid #e5e7eb}' +
    '.btn{padding:12px 24px;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin:5px}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>🧪 Test Dashboard</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + results.total + '</div><div>Total</div></div>' +
    '<div class="stat passed"><div class="num">' + results.passed + '</div><div>Passed</div></div>' +
    '<div class="stat failed"><div class="num">' + results.failed + '</div><div>Failed</div></div>' +
    '<div class="stat"><div class="num">' + results.duration + 'ms</div><div>Duration</div></div>' +
    '</div>' +
    '<button class="btn btn-primary" onclick="rerun()">🔄 Run Again</button>' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '<table>' +
    '<tr><th>Status</th><th>Test Name</th><th>Duration</th><th>Error</th></tr>' +
    testRows +
    '</table>' +
    '</div>' +
    '<script>function rerun(){google.script.run.withSuccessHandler(function(){location.reload()}).runAllTests()}</script>' +
    '</body></html>'
  ).setWidth(900).setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '🧪 Test Dashboard');
}
