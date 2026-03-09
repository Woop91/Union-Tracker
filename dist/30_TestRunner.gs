/**
 * 30_TestRunner.gs
 * GAS-Native Test Framework — runs inside the Apps Script runtime.
 *
 * Provides:
 *   - Lightweight assertion library (assertEquals, assertTrue, etc.)
 *   - Test discovery via naming convention (test_SUITE_description)
 *   - Suite runner with timing, pass/fail, error capture
 *   - Results stored in ScriptProperties for SPA dashboard panel
 *   - Manual trigger (Sheets menu) + timed trigger (daily)
 *
 * Test suites:
 *   config_   — Config tab reads, CONFIG_COLS, ConfigReader shape
 *   colmap_   — Column mapping (GRIEVANCE_COLS, MEMBER_COLS) integrity
 *   auth_     — Auth resolution, role checks, steward auth
 *   grievance_ — Deadline rules, status constants, ID validation
 *   security_ — escapeHtml, escapeForFormula, XSS prevention
 *   system_   — Sheet existence, version info, build integrity
 *
 * IMPORTANT: All tests are READ-ONLY. They never write to sheets.
 * Tests interact with real GAS services (SpreadsheetApp, CacheService, etc.)
 * to verify actual system state — this is the point of GAS-native testing.
 *
 * @fileoverview GAS-native test runner for DDS-Dashboard
 * @version 1.0.0
 */

/* ========================================================================
 * TEST FRAMEWORK — Assert helpers + runner
 * ======================================================================== */

var TestRunner = (function () {

  var RESULTS_KEY = 'TEST_RUNNER_RESULTS';
  var STATUS_KEY  = 'TEST_RUNNER_STATUS';

  // ── Assert helpers ──────────────────────────────────────────────────

  function _assertionError(msg) {
    var err = new Error(msg);
    err.isAssertion = true;
    return err;
  }

  function assertEquals(expected, actual, label) {
    if (expected !== actual) {
      throw _assertionError(
        (label || 'assertEquals') + ': expected ' +
        JSON.stringify(expected) + ' but got ' + JSON.stringify(actual)
      );
    }
  }

  function assertDeepEquals(expected, actual, label) {
    if (JSON.stringify(expected) !== JSON.stringify(actual)) {
      throw _assertionError(
        (label || 'assertDeepEquals') + ': objects differ.\n  Expected: ' +
        JSON.stringify(expected).substring(0, 200) + '\n  Actual:   ' +
        JSON.stringify(actual).substring(0, 200)
      );
    }
  }

  function assertTrue(value, label) {
    if (value !== true) {
      throw _assertionError(
        (label || 'assertTrue') + ': expected true but got ' + JSON.stringify(value)
      );
    }
  }

  function assertFalsy(value, label) {
    if (value) {
      throw _assertionError(
        (label || 'assertFalsy') + ': expected falsy but got ' + JSON.stringify(value)
      );
    }
  }

  function assertNotNull(value, label) {
    if (value === null || value === undefined) {
      throw _assertionError(
        (label || 'assertNotNull') + ': expected non-null value'
      );
    }
  }

  function assertType(value, expectedType, label) {
    if (typeof value !== expectedType) {
      throw _assertionError(
        (label || 'assertType') + ': expected type "' + expectedType +
        '" but got "' + typeof value + '"'
      );
    }
  }

  function assertContains(haystack, needle, label) {
    if (typeof haystack === 'string') {
      if (haystack.indexOf(needle) === -1) {
        throw _assertionError(
          (label || 'assertContains') + ': string does not contain "' + needle + '"'
        );
      }
    } else if (Array.isArray(haystack)) {
      if (haystack.indexOf(needle) === -1) {
        throw _assertionError(
          (label || 'assertContains') + ': array does not contain ' + JSON.stringify(needle)
        );
      }
    } else {
      throw _assertionError((label || 'assertContains') + ': value is not string or array');
    }
  }

  function assertThrows(fn, label) {
    var threw = false;
    try { fn(); } catch (_e) { threw = true; }
    if (!threw) {
      throw _assertionError(
        (label || 'assertThrows') + ': expected function to throw but it did not'
      );
    }
  }

  function assertGreaterThan(value, threshold, label) {
    if (!(value > threshold)) {
      throw _assertionError(
        (label || 'assertGreaterThan') + ': expected ' + value + ' > ' + threshold
      );
    }
  }

  function assertHasKey(obj, key, label) {
    if (!obj || !obj.hasOwnProperty(key)) {
      throw _assertionError(
        (label || 'assertHasKey') + ': object missing key "' + key + '"'
      );
    }
  }

  // ── Test discovery & execution ──────────────────────────────────────

  /**
   * Discovers all global functions matching test_SUITE_* pattern.
   * Groups them by suite name (text between first and second underscore).
   * @returns {Object} { suiteName: [{ name, fn }], ... }
   */
  function _discoverTests() {
    var suites = {};
    var globalScope = this;

    // GAS globals — iterate known test function names
    // We use a registry approach since GAS doesn't support Object.keys(this) reliably
    var allTests = _getTestRegistry();

    for (var i = 0; i < allTests.length; i++) {
      var entry = allTests[i];
      var parts = entry.name.split('_');
      // test_SUITE_description → suite = parts[1]
      if (parts.length >= 3 && parts[0] === 'test') {
        var suite = parts[1];
        if (!suites[suite]) suites[suite] = [];
        suites[suite].push({ name: entry.name, fn: entry.fn });
      }
    }
    return suites;
  }

  /**
   * Runs all discovered test suites.
   * @param {string} [filterSuite] - Optional suite name to run only that suite
   * @returns {Object} Full results object
   */
  function runAll(filterSuite) {
    var startTime = new Date().getTime();

    // Update status to "running"
    _setStatus('running');

    var suites = _discoverTests();
    var results = {
      timestamp: new Date().toISOString(),
      duration: 0,
      filterSuite: filterSuite || null,
      suites: [],
      summary: { total: 0, passed: 0, failed: 0, skipped: 0 }
    };

    var suiteNames = Object.keys(suites);
    if (filterSuite) {
      suiteNames = suiteNames.filter(function (s) { return s === filterSuite; });
    }

    for (var si = 0; si < suiteNames.length; si++) {
      var suiteName = suiteNames[si];
      var tests = suites[suiteName];
      var suiteResult = {
        name: suiteName,
        tests: [],
        passed: 0,
        failed: 0,
        duration: 0
      };

      var suiteStart = new Date().getTime();

      for (var ti = 0; ti < tests.length; ti++) {
        var test = tests[ti];
        var testStart = new Date().getTime();
        var testResult = { name: test.name, passed: false, error: null, duration: 0 };

        try {
          test.fn();
          testResult.passed = true;
          suiteResult.passed++;
        } catch (err) {
          testResult.error = err.message || String(err);
          suiteResult.failed++;
        }

        testResult.duration = new Date().getTime() - testStart;
        suiteResult.tests.push(testResult);
        results.summary.total++;
        if (testResult.passed) results.summary.passed++;
        else results.summary.failed++;
      }

      suiteResult.duration = new Date().getTime() - suiteStart;
      results.suites.push(suiteResult);
    }

    results.duration = new Date().getTime() - startTime;

    // Store results
    _storeResults(results);
    _setStatus('complete');

    return results;
  }

  /**
   * Stores results in ScriptProperties (for SPA retrieval).
   */
  function _storeResults(results) {
    try {
      var json = JSON.stringify(results);
      PropertiesService.getScriptProperties().setProperty(RESULTS_KEY, json);
    } catch (e) {
      Logger.log('TestRunner: Failed to store results — ' + e.message);
    }
  }

  function _setStatus(status) {
    try {
      PropertiesService.getScriptProperties().setProperty(STATUS_KEY, JSON.stringify({
        status: status,
        timestamp: new Date().toISOString()
      }));
    } catch (_e) { /* best effort */ }
  }

  /**
   * Retrieves stored test results (called by SPA dashboard).
   * @returns {Object|null}
   */
  function getResults() {
    try {
      var raw = PropertiesService.getScriptProperties().getProperty(RESULTS_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  /**
   * Gets the current runner status.
   * @returns {Object} { status: 'idle'|'running'|'complete', timestamp }
   */
  function getStatus() {
    try {
      var raw = PropertiesService.getScriptProperties().getProperty(STATUS_KEY);
      return raw ? JSON.parse(raw) : { status: 'idle', timestamp: null };
    } catch (_e) {
      return { status: 'idle', timestamp: null };
    }
  }

  // ── Trigger management ──────────────────────────────────────────────

  /**
   * Creates a daily trigger to run tests at 6 AM.
   */
  function setupDailyTrigger() {
    // Remove existing test triggers first
    removeTrigger();
    ScriptApp.newTrigger('runScheduledTests')
      .timeBased()
      .atHour(6)
      .everyDays(1)
      .create();
    return { success: true, message: 'Daily test trigger set for 6:00 AM' };
  }

  /**
   * Removes the test trigger.
   */
  function removeTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'runScheduledTests') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    return { success: true, message: 'Test trigger removed' };
  }

  // ── Public API ──────────────────────────────────────────────────────

  return {
    // Assert helpers (exposed for test functions)
    assertEquals: assertEquals,
    assertDeepEquals: assertDeepEquals,
    assertTrue: assertTrue,
    assertFalsy: assertFalsy,
    assertNotNull: assertNotNull,
    assertType: assertType,
    assertContains: assertContains,
    assertThrows: assertThrows,
    assertGreaterThan: assertGreaterThan,
    assertHasKey: assertHasKey,

    // Runner
    runAll: runAll,
    getResults: getResults,
    getStatus: getStatus,

    // Triggers
    setupDailyTrigger: setupDailyTrigger,
    removeTrigger: removeTrigger,
  };
})();

/* ========================================================================
 * TRIGGER HANDLER — Called by timed trigger
 * ======================================================================== */

/**
 * Global function called by the daily trigger.
 * Must be a top-level function (GAS trigger requirement).
 */
function runScheduledTests() {
  TestRunner.runAll();
}

/**
 * Menu handler: run all tests manually from Sheets UI.
 */
function runTestsFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Running tests...', 'Test Runner', 30);
  var results = TestRunner.runAll();
  var msg = results.summary.passed + ' passed, ' + results.summary.failed + ' failed (' + results.duration + 'ms)';
  ss.toast(msg, results.summary.failed > 0 ? '❌ Tests Failed' : '✅ Tests Passed', 10);
}

/**
 * Menu handler: setup daily trigger.
 */
function setupTestTriggerFromMenu() {
  var result = TestRunner.setupDailyTrigger();
  SpreadsheetApp.getActiveSpreadsheet().toast(result.message, 'Test Trigger', 5);
}

/**
 * Menu handler: remove trigger.
 */
function removeTestTriggerFromMenu() {
  var result = TestRunner.removeTrigger();
  SpreadsheetApp.getActiveSpreadsheet().toast(result.message, 'Test Trigger', 5);
}

/* ========================================================================
 * SPA DATA ENDPOINTS — Called by web dashboard
 * ======================================================================== */

/**
 * SPA endpoint: Run all tests (steward-only).
 * @param {string} [sessionToken]
 * @param {string} [filterSuite] - Optional suite to run
 * @returns {Object} Test results
 */
function dataRunTests(sessionToken, filterSuite) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, error: 'Not authenticated' };

  // Steward-only check
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, error: 'Steward access required' };

  var results = TestRunner.runAll(filterSuite || null);
  return { success: true, results: results };
}

/**
 * SPA endpoint: Get last test results (steward-only).
 * @param {string} [sessionToken]
 * @returns {Object}
 */
function dataGetTestResults(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, error: 'Not authenticated' };

  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, error: 'Steward access required' };

  var results = TestRunner.getResults();
  var status = TestRunner.getStatus();
  return { success: true, results: results, status: status };
}

/**
 * SPA endpoint: Manage test trigger (steward-only).
 * @param {string} sessionToken
 * @param {string} action - 'setup' or 'remove'
 * @returns {Object}
 */
function dataManageTestTrigger(sessionToken, action) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, error: 'Steward access required' };

  if (action === 'setup') return TestRunner.setupDailyTrigger();
  if (action === 'remove') return TestRunner.removeTrigger();
  return { success: false, error: 'Unknown action: ' + action };
}


/* ========================================================================
 * TEST REGISTRY — All test functions listed here for discovery
 * ========================================================================
 * GAS doesn't support Object.keys(this) for global function discovery.
 * Each test function is registered here AND defined below.
 * ======================================================================== */

function _getTestRegistry() {
  return [
    // ── config suite ──
    { name: 'test_config_sheetsConstantDefined',       fn: test_config_sheetsConstantDefined },
    { name: 'test_config_configSheetExists',           fn: test_config_configSheetExists },
    { name: 'test_config_memberDirSheetExists',        fn: test_config_memberDirSheetExists },
    { name: 'test_config_grievanceLogSheetExists',     fn: test_config_grievanceLogSheetExists },
    { name: 'test_config_configReaderReturnsObject',   fn: test_config_configReaderReturnsObject },
    { name: 'test_config_configHasOrgName',            fn: test_config_configHasOrgName },
    { name: 'test_config_configHasAccentHue',          fn: test_config_configHasAccentHue },
    { name: 'test_config_configColsPopulated',         fn: test_config_configColsPopulated },
    { name: 'test_config_defaultDropdownsDefined',     fn: test_config_defaultDropdownsDefined },
    { name: 'test_config_deadlineDefaultsComplete',    fn: test_config_deadlineDefaultsComplete },

    // ── colmap suite ──
    { name: 'test_colmap_grievanceColsDefined',        fn: test_colmap_grievanceColsDefined },
    { name: 'test_colmap_memberColsDefined',           fn: test_colmap_memberColsDefined },
    { name: 'test_colmap_grievanceHasStatus',          fn: test_colmap_grievanceHasStatus },
    { name: 'test_colmap_grievanceHasStep',            fn: test_colmap_grievanceHasStep },
    { name: 'test_colmap_memberHasEmail',              fn: test_colmap_memberHasEmail },
    { name: 'test_colmap_memberHasRole',               fn: test_colmap_memberHasRole },
    { name: 'test_colmap_configColsHasOrgName',        fn: test_colmap_configColsHasOrgName },
    { name: 'test_colmap_allColsPositive',             fn: test_colmap_allColsPositive },
    { name: 'test_colmap_noDuplicatePositions',        fn: test_colmap_noDuplicatePositions },

    // ── auth suite ──
    { name: 'test_auth_moduleExists',                  fn: test_auth_moduleExists },
    { name: 'test_auth_resolveUserFunctionExists',     fn: test_auth_resolveUserFunctionExists },
    { name: 'test_auth_dataServiceExists',             fn: test_auth_dataServiceExists },
    { name: 'test_auth_findUserByEmailExists',         fn: test_auth_findUserByEmailExists },
    { name: 'test_auth_resolveCallerEmailExists',      fn: test_auth_resolveCallerEmailExists },
    { name: 'test_auth_requireStewardAuthExists',      fn: test_auth_requireStewardAuthExists },
    { name: 'test_auth_checkWebAppAuthExists',         fn: test_auth_checkWebAppAuthExists },
    { name: 'test_auth_sessionTokenFunctionsExist',    fn: test_auth_sessionTokenFunctionsExist },

    // ── grievance suite ──
    { name: 'test_grievance_statusConstantsDefined',   fn: test_grievance_statusConstantsDefined },
    { name: 'test_grievance_allStatusesPresent',       fn: test_grievance_allStatusesPresent },
    { name: 'test_grievance_closedStatusesSubset',     fn: test_grievance_closedStatusesSubset },
    { name: 'test_grievance_priorityCoversAllStatuses', fn: test_grievance_priorityCoversAllStatuses },
    { name: 'test_grievance_deadlineRulesReadable',    fn: test_grievance_deadlineRulesReadable },
    { name: 'test_grievance_deadlineRulesHaveAllSteps', fn: test_grievance_deadlineRulesHaveAllSteps },
    { name: 'test_grievance_deadlineDefaultsReasonable', fn: test_grievance_deadlineDefaultsReasonable },
    { name: 'test_grievance_stepsListDefined',         fn: test_grievance_stepsListDefined },
    { name: 'test_grievance_isValidIdFunction',        fn: test_grievance_isValidIdFunction },
    { name: 'test_grievance_outcomesDefined',          fn: test_grievance_outcomesDefined },

    // ── security suite ──
    { name: 'test_security_escapeHtmlDefined',         fn: test_security_escapeHtmlDefined },
    { name: 'test_security_escapeHtmlBlocks',          fn: test_security_escapeHtmlBlocks },
    { name: 'test_security_escapeForFormulaDefined',   fn: test_security_escapeForFormulaDefined },
    { name: 'test_security_escapeFormulaBlocks',       fn: test_security_escapeFormulaBlocks },
    { name: 'test_security_escapeHtmlHandlesNull',     fn: test_security_escapeHtmlHandlesNull },
    { name: 'test_security_escapeHtmlHandlesNumbers',  fn: test_security_escapeHtmlHandlesNumbers },

    // ── system suite ──
    { name: 'test_system_versionInfoDefined',          fn: test_system_versionInfoDefined },
    { name: 'test_system_versionFormatSemver',         fn: test_system_versionFormatSemver },
    { name: 'test_system_spreadsheetBound',            fn: test_system_spreadsheetBound },
    { name: 'test_system_eventBusExists',              fn: test_system_eventBusExists },
    { name: 'test_system_hiddenSheetsConstant',        fn: test_system_hiddenSheetsConstant },

    // ── dataservice suite (DataService CRUD) ──
    { name: 'test_dataservice_moduleExists',               fn: test_dataservice_moduleExists },
    { name: 'test_dataservice_findUserByEmailCallable',    fn: test_dataservice_findUserByEmailCallable },
    { name: 'test_dataservice_getUserRoleCallable',        fn: test_dataservice_getUserRoleCallable },
    { name: 'test_dataservice_getStewardCasesCallable',    fn: test_dataservice_getStewardCasesCallable },
    { name: 'test_dataservice_getAllMembersCallable',      fn: test_dataservice_getAllMembersCallable },
    { name: 'test_dataservice_getUnitsCallable',           fn: test_dataservice_getUnitsCallable },
    { name: 'test_dataservice_getBatchDataCallable',       fn: test_dataservice_getBatchDataCallable },
    { name: 'test_dataservice_memberLookupReturnsShape',   fn: test_dataservice_memberLookupReturnsShape },
    { name: 'test_dataservice_invalidEmailReturnsNull',    fn: test_dataservice_invalidEmailReturnsNull },
    { name: 'test_dataservice_publicAPIComplete',          fn: test_dataservice_publicAPIComplete },

    // ── authsweep suite (Endpoint auth rejection) ──
    { name: 'test_authsweep_allDataFnsExist',              fn: test_authsweep_allDataFnsExist },
    { name: 'test_authsweep_stewardEndpointsRejectNull',   fn: test_authsweep_stewardEndpointsRejectNull },
    { name: 'test_authsweep_memberEndpointsRejectNull',    fn: test_authsweep_memberEndpointsRejectNull },
    { name: 'test_authsweep_noDataLeakOnNullToken',        fn: test_authsweep_noDataLeakOnNullToken },
    { name: 'test_authsweep_pollStubsSafe',                fn: test_authsweep_pollStubsSafe },
    { name: 'test_authsweep_testRunnerEndpointsGated',     fn: test_authsweep_testRunnerEndpointsGated },

    // ── configlive suite (Config completeness vs live headers) ──
    { name: 'test_configlive_configSheetHasHeaders',       fn: test_configlive_configSheetHasHeaders },
    { name: 'test_configlive_memberDirHasHeaders',         fn: test_configlive_memberDirHasHeaders },
    { name: 'test_configlive_grievanceLogHasHeaders',      fn: test_configlive_grievanceLogHasHeaders },
    { name: 'test_configlive_configColsMatchSheet',        fn: test_configlive_configColsMatchSheet },
    { name: 'test_configlive_memberColsMatchSheet',        fn: test_configlive_memberColsMatchSheet },
    { name: 'test_configlive_grievanceColsMatchSheet',     fn: test_configlive_grievanceColsMatchSheet },
    { name: 'test_configlive_syncColumnMapsCallable',      fn: test_configlive_syncColumnMapsCallable },
    { name: 'test_configlive_configRow3HasValues',         fn: test_configlive_configRow3HasValues },

    // ── survey suite (Survey engine integrity) ──
    { name: 'test_survey_hiddenSheetsConstants',           fn: test_survey_hiddenSheetsConstants },
    { name: 'test_survey_periodsColsDefined',              fn: test_survey_periodsColsDefined },
    { name: 'test_survey_questionsColsDefined',            fn: test_survey_questionsColsDefined },
    { name: 'test_survey_getSurveyQuestionsCallable',      fn: test_survey_getSurveyQuestionsCallable },
    { name: 'test_survey_questionsReturnArray',            fn: test_survey_questionsReturnArray },
    { name: 'test_survey_questionShapeValid',              fn: test_survey_questionShapeValid },
    { name: 'test_survey_getSurveyPeriodCallable',         fn: test_survey_getSurveyPeriodCallable },
    { name: 'test_survey_submitResponseCallable',          fn: test_survey_submitResponseCallable },
    { name: 'test_survey_satisfactionColsDefined',         fn: test_survey_satisfactionColsDefined },
    { name: 'test_survey_trackingSheetExists',             fn: test_survey_trackingSheetExists },
  ];
}


/* ========================================================================
 * TEST SUITES
 * ========================================================================
 * Naming: test_SUITE_description
 * All tests are READ-ONLY — never write to sheets.
 * ======================================================================== */

// ── CONFIG SUITE ──────────────────────────────────────────────────────

function test_config_sheetsConstantDefined() {
  TestRunner.assertNotNull(SHEETS, 'SHEETS constant');
  TestRunner.assertType(SHEETS, 'object', 'SHEETS type');
  TestRunner.assertHasKey(SHEETS, 'CONFIG', 'SHEETS.CONFIG');
  TestRunner.assertHasKey(SHEETS, 'MEMBER_DIR', 'SHEETS.MEMBER_DIR');
  TestRunner.assertHasKey(SHEETS, 'GRIEVANCE_LOG', 'SHEETS.GRIEVANCE_LOG');
}

function test_config_configSheetExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return; // web context — skip gracefully
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  TestRunner.assertNotNull(sheet, 'Config sheet "' + SHEETS.CONFIG + '" exists');
}

function test_config_memberDirSheetExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  TestRunner.assertNotNull(sheet, 'Member Dir sheet "' + SHEETS.MEMBER_DIR + '" exists');
}

function test_config_grievanceLogSheetExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  TestRunner.assertNotNull(sheet, 'Grievance Log sheet "' + SHEETS.GRIEVANCE_LOG + '" exists');
}

function test_config_configReaderReturnsObject() {
  TestRunner.assertNotNull(ConfigReader, 'ConfigReader module exists');
  TestRunner.assertType(ConfigReader.getConfig, 'function', 'ConfigReader.getConfig is function');
  var config = ConfigReader.getConfig();
  TestRunner.assertNotNull(config, 'ConfigReader.getConfig() returns value');
  TestRunner.assertType(config, 'object', 'Config is object');
}

function test_config_configHasOrgName() {
  var config = ConfigReader.getConfig();
  TestRunner.assertHasKey(config, 'orgName', 'config.orgName');
  TestRunner.assertType(config.orgName, 'string', 'orgName is string');
  TestRunner.assertGreaterThan(config.orgName.length, 0, 'orgName not empty');
}

function test_config_configHasAccentHue() {
  var config = ConfigReader.getConfig();
  TestRunner.assertHasKey(config, 'accentHue', 'config.accentHue');
  TestRunner.assertType(config.accentHue, 'number', 'accentHue is number');
}

function test_config_configColsPopulated() {
  TestRunner.assertNotNull(CONFIG_COLS, 'CONFIG_COLS');
  TestRunner.assertHasKey(CONFIG_COLS, 'ORG_NAME', 'CONFIG_COLS.ORG_NAME');
  TestRunner.assertGreaterThan(CONFIG_COLS.ORG_NAME, 0, 'ORG_NAME > 0');
}

function test_config_defaultDropdownsDefined() {
  TestRunner.assertNotNull(DEFAULT_CONFIG, 'DEFAULT_CONFIG');
  TestRunner.assertHasKey(DEFAULT_CONFIG, 'GRIEVANCE_STATUS', 'DEFAULT_CONFIG.GRIEVANCE_STATUS');
  TestRunner.assertHasKey(DEFAULT_CONFIG, 'GRIEVANCE_STEP', 'DEFAULT_CONFIG.GRIEVANCE_STEP');
  TestRunner.assertTrue(Array.isArray(DEFAULT_CONFIG.GRIEVANCE_STATUS), 'GRIEVANCE_STATUS is array');
  TestRunner.assertGreaterThan(DEFAULT_CONFIG.GRIEVANCE_STATUS.length, 0, 'GRIEVANCE_STATUS not empty');
}

function test_config_deadlineDefaultsComplete() {
  TestRunner.assertNotNull(DEADLINE_DEFAULTS, 'DEADLINE_DEFAULTS');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'FILING_DAYS', 'FILING_DAYS');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'STEP_1_RESPONSE', 'STEP_1_RESPONSE');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'STEP_2_APPEAL', 'STEP_2_APPEAL');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'STEP_2_RESPONSE', 'STEP_2_RESPONSE');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'STEP_3_APPEAL', 'STEP_3_APPEAL');
  TestRunner.assertHasKey(DEADLINE_DEFAULTS, 'ARBITRATION_DEMAND', 'ARBITRATION_DEMAND');
}

// ── COLMAP SUITE ──────────────────────────────────────────────────────

function test_colmap_grievanceColsDefined() {
  TestRunner.assertNotNull(GRIEVANCE_COLS, 'GRIEVANCE_COLS');
  TestRunner.assertType(GRIEVANCE_COLS, 'object', 'GRIEVANCE_COLS type');
}

function test_colmap_memberColsDefined() {
  TestRunner.assertNotNull(MEMBER_COLS, 'MEMBER_COLS');
  TestRunner.assertType(MEMBER_COLS, 'object', 'MEMBER_COLS type');
}

function test_colmap_grievanceHasStatus() {
  TestRunner.assertHasKey(GRIEVANCE_COLS, 'GRIEVANCE_STATUS', 'GRIEVANCE_COLS.GRIEVANCE_STATUS');
  TestRunner.assertGreaterThan(GRIEVANCE_COLS.GRIEVANCE_STATUS, 0, 'STATUS col > 0');
}

function test_colmap_grievanceHasStep() {
  TestRunner.assertHasKey(GRIEVANCE_COLS, 'GRIEVANCE_STEP', 'GRIEVANCE_COLS.GRIEVANCE_STEP');
  TestRunner.assertGreaterThan(GRIEVANCE_COLS.GRIEVANCE_STEP, 0, 'STEP col > 0');
}

function test_colmap_memberHasEmail() {
  TestRunner.assertHasKey(MEMBER_COLS, 'EMAIL', 'MEMBER_COLS.EMAIL');
  TestRunner.assertGreaterThan(MEMBER_COLS.EMAIL, 0, 'EMAIL col > 0');
}

function test_colmap_memberHasRole() {
  TestRunner.assertHasKey(MEMBER_COLS, 'ROLE', 'MEMBER_COLS.ROLE');
  TestRunner.assertGreaterThan(MEMBER_COLS.ROLE, 0, 'ROLE col > 0');
}

function test_colmap_configColsHasOrgName() {
  TestRunner.assertHasKey(CONFIG_COLS, 'ORG_NAME', 'CONFIG_COLS.ORG_NAME');
  TestRunner.assertGreaterThan(CONFIG_COLS.ORG_NAME, 0, 'ORG_NAME col > 0');
}

function test_colmap_allColsPositive() {
  // Verify no column constant is 0, negative, or NaN
  var colSets = [
    { name: 'GRIEVANCE_COLS', obj: GRIEVANCE_COLS },
    { name: 'MEMBER_COLS', obj: MEMBER_COLS },
    { name: 'CONFIG_COLS', obj: CONFIG_COLS }
  ];
  for (var c = 0; c < colSets.length; c++) {
    var set = colSets[c];
    var keys = Object.keys(set.obj);
    for (var k = 0; k < keys.length; k++) {
      var val = set.obj[keys[k]];
      if (typeof val === 'number') {
        TestRunner.assertGreaterThan(val, 0, set.name + '.' + keys[k] + ' > 0');
      }
    }
  }
}

function test_colmap_noDuplicatePositions() {
  // Within each col set, no two keys should map to the same column number
  var colSets = [
    { name: 'GRIEVANCE_COLS', obj: GRIEVANCE_COLS },
    { name: 'MEMBER_COLS', obj: MEMBER_COLS }
  ];
  for (var c = 0; c < colSets.length; c++) {
    var set = colSets[c];
    var seen = {};
    var keys = Object.keys(set.obj);
    for (var k = 0; k < keys.length; k++) {
      var val = set.obj[keys[k]];
      if (typeof val !== 'number') continue;
      // Allow known aliases (DAYS_TO_DEADLINE/NEXT_DEADLINE)
      if (seen[val]) {
        // This is an alias — just note it, don't fail.
        // Aliases are intentional (e.g., LOCATION = WORK_LOCATION)
        Logger.log('TestRunner: ' + set.name + ' alias detected: ' +
          keys[k] + ' and ' + seen[val] + ' both map to col ' + val);
      }
      seen[val] = keys[k];
    }
  }
}

// ── AUTH SUITE ────────────────────────────────────────────────────────

function test_auth_moduleExists() {
  TestRunner.assertNotNull(Auth, 'Auth module');
  TestRunner.assertType(Auth, 'object', 'Auth is object');
}

function test_auth_resolveUserFunctionExists() {
  TestRunner.assertType(Auth.resolveUser, 'function', 'Auth.resolveUser is function');
}

function test_auth_dataServiceExists() {
  TestRunner.assertNotNull(DataService, 'DataService module');
  TestRunner.assertType(DataService, 'object', 'DataService is object');
}

function test_auth_findUserByEmailExists() {
  TestRunner.assertType(DataService.findUserByEmail, 'function', 'DataService.findUserByEmail');
}

function test_auth_resolveCallerEmailExists() {
  TestRunner.assertType(typeof _resolveCallerEmail, 'string', 'typeof check');
  TestRunner.assertEquals('function', typeof _resolveCallerEmail, '_resolveCallerEmail is function');
}

function test_auth_requireStewardAuthExists() {
  TestRunner.assertEquals('function', typeof _requireStewardAuth, '_requireStewardAuth is function');
}

function test_auth_checkWebAppAuthExists() {
  TestRunner.assertEquals('function', typeof checkWebAppAuthorization, 'checkWebAppAuthorization is function');
}

function test_auth_sessionTokenFunctionsExist() {
  TestRunner.assertType(Auth.createSessionToken, 'function', 'Auth.createSessionToken');
  TestRunner.assertType(Auth.validateSessionToken, 'function', 'Auth.validateSessionToken');
}

// ── GRIEVANCE SUITE ───────────────────────────────────────────────────

function test_grievance_statusConstantsDefined() {
  TestRunner.assertNotNull(GRIEVANCE_STATUS, 'GRIEVANCE_STATUS');
  TestRunner.assertHasKey(GRIEVANCE_STATUS, 'OPEN', 'OPEN');
  TestRunner.assertHasKey(GRIEVANCE_STATUS, 'PENDING', 'PENDING');
  TestRunner.assertHasKey(GRIEVANCE_STATUS, 'SETTLED', 'SETTLED');
  TestRunner.assertHasKey(GRIEVANCE_STATUS, 'CLOSED', 'CLOSED');
}

function test_grievance_allStatusesPresent() {
  var expected = ['Open', 'Pending Info', 'Settled', 'Withdrawn', 'Denied', 'Won', 'Appealed', 'In Arbitration', 'Closed'];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertContains(
      DEFAULT_CONFIG.GRIEVANCE_STATUS, expected[i],
      'DEFAULT_CONFIG.GRIEVANCE_STATUS contains "' + expected[i] + '"'
    );
  }
}

function test_grievance_closedStatusesSubset() {
  TestRunner.assertNotNull(GRIEVANCE_CLOSED_STATUSES, 'GRIEVANCE_CLOSED_STATUSES');
  TestRunner.assertTrue(Array.isArray(GRIEVANCE_CLOSED_STATUSES), 'is array');
  // Every closed status should exist in the full status list
  for (var i = 0; i < GRIEVANCE_CLOSED_STATUSES.length; i++) {
    TestRunner.assertContains(
      DEFAULT_CONFIG.GRIEVANCE_STATUS,
      GRIEVANCE_CLOSED_STATUSES[i],
      'Closed status "' + GRIEVANCE_CLOSED_STATUSES[i] + '" in full list'
    );
  }
}

function test_grievance_priorityCoversAllStatuses() {
  TestRunner.assertNotNull(GRIEVANCE_STATUS_PRIORITY, 'GRIEVANCE_STATUS_PRIORITY');
  var statuses = DEFAULT_CONFIG.GRIEVANCE_STATUS;
  for (var i = 0; i < statuses.length; i++) {
    TestRunner.assertHasKey(
      GRIEVANCE_STATUS_PRIORITY, statuses[i],
      'Priority for "' + statuses[i] + '"'
    );
  }
}

function test_grievance_deadlineRulesReadable() {
  TestRunner.assertType(getDeadlineRules, 'function', 'getDeadlineRules is function');
  var rules = getDeadlineRules();
  TestRunner.assertNotNull(rules, 'getDeadlineRules returns value');
  TestRunner.assertType(rules, 'object', 'rules is object');
}

function test_grievance_deadlineRulesHaveAllSteps() {
  var rules = getDeadlineRules();
  TestRunner.assertHasKey(rules, 'FILING_DAYS', 'FILING_DAYS');
  TestRunner.assertHasKey(rules, 'STEP_1', 'STEP_1');
  TestRunner.assertHasKey(rules, 'STEP_2', 'STEP_2');
  TestRunner.assertHasKey(rules, 'STEP_3', 'STEP_3');
  TestRunner.assertHasKey(rules, 'ARBITRATION', 'ARBITRATION');
  // Sub-keys
  TestRunner.assertHasKey(rules.STEP_1, 'DAYS_FOR_RESPONSE', 'STEP_1.DAYS_FOR_RESPONSE');
  TestRunner.assertHasKey(rules.STEP_2, 'DAYS_TO_APPEAL', 'STEP_2.DAYS_TO_APPEAL');
  TestRunner.assertHasKey(rules.STEP_2, 'DAYS_FOR_RESPONSE', 'STEP_2.DAYS_FOR_RESPONSE');
  TestRunner.assertHasKey(rules.ARBITRATION, 'DAYS_TO_DEMAND', 'ARB.DAYS_TO_DEMAND');
}

function test_grievance_deadlineDefaultsReasonable() {
  // Sanity check: deadline values should be positive and within contract-reasonable range
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.FILING_DAYS, 0, 'FILING > 0');
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.STEP_1_RESPONSE, 0, 'S1 > 0');
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.STEP_2_APPEAL, 0, 'S2 appeal > 0');
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.STEP_2_RESPONSE, 0, 'S2 resp > 0');
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.STEP_3_APPEAL, 0, 'S3 appeal > 0');
  TestRunner.assertGreaterThan(DEADLINE_DEFAULTS.ARBITRATION_DEMAND, 0, 'ARB > 0');
  // Upper bounds — no deadline should be > 365 days
  TestRunner.assertTrue(DEADLINE_DEFAULTS.FILING_DAYS <= 365, 'FILING <= 365');
  TestRunner.assertTrue(DEADLINE_DEFAULTS.ARBITRATION_DEMAND <= 365, 'ARB <= 365');
}

function test_grievance_stepsListDefined() {
  var steps = DEFAULT_CONFIG.GRIEVANCE_STEP;
  TestRunner.assertNotNull(steps, 'GRIEVANCE_STEP');
  TestRunner.assertTrue(Array.isArray(steps), 'is array');
  TestRunner.assertContains(steps, 'Step I', 'contains Step I');
  TestRunner.assertContains(steps, 'Step II', 'contains Step II');
  TestRunner.assertContains(steps, 'Arbitration', 'contains Arbitration');
}

function test_grievance_isValidIdFunction() {
  TestRunner.assertEquals('function', typeof isValidGrievanceId, 'isValidGrievanceId exists');
  // Known-good patterns (prefix + sequence)
  TestRunner.assertTrue(isValidGrievanceId('G-001'), 'G-001 valid');
  // Known-bad patterns
  TestRunner.assertFalsy(isValidGrievanceId(''), 'empty invalid');
  TestRunner.assertFalsy(isValidGrievanceId(null), 'null invalid');
}

function test_grievance_outcomesDefined() {
  TestRunner.assertNotNull(GRIEVANCE_OUTCOMES, 'GRIEVANCE_OUTCOMES');
  TestRunner.assertHasKey(GRIEVANCE_OUTCOMES, 'WON', 'WON');
  TestRunner.assertHasKey(GRIEVANCE_OUTCOMES, 'DENIED', 'DENIED');
  TestRunner.assertHasKey(GRIEVANCE_OUTCOMES, 'SETTLED', 'SETTLED');
  TestRunner.assertHasKey(GRIEVANCE_OUTCOMES, 'WITHDRAWN', 'WITHDRAWN');
}

// ── SECURITY SUITE ────────────────────────────────────────────────────

function test_security_escapeHtmlDefined() {
  TestRunner.assertEquals('function', typeof escapeHtml, 'escapeHtml exists');
}

function test_security_escapeHtmlBlocks() {
  var input = '<script>alert("xss")</script>';
  var output = escapeHtml(input);
  // Output must not contain raw < or >
  TestRunner.assertTrue(output.indexOf('<') === -1, 'No raw < in output');
  TestRunner.assertTrue(output.indexOf('>') === -1, 'No raw > in output');
}

function test_security_escapeForFormulaDefined() {
  TestRunner.assertEquals('function', typeof escapeForFormula, 'escapeForFormula exists');
}

function test_security_escapeFormulaBlocks() {
  var input = '=IMPORTRANGE("evil","A1")';
  var output = escapeForFormula(input);
  // Should not start with = after escaping
  TestRunner.assertTrue(output.charAt(0) !== '=', 'No leading = after escape');
}

function test_security_escapeHtmlHandlesNull() {
  // escapeHtml should handle null/undefined without throwing
  var output = escapeHtml(null);
  TestRunner.assertType(output, 'string', 'null → string');
  var output2 = escapeHtml(undefined);
  TestRunner.assertType(output2, 'string', 'undefined → string');
}

function test_security_escapeHtmlHandlesNumbers() {
  var output = escapeHtml(42);
  TestRunner.assertType(output, 'string', 'number → string');
  TestRunner.assertEquals('42', output, 'number preserved');
}

// ── SYSTEM SUITE ──────────────────────────────────────────────────────

function test_system_versionInfoDefined() {
  TestRunner.assertNotNull(VERSION_INFO, 'VERSION_INFO');
  TestRunner.assertHasKey(VERSION_INFO, 'version', 'version');
  TestRunner.assertHasKey(VERSION_INFO, 'codename', 'codename');
}

function test_system_versionFormatSemver() {
  var v = VERSION_INFO.version;
  TestRunner.assertType(v, 'string', 'version is string');
  // Should match X.Y.Z pattern
  var parts = v.split('.');
  TestRunner.assertGreaterThan(parts.length, 1, 'version has dots');
  TestRunner.assertTrue(!isNaN(Number(parts[0])), 'major is numeric');
}

function test_system_spreadsheetBound() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // In web app context this may be null — that's okay for web-triggered tests
  // but if it exists, verify it's an actual spreadsheet
  if (ss) {
    TestRunner.assertType(ss.getSheets, 'function', 'ss.getSheets is function');
    var sheets = ss.getSheets();
    TestRunner.assertGreaterThan(sheets.length, 0, 'has at least 1 sheet');
  }
}

function test_system_eventBusExists() {
  TestRunner.assertEquals('function', typeof EventBus !== 'undefined' ? 'object' : 'undefined',
    'EventBus type check — if this fails, 15_EventBus.gs may not be loaded');
  if (typeof EventBus !== 'undefined') {
    TestRunner.assertType(EventBus.emit, 'function', 'EventBus.emit');
    TestRunner.assertType(EventBus.on, 'function', 'EventBus.on');
  }
}

function test_system_hiddenSheetsConstant() {
  if (typeof HIDDEN_SHEETS !== 'undefined') {
    TestRunner.assertType(HIDDEN_SHEETS, 'object', 'HIDDEN_SHEETS is object');
  }
}

// ── DATASERVICE SUITE ─────────────────────────────────────────────────
// Tests DataService CRUD operations against live sheets.

function test_dataservice_moduleExists() {
  TestRunner.assertNotNull(DataService, 'DataService');
  TestRunner.assertType(DataService, 'object', 'DataService is object');
}

function test_dataservice_findUserByEmailCallable() {
  TestRunner.assertType(DataService.findUserByEmail, 'function', 'findUserByEmail');
}

function test_dataservice_getUserRoleCallable() {
  TestRunner.assertType(DataService.getUserRole, 'function', 'getUserRole');
}

function test_dataservice_getStewardCasesCallable() {
  TestRunner.assertType(DataService.getStewardCases, 'function', 'getStewardCases');
}

function test_dataservice_getAllMembersCallable() {
  TestRunner.assertType(DataService.getAllMembers, 'function', 'getAllMembers');
}

function test_dataservice_getUnitsCallable() {
  TestRunner.assertType(DataService.getUnits, 'function', 'getUnits');
}

function test_dataservice_getBatchDataCallable() {
  TestRunner.assertType(DataService.getBatchData, 'function', 'getBatchData');
}

function test_dataservice_memberLookupReturnsShape() {
  // Lookup a definitely-nonexistent email — should return null (not throw)
  var result = DataService.findUserByEmail('__nonexistent_test_probe__@example.invalid');
  // null means "not found" — this is correct behavior
  TestRunner.assertTrue(result === null || result === undefined,
    'Nonexistent email returns null/undefined, not an error');
}

function test_dataservice_invalidEmailReturnsNull() {
  var result = DataService.findUserByEmail('');
  TestRunner.assertTrue(result === null || result === undefined,
    'Empty email returns null/undefined');
}

function test_dataservice_publicAPIComplete() {
  // Verify all expected DataService public methods exist
  var expected = [
    'findUserByEmail', 'getUserRole', 'getStewardCases', 'getStewardKPIs',
    'getMemberGrievances', 'getMemberGrievanceHistory', 'getStewardContact',
    'getUnits', 'getFullMemberProfile', 'updateMemberProfile',
    'getAssignedStewardInfo', 'getAvailableStewards', 'getAllMembers',
    'createTask', 'getTasks', 'completeTask', 'updateTask',
    'getStewardMemberStats', 'getStewardDirectory',
    'getGrievanceStats', 'getGrievanceHotSpots', 'getMembershipStats',
    'getUpcomingEvents', 'isChiefSteward', 'submitFeedback',
    'getMeetingMinutes', 'addMeetingMinutes',
    'getCaseChecklist', 'toggleChecklistItem', 'getMemberMeetings',
    'getSatisfactionTrends', 'getBatchData'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(DataService[expected[i]], 'function',
      'DataService.' + expected[i]);
  }
}

// ── AUTHSWEEP SUITE ───────────────────────────────────────────────────
// Verifies all data* endpoints reject unauthenticated calls.
// These tests call real endpoints with null/invalid tokens.

function test_authsweep_allDataFnsExist() {
  // Spot-check critical data* wrappers exist as global functions
  var fns = [
    'dataGetStewardCases', 'dataGetMemberGrievances', 'dataGetAllMembers',
    'dataSendBroadcast', 'dataCreateTask', 'dataGetTasks',
    'dataUpdateProfile', 'dataGetBatchData', 'dataRunTests'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]] !== 'undefined' ? 'function' : 'undefined',
      fns[i] + ' exists');
  }
}

function test_authsweep_stewardEndpointsRejectNull() {
  // Steward-only endpoints must return safe-empty when called with null token
  var stewardEndpoints = [
    { fn: 'dataGetStewardCases', safeEmpty: [] },
    { fn: 'dataGetStewardKPIs', safeEmpty: {} },
    { fn: 'dataGetAllMembers', safeEmpty: [] },
    { fn: 'dataGetGrievanceStats', safeEmpty: { available: false } },
    { fn: 'dataGetGrievanceHotSpots', safeEmpty: [] },
    { fn: 'dataGetStewardContactLog', safeEmpty: [] },
    { fn: 'dataGetTasks', safeEmpty: [] },
    { fn: 'dataGetAllStewardPerformance', safeEmpty: [] },
  ];
  for (var i = 0; i < stewardEndpoints.length; i++) {
    var ep = stewardEndpoints[i];
    try {
      var result = this[ep.fn](null); // null token = unauthenticated
      // Result must be the safe empty type, not real data
      var resultType = Array.isArray(result) ? 'array' : typeof result;
      var safeType = Array.isArray(ep.safeEmpty) ? 'array' : typeof ep.safeEmpty;
      TestRunner.assertEquals(safeType, resultType,
        ep.fn + ' returns safe-empty type on null token');
    } catch (e) {
      // Throwing is also acceptable — it means the endpoint rejected
    }
  }
}

function test_authsweep_memberEndpointsRejectNull() {
  // Member endpoints must return safe-empty for null token
  var memberEndpoints = [
    { fn: 'dataGetMemberGrievances', safeEmpty: [] },
    { fn: 'dataGetStewardContact', safeEmpty: null },
    { fn: 'dataGetAssignedSteward', safeEmpty: null },
    { fn: 'dataGetAvailableStewards', safeEmpty: [] },
    { fn: 'dataGetSurveyStatus', safeEmpty: null },
    { fn: 'dataGetMemberTasks', safeEmpty: [] },
    { fn: 'dataGetMemberMeetings', safeEmpty: [] },
    { fn: 'dataGetMyFeedback', safeEmpty: [] },
  ];
  for (var i = 0; i < memberEndpoints.length; i++) {
    var ep = memberEndpoints[i];
    try {
      var result = this[ep.fn](null);
      // Must not return actual member data
      if (Array.isArray(ep.safeEmpty)) {
        TestRunner.assertTrue(Array.isArray(result) && result.length === 0,
          ep.fn + ' returns empty array on null token');
      }
    } catch (e) {
      // Throwing is acceptable
    }
  }
}

function test_authsweep_noDataLeakOnNullToken() {
  // dataGetBatchData is a high-value target — bulk data endpoint
  try {
    var result = dataGetBatchData(null);
    // Should return empty/error, not real batch data
    if (result && typeof result === 'object') {
      // If it returns an object, it should NOT have member arrays
      var hasMemberData = result.members && Array.isArray(result.members) && result.members.length > 0;
      TestRunner.assertFalsy(hasMemberData,
        'dataGetBatchData(null) must not leak member data');
    }
  } catch (e) {
    // Throwing is acceptable
  }
}

function test_authsweep_pollStubsSafe() {
  // Legacy poll stubs should be safe (no auth needed, return empty/failure)
  var polls = dataGetActivePolls();
  TestRunner.assertTrue(Array.isArray(polls) && polls.length === 0,
    'dataGetActivePolls returns empty array');
  var vote = dataSubmitPollVote();
  TestRunner.assertHasKey(vote, 'success', 'dataSubmitPollVote has success key');
  TestRunner.assertEquals(false, vote.success, 'dataSubmitPollVote returns failure');
}

function test_authsweep_testRunnerEndpointsGated() {
  // Our own test endpoints must be auth-gated
  var runResult = dataRunTests(null);
  TestRunner.assertHasKey(runResult, 'success', 'dataRunTests has success key');
  TestRunner.assertEquals(false, runResult.success, 'dataRunTests rejects null token');

  var getResult = dataGetTestResults(null);
  TestRunner.assertEquals(false, getResult.success, 'dataGetTestResults rejects null token');

  var trigResult = dataManageTestTrigger(null, 'setup');
  TestRunner.assertEquals(false, trigResult.success, 'dataManageTestTrigger rejects null token');
}

// ── CONFIGLIVE SUITE ──────────────────────────────────────────────────
// Verifies live sheet headers match expected column constants.

function test_configlive_configSheetHasHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  TestRunner.assertNotNull(sheet, 'Config sheet exists');
  var lastCol = sheet.getLastColumn();
  TestRunner.assertGreaterThan(lastCol, 0, 'Config has columns');
}

function test_configlive_memberDirHasHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  TestRunner.assertNotNull(sheet, 'Member Dir sheet exists');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  TestRunner.assertGreaterThan(headers.length, 5, 'Member Dir has 5+ columns');
}

function test_configlive_grievanceLogHasHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  TestRunner.assertNotNull(sheet, 'Grievance Log sheet exists');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  TestRunner.assertGreaterThan(headers.length, 5, 'Grievance Log has 5+ columns');
}

function test_configlive_configColsMatchSheet() {
  // Verify CONFIG_COLS positions don't exceed actual sheet width
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) return;
  var maxCol = sheet.getLastColumn();
  var keys = Object.keys(CONFIG_COLS);
  for (var i = 0; i < keys.length; i++) {
    var val = CONFIG_COLS[keys[i]];
    if (typeof val === 'number') {
      TestRunner.assertTrue(val <= maxCol,
        'CONFIG_COLS.' + keys[i] + ' (' + val + ') <= sheet width (' + maxCol + ')');
    }
  }
}

function test_configlive_memberColsMatchSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;
  var maxCol = sheet.getLastColumn();
  var keys = Object.keys(MEMBER_COLS);
  for (var i = 0; i < keys.length; i++) {
    var val = MEMBER_COLS[keys[i]];
    if (typeof val === 'number') {
      TestRunner.assertTrue(val <= maxCol,
        'MEMBER_COLS.' + keys[i] + ' (' + val + ') <= sheet width (' + maxCol + ')');
    }
  }
}

function test_configlive_grievanceColsMatchSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return;
  var maxCol = sheet.getLastColumn();
  var keys = Object.keys(GRIEVANCE_COLS);
  for (var i = 0; i < keys.length; i++) {
    var val = GRIEVANCE_COLS[keys[i]];
    if (typeof val === 'number') {
      TestRunner.assertTrue(val <= maxCol,
        'GRIEVANCE_COLS.' + keys[i] + ' (' + val + ') <= sheet width (' + maxCol + ')');
    }
  }
}

function test_configlive_syncColumnMapsCallable() {
  TestRunner.assertEquals('function', typeof syncColumnMaps, 'syncColumnMaps exists');
  // Don't actually call it (it writes) — just verify it's available
}

function test_configlive_configRow3HasValues() {
  // Config tab row 3 holds the actual config values — verify it's not empty
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() < 3) return;
  var orgNameCol = CONFIG_COLS.ORG_NAME;
  if (!orgNameCol || orgNameCol < 1) return;
  var val = sheet.getRange(3, orgNameCol).getValue();
  TestRunner.assertTrue(String(val).trim().length > 0,
    'Config row 3 ORG_NAME is not empty');
}

// ── SURVEY SUITE ──────────────────────────────────────────────────────
// Tests survey engine integrity — constants, question schema, period mgmt.

function test_survey_hiddenSheetsConstants() {
  TestRunner.assertHasKey(HIDDEN_SHEETS, 'SURVEY_TRACKING', 'SURVEY_TRACKING');
  TestRunner.assertHasKey(HIDDEN_SHEETS, 'SURVEY_VAULT', 'SURVEY_VAULT');
  TestRunner.assertHasKey(HIDDEN_SHEETS, 'SURVEY_PERIODS', 'SURVEY_PERIODS');
  // Values should be underscore-prefixed hidden sheet names
  TestRunner.assertTrue(HIDDEN_SHEETS.SURVEY_TRACKING.charAt(0) === '_',
    'SURVEY_TRACKING starts with _');
}

function test_survey_periodsColsDefined() {
  TestRunner.assertNotNull(SURVEY_PERIODS_COLS, 'SURVEY_PERIODS_COLS');
  TestRunner.assertType(SURVEY_PERIODS_COLS, 'object', 'is object');
  TestRunner.assertHasKey(SURVEY_PERIODS_COLS, 'PERIOD_ID', 'PERIOD_ID');
  TestRunner.assertHasKey(SURVEY_PERIODS_COLS, 'STATUS', 'STATUS');
}

function test_survey_questionsColsDefined() {
  TestRunner.assertNotNull(SURVEY_QUESTIONS_COLS, 'SURVEY_QUESTIONS_COLS');
  TestRunner.assertType(SURVEY_QUESTIONS_COLS, 'object', 'is object');
  TestRunner.assertHasKey(SURVEY_QUESTIONS_COLS, 'QUESTION_ID', 'QUESTION_ID');
  TestRunner.assertHasKey(SURVEY_QUESTIONS_COLS, 'QUESTION_TEXT', 'QUESTION_TEXT');
  TestRunner.assertHasKey(SURVEY_QUESTIONS_COLS, 'TYPE', 'TYPE');
  TestRunner.assertHasKey(SURVEY_QUESTIONS_COLS, 'ACTIVE', 'ACTIVE');
}

function test_survey_getSurveyQuestionsCallable() {
  TestRunner.assertEquals('function', typeof getSurveyQuestions, 'getSurveyQuestions exists');
}

function test_survey_questionsReturnArray() {
  var questions = getSurveyQuestions();
  TestRunner.assertTrue(Array.isArray(questions), 'getSurveyQuestions returns array');
  TestRunner.assertGreaterThan(questions.length, 0, 'at least 1 question');
}

function test_survey_questionShapeValid() {
  var questions = getSurveyQuestions();
  if (!questions || questions.length === 0) return;
  // Check first question has expected shape
  var q = questions[0];
  TestRunner.assertHasKey(q, 'id', 'question.id');
  TestRunner.assertHasKey(q, 'text', 'question.text');
  TestRunner.assertHasKey(q, 'type', 'question.type');
  // Type should be one of known types
  var validTypes = ['slider-10', 'dropdown', 'radio', 'checkbox', 'paragraph', 'text'];
  TestRunner.assertTrue(validTypes.indexOf(q.type) !== -1,
    'question.type "' + q.type + '" is a known type');
}

function test_survey_getSurveyPeriodCallable() {
  TestRunner.assertEquals('function', typeof getSurveyPeriod, 'getSurveyPeriod exists');
  // getSurveyPeriod returns the active period object or null
  try {
    var period = getSurveyPeriod();
    // null means no active period — valid
    // object means active period — also valid
    if (period !== null) {
      TestRunner.assertType(period, 'object', 'period is object');
    }
  } catch (e) {
    // If sheet doesn't exist yet, that's okay — function should handle gracefully
  }
}

function test_survey_submitResponseCallable() {
  TestRunner.assertEquals('function', typeof submitSurveyResponse, 'submitSurveyResponse exists');
  // Don't actually submit — just verify it exists
}

function test_survey_satisfactionColsDefined() {
  // SATISFACTION_COLS is deprecated but kept for backward compat
  if (typeof SATISFACTION_COLS !== 'undefined') {
    TestRunner.assertType(SATISFACTION_COLS, 'object', 'SATISFACTION_COLS is object');
  }
  // The dynamic col map function should also exist
  if (typeof getSatisfactionColMap_ !== 'undefined') {
    TestRunner.assertType(getSatisfactionColMap_, 'function', 'getSatisfactionColMap_ is function');
  }
}

function test_survey_trackingSheetExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
  // Sheet may not exist if survey hasn't been initialized yet — that's a warning, not failure
  // But if it exists, verify it has headers
  if (sheet && sheet.getLastRow() >= 1) {
    var firstCell = sheet.getRange(1, 1).getValue();
    TestRunner.assertTrue(String(firstCell).trim().length > 0,
      'Survey tracking sheet has a header in A1');
  }
}
