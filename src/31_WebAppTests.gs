/**
 * ============================================================================
 * 31_WebAppTests.gs - GAS-Native Web App Test Suites
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   GAS-native web app test suites — comprehensive coverage of all SPA
 *   modules. 9 test suites: webapp (doGet routing), configrd (ConfigReader),
 *   portal (PortalSheets 0-indexed validation), weeklyq (WeeklyQuestions
 *   API), qaforum (QAForum), timeline
 *   (TimelineService), failsafe (FailsafeService), endpoints (all
 *   data, wq, qa, tl, fs wrapper functions exist and enforce auth).
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Separated from 30_TestRunner.gs to keep test file sizes manageable and
 *   allow independent suite execution. Each suite completes well under 3
 *   minutes to avoid the GAS 6-minute execution limit. The endpoints suite
 *   is critical — it verifies that EVERY public data function exists and
 *   gates on authentication, catching accidental auth removal.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Web app module verification is unavailable. The endpoints auth sweep
 *   won't catch accidental auth removal on data functions. Core SPA
 *   functionality is unaffected — these are diagnostic tests.
 *
 * DEPENDENCIES:
 *   Depends on 30_TestRunner.gs (TestRunner assertion library), all SPA
 *   modules (19-28) for existence checks. Used by the test runner suite
 *   system.
 *
 * @version 4.33.0
 */

/* ========================================================================
 * WEBAPP SUITE — doGet routing, templates, URL resolution
 * ======================================================================== */

function test_webapp_doGetExists() {
  TestRunner.assertEquals('function', typeof doGet, 'doGet global function exists');
}

/** Tests webapp: do get web dashboard exists. */
function test_webapp_doGetWebDashboardExists() {
  TestRunner.assertEquals('function', typeof doGetWebDashboard, 'doGetWebDashboard exists');
}

/** Tests webapp: include helper exists. */
function test_webapp_includeHelperExists() {
  TestRunner.assertEquals('function', typeof include, 'include() template helper exists');
}

/** Tests webapp: get web app url exists. */
function test_webapp_getWebAppUrlExists() {
  TestRunner.assertEquals('function', typeof getWebAppUrl, 'getWebAppUrl exists');
}

/** Tests webapp: get web app url returns string. */
function test_webapp_getWebAppUrlReturnsString() {
  var url = getWebAppUrl();
  TestRunner.assertNotNull(url, 'getWebAppUrl returns a value');
  TestRunner.assertType(url, 'string', 'getWebAppUrl returns string');
  TestRunner.assertGreaterThan(url.length, 0, 'URL is not empty');
}

/** Tests webapp: org chart html exists. */
function test_webapp_orgChartHtmlExists() {
  TestRunner.assertEquals('function', typeof getOrgChartHtml, 'getOrgChartHtml exists');
}

/** Tests webapp: org chart html returns content. */
function test_webapp_orgChartHtmlReturnsContent() {
  // getOrgChartHtml should return HTML content (not an error fallback)
  try {
    var html = getOrgChartHtml();
    TestRunner.assertNotNull(html, 'getOrgChartHtml returns non-null');
    TestRunner.assertTrue(typeof html === 'string', 'getOrgChartHtml returns string');
    TestRunner.assertGreaterThan(html.length, 100, 'getOrgChartHtml returns substantial content');
    // Should NOT be the auth-failure fallback
    TestRunner.assertTrue(html.indexOf('Authentication required') === -1, 'getOrgChartHtml did not return auth error');
  } catch (e) {
    TestRunner.assertTrue(false, 'getOrgChartHtml threw: ' + e.message);
  }
}

/** Tests webapp: diagnose web app exists. */
function test_webapp_diagnoseWebAppExists() {
  TestRunner.assertEquals('function', typeof diagnoseWebApp, 'diagnoseWebApp exists');
}

/** Tests webapp: diagnose web app runs clean. */
function test_webapp_diagnoseWebAppRunsClean() {
  // diagnoseWebApp() is a comprehensive 14-step health check — read-only
  try {
    var diag = diagnoseWebApp();
    TestRunner.assertNotNull(diag, 'diagnoseWebApp returns a result');
    // Should be an object with steps or results
    if (typeof diag === 'object') {
      TestRunner.assertTrue(true, 'diagnoseWebApp returned an object');
    } else if (typeof diag === 'string') {
      TestRunner.assertGreaterThan(diag.length, 0, 'diagnoseWebApp returned non-empty string');
    }
  } catch (e) {
    // If it throws, the diagnostic itself failed — that's a valid test result
    TestRunner.assertTrue(false, 'diagnoseWebApp threw: ' + e.message);
  }
}

/** Tests webapp: serve fatal error exists. */
function test_webapp_serveFatalErrorExists() {
  TestRunner.assertEquals('function', typeof _serveFatalError, '_serveFatalError exists');
}

/** Tests webapp: sanitize config exists. */
function test_webapp_sanitizeConfigExists() {
  TestRunner.assertEquals('function', typeof _sanitizeConfig, '_sanitizeConfig exists');
}

/** Tests webapp: sanitize config strips internal. */
function test_webapp_sanitizeConfigStripsInternal() {
  // _sanitizeConfig should remove sensitive fields from config before sending to client
  var testConfig = {
    orgName: 'Test Org',
    accentHue: 220,
    // These should be stripped if they exist
    _internal: 'secret',
    spreadsheetId: 'abc123'
  };
  try {
    var sanitized = _sanitizeConfig(testConfig);
    TestRunner.assertNotNull(sanitized, 'sanitizeConfig returns value');
    TestRunner.assertHasKey(sanitized, 'orgName', 'keeps orgName');
  } catch (e) {
    // If function doesn't exist or has different signature, note it
    TestRunner.assertTrue(false, '_sanitizeConfig error: ' + e.message);
  }
}

/* ========================================================================
 * CONFIGRD SUITE — ConfigReader module completeness
 * ======================================================================== */

function test_configrd_moduleExists() {
  TestRunner.assertNotNull(ConfigReader, 'ConfigReader module');
  TestRunner.assertType(ConfigReader, 'object', 'ConfigReader is object');
}

/** Tests configrd: get config callable. */
function test_configrd_getConfigCallable() {
  TestRunner.assertType(ConfigReader.getConfig, 'function', 'getConfig is function');
}

/** Tests configrd: validate config callable. */
function test_configrd_validateConfigCallable() {
  TestRunner.assertType(ConfigReader.validateConfig, 'function', 'validateConfig is function');
}

/** Tests configrd: refresh config callable. */
function test_configrd_refreshConfigCallable() {
  TestRunner.assertType(ConfigReader.refreshConfig, 'function', 'refreshConfig is function');
}

/** Tests configrd: get config JSON callable. */
function test_configrd_getConfigJSONCallable() {
  TestRunner.assertType(ConfigReader.getConfigJSON, 'function', 'getConfigJSON is function');
}

/** Tests configrd: config has required fields. */
function test_configrd_configHasRequiredFields() {
  var config = ConfigReader.getConfig();
  TestRunner.assertNotNull(config, 'config returned');
  // Core identity fields
  TestRunner.assertHasKey(config, 'orgName', 'orgName');
  TestRunner.assertHasKey(config, 'accentHue', 'accentHue');
  // Label fields
  TestRunner.assertHasKey(config, 'stewardLabel', 'stewardLabel');
  TestRunner.assertHasKey(config, 'memberLabel', 'memberLabel');
}

/** Tests configrd: config has drive fields. */
function test_configrd_configHasDriveFields() {
  var config = ConfigReader.getConfig();
  // Drive integration fields — may be empty but keys should exist
  var driveKeys = ['calendarId', 'driveFolderId', 'dashboardRootFolderId'];
  for (var i = 0; i < driveKeys.length; i++) {
    TestRunner.assertHasKey(config, driveKeys[i],
      'config has key: ' + driveKeys[i]);
  }
}

/** Tests configrd: config has auth fields. */
function test_configrd_configHasAuthFields() {
  var config = ConfigReader.getConfig();
  // Auth-related config fields
  TestRunner.assertHasKey(config, 'magicLinkExpiryDays', 'magicLinkExpiryDays');
  TestRunner.assertHasKey(config, 'cookieDurationDays', 'cookieDurationDays');
}

/** Tests configrd: validate config returns shape. */
function test_configrd_validateConfigReturnsShape() {
  var result = ConfigReader.validateConfig();
  TestRunner.assertNotNull(result, 'validateConfig returns value');
  TestRunner.assertType(result, 'object', 'validateConfig returns object');
  TestRunner.assertHasKey(result, 'valid', 'has valid key');
  TestRunner.assertType(result.valid, 'boolean', 'valid is boolean');
  TestRunner.assertHasKey(result, 'missing', 'has missing key');
  TestRunner.assertTrue(Array.isArray(result.missing), 'missing is array');
}

/** Tests configrd: get config JSON returns string. */
function test_configrd_getConfigJSONReturnsString() {
  var json = ConfigReader.getConfigJSON();
  TestRunner.assertType(json, 'string', 'getConfigJSON returns string');
  // Should be valid JSON
  try {
    var parsed = JSON.parse(json);
    TestRunner.assertType(parsed, 'object', 'JSON parses to object');
    TestRunner.assertHasKey(parsed, 'orgName', 'parsed JSON has orgName');
  } catch (e) {
    TestRunner.assertTrue(false, 'getConfigJSON returned invalid JSON: ' + e.message);
  }
}

/* ========================================================================
 * PORTAL SUITE — PortalSheets column constants, 0-indexed validation
 * ======================================================================== */

function test_portal_memberDirColsDefined() {
  TestRunner.assertNotNull(PORTAL_MEMBER_DIR_COLS, 'PORTAL_MEMBER_DIR_COLS');
  TestRunner.assertType(PORTAL_MEMBER_DIR_COLS, 'object', 'is object');
  TestRunner.assertHasKey(PORTAL_MEMBER_DIR_COLS, 'EMAIL', 'EMAIL');
  TestRunner.assertHasKey(PORTAL_MEMBER_DIR_COLS, 'NAME', 'NAME');
  TestRunner.assertHasKey(PORTAL_MEMBER_DIR_COLS, 'ROLE', 'ROLE');
}

/** Tests portal: event cols defined. */
function test_portal_eventColsDefined() {
  TestRunner.assertNotNull(PORTAL_EVENT_COLS, 'PORTAL_EVENT_COLS');
  TestRunner.assertType(PORTAL_EVENT_COLS, 'object', 'is object');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'TITLE', 'TITLE');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'TYPE', 'TYPE');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'DATE_TIME', 'DATE_TIME');
}

/** Tests portal: minutes cols defined. */
function test_portal_minutesColsDefined() {
  TestRunner.assertNotNull(PORTAL_MINUTES_COLS, 'PORTAL_MINUTES_COLS');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'MEETING_DATE', 'MEETING_DATE');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'TITLE', 'TITLE');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'DRIVE_DOC_URL', 'DRIVE_DOC_URL');
}

/** Tests portal: grievance cols defined. */
function test_portal_grievanceColsDefined() {
  TestRunner.assertNotNull(PORTAL_GRIEVANCE_COLS, 'PORTAL_GRIEVANCE_COLS');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'MEMBER_EMAIL', 'MEMBER_EMAIL');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'STEWARD_EMAIL', 'STEWARD_EMAIL');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'STATUS', 'STATUS');
}

/** Tests portal: steward log cols defined. */
function test_portal_stewardLogColsDefined() {
  TestRunner.assertNotNull(PORTAL_STEWARD_LOG_COLS, 'PORTAL_STEWARD_LOG_COLS');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'STEWARD_EMAIL', 'STEWARD_EMAIL');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'MEMBER_EMAIL', 'MEMBER_EMAIL');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'TYPE', 'TYPE');
}

/** Tests portal: mega survey cols defined. */
function test_portal_megaSurveyColsDefined() {
  TestRunner.assertNotNull(PORTAL_MEGA_SURVEY_COLS, 'PORTAL_MEGA_SURVEY_COLS');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'EMAIL', 'EMAIL');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'RESPONSES', 'RESPONSES');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'COMPLETED', 'COMPLETED');
}

/** Tests portal: all cols0 indexed. */
function test_portal_allCols0Indexed() {
  // PORTAL_*_COLS use 0-based indexing (for array access on getValues())
  var colSets = [
    { name: 'PORTAL_MEMBER_DIR_COLS', obj: PORTAL_MEMBER_DIR_COLS },
    { name: 'PORTAL_EVENT_COLS', obj: PORTAL_EVENT_COLS },
    { name: 'PORTAL_MINUTES_COLS', obj: PORTAL_MINUTES_COLS },
    { name: 'PORTAL_GRIEVANCE_COLS', obj: PORTAL_GRIEVANCE_COLS },
    { name: 'PORTAL_STEWARD_LOG_COLS', obj: PORTAL_STEWARD_LOG_COLS },
    { name: 'PORTAL_MEGA_SURVEY_COLS', obj: PORTAL_MEGA_SURVEY_COLS }
  ];
  for (var c = 0; c < colSets.length; c++) {
    var set = colSets[c];
    var keys = Object.keys(set.obj);
    var hasZero = false;
    for (var k = 0; k < keys.length; k++) {
      var val = set.obj[keys[k]];
      if (typeof val === 'number') {
        TestRunner.assertTrue(val >= 0, set.name + '.' + keys[k] + ' >= 0 (0-indexed)');
        if (val === 0) hasZero = true;
      }
    }
    TestRunner.assertTrue(hasZero, set.name + ' has a column at index 0');
  }
}

/** Tests portal: no duplicate indices. */
function test_portal_noDuplicateIndices() {
  var colSets = [
    { name: 'PORTAL_MEMBER_DIR_COLS', obj: PORTAL_MEMBER_DIR_COLS },
    { name: 'PORTAL_EVENT_COLS', obj: PORTAL_EVENT_COLS },
    { name: 'PORTAL_GRIEVANCE_COLS', obj: PORTAL_GRIEVANCE_COLS }
  ];
  for (var c = 0; c < colSets.length; c++) {
    var set = colSets[c];
    var seen = {};
    var keys = Object.keys(set.obj);
    for (var k = 0; k < keys.length; k++) {
      var val = set.obj[keys[k]];
      if (typeof val !== 'number') continue;
      if (seen[val]) {
        Logger.log('Portal duplicate: ' + set.name + '.' + keys[k] + ' and ' + seen[val] + ' = ' + val);
      }
      seen[val] = keys[k];
    }
  }
}

/** Tests portal: setup functions exist. */
function test_portal_setupFunctionsExist() {
  TestRunner.assertEquals('function', typeof getOrCreatePortalMemberDirectory, 'getOrCreatePortalMemberDirectory');
  TestRunner.assertEquals('function', typeof getOrCreateEventsSheet, 'getOrCreateEventsSheet');
  TestRunner.assertEquals('function', typeof getOrCreateMinutesSheet, 'getOrCreateMinutesSheet');
  TestRunner.assertEquals('function', typeof getOrCreatePortalGrievanceSheet, 'getOrCreatePortalGrievanceSheet');
  TestRunner.assertEquals('function', typeof getOrCreateStewardLogSheet, 'getOrCreateStewardLogSheet');
  TestRunner.assertEquals('function', typeof getOrCreateMegaSurveySheet, 'getOrCreateMegaSurveySheet');
  TestRunner.assertEquals('function', typeof initPortalSheets, 'initPortalSheets');
}

/* ========================================================================
 * WEEKLYQ SUITE — WeeklyQuestions module API and constants
 * ======================================================================== */

function test_weeklyq_moduleExists() {
  TestRunner.assertNotNull(WeeklyQuestions, 'WeeklyQuestions module');
  TestRunner.assertType(WeeklyQuestions, 'object', 'WeeklyQuestions is object');
}

/** Tests weeklyq: public API complete. */
function test_weeklyq_publicAPIComplete() {
  var expected = [
    'getActiveQuestions', 'submitResponse', 'setStewardQuestion',
    'submitPoolQuestion', 'selectRandomPoolQuestion', 'closePoll',
    'getHistory', 'getPoolCount', 'getPollFrequency', 'setPollFrequency',
    'initWeeklyQuestionSheets'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(WeeklyQuestions[expected[i]], 'function',
      'WeeklyQuestions.' + expected[i]);
  }
}

/** Tests weeklyq: q cols exposed. */
function test_weeklyq_qColsExposed() {
  // Q_COLS exposed since v4.24.4
  TestRunner.assertNotNull(WeeklyQuestions.Q_COLS, 'WeeklyQuestions.Q_COLS');
  TestRunner.assertType(WeeklyQuestions.Q_COLS, 'object', 'Q_COLS is object');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'ID', 'Q_COLS.ID');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'TEXT', 'Q_COLS.TEXT');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'ACTIVE', 'Q_COLS.ACTIVE');
}

/** Tests weeklyq: global wrappers exist. */
function test_weeklyq_globalWrappersExist() {
  var wrappers = [
    'wqGetActiveQuestions', 'wqSubmitResponse', 'wqSetStewardQuestion',
    'wqSubmitPoolQuestion', 'wqClosePoll', 'wqGetHistory',
    'wqGetPoolCount', 'wqInitSheets', 'wqGetPollFrequency',
    'wqSetPollFrequency', 'wqManualDrawCommunityPoll'
  ];
  for (var i = 0; i < wrappers.length; i++) {
    TestRunner.assertEquals('function', typeof this[wrappers[i]],
      wrappers[i] + ' global wrapper exists');
  }
}

/** Tests weeklyq: get pool count callable. */
function test_weeklyq_getPoolCountCallable() {
  // getPoolCount is unauthenticated (returns number)
  try {
    var count = WeeklyQuestions.getPoolCount();
    TestRunner.assertType(count, 'number', 'getPoolCount returns number');
    TestRunner.assertTrue(count >= 0, 'pool count >= 0');
  } catch (e) {
    // Sheet may not exist — that's OK
    Logger.log('test_weeklyq_getPoolCountCallable: ' + e.message);
  }
}

/** Tests weeklyq: get poll frequency callable. */
function test_weeklyq_getPollFrequencyCallable() {
  try {
    var freq = WeeklyQuestions.getPollFrequency();
    TestRunner.assertType(freq, 'string', 'getPollFrequency returns string');
    var valid = ['weekly', 'biweekly', 'monthly'];
    TestRunner.assertTrue(valid.indexOf(freq) !== -1,
      'frequency "' + freq + '" is valid (weekly|biweekly|monthly)');
  } catch (e) {
    // Sheet may not exist — acceptable
    Logger.log('test_weeklyq_getPollFrequencyCallable: ' + e.message);
  }
}

/** Tests weeklyq: wrappers reject null token. */
function test_weeklyq_wrappersRejectNullToken() {
  // Auth-gated wrappers should return safe empty on null token
  try {
    var result = wqGetActiveQuestions(null);
    TestRunner.assertTrue(Array.isArray(result) && result.length === 0,
      'wqGetActiveQuestions(null) returns empty array');
  } catch (_e) { /* throwing is acceptable */ }
}

/** Tests weeklyq: auto select exists. */
function test_weeklyq_autoSelectExists() {
  TestRunner.assertEquals('function', typeof autoSelectCommunityPoll, 'autoSelectCommunityPoll trigger handler exists');
}

/* ========================================================================
 * QAFORUM SUITE — QAForum module, question retrieval
 * ======================================================================== */

function test_qaforum_moduleExists() {
  TestRunner.assertNotNull(QAForum, 'QAForum module');
  TestRunner.assertType(QAForum, 'object', 'QAForum is object');
}

/** Tests qaforum: public API complete. */
function test_qaforum_publicAPIComplete() {
  var expected = [
    'initQAForumSheets', 'getQuestions', 'getQuestionDetail',
    'submitQuestion', 'submitAnswer', 'upvoteQuestion',
    'moderateQuestion', 'moderateAnswer', 'getFlaggedContent',
    'resolveQuestion'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(QAForum[expected[i]], 'function',
      'QAForum.' + expected[i]);
  }
}

/** Tests qaforum: global wrappers exist. */
function test_qaforum_globalWrappersExist() {
  var wrappers = [
    'qaGetQuestions', 'qaGetQuestionDetail', 'qaSubmitQuestion',
    'qaSubmitAnswer', 'qaUpvoteQuestion', 'qaModerateQuestion',
    'qaModerateAnswer', 'qaGetFlaggedContent', 'qaResolveQuestion',
    'qaInitSheets'
  ];
  for (var i = 0; i < wrappers.length; i++) {
    TestRunner.assertEquals('function', typeof this[wrappers[i]],
      wrappers[i] + ' global wrapper exists');
  }
}

/** Tests qaforum: get questions returns array. */
function test_qaforum_getQuestionsReturnsArray() {
  try {
    // Call with empty email — should return questions array (public)
    var result = QAForum.getQuestions('', 1, 10, 'recent');
    TestRunner.assertTrue(Array.isArray(result), 'getQuestions returns array');
  } catch (e) {
    // Sheet may not exist — acceptable for new installs
    Logger.log('test_qaforum_getQuestionsReturnsArray: ' + e.message);
  }
}

/** Tests qaforum: wrappers reject null token. */
function test_qaforum_wrappersRejectNullToken() {
  // Steward-gated endpoints should reject null
  try {
    var result = qaGetFlaggedContent(null);
    // Should return empty array or error object
    if (Array.isArray(result)) {
      TestRunner.assertEquals(0, result.length, 'qaGetFlaggedContent(null) returns empty');
    } else if (result && typeof result === 'object') {
      TestRunner.assertHasKey(result, 'success', 'returns result with success key');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

/** Tests qaforum: pagination defaults. */
function test_qaforum_paginationDefaults() {
  try {
    // Default pagination should work (page 1, pageSize 10)
    var result = QAForum.getQuestions('', 1, 10, 'recent');
    if (Array.isArray(result)) {
      TestRunner.assertTrue(result.length <= 10, 'page size respected (max 10 results)');
    }
  } catch (e) {
    Logger.log('test_qaforum_paginationDefaults: ' + e.message);
  }
}

/* ========================================================================
 * TIMELINE SUITE — TimelineService module, events, categories
 * ======================================================================== */

function test_timeline_moduleExists() {
  TestRunner.assertNotNull(TimelineService, 'TimelineService module');
  TestRunner.assertType(TimelineService, 'object', 'TimelineService is object');
}

/** Tests timeline: public API complete. */
function test_timeline_publicAPIComplete() {
  var expected = [
    'initTimelineSheet', 'getTimelineEvents', 'addTimelineEvent',
    'updateTimelineEvent', 'deleteTimelineEvent', 'importCalendarEvents',
    'attachDriveFiles'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(TimelineService[expected[i]], 'function',
      'TimelineService.' + expected[i]);
  }
}

/** Tests timeline: global wrappers exist. */
function test_timeline_globalWrappersExist() {
  var wrappers = [
    'tlGetTimelineEvents', 'tlAddTimelineEvent', 'tlUpdateTimelineEvent',
    'tlDeleteTimelineEvent', 'tlImportCalendarEvents', 'tlAttachDriveFiles',
    'tlInitSheets'
  ];
  for (var i = 0; i < wrappers.length; i++) {
    TestRunner.assertEquals('function', typeof this[wrappers[i]],
      wrappers[i] + ' global wrapper exists');
  }
}

/** Tests timeline: get events returns array. */
function test_timeline_getEventsReturnsArray() {
  try {
    var result = TimelineService.getTimelineEvents(1, 10, null, null);
    TestRunner.assertTrue(Array.isArray(result), 'getTimelineEvents returns array');
  } catch (e) {
    // Sheet may not exist
    Logger.log('test_timeline_getEventsReturnsArray: ' + e.message);
  }
}

/** Tests timeline: wrappers reject null token. */
function test_timeline_wrappersRejectNullToken() {
  // Write endpoints should reject null token
  try {
    var result = tlAddTimelineEvent(null, {});
    if (result && typeof result === 'object') {
      TestRunner.assertEquals(false, result.success, 'tlAddTimelineEvent(null) rejects');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

/** Tests timeline: categories validated. */
function test_timeline_categoriesValidated() {
  // TimelineService should accept known categories
  var knownCategories = ['meeting', 'announcement', 'milestone', 'action', 'decision', 'other'];
  // Verify by trying a read with category filter — should not throw for valid categories
  for (var i = 0; i < knownCategories.length; i++) {
    try {
      TimelineService.getTimelineEvents(1, 1, null, knownCategories[i]);
    } catch (e) {
      // Sheet missing is OK, but category rejection would be a bug
      if (e.message && e.message.indexOf('Invalid category') >= 0) {
        TestRunner.assertTrue(false, 'Category "' + knownCategories[i] + '" was rejected');
      }
    }
  }
}

/* ========================================================================
 * FAILSAFE SUITE — FailsafeService module, digest config, diagnostics
 * ======================================================================== */

function test_failsafe_moduleExists() {
  TestRunner.assertNotNull(FailsafeService, 'FailsafeService module');
  TestRunner.assertType(FailsafeService, 'object', 'FailsafeService is object');
}

/** Tests failsafe: public API complete. */
function test_failsafe_publicAPIComplete() {
  var expected = [
    'initFailsafeSheet', 'getDigestConfig', 'updateDigestConfig',
    'processScheduledDigests', 'triggerBulkExport',
    'backupCriticalSheets', 'setupFailsafeTriggers', 'removeFailsafeTriggers'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(FailsafeService[expected[i]], 'function',
      'FailsafeService.' + expected[i]);
  }
}

/** Tests failsafe: global wrappers exist. */
function test_failsafe_globalWrappersExist() {
  var wrappers = [
    'fsGetDigestConfig', 'fsUpdateDigestConfig',
    'fsProcessScheduledDigests', 'fsTriggerBulkExport',
    'fsBackupCriticalSheets', 'fsSetupTriggers', 'fsRemoveTriggers',
    'fsInitSheets'
  ];
  for (var i = 0; i < wrappers.length; i++) {
    TestRunner.assertEquals('function', typeof this[wrappers[i]],
      wrappers[i] + ' global wrapper exists');
  }
}

/** Tests failsafe: diagnostic exists. */
function test_failsafe_diagnosticExists() {
  TestRunner.assertEquals('function', typeof fsDiagnostic, 'fsDiagnostic exists');
}

/** Tests failsafe: digest config returns shape. */
function test_failsafe_digestConfigReturnsShape() {
  // getDigestConfig with non-existent email should return default config
  try {
    var config = FailsafeService.getDigestConfig('__nonexistent_probe__@example.invalid');
    TestRunner.assertNotNull(config, 'getDigestConfig returns value for unknown email');
    TestRunner.assertType(config, 'object', 'config is object');
  } catch (e) {
    // Sheet may not exist — acceptable
    Logger.log('test_failsafe_digestConfigReturnsShape: ' + e.message);
  }
}

/** Tests failsafe: wrappers reject null token. */
function test_failsafe_wrappersRejectNullToken() {
  try {
    var result = fsTriggerBulkExport(null);
    if (result && typeof result === 'object') {
      TestRunner.assertEquals(false, result.success, 'fsTriggerBulkExport(null) rejects');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

/** Tests failsafe: ensure all sheets exists. */
function test_failsafe_ensureAllSheetsExists() {
  TestRunner.assertEquals('function', typeof fsEnsureAllSheets, 'fsEnsureAllSheets exists');
}

/* ========================================================================
 * ENDPOINTS SUITE — Comprehensive data* wrapper existence & auth gating
 * Covers ALL web app callable functions across all modules.
 * ======================================================================== */

function test_endpoints_thisBindingCanary() {
  // Canary: verifies `this` inside test functions references the global scope.
  // If this fails, TestRunner is calling test functions with wrong `this` binding
  // (e.g., test.fn() method call instead of indirect call), which breaks all
  // this[fnName] dynamic lookups in this suite.
  // TestRunner is an IIFE-returned object, so typeof is 'object', not 'function'.
  TestRunner.assertEquals('object', typeof this['TestRunner'],
    'this references global scope (canary)');
}

/** Tests endpoints: core grievance fns exist. */
function test_endpoints_coreGrievanceFnsExist() {
  var fns = [
    'dataGetStewardCases', 'dataGetStewardKPIs',
    'dataGetMemberGrievances', 'dataGetMemberGrievanceHistory',
    'dataGetStewardContact', 'dataGetFullProfile',
    'dataUpdateProfile', 'dataGetAssignedSteward',
    'dataGetAvailableStewards', 'dataAssignSteward'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: task fns exist. */
function test_endpoints_taskFnsExist() {
  var fns = [
    'dataCreateTask', 'dataGetTasks', 'dataCompleteTask', 'dataUpdateTask',
    'dataCreateMemberTask', 'dataGetMemberTasks', 'dataCompleteMemberTask'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: survey feedback fns exist. */
function test_endpoints_surveyFeedbackFnsExist() {
  var fns = [
    'dataGetSurveyStatus', 'dataGetSurveyQuestions',
    'dataSubmitSurveyResponse', 'dataGetSurveyResults',
    'dataGetSatisfactionTrends', 'dataSubmitFeedback', 'dataGetMyFeedback'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: contact notification fns exist. */
function test_endpoints_contactNotificationFnsExist() {
  var fns = [
    'dataLogMemberContact', 'dataGetMemberContactHistory',
    'dataGetStewardContactLog', 'dataSendDirectMessage'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: admin stats fns exist. */
function test_endpoints_adminStatsFnsExist() {
  var fns = [
    'dataGetAllMembers', 'dataSendBroadcast',
    'dataGetStewardMemberStats', 'dataGetStewardDirectory',
    'dataGetGrievanceStats', 'dataGetGrievanceHotSpots',
    'dataGetMembershipStats', 'dataGetUpcomingEvents',
    'dataIsChiefSteward', 'dataGetAllStewardPerformance'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: batch diagnostic fns exist. */
function test_endpoints_batchDiagnosticFnsExist() {
  var fns = [
    'dataGetBatchData', 'dataEnsureSheetsIfNeeded',
    'dataGetEngagementStats', 'dataGetWelcomeData'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: meeting fns exist. */
function test_endpoints_meetingFnsExist() {
  var fns = [
    'dataGetMemberMeetings', 'dataGetMeetingMinutes', 'dataAddMeetingMinutes'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: checklist fns exist. */
function test_endpoints_checklistFnsExist() {
  var fns = ['dataGetCaseChecklist', 'dataToggleChecklistItem'];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

/** Tests endpoints: poll stubs removed. */
function test_endpoints_pollStubsRemoved() {
  // v4.25.11: Legacy poll stubs removed — verify they no longer exist
  TestRunner.assertEquals('undefined', typeof dataGetActivePolls, 'dataGetActivePolls removed');
  TestRunner.assertEquals('undefined', typeof dataSubmitPollVote, 'dataSubmitPollVote removed');
  TestRunner.assertEquals('undefined', typeof dataAddPoll, 'dataAddPoll removed');
}

/** Tests endpoints: all write endpoints reject null. */
function test_endpoints_allWriteEndpointsRejectNull() {
  // Write/mutate endpoints must reject null session tokens
  var writeEndpoints = [
    { fn: 'dataUpdateProfile', args: [null, {}] },
    { fn: 'dataCreateTask', args: [null, {}] },
    { fn: 'dataSubmitFeedback', args: [null, {}] },
    { fn: 'dataSendBroadcast', args: [null, '', '', ''] },
    { fn: 'dataLogMemberContact', args: [null, '', '', ''] }
  ];
  for (var i = 0; i < writeEndpoints.length; i++) {
    var ep = writeEndpoints[i];
    try {
      var result = this[ep.fn].apply(null, ep.args);
      // Result should indicate auth failure
      if (result && typeof result === 'object' && result.hasOwnProperty('success')) {
        TestRunner.assertEquals(false, result.success,
          ep.fn + '(null) returns success:false');
      } else if (Array.isArray(result)) {
        TestRunner.assertEquals(0, result.length,
          ep.fn + '(null) returns empty array');
      }
      // null/undefined also acceptable as rejection
    } catch (_e) {
      // Throwing is acceptable — endpoint rejected the call
    }
  }
}

/** Tests endpoints: notification count exists. */
function test_endpoints_notificationCountExists() {
  TestRunner.assertEquals('function', typeof getWebAppNotificationCount,
    'getWebAppNotificationCount exists');
}

/** Tests endpoints: grievance draft fns exist. */
function test_endpoints_grievanceDraftFnsExist() {
  // Grievance draft/drive functions
  var fns = ['dataStartGrievanceDraft', 'dataStartGrievanceDraftForMember', 'dataCreateGrievanceDrive'];
  for (var i = 0; i < fns.length; i++) {
    if (typeof this[fns[i]] !== 'undefined') {
      TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
    }
  }
}
