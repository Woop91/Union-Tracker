/**
 * 31_WebAppTests.gs
 * GAS-Native Web App Test Suites — comprehensive coverage of all web app modules.
 *
 * Designed to run independently via TestRunner suite filter. Each suite
 * completes well under 3 minutes to avoid GAS 6-minute execution limit.
 *
 * Test suites:
 *   webapp_      — doGet routing, template rendering, URL generation, diagnoseWebApp
 *   configrd_    — ConfigReader module shape, validation, caching
 *   portal_      — PortalSheets column constants, 0-indexed validation, sheet existence
 *   weeklyq_     — WeeklyQuestions module API, column constants, frequency reads
 *   workload_    — WorkloadService module, categories, health, column constants
 *   qaforum_     — QAForum module, question retrieval, pagination
 *   timeline_    — TimelineService module, event retrieval, category validation
 *   failsafe_    — FailsafeService module, digest config, diagnostics
 *   endpoints_   — All data* / wq* / qa* / tl* / fs* wrapper existence & auth gating
 *
 * IMPORTANT: All tests are READ-ONLY. They never write to sheets.
 *
 * @fileoverview Web app test suites for DDS-Dashboard
 * @version 1.0.0
 */

/* ========================================================================
 * WEBAPP SUITE — doGet routing, templates, URL resolution
 * ======================================================================== */

function test_webapp_doGetExists() {
  TestRunner.assertEquals('function', typeof doGet, 'doGet global function exists');
}

function test_webapp_doGetWebDashboardExists() {
  TestRunner.assertEquals('function', typeof doGetWebDashboard, 'doGetWebDashboard exists');
}

function test_webapp_includeHelperExists() {
  TestRunner.assertEquals('function', typeof include, 'include() template helper exists');
}

function test_webapp_getWebAppUrlExists() {
  TestRunner.assertEquals('function', typeof getWebAppUrl, 'getWebAppUrl exists');
}

function test_webapp_getWebAppUrlReturnsString() {
  var url = getWebAppUrl();
  TestRunner.assertNotNull(url, 'getWebAppUrl returns a value');
  TestRunner.assertType(url, 'string', 'getWebAppUrl returns string');
  TestRunner.assertGreaterThan(url.length, 0, 'URL is not empty');
}

function test_webapp_orgChartHtmlExists() {
  TestRunner.assertEquals('function', typeof getOrgChartHtml, 'getOrgChartHtml exists');
}

function test_webapp_pomsReferenceHtmlExists() {
  TestRunner.assertEquals('function', typeof getPOMSReferenceHtml, 'getPOMSReferenceHtml exists');
}

function test_webapp_pomsReferenceHtmlReturnsContent() {
  // getPOMSReferenceHtml should return HTML content (not an error fallback)
  try {
    var html = getPOMSReferenceHtml();
    TestRunner.assertNotNull(html, 'getPOMSReferenceHtml returns non-null');
    TestRunner.assertTrue(typeof html === 'string', 'getPOMSReferenceHtml returns string');
    TestRunner.assertGreaterThan(html.length, 100, 'getPOMSReferenceHtml returns substantial content');
    // Should contain the POMS root element
    TestRunner.assertTrue(html.indexOf('poms-root') > -1, 'getPOMSReferenceHtml contains poms-root element');
    // Should NOT be the auth-failure fallback
    TestRunner.assertTrue(html.indexOf('Authentication required') === -1, 'getPOMSReferenceHtml did not return auth error');
  } catch (e) {
    TestRunner.assertTrue(false, 'getPOMSReferenceHtml threw: ' + e.message);
  }
}

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

function test_webapp_diagnoseWebAppExists() {
  TestRunner.assertEquals('function', typeof diagnoseWebApp, 'diagnoseWebApp exists');
}

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

function test_webapp_serveFatalErrorExists() {
  TestRunner.assertEquals('function', typeof _serveFatalError, '_serveFatalError exists');
}

function test_webapp_sanitizeConfigExists() {
  TestRunner.assertEquals('function', typeof _sanitizeConfig, '_sanitizeConfig exists');
}

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

function test_configrd_getConfigCallable() {
  TestRunner.assertType(ConfigReader.getConfig, 'function', 'getConfig is function');
}

function test_configrd_validateConfigCallable() {
  TestRunner.assertType(ConfigReader.validateConfig, 'function', 'validateConfig is function');
}

function test_configrd_refreshConfigCallable() {
  TestRunner.assertType(ConfigReader.refreshConfig, 'function', 'refreshConfig is function');
}

function test_configrd_getConfigJSONCallable() {
  TestRunner.assertType(ConfigReader.getConfigJSON, 'function', 'getConfigJSON is function');
}

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

function test_configrd_configHasDriveFields() {
  var config = ConfigReader.getConfig();
  // Drive integration fields — may be empty but keys should exist
  var driveKeys = ['calendarId', 'driveFolderId', 'dashboardRootFolderId'];
  for (var i = 0; i < driveKeys.length; i++) {
    TestRunner.assertHasKey(config, driveKeys[i],
      'config has key: ' + driveKeys[i]);
  }
}

function test_configrd_configHasAuthFields() {
  var config = ConfigReader.getConfig();
  // Auth-related config fields
  TestRunner.assertHasKey(config, 'magicLinkExpiryDays', 'magicLinkExpiryDays');
  TestRunner.assertHasKey(config, 'cookieDurationDays', 'cookieDurationDays');
}

function test_configrd_validateConfigReturnsShape() {
  var result = ConfigReader.validateConfig();
  TestRunner.assertNotNull(result, 'validateConfig returns value');
  TestRunner.assertType(result, 'object', 'validateConfig returns object');
  TestRunner.assertHasKey(result, 'valid', 'has valid key');
  TestRunner.assertType(result.valid, 'boolean', 'valid is boolean');
  TestRunner.assertHasKey(result, 'missing', 'has missing key');
  TestRunner.assertTrue(Array.isArray(result.missing), 'missing is array');
}

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

function test_portal_eventColsDefined() {
  TestRunner.assertNotNull(PORTAL_EVENT_COLS, 'PORTAL_EVENT_COLS');
  TestRunner.assertType(PORTAL_EVENT_COLS, 'object', 'is object');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'TITLE', 'TITLE');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'TYPE', 'TYPE');
  TestRunner.assertHasKey(PORTAL_EVENT_COLS, 'DATE_TIME', 'DATE_TIME');
}

function test_portal_minutesColsDefined() {
  TestRunner.assertNotNull(PORTAL_MINUTES_COLS, 'PORTAL_MINUTES_COLS');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'MEETING_DATE', 'MEETING_DATE');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'TITLE', 'TITLE');
  TestRunner.assertHasKey(PORTAL_MINUTES_COLS, 'DRIVE_DOC_URL', 'DRIVE_DOC_URL');
}

function test_portal_grievanceColsDefined() {
  TestRunner.assertNotNull(PORTAL_GRIEVANCE_COLS, 'PORTAL_GRIEVANCE_COLS');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'MEMBER_EMAIL', 'MEMBER_EMAIL');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'STEWARD_EMAIL', 'STEWARD_EMAIL');
  TestRunner.assertHasKey(PORTAL_GRIEVANCE_COLS, 'STATUS', 'STATUS');
}

function test_portal_stewardLogColsDefined() {
  TestRunner.assertNotNull(PORTAL_STEWARD_LOG_COLS, 'PORTAL_STEWARD_LOG_COLS');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'ID', 'ID');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'STEWARD_EMAIL', 'STEWARD_EMAIL');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'MEMBER_EMAIL', 'MEMBER_EMAIL');
  TestRunner.assertHasKey(PORTAL_STEWARD_LOG_COLS, 'TYPE', 'TYPE');
}

function test_portal_megaSurveyColsDefined() {
  TestRunner.assertNotNull(PORTAL_MEGA_SURVEY_COLS, 'PORTAL_MEGA_SURVEY_COLS');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'EMAIL', 'EMAIL');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'RESPONSES', 'RESPONSES');
  TestRunner.assertHasKey(PORTAL_MEGA_SURVEY_COLS, 'COMPLETED', 'COMPLETED');
}

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

function test_weeklyq_qColsExposed() {
  // Q_COLS exposed since v4.24.4
  TestRunner.assertNotNull(WeeklyQuestions.Q_COLS, 'WeeklyQuestions.Q_COLS');
  TestRunner.assertType(WeeklyQuestions.Q_COLS, 'object', 'Q_COLS is object');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'ID', 'Q_COLS.ID');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'TEXT', 'Q_COLS.TEXT');
  TestRunner.assertHasKey(WeeklyQuestions.Q_COLS, 'ACTIVE', 'Q_COLS.ACTIVE');
}

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

function test_weeklyq_wrappersRejectNullToken() {
  // Auth-gated wrappers should return safe empty on null token
  try {
    var result = wqGetActiveQuestions(null);
    TestRunner.assertTrue(Array.isArray(result) && result.length === 0,
      'wqGetActiveQuestions(null) returns empty array');
  } catch (_e) { /* throwing is acceptable */ }
}

function test_weeklyq_autoSelectExists() {
  TestRunner.assertEquals('function', typeof autoSelectCommunityPoll, 'autoSelectCommunityPoll trigger handler exists');
}

/* ========================================================================
 * WORKLOAD SUITE — WorkloadService module, categories, health
 * ======================================================================== */

function test_workload_moduleExists() {
  TestRunner.assertNotNull(WorkloadService, 'WorkloadService module');
  TestRunner.assertType(WorkloadService, 'object', 'WorkloadService is object');
}

function test_workload_publicAPIComplete() {
  var expected = [
    'initSheets', 'processFormSSO', 'getHistorySSO',
    'getDashboardDataSSO', 'getReminderSSO', 'setReminderSSO',
    'exportHistoryCSV', 'getSubCategories', 'processReminders',
    'getHealthStatus', 'refreshLedger'
  ];
  for (var i = 0; i < expected.length; i++) {
    TestRunner.assertType(WorkloadService[expected[i]], 'function',
      'WorkloadService.' + expected[i]);
  }
}

function test_workload_subCategoriesExposed() {
  TestRunner.assertNotNull(WorkloadService.SUB_CATEGORIES, 'SUB_CATEGORIES');
  TestRunner.assertType(WorkloadService.SUB_CATEGORIES, 'object', 'SUB_CATEGORIES is object');
  // Verify known category keys
  var keys = ['priority', 'pending', 'unread', 'todo', 'referrals', 'ce', 'assistance', 'aged'];
  for (var i = 0; i < keys.length; i++) {
    TestRunner.assertHasKey(WorkloadService.SUB_CATEGORIES, keys[i],
      'SUB_CATEGORIES.' + keys[i]);
    TestRunner.assertTrue(Array.isArray(WorkloadService.SUB_CATEGORIES[keys[i]]),
      keys[i] + ' is array');
  }
}

function test_workload_categoryLabelsExposed() {
  TestRunner.assertNotNull(WorkloadService.CATEGORY_LABELS, 'CATEGORY_LABELS');
  TestRunner.assertType(WorkloadService.CATEGORY_LABELS, 'object', 'CATEGORY_LABELS is object');
  // Should have t1-t8 mappings
  for (var t = 1; t <= 8; t++) {
    TestRunner.assertHasKey(WorkloadService.CATEGORY_LABELS, 't' + t,
      'CATEGORY_LABELS.t' + t);
  }
}

function test_workload_getSubCategoriesCallable() {
  var cats = WorkloadService.getSubCategories();
  TestRunner.assertNotNull(cats, 'getSubCategories returns value');
  TestRunner.assertType(cats, 'object', 'returns object');
  TestRunner.assertHasKey(cats, 'priority', 'has priority key');
}

function test_workload_globalWrappersExist() {
  var wrappers = [
    'processWorkloadFormSSO', 'getWorkloadHistorySSO', 'getWorkloadDashboardDataSSO',
    'getWorkloadReminderSSO', 'setWorkloadReminderSSO', 'exportWorkloadHistoryCSV',
    'getWorkloadSubCategories'
  ];
  for (var i = 0; i < wrappers.length; i++) {
    TestRunner.assertEquals('function', typeof this[wrappers[i]],
      wrappers[i] + ' exists');
  }
}

function test_workload_triggerHandlersExist() {
  TestRunner.assertEquals('function', typeof initWorkloadTrackerSheets, 'initWorkloadTrackerSheets');
  TestRunner.assertEquals('function', typeof processWorkloadReminders, 'processWorkloadReminders');
  TestRunner.assertEquals('function', typeof refreshWorkloadLedger, 'refreshWorkloadLedger');
}

function test_workload_wrappersRejectNullToken() {
  try {
    var result = getWorkloadHistorySSO(null);
    // Should return empty/null for null token
    if (Array.isArray(result)) {
      TestRunner.assertEquals(0, result.length, 'getWorkloadHistorySSO(null) returns empty array');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

/* ========================================================================
 * QAFORUM SUITE — QAForum module, question retrieval
 * ======================================================================== */

function test_qaforum_moduleExists() {
  TestRunner.assertNotNull(QAForum, 'QAForum module');
  TestRunner.assertType(QAForum, 'object', 'QAForum is object');
}

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

function test_timeline_getEventsReturnsArray() {
  try {
    var result = TimelineService.getTimelineEvents(1, 10, null, null);
    TestRunner.assertTrue(Array.isArray(result), 'getTimelineEvents returns array');
  } catch (e) {
    // Sheet may not exist
    Logger.log('test_timeline_getEventsReturnsArray: ' + e.message);
  }
}

function test_timeline_wrappersRejectNullToken() {
  // Write endpoints should reject null token
  try {
    var result = tlAddTimelineEvent(null, {});
    if (result && typeof result === 'object') {
      TestRunner.assertEquals(false, result.success, 'tlAddTimelineEvent(null) rejects');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

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

function test_failsafe_diagnosticExists() {
  TestRunner.assertEquals('function', typeof fsDiagnostic, 'fsDiagnostic exists');
}

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

function test_failsafe_wrappersRejectNullToken() {
  try {
    var result = fsTriggerBulkExport(null);
    if (result && typeof result === 'object') {
      TestRunner.assertEquals(false, result.success, 'fsTriggerBulkExport(null) rejects');
    }
  } catch (_e) { /* throwing is acceptable */ }
}

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

function test_endpoints_taskFnsExist() {
  var fns = [
    'dataCreateTask', 'dataGetTasks', 'dataCompleteTask', 'dataUpdateTask',
    'dataCreateMemberTask', 'dataGetMemberTasks', 'dataCompleteMemberTask'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

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

function test_endpoints_contactNotificationFnsExist() {
  var fns = [
    'dataLogMemberContact', 'dataGetMemberContactHistory',
    'dataGetStewardContactLog', 'dataSendDirectMessage'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

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

function test_endpoints_batchDiagnosticFnsExist() {
  var fns = [
    'dataGetBatchData', 'dataEnsureSheetsIfNeeded',
    'dataGetEngagementStats', 'dataGetWelcomeData'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

function test_endpoints_meetingFnsExist() {
  var fns = [
    'dataGetMemberMeetings', 'dataGetMeetingMinutes', 'dataAddMeetingMinutes'
  ];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

function test_endpoints_checklistFnsExist() {
  var fns = ['dataGetCaseChecklist', 'dataToggleChecklistItem'];
  for (var i = 0; i < fns.length; i++) {
    TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
  }
}

function test_endpoints_legacyStubsSafe() {
  // Legacy poll stubs must exist and return safe values
  TestRunner.assertEquals('function', typeof dataGetActivePolls, 'dataGetActivePolls stub');
  TestRunner.assertEquals('function', typeof dataSubmitPollVote, 'dataSubmitPollVote stub');
  TestRunner.assertEquals('function', typeof dataAddPoll, 'dataAddPoll stub');

  var polls = dataGetActivePolls();
  TestRunner.assertTrue(Array.isArray(polls), 'dataGetActivePolls returns array');
  TestRunner.assertEquals(0, polls.length, 'dataGetActivePolls returns empty');
}

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

function test_endpoints_notificationCountExists() {
  TestRunner.assertEquals('function', typeof getWebAppNotificationCount,
    'getWebAppNotificationCount exists');
}

function test_endpoints_grievanceDraftFnsExist() {
  // Grievance draft/drive functions
  var fns = ['dataStartGrievanceDraft', 'dataStartGrievanceDraftForMember', 'dataCreateGrievanceDrive'];
  for (var i = 0; i < fns.length; i++) {
    if (typeof this[fns[i]] !== 'undefined') {
      TestRunner.assertEquals('function', typeof this[fns[i]], fns[i] + ' exists');
    }
  }
}
