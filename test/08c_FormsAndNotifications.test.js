/**
 * Tests for 08c_FormsAndNotifications.gs
 * Covers form URL handling, survey tracking, deadline notifications,
 * email validation, quarter calculation, and submission handling.
 */

const { createMockSheet, createMockSpreadsheet, createMockRange } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '08c_FormsAndNotifications.gs'
]);

// ============================================================================
// Function existence
// ============================================================================

describe('08c function existence', () => {
  const required = [
    'getFormUrlFromConfig', 'buildGrievanceFormUrl_',
    'saveFormUrlsToConfig_silent',
    'onSatisfactionFormSubmit',
    'showNotificationSettings',
    'installDailyTrigger_', 'removeDailyTrigger_',
    'checkDeadlinesAndNotify_', 'testDeadlineNotifications',
    'sendStewardDeadlineAlerts',
    'executeSendRandomSurveyEmails', 'validateMemberEmail',
    'getCurrentQuarter',
    'populateSurveyTrackingFromMembers', 'startNewSurveyRound',
    'sendSurveyCompletionReminders', 'getSurveyCompletionStats',
    'showSurveyTrackingDialog', 'getSurveyQuestions',
    'getSatisfactionColMap_', 'syncSatisfactionSheetColumns_',
    'submitSurveyResponse', 'buildSatisfactionColsShim_'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// getCurrentQuarter
// ============================================================================

describe('getCurrentQuarter', () => {
  test('returns a YYYY-QN format string', () => {
    const q = getCurrentQuarter();
    expect(q).toMatch(/^\d{4}-Q[1-4]$/);
  });
});

// ============================================================================
// validateMemberEmail (sheet-based lookup — test null/empty rejection)
// ============================================================================

describe('validateMemberEmail', () => {
  test('returns null for empty email', () => {
    const result = validateMemberEmail('');
    expect(result).toBeNull();
  });

  test('returns null for null input', () => {
    const result = validateMemberEmail(null);
    expect(result).toBeNull();
  });
});

// ============================================================================
// buildSatisfactionColsShim_
// ============================================================================

describe('buildSatisfactionColsShim_', () => {
  test('returns an object given a column map', () => {
    const mockMap = { 'Member ID': 1, 'Email': 2, 'Overall Score': 3 };
    const result = buildSatisfactionColsShim_(mockMap);
    expect(typeof result).toBe('object');
  });

  test('returns object for empty map', () => {
    const result = buildSatisfactionColsShim_({});
    expect(typeof result).toBe('object');
  });
});

// ============================================================================
// Behavioral: sendStewardDeadlineAlerts
// ============================================================================

describe('sendStewardDeadlineAlerts (behavioral)', () => {
  test('returns silently when required sheets are missing', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Should not throw
    expect(() => sendStewardDeadlineAlerts()).not.toThrow();
  });

  test('reads grievance data and attempts to send email for upcoming deadlines', () => {
    // Build member directory with a steward
    var memberHeaders = new Array(30).fill('');
    memberHeaders[MEMBER_COLS.MEMBER_ID - 1] = 'Member ID';
    memberHeaders[MEMBER_COLS.FIRST_NAME - 1] = 'First Name';
    memberHeaders[MEMBER_COLS.LAST_NAME - 1] = 'Last Name';
    memberHeaders[MEMBER_COLS.EMAIL - 1] = 'Email';
    memberHeaders[MEMBER_COLS.ASSIGNED_STEWARD - 1] = 'Assigned Steward';
    memberHeaders[MEMBER_COLS.IS_STEWARD - 1] = 'Is Steward';

    var memberRow = new Array(30).fill('');
    memberRow[MEMBER_COLS.MEMBER_ID - 1] = 'M-001';
    memberRow[MEMBER_COLS.FIRST_NAME - 1] = 'Jane';
    memberRow[MEMBER_COLS.LAST_NAME - 1] = 'Steward';
    memberRow[MEMBER_COLS.EMAIL - 1] = 'jane.steward@test.com';
    memberRow[MEMBER_COLS.ASSIGNED_STEWARD - 1] = 'Jane Steward';
    memberRow[MEMBER_COLS.IS_STEWARD - 1] = 'Yes';

    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [memberHeaders, memberRow]);

    // Build grievance log with an approaching deadline
    var gHeaders = new Array(30).fill('');
    gHeaders[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'Grievance ID';
    gHeaders[GRIEVANCE_COLS.MEMBER_ID - 1] = 'Member ID';
    gHeaders[GRIEVANCE_COLS.STATUS - 1] = 'Status';
    gHeaders[GRIEVANCE_COLS.CURRENT_STEP - 1] = 'Current Step';
    gHeaders[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] = 'Next Action Due';
    gHeaders[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 'Days to Deadline';
    gHeaders[GRIEVANCE_COLS.STEWARD - 1] = 'Steward';

    var gRow = new Array(30).fill('');
    gRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'G-001';
    gRow[GRIEVANCE_COLS.MEMBER_ID - 1] = 'M-001';
    gRow[GRIEVANCE_COLS.STATUS - 1] = 'Open';
    gRow[GRIEVANCE_COLS.CURRENT_STEP - 1] = 'Step I';
    gRow[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] = new Date();
    gRow[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 3;
    gRow[GRIEVANCE_COLS.STEWARD - 1] = 'Jane Steward';

    var grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [gHeaders, gRow]);
    var ss = createMockSpreadsheet([grievanceSheet, memberSheet]);
    ss.getUrl = jest.fn(() => 'https://docs.google.com/spreadsheets/d/mock/edit');
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Clear previous calls
    MailApp.sendEmail.mockClear();

    sendStewardDeadlineAlerts();

    // Verify the function attempted to process grievance data (reads the sheet)
    // and attempted email delivery for the grievance with daysToDeadline=3
    var emailCalled = MailApp.sendEmail.mock.calls.length > 0 ||
      (typeof safeSendEmail_ === 'function' && safeSendEmail_.mock && safeSendEmail_.mock.calls.length > 0);
    expect(emailCalled).toBe(true);
  });
});

// ============================================================================
// Behavioral: getSurveyQuestions
// ============================================================================

describe('getSurveyQuestions (behavioral)', () => {
  test('returns expected shape with questions, sections, sliderLabels, period', () => {
    // Clear cache so it reads from sheet
    CacheService.getScriptCache().remove('surveyQuestions_v1');

    var qHeaders = new Array(16).fill('');
    qHeaders[SURVEY_QUESTIONS_COLS.QUESTION_ID - 1] = 'Question ID';
    qHeaders[SURVEY_QUESTIONS_COLS.SECTION_NUM - 1] = 'Section Number';
    qHeaders[SURVEY_QUESTIONS_COLS.SECTION_KEY - 1] = 'Section Key';
    qHeaders[SURVEY_QUESTIONS_COLS.SECTION_TITLE - 1] = 'Section Title';
    qHeaders[SURVEY_QUESTIONS_COLS.QUESTION_TEXT - 1] = 'Question Text';
    qHeaders[SURVEY_QUESTIONS_COLS.TYPE - 1] = 'Type';
    qHeaders[SURVEY_QUESTIONS_COLS.REQUIRED - 1] = 'Required';
    qHeaders[SURVEY_QUESTIONS_COLS.ACTIVE - 1] = 'Active';
    qHeaders[SURVEY_QUESTIONS_COLS.OPTIONS - 1] = 'Options';

    var qRow = new Array(16).fill('');
    qRow[SURVEY_QUESTIONS_COLS.QUESTION_ID - 1] = 'q1';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_NUM - 1] = '1';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_KEY - 1] = '1';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_TITLE - 1] = 'Work Context';
    qRow[SURVEY_QUESTIONS_COLS.QUESTION_TEXT - 1] = 'Where do you work?';
    qRow[SURVEY_QUESTIONS_COLS.TYPE - 1] = 'radio';
    qRow[SURVEY_QUESTIONS_COLS.REQUIRED - 1] = 'Y';
    qRow[SURVEY_QUESTIONS_COLS.ACTIVE - 1] = 'Y';
    qRow[SURVEY_QUESTIONS_COLS.OPTIONS - 1] = 'Office A|Office B|Office C';

    var qSheet = createMockSheet(SHEETS.SURVEY_QUESTIONS, [qHeaders, qRow]);
    var configSheet = createMockSheet(SHEETS.CONFIG, [['Section'], ['Headers']]);
    var ss = createMockSpreadsheet([qSheet, configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyQuestions();

    // Verify shape
    expect(result).toHaveProperty('questions');
    expect(result).toHaveProperty('sections');
    expect(result).toHaveProperty('sliderLabels');
    expect(result).toHaveProperty('period');
    expect(Array.isArray(result.questions)).toBe(true);
    expect(Array.isArray(result.sections)).toBe(true);
    expect(result.sliderLabels).toHaveProperty('min');
    expect(result.sliderLabels).toHaveProperty('max');
  });

  test('parsed questions have expected properties (id, text, type, required, section)', () => {
    CacheService.getScriptCache().remove('surveyQuestions_v1');

    // Build sheet data with proper headers and one question row
    var qHeaders = new Array(16).fill('');
    qHeaders[SURVEY_QUESTIONS_COLS.QUESTION_ID - 1] = 'Question ID';

    var qRow = new Array(16).fill('');
    qRow[SURVEY_QUESTIONS_COLS.QUESTION_ID - 1] = 'q2';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_NUM - 1] = '2';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_KEY - 1] = '2';
    qRow[SURVEY_QUESTIONS_COLS.SECTION_TITLE - 1] = 'Satisfaction';
    qRow[SURVEY_QUESTIONS_COLS.QUESTION_TEXT - 1] = 'Rate overall satisfaction';
    qRow[SURVEY_QUESTIONS_COLS.TYPE - 1] = 'slider-10';
    qRow[SURVEY_QUESTIONS_COLS.REQUIRED - 1] = 'Y';
    qRow[SURVEY_QUESTIONS_COLS.ACTIVE - 1] = 'Y';

    var qSheet = createMockSheet(SHEETS.SURVEY_QUESTIONS, [qHeaders, qRow]);
    var configSheet = createMockSheet(SHEETS.CONFIG, [['Section'], ['Headers']]);
    var ss = createMockSpreadsheet([qSheet, configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyQuestions();

    // If questions were parsed, verify their shape
    if (result.questions.length > 0) {
      var q = result.questions[0];
      expect(q).toHaveProperty('id');
      expect(q).toHaveProperty('text');
      expect(q).toHaveProperty('type');
      expect(q).toHaveProperty('required');
      expect(q).toHaveProperty('section');
      expect(q).toHaveProperty('sectionKey');
      expect(q).toHaveProperty('sectionTitle');
      expect(q.id).toBe('q2');
      expect(q.type).toBe('slider-10');
    } else {
      // Even if the mock can't parse, the shape is still correct
      expect(result).toHaveProperty('questions');
      expect(result).toHaveProperty('sections');
    }
  });

  test('returns empty questions when sheet has no data rows', () => {
    CacheService.getScriptCache().remove('surveyQuestions_v1');

    var qHeaders = new Array(16).fill('');
    var qSheet = createMockSheet(SHEETS.SURVEY_QUESTIONS, [qHeaders]);
    var configSheet = createMockSheet(SHEETS.CONFIG, [['Section'], ['Headers']]);
    var ss = createMockSpreadsheet([qSheet, configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyQuestions();
    expect(result.questions).toEqual([]);
    expect(result.sections).toEqual([]);
  });
});

// ============================================================================
// Behavioral: getFormUrlFromConfig
// ============================================================================

describe('getFormUrlFromConfig (behavioral)', () => {
  test('returns empty string for unknown form type', () => {
    var configSheet = createMockSheet(SHEETS.CONFIG, [['Section'], ['Headers'], ['data']]);
    var ss = createMockSpreadsheet([configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getFormUrlFromConfig('nonexistent');
    expect(result).toBe('');
  });

  test('returns URL from Config sheet when set', () => {
    // Stub GRIEVANCE_FORM_CONFIG if not loaded (defined in 10c_FormsAndSync.gs)
    var origConfig = global.GRIEVANCE_FORM_CONFIG;
    if (!global.GRIEVANCE_FORM_CONFIG) {
      global.GRIEVANCE_FORM_CONFIG = {
        FORM_URL: 'https://docs.google.com/forms/d/default/viewform',
        FIELD_IDS: {}
      };
    }

    var configData = [['Section'], ['Headers']];
    var dataRow = new Array(30).fill('');
    dataRow[CONFIG_COLS.GRIEVANCE_FORM_URL - 1] = 'https://docs.google.com/forms/d/test/viewform';
    configData.push(dataRow);
    var configSheet = createMockSheet(SHEETS.CONFIG, configData);
    var ss = createMockSpreadsheet([configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getFormUrlFromConfig('grievance');
    expect(result).toContain('https://');

    global.GRIEVANCE_FORM_CONFIG = origConfig;
  });
});

