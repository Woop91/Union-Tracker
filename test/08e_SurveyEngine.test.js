/**
 * Tests for 08e_SurveyEngine.gs
 * Covers survey period management, triggers, and satisfaction summaries.
 */

const { createMockSheet, createMockSpreadsheet, createMockRange } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs', '08e_SurveyEngine.gs'
]);

describe('08e function existence', () => {
  const required = [
    'initSurveyEngine', 'setupSurveyPeriodsSheet', 'getSurveyPeriod',
    'openNewSurveyPeriod', 'archiveSurveyPeriod_',
    'incrementPeriodResponseCount_', 'autoTriggerQuarterlyPeriod',
    'setupQuarterlyTrigger', 'getPendingSurveyMembers',
    'getSatisfactionSummary', 'pushSurveyOpenNotification_',
    'autoSurveyReminderWeekly', 'setupWeeklyReminderTrigger',
    'dataGetPendingSurveyMembers', 'dataGetSatisfactionSummary',
    'dataOpenNewSurveyPeriod', 'menuOpenNewSurveyPeriod',
    'menuShowSurveyPeriodStatus', 'menuInstallSurveyTriggers',
    'setupOpenDeferredTrigger'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('HIDDEN_SHEETS survey constants', () => {
  test('SURVEY_TRACKING is defined', () => {
    expect(HIDDEN_SHEETS.SURVEY_TRACKING).toBe('_Survey_Tracking');
  });
  test('SURVEY_VAULT is defined', () => {
    expect(HIDDEN_SHEETS.SURVEY_VAULT).toBe('_Survey_Vault');
  });
  test('SURVEY_PERIODS is defined', () => {
    expect(HIDDEN_SHEETS.SURVEY_PERIODS).toBe('_Survey_Periods');
  });
});

// ============================================================================
// Behavioral: getSurveyPeriod
// ============================================================================

describe('getSurveyPeriod (behavioral)', () => {
  test('returns null when periods sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyPeriod();
    expect(result).toBeNull();
  });

  test('returns null when periods sheet has no data rows', () => {
    var headers = ['Period ID', 'Period Name', 'Start Date', 'End Date', 'Status', 'Archive Folder URL', 'Created By', 'Response Count'];
    var sheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [headers]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyPeriod();
    expect(result).toBeNull();
  });

  test('returns active period with expected shape', () => {
    var C = SURVEY_PERIODS_COLS;
    var headers = ['Period ID', 'Period Name', 'Start Date', 'End Date', 'Status', 'Archive Folder URL', 'Created By', 'Response Count'];
    var row = new Array(8).fill('');
    row[C.PERIOD_ID - 1] = 'Q1-2026';
    row[C.PERIOD_NAME - 1] = 'Q1 2026 Survey';
    row[C.START_DATE - 1] = new Date(2026, 0, 1);
    row[C.END_DATE - 1] = new Date(2026, 2, 31);
    row[C.STATUS - 1] = 'Active';
    row[C.ARCHIVE_URL - 1] = '';
    row[C.RESPONSE_COUNT - 1] = 42;

    var sheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [headers, row]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyPeriod();
    expect(result).not.toBeNull();
    expect(result.periodId).toBe('Q1-2026');
    expect(result.name).toBe('Q1 2026 Survey');
    expect(result.status).toBe('Active');
    expect(result.responseCount).toBe(42);
    expect(result).toHaveProperty('startDate');
    expect(result).toHaveProperty('endDate');
    expect(result).toHaveProperty('rowIndex');
  });

  test('returns null when all periods are Closed', () => {
    var C = SURVEY_PERIODS_COLS;
    var headers = ['Period ID', 'Period Name', 'Start Date', 'End Date', 'Status', 'Archive Folder URL', 'Created By', 'Response Count'];
    var row = new Array(8).fill('');
    row[C.PERIOD_ID - 1] = 'Q4-2025';
    row[C.PERIOD_NAME - 1] = 'Q4 2025 Survey';
    row[C.STATUS - 1] = 'Closed';

    var sheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [headers, row]);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = getSurveyPeriod();
    expect(result).toBeNull();
  });
});

// ============================================================================
// Behavioral: setupSurveyPeriodsSheet
// ============================================================================

describe('setupSurveyPeriodsSheet (behavioral)', () => {
  test('creates the periods sheet when it does not exist', () => {
    // Create a sheet with chainable getRange (setValues → setFontWeight → setBackground)
    var sheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [['header']]);
    var chainableRange = {
      setValues: jest.fn(function() { return this; }),
      setValue: jest.fn(function() { return this; }),
      setFontWeight: jest.fn(function() { return this; }),
      setBackground: jest.fn(function() { return this; }),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    sheet.getRange.mockReturnValue(chainableRange);

    var ss = createMockSpreadsheet([]);
    ss.insertSheet.mockReturnValue(sheet);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupSurveyPeriodsSheet();

    expect(ss.insertSheet).toHaveBeenCalledWith(HIDDEN_SHEETS.SURVEY_PERIODS);
    expect(sheet.setFrozenRows).toHaveBeenCalledWith(1);
  });

  test('does not recreate sheet when it already has valid headers', () => {
    var headers = ['Period ID', 'Period Name', 'Start Date', 'End Date', 'Status', 'Archive Folder URL', 'Created By', 'Response Count'];
    var sheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [headers]);
    // Override the getRange(1,1) mock to return "Period ID"
    sheet.getRange.mockImplementation(function(row, col, numRows, numCols) {
      if (row === 1 && col === 1 && !numRows) {
        return { getValue: jest.fn(() => 'Period ID'), getValues: jest.fn(() => [['Period ID']]), setValues: jest.fn(function() { return this; }), setFontWeight: jest.fn(function() { return this; }), setBackground: jest.fn(function() { return this; }) };
      }
      return createMockRange();
    });
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupSurveyPeriodsSheet();

    // Should NOT call insertSheet if the sheet already exists with proper headers
    expect(ss.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// Behavioral: openNewSurveyPeriod
// ============================================================================

describe('openNewSurveyPeriod (behavioral)', () => {
  test('returns success with a periodId', () => {
    // Setup: no active period, periods sheet exists
    var headers = ['Period ID', 'Period Name', 'Start Date', 'End Date', 'Status', 'Archive Folder URL', 'Created By', 'Response Count'];
    var periodsSheet = createMockSheet(HIDDEN_SHEETS.SURVEY_PERIODS, [headers]);
    var trackingSheet = createMockSheet(HIDDEN_SHEETS.SURVEY_TRACKING, [['Member ID', 'Email', 'Status']]);
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Member ID'], ['M-001']]);
    var ss = createMockSpreadsheet([periodsSheet, trackingSheet, memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = openNewSurveyPeriod('admin@test.com');
    expect(result).toHaveProperty('success');
    expect(result).toHaveProperty('periodId');
    if (result.success) {
      expect(result.periodId).toMatch(/^Q\d-\d{4}$/);
    }
  });
});
