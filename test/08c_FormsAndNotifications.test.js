/**
 * Tests for 08c_FormsAndNotifications.gs
 * Covers form URL handling, survey tracking, deadline notifications,
 * email validation, quarter calculation, and submission handling.
 */

require('./gas-mock');
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
    'getFormUrlFromConfig', 'buildGrievanceFormUrl_', 'saveFormUrlsToConfig',
    'saveFormUrlsToConfig_silent', 'getFormValue_', 'getFormMultiValue_',
    'parseFormDate_', 'sendContactInfoForm', 'onContactFormSubmit',
    'setupContactFormTrigger', 'setupGrievanceFormTrigger',
    'getSatisfactionSurveyLink', 'onSatisfactionFormSubmit',
    'setupSatisfactionFormTrigger', 'showNotificationSettings',
    'installDailyTrigger_', 'removeDailyTrigger_',
    'checkDeadlinesAndNotify_', 'testDeadlineNotifications',
    'sendStewardDeadlineAlerts', 'sendStewardAlertsNow',
    'configureAlertSettings', 'sendRandomSurveyEmails',
    'executeSendRandomSurveyEmails', 'validateMemberEmail',
    'getCurrentQuarter', 'getQuarterFromDate',
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

describe('getQuarterFromDate', () => {
  test('January is Q1', () => {
    expect(getQuarterFromDate(new Date(2026, 0, 15))).toBe('2026-Q1');
  });
  test('April is Q2', () => {
    expect(getQuarterFromDate(new Date(2026, 3, 1))).toBe('2026-Q2');
  });
  test('July is Q3', () => {
    expect(getQuarterFromDate(new Date(2026, 6, 30))).toBe('2026-Q3');
  });
  test('October is Q4', () => {
    expect(getQuarterFromDate(new Date(2026, 9, 1))).toBe('2026-Q4');
  });
  test('December is Q4', () => {
    expect(getQuarterFromDate(new Date(2026, 11, 31))).toBe('2026-Q4');
  });
  test('invalid date returns empty string', () => {
    expect(getQuarterFromDate('not-a-date')).toBe('');
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
// parseFormDate_
// ============================================================================

describe('parseFormDate_', () => {
  test('parses valid Date object', () => {
    const d = new Date(2026, 2, 11);
    const result = parseFormDate_(d);
    expect(result).toBeInstanceOf(Date);
  });

  test('returns null for empty input', () => {
    const result = parseFormDate_('');
    expect(result == null || result === '').toBeTruthy();
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
