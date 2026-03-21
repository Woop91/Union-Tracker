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
    'getFormUrlFromConfig', 'buildGrievanceFormUrl_',
    'saveFormUrlsToConfig_silent', 'getFormValue_', 'getFormMultiValue_',
    'parseFormDate_',
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
