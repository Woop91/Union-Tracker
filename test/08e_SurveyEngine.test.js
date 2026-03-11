/**
 * Tests for 08e_SurveyEngine.gs
 * Covers survey period management, triggers, and satisfaction summaries.
 */

require('./gas-mock');
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
