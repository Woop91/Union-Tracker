/**
 * Tests for 10b_SurveyDocSheets.gs
 * Covers survey question sheet, satisfaction sheet, feedback sheet,
 * FAQ, features reference, resources, and notification sheet creation.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '05_Integrations.gs', '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs', '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs'
]);

describe('10b function existence', () => {
  const required = [
    'createSurveyQuestionsSheet', 'clearSurveyQuestionsCache',
    'createSatisfactionSheet', 'createFeedbackSheet',
    'populateRoadmapItems', 'createFunctionChecklistSheet_',
    'createGettingStartedSheet', 'createFAQSheet',
    'createFeaturesReferenceSheet', 'createResourcesSheet',
    'createNotificationsSheet', 'createResourceConfigSheet'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('Notification sheet uses dynamic email', () => {
  test('createNotificationsSheet does not hardcode massability.org', () => {
    const fs = require('fs');
    const path = require('path');
    const code = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10b_SurveyDocSheets.gs'), 'utf8'
    );
    // Should use systemEmail variable, not hardcoded
    expect(code).not.toContain("'system@massability.org'");
  });
});

describe('Sheet name constants', () => {
  test('SHEETS.NOTIFICATIONS is defined', () => {
    expect(SHEETS.NOTIFICATIONS).toBeDefined();
    expect(typeof SHEETS.NOTIFICATIONS).toBe('string');
  });

  test('SHEETS.FEEDBACK is defined', () => {
    expect(SHEETS.FEEDBACK).toBeDefined();
  });

  test('SHEETS.SATISFACTION is defined', () => {
    expect(SHEETS.SATISFACTION).toBeDefined();
  });
});
