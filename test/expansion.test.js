/**
 * Expansion tests — Phase 7 seed data coverage
 *
 * Tests for v4.15.0 DevTools seed functions:
 *   - seedUnionStatsData, seedCalendarEvents, seedWeeklyQuestions
 *   - dataGetEngagementStats (requires seeded data from DevTools)
 *
 * NOTE: DataService tests (getGrievanceStats, getGrievanceHotSpots,
 * getStewardDirectory, isChiefSteward, welcome/broadcast helpers) were
 * moved to 21_WebDashDataService.test.js as their canonical home.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load in GAS load order — DevTools needs DataManagers + DataService
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '21_WebDashDataService.gs',
  '07_DevTools.gs',
]);

// ============================================================================
// Seed data functions (07_DevTools.gs)
// ============================================================================

describe('seedUnionStatsData', () => {
  beforeEach(() => {
    PropertiesService.getScriptProperties().deleteAllProperties();
  });

  test('stores engagement data in Script Properties', () => {
    seedUnionStatsData();
    const raw = PropertiesService.getScriptProperties().getProperty('SEEDED_UNION_STATS');
    expect(raw).toBeTruthy();
    const stats = JSON.parse(raw);
    expect(stats.engagement).toBeDefined();
    expect(stats.engagement.surveyParticipation).toBeDefined();
    expect(stats.engagement.weeklyQuestionVotes).toBeDefined();
  });

  test('stores membership trends array', () => {
    seedUnionStatsData();
    const stats = JSON.parse(PropertiesService.getScriptProperties().getProperty('SEEDED_UNION_STATS'));
    expect(stats.membershipTrends).toBeDefined();
    expect(Array.isArray(stats.membershipTrends)).toBe(true);
    expect(stats.membershipTrends.length).toBeGreaterThan(0);
  });
});

describe('seedCalendarEvents', () => {
  test('is a defined function', () => {
    expect(typeof seedCalendarEvents).toBe('function');
  });
});

describe('seedWeeklyQuestions', () => {
  test('is a defined function', () => {
    expect(typeof seedWeeklyQuestions).toBe('function');
  });
});

// ============================================================================
// dataGetEngagementStats — requires seeded data (integration with DevTools)
// ============================================================================

describe('dataGetEngagementStats (live sheet reads)', () => {
  test('returns engagement data from sheet when members exist', () => {
    // dataGetEngagementStats now reads directly from sheets (not PropertiesService).
    // Mock _resolveCallerEmail for auth and provide a Member Directory with active members.
    const origResolve = global._resolveCallerEmail;
    global._resolveCallerEmail = jest.fn(() => 'steward@test.com');

    const memberData = [
      ['Email', 'Name', 'Dues Status', 'Hire Date'],
      ['alice@test.com', 'Alice', 'Active', '2025-01-15'],
      ['bob@test.com', 'Bob', 'Active', '2025-06-01'],
    ];
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', memberData);
    // Override getRange to be data-aware (the source uses getRange, not getDataRange)
    memberSheet.getRange = jest.fn((row, col, numRows, numCols) => {
      const nr = numRows || 1;
      const nc = numCols || (memberData[0] ? memberData[0].length : 1);
      const slice = memberData.slice(row - 1, row - 1 + nr).map(r => r.slice(col - 1, col - 1 + nc));
      return {
        getValues: jest.fn(() => slice),
        getValue: jest.fn(() => (slice[0] && slice[0][0]) || ''),
        setValue: jest.fn(),
        setValues: jest.fn()
      };
    });
    const mockSs = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);

    const result = dataGetEngagementStats('test-session-token');
    expect(result).toBeDefined();
    expect(result).not.toBeNull();
    expect(typeof result.surveyParticipation).toBe('number');
    expect(result.membershipTrends).toBeDefined();
    expect(Array.isArray(result.membershipTrends)).toBe(true);

    global._resolveCallerEmail = origResolve;
  });
});
