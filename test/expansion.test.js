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

describe('dataGetEngagementStats (with seeded data)', () => {
  test('returns engagement data when seeded', () => {
    PropertiesService.getScriptProperties().deleteAllProperties();
    seedUnionStatsData();
    const result = dataGetEngagementStats();
    expect(result).toBeDefined();
    expect(result.surveyParticipation).toBeDefined();
    expect(result.membershipTrends).toBeDefined();
    expect(Array.isArray(result.membershipTrends)).toBe(true);
  });
});
