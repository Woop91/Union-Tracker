/**
 * Expansion tests — Phase 7 feature coverage
 *
 * Tests for v4.15.0 Phase 7 features:
 *   - _buildGrievanceRecord issueCategory + closedTimestamp fields
 *   - getGrievanceStats monthlyResolved + summary counts
 *   - Seed data functions (calendar events, weekly questions, union stats)
 *   - Chief steward task assignment
 *   - Login UX helpers (welcome experience)
 *   - Broadcast filter helpers
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '21_WebDashDataService.gs',
]);

// ============================================================================
// DataService.getGrievanceStats — monthlyResolved + summary counts
// ============================================================================

describe('getGrievanceStats', () => {
  beforeEach(() => {
    // Build a mock Grievance Log with known data
    const headers = [
      'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
      'Filed', 'Steward', 'Unit', 'Priority', 'Notes',
      'Issue Category', 'Resolution', 'Date Closed',
    ];

    // 12 rows to pass the minimum-10 threshold
    const rows = [
      ['G001', 'a@test.com', 'Won', 'Step II', '', '2025-06-15', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won at Step II', '2025-07-01'],
      ['G002', 'b@test.com', 'Denied', 'Step III', '', '2025-06-20', 'steward@test.com', 'Unit A', 'medium', '', 'Pay', 'Denied', '2025-08-01'],
      ['G003', 'c@test.com', 'Settled', 'Step I', '', '2025-07-01', 'steward@test.com', 'Unit B', 'low', '', 'Safety', 'Settlement reached', '2025-07-15'],
      ['G004', 'd@test.com', 'Active', 'Step I', '', '2025-07-10', 'steward@test.com', 'Unit A', 'high', '', 'Discipline', '', ''],
      ['G005', 'e@test.com', 'New', 'Step I', '', '2025-08-01', 'steward@test.com', 'Unit C', 'medium', '', 'Pay', '', ''],
      ['G006', 'f@test.com', 'Won', 'Step II', '', '2025-08-15', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won', '2025-09-01'],
      ['G007', 'g@test.com', 'Active', 'Step II', '', '2025-09-01', 'steward@test.com', 'Unit B', 'medium', '', 'Workload', '', ''],
      ['G008', 'h@test.com', 'New', 'Step I', '', '2025-09-10', 'steward@test.com', 'Unit C', 'low', '', 'Pay', '', ''],
      ['G009', 'i@test.com', 'Active', 'Step I', '', '2025-10-01', 'steward@test.com', 'Unit A', 'high', '', 'Safety', '', ''],
      ['G010', 'j@test.com', 'Settled', 'Step III', '', '2025-10-15', 'steward@test.com', 'Unit B', 'medium', '', 'Discipline', 'Settled', '2025-11-01'],
      ['G011', 'k@test.com', 'Won', 'Step I', '', '2025-11-01', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won', '2025-11-15'],
      ['G012', 'l@test.com', 'Active', 'Step II', '', '2025-12-01', 'steward@test.com', 'Unit C', 'medium', '', 'Pay', '', ''],
    ];

    const data = [headers, ...rows];
    const mockSheet = createMockSheet('Grievance Log', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
  });

  test('returns available: true with 12 records', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(true);
    expect(stats.total).toBe(12);
  });

  test('returns byCategory with real category names (not all Uncategorized)', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.byCategory).toBeDefined();
    const cats = Object.keys(stats.byCategory);
    // Should have multiple categories, not just "Uncategorized"
    expect(cats.length).toBeGreaterThan(1);
    expect(cats).toContain('Safety');
    expect(cats).toContain('Pay');
  });

  test('returns monthlyResolved array', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.monthlyResolved).toBeDefined();
    expect(Array.isArray(stats.monthlyResolved)).toBe(true);
    expect(stats.monthlyResolved.length).toBeGreaterThan(0);
    // Each entry has month and count
    stats.monthlyResolved.forEach(entry => {
      expect(entry).toHaveProperty('month');
      expect(entry).toHaveProperty('count');
    });
  });

  test('monthlyResolved and monthly arrays share the same month keys', () => {
    const stats = DataService.getGrievanceStats();
    const filedMonths = stats.monthly.map(m => m.month);
    const resolvedMonths = stats.monthlyResolved.map(m => m.month);
    expect(filedMonths).toEqual(resolvedMonths);
  });

  test('returns summary counts (openCount, wonCount, deniedCount, settledCount)', () => {
    const stats = DataService.getGrievanceStats();
    expect(typeof stats.openCount).toBe('number');
    expect(typeof stats.wonCount).toBe('number');
    expect(typeof stats.deniedCount).toBe('number');
    expect(typeof stats.settledCount).toBe('number');
    // Based on test data: 3 won, 1 denied, 2 settled, 4 active/new (open)
    expect(stats.wonCount).toBe(3);
    expect(stats.deniedCount).toBe(1);
    expect(stats.settledCount).toBe(2);
  });

  test('months are limited to at most 12', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.monthly.length).toBeLessThanOrEqual(12);
    expect(stats.monthlyResolved.length).toBeLessThanOrEqual(12);
  });

  test('byStep counts are accurate', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.byStep).toBeDefined();
    const totalByStep = Object.values(stats.byStep).reduce((a, b) => a + b, 0);
    expect(totalByStep).toBe(12);
  });

  test('returns available: false when fewer than 10 records', () => {
    // Clear IIFE in-memory cache so re-mocked spreadsheet is used
    DataService._invalidateSheetCache('Grievance Log');

    const headers = ['Grievance ID', 'Member Email', 'Status', 'Step', 'Filed'];
    const data = [headers,
      ['G001', 'a@test.com', 'Active', 'Step I', '2025-06-01'],
      ['G002', 'b@test.com', 'Active', 'Step I', '2025-06-02'],
    ];
    const mockSheet = createMockSheet('Grievance Log', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(false);
  });
});

// ============================================================================
// DataService.getGrievanceHotSpots
// ============================================================================

describe('getGrievanceHotSpots', () => {
  test('returns locations with 3+ grievances', () => {
    // Clear IIFE in-memory cache so re-mocked spreadsheet is used
    DataService._invalidateSheetCache('Grievance Log');

    const headers = ['Grievance ID', 'Unit'];
    const data = [headers,
      ['G1', 'Hot Unit'], ['G2', 'Hot Unit'], ['G3', 'Hot Unit'],
      ['G4', 'Cold Unit'], ['G5', 'Cold Unit'],
    ];
    const mockSheet = createMockSheet('Grievance Log', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const spots = DataService.getGrievanceHotSpots();
    expect(spots.length).toBe(1);
    expect(spots[0].location).toBe('Hot Unit');
    expect(spots[0].count).toBe(3);
  });

  test('returns empty when no unit has 3+ cases', () => {
    DataService._invalidateSheetCache('Grievance Log');

    const headers = ['Grievance ID', 'Unit'];
    const data = [headers, ['G1', 'A'], ['G2', 'B']];
    const mockSheet = createMockSheet('Grievance Log', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    expect(DataService.getGrievanceHotSpots()).toEqual([]);
  });
});

// ============================================================================
// DataService.getStewardDirectory
// ============================================================================

describe('getStewardDirectory', () => {
  test('returns steward records sorted by name', () => {
    // isSteward is derived from role column: 'steward' or 'both' = isSteward
    const headers = ['Email', 'Name', 'Role', 'Work Location', 'Office Days', 'Phone', 'Unit'];
    const data = [headers,
      ['z@test.com', 'Zara', 'Steward', 'Office B', 'Mon-Fri', '555-0001', 'Unit 1'],
      ['a@test.com', 'Alice', 'Steward', 'Office A', 'Mon-Wed', '555-0002', 'Unit 2'],
      ['m@test.com', 'Mike', 'Member', 'Office A', 'Mon-Fri', '555-0003', 'Unit 1'],
    ];
    const mockSheet = createMockSheet('Member Directory', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stewards = DataService.getStewardDirectory();
    expect(stewards.length).toBe(2); // Only Zara and Alice (role=Steward)
    expect(stewards[0].name).toBe('Alice'); // Sorted alphabetically
    expect(stewards[1].name).toBe('Zara');
    expect(stewards[0]).toHaveProperty('email');
    expect(stewards[0]).toHaveProperty('workLocation');
    expect(stewards[0]).toHaveProperty('phone');
  });
});

// ============================================================================
// Seed data functions
// ============================================================================

// Load DevTools for seed function tests
loadSources(['07_DevTools.gs']);

describe('seedUnionStatsData', () => {
  beforeEach(() => {
    // Clear any stored properties
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
// Chief steward task assignment
// ============================================================================

describe('DataService.isChiefSteward', () => {
  test('is exported from DataService', () => {
    expect(typeof DataService.isChiefSteward).toBe('function');
  });
});

// ============================================================================
// Welcome experience (Login UX)
// ============================================================================

describe('dataGetWelcomeData', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetWelcomeData).toBe('function');
  });
});

describe('dataMarkWelcomeDismissed', () => {
  test('is a defined global function', () => {
    expect(typeof dataMarkWelcomeDismissed).toBe('function');
  });

  test('stores dismissal timestamp in Script Properties', () => {
    PropertiesService.getScriptProperties().deleteAllProperties();
    const result = dataMarkWelcomeDismissed('test@example.com');
    expect(result.success).toBe(true);
    // Check that a WELCOME_DISMISSED_ key was stored
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const dismissKey = Object.keys(allProps).find(k => k.startsWith('WELCOME_DISMISSED_'));
    expect(dismissKey).toBeTruthy();
  });

  test('returns failure when no authenticated user', () => {
    // Mock Session to return empty email (unauthenticated)
    const origGetEmail = Session.getActiveUser().getEmail;
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => '') }));
    const result = dataMarkWelcomeDismissed();
    expect(result.success).toBe(false);
    // Restore
    Session.getActiveUser = jest.fn(() => ({ getEmail: origGetEmail }));
  });
});

// ============================================================================
// dataGetEngagementStats — reads seeded union stats
// ============================================================================

describe('dataGetEngagementStats', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetEngagementStats).toBe('function');
  });

  test('returns null when no stats are seeded', () => {
    PropertiesService.getScriptProperties().deleteAllProperties();
    const result = dataGetEngagementStats();
    expect(result).toBeNull();
  });

  test('returns engagement data when seeded', () => {
    seedUnionStatsData();
    const result = dataGetEngagementStats();
    expect(result).toBeDefined();
    expect(result.surveyParticipation).toBeDefined();
    expect(result.membershipTrends).toBeDefined();
    expect(Array.isArray(result.membershipTrends)).toBe(true);
  });
});

// ============================================================================
// Broadcast filter data
// ============================================================================

describe('dataGetBroadcastFilterOptions', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetBroadcastFilterOptions).toBe('function');
  });

  test('returns filter options from member directory', () => {
    // Clear IIFE in-memory cache so re-mocked spreadsheet is used
    DataService._invalidateSheetCache('Member Directory');

    // Mock steward auth check (CR-AUTH-3 moved to server-side identity)
    const origAuth = global.checkWebAppAuthorization;
    global.checkWebAppAuthorization = jest.fn(() => ({
      isAuthorized: true, email: 'steward@test.com', role: 'steward'
    }));

    const headers = ['Email', 'Name', 'Work Location', 'Office Days', 'Dues Status', 'Role', 'Assigned Steward'];
    const data = [headers,
      ['a@test.com', 'Alice', 'Office A', 'Mon-Wed', 'Active', 'Member', 'steward@test.com'],
      ['b@test.com', 'Bob', 'Office B', 'Thu-Fri', 'Inactive', 'Member', 'steward@test.com'],
      ['c@test.com', 'Carol', 'Office A', 'Mon-Fri', 'Active', 'Member', 'steward@test.com'],
    ];
    const mockSheet = createMockSheet('Member Directory', data);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const filters = dataGetBroadcastFilterOptions();
    expect(filters).toBeDefined();
    expect(filters.locations).toBeDefined();
    expect(Array.isArray(filters.locations)).toBe(true);
    expect(filters.locations).toContain('Office A');
    expect(filters.locations).toContain('Office B');

    // Restore original auth function
    global.checkWebAppAuthorization = origAuth;
  });
});
