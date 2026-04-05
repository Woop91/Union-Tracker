// Tests functions from: 16_DashboardEnhancements.gs, 09_Dashboards.gs
/**
 * Tests for 16_DashboardEnhancements.gs
 *
 * Covers date range presets, chart image export, scheduled email reports,
 * notifications, shared views, chart presets, and filtered dashboard data.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock dependencies before loading source
global.getUnifiedDashboardData = jest.fn(() => JSON.stringify({
  totalMembers: 100, openCases: 10, winRate: 60, moraleScore: 7,
  statusDistribution: { Open: 5, Closed: 3 },
  locationBreakdown: { 'HQ': 50, 'Branch': 30 },
  unitBreakdown: { 'Unit A': 30 },
  grievancesByCategory: { 'Discipline': 4, 'Contract': 6 },
  chartDrillDown: { statusByCase: {}, locationByCase: {} },
  stepProgression: { step1: 5, step2: 3, step3: 1, arb: 0 },
  stewardPerformance: [],
  monthlyFilings: []
}));
global.getUnifiedDashboardDataWithDateRange = jest.fn(() => getUnifiedDashboardData());
global.isTruthyValue = jest.fn(v => !!v);
global.getUserRole_ = jest.fn(() => 'steward');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '15_EventBus.gs', '09_Dashboards.gs']);

// Singleton PropertiesService mocks — the source code and tests must share the
// same jest.fn() instance so that spy.mock.calls reflects source-code calls.
let mockScriptProps, mockUserProps;

// Reset EventBus and properties between tests
beforeEach(() => {
  EventBus.reset();
  jest.clearAllMocks();

  // Create fresh singletons AFTER clearAllMocks so spy histories start clean
  const scriptStore = {};
  mockScriptProps = {
    setProperty: jest.fn((k, v) => { scriptStore[k] = v; }),
    getProperty: jest.fn(k => scriptStore[k] || null),
    deleteProperty: jest.fn(k => { delete scriptStore[k]; }),
    deleteAllProperties: jest.fn(() => { for (var k in scriptStore) delete scriptStore[k]; }),
    getProperties: jest.fn(() => Object.assign({}, scriptStore))
  };
  PropertiesService.getScriptProperties = jest.fn(() => mockScriptProps);

  const userStore = {};
  mockUserProps = {
    setProperty: jest.fn((k, v) => { userStore[k] = v; }),
    getProperty: jest.fn(k => userStore[k] || null),
    deleteProperty: jest.fn(k => { delete userStore[k]; }),
    deleteAllProperties: jest.fn(() => { for (var k in userStore) delete userStore[k]; }),
    getProperties: jest.fn(() => Object.assign({}, userStore))
  };
  PropertiesService.getUserProperties = jest.fn(() => mockUserProps);

  // Reset Session — clearAllMocks() clears call history but not implementations,
  // so tests that set mockImplementation on Session can leak into later tests.
  Session.getActiveUser = jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  }));

  // Re-set the default mock implementations after clearAllMocks
  getUnifiedDashboardData.mockImplementation(() => JSON.stringify({
    totalMembers: 100, openCases: 10, winRate: 60, moraleScore: 7,
    statusDistribution: { Open: 5, Closed: 3 },
    locationBreakdown: { 'HQ': 50, 'Branch': 30 },
    unitBreakdown: { 'Unit A': 30 },
    grievancesByCategory: { 'Discipline': 4, 'Contract': 6 },
    chartDrillDown: { statusByCase: {}, locationByCase: {} },
    stepProgression: { step1: 5, step2: 3, step3: 1, arb: 0 },
    stewardPerformance: [],
    monthlyFilings: []
  }));
  getUnifiedDashboardDataWithDateRange.mockImplementation(() => getUnifiedDashboardData());
  global.isTruthyValue = jest.fn(v => !!v);
  global.getUserRole_ = jest.fn(() => 'steward');
});

// ============================================================================
// 6. getUserNotifications — REMOVED (migrated to sheet-based system in v4.13.0+)
// ============================================================================

// ============================================================================
// 7. pushNotification (writes to Notifications sheet)
// ============================================================================

describe('pushNotification', () => {
  let mockNotifSheet;

  beforeEach(() => {
    // Mock Notifications sheet with header row + existing data
    const notifData = [
      ['Notification ID', 'Recipient', 'Type', 'Title', 'Message', 'Priority', 'Sent By', 'Sent By Name', 'Created Date', 'Expires Date', 'Dismissed By', 'Status', 'Dismiss Mode'],
      ['NOTIF-001', 'old@example.com', 'System', 'Old', 'Old msg', 'Normal', 'system', 'System', '2026-01-01', '', '', 'Active', 'Dismissible']
    ];
    mockNotifSheet = {
      getName: jest.fn(() => '\uD83D\uDCE2 Notifications'),
      getDataRange: jest.fn(() => ({
        getValues: jest.fn(() => notifData)
      })),
      appendRow: jest.fn(),
      getLastRow: jest.fn(() => notifData.length)
    };
    const mockSs = {
      getSheetByName: jest.fn(name => {
        if (name === SHEETS.NOTIFICATIONS) return mockNotifSheet;
        return null;
      }),
      insertSheet: jest.fn(),
      toast: jest.fn()
    };
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
  });

  test('adds notification to sheet and returns success', () => {
    const result = pushNotification('test@example.com', {
      title: 'Test', body: 'Hello', type: 'info'
    });
    expect(result.success).toBe(true);
    expect(result.id).toBeDefined();
    expect(mockNotifSheet.appendRow).toHaveBeenCalledTimes(1);
  });

  test('returns error when notification title is missing', () => {
    const result = pushNotification('test@example.com', { body: 'No title' });
    expect(result.success).toBe(false);
  });

  test('returns error when userEmail is empty', () => {
    const result = pushNotification('', { title: 'Test' });
    expect(result.success).toBe(false);
  });

  test('generates sequential NOTIF-NNN id', () => {
    const result = pushNotification('test@example.com', {
      title: 'Test', body: 'Hello', type: 'info'
    });
    expect(result.id).toBe('NOTIF-002');
  });

  test('emits notification:pushed event', () => {
    const handler = jest.fn();
    EventBus.on('notification:pushed', handler);

    pushNotification('test@example.com', { title: 'Test', body: 'msg', type: 'info' });
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({
      userId: 'test@example.com'
    }));
  });
});

// ============================================================================
// 8. markNotificationRead — REMOVED (migrated to sheet-based system)
// ============================================================================

// ============================================================================
// 9. saveSharedView / getSharedViews / deleteSharedView
// ============================================================================

describe('saveSharedView', () => {
  test('saves a view configuration', () => {
    const view = {
      name: 'My View',
      selectedCharts: ['chart1', 'chart2'],
      filters: { statuses: ['Open'] },
      sharedWith: []
    };
    const result = saveSharedView(view);
    expect(result.success).toBe(true);
    expect(result.view).toHaveProperty('id');
    expect(result.view.name).toBe('My View');
    expect(result.view.createdBy).toBe('test@example.com');
  });

  test('notifies shared users', () => {
    // pushNotification now writes to the Notifications sheet — mock it
    const notifData = [['Notification ID', 'Recipient', 'Type', 'Title', 'Message', 'Priority', 'Sent By', 'Sent By Name', 'Created Date', 'Expires Date', 'Dismissed By', 'Status', 'Dismiss Mode']];
    const mockNotifSheet = {
      getName: jest.fn(() => '\uD83D\uDCE2 Notifications'),
      getDataRange: jest.fn(() => ({ getValues: jest.fn(() => notifData) })),
      appendRow: jest.fn(),
      getLastRow: jest.fn(() => notifData.length)
    };
    const mockSs = {
      getSheetByName: jest.fn(name => {
        if (name === SHEETS.NOTIFICATIONS) return mockNotifSheet;
        return null;
      }),
      insertSheet: jest.fn(),
      toast: jest.fn()
    };
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);

    const handler = jest.fn();
    EventBus.on('notification:pushed', handler);

    const view = {
      name: 'Shared View',
      selectedCharts: ['chart1'],
      sharedWith: ['colleague@example.com']
    };
    saveSharedView(view);
    // pushNotification is called for each shared user, which emits notification:pushed
    expect(handler).toHaveBeenCalled();
  });

  test('emits collaboration:viewCreated event', () => {
    const handler = jest.fn();
    EventBus.on('collaboration:viewCreated', handler);

    saveSharedView({ name: 'Test', sharedWith: [] });
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({
      viewId: expect.any(String)
    }));
  });
});

// ============================================================================
// 11. getFilteredDashboardData
// ============================================================================

describe('getFilteredDashboardData', () => {
  test('returns full data when no filters provided', () => {
    const result = JSON.parse(getFilteredDashboardData(false, null));
    expect(result.totalMembers).toBe(100);
    expect(result.statusDistribution).toEqual({ Open: 5, Closed: 3 });
  });

  test('applies status filters', () => {
    const result = JSON.parse(getFilteredDashboardData(false, { statuses: ['Open'] }));
    expect(result.statusDistribution).toEqual({ Open: 5 });
    expect(result.statusDistribution.Closed).toBeUndefined();
  });

  test('applies location filters', () => {
    const result = JSON.parse(getFilteredDashboardData(false, { locations: ['HQ'] }));
    expect(result.locationBreakdown).toEqual({ 'HQ': 50 });
    expect(result.locationBreakdown['Branch']).toBeUndefined();
  });
});
