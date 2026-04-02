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
// 2. saveChartImageToDrive
// ============================================================================

describe('saveChartImageToDrive', () => {
  let mockFolder;
  let mockFile;

  beforeEach(() => {
    mockFile = {
      getUrl: jest.fn(() => 'https://drive.google.com/file/mock-id')
    };
    mockFolder = {
      createFile: jest.fn(() => mockFile),
      getId: jest.fn(() => 'folder-id')
    };
    // Mock getOrCreateExportFolder_ by making DriveApp.getFoldersByName return a folder
    DriveApp.getFoldersByName.mockImplementation(() => ({
      hasNext: jest.fn(() => true),
      next: jest.fn(() => mockFolder)
    }));
  });

  test('calls Utilities.base64Decode with the provided data', () => {
    saveChartImageToDrive('testChart', 'base64encodeddata');
    expect(Utilities.base64Decode).toHaveBeenCalledWith('base64encodeddata');
  });

  test('creates a file in the export folder', () => {
    saveChartImageToDrive('testChart', 'base64encodeddata');
    expect(mockFolder.createFile).toHaveBeenCalled();
  });

  test('returns the Drive file URL', () => {
    const url = saveChartImageToDrive('testChart', 'base64encodeddata');
    expect(url).toBe('https://drive.google.com/file/mock-id');
  });
});

// ============================================================================
// 3. scheduleEmailReport
// ============================================================================

describe('scheduleEmailReport', () => {
  test('rejects missing email', () => {
    const result = scheduleEmailReport({});
    expect(result.success).toBe(false);
    expect(result.error).toContain('Email');
  });

  test('rejects null config', () => {
    const result = scheduleEmailReport(null);
    expect(result.success).toBe(false);
  });

  test('rejects invalid email format', () => {
    const result = scheduleEmailReport({ email: 'not-an-email' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('Invalid email');
  });

  test('stores schedule in ScriptProperties', () => {
    const config = { email: 'steward@example.com', frequency: 'weekly', sections: ['summary'] };
    const result = scheduleEmailReport(config);
    expect(result.success).toBe(true);

    const props = PropertiesService.getScriptProperties();
    expect(props.setProperty).toHaveBeenCalled();
    const setCall = props.setProperty.mock.calls.find(c => c[0] === 'report_schedules');
    expect(setCall).toBeTruthy();
    const stored = JSON.parse(setCall[1]);
    expect(stored.length).toBeGreaterThan(0);
    expect(stored[0].email).toBe('steward@example.com');
  });

  test('installs a trigger', () => {
    const config = { email: 'steward@example.com', frequency: 'daily' };
    scheduleEmailReport(config);
    expect(ScriptApp.newTrigger).toHaveBeenCalledWith('sendScheduledReports');
  });

  test('creates schedule with correct structure', () => {
    const config = { email: 'steward@example.com', frequency: 'monthly', sections: ['summary', 'charts'] };
    const result = scheduleEmailReport(config);
    expect(result.success).toBe(true);
    expect(result.schedule).toHaveProperty('id');
    expect(result.schedule.email).toBe('steward@example.com');
    expect(result.schedule.frequency).toBe('monthly');
    expect(result.schedule.sections).toEqual(['summary', 'charts']);
    expect(result.schedule.active).toBe(true);
  });

  test('emits report:scheduled event', () => {
    const handler = jest.fn();
    EventBus.on('report:scheduled', handler);
    scheduleEmailReport({ email: 'steward@example.com', frequency: 'weekly' });
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({
      email: 'steward@example.com'
    }));
  });

  test('PII restriction for non-stewards', () => {
    getUserRole_.mockImplementation(() => 'member');
    // Self-send with PII should fail for non-steward
    Session.getActiveUser.mockImplementation(() => ({
      getEmail: jest.fn(() => 'member@example.com')
    }));
    const result = scheduleEmailReport({ email: 'other@example.com', includePII: true });
    expect(result.success).toBe(false);
    expect(result.error).toContain('PII');
  });

  // C-AUTH-7 regression: old `||` logic incorrectly blocked both of these
  test('C-AUTH-7 regression: PII report sent to self is allowed for non-steward', () => {
    getUserRole_.mockImplementation(() => 'member');
    Session.getActiveUser.mockImplementation(() => ({
      getEmail: jest.fn(() => 'member@example.com')
    }));
    // Sending PII to yourself should succeed even if you are not steward/admin
    const result = scheduleEmailReport({ email: 'member@example.com', includePII: true });
    expect(result.success).toBe(true);
  });

  test('C-AUTH-7 regression: non-PII report to any recipient is allowed for non-steward', () => {
    getUserRole_.mockImplementation(() => 'member');
    Session.getActiveUser.mockImplementation(() => ({
      getEmail: jest.fn(() => 'sender@example.com')
    }));
    // Sending a non-PII report to another person should succeed
    const result = scheduleEmailReport({ email: 'recipient@example.com', includePII: false });
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// 4. getScheduledReports
// ============================================================================

describe('getScheduledReports', () => {
  test('returns empty array when no schedules exist', () => {
    const result = JSON.parse(getScheduledReports());
    expect(result).toEqual([]);
  });

  test('returns only current user schedules', () => {
    // Set up schedules with different creators
    const schedules = [
      { id: 'rpt_1', email: 'test@example.com', createdBy: 'test@example.com' },
      { id: 'rpt_2', email: 'other@example.com', createdBy: 'other@example.com' }
    ];
    PropertiesService.getScriptProperties().setProperty('report_schedules', JSON.stringify(schedules));
    // Re-mock getProperty to read from the store
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(key => {
      if (key === 'report_schedules') return JSON.stringify(schedules);
      return null;
    });

    const result = JSON.parse(getScheduledReports());
    expect(result.length).toBe(1);
    expect(result[0].createdBy).toBe('test@example.com');
  });

  test('returns a JSON string', () => {
    const result = getScheduledReports();
    expect(typeof result).toBe('string');
    expect(() => JSON.parse(result)).not.toThrow();
  });
});

// ============================================================================
// 5. removeScheduledReport
// ============================================================================

describe('removeScheduledReport', () => {
  test('removes matching schedule', () => {
    const schedules = [
      { id: 'rpt_1', email: 'test@example.com', createdBy: 'test@example.com' },
      { id: 'rpt_2', email: 'test@example.com', createdBy: 'test@example.com' }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(key => {
      if (key === 'report_schedules') return JSON.stringify(schedules);
      return null;
    });

    const result = removeScheduledReport('rpt_1');
    expect(result.success).toBe(true);
    // Verify it was saved with the removed schedule filtered out
    const saveCall = props.setProperty.mock.calls.find(c => c[0] === 'report_schedules');
    expect(saveCall).toBeTruthy();
    const saved = JSON.parse(saveCall[1]);
    expect(saved.length).toBe(1);
    expect(saved[0].id).toBe('rpt_2');
  });

  test('only owner can remove their schedule', () => {
    const schedules = [
      { id: 'rpt_1', email: 'other@example.com', createdBy: 'other@example.com' }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(key => {
      if (key === 'report_schedules') return JSON.stringify(schedules);
      return null;
    });

    // Current user is test@example.com, schedule was created by other@example.com
    removeScheduledReport('rpt_1');
    const saveCall = props.setProperty.mock.calls.find(c => c[0] === 'report_schedules');
    const saved = JSON.parse(saveCall[1]);
    // Should still have the schedule because owner doesn't match
    expect(saved.length).toBe(1);
  });

  test('emits report:removed event', () => {
    const handler = jest.fn();
    EventBus.on('report:removed', handler);
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(() => '[]');

    removeScheduledReport('rpt_1');
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({ scheduleId: 'rpt_1' }));
  });
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

describe('getSharedViews', () => {
  test('returns views accessible to the current user as creator', () => {
    const views = [
      { id: 'v1', name: 'My View', createdBy: 'test@example.com', sharedWith: [] },
      { id: 'v2', name: 'Other View', createdBy: 'other@example.com', sharedWith: [] }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'shared_views') return JSON.stringify(views);
      return null;
    });

    const result = JSON.parse(getSharedViews());
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('v1');
  });

  test('returns views where current user is in sharedWith', () => {
    const views = [
      { id: 'v1', name: 'Shared To Me', createdBy: 'other@example.com', sharedWith: ['test@example.com'] }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'shared_views') return JSON.stringify(views);
      return null;
    });

    const result = JSON.parse(getSharedViews());
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('v1');
  });
});

describe('deleteSharedView', () => {
  test('deletes only views owned by the current user', () => {
    const views = [
      { id: 'v1', name: 'My View', createdBy: 'test@example.com', sharedWith: [] },
      { id: 'v2', name: 'Other View', createdBy: 'other@example.com', sharedWith: [] }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'shared_views') return JSON.stringify(views);
      return null;
    });

    const result = deleteSharedView('v1');
    expect(result.success).toBe(true);

    const saveCall = props.setProperty.mock.calls.find(c => c[0] === 'shared_views');
    const saved = JSON.parse(saveCall[1]);
    expect(saved.length).toBe(1);
    expect(saved[0].id).toBe('v2');
  });

  test('returns error when trying to delete view not owned by user', () => {
    const views = [
      { id: 'v2', name: 'Other View', createdBy: 'other@example.com', sharedWith: [] }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'shared_views') return JSON.stringify(views);
      return null;
    });

    const result = deleteSharedView('v2');
    expect(result.success).toBe(false);
    expect(result.error).toContain('not authorized');
  });

  test('emits collaboration:viewDeleted event on success', () => {
    const handler = jest.fn();
    EventBus.on('collaboration:viewDeleted', handler);

    const views = [
      { id: 'v1', name: 'My View', createdBy: 'test@example.com', sharedWith: [] }
    ];
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'shared_views') return JSON.stringify(views);
      return null;
    });

    deleteSharedView('v1');
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({ viewId: 'v1' }));
  });
});

// ============================================================================
// 10. saveChartPreset / getChartPresets / deleteChartPreset / updateChartPreset
// ============================================================================

describe('saveChartPreset', () => {
  test('saves a preset and returns it', () => {
    const result = saveChartPreset({
      name: 'My Preset',
      visibleCharts: ['statusChart', 'locationChart'],
      chartOptions: {},
      layout: { columns: 3 }
    });
    expect(result.success).toBe(true);
    expect(result.preset).toHaveProperty('id');
    expect(result.preset.name).toBe('My Preset');
    expect(result.preset.visibleCharts).toEqual(['statusChart', 'locationChart']);
  });

  test('emits preset:saved event', () => {
    const handler = jest.fn();
    EventBus.on('preset:saved', handler);

    saveChartPreset({ name: 'Test Preset' });
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({
      name: 'Test Preset'
    }));
  });
});

describe('getChartPresets', () => {
  test('returns presets as a JSON string', () => {
    const result = getChartPresets();
    expect(typeof result).toBe('string');
    expect(() => JSON.parse(result)).not.toThrow();
  });
});

describe('deleteChartPreset', () => {
  test('removes the matching preset', () => {
    const presets = [
      { id: 'preset_1', name: 'A' },
      { id: 'preset_2', name: 'B' }
    ];
    const props = PropertiesService.getUserProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'chart_presets') return JSON.stringify(presets);
      return null;
    });

    const result = deleteChartPreset('preset_1');
    expect(result.success).toBe(true);

    const saveCall = props.setProperty.mock.calls.find(c => c[0] === 'chart_presets');
    const saved = JSON.parse(saveCall[1]);
    expect(saved.length).toBe(1);
    expect(saved[0].id).toBe('preset_2');
  });

  test('emits preset:deleted event', () => {
    const handler = jest.fn();
    EventBus.on('preset:deleted', handler);
    const props = PropertiesService.getUserProperties();
    props.getProperty.mockImplementation(() => '[]');

    deleteChartPreset('preset_99');
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({ presetId: 'preset_99' }));
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
