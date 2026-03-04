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

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '15_EventBus.gs', '16_DashboardEnhancements.gs']);

// Reset EventBus and properties between tests
beforeEach(() => {
  EventBus.reset();
  jest.clearAllMocks();
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
// 1. getDateRangePresets
// ============================================================================

describe('getDateRangePresets', () => {
  test('returns an array', () => {
    const presets = getDateRangePresets();
    expect(Array.isArray(presets)).toBe(true);
  });

  test('includes last7, last30, ytd, and custom presets', () => {
    const presets = getDateRangePresets();
    const ids = presets.map(p => p.id);
    expect(ids).toContain('last7');
    expect(ids).toContain('last30');
    expect(ids).toContain('ytd');
    expect(ids).toContain('custom');
  });

  test('each preset has id, label, and days properties', () => {
    const presets = getDateRangePresets();
    for (const preset of presets) {
      expect(preset).toHaveProperty('id');
      expect(preset).toHaveProperty('label');
      expect(preset).toHaveProperty('days');
      expect(typeof preset.id).toBe('string');
      expect(typeof preset.label).toBe('string');
      expect(typeof preset.days).toBe('number');
    }
  });
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
// 6. getUserNotifications
// ============================================================================

describe('getUserNotifications', () => {
  test('returns empty array when no notifications exist', () => {
    const result = JSON.parse(getUserNotifications());
    expect(result).toEqual([]);
  });

  test('prunes notifications older than 30 days', () => {
    const oldDate = new Date();
    oldDate.setDate(oldDate.getDate() - 40);
    const recentDate = new Date().toISOString();

    const notifications = [
      { id: 'n1', timestamp: oldDate.toISOString(), read: false },
      { id: 'n2', timestamp: recentDate, read: false }
    ];
    const key = 'notifications_test_example_com';
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === key) return JSON.stringify(notifications);
      return null;
    });

    const result = JSON.parse(getUserNotifications());
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('n2');
  });

  test('returns a JSON string', () => {
    const result = getUserNotifications();
    expect(typeof result).toBe('string');
    expect(() => JSON.parse(result)).not.toThrow();
  });

  test('caps at 30 days (keeps recent notifications)', () => {
    const now = new Date();
    const within30 = new Date(now.getTime() - 25 * 24 * 60 * 60 * 1000).toISOString();
    const notifications = [
      { id: 'n1', timestamp: within30, read: false }
    ];
    const key = 'notifications_test_example_com';
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === key) return JSON.stringify(notifications);
      return null;
    });

    const result = JSON.parse(getUserNotifications());
    expect(result.length).toBe(1);
  });
});

// ============================================================================
// 7. pushNotification
// ============================================================================

describe('pushNotification', () => {
  test('adds notification to store', () => {
    const result = pushNotification('test@example.com', {
      title: 'Test', body: 'Hello', type: 'info'
    });
    expect(result.success).toBe(true);
    expect(result.id).toBeDefined();
  });

  test('auth check rejects non-steward cross-user push', () => {
    getUserRole_.mockImplementation(() => 'member');
    Session.getActiveUser.mockImplementation(() => ({
      getEmail: jest.fn(() => 'member@example.com')
    }));

    const result = pushNotification('other@example.com', {
      title: 'Test', body: 'Hello', type: 'info'
    });
    expect(result.success).toBe(false);
    expect(result.error).toContain('Not authorized');
  });

  test('creates notification with correct structure', () => {
    const result = pushNotification('test@example.com', {
      title: 'Alert', body: 'Important message', type: 'warning', link: 'https://example.com'
    });
    expect(result.success).toBe(true);

    // Verify the notification was stored correctly
    const key = 'notifications_test_example_com';
    const props = PropertiesService.getScriptProperties();
    const saveCall = props.setProperty.mock.calls.find(c => c[0] === key);
    expect(saveCall).toBeTruthy();
    const stored = JSON.parse(saveCall[1]);
    expect(stored[0].title).toBe('Alert');
    expect(stored[0].body).toBe('Important message');
    expect(stored[0].type).toBe('warning');
    expect(stored[0].link).toBe('https://example.com');
    expect(stored[0].read).toBe(false);
  });

  test('caps at 50 notifications', () => {
    const existing = [];
    for (let i = 0; i < 55; i++) {
      existing.push({ id: 'n_' + i, title: 'Old', body: 'old', type: 'info', timestamp: new Date().toISOString(), read: false });
    }
    const key = 'notifications_test_example_com';
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === key) return JSON.stringify(existing);
      return null;
    });

    pushNotification('test@example.com', { title: 'New', body: 'new', type: 'info' });

    const saveCall = props.setProperty.mock.calls.find(c => c[0] === key);
    const stored = JSON.parse(saveCall[1]);
    expect(stored.length).toBe(50);
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
// 8. markNotificationRead
// ============================================================================

describe('markNotificationRead', () => {
  test('marks notification as read', () => {
    const notifications = [
      { id: 'notif_123', title: 'Test', read: false }
    ];
    const key = 'notifications_test_example_com';
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(k => {
      if (k === key) return JSON.stringify(notifications);
      return null;
    });

    const result = markNotificationRead('notif_123');
    expect(result.success).toBe(true);

    const saveCall = props.setProperty.mock.calls.find(c => c[0] === key);
    const stored = JSON.parse(saveCall[1]);
    expect(stored[0].read).toBe(true);
  });

  test('handles missing notification gracefully', () => {
    const props = PropertiesService.getScriptProperties();
    props.getProperty.mockImplementation(() => '[]');

    const result = markNotificationRead('nonexistent');
    expect(result.success).toBe(true);
  });
});

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
    // Track pushNotification calls via event bus
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

describe('updateChartPreset', () => {
  test('updates specific fields of an existing preset', () => {
    const presets = [
      { id: 'preset_1', name: 'Old Name', visibleCharts: ['a'], chartOptions: {}, layout: { columns: 2 }, filters: {} }
    ];
    const props = PropertiesService.getUserProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'chart_presets') return JSON.stringify(presets);
      return null;
    });

    const result = updateChartPreset('preset_1', { name: 'New Name', layout: { columns: 4 } });
    expect(result.success).toBe(true);
    expect(result.preset.name).toBe('New Name');
    expect(result.preset.layout.columns).toBe(4);
    // Unchanged field should remain
    expect(result.preset.visibleCharts).toEqual(['a']);
  });

  test('returns error for nonexistent preset', () => {
    const props = PropertiesService.getUserProperties();
    props.getProperty.mockImplementation(() => '[]');

    const result = updateChartPreset('nonexistent', { name: 'X' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('not found');
  });

  test('emits preset:updated event', () => {
    const handler = jest.fn();
    EventBus.on('preset:updated', handler);

    const presets = [{ id: 'preset_1', name: 'A' }];
    const props = PropertiesService.getUserProperties();
    props.getProperty.mockImplementation(k => {
      if (k === 'chart_presets') return JSON.stringify(presets);
      return null;
    });

    updateChartPreset('preset_1', { name: 'B' });
    expect(handler).toHaveBeenCalledWith(expect.objectContaining({ presetId: 'preset_1' }));
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
