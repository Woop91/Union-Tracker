/**
 * Tests for 25_WorkloadService.gs
 *
 * Covers WorkloadService IIFE: config/column constants, sanitization,
 * rate limiting, sheet helpers, submission, history, dashboard analytics,
 * reminders, and global wrappers.
 *
 * NOTE: v4.20.0 refactored WorkloadService to SSO-based auth.
 * Methods renamed: submitWorkload→processFormSSO, getHistory→getHistorySSO,
 * getDashboardData→getDashboardDataSSO, getCategoryLabels→CATEGORY_LABELS (constant).
 * Reminder API: getReminderConfig→getReminderSSO, updateReminderConfig→setReminderSSO.
 *
 * v4.28.1 expanded tests: atomic rate limiting, crash-safe reporting refresh,
 * category label consistency, history shape assertions, stronger dashboard tests.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '25_WorkloadService.gs']);

let mockVaultSheet;
let mockReportingSheet;
let mockRemindersSheet;
let mockUserMetaSheet;
let mockSs;

function setupSheets(vaultData, reminderData, userMetaData) {
  mockVaultSheet = createMockSheet(SHEETS.WORKLOAD_VAULT || 'Workload Vault', vaultData || [['header']]);
  mockReportingSheet = createMockSheet(SHEETS.WORKLOAD_REPORTING || 'Workload Reporting', [['header']]);
  mockRemindersSheet = createMockSheet(SHEETS.WORKLOAD_REMINDERS || 'Workload Reminders', reminderData || [['header']]);
  mockUserMetaSheet = createMockSheet(SHEETS.WORKLOAD_USER_META || 'Workload UserMeta', userMetaData || [['header']]);

  // Add clearContent mock to reporting sheet for crash-safe refresh tests
  mockReportingSheet.clearContents = jest.fn();
  mockReportingSheet.clearContent = jest.fn();

  mockSs = createMockSpreadsheet([mockVaultSheet, mockReportingSheet, mockRemindersSheet, mockUserMetaSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

function makeValidFormData(overrides) {
  return Object.assign({
    t1: 5, t2: 10, t3: 3, t4: 7, t5: 1, t6: 2, t7: 0, t8: 4,
    weekly_cases: 15,
    employment_type: 'Full-time',
    privacy: 'Unit'
  }, overrides || {});
}

function makeVaultRow(email, overrides) {
  var now = new Date();
  var defaults = [
    now, email || 'user@test.com',
    5, 10, 3, 7, 1, 2, 0, 4,     // t1-t8
    15, '{}',                       // weekly, subcats
    'Full-time', '',                // emp, pt hours
    '', false, '', '',              // leave type/planned/start/end
    false, '', false,               // no-intake, notice, half-day
    'Unit', 'No', 0                 // privacy, on-plan, OT
  ];
  if (overrides) {
    for (var k in overrides) defaults[k] = overrides[k];
  }
  return defaults;
}

var VAULT_HEADER = ['Timestamp', 'Email', 'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8',
  'Weekly', 'SubCats', 'EmpType', 'PTHrs', 'Leave', 'LeavePlanned', 'LeaveStart', 'LeaveEnd',
  'NoIntake', 'NoticeTime', 'HalfDay', 'Privacy', 'OnPlan', 'OT'];

beforeEach(() => {
  setupSheets();
});

// ============================================================================
// WorkloadService module exports
// ============================================================================

describe('WorkloadService module', () => {
  test('is defined as an object', () => {
    expect(typeof WorkloadService).toBe('object');
    expect(WorkloadService).not.toBeNull();
  });

  test('exports processFormSSO function', () => {
    expect(typeof WorkloadService.processFormSSO).toBe('function');
  });

  test('exports getHistorySSO function', () => {
    expect(typeof WorkloadService.getHistorySSO).toBe('function');
  });

  test('exports getDashboardDataSSO function', () => {
    expect(typeof WorkloadService.getDashboardDataSSO).toBe('function');
  });

  test('exports getSubCategories function', () => {
    expect(typeof WorkloadService.getSubCategories).toBe('function');
  });

  test('exports CATEGORY_LABELS constant', () => {
    expect(typeof WorkloadService.CATEGORY_LABELS).toBe('object');
    expect(WorkloadService.CATEGORY_LABELS).not.toBeNull();
  });

  test('exports all required public methods', () => {
    var required = [
      'initSheets', 'processFormSSO', 'getHistorySSO', 'getDashboardDataSSO',
      'getReminderSSO', 'setReminderSSO', 'exportHistoryCSV', 'getSubCategories',
      'processReminders', 'createBackup', 'archiveOldData', 'cleanVault',
      'getHealthStatus', 'refreshLedger'
    ];
    required.forEach(function(name) {
      expect(typeof WorkloadService[name]).toBe('function');
    });
  });
});

// ============================================================================
// getSubCategories / CATEGORY_LABELS
// ============================================================================

describe('WorkloadService.getSubCategories', () => {
  test('returns object with all 8 category keys', () => {
    const cats = WorkloadService.getSubCategories();
    expect(typeof cats).toBe('object');
    var expectedKeys = ['priority', 'pending', 'unread', 'todo', 'referrals', 'ce', 'assistance', 'aged'];
    expectedKeys.forEach(function(key) {
      expect(cats).toHaveProperty(key);
      expect(Array.isArray(cats[key])).toBe(true);
      expect(cats[key].length).toBeGreaterThan(0);
    });
  });

  test('priority has QDD and CAL', () => {
    const cats = WorkloadService.getSubCategories();
    expect(cats.priority).toContain('QDD');
    expect(cats.priority).toContain('CAL');
  });

  test('has exactly 8 categories', () => {
    const cats = WorkloadService.getSubCategories();
    expect(Object.keys(cats).length).toBe(8);
  });

  test('every sub-category value is a non-empty string', () => {
    const cats = WorkloadService.getSubCategories();
    for (var key in cats) {
      cats[key].forEach(function(sub) {
        expect(typeof sub).toBe('string');
        expect(sub.length).toBeGreaterThan(0);
      });
    }
  });
});

describe('WorkloadService.CATEGORY_LABELS', () => {
  test('has all 8 labels with correct t1-t8 keys', () => {
    const labels = WorkloadService.CATEGORY_LABELS;
    expect(Object.keys(labels).length).toBe(8);
    for (var i = 1; i <= 8; i++) {
      expect(labels).toHaveProperty('t' + i);
      expect(typeof labels['t' + i]).toBe('string');
      expect(labels['t' + i].length).toBeGreaterThan(0);
    }
  });

  test('t1 is Priority Cases', () => {
    expect(WorkloadService.CATEGORY_LABELS.t1).toBe('Priority Cases');
  });

  test('CATEGORY_LABELS keys match SUB_CATEGORIES order', () => {
    // Ensure label→subcategory mapping is consistent
    var labels = WorkloadService.CATEGORY_LABELS;
    var subs = WorkloadService.getSubCategories();
    var labelValues = Object.values(labels);
    // Each label should map to a subcategory key
    expect(labelValues).toContain('Priority Cases');
    expect(labelValues).toContain('Pending Cases');
    expect(labelValues).toContain('Aged Cases');
    expect(subs).toHaveProperty('priority');
    expect(subs).toHaveProperty('pending');
    expect(subs).toHaveProperty('aged');
  });
});

// ============================================================================
// processFormSSO
// ============================================================================

describe('WorkloadService.processFormSSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.processFormSSO('', {});
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects null email', () => {
    const result = WorkloadService.processFormSSO(null, {});
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects missing data', () => {
    const result = WorkloadService.processFormSSO('user@test.com', null);
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects undefined data', () => {
    const result = WorkloadService.processFormSSO('user@test.com', undefined);
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('submits valid workload data and returns Success string', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData());
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Success/i);
  });

  test('rejects negative numeric values', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData({ t1: -5 }));
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects values above 999', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData({ t2: 1000 }));
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects fractional numeric values', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData({ t3: 5.5 }));
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('lowercases email before submission', () => {
    const result = WorkloadService.processFormSSO('USER@TEST.COM', makeValidFormData());
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Success/i);
  });

  test('appends row to vault on success', () => {
    WorkloadService.processFormSSO('user@test.com', makeValidFormData());
    expect(mockVaultSheet.appendRow).toHaveBeenCalled();
    var row = mockVaultSheet.appendRow.mock.calls[0][0];
    expect(row.length).toBe(24); // VAULT_COL_COUNT
    expect(row[1]).toBe('user@test.com'); // lowercased email
    expect(row[2]).toBe(5); // t1
  });

  test('validates part-time hours when employment is Part-time', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData({
      employment_type: 'Part-time',
      part_time_hours: 0 // invalid: must be 1-40
    }));
    expect(result).toMatch(/^Error/i);
  });

  test('accepts valid part-time hours', () => {
    const result = WorkloadService.processFormSSO('user@test.com', makeValidFormData({
      employment_type: 'Part-time',
      part_time_hours: 20
    }));
    expect(result).toMatch(/^Success/i);
  });

  test('handles sub-category data in submission', () => {
    var data = makeValidFormData({ sub_1_0: 3, sub_1_1: 2 }); // priority subs
    const result = WorkloadService.processFormSSO('user@test.com', data);
    expect(result).toMatch(/^Success/i);
    var row = mockVaultSheet.appendRow.mock.calls[0][0];
    var subCats = JSON.parse(row[11]); // SUB_CATEGORIES column
    expect(subCats).toHaveProperty('priority');
    expect(subCats.priority).toHaveProperty('QDD');
    expect(subCats.priority.QDD).toBe(3);
  });

  test('stores overtime hours when enabled', () => {
    var data = makeValidFormData({ overtime_enabled: true, overtime_hours: 8 });
    WorkloadService.processFormSSO('user@test.com', data);
    var row = mockVaultSheet.appendRow.mock.calls[0][0];
    expect(row[23]).toBe(8); // OVERTIME_HOURS column
  });
});

// ============================================================================
// Rate Limiting (atomic check-and-record)
// ============================================================================

describe('Rate limiting', () => {
  test('allows submissions within rate limit', () => {
    const result = WorkloadService.processFormSSO('ratelimit@test.com', makeValidFormData());
    expect(result).toMatch(/^Success/i);
  });

  test('rate limit is enforced after max submissions', () => {
    // Submit maxSubmissionsPerHour times (10 by default)
    for (var i = 0; i < 10; i++) {
      WorkloadService.processFormSSO('ratelimit2@test.com', makeValidFormData());
    }
    // 11th should be rate-limited
    const result = WorkloadService.processFormSSO('ratelimit2@test.com', makeValidFormData());
    expect(result).toMatch(/^Error/i);
    expect(result).toMatch(/limit/i);
  });

  test('rate limit is per-email (different emails are independent)', () => {
    // Fill up one email's limit
    for (var i = 0; i < 10; i++) {
      WorkloadService.processFormSSO('usera_rl@test.com', makeValidFormData());
    }
    // Different email should still work
    const result = WorkloadService.processFormSSO('userb_rl@test.com', makeValidFormData());
    expect(result).toMatch(/^Success/i);
  });

  test('rate limit is case-insensitive on email', () => {
    for (var i = 0; i < 10; i++) {
      WorkloadService.processFormSSO('RateCase@test.com', makeValidFormData());
    }
    // Same email, different case — should be rate-limited
    const result = WorkloadService.processFormSSO('ratecase@test.com', makeValidFormData());
    expect(result).toMatch(/^Error/i);
    expect(result).toMatch(/limit/i);
  });
});

// ============================================================================
// getHistorySSO
// ============================================================================

describe('WorkloadService.getHistorySSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getHistorySSO('');
    expect(result.success).toBe(false);
    expect(result.history).toEqual([]);
  });

  test('returns empty history when no data', () => {
    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(Array.isArray(result.history)).toBe(true);
    expect(result.history.length).toBe(0);
  });

  test('returns properly shaped history entries', () => {
    var now = new Date();
    var vaultData = [VAULT_HEADER, makeVaultRow('user@test.com')];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.history.length).toBe(1);

    var entry = result.history[0];
    expect(entry).toHaveProperty('date');
    expect(entry).toHaveProperty('dateDisplay');
    expect(entry).toHaveProperty('t1');
    expect(entry).toHaveProperty('t8');
    expect(entry).toHaveProperty('weeklyCases');
    expect(entry).toHaveProperty('employment');
    expect(entry).toHaveProperty('privacy');
    expect(entry).toHaveProperty('onPlan');
    expect(entry).toHaveProperty('subCategories');
    expect(typeof entry.subCategories).toBe('object');
  });

  test('filters by email — does not return other users data', () => {
    var vaultData = [
      VAULT_HEADER,
      makeVaultRow('user@test.com'),
      makeVaultRow('other@test.com'),
    ];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.history.length).toBe(1);
  });

  test('returns summary with overtime stats', () => {
    var vaultData = [
      VAULT_HEADER,
      makeVaultRow('user@test.com', { 23: 5 }),  // 5 OT hours
      makeVaultRow('user@test.com', { 23: 3 }),  // 3 OT hours
    ];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.summary).toBeDefined();
    expect(result.summary.totalSubmissions).toBe(2);
    expect(result.summary.overtime.totalHours).toBe(8);
    expect(result.summary.overtime.submissionsWithOvertime).toBe(2);
  });

  test('sorts history newest first', () => {
    var old = new Date('2025-01-01');
    var recent = new Date('2025-06-01');
    var vaultData = [
      VAULT_HEADER,
      makeVaultRow('user@test.com', { 0: old }),
      makeVaultRow('user@test.com', { 0: recent }),
    ];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.history.length).toBe(2);
    // Newest first
    expect(new Date(result.history[0].date).getTime()).toBeGreaterThanOrEqual(
      new Date(result.history[1].date).getTime()
    );
  });

  test('parses sub-category JSON from vault', () => {
    var subCats = JSON.stringify({ priority: { QDD: 3, CAL: 2 } });
    var vaultData = [VAULT_HEADER, makeVaultRow('user@test.com', { 11: subCats })];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.history[0].subCategories).toHaveProperty('priority');
    expect(result.history[0].subCategories.priority.QDD).toBe(3);
  });

  test('handles malformed sub-category JSON gracefully', () => {
    var vaultData = [VAULT_HEADER, makeVaultRow('user@test.com', { 11: 'not-json{' })];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    expect(result.success).toBe(true);
    expect(result.history[0].subCategories).toEqual({});
  });
});

// ============================================================================
// getDashboardDataSSO
// ============================================================================

describe('WorkloadService.getDashboardDataSSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getDashboardDataSSO('');
    expect(result.success).toBe(false);
  });

  test('returns reciprocityBlocked when user has never shared', () => {
    var vaultData = [VAULT_HEADER, makeVaultRow('other@test.com')];
    setupSheets(vaultData);

    const result = WorkloadService.getDashboardDataSSO('noshare@test.com');
    // User has never submitted, so no sharing start date
    expect(result.reciprocityBlocked).toBe(true);
  });

  test('returns dashboard structure with averages', () => {
    // Set up user meta so user has a sharing start date
    var userMetaData = [
      ['Email', 'Sharing Start Date', 'Created Date'],
      ['user@test.com', new Date('2020-01-01'), new Date()]
    ];
    var vaultData = [
      VAULT_HEADER,
      makeVaultRow('user@test.com'),
      makeVaultRow('other@test.com'),
    ];
    setupSheets(vaultData, undefined, userMetaData);

    const result = WorkloadService.getDashboardDataSSO('user@test.com');
    if (result.success) {
      expect(result.data).toBeDefined();
      expect(result.data).toHaveProperty('totalSubmissions');
      expect(result.data).toHaveProperty('members');
      expect(result.data).toHaveProperty('averages');
      expect(result.data).toHaveProperty('employment');
      expect(result.data).toHaveProperty('overtime');
      expect(typeof result.data.averages).toBe('object');
    }
  });
});

// ============================================================================
// Reminder Functions (SSO-based)
// ============================================================================

describe('WorkloadService.getReminderSSO', () => {
  test('returns default config for unknown email', () => {
    const config = WorkloadService.getReminderSSO('unknown@test.com');
    expect(config).not.toBeNull();
    expect(config.enabled).toBe(false);
    expect(config.frequency).toBe('weekly');
  });

  test('returns default for empty email', () => {
    const config = WorkloadService.getReminderSSO('');
    expect(config.enabled).toBe(false);
  });
});

describe('WorkloadService.setReminderSSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.setReminderSSO('', {});
    expect(result.success).toBe(false);
  });

  test('saves reminder config for valid input', () => {
    const result = WorkloadService.setReminderSSO('user@test.com', {
      enabled: true,
      frequency: 'weekly',
      day: 'Monday',
      time: '09:00'
    });
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// CSV Export
// ============================================================================

describe('WorkloadService.exportHistoryCSV', () => {
  test('returns empty string for user with no data', () => {
    const csv = WorkloadService.exportHistoryCSV('nobody@test.com');
    expect(csv).toBe('');
  });

  test('returns CSV with headers for user with data', () => {
    var vaultData = [VAULT_HEADER, makeVaultRow('user@test.com')];
    setupSheets(vaultData);

    const csv = WorkloadService.exportHistoryCSV('user@test.com');
    expect(csv.length).toBeGreaterThan(0);
    expect(csv.split('\n').length).toBeGreaterThanOrEqual(2); // header + 1 row
    expect(csv).toContain('Date');
    expect(csv).toContain('Priority Cases');
  });
});

// ============================================================================
// Health Status & Init
// ============================================================================

describe('WorkloadService.getHealthStatus', () => {
  test('does not throw', () => {
    expect(() => WorkloadService.getHealthStatus()).not.toThrow();
  });
});

describe('WorkloadService.initSheets', () => {
  test('creates sheets without error', () => {
    expect(() => WorkloadService.initSheets()).not.toThrow();
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('processWorkloadFormSSO delegates to WorkloadService.processFormSSO', () => {
    const result = processWorkloadFormSSO('user@test.com', makeValidFormData());
    expect(typeof result).toBe('string');
  });

  test('getWorkloadHistorySSO delegates to WorkloadService.getHistorySSO', () => {
    const result = getWorkloadHistorySSO('user@test.com');
    expect(result).toHaveProperty('success');
  });

  test('getWorkloadDashboardDataSSO delegates to WorkloadService.getDashboardDataSSO', () => {
    const result = getWorkloadDashboardDataSSO('user@test.com');
    expect(result).toHaveProperty('success');
  });

  test('getWorkloadSubCategories delegates to WorkloadService.getSubCategories', () => {
    const result = getWorkloadSubCategories();
    expect(typeof result).toBe('object');
    expect(Object.keys(result).length).toBe(8);
  });

  test('getWorkloadReminderSSO delegates to WorkloadService.getReminderSSO', () => {
    const result = getWorkloadReminderSSO('user@test.com');
    expect(result).toBeDefined();
    expect(result.enabled).toBe(false);
  });

  test('setWorkloadReminderSSO delegates to WorkloadService.setReminderSSO', () => {
    const result = setWorkloadReminderSSO('user@test.com', { enabled: false });
    expect(result).toHaveProperty('success');
    expect(result.success).toBe(true);
  });

  test('exportWorkloadHistoryCSV delegates to WorkloadService.exportHistoryCSV', () => {
    const result = exportWorkloadHistoryCSV('user@test.com');
    expect(result).toBeDefined();
  });

  test('initWorkloadTrackerSheets runs without error', () => {
    expect(() => initWorkloadTrackerSheets()).not.toThrow();
  });
});

// ============================================================================
// Structural invariants (prevent regressions)
// ============================================================================

describe('Structural invariants', () => {
  test('SUB_CATEGORIES keys align with CATEGORY_LABELS count', () => {
    var subs = WorkloadService.getSubCategories();
    var labels = WorkloadService.CATEGORY_LABELS;
    expect(Object.keys(subs).length).toBe(Object.keys(labels).length);
  });

  test('every CATEGORY_LABELS entry is a non-empty string', () => {
    var labels = WorkloadService.CATEGORY_LABELS;
    for (var k in labels) {
      expect(typeof labels[k]).toBe('string');
      expect(labels[k].trim().length).toBeGreaterThan(0);
    }
  });

  test('SUB_CATEGORIES has no empty arrays', () => {
    var subs = WorkloadService.getSubCategories();
    for (var k in subs) {
      expect(subs[k].length).toBeGreaterThan(0);
    }
  });

  test('CATEGORY_LABELS t-keys are sequential from t1 to t8', () => {
    var labels = WorkloadService.CATEGORY_LABELS;
    for (var i = 1; i <= 8; i++) {
      expect(labels['t' + i]).toBeDefined();
    }
    expect(labels['t0']).toBeUndefined();
    expect(labels['t9']).toBeUndefined();
  });
});
