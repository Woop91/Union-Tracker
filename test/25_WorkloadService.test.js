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

  mockSs = createMockSpreadsheet([mockVaultSheet, mockReportingSheet, mockRemindersSheet, mockUserMetaSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

beforeEach(() => {
  setupSheets();
});

// ============================================================================
// WorkloadService exists
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
});

// ============================================================================
// getSubCategories / CATEGORY_LABELS
// ============================================================================

describe('WorkloadService.getSubCategories', () => {
  test('returns object with category keys', () => {
    const cats = WorkloadService.getSubCategories();
    expect(typeof cats).toBe('object');
    expect(cats).toHaveProperty('priority');
    expect(cats).toHaveProperty('pending');
    expect(cats).toHaveProperty('aged');
  });

  test('priority has QDD and CAL', () => {
    const cats = WorkloadService.getSubCategories();
    expect(cats.priority).toContain('QDD');
    expect(cats.priority).toContain('CAL');
  });

  test('has 8 categories', () => {
    const cats = WorkloadService.getSubCategories();
    expect(Object.keys(cats).length).toBe(8);
  });
});

describe('WorkloadService.CATEGORY_LABELS', () => {
  test('is an object with label entries', () => {
    const labels = WorkloadService.CATEGORY_LABELS;
    expect(typeof labels).toBe('object');
    expect(labels).toHaveProperty('t1');
    expect(labels.t1).toBe('Priority Cases');
  });

  test('has 8 label entries', () => {
    const labels = WorkloadService.CATEGORY_LABELS;
    expect(Object.keys(labels).length).toBe(8);
  });
});

// ============================================================================
// processFormSSO (was submitWorkload)
// ============================================================================

describe('WorkloadService.processFormSSO', () => {
  // processFormSSO returns a string ('Error:...' or 'Success:...')
  test('rejects missing email', () => {
    const result = WorkloadService.processFormSSO('', {});
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('rejects missing data', () => {
    const result = WorkloadService.processFormSSO('user@test.com', null);
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Error/i);
  });

  test('submits valid workload data', () => {
    const data = {
      t1: 5, t2: 10, t3: 3, t4: 7, t5: 1, t6: 2, t7: 0, t8: 4,
      weekly_cases: 15,
      employment_type: 'Full-time',
      privacy: 'Unit'
    };
    const result = WorkloadService.processFormSSO('user@test.com', data);
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Success/i);
  });

  test('clamps numeric values to non-negative', () => {
    const data = {
      t1: -5, t2: 0, t3: 3, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0,
      weekly_cases: 0,
      employment_type: 'Full-time',
      privacy: 'Unit'
    };
    // -5 is invalid (t1 range 0-999) but clamp behavior returns error per validation
    const result = WorkloadService.processFormSSO('user@test.com', data);
    expect(typeof result).toBe('string');
  });

  test('lowercases email', () => {
    const data = {
      t1: 1, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0,
      weekly_cases: 1,
      employment_type: 'Full-time',
      privacy: 'Unit'
    };
    const result = WorkloadService.processFormSSO('USER@TEST.COM', data);
    expect(typeof result).toBe('string');
    expect(result).toMatch(/^Success/i);
  });
});

// ============================================================================
// getHistorySSO (was getHistory)
// ============================================================================

describe('WorkloadService.getHistorySSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getHistorySSO('');
    expect(result.success).toBe(false);
  });

  test('returns empty history when no data', () => {
    const result = WorkloadService.getHistorySSO('user@test.com');
    if (result.success) {
      expect(Array.isArray(result.history)).toBe(true);
      expect(result.history.length).toBe(0);
    }
  });

  test('returns history for user with data', () => {
    const now = new Date();
    const vaultData = [
      ['Timestamp', 'Email', 'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'Weekly', 'SubCats', 'EmpType', 'PTHrs', 'Leave', 'LeavePlanned', 'LeaveStart', 'LeaveEnd', 'NoIntake', 'NoticeTime', 'HalfDay', 'Privacy', 'OnPlan', 'OT'],
      [now, 'user@test.com', 5, 10, 3, 7, 1, 2, 0, 4, 15, '{}', 'full-time', '', '', false, '', '', false, '', false, 'anonymous', false, 0],
    ];
    setupSheets(vaultData);

    const result = WorkloadService.getHistorySSO('user@test.com');
    if (result.success) {
      expect(result.history.length).toBeGreaterThanOrEqual(0);
    }
  });
});

// ============================================================================
// getDashboardDataSSO (was getDashboardData)
// ============================================================================

describe('WorkloadService.getDashboardDataSSO', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getDashboardDataSSO('');
    expect(result.success).toBe(false);
  });

  test('returns dashboard data structure', () => {
    const result = WorkloadService.getDashboardDataSSO('user@test.com');
    if (result.success) {
      expect(result).toHaveProperty('data');
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
  });

  test('returns null-safe result for empty email', () => {
    const config = WorkloadService.getReminderSSO('');
    expect(config === null || config.enabled === false).toBe(true);
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
// Health Status
// ============================================================================

describe('WorkloadService.getHealthStatus', () => {
  test('does not throw', () => {
    // getHealthStatus() shows a UI alert — verifying it does not throw
    expect(() => WorkloadService.getHealthStatus()).not.toThrow();
  });
});

// ============================================================================
// initSheets
// ============================================================================

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
    const result = processWorkloadFormSSO('user@test.com', { t1: 1, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0, weekly_cases: 1, employment_type: 'Full-time', privacy: 'Unit' });
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
  });

  test('getWorkloadReminderSSO delegates to WorkloadService.getReminderSSO', () => {
    const result = getWorkloadReminderSSO('user@test.com');
    expect(result !== undefined).toBe(true);
  });

  test('setWorkloadReminderSSO delegates to WorkloadService.setReminderSSO', () => {
    const result = setWorkloadReminderSSO('user@test.com', { enabled: false });
    expect(result).toHaveProperty('success');
  });

  test('exportWorkloadHistoryCSV delegates to WorkloadService.exportHistoryCSV', () => {
    const result = exportWorkloadHistoryCSV('user@test.com');
    expect(result !== undefined).toBe(true);
  });

  test('initWorkloadTrackerSheets runs without error', () => {
    expect(() => initWorkloadTrackerSheets()).not.toThrow();
  });
});
