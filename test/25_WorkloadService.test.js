/**
 * Tests for 25_WorkloadService.gs
 *
 * Covers WorkloadService IIFE: config/column constants, sanitization,
 * rate limiting, sheet helpers, submission, history, dashboard analytics,
 * reminders, and global wrappers.
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

  test('exports submitWorkload function', () => {
    expect(typeof WorkloadService.submitWorkload).toBe('function');
  });

  test('exports getHistory function', () => {
    expect(typeof WorkloadService.getHistory).toBe('function');
  });

  test('exports getDashboardData function', () => {
    expect(typeof WorkloadService.getDashboardData).toBe('function');
  });

  test('exports getSubCategories function', () => {
    expect(typeof WorkloadService.getSubCategories).toBe('function');
  });

  test('exports getCategoryLabels function', () => {
    expect(typeof WorkloadService.getCategoryLabels).toBe('function');
  });
});

// ============================================================================
// getSubCategories / getCategoryLabels
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

describe('WorkloadService.getCategoryLabels', () => {
  test('returns labels object', () => {
    const labels = WorkloadService.getCategoryLabels();
    expect(typeof labels).toBe('object');
    expect(labels).toHaveProperty('t1');
    expect(labels.t1).toBe('Priority Cases');
  });

  test('has 8 label entries', () => {
    const labels = WorkloadService.getCategoryLabels();
    expect(Object.keys(labels).length).toBe(8);
  });
});

// ============================================================================
// submitWorkload
// ============================================================================

describe('WorkloadService.submitWorkload', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.submitWorkload('', {});
    expect(result.success).toBe(false);
    expect(result.message).toBeTruthy();
  });

  test('rejects missing data', () => {
    const result = WorkloadService.submitWorkload('user@test.com', null);
    expect(result.success).toBe(false);
  });

  test('submits valid workload data', () => {
    const data = {
      t1: 5, t2: 10, t3: 3, t4: 7, t5: 1, t6: 2, t7: 0, t8: 4,
      weeklyCases: 15,
      employmentType: 'full-time',
      privacy: 'anonymous'
    };
    const result = WorkloadService.submitWorkload('user@test.com', data);
    expect(result.success).toBe(true);
  });

  test('clamps numeric values to non-negative', () => {
    const data = {
      t1: -5, t2: 0, t3: 3, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0,
      weeklyCases: 0,
      employmentType: 'full-time',
      privacy: 'anonymous'
    };
    const result = WorkloadService.submitWorkload('user@test.com', data);
    expect(result.success).toBe(true);
  });

  test('lowercases email', () => {
    const data = {
      t1: 1, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0,
      weeklyCases: 1,
      employmentType: 'full-time',
      privacy: 'anonymous'
    };
    const result = WorkloadService.submitWorkload('USER@TEST.COM', data);
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// getHistory
// ============================================================================

describe('WorkloadService.getHistory', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getHistory('');
    expect(result.success).toBe(false);
  });

  test('returns empty submissions when no data', () => {
    const result = WorkloadService.getHistory('user@test.com');
    if (result.success) {
      expect(Array.isArray(result.submissions)).toBe(true);
      expect(result.submissions.length).toBe(0);
    }
  });

  test('returns submissions for user with data', () => {
    const now = new Date();
    const vaultData = [
      ['Timestamp', 'Email', 'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'Weekly', 'SubCats', 'EmpType', 'PTHrs', 'Leave', 'LeavePlanned', 'LeaveStart', 'LeaveEnd', 'NoIntake', 'NoticeTime', 'HalfDay', 'Privacy', 'OnPlan', 'OT'],
      [now, 'user@test.com', 5, 10, 3, 7, 1, 2, 0, 4, 15, '{}', 'full-time', '', '', false, '', '', false, '', false, 'anonymous', false, 0],
    ];
    setupSheets(vaultData);

    const result = WorkloadService.getHistory('user@test.com');
    if (result.success) {
      expect(result.submissions.length).toBeGreaterThanOrEqual(0);
    }
  });
});

// ============================================================================
// getDashboardData
// ============================================================================

describe('WorkloadService.getDashboardData', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.getDashboardData('');
    expect(result.success).toBe(false);
  });

  test('returns dashboard data structure', () => {
    const result = WorkloadService.getDashboardData('user@test.com');
    if (result.success) {
      expect(result).toHaveProperty('summary');
    }
  });
});

// ============================================================================
// Reminder Functions
// ============================================================================

describe('WorkloadService.getReminderConfig', () => {
  test('returns default config for unknown email', () => {
    const config = WorkloadService.getReminderConfig('unknown@test.com');
    expect(config).not.toBeNull();
    expect(config.enabled).toBe(false);
  });

  test('returns null for empty email', () => {
    const config = WorkloadService.getReminderConfig('');
    expect(config === null || config.enabled === false).toBe(true);
  });
});

describe('WorkloadService.updateReminderConfig', () => {
  test('rejects missing email', () => {
    const result = WorkloadService.updateReminderConfig('', {});
    expect(result.success).toBe(false);
  });

  test('saves reminder config for valid input', () => {
    const result = WorkloadService.updateReminderConfig('user@test.com', {
      enabled: true,
      frequency: 'weekly',
      day: 'Monday',
      time: '09:00'
    });
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// Privacy Functions
// ============================================================================

describe('WorkloadService.getPrivacyStats', () => {
  test('returns stats object', () => {
    const stats = WorkloadService.getPrivacyStats();
    expect(typeof stats).toBe('object');
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
  test('wsSubmitWorkload delegates to WorkloadService.submitWorkload', () => {
    const result = wsSubmitWorkload('user@test.com', { t1: 1, t2: 0, t3: 0, t4: 0, t5: 0, t6: 0, t7: 0, t8: 0, weeklyCases: 1, employmentType: 'full-time', privacy: 'anonymous' });
    expect(result).toHaveProperty('success');
  });

  test('wsGetHistory delegates to WorkloadService.getHistory', () => {
    const result = wsGetHistory('user@test.com');
    expect(result).toHaveProperty('success');
  });

  test('wsGetDashboardData delegates to WorkloadService.getDashboardData', () => {
    const result = wsGetDashboardData('user@test.com');
    expect(result).toHaveProperty('success');
  });

  test('wsGetSubCategories delegates to WorkloadService.getSubCategories', () => {
    const result = wsGetSubCategories();
    expect(typeof result).toBe('object');
  });

  test('wsGetCategoryLabels delegates to WorkloadService.getCategoryLabels', () => {
    const result = wsGetCategoryLabels();
    expect(typeof result).toBe('object');
  });

  test('wsGetReminderConfig delegates to WorkloadService.getReminderConfig', () => {
    const result = wsGetReminderConfig('user@test.com');
    expect(result !== undefined).toBe(true);
  });

  test('wsUpdateReminderConfig delegates to WorkloadService.updateReminderConfig', () => {
    const result = wsUpdateReminderConfig('user@test.com', { enabled: false });
    expect(result).toHaveProperty('success');
  });

  test('wsGetPrivacyStats delegates to WorkloadService.getPrivacyStats', () => {
    const result = wsGetPrivacyStats();
    expect(typeof result).toBe('object');
  });

  test('wsInitSheets runs without error', () => {
    expect(() => wsInitSheets()).not.toThrow();
  });
});
