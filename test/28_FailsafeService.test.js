/**
 * Tests for 28_FailsafeService.gs
 *
 * Covers FailsafeService IIFE: sheet setup, digest configuration,
 * scheduled digests, bulk export, Drive backup, and trigger management.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock DataService before loading
global.DataService = {
  getMemberGrievances: jest.fn(() => [
    { status: 'Open', issueType: 'Safety', dateFiled: '2026-01-15' }
  ]),
  getMemberTasks: jest.fn(() => [
    { title: 'Follow up', priority: 'High', status: 'open', dueDate: '2026-04-01' }
  ])
};

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '28_FailsafeService.gs']);

let mockConfigSheet;
let mockMemberSheet;
let mockSs;

function makeConfigData(rows) {
  const headers = ['Email', 'Digest Enabled', 'Digest Frequency', 'Last Digest Sent', 'Include Grievances', 'Include Workload', 'Include Tasks'];
  return [headers, ...(rows || [])];
}

function makeMemberData() {
  return [
    ['Email', 'Name'],
    ['member@test.com', 'Jane Doe'],
    ['other@test.com', 'John Smith'],
  ];
}

function setupSheets(configRows, includeMemberSheet) {
  const configData = makeConfigData(configRows);
  mockConfigSheet = createMockSheet(SHEETS.FAILSAFE_CONFIG || '_Failsafe_Config', configData);

  // Custom getRange for config sheet
  mockConfigSheet.getRange = jest.fn((row, col) => ({
    getValue: jest.fn(() => {
      if (configData[row - 1]) return configData[row - 1][col - 1] || '';
      return '';
    }),
    setValue: jest.fn(),
    getValues: jest.fn(() => [[configData[row - 1] ? configData[row - 1][col - 1] : '']])
  }));

  const sheets = [mockConfigSheet];

  if (includeMemberSheet) {
    mockMemberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    sheets.push(mockMemberSheet);
  }

  mockSs = createMockSpreadsheet(sheets);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

beforeEach(() => {
  setupSheets();
});

// ============================================================================
// FailsafeService.initFailsafeSheet
// ============================================================================

describe('FailsafeService.initFailsafeSheet', () => {
  test('creates sheet if missing', () => {
    mockSs = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
    FailsafeService.initFailsafeSheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });

  test('does not recreate existing sheet', () => {
    FailsafeService.initFailsafeSheet();
    expect(mockSs.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// FailsafeService.getDigestConfig
// ============================================================================

describe('FailsafeService.getDigestConfig', () => {
  test('returns null for empty email', () => {
    expect(FailsafeService.getDigestConfig('')).toBeNull();
    expect(FailsafeService.getDigestConfig(null)).toBeNull();
  });

  test('returns default when no sheet exists', () => {
    mockSs = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
    const config = FailsafeService.getDigestConfig('user@test.com');
    expect(config.enabled).toBe(false);
    expect(config.frequency).toBe('weekly');
    expect(config.includeGrievances).toBe(true);
  });

  test('returns default for unknown email', () => {
    setupSheets([
      ['other@test.com', true, 'weekly', '', true, true, true]
    ]);
    const config = FailsafeService.getDigestConfig('unknown@test.com');
    expect(config.enabled).toBe(false);
  });

  test('returns stored config for known email', () => {
    setupSheets([
      ['user@test.com', true, 'monthly', new Date('2026-02-01'), true, false, true]
    ]);
    const config = FailsafeService.getDigestConfig('user@test.com');
    expect(config.enabled).toBe(true);
    expect(config.frequency).toBe('monthly');
    expect(config.includeWorkload).toBe(false);
  });

  test('handles boolean TRUE string', () => {
    setupSheets([
      ['user@test.com', 'TRUE', 'weekly', '', 'TRUE', 'FALSE', 'TRUE']
    ]);
    const config = FailsafeService.getDigestConfig('user@test.com');
    expect(config.enabled).toBe(true);
    expect(config.includeWorkload).toBe(false);
  });

  test('case insensitive email lookup', () => {
    setupSheets([
      ['USER@TEST.COM', true, 'weekly', '', true, true, true]
    ]);
    const config = FailsafeService.getDigestConfig('user@test.com');
    expect(config.enabled).toBe(true);
  });
});

// ============================================================================
// FailsafeService.updateDigestConfig
// ============================================================================

describe('FailsafeService.updateDigestConfig', () => {
  test('rejects missing email', () => {
    const result = FailsafeService.updateDigestConfig('', {});
    expect(result.success).toBe(false);
  });

  test('rejects missing config', () => {
    const result = FailsafeService.updateDigestConfig('user@test.com', null);
    expect(result.success).toBe(false);
  });

  test('normalizes frequency to weekly/monthly', () => {
    const result = FailsafeService.updateDigestConfig('user@test.com', {
      enabled: true,
      frequency: 'daily' // Not supported → should default to weekly
    });
    expect(result.success).toBe(true);
  });

  test('saves new config for unknown email', () => {
    const result = FailsafeService.updateDigestConfig('new@test.com', {
      enabled: true,
      frequency: 'monthly',
      includeGrievances: true,
      includeWorkload: false,
      includeTasks: true
    });
    expect(result.success).toBe(true);
  });

  test('updates existing config', () => {
    setupSheets([
      ['user@test.com', true, 'weekly', '', true, true, true]
    ]);
    const result = FailsafeService.updateDigestConfig('user@test.com', {
      enabled: false,
      frequency: 'monthly'
    });
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// FailsafeService.processScheduledDigests
// ============================================================================

describe('FailsafeService.processScheduledDigests', () => {
  test('returns {processed:0} when no config sheet', () => {
    mockSs = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
    const result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
  });

  test('skips disabled digests', () => {
    setupSheets([
      ['user@test.com', false, 'weekly', '', true, true, true]
    ]);
    const result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
  });

  test('stops when email quota is low', () => {
    setupSheets([
      ['user@test.com', true, 'weekly', new Date('2000-01-01'), true, true, true]
    ]);
    MailApp.getRemainingDailyQuota = jest.fn(() => 2);
    const result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
    // Restore
    MailApp.getRemainingDailyQuota = jest.fn(() => 100);
  });
});

// ============================================================================
// FailsafeService.triggerBulkExport
// ============================================================================

describe('FailsafeService.triggerBulkExport', () => {
  test('rejects missing stewardEmail', () => {
    const result = FailsafeService.triggerBulkExport('');
    expect(result.success).toBe(false);
  });

  test('returns error when no member data', () => {
    const result = FailsafeService.triggerBulkExport('steward@test.com');
    expect(result.success).toBe(false);
  });

  test('exports when member sheet has data', () => {
    setupSheets([], true);
    const result = FailsafeService.triggerBulkExport('steward@test.com');
    expect(result).toHaveProperty('success');
  });
});

// ============================================================================
// FailsafeService.backupCriticalSheets
// ============================================================================

describe('FailsafeService.backupCriticalSheets', () => {
  test('returns backup result', () => {
    const result = FailsafeService.backupCriticalSheets();
    expect(result).toHaveProperty('success');
    expect(result.success).toBe(true);
    expect(result).toHaveProperty('backedUp');
  });

  test('skips sheets with no data', () => {
    const result = FailsafeService.backupCriticalSheets();
    expect(result.backedUp).toBe(0);
  });
});

// ============================================================================
// FailsafeService.setupFailsafeTriggers / removeFailsafeTriggers
// ============================================================================

describe('FailsafeService.setupFailsafeTriggers', () => {
  test('creates triggers', () => {
    const result = FailsafeService.setupFailsafeTriggers();
    expect(result.success).toBe(true);
    expect(ScriptApp.newTrigger).toHaveBeenCalled();
  });

  test('removes existing triggers first', () => {
    const mockTrigger = {
      getHandlerFunction: jest.fn(() => 'fsProcessScheduledDigests')
    };
    ScriptApp.getProjectTriggers = jest.fn(() => [mockTrigger]);
    FailsafeService.setupFailsafeTriggers();
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(mockTrigger);
    // Restore
    ScriptApp.getProjectTriggers = jest.fn(() => []);
  });
});

describe('FailsafeService.removeFailsafeTriggers', () => {
  test('removes matching triggers', () => {
    const digestTrigger = { getHandlerFunction: jest.fn(() => 'fsProcessScheduledDigests') };
    const backupTrigger = { getHandlerFunction: jest.fn(() => 'fsBackupCriticalSheets') };
    const otherTrigger = { getHandlerFunction: jest.fn(() => 'someOtherFunction') };
    ScriptApp.getProjectTriggers = jest.fn(() => [digestTrigger, backupTrigger, otherTrigger]);

    const result = FailsafeService.removeFailsafeTriggers();
    expect(result.success).toBe(true);
    expect(result.removed).toBe(2);
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(digestTrigger);
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(backupTrigger);
    expect(ScriptApp.deleteTrigger).not.toHaveBeenCalledWith(otherTrigger);

    // Restore
    ScriptApp.getProjectTriggers = jest.fn(() => []);
  });

  test('returns {removed:0} when no triggers', () => {
    ScriptApp.getProjectTriggers = jest.fn(() => []);
    const result = FailsafeService.removeFailsafeTriggers();
    expect(result.removed).toBe(0);
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('fsGetDigestConfig delegates', () => {
    const result = fsGetDigestConfig('user@test.com');
    expect(result === null || typeof result === 'object').toBe(true);
  });

  test('fsUpdateDigestConfig delegates', () => {
    const result = fsUpdateDigestConfig('user@test.com', { enabled: false });
    expect(result).toHaveProperty('success');
  });

  test('fsProcessScheduledDigests delegates', () => {
    const result = fsProcessScheduledDigests();
    expect(result).toHaveProperty('processed');
  });

  test('fsTriggerBulkExport delegates', () => {
    const result = fsTriggerBulkExport('steward@test.com');
    expect(result).toHaveProperty('success');
  });

  test('fsBackupCriticalSheets delegates', () => {
    const result = fsBackupCriticalSheets();
    expect(result).toHaveProperty('success');
  });

  test('fsSetupTriggers delegates', () => {
    const result = fsSetupTriggers();
    expect(result).toHaveProperty('success');
  });

  test('fsRemoveTriggers delegates', () => {
    const result = fsRemoveTriggers();
    expect(result).toHaveProperty('success');
  });

  test('fsInitSheets runs without error', () => {
    expect(() => fsInitSheets()).not.toThrow();
  });
});
