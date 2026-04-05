/**
 * Tests for 28_FailsafeService.gs
 *
 * Covers the FailsafeService IIFE: initFailsafeSheet, getDigestConfig,
 * updateDigestConfig, processScheduledDigests,
 * backupCriticalSheets, setupFailsafeTriggers, removeFailsafeTriggers,
 * and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '28_FailsafeService.gs']);

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

/** Standard failsafe config header row */
var FAILSAFE_HEADERS = [
  'Email', 'Digest Enabled', 'Digest Frequency', 'Last Digest Sent',
  'Include Grievances', 'Include Workload', 'Include Tasks'
];

/** Build a data set with headers + config rows */
function buildFailsafeData(rows) {
  return [FAILSAFE_HEADERS].concat(rows || []);
}

/** Create a mock failsafe config sheet with custom data rows */
function buildFailsafeSheet(rows) {
  var data = buildFailsafeData(rows);
  var sheet = createMockSheet(SHEETS.FAILSAFE_CONFIG, data);

  sheet.getRange = jest.fn(function (row, col, numRows, numCols) {
    return {
      getValue: jest.fn(function () {
        if (data[row - 1] && data[row - 1][col - 1] !== undefined) return data[row - 1][col - 1];
        return '';
      }),
      getValues: jest.fn(function () {
        var result = [];
        for (var r = row - 1; r < row - 1 + (numRows || 1); r++) {
          var rowArr = [];
          for (var c = col - 1; c < col - 1 + (numCols || 1); c++) {
            rowArr.push(data[r] && data[r][c] !== undefined ? data[r][c] : '');
          }
          result.push(rowArr);
        }
        return result;
      }),
      setValue: jest.fn(),
      setValues: jest.fn()
    };
  });

  return sheet;
}

/** Create a mock member directory sheet */
function buildMemberSheet(rows) {
  var headers = ['Name', 'Email', 'Department'];
  var data = [headers].concat(rows || []);
  var sheet = createMockSheet(SHEETS.MEMBER_DIR, data);
  return sheet;
}

/** Install a spreadsheet mock with the given sheets */
function installSS(sheets) {
  var ss = createMockSpreadsheet(sheets || []);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

// ---------------------------------------------------------------------------
// Global mocks
// ---------------------------------------------------------------------------

beforeEach(() => {
  jest.clearAllMocks();
  global.logAuditEvent = jest.fn();
  global.logIntegrityEvent = jest.fn();
  global.DataService = undefined;

  // Reset MailApp quota mock
  MailApp.getRemainingDailyQuota.mockReturnValue(100);
  MailApp.sendEmail.mockImplementation(() => {});

  // Reset DriveApp mocks
  var mockFolder = {
    getId: jest.fn(() => 'backup-folder-id'),
    getUrl: jest.fn(() => 'https://drive.google.com/backup-folder'),
    createFile: jest.fn(() => ({ getId: jest.fn(() => 'file-id') })),
    setDescription: jest.fn(),
    getFiles: jest.fn(() => ({ hasNext: jest.fn(() => false) }))
  };
  DriveApp.getFoldersByName.mockReturnValue({
    hasNext: jest.fn(() => false)
  });
  DriveApp.createFolder.mockReturnValue(mockFolder);
});

// ============================================================================
// FailsafeService.initFailsafeSheet
// ============================================================================

describe('FailsafeService.initFailsafeSheet', () => {
  test('creates sheet if missing', () => {
    var ss = installSS([]);
    var result = FailsafeService.initFailsafeSheet();

    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.FAILSAFE_CONFIG);
    expect(result).toBeTruthy();
  });

  test('does not recreate if sheet exists', () => {
    var existingSheet = buildFailsafeSheet([]);
    var ss = installSS([existingSheet]);
    var result = FailsafeService.initFailsafeSheet();

    expect(ss.insertSheet).not.toHaveBeenCalled();
    expect(result).toBe(existingSheet);
  });
});

// ============================================================================
// FailsafeService.getDigestConfig
// ============================================================================

describe('FailsafeService.getDigestConfig', () => {
  test('returns null for null email', () => {
    var result = FailsafeService.getDigestConfig(null);
    expect(result).toBeNull();
  });

  test('returns default when no config sheet exists', () => {
    installSS([]);
    var result = FailsafeService.getDigestConfig('user@test.com');
    expect(result.enabled).toBe(false);
    expect(result.frequency).toBe('weekly');
    expect(result.includeGrievances).toBe(true);
    expect(result.includeWorkload).toBe(true);
    expect(result.includeTasks).toBe(true);
  });

  test('returns stored config for known email', () => {
    var lastSent = new Date('2026-02-20T00:00:00');
    var sheet = buildFailsafeSheet([
      ['admin@test.com', true, 'monthly', lastSent, true, false, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.getDigestConfig('admin@test.com');
    expect(result.enabled).toBe(true);
    expect(result.frequency).toBe('monthly');
    expect(result.includeGrievances).toBe(true);
    expect(result.includeWorkload).toBe(false);
    expect(result.includeTasks).toBe(true);
    expect(result.lastSent).toBeTruthy();
  });

  test('returns default for unknown email', () => {
    var sheet = buildFailsafeSheet([
      ['other@test.com', true, 'weekly', '', true, true, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.getDigestConfig('unknown@test.com');
    expect(result.enabled).toBe(false);
    expect(result.frequency).toBe('weekly');
  });

  test('handles boolean/string TRUE values', () => {
    var sheet = buildFailsafeSheet([
      ['user@test.com', 'TRUE', 'weekly', '', 'TRUE', 'FALSE', true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.getDigestConfig('user@test.com');
    expect(result.enabled).toBe(true);
    expect(result.includeGrievances).toBe(true);
    expect(result.includeWorkload).toBe(false);
    expect(result.includeTasks).toBe(true);
  });
});

// ============================================================================
// FailsafeService.updateDigestConfig
// ============================================================================

describe('FailsafeService.updateDigestConfig', () => {
  test('rejects missing email', () => {
    var result = FailsafeService.updateDigestConfig(null, { enabled: true });
    expect(result.success).toBe(false);
    expect(result.message).toContain('Missing');
  });

  test('rejects missing config', () => {
    var result = FailsafeService.updateDigestConfig('user@test.com', null);
    expect(result.success).toBe(false);
    expect(result.message).toContain('Missing');
  });

  test('updates existing row', () => {
    var sheet = buildFailsafeSheet([
      ['user@test.com', false, 'weekly', '', true, true, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.updateDigestConfig('user@test.com', {
      enabled: true,
      frequency: 'monthly',
      includeGrievances: true,
      includeWorkload: false,
      includeTasks: true
    });
    expect(result.success).toBe(true);
    expect(result.message).toContain('updated');
    // Should have called setValue on row 2, columns 2, 3, 5, 6, 7
    expect(sheet.getRange).toHaveBeenCalledWith(2, 2);
    expect(sheet.getRange).toHaveBeenCalledWith(2, 3);
  });

  test('adds new row if email not found', () => {
    var sheet = buildFailsafeSheet([
      ['existing@test.com', true, 'weekly', '', true, true, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.updateDigestConfig('new@test.com', {
      enabled: true,
      frequency: 'weekly'
    });
    expect(result.success).toBe(true);
    expect(result.message).toContain('saved');
    expect(sheet.appendRow).toHaveBeenCalled();
    var appendArgs = sheet.appendRow.mock.calls[0][0];
    expect(appendArgs[0]).toBe('new@test.com');
  });

  test('normalizes frequency to weekly/monthly', () => {
    var sheet = buildFailsafeSheet([]);
    installSS([sheet]);

    // Frequency that is not 'monthly' should default to 'weekly'
    FailsafeService.updateDigestConfig('user@test.com', {
      enabled: true,
      frequency: 'daily'
    });
    var appendArgs = sheet.appendRow.mock.calls[0][0];
    expect(appendArgs[2]).toBe('weekly');
  });

  test('uses ScriptLock', () => {
    var sheet = buildFailsafeSheet([]);
    installSS([sheet]);

    FailsafeService.updateDigestConfig('user@test.com', { enabled: true });
    expect(LockService.getScriptLock).toHaveBeenCalled();
  });
});

// ============================================================================
// FailsafeService.processScheduledDigests
// ============================================================================

describe('FailsafeService.processScheduledDigests', () => {
  test('returns {processed:0} when no config sheet', () => {
    installSS([]);
    var result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
  });

  test('skips disabled digests', () => {
    var sheet = buildFailsafeSheet([
      ['user@test.com', false, 'weekly', '', true, true, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });

  test('respects frequency scheduling', () => {
    // Last sent 3 days ago, frequency is weekly -> should not send
    var threeDaysAgo = new Date();
    threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
    var sheet = buildFailsafeSheet([
      ['user@test.com', true, 'weekly', threeDaysAgo, true, true, true]
    ]);
    installSS([sheet]);

    var result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
  });

  test('stops on low email quota', () => {
    // Last sent 10 days ago -> eligible for weekly send
    var tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 10);

    var sheet = buildFailsafeSheet([
      ['user1@test.com', true, 'weekly', tenDaysAgo, true, true, true],
      ['user2@test.com', true, 'weekly', tenDaysAgo, true, true, true]
    ]);
    installSS([sheet]);

    // Low quota
    MailApp.getRemainingDailyQuota.mockReturnValue(3);

    var result = FailsafeService.processScheduledDigests();
    // Should stop before processing because quota < 5
    expect(result.processed).toBe(0);
  });

  test('updates lastSent date after sending', () => {
    // Last sent 10 days ago with DataService providing data
    var tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 10);

    var sheet = buildFailsafeSheet([
      ['user@test.com', true, 'weekly', tenDaysAgo, true, false, false]
    ]);
    installSS([sheet]);

    // Mock DataService to provide grievances so _composeMemberDigest returns content
    global.DataService = {
      getMemberGrievances: jest.fn(() => [
        { status: 'Open', issueType: 'Discipline', dateFiled: '2026-01-15' }
      ]),
      getMemberTasks: jest.fn(() => [])
    };

    var result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(1);
    // getRange(2, 4).setValue(now) for last sent date
    expect(sheet.getRange).toHaveBeenCalledWith(2, 4);
  });
});

// ============================================================================
// FailsafeService.backupCriticalSheets
// ============================================================================

describe('FailsafeService.backupCriticalSheets', () => {
  test('creates backup folder if it does not exist', () => {
    // DriveApp.getFoldersByName returns no folders
    DriveApp.getFoldersByName.mockReturnValue({
      hasNext: jest.fn(() => false)
    });

    // Create sheets with data so backup has something to process
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [
      ['Name', 'Email'],
      ['Alice', 'alice@test.com']
    ]);
    installSS([memberSheet]);

    FailsafeService.backupCriticalSheets();
    expect(DriveApp.createFolder).toHaveBeenCalledWith('DDS_Dashboard_Backups');
  });

  test('backs up sheets to CSV files', () => {
    var mockFolder = {
      createFile: jest.fn(() => ({ getId: jest.fn(() => 'csv-file-id') })),
      getFiles: jest.fn(() => ({ hasNext: jest.fn(() => false) }))
    };
    DriveApp.getFoldersByName.mockReturnValue({
      hasNext: jest.fn(() => true),
      next: jest.fn(() => mockFolder)
    });

    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [
      ['Name', 'Email'],
      ['Alice', 'alice@test.com']
    ]);
    var grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [
      ['ID', 'Status'],
      ['G1', 'Open']
    ]);
    installSS([memberSheet, grievanceSheet]);

    var result = FailsafeService.backupCriticalSheets();
    expect(result.success).toBe(true);
    expect(result.backedUp).toBe(2);
    expect(mockFolder.createFile).toHaveBeenCalledTimes(2);
  });

  test('skips empty sheets', () => {
    var mockFolder = {
      createFile: jest.fn(),
      getFiles: jest.fn(() => ({ hasNext: jest.fn(() => false) }))
    };
    DriveApp.getFoldersByName.mockReturnValue({
      hasNext: jest.fn(() => true),
      next: jest.fn(() => mockFolder)
    });

    // Sheet with only header row (getLastRow returns 1)
    var emptySheet = createMockSheet(SHEETS.MEMBER_DIR, [['Name', 'Email']]);
    emptySheet.getLastRow.mockReturnValue(1);
    installSS([emptySheet]);

    var result = FailsafeService.backupCriticalSheets();
    expect(result.backedUp).toBe(0);
    expect(mockFolder.createFile).not.toHaveBeenCalled();
  });

  test('returns backedUp count and folderName', () => {
    var mockFolder = {
      createFile: jest.fn(() => ({ getId: jest.fn(() => 'id') })),
      getFiles: jest.fn(() => ({ hasNext: jest.fn(() => false) }))
    };
    DriveApp.getFoldersByName.mockReturnValue({
      hasNext: jest.fn(() => true),
      next: jest.fn(() => mockFolder)
    });

    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [
      ['Name', 'Email'],
      ['Alice', 'alice@test.com']
    ]);
    installSS([memberSheet]);

    var result = FailsafeService.backupCriticalSheets();
    expect(result).toHaveProperty('backedUp');
    expect(result).toHaveProperty('folderName', 'DDS_Dashboard_Backups');
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// FailsafeService.setupFailsafeTriggers / removeFailsafeTriggers
// ============================================================================

describe('FailsafeService.setupFailsafeTriggers', () => {
  test('creates daily + weekly triggers', () => {
    var result = FailsafeService.setupFailsafeTriggers();
    expect(result.success).toBe(true);
    // Two newTrigger calls: one for daily digest, one for weekly backup
    expect(ScriptApp.newTrigger).toHaveBeenCalledTimes(2);
    expect(ScriptApp.newTrigger).toHaveBeenCalledWith('fsProcessScheduledDigests');
    expect(ScriptApp.newTrigger).toHaveBeenCalledWith('fsBackupCriticalSheets');
  });

  test('removes existing triggers before installing new ones', () => {
    // Mock existing triggers
    var existingTrigger1 = {
      getHandlerFunction: jest.fn(() => 'fsProcessScheduledDigests')
    };
    var existingTrigger2 = {
      getHandlerFunction: jest.fn(() => 'fsBackupCriticalSheets')
    };
    ScriptApp.getProjectTriggers.mockReturnValue([existingTrigger1, existingTrigger2]);

    FailsafeService.setupFailsafeTriggers();

    // Should have deleted the existing triggers first
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(existingTrigger1);
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(existingTrigger2);
    // Then created new ones
    expect(ScriptApp.newTrigger).toHaveBeenCalledTimes(2);
  });
});

describe('FailsafeService.removeFailsafeTriggers', () => {
  test('removes matching triggers', () => {
    var fsTrigger = {
      getHandlerFunction: jest.fn(() => 'fsProcessScheduledDigests')
    };
    var otherTrigger = {
      getHandlerFunction: jest.fn(() => 'someOtherFunction')
    };
    ScriptApp.getProjectTriggers.mockReturnValue([fsTrigger, otherTrigger]);

    var result = FailsafeService.removeFailsafeTriggers();
    expect(result.success).toBe(true);
    expect(result.removed).toBe(1);
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(fsTrigger);
    expect(ScriptApp.deleteTrigger).not.toHaveBeenCalledWith(otherTrigger);
  });

  test('returns success with 0 removed when no matching triggers', () => {
    ScriptApp.getProjectTriggers.mockReturnValue([]);
    var result = FailsafeService.removeFailsafeTriggers();
    expect(result.success).toBe(true);
    expect(result.removed).toBe(0);
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  beforeEach(() => {
    // v4.51.1: gas-mock defaults to deny — explicitly opt in for wrapper tests
    global._resolveCallerEmail = jest.fn(() => 'test@example.com');
    global._requireStewardAuth = jest.fn(() => 'steward@example.com');
  });

  test('fsGetDigestConfig delegates to FailsafeService', () => {
    // Auth resolves via _resolveCallerEmail() mock — returns config (not null)
    var result = fsGetDigestConfig(null);
    expect(result).not.toBeNull();
    expect(typeof result).toBe('object');
  });

  test('fsSetupTriggers delegates to FailsafeService', () => {
    var result = fsSetupTriggers();
    expect(result.success).toBe(true);
    expect(ScriptApp.newTrigger).toHaveBeenCalledTimes(2);
  });
});
