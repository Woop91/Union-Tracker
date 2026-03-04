/**
 * Tests for 18_WorkloadTracker.gs
 *
 * Covers WT_APP_CONFIG, WT_SECURITY_CONFIG, WT_VAULT_COLS, WT_SUB_CATEGORIES,
 * sanitizeString, rate limiting, withLock, and authenticateWorkloadMember_.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock verifyPIN before loading workload tracker
global.verifyPIN = jest.fn(() => true);
global.hashPIN = jest.fn(() => 'mockhash');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '18_WorkloadTracker.gs']);

// ============================================================================
// Configuration Constants
// ============================================================================

describe('WT_APP_CONFIG', () => {
  test('has version string', () => {
    expect(typeof WT_APP_CONFIG.version).toBe('string');
  });

  test('has name = Workload Tracker', () => {
    expect(WT_APP_CONFIG.name).toBe('Workload Tracker');
  });

  test('maxStringLength is 255', () => {
    expect(WT_APP_CONFIG.maxStringLength).toBe(255);
  });

  test('dataRetentionMonths is 24', () => {
    expect(WT_APP_CONFIG.dataRetentionMonths).toBe(24);
  });

  test('lockTimeoutMs is 30000', () => {
    expect(WT_APP_CONFIG.lockTimeoutMs).toBe(30000);
  });
});

describe('WT_SECURITY_CONFIG', () => {
  test('maxPinAttempts is 5', () => {
    expect(WT_SECURITY_CONFIG.maxPinAttempts).toBe(5);
  });

  test('pinLockoutMinutes is 15', () => {
    expect(WT_SECURITY_CONFIG.pinLockoutMinutes).toBe(15);
  });

  test('maxSubmissionsPerHour is 10', () => {
    expect(WT_SECURITY_CONFIG.maxSubmissionsPerHour).toBe(10);
  });

  test('maxHistoryRequestsPerHour is 20', () => {
    expect(WT_SECURITY_CONFIG.maxHistoryRequestsPerHour).toBe(20);
  });

  test('maxDashboardRequestsPerHour is 30', () => {
    expect(WT_SECURITY_CONFIG.maxDashboardRequestsPerHour).toBe(30);
  });
});

// ============================================================================
// Column Constants
// ============================================================================

describe('WT_VAULT_COLS', () => {
  test('TIMESTAMP is 0', () => {
    expect(WT_VAULT_COLS.TIMESTAMP).toBe(0);
  });

  test('EMAIL is 1', () => {
    expect(WT_VAULT_COLS.EMAIL).toBe(1);
  });

  test('PRIORITY_CASES is 2', () => {
    expect(WT_VAULT_COLS.PRIORITY_CASES).toBe(2);
  });

  test('OVERTIME_HOURS is 23', () => {
    expect(WT_VAULT_COLS.OVERTIME_HOURS).toBe(23);
  });

  test('has 24 columns', () => {
    expect(Object.keys(WT_VAULT_COLS).length).toBe(24);
  });
});

describe('WT_USERMETA_COLS', () => {
  test('EMAIL is 0', () => {
    expect(WT_USERMETA_COLS.EMAIL).toBe(0);
  });

  test('SHARING_START_DATE is 1', () => {
    expect(WT_USERMETA_COLS.SHARING_START_DATE).toBe(1);
  });
});

describe('WT_REMINDERS_COLS', () => {
  test('EMAIL is 0', () => {
    expect(WT_REMINDERS_COLS.EMAIL).toBe(0);
  });

  test('ENABLED is 1', () => {
    expect(WT_REMINDERS_COLS.ENABLED).toBe(1);
  });

  test('LAST_SENT is 5', () => {
    expect(WT_REMINDERS_COLS.LAST_SENT).toBe(5);
  });
});

// ============================================================================
// Sub-Categories
// ============================================================================

describe('WT_SUB_CATEGORIES', () => {
  test('has 8 category keys', () => {
    expect(Object.keys(WT_SUB_CATEGORIES).length).toBe(8);
  });

  test('priority has expected sub-categories', () => {
    expect(WT_SUB_CATEGORIES.priority).toContain('QDD');
    expect(WT_SUB_CATEGORIES.priority).toContain('Congressional');
  });

  test('pending has expected sub-categories', () => {
    expect(WT_SUB_CATEGORIES.pending).toContain('New Cases');
    expect(WT_SUB_CATEGORIES.pending).toContain('Approval Returns');
  });

  test('aged has day ranges', () => {
    expect(WT_SUB_CATEGORIES.aged).toContain('60-89 Days');
    expect(WT_SUB_CATEGORIES.aged).toContain('180+ Days');
  });

  test('all sub-categories are arrays', () => {
    Object.values(WT_SUB_CATEGORIES).forEach(val => {
      expect(Array.isArray(val)).toBe(true);
      expect(val.length).toBeGreaterThan(0);
    });
  });
});

// ============================================================================
// sanitizeString
// ============================================================================

describe('sanitizeString', () => {
  test('returns empty string for null', () => {
    expect(sanitizeString(null)).toBe('');
  });

  test('returns empty string for number', () => {
    expect(sanitizeString(123)).toBe('');
  });

  test('trims whitespace', () => {
    expect(sanitizeString('  hello  ')).toBe('hello');
  });

  test('strips control characters', () => {
    expect(sanitizeString('hello\x00world')).toBe('helloworld');
    expect(sanitizeString('test\x1Fchar')).toBe('testchar');
  });

  test('truncates to maxStringLength (default 255)', () => {
    const long = 'A'.repeat(300);
    expect(sanitizeString(long).length).toBe(255);
  });

  test('truncates to custom maxLength', () => {
    const long = 'B'.repeat(50);
    expect(sanitizeString(long, 20).length).toBe(20);
  });

  test('preserves valid strings', () => {
    expect(sanitizeString('Hello World!')).toBe('Hello World!');
  });

  test('preserves newlines and tabs', () => {
    // \n and \t are not stripped (only \x00-\x08, \x0B, \x0C, \x0E-\x1F, \x7F)
    expect(sanitizeString('line1\nline2')).toBe('line1\nline2');
  });
});

// ============================================================================
// Rate Limiting
// ============================================================================

describe('wtCheckRateLimit_', () => {
  test('allows first attempt', () => {
    const result = wtCheckRateLimit_('test_key', 5, 15);
    expect(result.allowed).toBe(true);
    expect(result.attemptsRemaining).toBe(5);
  });

  test('blocks when max attempts reached', () => {
    // Simulate max attempts by putting value in cache
    CacheService.getScriptCache().put('WT_RATE_test_block', '5', 900);
    const result = wtCheckRateLimit_('test_block', 5, 15);
    expect(result.allowed).toBe(false);
    expect(result.attemptsRemaining).toBe(0);
    expect(result.waitMinutes).toBe(15);
  });

  test('allows when under max attempts', () => {
    CacheService.getScriptCache().put('WT_RATE_test_under', '3', 900);
    const result = wtCheckRateLimit_('test_under', 5, 15);
    expect(result.allowed).toBe(true);
    expect(result.attemptsRemaining).toBe(2);
  });
});

describe('wtRecordRateLimitAttempt_', () => {
  test('increments counter', () => {
    wtRecordRateLimitAttempt_('test_record', 15);
    const result = CacheService.getScriptCache().get('WT_RATE_test_record');
    expect(result).toBe('1');
  });

  test('increments existing counter', () => {
    CacheService.getScriptCache().put('WT_RATE_test_incr', '2', 900);
    wtRecordRateLimitAttempt_('test_incr', 15);
    const result = CacheService.getScriptCache().get('WT_RATE_test_incr');
    expect(result).toBe('3');
  });
});

describe('wtClearRateLimit_', () => {
  test('removes rate limit entry', () => {
    CacheService.getScriptCache().put('WT_RATE_test_clear', '5', 900);
    wtClearRateLimit_('test_clear');
    expect(CacheService.getScriptCache().get('WT_RATE_test_clear')).toBeNull();
  });
});

describe('wtCheckPinRateLimit_', () => {
  test('lowercases email', () => {
    const result = wtCheckPinRateLimit_('User@Test.COM');
    expect(result.allowed).toBe(true);
  });
});

describe('wtCheckSubmissionRateLimit_', () => {
  test('uses 60 minute window with 10 max', () => {
    const result = wtCheckSubmissionRateLimit_('user@test.com');
    expect(result.allowed).toBe(true);
    expect(result.attemptsRemaining).toBe(10);
  });
});

// ============================================================================
// withLock
// ============================================================================

describe('withLock', () => {
  test('executes function under lock', () => {
    const fn = jest.fn(() => 'result');
    const result = withLock(fn);
    expect(fn).toHaveBeenCalled();
    expect(result).toBe('result');
  });

  test('releases lock after function', () => {
    const mockLock = {
      tryLock: jest.fn(() => true),
      releaseLock: jest.fn()
    };
    LockService.getScriptLock = jest.fn(() => mockLock);

    withLock(() => {});
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });

  test('releases lock even on error', () => {
    const mockLock = {
      tryLock: jest.fn(() => true),
      releaseLock: jest.fn()
    };
    LockService.getScriptLock = jest.fn(() => mockLock);

    expect(() => withLock(() => { throw new Error('boom'); })).toThrow('boom');
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });

  test('throws when lock acquisition fails', () => {
    const mockLock = {
      tryLock: jest.fn(() => false),
      releaseLock: jest.fn()
    };
    LockService.getScriptLock = jest.fn(() => mockLock);

    expect(() => withLock(() => {})).toThrow('System busy');
  });
});

// ============================================================================
// authenticateWorkloadMember_
// ============================================================================

describe('authenticateWorkloadMember_', () => {
  let mockMemberSheet;

  beforeEach(() => {
    // Reset verifyPIN mock
    global.verifyPIN = jest.fn(() => true);

    // Build member data with correct column indices
    const emailCol = (MEMBER_COLS.EMAIL || 9) - 1;
    const idCol = (MEMBER_COLS.MEMBER_ID || 1) - 1;
    const pinCol = (MEMBER_COLS.PIN_HASH || 33) - 1;

    const headerRow = new Array(Math.max(emailCol, idCol, pinCol) + 1).fill('');
    headerRow[emailCol] = 'Email';
    headerRow[idCol] = 'Member ID';
    headerRow[pinCol] = 'PIN Hash';

    const dataRow = new Array(Math.max(emailCol, idCol, pinCol) + 1).fill('');
    dataRow[emailCol] = 'user@test.com';
    dataRow[idCol] = 'MEM-001';
    dataRow[pinCol] = 'stored_hash_value';

    mockMemberSheet = createMockSheet(SHEETS.MEMBER_DIR, [headerRow, dataRow]);

    const mockSs = createMockSpreadsheet([mockMemberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
  });

  test('rejects empty email', () => {
    const result = authenticateWorkloadMember_('', '123456');
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('rejects empty PIN', () => {
    const result = authenticateWorkloadMember_('user@test.com', '');
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('returns success for valid credentials', () => {
    global.verifyPIN = jest.fn(() => true);
    const result = authenticateWorkloadMember_('user@test.com', '123456');
    expect(result.success).toBe(true);
    expect(result.memberId).toBe('MEM-001');
  });

  test('returns failure for invalid PIN', () => {
    global.verifyPIN = jest.fn(() => false);
    const result = authenticateWorkloadMember_('user@test.com', '000000');
    expect(result.success).toBe(false);
    expect(result.message).toContain('Invalid PIN');
  });

  test('lowercases email for comparison', () => {
    global.verifyPIN = jest.fn(() => true);
    const result = authenticateWorkloadMember_('USER@TEST.COM', '123456');
    expect(result.success).toBe(true);
  });

  test('blocks when rate limited', () => {
    CacheService.getScriptCache().put('WT_RATE_PIN_user@test.com', '5', 900);
    const result = authenticateWorkloadMember_('user@test.com', '123456');
    expect(result.success).toBe(false);
    expect(result.message).toContain('Too many');
  });

  test('clears rate limit on success', () => {
    CacheService.getScriptCache().put('WT_RATE_PIN_user@test.com', '2', 900);
    global.verifyPIN = jest.fn(() => true);
    authenticateWorkloadMember_('user@test.com', '123456');
    // Rate limit should be cleared
    expect(CacheService.getScriptCache().get('WT_RATE_PIN_user@test.com')).toBeNull();
  });

  test('records failed attempt on bad PIN', () => {
    global.verifyPIN = jest.fn(() => false);
    authenticateWorkloadMember_('user@test.com', '000000');
    const attempts = CacheService.getScriptCache().get('WT_RATE_PIN_user@test.com');
    expect(attempts).toBe('1');
  });
});
