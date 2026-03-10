/**
 * Negative Path / Error Handling Tests (TEST-02)
 *
 * Tests error paths that are critical for production reliability:
 *   - Permission escalation (member calling steward-only functions)
 *   - Malformed / missing data handling
 *   - Unicode / emoji injection safety
 *   - Quota exhaustion behavior
 *   - Spreadsheet unavailability
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load core modules + modules under test
// 06_Maintenance.gs provides logAuditEvent and maskEmail
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '06_Maintenance.gs',
  '26_QAForum.gs',
  '28_FailsafeService.gs',
]);


// ============================================================================
// Permission Escalation — Members calling steward-only functions
// ============================================================================

describe('Permission escalation: non-steward calling steward-only functions', () => {

  beforeEach(() => {
    // Mock _requireStewardAuth to throw (simulating non-steward caller)
    global._requireStewardAuth = jest.fn(() => {
      throw new Error('Access denied: steward role required.');
    });
    // Mock _resolveCallerEmail to return a member email
    global._resolveCallerEmail = jest.fn(() => 'member@example.com');
  });

  afterEach(() => {
    global._requireStewardAuth = jest.fn(() => 'steward@example.com');
    global._resolveCallerEmail = jest.fn(() => 'test@example.com');
  });

  test('qaSubmitAnswer rejects non-steward callers', () => {
    // QAForum.submitAnswer checks isSteward param; the wrapper resolves auth
    const result = QAForum.submitAnswer('member@example.com', 'Member', 'QA_123', 'My answer', false);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/steward/i);
  });

  test('fsSetupTriggers wrapper throws for non-steward', () => {
    expect(() => {
      fsSetupTriggers('invalid-token');
    }).toThrow(/steward|access denied/i);
  });

  test('fsRemoveTriggers wrapper throws for non-steward', () => {
    expect(() => {
      fsRemoveTriggers('invalid-token');
    }).toThrow(/steward|access denied/i);
  });

  test('fsBackupCriticalSheets wrapper throws for non-steward', () => {
    expect(() => {
      fsBackupCriticalSheets('invalid-token');
    }).toThrow(/steward|access denied/i);
  });

  test('fsTriggerBulkExport wrapper throws for non-steward', () => {
    expect(() => {
      fsTriggerBulkExport('invalid-token');
    }).toThrow(/steward|access denied/i);
  });
});


// ============================================================================
// Malformed / Missing Data Handling
// ============================================================================

describe('Malformed data handling', () => {

  test('QAForum.submitQuestion rejects empty text', () => {
    const result = QAForum.submitQuestion('user@test.com', 'User', '', false);
    expect(result.success).toBe(false);
  });

  test('QAForum.submitQuestion rejects null email', () => {
    const result = QAForum.submitQuestion(null, 'User', 'My question?', false);
    expect(result.success).toBe(false);
  });

  test('QAForum.submitQuestion rejects whitespace-only text', () => {
    const result = QAForum.submitQuestion('user@test.com', 'User', '   \n  ', false);
    expect(result.success).toBe(false);
  });

  test('QAForum.upvoteQuestion rejects missing questionId', () => {
    const result = QAForum.upvoteQuestion('user@test.com', null);
    expect(result.success).toBe(false);
  });

  test('QAForum.submitAnswer rejects empty text', () => {
    const result = QAForum.submitAnswer('s@test.com', 'Steward', 'QA_1', '', true);
    expect(result.success).toBe(false);
  });

  test('FailsafeService.getDigestConfig returns null for empty email', () => {
    const result = FailsafeService.getDigestConfig('');
    expect(result).toBeNull();
  });

  test('FailsafeService.getDigestConfig returns null for null email', () => {
    const result = FailsafeService.getDigestConfig(null);
    expect(result).toBeNull();
  });

  test('FailsafeService.updateDigestConfig rejects missing data', () => {
    const result = FailsafeService.updateDigestConfig(null, {});
    expect(result.success).toBe(false);
  });

  test('FailsafeService.updateDigestConfig rejects null config', () => {
    const result = FailsafeService.updateDigestConfig('user@test.com', null);
    expect(result.success).toBe(false);
  });
});


// ============================================================================
// Spreadsheet Unavailability
// ============================================================================

describe('Spreadsheet unavailability', () => {

  beforeEach(() => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
  });

  afterEach(() => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(
      createMockSpreadsheet([])
    );
  });

  test('QAForum.getQuestions returns empty when spreadsheet is null', () => {
    const result = QAForum.getQuestions('user@test.com', 1, 20, 'recent');
    expect(result.questions).toEqual([]);
    expect(result.total).toBe(0);
  });

  test('QAForum.getQuestionDetail returns null when spreadsheet is null', () => {
    const result = QAForum.getQuestionDetail('user@test.com', 'QA_123');
    expect(result).toBeNull();
  });

  test('FailsafeService.getDigestConfig returns defaults when spreadsheet is null', () => {
    const result = FailsafeService.getDigestConfig('user@test.com');
    expect(result).toBeDefined();
    expect(result.enabled).toBe(false);
    expect(result.frequency).toBe('weekly');
  });

  test('FailsafeService.updateDigestConfig returns error when spreadsheet is null', () => {
    const result = FailsafeService.updateDigestConfig('u@test.com', { enabled: true });
    expect(result.success).toBe(false);
  });

  test('FailsafeService.backupCriticalSheets returns error when spreadsheet is null', () => {
    const result = FailsafeService.backupCriticalSheets();
    expect(result.success).toBe(false);
  });

  test('FailsafeService.processScheduledDigests returns 0 when spreadsheet is null', () => {
    const result = FailsafeService.processScheduledDigests();
    expect(result.processed).toBe(0);
  });
});


// ============================================================================
// Unicode / Emoji Injection
// ============================================================================

describe('Unicode / emoji safety in text inputs', () => {

  beforeEach(() => {
    global._resolveCallerEmail = jest.fn(() => 'test@example.com');
    // Set up sheets for question submission
    const forumSheet = createMockSheet(SHEETS.QA_FORUM, [
      ['ID', 'Author Email', 'Author Name', 'Is Anonymous', 'Question Text',
        'Status', 'Upvote Count', 'Upvoters', 'Answer Count', 'Created', 'Updated']
    ]);
    const ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  });

  afterEach(() => {
    global._resolveCallerEmail = jest.fn(() => 'test@example.com');
  });

  test('submitQuestion handles emoji text without throwing', () => {
    const result = QAForum.submitQuestion(
      'user@test.com', 'User',
      'How do I file a grievance? \uD83D\uDE00\uD83D\uDC4D\u2764\uFE0F',
      false
    );
    expect(result.success).toBe(true);
    expect(result.questionId).toBeDefined();
  });

  test('submitQuestion handles zero-width characters', () => {
    const result = QAForum.submitQuestion(
      'user@test.com', 'User',
      'Normal\u200B\u200C\u200D\uFEFF text with zero-width chars',
      false
    );
    expect(result.success).toBe(true);
  });

  test('submitQuestion handles RTL override characters', () => {
    const result = QAForum.submitQuestion(
      'user@test.com', 'User',
      'Text with \u202E RTL override \u202C and normal',
      false
    );
    expect(result.success).toBe(true);
  });

  test('submitQuestion truncates extremely long input to 2000 chars', () => {
    const longText = 'A'.repeat(5000);
    const result = QAForum.submitQuestion('user@test.com', 'User', longText, false);
    expect(result.success).toBe(true);
    // Verify appendRow was called with truncated text
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    const appendCall = sheet.appendRow.mock.calls[0];
    expect(appendCall[0][4].length).toBeLessThanOrEqual(2000);
  });
});


// ============================================================================
// Rate Limiting
// ============================================================================

describe('Rate limiting on Q&A submissions', () => {

  beforeEach(() => {
    global._resolveCallerEmail = jest.fn(() => 'test@example.com');
    const forumSheet = createMockSheet(SHEETS.QA_FORUM, [
      ['ID', 'Author Email', 'Author Name', 'Is Anonymous', 'Question Text',
        'Status', 'Upvote Count', 'Upvoters', 'Answer Count', 'Created', 'Updated']
    ]);
    const ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  });

  test('blocks after 5 questions in rapid succession', () => {
    // Submit 5 questions (should all succeed)
    for (let i = 0; i < 5; i++) {
      const result = QAForum.submitQuestion('rapid@test.com', 'User', 'Question ' + i, false);
      expect(result.success).toBe(true);
    }

    // 6th should be rate-limited
    const result = QAForum.submitQuestion('rapid@test.com', 'User', 'Question 6', false);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/rate limit/i);
  });

  test('rate limit is per-user (different user succeeds)', () => {
    // Fill up rate limit for one user
    for (let i = 0; i < 5; i++) {
      QAForum.submitQuestion('full@test.com', 'User', 'Question ' + i, false);
    }

    // Different user should still succeed
    const result = QAForum.submitQuestion('other@test.com', 'Other', 'My question', false);
    expect(result.success).toBe(true);
  });
});


// ============================================================================
// Email Quota Exhaustion
// ============================================================================

describe('MailApp quota exhaustion', () => {

  test('processScheduledDigests stops when quota is low', () => {
    // Set up enabled digest user
    const configData = [
      ['Email', 'Digest Enabled', 'Digest Frequency', 'Last Digest Sent',
        'Include Grievances', 'Include Workload', 'Include Tasks'],
      ['user1@test.com', true, 'weekly', '', true, true, true],
      ['user2@test.com', true, 'weekly', '', true, true, true],
    ];
    const configSheet = createMockSheet(SHEETS.FAILSAFE_CONFIG, configData);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Email'], ['user1@test.com']]);
    const ss = createMockSpreadsheet([configSheet, memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Simulate low quota — returns 2 (below the threshold of 5)
    MailApp.getRemainingDailyQuota.mockReturnValue(2);

    const result = FailsafeService.processScheduledDigests();
    // Should process 0 because quota check happens before sending
    expect(result.processed).toBe(0);
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });
});


// ============================================================================
// Lock Contention
// ============================================================================

describe('Lock contention handling', () => {

  test('updateDigestConfig handles lock failure gracefully', () => {
    // Set up sheets
    const configSheet = createMockSheet(SHEETS.FAILSAFE_CONFIG, [
      ['Email', 'Digest Enabled', 'Digest Frequency', 'Last Digest Sent',
        'Include Grievances', 'Include Workload', 'Include Tasks'],
    ]);
    const ss = createMockSpreadsheet([configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Mock lock to throw (another process holds it)
    LockService.getScriptLock.mockReturnValue({
      tryLock: jest.fn(() => false),
      waitLock: jest.fn(() => { throw new Error('Lock timeout'); }),
      releaseLock: jest.fn()
    });

    expect(() => {
      FailsafeService.updateDigestConfig('user@test.com', { enabled: true });
    }).toThrow(/lock/i);

    // Restore lock mock
    LockService.getScriptLock.mockReturnValue({
      tryLock: jest.fn(() => true),
      waitLock: jest.fn(),
      releaseLock: jest.fn()
    });
  });
});
