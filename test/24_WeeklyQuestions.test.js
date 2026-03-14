/**
 * Tests for 24_WeeklyQuestions.gs
 *
 * Covers WeeklyQuestions IIFE: sheet setup, active questions, response
 * submission, pool questions, steward question setting, history, poll
 * closing, frequency settings, rate limiting, and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '24_WeeklyQuestions.gs']);

let mockQuestionsSheet;
let mockResponsesSheet;
let mockPoolSheet;
let mockSs;

function setupSheets(questionsData, responsesData, poolData) {
  mockQuestionsSheet = createMockSheet(SHEETS.WEEKLY_QUESTIONS || '_Weekly_Questions', questionsData || [['header']]);
  mockResponsesSheet = createMockSheet(SHEETS.WEEKLY_RESPONSES || '_Weekly_Responses', responsesData || [['header']]);
  mockPoolSheet = createMockSheet(SHEETS.QUESTION_POOL || '_Question_Pool', poolData || [['header']]);

  // Custom getRange for questions sheet
  mockQuestionsSheet.getRange = jest.fn((row, col, numRows, numCols) => {
    const range = {
      getValue: jest.fn(() => {
        if (questionsData && questionsData[row - 1]) return questionsData[row - 1][col - 1] || '';
        return '';
      }),
      getValues: jest.fn(() => {
        if (!questionsData) return [['']];
        if (numRows && numCols) {
          return questionsData.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols));
        }
        return [[questionsData[row - 1] ? questionsData[row - 1][col - 1] : '']];
      }),
      setValue: jest.fn(),
      setValues: jest.fn(),
      setBackground: jest.fn(function() { return this; }),
      setFontWeight: jest.fn(function() { return this; }),
      setFontColor: jest.fn(function() { return this; }),
      setFontFamily: jest.fn(function() { return this; }),
      setFontSize: jest.fn(function() { return this; }),
      setVerticalAlignment: jest.fn(function() { return this; }),
      setHorizontalAlignment: jest.fn(function() { return this; }),
      insertCheckboxes: jest.fn()
    };
    return range;
  });

  mockSs = createMockSpreadsheet([mockQuestionsSheet, mockResponsesSheet, mockPoolSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

beforeEach(() => {
  setupSheets(null, null, null);
});

// ============================================================================
// Module existence
// ============================================================================

describe('WeeklyQuestions module', () => {
  test('WeeklyQuestions is defined', () => {
    expect(typeof WeeklyQuestions).toBe('object');
  });

  test('exports expected methods', () => {
    expect(typeof WeeklyQuestions.initWeeklyQuestionSheets).toBe('function');
    expect(typeof WeeklyQuestions.getActiveQuestions).toBe('function');
    expect(typeof WeeklyQuestions.submitResponse).toBe('function');
    expect(typeof WeeklyQuestions.submitPoolQuestion).toBe('function');
    expect(typeof WeeklyQuestions.setStewardQuestion).toBe('function');
    expect(typeof WeeklyQuestions.selectRandomPoolQuestion).toBe('function');
    expect(typeof WeeklyQuestions.closePoll).toBe('function');
    expect(typeof WeeklyQuestions.getHistory).toBe('function');
    expect(typeof WeeklyQuestions.getPoolCount).toBe('function');
    expect(typeof WeeklyQuestions.getPollFrequency).toBe('function');
    expect(typeof WeeklyQuestions.setPollFrequency).toBe('function');
  });

  test('exports Q_COLS indices', () => {
    expect(WeeklyQuestions.Q_COLS).toBeDefined();
    expect(typeof WeeklyQuestions.Q_COLS.ID).toBe('number');
    expect(typeof WeeklyQuestions.Q_COLS.TEXT).toBe('number');
    expect(typeof WeeklyQuestions.Q_COLS.ACTIVE).toBe('number');
  });
});

// ============================================================================
// initWeeklyQuestionSheets
// ============================================================================

describe('WeeklyQuestions.initWeeklyQuestionSheets', () => {
  test('creates sheets if missing', () => {
    const emptySs = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => emptySs);
    WeeklyQuestions.initWeeklyQuestionSheets();
    expect(emptySs.insertSheet).toHaveBeenCalled();
  });

  test('does not recreate existing sheets', () => {
    const questionsData = [['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses']];
    const responsesData = [['ID', 'Question ID', 'Email', 'Response', 'Submitted']];
    setupSheets(questionsData, responsesData);
    WeeklyQuestions.initWeeklyQuestionSheets();
    expect(mockSs.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// getActiveQuestions
// ============================================================================

describe('WeeklyQuestions.getActiveQuestions', () => {
  test('returns object with questions array', () => {
    const result = WeeklyQuestions.getActiveQuestions('user@test.com');
    expect(result).toHaveProperty('questions');
    expect(Array.isArray(result.questions)).toBe(true);
  });

  test('returns empty questions when sheet has no data', () => {
    setupSheets([['header']], null);
    const result = WeeklyQuestions.getActiveQuestions('user@test.com');
    expect(result.questions).toEqual([]);
  });

  test('each question has id, text, options, source, stats', () => {
    // Build a question row matching the current period
    const now = new Date();
    const day = now.getDay();
    const monday = new Date(now);
    monday.setDate(now.getDate() - (day === 0 ? 6 : day - 1));
    monday.setHours(0, 0, 0, 0);
    const periodKey = monday.toISOString().split('T')[0];

    const qData = [
      ['ID', 'Text', 'Options', 'Source', 'SubmittedBy', 'WeekStart', 'Active', 'Created'],
      ['PL_test1', 'Favorite color?', '["Red","Blue","Green"]', 'steward', 'admin@test.com', periodKey, 'TRUE', new Date()],
    ];
    setupSheets(qData, [['header']], null);
    const result = WeeklyQuestions.getActiveQuestions('user@test.com');
    expect(result.questions.length).toBe(1);
    const q = result.questions[0];
    expect(q.id).toBe('PL_test1');
    expect(q.text).toBe('Favorite color?');
    expect(q.options).toEqual(['Red', 'Blue', 'Green']);
    expect(q.source).toBe('steward');
    expect(q.stats).toBeDefined();
    expect(q.stats.total).toBe(0);
    expect(q.hasResponded).toBe(false);
  });
});

// ============================================================================
// submitResponse
// ============================================================================

describe('WeeklyQuestions.submitResponse', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestions.submitResponse(null, 'QID', 'answer');
    expect(result.success).toBe(false);
  });

  test('rejects missing questionId', () => {
    const result = WeeklyQuestions.submitResponse('user@test.com', null, 'answer');
    expect(result.success).toBe(false);
  });

  test('rejects empty response', () => {
    const result = WeeklyQuestions.submitResponse('user@test.com', 'QID', '');
    expect(result.success).toBe(false);
  });

  test('successful response calls appendRow and returns stats', () => {
    setupSheets(null, [['header']], null);
    const result = WeeklyQuestions.submitResponse('user@test.com', 'QID1', 'Red');
    expect(result.success).toBe(true);
    expect(result.stats).toBeDefined();
    expect(typeof result.stats.total).toBe('number');
    expect(mockResponsesSheet.appendRow).toHaveBeenCalled();
  });

  test('detects duplicate vote via appendRow dedup', () => {
    // Pre-populate responses with a matching hash
    const email = 'user@test.com';
    const raw = email.trim().toLowerCase();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    const hash = digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

    const rData = [
      ['ID', 'QuestionID', 'EmailHash', 'Response', 'Timestamp'],
      ['R1', 'QID1', hash, 'Red', new Date()],
    ];
    setupSheets(null, rData, null);
    const result = WeeklyQuestions.submitResponse('user@test.com', 'QID1', 'Blue');
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/[Aa]lready/);
  });
});

// ============================================================================
// submitPoolQuestion
// ============================================================================

describe('WeeklyQuestions.submitPoolQuestion', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestions.submitPoolQuestion(null, 'Question text');
    expect(result.success).toBe(false);
  });

  test('rejects missing question text', () => {
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', '');
    expect(result.success).toBe(false);
  });

  test('rejects fewer than 2 options', () => {
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'What color?', ['Red']);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/2.*5/);
  });

  test('rejects more than 5 options', () => {
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'Pick one', ['A','B','C','D','E','F']);
    expect(result.success).toBe(false);
  });

  test('rejects duplicate options', () => {
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'Pick one', ['Red', 'red']);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/unique/i);
  });

  test('rejects options longer than 100 chars', () => {
    const longOpt = 'A'.repeat(101);
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'Pick one', [longOpt, 'Short']);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/100/);
  });

  test('rejects blank options', () => {
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'Pick one', ['Red', '']);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/blank/i);
  });

  test('successful submission with valid options', () => {
    setupSheets(null, null, [['header']]);
    const result = WeeklyQuestions.submitPoolQuestion('user@test.com', 'Best day?', ['Monday', 'Friday']);
    expect(result.success).toBe(true);
    expect(result.message).toMatch(/pool/i);
    expect(mockPoolSheet.appendRow).toHaveBeenCalled();
  });
});

// ============================================================================
// setStewardQuestion
// ============================================================================

describe('WeeklyQuestions.setStewardQuestion', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestions.setStewardQuestion(null, 'Question text');
    expect(result.success).toBe(false);
  });

  test('rejects missing text', () => {
    const result = WeeklyQuestions.setStewardQuestion('steward@test.com', '');
    expect(result.success).toBe(false);
  });

  test('rejects fewer than 2 options', () => {
    const result = WeeklyQuestions.setStewardQuestion('steward@test.com', 'Best color?', ['Red']);
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// closePoll
// ============================================================================

describe('WeeklyQuestions.closePoll', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestions.closePoll(null, 'PL_123');
    expect(result.success).toBe(false);
  });

  test('rejects missing pollId', () => {
    const result = WeeklyQuestions.closePoll('steward@test.com', null);
    expect(result.success).toBe(false);
  });

  test('returns not found for non-existent poll', () => {
    setupSheets([['header'], ['PL_other', 'Q?', '[]', 'steward', '', '', 'TRUE', new Date()]], null, null);
    const result = WeeklyQuestions.closePoll('steward@test.com', 'PL_nonexistent');
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/not found/i);
  });

  test('closes existing poll by setting Active to FALSE', () => {
    const qData = [
      ['ID', 'Text', 'Options', 'Source', 'SubmittedBy', 'WeekStart', 'Active', 'Created'],
      ['PL_abc', 'Color?', '["R","B"]', 'steward', 'admin@test.com', '2026-03-10', 'TRUE', new Date()],
    ];
    setupSheets(qData, null, null);
    const result = WeeklyQuestions.closePoll('steward@test.com', 'PL_abc');
    expect(result.success).toBe(true);
    expect(result.message).toMatch(/closed/i);
    // Verify setValue was called on the Active column
    expect(mockQuestionsSheet.getRange).toHaveBeenCalledWith(2, WeeklyQuestions.Q_COLS.ACTIVE + 1);
  });
});

// ============================================================================
// getPoolCount
// ============================================================================

describe('WeeklyQuestions.getPoolCount', () => {
  test('returns a number', () => {
    const result = WeeklyQuestions.getPoolCount();
    expect(typeof result).toBe('number');
  });

  test('returns 0 for empty pool', () => {
    setupSheets(null, null, [['header']]);
    expect(WeeklyQuestions.getPoolCount()).toBe(0);
  });

  test('counts only pending questions', () => {
    const poolData = [
      ['ID', 'Text', 'Options', 'Hash', 'Status', 'Created'],
      ['P1', 'Q1?', '["A","B"]', 'hash1', 'pending', new Date()],
      ['P2', 'Q2?', '["A","B"]', 'hash2', 'used', new Date()],
      ['P3', 'Q3?', '["A","B"]', 'hash3', 'pending', new Date()],
    ];
    setupSheets(null, null, poolData);
    expect(WeeklyQuestions.getPoolCount()).toBe(2);
  });
});

// ============================================================================
// getHistory
// ============================================================================

describe('WeeklyQuestions.getHistory', () => {
  test('returns questions array and hasMore flag', () => {
    const questionsData = [
      ['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses'],
    ];
    setupSheets(questionsData, null);
    const result = WeeklyQuestions.getHistory('user@test.com', 1, 10);
    expect(result).toHaveProperty('questions');
    expect(result).toHaveProperty('hasMore');
    expect(Array.isArray(result.questions)).toBe(true);
  });

  test('paginates correctly', () => {
    const rows = [['ID', 'Text', 'Options', 'Source', 'SubmittedBy', 'WeekStart', 'Active', 'Created']];
    for (let i = 0; i < 15; i++) {
      const d = new Date(2026, 0, 6 + (i * 7)); // successive Mondays
      rows.push(['PL_' + i, 'Q' + i + '?', '["A","B"]', 'steward', '', d, 'FALSE', d]);
    }
    setupSheets(rows, [['header']], null);

    const page1 = WeeklyQuestions.getHistory('user@test.com', 1, 5);
    expect(page1.questions.length).toBe(5);
    expect(page1.hasMore).toBe(true);

    const page3 = WeeklyQuestions.getHistory('user@test.com', 3, 5);
    expect(page3.questions.length).toBe(5);
    expect(page3.hasMore).toBe(false);
  });

  test('caps pageSize at 20', () => {
    const rows = [['ID', 'Text', 'Options', 'Source', 'SubmittedBy', 'WeekStart', 'Active', 'Created']];
    for (let i = 0; i < 25; i++) {
      rows.push(['PL_' + i, 'Q' + i, '["A","B"]', 'steward', '', new Date(2026, 0, 6), 'FALSE', new Date()]);
    }
    setupSheets(rows, [['header']], null);
    const result = WeeklyQuestions.getHistory('user@test.com', 1, 100);
    expect(result.questions.length).toBe(20);
  });
});

// ============================================================================
// Frequency settings
// ============================================================================

describe('WeeklyQuestions.getPollFrequency / setPollFrequency', () => {
  test('getPollFrequency returns a string', () => {
    const result = WeeklyQuestions.getPollFrequency();
    expect(['weekly', 'biweekly', 'monthly']).toContain(result);
  });

  test('setPollFrequency rejects invalid frequency', () => {
    const result = WeeklyQuestions.setPollFrequency('steward@test.com', 'daily');
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/[Ii]nvalid/);
  });

  test('setPollFrequency accepts weekly', () => {
    const result = WeeklyQuestions.setPollFrequency('steward@test.com', 'weekly');
    expect(result.success).toBe(true);
    expect(result.frequency).toBe('weekly');
  });

  test('setPollFrequency accepts biweekly', () => {
    const result = WeeklyQuestions.setPollFrequency('steward@test.com', 'biweekly');
    expect(result.success).toBe(true);
    expect(result.frequency).toBe('biweekly');
  });

  test('setPollFrequency accepts monthly', () => {
    const result = WeeklyQuestions.setPollFrequency('steward@test.com', 'monthly');
    expect(result.success).toBe(true);
    expect(result.frequency).toBe('monthly');
  });
});

// ============================================================================
// selectRandomPoolQuestion
// ============================================================================

describe('WeeklyQuestions.selectRandomPoolQuestion', () => {
  test('returns failure when pool is empty', () => {
    setupSheets(null, null, [['header']]);
    const result = WeeklyQuestions.selectRandomPoolQuestion();
    expect(result.success).toBe(false);
  });

  test('returns failure when no pending questions', () => {
    const poolData = [
      ['ID', 'Text', 'Options', 'Hash', 'Status', 'Created'],
      ['P1', 'Q1?', '["A","B"]', 'hash1', 'used', new Date()],
    ];
    setupSheets([['header']], null, poolData);
    const result = WeeklyQuestions.selectRandomPoolQuestion();
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global Wrappers', () => {
  test('wqGetActiveQuestions is defined', () => {
    expect(typeof wqGetActiveQuestions).toBe('function');
  });

  test('wqGetActiveQuestions returns questions array', () => {
    const result = wqGetActiveQuestions();
    expect(result).toHaveProperty('questions');
  });

  test('wqSubmitResponse is defined', () => {
    expect(typeof wqSubmitResponse).toBe('function');
  });

  test('wqSubmitResponse returns result object', () => {
    const result = wqSubmitResponse('user@test.com', 'QID', 'answer');
    expect(result).toHaveProperty('success');
  });

  test('wqSubmitPoolQuestion is defined', () => {
    expect(typeof wqSubmitPoolQuestion).toBe('function');
  });

  test('wqSetStewardQuestion is defined', () => {
    expect(typeof wqSetStewardQuestion).toBe('function');
  });

  test('wqGetPoolCount is defined', () => {
    expect(typeof wqGetPoolCount).toBe('function');
  });

  test('wqInitSheets runs without error', () => {
    expect(() => wqInitSheets()).not.toThrow();
  });

  test('wqGetHistory is defined', () => {
    expect(typeof wqGetHistory).toBe('function');
  });

  test('wqClosePoll is defined', () => {
    expect(typeof wqClosePoll).toBe('function');
  });

  test('wqGetPollFrequency is defined', () => {
    expect(typeof wqGetPollFrequency).toBe('function');
  });

  test('wqSetPollFrequency is defined', () => {
    expect(typeof wqSetPollFrequency).toBe('function');
  });

  test('wqManualDrawCommunityPoll is defined', () => {
    expect(typeof wqManualDrawCommunityPoll).toBe('function');
  });

  test('autoSelectCommunityPoll is defined', () => {
    expect(typeof autoSelectCommunityPoll).toBe('function');
  });

  test('setupCommunityPollTrigger is defined', () => {
    expect(typeof setupCommunityPollTrigger).toBe('function');
  });
});
