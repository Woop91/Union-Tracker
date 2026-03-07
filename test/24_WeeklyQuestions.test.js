/**
 * Tests for 24_WeeklyQuestions.gs
 *
 * Covers WeeklyQuestions IIFE: sheet setup, active questions, response
 * submission, pool questions, steward question setting, history, and
 * global wrappers.
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
});

// ============================================================================
// getQuestionStats
// ============================================================================

describe('WeeklyQuestions.getPoolCount', () => {
  test('returns a number', () => {
    const result = WeeklyQuestions.getPoolCount();
    expect(typeof result).toBe('number');
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
});
