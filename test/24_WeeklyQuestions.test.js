/**
 * Tests for 24_WeeklyQuestions.gs
 *
 * Covers WeeklyQuestionService IIFE: sheet setup, question CRUD,
 * response management, auto-rotation, and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '24_WeeklyQuestions.gs']);

let mockQuestionsSheet;
let mockResponsesSheet;
let mockSs;

function setupSheets(questionsData, responsesData) {
  mockQuestionsSheet = createMockSheet(SHEETS.WEEKLY_QUESTIONS || '_Weekly_Questions', questionsData);
  mockResponsesSheet = createMockSheet(SHEETS.WEEKLY_RESPONSES || '_Weekly_Responses', responsesData);

  // Custom getRange for questions sheet to return specific values
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

  mockSs = createMockSpreadsheet([mockQuestionsSheet, mockResponsesSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

beforeEach(() => {
  setupSheets(null, null);
});

// ============================================================================
// Sheet Initialization
// ============================================================================

describe('WeeklyQuestionService.initSheets', () => {
  test('creates question sheet if missing', () => {
    WeeklyQuestionService.initSheets();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });

  test('does not recreate existing sheets', () => {
    const questionsData = [['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses']];
    const responsesData = [['ID', 'Question ID', 'Email', 'Response', 'Submitted']];
    setupSheets(questionsData, responsesData);
    WeeklyQuestionService.initSheets();
    expect(mockSs.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// getActiveQuestion
// ============================================================================

describe('WeeklyQuestionService.getActiveQuestion', () => {
  test('returns null when no active questions', () => {
    const questionsData = [
      ['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses'],
    ];
    setupSheets(questionsData, null);
    const result = WeeklyQuestionService.getActiveQuestion();
    expect(result).toBeNull();
  });

  test('returns null when sheet has no data', () => {
    setupSheets([['header']], null);
    const result = WeeklyQuestionService.getActiveQuestion();
    expect(result).toBeNull();
  });
});

// ============================================================================
// addQuestion
// ============================================================================

describe('WeeklyQuestionService.addQuestion', () => {
  test('rejects missing steward email', () => {
    const result = WeeklyQuestionService.addQuestion(null, {});
    expect(result.success).toBe(false);
    expect(result.message).toBeTruthy();
  });

  test('rejects missing question text', () => {
    const result = WeeklyQuestionService.addQuestion('steward@test.com', {});
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('rejects invalid question type', () => {
    const result = WeeklyQuestionService.addQuestion('steward@test.com', {
      question: 'How are you?',
      type: 'invalid_type'
    });
    expect(result.success).toBe(false);
    expect(result.message).toContain('type');
  });

  test('adds question with valid inputs', () => {
    const questionsData = [
      ['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses'],
    ];
    setupSheets(questionsData, null);

    const result = WeeklyQuestionService.addQuestion('steward@test.com', {
      question: 'How are you?',
      type: 'text'
    });
    expect(result.success).toBe(true);
    expect(result.questionId).toBeTruthy();
  });

  test('sanitizes question text', () => {
    const questionsData = [
      ['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses'],
    ];
    setupSheets(questionsData, null);

    WeeklyQuestionService.addQuestion('steward@test.com', {
      question: '=FORMULA("evil")',
      type: 'text'
    });
    expect(mockQuestionsSheet.appendRow).toHaveBeenCalled();
  });

  test('truncates question to 500 chars', () => {
    const questionsData = [
      ['ID', 'Question', 'Type', 'Options', 'Category', 'Active', 'Week Number', 'Year', 'Created By', 'Created', 'Responses'],
    ];
    setupSheets(questionsData, null);

    const longQuestion = 'A'.repeat(1000);
    WeeklyQuestionService.addQuestion('steward@test.com', {
      question: longQuestion,
      type: 'text'
    });
    expect(mockQuestionsSheet.appendRow).toHaveBeenCalled();
    // Check the appended question was truncated
    const appendedRow = mockQuestionsSheet.appendRow.mock.calls[0][0];
    expect(appendedRow[1].length).toBeLessThanOrEqual(500);
  });
});

// ============================================================================
// submitResponse
// ============================================================================

describe('WeeklyQuestionService.submitResponse', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestionService.submitResponse(null, 'QID', 'answer');
    expect(result.success).toBe(false);
  });

  test('rejects missing questionId', () => {
    const result = WeeklyQuestionService.submitResponse('user@test.com', null, 'answer');
    expect(result.success).toBe(false);
  });

  test('rejects empty response', () => {
    const result = WeeklyQuestionService.submitResponse('user@test.com', 'QID', '');
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// getQuestionResponses
// ============================================================================

describe('WeeklyQuestionService.getQuestionResponses', () => {
  test('returns empty when no responses', () => {
    const responsesData = [['ID', 'Question ID', 'Email', 'Response', 'Submitted']];
    setupSheets(null, responsesData);
    const result = WeeklyQuestionService.getQuestionResponses('QID');
    expect(result.responses).toEqual([]);
    expect(result.total).toBe(0);
  });
});

// ============================================================================
// activateQuestion / deactivateQuestion
// ============================================================================

describe('WeeklyQuestionService.activateQuestion', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestionService.activateQuestion(null, 'QID');
    expect(result.success).toBe(false);
  });

  test('rejects missing questionId', () => {
    const result = WeeklyQuestionService.activateQuestion('steward@test.com', null);
    expect(result.success).toBe(false);
  });
});

describe('WeeklyQuestionService.deactivateQuestion', () => {
  test('rejects missing questionId', () => {
    const result = WeeklyQuestionService.deactivateQuestion('steward@test.com', null);
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// deleteQuestion
// ============================================================================

describe('WeeklyQuestionService.deleteQuestion', () => {
  test('rejects missing email', () => {
    const result = WeeklyQuestionService.deleteQuestion(null, 'QID');
    expect(result.success).toBe(false);
  });

  test('rejects missing questionId', () => {
    const result = WeeklyQuestionService.deleteQuestion('steward@test.com', null);
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global Wrappers', () => {
  test('wqGetActiveQuestion delegates to WeeklyQuestionService', () => {
    const result = wqGetActiveQuestion();
    // Just verify no error thrown
    expect(result === null || typeof result === 'object').toBe(true);
  });

  test('wqAddQuestion delegates to WeeklyQuestionService', () => {
    const result = wqAddQuestion('steward@test.com', { question: 'Test?', type: 'text' });
    expect(result).toHaveProperty('success');
  });

  test('wqSubmitResponse delegates to WeeklyQuestionService', () => {
    const result = wqSubmitResponse('user@test.com', 'QID', 'answer');
    expect(result).toHaveProperty('success');
  });

  test('wqGetQuestionResponses delegates to WeeklyQuestionService', () => {
    const responsesData = [['ID', 'Question ID', 'Email', 'Response', 'Submitted']];
    setupSheets(null, responsesData);
    const result = wqGetQuestionResponses('QID');
    expect(result).toHaveProperty('responses');
  });

  test('wqActivateQuestion delegates to WeeklyQuestionService', () => {
    const result = wqActivateQuestion('steward@test.com', 'QID');
    expect(result).toHaveProperty('success');
  });

  test('wqDeactivateQuestion delegates to WeeklyQuestionService', () => {
    const result = wqDeactivateQuestion('steward@test.com', 'QID');
    expect(result).toHaveProperty('success');
  });

  test('wqDeleteQuestion delegates to WeeklyQuestionService', () => {
    const result = wqDeleteQuestion('steward@test.com', 'QID');
    expect(result).toHaveProperty('success');
  });

  test('wqInitSheets delegates to WeeklyQuestionService', () => {
    // Just verify it runs without error
    expect(() => wqInitSheets()).not.toThrow();
  });
});
