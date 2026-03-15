/**
 * Tests for 26_QAForum.gs
 *
 * Covers Q&A Forum module: sheet initialization, question CRUD,
 * answers, upvoting, moderation, flagged content, and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '26_QAForum.gs']);

// ============================================================================
// Helpers
// ============================================================================

const FORUM_HEADERS = [
  'ID', 'Author Email', 'Author Name', 'Is Anonymous', 'Question Text',
  'Status', 'Upvote Count', 'Upvoters', 'Answer Count', 'Created', 'Updated'
];

const ANSWER_HEADERS = [
  'ID', 'Question ID', 'Author Email', 'Author Name', 'Is Steward',
  'Answer Text', 'Status', 'Created'
];

function makeForumData(rows) {
  return [FORUM_HEADERS, ...rows];
}

function makeAnswerData(rows) {
  return [ANSWER_HEADERS, ...rows];
}

function setupSheets(opts) {
  opts = opts || {};
  var sheets = [];
  if (opts.forumData) {
    sheets.push(createMockSheet(SHEETS.QA_FORUM, opts.forumData));
  }
  if (opts.answerData) {
    sheets.push(createMockSheet(SHEETS.QA_ANSWERS, opts.answerData));
  }
  var ss = createMockSpreadsheet(sheets);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

// ============================================================================
// initQAForumSheets
// ============================================================================

describe('QAForum.initQAForumSheets', () => {
  test('creates _QA_Forum sheet if missing', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.initQAForumSheets();

    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.QA_FORUM);
  });

  test('creates _QA_Answers sheet if missing', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.initQAForumSheets();

    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.QA_ANSWERS);
  });

  test('does not recreate existing sheets', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, [ANSWER_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet, answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.initQAForumSheets();

    expect(ss.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// getQuestions
// ============================================================================

describe('QAForum.getQuestions', () => {
  test('returns empty array when sheet is empty (header only)', () => {
    setupSheets({ forumData: [FORUM_HEADERS] });

    var result = QAForum.getQuestions('user@test.com');
    expect(result.questions).toEqual([]);
    expect(result.total).toBe(0);
  });

  test('returns empty array when sheet does not exist', () => {
    // No sheets at all — initQAForumSheets will be called, but the second
    // getSheetByName still returns null-equivalent with lastRow <= 1
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([]);
    // First call returns null (triggers init), second call returns sheet with only headers
    ss.getSheetByName
      .mockReturnValueOnce(null)
      .mockReturnValue(forumSheet);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.getQuestions('user@test.com');
    expect(result.questions).toEqual([]);
    expect(result.total).toBe(0);
  });

  test('returns questions from sheet data', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User One', false, 'How does this work?', 'active', 2, 'h1,h2', 1, now, now],
      ['QA_2', 'other@test.com', 'User Two', false, 'Another question?', 'active', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('user@test.com');
    expect(result.questions.length).toBe(2);
    expect(result.questions[0].questionText).toBeTruthy();
  });

  test('excludes deleted questions', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Active question', 'active', 0, '', 0, now, now],
      ['QA_2', 'user@test.com', 'User', false, 'Deleted question', 'deleted', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('user@test.com');
    expect(result.questions.length).toBe(1);
    expect(result.questions[0].id).toBe('QA_1');
  });

  test('supports recent sort (default)', () => {
    var older = new Date('2026-01-01T12:00:00Z');
    var newer = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Older question', 'active', 5, '', 0, older, older],
      ['QA_2', 'b@test.com', 'B', false, 'Newer question', 'active', 0, '', 0, newer, newer],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('a@test.com', 1, 20, 'recent');
    expect(result.questions[0].id).toBe('QA_2');
    expect(result.questions[1].id).toBe('QA_1');
  });

  test('supports popular sort', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Low votes', 'active', 1, 'h1', 0, now, now],
      ['QA_2', 'b@test.com', 'B', false, 'High votes', 'active', 10, 'h1,h2,h3', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('a@test.com', 1, 20, 'popular');
    expect(result.questions[0].id).toBe('QA_2');
    expect(result.questions[0].upvoteCount).toBe(10);
  });

  test('pagination works correctly (page, pageSize)', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var rows = [];
    for (var i = 0; i < 5; i++) {
      rows.push(['QA_' + i, 'u@test.com', 'U', false, 'Q' + i, 'active', 0, '', 0, now, now]);
    }
    setupSheets({ forumData: makeForumData(rows) });

    var page1 = QAForum.getQuestions('u@test.com', 1, 2);
    expect(page1.questions.length).toBe(2);
    expect(page1.total).toBe(5);
    expect(page1.page).toBe(1);
    expect(page1.pageSize).toBe(2);

    var page2 = QAForum.getQuestions('u@test.com', 2, 2);
    expect(page2.questions.length).toBe(2);

    var page3 = QAForum.getQuestions('u@test.com', 3, 2);
    expect(page3.questions.length).toBe(1);
  });

  test('returns total count', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Q1', 'active', 0, '', 0, now, now],
      ['QA_2', 'b@test.com', 'B', false, 'Q2', 'active', 0, '', 0, now, now],
      ['QA_3', 'c@test.com', 'C', false, 'Q3', 'deleted', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('a@test.com');
    expect(result.total).toBe(2); // excludes deleted
  });

  test('excludes resolved questions by default (showResolved=false)', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Active', 'active', 0, '', 0, now, now],
      ['QA_2', 'b@test.com', 'B', false, 'Resolved', 'resolved', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('a@test.com', 1, 20, 'recent');
    expect(result.questions.length).toBe(1);
    expect(result.questions[0].id).toBe('QA_1');
    expect(result.total).toBe(1);
  });

  test('includes resolved questions when showResolved=true', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Active', 'active', 0, '', 0, now, now],
      ['QA_2', 'b@test.com', 'B', false, 'Resolved', 'resolved', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('a@test.com', 1, 20, 'recent', true);
    expect(result.questions.length).toBe(2);
    expect(result.total).toBe(2);
  });

  test('sets isOwner correctly based on email', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'owner@test.com', 'Owner', false, 'My question', 'active', 0, '', 0, now, now],
      ['QA_2', 'other@test.com', 'Other', false, 'Their question', 'active', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData });

    var result = QAForum.getQuestions('owner@test.com');
    var own = result.questions.find(function (q) { return q.id === 'QA_1'; });
    var notOwn = result.questions.find(function (q) { return q.id === 'QA_2'; });
    expect(own.isOwner).toBe(true);
    expect(notOwn.isOwner).toBe(false);
  });
});

// ============================================================================
// getQuestionDetail
// ============================================================================

describe('QAForum.getQuestionDetail', () => {
  test('returns null for missing question', () => {
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Q1', 'active', 0, '', 0, new Date(), new Date()],
    ]);
    setupSheets({ forumData: forumData, answerData: [ANSWER_HEADERS] });

    var result = QAForum.getQuestionDetail('a@test.com', 'QA_NONEXISTENT');
    expect(result).toBeNull();
  });

  test('returns null for deleted question', () => {
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Q1', 'deleted', 0, '', 0, new Date(), new Date()],
    ]);
    setupSheets({ forumData: forumData, answerData: [ANSWER_HEADERS] });

    var result = QAForum.getQuestionDetail('a@test.com', 'QA_1');
    expect(result).toBeNull();
  });

  test('returns question data with answers', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'How does this work?', 'active', 2, 'h1,h2', 1, now, now],
    ]);
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 'steward@test.com', 'Steward', true, 'Like this.', 'active', now],
    ]);
    setupSheets({ forumData: forumData, answerData: answerData });

    var result = QAForum.getQuestionDetail('user@test.com', 'QA_1');
    expect(result).not.toBeNull();
    expect(result.id).toBe('QA_1');
    expect(result.questionText).toBe('How does this work?');
    expect(result.answers.length).toBe(1);
    expect(result.answers[0].answerText).toBe('Like this.');
  });

  test('excludes deleted answers', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Question', 'active', 0, '', 2, now, now],
    ]);
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'Steward', true, 'Good answer', 'active', now],
      ['ANS_2', 'QA_1', 't@test.com', 'Troll', false, 'Bad answer', 'deleted', now],
    ]);
    setupSheets({ forumData: forumData, answerData: answerData });

    var result = QAForum.getQuestionDetail('user@test.com', 'QA_1');
    expect(result.answers.length).toBe(1);
    expect(result.answers[0].id).toBe('ANS_1');
  });

  test('shows isSteward flag on answers', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Question', 'active', 0, '', 2, now, now],
    ]);
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'Steward', true, 'Steward reply', 'active', now],
      ['ANS_2', 'QA_1', 'm@test.com', 'Member', false, 'Member reply', 'active', now],
    ]);
    setupSheets({ forumData: forumData, answerData: answerData });

    var result = QAForum.getQuestionDetail('user@test.com', 'QA_1');
    expect(result.answers[0].isSteward).toBe(true);
    expect(result.answers[1].isSteward).toBe(false);
  });

  test('shows isAnonymous correctly', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', true, 'Anonymous Q', 'active', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData, answerData: [ANSWER_HEADERS] });

    var result = QAForum.getQuestionDetail('user@test.com', 'QA_1');
    expect(result.isAnonymous).toBe(true);
    expect(result.authorName).toBe('Anonymous');
  });
});

// ============================================================================
// Anonymous flag boolean normalization — regression for fa27b42
// Google Sheets stores booleans as boolean true/false OR as strings 'TRUE'/'FALSE'.
// isTruthyValue() must handle all representations consistently.
// ============================================================================

describe('QAForum anonymous flag — isTruthyValue normalization', () => {
  var now = new Date('2026-03-01T12:00:00Z');

  // Each pair: [Is Anonymous cell value, expected authorName, expected isAnonymous]
  var cases = [
    [true,    'Anonymous', true,  'boolean true'],
    ['TRUE',  'Anonymous', true,  'string TRUE'],
    ['True',  'Anonymous', true,  'string True'],
    ['yes',   'Anonymous', true,  'string yes'],
    ['1',     'Anonymous', true,  'string 1'],
    [false,   'Real Name', false, 'boolean false'],
    ['FALSE', 'Real Name', false, 'string FALSE'],
    ['false', 'Real Name', false, 'string false'],
    ['no',    'Real Name', false, 'string no'],
    ['',      'Real Name', false, 'empty string'],
  ];

  cases.forEach(function(c) {
    var anonValue = c[0], expectedName = c[1], expectedAnon = c[2], label = c[3];

    test('getQuestions: Is Anonymous=' + label + ' → isAnonymous=' + expectedAnon, function() {
      var forumData = makeForumData([
        ['QA_X', 'author@test.com', 'Real Name', anonValue, 'A question', 'active', 0, '', 0, now, now],
      ]);
      setupSheets({ forumData: forumData });

      var result = QAForum.getQuestions('user@test.com', 1, 10);
      expect(result.questions).toHaveLength(1);
      expect(result.questions[0].isAnonymous).toBe(expectedAnon);
      expect(result.questions[0].authorName).toBe(expectedName);
    });

    test('getQuestionDetail: Is Anonymous=' + label + ' → isAnonymous=' + expectedAnon, function() {
      var forumData = makeForumData([
        ['QA_Y', 'author@test.com', 'Real Name', anonValue, 'A question', 'active', 0, '', 0, now, now],
      ]);
      setupSheets({ forumData: forumData, answerData: [ANSWER_HEADERS] });

      var result = QAForum.getQuestionDetail('user@test.com', 'QA_Y');
      expect(result.isAnonymous).toBe(expectedAnon);
      expect(result.authorName).toBe(expectedName);
    });
  });
});

// ============================================================================
// submitQuestion
// ============================================================================

describe('QAForum.submitQuestion', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('rejects empty email', () => {
    var result = QAForum.submitQuestion('', 'Name', 'Some question?');
    expect(result.success).toBe(false);
  });

  test('rejects null email', () => {
    var result = QAForum.submitQuestion(null, 'Name', 'Some question?');
    expect(result.success).toBe(false);
  });

  test('rejects empty text', () => {
    var result = QAForum.submitQuestion('user@test.com', 'Name', '');
    expect(result.success).toBe(false);
    expect(result.message).toContain('required');
  });

  test('rejects whitespace-only text', () => {
    var result = QAForum.submitQuestion('user@test.com', 'Name', '   ');
    expect(result.success).toBe(false);
  });

  test('sanitizes text with escapeForFormula', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var spy = jest.spyOn(global, 'escapeForFormula');

    QAForum.submitQuestion('user@test.com', 'Name', '=IMPORTRANGE("evil")');

    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });

  test('truncates text to 2000 chars', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var longText = 'x'.repeat(3000);
    QAForum.submitQuestion('user@test.com', 'Name', longText);

    var appendCall = forumSheet.appendRow.mock.calls[0][0];
    // Column index 4 is the question text (0-indexed)
    expect(appendCall[4].length).toBeLessThanOrEqual(2000);
  });

  test('rate limits at 5 per hour', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Simulate cache returning count of 5 (already at limit)
    var mockCache = {
      get: jest.fn().mockReturnValue('5'),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValue(mockCache);

    var result = QAForum.submitQuestion('user@test.com', 'Name', 'Another question');
    expect(result.success).toBe(false);
    expect(result.message).toContain('Rate limit');

    // Restore default cache mock
    CacheService.getScriptCache.mockReturnValue({
      get: jest.fn(() => null),
      put: jest.fn(),
      remove: jest.fn()
    });
  });

  test('uses ScriptLock', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var mockLock = { tryLock: jest.fn(() => true), releaseLock: jest.fn() };
    LockService.getScriptLock.mockReturnValue(mockLock);

    QAForum.submitQuestion('user@test.com', 'Name', 'Question text?');

    expect(mockLock.tryLock).toHaveBeenCalledWith(10000);
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });

  test('returns {success:true, questionId}', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.submitQuestion('user@test.com', 'Name', 'Valid question?');
    expect(result.success).toBe(true);
    expect(result.questionId).toMatch(/^QA_/);
  });

  test('logs audit event', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.submitQuestion('user@test.com', 'Name', 'Audit question?');
    expect(logAuditEvent).toHaveBeenCalledWith(
      'QA_QUESTION_SUBMITTED',
      expect.stringContaining('QA_')
    );
  });
});

// ============================================================================
// submitAnswer
// ============================================================================

describe('QAForum.submitAnswer', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('rejects empty fields', () => {
    var result = QAForum.submitAnswer('', 'Name', 'QA_1', 'Answer text');
    expect(result.success).toBe(false);

    result = QAForum.submitAnswer('user@test.com', 'Name', '', 'Answer text');
    expect(result.success).toBe(false);

    result = QAForum.submitAnswer('user@test.com', 'Name', 'QA_1', '');
    expect(result.success).toBe(false);

    result = QAForum.submitAnswer('user@test.com', 'Name', 'QA_1', null);
    expect(result.success).toBe(false);
  });

  test('appends to answer sheet', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Question', 'active', 0, '', 0, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, [ANSWER_HEADERS]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var ss = createMockSpreadsheet([forumSheet, answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.submitAnswer('steward@test.com', 'Steward', 'QA_1', 'Here is the answer.', true);
    expect(result.success).toBe(true);
    expect(answerSheet.appendRow).toHaveBeenCalled();
  });

  test('increments answer count on question', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Question', 'active', 0, '', 0, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, [ANSWER_HEADERS]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);

    // Mock getRange to capture setValue calls
    var setValueCalls = [];
    var mockRange = {
      setValue: jest.fn(function (v) { setValueCalls.push(v); }),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);

    var ss = createMockSpreadsheet([forumSheet, answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.submitAnswer('steward@test.com', 'Steward', 'QA_1', 'Answer text', true);

    // Should call getRange for answer count (row 2, col 9) and updated (row 2, col 11)
    expect(forumSheet.getRange).toHaveBeenCalledWith(2, 9);
    expect(forumSheet.getRange).toHaveBeenCalledWith(2, 11);
    // First setValue should be new answer count (0 + 1 = 1)
    expect(setValueCalls[0]).toBe(1);
  });

  test('uses ScriptLock', () => {
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, [ANSWER_HEADERS]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Q', 'active', 0, '', 0, new Date(), new Date()],
    ]));
    var ss = createMockSpreadsheet([forumSheet, answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var mockLock = { tryLock: jest.fn(() => true), releaseLock: jest.fn() };
    LockService.getScriptLock.mockReturnValue(mockLock);

    QAForum.submitAnswer('u@test.com', 'U', 'QA_1', 'Answer text', true);

    expect(mockLock.tryLock).toHaveBeenCalledWith(10000);
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });

  test('returns {success:true, answerId}', () => {
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, [ANSWER_HEADERS]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Q', 'active', 0, '', 0, new Date(), new Date()],
    ]));
    var ss = createMockSpreadsheet([forumSheet, answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.submitAnswer('s@test.com', 'S', 'QA_1', 'Answer', true);
    expect(result.success).toBe(true);
    expect(result.answerId).toMatch(/^ANS_/);
  });
});

// ============================================================================
// upvoteQuestion
// ============================================================================

describe('QAForum.upvoteQuestion', () => {
  test('rejects missing email', () => {
    var result = QAForum.upvoteQuestion('', 'QA_1');
    expect(result.success).toBe(false);
  });

  test('rejects missing questionId', () => {
    var result = QAForum.upvoteQuestion('user@test.com', '');
    expect(result.success).toBe(false);
  });

  test('adds vote and returns upvoted:true', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'author@test.com', 'Author', false, 'Question', 'active', 0, '', 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var setValueCalls = [];
    var mockRange = {
      setValue: jest.fn(function (v) { setValueCalls.push(v); }),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.upvoteQuestion('voter@test.com', 'QA_1');
    expect(result.success).toBe(true);
    expect(result.upvoted).toBe(true);
    expect(result.newCount).toBe(1);
  });

  test('toggles vote off and returns upvoted:false', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    // Compute the hash for voter@test.com the same way the module does
    var emailHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'voter@test.com')
      .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 12);

    var forumData = makeForumData([
      ['QA_1', 'author@test.com', 'Author', false, 'Question', 'active', 1, emailHash, 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var mockRange = {
      setValue: jest.fn(),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.upvoteQuestion('voter@test.com', 'QA_1');
    expect(result.success).toBe(true);
    expect(result.upvoted).toBe(false);
    expect(result.newCount).toBe(0);
  });

  test('uses email hash for privacy (not raw email)', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'author@test.com', 'Author', false, 'Question', 'active', 0, '', 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var savedUpvoters = null;
    var mockRange = {
      setValue: jest.fn(function (v) {
        // Capture the upvoters string (column 8 call)
        if (typeof v === 'string') savedUpvoters = v;
      }),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.upvoteQuestion('voter@test.com', 'QA_1');

    // The stored upvoter should be a hash, not the raw email
    expect(savedUpvoters).not.toContain('voter@test.com');
    expect(savedUpvoters).toBeTruthy();
    expect(savedUpvoters.length).toBe(12); // 12-char hex substring
  });

  test('returns not found for missing question', () => {
    var forumData = makeForumData([
      ['QA_1', 'a@test.com', 'A', false, 'Q', 'active', 0, '', 0, new Date(), new Date()],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.upvoteQuestion('voter@test.com', 'QA_NONEXISTENT');
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });
});

// ============================================================================
// moderateQuestion
// ============================================================================

describe('QAForum.moderateQuestion', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('sets status to deleted on delete action', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Bad question', 'active', 0, '', 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var lastSetValue = null;
    var mockRange = {
      setValue: jest.fn(function (v) { lastSetValue = v; }),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateQuestion('steward@test.com', 'QA_1', 'delete');
    expect(result.success).toBe(true);
    // First getRange call is for status column (row 2, col 6)
    expect(forumSheet.getRange).toHaveBeenCalledWith(2, 6);
    // First setValue call should be 'deleted'
    expect(mockRange.setValue.mock.calls[0][0]).toBe('deleted');
  });

  test('sets status to flagged on flag action', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Suspicious question', 'active', 0, '', 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var mockRange = {
      setValue: jest.fn(),
      setValues: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateQuestion('steward@test.com', 'QA_1', 'flag');
    expect(result.success).toBe(true);
    expect(mockRange.setValue.mock.calls[0][0]).toBe('flagged');
  });

  test('rejects missing params', () => {
    var result = QAForum.moderateQuestion('', 'QA_1', 'delete');
    expect(result.success).toBe(false);

    result = QAForum.moderateQuestion('steward@test.com', '', 'delete');
    expect(result.success).toBe(false);

    result = QAForum.moderateQuestion('steward@test.com', 'QA_1', '');
    expect(result.success).toBe(false);
  });

  test('logs audit event', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'user@test.com', 'User', false, 'Question', 'active', 0, '', 0, now, now],
    ]);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    forumSheet.getRange.mockReturnValue({
      setValue: jest.fn(), setValues: jest.fn(),
      getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    });
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.moderateQuestion('steward@test.com', 'QA_1', 'delete');
    expect(logAuditEvent).toHaveBeenCalledWith(
      'QA_QUESTION_MODERATED',
      expect.stringContaining('QA_1')
    );
  });

  test('returns not found for missing question', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Q', 'active', 0, '', 0, new Date(), new Date()],
    ]));
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateQuestion('steward@test.com', 'QA_MISSING', 'delete');
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });
});

// ============================================================================
// moderateAnswer
// ============================================================================

describe('QAForum.moderateAnswer', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('sets status to deleted on delete action', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'Steward', true, 'Bad answer', 'active', now],
    ]);
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Question', 'active', 0, '', 1, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, answerData);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var mockRange = {
      setValue: jest.fn(), setValues: jest.fn(),
      getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    };
    answerSheet.getRange.mockReturnValue(mockRange);
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([answerSheet, forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateAnswer('steward@test.com', 'ANS_1', 'delete');
    expect(result.success).toBe(true);
    expect(answerSheet.getRange).toHaveBeenCalledWith(2, 7);
    expect(mockRange.setValue).toHaveBeenCalledWith('deleted');
  });

  test('decrements answer count on parent question when deleting an active answer', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'Steward', true, 'Answer to delete', 'active', now],
    ]);
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Question', 'active', 0, '', 2, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, answerData);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var setValueCalls = { answer: [], forum: [] };
    answerSheet.getRange.mockReturnValue({
      setValue: jest.fn(function(v) { setValueCalls.answer.push(v); }),
      setValues: jest.fn(), getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    });
    forumSheet.getRange.mockReturnValue({
      setValue: jest.fn(function(v) { setValueCalls.forum.push(v); }),
      setValues: jest.fn(), getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    });
    var ss = createMockSpreadsheet([answerSheet, forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.moderateAnswer('steward@test.com', 'ANS_1', 'delete');

    // Answer count should be decremented from 2 to 1
    expect(forumSheet.getRange).toHaveBeenCalledWith(2, 9);
    expect(setValueCalls.forum).toContain(1);
  });

  test('does not change answer count when flagging (not deleting)', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'Flaggable', 'active', now],
    ]);
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Question', 'active', 0, '', 1, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, answerData);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    answerSheet.getRange.mockReturnValue({
      setValue: jest.fn(), setValues: jest.fn(),
      getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    });
    var ss = createMockSpreadsheet([answerSheet, forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.moderateAnswer('steward@test.com', 'ANS_1', 'flag');

    // Forum sheet should NOT have getRange called for answer count update
    expect(forumSheet.getRange).not.toHaveBeenCalled();
  });

  test('sets status to flagged on flag action', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'Suspicious', 'active', now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, answerData);
    var mockRange = {
      setValue: jest.fn(), setValues: jest.fn(),
      getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    };
    answerSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateAnswer('steward@test.com', 'ANS_1', 'flag');
    expect(result.success).toBe(true);
    expect(mockRange.setValue).toHaveBeenCalledWith('flagged');
  });

  test('rejects missing params', () => {
    var result = QAForum.moderateAnswer('', 'ANS_1', 'delete');
    expect(result.success).toBe(false);

    result = QAForum.moderateAnswer('steward@test.com', '', 'delete');
    expect(result.success).toBe(false);

    result = QAForum.moderateAnswer('steward@test.com', 'ANS_1', '');
    expect(result.success).toBe(false);
  });

  test('logs audit event', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'Answer', 'active', now],
    ]);
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Q', 'active', 0, '', 1, now, now],
    ]);
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, answerData);
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, forumData);
    var mockRange = {
      setValue: jest.fn(), setValues: jest.fn(),
      getValue: jest.fn(() => ''), getValues: jest.fn(() => [['']])
    };
    answerSheet.getRange.mockReturnValue(mockRange);
    forumSheet.getRange.mockReturnValue(mockRange);
    var ss = createMockSpreadsheet([answerSheet, forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    QAForum.moderateAnswer('steward@test.com', 'ANS_1', 'delete');
    expect(logAuditEvent).toHaveBeenCalledWith(
      'QA_ANSWER_MODERATED',
      expect.stringContaining('ANS_1')
    );
  });

  test('returns not found for missing answer', () => {
    var answerSheet = createMockSheet(SHEETS.QA_ANSWERS, makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'A', 'active', new Date()],
    ]));
    var ss = createMockSpreadsheet([answerSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = QAForum.moderateAnswer('steward@test.com', 'ANS_MISSING', 'delete');
    expect(result.success).toBe(false);
    expect(result.message).toContain('not found');
  });
});

// ============================================================================
// getFlaggedContent
// ============================================================================

describe('QAForum.getFlaggedContent', () => {
  test('returns empty when no flagged content', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Normal Q', 'active', 0, '', 0, now, now],
    ]);
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'Normal A', 'active', now],
    ]);
    setupSheets({ forumData: forumData, answerData: answerData });

    var result = QAForum.getFlaggedContent('steward@test.com');
    expect(result.questions).toEqual([]);
    expect(result.answers).toEqual([]);
  });

  test('returns flagged questions', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var forumData = makeForumData([
      ['QA_1', 'u@test.com', 'U', false, 'Flagged Q', 'flagged', 0, '', 0, now, now],
      ['QA_2', 'u@test.com', 'U', false, 'Normal Q', 'active', 0, '', 0, now, now],
    ]);
    setupSheets({ forumData: forumData, answerData: [ANSWER_HEADERS] });

    var result = QAForum.getFlaggedContent('steward@test.com');
    expect(result.questions.length).toBe(1);
    expect(result.questions[0].id).toBe('QA_1');
    expect(result.questions[0].questionText).toContain('Flagged');
  });

  test('returns flagged answers', () => {
    var now = new Date('2026-03-01T12:00:00Z');
    var answerData = makeAnswerData([
      ['ANS_1', 'QA_1', 's@test.com', 'S', false, 'Flagged answer', 'flagged', now],
      ['ANS_2', 'QA_1', 'm@test.com', 'M', false, 'Normal answer', 'active', now],
    ]);
    setupSheets({ forumData: [FORUM_HEADERS], answerData: answerData });

    var result = QAForum.getFlaggedContent('steward@test.com');
    expect(result.answers.length).toBe(1);
    expect(result.answers[0].id).toBe('ANS_1');
    expect(result.answers[0].answerText).toContain('Flagged');
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('qaGetQuestions delegates to QAForum.getQuestions', () => {
    var forumSheet = createMockSheet(SHEETS.QA_FORUM, [FORUM_HEADERS]);
    var ss = createMockSpreadsheet([forumSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var spy = jest.spyOn(QAForum, 'getQuestions');
    qaGetQuestions('test-session-token', 1, 10, 'recent', true);
    // Wrapper resolves email via _resolveCallerEmail(sessionToken) — passes resolved email + showResolved
    expect(spy).toHaveBeenCalledWith('test@example.com', 1, 10, 'recent', true);
    spy.mockRestore();
  });

  test('qaSubmitQuestion delegates to QAForum.submitQuestion', () => {
    var spy = jest.spyOn(QAForum, 'submitQuestion').mockReturnValue({ success: true, questionId: 'QA_test' });
    var result = qaSubmitQuestion('test-session-token', 'Name', 'Question?', false);
    // Wrapper resolves email via _resolveCallerEmail(sessionToken) — passes resolved email
    expect(spy).toHaveBeenCalledWith('test@example.com', 'Name', 'Question?', false);
    expect(result.success).toBe(true);
    spy.mockRestore();
  });
});
