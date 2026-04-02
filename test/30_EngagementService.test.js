/**
 * Tests for 30_EngagementService.gs
 *
 * Covers EngagementService module: score computation for each dimension
 * (survey, meeting, Q&A, workload, contact freshness), composite scoring,
 * scoreboard retrieval, unit aggregation, and member lookup.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '30_EngagementService.gs']);

// ============================================================================
// Module existence
// ============================================================================

describe('EngagementService module existence', () => {
  test('EngagementService is defined', () => {
    expect(EngagementService).toBeDefined();
  });

  test('exposes expected public API methods', () => {
    expect(typeof EngagementService.initSheet).toBe('function');
    expect(typeof EngagementService.computeScoreForMember).toBe('function');
    expect(typeof EngagementService.computeAllScores).toBe('function');
    expect(typeof EngagementService.getScoreboard).toBe('function');
    expect(typeof EngagementService.getScoreByUnit).toBe('function');
    expect(typeof EngagementService.getMemberScore).toBe('function');
    expect(typeof EngagementService.getMemberReportCard).toBe('function');
  });

  test('global wrappers exist', () => {
    expect(typeof dataGetEngagementScoreboard).toBe('function');
    expect(typeof dataGetEngagementByUnit).toBe('function');
    expect(typeof dataComputeEngagementScores).toBe('function');
    expect(typeof dataGetMyReportCard).toBe('function');
    expect(typeof initEngagementSheet).toBe('function');
  });
});

// ============================================================================
// computeScoreForMember — pure computation logic
// ============================================================================

describe('EngagementService.computeScoreForMember', () => {
  test('returns default scores with empty cachedData', () => {
    var result = EngagementService.computeScoreForMember('user@test.com', {});
    expect(result).toBeDefined();
    expect(result.email).toBe('user@test.com');
    expect(result.scores).toBeDefined();
    expect(typeof result.composite).toBe('number');
    // Survey defaults to 50 when no survey data, workload defaults to 50
    expect(result.scores.survey).toBe(50);
    expect(result.scores.workload).toBe(50);
    // QA with no data should be 0
    expect(result.scores.qa).toBe(0);
    // Contact with no data should be 0
    expect(result.scores.contact).toBe(0);
    // Meeting with 0 totalMeetings90d defaults to 50
    expect(result.scores.meeting).toBe(50);
  });

  test('returns default scores with null cachedData', () => {
    var result = EngagementService.computeScoreForMember('user@test.com', null);
    expect(result).toBeDefined();
    expect(result.scores.survey).toBe(50);
  });

  test('computes survey score from survey data', () => {
    var surveyData = [
      // [col0, col1, EMAIL(col2), col3, col4, STATUS(col5)]
      ['', '', 'user@test.com', '', '', 'completed'],
      ['', '', 'user@test.com', '', '', 'pending'],
      ['', '', 'user@test.com', '', '', 'completed'],
      ['', '', 'other@test.com', '', '', 'completed'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', { surveyData: surveyData });
    // 2 completed / 3 total = 67%
    expect(result.scores.survey).toBe(67);
  });

  test('computes QA score from posts and answers', () => {
    var qaData = [
      // [col0, AUTHOR_EMAIL(col1)]
      ['', 'user@test.com'],
      ['', 'user@test.com'],
      ['', 'other@test.com'],
    ];
    var answerData = [
      // [col0, col1, AUTHOR_EMAIL(col2)]
      ['', '', 'user@test.com'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', {
      qaData: qaData,
      answerData: answerData,
    });
    // 2 posts * 10 + 1 answer * 15 = 35, capped at 100
    expect(result.scores.qa).toBe(35);
  });

  test('QA score is capped at 100', () => {
    var qaData = [];
    for (var i = 0; i < 15; i++) {
      qaData.push(['', 'user@test.com']);
    }
    var answerData = [];
    for (var j = 0; j < 10; j++) {
      answerData.push(['', '', 'user@test.com']);
    }
    var result = EngagementService.computeScoreForMember('user@test.com', {
      qaData: qaData,
      answerData: answerData,
    });
    // 15 * 10 + 10 * 15 = 300, capped at 100
    expect(result.scores.qa).toBe(100);
  });

  test('computes contact score based on freshness', () => {
    var now = new Date();
    // Contact 5 days ago => 100
    var recentContact = new Date(now.getTime() - 5 * 86400000);
    var contactData = [
      // [TIMESTAMP(col0), col1, MEMBER_EMAIL(col2)]
      [recentContact, '', 'user@test.com'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', {
      contactData: contactData,
    });
    expect(result.scores.contact).toBe(100);
  });

  test('contact score degrades with older contacts', () => {
    var now = new Date();
    // Contact 45 days ago => 50
    var oldContact = new Date(now.getTime() - 45 * 86400000);
    var contactData = [
      [oldContact, '', 'user@test.com'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', {
      contactData: contactData,
    });
    expect(result.scores.contact).toBe(50);
  });

  test('contact score is 0 for very old contacts (>90 days)', () => {
    var now = new Date();
    var veryOldContact = new Date(now.getTime() - 120 * 86400000);
    var contactData = [
      [veryOldContact, '', 'user@test.com'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', {
      contactData: contactData,
    });
    expect(result.scores.contact).toBe(0);
  });

  test('composite score is weighted average', () => {
    // With all defaults (survey=50, meeting=50, qa=0, workload=50, contact=0):
    // composite = (50*25 + 50*25 + 0*15 + 50*15 + 0*20) / 100 = (1250+1250+0+750+0)/100 = 33
    var result = EngagementService.computeScoreForMember('user@test.com', {});
    expect(result.composite).toBe(33); // (1250+1250+750)/100 = 32.5 rounds to 33
  });

  test('handles email case-insensitivity', () => {
    var surveyData = [
      ['', '', 'USER@TEST.COM', '', '', 'completed'],
    ];
    var result = EngagementService.computeScoreForMember('user@test.com', {
      surveyData: surveyData,
    });
    // 1/1 = 100%
    expect(result.scores.survey).toBe(100);
  });
});

// ============================================================================
// getScoreboard
// ============================================================================

describe('EngagementService.getScoreboard', () => {
  test('returns empty when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    var result = EngagementService.getScoreboard();
    expect(result.items).toEqual([]);
    expect(result.totalRows).toBe(0);
  });

  test('returns empty when sheet does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    var result = EngagementService.getScoreboard();
    expect(result.items).toEqual([]);
  });

  test('returns paginated items from score sheet', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice Smith', 'Unit A', 80, 70, 60, 50, 90, 72, new Date(), 'up'],
      ['bob@test.com', 'Bob Jones', 'Unit B', 40, 30, 20, 10, 50, 30, new Date(), 'down'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getScoreboard();
    expect(result.items.length).toBe(2);
    expect(result.totalRows).toBe(2);
    expect(result.items[0].composite).toBeGreaterThanOrEqual(result.items[1].composite);
  });

  test('respects filterUnit option', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice Smith', 'Unit A', 80, 70, 60, 50, 90, 72, new Date(), 'up'],
      ['bob@test.com', 'Bob Jones', 'Unit B', 40, 30, 20, 10, 50, 30, new Date(), 'down'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getScoreboard({ filterUnit: 'Unit A' });
    expect(result.items.length).toBe(1);
    expect(result.items[0].email).toBe('alice@test.com');
  });

  test('respects searchTerm option', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice Smith', 'Unit A', 80, 70, 60, 50, 90, 72, new Date(), 'up'],
      ['bob@test.com', 'Bob Jones', 'Unit B', 40, 30, 20, 10, 50, 30, new Date(), 'down'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getScoreboard({ searchTerm: 'bob' });
    expect(result.items.length).toBe(1);
    expect(result.items[0].email).toBe('bob@test.com');
  });
});

// ============================================================================
// getScoreByUnit
// ============================================================================

describe('EngagementService.getScoreByUnit', () => {
  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(EngagementService.getScoreByUnit()).toEqual([]);
  });

  test('returns unit aggregations from score sheet', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice', 'Unit A', 80, 70, 60, 50, 90, 70, new Date(), 'up'],
      ['bob@test.com', 'Bob', 'Unit A', 80, 70, 60, 50, 90, 80, new Date(), 'up'],
      ['carol@test.com', 'Carol', 'Unit B', 40, 30, 20, 10, 50, 40, new Date(), 'down'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getScoreByUnit();
    expect(result.length).toBe(2);
    // Sorted alphabetically
    expect(result[0].unit).toBe('Unit A');
    expect(result[0].memberCount).toBe(2);
    expect(result[0].avgScore).toBe(75); // (70+80)/2
    expect(result[1].unit).toBe('Unit B');
    expect(result[1].memberCount).toBe(1);
  });
});

// ============================================================================
// getMemberScore
// ============================================================================

describe('EngagementService.getMemberScore', () => {
  test('returns null when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(EngagementService.getMemberScore('user@test.com')).toBeNull();
  });

  test('returns null when member not found', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [HEADERS, ['other@test.com', 'Other', 'Unit A', 50, 50, 50, 50, 50, 50, new Date(), 'stable']];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(EngagementService.getMemberScore('notfound@test.com')).toBeNull();
  });

  test('returns score breakdown for known member', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice Smith', 'Unit A', 80, 70, 60, 50, 90, 72, new Date(), 'up'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getMemberScore('alice@test.com');
    expect(result).not.toBeNull();
    expect(result.email).toBe('alice@test.com');
    expect(result.name).toBe('Alice Smith');
    expect(result.scores.survey).toBe(80);
    expect(result.composite).toBe(72);
    expect(result.trend).toBe('up');
  });

  test('is case-insensitive on email lookup', () => {
    var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
      'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];
    var data = [
      HEADERS,
      ['alice@test.com', 'Alice', 'Unit A', 80, 70, 60, 50, 90, 72, new Date(), 'up'],
    ];
    var sheet = createMockSheet(SHEETS.ENGAGEMENT_SCORES, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = EngagementService.getMemberScore('ALICE@TEST.COM');
    expect(result).not.toBeNull();
  });
});

// ============================================================================
// Global wrappers — auth gating
// ============================================================================

describe('Engagement global wrappers', () => {
  test('dataGetEngagementScoreboard requires steward auth', () => {
    _requireStewardAuth.mockReturnValueOnce(null);
    var result = dataGetEngagementScoreboard('bad-token');
    expect(result.items).toEqual([]);
  });

  test('dataGetEngagementByUnit requires steward auth', () => {
    _requireStewardAuth.mockReturnValueOnce(null);
    var result = dataGetEngagementByUnit('bad-token');
    expect(result).toEqual([]);
  });

  test('dataComputeEngagementScores requires steward auth', () => {
    _requireStewardAuth.mockReturnValueOnce(null);
    var result = dataComputeEngagementScores('bad-token');
    expect(result.success).toBe(false);
  });

  test('dataGetMyReportCard requires authenticated caller', () => {
    _resolveCallerEmail.mockReturnValueOnce(null);
    var result = dataGetMyReportCard('bad-token');
    expect(result.success).toBe(false);
    expect(result.authError).toBe(true);
  });
});
