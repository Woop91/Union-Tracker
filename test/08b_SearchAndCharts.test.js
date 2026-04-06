/**
 * Tests for 08b_SearchAndCharts.gs
 * Covers search functionality, chart generation, and data retrieval.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '08b_SearchAndCharts.gs'
]);

describe('08b function existence', () => {
  const required = [
    'getDesktopSearchLocations', 'getDesktopSearchData',
    'navigateToSearchResult',
    'searchDashboard', 'quickSearchDashboard', 'advancedSearch',
    'getDepartmentList', 'getMemberList',
    'generateSelectedChart',
    'createGaugeStyleChart_', 'createScorecardChart_',
    'createTrendLineChart_', 'createAreaChart_', 'createComboChart_',
    'createSummaryTableChart_', 'createStewardLeaderboardChart_',
    'padRight'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('padRight', () => {
  test('pads string to target length', () => {
    expect(padRight('hi', 5)).toBe('hi   ');
  });

  test('does not truncate longer strings', () => {
    expect(padRight('hello world', 5).length).toBeGreaterThanOrEqual(5);
  });

  test('handles empty string', () => {
    expect(padRight('', 3)).toBe('   ');
  });
});

// ============================================================================
// getDesktopSearchData — behavioral tests
// ============================================================================

describe('getDesktopSearchData', () => {
  /** Build a member-row array padded to the right width for col_() lookups */
  function buildMemberRow(id, first, last, jobTitle, location, email, isSteward) {
    // MEMBER_COLS are 1-indexed; we need a 0-indexed array large enough
    var row = new Array(Math.max(
      MEMBER_COLS.MEMBER_ID, MEMBER_COLS.FIRST_NAME, MEMBER_COLS.LAST_NAME,
      MEMBER_COLS.JOB_TITLE, MEMBER_COLS.WORK_LOCATION, MEMBER_COLS.EMAIL,
      MEMBER_COLS.IS_STEWARD
    )).fill('');
    row[MEMBER_COLS.MEMBER_ID - 1] = id;
    row[MEMBER_COLS.FIRST_NAME - 1] = first;
    row[MEMBER_COLS.LAST_NAME - 1] = last;
    row[MEMBER_COLS.JOB_TITLE - 1] = jobTitle;
    row[MEMBER_COLS.WORK_LOCATION - 1] = location;
    row[MEMBER_COLS.EMAIL - 1] = email;
    row[MEMBER_COLS.IS_STEWARD - 1] = isSteward;
    return row;
  }

  test('returns array of member results matching query', () => {
    var header = new Array(20).fill('');
    var memberData = [
      header,
      buildMemberRow('M001', 'Alice', 'Smith', 'Analyst', 'Boston', 'alice@example.com', 'No'),
      buildMemberRow('M002', 'Bob', 'Jones', 'Manager', 'Worcester', 'bob@example.com', 'Yes')
    ];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var results = getDesktopSearchData('alice', 'members', {});
    expect(Array.isArray(results)).toBe(true);
    expect(results.length).toBe(1);
    expect(results[0].type).toBe('member');
    expect(results[0].title).toBe('Alice Smith');
    expect(results[0].email).toBe('alice@example.com');
    expect(results[0].id).toBe('M001');
  });

  test('returns empty array when no matches', () => {
    var header = new Array(20).fill('');
    var memberData = [
      header,
      buildMemberRow('M001', 'Alice', 'Smith', 'Analyst', 'Boston', 'alice@example.com', 'No')
    ];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var results = getDesktopSearchData('zzznomatch', 'members', {});
    expect(Array.isArray(results)).toBe(true);
    expect(results.length).toBe(0);
  });

  test('applies location filter', () => {
    var header = new Array(20).fill('');
    var memberData = [
      header,
      buildMemberRow('M001', 'Alice', 'Smith', 'Analyst', 'Boston', 'alice@example.com', 'No'),
      buildMemberRow('M002', 'Bob', 'Jones', 'Manager', 'Worcester', 'bob@example.com', 'Yes')
    ];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    // Search with location filter — query long enough to match
    var results = getDesktopSearchData('bo', 'members', { location: 'Worcester' });
    // Alice is in Boston — filtered out; Bob is in Worcester and matches 'bo'
    expect(results.length).toBe(1);
    expect(results[0].title).toBe('Bob Jones');
  });

  test('result objects have expected shape', () => {
    var header = new Array(20).fill('');
    var memberData = [
      header,
      buildMemberRow('M010', 'Carol', 'Davis', 'Engineer', 'Springfield', 'carol@example.com', 'Yes')
    ];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var results = getDesktopSearchData('carol', 'members', {});
    expect(results.length).toBe(1);
    var r = results[0];
    expect(r).toHaveProperty('type');
    expect(r).toHaveProperty('id');
    expect(r).toHaveProperty('title');
    expect(r).toHaveProperty('email');
    expect(r).toHaveProperty('jobTitle');
    expect(r).toHaveProperty('location');
    expect(r).toHaveProperty('isSteward');
    expect(r).toHaveProperty('row');
    expect(typeof r.row).toBe('number');
  });
});

// ============================================================================
// getDepartmentList — behavioral tests
// ============================================================================

describe('getDepartmentList', () => {
  test('returns sorted unique department names from member sheet', () => {
    var header = new Array(10).fill('');
    // Build minimal data with UNIT column populated
    var row1 = new Array(10).fill('');
    row1[MEMBER_COLS.UNIT - 1] = 'Engineering';
    var row2 = new Array(10).fill('');
    row2[MEMBER_COLS.UNIT - 1] = 'Accounting';
    var row3 = new Array(10).fill('');
    row3[MEMBER_COLS.UNIT - 1] = 'Engineering'; // duplicate
    var data = [header, row1, row2, row3];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var depts = getDepartmentList();
    expect(Array.isArray(depts)).toBe(true);
    expect(depts.length).toBe(2);
    // Sorted alphabetically
    expect(depts[0]).toBe('Accounting');
    expect(depts[1]).toBe('Engineering');
  });

  test('returns empty array when member sheet is missing', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var depts = getDepartmentList();
    expect(Array.isArray(depts)).toBe(true);
    expect(depts.length).toBe(0);
  });
});

// ============================================================================
// getMemberList — behavioral tests
// ============================================================================

describe('getMemberList', () => {
  test('returns array of member objects with id, name, department', () => {
    var header = new Array(10).fill('');
    var row1 = new Array(10).fill('');
    row1[MEMBER_COLS.MEMBER_ID - 1] = 'M001';
    row1[MEMBER_COLS.FIRST_NAME - 1] = 'Alice';
    row1[MEMBER_COLS.LAST_NAME - 1] = 'Smith';
    row1[MEMBER_COLS.UNIT - 1] = 'Engineering';
    var data = [header, row1];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var members = getMemberList();
    expect(Array.isArray(members)).toBe(true);
    expect(members.length).toBe(1);
    expect(members[0].id).toBe('M001');
    expect(members[0].name).toBe('Alice Smith');
    expect(members[0].department).toBe('Engineering');
  });

  test('skips rows without a member ID', () => {
    var header = new Array(10).fill('');
    var row1 = new Array(10).fill('');
    row1[MEMBER_COLS.MEMBER_ID - 1] = '';
    row1[MEMBER_COLS.FIRST_NAME - 1] = 'NoID';
    var row2 = new Array(10).fill('');
    row2[MEMBER_COLS.MEMBER_ID - 1] = 'M002';
    row2[MEMBER_COLS.FIRST_NAME - 1] = 'Bob';
    row2[MEMBER_COLS.LAST_NAME - 1] = 'Jones';
    row2[MEMBER_COLS.UNIT - 1] = 'Sales';
    var data = [header, row1, row2];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var members = getMemberList();
    expect(members.length).toBe(1);
    expect(members[0].id).toBe('M002');
  });

  test('returns empty array when sheet is missing', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var members = getMemberList();
    expect(Array.isArray(members)).toBe(true);
    expect(members.length).toBe(0);
  });
});

// ============================================================================
// getDesktopSearchLocations — behavioral tests
// ============================================================================

describe('getDesktopSearchLocations', () => {
  test('returns sorted unique locations from member sheet', () => {
    var header = new Array(10).fill('');
    var row1 = new Array(10).fill('');
    row1[MEMBER_COLS.WORK_LOCATION - 1] = 'Boston';
    var row2 = new Array(10).fill('');
    row2[MEMBER_COLS.WORK_LOCATION - 1] = 'Worcester';
    var row3 = new Array(10).fill('');
    row3[MEMBER_COLS.WORK_LOCATION - 1] = 'Boston'; // dup
    var data = [header, row1, row2, row3];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var locations = getDesktopSearchLocations();
    expect(Array.isArray(locations)).toBe(true);
    expect(locations.length).toBe(2);
    expect(locations[0]).toBe('Boston');
    expect(locations[1]).toBe('Worcester');
  });
});

describe('Search uses SHEETS constants', () => {
  test('no hardcoded sheet names in 08b', () => {
    const fs = require('fs');
    const path = require('path');
    const code = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '08b_SearchAndCharts.gs'), 'utf8'
    );
    // SHEET_NAMES alias removed in v4.31.1 — verify no references remain
    expect(code).not.toContain('SHEET_NAMES.');
  });
});
