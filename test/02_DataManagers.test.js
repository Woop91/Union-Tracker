/**
 * Tests for 02_DataManagers.gs
 *
 * Covers: validateGrievanceData, addBusinessDays, getDaysUntilDeadline,
 * calculateInitialDeadlines, calculateNextStepDeadline, calculateResponseDeadline,
 * findExistingMember, getGrievanceStats, addMember, getMemberById, searchMembers.
 */

const { createMockRange, createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs'
]);

// ============================================================================
// validateGrievanceData
// ============================================================================

describe('validateGrievanceData', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns valid when memberName and required fields are provided', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'This is a valid description for a grievance'
    });
    expect(result.valid).toBe(true);
    expect(result.error).toBeUndefined();
  });

  test('returns valid when memberId is provided instead of memberName', () => {
    const result = validateGrievanceData({
      memberId: 'MS-101-H',
      grievanceType: 'Scheduling',
      description: 'Description that is long enough'
    });
    expect(result.valid).toBe(true);
  });

  test('returns invalid when neither memberName nor memberId is provided', () => {
    const result = validateGrievanceData({
      grievanceType: 'Discipline',
      description: 'This is a valid description'
    });
    expect(result.valid).toBe(false);
    expect(result.error).toMatch(/member/i);
  });

  test('returns invalid when grievanceType is missing', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      description: 'This is a valid description'
    });
    expect(result.valid).toBe(false);
    expect(result.error).toMatch(/grievance type/i);
  });

  test('returns invalid when description is missing', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      grievanceType: 'Discipline'
    });
    expect(result.valid).toBe(false);
    expect(result.error).toMatch(/description/i);
  });

  test('returns invalid when description is too short', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'Short'
    });
    expect(result.valid).toBe(false);
    expect(result.error).toMatch(/10 characters/i);
  });

  test('returns invalid when description is only whitespace under 10 chars', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: '   abc   '
    });
    expect(result.valid).toBe(false);
  });

  test('returns valid when description is exactly 10 characters', () => {
    const result = validateGrievanceData({
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: '1234567890'
    });
    expect(result.valid).toBe(true);
  });
});

// ============================================================================
// addBusinessDays
// ============================================================================

describe('addBusinessDays', () => {
  beforeEach(() => jest.clearAllMocks());

  test('adds business days skipping weekends', () => {
    // Monday 2025-01-06
    const monday = new Date(2025, 0, 6);
    const result = addBusinessDays(monday, 5);
    // 5 business days from Monday = next Monday (2025-01-13)
    expect(result.getDay()).toBe(1); // Monday
    expect(result.getDate()).toBe(13);
  });

  test('adding 1 business day from Friday gives Monday', () => {
    // Friday 2025-01-10
    const friday = new Date(2025, 0, 10);
    const result = addBusinessDays(friday, 1);
    expect(result.getDay()).toBe(1); // Monday
    expect(result.getDate()).toBe(13);
  });

  test('adding 1 business day from Thursday gives Friday', () => {
    // Thursday 2025-01-09
    const thursday = new Date(2025, 0, 9);
    const result = addBusinessDays(thursday, 1);
    expect(result.getDay()).toBe(5); // Friday
    expect(result.getDate()).toBe(10);
  });

  test('adding 0 business days returns the same date', () => {
    const monday = new Date(2025, 0, 6);
    const result = addBusinessDays(monday, 0);
    expect(result.getDate()).toBe(6);
  });

  test('adding 10 business days spans two weeks', () => {
    // Monday 2025-01-06
    const monday = new Date(2025, 0, 6);
    const result = addBusinessDays(monday, 10);
    // 10 business days = 2 weeks = Monday 2025-01-20
    expect(result.getDate()).toBe(20);
    expect(result.getDay()).toBe(1);
  });

  test('adding business days from Saturday skips to next week', () => {
    // Saturday 2025-01-11
    const saturday = new Date(2025, 0, 11);
    const result = addBusinessDays(saturday, 1);
    // Should land on Monday 2025-01-13
    expect(result.getDay()).toBe(1);
    expect(result.getDate()).toBe(13);
  });

  test('adding business days from Sunday skips to next week', () => {
    // Sunday 2025-01-12
    const sunday = new Date(2025, 0, 12);
    const result = addBusinessDays(sunday, 1);
    // Should land on Monday 2025-01-13
    expect(result.getDay()).toBe(1);
    expect(result.getDate()).toBe(13);
  });

  test('does not modify the original date', () => {
    const original = new Date(2025, 0, 6);
    const originalTime = original.getTime();
    addBusinessDays(original, 5);
    expect(original.getTime()).toBe(originalTime);
  });
});

// ============================================================================
// getDaysUntilDeadline
// ============================================================================

describe('getDaysUntilDeadline', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns positive days for future deadline', () => {
    const future = new Date();
    future.setDate(future.getDate() + 10);
    const result = getDaysUntilDeadline(future);
    expect(result).toBe(10);
  });

  test('returns negative days for past deadline', () => {
    const past = new Date();
    past.setDate(past.getDate() - 5);
    const result = getDaysUntilDeadline(past);
    expect(result).toBe(-5);
  });

  test('returns 0 for today', () => {
    const today = new Date();
    today.setHours(12, 0, 0, 0);
    const result = getDaysUntilDeadline(today);
    expect(result).toBe(0);
  });

  test('returns 1 for tomorrow', () => {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    tomorrow.setHours(12, 0, 0, 0);
    const result = getDaysUntilDeadline(tomorrow);
    expect(result).toBe(1);
  });
});

// ============================================================================
// calculateInitialDeadlines
// ============================================================================

describe('calculateInitialDeadlines', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns step1Due using DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE', () => {
    const filingDate = new Date(2025, 0, 6); // Monday
    const result = calculateInitialDeadlines(filingDate);
    expect(result).toHaveProperty('step1Due');
    // 7 business days from Jan 6 = Jan 15 (Wednesday)
    const expected = addBusinessDays(filingDate, DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);
    expect(result.step1Due.getTime()).toBe(expected.getTime());
  });

  test('step1Due is a Date object', () => {
    const filingDate = new Date(2025, 5, 1);
    const result = calculateInitialDeadlines(filingDate);
    expect(result.step1Due).toBeInstanceOf(Date);
  });

  test('step1Due is in the future relative to filing date', () => {
    const filingDate = new Date(2025, 0, 6);
    const result = calculateInitialDeadlines(filingDate);
    expect(result.step1Due.getTime()).toBeGreaterThan(filingDate.getTime());
  });
});

// ============================================================================
// calculateNextStepDeadline
// ============================================================================

describe('calculateNextStepDeadline', () => {
  beforeEach(() => jest.clearAllMocks());

  test('step 1 uses STEP_2.DAYS_TO_APPEAL', () => {
    const stepDate = new Date(2025, 0, 6); // Monday
    const result = calculateNextStepDeadline(1, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.STEP_2.DAYS_TO_APPEAL);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('step 2 uses STEP_3.DAYS_TO_APPEAL', () => {
    const stepDate = new Date(2025, 0, 6);
    const result = calculateNextStepDeadline(2, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.STEP_3.DAYS_TO_APPEAL);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('step 3 uses ARBITRATION.DAYS_TO_DEMAND', () => {
    const stepDate = new Date(2025, 0, 6);
    const result = calculateNextStepDeadline(3, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.ARBITRATION.DAYS_TO_DEMAND);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('returns null for invalid step (0)', () => {
    const result = calculateNextStepDeadline(0, new Date());
    expect(result).toBeNull();
  });

  test('returns null for invalid step (4)', () => {
    const result = calculateNextStepDeadline(4, new Date());
    expect(result).toBeNull();
  });

  test('returns a Date for valid steps', () => {
    const result = calculateNextStepDeadline(1, new Date(2025, 0, 6));
    expect(result).toBeInstanceOf(Date);
  });
});

// ============================================================================
// calculateResponseDeadline
// ============================================================================

describe('calculateResponseDeadline', () => {
  beforeEach(() => jest.clearAllMocks());

  test('step 1 uses STEP_1.DAYS_FOR_RESPONSE (7 business days)', () => {
    const stepDate = new Date(2025, 0, 6);
    const result = calculateResponseDeadline(1, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.STEP_1.DAYS_FOR_RESPONSE);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('step 2 uses STEP_2.DAYS_FOR_RESPONSE (14 business days)', () => {
    const stepDate = new Date(2025, 0, 6);
    const result = calculateResponseDeadline(2, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.STEP_2.DAYS_FOR_RESPONSE);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('step 3 uses STEP_3.DAYS_FOR_RESPONSE (21 business days)', () => {
    const stepDate = new Date(2025, 0, 6);
    const result = calculateResponseDeadline(3, stepDate);
    const expected = addBusinessDays(stepDate, DEADLINE_RULES.STEP_3.DAYS_FOR_RESPONSE);
    expect(result.getTime()).toBe(expected.getTime());
  });

  test('returns null for invalid step (0)', () => {
    expect(calculateResponseDeadline(0, new Date())).toBeNull();
  });

  test('returns null for invalid step (4)', () => {
    expect(calculateResponseDeadline(4, new Date())).toBeNull();
  });
});

// ============================================================================
// findExistingMember
// ============================================================================

describe('findExistingMember', () => {
  beforeEach(() => jest.clearAllMocks());

  // Build a mock data array matching MEMBER_COLS layout
  // MEMBER_ID=1(idx0), FIRST_NAME=2(idx1), LAST_NAME=3(idx2), ... CUBICLE=7(idx6), OFFICE_DAYS=8(idx7), EMAIL=9(idx8)
  const headers = ['Member ID', 'First Name', 'Last Name', 'Job Title', 'Work Location', 'Unit', 'Cubicle', 'Office Days', 'Email'];
  const memberData = [
    headers,
    ['MS-101-H', 'John', 'Doe', 'Analyst', 'Main', 'Unit A', '', 'Mon', 'john.doe@example.com'],
    ['MS-102-H', 'Jane', 'Smith', 'Manager', 'Field', 'Unit B', '', 'Tue', 'jane.smith@example.com'],
    ['MS-103-H', 'Bob', 'Wilson', 'Tech', 'Remote', 'Unit A', '', 'Wed', 'bob.wilson@example.com']
  ];

  test('matches by Member ID with HIGH confidence', () => {
    const result = findExistingMember({ memberId: 'MS-101-H' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('MEMBER_ID');
    expect(result.confidence).toBe('HIGH');
    expect(result.row).toBe(2); // 1-indexed (row 2 in sheet)
  });

  test('matches by email with HIGH confidence', () => {
    const result = findExistingMember({ email: 'jane.smith@example.com' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('EMAIL');
    expect(result.confidence).toBe('HIGH');
    expect(result.row).toBe(3);
  });

  test('matches by name with MEDIUM confidence', () => {
    const result = findExistingMember({ firstName: 'Bob', lastName: 'Wilson' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('NAME');
    expect(result.confidence).toBe('MEDIUM');
    expect(result.row).toBe(4);
  });

  test('Member ID match takes priority over email match', () => {
    const result = findExistingMember({
      memberId: 'MS-101-H',
      email: 'jane.smith@example.com'
    }, memberData);
    expect(result.matchType).toBe('MEMBER_ID');
    expect(result.row).toBe(2);
  });

  test('email match takes priority over name match', () => {
    // Search with email of Jane but name of Bob
    const result = findExistingMember({
      email: 'jane.smith@example.com',
      firstName: 'Bob',
      lastName: 'Wilson'
    }, memberData);
    expect(result.matchType).toBe('EMAIL');
    expect(result.row).toBe(3);
  });

  test('returns null when no match is found', () => {
    const result = findExistingMember({
      memberId: 'NONEXISTENT',
      email: 'nobody@example.com',
      firstName: 'Nobody',
      lastName: 'Here'
    }, memberData);
    expect(result).toBeNull();
  });

  test('handles case-insensitive email matching', () => {
    const result = findExistingMember({ email: 'JOHN.DOE@EXAMPLE.COM' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('EMAIL');
  });

  test('handles case-insensitive name matching', () => {
    const result = findExistingMember({ firstName: 'JANE', lastName: 'SMITH' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('NAME');
  });

  test('returns null for empty search params', () => {
    const result = findExistingMember({}, memberData);
    expect(result).toBeNull();
  });

  test('trims whitespace from search values', () => {
    const result = findExistingMember({ memberId: '  MS-101-H  ' }, memberData);
    expect(result).not.toBeNull();
    expect(result.matchType).toBe('MEMBER_ID');
  });
});

// ============================================================================
// Mock-dependent functions: getGrievanceStats, addMember, getMemberById, searchMembers
// ============================================================================

describe('getGrievanceStats (mock spreadsheet)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns zeroed stats when sheet is missing', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const stats = getGrievanceStats();
    expect(stats.total).toBe(0);
    expect(stats.open).toBe(0);
  });

  test('returns zeroed stats when sheet has only header row', () => {
    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, [
      ['Grievance ID', 'Member ID', 'First Name', 'Last Name', 'Status']
    ]);
    sheet.getLastRow.mockReturnValue(1);
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const stats = getGrievanceStats();
    expect(stats.total).toBe(0);
    expect(stats.open).toBe(0);
  });

  test('counts open, pending, won, and resolved grievances correctly', () => {
    // Build data using GRIEVANCE_COLS (1-indexed, subtract 1 for array access)
    const headerRow = new Array(30).fill('');
    headerRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'Grievance ID';
    headerRow[GRIEVANCE_COLS.STATUS - 1] = 'Status';
    headerRow[GRIEVANCE_COLS.LAST_UPDATED - 1] = 'Last Updated';
    headerRow[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = 'Category';

    const makeRow = (status, category) => {
      const row = new Array(30).fill('');
      row[GRIEVANCE_COLS.STATUS - 1] = status;
      row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = category || 'General';
      row[GRIEVANCE_COLS.LAST_UPDATED - 1] = new Date();
      return row;
    };

    const data = [
      headerRow,
      makeRow(GRIEVANCE_STATUS.OPEN, 'Discipline'),
      makeRow(GRIEVANCE_STATUS.OPEN, 'Discipline'),
      makeRow(GRIEVANCE_STATUS.PENDING, 'Scheduling'),
      makeRow(GRIEVANCE_STATUS.WON, 'Safety'),
      makeRow(GRIEVANCE_STATUS.CLOSED, 'Other')
    ];

    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, data);
    sheet.getLastRow.mockReturnValue(data.length);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 30)
    });
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const stats = getGrievanceStats();
    expect(stats.total).toBe(5);
    expect(stats.open).toBe(2);
    expect(stats.pending).toBe(1);
    expect(stats.won).toBe(1);
  });

  test('categoryData includes header row and category entries', () => {
    const headerRow = new Array(30).fill('');
    headerRow[GRIEVANCE_COLS.STATUS - 1] = 'Status';
    headerRow[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = 'Category';

    const makeRow = (status, category) => {
      const row = new Array(30).fill('');
      row[GRIEVANCE_COLS.STATUS - 1] = status;
      row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = category;
      row[GRIEVANCE_COLS.LAST_UPDATED - 1] = new Date();
      return row;
    };

    const data = [
      headerRow,
      makeRow(GRIEVANCE_STATUS.OPEN, 'Discipline'),
      makeRow(GRIEVANCE_STATUS.OPEN, 'Discipline'),
      makeRow(GRIEVANCE_STATUS.OPEN, 'Safety')
    ];

    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, data);
    sheet.getLastRow.mockReturnValue(data.length);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 30)
    });
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const stats = getGrievanceStats();
    expect(stats.categoryData[0]).toEqual(['Category', 'Count']);
    expect(stats.categoryData.length).toBeGreaterThan(1);
  });
});

describe('addMember (mock spreadsheet)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('appends member data and returns member ID', () => {
    const memberDirData = [
      ['Member ID', 'First Name', 'Last Name', 'Job Title', 'Work Location', 'Unit', 'Office Days', 'Email'],
      ['MS-101-H', 'Existing', 'Member', '', '', '', '', '']
    ];

    const mockRange = {
      setValue: jest.fn(),
      setValues: jest.fn(),
      getValue: jest.fn(),
      getValues: jest.fn(() => [['']]),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 1),
      getSheet: jest.fn(),
      getNumColumns: jest.fn(() => 1),
      getNumRows: jest.fn(() => 1),
      getA1Notation: jest.fn(() => 'A2')
    };

    const sheet = createMockSheet(SHEETS.MEMBER_DIR, memberDirData);
    sheet.getLastRow.mockReturnValue(2);
    sheet.getLastColumn.mockReturnValue(8);
    sheet.getRange.mockReturnValue(mockRange);

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const memberId = addMember({
      memberId: 'MS-200-H',
      firstName: 'New',
      lastName: 'Member',
      email: 'new@example.com'
    });

    expect(memberId).toBe('MS-200-H');
    expect(sheet.getRange).toHaveBeenCalled();
    expect(mockRange.setValues).toHaveBeenCalled();
  });

  test('throws error when Member Directory sheet is not found', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    expect(() => addMember({ firstName: 'Test', lastName: 'User' }))
      .toThrow(/Member Directory sheet not found/);
  });
});

describe('getMemberById (mock spreadsheet)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns member object when found', () => {
    const memberData = [
      ['Member ID', 'First Name', 'Last Name', 'Email'],
      ['MS-101-H', 'John', 'Doe', 'john@example.com'],
      ['MS-102-H', 'Jane', 'Smith', 'jane@example.com']
    ];

    const sheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => memberData),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => memberData.length),
      getNumColumns: jest.fn(() => 4)
    });

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const member = getMemberById('MS-101-H');
    expect(member).not.toBeNull();
    expect(member['First Name']).toBe('John');
    expect(member['Last Name']).toBe('Doe');
    expect(member['Email']).toBe('john@example.com');
  });

  test('returns null when member is not found', () => {
    const memberData = [
      ['Member ID', 'First Name', 'Last Name', 'Email'],
      ['MS-101-H', 'John', 'Doe', 'john@example.com']
    ];

    const sheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => memberData),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => memberData.length),
      getNumColumns: jest.fn(() => 4)
    });

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const member = getMemberById('NONEXISTENT');
    expect(member).toBeNull();
  });

  test('returns null when sheet does not exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const member = getMemberById('MS-101-H');
    expect(member).toBeNull();
  });
});

describe('searchMembers (mock spreadsheet)', () => {
  beforeEach(() => jest.clearAllMocks());

  // Build data with proper column indices for MEMBER_COLS
  // MEMBER_ID=1(idx0), FIRST_NAME=2(idx1), LAST_NAME=3(idx2), ... EMAIL=8(idx7)
  const headers = new Array(10).fill('');
  headers[MEMBER_COLS.MEMBER_ID - 1] = 'Member ID';
  headers[MEMBER_COLS.FIRST_NAME - 1] = 'First Name';
  headers[MEMBER_COLS.LAST_NAME - 1] = 'Last Name';
  headers[MEMBER_COLS.EMAIL - 1] = 'Email';

  const makeRow = (id, first, last, email) => {
    const row = new Array(10).fill('');
    row[MEMBER_COLS.MEMBER_ID - 1] = id;
    row[MEMBER_COLS.FIRST_NAME - 1] = first;
    row[MEMBER_COLS.LAST_NAME - 1] = last;
    row[MEMBER_COLS.EMAIL - 1] = email;
    return row;
  };

  const memberData = [
    headers,
    makeRow('MS-101-H', 'John', 'Doe', 'john@example.com'),
    makeRow('MS-102-H', 'Jane', 'Smith', 'jane@example.com'),
    makeRow('MS-103-H', 'John', 'Adams', 'jadams@example.com')
  ];

  function setupSearchMock() {
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => memberData),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => memberData.length),
      getNumColumns: jest.fn(() => 10)
    });
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);
  }

  test('finds members by first name', () => {
    setupSearchMock();
    const results = searchMembers('John');
    expect(results.length).toBe(2);
  });

  test('finds members by last name', () => {
    setupSearchMock();
    const results = searchMembers('Smith');
    expect(results.length).toBe(1);
    expect(results[0]['First Name']).toBe('Jane');
  });

  test('finds members by email substring', () => {
    setupSearchMock();
    const results = searchMembers('jadams');
    expect(results.length).toBe(1);
  });

  test('finds members by member ID', () => {
    setupSearchMock();
    const results = searchMembers('MS-102');
    expect(results.length).toBe(1);
  });

  test('search is case-insensitive', () => {
    setupSearchMock();
    const results = searchMembers('jane');
    expect(results.length).toBe(1);
  });

  test('returns empty array when no match', () => {
    setupSearchMock();
    const results = searchMembers('Nonexistent');
    expect(results.length).toBe(0);
  });

  test('returns empty array when sheet is missing', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const results = searchMembers('John');
    expect(results).toEqual([]);
  });
});

// ============================================================================
// startNewGrievance (integration - exercises GRIEVANCE_OUTCOMES)
// ============================================================================

describe('startNewGrievance', () => {
  beforeEach(() => jest.clearAllMocks());

  function setupGrievanceMock() {
    const headerRow = new Array(40).fill('');
    headerRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'Grievance ID';

    const existingRow = new Array(40).fill('');
    existingRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'GRV-2026-0001';

    const data = [headerRow, existingRow];

    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, data);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 40)
    });
    sheet.getLastRow.mockReturnValue(data.length);

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    // Mock logAuditEvent if not already defined
    if (!global.logAuditEvent) {
      global.logAuditEvent = jest.fn();
    } else {
      global.logAuditEvent.mockClear();
    }

    return { sheet, mockSS };
  }

  test('returns success with grievance ID for valid input', () => {
    const { sheet } = setupGrievanceMock();

    const result = startNewGrievance({
      memberId: 'MS-101-H',
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'Valid description for the grievance'
    });

    expect(result.success).toBe(true);
    expect(result.grievanceId).toBeDefined();
    expect(result.grievanceId).toMatch(/^GRV-/);
    expect(sheet.appendRow).toHaveBeenCalledTimes(1);
  });

  test('row data uses GRIEVANCE_COLS positions and includes correct status', () => {
    const { sheet } = setupGrievanceMock();

    startNewGrievance({
      memberId: 'MS-101-H',
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'Valid description for the grievance'
    });

    const appendedRow = sheet.appendRow.mock.calls[0][0];
    // Status should be set at the correct GRIEVANCE_COLS position
    expect(appendedRow[GRIEVANCE_COLS.STATUS - 1]).toBe(GRIEVANCE_STATUS.OPEN);
    expect(appendedRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1]).toMatch(/^GRV-/);
    expect(appendedRow[GRIEVANCE_COLS.MEMBER_ID - 1]).toBe('MS-101-H');
    expect(appendedRow[GRIEVANCE_COLS.ISSUE_CATEGORY - 1]).toBe('Discipline');
    expect(appendedRow[GRIEVANCE_COLS.CURRENT_STEP - 1]).toBe(1);
  });

  test('returns error for invalid grievance data', () => {
    setupGrievanceMock();

    const result = startNewGrievance({
      // missing memberName/memberId
      grievanceType: 'Discipline',
      description: 'Valid description'
    });

    expect(result.success).toBe(false);
    expect(result.error).toBeDefined();
  });

  test('returns error when grievance sheet is missing', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = startNewGrievance({
      memberId: 'MS-101-H',
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'Valid description for the grievance'
    });

    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('calls logAuditEvent after creation', () => {
    setupGrievanceMock();

    startNewGrievance({
      memberId: 'MS-101-H',
      memberName: 'John Doe',
      grievanceType: 'Discipline',
      description: 'Valid description for the grievance'
    });

    expect(logAuditEvent).toHaveBeenCalledWith(
      AUDIT_EVENTS.GRIEVANCE_CREATED,
      expect.objectContaining({
        grievanceId: expect.stringMatching(/^GRV-/),
        memberId: 'MS-101-H'
      })
    );
  });
});

// ============================================================================
// resolveGrievance (integration)
// ============================================================================

describe('resolveGrievance', () => {
  beforeEach(() => jest.clearAllMocks());

  function setupResolveMock(grievanceId, currentStep) {
    const headerRow = new Array(40).fill('');
    headerRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'Grievance ID';
    headerRow[GRIEVANCE_COLS.STATUS - 1] = 'Status';

    const grievanceRow = new Array(40).fill('');
    grievanceRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = grievanceId || 'GRV-2026-0001';
    grievanceRow[GRIEVANCE_COLS.STATUS - 1] = GRIEVANCE_STATUS.OPEN;
    grievanceRow[GRIEVANCE_COLS.CURRENT_STEP - 1] = currentStep || 1;
    grievanceRow[GRIEVANCE_COLS.RESOLUTION - 1] = '';

    const data = [headerRow, grievanceRow];

    const setValueCalls = {};
    const mockRange = {
      setValue: jest.fn(function(val) { return val; }),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 40),
      getA1Notation: jest.fn(() => 'A2'),
      getSheet: jest.fn()
    };

    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, data);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 40)
    });
    sheet.getRange.mockReturnValue(mockRange);

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    if (!global.logAuditEvent) {
      global.logAuditEvent = jest.fn();
    } else {
      global.logAuditEvent.mockClear();
    }

    return { sheet, mockRange };
  }

  test('returns success for valid resolution', () => {
    setupResolveMock('GRV-2026-0001');

    const result = resolveGrievance('GRV-2026-0001', 'Won', 'Management conceded', 'Great outcome');

    expect(result.success).toBe(true);
    expect(result.grievanceId).toBe('GRV-2026-0001');
    expect(result.message).toContain('Won');
  });

  test('updates resolution, outcome, and status on the sheet', () => {
    const { mockRange } = setupResolveMock('GRV-2026-0001');

    resolveGrievance('GRV-2026-0001', 'Settled', 'Partial agreement reached', '');

    const setValueArgs = mockRange.setValue.mock.calls.map(c => c[0]);
    // Outcome and resolution are combined into one RESOLUTION column write
    const resolutionArg = setValueArgs.find(v => typeof v === 'string' && v.includes('Settled'));
    expect(resolutionArg).toContain('Settled');
    expect(resolutionArg).toContain('Partial agreement reached');
    expect(setValueArgs).toContain(GRIEVANCE_STATUS.RESOLVED);   // status
  });

  test('returns error when grievance ID not found', () => {
    setupResolveMock('GRV-2026-0001');

    const result = resolveGrievance('GRV-NONEXISTENT', 'Won', 'N/A', '');

    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('logs audit event on resolution', () => {
    setupResolveMock('GRV-2026-0001');

    resolveGrievance('GRV-2026-0001', 'Denied', 'No violation found', 'Reviewed thoroughly');

    expect(logAuditEvent).toHaveBeenCalledWith(
      AUDIT_EVENTS.GRIEVANCE_UPDATED,
      expect.objectContaining({
        grievanceId: 'GRV-2026-0001',
        action: 'RESOLVED',
        outcome: 'Denied'
      })
    );
  });

  test('appends timestamped notes when notes provided', () => {
    const { mockRange } = setupResolveMock('GRV-2026-0001');

    resolveGrievance('GRV-2026-0001', 'Won', 'Full remedy', 'Victory!');

    const setValueArgs = mockRange.setValue.mock.calls.map(c => c[0]);
    // Outcome, resolution, and notes are combined into one RESOLUTION column write
    const resolutionArg = setValueArgs.find(v => typeof v === 'string' && v.includes('Won'));
    expect(resolutionArg).toContain('Won');
    expect(resolutionArg).toContain('Victory!');
  });
});

// ============================================================================
// advanceGrievanceStep (integration)
// ============================================================================

describe('advanceGrievanceStep', () => {
  beforeEach(() => jest.clearAllMocks());

  function setupAdvanceMock(grievanceId, currentStep) {
    const headerRow = new Array(40).fill('');
    headerRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = 'Grievance ID';

    const grievanceRow = new Array(40).fill('');
    grievanceRow[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = grievanceId || 'GRV-2026-0001';
    grievanceRow[GRIEVANCE_COLS.CURRENT_STEP - 1] = currentStep || 1;
    grievanceRow[GRIEVANCE_COLS.STATUS - 1] = GRIEVANCE_STATUS.OPEN;
    grievanceRow[GRIEVANCE_COLS.RESOLUTION - 1] = '';

    const data = [headerRow, grievanceRow];

    const mockRange = {
      setValue: jest.fn(),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 40),
      getA1Notation: jest.fn(() => 'A2'),
      getSheet: jest.fn()
    };

    const sheet = createMockSheet(SHEET_NAMES.GRIEVANCE_TRACKER, data);
    sheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => data),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => data.length),
      getNumColumns: jest.fn(() => 40)
    });
    sheet.getRange.mockReturnValue(mockRange);

    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    if (!global.logAuditEvent) {
      global.logAuditEvent = jest.fn();
    } else {
      global.logAuditEvent.mockClear();
    }

    return { sheet, mockRange };
  }

  test('advances from step 1 to step 2', () => {
    setupAdvanceMock('GRV-2026-0001', 1);

    const result = advanceGrievanceStep('GRV-2026-0001', {});

    expect(result.success).toBe(true);
    expect(result.newStep).toBe(2);
    expect(result.message).toContain('Step 2');
  });

  test('advances from step 2 to step 3', () => {
    setupAdvanceMock('GRV-2026-0001', 2);

    const result = advanceGrievanceStep('GRV-2026-0001', {});

    expect(result.success).toBe(true);
    expect(result.newStep).toBe(3);
  });

  test('advances from step 3 to arbitration (step 4)', () => {
    setupAdvanceMock('GRV-2026-0001', 3);

    const result = advanceGrievanceStep('GRV-2026-0001', {});

    expect(result.success).toBe(true);
    expect(result.newStep).toBe(4);
  });

  test('returns error when already at arbitration (step 4)', () => {
    setupAdvanceMock('GRV-2026-0001', 4);

    const result = advanceGrievanceStep('GRV-2026-0001', {});

    expect(result.success).toBe(false);
    expect(result.error).toMatch(/arbitration/i);
  });

  test('returns error when grievance not found', () => {
    setupAdvanceMock('GRV-2026-0001', 1);

    const result = advanceGrievanceStep('GRV-NONEXISTENT', {});

    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('updates status to Appealed for steps 1-2', () => {
    const { mockRange } = setupAdvanceMock('GRV-2026-0001', 1);

    advanceGrievanceStep('GRV-2026-0001', {});

    const setValueArgs = mockRange.setValue.mock.calls.map(c => c[0]);
    expect(setValueArgs).toContain(GRIEVANCE_STATUS.APPEALED);
  });

  test('updates status to In Arbitration for step 3 advance', () => {
    const { mockRange } = setupAdvanceMock('GRV-2026-0001', 3);

    advanceGrievanceStep('GRV-2026-0001', {});

    const setValueArgs = mockRange.setValue.mock.calls.map(c => c[0]);
    expect(setValueArgs).toContain(GRIEVANCE_STATUS.AT_ARBITRATION);
  });

  test('logs audit event with step details', () => {
    setupAdvanceMock('GRV-2026-0001', 2);

    advanceGrievanceStep('GRV-2026-0001', { notes: 'Denied at step 2' });

    expect(logAuditEvent).toHaveBeenCalledWith(
      AUDIT_EVENTS.GRIEVANCE_STEP_ADVANCED,
      expect.objectContaining({
        grievanceId: 'GRV-2026-0001',
        fromStep: 2,
        toStep: 3
      })
    );
  });
});
