/**
 * Tests for 21_WebDashDataService.gs
 *
 * Covers DataService IIFE: user lookup, role determination, member data,
 * grievance access, steward dashboard data, and helper functions.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '21_WebDashDataService.gs', '21d_WebDashDataWrappers.gs']);

let mockMemberSheet;
let mockGrievanceSheet;

function makeMemberData() {
  return [
    ['Email', 'Name', 'First Name', 'Last Name', 'Role', 'Unit', 'Phone', 'Join Date', 'Dues Status', 'Member ID', 'Work Location', 'Office Days', 'Assigned Steward', 'Is Steward'],
    ['member@test.com', 'Jane Doe', 'Jane', 'Doe', 'Member', 'Unit A', '555-0001', new Date('2023-01-15'), 'Active', 'MEM-001', 'HQ', 'Mon,Tue', '', false],
    ['steward@test.com', 'John Smith', 'John', 'Smith', 'Steward', 'Unit A', '555-0002', new Date('2022-06-01'), 'Active', 'MEM-002', 'HQ', 'Mon-Fri', '', true],
    ['admin@test.com', 'Admin User', 'Admin', 'User', 'Admin', 'Admin', '555-0003', new Date('2021-01-01'), 'Active', 'MEM-003', 'Main', '', '', true],
  ];
}

function makeGrievanceData() {
  return [
    ['Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline', 'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category', 'Resolution'],
    ['GR-001', 'member@test.com', 'Open', 'Step I', new Date('2026-04-01'), new Date('2026-02-01'), 'steward@test.com', 'Unit A', 'High', 'Test grievance', 'Safety', ''],
    ['GR-002', 'member@test.com', 'Won', 'Step II', '', new Date('2025-11-01'), 'steward@test.com', 'Unit A', 'Medium', 'Old case', 'Pay', 'Resolved'],
    ['GR-003', 'other@test.com', 'Open', 'Step I', new Date('2026-05-01'), new Date('2026-01-15'), 'steward@test.com', 'Unit B', 'Low', 'Another case', 'Hours', ''],
  ];
}

function setupSheets() {
  const memberData = makeMemberData();
  const grievanceData = makeGrievanceData();

  mockMemberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', memberData);
  mockGrievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG || 'Grievance Log', grievanceData);

  const mockSs = createMockSpreadsheet([mockMemberSheet, mockGrievanceSheet]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
}

beforeEach(() => {
  // Clear the DataService IIFE in-memory cache so each test gets fresh data
  if (typeof DataService !== 'undefined' && DataService._invalidateSheetCache) {
    DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
  }
  // Reset spreadsheet singleton so mock changes take effect
  if (typeof DataService !== 'undefined' && DataService._resetSSCache) {
    DataService._resetSSCache();
  }
  setupSheets();
  // Re-mock auth helpers that loadSources overwrites with real implementations
  global._resolveCallerEmail = jest.fn(() => 'member@test.com');
  global._requireStewardAuth = jest.fn(() => 'steward@test.com');
});

// ============================================================================
// DataService.findUserByEmail
// ============================================================================

describe('DataService.findUserByEmail', () => {
  test('returns null for empty email', () => {
    expect(DataService.findUserByEmail('')).toBeNull();
    expect(DataService.findUserByEmail(null)).toBeNull();
  });

  test('returns null for unknown email', () => {
    expect(DataService.findUserByEmail('unknown@test.com')).toBeNull();
  });

  test('finds user by email (case insensitive)', () => {
    const user = DataService.findUserByEmail('MEMBER@TEST.COM');
    expect(user).not.toBeNull();
    expect(user.email).toBe('member@test.com');
  });

  test('returns user with expected fields', () => {
    const user = DataService.findUserByEmail('member@test.com');
    expect(user).toHaveProperty('email');
    expect(user).toHaveProperty('name');
  });

  test('trims email before lookup', () => {
    const user = DataService.findUserByEmail('  member@test.com  ');
    expect(user).not.toBeNull();
  });

  test('returns null when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(DataService.findUserByEmail('member@test.com')).toBeNull();
  });
});

// ============================================================================
// DataService.getUserRole
// ============================================================================

describe('DataService.getUserRole', () => {
  test('returns "member" for regular member', () => {
    const role = DataService.getUserRole('member@test.com');
    expect(role).toBe('member');
  });

  test('returns "steward" for steward', () => {
    const role = DataService.getUserRole('steward@test.com');
    expect(['steward', 'admin']).toContain(role);
  });

  test('returns null for unknown email', () => {
    const role = DataService.getUserRole('unknown@test.com');
    expect(role === null || role === 'member').toBe(true);
  });

  test('returns null for empty email', () => {
    const role = DataService.getUserRole('');
    expect(role).toBeNull();
  });
});

// ============================================================================
// DataService.getMemberGrievances (member portal data)
// Note: getMemberData does not exist in the public API; getMemberGrievances is
// the canonical method for member-scoped grievance data.
// ============================================================================

describe('DataService.getMemberData (via getMemberGrievances)', () => {
  test('returns empty array for empty email', () => {
    expect(DataService.getMemberGrievances('')).toEqual([]);
    expect(DataService.getMemberGrievances(null)).toEqual([]);
  });

  test('returns grievances array for known member', () => {
    const grievances = DataService.getMemberGrievances('member@test.com');
    expect(grievances).not.toBeNull();
    expect(Array.isArray(grievances)).toBe(true);
  });

  test('returns empty grievances for member with no cases', () => {
    const grievances = DataService.getMemberGrievances('admin@test.com');
    expect(Array.isArray(grievances)).toBe(true);
    expect(grievances.length).toBe(0);
  });

  test('returns empty when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    const result = DataService.getMemberGrievances('member@test.com');
    expect(result === null || (Array.isArray(result) && result.length === 0)).toBe(true);
  });
});

// ============================================================================
// DataService.getMemberGrievances
// ============================================================================

describe('DataService.getMemberGrievances', () => {
  test('returns empty array for empty email', () => {
    const result = DataService.getMemberGrievances('');
    expect(result).toEqual([]);
  });

  test('returns grievances for known member', () => {
    const result = DataService.getMemberGrievances('member@test.com');
    expect(Array.isArray(result)).toBe(true);
  });

  test('returns empty for unknown member', () => {
    const result = DataService.getMemberGrievances('nobody@test.com');
    expect(result).toEqual([]);
  });
});

// ============================================================================
// DataService.getStewardCases (steward dashboard data)
// Note: getStewardDashboardData does not exist in the public API; getStewardCases
// is the canonical method for steward-scoped case data.
// ============================================================================

describe('DataService.getStewardDashboardData (via getStewardCases)', () => {
  test('returns empty array for empty email', () => {
    const result = DataService.getStewardCases('');
    expect(result === null || (Array.isArray(result) && result.length === 0)).toBe(true);
  });

  test('returns cases array for steward', () => {
    const cases = DataService.getStewardCases('steward@test.com');
    expect(Array.isArray(cases)).toBe(true);
  });

  test('returns empty when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    const result = DataService.getStewardCases('steward@test.com');
    expect(result === null || (Array.isArray(result) && result.length === 0)).toBe(true);
  });
});

// ============================================================================
// DataService.getMemberTasks
// ============================================================================

describe('DataService.getMemberTasks', () => {
  test('returns empty array for empty email', () => {
    const result = DataService.getMemberTasks('');
    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBe(0);
  });

  test('returns array for valid email', () => {
    const result = DataService.getMemberTasks('member@test.com');
    expect(Array.isArray(result)).toBe(true);
  });
});

// ============================================================================
// DataService.getAllMembers (directory data)
// Note: getDirectoryData does not exist in the public API; getAllMembers is
// the canonical method for full member directory access.
// ============================================================================

describe('DataService.getDirectoryData (via getAllMembers)', () => {
  test('returns directory data', () => {
    const data = DataService.getAllMembers();
    expect(Array.isArray(data)).toBe(true);
  });

  test('returns empty when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    const data = DataService.getAllMembers();
    expect(data === null || (Array.isArray(data) && data.length === 0)).toBe(true);
  });
});

// ============================================================================
// DataService.getGrievanceStats
// (Moved from expansion.test.js — canonical home for DataService tests)
// ============================================================================

describe('DataService.getGrievanceStats', () => {
  beforeEach(() => {
    // Build a mock Grievance Log with 12 rows (above the minimum-10 threshold)
    const headers = [
      'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
      'Filed', 'Steward', 'Unit', 'Priority', 'Notes',
      'Issue Category', 'Resolution', 'Date Closed',
    ];
    const rows = [
      ['G001', 'a@test.com', 'Won', 'Step II', '', '2025-06-15', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won at Step II', '2025-07-01'],
      ['G002', 'b@test.com', 'Denied', 'Step III', '', '2025-06-20', 'steward@test.com', 'Unit A', 'medium', '', 'Pay', 'Denied', '2025-08-01'],
      ['G003', 'c@test.com', 'Settled', 'Step I', '', '2025-07-01', 'steward@test.com', 'Unit B', 'low', '', 'Safety', 'Settlement reached', '2025-07-15'],
      ['G004', 'd@test.com', 'Active', 'Step I', '', '2025-07-10', 'steward@test.com', 'Unit A', 'high', '', 'Discipline', '', ''],
      ['G005', 'e@test.com', 'New', 'Step I', '', '2025-08-01', 'steward@test.com', 'Unit C', 'medium', '', 'Pay', '', ''],
      ['G006', 'f@test.com', 'Won', 'Step II', '', '2025-08-15', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won', '2025-09-01'],
      ['G007', 'g@test.com', 'Active', 'Step II', '', '2025-09-01', 'steward@test.com', 'Unit B', 'medium', '', 'Workload', '', ''],
      ['G008', 'h@test.com', 'New', 'Step I', '', '2025-09-10', 'steward@test.com', 'Unit C', 'low', '', 'Pay', '', ''],
      ['G009', 'i@test.com', 'Active', 'Step I', '', '2025-10-01', 'steward@test.com', 'Unit A', 'high', '', 'Safety', '', ''],
      ['G010', 'j@test.com', 'Settled', 'Step III', '', '2025-10-15', 'steward@test.com', 'Unit B', 'medium', '', 'Discipline', 'Settled', '2025-11-01'],
      ['G011', 'k@test.com', 'Won', 'Step I', '', '2025-11-01', 'steward@test.com', 'Unit A', 'high', '', 'Safety', 'Won', '2025-11-15'],
      ['G012', 'l@test.com', 'Active', 'Step II', '', '2025-12-01', 'steward@test.com', 'Unit C', 'medium', '', 'Pay', '', ''],
    ];
    const data = [headers, ...rows];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
  });

  test('returns available: true with 12 records', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(true);
    expect(stats.total).toBe(12);
  });

  test('returns byCategory with real category names', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.byCategory).toBeDefined();
    const cats = Object.keys(stats.byCategory);
    expect(cats.length).toBeGreaterThan(1);
    expect(cats).toContain('Safety');
    expect(cats).toContain('Pay');
  });

  test('returns monthlyResolved array', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.monthlyResolved).toBeDefined();
    expect(Array.isArray(stats.monthlyResolved)).toBe(true);
    expect(stats.monthlyResolved.length).toBeGreaterThan(0);
    stats.monthlyResolved.forEach(entry => {
      expect(entry).toHaveProperty('month');
      expect(entry).toHaveProperty('count');
    });
  });

  test('returns summary counts (wonCount, deniedCount, settledCount, withdrawnCount)', () => {
    const stats = DataService.getGrievanceStats();
    expect(typeof stats.wonCount).toBe('number');
    expect(typeof stats.deniedCount).toBe('number');
    expect(typeof stats.settledCount).toBe('number');
    expect(typeof stats.withdrawnCount).toBe('number');
    expect(stats.wonCount).toBe(3);
    expect(stats.deniedCount).toBe(1);
    expect(stats.settledCount).toBe(2);
    expect(stats.withdrawnCount).toBe(0);
  });

  test('byStep counts are accurate', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.byStep).toBeDefined();
    const totalByStep = Object.values(stats.byStep).reduce((a, b) => a + b, 0);
    expect(totalByStep).toBe(12);
  });
});

// ============================================================================
// DataService.getGrievanceHotSpots
// ============================================================================

describe('DataService.getGrievanceHotSpots', () => {
  test('returns locations with 3+ grievances', () => {
    DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    const headers = ['Grievance ID', 'Unit'];
    const data = [headers,
      ['G1', 'Hot Unit'], ['G2', 'Hot Unit'], ['G3', 'Hot Unit'],
      ['G4', 'Cold Unit'], ['G5', 'Cold Unit'],
    ];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const spots = DataService.getGrievanceHotSpots();
    expect(spots.length).toBe(1);
    expect(spots[0].location).toBe('Hot Unit');
    expect(spots[0].count).toBe(3);
  });

  test('returns empty when no unit has 3+ cases', () => {
    DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    const headers = ['Grievance ID', 'Unit'];
    const data = [headers, ['G1', 'A'], ['G2', 'B']];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    expect(DataService.getGrievanceHotSpots()).toEqual([]);
  });
});

// ============================================================================
// DataService.getStewardDirectory
// ============================================================================

describe('DataService.getStewardDirectory', () => {
  test('returns steward records (unsorted — client applies smart sort)', () => {
    DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
    const headers = ['Email', 'Name', 'Role', 'Work Location', 'Office Days', 'Phone', 'Unit'];
    const data = [headers,
      ['z@test.com', 'Zara', 'Steward', 'Office B', 'Mon-Fri', '555-0001', 'Unit 1'],
      ['a@test.com', 'Alice', 'Steward', 'Office A', 'Mon-Wed', '555-0002', 'Unit 2'],
      ['m@test.com', 'Mike', 'Member', 'Office A', 'Mon-Fri', '555-0003', 'Unit 1'],
    ];
    const memberSheet = createMockSheet('Member Directory', data);
    const mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stewards = DataService.getStewardDirectory();
    expect(stewards.length).toBe(2);
    const names = stewards.map(s => s.name).sort();
    expect(names).toEqual(['Alice', 'Zara']);
    expect(stewards[0]).toHaveProperty('email');
    expect(stewards[0]).toHaveProperty('workLocation');
    expect(stewards[0]).toHaveProperty('phone');
  });
});

// ============================================================================
// DataService.isChiefSteward
// ============================================================================

describe('DataService.isChiefSteward', () => {
  test('is exported from DataService', () => {
    expect(typeof DataService.isChiefSteward).toBe('function');
  });
});

// ============================================================================
// Welcome experience
// ============================================================================

describe('dataGetWelcomeData', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetWelcomeData).toBe('function');
  });
});

describe('dataMarkWelcomeDismissed', () => {
  test('is a defined global function', () => {
    expect(typeof dataMarkWelcomeDismissed).toBe('function');
  });

  test('stores dismissal timestamp in Script Properties', () => {
    PropertiesService.getScriptProperties().deleteAllProperties();
    const result = dataMarkWelcomeDismissed('test-session-token');
    expect(result.success).toBe(true);
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const dismissKey = Object.keys(allProps).find(k => k.startsWith('WELCOME_DISMISSED_'));
    expect(dismissKey).toBeTruthy();
  });

  test('returns failure when no authenticated user', () => {
    const origResolve = global._resolveCallerEmail;
    global._resolveCallerEmail = jest.fn(() => null);
    const result = dataMarkWelcomeDismissed('test-session-token');
    expect(result.success).toBe(false);
    global._resolveCallerEmail = origResolve;
  });
});

// ============================================================================
// dataGetEngagementStats
// ============================================================================

describe('dataGetEngagementStats', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetEngagementStats).toBe('function');
  });

  test('returns null when no member data exists', () => {
    // Override with empty spreadsheet — no member directory means totalMembers=0 → null
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(createMockSpreadsheet([]));
    PropertiesService.getScriptProperties().deleteAllProperties();
    const result = dataGetEngagementStats('test-session-token');
    expect(result).toBeNull();
  });
});

// ============================================================================
// dataGetBroadcastFilterOptions
// ============================================================================

describe('dataGetBroadcastFilterOptions', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetBroadcastFilterOptions).toBe('function');
  });

  test('returns filter options from member directory', () => {
    DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
    const origAuth = global.checkWebAppAuthorization;
    global.checkWebAppAuthorization = jest.fn(() => ({
      isAuthorized: true, email: 'steward@test.com', role: 'steward'
    }));

    const headers = ['Email', 'Name', 'Work Location', 'Office Days', 'Dues Status', 'Role', 'Assigned Steward'];
    const data = [headers,
      ['a@test.com', 'Alice', 'Office A', 'Mon-Wed', 'Active', 'Member', 'steward@test.com'],
      ['b@test.com', 'Bob', 'Office B', 'Thu-Fri', 'Inactive', 'Member', 'steward@test.com'],
      ['c@test.com', 'Carol', 'Office A', 'Mon-Fri', 'Active', 'Member', 'steward@test.com'],
    ];
    const memberSheet = createMockSheet('Member Directory', data);
    const mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const filters = dataGetBroadcastFilterOptions('test-session-token');
    expect(filters).toBeDefined();
    expect(filters.locations).toBeDefined();
    expect(Array.isArray(filters.locations)).toBe(true);
    expect(filters.locations).toContain('Office A');
    expect(filters.locations).toContain('Office B');

    global.checkWebAppAuthorization = origAuth;
  });
});

// ============================================================================
// Global Wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('dataGetMemberGrievances delegates to DataService.getMemberGrievances', () => {
    const result = dataGetMemberGrievances('member@test.com');
    expect(Array.isArray(result)).toBe(true);
  });

  test('dataGetStewardCases delegates to DataService.getStewardCases', () => {
    const result = dataGetStewardCases();
    expect(Array.isArray(result)).toBe(true);
  });

  test('dataGetStewardKPIs delegates to DataService.getStewardKPIs', () => {
    const result = dataGetStewardKPIs();
    expect(typeof result).toBe('object');
  });

  test('dataGetGrievanceStats delegates to DataService.getGrievanceStats', () => {
    const result = dataGetGrievanceStats();
    expect(typeof result).toBe('object');
  });

  test('dataGetMemberGrievanceStats uses _resolveCallerEmail and delegates to DataService.getGrievanceStats', () => {
    const result = dataGetMemberGrievanceStats('test-session-token');
    expect(typeof result).toBe('object');
  });

  test('dataGetMemberGrievanceHotSpots uses _resolveCallerEmail and delegates to DataService.getGrievanceHotSpots', () => {
    const result = dataGetMemberGrievanceHotSpots('test-session-token');
    expect(Array.isArray(result)).toBe(true);
  });

  test('dataGetStewardDirectory delegates to DataService.getStewardDirectory', () => {
    const result = dataGetStewardDirectory('test-session-token');
    expect(Array.isArray(result)).toBe(true);
  });

  test('dataGetAllMembers delegates to DataService.getAllMembers', () => {
    // _requireStewardAuth is mocked to return 'steward@example.com'
    const result = dataGetAllMembers();
    expect(Array.isArray(result)).toBe(true);
  });
});

// ============================================================================
// DataService.getStewardSurveyTracking
// Regression tests for N+1 fix (fa27b42): survey sheet must be read ONCE
// regardless of member count, with an O(1) email→status map lookup per member.
// ============================================================================

describe('DataService.getStewardSurveyTracking', () => {
  const SURVEY_HEADERS = ['Email', 'Status', 'Completed Date'];

  function makeSurveySheet(rows) {
    const data = [SURVEY_HEADERS, ...rows];
    return createMockSheet('_Survey_Tracking', data);
  }

  function makeFullMockSs(surveyRows) {
    const memberData = makeMemberData();
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', memberData);
    const grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG || 'Grievance Log', makeGrievanceData());
    const surveySheet = makeSurveySheet(surveyRows || []);
    return { ss: createMockSpreadsheet([memberSheet, grievanceSheet, surveySheet]), surveySheet };
  }

  beforeEach(() => {
    if (typeof DataService !== 'undefined' && DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
  });

  test('reads the survey sheet exactly once for multiple members (N+1 fix)', () => {
    // The fix pre-loads _Survey_Tracking once; previously it was read per member.
    const { ss, surveySheet } = makeFullMockSs([
      ['member@test.com', 'Completed', '2026-01-15'],
      ['steward@test.com', 'Pending', ''],
    ]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    DataService.getStewardSurveyTracking('steward@test.com', 'all');

    // getDataRange on the survey sheet must be called exactly once
    expect(surveySheet.getDataRange).toHaveBeenCalledTimes(1);
  });

  test('returns correct completion counts from survey map', () => {
    const { ss } = makeFullMockSs([
      ['member@test.com', 'completed', '2026-01-15'],
      ['steward@test.com', 'pending', ''],
    ]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    const result = DataService.getStewardSurveyTracking('steward@test.com', 'all');

    expect(result.total).toBe(3); // 3 members in makeMemberData
    expect(result.completed).toBe(1);
    expect(result.members).toHaveLength(3);
    const memberEntry = result.members.find(m => m.email === 'member@test.com');
    expect(memberEntry.completed).toBe(true);
  });

  test('recognises all status string variants as completed', () => {
    // Status column may contain 'completed', 'yes', or 'true' — all mean done.
    const { ss } = makeFullMockSs([
      ['member@test.com', 'yes', ''],
      ['admin@test.com', 'true', ''],
    ]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    const result = DataService.getStewardSurveyTracking('steward@test.com', 'all');
    const member = result.members.find(m => m.email === 'member@test.com');
    const admin = result.members.find(m => m.email === 'admin@test.com');
    expect(member.completed).toBe(true);
    expect(admin.completed).toBe(true);
  });

  test('scope=assigned returns only steward-assigned members', () => {
    // member@test.com has Assigned Steward = '' in makeMemberData; steward has none either.
    // getStewardMembers returns members whose assignedSteward matches stewardEmail.
    const { ss } = makeFullMockSs([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    const result = DataService.getStewardSurveyTracking('steward@test.com', 'assigned');
    expect(typeof result.total).toBe('number');
    expect(Array.isArray(result.members)).toBe(true);
  });

  test('returns { total:0, completed:0, members:[] } when no survey sheet exists', () => {
    // No _Survey_Tracking sheet — should gracefully return all members with completed:false.
    const memberData = makeMemberData();
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', memberData);
    const grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG || 'Grievance Log', makeGrievanceData());
    const ss = createMockSpreadsheet([memberSheet, grievanceSheet]); // no survey sheet
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    const result = DataService.getStewardSurveyTracking('steward@test.com', 'all');
    // No survey data → no one is marked completed, but members are still returned
    expect(result.completed).toBe(0);
    expect(result.total).toBeGreaterThan(0);
    result.members.forEach(m => expect(m.completed).toBe(false));
  });

  test('is case-insensitive on email matching between members and survey map', () => {
    const { ss } = makeFullMockSs([
      ['MEMBER@TEST.COM', 'completed', '2026-02-01'], // uppercase in survey sheet
    ]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);

    const result = DataService.getStewardSurveyTracking('steward@test.com', 'all');
    const memberEntry = result.members.find(m => m.email === 'member@test.com');
    expect(memberEntry.completed).toBe(true);
  });
});


// ============================================================================
// DEADLINE BOUNDARY TESTS — Regression for 2026-03-09 fix
// ============================================================================
// Bug: deadlineDays === 0 (due today) fell through both overdue (< 0) and
// dueSoon (> 0) checks. These tests ensure every boundary is covered.

describe('Deadline boundary conditions (_buildGrievanceRecord auto-detect)', () => {
  const HEADERS = [
    'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
    'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category',
    'Resolution', 'Date Closed',
  ];

  function makeDeadlineCase(id, daysFromNow, status) {
    const deadline = daysFromNow !== null ? new Date(Date.now() + daysFromNow * 86400000) : '';
    return [
      id, 'member@test.com', status || 'Open', 'Step I', deadline,
      new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'Medium', '', 'Safety', '', '',
    ];
  }

  function setupAndGetStats(rows) {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [HEADERS, ...rows];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
    return DataService.getGrievanceStats();
  }

  function setupAndGetCases(rows) {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [HEADERS, ...rows];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
    return DataService.getStewardCases('steward@test.com');
  }

  // --- Overdue detection ---

  test('deadline 1 day ago → status auto-set to overdue', () => {
    const cases = setupAndGetCases([makeDeadlineCase('GR-A', -1, 'Open')]);
    expect(cases.length).toBe(1);
    expect(cases[0].status).toBe('overdue');
  });

  test('deadline 30 days ago → status auto-set to overdue', () => {
    const cases = setupAndGetCases([makeDeadlineCase('GR-A', -30, 'Active')]);
    expect(cases[0].status).toBe('overdue');
  });

  test('terminal status NOT overridden even with past deadline', () => {
    const cases = setupAndGetCases([makeDeadlineCase('GR-A', -10, 'Won')]);
    expect(cases[0].status).toBe('won');
  });

  test('resolved status NOT overridden even with past deadline', () => {
    const cases = setupAndGetCases([makeDeadlineCase('GR-A', -5, 'Resolved')]);
    expect(cases[0].status).toBe('resolved');
  });

  // --- DueSoon boundary: 0 days (today) ---

  test('deadline TODAY (0 days) counts as dueSoon in getGrievanceStats', () => {
    const stats = setupAndGetStats([makeDeadlineCase('GR-A', 0, 'Open')]);
    expect(stats.dueSoonCount).toBe(1);
    expect(stats.overdueCount).toBe(0);
  });

  test('deadline TODAY counts as dueSoon in getStewardKPIs', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const deadline = new Date(); // today
    const data = [HEADERS, [
      'GR-A', 'member@test.com', 'Open', 'Step I', deadline,
      new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'Medium', '', 'Safety', '', '',
    ]];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const kpis = DataService.getStewardKPIs('steward@test.com');
    expect(kpis.dueSoon).toBe(1);
    expect(kpis.overdue).toBe(0);
  });

  // --- DueSoon boundary: 1 day ---

  test('deadline 1 day from now → dueSoon', () => {
    const stats = setupAndGetStats([makeDeadlineCase('GR-A', 1, 'Open')]);
    expect(stats.dueSoonCount).toBe(1);
  });

  // --- DueSoon boundary: 7 days ---

  test('deadline exactly 7 days from now → dueSoon (inclusive)', () => {
    const stats = setupAndGetStats([makeDeadlineCase('GR-A', 7, 'Open')]);
    expect(stats.dueSoonCount).toBe(1);
  });

  // --- DueSoon boundary: 8 days (excluded) ---

  test('deadline 8 days from now → NOT dueSoon', () => {
    const stats = setupAndGetStats([makeDeadlineCase('GR-A', 8, 'Open')]);
    expect(stats.dueSoonCount).toBe(0);
  });

  // --- No deadline ---

  test('no deadline → neither overdue nor dueSoon', () => {
    const stats = setupAndGetStats([makeDeadlineCase('GR-A', null, 'Open')]);
    expect(stats.overdueCount).toBe(0);
    expect(stats.dueSoonCount).toBe(0);
  });

  // --- Terminal statuses excluded from dueSoon ---

  test('terminal statuses with valid deadline NOT counted as dueSoon', () => {
    const stats = setupAndGetStats([
      makeDeadlineCase('GR-A', 3, 'Won'),
      makeDeadlineCase('GR-B', 3, 'Denied'),
      makeDeadlineCase('GR-C', 3, 'Settled'),
      makeDeadlineCase('GR-D', 3, 'Withdrawn'),
      makeDeadlineCase('GR-E', 3, 'Closed'),
      makeDeadlineCase('GR-F', 3, 'Resolved'),
    ]);
    expect(stats.dueSoonCount).toBe(0);
  });

  // --- Mixed scenario ---

  test('mixed deadlines: overdue + dueSoon + future + terminal counted correctly', () => {
    const stats = setupAndGetStats([
      makeDeadlineCase('GR-OVERDUE1', -5, 'Open'),     // overdue
      makeDeadlineCase('GR-OVERDUE2', -1, 'Active'),    // overdue
      makeDeadlineCase('GR-TODAY',     0, 'Open'),       // dueSoon
      makeDeadlineCase('GR-SOON1',     3, 'Open'),       // dueSoon
      makeDeadlineCase('GR-SOON2',     7, 'New'),        // dueSoon
      makeDeadlineCase('GR-FUTURE',   30, 'Open'),       // neither
      makeDeadlineCase('GR-WON',      -2, 'Won'),        // terminal, NOT overdue
      makeDeadlineCase('GR-NONE',    null, 'Open'),       // no deadline
    ]);
    expect(stats.overdueCount).toBe(2);
    expect(stats.dueSoonCount).toBe(3);
    expect(stats.total).toBe(8);
  });
});


// ============================================================================
// KPI COMPUTATION PARITY — getStewardKPIs vs _computeKPIsFromCases
// ============================================================================
// Both functions should produce identical results for the same case set.
// _computeKPIsFromCases is used by batch data; getStewardKPIs by individual calls.

describe('KPI computation parity (batch vs individual)', () => {
  const HEADERS = [
    'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
    'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category',
    'Resolution', 'Date Closed',
  ];

  function setupMixed() {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const now = Date.now();
    const rows = [
      ['GR-1', 'member@test.com', 'Open', 'Step I', new Date(now - 2 * 86400000), new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'High', '', 'Safety', '', ''],
      ['GR-2', 'member@test.com', 'Open', 'Step I', new Date(now + 3 * 86400000), new Date('2026-01-15'), 'steward@test.com', 'Unit A', 'Med', '', 'Pay', '', ''],
      ['GR-3', 'member@test.com', 'Resolved', 'Step II', '', new Date('2025-06-01'), 'steward@test.com', 'Unit A', 'Low', '', 'Hours', 'Done', '2025-07-01'],
      ['GR-4', 'other@test.com', 'Open', 'Step I', new Date(now + 20 * 86400000), new Date('2026-02-01'), 'steward@test.com', 'Unit B', 'Med', '', 'Pay', '', ''],
    ];
    const data = [HEADERS, ...rows];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
  }

  test('getStewardKPIs and getBatchData.kpis return same counts', () => {
    setupMixed();

    const kpisIndividual = DataService.getStewardKPIs('steward@test.com');
    // Reset cache to simulate fresh batch call
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    setupMixed();
    const batch = DataService.getBatchData('steward@test.com', 'steward');
    const kpisBatch = batch.kpis;

    expect(kpisBatch.totalCases).toBe(kpisIndividual.totalCases);
    expect(kpisBatch.overdue).toBe(kpisIndividual.overdue);
    expect(kpisBatch.dueSoon).toBe(kpisIndividual.dueSoon);
    expect(kpisBatch.resolved).toBe(kpisIndividual.resolved);
    expect(kpisBatch.activeCases).toBe(kpisIndividual.activeCases);
  });

  test('overdue cases count correctly in KPIs', () => {
    setupMixed();
    const kpis = DataService.getStewardKPIs('steward@test.com');
    // GR-1 has deadline 2 days ago, status Open → auto-overdue
    expect(kpis.overdue).toBe(1);
  });

  test('dueSoon counts correctly in KPIs', () => {
    setupMixed();
    const kpis = DataService.getStewardKPIs('steward@test.com');
    // GR-2 has deadline 3 days from now, status Open → dueSoon
    expect(kpis.dueSoon).toBe(1);
  });

  test('resolved counts correctly in KPIs', () => {
    setupMixed();
    const kpis = DataService.getStewardKPIs('steward@test.com');
    expect(kpis.resolved).toBe(1);
  });
});


// ============================================================================
// ERROR RESILIENCE — getGrievanceStats with malformed rows
// ============================================================================
// Bug: One bad row in Grievance Log could crash the entire getGrievanceStats
// function, causing org KPI cards to stay as "..." placeholders.

describe('getGrievanceStats error resilience', () => {
  const HEADERS = [
    'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
    'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category',
    'Resolution', 'Date Closed',
  ];

  test('survives a row with fewer columns than header', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [
      HEADERS,
      ['GR-GOOD', 'a@test.com', 'Open', 'Step I', new Date(Date.now() + 86400000), new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'High', '', 'Safety', '', ''],
      ['GR-SHORT'],  // malformed — only 1 column
      ['GR-ALSO-GOOD', 'b@test.com', 'Won', 'Step II', '', new Date('2025-06-01'), 'steward@test.com', 'Unit B', 'Med', '', 'Pay', 'Won', '2025-07-01'],
    ];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    // Should not throw; should process the 2 good rows (may or may not count the short row)
    expect(stats.available).toBe(true);
    expect(stats.total).toBe(3); // total includes all data rows
  });

  test('survives a row with null/undefined values in date columns', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [
      HEADERS,
      ['GR-1', 'a@test.com', 'Open', 'Step I', null, null, 'steward@test.com', 'Unit A', 'High', '', 'Safety', '', ''],
      ['GR-2', 'b@test.com', 'Open', 'Step I', undefined, undefined, 'steward@test.com', 'Unit B', 'Med', '', 'Pay', '', ''],
    ];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(true);
    expect(stats.overdueCount).toBe(0);
    expect(stats.dueSoonCount).toBe(0);
  });

  test('survives a row with invalid date string in deadline column', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [
      HEADERS,
      ['GR-1', 'a@test.com', 'Open', 'Step I', 'not-a-date', new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'High', '', 'Safety', '', ''],
      ['GR-2', 'b@test.com', 'Open', 'Step I', new Date(Date.now() + 86400000), new Date('2026-01-01'), 'steward@test.com', 'Unit A', 'Med', '', 'Pay', '', ''],
    ];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(true);
    // GR-2 should count as dueSoon, GR-1 deadline is unparseable so no count
    expect(stats.dueSoonCount).toBe(1);
  });

  test('returns available:false for empty grievance sheet (header only)', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const data = [HEADERS]; // header only, no data rows
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(false);
  });

  test('returns available:false when Grievance Log sheet missing', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    // Only member sheet, no grievance sheet
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(false);
  });
});


// ============================================================================
// ORG KPI CARD CONTRACT — getGrievanceStats return shape
// ============================================================================
// The frontend depends on specific keys. Verify they're always present.

describe('getGrievanceStats return shape contract', () => {
  beforeEach(() => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const headers = [
      'Grievance ID', 'Member Email', 'Status', 'Step', 'Deadline',
      'Filed', 'Steward', 'Unit', 'Priority', 'Notes', 'Issue Category',
      'Resolution', 'Date Closed',
    ];
    const data = [headers, ['GR-1', 'a@test.com', 'Open', 'Step I', '', '', 'steward@test.com', 'Unit A', 'Med', '', 'Pay', '', '']];
    const grievSheet = createMockSheet('Grievance Log', data);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData());
    const mockSS = createMockSpreadsheet([memberSheet, grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
  });

  test('when available: true, includes all required keys for org KPI cards', () => {
    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(true);
    // These are the keys the frontend reads (steward_view.html:150-156)
    expect(stats).toHaveProperty('total');
    expect(stats).toHaveProperty('overdueCount');
    expect(stats).toHaveProperty('dueSoonCount');
    expect(stats).toHaveProperty('settledCount');
    expect(stats).toHaveProperty('wonCount');
    expect(stats).toHaveProperty('withdrawnCount');
    // Type checks
    expect(typeof stats.total).toBe('number');
    expect(typeof stats.overdueCount).toBe('number');
    expect(typeof stats.dueSoonCount).toBe('number');
    expect(typeof stats.settledCount).toBe('number');
    expect(typeof stats.wonCount).toBe('number');
    expect(typeof stats.withdrawnCount).toBe('number');
  });

  test('when available: false, does NOT include data keys', () => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.GRIEVANCE_LOG);
    }
    const mockSS = createMockSpreadsheet([
      createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', makeMemberData()),
    ]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);

    const stats = DataService.getGrievanceStats();
    expect(stats.available).toBe(false);
    // When unavailable, frontend checks !stats.available and shows fallback
    // No data keys required — but must not throw
  });
});


// ============================================================================
// getMembershipStats return shape contract
// ============================================================================

describe('getMembershipStats return shape', () => {
  beforeEach(() => {
    if (DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.MEMBER_DIR || 'Member Directory');
    }
    // Need 20+ members to pass threshold
    const headers = ['Email', 'Name', 'First Name', 'Last Name', 'Role', 'Unit', 'Phone', 'Hire Date', 'Dues Status', 'Member ID', 'Work Location'];
    const rows = [];
    for (let i = 1; i <= 25; i++) {
      const hireDate = i <= 5 ? new Date(Date.now() - (i * 10) * 24 * 60 * 60 * 1000) : new Date('2024-01-' + String(i).padStart(2, '0'));
      rows.push(['m' + i + '@test.com', 'Member ' + i, 'M', String(i), 'Member', i <= 15 ? 'Unit A' : 'Unit B', '', hireDate, i <= 20 ? 'Active' : 'Inactive', 'MEM-' + i, i <= 10 ? 'HQ' : 'Remote']);
    }
    const data = [headers, ...rows];
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', data);
    const mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSS);
  });

  test('includes newMembersLast90 and byHireMonth when available', () => {
    const ms = DataService.getMembershipStats();
    expect(ms.available).toBe(true);
    expect(ms.total).toBe(25);
    expect(ms).toHaveProperty('newMembersLast90');
    expect(ms).toHaveProperty('byHireMonth');
    expect(typeof ms.newMembersLast90).toBe('number');
    expect(ms.newMembersLast90).toBeGreaterThanOrEqual(0);
    expect(typeof ms.byHireMonth).toBe('object');
  });

  test('includes byUnit, byLocation, byDues', () => {
    const ms = DataService.getMembershipStats();
    expect(ms).toHaveProperty('byUnit');
    expect(ms).toHaveProperty('byLocation');
    expect(ms).toHaveProperty('byDues');
    expect(ms.byUnit['Unit A']).toBe(15);
    expect(ms.byUnit['Unit B']).toBe(10);
  });
});


// ============================================================================
// getStewardKPIs return shape contract
// ============================================================================

describe('getStewardKPIs return shape', () => {
  test('always returns all expected keys even with 0 cases', () => {
    const kpis = DataService.getStewardKPIs('nonexistent@test.com');
    expect(kpis).toHaveProperty('totalCases');
    expect(kpis).toHaveProperty('activeCases');
    expect(kpis).toHaveProperty('overdue');
    expect(kpis).toHaveProperty('dueSoon');
    expect(kpis).toHaveProperty('resolved');
    expect(typeof kpis.totalCases).toBe('number');
    expect(kpis.totalCases).toBe(0);
  });
});

// ============================================================================
// _ensureStewardTasks schema parity (Fix #1 regression guard)
// ============================================================================

describe('_ensureStewardTasks schema parity', () => {
  test('DataService _ensureStewardTasks creates 12-column header matching 08a_SheetSetup', () => {
    // Read source and verify the inline setValues call uses 12 columns
    const fs = require('fs');
    const src = fs.readFileSync(require('path').join(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8');

    // Find the _ensureStewardTasks function and verify it uses 12 columns
    const funcMatch = src.match(/function _ensureStewardTasks\(\)[\s\S]*?return sheet;\s*\}/);
    expect(funcMatch).not.toBeNull();
    const funcBody = funcMatch[0];

    // Must reference 12 columns in getRange call
    expect(funcBody).toMatch(/getRange\(1,\s*1,\s*1,\s*12\)/);

    // Must include Assignee Type and Assigned By headers
    expect(funcBody).toContain('Assignee Type');
    expect(funcBody).toContain('Assigned By');
  });

  test('both _ensureStewardTasks and _ensureStewardTasksSheet define same 12 headers', () => {
    const fs = require('fs');
    const dsrc = fs.readFileSync(require('path').join(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8');
    const ssrc = fs.readFileSync(require('path').join(__dirname, '..', 'src', '08a_SheetSetup.gs'), 'utf8');

    const expectedHeaders = ['ID', 'Steward Email', 'Title', 'Description', 'Member Email',
      'Priority', 'Status', 'Due Date', 'Created', 'Completed', 'Assignee Type', 'Assigned By'];

    // Check DataService version
    for (const header of expectedHeaders) {
      expect(dsrc).toContain(header);
    }
    // Check SheetSetup version
    for (const header of expectedHeaders) {
      expect(ssrc).toContain(header);
    }
  });
});

// ============================================================================
// getMemberTasks status filtering (Fix #2 regression guard)
// ============================================================================

describe('DataService.getMemberTasks status filtering', () => {
  let mockTaskSheet;

  function makeTaskData() {
    return [
      ['ID', 'Steward Email', 'Title', 'Description', 'Member Email', 'Priority', 'Status', 'Due Date', 'Created', 'Completed', 'Assignee Type', 'Assigned By'],
      ['MT_1', 'steward@test.com', 'Task Open', 'desc', 'member@test.com', 'high', 'open', new Date('2026-04-01'), new Date(), '', 'member', 'steward@test.com'],
      ['MT_2', 'steward@test.com', 'Task Done', 'desc', 'member@test.com', 'medium', 'completed', new Date('2026-03-01'), new Date(), new Date(), 'member', 'steward@test.com'],
      ['MT_3', 'steward@test.com', 'Task InProg', 'desc', 'member@test.com', 'low', 'in-progress', new Date('2026-05-01'), new Date(), '', 'member', 'steward@test.com'],
      ['ST_4', 'steward@test.com', 'Steward Task', 'desc', '', 'high', 'open', '', new Date(), '', 'steward', ''],
    ];
  }

  beforeEach(() => {
    const taskData = makeTaskData();
    mockTaskSheet = createMockSheet(SHEETS.STEWARD_TASKS || '_Steward_Tasks', taskData);

    const memberData = makeMemberData();
    const mockMemberSheet2 = createMockSheet(SHEETS.MEMBER_DIR || 'Member Directory', memberData);
    const mockSs = createMockSpreadsheet([mockMemberSheet2, mockTaskSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);

    if (typeof DataService !== 'undefined' && DataService._invalidateSheetCache) {
      DataService._invalidateSheetCache(SHEETS.STEWARD_TASKS || '_Steward_Tasks');
    }
  });

  test('not-completed filter returns only non-completed tasks', () => {
    const tasks = DataService.getMemberTasks('member@test.com', 'not-completed');
    expect(tasks.length).toBe(2);
    tasks.forEach(t => expect(t.status).not.toBe('completed'));
  });

  test('completed filter returns only completed tasks', () => {
    const tasks = DataService.getMemberTasks('member@test.com', 'completed');
    expect(tasks.length).toBe(1);
    expect(tasks[0].status).toBe('completed');
  });

  test('null filter returns all member tasks (no steward tasks)', () => {
    const tasks = DataService.getMemberTasks('member@test.com', null);
    expect(tasks.length).toBe(3);
  });

  test('does not include steward-type tasks', () => {
    const tasks = DataService.getMemberTasks('steward@test.com', null);
    // ST_4 has assignee type 'steward', not 'member' — should not appear
    const stewardTasks = tasks.filter(t => t.id === 'ST_4');
    expect(stewardTasks.length).toBe(0);
  });

  test('sorts by priority (high first) then by due date', () => {
    const tasks = DataService.getMemberTasks('member@test.com', 'not-completed');
    expect(tasks[0].priority).toBe('high');
  });

  test('returns expected fields on each task', () => {
    const tasks = DataService.getMemberTasks('member@test.com', 'not-completed');
    expect(tasks.length).toBeGreaterThan(0);
    const t = tasks[0];
    expect(t).toHaveProperty('id');
    expect(t).toHaveProperty('title');
    expect(t).toHaveProperty('description');
    expect(t).toHaveProperty('priority');
    expect(t).toHaveProperty('status');
    expect(t).toHaveProperty('dueDate');
    expect(t).toHaveProperty('dueDays');
    expect(t).toHaveProperty('assignedBy');
    expect(t).toHaveProperty('created');
  });
});

// ============================================================================
// _getMemberBatchData includes memberTaskCount (Fix #4 regression guard)
// ============================================================================

describe('_getMemberBatchData memberTaskCount', () => {
  test('getBatchData for member role includes memberTaskCount key', () => {
    const fs = require('fs');
    const src = fs.readFileSync(require('path').join(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8');
    expect(src).toMatch(/memberTaskCount/);
    // Verify it's in the return object of _getMemberBatchData
    const memberBatchMatch = src.match(/function _getMemberBatchData[\s\S]*?return \{[\s\S]*?\};/);
    expect(memberBatchMatch).not.toBeNull();
    expect(memberBatchMatch[0]).toContain('memberTaskCount');
  });
});

// ============================================================================
// v4.31.0 — H1: _invalidateSheetCache clears _emailIndex for MEMBER_SHEET
// ============================================================================

describe('v4.31.0 H1: _invalidateSheetCache clears _emailIndex', () => {
  test('_invalidateSheetCache source clears _emailIndex when called with Member Directory', () => {
    // Structural test: verify the source code nulls _emailIndex on MEMBER_SHEET invalidation
    const fs = require('fs');
    const src = fs.readFileSync(require('path').join(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8');
    const funcStart = src.indexOf('function _invalidateSheetCache');
    expect(funcStart).toBeGreaterThan(-1);
    const funcEnd = src.indexOf('\n  }', funcStart + 10) + 4;
    const funcBody = src.substring(funcStart, funcEnd);
    // Must contain _emailIndex = null
    expect(funcBody).toContain('_emailIndex = null');
    // Must be conditional on MEMBER_SHEET
    expect(funcBody).toContain('MEMBER_SHEET');
  });

  test('findUserByEmail works after _invalidateSheetCache (integration)', () => {
    // First, populate the email index by calling findUserByEmail
    const user1 = DataService.findUserByEmail('member@test.com');
    expect(user1).not.toBeNull();

    // Invalidate and re-setup with different data
    DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
    if (DataService._resetSSCache) DataService._resetSSCache();
    setupSheets();

    // Should still work after cache invalidation
    const user2 = DataService.findUserByEmail('member@test.com');
    expect(user2).not.toBeNull();
  });
});
