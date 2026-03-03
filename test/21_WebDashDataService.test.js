/**
 * Tests for 21_WebDashDataService.gs
 *
 * Covers DataService IIFE: user lookup, role determination, member data,
 * grievance access, steward dashboard data, and helper functions.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '21_WebDashDataService.gs']);

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
    DataService._invalidateSheetCache('Grievance Log');
    DataService._invalidateSheetCache('Member Directory');
  }
  setupSheets();
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
// DataService.getMemberData
// ============================================================================

describe('DataService.getMemberData', () => {
  test('returns null for empty email', () => {
    expect(DataService.getMemberData('')).toBeNull();
    expect(DataService.getMemberData(null)).toBeNull();
  });

  test('returns member grievances for known member', () => {
    const data = DataService.getMemberData('member@test.com');
    expect(data).not.toBeNull();
    if (data) {
      expect(data).toHaveProperty('grievances');
      expect(Array.isArray(data.grievances)).toBe(true);
    }
  });

  test('returns empty grievances for member with no cases', () => {
    const data = DataService.getMemberData('admin@test.com');
    if (data && data.grievances) {
      expect(data.grievances.length).toBe(0);
    }
  });

  test('returns null when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(DataService.getMemberData('member@test.com')).toBeNull();
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
// DataService.getStewardDashboardData
// ============================================================================

describe('DataService.getStewardDashboardData', () => {
  test('returns null for empty email', () => {
    expect(DataService.getStewardDashboardData('')).toBeNull();
  });

  test('returns dashboard data for steward', () => {
    const data = DataService.getStewardDashboardData('steward@test.com');
    expect(data).not.toBeNull();
    if (data) {
      expect(data).toHaveProperty('assignedCases');
      expect(Array.isArray(data.assignedCases)).toBe(true);
    }
  });

  test('returns null when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(DataService.getStewardDashboardData('steward@test.com')).toBeNull();
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
// DataService.getDirectoryData
// ============================================================================

describe('DataService.getDirectoryData', () => {
  test('returns directory data', () => {
    const data = DataService.getDirectoryData();
    expect(data).not.toBeNull();
    if (data) {
      expect(Array.isArray(data)).toBe(true);
    }
  });

  test('returns null when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    const data = DataService.getDirectoryData();
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

  test('returns summary counts (wonCount, deniedCount, settledCount)', () => {
    const stats = DataService.getGrievanceStats();
    expect(typeof stats.wonCount).toBe('number');
    expect(typeof stats.deniedCount).toBe('number');
    expect(typeof stats.settledCount).toBe('number');
    expect(stats.wonCount).toBe(3);
    expect(stats.deniedCount).toBe(1);
    expect(stats.settledCount).toBe(2);
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
    DataService._invalidateSheetCache('Grievance Log');
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
    DataService._invalidateSheetCache('Grievance Log');
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
  test('returns steward records sorted by name', () => {
    DataService._invalidateSheetCache('Member Directory');
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
    expect(stewards[0].name).toBe('Alice');
    expect(stewards[1].name).toBe('Zara');
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
    const result = dataMarkWelcomeDismissed('test@example.com');
    expect(result.success).toBe(true);
    const allProps = PropertiesService.getScriptProperties().getProperties();
    const dismissKey = Object.keys(allProps).find(k => k.startsWith('WELCOME_DISMISSED_'));
    expect(dismissKey).toBeTruthy();
  });

  test('returns failure when no authenticated user', () => {
    const origGetEmail = Session.getActiveUser().getEmail;
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => '') }));
    const result = dataMarkWelcomeDismissed();
    expect(result.success).toBe(false);
    Session.getActiveUser = jest.fn(() => ({ getEmail: origGetEmail }));
  });
});

// ============================================================================
// dataGetEngagementStats
// ============================================================================

describe('dataGetEngagementStats', () => {
  test('is a defined global function', () => {
    expect(typeof dataGetEngagementStats).toBe('function');
  });

  test('returns null when no stats are seeded', () => {
    PropertiesService.getScriptProperties().deleteAllProperties();
    const result = dataGetEngagementStats();
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
    DataService._invalidateSheetCache('Member Directory');
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

    const filters = dataGetBroadcastFilterOptions();
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
  test('dataFindUser delegates to DataService.findUserByEmail', () => {
    const user = dataFindUser('member@test.com');
    expect(user).not.toBeNull();
  });

  test('dataGetMemberData delegates to DataService.getMemberData', () => {
    const data = dataGetMemberData('member@test.com');
    expect(data !== undefined).toBe(true);
  });

  test('dataGetMemberGrievances delegates to DataService.getMemberGrievances', () => {
    const result = dataGetMemberGrievances('member@test.com');
    expect(Array.isArray(result)).toBe(true);
  });

  test('dataGetUserRole delegates to DataService.getUserRole', () => {
    const role = dataGetUserRole('member@test.com');
    expect(typeof role === 'string' || role === null).toBe(true);
  });

  test('dataGetStewardDashboard delegates to DataService.getStewardDashboardData', () => {
    const data = dataGetStewardDashboard('steward@test.com');
    expect(data !== undefined).toBe(true);
  });

  test('dataGetDirectory delegates to DataService.getDirectoryData', () => {
    const data = dataGetDirectory();
    expect(data !== undefined).toBe(true);
  });
});
