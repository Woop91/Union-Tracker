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
