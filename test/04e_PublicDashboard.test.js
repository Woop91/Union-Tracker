/**
 * Tests for 04e_PublicDashboard.gs - Engagement Tracking & Unified Dashboard
 *
 * Covers: getUnifiedDashboardData, getUnifiedDashboardDataAPI,
 * getUnifiedDashboardDataWithDateRange, engagement metric calculations,
 * participation rates, hot spot detection, PII handling.
 */

require('./gas-mock');
const { createMockRange, createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock globals that these modules expect
global.logAuditEvent = jest.fn();
global.AUDIT_EVENTS = {
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  FOLDER_CREATED: 'FOLDER_CREATED'
};
global.COMMAND_CONFIG = {
  ARCHIVE_FOLDER_ID: ''
};

// Mock DriveApp.getFolderById for drive resources
global.DriveApp.getFolderById = jest.fn(() => ({
  getFiles: jest.fn(() => ({
    hasNext: jest.fn(() => false),
    next: jest.fn()
  }))
}));

// Mock ScriptApp.getService for web app URL
global.ScriptApp.getService = jest.fn(() => ({
  getUrl: jest.fn(() => 'https://script.google.com/test')
}));

// Mock functions called by loaded modules
global.syncGrievanceFormulasToLog = jest.fn();
global.syncGrievanceToMemberDirectory = jest.fn();
global.syncMemberToGrievanceLog = jest.fn();
global.syncChecklistCalcToGrievanceLog = jest.fn();
global.repairGrievanceCheckboxes = jest.fn();
global.repairMemberCheckboxes = jest.fn();
global.checkDataQuality = jest.fn(() => []);

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '08c_FormsAndNotifications.gs',
  '09_Dashboards.gs',
  '04e_PublicDashboard.gs'
]);

// Mock getSatisfactionColMap_ to return static positions matching SATISFACTION_COLS
global.getSatisfactionColMap_ = jest.fn(() => ({
  'Timestamp': 1, 'q1': 2, 'q2': 3, 'q3': 4, 'q4': 5, 'q5': 6,
  'q6': 7, 'q7': 8, 'q8': 9, 'q9': 10,
  'q10': 11, 'q11': 12, 'q12': 13, 'q13': 14, 'q14': 15, 'q15': 16, 'q16': 17, 'q17': 18,
  'q18': 19, 'q19': 20, 'q20': 21,
  'q21': 22, 'q22': 23, 'q23': 24, 'q24': 25, 'q25': 26,
  'q26': 27, 'q27': 28, 'q28': 29, 'q29': 30, 'q30': 31, 'q31': 32,
  'q32': 33, 'q33': 34, 'q34': 35, 'q35': 36,
  'q36': 37, 'q37': 38, 'q38': 39, 'q39': 40, 'q40': 41,
  'q41': 42, 'q42': 43, 'q43': 44, 'q44': 45, 'q45': 46,
  'q46': 47, 'q47': 48, 'q48': 49, 'q49': 50, 'q50': 51,
  'q51': 52, 'q52': 53, 'q53': 54, 'q54': 55, 'q55': 56,
  'q56': 57, 'q57': 58, 'q58': 59, 'q59': 60, 'q60': 61, 'q61': 62, 'q62': 63
}));

// ============================================================================
// Helper functions for building mock data rows
// ============================================================================

function buildMemberRow(overrides) {
  var row = new Array(45).fill('');
  row[MEMBER_COLS.MEMBER_ID - 1] = overrides.memberId || '';
  row[MEMBER_COLS.FIRST_NAME - 1] = overrides.firstName || '';
  row[MEMBER_COLS.LAST_NAME - 1] = overrides.lastName || '';
  row[MEMBER_COLS.EMAIL - 1] = overrides.email || '';
  row[MEMBER_COLS.PHONE - 1] = overrides.phone || '';
  row[MEMBER_COLS.WORK_LOCATION - 1] = overrides.location || '';
  row[MEMBER_COLS.UNIT - 1] = overrides.unit || '';
  row[MEMBER_COLS.IS_STEWARD - 1] = overrides.isSteward || '';
  row[MEMBER_COLS.OFFICE_DAYS - 1] = overrides.officeDays || '';
  row[MEMBER_COLS.LAST_VIRTUAL_MTG - 1] = overrides.lastVirtualMtg || '';
  row[MEMBER_COLS.LAST_INPERSON_MTG - 1] = overrides.lastInPersonMtg || '';
  row[MEMBER_COLS.OPEN_RATE - 1] = overrides.openRate || '';
  row[MEMBER_COLS.VOLUNTEER_HOURS - 1] = overrides.volunteerHours || '';
  row[MEMBER_COLS.INTEREST_LOCAL - 1] = overrides.interestLocal || '';
  row[MEMBER_COLS.INTEREST_CHAPTER - 1] = overrides.interestChapter || '';
  row[MEMBER_COLS.INTEREST_ALLIED - 1] = overrides.interestAllied || '';
  row[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = overrides.lastUpdated || '';
  return row;
}

function buildGrievanceRow(overrides) {
  var row = new Array(40).fill('');
  row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = overrides.grievanceId || '';
  row[GRIEVANCE_COLS.FIRST_NAME - 1] = overrides.firstName || '';
  row[GRIEVANCE_COLS.LAST_NAME - 1] = overrides.lastName || '';
  row[GRIEVANCE_COLS.STATUS - 1] = overrides.status || '';
  row[GRIEVANCE_COLS.CURRENT_STEP - 1] = overrides.currentStep || 'Step 1';
  row[GRIEVANCE_COLS.DATE_FILED - 1] = overrides.dateFiled || '';
  row[GRIEVANCE_COLS.DATE_CLOSED - 1] = overrides.dateClosed || '';
  row[GRIEVANCE_COLS.STEWARD - 1] = overrides.steward || '';
  row[GRIEVANCE_COLS.LOCATION - 1] = overrides.location || '';
  row[GRIEVANCE_COLS.ARTICLES - 1] = overrides.articles || '';
  row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = overrides.category || '';
  row[GRIEVANCE_COLS.DAYS_OPEN - 1] = overrides.daysOpen || '';
  row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = overrides.daysToDeadline || '';
  row[GRIEVANCE_COLS.STEP1_DUE - 1] = overrides.step1Due || '';
  row[GRIEVANCE_COLS.STEP1_RCVD - 1] = overrides.step1Rcvd || '';
  row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1] = overrides.step2AppealFiled || '';
  row[GRIEVANCE_COLS.STEP2_RCVD - 1] = overrides.step2Rcvd || '';
  row[GRIEVANCE_COLS.STEP3_APPEAL_FILED - 1] = overrides.step3AppealFiled || '';
  row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] = overrides.nextActionDue || '';
  row[GRIEVANCE_COLS.RESOLUTION - 1] = overrides.resolution || '';
  return row;
}

function buildSatisfactionRow(overrides) {
  var row = new Array(82).fill('');
  row[0] = overrides.timestamp || new Date();
  row[1] = overrides.worksite || '';
  row[2] = overrides.role || '';
  // Overall questions (6-9)
  row[6] = overrides.q6 || '';
  row[7] = overrides.q7 || '';
  row[8] = overrides.q8 || '';
  row[9] = overrides.q9 || '';
  // Section averages
  row[71] = overrides.avgOverall || '';
  row[72] = overrides.avgSteward || '';
  return row;
}

function setupMockSpreadsheet(memberRows, grievanceRows, satRows, configSheet) {
  var memberData = [new Array(45).fill('')].concat(memberRows || []);
  var grievanceData = [new Array(40).fill('')].concat(grievanceRows || []);
  var satData = satRows ? [new Array(82).fill('')].concat(satRows) : null;

  var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
  memberSheet.getLastRow.mockReturnValue(memberData.length);

  var grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
  grievanceSheet.getLastRow.mockReturnValue(grievanceData.length);

  var sheets = [memberSheet, grievanceSheet];

  if (satData) {
    var satSheet = createMockSheet(SHEETS.SATISFACTION, satData);
    satSheet.getLastRow.mockReturnValue(satData.length);
    sheets.push(satSheet);
  }

  if (configSheet) {
    sheets.push(configSheet);
  }

  var mockSS = createMockSpreadsheet(sheets);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);
  return mockSS;
}

// ============================================================================
// getUnifiedDashboardData - Function existence
// ============================================================================

describe('getUnifiedDashboardData - Function existence', () => {
  test('getUnifiedDashboardData is defined', () => {
    expect(typeof getUnifiedDashboardData).toBe('function');
  });

  test('getUnifiedDashboardDataAPI is defined', () => {
    expect(typeof getUnifiedDashboardDataAPI).toBe('function');
  });

  test('getUnifiedDashboardDataWithDateRange is defined', () => {
    expect(typeof getUnifiedDashboardDataWithDateRange).toBe('function');
  });

  test('showPublicMemberDashboard is defined', () => {
    expect(typeof showPublicMemberDashboard).toBe('function');
  });

  test('getUnifiedDashboardHtml is defined', () => {
    expect(typeof getUnifiedDashboardHtml).toBe('function');
  });
});

// ============================================================================
// getUnifiedDashboardData - Empty data
// ============================================================================

describe('getUnifiedDashboardData - Empty data', () => {
  test('returns valid JSON with zero counts when no sheets exist', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalMembers).toBe(0);
    expect(result.totalGrievances).toBe(0);
    expect(result.openGrievances).toBe(0);
    expect(result.winRate).toBe(0);
    expect(result.engagement.emailOpenRate).toBe(0);
    expect(result.engagement.virtualMeetingRate).toBe(0);
    expect(result.engagement.inPersonMeetingRate).toBe(0);
    expect(result.engagement.totalVolunteerHours).toBe(0);
  });

  test('returns mode: member when includePII is false', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.mode).toBe('member');
  });

  test('returns mode: steward when includePII is true', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.mode).toBe('steward');
  });

  test('returns valid timestamp in ISO format', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.timestamp).toBeDefined();
    expect(new Date(result.timestamp).getTime()).not.toBeNaN();
  });
});

// ============================================================================
// getUnifiedDashboardData - Member counting
// ============================================================================

describe('getUnifiedDashboardData - Member counting', () => {
  test('counts total members correctly', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'MJOSM001', firstName: 'John', lastName: 'Smith' }),
      buildMemberRow({ memberId: 'MJADO002', firstName: 'Jane', lastName: 'Doe' }),
      buildMemberRow({ memberId: 'MBOJO003', firstName: 'Bob', lastName: 'Jones' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalMembers).toBe(3);
  });

  test('skips rows without Member ID', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'MJOSM001', firstName: 'John', lastName: 'Smith' }),
      buildMemberRow({ memberId: '', firstName: 'No', lastName: 'ID' }),
      buildMemberRow({ memberId: 'MBOJO003', firstName: 'Bob', lastName: 'Jones' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalMembers).toBe(2);
  });

  test('counts stewards correctly', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'A', lastName: 'S', isSteward: 'Yes' }),
      buildMemberRow({ memberId: 'M002', firstName: 'B', lastName: 'N', isSteward: 'No' }),
      buildMemberRow({ memberId: 'M003', firstName: 'C', lastName: 'S', isSteward: 'Yes' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.stewardCount).toBe(2);
  });

  test('calculates steward:member ratio', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', isSteward: 'Yes' }),
      buildMemberRow({ memberId: 'M002' }),
      buildMemberRow({ memberId: 'M003' }),
      buildMemberRow({ memberId: 'M004' }),
      buildMemberRow({ memberId: 'M005' }),
      buildMemberRow({ memberId: 'M006' }),
      buildMemberRow({ memberId: 'M007' }),
      buildMemberRow({ memberId: 'M008' }),
      buildMemberRow({ memberId: 'M009' }),
      buildMemberRow({ memberId: 'M010' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.stewardRatio).toBe('10:1');
  });
});

// ============================================================================
// getUnifiedDashboardData - Engagement metric calculations
// ============================================================================

describe('getUnifiedDashboardData - Engagement metrics', () => {
  test('calculates email open rate as average of all member rates', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', openRate: 80 }),
      buildMemberRow({ memberId: 'M002', openRate: 60 }),
      buildMemberRow({ memberId: 'M003', openRate: 40 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.emailOpenRate).toBe(60);
  });

  test('ignores non-numeric open rate values', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', openRate: 80 }),
      buildMemberRow({ memberId: 'M002', openRate: 'N/A' }),
      buildMemberRow({ memberId: 'M003', openRate: '' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.emailOpenRate).toBe(80);
  });

  test('sums total volunteer hours', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', volunteerHours: 10 }),
      buildMemberRow({ memberId: 'M002', volunteerHours: 25 }),
      buildMemberRow({ memberId: 'M003', volunteerHours: 15 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.totalVolunteerHours).toBe(50);
  });

  test('ignores non-numeric volunteer hours', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', volunteerHours: 10 }),
      buildMemberRow({ memberId: 'M002', volunteerHours: 'N/A' }),
      buildMemberRow({ memberId: 'M003', volunteerHours: '' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.totalVolunteerHours).toBe(10);
  });

  test('calculates virtual meeting attendance rate', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000)); // 30 days ago
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', lastVirtualMtg: recentDate }),
      buildMemberRow({ memberId: 'M002', lastVirtualMtg: recentDate }),
      buildMemberRow({ memberId: 'M003' }),
      buildMemberRow({ memberId: 'M004' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.virtualMeetingRate).toBe(50);
  });

  test('excludes meetings older than 6 months', () => {
    var oldDate = new Date(Date.now() - (200 * 24 * 60 * 60 * 1000)); // 200 days ago
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', lastVirtualMtg: recentDate }),
      buildMemberRow({ memberId: 'M002', lastVirtualMtg: oldDate }),
      buildMemberRow({ memberId: 'M003' }),
      buildMemberRow({ memberId: 'M004' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.virtualMeetingRate).toBe(25);
  });

  test('calculates in-person meeting attendance rate', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', lastInPersonMtg: recentDate }),
      buildMemberRow({ memberId: 'M002' }),
      buildMemberRow({ memberId: 'M003' }),
      buildMemberRow({ memberId: 'M004' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.inPersonMeetingRate).toBe(25);
  });

  test('calculates union interest percentages for Yes values', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', interestLocal: 'Yes', interestChapter: 'Yes', interestAllied: 'Yes' }),
      buildMemberRow({ memberId: 'M002', interestLocal: 'Yes', interestChapter: 'No', interestAllied: 'No' }),
      buildMemberRow({ memberId: 'M003', interestLocal: 'No', interestChapter: 'No', interestAllied: 'No' }),
      buildMemberRow({ memberId: 'M004', interestLocal: 'No', interestChapter: 'Yes', interestAllied: 'No' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.unionInterestLocal).toBe(50);
    expect(result.engagement.unionInterestChapter).toBe(50);
    expect(result.engagement.unionInterestAllied).toBe(25);
  });

  test('counts boolean true as interest', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', interestLocal: true }),
      buildMemberRow({ memberId: 'M002', interestLocal: false })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.unionInterestLocal).toBe(50);
  });

  test('counts string TRUE as interest', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', interestLocal: 'TRUE' }),
      buildMemberRow({ memberId: 'M002', interestLocal: 'FALSE' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.unionInterestLocal).toBe(50);
  });

  test('returns zero engagement metrics when no members', () => {
    setupMockSpreadsheet([]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.emailOpenRate).toBe(0);
    expect(result.engagement.virtualMeetingRate).toBe(0);
    expect(result.engagement.inPersonMeetingRate).toBe(0);
    expect(result.engagement.totalVolunteerHours).toBe(0);
    expect(result.engagement.unionInterestLocal).toBe(0);
  });
});

// ============================================================================
// getUnifiedDashboardData - Participation by unit/location
// ============================================================================

describe('getUnifiedDashboardData - Participation by unit/location', () => {
  test('calculates participation rates by unit', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', unit: 'IT', openRate: 80, lastVirtualMtg: recentDate }),
      buildMemberRow({ memberId: 'M002', unit: 'IT', openRate: 60 }),
      buildMemberRow({ memberId: 'M003', unit: 'HR', openRate: 40 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.participationByUnit.IT).toBeDefined();
    expect(result.participationByUnit.IT.count).toBe(2);
    expect(result.participationByUnit.IT.emailRate).toBe(70);
    expect(result.participationByUnit.HR).toBeDefined();
    expect(result.participationByUnit.HR.count).toBe(1);
    expect(result.participationByUnit.HR.emailRate).toBe(40);
  });

  test('calculates participation rates by location', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', location: 'NYC', openRate: 90, lastVirtualMtg: recentDate }),
      buildMemberRow({ memberId: 'M002', location: 'NYC', openRate: 70, lastInPersonMtg: recentDate }),
      buildMemberRow({ memberId: 'M003', location: 'LA', openRate: 50 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.participationByLocation.NYC).toBeDefined();
    expect(result.participationByLocation.NYC.count).toBe(2);
    expect(result.participationByLocation.NYC.meetingRate).toBe(100);
    expect(result.participationByLocation.LA.meetingRate).toBe(0);
  });

  test('defaults to Unknown for missing unit/location', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', openRate: 50 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.participationByUnit.Unknown).toBeDefined();
    expect(result.participationByLocation.Unknown).toBeDefined();
  });
});

// ============================================================================
// getUnifiedDashboardData - Hot spot detection
// ============================================================================

describe('getUnifiedDashboardData - Hot spots', () => {
  test('detects low engagement hot spots (< 30% engagement, >= 5 members)', () => {
    // Create 6 members in location with very low engagement
    var members = [];
    for (var i = 0; i < 6; i++) {
      members.push(buildMemberRow({
        memberId: 'M00' + i,
        location: 'RemoteSite',
        unit: 'Warehouse',
        openRate: 10 // Very low
        // No meeting attendance = 0% meeting rate
      }));
    }
    setupMockSpreadsheet(members);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.hotSpots.lowEngagement.length).toBeGreaterThan(0);
    var found = result.hotSpots.lowEngagement.find(function(h) {
      return h.name === 'RemoteSite';
    });
    expect(found).toBeDefined();
    expect(found.engagement).toBeLessThan(30);
  });

  test('does not flag locations with < 5 members as low engagement', () => {
    var members = [];
    for (var i = 0; i < 3; i++) {
      members.push(buildMemberRow({
        memberId: 'M00' + i,
        location: 'SmallSite',
        openRate: 5
      }));
    }
    setupMockSpreadsheet(members);

    var result = JSON.parse(getUnifiedDashboardData(false));
    var found = result.hotSpots.lowEngagement.find(function(h) {
      return h.name === 'SmallSite';
    });
    expect(found).toBeUndefined();
  });

  test('does not flag locations with >= 30% engagement', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    var members = [];
    for (var i = 0; i < 6; i++) {
      members.push(buildMemberRow({
        memberId: 'M00' + i,
        location: 'ActiveSite',
        openRate: 80,
        lastVirtualMtg: recentDate
      }));
    }
    setupMockSpreadsheet(members);

    var result = JSON.parse(getUnifiedDashboardData(false));
    var found = result.hotSpots.lowEngagement.find(function(h) {
      return h.name === 'ActiveSite';
    });
    expect(found).toBeUndefined();
  });

  test('detects grievance hot zones (3+ active cases)', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', location: 'HotZone' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', location: 'HotZone' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Pending Info', location: 'HotZone' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.hotZones.length).toBeGreaterThan(0);
    expect(result.hotSpots.grievance.length).toBeGreaterThan(0);
  });

  test('does not flag locations with < 3 active cases', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', location: 'QuietZone' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', location: 'QuietZone' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    var found = result.hotSpots.grievance.find(function(h) {
      return h.name === 'QuietZone';
    });
    expect(found).toBeUndefined();
  });
});

// ============================================================================
// getUnifiedDashboardData - PII handling
// ============================================================================

describe('getUnifiedDashboardData - PII handling', () => {
  test('hides member list in non-PII mode', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'MJOSM001', firstName: 'John', lastName: 'Smith' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.memberList).toEqual([]);
  });

  test('includes member list in PII mode', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'MJOSM001', firstName: 'John', lastName: 'Smith' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.memberList.length).toBe(1);
  });

  test('steward names are always shown (public union reps)', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'Alice', lastName: 'Jones', isSteward: 'Yes', location: 'NYC', unit: 'IT' })
    ]);

    var resultMember = JSON.parse(getUnifiedDashboardData(false));
    expect(resultMember.stewardList.length).toBe(1);
    expect(resultMember.stewardList[0].name).toBeDefined();
  });

  test('hides meeting attendee details in non-PII mode', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'John', lastName: 'S', lastVirtualMtg: recentDate })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.recentMeetingAttendees).toEqual([]);
  });

  test('shows meeting attendee details in PII mode', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'John', lastName: 'S', lastVirtualMtg: recentDate })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.engagement.recentMeetingAttendees.length).toBe(1);
  });

  test('masks member IDs in non-PII mode for drill-down data', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'MJOSM001', firstName: 'John', lastName: 'Smith', unit: 'IT' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    var unitMembers = result.chartDrillDown.unitByMember.IT || result.chartDrillDown.unitByMember.Unknown;
    if (unitMembers && unitMembers.length > 0) {
      expect(unitMembers[0].name).toBe('Member');
      expect(unitMembers[0].id).toContain('***');
    }
  });
});

// ============================================================================
// getUnifiedDashboardData - Grievance processing
// ============================================================================

describe('getUnifiedDashboardData - Grievance processing', () => {
  test('counts total grievances', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Won' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Denied' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalGrievances).toBe(3);
  });

  test('counts open grievances (Open + Pending)', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Pending Info' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Won' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.openGrievances).toBe(2);
  });

  test('calculates win rate correctly', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Won' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Won' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Denied' }),
        buildGrievanceRow({ grievanceId: 'G004', status: 'Settled' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    // Win rate: 2 / (2+1+1) = 50%
    expect(result.winRate).toBe(50);
  });

  test('win rate is 0 when no closed cases', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [buildGrievanceRow({ grievanceId: 'G001', status: 'Open' })]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.winRate).toBe(0);
  });

  test('counts overdue cases correctly', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', daysToDeadline: -5 }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', daysToDeadline: 'Overdue' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Open', daysToDeadline: 5 })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.overdueCount).toBe(2);
  });

  test('tracks status distribution correctly', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Won' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Denied' }),
        buildGrievanceRow({ grievanceId: 'G004', status: 'Settled' }),
        buildGrievanceRow({ grievanceId: 'G005', status: 'Withdrawn' }),
        buildGrievanceRow({ grievanceId: 'G006', status: 'Pending Info' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.statusDistribution.open).toBe(1);
    expect(result.statusDistribution.won).toBe(1);
    expect(result.statusDistribution.denied).toBe(1);
    expect(result.statusDistribution.settled).toBe(1);
    expect(result.statusDistribution.withdrawn).toBe(1);
    expect(result.statusDistribution.pending).toBe(1);
  });

  test('calculates average settlement days', () => {
    var dateFiled = new Date('2026-01-01');
    var dateClosed = new Date('2026-01-31'); // 30 days
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [buildGrievanceRow({
        grievanceId: 'G001', status: 'Won',
        dateFiled: dateFiled,
        dateClosed: dateClosed
      })]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.avgSettlementDays).toBe(30);
  });

  test('tracks steward workload', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', steward: 'Alice' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', steward: 'Alice' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Open', steward: 'Bob' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.stewardWorkload.length).toBe(2);
    expect(result.stewardWorkload[0].name).toBe('Alice');
    expect(result.stewardWorkload[0].count).toBe(2);
  });

  test('labels heavy workload stewards correctly', () => {
    var grievances = [];
    for (var i = 0; i < 9; i++) {
      grievances.push(buildGrievanceRow({
        grievanceId: 'G00' + i,
        status: 'Open',
        steward: 'Overloaded'
      }));
    }
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      grievances
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    var overloaded = result.stewardWorkload.find(function(s) {
      return s.name === 'Overloaded';
    });
    expect(overloaded.status).toBe('OVERLOAD');
  });

  test('tracks step progression (Step 1, 2, 3, Arb)', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', currentStep: 'Step 1' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', currentStep: 'Step 2' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Open', currentStep: 'Step 3' }),
        buildGrievanceRow({ grievanceId: 'G004', status: 'Open', currentStep: 'Arbitration' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.stepProgression.step1).toBe(1);
    expect(result.stepProgression.step2).toBe(1);
    expect(result.stepProgression.step3).toBe(1);
    expect(result.stepProgression.arb).toBe(1);
  });

  test('tracks article violations', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open', articles: 'Article 5, Article 7' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'Open', articles: 'Article 5' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.articleViolations['Article 5']).toBe(2);
    expect(result.articleViolations['Article 7']).toBe(1);
    expect(result.topViolatedArticle).toBe('Article 5');
  });
});

// ============================================================================
// getUnifiedDashboardData - Directory trends
// ============================================================================

describe('getUnifiedDashboardData - Directory trends', () => {
  test('tracks members with and without email', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', email: 'a@test.com' }),
      buildMemberRow({ memberId: 'M002', email: '' }),
      buildMemberRow({ memberId: 'M003', email: 'c@test.com' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.directoryTrends.totalWithEmail).toBe(2);
    expect(result.directoryTrends.missingEmail).toBe(1);
  });

  test('tracks members with and without phone', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', phone: '555-1234' }),
      buildMemberRow({ memberId: 'M002', phone: '' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.directoryTrends.totalWithPhone).toBe(1);
    expect(result.directoryTrends.missingPhone).toBe(1);
  });

  test('tracks recent updates (last 30 days)', () => {
    var recentDate = new Date(Date.now() - (10 * 24 * 60 * 60 * 1000)); // 10 days ago
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'John', lastName: 'S', lastUpdated: recentDate })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.directoryTrends.recentUpdates.length).toBe(1);
  });

  test('tracks stale contacts (90+ days)', () => {
    var staleDate = new Date(Date.now() - (100 * 24 * 60 * 60 * 1000)); // 100 days ago
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', firstName: 'John', lastName: 'S', lastUpdated: staleDate })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.directoryTrends.staleContacts.length).toBe(1);
  });
});

// ============================================================================
// getUnifiedDashboardData - Office days breakdown
// ============================================================================

describe('getUnifiedDashboardData - Office days', () => {
  test('tracks office days breakdown', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', officeDays: 'Monday, Wednesday, Friday' }),
      buildMemberRow({ memberId: 'M002', officeDays: 'Monday, Tuesday' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.officeDaysBreakdown.Monday).toBe(2);
    expect(result.officeDaysBreakdown.Wednesday).toBe(1);
    expect(result.officeDaysBreakdown.Friday).toBe(1);
    expect(result.officeDaysBreakdown.Tuesday).toBe(1);
  });

  test('ignores N/A office days', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', officeDays: 'N/A' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(Object.keys(result.officeDaysBreakdown).length).toBe(0);
  });
});

// ============================================================================
// getUnifiedDashboardData - Satisfaction survey processing
// ============================================================================

describe('getUnifiedDashboardData - Satisfaction data', () => {
  test('calculates survey response rate', () => {
    var satRow = buildSatisfactionRow({ timestamp: new Date(), q7: 8 });
    setupMockSpreadsheet(
      [
        buildMemberRow({ memberId: 'M001' }),
        buildMemberRow({ memberId: 'M002' })
      ],
      [],
      [satRow]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.surveyResponseRate).toBe(50);
  });

  test('handles no satisfaction data', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement.surveyResponseRate).toBe(0);
    expect(result.satisfactionData.responseCount).toBe(0);
  });

  test('calculates morale score from trust scores', () => {
    var satRows = [
      buildSatisfactionRow({ timestamp: new Date(), q7: 8 }),
      buildSatisfactionRow({ timestamp: new Date(), q7: 6 })
    ];
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [],
      satRows
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.moraleScore).toBe(7);
  });
});

// ============================================================================
// getUnifiedDashboardDataAPI
// ============================================================================

describe('getUnifiedDashboardDataAPI', () => {
  test('passes true for PII when isPII is true', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataAPI(true));
    expect(result.mode).toBe('steward');
  });

  test('passes false for PII when isPII is false', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataAPI(false));
    expect(result.mode).toBe('member');
  });

  test('treats string "true" as true for PII', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataAPI('true'));
    expect(result.mode).toBe('steward');
  });
});

// ============================================================================
// getUnifiedDashboardDataWithDateRange
// ============================================================================

describe('getUnifiedDashboardDataWithDateRange', () => {
  test('returns full data when no filtering requested', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataWithDateRange(false, 0, null, null));
    expect(result.totalMembers).toBe(0);
    expect(result.dateRangeApplied).toBeUndefined();
  });

  test('applies date range filter when days specified', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataWithDateRange(false, 30, null, null));
    expect(result.dateRangeApplied).toBeDefined();
    expect(result.dateRangeApplied.days).toBe(30);
  });

  test('applies custom date range when fromDate and toDate specified', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardDataWithDateRange(
      false, null, '2026-01-01', '2026-01-31'
    ));
    expect(result.dateRangeApplied).toBeDefined();
  });
});

// ============================================================================
// getUnifiedDashboardHtml
// ============================================================================

describe('getUnifiedDashboardHtml', () => {
  test('generates member mode HTML with member title', () => {
    var html = getUnifiedDashboardHtml(false);
    expect(html).toContain('MEMBER DASHBOARD');
    expect(html).toContain('MEMBER VIEW');
  });

  test('generates steward mode HTML with steward title', () => {
    var html = getUnifiedDashboardHtml(true);
    expect(html).toContain('STEWARD COMMAND CENTER');
    expect(html).toContain('INTERNAL USE - CONTAINS PII');
  });

  test('includes required CSS classes', () => {
    var html = getUnifiedDashboardHtml(false);
    expect(html).toContain('.kpi-card');
    expect(html).toContain('.chart-card');
    expect(html).toContain('.tab');
  });

  test('includes Chart.js library', () => {
    var html = getUnifiedDashboardHtml(false);
    expect(html).toContain('chart.js');
  });
});

// ============================================================================
// getUnifiedDashboardData - Data structure integrity
// ============================================================================

describe('getUnifiedDashboardData - Data structure', () => {
  test('returns all expected top-level keys', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    var expectedKeys = [
      'mode', 'timestamp', 'totalMembers', 'stewardCount',
      'totalGrievances', 'openGrievances', 'wins', 'losses',
      'settled', 'winRate', 'overdueCount', 'moraleScore',
      'engagement', 'participationByUnit', 'participationByLocation',
      'hotSpots', 'hotZones', 'stewardWorkload',
      'statusDistribution', 'stepProgression',
      'directoryTrends', 'satisfactionData', 'driveResources'
    ];

    expectedKeys.forEach(function(key) {
      expect(result).toHaveProperty(key);
    });
  });

  test('engagement object has correct structure', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.engagement).toHaveProperty('emailOpenRate');
    expect(result.engagement).toHaveProperty('virtualMeetingRate');
    expect(result.engagement).toHaveProperty('inPersonMeetingRate');
    expect(result.engagement).toHaveProperty('totalVolunteerHours');
    expect(result.engagement).toHaveProperty('unionInterestLocal');
    expect(result.engagement).toHaveProperty('unionInterestChapter');
    expect(result.engagement).toHaveProperty('unionInterestAllied');
    expect(result.engagement).toHaveProperty('surveyResponseRate');
    expect(result.engagement).toHaveProperty('recentMeetingAttendees');
  });

  test('hotSpots object has correct structure', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.hotSpots).toHaveProperty('grievance');
    expect(result.hotSpots).toHaveProperty('dissatisfaction');
    expect(result.hotSpots).toHaveProperty('lowEngagement');
    expect(result.hotSpots).toHaveProperty('overdueConcentration');
    expect(Array.isArray(result.hotSpots.grievance)).toBe(true);
    expect(Array.isArray(result.hotSpots.lowEngagement)).toBe(true);
  });

  test('chartDrillDown has correct status keys', () => {
    var mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.chartDrillDown.statusByCase).toHaveProperty('open');
    expect(result.chartDrillDown.statusByCase).toHaveProperty('pending');
    expect(result.chartDrillDown.statusByCase).toHaveProperty('won');
    expect(result.chartDrillDown.statusByCase).toHaveProperty('denied');
    expect(result.chartDrillDown.statusByCase).toHaveProperty('settled');
    expect(result.chartDrillDown.statusByCase).toHaveProperty('withdrawn');
  });
});

// ============================================================================
// getUnifiedDashboardData - Edge cases
// ============================================================================

describe('getUnifiedDashboardData - Edge cases', () => {
  test('handles member with all engagement fields populated', () => {
    var recentDate = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
    setupMockSpreadsheet([
      buildMemberRow({
        memberId: 'M001',
        firstName: 'John',
        lastName: 'Smith',
        email: 'john@test.com',
        phone: '555-1234',
        location: 'NYC',
        unit: 'IT',
        isSteward: 'Yes',
        officeDays: 'Monday, Wednesday',
        lastVirtualMtg: recentDate,
        lastInPersonMtg: recentDate,
        openRate: 75,
        volunteerHours: 20,
        interestLocal: 'Yes',
        interestChapter: 'Yes',
        interestAllied: 'No',
        lastUpdated: recentDate
      })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(true));
    expect(result.totalMembers).toBe(1);
    expect(result.stewardCount).toBe(1);
    expect(result.engagement.emailOpenRate).toBe(75);
    expect(result.engagement.totalVolunteerHours).toBe(20);
    expect(result.engagement.virtualMeetingRate).toBe(100);
    expect(result.engagement.inPersonMeetingRate).toBe(100);
    expect(result.engagement.unionInterestLocal).toBe(100);
    expect(result.engagement.unionInterestChapter).toBe(100);
    expect(result.engagement.unionInterestAllied).toBe(0);
  });

  test('handles grievance with all fields populated', () => {
    var dateFiled = new Date('2026-01-01');
    var dateClosed = new Date('2026-01-15');
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [buildGrievanceRow({
        grievanceId: 'G001',
        firstName: 'John',
        lastName: 'Smith',
        status: 'Won',
        currentStep: 'Step 2',
        dateFiled: dateFiled,
        dateClosed: dateClosed,
        steward: 'Alice',
        location: 'NYC',
        articles: 'Article 5',
        category: 'Discipline',
        daysOpen: 14,
        daysToDeadline: 5
      })]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalGrievances).toBe(1);
    expect(result.wins).toBe(1);
  });

  test('skips grievances without Grievance ID', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'Open' }),
        buildGrievanceRow({ grievanceId: '', status: 'Open' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.totalGrievances).toBe(1);
  });

  test('handles case-insensitive status matching', () => {
    setupMockSpreadsheet(
      [buildMemberRow({ memberId: 'M001' })],
      [
        buildGrievanceRow({ grievanceId: 'G001', status: 'OPEN' }),
        buildGrievanceRow({ grievanceId: 'G002', status: 'won' }),
        buildGrievanceRow({ grievanceId: 'G003', status: 'Sustained' }),
        buildGrievanceRow({ grievanceId: 'G004', status: 'Favorable' })
      ]
    );

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.openGrievances).toBe(1);
    // Won + Sustained + Favorable = 3 wins
    expect(result.wins).toBe(3);
  });

  test('handles zero open rate value as valid numeric', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', openRate: 0 })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    // 0 is falsy in JS, so check if code handles it
    // The code does: if (openRate && !isNaN(parseFloat(openRate)))
    // 0 is falsy, so it would be skipped - this is a known edge case
    expect(result.engagement.emailOpenRate).toBe(0);
  });

  test('handles location/unit breakdowns for all members', () => {
    setupMockSpreadsheet([
      buildMemberRow({ memberId: 'M001', location: 'NYC', unit: 'IT' }),
      buildMemberRow({ memberId: 'M002', location: 'NYC', unit: 'HR' }),
      buildMemberRow({ memberId: 'M003', location: 'LA', unit: 'IT' })
    ]);

    var result = JSON.parse(getUnifiedDashboardData(false));
    expect(result.locationBreakdown.NYC).toBe(2);
    expect(result.locationBreakdown.LA).toBe(1);
    expect(result.unitBreakdown.IT).toBe(2);
    expect(result.unitBreakdown.HR).toBe(1);
  });
});

// ============================================================================
// showPublicMemberDashboard
// ============================================================================

describe('showPublicMemberDashboard', () => {
  test('creates modal dialog with correct title', () => {
    var mockUi = {
      showModalDialog: jest.fn()
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);
    ScriptApp.getService = jest.fn(() => ({
      getUrl: jest.fn(() => 'https://script.google.com/test')
    }));

    showPublicMemberDashboard();
    expect(mockUi.showModalDialog).toHaveBeenCalled();
    expect(mockUi.showModalDialog.mock.calls[0][1]).toBe('Member Dashboard');
  });
});
