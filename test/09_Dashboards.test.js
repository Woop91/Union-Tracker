/**
 * Tests for 09_Dashboards.gs
 *
 * Covers: computeAverage_, getSheetLastRow, computeDashboardMetrics_,
 * getStewardCoverageStats, getPublicOverviewData, syncAllData,
 * computeSatisfactionRowAverages, computeSectionAverages_,
 * and syncSatisfactionValues.
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

// Mock functions called by syncAllData that live in other modules
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
  '09_Dashboards.gs'
]);

// ============================================================================
// computeAverage_ - pure math function
// ============================================================================

describe('computeAverage_', () => {
  test('computes average of all numeric values', () => {
    expect(computeAverage_([1, 2, 3, 4, 5], 0, 4)).toBe(3);
  });

  test('returns empty string when no valid numeric values', () => {
    expect(computeAverage_(['a', 'b', 'c'], 0, 2)).toBe('');
  });

  test('ignores non-numeric values in the range', () => {
    expect(computeAverage_([1, '', 3, null, 5], 0, 4)).toBe(3);
  });

  test('handles single value', () => {
    expect(computeAverage_([7], 0, 0)).toBe(7);
  });

  test('rounds to 2 decimal places', () => {
    // Average of [1, 2] = 1.5 -> rounds to 1.5
    expect(computeAverage_([1, 2], 0, 1)).toBe(1.5);
  });

  test('rounds correctly for repeating decimals', () => {
    // Average of [1, 1, 2] = 1.333... -> rounds to 1.33
    expect(computeAverage_([1, 1, 2], 0, 2)).toBe(1.33);
  });

  test('uses startIdx and endIdx as inclusive range', () => {
    // Only average indices 1-3: [20, 30, 40]
    expect(computeAverage_([10, 20, 30, 40, 50], 1, 3)).toBe(30);
  });

  test('handles all empty/null values', () => {
    expect(computeAverage_(['', null, undefined, 'text'], 0, 3)).toBe('');
  });

  test('handles zero values correctly (zero is numeric)', () => {
    expect(computeAverage_([0, 0, 0], 0, 2)).toBe(0);
  });

  test('handles mixed numeric and string-number values', () => {
    // '5' is a string, not typeof number, so should be skipped
    expect(computeAverage_([3, '5', 7], 0, 2)).toBe(5);
  });

  test('handles NaN values (skips them)', () => {
    expect(computeAverage_([2, NaN, 4], 0, 2)).toBe(3);
  });

  test('handles negative numbers', () => {
    expect(computeAverage_([-2, -4, -6], 0, 2)).toBe(-4);
  });

  test('handles decimal inputs', () => {
    expect(computeAverage_([1.5, 2.5], 0, 1)).toBe(2);
  });

  test('handles large arrays with partial range', () => {
    const row = new Array(100).fill(0);
    row[50] = 10;
    row[51] = 20;
    row[52] = 30;
    expect(computeAverage_(row, 50, 52)).toBe(20);
  });
});

// ============================================================================
// getSheetLastRow
// ============================================================================

describe('getSheetLastRow', () => {
  test('returns last row from sheet', () => {
    const mockSheet = createMockSheet('Test');
    mockSheet.getLastRow.mockReturnValue(3);

    const result = getSheetLastRow(mockSheet);
    expect(result).toBe(3);
  });

  test('returns 1 when only header exists', () => {
    const mockSheet = createMockSheet('Test');
    mockSheet.getLastRow.mockReturnValue(1);

    const result = getSheetLastRow(mockSheet);
    expect(result).toBe(1);
  });

  test('returns 0 for empty sheet', () => {
    const mockSheet = createMockSheet('Test');
    mockSheet.getLastRow.mockReturnValue(0);

    const result = getSheetLastRow(mockSheet);
    expect(result).toBe(0);
  });
});

// ============================================================================
// computeDashboardMetrics_ - data transform
// ============================================================================

describe('computeDashboardMetrics_', () => {
  // Build minimal header rows matching column constants
  function buildMemberRow(overrides) {
    const row = new Array(45).fill('');
    row[MEMBER_COLS.MEMBER_ID - 1] = overrides.memberId || '';
    row[MEMBER_COLS.IS_STEWARD - 1] = overrides.isSteward || '';
    row[MEMBER_COLS.OPEN_RATE - 1] = overrides.openRate || '';
    row[MEMBER_COLS.VOLUNTEER_HOURS - 1] = overrides.volunteerHours || '';
    row[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = overrides.contactDate || '';
    return row;
  }

  function buildGrievanceRow(overrides) {
    const row = new Array(40).fill('');
    row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] = overrides.grievanceId || '';
    row[GRIEVANCE_COLS.STATUS - 1] = overrides.status || '';
    row[GRIEVANCE_COLS.STEWARD - 1] = overrides.steward || '';
    row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] = overrides.category || '';
    row[GRIEVANCE_COLS.LOCATION - 1] = overrides.location || '';
    row[GRIEVANCE_COLS.DATE_FILED - 1] = overrides.dateFiled || '';
    row[GRIEVANCE_COLS.DATE_CLOSED - 1] = overrides.dateClosed || '';
    row[GRIEVANCE_COLS.DAYS_OPEN - 1] = overrides.daysOpen || '';
    row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = overrides.daysToDeadline || '';
    row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1] = overrides.nextActionDue || '';
    row[GRIEVANCE_COLS.RESOLUTION - 1] = overrides.resolution || '';
    return row;
  }

  test('returns metrics object with expected keys', () => {
    const memberData = [new Array(45).fill('')]; // header only
    const grievanceData = [new Array(40).fill('')]; // header only
    const configData = [];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, configData);

    expect(metrics).toHaveProperty('totalMembers');
    expect(metrics).toHaveProperty('activeStewards');
    expect(metrics).toHaveProperty('activeGrievances');
    expect(metrics).toHaveProperty('winRate');
    expect(metrics).toHaveProperty('overdueCases');
    expect(metrics).toHaveProperty('categories');
    expect(metrics).toHaveProperty('locations');
    expect(metrics).toHaveProperty('trends');
    expect(metrics).toHaveProperty('stewardSummary');
  });

  test('counts total members correctly', () => {
    const memberData = [
      new Array(45).fill(''), // header
      buildMemberRow({ memberId: 'M001' }),
      buildMemberRow({ memberId: 'M002' }),
      buildMemberRow({ memberId: 'M003' })
    ];
    const grievanceData = [new Array(40).fill('')];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.totalMembers).toBe(3);
  });

  test('counts stewards correctly', () => {
    const memberData = [
      new Array(45).fill(''), // header
      buildMemberRow({ memberId: 'M001', isSteward: 'Yes' }),
      buildMemberRow({ memberId: 'M002', isSteward: 'No' }),
      buildMemberRow({ memberId: 'M003', isSteward: 'Yes' })
    ];
    const grievanceData = [new Array(40).fill('')];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.activeStewards).toBe(2);
  });

  test('counts active grievances (Open + Pending Info)', () => {
    const memberData = [new Array(45).fill('')];
    const grievanceData = [
      new Array(40).fill(''), // header
      buildGrievanceRow({ grievanceId: 'GRV-001', status: 'Open' }),
      buildGrievanceRow({ grievanceId: 'GRV-002', status: 'Pending Info' }),
      buildGrievanceRow({ grievanceId: 'GRV-003', status: 'Won' }),
      buildGrievanceRow({ grievanceId: 'GRV-004', status: 'Denied' })
    ];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.activeGrievances).toBe(2);
    expect(metrics.open).toBe(1);
    expect(metrics.pendingInfo).toBe(1);
    expect(metrics.won).toBe(1);
    expect(metrics.denied).toBe(1);
  });

  test('calculates win rate as percentage', () => {
    const memberData = [new Array(45).fill('')];
    const grievanceData = [
      new Array(40).fill(''), // header
      buildGrievanceRow({ grievanceId: 'GRV-001', status: 'Won' }),
      buildGrievanceRow({ grievanceId: 'GRV-002', status: 'Won' }),
      buildGrievanceRow({ grievanceId: 'GRV-003', status: 'Denied' }),
      buildGrievanceRow({ grievanceId: 'GRV-004', status: 'Settled' })
    ];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    // Win rate: 2 won / 4 total outcomes = 50%
    expect(metrics.winRate).toBe('50%');
  });

  test('winRate is "-" when no outcomes', () => {
    const memberData = [new Array(45).fill('')];
    const grievanceData = [
      new Array(40).fill(''), // header
      buildGrievanceRow({ grievanceId: 'GRV-001', status: 'Open' })
    ];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.winRate).toBe('-');
  });

  test('counts overdue cases', () => {
    const memberData = [new Array(45).fill('')];
    const grievanceData = [
      new Array(40).fill(''), // header
      buildGrievanceRow({ grievanceId: 'GRV-001', status: 'Open', daysToDeadline: -5 }),
      buildGrievanceRow({ grievanceId: 'GRV-002', status: 'Open', daysToDeadline: 'Overdue' }),
      buildGrievanceRow({ grievanceId: 'GRV-003', status: 'Open', daysToDeadline: 3 })
    ];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.overdueCases).toBe(2);
    expect(metrics.dueThisWeek).toBe(1);
  });

  test('handles empty data gracefully', () => {
    const memberData = [new Array(45).fill('')]; // header only
    const grievanceData = [new Array(40).fill('')]; // header only

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.totalMembers).toBe(0);
    expect(metrics.activeGrievances).toBe(0);
    expect(metrics.avgDaysOpen).toBe(0);
  });

  test('sixMonthHistory has expected structure', () => {
    const memberData = [new Array(45).fill('')];
    const grievanceData = [new Array(40).fill('')];

    const metrics = computeDashboardMetrics_(memberData, grievanceData, []);
    expect(metrics.sixMonthHistory.casesFiled).toHaveLength(6);
    expect(metrics.sixMonthHistory.grievances).toHaveLength(6);
    expect(metrics.sixMonthHistory.members).toHaveLength(6);
  });
});

// ============================================================================
// getStewardCoverageStats
// ============================================================================

describe('getStewardCoverageStats', () => {
  test('returns zero stats when member sheet does not exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getStewardCoverageStats();
    expect(result.ratio).toBe(0);
    expect(result.stewardCount).toBe(0);
    expect(result.memberCount).toBe(0);
    expect(result.targetRatio).toBe(15);
  });

  test('returns zero stats when sheet has no data', () => {
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Header']]);
    memberSheet.getLastRow.mockReturnValue(1);
    const mockSS = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getStewardCoverageStats();
    expect(result.ratio).toBe(0);
    expect(result.memberCount).toBe(0);
  });
});

// ============================================================================
// getPublicOverviewData
// ============================================================================

describe('getPublicOverviewData', () => {
  test('returns default structure when no sheets exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getPublicOverviewData();
    expect(result.totalMembers).toBe(0);
    expect(result.totalStewards).toBe(0);
    expect(result.totalGrievances).toBe(0);
    expect(result.winRate).toBe(0);
    expect(result.locationBreakdown).toEqual([]);
  });

  test('is defined as a function', () => {
    expect(typeof getPublicOverviewData).toBe('function');
  });
});

// ============================================================================
// syncAllData - smoke test
// ============================================================================

describe('syncAllData', () => {
  test('is defined as a function', () => {
    expect(typeof syncAllData).toBe('function');
  });
});

// ============================================================================
// syncSatisfactionValues - smoke test
// ============================================================================

describe('syncSatisfactionValues', () => {
  test('is defined as a function', () => {
    expect(typeof syncSatisfactionValues).toBe('function');
  });

  test('handles missing satisfaction sheet gracefully', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    expect(() => syncSatisfactionValues()).not.toThrow();
  });
});

// ============================================================================
// computeSectionAverages_ - internal function
// ============================================================================

describe('computeSectionAverages_', () => {
  test('is defined as a function', () => {
    expect(typeof computeSectionAverages_).toBe('function');
  });

  test('returns empty array for empty input', () => {
    const result = computeSectionAverages_([]);
    expect(result).toEqual([]);
  });

  test('skips rows with empty first column (timestamp)', () => {
    const row = new Array(SATISFACTION_COLS.Q62_CONCERNS_SERIOUS).fill('');
    row[SATISFACTION_COLS.TIMESTAMP - 1] = ''; // no timestamp
    const result = computeSectionAverages_([row]);
    expect(result).toEqual([]);
  });

  test('returns 11 section averages per valid row', () => {
    const row = new Array(SATISFACTION_COLS.Q62_CONCERNS_SERIOUS).fill(5); // all 5s
    row[SATISFACTION_COLS.TIMESTAMP - 1] = new Date(); // timestamp
    const result = computeSectionAverages_([row]);
    expect(result).toHaveLength(1);
    expect(result[0]).toHaveLength(11);
  });

  test('computes correct section averages for known values', () => {
    const row = new Array(SATISFACTION_COLS.Q62_CONCERNS_SERIOUS).fill('');
    row[SATISFACTION_COLS.TIMESTAMP - 1] = new Date(); // timestamp

    // Overall Satisfaction (Q6-Q9): set to 4
    [SATISFACTION_COLS.Q6_SATISFIED_REP, SATISFACTION_COLS.Q7_TRUST_UNION,
     SATISFACTION_COLS.Q8_FEEL_PROTECTED, SATISFACTION_COLS.Q9_RECOMMEND].forEach(col => {
      row[col - 1] = 4;
    });

    // Steward Rating (Q10-Q16): set to 3
    [SATISFACTION_COLS.Q10_TIMELY_RESPONSE, SATISFACTION_COLS.Q11_TREATED_RESPECT,
     SATISFACTION_COLS.Q12_EXPLAINED_OPTIONS, SATISFACTION_COLS.Q13_FOLLOWED_THROUGH,
     SATISFACTION_COLS.Q14_ADVOCATED, SATISFACTION_COLS.Q15_SAFE_CONCERNS,
     SATISFACTION_COLS.Q16_CONFIDENTIALITY].forEach(col => {
      row[col - 1] = 3;
    });

    const result = computeSectionAverages_([row]);
    expect(result[0][0]).toBe(4);   // Overall Satisfaction avg
    expect(result[0][1]).toBe(3);   // Steward Rating avg
  });
});

// ============================================================================
// computeSatisfactionRowAverages
// ============================================================================

describe('computeSatisfactionRowAverages', () => {
  test('is defined as a function', () => {
    expect(typeof computeSatisfactionRowAverages).toBe('function');
  });
});

// ============================================================================
// Module-level function existence
// ============================================================================

describe('Dashboard module functions exist', () => {
  test('showSatisfactionDashboard is defined', () => {
    expect(typeof showSatisfactionDashboard).toBe('function');
  });

  test('writeSatisfactionDashboard_ is defined', () => {
    expect(typeof writeSatisfactionDashboard_).toBe('function');
  });

  test('writeDashboardValues_ is defined', () => {
    expect(typeof writeDashboardValues_).toBe('function');
  });
});
