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
// getVaultDataMap_ lives in 08d_AuditAndFormulas.gs (not loaded here) — default to empty map
global.getVaultDataMap_ = jest.fn(() => ({}));
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
  '08c_FormsAndNotifications.gs',
  '09_Dashboards.gs'
]);

// Mock getSatisfactionColMap_ to return static positions matching SATISFACTION_COLS
// (the dynamic version reads sheet headers which don't exist in test mocks)
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
    // '5' is a numeric string — computeAverage_ treats it as a number and includes it
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

// ============================================================================
// approveFlaggedSubmission / rejectFlaggedSubmission
// Regression tests for formula injection fix: callerEmail must be wrapped
// with escapeForFormula() before writing to REVIEWER_NOTES column.
// ============================================================================

describe('approveFlaggedSubmission / rejectFlaggedSubmission — formula injection protection', () => {
  // Vault column order (1-indexed, matches SURVEY_VAULT_HEADER_MAP_):
  // 1=Response Row, 2=Email Hash, 3=Verified, 4=Member ID Hash,
  // 5=Quarter, 6=Is Latest, 7=Superseded By, 8=Reviewer Notes
  const VAULT_HEADERS = ['Response Row', 'Email Hash', 'Verified', 'Member ID Hash', 'Quarter', 'Is Latest', 'Superseded By', 'Reviewer Notes'];

  var _origGetUserRole;
  var _origSyncSatisfaction;

  beforeEach(() => {
    _origGetUserRole = global.getUserRole_;
    global.getUserRole_ = jest.fn(() => 'admin');
    _origSyncSatisfaction = global.syncSatisfactionValues;
    global.syncSatisfactionValues = jest.fn(); // prevent cascade
  });

  afterEach(() => {
    global.getUserRole_ = _origGetUserRole;
    global.syncSatisfactionValues = _origSyncSatisfaction;
  });

  function makeVaultSs(extraRows) {
    var data = [VAULT_HEADERS, ...(extraRows || [])];
    var setValueSpy = jest.fn();
    var vaultSheet = createMockSheet('_Survey_Vault', data);
    // Override getRange to capture setValue calls
    vaultSheet.getRange = jest.fn(() => ({ setValue: setValueSpy, getValues: jest.fn(() => data) }));
    var ss = createMockSpreadsheet([vaultSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);
    return setValueSpy;
  }

  test('approve: escapeForFormula called on reviewer notes (newline injection blocked)', () => {
    // A crafted email with \n= would inject a formula on the next cell line
    // without escapeForFormula. escapeForFormula replaces \n with a space.
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'steward@org.com\n=MALICIOUS()') }));
    var setValueSpy = makeVaultSs([[1, 'hash', false, 'memhash', 'Q1', true, '', '']]);

    approveFlaggedSubmission(1);

    var reviewerNotesCall = setValueSpy.mock.calls.find(c => typeof c[0] === 'string' && c[0].includes('Approved by'));
    expect(reviewerNotesCall).toBeDefined();
    // escapeForFormula strips newlines — no raw \n in output
    expect(reviewerNotesCall[0]).not.toContain('\n');
    expect(reviewerNotesCall[0]).not.toMatch(/^=/);
  });

  test('approve: writes "Yes" to VERIFIED column', () => {
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'admin@org.com') }));
    var setValueSpy = makeVaultSs([[1, 'hash', false, 'memhash', 'Q1', true, '', '']]);

    approveFlaggedSubmission(1);

    expect(setValueSpy).toHaveBeenCalledWith('Yes');
  });

  test('reject: escapeForFormula called on reviewer notes', () => {
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'steward@org.com\n=MALICIOUS()') }));
    var setValueSpy = makeVaultSs([[2, 'hash', false, 'memhash', 'Q1', true, '', '']]);

    rejectFlaggedSubmission(2);

    var reviewerNotesCall = setValueSpy.mock.calls.find(c => typeof c[0] === 'string' && c[0].includes('Rejected by'));
    expect(reviewerNotesCall).toBeDefined();
    expect(reviewerNotesCall[0]).not.toContain('\n');
  });

  test('reject: writes "Rejected" to VERIFIED and "No" to IS_LATEST', () => {
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'admin@org.com') }));
    var setValueSpy = makeVaultSs([[3, 'hash', false, 'memhash', 'Q1', true, '', '']]);

    rejectFlaggedSubmission(3);

    expect(setValueSpy).toHaveBeenCalledWith('Rejected');
    expect(setValueSpy).toHaveBeenCalledWith('No');
  });

  test('approve: throws when caller has no role (auth check)', () => {
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'nobody@org.com') }));
    global.getUserRole_ = jest.fn(() => 'member');
    makeVaultSs([[1, 'hash', false, 'memhash', 'Q1', true, '', '']]);

    expect(() => approveFlaggedSubmission(1)).toThrow('Access denied');
  });
});

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
