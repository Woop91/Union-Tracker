// Tests functions from: 04d_ExecutiveDashboard.gs, 09_Dashboards.gs
/**
 * Tests for getExecutiveMetrics_ and friends — now exercises the real
 * production function against seeded mock grievance data instead of testing
 * a test-local regex copy that the production code doesn't even use.
 *
 * The previous v4.31.0 "M7: Step matching word-boundary regex" block was
 * deleted (v4.55.2). The production getExecutiveMetrics_ uses === against
 * GRIEVANCE_STATUS constants, not regex patterns, so the old tests could
 * not catch any regression in the real function.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '09_Dashboards.gs'
]);

describe('04d function existence', () => {
  const required = [
    'getDashboardStats',
    'getExecutiveMetrics_', 'showStewardDashboard',
    'checkDashboardAlerts',
    'createAutomationTriggers', 'setupMidnightTrigger',
    'removeMidnightTrigger', 'midnightAutoRefresh',
    'checkOverdueGrievances_', 'emailExecutivePDF'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('getDashboardStats', () => {
  test('returns defined result', () => {
    const result = getDashboardStats();
    expect(result).toBeDefined();
    const parsed = typeof result === 'string' ? JSON.parse(result) : result;
    expect(typeof parsed).toBe('object');
  });
});

describe('getExecutiveMetrics_ (real production function)', () => {
  // Helper: build a grievance-log sheet with one row per status.
  function seedGrievances(rows) {
    var headerCount = 0;
    for (var k in GRIEVANCE_COLS) {
      if (GRIEVANCE_COLS[k] > headerCount) headerCount = GRIEVANCE_COLS[k];
    }
    var header = new Array(headerCount).fill('Col');
    header[GRIEVANCE_COLS.STATUS - 1] = 'Status';
    header[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 'Days to Deadline';
    var sheetData = [header];
    rows.forEach(function(r) {
      var row = new Array(headerCount).fill('');
      row[GRIEVANCE_COLS.STATUS - 1] = r.status;
      if (r.daysToDeadline !== undefined) row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = r.daysToDeadline;
      sheetData.push(row);
    });
    var grievSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, sheetData);
    var ss = createMockSpreadsheet([grievSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    return { sheet: grievSheet, ss: ss };
  }

  test('counts OPEN and PENDING as active grievances', () => {
    seedGrievances([
      { status: GRIEVANCE_STATUS.OPEN },
      { status: GRIEVANCE_STATUS.OPEN },
      { status: GRIEVANCE_STATUS.PENDING },
      { status: GRIEVANCE_STATUS.SETTLED }
    ]);
    var metrics = getExecutiveMetrics_();
    expect(metrics.activeGrievances).toBe(3);
  });

  test('computes winRate from WON over total closed', () => {
    seedGrievances([
      { status: GRIEVANCE_STATUS.WON },
      { status: GRIEVANCE_STATUS.WON },
      { status: GRIEVANCE_STATUS.SETTLED },
      { status: GRIEVANCE_STATUS.DENIED }
    ]);
    var metrics = getExecutiveMetrics_();
    // 2 won out of 4 closed = 50%
    expect(metrics.winRate).toBe(50);
  });

  test('counts overdue grievances from negative daysToDeadline', () => {
    seedGrievances([
      { status: GRIEVANCE_STATUS.OPEN, daysToDeadline: -3 },
      { status: GRIEVANCE_STATUS.OPEN, daysToDeadline: 'Overdue' },
      { status: GRIEVANCE_STATUS.OPEN, daysToDeadline: 7 }
    ]);
    var metrics = getExecutiveMetrics_();
    expect(metrics.overdueSteps).toBe(2);
  });

  test('returns defaults when no grievance data is available', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    var metrics = getExecutiveMetrics_();
    expect(metrics.activeGrievances).toBe(0);
    expect(metrics.winRate).toBe(0);
    expect(metrics.overdueSteps).toBe(0);
  });
});
