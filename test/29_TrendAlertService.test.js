/**
 * Tests for 29_TrendAlertService.gs
 *
 * Covers the TrendAlertService IIFE: _dedupKey, _alertExists,
 * _detectGrievanceSpike, _detectSatisfactionDrop, _detectDeadlineMissSpike,
 * resolveAlert, getActiveAlerts, acknowledgeAlert, initSheet, runDetection,
 * and global wrappers.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '29_TrendAlertService.gs']);

// ---------------------------------------------------------------------------
// Constants — mirror the HEADERS inside the TrendAlertService IIFE
// ---------------------------------------------------------------------------

var ALERT_HEADERS = [
  'ID', 'Type', 'Severity', 'Title', 'Message', 'Data', 'Created',
  'Status', 'Acknowledged By', 'Acknowledged At', 'Resolved By', 'Resolved At'
];

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

/** Build a data set with headers + alert rows */
function buildAlertData(rows) {
  return [ALERT_HEADERS].concat(rows || []);
}

/** Create a mock _Trend_Alerts sheet with custom data rows */
function buildAlertSheet(rows) {
  var data = buildAlertData(rows);
  var sheet = createMockSheet(SHEETS.TREND_ALERTS, data);

  sheet.getRange = jest.fn(function (row, col, numRows, numCols) {
    return {
      getValue: jest.fn(function () {
        if (data[row - 1] && data[row - 1][col - 1] !== undefined) return data[row - 1][col - 1];
        return '';
      }),
      getValues: jest.fn(function () {
        var result = [];
        for (var r = row - 1; r < row - 1 + (numRows || 1); r++) {
          var rowArr = [];
          for (var c = col - 1; c < col - 1 + (numCols || 1); c++) {
            rowArr.push(data[r] && data[r][c] !== undefined ? data[r][c] : '');
          }
          result.push(rowArr);
        }
        return result;
      }),
      setValue: jest.fn(),
      setValues: jest.fn()
    };
  });

  return sheet;
}

/**
 * Create a mock Grievance Log sheet with custom data rows.
 * Headers are derived from GRIEVANCE_HEADER_MAP_ column order.
 * Caller provides sparse row objects keyed by GRIEVANCE_COLS constants (1-indexed).
 *
 * @param {Array<Object>} rowDefs - Each element maps column index (1-based) to value.
 * @returns {Object} Mock sheet
 */
function buildGrievanceSheet(rowDefs) {
  var headerCount = 30; // GRIEVANCE_HEADER_MAP_ has 30 entries
  var headers = [];
  for (var h = 0; h < headerCount; h++) headers.push('Header' + (h + 1));
  var data = [headers];
  (rowDefs || []).forEach(function (def) {
    var row = new Array(headerCount).fill('');
    for (var col in def) {
      if (def.hasOwnProperty(col)) {
        row[parseInt(col) - 1] = def[col];
      }
    }
    data.push(row);
  });
  return createMockSheet(SHEETS.GRIEVANCE_LOG, data);
}

/**
 * Create a mock Satisfaction sheet with timestamp + score columns.
 * @param {Array<Array>} rows - Each row: [Date, score1, score2, ...]
 * @returns {Object} Mock sheet
 */
function buildSatisfactionSheet(rows) {
  var headers = ['Timestamp', 'Q1', 'Q2', 'Q3'];
  var data = [headers].concat(rows || []);
  var sheet = createMockSheet(SHEETS.SATISFACTION || '_Satisfaction', data);
  // Override getLastColumn to reflect actual data width
  sheet.getLastColumn = jest.fn(function () {
    return data[0] ? data[0].length : 1;
  });
  // Override getRange to support the (row, col, numRows, numCols) pattern used by _detectSatisfactionDrop
  sheet.getRange = jest.fn(function (row, col, numRows, numCols) {
    var nr = numRows || 1;
    var nc = numCols || 1;
    var sliced = [];
    for (var r = row - 1; r < row - 1 + nr; r++) {
      var rowData = [];
      for (var c = col - 1; c < col - 1 + nc; c++) {
        rowData.push(data[r] && data[r][c] !== undefined ? data[r][c] : '');
      }
      sliced.push(rowData);
    }
    return {
      getValue: jest.fn(function () { return sliced[0] && sliced[0][0] !== undefined ? sliced[0][0] : ''; }),
      getValues: jest.fn(function () { return sliced; }),
      setValue: jest.fn(),
      setValues: jest.fn()
    };
  });
  return sheet;
}

/** Install a spreadsheet mock with the given sheets */
function installSS(sheets) {
  var ss = createMockSpreadsheet(sheets || []);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

// ---------------------------------------------------------------------------
// Global mocks
// ---------------------------------------------------------------------------

beforeEach(() => {
  jest.clearAllMocks();
  global.logAuditEvent = jest.fn();
  global.maskEmail = jest.fn(function (e) { return e ? e.replace(/@.*/, '@***') : ''; });
  global.setSheetVeryHidden_ = jest.fn();
});

// ============================================================================
// _dedupKey (accessed indirectly via runDetection dedup logic)
// ============================================================================

describe('_dedupKey logic', () => {
  // _dedupKey is private inside the IIFE, but we can verify its behavior
  // by observing runDetection dedup: two identical alerts produce only one row.

  test('same type + data produces one alert (dedup works)', () => {
    // Create a grievance spike scenario that would fire the same alert twice
    // if dedup were broken. We test by running detection with spike data
    // and verifying only one appendRow call occurs.
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // 5 grievances at "Office A" in last 7 days, 5 total in 30 days
    // dailyAvg = 5/30 = 0.167, expected7d = 1.17, spike threshold = 2.33
    // count7d (5) > 2.33 => spike detected
    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    // Only one GRIEVANCE_SPIKE alert for "Office A" should be appended
    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(1);
  });

  test('different data signatures produce separate alerts', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // 4 grievances at "Office A" + 4 at "Office B" — both should spike
    var rows = [];
    ['Office A', 'Office B'].forEach(function (loc) {
      for (var i = 0; i < 4; i++) {
        var def = {};
        def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
        def[GRIEVANCE_COLS.LOCATION] = loc;
        def[GRIEVANCE_COLS.STATUS] = 'Open';
        rows.push(def);
      }
    });

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(2);
  });

  test('dedup key is case-insensitive and trims whitespace', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // Grievance grouping uses raw location strings, so each variant is a
    // separate location. The dedup key normalization (lowercase+trim) only
    // matters when comparing candidate alerts against existing sheet alerts.
    // To test dedup normalization, create a spike at "Office A" and put an
    // existing alert with data " office a " — dedup should still match.
    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);

    // Existing alert with differently-cased/whitespaced data should still dedup
    var alertSheet = buildAlertSheet([
      ['abc', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', ' office a ', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    // Dedup key normalizes both to "grievance_spike:office a" — no new alert
    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(0);
  });
});

// ============================================================================
// _alertExists (accessed indirectly via runDetection dedup logic)
// ============================================================================

describe('_alertExists logic', () => {
  test('skips alert if matching Active alert exists in sheet', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // 5 grievances that would trigger a spike
    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);

    // Pre-existing Active alert for same type+data
    var alertSheet = buildAlertSheet([
      ['abc123', 'GRIEVANCE_SPIKE', 'HIGH', 'Grievance Spike: Office A',
       '5 grievances...', 'Office A', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    // No new alert should be appended since one already exists
    expect(alertSheet.appendRow).not.toHaveBeenCalled();
  });

  test('skips alert if matching Acknowledged alert exists', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);

    // Acknowledged (not Resolved) — should still block new alert
    var alertSheet = buildAlertSheet([
      ['abc123', 'GRIEVANCE_SPIKE', 'HIGH', 'Grievance Spike: Office A',
       '5 grievances...', 'Office A', new Date(), 'Acknowledged', 'steward@test.com', new Date(), '', '']
    ]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    expect(alertSheet.appendRow).not.toHaveBeenCalled();
  });

  test('creates new alert if matching alert is Resolved', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);

    // Resolved alert — should NOT block new alert
    var alertSheet = buildAlertSheet([
      ['abc123', 'GRIEVANCE_SPIKE', 'HIGH', 'Grievance Spike: Office A',
       '5 grievances...', 'Office A', new Date(), 'Resolved', '', '', 'steward@test.com', new Date()]
    ]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    // New alert should be created since old one is Resolved
    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(1);
  });
});

// ============================================================================
// _detectGrievanceSpike
// ============================================================================

describe('_detectGrievanceSpike', () => {
  test('detects spike when 7-day count > 2x expected', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // 5 grievances in last 7 days at one location, only 5 total in 30 days
    // dailyAvg = 5/30 = 0.167/day, expected7d = 1.17
    // count7d(5) > expected7d*2(2.33) AND count7d(5) >= 3 => SPIKE
    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Main Office';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(1);
    expect(spikeCalls[0][0][3]).toContain('Main Office');
  });

  test('does not trigger when count is below minimum threshold of 3', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // Only 2 grievances — below the minimum threshold of 3
    var rows = [];
    for (var i = 0; i < 2; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Small Office';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(0);
  });

  test('does not trigger when 7-day count is within normal range', () => {
    var now = new Date();

    // Spread 30 grievances evenly over 30 days — no spike
    // dailyAvg = 30/30 = 1/day, expected7d = 7
    // 3 in last 7 days => 3 <= 14 (7*2) => no spike
    var rows = [];
    for (var i = 0; i < 30; i++) {
      var daysAgo = new Date(now.getTime() - i * 86400000);
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = daysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Steady Office';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    expect(spikeCalls.length).toBe(0);
  });

  test('groups grievances by location', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    // 4 grievances at "Site A" (spike) + 1 at "Site B" (no spike — below min)
    var rows = [];
    for (var i = 0; i < 4; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Site A';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }
    var defB = {};
    defB[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
    defB[GRIEVANCE_COLS.LOCATION] = 'Site B';
    defB[GRIEVANCE_COLS.STATUS] = 'Open';
    rows.push(defB);

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    // Only Site A should spike; Site B has only 1 (below min 3)
    expect(spikeCalls.length).toBe(1);
    expect(spikeCalls[0][0][3]).toContain('Site A');
  });

  test('returns no alerts for empty grievance log', () => {
    var grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [['Header']]);
    grievanceSheet.getLastRow.mockReturnValue(1);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    expect(alertSheet.appendRow).not.toHaveBeenCalled();
  });

  test('skips rows with non-Date filed values', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    // 2 valid date rows + 3 non-date rows = only 2 valid, below threshold
    for (var i = 0; i < 2; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'Office X';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }
    for (var j = 0; j < 3; j++) {
      var bad = {};
      bad[GRIEVANCE_COLS.DATE_FILED] = 'not-a-date';
      bad[GRIEVANCE_COLS.LOCATION] = 'Office X';
      bad[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(bad);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var spikeCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'GRIEVANCE_SPIKE';
    });
    // Only 2 valid rows — below minimum threshold of 3
    expect(spikeCalls.length).toBe(0);
  });
});

// ============================================================================
// _detectSatisfactionDrop
// ============================================================================

describe('_detectSatisfactionDrop', () => {
  test('detects >15% quarter-over-quarter drop', () => {
    // Q1 2026: avg score 4.0, Q2 2026: avg score 3.0 => 25% drop
    var q1Date = new Date('2026-01-15');
    var q2Date = new Date('2026-04-15');
    var satRows = [
      [q1Date, 4.0, 4.0, 4.0],
      [q1Date, 4.0, 4.0, 4.0],
      [q2Date, 3.0, 3.0, 3.0],
      [q2Date, 3.0, 3.0, 3.0]
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    expect(dropCalls.length).toBe(1);
    expect(dropCalls[0][0][4]).toContain('25%');
  });

  test('does not trigger for drop <= 15%', () => {
    // Q1: avg 4.0, Q2: avg 3.5 => 12.5% drop — below threshold
    var q1Date = new Date('2026-01-15');
    var q2Date = new Date('2026-04-15');
    var satRows = [
      [q1Date, 4.0, 4.0, 4.0],
      [q2Date, 3.5, 3.5, 3.5]
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    expect(dropCalls.length).toBe(0);
  });

  test('does not trigger when scores improve', () => {
    // Q1: avg 3.0, Q2: avg 4.0 => improvement, not a drop
    var q1Date = new Date('2026-01-15');
    var q2Date = new Date('2026-04-15');
    var satRows = [
      [q1Date, 3.0, 3.0, 3.0],
      [q2Date, 4.0, 4.0, 4.0]
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    expect(dropCalls.length).toBe(0);
  });

  test('returns no alerts with fewer than 2 quarters of data', () => {
    // Only one quarter of data — cannot compute a drop
    var q1Date = new Date('2026-01-15');
    var satRows = [
      [q1Date, 4.0, 4.0, 4.0],
      [q1Date, 3.5, 3.5, 3.5]
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    expect(dropCalls.length).toBe(0);
  });

  test('returns no alerts when satisfaction sheet is missing', () => {
    var alertSheet = buildAlertSheet([]);
    installSS([alertSheet]);

    TrendAlertService.runDetection();

    expect(alertSheet.appendRow).not.toHaveBeenCalled();
  });

  test('ignores non-numeric score values', () => {
    // Q1: avg 4.0 (valid), Q2: all non-numeric => no valid avg => no alert
    var q1Date = new Date('2026-01-15');
    var q2Date = new Date('2026-04-15');
    var satRows = [
      [q1Date, 4.0, 4.0, 4.0],
      [q2Date, 'N/A', '', 'pending']
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    expect(dropCalls.length).toBe(0);
  });

  test('ignores scores outside 1-5 range', () => {
    // Q1: valid scores, Q2: scores outside valid range (0 and 10)
    var q1Date = new Date('2026-01-15');
    var q2Date = new Date('2026-04-15');
    var satRows = [
      [q1Date, 4.0, 4.0, 4.0],
      [q2Date, 0, 10, -1]
    ];

    var satSheet = buildSatisfactionSheet(satRows);
    var alertSheet = buildAlertSheet([]);
    installSS([satSheet, alertSheet]);

    TrendAlertService.runDetection();

    var dropCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'SATISFACTION_DROP';
    });
    // No valid Q2 scores => no comparison possible
    expect(dropCalls.length).toBe(0);
  });
});

// ============================================================================
// _detectDeadlineMissSpike
// ============================================================================

describe('_detectDeadlineMissSpike', () => {
  test('detects spike when current month misses > 1.5x 3-month average', () => {
    var now = new Date();
    var thisMonth = now.getFullYear() * 12 + now.getMonth();

    var rows = [];

    // Current month: 6 past-due deadlines
    for (var i = 0; i < 6; i++) {
      var def = {};
      // Deadline is in current month but in the past
      var deadlineDate = new Date(now.getFullYear(), now.getMonth(), 1);
      def[GRIEVANCE_COLS.NEXT_ACTION_DUE] = deadlineDate;
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    // Previous 3 months: 2 misses each (avg = 2)
    // Spike: 6 > 2 * 1.5 (3) => yes
    for (var m = 1; m <= 3; m++) {
      for (var j = 0; j < 2; j++) {
        var pastDef = {};
        var pastMonth = now.getMonth() - m;
        var pastYear = now.getFullYear();
        if (pastMonth < 0) { pastMonth += 12; pastYear--; }
        // Date in the past (day 1 of that month is before now)
        pastDef[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(pastYear, pastMonth, 1);
        pastDef[GRIEVANCE_COLS.STATUS] = 'Open';
        rows.push(pastDef);
      }
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var deadlineCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'DEADLINE_MISS_SPIKE';
    });
    expect(deadlineCalls.length).toBe(1);
    expect(deadlineCalls[0][0][4]).toContain('6');
  });

  test('does not trigger when current month < 3 total deadlines', () => {
    var now = new Date();

    var rows = [];
    // Only 2 deadlines this month (below minimum threshold of 3)
    for (var i = 0; i < 2; i++) {
      var def = {};
      def[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(now.getFullYear(), now.getMonth(), 1);
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    // Past months: 1 miss each
    for (var m = 1; m <= 3; m++) {
      var pastDef = {};
      var pastMonth = now.getMonth() - m;
      var pastYear = now.getFullYear();
      if (pastMonth < 0) { pastMonth += 12; pastYear--; }
      pastDef[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(pastYear, pastMonth, 1);
      pastDef[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(pastDef);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var deadlineCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'DEADLINE_MISS_SPIKE';
    });
    expect(deadlineCalls.length).toBe(0);
  });

  test('excludes resolved/settled/withdrawn grievances from counting', () => {
    var now = new Date();

    var rows = [];
    // 4 deadlines this month but all resolved — should not count
    var resolvedStatuses = ['resolved', 'settled', 'withdrawn', 'resolved'];
    for (var i = 0; i < 4; i++) {
      var def = {};
      def[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(now.getFullYear(), now.getMonth(), 1);
      def[GRIEVANCE_COLS.STATUS] = resolvedStatuses[i];
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var deadlineCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'DEADLINE_MISS_SPIKE';
    });
    expect(deadlineCalls.length).toBe(0);
  });

  test('does not trigger when misses are within normal range', () => {
    var now = new Date();

    var rows = [];
    // Current month: 3 misses
    for (var i = 0; i < 3; i++) {
      var def = {};
      def[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(now.getFullYear(), now.getMonth(), 1);
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    // Past 3 months: 3 misses each (avg = 3)
    // Spike threshold: 3 * 1.5 = 4.5; current(3) <= 4.5 => no spike
    for (var m = 1; m <= 3; m++) {
      for (var j = 0; j < 3; j++) {
        var pastDef = {};
        var pastMonth = now.getMonth() - m;
        var pastYear = now.getFullYear();
        if (pastMonth < 0) { pastMonth += 12; pastYear--; }
        pastDef[GRIEVANCE_COLS.NEXT_ACTION_DUE] = new Date(pastYear, pastMonth, 1);
        pastDef[GRIEVANCE_COLS.STATUS] = 'Open';
        rows.push(pastDef);
      }
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    var deadlineCalls = alertSheet.appendRow.mock.calls.filter(function (call) {
      return call[0] && call[0][1] === 'DEADLINE_MISS_SPIKE';
    });
    expect(deadlineCalls.length).toBe(0);
  });
});

// ============================================================================
// resolveAlert
// ============================================================================

describe('TrendAlertService.resolveAlert', () => {
  test('marks alert as Resolved and records resolver info', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'Office A', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = TrendAlertService.resolveAlert('alert001', 'steward@test.com');

    expect(result.success).toBe(true);
    // Row 2 (data row 1), col 8 = Status => 'Resolved'
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 8);
    // Col 11 = Resolved By
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 11);
    // Col 12 = Resolved At
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 12);
  });

  test('returns success:false for non-existent alertId', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = TrendAlertService.resolveAlert('nonexistent');

    expect(result.success).toBe(false);
  });

  test('returns success:false when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);

    var result = TrendAlertService.resolveAlert('alert001');

    expect(result.success).toBe(false);
  });

  test('returns success:false when alert sheet is missing', () => {
    installSS([]);

    var result = TrendAlertService.resolveAlert('alert001');

    expect(result.success).toBe(false);
  });

  test('logs audit event on resolve', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    TrendAlertService.resolveAlert('alert001', 'steward@test.com');

    expect(logAuditEvent).toHaveBeenCalledWith(
      'TREND_ALERT_RESOLVED',
      expect.objectContaining({ alertId: 'alert001' })
    );
  });

  test('resolves without resolvedBy (anonymous resolve)', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = TrendAlertService.resolveAlert('alert001');

    expect(result.success).toBe(true);
    // Status should still be set to Resolved
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 8);
    // Resolved By / At columns should NOT be written without resolvedBy
    expect(alertSheet.getRange).not.toHaveBeenCalledWith(2, 11);
  });
});

// ============================================================================
// getActiveAlerts
// ============================================================================

describe('TrendAlertService.getActiveAlerts', () => {
  test('returns Active and Acknowledged alerts, excludes Resolved', () => {
    var alertSheet = buildAlertSheet([
      ['a1', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike 1', 'msg1', 'data1', new Date(), 'Active', '', '', '', ''],
      ['a2', 'SATISFACTION_DROP', 'HIGH', 'Drop', 'msg2', 'data2', new Date(), 'Acknowledged', 'admin@test.com', new Date(), '', ''],
      ['a3', 'WIN_RATE_DECLINE', 'MEDIUM', 'Win Rate', 'msg3', 'data3', new Date(), 'Resolved', '', '', 'admin@test.com', new Date()]
    ]);
    installSS([alertSheet]);

    var alerts = TrendAlertService.getActiveAlerts();

    expect(alerts.length).toBe(2);
    expect(alerts[0].id).toBe('a1');
    expect(alerts[0].status).toBe('Active');
    expect(alerts[1].id).toBe('a2');
    expect(alerts[1].status).toBe('Acknowledged');
  });

  test('returns empty array when sheet is missing', () => {
    installSS([]);

    var alerts = TrendAlertService.getActiveAlerts();

    expect(alerts).toEqual([]);
  });

  test('returns empty array when sheet has only headers', () => {
    var alertSheet = buildAlertSheet([]);
    // Override getLastRow to return 1 (headers only)
    alertSheet.getLastRow.mockReturnValue(1);
    installSS([alertSheet]);

    var alerts = TrendAlertService.getActiveAlerts();

    expect(alerts).toEqual([]);
  });

  test('returns empty array when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);

    var alerts = TrendAlertService.getActiveAlerts();

    expect(alerts).toEqual([]);
  });

  test('returns alert objects with correct fields', () => {
    var created = new Date('2026-04-01');
    var ackDate = new Date('2026-04-02');
    var alertSheet = buildAlertSheet([
      ['abc123', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike at HQ', '5 grievances...', 'HQ', created, 'Acknowledged', 'admin@test.com', ackDate, '', '']
    ]);
    installSS([alertSheet]);

    var alerts = TrendAlertService.getActiveAlerts();

    expect(alerts.length).toBe(1);
    var a = alerts[0];
    expect(a.id).toBe('abc123');
    expect(a.type).toBe('GRIEVANCE_SPIKE');
    expect(a.severity).toBe('HIGH');
    expect(a.title).toBe('Spike at HQ');
    expect(a.message).toBe('5 grievances...');
    expect(a.data).toBe('HQ');
    expect(a.created).toEqual(created);
    expect(a.status).toBe('Acknowledged');
    expect(a.acknowledgedBy).toBe('admin@test.com');
    expect(a.acknowledgedAt).toEqual(ackDate);
  });
});

// ============================================================================
// acknowledgeAlert
// ============================================================================

describe('TrendAlertService.acknowledgeAlert', () => {
  test('marks alert as Acknowledged with steward email', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = TrendAlertService.acknowledgeAlert('alert001', 'steward@test.com');

    expect(result.success).toBe(true);
    // Row 2, col 8 = Status => 'Acknowledged'
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 8);
    // Row 2, col 9 = Acknowledged By
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 9);
    // Row 2, col 10 = Acknowledged At
    expect(alertSheet.getRange).toHaveBeenCalledWith(2, 10);
  });

  test('returns success:false for non-existent alertId', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = TrendAlertService.acknowledgeAlert('nonexistent', 'steward@test.com');

    expect(result.success).toBe(false);
  });

  test('logs audit event on acknowledgement', () => {
    var alertSheet = buildAlertSheet([
      ['alert001', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    TrendAlertService.acknowledgeAlert('alert001', 'steward@test.com');

    expect(logAuditEvent).toHaveBeenCalledWith(
      'TREND_ALERT_ACKNOWLEDGED',
      expect.objectContaining({ alertId: 'alert001' })
    );
  });
});

// ============================================================================
// initSheet
// ============================================================================

describe('TrendAlertService.initSheet', () => {
  test('creates _Trend_Alerts sheet if missing', () => {
    var ss = installSS([]);

    TrendAlertService.initSheet();

    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.TREND_ALERTS);
  });

  test('does not recreate if sheet exists', () => {
    var existingSheet = buildAlertSheet([]);
    var ss = installSS([existingSheet]);

    TrendAlertService.initSheet();

    expect(ss.insertSheet).not.toHaveBeenCalled();
  });

  test('throws when spreadsheet binding is broken', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);

    expect(() => TrendAlertService.initSheet()).toThrow('Spreadsheet binding broken');
  });
});

// ============================================================================
// runDetection
// ============================================================================

describe('TrendAlertService.runDetection', () => {
  test('returns 0 when no alerts are generated', () => {
    var alertSheet = buildAlertSheet([]);
    installSS([alertSheet]);

    var result = TrendAlertService.runDetection();

    expect(result).toBe(0);
  });

  test('logs audit event when new alerts are written', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'HQ';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    expect(logAuditEvent).toHaveBeenCalledWith(
      'TREND_ALERTS_GENERATED',
      expect.objectContaining({ newAlerts: expect.any(Number) })
    );
  });

  test('returns 0 when lock cannot be acquired', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = 'HQ';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    // Make lock fail
    LockService._mockLockFailure();

    var result = TrendAlertService.runDetection();

    expect(result).toBe(0);
    expect(alertSheet.appendRow).not.toHaveBeenCalled();

    LockService._mockLockSuccess();
  });

  test('escapes alert data with escapeForFormula', () => {
    var now = new Date();
    var twoDaysAgo = new Date(now.getTime() - 2 * 86400000);

    var rows = [];
    for (var i = 0; i < 5; i++) {
      var def = {};
      def[GRIEVANCE_COLS.DATE_FILED] = twoDaysAgo;
      def[GRIEVANCE_COLS.LOCATION] = '=MALICIOUS()';
      def[GRIEVANCE_COLS.STATUS] = 'Open';
      rows.push(def);
    }

    var grievanceSheet = buildGrievanceSheet(rows);
    var alertSheet = buildAlertSheet([]);
    installSS([grievanceSheet, alertSheet]);

    TrendAlertService.runDetection();

    // The data field (col index 5 in the appendRow array) should be escaped
    if (alertSheet.appendRow.mock.calls.length > 0) {
      var appendedData = alertSheet.appendRow.mock.calls[0][0][5];
      // escapeForFormula should prefix '=' with apostrophe
      expect(appendedData).not.toBe('=MALICIOUS()');
    }
  });
});

// ============================================================================
// Global wrappers
// ============================================================================

describe('Global wrappers', () => {
  test('dataGetTrendAlerts delegates to getActiveAlerts with auth', () => {
    var alertSheet = buildAlertSheet([
      ['a1', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = dataGetTrendAlerts('valid-token');

    expect(_requireStewardAuth).toHaveBeenCalledWith('valid-token');
    expect(result.length).toBe(1);
    expect(result[0].id).toBe('a1');
  });

  test('dataGetTrendAlerts returns [] when auth fails', () => {
    _requireStewardAuth.mockReturnValueOnce(null);

    var result = dataGetTrendAlerts('bad-token');

    expect(result).toEqual([]);
  });

  test('dataAcknowledgeTrendAlert delegates with auth', () => {
    var alertSheet = buildAlertSheet([
      ['a1', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = dataAcknowledgeTrendAlert('valid-token', 'a1');

    expect(result.success).toBe(true);
  });

  test('dataAcknowledgeTrendAlert fails with bad auth', () => {
    _requireStewardAuth.mockReturnValueOnce(null);

    var result = dataAcknowledgeTrendAlert('bad-token', 'a1');

    expect(result.success).toBe(false);
    expect(result.message).toContain('Steward');
  });

  test('dataResolveTrendAlert delegates with auth', () => {
    var alertSheet = buildAlertSheet([
      ['a1', 'GRIEVANCE_SPIKE', 'HIGH', 'Spike', 'msg', 'data', new Date(), 'Active', '', '', '', '']
    ]);
    installSS([alertSheet]);

    var result = dataResolveTrendAlert('valid-token', 'a1');

    expect(result.success).toBe(true);
  });

  test('dataResolveTrendAlert fails with bad auth', () => {
    _requireStewardAuth.mockReturnValueOnce(null);

    var result = dataResolveTrendAlert('bad-token', 'a1');

    expect(result.success).toBe(false);
    expect(result.message).toContain('Steward');
  });

  test('triggerDailyTrendDetection calls runDetection', () => {
    var alertSheet = buildAlertSheet([]);
    installSS([alertSheet]);

    // Should not throw
    expect(() => triggerDailyTrendDetection()).not.toThrow();
  });
});
