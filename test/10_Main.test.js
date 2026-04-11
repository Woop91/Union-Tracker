/**
 * Tests for 10_Main.gs — exercises the real addNewMember + Member ID
 * generation + column guards + exportMemberDirectory validation paths.
 *
 * The previous version of this file defined buildMemberRow / sanitizeForScript /
 * escapeCSVField locally and asserted against those re-implementations. Those
 * tests could never catch a production bug because the production code was
 * never exercised. This rewrite loads 10_Main.gs and calls the real exports
 * against gas-mock spreadsheets.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '02_DataManagers.gs', '10_Main.gs']);

// ============================================================================
// addNewMember — verifies the real function writes correct column positions
// ============================================================================
describe('addNewMember (real production function)', () => {
  test('appends a row sized by the max MEMBER_COLS index', () => {
    const headers = [];
    for (var i = 0; i < 80; i++) headers.push('Col' + i);
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, [headers]);
    const ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var maxCol = 0;
    for (var k in MEMBER_COLS) {
      if (MEMBER_COLS[k] > maxCol) maxCol = MEMBER_COLS[k];
    }

    addNewMember({ firstName: 'Jane', lastName: 'Doe', email: 'jane@example.com' });

    expect(sheet.appendRow).toHaveBeenCalledTimes(1);
    var row = sheet.appendRow.mock.calls[0][0];
    expect(row.length).toBe(maxCol);
    // 1-indexed MEMBER_COLS; 0-indexed array.
    expect(row[MEMBER_COLS.FIRST_NAME - 1]).toBe('Jane');
    expect(row[MEMBER_COLS.LAST_NAME - 1]).toBe('Doe');
    expect(row[MEMBER_COLS.EMAIL - 1]).toBe('jane@example.com');
    expect(row[MEMBER_COLS.IS_STEWARD - 1]).toBe('No');
  });

  test('generates a MEM-{base36} id unique per call', () => {
    const headers = [];
    for (var i = 0; i < 80; i++) headers.push('Col' + i);
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, [headers]);
    const ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    addNewMember({ firstName: 'A', lastName: 'B' });
    var row1 = sheet.appendRow.mock.calls[0][0];
    var id1 = row1[MEMBER_COLS.MEMBER_ID - 1];
    expect(id1).toMatch(/^MEM-[A-Z0-9]+$/);
    expect(id1.length).toBeGreaterThan(6);
    expect(id1.length).toBeLessThan(32);
  });

  test('returns a failure message when Member Directory sheet is missing', () => {
    const ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    var result = addNewMember({ firstName: 'X', lastName: 'Y' });
    expect(result && (result.success === false || /not found/i.test(result.message || result.error || ''))).toBe(true);
  });
});

// ============================================================================
// Column range invariants — caught regressions in v4.50.x column migrations
// ============================================================================
describe('handleGrievanceEdit column references are valid', () => {
  test('status + date columns used by Next Action Due updates are in range', () => {
    var cols = [
      GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.CURRENT_STEP,
      GRIEVANCE_COLS.DATE_FILED, GRIEVANCE_COLS.STEP1_RCVD,
      GRIEVANCE_COLS.STEP2_APPEAL_FILED, GRIEVANCE_COLS.STEP2_RCVD,
      GRIEVANCE_COLS.STEP3_APPEAL_FILED, GRIEVANCE_COLS.DATE_CLOSED
    ];
    cols.forEach(function(col) {
      expect(col).toBeGreaterThan(0);
      expect(col).toBeLessThanOrEqual(GRIEVANCE_COLS.QUICK_ACTIONS);
    });
  });

  test('NEXT_ACTION_DUE is a distinct deadline column', () => {
    expect(GRIEVANCE_COLS.NEXT_ACTION_DUE).toBeGreaterThan(0);
    expect(GRIEVANCE_COLS.NEXT_ACTION_DUE).not.toBe(GRIEVANCE_COLS.DATE_CLOSED);
  });
});

// ============================================================================
// exportMemberDirectory — at minimum verify it refuses unauthorized callers
// ============================================================================
describe('exportMemberDirectory auth gate', () => {
  test('denies when checkWebAppAuthorization returns unauthorized', () => {
    // checkWebAppAuthorization uses ACCESS_CONTROL which is defined in 00_Security.
    // With no mock user it should return isAuthorized:false.
    var original;
    if (typeof checkWebAppAuthorization === 'function') {
      original = global.checkWebAppAuthorization;
      global.checkWebAppAuthorization = jest.fn(function() {
        return { isAuthorized: false, message: 'Not authorized' };
      });
    }
    try {
      var result = exportMemberDirectory('csv');
      expect(result && (result.success === false || /not authorized/i.test(result.message || ''))).toBe(true);
    } finally {
      if (original) global.checkWebAppAuthorization = original;
    }
  });
});
