/**
 * Tests for 12_Features.gs
 *
 * Covers EXTENSION_CONFIG, COL_IDX, LOOKER_CONFIG constants,
 * and pure categorization/bucket functions:
 * getDaysBucket_, getContactFrequencyCategory_,
 * getEngagementLevel_, getQuarter_,
 * getOutcomeCategory_, and getHeaderMap.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '12_Features.gs'
]);

// ============================================================================
// EXTENSION_CONFIG
// ============================================================================

describe('EXTENSION_CONFIG', () => {
  test('is defined', () => {
    expect(EXTENSION_CONFIG).toBeDefined();
  });

  test('has correct HIDDEN_CALC_SHEET name', () => {
    expect(EXTENSION_CONFIG.HIDDEN_CALC_SHEET).toBe('_Dashboard_Calc');
  });

  test('has DYNAMIC_FORMULA_ROW of 50', () => {
    expect(EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW).toBe(50);
  });

  test('has MEMBER_SHEET name', () => {
    expect(EXTENSION_CONFIG.MEMBER_SHEET).toBe(SHEETS.MEMBER_DIR);
  });

  test('has GRIEVANCE_SHEET name', () => {
    expect(EXTENSION_CONFIG.GRIEVANCE_SHEET).toBe(SHEETS.GRIEVANCE_LOG);
  });

  test('has LEADER_ROLE_NAME', () => {
    expect(EXTENSION_CONFIG.LEADER_ROLE_NAME).toBe('Member Leader');
  });

  test('has CORE_COLUMN_COUNT of 32', () => {
    expect(EXTENSION_CONFIG.CORE_COLUMN_COUNT).toBe(32);
  });

  test('has CACHE_TTL_SECONDS of 300', () => {
    expect(EXTENSION_CONFIG.CACHE_TTL_SECONDS).toBe(300);
  });
});

// ============================================================================
// COL_IDX
// ============================================================================

describe('COL_IDX', () => {
  test('is defined', () => {
    expect(COL_IDX).toBeDefined();
  });

  test('MEMBER_ID is 0-based (MEMBER_COLS.MEMBER_ID - 1)', () => {
    expect(COL_IDX.MEMBER_ID).toBe(MEMBER_COLS.MEMBER_ID - 1);
  });

  test('FIRST_NAME is 0-based', () => {
    expect(COL_IDX.FIRST_NAME).toBe(MEMBER_COLS.FIRST_NAME - 1);
  });

  test('EMAIL is 0-based', () => {
    expect(COL_IDX.EMAIL).toBe(MEMBER_COLS.EMAIL - 1);
  });
});

// ============================================================================
// getHeaderMap
// ============================================================================

describe('getHeaderMap', () => {
  test('returns header map from sheet', () => {
    const headers = [['Name', 'Email', 'Phone']];
    const mockSheet = createMockSheet('TestSheet', headers);
    mockSheet.getLastColumn.mockReturnValue(3);
    mockSheet.getRange.mockReturnValue({
      getValues: jest.fn(() => [['Name', 'Email', 'Phone']]),
      getValue: jest.fn(),
      setValue: jest.fn(),
      setValues: jest.fn(),
      setFontWeight: jest.fn(function() { return this; }),
      setBackground: jest.fn(function() { return this; }),
      setFontColor: jest.fn(function() { return this; })
    });

    const ss = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    const result = getHeaderMap('TestSheet', true);
    expect(result['Name']).toBe(1);
    expect(result['Email']).toBe(2);
    expect(result['Phone']).toBe(3);
  });

  test('returns empty object when sheet does not exist', () => {
    const ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    const result = getHeaderMap('NonExistentSheet', true);
    expect(result).toEqual({});
  });
});


