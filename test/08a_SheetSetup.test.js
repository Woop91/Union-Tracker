/**
 * Tests for 08a_SheetSetup.gs
 * Covers sheet creation, data validation setup, multi-select handling,
 * config value retrieval, and dropdown validation.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '08a_SheetSetup.gs'
]);

describe('08a function existence', () => {
  const required = [
    'CREATE_DASHBOARD', '_ensureContactLogSheet', '_ensureStewardTasksSheet',
    'getOrCreateSheet', 'reorderSheetsToStandard', 'setupHiddenSheets',
    'setupDataValidations',
    'setDropdownValidation', 'setMultiSelectValidation',
    'getConfigValues', 'applyMultiSelectValue', 'clearMultiSelectState',
    'onEditMultiSelect', 'onSelectionChangeMultiSelect',
    'installMultiSelectTrigger', 'removeMultiSelectTrigger'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('_ensureContactLogSheet', () => {
  test('creates sheet if not found', () => {
    const mockSheet = {
      getRange: jest.fn(() => ({
        setValues: jest.fn(() => ({ setFontWeight: jest.fn(() => ({ setBackground: jest.fn(() => ({ setFontColor: jest.fn() })) })) }))
      })),
      hideSheet: jest.fn()
    };
    const mockSS = {
      getSheetByName: jest.fn(() => null),
      insertSheet: jest.fn(() => mockSheet)
    };
    const result = _ensureContactLogSheet(mockSS);
    expect(mockSS.insertSheet).toHaveBeenCalledWith(SHEETS.CONTACT_LOG);
  });

  test('returns existing sheet if found', () => {
    const existing = { getName: jest.fn(() => '_Contact_Log') };
    const mockSS = {
      getSheetByName: jest.fn(() => existing)
    };
    const result = _ensureContactLogSheet(mockSS);
    expect(result).toBe(existing);
    expect(mockSS.insertSheet).toBeUndefined();
  });
});

describe('SHEETS.CONTACT_LOG constant', () => {
  test('is defined', () => {
    expect(SHEETS.CONTACT_LOG).toBeDefined();
    expect(typeof SHEETS.CONTACT_LOG).toBe('string');
  });

  test('starts with underscore (hidden)', () => {
    expect(SHEETS.CONTACT_LOG).toMatch(/^_/);
  });
});

// ============================================================================
// getOrCreateSheet — behavioral tests
// ============================================================================

describe('getOrCreateSheet', () => {
  test('returns existing sheet when found', () => {
    var existingSheet = createMockSheet('TestSheet', [['Header1']]);
    var mockSS = createMockSpreadsheet([existingSheet]);

    var result = getOrCreateSheet(mockSS, 'TestSheet');
    expect(result).toBe(existingSheet);
    expect(mockSS.insertSheet).not.toHaveBeenCalled();
  });

  test('creates new sheet when not found', () => {
    var mockSS = createMockSpreadsheet([]);

    var result = getOrCreateSheet(mockSS, 'BrandNewSheet');
    expect(mockSS.insertSheet).toHaveBeenCalledWith('BrandNewSheet');
    // insertSheet returns a mock sheet with the given name
    expect(result).toBeDefined();
    expect(result.getName()).toBe('BrandNewSheet');
  });

  test('preserves existing content when sheet has data rows', () => {
    // Sheet with header + 1 data row — must NOT be cleared
    var sheetWithData = createMockSheet('DataSheet', [['Header'], ['Row1']]);
    var mockSS = createMockSpreadsheet([sheetWithData]);

    var result = getOrCreateSheet(mockSS, 'DataSheet');
    expect(result).toBe(sheetWithData);
    expect(sheetWithData.clear).not.toHaveBeenCalled();
  });
});

// ============================================================================
// getConfigValues — behavioral tests
// ============================================================================

describe('getConfigValues', () => {
  test('returns non-empty values from config column starting at row 3', () => {
    // Config sheet: row 1 = section header, row 2 = column header, rows 3+ = data
    var configData = [
      ['Section Header'],
      ['Column Header'],
      ['Option A'],
      ['Option B'],
      [''],           // blank — should be skipped
      ['Option C']
    ];
    var configSheet = createMockSheet(SHEETS.CONFIG, configData);

    var values = getConfigValues(configSheet, 1);
    expect(Array.isArray(values)).toBe(true);
    expect(values.length).toBe(3);
    expect(values[0]).toBe('Option A');
    expect(values[1]).toBe('Option B');
    expect(values[2]).toBe('Option C');
  });

  test('returns empty array when config sheet has no data rows', () => {
    // Only header rows — lastRow < 3
    var configData = [
      ['Section Header'],
      ['Column Header']
    ];
    var configSheet = createMockSheet('Config', configData);

    var values = getConfigValues(configSheet, 1);
    expect(Array.isArray(values)).toBe(true);
    expect(values.length).toBe(0);
  });
});

// ============================================================================
// setupDataValidations — behavioral tests
// ============================================================================

describe('setupDataValidations', () => {
  test('calls setDataValidation on member sheet Yes/No columns', () => {
    // Build minimal sheets with enough rows for validation range
    var configData = [['Section'], ['Header'], ['Active'], ['Inactive']];
    var memberData = [['Header'], ['Row1'], ['Row2']];
    var grievanceData = [['Header'], ['Row1']];

    var configSheet = createMockSheet(SHEETS.CONFIG, configData);
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, grievanceData);

    var mockSS = createMockSpreadsheet([configSheet, memberSheet, grievanceSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    setupDataValidations();

    // Verify getRange was called on member sheet (for Yes/No validation)
    expect(memberSheet.getRange).toHaveBeenCalled();
    // Verify setDataValidation was called (via the mock range returned by getRange)
    // The mock range factory always includes setDataValidation
    var rangeCalls = memberSheet.getRange.mock.results;
    var setDataValidationCalled = rangeCalls.some(r =>
      r.value && r.value.setDataValidation && r.value.setDataValidation.mock.calls.length > 0
    );
    expect(setDataValidationCalled).toBe(true);
  });
});

// ============================================================================
// setDropdownValidation — behavioral tests
// ============================================================================

describe('setDropdownValidation', () => {
  test('applies data validation from config values to target column', () => {
    var configData = [['Section'], ['Header'], ['Val1'], ['Val2'], ['Val3']];
    var targetData = [['Header'], [''], ['']];
    var configSheet = createMockSheet(SHEETS.CONFIG, configData);
    var targetSheet = createMockSheet('Target', targetData);

    setDropdownValidation(targetSheet, 2, configSheet, 1);

    // SpreadsheetApp.newDataValidation should have been called
    expect(SpreadsheetApp.newDataValidation).toHaveBeenCalled();
  });
});
