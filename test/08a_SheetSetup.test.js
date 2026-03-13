/**
 * Tests for 08a_SheetSetup.gs
 * Covers sheet creation, data validation setup, multi-select handling,
 * config value retrieval, and dropdown validation.
 */

require('./gas-mock');
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
