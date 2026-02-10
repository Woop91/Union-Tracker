/**
 * Tests for 08_SheetUtils.gs
 *
 * Covers: padRight, getCurrentQuarter, getQuarterFromDate, validateMemberEmail,
 * getConfigValues, setupAuditLogSheet, getAuditHistory, CREATE_509_DASHBOARD,
 * getOrCreateSheet, and related utilities.
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

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '08_SheetUtils.gs'
]);

// ============================================================================
// padRight
// ============================================================================

describe('padRight', () => {
  test('pads a short string with spaces', () => {
    expect(padRight('hi', 5)).toBe('hi   ');
  });

  test('returns exact string when length matches', () => {
    expect(padRight('hello', 5)).toBe('hello');
  });

  test('truncates a long string to specified length', () => {
    expect(padRight('hello world', 5)).toBe('hello');
  });

  test('handles empty string', () => {
    expect(padRight('', 3)).toBe('   ');
  });

  test('handles zero length', () => {
    expect(padRight('hello', 0)).toBe('');
  });

  test('converts non-string input via String()', () => {
    expect(padRight(42, 5)).toBe('42   ');
  });

  test('converts null via String()', () => {
    expect(padRight(null, 6)).toBe('null  ');
  });

  test('converts undefined via String()', () => {
    const result = padRight(undefined, 12);
    expect(result).toBe('undefined   ');
  });

  test('converts boolean via String()', () => {
    expect(padRight(true, 6)).toBe('true  ');
  });
});

// ============================================================================
// getCurrentQuarter
// ============================================================================

describe('getCurrentQuarter', () => {
  test('returns string in YYYY-QN format', () => {
    const result = getCurrentQuarter();
    expect(result).toMatch(/^\d{4}-Q[1-4]$/);
  });

  test('quarter is between Q1 and Q4', () => {
    const result = getCurrentQuarter();
    const quarter = parseInt(result.split('-Q')[1], 10);
    expect(quarter).toBeGreaterThanOrEqual(1);
    expect(quarter).toBeLessThanOrEqual(4);
  });

  test('year matches current year', () => {
    const result = getCurrentQuarter();
    const year = parseInt(result.split('-Q')[0], 10);
    expect(year).toBe(new Date().getFullYear());
  });
});

// ============================================================================
// getQuarterFromDate
// ============================================================================

describe('getQuarterFromDate', () => {
  test('January is Q1', () => {
    expect(getQuarterFromDate(new Date(2026, 0, 15))).toBe('2026-Q1');
  });

  test('March is Q1', () => {
    expect(getQuarterFromDate(new Date(2026, 2, 31))).toBe('2026-Q1');
  });

  test('April is Q2', () => {
    expect(getQuarterFromDate(new Date(2026, 3, 1))).toBe('2026-Q2');
  });

  test('June is Q2', () => {
    expect(getQuarterFromDate(new Date(2026, 5, 15))).toBe('2026-Q2');
  });

  test('July is Q3', () => {
    expect(getQuarterFromDate(new Date(2026, 6, 4))).toBe('2026-Q3');
  });

  test('September is Q3', () => {
    expect(getQuarterFromDate(new Date(2026, 8, 30))).toBe('2026-Q3');
  });

  test('October is Q4', () => {
    expect(getQuarterFromDate(new Date(2026, 9, 1))).toBe('2026-Q4');
  });

  test('December is Q4', () => {
    expect(getQuarterFromDate(new Date(2026, 11, 25))).toBe('2026-Q4');
  });

  test('handles different years', () => {
    expect(getQuarterFromDate(new Date(2024, 7, 10))).toBe('2024-Q3');
  });
});

// ============================================================================
// getConfigValues
// ============================================================================

describe('getConfigValues', () => {
  test('returns non-empty values from config column', () => {
    const mockSheet = createMockSheet('Config', [
      ['Header'],
      ['SubHeader'],
      ['Value1'],
      ['Value2'],
      ['Value3']
    ]);
    mockSheet.getLastRow.mockReturnValue(5);
    mockSheet.getRange.mockReturnValue(
      createMockRange([['Value1'], ['Value2'], ['Value3']])
    );

    const result = getConfigValues(mockSheet, 1);
    expect(result).toEqual(['Value1', 'Value2', 'Value3']);
  });

  test('filters out empty values', () => {
    const mockSheet = createMockSheet('Config');
    mockSheet.getLastRow.mockReturnValue(6);
    mockSheet.getRange.mockReturnValue(
      createMockRange([['Value1'], [''], ['Value3'], ['  ']])
    );

    const result = getConfigValues(mockSheet, 1);
    expect(result).toEqual(['Value1', 'Value3']);
  });

  test('returns empty array when lastRow < 3', () => {
    const mockSheet = createMockSheet('Config');
    mockSheet.getLastRow.mockReturnValue(2);

    const result = getConfigValues(mockSheet, 1);
    expect(result).toEqual([]);
  });

  test('converts non-string values to strings', () => {
    const mockSheet = createMockSheet('Config');
    mockSheet.getLastRow.mockReturnValue(4);
    mockSheet.getRange.mockReturnValue(
      createMockRange([[42], [true]])
    );

    const result = getConfigValues(mockSheet, 1);
    expect(result).toEqual(['42', 'true']);
  });
});

// ============================================================================
// getOrCreateSheet
// ============================================================================

describe('getOrCreateSheet', () => {
  test('returns existing sheet and clears it', () => {
    const existingSheet = createMockSheet('TestSheet');
    const ss = createMockSpreadsheet([existingSheet]);

    const result = getOrCreateSheet(ss, 'TestSheet');
    expect(result).toBe(existingSheet);
    expect(existingSheet.clear).toHaveBeenCalled();
  });

  test('creates new sheet when not found', () => {
    const ss = createMockSpreadsheet([]);

    const result = getOrCreateSheet(ss, 'NewSheet');
    expect(ss.insertSheet).toHaveBeenCalledWith('NewSheet');
    expect(result).toBeDefined();
  });

  test('does not call insertSheet when sheet exists', () => {
    const existingSheet = createMockSheet('Existing');
    const ss = createMockSpreadsheet([existingSheet]);

    getOrCreateSheet(ss, 'Existing');
    expect(ss.insertSheet).not.toHaveBeenCalled();
  });
});

// ============================================================================
// validateMemberEmail
// ============================================================================

describe('validateMemberEmail', () => {
  test('returns null for empty email', () => {
    expect(validateMemberEmail('')).toBeNull();
  });

  test('returns null for null/undefined email', () => {
    expect(validateMemberEmail(null)).toBeNull();
    expect(validateMemberEmail(undefined)).toBeNull();
  });

  test('returns null when member sheet does not exist', () => {
    // Default mock returns null for getSheetByName
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    expect(validateMemberEmail('test@example.com')).toBeNull();
  });

  test('returns member info when email matches', () => {
    // Build mock data matching MEMBER_COLS structure
    const headerRow = new Array(40).fill('');
    const memberRow = new Array(40).fill('');
    memberRow[MEMBER_COLS.MEMBER_ID - 1] = 'M001';
    memberRow[MEMBER_COLS.FIRST_NAME - 1] = 'John';
    memberRow[MEMBER_COLS.LAST_NAME - 1] = 'Doe';
    memberRow[MEMBER_COLS.EMAIL - 1] = 'john@example.com';

    const mockSheet = createMockSheet(SHEETS.MEMBER_DIR, [headerRow, memberRow]);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = validateMemberEmail('john@example.com');
    expect(result).toEqual({
      memberId: 'M001',
      firstName: 'John',
      lastName: 'Doe',
      email: 'john@example.com'
    });
  });

  test('email matching is case-insensitive', () => {
    const headerRow = new Array(40).fill('');
    const memberRow = new Array(40).fill('');
    memberRow[MEMBER_COLS.MEMBER_ID - 1] = 'M002';
    memberRow[MEMBER_COLS.FIRST_NAME - 1] = 'Jane';
    memberRow[MEMBER_COLS.LAST_NAME - 1] = 'Smith';
    memberRow[MEMBER_COLS.EMAIL - 1] = 'Jane@Example.COM';

    const mockSheet = createMockSheet(SHEETS.MEMBER_DIR, [headerRow, memberRow]);
    const mockSS = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = validateMemberEmail('jane@example.com');
    expect(result).not.toBeNull();
    expect(result.firstName).toBe('Jane');
  });
});

// ============================================================================
// setupAuditLogSheet
// ============================================================================

describe('setupAuditLogSheet', () => {
  test('creates audit log sheet if it does not exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    setupAuditLogSheet();

    expect(mockSS.insertSheet).toHaveBeenCalled();
  });

  test('clears existing audit log sheet', () => {
    const auditSheet = createMockSheet(SHEETS.AUDIT_LOG || '_Audit_Log');
    const mockSS = createMockSpreadsheet([auditSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    setupAuditLogSheet();

    expect(auditSheet.clear).toHaveBeenCalled();
  });

  test('sets frozen rows on audit log sheet', () => {
    const auditSheet = createMockSheet(SHEETS.AUDIT_LOG || '_Audit_Log');
    const mockSS = createMockSpreadsheet([auditSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    setupAuditLogSheet();

    expect(auditSheet.setFrozenRows).toHaveBeenCalledWith(1);
  });
});

// ============================================================================
// getAuditHistory
// ============================================================================

describe('getAuditHistory', () => {
  test('returns empty array when audit sheet does not exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getAuditHistory('GRV-001');
    expect(result).toEqual([]);
  });

  test('returns empty array when sheet has no data rows', () => {
    const auditSheet = createMockSheet(SHEETS.AUDIT_LOG || '_Audit_Log', [['Header']]);
    auditSheet.getLastRow.mockReturnValue(1);
    const mockSS = createMockSpreadsheet([auditSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getAuditHistory('GRV-001');
    expect(result).toEqual([]);
  });

  test('returns matching audit entries for a record ID', () => {
    const auditData = [
      ['Timestamp', 'User', 'Sheet', 'Row', 'Col', 'Field', 'Old', 'New', 'RecordID', 'Action'],
      ['2026-01-01', 'user@test.com', 'Grievance Log', 2, 3, 'Status', 'Open', 'Closed', 'GRV-001', 'Edit'],
      ['2026-01-02', 'other@test.com', 'Member Dir', 3, 4, 'Name', 'A', 'B', 'M-002', 'Edit'],
      ['2026-01-03', 'user@test.com', 'Grievance Log', 2, 5, 'Notes', '', 'Resolved', 'GRV-001', 'Add']
    ];
    const auditSheet = createMockSheet(SHEETS.AUDIT_LOG || '_Audit_Log', auditData);
    auditSheet.getLastRow.mockReturnValue(4);
    const mockSS = createMockSpreadsheet([auditSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getAuditHistory('GRV-001');
    expect(result).toHaveLength(2);
    expect(result[0].field).toBe('Status');
    expect(result[1].field).toBe('Notes');
    expect(result[0].user).toBe('user@test.com');
  });

  test('returns empty array when no records match', () => {
    const auditData = [
      ['Timestamp', 'User', 'Sheet', 'Row', 'Col', 'Field', 'Old', 'New', 'RecordID', 'Action'],
      ['2026-01-01', 'user@test.com', 'Grievance Log', 2, 3, 'Status', 'Open', 'Closed', 'GRV-999', 'Edit']
    ];
    const auditSheet = createMockSheet(SHEETS.AUDIT_LOG || '_Audit_Log', auditData);
    auditSheet.getLastRow.mockReturnValue(2);
    const mockSS = createMockSpreadsheet([auditSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const result = getAuditHistory('GRV-001');
    expect(result).toEqual([]);
  });
});

// ============================================================================
// CREATE_509_DASHBOARD - smoke test
// ============================================================================

describe('CREATE_509_DASHBOARD', () => {
  test('is defined as a function', () => {
    expect(typeof CREATE_509_DASHBOARD).toBe('function');
  });

  test('does not throw when UI cancels', () => {
    const ui = SpreadsheetApp.getUi();
    ui.alert.mockReturnValue('NO');
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    expect(() => CREATE_509_DASHBOARD()).not.toThrow();
  });
});

// ============================================================================
// reorderSheetsToStandard
// ============================================================================

describe('reorderSheetsToStandard', () => {
  test('is defined as a function', () => {
    expect(typeof reorderSheetsToStandard).toBe('function');
  });

  test('calls setActiveSheet and moveActiveSheet for found sheets', () => {
    const gettingStarted = createMockSheet(SHEETS.GETTING_STARTED);
    const faq = createMockSheet(SHEETS.FAQ);
    const ss = createMockSpreadsheet([gettingStarted, faq]);

    reorderSheetsToStandard(ss);

    expect(ss.setActiveSheet).toHaveBeenCalled();
    expect(ss.moveActiveSheet).toHaveBeenCalled();
  });
});

// ============================================================================
// Constants check
// ============================================================================

describe('Sheet utility constants', () => {
  test('SHEETS.AUDIT_LOG is defined', () => {
    expect(SHEETS.AUDIT_LOG).toBeDefined();
  });

  test('SHEETS.AUDIT_LOG starts with underscore', () => {
    expect(SHEETS.AUDIT_LOG[0]).toBe('_');
  });

  test('COLORS constant is available', () => {
    expect(COLORS).toBeDefined();
    expect(COLORS.PRIMARY_PURPLE).toBeDefined();
  });
});

// ============================================================================
// setupCalcFormulasSheet (exercises GRIEVANCE_OUTCOMES and GRIEVANCE_STATUS)
// ============================================================================

describe('setupCalcFormulasSheet', () => {
  beforeEach(() => jest.clearAllMocks());

  function createTrackingSheet() {
    const cells = {};
    const mockSheet = createMockSheet('_CalcFormulas');

    mockSheet.getRange.mockImplementation((a1) => {
      const range = createMockRange();
      range.setValue.mockImplementation((val) => { cells[a1] = val; });
      range.setFormula = jest.fn((val) => { cells[a1 + '_formula'] = val; });
      range.setFontWeight.mockReturnThis();
      return range;
    });

    return { mockSheet, cells };
  }

  test('is defined as a function', () => {
    expect(typeof setupCalcFormulasSheet).toBe('function');
  });

  test('does not throw (GRIEVANCE_OUTCOMES is defined)', () => {
    const { mockSheet } = createTrackingSheet();
    expect(() => setupCalcFormulasSheet(mockSheet)).not.toThrow();
  });

  test('writes STATUS_LIST from GRIEVANCE_STATUS values', () => {
    const { mockSheet, cells } = createTrackingSheet();

    setupCalcFormulasSheet(mockSheet);

    expect(cells['A6']).toBe('STATUS_LIST');
    expect(cells['B6']).toBe(Object.values(GRIEVANCE_STATUS).join(','));
  });

  test('writes OUTCOME_LIST from GRIEVANCE_OUTCOMES values', () => {
    const { mockSheet, cells } = createTrackingSheet();

    setupCalcFormulasSheet(mockSheet);

    expect(cells['A8']).toBe('OUTCOME_LIST');
    expect(cells['B8']).toBe(Object.values(GRIEVANCE_OUTCOMES).join(','));
    expect(cells['B8']).toContain('Pending');
    expect(cells['B8']).toContain('Won');
    expect(cells['B8']).toContain('Denied');
  });

  test('writes GRIEVANCE_TYPES list', () => {
    const { mockSheet, cells } = createTrackingSheet();

    setupCalcFormulasSheet(mockSheet);

    expect(cells['A10']).toBe('GRIEVANCE_TYPES');
    expect(cells['B10']).toContain('Discipline');
  });

  test('writes header and department formula', () => {
    const { mockSheet, cells } = createTrackingSheet();

    setupCalcFormulasSheet(mockSheet);

    expect(cells['A1']).toBe('Named Formula References');
    expect(cells['A4']).toBe('DEPARTMENT_LIST');
  });
});
