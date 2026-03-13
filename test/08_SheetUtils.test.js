/**
 * Tests for 08_SheetUtils (split into 08a-08d modules)
 *
 * Covers: padRight, getCurrentQuarter, validateMemberEmail,
 * getConfigValues, setupAuditLogSheet, CREATE_DASHBOARD,
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
  '08a_SheetSetup.gs',
  '08b_SearchAndCharts.gs',
  '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs'
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
// CREATE_DASHBOARD - smoke test
// ============================================================================

describe('CREATE_DASHBOARD', () => {
  test('is defined as a function', () => {
    expect(typeof CREATE_DASHBOARD).toBe('function');
  });

  test('does not throw when UI cancels', () => {
    const ui = SpreadsheetApp.getUi();
    ui.alert.mockReturnValue('NO');
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    expect(() => CREATE_DASHBOARD()).not.toThrow();
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

// ============================================================================
// hashForVault_ — upgraded HMAC-style double hash (v4.8.1)
// ============================================================================

describe('hashForVault_', () => {
  test('returns a string starting with V', () => {
    const hash = hashForVault_('test@example.com');
    expect(hash).toMatch(/^V/);
  });

  test('returns 23-character hash (V + 22 hex chars)', () => {
    const hash = hashForVault_('test@example.com');
    expect(hash.length).toBe(23);
  });

  test('hash is all uppercase hex after V prefix', () => {
    const hash = hashForVault_('user@domain.com');
    expect(hash).toMatch(/^V[0-9A-F]{22}$/);
  });

  test('same input produces same hash (deterministic)', () => {
    const hash1 = hashForVault_('same@email.com');
    const hash2 = hashForVault_('same@email.com');
    expect(hash1).toBe(hash2);
  });

  test('different inputs produce different hashes', () => {
    const hash1 = hashForVault_('user1@example.com');
    const hash2 = hashForVault_('user2@example.com');
    expect(hash1).not.toBe(hash2);
  });

  test('normalizes input to lowercase and trims whitespace', () => {
    const hash1 = hashForVault_('Test@Example.COM');
    const hash2 = hashForVault_('  test@example.com  ');
    expect(hash1).toBe(hash2);
  });
});

// ============================================================================
// hashForVaultLegacy_ — backward-compatible 11-char hash
// ============================================================================

describe('hashForVaultLegacy_', () => {
  test('returns a string starting with V', () => {
    const hash = hashForVaultLegacy_('test@example.com');
    expect(hash).toMatch(/^V/);
  });

  test('returns 11-character hash (V + 10 alphanumeric)', () => {
    const hash = hashForVaultLegacy_('test@example.com');
    expect(hash.length).toBe(11);
  });

  test('same input produces same hash', () => {
    const h1 = hashForVaultLegacy_('x@y.com');
    const h2 = hashForVaultLegacy_('x@y.com');
    expect(h1).toBe(h2);
  });

  test('new hash differs from legacy hash for same input', () => {
    const newHash = hashForVault_('test@example.com');
    const legacyHash = hashForVaultLegacy_('test@example.com');
    expect(newHash).not.toBe(legacyHash);
    expect(newHash.length).toBeGreaterThan(legacyHash.length);
  });
});

// ============================================================================
// computeHmacSha256_
// ============================================================================

describe('computeHmacSha256_', () => {
  test('returns a 64-character hex string', () => {
    const hmac = computeHmacSha256_('secret-key', 'hello world');
    expect(hmac).toMatch(/^[0-9a-f]{64}$/);
  });

  test('same key+message produces same HMAC (deterministic)', () => {
    const h1 = computeHmacSha256_('key', 'msg');
    const h2 = computeHmacSha256_('key', 'msg');
    expect(h1).toBe(h2);
  });

  test('different keys produce different HMACs', () => {
    // Use very different keys to ensure mock hash differentiates
    const h1 = computeHmacSha256_('aaaa-key-alpha', 'message');
    const h2 = computeHmacSha256_('zzzz-key-omega-totally-different', 'message');
    expect(h1).not.toBe(h2);
  });

  test('is a function that accepts two string arguments', () => {
    expect(typeof computeHmacSha256_).toBe('function');
    expect(() => computeHmacSha256_('k', 'm')).not.toThrow();
  });
});

// ============================================================================
// computeAuditRowHash_
// ============================================================================

describe('computeAuditRowHash_', () => {
  test('returns a 16-character hex string', () => {
    const hash = computeAuditRowHash_('', new Date(), 'TEST', 'user@x.com', '{}', 'session1');
    expect(hash).toMatch(/^[0-9a-f]{16}$/);
  });

  test('same inputs produce same hash (deterministic)', () => {
    const ts = '2026-01-15T10:00:00';
    const h1 = computeAuditRowHash_('prev', ts, 'EDIT', 'u@x.com', '{"a":1}', 's1');
    const h2 = computeAuditRowHash_('prev', ts, 'EDIT', 'u@x.com', '{"a":1}', 's1');
    expect(h1).toBe(h2);
  });

  test('empty previous hash does not throw', () => {
    expect(() => {
      computeAuditRowHash_('', new Date(), 'E', 'u@x.com', '{}', 's');
    }).not.toThrow();
  });
});

// ============================================================================
// verifySurveyVaultIntegrity
// ============================================================================

describe('verifySurveyVaultIntegrity', () => {
  test('returns valid for empty vault', () => {
    const result = verifySurveyVaultIntegrity();
    expect(result.valid).toBe(true);
    expect(result.stats.totalEntries).toBe(0);
  });
});
