/**
 * Tests for 08d_AuditAndFormulas.gs
 * Covers audit logging, hidden sheet setup, vault integrity,
 * HMAC computation, and formula sheet creation.
 */

const { createMockSheet, createMockSpreadsheet, createMockRange } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '08d_AuditAndFormulas.gs'
]);

// ============================================================================
// Function existence
// ============================================================================

describe('08d function existence', () => {
  const required = [
    'setupAuditLogSheet', 'protectAuditLogSheet_', 'onEditAudit',
    'setupGrievanceCalcSheet', 'setupGrievanceFormulasSheet',
    'setupMemberLookupSheet', 'setupStewardContactCalcSheet',
    'setupDashboardCalcSheet', 'setupStewardPerformanceCalcSheet',
    'setupAllHiddenSheets', 'repairAllHiddenSheets', 'verifyHiddenSheets',
    'refreshAllHiddenFormulas',
    'setupCalcMembersSheet', 'setupCalcGrievancesSheet',
    'setupCalcDeadlinesSheet', 'setupCalcStatsSheet',
    'setupCalcSyncSheet', 'setupCalcFormulasSheet',
    'setupSurveyTrackingSheet', 'setupSurveyVaultSheet',
    'getVaultDataMap_', 'getVaultDataFull_',
    'supersedePreviousVaultEntry_',
    'hashForVault_', 'hashForVaultLegacy_',
    'computeHmacSha256_', 'computeAuditRowHash_',
    'verifyAuditLogIntegrity', 'verifySurveyVaultIntegrity'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// HMAC / hashing
// ============================================================================

describe('computeHmacSha256_', () => {
  test('returns a string', () => {
    const result = computeHmacSha256_('test data', 'secret key');
    expect(typeof result).toBe('string');
  });

  test('different inputs produce different outputs', () => {
    const a = computeHmacSha256_('data A', 'key');
    const b = computeHmacSha256_('data B', 'key');
    expect(a).not.toBe(b);
  });

  test('same inputs produce same output (deterministic)', () => {
    const a = computeHmacSha256_('same data', 'same key');
    const b = computeHmacSha256_('same data', 'same key');
    expect(a).toBe(b);
  });
});

describe('hashForVault_', () => {
  test('returns a string', () => {
    const result = hashForVault_('member123', 'response data');
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });
});

describe('computeAuditRowHash_', () => {
  test('returns a string for array input', () => {
    const row = ['2026-03-11', 'test@example.com', 'EDIT', 'Changed field X'];
    const result = computeAuditRowHash_(row);
    expect(typeof result).toBe('string');
  });
});

// ============================================================================
// Hidden sheets constants
// ============================================================================

// ============================================================================
// Behavioral: setupAuditLogSheet
// ============================================================================

describe('setupAuditLogSheet (behavioral)', () => {
  test('creates the audit log sheet when it does not exist', () => {
    var ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupAuditLogSheet();

    // insertSheet should have been called with the audit log name
    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.AUDIT_LOG);
  });

  test('writes headers to row 1 when sheet is new', () => {
    var sheet = createMockSheet(SHEETS.AUDIT_LOG, [['header']]);
    sheet.getLastRow.mockReturnValue(1); // only header row
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupAuditLogSheet();

    // clear() should be called on an empty/header-only sheet
    expect(sheet.clear).toHaveBeenCalled();
    // setFrozenRows(1) should be called for the header row
    expect(sheet.setFrozenRows).toHaveBeenCalledWith(1);
  });

  test('does NOT clear a sheet with existing audit data (compliance)', () => {
    var data = [
      ['Timestamp', 'User Email', 'Sheet', 'Row', 'Column', 'Field Name', 'Old Value', 'New Value', 'Record ID', 'Action Type'],
      ['2026-01-01', 'user@test.com', 'Members', '2', '3', 'Name', 'Old', 'New', 'M001', 'Edit']
    ];
    var sheet = createMockSheet(SHEETS.AUDIT_LOG, data);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupAuditLogSheet();

    // Should NOT call clear when data exists (compliance requirement)
    expect(sheet.clear).not.toHaveBeenCalled();
  });
});

// ============================================================================
// Behavioral: onEditAudit
// ============================================================================

describe('onEditAudit (behavioral)', () => {
  test('returns silently for null event', () => {
    expect(() => onEditAudit(null)).not.toThrow();
  });

  test('returns silently for event without range', () => {
    expect(() => onEditAudit({ range: null })).not.toThrow();
  });

  test('skips header row edits (row < 2)', () => {
    // logAuditEvent should NOT be called for header edits
    var origLog = global.logAuditEvent;
    global.logAuditEvent = jest.fn();

    var mockSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Header']]);
    var mockRange = {
      getSheet: jest.fn(() => mockSheet),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => 1),
      getNumColumns: jest.fn(() => 1),
      getA1Notation: jest.fn(() => 'A1'),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };

    onEditAudit({ range: mockRange, oldValue: 'a', value: 'b' });

    expect(global.logAuditEvent).not.toHaveBeenCalled();
    global.logAuditEvent = origLog;
  });

  test('skips edits on non-tracked sheets', () => {
    var origLog = global.logAuditEvent;
    global.logAuditEvent = jest.fn();

    var mockSheet = createMockSheet('Random Sheet', [['Header']]);
    var mockRange = {
      getSheet: jest.fn(() => mockSheet),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => 1),
      getNumColumns: jest.fn(() => 1),
      getA1Notation: jest.fn(() => 'A2'),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };

    onEditAudit({ range: mockRange, oldValue: 'a', value: 'b' });

    expect(global.logAuditEvent).not.toHaveBeenCalled();
    global.logAuditEvent = origLog;
  });

  test('logs audit event for a valid edit on Member Directory', () => {
    var origLog = global.logAuditEvent;
    global.logAuditEvent = jest.fn();

    var memberData = [
      ['Member ID', 'First Name', 'Last Name'],
      ['M-001', 'John', 'Doe']
    ];
    var mockSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    var mockRange = {
      getSheet: jest.fn(() => mockSheet),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 2),
      getNumRows: jest.fn(() => 1),
      getNumColumns: jest.fn(() => 1),
      getA1Notation: jest.fn(() => 'B2'),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };

    onEditAudit({
      range: mockRange,
      oldValue: 'John',
      value: 'Jane'
    });

    expect(global.logAuditEvent).toHaveBeenCalled();
    var callArgs = global.logAuditEvent.mock.calls[0];
    // First argument is the event type string
    expect(callArgs[0]).toContain('MEMBER_DIRECTORY');
    // Second argument contains field details
    expect(callArgs[1]).toHaveProperty('oldValue', 'John');
    expect(callArgs[1]).toHaveProperty('newValue', 'Jane');

    global.logAuditEvent = origLog;
  });

  test('skips audit when old and new values are identical', () => {
    var origLog = global.logAuditEvent;
    global.logAuditEvent = jest.fn();

    var mockSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Header'], ['val']]);
    var mockRange = {
      getSheet: jest.fn(() => mockSheet),
      getRow: jest.fn(() => 2),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => 1),
      getNumColumns: jest.fn(() => 1),
      getA1Notation: jest.fn(() => 'A2'),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };

    onEditAudit({
      range: mockRange,
      oldValue: 'same',
      value: 'same'
    });

    expect(global.logAuditEvent).not.toHaveBeenCalled();
    global.logAuditEvent = origLog;
  });
});

// ============================================================================
// Behavioral: setupGrievanceCalcSheet
// ============================================================================

describe('setupGrievanceCalcSheet (behavioral)', () => {
  test('creates hidden sheet and sets formulas', () => {
    // Create a sheet whose getRange returns a chainable mock
    var sheet = createMockSheet(SHEETS.GRIEVANCE_CALC, [['header']]);
    // Override getRange to return chainable ranges (setValues chains to setFontWeight etc.)
    var chainableRange = {
      setValues: jest.fn(function() { return this; }),
      setValue: jest.fn(function() { return this; }),
      setFormula: jest.fn(function() { return this; }),
      setFontWeight: jest.fn(function() { return this; }),
      setBackground: jest.fn(function() { return this; }),
      setFontColor: jest.fn(function() { return this; }),
      getValue: jest.fn(() => ''),
      getValues: jest.fn(() => [['']])
    };
    sheet.getRange.mockReturnValue(chainableRange);

    var ss = createMockSpreadsheet([]);
    ss.insertSheet.mockReturnValue(sheet);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    setupGrievanceCalcSheet();

    // insertSheet should have been called
    expect(ss.insertSheet).toHaveBeenCalledWith(SHEETS.GRIEVANCE_CALC);
    // The sheet should have been cleared
    expect(sheet.clear).toHaveBeenCalled();
    // Formulas should have been set (at least the header row + 7 formula columns)
    expect(chainableRange.setFormula).toHaveBeenCalled();
    expect(chainableRange.setFormula.mock.calls.length).toBeGreaterThanOrEqual(5);
  });
});

// ============================================================================
// Hidden sheets constants
// ============================================================================

describe('Hidden sheet references use HIDDEN_SHEETS constants', () => {
  test('HIDDEN_SHEETS has all required calc sheets', () => {
    expect(HIDDEN_SHEETS.CALC_STATS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_FORMULAS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_MEMBERS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_GRIEVANCES).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_DEADLINES).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_SYNC).toBeDefined();
    expect(HIDDEN_SHEETS.AUDIT_LOG).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_TRACKING).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_VAULT).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_PERIODS).toBeDefined();
  });

  test('all hidden sheet names start with underscore', () => {
    Object.entries(HIDDEN_SHEETS).forEach(([key, name]) => {
      expect(name).toMatch(/^_/);
    });
  });
});
