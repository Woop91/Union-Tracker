// Tests cross-cutting code quality: sanitizeFolderName_, padRight, etc.
/**
 * Tests for 10_Code (split into 10a-10d modules)
 *
 * Covers form config constants, sanitizeFolderName_, padRight,
 * getCurrentStewardInfo_, and sheet creation smoke tests.
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
  '03_UIComponents.gs',
  '08a_SheetSetup.gs',
  '08b_SearchAndCharts.gs',
  '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs',
  '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs',
  '10c_FormsAndSync.gs'
]);

// ============================================================================
// GRIEVANCE_FORM_CONFIG
// ============================================================================

describe('GRIEVANCE_FORM_CONFIG', () => {
  test('is defined', () => {
    expect(GRIEVANCE_FORM_CONFIG).toBeDefined();
  });

  test('has a FORM_URL property (configured via Config sheet)', () => {
    expect(typeof GRIEVANCE_FORM_CONFIG.FORM_URL).toBe('string');
  });

  test('FIELD_IDS has expected number of entries', () => {
    // 18 fields defined in GRIEVANCE_FORM_CONFIG.FIELD_IDS (10c_FormsAndSync.gs):
    // MEMBER_ID, MEMBER_FIRST_NAME, MEMBER_LAST_NAME, JOB_TITLE,
    // AGENCY_DEPARTMENT, REGION, WORK_LOCATION, MANAGERS, MEMBER_EMAIL,
    // STEWARD_FIRST_NAME, STEWARD_LAST_NAME, STEWARD_EMAIL,
    // DATE_OF_INCIDENT, ARTICLES_VIOLATED, REMEDY_SOUGHT,
    // DATE_FILED, STEP, CONFIDENTIAL_WAIVER
    expect(Object.keys(GRIEVANCE_FORM_CONFIG.FIELD_IDS).length).toBe(18);
  });

  test('FIELD_IDS keys include MEMBER_ID, MEMBER_FIRST_NAME, STEWARD_EMAIL', () => {
    expect(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_ID).toBeDefined();
    expect(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_FIRST_NAME).toBeDefined();
    expect(GRIEVANCE_FORM_CONFIG.FIELD_IDS.STEWARD_EMAIL).toBeDefined();
  });

  test('all FIELD_IDS values start with "entry."', () => {
    Object.values(GRIEVANCE_FORM_CONFIG.FIELD_IDS).forEach(val => {
      expect(val).toMatch(/^entry\.\d+$/);
    });
  });
});

// CONTACT_FORM_CONFIG — removed (contact form deprecated)

// ============================================================================
// SATISFACTION_FORM_CONFIG
// ============================================================================

describe('SATISFACTION_FORM_CONFIG', () => {
  test('is defined', () => {
    expect(SATISFACTION_FORM_CONFIG).toBeDefined();
  });

  test('has FORM_URL property (configured via Config sheet)', () => {
    expect(typeof SATISFACTION_FORM_CONFIG.FORM_URL).toBe('string');
  });

  test('FIELD_IDS has entries', () => {
    expect(Object.keys(SATISFACTION_FORM_CONFIG.FIELD_IDS).length).toBeGreaterThan(0);
  });
});

// ============================================================================
// sanitizeFolderName_
// ============================================================================

describe('sanitizeFolderName_', () => {
  // sanitizeFolderName_ now delegates to sanitizeFolderName (05_Integrations.gs)
  // which replaces spaces with underscores for consistent folder names
  test('removes slashes', () => {
    expect(sanitizeFolderName_('John / Smith')).toBe('John_Smith');
  });

  test('removes backslashes', () => {
    expect(sanitizeFolderName_('John \\ Smith')).toBe('John_Smith');
  });

  test('removes colons, asterisks, question marks, quotes, angle brackets, pipes', () => {
    expect(sanitizeFolderName_('a:b*c?"d<e>f|g')).toBe('abcdefg');
  });

  test('collapses multiple spaces into single underscore', () => {
    expect(sanitizeFolderName_('John    Smith')).toBe('John_Smith');
  });

  test('trims leading and trailing whitespace', () => {
    expect(sanitizeFolderName_('  John Smith  ')).toBe('John_Smith');
  });

  test('combined: "  John / Smith  " becomes "John_Smith"', () => {
    expect(sanitizeFolderName_('  John / Smith  ')).toBe('John_Smith');
  });

  test('returns empty string for null', () => {
    expect(sanitizeFolderName_(null)).toBe('');
  });

  test('returns empty string for undefined', () => {
    expect(sanitizeFolderName_(undefined)).toBe('');
  });

  test('returns empty string for empty string', () => {
    expect(sanitizeFolderName_('')).toBe('');
  });

  test('truncates to 50 characters', () => {
    const long = 'a'.repeat(100);
    const result = sanitizeFolderName_(long);
    expect(result.length).toBe(50);
  });

  test('handles already-clean names (spaces become underscores)', () => {
    expect(sanitizeFolderName_('Clean Name')).toBe('Clean_Name');
  });

  test('converts non-string to string via toString', () => {
    expect(sanitizeFolderName_(12345)).toBe('12345');
  });
});

// ============================================================================
// padRight
// ============================================================================

describe('padRight', () => {
  test('pads short string with spaces', () => {
    expect(padRight('hi', 5)).toBe('hi   ');
  });

  test('returns exact length when string matches target', () => {
    expect(padRight('hello', 5)).toBe('hello');
  });

  test('truncates string longer than target length', () => {
    expect(padRight('hello world', 5)).toBe('hello');
  });

  test('handles empty string', () => {
    expect(padRight('', 3)).toBe('   ');
  });

  test('converts non-string input to string', () => {
    expect(padRight(42, 5)).toBe('42   ');
  });

  test('returns string of length 0 when len is 0', () => {
    expect(padRight('test', 0)).toBe('');
  });
});

// ============================================================================
// getCurrentStewardInfo_
// ============================================================================

describe('getCurrentStewardInfo_', () => {
  test('returns steward info when current user is a steward', () => {
    // Build member data with a steward whose email matches the mock Session user
    const emailCol = MEMBER_COLS.EMAIL - 1;
    const isStewardCol = MEMBER_COLS.IS_STEWARD - 1;
    const firstNameCol = MEMBER_COLS.FIRST_NAME - 1;
    const lastNameCol = MEMBER_COLS.LAST_NAME - 1;

    const headerRow = new Array(33).fill('Header');
    const memberRow = new Array(33).fill('');
    memberRow[emailCol] = 'test@example.com';
    memberRow[isStewardCol] = 'Yes';
    memberRow[firstNameCol] = 'Test';
    memberRow[lastNameCol] = 'Steward';

    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [headerRow, memberRow]);
    const ss = createMockSpreadsheet([memberSheet]);

    const result = getCurrentStewardInfo_(ss);
    expect(result.firstName).toBe('Test');
    expect(result.lastName).toBe('Steward');
    expect(result.email).toBe('test@example.com');
  });

  test('returns email only when user is not a steward', () => {
    const emailCol = MEMBER_COLS.EMAIL - 1;
    const isStewardCol = MEMBER_COLS.IS_STEWARD - 1;

    const headerRow = new Array(33).fill('Header');
    const memberRow = new Array(33).fill('');
    memberRow[emailCol] = 'other@example.com';
    memberRow[isStewardCol] = 'No';

    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [headerRow, memberRow]);
    const ss = createMockSpreadsheet([memberSheet]);

    const result = getCurrentStewardInfo_(ss);
    expect(result.firstName).toBe('');
    expect(result.lastName).toBe('');
    expect(result.email).toBe('test@example.com');
  });

  test('returns defaults when member sheet does not exist', () => {
    const ss = createMockSpreadsheet([]);

    const result = getCurrentStewardInfo_(ss);
    expect(result.firstName).toBe('');
    expect(result.lastName).toBe('');
    expect(result.email).toBe('test@example.com');
  });
});


// ============================================================================
// Sheet creation smoke tests
// ============================================================================

/**
 * Creates a fully chainable mock sheet where getRange returns
 * an object whose every method returns itself (for .setValues().setBackground() etc.)
 */
function createChainableMockSheet(name) {
  const chainable = new Proxy({}, {
    get: function(target, prop) {
      if (prop === 'then') return undefined; // prevent Promise-like behavior
      return jest.fn(function() { return chainable; });
    }
  });

  return {
    getName: jest.fn(() => name),
    getDataRange: jest.fn(() => chainable),
    getRange: jest.fn(() => chainable),
    getLastRow: jest.fn(() => 1),
    getLastColumn: jest.fn(() => 1),
    appendRow: jest.fn(),
    deleteRow: jest.fn(),
    deleteRows: jest.fn(),
    hideSheet: jest.fn(),
    showSheet: jest.fn(),
    setFrozenRows: jest.fn(),
    setColumnWidth: jest.fn(),
    setTabColor: jest.fn(),
    clear: jest.fn(),
    autoResizeColumns: jest.fn(),
    getColumnWidth: jest.fn(() => 80),
    getMaxColumns: jest.fn(() => 100),
    insertColumnsAfter: jest.fn(),
    getConditionalFormatRules: jest.fn(() => []),
    setConditionalFormatRules: jest.fn(),
    getFilter: jest.fn(() => null),
    createFilter: jest.fn(),
    setNumberFormat: jest.fn(),
    protect: jest.fn(() => ({
      setDescription: jest.fn(function() { return this; }),
      setWarningOnly: jest.fn(function() { return this; })
    }))
  };
}

describe('createConfigSheet (smoke test)', () => {
  test('does not throw when called with a mock spreadsheet', () => {
    const ss = createMockSpreadsheet([]);
    const mockSheet = createChainableMockSheet(SHEETS.CONFIG);
    global.getOrCreateSheet = jest.fn(() => mockSheet);
    global.applyConfigSheetStyling = jest.fn();

    expect(() => createConfigSheet(ss)).not.toThrow();
  });
});

describe('createMemberDirectory (smoke test)', () => {
  test('function is defined and callable', () => {
    expect(typeof createMemberDirectory).toBe('function');
  });
});

describe('createGrievanceLog (smoke test)', () => {
  test('function is defined and callable', () => {
    expect(typeof createGrievanceLog).toBe('function');
  });
});
