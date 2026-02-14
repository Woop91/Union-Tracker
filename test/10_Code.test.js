/**
 * Tests for 10_Code (split into 10a-10d modules)
 *
 * Covers form config constants, sanitizeFolderName_, padRight,
 * getExistingGrievanceIds_, getCurrentStewardInfo_,
 * createGrievanceFolderFromData_, and sheet creation smoke tests.
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
  '10c_FormHandlers.gs',
  '10d_SyncAndMaintenance.gs'
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

  test('FIELD_IDS has 18 entries', () => {
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

// ============================================================================
// CONTACT_FORM_CONFIG
// ============================================================================

describe('CONTACT_FORM_CONFIG', () => {
  test('is defined', () => {
    expect(CONTACT_FORM_CONFIG).toBeDefined();
  });

  test('has a FORM_URL property (configured via Config sheet)', () => {
    expect(typeof CONTACT_FORM_CONFIG.FORM_URL).toBe('string');
  });

  test('FIELD_IDS has 15 entries', () => {
    expect(Object.keys(CONTACT_FORM_CONFIG.FIELD_IDS).length).toBe(15);
  });

  test('FIELD_IDS includes FIRST_NAME, LAST_NAME, EMAIL, PHONE', () => {
    expect(CONTACT_FORM_CONFIG.FIELD_IDS.FIRST_NAME).toBeDefined();
    expect(CONTACT_FORM_CONFIG.FIELD_IDS.LAST_NAME).toBeDefined();
    expect(CONTACT_FORM_CONFIG.FIELD_IDS.EMAIL).toBeDefined();
    expect(CONTACT_FORM_CONFIG.FIELD_IDS.PHONE).toBeDefined();
  });
});

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
  test('removes slashes', () => {
    expect(sanitizeFolderName_('John / Smith')).toBe('John Smith');
  });

  test('removes backslashes', () => {
    expect(sanitizeFolderName_('John \\ Smith')).toBe('John Smith');
  });

  test('removes colons, asterisks, question marks, quotes, angle brackets, pipes', () => {
    expect(sanitizeFolderName_('a:b*c?"d<e>f|g')).toBe('abcdefg');
  });

  test('collapses multiple spaces into one', () => {
    expect(sanitizeFolderName_('John    Smith')).toBe('John Smith');
  });

  test('trims leading and trailing whitespace', () => {
    expect(sanitizeFolderName_('  John Smith  ')).toBe('John Smith');
  });

  test('combined: "  John / Smith  " becomes "John Smith"', () => {
    expect(sanitizeFolderName_('  John / Smith  ')).toBe('John Smith');
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

  test('handles already-clean names', () => {
    expect(sanitizeFolderName_('Clean Name')).toBe('Clean Name');
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
// getExistingGrievanceIds_
// ============================================================================

describe('getExistingGrievanceIds_', () => {
  test('returns an object mapping existing IDs from sheet data', () => {
    // Build mock data: header row + data rows with grievance IDs
    const idCol = GRIEVANCE_COLS.GRIEVANCE_ID - 1;
    const headerRow = new Array(20).fill('Header');
    const row1 = new Array(20).fill('');
    row1[idCol] = 'GRV-001';
    const row2 = new Array(20).fill('');
    row2[idCol] = 'GRV-002';
    const row3 = new Array(20).fill('');
    row3[idCol] = '';  // empty ID should be skipped

    const mockSheet = createMockSheet('Grievance Log', [headerRow, row1, row2, row3]);
    const result = getExistingGrievanceIds_(mockSheet);

    expect(result['GRV-001']).toBe(true);
    expect(result['GRV-002']).toBe(true);
    expect(Object.keys(result).length).toBe(2);
  });

  test('returns empty object for sheet with only header row', () => {
    const headerRow = new Array(20).fill('Header');
    const mockSheet = createMockSheet('Grievance Log', [headerRow]);
    const result = getExistingGrievanceIds_(mockSheet);

    expect(Object.keys(result).length).toBe(0);
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
// createGrievanceFolderFromData_ (Drive mock)
// ============================================================================

describe('createGrievanceFolderFromData_', () => {
  beforeEach(() => {
    // Reset DriveApp mocks
    const mockSubFolder = {
      getId: jest.fn(() => 'sub-folder-id'),
      getUrl: jest.fn(() => 'https://drive.google.com/sub-folder'),
      setDescription: jest.fn(),
      createFolder: jest.fn(),
      getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
      addEditor: jest.fn()
    };

    const mockRootFolder = {
      getId: jest.fn(() => 'root-folder-id'),
      getUrl: jest.fn(() => 'https://drive.google.com/root'),
      getFoldersByName: jest.fn(() => ({ hasNext: jest.fn(() => false) })),
      createFolder: jest.fn(() => mockSubFolder)
    };

    // Mock getOrCreateDashboardFolder_ to return our mock root
    global.getOrCreateDashboardFolder_ = jest.fn(() => mockRootFolder);
    // Mock shareWithCoordinators_ to avoid side effects
    global.shareWithCoordinators_ = jest.fn();
  });

  test('returns object with id and url on success', () => {
    const result = createGrievanceFolderFromData_(
      'GRV-001', 'MBR-001', 'John', 'Smith', 'Discipline', '2026-01-15'
    );
    expect(result).toHaveProperty('id');
    expect(result).toHaveProperty('url');
  });

  test('returns empty id/url when folder creation throws', () => {
    global.getOrCreateDashboardFolder_ = jest.fn(() => {
      throw new Error('Drive error');
    });

    const result = createGrievanceFolderFromData_(
      'GRV-001', 'MBR-001', 'John', 'Smith', 'Discipline', '2026-01-15'
    );
    expect(result.id).toBe('');
    expect(result.url).toBe('');
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
