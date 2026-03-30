/**
 * Tests for 10a_SheetCreation.gs
 * Covers sheet creation, config seeding, header validation,
 * and dashboard setup functions.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '08a_SheetSetup.gs', '08d_AuditAndFormulas.gs',
  '10a_SheetCreation.gs'
]);

// ============================================================================
// Function existence
// ============================================================================

describe('10a function existence', () => {
  const required = [
    'createConfigSheet', 'seedConfigDefault_',
    'populateConfigFromSheetData', 'deduplicateAndSortConfigColumns_',
    'applyConfigSheetStyling', 'applySectionColors_', 'applyConfigStyling',
    'createConfigGuideSheet', '_addMissingMemberHeaders_',
    'createMemberDirectory', '_addMissingGrievanceHeaders_',
    'createGrievanceLog',
    'createVolunteerHoursSheet', 'createMeetingAttendanceSheet',
    'createMeetingCheckInLogSheet', 'setupMeetingCheckInSheet'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// Config column definitions are dynamic
// ============================================================================

describe('Config column structure', () => {
  test('CONFIG_COLS is defined and non-empty', () => {
    expect(typeof CONFIG_COLS).toBe('object');
    expect(Object.keys(CONFIG_COLS).length).toBeGreaterThan(10);
  });

  test('CONFIG_COLS has ORG_NAME', () => {
    expect(CONFIG_COLS.ORG_NAME).toBeDefined();
    expect(typeof CONFIG_COLS.ORG_NAME).toBe('number');
  });

  test('CONFIG_COLS has MAIN_CONTACT_EMAIL', () => {
    expect(CONFIG_COLS.MAIN_CONTACT_EMAIL).toBeDefined();
  });

  test('CONFIG_COLS has ACCENT_HUE', () => {
    expect(CONFIG_COLS.ACCENT_HUE).toBeDefined();
  });
});

// ============================================================================
// Member and Grievance headers use constants
// ============================================================================

describe('Member header map completeness', () => {
  test('MEMBER_HEADER_MAP_ is defined', () => {
    expect(typeof MEMBER_HEADER_MAP_).not.toBe('undefined');
  });

  test('MEMBER_COLS has essential fields', () => {
    expect(MEMBER_COLS.MEMBER_ID).toBeDefined();
    expect(MEMBER_COLS.FIRST_NAME).toBeDefined();
    expect(MEMBER_COLS.LAST_NAME).toBeDefined();
    expect(MEMBER_COLS.EMAIL).toBeDefined();
    expect(MEMBER_COLS.PHONE).toBeDefined();
    expect(MEMBER_COLS.IS_STEWARD).toBeDefined();
  });
});

describe('Grievance header map completeness', () => {
  test('GRIEVANCE_COLS has essential fields', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_ID).toBeDefined();
    expect(GRIEVANCE_COLS.MEMBER_ID).toBeDefined();
    expect(GRIEVANCE_COLS.STATUS).toBeDefined();
    expect(GRIEVANCE_COLS.CURRENT_STEP).toBeDefined();
  });
});

// ============================================================================
// Config sheet section header parity
// ============================================================================

describe('Config section header parity', () => {
  // Read sectionHeaders array from source to count elements
  const fs = require('fs');
  const path = require('path');
  const src = fs.readFileSync(path.resolve(__dirname, '..', 'src', '10a_SheetCreation.gs'), 'utf8');

  test('sectionHeaders array exists in source', () => {
    expect(src).toContain('var sectionHeaders');
  });

  test('sectionHeaders element count matches CONFIG_HEADER_MAP_ length (81)', () => {
    // Extract the sectionHeaders array definition from source
    const match = src.match(/var sectionHeaders\s*=\s*\[([\s\S]*?)\];/);
    expect(match).not.toBeNull();
    // Count elements: section names + empty strings
    const body = match[1];
    // Count all string literals (both named sections and empty '')
    const elements = body.match(/'[^']*'/g) || [];
    expect(elements.length).toBe(CONFIG_HEADER_MAP_.length);
  });

  test('CONFIG_HEADER_MAP_ has exactly 81 entries', () => {
    expect(CONFIG_HEADER_MAP_.length).toBe(81);
  });
});

// ============================================================================
// Config seed defaults match declared column types
// ============================================================================

describe('Config seed defaults match column types', () => {
  // Known seed defaults — must match their column's declared type
  const seedDefaults = {
    STEWARD_LABEL: 'Steward',
    MEMBER_LABEL: 'Member',
    ACCENT_HUE: '30',
    LOGO_INITIALS: 'DDS',
    MAGIC_LINK_EXPIRY_DAYS: '7',
    COOKIE_DURATION_DAYS: '30',
    ORG_ABBREV: 'DDS',
    SHOW_GRIEVANCES: 'yes',
    ENABLE_CORRELATION: 'no',
    GRIEVANCE_ARCHIVE_DAYS: '90',
    AUDIT_ARCHIVE_DAYS: '365',
    ORG_WEBSITE: 'https://www.example.org/',
    MAIN_CONTACT_EMAIL: 'your-email@your-org.org'
  };

  Object.entries(seedDefaults).forEach(([key, value]) => {
    test(`seed default for ${key} ("${value}") passes type validation`, () => {
      const result = validateConfigValue_(key, value);
      expect(result.valid).toBe(true);
    });
  });

  // Negative: cross-column contamination tests
  const crossContamination = [
    { key: 'ACCENT_HUE', value: 'Steward', desc: 'label in number column' },
    { key: 'DRIVE_FOLDER_ID', value: 'Steward', desc: 'label in ID column' },
    { key: 'FILING_DEADLINE_DAYS', value: 'In Arbitration', desc: 'status text in days column' },
    { key: 'MAIN_CONTACT_EMAIL', value: '(000) 000-0000', desc: 'phone in email column' },
    { key: 'SHOW_GRIEVANCES', value: 'DDS', desc: 'abbreviation in boolean column' },
    { key: 'PDF_FOLDER_ID', value: 'Step II', desc: 'step name in ID column' }
  ];

  crossContamination.forEach(({ key, value, desc }) => {
    test(`DETECTS cross-contamination: ${desc}`, () => {
      const result = validateConfigValue_(key, value);
      expect(result.valid).toBe(false);
    });
  });
});

// ============================================================================
// _migrateOrphanedColumns logic
// ============================================================================

describe('_migrateOrphanedColumns', () => {
  test('function exists', () => {
    expect(typeof _migrateOrphanedColumns).toBe('function');
  });

  test('repairConfigData function exists', () => {
    expect(typeof repairConfigData).toBe('function');
  });

  test('returns 0 when sheet has correct column count', () => {
    const mockSheet = {
      getMaxColumns: () => CONFIG_HEADER_MAP_.length,
      getRange: () => ({ getValues: () => [getHeadersFromMap_(CONFIG_HEADER_MAP_)] })
    };
    expect(_migrateOrphanedColumns(mockSheet)).toBe(0);
  });

  test('detects extra trailing columns when headers match', () => {
    const headers = getHeadersFromMap_(CONFIG_HEADER_MAP_);
    headers.push('Old Orphan 1', 'Old Orphan 2'); // 81 columns
    let deletedCols = [];
    const mockSheet = {
      getMaxColumns: () => headers.length,
      getRange: () => ({ getValues: () => [headers] }),
      deleteColumn: (c) => { deletedCols.push(c); }
    };
    const result = _migrateOrphanedColumns(mockSheet);
    expect(result).toBe(2);
    expect(deletedCols.length).toBe(2);
  });

  test('detects orphaned column in the middle when old headers present', () => {
    const headers = getHeadersFromMap_(CONFIG_HEADER_MAP_);
    // Insert an orphan at position 5 (old Yes/No column)
    headers.splice(4, 0, 'Yes/No (Dropdowns)');
    let deletedCols = [];
    const mockSheet = {
      getMaxColumns: () => headers.length,
      getRange: () => ({ getValues: () => [headers] }),
      deleteColumn: (c) => {
        deletedCols.push(c);
        headers.splice(c - 1, 1); // simulate deletion
      }
    };
    const result = _migrateOrphanedColumns(mockSheet);
    expect(result).toBe(1);
    expect(deletedCols).toEqual([5]); // Yes/No was at column 5
  });

  test('detects TWO orphaned columns (Yes/No + Satisfaction Form URL)', () => {
    const headers = getHeadersFromMap_(CONFIG_HEADER_MAP_);
    // Insert Yes/No at position 5 and Satisfaction Form URL after Main Contact Email
    const emailIdx = headers.indexOf('Main Contact Email');
    headers.splice(emailIdx + 1, 0, 'Satisfaction Form URL');
    headers.splice(4, 0, 'Yes/No (Dropdowns)');
    let deletedCols = [];
    const mockSheet = {
      getMaxColumns: () => headers.length,
      getRange: () => ({ getValues: () => [headers] }),
      deleteColumn: (c) => {
        deletedCols.push(c);
        headers.splice(c - 1, 1);
      }
    };
    const result = _migrateOrphanedColumns(mockSheet);
    expect(result).toBe(2);
    // After first deletion (rightmost first), indices shift
    expect(deletedCols.length).toBe(2);
  });
});
