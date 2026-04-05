/**
 * Tests for 10a_SheetCreation.gs
 * Covers sheet creation, config seeding, header validation,
 * and dashboard setup functions.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
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
    'populateConfigFromSheetData',
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

  test('CONFIG_HEADER_MAP_ is non-empty', () => {
    // Redundant hardcoded count removed — dynamic assertion above is authoritative
    expect(CONFIG_HEADER_MAP_.length).toBeGreaterThan(0);
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
    let batchDeleted = [];
    const mockSheet = {
      getMaxColumns: () => headers.length,
      getRange: () => ({ getValues: () => [headers] }),
      deleteColumn: (c) => { deletedCols.push(c); },
      deleteColumns: (start, count) => { batchDeleted.push({ start, count }); }
    };
    const result = _migrateOrphanedColumns(mockSheet);
    expect(result).toBe(2);
    // May use batch deleteColumns for contiguous ranges or individual deleteColumn
    const totalDeleted = deletedCols.length + batchDeleted.reduce((sum, b) => sum + b.count, 0);
    expect(totalDeleted).toBe(2);
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

// ============================================================================
// getMemberHeaders behavioral tests
// ============================================================================

describe('getMemberHeaders behavioral', () => {
  test('returns headers matching MEMBER_HEADER_MAP_ length', () => {
    const headers = getMemberHeaders();
    expect(headers.length).toBe(MEMBER_HEADER_MAP_.length);
  });

  test('headers are all non-empty strings', () => {
    const headers = getMemberHeaders();
    headers.forEach((h, i) => {
      expect(typeof h).toBe('string');
      expect(h.trim().length).toBeGreaterThan(0);
    });
  });

  test('first header is Member ID', () => {
    const headers = getMemberHeaders();
    expect(headers[0]).toBe('Member ID');
  });

  test('contains all essential member fields in order', () => {
    const headers = getMemberHeaders();
    // First few must be identity fields
    expect(headers[0]).toBe('Member ID');
    expect(headers[1]).toBe('First Name');
    expect(headers[2]).toBe('Last Name');
    // Email and Phone are present
    expect(headers).toContain('Email');
    expect(headers).toContain('Phone');
    // Employment fields
    expect(headers).toContain('Job Title');
    expect(headers).toContain('Work Location');
    // Steward/role fields
    expect(headers).toContain('Is Steward');
    expect(headers).toContain('Assigned Steward');
    // Grievance integration fields
    expect(headers).toContain('Has Open Grievance?');
    expect(headers).toContain('Grievance Status');
    expect(headers).toContain('Days to Deadline');
    expect(headers).toContain('Start Grievance');
  });

  test('contains no duplicate headers', () => {
    const headers = getMemberHeaders();
    const unique = new Set(headers);
    expect(unique.size).toBe(headers.length);
  });

  test('headers match MEMBER_HEADER_MAP_ entries in order', () => {
    const headers = getMemberHeaders();
    MEMBER_HEADER_MAP_.forEach((entry, i) => {
      expect(headers[i]).toBe(entry.header);
    });
  });
});

// ============================================================================
// getGrievanceHeaders behavioral tests
// ============================================================================

describe('getGrievanceHeaders behavioral', () => {
  test('returns headers matching GRIEVANCE_HEADER_MAP_ length', () => {
    const headers = getGrievanceHeaders();
    expect(headers.length).toBe(GRIEVANCE_HEADER_MAP_.length);
  });

  test('headers are all non-empty strings', () => {
    const headers = getGrievanceHeaders();
    headers.forEach(h => {
      expect(typeof h).toBe('string');
      expect(h.trim().length).toBeGreaterThan(0);
    });
  });

  test('first header is Grievance ID', () => {
    const headers = getGrievanceHeaders();
    expect(headers[0]).toBe('Grievance ID');
  });

  test('contains all essential grievance fields', () => {
    const headers = getGrievanceHeaders();
    expect(headers).toContain('Grievance ID');
    expect(headers).toContain('Member ID');
    expect(headers).toContain('Status');
    expect(headers).toContain('Current Step');
    expect(headers).toContain('Incident Date');
    expect(headers).toContain('Filing Deadline');
    expect(headers).toContain('Date Filed');
    expect(headers).toContain('Resolution');
    expect(headers).toContain('Days Open');
    expect(headers).toContain('Days to Deadline');
    expect(headers).toContain('Assigned Steward');
  });

  test('contains step timeline fields in correct order', () => {
    const headers = getGrievanceHeaders();
    const step1Due = headers.indexOf('Step I Due');
    const step1Rcvd = headers.indexOf('Step I Rcvd');
    const step2AppealDue = headers.indexOf('Step II Appeal Due');
    const step3AppealDue = headers.indexOf('Step III Appeal Due');
    // All must exist
    expect(step1Due).toBeGreaterThan(-1);
    expect(step1Rcvd).toBeGreaterThan(-1);
    expect(step2AppealDue).toBeGreaterThan(-1);
    expect(step3AppealDue).toBeGreaterThan(-1);
    // Must be in order
    expect(step1Due).toBeLessThan(step1Rcvd);
    expect(step1Rcvd).toBeLessThan(step2AppealDue);
    expect(step2AppealDue).toBeLessThan(step3AppealDue);
  });

  test('contains no duplicate headers', () => {
    const headers = getGrievanceHeaders();
    const unique = new Set(headers);
    expect(unique.size).toBe(headers.length);
  });

  test('headers match GRIEVANCE_HEADER_MAP_ entries in order', () => {
    const headers = getGrievanceHeaders();
    GRIEVANCE_HEADER_MAP_.forEach((entry, i) => {
      expect(headers[i]).toBe(entry.header);
    });
  });
});

// ============================================================================
// createMemberDirectory creates sheet with correct headers
// ============================================================================

describe('createMemberDirectory creates sheet with correct headers', () => {
  test('writes all member headers to a new empty sheet', () => {
    const ss = createMockSpreadsheet([]);
    // createMemberDirectory calls getOrCreateSheet which calls ss.getSheetByName
    // then ss.insertSheet (since sheet doesn't exist). The insertSheet mock returns
    // a createMockSheet. We need to capture what setValues is called with.
    let capturedHeaders = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      // The sheet.getLastRow() returns 1 by default (mock), but createMemberDirectory
      // checks if (sheet.getLastRow() <= 1) to decide if it's a new sheet.
      // Default mock returns 1 which means "new sheet" path.
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 1 && col === 1 && numCols > 10) {
          // This is the header write call
          const origSetValues = range.setValues;
          range.setValues = jest.fn(function(vals) {
            capturedHeaders = vals[0];
            return range; // chain
          });
        }
        return range;
      });
      return sheet;
    });

    expect(() => createMemberDirectory(ss)).not.toThrow();
    expect(capturedHeaders).not.toBeNull();
    expect(capturedHeaders.length).toBe(45);
    expect(capturedHeaders[0]).toBe('Member ID');
    expect(capturedHeaders[1]).toBe('First Name');
    expect(capturedHeaders[2]).toBe('Last Name');
    expect(capturedHeaders).toContain('Email');
    expect(capturedHeaders).toContain('Is Steward');
  });

  test('sets frozen rows on the sheet', () => {
    const ss = createMockSpreadsheet([]);
    let frozenRowsCalled = false;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      sheet.setFrozenRows = jest.fn(() => { frozenRowsCalled = true; });
      return sheet;
    });

    createMemberDirectory(ss);
    expect(frozenRowsCalled).toBe(true);
  });
});

// ============================================================================
// createGrievanceLog creates sheet with correct headers
// ============================================================================

describe('createGrievanceLog creates sheet with correct headers', () => {
  test('writes all grievance headers to a new empty sheet', () => {
    const ss = createMockSpreadsheet([]);
    let capturedHeaders = null;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      const origGetRange = sheet.getRange;
      sheet.getRange = jest.fn(function(row, col, numRows, numCols) {
        const range = origGetRange(row, col, numRows, numCols);
        if (row === 1 && col === 1 && numCols > 10) {
          range.setValues = jest.fn(function(vals) {
            capturedHeaders = vals[0];
            return range;
          });
        }
        return range;
      });
      return sheet;
    });

    expect(() => createGrievanceLog(ss)).not.toThrow();
    expect(capturedHeaders).not.toBeNull();
    expect(capturedHeaders.length).toBe(46);
    expect(capturedHeaders[0]).toBe('Grievance ID');
    expect(capturedHeaders[1]).toBe('Member ID');
    expect(capturedHeaders).toContain('Status');
    expect(capturedHeaders).toContain('Current Step');
    expect(capturedHeaders).toContain('Resolution');
  });

  test('sets frozen rows on the sheet', () => {
    const ss = createMockSpreadsheet([]);
    let frozenRowsCalled = false;
    const origInsertSheet = ss.insertSheet;
    ss.insertSheet = jest.fn(name => {
      const sheet = origInsertSheet(name);
      sheet.setFrozenRows = jest.fn(() => { frozenRowsCalled = true; });
      return sheet;
    });

    createGrievanceLog(ss);
    expect(frozenRowsCalled).toBe(true);
  });
});

// ============================================================================
// _addMissingMemberHeaders_ detects and appends missing columns
// ============================================================================

describe('_addMissingMemberHeaders_ behavioral', () => {
  test('returns empty array when all headers already present', () => {
    const allHeaders = getMemberHeaders();
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, [allHeaders]);
    const result = _addMissingMemberHeaders_(sheet);
    expect(result).toEqual([]);
  });

  test('detects missing headers and returns their names', () => {
    // Provide only the first 3 headers — the rest should be detected as missing
    const partialHeaders = getMemberHeaders().slice(0, 3);
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, [partialHeaders]);
    const result = _addMissingMemberHeaders_(sheet);
    expect(result.length).toBe(42); // 45 - 3 = 42 missing
    expect(result).toContain('Email');
    expect(result).toContain('Phone');
    expect(result).toContain('Is Steward');
  });

  test('is case-insensitive when matching existing headers', () => {
    // Provide headers in different case
    const mixedCaseHeaders = ['member id', 'FIRST NAME', 'Last Name'];
    const sheet = createMockSheet(SHEETS.MEMBER_DIR, [mixedCaseHeaders]);
    const result = _addMissingMemberHeaders_(sheet);
    // These 3 should NOT be in the missing list (case-insensitive match)
    expect(result).not.toContain('Member ID');
    expect(result).not.toContain('First Name');
    expect(result).not.toContain('Last Name');
    expect(result.length).toBe(42); // 45 - 3 = 42 missing
  });
});

// ============================================================================
// _addMissingGrievanceHeaders_ detects and appends missing columns
// ============================================================================

describe('_addMissingGrievanceHeaders_ behavioral', () => {
  test('returns empty array when all headers already present', () => {
    const allHeaders = getGrievanceHeaders();
    const sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [allHeaders]);
    const result = _addMissingGrievanceHeaders_(sheet);
    expect(result).toEqual([]);
  });

  test('detects missing headers and returns their names', () => {
    const partialHeaders = getGrievanceHeaders().slice(0, 5);
    const sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [partialHeaders]);
    const result = _addMissingGrievanceHeaders_(sheet);
    expect(result.length).toBe(41); // 46 - 5 = 41 missing
    expect(result).toContain('Resolution');
    expect(result).toContain('Days Open');
    expect(result).toContain('Assigned Steward');
  });

  test('calls setValue for each missing header', () => {
    const partialHeaders = getGrievanceHeaders().slice(0, 2);
    const sheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [partialHeaders]);
    const result = _addMissingGrievanceHeaders_(sheet);
    // Each missing header triggers a getRange().setValue() call
    expect(result.length).toBe(44); // 46 - 2 = 44
    // Verify getRange was called for each missing header
    // (initial read + 44 missing header writes)
    expect(sheet.getRange.mock.calls.length).toBeGreaterThanOrEqual(44);
  });
});

// ============================================================================
// MEMBER_COLS and GRIEVANCE_COLS positional integrity
// ============================================================================

describe('MEMBER_COLS positional integrity', () => {
  test('MEMBER_ID is column 1', () => {
    expect(MEMBER_COLS.MEMBER_ID).toBe(1);
  });

  test('FIRST_NAME is column 2', () => {
    expect(MEMBER_COLS.FIRST_NAME).toBe(2);
  });

  test('LAST_NAME is column 3', () => {
    expect(MEMBER_COLS.LAST_NAME).toBe(3);
  });

  test('column count matches header count', () => {
    const maxCol = Math.max(...Object.values(MEMBER_COLS).filter(v => typeof v === 'number'));
    expect(maxCol).toBe(MEMBER_HEADER_MAP_.length);
  });

  test('alias LOCATION maps to WORK_LOCATION', () => {
    expect(MEMBER_COLS.LOCATION).toBe(MEMBER_COLS.WORK_LOCATION);
  });

  test('alias DIRECTOR maps to MANAGER', () => {
    expect(MEMBER_COLS.DIRECTOR).toBe(MEMBER_COLS.MANAGER);
  });
});

describe('GRIEVANCE_COLS positional integrity', () => {
  test('GRIEVANCE_ID is column 1', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_ID).toBe(1);
  });

  test('MEMBER_ID is column 2', () => {
    expect(GRIEVANCE_COLS.MEMBER_ID).toBe(2);
  });

  test('column count matches header count', () => {
    const maxCol = Math.max(...Object.values(GRIEVANCE_COLS).filter(v => typeof v === 'number'));
    expect(maxCol).toBe(GRIEVANCE_HEADER_MAP_.length);
  });

  test('alias GRIEVANCE_STATUS maps to STATUS', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_STATUS).toBe(GRIEVANCE_COLS.STATUS);
  });

  test('alias GRIEVANCE_STEP maps to CURRENT_STEP', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_STEP).toBe(GRIEVANCE_COLS.CURRENT_STEP);
  });
});
