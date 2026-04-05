/**
 * Write Safety Tests (v4.51.1)
 *
 * Validates that ALL write paths to the Member Directory sheet resolve column
 * positions from actual sheet headers at write time, rather than trusting
 * static MEMBER_COLS constants that can go stale.
 *
 * Three layers of protection:
 *   1. Unit tests for resolveColumnByHeader_ / resolveColumnsByHeader_
 *   2. Integration tests for every function that writes to Member Directory
 *   3. Structural guard: regex scan ensuring no NEW static-constant writes are added
 *
 * Background: Prior to v4.51.1, write paths used MEMBER_COLS.XXX directly.
 * These constants are initialized from MEMBER_HEADER_MAP_ array order and
 * updated by syncColumnMaps() + CacheService. When cache expired or columns
 * were appended out of canonical order, writes silently went to wrong columns,
 * corrupting the Member Directory (grievance status data in the deadline column,
 * deadline data overwriting checkboxes, dates in Open Rate, etc.).
 */

const fs = require('fs');
const path = require('path');

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load sources in GAS alphabetical load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '09_Dashboards.gs',
  '10a_SheetCreation.gs',
  '10c_FormsAndSync.gs'
]);

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Creates a mock sheet where getRange().setValues/insertCheckboxes/clearContent
 * calls are captured with their positional context.
 * Also supports getRange(headerRow, 1, 1, lastCol).getValues() for header reads.
 */
function buildTrackedSheet(name, data) {
  const sheet = createMockSheet(name, data);
  sheet.__setValuesCalls = [];
  sheet.__insertCheckboxCalls = [];
  sheet.__clearContentCalls = [];

  sheet.getLastColumn.mockReturnValue(data && data[0] ? data[0].length : 0);
  sheet.getLastRow.mockReturnValue(data ? data.length : 0);
  sheet.getName.mockReturnValue(name);

  sheet.getRange.mockImplementation((row, col, numRows, numCols) => {
    const range = {
      getValues: jest.fn(() => {
        if (data && row && numRows) {
          return data.slice(row - 1, row - 1 + (numRows || 1)).map(r =>
            numCols ? r.slice(col - 1, col - 1 + numCols) : r
          );
        }
        return [['']];
      }),
      getValue: jest.fn(() => ''),
      setValue: jest.fn(),
      setValues: jest.fn((values) => {
        sheet.__setValuesCalls.push({ row, col, numRows, numCols, values });
      }),
      insertCheckboxes: jest.fn(() => {
        sheet.__insertCheckboxCalls.push({ row, col, numRows, numCols });
      }),
      clearContent: jest.fn(() => {
        sheet.__clearContentCalls.push({ row, col, numRows, numCols });
      }),
      setFontWeight: jest.fn(function () { return this; }),
      setBackground: jest.fn(function () { return this; }),
      setFontColor: jest.fn(function () { return this; }),
      setHorizontalAlignment: jest.fn(function () { return this; }),
      setNumberFormat: jest.fn(function () { return this; }),
      setDataValidation: jest.fn(function () { return this; }),
      getRow: jest.fn(() => row),
      getColumn: jest.fn(() => col),
      getNumRows: jest.fn(() => numRows || 1),
      getNumColumns: jest.fn(() => numCols || 1)
    };
    return range;
  });

  return sheet;
}

/**
 * Builds a Member Directory header row from the canonical MEMBER_HEADER_MAP_.
 */
function buildCanonicalMemberHeaders() {
  return getHeadersFromMap_(MEMBER_HEADER_MAP_);
}

/**
 * Builds a Member Directory header row with columns SHUFFLED —
 * simulates what happens when _addMissingMemberHeaders_ appends new
 * columns to the end instead of inserting at canonical position.
 * Returns { headers: [...], posMap: { key: actualCol } }
 */
function buildShuffledMemberHeaders() {
  const canonical = getHeadersFromMap_(MEMBER_HEADER_MAP_);
  // Move 'Has Open Grievance?' to the end (simulates append)
  const shuffled = canonical.filter(h => h !== 'Has Open Grievance?');
  shuffled.push('Has Open Grievance?');
  const posMap = {};
  shuffled.forEach((h, i) => { posMap[h] = i + 1; });
  return { headers: shuffled, posMap };
}

/**
 * Builds a member data row with the right number of columns.
 */
function buildMemberRow(headers, memberId) {
  const row = new Array(headers.length).fill('');
  const idIdx = headers.indexOf('Member ID');
  if (idIdx >= 0) row[idIdx] = memberId;
  return row;
}

// ============================================================================
// 1. UNIT TESTS: resolveColumnByHeader_
// ============================================================================

describe('resolveColumnByHeader_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns correct 1-indexed position when header is found', () => {
    const headers = ['Member ID', 'First Name', 'Last Name', 'Email'];
    const data = [headers, ['M1', 'John', 'Doe', 'j@x.com']];
    const sheet = buildTrackedSheet('Test', data);

    expect(resolveColumnByHeader_(sheet, 'Email', 99)).toBe(4);
    expect(resolveColumnByHeader_(sheet, 'Member ID', 99)).toBe(1);
  });

  test('returns fallback when header is not found', () => {
    const headers = ['Member ID', 'First Name'];
    const data = [headers];
    const sheet = buildTrackedSheet('Test', data);

    expect(resolveColumnByHeader_(sheet, 'Nonexistent Column', 42)).toBe(42);
  });

  test('returns fallback for empty sheet (no columns)', () => {
    const sheet = buildTrackedSheet('Empty', []);

    expect(resolveColumnByHeader_(sheet, 'Any Header', 10)).toBe(10);
  });

  test('trims whitespace from header cells', () => {
    const headers = ['  Member ID  ', 'First Name', ' Email '];
    const data = [headers];
    const sheet = buildTrackedSheet('Test', data);

    // resolveColumnByHeader_ trims with String(headers[i]).trim()
    expect(resolveColumnByHeader_(sheet, 'Member ID', 99)).toBe(1);
    expect(resolveColumnByHeader_(sheet, 'Email', 99)).toBe(3);
  });

  test('returns 1 when both header and fallback are missing', () => {
    const sheet = buildTrackedSheet('Test', [['A']]);

    expect(resolveColumnByHeader_(sheet, 'Missing', undefined)).toBe(1);
  });

  test('handles getRange error gracefully, returns fallback', () => {
    const sheet = createMockSheet('Err', []);
    sheet.getLastColumn.mockImplementation(() => { throw new Error('quota exceeded'); });

    expect(resolveColumnByHeader_(sheet, 'Any', 7)).toBe(7);
  });

  test('handles null sheet gracefully', () => {
    expect(resolveColumnByHeader_(null, 'Any', 5)).toBe(5);
  });

  test('trims headerName input for symmetric matching', () => {
    const headers = ['Member ID', 'First Name'];
    const sheet = buildTrackedSheet('Test', [headers]);

    expect(resolveColumnByHeader_(sheet, '  First Name  ', 99)).toBe(2);
  });
});

// ============================================================================
// 2. UNIT TESTS: resolveColumnsByHeader_ (batch)
// ============================================================================

describe('resolveColumnsByHeader_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('resolves multiple headers in a single call', () => {
    const headers = ['Member ID', 'First Name', 'Open Rate %', 'Recent Contact Date'];
    const sheet = buildTrackedSheet('Test', [headers]);

    const result = resolveColumnsByHeader_(sheet, [
      { key: 'OPEN_RATE', header: 'Open Rate %', fallback: 99 },
      { key: 'CONTACT_DATE', header: 'Recent Contact Date', fallback: 98 }
    ]);

    expect(result.OPEN_RATE).toBe(3);
    expect(result.CONTACT_DATE).toBe(4);
  });

  test('returns fallback for missing headers, correct for present ones', () => {
    const headers = ['Member ID', 'Grievance Status'];
    const sheet = buildTrackedSheet('Test', [headers]);

    const result = resolveColumnsByHeader_(sheet, [
      { key: 'STATUS', header: 'Grievance Status', fallback: 30 },
      { key: 'DEADLINE', header: 'Days to Deadline', fallback: 31 }
    ]);

    expect(result.STATUS).toBe(2);       // Found at actual position
    expect(result.DEADLINE).toBe(31);     // Fallback — not in sheet
  });

  test('returns FIRST match on duplicate headers (consistent with resolveColumnByHeader_)', () => {
    const headers = ['Name', 'Email', 'Name', 'Phone'];
    const sheet = buildTrackedSheet('Test', [headers]);

    const result = resolveColumnsByHeader_(sheet, [
      { key: 'NAME', header: 'Name', fallback: 99 }
    ]);

    // Must return 1 (first match), not 3 (last match)
    expect(result.NAME).toBe(1);
  });

  test('handles empty sheet — all fallbacks', () => {
    const sheet = buildTrackedSheet('Empty', []);

    const result = resolveColumnsByHeader_(sheet, [
      { key: 'A', header: 'Alpha', fallback: 5 },
      { key: 'B', header: 'Bravo', fallback: 10 }
    ]);

    expect(result.A).toBe(5);
    expect(result.B).toBe(10);
  });

  test('reads header row only once for multiple lookups', () => {
    const headers = ['A', 'B', 'C', 'D', 'E'];
    const sheet = buildTrackedSheet('Test', [headers]);

    resolveColumnsByHeader_(sheet, [
      { key: 'X', header: 'A', fallback: 1 },
      { key: 'Y', header: 'C', fallback: 3 },
      { key: 'Z', header: 'E', fallback: 5 }
    ]);

    // Should read header row only once (getRange called once for the header read)
    expect(sheet.getRange).toHaveBeenCalledTimes(1);
  });
});

// ============================================================================
// 3. INTEGRATION: syncGrievanceToMemberDirectory writes to correct columns
// ============================================================================

describe('syncGrievanceToMemberDirectory — header-resolved writes', () => {
  beforeEach(() => jest.clearAllMocks());

  test('writes grievance data to header-resolved positions (canonical order)', () => {
    const headers = buildCanonicalMemberHeaders();
    const memberData = [headers, buildMemberRow(headers, 'MS-001')];

    // Build grievance data with one open grievance
    const gHeaders = getHeadersFromMap_(GRIEVANCE_HEADER_MAP_);
    const gRow = new Array(gHeaders.length).fill('');
    gRow[GRIEVANCE_COLS.MEMBER_ID - 1] = 'MS-001';
    gRow[GRIEVANCE_COLS.STATUS - 1] = 'Open';
    gRow[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 5;
    const grievanceData = [gHeaders, gRow];

    const memberSheet = buildTrackedSheet(SHEETS.MEMBER_DIR, memberData);
    const grievanceSheet = buildTrackedSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
    const ss = createMockSpreadsheet([memberSheet, grievanceSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    syncGrievanceToMemberDirectory();

    // Verify writes happened
    expect(memberSheet.__setValuesCalls.length).toBeGreaterThan(0);

    // The write should target the column where 'Has Open Grievance?' actually is in the header row
    const hogActualCol = headers.indexOf('Has Open Grievance?') + 1;
    const call = memberSheet.__setValuesCalls[0];
    expect(call.col).toBe(hogActualCol);
  });

  test('writes to correct columns when headers are shuffled (non-canonical order)', () => {
    const { headers, posMap } = buildShuffledMemberHeaders();
    const memberData = [headers, buildMemberRow(headers, 'MS-001')];

    const gHeaders = getHeadersFromMap_(GRIEVANCE_HEADER_MAP_);
    const gRow = new Array(gHeaders.length).fill('');
    gRow[GRIEVANCE_COLS.MEMBER_ID - 1] = 'MS-001';
    gRow[GRIEVANCE_COLS.STATUS - 1] = 'Open';
    gRow[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 3;
    const grievanceData = [gHeaders, gRow];

    const memberSheet = buildTrackedSheet(SHEETS.MEMBER_DIR, memberData);
    const grievanceSheet = buildTrackedSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
    const ss = createMockSpreadsheet([memberSheet, grievanceSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    syncGrievanceToMemberDirectory();

    expect(memberSheet.__setValuesCalls.length).toBeGreaterThan(0);

    // When headers are shuffled, Has Open Grievance? is at the END, not position 29.
    // The function should resolve from headers and write to the actual position.
    const hogActualCol = posMap['Has Open Grievance?'];
    // Find the write that contains grievance data
    // With non-consecutive columns, it should write 3 individual columns
    const writes = memberSheet.__setValuesCalls;
    const hogWrite = writes.find(w => w.col === hogActualCol);
    expect(hogWrite).toBeDefined();
  });

  test('data values are correct (Yes/Open/5 for open grievance)', () => {
    const headers = buildCanonicalMemberHeaders();
    const memberData = [headers, buildMemberRow(headers, 'MS-001')];

    const gHeaders = getHeadersFromMap_(GRIEVANCE_HEADER_MAP_);
    const gRow = new Array(gHeaders.length).fill('');
    gRow[GRIEVANCE_COLS.MEMBER_ID - 1] = 'MS-001';
    gRow[GRIEVANCE_COLS.STATUS - 1] = 'Open';
    gRow[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1] = 5;
    const grievanceData = [gHeaders, gRow];

    const memberSheet = buildTrackedSheet(SHEETS.MEMBER_DIR, memberData);
    const grievanceSheet = buildTrackedSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
    const ss = createMockSpreadsheet([memberSheet, grievanceSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    syncGrievanceToMemberDirectory();

    const call = memberSheet.__setValuesCalls[0];
    // Values: [hasOpen, status, deadline] for the 1 member
    expect(call.values[0][0]).toBe('Yes');   // Has Open Grievance
    expect(call.values[0][1]).toBe('Open');  // Grievance Status
    expect(call.values[0][2]).toBe(5);       // Days to Deadline
  });
});

// ============================================================================
// 4. INTEGRATION: repairMemberDirectoryColumns
// ============================================================================

describe('repairMemberDirectoryColumns', () => {
  beforeEach(() => jest.clearAllMocks());

  test('function exists and is callable', () => {
    expect(typeof repairMemberDirectoryColumns).toBe('function');
  });

  test('clears grievance columns before re-syncing', () => {
    const headers = buildCanonicalMemberHeaders();
    const memberData = [
      headers,
      buildMemberRow(headers, 'MS-001'),
      buildMemberRow(headers, 'MS-002')
    ];

    const gHeaders = getHeadersFromMap_(GRIEVANCE_HEADER_MAP_);
    const grievanceData = [gHeaders]; // no grievances

    const memberSheet = buildTrackedSheet(SHEETS.MEMBER_DIR, memberData);
    const grievanceSheet = buildTrackedSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
    const ss = createMockSpreadsheet([memberSheet, grievanceSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    // Mock UI
    SpreadsheetApp.getUi = jest.fn(() => ({
      alert: jest.fn(),
      ButtonSet: { OK: 0 }
    }));

    repairMemberDirectoryColumns();

    // Should have cleared the 3 grievance columns
    expect(memberSheet.__clearContentCalls.length).toBeGreaterThanOrEqual(3);

    // Should have re-applied checkboxes to Start Grievance
    expect(memberSheet.__insertCheckboxCalls.length).toBeGreaterThanOrEqual(1);
    const cbCall = memberSheet.__insertCheckboxCalls[0];
    const sgActualCol = headers.indexOf('Start Grievance') + 1;
    expect(cbCall.col).toBe(sgActualCol);
  });

  test('does not clear columns when Grievance Log is missing', () => {
    const headers = buildCanonicalMemberHeaders();
    const memberData = [headers, buildMemberRow(headers, 'MS-001')];

    const memberSheet = buildTrackedSheet(SHEETS.MEMBER_DIR, memberData);
    // No grievance sheet in the spreadsheet
    const ss = createMockSpreadsheet([memberSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    SpreadsheetApp.getUi = jest.fn(() => ({
      alert: jest.fn(),
      ButtonSet: { OK: 0 }
    }));

    repairMemberDirectoryColumns();

    // Must NOT have cleared any columns since there's no Grievance Log to re-sync from
    expect(memberSheet.__clearContentCalls.length).toBe(0);
  });
});

// ============================================================================
// 5. STRUCTURAL GUARD: Comprehensive static-constant write scanner
// ============================================================================

describe('Structural guard: no unsafe static-constant writes (comprehensive)', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');
  const allGsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs'));

  // All column constant systems that could cause write corruption
  const CONSTANT_SYSTEMS = [
    'MEMBER_COLS', 'GRIEVANCE_COLS', 'CONFIG_COLS', 'NMC_COLS',
    'CHECKLIST_COLS', 'SURVEY_TRACKING_COLS', 'AUDIT_LOG_COLS',
    'STEWARD_PERF_COLS', 'MEETING_CHECKIN_COLS', 'FEEDBACK_COLS',
    'RESOURCES_COLS', 'RESOURCE_CONFIG_COLS', 'SURVEY_VAULT_COLS'
  ];
  const CONSTANTS_PATTERN = CONSTANT_SYSTEMS.join('|');

  // Write methods that modify cell content
  const WRITE_METHODS = 'setValues|setValue|insertCheckboxes|clearContent';

  // Allowlist: patterns that are safe despite matching
  // { file: string|null, linePattern: RegExp, reason: string }
  const allowlist = [
    // Dev-only file excluded from prod
    { file: '07_DevTools.gs', linePattern: /.*/, reason: 'dev-only, PROD_EXCLUDE' },
    // Cosmetic operations that don't write data
    { file: null, linePattern: /setColumnWidth|hideColumns|autoResizeColumns|setFrozenRows|setNumberFormat|setBackground|setFontWeight|setFontColor|setHorizontalAlignment/, reason: 'cosmetic/formatting only' },
    // Full-row read-modify-write (internally consistent within execution)
    { file: '02_DataManagers.gs', linePattern: /setValues\(\[rowData\]\)/, reason: 'full-row write — col_/setCol_ consistent with same-execution read' },
    { file: '02_DataManagers.gs', linePattern: /appendRow\(rowData\)/, reason: 'full-row append — setCol_ consistent' },
    { file: '05_Integrations.gs', linePattern: /appendRow\(rowData\)/, reason: 'full-row append — setCol_ consistent' },
    // MEMBER_COLS.MEMBER_ID is always column 1 — never shifts
    { file: '02_DataManagers.gs', linePattern: /MEMBER_COLS\.MEMBER_ID/, reason: 'MEMBER_ID always column 1' },
    // Single-cell edit handlers (one row at a time, user-triggered)
    { file: '10_Main.gs', linePattern: /\.setValue\(/, reason: 'single-cell edit handler' },
    { file: '03_UIComponents.gs', linePattern: /\.setValue\(/, reason: 'single-cell UI action' },
    // Single-cell IS_STEWARD toggle (interactive, one cell)
    { file: '02_DataManagers.gs', linePattern: /MEMBER_COLS\.IS_STEWARD/, reason: 'single-cell steward toggle' },
    // Single-cell grievance field updates (interactive edit handlers)
    { file: '09_Dashboards.gs', linePattern: /e\.range/, reason: 'onEdit handler context' },
    { file: '12_Features.gs', linePattern: /GRIEVANCE_COLS\.CHECKLIST_PROGRESS\)\.setValue/, reason: 'single-cell progress update' },
    // Drive folder URL single-cell writes
    { file: null, linePattern: /DRIVE_FOLDER_ID|DRIVE_FOLDER_URL|\.getId\(\)|\.getUrl\(\)/, reason: 'single-cell Drive URL write' },
    // E-sign full-row creation (same pattern as addMember — internally consistent)
    { file: '05_Integrations.gs', linePattern: /setCol_/, reason: 'full-row build with setCol_' },
    // Conditional format / validation ranges (wrong column = cosmetic, not data corruption)
    { file: null, linePattern: /setDataValidation|ConditionalFormatRule|setConditionalFormatRules/, reason: 'validation/formatting' },
    // TOTAL_GRIEVANCES/ACTIVE_GRIEVANCES — guarded by existence check, may not exist
    { file: '02_DataManagers.gs', linePattern: /MEMBER_COLS\.TOTAL_GRIEVANCES|MEMBER_COLS\.ACTIVE_GRIEVANCES/, reason: 'guarded by existence check' },
    // PIN writes (single-cell, PIN_CONFIG resolves internally)
    { file: '13_MemberSelfService.gs', linePattern: /PIN_CONFIG|PIN_HASH|pinCol/, reason: 'single-cell PIN write' },
    // Signature status single-cell writes
    { file: '05_Integrations.gs', linePattern: /SIGNATURE_STATUS/, reason: 'single-cell signature status' },
    // Action type default value write uses dynamically-resolved actionTypeCol variable
    { file: '12_Features.gs', linePattern: /actionTypeCol/, reason: 'uses dynamically resolved variable' },
    // Single-cell config writes (Config sheet, not member/grievance data)
    { file: null, linePattern: /CONFIG_COLS\.[A-Z_]+\)\.setValue|CONFIG_COLS\.[A-Z_]+\)\.clearContent/, reason: 'single-cell config write' },
    // Single-cell survey vault writes (review status, not bulk data)
    { file: null, linePattern: /SURVEY_VAULT_COLS\.\w+\)\.setValue/, reason: 'single-cell vault write' },
    // Single-cell survey tracking writes
    { file: null, linePattern: /SURVEY_TRACKING_COLS\.\w+\)\.setValue/, reason: 'single-cell tracking write' },
    // Single-cell resource visibility toggles
    { file: null, linePattern: /RESOURCES_COLS\.VISIBLE\)\.setValue/, reason: 'single-cell resource visibility' },
    // Single-cell grievance field writes (one row at a time, interactive)
    { file: '02_DataManagers.gs', linePattern: /GRIEVANCE_COLS\.\w+\)\.setValue/, reason: 'single-cell grievance field write' },
    // CommandHub OCR-to-grievance single-cell writes
    { file: '11_CommandHub.gs', linePattern: /GRIEVANCE_COLS\.\w+\)\.setValue/, reason: 'single-cell OCR import write' },
    // Checklist sheet writes (separate sheet, single-cell/checkbox operations)
    { file: null, linePattern: /CHECKLIST_COLS\.\w+/, reason: 'checklist sheet — separate from Member/Grievance' },
    // Meeting check-in status writes (separate sheet, single-cell)
    { file: null, linePattern: /MEETING_CHECKIN_COLS\.\w+/, reason: 'meeting check-in sheet — separate sheet' },
    // Steward performance writes (separate sheet)
    { file: null, linePattern: /STEWARD_PERF_COLS\.\w+/, reason: 'steward perf sheet — separate sheet' },
    // Feedback sheet writes (separate sheet)
    { file: null, linePattern: /FEEDBACK_COLS\.\w+/, reason: 'feedback sheet — separate sheet' },
    // Non-Member Contacts writes (separate sheet, single-cell per field)
    { file: null, linePattern: /NMC_COLS\.\w+/, reason: 'NMC sheet — separate from Member/Grievance' },
  ];

  function isAllowlisted(file, line) {
    return allowlist.some(a =>
      (a.file === null || a.file === file) && a.linePattern.test(line)
    );
  }

  test('no bulk writes using static column constants across ALL source files', () => {
    const violations = [];

    allGsFiles.forEach(file => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const lines = content.split('\n');

      lines.forEach((line, idx) => {
        const trimmed = line.trim();
        if (trimmed.startsWith('//') || trimmed.startsWith('*')) return;

        // Pattern: getRange(x, CONSTANT.Y, ...).writeMethod(
        const directPattern = new RegExp(
          'getRange\\([^)]*(' + CONSTANTS_PATTERN + ')\\.[A-Z_]+[^)]*\\)\\s*\\.\\s*(' + WRITE_METHODS + ')\\s*\\('
        );
        if (directPattern.test(line) && !isAllowlisted(file, line)) {
          violations.push(`  ${file}:${idx + 1} [DIRECT]: ${trimmed.substring(0, 130)}`);
        }
      });
    });

    if (violations.length > 0) {
      throw new Error(
        `UNSAFE WRITES: ${violations.length} write(s) using static column constants found.\n` +
        `Each must use resolveColumnByHeader_() or resolveColumnsByHeader_() instead.\n\n` +
        violations.join('\n')
      );
    }
  });

  test('no multi-column span writes (2+ cols) anchored on static constants', () => {
    const violations = [];
    const spanPattern = new RegExp(
      'getRange\\([^,]+,\\s*(' + CONSTANTS_PATTERN + ')\\.[A-Z_]+\\s*,[^,]+,\\s*[2-9]\\s*\\)\\s*\\.\\s*(' + WRITE_METHODS + '|getValues)'
    );

    allGsFiles.forEach(file => {
      if (file === '07_DevTools.gs') return;
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const lines = content.split('\n');

      lines.forEach((line, idx) => {
        const trimmed = line.trim();
        if (trimmed.startsWith('//') || trimmed.startsWith('*')) return;
        if (spanPattern.test(line)) {
          violations.push(`  ${file}:${idx + 1}: ${trimmed.substring(0, 130)}`);
        }
      });
    });

    if (violations.length > 0) {
      throw new Error(
        `DANGEROUS: Multi-column span writes anchored on static constants.\n` +
        violations.join('\n')
      );
    }
  });
});

// ============================================================================
// 7. GUARD: Dues status includes Non Member
// ============================================================================

describe('Dues status seed defaults', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');

  test('10a_SheetCreation.gs seeds include Non Member', () => {
    const content = fs.readFileSync(path.join(srcDir, '10a_SheetCreation.gs'), 'utf8');
    expect(content).toContain("'Non Member'");
  });

  test('07_DevTools.gs seeds include Non Member', () => {
    const filePath = path.join(srcDir, '07_DevTools.gs');
    if (!fs.existsSync(filePath)) return; // SolidBase may exclude
    const content = fs.readFileSync(filePath, 'utf8');
    expect(content).toContain("'Non Member'");
  });
});

// ============================================================================
// 8. GUARD: resolveColumnByHeader_ and resolveColumnsByHeader_ are globally available
// ============================================================================

describe('Header resolution functions are defined', () => {
  test('resolveColumnByHeader_ is a function', () => {
    expect(typeof resolveColumnByHeader_).toBe('function');
  });

  test('resolveColumnsByHeader_ is a function', () => {
    expect(typeof resolveColumnsByHeader_).toBe('function');
  });

  test('repairMemberDirectoryColumns is a function', () => {
    expect(typeof repairMemberDirectoryColumns).toBe('function');
  });
});
