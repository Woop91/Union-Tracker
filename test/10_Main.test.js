/**
 * Tests for 10_Main.gs
 *
 * Covers addNewMember column mapping, member ID generation,
 * startGrievanceForMember sanitization, and trigger management.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load in GAS load order — we need constants defined before Main
loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs']);

// Derive 0-indexed constants for row position validation
const MEMBER_COLUMNS = Object.fromEntries(
  Object.entries(MEMBER_COLS).map(([k, v]) => [k, v - 1])
);

// We can't fully load 10_Main.gs because it has heavy GAS dependencies,
// so we test the key logic patterns by extracting them into isolated tests.

// ============================================================================
// addNewMember row construction logic
// ============================================================================

describe('addNewMember row construction', () => {
  // Simulate the logic from addNewMember to verify column mapping
  function buildMemberRow(memberData) {
    var rowData = new Array(MEMBER_COLS.QUICK_ACTIONS).fill('');
    rowData[MEMBER_COLS.MEMBER_ID - 1] = 'MEM-TEST123';
    rowData[MEMBER_COLS.FIRST_NAME - 1] = memberData.firstName || '';
    rowData[MEMBER_COLS.LAST_NAME - 1] = memberData.lastName || '';
    rowData[MEMBER_COLS.JOB_TITLE - 1] = memberData.jobTitle || '';
    rowData[MEMBER_COLS.WORK_LOCATION - 1] = memberData.workLocation || '';
    rowData[MEMBER_COLS.UNIT - 1] = memberData.unit || '';
    rowData[MEMBER_COLS.EMAIL - 1] = memberData.email || '';
    rowData[MEMBER_COLS.PHONE - 1] = memberData.phone || '';
    rowData[MEMBER_COLS.IS_STEWARD - 1] = 'No';
    rowData[MEMBER_COLS.RECENT_CONTACT_DATE - 1] = 'TODAY';
    return rowData;
  }

  test('row array has correct length (QUICK_ACTIONS columns)', () => {
    const row = buildMemberRow({ firstName: 'John', lastName: 'Smith' });
    expect(row.length).toBe(MEMBER_COLS.QUICK_ACTIONS);
    expect(row.length).toBe(33);
  });

  test('Member ID is in correct position', () => {
    const row = buildMemberRow({});
    expect(row[0]).toBe('MEM-TEST123'); // Column A = index 0
    expect(row[MEMBER_COLUMNS.MEMBER_ID]).toBe('MEM-TEST123');
  });

  test('First Name is in correct position', () => {
    const row = buildMemberRow({ firstName: 'Jane' });
    expect(row[MEMBER_COLUMNS.FIRST_NAME]).toBe('Jane');
  });

  test('Last Name is in correct position', () => {
    const row = buildMemberRow({ lastName: 'Doe' });
    expect(row[MEMBER_COLUMNS.LAST_NAME]).toBe('Doe');
  });

  test('Email is in correct position', () => {
    const row = buildMemberRow({ email: 'test@example.com' });
    expect(row[MEMBER_COLUMNS.EMAIL]).toBe('test@example.com');
  });

  test('Phone is in correct position', () => {
    const row = buildMemberRow({ phone: '555-123-4567' });
    expect(row[MEMBER_COLUMNS.PHONE]).toBe('555-123-4567');
  });

  test('Work Location is in correct position', () => {
    const row = buildMemberRow({ workLocation: 'Main Station' });
    expect(row[MEMBER_COLUMNS.WORK_LOCATION]).toBe('Main Station');
  });

  test('Unit is in correct position', () => {
    const row = buildMemberRow({ unit: 'Field Ops' });
    expect(row[MEMBER_COLUMNS.UNIT]).toBe('Field Ops');
  });

  test('Is Steward defaults to No', () => {
    const row = buildMemberRow({});
    expect(row[MEMBER_COLUMNS.IS_STEWARD]).toBe('No');
  });

  test('Recent Contact Date is set', () => {
    const row = buildMemberRow({});
    expect(row[MEMBER_COLUMNS.RECENT_CONTACT_DATE]).toBe('TODAY');
  });

  test('all other fields default to empty string', () => {
    const row = buildMemberRow({});
    // Check that fields not explicitly set are empty
    expect(row[MEMBER_COLUMNS.JOB_TITLE]).toBe('');
    expect(row[MEMBER_COLUMNS.SUPERVISOR]).toBe('');
    expect(row[MEMBER_COLUMNS.MANAGER]).toBe('');
    expect(row[MEMBER_COLUMNS.COMMITTEES]).toBe('');
    expect(row[MEMBER_COLUMNS.CONTACT_NOTES]).toBe('');
  });
});

// ============================================================================
// Member ID generation
// ============================================================================

describe('Member ID generation', () => {
  test('Date.now().toString(36) produces unique-ish IDs', () => {
    const id1 = Date.now().toString(36).toUpperCase();
    // Small delay to ensure different timestamp
    const id2 = (Date.now() + 1).toString(36).toUpperCase();
    expect(id1).not.toBe(id2);
  });

  test('generated ID format is MEM-{base36timestamp}', () => {
    const timestamp = Date.now().toString(36).toUpperCase();
    const newId = `MEM-${timestamp}`;
    expect(newId).toMatch(/^MEM-[A-Z0-9]+$/);
  });

  test('generated ID length is reasonable', () => {
    const timestamp = Date.now().toString(36).toUpperCase();
    const newId = `MEM-${timestamp}`;
    expect(newId.length).toBeGreaterThan(6);
    expect(newId.length).toBeLessThan(20);
  });
});

// ============================================================================
// startGrievanceForMember sanitization
// ============================================================================

describe('startGrievanceForMember sanitization', () => {
  // Simulate the sanitization logic
  function sanitizeForScript(value) {
    return String(value || '').replace(/['"\\<>&]/g, '');
  }

  test('strips single quotes', () => {
    expect(sanitizeForScript("O'Brien")).toBe('OBrien');
  });

  test('strips double quotes', () => {
    expect(sanitizeForScript('say "hello"')).toBe('say hello');
  });

  test('strips backslashes', () => {
    expect(sanitizeForScript('a\\b')).toBe('ab');
  });

  test('strips HTML special characters', () => {
    expect(sanitizeForScript('<script>alert(1)</script>')).toBe('scriptalert(1)/script');
  });

  test('strips ampersand', () => {
    expect(sanitizeForScript('A&B')).toBe('AB');
  });

  test('handles empty/null input', () => {
    expect(sanitizeForScript('')).toBe('');
    expect(sanitizeForScript(null)).toBe('');
    expect(sanitizeForScript(undefined)).toBe('');
  });

  test('preserves normal member IDs', () => {
    expect(sanitizeForScript('MEM-ABC123')).toBe('MEM-ABC123');
  });
});

// ============================================================================
// CSV export escaping logic
// ============================================================================

describe('CSV export escaping', () => {
  // Simulate the CSV escaping logic from exportMembersToCSV
  function escapeCSVField(value) {
    var str = String(value || '');
    if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  }

  test('wraps fields with commas in double quotes', () => {
    expect(escapeCSVField('Smith, John')).toBe('"Smith, John"');
  });

  test('escapes double quotes by doubling them', () => {
    expect(escapeCSVField('He said "hello"')).toBe('"He said ""hello"""');
  });

  test('wraps fields with newlines', () => {
    expect(escapeCSVField('line1\nline2')).toBe('"line1\nline2"');
  });

  test('leaves simple values unwrapped', () => {
    expect(escapeCSVField('John')).toBe('John');
    expect(escapeCSVField('12345')).toBe('12345');
  });

  test('handles empty/null input', () => {
    expect(escapeCSVField('')).toBe('');
    expect(escapeCSVField(null)).toBe('');
  });
});

// ============================================================================
// handleGrievanceEdit - Next Action Due columns
// ============================================================================

describe('handleGrievanceEdit column references', () => {
  test('status and date columns used for Next Action Due update are valid', () => {
    var statusAndDateCols = [
      GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.CURRENT_STEP,
      GRIEVANCE_COLS.DATE_FILED, GRIEVANCE_COLS.STEP1_RCVD,
      GRIEVANCE_COLS.STEP2_APPEAL_FILED, GRIEVANCE_COLS.STEP2_RCVD,
      GRIEVANCE_COLS.STEP3_APPEAL_FILED, GRIEVANCE_COLS.DATE_CLOSED
    ];

    statusAndDateCols.forEach(col => {
      expect(col).toBeGreaterThan(0);
      expect(col).toBeLessThanOrEqual(GRIEVANCE_COLS.QUICK_ACTIONS);
    });
  });

  test('NEXT_ACTION_DUE column is within range', () => {
    expect(GRIEVANCE_COLS.NEXT_ACTION_DUE).toBe(20);
  });

  test('step date columns map correctly for deadline recalculation', () => {
    const stepDateColumns = [
      GRIEVANCE_COLS.DATE_FILED,
      GRIEVANCE_COLS.STEP2_APPEAL_FILED,
      GRIEVANCE_COLS.STEP3_APPEAL_FILED
    ];
    const dueColumns = [
      GRIEVANCE_COLS.STEP1_DUE,
      GRIEVANCE_COLS.STEP2_DUE,
      GRIEVANCE_COLS.STEP3_APPEAL_DUE
    ];

    // Each step date should have a corresponding due column
    expect(stepDateColumns.length).toBe(dueColumns.length);
    stepDateColumns.forEach(col => expect(col).toBeGreaterThan(0));
    dueColumns.forEach(col => expect(col).toBeGreaterThan(0));
  });
});

// ============================================================================
// removeTriggers - scoped to dailyTrigger only
// ============================================================================

describe('removeTriggers scoping', () => {
  test('SHEETS has no INTERACTIVE key (removed in bug fix)', () => {
    expect(SHEETS.INTERACTIVE).toBeUndefined();
  });
});
