/**
 * Tests for Non-Member Contacts CRUD functions in 21_WebDashDataService.gs
 *
 * Covers: getNonMemberContacts, addNonMemberContact, updateNonMemberContact,
 * deleteNonMemberContact — including error cases, audit logging, and formula
 * injection escaping.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '06_Maintenance.gs', '21_WebDashDataService.gs']);

// ============================================================================
// Helpers
// ============================================================================

/** Build NMC sheet header row matching NMC_HEADER_MAP_ order */
var NMC_HEADERS = [
  'Contact ID', 'First Name', 'Last Name', 'Job Title', 'Work Location',
  'Unit', 'Union Name', 'Shirt Size', 'Steward', 'Email', 'Phone',
  'Category', 'Notes'
];

function makeNmcRow(overrides) {
  var defaults = {
    contactId: 'NMC-001',
    firstName: 'Alice',
    lastName: 'Johnson',
    jobTitle: 'HR Director',
    workLocation: 'Main Office',
    unit: 'Admin',
    unionName: '',
    shirtSize: 'M',
    isSteward: 'No',
    email: 'alice@example.com',
    phone: '555-0010',
    category: 'Management',
    notes: 'Primary HR contact'
  };
  var merged = Object.assign({}, defaults, overrides || {});
  return [
    merged.contactId, merged.firstName, merged.lastName, merged.jobTitle,
    merged.workLocation, merged.unit, merged.unionName, merged.shirtSize,
    merged.isSteward, merged.email, merged.phone, merged.category, merged.notes
  ];
}

function makeNmcData(rows) {
  return [NMC_HEADERS].concat(rows);
}

function setupNmcSheet(nmcRows, opts) {
  opts = opts || {};
  var sheets = [];
  if (nmcRows !== null) {
    var data = makeNmcData(nmcRows);
    sheets.push(createMockSheet(SHEETS.NON_MEMBER_CONTACTS, data));
  }
  // Always provide a member directory sheet to avoid unrelated errors
  var memberHeaders = ['Email', 'Name', 'First Name', 'Last Name', 'Role', 'Unit'];
  sheets.push(createMockSheet(SHEETS.MEMBER_DIR, [memberHeaders]));
  var ss = createMockSpreadsheet(sheets);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

beforeEach(() => {
  if (typeof DataService !== 'undefined' && DataService._resetSSCache) {
    DataService._resetSSCache();
  }
  // Ensure logAuditEvent is a spy we can track
  global.logAuditEvent = jest.fn();
  // Ensure ensureNonMemberContactsSheet_ is available as global
  if (typeof ensureNonMemberContactsSheet_ === 'undefined') {
    global.ensureNonMemberContactsSheet_ = jest.fn(function(ss) {
      var sheet = createMockSheet(SHEETS.NON_MEMBER_CONTACTS, [NMC_HEADERS]);
      return sheet;
    });
  }
});

// ============================================================================
// getNonMemberContacts
// ============================================================================

describe('DataService.getNonMemberContacts', () => {
  test('returns empty array when NMC sheet is missing', () => {
    // No NMC sheet, just member dir
    var ss = createMockSpreadsheet([
      createMockSheet(SHEETS.MEMBER_DIR, [['Email']])
    ]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    var result = DataService.getNonMemberContacts();
    expect(result).toEqual([]);
  });

  test('returns empty array when sheet has only headers', () => {
    setupNmcSheet([]);

    var result = DataService.getNonMemberContacts();
    expect(result).toEqual([]);
  });

  test('returns contacts with all expected fields', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001', firstName: 'Alice', lastName: 'Johnson', email: 'alice@example.com' })
    ]);

    var result = DataService.getNonMemberContacts();
    expect(result.length).toBe(1);
    expect(result[0]).toHaveProperty('contactId', 'NMC-001');
    expect(result[0]).toHaveProperty('firstName', 'Alice');
    expect(result[0]).toHaveProperty('lastName', 'Johnson');
    expect(result[0]).toHaveProperty('email', 'alice@example.com');
    expect(result[0]).toHaveProperty('jobTitle');
    expect(result[0]).toHaveProperty('workLocation');
    expect(result[0]).toHaveProperty('unit');
    expect(result[0]).toHaveProperty('unionName');
    expect(result[0]).toHaveProperty('shirtSize');
    expect(result[0]).toHaveProperty('isSteward');
    expect(result[0]).toHaveProperty('phone');
    expect(result[0]).toHaveProperty('category');
    expect(result[0]).toHaveProperty('notes');
  });

  test('filters out rows with no first or last name', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001', firstName: 'Alice', lastName: 'Johnson' }),
      makeNmcRow({ contactId: 'NMC-002', firstName: '', lastName: '' })
    ]);

    var result = DataService.getNonMemberContacts();
    expect(result.length).toBe(1);
    expect(result[0].contactId).toBe('NMC-001');
  });

  test('returns multiple contacts', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001', firstName: 'Alice', lastName: 'A' }),
      makeNmcRow({ contactId: 'NMC-002', firstName: 'Bob', lastName: 'B' }),
      makeNmcRow({ contactId: 'NMC-003', firstName: 'Carol', lastName: 'C' })
    ]);

    var result = DataService.getNonMemberContacts();
    expect(result.length).toBe(3);
  });

  test('returns empty array when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.getNonMemberContacts();
    expect(result).toEqual([]);
  });
});

// ============================================================================
// addNonMemberContact
// ============================================================================

describe('DataService.addNonMemberContact', () => {
  test('returns success with generated contactId', () => {
    setupNmcSheet([]);

    var result = DataService.addNonMemberContact({ firstName: 'New', lastName: 'Contact' });
    expect(result.success).toBe(true);
    expect(result.contactId).toMatch(/^NMC-\d{3}$/);
  });

  test('generated contactId increments from existing max', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-005' }),
      makeNmcRow({ contactId: 'NMC-003' })
    ]);

    var result = DataService.addNonMemberContact({ firstName: 'Next', lastName: 'Person' });
    expect(result.success).toBe(true);
    expect(result.contactId).toBe('NMC-006');
  });

  test('calls logAuditEvent with NMC_CREATED', () => {
    setupNmcSheet([]);

    var result = DataService.addNonMemberContact({ firstName: 'Audit', lastName: 'Test' });
    expect(result.success).toBe(true);
    expect(logAuditEvent).toHaveBeenCalledWith('NMC_CREATED', expect.objectContaining({
      contactId: result.contactId
    }));
  });

  test('applies escapeForFormula to user-supplied fields', () => {
    setupNmcSheet([]);
    var spy = jest.spyOn(global, 'escapeForFormula');

    DataService.addNonMemberContact({
      firstName: '=EVIL()',
      lastName: 'Normal',
      email: 'test@example.com'
    });

    // escapeForFormula should be called for each field in the contact row
    expect(spy).toHaveBeenCalled();
    // Verify it was called with the formula-injected value
    var calls = spy.mock.calls.map(function(c) { return c[0]; });
    expect(calls).toContain('=EVIL()');

    spy.mockRestore();
  });

  test('returns error when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.addNonMemberContact({ firstName: 'Test', lastName: 'User' });
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('creates sheet via ensureNonMemberContactsSheet_ when missing', () => {
    // Only member dir sheet, no NMC sheet
    var ss = createMockSpreadsheet([
      createMockSheet(SHEETS.MEMBER_DIR, [['Email']])
    ]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.addNonMemberContact({ firstName: 'New', lastName: 'User' });
    // Should succeed — ensureNonMemberContactsSheet_ creates the sheet
    expect(result.success).toBe(true);
  });
});

// ============================================================================
// updateNonMemberContact
// ============================================================================

describe('DataService.updateNonMemberContact', () => {
  test('returns success for existing contact', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001', firstName: 'Alice' })
    ]);

    var result = DataService.updateNonMemberContact('NMC-001', { firstName: 'Alicia' });
    expect(result.success).toBe(true);
  });

  test('returns error when contact not found', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001' })
    ]);

    var result = DataService.updateNonMemberContact('NMC-999', { firstName: 'Nope' });
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('returns error when sheet is missing', () => {
    var ss = createMockSpreadsheet([
      createMockSheet(SHEETS.MEMBER_DIR, [['Email']])
    ]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.updateNonMemberContact('NMC-001', { firstName: 'Test' });
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/sheet not found/i);
  });

  test('returns error when sheet has only headers (no data rows)', () => {
    setupNmcSheet([]);

    var result = DataService.updateNonMemberContact('NMC-001', { firstName: 'Test' });
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('calls logAuditEvent with NMC_UPDATED', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001' })
    ]);

    DataService.updateNonMemberContact('NMC-001', { firstName: 'Updated', phone: '555-9999' });
    expect(logAuditEvent).toHaveBeenCalledWith('NMC_UPDATED', expect.objectContaining({
      contactId: 'NMC-001',
      fields: expect.arrayContaining(['firstName', 'phone'])
    }));
  });

  test('returns error when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.updateNonMemberContact('NMC-001', { firstName: 'Test' });
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// deleteNonMemberContact
// ============================================================================

describe('DataService.deleteNonMemberContact', () => {
  test('returns success for existing contact', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001' })
    ]);

    var result = DataService.deleteNonMemberContact('NMC-001');
    expect(result.success).toBe(true);
  });

  test('returns error when contact not found', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001' })
    ]);

    var result = DataService.deleteNonMemberContact('NMC-999');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('returns error when sheet is missing', () => {
    var ss = createMockSpreadsheet([
      createMockSheet(SHEETS.MEMBER_DIR, [['Email']])
    ]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.deleteNonMemberContact('NMC-001');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/sheet not found/i);
  });

  test('returns error when sheet has only headers', () => {
    setupNmcSheet([]);

    var result = DataService.deleteNonMemberContact('NMC-001');
    expect(result.success).toBe(false);
    expect(result.error).toMatch(/not found/i);
  });

  test('calls logAuditEvent with NMC_DELETED', () => {
    setupNmcSheet([
      makeNmcRow({ contactId: 'NMC-001' })
    ]);

    DataService.deleteNonMemberContact('NMC-001');
    expect(logAuditEvent).toHaveBeenCalledWith('NMC_DELETED', expect.objectContaining({
      contactId: 'NMC-001'
    }));
  });

  test('calls deleteRow on the correct row', () => {
    var nmcSheet = createMockSheet(SHEETS.NON_MEMBER_CONTACTS, makeNmcData([
      makeNmcRow({ contactId: 'NMC-001' }),
      makeNmcRow({ contactId: 'NMC-002' })
    ]));
    var ss = createMockSpreadsheet([
      nmcSheet,
      createMockSheet(SHEETS.MEMBER_DIR, [['Email']])
    ]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    if (DataService._resetSSCache) DataService._resetSSCache();

    DataService.deleteNonMemberContact('NMC-002');
    // NMC-002 is at data index 1 (0-based), which is sheet row 3 (1-based header + 1-based data)
    expect(nmcSheet.deleteRow).toHaveBeenCalledWith(3);
  });

  test('returns error when spreadsheet is null', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    if (DataService._resetSSCache) DataService._resetSSCache();

    var result = DataService.deleteNonMemberContact('NMC-001');
    expect(result.success).toBe(false);
  });
});
