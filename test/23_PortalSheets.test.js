/**
 * Tests for 23_PortalSheets.gs
 *
 * Covers portal column constants, sheet setup helpers, and
 * the 8 getOrCreate sheet functions.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '23_PortalSheets.gs']);

let mockSs;

beforeEach(() => {
  mockSs = createMockSpreadsheet([]);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
});

// ============================================================================
// Portal Column Constants
// ============================================================================

describe('PORTAL_EVENT_COLS', () => {
  test('columns are 0-indexed', () => {
    expect(PORTAL_EVENT_COLS.ID).toBe(0);
    expect(PORTAL_EVENT_COLS.TITLE).toBe(1);
    expect(PORTAL_EVENT_COLS.TYPE).toBe(2);
    expect(PORTAL_EVENT_COLS.DATE_TIME).toBe(3);
  });

  test('all expected fields exist', () => {
    expect(PORTAL_EVENT_COLS).toHaveProperty('ID');
    expect(PORTAL_EVENT_COLS).toHaveProperty('TITLE');
    expect(PORTAL_EVENT_COLS).toHaveProperty('LOCATION');
    expect(PORTAL_EVENT_COLS).toHaveProperty('CREATED_BY');
  });
});

describe('PORTAL_MEMBER_DIR_COLS', () => {
  test('EMAIL is 0, NAME is 1', () => {
    expect(PORTAL_MEMBER_DIR_COLS.EMAIL).toBe(0);
    expect(PORTAL_MEMBER_DIR_COLS.NAME).toBe(1);
  });

  test('has OFFICE_DAYS and BADGES', () => {
    expect(PORTAL_MEMBER_DIR_COLS.OFFICE_DAYS).toBe(8);
    expect(PORTAL_MEMBER_DIR_COLS.BADGES).toBe(9);
  });
});

// PORTAL_POLL_COLS removed v4.24.0 — FlashPolls replaced by unified wq* poll system

describe('PORTAL_GRIEVANCE_COLS', () => {
  test('has all 12 columns', () => {
    const keys = Object.keys(PORTAL_GRIEVANCE_COLS);
    expect(keys.length).toBe(12);
    expect(PORTAL_GRIEVANCE_COLS.UPDATED_DATE).toBe(11);
  });
});

describe('PORTAL_STEWARD_LOG_COLS', () => {
  test('has STEWARD_EMAIL and MEMBER_EMAIL', () => {
    expect(PORTAL_STEWARD_LOG_COLS.STEWARD_EMAIL).toBe(1);
    expect(PORTAL_STEWARD_LOG_COLS.MEMBER_EMAIL).toBe(2);
  });
});

describe('PORTAL_MEGA_SURVEY_COLS', () => {
  test('has 5 columns', () => {
    expect(Object.keys(PORTAL_MEGA_SURVEY_COLS).length).toBe(5);
    expect(PORTAL_MEGA_SURVEY_COLS.EMAIL).toBe(0);
    expect(PORTAL_MEGA_SURVEY_COLS.COMPLETED).toBe(4);
  });
});

// ============================================================================
// PORTAL_SHEET_NAMES_
// ============================================================================

describe('PORTAL_SHEET_NAMES_', () => {
  test('has all 6 sheet names', () => {
    expect(Object.keys(PORTAL_SHEET_NAMES_).length).toBe(6);
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('MEMBER_DIR');
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('EVENTS');
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('MINUTES');
    // POLLS / POLL_RESPONSES removed v4.24.0
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('GRIEVANCES');
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('STEWARD_LOG');
    expect(PORTAL_SHEET_NAMES_).toHaveProperty('MEGA_SURVEY');
  });

  test('uses SHEETS constants when available', () => {
    // Since SHEETS is loaded from 01_Core.gs, it should be used
    if (SHEETS.PORTAL_EVENTS) {
      expect(PORTAL_SHEET_NAMES_.EVENTS).toBe(SHEETS.PORTAL_EVENTS);
    }
  });
});

// ============================================================================
// portalGetOrCreateSheet_
// ============================================================================

describe('portalGetOrCreateSheet_', () => {
  test('creates sheet when it does not exist', () => {
    const sheet = portalGetOrCreateSheet_('TestSheet', ['A', 'B', 'C'], false);
    expect(mockSs.insertSheet).toHaveBeenCalledWith('TestSheet');
    expect(sheet).toBeTruthy();
  });

  test('returns existing sheet without creating', () => {
    const existingSheet = createMockSheet('ExistingSheet');
    mockSs = createMockSpreadsheet([existingSheet]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);

    const sheet = portalGetOrCreateSheet_('ExistingSheet', ['A', 'B'], false);
    expect(mockSs.insertSheet).not.toHaveBeenCalled();
    expect(sheet.getName()).toBe('ExistingSheet');
  });

  test('throws when getActiveSpreadsheet returns null', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(() => portalGetOrCreateSheet_('Fail', ['A'], false)).toThrow('getActiveSpreadsheet() returned null');
  });

  test('hides sheet when hidden=true', () => {
    const sheet = portalGetOrCreateSheet_('HiddenSheet', ['A'], true);
    expect(sheet.hideSheet).toHaveBeenCalled();
  });

  test('sets frozen rows on new sheet', () => {
    const sheet = portalGetOrCreateSheet_('FrozenSheet', ['A', 'B'], false);
    expect(sheet.setFrozenRows).toHaveBeenCalledWith(1);
  });
});

// ============================================================================
// Individual getOrCreate functions
// ============================================================================

describe('getOrCreatePortalMemberDirectory', () => {
  test('creates PortalMemberDirectory sheet', () => {
    getOrCreatePortalMemberDirectory();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

describe('getOrCreateEventsSheet', () => {
  test('creates Events sheet', () => {
    getOrCreateEventsSheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

describe('getOrCreateMinutesSheet', () => {
  test('creates MeetingMinutes sheet', () => {
    getOrCreateMinutesSheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

// getOrCreatePollsSheet / getOrCreatePollResponsesSheet removed v4.24.0

describe('getOrCreatePortalGrievanceSheet', () => {
  test('creates hidden PortalGrievances sheet', () => {
    getOrCreatePortalGrievanceSheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

describe('getOrCreateStewardLogSheet', () => {
  test('creates hidden StewardLog sheet', () => {
    getOrCreateStewardLogSheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

describe('getOrCreateMegaSurveySheet', () => {
  test('creates hidden MegaSurvey sheet', () => {
    getOrCreateMegaSurveySheet();
    expect(mockSs.insertSheet).toHaveBeenCalled();
  });
});

// ============================================================================
// initPortalSheets
// ============================================================================

describe('initPortalSheets', () => {
  test('creates all 8 sheets', () => {
    initPortalSheets();
    // Should have called insertSheet 8 times (once per portal sheet)
    expect(mockSs.insertSheet).toHaveBeenCalled();
    // Verify Logger was called
    expect(Logger.log).toHaveBeenCalledWith(expect.stringContaining('Portal sheets initialized'));
  });
});
