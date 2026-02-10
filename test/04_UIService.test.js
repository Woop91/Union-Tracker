/**
 * Tests for 04_UIService.gs
 *
 * Covers: parseCSVLine_, mapImportColumns_, getCommonStyles,
 * getQuickCaptureNotes, saveQuickCaptureNotes, clearQuickCaptureNotes,
 * setBreakReminders, getDashboardStats, showVisualControlPanel,
 * startPomodoroTimer, and HTML generator smoke tests.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '04_UIService.gs'
]);

// ============================================================================
// Helper: set up PropertiesService.getUserProperties mock
// ============================================================================

function setupUserPropertiesMock(store) {
  const propsStore = store || {};
  const mockUserProps = {
    getProperty: jest.fn(key => propsStore[key] !== undefined ? propsStore[key] : null),
    setProperty: jest.fn((key, val) => { propsStore[key] = val; }),
    deleteProperty: jest.fn(key => { delete propsStore[key]; })
  };
  PropertiesService.getUserProperties = jest.fn(() => mockUserProps);
  return mockUserProps;
}

// ============================================================================
// parseCSVLine_
// ============================================================================

describe('parseCSVLine_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('parses simple comma-separated values', () => {
    const result = parseCSVLine_('a,b,c');
    expect(result).toEqual(['a', 'b', 'c']);
  });

  test('handles quoted values containing commas', () => {
    const result = parseCSVLine_('"a,b",c');
    expect(result).toEqual(['a,b', 'c']);
  });

  test('strips quotes from quoted values', () => {
    const result = parseCSVLine_('"hello",world');
    expect(result).toEqual(['hello', 'world']);
  });

  test('handles empty input', () => {
    const result = parseCSVLine_('');
    expect(result).toEqual(['']);
  });

  test('handles single value without comma', () => {
    const result = parseCSVLine_('onlyone');
    expect(result).toEqual(['onlyone']);
  });

  test('trims whitespace from values', () => {
    const result = parseCSVLine_('  a , b , c  ');
    expect(result).toEqual(['a', 'b', 'c']);
  });

  test('handles multiple quoted fields', () => {
    const result = parseCSVLine_('"first","second","third"');
    expect(result).toEqual(['first', 'second', 'third']);
  });

  test('handles mixed quoted and unquoted fields', () => {
    const result = parseCSVLine_('plain,"with,comma",another');
    expect(result).toEqual(['plain', 'with,comma', 'another']);
  });

  test('handles empty fields between commas', () => {
    const result = parseCSVLine_('a,,c');
    expect(result).toEqual(['a', '', 'c']);
  });
});

// ============================================================================
// mapImportColumns_
// ============================================================================

describe('mapImportColumns_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('maps "First Name" to firstName', () => {
    const result = mapImportColumns_(['First Name', 'Last Name']);
    expect(result.firstName).toBe(0);
    expect(result.lastName).toBe(1);
  });

  test('maps "Email" to email', () => {
    const result = mapImportColumns_(['Email']);
    expect(result.email).toBe(0);
  });

  test('maps "Email Address" to email', () => {
    const result = mapImportColumns_(['Email Address']);
    expect(result.email).toBe(0);
  });

  test('maps "Phone" to phone', () => {
    const result = mapImportColumns_(['Phone']);
    expect(result.phone).toBe(0);
  });

  test('maps "Phone Number" to phone', () => {
    const result = mapImportColumns_(['Phone Number']);
    expect(result.phone).toBe(0);
  });

  test('maps "Job Title" to jobTitle', () => {
    const result = mapImportColumns_(['Job Title']);
    expect(result.jobTitle).toBe(0);
  });

  test('maps "Title" and "Position" to jobTitle', () => {
    const resultTitle = mapImportColumns_(['Title']);
    expect(resultTitle.jobTitle).toBe(0);
    const resultPos = mapImportColumns_(['Position']);
    expect(resultPos.jobTitle).toBe(0);
  });

  test('maps "Unit", "Department", and "Dept" to unit', () => {
    expect(mapImportColumns_(['Unit']).unit).toBe(0);
    expect(mapImportColumns_(['Department']).unit).toBe(0);
    expect(mapImportColumns_(['Dept']).unit).toBe(0);
  });

  test('maps "Work Location", "Location", and "Worksite" to workLocation', () => {
    expect(mapImportColumns_(['Work Location']).workLocation).toBe(0);
    expect(mapImportColumns_(['Location']).workLocation).toBe(0);
    expect(mapImportColumns_(['Worksite']).workLocation).toBe(0);
  });

  test('maps "Supervisor" and "Manager" to supervisor', () => {
    expect(mapImportColumns_(['Supervisor']).supervisor).toBe(0);
    expect(mapImportColumns_(['Manager']).supervisor).toBe(0);
  });

  test('maps "Is Steward" and "Steward" to isSteward', () => {
    expect(mapImportColumns_(['Is Steward']).isSteward).toBe(0);
    expect(mapImportColumns_(['Steward']).isSteward).toBe(0);
  });

  test('maps "Dues Paying" and "Dues" to duesPaying', () => {
    expect(mapImportColumns_(['Dues Paying']).duesPaying).toBe(0);
    expect(mapImportColumns_(['Dues']).duesPaying).toBe(0);
  });

  test('maps a full set of common headers correctly', () => {
    const headers = ['First Name', 'Last Name', 'Email', 'Phone', 'Job Title', 'Unit', 'Work Location'];
    const result = mapImportColumns_(headers);
    expect(result.firstName).toBe(0);
    expect(result.lastName).toBe(1);
    expect(result.email).toBe(2);
    expect(result.phone).toBe(3);
    expect(result.jobTitle).toBe(4);
    expect(result.unit).toBe(5);
    expect(result.workLocation).toBe(6);
  });

  test('ignores unrecognized headers', () => {
    const result = mapImportColumns_(['Unknown Column', 'Random Header']);
    expect(Object.keys(result).length).toBe(0);
  });

  test('handles empty header array', () => {
    const result = mapImportColumns_([]);
    expect(result).toEqual({});
  });
});

// ============================================================================
// getCommonStyles
// ============================================================================

describe('getCommonStyles', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns a string containing CSS', () => {
    const result = getCommonStyles();
    expect(typeof result).toBe('string');
    expect(result).toContain('<style>');
  });

  test('includes btn class styling', () => {
    const result = getCommonStyles();
    expect(result).toContain('.btn');
  });
});

// ============================================================================
// Quick Capture Notes (PropertiesService-based)
// ============================================================================

describe('getQuickCaptureNotes', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns empty string when no notes are saved', () => {
    setupUserPropertiesMock({});
    const result = getQuickCaptureNotes();
    expect(result).toBe('');
  });

  test('returns saved notes from PropertiesService', () => {
    setupUserPropertiesMock({ quickCaptureNotes: 'My important note' });
    const result = getQuickCaptureNotes();
    expect(result).toBe('My important note');
  });
});

describe('saveQuickCaptureNotes', () => {
  beforeEach(() => jest.clearAllMocks());

  test('saves notes to PropertiesService and returns success', () => {
    const mockProps = setupUserPropertiesMock({});
    const result = saveQuickCaptureNotes('New notes here');
    expect(result.success).toBe(true);
    expect(mockProps.setProperty).toHaveBeenCalledWith('quickCaptureNotes', 'New notes here');
  });

  test('saves last-saved timestamp', () => {
    const mockProps = setupUserPropertiesMock({});
    saveQuickCaptureNotes('Test');
    expect(mockProps.setProperty).toHaveBeenCalledWith(
      'quickCaptureLastSaved',
      expect.any(String)
    );
  });

  test('saves empty string when notes is null or undefined', () => {
    const mockProps = setupUserPropertiesMock({});
    const result = saveQuickCaptureNotes(null);
    expect(result.success).toBe(true);
    expect(mockProps.setProperty).toHaveBeenCalledWith('quickCaptureNotes', '');
  });
});

describe('clearQuickCaptureNotes', () => {
  beforeEach(() => jest.clearAllMocks());

  test('deletes notes properties and returns success', () => {
    const mockProps = setupUserPropertiesMock({
      quickCaptureNotes: 'Some notes',
      quickCaptureLastSaved: '2025-01-01'
    });
    const result = clearQuickCaptureNotes();
    expect(result.success).toBe(true);
    expect(mockProps.deleteProperty).toHaveBeenCalledWith('quickCaptureNotes');
    expect(mockProps.deleteProperty).toHaveBeenCalledWith('quickCaptureLastSaved');
  });
});

// ============================================================================
// setBreakReminders
// ============================================================================

describe('setBreakReminders', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    setupUserPropertiesMock({});

    // Mock ScriptApp triggers
    const mockTrigger = {
      getHandlerFunction: jest.fn(() => 'someOtherFunction')
    };
    ScriptApp.getProjectTriggers.mockReturnValue([mockTrigger]);
    ScriptApp.newTrigger.mockReturnValue({
      timeBased: jest.fn(function() { return this; }),
      everyMinutes: jest.fn(function() { return this; }),
      create: jest.fn()
    });

    // Mock toast
    const mockSS = { toast: jest.fn() };
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);
  });

  test('creates a new trigger when minutes > 0', () => {
    setBreakReminders(30);
    expect(ScriptApp.newTrigger).toHaveBeenCalledWith('showBreakReminder');
  });

  test('does not create trigger when minutes is 0', () => {
    setBreakReminders(0);
    expect(ScriptApp.newTrigger).not.toHaveBeenCalled();
  });

  test('deletes existing showBreakReminder triggers', () => {
    const breakTrigger = {
      getHandlerFunction: jest.fn(() => 'showBreakReminder')
    };
    ScriptApp.getProjectTriggers.mockReturnValue([breakTrigger]);

    setBreakReminders(15);
    expect(ScriptApp.deleteTrigger).toHaveBeenCalledWith(breakTrigger);
  });
});

// ============================================================================
// getDashboardStats (04_UIService version)
// ============================================================================

describe('getDashboardStats', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns stats object with expected keys', () => {
    // Set up minimal mock sheets
    const grievanceData = [
      ['Grievance ID', 'Status', 'Current Step'],
      ['GRV-2025-0001', 'Open', 1]
    ];
    const memberData = [
      ['Member ID', 'First Name'],
      ['MS-101-H', 'John']
    ];

    const gSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, grievanceData);
    gSheet.getLastRow.mockReturnValue(2);
    gSheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => grievanceData),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => 2),
      getNumColumns: jest.fn(() => 3)
    });

    const mSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
    mSheet.getLastRow.mockReturnValue(2);
    mSheet.getDataRange.mockReturnValue({
      getValues: jest.fn(() => memberData),
      getRow: jest.fn(() => 1),
      getColumn: jest.fn(() => 1),
      getNumRows: jest.fn(() => 2),
      getNumColumns: jest.fn(() => 2)
    });

    const mockSS = createMockSpreadsheet([gSheet, mSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const rawResult = getDashboardStats();
    // getDashboardStats returns JSON.stringify(stats)
    const stats = JSON.parse(rawResult);
    expect(stats).toHaveProperty('totalGrievances');
    expect(stats).toHaveProperty('totalMembers');
    expect(stats).toHaveProperty('activeGrievances');
    expect(stats).toHaveProperty('outcomes');
    expect(stats).toHaveProperty('winRate');
  });
});

// ============================================================================
// showVisualControlPanel (smoke test)
// ============================================================================

describe('showVisualControlPanel', () => {
  beforeEach(() => jest.clearAllMocks());

  test('calls HtmlService.createHtmlOutput and showSidebar', () => {
    const mockSidebar = jest.fn();
    const mockUi = {
      showSidebar: mockSidebar,
      alert: jest.fn(),
      createMenu: jest.fn(() => ({
        addItem: jest.fn(function() { return this; }),
        addSubMenu: jest.fn(function() { return this; }),
        addSeparator: jest.fn(function() { return this; }),
        addToUi: jest.fn()
      })),
      ButtonSet: { OK: 'OK' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showVisualControlPanel();
    expect(HtmlService.createHtmlOutput).toHaveBeenCalled();
    expect(mockSidebar).toHaveBeenCalled();
  });
});

// ============================================================================
// HTML generator smoke tests
// ============================================================================

describe('HTML generator functions (smoke tests)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('getVisualControlPanelHtml returns HTML string', () => {
    const html = getVisualControlPanelHtml();
    expect(typeof html).toBe('string');
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('Visual Control Panel');
  });

  test('getCommonStyles returns a string with style tags', () => {
    const styles = getCommonStyles();
    expect(typeof styles).toBe('string');
    expect(styles).toContain('<style>');
    expect(styles).toContain('</style>');
  });

  test('getDashboardSidebarHtml returns HTML string', () => {
    if (typeof getDashboardSidebarHtml === 'function') {
      const html = getDashboardSidebarHtml();
      expect(typeof html).toBe('string');
      expect(html).toContain('<!DOCTYPE html>');
    }
  });

  test('getMultiSelectHtml returns HTML string', () => {
    if (typeof getMultiSelectHtml === 'function') {
      const items = [{ id: 'a', label: 'Item A', selected: false }];
      const html = getMultiSelectHtml(items, 'testCallback');
      expect(typeof html).toBe('string');
      expect(html).toContain('<!DOCTYPE html>');
    }
  });
});
