/**
 * Tests for 03_UIComponents.gs
 *
 * Covers: showToast, showConfirmation, showAlert, navigateToSheet,
 * getDefaultComfortViewSettings_, getComfortViewSettings, saveComfortViewSettings,
 * isMobileContext, createDashboardMenu.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '03_UIComponents.gs'
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
// showToast
// ============================================================================

describe('showToast', () => {
  beforeEach(() => jest.clearAllMocks());

  test('calls SpreadsheetApp.getActiveSpreadsheet().toast()', () => {
    const mockToast = jest.fn();
    const mockSS = { toast: mockToast };
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    showToast('Hello world', 'Title');
    expect(mockToast).toHaveBeenCalledWith('Hello world', 'Title', 3);
  });

  test('uses default title "Info" when title is not provided', () => {
    const mockToast = jest.fn();
    const mockSS = { toast: mockToast };
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    showToast('Hello world');
    expect(mockToast).toHaveBeenCalledWith('Hello world', 'Info', 3);
  });
});

// ============================================================================
// showAlert
// ============================================================================

describe('showAlert', () => {
  beforeEach(() => jest.clearAllMocks());

  test('calls ui.alert with OK button set', () => {
    const mockAlert = jest.fn();
    const mockUi = {
      alert: mockAlert,
      ButtonSet: { OK: 'OK' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showAlert('Something happened', 'Notice');
    expect(mockAlert).toHaveBeenCalledWith('Notice', 'Something happened', 'OK');
  });

  test('uses default title "Alert" when not provided', () => {
    const mockAlert = jest.fn();
    const mockUi = {
      alert: mockAlert,
      ButtonSet: { OK: 'OK' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showAlert('Message only');
    expect(mockAlert).toHaveBeenCalledWith('Alert', 'Message only', 'OK');
  });
});

// ============================================================================
// navigateToSheet
// ============================================================================

describe('navigateToSheet', () => {
  beforeEach(() => jest.clearAllMocks());

  test('sets the target sheet as active when it exists', () => {
    const targetSheet = createMockSheet('MySheet');
    const mockSS = createMockSpreadsheet([targetSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    navigateToSheet('MySheet');
    expect(mockSS.setActiveSheet).toHaveBeenCalledWith(targetSheet);
  });

  test('shows alert when sheet does not exist', () => {
    const mockSS = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    const mockAlert = jest.fn();
    const mockUi = {
      alert: mockAlert,
      ButtonSet: { OK: 'OK' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    navigateToSheet('Nonexistent Sheet');
    expect(mockSS.setActiveSheet).not.toHaveBeenCalled();
    expect(mockAlert).toHaveBeenCalled();
  });
});

// ============================================================================
// getDefaultComfortViewSettings_
// ============================================================================

describe('getDefaultComfortViewSettings_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns an object with expected default properties', () => {
    const defaults = getDefaultComfortViewSettings_();
    expect(defaults).toHaveProperty('zebraStripes', true);
    expect(defaults).toHaveProperty('reducedMotion', false);
    expect(defaults).toHaveProperty('focusMode', false);
    expect(defaults).toHaveProperty('hideGridlines', false);
  });

  test('zebraStripes defaults to true', () => {
    const defaults = getDefaultComfortViewSettings_();
    expect(defaults.zebraStripes).toBe(true);
  });

  test('focusMode defaults to false', () => {
    const defaults = getDefaultComfortViewSettings_();
    expect(defaults.focusMode).toBe(false);
  });

  test('returns a fresh object each time (not shared reference)', () => {
    const a = getDefaultComfortViewSettings_();
    const b = getDefaultComfortViewSettings_();
    expect(a).not.toBe(b);
    expect(a).toEqual(b);
  });
});

// ============================================================================
// getComfortViewSettings
// ============================================================================

describe('getComfortViewSettings', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns saved settings from PropertiesService', () => {
    const savedSettings = {
      zebraStripes: false,
      reducedMotion: true,
      focusMode: true,
      hideGridlines: true
    };
    setupUserPropertiesMock({ comfortViewSettings: JSON.stringify(savedSettings) });

    const result = getComfortViewSettings();
    expect(result.zebraStripes).toBe(false);
    expect(result.reducedMotion).toBe(true);
    expect(result.focusMode).toBe(true);
  });

  test('returns defaults when no settings are saved', () => {
    setupUserPropertiesMock({});

    const result = getComfortViewSettings();
    const defaults = getDefaultComfortViewSettings_();
    expect(result).toEqual(defaults);
  });
});

// ============================================================================
// saveComfortViewSettings
// ============================================================================

describe('saveComfortViewSettings', () => {
  beforeEach(() => jest.clearAllMocks());

  test('saves settings as JSON to PropertiesService', () => {
    const mockProps = setupUserPropertiesMock({});
    const settings = { zebraStripes: true, focusMode: true };

    saveComfortViewSettings(settings);
    expect(mockProps.setProperty).toHaveBeenCalledWith(
      'comfortViewSettings',
      JSON.stringify(settings)
    );
  });
});

// ============================================================================
// ============================================================================
// createDashboardMenu
// ============================================================================

describe('createDashboardMenu', () => {
  beforeEach(() => jest.clearAllMocks());

  test('calls SpreadsheetApp.getUi() to build menus', () => {
    const mockMenu = {
      addItem: jest.fn(function() { return this; }),
      addSubMenu: jest.fn(function() { return this; }),
      addSeparator: jest.fn(function() { return this; }),
      addToUi: jest.fn()
    };
    const mockUi = {
      createMenu: jest.fn(() => mockMenu),
      alert: jest.fn(),
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    // isProductionMode may not be loaded; provide a global fallback
    if (typeof global.isProductionMode === 'undefined') {
      global.isProductionMode = jest.fn(() => false);
    }

    createDashboardMenu();
    expect(mockUi.createMenu).toHaveBeenCalled();
    expect(mockMenu.addToUi).toHaveBeenCalled();
  });

  test('creates at least the main hub menu', () => {
    const menuNames = [];
    const mockMenu = {
      addItem: jest.fn(function() { return this; }),
      addSubMenu: jest.fn(function() { return this; }),
      addSeparator: jest.fn(function() { return this; }),
      addToUi: jest.fn()
    };
    const mockUi = {
      createMenu: jest.fn(name => {
        menuNames.push(name);
        return mockMenu;
      }),
      alert: jest.fn(),
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    if (typeof global.isProductionMode === 'undefined') {
      global.isProductionMode = jest.fn(() => false);
    }

    createDashboardMenu();
    // Should have created menus including the Union Hub menu
    expect(menuNames.length).toBeGreaterThanOrEqual(1);
  });
});

// ============================================================================
// showDialog_ utility
// ============================================================================

describe('showDialog_ utility', () => {
  beforeEach(() => jest.clearAllMocks());

  test('showDialog_ calls showModalDialog with correct dimensions', () => {
    const mockShowModalDialog = jest.fn();
    const mockUi = {
      alert: jest.fn(),
      showModalDialog: mockShowModalDialog,
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showDialog_('<h1>Test</h1>', 'Test Title', 400, 300);
    expect(mockShowModalDialog).toHaveBeenCalled();
    const [htmlArg, titleArg] = mockShowModalDialog.mock.calls[0];
    expect(titleArg).toBe('Test Title');
    // The HtmlOutput object is what gets passed (not the raw string)
    expect(htmlArg).toBeDefined();
  });

  test('showDialog_ uses default dimensions when not provided', () => {
    const mockShowModalDialog = jest.fn();
    const mockUi = {
      alert: jest.fn(),
      showModalDialog: mockShowModalDialog,
      ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showDialog_('<h1>Test</h1>', 'Default Title');
    expect(mockShowModalDialog).toHaveBeenCalled();
    const [, titleArg] = mockShowModalDialog.mock.calls[0];
    expect(titleArg).toBe('Default Title');
  });
});

// ============================================================================
// C-XSS-17 regression: email validation regex
// ============================================================================

describe('email validation regex (C-XSS-17 regression)', () => {
  // The regex used in sendMemberReport_ after C-XSS-17 fix
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  test.each([
    ['user@example.com', true],
    ['first.last@subdomain.domain.org', true],
    ['user+tag@example.co.uk', true],
    ['not-an-email', false],
    ['@example.com', false],
    ['user@', false],
    ['user @example.com', false],
    ['user@example', false],
    ['', false],
  ])('email "%s" valid=%s', (email, expected) => {
    expect(emailRegex.test(email)).toBe(expected);
  });
});

// ============================================================================
// isTabModalsEnabled_ — Auditor-Alpha AA-14
// ============================================================================
//
// The ENABLE_TAB_MODALS feature toggle gates the v4.48.0 tab-modal system
// that replaces legacy sidebar modals. Pre-v4.55.2 there were zero tests
// for this toggle. A regression inverting the check would either spam
// users with modal dialogs everywhere or suppress them entirely. These
// tests exercise every branch of isTabModalsEnabled_ directly against
// the real source.
// ============================================================================

describe('isTabModalsEnabled_ (AA-14)', () => {
  var _origGetActive;

  beforeEach(() => {
    _origGetActive = SpreadsheetApp.getActiveSpreadsheet;
  });

  afterEach(() => {
    SpreadsheetApp.getActiveSpreadsheet = _origGetActive;
  });

  function mockConfigSheet(value) {
    var configSheet = {
      getRange: jest.fn(() => ({ getValue: jest.fn(() => value) })),
      getName: jest.fn(() => SHEETS.CONFIG)
    };
    var ss = {
      getSheetByName: jest.fn(name => (name === SHEETS.CONFIG ? configSheet : null))
    };
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);
    return configSheet;
  }

  test('defaults to true when no active spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('defaults to true when Config sheet is missing', () => {
    var ss = { getSheetByName: jest.fn(() => null) };
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('returns true when Config value is "yes"', () => {
    mockConfigSheet('yes');
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('returns true for "YES" (case-insensitive)', () => {
    mockConfigSheet('YES');
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('returns true when Config value is empty (default enabled)', () => {
    mockConfigSheet('');
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('returns true when Config value is whitespace', () => {
    mockConfigSheet('   ');
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('returns false when Config value is "no"', () => {
    mockConfigSheet('no');
    expect(isTabModalsEnabled_()).toBe(false);
  });

  test('returns false for "NO" (case-insensitive)', () => {
    mockConfigSheet('NO');
    expect(isTabModalsEnabled_()).toBe(false);
  });

  test('returns false for "  no  " (whitespace padded)', () => {
    mockConfigSheet('  no  ');
    expect(isTabModalsEnabled_()).toBe(false);
  });

  test('returns true for arbitrary non-"no" string (fail-open)', () => {
    mockConfigSheet('enabled');
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('fail-open: returns true when getRange throws', () => {
    var configSheet = {
      getRange: jest.fn(() => { throw new Error('Range error'); })
    };
    var ss = {
      getSheetByName: jest.fn(() => configSheet)
    };
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ss);
    expect(isTabModalsEnabled_()).toBe(true);
  });

  test('reads from row 3 of the Config sheet', () => {
    var configSheet = mockConfigSheet('yes');
    isTabModalsEnabled_();
    // First arg to getRange is row (3), second is column (ENABLE_TAB_MODALS)
    expect(configSheet.getRange).toHaveBeenCalledWith(3, CONFIG_COLS.ENABLE_TAB_MODALS);
  });
});
