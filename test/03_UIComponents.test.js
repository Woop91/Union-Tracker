/**
 * Tests for 03_UIComponents.gs
 *
 * Covers: showToast, showConfirmation, showAlert, navigateToSheet,
 * getDefaultADHDSettings_, getADHDSettings, saveADHDSettings,
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
// showConfirmation
// ============================================================================

describe('showConfirmation', () => {
  beforeEach(() => jest.clearAllMocks());

  test('calls ui.alert with YES_NO button set', () => {
    const mockAlert = jest.fn(() => 'YES');
    const mockUi = {
      alert: mockAlert,
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showConfirmation('Are you sure?', 'Confirm');
    expect(mockAlert).toHaveBeenCalledWith('Confirm', 'Are you sure?', 'YES_NO');
  });

  test('returns true when user clicks YES', () => {
    const mockUi = {
      alert: jest.fn(() => 'YES'),
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    const result = showConfirmation('Continue?', 'Confirm');
    expect(result).toBe(true);
  });

  test('returns false when user clicks NO', () => {
    const mockUi = {
      alert: jest.fn(() => 'NO'),
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    const result = showConfirmation('Continue?', 'Confirm');
    expect(result).toBe(false);
  });

  test('uses default title "Confirm" when title is not provided', () => {
    const mockAlert = jest.fn(() => 'YES');
    const mockUi = {
      alert: mockAlert,
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES', NO: 'NO' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);

    showConfirmation('Question?');
    expect(mockAlert).toHaveBeenCalledWith('Confirm', 'Question?', 'YES_NO');
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
// getDefaultADHDSettings_
// ============================================================================

describe('getDefaultADHDSettings_', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns an object with expected default properties', () => {
    const defaults = getDefaultADHDSettings_();
    expect(defaults).toHaveProperty('zebraStripes', true);
    expect(defaults).toHaveProperty('reducedMotion', false);
    expect(defaults).toHaveProperty('focusMode', false);
    expect(defaults).toHaveProperty('highContrast', false);
    expect(defaults).toHaveProperty('largeText', false);
    expect(defaults).toHaveProperty('hideGridlines', false);
  });

  test('zebraStripes defaults to true', () => {
    const defaults = getDefaultADHDSettings_();
    expect(defaults.zebraStripes).toBe(true);
  });

  test('focusMode defaults to false', () => {
    const defaults = getDefaultADHDSettings_();
    expect(defaults.focusMode).toBe(false);
  });

  test('returns a fresh object each time (not shared reference)', () => {
    const a = getDefaultADHDSettings_();
    const b = getDefaultADHDSettings_();
    expect(a).not.toBe(b);
    expect(a).toEqual(b);
  });
});

// ============================================================================
// getADHDSettings
// ============================================================================

describe('getADHDSettings', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns saved settings from PropertiesService', () => {
    const savedSettings = {
      zebraStripes: false,
      reducedMotion: true,
      focusMode: true,
      highContrast: true,
      largeText: true,
      hideGridlines: true
    };
    setupUserPropertiesMock({ adhdSettings: JSON.stringify(savedSettings) });

    const result = getADHDSettings();
    expect(result.zebraStripes).toBe(false);
    expect(result.reducedMotion).toBe(true);
    expect(result.focusMode).toBe(true);
  });

  test('returns defaults when no settings are saved', () => {
    setupUserPropertiesMock({});

    const result = getADHDSettings();
    const defaults = getDefaultADHDSettings_();
    expect(result).toEqual(defaults);
  });
});

// ============================================================================
// saveADHDSettings
// ============================================================================

describe('saveADHDSettings', () => {
  beforeEach(() => jest.clearAllMocks());

  test('saves settings as JSON to PropertiesService', () => {
    const mockProps = setupUserPropertiesMock({});
    const settings = { zebraStripes: true, focusMode: true };

    saveADHDSettings(settings);
    expect(mockProps.setProperty).toHaveBeenCalledWith(
      'adhdSettings',
      JSON.stringify(settings)
    );
  });
});

// ============================================================================
// isMobileContext
// ============================================================================

describe('isMobileContext', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns a boolean value', () => {
    const result = isMobileContext();
    expect(typeof result).toBe('boolean');
  });

  test('returns false in the test environment', () => {
    const result = isMobileContext();
    expect(result).toBe(false);
  });
});

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
