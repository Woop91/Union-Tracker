/**
 * Tests for 11_CommandHub.gs
 *
 * Covers COMMAND_CENTER_CONFIG, GEMINI_CONFIG, darkenColor_,
 * safetyValveScrub, scrubObjectPII, isProductionMode,
 * getNextSequence, getCommandCenterConfig, and UI smoke tests.
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
  '11_CommandHub.gs'
]);

// ============================================================================
// COMMAND_CENTER_CONFIG constant
// ============================================================================

describe('COMMAND_CENTER_CONFIG', () => {
  test('is defined', () => {
    expect(COMMAND_CENTER_CONFIG).toBeDefined();
  });

  test('has SYSTEM_NAME', () => {
    expect(COMMAND_CENTER_CONFIG.SYSTEM_NAME).toBe('509 Strategic Command Center');
  });

  test('has LOG_SHEET_NAME', () => {
    expect(typeof COMMAND_CENTER_CONFIG.LOG_SHEET_NAME).toBe('string');
    expect(COMMAND_CENTER_CONFIG.LOG_SHEET_NAME.length).toBeGreaterThan(0);
  });

  test('has DIR_SHEET_NAME', () => {
    expect(typeof COMMAND_CENTER_CONFIG.DIR_SHEET_NAME).toBe('string');
  });

  test('has THEME with expected keys', () => {
    expect(COMMAND_CENTER_CONFIG.THEME).toBeDefined();
    expect(COMMAND_CENTER_CONFIG.THEME.HEADER_BG).toBeDefined();
    expect(COMMAND_CENTER_CONFIG.THEME.HEADER_TEXT).toBeDefined();
    expect(COMMAND_CENTER_CONFIG.THEME.ALT_ROW).toBeDefined();
    expect(COMMAND_CENTER_CONFIG.THEME.FONT).toBeDefined();
  });
});

// ============================================================================
// GEMINI_CONFIG constant
// ============================================================================

describe('GEMINI_CONFIG', () => {
  test('is defined', () => {
    expect(GEMINI_CONFIG).toBeDefined();
  });

  test('has SYSTEM_NAME matching COMMAND_CENTER_CONFIG', () => {
    expect(GEMINI_CONFIG.SYSTEM_NAME).toBe(COMMAND_CENTER_CONFIG.SYSTEM_NAME);
  });

  test('has legacy emoji-prefixed sheet names', () => {
    expect(GEMINI_CONFIG.LOG_SHEET_NAME).toContain('Grievance Log');
    expect(GEMINI_CONFIG.DIR_SHEET_NAME).toContain('Member Directory');
  });

  test('has UNIT_CODES object', () => {
    expect(typeof GEMINI_CONFIG.UNIT_CODES).toBe('object');
  });

  test('has THEME object', () => {
    expect(GEMINI_CONFIG.THEME).toBeDefined();
    expect(GEMINI_CONFIG.THEME.HEADER_BG).toBe('#1e293b');
  });
});

// ============================================================================
// darkenColor_
// ============================================================================

describe('darkenColor_', () => {
  test('darkens a light blue color', () => {
    const result = darkenColor_('#e3f2fd', 20);
    expect(result).toMatch(/^#[0-9a-f]{6}$/);
    // The result should be darker than the input
    expect(result).not.toBe('#e3f2fd');
  });

  test('returns a valid hex color string', () => {
    const result = darkenColor_('#ffffff', 10);
    expect(result).toMatch(/^#[0-9a-f]{6}$/);
  });

  test('clamps to black when percent is 100', () => {
    // 2.55 * 100 = 255, which should make all channels 0 or near 0
    const result = darkenColor_('#ffffff', 100);
    expect(result).toBe('#000000');
  });

  test('returns same color when percent is 0', () => {
    const result = darkenColor_('#aabbcc', 0);
    expect(result).toBe('#aabbcc');
  });

  test('clamps channels to minimum 0 (no negative values)', () => {
    // Very dark color darkened further should not produce negative channels
    const result = darkenColor_('#050505', 50);
    expect(result).toMatch(/^#[0-9a-f]{6}$/);
    // Channels should be clamped to 0: 5 - 128 = -123 -> 0
    expect(result).toBe('#000000');
  });

  test('correctly computes a known value', () => {
    // #e3f2fd with 20% -> amt = 51
    // R = 0xe3(227) - 51 = 176 = 0xb0
    // G = 0xf2(242) - 51 = 191 = 0xbf
    // B = 0xfd(253) - 51 = 202 = 0xca
    const result = darkenColor_('#e3f2fd', 20);
    expect(result).toBe('#b0bfca');
  });

  test('handles pure red', () => {
    // #ff0000 with 50% -> amt = Math.round(2.55 * 50) = 127 (float: 127.499...)
    // R = 255 - 127 = 128 = 0x80
    // G = 0 - 127 = -127 -> 0
    // B = 0 - 127 = -127 -> 0
    const result = darkenColor_('#ff0000', 50);
    expect(result).toBe('#800000');
  });
});

// ============================================================================
// safetyValveScrub
// ============================================================================

describe('safetyValveScrub', () => {
  test('redacts phone number in parenthesized format', () => {
    expect(safetyValveScrub('Call (555) 123-4567')).toBe('Call [REDACTED CONTACT]');
  });

  test('redacts phone number in dashed format', () => {
    expect(safetyValveScrub('Phone: 555-123-4567')).toBe('Phone: [REDACTED CONTACT]');
  });

  test('redacts phone number with +1 prefix', () => {
    const result = safetyValveScrub('Call +1 555-123-4567');
    expect(result).toContain('[REDACTED CONTACT]');
    expect(result).not.toContain('555');
  });

  test('redacts SSN pattern', () => {
    expect(safetyValveScrub('SSN: 123-45-6789')).toBe('SSN: [REDACTED ID]');
  });

  test('redacts multiple phone numbers', () => {
    const result = safetyValveScrub('Home: (555) 111-2222 Work: (555) 333-4444');
    expect(result).not.toContain('555');
    expect((result.match(/\[REDACTED CONTACT\]/g) || []).length).toBe(2);
  });

  test('returns non-string input unchanged (number)', () => {
    expect(safetyValveScrub(42)).toBe(42);
  });

  test('returns non-string input unchanged (null)', () => {
    expect(safetyValveScrub(null)).toBeNull();
  });

  test('returns non-string input unchanged (undefined)', () => {
    expect(safetyValveScrub(undefined)).toBeUndefined();
  });

  test('returns non-string input unchanged (boolean)', () => {
    expect(safetyValveScrub(true)).toBe(true);
  });

  test('leaves non-PII strings unchanged', () => {
    expect(safetyValveScrub('Hello World')).toBe('Hello World');
  });

  test('handles empty string', () => {
    expect(safetyValveScrub('')).toBe('');
  });
});

// ============================================================================
// scrubObjectPII
// ============================================================================

describe('scrubObjectPII', () => {
  test('scrubs phone numbers in object values', () => {
    const result = scrubObjectPII({ name: 'John', phone: '(555) 123-4567' });
    expect(result.name).toBe('John');
    expect(result.phone).toBe('[REDACTED CONTACT]');
  });

  test('scrubs SSN in object values', () => {
    const result = scrubObjectPII({ ssn: '123-45-6789' });
    expect(result.ssn).toBe('[REDACTED ID]');
  });

  test('leaves non-string values unchanged', () => {
    const result = scrubObjectPII({ count: 42, active: true });
    expect(result.count).toBe(42);
    expect(result.active).toBe(true);
  });

  test('returns non-object input as-is (null)', () => {
    expect(scrubObjectPII(null)).toBeNull();
  });

  test('returns non-object input as-is (string)', () => {
    expect(scrubObjectPII('hello')).toBe('hello');
  });

  test('returns non-object input as-is (number)', () => {
    expect(scrubObjectPII(42)).toBe(42);
  });

  test('handles empty object', () => {
    const result = scrubObjectPII({});
    expect(Object.keys(result).length).toBe(0);
  });
});

// ============================================================================
// isProductionMode
// ============================================================================

describe('isProductionMode', () => {
  test('returns false when PRODUCTION_MODE is not set', () => {
    expect(isProductionMode()).toBe(false);
  });

  test('returns true when PRODUCTION_MODE is "true"', () => {
    PropertiesService.getScriptProperties().setProperty('PRODUCTION_MODE', 'true');
    expect(isProductionMode()).toBe(true);
    // Cleanup
    PropertiesService.getScriptProperties().deleteProperty('PRODUCTION_MODE');
  });

  test('returns false when PRODUCTION_MODE is "false"', () => {
    PropertiesService.getScriptProperties().setProperty('PRODUCTION_MODE', 'false');
    expect(isProductionMode()).toBe(false);
    PropertiesService.getScriptProperties().deleteProperty('PRODUCTION_MODE');
  });
});

// ============================================================================
// getNextSequence
// ============================================================================

describe('getNextSequence', () => {
  test('returns 4-digit padded string starting at 0001', () => {
    const result = getNextSequence('TEST_PREFIX_UNIQUE');
    expect(result).toBe('0001');
  });

  test('increments on subsequent calls', () => {
    const prefix = 'SEQ_INC_TEST';
    const first = getNextSequence(prefix);
    const second = getNextSequence(prefix);
    expect(parseInt(second, 10)).toBe(parseInt(first, 10) + 1);
  });

  test('pads to at least 4 digits', () => {
    const result = getNextSequence('PAD_TEST_UNIQUE');
    expect(result.length).toBeGreaterThanOrEqual(4);
    expect(result).toMatch(/^\d{4,}$/);
  });
});

// ============================================================================
// getCommandCenterConfig
// ============================================================================

describe('getCommandCenterConfig', () => {
  test('returns an object', () => {
    const config = getCommandCenterConfig();
    expect(typeof config).toBe('object');
  });

  test('has SYSTEM_NAME matching COMMAND_CENTER_CONFIG', () => {
    const config = getCommandCenterConfig();
    expect(config.SYSTEM_NAME).toBe('509 Strategic Command Center');
  });

  test('has LOG_SHEET_NAME', () => {
    const config = getCommandCenterConfig();
    expect(config.LOG_SHEET_NAME).toBeDefined();
  });

  test('has THEME object', () => {
    const config = getCommandCenterConfig();
    expect(config.THEME).toBeDefined();
  });
});

// ============================================================================
// UI function smoke tests
// ============================================================================

describe('showDiagnosticReport (smoke test)', () => {
  test('does not throw when called', () => {
    const dashboardSheet = createMockSheet(SHEETS.DASHBOARD || 'Dashboard', [['Header']]);
    const memberSheet = createMockSheet(SHEETS.MEMBER_DIR, [['Header']]);
    const grievanceSheet = createMockSheet(SHEETS.GRIEVANCE_LOG, [['Header']]);
    const configSheet = createMockSheet(SHEETS.CONFIG, [['Header']]);

    const ss = createMockSpreadsheet([dashboardSheet, memberSheet, grievanceSheet, configSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(() => showDiagnosticReport()).not.toThrow();
  });
});

describe('navigateToDashboard (smoke test)', () => {
  test('sets active sheet when dashboard exists', () => {
    const dashboardSheet = createMockSheet(SHEETS.DASHBOARD || 'Dashboard');
    const ss = createMockSpreadsheet([dashboardSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(() => navigateToDashboard()).not.toThrow();
  });

  test('does not throw when dashboard sheet is missing', () => {
    const ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(() => navigateToDashboard()).not.toThrow();
  });
});
