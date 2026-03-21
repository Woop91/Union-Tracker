/**
 * Tests for 11_CommandHub.gs
 *
 * Covers deprecated config removal, darkenColor_,
 * safetyValveScrub, isProductionMode,
 * getCommandCenterConfig, and UI smoke tests.
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
// Deprecated config removal verification
// ============================================================================

describe('deprecated config removal', () => {
  test('COMMAND_CENTER_CONFIG is removed (was deprecated)', () => {
    expect(typeof COMMAND_CENTER_CONFIG).toBe('undefined');
  });

  test('GEMINI_CONFIG is removed (was deprecated)', () => {
    expect(typeof GEMINI_CONFIG).toBe('undefined');
  });

  test('COMMAND_CONFIG is the canonical config', () => {
    expect(COMMAND_CONFIG).toBeDefined();
    expect(COMMAND_CONFIG.SYSTEM_NAME).toBeDefined();
    expect(COMMAND_CONFIG.UNIT_CODES).toBeDefined();
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
// getCommandCenterConfig
// ============================================================================

describe('getCommandCenterConfig', () => {
  test('returns an object', () => {
    const config = getCommandCenterConfig();
    expect(typeof config).toBe('object');
  });

  test('has SYSTEM_NAME matching COMMAND_CENTER_CONFIG', () => {
    const config = getCommandCenterConfig();
    expect(config.SYSTEM_NAME).toBe('Strategic Command Center');
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

// navigateToDashboard removed v4.33.0 — deprecated Dashboard sheet cleaned up
