/**
 * Tests for 11_CommandHub.gs
 *
 * Covers deprecated config removal, isProductionMode,
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
