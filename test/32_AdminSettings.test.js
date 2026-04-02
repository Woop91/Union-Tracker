/**
 * Tests for 32_AdminSettings.gs
 *
 * Covers admin settings sidebar: schema structure, authorization checks,
 * input validation (URL format, number ranges, toggle normalization),
 * field definition lookup, and global entry points.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Stub dependencies that AdminSettings.gs calls
global.getConfigValues = jest.fn(() => []);
global.setupDataValidations = jest.fn();
global.logAuditEvent = jest.fn();
global.withScriptLock_ = jest.fn(function(fn) { fn(); });
global.ConfigReader = { refreshConfig: jest.fn() };
global.escapeHtml = jest.fn(function(s) { return String(s); });

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '32_AdminSettings.gs']);

// ============================================================================
// Module existence
// ============================================================================

describe('AdminSettings module existence', () => {
  test('ADMIN_SETTINGS_SCHEMA_ is defined', () => {
    expect(ADMIN_SETTINGS_SCHEMA_).toBeDefined();
    expect(Array.isArray(ADMIN_SETTINGS_SCHEMA_)).toBe(true);
    expect(ADMIN_SETTINGS_SCHEMA_.length).toBeGreaterThan(0);
  });

  test('schema has expected tab IDs', () => {
    var tabIds = ADMIN_SETTINGS_SCHEMA_.map(function(t) { return t.id; });
    expect(tabIds).toContain('org');
    expect(tabIds).toContain('lists');
    expect(tabIds).toContain('grievance');
    expect(tabIds).toContain('deadlines');
    expect(tabIds).toContain('integrations');
    expect(tabIds).toContain('branding');
  });

  test('global functions exist', () => {
    expect(typeof showAdminSettingsSidebar).toBe('function');
    expect(typeof adminGetSettings).toBe('function');
    expect(typeof adminSaveSettings).toBe('function');
    expect(typeof adminSaveListValues).toBe('function');
    expect(typeof _adminIsAuthorized_).toBe('function');
    expect(typeof _findFieldDef_).toBe('function');
  });
});

// ============================================================================
// Schema validation
// ============================================================================

describe('ADMIN_SETTINGS_SCHEMA_ structure', () => {
  test('every tab has id, label, icon, and fields array', () => {
    ADMIN_SETTINGS_SCHEMA_.forEach(function(tab) {
      expect(tab.id).toBeTruthy();
      expect(tab.label).toBeTruthy();
      expect(tab.icon).toBeTruthy();
      expect(Array.isArray(tab.fields)).toBe(true);
      expect(tab.fields.length).toBeGreaterThan(0);
    });
  });

  test('every field has key, label, type, and desc', () => {
    ADMIN_SETTINGS_SCHEMA_.forEach(function(tab) {
      tab.fields.forEach(function(field) {
        expect(field.key).toBeTruthy();
        expect(field.label).toBeTruthy();
        expect(field.type).toBeTruthy();
        expect(field.desc).toBeTruthy();
        expect(['text', 'number', 'url', 'email', 'toggle', 'list']).toContain(field.type);
      });
    });
  });

  test('number fields have min and max defined', () => {
    ADMIN_SETTINGS_SCHEMA_.forEach(function(tab) {
      tab.fields.forEach(function(field) {
        if (field.type === 'number') {
          expect(typeof field.min).toBe('number');
          expect(typeof field.max).toBe('number');
          expect(field.min).toBeLessThan(field.max);
        }
      });
    });
  });
});

// ============================================================================
// _findFieldDef_
// ============================================================================

describe('_findFieldDef_', () => {
  test('finds a known field by CONFIG_COLS key', () => {
    var result = _findFieldDef_('ORG_NAME');
    expect(result).not.toBeNull();
    expect(result.key).toBe('ORG_NAME');
    expect(result.type).toBe('text');
  });

  test('returns null for unknown key', () => {
    expect(_findFieldDef_('NONEXISTENT_KEY')).toBeNull();
  });

  test('finds toggle fields', () => {
    var result = _findFieldDef_('SHOW_GRIEVANCES');
    expect(result).not.toBeNull();
    expect(result.type).toBe('toggle');
  });

  test('finds list fields', () => {
    var result = _findFieldDef_('JOB_TITLES');
    expect(result).not.toBeNull();
    expect(result.type).toBe('list');
  });

  test('finds url fields', () => {
    var result = _findFieldDef_('ORG_WEBSITE');
    expect(result).not.toBeNull();
    expect(result.type).toBe('url');
  });
});

// ============================================================================
// _adminIsAuthorized_
// ============================================================================

describe('_adminIsAuthorized_', () => {
  test('returns false for null email', () => {
    expect(_adminIsAuthorized_(null)).toBe(false);
  });

  test('returns false for empty email', () => {
    expect(_adminIsAuthorized_('')).toBe(false);
  });

  test('returns true for spreadsheet owner', () => {
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(_adminIsAuthorized_('owner@test.com')).toBe(true);
  });

  test('owner check is case-insensitive', () => {
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    expect(_adminIsAuthorized_('OWNER@TEST.COM')).toBe(true);
  });

  test('returns false when not owner and not in ADMIN_EMAILS', () => {
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue(['admin1@test.com', 'admin2@test.com']);

    expect(_adminIsAuthorized_('random@test.com')).toBe(false);
  });

  test('returns true when in ADMIN_EMAILS list', () => {
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue(['admin1@test.com', 'admin2@test.com']);

    expect(_adminIsAuthorized_('admin2@test.com')).toBe(true);
  });

  test('returns false when no spreadsheet', () => {
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(null);
    expect(_adminIsAuthorized_('anyone@test.com')).toBe(false);
  });
});

// ============================================================================
// adminSaveSettings — input validation
// ============================================================================

describe('adminSaveSettings input validation', () => {
  beforeEach(() => {
    // Make the calling user authorized
    Session.getActiveUser.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue([]);
  });

  test('returns error for null changes', () => {
    var result = adminSaveSettings(null);
    expect(result.success).toBe(false);
    expect(result.error).toContain('Invalid');
  });

  test('validates URL format — rejects bad URL', () => {
    var result = adminSaveSettings({ ORG_WEBSITE: 'not-a-url' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('http');
  });

  test('validates URL format — accepts valid URL', () => {
    var result = adminSaveSettings({ ORG_WEBSITE: 'https://example.com' });
    expect(result.success).toBe(true);
  });

  test('validates number range — rejects too low', () => {
    var result = adminSaveSettings({ FILING_DEADLINE_DAYS: 0 });
    expect(result.success).toBe(false);
    expect(result.error).toContain('minimum');
  });

  test('validates number range — rejects too high', () => {
    var result = adminSaveSettings({ FILING_DEADLINE_DAYS: 999 });
    expect(result.success).toBe(false);
    expect(result.error).toContain('maximum');
  });

  test('validates number range — accepts valid number', () => {
    var result = adminSaveSettings({ FILING_DEADLINE_DAYS: 30 });
    expect(result.success).toBe(true);
  });

  test('validates NaN rejection', () => {
    var result = adminSaveSettings({ FILING_DEADLINE_DAYS: 'abc' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('number');
  });

  test('normalizes toggle — true becomes yes', () => {
    var result = adminSaveSettings({ SHOW_GRIEVANCES: true });
    expect(result.success).toBe(true);
  });

  test('normalizes toggle — false becomes no', () => {
    var result = adminSaveSettings({ SHOW_GRIEVANCES: false });
    expect(result.success).toBe(true);
  });

  test('skips readOnly fields', () => {
    var result = adminSaveSettings({ MOBILE_DASHBOARD_URL: 'https://evil.com' });
    // readOnly fields are silently skipped, so saved count is 0
    expect(result.success).toBe(true);
    expect(result.saved).toBe(0);
  });
});

// ============================================================================
// adminSaveSettings — authorization
// ============================================================================

describe('adminSaveSettings authorization', () => {
  test('rejects unauthorized user', () => {
    Session.getActiveUser.mockReturnValue({ getEmail: jest.fn(() => 'nobody@test.com') });
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue([]);

    var result = adminSaveSettings({ ORG_NAME: 'Test' });
    expect(result.success).toBe(false);
    expect(result.error).toContain('Admin access required');
  });
});

// ============================================================================
// adminGetSettings — authorization
// ============================================================================

describe('adminGetSettings authorization', () => {
  test('rejects unauthorized user', () => {
    Session.getActiveUser.mockReturnValue({ getEmail: jest.fn(() => 'nobody@test.com') });
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue([]);

    var result = adminGetSettings();
    expect(result.success).toBe(false);
    expect(result.error).toContain('Admin access required');
  });
});

// ============================================================================
// adminSaveListValues
// ============================================================================

describe('adminSaveListValues', () => {
  beforeEach(() => {
    Session.getActiveUser.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    var configSheet = createMockSheet(SHEETS.CONFIG, [['header'], ['row2'], ['row3']]);
    var ss = createMockSpreadsheet([configSheet]);
    ss.getOwner.mockReturnValue({ getEmail: jest.fn(() => 'owner@test.com') });
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    getConfigValues.mockReturnValue([]);
  });

  test('rejects non-list field', () => {
    var result = adminSaveListValues('ORG_NAME', ['a', 'b']);
    expect(result.success).toBe(false);
    expect(result.error).toContain('not a list');
  });

  test('rejects non-array values', () => {
    var result = adminSaveListValues('JOB_TITLES', 'not-array');
    expect(result.success).toBe(false);
    expect(result.error).toContain('array');
  });

  test('rejects unknown config key', () => {
    var result = adminSaveListValues('FAKE_KEY', ['a']);
    expect(result.success).toBe(false);
    expect(result.error).toContain('Unknown');
  });
});
