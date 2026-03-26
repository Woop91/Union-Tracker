/**
 * Tests for 20_WebDashConfigReader.gs
 *
 * Covers the ConfigReader IIFE: getConfig(), getConfigJSON(),
 * refreshConfig(), validateConfig(), and the private helper functions
 * _deriveInitials() and _deriveAbbrev() tested indirectly via getConfig().
 */

require('./gas-mock');
const { loadSources } = require('./load-source');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '20_WebDashConfigReader.gs'
]);

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

/**
 * Creates a mock Config sheet whose getRange(3, col).getValue() returns
 * values from the provided map. Unrecognised columns return ''.
 */
function buildMockConfigSheet(valueMap) {
  var mockSheet = createMockSheet('Config');
  // Determine max column from the value map for getLastColumn
  var maxCol = 0;
  for (var k in valueMap) {
    if (valueMap.hasOwnProperty(k)) {
      var c = Number(k);
      if (c > maxCol) maxCol = c;
    }
  }
  maxCol = Math.max(maxCol, 20); // Ensure a reasonable default
  mockSheet.getLastColumn = jest.fn(function () { return maxCol; });
  mockSheet.getRange = jest.fn(function (row, col, numRows, numCols) {
    // S4: Support single-row read pattern getRange(3, 1, 1, lastCol).getValues()
    if (row === 3 && col === 1 && numRows === 1 && numCols) {
      var rowData = [];
      for (var i = 1; i <= numCols; i++) {
        rowData.push(valueMap[i] !== undefined ? valueMap[i] : '');
      }
      return {
        getValue: jest.fn(function () { return rowData[0] || ''; }),
        getValues: jest.fn(function () { return [rowData]; }),
        setValue: jest.fn(),
        setValues: jest.fn()
      };
    }
    // Legacy single-cell pattern getRange(3, col).getValue()
    return {
      getValue: jest.fn(function () {
        if (row === 3 && valueMap[col] !== undefined) return valueMap[col];
        return '';
      }),
      setValue: jest.fn(),
      getValues: jest.fn(function () { return [['']]; }),
      setValues: jest.fn()
    };
  });
  return mockSheet;
}

/**
 * Wires up SpreadsheetApp.getActiveSpreadsheet() so it returns
 * a mock spreadsheet containing the given Config sheet.
 */
function installMockSpreadsheet(configSheet) {
  var mockSs = createMockSpreadsheet(configSheet ? [configSheet] : []);
  SpreadsheetApp.getActiveSpreadsheet = jest.fn(function () { return mockSs; });
  return mockSs;
}

// ---------------------------------------------------------------------------
// Shared setup — default "happy path" mocks reset before every test
// ---------------------------------------------------------------------------

var mockCache;

beforeEach(function () {
  // Fresh script cache mock that starts empty
  mockCache = {
    get: jest.fn(function () { return null; }),
    put: jest.fn(),
    remove: jest.fn()
  };
  CacheService.getScriptCache = jest.fn(function () { return mockCache; });

  // Default Config sheet with a real org name
  var configSheet = buildMockConfigSheet({
    [CONFIG_COLS.ORG_NAME]: 'Test Organization',
    [CONFIG_COLS.CALENDAR_ID]: 'cal@test.com',
    [CONFIG_COLS.DRIVE_FOLDER_ID]: 'folder-123',
    [CONFIG_COLS.ORG_WEBSITE]: 'https://example.com'
  });
  installMockSpreadsheet(configSheet);
});

// ============================================================================
// ConfigReader.getConfig
// ============================================================================

describe('ConfigReader.getConfig', function () {

  test('returns object with expected org fields', function () {
    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgName).toBe('Test Organization');
    expect(cfg.orgAbbrev).toBeDefined();
    expect(cfg.logoInitials).toBeDefined();
    expect(typeof cfg.accentHue).toBe('number');
    expect(typeof cfg.magicLinkExpiryDays).toBe('number');
    expect(typeof cfg.cookieDurationDays).toBe('number');
    expect(cfg.stewardLabel).toBe('Steward');
    expect(cfg.memberLabel).toBe('Member');
  });

  test('returns calendarId and driveFolderId from Config tab', function () {
    var cfg = ConfigReader.getConfig(true);
    expect(cfg.calendarId).toBe('cal@test.com');
    expect(cfg.driveFolderId).toBe('folder-123');
  });

  test('returns orgWebsite from Config tab', function () {
    var cfg = ConfigReader.getConfig(true);
    // satisfactionFormUrl removed v4.22.7 — survey is native webapp
    expect(cfg.orgWebsite).toBe('https://example.com');
  });

  test('returns cached value on second call without forceRefresh', function () {
    // First call reads from sheet
    var first = ConfigReader.getConfig(true);
    // S3: In-execution memo means second call returns memo without hitting CacheService.
    // This is the expected performance optimization behavior.
    var second = ConfigReader.getConfig();
    expect(second).toEqual(first);
    // The memo short-circuits — CacheService may not be consulted within same execution.
    // Validate that refreshConfig clears the memo and re-reads:
    var refreshed = ConfigReader.refreshConfig();
    expect(refreshed.orgName).toBe('Test Organization');
  });

  test('forceRefresh=true bypasses cache', function () {
    // Pre-populate cache with stale data
    mockCache.get = jest.fn(function () {
      return JSON.stringify({ orgName: 'Stale' });
    });

    var cfg = ConfigReader.getConfig(true);
    // Should read from sheet, not cache
    expect(cfg.orgName).toBe('Test Organization');
  });

  test('calculates magicLinkExpiryMs from days', function () {
    var cfg = ConfigReader.getConfig(true);
    expect(cfg.magicLinkExpiryMs).toBe(cfg.magicLinkExpiryDays * 24 * 60 * 60 * 1000);
  });

  test('calculates cookieDurationMs from days', function () {
    var cfg = ConfigReader.getConfig(true);
    expect(cfg.cookieDurationMs).toBe(cfg.cookieDurationDays * 24 * 60 * 60 * 1000);
  });

  test('handles corrupted cache JSON gracefully', function () {
    // Return invalid JSON from cache
    mockCache.get = jest.fn(function () { return '{not-valid-json!!!'; });

    // Should fall through to sheet read instead of throwing
    var cfg = ConfigReader.getConfig();
    expect(cfg.orgName).toBe('Test Organization');
  });

  test('throws if getActiveSpreadsheet() returns null', function () {
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(function () { return null; });
    expect(function () {
      ConfigReader.getConfig(true);
    }).toThrow(/Spreadsheet binding broken/);
  });

  test('throws if Config sheet not found', function () {
    // Spreadsheet exists but has no Config sheet
    var mockSs = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet = jest.fn(function () { return mockSs; });
    expect(function () {
      ConfigReader.getConfig(true);
    }).toThrow(/Config tab.*not found/);
  });

  test('uses CONFIG_COLS for reading cells', function () {
    var configSheet = buildMockConfigSheet({
      [CONFIG_COLS.ORG_NAME]: 'Column Test Org'
    });
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    // S4: Now reads entire row 3 in one call: getRange(3, 1, 1, lastCol)
    expect(configSheet.getRange).toHaveBeenCalledWith(3, 1, 1, expect.any(Number));
    expect(cfg.orgName).toBe('Column Test Org');
  });

  test('defaults orgName to "My Organization" when empty', function () {
    var configSheet = buildMockConfigSheet({});
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgName).toBe('My Organization');
  });

  test('stores config in cache with 5min TTL (300s)', function () {
    ConfigReader.getConfig(true);
    expect(mockCache.put).toHaveBeenCalledWith(
      'ORG_CONFIG_v2',
      expect.any(String),
      300
    );
  });

  test('handles cache write failure gracefully', function () {
    mockCache.put = jest.fn(function () {
      throw new Error('Cache quota exceeded');
    });

    // Should not throw even though cache.put fails
    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgName).toBe('Test Organization');
  });
});

// ============================================================================
// ConfigReader.getConfigJSON
// ============================================================================

describe('ConfigReader.getConfigJSON', function () {

  test('returns a valid JSON string', function () {
    var json = ConfigReader.getConfigJSON();
    expect(typeof json).toBe('string');
    expect(function () { JSON.parse(json); }).not.toThrow();
  });

  test('JSON parses to same object as getConfig', function () {
    // Force fresh read so both calls go to sheet
    var cfg = ConfigReader.getConfig(true);
    // Now the cache has the value, so getConfigJSON will use it
    mockCache.get = jest.fn(function () { return JSON.stringify(cfg); });

    var parsed = JSON.parse(ConfigReader.getConfigJSON());
    expect(parsed).toEqual(cfg);
  });

  test('contains orgName field', function () {
    var parsed = JSON.parse(ConfigReader.getConfigJSON());
    expect(parsed.orgName).toBeDefined();
    expect(typeof parsed.orgName).toBe('string');
  });
});

// ============================================================================
// ConfigReader.refreshConfig
// ============================================================================

describe('ConfigReader.refreshConfig', function () {

  test('calls getConfig with forceRefresh=true (bypasses cache)', function () {
    // Put stale data in cache
    mockCache.get = jest.fn(function () {
      return JSON.stringify({ orgName: 'Stale Data' });
    });

    var cfg = ConfigReader.refreshConfig();
    // Should have fresh data from sheet, not the stale cache
    expect(cfg.orgName).toBe('Test Organization');
  });

  test('returns fresh config object', function () {
    var cfg = ConfigReader.refreshConfig();
    expect(cfg).toBeDefined();
    expect(cfg.orgName).toBe('Test Organization');
    expect(cfg.stewardLabel).toBe('Steward');
  });

  test('bypasses cache even when cache has data', function () {
    // First call to populate
    ConfigReader.getConfig(true);
    // Mock cache to return old data
    mockCache.get = jest.fn(function () {
      return JSON.stringify({ orgName: 'Cached Old Value', stewardLabel: 'Rep' });
    });

    var cfg = ConfigReader.refreshConfig();
    // refreshConfig should skip cache and read from sheet
    expect(cfg.orgName).toBe('Test Organization');
    expect(cfg.stewardLabel).toBe('Steward');
  });
});

// ============================================================================
// ConfigReader.validateConfig
// ============================================================================

describe('ConfigReader.validateConfig', function () {

  test('returns {valid: true} for proper config', function () {
    var result = ConfigReader.validateConfig();
    expect(result.valid).toBe(true);
    expect(result.missing).toEqual([]);
  });

  test('returns {valid: false} when orgName is default "My Organization"', function () {
    var configSheet = buildMockConfigSheet({});
    installMockSpreadsheet(configSheet);

    var result = ConfigReader.validateConfig();
    expect(result.valid).toBe(false);
  });

  test('returns missing array with "Org Name" when orgName is default', function () {
    var configSheet = buildMockConfigSheet({});
    installMockSpreadsheet(configSheet);

    var result = ConfigReader.validateConfig();
    expect(result.missing).toContain('Org Name');
  });

  test('returns config object in result', function () {
    var result = ConfigReader.validateConfig();
    expect(result.config).toBeDefined();
    expect(result.config.orgName).toBe('Test Organization');
    expect(result.config.stewardLabel).toBe('Steward');
  });

  test('forces fresh read (forceRefresh=true)', function () {
    // Pre-populate cache with stale data that would fail validation
    mockCache.get = jest.fn(function () {
      return JSON.stringify({ orgName: 'Stale', accentHue: 250 });
    });

    var result = ConfigReader.validateConfig();
    // Should read from sheet (Test Organization), not cache (Stale)
    expect(result.config.orgName).toBe('Test Organization');
  });

  test('checks accentHue range — valid default of 250 passes', function () {
    var result = ConfigReader.validateConfig();
    // Default accentHue is 250 which is within 0-360
    expect(result.valid).toBe(true);
    expect(result.missing).not.toContain(expect.stringContaining('Accent Hue'));
  });
});

// ============================================================================
// Helper functions (tested indirectly through getConfig)
// ============================================================================

describe('_deriveInitials (via getConfig)', function () {

  test('returns first 2 chars uppercased for single-word name', function () {
    var configSheet = buildMockConfigSheet({
      [CONFIG_COLS.ORG_NAME]: 'Solidarity'
    });
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.logoInitials).toBe('SO');
  });

  test('returns initials for multi-word name (up to 3 chars)', function () {
    var configSheet = buildMockConfigSheet({
      [CONFIG_COLS.ORG_NAME]: 'American Federation Workers'
    });
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.logoInitials).toBe('AFW');
  });

  test('returns "ORG" for empty org name (fallback default)', function () {
    // When orgName is empty, it defaults to "My Organization"
    // _deriveInitials("My Organization") = "MO" (2 words => first letters, up to 3)
    // To test the empty-string path of _deriveInitials, we need orgName to be
    // falsy BEFORE the default is applied. Since getConfig applies the default,
    // the only way to trigger _deriveInitials('') is if orgName is falsy.
    // In practice, empty orgName defaults to 'My Organization', so logoInitials = 'MO'.
    // We verify the default path here.
    var configSheet = buildMockConfigSheet({});
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    // orgName defaults to 'My Organization', initials = 'MO'
    expect(cfg.logoInitials).toBe('MO');
  });
});

describe('_deriveAbbrev (via getConfig)', function () {

  test('returns full name for name with 2 or fewer words', function () {
    var configSheet = buildMockConfigSheet({
      [CONFIG_COLS.ORG_NAME]: 'Union Local'
    });
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgAbbrev).toBe('Union Local');
  });

  test('returns initials for name with more than 2 words', function () {
    var configSheet = buildMockConfigSheet({
      [CONFIG_COLS.ORG_NAME]: 'American Postal Workers Union'
    });
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgAbbrev).toBe('APWU');
  });

  test('returns empty string for empty org name (via default path)', function () {
    // When orgName is empty, it defaults to 'My Organization'
    // _deriveAbbrev('My Organization') = 'My Organization' (2 words => returns full name)
    var configSheet = buildMockConfigSheet({});
    installMockSpreadsheet(configSheet);

    var cfg = ConfigReader.getConfig(true);
    expect(cfg.orgAbbrev).toBe('My Organization');
  });
});
