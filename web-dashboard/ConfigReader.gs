/**
 * ConfigReader.gs
 * Reads org-specific configuration from the "Config" tab.
 * Caches in CacheService to avoid repeated sheet reads.
 *
 * Config Tab Layout (column-based):
 *   Row 1: Section headers (grouped categories)
 *   Row 2: Column headers (e.g., "Organization Name", "Accent Hue", etc.)
 *   Row 3+: Data values (row 3 = primary value for single-value settings)
 *
 * This reader maps row 2 headers to their row 3 values.
 */

var ConfigReader = (function () {

  var CACHE_KEY = 'ORG_CONFIG';
  var CACHE_TTL = 21600; // 6 hours in seconds
  var CONFIG_SHEET_NAME = 'Config';

  // Header names must match CONFIG_HEADER_MAP_ in 01_Core.gs
  var HEADER_KEYS = {
    'organization name':       'orgName',
    'logo initials':           'logoInitials',
    'accent hue':              'accentHue',
    'magic link expiry days':  'magicLinkExpiryDays',
    'cookie duration days':    'cookieDurationDays',
    'steward label':           'stewardLabel',
    'member label':            'memberLabel',
    'custom link 1 name':      'customLink1Name',
    'custom link 1 url':       'customLink1URL',
    'custom link 2 name':      'customLink2Name',
    'custom link 2 url':       'customLink2URL'
  };

  /**
   * Reads the Config tab and returns a settings object.
   * Uses CacheService for performance.
   * @param {boolean} [forceRefresh=false] - Bypass cache
   * @returns {Object} Config settings
   */
  function getConfig(forceRefresh) {
    // Try cache first
    if (!forceRefresh) {
      var cache = CacheService.getScriptCache();
      var cached = cache.get(CACHE_KEY);
      if (cached) {
        try {
          return JSON.parse(cached);
        } catch (e) {
          // Cache corrupted, fall through to read from sheet
        }
      }
    }

    // Read from sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

    if (!sheet) {
      throw new Error('Config tab "' + CONFIG_SHEET_NAME + '" not found. Please create it.');
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      throw new Error('Config tab is empty. Please run Setup from the menu.');
    }

    // Row 2 = column headers, Row 3 = primary data values
    var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var values = sheet.getRange(3, 1, 1, lastCol).getValues()[0];

    // Build header-to-value map
    var configMap = {};
    for (var i = 0; i < headers.length; i++) {
      var headerKey = String(headers[i]).trim().toLowerCase();
      if (headerKey && HEADER_KEYS[headerKey]) {
        configMap[HEADER_KEYS[headerKey]] = values[i];
      }
    }

    // Build config object with defaults
    var config = {
      orgName:             configMap.orgName ? String(configMap.orgName).trim() : 'My Organization',
      logoInitials:        configMap.logoInitials ? String(configMap.logoInitials).trim() : _deriveInitials(configMap.orgName || 'MO'),
      accentHue:           _parseInt(configMap.accentHue, 250),
      magicLinkExpiryDays: _parseInt(configMap.magicLinkExpiryDays, 7),
      cookieDurationDays:  _parseInt(configMap.cookieDurationDays, 30),
      stewardLabel:        configMap.stewardLabel ? String(configMap.stewardLabel).trim() : 'Steward',
      memberLabel:         configMap.memberLabel ? String(configMap.memberLabel).trim() : 'Member',
      customLink1Name:     configMap.customLink1Name ? String(configMap.customLink1Name).trim() : '',
      customLink1URL:      configMap.customLink1URL ? String(configMap.customLink1URL).trim() : '',
      customLink2Name:     configMap.customLink2Name ? String(configMap.customLink2Name).trim() : '',
      customLink2URL:      configMap.customLink2URL ? String(configMap.customLink2URL).trim() : '',
      // Derived
      magicLinkExpiryMs:   0,
      cookieDurationMs:    0,
    };

    config.magicLinkExpiryMs = config.magicLinkExpiryDays * 24 * 60 * 60 * 1000;
    config.cookieDurationMs = config.cookieDurationDays * 24 * 60 * 60 * 1000;

    // Cache it
    try {
      var cache = CacheService.getScriptCache();
      cache.put(CACHE_KEY, JSON.stringify(config), CACHE_TTL);
    } catch (e) {
      // Cache write failed — non-fatal, will just re-read next time
      Logger.log('ConfigReader: Cache write failed: ' + e.message);
    }

    return config;
  }

  /**
   * Returns config as JSON string (for injecting into HTML templates)
   * @returns {string} JSON string
   */
  function getConfigJSON() {
    return JSON.stringify(getConfig());
  }

  /**
   * Forces a cache refresh and returns fresh config
   * @returns {Object} Fresh config
   */
  function refreshConfig() {
    return getConfig(true);
  }

  /**
   * Validates that all required config fields are present
   * @returns {Object} { valid: boolean, missing: string[] }
   */
  function validateConfig() {
    var config = getConfig(true);
    var missing = [];

    if (!config.orgName || config.orgName === 'My Organization') missing.push('Organization Name');
    if (config.accentHue < 0 || config.accentHue > 360) missing.push('Accent Hue (must be 0-360)');

    return {
      valid: missing.length === 0,
      missing: missing,
      config: config
    };
  }

  // --- Helpers ---

  function _parseInt(val, defaultVal) {
    var parsed = parseInt(val, 10);
    return isNaN(parsed) ? defaultVal : parsed;
  }

  function _deriveInitials(name) {
    if (!name) return 'ORG';
    var words = String(name).trim().split(/\s+/);
    if (words.length === 1) return words[0].substring(0, 2).toUpperCase();
    return words.map(function (w) { return w[0]; }).join('').substring(0, 3).toUpperCase();
  }

  // Public API
  return {
    getConfig: getConfig,
    getConfigJSON: getConfigJSON,
    refreshConfig: refreshConfig,
    validateConfig: validateConfig,
  };

})();
