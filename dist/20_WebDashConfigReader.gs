/**
 * ConfigReader.gs
 * Reads org-specific configuration from the "Config" tab.
 * Caches in CacheService to avoid repeated sheet reads.
 *
 * Config Tab Layout (column-based):
 *   Row 2 = headers, Row 3 = values
 *   Indexed via CONFIG_COLS constants from 01_Core.gs
 *
 * SPA-specific settings (accent hue, labels, magic link / cookie
 * duration) are not in the Config tab — sensible defaults are used.
 */

var ConfigReader = (function () {

  var CACHE_KEY = 'ORG_CONFIG';
  var CACHE_TTL = 21600; // 6 hours in seconds

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
        } catch (_e) {
          // Cache corrupted, fall through to read from sheet
        }
      }
    }

    // Read from sheet — column-based layout via CONFIG_COLS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!sheet) {
      throw new Error('Config tab "' + SHEETS.CONFIG + '" not found. Please create it.');
    }

    // Read org-level values from row 3 using CONFIG_COLS
    var orgName = _readCell(sheet, CONFIG_COLS.ORG_NAME) || 'My Organization';

    var config = {
      orgName:             orgName,
      orgAbbrev:           _deriveAbbrev(orgName),
      logoInitials:        _deriveInitials(orgName),
      accentHue:           250,
      magicLinkExpiryDays: 7,
      cookieDurationDays:  30,
      stewardLabel:        'Steward',
      memberLabel:         'Member',
      // Org links — from Config tab columns
      calendarId:          _readCell(sheet, CONFIG_COLS.CALENDAR_ID) || '',
      driveFolderId:       _readCell(sheet, CONFIG_COLS.DRIVE_FOLDER_ID) || '',
      satisfactionFormUrl: _readCell(sheet, CONFIG_COLS.SATISFACTION_FORM_URL) || '',
      orgWebsite:          _readCell(sheet, CONFIG_COLS.ORG_WEBSITE) || '',
      // Derived (computed below)
      magicLinkExpiryMs:   0,
      cookieDurationMs:    0,
    };

    config.magicLinkExpiryMs = config.magicLinkExpiryDays * 24 * 60 * 60 * 1000;
    config.cookieDurationMs = config.cookieDurationDays * 24 * 60 * 60 * 1000;

    // Cache it
    try {
      cache = CacheService.getScriptCache();
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

    if (!config.orgName || config.orgName === 'My Organization') missing.push('Org Name');
    if (config.accentHue < 0 || config.accentHue > 360) missing.push('Accent Hue (must be 0-360)');

    return {
      valid: missing.length === 0,
      missing: missing,
      config: config
    };
  }

  // --- Helpers ---

  function _readCell(sheet, col) {
    if (!col) return '';
    try {
      return sheet.getRange(3, col).getValue();
    } catch (_e) {
      return '';
    }
  }

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

  function _deriveAbbrev(name) {
    if (!name) return '';
    var words = String(name).trim().split(/\s+/);
    if (words.length <= 2) return name;
    return words.map(function (w) { return w[0]; }).join('').toUpperCase();
  }

  // Public API
  return {
    getConfig: getConfig,
    getConfigJSON: getConfigJSON,
    refreshConfig: refreshConfig,
    validateConfig: validateConfig,
  };

})();
