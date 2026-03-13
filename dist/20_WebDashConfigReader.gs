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
  // S3: In-execution memo — avoids repeated cache.get() + JSON.parse() within same request
  var _memo = null;

  /**
   * Reads the Config tab and returns a settings object.
   * Uses in-execution memo + CacheService for performance.
   * @param {boolean} [forceRefresh=false] - Bypass cache
   * @returns {Object} Config settings
   */
  function getConfig(forceRefresh) {
    // S3: Return in-execution memo if available (same GAS execution)
    if (!forceRefresh && _memo) return _memo;

    // Try cache first
    if (!forceRefresh) {
      var cache = CacheService.getScriptCache();
      var cached = cache.get(CACHE_KEY);
      if (cached) {
        try {
          _memo = JSON.parse(cached);
          return _memo;
        } catch (_e) {
          // Cache corrupted, fall through to read from sheet
        }
      }
    }

    // Read from sheet — column-based layout via CONFIG_COLS
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('Spreadsheet binding broken — getActiveSpreadsheet() returned null. Is this script bound to a spreadsheet?');
    }
    var sheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!sheet) {
      throw new Error('Config tab "' + SHEETS.CONFIG + '" not found. Please create it.');
    }

    // S4: Single-row read — read entire row 3 once instead of 12 individual getRange calls
    var lastCol = sheet.getLastColumn();
    var row3 = lastCol > 0 ? sheet.getRange(3, 1, 1, lastCol).getValues()[0] : [];
    function _readRow(col) {
      if (!col || col < 1 || col > row3.length) return '';
      return row3[col - 1];
    }

    var orgName = _readRow(CONFIG_COLS.ORG_NAME) || 'My Organization';

    var config = {
      orgName:             orgName,
      orgAbbrev:           _deriveAbbrev(orgName),
      logoInitials:        _deriveInitials(orgName),
      accentHue:           250,
      magicLinkExpiryDays: 7,
      cookieDurationDays:  30,
      stewardLabel:        'Steward',
      memberLabel:         'Member',
      // Insights cache TTL in minutes (default 5). Admins can override via Config tab.
      insightsCacheTTLMin: Number(_readRow(CONFIG_COLS.INSIGHTS_CACHE_TTL_MIN) || 5) || 5,
      // Broadcast scope: 'yes' = stewards can send to all members, 'no' (default) = assigned only
      broadcastScopeAll:   _readRow(CONFIG_COLS.BROADCAST_SCOPE_ALL) || 'no',
      // Org links — from Config tab columns
      calendarId:          _readRow(CONFIG_COLS.CALENDAR_ID) || '',
      driveFolderId:       _readRow(CONFIG_COLS.DRIVE_FOLDER_ID) || '',
      // satisfactionFormUrl removed v4.22.7 — survey is native webapp (see member_view.html renderSurveyFormPage)
      orgWebsite:          _readRow(CONFIG_COLS.ORG_WEBSITE) || '',
      // Dashboard folder structure (v4.20.17)
      dashboardRootFolderId:  _readRow(CONFIG_COLS.DASHBOARD_ROOT_FOLDER_ID) || '',
      grievancesFolderId:     _readRow(CONFIG_COLS.GRIEVANCES_FOLDER_ID) || '',
      resourcesFolderId:      _readRow(CONFIG_COLS.RESOURCES_FOLDER_ID) || '',
      minutesFolderId:        _readRow(CONFIG_COLS.MINUTES_FOLDER_ID) || '',
      eventCheckinFolderId:   _readRow(CONFIG_COLS.EVENT_CHECKIN_FOLDER_ID) || '',
      // Derived (computed below)
      magicLinkExpiryMs:   0,
      cookieDurationMs:    0,
    };

    config.magicLinkExpiryMs = config.magicLinkExpiryDays * 24 * 60 * 60 * 1000;
    config.cookieDurationMs = config.cookieDurationDays * 24 * 60 * 60 * 1000;

    // S3: Store in-execution memo
    _memo = config;

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
    _memo = null;
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
