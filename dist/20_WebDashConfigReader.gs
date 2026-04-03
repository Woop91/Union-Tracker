/**
 * ConfigReader.gs — Configuration reader for the web app SPA
 *
 * WHAT THIS FILE DOES:
 *   Configuration reader for the web app SPA. ConfigReader.getConfig() reads
 *   organization-specific settings from the Config tab and returns them as a
 *   JavaScript object. Uses a three-tier caching strategy:
 *     (1) in-execution memo (avoids repeated JSON.parse within same request)
 *     (2) CacheService (5-minute TTL, persists across requests)
 *     (3) sheet read (fallback when cache misses)
 *   Config tab layout: row 2=headers, row 3=values, indexed via CONFIG_COLS
 *   from 01_Core.gs.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   The three-tier cache minimizes expensive sheet reads (each getDataRange()
 *   is an IPC call to Google's server). In-execution memo prevents redundant
 *   cache.get() + JSON.parse() within the same GAS execution. 6-hour
 *   CacheService TTL balances freshness with performance. The IIFE pattern
 *   creates a clean namespace. SPA-specific settings (accent hue, labels) are
 *   not stored in Config — sensible defaults are used to reduce config
 *   complexity.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The SPA can't read organization settings (org name, colors, feature
 *   flags). Falls back to hardcoded defaults which may not match the org's
 *   configuration. If CacheService is down, every page load reads from the
 *   sheet (slower but functional). If the Config sheet is missing, throws an
 *   error and the SPA shows a fatal error page.
 *
 * DEPENDENCIES:
 *   Depends on: CacheService, SpreadsheetApp (GAS built-ins),
 *               01_Core.gs (SHEETS, CONFIG_COLS).
 *   Used by: 22_WebDashApp.gs (injects config into SPA),
 *            21_WebDashDataService.gs, and all SPA views.
 */

var ConfigReader = (function () {

  var CACHE_KEY = 'ORG_CONFIG_v2';
  var CACHE_TTL = 300; // 5 minutes in seconds
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

    // Declared at function scope — used in both the cache-read and cache-write paths
    var cache = CacheService.getScriptCache();

    // Try cache first
    if (!forceRefresh) {
      var cached = cache.get(CACHE_KEY);
      if (cached) {
        try {
          _memo = JSON.parse(cached);
          return _memo;
        } catch (_e) { log_('_e', (_e.message || _e)); }
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
      orgAbbrev:           _readRow(CONFIG_COLS.ORG_ABBREV) || _deriveAbbrev(orgName),
      logoInitials:        _readRow(CONFIG_COLS.LOGO_INITIALS) || _deriveInitials(orgName),
      accentHue:           Number(_readRow(CONFIG_COLS.ACCENT_HUE)) || 250,
      magicLinkExpiryDays: Number(_readRow(CONFIG_COLS.MAGIC_LINK_EXPIRY_DAYS)) || 7,
      cookieDurationDays:  Number(_readRow(CONFIG_COLS.COOKIE_DURATION_DAYS)) || 30,
      stewardLabel:        _readRow(CONFIG_COLS.STEWARD_LABEL) || 'Steward',
      memberLabel:         _readRow(CONFIG_COLS.MEMBER_LABEL) || 'Member',
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

    // Cache it (reuses `cache` declared at function top)
    try {
      cache.put(CACHE_KEY, JSON.stringify(config), CACHE_TTL);
    } catch (e) {
      // Cache write failed — non-fatal, will just re-read next time
      log_('ConfigReader', 'Cache write failed: ' + e.message);
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
