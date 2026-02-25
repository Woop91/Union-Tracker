/**
 * ConfigReader.gs
 * Reads org-specific configuration from the "Config" tab.
 * Caches in CacheService to avoid repeated sheet reads.
 * 
 * Config Tab Layout (column-based):
 *   Row 1: Headers (matches CONFIG_HEADER_MAP_ keys)
 *   Row 2: Section labels
 *   Row 3+: Values
 * 
 * Uses CONFIG_COLS constants from 01_Core.gs for column positions.
 * Falls back to safe defaults when columns or values are missing.
 */

var ConfigReader = (function () {
  
  var CACHE_KEY = 'ORG_CONFIG';
  var CACHE_TTL = 21600; // 6 hours in seconds
  var CONFIG_SHEET_NAME = 'Config';
  var DATA_ROW = 3; // Values live in row 3 (row 1=headers, row 2=section labels)
  
  /**
   * Reads the Config tab and returns a settings object.
   * Uses CacheService for performance.
   * @param {boolean} [forceRefresh=false] - Bypass cache
   * @returns {Object} Config settings
   */
  function getConfig(forceRefresh) {
    // Try cache first
    if (!forceRefresh) {
      try {
        var cache = CacheService.getScriptCache();
        var cached = cache.get(CACHE_KEY);
        if (cached) {
          return JSON.parse(cached);
        }
      } catch (e) {
        // Cache read failed — non-fatal
      }
    }
    
    // Read from sheet using CONFIG_COLS (column-based layout)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('ConfigReader: Config tab not found, using defaults');
      return _defaults();
    }
    
    // Read entire data row at once for efficiency
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return _defaults();
    
    var dataRow = sheet.getRange(DATA_ROW, 1, 1, lastCol).getValues()[0];
    
    /**
     * Safe column reader — returns value at CONFIG_COLS position or default.
     * CONFIG_COLS are 1-indexed column numbers; dataRow is 0-indexed.
     */
    function _val(colConst, fallback) {
      if (typeof colConst !== 'number' || colConst < 1 || colConst > lastCol) return fallback;
      var v = dataRow[colConst - 1];
      if (v === '' || v === null || v === undefined) return fallback;
      return v;
    }
    
    // Check if CONFIG_COLS exists (loaded from 01_Core.gs)
    var C = (typeof CONFIG_COLS !== 'undefined') ? CONFIG_COLS : {};
    
    var orgName = String(_val(C.ORG_NAME, 'My Organization'));
    
    var config = {
      orgName:             orgName,
      orgAbbrev:           _deriveInitials(orgName),
      logoInitials:        String(_val(C.LOGO_INITIALS, _deriveInitials(orgName))),
      accentHue:           _parseInt(_val(C.ACCENT_HUE, 30), 30),
      magicLinkExpiryDays: _parseInt(_val(C.MAGIC_LINK_EXPIRY_DAYS, 7), 7),
      cookieDurationDays:  _parseInt(_val(C.COOKIE_DURATION_DAYS, 30), 30),
      stewardLabel:        String(_val(C.STEWARD_LABEL, 'Steward')),
      memberLabel:         String(_val(C.MEMBER_LABEL, 'Member')),
      localNumber:         String(_val(C.LOCAL_NUMBER, '')),
      mainPhone:           String(_val(C.MAIN_PHONE, '')),
      // Org links
      calendarId:          String(_val(C.CALENDAR_ID, '')),
      driveFolderId:       String(_val(C.DRIVE_FOLDER_ID, '')),
      satisfactionFormUrl: String(_val(C.SATISFACTION_FORM_URL, '')),
      orgWebsite:          String(_val(C.ORG_WEBSITE, '')),
      // Derived URLs
      calendarUrl:         '',
      driveFolderUrl:      '',
      // Derived ms
      magicLinkExpiryMs:   0,
      cookieDurationMs:    0,
    };
    
    // Build URLs from IDs
    if (config.calendarId) {
      config.calendarUrl = 'https://calendar.google.com/calendar/r?cid=' + encodeURIComponent(config.calendarId);
    }
    if (config.driveFolderId) {
      config.driveFolderUrl = 'https://drive.google.com/drive/folders/' + config.driveFolderId;
    }
    var driveFolderUrl = String(_val(C.DRIVE_FOLDER_URL, ''));
    if (driveFolderUrl && !config.driveFolderUrl) {
      config.driveFolderUrl = driveFolderUrl;
    }
    
    config.magicLinkExpiryMs = config.magicLinkExpiryDays * 86400000;
    config.cookieDurationMs = config.cookieDurationDays * 86400000;
    
    // Cache it
    try {
      var cache = CacheService.getScriptCache();
      cache.put(CACHE_KEY, JSON.stringify(config), CACHE_TTL);
    } catch (e) {
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
  
  // --- Defaults (when Config tab doesn't exist) ---
  
  function _defaults() {
    return {
      orgName: 'My Organization', orgAbbrev: 'MO', logoInitials: 'MO',
      accentHue: 30, magicLinkExpiryDays: 7, cookieDurationDays: 30,
      stewardLabel: 'Steward', memberLabel: 'Member',
      localNumber: '', mainPhone: '',
      calendarId: '', driveFolderId: '', satisfactionFormUrl: '', orgWebsite: '',
      calendarUrl: '', driveFolderUrl: '',
      magicLinkExpiryMs: 604800000, cookieDurationMs: 2592000000,
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
