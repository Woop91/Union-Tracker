/**
 * ConfigReader.gs
 * Reads org-specific configuration from the "Config" tab.
 * Caches in CacheService to avoid repeated sheet reads.
 * 
 * Config Tab Expected Layout:
 *   Column A: Setting Name
 *   Column B: Setting Value
 * 
 * Required rows (by name in Column A):
 *   - Org Name
 *   - Org Abbreviation
 *   - Logo Initials
 *   - Accent Hue          (0-360, integer)
 *   - Magic Link Expiry   (days, integer)
 *   - Cookie Duration      (days, integer)
 *   - Steward Label        (optional, default "Steward")
 *   - Member Label         (optional, default "Member")
 */

var ConfigReader = (function () {
  
  var CACHE_KEY = 'ORG_CONFIG';
  var CACHE_TTL = 21600; // 6 hours in seconds
  var CONFIG_SHEET_NAME = 'Config';
  
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
    
    var data = sheet.getDataRange().getValues();
    var configMap = {};
    
    // Build key-value map from columns A and B
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      var value = data[i][1];
      if (key) {
        configMap[key] = value;
      }
    }
    
    // Build config object with defaults
    var config = {
      orgName:             configMap['org name'] || configMap['organization name'] || 'My Organization',
      orgAbbrev:           configMap['org abbreviation'] || configMap['abbreviation'] || '',
      logoInitials:        configMap['logo initials'] || configMap['logo'] || _deriveInitials(configMap['org name'] || 'MO'),
      accentHue:           _parseInt(configMap['accent hue'] || configMap['accent color hue'], 250),
      magicLinkExpiryDays: _parseInt(configMap['magic link expiry'] || configMap['magic link expiry days'], 7),
      cookieDurationDays:  _parseInt(configMap['cookie duration'] || configMap['cookie duration days'], 30),
      stewardLabel:        configMap['steward label'] || 'Steward',
      memberLabel:         configMap['member label'] || 'Member',
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
    
    if (!config.orgName || config.orgName === 'My Organization') missing.push('Org Name');
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
