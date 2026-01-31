/**
 * ============================================================================
 * CacheManager.gs - Caching and Performance
 * ============================================================================
 *
 * This module handles all caching-related functions including:
 * - Multi-layer caching (memory + properties)
 * - Cache warming and invalidation
 * - Cached data loaders for common queries
 * - Cache status dashboard
 *
 * REFACTORED: Split from 06_Maintenance.gs for better maintainability
 *
 * @fileoverview Caching and performance optimization functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// CACHE CONFIGURATION
// ============================================================================

/**
 * Cache configuration constants
 * @const {Object}
 */
var CACHE_CONFIG = {
  /** Memory cache TTL in seconds */
  MEMORY_TTL: 300,
  /** Properties cache TTL in seconds */
  PROPS_TTL: 3600,
  /** Enable debug logging */
  ENABLE_LOGGING: false
};

/**
 * Cache key constants
 * @const {Object}
 */
var CACHE_KEYS = {
  ALL_GRIEVANCES: 'cache_grievances',
  ALL_MEMBERS: 'cache_members',
  ALL_STEWARDS: 'cache_stewards',
  DASHBOARD_METRICS: 'cache_metrics',
  CONFIG_VALUES: 'cache_config'
};

// ============================================================================
// CORE CACHING FUNCTIONS
// ============================================================================

/**
 * Gets data from cache or loads it using the provided loader function
 * Implements two-layer caching: memory (fast) -> properties (persistent)
 * @param {string} key - Cache key
 * @param {Function} loader - Function to load data if not cached
 * @param {number} [ttl=CACHE_CONFIG.MEMORY_TTL] - Time to live in seconds
 * @returns {*} The cached or loaded data
 */
function getCachedData(key, loader, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;

  try {
    // Check memory cache first (fastest)
    var memCache = CacheService.getScriptCache();
    var cached = memCache.get(key);
    if (cached) {
      if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT - Memory] ' + key);
      return JSON.parse(cached);
    }

    // Check properties cache (persistent)
    var propsCache = PropertiesService.getScriptProperties();
    var propsCached = propsCache.getProperty(key);
    if (propsCached) {
      var obj = JSON.parse(propsCached);
      if (obj.timestamp && (Date.now() - obj.timestamp) < (ttl * 1000)) {
        // Refresh memory cache from properties
        memCache.put(key, JSON.stringify(obj.data), ttl);
        if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT - Props] ' + key);
        return obj.data;
      }
    }

    // Cache miss - load fresh data
    if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE MISS] ' + key);
    var data = loader();
    setCachedData(key, data, ttl);
    return data;

  } catch (e) {
    Logger.log('Cache error for ' + key + ': ' + e.message);
    // Fallback to direct load on error
    return loader();
  }
}

/**
 * Sets data in both memory and properties cache
 * @param {string} key - Cache key
 * @param {*} data - Data to cache
 * @param {number} [ttl=CACHE_CONFIG.MEMORY_TTL] - Time to live in seconds
 * @returns {void}
 */
function setCachedData(key, data, ttl) {
  ttl = ttl || CACHE_CONFIG.MEMORY_TTL;

  try {
    var str = JSON.stringify(data);

    // Store in memory cache (max 6 hours)
    var memCache = CacheService.getScriptCache();
    memCache.put(key, str, Math.min(ttl, 21600));

    // Store in properties if under size limit (100KB)
    if (str.length < 100000) {
      var propsCache = PropertiesService.getScriptProperties();
      propsCache.setProperty(key, JSON.stringify({
        data: data,
        timestamp: Date.now()
      }));
    }
  } catch (e) {
    Logger.log('Set cache error for ' + key + ': ' + e.message);
  }
}

/**
 * Invalidates a specific cache key
 * @param {string} key - Cache key to invalidate
 * @returns {void}
 */
function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
    PropertiesService.getScriptProperties().deleteProperty(key);
    if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('Cache invalidated: ' + key);
  } catch (e) {
    Logger.log('Invalidate error for ' + key + ': ' + e.message);
  }
}

/**
 * Invalidates all caches
 * @returns {void}
 */
function invalidateAllCaches() {
  try {
    var keys = Object.keys(CACHE_KEYS).map(function(k) { return CACHE_KEYS[k]; });
    CacheService.getScriptCache().removeAll(keys);

    var props = PropertiesService.getScriptProperties();
    keys.forEach(function(k) { props.deleteProperty(k); });

    SpreadsheetApp.getActiveSpreadsheet().toast('All caches cleared', 'Cache', 3);
  } catch (e) {
    Logger.log('Clear all caches error: ' + e.message);
  }
}

/**
 * Warms up all caches by pre-loading common data
 * @returns {void}
 */
function warmUpCaches() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Warming caches...', 'Cache', -1);

  try {
    getCachedGrievances();
    getCachedMembers();
    getCachedStewards();
    getCachedDashboardMetrics();
    SpreadsheetApp.getActiveSpreadsheet().toast('Caches warmed successfully', 'Cache', 3);
  } catch (e) {
    Logger.log('Cache warmup error: ' + e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Cache warmup failed: ' + e.message, 'Error', 5);
  }
}

// ============================================================================
// CACHED DATA LOADERS
// ============================================================================

/**
 * Gets all grievances with caching
 * @returns {Array<Array>} 2D array of grievance data
 */
function getCachedGrievances() {
  return getCachedData(CACHE_KEYS.ALL_GRIEVANCES, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 300);
}

/**
 * Gets all members with caching
 * @returns {Array<Array>} 2D array of member data
 */
function getCachedMembers() {
  return getCachedData(CACHE_KEYS.ALL_MEMBERS, function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  }, 600);
}

/**
 * Gets all stewards with caching
 * @returns {Array<Array>} 2D array of steward data
 */
function getCachedStewards() {
  return getCachedData(CACHE_KEYS.ALL_STEWARDS, function() {
    var members = getCachedMembers();
    return members.filter(function(row) {
      return row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes';
    });
  }, 600);
}

/**
 * Gets dashboard metrics with caching
 * @returns {Object} Dashboard metrics object
 */
function getCachedDashboardMetrics() {
  return getCachedData(CACHE_KEYS.DASHBOARD_METRICS, function() {
    var grievances = getCachedGrievances();

    var metrics = {
      total: grievances.length,
      open: 0,
      closed: 0,
      overdue: 0,
      byStatus: {},
      byIssueType: {},
      bySteward: {}
    };

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    grievances.forEach(function(row) {
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var issue = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
      var steward = row[GRIEVANCE_COLS.STEWARD - 1];
      var deadline = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];

      // Calculate days to deadline
      var daysTo = null;
      if (deadline) {
        daysTo = Math.floor((new Date(deadline) - today) / (1000 * 60 * 60 * 24));
      }

      // Update counts
      if (status === 'Open') metrics.open++;
      if (status === 'Closed' || status === 'Resolved') metrics.closed++;
      if (daysTo !== null && daysTo < 0) metrics.overdue++;

      // Update breakdowns
      metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;
      if (issue) metrics.byIssueType[issue] = (metrics.byIssueType[issue] || 0) + 1;
      if (steward) metrics.bySteward[steward] = (metrics.bySteward[steward] || 0) + 1;
    });

    return metrics;
  }, 180);
}

// ============================================================================
// CACHE STATUS DASHBOARD
// ============================================================================

/**
 * Shows the cache status dashboard
 * @returns {void}
 */
function showCacheStatusDashboard() {
  var memCache = CacheService.getScriptCache();
  var propsCache = PropertiesService.getScriptProperties();

  var rows = Object.keys(CACHE_KEYS).map(function(name) {
    var key = CACHE_KEYS[name];
    var inMem = memCache.get(key) !== null;
    var inProps = propsCache.getProperty(key) !== null;
    var age = 'N/A';

    if (inProps) {
      try {
        var obj = JSON.parse(propsCache.getProperty(key));
        if (obj.timestamp) {
          age = Math.floor((Date.now() - obj.timestamp) / 1000) + 's';
        }
      } catch (e) {}
    }

    return '<tr>' +
      '<td>' + name + '</td>' +
      '<td>' + (inMem ? '✅' : '❌') + '</td>' +
      '<td>' + (inProps ? '✅' : '❌') + '</td>' +
      '<td>' + age + '</td>' +
      '</tr>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}' +
    'h2{color:#7c3aed;margin-bottom:20px}' +
    'table{width:100%;border-collapse:collapse;margin:20px 0}' +
    'th{background:#7c3aed;color:white;padding:12px;text-align:left}' +
    'td{padding:10px;border-bottom:1px solid #e0e0e0}' +
    '.btn{padding:12px 24px;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin:5px}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-danger{background:#ef4444;color:white}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>🗄️ Cache Status Dashboard</h2>' +
    '<table>' +
    '<tr><th>Cache</th><th>Memory</th><th>Props</th><th>Age</th></tr>' +
    rows +
    '</table>' +
    '<button class="btn btn-primary" onclick="warmUp()">🔥 Warm Up Caches</button>' +
    '<button class="btn btn-danger" onclick="clearAll()">🗑️ Clear All Caches</button>' +
    '</div>' +
    '<script>' +
    'function warmUp(){google.script.run.withSuccessHandler(function(){location.reload()}).warmUpCaches()}' +
    'function clearAll(){google.script.run.withSuccessHandler(function(){location.reload()}).invalidateAllCaches()}' +
    '</script></body></html>'
  ).setWidth(600).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '🗄️ Cache Status');
}
