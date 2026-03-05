/**
 * ============================================================================
 * 06_Maintenance.gs - System Diagnostics and Repair
 * ============================================================================
 *
 * This module handles all diagnostic and repair functions including:
 * - System diagnostics (DIAGNOSE_SETUP)
 * - Dashboard repair (REPAIR_DASHBOARD)
 * - Modal diagnostics
 * - Sheet verification
 *
 * REFACTORED: Split from 06_Maintenance.gs for better maintainability
 *
 * @fileoverview System diagnostics and repair functions
 * @version 4.7.0
 * @requires 01_Core.gs
 */

// ============================================================================
// SYSTEM DIAGNOSTICS
// ============================================================================

/**
 * Runs a complete diagnostic check on the dashboard setup
 * Checks all required sheets, columns, formulas, and configurations
 * @returns {Object} Diagnostic results with status, checks, warnings, and errors
 */
// PERF: Diagnostic reads run infrequently (manual trigger), bulk read acceptable
function DIAGNOSE_SETUP() {
  var results = {
    timestamp: new Date().toISOString(),
    status: 'OK',
    checks: [],
    warnings: [],
    errors: []
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check 1: Required sheets exist
  // Skip hidden sheets (checked in Check 2), deprecated/test-only/alias entries
  results.checks.push('Checking required sheets...');
  var existingSheets = ss.getSheets().map(function(s) { return s.getName(); });

  var skipKeys = {
    DASHBOARD: true,           // @deprecated v4.3.2 - modal dashboards now
    REPORTS: true,             // Backward-compat alias for DASHBOARD
    TEST_RESULTS: true,        // Created on-demand by test framework only
    GRIEVANCE_TRACKER: true,   // Backward-compat alias for GRIEVANCE_LOG
    MEMBER_DIRECTORY: true     // Backward-compat alias for MEMBER_DIR
  };

  var checkedValues = {};
  for (var key in SHEET_NAMES) {
    if (skipKeys[key]) continue;
    var sheetName = SHEET_NAMES[key];
    // Hidden sheets (underscore prefix) are checked separately in Check 2
    if (sheetName.charAt(0) === '_') continue;
    // Avoid reporting the same sheet name twice (e.g. aliases)
    if (checkedValues[sheetName]) continue;
    checkedValues[sheetName] = true;
    if (existingSheets.indexOf(sheetName) === -1) {
      results.errors.push('Missing required sheet: ' + sheetName);
      results.status = 'ERROR';
    }
  }

  // Check 2: Hidden sheets exist
  results.checks.push('Checking hidden calculation sheets...');
  for (var hiddenKey in HIDDEN_SHEETS) {
    var hiddenName = HIDDEN_SHEETS[hiddenKey];
    if (existingSheets.indexOf(hiddenName) === -1) {
      results.warnings.push('Missing hidden sheet: ' + hiddenName + ' (will be auto-created on repair)');
      if (results.status === 'OK') results.status = 'WARNING';
    }
  }

  // Check 3: Member Directory structure
  results.checks.push('Verifying Member Directory structure...');
  var memberSheet = ss.getSheetByName(SHEET_NAMES.MEMBER_DIRECTORY);
  if (memberSheet) {
    var headers = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    var requiredHeaders = [
      'Member ID', 'First Name', 'Last Name', 'Job Title',
      'Work Location', 'Unit', 'Email', 'Phone', 'Is Steward', 'Assigned Steward',
      'Employee ID', 'Department', 'Hire Date'
    ];

    requiredHeaders.forEach(function(header) {
      var found = headers.some(function(h) {
        return h.toString().toLowerCase().indexOf(header.toLowerCase()) !== -1;
      });
      if (!found) {
        results.warnings.push('Member Directory may be missing column: ' + header);
      }
    });
  }

  // Check 4: Grievance Tracker structure
  results.checks.push('Verifying Grievance Tracker structure...');
  var grievanceSheet = ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);
  if (grievanceSheet) {
    var gHeaders = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues()[0];
    var requiredGHeaders = [
      'Grievance ID', 'Member ID', 'Date Filed', 'Issue Category', 'Current Step',
      'Status', 'Next Action Due', 'Last Updated'
    ];

    requiredGHeaders.forEach(function(header) {
      var found = gHeaders.some(function(h) {
        return h.toString().toLowerCase().indexOf(header.toLowerCase()) !== -1;
      });
      if (!found) {
        results.warnings.push('Grievance Tracker may be missing column: ' + header);
      }
    });
  }

  // Check 5: Data integrity
  results.checks.push('Checking data integrity...');
  if (grievanceSheet) {
    var data = grievanceSheet.getDataRange().getValues();
    var orphanedGrievances = 0;
    var invalidDates = 0;

    for (var i = 1; i < data.length; i++) {
      var memberId = data[i][GRIEVANCE_COLS.MEMBER_ID - 1];
      var filingDate = data[i][GRIEVANCE_COLS.DATE_FILED - 1];

      if (!memberId && data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1]) {
        orphanedGrievances++;
      }

      if (filingDate && !(filingDate instanceof Date)) {
        invalidDates++;
      }
    }

    if (orphanedGrievances > 0) {
      results.warnings.push('Found ' + orphanedGrievances + ' grievances without member ID');
    }
    if (invalidDates > 0) {
      results.warnings.push('Found ' + invalidDates + ' grievances with invalid filing dates');
    }
  }

  // Check 6: Calendar permissions
  results.checks.push('Checking Calendar permissions...');
  try {
    CalendarApp.getAllCalendars();
    results.checks.push('Calendar access: OK');
  } catch (_e) {
    results.warnings.push('Cannot access Calendar - sync features may not work');
  }

  // Check 7: Drive permissions
  results.checks.push('Checking Drive permissions...');
  try {
    DriveApp.getRootFolder();
    results.checks.push('Drive access: OK');
  } catch (_e) {
    results.warnings.push('Cannot access Drive - folder features may not work');
  }

  // Finalize status
  if (results.errors.length > 0) {
    results.status = 'ERROR';
  } else if (results.warnings.length > 0) {
    results.status = 'WARNING';
  }

  results.summary = 'Completed ' + results.checks.length + ' checks: ' +
                    results.errors.length + ' errors, ' + results.warnings.length + ' warnings';

  // Log the diagnostic run
  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'DIAGNOSTICS',
    status: results.status,
    errors: results.errors.length,
    warnings: results.warnings.length,
    runBy: Session.getActiveUser().getEmail()
  });

  return results;
}

// ============================================================================
// REPAIR FUNCTIONS
// ============================================================================

/**
 * Repairs the dashboard by recreating missing sheets and fixing issues
 * @returns {void}
 */
function REPAIR_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    '🔧 Repair Dashboard',
    'This will:\n' +
    '• Recreate missing hidden sheets\n' +
    '• Reapply formulas and validations\n' +
    '• Fix broken references\n' +
    '• Install required triggers\n\n' +
    'Your data will NOT be deleted.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Repair cancelled.');
    return;
  }

  ss.toast('Starting repair...', '🔧 Repair', 5);

  try {
    // Recreate hidden sheets
    ss.toast('Recreating hidden sheets...', '🔧 Progress', 3);
    setupHiddenSheets(ss);

    // Reapply data validations
    ss.toast('Reapplying data validations...', '🔧 Progress', 3);
    setupDataValidations();

    // Install triggers
    ss.toast('Installing triggers...', '🔧 Progress', 3);
    installAutoSyncTrigger();

    // Sync data
    ss.toast('Syncing data...', '🔧 Progress', 3);
    syncAllData();

    ss.toast('Repair complete!', '✅ Success', 5);
    ui.alert('✅ Repair Complete',
      'Dashboard has been repaired successfully!\n\n' +
      '• Hidden sheets recreated\n' +
      '• Validations reapplied\n' +
      '• Triggers installed\n' +
      '• Data synced',
      ui.ButtonSet.OK);

    // Log the repair
    logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
      action: 'REPAIR_DASHBOARD',
      runBy: Session.getActiveUser().getEmail()
    });

  } catch (error) {
    Logger.log('Error in REPAIR_DASHBOARD: ' + error.message);
    ui.alert('❌ Error', 'Repair failed: ' + error.message, ui.ButtonSet.OK);

    logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
      action: 'REPAIR_DASHBOARD',
      status: 'FAILED',
      error: error.message,
      runBy: Session.getActiveUser().getEmail()
    });
  }
}

/**
 * Removes deprecated tabs from the spreadsheet
 * @returns {void}
 */
function removeDeprecatedTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Exact-match names — only delete if name matches exactly
  var exactMatches = [
    'Dashboard_OLD',
    'Member List_OLD'
  ];
  // Prefix patterns — delete if name starts with the pattern (intentional wildcard)
  var prefixPatterns = [
    'DEPRECATED_BACKUP_',
    'DEPRECATED_TEST_'
  ];

  var removed = [];
  var sheets = ss.getSheets();
  for (var i = sheets.length - 1; i >= 0; i--) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var shouldRemove = false;

    // Check exact matches
    if (exactMatches.indexOf(name) !== -1) {
      shouldRemove = true;
    } else {
      // Check prefix patterns
      for (var j = 0; j < prefixPatterns.length; j++) {
        if (name.indexOf(prefixPatterns[j]) === 0) {
          shouldRemove = true;
          break;
        }
      }
    }

    if (shouldRemove) {
      try {
        ss.deleteSheet(sheet);
        removed.push(name);
      } catch (_e) {
        Logger.log('Could not delete sheet: ' + name);
      }
    }
  }

  if (removed.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Removed ' + removed.length + ' deprecated sheets:\n' + removed.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('No deprecated sheets found.');
  }
}

// setupSheetStructure stub removed - initializeDashboard now delegates to CREATE_DASHBOARD()

/**
 * Shows the repair dialog
 * @returns {void}
 */
function showRepairDialog() {
  var diagnostics = DIAGNOSE_SETUP();

  var html = HtmlService.createHtmlOutput(getRepairDialogHtml_(diagnostics))
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, '🔧 Repair Dashboard');
}

/**
 * Generates HTML for repair dialog
 * @param {Object} diagnostics - Diagnostic results
 * @returns {string} HTML content
 * @private
 */
function getRepairDialogHtml_(diagnostics) {
  var statusColor = diagnostics.status === 'OK' ? '#10b981' :
                    diagnostics.status === 'WARNING' ? '#f59e0b' : '#ef4444';

  return '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;margin:0}' +
    '.status{padding:10px;border-radius:8px;text-align:center;margin-bottom:20px;color:white;background:' + statusColor + '}' +
    '.section{margin-bottom:15px}' +
    '.section-title{font-weight:600;margin-bottom:8px}' +
    '.item{padding:4px 0;font-size:13px}' +
    '.error{color:#ef4444}' +
    '.warning{color:#f59e0b}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin-right:10px}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151}' +
    '</style></head><body>' +
    '<div class="status">' + escapeHtml(diagnostics.status) + '</div>' +
    '<p>' + escapeHtml(diagnostics.summary) + '</p>' +
    (diagnostics.errors.length > 0 ?
      '<div class="section"><div class="section-title error">Errors</div>' +
      diagnostics.errors.map(function(e) { return '<div class="item error">' + escapeHtml(e) + '</div>'; }).join('') +
      '</div>' : '') +
    (diagnostics.warnings.length > 0 ?
      '<div class="section"><div class="section-title warning">Warnings</div>' +
      diagnostics.warnings.map(function(w) { return '<div class="item warning">' + escapeHtml(w) + '</div>'; }).join('') +
      '</div>' : '') +
    '<div style="margin-top:20px;text-align:right">' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="runRepair()">Run Repair</button>' +
    '</div>' +
    '<script>' +
    'function runRepair(){google.script.run.withSuccessHandler(function(){google.script.host.close()}).REPAIR_DASHBOARD()}' +
    '</script></body></html>';
}

// ============================================================================
// MODAL DIAGNOSTICS
// ============================================================================

/**
 * Diagnoses why modals might not be loading data
 * @returns {Object} Diagnostic results
 */
function diagnoseModalIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = {
    status: 'OK',
    sheetChecks: [],
    dataChecks: [],
    modalTests: [],
    errors: [],
    warnings: []
  };

  var actualSheets = ss.getSheets().map(function(s) { return s.getName(); });

  // Check required sheets
  var requiredSheets = {
    'Member Directory': SHEETS.MEMBER_DIR,
    'Grievance Log': SHEETS.GRIEVANCE_LOG,
    'Member Satisfaction': SHEETS.SATISFACTION,
    'Config': SHEETS.CONFIG
  };

  for (var displayName in requiredSheets) {
    var expectedName = requiredSheets[displayName];
    var found = actualSheets.indexOf(expectedName) !== -1;
    results.sheetChecks.push({
      name: expectedName,
      expected: expectedName,
      found: found,
      status: found ? 'OK' : 'MISSING'
    });
    if (!found) {
      results.errors.push('Sheet "' + expectedName + '" not found.');
      results.status = 'ERROR';
    }
  }

  // Check data structure
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet) {
    var memberCols = memberSheet.getLastColumn();
    var memberRows = memberSheet.getLastRow();
    results.dataChecks.push({
      sheet: 'Member Directory',
      columns: memberCols,
      rows: memberRows,
      dataRows: Math.max(0, memberRows - 1),
      status: memberCols >= 20 ? 'OK' : 'WARNING'
    });
  }

  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet) {
    var grievCols = grievanceSheet.getLastColumn();
    var grievRows = grievanceSheet.getLastRow();
    results.dataChecks.push({
      sheet: 'Grievance Log',
      columns: grievCols,
      rows: grievRows,
      dataRows: Math.max(0, grievRows - 1),
      status: grievCols >= 20 ? 'OK' : 'WARNING'
    });
  }

  results.summary = results.status === 'OK'
    ? 'All modal checks passed.'
    : results.status === 'WARNING'
    ? 'Some warnings detected.'
    : 'Errors detected.';

  return results;
}

/**
 * Shows modal diagnostic results in a dialog
 * @returns {void}
 */
function showModalDiagnostics() {
  var results = diagnoseModalIssues();

  var html = '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px}' +
    '.card{background:#f9fafb;border-radius:8px;padding:15px;margin-bottom:15px}' +
    '.title{font-size:18px;font-weight:600;margin-bottom:15px}' +
    '.status-ok{color:#059669;background:#d1fae5;padding:4px 12px;border-radius:20px}' +
    '.status-error{color:#dc2626;background:#fee2e2;padding:4px 12px;border-radius:20px}' +
    'table{width:100%;border-collapse:collapse;font-size:13px}' +
    'th,td{padding:8px;text-align:left;border-bottom:1px solid #e5e7eb}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;background:#7c3aed;color:white}' +
    '</style></head><body>' +
    '<div class="title">Modal Diagnostics</div>' +
    '<span class="status-' + (results.status === 'OK' ? 'ok' : 'error') + '">' + escapeHtml(results.status) + '</span>' +
    '<p>' + escapeHtml(results.summary) + '</p>' +
    '<div class="card"><strong>Sheet Checks</strong><table>' +
    '<tr><th>Sheet</th><th>Status</th></tr>' +
    results.sheetChecks.map(function(c) {
      return '<tr><td>' + escapeHtml(c.name) + '</td><td>' + (c.found ? '✅' : '❌') + '</td></tr>';
    }).join('') +
    '</table></div>' +
    '<button class="btn" onclick="google.script.host.close()">Close</button>' +
    '</body></html>';

  var output = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, '🔍 Modal Diagnostics');
}

/**
 * Shows the diagnostics dialog
 * @returns {void}
 */
function showDiagnosticsDialog() {
  var results = DIAGNOSE_SETUP();

  var html = HtmlService.createHtmlOutput(getDiagnosticsDialogHtml_(results))
    .setWidth(550)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '🩺 System Diagnostics');
}

/**
 * Generates HTML for diagnostics dialog
 * @param {Object} results - Diagnostic results
 * @returns {string} HTML content
 * @private
 */
function getDiagnosticsDialogHtml_(results) {
  var statusClass = results.status === 'OK' ? 'ok' :
                    results.status === 'WARNING' ? 'warning' : 'error';

  return '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;margin:0}' +
    '.status{display:inline-block;padding:8px 16px;border-radius:20px;font-weight:600;margin-bottom:15px}' +
    '.ok{background:#d1fae5;color:#065f46}' +
    '.warning{background:#fef3c7;color:#92400e}' +
    '.error{background:#fee2e2;color:#991b1b}' +
    '.section{margin:15px 0}' +
    '.section-title{font-weight:600;margin-bottom:8px}' +
    '.list{background:#f9fafb;padding:10px;border-radius:6px;max-height:120px;overflow-y:auto}' +
    '.item{padding:3px 0;font-size:13px}' +
    '.btn{padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#7c3aed;color:white}' +
    '.btn-secondary{background:#e5e7eb;color:#374151;margin-right:10px}' +
    '</style></head><body>' +
    '<div class="status ' + statusClass + '">' + escapeHtml(results.status) + '</div>' +
    '<p>' + escapeHtml(results.summary) + '</p>' +
    '<div class="section">' +
    '<div class="section-title">Checks Performed (' + results.checks.length + ')</div>' +
    '<div class="list">' +
    results.checks.map(function(c) { return '<div class="item">✓ ' + escapeHtml(c) + '</div>'; }).join('') +
    '</div></div>' +
    (results.errors.length > 0 ?
      '<div class="section"><div class="section-title error">Errors (' + results.errors.length + ')</div>' +
      '<div class="list">' + results.errors.map(function(e) { return '<div class="item" style="color:#991b1b">❌ ' + escapeHtml(e) + '</div>'; }).join('') + '</div></div>' : '') +
    (results.warnings.length > 0 ?
      '<div class="section"><div class="section-title warning">Warnings (' + results.warnings.length + ')</div>' +
      '<div class="list">' + results.warnings.map(function(w) { return '<div class="item" style="color:#92400e">⚠ ' + escapeHtml(w) + '</div>'; }).join('') + '</div></div>' : '') +
    '<div style="margin-top:20px;text-align:right">' +
    '<button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>' +
    '<button class="btn btn-primary" onclick="runRepair()">Run Repair</button>' +
    '</div>' +
    '<script>function runRepair(){google.script.run.withFailureHandler(function(e){alert("Repair failed: "+e.message)}).REPAIR_DASHBOARD();google.script.host.close()}</script>' +
    '</body></html>';
}



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
    var cached = null;
    try {
      cached = memCache.get(key);
    } catch (_memErr) {
      // CacheService may be unavailable
    }
    if (cached) {
      try {
        if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT - Memory] ' + key);
        return JSON.parse(cached);
      } catch (_parseErr) {
        // Corrupted cache entry, remove it
        try { memCache.remove(key); } catch (_e) {}
      }
    }

    // Check properties cache (persistent)
    try {
      var propsCache = PropertiesService.getScriptProperties();
      var propsCached = propsCache.getProperty(key);
      if (propsCached) {
        var obj = JSON.parse(propsCached);
        if (obj.timestamp && (Date.now() - obj.timestamp) < (ttl * 1000)) {
          // Refresh memory cache from properties
          var str = JSON.stringify(obj.data);
          if (str.length < 100000) {
            try { memCache.put(key, str, Math.min(ttl, 21600)); } catch (_e) {}
          }
          if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE HIT - Props] ' + key);
          return obj.data;
        } else {
          // Expired - clean up stale property
          try { propsCache.deleteProperty(key); } catch (_e) {}
        }
      }
    } catch (_propsErr) {
      // Properties unavailable, continue to loader
    }

    // Cache miss - load fresh data
    if (CACHE_CONFIG.ENABLE_LOGGING) Logger.log('[CACHE MISS] ' + key);
    var data = loader();
    setCachedData(key, data, ttl);
    return data;

  } catch (e) {
    Logger.log('Cache error for ' + key + ': ' + e.message);
    // Fallback to direct load on error
    try {
      return loader();
    } catch (loaderErr) {
      Logger.log('Loader also failed for ' + key + ': ' + loaderErr.message);
      return null;
    }
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
    if (!str) return;

    // CacheService limit is 100KB per key
    if (str.length < 100000) {
      try {
        var memCache = CacheService.getScriptCache();
        memCache.put(key, str, Math.min(ttl, 21600));
      } catch (_memErr) {
        Logger.log('Memory cache write failed for ' + key + ': data too large or service unavailable');
      }
    }

    // PropertiesService limit is ~9KB per property
    if (str.length < 9000) {
      try {
        var propsCache = PropertiesService.getScriptProperties();
        propsCache.setProperty(key, JSON.stringify({
          data: data,
          timestamp: Date.now()
        }));
      } catch (_propsErr) {
        Logger.log('Properties cache write failed for ' + key);
      }
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
      return isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1]);
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
      } catch (_e) { /* cached value may not be valid JSON; skip age display */ }
    }

    // M-13: Apply escapeHtml() to dynamic values embedded in HTML
    return '<tr>' +
      '<td>' + escapeHtml(String(name)) + '</td>' +
      '<td>' + (inMem ? '✅' : '❌') + '</td>' +
      '<td>' + (inProps ? '✅' : '❌') + '</td>' +
      '<td>' + escapeHtml(String(age)) + '</td>' +
      '</tr>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
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
    'function warmUp(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Warm-up failed: "+e.message)}).warmUpCaches()}' +
    'function clearAll(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Clear failed: "+e.message)}).invalidateAllCaches()}' +
    '</script></body></html>'
  ).setWidth(600).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '🗄️ Cache Status');
}



/**
 * ============================================================================
 * UndoManager.gs - Undo/Redo System
 * ============================================================================
 *
 * This module handles the undo/redo functionality including:
 * - Action recording
 * - State management
 * - History navigation
 * - Snapshot management
 *
 * REFACTORED: Split from 06_Maintenance.gs for better maintainability
 *
 * @fileoverview Undo/redo system functions
 * @version 1.0.0
 * @requires 01_Constants.gs
 */

// ============================================================================
// UNDO CONFIGURATION
// ============================================================================

/**
 * Undo system configuration
 * @const {Object}
 */
var UNDO_CONFIG = {
  /** Maximum number of actions to store */
  MAX_HISTORY: 50,
  /** Storage key for undo history */
  STORAGE_KEY: 'undoRedoHistory'
};

// ============================================================================
// HISTORY MANAGEMENT
// ============================================================================

/**
 * Gets the undo history from storage
 * @returns {Object} History object with actions array and currentIndex
 */
function getUndoHistory() {
  var props = PropertiesService.getUserProperties();
  var json = props.getProperty(UNDO_CONFIG.STORAGE_KEY);

  if (json) {
    return JSON.parse(json);
  }

  return { actions: [], currentIndex: 0 };
}

/**
 * Saves the undo history to storage
 * @param {Object} history - History object to save
 * @returns {void}
 */
function saveUndoHistory(history) {
  var props = PropertiesService.getUserProperties();

  // Trim history if over limit
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }

  // UserProperties has a 9KB per-property limit; progressively trim if too large
  var json = JSON.stringify(history);
  while (json.length > 8000 && history.actions.length > 1) {
    history.actions.shift();
    history.currentIndex = Math.max(0, history.currentIndex - 1);
    json = JSON.stringify(history);
  }

  props.setProperty(UNDO_CONFIG.STORAGE_KEY, json);
}

/**
 * Clears all undo history
 * @returns {void}
 */
function clearUndoHistory() {
  PropertiesService.getUserProperties().deleteProperty(UNDO_CONFIG.STORAGE_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('History cleared', 'Undo/Redo', 3);
}

// ============================================================================
// ACTION RECORDING
// ============================================================================

/**
 * Records an action for undo/redo
 * @param {string} type - Action type (EDIT_CELL, ADD_ROW, DELETE_ROW, etc.)
 * @param {string} description - Human-readable description
 * @param {Object} beforeState - State before the action
 * @param {Object} afterState - State after the action
 * @returns {void}
 */
function recordAction(type, description, beforeState, afterState) {
  var history = getUndoHistory();

  // Clear any redo history when new action is recorded
  if (history.currentIndex < history.actions.length) {
    history.actions = history.actions.slice(0, history.currentIndex);
  }

  history.actions.push({
    type: type,
    description: description,
    timestamp: new Date().toISOString(),
    beforeState: beforeState,
    afterState: afterState
  });

  history.currentIndex = history.actions.length;
  saveUndoHistory(history);
}

/**
 * Records a cell edit action
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {*} oldValue - Value before edit
 * @param {*} newValue - Value after edit
 * @returns {void}
 */
function recordCellEdit(row, col, oldValue, newValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getActiveSheet();
  var colName = sheet.getRange(1, col).getValue();

  recordAction(
    'EDIT_CELL',
    'Edited ' + colName + ' in row ' + row,
    { row: row, col: col, value: oldValue, sheet: sheet.getName() },
    { row: row, col: col, value: newValue, sheet: sheet.getName() }
  );
}

/**
 * Records a row addition action
 * @param {number} row - Row number
 * @param {Array} rowData - Data in the new row
 * @returns {void}
 */
function recordRowAddition(row, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getActiveSheet();

  recordAction(
    'ADD_ROW',
    'Added row ' + row,
    null,
    { row: row, data: rowData, sheet: sheet.getName() }
  );
}

/**
 * Records a row deletion action
 * @param {number} row - Row number
 * @param {Array} rowData - Data that was in the row
 * @returns {void}
 */
function recordRowDeletion(row, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getActiveSheet();

  recordAction(
    'DELETE_ROW',
    'Deleted row ' + row,
    { row: row, data: rowData, sheet: sheet.getName() },
    null
  );
}

// ============================================================================
// UNDO/REDO OPERATIONS
// ============================================================================

/**
 * Undoes the last action
 * @returns {void}
 * @throws {Error} If nothing to undo
 */
function undoLastAction() {
  var history = getUndoHistory();

  if (history.currentIndex === 0) {
    throw new Error('Nothing to undo');
  }

  var action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);

  history.currentIndex--;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast('Undone: ' + action.description, 'Undo', 3);
}

/**
 * Redoes the last undone action
 * @returns {void}
 * @throws {Error} If nothing to redo
 */
function redoLastAction() {
  var history = getUndoHistory();

  if (history.currentIndex >= history.actions.length) {
    throw new Error('Nothing to redo');
  }

  var action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);

  history.currentIndex++;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast('Redone: ' + action.description, 'Redo', 3);
}

/**
 * Undoes all actions up to a specific index
 * @param {number} targetIndex - Target index to undo to
 * @returns {void}
 */
function undoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex > targetIndex) {
    undoLastAction();
    history = getUndoHistory();
  }
}

/**
 * Redoes all actions up to a specific index
 * @param {number} targetIndex - Target index to redo to
 * @returns {void}
 */
function redoToIndex(targetIndex) {
  var history = getUndoHistory();
  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) {
    redoLastAction();
    history = getUndoHistory();
  }
}

/**
 * Applies a saved state to the spreadsheet
 * @param {Object} state - State to apply
 * @param {string} actionType - Type of action being applied
 * @returns {void}
 */
function applyState(state, actionType) {
  if (!state) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(state.sheet);

  if (!sheet) {
    throw new Error('Sheet ' + state.sheet + ' not found');
  }

  switch (actionType) {
    case 'EDIT_CELL':
      sheet.getRange(state.row, state.col).setValue(state.value);
      break;

    case 'ADD_ROW':
      if (state.row && state.row > 1 && state.row <= sheet.getLastRow()) {
        sheet.deleteRow(state.row);
      }
      break;

    case 'DELETE_ROW':
      if (state.row && state.data) {
        sheet.insertRowAfter(state.row - 1);
        sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]);
      }
      break;

    case 'BATCH_UPDATE':
      if (state.changes && state.changes.length > 0) {
        var data = sheet.getDataRange().getValues();
        state.changes.forEach(function(c) {
          if (c.row > 0 && c.row <= data.length && c.col > 0) {
            data[c.row - 1][c.col - 1] = c.value;
          }
        });
        sheet.getDataRange().setValues(data);
      }
      break;
  }
}

// ============================================================================
// SNAPSHOT MANAGEMENT
// ============================================================================

/**
 * Creates a snapshot of the grievance log
 * @returns {Object} Snapshot object
 */
function createGrievanceSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    throw new Error('Grievance Log not found');
  }

  return {
    timestamp: new Date().toISOString(),
    data: sheet.getDataRange().getValues(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn()
  };
}

/**
 * Restores the grievance log from a snapshot
 * CR-14: Requires explicit confirmed=true parameter and creates a persistent backup first
 * @param {Object} snapshot - Snapshot to restore
 * @param {boolean} confirmed - Must be true to proceed; adds programmatic confirmation gate
 * @returns {void}
 */
function restoreFromSnapshot(snapshot, confirmed) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    throw new Error('Grievance Log not found');
  }

  // Validate snapshot column structure matches current sheet
  if (snapshot.data && snapshot.data.length > 0 && sheet.getLastColumn() > 0) {
    if (snapshot.data[0].length !== sheet.getLastColumn()) {
      throw new Error('Snapshot column count (' + snapshot.data[0].length +
        ') does not match current sheet (' + sheet.getLastColumn() + ')');
    }
  }

  // CR-14: Require explicit confirmed parameter as a programmatic safety gate
  if (confirmed !== true) {
    // Confirm with user via UI dialog
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Restore Snapshot',
      'This will replace all current grievance data with the snapshot from ' +
      snapshot.timestamp + '. A backup of current data will be created first.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  // CR-14: Create a persistent backup of current data to a hidden sheet before restoring
  var backupSheetName = '';
  try {
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    backupSheetName = '_PreRestore_Backup_' + timestamp;
    if (sheet.getLastRow() > 1) {
      var backupSheet = ss.insertSheet(backupSheetName);
      var currentData = sheet.getDataRange().getValues();
      backupSheet.getRange(1, 1, currentData.length, currentData[0].length).setValues(currentData);
      try {
        setSheetVeryHidden_(backupSheet);
      } catch (_hideErr) {
        backupSheet.hideSheet();
      }
      Logger.log('Pre-restore backup saved to hidden sheet: ' + backupSheetName);
    }
  } catch (_backupErr) {
    Logger.log('Warning: Could not create pre-restore backup sheet: ' + _backupErr.message);
    // Continue with restore — backup failure should not block the operation,
    // but log a warning for audit trail
  }

  // Clear existing data (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // Restore data
  if (snapshot.data.length > 1) {
    sheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length)
      .setValues(snapshot.data.slice(1));
  }

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR || 'SNAPSHOT_RESTORED', {
    snapshotTimestamp: snapshot.timestamp,
    backupSheet: backupSheetName || 'N/A',
    restoredBy: Session.getActiveUser().getEmail()
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Snapshot restored', 'Undo/Redo', 3);
}

// ============================================================================
// EXPORT AND UI
// ============================================================================

/**
 * Exports undo history to a new sheet
 * @returns {string} URL of the exported sheet
 */
function exportUndoHistoryToSheet() {
  var history = getUndoHistory();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // M-17: Use a constant for the export sheet name instead of hardcoded string
  var UNDO_EXPORT_SHEET_NAME = (typeof SHEETS !== 'undefined' && SHEETS.UNDO_HISTORY_EXPORT)
    ? SHEETS.UNDO_HISTORY_EXPORT
    : 'Undo_History_Export';

  var sheet = ss.getSheetByName(UNDO_EXPORT_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(UNDO_EXPORT_SHEET_NAME);
  }

  var headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#7c3aed')
    .setFontColor('#fff');

  if (history.actions.length > 0) {
    var rows = history.actions.map(function(a, i) {
      return [
        i + 1,
        a.type,
        a.description,
        new Date(a.timestamp).toLocaleString(),
        i < history.currentIndex ? 'Applied' : 'Undone'
      ];
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }

  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

/**
 * Shows the undo/redo panel
 * @returns {void}
 */
function showUndoRedoPanel() {
  var history = getUndoHistory();

  var rows = history.actions.length === 0
    ? '<tr><td colspan="5" style="text-align:center;padding:40px;color:#999">No actions recorded</td></tr>'
    : history.actions.slice().reverse().map(function(a, i) {
        var idx = history.actions.length - i;
        var time = new Date(a.timestamp).toLocaleString();
        var canUndo = idx <= history.currentIndex;

        return '<tr>' +
          '<td>' + idx + '</td>' +
          '<td><span class="badge ' + escapeHtml(a.type.toLowerCase()) + '">' + escapeHtml(a.type) + '</span></td>' +
          '<td>' + escapeHtml(a.description) + '</td>' +
          '<td style="font-size:12px;color:#666">' + escapeHtml(time) + '</td>' +
          '<td>' + (canUndo ? '<button onclick="undo(' + (idx - 1) + ')">↩️</button>' : '') + '</td>' +
          '</tr>';
      }).join('');

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' + getMobileOptimizedHead() +
    '<style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:12px}' +
    'h2{color:#7c3aed}' +
    '.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin:20px 0}' +
    '.stat{background:#f8f9fa;padding:15px;border-radius:8px;text-align:center;border-left:4px solid #7c3aed}' +
    '.num{font-size:32px;font-weight:bold;color:#7c3aed}' +
    'table{width:100%;border-collapse:collapse;margin:20px 0}' +
    'th{background:#7c3aed;color:white;padding:12px;text-align:left}' +
    'td{padding:12px;border-bottom:1px solid #e0e0e0}' +
    'button{background:#7c3aed;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;margin:2px}' +
    '.badge{display:inline-block;padding:4px 12px;border-radius:12px;font-size:11px;font-weight:bold}' +
    '.edit_cell{background:#e3f2fd;color:#1976d2}' +
    '.add_row{background:#e8f5e9;color:#388e3c}' +
    '.delete_row{background:#ffebee;color:#d32f2f}' +
    '.actions{margin:20px 0;padding:20px;background:#f8f9fa;border-radius:8px}' +
    '.btn-danger{background:#ef4444}' +
    '.btn-success{background:#10b981}' +
    '</style></head><body>' +
    '<div class="container">' +
    '<h2>↩️ Undo/Redo History</h2>' +
    '<div class="stats">' +
    '<div class="stat"><div class="num">' + history.actions.length + '</div><div>Total Actions</div></div>' +
    '<div class="stat"><div class="num">' + history.currentIndex + '</div><div>Current Position</div></div>' +
    '<div class="stat"><div class="num">' + (history.actions.length - history.currentIndex) + '</div><div>Can Redo</div></div>' +
    '</div>' +
    '<div class="actions">' +
    '<button onclick="performUndo()">↩️ Undo Last</button>' +
    '<button onclick="performRedo()">↪️ Redo</button>' +
    '<button class="btn-danger" onclick="clearHistory()">🗑️ Clear All</button>' +
    '<button class="btn-success" onclick="exportHistory()">📥 Export</button>' +
    '</div>' +
    '<table>' +
    '<tr><th>#</th><th>Type</th><th>Description</th><th>Time</th><th></th></tr>' +
    rows +
    '</table>' +
    '</div>' +
    '<script>' +
    'function performUndo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Error: "+e.message)}).undoLastAction()}' +
    'function performRedo(){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Error: "+e.message)}).redoLastAction()}' +
    'function undo(i){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Undo failed: "+e.message)}).undoToIndex(i)}' +
    'function clearHistory(){if(confirm("Clear all history?")){google.script.run.withSuccessHandler(function(){location.reload()}).withFailureHandler(function(e){alert("Clear failed: "+e.message)}).clearUndoHistory()}}' +
    'function exportHistory(){google.script.run.withSuccessHandler(function(url){alert("Exported!");if(/^https:\\/\\/docs\\.google\\.com\\//.test(url))window.open(url,"_blank");else alert("Invalid URL")}).withFailureHandler(function(e){alert("Export failed: "+e.message)}).exportUndoHistoryToSheet()}' +
    '</script></body></html>'
  ).setWidth(800).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '↩️ Undo/Redo History');
}



/**
 * ============================================================================
 * Maintenance.gs - Admin & Diagnostic Tools
 * ============================================================================
 *
 * This module contains administrative and diagnostic functions including:
 * - System diagnostics (DIAGNOSE_SETUP)
 * - Dashboard repair (REPAIR_DASHBOARD)
 * - Hidden sheet verification
 * - Data quality fixes
 * - Audit logging
 *
 * IMPORTANT: These are "heavy" operations that should only be run by
 * administrators. Moving them to a separate file prevents regular stewards
 * from accidentally triggering high-intensity system scans.
 *
 * @fileoverview Administrative and maintenance utilities
 * @version 2.0.0
 * @requires Constants.gs
 */

// Note: DIAGNOSE_SETUP(), REPAIR_DASHBOARD(), removeDeprecatedTabs(), and
// showRepairDialog() implementations are defined earlier in this file.
// verifyHiddenSheets() and fixDataQualityIssues() are in HiddenSheets.gs
// and DataIntegrity.gs respectively.

// ============================================================================
// AUDIT LOGGING
// ============================================================================

/**
 * Logs an audit event to the Audit Log sheet.
 * v4.8.1: Adds an Integrity Hash column (F) that forms a hash chain —
 * each row's hash includes the previous row's hash, so any tampering
 * (edit, insert, or delete) breaks the chain and is detectable via
 * verifyAuditLogIntegrity().
 *
 * @param {string} eventType - The type of event from AUDIT_EVENTS
 * @param {Object} details - Event details object
 */
function logAuditEvent(eventType, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

    // Create audit sheet if it doesn't exist — use the canonical 10-col schema
    // H-43: Standardize to the 10-col AUDIT_LOG_HEADER_MAP_ schema used by
    // 08d_AuditAndFormulas.gs, plus an Integrity Hash column (col 11).
    if (!auditSheet) {
      auditSheet = ss.insertSheet(SHEET_NAMES.AUDIT_LOG);
      var headerRow = (typeof getHeadersFromMap_ === 'function' && typeof AUDIT_LOG_HEADER_MAP_ !== 'undefined')
        ? getHeadersFromMap_(AUDIT_LOG_HEADER_MAP_)
        : ['Timestamp', 'User Email', 'Sheet', 'Row', 'Column', 'Field Name', 'Old Value', 'New Value', 'Record ID', 'Action Type'];
      headerRow.push('Integrity Hash');
      auditSheet.appendRow(headerRow);
      auditSheet.setFrozenRows(1);
    } else {
      // Ensure Integrity Hash header exists (upgrade path for existing sheets)
      var headers = auditSheet.getRange(1, 1, 1, auditSheet.getLastColumn()).getValues()[0];
      var hasIntegrityCol = false;
      for (var h = 0; h < headers.length; h++) {
        if (String(headers[h]).trim() === 'Integrity Hash') {
          hasIntegrityCol = true;
          break;
        }
      }
      if (!hasIntegrityCol) {
        var nextCol = auditSheet.getLastColumn() + 1;
        auditSheet.getRange(1, nextCol).setValue('Integrity Hash');
      }
    }

    // Build log entry
    const timestamp = new Date();
    const rawEmail = Session.getActiveUser().getEmail() || 'Unknown';
    const user = (typeof maskEmail === 'function') ? maskEmail(rawEmail) : rawEmail;
    const detailsJson = JSON.stringify(details);
    const sessionId = Session.getTemporaryActiveUserKey() || '';

    // Compute integrity hash chained to previous row
    var previousHash = '';
    var lastRow = auditSheet.getLastRow();
    if (lastRow >= 2) {
      // Find Integrity Hash column index dynamically
      var hdrRow = auditSheet.getRange(1, 1, 1, auditSheet.getLastColumn()).getValues()[0];
      var integrityColIdx = -1;
      for (var i = 0; i < hdrRow.length; i++) {
        if (String(hdrRow[i]).trim() === 'Integrity Hash') {
          integrityColIdx = i + 1; // 1-indexed
          break;
        }
      }
      if (integrityColIdx > 0) {
        previousHash = String(auditSheet.getRange(lastRow, integrityColIdx).getValue() || '');
      }
    }

    var integrityHash = '';
    if (typeof computeAuditRowHash_ === 'function') {
      // Use rawEmail (not masked) in integrity hash so hash chain remains verifiable
      integrityHash = computeAuditRowHash_(previousHash, timestamp, eventType, rawEmail, detailsJson, sessionId);
    }

    // H-43: Write in 10-col schema: Timestamp, User Email, Sheet, Row, Column,
    // Field Name, Old Value, New Value, Record ID, Action Type, Integrity Hash.
    // For event-level logging, Sheet/Row/Column/Field Name are empty; details go
    // in New Value, sessionId in Record ID, eventType in Action Type.
    auditSheet.appendRow([
      timestamp,           // Timestamp
      user,                // User Email
      '',                  // Sheet (N/A for events)
      '',                  // Row (N/A for events)
      '',                  // Column (N/A for events)
      '',                  // Field Name (N/A for events)
      '',                  // Old Value (N/A for events)
      detailsJson,         // New Value (event details JSON)
      sessionId,           // Record ID (session key)
      eventType,           // Action Type (event type)
      integrityHash        // Integrity Hash
    ]);

    // Trim old entries if sheet gets too large (keep last 10,000)
    // CR-30 WARNING: Deleting rows breaks the integrity hash chain because each
    // row's hash is chained to the previous row. Before trimming, save the hash
    // of the last row being deleted as a "chain checkpoint" in ScriptProperties
    // so that verifyAuditLogIntegrity() can start verification from this checkpoint.
    const rowCount = auditSheet.getLastRow();
    if (rowCount > 10000) {
      var rowsToDelete = rowCount - 10000;
      // Save the hash of the last deleted row as a chain checkpoint
      try {
        var checkpointRow = 1 + rowsToDelete; // last row being deleted (1-indexed, row 1 is header)
        var trimHeaders = auditSheet.getRange(1, 1, 1, auditSheet.getLastColumn()).getValues()[0];
        var hashColIdx = -1;
        for (var ci = 0; ci < trimHeaders.length; ci++) {
          if (String(trimHeaders[ci]).trim() === 'Integrity Hash') {
            hashColIdx = ci + 1;
            break;
          }
        }
        if (hashColIdx > 0) {
          var checkpointHash = String(auditSheet.getRange(checkpointRow, hashColIdx).getValue() || '');
          PropertiesService.getScriptProperties().setProperty(
            'AUDIT_CHAIN_CHECKPOINT_HASH', checkpointHash
          );
          PropertiesService.getScriptProperties().setProperty(
            'AUDIT_CHAIN_CHECKPOINT_TIMESTAMP', new Date().toISOString()
          );
        }
      } catch (cpErr) {
        Logger.log('Warning: Could not save audit chain checkpoint: ' + cpErr.message);
      }
      auditSheet.deleteRows(2, rowsToDelete);
    }

  } catch (error) {
    console.error('Error logging audit event:', error);
    // Don't throw - audit logging shouldn't break main functionality
  }
}

/**
 * Gets recent audit log entries
 * @param {number} count - Number of entries to retrieve
 * @return {Array} Array of audit log entries
 */
function getRecentAuditLogs(count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);

  if (!auditSheet) return [];

  const lastRow = auditSheet.getLastRow();
  const startRow = Math.max(2, lastRow - count + 1);
  const numRows = lastRow - startRow + 1;

  if (numRows <= 0) return [];

  // Read all columns including Integrity Hash
  var numCols = EVENT_AUDIT_COLS.INTEGRITY_HASH;
  const data = auditSheet.getRange(startRow, 1, numRows, numCols).getValues();

  return data.map(row => ({
    timestamp: row[EVENT_AUDIT_COLS.TIMESTAMP - 1],
    eventType: row[EVENT_AUDIT_COLS.EVENT_TYPE - 1],
    user: row[EVENT_AUDIT_COLS.USER - 1],
    details: row[EVENT_AUDIT_COLS.DETAILS - 1],
    sessionId: row[EVENT_AUDIT_COLS.SESSION_ID - 1],
    integrityHash: row[EVENT_AUDIT_COLS.INTEGRITY_HASH - 1] || ''
  })).reverse();
}

// ============================================================================
// NUCLEAR OPTIONS (ADMIN ONLY)
// ============================================================================

/**
 * WARNING: Completely resets all hidden sheets
 * This is a destructive operation that should only be used as a last resort
 * @return {Object} Reset results
 */
function NUCLEAR_RESET_HIDDEN_SHEETS() {
  var isDev = typeof IS_DEV_ENVIRONMENT !== 'undefined' && IS_DEV_ENVIRONMENT === true;
  if (!isDev) {
    var _ui = SpreadsheetApp.getUi();
    var _response = _ui.alert('Production Safety Check',
      'This action is intended for development environments only. Are you SURE you want to proceed?',
      _ui.ButtonSet.YES_NO);
    if (_response !== _ui.Button.YES) return;
  }
  // Require double confirmation
  const ui = SpreadsheetApp.getUi();
  const response1 = ui.alert(
    '⚠️ NUCLEAR RESET',
    'This will DELETE and REBUILD all hidden calculation sheets. ' +
    'This may temporarily break formulas. Are you absolutely sure?',
    ui.ButtonSet.YES_NO
  );

  if (response1 !== ui.Button.YES) {
    return errorResponse('Cancelled by user');
  }

  const response2 = ui.alert(
    '⚠️ FINAL WARNING',
    'This action CANNOT be undone. Click YES to proceed with the nuclear reset.',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    return errorResponse('Cancelled by user');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete all hidden sheets
  for (const sheetName of Object.values(HIDDEN_SHEETS)) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  // Rebuild all hidden sheets
  const result = setupAllHiddenSheets();

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'NUCLEAR_RESET',
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    message: `Nuclear reset complete. Created ${result.created} sheets.`
  };
}

/**
 * WARNING: Wipes all grievance data
 * For testing/development use only
 */
function NUCLEAR_WIPE_GRIEVANCES() {
  // H-23: Authorization check — only admins may perform destructive data wipe
  if (typeof getUserRole_ === 'function') {
    var callerEmail = Session.getActiveUser().getEmail();
    var callerRole = getUserRole_(callerEmail);
    if (callerRole !== 'admin') {
      SpreadsheetApp.getUi().alert('Access Denied',
        'Only administrators may perform this destructive action.', SpreadsheetApp.getUi().ButtonSet.OK);
      return errorResponse('Unauthorized: admin role required');
    }
  }

  var isDev = typeof IS_DEV_ENVIRONMENT !== 'undefined' && IS_DEV_ENVIRONMENT === true;
  if (!isDev) {
    var _ui = SpreadsheetApp.getUi();
    var _response = _ui.alert('Production Safety Check',
      'This action is intended for development environments only. Are you SURE you want to proceed?',
      _ui.ButtonSet.YES_NO);
    if (_response !== _ui.Button.YES) return;
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ DANGER - DATA WIPE',
    'This will PERMANENTLY DELETE all grievance records. ' +
    'This action CANNOT be undone. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return errorResponse('Cancelled');
  }

  // Second confirmation required
  const response2 = ui.alert(
    '⚠️ FINAL WARNING',
    'All grievance data will be PERMANENTLY deleted. Are you absolutely certain?',
    ui.ButtonSet.YES_NO
  );
  if (response2 !== ui.Button.YES) {
    return errorResponse('Cancelled');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Try canonical SHEETS.GRIEVANCE_LOG first, fall back to SHEET_NAMES.GRIEVANCE_TRACKER
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG) || ss.getSheetByName(SHEET_NAMES.GRIEVANCE_TRACKER);

  if (!sheet) {
    return errorResponse('Grievance Log/Tracker not found');
  }

  const lastRow = sheet.getLastRow();

  // Safety snapshot: back up data to a hidden sheet before destructive wipe
  if (lastRow > 1) {
    var snapshotName = '_PreWipe_Snapshot_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var snapshotSheet = ss.insertSheet(snapshotName);
    var allData = sheet.getDataRange().getValues();
    snapshotSheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
    try {
      setSheetVeryHidden_(snapshotSheet);
    } catch (_hideErr) {
      snapshotSheet.hideSheet();
    }
    Logger.log('Pre-wipe snapshot saved to hidden sheet: ' + snapshotName);
  }

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  logAuditEvent(AUDIT_EVENTS.SYSTEM_REPAIR, {
    action: 'NUCLEAR_WIPE_GRIEVANCES',
    rowsDeleted: lastRow - 1,
    snapshotSheet: lastRow > 1 ? snapshotName : 'N/A (no data)',
    performedBy: Session.getActiveUser().getEmail()
  });

  return {
    success: true,
    message: `Deleted ${lastRow - 1} grievance records`
  };
}

// ============================================================================
// BACKUP & SNAPSHOT FUNCTIONS (Strategic Command Center)
// ============================================================================

/**
 * Creates a complete backup snapshot of the spreadsheet
 * Copies entire spreadsheet to the archive folder with timestamp
 */
function createWeeklySnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Get archive folder ID from Config
  var archiveFolderId = '';
  try {
    archiveFolderId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID) || COMMAND_CONFIG.ARCHIVE_FOLDER_ID;
  } catch (_e) {
    // Fall back to a default if config is not available
  }

  var folder;
  if (!archiveFolderId) {
    // Auto-create archive folder if not configured
    try {
      var folders = DriveApp.getFoldersByName('Dashboard Archive');
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder('Dashboard Archive');
      }
      archiveFolderId = folder.getId();
      // Save to Config sheet for future use
      try {
        var configSheet = ss.getSheetByName(SHEETS.CONFIG);
        if (configSheet) {
          configSheet.getRange(3, CONFIG_COLS.ARCHIVE_FOLDER_ID).setValue(archiveFolderId);
        }
      } catch (_saveErr) {
        Logger.log('Could not save archive folder ID to config: ' + _saveErr.message);
      }
      ss.toast('Auto-created archive folder in Google Drive.', 'Archive Setup', 3);
    } catch (createErr) {
      ui.alert('Archive Folder Error',
        'Could not auto-create archive folder: ' + createErr.message + '\n\n' +
        'You can manually set the Archive Folder ID in the Config sheet (column ' +
        (typeof getColumnLetter === 'function' && typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.ARCHIVE_FOLDER_ID
          ? getColumnLetter(CONFIG_COLS.ARCHIVE_FOLDER_ID) : 'ARCHIVE_FOLDER_ID') + ').',
        ui.ButtonSet.OK);
      return;
    }
  }

  try {
    if (!folder) {
      folder = DriveApp.getFolderById(archiveFolderId);
    }
    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    var snapshotName = 'SNAPSHOT_' + date;

    // Create the copy
    DriveApp.getFileById(ss.getId()).makeCopy(snapshotName, folder);

    // Log the backup
    logAuditEvent('SNAPSHOT_CREATED', {
      snapshotName: snapshotName,
      folderId: archiveFolderId,
      createdBy: Session.getActiveUser().getEmail()
    });

    ui.alert('Snapshot Created', 'Backup saved to archive folder.\n\nFile: ' + snapshotName, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', 'Failed to create snapshot: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Creates an automated snapshot (for scheduled triggers)
 * Does not show UI alerts - just logs the action
 */
function createAutomatedSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var archiveFolderId = '';
  try {
    archiveFolderId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID) || COMMAND_CONFIG.ARCHIVE_FOLDER_ID;
  } catch (e) {
    Logger.log('Could not get archive folder ID: ' + e.message);
    return errorResponse('Archive folder not configured');
  }

  if (!archiveFolderId) {
    Logger.log('Archive folder not configured');
    return errorResponse('Archive folder not configured');
  }

  try {
    var folder = DriveApp.getFolderById(archiveFolderId);
    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var snapshotName = 'AUTO_SNAPSHOT_' + date;

    DriveApp.getFileById(ss.getId()).makeCopy(snapshotName, folder);

    logAuditEvent('AUTO_SNAPSHOT_CREATED', {
      snapshotName: snapshotName,
      createdBy: 'Automated Trigger'
    });

    return { success: true, snapshotName: snapshotName };

  } catch (e) {
    Logger.log('Failed to create automated snapshot: ' + e.message);
    return errorResponse(e.message);
  }
}

/**
 * Sets up weekly snapshot trigger
 * Creates a time-based trigger to run every Sunday at 2am
 */
function setupWeeklySnapshotTrigger() {
  // Remove existing snapshot triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'createAutomatedSnapshot') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new weekly trigger (Sunday at 2am)
  ScriptApp.newTrigger('createAutomatedSnapshot')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();

  SpreadsheetApp.getUi().alert(
    'Weekly Snapshot Enabled',
    'Automated backups will be created every Sunday at 2:00 AM.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Sets up weekly Drive folder cleanup trigger.
 * Runs every Saturday at 3am to remove empty sub-folders for resolved grievances.
 */
function setupWeeklyDriveCleanupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'cleanupEmptyDriveFolders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('cleanupEmptyDriveFolders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(3)
    .create();

  SpreadsheetApp.getUi().alert(
    'Drive Cleanup Enabled',
    'Empty Drive folders for resolved grievances will be cleaned up every Saturday at 3:00 AM.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Cleans up old export files from Drive.
 * Called by a weekly trigger (set up via setupWeeklyExportCleanupTrigger).
 * Searches the user's Drive root for CSV/export files older than the retention period
 * and moves them to trash.
 */
function cleanupOldExportFiles() {
  var RETENTION_DAYS = 7;
  var cutoff = new Date(Date.now() - RETENTION_DAYS * 86400000);
  var trashed = 0;

  // Export file name patterns used across the codebase
  var patterns = [
    'title contains "grievance_export_" and mimeType = "text/csv"',
    'title contains "MemberDirectory_Export_" and mimeType = "text/csv"',
    'title contains "AUDIT_LOG_BACKUP_" and mimeType = "text/csv"',
    'title = "MemberDirectory.csv" and mimeType = "text/csv"'
  ];

  for (var p = 0; p < patterns.length; p++) {
    var query = patterns[p] + ' and trashed = false';
    try {
      var files = DriveApp.searchFiles(query);
      while (files.hasNext()) {
        var file = files.next();
        if (file.getDateCreated() < cutoff) {
          file.setTrashed(true);
          trashed++;
        }
      }
    } catch (_e) {
      Logger.log('Export cleanup search error (' + patterns[p] + '): ' + _e.message);
    }
  }

  if (trashed > 0) {
    logAuditEvent('EXPORT_FILES_CLEANED', 'Trashed ' + trashed + ' export file(s) older than ' + RETENTION_DAYS + ' days');
  }
  Logger.log('cleanupOldExportFiles: trashed ' + trashed + ' file(s)');
}

/**
 * Sets up weekly export file cleanup trigger.
 * Runs every Sunday at 4am to remove old CSV exports from Drive.
 */
function setupWeeklyExportCleanupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'cleanupOldExportFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('cleanupOldExportFiles')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(4)
    .create();

  SpreadsheetApp.getUi().alert(
    'Export Cleanup Enabled',
    'Old export files will be cleaned up every Sunday at 4:00 AM.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Navigates to the Audit Log sheet
 */
function navigateToAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!auditSheet) {
    // Create audit log if it doesn't exist
    auditSheet = ss.insertSheet(SHEETS.AUDIT_LOG);
    auditSheet.getRange('A1:E1').setValues([['Timestamp', 'Event Type', 'User', 'Details', 'Session ID']]);
    auditSheet.getRange('A1:E1')
      .setFontWeight('bold')
      .setBackground(COLORS.CARD_DARK_BG)
      .setFontColor(COLORS.CARD_DARK_TEXT);
    setSheetVeryHidden_(auditSheet);
  }

  // Unhide temporarily and activate
  if (auditSheet.isSheetHidden()) {
    setSheetVisible_(auditSheet);
  }
  auditSheet.activate();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Showing Audit Log. It will be hidden again when you switch sheets.',
    COMMAND_CONFIG.SYSTEM_NAME,
    5
  );
}

// ============================================================================
// SETTINGS MANAGEMENT
// ============================================================================

/**
 * Shows the settings dialog
 */
function showSettingsDialog() {
  const currentSettings = getSettings();

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .settings-container { padding: 20px; }
        .setting-group { margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid #eee; }
        .setting-title { font-weight: 600; margin-bottom: 8px; }
        .setting-desc { font-size: 12px; color: #666; margin-bottom: 10px; }
        .danger-zone { background: #fff5f5; padding: 15px; border-radius: 8px; border: 1px solid #feb2b2; margin-top: 20px; }
        .danger-title { color: #c53030; font-weight: 600; margin-bottom: 10px; }
      </style>
    </head>
    <body>
      <div class="settings-container">
        <div class="setting-group">
          <div class="setting-title">Calendar Integration</div>
          <div class="setting-desc">Automatically sync grievance deadlines to Google Calendar</div>
          <label>
            <input type="checkbox" id="autoSyncCalendar"
                   ${currentSettings.autoSyncCalendar ? 'checked' : ''}>
            Enable auto-sync on grievance changes
          </label>
        </div>

        <div class="setting-group">
          <div class="setting-title">Email Notifications</div>
          <div class="setting-desc">Send email reminders for upcoming deadlines</div>
          <label>
            <input type="checkbox" id="emailReminders"
                   ${currentSettings.emailReminders ? 'checked' : ''}>
            Enable deadline reminders
          </label>
          <div style="margin-top: 8px;">
            <label>Days before deadline:
              <select id="reminderDays">
                <option value="3" ${currentSettings.reminderDays === 3 ? 'selected' : ''}>3 days</option>
                <option value="5" ${currentSettings.reminderDays === 5 ? 'selected' : ''}>5 days</option>
                <option value="7" ${currentSettings.reminderDays === 7 ? 'selected' : ''}>7 days</option>
              </select>
            </label>
          </div>
        </div>

        <div class="setting-group">
          <div class="setting-title">Drive Integration</div>
          <div class="setting-desc">Automatically create Drive folders for new grievances</div>
          <label>
            <input type="checkbox" id="autoCreateFolders"
                   ${currentSettings.autoCreateFolders ? 'checked' : ''}>
            Auto-create folders
          </label>
        </div>

        <div class="danger-zone">
          <div class="danger-title">⚠️ Danger Zone</div>
          <p style="font-size: 13px; margin-bottom: 10px;">
            Administrative functions that should be used with caution
          </p>
          <button class="btn btn-secondary" onclick="runDiagnostics()">
            Run Diagnostics
          </button>
          <button class="btn btn-secondary" onclick="repairDashboard()">
            Repair Dashboard
          </button>
          <button class="btn btn-danger" onclick="nuclearReset()">
            Nuclear Reset
          </button>
        </div>

        <div style="margin-top: 20px; text-align: right;">
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
          <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
        </div>
      </div>

      <script>
        function saveSettings() {
          const settings = {
            autoSyncCalendar: document.getElementById('autoSyncCalendar').checked,
            emailReminders: document.getElementById('emailReminders').checked,
            reminderDays: parseInt(document.getElementById('reminderDays').value),
            autoCreateFolders: document.getElementById('autoCreateFolders').checked
          };

          google.script.run
            .withSuccessHandler(function() {
              alert('Settings saved!');
              google.script.host.close();
            })
            .withFailureHandler(function(e){alert("Save failed: "+e.message)})
            .saveSettings(settings);
        }

        function runDiagnostics() {
          google.script.run.withFailureHandler(function(e){alert(e.message)}).showDiagnosticsDialog();
          google.script.host.close();
        }

        function repairDashboard() {
          google.script.run.withFailureHandler(function(e){alert(e.message)}).showRepairDialog();
          google.script.host.close();
        }

        function nuclearReset() {
          if (confirm('This is an extreme action. Are you sure?')) {
            google.script.run.withFailureHandler(function(e){alert(e.message)}).NUCLEAR_RESET_HIDDEN_SHEETS();
            google.script.host.close();
          }
        }
      </script>
    </body>
    </html>
  `).setWidth(500).setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Settings');
}

/**
 * Gets current settings from document properties
 * @return {Object} Current settings
 */
function getSettings() {
  const props = PropertiesService.getDocumentProperties();
  return {
    autoSyncCalendar: props.getProperty('autoSyncCalendar') === 'true',
    emailReminders: props.getProperty('emailReminders') === 'true',
    reminderDays: parseInt(props.getProperty('reminderDays') || '7'),
    autoCreateFolders: props.getProperty('autoCreateFolders') === 'true'
  };
}

/**
 * Saves settings to document properties
 * @param {Object} settings - Settings to save
 */
function saveSettings(settings) {
  // M-58: Whitelist of allowed setting keys to prevent arbitrary property writes
  var ALLOWED_SETTINGS = {
    autoSyncCalendar: 'boolean',
    emailReminders: 'boolean',
    reminderDays: 'number',
    autoCreateFolders: 'boolean'
  };

  if (!settings || typeof settings !== 'object') {
    throw new Error('Invalid settings object');
  }

  const props = PropertiesService.getDocumentProperties();

  // Only write whitelisted keys, validate types
  var savedKeys = [];
  for (var key in ALLOWED_SETTINGS) {
    if (ALLOWED_SETTINGS.hasOwnProperty(key) && settings.hasOwnProperty(key)) {
      var expectedType = ALLOWED_SETTINGS[key];
      var value = settings[key];
      // Coerce to expected type for safety
      if (expectedType === 'boolean') {
        props.setProperty(key, String(!!value));
      } else if (expectedType === 'number') {
        var numVal = parseInt(value, 10);
        if (isNaN(numVal) || numVal < 0) numVal = 7; // safe default for reminderDays
        props.setProperty(key, String(numVal));
      } else {
        props.setProperty(key, String(value));
      }
      savedKeys.push(key);
    }
  }

  logAuditEvent(AUDIT_EVENTS.SETTINGS_CHANGED, {
    settingsKeys: savedKeys,
    changedBy: Session.getActiveUser().getEmail()
  });
}
/**
 * ============================================================================
 * PERFORMANCE CACHING & UNDO/REDO SYSTEM
 * ============================================================================
 * Data caching for performance + action history
 */

// ==================== CACHE CONFIGURATION ====================

// NOTE: CACHE_CONFIG, CACHE_KEYS, and UNDO_CONFIG are already declared above (lines 536-887).
// Duplicate declarations removed to prevent silent overwrites.

/**
 * Dashboard - Data Integrity and Performance Enhancements
 *
 * This module contains improvements for:
 * - Batch operations for optimized data writing
 * - Comprehensive error handling with retry logic
 * - Confirmation dialogs for destructive actions
 * - Dynamic validation ranges
 * - Duplicate ID validation
 * - Ghost validation for orphaned grievances
 * - Steward load balancing metrics
 * - Self-healing Config tool
 * - Enhanced audit logging
 * - Auto-archive for closed grievances
 *
 * @version 2.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// BATCH OPERATIONS UTILITIES
// ============================================================================

/**
 * Batch write utility - writes multiple values to a sheet in a single operation
 * Significantly faster than individual setValue() calls
 * @param {Sheet} sheet - Target sheet
 * @param {Array<Object>} updates - Array of {row, col, value} objects
 */
function batchSetValues(sheet, updates) {
  if (!updates || updates.length === 0) return;

  // H-4: Guard against updates targeting the header row (row <= 1)
  updates = updates.filter(function(u) {
    if (u.row <= 1) { Logger.log('batchSetValues: skipped update targeting header row=' + u.row); return false; }
    return true;
  });
  if (updates.length === 0) return;

  // CR-13: Wrap entire read-modify-write cycle in a script lock to prevent
  // concurrent writes from overwriting each other's changes
  withScriptLock_(function() {
    // Group updates by row for efficient writing
    var rowGroups = {};
    updates.forEach(function(update) {
      if (!rowGroups[update.row]) {
        rowGroups[update.row] = [];
      }
      rowGroups[update.row].push(update);
    });

    // Get current sheet data for merging
    var lastRow = Math.max.apply(null, updates.map(function(u) { return u.row; }));
    var lastCol = Math.max.apply(null, updates.map(function(u) { return u.col; }));

    // Read current data
    var currentData = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Apply updates to the array
    updates.forEach(function(update) {
      currentData[update.row - 1][update.col - 1] = update.value;
    });

    // Write back in single operation
    sheet.getRange(1, 1, lastRow, lastCol).setValues(currentData);
  });
}

/**
 * Batch write utility for a specific row - updates multiple columns in one operation
 * @param {Sheet} sheet - Target sheet
 * @param {number} row - Row number (1-indexed)
 * @param {Object} columnValues - Object mapping column numbers to values {col: value}
 */
function batchSetRowValues(sheet, row, columnValues) {
  var cols = Object.keys(columnValues);
  if (cols.length === 0) return;

  // Find range bounds
  var minCol = Math.min.apply(null, cols.map(Number));
  var maxCol = Math.max.apply(null, cols.map(Number));
  var numCols = maxCol - minCol + 1;

  // Read current row data
  var rowData = sheet.getRange(row, minCol, 1, numCols).getValues()[0];

  // Apply updates
  cols.forEach(function(col) {
    var colNum = parseInt(col);
    rowData[colNum - minCol] = columnValues[col];
  });

  // Write back in single operation
  sheet.getRange(row, minCol, 1, numCols).setValues([rowData]);
}

/**
 * Batch append rows - adds multiple rows in a single operation
 * Much faster than multiple appendRow() calls
 * @param {Sheet} sheet - Target sheet
 * @param {Array<Array>} rows - 2D array of row data
 */
function batchAppendRows(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var lastRow = sheet.getLastRow();
  var numCols = rows[0].length;

  // Write all rows in a single operation
  sheet.getRange(lastRow + 1, 1, rows.length, numCols).setValues(rows);
}

// ============================================================================
// ERROR HANDLING WITH RETRY LOGIC
// ============================================================================

/**
 * Execute a function with retry logic for transient failures
 * @param {Function} fn - Function to execute
 * @param {Object} options - Options: {maxRetries, baseDelay, onError}
 * @returns {*} Result of the function
 */
function executeWithRetry(fn, options) {
  options = options || {};
  var maxRetries = options.maxRetries || 3;
  var baseDelay = options.baseDelay || 1000; // 1 second
  var onError = options.onError || function() {};

  var lastError;

  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      // Log the error
      Logger.log('Attempt ' + (attempt + 1) + ' failed: ' + error.message);
      onError(error, attempt);

      // Check if we should retry
      if (attempt < maxRetries) {
        // Exponential backoff: 1s, 2s, 4s, 8s...
        var delay = baseDelay * Math.pow(2, attempt);
        Utilities.sleep(delay);
      }
    }
  }

  // All retries exhausted
  throw new Error('Operation failed after ' + (maxRetries + 1) + ' attempts: ' + lastError.message);
}

/**
 * Safe sheet operation wrapper with error handling
 * @param {Function} operation - Sheet operation to perform
 * @param {string} operationName - Name for logging
 * @returns {Object} {success: boolean, result: *, error: string}
 */
function safeSheetOperation(operation, operationName) {
  try {
    var result = executeWithRetry(operation, {
      maxRetries: 2,
      baseDelay: 500,
      onError: function(error, attempt) {
        Logger.log(operationName + ' - Retry ' + (attempt + 1) + ': ' + error.message);
      }
    });
    return { success: true, result: result, error: null };
  } catch (error) {
    Logger.log(operationName + ' FAILED: ' + error.message);
    return { success: false, result: null, error: error.message };
  }
}

// ============================================================================
// CONFIRMATION DIALOGS FOR DESTRUCTIVE ACTIONS
// ============================================================================

/**
 * Safe version of getOrCreateSheet that confirms before deleting existing data
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} name - Sheet name
 * @param {boolean} forceDelete - Skip confirmation if true
 * @returns {Sheet|null} The created sheet, or null if user cancelled
 */
function getOrCreateSheetSafe(ss, name, forceDelete) {
  var sheet = ss.getSheetByName(name);

  if (sheet) {
    // Check if sheet has data
    var hasData = sheet.getLastRow() > 1 || sheet.getLastColumn() > 1;

    if (hasData && !forceDelete) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        '⚠️ Confirm Data Deletion',
        'The sheet "' + name + '" already exists and contains data.\n\n' +
        'Deleting this sheet will permanently remove all data in it.\n\n' +
        'Are you sure you want to continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        Logger.log('User cancelled deletion of sheet: ' + name);
        return null;
      }
    }

    ss.deleteSheet(sheet);
    Logger.log('Deleted existing sheet: ' + name);
  }

  return ss.insertSheet(name);
}

/**
 * Confirm before performing any destructive operation
 * @param {string} actionName - Description of the action
 * @param {string} warningMessage - Detailed warning message
 * @returns {boolean} True if user confirmed, false otherwise
 */
function confirmDestructiveAction(actionName, warningMessage) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ ' + actionName,
    warningMessage + '\n\nThis action cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}

// ============================================================================
// DYNAMIC VALIDATION RANGES
// ============================================================================

/**
 * Set dropdown validation using dynamic row count instead of fixed 100
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */

/**
 * Set multi-select validation using dynamic row count
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */

// ============================================================================
// DUPLICATE MEMBER ID VALIDATION
// ============================================================================

/**
 * Check if a Member ID already exists in the Member Directory
 * @param {string} memberId - Member ID to check
 * @returns {Object} {exists: boolean, row: number|null}
 */
function checkDuplicateMemberId(memberId) {
  if (!memberId) return { exists: false, row: null };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return { exists: false, row: null };

  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return { exists: false, row: null };

  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1).getValues();

  for (var i = 0; i < memberIds.length; i++) {
    if (memberIds[i][0] === memberId) {
      return { exists: true, row: i + 2 }; // +2 for header and 0-index
    }
  }

  return { exists: false, row: null };
}

/**
 * Real-time duplicate Member ID validator for onEdit trigger
 * Call this from onEdit when Member ID column is modified
 * @param {Event} e - Edit event
 */
function validateMemberIdOnEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Only check Member Directory, Member ID column
  if (sheet.getName() !== SHEETS.MEMBER_DIR) return;
  if (range.getColumn() !== MEMBER_COLS.MEMBER_ID) return;

  var newValue = e.value;
  if (!newValue) return;

  var currentRow = range.getRow();
  var result = checkDuplicateMemberId(newValue);

  if (result.exists && result.row !== currentRow) {
    // Duplicate found in a different row
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Duplicate Member ID',
      'The Member ID "' + newValue + '" already exists in row ' + result.row + '.\n\n' +
      'Please use a unique Member ID.',
      ui.ButtonSet.OK
    );

    // Optionally highlight the cell
    range.setBackground('#FFCDD2'); // Light red

    // Revert to old value if available
    if (e.oldValue) {
      range.setValue(e.oldValue);
    }
  }
}

// ============================================================================
// GHOST VALIDATION FOR ORPHANED GRIEVANCES
// ============================================================================

/**
 * Find grievances with Member IDs that don't exist in Member Directory
 * @returns {Array<Object>} Array of orphaned grievance info
 */
function findOrphanedGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    return [];
  }

  if (memberSheet.getLastRow() < 2 || grievanceSheet.getLastRow() < 2) {
    return [];
  }

  // Build set of valid member IDs
  var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
  var validMemberIds = {};
  memberIds.forEach(function(row) {
    if (row[0]) validMemberIds[row[0]] = true;
  });

  // Check grievance member IDs
  var lastCol = Math.max(GRIEVANCE_COLS.GRIEVANCE_ID, GRIEVANCE_COLS.MEMBER_ID,
                         GRIEVANCE_COLS.FIRST_NAME, GRIEVANCE_COLS.LAST_NAME);
  var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, lastCol).getValues();
  var orphaned = [];

  grievanceData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];

    if (memberId && !validMemberIds[memberId]) {
      orphaned.push({
        row: index + 2,
        grievanceId: grievanceId,
        memberId: memberId,
        memberName: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')
      });
    }
  });

  return orphaned;
}

/**
 * Highlight orphaned grievances in the Grievance Log
 * Adds red background to rows with invalid Member IDs
 */
function highlightOrphanedGrievances() {
  var orphaned = findOrphanedGrievances();

  if (orphaned.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Data Integrity Check',
      'All grievances have valid Member IDs. No orphaned records found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Highlight orphaned rows
  orphaned.forEach(function(item) {
    grievanceSheet.getRange(item.row, 1, 1, grievanceSheet.getLastColumn())
      .setBackground('#FFCDD2'); // Light red
  });

  // Report findings
  var message = 'Found ' + orphaned.length + ' orphaned grievance(s):\n\n';
  orphaned.slice(0, 10).forEach(function(item) {
    message += '• Row ' + item.row + ': ' + item.grievanceId + ' (Member: ' + item.memberId + ')\n';
  });

  if (orphaned.length > 10) {
    message += '\n...and ' + (orphaned.length - 10) + ' more.';
  }

  message += '\n\nThese rows have been highlighted in red.';

  SpreadsheetApp.getUi().alert('⚠️ Orphaned Grievances Found', message, SpreadsheetApp.getUi().ButtonSet.OK);

  // Log to audit
  logIntegrityEvent('GHOST_VALIDATION', 'Found ' + orphaned.length + ' orphaned grievances');
}

/**
 * Run ghost validation automatically (for scheduled trigger)
 * Sends email to admin if orphans are found
 */
function runScheduledGhostValidation() {
  var orphaned = findOrphanedGrievances();

  if (orphaned.length > 0) {
    // Get admin emails from Config
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    var adminEmail = configSheet.getRange(3, CONFIG_COLS.ADMIN_EMAILS, 1, 1).getValue();

    if (adminEmail) {
      var subject = '⚠️ ' + COMMAND_CONFIG.SYSTEM_NAME + ': Orphaned Grievances Detected';
      var body = 'The scheduled data integrity check found ' + orphaned.length + ' grievances with invalid Member IDs.\n\n';

      orphaned.slice(0, 20).forEach(function(item) {
        body += '• Row ' + item.row + ': ' + item.grievanceId + ' - Member ID: ' + item.memberId + ' (' + item.memberName + ')\n';
      });

      if (orphaned.length > 20) {
        body += '\n...and ' + (orphaned.length - 20) + ' more.\n';
      }

      body += '\nPlease review and correct these records in the Grievance Log.';
      body += '\n\n--\n' + COMMAND_CONFIG.SYSTEM_NAME + ' Automated Alert';

      try {
        MailApp.sendEmail(adminEmail, subject, body);
        Logger.log('Orphan alert sent to: ' + adminEmail);
      } catch (e) {
        Logger.log('Failed to send orphan alert: ' + e.message);
      }
    }
  }
}

// ============================================================================
// STEWARD LOAD BALANCING METRICS
// ============================================================================

/**
 * Calculate steward workload metrics
 * @returns {Array<Object>} Array of steward workload data
 */
function calculateStewardWorkload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) return [];

  var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();

  // Count active grievances per steward
  var stewardCounts = {};
  var activeStatuses = ['Open', 'Pending Info', 'In Arbitration', 'Appealed'];

  grievanceData.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var steward = row[GRIEVANCE_COLS.STEWARD - 1];

    if (steward && activeStatuses.indexOf(status) !== -1) {
      if (!stewardCounts[steward]) {
        stewardCounts[steward] = {
          name: steward,
          activeCount: 0,
          urgentCount: 0,
          totalAssigned: 0
        };
      }
      stewardCounts[steward].activeCount++;
      stewardCounts[steward].totalAssigned++;

      // Check if urgent (days to deadline <= 7)
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
      if (typeof daysToDeadline === 'number' && daysToDeadline <= 7) {
        stewardCounts[steward].urgentCount++;
      }
    }
  });

  // Convert to array and calculate load scores
  var stewards = Object.keys(stewardCounts).map(function(name) {
    var data = stewardCounts[name];
    // Load score: active cases + (urgent cases * 2)
    data.loadScore = data.activeCount + (data.urgentCount * 2);
    return data;
  });

  // Sort by load score descending
  stewards.sort(function(a, b) { return b.loadScore - a.loadScore; });

  return stewards;
}

/**
 * Show steward workload dashboard dialog
 */
function showStewardWorkloadDashboard() {
  var stewards = calculateStewardWorkload();

  if (stewards.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Steward Workload',
      'No stewards have active grievances assigned.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Calculate statistics
  var totalActive = stewards.reduce(function(sum, s) { return sum + s.activeCount; }, 0);
  var avgLoad = totalActive / stewards.length;

  var html = '<html><head>' + getMobileOptimizedHead() + '</head><body><style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; margin-bottom: 10px; }' +
    '.stats { background: #F3F4F6; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    '.stat { display: inline-block; margin-right: 30px; }' +
    '.stat-value { font-size: 24px; font-weight: bold; color: #7C3AED; }' +
    '.stat-label { font-size: 12px; color: #6B7280; }' +
    'table { width: 100%; border-collapse: collapse; }' +
    'th { background: #7C3AED; color: white; padding: 10px; text-align: left; }' +
    'td { padding: 8px; border-bottom: 1px solid #E5E7EB; }' +
    '.high { background: #FEE2E2; }' +
    '.medium { background: #FEF3C7; }' +
    '.low { background: #D1FAE5; }' +
    '.load-bar { height: 10px; background: #E5E7EB; border-radius: 5px; }' +
    '.load-fill { height: 100%; border-radius: 5px; }' +
    '</style>';

  html += '<h2>Steward Workload Dashboard</h2>';

  html += '<div class="stats">' +
    '<div class="stat"><div class="stat-value">' + stewards.length + '</div><div class="stat-label">Active Stewards</div></div>' +
    '<div class="stat"><div class="stat-value">' + totalActive + '</div><div class="stat-label">Total Active Cases</div></div>' +
    '<div class="stat"><div class="stat-value">' + avgLoad.toFixed(1) + '</div><div class="stat-label">Avg Cases/Steward</div></div>' +
    '</div>';

  html += '<table><tr><th>Steward</th><th>Active</th><th>Urgent</th><th>Load Score</th><th>Status</th></tr>';

  var maxLoad = stewards[0] ? stewards[0].loadScore : 1;

  stewards.forEach(function(s) {
    var statusClass = s.loadScore > avgLoad * 1.5 ? 'high' : (s.loadScore > avgLoad ? 'medium' : 'low');
    var statusText = s.loadScore > avgLoad * 1.5 ? 'Overloaded' : (s.loadScore > avgLoad ? 'Busy' : 'Normal');
    var loadPct = (s.loadScore / maxLoad * 100).toFixed(0);
    var loadColor = statusClass === 'high' ? '#DC2626' : (statusClass === 'medium' ? '#F59E0B' : '#059669');

    html += '<tr class="' + statusClass + '">' +
      '<td><strong>' + escapeHtml(String(s.name)) + '</strong></td>' +
      '<td>' + escapeHtml(String(s.activeCount)) + '</td>' +
      '<td>' + escapeHtml(String(s.urgentCount)) + '</td>' +
      '<td>' +
        '<div class="load-bar"><div class="load-fill" style="width: ' + loadPct + '%; background: ' + loadColor + ';"></div></div>' +
        '<small>' + s.loadScore + '</small>' +
      '</td>' +
      '<td>' + statusText + '</td>' +
      '</tr>';
  });

  html += '</table>';
  html += '<p style="margin-top: 15px; color: #6B7280; font-size: 12px;">' +
    'Load Score = Active Cases + (Urgent Cases x 2). Urgent = deadline within 7 days.</p>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Steward Workload Analysis');
}

/**
 * Get steward with lowest workload for new case assignment
 * @returns {string|null} Name of steward with lowest load, or null if none
 */
function getStewardWithLowestWorkload() {
  var stewards = calculateStewardWorkload();

  // Also get all stewards from Config (some may have 0 cases)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet || configSheet.getLastRow() <= 1) return null;
  var stewardList = configSheet.getRange(3, CONFIG_COLS.STEWARDS, 50, 1).getValues()
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) { return row[0]; });

  // Find stewards with 0 active cases
  var activeStewardNames = stewards.map(function(s) { return s.name; });
  var availableStewards = stewardList.filter(function(name) {
    return activeStewardNames.indexOf(name) === -1;
  });

  if (availableStewards.length > 0) {
    return availableStewards[0]; // Return first available steward with 0 cases
  }

  // Otherwise return steward with lowest load score
  if (stewards.length > 0) {
    return stewards[stewards.length - 1].name;
  }

  return null;
}

// ============================================================================
// SELF-HEALING CONFIG VALIDATION TOOL
// ============================================================================

/**
 * Scan all data sheets for values not in Config dropdowns
 * @returns {Object} Report of missing config values by column
 */
function findMissingConfigValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet || !memberSheet || !grievanceSheet) {
    return { error: 'Required sheets not found' };
  }

  var report = { missingValues: [], autoFixable: [] };

  // Define field mappings to check
  var fieldsToCheck = [
    { sheet: memberSheet, col: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES, name: 'Job Title' },
    { sheet: memberSheet, col: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS, name: 'Work Location' },
    { sheet: memberSheet, col: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS, name: 'Unit' },
    { sheet: memberSheet, col: MEMBER_COLS.SUPERVISOR, configCol: CONFIG_COLS.SUPERVISORS, name: 'Supervisor' },
    { sheet: memberSheet, col: MEMBER_COLS.MANAGER, configCol: CONFIG_COLS.MANAGERS, name: 'Manager' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.STATUS, configCol: CONFIG_COLS.GRIEVANCE_STATUS, name: 'Grievance Status' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.CURRENT_STEP, configCol: CONFIG_COLS.GRIEVANCE_STEP, name: 'Grievance Step' },
    { sheet: grievanceSheet, col: GRIEVANCE_COLS.ISSUE_CATEGORY, configCol: CONFIG_COLS.ISSUE_CATEGORY, name: 'Issue Category' }
  ];

  fieldsToCheck.forEach(function(field) {
    // Get valid config values
    var configValues = configSheet.getRange(3, field.configCol, 100, 1).getValues()
      .filter(function(row) { return row[0] !== ''; })
      .map(function(row) { return row[0]; });

    var validSet = {};
    configValues.forEach(function(v) { validSet[v] = true; });

    // Get data values
    var lastRow = field.sheet.getLastRow();
    if (lastRow < 2) return;

    var dataValues = field.sheet.getRange(2, field.col, lastRow - 1, 1).getValues();

    // Find values not in config
    var missing = {};
    dataValues.forEach(function(row, index) {
      var value = row[0];
      // Skip pure-numeric values — they're data-entry errors, not text labels
      if (value && /^\d+$/.test(String(value).trim())) return;
      if (value && !validSet[value] && !missing[value]) {
        missing[value] = {
          field: field.name,
          value: value,
          configCol: field.configCol,
          exampleRow: index + 2
        };
      }
    });

    Object.values(missing).forEach(function(item) {
      report.missingValues.push(item);
      report.autoFixable.push(item);
    });
  });

  return report;
}

/**
 * Show Config health check dialog with auto-fix option
 */
function showConfigHealthCheck() {
  var report = findMissingConfigValues();

  if (report.error) {
    SpreadsheetApp.getUi().alert('Error', report.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (report.missingValues.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Config Health Check',
      'All dropdown values in your data sheets exist in the Config sheet.\n\nNo issues found!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var html = '<html><head>' + getMobileOptimizedHead() + '</head><body><style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #DC2626; }' +
    '.warning { background: #FEF3C7; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    'table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }' +
    'th { background: #7C3AED; color: white; padding: 10px; text-align: left; }' +
    'td { padding: 8px; border-bottom: 1px solid #E5E7EB; }' +
    '.btn { background: #059669; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; }' +
    '.btn:hover { background: #047857; }' +
    '</style>';

  html += '<h2>⚠️ Missing Config Values Found</h2>';
  html += '<div class="warning">' +
    '<strong>' + report.missingValues.length + ' value(s)</strong> in your data sheets are not in the Config dropdowns. ' +
    'This can cause validation errors and data inconsistency.' +
    '</div>';

  html += '<table><tr><th>Field</th><th>Missing Value</th><th>Example Row</th></tr>';

  report.missingValues.slice(0, 20).forEach(function(item) {
    html += '<tr><td>' + escapeHtml(String(item.field)) + '</td><td><strong>' + escapeHtml(String(item.value)) + '</strong></td><td>' + escapeHtml(String(item.exampleRow)) + '</td></tr>';
  });

  if (report.missingValues.length > 20) {
    html += '<tr><td colspan="3">...and ' + (report.missingValues.length - 20) + ' more</td></tr>';
  }

  html += '</table>';

  html += '<button class="btn" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).autoFixMissingConfigValues()">Auto-Add Missing Values to Config</button>';
  html += '<p style="margin-top: 10px; color: #6B7280; font-size: 12px;">This will add the missing values to the appropriate Config columns.</p>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Config Health Check');
}

/**
 * Auto-add missing values to Config sheet
 */
function autoFixMissingConfigValues() {
  var report = findMissingConfigValues();

  if (report.error || report.autoFixable.length === 0) {
    return;
  }

  // Group by config column
  var byColumn = {};
  report.autoFixable.forEach(function(item) {
    if (!byColumn[item.configCol]) {
      byColumn[item.configCol] = [];
    }
    byColumn[item.configCol].push(item.value);
  });

  // Add values to each column using the canonical write path
  Object.keys(byColumn).forEach(function(colStr) {
    var col = parseInt(colStr);
    var values = byColumn[col];

    values.forEach(function(value) {
      addToConfigDropdown_(col, value);
    });
  });

  // Log the fix
  logIntegrityEvent('CONFIG_AUTO_FIX', 'Added ' + report.autoFixable.length + ' missing values to Config');

  SpreadsheetApp.getUi().alert(
    '✅ Config Updated',
    'Successfully added ' + report.autoFixable.length + ' missing values to the Config sheet.\n\n' +
    'Please run "Setup Data Validations" from the Settings menu to refresh dropdowns.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// ENHANCED AUDIT LOGGING
// ============================================================================

/**
 * Log a data integrity event to the audit log with timestamp and user info
 * Note: This is separate from the detailed edit audit in Code.gs logAuditEvent()
 * @param {string} eventType - Type of event (e.g., 'STATUS_CHANGE', 'DATA_EDIT')
 * @param {string} details - Description of what happened
 * @param {Object} additionalInfo - Optional additional information
 */
/**
 * @deprecated Use logAuditEvent() instead. This is a legacy wrapper kept for backward compatibility.
 * Delegates to logAuditEvent() which handles email masking and integrity hashing.
 */
function logIntegrityEvent(eventType, details, additionalInfo) {
  // Consolidate into the primary logAuditEvent function
  var combinedDetails = {};
  if (details) combinedDetails.details = details;
  if (additionalInfo) {
    if (typeof additionalInfo === 'object') {
      for (var key in additionalInfo) {
        if (additionalInfo.hasOwnProperty(key)) {
          combinedDetails[key] = additionalInfo[key];
        }
      }
    } else {
      combinedDetails.additionalInfo = additionalInfo;
    }
  }
  logAuditEvent(eventType, combinedDetails);
}

/**
 * Log grievance status change (call from onEdit trigger)
 * @param {string} grievanceId - The grievance ID
 * @param {string} oldStatus - Previous status
 * @param {string} newStatus - New status
 */
function logGrievanceStatusChange(grievanceId, oldStatus, newStatus) {
  logIntegrityEvent('STATUS_CHANGE',
    'Grievance ' + grievanceId + ' status changed from "' + oldStatus + '" to "' + newStatus + '"',
    { grievanceId: grievanceId, oldStatus: oldStatus, newStatus: newStatus }
  );
}

/**
 * Log steward assignment change
 * @param {string} grievanceId - The grievance ID
 * @param {string} oldSteward - Previous steward
 * @param {string} newSteward - New steward
 */
function logStewardAssignmentChange(grievanceId, oldSteward, newSteward) {
  logIntegrityEvent('STEWARD_CHANGE',
    'Grievance ' + grievanceId + ' reassigned from "' + (oldSteward || 'None') + '" to "' + newSteward + '"',
    { grievanceId: grievanceId, oldSteward: oldSteward, newSteward: newSteward }
  );
}

/**
 * Show audit log viewer dialog
 */
function showAuditLogViewer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!auditSheet || auditSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert(
      'Audit Log',
      'No audit log entries found. The audit log records important changes like status updates and steward assignments.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Get recent entries
  var lastRow = auditSheet.getLastRow();
  var startRow = Math.max(2, lastRow - 49); // Last 50 entries
  var data = auditSheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();

  var html = '<html><head>' + getMobileOptimizedHead() + '</head><body><style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; }' +
    'table { width: 100%; border-collapse: collapse; font-size: 12px; }' +
    'th { background: #7C3AED; color: white; padding: 8px; text-align: left; position: sticky; top: 0; }' +
    'td { padding: 6px; border-bottom: 1px solid #E5E7EB; }' +
    'tr:hover { background: #F3F4F6; }' +
    '.status { background: #DBEAFE; padding: 2px 6px; border-radius: 4px; }' +
    '.steward { background: #D1FAE5; padding: 2px 6px; border-radius: 4px; }' +
    '.config { background: #FEF3C7; padding: 2px 6px; border-radius: 4px; }' +
    '</style>';

  html += '<h2>Recent Audit Log Entries</h2>';
  html += '<p style="color: #6B7280;">Showing last ' + data.length + ' entries</p>';

  html += '<table><tr><th>Timestamp</th><th>User</th><th>Event</th><th>Details</th></tr>';

  // Reverse to show most recent first
  data.reverse().forEach(function(row) {
    var timestamp = row[EVENT_AUDIT_COLS.TIMESTAMP - 1] instanceof Date ? row[EVENT_AUDIT_COLS.TIMESTAMP - 1].toLocaleString() : row[EVENT_AUDIT_COLS.TIMESTAMP - 1];
    var user = row[EVENT_AUDIT_COLS.USER - 1] ? row[EVENT_AUDIT_COLS.USER - 1].split('@')[0] : 'Unknown';
    var eventType = row[EVENT_AUDIT_COLS.EVENT_TYPE - 1] || '';
    var details = row[EVENT_AUDIT_COLS.DETAILS - 1] || '';

    var eventClass = eventType.indexOf('STATUS') !== -1 ? 'status' :
                     (eventType.indexOf('STEWARD') !== -1 ? 'steward' :
                     (eventType.indexOf('CONFIG') !== -1 ? 'config' : ''));

    html += '<tr>' +
      '<td>' + escapeHtml(String(timestamp)) + '</td>' +
      '<td>' + escapeHtml(String(user)) + '</td>' +
      '<td><span class="' + eventClass + '">' + escapeHtml(String(eventType)) + '</span></td>' +
      '<td>' + escapeHtml(String(details).substring(0, 100)) + (details.length > 100 ? '...' : '') + '</td>' +
      '</tr>';
  });

  html += '</table>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Audit Log Viewer');
}

// ============================================================================
// AUTO-ARCHIVE FOR CLOSED GRIEVANCES
// ============================================================================

/**
 * Archive closed grievances older than specified days
 * @param {number} daysOld - Archive grievances closed more than this many days ago (default 90)
 */
function archiveClosedGrievances(daysOld) {
  daysOld = daysOld || 90;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    Logger.log('Grievance Log not found');
    return { archived: 0 };
  }

  // Get or create archive sheet
  var archiveSheetName = '_Archive_Grievances';
  var archiveSheet = ss.getSheetByName(archiveSheetName);

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveSheetName);
    setSheetVeryHidden_(archiveSheet);

    // Copy headers
    var headers = grievanceSheet.getRange(1, 1, 1, grievanceSheet.getLastColumn()).getValues();
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
      .setFontWeight('bold')
      .setBackground('#6B7280')
      .setFontColor('#FFFFFF');
  }

  // Find rows to archive
  var closedStatuses = ['Closed', 'Won', 'Denied', 'Settled', 'Withdrawn'];
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysOld);

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return { archived: 0 };

  // H-18: Acquire lock so read-archive-delete is atomic; prevents duplicate archival on concurrent calls
  var archiveLock = LockService.getScriptLock();
  if (!archiveLock.tryLock(30000)) {
    Logger.log('archiveClosedGrievances: could not acquire lock');
    return { archived: 0, error: 'Lock unavailable' };
  }
  try {
    var data = grievanceSheet.getRange(2, 1, lastRow - 1, grievanceSheet.getLastColumn()).getValues();

    var rowsToArchive = [];
    var rowIndicesToDelete = [];

    data.forEach(function(row, index) {
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];

      if (closedStatuses.indexOf(status) !== -1 && dateClosed instanceof Date && dateClosed < cutoffDate) {
        rowsToArchive.push(row);
        rowIndicesToDelete.push(index + 2); // +2 for header and 0-index
      }
    });

    if (rowsToArchive.length === 0) {
      return { archived: 0 };
    }

    // Append to archive sheet
    var archiveLastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(archiveLastRow + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
      .setValues(rowsToArchive);

    // Delete from main sheet (in reverse order to maintain row indices)
    // Transaction pattern: track individual failures and report at the end
    var failedDeletes = [];
    var _successfulDeletes = 0;
    rowIndicesToDelete.reverse().forEach(function(rowIndex) {
      try {
        grievanceSheet.deleteRow(rowIndex);
        _successfulDeletes++;
      } catch (deleteErr) {
        failedDeletes.push({ row: rowIndex, error: deleteErr.message });
        Logger.log('Failed to delete row ' + rowIndex + ': ' + deleteErr.message);
      }
    });
    if (failedDeletes.length > 0) {
      var failedRows = failedDeletes.map(function(f) { return f.row; });
      Logger.log('Warning: ' + failedDeletes.length + ' rows could not be deleted and may exist in both archive and main sheet: ' + failedRows.join(', '));
    }

    // Log the archive operation
    logIntegrityEvent('AUTO_ARCHIVE',
      'Archived ' + rowsToArchive.length + ' closed grievances older than ' + daysOld + ' days',
      { count: rowsToArchive.length, daysOld: daysOld, deleteFailed: failedDeletes.length }
    );

    // F141: Include failedGrievanceIds in return for error reporting
    var failedGrievanceIds = [];
    if (failedDeletes.length > 0) {
      failedDeletes.forEach(function(rowIdx) {
        var dataIdx = rowIdx - 2;
        if (dataIdx >= 0 && dataIdx < data.length) {
          var gId = data[dataIdx][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
          if (gId) failedGrievanceIds.push(String(gId));
        }
      });
    }

    return {
      archived: rowsToArchive.length,
      failedDeletes: failedDeletes.length,
      failedGrievanceIds: failedGrievanceIds
    };
  } finally {
    archiveLock.releaseLock();
  }
}

/**
 * Show archive dialog with options
 */
function showArchiveDialog() {
  var html = '<html><head>' + getMobileOptimizedHead() + '</head><body><style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; }' +
    'h2 { color: #7C3AED; }' +
    '.info { background: #EFF6FF; padding: 15px; border-radius: 8px; margin-bottom: 20px; }' +
    'label { display: block; margin-bottom: 5px; font-weight: bold; }' +
    'input[type="number"] { width: 100px; padding: 8px; border: 1px solid #D1D5DB; border-radius: 4px; margin-bottom: 15px; }' +
    '.btn { background: #7C3AED; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; }' +
    '.btn:hover { background: #6D28D9; }' +
    '.btn-secondary { background: #6B7280; margin-left: 10px; }' +
    '</style>';

  html += '<h2>Archive Closed Grievances</h2>';

  html += '<div class="info">' +
    '<strong>What this does:</strong><br>' +
    'Moves closed grievances (Won, Denied, Settled, Withdrawn, Closed) to a hidden archive sheet. ' +
    'This keeps your active Grievance Log fast and focused on current cases.' +
    '</div>';

  html += '<label>Archive grievances closed more than:</label>';
  html += '<input type="number" id="daysOld" value="90" min="30" max="365"> days ago';

  html += '<br><br>';
  html += '<button class="btn" onclick="runArchive()">Archive Now</button>';
  html += '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>';

  html += '<script>' +
    'function runArchive() {' +
    '  var days = document.getElementById("daysOld").value;' +
    '  google.script.run.withSuccessHandler(function(result) {' +
    '    alert("Archived " + result.archived + " grievances.");' +
    '    google.script.host.close();' +
    '  }).archiveClosedGrievances(parseInt(days));' +
    '}' +
    '</script>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Archive Grievances');
}

/**
 * Restore grievances from archive
 * @param {Array<string>} grievanceIds - Array of grievance IDs to restore
 */
function restoreFromArchive(grievanceIds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var archiveSheet = ss.getSheetByName('_Archive_Grievances');

  if (!archiveSheet || !grievanceSheet) {
    return { restored: 0, error: 'Required sheets not found' };
  }

  if (archiveSheet.getLastRow() <= 1) {
    return { restored: 0 };
  }

  var archiveData = archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archiveSheet.getLastColumn()).getValues();

  var rowsToRestore = [];
  var rowIndicesToDelete = [];

  archiveData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    if (grievanceIds.indexOf(grievanceId) !== -1) {
      rowsToRestore.push(row);
      rowIndicesToDelete.push(index + 2);
    }
  });

  if (rowsToRestore.length === 0) {
    return { restored: 0 };
  }

  // Add back to main sheet
  var lastRow = grievanceSheet.getLastRow();
  grievanceSheet.getRange(lastRow + 1, 1, rowsToRestore.length, rowsToRestore[0].length)
    .setValues(rowsToRestore);

  // Remove from archive (track failures rather than stopping on first error)
  var restoreFailedDeletes = [];
  rowIndicesToDelete.reverse().forEach(function(rowIndex) {
    try {
      archiveSheet.deleteRow(rowIndex);
    } catch (deleteErr) {
      restoreFailedDeletes.push({ row: rowIndex, error: deleteErr.message });
      Logger.log('Failed to delete archive row ' + rowIndex + ': ' + deleteErr.message);
    }
  });

  logIntegrityEvent('ARCHIVE_RESTORE',
    'Restored ' + rowsToRestore.length + ' grievances from archive',
    { grievanceIds: grievanceIds, deleteFailed: restoreFailedDeletes.length }
  );

  return {
    restored: rowsToRestore.length,
    failedDeletes: restoreFailedDeletes.length,
    failedRows: restoreFailedDeletes.length > 0 ? restoreFailedDeletes : undefined
  };
}

// ============================================================================
// VISUAL DEADLINE HEATMAP WITH SPARKLINES
// ============================================================================

/**
 * Apply deadline heatmap conditional formatting to Grievance Log
 * Colors cells based on urgency level
 */
function applyDeadlineHeatmap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return;

  // Days to Deadline column
  var deadlineRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);

  // Clear existing conditional formatting for this range
  var rules = grievanceSheet.getConditionalFormatRules();
  var newRules = rules.filter(function(rule) {
    var ranges = rule.getRanges();
    return !ranges.some(function(r) {
      return r.getColumn() === GRIEVANCE_COLS.DAYS_TO_DEADLINE;
    });
  });

  // Create new rules for heatmap
  // Overdue (negative or "Overdue" text)
  var overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground('#DC2626')  // Bright red
    .setFontColor('#FFFFFF')
    .setBold(true)
    .setRanges([deadlineRange])
    .build();

  // Critical (1-3 days)
  var criticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 3)
    .setBackground('#F87171')  // Light red
    .setFontColor('#7F1D1D')
    .setBold(true)
    .setRanges([deadlineRange])
    .build();

  // Warning (4-7 days)
  var warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 7)
    .setBackground('#FBBF24')  // Yellow/orange
    .setFontColor('#78350F')
    .setRanges([deadlineRange])
    .build();

  // Caution (8-14 days)
  var cautionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(8, 14)
    .setBackground('#FEF3C7')  // Light yellow
    .setFontColor('#92400E')
    .setRanges([deadlineRange])
    .build();

  // Safe (15+ days)
  var safeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(14)
    .setBackground('#D1FAE5')  // Light green
    .setFontColor('#065F46')
    .setRanges([deadlineRange])
    .build();

  newRules.push(overdueRule, criticalRule, warningRule, cautionRule, safeRule);
  grievanceSheet.setConditionalFormatRules(newRules);

  // Also apply to status column for closed cases
  var statusRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1);

  var wonRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Won')
    .setBackground('#059669')  // Green
    .setFontColor('#FFFFFF')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  var deniedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Denied')
    .setBackground('#DC2626')  // Red
    .setFontColor('#FFFFFF')
    .setRanges([statusRange])
    .build();

  var settledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Settled')
    .setBackground('#7C3AED')  // Purple
    .setFontColor('#FFFFFF')
    .setRanges([statusRange])
    .build();

  var existingRules = grievanceSheet.getConditionalFormatRules();
  existingRules.push(wonRule, deniedRule, settledRule);
  grievanceSheet.setConditionalFormatRules(existingRules);

  SpreadsheetApp.getUi().alert(
    '✅ Heatmap Applied',
    'Deadline heatmap has been applied to the Grievance Log:\n\n' +
    '🔴 Red: Overdue or 1-3 days\n' +
    '🟡 Yellow: 4-7 days\n' +
    '🟢 Green: 15+ days',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// MOBILE STEWARD PORTAL VIEW
// ============================================================================

/**
 * Create or update the mobile steward portal sheet
 * Shows only essential columns for mobile access
 */
function createMobileStewardPortal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var portalSheetName = '📱 Steward Portal';

  // Get or create portal sheet
  var portalSheet = ss.getSheetByName(portalSheetName);
  if (portalSheet) {
    portalSheet.clear();
  } else {
    portalSheet = ss.insertSheet(portalSheetName);
  }

  // Get grievance data
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceSheet) {
    portalSheet.getRange('A1').setValue('Error: Grievance Log not found');
    return;
  }

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    portalSheet.getRange('A1').setValue('No grievances found');
    return;
  }

  // Title
  portalSheet.getRange('A1').setValue('📱 STEWARD PORTAL')
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor('#7C3AED');
  portalSheet.getRange('A1:E1').merge();

  portalSheet.getRange('A2').setValue('Quick access to your active cases. Updated: ' + new Date().toLocaleString())
    .setFontStyle('italic')
    .setFontColor('#6B7280');
  portalSheet.getRange('A2:E2').merge();

  // Section: Urgent Cases (deadline <= 7 days)
  portalSheet.getRange('A4').setValue('🚨 URGENT CASES')
    .setFontWeight('bold')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF');
  portalSheet.getRange('A4:E4').merge();

  // Headers for mobile view
  var headers = ['ID', 'Member', 'Status', 'Deadline', 'Steward'];
  portalSheet.getRange('A5:E5').setValues([headers])
    .setFontWeight('bold')
    .setBackground('#F3F4F6');

  // Get grievance data and filter
  var grievanceData = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var activeStatuses = ['Open', 'Pending Info', 'In Arbitration', 'Appealed'];

  var urgentCases = [];
  var normalCases = [];

  grievanceData.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    if (activeStatuses.indexOf(status) === -1) return;

    var mobileRow = [
      row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
      (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '').charAt(0) + '.',
      row[GRIEVANCE_COLS.STATUS - 1],
      row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1],
      row[GRIEVANCE_COLS.STEWARD - 1]
    ];

    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    if (typeof daysToDeadline === 'number' && daysToDeadline <= 7) {
      urgentCases.push(mobileRow);
    } else {
      normalCases.push(mobileRow);
    }
  });

  // Sort urgent cases by deadline ascending
  urgentCases.sort(function(a, b) { return (a[3] || 999) - (b[3] || 999); });
  normalCases.sort(function(a, b) { return (a[3] || 999) - (b[3] || 999); });

  var currentRow = 6;

  // Write urgent cases
  if (urgentCases.length > 0) {
    portalSheet.getRange(currentRow, 1, urgentCases.length, 5).setValues(urgentCases)
      .setBackground('#FEE2E2');
    currentRow += urgentCases.length + 1;
  } else {
    portalSheet.getRange(currentRow, 1).setValue('No urgent cases!')
      .setFontColor('#059669');
    currentRow += 2;
  }

  // Section: All Active Cases
  portalSheet.getRange(currentRow, 1).setValue('📋 ALL ACTIVE CASES')
    .setFontWeight('bold')
    .setBackground('#7C3AED')
    .setFontColor('#FFFFFF');
  portalSheet.getRange(currentRow, 1, 1, 5).merge();
  currentRow++;

  // Headers
  portalSheet.getRange(currentRow, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#F3F4F6');
  currentRow++;

  // Write normal cases
  if (normalCases.length > 0) {
    portalSheet.getRange(currentRow, 1, normalCases.length, 5).setValues(normalCases);
  } else {
    portalSheet.getRange(currentRow, 1).setValue('No other active cases');
  }

  // Format columns for mobile
  portalSheet.setColumnWidth(1, 90);  // ID
  portalSheet.setColumnWidth(2, 120); // Member
  portalSheet.setColumnWidth(3, 80);  // Status
  portalSheet.setColumnWidth(4, 70);  // Deadline
  portalSheet.setColumnWidth(5, 100); // Steward

  // Freeze header
  portalSheet.setFrozenRows(3);

  // Move to front
  ss.setActiveSheet(portalSheet);
  ss.moveActiveSheet(2);

  SpreadsheetApp.getUi().alert(
    '📱 Steward Portal Created',
    'The mobile-friendly Steward Portal has been created/updated.\n\n' +
    'This view shows:\n' +
    '• Urgent cases (7 days or less)\n' +
    '• All active cases\n' +
    '• Narrow columns optimized for mobile',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// ENHANCED onEdit TRIGGER WITH AUDIT LOGGING
// ============================================================================

/**
 * Enhanced onEdit handler that logs important changes
 * Add this to your onEdit trigger
 * @param {Event} e - Edit event
 */
function onEditWithAuditLogging(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var oldValue = e.oldValue;
  var newValue = e.value;

  // Only process single cell edits
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  var sheetName = sheet.getName();
  var col = range.getColumn();
  var row = range.getRow();

  // Skip header row
  if (row === 1) return;

  // Grievance Log changes
  if (sheetName === SHEETS.GRIEVANCE_LOG) {
    var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

    // Status change
    if (col === GRIEVANCE_COLS.STATUS && oldValue !== newValue) {
      logGrievanceStatusChange(grievanceId, oldValue, newValue);
    }

    // Steward assignment change
    if (col === GRIEVANCE_COLS.STEWARD && oldValue !== newValue) {
      logStewardAssignmentChange(grievanceId, oldValue, newValue);
    }
  }

  // Member Directory - duplicate ID check
  if (sheetName === SHEETS.MEMBER_DIR && col === MEMBER_COLS.MEMBER_ID) {
    validateMemberIdOnEdit(e);
  }
}

// ============================================================================
// MENU ADDITIONS
// ============================================================================

/**
 * Add Data Integrity menu items to the existing menu structure
 * Call this from onOpen or add to existing menu creation
 */
function addDataIntegrityMenuItems() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('🛡️ Data Integrity')
    .addItem('🔍 Find Orphaned Grievances', 'highlightOrphanedGrievances')
    .addItem('⚙️ Config Health Check', 'showConfigHealthCheck')
    .addSeparator()
    .addItem('📊 Steward Workload Dashboard', 'showStewardWorkloadDashboard')
    .addItem('📱 Create/Update Steward Portal', 'createMobileStewardPortal')
    .addSeparator()
    .addItem('🎨 Apply Deadline Heatmap', 'applyDeadlineHeatmap')
    .addItem('📦 Archive Closed Grievances', 'showArchiveDialog')
    .addSeparator()
    .addItem('📜 View Audit Log', 'showAuditLogViewer')
    .addToUi();
}
