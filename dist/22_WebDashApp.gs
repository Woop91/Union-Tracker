/**
 * WebApp.gs
 * Main entry point for the web app dashboard.
 *
 * doGet(e) handles:
 *   1. Auth resolution (SSO / magic link / session token)
 *   2. Role lookup from Member Directory
 *   3. Routing to the correct view (auth, steward, member)
 *   4. Injecting config + user data into the HTML template
 *
 * Deployment: Deploy as Web App
 *   - Execute as: Me (the script owner)
 *   - Who has access: Anyone (or anyone within org)
 *   - Note: Users arrive via Bitly redirect, not the raw URL
 *
 * Deep-link support: ?page=workload (or any tab name) opens the SPA
 * at that tab after authentication.
 */

/**
 * GAS web app entry point — serves the single-page app.
 * @param {Object} e - Event object with URL parameters
 * @returns {HtmlOutput}
 */
function doGet(e) {
  e = e || { parameter: {} };

  try {
    return doGetWebDashboard(e);
  } catch (fatalErr) {
    Logger.log('doGet FATAL: ' + fatalErr.message + '\n' + (fatalErr.stack || ''));
    // Log additional context for debugging
    try {
      Logger.log('doGet FATAL context: ConfigReader=' + (typeof ConfigReader) +
        ', Auth=' + (typeof Auth) +
        ', DataService=' + (typeof DataService) +
        ', SHEETS=' + (typeof SHEETS));
    } catch (_) { /* best-effort */ }
    return _serveFatalError(fatalErr.message);
  }
}

/**
 * Main dashboard handler — serves the single-page app.
 * @param {Object} e - Event object with URL parameters
 * @returns {HtmlOutput}
 */
function doGetWebDashboard(e) {
  e = e || { parameter: {} };

  // Load config once at top — reused in both success and error paths (Fix 2.1)
  var config;
  try {
    config = ConfigReader.getConfig();
  } catch (cfgErr) {
    Logger.log('doGetWebDashboard: config load failed: ' + cfgErr.message);
    config = { orgName: 'Dashboard', orgAbbrev: '', logoInitials: '', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' };
  }

  try {
    var user = Auth.resolveUser(e);

    if (!user) {
      // Explicit logout — show login page with signed-out message
      if (e.parameter.loggedout === '1') {
        return _serveAuth(config, e, 'loggedout');
      }
      // If the user explicitly clicked "Continue with Google" (sso=1 param)
      // but SSO failed, redirect back with an error flag so the login page
      // can display a helpful message.
      if (e.parameter.sso === '1') {
        return _serveAuth(config, e, 'sso_failed');
      }
      // Not authenticated — show login screen
      return _serveAuth(config, e);
    }

    // Authenticated — look up role
    var userRecord = DataService.findUserByEmail(user.email);

    if (!userRecord) {
      // Check if the visitor is the script owner — grant bootstrap access
      // so the app is usable before the Member Directory is populated.
      try {
        var ownerEmail = Session.getActiveUser().getEmail().toLowerCase();
        if (user.email && user.email.toLowerCase() === ownerEmail) {
          userRecord = {
            email: user.email,
            name: 'Admin',
            firstName: 'Admin',
            lastName: '',
            role: 'both',
            unit: 'Admin',
            joined: '',
            duesStatus: 'Active',
            phone: '',
            isBootstrapAdmin: true,
          };
        }
      } catch (_ownerErr) { /* SSO not available — fall through */ }
    }

    if (!userRecord) {
      // Email not in directory
      return _serveError(config, 'not_found', user.email);
    }

    // Handle "remember me" — create session token if requested
    // Also echo back an existing validated session token so the client always has
    // SESSION_TOKEN populated for non-SSO auth (magic link + remember me, or returning
    // session users). This is safe: the token was already validated by resolveUser().
    var sessionToken = null;
    if (e.parameter.remember === '1' && user.method === 'magic') {
      sessionToken = Auth.createSessionToken(user.email);
    } else if (user.method === 'session' && e.parameter.sessionToken) {
      // Echo back the already-validated session token so the client can use it
      sessionToken = e.parameter.sessionToken;
    }

    // Route to appropriate dashboard
    var role = userRecord.role; // 'steward', 'member', or 'both'
    var initialTab = e.parameter.page || null;

    return _serveDashboard(config, userRecord, role, sessionToken, initialTab);

  } catch (err) {
    Logger.log('WebApp doGet error: ' + err.message + '\n' + err.stack);
    return _serveError(config, 'error', err.message);
  }
}

/**
 * Serves the login/auth page.
 * @param {Object} config
 * @param {Object} e - Request event
 * @param {string} [authError] - Optional auth error code (e.g. 'sso_failed')
 */
function _serveAuth(config, e, authError) {
  var template = HtmlService.createTemplateFromFile('index');
  template.view = 'auth';

  template.pageData = JSON.stringify({
    view: 'auth',
    config: _sanitizeConfig(config),
    error: e.parameter.authError || authError || null,
    webAppUrl: _getWebAppUrlSafe(),
    tokenChecked: !!(e.parameter.sessionToken),
  });

  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Serves the dashboard (steward or member view).
 */
function _serveDashboard(config, userRecord, role, sessionToken, initialTab) {
  var template = HtmlService.createTemplateFromFile('index');
  template.view = role; // 'steward', 'member', or 'both'

  // Sanitize user record — strip sensitive fields
  var safeUser = {
    email: userRecord.email,
    name: userRecord.name,
    firstName: userRecord.firstName,
    lastName: userRecord.lastName,
    role: role,
    unit: userRecord.unit,
    joined: userRecord.joined,
    duesStatus: userRecord.duesStatus,
    // duesPaying: true=paying, false=not paying, null=column absent (treated as paying)
    duesPaying: (userRecord.duesPaying !== undefined) ? userRecord.duesPaying : null,
    phone: userRecord.phone,
    workLocation: userRecord.workLocation || '',
    officeDays: userRecord.officeDays || '',
    assignedSteward: userRecord.assignedSteward || '',
    hasOpenGrievance: userRecord.hasOpenGrievance || false,
    sharePhone: userRecord.sharePhone === true,  // steward phone opt-in; false if column absent
  };

  template.pageData = JSON.stringify({
    view: role === 'steward' || role === 'both' ? 'steward' : 'member',
    config: _sanitizeConfig(config),
    user: safeUser,
    isDualRole: role === 'steward' || role === 'both',
    sessionToken: sessionToken || null,
    initialTab: initialTab || null,
    webAppUrl: _getWebAppUrlSafe(),
  });

  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Serves an error page.
 */
function _serveError(config, type, detail) {
  var template = HtmlService.createTemplateFromFile('index');
  template.view = 'error';

  var messages = {
    'not_found': 'Your email was not found in the member directory. Please contact your steward for access.',
    'expired': 'Your sign-in link has expired. Please request a new one.',
    'error': 'Something went wrong. Please try again.',
  };

  // Log the actual error detail server-side; never expose raw error messages to the client
  if (type === 'error' && detail) {
    Logger.log('WebApp _serveError: ' + detail);
  }

  template.pageData = JSON.stringify({
    view: 'error',
    config: _sanitizeConfig(config),
    error: {
      type: type,
      message: messages[type] || messages['error'],
    },
  });

  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Strips internal-only fields from config before sending to client.
 */
function _sanitizeConfig(config) {
  return {
    orgName: config.orgName,
    orgAbbrev: config.orgAbbrev,
    logoInitials: config.logoInitials,
    accentHue: config.accentHue,
    stewardLabel: config.stewardLabel,
    memberLabel: config.memberLabel,
    magicLinkExpiryDays: config.magicLinkExpiryDays,
    cookieDurationDays: config.cookieDurationDays,
    calendarUrl: config.calendarId ? 'https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(config.calendarId) : '',
    calendarId: config.calendarId || '',
    driveFolderUrl: config.driveFolderId ? 'https://drive.google.com/drive/folders/' + config.driveFolderId : '',
    // surveyFormUrl removed v4.22.7 — survey is native webapp (renderSurveyFormPage in member_view.html)
    orgWebsite: config.orgWebsite || '',
    broadcastScopeAll: (String(config.broadcastScopeAll || '').trim().toLowerCase() === 'yes'),
    // v4.20.18: folder IDs needed client-side for Drive links and warnings
    minutesFolderId:    config.minutesFolderId    || '',
    grievancesFolderId: config.grievancesFolderId || '',
    // v4.20.18: insights cache TTL exposed so client can show staleness info
    insightsCacheTTLMin: config.insightsCacheTTLMin || 5,
  };
}

/**
 * Last-resort error page — zero external dependencies.
 * Rendered when doGet() itself throws (e.g. ConfigReader unavailable,
 * missing sheets, deployment misconfiguration).  Uses only HtmlService
 * so the user sees a helpful message instead of the generic Google
 * "Sorry, unable to open the file" page.
 * @param {string} detail - Internal error detail (logged, NOT shown to user)
 * @returns {HtmlOutput}
 */
function _serveFatalError(detail) {
  if (detail) Logger.log('_serveFatalError detail: ' + detail);

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Dashboard — Error</title>'
    + '<style>'
    + 'body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;'
    + 'font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#141414;color:#e0e0e0}'
    + '.card{max-width:440px;padding:40px 32px;background:#1e1e1e;border-radius:16px;text-align:center}'
    + 'h1{font-size:20px;margin:0 0 12px}p{color:#888;font-size:14px;line-height:1.6;margin:0 0 20px}'
    + 'a{display:inline-block;padding:10px 24px;background:hsl(250,70%,68%);color:#fff;'
    + 'border-radius:8px;text-decoration:none;font-weight:600}'
    + '</style></head><body><div class="card">'
    + '<h1>Something went wrong</h1>'
    + '<p>The dashboard could not load. This is usually temporary &mdash; '
    + 'please wait a moment and try again.</p>'
    + '<a href="javascript:location.reload()">Reload</a>'
    + '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Dashboard — Error')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Returns the web app URL, or empty string if unavailable.
 * Prevents null from being injected into pageData JSON.
 * @returns {string}
 * @private
 */
function _getWebAppUrlSafe() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (_e) {
    Logger.log('_getWebAppUrlSafe: ScriptApp.getService().getUrl() failed: ' + _e.message);
    return '';
  }
}

/**
 * Include helper — allows <?!= include('filename') ?> in HTML templates.
 * Used to compose the SPA from multiple HTML files.
 * @param {string} filename - Name of the HTML file (without .html extension)
 * @returns {string} Raw HTML content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Client-callable: Returns the org chart HTML content for lazy-loading.
 * Loaded on-demand when the user navigates to the Org Chart tab.
 * @returns {string} Raw HTML content (CSS-scoped under .oc-wrap), or error message
 */
function getOrgChartHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('org_chart').getContent();
  } catch (e) {
    Logger.log('getOrgChartHtml error: ' + e.message);
    return '<div class="empty-state">Org chart could not be loaded.</div>';
  }
}

/**
 * Client-callable: Returns the POMS Reference HTML for lazy-loading.
 * Loaded on-demand when the user navigates to the POMS Reference tab.
 * @returns {string} Raw HTML content (CSS-scoped under .poms-root), or error message
 */
function getPOMSReferenceHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('poms_reference').getContent();
  } catch (e) {
    Logger.log('getPOMSReferenceHtml error: ' + e.message);
    return '<div class="empty-state">POMS Reference could not be loaded.</div>';
  }
}

/**
 * Returns the published web app URL. Used by client-side logout
 * as a reload fallback when window.top.location.reload() is blocked.
 * @returns {string}
 */
function getWebAppUrl() {
  return _getWebAppUrlSafe();
}

/**
 * DIAGNOSTIC: Run from Apps Script editor to test the full doGet loading chain.
 * Reports each step's success/failure with timing info.
 * Usage: Open Apps Script editor → Select diagnoseWebApp → Run → View Logs.
 */
function diagnoseWebApp() {
  var results = [];
  var start = Date.now();

  function step(name, fn) {
    var t0 = Date.now();
    try {
      var val = fn();
      results.push({ step: name, ok: true, ms: Date.now() - t0, detail: val });
      Logger.log('✓ ' + name + ' (' + (Date.now() - t0) + 'ms)');
      return val;
    } catch (err) {
      results.push({ step: name, ok: false, ms: Date.now() - t0, error: err.message });
      Logger.log('✗ ' + name + ' (' + (Date.now() - t0) + 'ms): ' + err.message);
      return null;
    }
  }

  // 1. Spreadsheet binding
  var ss = step('SpreadsheetApp.getActiveSpreadsheet()', function() {
    var s = SpreadsheetApp.getActiveSpreadsheet();
    if (!s) throw new Error('null — script is not bound to a spreadsheet');
    return s;
  });

  // 2. Config sheet exists
  step('Config sheet lookup', function() {
    if (!ss) throw new Error('skipped — no spreadsheet');
    var sheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!sheet) throw new Error('Config sheet "' + SHEETS.CONFIG + '" not found');
    return 'found (' + sheet.getLastRow() + ' rows)';
  });

  // 3. ConfigReader
  var config = step('ConfigReader.getConfig()', function() {
    var c = ConfigReader.getConfig(true);
    return 'orgName=' + c.orgName;
  });

  // 4. Member Directory sheet
  step('Member Directory sheet', function() {
    if (!ss) throw new Error('skipped');
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) throw new Error('Sheet "' + SHEETS.MEMBER_DIR + '" not found');
    return 'found (' + sheet.getLastRow() + ' rows)';
  });

  // 5. Grievance Log sheet
  step('Grievance Log sheet', function() {
    if (!ss) throw new Error('skipped');
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) throw new Error('Sheet "' + SHEETS.GRIEVANCE_LOG + '" not found');
    return 'found (' + sheet.getLastRow() + ' rows)';
  });

  // 6. Auth module
  step('Auth module available', function() {
    if (typeof Auth === 'undefined') throw new Error('Auth is undefined');
    if (typeof Auth.resolveUser !== 'function') throw new Error('Auth.resolveUser missing');
    return 'OK';
  });

  // 7. DataService module
  step('DataService module available', function() {
    if (typeof DataService === 'undefined') throw new Error('DataService is undefined');
    if (typeof DataService.findUserByEmail !== 'function') throw new Error('findUserByEmail missing');
    if (typeof DataService.getBatchData !== 'function') throw new Error('getBatchData missing');
    return 'OK';
  });

  // 8. SSO identity
  step('Session.getActiveUser()', function() {
    var email = Session.getActiveUser().getEmail();
    return email || '(empty — expected when run from editor, works in web app)';
  });

  // 9. ScriptApp URL
  step('ScriptApp.getService().getUrl()', function() {
    var url = ScriptApp.getService().getUrl();
    if (!url) throw new Error('null — web app may not be deployed');
    return url;
  });

  // 10. HTML template evaluation (auth view — lightest)
  step('Template: auth view', function() {
    var template = HtmlService.createTemplateFromFile('index');
    template.view = 'auth';
    template.pageData = JSON.stringify({ view: 'auth', config: _sanitizeConfig(config || { orgName: 'Test', orgAbbrev: '', logoInitials: '', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' }), error: null, webAppUrl: '', tokenChecked: false });
    var output = template.evaluate();
    var size = output.getContent().length;
    return 'OK (' + Math.round(size / 1024) + ' KB)';
  });

  // 11. HTML template evaluation (dashboard view — heaviest)
  step('Template: dashboard view (steward+member)', function() {
    var template = HtmlService.createTemplateFromFile('index');
    template.view = 'steward';
    template.pageData = JSON.stringify({ view: 'steward', config: _sanitizeConfig(config || { orgName: 'Test', orgAbbrev: '', logoInitials: '', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' }), user: { email: 'test@test.com', name: 'Test', firstName: 'Test', lastName: 'User', role: 'steward', unit: 'Test', joined: '', duesStatus: 'Active', duesPaying: null, phone: '', workLocation: '', officeDays: '', assignedSteward: '', hasOpenGrievance: false, sharePhone: false }, isDualRole: true, sessionToken: null, initialTab: null, webAppUrl: '' });
    var output = template.evaluate();
    var size = output.getContent().length;
    if (size > 500000) {
      throw new Error('HTML output is ' + Math.round(size / 1024) + ' KB — may exceed GAS limit (500 KB). Consider lazy-loading views.');
    }
    return 'OK (' + Math.round(size / 1024) + ' KB)';
  });

  // 12. ScriptProperties accessible
  step('PropertiesService.getScriptProperties()', function() {
    var props = PropertiesService.getScriptProperties();
    if (!props) throw new Error('null');
    var keys = props.getKeys();
    return 'OK (' + keys.length + ' keys)';
  });

  // 13. CacheService accessible
  step('CacheService.getScriptCache()', function() {
    var cache = CacheService.getScriptCache();
    if (!cache) throw new Error('null');
    return 'OK';
  });

  Logger.log('\n=== DIAGNOSIS COMPLETE (' + (Date.now() - start) + 'ms total) ===');
  var failed = results.filter(function(r) { return !r.ok; });
  if (failed.length === 0) {
    Logger.log('All steps passed. If the web app still fails to load:');
    Logger.log('  1. Check that you created a NEW deployment version after clasp push');
    Logger.log('  2. Clear browser cache and try incognito window');
    Logger.log('  3. Check Apps Script editor Executions log for doGet errors');
    Logger.log('  4. If HTML output > 450 KB, the page may be too large for GAS');
  } else {
    Logger.log(failed.length + ' step(s) failed:');
    for (var i = 0; i < failed.length; i++) {
      Logger.log('  ✗ ' + failed[i].step + ': ' + failed[i].error);
    }
  }
  return results;
}
