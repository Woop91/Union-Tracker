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
    Logger.log('doGet FATAL: ' + fatalErr.message + '\n' + fatalErr.stack);
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

  try {
    var config = ConfigReader.getConfig();
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
    var sessionToken = null;
    if (e.parameter.remember === '1' && user.method === 'magic') {
      sessionToken = Auth.createSessionToken(user.email);
    }

    // Route to appropriate dashboard
    var role = userRecord.role; // 'steward', 'member', or 'both'
    var initialTab = e.parameter.page || null;

    return _serveDashboard(config, userRecord, role, sessionToken, initialTab);

  } catch (err) {
    Logger.log('WebApp doGet error: ' + err.message + '\n' + err.stack);
    // Safe config fetch — if ConfigReader itself threw, fall back to defaults
    // so _serveError can still render a page.
    var safeConfig;
    try {
      safeConfig = ConfigReader.getConfig();
    } catch (_cfgErr) {
      safeConfig = { orgName: 'Dashboard', orgAbbrev: '', logoInitials: '', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' };
    }
    return _serveError(safeConfig, 'error', err.message);
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
    webAppUrl: ScriptApp.getService().getUrl(),
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
    phone: userRecord.phone,
    workLocation: userRecord.workLocation || '',
    officeDays: userRecord.officeDays || '',
    assignedSteward: userRecord.assignedSteward || '',
    hasOpenGrievance: userRecord.hasOpenGrievance || false,
  };

  template.pageData = JSON.stringify({
    view: role === 'steward' || role === 'both' ? 'steward' : 'member',
    config: _sanitizeConfig(config),
    user: safeUser,
    isDualRole: role === 'both',
    sessionToken: sessionToken || null,
    initialTab: initialTab || null,
    webAppUrl: ScriptApp.getService().getUrl(),
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
    driveFolderUrl: config.driveFolderId ? 'https://drive.google.com/drive/folders/' + config.driveFolderId : '',
    surveyFormUrl: config.satisfactionFormUrl || '',
    orgWebsite: config.orgWebsite || '',
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
 * @returns {string} Raw HTML content (CSS-scoped under .oc-wrap)
 */
function getOrgChartHtml() {
  return HtmlService.createHtmlOutputFromFile('org_chart').getContent();
}

/**
 * Returns the published web app URL. Used by client-side logout
 * as a reload fallback when window.top.location.reload() is blocked.
 * @returns {string}
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
