/**
 * WebApp.gs — Main entry point for the web app SPA
 *
 * WHAT THIS FILE DOES:
 *   Main entry point for the web app SPA. doGet(e) is THE function Google
 *   Apps Script calls when a user visits the deployed web app URL. It handles:
 *     (1) Auth resolution via Auth.resolveUser(e) — tries SSO/magic link/
 *         session token
 *     (2) Role lookup from Member Directory (admin/steward/member)
 *     (3) Route to correct view (auth_view for login, steward_view for
 *         stewards, member_view for members)
 *     (4) Config + user data injection into HTML template
 *   Deep-link support: ?page=<tabId> opens specific SPA tab after auth.
 *   Token-authenticated pages: ?page=esign (e-sig), ?page=rsvp (meeting RSVP).
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   "Execute as: Me" deployment means the script runs with the owner's
 *   permissions, allowing access to all sheets. "Who has access: Anyone" is
 *   required so members can use the web app. Users arrive via Bitly redirects,
 *   not the raw Apps Script URL. The try/catch in doGet wraps
 *   doGetWebDashboard and falls back to _serveFatalError() — this ensures
 *   users always see a page, even if the SPA fails to load.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   The entire web app is down. Users see a fatal error page or a Google
 *   error. Nobody can access the SPA dashboard. If routing breaks, stewards
 *   might see the member view (less data than expected) or members might see
 *   the steward view (PII exposure — security issue).
 *
 * DEPENDENCIES:
 *   Depends on: 19_WebDashAuth.gs (Auth.resolveUser),
 *               20_WebDashConfigReader.gs (ConfigReader.getConfig),
 *               21_WebDashDataService.gs (DataService),
 *               HtmlService (GAS built-in).
 *   Used by: All HTML template files (index.html, steward_view.html,
 *            member_view.html, auth_view.html, error_view.html).
 */

/**
 * GAS web app entry point — serves the single-page app.
 * @param {Object} e - Event object with URL parameters
 * @returns {HtmlOutput}
 */
function doGet(e) {
  e = e || { parameter: {} };

  // v4.33.0 — E-Signature page (token-authenticated, no login required)
  if (e.parameter && e.parameter.page === 'esign') {
    try {
      return HtmlService.createHtmlOutputFromFile('esign')
        .setTitle('Grievance E-Signature — ' + (function() { try { return getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Your Local'; } catch(_) { return 'Your Local'; } })());
    } catch (esignErr) {
      Logger.log('doGet esign error: ' + esignErr.message);
      return _serveFatalError('E-Signature page unavailable.');
    }
  }

  // v4.36.0 — RSVP page (token-authenticated, no login required)
  if (e.parameter && e.parameter.page === 'rsvp') {
    try {
      return _serveRSVPPage(e);
    } catch (rsvpErr) {
      Logger.log('doGet rsvp error: ' + rsvpErr.message);
      return _serveFatalError('RSVP page unavailable.');
    }
  }

  // v4.43.0 — QR Code meeting check-in (mobile, no login required)
  if (e.parameter && e.parameter.page === 'qr-checkin') {
    try {
      return _serveQRCheckInPage(e);
    } catch (qrErr) {
      Logger.log('doGet qr-checkin error: ' + qrErr.message);
      return _serveFatalError('QR Check-In page unavailable.');
    }
  }

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
    } catch (_) { Logger.log('_: ' + (_.message || _)); }
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
    config = { orgName: 'SolidBase', orgAbbrev: 'SB', logoInitials: 'SB', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' };
  }

  var _doGetStart = Date.now();

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
      // v4.40.0: Magic link expired or already used — show helpful message
      // instead of silently dumping users back to login with no explanation
      if (e.parameter.token) {
        return _serveAuth(config, e, 'token_expired');
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
          // H5: Audit log bootstrap admin access
          if (typeof recordSecurityEvent === 'function') {
            recordSecurityEvent('BOOTSTRAP_ADMIN', 'MEDIUM', 'Script owner granted admin access without directory entry', { email: user.email });
          }
        }
      } catch (_ownerErr) { Logger.log('_ownerErr: ' + (_ownerErr.message || _ownerErr)); }
    }

    if (!userRecord) {
      // Email not in directory
      return _serveError(config, 'not_found', user.email);
    }

    // Handle "remember me" — create session token if requested
    // Also echo back an existing validated session token so the client always has
    // SESSION_TOKEN populated for non-SSO auth (magic link + remember me, or returning
    // session users). This is safe: the token was already validated by resolveUser().
    // v4.40.0: SSO users are auto-remembered — Google auth implies browser trust,
    // and losing the session on accidental tab close was a top user pain point.
    var sessionToken = null;
    if (user.method === 'sso') {
      // Auto-remember SSO users — no toggle needed; Google auth is already trusted
      var ssoToken = Auth.createSessionToken(user.email, 'sso');
      if (ssoToken && typeof ssoToken === 'object' && ssoToken.error) {
        Logger.log('SSO auto-session storage failed: ' + ssoToken.message);
      } else {
        sessionToken = ssoToken;
      }
    } else if (e.parameter.remember === '1' && user.method === 'magic') {
      var tokenResult = Auth.createSessionToken(user.email);
      // C3: Handle session storage failure — createSessionToken may return error object
      if (tokenResult && typeof tokenResult === 'object' && tokenResult.error) {
        Logger.log('Session token storage failed: ' + tokenResult.message);
        // Proceed without remember-me; user will need to re-authenticate next visit
      } else {
        sessionToken = tokenResult;
      }
    } else if ((user.method === 'session' || user.method === 'pin') && e.parameter.sessionToken) {
      // Echo back the already-validated session token so the client can use it
      // PIN sessions use the same echo-back — token was created during devAuthLoginByPIN()
      sessionToken = e.parameter.sessionToken;
    }

    // Route to appropriate dashboard
    var role = userRecord.role; // 'steward', 'member', or 'both'
    var initialTab = e.parameter.page || null;

    var elapsed = Date.now() - _doGetStart;
    if (elapsed > 20000) {
      Logger.log('doGet slow: ' + elapsed + 'ms before serving dashboard');
    }

    return _serveDashboard(config, userRecord, role, sessionToken, initialTab, user.method);

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
    isDevMode: !isProductionMode(),
  });

  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Serves the dashboard (steward or member view).
 * @param {string} [authMethod] - Auth method ('sso', 'magic', 'session', 'pin')
 */
function _serveDashboard(config, userRecord, role, sessionToken, initialTab, authMethod) {
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

  // Fetch user's color theme for unified sheet↔webapp theming
  var colorThemeData = {};
  try { colorThemeData = getUserColorTheme(); } catch (_e) { colorThemeData = { themeKey: 'default', accentHue: 250 }; }

  // Read default view preference for dual-role users (ScriptProperties, email-keyed).
  // Uses ScriptProperties (not UserProperties) because UserProperties returns the
  // script owner's props in Execute-as-Me webapps — shared, not per-user.
  var defaultView = 'steward';
  if (role === 'both') {
    try {
      var pref = PropertiesService.getScriptProperties().getProperty('defaultView_' + userRecord.email.toLowerCase());
      if (pref === 'member') defaultView = 'member';
    } catch (_) {}
  }

  template.pageData = JSON.stringify({
    view: role === 'steward' || role === 'both' ? 'steward' : 'member',
    config: _sanitizeConfig(config),
    user: safeUser,
    isDualRole: role === 'steward' || role === 'both',
    defaultView: defaultView,
    sessionToken: sessionToken || null,
    initialTab: initialTab || null,
    webAppUrl: _getWebAppUrlSafe(),
    colorTheme: colorThemeData.themeKey || 'default',
    colorThemes: (typeof getColorThemeList === 'function') ? getColorThemeList() : [],
    isDevMode: !isProductionMode(),
    authMethod: authMethod || 'sso',
    isAdmin: (function() {
      try {
        if (typeof _adminIsAuthorized_ === 'function') return _adminIsAuthorized_(userRecord.email);
        return userRecord.email.toLowerCase() === Session.getEffectiveUser().getEmail().toLowerCase();
      } catch (_) { return false; }
    })(),
    appVersion: (typeof VERSION_INFO !== 'undefined' && VERSION_INFO.version) ? VERSION_INFO.version : '',
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
  // Override accentHue with user's saved color theme preference if set
  var userHue = config.accentHue;
  try {
    var savedHue = PropertiesService.getUserProperties().getProperty('visual_accentHue');
    if (savedHue) {
      var parsed = JSON.parse(savedHue);
      if (typeof parsed === 'number' && parsed >= 0 && parsed <= 360) userHue = parsed;
    }
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  return {
    orgName: config.orgName,
    orgAbbrev: config.orgAbbrev,
    logoInitials: config.logoInitials,
    accentHue: userHue,
    stewardLabel: config.stewardLabel,
    memberLabel: config.memberLabel,
    magicLinkExpiryDays: config.magicLinkExpiryDays,
    cookieDurationDays: config.cookieDurationDays,
    calendarUrl: config.calendarId ? 'https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(config.calendarId) : '',
    // H8: Pre-build URLs server-side instead of exposing raw resource IDs
    calendarCreateUrl: config.calendarId ? 'https://calendar.google.com/calendar/render?action=TEMPLATE&src=' + encodeURIComponent(config.calendarId) : '',
    driveFolderUrl: config.driveFolderId ? 'https://drive.google.com/drive/folders/' + config.driveFolderId : '',
    // surveyFormUrl removed v4.22.7 — survey is native webapp (renderSurveyFormPage in member_view.html)
    orgWebsite: config.orgWebsite || '',
    broadcastScopeAll: (String(config.broadcastScopeAll || '').trim().toLowerCase() === 'yes'),
    // H8: v4.31.0 — replaced raw folder IDs with pre-built URLs and boolean flags
    minutesFolderUrl: config.minutesFolderId ? 'https://drive.google.com/drive/folders/' + config.minutesFolderId : '',
    hasMinutesFolder: !!config.minutesFolderId,
    hasGrievancesFolder: !!config.grievancesFolderId,
    // v4.20.18: insights cache TTL exposed so client can show staleness info
    insightsCacheTTLMin: config.insightsCacheTTLMin || 5,
    issueCategories: (typeof DEFAULT_CONFIG !== 'undefined' && Array.isArray(DEFAULT_CONFIG.ISSUE_CATEGORY)) ? DEFAULT_CONFIG.ISSUE_CATEGORY : [],
    showGrievances: _isGrievancesEnabled(),
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
    + 'body.light{background:#f2f2f5;color:#1a1a2e}'
    + '.card{max-width:440px;padding:40px 32px;background:#1e1e1e;border-radius:16px;text-align:center}'
    + 'body.light .card{background:#ffffff;box-shadow:0 2px 12px rgba(0,0,0,0.08)}'
    + 'h1{font-size:20px;margin:0 0 12px}p{color:#888;font-size:14px;line-height:1.6;margin:0 0 20px}'
    + 'body.light p{color:#5c5c7a}'
    + 'a{display:inline-block;padding:10px 24px;background:hsl(250,70%,68%);color:#fff;'
    + 'border-radius:8px;text-decoration:none;font-weight:600}'
    + '</style></head><body><div class="card">'
    + '<h1>Something went wrong</h1>'
    + '<p>The dashboard could not load. This is usually temporary &mdash; '
    + 'please wait a moment and try again.</p>'
    + '<a href="javascript:location.reload()">Reload</a>'
    + '</div><script>try{if(localStorage.getItem("dds_isDark")==="false")document.body.classList.add("light")}catch(e){}</script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Dashboard — Error')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Serves the RSVP confirmation page — token-authenticated, no login required.
 * Processes the RSVP response and shows a confirmation/error page.
 * @param {Object} e - Event object with token and response params
 * @returns {HtmlOutput}
 * @private
 */
function _serveRSVPPage(e) {
  var token = e.parameter.token || '';
  var response = e.parameter.response || '';

  if (!token || (response !== 'accept' && response !== 'decline')) {
    return _serveFatalError('Invalid RSVP link.');
  }

  var result = RSVPService.processRSVP(token, response);

  var orgName = '';
  try {
    var cfg = ConfigReader.getConfig();
    orgName = cfg.orgName || '';
  } catch (_) { /* ignore */ }

  var statusText = result.success
    ? (response === 'accept' ? 'Accepted' : 'Declined')
    : 'Error';
  var statusColor = response === 'accept' ? '#22c55e' : '#ef4444';
  var icon = result.success ? (response === 'accept' ? '\u2705' : '\u274C') : '\u26A0\uFE0F';

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Meeting RSVP' + (orgName ? ' — ' + escapeHtml(orgName) : '') + '</title>'
    + '<style>'
    + 'body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;'
    + 'font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#141414;color:#e0e0e0}'
    + 'body.light{background:#f2f2f5;color:#1a1a2e}'
    + '.card{max-width:440px;padding:40px 32px;background:#1e1e1e;border-radius:16px;text-align:center}'
    + 'body.light .card{background:#ffffff;box-shadow:0 2px 12px rgba(0,0,0,0.08)}'
    + '.icon{font-size:48px;margin-bottom:16px}'
    + 'h1{font-size:20px;margin:0 0 8px}'
    + '.meeting-name{font-size:16px;font-weight:600;margin:0 0 12px;color:' + statusColor + '}'
    + 'p{color:#888;font-size:14px;line-height:1.6;margin:0 0 12px}'
    + 'body.light p{color:#5c5c7a}'
    + '.status{display:inline-block;padding:6px 16px;border-radius:20px;font-weight:600;font-size:13px;'
    + 'background:' + statusColor + '22;color:' + statusColor + '}'
    + '</style></head><body><div class="card">'
    + '<div class="icon">' + icon + '</div>';

  if (result.success) {
    html += '<h1>RSVP ' + statusText + '</h1>'
      + '<div class="meeting-name">' + escapeHtml(result.meetingName || '') + '</div>'
      + '<p>' + (result.memberName ? escapeHtml(result.memberName) + ', your' : 'Your')
      + ' response has been recorded.</p>'
      + '<div class="status">' + escapeHtml(statusText) + '</div>';
  } else {
    html += '<h1>RSVP Link Expired</h1>'
      + '<p>' + escapeHtml(result.error || 'This link is no longer valid.') + '</p>'
      + '<p>Please contact your steward for a new invitation.</p>';
  }

  html += '</div><script>try{if(localStorage.getItem("dds_isDark")==="false")document.body.classList.add("light")}catch(e){}</script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Meeting RSVP' + (orgName ? ' — ' + escapeHtml(orgName) : ''))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Serve the QR Code mobile check-in page.
 * Members scan a QR code → land here → enter phone + PIN → attendance recorded.
 * No login/session required — the page authenticates via phone + PIN per check-in.
 *
 * @param {Object} e - Event object with e.parameter.meeting
 * @returns {HtmlOutput}
 * @private
 */
function _serveQRCheckInPage(e) {
  var meetingId = e.parameter.meeting || '';

  var orgName = '';
  try {
    var cfg = ConfigReader.getConfig();
    orgName = cfg.orgName || '';
  } catch (_) { /* ignore */ }

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">'
    + '<title>Meeting Check-In' + (orgName ? ' — ' + escapeHtml(orgName) : '') + '</title>'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;'
    + 'background:#141414;color:#e0e0e0;min-height:100vh;display:flex;flex-direction:column}'
    + 'body.light{background:#f2f2f5;color:#1a1a2e}'

    // Header
    + '.header{background:linear-gradient(135deg,#059669,#047857);color:white;padding:24px 20px;text-align:center}'
    + '.header h1{font-size:22px;margin-bottom:4px}'
    + '.header .subtitle{font-size:13px;opacity:0.9}'

    // Container
    + '.container{flex:1;max-width:440px;width:100%;margin:0 auto;padding:20px}'

    // Card
    + '.card{background:#1e1e1e;border-radius:16px;padding:28px 24px;margin-bottom:16px}'
    + 'body.light .card{background:#ffffff;box-shadow:0 2px 12px rgba(0,0,0,0.08)}'

    // Meeting info
    + '.meeting-badge{display:inline-block;padding:6px 14px;border-radius:20px;font-size:12px;'
    + 'font-weight:600;background:#059669;color:white;margin-bottom:16px}'
    + '.meeting-name{font-size:18px;font-weight:700;color:#fff;margin-bottom:4px;text-align:center}'
    + 'body.light .meeting-name{color:#1a1a2e}'
    + '.meeting-detail{font-size:13px;color:#888;text-align:center;margin-bottom:8px}'
    + 'body.light .meeting-detail{color:#5c5c7a}'

    // Fields
    + '.field{margin-bottom:18px}'
    + '.field label{display:block;margin-bottom:8px;font-weight:600;color:#ccc;font-size:14px}'
    + 'body.light .field label{color:#333}'
    + '.field input{width:100%;padding:14px 16px;border:2px solid #333;border-radius:10px;'
    + 'font-size:17px;background:#252525;color:#fff;transition:border-color 0.2s;-webkit-appearance:none}'
    + 'body.light .field input{background:#f8f8f8;border-color:#e0e0e0;color:#1a1a2e}'
    + '.field input:focus{outline:none;border-color:#059669}'
    + '.field .hint{font-size:12px;color:#666;margin-top:4px}'

    // Button
    + '.btn{width:100%;padding:16px;border:none;border-radius:10px;font-size:18px;font-weight:700;'
    + 'cursor:pointer;background:#059669;color:white;transition:all 0.2s;-webkit-appearance:none}'
    + '.btn:hover{background:#047857}'
    + '.btn:active{transform:scale(0.98)}'
    + '.btn:disabled{background:#444;cursor:not-allowed}'
    + 'body.light .btn:disabled{background:#ccc}'

    // Messages
    + '.error{color:#f87171;font-size:14px;margin-top:12px;text-align:center;padding:14px;'
    + 'background:#3b1010;border-radius:10px;display:none}'
    + 'body.light .error{background:#fee2e2;color:#dc2626}'
    + '.success-banner{text-align:center;padding:30px 20px;background:#052e1c;border-radius:16px;'
    + 'margin-bottom:16px;display:none}'
    + 'body.light .success-banner{background:#d1fae5}'
    + '.success-banner .checkmark{font-size:56px;margin-bottom:12px}'
    + '.success-banner .name{font-size:22px;font-weight:700;color:#34d399;margin-bottom:6px}'
    + 'body.light .success-banner .name{color:#059669}'
    + '.success-banner .msg{color:#6ee7b7;font-size:14px}'
    + 'body.light .success-banner .msg{color:#047857}'

    // Attendee count
    + '.attendee-count{text-align:center;padding:10px;background:#0a2e1c;border-radius:8px;'
    + 'color:#34d399;font-weight:600;font-size:14px;margin-top:12px}'
    + 'body.light .attendee-count{background:#f0fdf4;color:#059669}'

    // Loading
    + '.loading{text-align:center;padding:40px;color:#888}'
    + '.loading .spinner{display:inline-block;width:32px;height:32px;border:3px solid #333;'
    + 'border-top-color:#059669;border-radius:50%;animation:spin 0.8s linear infinite}'
    + '@keyframes spin{to{transform:rotate(360deg)}}'

    // No meeting
    + '.no-meeting{text-align:center;padding:40px 20px}'
    + '.no-meeting .icon{font-size:48px;margin-bottom:16px}'
    + '.no-meeting h2{font-size:18px;margin-bottom:8px;color:#e0e0e0}'
    + 'body.light .no-meeting h2{color:#1a1a2e}'
    + '.no-meeting p{color:#888;font-size:14px;line-height:1.6}'
    + 'body.light .no-meeting p{color:#5c5c7a}'

    + '</style></head><body>'

    + '<div class="header">'
    + '<h1>' + escapeHtml(orgName || 'Meeting Check-In') + '</h1>'
    + '<div class="subtitle">Scan &bull; Enter Phone &amp; PIN &bull; Check In</div>'
    + '</div>'

    + '<div class="container">'

    // Success banner
    + '<div id="successBanner" class="success-banner">'
    + '<div class="checkmark">&#10003;</div>'
    + '<div class="name" id="successName"></div>'
    + '<div class="msg">Checked in successfully!</div>'
    + '</div>'

    // Loading state
    + '<div id="loadingState" class="card loading">'
    + '<div class="spinner"></div>'
    + '<p style="margin-top:12px">Loading meeting...</p>'
    + '</div>'

    // Meeting info + form (hidden until loaded)
    + '<div id="meetingCard" class="card" style="display:none">'
    + '<div style="text-align:center"><span class="meeting-badge" id="meetingType"></span></div>'
    + '<div class="meeting-name" id="meetingName"></div>'
    + '<div class="meeting-detail" id="meetingDetail"></div>'
    + '</div>'

    + '<div id="checkinForm" class="card" style="display:none">'
    + '<div class="field">'
    + '<label for="phone">Phone Number</label>'
    + '<input type="tel" id="phone" placeholder="(555) 123-4567" autocomplete="tel" inputmode="tel">'
    + '</div>'
    + '<div class="field">'
    + '<label for="pin">PIN</label>'
    + '<input type="password" id="pin" placeholder="Enter your PIN" maxlength="6" pattern="[0-9]*" inputmode="numeric" autocomplete="off">'
    + '</div>'
    + '<button class="btn" id="checkinBtn" onclick="doCheckIn()">Check In</button>'
    + '<div id="error" class="error"></div>'
    + '<div id="attendeeCount" class="attendee-count" style="display:none"></div>'
    + '</div>'

    // No meeting state
    + '<div id="noMeeting" class="card no-meeting" style="display:none">'
    + '<div class="icon">&#128197;</div>'
    + '<h2>Meeting Not Found</h2>'
    + '<p>This check-in link may have expired or the meeting has ended. Please ask your steward for a new QR code.</p>'
    + '</div>'

    + '</div>'  // container

    + '<script>'
    + 'var MEETING_ID=' + JSON.stringify(meetingId) + ';'

    // Dark mode detection
    + 'try{if(localStorage.getItem("dds_isDark")==="false")document.body.classList.add("light")}catch(e){}'

    // On load: validate meeting
    + 'if(!MEETING_ID){'
    + '  document.getElementById("loadingState").style.display="none";'
    + '  document.getElementById("noMeeting").style.display="block";'
    + '}else{'
    + '  google.script.run'
    + '    .withSuccessHandler(function(r){'
    + '      document.getElementById("loadingState").style.display="none";'
    + '      if(!r.success||r.meetings.length===0){'
    + '        document.getElementById("noMeeting").style.display="block";'
    + '        return;'
    + '      }'
    + '      var found=null;'
    + '      r.meetings.forEach(function(m){if(m.id===MEETING_ID)found=m});'
    + '      if(!found){'
    + '        document.getElementById("noMeeting").style.display="block";'
    + '        return;'
    + '      }'
    + '      document.getElementById("meetingName").textContent=found.name;'
    + '      document.getElementById("meetingType").textContent=found.type;'
    + '      document.getElementById("meetingDetail").textContent=found.date+(found.time?" at "+found.time:"");'
    + '      document.getElementById("meetingCard").style.display="block";'
    + '      document.getElementById("checkinForm").style.display="block";'
    + '      document.getElementById("phone").focus();'
    + '      refreshCount();'
    + '    })'
    + '    .withFailureHandler(function(){'
    + '      document.getElementById("loadingState").style.display="none";'
    + '      document.getElementById("noMeeting").style.display="block";'
    + '    })'
    + '    .getCheckInEligibleMeetings();'
    + '}'

    // Attendee count refresh
    + 'function refreshCount(){'
    + '  google.script.run'
    + '    .withSuccessHandler(function(r){'
    + '      if(r.success&&r.count>0){'
    + '        var el=document.getElementById("attendeeCount");'
    + '        el.textContent=r.count+" member"+(r.count!==1?"s":"")+" checked in";'
    + '        el.style.display="block";'
    + '      }'
    + '    })'
    + '    .getMeetingAttendees(MEETING_ID);'
    + '}'

    // Check-in handler
    + 'function doCheckIn(){'
    + '  var phone=document.getElementById("phone").value.trim();'
    + '  var pin=document.getElementById("pin").value.trim();'
    + '  var errEl=document.getElementById("error");'
    + '  errEl.style.display="none";'
    + '  if(!phone){showErr("Please enter your phone number");return}'
    + '  if(!pin||pin.length<4){showErr("Please enter your PIN");return}'
    + '  var btn=document.getElementById("checkinBtn");'
    + '  btn.disabled=true;btn.textContent="Checking in...";'
    + '  google.script.run'
    + '    .withSuccessHandler(function(r){'
    + '      btn.disabled=false;btn.textContent="Check In";'
    + '      if(r.success){'
    + '        document.getElementById("successName").textContent=r.memberName;'
    + '        document.getElementById("successBanner").style.display="block";'
    + '        document.getElementById("phone").value="";'
    + '        document.getElementById("pin").value="";'
    + '        errEl.style.display="none";'
    + '        refreshCount();'
    + '        setTimeout(function(){'
    + '          document.getElementById("successBanner").style.display="none";'
    + '          document.getElementById("phone").focus();'
    + '        },3500);'
    + '      }else{showErr(r.error)}'
    + '    })'
    + '    .withFailureHandler(function(e){'
    + '      btn.disabled=false;btn.textContent="Check In";'
    + '      showErr("Error: "+e.message);'
    + '    })'
    + '    .processQRCheckIn(MEETING_ID,phone,pin);'
    + '}'

    + 'function showErr(msg){'
    + '  var el=document.getElementById("error");'
    + '  el.textContent=msg;el.style.display="block";'
    + '}'

    // Enter key handlers
    + 'document.getElementById("phone").addEventListener("keypress",function(e){if(e.key==="Enter")document.getElementById("pin").focus()});'
    + 'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter")doCheckIn()});'

    + '</script></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Meeting Check-In' + (orgName ? ' — ' + escapeHtml(orgName) : ''))
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
 * Client-callable: Returns the member view HTML for lazy-loading.
 * For dual-role users (steward + member), member_view.html is NOT included
 * in the initial template to stay under the GAS ~820KB HtmlOutput limit.
 * Loaded on-demand when the user switches to member view.
 * @returns {string} Raw HTML content (<script> block defining member view functions)
 */
function getMemberViewHtml() {
  try {
    // No auth check needed — user already authenticated via doGet().
    // Session.getActiveUser().getEmail() returns empty for magic-link/session-token
    // users (Execute-as-Me), which was causing this to return '' and silently
    // breaking the member view switch for dual-role users.
    // No CacheService — member_view.html exceeds the 100KB per-key limit.
    return HtmlService.createHtmlOutputFromFile('member_view').getContent();
  } catch (e) {
    Logger.log('getMemberViewHtml error: ' + e.message);
    return '';
  }
}

/**
 * Client-callable: Returns the org chart HTML content for lazy-loading.
 * Loaded on-demand when the user navigates to the Org Chart tab.
 * @returns {string} Raw HTML content (CSS-scoped under .madds-embed), or error message
 */
function getOrgChartHtml() {
  try {
    // No auth check — user already authenticated via doGet().
    // Session.getActiveUser().getEmail() returns empty for magic-link users.
    // PERF: Cache static HTML in CacheService — avoids re-reading the file on every tab click.
    // Version-keyed so deploys automatically bust the cache.
    var ver = (typeof VERSION_INFO !== 'undefined' && VERSION_INFO.version) ? VERSION_INFO.version : '';
    var cacheKey = 'HTML_org_chart_' + ver;
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return cached;
    var html = HtmlService.createHtmlOutputFromFile('org_chart').getContent();
    try { cache.put(cacheKey, html, 21600); } catch (_) { /* exceeds 100KB limit — skip cache */ }
    return html;
  } catch (e) {
    Logger.log('getOrgChartHtml error: ' + e.message);
    return '<div class="empty-state">Org chart could not be loaded.</div>';
  }
}

/**
 * Client-callable: Returns the Agency Org Chart HTML for lazy-loading.
 * Loaded on-demand when the user navigates to the Agency Org Chart tab.
 * @returns {string} Raw HTML content (CSS-scoped under .agency-oc), or error message
 */
function getAgencyOrgChartHtml() {
  try {
    var ver = (typeof VERSION_INFO !== 'undefined' && VERSION_INFO.version) ? VERSION_INFO.version : '';
    var cacheKey = 'HTML_agency_org_chart_' + ver;
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return cached;
    var html = HtmlService.createHtmlOutputFromFile('agency_org_chart').getContent();
    try { cache.put(cacheKey, html, 21600); } catch (_) { /* exceeds 100KB limit — skip cache */ }
    return html;
  } catch (e) {
    Logger.log('getAgencyOrgChartHtml error: ' + e.message);
    return '<div class="empty-state">Agency org chart could not be loaded.</div>';
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
