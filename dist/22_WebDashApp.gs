/**
 * WebApp.gs
 * Main entry point for the web app dashboard.
 * 
 * doGetWebDashboard(e) handles:
 *   1. Auth resolution (SSO / magic link / session token)
 *   2. Role lookup from Member Directory
 *   3. Routing to the correct view (auth, steward, member)
 *   4. Injecting config + user data into the HTML template
 * 
 * Deployment: Deploy as Web App
 *   - Execute as: Me (the script owner)
 *   - Who has access: Anyone (or anyone within org)
 *   - Note: Users arrive via Bitly redirect, not the raw URL
 */

/**
 * Main GET handler — serves the single-page app.
 * @param {Object} e - Event object with URL parameters
 * @returns {HtmlOutput}
 */
function doGetWebDashboard(e) {
  e = e || { parameter: {} };
  
  try {
    var config = ConfigReader.getConfig();
    var user = Auth.resolveUser(e);
    
    if (!user) {
      // Not authenticated — show login screen
      return _serveAuth(config, e);
    }
    
    // Authenticated — look up role
    var userRecord = DataService.findUserByEmail(user.email);
    
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
    
    return _serveDashboard(config, userRecord, role, sessionToken);
    
  } catch (err) {
    Logger.log('WebApp doGet error: ' + err.message + '\n' + err.stack);
    return _serveError(ConfigReader.getConfig(), 'error', err.message);
  }
}

/**
 * Serves the login/auth page.
 */
function _serveAuth(config, e) {
  var template = HtmlService.createTemplateFromFile('index');
  
  template.pageData = JSON.stringify({
    view: 'auth',
    config: _sanitizeConfig(config),
    error: e.parameter.authError || null,
  });
  
  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Serves the dashboard (steward or member view).
 */
function _serveDashboard(config, userRecord, role, sessionToken) {
  var template = HtmlService.createTemplateFromFile('index');
  
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
  };
  
  template.pageData = JSON.stringify({
    view: role === 'steward' || role === 'both' ? 'steward' : 'member',
    config: _sanitizeConfig(config),
    user: safeUser,
    isDualRole: role === 'both',
    sessionToken: sessionToken || null,
  });
  
  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Serves an error page.
 */
function _serveError(config, type, detail) {
  var template = HtmlService.createTemplateFromFile('index');
  
  var messages = {
    'not_found': 'Your email was not found in the member directory. Please contact your steward for access.',
    'expired': 'Your sign-in link has expired. Please request a new one.',
    'error': 'Something went wrong. Please try again.',
  };
  
  template.pageData = JSON.stringify({
    view: 'error',
    config: _sanitizeConfig(config),
    error: {
      type: type,
      message: messages[type] || messages['error'],
      detail: type === 'error' ? detail : null,
    },
  });
  
  return template.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
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
  };
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
