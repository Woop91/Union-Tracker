/**
 * WebApp.gs — Main entry point
 * doGet(e) → Auth → Role → Serve HTML
 * 
 * Role logic:
 *   "Is Steward" = Yes → role is "both" (steward + member)
 *   "Is Steward" = No  → role is "member"
 *   Dual-role users default to steward view with member toggle.
 */

function doGet(e) {
  e = e || { parameter: {} };
  
  try {
    var config = ConfigReader.getSafeConfig();
    var user = Auth.resolveUser(e);
    
    if (!user) {
      return _serveAuth(config, e);
    }
    
    var userRecord = DataService.findUserByEmail(user.email);
    if (!userRecord) {
      return _serveError(config, 'not_found', user.email);
    }
    
    // Create session if "remember me" and came via magic link
    var sessionToken = null;
    if (e.parameter.remember === '1' && user.method === 'magic') {
      sessionToken = Auth.createSessionToken(user.email);
    }
    
    // Role: getUserRole returns 'both' for stewards, 'member' for members
    var role = DataService.getUserRole(user.email);
    
    return _serveDashboard(config, userRecord, role, sessionToken);
    
  } catch (err) {
    Logger.log('WebApp doGet error: ' + err.message + '\n' + err.stack);
    var fallbackConfig;
    try { fallbackConfig = ConfigReader.getSafeConfig(); } catch (e2) {
      fallbackConfig = { orgName: 'Dashboard', accentHue: 250, logoInitials: 'D' };
    }
    return _serveError(fallbackConfig, 'error', err.message);
  }
}

function _serveAuth(config, e) {
  var t = HtmlService.createTemplateFromFile('index');
  t.pageData = JSON.stringify({
    view: 'auth', config: config,
    error: e.parameter.authError || null,
  });
  return t.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function _serveDashboard(config, userRecord, role, sessionToken) {
  var t = HtmlService.createTemplateFromFile('index');
  t.pageData = JSON.stringify({
    view: (role === 'both') ? 'steward' : 'member', // both defaults to steward
    config: config,
    user: {
      email: userRecord.email, name: userRecord.name,
      firstName: userRecord.firstName, lastName: userRecord.lastName,
      role: role, unit: userRecord.unit, hireDate: userRecord.hireDate,
      phone: userRecord.phone, jobTitle: userRecord.jobTitle,
      isSteward: userRecord.isSteward, workLocation: userRecord.workLocation,
      assignedSteward: userRecord.assignedSteward,
    },
    isDualRole: role === 'both',
    sessionToken: sessionToken || null,
  });
  return t.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function _serveError(config, type, detail) {
  var t = HtmlService.createTemplateFromFile('index');
  var msgs = {
    not_found: 'Your email was not found in the member directory. Contact your steward for access.',
    expired: 'Your sign-in link has expired. Please request a new one.',
    error: 'Something went wrong. Please try again.',
  };
  t.pageData = JSON.stringify({
    view: 'error', config: config,
    error: { type: type, message: msgs[type] || msgs.error, detail: type === 'error' ? detail : null },
  });
  return t.evaluate()
    .setTitle(config.orgName + ' Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
