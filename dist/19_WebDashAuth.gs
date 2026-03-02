/**
 * Auth.gs
 * Handles authentication for the web app dashboard.
 *
 * Two auth paths:
 *   1. Google SSO — Session.getActiveUser().getEmail()
 *   2. Magic Link — email-based token sent via MailApp
 *
 * Token storage: ScriptProperties (server-side)
 * Client persistence: localStorage token (acts as "cookie")
 *
 * Flow:
 *   doGet() → Auth.resolveUser(e) → returns { email, method } or null
 *   If null → serve login page
 *   If email → lookup role in directory → route to dashboard
 */

var Auth = (function () {

  var TOKEN_PREFIX = 'MAGIC_TOKEN_';
  var SESSION_PREFIX = 'SESSION_';

  /**
   * Attempts to resolve the current user from available auth signals.
   * Priority: (1) session token in URL, (2) SSO, (3) magic link token in URL
   * @param {Object} e - doGet event object
   * @returns {Object|null} { email: string, method: 'sso'|'magic'|'session' } or null
   */
  function resolveUser(e) {
    var params = e ? e.parameter || {} : {};

    // 0. Explicit logout — bypass all auth methods
    if (params.loggedout === '1') return null;

    // 1. Check for session token (returning user with "remember me")
    if (params.sessionToken) {
      var sessionEmail = _validateSessionToken(params.sessionToken);
      if (sessionEmail) {
        return { email: sessionEmail, method: 'session' };
      }
    }

    // 2. Check Google SSO
    try {
      var ssoUser = Session.getActiveUser().getEmail();
      if (ssoUser && ssoUser !== '') {
        return { email: ssoUser.toLowerCase(), method: 'sso' };
      }
    } catch (_err) {
      // SSO not available — continue
    }

    // 3. Check magic link token in URL
    if (params.token) {
      var magicEmail = _validateMagicToken(params.token);
      if (magicEmail) {
        return { email: magicEmail, method: 'magic' };
      }
    }

    return null;
  }

  /**
   * Generates a magic link and sends it to the given email.
   * Called from the client via google.script.run.
   * @param {string} email - The email to send the link to
   * @param {boolean} rememberMe - Whether to create a long-lived session
   * @returns {Object} { success: boolean, message: string }
   */
  function sendMagicLink(email, rememberMe) {
    email = String(email).trim().toLowerCase();

    // Rate limiting — max 3 magic links per email per 15 minutes
    var cache = CacheService.getScriptCache();
    var rateKey = 'MAGIC_RATE_' + email;
    var count = parseInt(cache.get(rateKey) || '0', 10);
    if (count >= 3) {
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }
    cache.put(rateKey, String(count + 1), 900);

    // Validate email exists in directory
    var userRecord = DataService.findUserByEmail(email);
    if (!userRecord) {
      // Don't reveal whether email exists — security best practice
      // Simulate processing time to prevent timing-based enumeration
      Utilities.sleep(500 + Math.floor(Math.random() * 500));
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }

    var config = ConfigReader.getConfig();
    var token = _generateMagicToken(email);
    var webAppUrl = ScriptApp.getService().getUrl();

    var linkParams = '?token=' + encodeURIComponent(token);
    if (rememberMe) {
      linkParams += '&remember=1';
    }

    var signInUrl = webAppUrl + linkParams;

    var subject = 'Sign in to ' + config.orgName + ' Dashboard';
    var htmlBody = _buildEmailHtml(config, signInUrl, email);

    // Check remaining email quota before attempting to send
    try {
      var remaining = MailApp.getRemainingDailyQuota();
      if (remaining <= 0) {
        Logger.log('Auth: Daily email quota exhausted');
        return { success: false, message: 'Email quota reached for today. Please use Google Sign-In or try again tomorrow.' };
      }
    } catch (_quotaErr) {
      // If quota check fails, proceed with send attempt anyway
    }

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        noReply: true,
      });

      return { success: true, message: 'Sign-in link sent to ' + email };
    } catch (err) {
      Logger.log('Auth: Failed to send magic link to ' + email + ': ' + err.message);
      var msg = 'Failed to send email. ';
      if (err.message && err.message.indexOf('quota') >= 0) {
        msg += 'Email quota exhausted. Please use Google Sign-In.';
      } else if (err.message && err.message.indexOf('invalid') >= 0) {
        msg += 'The email address may be invalid. Please check and try again.';
      } else {
        msg += 'Please try again or use Google Sign-In.';
      }
      return { success: false, message: msg };
    }
  }

  /**
   * Creates a session token for "remember me" functionality.
   * Called after successful auth if remember=1.
   * @param {string} email
   * @returns {string} Session token
   */
  function createSessionToken(email) {
    var config = ConfigReader.getConfig();
    var token = _generateToken();
    var expiry = Date.now() + config.cookieDurationMs;

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SESSION_PREFIX + token, JSON.stringify({
      email: email.toLowerCase(),
      expiry: expiry,
      created: Date.now(),
    }));

    return token;
  }

  /**
   * Invalidates a session token (logout).
   * @param {string} token
   */
  function invalidateSession(token) {
    if (!token) return;
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty(SESSION_PREFIX + token);
  }

  /**
   * Cleans up expired tokens from ScriptProperties.
   * Should be run periodically via a time-based trigger.
   */
  function cleanupExpiredTokens() {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var now = Date.now();
    var cleaned = 0;

    for (var key in all) {
      if (key.indexOf(TOKEN_PREFIX) === 0 || key.indexOf(SESSION_PREFIX) === 0) {
        try {
          var data = JSON.parse(all[key]);
          if (data.expiry && data.expiry < now) {
            props.deleteProperty(key);
            cleaned++;
          }
        } catch (_e) {
          // Malformed entry — delete it
          props.deleteProperty(key);
          cleaned++;
        }
      }
    }

    Logger.log('Auth: Cleaned up ' + cleaned + ' expired tokens.');
    return cleaned;
  }

  // ═══════════════════════════════════════
  // PRIVATE METHODS
  // ═══════════════════════════════════════

  function _generateMagicToken(email) {
    var config = ConfigReader.getConfig();
    var token = _generateToken();
    var expiry = Date.now() + config.magicLinkExpiryMs;

    var props = PropertiesService.getScriptProperties();
    props.setProperty(TOKEN_PREFIX + token, JSON.stringify({
      email: email.toLowerCase(),
      expiry: expiry,
      created: Date.now(),
      used: false,
    }));

    return token;
  }

  function _validateMagicToken(token) {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(TOKEN_PREFIX + token);

    if (!raw) return null;

    try {
      var data = JSON.parse(raw);

      // CR-03: Reject already-used tokens (prevent replay)
      if (data.used === true) {
        props.deleteProperty(TOKEN_PREFIX + token);
        return null;
      }

      // Check expiry
      if (data.expiry < Date.now()) {
        props.deleteProperty(TOKEN_PREFIX + token);
        return null;
      }

      // Mark as used
      data.used = true;
      props.setProperty(TOKEN_PREFIX + token, JSON.stringify(data));

      // Delete after short delay to prevent replay but allow page load
      // Token will be cleaned up by cleanupExpiredTokens()

      return data.email;
    } catch (_e) {
      return null;
    }
  }

  function _validateSessionToken(token) {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(SESSION_PREFIX + token);

    if (!raw) return null;

    try {
      var data = JSON.parse(raw);

      if (data.expiry < Date.now()) {
        props.deleteProperty(SESSION_PREFIX + token);
        return null;
      }

      return data.email;
    } catch (_e) {
      return null;
    }
  }

  function _generateToken() {
    // Generate a random token using Utilities
    // M-66: Note — base64-encodes the UUID string chars (not raw bytes), producing an
    // inflated token. This is acceptable because two concatenated UUIDs (64 hex chars)
    // still provide sufficient entropy (~128 bits) for token security.
    var bytes = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    return Utilities.base64EncodeWebSafe(bytes).replace(/[=]+$/, '');
  }

  function _buildEmailHtml(config, signInUrl, email) {
    // Dynamic accent color from config
    var hue = config.accentHue || 250;
    var accent = 'hsl(' + hue + ', 70%, 55%)';
    var accentLight = 'hsl(' + hue + ', 70%, 95%)';

    var safeInitials = escapeHtml(String(config.logoInitials || ''));
    var safeOrgName = escapeHtml(String(config.orgName || ''));
    var safeExpiry = escapeHtml(String(config.magicLinkExpiryDays || 7));

    return '<!DOCTYPE html><html><body style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin:0; padding:40px 20px; background:#f5f5f5;">'
      + '<div style="max-width:480px; margin:0 auto; background:#fff; border-radius:16px; padding:40px 32px; box-shadow:0 2px 12px rgba(0,0,0,0.08);">'
      + '<div style="text-align:center; margin-bottom:24px;">'
      + '<div style="display:inline-block; width:48px; height:48px; border-radius:12px; background:' + accent + '; color:#fff; font-size:20px; font-weight:700; line-height:48px;">' + safeInitials + '</div>'
      + '</div>'
      + '<h1 style="text-align:center; font-size:20px; color:#1a1a2e; margin:0 0 8px;">Sign in to ' + safeOrgName + '</h1>'
      + '<p style="text-align:center; color:#666; font-size:14px; margin:0 0 28px;">Click the button below to access your dashboard.</p>'
      + '<div style="text-align:center; margin-bottom:28px;">'
      + '<a href="' + signInUrl + '" style="display:inline-block; background:' + accent + '; color:#fff; text-decoration:none; padding:14px 36px; border-radius:10px; font-size:16px; font-weight:600;">Sign In</a>'
      + '</div>'
      + '<p style="color:#999; font-size:12px; text-align:center; line-height:1.6;">'
      + 'This link expires in ' + safeExpiry + ' days.<br>'
      + 'If you didn\'t request this, you can safely ignore this email.'
      + '</p>'
      + '<hr style="border:none; border-top:1px solid #eee; margin:24px 0;">'
      + '<p style="color:#bbb; font-size:11px; text-align:center;">' + safeOrgName + ' Grievance Dashboard</p>'
      + '</div></body></html>';
  }

  // Public API
  return {
    resolveUser: resolveUser,
    sendMagicLink: sendMagicLink,
    createSessionToken: createSessionToken,
    invalidateSession: invalidateSession,
    cleanupExpiredTokens: cleanupExpiredTokens,
  };

})();


// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

/**
 * Client-callable: Send a magic link to the given email
 */
function authSendMagicLink(email, rememberMe) {
  return Auth.sendMagicLink(email, rememberMe);
}

/**
 * Client-callable: Create a session token after successful auth.
 * CR-02: email is resolved server-side — never trust client-supplied email.
 */
function authCreateSessionToken() {
  var email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (_e) { /* SSO not available */ }
  if (!email) {
    return { error: 'Unable to resolve authenticated user. Please sign in again.' };
  }
  return Auth.createSessionToken(email);
}

/**
 * Client-callable: Invalidate session (logout)
 */
function authLogout(sessionToken) {
  Auth.invalidateSession(sessionToken);
  return { success: true };
}

/**
 * Script-editor verification: checks SSO, ScriptProperties, and Config tab.
 * Run from the Apps Script editor to confirm the auth module is ready.
 */
function initWebDashboardAuth() {
  var email = Session.getActiveUser().getEmail();
  Logger.log('SSO email: ' + (email || '(not available — deploy as web app to test SSO)'));

  var props = PropertiesService.getScriptProperties();
  Logger.log('ScriptProperties accessible: ' + (props !== null));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  Logger.log('Config tab found: ' + (configSheet !== null));

  Logger.log('Auth module ready.');
}

/**
 * Global wrapper: Clean up expired auth tokens.
 * Called by daily time-based trigger.
 */
function authCleanupExpiredTokens() {
  return Auth.cleanupExpiredTokens();
}

/**
 * Installs a daily trigger to auto-clean expired auth tokens.
 * Run once from the Apps Script editor. Idempotent — removes existing trigger first.
 */
function installTokenCleanupTrigger() {
  // Remove any existing cleanup triggers to prevent duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'authCleanupExpiredTokens') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new daily trigger (runs between 2-3 AM in script timezone)
  ScriptApp.newTrigger('authCleanupExpiredTokens')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  Logger.log('Token cleanup trigger installed — runs daily at ~2 AM.');
}
