/**
 * Auth.gs — Authentication system for the web app dashboard
 *
 * WHAT THIS FILE DOES:
 *   Authentication system for the web app dashboard. Two auth paths:
 *     (1) Google SSO via Session.getActiveUser().getEmail()
 *     (2) Magic link tokens emailed via GmailApp (primary) / MailApp (fallback)
 *   Auth.resolveUser(e) is the entry point — tries session token first, then
 *   SSO, then magic link token from URL params. Creates session tokens for
 *   "remember me" functionality stored in ScriptProperties.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Google SSO is preferred but Session.getActiveUser() returns empty in
 *   "Execute as: Me" web apps for non-Google users. Magic links solve this
 *   by emailing a one-time token. Token storage in ScriptProperties (not
 *   CacheService) ensures tokens survive cache eviction. GmailApp is primary
 *   sender because it provides better deliverability than MailApp; MailApp is
 *   kept as fallback. The Auth IIFE pattern creates a namespace without
 *   polluting the global scope.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Nobody can log into the web app dashboard. SSO failure means Google-
 *   authenticated users can't access the SPA. Magic link failure means
 *   non-Google users are locked out. If session token validation breaks,
 *   "remember me" stops working and users must re-authenticate every visit.
 *   If both GmailApp and MailApp fail, the deploy needs re-authorization.
 *
 * DEPENDENCIES:
 *   Depends on: Session, PropertiesService, GmailApp, MailApp (GAS built-ins).
 *   Used by: 22_WebDashApp.gs (doGet calls Auth.resolveUser),
 *            00_Security.gs (Auth.resolveEmailFromToken for token-based authorization).
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
    } catch (_err) { Logger.log('_err: ' + (_err.message || _err)); }

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
    var step = 'init';
    try {
    email = String(email).trim().toLowerCase();

    // Rate limiting — max 3 magic links per email per 15 minutes
    step = 'cache';
    var cache = CacheService.getScriptCache();
    var rateKey = 'MAGIC_RATE_' + email;
    var count = parseInt(cache.get(rateKey) || '0', 10);
    if (count >= 3) {
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }
    cache.put(rateKey, String(count + 1), 900);

    // Validate email exists in directory
    step = 'lookup';
    Logger.log('Auth.sendMagicLink STEP 0: looking up ' + email + ' in Member Directory');
    var userRecord = DataService.findUserByEmail(email);
    if (!userRecord) {
      Logger.log('Auth.sendMagicLink: email not found in directory — returning generic message');
      // Don't reveal whether email exists — security best practice
      // Simulate processing time to prevent timing-based enumeration
      Utilities.sleep(500 + Math.floor(Math.random() * 500));
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }

    step = 'config';
    Logger.log('Auth.sendMagicLink STEP 1: user found, building token for ' + email);
    // ConfigReader.getConfig() can throw if its CacheService cache expired and
    // getActiveSpreadsheet() returns null in certain web app execution contexts.
    // Magic link emails only need a few config fields — fall back to safe defaults
    // so the email still sends even when the config sheet is temporarily unreachable.
    var config;
    try {
      config = ConfigReader.getConfig();
    } catch (cfgErr) {
      Logger.log('Auth.sendMagicLink: ConfigReader config fetch failed (' + cfgErr.message + ') — using defaults');
      config = {
        orgName: 'Dashboard',
        logoInitials: '',
        accentHue: 250,
        magicLinkExpiryMs: 7 * 24 * 60 * 60 * 1000,
        magicLinkExpiryDays: 7,
      };
    }

    step = 'token';
    Logger.log('Auth.sendMagicLink STEP 2: config loaded, orgName=' + config.orgName);
    var token = _generateMagicToken(email, config);

    step = 'url';
    Logger.log('Auth.sendMagicLink STEP 3: token generated, fetching web app URL');
    var webAppUrl = ScriptApp.getService().getUrl();
    if (!webAppUrl) {
      Logger.log('Auth.sendMagicLink ERROR: ScriptApp.getService().getUrl() returned null — is this script deployed as a web app?');
      return { success: false, message: 'Web app URL could not be resolved. Please contact your administrator.' };
    }

    var linkParams = '?token=' + encodeURIComponent(token);
    if (rememberMe) {
      linkParams += '&remember=1';
    }

    var signInUrl = webAppUrl + linkParams;

    step = 'build-email';
    var subject = 'Sign in to ' + config.orgName + ' Dashboard';
    var htmlBody = _buildEmailHtml(config, signInUrl, email);

    step = 'send';
    Logger.log('Auth.sendMagicLink STEP 4: attempting email send to ' + email);

    // PRIMARY: GmailApp (uses gmail.send scope — authorized in web app deployment)
    // FALLBACK: MailApp (uses script.send_mail scope — requires separate re-auth)
    // Reason: gmail.send was in appsscript.json long before script.send_mail was added
    // (v4.24.9). If the deployment wasn't re-authorized after v4.24.9, MailApp throws.
    // GmailApp uses the pre-existing gmail.send scope and avoids that auth gap.
    var sendError = null;

    try {
      GmailApp.sendEmail(email, subject, '', {
        htmlBody: htmlBody,
        name: config.orgName + ' Dashboard',
        noReply: false,
      });
      Logger.log('Auth.sendMagicLink STEP 5: GmailApp send succeeded');
      return { success: true, message: 'Sign-in link sent to ' + email };
    } catch (gmailErr) {
      Logger.log('Auth.sendMagicLink GmailApp FAILED: ' + gmailErr.message + ' — trying MailApp fallback');
      sendError = gmailErr;
    }

    // Fallback: MailApp (works if script.send_mail scope is authorized in deployment)
    try {
      // Quota guard — MailApp has a daily limit
      var remaining = 0;
      try { remaining = MailApp.getRemainingDailyQuota(); } catch (_q) { remaining = 1; }
      if (remaining <= 0) {
        Logger.log('Auth.sendMagicLink: MailApp quota exhausted');
        return { success: false, message: 'Email quota reached for today. Please use Google Sign-In or try again tomorrow.' };
      }

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        noReply: true,
      });
      Logger.log('Auth.sendMagicLink STEP 5: MailApp fallback send succeeded');
      return { success: true, message: 'Sign-in link sent to ' + email };
    } catch (mailErr) {
      Logger.log('Auth.sendMagicLink BOTH senders FAILED. GmailApp: ' + sendError.message + ' | MailApp: ' + mailErr.message);

      var msg = 'Failed to send sign-in link. ';
      var combinedMsg = (sendError.message || '') + ' ' + (mailErr.message || '');
      if (combinedMsg.indexOf('quota') >= 0) {
        msg += 'Email quota exhausted. Please use Google Sign-In.';
      } else if (combinedMsg.indexOf('authorization') >= 0 || combinedMsg.indexOf('Permission') >= 0 || combinedMsg.indexOf('auth') >= 0) {
        msg += 'Email authorization error — the web app may need to be re-deployed. Please use Google Sign-In.';
      } else if (combinedMsg.indexOf('invalid') >= 0) {
        msg += 'The email address may be invalid. Please check and try again.';
      } else {
        msg += 'Please try again or use Google Sign-In.';
      }
      return { success: false, message: msg };
    }
    } catch (outerErr) {
      Logger.log('Auth.sendMagicLink OUTER ERROR at step [' + step + ']: ' + outerErr.message + '\n' + (outerErr.stack || ''));

      // Provide actionable error messages based on which step failed
      var outerMsg = 'Failed to send email. ';
      if (step === 'cache') {
        outerMsg += 'A temporary service error occurred. Please try again.';
      } else if (step === 'lookup') {
        outerMsg += 'Could not access the member directory. Please try again or use Google Sign-In.';
      } else if (step === 'config') {
        outerMsg += 'Could not load dashboard configuration. Please contact your administrator.';
      } else if (step === 'token' || step === 'url') {
        outerMsg += 'A server configuration error occurred. Please contact your administrator.';
      } else if (step === 'build-email' || step === 'send') {
        outerMsg += 'Please try again or use Google Sign-In.';
      } else {
        outerMsg += 'Please try again.';
      }
      return { success: false, message: outerMsg };
    }
  }

  /**
   * Creates a session token for session persistence.
   * Called after successful magic link auth. If shortLived is true, token
   * expires in 24 hours (for users who didn't check "remember me") instead
   * of the full cookieDuration (default 30 days).
   * @param {string} email
   * @param {boolean} [shortLived] - If true, 24-hour expiry instead of full duration
   * @returns {string} Session token
   */
  function createSessionToken(email, shortLived) {
    // Auto-evict expired tokens if approaching quota
    try {
      var props = PropertiesService.getScriptProperties();
      var allProps = props.getProperties();
      var propSize = JSON.stringify(allProps).length;
      if (propSize > 350000) { // 350KB of 500KB — proactive cleanup
        var now = Date.now();
        var keys = Object.keys(allProps);
        for (var i = 0; i < keys.length; i++) {
          var k = keys[i];
          if (k.indexOf('SESSION_') === 0 || k.indexOf('MAGIC_TOKEN_') === 0) {
            try {
              var val = JSON.parse(allProps[k]);
              if (val.expiry && val.expiry < now) {
                props.deleteProperty(k);
              }
            } catch (_parseErr) {
              props.deleteProperty(k); // corrupt entry
            }
          }
        }
      }
    } catch (_evictErr) {
      Logger.log('Auto token eviction failed: ' + _evictErr.message);
    }

    var config;
    try {
      config = ConfigReader.getConfig();
    } catch (_cfgErr) {
      Logger.log('Auth.createSessionToken: ConfigReader failed (' + _cfgErr.message + ') — using default cookie duration');
      config = { cookieDurationMs: 30 * 24 * 60 * 60 * 1000 }; // 30 days default
    }
    var token = _generateToken();
    var SHORT_SESSION_MS = 24 * 60 * 60 * 1000; // 24 hours
    var duration = shortLived ? SHORT_SESSION_MS : (config.cookieDurationMs || 30 * 24 * 60 * 60 * 1000);
    var expiry = Date.now() + duration;

    try {
      props = PropertiesService.getScriptProperties();
      props.setProperty(SESSION_PREFIX + token, JSON.stringify({
        email: email.toLowerCase(),
        expiry: expiry,
        created: Date.now(),
      }));
    } catch (propErr) {
      Logger.log('Auth.createSessionToken: PropertiesService write failed (' + propErr.message + ')');
      // C3: Return error instead of a token that doesn't exist server-side
      return { error: 'session_storage_failed', message: 'Could not persist session. Please try again.' };
    }

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

    // Fix 2.9: Monitor PropertiesService quota (500KB limit)
    // Reuse the already-loaded snapshot when no deletions occurred (avoids double getProperties call)
    var sizeSource = cleaned > 0 ? props.getProperties() : all;
    var totalSize = JSON.stringify(sizeSource).length;
    if (totalSize > 400000) {
      Logger.log('WARNING: ScriptProperties usage at ' + Math.round(totalSize / 1024) + 'KB / 500KB');
      // M9: Escalate via recordSecurityEvent so admins get notified
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('QUOTA_WARNING', 'ScriptProperties at ' + Math.round(totalSize / 1024) + 'KB / 500KB', { totalSize: totalSize }, 'HIGH');
      }
    }

    return cleaned;
  }

  // ═══════════════════════════════════════
  // PRIVATE METHODS
  // ═══════════════════════════════════════

  function _generateMagicToken(email, config) {
    config = config || ConfigReader.getConfig();
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

      // C1+C4: Immediately delete token after successful validation.
      // Eliminates TOCTOU race (two concurrent requests validating same token)
      // and prevents token accumulation in PropertiesService quota.
      props.deleteProperty(TOKEN_PREFIX + token);

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

  function _hslToHex(h, s, l) {
    s /= 100; l /= 100;
    var c = (1 - Math.abs(2 * l - 1)) * s;
    var x = c * (1 - Math.abs((h / 60) % 2 - 1));
    var m = l - c / 2;
    var r = 0, g = 0, b = 0;
    if (h < 60) { r = c; g = x; }
    else if (h < 120) { r = x; g = c; }
    else if (h < 180) { g = c; b = x; }
    else if (h < 240) { g = x; b = c; }
    else if (h < 300) { r = x; b = c; }
    else { r = c; b = x; }
    var toHex = function(v) { var hex = Math.round((v + m) * 255).toString(16); return hex.length === 1 ? '0' + hex : hex; };
    return '#' + toHex(r) + toHex(g) + toHex(b);
  }

  function _buildEmailHtml(config, signInUrl, email) {
    // Convert accent hue to hex — hsl() is not supported by many email clients
    // (Outlook, older Gmail, Yahoo) which silently strip it, making the button invisible
    var hue = config.accentHue || 250;
    var accent = _hslToHex(hue, 70, 55);
    var safeInitials = escapeHtml(String(config.logoInitials || ''));
    var safeOrgName = escapeHtml(String(config.orgName || ''));
    var safeExpiry = escapeHtml(String(config.magicLinkExpiryDays || 7));

    // "Bulletproof button" pattern: <table> with bgcolor works in all email clients
    // including Outlook, which ignores CSS background on <a> tags
    return '<!DOCTYPE html><html><body style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; margin:0; padding:40px 20px; background:#f5f5f5;">'
      + '<div style="max-width:480px; margin:0 auto; background:#fff; border-radius:16px; padding:40px 32px; box-shadow:0 2px 12px rgba(0,0,0,0.08);">'
      + '<div style="text-align:center; margin-bottom:24px;">'
      + '<div style="display:inline-block; width:48px; height:48px; border-radius:12px; background:' + accent + '; color:#fff; font-size:20px; font-weight:700; line-height:48px; text-align:center;">' + safeInitials + '</div>'
      + '</div>'
      + '<h1 style="text-align:center; font-size:20px; color:#1a1a2e; margin:0 0 8px;">Sign in to ' + safeOrgName + '</h1>'
      + '<p style="text-align:center; color:#666; font-size:14px; margin:0 0 28px;">Click the button below to access your dashboard.</p>'
      + '<table role="presentation" cellspacing="0" cellpadding="0" border="0" align="center" style="margin:0 auto 28px;">'
      + '<tr><td align="center" bgcolor="' + accent + '" style="border-radius:10px;">'
      + '<a href="' + signInUrl + '" target="_blank" style="display:inline-block; background:' + accent + '; color:#ffffff; text-decoration:none; padding:14px 36px; border-radius:10px; font-size:16px; font-weight:600; font-family:-apple-system,BlinkMacSystemFont,sans-serif;">Sign In</a>'
      + '</td></tr></table>'
      + '<p style="color:#999; font-size:12px; text-align:center; line-height:1.6;">'
      + 'This link expires in ' + safeExpiry + ' days.<br>'
      + 'If you didn\'t request this, you can safely ignore this email.'
      + '</p>'
      + '<hr style="border:none; border-top:1px solid #eee; margin:24px 0;">'
      + '<p style="color:#bbb; font-size:11px; text-align:center;">' + safeOrgName + ' Grievance Dashboard</p>'
      + '</div></body></html>';
  }

  // ═══════════════════════════════════════
  // DEV-ONLY PIN LOGIN
  // Steward generates a 6-digit PIN for a member. Member enters just the
  // PIN on the login screen — no email needed. PIN is reusable for 14 days.
  // Guarded by isProductionMode() — never available in prod.
  // ═══════════════════════════════════════

  var DEV_PIN_PREFIX = 'DEV_PIN_';
  var DEV_PIN_EXPIRY_MS = 14 * 24 * 60 * 60 * 1000; // 14 days

  /**
   * Generates a 6-digit login PIN for a member. Called by a steward.
   * DEV MODE ONLY — refused in production.
   * @param {string} memberEmail - The member to generate a PIN for
   * @param {string} sessionToken - Steward's session token for auth
   * @returns {Object} { success, pin?, memberName?, message }
   */
  function generateDevPin(memberEmail, sessionToken) {
    if (isProductionMode()) {
      return { success: false, message: 'PIN login is not available in production.' };
    }
    memberEmail = String(memberEmail).trim().toLowerCase();
    if (!memberEmail || memberEmail.indexOf('@') < 0) {
      return { success: false, message: 'Invalid email address.' };
    }

    // Validate member exists in directory
    var userRecord = DataService.findUserByEmail(memberEmail);
    if (!userRecord) {
      return { success: false, message: 'Member not found in directory.' };
    }

    // Remove any existing PIN for this member (search all DEV_PIN_ keys)
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var keys = Object.keys(allProps);
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].indexOf(DEV_PIN_PREFIX) === 0) {
        try {
          var existing = JSON.parse(allProps[keys[i]]);
          if (existing.email === memberEmail) {
            props.deleteProperty(keys[i]);
          }
        } catch (_e) { /* skip corrupt */ }
      }
    }

    // Generate unique 6-digit PIN — ensure no collision
    var pin;
    var attempts = 0;
    do {
      pin = String(100000 + Math.floor(Math.random() * 900000));
      attempts++;
    } while (props.getProperty(DEV_PIN_PREFIX + pin) && attempts < 10);

    // Store PIN → email mapping (keyed by PIN for fast lookup at login)
    props.setProperty(DEV_PIN_PREFIX + pin, JSON.stringify({
      email: memberEmail,
      expiry: Date.now() + DEV_PIN_EXPIRY_MS,
      created: Date.now(),
    }));

    var memberName = userRecord.name || userRecord.firstName || memberEmail;
    return { success: true, pin: pin, memberName: memberName, message: 'PIN generated for ' + memberName };
  }

  /**
   * Validates a dev login PIN. Returns the email if valid, null if not.
   * Member enters only the PIN — no email needed.
   * PIN is reusable until it expires (14 days).
   * DEV MODE ONLY.
   * @param {string} pin - 6-digit PIN
   * @returns {string|null} verified member email or null
   */
  function verifyDevPin(pin) {
    if (isProductionMode()) return null;
    pin = String(pin).trim();
    if (!pin || pin.length !== 6) return null;

    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(DEV_PIN_PREFIX + pin);
    if (!raw) return null;

    try {
      var data = JSON.parse(raw);
      if (data.expiry < Date.now()) {
        props.deleteProperty(DEV_PIN_PREFIX + pin);
        return null;
      }
      return data.email;
    } catch (_e) {
      return null;
    }
  }

  // Public API
  return {
    resolveUser: resolveUser,
    sendMagicLink: sendMagicLink,
    createSessionToken: createSessionToken,
    invalidateSession: invalidateSession,
    cleanupExpiredTokens: cleanupExpiredTokens,
    generateDevPin: generateDevPin,
    verifyDevPin: verifyDevPin,
    /**
     * Resolve a verified email from a session token.
     * Used by auth helpers when Session.getActiveUser() is empty
     * (magic link / session token users in Execute-as-Me web apps).
     * @param {string} token
     * @returns {string|null} verified email or null
     */
    resolveEmailFromToken: _validateSessionToken,
    validateSessionToken: _validateSessionToken,
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
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
 * Global wrapper: Clean up expired auth tokens.
 * Called by daily time-based trigger.
 */
function authCleanupExpiredTokens() {
  return Auth.cleanupExpiredTokens();
}

/**
 * Client-callable: Steward generates a login PIN for a member.
 * DEV MODE ONLY. Returns { success, pin, memberName, message }.
 */
function authGenerateDevPin(memberEmail, sessionToken) {
  if (isProductionMode()) {
    return { success: false, message: 'PIN login is not available in production.' };
  }
  // Require steward auth
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward authorization required.' };
  return Auth.generateDevPin(memberEmail, sessionToken);
}

/**
 * Client-callable: Member verifies a login PIN (no email needed).
 * DEV MODE ONLY — returns session token on success.
 */
function authVerifyDevPin(pin) {
  if (isProductionMode()) {
    return { success: false, message: 'PIN login is not available in production.' };
  }
  var verifiedEmail = Auth.verifyDevPin(pin);
  if (!verifiedEmail) {
    return { success: false, message: 'Invalid or expired PIN. Please ask your steward for a new one.' };
  }
  // Create a session token so the member stays logged in
  var sessionToken = Auth.createSessionToken(verifiedEmail);
  if (sessionToken && typeof sessionToken === 'object' && sessionToken.error) {
    return { success: false, message: 'PIN verified but session creation failed. Please try again.' };
  }
  return { success: true, sessionToken: sessionToken, email: verifiedEmail };
}

/**
 * DIAGNOSTIC: Test email sending for the magic link flow.
 * Run this from the Apps Script editor to verify GmailApp and MailApp work.
 * Reports which sender succeeded, which failed, and why.
 * Does NOT store a token — the email is sent but the link will not work.
 *
 * Usage: Open Apps Script editor → Select function → Run testAuthEmailSend
 * @param {string} [testEmail] - Override destination (default: script owner)
 */
function testAuthEmailSend(testEmail) {
  var to = testEmail || Session.getEffectiveUser().getEmail();
  var subject = '[SolidBase] Magic Link Email Test';
  var htmlBody = '<p>This is a test of the magic link email system.</p>'
    + '<p>If you received this, email sending is working correctly.</p>'
    + '<p>Sent at: ' + new Date().toISOString() + '</p>';

  var results = { to: to, gmail: null, mailapp: null };

  try {
    GmailApp.sendEmail(to, subject + ' (GmailApp)', '', { htmlBody: htmlBody, name: 'SolidBase Test' });
    results.gmail = 'SUCCESS';
    Logger.log('testAuthEmailSend: GmailApp → SUCCESS');
  } catch (e) {
    results.gmail = 'FAILED: ' + e.message;
    Logger.log('testAuthEmailSend: GmailApp → FAILED: ' + e.message);
  }

  try {
    MailApp.sendEmail({ to: to, subject: subject + ' (MailApp)', htmlBody: htmlBody, noReply: true });
    results.mailapp = 'SUCCESS';
    Logger.log('testAuthEmailSend: MailApp → SUCCESS');
  } catch (e) {
    results.mailapp = 'FAILED: ' + e.message;
    Logger.log('testAuthEmailSend: MailApp → FAILED: ' + e.message);
  }

  Logger.log('testAuthEmailSend results: ' + JSON.stringify(results));
  return results;
}
