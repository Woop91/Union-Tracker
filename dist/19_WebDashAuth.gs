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
  var DEVICE_KEY_PREFIX = 'DEVICE_KEY_';

  /**
   * Attempts to resolve the current user from available auth signals.
   * Priority: (1) session token in URL, (2) SSO, (3) magic link token in URL
   * @param {Object} e - doGet event object
   * @returns {Object|null} { email: string, method: 'sso'|'magic'|'session'|'pin' } or null
   */
  function resolveUser(e) {
    var params = e ? e.parameter || {} : {};

    // 0. Explicit logout — bypass all auth methods
    if (params.loggedout === '1') return null;

    // 1. Check for session token (returning user with "remember me")
    if (params.sessionToken) {
      var sessionData = _getSessionData(params.sessionToken);
      if (sessionData && sessionData.email) {
        // Preserve original auth method (e.g. 'pin') if stored in session
        return { email: sessionData.email, method: sessionData.authMethod || 'session' };
      }
    }

    // 2. Check Google SSO
    try {
      var ssoUser = Session.getActiveUser().getEmail();
      if (ssoUser && ssoUser !== '') {
        return { email: ssoUser.toLowerCase(), method: 'sso' };
      }
    } catch (_err) { log_('_err', (_err.message || _err)); }

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

    // Cache reference — needed for both rate limiters below
    step = 'cache';
    var cache = CacheService.getScriptCache();

    // Global session rate limit — prevent email enumeration via rapid attempts
    var sessionId = Session.getTemporaryActiveUserKey() || 'anon';
    var globalMagicKey = 'MAGIC_GLOBAL_' + sessionId;
    var globalMagicCount = parseInt(cache.get(globalMagicKey) || '0', 10);
    if (globalMagicCount >= 5) {
      Utilities.sleep(500 + Math.floor(Math.random() * 500));
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }
    cache.put(globalMagicKey, String(globalMagicCount + 1), 900);

    // Rate limiting — max 3 magic links per email per 15 minutes
    var rateKey = 'MAGIC_RATE_' + email;
    var count = parseInt(cache.get(rateKey) || '0', 10);
    if (count >= 3) {
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }
    cache.put(rateKey, String(count + 1), 900);

    // Validate email exists in directory
    step = 'lookup';
    log_('Auth.sendMagicLink STEP 0', 'looking up ' + email + ' in Member Directory');
    var userRecord = DataService.findUserByEmail(email);
    if (!userRecord) {
      log_('Auth.sendMagicLink', 'email not found in directory — returning generic message');
      // Don't reveal whether email exists — security best practice
      // Simulate processing time to prevent timing-based enumeration
      Utilities.sleep(500 + Math.floor(Math.random() * 500));
      return { success: true, message: 'If this email is in our directory, you will receive a sign-in link.' };
    }

    step = 'config';
    log_('Auth.sendMagicLink STEP 1', 'user found, building token for ' + email);
    // ConfigReader.getConfig() can throw if its CacheService cache expired and
    // getActiveSpreadsheet() returns null in certain web app execution contexts.
    // Magic link emails only need a few config fields — fall back to safe defaults
    // so the email still sends even when the config sheet is temporarily unreachable.
    var config;
    try {
      config = ConfigReader.getConfig();
    } catch (cfgErr) {
      log_('Auth.sendMagicLink', 'ConfigReader config fetch failed (' + cfgErr.message + ') — using defaults');
      config = {
        orgName: 'DDS',
        logoInitials: '',
        accentHue: 250,
        magicLinkExpiryMs: 7 * 24 * 60 * 60 * 1000,
        magicLinkExpiryDays: 7,
      };
    }

    step = 'token';
    log_('Auth.sendMagicLink STEP 2', 'config loaded, orgName=' + config.orgName);
    var token = _generateMagicToken(email, config);

    step = 'url';
    log_('Auth.sendMagicLink STEP 3', 'token generated, fetching web app URL');
    var webAppUrl = ScriptApp.getService().getUrl();
    if (!webAppUrl) {
      log_('Auth.sendMagicLink ERROR', 'ScriptApp.getService().getUrl() returned null — is this script deployed as a web app?');
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
    log_('Auth.sendMagicLink STEP 4', 'attempting email send to ' + email);

    return _sendMagicLinkEmail(email, subject, htmlBody, config.orgName);
    } catch (outerErr) {
      log_('sendMagicLink', 'Auth.sendMagicLink OUTER ERROR at step [' + step + ']: ' + outerErr.message + '\n' + (outerErr.stack || ''));

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
   * Creates a session token for "remember me" functionality.
   * Called after successful auth if remember=1.
   * @param {string} email
   * @param {string} [authMethod] - Original auth method ('sso', 'magic', 'pin'). Stored in session for access-level decisions.
   * @returns {string|{error: string, message: string}} Session token string on success, or error object on failure (e.g. PropertiesService write failure)
   */
  function createSessionToken(email, authMethod) {
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
          if (k.indexOf('SESSION_') === 0 || k.indexOf('MAGIC_TOKEN_') === 0 || k.indexOf('DEVICE_KEY_') === 0) {
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
      log_('Auto token eviction failed', _evictErr.message);
    }

    var config;
    try {
      config = ConfigReader.getConfig();
    } catch (_cfgErr) {
      log_('Auth.createSessionToken', 'ConfigReader failed (' + _cfgErr.message + ') — using default cookie duration');
      config = { cookieDurationMs: 30 * 24 * 60 * 60 * 1000 }; // 30 days default
    }
    var token = _generateToken();
    var expiry = Date.now() + (config.cookieDurationMs || 30 * 24 * 60 * 60 * 1000);

    try {
      props = PropertiesService.getScriptProperties();
      var sessionObj = {
        email: email.toLowerCase(),
        expiry: expiry,
        created: Date.now(),
      };
      if (authMethod) sessionObj.authMethod = authMethod;
      props.setProperty(SESSION_PREFIX + token, JSON.stringify(sessionObj));
    } catch (propErr) {
      log_('Auth.createSessionToken', 'PropertiesService write failed (' + propErr.message + ')');
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
    var key = SESSION_PREFIX + token;
    props.deleteProperty(key);

    // Verify deletion — a surviving session token is a security risk
    if (props.getProperty(key) !== null) {
      log_('SECURITY WARNING', 'session token ' + key.substring(0, 12) + '... was not deleted — forcing overwrite');
      try { props.setProperty(key, '{"invalidated":true}'); } catch (_) {}
    }
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
      if (key.indexOf(TOKEN_PREFIX) === 0 || key.indexOf(SESSION_PREFIX) === 0 || key.indexOf(DEVICE_KEY_PREFIX) === 0) {
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

    log_('Auth', 'Cleaned up ' + cleaned + ' expired tokens.');

    // Fix 2.9: Monitor PropertiesService quota (500KB limit)
    // Reuse the already-loaded snapshot when no deletions occurred (avoids double getProperties call)
    var sizeSource = cleaned > 0 ? props.getProperties() : all;
    var totalSize = JSON.stringify(sizeSource).length;
    if (totalSize > 400000) {
      log_('WARNING', 'ScriptProperties usage at ' + Math.round(totalSize / 1024) + 'KB / 500KB');
      // M9: Escalate via recordSecurityEvent so admins get notified
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('QUOTA_WARNING', 'HIGH', 'ScriptProperties at ' + Math.round(totalSize / 1024) + 'KB / 500KB', { totalSize: totalSize });
      }
    }

    return cleaned;
  }

  // ═══════════════════════════════════════
  // PRIVATE METHODS
  // ═══════════════════════════════════════

  /**
   * Sends a magic link email using GmailApp (primary) with MailApp fallback.
   * Extracted from sendMagicLink for clarity.
   *
   * PRIMARY: GmailApp (uses gmail.send scope — authorized in web app deployment)
   * FALLBACK: MailApp (uses script.send_mail scope — requires separate re-auth)
   * Reason: gmail.send was in appsscript.json long before script.send_mail was added
   * (v4.24.9). If the deployment wasn't re-authorized after v4.24.9, MailApp throws.
   * GmailApp uses the pre-existing gmail.send scope and avoids that auth gap.
   *
   * @param {string} email - Recipient email address
   * @param {string} subject - Email subject line
   * @param {string} htmlBody - HTML email body
   * @param {string} orgName - Organization name for sender display
   * @returns {Object} { success: boolean, message: string }
   * @private
   */
  function _sendMagicLinkEmail(email, subject, htmlBody, orgName) {
    var sendError = null;

    try {
      GmailApp.sendEmail(email, subject, '', {
        htmlBody: htmlBody,
        name: (orgName || 'Dashboard') + ' Dashboard',
        noReply: false,
      });
      log_('Auth._sendMagicLinkEmail', 'GmailApp send succeeded');
      return { success: true, message: 'Sign-in link sent to ' + email };
    } catch (gmailErr) {
      log_('Auth._sendMagicLinkEmail GmailApp FAILED', gmailErr.message + ' — trying MailApp fallback');
      sendError = gmailErr;
    }

    // Fallback: MailApp (works if script.send_mail scope is authorized in deployment)
    try {
      // Quota guard — MailApp has a daily limit
      var remaining = 0;
      try { remaining = MailApp.getRemainingDailyQuota(); } catch (_q) { remaining = 1; }
      if (remaining <= 0) {
        log_('Auth._sendMagicLinkEmail', 'MailApp quota exhausted');
        return { success: false, message: 'Email quota reached for today. Please use Google Sign-In or try again tomorrow.' };
      }

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        noReply: true,
      });
      log_('Auth._sendMagicLinkEmail', 'MailApp fallback send succeeded');
      return { success: true, message: 'Sign-in link sent to ' + email };
    } catch (mailErr) {
      log_('Auth._sendMagicLinkEmail BOTH senders FAILED. GmailApp', sendError.message + ' | MailApp: ' + mailErr.message);

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
  }

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
      // KNOWN LIMITATION: PropertiesService get+delete is not atomic.
      // Two concurrent requests could both validate the same token.
      // Risk is minimal: magic links are single-use email tokens, window is milliseconds.
      // LockService was considered but rejected for latency on the hot auth path.
      props.deleteProperty(TOKEN_PREFIX + token);

      return data.email;
    } catch (_e) {
      return null;
    }
  }

  /**
   * Reads full session data for a token. Returns null if expired or missing.
   * @param {string} token
   * @returns {Object|null} { email, expiry, created, authMethod } or null
   * @private
   */
  function _getSessionData(token) {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(SESSION_PREFIX + token);
    if (!raw) return null;
    try {
      var data = JSON.parse(raw);
      if (data.expiry < Date.now()) {
        props.deleteProperty(SESSION_PREFIX + token);
        return null;
      }
      return data;
    } catch (_e) {
      return null;
    }
  }

  /** @returns {string|null} Email or null — backward-compatible wrapper */
  function _validateSessionToken(token) {
    var data = _getSessionData(token);
    return data ? data.email : null;
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

  /**
   * Hash a device key for storage lookup.
   * @param {string} key
   * @returns {string} hex-encoded SHA-256 hash
   * @private
   */
  function _hashDeviceKey(key) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key);
    var hex = '';
    for (var i = 0; i < digest.length; i++) {
      hex += ('0' + (digest[i] & 0xff).toString(16)).slice(-2);
    }
    return hex;
  }

  /**
   * Registers a device key for biometric/quick sign-in.
   * Called after successful login to enable credential-manager-based re-auth.
   * The device key is returned to the client for storage via PasswordCredential.
   * Server stores only a hash → email mapping.
   * @param {string} email - Verified user email (resolved server-side)
   * @returns {Object} { success, deviceKey, displayName }
   */
  function registerDeviceKey(email) {
    if (!email) return { success: false, message: 'No email provided.' };
    email = String(email).trim().toLowerCase();

    var deviceKey = _generateToken();
    var hash = _hashDeviceKey(deviceKey);

    // Device keys last 3x session duration (default 90 days)
    var config;
    try { config = ConfigReader.getConfig(); } catch (_) { config = {}; }
    var ttl = (config.cookieDurationMs || 30 * 24 * 60 * 60 * 1000) * 3;

    try {
      var props = PropertiesService.getScriptProperties();
      props.setProperty(DEVICE_KEY_PREFIX + hash, JSON.stringify({
        email: email,
        expiry: Date.now() + ttl,
        created: Date.now()
      }));
    } catch (err) {
      log_('Auth.registerDeviceKey', 'storage failed: ' + err.message);
      return { success: false, message: 'Could not register device key.' };
    }

    // Look up display name for credential manager label
    var displayName = email;
    try {
      var userRecord = DataService.findUserByEmail(email);
      if (userRecord && userRecord.name) displayName = userRecord.name;
      else if (userRecord && userRecord.firstName) displayName = userRecord.firstName;
    } catch (_) {}

    return { success: true, deviceKey: deviceKey, displayName: displayName };
  }

  /**
   * Authenticates via device key (biometric/quick sign-in flow).
   * Looks up the hashed device key in ScriptProperties, creates a session token.
   * @param {string} deviceKey - The plaintext device key from credential manager
   * @returns {Object} { success, sessionToken, email, displayName } or error
   */
  function loginWithDeviceKey(deviceKey) {
    if (!deviceKey) return { success: false, message: 'No device key provided.' };
    deviceKey = String(deviceKey).trim();

    var hash = _hashDeviceKey(deviceKey);
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(DEVICE_KEY_PREFIX + hash);
    if (!raw) return { success: false, message: 'Device not recognized.' };

    try {
      var data = JSON.parse(raw);
      if (data.expiry < Date.now()) {
        props.deleteProperty(DEVICE_KEY_PREFIX + hash);
        return { success: false, message: 'Device key expired. Please sign in again.' };
      }

      // Update last-used timestamp
      data.lastUsed = Date.now();
      props.setProperty(DEVICE_KEY_PREFIX + hash, JSON.stringify(data));

      // Create session token
      var token = createSessionToken(data.email, 'device_key');
      if (token && typeof token === 'object' && token.error) {
        return { success: false, message: 'Session creation failed.' };
      }

      // Look up display name
      var displayName = data.email;
      try {
        var userRecord = DataService.findUserByEmail(data.email);
        if (userRecord && userRecord.name) displayName = userRecord.name;
        else if (userRecord && userRecord.firstName) displayName = userRecord.firstName;
      } catch (_) {}

      if (typeof logAuditEvent === 'function') {
        logAuditEvent('DEVICE_KEY_LOGIN', { email: data.email });
      }

      return { success: true, sessionToken: token, email: data.email, displayName: displayName };
    } catch (err) {
      log_('Auth.loginWithDeviceKey error', err.message);
      return { success: false, message: 'Device key validation failed.' };
    }
  }

  // Public API
  return {
    resolveUser: resolveUser,
    sendMagicLink: sendMagicLink,
    createSessionToken: createSessionToken,
    invalidateSession: invalidateSession,
    cleanupExpiredTokens: cleanupExpiredTokens,
    /**
     * Resolve a verified email from a session token.
     * Used by auth helpers when Session.getActiveUser() is empty
     * (magic link / session token users in Execute-as-Me web apps).
     * @param {string} token
     * @returns {string|null} verified email or null
     */
    resolveEmailFromToken: _validateSessionToken,
    validateSessionToken: _validateSessionToken,
    /**
     * Check if a session token was created via PIN login.
     * PIN sessions have restricted access — no personal data.
     * @param {string} token
     * @returns {boolean}
     */
    registerDeviceKey: registerDeviceKey,
    loginWithDeviceKey: loginWithDeviceKey,
    isPINSession: function(token) {
      if (!token) return false;
      var data = _getSessionData(token);
      return !!(data && data.authMethod === 'pin');
    },
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
  } catch (_e) { log_('_e', (_e.message || _e)); }
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
 * Client-callable: Register a device key for biometric quick sign-in.
 * Email resolved server-side from session — never trusts client-supplied email.
 * @param {string} sessionToken - Current session token for identity resolution
 * @returns {Object} { success, deviceKey, displayName }
 */
function authRegisterDeviceKey(sessionToken) {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (_) {}
  if (!email && sessionToken) {
    email = Auth.resolveEmailFromToken(sessionToken);
  }
  if (!email) return { success: false, message: 'Not authenticated.' };
  return Auth.registerDeviceKey(email);
}

/**
 * Client-callable: Authenticate via device key (biometric flow).
 * @param {string} deviceKey - The device key from credential manager
 * @returns {Object} { success, sessionToken, email, displayName }
 */
function authLoginWithDeviceKey(deviceKey) {
  return Auth.loginWithDeviceKey(deviceKey);
}

/**
 * Global wrapper: Clean up expired auth tokens.
 * Called by daily time-based trigger.
 */
function authCleanupExpiredTokens() {
  return Auth.cleanupExpiredTokens();
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
  var subject = '[DDS-Dashboard] Magic Link Email Test';
  var htmlBody = '<p>This is a test of the magic link email system.</p>'
    + '<p>If you received this, email sending is working correctly.</p>'
    + '<p>Sent at: ' + new Date().toISOString() + '</p>';

  var results = { to: to, gmail: null, mailapp: null };

  try {
    GmailApp.sendEmail(to, subject + ' (GmailApp)', '', { htmlBody: htmlBody, name: 'DDS-Dashboard Test' });
    results.gmail = 'SUCCESS';
    log_('testAuthEmailSend', 'GmailApp → SUCCESS');
  } catch (e) {
    results.gmail = 'FAILED: ' + e.message;
    log_('testAuthEmailSend', 'GmailApp → FAILED: ' + e.message);
  }

  try {
    MailApp.sendEmail({ to: to, subject: subject + ' (MailApp)', htmlBody: htmlBody, noReply: true });
    results.mailapp = 'SUCCESS';
    log_('testAuthEmailSend', 'MailApp → SUCCESS');
  } catch (e) {
    results.mailapp = 'FAILED: ' + e.message;
    log_('testAuthEmailSend', 'MailApp → FAILED: ' + e.message);
  }

  log_('testAuthEmailSend results', JSON.stringify(results));
  return results;
}
