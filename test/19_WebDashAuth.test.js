/**
 * Tests for 19_WebDashAuth.gs
 *
 * Covers the Auth IIFE: resolveUser(), sendMagicLink(), createSessionToken(),
 * invalidateSession(), cleanupExpiredTokens(), and global wrapper functions.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '19_WebDashAuth.gs'
]);

// ---------------------------------------------------------------------------
// Module-level mocks — set AFTER loadSources since 01_Core.gs defines stubs
// ---------------------------------------------------------------------------

var defaultConfig = {
  orgName: 'Test Org',
  magicLinkExpiryMs: 86400000,       // 1 day
  cookieDurationMs: 2592000000,       // 30 days
  magicLinkExpiryDays: 7,
  accentHue: 250,
  logoInitials: 'TO'
};

// ---------------------------------------------------------------------------
// Shared setup
// ---------------------------------------------------------------------------

beforeEach(function () {
  // Reset ConfigReader and DataService before each test
  global.ConfigReader = {
    getConfig: jest.fn(function () { return Object.assign({}, defaultConfig); })
  };
  global.DataService = {
    findUserByEmail: jest.fn(function () {
      return { email: 'user@test.com', name: 'Test User' };
    })
  };

  // Reset Session to return a valid SSO email by default
  Session.getActiveUser = jest.fn(function () {
    return { getEmail: jest.fn(function () { return 'test@example.com'; }) };
  });

  // Reset MailApp
  MailApp.sendEmail = jest.fn();
  MailApp.getRemainingDailyQuota = jest.fn(function () { return 100; });

  // Reset Logger
  Logger.log = jest.fn();

  // Reset ScriptApp
  ScriptApp.getService = jest.fn(function () {
    return { getUrl: jest.fn(function () { return 'https://script.google.com/macros/s/test/exec'; }) };
  });
  ScriptApp.getProjectTriggers = jest.fn(function () { return []; });
  ScriptApp.deleteTrigger = jest.fn();
  ScriptApp.newTrigger = jest.fn(function () {
    return {
      timeBased: jest.fn(function () { return this; }),
      everyDays: jest.fn(function () { return this; }),
      atHour: jest.fn(function () { return this; }),
      create: jest.fn()
    };
  });
});

// ============================================================================
// Auth.resolveUser
// ============================================================================

describe('Auth.resolveUser', function () {

  test('returns null for empty/undefined event', function () {
    // Override SSO so no auth signal is present
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });
    expect(Auth.resolveUser(undefined)).toBeNull();
    expect(Auth.resolveUser(null)).toBeNull();
  });

  test('returns null when loggedout=1', function () {
    var result = Auth.resolveUser({ parameter: { loggedout: '1' } });
    expect(result).toBeNull();
  });

  test('returns {email, method:"sso"} when Session.getActiveUser() returns email', function () {
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return 'sso@example.com'; }) };
    });
    var result = Auth.resolveUser({ parameter: {} });
    expect(result).toEqual({ email: 'sso@example.com', method: 'sso' });
  });

  test('returns null when no auth signals are present', function () {
    // No SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });
    var result = Auth.resolveUser({ parameter: {} });
    expect(result).toBeNull();
  });

  test('returns {method:"session"} for valid session token', function () {
    // Store a valid session token
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_mytoken123', JSON.stringify({
      email: 'session@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now()
    }));
    // Disable SSO so session is checked first and we can confirm method
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { sessionToken: 'mytoken123' } });
    expect(result).not.toBeNull();
    expect(result.email).toBe('session@example.com');
    expect(result.method).toBe('session');
  });

  test('returns null for expired session token', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_expiredtoken', JSON.stringify({
      email: 'session@example.com',
      expiry: Date.now() - 1000,  // expired
      created: Date.now() - 7200000
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { sessionToken: 'expiredtoken' } });
    expect(result).toBeNull();
  });

  test('returns {method:"magic"} for valid magic token', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_abc123', JSON.stringify({
      email: 'magic@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: false
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'abc123' } });
    expect(result).not.toBeNull();
    expect(result.email).toBe('magic@example.com');
    expect(result.method).toBe('magic');
  });

  test('returns null for expired magic token', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_expmagic', JSON.stringify({
      email: 'magic@example.com',
      expiry: Date.now() - 1000,
      created: Date.now() - 7200000,
      used: false
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'expmagic' } });
    expect(result).toBeNull();
  });

  test('returns null for used (replayed) magic token', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_usedtoken', JSON.stringify({
      email: 'magic@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: true
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'usedtoken' } });
    expect(result).toBeNull();
  });

  test('email is lowercased on SSO', function () {
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return 'Admin@Example.COM'; }) };
    });
    var result = Auth.resolveUser({ parameter: {} });
    expect(result.email).toBe('admin@example.com');
  });

  test('priority: session > SSO > magic', function () {
    // All three signals present — session should win
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_sesstoken', JSON.stringify({
      email: 'session@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now()
    }));
    props.setProperty('MAGIC_TOKEN_magictoken', JSON.stringify({
      email: 'magic@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: false
    }));
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return 'sso@example.com'; }) };
    });

    var result = Auth.resolveUser({
      parameter: { sessionToken: 'sesstoken', token: 'magictoken' }
    });
    expect(result.method).toBe('session');
    expect(result.email).toBe('session@example.com');
  });

  test('handles missing e.parameter gracefully', function () {
    // event object with no parameter property
    var result = Auth.resolveUser({});
    // Should not throw — SSO is available, so it returns sso result
    expect(result).not.toBeNull();
    expect(result.method).toBe('sso');
  });

  test('falls through to SSO when session token is invalid (not in store)', function () {
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return 'sso@example.com'; }) };
    });
    var result = Auth.resolveUser({ parameter: { sessionToken: 'nonexistent' } });
    // Invalid session token falls through to SSO
    expect(result.method).toBe('sso');
    expect(result.email).toBe('sso@example.com');
  });

  test('falls through to magic when SSO throws', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_fallback', JSON.stringify({
      email: 'magic@example.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: false
    }));
    Session.getActiveUser = jest.fn(function () {
      throw new Error('SSO not available in anonymous context');
    });

    var result = Auth.resolveUser({ parameter: { token: 'fallback' } });
    expect(result.method).toBe('magic');
    expect(result.email).toBe('magic@example.com');
  });
});

// ============================================================================
// Auth.sendMagicLink
// ============================================================================

describe('Auth.sendMagicLink', function () {

  beforeEach(function () {
    // Reset email mocks between tests to prevent accumulated calls
    GmailApp.sendEmail = jest.fn();
    MailApp.sendEmail = jest.fn();
    MailApp.getRemainingDailyQuota = jest.fn(function () { return 100; });
  });

  test('returns success for valid email', function () {
    var result = Auth.sendMagicLink('user@test.com', false);
    expect(result.success).toBe(true);
    expect(result.message).toContain('Sign-in link sent');
  });

  test('lowercases and trims email', function () {
    Auth.sendMagicLink('  User@Test.COM  ', false);
    expect(DataService.findUserByEmail).toHaveBeenCalledWith('user@test.com');
  });

  test('returns same message for unknown email (no enumeration)', function () {
    DataService.findUserByEmail = jest.fn(function () { return null; });
    var result = Auth.sendMagicLink('unknown@nowhere.com', false);
    expect(result.success).toBe(true);
    expect(result.message).toContain('If this email is in our directory');
  });

  test('rate limits at 3 per email per 15 min', function () {
    // Make getScriptCache return a singleton so the source code sees count '3'
    var rateLimitCache = {
      get: jest.fn(function () { return '3'; }),
      put: jest.fn(),
      remove: jest.fn()
    };
    CacheService.getScriptCache.mockReturnValueOnce(rateLimitCache);

    var result = Auth.sendMagicLink('user@test.com', false);
    // Should silently succeed (no enumeration) but not send email
    expect(result.success).toBe(true);
    expect(result.message).toContain('If this email is in our directory');
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });

  test('handles quota exhaustion', function () {
    // Force GmailApp to fail so MailApp fallback is exercised
    GmailApp.sendEmail = jest.fn(function () { throw new Error('GmailApp unavailable'); });
    MailApp.getRemainingDailyQuota = jest.fn(function () { return 0; });
    var result = Auth.sendMagicLink('user@test.com', false);
    expect(result.success).toBe(false);
    expect(result.message).toContain('quota');
  });

  test('handles invalid email sendEmail error', function () {
    // Force both GmailApp and MailApp to fail
    GmailApp.sendEmail = jest.fn(function () { throw new Error('invalid email address'); });
    MailApp.sendEmail = jest.fn(function () {
      throw new Error('invalid email address');
    });
    var result = Auth.sendMagicLink('user@test.com', false);
    expect(result.success).toBe(false);
    expect(result.message).toContain('invalid');
  });

  test('handles generic sendEmail error', function () {
    // Force both GmailApp and MailApp to fail
    GmailApp.sendEmail = jest.fn(function () { throw new Error('Service unavailable'); });
    MailApp.sendEmail = jest.fn(function () {
      throw new Error('Service unavailable');
    });
    var result = Auth.sendMagicLink('user@test.com', false);
    expect(result.success).toBe(false);
    expect(result.message).toContain('try again');
  });

  test('creates remember link when rememberMe=true', function () {
    Auth.sendMagicLink('user@test.com', true);
    // GmailApp is the primary sender — args: (to, subject, plainBody, options)
    var emailOpts = GmailApp.sendEmail.mock.calls[0][3];
    expect(emailOpts.htmlBody).toContain('remember=1');
  });

  test('calls GmailApp.sendEmail with correct parameters', function () {
    Auth.sendMagicLink('user@test.com', false);
    expect(GmailApp.sendEmail).toHaveBeenCalledTimes(1);
    // GmailApp.sendEmail(to, subject, plainBody, options)
    expect(GmailApp.sendEmail.mock.calls[0][0]).toBe('user@test.com');
    expect(GmailApp.sendEmail.mock.calls[0][1]).toContain('Test Org');
    expect(GmailApp.sendEmail.mock.calls[0][3].noReply).toBe(false);
  });

  test('does not include remember param when rememberMe=false', function () {
    Auth.sendMagicLink('user@test.com', false);
    var emailOpts = GmailApp.sendEmail.mock.calls[0][3];
    expect(emailOpts.htmlBody).not.toContain('remember=1');
  });

  test('handles quota check failure gracefully and proceeds to send via MailApp fallback', function () {
    // Force GmailApp to fail so MailApp fallback is exercised
    GmailApp.sendEmail = jest.fn(function () { throw new Error('GmailApp unavailable'); });
    MailApp.getRemainingDailyQuota = jest.fn(function () {
      throw new Error('Quota check unavailable');
    });
    // MailApp sendEmail should still be called since quota check failure is caught (defaults to 1)
    var result = Auth.sendMagicLink('user@test.com', false);
    expect(MailApp.sendEmail).toHaveBeenCalledTimes(1);
    expect(result.success).toBe(true);
  });

  test('email body uses escapeHtml for config values', function () {
    global.ConfigReader = {
      getConfig: jest.fn(function () {
        return Object.assign({}, defaultConfig, {
          orgName: 'Org <script>alert("xss")</script>',
          logoInitials: '<b>XSS</b>'
        });
      })
    };
    Auth.sendMagicLink('user@test.com', false);
    // GmailApp is the primary sender
    var emailOpts = GmailApp.sendEmail.mock.calls[0][3];
    // HTML entities should be escaped — raw tags should not appear
    expect(emailOpts.htmlBody).not.toContain('<script>alert');
    expect(emailOpts.htmlBody).not.toContain('<b>XSS</b>');
  });
});

// ============================================================================
// Auth.createSessionToken
// ============================================================================

describe('Auth.createSessionToken', function () {

  test('returns a token string', function () {
    var token = Auth.createSessionToken('user@test.com');
    expect(typeof token).toBe('string');
    expect(token.length).toBeGreaterThan(0);
  });

  test('stores token in ScriptProperties', function () {
    var token = Auth.createSessionToken('user@test.com');
    var props = PropertiesService.getScriptProperties();
    var stored = props.getProperty('SESSION_' + token);
    expect(stored).not.toBeNull();
  });

  test('token is retrievable and session is valid', function () {
    var token = Auth.createSessionToken('user@test.com');
    var props = PropertiesService.getScriptProperties();
    var data = JSON.parse(props.getProperty('SESSION_' + token));
    expect(data.email).toBe('user@test.com');
  });

  test('stores email as lowercase', function () {
    var token = Auth.createSessionToken('Admin@EXAMPLE.com');
    var props = PropertiesService.getScriptProperties();
    var data = JSON.parse(props.getProperty('SESSION_' + token));
    expect(data.email).toBe('admin@example.com');
  });

  test('stores expiry in the future', function () {
    var before = Date.now();
    var token = Auth.createSessionToken('user@test.com');
    var props = PropertiesService.getScriptProperties();
    var data = JSON.parse(props.getProperty('SESSION_' + token));
    expect(data.expiry).toBeGreaterThan(before);
    // Expiry should be ~30 days in the future (cookieDurationMs = 2592000000)
    expect(data.expiry).toBeGreaterThanOrEqual(before + defaultConfig.cookieDurationMs);
  });
});

// ============================================================================
// Auth.invalidateSession
// ============================================================================

describe('Auth.invalidateSession', function () {

  test('deletes session property', function () {
    // First create a session
    var token = Auth.createSessionToken('user@test.com');
    var props = PropertiesService.getScriptProperties();
    expect(props.getProperty('SESSION_' + token)).not.toBeNull();

    // Invalidate it
    Auth.invalidateSession(token);
    expect(props.getProperty('SESSION_' + token)).toBeNull();
  });

  test('does nothing for null token', function () {
    // Should not throw
    expect(function () {
      Auth.invalidateSession(null);
    }).not.toThrow();
  });

  test('makes token no longer valid after invalidation', function () {
    var token = Auth.createSessionToken('user@test.com');
    Auth.invalidateSession(token);

    // Disable SSO so resolveUser relies on session
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { sessionToken: token } });
    expect(result).toBeNull();
  });
});

// ============================================================================
// Auth.cleanupExpiredTokens
// ============================================================================

describe('Auth.cleanupExpiredTokens', function () {

  test('returns 0 for empty properties', function () {
    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(0);
  });

  test('cleans expired magic tokens', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_old', JSON.stringify({
      email: 'old@test.com',
      expiry: Date.now() - 100000,
      created: Date.now() - 200000,
      used: false
    }));

    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(1);
    expect(props.getProperty('MAGIC_TOKEN_old')).toBeNull();
  });

  test('cleans expired session tokens', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_old', JSON.stringify({
      email: 'old@test.com',
      expiry: Date.now() - 100000,
      created: Date.now() - 200000
    }));

    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(1);
    expect(props.getProperty('SESSION_old')).toBeNull();
  });

  test('keeps non-expired tokens', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('SESSION_valid', JSON.stringify({
      email: 'valid@test.com',
      expiry: Date.now() + 9999999,
      created: Date.now()
    }));

    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(0);
    expect(props.getProperty('SESSION_valid')).not.toBeNull();
  });

  test('cleans malformed JSON entries', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_broken', '{not valid json!!!');

    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(1);
    expect(props.getProperty('MAGIC_TOKEN_broken')).toBeNull();
  });

  test('returns count of cleaned tokens', function () {
    var props = PropertiesService.getScriptProperties();
    // 2 expired + 1 malformed = 3 cleaned
    props.setProperty('MAGIC_TOKEN_exp1', JSON.stringify({
      email: 'a@test.com',
      expiry: Date.now() - 1000,
      created: Date.now() - 5000,
      used: false
    }));
    props.setProperty('SESSION_exp2', JSON.stringify({
      email: 'b@test.com',
      expiry: Date.now() - 2000,
      created: Date.now() - 6000
    }));
    props.setProperty('MAGIC_TOKEN_bad', 'not-json');
    // 1 valid — should not be counted
    props.setProperty('SESSION_good', JSON.stringify({
      email: 'c@test.com',
      expiry: Date.now() + 9999999,
      created: Date.now()
    }));

    var result = Auth.cleanupExpiredTokens();
    expect(result).toBe(3);
  });
});

// ============================================================================
// Global wrapper functions
// ============================================================================

describe('Global wrappers', function () {

  test('authSendMagicLink delegates to Auth.sendMagicLink', function () {
    var result = authSendMagicLink('user@test.com', false);
    expect(result.success).toBe(true);
    expect(result.message).toContain('Sign-in link sent');
  });

  test('authCreateSessionToken resolves email from Session', function () {
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return 'sso@example.com'; }) };
    });
    var token = authCreateSessionToken();
    expect(typeof token).toBe('string');
    expect(token.length).toBeGreaterThan(0);
  });

  test('authCreateSessionToken returns error when SSO unavailable', function () {
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });
    var result = authCreateSessionToken();
    expect(result.error).toBeDefined();
    expect(result.error).toContain('Unable to resolve');
  });

  test('authCreateSessionToken returns error when getActiveUser throws', function () {
    Session.getActiveUser = jest.fn(function () {
      throw new Error('SSO not available');
    });
    var result = authCreateSessionToken();
    expect(result.error).toBeDefined();
    expect(result.error).toContain('Unable to resolve');
  });

  test('authLogout calls Auth.invalidateSession', function () {
    var token = Auth.createSessionToken('user@test.com');
    var result = authLogout(token);
    expect(result.success).toBe(true);
    // Token should be invalidated
    var props = PropertiesService.getScriptProperties();
    expect(props.getProperty('SESSION_' + token)).toBeNull();
  });

  test('authCleanupExpiredTokens delegates to Auth.cleanupExpiredTokens', function () {
    var result = authCleanupExpiredTokens();
    expect(typeof result).toBe('number');
  });

});

// ============================================================================
// Magic token security
// ============================================================================

describe('Magic token security', function () {

  test('used token is rejected (replay prevention)', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_replay', JSON.stringify({
      email: 'user@test.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: true
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'replay' } });
    expect(result).toBeNull();
  });

  test('token is deleted after first validate (C1+C4: immediate delete)', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_markused', JSON.stringify({
      email: 'user@test.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: false
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    // First use should succeed
    var result = Auth.resolveUser({ parameter: { token: 'markused' } });
    expect(result).not.toBeNull();
    expect(result.email).toBe('user@test.com');

    // v4.31.0 C1+C4: Token should be DELETED (not marked used)
    var stored = props.getProperty('MAGIC_TOKEN_markused');
    expect(stored).toBeNull();
  });

  test('expired token is deleted from properties', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_expiredel', JSON.stringify({
      email: 'user@test.com',
      expiry: Date.now() - 1000,
      created: Date.now() - 100000,
      used: false
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    Auth.resolveUser({ parameter: { token: 'expiredel' } });
    expect(props.getProperty('MAGIC_TOKEN_expiredel')).toBeNull();
  });

  test('invalid JSON in token returns null', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_badjson', 'this is not json');
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'badjson' } });
    expect(result).toBeNull();
  });

  test('token stores created timestamp', function () {
    var before = Date.now();
    Auth.sendMagicLink('user@test.com', false);

    // Find the magic token that was created
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var magicKeys = Object.keys(allProps).filter(function (k) {
      return k.indexOf('MAGIC_TOKEN_') === 0;
    });

    expect(magicKeys.length).toBeGreaterThan(0);
    var data = JSON.parse(allProps[magicKeys[0]]);
    expect(data.created).toBeDefined();
    expect(data.created).toBeGreaterThanOrEqual(before);
    expect(data.created).toBeLessThanOrEqual(Date.now());
  });
});

// ============================================================================
// v4.31.0 — C1+C4: Token validation deletes token (not set-used)
// ============================================================================

describe('v4.31.0 C1+C4: Token validation uses deleteProperty', function () {

  test('magic token is deleted (not marked used) after successful validation', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_deltest', JSON.stringify({
      email: 'user@test.com',
      expiry: Date.now() + 3600000,
      created: Date.now(),
      used: false
    }));
    // Disable SSO
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    var result = Auth.resolveUser({ parameter: { token: 'deltest' } });
    expect(result).not.toBeNull();
    expect(result.email).toBe('user@test.com');

    // Token should be DELETED, not still present with used=true
    var stored = props.getProperty('MAGIC_TOKEN_deltest');
    expect(stored).toBeNull();
  });

  test('expired magic token is deleted on validation attempt', function () {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MAGIC_TOKEN_expdeltest', JSON.stringify({
      email: 'user@test.com',
      expiry: Date.now() - 1000,
      created: Date.now() - 100000,
      used: false
    }));
    Session.getActiveUser = jest.fn(function () {
      return { getEmail: jest.fn(function () { return ''; }) };
    });

    Auth.resolveUser({ parameter: { token: 'expdeltest' } });
    expect(props.getProperty('MAGIC_TOKEN_expdeltest')).toBeNull();
  });
});

// ============================================================================
// v4.31.0 — C3: createSessionToken returns error on storage failure
// ============================================================================

describe('v4.31.0 C3: createSessionToken error on PropertiesService write failure', function () {

  test('returns error object when PropertiesService write fails', function () {
    // Make setProperty throw to simulate quota exceeded
    var origGetScriptProps = PropertiesService.getScriptProperties;
    PropertiesService.getScriptProperties = jest.fn(function () {
      return {
        getProperty: jest.fn(function () { return null; }),
        setProperty: jest.fn(function () { throw new Error('Quota exceeded'); }),
        deleteProperty: jest.fn(),
        getProperties: jest.fn(function () { return {}; })
      };
    });

    var result = Auth.createSessionToken('user@test.com');
    expect(result).toBeDefined();
    expect(result.error).toBeDefined();
    expect(result.message).toBeDefined();

    // Restore
    PropertiesService.getScriptProperties = origGetScriptProps;
  });
});

// ============================================================================
// v4.31.0 — M9: cleanupExpiredTokens calls recordSecurityEvent on quota warning
// ============================================================================

describe('v4.31.0 M9: cleanupExpiredTokens quota escalation', function () {

  test('calls recordSecurityEvent when quota exceeds 400KB', function () {
    // Mock recordSecurityEvent
    global.recordSecurityEvent = jest.fn();

    // Populate ScriptProperties with enough data to exceed 400KB
    var props = PropertiesService.getScriptProperties();
    // Create a large payload that will exceed 400KB when serialized
    var bigValue = new Array(50001).join('x'); // ~50KB string
    for (var i = 0; i < 10; i++) {
      props.setProperty('PADDING_' + i, bigValue);
    }

    Auth.cleanupExpiredTokens();

    expect(global.recordSecurityEvent).toHaveBeenCalled();
    var call = global.recordSecurityEvent.mock.calls[0];
    expect(call[0]).toBe('QUOTA_WARNING');

    // Cleanup
    delete global.recordSecurityEvent;
  });
});
