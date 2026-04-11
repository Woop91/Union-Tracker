/**
 * Auditor-Alpha AA-04 / AA-05 coverage: _resolveCallerEmail and _requireStewardAuth.
 *
 * Pre-v4.55.2 the full test suite stubbed both of these functions via
 * jest.spyOn().mockReturnValue(...) in 21_WebDashDataService.test.js and
 * auth-denial.test.js. That made every wrapper test verify DELEGATION to
 * the auth helpers instead of the helpers' actual behavior. Alpha AA-04
 * and AA-05 flagged this as a silent-pass risk: a regression that
 * lowercased email incorrectly, returned an unexpected type, or bypassed
 * the session-token path would have zero failing tests.
 *
 * This file loads the real implementations (NO global mocks of these
 * helpers) and exercises:
 *
 *   - session token priority (token wins over Session.getActiveUser)
 *   - SSO fallback when no token is supplied
 *   - SSO fallback when token resolves to empty string
 *   - lowercase normalization and whitespace trimming
 *   - graceful null when nothing resolves
 *   - _requireStewardAuth early-return on missing token
 *   - _requireStewardAuth delegation to checkWebAppAuthorization
 *   - _requireStewardAuth returns lowercased email on success
 *   - _requireStewardAuth returns null on auth failure
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load in dependency order. 21d_WebDashDataWrappers defines our targets;
// 00_Security is where checkWebAppAuthorization lives.
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '21d_WebDashDataWrappers.gs'
]);

// ============================================================================
// _resolveCallerEmail (AA-04)
// ============================================================================

describe('_resolveCallerEmail (AA-04)', () => {
  var _origAuth, _origSessionGetActiveUser, _origLog;

  beforeEach(() => {
    _origAuth = global.Auth;
    _origSessionGetActiveUser = global.Session.getActiveUser;
    _origLog = global.log_;
    global.log_ = jest.fn();
  });

  afterEach(() => {
    global.Auth = _origAuth;
    global.Session.getActiveUser = _origSessionGetActiveUser;
    global.log_ = _origLog;
  });

  test('session token path: resolves via Auth.resolveEmailFromToken', () => {
    global.Auth = {
      resolveEmailFromToken: jest.fn(() => 'Token.User@Example.COM')
    };
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => 'should-not-see@example.com')
    }));
    var email = _resolveCallerEmail('abc123');
    expect(email).toBe('token.user@example.com');
    expect(global.Auth.resolveEmailFromToken).toHaveBeenCalledWith('abc123');
  });

  test('token path takes priority over Session active user', () => {
    global.Auth = {
      resolveEmailFromToken: jest.fn(() => 'priority@example.com')
    };
    var activeSpy = jest.fn(() => ({
      getEmail: jest.fn(() => 'secondary@example.com')
    }));
    global.Session.getActiveUser = activeSpy;
    var email = _resolveCallerEmail('abc123');
    expect(email).toBe('priority@example.com');
    expect(activeSpy).not.toHaveBeenCalled();
  });

  test('falls back to Session active user when no token supplied', () => {
    global.Auth = { resolveEmailFromToken: jest.fn() };
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => 'sso@example.com')
    }));
    var email = _resolveCallerEmail(null);
    expect(email).toBe('sso@example.com');
    expect(global.Auth.resolveEmailFromToken).not.toHaveBeenCalled();
  });

  test('falls back to Session when token resolves to empty string', () => {
    global.Auth = { resolveEmailFromToken: jest.fn(() => '') };
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => 'fallback@example.com')
    }));
    var email = _resolveCallerEmail('bad-token');
    expect(email).toBe('fallback@example.com');
  });

  test('returns null when neither path yields an email', () => {
    global.Auth = { resolveEmailFromToken: jest.fn(() => null) };
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => '')
    }));
    expect(_resolveCallerEmail('bad')).toBeNull();
  });

  test('catches exceptions from Session.getActiveUser and logs', () => {
    global.Auth = undefined;
    global.Session.getActiveUser = jest.fn(() => {
      throw new Error('Session unavailable');
    });
    expect(_resolveCallerEmail(null)).toBeNull();
    expect(global.log_).toHaveBeenCalled();
  });

  test('lowercases and trims the returned email', () => {
    global.Auth = {
      resolveEmailFromToken: jest.fn(() => '  Mixed.Case@EXAMPLE.com  ')
    };
    expect(_resolveCallerEmail('t')).toBe('mixed.case@example.com');
  });

  test('works when Auth global is undefined (early boot)', () => {
    global.Auth = undefined;
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => 'sso@example.com')
    }));
    // Token is supplied but Auth is undefined — should still fall back
    expect(_resolveCallerEmail('t')).toBe('sso@example.com');
  });
});

// ============================================================================
// _requireStewardAuth (AA-05)
// ============================================================================

describe('_requireStewardAuth (AA-05)', () => {
  var _origCheck;

  beforeEach(() => {
    _origCheck = global.checkWebAppAuthorization;
  });

  afterEach(() => {
    global.checkWebAppAuthorization = _origCheck;
  });

  test('returns null when no sessionToken supplied', () => {
    // No mock needed — early return before calling checkWebAppAuthorization
    expect(_requireStewardAuth(null)).toBeNull();
    expect(_requireStewardAuth('')).toBeNull();
    expect(_requireStewardAuth(undefined)).toBeNull();
  });

  test('delegates to checkWebAppAuthorization with steward role', () => {
    var spy = jest.fn(() => ({ isAuthorized: true, email: 'Steward@Example.COM' }));
    global.checkWebAppAuthorization = spy;
    var result = _requireStewardAuth('token-xyz');
    expect(spy).toHaveBeenCalledWith('steward', 'token-xyz');
    expect(result).toBe('steward@example.com');
  });

  test('returns null when checkWebAppAuthorization denies', () => {
    global.checkWebAppAuthorization = jest.fn(() => ({
      isAuthorized: false,
      message: 'Steward access required'
    }));
    expect(_requireStewardAuth('member-token')).toBeNull();
  });

  test('returns lowercased + trimmed email on success', () => {
    global.checkWebAppAuthorization = jest.fn(() => ({
      isAuthorized: true,
      email: '   STEWARD@Example.COM   '
    }));
    expect(_requireStewardAuth('t')).toBe('steward@example.com');
  });

  test('returns empty string (falsy) when authorized but email missing', () => {
    // Edge case: checkWebAppAuthorization returns isAuthorized:true but no email.
    // Current implementation returns '' — a falsy but non-null value. Callers
    // treat both as denial because they use if (!steward). Verify the current
    // contract explicitly so a regression that returns {} or undefined would
    // fail this test.
    global.checkWebAppAuthorization = jest.fn(() => ({
      isAuthorized: true,
      email: null
    }));
    var result = _requireStewardAuth('t');
    expect(result).toBe('');
  });
});
