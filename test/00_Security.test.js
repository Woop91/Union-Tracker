/**
 * Tests for 00_Security.gs
 *
 * Covers XSS prevention, formula injection, PII masking, input validation,
 * and client-side security helper generation.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs']);

// ============================================================================
// escapeHtml
// ============================================================================

describe('escapeHtml', () => {
  test('escapes HTML special characters', () => {
    expect(escapeHtml('<script>alert("xss")</script>')).toBe(
      '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
    );
  });

  test('escapes single quotes', () => {
    expect(escapeHtml("it's")).toBe("it&#x27;s");
  });

  test('escapes backticks', () => {
    expect(escapeHtml('`code`')).toBe('&#x60;code&#x60;');
  });

  test('does not escape equals sign (not an XSS vector)', () => {
    expect(escapeHtml('a=b')).toBe('a=b');
  });

  test('does not escape forward slash (not an XSS vector)', () => {
    expect(escapeHtml('a/b')).toBe('a/b');
  });

  test('escapes ampersand', () => {
    expect(escapeHtml('a&b')).toBe('a&amp;b');
  });

  test('returns empty string for null', () => {
    expect(escapeHtml(null)).toBe('');
  });

  test('returns empty string for undefined', () => {
    expect(escapeHtml(undefined)).toBe('');
  });

  test('converts non-string input to string', () => {
    expect(escapeHtml(123)).toBe('123');
  });

  test('handles empty string', () => {
    expect(escapeHtml('')).toBe('');
  });

  test('handles combined attack vectors', () => {
    const input = '<img src=x onerror="alert(1)">';
    const result = escapeHtml(input);
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result).not.toContain('"');
  });

  test('escapes input with all special chars combined', () => {
    const input = '<div class="test">&\'`';
    const result = escapeHtml(input);
    // Raw dangerous chars should not appear unescaped
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result).not.toContain('"');
    expect(result).not.toContain("'");
    expect(result).not.toContain('`');
    // Verify each entity is present in the output
    expect(result).toContain('&lt;');
    expect(result).toContain('&gt;');
    expect(result).toContain('&quot;');
    expect(result).toContain('&amp;');
    expect(result).toContain('&#x27;');
    expect(result).toContain('&#x60;');
  });

  test('handles very long input string', () => {
    const longStr = '<script>'.repeat(1000);
    const result = escapeHtml(longStr);
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result.length).toBeGreaterThan(longStr.length);
  });

  test('handles input containing only special chars', () => {
    const input = '<>&"\'`';
    const result = escapeHtml(input);
    expect(result).toBe('&lt;&gt;&amp;&quot;&#x27;&#x60;');
  });
});

// sanitizeObjectForHtml — removed (zero callers in production code)

// ============================================================================
// escapeForFormula
// ============================================================================

describe('escapeForFormula', () => {
  test('returns empty string for null', () => {
    expect(escapeForFormula(null)).toBe('');
  });

  test('prefixes formula-starting characters at start of string', () => {
    expect(escapeForFormula('=SUM(A1)')).toBe("'=SUM(A1)");
    expect(escapeForFormula('+cmd')).toBe("'+cmd");
    expect(escapeForFormula('-cmd')).toBe("'-cmd");
    expect(escapeForFormula('@import')).toBe("'@import");
  });

  test('does NOT prefix formula characters mid-string (bug fix verification)', () => {
    expect(escapeForFormula('user@email.com')).toBe('user@email.com');
    expect(escapeForFormula('a+b')).toBe('a+b');
    expect(escapeForFormula('a-b')).toBe('a-b');
  });

  test('preserves single quotes (no SQL-style doubling)', () => {
    expect(escapeForFormula("it's")).toBe("it's");
  });

  test('preserves double quotes (no SQL-style doubling)', () => {
    expect(escapeForFormula('say "hi"')).toBe('say "hi"');
  });

  test('replaces newlines with spaces', () => {
    expect(escapeForFormula('line1\nline2')).toBe('line1 line2');
  });

  test('handles normal text without modification', () => {
    expect(escapeForFormula('Member Directory')).toBe('Member Directory');
  });
});

// ============================================================================
// safeSheetNameForFormula
// ============================================================================

describe('safeSheetNameForFormula', () => {
  test('returns empty string for empty input', () => {
    expect(safeSheetNameForFormula('')).toBe('');
  });

  test('wraps names with special characters in quotes', () => {
    expect(safeSheetNameForFormula('Member Directory')).toBe("'Member Directory'");
  });

  test('does not wrap simple names', () => {
    expect(safeSheetNameForFormula('Config')).toBe('Config');
  });
});

// ============================================================================
// PII Masking
// ============================================================================

describe('maskEmail', () => {
  test('masks email correctly', () => {
    expect(maskEmail('john@example.com')).toBe('j***n@example.com');
  });

  test('returns placeholder for non-string', () => {
    expect(maskEmail(null)).toBe('[no email]');
  });
});

describe('maskPhone', () => {
  test('masks phone number', () => {
    expect(maskPhone('555-123-4567')).toBe('***-***-4567');
  });

  test('returns placeholder for non-string', () => {
    expect(maskPhone(null)).toBe('[no phone]');
  });
});

describe('maskName', () => {
  test('masks name to initials', () => {
    expect(maskName('John', 'Smith')).toBe('J. S.');
  });

  test('handles missing parts', () => {
    expect(maskName('', '')).toBe('[anonymous]');
  });
});

// ============================================================================
// Input Validation
// ============================================================================

describe('isValidSafeString', () => {
  test('accepts normal strings', () => {
    expect(isValidSafeString('Hello World')).toBe(true);
  });

  test('rejects script tags', () => {
    expect(isValidSafeString('<script>alert(1)</script>')).toBe(false);
  });

  test('rejects event handlers', () => {
    expect(isValidSafeString('onerror=alert(1)')).toBe(false);
  });

  test('rejects strings over max length', () => {
    expect(isValidSafeString('a'.repeat(1001))).toBe(false);
  });

  test('rejects null/undefined', () => {
    expect(isValidSafeString(null)).toBe(false);
    expect(isValidSafeString(undefined)).toBe(false);
  });
});

describe('isValidGrievanceId', () => {
  test('accepts valid grievance IDs', () => {
    expect(isValidGrievanceId('GRV-001')).toBe(true);
  });

  test('rejects invalid IDs', () => {
    expect(isValidGrievanceId('')).toBe(false);
  });
});

// ============================================================================
// Client-Side Security Helpers
// ============================================================================

describe('getClientSideEscapeHtml', () => {
  test('returns JavaScript code with all 6 escape replacements', () => {
    const code = getClientSideEscapeHtml();
    expect(code).toContain('function escapeHtml');
    expect(code).toContain('&amp;');
    expect(code).toContain('&lt;');
    expect(code).toContain('&gt;');
    expect(code).toContain('&quot;');
    expect(code).toContain('&#x27;');
    expect(code).toContain('&#x60;');
    // / and = are NOT escaped (F107: not XSS vectors, caused data corruption)
    expect(code).not.toContain('&#x2F;');
    expect(code).not.toContain('&#x3D;');
  });
});

// ============================================================================
// ACCESS_CONTROL
// ============================================================================

describe('ACCESS_CONTROL', () => {
  test('has required mode and page lists', () => {
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('steward');
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('member');
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('selfservice');
    expect(ACCESS_CONTROL.ALLOWED_PAGES).toContain('dashboard');
    expect(ACCESS_CONTROL.ALLOWED_PAGES).toContain('portal');
  });
});

// ============================================================================
// SECURITY_SEVERITY
// ============================================================================

describe('SECURITY_SEVERITY', () => {
  test('defines all severity levels', () => {
    expect(SECURITY_SEVERITY.CRITICAL).toBe('CRITICAL');
    expect(SECURITY_SEVERITY.HIGH).toBe('HIGH');
    expect(SECURITY_SEVERITY.MEDIUM).toBe('MEDIUM');
    expect(SECURITY_SEVERITY.LOW).toBe('LOW');
  });
});

// ============================================================================
// recordSecurityEvent
// ============================================================================

describe('recordSecurityEvent', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('is a function', () => {
    expect(typeof recordSecurityEvent).toBe('function');
  });

  test('logs audit event for any severity', () => {
    recordSecurityEvent('TEST_EVENT', SECURITY_SEVERITY.LOW, 'Test description', { foo: 'bar' });
    expect(logAuditEvent).toHaveBeenCalledWith('SECURITY_TEST_EVENT', expect.objectContaining({
      _severity: 'LOW',
      _description: 'Test description',
      foo: 'bar'
    }));
  });

  test('calls sendSecurityAlertEmail_ for CRITICAL events', () => {
    // MailApp.sendEmail is mocked — just verify it doesn't throw
    expect(() => {
      recordSecurityEvent('CRITICAL_TEST', SECURITY_SEVERITY.CRITICAL, 'Critical event', {});
    }).not.toThrow();
  });

  test('queues HIGH events for digest', () => {
    recordSecurityEvent('HIGH_TEST', SECURITY_SEVERITY.HIGH, 'High event', { detail: 'x' });

    // Verify event was queued in properties
    const props = PropertiesService.getScriptProperties();
    const queue = JSON.parse(props.getProperty('SECURITY_DIGEST_QUEUE') || '[]');
    expect(queue.length).toBeGreaterThan(0);
    expect(queue[queue.length - 1].event).toBe('HIGH_TEST');
  });

  test('does not queue LOW or MEDIUM events for digest', () => {
    // Clear any existing queue
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SECURITY_DIGEST_QUEUE', '[]');

    recordSecurityEvent('LOW_TEST', SECURITY_SEVERITY.LOW, 'Low event', {});
    recordSecurityEvent('MED_TEST', SECURITY_SEVERITY.MEDIUM, 'Medium event', {});

    const queue = JSON.parse(props.getProperty('SECURITY_DIGEST_QUEUE') || '[]');
    expect(queue.length).toBe(0);
  });
});

// ============================================================================
// queueSecurityDigestEvent_ and sendDailySecurityDigest
// ============================================================================

describe('sendDailySecurityDigest', () => {
  test('is a function', () => {
    expect(typeof sendDailySecurityDigest).toBe('function');
  });

  test('does not throw when queue is empty', () => {
    expect(() => sendDailySecurityDigest()).not.toThrow();
  });

  test('does not send email when queue is empty', () => {
    MailApp.sendEmail.mockClear();
    sendDailySecurityDigest();
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });
});

// ============================================================================
// escapeForFormulaPreserveNewlines
// ============================================================================

describe('escapeForFormulaPreserveNewlines', () => {
  test('preserves newlines in multiline text', () => {
    expect(escapeForFormulaPreserveNewlines('line 1\nline 2\nline 3'))
      .toBe('line 1\nline 2\nline 3');
  });

  test('escapes formula chars at start of each line', () => {
    expect(escapeForFormulaPreserveNewlines('=SUM(A1)\n+B2\n-C3\n@D4'))
      .toBe("'=SUM(A1)\n'+B2\n'-C3\n'@D4");
  });

  test('escapes formula chars after leading whitespace on each line', () => {
    expect(escapeForFormulaPreserveNewlines('  =IMPORTRANGE()\n  +123'))
      .toBe("'  =IMPORTRANGE()\n'  +123");
  });

  test('strips carriage returns but preserves line feeds', () => {
    expect(escapeForFormulaPreserveNewlines('line 1\r\nline 2\rline 3'))
      .toBe('line 1\nline 2\nline 3');
  });

  test('replaces tabs with spaces per line', () => {
    expect(escapeForFormulaPreserveNewlines('col1\tcol2\ncol3\tcol4'))
      .toBe('col1 col2\ncol3 col4');
  });

  test('escapes backslashes per line', () => {
    expect(escapeForFormulaPreserveNewlines('path\\to\\file\nsecond\\line'))
      .toBe('path\\\\to\\\\file\nsecond\\\\line');
  });

  test('returns empty string for null/undefined', () => {
    expect(escapeForFormulaPreserveNewlines(null)).toBe('');
    expect(escapeForFormulaPreserveNewlines(undefined)).toBe('');
  });

  test('handles empty string', () => {
    expect(escapeForFormulaPreserveNewlines('')).toBe('');
  });

  test('handles single line without newlines', () => {
    expect(escapeForFormulaPreserveNewlines('Normal text')).toBe('Normal text');
  });

  test('handles bullet-formatted text (real-world use case)', () => {
    var input = '• Budget approved\n• 15 stewards attended\n• Next meeting May 4';
    expect(escapeForFormulaPreserveNewlines(input)).toBe(input);
  });
});

// ============================================================================
// checkWebAppAuthorization — Auditor-Alpha AA-01
// ============================================================================
//
// The central RBAC gate for every webapp endpoint had zero dedicated unit
// tests in the pre-v4.55.2 suite. Every test file that referenced it
// mocked it to always return { isAuthorized: true } or { isAuthorized: false },
// so the real role hierarchy, fail-secure behavior, owner-is-admin detection,
// and error-path swallowing were unverified.
//
// These tests exercise the actual implementation. We stub external globals
// (Auth, DataService, ACCESS_CONTROL, SHEETS, log_) at the top of each test
// and restore on teardown so the rest of the file stays unaffected.
// ============================================================================

describe('checkWebAppAuthorization (AA-01)', () => {
  var _origAccessControl, _origAuth, _origDataService, _origSHEETS, _origLog;
  var _origSessionGetActiveUser, _origSSgetActiveSpreadsheet;
  var _origCache;

  beforeEach(() => {
    _origAccessControl = global.ACCESS_CONTROL;
    _origAuth = global.Auth;
    _origDataService = global.DataService;
    _origSHEETS = global.SHEETS;
    _origLog = global.log_;
    _origSessionGetActiveUser = global.Session.getActiveUser;
    _origSSgetActiveSpreadsheet = global.SpreadsheetApp.getActiveSpreadsheet;
    _origCache = global.CacheService.getScriptCache;

    global.ACCESS_CONTROL = { ENABLED: true, AUTH_CACHE_DURATION: 300 };
    global.SHEETS = { MEMBER_DIR: 'Member Directory' };
    global.log_ = jest.fn();
    global.CacheService.getScriptCache = jest.fn(() => ({
      get: jest.fn(() => null),
      put: jest.fn(),
      remove: jest.fn()
    }));
  });

  afterEach(() => {
    global.ACCESS_CONTROL = _origAccessControl;
    global.Auth = _origAuth;
    global.DataService = _origDataService;
    global.SHEETS = _origSHEETS;
    global.log_ = _origLog;
    global.Session.getActiveUser = _origSessionGetActiveUser;
    global.SpreadsheetApp.getActiveSpreadsheet = _origSSgetActiveSpreadsheet;
    global.CacheService.getScriptCache = _origCache;
  });

  function mockOwner(email) {
    global.SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
      getOwner: jest.fn(() => ({ getEmail: jest.fn(() => email) })),
      getSheetByName: jest.fn(() => null)
    }));
  }

  function mockActiveUser(email) {
    global.Session.getActiveUser = jest.fn(() => ({
      getEmail: jest.fn(() => email)
    }));
  }

  function mockDataServiceRole(role) {
    global.DataService = { getUserRole: jest.fn(() => role) };
  }

  test('returns Authentication required when no email and no token', () => {
    mockActiveUser('');
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(false);
    expect(res.message).toMatch(/Authentication required/i);
    expect(res.email).toBeNull();
    expect(res.role).toBe('anonymous');
  });

  test('fail-secure: denies when ACCESS_CONTROL.ENABLED is false', () => {
    mockActiveUser('alice@example.com');
    mockOwner('admin@example.com');
    global.ACCESS_CONTROL = { ENABLED: false, AUTH_CACHE_DURATION: 300 };
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(false);
    expect(res.message).toMatch(/Access control is currently disabled/i);
    expect(res.email).toBe('alice@example.com');
    expect(global.log_).toHaveBeenCalled();
  });

  test('spreadsheet owner is treated as admin regardless of directory role', () => {
    mockActiveUser('owner@example.com');
    mockOwner('owner@example.com');
    var res = checkWebAppAuthorization('admin', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('admin');
  });

  test('admin role satisfies requiredRole=admin', () => {
    mockActiveUser('admin@example.com');
    mockOwner('admin@example.com'); // owner → admin path
    var res = checkWebAppAuthorization('admin', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('admin');
  });

  test('member role denied for requiredRole=admin with clear message', () => {
    mockActiveUser('member@example.com');
    mockOwner('someone-else@example.com');
    mockDataServiceRole('member');
    var res = checkWebAppAuthorization('admin', null);
    expect(res.isAuthorized).toBe(false);
    expect(res.message).toMatch(/Administrator access required/i);
    expect(res.role).toBe('member');
  });

  test('steward role satisfies requiredRole=steward', () => {
    mockActiveUser('steward@example.com');
    mockOwner('someone-else@example.com');
    mockDataServiceRole('steward');
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('steward');
  });

  test('member role denied for requiredRole=steward with clear message', () => {
    mockActiveUser('member@example.com');
    mockOwner('someone-else@example.com');
    mockDataServiceRole('member');
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(false);
    expect(res.message).toMatch(/Steward access required/i);
  });

  test('admin role can perform steward actions (hierarchy)', () => {
    mockActiveUser('admin@example.com');
    mockOwner('admin@example.com'); // owner → admin
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('admin');
  });

  test('dual-role user (both) satisfies requiredRole=steward', () => {
    mockActiveUser('dual@example.com');
    mockOwner('owner@example.com');
    mockDataServiceRole('both');
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('both');
  });

  test('no requiredRole: any authenticated user passes', () => {
    mockActiveUser('anyone@example.com');
    mockOwner('owner@example.com');
    mockDataServiceRole('member');
    var res = checkWebAppAuthorization(null, null);
    expect(res.isAuthorized).toBe(true);
  });

  test('session token path: resolves email via Auth.resolveEmailFromToken', () => {
    mockOwner('owner@example.com');
    mockDataServiceRole('steward');
    global.Auth = {
      resolveEmailFromToken: jest.fn(() => 'tokenuser@example.com')
    };
    // Deliberately no active user — test that token is used
    mockActiveUser('');
    var res = checkWebAppAuthorization('steward', 'fake-token');
    expect(global.Auth.resolveEmailFromToken).toHaveBeenCalledWith('fake-token');
    expect(res.email).toBe('tokenuser@example.com');
    expect(res.isAuthorized).toBe(true);
  });

  test('error path: catches exceptions and returns generic message (no leak)', () => {
    // Force getActiveUser to throw
    global.Session.getActiveUser = jest.fn(() => {
      throw new Error('sheet "Member Directory!A1:Z" error with stack trace');
    });
    var res = checkWebAppAuthorization('steward', null);
    expect(res.isAuthorized).toBe(false);
    expect(res.message).toMatch(/Authorization check failed/i);
    // Critical: internal details MUST NOT be in the returned message
    expect(res.message).not.toContain('Member Directory');
    expect(res.message).not.toContain('sheet');
    expect(global.log_).toHaveBeenCalled();
  });

  test('owner email comparison is case-insensitive', () => {
    mockActiveUser('Owner@Example.COM');
    mockOwner('owner@example.com');
    var res = checkWebAppAuthorization('admin', null);
    expect(res.isAuthorized).toBe(true);
    expect(res.role).toBe('admin');
  });
});

// ============================================================================
// getUserRole_ — Auditor-Alpha AA-02 / AA-11
// ============================================================================
//
// AA-02: getUserRole_ is never tested against real sheet data.
// AA-11: getUserRole_ CacheService caching is untested — no verification that
// a cache hit short-circuits sheet access, no verification that a cache miss
// populates the cache, no verification of cache-failure fallthrough.
// ============================================================================

describe('getUserRole_ cache behavior (AA-02/AA-11)', () => {
  var _origAccessControl, _origDataService, _origSHEETS, _origLog;
  var _origCacheGetScriptCache, _origSSgetActiveSpreadsheet, _origMemberCols;

  var mockCache;
  var cachePut;
  var cacheGet;

  beforeEach(() => {
    _origAccessControl = global.ACCESS_CONTROL;
    _origDataService = global.DataService;
    _origSHEETS = global.SHEETS;
    _origLog = global.log_;
    _origCacheGetScriptCache = global.CacheService.getScriptCache;
    _origSSgetActiveSpreadsheet = global.SpreadsheetApp.getActiveSpreadsheet;
    _origMemberCols = global.MEMBER_COLS;

    global.ACCESS_CONTROL = { ENABLED: true, AUTH_CACHE_DURATION: 300 };
    global.SHEETS = { MEMBER_DIR: 'Member Directory' };
    global.log_ = jest.fn();

    cacheGet = jest.fn(() => null); // default: cache miss
    cachePut = jest.fn();
    mockCache = { get: cacheGet, put: cachePut, remove: jest.fn() };
    global.CacheService.getScriptCache = jest.fn(() => mockCache);
  });

  afterEach(() => {
    global.ACCESS_CONTROL = _origAccessControl;
    global.DataService = _origDataService;
    global.SHEETS = _origSHEETS;
    global.log_ = _origLog;
    global.CacheService.getScriptCache = _origCacheGetScriptCache;
    global.SpreadsheetApp.getActiveSpreadsheet = _origSSgetActiveSpreadsheet;
    global.MEMBER_COLS = _origMemberCols;
  });

  function mockOwner(email) {
    var ownerEmail = email;
    global.SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
      getOwner: jest.fn(() => ({ getEmail: jest.fn(() => ownerEmail) })),
      getSheetByName: jest.fn(() => null)
    }));
  }

  test('returns anonymous immediately for empty email without touching cache', () => {
    expect(getUserRole_('')).toBe('anonymous');
    expect(cacheGet).not.toHaveBeenCalled();
  });

  test('returns anonymous for null email', () => {
    expect(getUserRole_(null)).toBe('anonymous');
    expect(cacheGet).not.toHaveBeenCalled();
  });

  test('cache HIT: returns cached role without calling SpreadsheetApp', () => {
    cacheGet.mockReturnValue('steward');
    var ssGetActive = jest.fn();
    global.SpreadsheetApp.getActiveSpreadsheet = ssGetActive;
    expect(getUserRole_('user@example.com')).toBe('steward');
    expect(cacheGet).toHaveBeenCalledWith('user_role_user@example.com');
    expect(ssGetActive).not.toHaveBeenCalled();
    expect(cachePut).not.toHaveBeenCalled();
  });

  test('cache MISS with owner path: populates cache with "admin"', () => {
    cacheGet.mockReturnValue(null);
    mockOwner('owner@example.com');
    expect(getUserRole_('owner@example.com')).toBe('admin');
    expect(cachePut).toHaveBeenCalledWith(
      'user_role_owner@example.com',
      'admin',
      300
    );
  });

  test('cache MISS with DataService: populates cache with delegated role', () => {
    cacheGet.mockReturnValue(null);
    mockOwner('someone-else@example.com');
    global.DataService = {
      getUserRole: jest.fn(() => 'steward')
    };
    expect(getUserRole_('steward@example.com')).toBe('steward');
    expect(global.DataService.getUserRole).toHaveBeenCalledWith('steward@example.com');
    expect(cachePut).toHaveBeenCalledWith(
      'user_role_steward@example.com',
      'steward',
      300
    );
  });

  test('DataService returns null → anonymous, cache NOT populated', () => {
    cacheGet.mockReturnValue(null);
    mockOwner('someone-else@example.com');
    global.DataService = {
      getUserRole: jest.fn(() => null)
    };
    expect(getUserRole_('ghost@example.com')).toBe('anonymous');
    // anonymous is NOT cached — a user who later joins should be detected.
    expect(cachePut).not.toHaveBeenCalled();
  });

  test('cache get failure (throws) → falls through to sheet scan', () => {
    cacheGet.mockImplementation(() => { throw new Error('Cache unavailable'); });
    mockOwner('owner@example.com');
    expect(getUserRole_('owner@example.com')).toBe('admin');
  });

  test('cache put failure is swallowed (role still returned)', () => {
    cacheGet.mockReturnValue(null);
    cachePut.mockImplementation(() => { throw new Error('Quota'); });
    mockOwner('owner@example.com');
    expect(() => getUserRole_('owner@example.com')).not.toThrow();
    expect(getUserRole_('owner@example.com')).toBe('admin');
  });

  test('cache key is email-lowercased (case-insensitive lookup)', () => {
    cacheGet.mockReturnValue('steward');
    getUserRole_('Mixed.Case@Example.COM');
    expect(cacheGet).toHaveBeenCalledWith('user_role_mixed.case@example.com');
  });

  test('no active spreadsheet (web app context) → anonymous', () => {
    cacheGet.mockReturnValue(null);
    global.SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => null);
    expect(getUserRole_('anyone@example.com')).toBe('anonymous');
  });
});

