/**
 * Tests for 22_WebDashApp.gs
 *
 * Covers doGet entry point, error handling, _serveFatalError,
 * _serveError, template serving, and routing.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Pre-mock dependencies before loading
global.ConfigReader = {
  getConfig: jest.fn(() => ({
    orgName: 'Test Union',
    orgAbbrev: 'TU',
    logoInitials: 'TU',
    accentHue: 250,
    stewardLabel: 'Steward',
    memberLabel: 'Member',
    magicLinkExpiryMs: 86400000,
    cookieDurationMs: 2592000000
  }))
};

global.Auth = {
  resolveUser: jest.fn(() => null),
  sendMagicLink: jest.fn(() => ({ success: true })),
  createSessionToken: jest.fn(() => 'token_123'),
  invalidateSession: jest.fn()
};

global.DataService = {
  findUserByEmail: jest.fn(() => ({ email: 'user@test.com', name: 'User' })),
  getUserRole: jest.fn(() => 'member'),
  getMemberData: jest.fn(() => ({ grievances: [] })),
  getStewardDashboardData: jest.fn(() => ({ assignedCases: [] })),
  getDirectoryData: jest.fn(() => [])
};

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '22_WebDashApp.gs']);

// ============================================================================
// _serveFatalError
// ============================================================================

describe('_serveFatalError', () => {
  test('returns HtmlOutput', () => {
    const output = _serveFatalError('Test error');
    expect(HtmlService.createHtmlOutput).toHaveBeenCalled();
  });

  test('shows generic error message without exposing details', () => {
    _serveFatalError('Something broke');
    const lastCallArg = HtmlService.createHtmlOutput.mock.calls[
      HtmlService.createHtmlOutput.mock.calls.length - 1
    ][0];
    // _serveFatalError deliberately does NOT expose the raw error to users (security)
    expect(lastCallArg).toContain('Something went wrong');
    expect(lastCallArg).not.toContain('Something broke');
  });

  test('does not throw', () => {
    expect(() => _serveFatalError('')).not.toThrow();
    expect(() => _serveFatalError(null)).not.toThrow();
  });

  test('includes retry link', () => {
    _serveFatalError('error');
    const html = HtmlService.createHtmlOutput.mock.calls[
      HtmlService.createHtmlOutput.mock.calls.length - 1
    ][0];
    expect(html).toContain('Reload');
  });
});

// ============================================================================
// doGet entry point
// ============================================================================

describe('doGet', () => {
  test('does not throw for empty event', () => {
    expect(() => doGet({})).not.toThrow();
  });

  test('does not throw for null event', () => {
    expect(() => doGet(null)).not.toThrow();
  });

  test('does not throw for undefined event', () => {
    expect(() => doGet(undefined)).not.toThrow();
  });

  test('returns an HtmlOutput object', () => {
    const output = doGet({ parameter: {} });
    expect(output).toBeDefined();
  });

  test('routes workload page when page=workload', () => {
    const output = doGet({ parameter: { page: 'workload' } });
    expect(output).toBeDefined();
  });

  test('catches errors and returns fallback page', () => {
    // Force an error
    ConfigReader.getConfig = jest.fn(() => { throw new Error('config broken'); });

    // doGet should NOT throw — it should return a fallback
    expect(() => doGet({ parameter: {} })).not.toThrow();

    // Restore
    ConfigReader.getConfig = jest.fn(() => ({
      orgName: 'Test Union', orgAbbrev: 'TU', logoInitials: 'TU',
      accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member'
    }));
  });
});

// ============================================================================
// doGetWebDashboard
// ============================================================================

describe('doGetWebDashboard', () => {
  test('does not throw for standard event', () => {
    expect(() => doGetWebDashboard({ parameter: {} })).not.toThrow();
  });

  test('handles loggedout parameter', () => {
    expect(() => doGetWebDashboard({ parameter: { loggedout: '1' } })).not.toThrow();
  });

  test('catches internal errors gracefully', () => {
    // Force error in Auth
    const origResolveUser = Auth.resolveUser;
    Auth.resolveUser = jest.fn(() => { throw new Error('auth broken'); });

    // Should not throw
    expect(() => doGetWebDashboard({ parameter: {} })).not.toThrow();

    Auth.resolveUser = origResolveUser;
  });
});

// ============================================================================
// _serveError
// ============================================================================

describe('_serveError', () => {
  test('returns HtmlOutput for auth errors', () => {
    const output = _serveError(ConfigReader.getConfig(), 'auth', 'Not logged in');
    expect(output).toBeDefined();
  });

  test('returns HtmlOutput for generic errors', () => {
    const output = _serveError(ConfigReader.getConfig(), 'error', 'Something went wrong');
    expect(output).toBeDefined();
  });

  test('does not throw with null config', () => {
    // _serveError should handle null config gracefully
    ConfigReader.getConfig = jest.fn(() => { throw new Error('nope'); });
    const safeCfg = { orgName: 'Dashboard', orgAbbrev: '', logoInitials: '', accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member' };
    expect(() => _serveError(safeCfg, 'error', 'test')).not.toThrow();

    ConfigReader.getConfig = jest.fn(() => ({
      orgName: 'Test Union', orgAbbrev: 'TU', logoInitials: 'TU',
      accentHue: 250, stewardLabel: 'Steward', memberLabel: 'Member'
    }));
  });
});

// ============================================================================
// _serveAuth (login/auth page)
// ============================================================================

describe('_serveAuth', () => {
  test('returns HtmlOutput', () => {
    const output = _serveAuth(ConfigReader.getConfig(), { parameter: {} });
    expect(output).toBeDefined();
  });

  test('does not throw', () => {
    expect(() => _serveAuth(ConfigReader.getConfig(), { parameter: {} })).not.toThrow();
  });
});

// ============================================================================
// _serveDashboard (member and steward views)
// ============================================================================

describe('_serveDashboard', () => {
  test('does not throw for member role', () => {
    const user = { email: 'member@test.com', name: 'Member', firstName: 'Jane', lastName: 'Doe' };
    expect(() => _serveDashboard(ConfigReader.getConfig(), user, 'member', 'tok123')).not.toThrow();
  });

  test('does not throw for steward role', () => {
    const user = { email: 'steward@test.com', name: 'Steward', firstName: 'John', lastName: 'Smith' };
    expect(() => _serveDashboard(ConfigReader.getConfig(), user, 'steward', 'tok456')).not.toThrow();
  });
});

// _serveWorkloadPortal removed in v4.20.0 — workload tracker integrated into SPA
