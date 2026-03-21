/**
 * Tests for 05_Integrations.gs
 *
 * Covers CALENDAR_CONFIG existence, step string comparison logic,
 * and DRIVE_CONFIG structure.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// We need to mock some functions that Integrations.gs calls
global.logAuditEvent = jest.fn();
global.getGrievanceById = jest.fn();
global.updateGrievanceFolderLink = jest.fn();
global.sanitizeFolderName = jest.fn(name => name.replace(/[^a-zA-Z0-9\s\-_.,]/g, ''));
global.BATCH_LIMITS = {
  MAX_EXECUTION_TIME_MS: 300000,
  MAX_API_CALLS_PER_BATCH: 50,
  PAUSE_BETWEEN_BATCHES_MS: 1000
};
global.AUDIT_EVENTS = {
  FOLDER_CREATED: 'FOLDER_CREATED',
  CALENDAR_SYNCED: 'CALENDAR_SYNCED'
};

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '05_Integrations.gs']);

// ============================================================================
// CALENDAR_CONFIG
// ============================================================================

describe('CALENDAR_CONFIG', () => {
  test('is defined (bug fix verification — was missing before)', () => {
    expect(CALENDAR_CONFIG).toBeDefined();
  });

  test('has CALENDAR_NAME', () => {
    expect(CALENDAR_CONFIG.CALENDAR_NAME).toBe('Grievance Deadlines');
  });

  test('has REMINDER_DAYS array', () => {
    expect(Array.isArray(CALENDAR_CONFIG.REMINDER_DAYS)).toBe(true);
    expect(CALENDAR_CONFIG.REMINDER_DAYS).toEqual([7, 3, 1]);
  });

  test('REMINDER_DAYS are in descending order', () => {
    const days = CALENDAR_CONFIG.REMINDER_DAYS;
    for (let i = 1; i < days.length; i++) {
      expect(days[i]).toBeLessThan(days[i - 1]);
    }
  });
});

// ============================================================================
// DRIVE_CONFIG
// ============================================================================

describe('DRIVE_CONFIG', () => {
  test('has ROOT_FOLDER_FALLBACK', () => {
    expect(DRIVE_CONFIG.ROOT_FOLDER_FALLBACK).toBeTruthy();
  });

  test('SUBFOLDER_TEMPLATE contains placeholders', () => {
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{lastName}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{firstName}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE).toContain('{date}');
  });

  test('SUBFOLDER_TEMPLATE_SIMPLE contains fallback placeholders', () => {
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE_SIMPLE).toContain('{grievanceId}');
    expect(DRIVE_CONFIG.SUBFOLDER_TEMPLATE_SIMPLE).toContain('{date}');
  });
});

// ============================================================================
// Calendar step comparison (string-based, not number)
// ============================================================================

describe('Calendar step comparison', () => {
  // The bug was: calendar sync compared step values as numbers (1, 2, 3)
  // when Grievance Log stores them as strings ('Step I', 'Step II', 'Step III')
  test('COMMAND_CONFIG.ESCALATION_STEPS uses string step values', () => {
    const steps = COMMAND_CONFIG.ESCALATION_STEPS;
    expect(steps).toContain('Step II');
    expect(steps).toContain('Step III');
    expect(steps).toContain('Arbitration');
    // Should NOT contain numeric values
    expect(steps).not.toContain(1);
    expect(steps).not.toContain(2);
    expect(steps).not.toContain(3);
  });
});

// ============================================================================
// Email HTML escaping
// ============================================================================

describe('Email HTML escaping', () => {
  test('escapeHtml prevents XSS in email content', () => {
    const memberName = '<script>alert(1)</script>';
    const escaped = escapeHtml(memberName);
    expect(escaped).not.toContain('<script>');
    expect(escaped).toContain('&lt;script&gt;');
  });
});

// ============================================================================
// sanitizeFolderName
// ============================================================================

describe('sanitizeFolderName', () => {
  test('returns Unknown for empty input', () => {
    expect(sanitizeFolderName('')).toBe('Unknown');
    expect(sanitizeFolderName(null)).toBe('Unknown');
    expect(sanitizeFolderName(undefined)).toBe('Unknown');
  });

  test('removes illegal characters', () => {
    expect(sanitizeFolderName('file<>:"/\\|?*name')).toBe('filename');
  });

  test('replaces spaces with underscores', () => {
    expect(sanitizeFolderName('John Doe Case')).toBe('John_Doe_Case');
  });

  test('collapses multiple spaces', () => {
    expect(sanitizeFolderName('John   Doe')).toBe('John_Doe');
  });

  test('truncates to 50 characters', () => {
    const longName = 'A'.repeat(60);
    expect(sanitizeFolderName(longName).length).toBe(50);
  });

  test('handles normal names without changes', () => {
    expect(sanitizeFolderName('GRV-001_Smith_2026')).toBe('GRV-001_Smith_2026');
  });
});

// ============================================================================
// BATCH_LIMITS configuration
// ============================================================================

describe('BATCH_LIMITS', () => {
  test('has reasonable execution time limit', () => {
    expect(BATCH_LIMITS.MAX_EXECUTION_TIME_MS).toBe(300000);
  });

  test('has max API calls per batch', () => {
    expect(BATCH_LIMITS.MAX_API_CALLS_PER_BATCH).toBe(50);
  });
});

// ============================================================================
// CALENDAR_CONFIG structure
// ============================================================================

describe('CALENDAR_CONFIG structure', () => {
  test('has MEETING_CALENDAR_NAME', () => {
    expect(CALENDAR_CONFIG.MEETING_CALENDAR_NAME).toBeTruthy();
  });

  test('has MEETING_DEACTIVATE_HOURS', () => {
    expect(typeof CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS).toBe('number');
    expect(CALENDAR_CONFIG.MEETING_DEACTIVATE_HOURS).toBeGreaterThan(0);
  });
});

// ============================================================================
// CC_CONFIG - Constant Contact Configuration
// ============================================================================

describe('CC_CONFIG', () => {
  test('is defined', () => {
    expect(CC_CONFIG).toBeDefined();
  });

  test('has required API endpoints', () => {
    expect(CC_CONFIG.API_BASE).toBe('https://api.cc.email/v3');
    expect(CC_CONFIG.AUTH_URL).toContain('constantcontact.com');
    expect(CC_CONFIG.TOKEN_URL).toContain('constantcontact.com');
    expect(CC_CONFIG.CONTACTS_ENDPOINT).toBe('/contacts');
  });

  test('ACTIVITY_SUMMARY_ENDPOINT has contact_id placeholder', () => {
    expect(CC_CONFIG.ACTIVITY_SUMMARY_ENDPOINT).toContain('{contact_id}');
  });

  test('has rate limit settings', () => {
    expect(CC_CONFIG.RATE_LIMIT_PER_SECOND).toBe(4);
    expect(CC_CONFIG.RATE_LIMIT_DELAY_MS).toBeGreaterThan(0);
  });

  test('has activity lookback period', () => {
    expect(CC_CONFIG.ACTIVITY_LOOKBACK_DAYS).toBe(365);
  });

  test('has property keys for credential storage', () => {
    expect(CC_CONFIG.PROP_API_KEY).toBeTruthy();
    expect(CC_CONFIG.PROP_API_SECRET).toBeTruthy();
    expect(CC_CONFIG.PROP_ACCESS_TOKEN).toBeTruthy();
    expect(CC_CONFIG.PROP_REFRESH_TOKEN).toBeTruthy();
    expect(CC_CONFIG.PROP_TOKEN_EXPIRY).toBeTruthy();
  });

  test('property keys are unique', () => {
    const keys = [
      CC_CONFIG.PROP_API_KEY,
      CC_CONFIG.PROP_API_SECRET,
      CC_CONFIG.PROP_ACCESS_TOKEN,
      CC_CONFIG.PROP_REFRESH_TOKEN,
      CC_CONFIG.PROP_TOKEN_EXPIRY
    ];
    const unique = new Set(keys);
    expect(unique.size).toBe(keys.length);
  });
});

// ============================================================================
// Constant Contact OAuth Token Management
// ============================================================================

describe('Constant Contact token management', () => {
  let scriptProps;

  beforeEach(() => {
    scriptProps = {};
    PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn(key => scriptProps[key] || null),
      setProperty: jest.fn((key, val) => { scriptProps[key] = val; }),
      deleteProperty: jest.fn(key => { delete scriptProps[key]; })
    });
  });

  test('storeConstantContactTokens_ stores access token', () => {
    storeConstantContactTokens_({
      access_token: 'test-token-123',
      refresh_token: 'refresh-456',
      expires_in: 7200
    });

    expect(scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN]).toBe('test-token-123');
    expect(scriptProps[CC_CONFIG.PROP_REFRESH_TOKEN]).toBe('refresh-456');
    expect(scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY]).toBeTruthy();
  });

  test('storeConstantContactTokens_ calculates expiry with buffer', () => {
    const before = Date.now();
    storeConstantContactTokens_({
      access_token: 'token',
      expires_in: 7200
    });
    const after = Date.now();

    const expiry = new Date(scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY]).getTime();
    // expires_in (7200) - buffer (300) = 6900 seconds
    expect(expiry).toBeGreaterThanOrEqual(before + 6900 * 1000 - 1000);
    expect(expiry).toBeLessThanOrEqual(after + 6900 * 1000 + 1000);
  });

  test('storeConstantContactTokens_ preserves existing refresh token when not provided', () => {
    scriptProps[CC_CONFIG.PROP_REFRESH_TOKEN] = 'existing-refresh';

    storeConstantContactTokens_({
      access_token: 'new-token'
      // no refresh_token in response
    });

    expect(scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN]).toBe('new-token');
    expect(scriptProps[CC_CONFIG.PROP_REFRESH_TOKEN]).toBe('existing-refresh');
  });

  test('getConstantContactToken_ returns token when valid', () => {
    scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN] = 'valid-token';
    scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY] = new Date(Date.now() + 3600000).toISOString();

    const token = getConstantContactToken_();
    expect(token).toBe('valid-token');
  });

  test('getConstantContactToken_ returns null when no token stored', () => {
    const token = getConstantContactToken_();
    expect(token).toBeNull();
  });

  test('getConstantContactToken_ attempts refresh when token expired', () => {
    scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN] = 'expired-token';
    scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY] = new Date(Date.now() - 1000).toISOString();
    scriptProps[CC_CONFIG.PROP_REFRESH_TOKEN] = 'refresh-token';
    scriptProps[CC_CONFIG.PROP_API_KEY] = 'api-key';
    scriptProps[CC_CONFIG.PROP_API_SECRET] = 'api-secret';

    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        access_token: 'refreshed-token',
        refresh_token: 'new-refresh',
        expires_in: 7200
      }))
    });

    const token = getConstantContactToken_();
    expect(token).toBe('refreshed-token');
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      CC_CONFIG.TOKEN_URL,
      expect.objectContaining({ method: 'post' })
    );
  });

  test('getConstantContactToken_ returns null when refresh fails', () => {
    scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN] = 'expired-token';
    scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY] = new Date(Date.now() - 1000).toISOString();
    scriptProps[CC_CONFIG.PROP_REFRESH_TOKEN] = 'refresh-token';
    scriptProps[CC_CONFIG.PROP_API_KEY] = 'api-key';
    scriptProps[CC_CONFIG.PROP_API_SECRET] = 'api-secret';

    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 401),
      getContentText: jest.fn(() => '{"error":"invalid_grant"}')
    });

    const token = getConstantContactToken_();
    expect(token).toBeNull();
  });

  test('getConstantContactToken_ returns null when no refresh token available', () => {
    scriptProps[CC_CONFIG.PROP_ACCESS_TOKEN] = 'expired-token';
    scriptProps[CC_CONFIG.PROP_TOKEN_EXPIRY] = new Date(Date.now() - 1000).toISOString();
    // No refresh token stored

    const token = getConstantContactToken_();
    expect(token).toBeNull();
  });
});

// ============================================================================
// Constant Contact API Helper
// ============================================================================

describe('ccApiGet_', () => {
  let scriptProps;

  beforeEach(() => {
    UrlFetchApp.fetch.mockClear();
    scriptProps = {
      [CC_CONFIG.PROP_ACCESS_TOKEN]: 'valid-token',
      [CC_CONFIG.PROP_TOKEN_EXPIRY]: new Date(Date.now() + 3600000).toISOString()
    };
    PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn(key => scriptProps[key] || null),
      setProperty: jest.fn((key, val) => { scriptProps[key] = val; }),
      deleteProperty: jest.fn(key => { delete scriptProps[key]; })
    });
  });

  test('makes GET request with Bearer token', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => '{"contacts":[]}')
    });

    const result = ccApiGet_('/contacts', { limit: 100 });

    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      expect.stringContaining('https://api.cc.email/v3/contacts?limit=100'),
      expect.objectContaining({
        method: 'get',
        headers: { 'Authorization': 'Bearer valid-token' }
      })
    );
    expect(result).toEqual({ contacts: [] });
  });

  test('returns null when no token available', () => {
    scriptProps = {}; // Clear all tokens

    const result = ccApiGet_('/contacts');
    expect(result).toBeNull();
  });

  test('handles API error responses', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 500),
      getContentText: jest.fn(() => 'Internal Server Error')
    });

    const result = ccApiGet_('/contacts');
    expect(result).toBeNull();
  });

  test('appends query params correctly', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => '{}')
    });

    ccApiGet_('/contacts', { limit: 50, include: 'email_address' });

    const calledUrl = UrlFetchApp.fetch.mock.calls[0][0];
    expect(calledUrl).toContain('limit=50');
    expect(calledUrl).toContain('include=email_address');
  });

  test('skips null/undefined params', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => '{}')
    });

    ccApiGet_('/contacts', { limit: 50, cursor: null, offset: undefined });

    const calledUrl = UrlFetchApp.fetch.mock.calls[0][0];
    expect(calledUrl).toContain('limit=50');
    expect(calledUrl).not.toContain('cursor');
    expect(calledUrl).not.toContain('offset');
  });
});

// ============================================================================
// Constant Contact Engagement Data Parsing
// ============================================================================

describe('fetchCCContactEngagement_', () => {
  let scriptProps;

  beforeEach(() => {
    UrlFetchApp.fetch.mockClear();
    scriptProps = {
      [CC_CONFIG.PROP_ACCESS_TOKEN]: 'valid-token',
      [CC_CONFIG.PROP_TOKEN_EXPIRY]: new Date(Date.now() + 3600000).toISOString()
    };
    PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn(key => scriptProps[key] || null),
      setProperty: jest.fn((key, val) => { scriptProps[key] = val; }),
      deleteProperty: jest.fn(key => { delete scriptProps[key]; })
    });
  });

  test('calculates open rate from campaign activities', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        campaign_activities: [
          { em_sends: 3, em_opens: 2, em_opens_date: '2026-01-15T10:00:00Z' },
          { em_sends: 2, em_opens: 1, em_clicks_date: '2026-02-01T14:00:00Z' }
        ]
      }))
    });

    const result = fetchCCContactEngagement_('contact-123');

    expect(result).not.toBeNull();
    // 3 opens / 5 sends = 60%
    expect(result.openRate).toBe(60);
  });

  test('returns 0 open rate when no sends', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        campaign_activities: [
          { em_sends: 0, em_opens: 0 }
        ]
      }))
    });

    const result = fetchCCContactEngagement_('contact-123');
    expect(result.openRate).toBe(0);
  });

  test('tracks most recent activity date', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        campaign_activities: [
          { em_sends: 1, em_opens: 1, em_sends_date: '2026-01-01T00:00:00Z', em_opens_date: '2026-01-02T00:00:00Z' },
          { em_sends: 1, em_opens: 1, em_sends_date: '2026-02-10T00:00:00Z', em_clicks_date: '2026-02-11T00:00:00Z' }
        ]
      }))
    });

    const result = fetchCCContactEngagement_('contact-123');
    expect(result.lastActivityDate).toEqual(new Date('2026-02-11T00:00:00Z'));
  });

  test('returns null when API returns no data', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => '{}')
    });

    const result = fetchCCContactEngagement_('contact-123');
    expect(result).toBeNull();
  });

  test('returns null when API call fails', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 500),
      getContentText: jest.fn(() => 'error')
    });

    const result = fetchCCContactEngagement_('contact-123');
    expect(result).toBeNull();
  });

  test('passes correct date range to activity summary endpoint', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({ campaign_activities: [] }))
    });

    fetchCCContactEngagement_('contact-456');

    const calledUrl = UrlFetchApp.fetch.mock.calls[0][0];
    expect(calledUrl).toContain('contact-456');
    expect(calledUrl).toContain('start=');
    expect(calledUrl).toContain('end=');
  });
});

// ============================================================================
// Constant Contact Contact Fetching
// ============================================================================

describe('fetchCCContacts_', () => {
  let scriptProps;

  beforeEach(() => {
    UrlFetchApp.fetch.mockClear();
    scriptProps = {
      [CC_CONFIG.PROP_ACCESS_TOKEN]: 'valid-token',
      [CC_CONFIG.PROP_TOKEN_EXPIRY]: new Date(Date.now() + 3600000).toISOString()
    };
    PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn(key => scriptProps[key] || null),
      setProperty: jest.fn((key, val) => { scriptProps[key] = val; }),
      deleteProperty: jest.fn(key => { delete scriptProps[key]; })
    });
  });

  test('builds email-to-contactId map', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        contacts: [
          { contact_id: 'id-1', email_address: { address: 'alice@union.org' } },
          { contact_id: 'id-2', email_address: { address: 'Bob@Union.org' } }
        ]
      }))
    });

    const result = fetchCCContacts_();

    expect(result['alice@union.org']).toBe('id-1');
    // Should be lowercased
    expect(result['bob@union.org']).toBe('id-2');
  });

  test('returns empty map when API fails', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 401),
      getContentText: jest.fn(() => '{"error":"unauthorized"}')
    });

    const result = fetchCCContacts_();
    expect(Object.keys(result).length).toBe(0);
  });

  test('skips contacts without email addresses', () => {
    UrlFetchApp.fetch.mockReturnValueOnce({
      getResponseCode: jest.fn(() => 200),
      getContentText: jest.fn(() => JSON.stringify({
        contacts: [
          { contact_id: 'id-1', email_address: { address: 'valid@test.com' } },
          { contact_id: 'id-2', email_address: null },
          { contact_id: 'id-3' }
        ]
      }))
    });

    const result = fetchCCContacts_();
    expect(Object.keys(result).length).toBe(1);
    expect(result['valid@test.com']).toBe('id-1');
  });
});

// ============================================================================
// Constant Contact Disconnect
// ============================================================================

// ============================================================================
// sendEmailToMember
// Regression tests for 3f88994: safeSubject/safeBody were undefined (ReferenceError
// on every call). Also covers HTML stripping, auth check, and member-not-found.
// ============================================================================

describe('sendEmailToMember', () => {
  var _origGetUserRole;
  beforeEach(() => {
    global.getMemberById = jest.fn(() => ({ Email: 'member@test.com', 'Member ID': 'MEM-001' }));
    // Mock getUserRole_ directly so auth passes without spreadsheet reads.
    _origGetUserRole = global.getUserRole_;
    global.getUserRole_ = jest.fn(() => 'admin');
    Session.getActiveUser = jest.fn(() => ({ getEmail: jest.fn(() => 'admin@test.com') }));
    MailApp.sendEmail = jest.fn();
    MailApp.getRemainingDailyQuota = jest.fn(() => 100);
  });
  afterEach(() => {
    global.getUserRole_ = _origGetUserRole;
  });

  test('returns success and calls safeSendEmail_ (safeSubject/safeBody no longer undefined)', () => {
    // Before 3f88994, safeSubject and safeBody were undefined → ReferenceError.
    // This test fails if that regression is reintroduced.
    const result = sendEmailToMember('MEM-001', 'Hello', '<p>Body text</p>');
    expect(result.success).toBe(true);
    expect(MailApp.sendEmail).toHaveBeenCalledTimes(1);
  });

  test('strips HTML tags from subject before sending', () => {
    sendEmailToMember('MEM-001', '<b>Important</b> Notice', 'Body');
    const callArgs = MailApp.sendEmail.mock.calls[0][0];
    expect(callArgs.subject).toBe('Important Notice');
    expect(callArgs.subject).not.toContain('<b>');
  });

  test('preserves subject text when no HTML tags present', () => {
    sendEmailToMember('MEM-001', 'Plain subject', 'Body');
    const callArgs = MailApp.sendEmail.mock.calls[0][0];
    expect(callArgs.subject).toBe('Plain subject');
  });

  test('returns error when member is not found', () => {
    global.getMemberById = jest.fn(() => null);
    const result = sendEmailToMember('NONEXISTENT', 'Subject', 'Body');
    expect(result.success).toBe(false);
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });

  test('returns error when member email is invalid', () => {
    global.getMemberById = jest.fn(() => ({ Email: 'not-a-valid-email', 'Member ID': 'MEM-X' }));
    const result = sendEmailToMember('MEM-X', 'Subject', 'Body');
    expect(result.success).toBe(false);
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });

  test('returns error when caller role is not admin or steward', () => {
    global.getUserRole_ = jest.fn(() => 'member');
    const result = sendEmailToMember('MEM-001', 'Subject', 'Body');
    expect(result.success).toBe(false);
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });
});

describe('disconnectConstantContact', () => {
  test('removes all CC properties when user confirms', () => {
    const deletedKeys = [];
    const mockProps = {
      getProperty: jest.fn(() => 'some-value'),
      setProperty: jest.fn(),
      deleteProperty: jest.fn(key => { deletedKeys.push(key); })
    };
    PropertiesService.getScriptProperties.mockReturnValue(mockProps);

    const mockUi = {
      alert: jest.fn(() => 'YES'),
      Button: { YES: 'YES', NO: 'NO' },
      ButtonSet: { YES_NO: 'YES_NO', OK: 'OK' }
    };
    SpreadsheetApp.getUi.mockReturnValue(mockUi);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({ toast: jest.fn() });

    disconnectConstantContact();

    expect(deletedKeys).toContain(CC_CONFIG.PROP_API_KEY);
    expect(deletedKeys).toContain(CC_CONFIG.PROP_API_SECRET);
    expect(deletedKeys).toContain(CC_CONFIG.PROP_ACCESS_TOKEN);
    expect(deletedKeys).toContain(CC_CONFIG.PROP_REFRESH_TOKEN);
    expect(deletedKeys).toContain(CC_CONFIG.PROP_TOKEN_EXPIRY);
  });
});

// ============================================================================
// LockService wrapping for WebApp resource functions
// ============================================================================

describe('LockService: addWebAppResource', () => {
  beforeEach(() => jest.clearAllMocks());

  test('acquires script lock during resource creation', () => {
    const mockLock = {
      tryLock: jest.fn(() => true),
      releaseLock: jest.fn()
    };
    LockService.getScriptLock.mockReturnValue(mockLock);

    // Mock auth to pass
    global.checkWebAppAuthorization = jest.fn(() => ({ isAuthorized: true, email: 'test@example.com' }));

    const resourceData = [['Resource ID', 'Title'], ['RES-001', 'Test']];
    const sheet = createMockSheet(SHEETS.RESOURCES, resourceData);
    sheet.appendRow = jest.fn();
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    addWebAppResource('token', { title: 'New Resource' });
    expect(mockLock.tryLock).toHaveBeenCalled();
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });
});

describe('LockService: deleteWebAppResource', () => {
  beforeEach(() => jest.clearAllMocks());

  test('acquires script lock during resource deletion', () => {
    const mockLock = {
      tryLock: jest.fn(() => true),
      releaseLock: jest.fn()
    };
    LockService.getScriptLock.mockReturnValue(mockLock);

    global.checkWebAppAuthorization = jest.fn(() => ({ isAuthorized: true, email: 'test@example.com' }));

    const resourceData = [['Resource ID', 'Visible'], ['RES-001', 'Yes']];
    const sheet = createMockSheet(SHEETS.RESOURCES, resourceData);
    const mockSS = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);

    deleteWebAppResource('token', 'RES-001');
    expect(mockLock.tryLock).toHaveBeenCalled();
    expect(mockLock.releaseLock).toHaveBeenCalled();
  });
});
