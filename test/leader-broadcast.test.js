/**
 * Tests for sendUnitBroadcast in 21_WebDashDataService.gs
 *
 * Covers: unit broadcast sending, rate limiting (3 per hour per leader),
 * sender exclusion from recipients, empty unit handling, safeSendEmail_ usage,
 * and formula escaping of subject/message.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '06_Maintenance.gs', '21_WebDashDataService.gs']);

// ============================================================================
// Helpers
// ============================================================================

/** Member Directory headers matching the HEADERS aliases in DataService */
var MEMBER_HEADERS = [
  'Email', 'Name', 'First Name', 'Last Name', 'Role', 'Unit', 'Phone',
  'Join Date', 'Dues Status', 'Member ID', 'Work Location', 'Office Days',
  'Assigned Steward', 'Is Steward'
];

function makeMemberRow(email, name, firstName, lastName, unit) {
  return [email, name, firstName, lastName, 'Member', unit, '555-0000', new Date('2023-01-01'), 'Active', 'MEM-001', 'HQ', '', '', false];
}

function setupBroadcastSheets(memberRows) {
  var memberData = [MEMBER_HEADERS].concat(memberRows || []);
  var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, memberData);
  var ss = createMockSpreadsheet([memberSheet]);
  SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
  return ss;
}

beforeEach(() => {
  // Reset DataService caches
  if (typeof DataService !== 'undefined' && DataService._resetSSCache) {
    DataService._resetSSCache();
  }
  if (typeof DataService !== 'undefined' && DataService._invalidateSheetCache) {
    DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
  }

  // Mock safeSendEmail_ as a global — the broadcast function checks typeof
  global.safeSendEmail_ = jest.fn(function(opts) {
    return { success: true };
  });

  // Mock secureLog so the broadcast function can call it
  global.secureLog = jest.fn();

  // Re-mock auth helpers that loadSources overwrites
  global._resolveCallerEmail = jest.fn(() => 'leader@test.com');
  global._requireStewardAuth = jest.fn(() => 'leader@test.com');
});

// ============================================================================
// sendUnitBroadcast — Basic sending
// ============================================================================

describe('DataService.sendUnitBroadcast', () => {
  test('sends to unit members and returns success with sentCount', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A'),
      makeMemberRow('member2@test.com', 'Member Two', 'Member', 'Two', 'Unit A')
    ]);

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Hello team!', 'Weekly Update'
    );

    expect(result.success).toBe(true);
    expect(result.sentCount).toBe(2);
    expect(result.failedCount).toBe(0);
  });

  test('uses safeSendEmail_ when available', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Test message', 'Subject'
    );

    expect(safeSendEmail_).toHaveBeenCalledWith(expect.objectContaining({
      to: 'member1@test.com',
      replyTo: 'leader@test.com'
    }));
  });

  test('returns failure message on send error', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);
    global.safeSendEmail_ = jest.fn(function() {
      return { success: false, error: 'Quota exceeded' };
    });

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Test', 'Subject'
    );

    expect(result.success).toBe(true); // overall call succeeds
    expect(result.sentCount).toBe(0);
    expect(result.failedCount).toBe(1);
  });
});

// ============================================================================
// sendUnitBroadcast — Sender excluded
// ============================================================================

describe('DataService.sendUnitBroadcast — sender exclusion', () => {
  test('excludes sender from recipients', () => {
    setupBroadcastSheets([
      makeMemberRow('leader@test.com', 'Leader Name', 'Leader', 'Name', 'Unit A'),
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Hello!', 'Update'
    );

    expect(result.sentCount).toBe(1);
    // Verify the sender was NOT included in send calls
    var sentTos = safeSendEmail_.mock.calls.map(function(c) { return c[0].to; });
    expect(sentTos).not.toContain('leader@test.com');
    expect(sentTos).toContain('member1@test.com');
  });

  test('sender exclusion is case-insensitive', () => {
    setupBroadcastSheets([
      makeMemberRow('leader@test.com', 'Leader Name', 'Leader', 'Name', 'Unit A'),
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    var result = DataService.sendUnitBroadcast(
      'LEADER@TEST.COM', 'Leader Name', 'Unit A', 'Hello!', 'Update'
    );

    expect(result.sentCount).toBe(1);
  });
});

// ============================================================================
// sendUnitBroadcast — Empty unit
// ============================================================================

describe('DataService.sendUnitBroadcast — empty unit', () => {
  test('returns error when no members in unit', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit B')
    ]);

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Hello!', 'Update'
    );

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/no members/i);
  });

  test('returns error when all unit members are the sender', () => {
    setupBroadcastSheets([
      makeMemberRow('leader@test.com', 'Leader Name', 'Leader', 'Name', 'Unit A')
    ]);

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Hello!', 'Update'
    );

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/no members/i);
  });
});

// ============================================================================
// sendUnitBroadcast — Rate limiting
// ============================================================================

describe('DataService.sendUnitBroadcast — rate limiting', () => {
  test('allows first 3 broadcasts', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    for (var i = 0; i < 3; i++) {
      // Reset caches between calls since DataService caches internally
      if (DataService._invalidateSheetCache) DataService._invalidateSheetCache(SHEETS.MEMBER_DIR);
      if (DataService._resetSSCache) DataService._resetSSCache();
      setupBroadcastSheets([
        makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
      ]);

      var result = DataService.sendUnitBroadcast(
        'leader@test.com', 'Leader Name', 'Unit A', 'Message ' + (i + 1), 'Subject'
      );
      expect(result.success).toBe(true);
    }
  });

  test('blocks 4th broadcast within the hour', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    // Seed the cache with a count of 3 already
    var cache = CacheService.getScriptCache();
    cache.put('BROADCAST_RATE_leader@test.com', '3', 3600);

    var result = DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Too many!', 'Subject'
    );

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/rate limit/i);
    expect(result.sentCount).toBe(0);
  });

  test('rate limit is per-leader (different leaders have separate limits)', () => {
    // Seed leader1 at limit
    var cache = CacheService.getScriptCache();
    cache.put('BROADCAST_RATE_leader1@test.com', '3', 3600);

    // leader2 should still be allowed
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    var result = DataService.sendUnitBroadcast(
      'leader2@test.com', 'Leader Two', 'Unit A', 'Hello!', 'Subject'
    );

    expect(result.success).toBe(true);
  });

  test('rate limit key is lowercase-trimmed', () => {
    var cache = CacheService.getScriptCache();
    cache.put('BROADCAST_RATE_leader@test.com', '3', 3600);

    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    // Calling with mixed case should still hit the rate limit
    var result = DataService.sendUnitBroadcast(
      '  LEADER@TEST.COM  ', 'Leader Name', 'Unit A', 'Test', 'Subject'
    );

    expect(result.success).toBe(false);
    expect(result.message).toMatch(/rate limit/i);
  });
});

// ============================================================================
// sendUnitBroadcast — Subject fallback and escaping
// ============================================================================

describe('DataService.sendUnitBroadcast — subject and message handling', () => {
  test('uses default subject when none provided', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Hello team!', null
    );

    expect(safeSendEmail_).toHaveBeenCalled();
    var subject = safeSendEmail_.mock.calls[0][0].subject;
    // Default subject includes sender name and unit
    expect(subject).toContain('Leader Name');
    expect(subject).toContain('Unit A');
  });

  test('uses provided subject when given', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);

    DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', 'Body text', 'Custom Subject'
    );

    var subject = safeSendEmail_.mock.calls[0][0].subject;
    expect(subject).toContain('Custom Subject');
  });

  test('applies escapeForFormula to subject and message', () => {
    setupBroadcastSheets([
      makeMemberRow('member1@test.com', 'Member One', 'Member', 'One', 'Unit A')
    ]);
    var spy = jest.spyOn(global, 'escapeForFormula');

    DataService.sendUnitBroadcast(
      'leader@test.com', 'Leader Name', 'Unit A', '=EVIL() body', '=EVIL() subject'
    );

    expect(spy).toHaveBeenCalled();
    spy.mockRestore();
  });
});
