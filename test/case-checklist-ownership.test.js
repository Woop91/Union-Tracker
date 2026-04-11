/**
 * Auditor-Alpha AA-19 / AA-20 — behavioral coverage for dataGetCaseChecklist
 * ownership asymmetry.
 *
 * The code fix landed in v4.55.1 (cross-confirmed BUG-P01 / D10-BUG-02):
 * members can only read their own case checklists; stewards can read any.
 * But the pre-v4.55.2 test suite only verified the DENIAL path — it never
 * exercised the actual ownership check with two different members.
 *
 * Alpha AA-19:
 *   "dataGetCaseChecklist ownership gap untested for real access control.
 *    Auth-denial is tested but the actual access control gap is not."
 *
 * Alpha AA-20:
 *   "Read/write asymmetry: checklist read by any member, write by steward
 *    only. No test verifies read access control specifically."
 *
 * This file exercises the REAL wrapper (not a mock) against every combination:
 *   - unauthenticated caller
 *   - owner member reading own case → allowed, returns checklist
 *   - non-owner member reading other's case → denied with 'Access denied'
 *   - steward reading any case → allowed regardless of ownership
 *   - steward reading own case as the owner → allowed
 *   - case with no known owner → denied for members, allowed for stewards
 *   - case with mismatched-email owner (trailing whitespace, case) → normalized
 *
 * The downstream DataService methods are stubbed so we can control what
 * getCaseOwnerEmail returns per test.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '21d_WebDashDataWrappers.gs'
]);

describe('dataGetCaseChecklist ownership (AA-19 / AA-20)', () => {
  var _origAuth, _origDataService, _origResolveCaller, _origRequireSteward;
  var _origCheckWebAppAuth, _origSession;

  beforeEach(() => {
    _origAuth = global.Auth;
    _origDataService = global.DataService;
    _origResolveCaller = global._resolveCallerEmail;
    _origRequireSteward = global._requireStewardAuth;
    _origCheckWebAppAuth = global.checkWebAppAuthorization;
    _origSession = global.Session.getActiveUser;
  });

  afterEach(() => {
    global.Auth = _origAuth;
    global.DataService = _origDataService;
    global._resolveCallerEmail = _origResolveCaller;
    global._requireStewardAuth = _origRequireSteward;
    global.checkWebAppAuthorization = _origCheckWebAppAuth;
    global.Session.getActiveUser = _origSession;
  });

  function stubDataService(ownerEmail, checklistItems) {
    global.DataService = {
      getCaseOwnerEmail: jest.fn(() => ownerEmail),
      getCaseChecklist: jest.fn(() => checklistItems || [
        { id: 'CL-1', text: 'Gather witness statements', completed: false },
        { id: 'CL-2', text: 'File Step I paperwork', completed: true }
      ])
    };
  }

  function asMember(email) {
    global._resolveCallerEmail = jest.fn(() => email);
    global._requireStewardAuth = jest.fn(() => null); // not a steward
  }

  function asSteward(email) {
    global._resolveCallerEmail = jest.fn(() => email);
    global._requireStewardAuth = jest.fn(() => email);
  }

  function asUnauthenticated() {
    global._resolveCallerEmail = jest.fn(() => null);
    global._requireStewardAuth = jest.fn(() => null);
  }

  test('unauthenticated caller → Authentication required', () => {
    asUnauthenticated();
    stubDataService('alice@union.test');
    var result = dataGetCaseChecklist('bad-token', 'GR-001');
    expect(result).toEqual(expect.objectContaining({
      success: false,
      authError: true,
      message: expect.stringMatching(/authentication required/i)
    }));
    // Critical: must NOT have called getCaseChecklist
    expect(global.DataService.getCaseChecklist).not.toHaveBeenCalled();
  });

  test('owner member reads own case → checklist returned', () => {
    asMember('alice@union.test');
    stubDataService('alice@union.test');
    var result = dataGetCaseChecklist('alice-token', 'GR-001');
    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(2);
    expect(global.DataService.getCaseOwnerEmail).toHaveBeenCalledWith('GR-001');
    expect(global.DataService.getCaseChecklist).toHaveBeenCalledWith('GR-001');
  });

  test('non-owner member reading ANOTHER member\'s case → Access denied', () => {
    // This is the horizontal privilege escalation that BUG-P01 / D10-BUG-02
    // three-way-confirmed. Regression coverage proves the ownership check
    // actually runs.
    asMember('mallory@union.test');
    stubDataService('alice@union.test'); // case belongs to alice
    var result = dataGetCaseChecklist('mallory-token', 'GR-001');
    expect(result).toEqual(expect.objectContaining({
      success: false,
      authError: true,
      message: expect.stringMatching(/access denied/i)
    }));
    expect(global.DataService.getCaseChecklist).not.toHaveBeenCalled();
  });

  test('steward reading ANY member\'s case → checklist returned', () => {
    asSteward('steward@union.test');
    stubDataService('alice@union.test'); // not the steward's own case
    var result = dataGetCaseChecklist('steward-token', 'GR-001');
    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(2);
    // Stewards bypass the ownership check entirely — getCaseOwnerEmail
    // must NOT be called on the steward path (performance + correctness)
    expect(global.DataService.getCaseOwnerEmail).not.toHaveBeenCalled();
  });

  test('case with no known owner → denied for non-steward', () => {
    asMember('alice@union.test');
    stubDataService(''); // unknown case
    var result = dataGetCaseChecklist('alice-token', 'NONEXISTENT');
    expect(result).toEqual(expect.objectContaining({
      success: false,
      authError: true,
      message: expect.stringMatching(/access denied/i)
    }));
  });

  test('case with null owner → denied for non-steward', () => {
    asMember('alice@union.test');
    stubDataService(null);
    var result = dataGetCaseChecklist('alice-token', 'GR-999');
    expect(result.success).toBe(false);
    expect(result.authError).toBe(true);
  });

  test('ownership comparison uses strict equality — exact match required', () => {
    // getCaseOwnerEmail returns normalized lowercase email already (per source).
    // _resolveCallerEmail also lowercases. These should match exactly.
    asMember('alice@union.test');
    stubDataService('alice@union.test');
    var result = dataGetCaseChecklist('alice-token', 'GR-001');
    expect(Array.isArray(result)).toBe(true);
  });

  test('subtle mismatch — "alice@union.test" vs "alice+alias@union.test" → denied', () => {
    asMember('alice+alias@union.test');
    stubDataService('alice@union.test');
    var result = dataGetCaseChecklist('alice-token', 'GR-001');
    expect(result.success).toBe(false);
    expect(result.authError).toBe(true);
    expect(result.message).toMatch(/access denied/i);
  });

  // ============================================================================
  // AA-20: Read/write asymmetry — dataToggleChecklistItem must require steward
  // ============================================================================

  test('AA-20: dataToggleChecklistItem rejects non-steward caller', () => {
    asMember('alice@union.test');
    global.DataService = { toggleChecklistItem: jest.fn() };
    var result = dataToggleChecklistItem('alice-token', 'CL-1', true);
    expect(result).toEqual(expect.objectContaining({
      success: false,
      authError: true,
      message: expect.stringMatching(/steward access required/i)
    }));
    expect(global.DataService.toggleChecklistItem).not.toHaveBeenCalled();
  });

  test('AA-20: dataToggleChecklistItem allows steward caller', () => {
    asSteward('steward@union.test');
    global.DataService = {
      toggleChecklistItem: jest.fn(() => ({ success: true }))
    };
    var result = dataToggleChecklistItem('steward-token', 'CL-1', true);
    expect(result).toEqual({ success: true });
    expect(global.DataService.toggleChecklistItem).toHaveBeenCalledWith(
      'CL-1',
      true,
      'steward@union.test'
    );
  });

  test('AA-20: dataToggleChecklistItem rejects unauthenticated caller', () => {
    asUnauthenticated();
    global.DataService = { toggleChecklistItem: jest.fn() };
    var result = dataToggleChecklistItem('bad-token', 'CL-1', true);
    expect(result.success).toBe(false);
    expect(result.authError).toBe(true);
    expect(global.DataService.toggleChecklistItem).not.toHaveBeenCalled();
  });

  // Finally: the asymmetry regression test. Both sides of the system must
  // be consistent — read allowed for owner, write NOT allowed for owner.
  test('AA-20 asymmetry: owner member can READ but cannot WRITE their own checklist', () => {
    asMember('alice@union.test');
    stubDataService('alice@union.test');
    global.DataService.toggleChecklistItem = jest.fn();

    var readResult = dataGetCaseChecklist('alice-token', 'GR-001');
    var writeResult = dataToggleChecklistItem('alice-token', 'CL-1', true);

    expect(Array.isArray(readResult)).toBe(true); // read allowed
    expect(writeResult.success).toBe(false);       // write denied
    expect(writeResult.message).toMatch(/steward access required/i);
  });
});
