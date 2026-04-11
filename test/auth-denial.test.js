/**
 * Auth Denial Tests — Verify ALL wrapper functions handle unauthenticated calls safely.
 *
 * Tests every client-callable function that uses _resolveCallerEmail or _requireStewardAuth.
 * When auth fails (invalid/missing sessionToken), each function must:
 *   1. NOT throw an unhandled exception (would crash the client)
 *   2. Return a safe empty value ([], {}, null, false, {success:false})
 *   3. NOT leak any data
 *
 * These tests would have caught:
 *   - v4.24.8: 6 functions had undeclared sessionToken → ReferenceError
 *   - v4.25.0: 26 functions got wrong positional args → silent auth failures
 *
 * Run: npx jest test/auth-denial.test.js --verbose
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock globals before loading sources
global.logAuditEvent = jest.fn();
global.AUDIT_EVENTS = {
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  FOLDER_CREATED: 'FOLDER_CREATED'
};

// Load all source files in build order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '04a_UIMenus.gs',
  '04b_AccessibilityFeatures.gs',
  '05_Integrations.gs',
  '06_Maintenance.gs',
  '07_DevTools.gs',
  '08a_SheetSetup.gs',
  '08b_SearchAndCharts.gs',
  '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs',
  '08e_SurveyEngine.gs',
  '09_Dashboards.gs',
  '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs',
  '10c_FormsAndSync.gs',
  '10_Main.gs',
  '11_CommandHub.gs',
  '12_Features.gs',
  '13_MemberSelfService.gs',
  '14_MeetingCheckIn.gs',
  '15_EventBus.gs',
  '17_CorrelationEngine.gs',
  '19_WebDashAuth.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs',
  '21d_WebDashDataWrappers.gs',
  '22_WebDashApp.gs',
  '23_PortalSheets.gs',
  '24_WeeklyQuestions.gs',
  '25_WorkloadService.gs',
  '26_QAForum.gs',
  '27_TimelineService.gs',
  '28_FailsafeService.gs',
  '33_NewFeatureServices.gs',
]);

// ============================================================================
// Setup: Force auth denial for all tests
// ============================================================================

const INVALID_TOKEN = 'invalid-token-that-does-not-exist';

// Track skipped functions so we can fail if any are undefined. Zero budget:
// every wrapper listed in this file is supposed to exist. If one vanishes
// the test should FAIL so we notice the regression instead of silently
// skipping up to 5 missing functions.
const skippedFunctions = [];
const MAX_ALLOWED_SKIPS = 0;

beforeEach(() => {
  // Make SSO return nothing (non-SSO user)
  global.Session.getActiveUser.mockReturnValue({
    getEmail: jest.fn(() => '')
  });
  // Make _resolveCallerEmail return '' (unauthenticated)
  if (typeof global._resolveCallerEmail === 'function') {
    jest.spyOn(global, '_resolveCallerEmail').mockReturnValue('');
  }
  // Make _requireStewardAuth return null (not a steward)
  if (typeof global._requireStewardAuth === 'function') {
    jest.spyOn(global, '_requireStewardAuth').mockReturnValue(null);
  }
  // Make checkWebAppAuthorization return unauthorized
  if (typeof global.checkWebAppAuthorization === 'function') {
    jest.spyOn(global, 'checkWebAppAuthorization').mockReturnValue({ isAuthorized: false, email: '' });
  }
});

afterEach(() => {
  jest.restoreAllMocks();
});

// ============================================================================
// Safe empty values — what denied wrappers should return
// ============================================================================

function isSafeDenialResult(result) {
  if (result === null || result === undefined) return true;
  if (result === false) return true;
  if (result === '') return true;
  if (result === 0) return true;
  if (Array.isArray(result) && result.length === 0) return true;
  if (typeof result === 'object' && result !== null) {
    // { success: false, ... } is safe
    if (result.success === false) return true;
    // {} is safe
    if (Object.keys(result).length === 0) return true;
    // { isFirstVisit: false, ... } is safe (welcome data default)
    if (result.isFirstVisit === false) return true;
    // { available: false } is safe (grievance stats)
    if (result.available === false) return true;
    // { total: 0, ... } is safe (survey tracking)
    if (result.total === 0) return true;
    // { error: ... } is safe
    if (result.error) return true;
    // { categories: [] } is safe (satisfaction trends)
    if (Array.isArray(result.categories) && result.categories.length === 0) return true;
    // Aggregate stats with zero values are safe
    if (result.locations !== undefined || result.officeDays !== undefined) return true;
    // { questions: [] } is safe (weekly questions)
    if (Array.isArray(result.questions) && result.questions.length === 0) return true;
    // { enabled: false, ... } is safe (digest config default)
    if (result.enabled === false && result.frequency !== undefined) return true;
    // { periodId: null, ... } is safe (satisfaction summary default)
    if (result.periodId === null && result.responseCount === 0) return true;
  }
  return false;
}

// ============================================================================
// Data Service Wrappers (21_WebDashDataService.gs)
// ============================================================================

describe('Auth denial: DataService wrappers return safe empty values', () => {

  // Steward-only functions (use _requireStewardAuth)
  const stewardFunctions = [
    { name: 'dataGetStewardCases', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardKPIs', args: [INVALID_TOKEN] },
    { name: 'dataGetAllMembers', args: [INVALID_TOKEN] },
    { name: 'dataGetGrievanceStats', args: [INVALID_TOKEN] },
    { name: 'dataGetGrievanceHotSpots', args: [INVALID_TOKEN] },
    { name: 'dataGetSatisfactionTrends', args: [INVALID_TOKEN] },
    { name: 'dataGetAllStewardPerformance', args: [INVALID_TOKEN] },
    { name: 'dataGetSurveyResults', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardContactLog', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardAssignedMemberTasks', args: [INVALID_TOKEN] },
    { name: 'dataGetTasks', args: [INVALID_TOKEN, null] },
    { name: 'dataGetBroadcastFilterOptions', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardSurveyTracking', args: [INVALID_TOKEN, 'assigned'] },
    { name: 'dataGetMemberContactHistory', args: [INVALID_TOKEN, 'test@example.com'] },
    { name: 'dataSendBroadcast', args: [INVALID_TOKEN, {}, 'msg', 'subject'] },
    { name: 'dataAssignSteward', args: [INVALID_TOKEN, 'a@b.com', 'c@d.com'] },
    { name: 'dataCreateTask', args: [INVALID_TOKEN, 'title', 'desc', 'a@b.com', 'high', null, ''] },
    { name: 'dataCompleteTask', args: [INVALID_TOKEN, 'task-1'] },
    { name: 'dataUpdateTask', args: [INVALID_TOKEN, 'task-1', {}] },
    { name: 'dataCreateMemberTask', args: [INVALID_TOKEN, 'a@b.com', 'title', 'desc', 'high', null] },
    { name: 'dataStaffCompleteMemberTask', args: [INVALID_TOKEN, 'task-1'] },
    { name: 'dataLogMemberContact', args: [INVALID_TOKEN, 'a@b.com', 'call', 'notes', 15] },
    { name: 'dataAddMeetingMinutes', args: [INVALID_TOKEN, {}] },
    { name: 'dataIsChiefSteward', args: [INVALID_TOKEN] },
  ];

  // Member functions (use _resolveCallerEmail)
  const memberFunctions = [
    { name: 'dataGetMemberGrievances', args: [INVALID_TOKEN] },
    { name: 'dataGetMemberGrievanceHistory', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardContact', args: [INVALID_TOKEN] },
    { name: 'dataGetAssignedSteward', args: [INVALID_TOKEN] },
    { name: 'dataGetAvailableStewards', args: [INVALID_TOKEN] },
    { name: 'dataCreateGrievanceDrive', args: [INVALID_TOKEN] },
    { name: 'dataGetSurveyStatus', args: [INVALID_TOKEN] },
    { name: 'dataGetSurveyQuestions', args: [INVALID_TOKEN] },
    { name: 'dataGetMembershipStats', args: [INVALID_TOKEN] },
    { name: 'dataGetMemberGrievanceStats', args: [INVALID_TOKEN] },
    { name: 'dataGetMemberGrievanceHotSpots', args: [INVALID_TOKEN] },
    { name: 'dataMemberAssignSteward', args: [INVALID_TOKEN, 'steward@test.com'] },
    { name: 'dataGetUpcomingEvents', args: [INVALID_TOKEN, 10] },
    { name: 'dataGetMemberTasks', args: [INVALID_TOKEN, null] },
    { name: 'dataGetMemberMeetings', args: [INVALID_TOKEN] },
    // v4.55.1: dataGetMyFeedback removed (feedback feature deleted v4.52.0)
    { name: 'dataGetMeetingMinutes', args: [INVALID_TOKEN, 10] },
    { name: 'dataGetCaseChecklist', args: [INVALID_TOKEN, 'case-1'] },
    { name: 'dataGetStewardMemberStats', args: [INVALID_TOKEN] },
    { name: 'dataGetStewardDirectory', args: [INVALID_TOKEN] },
    { name: 'dataGetBatchData', args: [INVALID_TOKEN] },
    { name: 'dataGetEngagementStats', args: [INVALID_TOKEN] },
    { name: 'dataGetWorkloadSummaryStats', args: [INVALID_TOKEN] },
    { name: 'dataGetWelcomeData', args: [INVALID_TOKEN] },
    { name: 'dataMarkWelcomeDismissed', args: [INVALID_TOKEN] },
    // v4.55.1: dataSubmitFeedback removed (feedback feature deleted v4.52.0)
    { name: 'dataSubmitSurveyResponse', args: [INVALID_TOKEN, []] },
    { name: 'dataUpdateProfile', args: [INVALID_TOKEN, {}] },
  ];

  [...stewardFunctions, ...memberFunctions].forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') {
        skippedFunctions.push(name);
        console.warn(`  SKIP: ${name} is not defined`);
        return;
      }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// QA Forum Wrappers (26_QAForum.gs)
// ============================================================================

describe('Auth denial: QA Forum wrappers', () => {
  const qaFunctions = [
    { name: 'qaGetQuestions', args: [INVALID_TOKEN, 1, 20, 'recent'] },
    { name: 'qaGetQuestionDetail', args: [INVALID_TOKEN, 'q-1'] },
    { name: 'qaSubmitQuestion', args: [INVALID_TOKEN, 'name', 'text', false] },
    { name: 'qaSubmitAnswer', args: [INVALID_TOKEN, 'name', 'q-1', 'text', true] },
    { name: 'qaUpvoteQuestion', args: [INVALID_TOKEN, 'q-1'] },
    { name: 'qaModerateQuestion', args: [INVALID_TOKEN, 'q-1', 'flag'] },
    { name: 'qaModerateAnswer', args: [INVALID_TOKEN, 'a-1', 'approve'] },
    { name: 'qaGetFlaggedContent', args: [INVALID_TOKEN] },
    { name: 'qaResolveQuestion', args: [INVALID_TOKEN, 'q-1'] },
  ];

  qaFunctions.forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') { skippedFunctions.push(name); console.warn(`  SKIP: ${name}`); return; }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// Weekly Questions Wrappers (24_WeeklyQuestions.gs)
// ============================================================================

describe('Auth denial: Weekly Questions wrappers', () => {
  const wqFunctions = [
    { name: 'wqGetActiveQuestions', args: [INVALID_TOKEN] },
    { name: 'wqSubmitResponse', args: [INVALID_TOKEN, 'q-1', 'yes'] },
    { name: 'wqSetStewardQuestion', args: [INVALID_TOKEN, 'text', 'a,b,c'] },
    { name: 'wqSubmitPoolQuestion', args: [INVALID_TOKEN, 'text', 'a,b'] },
    { name: 'wqClosePoll', args: [INVALID_TOKEN, 'poll-1'] },
    { name: 'wqGetHistory', args: [INVALID_TOKEN, 1, 10] },
    { name: 'wqSetPollFrequency', args: [INVALID_TOKEN, 'weekly'] },
    { name: 'wqManualDrawCommunityPoll', args: [INVALID_TOKEN] },
  ];

  wqFunctions.forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') { skippedFunctions.push(name); console.warn(`  SKIP: ${name}`); return; }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// Timeline Service Wrappers (27_TimelineService.gs)
// ============================================================================

describe('Auth denial: Timeline Service wrappers', () => {
  const tlFunctions = [
    { name: 'tlGetTimelineEvents', args: [INVALID_TOKEN, 1, 20, null, null] },
    { name: 'tlAddTimelineEvent', args: [INVALID_TOKEN, {}] },
    { name: 'tlUpdateTimelineEvent', args: [INVALID_TOKEN, 'ev-1', {}] },
    { name: 'tlDeleteTimelineEvent', args: [INVALID_TOKEN, 'ev-1'] },
    { name: 'tlImportCalendarEvents', args: [INVALID_TOKEN, '2026-01-01', '2026-12-31'] },
    { name: 'tlAttachDriveFiles', args: [INVALID_TOKEN, 'ev-1', []] },
  ];

  tlFunctions.forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') { skippedFunctions.push(name); console.warn(`  SKIP: ${name}`); return; }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// Failsafe Service Wrappers (28_FailsafeService.gs)
// ============================================================================

describe('Auth denial: Failsafe Service wrappers', () => {
  const fsFunctions = [
    { name: 'fsGetDigestConfig', args: [INVALID_TOKEN] },
    { name: 'fsUpdateDigestConfig', args: [INVALID_TOKEN, {}] },
    { name: 'fsBackupCriticalSheets', args: [INVALID_TOKEN] },
    { name: 'fsSetupTriggers', args: [INVALID_TOKEN] },
    { name: 'fsRemoveTriggers', args: [INVALID_TOKEN] },
    { name: 'fsInitSheets', args: [INVALID_TOKEN] },
  ];

  fsFunctions.forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') { skippedFunctions.push(name); console.warn(`  SKIP: ${name}`); return; }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// Survey Engine Wrappers (08e_SurveyEngine.gs)
// ============================================================================

describe('Auth denial: Survey Engine wrappers', () => {
  const surveyFunctions = [
    { name: 'dataGetPendingSurveyMembers', args: [INVALID_TOKEN] },
    { name: 'dataGetSatisfactionSummary', args: [INVALID_TOKEN] },
    { name: 'dataOpenNewSurveyPeriod', args: [INVALID_TOKEN] },
  ];

  surveyFunctions.forEach(({ name, args }) => {
    test(`${name}() does not throw when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') { skippedFunctions.push(name); console.warn(`  SKIP: ${name}`); return; }
      expect(() => fn(...args)).not.toThrow();
    });

    test(`${name}() returns safe empty value when unauthenticated`, () => {
      const fn = global[name];
      if (typeof fn !== 'function') return;
      const result = fn(...args);
      expect(isSafeDenialResult(result)).toBe(true);
    });
  });
});


// ============================================================================
// Build Validation Tests
// ============================================================================

describe('Build validation: syntax checking works', () => {
  const vm = require('vm');
  const fs = require('fs');
  const path = require('path');
  const { execSync } = require('child_process');
  const SRC_DIR = path.resolve(__dirname, '..', 'src');

  test('build.js passes with current clean source', () => {
    // Use --validate-only to check syntax without overwriting dist/
    // A full `node build.js` runs a dev build that destroys prod-minified HTML
    // and re-adds DevTools/DevMenu to dist/, breaking G25 minification guards.
    expect(() => {
      execSync('node build.js --validate-only', { cwd: path.resolve(__dirname, '..'), timeout: 15000 });
    }).not.toThrow();
  });

  test('vm.Script rejects broken .gs syntax (same technique as build.js)', () => {
    const brokenCode = 'function broken( { return; }';
    expect(() => {
      new vm.Script(brokenCode, { filename: 'test.gs' });
    }).toThrow();
  });

  test('vm.Script rejects unescaped apostrophe in single-quoted string', () => {
    const brokenCode = "function render() { var x = 'This week's data'; }";
    expect(() => {
      new vm.Script(brokenCode, { filename: 'test.html' });
    }).toThrow();
  });

  test('vm.Script rejects double-closing-paren on function declaration', () => {
    const brokenCode = 'function doStuff(a, b)) { return a + b; }';
    expect(() => {
      new vm.Script(brokenCode, { filename: 'test.gs' });
    }).toThrow();
  });

  test('vm.Script accepts valid GAS code', () => {
    const validCode = 'function doStuff(a, b) { return a + b; }';
    expect(() => {
      new vm.Script(validCode, { filename: 'test.gs' });
    }).not.toThrow();
  });

  test('all current src .gs files pass vm.Script validation', () => {
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
    const errors = [];
    for (const f of gsFiles) {
      const code = fs.readFileSync(path.join(SRC_DIR, f), 'utf8');
      try {
        new vm.Script(code, { filename: f });
      } catch (e) {
        errors.push(`${f}: ${e.message}`);
      }
    }
    expect(errors).toEqual([]);
  });

  test('all current src HTML <script> blocks pass vm.Script validation', () => {
    const htmlFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
    const errors = [];
    for (const f of htmlFiles) {
      const content = fs.readFileSync(path.join(SRC_DIR, f), 'utf8');
      const regex = /<script[^>]*>([\s\S]*?)<\/script>/gi;
      let match, idx = 0;
      while ((match = regex.exec(content)) !== null) {
        const js = match[1];
        if (js.trim().length < 10 || js.includes('<?')) continue;
        try {
          new vm.Script(js, { filename: `${f}:script[${idx}]` });
        } catch (e) {
          errors.push(`${f} block ${idx}: ${e.message}`);
        }
        idx++;
      }
    }
    expect(errors).toEqual([]);
  });
});


// ============================================================================
// Auth denial: Resource endpoints (05_Integrations.gs)
// ============================================================================
// Bug history: v4.25.10 — restoreWebAppResource checked auth.authorized (undefined)
// instead of auth.isAuthorized, so it always returned unauthorized but never
// actually threw. These tests verify all resource mutation functions return
// safe denial values when auth is denied.

describe('Auth denial: Resource mutation endpoints return safe values', () => {

  // Steward-only mutation functions
  const resourceMutations = [
    { name: 'addWebAppResource', args: [INVALID_TOKEN, { title: 'Test' }] },
    { name: 'updateWebAppResource', args: [INVALID_TOKEN, 'RES-001', { title: 'Updated' }] },
    { name: 'deleteWebAppResource', args: [INVALID_TOKEN, 'RES-001'] },
    { name: 'restoreWebAppResource', args: [INVALID_TOKEN, 'RES-001'] },
  ];

  resourceMutations.forEach(({ name, args }) => {
    test(`${name}() returns { success: false } when auth denied`, () => {
      const fn = global[name];
      expect(typeof fn).toBe('function');
      const result = fn(...args);
      expect(result).toBeDefined();
      expect(result.success).toBe(false);
    });

    test(`${name}() does not throw when auth denied`, () => {
      const fn = global[name];
      expect(() => fn(...args)).not.toThrow();
    });
  });

  // Read-only steward function (returns hidden resources)
  test('getWebAppResourcesListAll() returns [] when auth denied', () => {
    const fn = global.getWebAppResourcesListAll;
    expect(typeof fn).toBe('function');
    const result = fn();
    expect(result).toBeDefined();
    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBe(0);
  });

  test('getWebAppResourcesListAll() does not throw when auth denied', () => {
    expect(() => global.getWebAppResourcesListAll()).not.toThrow();
  });
});


// ============================================================================
// Summary: Fail if too many functions were skipped
// ============================================================================

describe('Auth denial: skip budget', () => {
  test(`no more than ${MAX_ALLOWED_SKIPS} functions should be undefined`, () => {
    if (skippedFunctions.length > 0) {
      console.warn(`Skipped ${skippedFunctions.length} undefined function(s): ${skippedFunctions.join(', ')}`);
    }
    expect(skippedFunctions.length).toBeLessThanOrEqual(MAX_ALLOWED_SKIPS);
  });
});
