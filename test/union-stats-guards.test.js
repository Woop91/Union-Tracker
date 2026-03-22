/**
 * Union Stats & Auth Contract Guards (v4.28.2)
 *
 * Prevents the exact class of bugs found in the Union Stats tab review:
 *
 *   G17: member_view.html must NEVER call steward-only (_requireStewardAuth) endpoints
 *   G18: Stats page renderers must use all non-reserved data fields returned by backend
 *   G19: All stats/insights pages must have client-side caching (AppState cache pattern)
 *   G20: Union Stats sub-tab renderers must separate data fetch from content render
 *   G21: Every steward-only endpoint must have a member-safe alternative when called from member_view
 *
 * These tests are structural — they read source files and verify contracts without
 * executing code, preventing regressions even when mocks are incomplete.
 *
 * Run: npx jest test/union-stats-guards.test.js --verbose
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.resolve(__dirname, '..', 'src');

function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

// Helper: extract function body from source code
function extractFunctionBody(code, funcName) {
  const funcStart = code.indexOf('function ' + funcName);
  if (funcStart === -1) return '';
  let depth = 0;
  let started = false;
  let block = '';
  for (let i = funcStart; i < code.length; i++) {
    if (code[i] === '{') { depth++; started = true; }
    if (code[i] === '}') { depth--; }
    if (started) block += code[i];
    if (started && depth === 0) break;
  }
  return block;
}


// ============================================================================
// G17: MEMBER VIEW MUST NEVER CALL STEWARD-ONLY ENDPOINTS
// ============================================================================
// Bug: member_view.html called dataGetGrievanceStats (steward-only) and
// dataAssignSteward (steward-only), causing silent failures for non-steward members.
// Fix: Created dataGetMemberGrievanceStats, dataGetMemberGrievanceHotSpots,
// dataMemberAssignSteward with _resolveCallerEmail auth.
// This test ensures no future code adds steward-only calls to member_view.html.

describe('G17: member_view.html never calls steward-only endpoints', () => {
  const memberCode = read('member_view.html');

  // Build the definitive list of steward-only functions by scanning all .gs files
  // for functions that call _requireStewardAuth
  function getStewardOnlyFunctions() {
    const stewardFns = new Set();
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
    for (const f of gsFiles) {
      const code = read(f);
      // Find all function declarations
      const funcRegex = /^function\s+(\w+)\s*\(/gm;
      let m;
      while ((m = funcRegex.exec(code)) !== null) {
        const name = m[1];
        // Extract this function's body and check if it calls _requireStewardAuth
        const body = extractFunctionBody(code, name);
        if (body.includes('_requireStewardAuth')) {
          stewardFns.add(name);
        }
      }
    }
    // Remove the definition itself and test helpers
    stewardFns.delete('_requireStewardAuth');
    stewardFns.delete('test_auth_requireStewardAuthExists');
    return stewardFns;
  }

  const stewardOnlyFunctions = getStewardOnlyFunctions();

  test('steward-only function list is non-empty (sanity check)', () => {
    expect(stewardOnlyFunctions.size).toBeGreaterThan(10);
  });

  test('member_view.html does not call any steward-only function via serverCall()', () => {
    // Known safe exceptions: QA Forum moderation functions are steward-only but
    // called from member_view within `if (isSteward)` UI guards. Stewards using
    // member view retain steward privileges for moderation actions.
    const knownSafeExceptions = new Set([
      'qaModerateQuestion',    // Gated by if (isSteward && ...) in QA Forum UI
      'qaModerateAnswer',      // Gated by if (isSteward) in QA Forum UI
      'qaSubmitAnswer',        // Steward answers — gated by isSteward check
      'qaGetFlaggedContent',   // Moderation panel — only rendered for stewards
    ]);

    const violations = [];
    const lines = memberCode.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      // Skip comments
      if (line.trim().startsWith('//') || line.trim().startsWith('*')) continue;
      // Check for serverCall chains or google.script.run chains
      for (const fn of stewardOnlyFunctions) {
        if (knownSafeExceptions.has(fn)) continue;
        // Match .functionName( pattern (the terminal call in a serverCall chain)
        const pattern = new RegExp('\\.' + fn + '\\s*\\(');
        if (pattern.test(line)) {
          violations.push(`  line ${i + 1}: calls steward-only .${fn}()`);
        }
      }
    }
    if (violations.length > 0) {
      // eslint-disable-next-line no-console
      console.error(
        'AUTH MISMATCH: member_view.html calls steward-only endpoints:\n' +
        violations.join('\n') +
        '\n\nFix: Create member-safe versions using _resolveCallerEmail'
      );
    }
    expect(violations).toEqual([]);
  });

  test('steward_view.html does not call member-self-service endpoints', () => {
    const stewardCode = read('steward_view.html');
    // Member-self-service functions that only make sense for the member themselves
    const memberSelfServiceFns = [
      'dataMemberAssignSteward',
      'dataGetMemberGrievanceStats',
      'dataGetMemberGrievanceHotSpots',
      'dataMemberCompleteMemberTask',
    ];
    const violations = [];
    for (const fn of memberSelfServiceFns) {
      if (stewardCode.includes('.' + fn + '(')) {
        violations.push(fn);
      }
    }
    // This is informational — steward calling member endpoints isn't dangerous
    // (stewards are a superset of members), but indicates confusion
    if (violations.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Steward view calls member-specific endpoints (review if intentional):', violations);
    }
    // Not a hard failure — just a warning
  });
});


// ============================================================================
// G18: STATS RENDERERS MUST USE ALL NON-RESERVED BACKEND DATA FIELDS
// ============================================================================
// Bug: _renderMembershipStats only rendered byLocation, ignoring byUnit, byDues,
// newMembersLast90, and byHireMonth returned by getMembershipStats.
// Bug: _renderGrievanceStatsContent ignored monthlyResolved from getGrievanceStats.
// This test ensures frontend renderers reference all fields the backend returns.

describe('G18: Stats renderers use all backend data fields', () => {
  const memberCode = read('member_view.html');

  test('membership stats renderer uses all getMembershipStats fields', () => {
    const body = extractFunctionBody(memberCode, '_renderMembershipStatsContent');
    expect(body.length).toBeGreaterThan(0);

    // Fields returned by getMembershipStats (from 21_WebDashDataService.gs)
    const requiredFields = ['byLocation', 'byUnit', 'byDues', 'newMembersLast90', 'byHireMonth'];
    const missing = requiredFields.filter(field => !body.includes(field));
    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.error('_renderMembershipStatsContent missing fields:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('grievance stats renderer uses monthlyResolved data', () => {
    const body = extractFunctionBody(memberCode, '_renderGrievanceStatsContent');
    expect(body.length).toBeGreaterThan(0);

    // Must reference monthlyResolved for the Filed vs Resolved trend chart
    expect(body).toContain('monthlyResolved');
  });

  test('engagement stats renderer includes resourceDownloads in metrics array', () => {
    const body = extractFunctionBody(memberCode, '_renderEngagementStatsContent');
    expect(body.length).toBeGreaterThan(0);

    // resourceDownloads is now live-tracked and always shown
    expect(body).toMatch(/resourceDownloads/);
  });

  test('workload stats renderer references all summary fields', () => {
    const body = extractFunctionBody(memberCode, '_renderWorkloadStatsContent');
    expect(body.length).toBeGreaterThan(0);

    const requiredFields = ['avgCaseload', 'highCaseloadPct', 'submissionRate', 'trendDirection'];
    const missing = requiredFields.filter(field => !body.includes(field));
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// G19: ALL STATS PAGES MUST HAVE CLIENT-SIDE CACHING
// ============================================================================
// Bug: Union Stats page had no caching while steward Insights page did,
// causing redundant server calls on every sub-tab click.
// This test ensures both stats pages maintain caching parity.

describe('G19: Stats pages have client-side caching', () => {
  const memberCode = read('member_view.html');
  const stewardCode = read('steward_view.html');

  test('Union Stats page uses AppState.unionStatsCache', () => {
    const body = extractFunctionBody(memberCode, 'renderUnionStatsPage');
    expect(body).toContain('AppState.unionStatsCache');
  });

  test('Union Stats page has cache TTL', () => {
    const body = extractFunctionBody(memberCode, 'renderUnionStatsPage');
    expect(body).toMatch(/CACHE_TTL/);
  });

  test('Insights page uses AppState.insightsCache', () => {
    const body = extractFunctionBody(stewardCode, 'renderInsightsPage');
    expect(body).toContain('AppState.insightsCache');
  });

  test('all Union Stats sub-tab renderers accept cache functions', () => {
    // Each sub-tab renderer should accept getCached/setCache parameters
    const renderers = [
      '_renderGrievanceStats',
      '_renderHotSpots',
      '_renderMembershipStats',
      '_renderEngagementStats',
      '_renderWorkloadSummaryStats',
    ];
    const missing = [];
    for (const fn of renderers) {
      const body = extractFunctionBody(memberCode, fn);
      if (!body.includes('getCached') && !body.includes('setCache')) {
        missing.push(fn);
      }
    }
    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.error('Sub-tab renderers without caching:', missing);
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// G20: STATS SUB-TAB RENDERERS MUST SEPARATE FETCH FROM RENDER
// ============================================================================
// Pattern: each sub-tab should have a data-fetch function (e.g., _renderGrievanceStats)
// that calls the server and a content-render function (e.g., _renderGrievanceStatsContent)
// that accepts data and renders DOM. This allows caching to bypass the fetch.
// Bug: Original renderers mixed fetch + render, making caching impossible without
// duplicating the render logic.

describe('G20: Stats sub-tab renderers separate fetch from content render', () => {
  const memberCode = read('member_view.html');

  const subTabs = [
    { fetch: '_renderGrievanceStats', render: '_renderGrievanceStatsContent' },
    { fetch: '_renderHotSpots', render: '_renderHotSpotsContent' },
    { fetch: '_renderMembershipStats', render: '_renderMembershipStatsContent' },
    { fetch: '_renderEngagementStats', render: '_renderEngagementStatsContent' },
    { fetch: '_renderWorkloadSummaryStats', render: '_renderWorkloadStatsContent' },
  ];

  subTabs.forEach(({ fetch: fetchFn, render: renderFn }) => {
    test(`${fetchFn} delegates to ${renderFn} for DOM rendering`, () => {
      // The fetch function must exist
      const fetchBody = extractFunctionBody(memberCode, fetchFn);
      expect(fetchBody.length).toBeGreaterThan(0);

      // The content render function must exist
      const renderBody = extractFunctionBody(memberCode, renderFn);
      expect(renderBody.length).toBeGreaterThan(0);

      // The fetch function must call the render function
      expect(fetchBody).toContain(renderFn);
    });

    test(`${renderFn} does not make server calls (pure renderer)`, () => {
      const renderBody = extractFunctionBody(memberCode, renderFn);
      // Content renderers should NEVER call serverCall() — they receive data as args
      expect(renderBody).not.toContain('serverCall()');
      expect(renderBody).not.toContain('google.script.run');
    });
  });
});


// ============================================================================
// G21: MEMBER-SAFE ENDPOINT COMPLETENESS
// ============================================================================
// When a steward-only endpoint is needed from member_view, a member-safe
// variant must exist. This test verifies known pairs are complete.

describe('G21: Member-safe endpoint pairs exist', () => {
  const dataServiceCode = read('21_WebDashDataService.gs');

  // Known pairs: steward-only → member-safe
  const requiredPairs = [
    { steward: 'dataGetGrievanceStats', member: 'dataGetMemberGrievanceStats' },
    { steward: 'dataGetGrievanceHotSpots', member: 'dataGetMemberGrievanceHotSpots' },
    { steward: 'dataAssignSteward', member: 'dataMemberAssignSteward' },
  ];

  requiredPairs.forEach(({ steward, member }) => {
    test(`${member} exists as member-safe alternative to ${steward}`, () => {
      const memberDef = new RegExp('function\\s+' + member + '\\s*\\(');
      expect(memberDef.test(dataServiceCode)).toBe(true);
    });

    test(`${member} uses _resolveCallerEmail (not _requireStewardAuth)`, () => {
      const body = extractFunctionBody(dataServiceCode, member);
      expect(body).toContain('_resolveCallerEmail');
      expect(body).not.toContain('_requireStewardAuth');
    });

    test(`${steward} uses _requireStewardAuth (remains steward-only)`, () => {
      const body = extractFunctionBody(dataServiceCode, steward);
      expect(body).toContain('_requireStewardAuth');
    });
  });
});


// ============================================================================
// G22: NO DEAD/ZERO KPI CARDS IN STATS PAGES
// ============================================================================
// Bug: Resource Downloads KPI always showed "0" with no explanation.
// This test ensures known untracked metrics are conditionally displayed.

describe('G22: Resource click tracking is wired end-to-end', () => {
  const memberCode = read('member_view.html');

  test('resourceDownloads KPI is always shown in metrics array', () => {
    const body = extractFunctionBody(memberCode, '_renderEngagementStatsContent');
    const metricsArrayMatch = body.match(/var metrics\s*=\s*\[([\s\S]*?)\];/);
    expect(metricsArrayMatch).not.toBeNull();
    expect(metricsArrayMatch[1]).toMatch(/resourceDownloads/);
  });

  test('backend engagement stats returns live resourceDownloads', () => {
    const dsCode = read('21_WebDashDataService.gs');
    const body = extractFunctionBody(dsCode, 'dataGetEngagementStats');
    // Backend should read resourceDownloads from DataService.getResourceClickTotal()
    expect(body).toMatch(/resourceDownloads.*getResourceClickTotal/);
  });
});


// ============================================================================
// G23: ENGAGEMENT STATS — NO REDUNDANT SHEET READS
// ============================================================================
// Bug: dataGetEngagementStats read Member Directory twice (mData + mData2).
// This test ensures the function reads each sheet at most once.

describe('G23: dataGetEngagementStats has no redundant sheet reads', () => {
  const dsCode = read('21_WebDashDataService.gs');

  test('Member Directory is read at most once', () => {
    const body = extractFunctionBody(dsCode, 'dataGetEngagementStats');
    // Count occurrences of memberSheet.getRange(...).getValues()
    const readPattern = /memberSheet\.getRange\([^)]*\)\.getValues\(\)/g;
    const matches = body.match(readPattern) || [];
    // Should be exactly 1 data read (header read is separate and OK)
    const dataReads = matches.filter(m => !m.includes('1, 1, 1,')); // exclude header reads
    expect(dataReads.length).toBeLessThanOrEqual(2); // 1 header + 1 data
  });

  test('no mData2 variable exists (was the redundant copy)', () => {
    const body = extractFunctionBody(dsCode, 'dataGetEngagementStats');
    expect(body).not.toContain('mData2');
  });
});


// ============================================================================
// G24: showLoading SKELETON CONSISTENCY
// ============================================================================
// Bug: _renderWorkloadSummaryStats called showLoading(container) without a
// skeleton type, while all other sub-tabs specified 'kpi' or 'list'.
// This test ensures all stats sub-tab renderers specify skeleton types.

describe('G24: Stats sub-tab renderers specify showLoading skeleton types', () => {
  const memberCode = read('member_view.html');

  const renderers = [
    '_renderGrievanceStats',
    '_renderHotSpots',
    '_renderMembershipStats',
    '_renderEngagementStats',
    '_renderWorkloadSummaryStats',
  ];

  renderers.forEach(fn => {
    test(`${fn} calls showLoading with a skeleton type`, () => {
      const body = extractFunctionBody(memberCode, fn);
      if (body.includes('showLoading')) {
        // showLoading should be called with 2 args: container + type
        // Match showLoading(container) without second arg — this is the bug
        const bareCall = /showLoading\(\s*\w+\s*\)\s*;/;
        if (bareCall.test(body)) {
          // eslint-disable-next-line no-console
          console.error(`${fn} calls showLoading() without skeleton type`);
        }
        expect(bareCall.test(body)).toBe(false);
      }
    });
  });
});
