/**
 * SPA Integrity Guards — Prevents the exact bugs found on 2026-03-08/09:
 *
 *   G8:  Every sidebar tab must have a matching route handler in _handleTabNav
 *   G9:  Mobile More menus must include all sidebar tabs (no desktop-only features)
 *   G10: isDualRole must include stewards (all stewards are members)
 *   G11: _ensureAllSheetsInternal must cover all feature sheets
 *   G12: Every client serverCall().funcName() must have a server-side function
 *
 * Run: npx jest test/spa-integrity.test.js --verbose
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.resolve(__dirname, '..', 'src');

// Helper: read file content
function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

// Helper: extract array items from a pattern like { id: 'xxx', ... }
function extractIds(code, startPattern) {
  const startIdx = code.indexOf(startPattern);
  if (startIdx === -1) return [];
  // Find the matching closing bracket
  let depth = 0;
  let started = false;
  let block = '';
  for (let i = startIdx; i < code.length; i++) {
    if (code[i] === '[') { depth++; started = true; }
    if (code[i] === ']') { depth--; }
    if (started) block += code[i];
    if (started && depth === 0) break;
  }
  const ids = [];
  const regex = /id:\s*'([^']+)'/g;
  let m;
  while ((m = regex.exec(block)) !== null) {
    ids.push(m[1]);
  }
  return ids;
}


// ============================================================================
// G8: SIDEBAR TABS → ROUTE HANDLERS
// ============================================================================
// Bug: POMS was in sidebar tabs but had no case in the steward switch block,
// causing it to fall through to default (renderStewardDashboard).
// Fix: POMS was actually handled as a shared tab before the role switch.
// This test ensures every sidebar tab ID maps to a handler somewhere.

describe('G8: Every sidebar tab has a route handler', () => {
  const indexCode = read('index.html');

  // Extract steward sidebar tabs
  const stewardTabs = extractIds(indexCode, "if (role === 'steward') return [");
  // Extract member sidebar tabs
  const memberTabs = extractIds(indexCode, "return [\n        { id: 'home'");

  // Extract all handled tab IDs from _handleTabNav
  const handledIds = new Set();
  // Shared tabs (before role switch)
  const sharedPattern = /if\s*\(tabId\s*===\s*'([^']+)'\)/g;
  let sm;
  while ((sm = sharedPattern.exec(indexCode)) !== null) {
    handledIds.add(sm[1]);
  }
  // Role-specific switch cases
  const casePattern = /case\s*'([^']+)':/g;
  while ((sm = casePattern.exec(indexCode)) !== null) {
    handledIds.add(sm[1]);
  }
  // 'more' and 'more_steward' are special mobile-only IDs
  handledIds.add('more');
  handledIds.add('more_steward');

  test('all steward sidebar tabs have route handlers', () => {
    const unhandled = stewardTabs.filter(id => !handledIds.has(id));
    if (unhandled.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Unhandled steward tabs:', unhandled);
    }
    expect(unhandled).toEqual([]);
  });

  test('all member sidebar tabs have route handlers', () => {
    const unhandled = memberTabs.filter(id => !handledIds.has(id));
    if (unhandled.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Unhandled member tabs:', unhandled);
    }
    expect(unhandled).toEqual([]);
  });
});


// ============================================================================
// G9: MOBILE MORE MENU PARITY
// ============================================================================
// Bug: POMS was in sidebar (desktop) but missing from steward and member
// More menus (mobile). Mobile users couldn't access POMS at all.

describe('G9: Mobile More menus cover all sidebar tabs', () => {
  const indexCode = read('index.html');
  const stewardCode = read('steward_view.html');
  const memberCode = read('member_view.html');

  // Sidebar tabs from _getSidebarTabs
  const stewardSidebarTabs = extractIds(indexCode, "if (role === 'steward') return [");
  const memberSidebarTabs = extractIds(indexCode, "return [\n        { id: 'home'");

  // Tabs visible in mobile bottom nav (always shown, never need More menu)
  const stewardBottomNav = ['cases', 'members', 'insights'];
  const memberBottomNav = ['home', 'cases', 'contact'];

  // Extract tab IDs referenced in More menus by looking for _handleTabNav calls
  // and render*() calls that map to specific tab IDs
  function extractMoreMenuTabIds(code, functionName) {
    const funcStart = code.indexOf('function ' + functionName);
    if (funcStart === -1) return [];
    // Read until next top-level function
    let depth = 0;
    let started = false;
    let block = '';
    for (let i = funcStart; i < code.length; i++) {
      if (code[i] === '{') { depth++; started = true; }
      if (code[i] === '}') { depth--; }
      if (started) block += code[i];
      if (started && depth === 0) break;
    }
    const ids = new Set();
    // Match _handleTabNav('role', 'tabId')
    const navRegex = /_handleTabNav\(\s*'[^']*'\s*,\s*'([^']+)'\s*\)/g;
    let m;
    while ((m = navRegex.exec(block)) !== null) ids.add(m[1]);
    // Match render*Page or render*() calls and map to known tab IDs
    // Map render functions to the sidebar tab IDs they satisfy.
    // Some render functions handle multiple tabs (e.g., renderStewardContact handles
    // both 'contact' and 'stewarddirectory').
    const renderMap = {
      renderStewardDashboard: ['cases'],
      renderStewardMembers: ['members'],
      renderInsightsPage: ['insights'],
      renderBroadcast: ['broadcast'],
      renderStewardNotifications: ['notifications'],
      renderStewardResources: ['resources'],
      renderManagePolls: ['polls'],
      renderSurveyTracking: ['surveytrack'],
      renderContactLog: ['contactlog'],
      renderStewardTasks: ['tasks'],
      renderEventsPage: ['events'],
      renderStewardMinutes: ['minutes'],
      renderStewardDirectoryPage: ['stewarddirectory'],
      renderTimelinePage: ['timeline'],
      renderStewardTimelinePage: ['timeline'],
      renderFeedbackPage: ['feedback'],
      renderFailsafePage: ['failsafe'],
      renderTestRunnerPage: ['testrunner'],
      renderQAForum: ['qaforum'],
      renderMemberHome: ['home'],
      renderMyCases: ['cases'],
      renderStewardContact: ['contact', 'stewarddirectory'],
      renderUpdateProfile: ['profile'],
      renderMemberResources: ['resources'],
      renderWorkloadTracker: ['workload'],
      renderPollsPage: ['polls'],
      renderSurveyResultsPage: ['survey'],
      renderUnionStatsPage: ['unionstats'],
      renderMemberNotifications: ['notifications'],
      renderMeetingsPage: ['meetings'],
      renderMinutesPage: ['minutes'],
      renderMemberTasks: ['mytasks'],
      renderSurveyFormPage: ['surveyform'],
    };
    for (const [funcName, tabIds] of Object.entries(renderMap)) {
      if (block.includes(funcName)) tabIds.forEach(id => ids.add(id));
    }
    return [...ids];
  }

  const stewardMoreIds = extractMoreMenuTabIds(stewardCode, 'renderStewardMore');
  const memberMoreIds = extractMoreMenuTabIds(memberCode, 'renderMemberMore');

  test('steward More menu covers all non-bottom-nav sidebar tabs', () => {
    const needsMoreMenu = stewardSidebarTabs.filter(id => !stewardBottomNav.includes(id));
    const missing = needsMoreMenu.filter(id => !stewardMoreIds.includes(id));
    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Steward sidebar tabs missing from mobile More menu:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('member More menu covers all non-bottom-nav sidebar tabs', () => {
    const needsMoreMenu = memberSidebarTabs.filter(id => !memberBottomNav.includes(id));
    const missing = needsMoreMenu.filter(id => !memberMoreIds.includes(id));
    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Member sidebar tabs missing from mobile More menu:', missing);
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// G10: DUAL-ROLE TOGGLE — STEWARDS MUST BE DUAL-ROLE
// ============================================================================
// Bug: isDualRole was only true when role === 'both'. Since most stewards have
// role === 'steward' (not 'both'), the member view toggle never appeared.

describe('G10: isDualRole includes stewards', () => {
  const webAppCode = read('22_WebDashApp.gs');

  test('isDualRole condition includes role === steward', () => {
    // Find the isDualRole line in pageData construction
    const match = webAppCode.match(/isDualRole:\s*(.+)/);
    expect(match).not.toBeNull();

    const condition = match[1];
    // Must include steward check — not just 'both'
    expect(condition).toMatch(/role\s*===\s*'steward'/);
  });

  test('isDualRole is NOT just role === both', () => {
    const match = webAppCode.match(/isDualRole:\s*(.+)/);
    const condition = match[1].trim().replace(/,$/, '');
    // Should not be exactly "role === 'both'" — that's the old broken condition
    expect(condition).not.toBe("role === 'both'");
  });
});


// ============================================================================
// G11: SHEET INIT COMPLETENESS
// ============================================================================
// Bug: Timeline, Q&A Forum, and other sheets weren't auto-created because
// _ensureAllSheetsInternal didn't exist. Now it does, but future features
// could add new sheets and forget to register them.

describe('G11: _ensureAllSheetsInternal covers all feature sheets', () => {
  const dataServiceCode = read('21_WebDashDataService.gs');
  const coreCode = read('01_Core.gs');

  // Extract hidden/feature sheet names from SHEETS constant
  function extractSheetConstants(code) {
    const sheets = {};
    // Match patterns like: TIMELINE_EVENTS: '_Timeline_Events',
    // and: NOTIFICATIONS: '📢 Notifications',
    const regex = /(\w+):\s*['"]([^'"]+)['"]/g;
    // Find the SHEETS = { ... } block
    const sheetsStart = code.indexOf('var SHEETS = {');
    if (sheetsStart === -1) return sheets;
    let depth = 0;
    let started = false;
    let block = '';
    for (let i = sheetsStart; i < code.length; i++) {
      if (code[i] === '{') { depth++; started = true; }
      if (code[i] === '}') { depth--; }
      if (started) block += code[i];
      if (started && depth === 0) break;
    }
    let m;
    while ((m = regex.exec(block)) !== null) {
      sheets[m[1]] = m[2];
    }
    return sheets;
  }

  const allSheets = extractSheetConstants(coreCode);

  // Feature sheets that are created by specific init functions and should be
  // in _ensureAllSheetsInternal. Excludes core sheets (Member Dir, Grievance Log, Config)
  // which are created by CREATE_DASHBOARD's main flow, and documentation/reference sheets.
  const featureSheets = [
    'CONTACT_LOG', 'STEWARD_TASKS', 'QA_FORUM', 'QA_ANSWERS',
    'TIMELINE_EVENTS', 'FAILSAFE_CONFIG', 'WEEKLY_QUESTIONS',
    'WEEKLY_RESPONSES', 'QUESTION_POOL', 'NOTIFICATIONS', 'AUDIT_LOG',
    'RESOURCES', 'RESOURCE_CONFIG', 'FEEDBACK', 'CASE_CHECKLIST',
    'SURVEY_QUESTIONS', 'SATISFACTION',
  ];

  test('_ensureAllSheetsInternal references all feature sheet init functions', () => {
    // Extract the _ensureAllSheetsInternal function body
    const funcStart = dataServiceCode.indexOf('function _ensureAllSheetsInternal()');
    expect(funcStart).toBeGreaterThan(-1);

    let depth = 0;
    let started = false;
    let body = '';
    for (let i = funcStart; i < dataServiceCode.length; i++) {
      if (dataServiceCode[i] === '{') { depth++; started = true; }
      if (dataServiceCode[i] === '}') { depth--; }
      if (started) body += dataServiceCode[i];
      if (started && depth === 0) break;
    }

    // Map: init function/module calls → sheets they create.
    // If the init function is in the body, all its child sheets are covered.
    const initCoversSheets = {
      'initQAForumSheets':         ['QA_FORUM', 'QA_ANSWERS'],
      'initTimelineSheet':         ['TIMELINE_EVENTS'],
      'initFailsafeSheet':         ['FAILSAFE_CONFIG'],
      'initWeeklyQuestionSheets':  ['WEEKLY_QUESTIONS', 'WEEKLY_RESPONSES', 'QUESTION_POOL'],
      'initPortalSheets':          [], // portal sheets not in SHEETS constant
      'initWorkloadTrackerSheets': ['WORKLOAD_VAULT', 'WORKLOAD_REPORTING', 'WORKLOAD_REMINDERS', 'WORKLOAD_USERMETA'],
      'setupHiddenSheets':         [], // hidden calc sheets
      'createResourcesSheet':      ['RESOURCES'],
      'createResourceConfigSheet': ['RESOURCE_CONFIG'],
      'createSurveyQuestionsSheet':['SURVEY_QUESTIONS'],
      'createSatisfactionSheet':   ['SATISFACTION'],
      'createFeedbackSheet':       ['FEEDBACK'],
      'getOrCreateChecklistSheet': ['CASE_CHECKLIST'],
      '_ensureContactLogSheet':    ['CONTACT_LOG'],
      '_ensureStewardTasksSheet':  ['STEWARD_TASKS'],
    };

    // Build set of all sheets covered by init functions present in the body
    const coveredSheets = new Set();
    for (const [funcName, sheets] of Object.entries(initCoversSheets)) {
      if (body.includes(funcName)) {
        sheets.forEach(s => coveredSheets.add(s));
      }
    }
    // Also check for direct SHEETS.* references or sheet name strings
    for (const key of featureSheets) {
      const sheetName = allSheets[key];
      if (body.includes(sheetName) || body.includes('SHEETS.' + key)) {
        coveredSheets.add(key);
      }
    }

    const missing = featureSheets.filter(key => !coveredSheets.has(key));

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Feature sheets missing from auto-init:\n  ' +
        missing.map(k => `${k} (${allSheets[k]})`).join('\n  '));
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// G12: EVERY CLIENT SERVER CALL HAS A MATCHING SERVER FUNCTION
// ============================================================================
// Bug: If a client calls serverCall().someFunc() but someFunc doesn't exist
// server-side, the call fails silently (or with opaque "Something went wrong").

describe('G12: All client server calls have matching server functions', () => {
  // Build set of all top-level server function names across all .gs files
  const serverFunctions = new Set();
  const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
  for (const f of gsFiles) {
    const code = read(f);
    const regex = /^function\s+(\w+)\s*\(/gm;
    let m;
    while ((m = regex.exec(code)) !== null) {
      serverFunctions.add(m[1]);
    }
  }

  // Standard JS methods to exclude from matching
  const JS_BUILTINS = new Set([
    'withSuccessHandler', 'withFailureHandler', 'withUserObject',
    'getElementById', 'querySelector', 'querySelectorAll', 'addEventListener',
    'appendChild', 'removeChild', 'insertBefore', 'replaceChild',
    'setAttribute', 'getAttribute', 'removeAttribute', 'remove',
    'toString', 'toLowerCase', 'toUpperCase', 'trim', 'replace',
    'split', 'join', 'filter', 'map', 'forEach', 'push', 'pop',
    'slice', 'splice', 'sort', 'concat', 'indexOf', 'includes',
    'find', 'findIndex', 'reduce', 'some', 'every', 'keys', 'values',
    'entries', 'stringify', 'parse', 'log', 'error', 'warn', 'info',
    'getTime', 'getFullYear', 'getMonth', 'getDate', 'getDay',
    'getHours', 'getMinutes', 'toISOString', 'toLocaleDateString',
    'clearInterval', 'setInterval', 'setTimeout', 'clearTimeout',
    'reload', 'focus', 'blur', 'click', 'contains', 'padStart',
    'preventDefault', 'stopPropagation', 'toFixed', 'charAt',
    'substring', 'localeCompare', 'startsWith', 'endsWith',
    'apply', 'call', 'bind', 'resolve', 'reject', 'then', 'catch',
  ]);

  const clientFiles = ['index.html', 'steward_view.html', 'member_view.html', 'auth_view.html'];

  test('no client calls reference non-existent server functions', () => {
    const missing = [];

    for (const file of clientFiles) {
      const filePath = path.join(SRC_DIR, file);
      if (!fs.existsSync(filePath)) continue;
      const code = read(file);
      const lines = code.split('\n');

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trim().startsWith('//') || line.trim().startsWith('*')) continue;

        // Only check lines that are part of server call chains
        const isServerLine = line.includes('serverCall(') ||
          line.includes('google.script.run') ||
          line.includes('.withSuccessHandler') ||
          line.includes('.withFailureHandler');
        if (!isServerLine) continue;

        // Match terminal .funcName( calls
        const regex = /\.(\w+)\s*\(/g;
        let m;
        while ((m = regex.exec(line)) !== null) {
          const name = m[1];
          if (JS_BUILTINS.has(name)) continue;
          if (serverFunctions.has(name)) continue;
          // Skip if it looks like a client-side function (starts lowercase, common patterns)
          if (['serverCall', 'showLoading', 'showToast', 'el', 'escHtml_'].includes(name)) continue;

          missing.push(`${file}:${i + 1} calls .${name}() — no matching server function`);
        }
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Missing server functions:\n  ' + missing.join('\n  '));
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// G13: SERVER CALL ERROR HANDLER SHOWS ACTUAL ERRORS
// ============================================================================
// Bug: serverCall() default failure handler showed "Something went wrong"
// with no detail, making all server errors impossible to diagnose.

describe('G13: serverCall failure handler surfaces error details', () => {
  const indexCode = read('index.html');

  test('serverCall failure handler includes err.message', () => {
    // Find the serverCall function
    const funcMatch = indexCode.match(/function serverCall\(\)\s*\{[\s\S]*?return runner;\s*\}/);
    expect(funcMatch).not.toBeNull();
    const body = funcMatch[0];

    // Must reference err.message or err to show actual error
    expect(body).toMatch(/err\.message|err\b/);
    // Must NOT just say "Something went wrong" without the error
    // (it's OK if it says that AND includes the error detail)
    const hasGenericOnly = body.includes('Something went wrong') && !body.includes('err');
    expect(hasGenericOnly).toBe(false);
  });
});


// ============================================================================
// G14: MEMBER LIST ITEMS ARE CLICKABLE
// ============================================================================
// Bug: Member list items had click handlers in JS but no CSS cursor/hover,
// making them look non-interactive.

describe('G14: Member list items have click affordance', () => {
  const stylesCode = read('styles.html');
  const stewardCode = read('steward_view.html');

  test('member-list-item has cursor: pointer in CSS', () => {
    // Extract .member-list-item rule
    const ruleMatch = stylesCode.match(/\.member-list-item\s*\{[^}]+\}/);
    expect(ruleMatch).not.toBeNull();
    expect(ruleMatch[0]).toMatch(/cursor:\s*pointer/);
  });

  test('member-list-item has a hover state', () => {
    expect(stylesCode).toMatch(/\.member-list-item:hover\s*\{/);
  });

  test('member rows have click event listeners', () => {
    // renderFilteredMembers must attach click handlers
    expect(stewardCode).toMatch(/item\.addEventListener\('click'/);
    // And the handler must call _toggleMemberDetail
    expect(stewardCode).toMatch(/_toggleMemberDetail\(container,\s*item,\s*m\)/);
  });
});


// ============================================================================
// G15: BOTTOM NAV TAB IDs HAVE ROUTE HANDLERS
// ============================================================================
// Bug (2026-03-11): renderBottomNav member tabs had id:'contact' but
// _handleTabNav only handled case 'stewarddirectory'. Unknown IDs fell through
// to default (renderMemberHome), silently sending the user back to home on
// every Contact tab click. G8 only checked sidebar tabs — not bottom nav tabs.
// This guard catches the same class of mismatch for bottom nav specifically.

describe('G15: Bottom nav tab IDs have route handlers', () => {
  const stewardCode = read('steward_view.html');
  const indexCode   = read('index.html');

  // Extract tab IDs from renderBottomNav (both steward and member arrays)
  function extractBottomNavIds(code) {
    const funcStart = code.indexOf('function renderBottomNav');
    if (funcStart === -1) return [];
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < code.length; i++) {
      if (code[i] === '{') { depth++; started = true; }
      if (code[i] === '}') depth--;
      if (started) block += code[i];
      if (started && depth === 0) break;
    }
    return [...block.matchAll(/id:\s*'([^']+)'/g)].map(m => m[1]);
  }

  // Extract all handled tab IDs from _handleTabNav
  function extractHandledIds(code) {
    const handled = new Set();
    // Shared if-checks (e.g. if (tabId === 'orgchart'))
    for (const m of code.matchAll(/if\s*\(tabId\s*===\s*'([^']+)'\)/g)) handled.add(m[1]);
    // Switch case labels
    for (const m of code.matchAll(/case\s*'([^']+)':/g)) handled.add(m[1]);
    // Special internal IDs: More menus trigger by their internal name
    handled.add('_member_more');
    handled.add('_steward_more');
    // 'more' and 'more_steward' are translated inside renderBottomNav onClick
    // before _handleTabNav is called — they are not routed directly
    handled.add('more');
    handled.add('more_steward');
    return handled;
  }

  const bottomNavIds = extractBottomNavIds(stewardCode);
  const handledIds   = extractHandledIds(indexCode);

  test('all bottom nav tab IDs have route handlers in _handleTabNav', () => {
    // 'more' / 'more_steward' are UI-only sentinels — translated to
    // '_member_more' / '_steward_more' before _handleTabNav is called.
    // Everything else must have an explicit handler.
    const sentinels = new Set(['more', 'more_steward']);
    const unhandled = bottomNavIds.filter(id => !sentinels.has(id) && !handledIds.has(id));
    if (unhandled.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Bottom nav tab IDs with no route handler:', unhandled);
    }
    expect(unhandled).toEqual([]);
  });

  test('no duplicate tab IDs in bottom nav', () => {
    const seen = new Set();
    const dupes = [];
    for (const id of bottomNavIds) {
      if (seen.has(id) && !['cases'].includes(id)) dupes.push(id); // 'cases' appears once per role array — OK
      seen.add(id);
    }
    expect(dupes).toEqual([]);
  });
});


// ============================================================================
// G16: renderPageLayout ALWAYS RENDERS BOTH SIDEBAR AND BOTTOM NAV
// ============================================================================
// Bug (2026-03-11): renderPageLayout used an if/else branch driven by
// isSidebarLayout() (JS viewport detection). If JS picked the wrong branch,
// ZERO nav was rendered into the DOM — user had no navigation at all.
// Fix: always add both sidebar-nav and bottom-nav; CSS controls visibility.
// This guard prevents a revert to the conditional pattern.

describe('G16: renderPageLayout always renders both sidebar and bottom nav', () => {
  const indexCode = read('index.html');

  // Extract renderPageLayout function body
  function extractRenderPageLayout(code) {
    const funcStart = code.indexOf('function renderPageLayout(');
    if (funcStart === -1) return '';
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < code.length; i++) {
      if (code[i] === '{') { depth++; started = true; }
      if (code[i] === '}') depth--;
      if (started) block += code[i];
      if (started && depth === 0) break;
    }
    return block;
  }

  const body = extractRenderPageLayout(indexCode);

  test('renderPageLayout calls renderSidebarItems unconditionally', () => {
    expect(body).toContain('renderSidebarItems(');
    // Must NOT be inside an if (isSidebarLayout()) conditional
    // Simple check: renderSidebarItems must appear without a preceding isSidebarLayout guard
    const guardedPattern = /if\s*\(\s*isSidebarLayout\(\)\s*\)[\s\S]*?renderSidebarItems/;
    expect(guardedPattern.test(body)).toBe(false);
  });

  test('renderPageLayout calls renderBottomNav unconditionally', () => {
    expect(body).toContain('renderBottomNav(');
    // Must NOT be inside an else block tied to isSidebarLayout
    const guardedPattern = /else\s*\{[\s\S]*?renderBottomNav/;
    expect(guardedPattern.test(body)).toBe(false);
  });

  test('renderPageLayout does not use if/else to choose between navs', () => {
    // The old bug: if (isSidebarLayout()) { sidebar only } else { bottomNav only }
    // This pattern should no longer exist
    const branchPattern = /if\s*\(\s*isSidebarLayout\(\)\s*\)\s*\{[\s\S]*?\}\s*else\s*\{/;
    expect(branchPattern.test(body)).toBe(false);
  });
});
