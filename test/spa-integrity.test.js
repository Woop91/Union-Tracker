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
  const sharedPattern = /if\s*\(tabId\s*===\s*'([^']+)'/g;
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
      renderMemberTimelinePage: ['timeline'],
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
    // serverCall() may delegate error handling to _showFailureUI or inline handler.
    // Check that the failure UI function (wherever it lives) references err.message.
    const failureUIMatch = indexCode.match(/function _showFailureUI\(err\)\s*\{[\s\S]*?\n {4}\}/);
    const inlineMatch = indexCode.match(/function serverCall\(\)\s*\{[\s\S]*?return runner;\s*\}/);
    const body = failureUIMatch ? failureUIMatch[0] : (inlineMatch ? inlineMatch[0] : '');
    expect(body.length).toBeGreaterThan(0);

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

// ====================================================================
// G21: My Tasks tab regression guards
// ============================================================================

describe('G21: My Tasks tab integrity', () => {
  const memberCode = read('member_view.html');
  const stewardCode = read('steward_view.html');
  const dataServiceCode = read('21_WebDashDataService.gs');

  test('renderMemberTasks passes explicit status filter (not null) for open tab', () => {
    // Extract the renderMemberTasks function body
    const funcStart = memberCode.indexOf('function renderMemberTasks');
    expect(funcStart).not.toBe(-1);
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < memberCode.length; i++) {
      if (memberCode[i] === '{') { depth++; started = true; }
      if (memberCode[i] === '}') depth--;
      if (started) block += memberCode[i];
      if (started && depth === 0) break;
    }
    // Should NOT pass null as statusFilter for the open tab
    // The old bug: var statusFilter = activeSubTab === 'completed' ? 'completed' : null;
    expect(block).not.toMatch(/statusFilter\s*=.*\?\s*'completed'\s*:\s*null/);
    // Should use 'not-completed' for the open tab
    expect(block).toContain('not-completed');
  });

  test('renderMemberTasks does not redundantly filter tasks client-side', () => {
    const funcStart = memberCode.indexOf('function renderMemberTasks');
    expect(funcStart).not.toBe(-1);
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < memberCode.length; i++) {
      if (memberCode[i] === '{') { depth++; started = true; }
      if (memberCode[i] === '}') depth--;
      if (started) block += memberCode[i];
      if (started && depth === 0) break;
    }
    // Should NOT have client-side filter like: tasks.filter(function(t) { return t.status !== 'completed'; })
    expect(block).not.toMatch(/tasks\.filter\(function\(t\)\s*\{\s*return t\.status !== 'completed'/);
  });

  test('description truncation includes ellipsis indicator', () => {
    // The old bug: task.description.substring(0, 100) — no ellipsis
    const funcStart = memberCode.indexOf('function renderMemberTasks');
    expect(funcStart).not.toBe(-1);
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < memberCode.length; i++) {
      if (memberCode[i] === '{') { depth++; started = true; }
      if (memberCode[i] === '}') depth--;
      if (started) block += memberCode[i];
      if (started && depth === 0) break;
    }
    // Must check length before truncating (not blind substring)
    expect(block).toMatch(/description\.length\s*>/);
    // Must include ellipsis character or '...'
    expect(block).toMatch(/\\u2026|\.{3}|\u2026/);
  });

  test('getMemberTasks supports not-completed filter', () => {
    // The backend must handle 'not-completed' as a special status filter
    const funcMatch = dataServiceCode.match(/function getMemberTasks[\s\S]*?return tasks;\s*\}/);
    expect(funcMatch).not.toBeNull();
    expect(funcMatch[0]).toContain('not-completed');
  });

  test('renderBottomNav shows member task badge on More tab', () => {
    // The steward_view.html renderBottomNav should have a badge for member task count
    const funcStart = stewardCode.indexOf('function renderBottomNav');
    expect(funcStart).not.toBe(-1);
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < stewardCode.length; i++) {
      if (stewardCode[i] === '{') { depth++; started = true; }
      if (stewardCode[i] === '}') depth--;
      if (started) block += stewardCode[i];
      if (started && depth === 0) break;
    }
    expect(block).toContain('memberTaskCount');
  });

  test('_getMemberBatchData includes memberTaskCount', () => {
    const funcMatch = dataServiceCode.match(/function _getMemberBatchData[\s\S]*?return \{[\s\S]*?\};/);
    expect(funcMatch).not.toBeNull();
    expect(funcMatch[0]).toContain('memberTaskCount');
  });

  test('initMemberView seeds AppState.memberTaskCount from batch', () => {
    expect(memberCode).toContain('AppState.memberTaskCount');
  });

  test('My Tasks menu item in renderMemberMore has badge property', () => {
    // Verify the More menu My Tasks entry includes a badge
    expect(memberCode).toMatch(/label:\s*'My Tasks'[\s\S]*?badge:/);
  });
});


// ============================================================================
// G18: SHARED TABS ROUTE FOR BOTH ROLES
// ============================================================================
// Bug (2026-03-14): POMS was a shared tab in both steward and member More menus,
// but _handleTabNav only handled it for role === 'member'. Stewards were silently
// redirected to their dashboard. This guard ensures any tab referenced by BOTH
// views is routed before the role-specific switch blocks (i.e., as a shared tab).

describe('G18: Shared tabs route for both roles', () => {
  const indexCode = read('index.html');
  const stewardCode = read('steward_view.html');
  const memberCode = read('member_view.html');

  // Extract tab IDs from More menus
  function extractMenuTabIds(code) {
    const ids = new Set();
    const navRegex = /_handleTabNav\(\s*'[^']*'\s*,\s*'([^']+)'\s*\)/g;
    let m;
    while ((m = navRegex.exec(code)) !== null) ids.add(m[1]);
    return ids;
  }

  const stewardMenuTabs = extractMenuTabIds(stewardCode);
  const memberMenuTabs = extractMenuTabIds(memberCode);

  // Tabs that appear in BOTH views
  const sharedTabs = [...stewardMenuTabs].filter(id => memberMenuTabs.has(id));

  // Extract shared tab handlers (if-checks before the role switch blocks)
  function extractSharedHandlers(code) {
    // Shared handlers are if (tabId === 'xxx') blocks before "if (role === 'member')"
    const roleSwitch = code.indexOf("if (role === 'member')");
    if (roleSwitch === -1) return new Set();
    const beforeSwitch = code.substring(
      code.indexOf('function _handleTabNav'),
      roleSwitch
    );
    const ids = new Set();
    const pattern = /if\s*\(tabId\s*===\s*'([^']+)'\)/g;
    let m;
    while ((m = pattern.exec(beforeSwitch)) !== null) ids.add(m[1]);
    return ids;
  }

  const sharedHandlers = extractSharedHandlers(indexCode);

  test('found shared tabs between steward and member views', () => {
    expect(sharedTabs.length).toBeGreaterThan(0);
  });

  sharedTabs.forEach(tabId => {
    test(`shared tab '${tabId}' is handled before role-specific switch (not role-gated)`, () => {
      // Must be in shared handlers OR in BOTH role-specific switch blocks
      if (sharedHandlers.has(tabId)) return; // handled as shared — OK

      // Otherwise check both switches have it
      const memberSwitch = indexCode.match(/if\s*\(role\s*===\s*'member'\)\s*\{[\s\S]*?^\s{6}\}/m);
      const stewardSwitch = indexCode.match(/\}\s*else\s*\{[\s\S]*?case\s*'([^']+)'/gm);

      const memberCases = new Set();
      const stewardCases = new Set();
      // Member switch cases
      for (const m of indexCode.matchAll(/if\s*\(role\s*===\s*'member'\)\s*\{[\s\S]*?case\s*'([^']+)'/g)) {
        memberCases.add(m[1]);
      }
      // All case labels (steward + member)
      for (const m of indexCode.matchAll(/case\s*'([^']+)':/g)) {
        stewardCases.add(m[1]);
      }

      const inShared = sharedHandlers.has(tabId);
      const msg = `Tab '${tabId}' is in both steward and member menus but not handled as a shared tab. ` +
        'Add it before the role-specific switch in _handleTabNav to prevent role-gated routing.';
      expect(inShared).toBe(true); // enforce: shared tabs must be handled as shared
    });
  });
});


// ============================================================================
// G19: MORE MENU ITEMS MUST HAVE ROUTE HANDLERS
// ============================================================================
// Bug class: A tab ID in a More menu's _handleTabNav call has no corresponding
// handler in _handleTabNav, causing silent redirect to default (dashboard/home).
// This is a superset of G8 — checks More menus in the view files, not just sidebar.

describe('G19: More menu items have route handlers', () => {
  const indexCode = read('index.html');
  const stewardCode = read('steward_view.html');
  const memberCode = read('member_view.html');

  function extractMenuTabIds(code) {
    const ids = [];
    const navRegex = /_handleTabNav\(\s*'[^']*'\s*,\s*'([^']+)'\s*\)/g;
    let m;
    while ((m = navRegex.exec(code)) !== null) ids.push(m[1]);
    return ids;
  }

  function extractHandledIds(code) {
    const handled = new Set();
    for (const m of code.matchAll(/if\s*\(tabId\s*===\s*'([^']+)'\)/g)) handled.add(m[1]);
    for (const m of code.matchAll(/case\s*'([^']+)':/g)) handled.add(m[1]);
    handled.add('_member_more');
    handled.add('_steward_more');
    handled.add('more');
    handled.add('more_steward');
    return handled;
  }

  const handledIds = extractHandledIds(indexCode);

  test('all steward More menu tab IDs have route handlers', () => {
    const stewardMenuTabs = extractMenuTabIds(stewardCode);
    const unhandled = stewardMenuTabs.filter(id => !handledIds.has(id));
    if (unhandled.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Steward More menu tabs with no route handler:', unhandled);
    }
    expect(unhandled).toEqual([]);
  });

  test('all member More menu tab IDs have route handlers', () => {
    const memberMenuTabs = extractMenuTabIds(memberCode);
    const unhandled = memberMenuTabs.filter(id => !handledIds.has(id));
    if (unhandled.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Member More menu tabs with no route handler:', unhandled);
    }
    expect(unhandled).toEqual([]);
  });
});


// ============================================================================
// G20: POMS DESCRIPTION ACCURACY
// ============================================================================
// Bug (2026-03-14): Both views described POMS as "Postal Operations Manual"
// instead of "Program Operations Manual System". This guard prevents the wrong
// acronym expansion from reappearing.

describe('G20: POMS description accuracy', () => {
  const stewardCode = read('steward_view.html');
  const memberCode = read('member_view.html');

  test('steward view does not say "Postal Operations Manual"', () => {
    expect(stewardCode).not.toContain('Postal Operations Manual');
  });

  test('member view does not say "Postal Operations Manual"', () => {
    expect(memberCode).not.toContain('Postal Operations Manual');
  });

  test('steward view has correct POMS description', () => {
    expect(stewardCode).toContain('Program Operations Manual System');
  });

  test('member view has correct POMS description', () => {
    expect(memberCode).toContain('Program Operations Manual System');
  });
});


// ============================================================================
// G21: MEMBER DUES-GATED TABS HAVE _isDuesPaying() GUARD
// ============================================================================
// Bug history: v4.25.10 — Resources tab banner claimed dues restriction but
// renderMemberResources() had no _isDuesPaying() check, allowing non-paying
// members full access. This test ensures all tabs that should be dues-gated
// actually call _isDuesPaying() at entry.

describe('G21: Member dues-gated tabs all have _isDuesPaying() guard', () => {
  const memberCode = read('member_view.html');

  // Tabs that MUST be dues-gated — any member tab that exposes premium content.
  // If you add a new dues-gated tab, add its render function name here.
  const duesGatedTabs = [
    { fn: 'renderMemberResources', label: 'Resources' },
    { fn: 'renderSurveyFormPage', label: 'Quarterly Survey' },
    { fn: 'renderUnionStatsPage', label: 'Union Stats' },
    { fn: 'renderPollsPage', label: 'Polls' },
  ];

  // Tabs that are intentionally NOT dues-gated (public/core features)
  const publicTabs = [
    'renderMemberHome',
    'renderMyCases',
    'renderStewardContact',
    'renderUpdateProfile',
    'renderMemberNotifications',
  ];

  duesGatedTabs.forEach(({ fn, label }) => {
    test(`${fn}() (${label}) calls _isDuesPaying() before rendering`, () => {
      // Extract the function body
      const fnStart = memberCode.indexOf('function ' + fn + '(');
      expect(fnStart).toBeGreaterThan(-1);

      // Get the first ~200 chars of the function body (gate must be at entry)
      const bodyStart = memberCode.indexOf('{', fnStart);
      const earlyBody = memberCode.substring(bodyStart, bodyStart + 300);

      expect(earlyBody).toContain('_isDuesPaying()');
      expect(earlyBody).toContain('_renderDuesGate');
    });
  });

  publicTabs.forEach(fn => {
    test(`${fn}() is intentionally NOT dues-gated (exists for whitelist)`, () => {
      expect(memberCode).toContain('function ' + fn + '(');
    });
  });
});

// ============================================================================
// G22: Workload Tracker frontend invariants
// ============================================================================

describe('G22 — Workload Tracker frontend invariants', () => {
  const memberView = read('member_view.html');

  test('WT_CAT_KEY_LABELS map is defined from WT_CATEGORIES', () => {
    expect(memberView).toContain('var WT_CAT_KEY_LABELS = {}');
    expect(memberView).toContain('WT_CATEGORIES.forEach');
    expect(memberView).toContain('WT_CAT_KEY_LABELS[c.key] = c.label');
  });

  test('history sub-categories use WT_CAT_KEY_LABELS not raw keys', () => {
    // Must NOT use ck + ' > ' (raw key concatenation)
    expect(memberView).not.toMatch(/ck \+ ' > ' \+ sk/);
    // Must use label lookup
    expect(memberView).toContain('WT_CAT_KEY_LABELS[ck]');
  });

  test('stats sub-category headings use WT_CAT_KEY_LABELS', () => {
    expect(memberView).toContain('WT_CAT_KEY_LABELS[sck]');
  });

  test('bar chart uses dynamic max instead of hardcoded 100', () => {
    // Must NOT have (val / 100) * 100 pattern
    expect(memberView).not.toMatch(/val\s*\/\s*100\)\s*\*\s*100/);
    // Must have barMax calculation
    expect(memberView).toContain('barMax');
    expect(memberView).toMatch(/val\s*\/\s*barMax/);
  });

  test('auto-save draft functions are defined', () => {
    expect(memberView).toContain('function _saveDraft()');
    expect(memberView).toContain('function _loadDraft()');
    expect(memberView).toContain('function _clearDraft()');
  });

  test('draft is cleared on successful submission', () => {
    expect(memberView).toContain('_clearDraft()');
  });

  test('draft is restored on page load', () => {
    expect(memberView).toContain('_loadDraft()');
    expect(memberView).toContain('_wtRestoreForm(inputs, draft)');
  });

  test('last-submitted indicator is rendered', () => {
    expect(memberView).toContain('wt-last-submitted');
    expect(memberView).toContain('wt_lastSub_');
  });

  test('WT_CATEGORIES has exactly 8 entries with required fields', () => {
    const catMatch = memberView.match(/var WT_CATEGORIES = \[([\s\S]*?)\];/);
    expect(catMatch).not.toBeNull();
    const idMatches = catMatch[1].match(/id:\s*'t\d'/g);
    expect(idMatches).not.toBeNull();
    expect(idMatches.length).toBe(8);
    const keyMatches = catMatch[1].match(/key:\s*'[a-z]+'/g);
    expect(keyMatches).not.toBeNull();
    expect(keyMatches.length).toBe(8);
  });

  test('WT_CAT_KEYS has exactly 8 entries matching WT_CATEGORIES keys', () => {
    const keysMatch = memberView.match(/var WT_CAT_KEYS = \[([^\]]+)\]/);
    expect(keysMatch).not.toBeNull();
    const keys = keysMatch[1].match(/'([a-z]+)'/g).map(k => k.replace(/'/g, ''));
    expect(keys.length).toBe(8);
    // Each key must appear in WT_CATEGORIES
    keys.forEach(k => {
      expect(memberView).toContain("key: '" + k + "'");
    });
  });
});
