/**
 * Home Tab Integrity Guards — Prevents home tab bugs:
 *
 *   H1:  Quick Links must have onClick handlers (navigation targets)
 *   H2:  Steward _handleTabNav must have explicit 'home' case
 *   H3:  DataCache.cachedCall must pass SESSION_TOKEN, not CURRENT_USER.email, to server functions
 *   H4:  Member home must show an empty state when no active grievances
 *   H5:  All card-clickable elements must have onClick or addEventListener handlers
 *
 * Run: npx jest test/home-tab-integrity.test.js --verbose
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.resolve(__dirname, '..', 'src');

function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

// Helper: extract a function body by name from code
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
// H1: QUICK LINKS MUST HAVE CLICK HANDLERS
// ============================================================================
// Bug: Quick Links on member home had card-clickable styling but no onClick
// handlers — clicking them did nothing.

describe('H1: Quick Links have onClick navigation', () => {
  const memberCode = read('member_view.html');
  const homeBody = extractFunctionBody(memberCode, 'renderMemberHome');

  test('Quick Links section exists', () => {
    expect(homeBody).toContain('Quick Links');
  });

  test('Quick Links items have tab targets', () => {
    // Each link object must have a tab property
    expect(homeBody).toMatch(/label:\s*'Know Your Rights'.*tab:\s*'/s);
    expect(homeBody).toMatch(/label:\s*'File a Grievance'.*tab:\s*'/s);
    expect(homeBody).toMatch(/label:\s*'FAQ'.*tab:\s*'/s);
  });

  test('Quick Links forEach attaches onClick with _handleTabNav', () => {
    // The forEach loop for links must include an onClick that calls _handleTabNav
    expect(homeBody).toMatch(/onClick:\s*function\s*\(\)\s*\{\s*_handleTabNav\(/);
  });

  test('Quick Links card-clickable elements have cursor pointer', () => {
    // card-clickable items used in Quick Links must include cursor style
    expect(homeBody).toMatch(/card-clickable.*cursor.*pointer/s);
  });
});


// ============================================================================
// H2: STEWARD HOME TAB ROUTING
// ============================================================================
// Bug: Steward _handleTabNav had no case 'home', so navigating to 'home'
// fell through to default. While the default rendered the correct page,
// the sidebar wouldn't highlight any tab since 'home' didn't match 'cases'.

describe('H2: Steward _handleTabNav has explicit home case', () => {
  const indexCode = read('index.html');

  test('steward switch block has case home', () => {
    // Tab routing may be inline in _handleTabNav or extracted to _getTabRenderFn.
    // Either way, there must be a steward switch block with case 'home'.
    const hasInline = indexCode.match(/\}\s*else\s*\{[\s\S]*?switch\s*\(tabId\)\s*\{[^}]*case\s*'home'/);
    const hasExtracted = indexCode.match(/function _getTabRenderFn[\s\S]*?\}\s*else\s*\{[\s\S]*?case\s*'home'/);
    expect(hasInline || hasExtracted).not.toBeNull();
  });

  test('steward home case renders dashboard', () => {
    // case 'home' should map to renderStewardDashboard — either via direct call or return
    const hasInline = indexCode.match(/case\s*'home':\s*renderStewardDashboard/);
    const hasReturn = indexCode.match(/case\s*'home'[\s\S]*?case\s*'cases':\s*return\s+renderStewardDashboard/);
    expect(hasInline || hasReturn).not.toBeNull();
  });
});


// ============================================================================
// H3: DATACACHE.CACHEDCALL MUST PASS SESSION_TOKEN, NOT EMAIL
// ============================================================================
// Bug: DataCache.cachedCall for member data functions passed CURRENT_USER.email
// as the function argument instead of SESSION_TOKEN. The server functions
// (dataGetSurveyStatus, dataGetMemberGrievanceHistory, dataGetAssignedSteward)
// expect sessionToken as the first parameter for _resolveCallerEmail().
// Passing email instead of token would fail in non-SSO (magic link) deployments.

describe('H3: DataCache.cachedCall passes SESSION_TOKEN to server functions', () => {
  const memberCode = read('member_view.html');

  // Server functions that require sessionToken (not email) as first arg
  const tokenFunctions = [
    'dataGetSurveyStatus',
    'dataGetMemberGrievanceHistory',
    'dataGetAssignedSteward',
    'dataGetUpcomingEvents',
  ];

  test.each(tokenFunctions)('%s is called with SESSION_TOKEN', (fnName) => {
    // Find all DataCache.cachedCall invocations for this function
    const pattern = new RegExp(
      `DataCache\\.cachedCall\\([^)]*,\\s*'${fnName}'\\s*,\\s*\\[([^\\]]+)\\]`,
      'g'
    );
    let match;
    const argsList = [];
    while ((match = pattern.exec(memberCode)) !== null) {
      argsList.push(match[1].trim());
    }
    expect(argsList.length).toBeGreaterThan(0);
    for (const args of argsList) {
      // First argument should be SESSION_TOKEN, not CURRENT_USER.email
      expect(args).toMatch(/^SESSION_TOKEN/);
      expect(args).not.toMatch(/^CURRENT_USER\.email/);
    }
  });

  test('no DataCache.cachedCall passes CURRENT_USER.email as server function arg', () => {
    // Generic guard: no cachedCall should pass CURRENT_USER.email in the args array
    // (cache key using CURRENT_USER.email is fine — it's the args array that matters)
    const pattern = /DataCache\.cachedCall\([^,]+,\s*'[^']+'\s*,\s*\[CURRENT_USER\.email\]/g;
    const matches = memberCode.match(pattern) || [];
    if (matches.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('DataCache.cachedCall passing CURRENT_USER.email as server arg:', matches);
    }
    expect(matches).toEqual([]);
  });
});


// ============================================================================
// H4: MEMBER HOME EMPTY STATE FOR GRIEVANCES
// ============================================================================
// Bug: When no active grievances existed, the member home showed nothing in
// that section — no feedback to the user at all.

describe('H4: Member home shows empty state for no active grievances', () => {
  const memberCode = read('member_view.html');
  const homeBody = extractFunctionBody(memberCode, 'renderMemberHome');

  test('has else branch when no active grievance', () => {
    // The active grievance section must have an else branch
    expect(homeBody).toMatch(/if\s*\(activeG\)\s*\{[\s\S]*?\}\s*else\s*\{/);
  });

  test('else branch shows "No active cases" message', () => {
    expect(homeBody).toContain('No active cases');
  });
});


// ============================================================================
// H5: QUICK LINKS IN renderMemberHome MUST HAVE CLICK HANDLERS
// ============================================================================
// Targeted guard: the Quick Links section in renderMemberHome must wire
// onClick handlers to each link item. This prevents reintroduction of the
// original bug where card-clickable cards had no navigation.

describe('H5: Quick Links in renderMemberHome have click navigation', () => {
  const memberCode = read('member_view.html');
  const homeBody = extractFunctionBody(memberCode, 'renderMemberHome');

  test('every Quick Link item has a tab property', () => {
    // Extract the links array from the Quick Links section
    const linksMatch = homeBody.match(/var links\s*=\s*\[([\s\S]*?)\];/);
    expect(linksMatch).not.toBeNull();
    const linksBlock = linksMatch[1];
    // Count link objects (each has a label)
    const labelCount = (linksBlock.match(/label:/g) || []).length;
    const tabCount = (linksBlock.match(/tab:/g) || []).length;
    expect(tabCount).toBe(labelCount);
    expect(tabCount).toBeGreaterThan(0);
  });

  test('links forEach has onClick calling _handleTabNav with r.tab', () => {
    // The forEach that renders links must reference r.tab in the onClick
    expect(homeBody).toMatch(/_handleTabNav\(\s*'member'\s*,\s*r\.tab\s*\)/);
  });
});


// ============================================================================
// H6: SERVER DATA FUNCTIONS REQUIRE AUTH
// ============================================================================
// Guard: All data* functions in WebDashDataService must call _resolveCallerEmail
// or _requireStewardAuth. This is already in CLAUDE.md but adding a test.

describe('H6: Home tab server functions have auth checks', () => {
  const dataServiceCode = read('21_WebDashDataService.gs');

  const homeFunctions = [
    'dataGetSurveyStatus',
    'dataGetMemberGrievanceHistory',
    'dataGetAssignedSteward',
    'dataGetUpcomingEvents',
    'dataGetWelcomeData',
  ];

  test.each(homeFunctions)('%s has auth check', (fnName) => {
    const body = extractFunctionBody(dataServiceCode, fnName);
    expect(body.length).toBeGreaterThan(0);
    const hasAuth = body.includes('_resolveCallerEmail') || body.includes('_requireStewardAuth');
    expect(hasAuth).toBe(true);
  });
});
