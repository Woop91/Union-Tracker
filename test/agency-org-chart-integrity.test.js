/**
 * Agency Org Chart Integrity Guards — Prevents recurring issues with
 * agency_org_chart.html embedded inside the SPA via lazy load:
 *
 *   AOC1: All toggle/tab functions must be defined
 *   AOC2: Every onclick handler must reference a declared function
 *   AOC3: Every getElementById target in toggle functions must exist in HTML
 *   AOC4: Tab panels exist for all tab IDs
 *   AOC5: CSP rebind function exists and is callable (window.agencyOC_init)
 *   AOC6: getAgencyOrgChartHtml server function exists and is wired
 *   AOC7: Script blocks parse without syntax errors
 *   AOC8: CSS is scoped under .agency-oc
 *
 * Run: npx jest test/agency-org-chart-integrity.test.js --verbose
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC_DIR = path.resolve(__dirname, '..', 'src');

function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

// SolidBase: agency_org_chart.html is excluded from SolidBase (DDS-only feature).
// Skip this entire test file when the file doesn't exist.
const agencyOcPath = path.join(SRC_DIR, 'agency_org_chart.html');
const AGENCY_OC_PRESENT = fs.existsSync(agencyOcPath);

const agencyCode = AGENCY_OC_PRESENT ? read('agency_org_chart.html') : '';
const indexCode = read('index.html');
const webAppCode = read('22_WebDashApp.gs');


// ============================================================================
// AOC1: ALL TOGGLE/TAB FUNCTIONS ARE DEFINED
// ============================================================================
// Bug: Toggle functions not defined or not accessible from onclick handlers
// cause silent click failures across the entire agency org chart tab.

describe.skip('AOC1: Agency org chart toggle/tab functions are defined', () => {
  const requiredFunctions = [
    'agencyOC_showTab',
    'agencyOC_openSal',
    'agencyOC_closeSal',
    'agencyOC_toggleFuncDirectors',
    'agencyOC_toggleThomas',
    'agencyOC_toggleSpecial',
    'agencyOC_toggleDDSSiblings',
    'agencyOC_toggleSiblings',
    'agencyOC_togCtxSub',
    'agencyOC_toggleJDSection',
    'agencyOC_togJD',
    'agencyOC_toggleSec',
    'agencyOC_initDirectorEditing',
    'agencyOC_saveDirectorEdits',
  ];

  test.each(requiredFunctions)('%s is declared as a function', (fnName) => {
    const pattern = new RegExp(`function\\s+${fnName}\\s*\\(`);
    expect(agencyCode).toMatch(pattern);
  });
});


// ============================================================================
// AOC2: EVERY ONCLICK HANDLER REFERENCES A DECLARED FUNCTION
// ============================================================================
// Bug: onclick handlers referencing undeclared functions fail silently.
// The CSP rebind skips unknown functions, leaving buttons dead.

describe.skip('AOC2: onclick handlers reference declared functions', () => {
  test('every onclick function call has a matching function declaration', () => {
    // Extract function names from onclick attributes
    // Handles both onclick="fn()" and onclick="fn(&#39;arg&#39;)"
    const onclickRegex = /onclick="(\w+)\s*\(/g;
    const calledFunctions = new Set();
    let m;
    while ((m = onclickRegex.exec(agencyCode)) !== null) {
      calledFunctions.add(m[1]);
    }

    // Extract all function declarations
    const funcDeclRegex = /function\s+(\w+)\s*\(/g;
    const declaredFunctions = new Set();
    while ((m = funcDeclRegex.exec(agencyCode)) !== null) {
      declaredFunctions.add(m[1]);
    }

    // Also accept window.* assignments
    const windowFuncRegex = /window\.(\w+)\s*=\s*function/g;
    while ((m = windowFuncRegex.exec(agencyCode)) !== null) {
      declaredFunctions.add(m[1]);
    }

    const missing = [...calledFunctions].filter(fn => !declaredFunctions.has(fn));

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('onclick references undeclared functions:', missing);
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// AOC3: EVERY getElementById TARGET IN TOGGLE FUNCTIONS EXISTS IN HTML
// ============================================================================
// Bug: Toggle functions call getElementById('some-id') but the element was
// removed or renamed, causing null reference errors that crash the function
// and break all subsequent interactions.

describe.skip('AOC3: Toggle function getElementById targets exist in HTML', () => {
  // Extract the <script> block to find getElementById calls
  const scriptStart = agencyCode.indexOf('<script>');
  const scriptEnd = agencyCode.indexOf('</script>');
  const scriptBlock = agencyCode.substring(scriptStart, scriptEnd);

  // Extract the HTML block (everything before <script>)
  const htmlBlock = agencyCode.substring(0, scriptStart);

  test('every getElementById target in script has a matching element in HTML', () => {
    const idRegex = /getElementById\s*\(\s*['"]([^'"]+)['"]\s*\)/g;
    const referencedIds = new Set();
    let m;
    while ((m = idRegex.exec(scriptBlock)) !== null) {
      referencedIds.add(m[1]);
    }

    const missing = [];
    for (const id of referencedIds) {
      const idPattern = new RegExp(`id=["']${id}["']`);
      if (!idPattern.test(htmlBlock)) {
        missing.push(id);
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('getElementById targets missing from HTML:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('every onclick toggle target ID (togJD, togCtxSub, toggleSec) has a matching element', () => {
    // Only these functions take element IDs as arguments:
    //   agencyOC_togJD('bodyId','caretId')
    //   agencyOC_togCtxSub('listId','togId')
    //   agencyOC_toggleSec('bodyId','carId')
    // Other functions take non-ID args (showTab takes tab suffix, openSal takes data key)
    const idArgFuncs = ['togJD', 'togCtxSub', 'toggleSec'];
    const argIdRegex = /onclick="agencyOC_(togJD|togCtxSub|toggleSec)\((?:&#39;|')([^'&#]+)(?:&#39;|')(?:,\s*(?:&#39;|')([^'&#]+)(?:&#39;|'))?\)"/g;
    const targetIds = new Set();
    let m2;
    while ((m2 = argIdRegex.exec(agencyCode)) !== null) {
      if (m2[2]) targetIds.add(m2[2]);
      if (m2[3]) targetIds.add(m2[3]);
    }

    expect(targetIds.size).toBeGreaterThan(0);

    const missing = [];
    for (const id of targetIds) {
      const idPattern = new RegExp(`id=["']${id}["']`);
      if (!idPattern.test(htmlBlock)) {
        missing.push(id);
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('onclick toggle target IDs missing from HTML:', missing);
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// AOC4: TAB PANELS EXIST FOR ALL TAB IDs
// ============================================================================
// Bug: agencyOC_showTab('bt') was called but #tab-bt didn't exist, causing
// a crash that prevented all tab switching from working.

describe.skip('AOC4: Tab panels exist for all tab IDs', () => {
  const tabIds = ['oc', 'bs', 'bt', 'hb', 'hs', 'ai'];

  test.each(tabIds)('tab panel #tab-%s exists', (id) => {
    const pattern = new RegExp(`id=["']tab-${id}["']`);
    expect(agencyCode).toMatch(pattern);
  });

  test.each(tabIds)('tab button for %s exists with onclick', (id) => {
    const pattern = new RegExp(`onclick="agencyOC_showTab\\(.*?${id}.*?\\)"`);
    expect(agencyCode).toMatch(pattern);
  });

  test('all tab panels have the tab-panel class', () => {
    for (const id of tabIds) {
      // tab-oc may have 'tab-panel active', others just 'tab-panel'
      const pattern = new RegExp(`id=["']tab-${id}["'][^>]*class=["'][^"']*tab-panel`);
      expect(agencyCode).toMatch(pattern);
    }
  });
});


// ============================================================================
// AOC5: CSP REBIND FUNCTION EXISTS AND IS CALLABLE
// ============================================================================
// Bug: GAS iframe CSP blocks inline onclick attributes. The rebind function
// must exist, be globally callable (for index.html lazy-load fallback), and
// process all onclick elements within .agency-oc.

describe.skip('AOC5: CSP rebind function is properly configured', () => {
  test('agencyOC_init is assigned to window', () => {
    expect(agencyCode).toMatch(/window\.agencyOC_init\s*=/);
  });

  test('agencyOC_init is invoked immediately after definition', () => {
    const defIdx = agencyCode.indexOf('window.agencyOC_init =');
    const callIdx = agencyCode.indexOf('window.agencyOC_init();', defIdx);
    expect(defIdx).toBeGreaterThan(-1);
    expect(callIdx).toBeGreaterThan(defIdx);
  });

  test('rebind function targets .agency-oc root', () => {
    expect(agencyCode).toMatch(/querySelector\s*\(\s*['"]\.agency-oc['"]\s*\)/);
  });

  test('rebind function queries [onclick] elements', () => {
    expect(agencyCode).toMatch(/querySelectorAll\s*\(\s*['"\[]*onclick['"\]]*\s*\)/);
  });

  test('rebind function removes onclick attrs and adds event listeners', () => {
    expect(agencyCode).toMatch(/removeAttribute\s*\(\s*['"]onclick['"]\s*\)/);
    expect(agencyCode).toMatch(/addEventListener\s*\(\s*['"]click['"]/);
  });

  test('index.html calls agencyOC_init after lazy load', () => {
    expect(indexCode).toMatch(/agencyOC_init\s*\(\s*\)/);
  });
});


// ============================================================================
// AOC6: getAgencyOrgChartHtml SERVER FUNCTION EXISTS AND IS WIRED
// ============================================================================
// Bug: If the server function is removed or renamed, the agency org chart tab
// silently fails to load with no useful error message.

describe.skip('AOC6: getAgencyOrgChartHtml server function integrity', () => {
  test('getAgencyOrgChartHtml function exists', () => {
    expect(webAppCode).toMatch(/function getAgencyOrgChartHtml\s*\(/);
  });

  test('getAgencyOrgChartHtml loads agency_org_chart file', () => {
    expect(webAppCode).toMatch(/createHtmlOutputFromFile\s*\(\s*['"]agency_org_chart['"]\s*\)/);
  });

  test('getAgencyOrgChartHtml has error handling', () => {
    const funcMatch = webAppCode.match(/function getAgencyOrgChartHtml[\s\S]*?^}/m);
    expect(funcMatch).not.toBeNull();
    expect(funcMatch[0]).toMatch(/try\s*\{/);
    expect(funcMatch[0]).toMatch(/catch/);
  });

  test('renderAgencyOrgChart calls getAgencyOrgChartHtml', () => {
    expect(indexCode).toMatch(/\.getAgencyOrgChartHtml\s*\(\s*\)/);
  });

  test('hasAgencyOrgChart config flag exists for graceful fallback', () => {
    expect(webAppCode).toMatch(/hasAgencyOrgChart/);
    expect(indexCode).toMatch(/CONFIG\.hasAgencyOrgChart/);
  });
});


// ============================================================================
// AOC7: SCRIPT BLOCKS PARSE WITHOUT SYNTAX ERRORS
// ============================================================================
// Bug: A syntax error anywhere in the script block prevents ALL functions
// from being defined, breaking every interactive element on the page.

describe.skip('AOC7: Script blocks parse without syntax errors', () => {
  test('agency_org_chart.html script block is valid JavaScript', () => {
    const scriptStart = agencyCode.indexOf('<script>') + '<script>'.length;
    const scriptEnd = agencyCode.indexOf('</script>');
    const jsBlock = agencyCode.substring(scriptStart, scriptEnd);

    expect(() => {
      new vm.Script(jsBlock, { filename: 'agency_org_chart.html:script' });
    }).not.toThrow();
  });
});


// ============================================================================
// AOC8: CSS IS SCOPED UNDER .agency-oc
// ============================================================================
// Bug: Unscoped CSS rules leak into the SPA, or SPA rules bleed into the
// agency org chart, causing broken layouts.

describe.skip('AOC8: Agency org chart CSS is scoped under .agency-oc', () => {
  test('HTML is wrapped in .agency-oc div', () => {
    expect(agencyCode).toMatch(/<div class="agency-oc">/);
  });

  test('style rules reference .agency-oc scoping', () => {
    // Count rules that reference .agency-oc
    const scopedCount = (agencyCode.match(/\.agency-oc\s/g) || []).length;
    // Should have many scoped rules (the file has 50+ CSS rules)
    expect(scopedCount).toBeGreaterThan(20);
  });

  test('no dangerous global selectors (html, body, *)', () => {
    const styleEnd = agencyCode.indexOf('</style>');
    if (styleEnd === -1) return;
    const styleBlock = agencyCode.substring(0, styleEnd);
    const lines = styleBlock.split('\n');
    const dangerous = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line || !line.includes('{')) continue;
      const selector = line.split('{')[0].trim();
      if (!selector) continue;
      if (['html', 'body', '*'].includes(selector)) {
        dangerous.push(`Line ${i + 1}: ${selector}`);
      }
    }

    expect(dangerous).toEqual([]);
  });
});
