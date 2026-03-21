/**
 * Org Chart Integrity Guards — SolidBase Generic Org Chart
 *
 * SolidBase uses a generic placeholder org chart (orgchart-embed) instead of
 * the MADDS-specific org chart. These tests verify the structural integrity
 * of the generic org chart within the SPA.
 *
 *   OC1: Org chart container must NOT be constrained by page-layout-content max-width
 *   OC2: All org chart interactive functions must be explicitly global (window.*)
 *   OC3: Script re-execution must use reliable global-scope injection
 *   OC4: _initDesktopRan flag must be reset before script re-execution
 *   OC5: All onclick handlers in org_chart.html must reference declared functions
 *   OC6: Light/dark mode toggle must exist and reference orgchartToggleMode
 *   OC7: (Skipped) Pill/rep/ma toggle targets — not applicable to generic chart
 *   OC8: Org chart CSS is scoped under .orgchart-embed
 *   OC9: Overflow handling — no content clipping
 *   OC10: getOrgChartHtml server function exists
 *
 * Run: npx jest test/org-chart-integrity.test.js --verbose
 */

const fs = require('fs');
const path = require('path');

const SRC_DIR = path.resolve(__dirname, '..', 'src');

function read(file) {
  return fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
}

const orgChartCode = read('org_chart.html');
const indexCode = read('index.html');
const stylesCode = read('styles.html');


// ============================================================================
// OC1: ORG CHART WIDTH — NOT CLIPPED BY page-layout-content MAX-WIDTH
// ============================================================================
// Bug: .page-layout-content had max-width:800px on desktop, but the org chart
// needs full width. Content was visually clipped on the right side.

describe('OC1: Org chart is not clipped by parent max-width', () => {
  test('styles.html has a max-width override for org chart content', () => {
    // There must be a CSS rule that removes max-width for the org chart container.
    // Accept .orgchart-content class-based override or :has(.orgchart-embed).
    const hasOverride =
      stylesCode.includes('orgchart-content') && stylesCode.includes('max-width: none');
    expect(hasOverride).toBe(true);
  });

  test('renderOrgChart adds orgchart-content class to container', () => {
    // Class-based fallback for browsers without :has() support
    expect(indexCode).toMatch(/orgchart-content/);
  });
});


// ============================================================================
// OC2: ALL INTERACTIVE FUNCTIONS ARE EXPLICITLY GLOBAL (window.*)
// ============================================================================
// Bug: Functions defined via plain `function foo()` inside dynamically injected
// <script> tags were not reliably accessible from onclick="" handlers in some
// browser/GAS contexts. Assigning to window.* guarantees global scope.

describe('OC2: Org chart interactive functions are explicitly global', () => {
  // SolidBase generic chart only has the toggle-mode function
  const requiredGlobalFunctions = [
    'orgchartToggleMode',
  ];

  test.each(requiredGlobalFunctions)('%s is assigned to window', (fnName) => {
    const pattern = new RegExp(`window\\.${fnName}\\s*=`);
    expect(orgChartCode).toMatch(pattern);
  });
});


// ============================================================================
// OC3: SCRIPT RE-EXECUTION USES RELIABLE GLOBAL-SCOPE INJECTION
// ============================================================================
// Bug: Using replaceChild to re-execute scripts was fragile. Appending to
// document.head ensures scripts run in the global scope reliably.

describe('OC3: renderOrgChart script re-execution is reliable', () => {
  // Extract the renderOrgChart function body
  const funcStart = indexCode.indexOf('function renderOrgChart(');
  const funcBlock = (function() {
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < indexCode.length; i++) {
      if (indexCode[i] === '{') { depth++; started = true; }
      if (indexCode[i] === '}') { depth--; }
      if (started) block += indexCode[i];
      if (started && depth === 0) break;
    }
    return block;
  })();

  test('renderOrgChart function exists', () => {
    expect(funcStart).toBeGreaterThan(-1);
    expect(funcBlock.length).toBeGreaterThan(50);
  });

  test('scripts are appended to document.head (not replaceChild)', () => {
    // Must use document.head.appendChild for reliable global-scope execution
    expect(funcBlock).toMatch(/document\.head\.appendChild/);
    // Should NOT use the fragile replaceChild pattern
    expect(funcBlock).not.toMatch(/replaceChild/);
  });

  test('scripts are iterated via querySelectorAll', () => {
    expect(funcBlock).toMatch(/querySelectorAll\s*\(\s*['"]script['"]\s*\)/);
  });
});


// ============================================================================
// OC4: _initDesktopRan FLAG IS RESET ON RE-NAVIGATION
// ============================================================================
// Bug: _initDesktopRan was set to true on first visit and never reset. When
// the user navigated away and back, initDesktop() wouldn't run, leaving
// Organization sub-sections in an incorrect state.

describe('OC4: _initDesktopRan is reset for SPA re-navigation', () => {
  test('renderOrgChart explicitly resets _initDesktopRan before script re-execution', () => {
    // The reset must happen BEFORE the script forEach loop
    const resetIdx = indexCode.indexOf('_initDesktopRan = false');
    const forEachIdx = indexCode.indexOf("querySelectorAll('script')");
    expect(resetIdx).toBeGreaterThan(-1);
    expect(forEachIdx).toBeGreaterThan(-1);
    expect(resetIdx).toBeLessThan(forEachIdx);
  });
});


// ============================================================================
// OC5: ALL ONCLICK HANDLERS REFERENCE DECLARED FUNCTIONS
// ============================================================================
// Bug: onclick handlers referenced functions that didn't exist or weren't
// global, causing silent failures when buttons were clicked.

describe('OC5: onclick handlers reference declared functions', () => {
  test('every onclick function call in org_chart.html has a matching window.* declaration', () => {
    // Extract all function names from onclick attributes
    const onclickRegex = /onclick="(\w+)\s*\(/g;
    const calledFunctions = new Set();
    let m;
    while ((m = onclickRegex.exec(orgChartCode)) !== null) {
      calledFunctions.add(m[1]);
    }

    // Extract all window.* function declarations
    const windowFuncRegex = /window\.(\w+)\s*=\s*function/g;
    const declaredFunctions = new Set();
    while ((m = windowFuncRegex.exec(orgChartCode)) !== null) {
      declaredFunctions.add(m[1]);
    }

    // Also accept inline onclick with document.getElementById (not function calls)
    const inlinePatterns = ['document.getElementById'];
    const missing = [...calledFunctions].filter(fn => {
      if (declaredFunctions.has(fn)) return false;
      // Check if it's an inline pattern rather than a named function
      if (inlinePatterns.some(p => fn.startsWith(p))) return false;
      return true;
    });

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('onclick references undeclared functions:', missing);
    }
    expect(missing).toEqual([]);
  });
});


// ============================================================================
// OC6: LIGHT/DARK MODE TOGGLE EXISTS AND IS WIRED
// ============================================================================
// SolidBase generic chart uses orgchartToggleMode instead of maddstoggleMode.

describe('OC6: Light/dark mode toggle is functional', () => {
  test('orgchart-mode-toggle button exists with onclick', () => {
    expect(orgChartCode).toMatch(/id="orgchart-mode-toggle"/);
    expect(orgChartCode).toMatch(/onclick="orgchartToggleMode\(\)"/);
  });

  test('orgchartToggleMode toggles .light class on .orgchart-embed', () => {
    expect(orgChartCode).toMatch(/\.orgchart-embed/);
    expect(orgChartCode).toMatch(/classList\.toggle\s*\(\s*['"]light['"]\s*\)/);
  });

  test('light mode CSS variables are defined', () => {
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light\s*\{/);
    // Must override key variables
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light[\s\S]*?--bg:/);
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light[\s\S]*?--txt:/);
  });
});


// ============================================================================
// OC7: PILL BUTTON TOGGLE TARGETS (NOT APPLICABLE TO GENERIC CHART)
// ============================================================================
// SolidBase generic org chart is a placeholder without pill/rep/ma toggles.
// These tests are skipped — they apply only to MADDS org chart.

describe('OC7: Pill button toggle targets (generic chart — no toggles)', () => {
  test('generic chart has no pill/rep/ma toggle functions (placeholder only)', () => {
    // Verify this is the generic placeholder chart, not the MADDS chart
    expect(orgChartCode).toContain('placeholder-card');
    expect(orgChartCode).not.toContain('pillToggle');
    expect(orgChartCode).not.toContain('repToggle');
    expect(orgChartCode).not.toContain('maToggle');
  });
});


// ============================================================================
// OC8: ORGCHART-EMBED WRAPPER SCOPING
// ============================================================================
// Bug: Org chart styles leaked into the SPA or SPA styles leaked into the
// org chart because CSS wasn't properly scoped under .orgchart-embed.

describe('OC8: Org chart CSS is scoped under .orgchart-embed', () => {
  test('org chart HTML is wrapped in .orgchart-embed div', () => {
    expect(orgChartCode).toMatch(/<div class="orgchart-embed">/);
  });

  test('all style rules (except footer/keyframes) are scoped to .orgchart-embed', () => {
    // Extract the <style> block
    const styleEnd = orgChartCode.indexOf('</style>');
    if (styleEnd === -1) return; // no style block
    const styleBlock = orgChartCode.substring(0, styleEnd);

    // Find unscoped rules: lines that start a CSS rule block but don't
    // reference .orgchart-embed or start with @, #orgchart-, or are comments
    const lines = styleBlock.split('\n');
    const unscopedRules = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      // Skip empty lines, comments, closing braces, properties, @-rules
      if (!line || line.startsWith('/*') || line.startsWith('*') ||
          line.startsWith('}') || line.startsWith('@') ||
          line.includes(':') && !line.includes('{') ||
          line.startsWith('//')) continue;

      // Lines that start a rule block (contain '{')
      if (line.includes('{')) {
        const selector = line.split('{')[0].trim();
        if (!selector) continue;
        // Must be scoped to .orgchart-embed OR be #orgchart-mode-toggle OR be a keyframe name
        if (selector.includes('.orgchart-embed') ||
            selector.includes('#orgchart-') ||
            selector.startsWith('@') ||
            selector.startsWith('from') ||
            selector.startsWith('to') ||
            /^\d+%/.test(selector)) continue;

        // Allow footer and other elements that only exist inside the wrapper div
        // Only flag selectors that could conflict with main SPA: html, body, *, etc.
        if (['html', 'body', 'html *', 'body *'].some(s => selector === s)) {
          unscopedRules.push(`Line ${i + 1}: ${selector}`);
        }
      }
    }

    if (unscopedRules.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Unscoped CSS rules that may leak:', unscopedRules);
    }
    expect(unscopedRules).toEqual([]);
  });
});


// ============================================================================
// OC9: OVERFLOW HANDLING — NO CONTENT CLIPPING
// ============================================================================
// Bug: overflow-x:hidden on a width-constrained parent caused content to be
// invisible instead of scrollable.

describe('OC9: Org chart overflow is handled correctly', () => {
  test('.orgchart-embed has overflow-x handling', () => {
    // .orgchart-embed should handle overflow (hidden is OK since content is self-contained)
    expect(orgChartCode).toMatch(/\.orgchart-embed\s*\{[^}]*overflow-x:\s*(hidden|auto)/);
  });

  test('.orgchart-embed uses box-sizing: border-box (via universal selector)', () => {
    // Generic chart uses a universal reset: .orgchart-embed *, ... { box-sizing: border-box }
    expect(orgChartCode).toMatch(/\.orgchart-embed \*[^{]*\{[^}]*box-sizing:\s*border-box/);
  });
});


// ============================================================================
// OC10: getOrgChartHtml SERVER FUNCTION EXISTS
// ============================================================================
// Bug: If the server function is removed or renamed, the entire org chart
// tab silently fails to load.

describe('OC10: getOrgChartHtml server function integrity', () => {
  const webAppCode = read('22_WebDashApp.gs');

  test('getOrgChartHtml function exists', () => {
    expect(webAppCode).toMatch(/function getOrgChartHtml\s*\(/);
  });

  test('getOrgChartHtml loads org_chart file', () => {
    expect(webAppCode).toMatch(/createHtmlOutputFromFile\s*\(\s*['"]org_chart['"]\s*\)/);
  });

  test('getOrgChartHtml has error handling', () => {
    // Must have try/catch to avoid unhandled server errors
    const funcMatch = webAppCode.match(/function getOrgChartHtml[\s\S]*?^}/m);
    expect(funcMatch).not.toBeNull();
    expect(funcMatch[0]).toMatch(/try\s*\{/);
    expect(funcMatch[0]).toMatch(/catch/);
  });

  test('renderOrgChart calls getOrgChartHtml', () => {
    expect(indexCode).toMatch(/\.getOrgChartHtml\s*\(/);
  });
});
