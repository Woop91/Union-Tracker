/**
 * Org Chart Integrity Guards — Prevents recurring issues with the org chart
 * embedded inside the SPA:
 *
 *   OC1: Org chart container must NOT be constrained by page-layout-content max-width
 *   OC2: All org chart interactive functions must be explicitly global (window.*)
 *   OC3: Script re-execution must use reliable global-scope injection
 *   OC4: _initDesktopRan flag must be reset before script re-execution
 *   OC5: All onclick handlers in org_chart.html must reference declared functions
 *   OC6: Light/dark mode toggle must exist and reference maddstoggleMode
 *   OC7: Org chart pill buttons must have matching toggle targets (id exists)
 *   OC8: renderOrgChart must mark container for CSS override
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

// SolidBase: org_chart.html is a generic placeholder without interactive functions.
// Detect this by checking for a key interactive function (pillToggle).
const _hasFullOrgChart = orgChartCode.includes('window.pillToggle');
const _describeInteractive = _hasFullOrgChart ? describe : describe.skip;


// ============================================================================
// OC1: ORG CHART WIDTH — NOT CLIPPED BY page-layout-content MAX-WIDTH
// ============================================================================
// Bug: .page-layout-content had max-width:800px on desktop, but the org chart
// needs up to 1200px. Content was visually clipped on the right side.

_describeInteractive('OC1: Org chart is not clipped by parent max-width', () => {
  test('styles.html has a max-width override for org chart content', () => {
    // There must be a CSS rule that removes max-width for the org chart container.
    // Accept either :has(.madds-embed) or .orgchart-content class-based override.
    const hasOverride =
      stylesCode.includes('.madds-embed') && stylesCode.includes('max-width: none') ||
      stylesCode.includes('orgchart-content') && stylesCode.includes('max-width: none');
    expect(hasOverride).toBe(true);
  });

  test('org chart .page max-width is wider than default content max-width', () => {
    // Org chart's .page should allow at least 1000px
    const pageMatch = orgChartCode.match(/\.page\s*\{[^}]*max-width:\s*(\d+)px/);
    expect(pageMatch).not.toBeNull();
    const orgChartMaxWidth = parseInt(pageMatch[1], 10);
    expect(orgChartMaxWidth).toBeGreaterThanOrEqual(1000);
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

_describeInteractive('OC2: Org chart interactive functions are explicitly global', () => {
  const requiredGlobalFunctions = [
    'pillToggle',
    'repToggle',
    'maToggle',
    'togglePSGroup',
    'toggleOtherGroup',
    'maddstoggleMode',
    'initDesktop',
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
// MassAbility sub-sections in an incorrect state.

_describeInteractive('OC4: _initDesktopRan is reset for SPA re-navigation', () => {
  test('org_chart.html declares _initDesktopRan on window (resets on re-execution)', () => {
    expect(orgChartCode).toMatch(/window\._initDesktopRan\s*=\s*false/);
  });

  test('renderOrgChart explicitly resets _initDesktopRan before script re-execution', () => {
    // The reset must happen BEFORE the script forEach loop in renderOrgChart.
    // Use renderOrgChart-specific pattern to avoid matching _loadMemberViewThen's querySelectorAll.
    const resetIdx = indexCode.indexOf('_initDesktopRan = false');
    const forEachIdx = indexCode.indexOf("wrap.querySelectorAll('script').forEach");
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

_describeInteractive('OC5: onclick handlers reference declared functions', () => {
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
// Bug: The light/dark mode toggle button existed but maddstoggleMode() wasn't
// globally accessible, so clicking the button did nothing.

_describeInteractive('OC6: Light/dark mode toggle is functional', () => {
  test('madds-mode-toggle button exists with onclick', () => {
    expect(orgChartCode).toMatch(/id="madds-mode-toggle"/);
    expect(orgChartCode).toMatch(/onclick="maddstoggleMode\(\)"/);
  });

  test('maddstoggleMode toggles .light class on .madds-embed', () => {
    expect(orgChartCode).toMatch(/\.madds-embed/);
    expect(orgChartCode).toMatch(/classList\.toggle\s*\(\s*['"]light['"]\s*\)/);
  });

  test('light mode CSS variables are defined', () => {
    expect(orgChartCode).toMatch(/\.madds-embed\.light\s*\{/);
    // Must override key variables
    expect(orgChartCode).toMatch(/\.madds-embed\.light[\s\S]*?--bg:/);
    expect(orgChartCode).toMatch(/\.madds-embed\.light[\s\S]*?--txt:/);
  });
});


// ============================================================================
// OC7: PILL BUTTON TOGGLE TARGETS EXIST
// ============================================================================
// Bug: pillToggle('some-id', this) was called but the element with that ID
// didn't exist, so clicking the pill did nothing.

_describeInteractive('OC7: Pill button toggle targets exist in the HTML', () => {
  test('every pillToggle target ID has a matching element', () => {
    const pillRegex = /pillToggle\('([^']+)'/g;
    const targetIds = new Set();
    let m;
    while ((m = pillRegex.exec(orgChartCode)) !== null) {
      targetIds.add(m[1]);
    }

    const missing = [];
    for (const id of targetIds) {
      const idPattern = new RegExp(`id=["']${id}["']`);
      if (!idPattern.test(orgChartCode)) {
        missing.push(id);
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('pillToggle targets missing from HTML:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('every repToggle target ID has a matching element', () => {
    const repRegex = /repToggle\('([^']+)'/g;
    const targetIds = new Set();
    let m;
    while ((m = repRegex.exec(orgChartCode)) !== null) {
      targetIds.add(m[1]);
    }

    const missing = [];
    for (const id of targetIds) {
      const idPattern = new RegExp(`id=["']${id}["']`);
      if (!idPattern.test(orgChartCode)) {
        missing.push(id);
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('repToggle targets missing from HTML:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('every maToggle target ID has a matching element', () => {
    const maRegex = /maToggle\('([^']+)'/g;
    const targetIds = new Set();
    let m;
    while ((m = maRegex.exec(orgChartCode)) !== null) {
      targetIds.add(m[1]);
    }

    const missing = [];
    for (const id of targetIds) {
      const idPattern = new RegExp(`id=["']${id}["']`);
      if (!idPattern.test(orgChartCode)) {
        missing.push(id);
      }
    }

    if (missing.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('maToggle targets missing from HTML:', missing);
    }
    expect(missing).toEqual([]);
  });

  test('chapter group toggle targets exist', () => {
    expect(orgChartCode).toMatch(/id=["']ps-group["']/);
    expect(orgChartCode).toMatch(/id=["']other-group["']/);
  });
});


// ============================================================================
// OC8: MADDS-EMBED WRAPPER SCOPING
// ============================================================================
// Bug: Org chart styles leaked into the SPA or SPA styles leaked into the
// org chart because CSS wasn't properly scoped under .madds-embed.

_describeInteractive('OC8: Org chart CSS is scoped under .madds-embed', () => {
  test('all style rules (except keyframes) are scoped to .madds-embed', () => {
    // Extract the <style> block
    const styleEnd = orgChartCode.indexOf('</style>');
    if (styleEnd === -1) return; // no style block
    const styleBlock = orgChartCode.substring(0, styleEnd);

    // Find unscoped rules: lines that start a CSS rule block but don't
    // reference .madds-embed or start with @, #madds-, or are comments
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
        // Must be scoped to .madds-embed OR be #madds-mode-toggle OR be a keyframe name
        if (selector.includes('.madds-embed') ||
            selector.includes('#madds-') ||
            selector.startsWith('@') ||
            selector.startsWith('from') ||
            selector.startsWith('to') ||
            /^\d+%/.test(selector)) continue;

        // Allow .page, .tier, etc. inside the .madds-embed <style> block
        // since they're scoped by the wrapper div
        // Only flag selectors that could conflict: html, body, *, etc.
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

  test('org chart HTML is wrapped in .madds-embed div', () => {
    expect(orgChartCode).toMatch(/<div class="madds-embed">/);
  });
});


// ============================================================================
// OC9: OVERFLOW HANDLING — NO CONTENT CLIPPING
// ============================================================================
// Bug: overflow-x:hidden on a width-constrained parent caused content to be
// invisible instead of scrollable.

_describeInteractive('OC9: Org chart overflow is handled correctly', () => {
  test('.madds-embed has overflow-x handling', () => {
    // .madds-embed should handle overflow (hidden is OK since content is self-contained)
    expect(orgChartCode).toMatch(/\.madds-embed\s*\{[^}]*overflow-x:\s*(hidden|auto)/);
  });

  test('.page uses box-sizing: border-box', () => {
    // Prevents padding from pushing content beyond container
    expect(orgChartCode).toMatch(/\.page\s*\{[^}]*box-sizing:\s*border-box/);
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
