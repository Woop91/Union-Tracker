/**
 * Org Chart Integrity Guards — SolidBase generic template
 *
 * Verifies the org chart template structure:
 *   OC1: CSS is scoped under .orgchart-embed
 *   OC2: Light/dark mode toggle exists and is wired
 *   OC3: renderOrgChart in index.html uses reliable script injection
 *   OC4: Container class override for org chart width
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
// OC1: CSS SCOPING
// ============================================================================

describe('OC1: Org chart CSS is scoped under .orgchart-embed', () => {
  test('org chart HTML is wrapped in .orgchart-embed div', () => {
    expect(orgChartCode).toMatch(/class="orgchart-embed/);
  });

  test('.orgchart-embed has overflow-x handling', () => {
    expect(orgChartCode).toMatch(/overflow-x:\s*hidden/);
  });

  test('light mode CSS variables are defined', () => {
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light\s*\{/);
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light[\s\S]*?--bg:/);
    expect(orgChartCode).toMatch(/\.orgchart-embed\.light[\s\S]*?--txt:/);
  });
});


// ============================================================================
// OC2: LIGHT/DARK MODE TOGGLE
// ============================================================================

describe('OC2: Light/dark mode toggle is functional', () => {
  test('orgchartToggleMode is assigned to window', () => {
    expect(orgChartCode).toMatch(/window\.orgchartToggleMode\s*=/);
  });

  test('toggle switches .light class on .orgchart-embed', () => {
    expect(orgChartCode).toMatch(/classList\.toggle\s*\(\s*['"]light['"]\s*\)/);
  });

  test('auto-detect reads AppState.isDark or localStorage', () => {
    expect(orgChartCode).toMatch(/AppState/);
    expect(orgChartCode).toMatch(/localStorage/);
  });
});


// ============================================================================
// OC3: RENDER FUNCTION IN INDEX.HTML
// ============================================================================

describe('OC3: renderOrgChart script re-execution is reliable', () => {
  test('renderOrgChart function exists in index.html', () => {
    expect(indexCode).toContain('function renderOrgChart(');
  });

  test('scripts are appended to document.head (not replaceChild)', () => {
    const funcStart = indexCode.indexOf('function renderOrgChart(');
    expect(funcStart).toBeGreaterThan(-1);
    let depth = 0, started = false, block = '';
    for (let i = funcStart; i < indexCode.length; i++) {
      if (indexCode[i] === '{') { depth++; started = true; }
      if (indexCode[i] === '}') { depth--; }
      if (started) block += indexCode[i];
      if (started && depth === 0) break;
    }
    expect(block).toMatch(/document\.head\.appendChild/);
    expect(block).not.toMatch(/replaceChild/);
  });

  test('scripts are iterated via querySelectorAll', () => {
    expect(indexCode).toMatch(/querySelectorAll\s*\(\s*['"]script['"]\s*\)/);
  });
});


// ============================================================================
// OC4: ORG CHART CONTAINER WIDTH OVERRIDE
// ============================================================================

describe('OC4: Org chart container allows wide content', () => {
  test('renderOrgChart adds orgchart-content class to container', () => {
    expect(indexCode).toMatch(/orgchart-content/);
  });

  test('styles.html has max-width override for org chart', () => {
    const hasOverride =
      (stylesCode.includes('orgchart-content') || stylesCode.includes('orgchart-embed') || stylesCode.includes('madds-embed'))
      && stylesCode.includes('max-width');
    expect(hasOverride).toBe(true);
  });
});


// ============================================================================
// OC5: ONCLICK HANDLERS REFERENCE DECLARED FUNCTIONS
// ============================================================================

describe('OC5: onclick handlers reference declared functions', () => {
  test('every onclick function call has a matching window.* declaration', () => {
    const onclickRegex = /onclick="(\w+)\s*\(/g;
    const calledFunctions = new Set();
    let m;
    while ((m = onclickRegex.exec(orgChartCode)) !== null) {
      calledFunctions.add(m[1]);
    }

    const windowFuncRegex = /window\.(\w+)\s*=\s*function/g;
    const declaredFunctions = new Set();
    while ((m = windowFuncRegex.exec(orgChartCode)) !== null) {
      declaredFunctions.add(m[1]);
    }

    const missing = [...calledFunctions].filter(fn => !declaredFunctions.has(fn));
    expect(missing).toEqual([]);
  });
});
