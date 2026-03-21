/**
 * Org Chart Integrity Guards — Prevents recurring issues with the org chart
 * embedded inside the SPA:
 *
 *   OC1: Org chart container must NOT be constrained by page-layout-content max-width
 *   OC8: Org chart HTML must be wrapped in .madds-embed div
 *   OC10: getOrgChartHtml server function must exist and load the file
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

describe('OC1: Org chart is not clipped by parent max-width', () => {
  test('styles.html has a max-width override for org chart content', () => {
    const hasOverride =
      stylesCode.includes('.madds-embed') && stylesCode.includes('max-width: none') ||
      stylesCode.includes('orgchart-content') && stylesCode.includes('max-width: none');
    expect(hasOverride).toBe(true);
  });

  test('renderOrgChart adds orgchart-content class to container', () => {
    expect(indexCode).toMatch(/orgchart-content/);
  });
});


// ============================================================================
// OC8: MADDS-EMBED WRAPPER SCOPING
// ============================================================================

describe('OC8: Org chart CSS is scoped under .madds-embed', () => {
  test('org chart HTML is wrapped in .madds-embed div', () => {
    expect(orgChartCode).toMatch(/<div class="madds-embed">/);
  });
});


// ============================================================================
// OC10: getOrgChartHtml SERVER FUNCTION EXISTS
// ============================================================================

describe('OC10: getOrgChartHtml server function integrity', () => {
  const webAppCode = read('22_WebDashApp.gs');

  test('getOrgChartHtml function exists', () => {
    expect(webAppCode).toMatch(/function getOrgChartHtml\s*\(/);
  });

  test('getOrgChartHtml loads org_chart file', () => {
    expect(webAppCode).toMatch(/createHtmlOutputFromFile\s*\(\s*['"]org_chart['"]\s*\)/);
  });

  test('getOrgChartHtml has error handling', () => {
    const funcMatch = webAppCode.match(/function getOrgChartHtml[\s\S]*?^}/m);
    expect(funcMatch).not.toBeNull();
    expect(funcMatch[0]).toMatch(/try\s*\{/);
    expect(funcMatch[0]).toMatch(/catch/);
  });

  test('renderOrgChart calls getOrgChartHtml', () => {
    expect(indexCode).toMatch(/\.getOrgChartHtml\s*\(/);
  });
});
