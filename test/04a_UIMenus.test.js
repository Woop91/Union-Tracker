/**
 * Tests for 04a_UIMenus.gs
 * Covers visual control panel, multi-select editor, and dashboard sidebar.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '04b_AccessibilityFeatures.gs', '04a_UIMenus.gs'
]);

describe('04a function existence', () => {
  const required = [
    'showVisualControlPanel', 'getVisualControlPanelHtml',
    'showExecutiveDashboard', 'showStewardDirectory',
    'refreshVisualsSimple_', 'openCellMultiSelectEditor',
    'showMultiSelectDialog', 'getMultiSelectHtml',
    'showDashboardSidebar', 'getDashboardSidebarHtml'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('getVisualControlPanelHtml', () => {
  test('returns an HTML string', () => {
    const html = getVisualControlPanelHtml();
    expect(typeof html).toBe('string');
    expect(html.length).toBeGreaterThan(100);
  });
});

describe('getMultiSelectHtml', () => {
  test('returns HTML for valid callback', () => {
    const html = getMultiSelectHtml(['A', 'B', 'C'], 'applyMultiSelectValue');
    expect(typeof html).toBe('string');
    expect(html).toContain('multi-select');
  });

  test('throws for invalid callback', () => {
    expect(() => getMultiSelectHtml(['A'], 'invalidFunc')).toThrow('Invalid callback');
  });
});
