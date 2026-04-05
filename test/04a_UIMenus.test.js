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
    'showStewardDirectory',
    'openCellMultiSelectEditor',
    'showMultiSelectDialog', 'getMultiSelectHtml'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// getVisualControlPanelHtml — behavioral tests
// ============================================================================

describe('getVisualControlPanelHtml', () => {
  test('returns an HTML string', () => {
    const html = getVisualControlPanelHtml();
    expect(typeof html).toBe('string');
    expect(html.length).toBeGreaterThan(100);
  });

  test('contains DOCTYPE and html structure', () => {
    const html = getVisualControlPanelHtml();
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('<html>');
    expect(html).toContain('</html>');
  });

  test('contains Dark Mode toggle', () => {
    const html = getVisualControlPanelHtml();
    expect(html).toContain('Dark Mode');
    expect(html).toContain('id="darkMode"');
  });

  test('contains Focus Mode toggle', () => {
    const html = getVisualControlPanelHtml();
    expect(html).toContain('Focus Mode');
    expect(html).toContain('id="focusMode"');
  });

  test('contains theme selection buttons', () => {
    const html = getVisualControlPanelHtml();
    expect(html).toContain('Theme Selection');
    expect(html).toContain('theme-default');
    expect(html).toContain('theme-dark');
    expect(html).toContain('High Contrast');
  });

  test('contains Refresh Dashboard button', () => {
    const html = getVisualControlPanelHtml();
    expect(html).toContain('Refresh Dashboard');
    expect(html).toContain('refreshDashboard()');
  });
});

// ============================================================================
// getMultiSelectHtml — behavioral tests
// ============================================================================

describe('getMultiSelectHtml', () => {
  test('returns HTML for valid callback', () => {
    const html = getMultiSelectHtml(['A', 'B', 'C'], 'applyMultiSelectValue');
    expect(typeof html).toBe('string');
    expect(html).toContain('multi-select');
  });

  test('throws for invalid callback', () => {
    expect(() => getMultiSelectHtml(['A'], 'invalidFunc')).toThrow('Invalid callback');
  });

  test('embeds items as JSON in the HTML', () => {
    var items = [
      { id: 'opt1', label: 'Option One', selected: false },
      { id: 'opt2', label: 'Option Two', selected: true }
    ];
    var html = getMultiSelectHtml(items, 'applyMultiSelectValue');
    expect(html).toContain('opt1');
    expect(html).toContain('Option One');
    expect(html).toContain('opt2');
    expect(html).toContain('Option Two');
  });

  test('contains Select All and Select None buttons', () => {
    var items = [{ id: 'a', label: 'A', selected: false }];
    var html = getMultiSelectHtml(items, 'applyMultiSelectValue');
    expect(html).toContain('Select All');
    expect(html).toContain('Select None');
  });

  test('contains Apply and Cancel buttons', () => {
    var items = [{ id: 'a', label: 'A', selected: false }];
    var html = getMultiSelectHtml(items, 'applyMultiSelectValue');
    expect(html).toContain('Apply');
    expect(html).toContain('Cancel');
  });

  test('uses the specified callback name in script', () => {
    var items = [{ id: 'a', label: 'A', selected: false }];
    var html = getMultiSelectHtml(items, 'handleBulkStatusSelection');
    expect(html).toContain("'handleBulkStatusSelection'");
  });
});
