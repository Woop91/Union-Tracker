/**
 * Tests for 04b_AccessibilityFeatures.gs
 * Covers accessibility features, pomodoro, quick capture, import/export.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '04b_AccessibilityFeatures.gs'
]);

describe('04b function existence', () => {
  const required = [
    'getCommonStyles', 'startPomodoroTimer',
    'getQuickCaptureNotes', 'saveQuickCaptureNotes',
    'clearQuickCaptureNotes',
    'showQuickCaptureNotepad',
    'processMemberImport',
    'parseCSVLine_', 'mapImportColumns_',
    'showThemeManager'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('getCommonStyles', () => {
  test('returns a CSS string', () => {
    const css = getCommonStyles();
    expect(typeof css).toBe('string');
    expect(css.length).toBeGreaterThan(50);
  });
});

describe('parseCSVLine_', () => {
  test('splits simple CSV', () => {
    const result = parseCSVLine_('a,b,c');
    expect(result).toEqual(['a', 'b', 'c']);
  });

  test('handles quoted fields', () => {
    const result = parseCSVLine_('"hello, world",b,c');
    expect(result[0]).toBe('hello, world');
    expect(result.length).toBe(3);
  });

  test('handles empty string', () => {
    const result = parseCSVLine_('');
    expect(Array.isArray(result)).toBe(true);
  });
});

describe('Quick Capture (PropertiesService)', () => {
  test('saveQuickCaptureNotes stores text', () => {
    saveQuickCaptureNotes('Test note');
    // Verify it used PropertiesService
    expect(PropertiesService.getUserProperties).toHaveBeenCalled();
  });
});
