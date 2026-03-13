/**
 * Tests for 10c_FormHandlers.gs
 * Covers form field management, timeline views, and column operations.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '05_Integrations.gs', '10c_FormHandlers.gs'
]);

describe('10c function existence', () => {
  const required = [
    'getCurrentStewardInfo_',
    'sanitizeFolderName_', 'shareWithCoordinators_',
    'refreshMemberDirectoryFormulas',
    'rebuildDashboard', 'refreshAllFormulas'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('sanitizeFolderName_', () => {
  test('removes special characters', () => {
    const result = sanitizeFolderName_('Test / Folder : Name');
    expect(result).not.toContain('/');
    expect(result).not.toContain(':');
  });

  test('handles empty string', () => {
    const result = sanitizeFolderName_('');
    expect(typeof result).toBe('string');
  });
});
