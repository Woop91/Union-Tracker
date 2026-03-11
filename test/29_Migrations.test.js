/**
 * Tests for 29_Migrations.gs
 * Covers one-time migration functions.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '29_Migrations.gs'
]);

describe('29 function existence', () => {
  test('migrateContactLogFolderUrlColumn is defined', () => {
    expect(typeof migrateContactLogFolderUrlColumn).toBe('function');
  });
});
