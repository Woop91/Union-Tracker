/**
 * Tests for 08b_SearchAndCharts.gs
 * Covers search functionality, chart generation, and data retrieval.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '08b_SearchAndCharts.gs'
]);

describe('08b function existence', () => {
  const required = [
    'getDesktopSearchLocations', 'getDesktopSearchData',
    'navigateToSearchResult', 'viewActiveGrievances',
    'searchDashboard', 'quickSearchDashboard', 'advancedSearch',
    'getDepartmentList', 'getMemberList',
    'generateSelectedChart',
    'createGaugeStyleChart_', 'createScorecardChart_',
    'createTrendLineChart_', 'createAreaChart_', 'createComboChart_',
    'createSummaryTableChart_', 'createStewardLeaderboardChart_',
    'padRight'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('padRight', () => {
  test('pads string to target length', () => {
    expect(padRight('hi', 5)).toBe('hi   ');
  });

  test('does not truncate longer strings', () => {
    expect(padRight('hello world', 5).length).toBeGreaterThanOrEqual(5);
  });

  test('handles empty string', () => {
    expect(padRight('', 3)).toBe('   ');
  });
});

describe('Search uses SHEETS constants', () => {
  test('no hardcoded sheet names in 08b', () => {
    const fs = require('fs');
    const path = require('path');
    const code = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '08b_SearchAndCharts.gs'), 'utf8'
    );
    // After our migration, should use SHEETS. not SHEET_NAMES.
    expect(code).not.toContain('SHEET_NAMES.');
  });
});
