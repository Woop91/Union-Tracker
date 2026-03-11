/**
 * Tests for 04c_InteractiveDashboard.gs
 * Covers dashboard data retrieval, filters, analytics, and member save.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '04c_InteractiveDashboard.gs'
]);

describe('04c function existence', () => {
  const required = [
    'showInteractiveDashboardTab', 'getInteractiveDashboardHtml',
    'getInteractiveOverviewData', 'getConfigDropdownValues',
    'getInteractiveMemberData', 'getInteractiveGrievanceData',
    'getMyStewardCases', 'getInteractiveAnalyticsData',
    'getInteractiveResourceLinks', 'getInteractiveMemberFilters',
    'saveInteractiveMember'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('getConfigDropdownValues', () => {
  test('returns an object', () => {
    const result = getConfigDropdownValues();
    expect(typeof result).toBe('object');
  });
});

describe('getInteractiveOverviewData', () => {
  test('returns an object with expected shape', () => {
    const result = getInteractiveOverviewData();
    expect(typeof result).toBe('object');
    // Should have member and grievance counts
    expect(result).toHaveProperty('totalMembers');
  });
});
