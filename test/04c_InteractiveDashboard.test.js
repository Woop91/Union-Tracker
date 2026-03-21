/**
 * Tests for 04c_InteractiveDashboard.gs
 * The Interactive Dashboard was deprecated and removed in v4.31.0.
 * All functionality consolidated into Steward Dashboard (04d) and web app.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '04c_InteractiveDashboard.gs'
]);

describe('04c_InteractiveDashboard.gs (deprecated)', () => {
  test('file loads without error (empty stub)', () => {
    // File is intentionally empty — all Interactive Dashboard code removed in v4.31.0
    expect(true).toBe(true);
  });

  test('deprecated functions no longer exist', () => {
    const removed = [
      'showInteractiveDashboardTab', 'getInteractiveDashboardHtml',
      'getInteractiveOverviewData', 'getConfigDropdownValues',
      'getInteractiveMemberData', 'getInteractiveGrievanceData',
      'getMyStewardCases', 'getInteractiveAnalyticsData',
      'getInteractiveResourceLinks', 'saveInteractiveMember'
    ];

    removed.forEach(fn => {
      expect(typeof global[fn]).not.toBe('function');
    });
  });
});
