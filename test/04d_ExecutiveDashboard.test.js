/**
 * Tests for 04d_ExecutiveDashboard.gs
 * Covers executive metrics, trigger management, and alert checking.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '04d_ExecutiveDashboard.gs'
]);

describe('04d function existence', () => {
  const required = [
    'rebuildExecutiveDashboard',
    'getDashboardStats',
    'getExecutiveMetrics_', 'showStewardDashboard',
    'checkDashboardAlerts', 'renderBargainingCheatSheet',
    'renderHotZones', 'identifyRisingStars', 'renderHostilityFunnel',
    'createAutomationTriggers', 'setupMidnightTrigger',
    'removeMidnightTrigger', 'midnightAutoRefresh',
    'checkOverdueGrievances_', 'emailExecutivePDF'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

describe('getDashboardStats', () => {
  test('returns defined result', () => {
    const result = getDashboardStats();
    expect(result).toBeDefined();
    // In mock environment may return string (JSON) or object
    const parsed = typeof result === 'string' ? JSON.parse(result) : result;
    expect(typeof parsed).toBe('object');
  });
});
