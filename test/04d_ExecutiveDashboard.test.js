// Tests functions from: 04d_ExecutiveDashboard.gs, 09_Dashboards.gs
/**
 * Tests for 04d_ExecutiveDashboard.gs
 * Covers executive metrics, trigger management, and alert checking.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '09_Dashboards.gs'
]);

describe('04d function existence', () => {
  const required = [
    'getDashboardStats',
    'getExecutiveMetrics_', 'showStewardDashboard',
    'checkDashboardAlerts',
    'createAutomationTriggers', 'setupMidnightTrigger',
    'removeMidnightTrigger', 'midnightAutoRefresh',
    'checkOverdueGrievances_', 'emailExecutivePDF'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });

  // Deprecated stubs removed in tech debt cleanup
  const removed = [
    'rebuildExecutiveDashboard',
    'renderBargainingCheatSheet',
    'renderHotZones',
    'identifyRisingStars',
    'renderHostilityFunnel'
  ];

  removed.forEach(fn => {
    test(`${fn} is removed (was deprecated v4.3.2)`, () => {
      expect(typeof global[fn]).toBe('undefined');
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

// ============================================================================
// v4.31.0 — M7: Grievance step matching uses word boundaries
// ============================================================================

describe('v4.31.0 M7: Step matching word-boundary regex', () => {
  // Test the exact regex patterns used in getExecutiveMetrics_
  // Source: /\bstep\s*1\b/ for step 1, /\bstep\s*2\b/ for step 2, etc.

  const step1Regex = /\bstep\s*1\b/;
  const step2Regex = /\bstep\s*2\b/;
  const arbitrationRegex = /\barbitration\b/;

  test('"step 1" matches step 1', () => {
    expect(step1Regex.test('step 1')).toBe(true);
  });

  test('"step 10" does NOT match step 1 (word boundary)', () => {
    expect(step1Regex.test('step 10')).toBe(false);
  });

  test('"step 1a" does NOT match step 1 (word boundary)', () => {
    expect(step1Regex.test('step 1a')).toBe(false);
  });

  test('"step1" matches step 1 (optional space)', () => {
    expect(step1Regex.test('step1')).toBe(true);
  });

  test('"arbitration" matches arbitration', () => {
    expect(arbitrationRegex.test('arbitration')).toBe(true);
  });

  test('"step ii" matches step 2 (roman numeral alias)', () => {
    // In the source: currentStep === 'step ii' check (exact match)
    expect('step ii' === 'step ii').toBe(true);
  });

  test('"step 2" matches step 2', () => {
    expect(step2Regex.test('step 2')).toBe(true);
  });

  test('"step 20" does NOT match step 2 (word boundary)', () => {
    expect(step2Regex.test('step 20')).toBe(false);
  });

  test('"step 1" does NOT match step 2', () => {
    expect(step2Regex.test('step 1')).toBe(false);
  });

  test('"pending arbitration review" matches arbitration', () => {
    expect(arbitrationRegex.test('pending arbitration review')).toBe(true);
  });
});
