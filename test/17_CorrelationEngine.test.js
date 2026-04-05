/**
 * Tests for 17_CorrelationEngine.gs
 *
 * Covers statistical primitives (mean, stddev, Pearson, Spearman, ranks,
 * chi-square), classification, data extraction helpers, insight generation,
 * data point building, and the top-level getCorrelationInsights entry point.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock getUnifiedDashboardData and isTruthyValue before loading source
global.getUnifiedDashboardData = jest.fn(() => JSON.stringify({
  totalMembers: 100, openCases: 10, winRate: 60, moraleScore: 7,
  statusDistribution: { Open: 5, Closed: 3 },
  locationBreakdown: { HQ: 50, Branch: 30 },
  unitBreakdown: { 'Unit A': 30 },
  grievancesByCategory: {},
  chartDrillDown: { statusByCase: {}, locationByCase: {} },
  stepProgression: { step1: 5, step2: 3, step3: 1, arb: 0 },
  stewardPerformance: [],
  monthlyFilings: [],
  satisfactionByLocation: {},
  participationByLocation: {},
  satisfactionByUnit: {},
  avgDaysAtStep: {}
}));
global.isTruthyValue = jest.fn(v => !!v);
// Mock getConfigValue_ (defined in 10_Main.gs) — enable correlation by default for tests
global.getConfigValue_ = jest.fn(() => 'yes');

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '17_CorrelationEngine.gs']);

// Clean up module-level mocks between tests
afterEach(() => {
  jest.restoreAllMocks();
});

// ============================================================================
// 1. statMean_
// ============================================================================

describe('statMean_', () => {
  test('returns 0 for empty array', () => {
    expect(statMean_([])).toBe(0);
  });

  test('returns 0 for null/undefined', () => {
    expect(statMean_(null)).toBe(0);
    expect(statMean_(undefined)).toBe(0);
  });

  test('returns the value itself for single-element array', () => {
    expect(statMean_([42])).toBe(42);
  });

  test('calculates mean of multiple values', () => {
    expect(statMean_([2, 4, 6, 8, 10])).toBe(6);
  });

  test('handles negative values', () => {
    expect(statMean_([-5, -3, 0, 3, 5])).toBe(0);
  });

  test('handles decimal precision', () => {
    const result = statMean_([1.5, 2.5, 3.5]);
    expect(result).toBeCloseTo(2.5, 10);
  });
});

// ============================================================================
// ============================================================================
// 3. pearsonCorrelation_
// ============================================================================

describe('pearsonCorrelation_', () => {
  test('returns 1.0 for perfect positive correlation', () => {
    const r = pearsonCorrelation_([1, 2, 3, 4, 5], [2, 4, 6, 8, 10]);
    expect(r).toBeCloseTo(1.0, 5);
  });

  test('returns -1.0 for perfect negative correlation', () => {
    const r = pearsonCorrelation_([1, 2, 3, 4, 5], [10, 8, 6, 4, 2]);
    expect(r).toBeCloseTo(-1.0, 5);
  });

  test('returns ~0 for no correlation', () => {
    // Deliberately uncorrelated data
    const r = pearsonCorrelation_([1, 2, 3, 4, 5], [2, 4, 1, 5, 3]);
    expect(Math.abs(r)).toBeLessThan(0.5);
  });

  test('returns 0 for fewer than 5 pairs', () => {
    const r = pearsonCorrelation_([1, 2, 3, 4], [5, 6, 7, 8]);
    expect(r).toBe(0);
  });

  test('returns 1.0 for identical arrays', () => {
    const r = pearsonCorrelation_([3, 7, 1, 9, 5], [3, 7, 1, 9, 5]);
    expect(r).toBeCloseTo(1.0, 5);
  });

  test('returns 0 for constant x array (zero variance)', () => {
    const r = pearsonCorrelation_([5, 5, 5, 5, 5], [1, 2, 3, 4, 5]);
    expect(r).toBe(0);
  });

  test('uses min length when arrays are mismatched', () => {
    // 6 vs 5 elements; should use first 5 of each
    const r = pearsonCorrelation_([1, 2, 3, 4, 5, 6], [2, 4, 6, 8, 10]);
    expect(r).toBeCloseTo(1.0, 5);
  });

  test('computes moderate correlation', () => {
    // Data with moderate positive correlation
    const x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    const y = [2, 1, 4, 3, 6, 5, 8, 7, 10, 9];
    const r = pearsonCorrelation_(x, y);
    expect(r).toBeGreaterThan(0.5);
    expect(r).toBeLessThan(1.0);
  });
});

// ============================================================================
// 4. spearmanCorrelation_
// ============================================================================

describe('spearmanCorrelation_', () => {
  test('returns 0 for fewer than 3 pairs', () => {
    expect(spearmanCorrelation_([1, 2], [3, 4])).toBe(0);
  });

  test('returns ~1 for perfectly monotone increasing data', () => {
    // Needs >= 5 for the inner pearson call
    const rho = spearmanCorrelation_([10, 20, 30, 40, 50], [100, 200, 300, 400, 500]);
    expect(rho).toBeCloseTo(1.0, 5);
  });

  test('handles ties in input data', () => {
    // Ties: [10, 20, 20, 30, 40] — should still work
    const rho = spearmanCorrelation_([10, 20, 20, 30, 40], [1, 2, 3, 4, 5]);
    expect(rho).toBeGreaterThan(0.8);
  });

  test('returns ~-1 for perfectly monotone decreasing data', () => {
    const rho = spearmanCorrelation_([1, 2, 3, 4, 5], [50, 40, 30, 20, 10]);
    expect(rho).toBeCloseTo(-1.0, 5);
  });
});

// ============================================================================
// 5. toRanks_
// ============================================================================

describe('toRanks_', () => {
  test('ranks simple ascending array', () => {
    expect(toRanks_([10, 20, 30])).toEqual([1, 2, 3]);
  });

  test('handles ties with fractional ranks', () => {
    // [10, 20, 20, 30] -> positions sorted: 10(1), 20(2), 20(3), 30(4)
    // Ties at 20 get avg of positions 2,3 = 2.5
    expect(toRanks_([10, 20, 20, 30])).toEqual([1, 2.5, 2.5, 4]);
  });

  test('ranks single element', () => {
    expect(toRanks_([42])).toEqual([1]);
  });

  test('ranks descending order correctly', () => {
    // [30, 20, 10] -> sorted positions: 10(1)->idx2, 20(2)->idx1, 30(3)->idx0
    expect(toRanks_([30, 20, 10])).toEqual([3, 2, 1]);
  });
});

// ============================================================================
// 7. classifyCorrelation_
// ============================================================================

describe('classifyCorrelation_', () => {
  test('classifies strong correlation (r=0.8, n=30)', () => {
    const cls = classifyCorrelation_(0.8, 30);
    expect(cls.strength).toBe('strong');
    expect(cls.confidence).toBe('high');
    expect(cls.reliable).toBe(true);
  });

  test('classifies moderate correlation (r=0.5, n=30)', () => {
    const cls = classifyCorrelation_(0.5, 30);
    expect(cls.strength).toBe('moderate');
    expect(cls.confidence).toBe('high');
    expect(cls.reliable).toBe(true);
  });

  test('classifies weak correlation (r=0.3, n=30)', () => {
    const cls = classifyCorrelation_(0.3, 30);
    expect(cls.strength).toBe('weak');
    expect(cls.confidence).toBe('high');
    expect(cls.reliable).toBe(true);
  });

  test('classifies negligible correlation (r=0.1, n=30)', () => {
    const cls = classifyCorrelation_(0.1, 30);
    expect(cls.strength).toBe('negligible');
    expect(cls.confidence).toBe('high');
    expect(cls.reliable).toBe(false); // absR < 0.2
  });

  test('returns insufficient confidence for n<5', () => {
    const cls = classifyCorrelation_(0.9, 3);
    expect(cls.confidence).toBe('insufficient');
    expect(cls.reliable).toBe(false);
  });

  test('reliable depends on sample size (moderate conf, n=15)', () => {
    // n=15 => moderate confidence => reliable if absR >= 0.35
    const clsReliable = classifyCorrelation_(0.5, 15);
    expect(clsReliable.confidence).toBe('moderate');
    expect(clsReliable.reliable).toBe(true);

    const clsUnreliable = classifyCorrelation_(0.3, 15);
    expect(clsUnreliable.confidence).toBe('moderate');
    expect(clsUnreliable.reliable).toBe(false);
  });
});

// ============================================================================
// 9. generateInsight_
// ============================================================================

describe('generateInsight_', () => {
  test('returns insufficient data message for insufficient confidence', () => {
    const cls = { strength: 'negligible', confidence: 'insufficient', reliable: false };
    const result = generateInsight_('X', 'Y', 0.1, cls, 'location', 3);
    expect(result).toContain('Not enough data points');
    expect(result).toContain('location');
  });

  test('mentions sample size in unreliable weak result', () => {
    const cls = { strength: 'weak', confidence: 'low', reliable: false };
    const result = generateInsight_('X', 'Y', 0.25, cls, 'unit', 8);
    expect(result).toContain('N=8');
  });

  test('positive direction says "higher"', () => {
    const cls = { strength: 'strong', confidence: 'high', reliable: true };
    const result = generateInsight_('engagement', 'satisfaction', 0.8, cls, 'location', 30);
    expect(result).toContain('higher');
  });

  test('negative direction says "lower"', () => {
    const cls = { strength: 'strong', confidence: 'high', reliable: true };
    const result = generateInsight_('caseload', 'win rate', -0.8, cls, 'steward', 30);
    expect(result).toContain('lower');
  });

  test('strong reliable correlation says "Strong pattern"', () => {
    const cls = { strength: 'strong', confidence: 'high', reliable: true };
    const result = generateInsight_('X', 'Y', 0.8, cls, 'unit', 30);
    expect(result).toContain('Strong pattern');
  });
});

// ============================================================================
// 10. buildDataPoints_
// ============================================================================

describe('buildDataPoints_', () => {
  test('returns array of objects with label, x, y', () => {
    const points = buildDataPoints_(['A', 'B'], [1, 2], [3, 4]);
    expect(points).toEqual([
      { label: 'A', x: 1, y: 3 },
      { label: 'B', x: 2, y: 4 }
    ]);
  });

  test('rounds values to 2 decimal places', () => {
    const points = buildDataPoints_(['A'], [1.23456], [7.89012]);
    expect(points[0].x).toBe(1.23);
    expect(points[0].y).toBe(7.89);
  });

  test('handles empty arrays', () => {
    const points = buildDataPoints_([], [], []);
    expect(points).toEqual([]);
  });
});

// ============================================================================
// 11. getCorrelationInsights (integration)
// ============================================================================

describe('getCorrelationInsights', () => {
  test('returns a JSON string', () => {
    const result = getCorrelationInsights(false);
    expect(typeof result).toBe('string');
    const parsed = JSON.parse(result);
    expect(Array.isArray(parsed)).toBe(true);
  });

  test('calls getUnifiedDashboardData when no cachedData provided', () => {
    getUnifiedDashboardData.mockClear();
    getCorrelationInsights(false);
    expect(getUnifiedDashboardData).toHaveBeenCalled();
  });
});
