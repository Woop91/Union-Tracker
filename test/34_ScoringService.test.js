/**
 * Tests for 34_ScoringService.gs
 *
 * Covers all ScoringService functions: calculateEngagementScore,
 * calculateProfileScore, calculateGrievanceScore, calculateCompositeScore,
 * getScoreColor, autoAssignMembers, and locationBasedAssign.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Stub ConfigReader before loading ScoringService
global.ConfigReader = {
  getConfig: jest.fn(function() {
    return {
      maxVolunteerHours: 20,
      scoreWeightEngagement: 70,
      scoreWeightProfile: 20,
      scoreWeightGrievance: 10,
      scoreThresholdGreen: 70,
      scoreThresholdYellow: 40,
      grievanceScoreDirection: 'Negative',
    };
  })
};

loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '34_ScoringService.gs']);

// ============================================================================
// Module existence
// ============================================================================

describe('ScoringService module', function() {
  test('ScoringService is defined with expected API', function() {
    expect(ScoringService).toBeDefined();
    expect(typeof ScoringService.calculateEngagementScore).toBe('function');
    expect(typeof ScoringService.calculateProfileScore).toBe('function');
    expect(typeof ScoringService.calculateGrievanceScore).toBe('function');
    expect(typeof ScoringService.calculateCompositeScore).toBe('function');
    expect(typeof ScoringService.getScoreColor).toBe('function');
    expect(typeof ScoringService.autoAssignMembers).toBe('function');
    expect(typeof ScoringService.locationBasedAssign).toBe('function');
  });
});

// ============================================================================
// calculateEngagementScore
// ============================================================================

describe('ScoringService.calculateEngagementScore', function() {
  test('returns 0 for all-zero inputs', function() {
    var score = ScoringService.calculateEngagementScore(0, 0, '');
    expect(score).toBe(0);
  });

  test('returns 100 for perfect inputs (max hours, 100% open, committees)', function() {
    var score = ScoringService.calculateEngagementScore(20, 100, 'Safety');
    expect(score).toBeCloseTo(100, 1);
  });

  test('caps volunteer hours at maxVolunteerHours', function() {
    // hours = 40 (double max), openRate = 0, no committees => hoursScore = 100, rest = 0 => ~33.3
    var score = ScoringService.calculateEngagementScore(40, 0, '');
    expect(score).toBeCloseTo(100 / 3, 1);
  });

  test('caps openRate at 100', function() {
    // openRate = 200 should be capped to 100
    var score = ScoringService.calculateEngagementScore(0, 200, '');
    expect(score).toBeCloseTo(100 / 3, 1);
  });

  test('committees empty string gives 0 committee score', function() {
    var score = ScoringService.calculateEngagementScore(0, 0, '   ');
    expect(score).toBe(0);
  });

  test('handles non-numeric inputs gracefully', function() {
    var score = ScoringService.calculateEngagementScore('', '', null);
    expect(score).toBe(0);
  });

  test('partial hours gives proportional score', function() {
    // 10 out of 20 hours = 50% hoursScore, openRate = 0, no committees => 50/3
    var score = ScoringService.calculateEngagementScore(10, 0, '');
    expect(score).toBeCloseTo(50 / 3, 1);
  });
});

// ============================================================================
// calculateProfileScore
// ============================================================================

describe('ScoringService.calculateProfileScore', function() {
  test('returns 0 for all empty fields', function() {
    var score = ScoringService.calculateProfileScore({});
    expect(score).toBe(0);
  });

  test('returns 100 for all 10 fields filled', function() {
    var score = ScoringService.calculateProfileScore({
      phone: '555-1234',
      email: 'test@example.com',
      street: '123 Main St',
      city: 'Springfield',
      state: 'MA',
      zip: '01234',
      shirtSize: 'M',
      preferredComm: 'Email',
      bestTime: 'Morning',
      officeDays: 'Mon,Tue',
    });
    expect(score).toBe(100);
  });

  test('returns 50 for 5 out of 10 fields filled', function() {
    var score = ScoringService.calculateProfileScore({
      phone: '555-1234',
      email: 'test@example.com',
      street: '123 Main St',
      city: 'Springfield',
      state: 'MA',
    });
    expect(score).toBe(50);
  });

  test('ignores whitespace-only values', function() {
    var score = ScoringService.calculateProfileScore({
      phone: '   ',
      email: 'test@example.com',
    });
    expect(score).toBe(10); // only email counts (1/10)
  });

  test('returns 10 for one filled field', function() {
    var score = ScoringService.calculateProfileScore({ email: 'a@b.com' });
    expect(score).toBe(10);
  });
});

// ============================================================================
// calculateGrievanceScore — Negative direction (default)
// ============================================================================

describe('ScoringService.calculateGrievanceScore — Negative direction', function() {
  // v4.55.1 A08-BUG-01: progressive scaling replaced the 0/40/70 step function.
  // New anchors (negative direction, hasOpenGrievance=true):
  //   < 0 days  → 0 (overdue)
  //   0-7 days  → 30..40 (linear)
  //   7-30 days → 50..70 (linear)
  //  30-90 days → 70..85 (linear)
  //    > 90     → 90
  test('no open grievance gives 100', function() {
    var score = ScoringService.calculateGrievanceScore(false, '', 30, 'Negative');
    expect(score).toBe(100);
  });

  test('open grievance with 10 days to deadline scales between 50 and 70', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Open', 10, 'Negative');
    expect(score).toBeGreaterThanOrEqual(50);
    expect(score).toBeLessThanOrEqual(70);
  });

  test('open grievance with less than 7 days gives 30-40', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Open', 5, 'Negative');
    expect(score).toBeGreaterThanOrEqual(30);
    expect(score).toBeLessThanOrEqual(40);
  });

  test('open grievance with negative days gives 0 (overdue)', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Open', -3, 'Negative');
    expect(score).toBe(0);
  });

  test('open grievance with exactly 7 days gives 40', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Open', 7, 'Negative');
    expect(score).toBe(40);
  });

  test('open grievance with 60 days scores higher than 8 days (progressive scaling)', function() {
    var short = ScoringService.calculateGrievanceScore(true, 'Open', 8, 'Negative');
    var longD = ScoringService.calculateGrievanceScore(true, 'Open', 60, 'Negative');
    expect(longD).toBeGreaterThan(short);
  });
});

// ============================================================================
// calculateGrievanceScore — Positive direction
// ============================================================================

describe('ScoringService.calculateGrievanceScore — Positive direction', function() {
  test('open non-closed grievance gives 100 (active engagement is good)', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Open', 10, 'Positive');
    expect(score).toBe(100);
  });

  test('no open grievance gives 50', function() {
    var score = ScoringService.calculateGrievanceScore(false, '', 0, 'Positive');
    expect(score).toBe(50);
  });

  test('closed grievance gives 50', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Closed', 0, 'Positive');
    expect(score).toBe(50);
  });

  test('withdrawn grievance gives 50', function() {
    var score = ScoringService.calculateGrievanceScore(true, 'Withdrawn', 0, 'Positive');
    expect(score).toBe(50);
  });
});

// ============================================================================
// calculateCompositeScore
// ============================================================================

describe('ScoringService.calculateCompositeScore', function() {
  test('calculates weighted composite with default weights (70/20/10)', function() {
    // engagement=100, profile=100, grievance=100 => 100
    expect(ScoringService.calculateCompositeScore(100, 100, 100)).toBe(100);
  });

  test('returns 0 for all-zero inputs', function() {
    expect(ScoringService.calculateCompositeScore(0, 0, 0)).toBe(0);
  });

  test('weights engagement at 70%', function() {
    // engagement=100, profile=0, grievance=0 => 70
    expect(ScoringService.calculateCompositeScore(100, 0, 0)).toBe(70);
  });

  test('weights profile at 20%', function() {
    // engagement=0, profile=100, grievance=0 => 20
    expect(ScoringService.calculateCompositeScore(0, 100, 0)).toBe(20);
  });

  test('weights grievance at 10%', function() {
    // engagement=0, profile=0, grievance=100 => 10
    expect(ScoringService.calculateCompositeScore(0, 0, 100)).toBe(10);
  });

  test('rounds result to nearest integer', function() {
    // 50*0.7 + 50*0.2 + 50*0.1 = 35+10+5 = 50 (exact)
    expect(ScoringService.calculateCompositeScore(50, 50, 50)).toBe(50);
  });

  test('typical mixed score rounds correctly', function() {
    // engagement=80*0.7=56, profile=60*0.2=12, grievance=100*0.1=10 => 78
    expect(ScoringService.calculateCompositeScore(80, 60, 100)).toBe(78);
  });
});

// ============================================================================
// getScoreColor
// ============================================================================

describe('ScoringService.getScoreColor', function() {
  test('returns green (#4CAF50) for score >= 70', function() {
    expect(ScoringService.getScoreColor(70)).toBe('#4CAF50');
    expect(ScoringService.getScoreColor(100)).toBe('#4CAF50');
    expect(ScoringService.getScoreColor(85)).toBe('#4CAF50');
  });

  test('returns yellow (#FFC107) for score >= 40 and < 70', function() {
    expect(ScoringService.getScoreColor(40)).toBe('#FFC107');
    expect(ScoringService.getScoreColor(55)).toBe('#FFC107');
    expect(ScoringService.getScoreColor(69)).toBe('#FFC107');
  });

  test('returns red (#f44336) for score < 40', function() {
    expect(ScoringService.getScoreColor(0)).toBe('#f44336');
    expect(ScoringService.getScoreColor(20)).toBe('#f44336');
    expect(ScoringService.getScoreColor(39)).toBe('#f44336');
  });
});

// ============================================================================
// autoAssignMembers
// ============================================================================

describe('ScoringService.autoAssignMembers', function() {
  test('returns empty array when stewards list is empty', function() {
    var members = [{ email: 'a@b.com', steward: '' }];
    expect(ScoringService.autoAssignMembers(members, [])).toEqual([]);
  });

  test('skips members that already have a steward', function() {
    var members = [
      { email: 'a@b.com', steward: 'steward1' },
      { email: 'b@b.com', steward: '' },
    ];
    var result = ScoringService.autoAssignMembers(members, ['steward1']);
    expect(result.length).toBe(1);
    expect(result[0].email).toBe('b@b.com');
  });

  test('round-robins across stewards', function() {
    var members = [
      { email: 'a@b.com', steward: '' },
      { email: 'b@b.com', steward: '' },
      { email: 'c@b.com', steward: '' },
    ];
    var result = ScoringService.autoAssignMembers(members, ['s1', 's2']);
    expect(result[0].steward).toBe('s1');
    expect(result[1].steward).toBe('s2');
    expect(result[2].steward).toBe('s1');
  });

  test('handles single steward assigning all unassigned', function() {
    var members = [
      { email: 'a@b.com', steward: '' },
      { email: 'b@b.com', steward: '' },
    ];
    var result = ScoringService.autoAssignMembers(members, ['s1']);
    expect(result.every(function(r) { return r.steward === 's1'; })).toBe(true);
  });
});

// ============================================================================
// locationBasedAssign
// ============================================================================

describe('ScoringService.locationBasedAssign', function() {
  test('matches members by location to steward from map', function() {
    var members = [{ email: 'a@b.com', steward: '', location: 'Boston' }];
    var result = ScoringService.locationBasedAssign(members, 'Boston:s1,Springfield:s2', ['s1', 's2']);
    expect(result[0].steward).toBe('s1');
  });

  test('is case-insensitive for location matching', function() {
    var members = [{ email: 'a@b.com', steward: '', location: 'BOSTON' }];
    var result = ScoringService.locationBasedAssign(members, 'boston:s1', ['s1']);
    expect(result[0].steward).toBe('s1');
  });

  test('falls back to round-robin for unmatched locations', function() {
    var members = [
      { email: 'a@b.com', steward: '', location: 'Unknown City' },
      { email: 'b@b.com', steward: '', location: 'Unknown City 2' },
    ];
    var result = ScoringService.locationBasedAssign(members, 'Boston:s1', ['s1', 's2']);
    expect(result[0].steward).toBe('s1');
    expect(result[1].steward).toBe('s2');
  });

  test('skips members that already have a steward', function() {
    var members = [
      { email: 'a@b.com', steward: 'existing', location: 'Boston' },
      { email: 'b@b.com', steward: '', location: 'Boston' },
    ];
    var result = ScoringService.locationBasedAssign(members, 'Boston:s1', ['s1']);
    expect(result.length).toBe(1);
    expect(result[0].email).toBe('b@b.com');
  });

  test('handles empty locationMapStr gracefully', function() {
    var members = [{ email: 'a@b.com', steward: '', location: 'Boston' }];
    var result = ScoringService.locationBasedAssign(members, '', ['s1']);
    // No map match, falls back to round-robin
    expect(result[0].steward).toBe('s1');
  });

  test('handles null locationMapStr gracefully', function() {
    var members = [{ email: 'a@b.com', steward: '', location: 'Boston' }];
    var result = ScoringService.locationBasedAssign(members, null, ['s1']);
    expect(result[0].steward).toBe('s1');
  });
});
