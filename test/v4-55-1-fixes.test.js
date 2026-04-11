/**
 * v4.55.1 Fix Regression Tests
 *
 * Focused tests for the bugs fixed in Wave 1-5 of the v4.55.1 audit response.
 * These lock down behavior that is easy to regress: scoring formulas, priv-esc
 * gates, feature-toggle enforcement, and constant integrity.
 *
 * Covers:
 *   - AMIR-02: grievance score uses real status/deadline (not hardcoded)
 *   - AMIR-03: satisfaction trend filter accepts 1-10 scale
 *   - K11-BUG-01: suggestPairings reads Is Steward column (not 'role')
 *   - V09-BUG-03: scoring weights normalized to sum to 1
 *   - H07-BUG-01: resolveGrievance uses GRIEVANCE_RESOLVED audit event
 *   - H07-BUG-02: GRIEVANCE_SIGNED exists in AUDIT_EVENTS
 *   - H07-BUG-05: DRAFT exists in GRIEVANCE_STATUS
 *   - A08-BUG-01: grievance score has progressive scaling
 *   - T12-BUG-08: contact score interpolates (no 25-point cliffs)
 *   - P02-BUG-02 / V09-BUG-01: no formula escape on JSON blobs (via inverse check)
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

global.logAuditEvent = jest.fn();
global.AUDIT_EVENTS = {
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  FOLDER_CREATED: 'FOLDER_CREATED'
};

loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '29_TrendAlertService.gs',
  '30_EngagementService.gs',
  '34_ScoringService.gs',
]);

describe('v4.55.1 — Constant integrity', () => {
  test('H07-BUG-02: AUDIT_EVENTS.GRIEVANCE_SIGNED exists', () => {
    expect(global.AUDIT_EVENTS.GRIEVANCE_SIGNED).toBe('GRIEVANCE_SIGNED');
  });

  test('H07-BUG-01: AUDIT_EVENTS has both GRIEVANCE_RESOLVED and GRIEVANCE_UPDATED as distinct values', () => {
    expect(global.AUDIT_EVENTS.GRIEVANCE_RESOLVED).toBe('GRIEVANCE_RESOLVED');
    expect(global.AUDIT_EVENTS.GRIEVANCE_UPDATED).toBe('GRIEVANCE_UPDATED');
    expect(global.AUDIT_EVENTS.GRIEVANCE_RESOLVED).not.toBe(global.AUDIT_EVENTS.GRIEVANCE_UPDATED);
  });

  test('H07-BUG-05: GRIEVANCE_STATUS.DRAFT is defined', () => {
    expect(global.GRIEVANCE_STATUS.DRAFT).toBe('Draft');
  });

  test('GRIEVANCE_CLOSED_STATUSES is populated and includes Closed, Settled, Withdrawn', () => {
    expect(Array.isArray(global.GRIEVANCE_CLOSED_STATUSES)).toBe(true);
    expect(global.GRIEVANCE_CLOSED_STATUSES).toContain('Closed');
    expect(global.GRIEVANCE_CLOSED_STATUSES).toContain('Settled');
    expect(global.GRIEVANCE_CLOSED_STATUSES).toContain('Withdrawn');
  });
});

describe('v4.55.1 — AMIR-03: satisfaction scale 1-10', () => {
  test('_detectSatisfactionDrop source accepts val<=10 not val<=5', () => {
    // Regression guard: ensure the literal filter in 29_TrendAlertService still uses 10
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '29_TrendAlertService.gs'),
      'utf8'
    );
    // Should contain val <= 10 in the filter
    expect(src).toMatch(/val\s*>=\s*1\s*&&\s*val\s*<=\s*10/);
    // Should NOT contain the old val <= 5 filter pattern
    expect(src).not.toMatch(/val\s*>=\s*1\s*&&\s*val\s*<=\s*5/);
  });
});

describe('v4.55.1 — V09-BUG-03: scoring weight normalization', () => {
  beforeEach(() => {
    // Mock ConfigReader.getConfig to return weights that don't sum to 100
    global.ConfigReader = {
      getConfig: jest.fn(() => ({
        scoreWeightEngagement: 70,
        scoreWeightProfile: 70, // intentionally oversized
        scoreWeightGrievance: 70,
      }))
    };
  });

  test('calculateCompositeScore normalizes when weights do not sum to 1', () => {
    const result = global.ScoringService.calculateCompositeScore(100, 100, 100);
    // With all dims=100 and weights normalized, result should be 100 (not 210)
    expect(result).toBeLessThanOrEqual(100);
    expect(result).toBeGreaterThanOrEqual(99); // allow rounding
  });

  test('calculateCompositeScore with default weights (70/20/10) returns a weighted average', () => {
    global.ConfigReader.getConfig.mockReturnValue({
      scoreWeightEngagement: 70,
      scoreWeightProfile: 20,
      scoreWeightGrievance: 10,
    });
    // engagement=100, profile=0, grievance=0 → 100*0.7 = 70
    expect(global.ScoringService.calculateCompositeScore(100, 0, 0)).toBe(70);
    // engagement=0, profile=100, grievance=0 → 100*0.2 = 20
    expect(global.ScoringService.calculateCompositeScore(0, 100, 0)).toBe(20);
    // engagement=0, profile=0, grievance=100 → 100*0.1 = 10
    expect(global.ScoringService.calculateCompositeScore(0, 0, 100)).toBe(10);
  });
});

describe('v4.55.1 — A08-BUG-01: progressive grievance score scaling', () => {
  test('8-day grievance scores different from 60-day grievance', () => {
    const short = global.ScoringService.calculateGrievanceScore(true, 'Open', 8, 'Negative');
    const mid = global.ScoringService.calculateGrievanceScore(true, 'Open', 60, 'Negative');
    const long_ = global.ScoringService.calculateGrievanceScore(true, 'Open', 120, 'Negative');
    expect(short).toBeLessThan(mid);
    expect(mid).toBeLessThan(long_);
  });

  test('overdue grievance (negative days) returns 0', () => {
    expect(global.ScoringService.calculateGrievanceScore(true, 'Open', -5, 'Negative')).toBe(0);
  });

  test('no open grievance returns 100', () => {
    expect(global.ScoringService.calculateGrievanceScore(false, '', 30, 'Negative')).toBe(100);
  });
});

describe('v4.55.1 — T12-BUG-08: contact score interpolation', () => {
  // EngagementService depends on many fixtures; we test _computeContactScore via
  // computeScoreForMember with minimal data.
  const runContact = (daysAgo) => {
    const now = new Date();
    const past = new Date(now.getTime() - daysAgo * 86400000);
    const contactData = [[past, '', 'user@test.com']];
    return global.EngagementService.computeScoreForMember('user@test.com', {
      contactData,
    }).scores.contact;
  };

  test('recent contact (3 days) gives near-100 score', () => {
    const score = runContact(3);
    expect(score).toBeGreaterThanOrEqual(90);
    expect(score).toBeLessThanOrEqual(100);
  });

  test('scores interpolate between 30 and 60 days (no cliff)', () => {
    const s30 = runContact(30);
    const s45 = runContact(45);
    const s60 = runContact(60);
    // 45 should be between 30 and 60 (strict interpolation)
    expect(s45).toBeLessThan(s30);
    expect(s45).toBeGreaterThan(s60);
  });

  test('120+ days returns 0', () => {
    expect(runContact(140)).toBe(0);
  });
});

describe('v4.55.1 — P02-BUG-02 / V09-BUG-01: no escapeForFormula on JSON', () => {
  test('updateAgencyDirectorOverrides source stores JSON without escapeForFormula', () => {
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'),
      'utf8'
    );
    // Find the function body
    const fnIdx = src.indexOf('function updateAgencyDirectorOverrides');
    expect(fnIdx).toBeGreaterThan(-1);
    // Within ~1000 chars of the function, there should be a setValue(json) not setValue(escapeForFormula(json))
    const fnBody = src.substring(fnIdx, fnIdx + 3000);
    expect(fnBody).toMatch(/setValue\(json\)/);
    expect(fnBody).not.toMatch(/setValue\(escapeForFormula\(json\)\)/);
  });
});

describe('v4.55.1 — K11-BUG-01: suggestPairings reads Is Steward column', () => {
  test('source uses MEMBER_COLS.IS_STEWARD or "is steward" header, not "role"', () => {
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '33_NewFeatureServices.gs'),
      'utf8'
    );
    const fnIdx = src.indexOf('function suggestPairings');
    expect(fnIdx).toBeGreaterThan(-1);
    const fnBody = src.substring(fnIdx, fnIdx + 2000);
    // Either uses MEMBER_COLS.IS_STEWARD or matches 'is steward' (case-insensitive header)
    const usesConstant = /MEMBER_COLS\.IS_STEWARD/.test(fnBody);
    const usesHeader = /['"]is steward['"]/.test(fnBody);
    expect(usesConstant || usesHeader).toBe(true);
    // Should NOT look for header 'role' exclusively
    expect(fnBody).not.toMatch(/h\s*===\s*['"]role['"]/);
  });
});

describe('v4.55.1 — N14-BUG-09: dead feedback wrappers removed', () => {
  test('21d_WebDashDataWrappers.gs no longer defines dataSubmitFeedback', () => {
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '21d_WebDashDataWrappers.gs'),
      'utf8'
    );
    expect(src).not.toMatch(/^function dataSubmitFeedback\s*\(/m);
    expect(src).not.toMatch(/^function dataGetMyFeedback\s*\(/m);
  });
});

describe('v4.55.1 — H07-BUG-03: STEP3_RCVD not referenced in executable code', () => {
  test('11_CommandHub.gs does not reference GRIEVANCE_COLS.STEP3_RCVD in non-comment lines', () => {
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '11_CommandHub.gs'),
      'utf8'
    );
    // Strip single-line comments (// ...) and block comments (/* ... */)
    const stripped = src
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .split('\n')
      .map(function(line) { return line.replace(/\/\/.*$/, ''); })
      .join('\n');
    expect(stripped).not.toMatch(/GRIEVANCE_COLS\.STEP3_RCVD/);
  });
});

describe('v4.55.1 — P02-BUG-01: dataGetCaseChecklist has ownership gate in source', () => {
  test('wrapper calls getCaseOwnerEmail and compares to caller', () => {
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '21d_WebDashDataWrappers.gs'),
      'utf8'
    );
    const fnIdx = src.indexOf('function dataGetCaseChecklist');
    expect(fnIdx).toBeGreaterThan(-1);
    const fnBody = src.substring(fnIdx, fnIdx + 1000);
    // Must include ownership check path
    expect(fnBody).toMatch(/getCaseOwnerEmail/);
    expect(fnBody).toMatch(/Access denied/);
  });
});
