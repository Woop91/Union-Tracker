/**
 * Tests for 00_DataAccess.gs
 *
 * Covers TIME_CONSTANTS, deadline calculation utilities,
 * and the DataAccess layer interface.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs', '00_DataAccess.gs']);

// ============================================================================
// TIME_CONSTANTS
// ============================================================================

describe('TIME_CONSTANTS', () => {
  test('MS_PER_DAY is correct', () => {
    expect(TIME_CONSTANTS.MS_PER_DAY).toBe(86400000);
  });

  test('MS_PER_WEEK is 7 * MS_PER_DAY', () => {
    expect(TIME_CONSTANTS.MS_PER_WEEK).toBe(7 * TIME_CONSTANTS.MS_PER_DAY);
  });
});

describe('TIME_CONSTANTS.DEADLINE_DAYS', () => {
  const dd = TIME_CONSTANTS.DEADLINE_DAYS;

  test('FILING is 21 days', () => { expect(dd.FILING).toBe(21); });
  test('STEP1_RESPONSE is 7 days', () => { expect(dd.STEP1_RESPONSE).toBe(7); });
  test('STEP2_APPEAL is 7 days', () => { expect(dd.STEP2_APPEAL).toBe(7); });
  test('STEP2_RESPONSE is 14 days', () => { expect(dd.STEP2_RESPONSE).toBe(14); });
  test('STEP3_APPEAL is 10 days', () => { expect(dd.STEP3_APPEAL).toBe(10); });
  test('WARNING_THRESHOLD is 5', () => { expect(dd.WARNING_THRESHOLD).toBe(5); });
  test('CRITICAL_THRESHOLD is 2', () => { expect(dd.CRITICAL_THRESHOLD).toBe(2); });
});

describe('TIME_CONSTANTS.REMINDER_DAYS', () => {
  test('reminder days decrease correctly', () => {
    const rd = TIME_CONSTANTS.REMINDER_DAYS;
    expect(rd.FIRST).toBe(7);
    expect(rd.SECOND).toBe(3);
    expect(rd.FINAL).toBe(1);
  });
});

// ============================================================================
// calculateDeadline
// ============================================================================

describe('calculateDeadline', () => {
  test('adds correct number of days', () => {
    const start = new Date('2026-01-01T00:00:00Z');
    const result = calculateDeadline(start, 21);
    expect(result.getTime()).toBe(start.getTime() + 21 * TIME_CONSTANTS.MS_PER_DAY);
  });

  test('defaults to current date if startDate is invalid', () => {
    const before = Date.now();
    const result = calculateDeadline(null, 7);
    const after = Date.now();
    const sevenDaysMs = 7 * TIME_CONSTANTS.MS_PER_DAY;
    expect(result.getTime()).toBeGreaterThanOrEqual(before + sevenDaysMs);
    expect(result.getTime()).toBeLessThanOrEqual(after + sevenDaysMs);
  });

  test('handles zero days', () => {
    const start = new Date('2026-06-15T12:00:00Z');
    expect(calculateDeadline(start, 0).getTime()).toBe(start.getTime());
  });
});

// ============================================================================
// daysBetween
// ============================================================================

describe('daysBetween', () => {
  test('calculates positive difference', () => {
    expect(daysBetween(new Date('2026-01-01'), new Date('2026-01-08'))).toBe(7);
  });

  test('calculates negative difference', () => {
    expect(daysBetween(new Date('2026-01-08'), new Date('2026-01-01'))).toBe(-7);
  });

  test('returns 0 for null inputs', () => {
    expect(daysBetween(null, new Date())).toBe(0);
  });

  test('accepts string dates', () => {
    expect(daysBetween('2026-01-01', '2026-01-11')).toBe(10);
  });
});

// ============================================================================
// getDeadlineUrgency
// ============================================================================

describe('getDeadlineUrgency', () => {
  test('returns overdue for negative days', () => {
    expect(getDeadlineUrgency(-1)).toBe('overdue');
  });

  test('returns critical for 0-2 days', () => {
    expect(getDeadlineUrgency(0)).toBe('critical');
    expect(getDeadlineUrgency(2)).toBe('critical');
  });

  test('returns warning for 3-5 days', () => {
    expect(getDeadlineUrgency(3)).toBe('warning');
    expect(getDeadlineUrgency(5)).toBe('warning');
  });

  test('returns normal for 6+ days', () => {
    expect(getDeadlineUrgency(6)).toBe('normal');
    expect(getDeadlineUrgency(100)).toBe('normal');
  });
});

// ============================================================================
// DataAccess
// ============================================================================

describe('DataAccess', () => {
  test('has required methods', () => {
    expect(typeof DataAccess.getSpreadsheet).toBe('function');
    expect(typeof DataAccess.getSheet).toBe('function');
    expect(typeof DataAccess.getAllData).toBe('function');
    expect(typeof DataAccess.findRow).toBe('function');
    expect(typeof DataAccess.appendRow).toBe('function');
    expect(typeof DataAccess.getMemberById).toBe('function');
    expect(typeof DataAccess.getGrievanceById).toBe('function');
  });
});
