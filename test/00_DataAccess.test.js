/**
 * Tests for 00_DataAccess.gs
 *
 * Covers TIME_CONSTANTS and deadline calculation utilities.
 * Note: DataAccess namespace was removed (zero callers — refactored away).
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
  test('STEP1_RESPONSE is 30 days', () => { expect(dd.STEP1_RESPONSE).toBe(30); });
  test('STEP2_APPEAL is 10 days', () => { expect(dd.STEP2_APPEAL).toBe(10); });
  test('STEP2_RESPONSE is 30 days', () => { expect(dd.STEP2_RESPONSE).toBe(30); });
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


