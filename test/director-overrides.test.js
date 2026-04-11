/**
 * Director Overrides — Tests the real updateAgencyDirectorOverrides function
 * (loaded from src/21_WebDashDataService.gs), not a test-local copy.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs'
]);

// updateAgencyDirectorOverrides is a top-level function in the real file.
// In the mocked environment its Config-sheet write either succeeds (when
// we provide a mock sheet) or fails with a benign "Config sheet unavailable"
// message. Either way, validation errors short-circuit BEFORE the sheet
// write, so the validation assertions below don't depend on mock setup.

describe('Director overrides validation (real production function)', () => {
  test('valid name override passes validation', () => {
    var result = updateAgencyDirectorOverrides({ dds_commissioner: { name: 'Jane Doe' } });
    // Either success (mock sheet present) or "Config sheet unavailable" —
    // validation passed either way, never "Name must be".
    expect(result.success === true || /config sheet/i.test(result.message || '')).toBe(true);
  });

  test('empty name is rejected with specific error', () => {
    var result = updateAgencyDirectorOverrides({ dds_commissioner: { name: '  ' } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/name/i);
  });

  test('HTML in name is rejected', () => {
    var result = updateAgencyDirectorOverrides({ dds_commissioner: { name: '<script>alert(1)</script>' } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/html/i);
  });

  test('name over 100 chars is rejected', () => {
    var result = updateAgencyDirectorOverrides({ dds_commissioner: { name: 'A'.repeat(101) } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/too long|100/i);
  });

  test('unknown position ID is rejected', () => {
    var result = updateAgencyDirectorOverrides({ fake_position: { name: 'Test' } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/unknown position/i);
  });

  test('valid 4-digit removedPayHx passes validation', () => {
    var result = updateAgencyDirectorOverrides({ massability_commissioner: { removedPayHx: ['2017', '2018'] } });
    expect(result.success === true || /config sheet/i.test(result.message || '')).toBe(true);
  });

  test('non-array removedPayHx is rejected', () => {
    var result = updateAgencyDirectorOverrides({ massability_commissioner: { removedPayHx: '2017' } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/array/i);
  });

  test('invalid year format in removedPayHx is rejected', () => {
    var result = updateAgencyDirectorOverrides({ massability_commissioner: { removedPayHx: ['2017-01-15'] } });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/4-digit/i);
  });

  test('empty overrides object passes validation', () => {
    var result = updateAgencyDirectorOverrides({});
    expect(result.success === true || /config sheet/i.test(result.message || '')).toBe(true);
  });

  test('null input is rejected', () => {
    var result = updateAgencyDirectorOverrides(null);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/invalid/i);
  });

  // v4.55.2 Wave 29 / Auditor-Beta B-19: additional negative-path coverage.
  // The existing suite covered per-field validation failures but not input-
  // shape edge cases, nested-null handling, or the interaction between
  // multiple invalid fields in a single payload. These fill that gap.

  test('B-19: undefined input is rejected (treated as missing)', () => {
    var result = updateAgencyDirectorOverrides(undefined);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/invalid/i);
  });

  test('B-19: string input is rejected (not an object)', () => {
    var result = updateAgencyDirectorOverrides('hello');
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/invalid/i);
  });

  test('B-19: number input is rejected (not an object)', () => {
    var result = updateAgencyDirectorOverrides(42);
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/invalid/i);
  });

  test('B-19: non-string name type (number) is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: 12345 }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/name/i);
  });

  test('B-19: non-string name type (boolean) is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: true }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/name/i);
  });

  test('B-19: non-string name type (object) is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: { first: 'Jane', last: 'Doe' } }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/name/i);
  });

  test('B-19: mixed valid + invalid rejects on first invalid entry', () => {
    // One valid entry then one invalid — the function must reject the whole
    // payload before writing (atomicity). A silent partial-write would be
    // worse than the outright rejection.
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: 'Jane Doe' },
      fake_position: { name: 'Whatever' }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/unknown position/i);
  });

  test('B-19: valid + html-in-name rejects whole payload', () => {
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: 'Jane Doe' },
      massability_commissioner: { name: '<img src=x>' }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/html/i);
  });

  test('B-19: 3-digit year in removedPayHx is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      massability_commissioner: { removedPayHx: ['201'] }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/4-digit/i);
  });

  test('B-19: 5-digit year in removedPayHx is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      massability_commissioner: { removedPayHx: ['20170'] }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/4-digit/i);
  });

  test('B-19: non-numeric year string in removedPayHx is rejected', () => {
    var result = updateAgencyDirectorOverrides({
      massability_commissioner: { removedPayHx: ['abcd'] }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/4-digit/i);
  });

  test('B-19: numeric (not string) year coerces correctly', () => {
    // 2017 as a number should be coerced to "2017" via String() and pass.
    var result = updateAgencyDirectorOverrides({
      massability_commissioner: { removedPayHx: [2017, 2018] }
    });
    // Validation should pass — source does String(entry) before regex match.
    expect(
      result.success === true || /config sheet/i.test(result.message || '')
    ).toBe(true);
  });

  test('B-19: one valid year + one invalid year rejects whole payload', () => {
    var result = updateAgencyDirectorOverrides({
      massability_commissioner: { removedPayHx: ['2017', 'bad'] }
    });
    expect(result.success).toBe(false);
    expect(result.message).toMatch(/4-digit/i);
  });

  test('B-19: extra unknown key on entry is silently allowed', () => {
    // The source only validates 'name' and 'removedPayHx'. Extra keys should
    // not block the update — they're simply not serialized into the final
    // override (the whole overrides object IS serialized, but extra keys
    // are harmless no-ops on the read side). This test pins that contract.
    var result = updateAgencyDirectorOverrides({
      dds_commissioner: { name: 'Jane Doe', unknownExtraField: 'xyz' }
    });
    expect(
      result.success === true || /config sheet/i.test(result.message || '')
    ).toBe(true);
  });
});
