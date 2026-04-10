/**
 * Director Overrides — Tests for validation logic.
 */

var VALID_POS_IDS = [
  'eohhs_secretary', 'massability_commissioner', 'dds_commissioner',
  'dds_regional_director_west', 'dds_regional_director_metro', 'dds_regional_director_east',
  'dds_asst_commissioner', 'cl_director', 'vr_director',
  'dds_dep_commissioner', 'massability_dep_commissioner',
];

function validateOverrides(overrides) {
  if (!overrides || typeof overrides !== 'object') return { valid: false, error: 'Invalid overrides.' };
  for (var posId in overrides) {
    if (!Object.prototype.hasOwnProperty.call(overrides, posId)) continue;
    if (VALID_POS_IDS.indexOf(posId) === -1) return { valid: false, error: 'Unknown position: ' + posId };
    var entry = overrides[posId];
    if (entry.name !== undefined) {
      if (typeof entry.name !== 'string' || entry.name.trim().length === 0) return { valid: false, error: 'Name must be non-empty string.' };
      if (entry.name.length > 100) return { valid: false, error: 'Name too long (max 100).' };
      if (/<[^>]+>/.test(entry.name)) return { valid: false, error: 'HTML not allowed in name.' };
    }
    if (entry.removedPayHx !== undefined) {
      if (!Array.isArray(entry.removedPayHx)) return { valid: false, error: 'removedPayHx must be array.' };
      for (var i = 0; i < entry.removedPayHx.length; i++) {
        if (!/^\d{4}$/.test(String(entry.removedPayHx[i]))) return { valid: false, error: 'removedPayHx entries must be 4-digit years.' };
      }
    }
  }
  return { valid: true };
}

describe('Director overrides validation', () => {
  test('valid name override passes', () => {
    expect(validateOverrides({ dds_commissioner: { name: 'Jane Doe' } }).valid).toBe(true);
  });
  test('empty name rejected', () => {
    expect(validateOverrides({ dds_commissioner: { name: '  ' } }).valid).toBe(false);
  });
  test('HTML in name rejected', () => {
    expect(validateOverrides({ dds_commissioner: { name: '<script>alert(1)</script>' } }).valid).toBe(false);
  });
  test('name over 100 chars rejected', () => {
    expect(validateOverrides({ dds_commissioner: { name: 'A'.repeat(101) } }).valid).toBe(false);
  });
  test('unknown position ID rejected', () => {
    expect(validateOverrides({ fake_position: { name: 'Test' } }).valid).toBe(false);
  });
  test('valid removedPayHx passes', () => {
    expect(validateOverrides({ massability_commissioner: { removedPayHx: ['2017', '2018'] } }).valid).toBe(true);
  });
  test('non-array removedPayHx rejected', () => {
    expect(validateOverrides({ massability_commissioner: { removedPayHx: '2017' } }).valid).toBe(false);
  });
  test('invalid year format rejected', () => {
    expect(validateOverrides({ massability_commissioner: { removedPayHx: ['2017-01-15'] } }).valid).toBe(false);
  });
  test('empty overrides passes', () => {
    expect(validateOverrides({}).valid).toBe(true);
  });
  test('null rejected', () => {
    expect(validateOverrides(null).valid).toBe(false);
  });
});
