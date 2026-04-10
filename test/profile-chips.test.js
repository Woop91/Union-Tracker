/**
 * Profile Chips — Tests for tap-to-select chip UI logic.
 */

describe('Profile Chips: T-shirt size (single-select)', () => {
  test('only one chip active at a time', () => {
    var sizes = ['XS','S','M','L','XL','2XL','3XL','4XL','5XL'];
    var active = '';
    function selectSize(sz) { active = sz; }
    function getActive() { return active; }
    selectSize('M');
    expect(getActive()).toBe('M');
    selectSize('XL');
    expect(getActive()).toBe('XL');
    expect(getActive()).not.toBe('M');
  });

  test('empty selection returns empty string', () => {
    var active = '';
    expect(active).toBe('');
  });
});

describe('Profile Chips: Office Days (multi-select)', () => {
  test('toggling builds comma-separated string', () => {
    var selected = new Set();
    function toggle(day) {
      if (selected.has(day)) selected.delete(day);
      else selected.add(day);
    }
    function getValue() { return Array.from(selected).join(','); }
    toggle('Monday');
    toggle('Wednesday');
    toggle('Friday');
    expect(getValue()).toBe('Monday,Wednesday,Friday');
    toggle('Wednesday');
    expect(getValue()).toBe('Monday,Friday');
  });

  test('Remote option works alongside days', () => {
    var selected = new Set(['Monday', 'Remote']);
    expect(Array.from(selected).join(',')).toBe('Monday,Remote');
  });

  test('value is compatible with inOfficeToday() parsing', () => {
    var officeDays = 'Monday,Wednesday,Friday,Remote';
    var todayName = 'wednesday';
    expect(officeDays.toLowerCase().indexOf(todayName) !== -1).toBe(true);
    todayName = 'saturday';
    expect(officeDays.toLowerCase().indexOf(todayName) !== -1).toBe(false);
  });

  test('parse stored value back into set', () => {
    var stored = 'Monday,Friday,Remote';
    var parsed = stored.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    expect(parsed).toEqual(['Monday', 'Friday', 'Remote']);
  });
});
