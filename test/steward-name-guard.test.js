/**
 * steward-name guard — sub-project A Zone 2
 *
 * Tests the _looksLikeStewardName helper that rejects non-name values
 * (empty, Date serials, ISO dates, weekday-prefixed strings, overlong
 * junk) before they become grouping keys in the Avg Resolution chart.
 *
 * The helper is a pure function defined inside the 21_WebDashDataService.gs
 * IIFE. We extract it via the same pattern as fuzzyMatch tests.
 */

const fs = require('fs');
const path = require('path');

function loadStewardNameGuard() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'),
    'utf8'
  );
  var startIdx = src.indexOf('function _looksLikeStewardName(');
  if (startIdx === -1) throw new Error('_looksLikeStewardName not found in 21_WebDashDataService.gs');
  var depth = 0;
  var inFn = false;
  var endIdx = startIdx;
  for (var i = startIdx; i < src.length; i++) {
    var ch = src.charAt(i);
    if (ch === '{') { depth++; inFn = true; continue; }
    if (ch === '}') {
      depth--;
      if (inFn && depth === 0) { endIdx = i + 1; break; }
    }
  }
  var body = src.substring(startIdx, endIdx);
  var wrapper = body + '\n; return _looksLikeStewardName;';
  return new Function(wrapper)();
}

const _looksLikeStewardName = loadStewardNameGuard();

describe('_looksLikeStewardName (sub-project A Zone 2)', () => {
  test('empty string is rejected', () => {
    expect(_looksLikeStewardName('')).toBe(false);
  });

  test('null is rejected', () => {
    expect(_looksLikeStewardName(null)).toBe(false);
  });

  test('numeric-only serial string is rejected', () => {
    expect(_looksLikeStewardName('46076.51939715278')).toBe(false);
  });

  test('negative numeric serial is rejected', () => {
    expect(_looksLikeStewardName('-42.5')).toBe(false);
  });

  test('scientific notation serial is rejected', () => {
    expect(_looksLikeStewardName('4.607651939715278e+4')).toBe(false);
  });

  test('scientific notation with negative exponent is rejected', () => {
    expect(_looksLikeStewardName('1.5e-3')).toBe(false);
  });

  test('scientific notation uppercase E is rejected', () => {
    expect(_looksLikeStewardName('1E10')).toBe(false);
  });

  test('ISO date string is rejected', () => {
    expect(_looksLikeStewardName('2026-03-18t14:22:00.000z')).toBe(false);
  });

  test('weekday-prefixed date string is rejected', () => {
    expect(_looksLikeStewardName('wed mar 18 2026 10:25:00 gmt-0500 (est)')).toBe(false);
  });

  test('overlong string (> 120 chars) is rejected', () => {
    var longJunk = new Array(130).join('x');
    expect(_looksLikeStewardName(longJunk)).toBe(false);
  });

  test('legitimate lowercase name is accepted', () => {
    expect(_looksLikeStewardName('jacob sanders')).toBe(true);
  });
});
