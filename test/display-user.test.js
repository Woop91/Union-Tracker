/**
 * displayUser helper — sub-project A Zone 3
 *
 * Pure function that humanizes Access Log user values. Input may be a
 * SHA-256 hash (64 hex chars), an email, empty/null, or arbitrary text.
 * Output is always a short, readable string suitable for a table cell.
 *
 * Extracted from src/index.html the same way renderSafeText is.
 */

const fs = require('fs');
const path = require('path');

function loadDisplayUser() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'),
    'utf8'
  );
  var startIdx = src.indexOf('function displayUser(');
  if (startIdx === -1) throw new Error('displayUser not found in src/index.html');
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
  var wrapper = body + '\n; return displayUser;';
  return new Function(wrapper)();
}

const displayUser = loadDisplayUser();

describe('displayUser (sub-project A Zone 3)', () => {
  test('null returns em-dash placeholder', () => {
    expect(displayUser(null)).toBe('—');
  });

  test('empty string returns em-dash placeholder', () => {
    expect(displayUser('')).toBe('—');
  });

  test('undefined returns em-dash placeholder', () => {
    expect(displayUser(undefined)).toBe('—');
  });

  test('whitespace-only input returns em-dash placeholder', () => {
    expect(displayUser('   ')).toBe('—');
  });

  test('64-char hex hash returns abbreviated User #xxxxxxxx form', () => {
    const hash = 'ee048fdd37b45f43e3cc494bed87e7431ccc55d52f9eded4659d413678183373';
    expect(displayUser(hash)).toBe('User #ee048fdd');
  });

  test('email passes through unchanged (lowercased)', () => {
    expect(displayUser('Jacob.Sanders@example.org')).toBe('jacob.sanders@example.org');
  });

  test('short plain-text identifier passes through unchanged', () => {
    expect(displayUser('steward-alpha')).toBe('steward-alpha');
  });
});
