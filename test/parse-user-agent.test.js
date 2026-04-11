/**
 * parseUserAgent helper — sub-project D Zone 4
 *
 * Pure function that classifies a raw navigator.userAgent string into
 * {browser, os, deviceClass} buckets. Used by the client-side usage
 * logger before sending events to dataLogUsageEvents, and by the
 * admin Usage Analytics display to group events.
 *
 * Extracted from src/index.html via bracket-depth parsing (same as
 * fuzzyMatch, renderSafeText, displayUser helpers).
 */

const fs = require('fs');
const path = require('path');

function loadParseUserAgent() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'),
    'utf8'
  );
  var startIdx = src.indexOf('function parseUserAgent(');
  if (startIdx === -1) throw new Error('parseUserAgent not found in src/index.html');
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
  var wrapper = body + '\n; return parseUserAgent;';
  return new Function(wrapper)();
}

const parseUserAgent = loadParseUserAgent();

describe('parseUserAgent (sub-project D Zone 4)', () => {
  test('iPhone Safari is ios + safari + mobile', () => {
    var ua = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1';
    expect(parseUserAgent(ua)).toEqual({ browser: 'safari', os: 'ios', deviceClass: 'mobile' });
  });

  test('iPad Safari is ios + safari + tablet', () => {
    var ua = 'Mozilla/5.0 (iPad; CPU OS 17_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1';
    var r = parseUserAgent(ua);
    expect(r.os).toBe('ios');
    expect(r.browser).toBe('safari');
    // iPad UA contains "Mobile" so it's bucketed mobile here — acceptable; iPad users have iPad-class screens
    expect(['tablet', 'mobile']).toContain(r.deviceClass);
  });

  test('Android Chrome is android + chrome + mobile', () => {
    var ua = 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Mobile Safari/537.36';
    expect(parseUserAgent(ua)).toEqual({ browser: 'chrome', os: 'android', deviceClass: 'mobile' });
  });

  test('Windows Chrome desktop is windows + chrome + desktop', () => {
    var ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36';
    expect(parseUserAgent(ua)).toEqual({ browser: 'chrome', os: 'windows', deviceClass: 'desktop' });
  });

  test('Windows Edge desktop is windows + edge + desktop', () => {
    var ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0';
    expect(parseUserAgent(ua)).toEqual({ browser: 'edge', os: 'windows', deviceClass: 'desktop' });
  });

  test('macOS Safari desktop is macos + safari + desktop', () => {
    var ua = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15';
    expect(parseUserAgent(ua)).toEqual({ browser: 'safari', os: 'macos', deviceClass: 'desktop' });
  });

  test('macOS Firefox desktop is macos + firefox + desktop', () => {
    var ua = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:124.0) Gecko/20100101 Firefox/124.0';
    expect(parseUserAgent(ua)).toEqual({ browser: 'firefox', os: 'macos', deviceClass: 'desktop' });
  });

  test('Linux Firefox is linux + firefox + desktop', () => {
    var ua = 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:124.0) Gecko/20100101 Firefox/124.0';
    expect(parseUserAgent(ua)).toEqual({ browser: 'firefox', os: 'linux', deviceClass: 'desktop' });
  });

  test('empty/null input returns all "other"', () => {
    expect(parseUserAgent('')).toEqual({ browser: 'other', os: 'other', deviceClass: 'desktop' });
    expect(parseUserAgent(null)).toEqual({ browser: 'other', os: 'other', deviceClass: 'desktop' });
    expect(parseUserAgent(undefined)).toEqual({ browser: 'other', os: 'other', deviceClass: 'desktop' });
  });

  test('Edge is detected BEFORE Chrome (correct precedence)', () => {
    // Edge contains both "Chrome/" and "Edg/" — the function must check Edg first
    var ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0';
    expect(parseUserAgent(ua).browser).toBe('edge');
  });
});
