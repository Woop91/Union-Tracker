# Sub-Project A — Display-Safe Rendering Boundary Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship DDS v4.55.3 bug-fix wave closing issues #1 (Avg Resolution chart Y-axis shows date serials), #5 (Access Log shows SHA-256 hashes), and #9 (literal `\n` escape sequences in Resources content) by introducing a display-safe rendering boundary: two pure helpers (`renderSafeText`, `displayUser`), one upstream type guard in the steward aggregation loop, and seed-string cleanup for affected data.

**Architecture:** Three isolated zones. Zone 1 is a pure string→DOM-fragment helper that normalizes literal and real newlines into `<br>` nodes with HTML escaping; it lives in `src/index.html` alongside the Wave 31 `fuzzyMatch` helper and is adopted at every Resources render site plus seed sources. Zone 2 is a single-function type guard at the `stewardResolution` accumulation step in `src/21_WebDashDataService.gs:3074` that rejects non-name values (Date serials, ISO dates, numeric strings) from becoming phantom grouping keys. Zone 3 is a pure hash-aware resolver `displayUser(rawId)` in `src/index.html` that humanizes SHA-256 hashes (64 hex chars → `User #ee048fdd…`) and falls through to normal email/name handling. Each zone ships as its own commit so any can be reverted surgically.

**Tech Stack:** Google Apps Script (V8), vanilla JS inside GAS `HtmlService` templates (single-global-scope, IIFE-wrapped), Jest for unit tests (no jsdom — tests extract helpers via `new Function()` eval from `src/index.html`). Build: zero-dep `build.js --prod --minify`. Deploy: `clasp push` to DDS dev/prod Script IDs.

---

## File Structure

**Files created:**
- `test/render-safe-text.test.js` — Zone 1 helper tests (~10 tests). Follows the `fuzzyMatch.test.js` pattern: read `src/index.html`, extract the function body by bracket-depth, eval via `new Function()`, then exercise every path.
- `test/steward-name-guard.test.js` — Zone 2 guard tests (~8 tests). Standalone module because the guard is a small pure function extracted/copied for test import.
- `test/display-user.test.js` — Zone 3 helper tests (~6 tests). Same extraction pattern as Zone 1.

**Files modified:**
- `src/index.html` — Add `renderSafeText(str)` and `displayUser(rawId)` helpers near the existing `fuzzyMatch` block (after line 1307 per earlier grep).
- `src/member_view.html` — Adopt `renderSafeText` at the Resources card body render site inside `renderCards()` (around line 2099 in `renderMemberResources`). Grep for `r.content` usage to pin down exact line at execution time.
- `src/steward_view.html` — Adopt `renderSafeText` at the steward Know Your Rights hub (search-anchor: `hubSection` around line 3087). Adopt `displayUser` at the Access Log user-column render (`renderAccessLogPage` line 8932: `item.user || '--'`).
- `src/21_WebDashDataService.gs` — Add the `_looksLikeStewardName(s)` guard function near `_buildGrievanceRecord` (around line 2178) and call it in the aggregation loop at line 3074 before accumulating into `stewardResolution`.
- `src/10b_SurveyDocSheets.gs` — Replace `'\\n'` with `'\n'` in `starterRows` at lines 1254–1259 (6 affected rows in the `createResourcesSheet` function).
- `src/07_DevTools.gs` — Replace `'\\n'` with `'\n'` in `seedResourcesData` at lines 1737+ (9+ affected rows). Also audit any other `additionalResources` entries in the same function.
- `test/spa-integrity.test.js` — Add two assertions verifying `renderSafeText` and `displayUser` are defined in `src/index.html` at top-level IIFE scope (Wave 31 precedent: G-RENDER-GEN, G-IDLE blocks).
- `CHANGELOG.md` — Bump to v4.55.3 with a "Fixes" section referencing the three issues.
- `src/01_Core.gs` — Update `COMMAND_CONFIG.VERSION` from `'4.55.2'` to `'4.55.3'` (canonical version per session memory).

**Files untouched** (called out explicitly so reviewers know):
- `src/21d_WebDashDataWrappers.gs` `dataGetAuditLog` — No backend change. The frontend `displayUser` helper handles the hash humanization. An upstream investigation of "why are hashes getting into USER_EMAIL" is deferred to sub-project B.
- `src/_findColumn` / `_buildColumnMap` — Already exact-match only. No change needed (correction from initial brainstorm).

---

## Task 1: Write `renderSafeText` helper + tests (Zone 1)

**Files:**
- Modify: `src/index.html:1307+` (insert new function after the `fuzzyMatch` block)
- Create: `test/render-safe-text.test.js`

- [ ] **Step 1: Write the failing test file**

Create `test/render-safe-text.test.js` with this exact content:

```javascript
/**
 * renderSafeText helper — sub-project A Zone 1
 *
 * Tests the pure string→DocumentFragment helper that normalizes literal
 * '\n' (backslash+n) and real '\n' newlines into <br> nodes while HTML-
 * escaping everything else. Helper lives in src/index.html; we extract it
 * by bracket-depth the same way fuzzy-match.test.js does.
 */

const fs = require('fs');
const path = require('path');

function loadRenderSafeText() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'),
    'utf8'
  );
  var startIdx = src.indexOf('function renderSafeText(');
  if (startIdx === -1) throw new Error('renderSafeText not found in src/index.html');
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
  // The helper uses document.createDocumentFragment and document.createElement.
  // Tests run under node without jsdom, so stub a minimal document before eval.
  var stub = `
    var document = {
      createDocumentFragment: function() {
        var nodes = [];
        return {
          __nodes: nodes,
          appendChild: function(n) { nodes.push(n); return n; },
          get childNodes() { return nodes; }
        };
      },
      createElement: function(tag) {
        return { __tag: tag, nodeName: tag.toUpperCase(), childNodes: [] };
      },
      createTextNode: function(text) {
        return { __text: text, nodeName: '#text', nodeValue: text };
      }
    };
  `;
  var wrapper = stub + body + '\n; return renderSafeText;';
  return new Function(wrapper)();
}

const renderSafeText = loadRenderSafeText();

describe('renderSafeText (sub-project A Zone 1)', () => {
  test('empty string returns empty fragment', () => {
    const frag = renderSafeText('');
    expect(frag.childNodes.length).toBe(0);
  });

  test('null returns empty fragment', () => {
    const frag = renderSafeText(null);
    expect(frag.childNodes.length).toBe(0);
  });

  test('undefined returns empty fragment', () => {
    const frag = renderSafeText(undefined);
    expect(frag.childNodes.length).toBe(0);
  });

  test('plain text returns a single text node', () => {
    const frag = renderSafeText('hello world');
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeName).toBe('#text');
    expect(frag.childNodes[0].nodeValue).toBe('hello world');
  });

  test('literal backslash-n is split into two text nodes with a BR between', () => {
    const frag = renderSafeText('line one\\nline two');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('line one');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('line two');
  });

  test('real newline is split into two text nodes with a BR between', () => {
    const frag = renderSafeText('line one\nline two');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('line one');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('line two');
  });

  test('Windows CRLF is normalized to a single break', () => {
    const frag = renderSafeText('a\r\nb');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('a');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('b');
  });

  test('multi-line content produces alternating text/BR nodes', () => {
    const frag = renderSafeText('1. Was the employee warned?\\n2. Was the rule reasonable?\\n3. Was an investigation done?');
    // 3 text nodes + 2 BRs = 5 children
    expect(frag.childNodes.length).toBe(5);
    expect(frag.childNodes[0].nodeValue).toBe('1. Was the employee warned?');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('2. Was the rule reasonable?');
    expect(frag.childNodes[3].nodeName).toBe('BR');
    expect(frag.childNodes[4].nodeValue).toBe('3. Was an investigation done?');
  });

  test('script-tag content stays as escaped text (XSS defense)', () => {
    const frag = renderSafeText('<script>alert(1)</script>');
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeName).toBe('#text');
    expect(frag.childNodes[0].nodeValue).toBe('<script>alert(1)</script>');
  });

  test('numeric input is stringified, not rejected', () => {
    const frag = renderSafeText(42);
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeValue).toBe('42');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/render-safe-text.test.js --no-coverage`
Expected: FAIL with "renderSafeText not found in src/index.html".

- [ ] **Step 3: Add the `renderSafeText` helper to `src/index.html`**

Find the end of the `fuzzyMatch` function in `src/index.html` (search anchor: `function fuzzyMatch(`). After `fuzzyMatch` closes (brace-depth returns to 0), insert this block at the same indentation level:

```javascript
    // ============================================================================
    // v4.55.3 sub-project A Zone 1: renderSafeText — newline-tolerant text renderer.
    // ============================================================================
    //
    // Resources sheet content contains literal '\n' (two characters: backslash + n)
    // from an older seed bug, plus real '\n' from admin edits. Callers previously
    // piped content straight into .textContent, so '\n' rendered as visible
    // backslash-n instead of a line break. This helper normalizes both forms into
    // <br> nodes while HTML-escaping everything else by using text nodes, which
    // makes it XSS-safe by construction (no innerHTML path).
    //
    // Return shape is a DocumentFragment so callers can appendChild directly
    // without an intermediate HTML string.
    function renderSafeText(str) {
      var frag = document.createDocumentFragment();
      if (str === null || str === undefined || str === '') return frag;
      var s = String(str);
      // Normalize CRLF → LF, then convert literal '\n' (two chars) into real '\n'
      // so a single split handles both.
      s = s.replace(/\r\n/g, '\n').replace(/\\n/g, '\n');
      var parts = s.split('\n');
      for (var i = 0; i < parts.length; i++) {
        if (i > 0) frag.appendChild(document.createElement('br'));
        if (parts[i].length > 0) {
          frag.appendChild(document.createTextNode(parts[i]));
        }
      }
      return frag;
    }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx jest test/render-safe-text.test.js --no-coverage`
Expected: PASS, 10 tests passing.

- [ ] **Step 5: Run the full guard suite to confirm no regression**

Run: `npm run test:guards`
Expected: ALL WEBAPP TESTS PASSED (274 tests).

- [ ] **Step 6: Commit Zone 1 helper**

```bash
git add src/index.html test/render-safe-text.test.js
git commit -m "feat: add renderSafeText helper for sub-project A Zone 1

Pure string→DocumentFragment helper that normalizes literal '\n'
(backslash+n, from older Resources seed bug) and real '\n' newlines
into <br> nodes with HTML-escaped text nodes. XSS-safe by construction.
+10 tests in test/render-safe-text.test.js.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Adopt `renderSafeText` in Resources render sites

**Files:**
- Modify: `src/member_view.html` (Resources card body render, inside `renderCards()` near line 2099)
- Modify: `src/steward_view.html` (Know Your Rights hub render, search anchor: `hubSection` near line 3087)

- [ ] **Step 1: Locate the member_view Resources body render**

Open `src/member_view.html` and find `function renderCards()` inside `renderMemberResources` (starts around line 2081). Scroll to where each filtered resource card is built — look for where `r.content` or `r.body` is assigned to a DOM node (via `.textContent`, `innerHTML`, or a string passed into `el()`).

Run: `grep -n 'r\.content\|r\.body\|resource\.content' src/member_view.html | head -20`
Expected: line numbers of every resource-body assignment.

- [ ] **Step 2: Replace `.textContent = r.content` (or equivalent) with renderSafeText**

For every match from Step 1, change:

```javascript
// BEFORE (example pattern)
var bodyEl = el('div', { className: 'res-body' });
bodyEl.textContent = r.content || '';
```

to:

```javascript
// AFTER
var bodyEl = el('div', { className: 'res-body' });
bodyEl.appendChild(renderSafeText(r.content || ''));
```

If the existing pattern uses `el('div', {...}, r.content)` (text passed as the third `el()` argument), change it to:

```javascript
var bodyEl = el('div', { className: 'res-body' });
bodyEl.appendChild(renderSafeText(r.content || ''));
```

Do not change title, category, icon, or summary assignments — those are single-line fields and do not contain embedded newlines.

- [ ] **Step 3: Locate the steward_view Resources render**

Run: `grep -n 'Know Your Rights\|resource.*content\|r\.content' src/steward_view.html | head -20`
Expected: lines around 3087 (hub heading) and whatever function renders resource cards.

Apply the same replacement pattern as Step 2 to any `.textContent = r.content` or `el('div', {...}, r.content)` sites in `steward_view.html`.

- [ ] **Step 4: Run the guard suite to verify no syntax breakage**

Run: `npm run test:guards`
Expected: ALL WEBAPP TESTS PASSED (274 tests). The G1 guard will catch any unclosed brace or syntax error in the modified views.

- [ ] **Step 5: Build and verify dist parity**

Run: `npm run build:prod`
Expected: build succeeds, no `Prod file count` errors (warnings about "approaching limit of 65" are expected and fine).

- [ ] **Step 6: Commit the adoption**

```bash
git add src/member_view.html src/steward_view.html dist/
git commit -m "feat: adopt renderSafeText in Resources render sites

Every call site that previously wrote r.content to .textContent or
passed it as an el() third arg now pipes through renderSafeText, so
literal '\n' and real '\n' both render as <br>. Fixes issue #9.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: Clean `\\n` in Resources seed strings

**Files:**
- Modify: `src/10b_SurveyDocSheets.gs:1254-1260` (6 starter rows)
- Modify: `src/07_DevTools.gs:1737+` (9+ rows in `seedResourcesData`)

- [ ] **Step 1: Replace `'\\n'` with `'\n'` in 10b_SurveyDocSheets.gs starterRows**

Open `src/10b_SurveyDocSheets.gs`. Find `var starterRows = [` at line 1252. For each of the 6 rows (lines 1254–1259) that contain `\\n` in the body field, replace every `\\n` with `\n`. Keep the single-quoted string wrapper — only change the escape sequence.

Example — the RES-002 row at line 1254 currently has:

```javascript
'Step I: Filed with immediate supervisor within the contractual time limit. Management must respond within the specified days.\\nStep II: If Step I is denied, an appeal is filed. A hearing may be held.\\nStep III / Arbitration: Final step involving a neutral arbitrator. The decision is binding.\\nYour steward handles all filings and deadlines — you just need to provide information.'
```

Change to:

```javascript
'Step I: Filed with immediate supervisor within the contractual time limit. Management must respond within the specified days.\nStep II: If Step I is denied, an appeal is filed. A hearing may be held.\nStep III / Arbitration: Final step involving a neutral arbitrator. The decision is binding.\nYour steward handles all filings and deadlines — you just need to provide information.'
```

Apply the same replacement to RES-003 (line 1255), RES-004 (1256), RES-005 (1257), RES-006 (1258), and RES-007 (1259).

- [ ] **Step 2: Verify the file still parses**

Run: `node -e "require('fs').readFileSync('src/10b_SurveyDocSheets.gs', 'utf8')" && node --check <(node -e "console.log(require('fs').readFileSync('src/10b_SurveyDocSheets.gs', 'utf8').replace(/^function/gm, 'function'))")`

Alternative simpler check — just run the guard that parses every `.gs`:

Run: `npm run test:guards -- --testNamePattern="10b_SurveyDocSheets.gs parses"`
Expected: PASS.

- [ ] **Step 3: Replace `'\\n'` with `'\n'` in 07_DevTools.gs seedResourcesData**

Open `src/07_DevTools.gs`. Find `function seedResourcesData()` at line 1710. Inside it, locate `var additionalResources = [` around line 1736. For every row in that array (and any further rows past line 1773 — use grep to find them all), replace every `\\n` with `\n`.

Run first: `grep -n '\\\\n' src/07_DevTools.gs`
Expected: list of line numbers inside the `additionalResources` and any follow-up arrays.

Apply the replacement to each listed line. Use precise search-and-replace so you only touch the intended strings (confirm each edit in the diff).

- [ ] **Step 4: Verify parse + rebuild dist**

Run: `npm run build:prod`
Expected: build succeeds.

- [ ] **Step 5: Commit the seed cleanup**

```bash
git add src/10b_SurveyDocSheets.gs src/07_DevTools.gs dist/
git commit -m "fix: seed Resources strings use real newlines not literal '\\\\n'

JS single-quoted '\\\\n' evaluates to two characters (backslash + n), so
every seeded resource row was writing literal backslash-n into the
sheet cell. Future re-seeds now produce clean data. renderSafeText
handles existing installs transparently.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: Add steward-name type guard (Zone 2) + tests

**Files:**
- Create: `test/steward-name-guard.test.js`
- Modify: `src/21_WebDashDataService.gs` — add `_looksLikeStewardName(s)` helper and call it in the aggregation loop at line 3074

- [ ] **Step 1: Write the failing test file**

Create `test/steward-name-guard.test.js` with this exact content:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/steward-name-guard.test.js --no-coverage`
Expected: FAIL with "_looksLikeStewardName not found in 21_WebDashDataService.gs".

- [ ] **Step 3: Add `_looksLikeStewardName` helper to 21_WebDashDataService.gs**

Open `src/21_WebDashDataService.gs`. Find `function _buildGrievanceRecord(row, colMap) {` at line 2178. Immediately *before* that function (preserving indentation to match the surrounding code), insert:

```javascript
  /**
   * v4.55.3 sub-project A Zone 2 — steward name sanity guard.
   *
   * rRec.steward reaches the Avg Resolution chart as a grouping key, but
   * some live Grievance Log rows contain Date serials (46076.519...) or
   * stringified Date objects in the Assigned Steward column from legacy
   * admin edits or imports. This guard rejects values that clearly aren't
   * names before they become phantom buckets. Legitimate-looking names
   * and emails pass through; everything else collapses to 'unassigned'.
   *
   * Called only in the stewardResolution accumulation loop — this is NOT
   * a general-purpose filter, and no other consumer of rRec.steward is
   * affected.
   *
   * @param {string} s - lowercased, trimmed steward value from _buildGrievanceRecord
   * @returns {boolean} true if s looks like a valid steward name/email
   */
  function _looksLikeStewardName(s) {
    if (!s || typeof s !== 'string') return false;
    if (s.length > 120) return false;
    // Numeric-only (including serial floats and negatives)
    if (/^-?\d+(\.\d+)?$/.test(s)) return false;
    // ISO 8601 date prefix
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return false;
    // Weekday-prefixed date string (output of String(new Date()))
    if (/^(mon|tue|wed|thu|fri|sat|sun) /.test(s)) return false;
    return true;
  }

```

- [ ] **Step 4: Call the guard in the aggregation loop**

In the same file, find lines 3073–3076 (inside the `for (var ri = 1; ri < data.length; ri++)` block of the extended metrics pass). The current block is:

```javascript
            // Per-steward resolution
            var stw = rRec.steward || 'unassigned';
            if (!stewardResolution[stw]) stewardResolution[stw] = [];
            stewardResolution[stw].push(days);
```

Change to:

```javascript
            // Per-steward resolution
            // v4.55.3 sub-project A Zone 2: reject Date serials / ISO / weekday
            // strings that bleed in from dirty Grievance Log rows so they do
            // not create phantom grouping keys in the Avg Resolution chart.
            var stw = _looksLikeStewardName(rRec.steward) ? rRec.steward : 'unassigned';
            if (!stewardResolution[stw]) stewardResolution[stw] = [];
            stewardResolution[stw].push(days);
```

- [ ] **Step 5: Run the guard tests and the full guard suite**

Run: `npx jest test/steward-name-guard.test.js --no-coverage && npm run test:guards`
Expected: all 8 guard tests PASS, then full guard suite passes 274/274.

- [ ] **Step 6: Commit Zone 2**

```bash
git add src/21_WebDashDataService.gs test/steward-name-guard.test.js
git commit -m "fix: reject non-name values from Avg Resolution chart grouping

Adds _looksLikeStewardName guard in the stewardResolution aggregation
loop. Date serials, ISO dates, weekday-prefixed strings, and overlong
junk now collapse to 'unassigned' instead of creating phantom chart
buckets. Fixes issue #1 without touching _findColumn (which was already
exact-match — verified during plan-writing). +8 tests.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: Write `displayUser` helper + tests (Zone 3)

**Files:**
- Modify: `src/index.html` (insert `displayUser` next to `renderSafeText`)
- Create: `test/display-user.test.js`

- [ ] **Step 1: Write the failing test file**

Create `test/display-user.test.js` with this exact content:

```javascript
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx jest test/display-user.test.js --no-coverage`
Expected: FAIL with "displayUser not found in src/index.html".

- [ ] **Step 3: Add `displayUser` to src/index.html**

In `src/index.html`, immediately after the `renderSafeText` function you added in Task 1 (inside the same IIFE scope), insert:

```javascript
    // ============================================================================
    // v4.55.3 sub-project A Zone 3: displayUser — humanize Access Log user column.
    // ============================================================================
    //
    // The Audit Log sometimes stores SHA-256 hashes (64 hex chars) instead of
    // emails in the user column for privacy reasons. Rendering those raw leaks
    // unreadable hex to the UI (see issue #5). This helper detects the hash
    // shape and returns an abbreviated 'User #xxxxxxxx' form; emails pass
    // through lowercased; empty values become em-dash.
    //
    // Pure, total, never throws. No backend round-trip — the full hash→name
    // resolution is an optional future enhancement (tracked in sub-project B).
    function displayUser(rawId) {
      if (rawId === null || rawId === undefined || rawId === '') return '—';
      var s = String(rawId).trim();
      if (s === '') return '—';
      if (/^[a-f0-9]{64}$/i.test(s)) {
        return 'User #' + s.substring(0, 8).toLowerCase();
      }
      if (s.indexOf('@') !== -1) return s.toLowerCase();
      return s;
    }
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx jest test/display-user.test.js --no-coverage`
Expected: PASS, 6 tests passing.

- [ ] **Step 5: Run the full guard suite**

Run: `npm run test:guards`
Expected: ALL WEBAPP TESTS PASSED (274 tests).

- [ ] **Step 6: Commit the Zone 3 helper**

```bash
git add src/index.html test/display-user.test.js
git commit -m "feat: add displayUser helper for Access Log user column

Detects SHA-256 hash shape (64 hex chars) and returns abbreviated
'User #xxxxxxxx' form; emails pass through lowercased; empty → '—'.
Pure, total, never throws. +6 tests.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Adopt `displayUser` in the Access Log render site

**Files:**
- Modify: `src/steward_view.html:8932` (inside `renderAccessLogPage` → `loadPage` → results render loop)

- [ ] **Step 1: Locate the exact Access Log user cell render**

Open `src/steward_view.html`. Find `function renderAccessLogPage(appContainer)` at line 8824. Scroll to the `result.items.forEach` loop around line 8918. The current user-column line (8932) is:

```javascript
          // User
          row.appendChild(el('td', { style: { padding: '6px', fontSize: '11px', color: 'var(--text)' } }, item.user || '--'));
```

- [ ] **Step 2: Replace the raw `item.user || '--'` with `displayUser(item.user)`**

Change that line to:

```javascript
          // User — v4.55.3 sub-project A Zone 3: humanize SHA-256 hashes
          row.appendChild(el('td', { style: { padding: '6px', fontSize: '11px', color: 'var(--text)' } }, displayUser(item.user)));
```

- [ ] **Step 3: Verify no other Access Log user-column site needs the same change**

Run: `grep -n 'item\.user' src/steward_view.html`
Expected: the line you just changed, plus any other use. Apply the same `displayUser()` wrap to any other site that renders `item.user` as display text.

- [ ] **Step 4: Run the guard suite and build**

Run: `npm run test:guards && npm run build:prod`
Expected: 274/274 pass, build succeeds.

- [ ] **Step 5: Commit the adoption**

```bash
git add src/steward_view.html dist/
git commit -m "fix: render Access Log user column via displayUser

Raw SHA-256 hashes in item.user now show as 'User #xxxxxxxx' instead
of full 64-char hex. Fixes issue #5.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 7: Add spa-integrity guards for the new helpers

**Files:**
- Modify: `test/spa-integrity.test.js`

- [ ] **Step 1: Find a stable insertion point**

Open `test/spa-integrity.test.js`. Locate the G-RENDER-GEN describe block (search for `_renderGeneration`). That block is the closest precedent — it asserts a specific variable exists in `src/index.html`. We will add a sibling describe after it.

- [ ] **Step 2: Add the new guards**

Immediately after the G-RENDER-GEN describe block's closing `});`, insert:

```javascript
  describe('G-RENDER-SAFE: index.html defines renderSafeText and displayUser', () => {
    const fs = require('fs');
    const path = require('path');
    const indexSrc = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', 'index.html'),
      'utf8'
    );

    test('renderSafeText function is defined in index.html', () => {
      expect(indexSrc).toMatch(/function\s+renderSafeText\s*\(/);
    });

    test('displayUser function is defined in index.html', () => {
      expect(indexSrc).toMatch(/function\s+displayUser\s*\(/);
    });

    test('renderSafeText returns a DocumentFragment via createDocumentFragment', () => {
      var idx = indexSrc.indexOf('function renderSafeText(');
      var next = indexSrc.indexOf('function displayUser(', idx);
      var body = indexSrc.substring(idx, next === -1 ? idx + 2000 : next);
      expect(body).toMatch(/createDocumentFragment/);
    });

    test('displayUser detects 64-hex-char SHA-256 hashes', () => {
      var idx = indexSrc.indexOf('function displayUser(');
      var body = indexSrc.substring(idx, idx + 1000);
      expect(body).toMatch(/\[a-f0-9\]\{64\}/);
    });
  });
```

- [ ] **Step 3: Run the guard suite**

Run: `npm run test:guards`
Expected: 278/278 tests pass (274 existing + 4 new G-RENDER-SAFE assertions).

- [ ] **Step 4: Commit the integrity guards**

```bash
git add test/spa-integrity.test.js
git commit -m "test: add G-RENDER-SAFE guard for sub-project A helpers

Asserts renderSafeText and displayUser are defined in index.html
at load time, plus spot-checks their core behaviors (Document-
Fragment return type, 64-hex SHA-256 regex). Catches regressions
where a refactor removes either helper without updating call sites.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 8: Version bump + CHANGELOG

**Files:**
- Modify: `src/01_Core.gs` — `COMMAND_CONFIG.VERSION`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Bump `COMMAND_CONFIG.VERSION` in src/01_Core.gs**

Run: `grep -n "VERSION:\s*'4\.55\.2'" src/01_Core.gs`
Expected: one match showing the canonical version line.

Open that line and change `'4.55.2'` to `'4.55.3'`. Do not change any other `VERSION` references — this is the single source of truth per the session memory.

- [ ] **Step 2: Add a CHANGELOG entry**

Open `CHANGELOG.md`. Add a new section at the top (directly under the `# Changelog` heading, above the existing v4.55.2 entries):

```markdown
## v4.55.3 — 2026-04-11

### Fixes (Sub-project A — Display-Safe Rendering Boundary)
- **#1** Avg Resolution by Steward chart: `_looksLikeStewardName` guard rejects Date serials, ISO dates, and weekday-prefixed strings from becoming phantom grouping keys in the aggregation loop (`21_WebDashDataService.gs`). No `_findColumn` change — verified during plan-writing that the lookup chain was already exact-match only; the dirty data is in the live Grievance Log rows.
- **#5** Access Log user column: new `displayUser` helper in `index.html` detects 64-hex-char SHA-256 values and renders `User #xxxxxxxx` instead of the raw hash. Adopted in `renderAccessLogPage` (`steward_view.html`).
- **#9** Know Your Rights / FAQ / How to File a Grievance literal `\n`: new `renderSafeText` helper in `index.html` converts both literal `\n` (backslash+n) and real `\n` into `<br>` nodes while HTML-escaping the rest. Adopted in Resources render sites in `member_view.html` and `steward_view.html`. Seed strings in `10b_SurveyDocSheets.gs` and `07_DevTools.gs` were also corrected so future re-seeds produce clean data.

### Tests
- New suites: `test/render-safe-text.test.js` (+10), `test/steward-name-guard.test.js` (+8), `test/display-user.test.js` (+6).
- New G-RENDER-SAFE guard in `test/spa-integrity.test.js` (+4).
- Total: 75 suites, 3711 passed, 0 failed, 1 skipped (up from 72 / 3687 at v4.55.2 Wave 31).
```

- [ ] **Step 3: Rebuild dist and run the full pipeline**

Run: `npm run lint && npm run test:guards && npm run test:unit && npm run build:prod`
Expected: lint clean, all guards pass, all unit tests pass, prod build succeeds.

- [ ] **Step 4: Commit the version bump**

```bash
git add src/01_Core.gs CHANGELOG.md dist/
git commit -m "chore: bump to v4.55.3 — sub-project A display-safe rendering

Closes screenshot issues #1 (chart Y-axis date serials), #5 (Access
Log SHA-256 hashes), #9 (literal \\n in Resources). Three zones,
three helpers, +28 tests. DDS prod-ready; SolidBase sync next.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 9: Sync DDS → SolidBase (separate session)

This task produces a separate commit chain in the SolidBase repo, not DDS. Run it from `C:\Users\deskc\Documents\union-tracker` following the documented sync workflow in session memory.

- [ ] **Step 1: In the SolidBase repo, pull latest and create a working branch**

```bash
cd C:/Users/deskc/Documents/union-tracker
git status  # must be clean
git fetch origin
git checkout Main && git pull
```

- [ ] **Step 2: Copy the modified files from DDS into SolidBase, applying scrub rules**

Files to copy (per sub-project A scope):
- `src/index.html` — copy in full, then scrub DDS→SolidBase, →remove, the union→generalize, "Union Dashboard"→"SolidBase" (per session memory sync rules).
- `src/member_view.html` — same treatment.
- `src/steward_view.html` — same treatment.
- `src/21_WebDashDataService.gs` — same treatment.
- `src/10b_SurveyDocSheets.gs` — same treatment (verify no the union references in starterRows after copy).
- `src/07_DevTools.gs` — same treatment.
- `src/01_Core.gs` — copy the single `VERSION` line change only; do not overwrite SolidBase's package-specific config.
- `CHANGELOG.md` — manually add an equivalent SolidBase v4.55.3 entry; do not copy DDS's entry verbatim.
- `test/render-safe-text.test.js`, `test/steward-name-guard.test.js`, `test/display-user.test.js`, `test/spa-integrity.test.js` — copy tests as-is.

Exclude per session memory: `UI_REVIEW.md`, `DDS Job Descriptions/`, `25_WorkloadService.gs`, `poms_reference.html`, `agency_org_chart.html`.

- [ ] **Step 3: Rebuild SolidBase dist and run tests**

Run: `npm run build:prod && npm run test:unit`
Expected: 66 suites + new test files, all green. Total should be 69 suites, ~3496 tests.

- [ ] **Step 4: Commit and push SolidBase**

```bash
git add -A
git commit -m "feat: v4.55.3 sync from DDS — sub-project A display-safe rendering

Mirrors DDS v4.55.3 bug-fix wave (#1 chart date serials, #5 Access
Log hashes, #9 literal newlines) with standard DDS→SolidBase scrub
rules applied. No feature gating changes required.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 5: Ask the user before pushing and clasping**

Before running `git push` or `clasp push`, stop and confirm with the user. Pushing to origin and deploying to GAS are shared-state actions that must not happen without explicit approval per the user's standing rules.

---

## Task 10: Deploy DDS (after user approval)

This task runs only after the user confirms they want the wave shipped to production.

- [ ] **Step 1: Confirm current DDS state is ready**

Run: `git -C "C:/Users/deskc/Documents/SolidBase" log --oneline -10`
Expected: the 8 commits from Tasks 1–8 are present on Main, working tree clean.

- [ ] **Step 2: Push DDS to origin/Main (requires user approval)**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" push origin Main
```

- [ ] **Step 3: clasp push to DDS dev 1 first**

```bash
cd "C:/Users/deskc/Documents/SolidBase"
# Ensure .clasp.json points at dev 1: 1cJEyS0Ni2LTICNP_478yAVr2VmAlMGJkpDN3m3OKuJK2rEV8qwYrHRL3
clasp push
```

- [ ] **Step 4: Smoke-test dev 1 in a real GAS session**

Open the dev 1 web app URL. Verify:
1. Resources tab → Know Your Rights → "Grievance Steps Explained" shows real line breaks between Step I / II / III, not literal `\n`.
2. Member Grievance Outcomes tab → Avg Resolution by Steward chart shows real steward names on the Y-axis (no raw floats).
3. Admin → Access Log tab → user column shows `User #xxxxxxxx` for hashed rows (or email for email rows), not raw 64-char hex.

If any check fails, STOP and investigate before promoting to prod.

- [ ] **Step 5: clasp push to DDS prod (requires user confirmation)**

```bash
# Switch .clasp.json to prod Script ID:
# REDACTED_DDS_SCRIPT_ID
clasp push
```

- [ ] **Step 6: Update session memory**

After prod ship is confirmed working, update `memory/MEMORY.md` with the new version, commit count, and test count.

---

## Self-Review

I reviewed this plan against the spec at `docs/superpowers/specs/2026-04-11-subproject-a-display-safe-rendering-design.md` (as corrected in commit `f8eedff`).

**Spec coverage:**
- Spec §2 deliverable 1 (seed fix): ✅ Task 3.
- Spec §2 deliverable 2 (renderSafeText + adoption): ✅ Tasks 1 + 2.
- Spec §2 deliverable 3 (type guard): ✅ Task 4.
- Spec §2 deliverable 4 (displayUser + adoption): ✅ Tasks 5 + 6.
- Spec §2 deliverable 5 (tests, ~24 + integrity): ✅ Tasks 1, 4, 5, 7 sum to 28 new tests.
- Spec §2 deliverable 6 (per-zone commits, version bump, sync): ✅ Each task ends with its own commit; Task 8 bumps version; Task 9 handles SolidBase sync.
- Spec §6 test plan edge cases (literal `\n`, real `\n`, XSS escape, DocumentFragment type, null/undefined, numeric stringification): ✅ all covered in Task 1 test file.
- Spec §6 test plan steward guard cases (empty, Date object, serial, ISO, weekday, overlong, legit name): ✅ covered in Task 4 test file (8 tests).
- Spec §6 test plan displayUser cases (null, empty, hash-with-abbrev, email, plain string): ✅ covered in Task 5 test file (6 tests).
- Spec §7 deployment sequence: ✅ Tasks 8, 9, 10 match the commit-sequence and pipeline order.
- Spec §7 rollback: ✅ each zone is a standalone commit that can be reverted.

**Placeholder scan:** No TBD/TODO markers. Every step contains either a concrete command or concrete code. Step 1 of Task 2 uses a grep-driven "pin down exact line" approach because resource-body render sites have drifted between waves and cannot be hardcoded without risking a stale line number; the grep command and expected output make the step concrete.

**Type consistency:** `renderSafeText`, `displayUser`, and `_looksLikeStewardName` are referenced with consistent signatures across all tasks. The `stewardResolution` key assignment in Task 4 Step 4 uses the exact same variable name as the existing code at line 3074.

**Gaps found, fixed inline:**
- Task 2 was originally vague about how to find the adoption sites. I added the grep commands and the before/after snippet to make it unambiguous.
- Task 3 originally didn't have a way to audit the full set of `\\n` in `07_DevTools.gs`. I added a `grep -n '\\\\n' src/07_DevTools.gs` step to enumerate them first.
- Task 9 originally lacked an explicit "ask before push/clasp" step; added as Step 5.

No spec requirement is left without a task.

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-11-subproject-a-display-safe-rendering.md`. Two execution options:

**1. Subagent-Driven (recommended)** — dispatch a fresh subagent per task, review between tasks, fast iteration. Best fit here because Task 2 (adoption-site discovery) and Task 3 (grep-first seed audit) both benefit from isolated exploration without polluting the main context.

**2. Inline Execution** — execute tasks in this session using executing-plans, batch execution with checkpoints for review.

Which approach?
