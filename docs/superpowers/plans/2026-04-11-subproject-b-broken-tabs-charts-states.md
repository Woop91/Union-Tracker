# Sub-Project B — Broken Tab/Chart/State Logic Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship DDS v4.55.4 closing issues #2 (Monthly Trend chart black fill), #4 (radar empty-state misleading message), #7 (`renderQAForum is not defined` for pure stewards), #8 + survey open/close mismatch (period state coherence), plus a new `G-TAB-ROUTES-SAFE` integrity guard that prevents future tab-router regressions.

**Architecture:** Five isolated zones. Zone 1 adds explicit hex colors to the Monthly Trend `datasets` so the chart never falls through to `_palette()` and can't inherit theme-dependent blacks. Zone 2 updates the radar empty-state message to clearly explain it's a section-count visualization limit (not a response-count failure) and hides the radar view option when `sectionList.length < 3`. Zone 3 wraps the `qaforum` tab router calls in the existing `_loadMemberViewThen(container, callback)` lazy-loader so pure stewards load member_view.html on demand (the function `renderQAForum(appContainer, role)` only lives in member_view.html). Zone 4 normalizes "is there an active survey period?" to one backend-authoritative field (`periodActive` + `periodName`) consumed by both member and steward views. Zone 5 adds a test-time guard that parses every `case '…':` in `_handleTabNav` and asserts each routed function is defined somewhere in the concatenated src bundle.

**Tech Stack:** Google Apps Script (V8), vanilla JS inside GAS HtmlService templates (single-global-scope, IIFE-wrapped), Jest with no jsdom (helpers extracted from `src/index.html` via bracket-depth `new Function()` eval, same pattern as `test/fuzzy-match.test.js`). Build: zero-dep `build.js --prod --minify`. Pre-commit hook rebuilds `dist/` and runs 278 deploy guards. Deploy: `clasp push` to DDS prod Script ID `REDACTED_DDS_SCRIPT_ID`.

---

## File Structure

**Files created:**
- `test/radar-section-threshold.test.js` — tests the new section-count-aware empty-state message + hide-radar-when-<3 logic (~5 tests).
- `test/survey-period-state.test.js` — tests the normalized `periodActive`/`periodName` contract between backend and both views (~7 tests).

**Files modified:**
- `src/member_view.html` — Zone 1 (line 5372, Monthly Trend `datasets` colors), Zone 2 (line 3803, radar empty-state message + view picker hide), Zone 4 (line 2823/2827, survey open message cleanup — reads `surveyData.period.status` directly, which is what Zone 4 preserves as the authoritative check).
- `src/steward_view.html` — Zone 4 (line 4212, empty-state split into "no period active" vs "no members in scope"; reads new `periodActive` + `periodName` fields from `dataGetStewardSurveyTracking` response).
- `src/index.html` — Zone 3 (lines 4139 + 4167, wrap `renderQAForum` tab router calls in `_loadMemberViewThen`).
- `src/08e_SurveyEngine.gs` — Zone 4 diagnostic: read `getSurveyPeriod()` internals (lines 213+) to confirm whether it returns null for the "Q2 2026 Active" case. If it does, repair the period lookup. If it doesn't, the bug is elsewhere and no change is needed here.
- `src/21_WebDashDataService.gs` — Zone 4: add `periodActive: boolean` + `periodName: string | null` to `getStewardSurveyTracking()` return object so steward view can distinguish "no period" from "no members in scope."
- `test/spa-integrity.test.js` — Zone 5: new `G-TAB-ROUTES-SAFE` describe block with 3+ assertions covering every `case '…':` in `_handleTabNav` plus explicit coverage for `renderQAForum` (#7 regression).
- `src/01_Core.gs` — Task 6 (Version bump `COMMAND_CONFIG.VERSION` from `"4.55.3"` to `"4.55.4"` and `VERSION_INFO` fallback from `'4.55.3'` to `'4.55.4'`).
- `CHANGELOG.md` — Task 6 (new v4.55.4 entry at top).

**Files deliberately not touched:**
- `src/26_QAForum.gs` — the backend QAForum service. Zone 3's fix is purely a frontend tab-router change; no backend surface touched.
- `_palette()` in `src/index.html` — Zone 1 bypasses it rather than repairing it; repairing is out of scope.
- `renderQAForum` itself (the 400+ line function in `src/member_view.html`) — Zone 3 does NOT move it. The fix is the tab router, not the function.

---

## Task 1: Zone 1 — Monthly Trend explicit colors

**Files:**
- Modify: `src/member_view.html:5372`

- [ ] **Step 1: Read the current dataset definition**

Run: `grep -n "var datasets = \[{ label: 'Filed'" src/member_view.html`
Expected: one match at line 5372.

Confirm the current shape matches:

```javascript
        var datasets = [{ label: 'Filed', data: monthData }];
        if (data.monthlyResolved && data.monthlyResolved.length > 0) {
          var resolvedData = data.monthlyResolved.map(function(m) { return m.count; });
          datasets.push({ label: 'Resolved', data: resolvedData });
        }
```

- [ ] **Step 2: Add explicit color properties**

Change those lines to:

```javascript
        // v4.55.4 sub-project B Zone 1: explicit hex colors so the chart never
        // falls through to _palette() and can't inherit a theme-resolved black.
        var datasets = [{ label: 'Filed', data: monthData, color: '#f59e0b' }];
        if (data.monthlyResolved && data.monthlyResolved.length > 0) {
          var resolvedData = data.monthlyResolved.map(function(m) { return m.count; });
          datasets.push({ label: 'Resolved', data: resolvedData, color: '#10b981' });
        }
```

- [ ] **Step 3: Check for mirror render site in steward_view.html**

Run: `grep -n "Monthly Trend\|monthLabels\|data\.monthly" src/steward_view.html | head -10`

If there is a mirror site with identical `datasets = [{ label: 'Filed', data: ... }]` shape, apply the same change. If no match, skip — steward view doesn't render this chart.

- [ ] **Step 4: Run the guard suite**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:guards`
Expected: 278/278 pass (baseline; the G-RENDER-SAFE-MONTHLY-TREND guard is added in Task 5).

- [ ] **Step 5: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/member_view.html dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "fix: explicit Monthly Trend colors to bypass theme-dependent palette

datasets array now carries explicit #f59e0b (amber) for Filed and
#10b981 (emerald) for Resolved, so the chart never falls through
to _palette() and can't inherit a CSS-variable-resolved-to-black.
Fixes issue #2.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Zone 2 — Radar empty-state message clarity

**Files:**
- Modify: `src/member_view.html:3803` (radar empty-state message) and `:3798+` (the view picker that exposes radar as an option)
- Create: `test/radar-section-threshold.test.js`

- [ ] **Step 1: Write the failing test file**

Create `test/radar-section-threshold.test.js` with this exact content:

```javascript
/**
 * Radar section threshold — sub-project B Zone 2
 *
 * The radar chart needs ≥3 distinct sections (axes) to render as a
 * polygon. When there are fewer, the render short-circuits with a
 * clear message. The message must explain that this is a visualization
 * limit about section count, not a bug about response count.
 *
 * Static test: grep src/member_view.html for the old misleading string
 * and the new clearer string.
 */

const fs = require('fs');
const path = require('path');

const memberSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'member_view.html'),
  'utf8'
);

describe('Radar section threshold (sub-project B Zone 2)', () => {
  test('old misleading "Need 3+ sections for radar" string is gone', () => {
    expect(memberSrc).not.toMatch(/Need 3\+ sections for radar/);
  });

  test('new clearer message explains section-count visualization limit', () => {
    expect(memberSrc).toMatch(/radar chart needs at least 3 distinct sections/i);
  });

  test('new message mentions that this is a visualization limit not a data issue', () => {
    expect(memberSrc).toMatch(/visualization limit|visualization requirement|not a data issue/i);
  });

  test('radar view option is conditionally hidden when fewer than 3 sections', () => {
    // The view picker should only expose the radar option when at least 3 sections exist.
    // Look for a conditional guard around the radar option.
    expect(memberSrc).toMatch(/sectionList\.length\s*>=\s*3[\s\S]{0,200}radar/);
  });

  test('threshold constant is a named variable or inline 3 — not removed', () => {
    // Sanity: if someone accidentally removes the guard entirely, radar would crash on <3 sections.
    // Keep the guard, just improve the message.
    expect(memberSrc).toMatch(/sectionList\.length\s*<\s*3|n\s*<\s*3/);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/radar-section-threshold.test.js --no-coverage`
Expected: FAIL — "old misleading 'Need 3+ sections for radar' string is gone" will fail because the string still exists.

- [ ] **Step 3: Update the radar empty-state message**

Open `src/member_view.html`. Find line 3803 (inside `function _surveyViewRadar(vc, ctx)`):

```javascript
  if (n < 3) { cd.appendChild(el('div', { style: { fontSize: '11px', color: 'var(--muted)', textAlign: 'center', padding: '20px' } }, 'Need 3+ sections for radar')); vc.appendChild(cd); return; }
```

Change to:

```javascript
  if (n < 3) {
    // v4.55.4 sub-project B Zone 2: the previous "Need 3+ sections for radar"
    // string was ambiguous — users read it as "need 3+ responses" when they
    // had 123 responses but only 1–2 section groupings. This phrasing pins
    // the blame on the visualization limit, not the data.
    var msg = 'A radar chart needs at least 3 distinct sections to render. Your survey has ' + n + (n === 1 ? ' section' : ' sections') + ' — this is a visualization limit, not a data issue.';
    cd.appendChild(el('div', { style: { fontSize: '11px', color: 'var(--muted)', textAlign: 'center', padding: '20px', lineHeight: '1.5' } }, msg));
    vc.appendChild(cd);
    return;
  }
```

- [ ] **Step 4: Find the view picker and hide radar when sections < 3**

Run: `grep -n "_surveyViewRadar\|radar.*view\|view.*radar\|'radar'" src/member_view.html | head -20`

Expected: a list of references including the `_surveyViewRadar` definition (line 3798), the call site(s), and the view-switcher UI that exposes radar as a selectable view.

Locate the view-picker array (likely looks like `var views = [{ id: 'story', ... }, { id: 'heatmap', ... }, { id: 'sections', ... }, { id: 'radar', ... }]`). Add a conditional so the radar entry is only included when `sectionList.length >= 3`.

Example (adapt to the actual structure — the goal is to ensure the radar option doesn't appear if there are fewer than 3 sections):

```javascript
// BEFORE (example pattern)
var views = [
  { id: 'story', label: 'Story' },
  { id: 'heatmap', label: 'Heatmap' },
  { id: 'sections', label: 'Sections' },
  { id: 'radar', label: 'Radar' }
];

// AFTER — v4.55.4 sub-project B Zone 2: hide radar option when insufficient sections
var views = [
  { id: 'story', label: 'Story' },
  { id: 'heatmap', label: 'Heatmap' },
  { id: 'sections', label: 'Sections' }
];
if (sectionList.length >= 3) views.push({ id: 'radar', label: 'Radar' });
```

If the view picker structure is different (e.g., a switch statement or a hardcoded set of buttons), adapt the conditional to match. The invariant is: users cannot select the radar view when there are fewer than 3 sections.

- [ ] **Step 5: Run test to verify it passes**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/radar-section-threshold.test.js --no-coverage`
Expected: 5/5 PASS.

- [ ] **Step 6: Run the full guard suite**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:guards`
Expected: 278/278 pass.

- [ ] **Step 7: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/member_view.html test/radar-section-threshold.test.js dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "fix: radar empty-state message clarity + hide option when <3 sections

'Need 3+ sections for radar' was read by users as 'need 3+ responses'.
New phrasing explicitly cites the section-count visualization limit
and calls out that it is not a data issue. Radar option is also
hidden from the view picker when sectionList.length < 3 so users
cannot land on the empty state in the first place. +5 tests.
Fixes issue #4.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: Zone 3 — `renderQAForum` lazy-load wrap

**Files:**
- Modify: `src/index.html:4139` (member-role tab router) and `:4167` (steward-role tab router)

- [ ] **Step 1: Confirm the current call sites**

Run: `grep -n "case 'qaforum':" src/index.html`
Expected: 2 matches at lines ~4139 and ~4167.

Read the current shape:

```javascript
// Line 4139 (member-role router)
          case 'qaforum': return function(a) { renderQAForum(a, 'member'); };
// Line 4167 (steward-role router)
          case 'qaforum': return function(a) { renderQAForum(a, 'steward'); };
```

- [ ] **Step 2: Wrap both call sites in `_loadMemberViewThen`**

`_loadMemberViewThen(container, callback)` is defined in `src/index.html:3438`. For pure members it's a synchronous pass-through (member_view is baked in). For dual-role users and pure stewards, it fetches member_view.html, injects its script, then fires the callback. That's the existing mechanism — we just need to use it in the tab router.

Change line 4139 to:

```javascript
          case 'qaforum': return function(a) {
            // v4.55.4 sub-project B Zone 3: renderQAForum lives in member_view.html
            // so pure stewards need member_view lazy-loaded before invocation.
            // Pure members get a synchronous pass-through.
            _loadMemberViewThen(a, function() { renderQAForum(a, 'member'); });
          };
```

Change line 4167 to:

```javascript
          case 'qaforum': return function(a) {
            // v4.55.4 sub-project B Zone 3: see member-side note above.
            _loadMemberViewThen(a, function() { renderQAForum(a, 'steward'); });
          };
```

- [ ] **Step 3: Run the guard suite**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:guards`
Expected: 278/278 pass. The G-RENDER-SAFE guard from sub-project A doesn't specifically target the tab router, so the existing guard suite should remain green. Zone 5 below adds the comprehensive router guard.

- [ ] **Step 4: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/index.html dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "fix: wrap qaforum tab router in _loadMemberViewThen (pure steward fix)

renderQAForum is defined only in src/member_view.html but was called
from both the member and steward role branches of _handleTabNav.
Pure stewards never load member_view.html so the call crashed with
'renderQAForum is not defined'. Wrapping both call sites in the
existing _loadMemberViewThen lazy-loader fixes this for stewards
and is a no-op for pure members (sync pass-through). Fixes #7.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: Zone 4 — Survey period state coherence

**Files:**
- Modify: `src/08e_SurveyEngine.gs:213+` (diagnostic first; may or may not need changes)
- Modify: `src/21_WebDashDataService.gs` (`getStewardSurveyTracking` — add `periodActive` + `periodName` to return)
- Modify: `src/steward_view.html:4212` (split empty state into two messages)
- Modify: `src/member_view.html:2823-2830` (remove misleading quarterly hardcoded message; the existing `surveyData.period.status !== 'Active'` check IS the right question, just fix the messaging)
- Create: `test/survey-period-state.test.js`

- [ ] **Step 1: Diagnose `getSurveyPeriod()` behavior**

Open `src/08e_SurveyEngine.gs` starting at line 213. Read the entire `getSurveyPeriod()` function through its return. Determine how it decides whether a period is Active. Look for:

- Sheet reads from `HIDDEN_SHEETS.SURVEY_PERIODS`
- Row iteration and status-column lookup
- Any `new Date()`/month-modulo logic that could override a manually-Active status

Report back in your implementation notes: does `getSurveyPeriod()` correctly return the row when status is 'Active', or does it have an extra month-based filter that suppresses the Q2 2026 case?

Run: `sed -n '213,280p' src/08e_SurveyEngine.gs`
(use whatever paging tool works in your shell; or Read the file with the Read tool over that range)

**Decision gate:**
- If `getSurveyPeriod()` is correct (no extra filter), skip Step 2 entirely. The root cause of the survey open/close mismatch is elsewhere (most likely the member view's check or an endpoint stale-cache issue that was already addressed in v4.55.1's cache-key v2 bump).
- If `getSurveyPeriod()` has a month-based filter, proceed to Step 2.

- [ ] **Step 2: Repair `getSurveyPeriod()` if needed**

Only if Step 1 found a bug. Remove any month-modulo / quarterly-window check that suppresses an Active status. The source of truth must be the sheet's status column, nothing else.

Show the exact before/after edit in your implementation log.

- [ ] **Step 3: Write the failing test file**

Create `test/survey-period-state.test.js`:

```javascript
/**
 * Survey period state coherence — sub-project B Zone 4
 *
 * Tests the contract that both member_view and steward_view derive
 * "is there an active survey period?" from one backend-authoritative
 * answer, not from client-side month logic.
 *
 * Static tests verify the source files:
 * - No hardcoded "Jan, Apr, Jul, Oct" quarterly check remains in member_view.html
 * - steward_view.html distinguishes "no period active" vs "no members in scope"
 * - steward_view.html reads periodActive / periodName from the backend response
 */

const fs = require('fs');
const path = require('path');

const memberSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'member_view.html'),
  'utf8'
);
const stewardSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'steward_view.html'),
  'utf8'
);
const dataServiceSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'),
  'utf8'
);

describe('Survey period state coherence (sub-project B Zone 4)', () => {
  test('member_view: hardcoded quarterly "Jan, Apr, Jul, Oct" message is gone', () => {
    expect(memberSrc).not.toMatch(/Jan, Apr, Jul, Oct/);
  });

  test('member_view: still guards on period.status Active (backend source of truth)', () => {
    expect(memberSrc).toMatch(/period\.status\s*!==\s*['"]Active['"]|period\.status\s*===\s*['"]Active['"]/);
  });

  test('member_view: new "no survey open" message explains steward control, not quarterly schedule', () => {
    expect(memberSrc).toMatch(/your steward|steward will notify|opens when.*steward|admin activat/i);
  });

  test('steward_view: distinguishes "no period active" from "no members in scope"', () => {
    // Two distinct messages in the empty-state area around line 4212
    expect(stewardSrc).toMatch(/No survey period is currently active/);
    expect(stewardSrc).toMatch(/no members.*scope|no members.*assigned|expand.*scope/i);
  });

  test('getStewardSurveyTracking backend returns periodActive + periodName', () => {
    // The function should include these new fields in its return object
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?periodActive/);
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?periodName/);
  });

  test('getStewardSurveyTracking periodActive is derived from getSurveyPeriod() not client-side logic', () => {
    // Source the periodActive boolean from getSurveyPeriod() return, not from a month-modulo
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?getSurveyPeriod/);
    // Guard: no suspicious month-modulo arithmetic in this function
    const fn = dataServiceSrc.match(/function getStewardSurveyTracking[\s\S]*?(?=\n  function |\nfunction )/);
    if (fn) {
      expect(fn[0]).not.toMatch(/getMonth\(\)\s*%/);
    }
  });

  test('backend return object has both periodActive (boolean) and periodName (string or null)', () => {
    // Narrow shape check: the function literally assigns these keys
    expect(dataServiceSrc).toMatch(/periodActive:\s*!!\w+|periodActive:\s*\(.*\.status\s*===\s*['"]Active['"]\)|periodActive:\s*Boolean/);
    expect(dataServiceSrc).toMatch(/periodName:\s*\w+\.name\s*\|\|\s*null|periodName:\s*\(.*\.name/);
  });
});
```

- [ ] **Step 4: Run test to verify all 7 fail**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/survey-period-state.test.js --no-coverage`
Expected: FAIL — several or all 7 tests fail because the source changes haven't been made yet.

- [ ] **Step 5: Add `periodActive` + `periodName` to `getStewardSurveyTracking` backend**

Open `src/21_WebDashDataService.gs`. Find `function getStewardSurveyTracking(stewardEmail, scope)` (grep if the line number has shifted). At the top of the function body, fetch the current period:

```javascript
    // v4.55.4 sub-project B Zone 4: compute period state once from the backend
    // source of truth so both views can trust the same answer.
    var period = null;
    try { period = (typeof getSurveyPeriod === 'function') ? getSurveyPeriod() : null; } catch (_) { period = null; }
    var periodActive = !!(period && period.status === 'Active');
    var periodName = (period && period.name) ? period.name : null;
```

Then find the return statement at the end of the function and add both fields. Before:

```javascript
    return { total: members.length, completed: completedCount, members: members };
```

After:

```javascript
    return {
      total: members.length,
      completed: completedCount,
      members: members,
      periodActive: periodActive,
      periodName: periodName
    };
```

If there are multiple return statements in the function (e.g., early returns for "no spreadsheet" or "no sheet"), add `periodActive: false, periodName: null` to each so the shape is consistent across all paths.

Also update the `log_` / try-catch fallback in the wrapper `src/21d_WebDashDataWrappers.gs:197` to include the new fields on failure:

```javascript
function dataGetStewardSurveyTracking(sessionToken, scope) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; try { return DataService.getStewardSurveyTracking(s, scope); } catch (e) { log_('dataGetStewardSurveyTracking error', e.message + '\n' + (e.stack || '')); return { total: 0, completed: 0, members: [], periodActive: false, periodName: null }; } }
```

- [ ] **Step 6: Update `src/steward_view.html:4212` empty-state logic**

Find the current empty-state block around line 4212:

```javascript
        // ── No data ───────────────────────────────────────────────────────
        if (!total) {
          dataWrap.appendChild(el('div', { className: 'empty-state-sm' }, 'No survey tracking data. Open a period and populate tracking.'));
          return;
        }
```

Change to:

```javascript
        // ── No data ───────────────────────────────────────────────────────
        // v4.55.4 sub-project B Zone 4: distinguish "no period active" from
        // "period is active but no members match the current scope". The two
        // states want different user-facing messages and different actions.
        if (!result.periodActive) {
          dataWrap.appendChild(el('div', { className: 'empty-state-sm' }, 'No survey period is currently active. Open a new period from the Survey Admin panel.'));
          return;
        }
        if (!total) {
          dataWrap.appendChild(el('div', { className: 'empty-state-sm' }, 'Period "' + (result.periodName || 'active period') + '" is active but no members are assigned in your current scope. Expand the scope filter above.'));
          return;
        }
```

Note: you'll need to confirm the variable name `result` matches whatever the outer scope uses for the `dataGetStewardSurveyTracking` response. Grep for `dataGetStewardSurveyTracking` in the file and trace back the variable name. If it's called something else (e.g., `tracking`, `response`, `data`), substitute.

- [ ] **Step 7: Update `src/member_view.html:2823-2830` survey open message**

The existing check at line 2823 is CORRECT — it derives "is open" from `surveyData.period.status === 'Active'`, which is the backend source of truth. We only need to fix the misleading message text at line 2827.

Current:

```javascript
      if (!surveyData || !surveyData.period || surveyData.period.status !== 'Active') {
        var closedCard = el('div', { className: 'card card-glass', style: { padding: '28px', textAlign: 'center' } });
        closedCard.appendChild(el('div', { style: { fontSize: '40px', marginBottom: '12px' } }, '\uD83D\uDCCB'));
        closedCard.appendChild(el('div', { style: { fontFamily: 'var(--fontDisplay)', fontSize: '16px', fontWeight: '600', color: 'var(--text)', marginBottom: '8px' } }, 'No Survey Open'));
        closedCard.appendChild(el('div', { style: { fontSize: '13px', color: 'var(--muted)', lineHeight: '1.6' } }, 'The quarterly satisfaction survey is not currently open. Surveys open at the start of each quarter — Jan, Apr, Jul, Oct.'));
        content.appendChild(closedCard);
        return;
      }
```

Change the misleading message body (line 2827) to:

```javascript
      if (!surveyData || !surveyData.period || surveyData.period.status !== 'Active') {
        var closedCard = el('div', { className: 'card card-glass', style: { padding: '28px', textAlign: 'center' } });
        closedCard.appendChild(el('div', { style: { fontSize: '40px', marginBottom: '12px' } }, '\uD83D\uDCCB'));
        closedCard.appendChild(el('div', { style: { fontFamily: 'var(--fontDisplay)', fontSize: '16px', fontWeight: '600', color: 'var(--text)', marginBottom: '8px' } }, 'No Survey Open'));
        // v4.55.4 sub-project B Zone 4: drop the misleading "Jan, Apr, Jul, Oct"
        // hardcoded quarterly schedule — surveys are opened manually by stewards,
        // not on a calendar. The message must match the actual control model.
        closedCard.appendChild(el('div', { style: { fontSize: '13px', color: 'var(--muted)', lineHeight: '1.6' } }, 'No survey period is currently active. Your steward will notify you when the next survey opens.'));
        content.appendChild(closedCard);
        return;
      }
```

- [ ] **Step 8: Run the tests**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/survey-period-state.test.js --no-coverage`
Expected: 7/7 PASS.

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:guards`
Expected: 278/278 pass.

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:unit 2>&1 | tail -15`
Expected: all suites green.

- [ ] **Step 9: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/21_WebDashDataService.gs src/21d_WebDashDataWrappers.gs src/steward_view.html src/member_view.html test/survey-period-state.test.js dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "fix: coherent survey period state across member + steward views

getStewardSurveyTracking now returns periodActive + periodName derived
from getSurveyPeriod() so the steward empty state can distinguish
'no period' from 'no members in scope'. Member view drops the
misleading 'Jan, Apr, Jul, Oct' quarterly hardcoded message — surveys
are opened manually by stewards, not on a calendar. Both views now
resolve 'is the survey open?' to the same backend field. +7 tests.
Fixes issue #8 and the survey open/close mismatch.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: Zone 5 — G-TAB-ROUTES-SAFE integrity guard

**Files:**
- Modify: `test/spa-integrity.test.js`

- [ ] **Step 1: Find the insertion point**

Open `test/spa-integrity.test.js`. Locate the `G-RENDER-SAFE` describe block added in sub-project A (search for `G-RENDER-SAFE`). Insert the new block immediately after it at the same scope level.

- [ ] **Step 2: Add the new describe block**

After the G-RENDER-SAFE block's closing `});`, insert:

```javascript
  describe('G-TAB-ROUTES-SAFE: every _handleTabNav case routes to a defined function', () => {
    const fs = require('fs');
    const path = require('path');

    // Concatenate all src/*.html and src/*.gs into one big string so we can
    // check whether any function name is defined *somewhere* in the bundle.
    const srcDir = path.resolve(__dirname, '..', 'src');
    const files = fs.readdirSync(srcDir).filter(f => f.endsWith('.html') || f.endsWith('.gs'));
    const bundle = files.map(f => fs.readFileSync(path.join(srcDir, f), 'utf8')).join('\n');

    const indexSrc = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'index.html'), 'utf8');

    // Extract the body of _handleTabNav via bracket-depth parsing.
    function extractHandleTabNavBody() {
      const startIdx = indexSrc.indexOf('function _handleTabNav');
      if (startIdx === -1) throw new Error('_handleTabNav not found in src/index.html');
      let depth = 0;
      let inFn = false;
      let endIdx = startIdx;
      for (let i = startIdx; i < indexSrc.length; i++) {
        const ch = indexSrc.charAt(i);
        if (ch === '{') { depth++; inFn = true; continue; }
        if (ch === '}') {
          depth--;
          if (inFn && depth === 0) { endIdx = i + 1; break; }
        }
      }
      return indexSrc.substring(startIdx, endIdx);
    }

    const bodyText = extractHandleTabNavBody();

    // Find every function name invoked from a case block. Pattern: capture
    // identifiers that appear as `IDENTIFIER(` inside the body, then filter
    // to the ones that look like render* or similar view entry points.
    // For v1 of this guard we walk each case and capture the first identifier
    // followed by `(` inside the case body.
    function extractRoutedFunctionNames() {
      const caseRegex = /case\s*['"]([\w-]+)['"]\s*:\s*([\s\S]*?)(?=\bcase\s|\bdefault\s|\}\s*$)/g;
      const names = new Set();
      let m;
      while ((m = caseRegex.exec(bodyText)) !== null) {
        const caseBody = m[2];
        // Match identifiers that are called as functions, skipping _loadMemberViewThen wrappers
        const callRegex = /(?:^|[^\w.])(render[A-Z]\w*|init[A-Z]\w*|_?show[A-Z]\w*)\s*\(/g;
        let c;
        while ((c = callRegex.exec(caseBody)) !== null) {
          names.add(c[1]);
        }
      }
      return Array.from(names);
    }

    const routedNames = extractRoutedFunctionNames();

    function isDefinedInBundle(name) {
      const patterns = [
        new RegExp('function\\s+' + name + '\\s*\\('),
        new RegExp('\\b' + name + '\\s*=\\s*function'),
        new RegExp('\\b(var|let|const)\\s+' + name + '\\s*=\\s*function'),
      ];
      return patterns.some(p => p.test(bundle));
    }

    test('at least one routed function was extracted (sanity)', () => {
      expect(routedNames.length).toBeGreaterThan(0);
    });

    test('every routed function name is defined somewhere in src/', () => {
      const missing = routedNames.filter(n => !isDefinedInBundle(n));
      if (missing.length > 0) {
        throw new Error('Tab router references undefined function(s): ' + missing.join(', '));
      }
      expect(missing).toEqual([]);
    });

    test('renderQAForum specifically is defined (sub-project B #7 regression guard)', () => {
      expect(isDefinedInBundle('renderQAForum')).toBe(true);
    });

    test('Monthly Trend datasets contain explicit color properties (Zone 1 regression guard)', () => {
      const memberSrc = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'member_view.html'), 'utf8');
      // The datasets array at the Monthly Trend render site must include color: '#...' properties.
      const idx = memberSrc.indexOf("label: 'Filed'");
      expect(idx).toBeGreaterThan(-1);
      // Look at a ~300-char window around the "Filed" label for an explicit hex color.
      const window = memberSrc.substring(idx, idx + 300);
      expect(window).toMatch(/color:\s*['"]#[0-9a-fA-F]{3,8}['"]/);
    });
  });
```

- [ ] **Step 3: Run the guard suite**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run test:guards`
Expected: 282/282 pass (278 baseline + 4 new G-TAB-ROUTES-SAFE assertions).

If the "every routed function name is defined somewhere in src/" test fails with missing names, read the error — it lists the dangling references. For each one, verify whether it's a legitimate oversight (another latent bug to fix) or a false positive from the extractor regex. If it's a real bug, fix the tab router or the function definition in a follow-up commit (note: this counts as a sub-project B bonus find).

- [ ] **Step 4: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add test/spa-integrity.test.js
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "test: add G-TAB-ROUTES-SAFE guard for _handleTabNav

Parses _handleTabNav via bracket-depth, extracts every routed render*
function name, and asserts each is defined somewhere in the
concatenated src bundle. Catches renderQAForum-style regressions
before they ship. Plus explicit regression guards for sub-project A
Monthly Trend color literal (Zone 1) and issue #7 (renderQAForum
defined). +4 tests (278 → 282 total).

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Version bump + CHANGELOG

**Files:**
- Modify: `src/01_Core.gs`
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Bump `COMMAND_CONFIG.VERSION` and the fallback**

Open `src/01_Core.gs`. Find the line `VERSION: "4.55.3",` (around line 336) and change to:

```javascript
  VERSION: "4.55.4",
```

Then find the fallback line around 508:

```javascript
  var ver = (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.VERSION) ? COMMAND_CONFIG.VERSION : '4.55.3';
```

Change the fallback to `'4.55.4'`:

```javascript
  var ver = (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.VERSION) ? COMMAND_CONFIG.VERSION : '4.55.4';
```

- [ ] **Step 2: Add CHANGELOG entry**

Open `CHANGELOG.md`. Add a new section at the top, directly under the existing v4.55.3 entry:

```markdown
## [4.55.4] - 2026-04-11

### Fixed (Sub-project B — Broken Tab/Chart/State Logic)
- **#2** Monthly Trend chart "Filed" series rendered as solid black fill. Explicit hex colors (`#f59e0b` amber for Filed, `#10b981` emerald for Resolved) added to the `datasets` array in `member_view.html` so the chart never falls through to `_palette()` and can't inherit a theme-resolved black.
- **#4** Section Balance radar chart showed "Need 3+ sections for radar" in a way users read as "need 3+ responses." New message: "A radar chart needs at least 3 distinct sections to render. Your survey has N sections — this is a visualization limit, not a data issue." Radar option is also hidden from the view picker when fewer than 3 sections exist, so the empty state is unreachable by normal navigation.
- **#7** `renderQAForum is not defined` tab load failure for pure stewards. `renderQAForum` only lives in `member_view.html`, which pure stewards never load. Both `case 'qaforum':` branches in `_handleTabNav` now wrap the call in the existing `_loadMemberViewThen(container, callback)` lazy-loader (synchronous pass-through for pure members, lazy-fetch for dual-role and pure stewards).
- **#8** + survey open/close mismatch: `getStewardSurveyTracking` now returns `periodActive: boolean` and `periodName: string | null` derived from `getSurveyPeriod()` — one backend source of truth for both views. Steward empty state split into "no survey period active" and "period active but no members in scope" messages. Member view drops the misleading "Jan, Apr, Jul, Oct" hardcoded quarterly message — surveys are opened manually by stewards, not on a calendar.

### Tests
- New suites: `test/radar-section-threshold.test.js` (+5), `test/survey-period-state.test.js` (+7).
- New `G-TAB-ROUTES-SAFE` guard in `test/spa-integrity.test.js` (+4): parses every `case '…':` in `_handleTabNav` and asserts each routed function name is defined somewhere in the concatenated src bundle. Catches the #7 regression class at CI time.
- Total deploy guards: 282 (up from 278).
```

- [ ] **Step 3: Run the full pipeline**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npm run lint && npm run test:guards && npm run test:unit && npm run build:prod`
Expected: lint clean, 282/282 guards green, all unit tests green, prod build succeeds.

- [ ] **Step 4: Commit**

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/01_Core.gs CHANGELOG.md dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "chore: bump to v4.55.4 — sub-project B broken tabs/charts/states

Closes screenshot issues #2 (Monthly Trend black fill), #4 (radar
empty-state clarity), #7 (renderQAForum undefined for pure stewards),
and #8 + survey open/close mismatch (period state coherence). Adds
G-TAB-ROUTES-SAFE regression guard. Five zones, +16 tests (282
total guards). DDS prod-ready; SolidBase per-wave sync next.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Task 7: DDS push + clasp to prod

- [ ] **Step 1: Fetch origin before pushing (lesson from sub-project A)**

Run: `git -C "C:/Users/deskc/Documents/SolidBase" fetch origin && git -C "C:/Users/deskc/Documents/SolidBase" rev-list --left-right --count origin/Main...HEAD`
Expected: `0    N` (zero behind, N ahead where N is the number of commits in this wave). If left > 0, rebase: `git -C "C:/Users/deskc/Documents/SolidBase" pull --rebase origin Main`.

- [ ] **Step 2: Push to origin/Main**

Run: `git -C "C:/Users/deskc/Documents/SolidBase" push origin Main 2>&1 | tail -5`
Expected: `N..M  Main -> Main` fast-forward push. Pre-push hook runs the full test suite and must pass before the push completes.

- [ ] **Step 3: Verify .clasp.json target is DDS prod**

Run: `cat "C:/Users/deskc/Documents/SolidBase/.clasp.json"`
Expected: `{"scriptId":"REDACTED_DDS_SCRIPT_ID","rootDir":"dist"}`

- [ ] **Step 4: clasp push to DDS prod**

Run: `cd "C:/Users/deskc/Documents/SolidBase" && clasp push --force 2>&1 | grep -E 'Pushed|error|Error|fail' | tail -5`
Expected: `Pushed 63 files at HH:MM:SS.`

---

## Task 8: DDS → SolidBase per-wave sync (Option X)

**Files to copy (same list as DDS modifications in Tasks 1–6):**
- src/member_view.html
- src/steward_view.html
- src/index.html
- src/08e_SurveyEngine.gs (only if Step 1 of Task 4 required a fix)
- src/21_WebDashDataService.gs
- src/21d_WebDashDataWrappers.gs
- src/01_Core.gs (VERSION + fallback bump — but only the lines, not the full file; SB's version chain is ahead by v4.55.1 already)
- test/radar-section-threshold.test.js
- test/survey-period-state.test.js
- test/spa-integrity.test.js
- CHANGELOG.md (add equivalent v4.55.4 entry to SB)

- [ ] **Step 1: Verify SolidBase is clean and at v4.55.1**

Run: `cd "C:/Users/deskc/Documents/union-tracker" && git status --short && git log --oneline -3`
Expected: working tree clean, latest commit is `5772a7e feat: v4.55.1 sync from DDS — committed interrupted work`.

- [ ] **Step 2: Apply the zone fixes to SB**

Sub-project B's changes are almost entirely non-DDS-specific (no the union, no , no POMS/WorkloadService references). Copy each modified file from DDS into SB with these rules:

- **Zone 1** (member_view.html color literals): copy the two `color: '#f59e0b'` and `color: '#10b981'` changes into SB's `src/member_view.html` at the Monthly Trend render site (line number may differ).
- **Zone 2** (member_view.html radar message + view picker): copy the radar empty-state message change + view-picker conditional into SB.
- **Zone 3** (index.html qaforum router): copy the `_loadMemberViewThen` wrap into SB's `src/index.html` tab router.
- **Zone 4** (21_WebDashDataService.gs, 21d_WebDashDataWrappers.gs, steward_view.html, member_view.html): copy the periodActive/periodName backend changes and both view-side updates into SB.
- **Zone 5** (spa-integrity.test.js): copy the G-TAB-ROUTES-SAFE describe block into SB's test file.
- **Test files**: copy `test/radar-section-threshold.test.js` and `test/survey-period-state.test.js` verbatim into SB's `test/` directory.

**Scrub rules (standard DDS→SB):**
- Any mention of "DDS" in new content → "SolidBase"
- Any mention of "the union" in new content → generalize (e.g., "the union")
- Any mention of "" → remove

The new content in sub-project B doesn't contain any of these strings (verified during spec authoring), so the scrub rules are no-ops for this wave. Double-check during the copy.

- [ ] **Step 3: Bump SB VERSION to 4.55.4**

Open SB's `src/01_Core.gs`. Find the `VERSION: "4.55.1"` line and change to `"4.55.4"`. Find the fallback and change from whatever it is to `'4.55.4'`.

Note: SB is jumping from v4.55.1 directly to v4.55.4, skipping v4.55.2 and v4.55.3. The CHANGELOG entry needs to cover that — add entries for v4.55.2 and v4.55.3 AT THE END of this task's CHANGELOG update, OR add a single combined entry noting the skip.

Simpler approach for this wave: add only the v4.55.4 entry to the SB CHANGELOG, and in a separate follow-up task (tracked as "v4.55.2+v4.55.3 catchup sync") we do the full catchup. For v4.55.4 alone, the CHANGELOG entry should mirror DDS's.

- [ ] **Step 4: Add v4.55.4 CHANGELOG entry to SB**

Copy the v4.55.4 CHANGELOG entry from DDS verbatim to the top of SB's `CHANGELOG.md` (under the existing v4.55.1 entry, which should now be at the top).

- [ ] **Step 5: Run SB tests**

Run: `cd "C:/Users/deskc/Documents/union-tracker" && npm run build:prod && npm run test:unit 2>&1 | tail -10`
Expected: all tests green. Test count should grow by +16 (5 + 7 + 4) from SB's current baseline.

- [ ] **Step 6: Commit SB**

```bash
git -C "C:/Users/deskc/Documents/union-tracker" add -A
git -C "C:/Users/deskc/Documents/union-tracker" commit -m "feat: v4.55.4 sync from DDS — sub-project B broken tabs/charts/states

Mirrors DDS v4.55.4 bug-fix wave (#2 Monthly Trend black fill, #4
radar empty-state clarity, #7 renderQAForum pure-steward lazy-load
wrap, #8 + survey open/close period state coherence) plus the
G-TAB-ROUTES-SAFE regression guard. No the union / /
POMS / Workload content in this wave; scrub rules are no-ops.

SB is jumping v4.55.1 → v4.55.4 in this commit. The v4.55.2 + v4.55.3
catchup (Waves 18–31 + sub-project A Zones 1–5) is still pending and
tracked for a dedicated catchup sync pass.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 7: Push SB**

Run: `git -C "C:/Users/deskc/Documents/union-tracker" fetch origin && git -C "C:/Users/deskc/Documents/union-tracker" rev-list --left-right --count origin/Main...HEAD`
Expected: `0    1` (clean ahead, zero behind). If any divergence, rebase.

Then: `git -C "C:/Users/deskc/Documents/union-tracker" push origin Main 2>&1 | tail -5`

- [ ] **Step 8: clasp push SB prod**

Verify .clasp.json: `cat "C:/Users/deskc/Documents/union-tracker/.clasp.json"`
Expected: `{"scriptId":"1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl","rootDir":"./dist"}`

Run: `cd "C:/Users/deskc/Documents/union-tracker" && clasp push --force 2>&1 | grep -E 'Pushed|error|Error|fail' | tail -5`
Expected: `Pushed 59 files at HH:MM:SS.` (or 60 if the build.js new entry caused another file to mirror).

---

## Self-Review

Checking the plan against the spec at `docs/superpowers/specs/2026-04-11-subproject-b-broken-tabs-charts-states-design.md`.

**Spec coverage:**
- Spec §2 deliverable 1 (Zone 1 Monthly Trend colors): ✅ Task 1.
- Spec §2 deliverable 2 (Zone 2 radar): ✅ Task 2 (message + view-picker hide, not the aggregation fix — I confirmed during plan-writing that the aggregation is already correct, so the bug is messaging + view-picker UX).
- Spec §2 deliverable 3 (Zone 3 renderQAForum): ✅ Task 3 — fix shape decided: `_loadMemberViewThen` wrap, reusing existing lazy-loader infrastructure.
- Spec §2 deliverable 4 (Zone 4 survey period state): ✅ Task 4, with the `08e_SurveyEngine.gs` diagnostic step explicitly called out as a decision gate in Step 1.
- Spec §2 deliverable 5 (Zone 5 G-TAB-ROUTES-SAFE): ✅ Task 5.
- Spec §2 deliverable 6 (version bump, commits, push, clasp, SB sync): ✅ Tasks 6, 7, 8.
- Spec §6 test plan: ✅ ~17 new tests across 3 files (5 radar, 7 period, 4 G-TAB-ROUTES-SAFE — 1 more than the spec said because I added the at-least-one-routed-function sanity test).

**Placeholder scan:** No TBD/TODO/placeholders in binding steps. Task 4 Step 1 is a diagnostic-first step by design — its decision gate is explicit and the alternative outcomes are listed with concrete conditions.

**Type consistency:**
- `periodActive` (boolean) and `periodName` (string | null) — used consistently across Tasks 4 Steps 5, 6, 7, and in test/survey-period-state.test.js assertions.
- `_loadMemberViewThen(container, callback)` — referenced consistently in Task 3 Step 2 (both call sites pass `a` as container).
- Monthly Trend colors `#f59e0b` and `#10b981` — used consistently in Task 1 Step 2 and Task 5 Step 2's regression guard regex.

No issues found. Plan is ready for execution.

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-04-11-subproject-b-broken-tabs-charts-states.md`. Given the user's explicit "continue w/ asking for my approval" and "all" directives, execution proceeds via **superpowers:subagent-driven-development** — fresh subagent per task with two-stage review (spec compliance first, code quality second) after each task, using the implementer/spec-reviewer/code-quality-reviewer prompt templates.
