# Sub-Project B — Broken Tab / Chart / Empty-State Logic

**Date:** 2026-04-11
**Target release:** DDS v4.55.4 (patch wave), synced to SolidBase after DDS merges per Option X
**Status:** Design approved (user deferred section-by-section gating — whole-spec approval mode)
**Closes issues:** #2 (Monthly Trend chart "Filed" series renders as solid black fill), #4 (Section Balance radar shows "Need 3+ sections" despite 123 responses), #7 (`renderQAForum is not defined` tab load failure), #8 (Survey Tracking "No survey tracking data" despite Q2 2026 being active), Survey "shows open" vs "not currently open" mismatch

---

## 1. Motivation

Five production defects in the member and steward views share one structural trait: each is a case where the frontend fails to render real data because a threshold, guard, color lookup, or function name doesn't match what the backend (or the rest of the codebase) produces. Each looks isolated; together they reveal that DDS lacks integrity guards that would catch this class of bug before it ships. Sub-project B fixes the five bugs AND adds one new test-time integrity guard (`G-TAB-ROUTES-SAFE`) that would have caught #7 at CI time.

- **#2** — The Monthly Trend line/area chart's "Filed" series paints as a solid black fill obscuring the chart. The chart code falls through to `_palette()` when a dataset has no explicit `color`, and the palette resolves a CSS custom property that evaluates to black in some theme states.
- **#4** — The Section Balance radar chart shows "Need 3+ sections for radar" even when the survey has 123 responses. The threshold check counts distinct *section labels* (not responses), and the empty-state message is misleading. Either the threshold is wrong, the aggregation is wrong, or both.
- **#7** — The Q&A forum tab fails with `renderQAForum is not defined` at runtime. The tab router in `src/index.html` invokes a function name that either was renamed in Wave 31 (v4.55.2), never got registered in the global view scope, or loses its definition due to script load order.
- **#8** — The steward Survey Tracking tab shows "No survey tracking data. Open a period and populate tracking." even though the Q2 2026 Member Satisfaction Survey is marked Active. The empty-state check conflates "no active period" with "no members in the current scope."
- **Survey open/close mismatch** — The member-facing survey card says the survey is "currently open," but opening it lands on a page that says "The quarterly satisfaction survey is not currently open. Surveys open at the start of each quarter — Jan, Apr, Jul, Oct." The two views derive "is the survey open?" from different sources.

Fixing all five as one wave produces the unifying invariant: **every "is X available?" check on the frontend must resolve to the same backend-authoritative answer, and every tab-route target function must exist at bundle time.**

---

## 2. Scope

### In scope

1. **Zone 1** — Monthly Trend chart: add explicit `color` properties on the `datasets` array in `src/member_view.html` (and any mirror site in `steward_view.html`) so the chart never falls through to the theme-dependent palette.
2. **Zone 2** — Radar chart Section Balance: verify backend section aggregation first, then fix either the threshold check or the aggregation (decision made during implementation based on actual data). Clean up the misleading empty-state message.
3. **Zone 3** — `renderQAForum is not defined`: investigate the failure mode (rename vs load order), apply the minimum correct fix (rename the stale call site OR add a `typeof` guard OR force load order).
4. **Zone 4** — Survey period state coherence: add `periodActive: boolean` + `periodName: string` to the backend response(s) that both views consume, force both views to derive "is there a survey open?" from those backend fields, and remove the misleading "Jan, Apr, Jul, Oct" hardcoded message from `src/member_view.html`.
5. **Zone 5** — New `G-TAB-ROUTES-SAFE` integrity guard in `test/spa-integrity.test.js` that parses every `case '…':` in `_handleTabNav` and asserts the routed function is defined somewhere in the concatenated src bundle.
6. Version bump v4.55.3 → v4.55.4, CHANGELOG entry, per-zone commits, push + clasp DDS prod, DDS→SB per-wave sync.

### Out of scope (explicitly deferred)

- Broken image icon near "How to File a Grievance" (likely a missing emoji fallback or asset — investigate opportunistically during implementation, otherwise defer).
- Full v4.55.2 DDS→SB catchup sync (Option X schedules this for after all sub-projects ship).
- Any refactor of `_palette()` or `ChartHelper` beyond what Zone 1 strictly needs.

---

## 3. Architecture & Zones

### Zone 1 — Monthly Trend chart colors (#2)

**File:** `src/member_view.html:~5372` (Monthly Trend `datasets` definition). Grep during implementation for any mirror render site.

**Current code shape:**
```javascript
datasets: [
  { label: 'Filed', data: monthData.filed },
  { label: 'Resolved', data: monthData.resolved }
]
```

**Change:** Add explicit `color` properties so the chart never calls `_palette()`.
```javascript
datasets: [
  { label: 'Filed', data: monthData.filed, color: '#f59e0b' },   // amber
  { label: 'Resolved', data: monthData.resolved, color: '#10b981' } // green
]
```

**Why these colors:** the legend already says "Filed" is orange/amber and "Resolved" is green per the screenshot. Matching the legend means no visual shock for existing users. `#f59e0b` (Tailwind amber-500) and `#10b981` (Tailwind emerald-500) are stable hex values that do not depend on any CSS variable or theme state.

**Sanity check during implementation:** read `createLineChart` in `src/index.html` and verify nothing silently overrides an explicit `ds.color` property.

### Zone 2 — Radar section threshold (#4)

**Files:** `src/member_view.html:~3799–3803` (`_surveyViewRadar` section count check) + possibly `src/08e_SurveyEngine.gs` (section aggregation).

**Diagnostic step (required before writing the fix):** read the code path that builds `data.sections` from survey responses. Two branches:

- **Branch A — threshold wrong:** backend correctly aggregates sections but the frontend guard is `sectionList.length < 3` which is too aggressive. Fix: change to `sectionList.length < 2` or eliminate the minimum and let the radar render with however many sections exist.
- **Branch B — aggregation wrong:** backend returns zero/one sections because it groups questions by the wrong column, or because the schema changed and the section column was renamed. Fix: repair the aggregation in `08e_SurveyEngine.gs` to group by the current section column correctly.

**Empty-state message:** regardless of branch, replace the generic "Need 3+ sections for radar" with one of:
- "Survey has no section groupings — radar unavailable." (branch B case)
- "Radar chart needs at least N distinct sections." (branch A case, N from the new threshold)

**Blast radius:** 1 file if branch A, 2–3 files if branch B.

### Zone 3 — `renderQAForum is not defined` (#7)

**Files to investigate:** `src/index.html:~4139,4167` (tab router `case 'qaforum':`), `src/member_view.html:~6926` (where Explore agent reported the function IS defined), `src/26_QAForum.gs` (backend), git log for any Wave 31 renames.

**Diagnostic step:** confirm the current function name at `src/member_view.html:6926`. Grep for any other `renderQAForum*` references. Verify whether the function is inside an IIFE that shadows the view scope.

**Fix branches:**

- **Branch A — name mismatch:** the tab router invokes `renderQAForum` but the current function is named `renderQAForumPage` (or similar). Fix: rename both tab router call sites.
- **Branch B — load order:** the function exists but isn't registered before `_handleTabNav` runs. Fix: either wrap the call site in `typeof renderQAForum === 'function'` with a one-shot deferred retry, or hoist the function declaration, or load member_view.html before index.html's router is wired.
- **Branch C — scope isolation:** the function is defined inside a closure that the tab router can't reach. Fix: export the function to `window` or to a shared registry.

**Decision point:** implementation-time. Whichever branch the evidence supports gets the fix. Default assumption: Branch A (name mismatch, because Wave 31's Q&A editing work is the most recent touch to this area).

### Zone 4 — Survey period state coherence (#8 + survey open/close mismatch)

**Files:** `src/21_WebDashDataService.gs` or `src/21d_WebDashDataWrappers.gs` (backend response shape), `src/steward_view.html:~4212` (steward survey tracking empty state), `src/member_view.html:~2823` (member survey page empty state + hardcoded quarterly message).

**Invariant this zone establishes:**
> Both the member survey page and the steward Survey Tracking tab must derive "is there an active survey period?" from the same field in the same backend response. No client-side month-modulo logic, no separate paths.

**Backend change:** wherever the survey-related `data*` endpoints currently compute or expose "active period" info, normalize the response to include:
- `periodActive: boolean` — true iff a Period row exists with status = `Active`
- `periodName: string | null` — the human name of the active period (e.g., `"Q2 2026 Member Satisfaction Survey"`) or null if none

**Frontend changes:**
- `src/member_view.html:~2823`: replace any local quarterly-window check with a direct `if (!surveyData.periodActive)` test.
- `src/member_view.html:~2827`: delete the hardcoded "Surveys open at the start of each quarter — Jan, Apr, Jul, Oct" message. Replace with "No survey period is currently active. Your steward will notify you when the next survey opens."
- `src/steward_view.html:~4212`: split the empty state into two messages — one for `!periodActive` ("No survey period is currently active. Open a new period from the Survey Admin panel."), one for `periodActive && total === 0` ("Period is active but no members are assigned in your current scope — expand the scope filter above.")

**Consistency test:** the new `test/survey-period-state.test.js` file asserts that for any given backend state, both views derive the same boolean.

### Zone 5 — Tab router integrity guard (prevents #7 class)

**File:** `test/spa-integrity.test.js`

**New describe block `G-TAB-ROUTES-SAFE`:**

- Read `src/index.html` as a string.
- Extract the body of `function _handleTabNav(` via bracket-depth parsing.
- Regex-find every `case '…':` inside that body and capture the subsequent function call expression.
- Build the set of routed function names.
- Concatenate all `src/*.html` and all `src/*.gs` files into one string.
- For each routed function name, assert it matches `/function\s+NAME\s*\(|NAME\s*=\s*function|var\s+NAME\s*=\s*function|const\s+NAME\s*=/` somewhere in the concatenated bundle.
- Include an explicit test that routes which include `renderQAForum` (by whatever final name) are defined — the regression case for #7.

This guard, if it had existed, would have caught Wave 31's regression at pre-commit time.

---

## 4. Data Flow & Invariants

**Zone 4 is the largest data-flow change**, so it deserves the invariant spelled out:

```
[Survey Periods sheet]
        │
        ▼
 [getSurveyPeriod() — single source of truth for period status]
        │
        ▼
 [backend normalizer: {periodActive, periodName, ...}]
        │
        ▼
 [dataGetSurveyQuestions / dataGetStewardSurveyTracking / ...]
        │
        ▼
 [member_view.html and steward_view.html both read periodActive + periodName]
        │
        ▼
 [rendered UI: consistent across both views]
```

No view derives "period active" from any other source — not from hash, not from month, not from cache directly, not from a client-side fallback. Both views trust the backend field.

---

## 5. Error Handling

- Zone 1 helper: no error handling needed. Hardcoded string literals, no runtime path.
- Zone 2: if aggregation is repaired, the render path stays identical. Edge case: what if the survey has 0 sections? Empty-state message handles it.
- Zone 3: if rename is the fix, no runtime error path. If guard is the fix, the guard short-circuits with a user-friendly "Q&A Forum unavailable" toast instead of a crash.
- Zone 4: if backend returns `periodActive: undefined` (shouldn't, but defend), treat as `false`. If `periodName` is null when `periodActive` is true, display a generic "Survey is open" label.
- Zone 5: test-only, no runtime paths.

---

## 6. Test Plan

| File | Zone | Count | Key cases |
|---|---|---|---|
| `test/radar-section-threshold.test.js` (new) | 2 | ~5 | threshold fires at <N sections; passes at ≥N sections; handles empty `data.sections` object; handles undefined `data.sections`; differentiates "no sections" vs "1–(N-1) sections" empty-state messages |
| `test/survey-period-state.test.js` (new) | 4 | ~7 | backend returns `periodActive: true` + `periodName` when Period row has status=Active; `periodActive: false` when no period or status=Archived; member view and steward view derive the same boolean for the same backend state; no hardcoded quarterly month check remains in `src/member_view.html`; `dataGetStewardSurveyTracking` distinguishes "no period" from "no members assigned"; empty-state messages match the backend state; regression guard that legacy quarterly logic is gone |
| `test/spa-integrity.test.js` additions (G-TAB-ROUTES-SAFE) | 5 | ~3 | every `case '…':` in `_handleTabNav` routes to a function defined in the concatenated src bundle; explicit assertion that `renderQAForum` (or its final name) is defined; assertion that the Monthly Trend `datasets` array contains explicit `color` properties (covers Zone 1) |

Plus implicit coverage: Zone 1's fix is asserted by the G-TAB-ROUTES-SAFE extension (color literal regex), and Zone 3's fix is asserted by the same guard's function-defined check.

---

## 7. Deployment & Rollback

**Commit sequence (one wave, six commits):**

1. **Zone 1** — Monthly Trend color literals in `src/member_view.html` (+ mirror if present).
2. **Zone 2** — Radar section threshold + aggregation fix (whichever branch) + `radar-section-threshold.test.js`.
3. **Zone 3** — `renderQAForum` tab router fix (rename or guard).
4. **Zone 4** — Survey period state coherence: backend normalizer + both views + `survey-period-state.test.js`.
5. **Zone 5** — `G-TAB-ROUTES-SAFE` describe block + color-literal guard in `test/spa-integrity.test.js`.
6. **Version bump** v4.55.3 → v4.55.4 in `src/01_Core.gs` + CHANGELOG entry.

**Pipeline:** lint → test:guards → test:unit → build:prod → git push origin Main → clasp push to SolidBase prod (`REDACTED_DDS_SCRIPT_ID…`).

**SolidBase per-wave sync (Option X):** after DDS prod clasp, copy the modified files from DDS to SB with scrub rules applied. One SB commit, one SB push, one SB clasp. Files touched in SB will be roughly the same list.

**Rollback per zone:**
- Zone 1: safe (revert the color literals).
- Zone 2: safe if threshold-only; higher risk if aggregation — mitigated by `test:unit` running before push.
- Zone 3: safe (revert rename or guard).
- Zone 4: highest risk because backend response shape changes. Mitigation: the new fields are additive — the old response still parses. Revert restores pre-wave state; no consumer can break on missing new fields because the old code doesn't read them.
- Zone 5: test-only, safe.

**Version bump: v4.55.3 → v4.55.4** (patch — bug fixes only, no feature additions).

---

## 8. Open Questions / Implementation-Time Decisions

- Zone 2 branch (A vs B) — decide after reading `08e_SurveyEngine.gs`.
- Zone 3 branch (rename vs guard vs hoist) — decide after grepping `renderQAForum` call sites and definitions.
- Whether to fix the broken image icon near "How to File a Grievance" opportunistically — decide after a quick grep during Zone 1 or Zone 2 work; include it only if it's a 1-line fix.

---

## 9. Success Criteria

- All five production bugs fixed when verified in a live DDS prod session.
- Test suite count grows by ~15 (from 3719 → ~3734) plus 3+ G-TAB-ROUTES-SAFE assertions.
- Both member view and steward view return identical answers to "is there an active survey period?" for every backend state.
- `G-TAB-ROUTES-SAFE` fires and catches the regression on any future wave that introduces a dangling tab router reference.
- DDS prod clasped successfully; SolidBase synced and clasped without issue.
