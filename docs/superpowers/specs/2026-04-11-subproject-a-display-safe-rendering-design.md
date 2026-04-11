# Sub-Project A — Display-Safe Rendering Boundary

**Date:** 2026-04-11
**Target release:** DDS v4.55.3 (patch wave), synced to SolidBase after DDS merges
**Status:** Design approved, awaiting user review before writing-plans handoff
**Closes issues:** #1 (Avg Resolution by Steward chart Y-axis shows raw floats), #5 (Access Log shows SHA-256 hashes instead of names), #9 (Literal `\n` escape sequences in Know Your Rights / FAQ / How to File a Grievance content)

---

## 1. Motivation

Three visible production defects share one structural cause: raw backend values reach the DOM without passing through a display-safe rendering layer. Each looks like a separate bug; together they prove that the codebase lacks a consistent boundary between "data as produced by the sheet" and "data as shown to the user."

- **#1** — The grievance chart's Y-axis plots `Date` serial numbers (e.g., `46076.51939715278` ≈ March 2026) as steward group keys. After verification during plan-writing: `_buildColumnMap` (line 1995) and `_findColumn` (line 2013) both use exact, case-insensitive matching — there is no fuzzy/substring logic to blame. The real root cause is that the "Assigned Steward" column in the live Grievance Log contains numeric values (serials) for at least some rows — either because an admin wrote a timestamp into the wrong column, an import mapped dates there, or the column was repurposed. We cannot control what's in the cell; we *can* control what reaches the chart. The fix is a downstream type guard in the aggregation loop that rejects numeric/Date values before they become phantom grouping keys. One real steward (Jacob Sanders) renders correctly; every other bucket is a phantom group keyed by a unique timestamp serial that the guard will drop.
- **#5** — The Access Log renders user hashes (64 hex chars) in the user column. The backend stores hashes to avoid raw PII retention, but the frontend shows them verbatim instead of resolving them back to friendly names at display time.
- **#9** — The Resources sheet contains literal `\n` text (two characters: backslash + lowercase n) in the body of multiple seeded rows because the seed source strings in `src/10b_SurveyDocSheets.gs` and `src/07_DevTools.gs` use `'\\n'` (which JS evaluates to `\n`-as-text, not as a newline) inside single-quoted literals. Because production sheets are already populated from this seed, any pure code fix must also handle existing bad data without forcing a destructive re-seed.

Fixing all three as one wave produces the unifying invariant: **no raw sheet value reaches the DOM without passing through a display-safe helper.**

---

## 2. Scope

### In scope

1. Repair literal-`\n` seed strings in `src/10b_SurveyDocSheets.gs` (`starterRows` block, ~6 rows) and `src/07_DevTools.gs` (`seedResourcesData`, 9+ rows).
2. Add a newline-tolerant renderer helper (`renderSafeText`) and adopt it at every Resources rendering site in `src/member_view.html` and `src/steward_view.html`.
3. Add a type guard in the steward resolution aggregation loop in `src/21_WebDashDataService.gs` (around line 3074) that rejects `Date`/number/numeric-string values before they become grouping keys. (No `_findColumn` change — verified during plan-writing that it is already exact-match only, so there is no precedence bug to fix there.)
4. Add a `displayUser(rawId)` helper that resolves SHA-256 hash, email, or empty → friendly name, and adopt it in the Access Log renderer. Backend ships a `hashMap` payload alongside access log rows to avoid N×lookup cost.
5. Tests: three new test files (~24 tests total) plus one regression guard in `test/spa-integrity.test.js`.
6. Per-zone commits, version bump to v4.55.3, DDS→SolidBase sync following the documented workflow.

### Out of scope (explicitly deferred)

- The broken image icon near "How to File a Grievance" from the screenshot — belongs to sub-project B (chart / tab / empty-state bugs).
- The "survey shows Open but opening it says No Survey Open" mismatch — sub-project B.
- Rich text (markdown) support in resource content — a separate design, not required here.
- Non-SHA-256 legacy hash formats — per session memory, the Access Log has only ever used SHA-256.
- Refactoring `_findColumn`'s broader call sites beyond the precedence fix — targeted improvement only.

---

## 3. Architecture & Zones

The design is partitioned into three isolated zones. Each zone has its own helper (if applicable), its own adoption sites, its own tests, and its own commit. A zone can be reverted without touching the other two.

```
[sheet data / raw string]
        │
        ▼
 [backend extraction]  ← type-guarded in Zone 2
        │                    and Zone 3 (hashMap payload)
        ▼
 [wire to frontend via google.script.run]
        │
        ▼
 [display-safe boundary]  ← renderSafeText / displayUser (Zones 1, 3)
        │
        ▼
 [DOM / chart axis labels]
```

The invariant the wave guarantees: **no raw sheet value reaches the DOM without passing through one of the boundary helpers.** This is the single rule code review can check.

### Zone 1 — Newline-tolerant rendering (#9)

**New helper** `renderSafeText(str)` in `src/index.html`:

- Input: any value (string, null, undefined, number, etc.).
- Behavior: coerces to string, normalizes `\r\n` → `\n`, replaces literal `\n` (backslash+n, two chars) and real `\n` newlines with `<br>` nodes, HTML-escapes everything else.
- Output: a `DocumentFragment` (not an innerHTML string) so callers can `appendChild` without opening an XSS path.
- Contract: never throws; null/undefined → empty fragment; non-string → stringified first.

**Why in `index.html`**: GAS single-global-scope pattern — `index.html` is the outer shell that wraps every view (`member_view.html`, `steward_view.html`), so helpers defined there are reachable from every view. Wave 31 used this pattern for `fuzzyMatch`; we follow the same convention.

**Adoption sites** (to be enumerated precisely during plan-writing):
- Every `textContent = resource.body`-style assignment in `member_view.html` Resources rendering.
- Every equivalent in `steward_view.html` Know Your Rights hub.
- The "How to File a Grievance" card.

**Seed cleanup**:
- `src/10b_SurveyDocSheets.gs` starterRows block (lines 1254–1259): replace `'\\n'` with `'\n'`.
- `src/07_DevTools.gs` `seedResourcesData` (lines 1737+): same replacement across all affected rows.

Existing production sheets already contain the literal `\n` characters. `renderSafeText` handles them transparently at display time, so no destructive data migration is needed. A future admin-triggered re-seed would produce clean data from then on.

### Zone 2 — Steward value type guard (#1)

**Verified during plan-writing**: `_buildColumnMap` and `_findColumn` already perform exact, case-insensitive matching. There is no fuzzy-matching bug to fix in the lookup layer. The float values reaching the chart Y-axis are not the result of a wrong column being selected; they are real values stored in the correct "Assigned Steward" column of at least some rows of the live Grievance Log (admin data-entry drift, import mapping, or column repurposing).

**Type guard** in the aggregation loop around `src/21_WebDashDataService.gs:3074` (`stewardResolution` accumulation): before using `rRec.steward` as a grouping key, validate that it "looks like a name." Reject and coerce to `'unassigned'` if:

- It is falsy (empty string, null, undefined).
- It matches the numeric-only regex `/^-?\d+(\.\d+)?$/` (catches serials like `46076.519`).
- It matches a date-like pattern (`/^\d{4}-\d{2}-\d{2}/` or starts with a weekday name like `"mon "`, `"tue "`, etc. — catches `String(dateObject)` output).
- It is longer than a reasonable name length cap (e.g., > 120 chars).

Note: `rRec.steward` is already lowercased via `String(_getVal(...)).trim().toLowerCase()` at line 2260, so the guard operates on a lowercase string. The guard is placed at the point where the key is first used (line 3074), not inside `_buildGrievanceRecord`, so other consumers of `.steward` are unaffected.

This does not recover the "lost" cases from affected rows — those cases just get bucketed under `'unassigned'` along with any legitimately unassigned cases. But it cleans up the chart so legitimate stewards render correctly, and it prevents one bad row from creating a phantom bucket per unique timestamp.

### Zone 3 — Access Log display resolution (#5)

**New helper** `displayUser(rawId, hashMap)` in `src/index.html`:

- Input: raw value from Access Log user column + an optional `hashMap` object `{hash → name}`.
- Resolution order:
  1. If empty/null → `"—"` (em dash, matches existing empty-user convention).
  2. If looks like a SHA-256 hash (64 hex chars) → look up in `hashMap`; on miss, show abbreviated hash (first 8 chars + `…`).
  3. If looks like an email → run through the view's existing `findUserByEmail`-equivalent resolver.
  4. Else → return stringified as-is.

**Backend change** in `src/21d_WebDashDataWrappers.gs` (or wherever the access log payload is assembled): ship a lightweight `{hashMap: {hash → name}}` alongside the log rows. Single batched resolution pass on the backend, not per-row.

**Adoption site**: the Access Log renderer in whichever admin pane currently shows it (to be pinned down in plan-writing — likely `steward_view.html` admin section).

---

## 4. Data Flow & Invariants

Every display path flows through the boundary:

1. Backend reads sheet data via `_getCachedSheetData` / `_buildGrievanceRecord`.
2. Extraction respects the hardened `_findColumn` precedence (Zone 2).
3. Backend assembles response including `hashMap` for access log (Zone 3).
4. Frontend receives payload via `google.script.run`.
5. Before any DOM write, values flow through `renderSafeText` (text content) or `displayUser` (user identifiers).
6. DOM/chart receives only display-safe strings or fragments.

**Invariant**: the only textual data reaching the DOM from sheet-sourced content is either (a) a DOM fragment produced by `renderSafeText`, or (b) a string returned from `displayUser`, or (c) already-safe primitives like numeric KPI values. Lint/review can check for raw `.textContent = rec.something` assignments against resource body or user-identifier fields.

---

## 5. Error Handling

Both new helpers are **pure and total** — they accept any input and always return a safe value.

| Helper | Input | Output |
|---|---|---|
| `renderSafeText(null)` | — | empty `DocumentFragment` |
| `renderSafeText(undefined)` | — | empty `DocumentFragment` |
| `renderSafeText(42)` | — | fragment with `"42"` text node |
| `renderSafeText("<script>x</script>")` | — | fragment with text node `"<script>x</script>"` (escaped) |
| `displayUser(null)` | — | `"—"` |
| `displayUser("")` | — | `"—"` |
| `displayUser(nonString)` | — | `String(nonString)` then fall through resolution |

The "never throw" contract is tested explicitly — one test per helper asserting every primitive type (null, undefined, number, boolean, object, array) returns a safe value without throwing.

`_findColumn` retains its existing null-return contract for "no match"; the change is precedence, not error behavior. No caller-side changes are required.

---

## 6. Test Plan

Target: ~24 new tests across three new files plus one regression guard.

| File | Zone | Count | Key cases |
|---|---|---|---|
| `test/render-safe-text.test.js` (new) | 1 | ~10 | literal `\n` → `<br>`; real `\n` → `<br>`; mixed `\\n` and `\n`; empty string; `<script>` injection stays escaped; `&amp;` double-escape regression; multi-paragraph Weingarten content; returns `DocumentFragment` type; null/undefined → empty fragment; preserves leading/trailing whitespace |
| `test/steward-name-guard.test.js` (new) | 2 | ~8 | empty string → `unassigned`; `Date` object stringified → `unassigned`; numeric-only serial string (`"46076.519"`) → `unassigned`; ISO date string (`"2026-03-18..."`) → `unassigned`; weekday-prefixed date string (`"wed mar 18 2026 ..."`) → `unassigned`; overlong string (> 120 chars) → `unassigned`; legitimate name (`"jacob sanders"`) passes through unchanged; legitimate email (`"jsanders@ward.org"`) passes through unchanged |
| `test/display-user.test.js` (new) | 3 | ~6 | 64-hex-char input resolves via `hashMap`; email input resolves via `findUserByEmail`; empty input → `"—"`; hash not in `hashMap` → abbreviated hash; non-hex 64-char falls through to email branch; missing `hashMap` prop → graceful fallback |

**Regression guard** in `test/spa-integrity.test.js`: assert that `renderSafeText` and `displayUser` are defined on the view's global scope at load time. This catches the Wave 31 `renderQAForum is not defined`-style failure for these new helpers.

**Edge cases explicitly out of scope**:
- Markdown (`**bold**`, `[links](url)`) — not supported.
- Windows CRLF (`\r\n`) — handled via normalize inside the helper.
- Emoji in steward names — works fine, no special handling.
- Non-SHA-256 hashes — never existed in this system.

---

## 7. Deployment & Rollback

### Commit sequence (one wave, five commits)

1. **Zone 1**: `renderSafeText` helper + adoption in member_view/steward_view + seed string cleanup in `10b_SurveyDocSheets.gs` and `07_DevTools.gs`.
2. **Zone 2**: steward-name type guard in the aggregation loop in `21_WebDashDataService.gs`.
3. **Zone 3**: `displayUser` helper + `hashMap` payload wiring + Access Log adoption.
4. **Tests**: three new test files + `spa-integrity.test.js` guard additions.
5. **CHANGELOG + version bump** to v4.55.3.

### Pipeline (standard)

- `npm run lint && npm run test:guards && npm run test:unit && npm run build:prod`
- Commit via pre-commit hook (runs the same pipeline).
- Push to `origin/Main`.
- `clasp push` to DDS dev 1 (`1cJEyS0Ni2LTICNP_478yAVr2VmAlMGJkpDN3m3OKuJK2rEV8qwYrHRL3`).
- Smoke verify in dev: Resources tab renders newlines correctly, Avg Resolution by Steward chart shows real names, Access Log shows friendly user identifiers.
- `clasp push` to DDS prod (`REDACTED_DDS_SCRIPT_ID`) once dev passes.
- Sync DDS → SolidBase per documented sync workflow. SHOW_GRIEVANCES is already present in SB — no feature-toggle exclusions needed for this wave.
- `clasp push` to SB prod (`1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`).

### Rollback plan

Per-zone commits enable surgical revert:

- **Zone 1 rollback** — safe. `renderSafeText` is pure and has no other callers; removing it only reverts adoption sites and leaves existing bad-data rendering as before the fix (i.e., restores pre-wave behavior, doesn't make things worse).
- **Zone 2 rollback** — safe. The type guard is scoped to a single loop iteration (`stewardResolution` accumulation) and cannot affect any other consumer of `rRec.steward` (which is unchanged upstream). Reverting the zone restores the pre-wave buggy chart but breaks nothing else.
- **Zone 3 rollback** — safe. Affects only the Access Log tab.

Version bump: v4.55.2 → v4.55.3 (patch — bug fixes only, no feature additions).

---

## 8. Open Questions / Implementation-Time Decisions

These are not design blockers; they're details to resolve during plan-writing:

1. Exact set of Resources rendering call sites (grep for `textContent` + resource body pattern in both view files).
2. Exact location of the Access Log renderer — likely `steward_view.html` admin section; confirm during plan-writing.
3. Whether `displayUser` should accept `hashMap` as a parameter or read it from a view-global. Recommendation: parameter, for testability.
4. Whether the seed-string cleanup should also ship a one-shot admin-triggered "re-seed Resources" menu action for users who want clean data in their existing sheets. Default: no — `renderSafeText` makes the existing data render correctly, so this is optional polish, not required by the wave.

---

## 9. Success Criteria

The wave is successful when:

- All three screenshot defects (#1, #5, #9) are visibly fixed in DDS prod.
- 72 test suites remain green; new test count is ≥ 3687 + 24 = 3711.
- DDS and SolidBase prod clasp deploys complete without regression.
- No new violations of the display-safe rendering invariant introduced by this wave.
- Zone 2's type guard does not affect any non-steward-chart consumer of `rRec.steward` (guarded by running the full `test:unit` suite before clasp push, plus the targeted `test/steward-name-guard.test.js` suite).
