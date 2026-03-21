# UI Review — v4.30.2

**Date**: 2026-03-21
**Scope**: All 10 HTML views + 2 server-side UI files (03_UIComponents.gs, 04b_AccessibilityFeatures.gs)
**Files reviewed**: ~1.77MB of frontend code

---

## Summary

The webapp is a sophisticated SPA running inside a Google Apps Script iframe with no framework — all hand-rolled with a clean `el()` DOM builder and IIFE modules. Performance engineering is excellent (preloading, SWR, circuit breaker, LRU pane cache, LazyList). XSS hygiene is strong (`safeText()`, `el()` textContent, server-side `escapeHtml()`). The theme system is comprehensive (8 visual styles, dark/light per style, cyberpunk multi-neon).

**The primary gap is accessibility (WCAG compliance)** — critical for a union member-facing app. Secondary concerns include CSS conflicts, code duplication, and UX inconsistencies.

---

## 1. CRITICAL — Accessibility (WCAG)

### 1.1 No ARIA Landmarks
- `index.html`: Sidebar `<nav>` has no `aria-label`. Main content has no `role="main"` or `<main>` element. Bottom nav has no `role="navigation"`.
- No skip-to-content link exists anywhere.

### 1.2 Non-Semantic Interactive Elements
- Sidebar items and bottom nav items are `<div>` with `onClick` — not `<button>` or `<a>`. Lack `role="button"`, `tabindex="0"`, and keyboard handlers.
- Star rating widget (`member_view.html` ~line 411-427): `<span>` elements with `onClick`. No `role="radiogroup"`, not keyboard-operable.
- Toggle switch (`steward_view.html` ~lines 2614-2691): Hidden checkbox with `opacity:0, width:0, height:0` — not focusable.
- Checklist (`steward_view.html` ~lines 4358-4386): Unicode chars (checkboxes) instead of `<input type="checkbox">`.

### 1.3 Missing Focus Indicators
- `styles.html`: Every form input has `outline: none` (~lines 1987, 2068, 2081, 2172, 2327, 2383, 2449, 2595) with only `border-color` change as replacement. Buttons have no `:focus-visible` styles.
- No `.sr-only` / `.visually-hidden` utility class defined anywhere.

### 1.4 Missing ARIA States
- No `aria-expanded` on any collapsible element across steward_view, member_view, org_chart, poms_reference.
- No `role="tablist"` / `role="tab"` / `aria-selected` on any sub-tab bar (8 instances in steward_view alone).
- Notification bell has no `aria-label` or `aria-live` region for badge count.

### 1.5 Form Label Associations
- `member_view.html` profile form: `<label>` elements without `for`/`id` attributes.
- `grievance_form.html`: Text inputs lack `<label>` associations (visual `.lbl` spans are not semantic).
- `poms_reference.html`: Search input has no `<label>` or `aria-label`.

### 1.6 Org Chart Font Sizes
- Multiple elements use `0.44rem` (7px), `0.48rem`, `0.5rem` — below WCAG 1.4.4 minimum.

### 1.7 Dead Accessibility Code
- `03_UIComponents.gs` ~lines 858-859: `highContrast` and `largeText` ADHD settings are defined but never read or applied.
- `04b_AccessibilityFeatures.gs` contains zero `aria-*` or `role` attributes. Named "accessibility" but provides comfort/productivity tools, not WCAG compliance.

### 1.8 Reduced Motion Gaps
- `styles.html`: Survey animations (surveyFloat, surveyBounce) not included in `prefers-reduced-motion` block.
- `.spinner` only reduces `animation-duration` instead of `animation: none` (inconsistent with other elements).
- No `@media (forced-colors: active)` or `@media (prefers-contrast: more)` support.

### 1.9 Missing Focus Management
- No focus trapping in modal overlays (idle-logout modal, workload sub-category modal).
- No focus restoration after async content swaps — keyboard users lose their place.
- Escape key doesn't dismiss the idle-logout modal.

### 1.10 esign.html
- Canvas has no `aria-label` or keyboard alternative for drawing.
- `user-scalable=no` in viewport meta — accessibility concern (disables pinch-zoom).

---

## 2. HIGH — Security

### 2.1 POMS Self-XSS
- `poms_reference.html` ~line 792: User search query inserted via innerHTML without escaping: `'No results for "'+P.q+'"'`. Should use `textContent`.

### 2.2 PAGE_DATA Injection Pattern
- `index.html` ~line 37: `var PAGE_DATA = <?!= pageData ?>` relies on server `JSON.stringify()`. A `<script type="application/json">` + client-side parse would be safer defense-in-depth.

### 2.3 innerHTML Anti-Patterns
- `index.html` ~line 3118: `container.innerHTML += ...` destroys existing DOM nodes and event listeners. Should use `insertAdjacentHTML` or `el()`.
- `member_view.html` ~line 1340: `innerHTML` with interpolated content. Uses `safeText()` but DOM construction preferred.

### 2.4 Strengths (preserve these)
- `safeText()` / `el()` with `textContent` — XSS-safe by design.
- Server-side `escapeHtml()` covers `&`, `<`, `>`, `"`, `'`, backtick.
- `JSON.stringify()` correctly used for JS context embedding (documented at CR-XSS-6).
- Session token `encodeURIComponent`'d in URLs.

---

## 3. HIGH — CSS Conflicts & Dead Code

### 3.1 Duplicate Dark Theme Blocks
- `styles.html`: Brutalist dark mode defined twice with conflicting values:
  - Descendant selector `[data-theme-mode="dark"] .theme-brutalist` → bg `#0a0a08` (~line 938)
  - Compound selector `[data-theme-mode="dark"].theme-brutalist` → bg `#1a1a1a` (~line 1042)
- Same duplication for retroOS and comic dark themes.

### 3.2 Conflicting Timeline Styles
- `styles.html` ~lines 1621-1661 (flex-based) vs ~lines 3014-3018 (absolute-positioned) both define `.timeline-line`, `.timeline-dot`, `.timeline-date`. Later block silently overrides.

### 3.3 will-change / Animation Mismatch
- `cpScanBeam` animates `top` (~line 434) but declares `will-change: transform` (~line 458). No GPU benefit; should use `transform: translateY()`.

### 3.4 Performance-Heavy Animation
- `heatWave` (~line 363): `filter: hue-rotate()` on fullscreen `::before` triggers repaint every frame in blob lava theme.
- No `contain` on blob lava or liquid pour animated elements (only cyberpunk CRT has containment).

### 3.5 Dead CSS
- `styles.html` ~line 117: `.page-header-left {}` — empty rule.
- ~line 1920: `.loading-text { display: none; }` — legacy compat.
- ~lines 2572-2596: `.wt-category-*` — legacy single-row category styles.

---

## 4. MEDIUM — Code Quality & DRY

### 4.1 Inline Style Explosion
- `member_view.html`: **924 inline style objects**. `style: { fontSize: '11px', color: 'var(--muted)' }` repeated 100+ times.
- `steward_view.html`: Similar density. Nearly every `el()` call has inline styles.
- `org_chart.html`: Thousands of lines of inline `style="..."` attributes.
- **Impact**: Massive duplication, hard to maintain, inflates file size.

### 4.2 Duplicated Patterns

| Pattern | Count | Locations |
|---------|-------|-----------|
| KPI strip rendering | 6+ | steward_view ~lines 187, 640, 789, 893, 3089, 3228 |
| Sub-tab bar | 8 | steward_view ~lines 1499, 1870, 2804, 3053, 4417, 4698, 4912, 5577 |
| Autocomplete/typeahead | 3 | steward_view ~lines 2837, 2965, 6089 |
| Pagination controls | 4+ | Across steward_view and member_view |
| vCard generation | 2 | member_view ~lines 1099-1110 and 1354-1365 (verbatim copy) |
| Email sending boilerplate | 3 | 03_UIComponents.gs ~lines 1731-1845 |

### 4.3 Duplicate Theme Systems
- `THEME_PRESETS` (03_UIComponents.gs ~line 590) vs `THEME_CONFIG.THEMES` (04b ~line 214): Two sources of truth.
- `showThemePresetPicker()` (03 ~line 694) vs `showThemeManager()` (04b ~line 698): Two different UIs for the same job.
- Three different button color systems: `#7C3AED` (purple), `#1a73e8` (Google blue), purple gradient.

### 4.4 Function Length
- `_renderSurveyWizard()` in member_view: ~700 lines.
- `renderSurveyResultsPage()` in member_view: ~600 lines.
- `_renderMemberDirectory()` in member_view: ~240 lines.

### 4.5 Test Runner Bypasses serverCall()
- `steward_view.html` ~lines 5252-5406: Test Runner directly uses `google.script.run` instead of `serverCall()`, bypassing retry, throttling, and circuit-breaker protection.

### 4.6 Server-Side HTML Generation
- `03_UIComponents.gs`: Mobile dashboard functions use monolithic string concatenation (~20 lines of unbroken concatenation). Hard to maintain.
- `getBulkActionsDialogHtml_()`: 110 lines of string concatenation with embedded JavaScript.

---

## 5. MEDIUM — UX Issues

### 5.1 Navigation
- 20+ tabs behind a single "More" menu in steward view — poor discoverability, not searchable.
- "More" tab has 4 different identifiers (`more`, `more_steward`, `_member_more`, `_steward_more`) — confusing.
- Tab visit analytics (`dataLogTabVisit`) fires on every click including cached tab switches — should debounce/batch.
- Theme toggle re-renders entire tab — always-fresh tabs trigger full server round-trip unnecessarily.

### 5.2 Inconsistencies
- Two pagination models: "Show more" (append) vs Prev/Next (replace).
- `alert()` used in some forms (steward_view ~lines 2469, 3157) vs inline status messages everywhere else.
- Inconsistent empty states: some use large emoji + styled text, others plain text.
- Auth view email validation too permissive: `email.indexOf('@') === -1`.

### 5.3 Missing Safeguards
- Task completion has no undo and no confirmation dialog (member_view ~lines 5724-5729).
- Grievance intake only validates title — category and description should be required (member_view ~line 791).
- Profile form has zero field validation — zip codes, states accept any text (member_view ~lines 1670-1685).
- Survey draft lost on page reload (localStorage blocked in GAS iframe).

### 5.4 Navigation Bugs
- Steward assignment success navigates to directory tab instead of contact view (member_view ~line 1174).
- Search result click for grievance: `role === 'steward' ? 'cases' : 'cases'` is a no-op ternary (steward_view ~line 4037).

### 5.5 Missing Print Styles
- Zero `@media print` in the entire stylesheet. Union dashboard likely needs printable reports, grievance summaries, member lists.

### 5.6 esign.html Gaps
- No dark mode — always light theme (inconsistent with rest of app).
- `@media print { body { display: none } }` prevents printing signed grievance.
- Canvas resize on device rotation may silently lose drawn signature.

### 5.7 Hardcoded Data
- `grievance_form.html` ~line 579: Manager names hardcoded in JavaScript.
- `grievance_form.html` ~line 597: Only 2 work locations hardcoded (Everett, Worcester).

---

## 6. LOW — Mobile

| Issue | Location |
|-------|----------|
| Member-view switch button too small (10px font, 4x10px padding) | steward_view ~line 100 |
| Star rating touch targets (20px font, no padding) | member_view ~line 413 |
| Grievance preview fixed 140px column (overflow <320px) | member_view ~line 6483 |
| CSS selects elements by inline style value — extremely fragile | org_chart ~lines 697-704 |
| Advanced search fixed 280px sidebar breaks on mobile | 03_UIComponents.gs ~line 2459 |

**Strengths (preserve)**: Breakpoints at 640/1024px, touch scroll-nav, safe-area insets, passive listeners, 44px touch targets, responsive grid with `auto-fit`/`minmax`.

---

## 7. LOW — Performance

| Concern | Location |
|---------|----------|
| SWR JSON.stringify diff is O(n) on large batch data | index.html ~line 114 |
| POMS 100KB+ inline database | poms_reference.html — ~560 lines of data |
| 386 `!important` declarations (specificity ceiling) | styles.html (mostly theme overrides — architecturally justified) |
| All font-size values in px — no responsive rem/clamp | styles.html — 172 font-size declarations |

**Strengths (preserve)**: Preload-at-parse, SWR with JSON diff, stale-callback prevention (`_navSwitchId` + `_renderGeneration`), circuit breaker with manual half-open, retry with jitter + throttle (max 4 concurrent), LRU pane cache (8 max), LazyList with IntersectionObserver (400px rootMargin), idle callbacks, layout shell preservation, Chart.js lazy loading with CDN fallback, conditional view inclusion (~300KB saved for member-only users), HTML minification (13-26% per file).

---

## Recommended Priority Order

1. **Accessibility critical fixes** (ARIA landmarks, keyboard nav, focus indicators, semantic elements)
2. **POMS XSS fix** (line 792 — `textContent` instead of `innerHTML`)
3. **CSS conflict resolution** (duplicate dark theme blocks, timeline conflicts)
4. **Extract shared components** (sub-tabs, pagination, autocomplete, KPI strip)
5. **Consolidate theme systems** (single source of truth)
6. **Add print stylesheet**
7. **Form validation improvements**
8. **Performance: cpScanBeam → transform, heatWave containment**
9. **Mobile touch target sizing**
10. **Inline style reduction** (extract common patterns to CSS classes)
