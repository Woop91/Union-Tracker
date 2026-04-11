# Sub-Project C — Mobile Audit & Horizontal-Scroll Kills

**Date:** 2026-04-11
**Target release:** DDS v4.55.7 (patch wave)
**Status:** Compact spec, fast-path execution per user's "all/continue" directive
**Closes:**
- Issue #3 (Survey Results card — one-word-per-line text wrap)
- Access Log tab not mobile-friendly
- Wide-table horizontal scroll on Org Health Tree / Explorer area (`minWidth: 1200px`)
- Partial: the general "reduce sideways-scroll" directive — scoped to the biggest offenders, not a full per-tab audit

## 1. Motivation

Investigation showed the mobile issues concentrate on three specific culprits rather than a systemic layout bug: (a) one flex layout missing `flex: 1` / `min-width: 0` on its text column, (b) a specific admin table not hiding columns on narrow screens, and (c) one hardcoded `minWidth: 1200px` on a data table inside steward_view. The viewport meta tag, body container, and global CSS are all correctly configured — no global fix is needed. Sub-project C kills the three concrete offenders and adds a lightweight spa-integrity guard (`G-NO-WIDE-TABLES`) to catch `minWidth: 1200px`-style regressions in the future.

## 2. Scope

**Zone 1 (#3) — Survey Results card flex fix.** `src/member_view.html:3614-3628` renders a flex row with text on the left and a "Take Survey" button on the right. The text `<div>` has no `flex: 1` or `min-width: 0`, so on narrow screens the text shrinks below content width, causing one-word-per-line wrap. Fix: add `flex: 1`, `min-width: 0`, and a media-query fallback to `flex-direction: column` at ≤ 480px.

**Zone 2 — Access Log mobile column hiding.** `src/steward_view.html:8906-8960` renders a 6-column table (Time, Event, User, Sheet, Field, Details) inside `overflowX: 'auto'`. Fix: on screens ≤ 640px, hide `Sheet` and `Field` columns via CSS so the remaining 4 fit without horizontal scroll. Add a small "swipe for more columns" hint if the `overflowX` still fires (it shouldn't after hiding 2 cols, but defensive).

**Zone 3 — Wide table minWidth reduction.** `src/steward_view.html:6406` has `minWidth: '1200px'` on a data table inside the Org Health Tree / Explorer surface. This forces horizontal scroll below 1200px viewport (every phone and small laptop). Fix: change `minWidth: '1200px'` to `minWidth: '900px'` (still wide enough for desktop-style data display, narrow enough that only the smallest phones trigger horizontal scroll). If the table truly needs 1200px of content, consider a per-column-hide CSS approach matching Zone 2.

**Zone 4 — Sidebar scrollbar investigation.** Per explore report, no *separate* header scroll exists — `.sidebar-nav` has `overflow-y: auto` on the whole sidebar, which is expected for long nav lists. The user's screenshot showing a scrollbar inside the "SolidBase" header block is likely (a) the main sidebar scroll fired at a viewport height where the header's new padding pushes content past 100vh, OR (b) a visual artifact from a specific browser. Fix: verify during implementation by reading the CSS; if it's a real secondary scroll, fix it. If not, document the finding in the CHANGELOG and close the item.

**Zone 5 — `G-NO-WIDE-TABLES` integrity guard.** Add a test in `test/spa-integrity.test.js` that greps all `src/*.html` for `minWidth:\s*['"]?1[0-9]{3}` and fails if any match is found with a value ≥ 1100px. This prevents the Zone 3 regression class from shipping again. Whitelist can be added if any legitimate desktop-only surface needs wider content.

## 3. Deploy

Per-zone commits, version bump v4.55.6 → v4.55.7, push + clasp DDS prod. Final sub-project — the combined DDS→SB catchup sync happens after this wave ships per Option Y.

## 4. Success criteria

- Survey Results card on a 375px screen: text wraps at word boundaries within a full-width column, not one-word-per-line.
- Access Log on a 375px screen: 4 columns fit without horizontal scroll.
- Org Health Tree / Explorer area on a 375px screen: either fits without horizontal scroll, or narrows the minWidth to something phones can handle.
- Spa-integrity guard catches any future `minWidth: >=1100px` additions.
- DDS v4.55.7 live on prod.
