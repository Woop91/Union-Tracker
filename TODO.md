# SolidBase — TODO

A running list of deferred decisions, design reviews, and follow-up tasks.
Items are not prioritized — review and sort as needed.

---

## 🎨 Design Reviews

### [ ] Glow-bar animation options — review again with fresh eyes
**Added:** 2026-03-13  
**Context:** During the progress bar / loading indicator audit, four glow-bar animation
options were evaluated. The sweep animation was selected and implemented, but the
decision was made quickly during a session.

**Options previewed:**
| Option | Description | Status |
|--------|-------------|--------|
| Static | Centered gradient, no movement (original) | Was in use |
| **Sweep** ✅ | Bright peak travels left→right, 2.5s cycle | **Implemented** |
| Pulse | Whole bar fades in/out, 2s cycle | Not chosen |
| Breathe | Scales 85%→100% width + fade, 2.4s cycle | Not chosen |

**To review:** Open a grievance card in both steward and member views.
Compare how the sweep looks on warning (active), danger (overdue), and success (resolved) states.
Re-run the visual comparison if needed — the widget was built and can be regenerated.

**Files affected:** `src/styles.html` — `.glow-bar` + `@keyframes glowSweep`

**2026-07-21 verification:** Blocked by the production web deployment returning
HTTP 403 before application HTML loads. Re-run this visual review after public
deployment access is repaired; do not close from the standalone preview because
it does not render the production grievance-card glow states.

---

## 🔧 Technical Debt

### [ ] Raise real Google Apps Script coverage baseline
**Added:** 2026-07-21
**Current enforced baseline:** 32% statements, 25% branches, 39% functions, 33% lines

The old 70%/60% thresholds silently measured `0/0` because Jest did not
instrument `.gs` files loaded through `eval`. Version 4.56.2 added explicit
Istanbul instrumentation and set honest regression floors from the first real
full-suite measurement. Raise each floor as new tests land; never restore an
unmeasured aspirational threshold.

### [ ] Husky v10 Migration
**Added:** 2026-03-17
**Current version:** 9.1.7 (deprecation warning is forward-looking, no action needed now)
**Checked:** 2026-07-21 — npm still reports 9.1.7 as latest; migration remains future-only.

When Husky v10 is released:
- [ ] Remove `#!/bin/sh` shebang from `.husky/pre-commit`, `.husky/pre-push`, `.husky/commit-msg`
- [ ] Run `npx husky init` to regenerate the `_/` directory structure
- [ ] Verify commitlint and lint-staged still trigger correctly

---

## 💡 Feature Ideas

*(empty — add items here as they come up)*

---

## ✅ Recently Completed (for reference)

- [x] **ATLAS-001** Import complete public-safe USUnions organizing and research datasets.
- [x] **ATLAS-002** Add weekly source freshness workflow and exact package hashes.
- [x] **ATLAS-003** Replace eager 9.76 MiB JavaScript payload with 1.61 MiB compressed worker package.
- [x] **ATLAS-004** Geocode public headquarters through Census; retain explicit ZCTA fallback precision.
- [x] **ATLAS-005** Surface exact membership, website, and coordinate gaps without synthetic values.
- [x] **ATLAS-006** Remove hidden 24/50-record result caps through explicit compass labeling and pagination.
- [x] **ATLAS-007** Generate a true one-file offline export.
- [x] **ATLAS-008** Add automated mobile, desktop, interaction, theme, and offline browser tests.
- [x] Restore `.spinner` CSS — was `display:none` since v4.25.11 skeleton migration
- [x] Replace opacity-pulse skeleton with shimmer sweep (`@keyframes shimmer`)
- [x] Replace rotating ring spinner with three-dot `dotPulse` animation
- [x] Build shared `.prog-track` / `.prog-fill` CSS system — standardize all 9 inline bars
- [x] Add missing `transition` to 2 poll bar instances (steward_view.html)
- [x] Add light-mode support to all progress bar tracks
- [x] Animate glow-bar with sweep (was static)
