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

---

## 🔧 Technical Debt

### [ ] Husky v10 Migration
**Added:** 2026-03-17
**Current version:** 9.1.7 (deprecation warning is forward-looking, no action needed now)

When Husky v10 is released:
- [ ] Remove `#!/bin/sh` shebang from `.husky/pre-commit`, `.husky/pre-push`, `.husky/commit-msg`
- [ ] Run `npx husky init` to regenerate the `_/` directory structure
- [ ] Verify commitlint and lint-staged still trigger correctly

---

## 💡 Feature Ideas

*(empty — add items here as they come up)*

---

## ✅ Recently Completed (for reference)

- [x] Restore `.spinner` CSS — was `display:none` since v4.25.11 skeleton migration
- [x] Replace opacity-pulse skeleton with shimmer sweep (`@keyframes shimmer`)
- [x] Replace rotating ring spinner with three-dot `dotPulse` animation
- [x] Build shared `.prog-track` / `.prog-fill` CSS system — standardize all 9 inline bars
- [x] Add missing `transition` to 2 poll bar instances (steward_view.html)
- [x] Add light-mode support to all progress bar tracks
- [x] Animate glow-bar with sweep (was static)
