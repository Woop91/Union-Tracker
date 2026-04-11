# Sub-Project E — Org Health Tree Layout Fix

**Date:** 2026-04-11
**Target release:** DDS v4.55.5 (patch wave)
**Status:** Compact spec (1-line fix + 1 regression test)
**Closes issue:** #11 (Org Health Tree heavily skewed left, tiny trunk, enormous empty right side)

---

## 1. Motivation

The Org Health Tree SVG at `src/org_health_tree.html` renders stewards as branches radiating from a central trunk. Visible bug #11: with certain org configurations the tree clusters all branches on the left side of the canvas, leaving the right side empty — the user's screenshot shows this with "Org Score: 15" (a small org). The reference screenshot from SolidBase (#12) shows the expected balanced tree.

## 2. Root cause (verified during brainstorming)

`src/org_health_tree.html:117`:
```javascript
var startAngle = Math.PI / 2 + angleSpread / 2; // start from upper-left
```

`src/org_health_tree.html:139` branch angle calculation:
```javascript
var angle = startAngle - (branchCount > 1 ? (bi / (branchCount - 1)) * angleSpread : 0);
```

For `branchCount > 1`, the formula produces angles symmetric around `π/2` (straight up), spanning [π/8, 7π/8]. This is correct. For `branchCount == 1`, the ternary's false branch evaluates to `0`, so `angle = startAngle = 7π/8 ≈ 157.5°`. The single branch aims upper-left; the trunk remains centered; the right half of the canvas is empty. Any org with exactly one steward (or one steward with all the members) hits this bug.

## 3. Fix

Change line 139's ternary false branch from `: 0` to `: angleSpread / 2`:

```javascript
var angle = startAngle - (branchCount > 1 ? (bi / (branchCount - 1)) * angleSpread : angleSpread / 2);
```

For `branchCount == 1`: `angle = 7π/8 - 3π/8 = π/2` (90°, straight up). The single branch rises from the trunk and leaves are distributed by the existing `leafAngle` jitter formula (line 179), which already adds ±30° per leaf. The result is a centered trunk with a single vertical branch bushing out with leaves — balanced and correct.

For `branchCount > 1`: behavior is unchanged.

## 4. Test plan

New test file `test/org-health-tree-layout.test.js` (~3 tests):

1. When `branchCount == 1`, the single branch angle equals `π/2` (not `7π/8`).
2. When `branchCount == 2`, angles are [7π/8, π/8] (symmetric left/right — regression guard).
3. When `branchCount == 4`, angles are [7π/8, 5π/8, 3π/8, π/8] (symmetric fan — regression guard).

The test extracts the angle-computation formula from `src/org_health_tree.html` by reading the file as a string and locating the `startAngle - (branchCount > 1 ? ...)` expression via regex. Alternatively, because the formula is pure math, the test can re-implement the formula inline and assert expected values — that's more brittle but simpler.

Decision: use inline re-implementation for simplicity (3 tests, no bracket-depth parsing needed).

## 5. Deploy

Single commit on DDS Main, version bump to v4.55.5, push to origin/Main, clasp push to DDS prod. SolidBase sync deferred to the final catchup pass per Option Y.

## 6. Success criteria

- Any org with 1 steward shows the single branch rising straight up from the trunk, leaves distributed around it.
- Any org with 2+ stewards is unchanged.
- +3 tests, 282 → 285 deploy guards.
- DDS v4.55.5 live on prod.
