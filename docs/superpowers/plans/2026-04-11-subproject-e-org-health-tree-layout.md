# Sub-Project E — Org Health Tree Layout Fix Implementation Plan

> **For agentic workers:** Inline-executed in the brainstorming session because the fix is 1 line + 3 tests + version bump. No subagents required.

**Goal:** Fix the `branchCount == 1` fallback in the Org Health Tree SVG angle formula so a single-steward org renders its one branch straight up instead of in the upper-left corner.

**Architecture:** 1-line change in the angle-formula ternary at `src/org_health_tree.html:139`. Changes `: 0` to `: angleSpread / 2` so the single-branch fallback picks the fan midpoint instead of the fan edge.

**Tech Stack:** Vanilla JS in a GAS HtmlService template. Jest for test. No DOM needed — the test re-implements the formula inline and asserts expected angles.

---

## Task 1: Write regression test file

**File:** Create `test/org-health-tree-layout.test.js`

- [ ] Write the following file verbatim:

```javascript
/**
 * Org Health Tree layout — sub-project E
 *
 * Verifies the branch-angle formula at src/org_health_tree.html:139
 * produces symmetric layouts for branchCount >= 2 AND a centered
 * single-branch layout for branchCount == 1.
 *
 * The formula is pure math so we re-implement it inline and assert
 * expected angles. A static grep also verifies the source file still
 * uses the corrected fallback (not the buggy `: 0`).
 */

const fs = require('fs');
const path = require('path');

const TREE_SRC = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'org_health_tree.html'),
  'utf8'
);

// Re-implement the formula the file uses at line 139:
//   var angle = startAngle - (branchCount > 1 ? (bi / (branchCount - 1)) * angleSpread : angleSpread / 2);
function computeAngle(bi, branchCount) {
  const angleSpread = Math.PI * 0.75; // 135 degrees fan
  const startAngle = Math.PI / 2 + angleSpread / 2; // upper-left edge
  const progress = branchCount > 1
    ? (bi / (branchCount - 1)) * angleSpread
    : angleSpread / 2; // v4.55.5 sub-project E fix — was `0` (buggy)
  return startAngle - progress;
}

describe('Org Health Tree branch layout (sub-project E)', () => {
  test('branchCount=1: single branch points straight up (π/2)', () => {
    const a = computeAngle(0, 1);
    expect(a).toBeCloseTo(Math.PI / 2, 10);
  });

  test('branchCount=2: branches span [7π/8, π/8] (symmetric around π/2)', () => {
    expect(computeAngle(0, 2)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 2)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('branchCount=3: middle branch is straight up (π/2)', () => {
    expect(computeAngle(0, 3)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 3)).toBeCloseTo(Math.PI / 2, 10);
    expect(computeAngle(2, 3)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('branchCount=4: angles are evenly distributed [7π/8, 5π/8, 3π/8, π/8]', () => {
    expect(computeAngle(0, 4)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 4)).toBeCloseTo(5 * Math.PI / 8, 10);
    expect(computeAngle(2, 4)).toBeCloseTo(3 * Math.PI / 8, 10);
    expect(computeAngle(3, 4)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('source file uses the corrected fallback, not `: 0`', () => {
    // The buggy line was: `... : 0)`
    // The fix changes it to: `... : angleSpread / 2)`
    // The angle computation line must contain `angleSpread / 2` as the ternary false branch.
    const idx = TREE_SRC.indexOf('var angle = startAngle - (branchCount');
    expect(idx).toBeGreaterThan(-1);
    const line = TREE_SRC.substring(idx, TREE_SRC.indexOf('\n', idx));
    expect(line).toMatch(/:\s*angleSpread\s*\/\s*2/);
    expect(line).not.toMatch(/\?\s*\(bi[\s\S]*\)\s*\*\s*angleSpread\s*:\s*0\s*\)/);
  });
});
```

## Task 2: Run failing test

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/org-health-tree-layout.test.js --no-coverage`

Expected: 4/5 PASS (the pure-math tests pass because the test's inline formula is already correct), 1/5 FAIL (the source-file grep test fails because `src/org_health_tree.html` still has `: 0`).

## Task 3: Apply the 1-line fix

**File:** `src/org_health_tree.html:139`

Change this line:
```javascript
          var angle = startAngle - (branchCount > 1 ? (bi / (branchCount - 1)) * angleSpread : 0);
```

To:
```javascript
          // v4.55.5 sub-project E: fallback uses angleSpread/2 (fan midpoint)
          // so a single-branch tree points straight up (π/2) instead of
          // upper-left (7π/8). branchCount>=2 behavior unchanged.
          var angle = startAngle - (branchCount > 1 ? (bi / (branchCount - 1)) * angleSpread : angleSpread / 2);
```

## Task 4: Run tests + guards

Run: `cd "C:/Users/deskc/Documents/SolidBase" && npx jest test/org-health-tree-layout.test.js --no-coverage && npm run test:guards`

Expected: 5/5 org-health-tree tests pass, 282/282 guards pass.

## Task 5: Commit the fix

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/org_health_tree.html test/org-health-tree-layout.test.js dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "fix: single-branch Org Health Tree points straight up (not upper-left)

The branchCount==1 fallback in the angle ternary was `: 0`, which
defaults the single branch angle to 7π/8 (upper-left corner). Fix:
`: angleSpread / 2` so the single branch lands at π/2 (straight up).
branchCount>=2 behavior unchanged. +5 regression tests. Fixes #11.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

## Task 6: Version bump v4.55.5

- Modify `src/01_Core.gs` line 336: `VERSION: "4.55.4",` → `VERSION: "4.55.5",`
- Modify `src/01_Core.gs` line 508: `: '4.55.4'` → `: '4.55.5'`
- Add CHANGELOG.md entry at top:

```markdown
## [4.55.5] - 2026-04-11

### Fixed (Sub-project E — Org Health Tree Layout)
- **#11** Org Health Tree heavily skewed left for small orgs. The `branchCount==1` fallback in the angle formula at `src/org_health_tree.html:139` defaulted to `: 0`, which set `angle = startAngle = 7π/8` (upper-left corner). Changed to `: angleSpread / 2` so a single-branch tree points straight up (`π/2`) instead. For `branchCount >= 2` the formula is unchanged and remains symmetric.

### Tests
- New suite: `test/org-health-tree-layout.test.js` (+5).
- Total deploy guards: 282 (unchanged — the new file is a unit test, not a deploy guard).
```

Commit:

```bash
git -C "C:/Users/deskc/Documents/SolidBase" add src/01_Core.gs CHANGELOG.md dist/
git -C "C:/Users/deskc/Documents/SolidBase" commit -m "chore: bump to v4.55.5 — sub-project E Org Health Tree layout fix

Closes screenshot issue #11 (tree heavily skewed left for small orgs).
Single-commit zone: one-line angle-fallback fix + 5 regression tests.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

## Task 7: Push + clasp DDS prod

- `git fetch origin && git rev-list --left-right --count origin/Main...HEAD` — verify clean ahead
- If behind, `git pull --rebase origin Main`
- `git push origin Main`
- `clasp push --force` (DDS prod Script ID already in .clasp.json)
