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
  test('branchCount=1: single branch points straight up (pi/2)', () => {
    const a = computeAngle(0, 1);
    expect(a).toBeCloseTo(Math.PI / 2, 10);
  });

  test('branchCount=2: branches span [7pi/8, pi/8] (symmetric around pi/2)', () => {
    expect(computeAngle(0, 2)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 2)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('branchCount=3: middle branch is straight up (pi/2)', () => {
    expect(computeAngle(0, 3)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 3)).toBeCloseTo(Math.PI / 2, 10);
    expect(computeAngle(2, 3)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('branchCount=4: angles are evenly distributed [7pi/8, 5pi/8, 3pi/8, pi/8]', () => {
    expect(computeAngle(0, 4)).toBeCloseTo(7 * Math.PI / 8, 10);
    expect(computeAngle(1, 4)).toBeCloseTo(5 * Math.PI / 8, 10);
    expect(computeAngle(2, 4)).toBeCloseTo(3 * Math.PI / 8, 10);
    expect(computeAngle(3, 4)).toBeCloseTo(Math.PI / 8, 10);
  });

  test('source file uses the corrected fallback, not `: 0`', () => {
    const idx = TREE_SRC.indexOf('var angle = startAngle - (branchCount');
    expect(idx).toBeGreaterThan(-1);
    const line = TREE_SRC.substring(idx, TREE_SRC.indexOf('\n', idx));
    expect(line).toMatch(/:\s*angleSpread\s*\/\s*2/);
    expect(line).not.toMatch(/\?\s*\(bi[\s\S]*\)\s*\*\s*angleSpread\s*:\s*0\s*\)/);
  });
});
