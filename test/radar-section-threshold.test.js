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
