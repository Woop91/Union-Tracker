/**
 * Survey period state coherence — sub-project B Zone 4
 *
 * Tests the contract that both member_view and steward_view derive
 * "is there an active survey period?" from one backend-authoritative
 * answer, not from client-side month logic.
 */

const fs = require('fs');
const path = require('path');

const memberSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'member_view.html'),
  'utf8'
);
const stewardSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', 'steward_view.html'),
  'utf8'
);
const dataServiceSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'),
  'utf8'
);

describe('Survey period state coherence (sub-project B Zone 4)', () => {
  test('member_view: hardcoded quarterly "Jan, Apr, Jul, Oct" message is gone', () => {
    expect(memberSrc).not.toMatch(/Jan, Apr, Jul, Oct/);
  });

  test('member_view: still guards on period.status Active (backend source of truth)', () => {
    expect(memberSrc).toMatch(/period\.status\s*!==\s*['"]Active['"]|period\.status\s*===\s*['"]Active['"]/);
  });

  test('member_view: new "no survey open" message explains steward control, not quarterly schedule', () => {
    expect(memberSrc).toMatch(/your steward|steward will notify|opens when.*steward|admin activat/i);
  });

  test('steward_view: distinguishes "no period active" from "no members in scope"', () => {
    expect(stewardSrc).toMatch(/No survey period is currently active/);
    expect(stewardSrc).toMatch(/no members.*scope|no members.*assigned|expand.*scope/i);
  });

  test('getStewardSurveyTracking backend returns periodActive + periodName', () => {
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?periodActive/);
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?periodName/);
  });

  test('getStewardSurveyTracking periodActive is derived from getSurveyPeriod() not client-side logic', () => {
    expect(dataServiceSrc).toMatch(/getStewardSurveyTracking[\s\S]*?getSurveyPeriod/);
    const fn = dataServiceSrc.match(/function getStewardSurveyTracking[\s\S]*?(?=\n  function |\nfunction )/);
    if (fn) {
      expect(fn[0]).not.toMatch(/getMonth\(\)\s*%/);
    }
  });

  test('backend return object has both periodActive and periodName keys', () => {
    expect(dataServiceSrc).toMatch(/periodActive:\s*!!\w+|periodActive:\s*\(.*\.status\s*===\s*['"]Active['"]\)|periodActive:\s*Boolean/);
    expect(dataServiceSrc).toMatch(/periodName:\s*\w+\.name\s*\|\|\s*null|periodName:\s*\(.*\.name|periodName:\s*periodName/);
  });
});
