const fs = require('fs');
const path = require('path');

const atlasPath = path.join(__dirname, '..', 'docs', 'union-atlas.html');
const atlas = fs.readFileSync(atlasPath, 'utf8');
const dataPath = path.join(__dirname, '..', 'data', 'union-atlas.json');
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const tokensPath = path.join(__dirname, '..', 'design', 'solid-ground', 'tokens.css');
const tokens = fs.readFileSync(tokensPath, 'utf8').trim();
const themeReadmePath = path.join(__dirname, '..', 'design', 'solid-ground', 'README.md');
const themeReadme = fs.readFileSync(themeReadmePath, 'utf8');

describe('Union Atlas USUnions data bridge', () => {
  test('includes the exact seeded aggregate snapshot', () => {
    expect(atlas).toContain('surveyParticipation:62');
    expect(atlas).toContain('weeklyQuestionVotes:147');
    expect(atlas).toContain('{month:"Feb",total:878,newMembers:16,departed:7}');
    expect(atlas).toContain('avgCaseload:23.4');
    expect(atlas).toContain('submissionRate:71');
  });

  test('labels the source as demo rather than production', () => {
    expect(atlas).toContain('mode:"repo-demo"');
    expect(atlas).toContain('production:false');
    expect(atlas).toContain('Demo aggregate');
    expect(atlas).toContain('not production');
  });

  test('states and enforces the privacy boundary', () => {
    expect(atlas).toContain('Member names, roster rows or personal identifiers');
    expect(atlas).toContain('Grievance, discipline or case-level records');
    expect(atlas).toContain('Street addresses, steward names or authentication data');
    expect(atlas).not.toMatch(/[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i);
  });

  test('keeps the imported total separate from Atlas counts', () => {
    expect(atlas).toContain('keeps local operations separate from Atlas-wide organization and proximity counts');
    expect(atlas).toContain('Demo members');
  });

  test('includes the complete standalone data package', () => {
    expect(data.metadata.selfContained).toBe(true);
    expect(data.organizations).toHaveLength(45);
    expect(data.platformGroups).toHaveLength(14);
    expect(data.suggestions).toHaveLength(4);
    expect(data.admin.coverageGaps).toHaveLength(5);
    expect(data.usUnions.membershipTrends).toHaveLength(6);
    expect(data.usUnions.publicIdentitySchema).toHaveLength(7);
  });
});

describe('Union Atlas Solid Ground visual system', () => {
  test('records the exact design source and Claude session', () => {
    expect(themeReadme).toContain('42d54f1619fab56717f4cc2f0b3f4f289aa57b3c');
    expect(themeReadme).toContain('05fd27e3-7aa0-4384-934c-1c26921914cf');
  });

  test('embeds the copied token source in the standalone page', () => {
    const embedded = atlas
      .split('<!-- BEGIN SOLID GROUND TOKENS -->')[1]
      .split('<!-- END SOLID GROUND TOKENS -->')[0]
      .trim();
    expect(embedded).toContain(tokens);
  });

  test('supports explicit light and dark themes', () => {
    expect(atlas).toContain('<html lang="en" data-theme="light">');
    expect(atlas).toContain('html[data-theme="dark"]');
    expect(atlas).toContain('id="theme-toggle"');
    expect(atlas).toContain('localStorage.setItem("union-atlas-theme"');
  });

  test('keeps the visual runtime self-contained', () => {
    expect(atlas).not.toMatch(/<script[^>]+src=["']https?:/i);
    expect(atlas).not.toMatch(/<link[^>]+href=["']https?:/i);
    expect(atlas).not.toMatch(/url\(\s*["']?https?:/i);
  });

  test('uses semantic tokens instead of raw colors in application styles', () => {
    const application = atlas.split('<!-- END SOLID GROUND TOKENS -->')[1];
    expect(application).not.toMatch(/#[0-9a-f]{3,8}\b/i);
    expect(application).not.toMatch(/rgba?\([^)]*\)/i);
  });

  test('keeps compass points keyboard operable without SVG click shims', () => {
    expect(atlas).toContain('el.addEventListener("keydown"');
    expect(atlas).toContain('activate();');
    expect(atlas).not.toContain('el.click();');
  });

  test('defines narrow-phone, phone, and tablet layouts', () => {
    expect(atlas).toContain('@media(max-width:360px)');
    expect(atlas).toContain('@media(max-width:520px)');
    expect(atlas).toContain('@media(max-width:820px)');
    expect(atlas).toContain('env(safe-area-inset-top)');
    expect(atlas).toContain('env(safe-area-inset-bottom)');
  });

  test('contains overflow and touch-target safeguards for mobile', () => {
    expect(atlas).toContain('html,body{max-width:100%;overflow-x:hidden}');
    expect(atlas).toContain('min-height:44px');
    expect(atlas).toContain('overscroll-behavior-inline:contain');
    expect(atlas).toContain('#whitespace table{min-width:560px}');
  });

  test('uses compact admin tables and a mobile bottom-sheet dialog', () => {
    expect(atlas).toContain('sector:matchMedia("(max-width:520px)").matches');
    expect(atlas).toContain('region:matchMedia("(max-width:520px)").matches');
    expect(atlas).toContain('max-height:94dvh');
    expect(atlas).toContain('class="modal-actions"');
  });
});
