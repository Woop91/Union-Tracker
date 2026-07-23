const fs = require('fs');
const path = require('path');

const atlasPath = path.join(__dirname, '..', 'docs', 'union-atlas.html');
const atlas = fs.readFileSync(atlasPath, 'utf8');
const dataPath = path.join(__dirname, '..', 'data', 'union-atlas.json');
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const manifest = JSON.parse(
  fs.readFileSync(path.join(__dirname, '..', 'data', 'manifest.json'), 'utf8')
);
const browserData = fs.readFileSync(
  path.join(__dirname, '..', 'docs', 'union-atlas-data.js'),
  'utf8'
);
const tokens = fs.readFileSync(
  path.join(__dirname, '..', 'design', 'solid-ground', 'tokens.css'),
  'utf8'
).trim();
const themeReadme = fs.readFileSync(
  path.join(__dirname, '..', 'design', 'solid-ground', 'README.md'),
  'utf8'
);

describe('Union Atlas source-backed data package', () => {
  test('contains the complete canonical USUnions directory', () => {
    expect(data.metadata.selfContained).toBe(true);
    expect(data.metadata.dataType).toBe('public-filing union directory');
    expect(data.directory).toHaveLength(20699);
    expect(new Set(data.directory.map(row => row.id)).size).toBe(20699);
    expect(data.quality.uniqueIds).toBe(20699);
  });

  test('pins traceable source identities', () => {
    expect(data.metadata.sources[0].repository).toBe('Woop91/USUnions');
    expect(data.metadata.sources[0].commit).toMatch(/^[0-9a-f]{40}$/);
    expect(data.metadata.sources[0].manifestSha256).toMatch(/^[0-9a-f]{64}$/);
    expect(data.metadata.sources[1].url).toContain('census.gov');
    expect(manifest.source.commit).toBe(data.metadata.sources[0].commit);
  });

  test('contains no obsolete generated Atlas collections', () => {
    for (const key of ['platformGroups', 'suggestions', 'network', 'admin', 'usUnions']) {
      expect(data).not.toHaveProperty(key);
    }
    expect(manifest.syntheticRecords).toBe(0);
  });

  test('preserves missingness and public-data privacy boundary', () => {
    expect(data.directory.some(row => row.members === null)).toBe(true);
    expect(data.metadata.privacy).toContain('Organization-level public filing data only');
    expect(JSON.stringify(data)).not.toMatch(
      /[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i
    );
  });

  test('has strong location and membership coverage without filling gaps', () => {
    expect(data.quality.withMembership).toBeGreaterThanOrEqual(19000);
    expect(data.quality.withZctaCoordinates).toBeGreaterThanOrEqual(18500);
    expect(data.directory.every(row => row.city && row.state)).toBe(true);
    expect(data.directory.filter(row => row.zip === null)).toHaveLength(9);
  });

  test('keeps browser data synchronized with the repository package', () => {
    const serialized = fs.readFileSync(dataPath, 'utf8').trim();
    expect(browserData).toContain(`window.UNION_ATLAS_DATA=${serialized};`);
    expect(atlas).toContain('<script src="union-atlas-data.js"></script>');
  });
});

describe('Union Atlas Solid Ground visual system', () => {
  test('records the exact design source and Claude session', () => {
    expect(themeReadme).toContain('42d54f1619fab56717f4cc2f0b3f4f289aa57b3c');
    expect(themeReadme).toContain('05fd27e3-7aa0-4384-934c-1c26921914cf');
  });

  test('embeds the copied token source', () => {
    const embedded = atlas
      .split('/* BEGIN SOLID GROUND TOKENS */')[1]
      .split('/* END SOLID GROUND TOKENS */')[0]
      .trim();
    expect(embedded).toContain(tokens);
  });

  test('supports explicit light and dark themes', () => {
    expect(atlas).toContain('html[data-theme="dark"]');
    expect(atlas).toContain('id="theme"');
    expect(atlas).toContain('localStorage.setItem("union-atlas-theme"');
  });

  test('keeps runtime assets local', () => {
    expect(atlas).not.toMatch(/<script[^>]+src=["']https?:/i);
    expect(atlas).not.toMatch(/<link[^>]+href=["']https?:/i);
    expect(atlas).not.toMatch(/url\(\s*["']?https?:/i);
  });

  test('uses semantic tokens instead of raw application colors', () => {
    const application = atlas
      .split('/* END SOLID GROUND TOKENS */')[1]
      .split('</style>')[0];
    expect(application).not.toMatch(/#[0-9a-f]{3,8}\b/i);
    expect(application).not.toMatch(/rgba?\([^)]*\)/i);
  });

  test('keeps compass points keyboard discoverable', () => {
    expect(atlas).toContain('class="point" tabindex="0" role="img"');
    expect(atlas).toContain('<title id="compass-title">');
  });

  test('defines narrow-phone, phone, and tablet layouts', () => {
    expect(atlas).toContain('@media(max-width:360px)');
    expect(atlas).toContain('@media(max-width:520px)');
    expect(atlas).toContain('@media(max-width:820px)');
    expect(atlas).toContain('env(safe-area-inset-top)');
    expect(atlas).toContain('env(safe-area-inset-bottom)');
  });

  test('contains mobile overflow and touch safeguards', () => {
    expect(atlas).toContain('html,body{margin:0;max-width:100%;overflow-x:hidden}');
    expect(atlas).toContain('min-height:44px');
    expect(atlas).toContain('overscroll-behavior-inline:contain');
    expect(atlas).toContain('.tablewrap{overflow:auto');
  });
});
