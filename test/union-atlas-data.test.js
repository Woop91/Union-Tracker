const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const root = path.join(__dirname, '..');
const atlas = fs.readFileSync(path.join(root, 'docs', 'union-atlas.html'), 'utf8');
const standalone = fs.readFileSync(
  path.join(root, 'docs', 'union-atlas-standalone.html'),
  'utf8'
);
const worker = fs.readFileSync(path.join(root, 'docs', 'union-atlas-worker.js'), 'utf8');
const dataPath = path.join(root, 'data', 'union-atlas.json');
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'data', 'manifest.json'), 'utf8'));
const sourceManifest = JSON.parse(
  fs.readFileSync(path.join(root, 'data', 'source', 'usunions', 'manifest.json'), 'utf8')
);
const directoryPackage = JSON.parse(zlib.gunzipSync(
  fs.readFileSync(path.join(root, 'docs', 'union-atlas-data.json.gz'))
));
const evidencePackage = JSON.parse(zlib.gunzipSync(
  fs.readFileSync(path.join(root, 'docs', 'union-atlas-evidence.json.gz'))
));
const tokens = fs.readFileSync(
  path.join(root, 'design', 'solid-ground', 'tokens.css'),
  'utf8'
).trim();
const themeReadme = fs.readFileSync(
  path.join(root, 'design', 'solid-ground', 'README.md'),
  'utf8'
);

describe('Union Atlas complete real-data package', () => {
  test('contains the full canonical directory and exact upstream identity', () => {
    expect(data.metadata.repositoryLocal).toBe(true);
    expect(data.metadata.formatVersion).toBe(3);
    expect(data.directory).toHaveLength(20699);
    expect(new Set(data.directory.map(row => row.id)).size).toBe(20699);
    expect(data.metadata.sources[0].repository).toBe('Woop91/USUnions');
    expect(data.metadata.sources[0].commit).toBe(
      '55a20f1433d7a38f7af3077de7464760f1bc188f'
    );
    expect(sourceManifest.commit).toBe(data.metadata.sources[0].commit);
  });

  test('imports every organizing and research record', () => {
    expect(data.datasets.organizing.records).toBe(23497);
    expect(data.datasets.organizing.rcAttempts).toBe(18372);
    expect(data.datasets.organizing.groups).toBe(5714);
    expect(data.datasets.research).toEqual({
      articles: 64,
      statistics: 97,
      sources: 10
    });
    expect(evidencePackage.organizing.attempts).toHaveLength(23497);
    expect(evidencePackage.organizing.groups).toHaveLength(5714);
    expect(evidencePackage.research.articles).toHaveLength(64);
    expect(evidencePackage.research.statistics).toHaveLength(97);
    expect(evidencePackage.research.sources).toHaveLength(10);
  });

  test('uses address geocodes first and honest ZIP fallbacks', () => {
    expect(data.quality.withAddressCoordinates).toBe(15243);
    expect(data.quality.withZctaCoordinates).toBe(3809);
    expect(data.quality.withCoordinates).toBe(19052);
    expect(data.quality.withoutCoordinates).toBe(1647);
    expect(data.directory.filter(row => row.zip === null)).toHaveLength(9);
    expect(data.directory.every(row => [null, 'address', 'zcta'].includes(row.geocodePrecision))).toBe(true);
  });

  test('retains public organization fields while excluding private data', () => {
    const sample = data.directory.find(row => row.membership.history.length > 1);
    expect(sample.membership.history.length).toBeGreaterThan(1);
    expect(sample).toHaveProperty('finances');
    expect(sample).toHaveProperty('sector');
    expect(sample).toHaveProperty('sources');
    expect(data.directory.some(row => row.membership.latest === null)).toBe(true);
    expect(data.directory.every(row => !Object.hasOwn(row, 'street'))).toBe(true);
    expect(JSON.stringify(data)).not.toMatch(
      /[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i
    );
    expect(evidencePackage.organizing.attempts.every(
      row => row.privacy.participant_contacts_removed
    )).toBe(true);
  });

  test('contains no seeded Atlas collections or synthetic records', () => {
    for (const key of ['platformGroups', 'suggestions', 'network', 'admin', 'usUnions']) {
      expect(data).not.toHaveProperty(key);
    }
    expect(manifest.syntheticRecords).toBe(0);
  });

  test('uses compact lazy browser packages', () => {
    expect(directoryPackage.directory).toHaveLength(20699);
    expect(manifest.browserPackages.directory.bytes).toBeLessThan(2 * 1024 * 1024);
    expect(manifest.browserPackages.details.bytes).toBeLessThan(4 * 1024 * 1024);
    expect(atlas).toContain('new Worker("union-atlas-worker.js")');
    expect(worker).toContain('directory-search');
    expect(worker).toContain('record-detail');
    expect(worker).toContain('organizing-search');
    expect(worker).toContain('research-search');
    expect(atlas).not.toContain('union-atlas-data.js');
  });

  test('provides a true one-file offline export', () => {
    expect(standalone).not.toMatch(/<script\s+src=/i);
    expect(standalone).toContain('window.UNION_ATLAS_INDEX=');
    expect(standalone).toContain('window.UNION_ATLAS_EMBEDDED=');
    expect(standalone).toContain('"details":');
    expect(standalone).toContain('"evidence":');
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

  test('defines complete semantic tokens in light and dark themes', () => {
    for (const token of [
      '--background', '--foreground', '--card', '--card-foreground',
      '--popover', '--popover-foreground', '--primary', '--primary-foreground',
      '--secondary', '--secondary-foreground', '--muted', '--muted-foreground',
      '--accent', '--accent-foreground', '--destructive', '--destructive-foreground',
      '--border', '--input', '--ring'
    ]) {
      expect(atlas).toContain(token);
    }
    expect(atlas).toContain('html[data-theme="dark"]');
    expect(atlas).toContain('localStorage.setItem("union-atlas-theme"');
  });

  test('keeps runtime assets local and application colors semantic', () => {
    expect(atlas).not.toMatch(/<script[^>]+src=["']https?:/i);
    expect(atlas).not.toMatch(/<link[^>]+href=["']https?:/i);
    expect(atlas).not.toMatch(/url\(\s*["']?https?:/i);
    const application = atlas
      .split('/* END SOLID GROUND TOKENS */')[1]
      .split('</style>')[0];
    expect(application).not.toMatch(/#[0-9a-f]{3,8}\b/i);
    expect(application).not.toMatch(/rgba?\([^)]*\)/i);
  });

  test('keeps compass points and complete results accessible', () => {
    expect(atlas).toContain('class="point" tabindex="0" role="img"');
    expect(atlas).toContain('<title id="compass-title">');
    expect(atlas).toContain('The paginated list contains every result.');
    expect(atlas).toContain('aria-labelledby="detail-title"');
  });

  test('defines required mobile layouts, touch targets, and reduced motion', () => {
    expect(atlas).toContain('@media(max-width:360px)');
    expect(atlas).toContain('@media(max-width:520px)');
    expect(atlas).toContain('@media(max-width:768px)');
    expect(atlas).toContain('env(safe-area-inset-top)');
    expect(atlas).toContain('env(safe-area-inset-bottom)');
    expect(atlas).toContain('min-height:44px');
    expect(atlas).toContain('overscroll-behavior-inline:contain');
    expect(atlas).toContain('@media(prefers-reduced-motion:reduce)');
  });

  test('contains loading, empty, error, disabled, selected, and focus states', () => {
    expect(atlas).toContain('.loading');
    expect(atlas).toContain('.empty');
    expect(atlas).toContain('.error');
    expect(atlas).toContain('button:disabled');
    expect(atlas).toContain('[aria-selected="true"]');
    expect(atlas).toContain(':focus-visible');
  });
});
