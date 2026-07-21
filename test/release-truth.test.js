/**
 * Release-truth guards: versions stay aligned and production tooling cannot
 * silently claim success without an exact deployed Git SHA.
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const read = (file) => fs.readFileSync(path.join(ROOT, file), 'utf8');

describe('release version alignment', () => {
  const pkg = JSON.parse(read('package.json'));
  const lock = JSON.parse(read('package-lock.json'));
  const version = pkg.version;

  test('package and lockfile versions match', () => {
    expect(lock.version).toBe(version);
    expect(lock.packages[''].version).toBe(version);
  });

  test.each([
    ['src/01_Core.gs', new RegExp(`VERSION:\\s*["']${version.replace(/\./g, '\\.')}["']`)],
    ['README.md', new RegExp(`Current version:\\s*${version.replace(/\./g, '\\.')}`)],
    ['FEATURES.md', new RegExp(`Version:\\*\\*\\s*${version.replace(/\./g, '\\.')}`)],
    ['CLAUDE.md', new RegExp(`Version:\\*\\*\\s*${version.replace(/\./g, '\\.')}`)],
    ['CHANGELOG.md', new RegExp(`## \\[${version.replace(/\./g, '\\.')}\\]`)],
  ])('%s reports package version', (file, pattern) => {
    expect(read(file)).toMatch(pattern);
  });
});

describe('exact-SHA deployment contract', () => {
  test('source contains one release SHA placeholder', () => {
    const core = read('src/01_Core.gs');
    expect(core.match(/__SOLIDBASE_GIT_SHA__/g)).toHaveLength(1);
  });

  test('build validates and replaces an exact source SHA', () => {
    const build = read('build.js');
    expect(build).toContain("args.indexOf('--source-sha')");
    expect(build).toContain('/^[0-9a-f]{40}$/i');
    expect(build).toContain('content.replace(BUILD_SHA_TOKEN, sourceSha)');
  });

  test('deploy creates an immutable version, updates deployment, and verifies health', () => {
    const deploy = read('scripts/deploy.sh');
    expect(deploy).toContain('clasp version');
    expect(deploy).toContain('clasp deploy --deploymentId');
    expect(deploy).toContain('data.sourceSha !== expectedSha');
    expect(deploy).toContain('data.version !== expectedVersion');
  });

  test('post-deploy smoke fails closed and compares expected SHA', () => {
    const workflow = read('.github/workflows/post-deploy-smoke.yml');
    expect(workflow).toContain('WEBAPP_URL secret not configured');
    expect(workflow).toContain('exit 1');
    expect(workflow).toContain('Deployed SHA mismatch');
    expect(workflow).not.toContain('skipped=true');
  });
});
