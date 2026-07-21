/**
 * Cross-platform Jest coverage launcher for Google Apps Script .gs sources.
 * load-source.js needs an explicit signal because .gs files execute through
 * eval and therefore bypass Jest's normal transform/instrumentation pipeline.
 */

const { spawnSync } = require('child_process');

const jestBin = require.resolve('jest/bin/jest');
const args = [jestBin, '--coverage', '--verbose', ...process.argv.slice(2)];
const result = spawnSync(process.execPath, args, {
  cwd: process.cwd(),
  env: { ...process.env, SOLIDBASE_COVERAGE: '1' },
  stdio: 'inherit',
});

if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}

process.exit(result.status == null ? 1 : result.status);
