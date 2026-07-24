import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { readJson, sha256 } from './atlas-source-utils.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const args = process.argv.slice(2);
const valueAfter = flag => {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : null;
};
const sourceRoot = path.resolve(
  valueAfter('--source') || process.env.USUNIONS_SOURCE || ''
);
if (!sourceRoot || !fs.existsSync(path.join(sourceRoot, '.git'))) {
  throw new Error('Pass --source <USUnions checkout> or set USUNIONS_SOURCE.');
}

const manifest = readJson(path.join(root, 'data', 'manifest.json'));
const sourcePackage = readJson(
  path.join(root, 'data', 'source', 'usunions', 'manifest.json')
);
const currentCommit = execFileSync(
  'git',
  ['-C', sourceRoot, 'rev-parse', 'HEAD'],
  { encoding: 'utf8' }
).trim();
const pinnedCommit = manifest.source.commit;
const relevantPaths = ['data/unions', 'data/organizing', 'data/research'];
let changed = [];
try {
  changed = execFileSync(
    'git',
    ['-C', sourceRoot, 'diff', '--name-only', `${pinnedCommit}..${currentCommit}`, '--', ...relevantPaths],
    { encoding: 'utf8' }
  ).split(/\r?\n/).filter(Boolean);
} catch {
  changed = ['Source history cannot compare to pinned commit.'];
}

const packageFailures = [];
for (const file of sourcePackage.files) {
  const sourcePath = path.join(sourceRoot, ...file.path.split('/'));
  if (!fs.existsSync(sourcePath)) {
    packageFailures.push(`${file.path}: missing upstream`);
    continue;
  }
  const digest = sha256(fs.readFileSync(sourcePath));
  if (digest !== file.sourceSha256) {
    packageFailures.push(`${file.path}: source SHA changed`);
  }
}

if (changed.length || packageFailures.length) {
  console.error(`STALE Atlas source ${pinnedCommit.slice(0, 12)} -> ${currentCommit.slice(0, 12)}`);
  changed.slice(0, 20).forEach(file => console.error(`CHANGED ${file}`));
  packageFailures.forEach(failure => console.error(`PACKAGE ${failure}`));
  process.exit(1);
}

console.log(
  `FRESH Atlas source ${pinnedCommit}: directory, organizing, research, and local package hashes match.`
);
