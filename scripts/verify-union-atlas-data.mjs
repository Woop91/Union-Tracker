import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const htmlPath = path.join(root, 'docs', 'union-atlas.html');
const browserDataPath = path.join(root, 'docs', 'union-atlas-data.js');
const manifestPath = path.join(root, 'data', 'manifest.json');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);
const html = fs.readFileSync(htmlPath, 'utf8');
const browserData = fs.readFileSync(browserDataPath, 'utf8');
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
const failures = [];
const check = (condition, message) => { if (!condition) failures.push(message); };

const ids = new Set(data.directory.map(row => row.id));
check(data.metadata.formatVersion === 2, 'expected source-backed formatVersion 2');
check(data.metadata.selfContained === true, 'metadata.selfContained must be true');
check(data.metadata.dataType === 'public-filing union directory', 'unexpected data type');
check(data.directory.length === 20699, 'expected all 20,699 canonical USUnions records');
check(ids.size === data.directory.length, 'union IDs must be unique');
check(data.quality.uniqueIds === ids.size, 'quality.uniqueIds mismatch');
check(data.quality.withMembership >= 19000, 'membership coverage unexpectedly low');
check(data.quality.withZctaCoordinates >= 18500, 'Census coordinate coverage unexpectedly low');
check(data.metadata.sources[0].repository === 'Woop91/USUnions', 'wrong canonical source repository');
check(/^[0-9a-f]{40}$/.test(data.metadata.sources[0].commit), 'source commit must be exact');
check(data.metadata.sources[1].url.includes('census.gov'), 'coordinate source must be Census');

for (const row of data.directory) {
  check(/^olms-\d+$/.test(row.id), `invalid stable ID ${row.id}`);
  check(Number.isInteger(row.olmsFileNumber), `invalid OLMS number for ${row.id}`);
  check(Boolean(row.name), `missing name for ${row.id}`);
  check(Boolean(row.city && row.state), `missing public HQ city/state for ${row.id}`);
  check(!Number.isFinite(row.members) || row.members >= 0, `negative membership for ${row.id}`);
  check(['enriched', 'no_web_presence'].includes(row.enrichmentStatus), `unexpected enrichment state for ${row.id}`);
}

for (const key of ['platformGroups', 'suggestions', 'network', 'admin', 'usUnions']) {
  check(!(key in data), `obsolete generated dataset still present: ${key}`);
}

const sensitivePatterns = [
  /[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i,
  /\b(?:password|access[_-]?token|refresh[_-]?token|api[_-]?key)\b\s*[:=]\s*["'][^"']+/i,
  /-----BEGIN [A-Z ]*PRIVATE KEY-----/
];
for (const pattern of sensitivePatterns) {
  check(!pattern.test(source.toString()), `sensitive pattern found: ${pattern}`);
}

check(html.includes('<script src="union-atlas-data.js"></script>'), 'HTML does not load local data package');
check(html.includes('Nearest union headquarters'), 'HTML missing geographic discovery view');
check(html.includes('Missing values remain missing'), 'HTML missing no-fabrication statement');
check(!/\b(?:src|href)=["']https?:\/\//i.test(html), 'HTML has an external runtime dependency');
check(browserData.includes(`window.UNION_ATLAS_DATA=${source.toString().trim()};`), 'browser data is out of sync');

const digest = crypto.createHash('sha256').update(source).digest('hex');
check(manifest.sha256 === digest, 'manifest SHA-256 does not match union-atlas.json');
check(manifest.counts.unions === data.directory.length, 'manifest union count mismatch');
check(manifest.counts.uniqueIds === ids.size, 'manifest unique ID count mismatch');
check(manifest.syntheticRecords === 0, 'manifest must report zero synthetic records');

if (failures.length) {
  failures.forEach(message => console.error(`FAIL ${message}`));
  process.exit(1);
}
console.log(
  `PASS source-backed Atlas: ${data.directory.length} unions, ${data.quality.withZctaCoordinates} geocoded, SHA-256 ${digest}`
);
