import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import {
  readGzipJsonLines,
  readJson,
  sha256
} from './atlas-source-utils.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const htmlPath = path.join(root, 'docs', 'union-atlas.html');
const standalonePath = path.join(root, 'docs', 'union-atlas-standalone.html');
const workerPath = path.join(root, 'docs', 'union-atlas-worker.js');
const indexPath = path.join(root, 'docs', 'union-atlas-index.js');
const manifestPath = path.join(root, 'data', 'manifest.json');
const sourcePackageRoot = path.join(root, 'data', 'source', 'usunions');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);
const html = fs.readFileSync(htmlPath, 'utf8');
const standalone = fs.readFileSync(standalonePath, 'utf8');
const worker = fs.readFileSync(workerPath, 'utf8');
const indexSource = fs.readFileSync(indexPath, 'utf8');
const manifest = readJson(manifestPath);
const sourcePackage = readJson(path.join(sourcePackageRoot, 'manifest.json'));
const failures = [];
const check = (condition, message) => { if (!condition) failures.push(message); };

const ids = new Set(data.directory.map(row => row.id));
check(data.metadata.formatVersion === 3, 'expected complete-data formatVersion 3');
check(data.metadata.repositoryLocal === true, 'metadata.repositoryLocal must be true');
check(data.metadata.dataType.includes('organizing evidence'), 'data type must include organizing evidence');
check(data.directory.length === 20699, 'expected all 20,699 canonical USUnions records');
check(ids.size === data.directory.length, 'union IDs must be unique');
check(data.quality.uniqueIds === ids.size, 'quality.uniqueIds mismatch');
check(data.quality.withMembership === 19554, 'membership coverage changed unexpectedly');
check(data.quality.withWebsites === 5276, 'website coverage changed unexpectedly');
check(data.quality.withAddressCoordinates >= 15000, 'address geocode coverage unexpectedly low');
check(data.quality.withCoordinates >= 19000, 'combined coordinate coverage unexpectedly low');
check(data.quality.withoutCoordinates === data.directory.length - data.quality.withCoordinates, 'coordinate gap count mismatch');
check(data.datasets.organizing.records === 23497, 'expected all NLRB representation cases');
check(data.datasets.organizing.rcAttempts === 18372, 'expected all RC organizing attempts');
check(data.datasets.organizing.groups === 5714, 'expected all organizing group rollups');
check(data.datasets.research.articles === 64, 'expected complete research articles');
check(data.datasets.research.statistics === 97, 'expected complete research statistics');
check(data.datasets.research.sources === 10, 'expected complete research sources');
check(data.metadata.sources[0].repository === 'Woop91/USUnions', 'wrong source repository');
check(/^[0-9a-f]{40}$/.test(data.metadata.sources[0].commit), 'source commit must be exact');
check(data.metadata.sources[1].url.includes('census.gov'), 'address coordinate source must be Census');
check(data.metadata.sources[2].url.includes('census.gov'), 'fallback coordinate source must be Census');

for (const row of data.directory) {
  check(/^olms-\d+$/.test(row.id), `invalid stable ID ${row.id}`);
  check(Number.isInteger(row.olmsFileNumber), `invalid OLMS number for ${row.id}`);
  check(Boolean(row.name), `missing name for ${row.id}`);
  check(Boolean(row.city && row.state), `missing public HQ city/state for ${row.id}`);
  check(
    !Number.isFinite(row.membership.latest?.count) || row.membership.latest.count >= 0,
    `negative membership for ${row.id}`
  );
  check(
    [null, 'address', 'zcta'].includes(row.geocodePrecision),
    `unexpected geocode precision for ${row.id}`
  );
  check(!('street' in row), `street address persisted for ${row.id}`);
  check(!('officers' in row), `officer records persisted for ${row.id}`);
  check(!('staff' in row), `staff records persisted for ${row.id}`);
  check(!('stewards' in row), `steward records persisted for ${row.id}`);
}

for (const key of ['platformGroups', 'suggestions', 'network', 'admin', 'usUnions']) {
  check(!(key in data), `obsolete seeded dataset still present: ${key}`);
}

const sensitivePatterns = [
  /[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i,
  /\b(?:password|access[_-]?token|refresh[_-]?token|api[_-]?key)\b\s*[:=]\s*["'][^"']+/i,
  /-----BEGIN [A-Z ]*PRIVATE KEY-----/
];
for (const pattern of sensitivePatterns) {
  check(!pattern.test(source.toString()), `sensitive pattern found in directory: ${pattern}`);
}

const sourceFile = relative => path.join(sourcePackageRoot, ...relative.split('/'));
const attempts = readGzipJsonLines(sourceFile('data/organizing/attempts.jsonl.gz'));
const groups = readGzipJsonLines(sourceFile('data/organizing/groups.jsonl.gz'));
const articles = readGzipJsonLines(sourceFile('data/research/articles.jsonl.gz'));
const researchSources = readGzipJsonLines(sourceFile('data/research/sources.jsonl.gz'));
const statistics = readGzipJsonLines(sourceFile('data/research/statistics.jsonl.gz'));
check(attempts.length === manifest.counts.organizingCases, 'organizing attempts package mismatch');
check(groups.length === manifest.counts.organizingGroups, 'organizing groups package mismatch');
check(articles.length === manifest.counts.researchArticles, 'research article package mismatch');
check(researchSources.length === manifest.counts.researchSources, 'research source package mismatch');
check(statistics.length === manifest.counts.researchStatistics, 'research statistic package mismatch');
check(attempts.every(row => row.privacy?.participant_contacts_removed === true), 'organizing contacts not removed');

for (const entry of sourcePackage.files) {
  const local = fs.readFileSync(sourceFile(entry.localPath));
  check(sha256(local) === entry.localSha256, `local source package SHA mismatch: ${entry.localPath}`);
  const expanded = entry.compression === 'gzip' ? zlib.gunzipSync(local) : local;
  check(sha256(expanded) === entry.sourceSha256, `source package content mismatch: ${entry.path}`);
}

check(html.includes('<script src="union-atlas-index.js"></script>'), 'HTML missing compact index');
check(html.includes('new Worker("union-atlas-worker.js")'), 'HTML missing background worker');
check(html.includes('Organizing attempts'), 'HTML missing organizing view');
check(html.includes('Labor research'), 'HTML missing research view');
check(html.includes('The paginated list contains every result.'), 'HTML missing full-results statement');
check(!html.includes('union-atlas-data.js'), 'legacy eager JavaScript package still referenced');
check(!/<script[^>]+src=["']https?:/i.test(html), 'HTML has external script dependency');
check(!/<link[^>]+href=["']https?:/i.test(html), 'HTML has external stylesheet dependency');
check(worker.includes('DecompressionStream'), 'worker does not decompress packages');
check(worker.includes('directory-search'), 'worker missing directory search');
check(worker.includes('organizing-search'), 'worker missing organizing search');
check(worker.includes('record-detail'), 'worker missing lazy record details');
check(indexSource.includes('window.UNION_ATLAS_INDEX='), 'browser index missing');
check(!standalone.match(/<script\s+src=/i), 'standalone HTML still has companion scripts');
check(standalone.includes('window.UNION_ATLAS_EMBEDDED='), 'standalone packages not embedded');

for (const packageEntry of Object.values(manifest.browserPackages)) {
  const file = fs.readFileSync(path.join(root, 'docs', packageEntry.path));
  check(file.length === packageEntry.bytes, `browser package size mismatch: ${packageEntry.path}`);
  check(sha256(file) === packageEntry.sha256, `browser package SHA mismatch: ${packageEntry.path}`);
}

const digest = sha256(source);
check(manifest.sha256 === digest, 'manifest SHA-256 does not match union-atlas.json');
check(manifest.counts.unions === data.directory.length, 'manifest union count mismatch');
check(manifest.counts.uniqueIds === ids.size, 'manifest unique ID count mismatch');
check(manifest.syntheticRecords === 0, 'manifest must report zero synthetic records');
check(manifest.source.commit === sourcePackage.commit, 'source package commit mismatch');

if (failures.length) {
  failures.forEach(message => console.error(`FAIL ${message}`));
  process.exit(1);
}
console.log(
  `PASS complete Atlas: ${data.directory.length} unions, ${attempts.length} organizing cases, `
  + `${articles.length} articles, ${data.quality.withAddressCoordinates} address geocodes, zero synthetic records.`
);
