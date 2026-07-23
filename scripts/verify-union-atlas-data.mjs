import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const htmlPath = path.join(root, 'docs', 'union-atlas.html');
const manifestPath = path.join(root, 'data', 'manifest.json');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);
const html = fs.readFileSync(htmlPath, 'utf8');
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
const failures = [];
const check = (condition, message) => { if (!condition) failures.push(message); };

const organizationIds = new Set(data.organizations.map(row => row.id));
const groupIds = new Set(data.platformGroups.map(row => row.id));
check(data.metadata.selfContained === true, 'metadata.selfContained must be true');
check(data.organizations.length === 45, 'expected all 45 Atlas organizations');
check(data.platformGroups.length === 14, 'expected all 14 platform groups');
check(organizationIds.size === data.organizations.length, 'organization IDs must be unique');
check(groupIds.size === data.platformGroups.length, 'platform group IDs must be unique');

for (const row of data.organizations) {
  check(!row.parent || organizationIds.has(row.parent), `missing parent ${row.parent} for ${row.id}`);
}
for (const row of data.platformGroups) {
  check(organizationIds.has(row.node), `missing organization ${row.node} for group ${row.id}`);
  for (const sector of row.sectors) check(Boolean(data.taxonomies.sectors[sector]), `unknown sector ${sector}`);
  for (const issue of row.issues) check(Boolean(data.taxonomies.issues[issue]), `unknown issue ${issue}`);
}
for (const row of data.suggestions) check(groupIds.has(row.id), `unknown suggested group ${row.id}`);
for (const id of data.network.alliance.groupIds) check(groupIds.has(id), `unknown alliance group ${id}`);

check(data.usUnions.provenance.production === false, 'USUnions snapshot must remain labeled non-production');
check(data.usUnions.membershipTrends.length === 6, 'expected all six USUnions membership rows');
check(data.usUnions.membershipTrends.at(-1).total === 878, 'expected February demo total of 878');
check(data.usUnions.publicIdentitySchema.length === 7, 'expected all seven safe identity fields');

const sensitivePatterns = [
  /[A-Z0-9._%+-]+@(?!example\.com)[A-Z0-9.-]+\.[A-Z]{2,}/i,
  /\b(?:password|access[_-]?token|refresh[_-]?token|api[_-]?key)\b\s*[:=]\s*["'][^"']+/i
];
for (const pattern of sensitivePatterns) check(!pattern.test(source.toString()), `sensitive pattern found: ${pattern}`);

for (const row of data.organizations) check(html.includes(`id:"${row.id}"`), `HTML missing organization ${row.id}`);
for (const row of data.platformGroups) check(html.includes(`id:"${row.id}"`), `HTML missing group ${row.id}`);
check(html.includes('surveyParticipation:62'), 'HTML missing USUnions aggregate snapshot');
check(html.includes('data/union-atlas.json'), 'HTML does not identify its local standalone data package');
check(!/\b(?:src|href)=["']https?:\/\//i.test(html), 'HTML has an external runtime asset dependency');

const digest = crypto.createHash('sha256').update(source).digest('hex');
check(manifest.sha256 === digest, 'manifest SHA-256 does not match union-atlas.json');
check(manifest.counts.organizations === data.organizations.length, 'manifest organization count mismatch');
check(manifest.counts.platformGroups === data.platformGroups.length, 'manifest group count mismatch');

if (failures.length) {
  failures.forEach(message => console.error(`FAIL ${message}`));
  process.exit(1);
}
console.log(`PASS standalone Atlas data: ${data.organizations.length} organizations, ${data.platformGroups.length} groups, SHA-256 ${digest}`);
