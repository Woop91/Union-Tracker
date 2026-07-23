import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const args = process.argv.slice(2);
const valueAfter = flag => {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : null;
};

const sourceRoot = path.resolve(
  valueAfter('--source') || process.env.USUNIONS_SOURCE || ''
);
const zctaPath = path.resolve(
  valueAfter('--zcta') || process.env.CENSUS_ZCTA_SOURCE || ''
);
const outputPath = path.join(root, 'data', 'union-atlas.json');

if (!sourceRoot || !fs.existsSync(path.join(sourceRoot, 'data', 'unions'))) {
  throw new Error('Pass --source <USUnions repo> or set USUNIONS_SOURCE.');
}
if (!zctaPath || !fs.existsSync(zctaPath)) {
  throw new Error('Pass --zcta <Census Gazetteer ZCTA text file> or set CENSUS_ZCTA_SOURCE.');
}

const sourceCommit = execFileSync(
  'git',
  ['-C', sourceRoot, 'rev-parse', 'HEAD'],
  { encoding: 'utf8' }
).trim();
const sourceManifestPath = path.join(sourceRoot, 'data', 'manifest.jsonl');
const sourceManifest = fs.readFileSync(sourceManifestPath);

const zctas = Object.create(null);
for (const line of fs.readFileSync(zctaPath, 'utf8').split(/\r?\n/).slice(1)) {
  if (!line.trim()) continue;
  const [zip, , , , , , latitude, longitude] = line.split('|');
  if (/^\d{5}$/.test(zip) && latitude && longitude) {
    zctas[zip] = [Number(latitude), Number(longitude)];
  }
}

const unionFiles = [];
const unionRoot = path.join(sourceRoot, 'data', 'unions');
for (const affiliation of fs.readdirSync(unionRoot).sort()) {
  const directory = path.join(unionRoot, affiliation);
  if (!fs.statSync(directory).isDirectory()) continue;
  for (const file of fs.readdirSync(directory).sort()) {
    if (file.endsWith('.json')) unionFiles.push(path.join(directory, file));
  }
}

const records = unionFiles.map(file => JSON.parse(fs.readFileSync(file, 'utf8')));
const normalizeZip = value => {
  const match = String(value || '').match(/\d{5}/);
  return match ? match[0] : null;
};

const directory = records.map(record => {
  const zip = normalizeZip(record.hq?.zip);
  const point = zip ? zctas[zip] : null;
  return {
    id: record.id,
    olmsFileNumber: record.olms_file_number,
    name: record.name,
    affiliation: record.aff_abbr || null,
    designation: record.designation || null,
    localNumber: record.designation_number || null,
    level: record.level,
    parentOlmsFileNumber: record.parent?.olms_file_number || null,
    city: record.hq?.city || null,
    state: record.hq?.state || null,
    zip,
    latitude: point?.[0] ?? null,
    longitude: point?.[1] ?? null,
    website: record.website || null,
    members: record.membership?.latest?.count ?? null,
    membershipYear: record.membership?.latest?.year ?? null,
    receipts: record.finances?.latest?.total_receipts ?? null,
    filingForm: record.finances?.latest?.form || null,
    enrichmentStatus: record.enrichment.status,
    updatedAt: record.updated_at
  };
});

const rollup = (rows, keyFor) => {
  const groups = new Map();
  for (const row of rows) {
    const key = keyFor(row);
    if (!key) continue;
    const current = groups.get(key) || {
      key,
      unions: 0,
      reportedMembers: 0,
      membershipRecords: 0,
      geocoded: 0,
      websites: 0
    };
    current.unions += 1;
    if (Number.isFinite(row.members)) {
      current.reportedMembers += row.members;
      current.membershipRecords += 1;
    }
    if (Number.isFinite(row.latitude) && Number.isFinite(row.longitude)) current.geocoded += 1;
    if (row.website) current.websites += 1;
    groups.set(key, current);
  }
  return [...groups.values()].sort((a, b) => b.unions - a.unions || a.key.localeCompare(b.key));
};

const uniqueIds = new Set(directory.map(row => row.id));
const quality = {
  records: directory.length,
  uniqueIds: uniqueIds.size,
  withMembership: directory.filter(row => Number.isFinite(row.members)).length,
  withWebsites: directory.filter(row => Boolean(row.website)).length,
  withZctaCoordinates: directory.filter(row => Number.isFinite(row.latitude)).length,
  enrichmentStatuses: Object.fromEntries(
    [...new Set(directory.map(row => row.enrichmentStatus))]
      .sort()
      .map(status => [status, directory.filter(row => row.enrichmentStatus === status).length])
  )
};

if (quality.records !== 20699 || quality.uniqueIds !== quality.records) {
  throw new Error(`Unexpected USUnions source shape: ${quality.records} rows, ${quality.uniqueIds} unique IDs.`);
}

const data = {
  metadata: {
    formatVersion: 2,
    project: 'Common Ground Union Atlas',
    selfContained: true,
    generatedAt: new Date().toISOString(),
    dataType: 'public-filing union directory',
    recordGrain: 'one canonical OLMS union record',
    privacy: 'Organization-level public filing data only. No member, grievance, authentication, case, or private-contact records.',
    sources: [
      {
        name: 'USUnions canonical directory',
        repository: 'Woop91/USUnions',
        commit: sourceCommit,
        path: 'data/unions/**/*.json',
        manifestSha256: crypto.createHash('sha256').update(sourceManifest).digest('hex')
      },
      {
        name: '2025 U.S. Census Gazetteer ZIP Code Tabulation Areas',
        url: 'https://www2.census.gov/geo/docs/maps-data/data/gazetteer/2025_Gazetteer/2025_Gaz_zcta_national.zip',
        path: path.basename(zctaPath),
        purpose: 'Representative latitude and longitude for distance and bearing calculations'
      }
    ]
  },
  quality,
  affiliations: rollup(directory, row => row.affiliation),
  states: rollup(directory, row => row.state),
  zctas,
  directory
};

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, `${JSON.stringify(data)}\n`, 'utf8');
console.log(
  `Imported ${quality.records} source-backed unions; ${quality.withZctaCoordinates} have Census ZCTA coordinates.`
);
