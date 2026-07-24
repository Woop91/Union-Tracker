import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
  copySourceFile,
  readJson,
  sha256,
  writeJson
} from './atlas-source-utils.mjs';

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
const geocodePath = path.resolve(
  valueAfter('--geocodes')
    || process.env.CENSUS_ADDRESS_GEOCODES
    || path.join(root, 'data', 'geocodes', 'census-address-geocodes.json')
);
const outputPath = path.join(root, 'data', 'union-atlas.json');
const sourcePackageRoot = path.join(root, 'data', 'source', 'usunions');

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
const addressGeocodes = fs.existsSync(geocodePath) ? readJson(geocodePath) : {
  geocodes: {},
  counts: { matched: 0, unmatched: 0 }
};

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

const records = unionFiles.map(file => readJson(file));
const normalizeZip = value => String(value || '').match(/\d{5}/)?.[0] || null;
const contactEmail = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const contactPhone = /(?:\+?1[-.\s]?)?(?:\(\d{3}\)|\d{3})[-.\s]\d{3}[-.\s]\d{4}/g;
const cleanPublicText = value => value
  ? String(value)
    .replace(contactEmail, '[contact removed]')
    .replace(contactPhone, '[contact removed]')
  : null;
const cleanPublicUrl = value => {
  if (!/^https?:\/\//i.test(value || '')) return null;
  let decoded = String(value);
  try {
    decoded = decodeURIComponent(decoded);
  } catch {
    // Keep the original URL when malformed percent-encoding prevents decoding.
  }
  return /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(decoded)
    ? null
    : String(value);
};
const cleanSources = sources => (sources || [])
  .map(source => ({
    url: cleanPublicUrl(source.url),
    retrievedAt: source.retrieved_at || null
  }))
  .filter(source => source.url)
  .map(source => ({
    ...source
  }));

const directory = records.map(record => {
  const zip = normalizeZip(record.hq?.zip);
  const addressPoint = addressGeocodes.geocodes?.[record.id];
  const zctaPoint = zip ? zctas[zip] : null;
  const hasAddressPoint = Number.isFinite(addressPoint?.latitude)
    && Number.isFinite(addressPoint?.longitude);
  const point = hasAddressPoint
    ? [addressPoint.latitude, addressPoint.longitude]
    : zctaPoint;
  const latestFinances = record.finances?.latest || null;
  const background = cleanPublicText(record.background);
  return {
    id: record.id,
    olmsFileNumber: record.olms_file_number,
    name: record.name,
    affiliation: record.aff_abbr || null,
    designation: record.designation || null,
    localNumber: record.designation_number || null,
    level: record.level,
    parent: record.parent ? {
      affiliation: record.parent.aff_abbr || null,
      olmsFileNumber: record.parent.olms_file_number || null,
      name: record.parent.name || null
    } : null,
    sector: {
      publicPrivate: record.sector?.public_private || null,
      tags: record.sector?.tags || []
    },
    city: record.hq?.city || null,
    state: record.hq?.state || null,
    zip,
    latitude: point?.[0] ?? null,
    longitude: point?.[1] ?? null,
    geocodePrecision: hasAddressPoint ? 'address' : (zctaPoint ? 'zcta' : null),
    geocodeMatchType: hasAddressPoint ? addressPoint.matchType : null,
    website: cleanPublicUrl(record.website),
    membership: {
      latest: record.membership?.latest ? {
        count: record.membership.latest.count,
        year: record.membership.latest.year
      } : null,
      history: record.membership?.history || []
    },
    finances: latestFinances ? {
      fiscalYearEnd: latestFinances.fiscal_year_end || null,
      totalReceipts: latestFinances.total_receipts ?? null,
      totalDisbursements: latestFinances.total_disbursements ?? null,
      totalAssets: latestFinances.total_assets ?? null,
      form: latestFinances.form || null
    } : null,
    publicRecordCoverage: {
      filedOfficers: record.officers_olms?.length || 0,
      publishedStaff: record.staff_web?.length || 0,
      publishedStewards: record.stewards_web?.length || 0,
      chapters: record.chapters?.length || 0,
      contracts: record.contracts?.length || 0,
      sources: record.sources?.length || 0,
      background: Boolean(background)
    },
    background,
    sources: cleanSources(record.sources),
    enrichmentStatus: record.enrichment?.status || null,
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
      addressGeocoded: 0,
      zctaGeocoded: 0,
      websites: 0
    };
    current.unions += 1;
    if (Number.isFinite(row.membership.latest?.count)) {
      current.reportedMembers += row.membership.latest.count;
      current.membershipRecords += 1;
    }
    if (row.geocodePrecision === 'address') current.addressGeocoded += 1;
    if (row.geocodePrecision === 'zcta') current.zctaGeocoded += 1;
    if (row.website) current.websites += 1;
    groups.set(key, current);
  }
  return [...groups.values()].sort(
    (a, b) => b.unions - a.unions || a.key.localeCompare(b.key)
  );
};

const uniqueIds = new Set(directory.map(row => row.id));
const quality = {
  records: directory.length,
  uniqueIds: uniqueIds.size,
  withMembership: directory.filter(row => Number.isFinite(row.membership.latest?.count)).length,
  withWebsites: directory.filter(row => Boolean(row.website)).length,
  withCoordinates: directory.filter(row => Number.isFinite(row.latitude)).length,
  withAddressCoordinates: directory.filter(row => row.geocodePrecision === 'address').length,
  withZctaCoordinates: directory.filter(row => row.geocodePrecision === 'zcta').length,
  withoutCoordinates: directory.filter(row => !Number.isFinite(row.latitude)).length,
  withoutZip: directory.filter(row => !row.zip).length,
  enrichmentStatuses: Object.fromEntries(
    [...new Set(directory.map(row => row.enrichmentStatus))]
      .filter(Boolean)
      .sort()
      .map(status => [status, directory.filter(row => row.enrichmentStatus === status).length])
  )
};

if (quality.records !== 20699 || quality.uniqueIds !== quality.records) {
  throw new Error(
    `Unexpected USUnions source shape: ${quality.records} rows, ${quality.uniqueIds} unique IDs.`
  );
}

const evidenceFiles = [
  'data/organizing/attempts.jsonl',
  'data/organizing/groups.jsonl',
  'data/organizing/summary.json',
  'data/organizing/trends.json',
  'data/research/articles.jsonl',
  'data/research/sources.jsonl',
  'data/research/statistics.jsonl',
  'data/research/summary.json'
];
const evidenceManifest = evidenceFiles.map(relativePath => copySourceFile(
  sourceRoot,
  relativePath,
  sourcePackageRoot,
  relativePath.endsWith('.jsonl') || relativePath.endsWith('trends.json')
));
const organizingSummary = readJson(path.join(sourceRoot, 'data', 'organizing', 'summary.json'));
const researchSummary = readJson(path.join(sourceRoot, 'data', 'research', 'summary.json'));
const sourcePackage = {
  formatVersion: 1,
  repository: 'Woop91/USUnions',
  commit: sourceCommit,
  generatedAt: new Date().toISOString(),
  privacy: 'Organization-level public filing, NLRB case, and research data only. Private contacts, authentication data, and personal records are excluded.',
  files: evidenceManifest
};
writeJson(path.join(sourcePackageRoot, 'manifest.json'), sourcePackage);

const data = {
  metadata: {
    formatVersion: 3,
    project: 'Common Ground Union Atlas',
    repositoryLocal: true,
    generatedAt: new Date().toISOString(),
    dataType: 'public union directory, organizing evidence, and labor research',
    recordGrain: 'one canonical OLMS union record',
    privacy: 'Organization-level public records only. Street addresses are used transiently for Census geocoding but are not stored. Member, grievance, authentication, private-contact, and personal records remain excluded.',
    sources: [
      {
        name: 'USUnions canonical public-data repository',
        repository: 'Woop91/USUnions',
        commit: sourceCommit,
        paths: ['data/unions/**/*.json', 'data/organizing/**', 'data/research/**'],
        manifestSha256: sha256(sourceManifest)
      },
      {
        name: 'U.S. Census Geocoding Services',
        url: 'https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html',
        benchmark: addressGeocodes.geocoder?.benchmark || 'Public_AR_Current',
        purpose: 'Address-range coordinates for public union headquarters; street addresses are not persisted.'
      },
      {
        name: '2025 U.S. Census Gazetteer ZIP Code Tabulation Areas',
        url: 'https://www2.census.gov/geo/docs/maps-data/data/gazetteer/2025_Gazetteer/2025_Gaz_zcta_national.zip',
        path: path.basename(zctaPath),
        purpose: 'Fallback representative coordinates when an address match is unavailable.'
      }
    ]
  },
  quality,
  datasets: {
    directory: {
      records: directory.length,
      coverage: 'Current canonical OLMS directory from the pinned USUnions commit.'
    },
    organizing: {
      records: organizingSummary.counts.representation_cases,
      rcAttempts: organizingSummary.counts.new_representation_attempts_rc,
      groups: organizingSummary.counts.reported_petitioner_group_rollups,
      filedFrom: organizingSummary.coverage.actual_case_filed_from,
      filedThrough: organizingSummary.coverage.actual_case_filed_through,
      jurisdiction: organizingSummary.coverage.jurisdiction,
      limitations: organizingSummary.limitations
    },
    research: {
      articles: researchSummary.counts?.articles
        ?? researchSummary.article_count
        ?? evidenceManifest.find(file => file.path.endsWith('articles.jsonl'))?.records,
      statistics: researchSummary.counts?.statistics
        ?? researchSummary.statistic_count
        ?? evidenceManifest.find(file => file.path.endsWith('statistics.jsonl'))?.records,
      sources: evidenceManifest.find(file => file.path.endsWith('sources.jsonl'))?.records
    }
  },
  affiliations: rollup(directory, row => row.affiliation),
  states: rollup(directory, row => row.state),
  zctas,
  directory
};

writeJson(outputPath, data, true);
console.log(
  `Imported ${quality.records} unions, ${data.datasets.organizing.records} organizing cases, `
  + `${data.datasets.research.articles} research articles; ${quality.withAddressCoordinates} address-geocoded.`
);
