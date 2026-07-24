import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import {
  deterministicGzip,
  readGzipJsonLines,
  readJson,
  sha256,
  writeJson
} from './atlas-source-utils.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const manifestPath = path.join(root, 'data', 'manifest.json');
const sourcePackageRoot = path.join(root, 'data', 'source', 'usunions');
const browserDirectoryPath = path.join(root, 'docs', 'union-atlas-data.json.gz');
const browserDetailsPath = path.join(root, 'docs', 'union-atlas-details.json.gz');
const browserEvidencePath = path.join(root, 'docs', 'union-atlas-evidence.json.gz');
const browserIndexPath = path.join(root, 'docs', 'union-atlas-index.js');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);
const sourcePackage = readJson(path.join(sourcePackageRoot, 'manifest.json'));

const sourcePath = relativePath => path.join(sourcePackageRoot, ...relativePath.split('/'));
const readMaybeGzipJson = relativePath => {
  const file = sourcePath(relativePath);
  const buffer = relativePath.endsWith('.gz')
    ? zlib.gunzipSync(fs.readFileSync(file))
    : fs.readFileSync(file);
  return JSON.parse(buffer.toString('utf8'));
};

const organizingAttempts = readGzipJsonLines(
  sourcePath('data/organizing/attempts.jsonl.gz')
);
const organizingGroups = readGzipJsonLines(
  sourcePath('data/organizing/groups.jsonl.gz')
);
const organizingSummary = readMaybeGzipJson('data/organizing/summary.json');
const organizingTrends = readMaybeGzipJson('data/organizing/trends.json.gz');
const researchArticles = readGzipJsonLines(
  sourcePath('data/research/articles.jsonl.gz')
);
const researchSources = readGzipJsonLines(
  sourcePath('data/research/sources.jsonl.gz')
);
const researchStatistics = readGzipJsonLines(
  sourcePath('data/research/statistics.jsonl.gz')
);
const researchSummary = readMaybeGzipJson('data/research/summary.json');

const directoryPayload = {
  metadata: data.metadata,
  quality: data.quality,
  zctas: data.zctas,
  directory: data.directory.map(row => ({
    id: row.id,
    olmsFileNumber: row.olmsFileNumber,
    name: row.name,
    affiliation: row.affiliation,
    designation: row.designation,
    localNumber: row.localNumber,
    level: row.level,
    parent: row.parent,
    sector: row.sector,
    city: row.city,
    state: row.state,
    zip: row.zip,
    latitude: row.latitude,
    longitude: row.longitude,
    geocodePrecision: row.geocodePrecision,
    website: row.website,
    membership: { latest: row.membership.latest },
    finances: row.finances,
    updatedAt: row.updatedAt
  }))
};
const detailsPayload = {
  metadata: {
    formatVersion: 1,
    source: data.metadata.sources[0],
    privacy: data.metadata.privacy
  },
  records: Object.fromEntries(data.directory.map(row => [row.id, {
    membershipHistory: row.membership.history,
    geocodeMatchType: row.geocodeMatchType,
    publicRecordCoverage: row.publicRecordCoverage,
    background: row.background,
    sources: row.sources
  }]))
};
const evidencePayload = {
  metadata: {
    formatVersion: 1,
    source: data.metadata.sources[0],
    privacy: sourcePackage.privacy
  },
  organizing: {
    summary: organizingSummary,
    trends: organizingTrends,
    attempts: organizingAttempts,
    groups: organizingGroups
  },
  research: {
    summary: researchSummary,
    articles: researchArticles,
    sources: researchSources,
    statistics: researchStatistics
  }
};
const directoryJson = JSON.stringify(directoryPayload);
const detailsJson = JSON.stringify(detailsPayload);
const evidenceJson = JSON.stringify(evidencePayload);
const directoryGzip = deterministicGzip(directoryJson);
const detailsGzip = deterministicGzip(detailsJson);
const evidenceGzip = deterministicGzip(evidenceJson);
fs.writeFileSync(browserDirectoryPath, directoryGzip);
fs.writeFileSync(browserDetailsPath, detailsGzip);
fs.writeFileSync(browserEvidencePath, evidenceGzip);

const browserIndex = {
  metadata: data.metadata,
  quality: data.quality,
  datasets: data.datasets,
  affiliations: data.affiliations,
  states: data.states,
  packages: {
    directory: {
      path: 'union-atlas-data.json.gz',
      bytes: directoryGzip.length,
      uncompressedBytes: Buffer.byteLength(directoryJson),
      sha256: sha256(directoryGzip)
    },
    details: {
      path: 'union-atlas-details.json.gz',
      bytes: detailsGzip.length,
      uncompressedBytes: Buffer.byteLength(detailsJson),
      sha256: sha256(detailsGzip)
    },
    evidence: {
      path: 'union-atlas-evidence.json.gz',
      bytes: evidenceGzip.length,
      uncompressedBytes: Buffer.byteLength(evidenceJson),
      sha256: sha256(evidenceGzip)
    }
  }
};
fs.writeFileSync(
  browserIndexPath,
  `/* Generated source-backed Atlas index. */\nwindow.UNION_ATLAS_INDEX=${JSON.stringify(browserIndex)};\n`,
  'utf8'
);

const manifest = {
  formatVersion: data.metadata.formatVersion,
  dataFile: 'union-atlas.json',
  sha256: sha256(source),
  counts: {
    unions: data.directory.length,
    uniqueIds: data.quality.uniqueIds,
    affiliations: data.affiliations.length,
    statesAndTerritories: data.states.length,
    zctaCoordinates: Object.keys(data.zctas).length,
    geocodedUnions: data.quality.withCoordinates,
    addressGeocodedUnions: data.quality.withAddressCoordinates,
    zctaGeocodedUnions: data.quality.withZctaCoordinates,
    membershipRecords: data.quality.withMembership,
    websites: data.quality.withWebsites,
    organizingCases: organizingAttempts.length,
    organizingGroups: organizingGroups.length,
    researchArticles: researchArticles.length,
    researchStatistics: researchStatistics.length,
    researchSources: researchSources.length
  },
  source: data.metadata.sources[0],
  coordinateSources: data.metadata.sources.slice(1),
  browserPackages: browserIndex.packages,
  repositoryLocal: data.metadata.repositoryLocal,
  syntheticRecords: 0,
  privacy: data.metadata.privacy
};

writeJson(manifestPath, manifest);
console.log(
  `Built Atlas packages: directory ${(directoryGzip.length / 1048576).toFixed(2)} MiB, `
  + `details ${(detailsGzip.length / 1048576).toFixed(2)} MiB, `
  + `evidence ${(evidenceGzip.length / 1048576).toFixed(2)} MiB.`
);
