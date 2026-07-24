import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const manifestPath = path.join(root, 'data', 'manifest.json');
const browserDataPath = path.join(root, 'docs', 'union-atlas-data.js');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);

const manifest = {
  formatVersion: data.metadata.formatVersion,
  dataFile: 'union-atlas.json',
  browserDataFile: '../docs/union-atlas-data.js',
  sha256: crypto.createHash('sha256').update(source).digest('hex'),
  counts: {
    unions: data.directory.length,
    uniqueIds: data.quality.uniqueIds,
    affiliations: data.affiliations.length,
    statesAndTerritories: data.states.length,
    zctaCoordinates: Object.keys(data.zctas).length,
    geocodedUnions: data.quality.withZctaCoordinates,
    membershipRecords: data.quality.withMembership,
    websites: data.quality.withWebsites
  },
  source: data.metadata.sources[0],
  coordinateSource: data.metadata.sources[1],
  selfContained: data.metadata.selfContained,
  syntheticRecords: 0
};

fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
fs.writeFileSync(
  browserDataPath,
  `/* Generated from data/union-atlas.json. Do not edit manually. */\nwindow.UNION_ATLAS_DATA=${JSON.stringify(data)};\n`,
  'utf8'
);
console.log(`Built source-backed Atlas package: ${manifest.sha256}`);
