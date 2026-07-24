import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { readJson, sha256, writeJson } from './atlas-source-utils.mjs';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const args = process.argv.slice(2);
const valueAfter = flag => {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : null;
};
const sourceRoot = path.resolve(
  valueAfter('--source') || process.env.USUNIONS_SOURCE || ''
);
const outputPath = path.resolve(
  valueAfter('--output') || path.join(root, 'data', 'geocodes', 'census-address-geocodes.json')
);
const force = args.includes('--force');
const batchSize = Math.min(Number(valueAfter('--batch-size') || 9000), 10000);

if (!sourceRoot || !fs.existsSync(path.join(sourceRoot, 'data', 'unions'))) {
  throw new Error('Pass --source <USUnions repo> or set USUNIONS_SOURCE.');
}

const sourceCommit = execFileSync(
  'git',
  ['-C', sourceRoot, 'rev-parse', 'HEAD'],
  { encoding: 'utf8' }
).trim();

const normalize = value => String(value || '').trim().replace(/\s+/g, ' ');
const fingerprint = record => sha256([
  normalize(record.hq?.street).toUpperCase(),
  normalize(record.hq?.city).toUpperCase(),
  normalize(record.hq?.state).toUpperCase(),
  normalize(record.hq?.zip).toUpperCase()
].join('|'));

const records = [];
const unionRoot = path.join(sourceRoot, 'data', 'unions');
for (const affiliation of fs.readdirSync(unionRoot).sort()) {
  const directory = path.join(unionRoot, affiliation);
  if (!fs.statSync(directory).isDirectory()) continue;
  for (const file of fs.readdirSync(directory).sort()) {
    if (!file.endsWith('.json')) continue;
    const record = readJson(path.join(directory, file));
    if (!record.hq?.street) continue;
    records.push({
      id: record.id,
      street: normalize(record.hq.street),
      city: normalize(record.hq.city),
      state: normalize(record.hq.state),
      zip: normalize(record.hq.zip).match(/\d{5}/)?.[0] || '',
      sourceFingerprint: fingerprint(record)
    });
  }
}

const previous = fs.existsSync(outputPath) ? readJson(outputPath) : { geocodes: {} };
const geocodes = force ? {} : { ...(previous.geocodes || {}) };
const pending = records.filter(record => (
  force || geocodes[record.id]?.sourceFingerprint !== record.sourceFingerprint
));

const csvEscape = value => {
  const string = String(value ?? '');
  return /[",\r\n]/.test(string) ? `"${string.replaceAll('"', '""')}"` : string;
};

const parseCsvLine = line => {
  const fields = [];
  let value = '';
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    if (character === '"') {
      if (quoted && line[index + 1] === '"') {
        value += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (character === ',' && !quoted) {
      fields.push(value);
      value = '';
    } else {
      value += character;
    }
  }
  fields.push(value);
  return fields;
};

const endpoint = 'https://geocoding.geo.census.gov/geocoder/locations/addressbatch';
let matched = 0;
let unmatched = 0;
for (let offset = 0; offset < pending.length; offset += batchSize) {
  const batch = pending.slice(offset, offset + batchSize);
  const csv = `${batch.map(record => [
    record.id,
    record.street,
    record.city,
    record.state,
    record.zip
  ].map(csvEscape).join(',')).join('\n')}\n`;
  const form = new FormData();
  form.append('benchmark', 'Public_AR_Current');
  form.append('addressFile', new Blob([csv], { type: 'text/csv' }), 'union-headquarters.csv');
  const response = await fetch(endpoint, { method: 'POST', body: form });
  if (!response.ok) {
    throw new Error(`Census batch geocoder failed: HTTP ${response.status}`);
  }
  const responseText = await response.text();
  const byId = new Map(batch.map(record => [record.id, record]));
  for (const line of responseText.split(/\r?\n/).filter(Boolean)) {
    const fields = parseCsvLine(line);
    const id = fields[0]?.replace(/^"|"$/g, '');
    const source = byId.get(id);
    if (!source) continue;
    const status = fields[2];
    const coordinates = fields[5]?.split(',').map(Number);
    if (status === 'Match' && coordinates?.length === 2 && coordinates.every(Number.isFinite)) {
      geocodes[id] = {
        latitude: coordinates[1],
        longitude: coordinates[0],
        matchType: fields[3] || null,
        tigerLineId: fields[6] || null,
        sourceFingerprint: source.sourceFingerprint
      };
      matched += 1;
    } else {
      geocodes[id] = {
        latitude: null,
        longitude: null,
        matchType: status || 'No_Match',
        tigerLineId: null,
        sourceFingerprint: source.sourceFingerprint
      };
      unmatched += 1;
    }
  }
  console.log(`Census geocoder batch ${Math.floor(offset / batchSize) + 1}: ${batch.length} submitted.`);
}

const output = {
  formatVersion: 1,
  sourceRepository: 'Woop91/USUnions',
  sourceCommit,
  generatedAt: new Date().toISOString(),
  geocoder: {
    name: 'U.S. Census Geocoding Services',
    endpoint,
    benchmark: 'Public_AR_Current',
    documentation: 'https://geocoding.geo.census.gov/geocoder/Geocoding_Services_API.html'
  },
  privacy: 'Public filing street addresses were sent to the Census geocoder but are not stored in this file.',
  counts: {
    sourceRecords: records.length,
    matched: Object.values(geocodes).filter(row => Number.isFinite(row.latitude)).length,
    unmatched: Object.values(geocodes).filter(row => !Number.isFinite(row.latitude)).length,
    refreshed: pending.length,
    refreshedMatched: matched,
    refreshedUnmatched: unmatched
  },
  geocodes
};
writeJson(outputPath, output);
console.log(
  `Stored ${output.counts.matched} address geocodes; ${output.counts.unmatched} unmatched. No street addresses persisted.`
);
