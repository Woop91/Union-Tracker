import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const dataPath = path.join(root, 'data', 'union-atlas.json');
const manifestPath = path.join(root, 'data', 'manifest.json');
const source = fs.readFileSync(dataPath);
const data = JSON.parse(source);

const manifest = {
  formatVersion: data.metadata.formatVersion,
  dataFile: 'union-atlas.json',
  sha256: crypto.createHash('sha256').update(source).digest('hex'),
  counts: {
    organizations: data.organizations.length,
    platformGroups: data.platformGroups.length,
    sectors: Object.keys(data.taxonomies.sectors).length,
    issues: Object.keys(data.taxonomies.issues).length,
    suggestions: data.suggestions.length,
    coverageGaps: data.admin.coverageGaps.length,
    membershipTrendRows: data.usUnions.membershipTrends.length,
    publicIdentityFields: data.usUnions.publicIdentitySchema.length
  },
  selfContained: data.metadata.selfContained,
  productionData: data.usUnions.provenance.production
};

fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
console.log(`Built standalone data manifest: ${manifest.sha256}`);
