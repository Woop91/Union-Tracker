import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const htmlPath = path.join(root, 'docs', 'union-atlas.html');
const outputPath = path.join(root, 'data', 'union-atlas.json');
const html = fs.readFileSync(htmlPath, 'utf8');

function extractConst(name) {
  const marker = `const ${name}`;
  const start = html.indexOf(marker);
  if (start < 0) throw new Error(`Missing ${marker}`);
  const equals = html.indexOf('=', start + marker.length);
  let quote = '';
  let escaped = false;
  let depth = 0;
  for (let i = equals + 1; i < html.length; i += 1) {
    const char = html[i];
    if (quote) {
      if (escaped) escaped = false;
      else if (char === '\\') escaped = true;
      else if (char === quote) quote = '';
      continue;
    }
    if (char === '"' || char === "'" || char === '`') {
      quote = char;
      continue;
    }
    if ('([{'.includes(char)) depth += 1;
    else if (')]}'.includes(char)) depth -= 1;
    else if (char === ';' && depth === 0) {
      const expression = html.slice(equals + 1, i).trim();
      return vm.runInNewContext(`(${expression})`, Object.create(null), { timeout: 1000 });
    }
  }
  throw new Error(`Unterminated ${name}`);
}

const embeddedSnapshot = extractConst('USUNIONS_SNAPSHOT');
const snapshot = {
  ...embeddedSnapshot,
  provenance: {
    ...embeddedSnapshot.provenance,
    repository: 'Woop91/SolidBase',
    sourceFile: 'src/07_DevTools.gs',
    sourceFunction: 'seedUnionStatsData',
    mode: 'repo-demo',
    production: false
  },
  publicIdentitySchema: [
  { key: 'ORG_NAME', label: 'Organization Name', type: 'text' },
  { key: 'ORG_ABBREV', label: 'Organization Abbreviation', type: 'text' },
  { key: 'LOCAL_NUMBER', label: 'Local Number', type: 'text' },
  { key: 'UNION_PARENT', label: 'Parent Union', type: 'text' },
  { key: 'STATE_REGION', label: 'State or Region', type: 'text' },
  { key: 'ORG_WEBSITE', label: 'Public Organization Website', type: 'url' },
    { key: 'OFFICE_LOCATIONS', label: 'Public Office Locations', type: 'list' }
  ]
};

const data = {
  metadata: {
    formatVersion: 1,
    project: 'Common Ground Union Atlas',
    selfContained: true,
    generatedAt: '2026-07-23',
    sources: [
      'docs/union-atlas.html legacy embedded dataset',
      'src/07_DevTools.gs seedUnionStatsData',
      'src/01_Core.gs CONFIG_HEADER_MAP_'
    ],
    notice: 'Atlas organizations are public bodies. Platform activity and USUnions metrics are demo data.',
    privacy: 'No member roster, grievance, contact, address, authentication, or case-level data.'
  },
  taxonomies: {
    sectors: extractConst('SECTORS'),
    issues: extractConst('ISSUES'),
    organizationTypes: extractConst('TYPE_LABEL')
  },
  organizations: extractConst('NODES'),
  platformGroups: extractConst('GROUPS'),
  suggestions: extractConst('SUGGESTIONS'),
  admin: {
    tiles: [
      { label: 'Groups placed', value: '214', delta: '+18 this month' },
      { label: 'Placement rate', value: '82%', delta: '+4 pts' },
      { label: 'Members mapped', value: '6,430', delta: '+310' },
      { label: 'Cross-group links', value: '391', delta: '+42' }
    ],
    sectorDistribution: extractConst('DIST_SECTOR'),
    regionDistribution: extractConst('DIST_REGION'),
    coverageGaps: extractConst('WHITESPACE')
  },
  network: {
    mutualAid: {
      id: 'aid1',
      kind: 'Mutual aid · turnout ask',
      title: 'Informational picket at Baystate — Thu 5pm',
      sourceGroupId: 'g_mna',
      issue: 'Understaffing'
    },
    jointAction: {
      id: 'act1',
      kind: 'Joint action',
      title: 'Pack the State House caseloads hearing',
      sourceGroupId: 'g_dds',
      participatingGroups: 5
    },
    alliance: {
      name: 'Massachusetts Human Services Table',
      type: 'Sector council',
      combinedMembers: '900+',
      unionFamilies: 3,
      workerCenters: 1,
      groupIds: ['g_dds', 'g_cpc', 'g_1199', 'g_pvwc', 'g_ghjc']
    }
  },
  usUnions: snapshot
};

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, `${JSON.stringify(data, null, 2)}\n`, 'utf8');
console.log(`Imported ${data.organizations.length} organizations, ${data.platformGroups.length} groups, and the complete safe USUnions snapshot.`);
