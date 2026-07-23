import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const html = fs.readFileSync(path.join(here, '..', 'docs', 'union-atlas.html'), 'utf8');
const root = html.match(/:root\s*\{([\s\S]*?)\}/)?.[1] || '';
const tokens = Object.fromEntries(
  [...root.matchAll(/--([\w-]+)\s*:\s*([^;]+);/g)].map(([, key, value]) => [key, value.trim()])
);

function resolve(value, seen = new Set()) {
  const match = value?.match(/^var\(--([\w-]+)\)$/);
  if (!match) return value;
  if (seen.has(match[1])) throw new Error(`Circular token: ${match[1]}`);
  seen.add(match[1]);
  return resolve(tokens[match[1]], seen);
}

function luminance(hex) {
  const rgb = hex.replace('#', '').match(/.{2}/g).map(x => parseInt(x, 16) / 255);
  const linear = rgb.map(c => c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
}

function ratio(foreground, background) {
  const a = luminance(resolve(tokens[foreground]));
  const b = luminance(resolve(tokens[background]));
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}

const pairs = [
  ['foreground', 'background', 4.5],
  ['card-foreground', 'card', 4.5],
  ['popover-foreground', 'popover', 4.5],
  ['primary-foreground', 'primary', 4.5],
  ['secondary-foreground', 'secondary', 4.5],
  ['muted-foreground', 'background', 4.5],
  ['accent-foreground', 'accent', 4.5],
  ['destructive-foreground', 'destructive', 4.5],
  ['focus-ring', 'background', 3],
  ['input', 'background', 3],
  ['border', 'card', 3]
];

let failed = false;
for (const [fg, bg, minimum] of pairs) {
  const value = ratio(fg, bg);
  const ok = value >= minimum;
  failed ||= !ok;
  console.log(`${ok ? 'PASS' : 'FAIL'} ${fg}/${bg}: ${value.toFixed(2)}:1 (min ${minimum}:1)`);
}
if (failed) process.exitCode = 1;
