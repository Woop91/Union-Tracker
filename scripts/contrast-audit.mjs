import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const html = fs.readFileSync(path.join(here, '..', 'docs', 'union-atlas.html'), 'utf8');
const declarations = body => Object.fromEntries(
  [...body.matchAll(/--([\w-]+)\s*:\s*([^;]+);/g)].map(([, key, value]) => [key, value.trim()])
);
const rootBlocks = [...html.matchAll(/:root\s*\{([\s\S]*?)\}/g)].map(match => declarations(match[1]));
const base = Object.assign({}, ...rootBlocks);
const darkBody = html.match(/html\[data-theme="dark"\]\s*\{([\s\S]*?)\}/)?.[1] || '';
const themes = {
  light: { ...base },
  dark: { ...base, ...declarations(darkBody) }
};

function resolve(tokens, value, seen = new Set()) {
  const match = value?.match(/^var\(--([\w-]+)\)$/);
  if (!match) return value;
  if (seen.has(match[1])) throw new Error(`Circular token: ${match[1]}`);
  if (!tokens[match[1]]) throw new Error(`Missing token: ${match[1]}`);
  seen.add(match[1]);
  return resolve(tokens, tokens[match[1]], seen);
}

function luminance(hex) {
  if (!/^#[0-9a-f]{6}$/i.test(hex)) throw new Error(`Unsupported resolved color: ${hex}`);
  const rgb = hex.replace('#', '').match(/.{2}/g).map(x => parseInt(x, 16) / 255);
  const linear = rgb.map(c => c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2];
}

function contrast(tokens, foreground, background) {
  const fg = resolve(tokens, tokens[foreground]);
  const bg = resolve(tokens, tokens[background]);
  const a = luminance(fg);
  const b = luminance(bg);
  return { fg, bg, ratio: (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05) };
}

const pairs = [
  ['foreground', 'background', 4.5, 'page text'],
  ['card-foreground', 'card', 4.5, 'cards'],
  ['popover-foreground', 'popover', 4.5, 'modal and tooltip surfaces'],
  ['primary-foreground', 'primary', 4.5, 'selected controls'],
  ['secondary-foreground', 'secondary', 4.5, 'secondary panels'],
  ['muted-foreground', 'background', 4.5, 'informative captions'],
  ['muted-foreground', 'muted', 4.5, 'muted panels'],
  ['accent-foreground', 'accent', 4.5, 'gold actions'],
  ['destructive-foreground', 'destructive', 4.5, 'privacy and destructive states'],
  ['success-foreground', 'success', 4.5, 'success states'],
  ['warning-foreground', 'warning', 4.5, 'warning states'],
  ['foreground', 'muted', 4.5, 'table hover and quiet panels'],
  ['foreground', 'background', 7, 'preferred page-body target'],
  ['input', 'background', 3, 'input outline'],
  ['border', 'card', 3, 'meaningful card boundaries'],
  ['ring', 'background', 3, 'focus indicator']
];

let failed = false;
for (const [theme, tokens] of Object.entries(themes)) {
  console.log(`\n${theme.toUpperCase()} THEME`);
  for (const [fgName, bgName, minimum, usage] of pairs) {
    const result = contrast(tokens, fgName, bgName);
    const ok = result.ratio >= minimum;
    failed ||= !ok;
    const aa = result.ratio >= 4.5 ? 'AA' : result.ratio >= 3 ? '3:1' : 'FAIL';
    const aaa = result.ratio >= 7 ? 'AAA' : 'no-AAA';
    console.log(`${ok ? 'PASS' : 'FAIL'} ${fgName}/${bgName} ${result.fg}/${result.bg} ${result.ratio.toFixed(2)}:1 ${aa} ${aaa} — ${usage}`);
  }
}
if (failed) process.exitCode = 1;
