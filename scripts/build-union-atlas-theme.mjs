import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const tokensPath = path.join(root, 'design', 'solid-ground', 'tokens.css');
const atlasPath = path.join(root, 'docs', 'union-atlas.html');
const start = '<!-- BEGIN SOLID GROUND TOKENS -->';
const end = '<!-- END SOLID GROUND TOKENS -->';
const tokens = fs.readFileSync(tokensPath, 'utf8').trim();
const html = fs.readFileSync(atlasPath, 'utf8');
const pattern = new RegExp(`${start}[\\s\\S]*?${end}`);

if (!pattern.test(html)) throw new Error('Solid Ground token markers missing from Atlas HTML');
const block = `${start}\n<style id="solid-ground-tokens">\n${tokens}\n</style>\n${end}`;
fs.writeFileSync(atlasPath, html.replace(pattern, block), 'utf8');
console.log(`Inlined ${tokensPath} into ${atlasPath}`);
