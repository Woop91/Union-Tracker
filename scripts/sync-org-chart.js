#!/usr/bin/env node
/**
 * sync-org-chart.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Fetches MADDS.html from Woop91/509d and transforms it into the SPA-compatible
 * org_chart.html, then commits it to ALL branches of BOTH repos:
 *
 *   DDS-Dashboard (Woop91/DDS-Dashboard) → Main, staging
 *   Union-Tracker (Woop91/Union-Tracker) → Main, staging
 *
 * MADDS is the only org chart across all repos and branches.
 *
 * Usage:  npm run sync-org-chart
 *
 * REQUIRES: .env at repo root containing:
 *   GITHUB_509D_TOKEN=ghp_...   (PAT with repo scope)
 *
 * DOCUMENT STRUCTURE WARNING:
 *   In MADDS.html, <head> wraps the ENTIRE <style> block (closes AFTER </style>).
 *   Never use <head>...</head> greedy match — it strips all CSS.
 *   Strip individual preamble tags instead, then strip </head> separately.
 *
 * TRANSFORMATION ORDER (steps 8+9 MUST precede step 10)
 *  1.  Strip leading block comment
 *  2.  Strip <!DOCTYPE>, <html>
 *  3.  Strip preamble inside <head>: <meta charset>, <meta viewport>, <title>, Fonts <link>
 *  4.  Strip </head>
 *  5.  CSS: :root { → .madds-embed {
 *  6.  CSS: html, body { → .madds-embed {
 *  7.  CSS: body.light → .madds-embed.light
 *  8.  HTML: replace <body> open + first button  [BEFORE rename at step 10]
 *  9.  JS: toggleMode → maddstoggleMode          [BEFORE rename at step 10]
 * 10.  CSS: rename #mode-toggle → #madds-mode-toggle
 * 11.  CSS: position:fixed → position:sticky in #madds-mode-toggle block
 * 12.  Replace closing </script></body></html> → </script></div><font-loader>  [BEFORE html strip]
 * 13.  Strip residual </html> / </body>
 */

'use strict';

const https    = require('https');
const fs       = require('fs');
const path     = require('path');
const os       = require('os');
const { execSync } = require('child_process');

const SOURCE = { owner: 'Woop91', repo: '509d', file: 'org-chart/MADDS.html', branch: 'main' };
const TARGETS = [
  {
    owner: 'Woop91', repo: 'DDS-Dashboard',
    branches: ['Main', 'staging'],
    files: ['src/org_chart.html', 'dist/org_chart.html'],
  },
  {
    owner: 'Woop91', repo: 'Union-Tracker',
    branches: ['Main', 'staging'],
    files: ['src/org_chart.html', 'dist/org_chart.html'],
  },
];
const COMMIT_MSG = 'sync: update org_chart.html from 509d/MADDS.html';

function loadEnv() {
  const envFile = path.join(__dirname, '..', '.env');
  if (!fs.existsSync(envFile)) return;
  fs.readFileSync(envFile, 'utf8').split('\n').forEach(line => {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.+?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^['"]|['"]$/g, '');
  });
}

loadEnv();

const TOKEN = process.env.GITHUB_509D_TOKEN || process.env.GITHUB_TOKEN;
if (!TOKEN) {
  console.error('ERROR: No GitHub token. Add GITHUB_509D_TOKEN=ghp_... to .env');
  process.exit(1);
}

function fetchFile(owner, repo, filePath, branch, token) {
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.github.com',
      path:     `/repos/${owner}/${repo}/contents/${filePath}?ref=${branch}`,
      method:   'GET',
      headers: {
        'Authorization': `token ${token}`,
        'Accept':        'application/vnd.github.v3+json',
        'User-Agent':    'dds-dashboard-sync-org-chart/1.0',
      },
    }, res => {
      let body = '';
      res.on('data', c => { body += c; });
      res.on('end', () => {
        if (res.statusCode !== 200) return reject(new Error(`GitHub ${res.statusCode}: ${body.slice(0,200)}`));
        try {
          const json = JSON.parse(body);
          resolve({ content: Buffer.from(json.content, 'base64').toString('utf8'), sha: json.sha });
        } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

const FONT_LOADER = [
  '<script>',
  '(function() {',
  "  if (!document.querySelector('link[href*=\"DM+Serif\"]')) {",
  "    var lnk = document.createElement('link');",
  "    lnk.rel = 'stylesheet';",
  "    lnk.href = 'https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,700&display=swap';",
  '    document.head.appendChild(lnk);',
  '  }',
  '})();',
  '</script>',
].join('\n');

function transform(source) {
  let t = source;
  const moon = '\u{1F319} Dark';
  const sun  = '\u2600 Light';

  t = t.replace(/^<!--[\s\S]*?-->\s*/m, '');
  t = t.replace(/<!DOCTYPE html>\s*/i, '');
  t = t.replace(/<html[^>]*>\s*/i, '');
  t = t.replace(/<head>\s*/i, '');
  t = t.replace(/<meta charset="[^"]*">\s*/i, '');
  t = t.replace(/<meta name="viewport"[^>]*>\s*/i, '');
  t = t.replace(/<title>[^<]*<\/title>\s*/i, '');
  t = t.replace(/<link\s+href="https:\/\/fonts\.googleapis\.com[^"]*"[^>]*>\s*/g, '');
  t = t.replace(/\s*<\/head>\s*/i, '\n');
  t = t.replace(/^:root \{/m, '.madds-embed {');
  t = t.replace(/^html,\s*body \{/m, '.madds-embed {');
  t = t.replace(/body\.light/g, '.madds-embed.light');
  t = t.replace(
    '<body>\n<button id="mode-toggle" onclick="toggleMode()">',
    '<div class="madds-embed">\n<button id="madds-mode-toggle" onclick="maddstoggleMode()">'
  );
  t = t.replace(
    `function toggleMode() {\n  document.body.classList.toggle('light');\n  const btn = document.getElementById('mode-toggle');\n  btn.textContent = document.body.classList.contains('light') ? '${moon}' : '${sun}';\n}`,
    `function maddstoggleMode() {\n  var _me = document.querySelector('.madds-embed'); if(_me) _me.classList.toggle('light');;\n  const btn = document.getElementById('madds-mode-toggle');\n  btn.textContent = (_me && _me.classList.contains('light')) ? '${moon}' : '${sun}';\n}`
  );
  t = t.replace(/#mode-toggle/g, '#madds-mode-toggle');
  t = t.replace('#madds-mode-toggle {\n  position: fixed', '#madds-mode-toggle {\n  position: sticky');
  t = t.replace('</script>\n</body>\n</html>', `</script>\n</div>\n${FONT_LOADER}`);
  t = t.replace(/\s*<\/html>\s*$/, '');
  t = t.replace(/\s*<\/body>\s*$/, '');
  return t.trimEnd() + '\n';
}

function verify(output) {
  return [
    [!output.includes(':root {'),                   '":root {" not removed'],
    [!output.includes('body.light'),                '"body.light" not replaced'],
    [!output.includes('id="mode-toggle"'),          '"id=mode-toggle" not renamed'],
    [!output.includes('onclick="toggleMode'),       '"toggleMode" onclick not renamed'],
    [!output.includes('function toggleMode'),       '"toggleMode" fn not renamed'],
    [!output.includes('<body>'),                    '"<body>" not replaced'],
    [!output.includes('<!DOCTYPE'),                 '"<!DOCTYPE" not stripped'],
    [output.includes('<div class="madds-embed">'),  'madds-embed wrapper missing'],
    [output.includes('maddstoggleMode'),            'maddstoggleMode missing'],
    [output.includes('DM+Serif'),                   'font loader missing'],
    [output.includes('</div>'),                     'closing </div> missing'],
    [output.includes('position: sticky'),           'position:sticky not applied'],
  ].filter(([ok]) => !ok).map(([, msg]) => msg);
}

function run(cmd, cwd) {
  return execSync(cmd, { cwd, encoding: 'utf8', stdio: ['pipe','pipe','pipe'] }).trim();
}

let OUTPUT_CONTENT = '';

function processTarget(target) {
  console.log('\n[' + target.repo + '] Cloning...');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sync-' + target.repo + '-'));
  try {
    run(`git clone https://x-access-token:${TOKEN}@github.com/${target.owner}/${target.repo}.git .`, tmpDir);
    run('git config user.email "sync-org-chart@dds.local"', tmpDir);
    run('git config user.name "DDS Org Chart Sync"', tmpDir);

    for (const branch of target.branches) {
      try { run(`git checkout ${branch}`, tmpDir); }
      catch { run(`git checkout -b ${branch} origin/${branch}`, tmpDir); }

      let changed = false;
      for (const filePath of target.files) {
        const abs = path.join(tmpDir, filePath);
        if (!fs.existsSync(path.dirname(abs))) {
          console.log('[' + target.repo + '@' + branch + '] skip ' + filePath + ' (dir not found)');
          continue;
        }
        const existing = fs.existsSync(abs) ? fs.readFileSync(abs, 'utf8') : null;
        if (existing === OUTPUT_CONTENT) {
          console.log('[' + target.repo + '@' + branch + '] skip ' + filePath + ' (already current)');
          continue;
        }
        fs.writeFileSync(abs, OUTPUT_CONTENT, 'utf8');
        run('git add ' + filePath, tmpDir);
        changed = true;
        console.log('[' + target.repo + '@' + branch + '] wrote ' + filePath);
      }

      if (changed) {
        run('git commit -m "' + COMMIT_MSG + '"', tmpDir);
        run('git push origin ' + branch, tmpDir);
        console.log('[' + target.repo + '@' + branch + '] pushed');
      } else {
        console.log('[' + target.repo + '@' + branch + '] no changes — skipped');
      }
    }
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
  }
}

async function main() {
  console.log('\n[sync-org-chart] Fetching ' + SOURCE.file + ' from ' + SOURCE.owner + '/' + SOURCE.repo);
  let source, sha;
  try {
    ({ content: source, sha } = await fetchFile(SOURCE.owner, SOURCE.repo, SOURCE.file, SOURCE.branch, TOKEN));
  } catch(err) {
    console.error('[sync-org-chart] FETCH FAILED:', err.message);
    process.exit(1);
  }
  console.log('[sync-org-chart] ' + source.length.toLocaleString() + ' chars (sha: ' + sha.slice(0,7) + ')');

  OUTPUT_CONTENT = transform(source);
  const failures = verify(OUTPUT_CONTENT);
  if (failures.length) {
    console.error('[sync-org-chart] VERIFICATION FAILED:');
    failures.forEach(f => console.error('  - ' + f));
    process.exit(1);
  }
  console.log('[sync-org-chart] Verified OK (' + OUTPUT_CONTENT.split('\n').length + ' lines)');

  for (const target of TARGETS) {
    processTarget(target);
  }

  console.log('\n[sync-org-chart] Complete — all repos and branches updated.\n');
}

main().catch(err => {
  console.error('[sync-org-chart] Error:', err.message);
  process.exit(1);
});
