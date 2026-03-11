/**
 * Deploy Guards — Automated Regression Tests
 *
 * Catches the exact class of bugs that caused the v4.24.8/v4.24.9 outage:
 *   G1: HTML view file JS syntax errors (killed initStewardView → white screen)
 *   G2: Missing sessionToken params on server wrappers (broke magic-link auth)
 *   G3: OAuth scope gaps (MailApp needs script.send_mail, not gmail.send)
 *   G4: Client-server parameter mismatches (client sends args server doesn't expect)
 *   G5: Unescaped quotes in JS strings inside HTML (the Week's apostrophe)
 *   G6: dist/ parity with src/ (stale dist = deploying old code)
 *
 * Run: npm test (included in default test suite)
 * Run standalone: npx jest test/deploy-guards.test.js --verbose
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC_DIR = path.resolve(__dirname, '..', 'src');
const DIST_DIR = path.resolve(__dirname, '..', 'dist');

// HTML view files that contain <script> blocks
const VIEW_FILES = [
  'index.html',
  'auth_view.html',
  'steward_view.html',
  'member_view.html',
  'error_view.html',
];

// .gs files that contain google.script.run wrapper functions
const WRAPPER_FILES = [
  '21_WebDashDataService.gs',
  '26_QAForum.gs',
  '27_TimelineService.gs',
  '28_FailsafeService.gs',
  '25_WorkloadService.gs',
  '08e_SurveyEngine.gs',
];

// HTML view files that call google.script.run (client-side)
const CLIENT_VIEW_FILES = [
  'index.html',
  'steward_view.html',
  'member_view.html',
  'auth_view.html',
];

/**
 * Extract all <script> blocks from an HTML file, skipping blocks
 * that contain GAS template tags (<?= ?>) which aren't valid JS.
 */
function extractScriptBlocks(filePath) {
  const content = fs.readFileSync(filePath, 'utf8');
  const blocks = [];
  const regex = /<script[^>]*>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = regex.exec(content)) !== null) {
    const js = match[1];
    // Skip empty blocks or blocks with GAS template tags
    if (js.trim().length < 10) continue;
    if (js.includes('<?')) continue;
    blocks.push({ js, startLine: content.substring(0, match.index).split('\n').length });
  }
  return blocks;
}

// ============================================================================
// G1: HTML VIEW FILE SYNTAX VALIDATION
// ============================================================================
// Reason: An unescaped apostrophe in steward_view.html killed the entire
// <script> block, making initStewardView undefined → white screen.

describe('G1: HTML view files have valid JavaScript', () => {
  VIEW_FILES.forEach(file => {
    const filePath = path.join(SRC_DIR, file);
    if (!fs.existsSync(filePath)) return;

    test(`${file} — all <script> blocks parse without syntax errors`, () => {
      const blocks = extractScriptBlocks(filePath);
      blocks.forEach((block, i) => {
        expect(() => {
          new vm.Script(block.js, { filename: `${file}:block${i}` });
        }).not.toThrow();
      });
    });
  });
});


// ============================================================================
// G2: WRAPPER FUNCTIONS DECLARE sessionToken PARAMETER
// ============================================================================
// Reason: v4.23.1 auth sweep rewrote function bodies to use
// _resolveCallerEmail(sessionToken) but forgot to add sessionToken to 6
// function parameter lists. Server silently received undefined.

describe('G2: Server wrapper functions declare sessionToken when used', () => {

  WRAPPER_FILES.forEach(file => {
    const filePath = path.join(SRC_DIR, file);
    if (!fs.existsSync(filePath)) return;

    test(`${file} — no undeclared sessionToken references`, () => {
      const code = fs.readFileSync(filePath, 'utf8');
      const lines = code.split('\n');
      const errors = [];

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // Match global function declarations (not inside IIFEs — starts at column 0)
        const funcMatch = line.match(/^function\s+(\w+)\s*\(([^)]*)\)/);
        if (!funcMatch) continue;

        const funcName = funcMatch[1];
        const params = funcMatch[2].split(',').map(p => p.trim()).filter(Boolean);

        // Scan the function body (until next top-level function or EOF)
        let body = '';
        let braceCount = 0;
        let started = false;
        for (let j = i; j < lines.length; j++) {
          for (const ch of lines[j]) {
            if (ch === '{') { braceCount++; started = true; }
            if (ch === '}') braceCount--;
          }
          body += lines[j] + '\n';
          if (started && braceCount === 0) break;
        }

        // Check: body uses sessionToken but params don't include it
        const usesSessionToken = /\bsessionToken\b/.test(body) &&
          // Exclude comment-only references
          body.split('\n').some(l => {
            const trimmed = l.trim();
            return !trimmed.startsWith('//') && !trimmed.startsWith('*') && /\bsessionToken\b/.test(trimmed);
          });

        const declaresSessionToken = params.includes('sessionToken');

        // Stub functions that return constants and don't actually use sessionToken
        // are OK (e.g., dataGetActivePolls returns [])
        const isStub = body.includes('return []') || body.includes('return { success: false') ||
          (body.match(/return/g) || []).length === 1 && !body.includes('_resolveCallerEmail') && !body.includes('_requireStewardAuth');

        if (usesSessionToken && !declaresSessionToken && !isStub) {
          errors.push(`${funcName}() at line ${i + 1}: body references sessionToken but params are (${funcMatch[2]})`);
        }
      }

      expect(errors).toEqual([]);
    });
  });
});


// ============================================================================
// G3: OAUTH SCOPES MATCH API USAGE
// ============================================================================
// Reason: MailApp.sendEmail needs script.send_mail scope, not gmail.send.
// Both appsscript.json files must declare every scope used in code.

describe('G3: OAuth scopes cover all API usage', () => {
  // Map of GAS API calls to required OAuth scopes
  const SCOPE_MAP = {
    'MailApp.sendEmail':           'script.send_mail',
    'MailApp.getRemainingDailyQuota': 'script.send_mail',
    'GmailApp.sendEmail':          'gmail.send',
    'GmailApp.search':             'gmail.readonly',
    'SpreadsheetApp':              'spreadsheets',
    'DriveApp':                    'drive',
    'DocumentApp':                 'documents',
    'CalendarApp':                 'calendar',
    'ScriptApp.getService':        'script.scriptapp',
    'Session.getActiveUser':       'userinfo.email',
    'UrlFetchApp.fetch':           'script.external_request',
  };

  // Check both appsscript.json files
  ['appsscript.json', 'dist/appsscript.json'].forEach(jsonFile => {
    const jsonPath = path.resolve(__dirname, '..', jsonFile);
    if (!fs.existsSync(jsonPath)) return;

    test(`${jsonFile} — declares all required scopes`, () => {
      const manifest = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
      const scopes = (manifest.oauthScopes || []).join(' ');

      // Scan all .gs files for API usage
      const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
      const missing = [];

      for (const [apiCall, requiredScope] of Object.entries(SCOPE_MAP)) {
        // Check if any .gs file uses this API
        const apiUsed = gsFiles.some(f => {
          const code = fs.readFileSync(path.join(SRC_DIR, f), 'utf8');
          // Skip comments
          const lines = code.split('\n').filter(l => !l.trim().startsWith('//') && !l.trim().startsWith('*'));
          return lines.some(l => l.includes(apiCall));
        });

        if (apiUsed && !scopes.includes(requiredScope)) {
          missing.push(`${apiCall} requires scope "${requiredScope}"`);
        }
      }

      expect(missing).toEqual([]);
    });
  });

  test('both appsscript.json files have identical scopes', () => {
    const rootPath = path.resolve(__dirname, '..', 'appsscript.json');
    const distPath = path.resolve(__dirname, '..', 'dist', 'appsscript.json');
    if (!fs.existsSync(rootPath) || !fs.existsSync(distPath)) return;

    const rootScopes = JSON.parse(fs.readFileSync(rootPath, 'utf8')).oauthScopes || [];
    const distScopes = JSON.parse(fs.readFileSync(distPath, 'utf8')).oauthScopes || [];
    expect(rootScopes.sort()).toEqual(distScopes.sort());
  });
});


// ============================================================================
// G4: CLIENT-SERVER PARAMETER MATCHING
// ============================================================================
// Reason: Client called dataGetBatchData(email, view) but server expected
// dataGetBatchData(sessionToken). Args were silently shifted.

describe('G4: Client google.script.run calls match server function signatures', () => {

  test('client calls pass correct number of arguments to server functions', () => {
    // Build map of server function signatures from .gs files
    const serverFuncs = {};
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));

    for (const f of gsFiles) {
      const code = fs.readFileSync(path.join(SRC_DIR, f), 'utf8');
      // Match top-level function declarations
      const regex = /^function\s+(\w+)\s*\(([^)]*)\)/gm;
      let match;
      while ((match = regex.exec(code)) !== null) {
        const name = match[1];
        const params = match[2].split(',').map(p => p.trim()).filter(Boolean);
        serverFuncs[name] = { params, file: f };
      }
    }

    // Scan client HTML for google.script.run / serverCall() invocations
    const errors = [];

    for (const viewFile of CLIENT_VIEW_FILES) {
      const filePath = path.join(SRC_DIR, viewFile);
      if (!fs.existsSync(filePath)) continue;

      const code = fs.readFileSync(filePath, 'utf8');
      const lines = code.split('\n');

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        // Skip comments
        if (line.trim().startsWith('//') || line.trim().startsWith('*')) continue;

        // Match patterns: .functionName(args) at end of google.script.run chain
        // Only match lines that contain google.script.run or serverCall() chains
        // to avoid false positives on client-side functions (e.g., ThemeEngine.applyTheme)
        const isServerCallLine = line.includes('google.script.run') || line.includes('serverCall(') ||
          line.includes('.withSuccessHandler') || line.includes('.withFailureHandler') ||
          line.includes('DataCache.cachedCall');
        if (!isServerCallLine) continue;

        // Pattern: .funcName(arg1, arg2, ...)
        const callRegex = /\.(\w+)\(([^)]*)\)\s*;/g;
        let callMatch;
        while ((callMatch = callRegex.exec(line)) !== null) {
          const funcName = callMatch[1];
          const argStr = callMatch[2].trim();

          // Skip non-server calls (standard JS methods, handler setups)
          if (['withSuccessHandler', 'withFailureHandler', 'withUserObject',
            'getElementById', 'querySelector', 'addEventListener', 'appendChild',
            'removeChild', 'setAttribute', 'toString', 'toLowerCase', 'toUpperCase',
            'trim', 'replace', 'split', 'join', 'filter', 'map', 'forEach',
            'push', 'pop', 'slice', 'splice', 'sort', 'concat', 'indexOf',
            'includes', 'find', 'reduce', 'some', 'every', 'keys', 'values',
            'entries', 'stringify', 'parse', 'log', 'error', 'warn',
            'getTime', 'getFullYear', 'getMonth', 'getDate', 'toISOString',
            'clearInterval', 'setInterval', 'setTimeout', 'clearTimeout',
            'reload', 'focus', 'blur', 'click', 'remove', 'contains',
            'preventDefault', 'stopPropagation', 'getComputedStyle',
            'scrollIntoView', 'getBoundingClientRect', 'apply', 'call', 'bind',
            'resolve', 'reject', 'then', 'catch',
          ].includes(funcName)) continue;

          // Only check functions that exist as server-side declarations
          if (!serverFuncs[funcName]) continue;

          // Count client args (handle nested parens, strings, etc.)
          const clientArgCount = argStr === '' ? 0 : countTopLevelArgs(argStr);
          const serverParamCount = serverFuncs[funcName].params.length;

          // GAS allows extra args (they're ignored), but FEWER args
          // mean the server gets undefined where it expects data.
          // We flag if counts don't match at all.
          if (clientArgCount !== serverParamCount) {
            errors.push(
              `${viewFile}:${i + 1} calls ${funcName}(${clientArgCount} args) ` +
              `but server declares ${funcName}(${serverParamCount} params: ${serverFuncs[funcName].params.join(', ')}) ` +
              `in ${serverFuncs[funcName].file}`
            );
          }
        }
        // Also check DataCache.cachedCall('key', 'funcName', [args], cb) pattern
        const cacheCallRegex = /DataCache\.cachedCall\(\s*'[^']*'\s*,\s*'(\w+)'\s*,\s*\[([^\]]*)\]/g;
        let cacheMatch;
        while ((cacheMatch = cacheCallRegex.exec(line)) !== null) {
          const funcName = cacheMatch[1];
          const argStr = cacheMatch[2].trim();
          if (!serverFuncs[funcName]) continue;
          const clientArgCount = argStr === '' ? 0 : countTopLevelArgs(argStr);
          const serverParamCount = serverFuncs[funcName].params.length;
          if (clientArgCount !== serverParamCount) {
            errors.push(
              `${viewFile}:${i + 1} cachedCall to ${funcName}(${clientArgCount} args) ` +
              `but server declares ${funcName}(${serverParamCount} params: ${serverFuncs[funcName].params.join(', ')}) ` +
              `in ${serverFuncs[funcName].file}`
            );
          }
        }
      }
    }

    // Report all mismatches (not just first)
    if (errors.length > 0) {
      // eslint-disable-next-line no-console
      console.warn('Parameter mismatches found:\n  ' + errors.join('\n  '));
    }
    expect(errors).toEqual([]);
  });
});


// ============================================================================
// G5: NO UNESCAPED QUOTES IN JS STRINGS INSIDE HTML
// ============================================================================
// Reason: 'Create This Week's ' had a raw apostrophe inside single quotes
// that terminated the string early → fatal syntax error.

describe('G5: No unescaped apostrophes in single-quoted JS strings', () => {
  VIEW_FILES.forEach(file => {
    const filePath = path.join(SRC_DIR, file);
    if (!fs.existsSync(filePath)) return;

    test(`${file} — no raw contractions/possessives in single-quoted strings`, () => {
      const code = fs.readFileSync(filePath, 'utf8');
      const lines = code.split('\n');
      const errors = [];

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        // Skip pure comment lines
        if (line.trim().startsWith('//') || line.trim().startsWith('*')) continue;

        // Detect pattern: single-quoted string containing word's (apostrophe)
        // This regex finds: 'text word's text' where ' breaks the string
        // Look for: '<content><letter>'<letter><content>'
        // e.g., 'This Week's data'
        const dangerPattern = /'[^']*[a-zA-Z]'[a-zA-Z][^']*'/g;
        let match;
        while ((match = dangerPattern.exec(line)) !== null) {
          // Exclude CSS property values like 'var(--fontBody)'
          if (match[0].includes('var(--')) continue;
          // Exclude template expressions
          if (match[0].includes('${')) continue;
          // Exclude matches spanning across double-quoted strings (apostrophe is safely inside double quotes)
          if (match[0].includes('"')) continue;
          errors.push(`Line ${i + 1}: Possible unescaped apostrophe in single-quoted string: ${match[0].substring(0, 80)}`);
        }
      }

      expect(errors).toEqual([]);
    });
  });
});


// ============================================================================
// G6: DIST PARITY — dist/ must match src/ after build
// ============================================================================
// Reason: dist/ wasn't rebuilt after fixes → clasp push deployed old broken code.

describe('G6: dist/ files are in sync with src/', () => {

  test('every src .gs file has identical copy in dist', () => {
    const PROD_EXCLUDED = ['07_DevTools.gs']; // excluded by --prod build
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs') && !PROD_EXCLUDED.includes(f));
    const stale = [];

    for (const f of gsFiles) {
      const srcPath = path.join(SRC_DIR, f);
      const distPath = path.join(DIST_DIR, f);

      if (!fs.existsSync(distPath)) {
        stale.push(`${f}: missing from dist/`);
        continue;
      }

      const srcContent = fs.readFileSync(srcPath, 'utf8');
      const distContent = fs.readFileSync(distPath, 'utf8');
      if (srcContent !== distContent) {
        stale.push(`${f}: dist/ differs from src/ — run "node build.js"`);
      }
    }

    expect(stale).toEqual([]);
  });

  test('every src .html file has identical copy in dist', () => {
    const htmlFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
    const stale = [];

    for (const f of htmlFiles) {
      const srcPath = path.join(SRC_DIR, f);
      const distPath = path.join(DIST_DIR, f);

      if (!fs.existsSync(distPath)) {
        stale.push(`${f}: missing from dist/`);
        continue;
      }

      const srcContent = fs.readFileSync(srcPath, 'utf8');
      const distContent = fs.readFileSync(distPath, 'utf8');
      if (srcContent !== distContent) {
        stale.push(`${f}: dist/ differs from src/ — run "node build.js"`);
      }
    }

    expect(stale).toEqual([]);
  });
});


// ============================================================================
// G7: .gs FILES — NO SYNTAX ERRORS
// ============================================================================
// Reason: QAForum wrapper functions had )) double-closing-paren syntax errors
// that weren't caught because ESLint doesn't lint dist/ and the error was only
// in dist (src was fixed but dist wasn't rebuilt).

describe('G7: All .gs files have valid JavaScript syntax', () => {
  const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));

  gsFiles.forEach(file => {
    test(`src/${file} parses without syntax errors`, () => {
      const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
      expect(() => {
        new vm.Script(code, { filename: file });
      }).not.toThrow();
    });
  });
});


// ============================================================================
// G8: build.js file arrays match actual src/ contents
// ============================================================================
// Reason: poms_reference.html was added to src/ and dist/ but NOT to the
// HTML_FILES array in build.js. A fresh --prod build would delete it from dist/,
// breaking the POMS Reference tab at runtime.

describe('G8: build.js file arrays match src/ contents', () => {
  const buildCode = fs.readFileSync(path.resolve(__dirname, '..', 'build.js'), 'utf8');

  // Extract HTML_FILES array from build.js
  const htmlMatch = buildCode.match(/const HTML_FILES\s*=\s*\[([\s\S]*?)\];/);
  const registeredHtml = htmlMatch
    ? htmlMatch[1].match(/'([^']+\.html)'/g).map(m => m.replace(/'/g, ''))
    : [];

  // Extract BUILD_ORDER (GS files) array from build.js
  const gsMatch = buildCode.match(/const BUILD_ORDER\s*=\s*\[([\s\S]*?)\];/);
  const registeredGs = gsMatch
    ? gsMatch[1].match(/'([^']+\.gs)'/g).map(m => m.replace(/'/g, ''))
    : [];

  test('every .html file in src/ is registered in HTML_FILES', () => {
    const srcHtml = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
    const missing = srcHtml.filter(f => !registeredHtml.includes(f));
    expect(missing).toEqual([]);
  });

  test('every .gs file in src/ is registered in BUILD_ORDER', () => {
    const srcGs = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
    const missing = srcGs.filter(f => !registeredGs.includes(f));
    expect(missing).toEqual([]);
  });

  test('no phantom entries in HTML_FILES (file must exist in src/)', () => {
    const srcHtml = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.html'));
    const phantom = registeredHtml.filter(f => !srcHtml.includes(f));
    expect(phantom).toEqual([]);
  });

  test('no phantom entries in BUILD_ORDER (file must exist in src/)', () => {
    const srcGs = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
    const phantom = registeredGs.filter(f => !srcGs.includes(f));
    expect(phantom).toEqual([]);
  });
});


// ============================================================================
// G9: Nav tab entries have corresponding handlers
// ============================================================================
// Reason: Adding a tab to _getSidebarTabs without a case in _handleTabNav
// gives a white screen or falls through to default. Adding a handler without
// the nav entry makes the feature unreachable.

describe('G9: Nav tabs and handlers are in sync', () => {
  const indexCode = fs.readFileSync(path.join(SRC_DIR, 'index.html'), 'utf8');

  // Extract all tab ids from _getSidebarTabs
  const tabIdMatches = indexCode.match(/\{\s*id:\s*'([^']+)'/g) || [];
  const navTabIds = [...new Set(
    tabIdMatches.map(m => m.match(/id:\s*'([^']+)'/)[1])
      .filter(id => !id.startsWith('_')) // skip _member_more, _steward_more
  )];

  // Extract handled tab ids from _handleTabNav (case 'xxx': and tabId === 'xxx')
  const caseMatches = indexCode.match(/case\s+'([^']+)':/g) || [];
  const ifMatches = indexCode.match(/tabId\s*===\s*'([^']+)'/g) || [];
  const handledIds = [...new Set([
    ...caseMatches.map(m => m.match(/case\s+'([^']+)'/)[1]),
    ...ifMatches.map(m => m.match(/tabId\s*===\s*'([^']+)'/)[1]),
  ])];

  test('every nav tab id has a handler in _handleTabNav', () => {
    const unhandled = navTabIds.filter(id => !handledIds.includes(id));
    expect(unhandled).toEqual([]);
  });

  // Shared tabs that should appear in BOTH steward and member nav
  const sharedTabs = ['orgchart', 'poms'];
  sharedTabs.forEach(tabId => {
    test(`shared tab '${tabId}' appears in both steward and member nav arrays`, () => {
      // Count occurrences — should be 2 (once per role)
      const count = tabIdMatches.filter(m => m.includes(`'${tabId}'`)).length;
      expect(count).toBeGreaterThanOrEqual(2);
    });
  });
});


// ============================================================================
// G10: Lazy-loaded HTML files have server functions
// ============================================================================
// Reason: renderPOMSReference calls google.script.run.getPOMSReferenceHtml().
// If the server function doesn't exist, the tab shows "Failed to load".

describe('G10: Lazy-loaded views have matching server functions', () => {
  const indexCode = fs.readFileSync(path.join(SRC_DIR, 'index.html'), 'utf8');

  // Strip block comments
  const codeOnly = indexCode.replace(/\/\*[\s\S]*?\*\//g, '');

  // Find all HtmlService-pattern server functions: lines like
  //   .getOrgChartHtml();  or  .getPOMSReferenceHtml();
  // These are the terminal call in a google.script.run chain.
  // Pattern: line has only whitespace, a dot, a function name, parens, semicolon
  const terminalCalls = codeOnly.match(/^\s+\.(\w+)\(\s*\)\s*;/gm) || [];
  const serverFunctions = new Set();
  terminalCalls.forEach(line => {
    const name = line.match(/\.(\w+)\(/)[1];
    if (!name.startsWith('with') && !['forEach', 'map', 'filter', 'join',
        'appendChild', 'remove', 'push', 'pop', 'trim', 'replace',
        'toString', 'slice', 'splice'].includes(name)) {
      serverFunctions.add(name);
    }
  });

  // Collect all function declarations from .gs files
  const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
  const allGsFunctions = new Set();
  gsFiles.forEach(file => {
    const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
    const fnMatches = code.match(/^function\s+(\w+)\s*\(/gm) || [];
    fnMatches.forEach(m => allGsFunctions.add(m.match(/function\s+(\w+)/)[1]));
  });

  // Must find at least the known lazy-loaders
  test('detects lazy-loaded server functions (getOrgChartHtml, getPOMSReferenceHtml)', () => {
    expect(serverFunctions.has('getOrgChartHtml')).toBe(true);
    expect(serverFunctions.has('getPOMSReferenceHtml')).toBe(true);
  });

  [...serverFunctions].forEach(fn => {
    test(`server function ${fn}() exists in .gs files`, () => {
      expect(allGsFunctions.has(fn)).toBe(true);
    });
  });
});


// ============================================================================
// G11: poms_reference.html is self-contained and parseable
// ============================================================================
// Reason: poms_reference.html is loaded via HtmlService.createHtmlOutputFromFile
// and injected into the SPA. If its <script> block has syntax errors, the entire
// POMS tab fails silently.

describe('G11: poms_reference.html integrity', () => {
  const pomsPath = path.join(SRC_DIR, 'poms_reference.html');

  test('poms_reference.html exists', () => {
    expect(fs.existsSync(pomsPath)).toBe(true);
  });

  test('contains POMS_DATA array with entries', () => {
    const code = fs.readFileSync(pomsPath, 'utf8');
    const match = code.match(/POMS_DATA\s*=\s*\[/);
    expect(match).not.toBeNull();
    // Count entries
    const entries = (code.match(/\{id:/g) || []).length;
    expect(entries).toBeGreaterThanOrEqual(78);
  });

  test('contains FLOWS object with flowcharts', () => {
    const code = fs.readFileSync(pomsPath, 'utf8');
    const match = code.match(/FLOWS\s*=\s*\{/);
    expect(match).not.toBeNull();
    // Count flowcharts (title: entries)
    const charts = (code.match(/title:"/g) || []).length;
    expect(charts).toBeGreaterThanOrEqual(17);
  });

  test('all 4 tabs are handled (search, bookmarks, rated, stats)', () => {
    const code = fs.readFileSync(pomsPath, 'utf8');
    ['search', 'bookmarks', 'rated', 'stats'].forEach(tab => {
      expect(code).toContain(`P.tab==='${tab}'`);
    });
  });

  test('<script> block parses without syntax errors', () => {
    const code = fs.readFileSync(pomsPath, 'utf8');
    const scriptMatch = code.match(/<script>([\s\S]*)<\/script>/);
    expect(scriptMatch).not.toBeNull();
    expect(() => {
      new vm.Script(scriptMatch[1], { filename: 'poms_reference.html' });
    }).not.toThrow();
  });

  test('no DDS Apps Script ID present', () => {
    const code = fs.readFileSync(pomsPath, 'utf8');
    expect(code).not.toContain('18hHHX');
  });
});


// ============================================================================
// G12: No sensitive IDs leaked in source files
// ============================================================================
// Reason: DDS-Dashboard Apps Script ID (18hHHX...) must never appear in source
// files that get synced to the public Union-Tracker repo.

describe('G12: No sensitive ID leaks', () => {
  const DDS_SCRIPT_ID_PREFIX = '18hHHX';
  const srcFiles = fs.readdirSync(SRC_DIR);

  // AI_REFERENCE.md legitimately contains Script IDs in the private DDS repo;
  // in public Union-Tracker the ID is redacted so this guard still protects there.
  const G12_EXCLUDES = ['AI_REFERENCE.md'];

  srcFiles.filter(f => !G12_EXCLUDES.includes(f)).forEach(file => {
    test(`src/${file} does not contain DDS Script ID`, () => {
      const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
      expect(code).not.toContain(DDS_SCRIPT_ID_PREFIX);
    });
  });
});


// ============================================================================
// G13: MODULE LOAD ORDER SAFETY
// ============================================================================
// Reason: GAS V8 loads files alphabetically. IIFE modules (var X = (function(){})())
// must not reference other IIFE modules at init time (top-level execution outside
// function bodies) if those modules are defined in later-loading files.
// References inside function bodies are safe — by the time any function is called,
// all files have loaded.

describe('G13: Module load order is safe', () => {

  test('no duplicate IIFE module names across .gs files', () => {
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs')).sort();
    const iifeNames = {};
    const duplicates = [];

    gsFiles.forEach(file => {
      const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
      const iifeMatch = code.match(/^var\s+(\w+)\s*=\s*\(function/m);
      if (iifeMatch) {
        const name = iifeMatch[1];
        if (iifeNames[name]) {
          duplicates.push(`"${name}" defined in both ${iifeNames[name]} and ${file}`);
        }
        iifeNames[name] = file;
      }
    });

    expect(duplicates).toEqual([]);
  });

  test('no duplicate top-level function names across .gs files', () => {
    const gsFiles = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.gs')).sort();
    const funcMap = {};
    const duplicates = [];

    gsFiles.forEach(file => {
      const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
      const regex = /^function\s+(\w+)\s*\(/gm;
      let match;
      while ((match = regex.exec(code)) !== null) {
        const name = match[1];
        if (funcMap[name]) {
          duplicates.push(`"${name}" declared in both ${funcMap[name]} and ${file}`);
        }
        funcMap[name] = file;
      }
    });

    expect(duplicates).toEqual([]);
  });
});


// ============================================================================
// HELPERS
// ============================================================================

/**
 * Counts top-level comma-separated arguments in a string,
 * respecting nested parentheses, brackets, and string literals.
 * e.g., "a, fn(b, c), 'd,e'" → 3
 */
function countTopLevelArgs(str) {
  if (!str || str.trim() === '') return 0;

  let depth = 0;       // Paren/bracket nesting depth
  let inSingle = false; // Inside single-quoted string
  let inDouble = false; // Inside double-quoted string
  let inTemplate = false; // Inside template literal
  let count = 1;       // Start at 1 (first arg)
  let escaped = false;

  for (let i = 0; i < str.length; i++) {
    const ch = str[i];

    if (escaped) { escaped = false; continue; }
    if (ch === '\\') { escaped = true; continue; }

    if (inSingle) { if (ch === "'") inSingle = false; continue; }
    if (inDouble) { if (ch === '"') inDouble = false; continue; }
    if (inTemplate) { if (ch === '`') inTemplate = false; continue; }

    if (ch === "'") { inSingle = true; continue; }
    if (ch === '"') { inDouble = true; continue; }
    if (ch === '`') { inTemplate = true; continue; }

    if (ch === '(' || ch === '[' || ch === '{') { depth++; continue; }
    if (ch === ')' || ch === ']' || ch === '}') { depth--; continue; }

    if (ch === ',' && depth === 0) count++;
  }

  return count;
}
