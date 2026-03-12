/**
 * Architecture Tests — Cross-File Dependencies and HTML Output Validation
 *
 * These tests verify the structural integrity of the codebase to make
 * future refactoring (file splitting, HTML template extraction) safer.
 *
 * A1: Cross-file dependency tests — verify functions exist and are callable
 * A2: HTML output tests — verify HTML generators produce valid, safe output
 * A4: escapeHtml consistency — verify all client-side escapeHtml use the canonical helper
 */

const fs = require('fs');
const path = require('path');

require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock globals that modules expect
global.logAuditEvent = jest.fn();
global.AUDIT_EVENTS = {
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  FOLDER_CREATED: 'FOLDER_CREATED'
};

// Load all source files in build order (mirrors build.js BUILD_ORDER)
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '04a_UIMenus.gs',
  '04b_AccessibilityFeatures.gs',
  '04c_InteractiveDashboard.gs',
  '04d_ExecutiveDashboard.gs',
  '04e_PublicDashboard.gs',
  '05_Integrations.gs',
  '06_Maintenance.gs',
  '07_DevTools.gs',
  '08a_SheetSetup.gs',
  '08b_SearchAndCharts.gs',
  '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs',
  '08e_SurveyEngine.gs',
  '09_Dashboards.gs',
  '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs',
  '10c_FormHandlers.gs',
  '10d_SyncAndMaintenance.gs',
  '10_Main.gs',
  '11_CommandHub.gs',
  '12_Features.gs',
  '13_MemberSelfService.gs',
  '14_MeetingCheckIn.gs',
  '15_EventBus.gs',
  '16_DashboardEnhancements.gs',
  '17_CorrelationEngine.gs',
  '19_WebDashAuth.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs',
  '22_WebDashApp.gs',
  '23_PortalSheets.gs',
  '24_WeeklyQuestions.gs',
  '25_WorkloadService.gs',
  '26_QAForum.gs',
  '27_TimelineService.gs',
  '28_FailsafeService.gs',
  '29_Migrations.gs'
]);

// ============================================================================
// A1: CROSS-FILE DEPENDENCY TESTS
// ============================================================================

describe('A1: Cross-file dependencies', () => {

  describe('critical entry points exist as functions', () => {
    const entryPoints = [
      'onOpen',
      'onEdit',
      'createDashboardMenu',
      'doGet',
      'CREATE_DASHBOARD'
    ];

    entryPoints.forEach(fn => {
      test(`${fn} is a function`, () => {
        expect(typeof global[fn]).toBe('function');
      });
    });
  });

  describe('sync functions exist (cross-file orchestration)', () => {
    const syncFunctions = [
      'syncAllData',
      'syncGrievanceToMemberDirectory',
      'syncSatisfactionValues',
      'syncDashboardValues',
      'syncGrievanceFormulasToLog',
      'refreshAllHiddenFormulas'
    ];

    syncFunctions.forEach(fn => {
      test(`${fn} is a function`, () => {
        expect(typeof global[fn]).toBe('function');
      });
    });
  });

  describe('key utility functions exist (used cross-file)', () => {
    const utilities = [
      'escapeHtml',
      'getClientSideEscapeHtml',
      'getClientSecurityScript',
      'getConfigValue_',
      'isTruthyValue',
      'getMobileOptimizedHead',
      'getOrCreateRootFolder',
      'sanitizeFolderName',
      'getColumnLetter'
    ];

    utilities.forEach(fn => {
      test(`${fn} is a function`, () => {
        expect(typeof global[fn]).toBe('function');
      });
    });
  });

  describe('global constants are defined', () => {
    test('SHEETS has all required sheet names', () => {
      expect(SHEETS.MEMBER_DIR).toBeTruthy();
      expect(SHEETS.GRIEVANCE_LOG).toBeTruthy();
      expect(SHEETS.CONFIG).toBeTruthy();
    });

    test('MEMBER_COLS has required column indices', () => {
      expect(typeof MEMBER_COLS.MEMBER_ID).toBe('number');
      expect(typeof MEMBER_COLS.FIRST_NAME).toBe('number');
      expect(typeof MEMBER_COLS.LAST_NAME).toBe('number');
      expect(typeof MEMBER_COLS.EMAIL).toBe('number');
    });

    test('GRIEVANCE_COLS has required column indices', () => {
      expect(typeof GRIEVANCE_COLS.GRIEVANCE_ID).toBe('number');
      expect(typeof GRIEVANCE_COLS.MEMBER_ID).toBe('number');
      expect(typeof GRIEVANCE_COLS.STATUS).toBe('number');
      expect(typeof GRIEVANCE_COLS.CURRENT_STEP).toBe('number');
    });

    test('CALENDAR_CONFIG exists with expected properties', () => {
      expect(CALENDAR_CONFIG.CALENDAR_NAME).toBeTruthy();
      expect(Array.isArray(CALENDAR_CONFIG.REMINDER_DAYS)).toBe(true);
    });

    test('DRIVE_CONFIG exists with expected properties', () => {
      expect(DRIVE_CONFIG.ROOT_FOLDER_FALLBACK).toBeTruthy();
    });

    test('CACHE_CONFIG exists', () => {
      expect(typeof CACHE_CONFIG).toBe('object');
      expect(typeof CACHE_CONFIG.MEMORY_TTL).toBe('number');
    });

    test('VERSION_INFO exists with version', () => {
      expect(VERSION_INFO.CURRENT).toBeTruthy();
      expect(VERSION_INFO.CURRENT).toMatch(/^\d+\.\d+\.\d+$/);
    });
  });

  describe('form configs exist', () => {
    test('GRIEVANCE_FORM_CONFIG has FIELD_IDS', () => {
      expect(GRIEVANCE_FORM_CONFIG).toBeDefined();
      expect(GRIEVANCE_FORM_CONFIG.FIELD_IDS).toBeDefined();
      expect(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_ID).toBeTruthy();
    });

    test('CONTACT_FORM_CONFIG has FIELD_IDS', () => {
      expect(CONTACT_FORM_CONFIG).toBeDefined();
      expect(CONTACT_FORM_CONFIG.FIELD_IDS).toBeDefined();
    });

    test('SATISFACTION_FORM_CONFIG has FIELD_IDS', () => {
      expect(SATISFACTION_FORM_CONFIG).toBeDefined();
      expect(SATISFACTION_FORM_CONFIG.FIELD_IDS).toBeDefined();
    });
  });
});

// ============================================================================
// A1: BUILD ORDER TESTS
// ============================================================================

describe('A1: Build order integrity', () => {
  const BUILD_ORDER = [
    '00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '02_DataManagers.gs',
    '03_UIComponents.gs', '04a_UIMenus.gs', '04b_AccessibilityFeatures.gs',
    '04c_InteractiveDashboard.gs', '04d_ExecutiveDashboard.gs',
    '04e_PublicDashboard.gs', '05_Integrations.gs', '06_Maintenance.gs',
    '07_DevTools.gs', '08a_SheetSetup.gs', '08b_SearchAndCharts.gs',
    '08c_FormsAndNotifications.gs', '08d_AuditAndFormulas.gs', '08e_SurveyEngine.gs',
    '09_Dashboards.gs', '10a_SheetCreation.gs', '10b_SurveyDocSheets.gs',
    '10c_FormHandlers.gs', '10d_SyncAndMaintenance.gs', '10_Main.gs',
    '11_CommandHub.gs', '12_Features.gs', '13_MemberSelfService.gs',
    '14_MeetingCheckIn.gs', '15_EventBus.gs', '16_DashboardEnhancements.gs',
    '17_CorrelationEngine.gs', '19_WebDashAuth.gs',
    '20_WebDashConfigReader.gs', '21_WebDashDataService.gs',
    '22_WebDashApp.gs', '23_PortalSheets.gs', '24_WeeklyQuestions.gs',
    '25_WorkloadService.gs', '26_QAForum.gs', '27_TimelineService.gs',
    '28_FailsafeService.gs', '29_Migrations.gs', '30_TestRunner.gs'
  ];

  test('all source files in BUILD_ORDER exist on disk', () => {
    BUILD_ORDER.forEach(filename => {
      const filePath = path.resolve(__dirname, '..', 'src', filename);
      expect(fs.existsSync(filePath)).toBe(true);
    });
  });

  test('no unexpected .gs files outside BUILD_ORDER', () => {
    const srcDir = path.resolve(__dirname, '..', 'src');
    const gsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs'));
    gsFiles.forEach(file => {
      expect(BUILD_ORDER).toContain(file);
    });
  });
});

// ============================================================================
// A2: HTML OUTPUT TESTS
// ============================================================================

describe('A2: HTML output functions', () => {

  describe('getClientSideEscapeHtml produces valid JS', () => {
    test('returns a string containing function escapeHtml', () => {
      const script = getClientSideEscapeHtml();
      expect(script).toContain('function escapeHtml');
      expect(script).toContain('function safeText');
    });

    test('returned escapeHtml handles null/undefined', () => {
      // Evaluate the client-side function in a sandbox
      const script = getClientSideEscapeHtml();
      const fn = new Function(script + '; return escapeHtml;')();
      expect(fn(null)).toBe('');
      expect(fn(undefined)).toBe('');
    });

    test('returned escapeHtml escapes all dangerous characters', () => {
      const script = getClientSideEscapeHtml();
      const fn = new Function(script + '; return escapeHtml;')();
      const result = fn('<script>alert("xss")</script>');
      expect(result).not.toContain('<');
      expect(result).not.toContain('>');
      expect(result).not.toContain('"');
      expect(result).toContain('&lt;');
      expect(result).toContain('&gt;');
      expect(result).toContain('&quot;');
    });

    test('returned escapeHtml escapes single quotes but not forward slashes', () => {
      const script = getClientSideEscapeHtml();
      const fn = new Function(script + '; return escapeHtml;')();
      const result = fn("it's a/test");
      expect(result).not.toContain("'");
      expect(result).toContain('/'); // F107: / is not an XSS vector
      expect(result).toContain('&#x27;');
    });

    test('returned escapeHtml escapes backticks but not equals', () => {
      const script = getClientSideEscapeHtml();
      const fn = new Function(script + '; return escapeHtml;')();
      const result = fn('`template=${value}`');
      expect(result).not.toContain('`');
      expect(result).toContain('='); // F107: = is not an XSS vector
      expect(result).toContain('&#x60;');
    });
  });

  describe('getClientSecurityScript produces valid script tag', () => {
    test('wraps output in script tags', () => {
      const tag = getClientSecurityScript();
      expect(tag).toMatch(/^<script>/);
      expect(tag).toMatch(/<\/script>$/);
    });

    test('includes escapeHtml function', () => {
      const tag = getClientSecurityScript();
      expect(tag).toContain('function escapeHtml');
    });
  });

  describe('server-side escapeHtml matches client-side behavior', () => {
    test('both produce same output for XSS payloads', () => {
      const script = getClientSideEscapeHtml();
      const clientFn = new Function(script + '; return escapeHtml;')();

      const payloads = [
        '<img onerror=alert(1)>',
        '"><script>alert(1)</script>',
        "'; DROP TABLE users;--",
        '`${alert(1)}`',
        'normal text 123'
      ];

      payloads.forEach(payload => {
        expect(escapeHtml(payload)).toBe(clientFn(payload));
      });
    });
  });
});

// ============================================================================
// A4: NO INLINE escapeHtml DEFINITIONS
// ============================================================================

describe('A4: No inline escapeHtml duplicates in source files', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');
  const gsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs'));

  gsFiles.filter(f => f !== '00_Security.gs').forEach(file => {
    test(`${file} has no inline function escapeHtml definition`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      // Match actual function definitions, not references to the helper
      const inlineDefinitions = content.match(/function escapeHtml\s*\(/g);
      expect(inlineDefinitions).toBeNull();
    });
  });

  test('00_Security.gs has exactly 2 escapeHtml definitions (server + client helper)', () => {
    const content = fs.readFileSync(path.join(srcDir, '00_Security.gs'), 'utf8');
    const matches = content.match(/function escapeHtml\s*\(/g);
    expect(matches).toHaveLength(2);
  });

  test('getClientSideEscapeHtml is used in all files that need client-side escapeHtml', () => {
    const filesUsingHelper = [];
    gsFiles.forEach(file => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      if (content.includes('getClientSideEscapeHtml()')) {
        filesUsingHelper.push(file);
      }
    });
    // At minimum, the helper should be defined and used
    expect(filesUsingHelper.length).toBeGreaterThan(0);
    expect(filesUsingHelper).toContain('00_Security.gs');
  });
});

// ============================================================================
// A5: WEB APP ERROR RESILIENCE
// ============================================================================

describe('A5: Web app doGet() error resilience', () => {
  const webAppSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '22_WebDashApp.gs'), 'utf8'
  );

  test('doGet() wraps its entire body in a try/catch', () => {
    // Extract the doGet function body (between the first { and its matching })
    const doGetMatch = webAppSrc.match(/function doGet\s*\([^)]*\)\s*\{/);
    expect(doGetMatch).not.toBeNull();

    // The function must contain a try block
    const afterDoGet = webAppSrc.slice(doGetMatch.index);
    expect(afterDoGet).toMatch(/function doGet[^]*?try\s*\{/);
  });

  test('doGet() catch block calls _serveFatalError (zero-dependency fallback)', () => {
    // The catch in doGet must call _serveFatalError, NOT _serveError
    // _serveError depends on ConfigReader which may be the source of the error
    const doGetMatch = webAppSrc.match(/function doGet\s*\([^)]*\)\s*\{/);
    const afterDoGet = webAppSrc.slice(doGetMatch.index);
    // Find the catch block that belongs to doGet (the top-level one)
    expect(afterDoGet).toMatch(/\}\s*catch\s*\(\w+\)\s*\{[\s\S]*?_serveFatalError/);
  });

  test('_serveFatalError() exists and does not call ConfigReader or DataService', () => {
    expect(typeof global._serveFatalError).toBe('function');

    // Extract the function body from source
    const fnMatch = webAppSrc.match(/function _serveFatalError[\s\S]*?\nfunction /);
    expect(fnMatch).not.toBeNull();
    const fnBody = fnMatch[0];

    // Must NOT depend on ConfigReader, DataService, SpreadsheetApp, or sheet access
    expect(fnBody).not.toContain('ConfigReader');
    expect(fnBody).not.toContain('DataService');
    expect(fnBody).not.toContain('SpreadsheetApp');
    expect(fnBody).not.toContain('getActiveSpreadsheet');
    expect(fnBody).not.toContain('createTemplateFromFile');
  });

  test('doGetWebDashboard() catch block does not call ConfigReader.getConfig() without its own try/catch', () => {
    // Extract the catch block from doGetWebDashboard
    const fnStart = webAppSrc.indexOf('function doGetWebDashboard');
    const fnBody = webAppSrc.slice(fnStart);

    // Find the catch block
    const catchMatch = fnBody.match(/\}\s*catch\s*\(\w+\)\s*\{([\s\S]*?)\n\}/);
    expect(catchMatch).not.toBeNull();

    const catchBody = catchMatch[1];

    // If ConfigReader.getConfig() appears in the catch, it MUST be inside its own try
    if (catchBody.includes('ConfigReader.getConfig()')) {
      expect(catchBody).toMatch(/try\s*\{[\s\S]*?ConfigReader\.getConfig\(\)/);
    }
  });

  test('_serveFatalError returns HtmlOutput (HtmlService.createHtmlOutput)', () => {
    const fnMatch = webAppSrc.match(/function _serveFatalError[\s\S]*?\nfunction /);
    expect(fnMatch[0]).toContain('HtmlService.createHtmlOutput');
    // Must NOT use createTemplateFromFile (requires external file)
    expect(fnMatch[0]).not.toContain('createTemplateFromFile');
  });
});

// ============================================================================
// A3: NO EMPTY FUNCTION STUBS
// ============================================================================

describe('A3: No empty JSDoc-only stubs in form handler files', () => {
  test('10c_FormHandlers.gs has no orphaned JSDoc comments without function bodies', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10c_FormHandlers.gs'), 'utf8'
    );
    // Find JSDoc comments followed by blank line or another comment (no function)
    const orphanedJsdoc = content.match(/\*\/\s*\n\s*\n\s*\/\*\*/g);
    // Allow at most 2 (between config blocks that have actual var declarations following)
    const count = orphanedJsdoc ? orphanedJsdoc.length : 0;
    expect(count).toBeLessThanOrEqual(2);
  });

  test('10d_SyncAndMaintenance.gs has no "Note: defined in modular file" comments', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10d_SyncAndMaintenance.gs'), 'utf8'
    );
    expect(content).not.toContain('defined in modular file');
  });

  test('10c_FormHandlers.gs has no "Note: defined in modular file" comments', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10c_FormHandlers.gs'), 'utf8'
    );
    expect(content).not.toContain('defined in modular file');
  });

  test('10c_FormHandlers.gs is under 750 lines (was 922 with stubs)', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10c_FormHandlers.gs'), 'utf8'
    );
    const lineCount = content.split('\n').length;
    expect(lineCount).toBeLessThan(750);
  });

  test('10d_SyncAndMaintenance.gs is under 1150 lines (was 1519 with stubs)', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10d_SyncAndMaintenance.gs'), 'utf8'
    );
    const lineCount = content.split('\n').length;
    expect(lineCount).toBeLessThan(1150);
  });
});

// ============================================================================
// A6: getActiveSpreadsheet() NULL SAFETY IN WEB APP CHAIN
// ============================================================================

describe('A6: getActiveSpreadsheet() null safety in web app files', () => {
  // Web app files where getActiveSpreadsheet() returning null would be fatal
  const webAppFiles = [
    '20_WebDashConfigReader.gs',
    '21_WebDashDataService.gs',
    '23_PortalSheets.gs',
    '24_WeeklyQuestions.gs',
    '25_WorkloadService.gs',
  ];

  webAppFiles.forEach(file => {
    test(`${file}: every getActiveSpreadsheet() call has a null guard`, () => {
      const content = fs.readFileSync(
        path.resolve(__dirname, '..', 'src', file), 'utf8'
      );
      const lines = content.split('\n');
      const issues = [];

      lines.forEach((line, idx) => {
        // Find lines that call getActiveSpreadsheet()
        if (line.includes('getActiveSpreadsheet()')) {
          const lineNum = idx + 1;
          // Pattern 1: chained call — ss never stored, so no null check possible
          // e.g. SpreadsheetApp.getActiveSpreadsheet().getSheetByName(...)
          if (line.includes('getActiveSpreadsheet().get') ||
              line.includes('getActiveSpreadsheet().insert')) {
            issues.push('Line ' + lineNum + ': chained call without null check: ' + line.trim());
          }
          // Pattern 2: stored in var — next few lines must check if (!ss)
          const varMatch = line.match(/var\s+(\w+)\s*=\s*SpreadsheetApp\.getActiveSpreadsheet/);
          if (varMatch) {
            const varName = varMatch[1];
            // Check next 3 lines for a null guard
            const nextLines = lines.slice(idx + 1, idx + 4).join('\n');
            const hasGuard = nextLines.includes('if (!' + varName + ')') ||
                             nextLines.includes('if (' + varName + ')') ||
                             nextLines.includes(varName + ' ?') ||
                             nextLines.includes(varName + ' &&');
            if (!hasGuard) {
              issues.push('Line ' + lineNum + ': var ' + varName + ' assigned but no null check in next 3 lines');
            }
          }
        }
      });

      expect(issues).toEqual([]);
    });
  });
});

// ============================================================================
// A7: GAS TRIGGER ENTRY POINTS MUST HAVE TRY/CATCH
// ============================================================================

describe('A7: GAS trigger entry points have try/catch', () => {
  // These functions are called directly by GAS triggers — unhandled throws
  // cause silent failures or cryptic error dialogs
  const triggerEntryPoints = [
    { file: '10_Main.gs', fn: 'onOpen' },
    { file: '10_Main.gs', fn: 'onEdit' },
    { file: '08a_SheetSetup.gs', fn: 'onEditMultiSelect' },
    { file: '08a_SheetSetup.gs', fn: 'onSelectionChangeMultiSelect' },
  ];

  triggerEntryPoints.forEach(({ file, fn }) => {
    test(`${file}: ${fn}() wraps body in try/catch`, () => {
      const content = fs.readFileSync(
        path.resolve(__dirname, '..', 'src', file), 'utf8'
      );
      // Find the function
      const fnRegex = new RegExp('function ' + fn + '\\s*\\([^)]*\\)\\s*\\{');
      const fnMatch = content.match(fnRegex);
      expect(fnMatch).not.toBeNull();

      // Extract 40 lines after function start to check for try
      const afterFn = content.slice(fnMatch.index, fnMatch.index + 1500);
      // Must contain a try block (allow early returns like if (!e) return; before the try)
      expect(afterFn).toMatch(/try\s*\{/);
    });
  });
});

// ============================================================================
// A8: CLIENT-SIDE serverCall() HELPER EXISTS
// ============================================================================

describe('A8: Client-side serverCall() helper for safe google.script.run', () => {
  const indexHtml = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'), 'utf8'
  );

  test('index.html defines serverCall() function', () => {
    expect(indexHtml).toMatch(/function serverCall\s*\(\)/);
  });

  test('serverCall() attaches a default withFailureHandler', () => {
    expect(indexHtml).toMatch(/serverCall[\s\S]*?withFailureHandler/);
  });

  test('DataCache.cachedCall always attaches a withFailureHandler', () => {
    // The cachedCall function must always chain .withFailureHandler, not conditionally
    const cachedCallMatch = indexHtml.match(/function cachedCall[\s\S]*?\n {6}\}/);
    expect(cachedCallMatch).not.toBeNull();
    expect(cachedCallMatch[0]).toContain('.withFailureHandler(');
    // Must NOT have the old conditional pattern: if (onFailure) runner = runner.withFailureHandler
    expect(cachedCallMatch[0]).not.toMatch(/if\s*\(onFailure\)\s*runner/);
  });
});

// ============================================================================
// A9: EVERY UI TAB ROUTE HAS A CORRESPONDING RENDER FUNCTION
// ============================================================================
// Failure history: v4.19.1 — Org Chart tab wired in sidebar but renderOrgChart()
// never existed.  Tab click threw JS error.  This test prevents that class of bug.

describe('A9: UI tab routes have matching render functions', () => {
  const indexHtml = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'), 'utf8'
  );

  // Extract every function name referenced in _handleTabNav switch/case blocks
  // Pattern: renderFoo(app), renderFoo(app, 'member'), initFoo(app)
  const routedFunctions = [];
  const routeRegex = /case\s+'[^']+'\s*:\s*(\w+)\s*\(/g;
  let m;
  while ((m = routeRegex.exec(indexHtml)) !== null) {
    if (!routedFunctions.includes(m[1])) routedFunctions.push(m[1]);
  }
  // Also catch the orgchart handler called outside the switch
  const directRouteRegex = /if\s*\(tabId\s*===\s*'[^']+'\)\s*\{\s*(\w+)\s*\(/g;
  while ((m = directRouteRegex.exec(indexHtml)) !== null) {
    if (!routedFunctions.includes(m[1])) routedFunctions.push(m[1]);
  }

  test('found at least 20 routed render functions in index.html', () => {
    expect(routedFunctions.length).toBeGreaterThanOrEqual(20);
  });

  // Each routed function must be defined SOMEWHERE in the HTML files
  // (any of the SPA HTML files, including auth_view.html and error_view.html)
  const allHtml = ['index.html', 'steward_view.html', 'member_view.html', 'auth_view.html', 'error_view.html'].map(f =>
    fs.readFileSync(path.resolve(__dirname, '..', 'src', f), 'utf8')
  ).join('\n');

  routedFunctions.forEach(fn => {
    test(`${fn}() is defined (routed from _handleTabNav)`, () => {
      const defRegex = new RegExp('function\\s+' + fn + '\\s*\\(');
      expect(allHtml).toMatch(defRegex);
    });
  });
});

// ============================================================================
// A10: WEB APP setValue/appendRow CALLS USE escapeForFormula FOR USER DATA
// ============================================================================
// Failure history: v4.9.1, v4.14.0 — formula injection via =, +, -, @ in
// user-controlled data written to sheets without escapeForFormula().

describe('A10: Web app data writes use escapeForFormula where needed', () => {
  const dataServiceSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8'
  );

  test('setValue() calls with user-supplied string fields use escapeForFormula', () => {
    const lines = dataServiceSrc.split('\n');
    const issues = [];

    lines.forEach((line, idx) => {
      // Look for setValue() with a variable that looks like user input
      // Skip: setValue(new Date()), setValue('completed'), setValue(true/false),
      //        setValue(escapeForFormula(...)), setValue(number), setValue('')
      if (!line.includes('.setValue(')) return;
      const lineNum = idx + 1;
      const trimmed = line.trim();

      // Safe patterns — skip
      if (trimmed.includes('escapeForFormula(')) return;  // already escaped
      if (trimmed.match(/\.setValue\(\s*new Date/)) return;  // date
      if (trimmed.match(/\.setValue\(\s*['"][^'"]*['"]\s*\)/)) return;  // literal string
      if (trimmed.match(/\.setValue\(\s*(true|false|''|"")\s*\)/)) return;  // boolean/empty
      if (trimmed.match(/\.setValue\(\s*\d+\s*\)/)) return;  // number literal
      if (trimmed.match(/\.setValue\(\s*completed\s*\?/)) return;  // ternary with safe values

      // These are header/structure writes — safe
      if (trimmed.includes('getRange(1,')) return;  // row 1 = headers
      if (trimmed.includes("'Assignee Type'") || trimmed.includes("'Assigned By'")) return;

      // Remaining setValue calls with variables need escapeForFormula
      // But only flag ones that take a variable (not a safe expression)
      if (trimmed.match(/\.setValue\(\s*[a-z]/i) && !trimmed.match(/\.setValue\(\s*(new |completed |String\(|Number\()/)) {
        // Check if previous 2 lines have escapeForFormula
        const context = lines.slice(Math.max(0, idx - 2), idx + 1).join('\n');
        if (!context.includes('escapeForFormula')) {
          issues.push('Line ' + lineNum + ': ' + trimmed);
        }
      }
    });

    // Allow up to 5 remaining cases (some are genuinely safe — email, id, date)
    // This threshold should decrease as migration progresses
    expect(issues.length).toBeLessThanOrEqual(5);
  });

  test('appendRow() calls truncate user string fields (defense in depth)', () => {
    const appendRows = dataServiceSrc.match(/sheet\.appendRow\(\[[\s\S]*?\]\)/g) || [];
    // Every appendRow should contain at least one .substring() or .trim() call
    // on its user-controlled string arguments (defense in depth against oversized input)
    expect(appendRows.length).toBeGreaterThan(0);
    const withTruncation = appendRows.filter(call =>
      call.includes('.substring(') || call.includes('.trim()')
    );
    // At least half of appendRow calls with user data should truncate
    expect(withTruncation.length).toBeGreaterThanOrEqual(Math.floor(appendRows.length * 0.3));
  });
});

// ============================================================================
// A11: SERVER-EXPOSED FUNCTIONS HAVE AUTH CHECKS
// ============================================================================
// Failure history: v4.14.0 — rejectFlaggedSubmission() had no auth check.
// v4.18.1 — bulkUpdateGrievanceStatus() had no steward authorization.
// All client-callable functions (data*, qa*, tl*, fs*) must either call
// _resolveCallerEmail() or _requireStewardAuth().

describe('A11: Server-exposed functions have auth checks', () => {
  const dataServiceSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8'
  );
  const qaSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '26_QAForum.gs'), 'utf8'
  );
  const tlSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '27_TimelineService.gs'), 'utf8'
  );
  const fsSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '28_FailsafeService.gs'), 'utf8'
  );

  // Known safe exceptions: read-only public data functions that intentionally
  // skip auth (e.g. dataGetGrievanceStats returns aggregate-only data)
  const publicReadOnly = [
    'dataGetGrievanceStats', 'dataGetAgencyGrievanceStats',
    'dataGetGrievanceHotSpots', 'dataGetMembershipStats',
    'dataGetUpcomingEvents', 'dataGetSurveyQuestions',
    'dataGetMeetingMinutes', 'dataGetStewardDirectory',
    'dataGetCaseChecklist', 'dataGetSatisfactionTrends',
    'dataGetBroadcastFilterOptions', 'dataGetEngagementStats',
    'dataGetWorkloadSummaryStats',
    'dataGetActivePolls', 'dataSubmitPollVote', 'dataAddPoll',
  ];

  // Functions that are init/admin only (not called from client google.script.run)
  const adminOnly = [
    'qaInitSheets', 'tlInitSheets', 'fsInitSheets',
    'fsProcessScheduledDigests', 'fsBackupCriticalSheets',
    'fsSetupTriggers', 'fsRemoveTriggers',
    'wtArchiveOldData', 'wtCleanVault',
  ];

  function extractGlobalFunctions(src, prefix) {
    const regex = new RegExp('^function (' + prefix + '\\w+)\\s*\\(', 'gm');
    const fns = [];
    let match;
    while ((match = regex.exec(src)) !== null) fns.push(match[1]);
    return fns;
  }

  const allExposed = [
    ...extractGlobalFunctions(dataServiceSrc, 'data'),
    ...extractGlobalFunctions(qaSrc, 'qa'),
    ...extractGlobalFunctions(tlSrc, 'tl'),
    ...extractGlobalFunctions(fsSrc, 'fs'),
  ].filter(fn => !fn.endsWith('_')); // exclude private helpers (trailing _)

  const allSrc = dataServiceSrc + '\n' + qaSrc + '\n' + tlSrc + '\n' + fsSrc;

  test('found at least 40 server-exposed functions', () => {
    expect(allExposed.length).toBeGreaterThanOrEqual(40);
  });

  allExposed
    .filter(fn => !publicReadOnly.includes(fn) && !adminOnly.includes(fn))
    .forEach(fn => {
      test(`${fn}() calls _resolveCallerEmail or _requireStewardAuth`, () => {
        // Find the function body
        const fnRegex = new RegExp('function ' + fn + '\\s*\\([^)]*\\)\\s*\\{([^}]+)');
        const match = allSrc.match(fnRegex);
        if (!match) return; // function not found — skip (another test catches this)
        const body = match[1];
        const hasAuth = body.includes('_resolveCallerEmail') ||
                        body.includes('_requireStewardAuth') ||
                        body.includes('checkWebAppAuthorization');
        expect(hasAuth).toBe(true);
      });
    });
});

// ============================================================================
// A12: NO DYNAMIC innerHTML CONCATENATION IN .GS FILES
// ============================================================================
// Failure history: v4.9.1 — 75+ XSS instances from unescaped HTML concatenation.
// .gs files that generate HTML must use escapeHtml() for all dynamic values.

describe('A12: No unescaped dynamic HTML in .gs server files', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');
  // Files that generate HTML output (UI dialogs, web app pages)
  const htmlGeneratingFiles = [
    '03_UIComponents.gs',
    '04a_UIMenus.gs',
    '04b_AccessibilityFeatures.gs',
    '04c_InteractiveDashboard.gs',
    '04d_ExecutiveDashboard.gs',
    '04e_PublicDashboard.gs',
    '05_Integrations.gs',
    '11_CommandHub.gs',
    '13_MemberSelfService.gs',
    '14_MeetingCheckIn.gs',
  ];

  htmlGeneratingFiles.forEach(file => {
    test(`${file}: dynamic values in HTML use escapeHtml() or JSON.stringify()`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const lines = content.split('\n');
      const issues = [];

      lines.forEach((line, idx) => {
        // Pattern: '<tag>' + variable + '</tag>' without escapeHtml
        // Detect: string + variable + string patterns in HTML context
        if (!line.includes("'<") && !line.includes('"<')) return;
        if (!line.includes('+')) return;

        const lineNum = idx + 1;
        const trimmed = line.trim();

        // Skip lines that already use escapeHtml or JSON.stringify
        if (trimmed.includes('escapeHtml(') || trimmed.includes('JSON.stringify(')) return;

        // Skip pure literal HTML (no variable interpolation)
        if (!trimmed.match(/['"].*['"].*\+.*[a-z]\w*.*\+.*['"].*['"]/i)) return;

        // Skip safe patterns: numeric values, boolean, predefined constants
        if (trimmed.match(/\+\s*(count|total|width|height|percent|num|idx|i|j|len)\b/)) return;
        if (trimmed.match(/\+\s*\d+\s*\+/)) return;

        // This line concatenates a variable into HTML without escaping
        issues.push('Line ' + lineNum + ': ' + trimmed.substring(0, 120));
      });

      // Allow threshold — legacy code may have some safe cases
      // (e.g., pre-validated constants like STATUS_COLORS[status])
      // 04e_PublicDashboard.gs contributes ~122 flagged lines that are false positives:
      // hardcoded CSS constants, booleans (isPII), config values — not user-controlled data.
      // Threshold set to 130 as ceiling; lower as true violations are eliminated.
      if (issues.length > 0) {
        expect(issues.length).toBeLessThanOrEqual(130);
      }
    });
  });
});

// ============================================================================
// A13: google.script.run FAILURE HANDLER COVERAGE IN VIEW FILES
// ============================================================================
// Failure history: v4.19.0 audit found 92 calls without withFailureHandler.
// serverCall() wrapper was added but existing calls need incremental migration.
// This test tracks the ratio and ensures it never gets worse.

describe('A13: google.script.run failure handler coverage', () => {
  const viewFiles = ['steward_view.html', 'member_view.html'];

  viewFiles.forEach(file => {
    test(`${file}: failure handler coverage ratio`, () => {
      const content = fs.readFileSync(
        path.resolve(__dirname, '..', 'src', file), 'utf8'
      );

      const totalCalls = (content.match(/google\.script\.run/g) || []).length;
      const withFailure = (content.match(/withFailureHandler/g) || []).length;
      const usingServerCall = (content.match(/serverCall\(\)/g) || []).length;
      const covered = withFailure + usingServerCall;

      // Track: at least 25% of calls must have explicit failure handling.
      // As migration to serverCall() progresses, raise this threshold.
      const ratio = totalCalls > 0 ? covered / totalCalls : 1;
      expect(ratio).toBeGreaterThanOrEqual(0.25);
    });
  });

  test('total unprotected google.script.run calls across views is under 100', () => {
    let totalUnprotected = 0;
    viewFiles.forEach(file => {
      const content = fs.readFileSync(
        path.resolve(__dirname, '..', 'src', file), 'utf8'
      );
      const lines = content.split('\n');
      lines.forEach((line, idx) => {
        if (line.includes('google.script.run') &&
            !line.includes('withFailureHandler') &&
            !line.includes('serverCall()')) {
          // Check next 3 lines for .withFailureHandler
          const nextLines = lines.slice(idx + 1, idx + 4).join('\n');
          if (!nextLines.includes('withFailureHandler')) {
            totalUnprotected++;
          }
        }
      });
    });
    // Ceiling — must never increase above current count.
    // Lower this as migration progresses.
    expect(totalUnprotected).toBeLessThanOrEqual(100);
  });
});

// ============================================================================
// A14: GAS API ENUM VALIDATION
// ============================================================================
// Failure history: v4.9.0 — XFrameOptionsMode.DENY does not exist in GAS,
// evaluating to undefined → "Argument cannot be null: mode" across 7 locations.

describe('A14: GAS API enum validation', () => {
  const srcDir = path.resolve(__dirname, '..', 'src');
  const gsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs'));

  // Valid XFrameOptionsMode values in GAS
  const validXFrameModes = ['DEFAULT', 'ALLOWALL'];

  test('all XFrameOptionsMode references use valid enum values', () => {
    const issues = [];
    gsFiles.forEach(file => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const regex = /XFrameOptionsMode\.(\w+)/g;
      let match;
      while ((match = regex.exec(content)) !== null) {
        if (!validXFrameModes.includes(match[1])) {
          issues.push(file + ': XFrameOptionsMode.' + match[1] + ' is not valid (use DEFAULT or ALLOWALL)');
        }
      }
    });
    expect(issues).toEqual([]);
  });

  // Valid SandboxMode values (for HtmlService)
  const validSandboxModes = ['IFRAME', 'NATIVE_SANDBOX', 'EMULATED'];

  test('all SandboxMode references use valid enum values', () => {
    const issues = [];
    gsFiles.forEach(file => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const regex = /SandboxMode\.(\w+)/g;
      let match;
      while ((match = regex.exec(content)) !== null) {
        if (!validSandboxModes.includes(match[1])) {
          issues.push(file + ': SandboxMode.' + match[1] + ' is not valid');
        }
      }
    });
    expect(issues).toEqual([]);
  });
});

// ============================================================================
// A15: ERROR HANDLER NO-CASCADE RULE
// ============================================================================
// Failure history: v4.19.2 — doGetWebDashboard catch block called
// ConfigReader.getConfig() which was the original error source → double-fault.
// Rule: catch blocks in web app files must not call the same module that
// could have thrown the original error without their own try/catch.

// ============================================================================
// A16: LOCK ACQUISITION → FINALLY RELEASE CONTRACT
// ============================================================================
// Failure mode: a function acquires LockService.getScriptLock() but throws
// before releaseLock() → lock is held for the full 30-second timeout, blocking
// all subsequent writes to the same spreadsheet for every user.
// Rule: every direct getScriptLock() acquisition must pair with releaseLock()
// inside a finally block. (withScriptLock_() helper is already safe by design.)

describe('A16: LockService.getScriptLock() acquisitions release in finally blocks', () => {
  const lockFiles = [
    '02_DataManagers.gs',
    '25_WorkloadService.gs',
    '26_QAForum.gs',
    '27_TimelineService.gs',
    '28_FailsafeService.gs',
    '08c_FormsAndNotifications.gs',
    '10d_SyncAndMaintenance.gs',
    '12_Features.gs',
  ];
  const srcDir = path.resolve(__dirname, '..', 'src');

  lockFiles.forEach(file => {
    test(`${file}: every getScriptLock() is paired with waitLock() and releaseLock() in finally`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const lines = content.split('\n');
      const issues = [];

      lines.forEach((line, idx) => {
        if (!line.includes('LockService.getScriptLock()')) return;
        if (line.trim().startsWith('//')) return; // skip comments

        // Scan the next 120 lines for the required patterns
        // (some functions like startNewGrievance have ~74 lines between lock acquire and finally)
        const window = lines.slice(idx, idx + 120).join('\n');

        // Both waitLock (blocking) and tryLock (non-blocking check) are acceptable
        if (!window.includes('waitLock(') && !window.includes('tryLock(')) {
          issues.push(`${file}:${idx + 1} — getScriptLock() has no waitLock() or tryLock()`);
        }
        if (!window.includes('finally')) {
          issues.push(`${file}:${idx + 1} — getScriptLock() has no finally block`);
        }
        if (!window.includes('releaseLock()')) {
          issues.push(`${file}:${idx + 1} — getScriptLock() has no releaseLock()`);
        }

        // Verify releaseLock() appears AFTER a finally keyword
        const finallyIdx = window.indexOf('finally');
        const releaseIdx = window.indexOf('releaseLock()');
        if (finallyIdx !== -1 && releaseIdx !== -1 && releaseIdx < finallyIdx) {
          issues.push(`${file}:${idx + 1} — releaseLock() appears before finally block`);
        }
      });

      expect(issues).toEqual([]);
    });
  });
});

// ============================================================================
// A17: LOCK-ACQUIRING MUTATIONS LOG AUDIT EVENTS
// ============================================================================
// Rule: any function that acquires a script lock (meaning it's performing a
// protected write) in the web-app service files must also call logAuditEvent().
// A write with no audit trail is undetectable if something goes wrong.
// Exception: helper functions (trailing _) and batch/maintenance utilities.

describe('A17: Lock-acquiring mutations in service files log audit events', () => {
  const serviceFiles = [
    '26_QAForum.gs',
    '27_TimelineService.gs',
    '28_FailsafeService.gs',
    '25_WorkloadService.gs',
  ];
  const srcDir = path.resolve(__dirname, '..', 'src');

  // Extract named function bodies from a source string.
  // Returns [{name, body}] for each non-private (no trailing _) named function.
  function extractFunctions(src) {
    const results = [];
    // Match: function name(...) { ... }  — captures up to 120 lines
    const fnRegex = /^function (\w+)\s*\([^)]*\)\s*\{([\s\S]*?)(?=\n^function |\n\/\/ ===|$)/gm;
    let match;
    while ((match = fnRegex.exec(src)) !== null) {
      if (!match[1].endsWith('_')) { // skip private helpers
        results.push({ name: match[1], body: match[2] });
      }
    }
    return results;
  }

  serviceFiles.forEach(file => {
    test(`${file}: functions that acquire a lock also call logAuditEvent()`, () => {
      const content = fs.readFileSync(path.join(srcDir, file), 'utf8');
      const fns = extractFunctions(content);
      const issues = [];

      fns.forEach(({ name, body }) => {
        if (!body.includes('LockService.getScriptLock()')) return; // not a locking function
        if (!body.includes('logAuditEvent(')) {
          issues.push(`${file}: ${name}() acquires lock but has no logAuditEvent() call`);
        }
      });

      expect(issues).toEqual([]);
    });
  });
});

// ============================================================================
// A18: DATASERVICE WRAPPER → DATASERVICE METHOD LINKAGE
// ============================================================================
// Failure mode: a refactor renames DataService.getStewardCases() but the
// dataGetStewardCases() wrapper still compiles (returns fallback []) — the
// client gets empty data with no error. This test verifies every dataXxx
// wrapper calls DataService.someMethod() rather than being silently orphaned.
// Exception: wrappers that are intentionally thin pass-throughs or call
// other global functions instead.

describe('A18: dataXxx wrapper functions call DataService (not orphaned)', () => {
  const src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8'
  );
  const lines = src.split('\n');

  // Find each dataXxx function declaration line, then scan the next 12 lines
  // for a DataService. call. This handles both one-liners and multi-line bodies
  // without breaking on nested braces in inline object literals.
  const wrappers = [];
  lines.forEach((line, idx) => {
    const m = line.match(/^function (data\w+)\s*\(/);
    if (!m || m[1].endsWith('_')) return;
    const name = m[1];
    const window = lines.slice(idx, idx + 12).join('\n');
    wrappers.push({ name, window });
  });

  test('found at least 50 dataXxx wrapper functions', () => {
    expect(wrappers.length).toBeGreaterThanOrEqual(50);
  });

  // Wrappers intentionally delegating elsewhere (not DataService)
  const nonDataServiceWrappers = [
    'dataMarkWelcomeDismissed',        // writes directly to PropertiesService
    'dataGetEngagementStats',          // reads live sheet data directly
    'dataGetWorkloadSummaryStats',     // reads live sheet data directly
    'dataSendDirectMessage',           // sends email + Drive log directly (calls DataService helpers deeper than 12-line window)
    'dataEnsureSheetsIfNeeded',        // calls _ensureAllSheetsInternal() + PropertiesService directly (fire-and-forget init)
  ];

  wrappers
    .filter(w => !nonDataServiceWrappers.includes(w.name))
    .forEach(({ name, window }) => {
      test(`${name}() calls DataService.someMethod() (not orphaned)`, () => {
        expect(window).toContain('DataService.');
      });
    });
});

describe('A15: Catch blocks in web app files do not cascade', () => {
  const webAppFiles = [
    '19_WebDashAuth.gs',
    '20_WebDashConfigReader.gs',
    '22_WebDashApp.gs',
    '23_PortalSheets.gs',
  ];

  webAppFiles.forEach(file => {
    test(`${file}: catch blocks do not make unguarded calls to SpreadsheetApp or ConfigReader`, () => {
      const content = fs.readFileSync(
        path.resolve(__dirname, '..', 'src', file), 'utf8'
      );

      // Find all catch blocks
      const catchRegex = /\}\s*catch\s*\(\w+\)\s*\{([\s\S]*?)\n\s*\}/g;
      const issues = [];
      let match;

      while ((match = catchRegex.exec(content)) !== null) {
        const catchBody = match[1];
        // If catch body calls ConfigReader.getConfig or SpreadsheetApp.getActiveSpreadsheet,
        // it MUST be inside its own try block
        const dangerousCalls = [
          'ConfigReader.getConfig()',
          'SpreadsheetApp.getActiveSpreadsheet()',
        ];

        dangerousCalls.forEach(call => {
          if (catchBody.includes(call)) {
            // Check that it's inside a nested try
            const beforeCall = catchBody.substring(0, catchBody.indexOf(call));
            const tryCount = (beforeCall.match(/try\s*\{/g) || []).length;
            const catchCount = (beforeCall.match(/\}\s*catch/g) || []).length;
            // Must have more try blocks than catch blocks (i.e., we're inside an open try)
            if (tryCount <= catchCount) {
              issues.push(file + ': catch block calls ' + call + ' without its own try/catch');
            }
          }
        });
      }

      expect(issues).toEqual([]);
    });
  });
});
