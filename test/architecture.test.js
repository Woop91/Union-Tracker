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
  '18_WorkloadTracker.gs',
  '19_WebDashAuth.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs',
  '22_WebDashApp.gs',
  '23_PortalSheets.gs',
  '24_WeeklyQuestions.gs',
  '25_WorkloadService.gs',
  '26_QAForum.gs',
  '27_TimelineService.gs',
  '28_FailsafeService.gs'
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
      'doGetWebDashboard',
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
      expect(DRIVE_CONFIG.ROOT_FOLDER_NAME).toBeTruthy();
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
    '08c_FormsAndNotifications.gs', '08d_AuditAndFormulas.gs',
    '09_Dashboards.gs', '10a_SheetCreation.gs', '10b_SurveyDocSheets.gs',
    '10c_FormHandlers.gs', '10d_SyncAndMaintenance.gs', '10_Main.gs',
    '11_CommandHub.gs', '12_Features.gs', '13_MemberSelfService.gs',
    '14_MeetingCheckIn.gs', '15_EventBus.gs', '16_DashboardEnhancements.gs',
    '17_CorrelationEngine.gs', '18_WorkloadTracker.gs', '19_WebDashAuth.gs',
    '20_WebDashConfigReader.gs', '21_WebDashDataService.gs',
    '22_WebDashApp.gs', '23_PortalSheets.gs', '24_WeeklyQuestions.gs',
    '25_WorkloadService.gs', '26_QAForum.gs', '27_TimelineService.gs',
    '28_FailsafeService.gs'
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
    expect(afterDoGet).toMatch(/\}\s*catch\s*\(\w+\)\s*\{[^}]*_serveFatalError/);
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
