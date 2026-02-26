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
  '19_WebDashAuth.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs',
  '22_WebDashApp.gs',
  '23_PortalSheets.gs',
  '24_WeeklyQuestions.gs',
  '25_WorkloadService.gs'
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
      'sanitizeForHtml',
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
    '17_CorrelationEngine.gs',
    // Web-dashboard SPA modules (Union-Tracker excludes 18_WorkloadTracker.gs)
    '19_WebDashAuth.gs', '20_WebDashConfigReader.gs', '21_WebDashDataService.gs',
    '22_WebDashApp.gs', '23_PortalSheets.gs', '24_WeeklyQuestions.gs',
    '25_WorkloadService.gs'
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

  test('10c_FormHandlers.gs is under 700 lines (was 922 with stubs)', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10c_FormHandlers.gs'), 'utf8'
    );
    const lineCount = content.split('\n').length;
    expect(lineCount).toBeLessThan(700);
  });

  test('10d_SyncAndMaintenance.gs is under 1100 lines (was 1519 with stubs)', () => {
    const content = fs.readFileSync(
      path.resolve(__dirname, '..', 'src', '10d_SyncAndMaintenance.gs'), 'utf8'
    );
    const lineCount = content.split('\n').length;
    expect(lineCount).toBeLessThan(1100);
  });
});
