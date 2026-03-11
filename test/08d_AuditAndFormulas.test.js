/**
 * Tests for 08d_AuditAndFormulas.gs
 * Covers audit logging, hidden sheet setup, vault integrity,
 * HMAC computation, and formula sheet creation.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '08d_AuditAndFormulas.gs'
]);

// ============================================================================
// Function existence
// ============================================================================

describe('08d function existence', () => {
  const required = [
    'setupAuditLogSheet', 'protectAuditLogSheet_', 'onEditAudit',
    'installAuditTrigger', 'removeAuditTrigger', 'viewAuditLog',
    'clearOldAuditEntries', 'getAuditHistory',
    'setupLiveGrievanceFormulas', 'setupGrievanceMemberDropdown',
    'setupGrievanceCalcSheet', 'setupGrievanceFormulasSheet',
    'setupMemberLookupSheet', 'setupStewardContactCalcSheet',
    'setupDashboardCalcSheet', 'setupStewardPerformanceCalcSheet',
    'setupAllHiddenSheets', 'repairAllHiddenSheets', 'verifyHiddenSheets',
    'refreshAllHiddenFormulas',
    'setupCalcMembersSheet', 'setupCalcGrievancesSheet',
    'setupCalcDeadlinesSheet', 'setupCalcStatsSheet',
    'setupCalcSyncSheet', 'setupCalcFormulasSheet',
    'setupSurveyTrackingSheet', 'setupSurveyVaultSheet',
    'getVaultDataMap_', 'getVaultDataFull_',
    'writeVaultEntry_', 'supersedePreviousVaultEntry_',
    'hashForVault_', 'hashForVaultLegacy_',
    'computeHmacSha256_', 'computeAuditRowHash_',
    'verifyAuditLogIntegrity', 'verifySurveyVaultIntegrity'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// HMAC / hashing
// ============================================================================

describe('computeHmacSha256_', () => {
  test('returns a string', () => {
    const result = computeHmacSha256_('test data', 'secret key');
    expect(typeof result).toBe('string');
  });

  test('different inputs produce different outputs', () => {
    const a = computeHmacSha256_('data A', 'key');
    const b = computeHmacSha256_('data B', 'key');
    expect(a).not.toBe(b);
  });

  test('same inputs produce same output (deterministic)', () => {
    const a = computeHmacSha256_('same data', 'same key');
    const b = computeHmacSha256_('same data', 'same key');
    expect(a).toBe(b);
  });
});

describe('hashForVault_', () => {
  test('returns a string', () => {
    const result = hashForVault_('member123', 'response data');
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });
});

describe('computeAuditRowHash_', () => {
  test('returns a string for array input', () => {
    const row = ['2026-03-11', 'test@example.com', 'EDIT', 'Changed field X'];
    const result = computeAuditRowHash_(row);
    expect(typeof result).toBe('string');
  });
});

// ============================================================================
// Hidden sheets constants
// ============================================================================

describe('Hidden sheet references use HIDDEN_SHEETS constants', () => {
  test('HIDDEN_SHEETS has all required calc sheets', () => {
    expect(HIDDEN_SHEETS.CALC_STATS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_FORMULAS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_MEMBERS).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_GRIEVANCES).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_DEADLINES).toBeDefined();
    expect(HIDDEN_SHEETS.CALC_SYNC).toBeDefined();
    expect(HIDDEN_SHEETS.AUDIT_LOG).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_TRACKING).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_VAULT).toBeDefined();
    expect(HIDDEN_SHEETS.SURVEY_PERIODS).toBeDefined();
  });

  test('all hidden sheet names start with underscore', () => {
    Object.entries(HIDDEN_SHEETS).forEach(([key, name]) => {
      expect(name).toMatch(/^_/);
    });
  });
});
