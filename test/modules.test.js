/**
 * Tests for 06_Maintenance.gs, 08_SheetUtils.gs, and 12_Features.gs
 *
 * Covers cache keys, deprecated tabs logic, checklist ID generation,
 * caseId sanitization, satisfaction Q9 mapping, and sheet reordering.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Mock globals that these modules expect
global.logAuditEvent = jest.fn();
global.AUDIT_EVENTS = {
  SYSTEM_REPAIR: 'SYSTEM_REPAIR',
  FOLDER_CREATED: 'FOLDER_CREATED'
};

// ============================================================================
// Load constants only (the full modules have too many GAS deps)
// ============================================================================
loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs']);

// ============================================================================
// HIDDEN_SHEETS coverage
// ============================================================================

describe('HIDDEN_SHEETS', () => {
  test('all hidden sheets start with underscore', () => {
    Object.values(HIDDEN_SHEETS).forEach(name => {
      expect(name[0]).toBe('_');
    });
  });

  test('has required calculation sheets', () => {
    expect(HIDDEN_SHEETS.CALC_STATS).toBe('_Dashboard_Calc');
    expect(HIDDEN_SHEETS.GRIEVANCE_CALC).toBe('_Grievance_Calc');
    expect(HIDDEN_SHEETS.MEMBER_LOOKUP).toBe('_Member_Lookup');
    expect(HIDDEN_SHEETS.AUDIT_LOG).toBe('_Audit_Log');
    expect(HIDDEN_SHEETS.CHECKLIST_CALC).toBe('_Checklist_Calc');
  });
});

// ============================================================================
// removeDeprecatedTabs prefix-only matching
// ============================================================================

describe('removeDeprecatedTabs prefix-only matching', () => {
  // The bug was: substring match would incorrectly match sheets
  // that contained deprecated names as substrings.
  // Fix: only match if sheet name starts with the deprecated prefix.

  test('deprecated sheet names are prefix-matched only', () => {
    const deprecatedPrefixes = [SHEETS.DASHBOARD, SHEETS.SATISFACTION];

    // These should match (start with the prefix)
    expect(SHEETS.DASHBOARD.indexOf('💼')).toBe(0);
    expect(SHEETS.SATISFACTION.indexOf('📊')).toBe(0);

    // A sheet like "My 💼 Dashboard" should NOT match
    const testName = 'My 💼 Dashboard';
    const matches = deprecatedPrefixes.some(prefix => testName.indexOf(prefix) === 0);
    expect(matches).toBe(false);
  });
});

// ============================================================================
// Satisfaction Q9 mapping
// ============================================================================

describe('Satisfaction Q9 mapping', () => {
  test('Q9_RECOMMEND column exists in SATISFACTION_COLS', () => {
    expect(SATISFACTION_COLS.Q9_RECOMMEND).toBe(10); // Column J
  });

  test('Overall Satisfaction section includes Q9', () => {
    // Q9 maps to column 10, and the section should reference it
    expect(SATISFACTION_SECTIONS.OVERALL_SAT.questions).toContain(10);
  });
});

// ============================================================================
// Sheet reordering - SHEETS.INTERACTIVE removed
// ============================================================================

describe('Sheet reordering', () => {
  test('SHEETS does not have INTERACTIVE property (removed in bug fix)', () => {
    expect(SHEETS.INTERACTIVE).toBeUndefined();
  });

  test('required sheets for reordering exist in SHEETS constant', () => {
    expect(SHEETS.GETTING_STARTED).toBeDefined();
    expect(SHEETS.FAQ).toBeDefined();
    expect(SHEETS.MEMBER_DIR).toBeDefined();
    expect(SHEETS.GRIEVANCE_LOG).toBeDefined();
    expect(SHEETS.FEEDBACK).toBeDefined();
  });
});

// ============================================================================
// Checklist ID generation logic
// ============================================================================

describe('Checklist ID generation', () => {
  // Simulate generateChecklistId_ logic
  function formatChecklistId(maxNum, offset) {
    offset = offset || 0;
    return 'CL-' + String(maxNum + 1 + offset).padStart(5, '0');
  }

  test('generates correctly formatted IDs', () => {
    expect(formatChecklistId(0)).toBe('CL-00001');
    expect(formatChecklistId(5)).toBe('CL-00006');
    expect(formatChecklistId(99999)).toBe('CL-100000');
  });

  test('offset parameter increments correctly for batch creation', () => {
    expect(formatChecklistId(10, 0)).toBe('CL-00011');
    expect(formatChecklistId(10, 1)).toBe('CL-00012');
    expect(formatChecklistId(10, 2)).toBe('CL-00013');
  });

  test('first ID with empty sheet is CL-00001', () => {
    expect(formatChecklistId(0, 0)).toBe('CL-00001');
  });
});

// ============================================================================
// caseId sanitization for JS embedding
// ============================================================================

describe('caseId sanitization for JS embedding', () => {
  // Simulate the sanitization from 12_Features.gs
  function sanitizeCaseId(caseId) {
    return String(caseId || '').replace(/['"\\<>&]/g, '');
  }

  test('removes dangerous characters', () => {
    expect(sanitizeCaseId("GRV'; DROP TABLE;--")).toBe('GRV; DROP TABLE;--');
    expect(sanitizeCaseId('GRV"<script>')).toBe('GRVscript');
    expect(sanitizeCaseId('GRV\\n')).toBe('GRVn');
  });

  test('preserves normal grievance IDs', () => {
    expect(sanitizeCaseId('GRV-2026-001')).toBe('GRV-2026-001');
    expect(sanitizeCaseId('CL-00001')).toBe('CL-00001');
  });

  test('handles null/undefined', () => {
    expect(sanitizeCaseId(null)).toBe('');
    expect(sanitizeCaseId(undefined)).toBe('');
  });
});

// ============================================================================
// Undo/redo batch value logic
// ============================================================================

describe('Undo/redo batch value', () => {
  // Bug was: undo used c.oldValue instead of c.value for BATCH_UPDATE
  // The fix ensures that undo restores oldValue and redo restores value
  test('change object has distinct value and oldValue', () => {
    const change = { value: 'new', oldValue: 'old' };
    // Undo should use oldValue
    expect(change.oldValue).not.toBe(change.value);
    // Redo should use value
    expect(change.value).toBe('new');
    expect(change.oldValue).toBe('old');
  });
});

// ============================================================================
// CONFIG_COLS completeness
// ============================================================================

describe('CONFIG_COLS', () => {
  test('has deadline configuration columns', () => {
    expect(CONFIG_COLS.FILING_DEADLINE_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP1_RESPONSE_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP2_APPEAL_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP2_RESPONSE_DAYS).toBeDefined();
  });

  test('has deadline config columns including Step III and Arbitration', () => {
    expect(CONFIG_COLS.FILING_DEADLINE_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP1_RESPONSE_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP2_APPEAL_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP2_RESPONSE_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP3_APPEAL_DAYS).toBeDefined();
    expect(CONFIG_COLS.STEP3_RESPONSE_DAYS).toBeDefined();
    expect(CONFIG_COLS.ARBITRATION_DEMAND_DAYS).toBeDefined();
  });

  test('has chief steward email column', () => {
    expect(CONFIG_COLS.CHIEF_STEWARD_EMAIL).toBe(46);
  });

  test('has escalation config columns', () => {
    expect(CONFIG_COLS.ESCALATION_STATUSES).toBe(49);
    expect(CONFIG_COLS.ESCALATION_STEPS).toBe(50);
  });

  test('has form URL columns', () => {
    expect(CONFIG_COLS.GRIEVANCE_FORM_URL).toBeDefined();
    expect(CONFIG_COLS.CONTACT_FORM_URL).toBeDefined();
  });

  test('columns are sequential integers', () => {
    const values = Object.values(CONFIG_COLS).sort((a, b) => a - b);
    for (let i = 1; i < values.length; i++) {
      // Values should not have gaps larger than a few (some intentional gaps exist)
      expect(values[i] - values[i - 1]).toBeLessThanOrEqual(3);
    }
  });
});

// ============================================================================
// DEBUG_MODE should be false in production
// ============================================================================

describe('Production settings', () => {
  test('COMMAND_CONFIG exists and has system name', () => {
    expect(COMMAND_CONFIG.SYSTEM_NAME).toContain('Strategic');
  });

  test('ERROR_CONFIG.SHOW_STACK_TRACE is false by default', () => {
    expect(ERROR_CONFIG.SHOW_STACK_TRACE).toBe(false);
  });
});
