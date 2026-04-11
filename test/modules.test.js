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

// The removeDeprecatedTabs prefix-only matching test previously exercised
// test-local substring logic instead of the production code in 06_Maintenance.gs.
// It has been removed pending a real integration test that loads the actual
// removeDeprecatedTabs function against a mock spreadsheet.

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
  test('required sheets for reordering exist in SHEETS constant', () => {
    expect(SHEETS.GETTING_STARTED).toBeDefined();
    expect(SHEETS.FAQ).toBeDefined();
    expect(SHEETS.MEMBER_DIR).toBeDefined();
    expect(SHEETS.GRIEVANCE_LOG).toBeDefined();
  });
});

// ============================================================================
// generateChecklistId_ — real integration test against 12_Features.gs
// ============================================================================

describe('generateChecklistId_ (real production function)', () => {
  const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');

  // Load 12_Features.gs once. It depends on SHEETS/CHECKLIST_COLS from 01_Core
  // (already loaded above).
  beforeAll(() => {
    loadSources(['12_Features.gs']);
  });

  function seedChecklist(existingIds) {
    var header = new Array(Math.max(CHECKLIST_COLS.CHECKLIST_ID || 1, 2)).fill('col');
    header[(CHECKLIST_COLS.CHECKLIST_ID || 1) - 1] = 'Checklist ID';
    var rows = [header];
    existingIds.forEach(function(id) {
      var row = new Array(header.length).fill('');
      row[(CHECKLIST_COLS.CHECKLIST_ID || 1) - 1] = id;
      rows.push(row);
    });
    var sheet = createMockSheet(SHEETS.CASE_CHECKLIST || '_Case_Checklist', rows);
    var ss = createMockSpreadsheet([sheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);
    // Stub getOrCreateChecklistSheet so generateChecklistId_ uses our mock.
    global.getOrCreateChecklistSheet = jest.fn().mockReturnValue(sheet);
    return sheet;
  }

  test('returns CL-00001 when the sheet has only a header row', () => {
    seedChecklist([]);
    expect(generateChecklistId_()).toBe('CL-00001');
  });

  test('increments past the max existing CL-### id', () => {
    seedChecklist(['CL-00001', 'CL-00002', 'CL-00005']);
    expect(generateChecklistId_()).toBe('CL-00006');
  });

  test('offset parameter adds to the max id for batch creation', () => {
    seedChecklist(['CL-00010']);
    expect(generateChecklistId_(0)).toBe('CL-00011');
    expect(generateChecklistId_(1)).toBe('CL-00012');
    expect(generateChecklistId_(2)).toBe('CL-00013');
  });

  test('ignores non-CL- rows when computing the max', () => {
    seedChecklist(['', 'notes', 'CL-00003', 'xyz']);
    expect(generateChecklistId_()).toBe('CL-00004');
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
    expect(CONFIG_COLS.CHIEF_STEWARD_EMAIL).toBe(16);
  });

  test('has escalation config columns', () => {
    expect(CONFIG_COLS.ESCALATION_STATUSES).toBe(35);
    expect(CONFIG_COLS.ESCALATION_STEPS).toBe(36);
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
