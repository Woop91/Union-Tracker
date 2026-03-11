/**
 * Tests for 10a_SheetCreation.gs
 * Covers sheet creation, config seeding, header validation,
 * and dashboard setup functions.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '08a_SheetSetup.gs', '08d_AuditAndFormulas.gs',
  '10a_SheetCreation.gs'
]);

// ============================================================================
// Function existence
// ============================================================================

describe('10a function existence', () => {
  const required = [
    'createConfigSheet', 'seedConfigDefault_',
    'populateConfigFromSheetData', 'deduplicateAndSortConfigColumns_',
    'applyConfigSheetStyling', 'applySectionColors_', 'applyConfigStyling',
    'createConfigGuideSheet', '_addMissingMemberHeaders_',
    'createMemberDirectory', '_addMissingGrievanceHeaders_',
    'createGrievanceLog', 'createDashboard',
    'createVolunteerHoursSheet', 'createMeetingAttendanceSheet',
    'createMeetingCheckInLogSheet', 'setupMeetingCheckInSheet'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// Config column definitions are dynamic
// ============================================================================

describe('Config column structure', () => {
  test('CONFIG_COLS is defined and non-empty', () => {
    expect(typeof CONFIG_COLS).toBe('object');
    expect(Object.keys(CONFIG_COLS).length).toBeGreaterThan(10);
  });

  test('CONFIG_COLS has ORG_NAME', () => {
    expect(CONFIG_COLS.ORG_NAME).toBeDefined();
    expect(typeof CONFIG_COLS.ORG_NAME).toBe('number');
  });

  test('CONFIG_COLS has MAIN_CONTACT_EMAIL', () => {
    expect(CONFIG_COLS.MAIN_CONTACT_EMAIL).toBeDefined();
  });

  test('CONFIG_COLS has ACCENT_HUE', () => {
    expect(CONFIG_COLS.ACCENT_HUE).toBeDefined();
  });
});

// ============================================================================
// Member and Grievance headers use constants
// ============================================================================

describe('Member header map completeness', () => {
  test('MEMBER_HEADER_MAP_ is defined', () => {
    expect(typeof MEMBER_HEADER_MAP_).not.toBe('undefined');
  });

  test('MEMBER_COLS has essential fields', () => {
    expect(MEMBER_COLS.MEMBER_ID).toBeDefined();
    expect(MEMBER_COLS.FIRST_NAME).toBeDefined();
    expect(MEMBER_COLS.LAST_NAME).toBeDefined();
    expect(MEMBER_COLS.EMAIL).toBeDefined();
    expect(MEMBER_COLS.PHONE).toBeDefined();
    expect(MEMBER_COLS.IS_STEWARD).toBeDefined();
  });
});

describe('Grievance header map completeness', () => {
  test('GRIEVANCE_COLS has essential fields', () => {
    expect(GRIEVANCE_COLS.GRIEVANCE_ID).toBeDefined();
    expect(GRIEVANCE_COLS.MEMBER_ID).toBeDefined();
    expect(GRIEVANCE_COLS.STATUS).toBeDefined();
    expect(GRIEVANCE_COLS.CURRENT_STEP).toBeDefined();
  });
});
