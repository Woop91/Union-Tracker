// Tests functions from: 10c_FormsAndSync.gs
/**
 * Tests for 10c_FormHandlers.gs
 * Covers form field management, timeline views, and column operations.
 */

const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources([
  '00_Security.gs', '00_DataAccess.gs', '01_Core.gs',
  '02_DataManagers.gs', '03_UIComponents.gs',
  '05_Integrations.gs', '10c_FormsAndSync.gs'
]);

describe('10c function existence', () => {
  const required = [
    'getCurrentStewardInfo_',
    'sanitizeFolderName_', 'shareWithCoordinators_',
    'refreshMemberDirectoryFormulas',
    'refreshAllFormulas'
  ];

  required.forEach(fn => {
    test(`${fn} is defined`, () => {
      expect(typeof global[fn]).toBe('function');
    });
  });
});

// ============================================================================
// sanitizeFolderName_ — behavioral tests
// ============================================================================

describe('sanitizeFolderName_', () => {
  test('removes special characters', () => {
    const result = sanitizeFolderName_('Test / Folder : Name');
    expect(result).not.toContain('/');
    expect(result).not.toContain(':');
  });

  test('handles empty string', () => {
    const result = sanitizeFolderName_('');
    expect(typeof result).toBe('string');
  });

  test('replaces spaces with underscores', () => {
    const result = sanitizeFolderName_('Hello World Test');
    expect(result).toBe('Hello_World_Test');
  });

  test('collapses multiple spaces into single underscore', () => {
    const result = sanitizeFolderName_('Hello    World');
    expect(result).toBe('Hello_World');
  });

  test('strips all GAS-invalid characters: < > : " / \\ | ? *', () => {
    const result = sanitizeFolderName_('a<b>c:d"e/f\\g|h?i*j');
    expect(result).toBe('abcdefghij');
  });

  test('truncates to 50 characters', () => {
    const longName = 'A'.repeat(80);
    const result = sanitizeFolderName_(longName);
    expect(result.length).toBeLessThanOrEqual(50);
  });

  test('trims leading and trailing whitespace', () => {
    const result = sanitizeFolderName_('  hello  ');
    expect(result).toBe('hello');
  });

  test('handles null/undefined input gracefully', () => {
    // sanitizeFolderName_ delegates to sanitizeFolderName which returns 'Unknown' for falsy
    const result = sanitizeFolderName_(null);
    expect(typeof result).toBe('string');
    expect(result).toBe('Unknown');
  });
});

// ============================================================================
// getCurrentStewardInfo_ — behavioral tests
// ============================================================================

describe('getCurrentStewardInfo_', () => {
  test('returns steward info when current user is a steward', () => {
    var header = new Array(20).fill('');
    var row1 = new Array(20).fill('');
    row1[MEMBER_COLS.EMAIL - 1] = 'test@example.com';
    row1[MEMBER_COLS.IS_STEWARD - 1] = 'Yes';
    row1[MEMBER_COLS.FIRST_NAME - 1] = 'Jane';
    row1[MEMBER_COLS.LAST_NAME - 1] = 'Doe';
    var data = [header, row1];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);

    var info = getCurrentStewardInfo_(mockSS);
    expect(info.firstName).toBe('Jane');
    expect(info.lastName).toBe('Doe');
    expect(info.email).toBe('test@example.com');
  });

  test('returns email-only info when user is not a steward', () => {
    var header = new Array(20).fill('');
    var row1 = new Array(20).fill('');
    row1[MEMBER_COLS.EMAIL - 1] = 'test@example.com';
    row1[MEMBER_COLS.IS_STEWARD - 1] = 'No';
    row1[MEMBER_COLS.FIRST_NAME - 1] = 'Jane';
    var data = [header, row1];
    var memberSheet = createMockSheet(SHEETS.MEMBER_DIR, data);
    var mockSS = createMockSpreadsheet([memberSheet]);

    var info = getCurrentStewardInfo_(mockSS);
    expect(info.firstName).toBe('');
    expect(info.lastName).toBe('');
    expect(info.email).toBe('test@example.com');
  });

  test('returns email-only when member sheet is missing', () => {
    var mockSS = createMockSpreadsheet([]);

    var info = getCurrentStewardInfo_(mockSS);
    expect(info.email).toBe('test@example.com');
    expect(info.firstName).toBe('');
  });
});
