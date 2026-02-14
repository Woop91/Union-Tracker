/**
 * Tests for 12_Features.gs
 *
 * Covers EXTENSION_CONFIG, COL_IDX, LOOKER_CONFIG constants,
 * and pure categorization/bucket functions: generateAnonHash_,
 * getDaysBucket_, getContactFrequencyCategory_, getVolunteerHoursBucket_,
 * getEngagementLevel_, categorizeRole_, categorizeTenure_,
 * getScoreBucket_, calculateSectionAvg_, getQuarter_,
 * getOutcomeCategory_, getHeaderMap, and invalidateHeaderCache.
 */

require('./gas-mock');
const { createMockSheet, createMockSpreadsheet } = require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '12_Features.gs'
]);

// ============================================================================
// EXTENSION_CONFIG
// ============================================================================

describe('EXTENSION_CONFIG', () => {
  test('is defined', () => {
    expect(EXTENSION_CONFIG).toBeDefined();
  });

  test('has correct HIDDEN_CALC_SHEET name', () => {
    expect(EXTENSION_CONFIG.HIDDEN_CALC_SHEET).toBe('_Dashboard_Calc');
  });

  test('has DYNAMIC_FORMULA_ROW of 50', () => {
    expect(EXTENSION_CONFIG.DYNAMIC_FORMULA_ROW).toBe(50);
  });

  test('has MEMBER_SHEET name', () => {
    expect(EXTENSION_CONFIG.MEMBER_SHEET).toBe('Member Directory');
  });

  test('has GRIEVANCE_SHEET name', () => {
    expect(EXTENSION_CONFIG.GRIEVANCE_SHEET).toBe('Grievance Log');
  });

  test('has LEADER_ROLE_NAME', () => {
    expect(EXTENSION_CONFIG.LEADER_ROLE_NAME).toBe('Member Leader');
  });

  test('has CORE_COLUMN_COUNT of 32', () => {
    expect(EXTENSION_CONFIG.CORE_COLUMN_COUNT).toBe(32);
  });

  test('has CACHE_TTL_SECONDS of 300', () => {
    expect(EXTENSION_CONFIG.CACHE_TTL_SECONDS).toBe(300);
  });
});

// ============================================================================
// COL_IDX
// ============================================================================

describe('COL_IDX', () => {
  test('is defined', () => {
    expect(COL_IDX).toBeDefined();
  });

  test('MEMBER_ID is 0-based (MEMBER_COLS.MEMBER_ID - 1)', () => {
    expect(COL_IDX.MEMBER_ID).toBe(MEMBER_COLS.MEMBER_ID - 1);
  });

  test('FIRST_NAME is 0-based', () => {
    expect(COL_IDX.FIRST_NAME).toBe(MEMBER_COLS.FIRST_NAME - 1);
  });

  test('EMAIL is 0-based', () => {
    expect(COL_IDX.EMAIL).toBe(MEMBER_COLS.EMAIL - 1);
  });
});

// ============================================================================
// LOOKER_CONFIG
// ============================================================================

describe('LOOKER_CONFIG', () => {
  test('is defined', () => {
    expect(LOOKER_CONFIG).toBeDefined();
  });

  test('SHEETS has GRIEVANCES, MEMBERS, SATISFACTION keys', () => {
    expect(LOOKER_CONFIG.SHEETS.GRIEVANCES).toBeDefined();
    expect(LOOKER_CONFIG.SHEETS.MEMBERS).toBeDefined();
    expect(LOOKER_CONFIG.SHEETS.SATISFACTION).toBeDefined();
  });

  test('SHEETS_ANON has GRIEVANCES, MEMBERS, SATISFACTION keys', () => {
    expect(LOOKER_CONFIG.SHEETS_ANON.GRIEVANCES).toBeDefined();
    expect(LOOKER_CONFIG.SHEETS_ANON.MEMBERS).toBeDefined();
    expect(LOOKER_CONFIG.SHEETS_ANON.SATISFACTION).toBeDefined();
  });

  test('ALLOWED_SOURCES has 3 entries', () => {
    expect(LOOKER_CONFIG.ALLOWED_SOURCES.length).toBe(3);
  });

  test('AUTO_REFRESH_HOUR is a number', () => {
    expect(typeof LOOKER_CONFIG.AUTO_REFRESH_HOUR).toBe('number');
  });
});

// ============================================================================
// generateAnonHash_
// ============================================================================

describe('generateAnonHash_', () => {
  test('returns a string starting with "A"', () => {
    const result = generateAnonHash_('MBR-001');
    expect(result[0]).toBe('A');
  });

  test('is deterministic (same input produces same output)', () => {
    const first = generateAnonHash_('MBR-001');
    const second = generateAnonHash_('MBR-001');
    expect(first).toBe(second);
  });

  test('different inputs produce different outputs', () => {
    const hash1 = generateAnonHash_('MBR-001');
    const hash2 = generateAnonHash_('MBR-002');
    expect(hash1).not.toBe(hash2);
  });

  test('result length is at most 9 characters (A + up to 8)', () => {
    const result = generateAnonHash_('MBR-001');
    expect(result.length).toBeLessThanOrEqual(9);
  });

  test('handles empty string input', () => {
    const result = generateAnonHash_('');
    expect(typeof result).toBe('string');
    expect(result[0]).toBe('A');
  });
});

// ============================================================================
// getDaysBucket_
// ============================================================================

describe('getDaysBucket_', () => {
  test('0 days -> Within Week', () => {
    expect(getDaysBucket_(0)).toBe('Within Week');
  });

  test('7 days -> Within Week', () => {
    expect(getDaysBucket_(7)).toBe('Within Week');
  });

  test('8 days -> Within Month', () => {
    expect(getDaysBucket_(8)).toBe('Within Month');
  });

  test('30 days -> Within Month', () => {
    expect(getDaysBucket_(30)).toBe('Within Month');
  });

  test('31 days -> 1-3 Months', () => {
    expect(getDaysBucket_(31)).toBe('1-3 Months');
  });

  test('90 days -> 1-3 Months', () => {
    expect(getDaysBucket_(90)).toBe('1-3 Months');
  });

  test('91 days -> 3-6 Months', () => {
    expect(getDaysBucket_(91)).toBe('3-6 Months');
  });

  test('180 days -> 3-6 Months', () => {
    expect(getDaysBucket_(180)).toBe('3-6 Months');
  });

  test('181 days -> 6-12 Months', () => {
    expect(getDaysBucket_(181)).toBe('6-12 Months');
  });

  test('365 days -> 6-12 Months', () => {
    expect(getDaysBucket_(365)).toBe('6-12 Months');
  });

  test('366 days -> Over 1 Year', () => {
    expect(getDaysBucket_(366)).toBe('Over 1 Year');
  });

  test('1000 days -> Over 1 Year', () => {
    expect(getDaysBucket_(1000)).toBe('Over 1 Year');
  });
});

// ============================================================================
// getContactFrequencyCategory_
// ============================================================================

describe('getContactFrequencyCategory_', () => {
  const now = new Date('2026-02-10');

  test('non-Date input returns "No Contact"', () => {
    expect(getContactFrequencyCategory_('not a date', now)).toBe('No Contact');
  });

  test('null input returns "No Contact"', () => {
    expect(getContactFrequencyCategory_(null, now)).toBe('No Contact');
  });

  test('contact within 30 days -> Active', () => {
    const recent = new Date('2026-01-20');
    expect(getContactFrequencyCategory_(recent, now)).toBe('Active');
  });

  test('contact 31-90 days ago -> Regular', () => {
    const contact = new Date('2025-12-01');
    expect(getContactFrequencyCategory_(contact, now)).toBe('Regular');
  });

  test('contact 91-180 days ago -> Occasional', () => {
    const contact = new Date('2025-09-01');
    expect(getContactFrequencyCategory_(contact, now)).toBe('Occasional');
  });

  test('contact over 180 days ago -> Inactive', () => {
    const contact = new Date('2025-01-01');
    expect(getContactFrequencyCategory_(contact, now)).toBe('Inactive');
  });
});

// ============================================================================
// getVolunteerHoursBucket_
// ============================================================================

describe('getVolunteerHoursBucket_', () => {
  test('0 hours -> None', () => {
    expect(getVolunteerHoursBucket_(0)).toBe('None');
  });

  test('1 hour -> 1-5 Hours', () => {
    expect(getVolunteerHoursBucket_(1)).toBe('1-5 Hours');
  });

  test('5 hours -> 1-5 Hours', () => {
    expect(getVolunteerHoursBucket_(5)).toBe('1-5 Hours');
  });

  test('6 hours -> 6-20 Hours', () => {
    expect(getVolunteerHoursBucket_(6)).toBe('6-20 Hours');
  });

  test('20 hours -> 6-20 Hours', () => {
    expect(getVolunteerHoursBucket_(20)).toBe('6-20 Hours');
  });

  test('21 hours -> 21-50 Hours', () => {
    expect(getVolunteerHoursBucket_(21)).toBe('21-50 Hours');
  });

  test('50 hours -> 21-50 Hours', () => {
    expect(getVolunteerHoursBucket_(50)).toBe('21-50 Hours');
  });

  test('51 hours -> 50+ Hours', () => {
    expect(getVolunteerHoursBucket_(51)).toBe('50+ Hours');
  });
});

// ============================================================================
// getEngagementLevel_
// ============================================================================

describe('getEngagementLevel_', () => {
  test('all zeros -> Not Engaged', () => {
    expect(getEngagementLevel_(0, null, 'No', 0)).toBe('Not Engaged');
  });

  test('minimal involvement (1 grievance only) -> Low Engagement (score=1)', () => {
    expect(getEngagementLevel_(0, null, 'No', 1)).toBe('Low Engagement');
  });

  test('moderate volunteer + recent contact -> Somewhat Engaged or higher', () => {
    const recentDate = new Date(Date.now() - 15 * 86400000); // 15 days ago
    const result = getEngagementLevel_(6, recentDate, 'No', 0);
    // score: volunteer 6>5 = 1, contact <=30 = 2, total = 3
    expect(result).toBe('Somewhat Engaged');
  });

  test('steward with moderate hours -> Engaged', () => {
    const recentDate = new Date(Date.now() - 60 * 86400000); // 60 days ago
    // score: volunteer 21>20 = 2, contact 60<=90 = 1, steward Yes = 2, total = 5
    expect(getEngagementLevel_(21, recentDate, 'Yes', 0)).toBe('Engaged');
  });

  test('steward with high hours and recent contact -> Highly Engaged', () => {
    const recentDate = new Date(Date.now() - 10 * 86400000); // 10 days ago
    // score: volunteer 51>50 = 3, contact <=30 = 2, steward Yes = 2, total = 7
    expect(getEngagementLevel_(51, recentDate, 'Yes', 0)).toBe('Highly Engaged');
  });

  test('Member Leader role counts as leadership', () => {
    // score: volunteer 0 = 0, no contact = 0, leader = 2, grievance = 1, total = 3
    expect(getEngagementLevel_(0, null, 'Member Leader', 1)).toBe('Somewhat Engaged');
  });

  test('high hours alone can reach Somewhat Engaged', () => {
    // score: volunteer 51>50 = 3, no contact = 0, no steward = 0, no grievance = 0, total = 3
    expect(getEngagementLevel_(51, null, 'No', 0)).toBe('Somewhat Engaged');
  });
});

// ============================================================================
// categorizeRole_
// ============================================================================

describe('categorizeRole_', () => {
  test('"Steward" -> Leadership', () => {
    expect(categorizeRole_('Steward')).toBe('Leadership');
  });

  test('"Chief Steward" -> Leadership', () => {
    expect(categorizeRole_('Chief Steward')).toBe('Leadership');
  });

  test('"Member Leader" -> Leadership', () => {
    expect(categorizeRole_('Member Leader')).toBe('Leadership');
  });

  test('"Team Leader" -> Leadership', () => {
    expect(categorizeRole_('Team Leader')).toBe('Leadership');
  });

  test('"RN" -> Nursing', () => {
    expect(categorizeRole_('RN')).toBe('Nursing');
  });

  test('"Nurse Practitioner" -> Nursing', () => {
    expect(categorizeRole_('Nurse Practitioner')).toBe('Nursing');
  });

  test('"LPN" -> Nursing', () => {
    expect(categorizeRole_('LPN')).toBe('Nursing');
  });

  test('"Lab Tech" -> Technical/Support', () => {
    expect(categorizeRole_('Lab Tech')).toBe('Technical/Support');
  });

  test('"Home Health Aide" -> Technical/Support', () => {
    expect(categorizeRole_('Home Health Aide')).toBe('Technical/Support');
  });

  test('"Admin Assistant" -> Administrative', () => {
    expect(categorizeRole_('Admin Assistant')).toBe('Administrative');
  });

  test('"Unit Clerk" -> Administrative', () => {
    expect(categorizeRole_('Unit Clerk')).toBe('Administrative');
  });

  test('"Social Worker" -> Other', () => {
    expect(categorizeRole_('Social Worker')).toBe('Other');
  });

  test('empty string -> Other', () => {
    expect(categorizeRole_('')).toBe('Other');
  });
});

// ============================================================================
// categorizeTenure_
// ============================================================================

describe('categorizeTenure_', () => {
  test('"Less than 1 year" -> New (< 1 year)', () => {
    expect(categorizeTenure_('Less than 1 year')).toBe('New (< 1 year)');
  });

  test('"< 1 year" -> New (< 1 year)', () => {
    expect(categorizeTenure_('< 1 year')).toBe('New (< 1 year)');
  });

  test('"1-3 years" -> 1-3 Years', () => {
    expect(categorizeTenure_('1-3 years')).toBe('1-3 Years');
  });

  test('"1 to 3 years" -> 1-3 Years', () => {
    expect(categorizeTenure_('1 to 3 years')).toBe('1-3 Years');
  });

  test('"3-5 years" -> 3-5 Years', () => {
    expect(categorizeTenure_('3-5 years')).toBe('3-5 Years');
  });

  test('"3 to 5 years" -> 3-5 Years', () => {
    expect(categorizeTenure_('3 to 5 years')).toBe('3-5 Years');
  });

  test('"5-10 years" -> 5-10 Years', () => {
    expect(categorizeTenure_('5-10 years')).toBe('5-10 Years');
  });

  test('"5 to 10 years" -> 5-10 Years', () => {
    expect(categorizeTenure_('5 to 10 years')).toBe('5-10 Years');
  });

  test('"Over 10 years" -> 10+ Years', () => {
    expect(categorizeTenure_('Over 10 years')).toBe('10+ Years');
  });

  test('"20 years" -> 10+ Years (default)', () => {
    expect(categorizeTenure_('20 years')).toBe('10+ Years');
  });

  test('empty string -> 10+ Years (default)', () => {
    expect(categorizeTenure_('')).toBe('10+ Years');
  });
});

// ============================================================================
// getScoreBucket_
// ============================================================================

describe('getScoreBucket_', () => {
  test('NaN input -> No Response', () => {
    expect(getScoreBucket_('N/A')).toBe('No Response');
  });

  test('empty string -> No Response', () => {
    expect(getScoreBucket_('')).toBe('No Response');
  });

  test('null -> No Response', () => {
    expect(getScoreBucket_(null)).toBe('No Response');
  });

  test('undefined -> No Response', () => {
    expect(getScoreBucket_(undefined)).toBe('No Response');
  });

  test('score 10 -> High (8-10)', () => {
    expect(getScoreBucket_(10)).toBe('High (8-10)');
  });

  test('score 8 -> High (8-10)', () => {
    expect(getScoreBucket_(8)).toBe('High (8-10)');
  });

  test('score 7.9 -> Medium (5-7)', () => {
    expect(getScoreBucket_(7.9)).toBe('Medium (5-7)');
  });

  test('score 5 -> Medium (5-7)', () => {
    expect(getScoreBucket_(5)).toBe('Medium (5-7)');
  });

  test('score 4.9 -> Low (1-4)', () => {
    expect(getScoreBucket_(4.9)).toBe('Low (1-4)');
  });

  test('score 1 -> Low (1-4)', () => {
    expect(getScoreBucket_(1)).toBe('Low (1-4)');
  });

  test('string "9" is parsed as number -> High (8-10)', () => {
    expect(getScoreBucket_('9')).toBe('High (8-10)');
  });
});

// ============================================================================
// calculateSectionAvg_
// ============================================================================

describe('calculateSectionAvg_', () => {
  test('calculates average of valid values (1-10)', () => {
    const row = [0, 8, 0, 6, 0, 10];
    const result = calculateSectionAvg_(row, [1, 3, 5]);
    // (8 + 6 + 10) / 3 = 8.0
    expect(result).toBe(8);
  });

  test('ignores values outside 1-10 range', () => {
    const row = [0, 8, 0, 0, 15, 5];
    // col 1 = 8 (valid), col 3 = 0 (invalid), col 4 = 15 (invalid), col 5 = 5 (valid)
    const result = calculateSectionAvg_(row, [1, 3, 4, 5]);
    // (8 + 5) / 2 = 6.5
    expect(result).toBe(6.5);
  });

  test('returns empty string when no valid values', () => {
    const row = [0, 0, 0, 11, -1];
    const result = calculateSectionAvg_(row, [0, 1, 2, 3, 4]);
    expect(result).toBe('');
  });

  test('rounds to one decimal place', () => {
    const row = [7, 8, 9];
    // (7 + 8 + 9) / 3 = 8.0
    const result = calculateSectionAvg_(row, [0, 1, 2]);
    expect(result).toBe(8);
  });

  test('handles single valid value', () => {
    const row = [0, 0, 7];
    const result = calculateSectionAvg_(row, [0, 1, 2]);
    expect(result).toBe(7);
  });

  test('handles string numeric values', () => {
    const row = ['', '8', '', '6'];
    const result = calculateSectionAvg_(row, [1, 3]);
    expect(result).toBe(7);
  });
});

// ============================================================================
// getQuarter_
// ============================================================================

describe('getQuarter_', () => {
  test('January -> Q1', () => {
    expect(getQuarter_(new Date('2026-01-15'))).toBe('2026-Q1');
  });

  test('March -> Q1', () => {
    expect(getQuarter_(new Date('2026-03-31'))).toBe('2026-Q1');
  });

  test('April -> Q2', () => {
    expect(getQuarter_(new Date('2026-04-01'))).toBe('2026-Q2');
  });

  test('June -> Q2', () => {
    expect(getQuarter_(new Date('2026-06-30'))).toBe('2026-Q2');
  });

  test('July -> Q3', () => {
    expect(getQuarter_(new Date('2026-07-15'))).toBe('2026-Q3');
  });

  test('October -> Q4', () => {
    expect(getQuarter_(new Date('2026-10-01'))).toBe('2026-Q4');
  });

  test('December -> Q4', () => {
    expect(getQuarter_(new Date('2026-12-31'))).toBe('2026-Q4');
  });
});

// ============================================================================
// getOutcomeCategory_
// ============================================================================

describe('getOutcomeCategory_', () => {
  test('"Won" -> Win', () => {
    expect(getOutcomeCategory_('Won')).toBe('Win');
  });

  test('"Denied" -> Loss', () => {
    expect(getOutcomeCategory_('Denied')).toBe('Loss');
  });

  test('"Settled" -> Settlement', () => {
    expect(getOutcomeCategory_('Settled')).toBe('Settlement');
  });

  test('"Withdrawn" -> Withdrawn', () => {
    expect(getOutcomeCategory_('Withdrawn')).toBe('Withdrawn');
  });

  test('"Closed" -> Closed', () => {
    expect(getOutcomeCategory_('Closed')).toBe('Closed');
  });

  test('"Open" -> Active (default)', () => {
    expect(getOutcomeCategory_('Open')).toBe('Active');
  });

  test('"In Progress" -> Active (default)', () => {
    expect(getOutcomeCategory_('In Progress')).toBe('Active');
  });

  test('empty string -> Active (default)', () => {
    expect(getOutcomeCategory_('')).toBe('Active');
  });
});

// ============================================================================
// getHeaderMap
// ============================================================================

describe('getHeaderMap', () => {
  test('returns header map from sheet', () => {
    const headers = [['Name', 'Email', 'Phone']];
    const mockSheet = createMockSheet('TestSheet', headers);
    mockSheet.getLastColumn.mockReturnValue(3);
    mockSheet.getRange.mockReturnValue({
      getValues: jest.fn(() => [['Name', 'Email', 'Phone']]),
      getValue: jest.fn(),
      setValue: jest.fn(),
      setValues: jest.fn(),
      setFontWeight: jest.fn(function() { return this; }),
      setBackground: jest.fn(function() { return this; }),
      setFontColor: jest.fn(function() { return this; })
    });

    const ss = createMockSpreadsheet([mockSheet]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    const result = getHeaderMap('TestSheet', true);
    expect(result['Name']).toBe(1);
    expect(result['Email']).toBe(2);
    expect(result['Phone']).toBe(3);
  });

  test('returns empty object when sheet does not exist', () => {
    const ss = createMockSpreadsheet([]);
    SpreadsheetApp.getActiveSpreadsheet.mockReturnValue(ss);

    const result = getHeaderMap('NonExistentSheet', true);
    expect(result).toEqual({});
  });
});

// ============================================================================
// invalidateHeaderCache
// ============================================================================

describe('invalidateHeaderCache', () => {
  test('does not throw', () => {
    expect(() => invalidateHeaderCache('TestSheet')).not.toThrow();
  });

  test('invalidateHeaderCache is callable and does not throw', () => {
    expect(typeof invalidateHeaderCache).toBe('function');
    expect(() => invalidateHeaderCache('Member Directory')).not.toThrow();
  });
});
