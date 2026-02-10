/**
 * Tests for 07_DevTools.gs
 *
 * Covers pure utility functions (randomChoice, shuffleArray, randomDate, addDays),
 * validation patterns and messages (VALIDATION_PATTERNS, VALIDATION_MESSAGES),
 * the Assert object, validation functions (validateEmailAddress, validatePhoneNumber,
 * formatUSPhone, validateRequired), and test helper functions (isValidEmail_,
 * isValidPhone_, formatDate_, trimString_, isValidMemberId_, isValidGrievanceId_,
 * getUniqueValues_, flattenArray_, hasProperty_).
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load files in GAS load order
loadSources([
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '07_DevTools.gs'
]);

// ============================================================================
// Pure Utility - randomChoice
// ============================================================================

describe('randomChoice', () => {
  test('returns an element from the array', () => {
    const arr = ['a', 'b', 'c', 'd'];
    const result = randomChoice(arr);
    expect(arr).toContain(result);
  });

  test('returns the only element from single-element array', () => {
    expect(randomChoice([42])).toBe(42);
  });

  test('works with various data types', () => {
    const mixed = [1, 'hello', true, null];
    const result = randomChoice(mixed);
    expect(mixed).toContain(result);
  });
});

// ============================================================================
// Pure Utility - shuffleArray
// ============================================================================

describe('shuffleArray', () => {
  test('returns array of same length', () => {
    const arr = [1, 2, 3, 4, 5];
    const shuffled = shuffleArray(arr);
    expect(shuffled.length).toBe(arr.length);
  });

  test('does not mutate original array', () => {
    const arr = [1, 2, 3, 4, 5];
    const original = [...arr];
    shuffleArray(arr);
    expect(arr).toEqual(original);
  });

  test('contains all original elements', () => {
    const arr = [10, 20, 30, 40, 50];
    const shuffled = shuffleArray(arr);
    expect(shuffled.sort((a, b) => a - b)).toEqual(arr.sort((a, b) => a - b));
  });

  test('returns new array reference', () => {
    const arr = [1, 2, 3];
    const shuffled = shuffleArray(arr);
    expect(shuffled).not.toBe(arr);
  });

  test('handles empty array', () => {
    expect(shuffleArray([])).toEqual([]);
  });

  test('handles single element array', () => {
    expect(shuffleArray([99])).toEqual([99]);
  });
});

// ============================================================================
// Pure Utility - randomDate
// ============================================================================

describe('randomDate', () => {
  test('returns a Date object', () => {
    const start = new Date(2020, 0, 1);
    const end = new Date(2025, 0, 1);
    const result = randomDate(start, end);
    expect(result instanceof Date).toBe(true);
  });

  test('returns date within range', () => {
    const start = new Date(2023, 0, 1);
    const end = new Date(2023, 11, 31);
    const result = randomDate(start, end);
    expect(result.getTime()).toBeGreaterThanOrEqual(start.getTime());
    expect(result.getTime()).toBeLessThanOrEqual(end.getTime());
  });

  test('works when start equals end', () => {
    const date = new Date(2023, 5, 15);
    const result = randomDate(date, date);
    expect(result.getTime()).toBe(date.getTime());
  });
});

// ============================================================================
// Pure Utility - addDays
// ============================================================================

describe('addDays', () => {
  test('adds positive days to a date', () => {
    const date = new Date(2023, 0, 1); // Jan 1
    const result = addDays(date, 10);
    expect(result instanceof Date).toBe(true);
    expect(result.getDate()).toBe(11);
  });

  test('subtracts days with negative value', () => {
    const date = new Date(2023, 0, 15); // Jan 15
    const result = addDays(date, -5);
    expect(result.getDate()).toBe(10);
  });

  test('returns empty string for falsy date input', () => {
    expect(addDays(null, 5)).toBe('');
    expect(addDays(undefined, 5)).toBe('');
    expect(addDays('', 5)).toBe('');
  });

  test('does not mutate the original date', () => {
    const date = new Date(2023, 0, 1);
    const originalTime = date.getTime();
    addDays(date, 30);
    expect(date.getTime()).toBe(originalTime);
  });

  test('handles month boundary crossing', () => {
    const date = new Date(2023, 0, 28); // Jan 28
    const result = addDays(date, 5);
    expect(result.getMonth()).toBe(1); // February
    expect(result.getDate()).toBe(2);
  });
});

// ============================================================================
// Constants - VALIDATION_PATTERNS (from 07_DevTools.gs, overrides earlier defs)
// ============================================================================

describe('VALIDATION_PATTERNS', () => {
  test('VALIDATION_PATTERNS is defined', () => {
    expect(VALIDATION_PATTERNS).toBeDefined();
  });

  test('EMAIL regex matches valid emails', () => {
    expect(VALIDATION_PATTERNS.EMAIL.test('user@example.com')).toBe(true);
    expect(VALIDATION_PATTERNS.EMAIL.test('first.last+tag@domain.co.uk')).toBe(true);
  });

  test('EMAIL regex rejects invalid emails', () => {
    expect(VALIDATION_PATTERNS.EMAIL.test('noatsign')).toBe(false);
    expect(VALIDATION_PATTERNS.EMAIL.test('@nodomain.com')).toBe(false);
    expect(VALIDATION_PATTERNS.EMAIL.test('user@')).toBe(false);
  });

  test('MEMBER_ID regex matches M + 4 uppercase + 3 digits', () => {
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM123')).toBe(true);
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('MMAJO456')).toBe(true);
  });

  test('MEMBER_ID regex rejects invalid formats', () => {
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('JOSM123')).toBe(false);   // missing M prefix
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('MJOS123')).toBe(false);   // only 3 letters
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('MJOSM12')).toBe(false);   // only 2 digits
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('mjosm123')).toBe(false);  // lowercase
    expect(VALIDATION_PATTERNS.MEMBER_ID.test('M123456')).toBe(false);   // old format
  });

  test('GRIEVANCE_ID regex matches G + 4 uppercase + 3 digits', () => {
    expect(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM789')).toBe(true);
    expect(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GROWI001')).toBe(true);
  });

  test('GRIEVANCE_ID regex rejects invalid formats', () => {
    expect(VALIDATION_PATTERNS.GRIEVANCE_ID.test('JOSM789')).toBe(false);   // missing G prefix
    expect(VALIDATION_PATTERNS.GRIEVANCE_ID.test('G-123456')).toBe(false);  // old format
    expect(VALIDATION_PATTERNS.GRIEVANCE_ID.test('GJOSM1234')).toBe(false); // 4 digits
  });

  test('PHONE_US regex is defined', () => {
    expect(VALIDATION_PATTERNS.PHONE_US).toBeDefined();
  });
});

// ============================================================================
// Constants - VALIDATION_MESSAGES
// ============================================================================

describe('VALIDATION_MESSAGES', () => {
  test('has EMAIL_INVALID message', () => {
    expect(typeof VALIDATION_MESSAGES.EMAIL_INVALID).toBe('string');
    expect(VALIDATION_MESSAGES.EMAIL_INVALID.length).toBeGreaterThan(0);
  });

  test('has EMAIL_EMPTY message', () => {
    expect(typeof VALIDATION_MESSAGES.EMAIL_EMPTY).toBe('string');
  });

  test('has PHONE_INVALID message', () => {
    expect(typeof VALIDATION_MESSAGES.PHONE_INVALID).toBe('string');
  });

  test('has MEMBER_ID_INVALID message', () => {
    expect(typeof VALIDATION_MESSAGES.MEMBER_ID_INVALID).toBe('string');
  });

  test('has GRIEVANCE_ID_INVALID message', () => {
    expect(typeof VALIDATION_MESSAGES.GRIEVANCE_ID_INVALID).toBe('string');
  });
});

// ============================================================================
// Assert object
// ============================================================================

describe('Assert object', () => {
  test('Assert is defined', () => {
    expect(Assert).toBeDefined();
  });

  test('Assert has isTrue method', () => {
    expect(typeof Assert.isTrue).toBe('function');
  });

  test('Assert has isFalse method', () => {
    expect(typeof Assert.isFalse).toBe('function');
  });

  test('Assert has equals method', () => {
    expect(typeof Assert.equals).toBe('function');
  });

  test('Assert has isDefined method', () => {
    expect(typeof Assert.isDefined).toBe('function');
  });
});

// ============================================================================
// Validation - validateEmailAddress
// ============================================================================

describe('validateEmailAddress', () => {
  test('valid email returns valid: true', () => {
    const result = validateEmailAddress('test@example.com');
    expect(result.valid).toBe(true);
  });

  test('empty email returns valid: false with empty message', () => {
    const result = validateEmailAddress('');
    expect(result.valid).toBe(false);
    expect(result.message).toBe(VALIDATION_MESSAGES.EMAIL_EMPTY);
  });

  test('null email returns valid: false', () => {
    const result = validateEmailAddress(null);
    expect(result.valid).toBe(false);
  });

  test('invalid email returns valid: false with invalid message', () => {
    const result = validateEmailAddress('not-an-email');
    expect(result.valid).toBe(false);
    expect(result.message).toBe(VALIDATION_MESSAGES.EMAIL_INVALID);
  });

  test('detects gmial.com typo and suggests gmail.com', () => {
    const result = validateEmailAddress('user@gmial.com');
    expect(result.valid).toBe(true);
    expect(result.suggestion).toBe('user@gmail.com');
  });

  test('trims and lowercases email before validation', () => {
    const result = validateEmailAddress('  User@Example.COM  ');
    expect(result.valid).toBe(true);
  });
});

// ============================================================================
// Validation - validatePhoneNumber
// ============================================================================

describe('validatePhoneNumber', () => {
  test('valid 10-digit phone returns valid: true', () => {
    const result = validatePhoneNumber('5551234567');
    expect(result.valid).toBe(true);
    expect(result.formatted).toBeDefined();
  });

  test('formatted phone with parens/dashes returns valid: true', () => {
    const result = validatePhoneNumber('(555) 123-4567');
    expect(result.valid).toBe(true);
  });

  test('empty phone returns valid: false', () => {
    const result = validatePhoneNumber('');
    expect(result.valid).toBe(false);
  });

  test('too short phone returns valid: false', () => {
    const result = validatePhoneNumber('123');
    expect(result.valid).toBe(false);
  });

  test('phone with too many digits returns valid: false', () => {
    const result = validatePhoneNumber('1234567890123456');
    expect(result.valid).toBe(false);
  });

  test('null phone returns valid: false', () => {
    const result = validatePhoneNumber(null);
    expect(result.valid).toBe(false);
  });
});

// ============================================================================
// Validation - formatUSPhone
// ============================================================================

describe('formatUSPhone', () => {
  test('formats 10-digit string as (XXX) XXX-XXXX', () => {
    expect(formatUSPhone('5551234567')).toBe('(555) 123-4567');
  });

  test('strips leading 1 from 11-digit number and formats', () => {
    expect(formatUSPhone('15551234567')).toBe('(555) 123-4567');
  });

  test('returns digits as-is for non-10/11 digit input', () => {
    expect(formatUSPhone('12345')).toBe('12345');
  });
});

// ============================================================================
// Validation - validateRequired
// ============================================================================

describe('validateRequired', () => {
  test('throws Error for null value', () => {
    expect(() => validateRequired(null, 'Name')).toThrow('Name is required');
  });

  test('throws Error for undefined value', () => {
    expect(() => validateRequired(undefined, 'Email')).toThrow('Email is required');
  });

  test('throws Error for empty string', () => {
    expect(() => validateRequired('', 'Phone')).toThrow('Phone is required');
  });

  test('returns value when valid', () => {
    expect(validateRequired('hello', 'Field')).toBe('hello');
  });

  test('returns numeric zero (0 is falsy but not null/undefined/empty)', () => {
    expect(validateRequired(0, 'Count')).toBe(0);
  });
});

// ============================================================================
// Test Helper - isValidEmail_
// ============================================================================

describe('isValidEmail_', () => {
  test('returns true for valid email', () => {
    expect(isValidEmail_('user@example.com')).toBe(true);
  });

  test('returns false for invalid email', () => {
    expect(isValidEmail_('not-email')).toBe(false);
  });

  test('returns false for null', () => {
    expect(isValidEmail_(null)).toBe(false);
  });

  test('returns false for non-string', () => {
    expect(isValidEmail_(12345)).toBe(false);
  });
});

// ============================================================================
// Test Helper - isValidPhone_
// ============================================================================

describe('isValidPhone_', () => {
  test('returns true for 10-digit phone', () => {
    expect(isValidPhone_('5551234567')).toBe(true);
  });

  test('returns true for 11-digit phone', () => {
    expect(isValidPhone_('15551234567')).toBe(true);
  });

  test('returns false for short phone', () => {
    expect(isValidPhone_('12345')).toBe(false);
  });

  test('returns false for null', () => {
    expect(isValidPhone_(null)).toBe(false);
  });

  test('returns false for non-string', () => {
    expect(isValidPhone_(555)).toBe(false);
  });
});

// ============================================================================
// Test Helper - formatDate_
// ============================================================================

describe('formatDate_', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('formats a Date object using Utilities.formatDate', () => {
    const date = new Date(2023, 5, 15);
    const result = formatDate_(date);
    expect(Utilities.formatDate).toHaveBeenCalled();
    expect(typeof result).toBe('string');
  });

  test('returns empty string for null', () => {
    expect(formatDate_(null)).toBe('');
  });

  test('returns empty string for non-Date', () => {
    expect(formatDate_('not-a-date')).toBe('');
  });
});

// ============================================================================
// Test Helper - trimString_
// ============================================================================

describe('trimString_', () => {
  test('trims whitespace from string', () => {
    expect(trimString_('  hello  ')).toBe('hello');
  });

  test('returns empty string for null', () => {
    expect(trimString_(null)).toBe('');
  });

  test('returns empty string for undefined', () => {
    expect(trimString_(undefined)).toBe('');
  });

  test('returns empty string for non-string', () => {
    expect(trimString_(12345)).toBe('');
  });

  test('returns unchanged string when no trimming needed', () => {
    expect(trimString_('hello')).toBe('hello');
  });
});

// ============================================================================
// Test Helper - isValidMemberId_
// ============================================================================

describe('isValidMemberId_', () => {
  test('returns true for valid MBR- format', () => {
    expect(isValidMemberId_('MBR-001')).toBe(true);
    expect(isValidMemberId_('MBR-12345')).toBe(true);
  });

  test('returns false for invalid format', () => {
    expect(isValidMemberId_('MEMBER-001')).toBe(false);
    expect(isValidMemberId_('MBR001')).toBe(false);
  });

  test('returns false for null', () => {
    expect(isValidMemberId_(null)).toBe(false);
  });

  test('returns false for non-string', () => {
    expect(isValidMemberId_(123)).toBe(false);
  });
});

// ============================================================================
// Test Helper - isValidGrievanceId_
// ============================================================================

describe('isValidGrievanceId_', () => {
  test('returns true for valid GRV-YYYY-N format', () => {
    expect(isValidGrievanceId_('GRV-2023-001')).toBe(true);
    expect(isValidGrievanceId_('GRV-2024-1')).toBe(true);
  });

  test('returns false for invalid format', () => {
    expect(isValidGrievanceId_('GRV-23-001')).toBe(false);
    expect(isValidGrievanceId_('GRIEV-2023-001')).toBe(false);
  });

  test('returns false for null', () => {
    expect(isValidGrievanceId_(null)).toBe(false);
  });

  test('returns false for non-string', () => {
    expect(isValidGrievanceId_(123)).toBe(false);
  });
});

// ============================================================================
// Test Helper - getUniqueValues_
// ============================================================================

describe('getUniqueValues_', () => {
  test('removes duplicates', () => {
    expect(getUniqueValues_([1, 2, 2, 3, 3, 3])).toEqual([1, 2, 3]);
  });

  test('preserves order of first occurrence', () => {
    expect(getUniqueValues_(['c', 'a', 'b', 'a', 'c'])).toEqual(['c', 'a', 'b']);
  });

  test('returns empty array for empty input', () => {
    expect(getUniqueValues_([])).toEqual([]);
  });

  test('handles array with all unique elements', () => {
    expect(getUniqueValues_([1, 2, 3])).toEqual([1, 2, 3]);
  });
});

// ============================================================================
// Test Helper - flattenArray_
// ============================================================================

describe('flattenArray_', () => {
  test('flattens one level of nesting', () => {
    expect(flattenArray_([1, [2, 3], 4])).toEqual([1, 2, 3, 4]);
  });

  test('flattens deeply nested arrays', () => {
    expect(flattenArray_([1, [2, [3, [4]]]])).toEqual([1, 2, 3, 4]);
  });

  test('returns empty array for empty input', () => {
    expect(flattenArray_([])).toEqual([]);
  });

  test('returns same elements for already flat array', () => {
    expect(flattenArray_([1, 2, 3])).toEqual([1, 2, 3]);
  });
});

// ============================================================================
// Test Helper - hasProperty_
// ============================================================================

describe('hasProperty_', () => {
  test('returns true for own property', () => {
    expect(hasProperty_({ name: 'test' }, 'name')).toBe(true);
  });

  test('returns false for non-existent property', () => {
    expect(hasProperty_({ name: 'test' }, 'age')).toBe(false);
  });

  test('returns false for inherited property', () => {
    const obj = Object.create({ inherited: true });
    expect(hasProperty_(obj, 'inherited')).toBe(false);
  });

  test('returns falsy for null object', () => {
    expect(hasProperty_(null, 'prop')).toBeFalsy();
  });

  test('returns falsy for undefined object', () => {
    expect(hasProperty_(undefined, 'prop')).toBeFalsy();
  });
});
