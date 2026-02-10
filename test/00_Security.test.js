/**
 * Tests for 00_Security.gs
 *
 * Covers XSS prevention, formula injection, PII masking, input validation,
 * and client-side security helper generation.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

loadSources(['00_Security.gs']);

// ============================================================================
// escapeHtml
// ============================================================================

describe('escapeHtml', () => {
  test('escapes HTML special characters', () => {
    expect(escapeHtml('<script>alert("xss")</script>')).toBe(
      '&lt;script&gt;alert(&quot;xss&quot;)&lt;&#x2F;script&gt;'
    );
  });

  test('escapes single quotes', () => {
    expect(escapeHtml("it's")).toBe("it&#x27;s");
  });

  test('escapes backticks', () => {
    expect(escapeHtml('`code`')).toBe('&#x60;code&#x60;');
  });

  test('escapes equals sign', () => {
    expect(escapeHtml('a=b')).toBe('a&#x3D;b');
  });

  test('escapes forward slash', () => {
    expect(escapeHtml('a/b')).toBe('a&#x2F;b');
  });

  test('escapes ampersand', () => {
    expect(escapeHtml('a&b')).toBe('a&amp;b');
  });

  test('returns empty string for null', () => {
    expect(escapeHtml(null)).toBe('');
  });

  test('returns empty string for undefined', () => {
    expect(escapeHtml(undefined)).toBe('');
  });

  test('converts non-string input to string', () => {
    expect(escapeHtml(123)).toBe('123');
  });

  test('handles empty string', () => {
    expect(escapeHtml('')).toBe('');
  });

  test('handles combined attack vectors', () => {
    const input = '<img src=x onerror="alert(1)">';
    const result = escapeHtml(input);
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result).not.toContain('"');
  });
});

// ============================================================================
// sanitizeForHtml (alias)
// ============================================================================

describe('sanitizeForHtml', () => {
  test('delegates to escapeHtml', () => {
    expect(sanitizeForHtml('<b>bold</b>')).toBe(escapeHtml('<b>bold</b>'));
  });
});

// ============================================================================
// sanitizeObjectForHtml
// ============================================================================

describe('sanitizeObjectForHtml', () => {
  test('escapes string values in object', () => {
    const result = sanitizeObjectForHtml({ name: '<script>', age: 30 });
    expect(result.name).toBe(escapeHtml('<script>'));
    expect(result.age).toBe(30);
  });

  test('recursively sanitizes nested objects', () => {
    const result = sanitizeObjectForHtml({ data: { name: '<b>test</b>' } });
    expect(result.data.name).toBe(escapeHtml('<b>test</b>'));
  });

  test('returns non-object input as-is', () => {
    expect(sanitizeObjectForHtml(null)).toBe(null);
    expect(sanitizeObjectForHtml(42)).toBe(42);
  });
});

// ============================================================================
// escapeForFormula
// ============================================================================

describe('escapeForFormula', () => {
  test('returns empty string for null', () => {
    expect(escapeForFormula(null)).toBe('');
  });

  test('prefixes formula-starting characters at start of string', () => {
    expect(escapeForFormula('=SUM(A1)')).toBe("'=SUM(A1)");
    expect(escapeForFormula('+cmd')).toBe("'+cmd");
    expect(escapeForFormula('-cmd')).toBe("'-cmd");
    expect(escapeForFormula('@import')).toBe("'@import");
  });

  test('does NOT prefix formula characters mid-string (bug fix verification)', () => {
    expect(escapeForFormula('user@email.com')).toBe('user@email.com');
    expect(escapeForFormula('a+b')).toBe('a+b');
    expect(escapeForFormula('a-b')).toBe('a-b');
  });

  test('escapes single quotes', () => {
    expect(escapeForFormula("it's")).toBe("it''s");
  });

  test('escapes double quotes', () => {
    expect(escapeForFormula('say "hi"')).toBe('say ""hi""');
  });

  test('replaces newlines with spaces', () => {
    expect(escapeForFormula('line1\nline2')).toBe('line1 line2');
  });

  test('handles normal text without modification', () => {
    expect(escapeForFormula('Member Directory')).toBe('Member Directory');
  });
});

// ============================================================================
// safeSheetNameForFormula
// ============================================================================

describe('safeSheetNameForFormula', () => {
  test('returns empty string for empty input', () => {
    expect(safeSheetNameForFormula('')).toBe('');
  });

  test('wraps names with special characters in quotes', () => {
    expect(safeSheetNameForFormula('Member Directory')).toBe("'Member Directory'");
  });

  test('does not wrap simple names', () => {
    expect(safeSheetNameForFormula('Config')).toBe('Config');
  });
});

// ============================================================================
// PII Masking
// ============================================================================

describe('maskEmail', () => {
  test('masks email correctly', () => {
    expect(maskEmail('john@example.com')).toBe('j***n@example.com');
  });

  test('returns placeholder for non-string', () => {
    expect(maskEmail(null)).toBe('[no email]');
  });
});

describe('maskPhone', () => {
  test('masks phone number', () => {
    expect(maskPhone('555-123-4567')).toBe('***-***-4567');
  });

  test('returns placeholder for non-string', () => {
    expect(maskPhone(null)).toBe('[no phone]');
  });
});

describe('maskName', () => {
  test('masks name to initials', () => {
    expect(maskName('John', 'Smith')).toBe('J. S.');
  });

  test('handles missing parts', () => {
    expect(maskName('', '')).toBe('[anonymous]');
  });
});

// ============================================================================
// Input Validation
// ============================================================================

describe('isValidSafeString', () => {
  test('accepts normal strings', () => {
    expect(isValidSafeString('Hello World')).toBe(true);
  });

  test('rejects script tags', () => {
    expect(isValidSafeString('<script>alert(1)</script>')).toBe(false);
  });

  test('rejects event handlers', () => {
    expect(isValidSafeString('onerror=alert(1)')).toBe(false);
  });

  test('rejects strings over max length', () => {
    expect(isValidSafeString('a'.repeat(1001))).toBe(false);
  });

  test('accepts null/undefined', () => {
    expect(isValidSafeString(null)).toBe(true);
  });
});

describe('isValidMemberId', () => {
  test('accepts valid member IDs', () => {
    expect(isValidMemberId('MEM-abc123')).toBe(true);
  });

  test('rejects empty/null', () => {
    expect(isValidMemberId('')).toBe(false);
    expect(isValidMemberId(null)).toBe(false);
  });

  test('rejects IDs with special characters', () => {
    expect(isValidMemberId('MEM <script>')).toBe(false);
  });
});

describe('isValidGrievanceId', () => {
  test('accepts valid grievance IDs', () => {
    expect(isValidGrievanceId('GRV-001')).toBe(true);
  });

  test('rejects invalid IDs', () => {
    expect(isValidGrievanceId('')).toBe(false);
  });
});

// ============================================================================
// Client-Side Security Helpers
// ============================================================================

describe('getClientSideEscapeHtml', () => {
  test('returns JavaScript code with all 8 escape replacements', () => {
    const code = getClientSideEscapeHtml();
    expect(code).toContain('function escapeHtml');
    expect(code).toContain('&amp;');
    expect(code).toContain('&lt;');
    expect(code).toContain('&gt;');
    expect(code).toContain('&quot;');
    expect(code).toContain('&#x27;');
    expect(code).toContain('&#x2F;');
    expect(code).toContain('&#x60;');
    expect(code).toContain('&#x3D;');
  });
});

describe('getClientSecurityScript', () => {
  test('includes safeAttr that delegates to escapeHtml (bug fix verification)', () => {
    const result = getClientSecurityScript();
    expect(result).toContain('function safeAttr(t)');
    expect(result).toContain('escapeHtml(t)');
  });
});

describe('safeJsonForHtml', () => {
  test('returns empty object for null', () => {
    expect(safeJsonForHtml(null)).toBe('{}');
  });

  test('escapes HTML in string values', () => {
    const result = safeJsonForHtml({ name: '<script>' });
    expect(result).not.toContain('<script>');
  });

  test('escapes closing script tags', () => {
    const result = safeJsonForHtml({ html: '</script>' });
    expect(result).not.toContain('</script>');
  });
});

describe('sanitizeDataForClient', () => {
  test('sanitizes specified fields only', () => {
    const data = [{ name: '<b>John</b>', age: 30 }];
    const result = sanitizeDataForClient(data, ['name']);
    expect(result[0].name).toBe(escapeHtml('<b>John</b>'));
    expect(result[0].age).toBe(30);
  });

  test('returns empty array for non-array input', () => {
    expect(sanitizeDataForClient(null)).toEqual([]);
  });
});

// ============================================================================
// ACCESS_CONTROL
// ============================================================================

describe('ACCESS_CONTROL', () => {
  test('has required mode and page lists', () => {
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('steward');
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('member');
    expect(ACCESS_CONTROL.ALLOWED_MODES).toContain('selfservice');
    expect(ACCESS_CONTROL.ALLOWED_PAGES).toContain('dashboard');
    expect(ACCESS_CONTROL.ALLOWED_PAGES).toContain('portal');
  });
});
