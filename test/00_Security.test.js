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
      '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
    );
  });

  test('escapes single quotes', () => {
    expect(escapeHtml("it's")).toBe("it&#x27;s");
  });

  test('escapes backticks', () => {
    expect(escapeHtml('`code`')).toBe('&#x60;code&#x60;');
  });

  test('does not escape equals sign (not an XSS vector)', () => {
    expect(escapeHtml('a=b')).toBe('a=b');
  });

  test('does not escape forward slash (not an XSS vector)', () => {
    expect(escapeHtml('a/b')).toBe('a/b');
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

  test('escapes input with all special chars combined', () => {
    const input = '<div class="test">&\'`';
    const result = escapeHtml(input);
    // Raw dangerous chars should not appear unescaped
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result).not.toContain('"');
    expect(result).not.toContain("'");
    expect(result).not.toContain('`');
    // Verify each entity is present in the output
    expect(result).toContain('&lt;');
    expect(result).toContain('&gt;');
    expect(result).toContain('&quot;');
    expect(result).toContain('&amp;');
    expect(result).toContain('&#x27;');
    expect(result).toContain('&#x60;');
  });

  test('handles very long input string', () => {
    const longStr = '<script>'.repeat(1000);
    const result = escapeHtml(longStr);
    expect(result).not.toContain('<');
    expect(result).not.toContain('>');
    expect(result.length).toBeGreaterThan(longStr.length);
  });

  test('handles input containing only special chars', () => {
    const input = '<>&"\'`';
    const result = escapeHtml(input);
    expect(result).toBe('&lt;&gt;&amp;&quot;&#x27;&#x60;');
  });
});

// sanitizeObjectForHtml — removed (zero callers in production code)

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

  test('preserves single quotes (no SQL-style doubling)', () => {
    expect(escapeForFormula("it's")).toBe("it's");
  });

  test('preserves double quotes (no SQL-style doubling)', () => {
    expect(escapeForFormula('say "hi"')).toBe('say "hi"');
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

  test('rejects null/undefined', () => {
    expect(isValidSafeString(null)).toBe(false);
    expect(isValidSafeString(undefined)).toBe(false);
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
  test('returns JavaScript code with all 6 escape replacements', () => {
    const code = getClientSideEscapeHtml();
    expect(code).toContain('function escapeHtml');
    expect(code).toContain('&amp;');
    expect(code).toContain('&lt;');
    expect(code).toContain('&gt;');
    expect(code).toContain('&quot;');
    expect(code).toContain('&#x27;');
    expect(code).toContain('&#x60;');
    // / and = are NOT escaped (F107: not XSS vectors, caused data corruption)
    expect(code).not.toContain('&#x2F;');
    expect(code).not.toContain('&#x3D;');
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

// ============================================================================
// SECURITY_SEVERITY
// ============================================================================

describe('SECURITY_SEVERITY', () => {
  test('defines all severity levels', () => {
    expect(SECURITY_SEVERITY.CRITICAL).toBe('CRITICAL');
    expect(SECURITY_SEVERITY.HIGH).toBe('HIGH');
    expect(SECURITY_SEVERITY.MEDIUM).toBe('MEDIUM');
    expect(SECURITY_SEVERITY.LOW).toBe('LOW');
  });
});

// ============================================================================
// recordSecurityEvent
// ============================================================================

describe('recordSecurityEvent', () => {
  beforeEach(() => {
    global.logAuditEvent = jest.fn();
  });

  test('is a function', () => {
    expect(typeof recordSecurityEvent).toBe('function');
  });

  test('logs audit event for any severity', () => {
    recordSecurityEvent('TEST_EVENT', SECURITY_SEVERITY.LOW, 'Test description', { foo: 'bar' });
    expect(logAuditEvent).toHaveBeenCalledWith('SECURITY_TEST_EVENT', expect.objectContaining({
      _severity: 'LOW',
      _description: 'Test description',
      foo: 'bar'
    }));
  });

  test('calls sendSecurityAlertEmail_ for CRITICAL events', () => {
    // MailApp.sendEmail is mocked — just verify it doesn't throw
    expect(() => {
      recordSecurityEvent('CRITICAL_TEST', SECURITY_SEVERITY.CRITICAL, 'Critical event', {});
    }).not.toThrow();
  });

  test('queues HIGH events for digest', () => {
    recordSecurityEvent('HIGH_TEST', SECURITY_SEVERITY.HIGH, 'High event', { detail: 'x' });

    // Verify event was queued in properties
    const props = PropertiesService.getScriptProperties();
    const queue = JSON.parse(props.getProperty('SECURITY_DIGEST_QUEUE') || '[]');
    expect(queue.length).toBeGreaterThan(0);
    expect(queue[queue.length - 1].event).toBe('HIGH_TEST');
  });

  test('does not queue LOW or MEDIUM events for digest', () => {
    // Clear any existing queue
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SECURITY_DIGEST_QUEUE', '[]');

    recordSecurityEvent('LOW_TEST', SECURITY_SEVERITY.LOW, 'Low event', {});
    recordSecurityEvent('MED_TEST', SECURITY_SEVERITY.MEDIUM, 'Medium event', {});

    const queue = JSON.parse(props.getProperty('SECURITY_DIGEST_QUEUE') || '[]');
    expect(queue.length).toBe(0);
  });
});

// ============================================================================
// queueSecurityDigestEvent_ and sendDailySecurityDigest
// ============================================================================

describe('sendDailySecurityDigest', () => {
  test('is a function', () => {
    expect(typeof sendDailySecurityDigest).toBe('function');
  });

  test('does not throw when queue is empty', () => {
    expect(() => sendDailySecurityDigest()).not.toThrow();
  });

  test('does not send email when queue is empty', () => {
    MailApp.sendEmail.mockClear();
    sendDailySecurityDigest();
    expect(MailApp.sendEmail).not.toHaveBeenCalled();
  });
});

// ============================================================================
// escapeForFormulaPreserveNewlines
// ============================================================================

describe('escapeForFormulaPreserveNewlines', () => {
  test('preserves newlines in multiline text', () => {
    expect(escapeForFormulaPreserveNewlines('line 1\nline 2\nline 3'))
      .toBe('line 1\nline 2\nline 3');
  });

  test('escapes formula chars at start of each line', () => {
    expect(escapeForFormulaPreserveNewlines('=SUM(A1)\n+B2\n-C3\n@D4'))
      .toBe("'=SUM(A1)\n'+B2\n'-C3\n'@D4");
  });

  test('escapes formula chars after leading whitespace on each line', () => {
    expect(escapeForFormulaPreserveNewlines('  =IMPORTRANGE()\n  +123'))
      .toBe("'  =IMPORTRANGE()\n'  +123");
  });

  test('strips carriage returns but preserves line feeds', () => {
    expect(escapeForFormulaPreserveNewlines('line 1\r\nline 2\rline 3'))
      .toBe('line 1\nline 2\nline 3');
  });

  test('replaces tabs with spaces per line', () => {
    expect(escapeForFormulaPreserveNewlines('col1\tcol2\ncol3\tcol4'))
      .toBe('col1 col2\ncol3 col4');
  });

  test('escapes backslashes per line', () => {
    expect(escapeForFormulaPreserveNewlines('path\\to\\file\nsecond\\line'))
      .toBe('path\\\\to\\\\file\nsecond\\\\line');
  });

  test('returns empty string for null/undefined', () => {
    expect(escapeForFormulaPreserveNewlines(null)).toBe('');
    expect(escapeForFormulaPreserveNewlines(undefined)).toBe('');
  });

  test('handles empty string', () => {
    expect(escapeForFormulaPreserveNewlines('')).toBe('');
  });

  test('handles single line without newlines', () => {
    expect(escapeForFormulaPreserveNewlines('Normal text')).toBe('Normal text');
  });

  test('handles bullet-formatted text (real-world use case)', () => {
    var input = '• Budget approved\n• 15 stewards attended\n• Next meeting May 4';
    expect(escapeForFormulaPreserveNewlines(input)).toBe(input);
  });
});

