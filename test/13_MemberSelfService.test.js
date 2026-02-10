/**
 * Tests for 13_MemberSelfService.gs
 *
 * Covers PIN generation, PIN hashing, reset tokens,
 * and configuration constants.
 */

require('./gas-mock');
const { loadSources } = require('./load-source');

// Load dependencies then self-service module
loadSources(['00_Security.gs', '00_DataAccess.gs', '01_Core.gs', '13_MemberSelfService.gs']);

// ============================================================================
// PIN_CONFIG
// ============================================================================

describe('PIN_CONFIG', () => {
  test('PIN_LENGTH is 6', () => {
    expect(PIN_CONFIG.PIN_LENGTH).toBe(6);
  });

  test('MAX_ATTEMPTS is 5', () => {
    expect(PIN_CONFIG.MAX_ATTEMPTS).toBe(5);
  });

  test('LOCKOUT_MINUTES is 15', () => {
    expect(PIN_CONFIG.LOCKOUT_MINUTES).toBe(15);
  });

  test('PIN_COLUMN matches MEMBER_COLS.PIN_HASH', () => {
    expect(PIN_CONFIG.PIN_COLUMN).toBe(MEMBER_COLS.PIN_HASH);
  });

  test('RESET_TOKEN_EXPIRY_MINUTES is 30', () => {
    expect(PIN_CONFIG.RESET_TOKEN_EXPIRY_MINUTES).toBe(30);
  });
});

// ============================================================================
// generateMemberPIN
// ============================================================================

describe('generateMemberPIN', () => {
  test('generates a string', () => {
    const pin = generateMemberPIN();
    expect(typeof pin).toBe('string');
  });

  test('has correct length (6 digits)', () => {
    const pin = generateMemberPIN();
    expect(pin.length).toBe(6);
  });

  test('uses Utilities.getUuid() (not Math.random)', () => {
    generateMemberPIN();
    expect(Utilities.getUuid).toHaveBeenCalled();
  });

  test('only contains digits', () => {
    // Mock UUID with known value that has digits
    Utilities.getUuid.mockReturnValue('12345678-abcd-efgh-ijkl-mnopqrstuvwx');
    const pin = generateMemberPIN();
    expect(pin).toMatch(/^\d{6}$/);
  });

  test('pads with zeros if not enough digits', () => {
    // UUID with fewer than 6 digits
    Utilities.getUuid.mockReturnValue('abcde1fg-hijk-lmno-pqrs-tuvwxyz00000');
    const pin = generateMemberPIN();
    expect(pin.length).toBe(6);
  });
});

// ============================================================================
// hashPIN
// ============================================================================

describe('hashPIN', () => {
  test('returns empty string for missing pin', () => {
    expect(hashPIN('', 'MEM-001')).toBe('');
    expect(hashPIN(null, 'MEM-001')).toBe('');
  });

  test('returns empty string for missing memberId', () => {
    expect(hashPIN('123456', '')).toBe('');
    expect(hashPIN('123456', null)).toBe('');
  });

  test('returns a hex string', () => {
    const hash = hashPIN('123456', 'MEM-001');
    expect(hash).toMatch(/^[0-9a-f]+$/);
  });

  test('produces consistent output for same inputs', () => {
    const hash1 = hashPIN('123456', 'MEM-001');
    const hash2 = hashPIN('123456', 'MEM-001');
    expect(hash1).toBe(hash2);
  });

  test('produces different output for different PINs', () => {
    const hash1 = hashPIN('123456', 'MEM-001');
    const hash2 = hashPIN('654321', 'MEM-001');
    expect(hash1).not.toBe(hash2);
  });

  test('produces different output for different member IDs', () => {
    const hash1 = hashPIN('123456', 'MEM-001');
    const hash2 = hashPIN('123456', 'MEM-002');
    expect(hash1).not.toBe(hash2);
  });

  test('uses Utilities.computeDigest with SHA_256', () => {
    hashPIN('123456', 'MEM-001');
    expect(Utilities.computeDigest).toHaveBeenCalledWith(
      'SHA_256',
      expect.any(String),
      'UTF_8'
    );
  });
});

// ============================================================================
// verifyPIN
// ============================================================================

describe('verifyPIN', () => {
  test('returns true for matching PIN', () => {
    const hash = hashPIN('123456', 'MEM-001');
    expect(verifyPIN('123456', 'MEM-001', hash)).toBe(true);
  });

  test('returns false for wrong PIN', () => {
    const hash = hashPIN('123456', 'MEM-001');
    expect(verifyPIN('000000', 'MEM-001', hash)).toBe(false);
  });

  test('returns false for missing inputs', () => {
    expect(verifyPIN('', 'MEM-001', 'hash')).toBe(false);
    expect(verifyPIN('123456', '', 'hash')).toBe(false);
    expect(verifyPIN('123456', 'MEM-001', '')).toBe(false);
    expect(verifyPIN(null, null, null)).toBe(false);
  });
});

// ============================================================================
// MEMBER_PIN_COLS
// ============================================================================

describe('MEMBER_PIN_COLS', () => {
  test('PIN_HASH matches PIN_CONFIG.PIN_COLUMN', () => {
    expect(MEMBER_PIN_COLS.PIN_HASH).toBe(PIN_CONFIG.PIN_COLUMN);
  });

  test('PIN_HASH matches MEMBER_COLS.PIN_HASH', () => {
    expect(MEMBER_PIN_COLS.PIN_HASH).toBe(MEMBER_COLS.PIN_HASH);
  });
});

// ============================================================================
// completePINReset validation
// ============================================================================

describe('completePINReset', () => {
  test('rejects missing inputs', () => {
    const result = completePINReset('', '', '');
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });

  test('rejects invalid PIN format (not 6 digits)', () => {
    const result = completePINReset('MEM-001', 'ABCD1234', '12345');
    expect(result.success).toBe(false);
    expect(result.error).toContain('6 digits');
  });

  test('rejects non-numeric PIN', () => {
    const result = completePINReset('MEM-001', 'ABCD1234', 'abcdef');
    expect(result.success).toBe(false);
  });
});

// ============================================================================
// requestPINReset
// ============================================================================

describe('requestPINReset', () => {
  test('rejects empty memberId', () => {
    const result = requestPINReset('');
    expect(result.success).toBe(false);
    expect(result.error).toBeTruthy();
  });
});
