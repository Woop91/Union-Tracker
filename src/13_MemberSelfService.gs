/**
 * ============================================================================
 * 13_MemberSelfService.gs - MEMBER SELF-SERVICE PORTAL
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Member Self-Service Portal with PIN authentication. Members authenticate
 *   with a 6-digit PIN (SHA-256 hashed before storage) to view their own
 *   info, update contact details, and check grievance status. Features: PIN
 *   generation, rate limiting (5 failed attempts per 15 minutes), audit
 *   logging of all access, and members can ONLY see their own data.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   PIN-based auth was chosen over Google SSO for the self-service portal
 *   because many union members use shared workplace computers and may not
 *   have Google accounts. PINs are simpler for non-technical users. SHA-256
 *   hashing ensures even if the sheet is compromised, PINs can't be read.
 *   Rate limiting prevents brute-force attacks. validateSelfServiceInput_()
 *   validates all input fields using isValidSafeString() to prevent injection.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Members can't log in to the self-service portal. PIN verification fails.
 *   Contact info updates from members stop working. If the hashing is broken,
 *   PINs are stored in plaintext (security vulnerability). If rate limiting
 *   breaks, brute-force PIN guessing becomes possible.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS, MEMBER_COLS),
 *   00_Security.gs (isValidSafeString, secureLog).
 *   Used by 14_MeetingCheckIn.gs (authenticateMember, verifyPIN, hashPIN),
 *   the SPA self-service view, and menu items.
 *
 * @version 4.51.0
 * @license Free for use by non-profit collective bargaining groups and unions
 * ============================================================================
 */

/**
 * Validates a self-service input field using isValidSafeString and length checks.
 * @param {string} field - Field name for error messages
 * @param {*} value - Value to validate
 * @param {number} [maxLength=500] - Maximum allowed length
 * @returns {{ valid: boolean, error?: string }}
 * @private
 */
function validateSelfServiceInput_(field, value, maxLength) {
  maxLength = maxLength || 500;
  if (value === null || value === undefined || value === '') return { valid: true };
  var str = String(value);
  if (str.length > maxLength) {
    return { valid: false, error: field + ' exceeds maximum length (' + maxLength + ')' };
  }
  if (typeof isValidSafeString === 'function' && !isValidSafeString(str, maxLength)) {
    return { valid: false, error: field + ' contains disallowed content' };
  }
  return { valid: true };
}

// ============================================================================
// PIN SYSTEM CONSTANTS
// ============================================================================

/**
 * PIN system configuration
 * @const {Object}
 */
var PIN_CONFIG = {
  PIN_LENGTH: 6,                    // 6-digit PIN
  MAX_ATTEMPTS: 5,                  // Max failed attempts before lockout
  LOCKOUT_MINUTES: 15,              // Lockout duration in minutes
  SESSION_DURATION_MINUTES: 30,     // Session duration before re-auth required
  get PIN_COLUMN() { return (typeof MEMBER_COLS !== 'undefined' && MEMBER_COLS.PIN_HASH) ? MEMBER_COLS.PIN_HASH : 34; },
  SALT_PROPERTY: 'MEMBER_PIN_SALT', // Property key for salt storage
  RESET_TOKEN_EXPIRY_MINUTES: 30,   // Reset token expiration time
  RESET_TOKEN_PREFIX: 'pin_reset_'  // Cache key prefix for reset tokens
};

/**
 * Member self-service column mapping
 * Column AH (34) stores the hashed PIN
 * @const {Object}
 */
var MEMBER_PIN_COLS = {
  get PIN_HASH() { return (typeof MEMBER_COLS !== 'undefined' && MEMBER_COLS.PIN_HASH) ? MEMBER_COLS.PIN_HASH : 34; }
};

// ============================================================================
// PIN GENERATION AND HASHING
// ============================================================================

/**
 * Generate a random 6-digit PIN using Utilities.getUuid() for
 * cryptographic-quality randomness (avoids Math.random() weakness).
 * Leading zeros are valid (e.g. '012345').
 * @returns {string} 6-digit PIN
 */
function generateMemberPIN() {
  var pin = '';
  while (pin.length < 6) {
    var hex = Utilities.getUuid().replace(/-/g, '');
    for (var j = 0; j + 1 < hex.length && pin.length < 6; j += 2) {
      var byte = parseInt(hex.charAt(j) + hex.charAt(j + 1), 16);
      // Rejection sampling: discard 250-255 so 0-249 maps uniformly to 0-9
      if (byte < 250) {
        pin += (byte % 10).toString();
      }
    }
  }
  return pin;
}

/**
 * Get or create the system salt for PIN hashing
 * @returns {string} Salt value
 * @private
 */
function getPINSalt_() {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty(PIN_CONFIG.SALT_PROPERTY);

  if (!salt) {
    // Generate a new salt if none exists
    salt = Utilities.getUuid() + '-' + Date.now();
    props.setProperty(PIN_CONFIG.SALT_PROPERTY, salt);

    // Verify persistence — if this fails, all PIN hashes will silently mismatch
    var persisted = props.getProperty(PIN_CONFIG.SALT_PROPERTY);
    if (persisted !== salt) {
      throw new Error('CRITICAL: PIN salt failed to persist to ScriptProperties. PIN authentication will not work until this is resolved.');
    }
  }

  return salt;
}

/**
 * Hash a PIN using SHA-256 with salt
 * @param {string} pin - The plaintext PIN
 * @param {string} memberId - The member ID (used as additional salt)
 * @returns {string} Hashed PIN
 */
function hashPIN(pin, memberId) {
  if (!pin || !memberId) return '';

  var salt = getPINSalt_();
  var dataToHash = salt + ':' + memberId + ':' + pin;

  var hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    dataToHash,
    Utilities.Charset.UTF_8
  );

  // Convert to hex string
  return hash.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/**
 * Verify a PIN against stored hash
 * @param {string} pin - The plaintext PIN to verify
 * @param {string} memberId - The member ID
 * @param {string} storedHash - The stored hash to compare against
 * @returns {boolean} True if PIN matches
 */
function verifyPIN(pin, memberId, storedHash) {
  if (!pin || !memberId || !storedHash) return false;

  var computedHash = hashPIN(pin, memberId);

  // Constant-time comparison to prevent timing attacks
  if (computedHash.length !== storedHash.length) return false;
  var result = 0;
  for (var i = 0; i < computedHash.length; i++) {
    result |= computedHash.charCodeAt(i) ^ storedHash.charCodeAt(i);
  }
  return result === 0;
}

/**
 * Assign an initial PIN to a member and optionally email it
 * @param {string} memberId - The member ID
 * @param {Object} [options] - Optional settings
 * @param {boolean} [options.sendEmail] - Whether to email the PIN to the member
 * @returns {Object} Result with success status and generated PIN
 */
function assignMemberPIN(memberId, options) {
  options = options || {};

  // Auth check: only stewards/admins may assign PINs
  var callerRole = typeof getUserRole_ === 'function' ? getUserRole_(Session.getActiveUser().getEmail()) : null;
  if (callerRole !== 'steward' && callerRole !== 'admin' && callerRole !== 'both') {
    return errorResponse('Authorization required: steward or admin access needed', 'assignMemberPIN');
  }

  if (!memberId) {
    return errorResponse('Member ID is required', 'assignMemberPIN');
  }

  memberId = String(memberId).trim().toUpperCase();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return errorResponse('System configuration error. Please contact your steward.', 'assignMemberPIN');
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;
  var memberEmail = null;
  var memberName = null;

  for (var i = 1; i < data.length; i++) {
    var rowId = String(col_(data[i], MEMBER_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowId === memberId) {
      memberRow = i + 1;
      memberEmail = col_(data[i], MEMBER_COLS.EMAIL);
      memberName = ((col_(data[i], MEMBER_COLS.FIRST_NAME) || '') + ' ' + (col_(data[i], MEMBER_COLS.LAST_NAME) || '')).trim();
      break;
    }
  }

  if (memberRow === -1) {
    return errorResponse('Member not found', 'assignMemberPIN');
  }

  // Check if member already has a PIN
  var existingHash = data[memberRow - 1][PIN_CONFIG.PIN_COLUMN - 1];
  if (existingHash) {
    return errorResponse('Member already has a PIN assigned. Use PIN reset instead.', 'assignMemberPIN');
  }

  // Generate random 6-digit PIN
  var newPin = generateMemberPIN();
  var hashedPin = hashPIN(newPin, memberId);
  sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN).setValue(hashedPin);

  // Log the assignment
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('PIN_ASSIGNED', {
      memberId: memberId,
      assignedBy: Session.getActiveUser().getEmail()
    });
  }

  // Optionally email the PIN to the member
  var emailSent = false;
  if (options.sendEmail && memberEmail) {
    try {
      MailApp.sendEmail({
        to: memberEmail,
        subject: COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + 'Your Self-Service Portal PIN',
        body: 'Hello ' + memberName + ',\n\n' +
          'A PIN has been set up for you to access the Member Self-Service Portal.\n\n' +
          'Your PIN is: ' + newPin + '\n\n' +
          'Please change your PIN after your first login for security.\n\n' +
          'If you did not expect this, please contact your union steward.\n\n' +
          'Best regards,\n' + (typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Your Union') : 'Your Union')
      });
      emailSent = true;
    } catch (emailError) {
      log_('Could not send PIN email to member', emailError.message);
    }
  }

  var result = { success: true, memberId: memberId, emailSent: emailSent, message: 'PIN assigned successfully' + (emailSent ? ' and emailed to member' : '') };
  if (!emailSent) {
    result.pin = newPin; // Only return PIN when email delivery failed
    result.pinDeliveryNote = 'Email delivery failed. Provide PIN directly to member.';
  }
  return result;
}

// ============================================================================
// EMAIL-BASED PIN RESET
// ============================================================================

/**
 * Generate a secure reset token
 * @returns {string} Random token string
 * @private
 */
function generateResetToken_() {
  // Use Utilities.getUuid() for better randomness than Math.random()
  var uuid = Utilities.getUuid().replace(/-/g, '').toUpperCase();
  return uuid.substring(0, 16);
}

/**
 * Request a PIN reset - sends reset token via email
 * @param {string} memberId - The member ID
 * @returns {Object} Result with success status
 */
function requestPINReset(memberId) {
  if (!memberId) {
    return errorResponse('Member ID is required', 'requestPINReset');
  }

  memberId = String(memberId).trim().toUpperCase();

  // Rate limit per member ID to prevent email flooding and member enumeration.
  // Returns the same generic success message on rate-limit so attackers cannot
  // detect valid member IDs by timing.
  try {
    var _rateCache = CacheService.getScriptCache();
    var _rateKey = 'PIN_RESET_RATE_' + memberId;
    var _rateCount = parseInt(_rateCache.get(_rateKey) || '0', 10);
    if (_rateCount >= 3) {
      return { success: true, message: 'If your Member ID is valid, a reset email has been sent.' };
    }
    _rateCache.put(_rateKey, String(_rateCount + 1), 900); // 15-minute window
  } catch (_rlErr) { /* non-critical */ }

  // Find member and get their email
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return errorResponse('System configuration error. Please contact your steward.', 'requestPINReset');
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = null;
  var memberEmail = null;
  var memberName = null;

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], MEMBER_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      memberEmail = col_(data[i], MEMBER_COLS.EMAIL);
      memberName = (col_(data[i], MEMBER_COLS.FIRST_NAME) || '') + ' ' + (col_(data[i], MEMBER_COLS.LAST_NAME) || '');
      break;
    }
  }

  if (!memberRow) {
    // Don't reveal if member exists or not for security
    return { success: true, message: 'If your Member ID is valid, a reset email has been sent.' };
  }

  if (!memberEmail) {
    // Log but don't reveal to user
    if (typeof secureLog === 'function') {
      secureLog('PINResetRequest', 'Reset requested but no email on file', { memberId: memberId });
    }
    return { success: true, message: 'If your Member ID is valid, a reset email has been sent.' };
  }

  // Check if member even has a PIN set
  var currentPinHash = data[memberRow - 1][PIN_CONFIG.PIN_COLUMN - 1];
  if (!currentPinHash) {
    // No PIN set - they need to contact a steward
    return { success: true, message: 'If your Member ID is valid, a reset email has been sent.' };
  }

  // Generate reset token and store it (PropertiesService — survives cache eviction)
  var resetToken = generateResetToken_();
  var props = PropertiesService.getScriptProperties();
  var propKey = PIN_CONFIG.RESET_TOKEN_PREFIX + memberId;

  var tokenData = JSON.stringify({
    token: resetToken,
    memberId: memberId,
    created: Date.now(),
    expiresAt: Date.now() + (PIN_CONFIG.RESET_TOKEN_EXPIRY_MINUTES * 60 * 1000)
  });
  props.setProperty(propKey, tokenData);

  // Send reset email
  try {
    var subject = 'Union - PIN Reset Request';
    var body = 'Hello ' + memberName.trim() + ',\n\n' +
      'You requested a PIN reset for the Member Self-Service Portal.\n\n' +
      'Your reset code is: ' + resetToken + '\n\n' +
      'This code will expire in ' + PIN_CONFIG.RESET_TOKEN_EXPIRY_MINUTES + ' minutes.\n\n' +
      'If you did not request this reset, please ignore this email or contact your union steward.\n\n' +
      'Best regards,\n' +
      (typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Your Union') : 'Your Union');

    MailApp.sendEmail({
      to: memberEmail,
      subject: subject,
      body: body
    });

    if (typeof secureLog === 'function') {
      secureLog('PINResetSent', 'Reset token sent via email', { memberId: memberId });
    }

  } catch (e) {
    log_('Failed to send PIN reset email', e.message);
    if (typeof secureLog === 'function') {
      secureLog('PINResetError', 'Failed to send reset email', { memberId: memberId, error: e.message });
    }
    return errorResponse('Failed to send email. Please try again or contact your steward.');
  }

  return { success: true, message: 'If your Member ID is valid, a reset email has been sent.' };
}

/**
 * Complete PIN reset with token
 * @param {string} memberId - The member ID
 * @param {string} token - The reset token from email
 * @param {string} newPin - The new PIN to set
 * @returns {Object} Result with success status
 */
function completePINReset(memberId, token, newPin) {
  if (!memberId || !token || !newPin) {
    return errorResponse('All fields are required');
  }

  memberId = String(memberId).trim().toUpperCase();
  token = String(token).trim().toUpperCase();
  newPin = String(newPin).trim();

  // Validate new PIN format
  if (!/^\d{6}$/.test(newPin)) {
    return errorResponse('PIN must be exactly 6 digits');
  }

  // Retrieve and verify token (PropertiesService — survives cache eviction)
  var props = PropertiesService.getScriptProperties();
  var propKey = PIN_CONFIG.RESET_TOKEN_PREFIX + memberId;
  var storedData = props.getProperty(propKey);

  if (!storedData) {
    return errorResponse('Reset code has expired or is invalid. Please request a new one.');
  }

  var tokenData;
  try {
    tokenData = JSON.parse(storedData);
  } catch (_e) {
    props.deleteProperty(propKey);
    return errorResponse('Invalid reset data. Please request a new code.');
  }

  // Check expiry (PropertiesService has no auto-TTL, so we check manually)
  if (tokenData.expiresAt && tokenData.expiresAt < Date.now()) {
    props.deleteProperty(propKey);
    return errorResponse('Reset code has expired or is invalid. Please request a new one.');
  }

  // Verify token matches using constant-time comparison. String !== would
  // short-circuit on the first differing character and leak progressive
  // information about the stored token. verifyPIN uses the same pattern.
  var _storedToken = String(tokenData.token || '');
  var _inputToken = String(token || '');
  var _tokenMatches = _storedToken.length === _inputToken.length;
  if (_tokenMatches) {
    var _diff = 0;
    for (var _ti = 0; _ti < _storedToken.length; _ti++) {
      _diff |= (_storedToken.charCodeAt(_ti) ^ _inputToken.charCodeAt(_ti));
    }
    _tokenMatches = _diff === 0;
  }
  if (!_tokenMatches || tokenData.memberId !== memberId) {
    if (typeof secureLog === 'function') {
      secureLog('PINResetFailed', 'Invalid reset token', { memberId: memberId });
    }
    return errorResponse('Invalid reset code. Please check and try again.');
  }

  // Token consumed — delete immediately to prevent reuse
  props.deleteProperty(propKey);

  // Token is valid - update the PIN
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return errorResponse('System configuration error. Please try again later.', 'completePINReset');
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = null;

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], MEMBER_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      break;
    }
  }

  if (!memberRow) {
    return errorResponse('Member not found');
  }

  // Hash and store new PIN
  var newHash = hashPIN(newPin, memberId);
  sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN).setValue(newHash);

  // Reset token already consumed above (props.deleteProperty)

  // Clear any lockouts for this member
  _clearRateCount('pin_lockout_' + memberId);
  _clearRateCount('pin_attempts_' + memberId);

  // Invalidate any existing session tokens for this member's email so a
  // compromised session can't outlive the credential rotation.
  try {
    var _memberEmail = col_(data[memberRow - 1], MEMBER_COLS.EMAIL);
    if (_memberEmail && typeof Auth !== 'undefined' && typeof Auth.invalidateAllSessionsForEmail === 'function') {
      Auth.invalidateAllSessionsForEmail(_memberEmail);
    }
  } catch (_sessErr) { log_('completePINReset', 'session revoke: ' + _sessErr.message); }

  if (typeof secureLog === 'function') {
    secureLog('PINResetComplete', 'PIN successfully reset via email token', { memberId: memberId });
  }

  return { success: true, message: 'Your PIN has been reset successfully. You can now log in with your new PIN.' };
}

// ============================================================================
// RATE LIMITING
// ============================================================================

/**
 * Check if a member is currently locked out due to failed attempts
 * @param {string} memberId - The member ID
 * @returns {Object} { isLocked: boolean, remainingMinutes: number }
 */
function checkPINLockout(memberId) {
  var lockoutKey = 'pin_lockout_' + memberId;
  var lockoutEnd = _getRateCount(lockoutKey);
  if (lockoutEnd > 0) {
    var now = Date.now();
    if (now < lockoutEnd) {
      var remainingMs = lockoutEnd - now;
      return {
        isLocked: true,
        remainingMinutes: Math.ceil(remainingMs / 60000)
      };
    }
    // Expired — clean up
    _clearRateCount(lockoutKey);
  }

  return { isLocked: false, remainingMinutes: 0 };
}

/**
 * Record a failed PIN attempt
 * @param {string} memberId - The member ID
 * @returns {Object} { attemptsRemaining: number, isNowLocked: boolean }
 */
function recordFailedPINAttempt(memberId) {
  var attemptsKey = 'pin_attempts_' + memberId;
  var lockoutKey = 'pin_lockout_' + memberId;
  var lockoutTtl = PIN_CONFIG.LOCKOUT_MINUTES * 60;

  var attempts = _getRateCount(attemptsKey) + 1;

  if (attempts >= PIN_CONFIG.MAX_ATTEMPTS) {
    // Lock out the member — store lockout-end timestamp as the "count"
    var lockoutEnd = Date.now() + (PIN_CONFIG.LOCKOUT_MINUTES * 60 * 1000);
    _setRateCount(lockoutKey, lockoutEnd, lockoutTtl);
    _clearRateCount(attemptsKey);

    // Log security event
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('PIN_LOCKOUT', {
        memberId: memberId,
        attempts: attempts,
        lockoutMinutes: PIN_CONFIG.LOCKOUT_MINUTES
      });
    }

    // Alert admin of potential brute force attempt
    if (typeof recordSecurityEvent === 'function') {
      recordSecurityEvent('PIN_BRUTE_FORCE', typeof SECURITY_SEVERITY !== 'undefined' ? SECURITY_SEVERITY.HIGH : 'HIGH',
        'Member account locked after ' + attempts + ' failed PIN attempts (possible brute force)',
        { memberId: memberId, attempts: attempts, lockoutMinutes: PIN_CONFIG.LOCKOUT_MINUTES });
    }

    return { attemptsRemaining: 0, isNowLocked: true };
  }

  // Store attempt count (expires after lockout period)
  _setRateCount(attemptsKey, attempts, lockoutTtl);

  return {
    attemptsRemaining: PIN_CONFIG.MAX_ATTEMPTS - attempts,
    isNowLocked: false
  };
}

/**
 * Clear failed PIN attempts after successful login
 * @param {string} memberId - The member ID
 */
function clearPINAttempts(memberId) {
  _clearRateCount('pin_attempts_' + memberId);
  _clearRateCount('pin_lockout_' + memberId);
}

// ============================================================================
// MEMBER AUTHENTICATION
// ============================================================================
/**
 * Authenticate a member with their ID and PIN
 * @param {string} memberId - The member ID
 * @param {string} pin - The PIN to verify
 * @returns {Object} Authentication result
 */
function authenticateMember(memberId, pin) {
  // Input validation
  if (!memberId || !pin) {
    return {
      success: false,
      error: 'Member ID and PIN are required'
    };
  }

  // Sanitize inputs
  memberId = String(memberId).trim().toUpperCase();
  pin = String(pin).trim();

  // Check for lockout
  var lockoutStatus = checkPINLockout(memberId);
  if (lockoutStatus.isLocked) {
    return {
      success: false,
      error: 'Account temporarily locked. Try again in ' + lockoutStatus.remainingMinutes + ' minutes.',
      isLocked: true
    };
  }

  // Find member in directory
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return errorResponse('System temporarily unavailable. Please try again later.', 'authenticateMember');
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;
  var storedHash = '';

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], MEMBER_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1; // 1-indexed row
      storedHash = String(col_(data[i], MEMBER_PIN_COLS.PIN_HASH) || '');
      break;
    }
  }

  if (memberRow === -1) {
    // Don't reveal whether member exists - record as failed attempt
    recordFailedPINAttempt(memberId);
    return errorResponse('Invalid Member ID or PIN');
  }

  // Check if PIN is set
  if (!storedHash) {
    return {
      success: false,
      error: 'PIN not set. Please contact your steward to set up your PIN.',
      needsSetup: true
    };
  }

  // Verify PIN
  if (!verifyPIN(pin, memberId, storedHash)) {
    var attemptResult = recordFailedPINAttempt(memberId);

    if (attemptResult.isNowLocked) {
      return {
        success: false,
        error: 'Too many failed attempts. Account locked for ' + PIN_CONFIG.LOCKOUT_MINUTES + ' minutes.',
        isLocked: true
      };
    }

    return {
      success: false,
      error: 'Invalid Member ID or PIN. Please try again.'
    };
  }

  // Success! Clear failed attempts
  clearPINAttempts(memberId);

  // Create session token
  var sessionToken = createMemberSession(memberId);

  // Log successful authentication
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('MEMBER_LOGIN', { memberId: memberId });
  }

  return {
    success: true,
    memberId: memberId,
    sessionToken: sessionToken,
    row: memberRow
  };
}

/**
 * Create a session token for an authenticated member.
 *
 * PIN portal sessions are intentionally stored in CacheService (volatile, ~6h max)
 * rather than PropertiesService (persistent, 500KB quota). This is acceptable because:
 * 1. PIN sessions are short-lived (configurable, default 60 minutes)
 * 2. Cache eviction = user re-enters PIN, which is low-friction
 * 3. Avoids contributing to PropertiesService quota pressure from long-lived web sessions
 *
 * @param {string} memberId - The member ID
 * @returns {string} Session token
 */
function createMemberSession(memberId) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  var sessionData = JSON.stringify({
    memberId: memberId,
    created: Date.now()
  });

  // Store session (expires after configured duration)
  cache.put('member_session_' + token, sessionData, PIN_CONFIG.SESSION_DURATION_MINUTES * 60);

  return token;
}

/**
 * Validate a session token
 * @param {string} token - The session token
 * @returns {Object} { valid: boolean, memberId: string }
 */
function validateMemberSession(token) {
  if (!token) return { valid: false };

  var cache = CacheService.getScriptCache();
  var sessionData = cache.get('member_session_' + token);

  if (!sessionData) {
    return { valid: false };
  }

  try {
    var session = JSON.parse(sessionData);
    return {
      valid: true,
      memberId: session.memberId
    };
  } catch (_e) {
    return { valid: false };
  }
}

// Dead code removed: invalidateMemberSession() — zero callers in src

// ============================================================================
// PIN MANAGEMENT (Steward Functions)
// ============================================================================

/**
 * Generate and set a new PIN for a member
 * Called by stewards to set up member access
 * @param {string} memberId - The member ID
 * @returns {Object} { success: boolean, pin: string (only on success) }
 */
function generateMemberPINForSteward(memberId) {
  // Verify caller is a steward
  var userEmail = Session.getActiveUser().getEmail();
  var userRole = getUserRole_(userEmail);

  if (userRole !== 'admin' && userRole !== 'steward' && userRole !== 'both') {
    return errorResponse('Only stewards can generate member PINs');
  }

  if (!memberId) {
    return errorResponse('Member ID is required');
  }

  memberId = String(memberId).trim().toUpperCase();

  // Find member
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return errorResponse('Member directory not found');
  }

  ensureMinimumColumns(sheet, getMemberHeaders().length);

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;
  var memberName = '';

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], MEMBER_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      memberName = (col_(data[i], MEMBER_COLS.FIRST_NAME) || '') + ' ' + (col_(data[i], MEMBER_COLS.LAST_NAME) || '');
      break;
    }
  }

  if (memberRow === -1) {
    return errorResponse('Member not found: ' + memberId);
  }

  // Generate random 6-digit PIN
  var newPIN = generateMemberPIN();
  var hashedPIN = hashPIN(newPIN, memberId);

  // Store hashed PIN
  sheet.getRange(memberRow, MEMBER_PIN_COLS.PIN_HASH).setValue(hashedPIN);

  // Log the action
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('PIN_GENERATED', {
      memberId: memberId,
      generatedBy: userEmail
    });
  }

  return {
    success: true,
    pin: newPIN,
    memberId: memberId,
    memberName: memberName.trim(),
    message: 'PIN generated successfully. Please provide this PIN to the member securely.'
  };
}
/**
 * Show dialog for generating member PIN
 */
function showGeneratePINDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if on member directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Generate PIN', 'Please select a member in the Member Directory first.', ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Generate PIN', 'Please select a member row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();
  var memberName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue() + ' ' +
                   sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();

  if (!memberId) {
    ui.alert('Generate PIN', 'No Member ID found in selected row.', ui.ButtonSet.OK);
    return;
  }

  var confirm = ui.alert(
    'Generate PIN for ' + memberName,
    'This will generate a new 6-digit PIN for ' + memberName + ' (' + memberId + ').\n\n' +
    'If they already have a PIN, it will be replaced.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  var result = generateMemberPINForSteward(memberId);

  if (result.success) {
    // Email the PIN to the member instead of displaying in plaintext
    var memberEmail = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
    if (memberEmail && String(memberEmail).includes('@')) {
      try {
        var orgName = '';
        try { orgName = getConfigValue_(CONFIG_COLS.ORG_NAME) || 'SolidBase'; } catch (_e) { orgName = 'SolidBase'; }
        MailApp.sendEmail({
          to: String(memberEmail),
          subject: orgName + ' - Your Self-Service Portal PIN',
          body: 'Hello ' + result.memberName + ',\n\n' +
                'Your new self-service portal PIN is: ' + result.pin + '\n\n' +
                'Use this PIN along with your Member ID (' + result.memberId + ') to access the member portal.\n\n' +
                'If you did not request this PIN, please contact your steward.\n\n' +
                '- ' + orgName
        });
        ui.alert('PIN Generated & Emailed',
          'A new PIN for ' + result.memberName + ' has been generated and emailed to ' + memberEmail + '.\n\n' +
          'The PIN was NOT displayed here for security.',
          ui.ButtonSet.OK);
      } catch (emailErr) {
        // Fall back to showing PIN if email fails
        ui.alert('PIN Generated (Email Failed)',
          'New PIN for ' + result.memberName + ': ' + result.pin + '\n\n' +
          'Email delivery failed (' + emailErr.message + '). Please provide this PIN securely.',
          ui.ButtonSet.OK);
      }
    } else {
      // No email on file - must show PIN
      ui.alert('PIN Generated (No Email)',
        'New PIN for ' + result.memberName + ': ' + result.pin + '\n\n' +
        'No email address on file. Please provide this PIN to the member securely.',
        ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Error', result.error, ui.ButtonSet.OK);
  }
}

/**
 * Show dialog for resetting member PIN (same as generate, different messaging)
 */
function showResetPINDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if on member directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Reset PIN', 'Please select a member in the Member Directory first.', ui.ButtonSet.OK);
    return;
  }

  var range = ss.getActiveRange();
  var row = range.getRow();

  if (row < 2) {
    ui.alert('Reset PIN', 'Please select a member row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var memberId = sheet.getRange(row, MEMBER_COLS.MEMBER_ID).getValue();
  var memberName = sheet.getRange(row, MEMBER_COLS.FIRST_NAME).getValue() + ' ' +
                   sheet.getRange(row, MEMBER_COLS.LAST_NAME).getValue();

  if (!memberId) {
    ui.alert('Reset PIN', 'No Member ID found in selected row.', ui.ButtonSet.OK);
    return;
  }

  // Check if member has a PIN
  var pinHash = sheet.getRange(row, PIN_CONFIG.PIN_COLUMN).getValue();

  if (!pinHash) {
    ui.alert('Reset PIN', 'This member does not have a PIN set. Use "Generate Member PIN" instead.', ui.ButtonSet.OK);
    return;
  }

  var confirm = ui.alert(
    'Reset PIN for ' + memberName,
    'This will generate a new PIN for ' + memberName + ', replacing their current PIN.\n\n' +
    'The member will need to use the new PIN to log in.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  var result = generateMemberPINForSteward(memberId);

  if (result.success) {
    // Log the PIN reset
    if (typeof secureLog === 'function') {
      secureLog('PINReset', 'PIN reset for member', { memberId: memberId });
    }

    // Send new PIN via email instead of displaying in plaintext UI alert
    var memberEmail = sheet.getRange(row, MEMBER_COLS.EMAIL).getValue();
    if (memberEmail && String(memberEmail).includes('@')) {
      try {
        var orgName = '';
        try { orgName = getConfigValue_(CONFIG_COLS.ORG_NAME) || 'SolidBase'; } catch (_e) { orgName = 'SolidBase'; }
        MailApp.sendEmail({
          to: String(memberEmail),
          subject: orgName + ' - Your PIN Has Been Reset',
          body: 'Hello ' + result.memberName + ',\n\n' +
                'Your self-service portal PIN has been reset.\n\n' +
                'Your new PIN is: ' + result.pin + '\n\n' +
                'Use this PIN along with your Member ID (' + result.memberId + ') to access the member portal.\n\n' +
                'If you did not request this reset, please contact your steward immediately.\n\n' +
                '- ' + orgName
        });
        ui.alert('PIN Reset Successful',
          'A new PIN for ' + result.memberName + ' has been generated and emailed to ' + memberEmail + '.\n\n' +
          'The PIN was NOT displayed here for security. Their old PIN will no longer work.',
          ui.ButtonSet.OK);
      } catch (emailErr) {
        // Fall back to showing PIN if email fails
        ui.alert('PIN Reset (Email Failed)',
          'New PIN for ' + result.memberName + ': ' + result.pin + '\n\n' +
          'Email delivery failed (' + emailErr.message + '). Please provide this PIN securely.\n' +
          'Their old PIN will no longer work.',
          ui.ButtonSet.OK);
      }
    } else {
      // No email on file - must show PIN
      ui.alert('PIN Reset (No Email)',
        'New PIN for ' + result.memberName + ': ' + result.pin + '\n\n' +
        'No email address on file. Please provide this PIN to the member securely.\n' +
        'Their old PIN will no longer work.',
        ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Error', result.error, ui.ButtonSet.OK);
  }
}

/**
 * Show dialog for bulk generating PINs for multiple members
 */
function showBulkGeneratePINDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Check if on member directory
  if (sheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert('Bulk Generate PINs', 'Please go to the Member Directory first.', ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert(
    'Bulk Generate PINs',
    'This will generate PINs for all members who do not currently have one.\n\n' +
    'Members who already have a PIN will be skipped.\n\n' +
    'Each member will receive their PIN via email.\n' +
    'Members without an email on file will be listed for manual distribution.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var data = sheet.getDataRange().getValues();
  var emailed = 0;
  var noEmailList = [];
  var skipped = 0;
  var errors = [];

  var orgName = '';
  try { orgName = getConfigValue_(CONFIG_COLS.ORG_NAME) || 'SolidBase'; } catch (_e) { orgName = 'SolidBase'; }

  // Start from row 2 (index 1) to skip header
  for (var i = 1; i < data.length; i++) {
    var memberId = col_(data[i], MEMBER_COLS.MEMBER_ID);
    var memberName = ((col_(data[i], MEMBER_COLS.FIRST_NAME) || '') + ' ' +
                     (col_(data[i], MEMBER_COLS.LAST_NAME) || '')).trim();
    var memberEmail = col_(data[i], MEMBER_COLS.EMAIL);
    var existingPin = data[i][PIN_CONFIG.PIN_COLUMN - 1];

    if (!memberId) continue;

    if (existingPin) {
      skipped++;
      continue;
    }

    var result = generateMemberPINForSteward(memberId);
    if (result.success) {
      // Try to email the PIN to the member
      if (memberEmail && String(memberEmail).includes('@')) {
        try {
          MailApp.sendEmail({
            to: String(memberEmail),
            subject: orgName + ' - Your Self-Service Portal PIN',
            body: 'Hello ' + memberName + ',\n\n' +
                  'Your new self-service portal PIN is: ' + result.pin + '\n\n' +
                  'Use this PIN along with your Member ID (' + memberId + ') to access the member portal.\n\n' +
                  'If you did not request this PIN, please contact your steward.\n\n' +
                  '- ' + orgName
          });
          emailed++;
        } catch (emailErr) {
          noEmailList.push(memberName + ' (' + memberId + '): email failed - ' + emailErr.message);
        }
      } else {
        noEmailList.push(memberName + ' (' + memberId + '): no email on file');
      }
    } else {
      errors.push(memberId + ': ' + result.error);
    }
  }

  // Audit log
  if ((emailed + noEmailList.length) > 0 && typeof logAuditEvent === 'function') {
    logAuditEvent('BULK_PIN_GENERATION', {
      emailed: emailed,
      noEmail: noEmailList.length,
      generatedBy: (function() { try { return Session.getActiveUser().getEmail(); } catch (_e) { return 'Unknown'; } })()
    });
  }

  // Show summary (no PINs displayed)
  var htmlContent = '<html><head>' + getMobileOptimizedHead() + '</head><body style="font-family:sans-serif;padding:10px;">';
  htmlContent += '<h3>Bulk PIN Generation Complete</h3>';
  htmlContent += '<p><strong>Emailed:</strong> ' + emailed + ' PINs sent directly to members<br>';
  htmlContent += '<strong>Skipped:</strong> ' + skipped + ' (already had PINs)<br>';
  if (errors.length > 0) {
    htmlContent += '<strong>Errors:</strong> ' + errors.length + '<br>';
  }
  htmlContent += '</p>';

  if (noEmailList.length > 0) {
    htmlContent += '<h4>Members needing manual PIN distribution (' + noEmailList.length + '):</h4>';
    htmlContent += '<textarea readonly style="width:100%;height:200px;font-family:monospace;">';
    for (var n = 0; n < noEmailList.length; n++) {
      htmlContent += escapeHtml(noEmailList[n]) + '\n';
    }
    htmlContent += '</textarea>';
    htmlContent += '<p><em>Contact these members directly to provide portal access.</em></p>';
  }

  if (errors.length > 0) {
    htmlContent += '<h4>Errors:</h4><pre>';
    for (var e = 0; e < Math.min(errors.length, 10); e++) {
      htmlContent += escapeHtml(errors[e]) + '\n';
    }
    htmlContent += '</pre>';
  }

  htmlContent += '</body></html>';
  showDialog_(htmlContent, 'Bulk PIN Generation Results', 500, 400);
}

// ============================================================================
// MEMBER SELF-SERVICE DATA ACCESS
// ============================================================================

/**
 * Get member's own grievances
 * @param {string} sessionToken - Valid session token
 * @returns {Object} Member's grievances
 */
function getMemberGrievances(sessionToken) {
  var session = validateMemberSession(sessionToken);
  if (!session.valid) {
    return errorResponse('Session expired. Please log in again.');
  }

  var memberId = session.memberId;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    return { success: true, grievances: [] };
  }

  var data = sheet.getDataRange().getValues();
  var grievances = [];

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], GRIEVANCE_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      grievances.push({
        grievanceId: col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID) || '',
        status: col_(data[i], GRIEVANCE_COLS.STATUS) || '',
        currentStep: col_(data[i], GRIEVANCE_COLS.CURRENT_STEP) || '',
        issueCategory: col_(data[i], GRIEVANCE_COLS.ISSUE_CATEGORY) || '',
        incidentDate: formatDateMSS_(col_(data[i], GRIEVANCE_COLS.INCIDENT_DATE)),
        filedDate: formatDateMSS_(col_(data[i], GRIEVANCE_COLS.DATE_FILED)),
        steward: col_(data[i], GRIEVANCE_COLS.STEWARD) || '',
        nextDeadline: formatDateMSS_(col_(data[i], GRIEVANCE_COLS.NEXT_ACTION_DUE)),
        resolution: col_(data[i], GRIEVANCE_COLS.RESOLUTION) || ''
      });
    }
  }

  // Sort by filed date descending (most recent first)
  grievances.sort(function(a, b) {
    return new Date(b.filedDate || 0) - new Date(a.filedDate || 0);
  });

  return {
    success: true,
    grievances: grievances,
    count: grievances.length
  };
}

/**
 * Get member's resolved/closed grievance history (PIN-auth portal)
 * Returns only non-sensitive fields: case ID, category, status, outcome, dates.
 * @param {string} sessionTokenOrEmail - Session token or email address
 * @returns {Object} { success: boolean, history: Array }
 */
function getMemberGrievanceHistory(sessionTokenOrEmail) {
  // Determine if input is a session token or email
  var memberId;
  var ss; // hoisted — assigned in both branches and used after the if/else
  if (sessionTokenOrEmail && sessionTokenOrEmail.indexOf('@') !== -1) {
    // Email-based lookup (SPA context) — verify caller authorization
    var email = String(sessionTokenOrEmail).trim().toLowerCase();
    var callerEmail = '';
    try { callerEmail = Session.getActiveUser().getEmail().toLowerCase(); } catch (_e) { log_('getMemberGrievanceHistory', 'Error resolving caller: ' + (_e.message || _e)); }

    // Verify caller authorization — if we can't identify the caller in
    // Execute-as-Me context, deny access (prevents unauthenticated lookups)
    if (!callerEmail) {
      return { success: false, history: [], error: 'Unable to verify caller identity.' };
    }
    if (callerEmail !== email) {
      var callerRole = typeof getUserRole_ === 'function' ? getUserRole_(callerEmail) : 'member';
      if (callerRole !== 'admin' && callerRole !== 'steward') {
        return { success: false, history: [], error: 'Not authorized to view another member\'s history.' };
      }
    }

    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, history: [] };
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet) return { success: true, history: [] };
    var memberData = memberSheet.getDataRange().getValues();
    var emailCol = MEMBER_COLS.EMAIL - 1;
    var idCol = MEMBER_COLS.MEMBER_ID - 1;
    for (var r = 1; r < memberData.length; r++) {
      if (String(memberData[r][emailCol]).trim().toLowerCase() === email) {
        memberId = String(memberData[r][idCol >= 0 ? idCol : emailCol] || '').trim().toUpperCase();
        break;
      }
    }
    if (!memberId) return { success: true, history: [] };
  } else {
    // Session token (PIN portal context)
    var session = validateMemberSession(sessionTokenOrEmail);
    if (!session.valid) {
      return errorResponse('Session expired. Please log in again.');
    }
    memberId = session.memberId;
  }

  ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: true, history: [] };
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { success: true, history: [] };

  var data = sheet.getDataRange().getValues();
  var closedStatuses = ['settled', 'won', 'denied', 'withdrawn', 'closed'];
  var history = [];

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(col_(data[i], GRIEVANCE_COLS.MEMBER_ID) || '').trim().toUpperCase();
    if (rowMemberId !== memberId) continue;

    var status = String(col_(data[i], GRIEVANCE_COLS.STATUS) || '').trim().toLowerCase();
    if (closedStatuses.indexOf(status) === -1) continue;

    history.push({
      grievanceId: col_(data[i], GRIEVANCE_COLS.GRIEVANCE_ID) || '',
      issueCategory: col_(data[i], GRIEVANCE_COLS.ISSUE_CATEGORY) || '',
      status: status.charAt(0).toUpperCase() + status.slice(1),
      outcome: col_(data[i], GRIEVANCE_COLS.RESOLUTION) || '',
      dateFiled: formatDateMSS_(col_(data[i], GRIEVANCE_COLS.DATE_FILED)),
      dateClosed: formatDateMSS_(col_(data[i], GRIEVANCE_COLS.DATE_CLOSED))
    });
  }

  // Sort by filed date descending
  history.sort(function(a, b) {
    return new Date(b.dateFiled || 0) - new Date(a.dateFiled || 0);
  });

  return { success: true, history: history };
}

/**
 * Global wrapper for SPA to call member grievance history.
 * C-AUTH-4: Resolves identity server-side — never accepts client-supplied email.
 * @returns {Object} { success: boolean, history: Array }
 */
function dataGetMemberGrievanceHistoryPortal(sessionToken) {
  var e = (typeof _resolveCallerEmail === 'function') ? _resolveCallerEmail(sessionToken) : '';
  if (!e) {
    try { e = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch (_err) { log_('_err', (_err.message || _err)); }
  }
  return e ? getMemberGrievanceHistory(e) : { success: false, history: [], error: 'Not authenticated.' };
}

/**
 * Format date for display (Member Self Service module)
 * @private
 */
function formatDateMSS_(date) {
  if (!date) return '';
  try {
    var d = new Date(date);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
  } catch (_e) {
    return '';
  }
}

// ============================================================================
// ONBOARDING WIZARD
// ============================================================================

/**
 * Gets onboarding completion status for a member.
 * @param {string} memberEmail
 * @returns {Object} {hasPIN, hasContactInfo, hasNotificationPref, isComplete}
 */
function getOnboardingStatus_(memberEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { hasPIN: false, hasContactInfo: false, hasNotificationPref: false, isComplete: false };
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { hasPIN: false, hasContactInfo: false, hasNotificationPref: false, isComplete: false };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var emailCol = -1, pinCol = -1, phoneCol = -1, addressCol = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase().trim();
    if (h === 'email' || h === 'email address') emailCol = c;
    if (h === 'pin hash' || h === 'pin') pinCol = c;
    if (h === 'phone' || h === 'phone number') phoneCol = c;
    if (h === 'street address' || h === 'address') addressCol = c;
  }

  var target = String(memberEmail).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).trim().toLowerCase() === target) {
      var hasPIN = pinCol >= 0 && String(data[i][pinCol]).trim().length > 0;
      var hasPhone = phoneCol >= 0 && String(data[i][phoneCol]).trim().length > 0;
      var hasAddress = addressCol >= 0 && String(data[i][addressCol]).trim().length > 0;
      var hasContact = hasPhone || hasAddress;

      // Check notification preferences
      var hasNotifPref = false;
      if (typeof DigestService !== 'undefined') {
        var prefs = DigestService.getPreferences(memberEmail);
        hasNotifPref = prefs.frequency !== 'immediate' || prefs.types !== 'all';
      }

      return {
        hasPIN: hasPIN,
        hasContactInfo: hasContact,
        hasNotificationPref: hasNotifPref,
        isComplete: hasPIN && hasContact
      };
    }
  }
  return { hasPIN: false, hasContactInfo: false, hasNotificationPref: false, isComplete: false };
}

/** Global wrapper — get onboarding status. */
function dataGetOnboardingStatus(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { hasPIN: false, hasContactInfo: false, hasNotificationPref: false, isComplete: false };
  return getOnboardingStatus_(e);
}

/** Global wrapper — complete an onboarding step. */
function dataCompleteOnboardingStep(sessionToken, step, data) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };

  if (step === 'pin' && data && data.pin) {
    // Member self-set PIN: look up member ID from email, hash supplied PIN, store it
    var pin = String(data.pin).trim();
    if (!/^\d{6}$/.test(pin)) {
      return { success: false, message: 'PIN must be exactly 6 digits.' };
    }
    // v4.51.1: PIN strength validation — reject trivially guessable patterns
    var weakPins = ['000000', '111111', '222222', '333333', '444444', '555555',
      '666666', '777777', '888888', '999999', '123456', '654321', '012345', '234567',
      '345678', '456789', '567890', '098765', '987654', '876543', '765432'];
    if (weakPins.indexOf(pin) !== -1) {
      return { success: false, message: 'PIN is too common. Please choose a less predictable PIN.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'System error.' };
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!mSheet) return { success: false, message: 'System configuration error.' };
    var mData = mSheet.getDataRange().getValues();
    // Dynamic header-based lookup (matches getOnboardingStatus_ pattern) —
    // immune to column reorder, unlike static MEMBER_COLS constants which
    // depend on syncColumnMaps cache being warm.
    var headers = mData[0];
    var emailCol = -1, memberIdCol = -1, pinHashCol = -1;
    for (var ci = 0; ci < headers.length; ci++) {
      var h = String(headers[ci]).toLowerCase().trim();
      if (h === 'email' || h === 'email address') emailCol = ci;
      if (h === 'member id') memberIdCol = ci;
      if (h === 'pin hash' || h === 'pin') pinHashCol = ci;
    }
    if (emailCol === -1) return { success: false, message: 'System configuration error.' };
    var memberId = null;
    var memberRow = -1;
    for (var mi = 1; mi < mData.length; mi++) {
      if (String(mData[mi][emailCol] || '').toLowerCase().trim() === e.toLowerCase().trim()) {
        memberId = memberIdCol >= 0 ? String(mData[mi][memberIdCol] || '').trim() : '';
        memberRow = mi + 1;
        break;
      }
    }
    if (memberRow === -1) return { success: false, message: 'Account not found. Please contact your steward.' };

    // v4.55.1: Self-heal missing Member ID. A row may exist with no Member ID
    // (manual import, partial seed, legacy data). Rather than blocking PIN setup,
    // generate one in our existing M-prefix format and write it back so the row
    // becomes consistent with the rest of the directory. The hash seed must be
    // stable across set/verify, so we use whatever ID lives on the row after this
    // self-heal — never email, never row index.
    if (!memberId) {
      if (memberIdCol === -1) {
        return { success: false, message: 'Directory schema error: Member ID column missing. Contact your steward.' };
      }
      var existingIds = {};
      for (var ei = 1; ei < mData.length; ei++) {
        var eid = String(mData[ei][memberIdCol] || '').trim();
        if (eid) existingIds[eid] = true;
      }
      var generatedId = '';
      for (var attempt = 0; attempt < 20; attempt++) {
        var candidate = 'M' + Math.floor(100000 + Math.random() * 900000);
        if (!existingIds[candidate]) { generatedId = candidate; break; }
      }
      if (!generatedId) {
        return { success: false, message: 'Could not generate a unique Member ID. Please try again.' };
      }
      mSheet.getRange(memberRow, memberIdCol + 1).setValue(generatedId);
      memberId = generatedId;
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('MEMBER_ID_AUTOGEN', { memberId: generatedId, email: e, reason: 'self_heal_pin_setup' });
      }
    }
    var hashSeed = memberId;
    var pinCol = pinHashCol >= 0 ? pinHashCol + 1 : PIN_CONFIG.PIN_COLUMN;
    var existingHash = mData[memberRow - 1][pinCol - 1];
    if (existingHash) return { success: false, message: 'PIN already set. Use PIN reset instead.' };
    var hashedPin = hashPIN(pin, hashSeed);
    mSheet.getRange(memberRow, pinCol).setValue(hashedPin);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('PIN_SELF_SET', { memberId: memberId, email: e });
    }
    return { success: true, message: 'PIN set successfully.' };
  }

  if (step === 'contact' && data) {
    // Use existing profile update
    if (typeof updateMemberProfile === 'function') {
      return updateMemberProfile(e, data);
    }
    return { success: false, message: 'Profile service unavailable.' };
  }

  if (step === 'notifications' && data) {
    if (typeof DigestService !== 'undefined') {
      return DigestService.setPreferences(e, data.frequency || 'immediate', data.types || 'all');
    }
    return { success: true }; // Silently succeed if digest service not yet available
  }

  return { success: false, message: 'Unknown step: ' + step };
}

// ============================================================================
// WEB APP INTEGRATION
// ============================================================================

// ============================================================================
// PIN LOGIN + STEWARD PIN MANAGEMENT
// Available in all environments. PIN authentication for members.
// ============================================================================

/**
 * Authenticate a user by scanning all member PIN hashes.
 * No email required. Scans entire Member Directory, tries hashPIN() against
 * each row that has a stored hash. Returns a session token on first match.
 *
 * Rate limited globally (10 attempts / 15 min) since memberId is unknown
 * before the scan completes — per-member lockout is applied post-match.
 *
 * @param {string} pin - The plaintext PIN entered by the user
 * @returns {Object} { success, sessionToken, email, memberName } or { success: false, message }
 */
function devAuthLoginByPIN(pin) {
  // ── Input validation ──────────────────────────────────────────────────────
  pin = String(pin || '').trim();
  if (!pin || !/^\d{6}$/.test(pin)) {
    Utilities.sleep(100 + Math.floor(Math.random() * 200));
    return { success: false, message: 'PIN must be exactly 6 digits.' };
  }

  var startTime = Date.now();

  // ── Rate limit — prevent blind brute-force before any row is matched ─────
  // SECURITY NOTE: This function scans ALL members against a single PIN, so
  // each guess tests the entire population. We gate with BOTH a per-session
  // key and a global cap:
  //   - Per-session key (getTemporaryActiveUserKey) prevents one attacker from
  //     locking out every legitimate user by tripping the global limit.
  //   - Global cap is a defense-in-depth ceiling.
  // With 200 members and 5 per-session attempts / 15 min, brute force remains
  // infeasible while legitimate users can still log in during an attack.
  var lockoutMinutes = (typeof PIN_CONFIG !== 'undefined' && PIN_CONFIG.LOCKOUT_MINUTES) || 15;
  var perSessionKey = 'DEV_PIN_SCAN_SESSION_' + (Session.getTemporaryActiveUserKey() || 'anon');
  var perSessionAttempts = _getRateCount(perSessionKey);
  if (perSessionAttempts >= 5) {
    Utilities.sleep(500 + Math.floor(Math.random() * 500));
    return { success: false, message: 'Too many attempts. Try again in ' + lockoutMinutes + ' minutes.' };
  }
  var globalRateKey = 'DEV_PIN_SCAN_RATE';
  var globalAttempts = _getRateCount(globalRateKey);
  if (globalAttempts >= 50) {
    Utilities.sleep(500 + Math.floor(Math.random() * 500));
    return { success: false, message: 'System is temporarily rate-limited. Try again in ' + lockoutMinutes + ' minutes.' };
  }
  _setRateCount(perSessionKey, perSessionAttempts + 1, 900); // 15 min window
  _setRateCount(globalRateKey, globalAttempts + 1, 900);

  // ── Sheet scan ────────────────────────────────────────────────────────────
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };

  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { success: false, message: 'Member directory not found.' };

  var data = sheet.getDataRange().getValues();
  var pinCol  = PIN_CONFIG.PIN_COLUMN - 1;         // 0-indexed
  var idCol   = MEMBER_COLS.MEMBER_ID  - 1;
  var emailCol= MEMBER_COLS.EMAIL      - 1;
  var fNameCol= MEMBER_COLS.FIRST_NAME - 1;
  var lNameCol= MEMBER_COLS.LAST_NAME  - 1;

  // ── Constant-time scan — always iterate ALL members ──────────────────────
  var matchResult = null;

  for (var i = 1; i < data.length; i++) {
    var storedHash = String(data[i][pinCol] || '').trim();
    if (!storedHash) continue; // skip members with no PIN set

    var memberId = String(data[i][idCol] || '').trim().toUpperCase();
    if (!memberId) continue;

    // Per-member lockout check
    var isLocked = false;
    if (typeof checkPINLockout === 'function') {
      var lockout = checkPINLockout(memberId);
      if (lockout.isLocked) isLocked = true;
    }

    // Always verify PIN (maintain constant iteration) but only accept if not locked
    var isMatch = !isLocked && verifyPIN(pin, memberId, storedHash);

    if (isMatch && !matchResult) {
      matchResult = {
        email: String(data[i][emailCol] || '').trim().toLowerCase(),
        firstName: String(data[i][fNameCol] || '').trim(),
        lastName: String(data[i][lNameCol] || '').trim(),
        memberId: memberId
      };
      // Continue iterating — don't break early (constant time)
    }
  }

  // ── Normalize response timing ─────────────────────────────────────────────
  var elapsed = Date.now() - startTime;
  var targetMs = 1000 + Math.floor(Math.random() * 500);
  if (elapsed < targetMs) {
    Utilities.sleep(targetMs - elapsed);
  }

  // ── Handle match ──────────────────────────────────────────────────────────
  if (matchResult) {
    if (!matchResult.email) {
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('DEV_PIN_LOGIN_NO_EMAIL', { memberId: matchResult.memberId });
      }
      return { success: false, message: 'Account has no email on file. Contact your steward.' };
    }

    // Clear rate counters on success
    try { _clearRateCount(globalRateKey); } catch (_) {}
    try { if (typeof clearPINAttempts === 'function') clearPINAttempts(matchResult.memberId); } catch (_) {}

    // Create session token
    var token = Auth.createSessionToken(matchResult.email, 'pin');
    if (token && typeof token === 'object' && token.error) {
      return { success: false, message: 'Session creation failed. Try again.' };
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('DEV_PIN_LOGIN', { memberId: matchResult.memberId, email: matchResult.email });
    }

    var memberName = (matchResult.firstName + ' ' + matchResult.lastName).trim() || matchResult.memberId;
    return { success: true, sessionToken: token, email: matchResult.email, memberName: memberName };
  }

  // ── No match ──────────────────────────────────────────────────────────────
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('DEV_PIN_LOGIN_FAIL', { reason: 'no_match' });
  }
  // Report against the per-session limit (5) so the user sees an accurate
  // countdown for their own session. Global limit is a defense-in-depth cap.
  var remaining = Math.max(0, 5 - (perSessionAttempts + 1));
  return {
    success: false,
    message: remaining > 0
      ? 'Incorrect PIN. ' + remaining + ' attempt' + (remaining === 1 ? '' : 's') + ' remaining.'
      : 'Too many attempts. Try again in ' + lockoutMinutes + ' minutes.'
  };
}

/**
 * Self-service PIN reset — sends a reset email with a new temporary PIN.
 * Rate-limited to prevent abuse. Uses the same generic response for
 * both valid and invalid emails to prevent enumeration.
 *
 * @param {string} email - Member's email address
 * @returns {Object} { success, message }
 */
function devRequestPINReset(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    Utilities.sleep(200 + Math.floor(Math.random() * 300));
    return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
  }

  // Rate limit — 2 per email per hour
  var rateKey = 'PIN_RESET_RATE_' + email;
  var count = _getRateCount(rateKey);
  if (count >= 2) {
    Utilities.sleep(500 + Math.floor(Math.random() * 500));
    return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
  }
  _setRateCount(rateKey, count + 1, 3600);

  // Global rate limit — 10 resets per hour across all users
  var globalKey = 'PIN_RESET_GLOBAL';
  var globalCount = _getRateCount(globalKey);
  if (globalCount >= 10) {
    Utilities.sleep(500 + Math.floor(Math.random() * 500));
    return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
  }
  _setRateCount(globalKey, globalCount + 1, 3600);

  // Look up member
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };

  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };

  var data = sheet.getDataRange().getValues();
  var emailCol = MEMBER_COLS.EMAIL - 1;
  var idCol    = MEMBER_COLS.MEMBER_ID - 1;
  var fNameCol = MEMBER_COLS.FIRST_NAME - 1;
  var memberRow = -1;
  var memberId = '';
  var firstName = '';

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol] || '').trim().toLowerCase() === email) {
      memberRow = i + 1;
      memberId = String(data[i][idCol] || '').trim().toUpperCase();
      firstName = String(data[i][fNameCol] || '').trim();
      break;
    }
  }

  if (memberRow === -1 || !memberId) {
    // Simulate processing time to prevent enumeration
    Utilities.sleep(500 + Math.floor(Math.random() * 500));
    return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
  }

  // Generate new PIN and hash it
  var newPin = generateMemberPIN();
  var hashedPin = hashPIN(newPin, memberId);

  // Save old PIN hash so we can restore on email failure
  var pinCell = sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN);
  var oldHash = pinCell.getValue();

  // Write new hash
  pinCell.setValue(hashedPin);

  // Clear any lockout
  try { if (typeof clearPINAttempts === 'function') clearPINAttempts(memberId); } catch (_) {}

  // Send email with new PIN — if delivery fails, restore old hash
  var orgName = typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Your Union') : 'Your Union';
  try {
    if (orgName === 'Your Union' && typeof ConfigReader !== 'undefined') {
      var config = ConfigReader.getConfig();
      orgName = config.orgName || orgName;
    }
  } catch (_) {}

  var subject = 'Your new PIN — ' + orgName;
  var body = 'Hi ' + (firstName || 'Member') + ',\n\n' +
    'Your PIN has been reset. Your new PIN is:\n\n' +
    '    ' + newPin + '\n\n' +
    'Use this PIN to sign in at the member portal.\n' +
    'For security, please memorize this PIN — it will not be shown again.\n\n' +
    '— ' + orgName;

  var emailSent = false;
  try {
    GmailApp.sendEmail(email, subject, body);
    emailSent = true;
  } catch (_gmailErr) {
    try {
      MailApp.sendEmail(email, subject, body);
      emailSent = true;
    } catch (_) {}
  }

  if (!emailSent) {
    // Restore old PIN hash — member keeps existing credentials
    pinCell.setValue(oldHash);
    return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
  }

  if (typeof logAuditEvent === 'function') {
    logAuditEvent('PIN_SELF_RESET', { memberId: memberId, email: email });
  }

  return { success: true, message: 'If this email is in our directory, you will receive a new PIN.' };
}

/**
 * DEV ONLY — Steward generates or resets a member's PIN from the web app.
 * Requires steward or admin role. Returns plaintext PIN once — never stored.
 * The PIN is shown to the steward in the UI to pass to the member in person.
 *
 * @param {string} sessionToken - Steward's session token
 * @param {string} memberEmail  - Email of the member to generate PIN for
 * @returns {Object} { success, pin, memberName, isReset, message }
 */
/**
 * Member sets their own PIN from the web app PIN Setup page.
 * Requires member auth via session token. PIN must be exactly 6 digits.
 *
 * @param {string} sessionToken - Member's session token
 * @param {string} pin          - The new 6-digit PIN chosen by the member
 * @param {string} [idemKey]    - Idempotency key (optional)
 * @returns {Object} { success, message }
 */
function memberSetOwnPIN(sessionToken, pin, idemKey) {
  // ── Auth: resolve member email from session ────────────────────────────────
  var email;
  try {
    email = Session.getActiveUser().getEmail();
    if (!email && typeof Auth !== 'undefined' && typeof Auth.resolveEmailFromToken === 'function') {
      email = Auth.resolveEmailFromToken(sessionToken);
    }
  } catch (_e) {
    if (typeof Auth !== 'undefined' && typeof Auth.resolveEmailFromToken === 'function') {
      email = Auth.resolveEmailFromToken(sessionToken);
    }
  }
  if (!email) return { success: false, message: 'Authentication required.' };
  email = String(email).trim().toLowerCase();

  // v4.55.1 N14-BUG-10: honor idemKey to prevent duplicate audit entries on double-click
  if (idemKey) {
    try {
      var idemCache = CacheService.getScriptCache();
      var idemCacheKey = 'pin_idem_' + email + '_' + idemKey;
      var cached = idemCache.get(idemCacheKey);
      if (cached) return JSON.parse(cached);
      var resultBody = _memberSetOwnPINCore(email, pin, sessionToken);
      try { idemCache.put(idemCacheKey, JSON.stringify(resultBody), 600); } catch (_ce) {}
      return resultBody;
    } catch (_idemErr) { /* fall through to non-idempotent path */ }
  }
  return _memberSetOwnPINCore(email, pin, sessionToken);
}

function _memberSetOwnPINCore(email, pin, sessionToken) {

  // ── Validate PIN ──────────────────────────────────────────────────────────
  pin = String(pin || '').trim();
  if (!/^\d{6}$/.test(pin)) {
    return { success: false, message: 'PIN must be exactly 6 digits.' };
  }
  // v4.51.1: PIN strength validation — reject trivially guessable patterns
  var weakPins = ['000000', '111111', '222222', '333333', '444444', '555555',
    '666666', '777777', '888888', '999999', '123456', '654321', '012345', '234567',
    '345678', '456789', '567890', '098765', '987654', '876543', '765432'];
  if (weakPins.indexOf(pin) !== -1) {
    return { success: false, message: 'PIN is too common. Please choose a less predictable PIN.' };
  }

  // ── Find member row ───────────────────────────────────────────────────────
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };

  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { success: false, message: 'Member directory not found.' };

  var data     = sheet.getDataRange().getValues();
  var emailCol = MEMBER_COLS.EMAIL     - 1;
  var idCol    = MEMBER_COLS.MEMBER_ID - 1;
  var pinCol   = PIN_CONFIG.PIN_COLUMN - 1;

  var memberRow = -1;
  var memberId  = '';
  var hadPIN    = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol] || '').trim().toLowerCase() !== email) continue;
    memberRow = i + 1;
    memberId  = String(data[i][idCol] || '').trim().toUpperCase();
    hadPIN    = String(data[i][pinCol] || '').trim().length > 0;
    break;
  }

  if (memberRow === -1) return { success: false, message: 'Account not found. Please contact your steward.' };
  if (!memberId) return { success: false, message: 'Member ID required before setting PIN.' };
  var hashSeed = memberId;

  // ── Hash and store ────────────────────────────────────────────────────────
  var hashedPin = hashPIN(pin, hashSeed);
  sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN).setValue(hashedPin);

  var action = hadPIN ? 'PIN_RESET_BY_MEMBER' : 'PIN_SET_BY_MEMBER';
  if (typeof logAuditEvent === 'function') {
    logAuditEvent(action, { memberId: memberId, memberEmail: email });
  }

  // Invalidate any OTHER session tokens for this email so a compromised
  // session can't outlive the credential rotation. Keep the current session
  // (the user just used it to set their PIN) — we do a targeted invalidate
  // of all sessions, then let the caller re-authenticate if needed.
  if (hadPIN) {
    try {
      if (typeof Auth !== 'undefined' && typeof Auth.invalidateAllSessionsForEmail === 'function') {
        Auth.invalidateAllSessionsForEmail(email);
      }
    } catch (_sessErr) { log_('memberSetOwnPIN', 'session revoke: ' + _sessErr.message); }
  }

  return { success: true, message: hadPIN ? 'PIN updated successfully.' : 'PIN set successfully.' };
}

function devStewardManageMemberPIN(sessionToken, memberEmail) {
  // ── Auth: steward/admin only ───────────────────────────────────────────────
  var stewardEmail = _requireStewardAuth(sessionToken);
  if (!stewardEmail) {
    return { success: false, message: 'Steward authorization required.' };
  }

  if (!memberEmail) {
    return { success: false, message: 'Member email is required.' };
  }
  memberEmail = String(memberEmail).trim().toLowerCase();

  // ── Locate member by email ────────────────────────────────────────────────
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };

  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { success: false, message: 'Member directory not found.' };

  var data    = sheet.getDataRange().getValues();
  var emailCol= MEMBER_COLS.EMAIL      - 1;
  var idCol   = MEMBER_COLS.MEMBER_ID  - 1;
  var fNameCol= MEMBER_COLS.FIRST_NAME - 1;
  var lNameCol= MEMBER_COLS.LAST_NAME  - 1;
  var pinCol  = PIN_CONFIG.PIN_COLUMN  - 1;

  var memberRow = -1;
  var memberId  = '';
  var memberName= '';
  var hadPIN    = false;

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol] || '').trim().toLowerCase();
    if (rowEmail !== memberEmail) continue;

    memberRow  = i + 1; // 1-indexed sheet row
    memberId   = String(data[i][idCol]   || '').trim().toUpperCase();
    memberName = (String(data[i][fNameCol] || '') + ' ' + String(data[i][lNameCol] || '')).trim();
    hadPIN     = String(data[i][pinCol] || '').trim().length > 0;
    break;
  }

  if (memberRow === -1) {
    return { success: false, message: 'Member not found: ' + memberEmail };
  }
  if (!memberId) {
    return { success: false, message: 'Member has no Member ID — cannot generate PIN.' };
  }

  // ── Generate, hash, store ─────────────────────────────────────────────────
  var newPin    = generateMemberPIN();         // 6-digit numeric string
  var hashedPin = hashPIN(newPin, memberId);   // salted SHA-256
  sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN).setValue(hashedPin);

  var action = hadPIN ? 'PIN_RESET_BY_STEWARD' : 'PIN_GENERATED_BY_STEWARD';
  if (typeof logAuditEvent === 'function') {
    logAuditEvent(action, {
      memberId:     memberId,
      memberEmail:  memberEmail,
      generatedBy:  stewardEmail
    });
  }

  return {
    success:    true,
    pin:        newPin,       // plaintext — shown once in UI, never persisted
    memberId:   memberId,
    memberName: memberName || memberEmail,
    isReset:    hadPIN,
    message:    hadPIN
      ? 'PIN reset for ' + (memberName || memberEmail) + '. Share this PIN with the member in person.'
      : 'PIN generated for ' + (memberName || memberEmail) + '. Share this PIN with the member in person.'
  };
}
