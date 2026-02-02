/**
 * 509 Dashboard - Member Self-Service Portal
 *
 * This module provides PIN-based authentication for members to:
 * - Look up their own information
 * - Update their contact information
 * - View their grievance claims
 *
 * Security Features:
 * - PINs are hashed using SHA-256 before storage
 * - Rate limiting on failed PIN attempts (5 attempts per 15 minutes)
 * - All access and changes are audit logged
 * - Members can only view/edit their own data
 *
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

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
  PIN_COLUMN: 33,                   // AG column for PIN hash storage
  SALT_PROPERTY: 'MEMBER_PIN_SALT', // Property key for salt storage
  RESET_TOKEN_EXPIRY_MINUTES: 30,   // Reset token expiration time
  RESET_TOKEN_PREFIX: 'pin_reset_'  // Cache key prefix for reset tokens
};

/**
 * Member self-service column mapping
 * Column AG (33) stores the hashed PIN
 * @const {Object}
 */
var MEMBER_PIN_COLS = {
  PIN_HASH: 33  // AG - Hashed PIN (never store plaintext)
};

// ============================================================================
// PIN GENERATION AND HASHING
// ============================================================================

/**
 * Generate a random 6-digit PIN
 * @returns {string} 6-digit PIN
 */
function generateMemberPIN() {
  var pin = '';
  for (var i = 0; i < PIN_CONFIG.PIN_LENGTH; i++) {
    pin += Math.floor(Math.random() * 10).toString();
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
  return computedHash === storedHash;
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
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // Removed ambiguous characters
  var token = '';
  for (var i = 0; i < 8; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

/**
 * Request a PIN reset - sends reset token via email
 * @param {string} memberId - The member ID
 * @returns {Object} Result with success status
 */
function requestPINReset(memberId) {
  if (!memberId) {
    return { success: false, error: 'Member ID is required' };
  }

  memberId = String(memberId).trim().toUpperCase();

  // Find member and get their email
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return { success: false, error: 'System error' };
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = null;
  var memberEmail = null;
  var memberName = null;

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      memberEmail = data[i][MEMBER_COLS.EMAIL - 1];
      memberName = (data[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (data[i][MEMBER_COLS.LAST_NAME - 1] || '');
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

  // Generate reset token and store it
  var resetToken = generateResetToken_();
  var cache = CacheService.getScriptCache();
  var cacheKey = PIN_CONFIG.RESET_TOKEN_PREFIX + memberId;

  // Store token with member ID verification (expires in 30 minutes)
  var tokenData = JSON.stringify({
    token: resetToken,
    memberId: memberId,
    created: Date.now()
  });
  cache.put(cacheKey, tokenData, PIN_CONFIG.RESET_TOKEN_EXPIRY_MINUTES * 60);

  // Send reset email
  try {
    var subject = '509 Union - PIN Reset Request';
    var body = 'Hello ' + memberName.trim() + ',\n\n' +
      'You requested a PIN reset for the Member Self-Service Portal.\n\n' +
      'Your reset code is: ' + resetToken + '\n\n' +
      'This code will expire in ' + PIN_CONFIG.RESET_TOKEN_EXPIRY_MINUTES + ' minutes.\n\n' +
      'If you did not request this reset, please ignore this email or contact your union steward.\n\n' +
      'Best regards,\n' +
      'WFSE Local 509';

    MailApp.sendEmail({
      to: memberEmail,
      subject: subject,
      body: body
    });

    if (typeof secureLog === 'function') {
      secureLog('PINResetSent', 'Reset token sent via email', { memberId: memberId });
    }

  } catch (e) {
    Logger.log('Failed to send PIN reset email: ' + e.message);
    if (typeof secureLog === 'function') {
      secureLog('PINResetError', 'Failed to send reset email', { memberId: memberId, error: e.message });
    }
    return { success: false, error: 'Failed to send email. Please try again or contact your steward.' };
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
    return { success: false, error: 'All fields are required' };
  }

  memberId = String(memberId).trim().toUpperCase();
  token = String(token).trim().toUpperCase();
  newPin = String(newPin).trim();

  // Validate new PIN format
  if (!/^\d{6}$/.test(newPin)) {
    return { success: false, error: 'PIN must be exactly 6 digits' };
  }

  // Retrieve and verify token
  var cache = CacheService.getScriptCache();
  var cacheKey = PIN_CONFIG.RESET_TOKEN_PREFIX + memberId;
  var storedData = cache.get(cacheKey);

  if (!storedData) {
    return { success: false, error: 'Reset code has expired or is invalid. Please request a new one.' };
  }

  var tokenData;
  try {
    tokenData = JSON.parse(storedData);
  } catch (e) {
    return { success: false, error: 'Invalid reset data. Please request a new code.' };
  }

  // Verify token matches
  if (tokenData.token !== token || tokenData.memberId !== memberId) {
    // Log failed attempt
    if (typeof secureLog === 'function') {
      secureLog('PINResetFailed', 'Invalid reset token', { memberId: memberId });
    }
    return { success: false, error: 'Invalid reset code. Please check and try again.' };
  }

  // Token is valid - update the PIN
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return { success: false, error: 'System error' };
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = null;

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      break;
    }
  }

  if (!memberRow) {
    return { success: false, error: 'Member not found' };
  }

  // Hash and store new PIN
  var newHash = hashPIN(newPin, memberId);
  sheet.getRange(memberRow, PIN_CONFIG.PIN_COLUMN).setValue(newHash);

  // Invalidate the reset token
  cache.remove(cacheKey);

  // Clear any lockouts for this member
  cache.remove('pin_lockout_' + memberId);
  cache.remove('pin_attempts_' + memberId);

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
  var cache = CacheService.getScriptCache();
  var lockoutKey = 'pin_lockout_' + memberId;
  var attemptsKey = 'pin_attempts_' + memberId;

  var lockoutTime = cache.get(lockoutKey);
  if (lockoutTime) {
    var lockoutEnd = parseInt(lockoutTime, 10);
    var now = Date.now();
    if (now < lockoutEnd) {
      var remainingMs = lockoutEnd - now;
      return {
        isLocked: true,
        remainingMinutes: Math.ceil(remainingMs / 60000)
      };
    }
  }

  return { isLocked: false, remainingMinutes: 0 };
}

/**
 * Record a failed PIN attempt
 * @param {string} memberId - The member ID
 * @returns {Object} { attemptsRemaining: number, isNowLocked: boolean }
 */
function recordFailedPINAttempt(memberId) {
  var cache = CacheService.getScriptCache();
  var attemptsKey = 'pin_attempts_' + memberId;
  var lockoutKey = 'pin_lockout_' + memberId;

  var attempts = parseInt(cache.get(attemptsKey) || '0', 10) + 1;

  if (attempts >= PIN_CONFIG.MAX_ATTEMPTS) {
    // Lock out the member
    var lockoutEnd = Date.now() + (PIN_CONFIG.LOCKOUT_MINUTES * 60 * 1000);
    cache.put(lockoutKey, lockoutEnd.toString(), PIN_CONFIG.LOCKOUT_MINUTES * 60);
    cache.remove(attemptsKey);

    // Log security event
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('PIN_LOCKOUT', {
        memberId: memberId,
        attempts: attempts,
        lockoutMinutes: PIN_CONFIG.LOCKOUT_MINUTES
      });
    }

    return { attemptsRemaining: 0, isNowLocked: true };
  }

  // Store attempt count (expires after lockout period)
  cache.put(attemptsKey, attempts.toString(), PIN_CONFIG.LOCKOUT_MINUTES * 60);

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
  var cache = CacheService.getScriptCache();
  cache.remove('pin_attempts_' + memberId);
  cache.remove('pin_lockout_' + memberId);
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
    return { success: false, error: 'System error: Member directory not found' };
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;
  var storedHash = '';

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1; // 1-indexed row
      storedHash = String(data[i][MEMBER_PIN_COLS.PIN_HASH - 1] || '');
      break;
    }
  }

  if (memberRow === -1) {
    // Don't reveal whether member exists - record as failed attempt
    recordFailedPINAttempt(memberId);
    return { success: false, error: 'Invalid Member ID or PIN' };
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
      error: 'Invalid Member ID or PIN. ' + attemptResult.attemptsRemaining + ' attempts remaining.'
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
 * Create a session token for an authenticated member
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
  } catch (e) {
    return { valid: false };
  }
}

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

  if (userRole !== 'admin' && userRole !== 'steward') {
    return { success: false, error: 'Only stewards can generate member PINs' };
  }

  if (!memberId) {
    return { success: false, error: 'Member ID is required' };
  }

  memberId = String(memberId).trim().toUpperCase();

  // Find member
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return { success: false, error: 'Member directory not found' };
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;
  var memberName = '';

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      memberName = (data[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (data[i][MEMBER_COLS.LAST_NAME - 1] || '');
      break;
    }
  }

  if (memberRow === -1) {
    return { success: false, error: 'Member not found: ' + memberId };
  }

  // Generate new PIN
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
 * Reset a member's PIN (generates new one)
 * @param {string} memberId - The member ID
 * @returns {Object} Result with new PIN
 */
function resetMemberPIN(memberId) {
  // Same as generate - just calls it
  return generateMemberPINForSteward(memberId);
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
    ui.alert(
      'PIN Generated',
      'New PIN for ' + result.memberName + ':\n\n' +
      '    ' + result.pin + '\n\n' +
      'Please provide this PIN to the member securely.\n' +
      'They can use it with their Member ID to access the self-service portal.',
      ui.ButtonSet.OK
    );
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

    ui.alert(
      'PIN Reset Successful',
      'New PIN for ' + result.memberName + ':\n\n' +
      '    ' + result.pin + '\n\n' +
      'Please provide this new PIN to the member securely.\n' +
      'Their old PIN will no longer work.',
      ui.ButtonSet.OK
    );
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
    'A report will be generated showing all new PINs.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var data = sheet.getDataRange().getValues();
  var generated = [];
  var skipped = 0;
  var errors = [];

  // Start from row 2 (index 1) to skip header
  for (var i = 1; i < data.length; i++) {
    var memberId = data[i][MEMBER_COLS.MEMBER_ID - 1];
    var memberName = (data[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                     (data[i][MEMBER_COLS.LAST_NAME - 1] || '');
    var existingPin = data[i][PIN_CONFIG.PIN_COLUMN - 1];

    if (!memberId) continue;

    if (existingPin) {
      skipped++;
      continue;
    }

    var result = generateMemberPINForSteward(memberId);
    if (result.success) {
      generated.push({
        memberId: memberId,
        memberName: memberName.trim(),
        pin: result.pin
      });
    } else {
      errors.push(memberId + ': ' + result.error);
    }
  }

  // Show results
  var message = 'Bulk PIN Generation Complete\n\n';
  message += 'Generated: ' + generated.length + ' new PINs\n';
  message += 'Skipped: ' + skipped + ' (already had PINs)\n';
  if (errors.length > 0) {
    message += 'Errors: ' + errors.length + '\n';
  }

  if (generated.length > 0) {
    message += '\n--- New PINs ---\n';
    for (var j = 0; j < generated.length && j < 50; j++) {
      message += generated[j].memberName + ' (' + generated[j].memberId + '): ' + generated[j].pin + '\n';
    }
    if (generated.length > 50) {
      message += '... and ' + (generated.length - 50) + ' more\n';
    }
    message += '\nIMPORTANT: Note these PINs to distribute to members.\n';
    message += 'They will not be shown again.';
  }

  // Use a scrollable HTML dialog for large results
  if (generated.length > 10) {
    var htmlContent = '<html><body style="font-family:sans-serif;padding:10px;">';
    htmlContent += '<h3>Bulk PIN Generation Complete</h3>';
    htmlContent += '<p><strong>Generated:</strong> ' + generated.length + ' new PINs<br>';
    htmlContent += '<strong>Skipped:</strong> ' + skipped + ' (already had PINs)</p>';
    if (generated.length > 0) {
      htmlContent += '<h4>New PINs (copy this list!):</h4>';
      htmlContent += '<textarea readonly style="width:100%;height:300px;font-family:monospace;">';
      for (var k = 0; k < generated.length; k++) {
        htmlContent += generated[k].memberName + ' (' + generated[k].memberId + '): ' + generated[k].pin + '\n';
      }
      htmlContent += '</textarea>';
    }
    htmlContent += '</body></html>';

    var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(500)
      .setHeight(500);
    ui.showModalDialog(htmlOutput, 'Bulk PIN Generation Results');
  } else {
    ui.alert('Bulk PIN Generation', message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// MEMBER SELF-SERVICE DATA ACCESS
// ============================================================================

/**
 * Get member's own profile data
 * @param {string} sessionToken - Valid session token
 * @returns {Object} Member profile data
 */
function getMemberProfile(sessionToken) {
  var session = validateMemberSession(sessionToken);
  if (!session.valid) {
    return { success: false, error: 'Session expired. Please log in again.' };
  }

  var memberId = session.memberId;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return { success: false, error: 'System error' };
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      return {
        success: true,
        profile: {
          memberId: data[i][MEMBER_COLS.MEMBER_ID - 1] || '',
          firstName: data[i][MEMBER_COLS.FIRST_NAME - 1] || '',
          lastName: data[i][MEMBER_COLS.LAST_NAME - 1] || '',
          jobTitle: data[i][MEMBER_COLS.JOB_TITLE - 1] || '',
          workLocation: data[i][MEMBER_COLS.WORK_LOCATION - 1] || '',
          unit: data[i][MEMBER_COLS.UNIT - 1] || '',
          email: data[i][MEMBER_COLS.EMAIL - 1] || '',
          phone: data[i][MEMBER_COLS.PHONE - 1] || '',
          preferredComm: data[i][MEMBER_COLS.PREFERRED_COMM - 1] || '',
          bestTime: data[i][MEMBER_COLS.BEST_TIME - 1] || '',
          assignedSteward: data[i][MEMBER_COLS.ASSIGNED_STEWARD - 1] || '',
          hasOpenGrievance: data[i][MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] || 'No'
        }
      };
    }
  }

  return { success: false, error: 'Member not found' };
}

/**
 * Update member's own contact information
 * @param {string} sessionToken - Valid session token
 * @param {Object} updates - Fields to update
 * @returns {Object} Result
 */
function updateMemberContact(sessionToken, updates) {
  var session = validateMemberSession(sessionToken);
  if (!session.valid) {
    return { success: false, error: 'Session expired. Please log in again.' };
  }

  var memberId = session.memberId;

  // Validate updates - only allow contact fields
  var allowedFields = ['email', 'phone', 'preferredComm', 'bestTime'];
  var fieldMapping = {
    email: MEMBER_COLS.EMAIL,
    phone: MEMBER_COLS.PHONE,
    preferredComm: MEMBER_COLS.PREFERRED_COMM,
    bestTime: MEMBER_COLS.BEST_TIME
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    return { success: false, error: 'System error' };
  }

  var data = sheet.getDataRange().getValues();
  var memberRow = -1;

  for (var i = 1; i < data.length; i++) {
    var rowMemberId = String(data[i][MEMBER_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      memberRow = i + 1;
      break;
    }
  }

  if (memberRow === -1) {
    return { success: false, error: 'Member not found' };
  }

  // Apply updates
  var updated = [];
  for (var field in updates) {
    if (allowedFields.indexOf(field) >= 0 && fieldMapping[field]) {
      var value = String(updates[field] || '').trim();

      // Basic validation
      if (field === 'email' && value && !isValidEmailMSS_(value)) {
        return { success: false, error: 'Invalid email format' };
      }
      if (field === 'phone' && value) {
        value = formatPhoneNumber_(value);
      }

      sheet.getRange(memberRow, fieldMapping[field]).setValue(value);
      updated.push(field);
    }
  }

  if (updated.length === 0) {
    return { success: false, error: 'No valid fields to update' };
  }

  // Log the update
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('MEMBER_SELF_UPDATE', {
      memberId: memberId,
      fields: updated.join(', ')
    });
  }

  return {
    success: true,
    message: 'Contact information updated',
    updatedFields: updated
  };
}

/**
 * Simple email validation (Member Self Service module)
 * @private
 */
function isValidEmailMSS_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/**
 * Format phone number
 * @private
 */
function formatPhoneNumber_(phone) {
  var digits = phone.replace(/\D/g, '');
  if (digits.length === 10) {
    return '(' + digits.slice(0,3) + ') ' + digits.slice(3,6) + '-' + digits.slice(6);
  }
  return phone;
}

/**
 * Get member's own grievances
 * @param {string} sessionToken - Valid session token
 * @returns {Object} Member's grievances
 */
function getMemberGrievances(sessionToken) {
  var session = validateMemberSession(sessionToken);
  if (!session.valid) {
    return { success: false, error: 'Session expired. Please log in again.' };
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
    var rowMemberId = String(data[i][GRIEVANCE_COLS.MEMBER_ID - 1] || '').trim().toUpperCase();
    if (rowMemberId === memberId) {
      grievances.push({
        grievanceId: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '',
        status: data[i][GRIEVANCE_COLS.STATUS - 1] || '',
        currentStep: data[i][GRIEVANCE_COLS.CURRENT_STEP - 1] || '',
        issueCategory: data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '',
        incidentDate: formatDateMSS_(data[i][GRIEVANCE_COLS.INCIDENT_DATE - 1]),
        filedDate: formatDateMSS_(data[i][GRIEVANCE_COLS.DATE_FILED - 1]),
        steward: data[i][GRIEVANCE_COLS.STEWARD - 1] || '',
        nextDeadline: formatDateMSS_(data[i][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1]),
        resolution: data[i][GRIEVANCE_COLS.RESOLUTION - 1] || '',
        outcome: data[i][GRIEVANCE_COLS.OUTCOME - 1] || ''
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
 * Format date for display (Member Self Service module)
 * @private
 */
function formatDateMSS_(date) {
  if (!date) return '';
  try {
    var d = new Date(date);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d, yyyy');
  } catch (e) {
    return '';
  }
}

// ============================================================================
// MEMBER SELF-SERVICE PORTAL UI
// ============================================================================

/**
 * Get the Member Self-Service Portal HTML
 * @returns {string} HTML content
 */
function getMemberSelfServicePortalHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Member Self-Service Portal</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header
    '.header{background:linear-gradient(135deg,#059669,#047857);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:13px;opacity:0.9}' +

    // Container
    '.container{max-width:600px;margin:0 auto;padding:20px}' +

    // Login Form
    '.login-card{background:white;border-radius:12px;padding:30px;box-shadow:0 2px 12px rgba(0,0,0,0.1)}' +
    '.login-card h2{color:#059669;margin-bottom:20px;text-align:center}' +
    '.field{margin-bottom:20px}' +
    '.field label{display:block;margin-bottom:8px;font-weight:600;color:#333}' +
    '.field input{width:100%;padding:12px;border:2px solid #e0e0e0;border-radius:8px;font-size:16px;transition:border-color 0.2s}' +
    '.field input:focus{outline:none;border-color:#059669}' +
    '.btn{width:100%;padding:14px;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;transition:all 0.2s}' +
    '.btn-primary{background:#059669;color:white}' +
    '.btn-primary:hover{background:#047857}' +
    '.btn-primary:disabled{background:#ccc;cursor:not-allowed}' +
    '.error{color:#dc2626;font-size:14px;margin-top:10px;text-align:center;padding:10px;background:#fee2e2;border-radius:8px}' +

    // Tabs
    '.tabs{display:flex;background:white;border-radius:12px 12px 0 0;overflow:hidden;margin-top:20px}' +
    '.tab{flex:1;padding:15px;text-align:center;font-weight:600;color:#666;background:white;border:none;cursor:pointer;transition:all 0.2s}' +
    '.tab.active{color:#059669;background:#f0fdf4;border-bottom:3px solid #059669}' +
    '.tab:hover{background:#f0fdf4}' +

    // Content
    '.content{background:white;border-radius:0 0 12px 12px;padding:20px;box-shadow:0 2px 12px rgba(0,0,0,0.1)}' +
    '.section{margin-bottom:20px}' +
    '.section-title{font-size:14px;font-weight:600;color:#059669;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #f0fdf4}' +

    // Profile display
    '.profile-field{display:flex;justify-content:space-between;padding:12px 0;border-bottom:1px solid #f0f0f0}' +
    '.profile-field:last-child{border-bottom:none}' +
    '.profile-label{color:#666;font-size:14px}' +
    '.profile-value{font-weight:500;color:#333}' +

    // Grievance cards
    '.grievance-card{background:#f8fafc;border-radius:8px;padding:15px;margin-bottom:12px;border-left:4px solid #059669}' +
    '.grievance-card.open{border-left-color:#f59e0b}' +
    '.grievance-card.closed{border-left-color:#10b981}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}' +
    '.grievance-id{font-weight:600;color:#333}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:12px;font-weight:600}' +
    '.status-open{background:#fef3c7;color:#d97706}' +
    '.status-pending{background:#e0f2fe;color:#0284c7}' +
    '.status-closed{background:#d1fae5;color:#059669}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:8px}' +

    // Edit form
    '.edit-field{margin-bottom:15px}' +
    '.edit-field label{display:block;margin-bottom:5px;font-weight:500;color:#333;font-size:14px}' +
    '.edit-field input,.edit-field select{width:100%;padding:10px;border:1px solid #e0e0e0;border-radius:6px;font-size:14px}' +
    '.btn-row{display:flex;gap:10px;margin-top:20px}' +
    '.btn-secondary{background:#e0e0e0;color:#333}' +
    '.btn-secondary:hover{background:#d0d0d0}' +
    '.btn-small{flex:1}' +

    // Loading/Success states
    '.loading{text-align:center;padding:40px;color:#666}' +
    '.spinner{width:40px;height:40px;border:4px solid #f0f0f0;border-top-color:#059669;border-radius:50%;animation:spin 1s linear infinite;margin:0 auto 15px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.success{color:#059669;background:#d1fae5;padding:15px;border-radius:8px;text-align:center;margin-bottom:15px}' +
    '.empty{text-align:center;padding:40px;color:#666}' +

    // Logout
    '.logout-btn{position:absolute;top:15px;right:15px;background:rgba(255,255,255,0.2);color:white;border:none;padding:8px 15px;border-radius:6px;cursor:pointer;font-size:13px}' +
    '.logout-btn:hover{background:rgba(255,255,255,0.3)}' +
    '.header{position:relative}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<button class="logout-btn" id="logoutBtn" style="display:none" onclick="logout()">Logout</button>' +
    '<h1>Member Self-Service Portal</h1>' +
    '<div class="subtitle">View and update your information</div>' +
    '</div>' +

    '<div class="container">' +
    // Login view
    '<div id="loginView">' +
    '<div class="login-card">' +
    '<h2>🔐 Member Login</h2>' +
    '<div class="field">' +
    '<label for="memberId">Member ID</label>' +
    '<input type="text" id="memberId" placeholder="Enter your Member ID" autocomplete="username">' +
    '</div>' +
    '<div class="field">' +
    '<label for="pin">PIN</label>' +
    '<input type="password" id="pin" placeholder="Enter your 6-digit PIN" maxlength="6" pattern="[0-9]*" inputmode="numeric" autocomplete="current-password">' +
    '</div>' +
    '<button class="btn btn-primary" id="loginBtn" onclick="login()">Login</button>' +
    '<div id="loginError" class="error" style="display:none"></div>' +
    '<p style="text-align:center;margin-top:20px;font-size:13px;color:#666">' +
    'Don\'t have a PIN? Contact your union steward.' +
    '</p>' +
    '<p style="text-align:center;margin-top:10px">' +
    '<a href="#" onclick="showResetRequest();return false" style="color:#059669;font-size:13px">Forgot your PIN?</a>' +
    '</p>' +
    '</div>' +
    '</div>' +

    // PIN Reset Request view
    '<div id="resetRequestView" style="display:none">' +
    '<div class="login-card">' +
    '<h2>🔑 Reset PIN</h2>' +
    '<p style="text-align:center;color:#666;margin-bottom:20px;font-size:14px">Enter your Member ID and we\'ll send a reset code to your registered email.</p>' +
    '<div class="field">' +
    '<label for="resetMemberId">Member ID</label>' +
    '<input type="text" id="resetMemberId" placeholder="Enter your Member ID" autocomplete="username">' +
    '</div>' +
    '<button class="btn btn-primary" id="requestResetBtn" onclick="requestReset()">Send Reset Code</button>' +
    '<div id="resetRequestError" class="error" style="display:none"></div>' +
    '<div id="resetRequestSuccess" class="success" style="display:none"></div>' +
    '<p style="text-align:center;margin-top:20px">' +
    '<a href="#" onclick="showLogin();return false" style="color:#059669;font-size:13px">← Back to Login</a>' +
    '</p>' +
    '</div>' +
    '</div>' +

    // PIN Reset Complete view (enter token + new PIN)
    '<div id="resetCompleteView" style="display:none">' +
    '<div class="login-card">' +
    '<h2>🔐 Enter New PIN</h2>' +
    '<p style="text-align:center;color:#666;margin-bottom:20px;font-size:14px">Enter the reset code from your email and choose a new PIN.</p>' +
    '<div class="field">' +
    '<label for="resetToken">Reset Code</label>' +
    '<input type="text" id="resetToken" placeholder="8-character code from email" maxlength="8" style="text-transform:uppercase">' +
    '</div>' +
    '<div class="field">' +
    '<label for="newPin">New PIN</label>' +
    '<input type="password" id="newPin" placeholder="Enter new 6-digit PIN" maxlength="6" pattern="[0-9]*" inputmode="numeric">' +
    '</div>' +
    '<div class="field">' +
    '<label for="confirmPin">Confirm PIN</label>' +
    '<input type="password" id="confirmPin" placeholder="Confirm new PIN" maxlength="6" pattern="[0-9]*" inputmode="numeric">' +
    '</div>' +
    '<button class="btn btn-primary" id="completeResetBtn" onclick="completeReset()">Set New PIN</button>' +
    '<div id="resetCompleteError" class="error" style="display:none"></div>' +
    '<div id="resetCompleteSuccess" class="success" style="display:none"></div>' +
    '<p style="text-align:center;margin-top:20px">' +
    '<a href="#" onclick="showResetRequest();return false" style="color:#059669;font-size:13px">← Request New Code</a> | ' +
    '<a href="#" onclick="showLogin();return false" style="color:#059669;font-size:13px">Back to Login</a>' +
    '</p>' +
    '</div>' +
    '</div>' +

    // Authenticated view
    '<div id="authView" style="display:none">' +
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(&apos;profile&apos;,this)">👤 My Profile</button>' +
    '<button class="tab" onclick="switchTab(&apos;grievances&apos;,this)">📋 My Cases</button>' +
    '<button class="tab" onclick="switchTab(&apos;edit&apos;,this)">✏️ Update Info</button>' +
    '</div>' +
    '<div class="content">' +

    // Profile Tab
    '<div id="tab-profile" class="tab-content">' +
    '<div id="profileContent"><div class="loading"><div class="spinner"></div>Loading profile...</div></div>' +
    '</div>' +

    // Grievances Tab
    '<div id="tab-grievances" class="tab-content" style="display:none">' +
    '<div id="grievancesContent"><div class="loading"><div class="spinner"></div>Loading cases...</div></div>' +
    '</div>' +

    // Edit Tab
    '<div id="tab-edit" class="tab-content" style="display:none">' +
    '<div id="editContent"><div class="loading"><div class="spinner"></div>Loading...</div></div>' +
    '</div>' +

    '</div>' +
    '</div>' +
    '</div>' +

    '<script>' +
    'function escapeHtml(t){if(t==null)return"";return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/\'/g,"&#x27;").replace(/\\//g,"&#x2F;");}' +
    'var sessionToken=null;' +
    'var profileData=null;' +

    // Login
    'function login(){' +
    '  var memberId=document.getElementById("memberId").value.trim();' +
    '  var pin=document.getElementById("pin").value.trim();' +
    '  if(!memberId||!pin){showLoginError("Please enter your Member ID and PIN");return}' +
    '  document.getElementById("loginBtn").disabled=true;' +
    '  document.getElementById("loginBtn").textContent="Logging in...";' +
    '  google.script.run' +
    '    .withSuccessHandler(handleLoginResult)' +
    '    .withFailureHandler(function(e){showLoginError("Login failed: "+e.message);resetLoginBtn()})' +
    '    .authenticateMember(memberId,pin);' +
    '}' +

    'function handleLoginResult(result){' +
    '  resetLoginBtn();' +
    '  if(result.success){' +
    '    sessionToken=result.sessionToken;' +
    '    document.getElementById("loginView").style.display="none";' +
    '    document.getElementById("authView").style.display="block";' +
    '    document.getElementById("logoutBtn").style.display="block";' +
    '    loadProfile();' +
    '  }else{' +
    '    showLoginError(result.error);' +
    '  }' +
    '}' +

    'function showLoginError(msg){' +
    '  var el=document.getElementById("loginError");' +
    '  el.textContent=msg;' +
    '  el.style.display="block";' +
    '}' +

    'function resetLoginBtn(){' +
    '  document.getElementById("loginBtn").disabled=false;' +
    '  document.getElementById("loginBtn").textContent="Login";' +
    '}' +

    // Logout
    'function logout(){' +
    '  sessionToken=null;' +
    '  profileData=null;' +
    '  document.getElementById("loginView").style.display="block";' +
    '  document.getElementById("authView").style.display="none";' +
    '  document.getElementById("logoutBtn").style.display="none";' +
    '  document.getElementById("loginError").style.display="none";' +
    '  document.getElementById("memberId").value="";' +
    '  document.getElementById("pin").value="";' +
    '}' +

    // PIN Reset - View switching
    'var resetMemberIdStored="";' +
    'function showLogin(){' +
    '  document.getElementById("loginView").style.display="block";' +
    '  document.getElementById("resetRequestView").style.display="none";' +
    '  document.getElementById("resetCompleteView").style.display="none";' +
    '  document.getElementById("loginError").style.display="none";' +
    '}' +
    'function showResetRequest(){' +
    '  document.getElementById("loginView").style.display="none";' +
    '  document.getElementById("resetRequestView").style.display="block";' +
    '  document.getElementById("resetCompleteView").style.display="none";' +
    '  document.getElementById("resetRequestError").style.display="none";' +
    '  document.getElementById("resetRequestSuccess").style.display="none";' +
    '}' +
    'function showResetComplete(){' +
    '  document.getElementById("loginView").style.display="none";' +
    '  document.getElementById("resetRequestView").style.display="none";' +
    '  document.getElementById("resetCompleteView").style.display="block";' +
    '  document.getElementById("resetCompleteError").style.display="none";' +
    '  document.getElementById("resetCompleteSuccess").style.display="none";' +
    '}' +

    // Request reset code
    'function requestReset(){' +
    '  var memberId=document.getElementById("resetMemberId").value.trim();' +
    '  if(!memberId){document.getElementById("resetRequestError").textContent="Please enter your Member ID";document.getElementById("resetRequestError").style.display="block";return}' +
    '  resetMemberIdStored=memberId;' +
    '  document.getElementById("requestResetBtn").disabled=true;' +
    '  document.getElementById("requestResetBtn").textContent="Sending...";' +
    '  document.getElementById("resetRequestError").style.display="none";' +
    '  document.getElementById("resetRequestSuccess").style.display="none";' +
    '  google.script.run' +
    '    .withSuccessHandler(handleResetRequest)' +
    '    .withFailureHandler(function(e){document.getElementById("resetRequestError").textContent="Error: "+e.message;document.getElementById("resetRequestError").style.display="block";resetRequestBtn()})' +
    '    .requestPINReset(memberId);' +
    '}' +
    'function handleResetRequest(result){' +
    '  resetRequestBtn();' +
    '  if(result.success){' +
    '    document.getElementById("resetRequestSuccess").textContent=result.message+" Check your email for the reset code.";' +
    '    document.getElementById("resetRequestSuccess").style.display="block";' +
    '    setTimeout(function(){showResetComplete()},2000);' +
    '  }else{' +
    '    document.getElementById("resetRequestError").textContent=result.error;' +
    '    document.getElementById("resetRequestError").style.display="block";' +
    '  }' +
    '}' +
    'function resetRequestBtn(){' +
    '  document.getElementById("requestResetBtn").disabled=false;' +
    '  document.getElementById("requestResetBtn").textContent="Send Reset Code";' +
    '}' +

    // Complete reset with token
    'function completeReset(){' +
    '  var token=document.getElementById("resetToken").value.trim().toUpperCase();' +
    '  var newPin=document.getElementById("newPin").value.trim();' +
    '  var confirmPin=document.getElementById("confirmPin").value.trim();' +
    '  var errEl=document.getElementById("resetCompleteError");' +
    '  if(!token||!newPin||!confirmPin){errEl.textContent="All fields are required";errEl.style.display="block";return}' +
    '  if(newPin!==confirmPin){errEl.textContent="PINs do not match";errEl.style.display="block";return}' +
    '  if(!/^\\d{6}$/.test(newPin)){errEl.textContent="PIN must be exactly 6 digits";errEl.style.display="block";return}' +
    '  document.getElementById("completeResetBtn").disabled=true;' +
    '  document.getElementById("completeResetBtn").textContent="Setting PIN...";' +
    '  errEl.style.display="none";' +
    '  google.script.run' +
    '    .withSuccessHandler(handleResetComplete)' +
    '    .withFailureHandler(function(e){errEl.textContent="Error: "+e.message;errEl.style.display="block";resetCompleteBtn()})' +
    '    .completePINReset(resetMemberIdStored,token,newPin);' +
    '}' +
    'function handleResetComplete(result){' +
    '  resetCompleteBtn();' +
    '  if(result.success){' +
    '    document.getElementById("resetCompleteSuccess").textContent=result.message;' +
    '    document.getElementById("resetCompleteSuccess").style.display="block";' +
    '    setTimeout(function(){showLogin();document.getElementById("memberId").value=resetMemberIdStored},2000);' +
    '  }else{' +
    '    document.getElementById("resetCompleteError").textContent=result.error;' +
    '    document.getElementById("resetCompleteError").style.display="block";' +
    '  }' +
    '}' +
    'function resetCompleteBtn(){' +
    '  document.getElementById("completeResetBtn").disabled=false;' +
    '  document.getElementById("completeResetBtn").textContent="Set New PIN";' +
    '}' +

    // Tab switching
    'function switchTab(tab,btn){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.style.display="none"});' +
    '  btn.classList.add("active");' +
    '  document.getElementById("tab-"+tab).style.display="block";' +
    '  if(tab==="grievances")loadGrievances();' +
    '  if(tab==="edit")loadEditForm();' +
    '}' +

    // Load profile
    'function loadProfile(){' +
    '  google.script.run' +
    '    .withSuccessHandler(renderProfile)' +
    '    .withFailureHandler(function(e){document.getElementById("profileContent").innerHTML="<div class=\\"error\\">Failed to load profile</div>"})' +
    '    .getMemberProfile(sessionToken);' +
    '}' +

    'function renderProfile(result){' +
    '  if(!result.success){document.getElementById("profileContent").innerHTML="<div class=\\"error\\">"+escapeHtml(result.error)+"</div>";return}' +
    '  profileData=result.profile;' +
    '  var p=result.profile;' +
    '  var html="<div class=\\"section\\"><div class=\\"section-title\\">Personal Information</div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Name</span><span class=\\"profile-value\\">"+escapeHtml(p.firstName)+" "+escapeHtml(p.lastName)+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Member ID</span><span class=\\"profile-value\\">"+escapeHtml(p.memberId)+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Job Title</span><span class=\\"profile-value\\">"+escapeHtml(p.jobTitle||"Not set")+"</span></div>";' +
    '  html+="</div>";' +
    '  html+="<div class=\\"section\\"><div class=\\"section-title\\">Work Information</div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Location</span><span class=\\"profile-value\\">"+escapeHtml(p.workLocation||"Not set")+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Unit</span><span class=\\"profile-value\\">"+escapeHtml(p.unit||"Not set")+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Assigned Steward</span><span class=\\"profile-value\\">"+escapeHtml(p.assignedSteward||"Not assigned")+"</span></div>";' +
    '  html+="</div>";' +
    '  html+="<div class=\\"section\\"><div class=\\"section-title\\">Contact Information</div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Email</span><span class=\\"profile-value\\">"+escapeHtml(p.email||"Not set")+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Phone</span><span class=\\"profile-value\\">"+escapeHtml(p.phone||"Not set")+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Preferred Contact</span><span class=\\"profile-value\\">"+escapeHtml(p.preferredComm||"Not set")+"</span></div>";' +
    '  html+="<div class=\\"profile-field\\"><span class=\\"profile-label\\">Best Time</span><span class=\\"profile-value\\">"+escapeHtml(p.bestTime||"Not set")+"</span></div>";' +
    '  html+="</div>";' +
    '  document.getElementById("profileContent").innerHTML=html;' +
    '}' +

    // Load grievances
    'function loadGrievances(){' +
    '  document.getElementById("grievancesContent").innerHTML="<div class=\\"loading\\"><div class=\\"spinner\\"></div>Loading cases...</div>";' +
    '  google.script.run' +
    '    .withSuccessHandler(renderGrievances)' +
    '    .withFailureHandler(function(e){document.getElementById("grievancesContent").innerHTML="<div class=\\"error\\">Failed to load cases</div>"})' +
    '    .getMemberGrievances(sessionToken);' +
    '}' +

    'function renderGrievances(result){' +
    '  if(!result.success){document.getElementById("grievancesContent").innerHTML="<div class=\\"error\\">"+escapeHtml(result.error)+"</div>";return}' +
    '  if(result.grievances.length===0){' +
    '    document.getElementById("grievancesContent").innerHTML="<div class=\\"empty\\"><p>📋 No grievance cases found</p><p style=\\"font-size:13px;margin-top:10px\\">If you believe you should have a case on file, please contact your steward.</p></div>";' +
    '    return;' +
    '  }' +
    '  var html="<div class=\\"section-title\\">Your Cases ("+result.count+")</div>";' +
    '  result.grievances.forEach(function(g){' +
    '    var statusClass=g.status.toLowerCase().indexOf("open")>=0?"open":(g.status.toLowerCase().indexOf("closed")>=0||g.status.toLowerCase().indexOf("resolved")>=0?"closed":"");' +
    '    var badgeClass=g.status.toLowerCase().indexOf("open")>=0?"status-open":(g.status.toLowerCase().indexOf("pending")>=0?"status-pending":"status-closed");' +
    '    html+="<div class=\\"grievance-card "+statusClass+"\\">";' +
    '    html+="<div class=\\"grievance-header\\"><span class=\\"grievance-id\\">"+escapeHtml(g.grievanceId)+"</span><span class=\\"grievance-status "+badgeClass+"\\">"+escapeHtml(g.status)+"</span></div>";' +
    '    html+="<div class=\\"grievance-detail\\"><strong>Issue:</strong> "+escapeHtml(g.issueCategory||"Not specified")+"</div>";' +
    '    html+="<div class=\\"grievance-detail\\"><strong>Current Step:</strong> "+escapeHtml(g.currentStep||"N/A")+"</div>";' +
    '    html+="<div class=\\"grievance-detail\\"><strong>Filed:</strong> "+escapeHtml(g.filedDate||"N/A")+"</div>";' +
    '    html+="<div class=\\"grievance-detail\\"><strong>Steward:</strong> "+escapeHtml(g.steward||"Not assigned")+"</div>";' +
    '    if(g.nextDeadline)html+="<div class=\\"grievance-detail\\"><strong>Next Deadline:</strong> "+escapeHtml(g.nextDeadline)+"</div>";' +
    '    if(g.resolution)html+="<div class=\\"grievance-detail\\" style=\\"margin-top:10px;padding-top:10px;border-top:1px solid #e0e0e0\\"><strong>Resolution:</strong> "+escapeHtml(g.resolution)+"</div>";' +
    '    if(g.outcome)html+="<div class=\\"grievance-detail\\"><strong>Outcome:</strong> "+escapeHtml(g.outcome)+"</div>";' +
    '    html+="</div>";' +
    '  });' +
    '  document.getElementById("grievancesContent").innerHTML=html;' +
    '}' +

    // Edit form
    'function loadEditForm(){' +
    '  if(!profileData){loadProfile();setTimeout(loadEditForm,500);return}' +
    '  var p=profileData;' +
    '  var html="<div class=\\"section-title\\">Update Contact Information</div>";' +
    '  html+="<div id=\\"updateSuccess\\" class=\\"success\\" style=\\"display:none\\">Contact information updated!</div>";' +
    '  html+="<div class=\\"edit-field\\"><label>Email</label><input type=\\"email\\" id=\\"editEmail\\" value=\\""+escapeHtml(p.email||"")+"\\"></div>";' +
    '  html+="<div class=\\"edit-field\\"><label>Phone</label><input type=\\"tel\\" id=\\"editPhone\\" value=\\""+escapeHtml(p.phone||"")+"\\"></div>";' +
    '  html+="<div class=\\"edit-field\\"><label>Preferred Contact Method</label><select id=\\"editPreferred\\"><option value=\\"\\">Select...</option><option value=\\"Email\\"'+(p.preferredComm==="Email"?" selected":"")+'>Email</option><option value=\\"Phone\\"'+(p.preferredComm==="Phone"?" selected":"")+'>Phone</option><option value=\\"Text\\"'+(p.preferredComm==="Text"?" selected":"")+'>Text</option></select></div>";' +
    '  html+="<div class=\\"edit-field\\"><label>Best Time to Reach</label><select id=\\"editBestTime\\"><option value=\\"\\">Select...</option><option value=\\"Morning\\"'+(p.bestTime==="Morning"?" selected":"")+'>Morning</option><option value=\\"Afternoon\\"'+(p.bestTime==="Afternoon"?" selected":"")+'>Afternoon</option><option value=\\"Evening\\"'+(p.bestTime==="Evening"?" selected":"")+'>Evening</option></select></div>";' +
    '  html+="<div class=\\"btn-row\\"><button class=\\"btn btn-primary btn-small\\" onclick=\\"saveChanges()\\">💾 Save Changes</button><button class=\\"btn btn-secondary btn-small\\" onclick=\\"loadEditForm()\\">↩️ Reset</button></div>";' +
    '  document.getElementById("editContent").innerHTML=html;' +
    '}' +

    'function saveChanges(){' +
    '  var updates={' +
    '    email:document.getElementById("editEmail").value,' +
    '    phone:document.getElementById("editPhone").value,' +
    '    preferredComm:document.getElementById("editPreferred").value,' +
    '    bestTime:document.getElementById("editBestTime").value' +
    '  };' +
    '  google.script.run' +
    '    .withSuccessHandler(function(result){' +
    '      if(result.success){' +
    '        document.getElementById("updateSuccess").style.display="block";' +
    '        profileData.email=updates.email;' +
    '        profileData.phone=updates.phone;' +
    '        profileData.preferredComm=updates.preferredComm;' +
    '        profileData.bestTime=updates.bestTime;' +
    '        setTimeout(function(){document.getElementById("updateSuccess").style.display="none"},3000);' +
    '      }else{alert("Error: "+result.error)}' +
    '    })' +
    '    .withFailureHandler(function(e){alert("Failed to save: "+e.message)})' +
    '    .updateMemberContact(sessionToken,updates);' +
    '}' +

    // Enter key handlers
    'document.getElementById("memberId").addEventListener("keypress",function(e){if(e.key==="Enter")document.getElementById("pin").focus()});' +
    'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter")login()});' +

    '</script></body></html>';
}

// ============================================================================
// WEB APP INTEGRATION
// ============================================================================

/**
 * Handle member self-service portal request from doGet
 * @returns {HtmlOutput} Portal HTML
 */
function getMemberSelfServicePortal() {
  var html = getMemberSelfServicePortalHtml();
  return HtmlService.createHtmlOutput(html)
    .setTitle('Member Self-Service Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DENY);
}

/**
 * Show the self-service portal URL
 */
function showSelfServicePortalUrl() {
  var ui = SpreadsheetApp.getUi();
  var baseUrl = ScriptApp.getService().getUrl();
  var portalUrl = baseUrl + '?page=selfservice';

  ui.alert(
    'Member Self-Service Portal',
    'Share this URL with members so they can access their information:\n\n' +
    portalUrl + '\n\n' +
    'Members will need:\n' +
    '1. Their Member ID\n' +
    '2. A PIN (generated by a steward)\n\n' +
    'To generate a PIN for a member, select them in the Member Directory and use:\n' +
    'Dashboard → Members → Generate Member PIN',
    ui.ButtonSet.OK
  );
}
