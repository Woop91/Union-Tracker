/**
 * Web Dashboard Data Service — Global Wrappers
 *
 * These functions are called from the client via google.script.run.
 * They delegate to DataService (defined in 21_WebDashDataService.gs).
 * Split from 21_WebDashDataService.gs for maintainability.
 *
 * Includes:
 *   - _resolveCallerEmail() / _requireStewardAuth() / _isPINSession() — auth helpers
 *   - data*() wrappers — authenticated endpoints for steward & member operations
 *   - _requireLeaderAuth() + leader endpoints — Member Leader Hub (v4.46.0)
 *
 * @fileoverview Global data service wrappers (google.script.run endpoints)
 * @requires 21_WebDashDataService.gs (DataService IIFE)
 */

// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

/**
 * Resolves the caller's verified email server-side (never trust client-supplied email).
 * Priority: (1) Session.getActiveUser() for Google SSO, (2) sessionToken for magic link auth.
 *
 * Design: SSO email takes precedence over session token. In "Execute as: Me" deployment,
 * getActiveUser() returns the visitor's Google account (if signed in). This means a user
 * who authenticates via magic link as alice@union.org but is signed into Google as
 * bob@personal.com will resolve as bob@personal.com. This is intentional — their Google
 * identity IS their identity for SSO users. Non-Google users get empty from getActiveUser()
 * and fall through to token-based resolution.
 *
 * @param {string=} sessionToken - Optional client-supplied session token
 * @returns {string} Verified email or empty string
 * @private
 */
function _resolveCallerEmail(sessionToken) {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email.toLowerCase().trim();
  } catch (_e) { log_('_e', (_e.message || _e)); }
  // Fallback: verify session token server-side — never trust plain email from client
  if (sessionToken && typeof Auth !== 'undefined' && typeof Auth.resolveEmailFromToken === 'function') {
    var tokenEmail = Auth.resolveEmailFromToken(sessionToken);
    if (tokenEmail) return tokenEmail.toLowerCase().trim();
  }
  return null;
}

/**
 * Resolves the caller's email and verifies steward role.
 * Use for all steward-only operations.
 * @param {string=} sessionToken - Optional session token for non-SSO auth
 * @returns {string|null} Steward's email if authorized, null otherwise.
 * @private
 */
function _requireStewardAuth(sessionToken) {
  if (sessionToken === null || sessionToken === undefined) return null;
  var auth = checkWebAppAuthorization('steward', sessionToken);
  if (!auth.isAuthorized) return null;
  return (auth.email || '').toLowerCase().trim();
}

// ═══════════════════════════════════════
// AUTHENTICATED DATA SERVICE WRAPPERS
// ═══════════════════════════════════════
// Security model (CR-AUTH-3):
//   - Steward ops: _requireStewardAuth(sessionToken) verifies steward role, uses server email
//   - Member self-service: _resolveCallerEmail(sessionToken) provides server-verified identity
//   - Public reads: no auth required (aggregate/non-PII data only)
//   - PIN sessions: restricted from personal data (profile, grievances, Drive, steward contact)

/**
 * Check if the session token was created via PIN login.
 * PIN sessions are restricted from viewing or editing personal data.
 * @param {string} sessionToken
 * @returns {boolean}
 * @private
 */
function _isPINSession(sessionToken) {
  if (!sessionToken) return false;
  if (typeof Auth !== 'undefined' && typeof Auth.isPINSession === 'function') {
    return Auth.isPINSession(sessionToken);
  }
  return false;
}

/** Standard response for PIN-restricted endpoints */
var _PIN_RESTRICTED_RESPONSE = { success: false, pinRestricted: true, message: 'Personal information is not available with PIN login. Sign in with Google or email link for full access.' };

/** @param {string} sessionToken @returns {Object[]} Steward cases. Requires steward auth. */
function dataGetStewardCases(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; if (!_isGrievancesEnabled()) return { success: true, cases: [] }; return DataService.getStewardCases(s); }
/** @param {string} sessionToken @returns {Object} Steward KPIs. Requires steward auth. */
function dataGetStewardKPIs(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; if (!_isGrievancesEnabled()) return { totalCases: 0, overdue: 0, dueSoon: 0, resolved: 0, activeCases: 0 }; return DataService.getStewardKPIs(s); }
/** @param {string} sessionToken @returns {Object[]} Member's own active grievances. Requires auth. PIN-restricted. */
function dataGetMemberGrievances(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: true, grievances: [] }; return DataService.getMemberGrievances(e); }
/** @param {string} sessionToken @returns {Object} Member's closed grievance history. Requires auth. PIN-restricted. */
function dataGetMemberGrievanceHistory(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: true, history: [] }; return DataService.getMemberGrievanceHistory(e); }
/** @param {string} sessionToken @param {string} stewardEmail @returns {Object|null} Steward contact info. Requires auth. PIN-restricted. */
function dataGetStewardContact(sessionToken, stewardEmail) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) { log_('dataGetStewardContact', 'auth failed'); return null; } return DataService.getStewardContact(stewardEmail || e); }

// v4.11.0 — data service wrappers (CR-AUTH-3: server-side identity + role checks)
// Steward: view any member's full profile; Member: view own profile only
// FIX-WDS-01: v4.25.8 — Parameter was named 'email' but body referenced undefined 'sessionToken'.
// Renamed first param to sessionToken; email is now second param (optional, steward override).
/**
 * Returns a member's full profile. Stewards can view any member; members see only their own.
 * @param {string} sessionToken
 * @param {string} [email] - Target email (steward override)
 * @returns {Object} Profile data or error object
 */
function dataGetFullProfile(sessionToken, email) {
  var caller = _resolveCallerEmail(sessionToken);
  if (!caller) return { success: false, message: 'Not authenticated.' };
  // PIN sessions: return minimal non-PII profile only
  if (_isPINSession(sessionToken)) {
    var limitedProfile = DataService.getFullMemberProfile(caller);
    if (!limitedProfile) return { success: false, message: 'Member not found.' };
    return { success: true, pinRestricted: true, name: limitedProfile.name, firstName: limitedProfile.firstName, lastName: limitedProfile.lastName, unit: limitedProfile.unit, role: limitedProfile.role };
  }
  var isSteward = checkWebAppAuthorization('steward', sessionToken).isAuthorized;
  // Members may only fetch their own profile; stewards can fetch any member's
  var targetEmail = (isSteward && email) ? email : caller;
  var profile = DataService.getFullMemberProfile(targetEmail);
  if (!profile) return { success: false, message: 'Member not found.' };
  profile.success = true;
  return profile;
}
// Member self-service: update own safe fields (address, workLocation, officeDays only)
// Stewards can also update member profiles; both paths use updateMemberProfile's field allowlist
/**
 * Updates member profile fields (address, location, etc.). Requires auth; locked for concurrency.
 * @param {string} sessionToken
 * @param {Object} updates - Fields to update
 * @returns {Object} { success: boolean, message: string }
 */
function dataUpdateProfile(sessionToken, updates) {
  if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE;
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  var isSteward = checkWebAppAuthorization('steward', sessionToken).isAuthorized;
  // Members can only update their own record; stewards can pass a target email via updates._targetEmail
  var targetEmail = (isSteward && updates && updates._targetEmail) ? updates._targetEmail : e;
  if (updates && updates._targetEmail) delete updates._targetEmail; // strip internal routing field
  return withScriptLock_(function() { return DataService.updateMemberProfile(targetEmail, updates); });
}
/** @param {string} sessionToken @returns {Object|null} Assigned steward info for caller. Requires auth. PIN-restricted. */
function dataGetAssignedSteward(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return DataService.getAssignedStewardInfo(e); }
/** @param {string} sessionToken @returns {Object[]} Available stewards for self-assign. Requires auth. */
function dataGetAvailableStewards(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return DataService.getAvailableStewards(e); }
/** @param {string} sessionToken @param {string} memberEmail @param {string} stewardEmail @returns {Object} Assigns steward to member. Requires steward auth. */
function dataAssignSteward(sessionToken, memberEmail, stewardEmail) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.assignStewardToMember(memberEmail, stewardEmail); }); }
// v4.28.2 — Member-safe self-assign: members can assign a steward to THEMSELVES only.
/** @param {string} sessionToken @param {string} stewardEmail @returns {Object} Self-assigns a steward. Requires auth. */
function dataMemberAssignSteward(sessionToken, stewardEmail) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; return withScriptLock_(function() { return DataService.assignStewardToMember(e, stewardEmail); }); }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Starts a grievance draft. Requires auth. PIN-restricted. */
function dataStartGrievanceDraft(sessionToken, data, idemKey) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return withScriptLock_(function() { return DataService.startGrievanceDraft(e, data, idemKey); }); }
/** @param {string} sessionToken @returns {Object} Creates Drive folder for member's grievance. Requires auth. PIN-restricted. */
function dataCreateGrievanceDrive(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return DataService.createGrievanceDriveFolder(e); }
// v4.31.1 — Resource click tracking moved to line ~5423 (3-param version with resourceTitle)
/** @param {string} sessionToken @returns {Object} Survey completion status for caller. Requires auth. */
function dataGetSurveyStatus(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return DataService.getMemberSurveyStatus(e); }
/** @param {string} sessionToken @returns {Object[]} All members from directory. Requires steward auth. */
function dataGetAllMembers(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getAllMembers(); }
/** @param {string} sessionToken @param {string} [scope] @returns {Object} Survey tracking for steward's members. Requires steward auth. */
function dataGetStewardSurveyTracking(sessionToken, scope) { var s = _requireStewardAuth(sessionToken); if (!s) return { total: 0, completed: 0, members: [] }; try { return DataService.getStewardSurveyTracking(s, scope); } catch (e) { log_('dataGetStewardSurveyTracking error', e.message + '\n' + (e.stack || '')); return { total: 0, completed: 0, members: [] }; } }
/** @param {string} sessionToken @param {Object} filter @param {string} msg @param {string} subject @returns {Object} Sends broadcast email. Requires steward auth. */
function dataSendBroadcast(sessionToken, filter, msg, subject) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.sendBroadcastMessage(s, filter, msg, subject); }
/** @param {string} sessionToken @returns {Object} Aggregated survey results. Requires steward auth. */
function dataGetSurveyResults(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getSurveyResults(); }
// v4.21.0 — Native survey engine wrappers
/** @param {string} sessionToken @returns {Object} Survey questions. Requires auth. */
function dataGetSurveyQuestions(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return getSurveyQuestions(); }
/** @param {string} sessionToken @param {Object} responses @returns {Object} Submits survey response. Requires auth. */
function dataSubmitSurveyResponse(sessionToken, responses) { var e = _resolveCallerEmail(sessionToken); return e ? submitSurveyResponse(e, responses) : { success: false, message: 'Not authenticated.' }; }
// dataGetPendingSurveyMembers, dataGetSatisfactionSummary, dataOpenNewSurveyPeriod are in 08e_SurveyEngine.gs
/** @param {string} sessionToken @param {string} memberEmail @param {string} type @param {string} notes @param {string} duration @param {string} memberName @returns {Object} Logs member contact. Requires steward auth. */
function dataLogMemberContact(sessionToken, memberEmail, type, notes, duration, memberName) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.logMemberContact(s, memberEmail, type, notes, duration, memberName); }); }
/** @param {string} sessionToken @param {string} memberEmail @returns {Object[]} Contact history for a member. Requires steward auth. */
function dataGetMemberContactHistory(sessionToken, memberEmail) { var s = _requireStewardAuth(sessionToken); if (!s) { log_('dataGetMemberContactHistory', 'auth failed'); return []; } return DataService.getMemberContactHistory(s, memberEmail); }
/** @param {string} sessionToken @returns {Object[]} Full contact log for a steward. Requires steward auth. */
function dataGetStewardContactLog(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) { log_('dataGetStewardContactLog', 'auth failed'); return []; } return DataService.getStewardContactLog(s); }

// S2: Batch badge counts — replaces 3 serial client calls with 1 round-trip
/**
 * Returns notification, task, and Q&A badge counts in a single round-trip. Requires auth.
 * @param {string} sessionToken
 * @returns {Object} { notificationCount, taskCount, overdueTaskCount, qaUnansweredCount }
 */
function dataGetBadgeCounts(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { notificationCount: 0, taskCount: 0, overdueTaskCount: 0, qaUnansweredCount: 0 };
  var role = 'member';
  var auth = checkWebAppAuthorization('steward', sessionToken);
  if (auth.isAuthorized) role = 'steward';
  return DataService.getBadgeCounts(e, role);
}

/**
 * Sends a direct email to a single member and logs it to Drive contact sheet. Requires steward auth.
 * @param {string} sessionToken
 * @param {string} memberEmail
 * @param {string} subject
 * @param {string} body
 * @returns {Object} { success: boolean, message: string }
 */
function dataSendDirectMessage(sessionToken, memberEmail, subject, body) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!memberEmail || !subject || !body) return { success: false, message: 'Missing required fields.' };
  try {
    var config = (typeof ConfigReader !== 'undefined') ? ConfigReader.getConfig() : {};
    var fullSubject = (config.orgAbbrev || 'Union') + ' \u2014 ' + subject;
    MailApp.sendEmail(memberEmail.trim(), fullSubject, body);
    if (typeof logAuditEvent === 'function') logAuditEvent('DIRECT_MESSAGE_SENT', { steward: s, member: memberEmail, subject: subject });

    // Log to per-member Drive contact sheet — type 'Email', notes = subject + body preview
    try {
      var sRecord = (typeof findUserByEmail === 'function') ? findUserByEmail(s) : null;
      var sName   = (sRecord && sRecord.name) ? sRecord.name : s;
      var adminResult  = DataService.getOrCreateMemberContactFolderPublic(memberEmail.trim().toLowerCase());
      var memberFolder = (adminResult && adminResult.masterFolder) ? adminResult.masterFolder : null;
      if (memberFolder) {
        var folderName = memberFolder.getName();
        var contactSS  = DataService.getOrCreateMemberContactSheetPublic(memberFolder, folderName);
        if (contactSS) {
          var logSheet = contactSS.getSheetByName(CONTACT_SHEET_TAB_) || contactSS.getActiveSheet();
          var noteText = 'Subject: ' + subject + (body ? ' | ' + String(body).substring(0, 300) : '');
          logSheet.appendRow([new Date(), sName, 'Email', noteText, '']);
        }
      }
    } catch (driveErr) {
      log_('dataSendDirectMessage Drive log error', driveErr.message);
      // Non-fatal — email already sent
    }

    return { success: true, message: 'Message sent.' };
  } catch (e) {
    log_('dataSendDirectMessage error', e.message);
    return { success: false, message: 'Failed to send: ' + e.message };
  }
}

/**
 * Returns the Drive folder URL for a member's active (non-resolved) grievance.
 * Steward-only — requires steward auth via _requireStewardAuth().
 *
 * DEPENDENCY: This function reads target.driveFolderUrl from the grievance
 * record built by _buildGrievanceRecord(). That field was wired in v4.32.1
 * via HEADERS.grievanceDriveFolderUrl → GRIEVANCE_COLS.DRIVE_FOLDER_URL (col 33).
 * Before v4.32.1, driveFolderUrl was never populated in the record, so this
 * function always returned { success: false, message: 'No Drive folder...' }.
 *
 * IF THIS BREAKS: Returns { success: false, url: null }. Steward sees
 * "No Drive folder linked to this case" in the UI. Non-destructive.
 *
 * @param {string} sessionToken — steward session token
 * @param {string} memberEmail — member whose grievance folder to look up
 * @returns {{ success: boolean, url: string|null, grievanceId?: string, message?: string }}
 */
function dataGetMemberCaseFolderUrl(sessionToken, memberEmail) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, url: null, message: 'Steward access required.' };
  if (!memberEmail) return { success: false, url: null, message: 'Member email required.' };
  try {
    var grievances = DataService.getMemberGrievances(memberEmail.trim().toLowerCase());
    if (!grievances || grievances.length === 0) return { success: false, url: null, message: 'No grievances found.' };
    var active = grievances.find(function(g) {
      var st = (g.status || '').toLowerCase();
      return st !== 'resolved' && st !== 'closed' && st !== 'withdrawn' && st !== 'denied';
    });
    var target = active || grievances[0];
    if (target.driveFolderUrl) return { success: true, url: target.driveFolderUrl, grievanceId: target.grievanceId };
    return { success: false, url: null, message: 'No Drive folder linked to this case.' };
  } catch (e) {
    log_('dataGetMemberCaseFolderUrl error', e.message);
    return { success: false, url: null, message: 'Error fetching case folder.' };
  }
}
// A4: LockService for concurrent write safety
/** @param {string} sessionToken @param {string} title @param {string} desc @param {string} memberEmail @param {string} priority @param {string} dueDate @param {string} assignToEmail @param {string} idemKey @returns {Object} Creates a steward task. Requires steward auth. */
function dataCreateTask(sessionToken, title, desc, memberEmail, priority, dueDate, assignToEmail, idemKey) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assignToEmail || '', idemKey); }); }
/**
 * Creates a task assigned to a specific steward. Requires chief steward auth.
 * @param {string} sessionToken
 * @param {string} assigneeEmail - Target steward
 * @param {string} title
 * @param {string} desc
 * @param {string} memberEmail
 * @param {string} priority
 * @param {string} dueDate
 * @param {string} idemKey
 * @returns {Object} { success: boolean, message: string }
 */
function dataCreateTaskForSteward(sessionToken, assigneeEmail, title, desc, memberEmail, priority, dueDate, idemKey) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!DataService.isChiefSteward(s)) return { success: false, message: 'Not authorized.' };
  return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assigneeEmail, idemKey);
}
/** @param {string} sessionToken @param {string} [statusFilter] @returns {Object[]} Steward tasks. Requires steward auth. */
function dataGetTasks(sessionToken, statusFilter) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return DataService.getTasks(s, statusFilter); }
/** @param {string} sessionToken @param {string} taskId @returns {Object} Completes a steward task. Requires steward auth. */
function dataCompleteTask(sessionToken, taskId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.completeTask(s, taskId); }); }
/** @param {string} sessionToken @returns {Object} Member stats for steward's caseload. Requires auth. */
function dataGetStewardMemberStats(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return {}; try { return DataService.getStewardMemberStats(e); } catch (err) { log_('dataGetStewardMemberStats error', err.message + '\n' + (err.stack || '')); return { total: 0, byLocation: {}, byDues: {} }; } }
/**
 * Returns the steward directory with phone visibility based on caller's role. Requires auth.
 * @param {string} sessionToken
 * @returns {Object[]} Array of steward contact entries
 */
function dataGetStewardDirectory(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return [];
  try {
    var callerRec = DataService.findUserByEmail(e);
    var callerIsSteward = callerRec && (callerRec.isSteward === true);
    return DataService.getStewardDirectory(callerIsSteward);
  } catch (err) {
    log_('dataGetStewardDirectory error', err.message + '\n' + (err.stack || ''));
    return [];
  }
}
// ═══════════════════════════════════════
// NON-MEMBER CONTACTS (v4.48.0)
// Steward-only CRUD for external contacts (management, legal, HR, allies).
// ═══════════════════════════════════════

/** @param {string} sessionToken @returns {Object[]} Non-member contacts. Steward-only. */
function dataGetNonMemberContacts(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return []; try { return DataService.getNonMemberContacts(); } catch (err) { log_('dataGetNonMemberContacts error', err.message); return []; } }
/** @param {string} sessionToken @param {Object} data @returns {Object} Add non-member contact. Steward-only. */
function dataAddNonMemberContact(sessionToken, data) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, error: 'Not authorized.' }; return withScriptLock_(function() { try { return DataService.addNonMemberContact(data); } catch (err) { log_('dataAddNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }
/** @param {string} sessionToken @param {string} contactId @param {Object} data @returns {Object} Update non-member contact. Steward-only. */
function dataUpdateNonMemberContact(sessionToken, contactId, data) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, error: 'Not authorized.' }; return withScriptLock_(function() { try { return DataService.updateNonMemberContact(contactId, data); } catch (err) { log_('dataUpdateNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }
/** @param {string} sessionToken @param {string} contactId @returns {Object} Delete non-member contact. Steward-only. */
function dataDeleteNonMemberContact(sessionToken, contactId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, error: 'Not authorized.' }; return withScriptLock_(function() { try { return DataService.deleteNonMemberContact(contactId); } catch (err) { log_('dataDeleteNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }

// ─── ADD MEMBER (webapp) ──────────────────────────────────────────────────────
/** @param {string} sessionToken @returns {Object} Config-driven dropdown options for the Add Member form. Steward-only. */
function dataGetAddMemberOptions(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, error: 'Not authorized.' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'Spreadsheet not found.' };
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    function _configList(col) {
      if (!configSheet || !col || col < 1) return [];
      try {
        var lr = configSheet.getLastRow();
        if (lr < 3) return [];
        return configSheet.getRange(3, col, lr - 2, 1).getValues()
          .map(function(r) { return String(r[0]).trim(); })
          .filter(function(v) { return v !== ''; });
      } catch (_e) { return []; }
    }
    return {
      success: true,
      jobTitles:     _configList(CONFIG_COLS.JOB_TITLES),
      locations:     _configList(CONFIG_COLS.OFFICE_LOCATIONS),
      units:         _configList(CONFIG_COLS.UNITS),
      supervisors:   _configList(CONFIG_COLS.SUPERVISORS),
      managers:      _configList(CONFIG_COLS.MANAGERS),
      duesStatuses:  _configList(CONFIG_COLS.DUES_STATUSES)
    };
  } catch (err) {
    log_('dataGetAddMemberOptions error', err.message);
    return { success: false, error: err.message };
  }
}

/**
 * @param {string} sessionToken
 * @param {Object} memberData  — { firstName, lastName, email, phone, jobTitle, workLocation,
 *                                 unit, supervisor, manager, employeeId, hireDate, duesStatus }
 * @returns {Object} { success, memberId, message } or { success: false, error }
 * Steward-only. Calls addMember() in 02_DataManagers.gs via withScriptLock_.
 */
function dataAddMemberFromWebapp(sessionToken, memberData) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, error: 'Not authorized.' };
  if (!memberData || !memberData.firstName || !memberData.lastName) {
    return { success: false, error: 'First and Last Name are required.' };
  }

  // Input validation — reject bad data before touching the sheet
  if (!memberData.firstName.trim()) {
    return { success: false, error: 'First name is required.' };
  }
  if (!memberData.lastName.trim()) {
    return { success: false, error: 'Last name is required.' };
  }
  if (memberData.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(memberData.email).trim())) {
    return { success: false, error: 'Valid email address is required.' };
  }
  if (memberData.email && String(memberData.email).trim().length > 254) {
    return { success: false, error: 'Email address too long.' };
  }
  // Sanitize string lengths to prevent oversized cell writes
  if (memberData.firstName.length > 100) memberData.firstName = memberData.firstName.substring(0, 100);
  if (memberData.lastName.length > 100) memberData.lastName = memberData.lastName.substring(0, 100);

  return withScriptLock_(function() {
    try {
      var memberId = addMember({
        firstName:    String(memberData.firstName  || '').trim(),
        lastName:     String(memberData.lastName   || '').trim(),
        email:        String(memberData.email       || '').trim().toLowerCase(),
        phone:        String(memberData.phone       || '').trim(),
        jobTitle:     String(memberData.jobTitle    || '').trim(),
        workLocation: String(memberData.workLocation|| '').trim(),
        unit:         String(memberData.unit        || '').trim(),
        supervisor:   String(memberData.supervisor  || '').trim(),
        manager:      String(memberData.manager     || '').trim(),
        employeeId:   String(memberData.employeeId  || '').trim(),
        hireDate:     memberData.hireDate ? memberData.hireDate : '',
        duesStatus:   String(memberData.duesStatus  || 'Current').trim()
      });
      logAuditEvent(AUDIT_EVENTS.MEMBER_ADDED, {
        memberId: memberId,
        name: memberData.firstName + ' ' + memberData.lastName,
        addedBy: s
      });
      // Programmatic writes don't trigger onEdit, so bidirectional sync
      // (syncDropdownToConfig_) never fires. Explicitly sync dropdown values
      // to Config so they appear in future Add Member dropdowns.
      try {
        var _ddFields = [
          { val: memberData.jobTitle,     col: CONFIG_COLS.JOB_TITLES },
          { val: memberData.workLocation, col: CONFIG_COLS.OFFICE_LOCATIONS },
          { val: memberData.unit,         col: CONFIG_COLS.UNITS },
          { val: memberData.supervisor,   col: CONFIG_COLS.SUPERVISORS },
          { val: memberData.manager,      col: CONFIG_COLS.MANAGERS },
          { val: memberData.duesStatus,   col: CONFIG_COLS.DUES_STATUSES }
        ];
        for (var _d = 0; _d < _ddFields.length; _d++) {
          var _v = String(_ddFields[_d].val || '').trim();
          if (_v && _ddFields[_d].col) addToConfigDropdown_(_ddFields[_d].col, _v);
        }
      } catch (_syncErr) {
        log_('dataAddMemberFromWebapp config sync', _syncErr.message);
      }
      return { success: true, memberId: memberId, message: 'Member added successfully.' };
    } catch (err) {
      log_('dataAddMemberFromWebapp error', err.message);
      return { success: false, error: err.message };
    }
  });
}

/** @param {string} sessionToken @param {string} memberEmail @param {Object} updates @returns {Object} Steward updates member directory fields. Steward-only. */
function dataUpdateMemberBySteward(sessionToken, memberEmail, updates) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; if (!memberEmail) return { success: false, message: 'Member email required.' }; return withScriptLock_(function() { return DataService.updateMemberBySteward(memberEmail, updates); }); }

/** @param {string} sessionToken @returns {Object} Org-wide grievance statistics. Requires steward auth. */
function dataGetGrievanceStats(sessionToken) { if (!_isGrievancesEnabled()) return { available: false }; var s = _requireStewardAuth(sessionToken); if (!s) return { available: false }; return DataService.getGrievanceStats(); }
/** @param {string} sessionToken @returns {Object[]} Grievance hotspot locations. Requires steward auth. */
function dataGetGrievanceHotSpots(sessionToken) { if (!_isGrievancesEnabled()) return []; var s = _requireStewardAuth(sessionToken); if (!s) return []; return DataService.getGrievanceHotSpots(); }
/** @param {string} sessionToken @returns {Object|null} Membership statistics. Requires auth. */
function dataGetMembershipStats(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMembershipStats() : null; }

// v4.28.1 — Member-safe grievance endpoints for Union Stats page.
// Uses _resolveCallerEmail (any authenticated member) instead of _requireStewardAuth.
// Data is already anonymized (aggregate counts only); hotspots require 3+ per location.
/** @param {string} sessionToken @returns {Object} Anonymized grievance stats (member-safe). Requires auth. */
function dataGetMemberGrievanceStats(sessionToken) { if (!_isGrievancesEnabled()) return { available: false }; var e = _resolveCallerEmail(sessionToken); if (!e) return { available: false }; return DataService.getGrievanceStats(); }
/** @param {string} sessionToken @returns {Object[]} Anonymized grievance hotspots (member-safe). Requires auth. */
function dataGetMemberGrievanceHotSpots(sessionToken) { if (!_isGrievancesEnabled()) return []; var e = _resolveCallerEmail(sessionToken); if (!e) return []; return DataService.getGrievanceHotSpots(); }

// v4.32.0 — Grievance Feedback wrappers
/** @param {string} sessionToken @returns {Object|null} Pending grievance feedback prompt for caller. Requires auth. */
function dataGetPendingGrievanceFeedback(sessionToken) { if (!_isGrievancesEnabled()) return null; var e = _resolveCallerEmail(sessionToken); return e ? DataService.getPendingGrievanceFeedback(e) : null; }
/** @param {string} sessionToken @param {string} grievanceId @param {Object} ratings @param {string} comment @returns {Object} Submits grievance feedback. Requires auth. */
function dataSubmitGrievanceFeedback(sessionToken, grievanceId, ratings, comment) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return withScriptLock_(function() { return DataService.submitGrievanceFeedback(e, grievanceId, ratings, comment); }); }
/** @param {string} sessionToken @returns {Object|null} Aggregate grievance feedback stats. Requires auth. */
function dataGetGrievanceFeedbackStats(sessionToken) { if (!_isGrievancesEnabled()) return null; var e = _resolveCallerEmail(sessionToken); return e ? DataService.getGrievanceFeedbackStats() : null; }
/** @param {string} sessionToken @returns {Object|null} Feedback summary for calling steward. Requires steward auth. */
function dataGetStewardFeedbackSummary(sessionToken) { var s = _requireStewardAuth(sessionToken); return s ? DataService.getStewardFeedbackSummary(s) : null; }
/** @param {string} sessionToken @param {number} [limit] @returns {Object[]} Upcoming events. Requires auth. */
function dataGetUpcomingEvents(sessionToken, limit) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getUpcomingEvents(limit) : []; }
// dataGetSurveyQuestions and dataSubmitSurveyResponse are defined in the v4.21.0 block above (single canonical definition)
/** @param {string} sessionToken @returns {boolean} Whether caller is the chief steward. Requires auth. */
function dataIsChiefSteward(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.isChiefSteward(e) : false; }
// dataGetAgencyGrievanceStats — alias removed; frontend uses dataGetGrievanceStats directly

// v4.17.0 — member task assignment wrappers (CR-AUTH-3: server-side identity)
/** @param {string} sessionToken @param {string} memberEmail @param {string} title @param {string} desc @param {string} priority @param {string} dueDate @returns {Object} Creates member task. Requires steward auth. */
function dataCreateMemberTask(sessionToken, memberEmail, title, desc, priority, dueDate) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.createMemberTask(s, memberEmail, title, desc, priority, dueDate); }); }
/** @param {string} sessionToken @param {string} [statusFilter] @returns {Object[]} Tasks assigned to the calling member. Requires auth. */
function dataGetMemberTasks(sessionToken, statusFilter) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMemberTasks(e, statusFilter) : []; }
/** @param {string} sessionToken @param {string} taskId @returns {Object} Completes a member task. Requires auth. */
function dataCompleteMemberTask(sessionToken, taskId) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.completeMemberTask(e, taskId) : { success: false, message: 'Not authenticated.' }; }
/** @param {string} sessionToken @returns {Object[]} Member tasks assigned by calling steward. Requires steward auth. */
function dataGetStewardAssignedMemberTasks(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return DataService.getStewardAssignedMemberTasks(s); }
// BUG-TASKS-03: steward completing a member task on the member's behalf
/** @param {string} sessionToken @param {string} taskId @returns {Object} Steward marks member task complete. Requires steward auth. */
function dataStaffCompleteMemberTask(sessionToken, taskId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.stewardCompleteMemberTask(s, taskId); }

// v4.16.0 — unwired sheet wrappers (CR-AUTH-3: server-side identity + role checks)
/** @param {string} sessionToken @param {string} taskId @param {Object} updates @returns {Object} Updates a steward task. Requires steward auth. */
function dataUpdateTask(sessionToken, taskId, updates) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.updateTask(s, taskId, updates); }); }
/** @param {string} sessionToken @returns {Object[]} All steward performance metrics. Requires steward auth. */
function dataGetAllStewardPerformance(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return DataService.getAllStewardPerformance(); }
/** @param {string} sessionToken @param {string} caseId @returns {Object[]} Checklist items for a case. Requires auth. */
function dataGetCaseChecklist(sessionToken, caseId) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getCaseChecklist(caseId) : []; }
/** @param {string} sessionToken @param {string} checklistId @param {boolean} completed @returns {Object} Toggles checklist item. Requires auth. */
function dataToggleChecklistItem(sessionToken, checklistId, completed) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.toggleChecklistItem(checklistId, completed, e) : { success: false, message: 'Not authenticated.' }; }
/** @param {string} sessionToken @returns {Object[]} Meetings the caller has attended. Requires auth. */
function dataGetMemberMeetings(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMemberMeetings(e) : []; }
/** @param {string} sessionToken @returns {Object} Satisfaction survey trends. Requires steward auth. */
function dataGetSatisfactionTrends(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { categories: [] }; return DataService.getSatisfactionTrends(); }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Submits user feedback. Requires auth. */
function dataSubmitFeedback(sessionToken, data, idemKey) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.submitFeedback(e, data, idemKey) : { success: false, message: 'Not authenticated.' }; }
/** @param {string} sessionToken @returns {Object[]} Caller's submitted feedback. Requires auth. */
function dataGetMyFeedback(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMyFeedback(e) : []; }

// v4.33.0 — Insights batch: 6 parallel server calls in 1 round-trip
/**
 * Returns all Insights page data in a single round-trip. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { stats, hotSpots, perf, sat, memberStats, workload }
 */
function dataGetInsightsBatch(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { stats: { available: false }, hotSpots: [], perf: [], sat: { categories: [] }, memberStats: null, workload: { available: false } };
  var result = {};
  try { result.stats = DataService.getGrievanceStats(); } catch (_e) { result.stats = { available: false }; log_('InsightsBatch stats', _e.message); }
  try { result.hotSpots = DataService.getGrievanceHotSpots(); } catch (_e) { result.hotSpots = []; log_('InsightsBatch hotSpots', _e.message); }
  try { result.perf = DataService.getAllStewardPerformance(); } catch (_e) { result.perf = []; log_('InsightsBatch perf', _e.message); }
  try { result.sat = DataService.getSatisfactionTrends(); } catch (_e) { result.sat = { categories: [] }; log_('InsightsBatch sat', _e.message); }
  try { result.memberStats = DataService.getMembershipStats(); } catch (_e) { result.memberStats = null; log_('InsightsBatch memberStats', _e.message); }
  result.workload = { available: false };
  return result;
}

// v4.33.0 — Nav refresh batch: KPIs + badge counts in 1 round-trip
/**
 * Returns steward KPIs and badge counts in a single round-trip. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { kpis: Object|null, badges: Object }
 */
function dataRefreshNavData(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { kpis: null, badges: { notificationCount: 0, taskCount: 0, overdueTaskCount: 0, qaUnansweredCount: 0 } };
  var kpis = null;
  try { kpis = DataService.getStewardKPIs(s); } catch (_e) { log_('RefreshNavData kpis', _e.message); }
  var badges = DataService.getBadgeCounts(s, 'steward');
  return { kpis: kpis, badges: badges };
}

// v4.33.0 — Grievance e-signature + form option wrappers
/**
 * Retrieves grievance data for e-signature workflow.
 * @param {string} sigToken - Signature token from the e-sign URL
 * @returns {Object} Grievance data for signing or error object
 */
function dataGetGrievanceForSigning(sigToken) {
  return getGrievanceForSigning(sigToken);
}

/**
 * Submits a grievance e-signature.
 * @param {string} sigToken - Signature token from the e-sign URL
 * @param {string} sigBase64 - Base64-encoded signature image
 * @returns {Object} Result object with success status
 */
function dataSubmitGrievanceSignature(sigToken, sigBase64) {
  return submitGrievanceSignature(sigToken, sigBase64);
}

/**
 * Returns grievance form dropdown options (steps, statuses, categories).
 * @param {string} sessionToken - Session token for auth
 * @returns {Object} Form options or error object
 */
function dataGetGrievanceFormOptions(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' };
  return getGrievanceFormOptions();
}

/**
 * Authenticated wrapper for initiateGrievance().
 * Called by renderNewGrievanceForm() in steward_view.html when a steward submits
 * the New Grievance intake form. Requires steward-level session.
 * @param {string} sessionToken - Active steward session token
 * @param {Object} data - Grievance payload from _collectFormData():
 *   { memberEmail, step, incidentDate, issueCategory, articles, description, remedy, formOverrides }
 * @param {string} idemKey - Idempotency key (format: GRV_<timestamp>_<random>)
 * @returns {Object} { success, grievanceId, driveFolderUrl, memberName, message } on success
 *                  { success: false, message } on failure
 *                  { duplicate: true, message } if idemKey already processed
 */
function dataInitiateGrievance(sessionToken, data, idemKey) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' };
  return withScriptLock_(function() { return initiateGrievance(s, data, idemKey); });
}

/** @param {string} sessionToken @param {number} [limit] @returns {Object[]} Meeting minutes. Requires auth. */
function dataGetMeetingMinutes(sessionToken, limit) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMeetingMinutes(limit) : []; }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Adds meeting minutes. Requires steward auth. */
function dataAddMeetingMinutes(sessionToken, data, idemKey) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.addMeetingMinutes(s, data, idemKey); }

/**
 * BACKFILL: Generates Drive docs for any existing MeetingMinutes rows
 * that pre-date v4.20.18 and therefore have an empty DriveDocUrl column.
 *
 * Run once from Apps Script editor or from the menu after upgrading to v4.20.18.
 * Safe to re-run — skips rows that already have a URL.
 * Processes up to 50 rows per call to avoid the 6-min GAS timeout.
 *
 * @returns {Object} { processed, skipped, errors, message }
 */
function BACKFILL_MINUTES_DRIVE_DOCS() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { log_('_e', (_e.message || _e)); }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { processed: 0, skipped: 0, errors: 0, message: 'No active spreadsheet.' };
  var sheet = (typeof getOrCreateMinutesSheet === 'function') ? getOrCreateMinutesSheet() : null;
  if (!sheet) {
    var msg = 'MeetingMinutes sheet not found \u2014 nothing to backfill.';
    if (ui) ui.alert('\uD83D\uDCC4 Minutes Backfill', msg, ui.ButtonSet.OK);
    return { processed: 0, skipped: 0, errors: 0, message: msg };
  }

  // Resolve Minutes/ Drive folder ID
  var minutesFolderId = '';
  try {
    if (typeof getConfigValue_ === 'function' && typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.MINUTES_FOLDER_ID) {
      minutesFolderId = getConfigValue_(CONFIG_COLS.MINUTES_FOLDER_ID) || '';
    }
    if (!minutesFolderId) {
      minutesFolderId = PropertiesService.getScriptProperties().getProperty('MINUTES_FOLDER_ID') || '';
    }
  } catch (_re) { log_('_re', (_re.message || _re)); }

  var data      = sheet.getDataRange().getValues();
  var tz        = Session.getScriptTimeZone();
  var totalRows = data.length - 1; // rows excluding header

  // Pre-scan: count rows that still need a doc so progress toasts show X of Y
  var needsDoc = 0;
  for (var c = 1; c < data.length; c++) {
    if (!String(data[c][PORTAL_MINUTES_COLS.DRIVE_DOC_URL] || '').trim()) needsDoc++;
  }

  if (needsDoc === 0) {
    var allDoneMsg = '\u2705 All ' + totalRows + ' rows already have Drive doc URLs \u2014 nothing to do.';
    if (ui) ui.alert('\uD83D\uDCC4 Minutes Backfill', allDoneMsg, ui.ButtonSet.OK);
    else ss.toast(allDoneMsg, '\uD83D\uDCC4 Minutes Backfill', 6);
    return { processed: 0, skipped: totalRows, errors: 0, message: allDoneMsg };
  }

  ss.toast('Starting backfill of ' + needsDoc + ' rows\u2026', '\uD83D\uDCC4 Minutes Backfill', 5);

  var processed = 0, skipped = 0, errors = 0;
  // Flush every FLUSH_EVERY docs: commits URL writes to the sheet so that any
  // GAS 6-minute timeout preserves work done so far. Re-running the function
  // safely skips already-written rows and continues from where it left off.
  var FLUSH_EVERY = 10;

  for (var i = 1; i < data.length; i++) {
    var existingUrl = String(data[i][PORTAL_MINUTES_COLS.DRIVE_DOC_URL] || '').trim();
    if (existingUrl) { skipped++; continue; }

    var meetingDate = data[i][PORTAL_MINUTES_COLS.MEETING_DATE];
    var title       = String(data[i][PORTAL_MINUTES_COLS.TITLE]        || '(Untitled)');
    var bullets     = String(data[i][PORTAL_MINUTES_COLS.BULLETS]      || '');
    var fullMins    = String(data[i][PORTAL_MINUTES_COLS.FULL_MINUTES]  || '');
    var createdBy   = String(data[i][PORTAL_MINUTES_COLS.CREATED_BY]   || 'unknown');

    if (isNaN(new Date(meetingDate).getTime())) meetingDate = new Date();
    var dateStr  = Utilities.formatDate(new Date(meetingDate), tz, 'yyyy-MM-dd');
    var docTitle = title + ' \u2014 ' + dateStr;

    try {
      var doc  = DocumentApp.create(docTitle);
      var body = doc.getBody();
      body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('Recorded by: ' + createdBy);
      body.appendParagraph('Date: ' + Utilities.formatDate(new Date(meetingDate), tz, 'MMMM d, yyyy'));
      body.appendParagraph('');
      if (bullets) {
        body.appendParagraph('Key Points').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        bullets.split('\n').forEach(function(line) {
          if (line.trim()) body.appendListItem(line.replace(/^[-\u2022*]\s*/, '').trim());
        });
        body.appendParagraph('');
      }
      if (fullMins) {
        body.appendParagraph('Full Minutes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(fullMins);
      }
      doc.saveAndClose();

      var docFile = DriveApp.getFileById(doc.getId());
      if (minutesFolderId) {
        try {
          DriveApp.getFolderById(minutesFolderId).addFile(docFile);
          DriveApp.getRootFolder().removeFile(docFile);
        } catch (_mv) { log_('_mv', (_mv.message || _mv)); }
      }

      // Write URL back to sheet immediately (0-indexed col + 1 = 1-indexed for getRange)
      sheet.getRange(i + 1, PORTAL_MINUTES_COLS.DRIVE_DOC_URL + 1).setValue(docFile.getUrl());
      processed++;

      // Progress checkpoint: flush writes + show X-of-Y toast every FLUSH_EVERY docs.
      // SpreadsheetApp.flush() ensures partial progress survives a GAS 6-min timeout.
      if (processed % FLUSH_EVERY === 0) {
        SpreadsheetApp.flush();
        ss.toast('Created ' + processed + ' of ' + needsDoc + ' docs\u2026', '\uD83D\uDCC4 Minutes Backfill', 3);
      }

    } catch (docErr) {
      log_('BACKFILL_MINUTES_DRIVE_DOCS', 'BACKFILL_MINUTES_DRIVE_DOCS row ' + i + ': ' + docErr.message);
      errors++;
    }
  }

  // Final flush \u2014 commit the last partial batch before showing the result dialog
  SpreadsheetApp.flush();

  var summary = '\u2705 Backfill complete!\n\n' +
    'Docs created:              ' + processed + '\n' +
    'Already had URL (skipped): ' + skipped   + '\n' +
    'Errors:                    ' + errors +
    (errors > 0 ? '\n\nCheck Apps Script logs (Extensions > Apps Script > Executions) for details.' : '');

  if (ui) ui.alert('\uD83D\uDCC4 Minutes Backfill', summary, ui.ButtonSet.OK);
  else ss.toast('\u2705 Backfill done: ' + processed + ' created, ' + errors + ' errors.', '\uD83D\uDCC4 Minutes Backfill', 8);
  log_('BACKFILL_MINUTES_DRIVE_DOCS', 'processed=' + processed + ' skipped=' + skipped + ' errors=' + errors);
  return { processed: processed, skipped: skipped, errors: errors, message: summary };
}
// OPT-1: Dedicated steward dashboard init — single round-trip combining cases, KPIs, badges, member count.
// Used as fallback when the preloaded batch from dataGetBatchData is unavailable.
/**
 * Returns steward dashboard init data (cases, KPIs, badges, member count) in one call. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { cases, kpis, badges, memberCount }
 */
function dataGetStewardDashboardInit(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { cases: [], kpis: {}, badges: { notificationCount: 0, taskCount: 0, overdueTaskCount: 0, qaUnansweredCount: 0 }, memberCount: 0 };
  var batch = DataService.getBatchData(s, 'steward');
  return {
    cases: batch.cases || [],
    kpis: batch.kpis || {},
    badges: {
      notificationCount: batch.notificationCount || 0,
      taskCount: batch.taskCount || 0,
      overdueTaskCount: batch.overdueTaskCount || 0,
      qaUnansweredCount: batch.qaUnansweredCount || 0,
    },
    memberCount: batch.memberCount || 0,
  };
}

// Batch data fetch — single round-trip for SPA init (CR-AUTH-3: server-side identity + role)
// Role is re-verified server-side from the Member Directory; client-supplied role is ignored.
/**
 * Returns all data needed for the initial SPA render in one round-trip. Requires auth.
 * @param {string} sessionToken
 * @returns {Object} Batch data payload (role determined server-side)
 */
function dataGetBatchData(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return {};
  // Re-derive role from directory — never trust the client-supplied value
  var serverRole = DataService.getUserRole(e) || 'member';
  // Normalize 'both' → steward view (steward functions are a superset)
  if (serverRole === 'both' || serverRole === 'admin') serverRole = 'steward';
  return DataService.getBatchData(e, serverRole);
}

/**
 * Lightweight check + auto-init of missing sheets.
 * Called fire-and-forget from client AFTER view renders — never blocks initial load.
 * Version-keyed Script Property prevents re-running on every page load.
 */
function dataEnsureSheetsIfNeeded(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { skipped: true };
  var initKey = 'sheetsInitialized_' + (typeof VERSION !== 'undefined' ? VERSION : 'unknown');
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(initKey)) return { skipped: true, reason: 'already initialized' };
  try {
    _ensureAllSheetsInternal();
    props.setProperty(initKey, new Date().toISOString());
    return { initialized: true };
  } catch (err) {
    log_('Auto-init sheets warning', err.message);
    return { initialized: false, error: err.message };
  }
}

/**
 * Canonical sheet initialization — single source of truth.
 * Non-destructive: all init functions skip if sheet already exists.
 * Returns { created: string[], failed: string[] } for tracking.
 * @private
 */
function _ensureAllSheetsInternal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { created: [], failed: ['Spreadsheet binding broken'] };

  var created = [];
  var failed = [];

  var inits = [
    ['Hidden sheets',     function() { if (typeof setupHiddenSheets === 'function') setupHiddenSheets(ss); }],
    ['Contact Log',       function() { if (typeof _ensureContactLogSheet === 'function') _ensureContactLogSheet(ss); }],
    ['Steward Tasks',     function() { if (typeof _ensureStewardTasksSheet === 'function') _ensureStewardTasksSheet(ss); }],
    ['QA Forum',          function() { if (typeof QAForum !== 'undefined' && QAForum.initQAForumSheets) QAForum.initQAForumSheets(); }],
    ['Timeline',          function() { if (typeof TimelineService !== 'undefined' && TimelineService.initTimelineSheet) TimelineService.initTimelineSheet(); }],
    ['Failsafe Config',   function() { if (typeof FailsafeService !== 'undefined' && FailsafeService.initFailsafeSheet) FailsafeService.initFailsafeSheet(); }],
    ['Weekly Questions',  function() { if (typeof WeeklyQuestions !== 'undefined' && WeeklyQuestions.initWeeklyQuestionSheets) WeeklyQuestions.initWeeklyQuestionSheets(); }],
    ['Portal sheets',     function() { if (typeof initPortalSheets === 'function') initPortalSheets(); }],
    ['Workload Tracker',  function() { if (typeof initWorkloadTrackerSheets === 'function') initWorkloadTrackerSheets(); }],
    ['Resources',         function() { if (typeof createResourcesSheet === 'function') createResourcesSheet(ss); }],
    ['Resource Config',   function() { if (typeof createResourceConfigSheet === 'function') createResourceConfigSheet(ss); }],
    ['Survey Questions',  function() { if (typeof createSurveyQuestionsSheet === 'function') createSurveyQuestionsSheet(ss); }],
    ['Satisfaction',      function() { if (typeof createSatisfactionSheet === 'function') createSatisfactionSheet(ss); }],
    ['Feedback',          function() { if (typeof createFeedbackSheet === 'function') createFeedbackSheet(ss); }],
    ['Case Checklist',    function() { if (typeof getOrCreateChecklistSheet === 'function') getOrCreateChecklistSheet(); }],
    ['Notifications',     function() {
      if (!ss.getSheetByName(SHEETS.NOTIFICATIONS)) {
        var s = ss.insertSheet(SHEETS.NOTIFICATIONS);
        s.getRange(1, 1, 1, 10).setValues([['ID', 'Type', 'Subject', 'Body', 'Sender', 'Recipients', 'Created', 'Status', 'Priority', 'Metadata']]);
        s.hideSheet();
      }
    }],
    ['Grievance Feedback', function() { if (typeof _ensureGrievanceFeedbackSheet === 'function') _ensureGrievanceFeedbackSheet(ss); }],
    ['Audit Log',         function() {
      if (!ss.getSheetByName(SHEETS.AUDIT_LOG)) {
        var s = ss.insertSheet(SHEETS.AUDIT_LOG);
        s.getRange(1, 1, 1, 6).setValues([['Timestamp', 'User', 'Action', 'Target', 'Details', 'IP']]);
        s.hideSheet();
      }
    }],
  ];

  inits.forEach(function(pair) {
    try {
      pair[1]();
      created.push(pair[0]);
    } catch (err) {
      failed.push(pair[0] + ': ' + err.message);
      log_('_ensureAllSheetsInternal', 'Auto-init ' + pair[0] + ' failed: ' + err.message);
    }
  });

  return { created: created, failed: failed };
}

// Broadcast filter options (CR-AUTH-3: steward auth required)
/**
 * Returns available broadcast filter options (locations, office days, etc.). Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { locations, officeDays, hasDuesPayingColumn, broadcastScopeAll, totalMembers }
 */
function dataGetBroadcastFilterOptions(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { locations: [], officeDays: [], hasDuesPayingColumn: false, broadcastScopeAll: false, totalMembers: 0 };
  try {
    var members = DataService.getAllMembers();
    var locations = {};
    var officeDays = {};
    var hasDuesPayingColumn = false;
    members.forEach(function(m) {
      if (m.workLocation) locations[m.workLocation] = true;
      var days = String(m.officeDays || '');
      if (days) {
        days.split(/[,;]/).forEach(function(d) {
          var day = d.trim();
          if (day) officeDays[day] = true;
        });
      }
      // Detect if dues paying column exists (null = absent, true/false = present)
      if (m.duesPaying !== null && m.duesPaying !== undefined) hasDuesPayingColumn = true;
    });
    // Read broadcastScopeAll from config — determines if All Members scope toggle is shown
    var broadcastScopeAll = false;
    try {
      var config = ConfigReader.getConfig();
      broadcastScopeAll = (String(config.broadcastScopeAll || '').trim().toLowerCase() === 'yes');
    } catch (_e) { log_('_e', (_e.message || _e)); }
    return {
      locations: Object.keys(locations).sort(),
      officeDays: Object.keys(officeDays).sort(),
      hasDuesPayingColumn: hasDuesPayingColumn,
      broadcastScopeAll: broadcastScopeAll,
      totalMembers: members.length
    };
  } catch (e) {
    log_('dataGetBroadcastFilterOptions error', e.message + '\n' + (e.stack || ''));
    return { locations: [], officeDays: [], hasDuesPayingColumn: false, broadcastScopeAll: false, totalMembers: 0 };
  }
}

// ═══════════════════════════════════════
// Resource Click Tracking (v4.32.1)
// ═══════════════════════════════════════
//
// PURPOSE:
//   Tracks when members view/expand a resource card in the Resources tab.
//   This data feeds the "Resource Views" KPI in the engagement stats dashboard,
//   replacing the hardcoded 0 that was returned before v4.32.1.
//
// ARCHITECTURE DECISION — Why a separate sheet instead of _Audit_Log?
//   Click events are high-frequency (every resource card expand triggers one).
//   The _Audit_Log uses integrity hash chaining (logAuditEvent) which adds
//   overhead per row. A lightweight append-only sheet avoids that cost and
//   keeps click analytics separate from security-sensitive audit records.
//   Pattern matches _Contact_Log and _Survey_Tracking (dedicated tracking sheets).
//
// SHEET SCHEMA (_Resource_Click_Log):
//   Col A: Timestamp   (Date)   — server-side Date() at time of click
//   Col B: User Email  (String) — resolved from session token (not user-supplied)
//   Col C: Resource ID (String) — e.g. "RES-003" from 📚 Resources sheet
//   Col D: Resource Title (String) — human-readable title, for manual inspection
//
// DATA FLOW:
//   member_view.html (card expand) → google.script.run.dataLogResourceClick()
//   → _resolveCallerEmail() → appendRow() to _Resource_Click_Log
//   → dataGetEngagementStats() reads row count as resourceDownloads KPI
//
// FRONTEND DEDUPLICATION:
//   The member_view.html click handler sets card._tracked = true after the
//   first expand, preventing duplicate server calls for the same card in the
//   same page session. A page reload resets this (by design — repeat visits
//   on different days are meaningful engagement signals).
//
// FAILURE MODES:
//   - Session expired / invalid token → returns { success: false }, no row written
//   - Spreadsheet unavailable (web app context) → returns { success: false }
//   - Sheet creation fails (permissions) → caught, logged, returns { success: false }
//   - appendRow fails (quota/lock) → caught, logged, returns { success: false }
//   In ALL failure cases the UI is unaffected — the click handler uses
//   withFailureHandler(function() {}) so errors are silently swallowed.
//   The resource card still expands normally. Only the analytics row is lost.
//
// SCALING:
//   At ~50 members × ~8 resources × daily use, expect ~400 rows/day.
//   At 10,000 rows the sheet should be reviewed for trimming (same threshold
//   as _Audit_Log). No auto-trim is implemented yet — add if needed.

/**
 * Logs a resource view/click to the _Resource_Click_Log hidden sheet.
 * Called from member_view.html when a resource card is expanded for the first time.
 * Auto-creates the hidden sheet with headers if it doesn't exist yet.
 *
 * This is a top-level data* wrapper (not routed through DataService) because
 * it's a simple append-only write with no business logic — same pattern as
 * dataMarkWelcomeDismissed and dataApplyColorTheme.
 *
 * @param {string} sessionToken — session token for caller authentication
 * @param {string} resourceId — resource ID from 📚 Resources sheet (e.g. "RES-003")
 * @param {string} [resourceTitle] — optional human-readable title for log readability
 * @returns {{ success: boolean }} — always returns object (never throws to caller)
 */
function dataLogResourceClick(sessionToken, resourceId, resourceTitle) {
  // Auth: resolve email from session token. Reject unauthenticated requests.
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false };
  if (!resourceId) return { success: false };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false };  // null in web app context if container unbound

    // Resolve sheet name from SHEETS constant; fallback string for safety
    var sheetName = SHEETS.RESOURCE_CLICK_LOG || '_Resource_Click_Log';
    var sheet = ss.getSheetByName(sheetName);

    // Auto-create hidden sheet on first click (lazy initialization).
    // This avoids requiring a setup step — the sheet appears only when
    // the feature is actually used. Hidden via hideSheet() to keep the
    // tab bar clean (matches _Contact_Log, _Survey_Tracking pattern).
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.hideSheet();
      sheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'User Email', 'Resource ID', 'Resource Title']]);
      sheet.setFrozenRows(1);
    }

    // Append the click record. appendRow is atomic and handles concurrency.
    // No LockService needed — appendRow is naturally safe for concurrent appends.
    sheet.appendRow([new Date(), email, String(resourceId), String(resourceTitle || '')]);
    return { success: true };
  } catch (e) {
    log_('dataLogResourceClick error', e.message);
    return { success: false };
  }
}

// Engagement stats — reads seeded union stats from Script Properties
/**
 * v4.22.0 — LIVE engagement stats from real sheets.
 * Replaces the SEEDED_UNION_STATS property stub.
 *
 * Metrics:
 *   surveyParticipation  — % of active members with status 'Completed' in _Survey_Tracking
 *   weeklyQuestionVotes  — total rows in _Weekly_Responses
 *   eventAttendance      — unique members who checked in to any meeting (Meeting Check-In Log)
 *   grievanceFilingRate  — % of active members with at least one grievance row
 *   stewardContactRate   — % of active members who appear as a member email in _Contact_Log
 *   resourceDownloads    — total resource click count from CacheService (RC_TOTAL)
 *   membershipTrends     — monthly total/new member counts from Member Directory HIRE_DATE (last 6 mo)
 */
function dataGetEngagementStats(sessionToken) {
  var _caller = _resolveCallerEmail(sessionToken);
  if (!_caller) return null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;

    // ── Helper: get sheet data as 2-D array (skip header row) ──────────────
    function _rows(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return [];
      return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    }

    // ── Active members ──────────────────────────────────────────────────────
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    var totalMembers = 0;
    var memberEmails = [];
    if (memberSheet && memberSheet.getLastRow() >= 2) {
      var mData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
      var mHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
      var emailIdx   = mHeaders.indexOf('Email');
      var duesIdx    = mHeaders.indexOf('Dues Status');
      var hireIdx    = mHeaders.indexOf('Hire Date');
      for (var mi = 0; mi < mData.length; mi++) {
        var dues = duesIdx >= 0 ? String(mData[mi][duesIdx]).trim() : '';
        if (dues !== '' && dues.toLowerCase() !== 'inactive') {
          totalMembers++;
          if (emailIdx >= 0 && mData[mi][emailIdx]) {
            memberEmails.push(String(mData[mi][emailIdx]).toLowerCase().trim());
          }
        }
      }
    }
    if (totalMembers === 0) return null; // no member data yet

    // ── Survey participation ────────────────────────────────────────────────
    var surveyParticipation = 0;
    try {
      var stRows = _rows(SHEETS.SURVEY_TRACKING || '_Survey_Tracking');
      var completedCount = 0;
      for (var si = 0; si < stRows.length; si++) {
        // SURVEY_TRACKING_COLS.CURRENT_STATUS is col 6 (1-indexed) → array index 5
        var status = String(stRows[si][5]).trim().toLowerCase();
        if (status === 'completed') completedCount++;
      }
      surveyParticipation = stRows.length > 0 ? Math.round((completedCount / Math.max(totalMembers, 1)) * 100) : 0;
    } catch (_se) { log_('_se', (_se.message || _se)); }

    // ── Weekly question votes ───────────────────────────────────────────────
    var weeklyQuestionVotes = 0;
    try {
      var wqRows = _rows(SHEETS.WEEKLY_RESPONSES || '_Weekly_Responses');
      weeklyQuestionVotes = wqRows.length;
    } catch (_we) { log_('_we', (_we.message || _we)); }

    // ── Event attendance (unique members at any meeting) ────────────────────
    var eventAttendance = 0;
    try {
      var ciRows = _rows(SHEETS.MEETING_CHECKIN_LOG);
      var attendeeSet = {};
      // MEETING_CHECKIN_COLS.EMAIL is col 8 (1-indexed) → array index 7
      for (var ci = 0; ci < ciRows.length; ci++) {
        var ciEmail = String(ciRows[ci][7]).toLowerCase().trim();
        if (ciEmail) attendeeSet[ciEmail] = true;
      }
      eventAttendance = Object.keys(attendeeSet).length;
    } catch (_ce) { log_('_ce', (_ce.message || _ce)); }

    // ── Grievance filing rate ───────────────────────────────────────────────
    var grievanceFilingRate = 0;
    try {
      var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (gSheet && gSheet.getLastRow() >= 2) {
        var gHeaders = gSheet.getRange(1, 1, 1, gSheet.getLastColumn()).getValues()[0];
        var gEmailIdx = gHeaders.indexOf('Member Email');
        if (gEmailIdx >= 0) {
          var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, gSheet.getLastColumn()).getValues();
          var grievantSet = {};
          for (var gi = 0; gi < gData.length; gi++) {
            var ge = String(gData[gi][gEmailIdx]).toLowerCase().trim();
            if (ge) grievantSet[ge] = true;
          }
          var grievants = Object.keys(grievantSet).filter(function(e) { return memberEmails.indexOf(e) >= 0; });
          grievanceFilingRate = Math.round((grievants.length / totalMembers) * 100);
        }
      }
    } catch (_ge) { log_('_ge', (_ge.message || _ge)); }

    // ── Steward contact rate ────────────────────────────────────────────────
    var stewardContactRate = 0;
    try {
      var clSheet = ss.getSheetByName(SHEETS.CONTACT_LOG || '_Contact_Log');
      if (clSheet && clSheet.getLastRow() >= 2) {
        var clData = clSheet.getRange(2, 1, clSheet.getLastRow() - 1, clSheet.getLastColumn()).getValues();
        // _Contact_Log col layout: ID[0] StewardEmail[1] MemberEmail[2] Type[3] Date[4] Notes[5] Duration[6] Created[7]
        var contactedSet = {};
        for (var cli = 0; cli < clData.length; cli++) {
          var cme = String(clData[cli][2]).toLowerCase().trim();
          if (cme) contactedSet[cme] = true;
        }
        var contacted = Object.keys(contactedSet).filter(function(e) { return memberEmails.indexOf(e) >= 0; });
        stewardContactRate = Math.round((contacted.length / totalMembers) * 100);
      }
    } catch (_cl) { log_('_cl', (_cl.message || _cl)); }

    // ── Resource views (v4.32.1 — total clicks from _Resource_Click_Log) ────
    // ── Membership trends (last 6 months, by hire date) ─────────────────────
    var membershipTrends = [];
    try {
      if (memberSheet && hireIdx >= 0 && mData) {
        var now = new Date();
        var monthMap = {};
        for (var ti = 5; ti >= 0; ti--) {
          var d = new Date(now.getFullYear(), now.getMonth() - ti, 1);
          var key = (d.getMonth() + 1) + '/' + (d.getFullYear() - 2000);
          monthMap[key] = { month: key, total: 0, new: 0 };
        }
        var firstDate = new Date(now.getFullYear(), now.getMonth() - 5, 1);
        for (var ri = 0; ri < mData.length; ri++) {
          var hire = mData[ri][hireIdx];
          if (!hire) continue;
          var hd = hire instanceof Date ? hire : new Date(hire);
          if (isNaN(hd.getTime())) continue;
          // Count total active members in each month window (cumulative)
          // Simple approach: count members hired on or before end of each month
          for (var mk in monthMap) {
            var mEnd = new Date(parseInt('20' + mk.split('/')[1]), parseInt(mk.split('/')[0]), 0);
            if (hd <= mEnd) monthMap[mk].total++;
          }
          // New = hired within the 6-month window
          if (hd >= firstDate) {
            var hKey = (hd.getMonth() + 1) + '/' + (hd.getFullYear() - 2000);
            if (monthMap[hKey]) monthMap[hKey].new++;
          }
        }
        for (var mk2 in monthMap) membershipTrends.push(monthMap[mk2]);
      }
    } catch (_te) { log_('_te', (_te.message || _te)); }

    // ── Meeting attendance trends (last 6 months) ─────────────────────────
    var meetingTrends = [];
    var totalMeetings = 0;
    var avgMeetingAttendance = 0;
    try {
      var ciSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
      if (ciSheet && ciSheet.getLastRow() >= 2) {
        var ciData = ciSheet.getRange(2, 1, ciSheet.getLastRow() - 1, ciSheet.getLastColumn()).getValues();
        var ciHeaders = ciSheet.getRange(1, 1, 1, ciSheet.getLastColumn()).getValues()[0];
        var ciDateIdx = ciHeaders.indexOf('Meeting Date');
        var ciEmailIdx2 = ciHeaders.indexOf('Email');
        var ciMeetingIdIdx = ciHeaders.indexOf('Meeting ID');
        if (ciDateIdx >= 0) {
          var mtMap = {};
          var meetingIds = {};
          var sixMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 5, 1);
          for (var mi2 = 0; mi2 < ciData.length; mi2++) {
            var md = ciData[mi2][ciDateIdx];
            if (!md) continue;
            var mdDate = md instanceof Date ? md : new Date(md);
            if (isNaN(mdDate.getTime()) || mdDate < sixMonthsAgo) continue;
            var mtKey = (mdDate.getMonth() + 1) + '/' + (mdDate.getFullYear() - 2000);
            if (!mtMap[mtKey]) mtMap[mtKey] = {};
            var mtEmail = ciEmailIdx2 >= 0 ? String(ciData[mi2][ciEmailIdx2]).toLowerCase().trim() : '';
            if (mtEmail) mtMap[mtKey][mtEmail] = 1;
            if (ciMeetingIdIdx >= 0 && ciData[mi2][ciMeetingIdIdx]) meetingIds[String(ciData[mi2][ciMeetingIdIdx])] = 1;
          }
          totalMeetings = Object.keys(meetingIds).length;
          var mtAttTotal = 0;
          for (var mtk in mtMap) {
            var mtCount = Object.keys(mtMap[mtk]).length;
            meetingTrends.push({ month: mtk, attendees: mtCount });
            mtAttTotal += mtCount;
          }
          avgMeetingAttendance = meetingTrends.length > 0 ? Math.round(mtAttTotal / meetingTrends.length) : 0;
        }
      }
    } catch (_mt) { log_('_mt', (_mt.message || _mt)); }

    // ── Contact Log volume (last 6 months) ──────────────────────────────
    var contactTrends = [];
    var totalContacts = 0;
    try {
      var clSheet2 = ss.getSheetByName(SHEETS.CONTACT_LOG || '_Contact_Log');
      if (clSheet2 && clSheet2.getLastRow() >= 2) {
        var clData2 = clSheet2.getRange(2, 1, clSheet2.getLastRow() - 1, clSheet2.getLastColumn()).getValues();
        // _Contact_Log: ID[0] StewardEmail[1] MemberEmail[2] Type[3] Date[4] Notes[5] Duration[6] Created[7]
        var clMap = {};
        var sixMoAgo = new Date(now.getFullYear(), now.getMonth() - 5, 1);
        for (var ci2 = 0; ci2 < clData2.length; ci2++) {
          var clDate = clData2[ci2][4]; // Date column
          if (!clDate) continue;
          var cld = clDate instanceof Date ? clDate : new Date(clDate);
          if (isNaN(cld.getTime())) continue;
          totalContacts++;
          if (cld >= sixMoAgo) {
            var clKey = (cld.getMonth() + 1) + '/' + (cld.getFullYear() - 2000);
            clMap[clKey] = (clMap[clKey] || 0) + 1;
          }
        }
        for (var clk in clMap) contactTrends.push({ month: clk, count: clMap[clk] });
      }
    } catch (_cl2) { log_('_cl2', (_cl2.message || _cl2)); }

    // ── Satisfaction trends ──────────────────────────────────────────────
    var satisfactionData = null;
    try {
      satisfactionData = DataService.getSatisfactionTrends();
    } catch (_sat) { log_('_sat', (_sat.message || _sat)); }

    // ── Steward task completion rate ─────────────────────────────────────
    var taskCompletionRate = null;
    var totalTasks = 0;
    var completedTasks = 0;
    try {
      var taskSheet = ss.getSheetByName(SHEETS.STEWARD_TASKS || '_Steward_Tasks');
      if (taskSheet && taskSheet.getLastRow() >= 2) {
        var taskData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).getValues();
        // _Steward_Tasks: ID[0] StewardEmail[1] Title[2] Description[3] MemberEmail[4] Priority[5] Status[6] DueDate[7] Created[8] Completed[9]
        for (var tki = 0; tki < taskData.length; tki++) {
          var tkStatus = String(taskData[tki][6]).trim().toLowerCase();
          if (tkStatus) {
            totalTasks++;
            if (tkStatus === 'completed' || tkStatus === 'done') completedTasks++;
          }
        }
        taskCompletionRate = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : null;
      }
    } catch (_tk) { log_('_tk', (_tk.message || _tk)); }

    // ── Notification engagement ──────────────────────────────────────────
    var notifTotal = 0;
    var notifDismissed = 0;
    var notifActive = 0;
    try {
      var notifSheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
      if (notifSheet && notifSheet.getLastRow() >= 2) {
        var notifData = notifSheet.getRange(2, 1, notifSheet.getLastRow() - 1, notifSheet.getLastColumn()).getValues();
        var notifHeaders = notifSheet.getRange(1, 1, 1, notifSheet.getLastColumn()).getValues()[0];
        var notifStatusIdx = notifHeaders.indexOf('Status');
        var notifDismissedIdx = notifHeaders.indexOf('Dismissed_By');
        for (var ni = 0; ni < notifData.length; ni++) {
          notifTotal++;
          var nStatus = notifStatusIdx >= 0 ? String(notifData[ni][notifStatusIdx]).trim().toLowerCase() : '';
          var nDismissed = notifDismissedIdx >= 0 ? String(notifData[ni][notifDismissedIdx]).trim() : '';
          if (nStatus === 'active' || nStatus === 'sent') notifActive++;
          if (nDismissed) notifDismissed++;
        }
      }
    } catch (_ni) { log_('_ni', (_ni.message || _ni)); }

    // ── QA Forum activity ────────────────────────────────────────────────
    var qaQuestions = 0;
    var qaAnswered = 0;
    var qaUpvotes = 0;
    try {
      var qaSheet = ss.getSheetByName(SHEETS.QA_FORUM || '_QA_Forum');
      if (qaSheet && qaSheet.getLastRow() >= 2) {
        var qaData = qaSheet.getRange(2, 1, qaSheet.getLastRow() - 1, qaSheet.getLastColumn()).getValues();
        // _QA_Forum: ID[0] AuthorEmail[1] AuthorName[2] IsAnonymous[3] QuestionText[4] Status[5] UpvoteCount[6] Upvoters[7] AnswerCount[8] Created[9] Updated[10]
        for (var qi = 0; qi < qaData.length; qi++) {
          qaQuestions++;
          var ansCount = parseInt(qaData[qi][8], 10) || 0;
          if (ansCount > 0) qaAnswered++;
          qaUpvotes += parseInt(qaData[qi][6], 10) || 0;
        }
      }
    } catch (_qa) { log_('_qa', (_qa.message || _qa)); }

    // ── Steward coverage ratio & workload equity (v4.31.5) ───────────────
    var stewardCoverage = [];
    var workloadEquity = null;
    try {
      // Re-read member sheet for steward assignments
      if (memberSheet && memberSheet.getLastRow() >= 2) {
        var memHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
        var assignedStewardIdx = memHeaders.indexOf('Assigned Steward');
        var isStewardIdx = memHeaders.indexOf('Is Steward');
        var memLocIdx = memHeaders.indexOf('Work Location');
        if (assignedStewardIdx >= 0) {
          var memAllData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
          var stewardMemberCount = {}; // stewardEmail -> count
          var locationStewards = {}; // location -> { stewards: {}, members: 0 }
          for (var sci = 0; sci < memAllData.length; sci++) {
            var assignedTo = String(memAllData[sci][assignedStewardIdx]).trim().toLowerCase();
            var isSteward = isStewardIdx >= 0 ? String(memAllData[sci][isStewardIdx]).trim().toLowerCase() : '';
            var memLoc = memLocIdx >= 0 ? String(memAllData[sci][memLocIdx]).trim() : 'Unknown';
            if (!locationStewards[memLoc]) locationStewards[memLoc] = { stewards: {}, members: 0 };
            locationStewards[memLoc].members++;
            if (isSteward === 'yes' || isSteward === 'true') {
              var sEmail = emailIdx >= 0 ? String(memAllData[sci][emailIdx]).toLowerCase().trim() : '';
              if (sEmail) locationStewards[memLoc].stewards[sEmail] = 1;
            }
            if (assignedTo && assignedTo !== 'unassigned' && assignedTo !== '') {
              stewardMemberCount[assignedTo] = (stewardMemberCount[assignedTo] || 0) + 1;
            }
          }
          // Coverage by location
          for (var locName in locationStewards) {
            var ls = locationStewards[locName];
            var stewardCount = Object.keys(ls.stewards).length;
            if (ls.members > 0) {
              stewardCoverage.push({
                location: locName,
                members: ls.members,
                stewards: stewardCount,
                ratio: stewardCount > 0 ? Math.round((ls.members / stewardCount) * 10) / 10 : null
              });
            }
          }
          stewardCoverage.sort(function(a, b) { return (b.ratio || 999) - (a.ratio || 999); });

          // Workload equity (Gini coefficient of caseload distribution)
          var caseloads = [];
          for (var ste in stewardMemberCount) caseloads.push(stewardMemberCount[ste]);
          if (caseloads.length >= 2) {
            caseloads.sort(function(a, b) { return a - b; });
            var n = caseloads.length;
            var totalSum = caseloads.reduce(function(a, b) { return a + b; }, 0);
            var giniSum = 0;
            for (var gj = 0; gj < n; gj++) giniSum += (2 * (gj + 1) - n - 1) * caseloads[gj];
            var gini = totalSum > 0 ? Math.round((giniSum / (n * totalSum)) * 100) / 100 : 0;
            workloadEquity = {
              gini: gini,
              equityLabel: gini < 0.2 ? 'Very Even' : gini < 0.35 ? 'Fairly Even' : gini < 0.5 ? 'Moderate' : 'Uneven',
              stewardCount: n,
              avgCaseload: Math.round((totalSum / n) * 10) / 10,
              maxCaseload: caseloads[n - 1],
              minCaseload: caseloads[0]
            };
          }
        }
      }
    } catch (_sc) { log_('_sc', (_sc.message || _sc)); }

    return {
      surveyParticipation:  surveyParticipation,
      weeklyQuestionVotes:  weeklyQuestionVotes,
      eventAttendance:      eventAttendance,
      grievanceFilingRate:  grievanceFilingRate,
      stewardContactRate:   stewardContactRate,
      resourceDownloads:    DataService.getResourceClickTotal(),
      membershipTrends:     membershipTrends,
      // New metrics (v4.31.4)
      meetingTrends:        meetingTrends,
      totalMeetings:        totalMeetings,
      avgMeetingAttendance: avgMeetingAttendance,
      contactTrends:        contactTrends,
      totalContacts:        totalContacts,
      satisfactionData:     satisfactionData,
      taskCompletionRate:   taskCompletionRate,
      totalTasks:           totalTasks,
      completedTasks:       completedTasks,
      notifTotal:           notifTotal,
      notifDismissed:       notifDismissed,
      notifActive:          notifActive,
      qaQuestions:           qaQuestions,
      qaAnswered:            qaAnswered,
      qaUpvotes:             qaUpvotes,
      // Operational health (v4.31.5)
      stewardCoverage:       stewardCoverage,
      workloadEquity:        workloadEquity,
    };
  } catch (e) {
    log_('dataGetEngagementStats error', e.message + '\n' + (e.stack || ''));
    return null;
  }
}

/**
 * v4.31.5 — Per-member engagement score. Private — only shows the caller's own data.
 * Composite score (0-100) based on: survey participation, meeting attendance,
 * resource views, Q&A activity, and steward contact.
 */
function dataGetMyEngagementScore(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var email = e.toLowerCase().trim();

    var score = 0;
    var breakdown = [];

    // 1. Survey participation (0-25 points)
    try {
      var stSheet = ss.getSheetByName(SHEETS.SURVEY_TRACKING || '_Survey_Tracking');
      if (stSheet && stSheet.getLastRow() >= 2) {
        var stData = stSheet.getRange(2, 1, stSheet.getLastRow() - 1, stSheet.getLastColumn()).getValues();
        for (var si = 0; si < stData.length; si++) {
          if (String(stData[si][1]).toLowerCase().trim() === email) { // col 2 = email (0-indexed: 1)
            var status = String(stData[si][5]).toLowerCase().trim(); // col 6 = status
            if (status === 'completed') { score += 25; breakdown.push({ label: 'Survey Completed', points: 25, max: 25 }); }
            else if (status === 'in progress') { score += 10; breakdown.push({ label: 'Survey In Progress', points: 10, max: 25 }); }
            else { breakdown.push({ label: 'Survey Not Started', points: 0, max: 25 }); }
            break;
          }
        }
      }
      if (!breakdown.length) breakdown.push({ label: 'Survey', points: 0, max: 25 });
    } catch (_s) { breakdown.push({ label: 'Survey', points: 0, max: 25 }); }

    // 2. Meeting attendance (0-25 points): 5 pts per meeting attended (max 25)
    var meetingPts = 0;
    try {
      var ciSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
      if (ciSheet && ciSheet.getLastRow() >= 2) {
        var ciData = ciSheet.getRange(2, 1, ciSheet.getLastRow() - 1, ciSheet.getLastColumn()).getValues();
        var ciHeaders = ciSheet.getRange(1, 1, 1, ciSheet.getLastColumn()).getValues()[0];
        var ciEmailIdx = ciHeaders.indexOf('Email');
        if (ciEmailIdx >= 0) {
          var meetingsAttended = 0;
          for (var mi = 0; mi < ciData.length; mi++) {
            if (String(ciData[mi][ciEmailIdx]).toLowerCase().trim() === email) meetingsAttended++;
          }
          meetingPts = Math.min(meetingsAttended * 5, 25);
        }
      }
    } catch (_m) { /* skip */ }
    score += meetingPts;
    breakdown.push({ label: 'Meeting Attendance', points: meetingPts, max: 25 });

    // 3. Q&A activity (0-20 points): 5 pts per question asked, 3 pts per upvote given
    var qaPts = 0;
    try {
      var qaSheet = ss.getSheetByName(SHEETS.QA_FORUM || '_QA_Forum');
      if (qaSheet && qaSheet.getLastRow() >= 2) {
        var qaData = qaSheet.getRange(2, 1, qaSheet.getLastRow() - 1, qaSheet.getLastColumn()).getValues();
        for (var qi = 0; qi < qaData.length; qi++) {
          if (String(qaData[qi][1]).toLowerCase().trim() === email) qaPts += 5;
          // Check upvoters for this user
          var upvoters = String(qaData[qi][7]);
          if (upvoters.toLowerCase().indexOf(email) >= 0) qaPts += 3;
        }
        qaPts = Math.min(qaPts, 20);
      }
    } catch (_q) { /* skip */ }
    score += qaPts;
    breakdown.push({ label: 'Q&A Participation', points: qaPts, max: 20 });

    // 4. Resource engagement (0-15 points): based on total views by this user
    // We can't track per-user views with current RC_ scheme, so give 15 if they've used resources tab
    var resourcePts = 0;
    try {
      var props = PropertiesService.getScriptProperties();
      var allProps = props.getProperties();
      // Check usage analytics for resource tab visits by this user
      for (var uaKey in allProps) {
        if (uaKey.indexOf('UA_D_') === 0) {
          try {
            var parsed = JSON.parse(allProps[uaKey]);
            if (parsed.users && parsed.users[email]) { resourcePts = 10; break; }
          } catch (_p) { /* skip */ }
        }
      }
      // Bonus if they've viewed resources tab specifically
      for (var tKey in allProps) {
        if (tKey.indexOf('UA_T_') >= 0 && tKey.indexOf('_resources') >= 0) {
          resourcePts = 15;
          break;
        }
      }
    } catch (_r) { /* skip */ }
    score += resourcePts;
    breakdown.push({ label: 'Resource Engagement', points: resourcePts, max: 15 });

    // 5. Steward contact (0-15 points): contacted steward at least once
    var contactPts = 0;
    try {
      var clSheet = ss.getSheetByName(SHEETS.CONTACT_LOG || '_Contact_Log');
      if (clSheet && clSheet.getLastRow() >= 2) {
        var clData = clSheet.getRange(2, 1, clSheet.getLastRow() - 1, clSheet.getLastColumn()).getValues();
        for (var ci = 0; ci < clData.length; ci++) {
          if (String(clData[ci][2]).toLowerCase().trim() === email) { contactPts = 15; break; }
        }
      }
    } catch (_c) { /* skip */ }
    score += contactPts;
    breakdown.push({ label: 'Steward Contact', points: contactPts, max: 15 });

    // Percentile (approximate — based on active member count)
    var percentile = null;
    try {
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberSheet && memberSheet.getLastRow() >= 2) {
        // Very rough: assume normal distribution centered at 40 with std of 20
        var memberCount = memberSheet.getLastRow() - 1;
        if (memberCount > 0) {
          // z-score approximation
          var z = (score - 40) / 20;
          percentile = Math.min(99, Math.max(1, Math.round(50 + 50 * (z / Math.sqrt(1 + z * z)))));
        }
      }
    } catch (_pe) { /* skip */ }

    return {
      score: Math.min(score, 100),
      breakdown: breakdown,
      percentile: percentile,
      maxScore: 100
    };
  } catch (err) {
    log_('dataGetMyEngagementScore error', err.message);
    return null;
  }
}

/**
 * v4.25.11 — Resource usage stats for Union Stats > Resources sub-tab.
 * Returns per-resource views, top resources, category breakdown.
 */
function dataGetResourceStats(sessionToken) {
  var _caller = _resolveCallerEmail(sessionToken);
  if (!_caller) return null;
  try {
    return DataService.getResourceStats();
  } catch (e) {
    log_('dataGetResourceStats error', e.message);
    return null;
  }
}

/**
 * v4.31.3 — Log a tab visit for webapp usage analytics.
 * Fire-and-forget; client does not wait for response.
 */
function dataLogTabVisit(sessionToken, tab, role) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false };
  try {
    return DataService.logTabVisit(e, tab, role);
  } catch (err) {
    log_('dataLogTabVisit error', err.message);
    return { success: false };
  }
}

/**
 * v4.31.3 — Aggregated webapp usage stats for Union Stats > Usage sub-tab.
 */
function dataGetUsageStats(sessionToken) {
  var _caller = _resolveCallerEmail(sessionToken);
  if (!_caller) return null;
  try {
    return DataService.getUsageStats();
  } catch (e) {
    log_('dataGetUsageStats error', e.message);
    return null;
  }
}

/**
 * v4.22.0 — LIVE workload summary from Workload Vault.
 * Replaces the SEEDED_UNION_STATS property stub.
 *
 * Returns:
 *   avgCaseload       — average of PRIORITY_CASES across most-recent submission per steward
 *   highCaseloadPct   — % of stewards with priority_cases > 5
 *   submissionRate    — % of stewards (IS_STEWARD = 'Yes') who have submitted at least once
 *   trendDirection    — 'increasing' | 'decreasing' | 'stable' based on avg last 4 wks vs prior 4 wks
 */
/**
 * Returns lightweight member count and active-grievance count. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { total: number, withGrievances: number }
 */
function dataGetMemberCount(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try { return DataService.getMemberCount(); } catch (e) { log_('dataGetMemberCount error', e.message); return { total: 0, withGrievances: 0 }; }
}

/** Returns distinct filter dropdown values for the Members finder panel. Requires steward auth. */
function dataGetFilterDropdownValues(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { locations: [], units: [], stewards: [] };
  try { return DataService.getFilterDropdownValues(); } catch (e) { log_('dataGetFilterDropdownValues error', e.message); return { locations: [], units: [], stewards: [] }; }
}

/**
 * Returns a paginated page of members with search/filter support. Requires steward auth.
 * @param {string} sessionToken
 * @param {Object} [opts] - { page, pageSize, search, filter }
 * @returns {Object} { items, total, page, pageSize, totalPages }
 */
function dataGetMembersPaginated(sessionToken, opts) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try { return DataService.getMembersPaginated(s, opts || {}); } catch (e) { log_('dataGetMembersPaginated error', e.message); return { members: [], total: 0, page: 1, pageSize: 25 }; }
}

/**
 * Returns row counts and scale status for key sheets. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { members: { rows, status }, grievances: { rows, status } }
 */
function dataGetSheetHealth(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try { return DataService.getSheetHealth(); } catch (e) { log_('dataGetSheetHealth error', e.message); return { members: { rows: 0, status: 'ok' }, grievances: { rows: 0, status: 'ok' } }; }
}

/**
 * Returns live workload summary stats. Stub — WorkloadService excluded from this edition.
 * @param {string} sessionToken
 * @returns {Object} { success: true, data: { totalEntries: 0, categories: {} } }
 */
function dataGetWorkloadSummaryStats(sessionToken) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };
  return { success: true, data: { totalEntries: 0, categories: {} } };
}

// ═══════════════════════════════════════
// WELCOME EXPERIENCE (PHASE2)
// Uses PropertiesService to track first-visit state per user.
// ═══════════════════════════════════════

/**
 * Returns welcome/onboarding data for the current user.
 * Checks if this is the user's first visit by looking up a property key.
 * @param {string} email - User email
 * @returns {Object} { isFirstVisit, userName, role, quickActions }
 */
function dataGetWelcomeData(sessionToken) {
  // CR-AUTH-3: Use server-side identity instead of client-supplied email
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { isFirstVisit: false, userName: '', role: 'member', quickActions: [] };

  var userRecord = DataService.findUserByEmail(email);
  var firstName = '';
  var role = 'member';
  if (userRecord) {
    firstName = userRecord.firstName || (userRecord.name || '').split(' ')[0] || '';
    role = userRecord.role || 'member';
  }

  // Check first-visit flag using script properties (user properties not available
  // in web app context running as "me"). Use a hash of email as key.
  var emailHash = _welcomeEmailHash(email);
  var propKey = 'WELCOME_DISMISSED_' + emailHash;
  var props = PropertiesService.getScriptProperties();
  var dismissed = props.getProperty(propKey);
  var isFirstVisit = !dismissed;

  // Build role-appropriate quick actions
  var quickActions = [];
  if (role === 'steward' || role === 'both') {
    quickActions = [
      { label: 'View Cases', icon: '\uD83D\uDCCB', action: 'cases' },
      { label: 'Check Deadlines', icon: '\u23F0', action: 'cases' },
      { label: 'Member Directory', icon: '\uD83D\uDC65', action: 'members' },
      { label: 'Manage Tasks', icon: '\u2705', action: 'tasks' },
    ];
  } else {
    quickActions = [
      { label: 'View My Cases', icon: '\uD83D\uDCCB', action: 'cases' },
      { label: 'Update Contact Info', icon: '\uD83D\uDC64', action: 'profile' },
      { label: 'Check Resources', icon: '\uD83D\uDCDA', action: 'resources' },
      { label: 'Contact ' + (ConfigReader.getConfig().stewardLabel || 'Steward'), icon: '\uD83D\uDCAC', action: 'stewarddirectory' },
    ];
  }

  return {
    isFirstVisit: isFirstVisit,
    userName: firstName,
    role: role,
    quickActions: quickActions,
  };
}

/**
 * Marks the welcome experience as dismissed for a user.
 * @param {string} email - User email
 * @returns {Object} { success: boolean }
 */
function dataMarkWelcomeDismissed(sessionToken) {
  // CR-AUTH-3: Use server-side identity instead of client-supplied email
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false };
  var emailHash = _welcomeEmailHash(email);
  var propKey = 'WELCOME_DISMISSED_' + emailHash;
  var props = PropertiesService.getScriptProperties();
  props.setProperty(propKey, new Date().toISOString());
  return { success: true };
}

/**
 * Applies a color theme preset (updates sheets and webapp accent). Requires auth.
 * @param {string} sessionToken
 * @param {string} themeKey - Theme preset key from THEME_PRESETS
 * @returns {Object} { success: boolean, themeKey: string, accentHue: number }
 */
function dataApplyColorTheme(sessionToken, themeKey) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  if (!THEME_PRESETS || !THEME_PRESETS[themeKey]) return { success: false, message: 'Unknown theme.' };
  var preset = THEME_PRESETS[themeKey];
  saveVisualSetting('theme', themeKey);
  if (preset.accentHue !== undefined) {
    saveVisualSetting('accentHue', preset.accentHue);
  }
  return { success: true, themeKey: themeKey, accentHue: preset.accentHue || 250 };
}

/**
 * Client-callable: Saves the user's default view preference for dual-role users.
 * Stored in ScriptProperties (email-keyed) because UserProperties returns the
 * script owner's props in Execute-as-Me webapps.
 * @param {string} sessionToken - Session token for auth
 * @param {string} viewPref - 'steward' or 'member'
 * @returns {{ success: boolean, defaultView?: string, message?: string }}
 */
function dataSetDefaultView(sessionToken, viewPref) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, message: 'Not authenticated.' };
  if (viewPref !== 'steward' && viewPref !== 'member') {
    return { success: false, message: 'Invalid view preference.' };
  }
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('defaultView_' + email.toLowerCase(), viewPref);
    return { success: true, defaultView: viewPref };
  } catch (err) {
    log_('dataSetDefaultView error', err.message);
    return { success: false, message: 'Failed to save preference.' };
  }
}

/**
 * Computes a simple hash of an email for use as a ScriptProperties key.
 * @param {string} email
 * @returns {string} Base-36 hash string
 * @private
 */
function _welcomeEmailHash(email) {
  var hash = 0;
  var str = String(email).toLowerCase();
  for (var i = 0; i < str.length; i++) {
    var c = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + c;
    hash = hash & hash; // Convert to 32-bit integer
  }
  return Math.abs(hash).toString(36);
}

// ═══════════════════════════════════════
// SEARCH WRAPPER (v4.33.0)
// ═══════════════════════════════════════

/**
 * Webapp endpoint: cross-tab search for members and grievances.
 * @param {string} sessionToken - Session token
 * @param {string} query - Search query (min 2 chars)
 * @param {string} tab - Filter: 'all', 'members', or 'grievances'
 * @returns {Object} { success, results: Array<{type,title,subtitle,detail}> }
 */
function dataGetWebAppSearchResults(sessionToken, query, tab) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, authError: true, message: 'Not authenticated.' };
  try {
    var results = getWebAppSearchResults(query || '', tab || 'all');
    return { success: true, results: results || [] };
  } catch (e) {
    log_('dataGetWebAppSearchResults error', e.message);
    return { success: false, message: 'Search failed: ' + e.message };
  }
}

// ═══════════════════════════════════════
// UNDO SYSTEM WRAPPERS (v4.33.0)
// ═══════════════════════════════════════

/**
 * Reverts undo history to the specified index. Requires steward auth.
 * @param {string} sessionToken
 * @param {number} targetIndex
 * @returns {Object} { success: boolean, message: string }
 */
function dataUndoToIndex(sessionToken, targetIndex) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    undoToIndex(Math.floor(targetIndex));
    return { success: true, message: 'Undone to index ' + targetIndex };
  } catch (e) {
    log_('dataUndoToIndex error', e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Exports undo history to a new sheet. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { success: boolean, sheetUrl?: string, message?: string }
 */
function dataExportUndoHistory(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    var url = exportUndoHistoryToSheet();
    return { success: true, sheetUrl: url };
  } catch (e) {
    log_('dataExportUndoHistory error', e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Returns the undo history stack. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { success: boolean, history?: Array, message?: string }
 */
function dataGetUndoHistory(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    var history = getUndoHistory();
    return { success: true, history: history };
  } catch (e) {
    log_('dataGetUndoHistory error', e.message);
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════
// MEETING CHECK-IN WRAPPER (v4.33.0)
// ═══════════════════════════════════════

/**
 * Checks in a member to a meeting via the webapp. Requires auth.
 * @param {string} sessionToken
 * @param {string} meetingId
 * @param {string} pin
 * @returns {Object} { success: boolean, message: string }
 */
function dataWebCheckInMember(sessionToken, meetingId, pin) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, authError: true, message: 'Not authenticated.' };
  try {
    return webCheckInMember(meetingId, email, pin);
  } catch (e) {
    log_('dataWebCheckInMember error', e.message);
    return { success: false, message: 'Check-in failed: ' + e.message };
  }
}

// ═══════════════════════════════════════
// CASE ACTIVITY LOG (v4.34.0 — Feature 5)
// ═══════════════════════════════════════

/**
 * Get case activity log from audit log. Steward-only.
 * Returns all audit log entries whose Record ID matches caseId,
 * sorted newest-first.
 * @param {string} sessionToken
 * @param {string} caseId
 * @returns {Array<Object>} Array of activity events
 */
function dataGetCaseActivityLog(sessionToken, caseId) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Audit log columns (0-indexed): Timestamp(0), User Email(1), Sheet(2), Row(3),
  // Column(4), Field Name(5), Old Value(6), New Value(7), Record ID(8), Action Type(9)
  var targetId = String(caseId).trim();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var recordId = String(data[i][AUDIT_LOG_COLS.RECORD_ID - 1] || '').trim();
    if (recordId === targetId) {
      results.push({
        timestamp: data[i][AUDIT_LOG_COLS.TIMESTAMP - 1],
        userEmail: String(data[i][AUDIT_LOG_COLS.USER_EMAIL - 1] || ''),
        field: String(data[i][AUDIT_LOG_COLS.FIELD_NAME - 1] || ''),
        oldValue: String(data[i][AUDIT_LOG_COLS.OLD_VALUE - 1] || ''),
        newValue: String(data[i][AUDIT_LOG_COLS.NEW_VALUE - 1] || ''),
        actionType: String(data[i][AUDIT_LOG_COLS.ACTION_TYPE - 1] || 'edit')
      });
    }
  }
  results.sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
  return results;
}

// ═══════════════════════════════════════
// DEADLINE CALENDAR WRAPPER (v4.33.0)
// ═══════════════════════════════════════

/**
 * Returns grievance deadline calendar data for the steward view. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { success: boolean, data?: Object, message?: string }
 */
function dataGetDeadlineCalendarData(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    return { success: true, data: getDeadlineCalendarData() };
  } catch (e) {
    log_('dataGetDeadlineCalendarData error', e.message);
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════
// CORRELATION ENGINE WRAPPERS (v4.33.0)
// ═══════════════════════════════════════

/**
 * Returns active correlation alerts from the correlation engine. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { success: boolean, alerts: Array }
 */
function dataGetCorrelationAlerts(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    return { success: true, alerts: JSON.parse(getCorrelationAlerts(false)) };
  } catch (e) {
    log_('dataGetCorrelationAlerts error', e.message);
    return { success: true, alerts: [] };
  }
}

/**
 * Returns a summary of all computed correlations. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { success: boolean, summary: Object }
 */
function dataGetCorrelationSummary(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try {
    return { success: true, summary: JSON.parse(getCorrelationSummary(false)) };
  } catch (e) {
    log_('dataGetCorrelationSummary error', e.message);
    return { success: true, summary: { total: 0, strong: 0, moderate: 0, weak: 0, negligible: 0, insufficientData: 0, topInsights: [], actionableCount: 0, disabled: true } };
  }
}

// ═══════════════════════════════════════
// ACCESS LOG VIEWER (v4.36.0)
// ═══════════════════════════════════════

/**
 * Returns paginated, filtered audit log entries for the Access Log Viewer tab.
 * PII is masked for non-admin callers.
 * @param {string} sessionToken
 * @param {Object} [options] - { page, pageSize, dateFrom, dateTo, userFilter, eventTypeFilter, searchTerm }
 * @returns {Object} { items, totalRows, page, pageSize, eventTypes }
 */
function dataGetAuditLog(sessionToken, options) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { items: [], totalRows: 0, page: 0, pageSize: 20, eventTypes: [] };
  options = options || {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { items: [], totalRows: 0, page: 0, pageSize: 20, eventTypes: [] };

  // Log this access
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('ACCESS_LOG_VIEWED', { viewer: typeof maskEmail === 'function' ? maskEmail(s) : s });
  }

  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG) || ss.getSheetByName('_Audit_Log');
  if (!sheet || sheet.getLastRow() < 2) return { items: [], totalRows: 0, page: 0, pageSize: 20, eventTypes: [] };

  var page = Math.max(0, parseInt(options.page, 10) || 0);
  var pageSize = Math.min(100, Math.max(10, parseInt(options.pageSize, 10) || 20));
  var data = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 5000), sheet.getLastColumn()).getValues();

  // Determine if caller is admin (for PII visibility)
  var isAdmin = false;
  try {
    if (typeof _adminIsAuthorized_ === 'function') isAdmin = _adminIsAuthorized_(s);
    else isAdmin = (s === Session.getEffectiveUser().getEmail());
  } catch (_) { /* ignore */ }

  // Collect distinct event types and filter
  var eventTypeSet = {};
  var filtered = [];
  var dateFrom = options.dateFrom ? new Date(options.dateFrom) : null;
  var dateTo = options.dateTo ? new Date(options.dateTo) : null;
  if (dateTo) dateTo.setHours(23, 59, 59, 999);
  var userFilter = options.userFilter ? String(options.userFilter).toLowerCase().trim() : '';
  var eventTypeFilter = options.eventTypeFilter ? String(options.eventTypeFilter).trim() : '';
  var searchTerm = options.searchTerm ? String(options.searchTerm).toLowerCase().trim() : '';

  // Process rows in reverse (newest first)
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var ts = row[0]; // Timestamp
    var eventType = String(row[9] || row[1] || '').trim(); // ACTION_TYPE or EVENT_TYPE
    var user = String(row[1] || '').trim(); // USER_EMAIL
    var details = String(row[7] || row[3] || '').trim(); // NEW_VALUE or DETAILS

    if (eventType) eventTypeSet[eventType] = true;

    // Apply filters
    if (dateFrom && ts instanceof Date && ts < dateFrom) continue;
    if (dateTo && ts instanceof Date && ts > dateTo) continue;
    if (userFilter && user.toLowerCase().indexOf(userFilter) === -1) continue;
    if (eventTypeFilter && eventType !== eventTypeFilter) continue;
    if (searchTerm && (user + ' ' + eventType + ' ' + details).toLowerCase().indexOf(searchTerm) === -1) continue;

    // Mask PII for non-admins
    var displayUser = user;
    var displayDetails = details;
    if (!isAdmin) {
      if (typeof maskEmail === 'function' && user.indexOf('@') !== -1) displayUser = maskEmail(user);
      if (typeof maskObjectPII_ === 'function') {
        try {
          var parsed = JSON.parse(details);
          displayDetails = JSON.stringify(maskObjectPII_(parsed));
        } catch (_) { /* not JSON, show as-is */ }
      }
    }

    filtered.push({
      timestamp: ts instanceof Date ? ts.toISOString() : String(ts),
      eventType: eventType,
      user: displayUser,
      sheet: String(row[2] || '').trim(),
      fieldName: String(row[5] || '').trim(),
      oldValue: String(row[6] || '').substring(0, 200),
      newValue: displayDetails.substring(0, 500),
      recordId: String(row[8] || '').trim(),
      actionType: eventType
    });
  }

  var totalRows = filtered.length;
  var start = page * pageSize;
  var items = filtered.slice(start, start + pageSize);

  return {
    items: items,
    totalRows: totalRows,
    page: page,
    pageSize: pageSize,
    eventTypes: Object.keys(eventTypeSet).sort()
  };
}

// ═══════════════════════════════════════
// BULK ACTIONS (v4.32.0 — Feature 3)
// ═══════════════════════════════════════

/** Bulk update status for multiple grievances. Steward-only. */
function dataBulkUpdateStatus(sessionToken, caseIds, newStatus) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!Array.isArray(caseIds) || !caseIds.length || !newStatus) return { success: false, message: 'Invalid parameters.' };

  // v4.36.0 — 2FA required for bulk operations
  if (typeof TwoFactorService !== 'undefined' && !TwoFactorService.hasValidSession(s)) {
    return { success: false, requires2FA: true, message: 'Verification required for bulk operations.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: 'No spreadsheet.' };
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { success: false, message: 'Sheet not found.' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = -1, statusCol = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).toLowerCase().trim();
    if (h === 'id' || h === 'grievance id') idCol = c;
    if (h === 'status') statusCol = c;
  }
  if (idCol === -1 || statusCol === -1) return { success: false, message: 'Column not found.' };

  var updated = 0;
  var safeStatus = typeof escapeForFormula === 'function' ? escapeForFormula(newStatus) : newStatus;
  for (var i = 1; i < data.length; i++) {
    if (caseIds.indexOf(String(data[i][idCol]).trim()) !== -1) {
      sheet.getRange(i + 1, statusCol + 1).setValue(safeStatus);
      updated++;
    }
  }
  if (typeof _refreshNavBadges === 'function') _refreshNavBadges();
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('BULK_STATUS_UPDATE', { count: updated, newStatus: newStatus, steward: s });
  }
  return { success: true, updated: updated };
}

/** Bulk export cases as CSV string. Steward-only. */
function dataBulkExportCsv(sessionToken, caseIds) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, csv: '' };
  if (!Array.isArray(caseIds) || !caseIds.length) return { success: false, csv: '' };

  // v4.36.0 — 2FA required for PII export
  if (typeof TwoFactorService !== 'undefined' && !TwoFactorService.hasValidSession(s)) {
    return { success: false, requires2FA: true, message: 'Verification required for data export.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, csv: '' };
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { success: false, csv: '' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).toLowerCase().trim() === 'id' || String(headers[c]).toLowerCase().trim() === 'grievance id') { idCol = c; break; }
  }
  if (idCol === -1) return { success: false, csv: '' };

  // Build CSV: headers + matching rows
  var csv = headers.map(function(h) { return '"' + String(h).replace(/"/g, '""') + '"'; }).join(',') + '\n';
  for (var i = 1; i < data.length; i++) {
    if (caseIds.indexOf(String(data[i][idCol]).trim()) !== -1) {
      csv += data[i].map(function(cell) { return '"' + String(cell).replace(/"/g, '""') + '"'; }).join(',') + '\n';
    }
  }
  return { success: true, csv: csv };
}

/** Bulk create Drive folders for cases. Steward-only. */
function dataBulkCreateFolders(sessionToken, caseIds) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!Array.isArray(caseIds) || !caseIds.length) return { success: false, message: 'No cases selected.' };

  var created = 0;
  var errors = [];
  for (var i = 0; i < caseIds.length; i++) {
    try {
      if (typeof setupDriveFolderForGrievance === 'function') {
        setupDriveFolderForGrievance(caseIds[i]);
        created++;
      }
    } catch (e) {
      errors.push(caseIds[i] + ': ' + e.message);
    }
  }
  return { success: true, created: created, errors: errors };
}

/** Bulk send email to members. Steward-only. */
function dataBulkSendEmail(sessionToken, memberEmails, subject, body) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!Array.isArray(memberEmails) || !memberEmails.length || !subject) return { success: false, message: 'Invalid parameters.' };

  var sent = 0;
  var safeSubject = typeof escapeHtml === 'function' ? escapeHtml(subject) : subject;
  var safeBody = typeof escapeHtml === 'function' ? escapeHtml(body || '') : (body || '');
  var htmlBody = '<p>' + safeBody.replace(/\n/g, '<br>') + '</p>';

  for (var i = 0; i < memberEmails.length; i++) {
    try {
      if (typeof safeSendEmail_ === 'function') {
        safeSendEmail_({ to: memberEmails[i], subject: safeSubject, htmlBody: htmlBody });
      } else if (MailApp.getRemainingDailyQuota() > 5) {
        MailApp.sendEmail({ to: memberEmails[i], subject: safeSubject, htmlBody: htmlBody });
      }
      sent++;
    } catch (e) {
      log_('dataBulkSendEmail', 'Bulk email failed for ' + memberEmails[i] + ': ' + e.message);
    }
  }
  return { success: true, sent: sent };
}


// ═══════════════════════════════════════
// POMS REFERENCE DATA (stub — excluded from this edition)
// ═══════════════════════════════════════

/**
 * Returns the POMS reference database. Stub — POMS excluded from this edition.
 * @param {string} sessionToken
 * @returns {Object} { success: true, data: [], message: ... }
 */
function dataGetPomsReference(sessionToken) {
  return { success: true, data: [], message: 'POMS Reference is not available in this edition.' };
}

// ═══════════════════════════════════════
// USAGE TRACKING (v4.40.0 — production-enabled)
// ═══════════════════════════════════════

/**
 * Logs usage events from the client-side UsageTracker.
 * v4.40.0: Enabled for all users (was dev-only). Tracks session time, tabs,
 * performance metrics, errors, and navigation patterns.
 * Writes to a hidden _Usage_Log sheet. Auto-trims to 10,000 rows max.
 *
 * @param {string} sessionToken - Caller's session token
 * @param {Object} payload - { sessionId, sessionStart, elapsed, userAgent, events: [...] }
 * @returns {Object} { success: boolean }
 */
function dataLogUsageEvents(sessionToken, payload) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, message: 'Not authenticated.' };

  if (!payload || !payload.events || !Array.isArray(payload.events)) {
    return { success: false, message: 'No events.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };

    var sheetName = (typeof HIDDEN_SHEETS !== 'undefined' && HIDDEN_SHEETS.USAGE_LOG) || '_Usage_Log';
    var sheet = ss.getSheetByName(sheetName);

    // Auto-create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['Timestamp', 'Email', 'Session ID', 'Event Type', 'Event Data', 'Session Elapsed (ms)', 'User Agent']);
      sheet.setFrozenRows(1);
      try { sheet.hideSheet(); } catch (_h) {}
    }

    // Batch-write events
    var rows = [];
    var sid = String(payload.sessionId || '').slice(0, 20);
    var ua = String(payload.userAgent || '').slice(0, 150);
    var elapsed = payload.elapsed || 0;

    for (var i = 0; i < payload.events.length && i < 100; i++) {
      var ev = payload.events[i];
      var ts = ev.t ? new Date(ev.t) : new Date();
      rows.push([
        ts,
        email,
        sid,
        String(ev.type || '').slice(0, 30),
        JSON.stringify(ev.data || {}).slice(0, 500),
        elapsed,
        ua
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    }

    // Auto-trim: keep only the latest 10,000 rows to prevent sheet bloat
    var rowCount = sheet.getLastRow();
    if (rowCount > 12000) {
      var deleteCount = rowCount - 10000;
      sheet.deleteRows(2, deleteCount); // keep header row
    }

    return { success: true, logged: rows.length };
  } catch (err) {
    log_('dataLogUsageEvents error', err.message);
    return { success: false, message: 'Logging failed.' };
  }
}

/**
 * Returns aggregated usage analytics for admin dashboard.
 * ADMIN ONLY — gated by _adminIsAuthorized_. Returns summary stats
 * from the _Usage_Log sheet for the requested time range.
 *
 * @param {string} sessionToken - Caller's session token
 * @param {number} [days] - Number of days to analyze (default 7)
 * @returns {Object} Aggregated analytics data
 */
function dataGetUsageAnalytics(sessionToken, days) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, message: 'Not authenticated.' };

  // Admin-only gate
  try {
    var isAdmin = false;
    if (typeof _adminIsAuthorized_ === 'function') {
      isAdmin = _adminIsAuthorized_(email);
    } else {
      isAdmin = (email.toLowerCase() === Session.getEffectiveUser().getEmail().toLowerCase());
    }
    if (!isAdmin) return { success: false, message: 'Admin access required.' };
  } catch (_authErr) {
    return { success: false, message: 'Authorization check failed.' };
  }

  days = Math.min(Math.max(parseInt(days, 10) || 7, 1), 90);
  var cutoff = new Date(Date.now() - days * 86400000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheetName = (typeof HIDDEN_SHEETS !== 'undefined' && HIDDEN_SHEETS.USAGE_LOG) || '_Usage_Log';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, days: days, totalEvents: 0, message: 'No usage data yet.' };
    }

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

    // Aggregate metrics
    var uniqueUsers = {};
    var sessions = {};
    var tabCounts = {};
    var authMethods = { sso: 0, magic: 0, pin: 0, session: 0, unknown: 0 };
    var perfLoads = [];     // page load times in ms
    var perfBatches = [];   // batch fetch times in ms
    var swrHits = 0;
    var swrMisses = 0;
    var backSwipes = 0;
    var errors = [];
    var deviceTypes = { mobile: 0, tablet: 0, desktop: 0 };
    var dailySessions = {};
    var userLoadTimes = {};
    var totalEvents = 0;
    var filteredEvents = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var ts = row[0];
      totalEvents++;

      // Filter by date range
      if (ts instanceof Date && ts < cutoff) continue;
      filteredEvents++;

      var userEmail = String(row[1] || '');
      var sid = String(row[2] || '');
      var eventType = String(row[3] || '');
      var eventData = {};
      try { eventData = JSON.parse(row[4] || '{}'); } catch (_) {}
      var ua = String(row[6] || '');

      // Track unique users
      if (userEmail) uniqueUsers[userEmail] = (uniqueUsers[userEmail] || 0) + 1;

      // Track sessions
      if (sid) sessions[sid] = true;

      // Daily sessions (for chart)
      if (ts instanceof Date) {
        var dayKey = ts.toISOString().slice(0, 10);
        if (!dailySessions[dayKey]) dailySessions[dayKey] = {};
        if (sid) dailySessions[dayKey][sid] = true;
      }

      // Parse events by type
      switch (eventType) {
        case 'tab':
          if (eventData.tab) tabCounts[eventData.tab] = (tabCounts[eventData.tab] || 0) + 1;
          break;
        case 'session_start':
          if (eventData.auth) {
            var method = eventData.auth;
            if (authMethods.hasOwnProperty(method)) authMethods[method]++;
            else authMethods.unknown++;
          }
          // Device detection from UA
          if (/Mobile|Android|iPhone/i.test(ua)) deviceTypes.mobile++;
          else if (/iPad|Tablet/i.test(ua)) deviceTypes.tablet++;
          else deviceTypes.desktop++;
          break;
        case 'perf_load':
          if (eventData.ms) {
            perfLoads.push(eventData.ms);
            if (userEmail) {
              if (!userLoadTimes[userEmail]) userLoadTimes[userEmail] = [];
              userLoadTimes[userEmail].push(eventData.ms);
            }
          }
          break;
        case 'perf_batch':
          if (eventData.ms) perfBatches.push(eventData.ms);
          break;
        case 'perf_swr':
          if (eventData.hit) swrHits++;
          else swrMisses++;
          break;
        case 'back_swipe':
          backSwipes++;
          break;
        case 'error':
          errors.push({
            user: userEmail,
            msg: String(eventData.msg || '').slice(0, 100),
            ts: ts instanceof Date ? ts.toISOString() : ''
          });
          break;
      }
    }

    // Compute averages
    var avgLoadMs = perfLoads.length > 0
      ? Math.round(perfLoads.reduce(function(a, b) { return a + b; }, 0) / perfLoads.length)
      : null;
    var avgBatchMs = perfBatches.length > 0
      ? Math.round(perfBatches.reduce(function(a, b) { return a + b; }, 0) / perfBatches.length)
      : null;
    var p95LoadMs = perfLoads.length > 0
      ? (function() { var s = perfLoads.slice().sort(function(a, b) { return a - b; }); return s[Math.floor(s.length * 0.95)]; })()
      : null;

    // Per-user load time averages (sorted slowest first)
    var userPerfList = [];
    for (var ue in userLoadTimes) {
      var times = userLoadTimes[ue];
      var avg = Math.round(times.reduce(function(a, b) { return a + b; }, 0) / times.length);
      userPerfList.push({ email: ue, avgMs: avg, samples: times.length });
    }
    userPerfList.sort(function(a, b) { return b.avgMs - a.avgMs; });

    // Daily session counts (for trend chart)
    var dailyData = [];
    for (var dk in dailySessions) {
      dailyData.push({ date: dk, sessions: Object.keys(dailySessions[dk]).length });
    }
    dailyData.sort(function(a, b) { return a.date < b.date ? -1 : 1; });

    // Top tabs (sorted by count)
    var topTabs = [];
    for (var tk in tabCounts) {
      topTabs.push({ tab: tk, count: tabCounts[tk] });
    }
    topTabs.sort(function(a, b) { return b.count - a.count; });

    return {
      success: true,
      days: days,
      totalEvents: totalEvents,
      filteredEvents: filteredEvents,
      uniqueUsers: Object.keys(uniqueUsers).length,
      totalSessions: Object.keys(sessions).length,
      authMethods: authMethods,
      performance: {
        avgLoadMs: avgLoadMs,
        avgBatchMs: avgBatchMs,
        p95LoadMs: p95LoadMs,
        swrHitRate: (swrHits + swrMisses) > 0 ? Math.round(100 * swrHits / (swrHits + swrMisses)) : null,
        sampleCount: perfLoads.length
      },
      backSwipes: backSwipes,
      deviceTypes: deviceTypes,
      topTabs: topTabs.slice(0, 20),
      dailySessions: dailyData,
      userPerf: userPerfList.slice(0, 20),
      recentErrors: errors.slice(-20).reverse(),
    };
  } catch (err) {
    log_('dataGetUsageAnalytics error', err.message);
    return { success: false, message: 'Analytics query failed.' };
  }
}

// ============================================================================
// WEBAPP MEETING CHECK-IN (v4.43.0)
// ============================================================================

/**
 * Returns the first active/eligible meeting for today, or null.
 * Also checks whether the given email has already checked in.
 * Used by batch data to power the in-app check-in banner.
 *
 * @param {string} email - Member email
 * @returns {Object|null} { meetingId, meetingName, meetingTime, meetingType, alreadyCheckedIn }
 * @private
 */
function _getActiveMeetingForCheckIn(email) {
  try {
    if (typeof getCheckInEligibleMeetings !== 'function') return null;
    var result = getCheckInEligibleMeetings();
    if (!result || !result.success || !result.meetings || result.meetings.length === 0) return null;

    var meeting = result.meetings[0]; // First eligible meeting

    // Check if this member already checked in
    var alreadyCheckedIn = false;
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
        if (sheet && sheet.getLastRow() >= 2) {
          var data = sheet.getDataRange().getValues();
          var emailLower = String(email).toLowerCase().trim();
          for (var i = 1; i < data.length; i++) {
            var rowMeetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
            var rowEmail = String(data[i][MEETING_CHECKIN_COLS.EMAIL - 1] || '').toLowerCase().trim();
            if (rowMeetingId === meeting.id && rowEmail === emailLower) {
              alreadyCheckedIn = true;
              break;
            }
          }
        }
      }
    } catch (_e) { log_('_getActiveMeetingForCheckIn alreadyCheckedIn check', _e.message); }

    return {
      meetingId: meeting.id,
      meetingName: meeting.name,
      meetingTime: meeting.time || '',
      meetingType: meeting.type || '',
      alreadyCheckedIn: alreadyCheckedIn
    };
  } catch (_e) {
    log_('_getActiveMeetingForCheckIn error', _e.message);
    return null;
  }
}

/**
 * One-tap meeting check-in for authenticated webapp users.
 * The user is already logged in (session-authenticated), so no PIN is needed.
 * Looks up the member by their session email and records attendance.
 *
 * @param {string} sessionToken - Client session token
 * @param {string} meetingId - The meeting to check into
 * @returns {Object} { success, memberName, message } or { success: false, error }
 */
function dataWebAppCheckIn(sessionToken, meetingId) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return { success: false, error: 'Please log in to check in.', authError: true };

  if (!meetingId) return errorResponse('Meeting ID is required');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return errorResponse('System temporarily unavailable');

  // Look up member by email
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberSheet) return errorResponse('System error: Member directory not found');

  var memberData = memberSheet.getDataRange().getValues();
  var memberId = null;
  var memberName = '';
  var emailLower = String(email).toLowerCase().trim();

  for (var i = 1; i < memberData.length; i++) {
    var rowEmail = String(memberData[i][MEMBER_COLS.EMAIL - 1] || '').trim().toLowerCase();
    if (rowEmail === emailLower) {
      memberId = String(memberData[i][MEMBER_COLS.MEMBER_ID - 1] || '');
      memberName = String(memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' +
                   String(memberData[i][MEMBER_COLS.LAST_NAME - 1] || '');
      break;
    }
  }

  if (!memberId) return errorResponse('No member found with your email address');

  // Acquire lock to prevent TOCTOU race
  var checkInLock = LockService.getScriptLock();
  if (!checkInLock.tryLock(10000)) {
    return errorResponse('Check-in temporarily unavailable — please try again.');
  }
  try {
    var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!checkInSheet) return errorResponse('Meeting check-in sheet not found');

    var checkInData = checkInSheet.getDataRange().getValues();

    // Check duplicate
    for (var j = 1; j < checkInData.length; j++) {
      var rowMeetingId = String(checkInData[j][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      var rowMemberId = String(checkInData[j][MEETING_CHECKIN_COLS.MEMBER_ID - 1] || '');
      if (rowMeetingId === meetingId && rowMemberId === memberId) {
        return { success: true, memberName: memberName.trim(), message: 'You are already checked in!', alreadyCheckedIn: true };
      }
    }

    // Find meeting details
    var meetingName = '';
    var meetingDate = '';
    var meetingType = '';
    var meetingFound = false;
    for (var k = 1; k < checkInData.length; k++) {
      if (String(checkInData[k][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
        meetingName = checkInData[k][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '';
        meetingDate = checkInData[k][MEETING_CHECKIN_COLS.MEETING_DATE - 1] || '';
        meetingType = checkInData[k][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || '';
        meetingFound = true;
        break;
      }
    }

    if (!meetingFound) return errorResponse('Meeting not found or no longer active.');

    // Record check-in
    checkInSheet.appendRow([
      meetingId,
      meetingName,
      meetingDate,
      meetingType,
      memberId,
      escapeForFormula(memberName.trim()),
      new Date(),
      escapeForFormula(email)
    ]);

    // Auto-activate Scheduled meetings
    for (var m = 1; m < checkInData.length; m++) {
      if (String(checkInData[m][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
        var currentStatus = String(checkInData[m][MEETING_CHECKIN_COLS.EVENT_STATUS - 1] || '');
        if (currentStatus === MEETING_STATUS.SCHEDULED) {
          checkInSheet.getRange(m + 1, MEETING_CHECKIN_COLS.EVENT_STATUS).setValue(MEETING_STATUS.ACTIVE);
        }
        break;
      }
    }
  } finally {
    checkInLock.releaseLock();
  }

  // Audit log
  if (typeof logAuditEvent === 'function') {
    logAuditEvent('MEETING_WEBAPP_CHECKIN', {
      meetingId: meetingId,
      memberId: memberId,
      method: 'webapp_session'
    });
  }

  // Badge refresh
  if (typeof _refreshNavBadges === 'function') _refreshNavBadges();

  return {
    success: true,
    memberName: memberName.trim(),
    message: memberName.trim() + ' checked in successfully!'
  };
}

// ═══════════════════════════════════════
// MEMBER LEADER DATA ENDPOINTS
// ═══════════════════════════════════════
// v4.46.0 — Leader Hub: unit-scoped engagement, broadcast, outreach log

/**
 * Auth helper for Member Leader endpoints. Returns leader info or null.
 * @param {string} sessionToken
 * @returns {Object|null} { email, unit, name } or null
 */
function _requireLeaderAuth(sessionToken) {
  var email = _resolveCallerEmail(sessionToken);
  if (!email) return null;
  var user = DataService.findUserByEmail(email);
  if (!user) return null;
  // Allow leaders and stewards/admins to access leader endpoints
  if (user.isLeader || user.isSteward) return { email: email, unit: user.unit || '', name: user.name || '' };
  return null;
}

/**
 * Returns engagement metrics for the leader's unit members.
 * @param {string} sessionToken
 * @returns {Object} { success, unitName, memberCount, avgScore, members[] }
 */
function dataGetLeaderDashboard(sessionToken) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return { success: false, message: 'Leader access required.' };
  if (!leader.unit) return { success: false, message: 'No unit assigned to your profile.' };
  // DataService.getUnitMembers provides fallback roster; EngagementService enriches with scores
  var roster = DataService.getUnitMembers(leader.unit);
  if (typeof EngagementService !== 'undefined' && typeof EngagementService.getScoreboard === 'function') {
    var scores = EngagementService.getScoreboard({ filterUnit: leader.unit, pageSize: 100, sortBy: 'composite' });
    return { success: true, unitName: leader.unit, memberCount: scores.totalRows || 0, avgScore: scores.avgScore || 0, members: scores.items || [] };
  }
  var members = roster.map(function(m) { return { name: m.name, email: m.email, unit: m.unit, composite: null, trend: '' }; });
  return { success: true, unitName: leader.unit, memberCount: members.length, avgScore: 0, members: members };
}

/**
 * Returns members in the leader's unit for broadcast/outreach autocomplete.
 * @param {string} sessionToken
 * @returns {Object[]} [{ name, email }]
 */
function dataGetLeaderUnitMembers(sessionToken) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader || !leader.unit) return [];
  return DataService.getUnitMembers(leader.unit).filter(function(m) {
    return m.email.toLowerCase() !== leader.email;
  }).map(function(m) { return { name: m.name, email: m.email }; });
}

/**
 * Sends a broadcast email to members in the leader's unit.
 * @param {string} sessionToken
 * @param {string} message - Email body
 * @param {string} [subject] - Email subject (defaults to unit update)
 * @returns {Object} { success, sentCount, failedCount, message }
 */
function dataLeaderBroadcast(sessionToken, message, subject) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return { success: false, sentCount: 0, message: 'Leader access required.' };
  if (!message || !message.trim()) return { success: false, sentCount: 0, message: 'Message is required.' };
  if (!leader.unit) return { success: false, sentCount: 0, message: 'No unit assigned to your profile.' };

  var result = DataService.sendUnitBroadcast(leader.email, leader.name, leader.unit, message, subject);

  if (typeof logAuditEvent === 'function' && typeof AUDIT_EVENTS !== 'undefined') {
    logAuditEvent(AUDIT_EVENTS.BROADCAST_SENT || 'LEADER_BROADCAST', {
      leaderEmail: leader.email,
      unit: leader.unit,
      sentCount: result.sentCount,
      failedCount: result.failedCount,
      scope: 'unit'
    });
  }

  return result;
}

/**
 * Logs a member outreach contact from a leader. Reuses the steward contact log.
 * @param {string} sessionToken
 * @param {string} memberEmail
 * @param {string} contactType - 'Phone', 'Email', 'In Person', 'Text'
 * @param {string} notes
 * @param {string} [memberName]
 * @returns {Object} { success, message }
 */
function dataLogLeaderOutreach(sessionToken, memberEmail, contactType, notes, memberName) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return { success: false, message: 'Leader access required.' };
  if (!memberEmail) return { success: false, message: 'Member email is required.' };
  if (!contactType) return { success: false, message: 'Contact type is required.' };
  return withScriptLock_(function() {
    return DataService.logMemberContact(leader.email, memberEmail, contactType, notes || '', '', memberName || '');
  });
}

/**
 * Returns the leader's outreach log (most recent 100 entries).
 * @param {string} sessionToken
 * @returns {Object[]} Array of contact log entries
 */
function dataGetLeaderOutreachLog(sessionToken) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return [];
  return DataService.getStewardContactLog(leader.email);
}

/**
 * Returns the leader's mentor steward info (from mentorship pairings).
 * @param {string} sessionToken
 * @returns {Object|null} { mentorEmail, mentorName, started, notes } or null
 */
function dataGetLeaderMentor(sessionToken) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return null;
  if (typeof MentorshipService === 'undefined' || typeof MentorshipService.getPairings !== 'function') return null;
  var pairings = MentorshipService.getPairings();
  for (var i = 0; i < pairings.length; i++) {
    if (pairings[i].menteeEmail && pairings[i].menteeEmail.toLowerCase() === leader.email) {
      // Resolve mentor name from DataService.findUserByEmail if available
      var mentorRec = DataService.findUserByEmail(pairings[i].mentorEmail);
      return {
        mentorEmail: pairings[i].mentorEmail,
        mentorName: (mentorRec && mentorRec.name) || pairings[i].mentorEmail,
        started: pairings[i].started || '',
        notes: pairings[i].notes || ''
      };
    }
  }
  return null;
}
