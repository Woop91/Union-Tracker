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
  } catch (_e) { log_('_resolveCallerEmail', 'Error resolving active user: ' + (_e.message || _e)); }
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
  if (!sessionToken) return null;
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
Object.freeze(_PIN_RESTRICTED_RESPONSE);

/** @param {string} sessionToken @returns {Object[]} Steward cases. Requires steward auth. */
function dataGetStewardCases(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; if (!_isGrievancesEnabled()) return { success: true, cases: [] }; return DataService.getStewardCases(s); }
/** @param {string} sessionToken @returns {Object} Steward KPIs. Requires steward auth. */
function dataGetStewardKPIs(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; if (!_isGrievancesEnabled()) return { totalCases: 0, overdue: 0, dueSoon: 0, resolved: 0, activeCases: 0 }; return DataService.getStewardKPIs(s); }
/** @param {string} sessionToken @returns {Object[]} Member's own active grievances. Requires auth. PIN-restricted. */
function dataGetMemberGrievances(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; if (!_isGrievancesEnabled()) return []; return DataService.getMemberGrievances(e); }
/** @param {string} sessionToken @returns {Object} Member's closed grievance history. Requires auth. PIN-restricted. */
function dataGetMemberGrievanceHistory(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; if (!_isGrievancesEnabled()) return { success: true, history: [] }; return DataService.getMemberGrievanceHistory(e); }
/** @param {string} sessionToken @param {string} stewardEmail @returns {Object|null} Steward contact info. Requires auth. PIN-restricted. */
function dataGetStewardContact(sessionToken, stewardEmail) {
  if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE;
  var e = _resolveCallerEmail(sessionToken);
  if (!e) { log_('dataGetStewardContact', 'auth failed'); return { success: false, authError: true, message: 'Authentication required.' }; }
  // Security: only allow lookup if caller is a steward OR the target is the caller's assigned steward
  var callerRec = DataService.findUserByEmail(e);
  if (!callerRec) return null;
  var target = (stewardEmail || e).toLowerCase().trim();
  if (!callerRec.isSteward && callerRec.assignedSteward !== target) {
    log_('dataGetStewardContact', 'unauthorized: ' + e + ' tried to look up ' + target);
    return null;
  }
  return DataService.getStewardContact(target);
}

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
  if (!caller) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
  var isSteward = checkWebAppAuthorization('steward', sessionToken).isAuthorized;
  // Members can only update their own record; stewards can pass a target email via updates._targetEmail
  var targetEmail = (isSteward && updates && updates._targetEmail) ? updates._targetEmail : e;
  if (updates && updates._targetEmail) delete updates._targetEmail; // strip internal routing field
  return withScriptLock_(function() { return DataService.updateMemberProfile(targetEmail, updates); });
}
/** @param {string} sessionToken @returns {Object|null} Assigned steward info for caller. Requires auth. PIN-restricted. */
function dataGetAssignedSteward(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return DataService.getAssignedStewardInfo(e); }
/** @param {string} sessionToken @returns {Object[]} Available stewards for self-assign. Requires auth. */
function dataGetAvailableStewards(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; return DataService.getAvailableStewards(e); }
/** @param {string} sessionToken @param {string} memberEmail @param {string} stewardEmail @returns {Object} Assigns steward to member. Requires steward auth. */
function dataAssignSteward(sessionToken, memberEmail, stewardEmail) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.assignStewardToMember(memberEmail, stewardEmail); }); }
// v4.28.2 — Member-safe self-assign: members can assign a steward to THEMSELVES only.
/** @param {string} sessionToken @param {string} stewardEmail @returns {Object} Self-assigns a steward. Requires auth. */
function dataMemberAssignSteward(sessionToken, stewardEmail) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; return withScriptLock_(function() { return DataService.assignStewardToMember(e, stewardEmail); }); }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Starts a grievance draft. Requires auth. PIN-restricted. */
function dataStartGrievanceDraft(sessionToken, data, idemKey) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return withScriptLock_(function() { return DataService.startGrievanceDraft(e, data, idemKey); }); }
/** @param {string} sessionToken @returns {Object} Creates Drive folder for member's grievance. Requires auth. PIN-restricted. */
function dataCreateGrievanceDrive(sessionToken) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return DataService.createGrievanceDriveFolder(e); }
// v4.31.1 — Resource click tracking moved to line ~5423 (3-param version with resourceTitle)
/** @param {string} sessionToken @returns {Object} Survey completion status for caller. Requires auth. */
function dataGetSurveyStatus(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return DataService.getMemberSurveyStatus(e); }
/** @param {string} sessionToken @returns {Object[]} All members from directory. Requires steward auth. */
function dataGetAllMembers(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getAllMembers(); }
/** @param {string} sessionToken @param {string} [scope] @returns {Object} Survey tracking for steward's members. Requires steward auth. */
function dataGetStewardSurveyTracking(sessionToken, scope) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; try { return DataService.getStewardSurveyTracking(s, scope); } catch (e) { log_('dataGetStewardSurveyTracking error', e.message + '\n' + (e.stack || '')); return { total: 0, completed: 0, members: [] }; } }
/** @param {string} sessionToken @param {Object} filter @param {string} msg @param {string} subject @returns {Object} Sends broadcast email. Requires steward auth. */
function dataSendBroadcast(sessionToken, filter, msg, subject) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.sendBroadcastMessage(s, filter, msg, subject); }
/** @param {string} sessionToken @returns {Object} Aggregated survey results. Requires steward auth. */
function dataGetSurveyResults(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getSurveyResults(); }
// v4.21.0 — Native survey engine wrappers
/** @param {string} sessionToken @returns {Object} Survey questions. Requires auth. */
function dataGetSurveyQuestions(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Not authenticated.' }; return getSurveyQuestions(); }
/** @param {string} sessionToken @param {Object} responses @returns {Object} Submits survey response. Requires auth. */
function dataSubmitSurveyResponse(sessionToken, responses) { var e = _resolveCallerEmail(sessionToken); return e ? submitSurveyResponse(e, responses) : { success: false, authError: true, message: 'Authentication required.' }; }
// dataGetPendingSurveyMembers, dataGetSatisfactionSummary, dataOpenNewSurveyPeriod are in 08e_SurveyEngine.gs
/** @param {string} sessionToken @param {string} memberEmail @param {string} type @param {string} notes @param {string} duration @param {string} memberName @returns {Object} Logs member contact. Requires steward auth. */
function dataLogMemberContact(sessionToken, memberEmail, type, notes, duration, memberName) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.logMemberContact(s, memberEmail, type, notes, duration, memberName); }); }
/** @param {string} sessionToken @param {string} memberEmail @returns {Object[]} Contact history for a member. Requires steward auth. */
function dataGetMemberContactHistory(sessionToken, memberEmail) { var s = _requireStewardAuth(sessionToken); if (!s) { log_('dataGetMemberContactHistory', 'auth failed'); return { success: false, authError: true, message: 'Steward access required.' }; } return DataService.getMemberContactHistory(s, memberEmail); }
/** @param {string} sessionToken @returns {Object[]} Full contact log for a steward. Requires steward auth. */
function dataGetStewardContactLog(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) { log_('dataGetStewardContactLog', 'auth failed'); return { success: false, authError: true, message: 'Steward access required.' }; } return DataService.getStewardContactLog(s); }

// S2: Batch badge counts — replaces 3 serial client calls with 1 round-trip
/**
 * Returns notification, task, and Q&A badge counts in a single round-trip. Requires auth.
 * @param {string} sessionToken
 * @returns {Object} { notificationCount, taskCount, overdueTaskCount, qaUnansweredCount }
 */
function dataGetBadgeCounts(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
          logSheet.appendRow([new Date(), sName, 'Email', escapeForFormula(noteText), '']);
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
  if (!s) return { success: false, authError: true, url: null, message: 'Steward access required.' };
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
function dataCreateTask(sessionToken, title, desc, memberEmail, priority, dueDate, assignToEmail, idemKey) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assignToEmail || '', idemKey); }); }
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  if (!DataService.isChiefSteward(s)) return { success: false, message: 'Not authorized.' };
  return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assigneeEmail, idemKey);
}
/** @param {string} sessionToken @param {string} [statusFilter] @returns {Object[]} Steward tasks. Requires steward auth. */
function dataGetTasks(sessionToken, statusFilter) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getTasks(s, statusFilter); }
/** @param {string} sessionToken @param {string} taskId @returns {Object} Completes a steward task. Requires steward auth. */
function dataCompleteTask(sessionToken, taskId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.completeTask(s, taskId); }); }
/** @param {string} sessionToken @returns {Object} Member stats for steward's caseload. Requires auth. */
function dataGetStewardMemberStats(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; try { return DataService.getStewardMemberStats(s); } catch (err) { log_('dataGetStewardMemberStats error', err.message + '\n' + (err.stack || '')); return { total: 0, byLocation: {}, byDues: {} }; } }
/**
 * Returns the steward directory with phone visibility based on caller's role. Requires auth.
 * @param {string} sessionToken
 * @returns {Object[]} Array of steward contact entries
 */
function dataGetStewardDirectory(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
function dataGetNonMemberContacts(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; try { return DataService.getNonMemberContacts(); } catch (err) { log_('dataGetNonMemberContacts error', err.message); return []; } }
/** @param {string} sessionToken @param {Object} data @returns {Object} Add non-member contact. Steward-only. */
function dataAddNonMemberContact(sessionToken, data) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { try { return DataService.addNonMemberContact(data); } catch (err) { log_('dataAddNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }
/** @param {string} sessionToken @param {string} contactId @param {Object} data @returns {Object} Update non-member contact. Steward-only. */
function dataUpdateNonMemberContact(sessionToken, contactId, data) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { try { return DataService.updateNonMemberContact(contactId, data); } catch (err) { log_('dataUpdateNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }
/** @param {string} sessionToken @param {string} contactId @returns {Object} Delete non-member contact. Steward-only. */
function dataDeleteNonMemberContact(sessionToken, contactId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { try { return DataService.deleteNonMemberContact(contactId); } catch (err) { log_('dataDeleteNonMemberContact error', err.message); return { success: false, error: err.message }; } }); }

// ─── ADD MEMBER (webapp) ──────────────────────────────────────────────────────
/** @param {string} sessionToken @returns {Object} Config-driven dropdown options for the Add Member form. Steward-only. */
function dataGetAddMemberOptions(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
function dataUpdateMemberBySteward(sessionToken, memberEmail, updates) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; if (!memberEmail) return { success: false, message: 'Member email required.' }; return withScriptLock_(function() { return DataService.updateMemberBySteward(memberEmail, updates); }); }

/** @param {string} sessionToken @returns {Object} Org-wide grievance statistics. Requires steward auth. */
function dataGetGrievanceStats(sessionToken) { if (!_isGrievancesEnabled()) return { available: false }; var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getGrievanceStats(); }
/** @param {string} sessionToken @returns {Object[]} Grievance hotspot locations. Requires steward auth. */
function dataGetGrievanceHotSpots(sessionToken) { if (!_isGrievancesEnabled()) return []; var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getGrievanceHotSpots(); }
/** @param {string} sessionToken @returns {Object|null} Membership statistics. Requires auth. */
function dataGetMembershipStats(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMembershipStats() : { success: false, authError: true, message: 'Authentication required.' }; }

// v4.28.1 — Member-safe grievance endpoints for Union Stats page.
// Uses _resolveCallerEmail (any authenticated member) instead of _requireStewardAuth.
// Data is already anonymized (aggregate counts only); hotspots require 3+ per location.
/** @param {string} sessionToken @returns {Object} Anonymized grievance stats (member-safe). Requires auth. */
function dataGetMemberGrievanceStats(sessionToken) { if (!_isGrievancesEnabled()) return { available: false }; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; return DataService.getGrievanceStats(); }
/** @param {string} sessionToken @returns {Object[]} Anonymized grievance hotspots (member-safe). Requires auth. */
function dataGetMemberGrievanceHotSpots(sessionToken) { if (!_isGrievancesEnabled()) return []; var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; return DataService.getGrievanceHotSpots(); }

// v4.32.0 — Grievance Feedback wrappers
/** @param {string} sessionToken @returns {Object|null} Pending grievance feedback prompt for caller. Requires auth. */
function dataGetPendingGrievanceFeedback(sessionToken) { if (!_isGrievancesEnabled()) return null; var e = _resolveCallerEmail(sessionToken); return e ? DataService.getPendingGrievanceFeedback(e) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @param {string} grievanceId @param {Object} ratings @param {string} comment @returns {Object} Submits grievance feedback. Requires auth. */
function dataSubmitGrievanceFeedback(sessionToken, grievanceId, ratings, comment) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, authError: true, message: 'Authentication required.' }; if (!_isGrievancesEnabled()) return { success: false, message: 'Grievances disabled.' }; return withScriptLock_(function() { return DataService.submitGrievanceFeedback(e, grievanceId, ratings, comment); }); }
/** @param {string} sessionToken @returns {Object|null} Aggregate grievance feedback stats. Requires auth. */
function dataGetGrievanceFeedbackStats(sessionToken) { if (!_isGrievancesEnabled()) return null; var e = _resolveCallerEmail(sessionToken); return e ? DataService.getGrievanceFeedbackStats() : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @returns {Object|null} Feedback summary for calling steward. Requires steward auth. */
function dataGetStewardFeedbackSummary(sessionToken) { var s = _requireStewardAuth(sessionToken); return s ? DataService.getStewardFeedbackSummary(s) : { success: false, authError: true, message: 'Steward access required.' }; }
/** @param {string} sessionToken @param {number} [limit] @returns {Object[]} Upcoming events. Requires auth. */
function dataGetUpcomingEvents(sessionToken, limit) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getUpcomingEvents(limit) : { success: false, authError: true, message: 'Authentication required.' }; }
// dataGetSurveyQuestions and dataSubmitSurveyResponse are defined in the v4.21.0 block above (single canonical definition)
/** @param {string} sessionToken @returns {boolean} Whether caller is the chief steward. Requires auth. */
function dataIsChiefSteward(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.isChiefSteward(e) : { success: false, authError: true, message: 'Authentication required.' }; }
// dataGetAgencyGrievanceStats — alias removed; frontend uses dataGetGrievanceStats directly

// v4.17.0 — member task assignment wrappers (CR-AUTH-3: server-side identity)
/** @param {string} sessionToken @param {string} memberEmail @param {string} title @param {string} desc @param {string} priority @param {string} dueDate @returns {Object} Creates member task. Requires steward auth. */
function dataCreateMemberTask(sessionToken, memberEmail, title, desc, priority, dueDate) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.createMemberTask(s, memberEmail, title, desc, priority, dueDate); }); }
/** @param {string} sessionToken @param {string} [statusFilter] @returns {Object[]} Tasks assigned to the calling member. Requires auth. */
function dataGetMemberTasks(sessionToken, statusFilter) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMemberTasks(e, statusFilter) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @param {string} taskId @returns {Object} Completes a member task. Requires auth. */
function dataCompleteMemberTask(sessionToken, taskId) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.completeMemberTask(e, taskId) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @returns {Object[]} Member tasks assigned by calling steward. Requires steward auth. */
function dataGetStewardAssignedMemberTasks(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getStewardAssignedMemberTasks(s); }
// BUG-TASKS-03: steward completing a member task on the member's behalf
/** @param {string} sessionToken @param {string} taskId @returns {Object} Steward marks member task complete. Requires steward auth. */
function dataStaffCompleteMemberTask(sessionToken, taskId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.stewardCompleteMemberTask(s, taskId); }

// v4.16.0 — unwired sheet wrappers (CR-AUTH-3: server-side identity + role checks)
/** @param {string} sessionToken @param {string} taskId @param {Object} updates @returns {Object} Updates a steward task. Requires steward auth. */
function dataUpdateTask(sessionToken, taskId, updates) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return withScriptLock_(function() { return DataService.updateTask(s, taskId, updates); }); }
/** @param {string} sessionToken @returns {Object[]} All steward performance metrics. Requires steward auth. */
function dataGetAllStewardPerformance(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getAllStewardPerformance(); }
/** @param {string} sessionToken @param {string} caseId @returns {Object[]} Checklist items for a case. Requires auth. */
function dataGetCaseChecklist(sessionToken, caseId) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getCaseChecklist(caseId) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @param {string} checklistId @param {boolean} completed @returns {Object} Toggles checklist item. Requires auth. */
function dataToggleChecklistItem(sessionToken, checklistId, completed) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.toggleChecklistItem(checklistId, completed, e) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @returns {Object[]} Meetings the caller has attended. Requires auth. */
function dataGetMemberMeetings(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMemberMeetings(e) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @returns {Object} Satisfaction survey trends. Requires steward auth. */
function dataGetSatisfactionTrends(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.getSatisfactionTrends(); }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Submits user feedback. Requires auth. */
function dataSubmitFeedback(sessionToken, data, idemKey) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.submitFeedback(e, data, idemKey) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @returns {Object[]} Caller's submitted feedback. Requires auth. */
function dataGetMyFeedback(sessionToken) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMyFeedback(e) : { success: false, authError: true, message: 'Authentication required.' }; }

// v4.33.0 — Insights batch: 6 parallel server calls in 1 round-trip
/**
 * Returns all Insights page data in a single round-trip. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { stats, hotSpots, perf, sat, memberStats, workload }
 */
function dataGetInsightsBatch(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
function dataGetMeetingMinutes(sessionToken, limit) { var e = _resolveCallerEmail(sessionToken); return e ? DataService.getMeetingMinutes(limit) : { success: false, authError: true, message: 'Authentication required.' }; }
/** @param {string} sessionToken @param {Object} data @param {string} idemKey @returns {Object} Adds meeting minutes. Requires steward auth. */
function dataAddMeetingMinutes(sessionToken, data, idemKey) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, authError: true, message: 'Steward access required.' }; return DataService.addMeetingMinutes(s, data, idemKey); }

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
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { log_('BACKFILL_MINUTES_DRIVE_DOCS', 'Error getting UI: ' + (_e.message || _e)); }

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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
    } catch (_e) { log_('dataGetBroadcastFilterOptions', 'Error reading config: ' + (_e.message || _e)); }
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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!_caller) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!_caller) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!_caller) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
 * Returns live workload summary stats from the Workload Vault sheet. Requires auth.
 * @param {string} sessionToken
 * @returns {Object|null} { avgCaseload, highCaseloadPct, submissionRate, trendDirection }
 */
function dataGetWorkloadSummaryStats(sessionToken) {
  var _caller = _resolveCallerEmail(sessionToken);
  if (!_caller) return { success: false, authError: true, message: 'Authentication required.' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var vault = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
    if (!vault || vault.getLastRow() <= 1) return null;

    var data = vault.getRange(2, 1, vault.getLastRow() - 1, 24).getValues();

    // VAULT_COLS (0-indexed): TIMESTAMP=0 EMAIL=1 PRIORITY_CASES=2 PENDING_CASES=3
    // Most-recent submission per steward
    var latestByEmail = {};
    for (var i = 0; i < data.length; i++) {
      var email = String(data[i][1]).toLowerCase().trim();
      if (!email) continue;
      var ts = data[i][0] instanceof Date ? data[i][0].getTime() : new Date(data[i][0]).getTime();
      if (!latestByEmail[email] || ts > latestByEmail[email].ts) {
        latestByEmail[email] = { ts: ts, priority: Number(data[i][2]) || 0 };
      }
    }

    var stewardEmails = Object.keys(latestByEmail);
    if (stewardEmails.length === 0) return null;

    var totalPriority = 0;
    var highCount = 0;
    var HIGH_THRESHOLD = 5;
    for (var j = 0; j < stewardEmails.length; j++) {
      var p = latestByEmail[stewardEmails[j]].priority;
      totalPriority += p;
      if (p > HIGH_THRESHOLD) highCount++;
    }
    var avgCaseload = totalPriority / stewardEmails.length;
    var highCaseloadPct = Math.round((highCount / stewardEmails.length) * 100);

    // Submission rate — stewards in Member Directory vs those who submitted
    var submissionRate = 100;
    try {
      var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberSheet && memberSheet.getLastRow() >= 2) {
        var mHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
        var isStewardIdx = mHeaders.indexOf('Is Steward');
        var emailIdx     = mHeaders.indexOf('Email');
        if (isStewardIdx >= 0 && emailIdx >= 0) {
          var mData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
          var stewardSet = {};
          for (var mi = 0; mi < mData.length; mi++) {
            if (String(mData[mi][isStewardIdx]).toLowerCase() === 'yes') {
              var se = String(mData[mi][emailIdx]).toLowerCase().trim();
              if (se) stewardSet[se] = true;
            }
          }
          var totalStewards = Object.keys(stewardSet).length;
          if (totalStewards > 0) {
            var submitters = stewardEmails.filter(function(e) { return stewardSet[e]; }).length;
            submissionRate = Math.round((submitters / totalStewards) * 100);
          }
        }
      }
    } catch (_sr) { log_('_sr', (_sr.message || _sr)); }

    // Trend: compare avg priority in last 4 weeks vs prior 4 weeks
    var trendDirection = 'stable';
    try {
      var now = new Date();
      var fourWeeksAgo = new Date(now.getTime() - 28 * 24 * 3600 * 1000);
      var eightWeeksAgo = new Date(now.getTime() - 56 * 24 * 3600 * 1000);
      var recentTotal = 0, recentCount = 0, priorTotal = 0, priorCount = 0;
      for (var ti = 0; ti < data.length; ti++) {
        var tts = data[ti][0] instanceof Date ? data[ti][0] : new Date(data[ti][0]);
        var tp  = Number(data[ti][2]) || 0;
        if (tts >= fourWeeksAgo)    { recentTotal += tp; recentCount++; }
        else if (tts >= eightWeeksAgo) { priorTotal  += tp; priorCount++; }
      }
      if (recentCount > 0 && priorCount > 0) {
        var recentAvg = recentTotal / recentCount;
        var priorAvg  = priorTotal  / priorCount;
        var delta = recentAvg - priorAvg;
        if (delta > 0.5)       trendDirection = 'increasing';
        else if (delta < -0.5) trendDirection = 'decreasing';
      }
    } catch (_td) { log_('_td', (_td.message || _td)); }

    return {
      avgCaseload:      Math.round(avgCaseload * 10) / 10,
      highCaseloadPct:  highCaseloadPct,
      submissionRate:   submissionRate,
      trendDirection:   trendDirection,
    };
  } catch (e) {
    log_('dataGetWorkloadSummaryStats error', e.message + '\n' + (e.stack || ''));
    return null;
  }
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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };

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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };
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
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
    var recordId = String(col_(data[i], AUDIT_LOG_COLS.RECORD_ID) || '').trim();
    if (recordId === targetId) {
      results.push({
        timestamp: col_(data[i], AUDIT_LOG_COLS.TIMESTAMP),
        userEmail: String(col_(data[i], AUDIT_LOG_COLS.USER_EMAIL) || ''),
        field: String(col_(data[i], AUDIT_LOG_COLS.FIELD_NAME) || ''),
        oldValue: String(col_(data[i], AUDIT_LOG_COLS.OLD_VALUE) || ''),
        newValue: String(col_(data[i], AUDIT_LOG_COLS.NEW_VALUE) || ''),
        actionType: String(col_(data[i], AUDIT_LOG_COLS.ACTION_TYPE) || 'edit')
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  if (!Array.isArray(caseIds) || !caseIds.length || !newStatus) return { success: false, message: 'Invalid parameters.' };

  // v4.36.0 — 2FA required for bulk operations
  if (typeof TwoFactorService !== 'undefined' && !TwoFactorService.hasValidSession(s)) {
    return { success: false, requires2FA: true, message: 'Verification required for bulk operations.' };
  }

  return withScriptLock_(function() {
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
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('BULK_STATUS_UPDATE', { count: updated, newStatus: newStatus, steward: s });
    }
    return { success: true, updated: updated };
  });
}

/** Bulk export cases as CSV string. Steward-only. */
function dataBulkExportCsv(sessionToken, caseIds) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
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
// POMS REFERENCE DATA (lazy-loaded by poms_reference.html)
// ═══════════════════════════════════════

/**
 * Returns the POMS reference database for the client-side search tool.
 * Previously inlined as a 49KB const in poms_reference.html — extracted
 * to reduce initial steward-view payload by ~49KB.
 * @param {string} sessionToken
 * @returns {Array} POMS reference records
 */
function dataGetPomsReference(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, authError: true, message: 'Authentication required.' };
  return [
  // ── SEQUENTIAL EVAL ──
  {id:"di-11",s:"DI 22001",t:"Sequential Evaluation - Overview",c:"Sequential Eval",sh:"DI",r:3,tp:"5 step process evaluation framework sequential disability determination",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0422001000",fl:"sequential-eval",
    ex:"THE core framework. All adult claims follow 5 steps in order — cannot skip. If determinable at any step, stop. Step 1=SGA, 2=Severity, 3=Listings, 4=Past Work, 5=Other Work. Children: Steps 1-3 then functional equivalence.",
    rm:["Never skip a step","Favorable at Step 3 → stop, no RFC","Steps 4-5 adults only","Document rationale at each step"],rl:["DI 22005","DI 22010","DI 22015","DI 22020","DI 22025"]},
  {id:"di-12",s:"DI 22005",t:"Step 1 - SGA",c:"Sequential Eval",sh:"DI",r:3,tp:"step 1 SGA substantial gainful activity earnings",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0422005000",
    ex:"Working at SGA? If yes → deny at Step 1. SGA = monthly earnings threshold, updated annually. Blind SGA higher. Subsidies/IRWEs reduce earnings. Usually verified by FO.",
    rm:["SGA changes annually","Blind SGA higher","IRWEs/subsidies reduce earnings","Self-employment: 3 tests"],rl:["DI 10501.015","DI 10505","DI 10520"]},
  {id:"di-13",s:"DI 22010",t:"Step 2 - Severity",c:"Sequential Eval",sh:"DI",r:3,tp:"step 2 severe impairment severity basic work activities de minimis",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0422010000",
    ex:"Severe = significantly limits basic work activities. LOW bar (de minimis). More than minimal effect = severe. Consider ALL impairments combined. Step 2 denial rare.",
    rm:["LOW bar — more than minimal = severe","ALL impairments combined","Step 2 denial rare","Duration (12mo) also here"],rl:["DI 25505"]},
  {id:"di-14",s:"DI 22015",t:"Step 3 - Listings",c:"Sequential Eval",sh:"DI",r:3,tp:"step 3 meets equals listings impairment medical equivalence",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0422015000",
    ex:"Meets = every criterion satisfied. Equals = findings at least equal severity. MC/PC makes equivalence determination. If met → allow without RFC. Part A adults, Part B children.",
    rm:["Check BOTH meets AND equals","MC/PC signature for equivalence","Meets → ALLOW, skip Steps 4-5","Mental: check B AND C criteria"],rl:["DI 34001","DI 34002","DI 24505"]},
  {id:"di-15",s:"DI 22020",t:"Step 4 - Past Relevant Work",c:"Sequential Eval",sh:"DI",r:3,tp:"step 4 PRW past relevant work RFC job demands",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0422020000",
    ex:"Can claimant do past work given RFC? PRW = last 15 years, long enough to learn, at SGA. Compare RFC to demands as actually AND generally performed (DOT). If either works → deny.",
    rm:["PRW: 15 years, learned it, SGA","Compare BOTH ways","Complete RFC before Step 4","Not for children"],rl:["DI 24510","DI 24515","DI 25001"]},
  {id:"di-16",s:"DI 22025",t:"Step 5 - Other Work (Grid Rules)",c:"Sequential Eval",sh:"DI",r:3,tp:"step 5 grid rules vocational age education skill other work",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0422025000",fl:"sequential-eval",
    ex:"Grid tables combine RFC + age + education + experience → disabled/not disabled. Three tables: sedentary, light, medium. Exertional-only → directs. Nonexertional → framework. Burden shifts to SSA.",
    rm:["Grids: sedentary, light, medium","Nonexertional = framework only","55+ sedentary unskilled = usually disabled","Burden on SSA at Step 5"],rl:["DI 25001","DI 25005","DI 25015"]},

  // ── SGA ──
  {id:"di-2",s:"DI 10501.015",t:"SGA General",c:"SGA",sh:"DI",r:3,tp:"SGA definition earnings substantial gainful activity monthly threshold",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0410501015",
    ex:"SGA = earnings threshold. Above SGA → not disabled at Step 1. Updated annually. Separate blind threshold. Applies differently during TWP and EPE.",
    rm:["Changes annually","Blind SGA higher","Gross earnings","Subsidies/IRWEs reduce countable"],rl:["DI 10505","DI 10520"]},
  {id:"di-3",s:"DI 10505",t:"SGA Earnings Guidelines",c:"SGA",sh:"DI",r:3,tp:"monthly SGA amounts blind trial work period TWP EPE",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0410505000",
    ex:"Contains actual SGA dollar amounts by year, TWP service month amounts, and blind SGA thresholds. Reference this for current year figures.",
    rm:["Verify current year amounts","TWP threshold different from SGA","Blind threshold higher","Updated each January"],rl:["DI 10501.015"]},
  {id:"di-4",s:"DI 10520",t:"SGA Deductions (IRWEs/Subsidies)",c:"SGA",sh:"DI",r:3,tp:"IRWEs subsidies impairment related work expenses deductions",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0410520000",
    ex:"IRWEs = impairment-related costs needed to work. Subsidies = difference between pay and actual work value. Both reduce countable earnings, potentially below SGA.",
    rm:["IRWEs: related to impairment AND needed for work","Subsidies need employer docs","Self-employment separate rules","Also SSI: SI 00820.545"],rl:["DI 10501.015","SI 00820.545"]},
  {id:"di-10515",s:"DI 10515",t:"SGA - Self-Employment",c:"SGA",sh:"DI",r:2,tp:"self employment SGA three tests significant services comparable earnings worth",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0410515000",
    ex:"Self-employment SGA uses 3 tests: Test One (significant services + substantial income), Test Two (comparable to non-disabled), Test Three (work worth SGA in terms of value). Applied in order.",
    rm:["Three tests applied in order","'Significant services' = substantial involvement","Different from wage-earner SGA analysis","IRWEs and unincurred expenses apply"],rl:["DI 10501.015","DI 10520"]},

  // ── MEDICAL EVAL ──
  {id:"di-22",s:"DI 24501",t:"Medical Evaluation - General",c:"Medical Eval",sh:"DI",r:3,tp:"RFC assessment symptoms pain MC PC medical consultant role",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0424501000",
    ex:"Master section for MC/PC role. MC (physician) handles physical; PC (psychologist/psychiatrist) handles mental. Every case requires MC/PC involvement: evaluate evidence, review CE requests, assess RFC, check listings.",
    rm:["MC = physical. PC = mental.","MC/PC signs EVERY determination","MC/PC reviews CE requests","SDM authority ended 12/28/2018"],rl:["DI 24505","DI 24510","DI 24515"]},
  {id:"di-23",s:"DI 24505",t:"Mental Impairments (PRT)",c:"Medical Eval",sh:"DI",r:3,tp:"PRT psychiatric review technique B criteria C criteria mental disorders",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0424505000",fl:"listing-12",
    ex:"PRT documents: MDI exists, B criteria ratings (4 areas: understand/remember, interact, concentrate/persist, adapt — none→extreme). 2 marked OR 1 extreme = meets B. C criteria: serious/persistent + marginal adjustment.",
    rm:["PRT REQUIRED for mental cases","B: 4 areas, need 2 marked OR 1 extreme","C: alternative to B","MRFC separate from PRT"],rl:["DI 24515","DI 34001"]},
  {id:"di-24",s:"DI 24510",t:"Physical RFC Assessment",c:"Medical Eval",sh:"DI",r:3,tp:"physical RFC exertional nonexertional limitations residual functional capacity",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0424510000",
    ex:"Physical RFC = maximum still doable. Exertional: sit/stand/walk/lift/carry/push/pull. Nonexertional: postural, manipulative, visual, communicative, environmental. Expressed as exertional level (sedentary→very heavy). Only at Step 4+.",
    rm:["RFC = what they CAN do","Sedentary <10 lbs → very heavy 100+","Nonexertional erodes occupational base","Not needed if Step 3 allow"],rl:["DI 25001","DI 25005"]},
  {id:"di-25",s:"DI 24515",t:"Mental RFC Assessment",c:"Medical Eval",sh:"DI",r:3,tp:"mental RFC MRFC work functions psychological limitations",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0424515000",
    ex:"MRFC = function-by-function mental assessment for Steps 4-5. Areas: understanding/memory, concentration/persistence, social interaction, adaptation. Different from PRT. Both required for mental cases.",
    rm:["MRFC ≠ PRT — both needed","Function-by-function","Moderate PRT ≠ same RFC limitation","Must be consistent with PRT"],rl:["DI 24505"]},

  // ── MED-VOC ──
  {id:"di-26",s:"DI 25001",t:"Medical-Vocational (Grid Rules)",c:"Med-Voc",sh:"DI",r:3,tp:"grid rules RFC vocational factors tables framework directing",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0425001000",
    ex:"Grid combines RFC level + age + education + experience. Three tables: sedentary, light, medium. Exertional-only → directs. Nonexertional → framework. No grids for heavy/very heavy.",
    rm:["Three tables: sedentary, light, medium","Exertional-only → directs","Nonexertional → framework","Advanced age + sedentary + unskilled = disabled"],rl:["DI 25005","DI 25010","DI 25015"]},
  {id:"di-27",s:"DI 25005",t:"Exertional Levels",c:"Med-Voc",sh:"DI",r:3,tp:"sedentary light medium heavy very heavy exertional definitions lifting",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0425005000",
    ex:"Sedentary: lift max 10 lbs, mostly sitting. Light: lift max 20 lbs, frequent 10 lbs, standing/walking 6hrs. Medium: lift max 50 lbs, frequent 25 lbs. Heavy: lift max 100 lbs. Very Heavy: 100+ lbs.",
    rm:["Sedentary ≠ desk job — still lifting up to 10 lbs","Light requires 6 hrs standing/walking","These are SSA definitions, not DOL","RFC must specify exertional level"],rl:["DI 25001","DI 24510"]},
  {id:"di-28",s:"DI 25010",t:"Transferability of Skills",c:"Med-Voc",sh:"DI",r:3,tp:"skill analysis transferability vocational semi-skilled skilled SVP",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0425010000",
    ex:"Skills from past work that can transfer to other jobs within the RFC. Unskilled (SVP 1-2), semi-skilled (SVP 3-4), skilled (SVP 5-9). Transfer requires: similar tools/processes, same industry, minimal vocational adjustment. Advanced age + no transfer = favorable.",
    rm:["SVP 1-2 = unskilled, no transferable skills","Transfer requires similar tools/processes","Advanced age narrows transferability","Very closely related = same/similar industry"],rl:["DI 25001","DI 25015"]},
  {id:"di-29",s:"DI 25015",t:"Age Categories",c:"Med-Voc",sh:"DI",r:3,tp:"younger closely approaching advanced age 50 55 borderline",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0425015000",
    ex:"<50 = younger. 50-54 = closely approaching. 55+ = advanced age. Borderline (within months) may use higher category. Age + RFC + education interact heavily.",
    rm:["<50 younger, 50-54 closely approaching, 55+ advanced","Borderline → consider higher category","60+ sedentary no skills = almost always disabled","Age is vocational, not medical"],rl:["DI 25001","DI 25010"]},
  {id:"di-25020",s:"DI 25020",t:"Education Categories",c:"Med-Voc",sh:"DI",r:2,tp:"education illiterate marginal limited high school college vocational",ct:"DI/DIB",url:"https://secure.ssa.gov/poms.nsf/lnx/0425020000",
    ex:"Illiteracy, marginal education (6th grade or less), limited education (7th-11th), high school+ (GED counts). Education interacts with age and skills. Lower education = more favorable for Grid rules.",
    rm:["Illiterate = cannot read/write simple message","Marginal = 6th grade or less","Limited = 7th-11th grade","GED = high school equivalent"],rl:["DI 25001","DI 25015"]},

  // ── ONSET/DURATION ──
  {id:"di-33",s:"DI 25501",t:"Onset Date",c:"Onset/Duration",sh:"DI",r:3,tp:"EOD establishment AOD alleged onset established onset date",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0425501000",
    ex:"EOD = when disability began per evidence. AOD = claimant's claim. EOD can equal or be later than AOD. Cannot be earlier without consent. Title II: determines waiting period. SSI: benefits from filing month only.",
    rm:["EOD ≤ AOD without consent prohibited","SSI: no retroactivity beyond filing","Title II: 5-month waiting period","Amendment requires SSA-831"],rl:["DI 25505"]},
  {id:"di-34",s:"DI 25505",t:"Duration Requirement",c:"Onset/Duration",sh:"DI",r:3,tp:"12 month duration closed periods continuous impairment",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0425505000",
    ex:"Impairment must last or be expected to last 12+ continuous months or result in death. Closed periods (disability began and ended) are possible. Duration is evaluated at Step 2.",
    rm:["12 months continuous OR expected to result in death","Closed periods allowed","Evaluated at Step 2","Consider combined effect of all impairments"],rl:["DI 25501","DI 22010"]},

  // ── CHILD CLAIMS ──
  {id:"di-30",s:"DI 25201",t:"Title XVI Child Claims",c:"Child Claims",sh:"DI",r:3,tp:"childhood disability functional equivalence child SSI under 18",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0425201000",fl:"child-eval",
    ex:"Different eval than adults. Steps 1-3, then functional equivalence (not RFC/vocational). Disabled if marked in 2 of 6 domains OR extreme in 1. Under 18. Age-18 redetermination uses adult standards.",
    rm:["No RFC or vocational analysis","6 domains evaluated","Marked 2 OR extreme 1","Age-18 → adult standards"],rl:["DI 25205","DI 25210","DI 34002"]},
  {id:"di-31",s:"DI 25205",t:"Functional Equivalence - 6 Domains",c:"Child Claims",sh:"DI",r:3,tp:"six domains acquiring attending interacting moving caring health marked extreme",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0425205000",fl:"child-eval",
    ex:"(1) Acquiring/Using Info (2) Attending/Completing Tasks (3) Interacting (4) Moving/Manipulating (5) Caring for Self (6) Health/Well-Being. Rated: none → extreme. Need marked in 2 OR extreme in 1.",
    rm:["All impairments combined across domains","Age-appropriate peer comparison","Teachers/parents key evidence","Rated: none, less than marked, marked, extreme"],rl:["DI 25201","DI 25210"]},
  {id:"di-32",s:"DI 25210",t:"Age-18 Redetermination",c:"Child Claims",sh:"DI",r:3,tp:"transition adult standard 18 redetermination child to adult",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0425210000",
    ex:"When SSI child turns 18, disability is redetermined using adult standards (5-step sequential evaluation). No longer uses functional equivalence. This is a new determination, not a CDR.",
    rm:["New determination, not a CDR","Adult standards apply","Functional equivalence no longer used","Must be completed within 1 year of 18th birthday"],rl:["DI 25201","DI 22001"]},

  // ── CDR ──
  {id:"di-35",s:"DI 28001",t:"CDR Overview",c:"CDR",sh:"DI",r:3,tp:"medical improvement CDR continuing disability review MIRS diary",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0428001000",fl:"cdr-mirs",
    ex:"CDR asks: medical improvement related to work ability? Uses MIRS (different from initial). Diary: MIE (6-18mo), MIP (3yr), MINE (5-7yr). CPD = prior favorable decision.",
    rm:["MIRS ≠ initial eval standard","Must show improvement related to work","CPD = comparison point","MIE/MIP/MINE set at allowance"],rl:["DI 28005","DI 28075.005"]},
  {id:"di-36",s:"DI 28005",t:"Medical Improvement Review Standard",c:"CDR",sh:"DI",r:3,tp:"MIRS comparison point exceptions medical improvement standard groups 8 step CDR process",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0428005000",fl:"cdr-mirs",
    ex:"MIRS has 8 steps. Compare current medical condition to CPD. If improved AND improvement related to work ability → evaluate RFC. Exceptions exist (e.g., fraud, new medical evidence, technical improvement). Two groups of exceptions.",
    rm:["8-step CDR process","Compare to CPD","Group I exceptions: benefits cease","Group II exceptions: continue benefits review"],rl:["DI 28001"]},

  // ── LISTINGS ──
  {id:"di-37",s:"DI 34001",t:"Listing of Impairments - Adult (Part A)",c:"Listings",sh:"DI",r:3,tp:"adult listings body systems musculoskeletal respiratory cardiovascular mental neurological cancer",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",
    ex:"Part A = adult Listings by body system: 1.00 Musculoskeletal, 2.00 Senses/Speech, 3.00 Respiratory, 4.00 Cardiovascular, 5.00 Digestive, 6.00 Genitourinary, 7.00 Hematological, 8.00 Skin, 9.00 Endocrine, 10.00 Congenital, 11.00 Neurological, 12.00 Mental, 13.00 Cancer, 14.00 Immune.",
    rm:["Body systems 1.00-14.00","Check meets AND equals","Updated periodically","Some require specific test results"],rl:["DI 22015","DI 34002"]},
  {id:"di-38",s:"DI 34002",t:"Listing of Impairments - Child (Part B)",c:"Listings",sh:"DI",r:3,tp:"childhood listings Part B body systems child impairments",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434002000",
    ex:"Part B = childhood Listings. Similar body system structure to Part A but with age-appropriate criteria. Some listings reference Part A with modifications. Low birth weight listings unique to Part B.",
    rm:["Age-appropriate criteria","Some reference Part A","Low birth weight unique to Part B","If doesn't meet listing → functional equivalence"],rl:["DI 34001","DI 25201"]},

  // ── DAA ──
  {id:"di-39",s:"DI 90070",t:"DAA Evaluation Steps",c:"DAA",sh:"DI",r:3,tp:"drug addiction alcoholism DAA materiality substance abuse",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0490070000",fl:"daa-eval",
    ex:"First: disabled with ALL impairments (incl DAA)? If yes: would remaining limitations disable WITHOUT DAA? Yes → not material (allowed). No → material (denied). If allowed with DAA → mandatory rep payee + treatment referral.",
    rm:["Evaluated AFTER finding disability","Key: disabled WITHOUT DAA?","Material → denied","Allowed w/DAA → rep payee"],rl:[]},

  // ── SPECIAL ISSUES ──
  {id:"di-20",s:"DI 23007",t:"Failure to Cooperate - Overview",c:"Special Issues",sh:"DI",r:3,tp:"FTC failure cooperate insufficient evidence",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007000",fl:"ftc-process",
    ex:"FTC: claimant won't cooperate with evidence/CE. Purpose: ENCOURAGE cooperation. Special handling for vulnerable populations. Document all contacts. After exhausting procedures → determine on file.",
    rm:["Purpose: ENCOURAGE, not punish","Special: <18, ≥65, homeless, MI, LEP","Phone → letter → third party","PDN must explain"],rl:["DI 23007.001","DI 23007.005","DI 23007.009","DI 23007.010"]},
  {id:"di-21b",s:"DI 23022",t:"Terminal Illness (TERI)",c:"Special Issues",sh:"DI",r:3,tp:"TERI terminal illness expedited processing flag priority",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0423022000",
    ex:"TERI = expected to result in death. Process ASAP. Don't wait for evidence if allowance supportable. Prioritize above all other work.",
    rm:["Process ASAP — top priority","Don't wait if allowable","CAL often overlaps with TERI","FO or DDS can set TERI flag"],rl:[]},
  {id:"di-cal",s:"DI 23020",t:"Compassionate Allowances (CAL)",c:"Special Issues",sh:"DI",r:3,tp:"CAL compassionate allowances expedited conditions list",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0423020000",
    ex:"CAL conditions are so severe they obviously meet disability standard with minimal objective medical evidence. Over 260 conditions (rare diseases, cancers, brain disorders). Identified by system flag or DDS. Process quickly — usually allow at Step 3.",
    rm:["260+ conditions on CAL list","Minimal evidence needed","Usually allow at Step 3","Check SSA CAL list online"],rl:["DI 23022"]},
  {id:"di-qdd",s:"DI 11011",t:"Quick Disability Determination (QDD)",c:"Initial Claims",sh:"DI",r:3,tp:"QDD predictive model expedited fast track quick",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0411011000",
    ex:"QDD uses a predictive model to identify claims with high probability of allowance. Flagged cases should be processed within 20 days if possible. DDS still makes full determination — QDD is just prioritization, not a predetermined outcome.",
    rm:["Predictive model — not predetermined","Target 20 days processing","DDS still does full evaluation","Flag doesn't guarantee allowance"],rl:["DI 23020"]},

  // ── INITIAL CLAIMS ──
  {id:"di-5",s:"DI 11005",t:"Disability Interviews",c:"Initial Claims",sh:"DI",r:2,tp:"interview procedures forms documentation field office",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0411005000",
    ex:"FO conducts disability interviews, collects forms (SSA-3368, SSA-3369, SSA-827, etc.), documents work history, and routes case to DDS. Quality of interview affects DDS development needs.",
    rm:["SSA-3368 = Function Report","SSA-3369 = Work History","SSA-827 = Authorization for Source","Quality intake → less DDS development"],rl:["DI 11010"]},
  {id:"di-8",s:"DI 11015",t:"Presumptive Disability/Blindness",c:"Initial Claims",sh:"DI",r:3,tp:"PD PB presumptive categories emergency advance SSI",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0411015000",
    ex:"PD/PB allows immediate SSI payments before formal determination if the impairment is likely to be found disabling. Categories include: total blindness, total deafness, amputation of leg at hip, bed-bound/immobile, allegation of HIV with specific findings. Up to 6 months of presumptive payments.",
    rm:["SSI only — not Title II","Up to 6 months payments","Specific categories qualify","Still need formal determination"],rl:["SI 00501"]},
  {id:"di-9",s:"DI 11055",t:"SSI Disability Cases",c:"Initial Claims",sh:"DI",r:2,tp:"SSI specific processing concurrent claims title XVI",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0411055000",
    ex:"SSI-specific disability processing rules. Concurrent claims (Title II + XVI filed together) follow specific routing. SSI has no 5-month waiting period, no retroactive benefits beyond filing month.",
    rm:["No waiting period for SSI","No retroactivity beyond filing month","Concurrent = both Title II + XVI","Special routing rules"],rl:["SI 00501"]},

  // ── DDS PROCEDURES ──
  {id:"di-10",s:"DI 20101",t:"DDS Jurisdiction",c:"DDS Procedures",sh:"DI",r:3,tp:"DDS authority jurisdiction transfers state agency",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0420101000",
    ex:"DDS has jurisdiction over initial disability determinations and most CDRs within its state. Cases can be transferred between DDSs when claimant moves. Federal components (ODAR, OQR) handle certain cases.",
    rm:["State DDS handles initials + most CDRs","Transfers when claimant moves","Federal components for appeals","Jurisdiction based on claimant residence"],rl:[]},
  {id:"di-22505",s:"DI 22505",t:"Case Development - General",c:"Case Development",sh:"DI",r:3,tp:"evidence gathering MER medical evidence CE ordering development sufficiency",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0422505000",
    ex:"Master section for evidence development. Priority: (1) MER from treating sources, (2) Other medical sources, (3) CE only if needed. Sufficiency = enough evidence to determine. Document all development efforts. 12-month lookback for MER.",
    rm:["MER first, always","12-month lookback for MER","Document all development efforts","CE is last resort"],rl:["DI 22510.001","DI 22515"]},

  // ── APPEALS ──
  {id:"di-recon",s:"DI 12015",t:"Reconsideration",c:"Appeals",sh:"DI",r:2,tp:"reconsideration recon appeal first level disability hearing",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0412015000",
    ex:"First level of appeal. New examiner and MC/PC at DDS review entire case de novo. Claimant can submit new evidence. Some states have disability hearing (face-to-face) at recon level. 60-day filing deadline from initial determination.",
    rm:["New examiner + MC/PC review","De novo — fresh look at everything","60-day filing deadline","Some states: disability hearing at recon"],rl:[]},

  // ═══ CE SECTIONS ═══
  {id:"ce-1",s:"DI 22510.001",t:"Introduction to CEs",c:"CE Overview",sh:"CE",r:3,tp:"CE definition purchased exam SSA telehealth THCE",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510001",wf:"1-Before ordering",fl:"ce-workflow",
    ex:"CE = purchased exam when MER insufficient. In-person, telehealth, video. Source: licensed, no sanctions. SSA pays. LAST RESORT after MER.",
    rm:["Last resort — MER first","SSA pays, free to claimant","Licensed, no sanctions","Telehealth = claimant agreement"],rl:["DI 22510.005","DI 22505"]},
  {id:"ce-2",s:"DI 22510.005",t:"When to Purchase a CE",c:"CE Overview",sh:"CE",r:3,tp:"order CE insufficient conflict ambiguity MER first",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510005",wf:"1-Before ordering",
    ex:"Two grounds: insufficient evidence OR conflict/ambiguity. Exhaust MER. MC/PC approves. Don't order if decidable on current evidence.",
    rm:["Insufficient OR conflict","MER first, document efforts","MC/PC approves","Don't order if decidable"],rl:["DI 22510.001","DI 22510.006"]},
  {id:"ce-3",s:"DI 22510.006",t:"When NOT to Purchase a CE",c:"CE Overview",sh:"CE",r:3,tp:"do not order sufficient evidence claimant refuses DLI favorable",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510006",wf:"1-Before ordering",
    ex:"Don't order if: sufficient evidence exists, claimant refuses (see 23007.009), DLI passed with no retroactive possibility, or favorable possible on current evidence.",
    rm:["Sufficient evidence = no CE","Refusal → DI 23007.009","Past DLI + no retroactive = don't","Favorable possible = don't"],rl:["DI 22510.005","DI 23007.009"]},
  {id:"ce-4",s:"DI 22510.010",t:"Selecting CE Source",c:"Source Selection",sh:"CE",r:3,tp:"treating source preferred independent licensed qualified provider",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510010",wf:"2-Source selection",
    ex:"Preference order: (1) Treating source (preferred — knows the claimant), (2) Independent qualified source. Source must be qualified for the specific exam type. Cannot use sources with disqualification or exclusion from SSA programs.",
    rm:["Treating source preferred","Must be qualified for specific exam","No disqualified/excluded sources","Check SSA provider list"],rl:["DI 22510.013"]},
  {id:"ce-5",s:"DI 22510.013",t:"Telehealth CE (THCE)",c:"Source Selection",sh:"CE",r:2,tp:"telehealth THCE audio video remote CE agreement",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510013",wf:"2-Source selection",
    ex:"THCE = audio-video CE conducted remotely. Requires claimant agreement. Some exams not feasible via telehealth (those requiring physical contact). DDS determines if THCE is appropriate for the specific exam type.",
    rm:["Requires claimant agreement","Not all exams feasible remotely","DDS determines appropriateness","Physical contact exams = in-person only"],rl:["DI 22510.010"]},
  {id:"ce-6",s:"DI 22510.016",t:"CE Notice & Confirmation",c:"CE Scheduling",sh:"CE",r:3,tp:"phone first consequences written notice third party special handling interpreter",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510016",wf:"3-Scheduling",fl:"ce-workflow",
    ex:"(1) Phone first (2) Explain consequences (3) Written notice (4) Special handling: third party, free interpreter, accommodations (5) Notify rep.",
    rm:["ALWAYS phone first","Must explain consequences","Free interpreter for LEP","Special handling: proactive third party"],rl:["DI 22510.019"],al:"Must explain consequences of non-attendance"},
  {id:"ce-7",s:"DI 22510.016 (Alt)",t:"Short-Notice CE (≤10 days)",c:"CE Scheduling",sh:"CE",r:3,tp:"short notice 10 days not FTC must reschedule",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510016",wf:"3-Scheduling",
    ex:"≤10 days notice: standard follow-up rules don't apply. CRITICAL: no-show is NOT FTC per DI 23007.010B. MUST reschedule. Protects claimants.",
    rm:["≤10 days = short notice","Follow-up rules don't apply","No-show is NOT FTC","MUST reschedule"],rl:["DI 23007.010"],al:"Short-notice CE (≤10 days): no-show is NOT FTC — MUST reschedule"},
  {id:"ce-9",s:"DI 22510.019",t:"CE Follow-Up & Reminder",c:"CE Follow-Up",sh:"CE",r:3,tp:"follow up reminder third party homeless mental immediately",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510019",wf:"4-Follow-up",
    ex:"After notice: confirm call, reminder. Homeless/MI → third party IMMEDIATELY. Assist homeless with travel. Document all follow-up.",
    rm:["Homeless/MI: third party NOW","Assist with travel","Reminder before CE","Document everything"],rl:["DI 22510.016","DI 23007.010"],al:"Homeless/MI: involve third party IMMEDIATELY"},
  {id:"ce-10",s:"DI 23007.001",t:"FTC Definitions",c:"Missed CE",sh:"CE",r:3,tp:"failure cooperate definition special handling under 18 over 65 homeless MI LEP third party good reason",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007001",wf:"5-Missed CE",fl:"missed-ce",
    ex:"FTC = won't cooperate. Special handling: ALL <18, ≥65 w/o rep, homeless, MI, LEP. Third party: SSA-3368 §2 or SSA-3441 §1.D. Good reasons: illness, transport, weather, death, limitations.",
    rm:["Special: <18, ≥65, homeless, MI, LEP","Third party: SSA-3368/3441","Don't reuse prior third parties","Good reasons: illness, transport, weather"],rl:["DI 23007.005","DI 23007.009","DI 23007.010"]},
  {id:"ce-11",s:"DI 23007.005",t:"Contacting Claimant for FTC",c:"Missed CE",sh:"CE",r:3,tp:"phone 10 days call-in letter representative third party document",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007005",wf:"5-Missed CE",
    ex:"Phone → explain consequences → 10 calendar days. If unreachable → call-in letter → 10 more days. Contact rep if appointed. For special handling: third party simultaneously. Document every attempt.",
    rm:["Phone first, always","10 calendar days each round","Document EVERY attempt","Special handling: simultaneous third party"],rl:["DI 23007.001","DI 23007.009"],al:"10 calendar days to comply after each contact"},
  {id:"ce-12",s:"DI 23007.009",t:"Refusal to Attend CE",c:"Missed CE",sh:"CE",r:3,tp:"refusal will not attend good reason treating source reschedule PDN",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007009",wf:"5-Missed CE",fl:"missed-ce",
    ex:"REFUSAL = explicit 'I will not attend' without good reason. ≠ no-show. Good reason → reschedule. Treating source says no → STOP CE entirely. No good reason → determine on file.",
    rm:["Refusal ≠ no-show","Good reason → reschedule","TS advised no → STOP development","No reason → determine on file"],rl:["DI 23007.010","DI 23007.015"],al:"Good reason = reschedule. TS no = STOP CE."},
  {id:"ce-13",s:"DI 23007.010",t:"Failure to Attend CE",c:"Missed CE",sh:"CE",r:3,tp:"no show failure attend different refusal third party reschedule short notice",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007010",wf:"5-Missed CE",fl:"missed-ce",
    ex:"No-show ≠ refusal. Phone → 10d → letter → 10d. Special handling → third party simultaneously. Good reason → reschedule. ≤10-day CE no-show = NOT FTC, must reschedule.",
    rm:["No-show ≠ refusal","Phone → 10d → letter → 10d","≤10-day no-show = NOT FTC","Special: third party simultaneous"],rl:["DI 23007.009","DI 23007.015"],al:"No-show ≠ refusal. ≤10-day = NOT FTC."},
  {id:"ce-14",s:"DI 23007.015",t:"Determining on Evidence in File",c:"Missed CE",sh:"CE",r:3,tp:"FTC exhausted determine evidence file PDN favorable unfavorable",url:"https://secure.ssa.gov/poms.nsf/lnx/0423007015",wf:"6-Determination",
    ex:"After ALL procedures exhausted → determine on available evidence. NOT automatic denial. Normal sequential eval. PDN must explain: what requested, FTC, basis of determination.",
    rm:["Not automatic denial","PDN explains FTC","Exhaust ALL procedures first","Normal sequential eval"],rl:["DI 23007.009","DI 23007.010"]},
  {id:"ce-15",s:"DI 28075.005",t:"CDR FTC/WU",c:"CDR FTC",sh:"CE",r:3,tp:"CDR FTC cessation benefits stopped first month aware whereabouts unknown",url:"https://secure.ssa.gov/poms.nsf/lnx/0428075005",wf:"CDR context",fl:"cdr-ftc",
    ex:"CDR FTC → cessation. Cessation month = first month aware + knew repercussions + failed. Benefits STOP. Different from initial FTC.",
    rm:["CDR FTC → CESSATION","First month aware + knew + failed","Different from initial","Document everything"],rl:["DI 28001"],al:"CDR FTC → cessation of benefits"},
  {id:"ce-16",s:"DI 22510.020",t:"Reviewing CE Reports",c:"CE Quality",sh:"CE",r:3,tp:"CE report review adequacy deficient correction completeness",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510020",wf:"7-Post-CE",
    ex:"Check: completeness, clinical findings (objective), consistency, supported diagnoses. Deficient → corrections. Don't accept inadequate reports.",
    rm:["Completeness, findings, consistency","Deficient → request corrections","Don't accept inadequate reports","Delays case if redo"],rl:["DI 22510.070"]},
  {id:"ce-17",s:"DI 22510.070",t:"Life-Threatening CE Finding",c:"CE Quality",sh:"CE",r:3,tp:"life threatening immediate notification emergency",url:"https://secure.ssa.gov/poms.nsf/lnx/0422510070",wf:"7-Post-CE",
    ex:"CE reveals life-threatening condition → IMMEDIATE notification. CE source should notify. If not, DDS ensures it. Safety obligation.",
    rm:["IMMEDIATE — no delay","Source notifies claimant","DDS ensures if source didn't","Document in file"],rl:["DI 22510.020"],al:"IMMEDIATE action — life-threatening"},
  {id:"ce-18",s:"DI 39545.275",t:"Missed CE Payment",c:"CE Fiscal",sh:"CE",r:2,tp:"no pay missed CE nominal fee record review",url:"https://secure.ssa.gov/poms.nsf/lnx/0439545275",wf:"Admin",
    ex:"No-pay for missed CEs. Nominal fee possible for record review time only.",
    rm:["No-pay policy","Nominal fee: record review only","State fee schedule","Administrative matter"],rl:[]},
  {id:"ce-19",s:"DI 11018.005",t:"FTC at Field Office",c:"FO FTC",sh:"CE",r:2,tp:"FO denial codes 000M5 000M6 field office DLI closeout",url:"https://secure.ssa.gov/poms.nsf/lnx/0411018005",wf:"FO context",
    ex:"FO handles certain FTC situations (failure to file, whereabouts unknown). Denial codes 000M5 (FTC) and 000M6 (whereabouts unknown). FO does the closeout, not DDS.",
    rm:["FO jurisdiction, not DDS","000M5 = FTC at FO","000M6 = whereabouts unknown","Check DLI before closeout"],rl:[]},
  {id:"ce-20",s:"DI 25205.020",t:"Child Claim FTC",c:"Child CE",sh:"CE",r:3,tp:"child under 18 special efforts FTC childhood additional steps",url:"https://secure.ssa.gov/poms.nsf/lnx/0425205020",wf:"Special populations",
    ex:"ALL child claims (<18) require special FTC efforts. Additional steps beyond standard FTC. Must contact parents/guardians. Document additional efforts for children.",
    rm:["ALL <18 = special handling","Contact parents/guardians","Additional steps required","Document extra efforts"],rl:["DI 23007.001"],al:"ALL <18 require special FTC handling"},

  // ═══ SI SECTIONS ═══
  {id:"si-1",s:"SI 00501",t:"SSI Eligibility Overview",c:"Eligibility",sh:"SI",r:3,tp:"age blindness disability requirements residency citizenship SSI",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500501000",
    ex:"SSI: age 65+ OR blind OR disabled + income limits + resources ($2K/$3K) + US resident + citizen/qualified alien. Needs-based, no work history. General revenue funded.",
    rm:["Needs-based — no work history","$2K/$3K resources","US resident + citizen/alien","Benefits: month after filing"],rl:["SI 00810","SI 01110"]},
  {id:"si-2",s:"SI 00510",t:"Living Arrangements",c:"Eligibility",sh:"SI",r:2,tp:"FLA codes household ISM living arrangement in-kind",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500510000",
    ex:"Living arrangement codes affect SSI payment amount. Determines if ISM (In-Kind Support and Maintenance) rules apply. Own household vs. another's household vs. institution. FLA code set by FO.",
    rm:["FLA affects payment amount","Own household vs. another's","Institutional rules different","FO determines FLA code"],rl:["SI 00835"]},
  {id:"si-3",s:"SI 00520",t:"Institutionalization",c:"Eligibility",sh:"SI",r:2,tp:"public institutions $30 payment Medicaid institutionalized",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500520000",
    ex:"Public institution residents generally ineligible for SSI. Exception: $30 payment limit for Medicaid-eligible institutionalized individuals. Temporary absences may maintain eligibility.",
    rm:["Public institution = generally ineligible","$30 payment exception","Medicaid-eligible exception","Temporary absences may maintain"],rl:["SI 00501"]},
  {id:"si-4",s:"SI 00810",t:"Income Rules - General",c:"Income",sh:"SI",r:3,tp:"income definition types earned unearned in-kind deemed",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500810000",
    ex:"Income = cash/in-kind for food/shelter. Earned (wages, self-employment) + unearned (pensions, SS, gifts). Exclusions: $20 general + $65 earned + 50% remaining earned.",
    rm:["Two types: earned + unearned","$20 + $65 + ½ exclusions","ISM is also income","See SI 00815 for exclusions"],rl:["SI 00815","SI 00820","SI 00835"]},
  {id:"si-5",s:"SI 00815",t:"What Is Not Income",c:"Income",sh:"SI",r:2,tp:"medical social services exclusions not income",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500815000",
    ex:"Items that are NOT income for SSI: medical care/services, social services, receipts from sale of resources, income tax refunds, weatherization assistance, and many others.",
    rm:["Medical care not income","Tax refunds not income","Replacement of lost resource not income","Check full list"],rl:["SI 00810"]},
  {id:"si-6",s:"SI 00820",t:"Earned Income",c:"Income",sh:"SI",r:3,tp:"wages self-employment sheltered workshop earned income",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500820000",
    ex:"Earned income = wages, net earnings from self-employment, payments from sheltered workshops, royalties (active), and certain other forms. $65 exclusion + ½ remaining earned income before reducing SSI payment.",
    rm:["$65 exclusion + ½ remaining","Sheltered workshop = earned","Self-employment = net earnings","Applied after $20 general exclusion"],rl:["SI 00810","SI 00820.545"]},
  {id:"si-7",s:"SI 00830",t:"Unearned Income",c:"Income",sh:"SI",r:3,tp:"unearned income pension annuity Social Security deemed",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500830000",
    ex:"Unearned income: pensions, SS benefits, annuities, gifts, prizes, interest, dividends, rents, deemed income from spouse/parent/sponsor. $20 general exclusion applied first.",
    rm:["SS benefits = unearned income","$20 general exclusion applies","Deemed income counted here","In-kind = unearned"],rl:["SI 00810","SI 00835"]},
  {id:"si-9",s:"SI 00835",t:"In-Kind Support & Maintenance",c:"ISM",sh:"SI",r:2,tp:"ISM PMV presumed maximum value VTR third party",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500835000",
    ex:"ISM = food/shelter provided by someone else. Valued using PMV (Presumed Maximum Value) rule or actual value (if less). PMV = ⅓ FBR + $20. Can be rebutted with actual value evidence.",
    rm:["ISM = food/shelter from others","PMV = ⅓ FBR + $20","Can rebut PMV with actual value","Living arrangement affects ISM"],rl:["SI 00510","SI 00810"]},
  {id:"si-10",s:"SI 01110",t:"Resources - General",c:"Resources",sh:"SI",r:3,tp:"resource $2000 $3000 limits countable assets",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0501110000",
    ex:"Resources = things owned convertible to cash. $2K individual / $3K couple. Counted on 1st of month. Excluded: home, vehicle (usually), household goods, burial $1,500, life insurance ≤$1,500.",
    rm:["$2K/$3K on 1st of month","Home excluded","One vehicle usually excluded","Many exclusions — SI 01130"],rl:["SI 01130","SI 01140"]},
  {id:"si-11",s:"SI 01130",t:"Resource Exclusions",c:"Resources",sh:"SI",r:3,tp:"home vehicle burial PASS life insurance excluded resources",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0501130000",
    ex:"Excluded resources: home (regardless of value), one vehicle (with exceptions), household goods/personal effects, burial fund up to $1,500, life insurance ≤$1,500 face value, PASS resources, property for self-support.",
    rm:["Home always excluded","Vehicle rules have exceptions","Burial + life insurance limits","PASS resources excluded"],rl:["SI 01110","SI 01140"]},
  {id:"si-12",s:"SI 01140",t:"Trusts",c:"Resources",sh:"SI",r:3,tp:"trust evaluation special needs SNT countable resource rules",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0501140000",
    ex:"Trust evaluation for SSI resources: revocable trusts = countable resource. Irrevocable trusts: depends on circumstances. Special Needs Trusts (SNTs): certain types excluded (d)(4)(A) self-settled, (d)(4)(C) pooled trusts. POMS SI 01120.200+ covers exceptions.",
    rm:["Revocable = countable","(d)(4)(A) = self-settled SNT excluded","(d)(4)(C) = pooled trust excluded","Complex rules — check SI 01120.200"],rl:["SI 01110","SI 01130"]},
  {id:"si-13",s:"SI 01310",t:"Deeming - General",c:"Deeming",sh:"SI",r:3,tp:"spouse parent child sponsor deeming income resources",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0501310000",
    ex:"Deeming: non-SSI person's income/resources counted for SSI claimant. Types: spouse-to-spouse, parent-to-child (<18), sponsor-to-alien. Allocations for others reduce deemed amount. Stops when relationship ends.",
    rm:["Three types","Allocations reduce","Stops at separation/age 18","Complex — SI 01320 for calcs"],rl:["SI 01320"]},
  {id:"si-14",s:"SI 02302",t:"Section 1619(a) and (b)",c:"Work Incentives",sh:"SI",r:3,tp:"continued SSI Medicaid working 1619 thresholds",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0502302000",
    ex:"1619(a) = SSI continues (reduced) above SGA. 1619(b) = Medicaid continues even at $0 SSI. State-specific thresholds. Removes Medicaid loss fear.",
    rm:["1619(a) = SSI cash continues","1619(b) = Medicaid at $0 SSI","State thresholds for 1619(b)","Critical work incentive"],rl:["SI 00870","SI 00820.545"]},
  {id:"si-15",s:"SI 00870",t:"PASS Plans",c:"Work Incentives",sh:"SI",r:3,tp:"PASS plan self-support excludable resources income",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500870000",
    ex:"Plan to Achieve Self-Support: allows SSI recipients to set aside income/resources for a work goal without affecting SSI eligibility. PASS must have specific occupational goal, be approved by SSA, and be regularly reviewed.",
    rm:["Excludes income AND resources","Must have specific work goal","SSA must approve","Reviewed periodically"],rl:["SI 02302"]},
  {id:"si-16",s:"SI 00820.545",t:"IRWEs for SSI",c:"Work Incentives",sh:"SI",r:3,tp:"IRWE impairment related work expenses SSI deductions",ct:"SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0500820545",
    ex:"IRWEs deducted from earned income for SSI payment calculation. Must be: directly related to impairment, necessary for work, paid by claimant, reasonable cost, not reimbursed. Same concept as DI 10520 but applied to SSI income rules.",
    rm:["Deducted from earned income","Must relate to impairment + work","Not reimbursed","Same concept as DI 10520"],rl:["DI 10520","SI 00820"]},

  // ── GN CROSS-REFS ──
  {id:"gn-1",s:"GN 00201",t:"Applications - General",c:"GN Cross-Ref",sh:"DI",r:2,tp:"applications filing date protective filing Title II XVI",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0200201000",
    ex:"Application rules: filing date vs protective filing date, which applications to use, deemed filing (filing for one title may be deemed filing for both). Protective filing preserves earliest possible entitlement date.",
    rm:["Filing date matters for benefits","Protective filing = earliest date","Deemed filing for concurrent","FO handles applications"],rl:[]},
  {id:"gn-2",s:"GN 00301",t:"Evidence Requirements",c:"GN Cross-Ref",sh:"DI",r:2,tp:"evidence standards proof identity age citizenship documentation",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0200301000",
    ex:"General evidence requirements for all claims: proof of identity, age, citizenship/alien status, SSN, and other non-disability factors. Standards for what constitutes acceptable evidence.",
    rm:["Non-disability evidence rules","Identity + age + citizenship","FO primarily handles","Standards for acceptability"],rl:[]},
  {id:"gn-3",s:"GN 03101",t:"Authorized Representatives",c:"GN Cross-Ref",sh:"DI",r:2,tp:"authorized representative appointed rep claimant advocate attorney",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0203101000",
    ex:"Rules for appointed representatives (attorneys and non-attorneys). Must have SSA Form SSA-1696. Rep can act on claimant's behalf, receive notices, submit evidence. DDS must recognize and communicate with appointed reps.",
    rm:["SSA-1696 = appointment form","Rep acts on claimant's behalf","Must send copies of notices to rep","Can be attorney or non-attorney"],rl:[]},

  // ── BODY SYSTEM LISTING ENTRIES ──
  {id:"di-ls2",s:"Listing 2.00",t:"Special Senses & Speech Listings",c:"Listings",sh:"DI",r:2,tp:"vision hearing speech blindness acuity visual field audiometry cochlear implant labyrinthine vestibular",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-2",
    ex:"Body system 2.00 covers vision loss (central acuity, visual field contraction, visual efficiency), hearing loss (with and without cochlear implant), vestibular/balance disorders, and loss of speech. Vision uses BEST corrected acuity. Hearing requires audiometry meeting SSA standards. Cochlear implant recipients are considered disabled for 1 year post-implant, then re-evaluated. Statutory blindness (20/200 or less in better eye) has special SGA and benefit rules.",
    rm:["Vision = best CORRECTED acuity","20/200 or less = statutory blindness","Cochlear implant = 1 year automatic","Audiometry must meet SSA standards","Statutory blindness has different SGA threshold"],rl:["DI 34001","DI 10501.015"]},
  {id:"di-ls9",s:"Listing 9.00",t:"Endocrine Disorders Listings",c:"Listings",sh:"DI",r:2,tp:"endocrine diabetes thyroid adrenal pituitary hormonal no specific listing cross-reference affected system",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-9",
    ex:"CRITICAL: There are NO specific endocrine listings. Endocrine disorders are evaluated under the body system most affected. Diabetes → evaluate retinopathy (2.00), cardiovascular (4.00), nephropathy (6.00), neuropathy (11.00), Charcot arthropathy (1.00). Thyroid → cognitive/fatigue (12.00), cardiovascular (4.00), cancer (13.00). Adrenal → affected system. This is a common adjudicator error — looking for an endocrine listing that doesn't exist.",
    rm:["NO specific endocrine listings exist","Evaluate under AFFECTED body system","Diabetes complications → multiple possible listings","Common error: looking for endocrine listing"],rl:["DI 34001"],al:"No endocrine listings exist — evaluate under affected body system"},
  {id:"di-ls5",s:"Listing 5.00",t:"Digestive System Listings",c:"Listings",sh:"DI",r:2,tp:"digestive GI liver IBD Crohn ulcerative colitis hemorrhage transplant weight loss BMI bowel",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-5",
    ex:"Covers GI hemorrhaging, chronic liver disease, IBD (Crohn's/UC), short bowel syndrome, weight loss (BMI < 17.50), and liver transplant (1 year automatic). Most require specific lab values, hospitalization records, or objective clinical findings.",
    rm:["GI hemorrhage: need transfusion records","IBD: obstruction or 2+ hospitalizations","BMI < 17.50 on 2+ occasions 60 days apart","Liver transplant = 1 year automatic"],rl:["DI 34001"]},
  {id:"di-ls6",s:"Listing 6.00",t:"Genitourinary Listings",c:"Listings",sh:"DI",r:2,tp:"kidney CKD dialysis hemodialysis transplant nephrotic nephropathy genitourinary renal",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-6",
    ex:"Covers nephrotic syndrome, CKD on chronic dialysis, kidney transplant (1 year automatic), and nephrogenic systemic fibrosis. Dialysis and transplant listings are among the most straightforward — verify dialysis is chronic, not temporary.",
    rm:["Dialysis must be chronic, not temporary","Kidney transplant = 1 year automatic","Nephrotic syndrome: 3+ months persistent","Diabetes nephropathy often evaluated here"],rl:["DI 34001"]},
  {id:"di-ls7",s:"Listing 7.00",t:"Hematological Disorder Listings",c:"Listings",sh:"DI",r:2,tp:"sickle cell anemia hemolytic thrombosis aplastic bone marrow stem cell transplant hematological blood",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-7",
    ex:"Covers sickle cell disease (painful crises 3+ in 5 months), hemolytic anemias (Hgb ≤ 7.0), thrombosis/hemostasis disorders, aplastic anemia/MDS, and bone marrow/stem cell transplant (12 months automatic). Sickle cell requires detailed hospital documentation.",
    rm:["Sickle cell: need dates, duration, treatment records","Hgb ≤ 7.0 on 2+ evals 3+ months apart","Bone marrow transplant = 12 months","Aplastic anemia: specific blood count criteria"],rl:["DI 34001"]},
  {id:"di-ls8",s:"Listing 8.00",t:"Skin Disorder Listings",c:"Listings",sh:"DI",r:2,tp:"skin dermatitis ichthyosis bullous hidradenitis photosensitivity chronic infection lesions",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-8",
    ex:"Covers ichthyosis, bullous disease, chronic skin infections, dermatitis, hidradenitis suppurativa, and genetic photosensitivity. Most require 3+ months duration despite prescribed treatment. 'Extensive' means multiple body sites or critical areas. Burns are evaluated under 1.00 Musculoskeletal.",
    rm:["3+ months despite treatment required","'Extensive' = multiple sites or critical areas","Burns → evaluate under 1.00, not 8.00","Photosensitivity: must show inability to function outside"],rl:["DI 34001"],al:"Burns evaluated under 1.00, NOT 8.00"},
  {id:"di-ls10",s:"Listing 10.00",t:"Congenital Disorder Listings (Adult)",c:"Listings",sh:"DI",r:2,tp:"congenital Down syndrome karyotype birth defect genetic chromosomal",ct:"DI/DIB/SSI",url:"https://secure.ssa.gov/poms.nsf/lnx/0434001000",fl:"listing-10",
    ex:"ONLY non-mosaic Down syndrome (10.01) has a specific adult listing — confirmed by karyotype or molecular testing. All other congenital disorders are evaluated under the affected body system (heart defects → 4.06, spina bifida → 11.00, cleft palate → 2.00). Part B (children) has additional congenital listings including low birth weight.",
    rm:["Only Down syndrome has specific adult listing","All other congenital → affected body system","Must be non-mosaic (confirmed by testing)","Children: Part B has additional listings"],rl:["DI 34001","DI 34002"]},
];
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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };

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
  if (!email) return { success: false, authError: true, message: 'Authentication required.' };

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
            if (Object.prototype.hasOwnProperty.call(authMethods, method)) authMethods[method]++;
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
            var rowMeetingId = String(col_(data[i], MEETING_CHECKIN_COLS.MEETING_ID) || '');
            var rowEmail = String(col_(data[i], MEETING_CHECKIN_COLS.EMAIL) || '').toLowerCase().trim();
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
    var rowEmail = String(col_(memberData[i], MEMBER_COLS.EMAIL) || '').trim().toLowerCase();
    if (rowEmail === emailLower) {
      memberId = String(col_(memberData[i], MEMBER_COLS.MEMBER_ID) || '');
      memberName = String(col_(memberData[i], MEMBER_COLS.FIRST_NAME) || '') + ' ' +
                   String(col_(memberData[i], MEMBER_COLS.LAST_NAME) || '');
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
      var rowMeetingId = String(col_(checkInData[j], MEETING_CHECKIN_COLS.MEETING_ID) || '');
      var rowMemberId = String(col_(checkInData[j], MEETING_CHECKIN_COLS.MEMBER_ID) || '');
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
      if (String(col_(checkInData[k], MEETING_CHECKIN_COLS.MEETING_ID) || '') === meetingId) {
        meetingName = col_(checkInData[k], MEETING_CHECKIN_COLS.MEETING_NAME) || '';
        meetingDate = col_(checkInData[k], MEETING_CHECKIN_COLS.MEETING_DATE) || '';
        meetingType = col_(checkInData[k], MEETING_CHECKIN_COLS.MEETING_TYPE) || '';
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
      if (String(col_(checkInData[m], MEETING_CHECKIN_COLS.MEETING_ID) || '') === meetingId) {
        var currentStatus = String(col_(checkInData[m], MEETING_CHECKIN_COLS.EVENT_STATUS) || '');
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
  if (!leader) return { success: false, authError: true, message: 'Leader access required.' };
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
  if (!leader || !leader.unit) return { success: false, authError: true, message: 'Leader access required.' };
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
  if (!leader) return { success: false, authError: true, sentCount: 0, message: 'Leader access required.' };
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
  if (!leader) return { success: false, authError: true, message: 'Leader access required.' };
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
  if (!leader) return { success: false, authError: true, message: 'Leader access required.' };
  return DataService.getStewardContactLog(leader.email);
}

/**
 * Returns the leader's mentor steward info (from mentorship pairings).
 * @param {string} sessionToken
 * @returns {Object|null} { mentorEmail, mentorName, started, notes } or null
 */
function dataGetLeaderMentor(sessionToken) {
  var leader = _requireLeaderAuth(sessionToken);
  if (!leader) return { success: false, authError: true, message: 'Leader access required.' };
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
