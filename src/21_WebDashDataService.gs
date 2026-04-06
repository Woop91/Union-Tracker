/**
 * DataService.gs
 * Data access layer for the web app dashboard.
 *
 * Reads from:
 *   - Member Directory sheet
 *   - Grievance Log sheet
 *
 * All column positions are resolved dynamically by header name.
 * NEVER hardcode column indices.
 *
 * Access control:
 *   - Stewards see all cases assigned to them
 *   - Members see only their own grievances
 *   - No member PII is exposed to other members
 */

// Tab name used inside per-member contact spreadsheets (NOT the hidden _Contact_Log sheet)
var CONTACT_SHEET_TAB_ = 'Contact Log';

var DataService = (function () {

  // Sheet names — read from SHEETS constants if available, fallback to defaults
  var MEMBER_SHEET = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) ? SHEETS.MEMBER_DIR : 'Member Directory';
  var GRIEVANCE_SHEET = (typeof SHEETS !== 'undefined' && SHEETS.GRIEVANCE_LOG) ? SHEETS.GRIEVANCE_LOG : 'Grievance Log';

  // PERF: Compute lowercased closed statuses once at module init (avoids 4× redundant .map() calls)
  var _closedStatusesLower = GRIEVANCE_CLOSED_STATUSES.map(function(s) { return s.toLowerCase(); });

  // PERF: Spreadsheet singleton — avoids redundant getActiveSpreadsheet() IPC calls.
  // Cached per execution (GAS module-scope resets between requests).
  var _cachedSS = null;
  /**
   * Returns the cached spreadsheet singleton, initializing on first call.
   * @returns {Spreadsheet}
   */
  function _getSS() {
    if (_cachedSS) return _cachedSS;
    _cachedSS = SpreadsheetApp.getActiveSpreadsheet();
    return _cachedSS;
  }
  /** @private Reset cached spreadsheet reference (used by test harness). */
  function _resetSSCache() { _cachedSS = null; }

  // Header name mappings — these are the expected header labels
  // If your sheet uses different labels, update these OR add aliases
  var HEADERS = {
    // Member Directory
    memberEmail:     ['email', 'email address', 'member email', 'e-mail', 'work email'],
    memberName:      ['name', 'full name', 'member name'],
    memberFirstName: ['first name', 'first'],
    memberLastName:  ['last name', 'last'],
    memberRole:      ['role', 'member role', 'type', 'member type', 'membership type'],
    memberUnit:      ['unit', 'workplace unit', 'department'],
    memberPhone:     ['phone', 'phone number', 'cell', 'mobile'],
    memberJoined:    ['joined', 'join date', 'member since', 'date joined', 'hire date'],
    memberDuesStatus:['dues status', 'dues', 'status'],
    memberDuesPaying:['dues paying', 'is dues paying', 'dues paid', 'paying dues'],
    memberId:        ['member id', 'id', 'member number'],
    memberWorkLocation: ['work location', 'location', 'office location'],
    memberOfficeDays:   ['office days', 'in-office days', 'days in office'],
    memberAssignedSteward: ['assigned steward', 'steward assignment'],
    memberIsSteward: ['is steward', 'steward'],
    memberStreet:    ['street address', 'street', 'address'],
    memberCity:      ['city'],
    memberState:     ['state'],
    memberZip:       ['zip code', 'zip', 'postal code'],
    memberHasOpenGrievance: ['has open grievance?', 'has open grievance', 'open grievance'],
    memberSupervisor: ['supervisor'],
    memberJobTitle:  ['job title', 'title', 'position'],
    memberAdminFolderUrl: ['member admin folder url', 'member admin folder', 'contact log folder url'],  // 'contact log folder url' retained as fallback alias for backward compat
    memberSharePhone:    ['share phone', 'share phone number', 'phone visible', 'public phone', 'share contact'],
    memberCubicle:       ['cubicle', 'cube', 'workstation'],
    memberEmployeeId:    ['employee id', 'employee number', 'emp id', 'emp no'],
    memberHireDate:      ['hire date', 'date hired', 'start date'],
    memberOpenRate:      ['open rate %', 'open rate', 'email open rate'],
    memberShirtSize:     ['shirt size', 'tshirt size', 't-shirt size'],
    memberManager:       ['director', 'manager', 'director (manager)'],

    // Grievance Log
    grievanceId:     ['grievance id', 'id', 'case id', 'gr id'],
    grievanceMemberEmail: ['member email', 'email', 'filed by email', 'grievant email'],
    grievanceMemberFirstName: ['first name', 'first'],
    grievanceMemberLastName:  ['last name', 'last'],
    grievanceStatus: ['status', 'grievance status', 'case status'],
    grievanceStep:   ['step', 'current step', 'grievance step'],
    grievanceDeadline: ['deadline', 'next deadline', 'due date', 'next action due', 'filing deadline'],
    grievanceFiled:  ['filed', 'filed date', 'date filed', 'created'],
    grievanceSteward:['assigned steward', 'steward', 'steward email', 'assigned to'],
    grievanceUnit:   ['unit', 'workplace unit', 'work location', 'location'],
    grievancePriority: ['priority', 'urgency'],
    grievanceNotes:  ['notes', 'description', 'summary', 'resolution'],
    grievanceIssueCategory: ['issue category', 'category', 'issue type'],
    grievanceResolution: ['resolution', 'outcome', 'result'],
    grievanceDateClosed: ['date closed', 'closed date', 'closed', 'resolved date'],

    // v4.32.1 — Drive Folder URL lookup for grievance records.
    // Maps to the "Drive Folder URL" column (GRIEVANCE_COLS.DRIVE_FOLDER_URL, col 33)
    // in the Grievance Log sheet. This column is auto-populated by
    // setupDriveFolderForGrievance() in 05_Integrations.gs when a steward creates
    // a case folder. Multiple header aliases allow fuzzy matching via _findColumn().
    //
    // WHY: _buildGrievanceRecord() needs this to include driveFolderUrl in its
    // return object. Without it, getMemberGrievanceDriveUrl() and
    // dataGetMemberCaseFolderUrl() cannot resolve the folder URL and return null.
    //
    // IF THIS BREAKS: The Grievance Log "Drive Folder URL" column was renamed or
    // removed. Check GRIEVANCE_HEADER_MAP_ in 01_Core.gs (key: DRIVE_FOLDER_URL)
    // and verify the header text matches one of the aliases below. If none match,
    // _getVal() returns '' and driveFolderUrl defaults to null — safe but non-functional.
    grievanceDriveFolderUrl: ['drive folder url', 'drive url', 'folder url', 'case folder url'],
    grievanceDriveFolderId: ['drive folder id', 'folder id', 'case folder id'],
    grievanceArticles: ['articles violated', 'articles', 'contract articles'],
    grievanceStep1Due: ['step i due', 'step 1 due'],
    grievanceStep1Rcvd: ['step i rcvd', 'step 1 rcvd', 'step i received', 'step 1 received'],
    grievanceStep2AppealDue: ['step ii appeal due', 'step 2 appeal due'],
    grievanceStep2Due: ['step ii due', 'step 2 due'],
    grievanceStep2Rcvd: ['step ii rcvd', 'step 2 rcvd', 'step ii received'],
    grievanceSignatureStatus: ['signature status'],
    grievanceSignatureToken: ['signature token'],
    grievanceSignedDate: ['signed date'],
  };

  // ═══════════════════════════════════════
  // PUBLIC: User Lookup
  // ═══════════════════════════════════════

  // ─── Email Index — O(1) lookup replacing O(n) linear scan ───
  var _emailIndex = null; // { email: rowIndex } — built once per execution

  /**
   * Builds and returns the O(1) email-to-row-index lookup map for the Member Directory.
   * @returns {Object} Map of lowercase email to row index
   */
  function _getEmailIndex() {
    if (_emailIndex) return _emailIndex;
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return {};
    var emailCol = _findColumn(cached.colMap, HEADERS.memberEmail);
    if (emailCol === -1) return {};
    _emailIndex = {};
    for (var i = 1; i < cached.data.length; i++) {
      var e = String(cached.data[i][emailCol]).trim().toLowerCase();
      if (e) _emailIndex[e] = i;
    }
    return _emailIndex;
  }

  /**
   * Finds a user in the Member Directory by email using O(1) index lookup.
   * @param {string} email
   * @returns {Object|null} Sanitized user record or null
   */
  function findUserByEmail(email) {
    if (!email) return null;
    email = String(email).trim().toLowerCase();
    var idx = _getEmailIndex();
    var rowIdx = idx[email];
    if (rowIdx === undefined) return null;
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return null;
    return _buildUserRecord(cached.data[rowIdx], cached.colMap);
  }

  /**
   * Returns the role for a given email.
   * @param {string} email
   * @returns {string|null} 'steward', 'member', 'both', or null
   */
  function getUserRole(email) {
    var user = findUserByEmail(email);
    if (!user) return null;
    return user.role;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Data
  // ═══════════════════════════════════════

  /**
   * Returns all grievances assigned to a steward.
   * @param {string} stewardEmail
   * @returns {Object[]} Array of grievance records
   */
  function getStewardCases(stewardEmail) {
    stewardEmail = String(stewardEmail).trim().toLowerCase();

    // Resolve steward's display name for name-based matching
    // (Grievance Log "Assigned Steward" stores names, not emails)
    var stewardName = '';
    var stewardRecord = findUserByEmail(stewardEmail);
    if (stewardRecord) {
      stewardName = String(stewardRecord.name || '').trim().toLowerCase();
    }

    var cached = _getCachedSheetData(GRIEVANCE_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap;

    var stewardCol = _findColumn(colMap, HEADERS.grievanceSteward);
    if (stewardCol === -1) {
      log_('DataService', 'Steward column not found in Grievance Log');
      return [];
    }

    var cases = [];
    for (var i = 1; i < data.length; i++) {
      var assignedTo = String(data[i][stewardCol]).trim().toLowerCase();
      // Dual-match: accept email OR name (sheet stores names; email is fallback)
      var matchesEmail = assignedTo === stewardEmail;
      var matchesName  = stewardName && assignedTo === stewardName;
      if (matchesEmail || matchesName) {
        cases.push(_buildGrievanceRecord(data[i], colMap));
      }
    }

    // Sort: overdue first, then by deadline ascending
    cases.sort(function (a, b) {
      if (a.status === 'overdue' && b.status !== 'overdue') return -1;
      if (b.status === 'overdue' && a.status !== 'overdue') return 1;
      return (a.deadlineDays || 999) - (b.deadlineDays || 999);
    });

    return cases;
  }

  /**
   * Returns KPIs for a steward's personal caseload.
   * @param {string} stewardEmail
   * @returns {Object} { totalCases, overdue, dueSoon, resolved }
   */
  function getStewardKPIs(stewardEmail) {
    var cases = getStewardCases(stewardEmail);
    
    var total = cases.length;
    var overdue = 0;
    var dueSoon = 0;
    var resolved = 0;
    var active = 0;

    for (var i = 0; i < cases.length; i++) {
      var c = cases[i];
      var status = String(c.status).toLowerCase();

      if (_closedStatusesLower.indexOf(status) !== -1) {
        resolved++;
      } else {
        active++;
        if (status === 'overdue') {
          overdue++;
        }
        if (c.deadlineDays !== null && c.deadlineDays <= 7 && c.deadlineDays >= 0) {
          dueSoon++;
        }
      }
    }

    return {
      totalCases: total,
      activeCases: active,
      overdue: overdue,
      dueSoon: dueSoon,
      resolved: resolved,
    };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Member Data
  // ═══════════════════════════════════════

  /**
   * Returns grievances for a specific member (their own only).
   * @param {string} memberEmail
   * @returns {Object[]} Array of grievance records
   */
  function getMemberGrievances(memberEmail) {
    memberEmail = String(memberEmail).trim().toLowerCase();

    var cached = _getCachedSheetData(GRIEVANCE_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap;

    var memberCol = _findColumn(colMap, HEADERS.grievanceMemberEmail);
    if (memberCol === -1) return [];

    var closedStatuses = _closedStatusesLower;
    var grievances = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][memberCol]).trim().toLowerCase();
      if (rowEmail !== memberEmail) continue;
      var rowStatus = String(_getVal(data[i], colMap, HEADERS.grievanceStatus, '')).trim().toLowerCase();
      if (closedStatuses.indexOf(rowStatus) !== -1) continue;
      grievances.push(_buildGrievanceRecord(data[i], colMap));
    }

    // Sort by filed date, most recent first
    grievances.sort(function (a, b) {
      return (b.filedTimestamp || 0) - (a.filedTimestamp || 0);
    });

    return grievances;
  }

  /**
   * Returns resolved/closed grievance history for a member.
   * Only includes cases with terminal statuses. Excludes internal notes,
   * steward names, and other sensitive fields for member privacy.
   * @param {string} memberEmail
   * @returns {Object} { success: true, history: [...] }
   */
  function getMemberGrievanceHistory(memberEmail) {
    if (!memberEmail) return { success: true, history: [] };
    memberEmail = String(memberEmail).trim().toLowerCase();

    // Use combined active + archive for complete member history
    var cached = _getAllGrievanceData();
    if (!cached) return { success: true, history: [] };

    var data = cached.data;
    var colMap = cached.colMap;

    var memberCol = _findColumn(colMap, HEADERS.grievanceMemberEmail);
    if (memberCol === -1) return { success: true, history: [] };

    var closedStatuses = _closedStatusesLower;
    var history = [];

    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][memberCol]).trim().toLowerCase();
      if (rowEmail !== memberEmail) continue;

      var status = String(_getVal(data[i], colMap, HEADERS.grievanceStatus, '')).trim().toLowerCase();
      if (closedStatuses.indexOf(status) === -1) continue;

      var filedRaw = _getVal(data[i], colMap, HEADERS.grievanceFiled, null);
      var filedFormatted = '';
      var filedTimestamp = 0;
      if (filedRaw instanceof Date) {
        filedFormatted = _formatDate(filedRaw);
        filedTimestamp = filedRaw.getTime();
      } else if (filedRaw) {
        var parsed = new Date(filedRaw);
        if (!isNaN(parsed.getTime())) {
          filedFormatted = _formatDate(parsed);
          filedTimestamp = parsed.getTime();
        }
      }

      var closedRaw = _getVal(data[i], colMap, HEADERS.grievanceDateClosed, null);
      var closedFormatted = '';
      if (closedRaw instanceof Date) {
        closedFormatted = _formatDate(closedRaw);
      } else if (closedRaw) {
        var parsedClosed = new Date(closedRaw);
        if (!isNaN(parsedClosed.getTime())) {
          closedFormatted = _formatDate(parsedClosed);
        }
      }

      history.push({
        grievanceId: String(_getVal(data[i], colMap, HEADERS.grievanceId, '')).trim(),
        issueCategory: String(_getVal(data[i], colMap, HEADERS.grievanceIssueCategory, '')).trim(),
        status: status.charAt(0).toUpperCase() + status.slice(1),
        outcome: String(_getVal(data[i], colMap, HEADERS.grievanceResolution, '')).trim(),
        dateFiled: filedFormatted,
        dateClosed: closedFormatted,
        filedTimestamp: filedTimestamp,
      });
    }

    // Sort by filed date descending (most recent first)
    history.sort(function(a, b) {
      return (b.filedTimestamp || 0) - (a.filedTimestamp || 0);
    });

    // Strip timestamp from returned data
    history.forEach(function(h) { delete h.filedTimestamp; });

    return { success: true, history: history };
  }

  /**
   * Returns steward contact info for a member.
   * Only returns email (always) and phone (if listed).
   * @param {string} stewardEmail
   * @returns {Object|null} { name, email, phone }
   */
  function getStewardContact(stewardEmail) {
    if (!stewardEmail) return null;
    var user = findUserByEmail(stewardEmail);
    if (!user) return null;

    return {
      name: user.name || user.firstName + ' ' + user.lastName,
      email: user.email,
      phone: user.phone || null, // Only if listed
    };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Full Profile & Profile Updates
  // ═══════════════════════════════════════

  /**
   * Returns full member profile for self-service profile page.
   * Only returns data for the authenticated user (email match).
   * @param {string} email
   * @returns {Object|null}
   */
  function getFullMemberProfile(email) {
    var user = findUserByEmail(email);
    if (!user) return null;
    return {
      email: user.email,
      name: user.name,
      firstName: user.firstName,
      lastName: user.lastName,
      unit: user.unit,
      workLocation: user.workLocation,
      officeDays: user.officeDays,
      street: user.street,
      city: user.city,
      state: user.state,
      zip: user.zip,
      phone: user.phone,
      supervisor: user.supervisor,
      director: user.director,
      jobTitle: user.jobTitle,
      joined: user.joined,
      memberId: user.memberId || '',
      memberAdminFolderUrl: user.memberAdminFolderUrl || '',
      cubicle: user.cubicle || '',
      employeeId: user.employeeId || '',
      hireDate: user.hireDate || '',
      openRate: user.openRate || '',
      shirtSize: user.shirtSize || '',
    };
  }

  /**
   * Updates a member's self-service profile fields.
   * Only allows editing specific fields: address, work location, office days.
   * @param {string} email - The member's email (verified against session)
   * @param {Object} updates - { street, city, state, zip, workLocation, officeDays }
   * @returns {Object} { success: boolean, message: string }
   */
  function updateMemberProfile(email, updates) {
    if (!email || !updates) return { success: false, message: 'Invalid request.' };
    email = String(email).trim().toLowerCase();

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return { success: false, message: 'Member directory unavailable.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    if (emailCol === -1) return { success: false, message: 'System configuration error.' };

    // Editable field mappings
    var editableFields = {
      phone:        HEADERS.memberPhone,        // v4.51.1: member can update phone (BUG-3-001)
      street:       HEADERS.memberStreet,
      city:         HEADERS.memberCity,
      state:        HEADERS.memberState,
      zip:          HEADERS.memberZip,
      workLocation: HEADERS.memberWorkLocation,
      officeDays:   HEADERS.memberOfficeDays,
      sharePhone:   HEADERS.memberSharePhone,   // steward opt-in: phone visible to members
      shirtSize:    HEADERS.memberShirtSize,    // member shirt size (self-service)
    };

    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][emailCol]).trim().toLowerCase();
      if (rowEmail !== email) continue;

      var rowNum = i + 1; // 1-indexed for sheet ops
      var rowData = data[i].slice(); // copy row; apply all field edits in-memory
      for (var field in updates) {
        if (!editableFields[field]) continue;
        var col = _findColumn(colMap, editableFields[field]);
        if (col === -1) continue;
        var val = String(updates[field] || '').trim().substring(0, 255);
        rowData[col] = escapeForFormula(val);
      }
      // Write entire row in one API call instead of one setValue per field
      sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);

      _invalidateSheetCache(MEMBER_SHEET);
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('PROFILE_UPDATE', { email: email, fields: Object.keys(updates).join(',') });
      }

      // Sync shirt size to the dedicated log sheet when updated
      if (updates.shirtSize !== undefined) {
        try {
          _syncShirtSizeLog(email, updates.shirtSize);
        } catch (logErr) {
          log_('Shirt size log sync error', logErr.message);
        }
      }

      return { success: true, message: 'Profile updated.' };
    }

    return { success: false, message: 'Member not found.' };
  }

  /**
   * Steward-level member field update. Supports a broader set of fields than
   * the self-service updateMemberProfile (which is limited to address/location).
   * Used by the Directory Explorer inline editor.
   * @param {string} email - Member email to update
   * @param {Object} updates - { director, supervisor, unit, workLocation, jobTitle, officeDays }
   * @returns {{ success: boolean, message: string }}
   */
  function updateMemberBySteward(email, updates) {
    if (!email || !updates) return { success: false, message: 'Invalid request.' };
    email = String(email).trim().toLowerCase();

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return { success: false, message: 'Member directory unavailable.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    if (emailCol === -1) return { success: false, message: 'System configuration error.' };

    var editableFields = {
      director:     HEADERS.memberManager,
      supervisor:   HEADERS.memberSupervisor,
      unit:         HEADERS.memberUnit,
      workLocation: HEADERS.memberWorkLocation,
      jobTitle:     HEADERS.memberJobTitle,
      officeDays:   HEADERS.memberOfficeDays,
    };

    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][emailCol]).trim().toLowerCase();
      if (rowEmail !== email) continue;

      var rowNum = i + 1;
      var oldValues = {};
      for (var _f in updates) {
        if (!editableFields[_f]) continue;
        var _oldCol = _findColumn(colMap, editableFields[_f]);
        if (_oldCol !== -1) oldValues[_f] = data[i][_oldCol];
      }
      var changedCols = []; // {col: number, val: string}
      for (var field in updates) {
        if (!editableFields[field]) continue;
        var col = _findColumn(colMap, editableFields[field]);
        if (col === -1) continue;
        var val = String(updates[field] || '').trim().substring(0, 255);
        changedCols.push({ col: col, val: escapeForFormula(val) });
      }
      if (changedCols.length === 0) return { success: false, message: 'No valid fields to update.' };

      // Write only changed columns to preserve formulas in untouched cells (GAMMA-05)
      // Values already sanitized via escapeForFormula above
      for (var ci = 0; ci < changedCols.length; ci++) {
        sheet.getRange(rowNum, changedCols[ci].col + 1).setValue(changedCols[ci].val); // escapeForFormula applied in changedCols.push
      }
      _invalidateSheetCache(MEMBER_SHEET);
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('STEWARD_MEMBER_UPDATE', {
          email: email,
          changes: Object.keys(updates).reduce(function(acc, f) {
            if (editableFields[f]) {
              acc[f] = { from: String(oldValues[f] || ''), to: String(updates[f] || '').trim() };
            }
            return acc;
          }, {})
        });
      }
      return { success: true, message: 'Member updated.' };
    }

    return { success: false, message: 'Member not found.' };
  }

  /**
   * Syncs a member's shirt size to the Shirt Size Log sheet.
   * Creates the sheet with headers if it doesn't exist.
   * Upserts: updates existing row for the email, or appends a new one.
   * @param {string} email
   * @param {string} shirtSize
   * @private
   */
  function _syncShirtSizeLog(email, shirtSize) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheetName = SHEETS.SHIRT_SIZE_LOG;
    var logSheet = ss.getSheetByName(sheetName);
    var LOG_HEADERS = ['Name', 'Email', 'Shirt Size', 'Work Location', 'Unit', 'Assigned Steward', 'Updated'];

    if (!logSheet) {
      logSheet = ss.insertSheet(sheetName);
      logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
      logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    // Look up member details from directory
    var user = findUserByEmail(email);
    var memberName = user ? user.name : email;
    var workLocation = user ? (user.workLocation || '') : '';
    var unit = user ? (user.unit || '') : '';
    var steward = user ? (user.assignedSteward || '') : '';

    // Resolve steward name from email if possible
    if (steward && steward.indexOf('@') !== -1) {
      var stewardUser = findUserByEmail(steward);
      if (stewardUser && stewardUser.name) steward = stewardUser.name;
    }

    var newRow = [memberName, email, shirtSize || '', workLocation, unit, steward, new Date()];

    // Upsert: find existing row by email (col 2)
    var data = logSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === email) {
        logSheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        return;
      }
    }
    // Append new row
    logSheet.appendRow(newRow);
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Assignment
  // ═══════════════════════════════════════

  /**
   * Returns the assigned steward info for a member.
   * @param {string} email
   * @returns {Object|null} { name, email, phone, workLocation, officeDays }
   */
  function getAssignedStewardInfo(email) {
    var user = findUserByEmail(email);
    if (!user || !user.assignedSteward) return null;

    // Try email lookup first
    var steward = findUserByEmail(user.assignedSteward);

    // Fallback: if assignedSteward is a name (no '@'), scan by name
    if (!steward && user.assignedSteward.indexOf('@') === -1) {
      steward = _findUserByName(user.assignedSteward);
    }

    if (!steward) return null;
    return {
      name: steward.name,
      email: steward.email,
      phone: steward.phone,
      workLocation: steward.workLocation,
      officeDays: steward.officeDays,
    };
  }

  /**
   * Returns stewards at the same location as the member.
   * Used for self-assign steward picker.
   * @param {string} memberEmail
   * @returns {Object[]} Array of { name, email, workLocation, officeDays }
   */
  function getAvailableStewards(memberEmail) {
    var member = findUserByEmail(memberEmail);
    if (!member) return [];

    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap;
    var stewards = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      if (!rec.isSteward) continue;
      if (rec.email === memberEmail) continue;
      // Prioritize same location, but include all stewards
      stewards.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        sameLocation: rec.workLocation && member.workLocation &&
          rec.workLocation.toLowerCase() === member.workLocation.toLowerCase(),
      });
    }

    // Sort: same location first, then alphabetical
    stewards.sort(function(a, b) {
      if (a.sameLocation && !b.sameLocation) return -1;
      if (!a.sameLocation && b.sameLocation) return 1;
      return (a.name || '').localeCompare(b.name || '');
    });

    return stewards;
  }

  /**
   * Assigns a steward to a member (self-service).
   * @param {string} memberEmail
   * @param {string} stewardEmail
   * @returns {Object} { success, message }
   */
  function assignStewardToMember(memberEmail, stewardEmail) {
    if (!memberEmail || !stewardEmail) return { success: false, message: 'Invalid request.' };
    memberEmail = String(memberEmail).trim().toLowerCase();
    stewardEmail = String(stewardEmail).trim().toLowerCase();

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return { success: false, message: 'System error.' };

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
    if (emailCol === -1 || stewardCol === -1) return { success: false, message: 'Configuration error.' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).trim().toLowerCase() === memberEmail) {
        sheet.getRange(i + 1, stewardCol + 1).setValue(stewardEmail);
        _invalidateSheetCache(MEMBER_SHEET);
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('STEWARD_SELF_ASSIGN', { member: memberEmail, steward: stewardEmail });
        }
        return { success: true, message: 'Steward assigned.' };
      }
    }
    return { success: false, message: 'Member not found.' };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Grievance Drive & Survey Status
  // ═══════════════════════════════════════

  /**
   * Returns the Google Drive folder URL for a member's grievance case files.
   *
   * WHAT: Looks up the member's grievances and returns the Drive folder URL
   * from the most relevant case — preferring active (non-terminal) cases,
   * falling back to the most recent case if all are closed.
   *
   * WHY THIS APPROACH: Rather than querying the Grievance Log sheet directly
   * for the DRIVE_FOLDER_URL column, this reuses getMemberGrievances() which
   * already handles caching, column resolution, and access filtering. The
   * driveFolderUrl field was added to _buildGrievanceRecord() in v4.32.1
   * specifically to enable this lookup.
   *
   * DATA FLOW:
   *   Grievance Log sheet → _getCachedSheetData() → _buildGrievanceRecord()
   *   → grievanceRecord.driveFolderUrl → returned here
   *
   * CALLERS:
   *   - DataService.getMemberGrievanceDriveUrl (exposed in DataService public API, line ~3387)
   *   - dataGetMemberCaseFolderUrl() (steward-only wrapper, line ~3608)
   *
   * IF THIS BREAKS:
   *   Returns null. The Drive folder link in the UI will show
   *   "No Drive folder linked to this case." This is safe — the folder
   *   still exists in Drive, it just can't be resolved from the sheet.
   *   Root causes: HEADERS.grievanceDriveFolderUrl aliases don't match
   *   the actual column header, or the column was removed from the sheet.
   *
   * @param {string} email — member email address (case-insensitive)
   * @returns {string|null} Full Google Drive folder URL, or null if no folder exists
   */
  function getMemberGrievanceDriveUrl(email) {
    var grievances = getMemberGrievances(email);
    if (!grievances || grievances.length === 0) return null;

    // Prefer an active grievance (one that hasn't reached a terminal status).
    // Terminal statuses mirror GRIEVANCE_CLOSED_STATUSES used elsewhere.
    // If ALL grievances are terminal, fall back to the first (most recent by filing date,
    // since getMemberGrievances() sorts by filedTimestamp descending).
    var TERMINAL = ['resolved', 'closed', 'withdrawn', 'denied', 'won', 'settled'];
    var activeG = grievances.find(function(g) {
      return TERMINAL.indexOf((g.status || '').toLowerCase()) === -1;
    });
    var target = activeG || grievances[0];

    // Return the Drive Folder URL column value if populated
    if (target.driveFolderUrl) return target.driveFolderUrl;

    // Fallback: build URL from Drive Folder ID column if present
    if (target.driveFolderId) {
      return 'https://drive.google.com/drive/folders/' + target.driveFolderId;
    }

    return null;
  }

  /**
   * Creates a grievance draft for a member (member-initiated).
   * Writes to the Grievance Log with status 'Draft' so steward can review.
   * @param {string} email — member email (server-resolved)
   * @param {Object} data — { title, category, description }
   * @returns {Object} { success: boolean, message: string }
   */
  function startGrievanceDraft(email, data, idemKey) {
    if (idemKey) {
      var idemCache = CacheService.getScriptCache();
      if (idemCache.get('IDEM_' + idemKey)) return { duplicate: true, message: 'Duplicate request ignored' };
      idemCache.put('IDEM_' + idemKey, '1', 600); // 10-minute TTL to reduce duplicate mutation risk
    }
    if (!email || !data) return { success: false, message: 'Missing required data.' };
    if (!data.title || !data.description) return { success: false, message: 'Title and description are required.' };

    try {
      var ss = _getSS();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet) return { success: false, message: 'Grievance sheet not found.' };

      // Resolve member identity dynamically from directory
      var memberId = '';
      var memberFirstName = '';
      var memberLastName = '';
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberDir && memberDir.getLastRow() > 1) {
        var mData = memberDir.getDataRange().getValues();
        var mColMap = _buildColumnMap(mData[0]);
        var mEmailCol = _findColumn(mColMap, HEADERS.memberEmail);
        var mIdCol    = _findColumn(mColMap, HEADERS.memberId);
        var mFirstCol = _findColumn(mColMap, HEADERS.memberFirstName);
        var mLastCol  = _findColumn(mColMap, HEADERS.memberLastName);
        var emailLower = email.toLowerCase().trim();
        for (var i = 1; i < mData.length; i++) {
          if (mEmailCol === -1) break;
          if (String(mData[i][mEmailCol] || '').toLowerCase().trim() === emailLower) {
            memberId       = mIdCol    !== -1 ? String(mData[i][mIdCol]    || '').trim() : '';
            memberFirstName = mFirstCol !== -1 ? String(mData[i][mFirstCol] || '').trim() : '';
            memberLastName  = mLastCol  !== -1 ? String(mData[i][mLastCol]  || '').trim() : '';
            break;
          }
        }
      }

      // Build draft row using dynamic column resolution
      var gData    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
      var gColMap  = _buildColumnMap(gData[0]);
      var totalCols = sheet.getLastColumn();
      var row = new Array(totalCols).fill('');

      var colGrievanceId = _findColumn(gColMap, HEADERS.grievanceId);
      var colMemberId    = _findColumn(gColMap, ['member id']);
      var colFirstName   = _findColumn(gColMap, HEADERS.grievanceMemberFirstName);
      var colLastName    = _findColumn(gColMap, HEADERS.grievanceMemberLastName);
      var colStatus      = _findColumn(gColMap, HEADERS.grievanceStatus);
      var colFiled       = _findColumn(gColMap, HEADERS.grievanceFiled);
      var colUpdated     = _findColumn(gColMap, ['last updated']);
      var colCategory    = _findColumn(gColMap, HEADERS.grievanceIssueCategory);
      var colResolution  = _findColumn(gColMap, HEADERS.grievanceResolution);
      var colEmail       = _findColumn(gColMap, HEADERS.grievanceMemberEmail);

      if (colGrievanceId !== -1) row[colGrievanceId] = 'DRAFT-' + Utilities.getUuid().substring(0, 8);
      if (colMemberId    !== -1) row[colMemberId]    = escapeForFormula(memberId);
      if (colFirstName   !== -1) row[colFirstName]   = escapeForFormula(memberFirstName);
      if (colLastName    !== -1) row[colLastName]    = escapeForFormula(memberLastName);
      if (colStatus      !== -1) row[colStatus]      = 'Draft';
      if (colFiled       !== -1) row[colFiled]       = new Date();
      if (colUpdated     !== -1) row[colUpdated]     = new Date();
      if (colCategory    !== -1) row[colCategory]    = escapeForFormula(data.category || '');
      if (colResolution  !== -1) row[colResolution]  = escapeForFormula('[Draft] ' + data.title + ': ' + data.description);
      if (colEmail       !== -1) row[colEmail]       = escapeForFormula(email.toLowerCase().trim());

      sheet.appendRow(row);
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('GRIEVANCE_DRAFT_CREATED', { email: email, grievanceId: row[colGrievanceId] || 'unknown' });
      }
      return { success: true, message: 'Draft submitted.' };
    } catch (e) {
      handleError(e, 'DataService.startGrievanceDraft', ERROR_LEVEL.ERROR);
      return { success: false, message: 'Error creating draft.' };
    }
  }

  /**
   * Creates a Drive folder for the member's most recent active grievance.
   * Delegates to the global setupDriveFolderForGrievance() in 05_Integrations.gs.
   * @param {string} email - Member email
   * @returns {Object} { success, folderUrl?, message }
   */
  function createGrievanceDriveFolder(email) {
    if (!email) return { success: false, message: 'Email is required.' };
    try {
      var ss = _getSS();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'No grievances found.' };

      var emailLower = String(email).toLowerCase().trim();

      // ── Resolve memberId dynamically from Member Directory ────────────────
      var memberId = '';
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberDir && memberDir.getLastRow() > 1) {
        var mData   = memberDir.getDataRange().getValues();
        var mColMap = _buildColumnMap(mData[0]);
        var mEmailCol = _findColumn(mColMap, HEADERS.memberEmail);
        var mIdCol    = _findColumn(mColMap, HEADERS.memberId);
        if (mEmailCol !== -1) {
          for (var m = 1; m < mData.length; m++) {
            if (String(mData[m][mEmailCol] || '').toLowerCase().trim() === emailLower) {
              memberId = mIdCol !== -1 ? String(mData[m][mIdCol] || '').trim() : '';
              break;
            }
          }
        }
      }

      // ── Find most recent grievance via member email or memberId ───────────
      // Prefer direct email match (more reliable); fall back to memberId match.
      var gData   = sheet.getDataRange().getValues();
      var gColMap = _buildColumnMap(gData[0]);
      var gEmailCol      = _findColumn(gColMap, HEADERS.grievanceMemberEmail);
      var gIdCol         = _findColumn(gColMap, HEADERS.grievanceId);
      var gMemberIdCol   = _findColumn(gColMap, ['member id']);

      var grievanceId = '';
      for (var i = gData.length - 1; i >= 1; i--) {
        var rowEmail    = gEmailCol    !== -1 ? String(gData[i][gEmailCol]    || '').toLowerCase().trim() : '';
        var rowMemberId = gMemberIdCol !== -1 ? String(gData[i][gMemberIdCol] || '').trim()              : '';
        var emailMatch    = rowEmail && rowEmail === emailLower;
        var memberIdMatch = memberId && rowMemberId === memberId;
        if (emailMatch || memberIdMatch) {
          grievanceId = gIdCol !== -1 ? String(gData[i][gIdCol] || '').trim() : '';
          if (grievanceId) break;
        }
      }
      if (!grievanceId) return { success: false, message: 'No grievances found for this member.' };

      // Delegate to the global Drive folder setup function in 05_Integrations.gs
      return setupDriveFolderForGrievance(grievanceId);
    } catch (e) {
      log_('DataService.createGrievanceDriveFolder error', e.message);
      return { success: false, message: 'Error creating Drive folder.' };
    }
  }

  /**
   * Returns survey completion status for a member.
   * @param {string} email
   * @returns {Object} { hasCompleted: boolean, lastCompleted: string|null }
   */
  function getMemberSurveyStatus(email) {
    if (!email) return { hasCompleted: false, lastCompleted: null };
    email = String(email).trim().toLowerCase();

    try {
      var ss = _getSS();
      if (!ss) return { hasCompleted: false, lastCompleted: null };
      var trackSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
      if (!trackSheet || trackSheet.getLastRow() <= 1) {
        return { hasCompleted: false, lastCompleted: null };
      }

      var data = trackSheet.getDataRange().getValues();
      var headers = data[0];
      var colMap = _buildColumnMap(headers);
      var emailCol = _findColumn(colMap, ['email', 'member email']);
      var statusCol = _findColumn(colMap, ['status', 'completion status', 'completed']);
      var dateCol = _findColumn(colMap, ['completed date', 'date completed', 'completion date']);

      if (emailCol === -1) return { hasCompleted: false, lastCompleted: null };

      for (var i = 1; i < data.length; i++) {
        if (String(data[i][emailCol]).trim().toLowerCase() !== email) continue;
        var status = statusCol !== -1 ? String(data[i][statusCol]).trim().toLowerCase() : '';
        var completed = status === 'completed' || status === 'yes' || status === 'true';
        var dateVal = dateCol !== -1 ? data[i][dateCol] : null;
        return {
          hasCompleted: completed,
          lastCompleted: dateVal instanceof Date ? _formatDate(dateVal) : (dateVal ? String(dateVal) : null),
        };
      }
    } catch (e) {
      log_('DataService.getMemberSurveyStatus error', e.message);
    }
    return { hasCompleted: false, lastCompleted: null };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Resource Click Tracking (v4.31.1)
  // ═══════════════════════════════════════

  /**
   * Logs a resource click/view event. Uses CacheService for lightweight counting.
   * Each resource gets its own cache key:
   *   RC_<resourceId> = cumulative click count
   *   RC_TOTAL        = grand total across all resources
   *
   * Design: CacheService (max 6h TTL, auto-eviction) is used instead of
   * ScriptProperties to avoid unbounded key growth that pressures the 500KB
   * PropertiesService quota. Click analytics are non-critical — natural cache
   * expiration is acceptable; counts reset periodically, which is fine for
   * relative popularity ranking.
   *
   * @param {string} email  — caller email (server-resolved)
   * @param {string} resourceId — e.g. 'RES-001' or 'quick:calendar'
   * @returns {Object} { success: boolean }
   */
  function logResourceClick(email, resourceId) {
    if (!email || !resourceId) return { success: false };
    try {
      var cache = CacheService.getScriptCache();
      var rcKey = 'RC_' + String(resourceId).replace(/[^A-Za-z0-9_:-]/g, '');

      // Increment per-resource counter (6h TTL — max allowed by CacheService)
      var current = parseInt(cache.get(rcKey) || '0', 10) || 0;
      cache.put(rcKey, String(current + 1), 21600);

      // Increment grand total
      var total = parseInt(cache.get('RC_TOTAL') || '0', 10) || 0;
      cache.put('RC_TOTAL', String(total + 1), 21600);

      return { success: true };
    } catch (e) {
      log_('logResourceClick error', e.message);
      return { success: false };
    }
  }

  /**
   * Returns the total resource click count across all resources.
   * @returns {number}
   */
  function getResourceClickTotal() {
    try {
      var cache = CacheService.getScriptCache();
      return parseInt(cache.get('RC_TOTAL') || '0', 10) || 0;
    } catch (_e) {
      return 0;
    }
  }

  /**
   * Returns detailed resource usage stats for Union Stats > Resources sub-tab.
   * Reads per-resource click counts from CacheService (RC_<id>) and
   * cross-references with the Resources sheet for titles and categories.
   *
   * Returns:
   *   totalViews      — grand total clicks across all resources
   *   totalResources  — count of visible resources in Resources sheet
   *   uniqueViewed    — count of distinct resources that have at least 1 click
   *   avgViews        — average views per resource (totalViews / totalResources)
   *   topResources    — array of { id, title, category, views } sorted desc by views (top 10)
   *   byCategory      — { categoryName: totalViews } aggregated by category
   *   byCategoryCount — { categoryName: resourceCount } resources per category
   */
  function getResourceStats() {
    try {
      var ss = _getSS();
      if (!ss) return null;

      // Read Resources sheet
      var resSheet = ss.getSheetByName(SHEETS.RESOURCES);
      var resourceMap = {}; // id -> { title, category, visible }
      if (resSheet && resSheet.getLastRow() >= 2) {
        var resData = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, resSheet.getLastColumn()).getValues();
        for (var ri = 0; ri < resData.length; ri++) {
          var rId = String(col_(resData[ri], RESOURCES_COLS.RESOURCE_ID)).trim();
          var visible = String(col_(resData[ri], RESOURCES_COLS.VISIBLE)).trim().toLowerCase();
          if (rId && visible === 'yes') {
            resourceMap[rId] = {
              title: String(col_(resData[ri], RESOURCES_COLS.TITLE)).trim(),
              category: String(col_(resData[ri], RESOURCES_COLS.CATEGORY)).trim()
            };
          }
        }
      }

      // Read per-resource click counts from CacheService
      // Note: CacheService doesn't support getAll/enumerate, so we probe known resource IDs
      var cache = CacheService.getScriptCache();
      var clickMap = {}; // resourceId -> clickCount
      var totalViews = 0;
      var uniqueViewed = 0;

      // Build list of RC_ keys to probe from known resource IDs
      var knownIds = Object.keys(resourceMap);
      var rcKeys = knownIds.map(function(id) { return 'RC_' + id; });
      var cachedCounts = rcKeys.length > 0 ? (cache.getAll(rcKeys) || {}) : {};
      for (var ci = 0; ci < knownIds.length; ci++) {
        var rcKey = 'RC_' + knownIds[ci];
        var countStr = cachedCounts[rcKey];
        if (countStr) {
          var clicks = parseInt(countStr, 10) || 0;
          if (clicks > 0) {
            clickMap[knownIds[ci]] = clicks;
            totalViews += clicks;
            uniqueViewed++;
          }
        }
      }

      // Use RC_TOTAL as authoritative grand total (may include quick-link clicks)
      var rcTotal = parseInt(cache.get('RC_TOTAL') || '0', 10) || 0;
      if (rcTotal > totalViews) totalViews = rcTotal;

      var totalResources = Object.keys(resourceMap).length;
      var avgViews = totalResources > 0 ? Math.round((totalViews / totalResources) * 10) / 10 : 0;

      // Build top resources list and category aggregates
      var topList = [];
      var byCategory = {};
      var byCategoryCount = {};

      // Aggregate from Resources sheet (visible resources)
      for (var id in resourceMap) {
        var cat = resourceMap[id].category || 'Uncategorized';
        var views = clickMap[id] || 0;
        topList.push({ id: id, title: resourceMap[id].title, category: cat, views: views });
        byCategory[cat] = (byCategory[cat] || 0) + views;
        byCategoryCount[cat] = (byCategoryCount[cat] || 0) + 1;
      }

      // Also count clicks on quick-link resources (quick:*) not in the sheet
      for (var ck in clickMap) {
        if (ck.indexOf('quick:') === 0 && !resourceMap[ck]) {
          var quickLabel = ck.substring(6).replace(/_/g, ' ');
          quickLabel = quickLabel.charAt(0).toUpperCase() + quickLabel.slice(1);
          topList.push({ id: ck, title: quickLabel, category: 'Quick Links', views: clickMap[ck] });
          byCategory['Quick Links'] = (byCategory['Quick Links'] || 0) + clickMap[ck];
          byCategoryCount['Quick Links'] = (byCategoryCount['Quick Links'] || 0) + 1;
        }
      }

      // Sort top resources by views desc, take top 10
      topList.sort(function(a, b) { return b.views - a.views; });
      topList = topList.slice(0, 10);

      return {
        totalViews: totalViews,
        totalResources: totalResources,
        uniqueViewed: uniqueViewed,
        avgViews: avgViews,
        topResources: topList,
        byCategory: byCategory,
        byCategoryCount: byCategoryCount
      };
    } catch (e) {
      log_('getResourceStats error', e.message);
      return null;
    }
  }

  // ═══════════════════════════════════════
  // PUBLIC: Webapp Usage Analytics (v4.31.3)
  // ═══════════════════════════════════════
  //
  // Lightweight analytics stored in ScriptProperties with daily aggregation.
  // Property key scheme (all prefixed to avoid collisions):
  //   UA_D_YYYYMMDD         = JSON: { users: {email:1,...}, total: N }
  //   UA_T_YYYYMMDD_tabname = visit count for tab on date
  //   UA_H_YYYYMMDD_HH      = unique user count in that hour (for peak detection)
  //   UA_L_YYYYMMDD_location = unique user count by location on date
  //
  // Data is kept for 90 days max. Cleanup runs inside getUsageStats().

  /**
   * Logs a tab visit event for usage analytics.
   * @param {string} email — caller email (server-resolved)
   * @param {string} tab — tab identifier (e.g. 'home', 'grievances', 'resources')
   * @param {string} role — user role ('member' or 'steward')
   * @returns {{ success: boolean }}
   */
  function logTabVisit(email, tab, role) {
    if (!email || !tab) return { success: false };
    try {
      var props = PropertiesService.getScriptProperties();
      var tz = Session.getScriptTimeZone() || 'America/New_York';
      var now = new Date();
      var dateKey = Utilities.formatDate(now, tz, 'yyyyMMdd');
      var hourKey = Utilities.formatDate(now, tz, 'HH');
      var safeTab = String(tab).replace(/[^a-z0-9_-]/gi, '').substring(0, 30).toLowerCase();
      var safeEmail = String(email).toLowerCase().trim();
      // Hash email to avoid storing PII in ScriptProperties
      var emailHash = Utilities.base64Encode(
        Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, safeEmail)
      ).substring(0, 16);

      // 1. Daily unique users
      var dailyKey = 'UA_D_' + dateKey;
      var dailyRaw = props.getProperty(dailyKey);
      var daily = dailyRaw ? JSON.parse(dailyRaw) : { users: {}, total: 0 };
      if (!daily.users[emailHash]) {
        daily.users[emailHash] = 1;
      }
      daily.total = (daily.total || 0) + 1;
      props.setProperty(dailyKey, JSON.stringify(daily));

      // 2. Tab visit count
      var tabKey = 'UA_T_' + dateKey + '_' + safeTab;
      var tabCount = parseInt(props.getProperty(tabKey) || '0', 10) || 0;
      props.setProperty(tabKey, String(tabCount + 1));

      // 3. Hourly unique users (for peak detection)
      var hourlyKey = 'UA_H_' + dateKey + '_' + hourKey;
      var hourlyRaw = props.getProperty(hourlyKey);
      var hourly = hourlyRaw ? JSON.parse(hourlyRaw) : {};
      hourly[emailHash] = 1;
      props.setProperty(hourlyKey, JSON.stringify(hourly));

      // 4. Location-based usage (look up member's work location)
      try {
        var ss = _getSS();
        if (ss) {
          var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
          if (memberSheet && memberSheet.getLastRow() >= 2) {
            // Use MEMBER_COLS constants (1-indexed) for dynamic column access
            var emailIdx = MEMBER_COLS.EMAIL - 1;
            var locIdx = MEMBER_COLS.WORK_LOCATION - 1;
            if (emailIdx >= 0 && locIdx >= 0) {
              var mData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
              for (var i = 0; i < mData.length; i++) {
                if (String(mData[i][emailIdx]).toLowerCase().trim() === safeEmail) {
                  var loc = String(mData[i][locIdx]).trim();
                  if (loc) {
                    var locKey = 'UA_L_' + dateKey + '_' + loc.replace(/[^a-zA-Z0-9 _-]/g, '').substring(0, 40);
                    var locRaw = props.getProperty(locKey);
                    var locData = locRaw ? JSON.parse(locRaw) : {};
                    locData[emailHash] = 1;
                    props.setProperty(locKey, JSON.stringify(locData));
                  }
                  break;
                }
              }
            }
          }
        }
      } catch (_locErr) { /* non-critical */ }

      return { success: true };
    } catch (e) {
      log_('logTabVisit error', e.message);
      return { success: false };
    }
  }

  /**
   * Returns aggregated webapp usage stats for the Union Stats > Usage sub-tab.
   * Reads UA_* keys from ScriptProperties and computes:
   *   dau           — daily active users (today)
   *   wau           — weekly active users (last 7 days)
   *   mau           — monthly active users (last 30 days)
   *   peakHourly    — max concurrent users in a single hour (last 30 days)
   *   peakDate      — date of the peak
   *   repeatRate    — % of users who visited on 2+ distinct days (last 30 days)
   *   totalPageViews — total page views last 30 days
   *   avgDailyUsers — average DAU over last 30 days
   *   tabHeatmap    — { tabName: totalVisits } last 30 days
   *   byLocation    — { location: uniqueUsers } last 30 days
   *   dailyTrend    — [{ date, users, views }] last 30 days
   *   hourlyPattern — [{ hour, avgUsers }] averaged across last 30 days
   */
  function getUsageStats() {
    try {
      var props = PropertiesService.getScriptProperties();
      var allProps = props.getProperties();
      var tz = Session.getScriptTimeZone() || 'America/New_York';
      var now = new Date();
      var todayKey = Utilities.formatDate(now, tz, 'yyyyMMdd');

      // Date range helpers
      function _dateKey(daysAgo) {
        var d = new Date(now.getTime() - daysAgo * 86400000);
        return Utilities.formatDate(d, tz, 'yyyyMMdd');
      }
      function _formatDate(key) {
        return key.substring(0, 4) + '-' + key.substring(4, 6) + '-' + key.substring(6, 8);
      }

      var last30Keys = [];
      for (var d = 0; d < 30; d++) last30Keys.push(_dateKey(d));

      // Collect all user sets per day, page views per day
      var allUsers30 = {};    // email -> count of distinct days
      var allUsers7 = {};
      var dau = 0;
      var totalPageViews = 0;
      var dailyTrend = [];
      var activeDays = 0;     // days with any activity
      var totalDailyUsers = 0;

      for (var di = 0; di < last30Keys.length; di++) {
        var dk = last30Keys[di];
        var dailyProp = allProps['UA_D_' + dk];
        var dayUsers = 0;
        var dayViews = 0;

        if (dailyProp) {
          try {
            var parsed = JSON.parse(dailyProp);
            var emails = Object.keys(parsed.users || {});
            dayUsers = emails.length;
            dayViews = parsed.total || 0;
            for (var ei = 0; ei < emails.length; ei++) {
              allUsers30[emails[ei]] = (allUsers30[emails[ei]] || 0) + 1;
              if (di < 7) allUsers7[emails[ei]] = 1;
            }
            if (dk === todayKey) dau = dayUsers;
          } catch (_pe) { /* skip corrupt entries */ }
        }

        if (dayUsers > 0) {
          activeDays++;
          totalDailyUsers += dayUsers;
        }

        totalPageViews += dayViews;
        dailyTrend.push({ date: _formatDate(dk), users: dayUsers, views: dayViews });
      }
      dailyTrend.reverse(); // oldest first

      var wau = Object.keys(allUsers7).length;
      var mau = Object.keys(allUsers30).length;
      var avgDailyUsers = activeDays > 0 ? Math.round((totalDailyUsers / activeDays) * 10) / 10 : 0;

      // Repeat rate: users who visited on 2+ distinct days
      var repeatCount = 0;
      for (var u in allUsers30) {
        if (allUsers30[u] >= 2) repeatCount++;
      }
      var repeatRate = mau > 0 ? Math.round((repeatCount / mau) * 100) : 0;

      // Peak hourly users
      var peakHourly = 0;
      var peakDate = '';
      var hourlyBuckets = {}; // hour (0-23) -> total unique users across all days
      var hourlyDaysWithData = {}; // hour -> count of days with data for averaging

      for (var key in allProps) {
        if (key.indexOf('UA_H_') === 0) {
          var parts = key.substring(5).split('_'); // YYYYMMDD_HH
          if (parts.length === 2) {
            var hDateKey = parts[0];
            var hour = parts[1];
            // Only include last 30 days
            if (last30Keys.indexOf(hDateKey) >= 0) {
              try {
                var hUsers = Object.keys(JSON.parse(allProps[key]));
                var hCount = hUsers.length;
                if (hCount > peakHourly) {
                  peakHourly = hCount;
                  peakDate = _formatDate(hDateKey) + ' ' + hour + ':00';
                }
                var hNum = parseInt(hour, 10);
                hourlyBuckets[hNum] = (hourlyBuckets[hNum] || 0) + hCount;
                hourlyDaysWithData[hNum] = (hourlyDaysWithData[hNum] || 0) + 1;
              } catch (_he) { /* skip */ }
            }
          }
        }
      }

      // Hourly pattern (average users per hour)
      var hourlyPattern = [];
      for (var h = 0; h < 24; h++) {
        var hTotal = hourlyBuckets[h] || 0;
        var hDays = hourlyDaysWithData[h] || 1;
        hourlyPattern.push({ hour: String(h).length < 2 ? '0' + h : String(h), avgUsers: Math.round((hTotal / hDays) * 10) / 10 });
      }

      // Tab heatmap
      var tabHeatmap = {};
      for (var tKey in allProps) {
        if (tKey.indexOf('UA_T_') === 0) {
          var tParts = tKey.substring(5); // YYYYMMDD_tabname
          var tDateEnd = tParts.indexOf('_');
          if (tDateEnd > 0) {
            var tDate = tParts.substring(0, tDateEnd);
            if (last30Keys.indexOf(tDate) >= 0) {
              var tName = tParts.substring(tDateEnd + 1);
              tabHeatmap[tName] = (tabHeatmap[tName] || 0) + (parseInt(allProps[tKey], 10) || 0);
            }
          }
        }
      }

      // Location breakdown
      var byLocation = {};
      for (var lKey in allProps) {
        if (lKey.indexOf('UA_L_') === 0) {
          var lParts = lKey.substring(5); // YYYYMMDD_location
          var lDateEnd = lParts.indexOf('_');
          if (lDateEnd > 0) {
            var lDate = lParts.substring(0, lDateEnd);
            if (last30Keys.indexOf(lDate) >= 0) {
              var lName = lParts.substring(lDateEnd + 1);
              try {
                var lUsers = Object.keys(JSON.parse(allProps[lKey]));
                // Merge into location-level unique set
                if (!byLocation[lName]) byLocation[lName] = {};
                for (var li = 0; li < lUsers.length; li++) byLocation[lName][lUsers[li]] = 1;
              } catch (_le) { /* skip */ }
            }
          }
        }
      }
      // Convert location sets to counts
      var byLocationCounts = {};
      for (var loc in byLocation) {
        byLocationCounts[loc] = Object.keys(byLocation[loc]).length;
      }

      // Cleanup: remove keys older than 90 days
      var cutoff = _dateKey(90);
      var keysToDelete = [];
      for (var oldKey in allProps) {
        if (oldKey.indexOf('UA_') === 0) {
          // Extract date portion
          var datePortion = oldKey.replace(/^UA_[DTHL]_/, '').substring(0, 8);
          if (datePortion < cutoff && /^\d{8}$/.test(datePortion)) {
            keysToDelete.push(oldKey);
          }
        }
      }
      if (keysToDelete.length > 0) {
        props.deleteProperties(keysToDelete);
      }

      return {
        dau: dau,
        wau: wau,
        mau: mau,
        peakHourly: peakHourly,
        peakDate: peakDate,
        repeatRate: repeatRate,
        totalPageViews: totalPageViews,
        avgDailyUsers: avgDailyUsers,
        tabHeatmap: tabHeatmap,
        byLocation: byLocationCounts,
        dailyTrend: dailyTrend,
        hourlyPattern: hourlyPattern
      };
    } catch (e) {
      log_('getUsageStats error', e.message);
      return null;
    }
  }

  /**
   * Get or create a restricted Drive folder named after the spreadsheet.
   * Users shared on this folder cannot browse outside it.
   * The folder ID is cached in Script Properties for reuse.
   * @private
   * @returns {string} Folder URL (empty string on failure)
   */
  function getOrCreateSheetFolder_() {
    var props = PropertiesService.getScriptProperties();
    var storedId = props.getProperty('SHEET_DRIVE_FOLDER_ID');

    if (storedId) {
      try {
        var existing = DriveApp.getFolderById(storedId);
        return 'https://drive.google.com/drive/folders/' + existing.getId();
      } catch (_e) { log_('getOrCreateSheetFolder_', 'Error: ' + (_e.message || _e)); }
    }

    try {
      var ss = _getSS();
      if (!ss) return '';
      var folderName = ss.getName();

      // Search for existing folder with same name owned by this user
      var folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        var found = folders.next();
        props.setProperty('SHEET_DRIVE_FOLDER_ID', found.getId());
        return 'https://drive.google.com/drive/folders/' + found.getId();
      }

      // Create new folder
      var newFolder = DriveApp.createFolder(folderName);
      newFolder.setDescription('Restricted shared folder for ' + folderName + ' — auto-managed by Dashboard');
      props.setProperty('SHEET_DRIVE_FOLDER_ID', newFolder.getId());

      if (typeof logAuditEvent === 'function') {
        logAuditEvent('SHEET_DRIVE_FOLDER_CREATED', {
          folderId: newFolder.getId(),
          folderName: folderName
        });
      }

      return 'https://drive.google.com/drive/folders/' + newFolder.getId();
    } catch (e) {
      log_('getOrCreateSheetFolder_ error', e.message);
      return '';
    }
  }

  /**
   * Returns org-wide links (calendar, drive, survey form, etc.)
   * Drive link points to a restricted folder named after the spreadsheet.
   * @returns {Object}
   */
  function getOrgLinks() {
    try {
      var config = ConfigReader.getConfig();
      return {
        calendarUrl: config.calendarId ? 'https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(config.calendarId) : '',
        driveFolderUrl: getOrCreateSheetFolder_(),
        // surveyFormUrl removed v4.22.7 — survey is native webapp
        orgWebsite: config.orgWebsite || '',
      };
    } catch (_e) {
      return { calendarUrl: '', driveFolderUrl: '', orgWebsite: '' };
    }
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward — Member Management
  // ═══════════════════════════════════════

  /**
   * Returns all members assigned to a steward with summary stats.
   * @param {string} stewardEmail
   * @returns {Object[]} Array of member summaries
   */
  function getStewardMembers(stewardEmail) {
    if (!stewardEmail) return [];
    stewardEmail = String(stewardEmail).trim().toLowerCase();

    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap;
    var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
    if (stewardCol === -1) return [];

    var members = [];
    for (var i = 1; i < data.length; i++) {
      var assigned = String(data[i][stewardCol]).trim().toLowerCase();
      if (assigned !== stewardEmail) continue;
      var rec = _buildUserRecord(data[i], colMap);
      members.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        hasOpenGrievance: rec.hasOpenGrievance,
        duesStatus: rec.duesStatus,
        duesPaying: rec.duesPaying,
      });
    }

    members.sort(function(a, b) { return (a.name || '').localeCompare(b.name || ''); });
    return members;
  }

  /**
   * Returns ALL members from the directory (not filtered by steward assignment).
   * Same shape as getStewardMembers() for interchangeability.
   * @returns {Object[]} Array of member summaries
   */
  // S1: Route through _getCachedSheetData to eliminate redundant sheet reads
  function getAllMembers() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap || _buildColumnMap(data[0]);
    var members = [];
    var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      if (!rec.email && !rec.name) continue;
      members.push({
        name: rec.name,
        email: rec.email,
        role: rec.role,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        cubicle: rec.cubicle,
        hireDate: rec.hireDate,
        hasOpenGrievance: rec.hasOpenGrievance,
        duesStatus: rec.duesStatus,
        duesPaying: rec.duesPaying,
        director: rec.director,
        unit: rec.unit,
        supervisor: rec.supervisor,
        jobTitle: rec.jobTitle,
        assignedSteward: stewardCol !== -1 ? String(data[i][stewardCol]).trim() : '',
      });
    }
    members.sort(function(a, b) { return (a.name || '').localeCompare(b.name || ''); });
    return members;
  }

  /**
   * Returns survey completion tracking for a steward's assigned members.
   * Does NOT return actual survey responses — only completion status.
   * @param {string} stewardEmail
   * @param {string} [scope] - 'assigned' (default), 'location', or 'all'
   * @returns {Object} { total, completed, members: [{name, email, completed}] }
   */
  function getStewardSurveyTracking(stewardEmail, scope) {
    scope = scope || 'assigned';
    var members;
    if (scope === 'all') {
      members = getAllMembers();
    } else if (scope === 'location') {
      var allMembers = getAllMembers();
      // Find steward's location
      var stewardLoc = '';
      for (var l = 0; l < allMembers.length; l++) {
        if (String(allMembers[l].email).toLowerCase() === String(stewardEmail).toLowerCase()) {
          stewardLoc = allMembers[l].workLocation || '';
          break;
        }
      }
      if (stewardLoc) {
        members = allMembers.filter(function(m) {
          return m.workLocation && m.workLocation.toLowerCase() === stewardLoc.toLowerCase();
        });
      } else {
        members = getStewardMembers(stewardEmail);
      }
    } else {
      members = getStewardMembers(stewardEmail);
    }
    if (members.length === 0) return { total: 0, completed: 0, members: [] };

    // Pre-load _Survey_Tracking once and build email → status + participation maps
    // Single sheet read to avoid N+1 pattern.
    var surveyMap = {};
    var participationMap = {};
    try {
      var ss = _getSS();
      if (!ss) throw new Error('Spreadsheet unavailable');
      var trackSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
      if (trackSheet && trackSheet.getLastRow() > 1) {
        var tData = trackSheet.getDataRange().getValues();
        var tColMap = _buildColumnMap(tData[0]);
        var tEmailCol = _findColumn(tColMap, ['email', 'member email']);
        var tStatusCol = _findColumn(tColMap, ['status', 'completion status', 'completed']);
        var tDateCol = _findColumn(tColMap, ['completed date', 'date completed', 'completion date']);
        var tMissedCol = _findColumn(tColMap, ['total missed']);
        var tCompletedCol = _findColumn(tColMap, ['total completed']);
        if (tEmailCol !== -1) {
          for (var r = 1; r < tData.length; r++) {
            var rowEmail = String(tData[r][tEmailCol]).trim().toLowerCase();
            if (!rowEmail) continue;
            var rowStatus = tStatusCol !== -1 ? String(tData[r][tStatusCol]).trim().toLowerCase() : '';
            var rowCompleted = rowStatus === 'completed' || rowStatus === 'yes' || rowStatus === 'true';
            var rowDate = tDateCol !== -1 ? tData[r][tDateCol] : null;
            surveyMap[rowEmail] = {
              hasCompleted: rowCompleted,
              lastCompleted: rowDate instanceof Date ? _formatDate(rowDate) : (rowDate ? String(rowDate) : null),
            };
            var missed = tMissedCol !== -1 ? (parseInt(tData[r][tMissedCol], 10) || 0) : 0;
            var comp = tCompletedCol !== -1 ? (parseInt(tData[r][tCompletedCol], 10) || 0) : 0;
            participationMap[rowEmail] = { missed: missed, totalCompleted: comp };
          }
        }
      }
    } catch (e) {
      log_('getStewardSurveyTracking', 'survey pre-load error: ' + e.message);
    }

    // Find steward's own location + cubicle for proximity scoring
    var stewardInfo = { workLocation: '', officeDays: '', cubicle: '' };
    for (var si = 0; si < members.length; si++) {
      if (String(members[si].email).toLowerCase() === String(stewardEmail).toLowerCase()) {
        stewardInfo.workLocation = members[si].workLocation || '';
        stewardInfo.officeDays = members[si].officeDays || '';
        stewardInfo.cubicle = members[si].cubicle || '';
        break;
      }
    }
    // If steward not in members list, look them up directly
    if (!stewardInfo.workLocation) {
      var sUser = findUserByEmail(stewardEmail);
      if (sUser) {
        stewardInfo.workLocation = sUser.workLocation || '';
        stewardInfo.officeDays = sUser.officeDays || '';
        stewardInfo.cubicle = sUser.cubicle || '';
      }
    }

    var tracking = [];
    var completedCount = 0;

    for (var i = 0; i < members.length; i++) {
      var memberEmail = String(members[i].email || '').trim().toLowerCase();
      var status = surveyMap[memberEmail] || { hasCompleted: false, lastCompleted: null };
      var participation = participationMap[memberEmail] || { missed: 0, totalCompleted: 0 };
      var totalSurveys = participation.missed + participation.totalCompleted;
      var participationRate = totalSurveys > 0 ? Math.round((participation.totalCompleted / totalSurveys) * 100) : null;

      tracking.push({
        name: members[i].name,
        email: members[i].email,
        completed: status.hasCompleted,
        lastCompleted: status.lastCompleted,
        workLocation: members[i].workLocation || '',
        officeDays: members[i].officeDays || '',
        cubicle: members[i].cubicle || '',
        hireDate: members[i].hireDate || '',
        participationRate: participationRate,
        totalCompleted: participation.totalCompleted,
        totalMissed: participation.missed,
      });
      if (status.hasCompleted) completedCount++;
    }

    return {
      total: members.length,
      completed: completedCount,
      stewardInfo: stewardInfo,
      members: tracking
    };
  }

  /**
   * Sends a broadcast email to filtered members assigned to a steward.
   * @param {string} stewardEmail - The steward sending the broadcast
   * @param {Object} filter - { location, officeDays }
   * @param {string} message - The message body
   * @returns {Object} { success, sentCount, message }
   */
  function sendBroadcastMessage(stewardEmail, filter, message, customSubject) {
    try {
    if (!stewardEmail || !message) return { success: false, sentCount: 0, message: 'Missing required fields.' };

    // T6-1: Per-user rate limiting — max 3 broadcasts per hour (matches sendUnitBroadcast)
    var _bcRateKey = 'BROADCAST_RATE_' + stewardEmail.toLowerCase().trim();
    var _bcCache = CacheService.getScriptCache();
    var _bcRateCount = parseInt(_bcCache.get(_bcRateKey) || '0', 10);
    if (_bcRateCount >= 3) {
      return { success: false, sentCount: 0, message: 'Broadcast rate limit reached (3 per hour). Please try again later.' };
    }
    _bcCache.put(_bcRateKey, String(_bcRateCount + 1), 3600);

    // Scope: 'mine' = only members assigned to this steward, 'all' = all members
    var scope = (filter && filter.scope === 'all') ? 'all' : 'mine';
    var members = (scope === 'all') ? getAllMembers() : getStewardMembers(stewardEmail);
    if (members.length === 0) return { success: false, sentCount: 0, message: 'No members found.' };

    // Apply filters
    var filtered = members.filter(function(m) {
      // Multi-location: split comma-delimited string and check any match
      if (filter && filter.location) {
        var filterLocs = filter.location.split(',').map(function(l) { return l.trim().toLowerCase(); });
        var memberLoc = (m.workLocation || '').toLowerCase();
        if (!filterLocs.some(function(l) { return l === memberLoc; })) return false;
      }
      // Office days: any match
      if (filter && filter.officeDays && m.officeDays) {
        var filterDays = filter.officeDays.toLowerCase().split(',').map(function(d) { return d.trim(); });
        var memberDays = m.officeDays.toLowerCase();
        if (!filterDays.some(function(d) { return memberDays.indexOf(d) !== -1; })) return false;
      }
      // Dues paying: 'paying' = only dues paying, 'nonpaying' = only non-dues paying
      // null means column is absent — treat as paying (benefit of the doubt)
      if (filter && filter.duesPaying) {
        var dp = m.duesPaying; // true, false, or null (column absent)
        if (filter.duesPaying === 'paying' && dp === false) return false;
        if (filter.duesPaying === 'nonpaying' && dp !== false) return false;
      }
      return true;
    });

    if (filtered.length === 0) return { success: false, sentCount: 0, message: 'No members match the selected filters.' };

    var sentCount = 0;
    var failedCount = 0;
    var failures = [];
    var config = ConfigReader.getConfig();
    var autoSubject = config.orgAbbrev + ' - Message from your ' + config.stewardLabel;
    var subject = (customSubject && String(customSubject).trim()) ? String(customSubject).trim() : autoSubject;
    // T6-1: Sanitize subject and message for formula injection
    subject = typeof escapeForFormula === 'function' ? escapeForFormula(subject) : subject;
    message = typeof escapeForFormula === 'function' ? escapeForFormula(message) : message;

    for (var i = 0; i < filtered.length; i++) {
      try {
        MailApp.sendEmail(filtered[i].email, subject, message);
        sentCount++;
      } catch (e) {
        failedCount++;
        failures.push({ email: filtered[i].email, error: e.message });
        log_('sendBroadcastMessage', 'Broadcast send error for ' + maskEmail(filtered[i].email) + ': ' + e.message);
      }
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('BROADCAST_SENT', {
        steward: stewardEmail,
        scope: scope,
        recipientCount: sentCount,
        failedCount: failedCount,
        filter: JSON.stringify(filter || {}),
      });
    }

    var msg = 'Sent to ' + sentCount + ' member(s).';
    if (failedCount > 0) msg += ' Failed: ' + failedCount + '.';
    return { success: sentCount > 0, sentCount: sentCount, failedCount: failedCount, failures: failures, message: msg };
    } catch (outerErr) {
      handleError(outerErr, 'DataService.sendBroadcastMessage', ERROR_LEVEL.ERROR);
      return { success: false, sentCount: 0, message: 'An error occurred while sending the broadcast.' };
    }
  }

  /**
   * Returns aggregated quarterly survey results.
   * Privacy threshold: only returns data if 30+ responses.
   * @returns {Object} { available, count, threshold, sections }
   */
  function getSurveyResults() {
    try {
      var ss = _getSS();
      if (!ss) return { available: false, count: 0, threshold: 30, sections: [] };
      var satSheet = ss.getSheetByName(SHEETS.SATISFACTION);
      if (!satSheet) return { available: false, count: 0, threshold: 30, sections: [] };

      var data = satSheet.getDataRange().getValues();
      if (data.length <= 1) return { available: false, count: 0, threshold: 30, sections: [] };

      var responseCount = data.length - 1;
      if (responseCount < 30) {
        return { available: false, count: responseCount, threshold: 30, sections: [] };
      }

      // Use SATISFACTION_SECTIONS from 01_Core.gs to aggregate
      var sections = [];
      if (typeof SATISFACTION_SECTIONS !== 'undefined') {
        for (var key in SATISFACTION_SECTIONS) {
          var sec = SATISFACTION_SECTIONS[key];
          if (!sec.scale) continue; // Skip non-scale sections

          var qCols = sec.questions;
          var totalScore = 0;
          var totalCount = 0;
          var questionAvgs = [];

          for (var q = 0; q < qCols.length; q++) {
            var colIdx = qCols[q] - 1; // 0-indexed
            var qSum = 0;
            var qCount = 0;
            for (var r = 1; r < data.length; r++) {
              var val = Number(data[r][colIdx]);
              if (!isNaN(val) && val > 0) {
                qSum += val;
                qCount++;
              }
            }
            var avg = qCount > 0 ? qSum / qCount : 0;
            var text = sec.questionTexts && sec.questionTexts[q] ? sec.questionTexts[q] : '';
            questionAvgs.push({ question: colIdx + 1, text: text, avg: Math.round(avg * 10) / 10 });
            totalScore += qSum;
            totalCount += qCount;
          }

          sections.push({
            name: sec.name,
            key: key,
            avg: totalCount > 0 ? Math.round((totalScore / totalCount) * 10) / 10 : 0,
            questions: questionAvgs,
          });
        }
      }

      return { available: true, count: responseCount, threshold: 30, sections: sections };
    } catch (e) {
      log_('DataService.getSurveyResults error', e.message);
      return { available: false, count: 0, threshold: 30, sections: [], error: e.message };
    }
  }

  // ═══════════════════════════════════════
  // PUBLIC: Shared Data (non-PII)
  // ═══════════════════════════════════════

  /**
   * Returns the list of units from the directory (no PII).
   * Useful for filter dropdowns.
   * @returns {string[]} Unique unit names
   */
  function getUnits() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];

    var data = cached.data;
    var colMap = cached.colMap;
    var unitCol = _findColumn(colMap, HEADERS.memberUnit);
    if (unitCol === -1) return [];

    var units = {};
    for (var i = 1; i < data.length; i++) {
      var unit = String(data[i][unitCol]).trim();
      if (unit) units[unit] = true;
    }

    return Object.keys(units).sort();
  }

  // ═══════════════════════════════════════
  // PRIVATE: Sheet & Column Helpers
  // ═══════════════════════════════════════

  /**
   * Returns a sheet by name from the active spreadsheet, logging errors if not found.
   * @param {string} name - Sheet name
   * @returns {Sheet|null}
   */
  function _getSheet(name) {
    var ss = _getSS();
    if (!ss) {
      log_('DataService', 'getActiveSpreadsheet() returned null for sheet "' + name + '"');
      return null;
    }
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      log_('DataService', 'Sheet "' + name + '" not found.');
    }
    return sheet;
  }

  /**
   * Builds a map of normalized header names to column indices.
   * @param {Array} headerRow
   * @returns {Object} { normalizedName: columnIndex }
   */
  function _buildColumnMap(headerRow) {
    var map = {};
    for (var i = 0; i < headerRow.length; i++) {
      var normalized = String(headerRow[i]).trim().toLowerCase();
      if (normalized) {
        map[normalized] = i;
      }
    }
    return map;
  }

  /**
   * Finds a column index given an array of possible header names.
   * Returns the index of the first match, or -1 if none found.
   * @param {Object} colMap - from _buildColumnMap
   * @param {string[]} aliases - possible header names
   * @returns {number} Column index or -1
   */
  function _findColumn(colMap, aliases) {
    for (var i = 0; i < aliases.length; i++) {
      var key = aliases[i].toLowerCase();
      if (Object.prototype.hasOwnProperty.call(colMap, key)) {
        return colMap[key];
      }
    }
    return -1;
  }

  /**
   * Safe getter — returns cell value or default.
   * @param {Array} row - Row data array
   * @param {Object} colMap - Column name-to-index map
   * @param {string[]} aliases - Possible header names
   * @param {*} defaultVal - Fallback value if column missing
   * @returns {*} Cell value or defaultVal
   */
  function _getVal(row, colMap, aliases, defaultVal) {
    var col = _findColumn(colMap, aliases);
    if (col === -1 || col >= row.length) return defaultVal !== undefined ? defaultVal : '';
    var val = row[col];
    return (val === null || val === undefined) ? (defaultVal !== undefined ? defaultVal : '') : val;
  }

  /**
   * Constructs a normalized user record from a Member Directory row.
   * @param {Array} row - Row data array
   * @param {Object} colMap - Column name-to-index map
   * @returns {Object} User record with email, name, role, contact info, etc.
   */
  function _buildUserRecord(row, colMap) {
    var firstName = String(_getVal(row, colMap, HEADERS.memberFirstName, '')).trim();
    var lastName = String(_getVal(row, colMap, HEADERS.memberLastName, '')).trim();
    var fullName = String(_getVal(row, colMap, HEADERS.memberName, '')).trim();

    if (!fullName && (firstName || lastName)) {
      fullName = (firstName + ' ' + lastName).trim();
    }

    var roleRaw = String(_getVal(row, colMap, HEADERS.memberRole, 'member')).trim().toLowerCase();
    // Normalize role
    var role = 'member';
    if (roleRaw === 'steward' || roleRaw === 'rep' || roleRaw === 'representative') role = 'steward';
    if (roleRaw === 'both' || roleRaw === 'steward/member') role = 'both';

    // Also check the IS_STEWARD column (set by seed and manual entry)
    var isStewardRaw = String(_getVal(row, colMap, HEADERS.memberIsSteward, '')).trim().toLowerCase();
    if (isStewardRaw === 'yes' || isStewardRaw === 'true' || isStewardRaw === '1') {
      if (role === 'member') role = 'both';  // Upgrade to dual-role (member + steward)
    }
    var isLeader = isStewardRaw === 'member leader';

    var hasGrievance = String(_getVal(row, colMap, HEADERS.memberHasOpenGrievance, '')).trim().toLowerCase();

    return {
      email: String(_getVal(row, colMap, HEADERS.memberEmail, '')).trim().toLowerCase(),
      name: fullName,
      firstName: firstName,
      lastName: lastName,
      role: role,
      unit: String(_getVal(row, colMap, HEADERS.memberUnit, '')).trim(),
      phone: String(_getVal(row, colMap, HEADERS.memberPhone, '')).trim() || null,
      joined: _getVal(row, colMap, HEADERS.memberJoined, ''),
      duesStatus: String(_getVal(row, colMap, HEADERS.memberDuesStatus, '')).trim(),
      duesPaying: (function() {
        // First: check explicit Dues Paying column (if it exists)
        var raw = _getVal(row, colMap, HEADERS.memberDuesPaying, null);
        if (raw !== null && raw !== '') {
          if (typeof raw === 'boolean') return raw;
          var s = String(raw).trim().toLowerCase();
          return s === 'true' || s === 'yes' || s === '1';
        }
        // Fallback: derive from Dues Status column ('Current' → paying; 'Past Due'/'Inactive' → not paying)
        var duesStatus = String(_getVal(row, colMap, HEADERS.memberDuesStatus, '')).trim().toLowerCase();
        if (duesStatus === '') return null; // column absent — unknown
        var NON_PAYING = ['past due', 'inactive', 'delinquent', 'lapsed', 'non-paying', 'no', 'non-member'];
        for (var ni = 0; ni < NON_PAYING.length; ni++) {
          if (duesStatus === NON_PAYING[ni]) return false;
        }
        // 'current', 'active', 'paid', 'yes', or any unrecognized value → treat as paying
        return true;
      })(),
      memberId: String(_getVal(row, colMap, HEADERS.memberId, '')).trim(),
      workLocation: String(_getVal(row, colMap, HEADERS.memberWorkLocation, '')).trim(),
      officeDays: String(_getVal(row, colMap, HEADERS.memberOfficeDays, '')).trim(),
      assignedSteward: String(_getVal(row, colMap, HEADERS.memberAssignedSteward, '')).trim().toLowerCase(),
      isSteward: role === 'steward' || role === 'both',
      isLeader: isLeader,
      hasOpenGrievance: hasGrievance === 'yes' || hasGrievance === 'true' || hasGrievance === '1',
      street: String(_getVal(row, colMap, HEADERS.memberStreet, '')).trim(),
      city: String(_getVal(row, colMap, HEADERS.memberCity, '')).trim(),
      state: String(_getVal(row, colMap, HEADERS.memberState, '')).trim(),
      zip: String(_getVal(row, colMap, HEADERS.memberZip, '')).trim(),
      supervisor: String(_getVal(row, colMap, HEADERS.memberSupervisor, '')).trim(),
      director: String(_getVal(row, colMap, HEADERS.memberManager, '')).trim(),
      jobTitle: String(_getVal(row, colMap, HEADERS.memberJobTitle, '')).trim(),
      memberAdminFolderUrl: String(_getVal(row, colMap, HEADERS.memberAdminFolderUrl, '')).trim(),
      sharePhone: (function() {
        var raw = _getVal(row, colMap, HEADERS.memberSharePhone, null);
        // Column absent → treat as false (opt-in required)
        if (raw === null || raw === '') return false;
        if (typeof raw === 'boolean') return raw;
        var s = String(raw).trim().toLowerCase();
        return s === 'yes' || s === 'true' || s === '1';
      })(),
      cubicle: String(_getVal(row, colMap, HEADERS.memberCubicle, '')).trim(),
      employeeId: String(_getVal(row, colMap, HEADERS.memberEmployeeId, '')).trim(),
      hireDate: (function() {
        var raw = _getVal(row, colMap, HEADERS.memberHireDate, '');
        if (raw instanceof Date) return _formatDate(raw);
        return String(raw || '').trim();
      })(),
      openRate: String(_getVal(row, colMap, HEADERS.memberOpenRate, '')).trim(),
      shirtSize: String(_getVal(row, colMap, HEADERS.memberShirtSize, '')).trim(),
    };
  }

  /**
   * Finds a user by full name (case-insensitive).
   * Used as fallback when assignedSteward contains a name instead of email.
   * @param {string} name - Full name to search for
   * @returns {Object|null} User record or null
   */
  // S5: Route through _getCachedSheetData to avoid redundant sheet reads
  function _findUserByName(name) {
    if (!name) return null;
    name = String(name).trim().toLowerCase();

    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return null;

    var data = cached.data;
    var colMap = cached.colMap || _buildColumnMap(data[0]);
    var firstNameCol = _findColumn(colMap, HEADERS.memberFirstName);
    var lastNameCol = _findColumn(colMap, HEADERS.memberLastName);
    var fullNameCol = _findColumn(colMap, HEADERS.memberName);

    for (var i = 1; i < data.length; i++) {
      // Check full name column first
      if (fullNameCol !== -1) {
        var full = String(data[i][fullNameCol]).trim().toLowerCase();
        if (full === name) return _buildUserRecord(data[i], colMap);
      }
      // Check first+last name combination
      if (firstNameCol !== -1 && lastNameCol !== -1) {
        var combined = (String(data[i][firstNameCol]).trim() + ' ' + String(data[i][lastNameCol]).trim()).toLowerCase();
        if (combined === name) return _buildUserRecord(data[i], colMap);
      }
    }

    return null;
  }

  /**
   * Constructs a normalized grievance record from a Grievance Log row.
   * @param {Array} row - Row data array
   * @param {Object} colMap - Column name-to-index map
   * @returns {Object} Grievance record with id, status, deadlines, etc.
   */
  function _buildGrievanceRecord(row, colMap) {
    var deadlineRaw = _getVal(row, colMap, HEADERS.grievanceDeadline, null);
    var deadlineDays = null;
    var deadlineFormatted = '—';

    if (deadlineRaw && deadlineRaw instanceof Date) {
      var now = new Date();
      var diff = deadlineRaw.getTime() - now.getTime();
      deadlineDays = Math.ceil(diff / (24 * 60 * 60 * 1000));
      deadlineFormatted = _formatDate(deadlineRaw);
    } else if (deadlineRaw) {
      var parsed = new Date(deadlineRaw);
      if (!isNaN(parsed.getTime())) {
        now = new Date();
        diff = parsed.getTime() - now.getTime();
        deadlineDays = Math.ceil(diff / (24 * 60 * 60 * 1000));
        deadlineFormatted = _formatDate(parsed);
      }
    }

    // Fallback: if no date column resolved deadlineDays, read 'Days to Deadline'
    // which is a computed numeric column (or the text "Overdue") in the Grievance Log.
    if (deadlineDays === null) {
      var dtdRaw = _getVal(row, colMap, ['days to deadline'], null);
      if (dtdRaw !== null && dtdRaw !== '') {
        var dtdStr = String(dtdRaw).trim().toLowerCase();
        if (dtdStr === 'overdue') {
          deadlineDays = -1;
          deadlineFormatted = 'Overdue';
        } else {
          var dtdNum = parseFloat(dtdStr);
          if (!isNaN(dtdNum)) deadlineDays = Math.ceil(dtdNum);
        }
      }
    }

    var filedRaw = _getVal(row, colMap, HEADERS.grievanceFiled, null);
    var filedFormatted = '';
    var filedTimestamp = 0;
    if (filedRaw && filedRaw instanceof Date) {
      filedFormatted = _formatDate(filedRaw);
      filedTimestamp = filedRaw.getTime();
    } else if (filedRaw) {
      parsed = new Date(filedRaw);
      if (!isNaN(parsed.getTime())) {
        filedFormatted = _formatDate(parsed);
        filedTimestamp = parsed.getTime();
      }
    }

    var status = String(_getVal(row, colMap, HEADERS.grievanceStatus, 'new')).trim().toLowerCase();

    // Auto-detect overdue: only flag active/open cases whose deadline has passed
    var TERMINAL_STATUSES = ['resolved', 'won', 'denied', 'settled', 'withdrawn', 'closed'];
    if (deadlineDays !== null && deadlineDays < 0 && TERMINAL_STATUSES.indexOf(status) === -1) {
      status = 'overdue';
    }

    // Parse closed date (same pattern as filed date above)
    var closedRaw = _getVal(row, colMap, HEADERS.grievanceDateClosed, null);
    var closedTimestamp = 0;
    if (closedRaw instanceof Date) {
      closedTimestamp = closedRaw.getTime();
    } else if (closedRaw) {
      var parsedClosed = new Date(closedRaw);
      if (!isNaN(parsedClosed.getTime())) {
        closedTimestamp = parsedClosed.getTime();
      }
    }

    return {
      id: String(_getVal(row, colMap, HEADERS.grievanceId, '')).trim(),
      memberEmail: String(_getVal(row, colMap, HEADERS.grievanceMemberEmail, '')).trim().toLowerCase(),
      memberName: (String(_getVal(row, colMap, HEADERS.grievanceMemberFirstName, '')).trim() + ' ' +
                   String(_getVal(row, colMap, HEADERS.grievanceMemberLastName, '')).trim()).trim(),
      status: status,
      step: String(_getVal(row, colMap, HEADERS.grievanceStep, '')).trim(),
      deadline: deadlineFormatted,
      deadlineDays: deadlineDays,
      filed: filedFormatted,
      filedTimestamp: filedTimestamp,
      closedTimestamp: closedTimestamp,
      steward: String(_getVal(row, colMap, HEADERS.grievanceSteward, '')).trim().toLowerCase(),
      unit: String(_getVal(row, colMap, HEADERS.grievanceUnit, '')).trim(),
      issueCategory: String(_getVal(row, colMap, HEADERS.grievanceIssueCategory, '')).trim(),
      priority: String(_getVal(row, colMap, HEADERS.grievancePriority, 'medium')).trim().toLowerCase(),
      notes: String(_getVal(row, colMap, HEADERS.grievanceNotes, '')).trim(),

      // v4.32.1 — Google Drive folder URL for this grievance's case files.
      // Reads from GRIEVANCE_COLS.DRIVE_FOLDER_URL (col 33, header "Drive Folder URL").
      // Value is a full URL like "https://drive.google.com/drive/folders/{folderId}"
      // or null if no folder has been created yet (e.g., Draft grievances).
      //
      // Consumed by:
      //   - getMemberGrievanceDriveUrl() → returns URL for member's active case
      //   - dataGetMemberCaseFolderUrl() → steward-only wrapper that resolves case folder
      //
      // The || null coercion ensures callers can do a simple truthy check
      // (e.g., `if (target.driveFolderUrl)`) rather than checking for empty string.
      //
      // IF THIS BREAKS: driveFolderUrl will be null. getMemberGrievanceDriveUrl()
      // returns null, dataGetMemberCaseFolderUrl() returns { success: false }.
      // UI degrades gracefully: "No Drive folder linked to this case" message shown.
      driveFolderUrl: String(_getVal(row, colMap, HEADERS.grievanceDriveFolderUrl, '')).trim() || null,
      driveFolderId: String(_getVal(row, colMap, HEADERS.grievanceDriveFolderId, '')).trim() || null,
      signatureStatus: String(_getVal(row, colMap, HEADERS.grievanceSignatureStatus, '')).trim() || null,
      signatureToken: String(_getVal(row, colMap, HEADERS.grievanceSignatureToken, '')).trim() || null,
      signedDate: _getVal(row, colMap, HEADERS.grievanceSignedDate, null),
    };
  }

  /**
   * Formats a Date object as "Mon DD, YYYY" string.
   * @param {Date} date
   * @returns {string} Formatted date string
   */
  function _formatDate(date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  }

  // ═══════════════════════════════════════
  // PRIVATE: Member Contact Drive helpers (v4.20.22)
  // ═══════════════════════════════════════

  /**
   * Bridge to getOrCreateMemberAdminFolder() in 05_Integrations.gs.
   * Returns the per-member master admin folder (Members/LastName, FirstName/).
   * Contact Log sheet lives directly inside this folder.
   * Steward-only — never shared with member.
   * @param {string} memberEmail
   * @returns {Folder|null}
   */
  function _getMemberAdminFolder_(memberEmail) {
    if (typeof getOrCreateMemberAdminFolder !== 'function') {
      log_('_getMemberAdminFolder_', 'getOrCreateMemberAdminFolder not available');
      return null;
    }
    return getOrCreateMemberAdminFolder(memberEmail);
  }

  /**
   * Legacy stub — delegates to _getMemberAdminFolder_.
   * @param {string} memberEmail
   * @returns {Folder|null}
   */
  function getOrCreateMemberContactFolder_(memberEmail) {
    return _getMemberAdminFolder_(memberEmail);
  }

  /**
   * Gets or creates the Contact Log Google Sheet inside the member's Drive folder.
   * Sheet name: "Contact Log — [folder name]"
   * Columns: Date | Steward | Contact Type | Notes | Duration
   * @param {Folder} memberFolder
   * @param {string} memberFolderName  — used for sheet title
   * @returns {Spreadsheet|null}
   */
  function getOrCreateMemberContactSheet_(memberFolder, memberFolderName) {
    try {
      var sheetTitle = 'Contact Log \u2014 ' + memberFolderName;
      var fileIter = memberFolder.getFilesByName(sheetTitle);
      if (fileIter.hasNext()) {
        return SpreadsheetApp.open(fileIter.next());
      }
      // Create new sheet in member's folder
      var ss = SpreadsheetApp.create(sheetTitle);
      var ssFile = DriveApp.getFileById(ss.getId());
      memberFolder.addFile(ssFile);
      DriveApp.getRootFolder().removeFile(ssFile); // move out of My Drive root

      var sheet = ss.getActiveSheet();
      sheet.setName(CONTACT_SHEET_TAB_);
      var headers = ['Date', 'Steward', 'Contact Type', 'Notes', 'Duration'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
           .setFontWeight('bold')
           .setBackground('#1a365d')
           .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 130); // Date
      sheet.setColumnWidth(2, 180); // Steward
      sheet.setColumnWidth(3, 120); // Contact Type
      sheet.setColumnWidth(4, 320); // Notes
      sheet.setColumnWidth(5, 100); // Duration
      return ss;
    } catch (e) {
      log_('getOrCreateMemberContactSheet_ error', e.message);
      return null;
    }
  }

  // ═══════════════════════════════════════
  // PUBLIC: Contact Log (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns the _Contact_Log hidden sheet, creating it with headers if absent.
   * @returns {Sheet|null}
   */
  function _ensureContactLog() {
    var ss = _getSS();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.CONTACT_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.CONTACT_LOG);
      sheet.getRange(1, 1, 1, 9).setValues([['ID', 'Steward Email', 'Member Email', 'Contact Type', 'Date', 'Notes', 'Duration', 'Created', 'Member Name']]);
      sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Logs a steward-member contact to the hidden sheet, Member Directory snapshot, and per-member Drive sheet.
   * @param {string} stewardEmail
   * @param {string} memberEmail
   * @param {string} contactType - e.g. 'Phone', 'Email', 'In-Person'
   * @param {string} notes
   * @param {string} duration
   * @param {string} memberName - Optional display name
   * @returns {Object} { success: boolean, message: string }
   */
  function logMemberContact(stewardEmail, memberEmail, contactType, notes, duration, memberName) {
    if (!stewardEmail || !memberEmail || !contactType) return { success: false, message: 'Missing fields.' };

    // ── 0. Resolve member display name if not provided by client ─────────
    var resolvedName = (memberName || '').trim();
    if (!resolvedName && typeof findUserByEmail === 'function') {
      var mRecord = findUserByEmail(memberEmail);
      if (mRecord && mRecord.name) resolvedName = mRecord.name;
    }

    // ── 1. Append to _Contact_Log hidden sheet (fast dashboard queries) ────
    var sheet = _ensureContactLog();
    var id = 'CL_' + Date.now().toString(36);
    sheet.appendRow([id, stewardEmail.toLowerCase().trim(), memberEmail.toLowerCase().trim(), escapeForFormula(contactType), new Date(), escapeForFormula((notes || '').substring(0, 500)), duration || '', new Date(), resolvedName]);
    if (typeof logAuditEvent === 'function') logAuditEvent('CONTACT_LOG', { steward: stewardEmail, member: memberEmail, type: contactType });

    // ── 2. Resolve steward display name once (used in both writeback and Drive sheet) ──
    var sRecord = (typeof findUserByEmail === 'function') ? findUserByEmail(stewardEmail) : null;
    var sName   = (sRecord && sRecord.name) ? sRecord.name : stewardEmail;

    // ── 3. Writeback: update Member Directory snapshot columns ────────────
    try {
      var ss = _getSS();
      if (ss) {
        var memberDir = ss.getSheetByName(MEMBER_SHEET);
        if (memberDir) {
          var mData    = memberDir.getDataRange().getValues();
          // Use MEMBER_COLS constants (1-indexed, subtract 1 for array access)
          // Resolve contact columns from actual headers at write time (v4.51.1)
          var _lmcCols = resolveColumnsByHeader_(memberDir, [
            { key: 'EMAIL',    header: 'Email',                fallback: MEMBER_COLS.EMAIL },
            { key: 'CONTACT',  header: 'Recent Contact Date',  fallback: MEMBER_COLS.RECENT_CONTACT_DATE },
            { key: 'STEWARD',  header: 'Contact Steward',      fallback: MEMBER_COLS.CONTACT_STEWARD },
            { key: 'NOTES',    header: 'Contact Notes',        fallback: MEMBER_COLS.CONTACT_NOTES }
          ]);
          var emailIdx          = _lmcCols.EMAIL - 1;
          var recentContactIdx  = _lmcCols.CONTACT - 1;
          var contactStewardIdx = _lmcCols.STEWARD - 1;
          var contactNotesIdx   = _lmcCols.NOTES - 1;
          if (emailIdx >= 0) {
            var mEmailNorm = memberEmail.toLowerCase().trim();
            for (var r = 1; r < mData.length; r++) {
              if (String(mData[r][emailIdx]).toLowerCase().trim() === mEmailNorm) {
                if (recentContactIdx  !== -1) memberDir.getRange(r + 1, recentContactIdx  + 1).setValue(new Date());
                if (contactStewardIdx !== -1) memberDir.getRange(r + 1, contactStewardIdx + 1).setValue(sName);
                if (contactNotesIdx   !== -1 && notes) memberDir.getRange(r + 1, contactNotesIdx + 1).setValue(escapeForFormula(String(notes).substring(0, 500)));
                break;
              }
            }
          }
        }
      }
    } catch (wbErr) {
      log_('logMemberContact writeback error', wbErr.message);
    }

    // ── 4. Append to per-member Drive contact log sheet ───────────────────
    // getOrCreateMemberContactFolder_ → _getMemberAdminFolder_ → getOrCreateMemberAdminFolder
    // which returns { masterFolder, grievancesFolder } — extract masterFolder explicitly.
    try {
      var adminResult = getOrCreateMemberContactFolder_(memberEmail.toLowerCase().trim());
      var masterFolder = (adminResult && adminResult.masterFolder) ? adminResult.masterFolder : null;
      if (masterFolder) {
        var folderName = masterFolder.getName();
        var contactSS = getOrCreateMemberContactSheet_(masterFolder, folderName);
        if (contactSS) {
          var logSheet = contactSS.getSheetByName(CONTACT_SHEET_TAB_) || contactSS.getActiveSheet();
          logSheet.appendRow([new Date(), sName, escapeForFormula(contactType), escapeForFormula((notes || '').substring(0, 500)), duration || '']);
        }
      }
    } catch (driveErr) {
      log_('logMemberContact Drive sheet error', driveErr.message);
      // Non-fatal — _Contact_Log and Member Directory snapshot already updated
    }

    return { success: true, message: 'Contact logged.' };
  }

  /**
   * Returns contact history between a steward and a specific member, sorted newest first.
   * @param {string} stewardEmail
   * @param {string} memberEmail
   * @returns {Object[]} Array of contact records
   */
  function getMemberContactHistory(stewardEmail, memberEmail) {
    if (!stewardEmail || !memberEmail) return [];
    var sheet = _ensureContactLog();
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    var mEmail = memberEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === sEmail && String(data[i][2]).toLowerCase().trim() === mEmail) {
        var rawDate = data[i][4];
        results.push({ id: data[i][0], type: data[i][3],
          date: rawDate instanceof Date ? _formatDate(rawDate) : String(rawDate),
          _ts: rawDate instanceof Date ? rawDate.getTime() : 0,
          notes: data[i][5], duration: data[i][6], memberName: data[i][8] || '' });
      }
    }
    // H-16: sort by numeric timestamp — string comparison of formatted dates is not chronological
    results.sort(function(a, b) { return (b._ts || 0) - (a._ts || 0); });
    results.forEach(function(r) { delete r._ts; });
    return results;
  }

  /**
   * Returns all contact log entries for a steward (max 100), sorted newest first.
   * @param {string} stewardEmail
   * @returns {Object[]} Array of contact records
   */
  function getStewardContactLog(stewardEmail) {
    if (!stewardEmail) return [];
    var sheet = _ensureContactLog();
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === sEmail) {
        var rawDate2 = data[i][4];
        results.push({ id: data[i][0], memberEmail: data[i][2], memberName: data[i][8] || '', type: data[i][3],
          date: rawDate2 instanceof Date ? _formatDate(rawDate2) : String(rawDate2),
          _ts: rawDate2 instanceof Date ? rawDate2.getTime() : 0,
          notes: data[i][5], duration: data[i][6] });
      }
    }
    // H-16: sort by numeric timestamp — string comparison of formatted dates is not chronological
    results.sort(function(a, b) { return (b._ts || 0) - (a._ts || 0); });
    results.forEach(function(r) { delete r._ts; });
    return results.slice(0, 100);
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Tasks (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns the _Steward_Tasks hidden sheet, creating it with headers if absent.
   * @returns {Sheet|null}
   */
  function _ensureStewardTasks() {
    var ss = _getSS();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.STEWARD_TASKS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.STEWARD_TASKS);
      sheet.getRange(1, 1, 1, 12).setValues([[
        'ID', 'Steward Email', 'Title', 'Description', 'Member Email',
        'Priority', 'Status', 'Due Date', 'Created', 'Completed',
        'Assignee Type', 'Assigned By'
      ]]);
      sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Creates a steward task with optional delegation to another steward (chief only).
   * @param {string} stewardEmail
   * @param {string} title
   * @param {string} description
   * @param {string} memberEmail - Related member email
   * @param {string} priority - 'high', 'medium', or 'low'
   * @param {string} dueDate
   * @param {string} assignToEmail - Delegate target (chief steward only)
   * @param {string} idemKey - Idempotency key to prevent duplicates
   * @returns {Object} { success: boolean, message: string, taskId?: string }
   */
  function createTask(stewardEmail, title, description, memberEmail, priority, dueDate, assignToEmail, idemKey) {
    if (idemKey) {
      var idemCache = CacheService.getScriptCache();
      if (idemCache.get('IDEM_' + idemKey)) return { duplicate: true, message: 'Duplicate request ignored' };
      idemCache.put('IDEM_' + idemKey, '1', 600); // 10-minute TTL to reduce duplicate mutation risk
    }
    if (!stewardEmail || !title) return { success: false, message: 'Missing fields.' };
    // If assignToEmail is provided and caller is chief steward, assign to that steward
    var ownerEmail = stewardEmail.toLowerCase().trim();
    if (assignToEmail) {
      var chiefEmail = _getChiefStewardEmail();
      if (chiefEmail && ownerEmail === chiefEmail) {
        ownerEmail = assignToEmail.toLowerCase().trim();
      }
    }
    var sheet = _ensureStewardTasks();
    var id = 'ST_' + Date.now().toString(36);
    sheet.appendRow([id, ownerEmail, escapeForFormula(title.substring(0, 200)), escapeForFormula((description || '').substring(0, 500)), (memberEmail || '').toLowerCase().trim(), priority || 'medium', 'open', dueDate || '', new Date(), '']);
    _invalidateSheetCache(SHEETS.STEWARD_TASKS);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('TASK_CREATED', { taskId: id, owner: ownerEmail, createdBy: stewardEmail });
    }
    return { success: true, message: 'Task created.', taskId: id };
  }

  /**
   * Returns steward tasks, optionally filtered by status, sorted by priority then due date.
   * @param {string} stewardEmail
   * @param {string} [statusFilter] - e.g. 'open', 'completed'
   * @returns {Object[]} Array of task records
   */
  function getTasks(stewardEmail, statusFilter) {
    // Use _getCachedSheetData for same-execution and cross-request caching.
    var cached = _getCachedSheetData(SHEETS.STEWARD_TASKS);
    var data = cached ? cached.data : null;
    if (!data || data.length <= 1) return [];
    var tasks = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() !== sEmail) continue;
      // Only steward tasks (col 11 blank or not 'member')
      if (String(data[i][10] || '').toLowerCase().trim() === 'member') continue;
      var status = String(data[i][6]).toLowerCase().trim();
      if (statusFilter && status !== statusFilter) continue;
      var dueDateRaw = data[i][7];
      var dueStr = dueDateRaw instanceof Date ? _formatDate(dueDateRaw) : String(dueDateRaw || '');
      var dueDays = null;
      if (dueDateRaw instanceof Date) {
        dueDays = Math.ceil((dueDateRaw.getTime() - Date.now()) / 86400000);
      }
      tasks.push({ id: data[i][0], title: data[i][2], description: data[i][3], memberEmail: data[i][4], priority: data[i][5], status: status, dueDate: dueStr, dueDays: dueDays, created: data[i][8] instanceof Date ? _formatDate(data[i][8]) : '' });
    }
    tasks.sort(function(a, b) {
      var pa = a.priority === 'high' ? 0 : a.priority === 'medium' ? 1 : 2;
      var pb = b.priority === 'high' ? 0 : b.priority === 'medium' ? 1 : 2;
      if (pa !== pb) return pa - pb;
      return (a.dueDays || 999) - (b.dueDays || 999);
    });
    return tasks;
  }

  /**
   * Updates a steward task's status, priority, title, or due date.
   * @param {string} stewardEmail
   * @param {string} taskId
   * @param {Object} updates - { status?, priority?, title?, dueDate? }
   * @returns {Object} { success: boolean, message?: string }
   */
  function updateTask(stewardEmail, taskId, updates) {
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return { success: false, message: 'No tasks.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId && String(data[i][1]).toLowerCase().trim() === stewardEmail.toLowerCase().trim()) {
        if (updates.status)   sheet.getRange(i + 1, 7).setValue(escapeForFormula(updates.status));
        if (updates.priority) sheet.getRange(i + 1, 6).setValue(escapeForFormula(updates.priority));
        if (updates.title)    sheet.getRange(i + 1, 3).setValue(escapeForFormula(updates.title.substring(0, 200)));
        if (updates.dueDate !== undefined) sheet.getRange(i + 1, 8).setValue(updates.dueDate || '');
        _invalidateSheetCache(SHEETS.STEWARD_TASKS);
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('TASK_UPDATED', { taskId: taskId, updatedBy: stewardEmail, fields: Object.keys(updates) });
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
  }

  /**
   * Marks a steward task as completed with a timestamp.
   * @param {string} stewardEmail
   * @param {string} taskId
   * @returns {Object} { success: boolean, message?: string }
   */
  function completeTask(stewardEmail, taskId) {
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return { success: false, message: 'No tasks.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId && String(data[i][1]).toLowerCase().trim() === stewardEmail.toLowerCase().trim()) {
        sheet.getRange(i + 1, 7).setValue('completed');
        sheet.getRange(i + 1, 10).setValue(new Date());
        _invalidateSheetCache(SHEETS.STEWARD_TASKS);
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('TASK_COMPLETED', { taskId: taskId, completedBy: stewardEmail });
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
  }

  // ═══════════════════════════════════════
  // Member Task Assignment (v4.17.0)
  // ═══════════════════════════════════════

  /**
   * Creates a task assigned to a member by a steward.
   * @param {string} stewardEmail - Assigning steward
   * @param {string} memberEmail - Member receiving the task
   * @param {string} title
   * @param {string} desc
   * @param {string} priority - 'high', 'medium', or 'low'
   * @param {string} dueDate
   * @returns {Object} { success: boolean, message: string, taskId?: string }
   */
  function createMemberTask(stewardEmail, memberEmail, title, desc, priority, dueDate) {
    if (!stewardEmail || !memberEmail || !title) return { success: false, message: 'Missing required fields.' };
    var sheet = _ensureStewardTasks();
    // Migrate headers if needed
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 12)).getValues()[0];
    var assigneeTypeCol = 11;  // position in 12-col Steward Tasks schema
    var assignedByCol   = 12;
    if (!headers[10] || String(headers[10]).trim() !== 'Assignee Type') {
      sheet.getRange(1, assigneeTypeCol).setValue('Assignee Type');
      sheet.getRange(1, assignedByCol).setValue('Assigned By');
    }
    var id = 'MT_' + Date.now().toString(36);
    sheet.appendRow([
      id, memberEmail.toLowerCase().trim(), escapeForFormula(title.substring(0, 200)),
      escapeForFormula((desc || '').substring(0, 500)), memberEmail.toLowerCase().trim(),
      priority || 'medium', 'open', dueDate || '', new Date(), '',
      'member', stewardEmail.toLowerCase().trim()
    ]);
    _invalidateSheetCache(SHEETS.STEWARD_TASKS);
    if (typeof logAuditEvent === 'function') logAuditEvent('MEMBER_TASK_CREATED', 'Task ' + id + ' assigned to ' + memberEmail + ' by ' + stewardEmail);
    // v4.51.1: Notify member of task assignment (BUG-12-004)
    if (typeof pushNotification === 'function') {
      try { pushNotification(memberEmail.toLowerCase().trim(), { title: 'New task assigned: ' + title.substring(0, 100), body: (desc || '').substring(0, 200), type: 'task' }); } catch (_) { /* non-critical */ }
    }
    return { success: true, message: 'Task assigned to member.', taskId: id };
  }

  /**
   * Returns tasks assigned to a member, optionally filtered by status.
   * @param {string} memberEmail
   * @param {string} [statusFilter] - 'open', 'completed', or 'not-completed'
   * @returns {Object[]} Array of task records
   */
  function getMemberTasks(memberEmail, statusFilter) {
    var cached = _getCachedSheetData(SHEETS.STEWARD_TASKS);
    var data = cached ? cached.data : null;
    if (!data || data.length <= 1) return [];
    var tasks = [];
    var mEmail = memberEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][10]).toLowerCase().trim() !== 'member') continue;
      if (String(data[i][4]).toLowerCase().trim() !== mEmail) continue;
      var status = String(data[i][6]).toLowerCase().trim();
      if (statusFilter === 'not-completed' && status === 'completed') continue;
      else if (statusFilter && statusFilter !== 'not-completed' && status !== statusFilter) continue;
      var dueDateRaw = data[i][7];
      var dueStr = dueDateRaw instanceof Date ? _formatDate(dueDateRaw) : String(dueDateRaw || '');
      var dueDays = null;
      if (dueDateRaw instanceof Date) dueDays = Math.ceil((dueDateRaw.getTime() - Date.now()) / 86400000);
      tasks.push({
        id: data[i][0], title: data[i][2], description: data[i][3],
        priority: data[i][5], status: status, dueDate: dueStr, dueDays: dueDays,
        assignedBy: String(data[i][11] || ''),
        created: data[i][8] instanceof Date ? _formatDate(data[i][8]) : ''
      });
    }
    tasks.sort(function(a, b) {
      var pa = a.priority === 'high' ? 0 : a.priority === 'medium' ? 1 : 2;
      var pb = b.priority === 'high' ? 0 : b.priority === 'medium' ? 1 : 2;
      if (pa !== pb) return pa - pb;
      return (a.dueDays || 999) - (b.dueDays || 999);
    });
    return tasks;
  }

  /**
   * Marks a member task as completed by the member themselves.
   * @param {string} memberEmail
   * @param {string} taskId
   * @returns {Object} { success: boolean, message?: string }
   */
  function completeMemberTask(memberEmail, taskId) {
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return { success: false, message: 'No tasks.' };
    var data = sheet.getDataRange().getValues();
    var mEmail = memberEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId && String(data[i][10]).toLowerCase().trim() === 'member' &&
          String(data[i][4]).toLowerCase().trim() === mEmail) {
        sheet.getRange(i + 1, 7).setValue('completed');
        sheet.getRange(i + 1, 10).setValue(new Date());
        _invalidateSheetCache(SHEETS.STEWARD_TASKS);
        if (typeof logAuditEvent === 'function') logAuditEvent('MEMBER_TASK_COMPLETED', 'Task ' + taskId + ' completed by ' + memberEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
  }

  /**
   * Allows the assigning steward to mark a member task complete on the member's behalf.
   * @param {string} stewardEmail
   * @param {string} taskId
   * @returns {Object} { success: boolean, message?: string }
   */
  function stewardCompleteMemberTask(stewardEmail, taskId) {
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return { success: false, message: 'No tasks.' };
    var data = sheet.getDataRange().getValues();
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === taskId &&
          String(data[i][10] || '').toLowerCase().trim() === 'member' &&
          String(data[i][11] || '').toLowerCase().trim() === sEmail) {
        sheet.getRange(i + 1, 7).setValue('completed');
        sheet.getRange(i + 1, 10).setValue(new Date());
        _invalidateSheetCache(SHEETS.STEWARD_TASKS);
        if (typeof logAuditEvent === 'function') logAuditEvent('MEMBER_TASK_COMPLETED_BY_STEWARD', 'Task ' + taskId + ' marked complete by steward ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found or not yours to complete.' };
  }

  /**
   * Returns all member tasks assigned by this steward.
   * @param {string} stewardEmail
   * @returns {Object[]} Array of member task records
   */
  function getStewardAssignedMemberTasks(stewardEmail) {
    var cached = _getCachedSheetData(SHEETS.STEWARD_TASKS);
    var data = cached ? cached.data : null;
    if (!data || data.length <= 1) return [];
    var tasks = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][10]).toLowerCase().trim() !== 'member') continue;
      if (String(data[i][11]).toLowerCase().trim() !== sEmail) continue;
      var status = String(data[i][6]).toLowerCase().trim();
      var dueDateRaw = data[i][7];
      var dueStr = dueDateRaw instanceof Date ? _formatDate(dueDateRaw) : String(dueDateRaw || '');
      var dueDays = null;
      if (dueDateRaw instanceof Date) dueDays = Math.ceil((dueDateRaw.getTime() - Date.now()) / 86400000);
      tasks.push({
        id: data[i][0], title: data[i][2], description: data[i][3],
        memberEmail: String(data[i][4] || ''), priority: data[i][5],
        status: status, dueDate: dueStr, dueDays: dueDays,
        created: data[i][8] instanceof Date ? _formatDate(data[i][8]) : ''
      });
    }
    return tasks;
  }

  // ═══════════════════════════════════════
  // Chief Steward Task View (v4.15.0)
  // ═══════════════════════════════════════

  /**
   * Returns the chief steward email from the config sheet.
   * @private
   */
  function _getChiefStewardEmail() {
    try {
      var email = '';
      if (typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.CHIEF_STEWARD_EMAIL) {
        var ss = _getSS();
        if (!ss) return null;
        var cfgSheet = ss.getSheetByName(SHEETS.CONFIG);
        if (cfgSheet) email = String(cfgSheet.getRange(3, CONFIG_COLS.CHIEF_STEWARD_EMAIL).getValue() || '').toLowerCase().trim();
      }
      if (!email && typeof COMMAND_CONFIG !== 'undefined') email = String(COMMAND_CONFIG.CHIEF_STEWARD_EMAIL || '').toLowerCase().trim();
      return email || null;
    } catch (_e) { return null; }
  }

  /**
   * Checks if the given email belongs to the chief steward.
   */
  function isChiefSteward(email) {
    if (!email) return false;
    var chief = _getChiefStewardEmail();
    return chief ? email.toLowerCase().trim() === chief : false;
  }

  /**
   * Gets ALL steward tasks (for chief steward overview).
   * Returns tasks from all stewards, grouped by steward.
   */
  function getChiefStewardTaskView(chiefEmail) {
    if (!isChiefSteward(chiefEmail)) return { authorized: false, tasks: [] };
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return { authorized: true, tasks: [] };
    var data = sheet.getDataRange().getValues();
    var tasks = [];
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][6]).toLowerCase().trim();
      var dueDateRaw = data[i][7];
      var dueStr = dueDateRaw instanceof Date ? _formatDate(dueDateRaw) : String(dueDateRaw || '');
      var dueDays = null;
      if (dueDateRaw instanceof Date) dueDays = Math.ceil((dueDateRaw.getTime() - Date.now()) / 86400000);
      tasks.push({ id: data[i][0], assignedTo: data[i][1], title: data[i][2], description: data[i][3], memberEmail: data[i][4], priority: data[i][5], status: status, dueDate: dueStr, dueDays: dueDays, created: data[i][8] instanceof Date ? _formatDate(data[i][8]) : '' });
    }
    return { authorized: true, tasks: tasks };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Member Stats (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns aggregate stats for a steward's members (by location, by dues status).
   * @param {string} stewardEmail
   * @returns {Object} { total: number, byLocation: Object, byDues: Object, scope: string }
   */
  function getStewardMemberStats(stewardEmail) {
    var members = getStewardMembers(stewardEmail);
    var scope = 'assigned';
    if (members.length === 0) {
      members = getAllMembers();
      scope = 'all';
    }
    if (members.length === 0) return { total: 0, byLocation: {}, byDues: {}, scope: scope };
    var byLocation = {};
    var byDues = {};
    members.forEach(function(m) {
      var loc = m.workLocation || 'Unknown';
      byLocation[loc] = (byLocation[loc] || 0) + 1;
      var dues = m.duesStatus || 'Unknown';
      byDues[dues] = (byDues[dues] || 0) + 1;
    });
    return { total: members.length, byLocation: byLocation, byDues: byDues, scope: scope };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Directory (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns a directory of all stewards with phone visibility based on caller role.
   * @param {boolean} callerIsSteward - Whether the caller is a steward (controls phone visibility)
   * @returns {Object[]} Array of steward contact entries
   */
  function getStewardDirectory(callerIsSteward) {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];
    var data = cached.data;
    var colMap = cached.colMap;
    var stewards = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      if (!rec.isSteward) continue;
      // Phone: stewards always see it; members only see it if the steward opted in
      var phoneToReturn = (callerIsSteward || rec.sharePhone) ? rec.phone : null;
      stewards.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        phone: phoneToReturn,
        unit: rec.unit,
      });
    }
    // Return unsorted — client-side renderList applies smart sort:
    // (1) same work location as current user, (2) in-office today, (3) alphabetical
    return stewards;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Grievance Stats (v4.12.0) — anonymized
  // ═══════════════════════════════════════

  /**
   * Returns anonymized org-wide grievance statistics including status, step, category, monthly trends, and deep analytics.
   * @returns {Object} Grievance stats with byStatus, byStep, byCategory, monthly, winRate, resolution metrics, etc.
   */
  function getGrievanceStats() {
    // Use combined active + archive data for complete statistics
    var cached = _getAllGrievanceData();
    if (!cached) return { available: false };
    var data = cached.data;
    if (data.length <= 1) return { available: false };
    var colMap = cached.colMap;
    var total = data.length - 1;

    var byStatus = {};
    var byStep = {};
    var byUnit = {};
    var byCategory = {};
    var monthly = {};
    var monthlyResolved = {};
    var openCount = 0, wonCount = 0, deniedCount = 0, settledCount = 0, withdrawnCount = 0;
    var overdueCount = 0, dueSoonCount = 0;
    for (var i = 1; i < data.length; i++) {
      try {
      var rec = _buildGrievanceRecord(data[i], colMap);
      var s = rec.status || 'unknown';
      byStatus[s] = (byStatus[s] || 0) + 1;
      var step = rec.step || 'Unknown';
      byStep[step] = (byStep[step] || 0) + 1;
      var u = rec.unit || 'Unknown';
      byUnit[u] = (byUnit[u] || 0) + 1;
      var cat = rec.issueCategory || 'Uncategorized';
      byCategory[cat] = (byCategory[cat] || 0) + 1;
      // Summary counts
      if (s === 'won') wonCount++;
      else if (s === 'denied') deniedCount++;
      else if (s === 'settled') settledCount++;
      else if (s === 'withdrawn') withdrawnCount++;
      else if (s !== 'resolved' && s !== 'closed') openCount++;
      // Deadline-based counts for org KPI cards
      if (s === 'overdue') {
        overdueCount++;
      } else if (rec.deadlineDays !== null && rec.deadlineDays >= 0 && rec.deadlineDays <= 7 &&
                 s !== 'resolved' && s !== 'withdrawn' && s !== 'closed' && s !== 'won' && s !== 'denied' && s !== 'settled') {
        dueSoonCount++;
      }
      // Monthly filings
      if (rec.filedTimestamp) {
        var d = new Date(rec.filedTimestamp);
        var key = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
        monthly[key] = (monthly[key] || 0) + 1;
      }
      // Monthly resolutions
      if (rec.closedTimestamp) {
        var dc = new Date(rec.closedTimestamp);
        var rKey = dc.getFullYear() + '-' + ('0' + (dc.getMonth() + 1)).slice(-2);
        monthlyResolved[rKey] = (monthlyResolved[rKey] || 0) + 1;
      }
      } catch (rowErr) {
        log_('getGrievanceStats', 'skipped row ' + i + ' — ' + rowErr.message);
      }
    }
    // Merge month keys from both maps, sort, take last 12
    var allMonths = {};
    Object.keys(monthly).forEach(function(k) { allMonths[k] = true; });
    Object.keys(monthlyResolved).forEach(function(k) { allMonths[k] = true; });
    var sortedMonths = Object.keys(allMonths).sort().slice(-12);
    var monthlyArr = sortedMonths.map(function(k) { return { month: k, count: monthly[k] || 0 }; });
    var monthlyResolvedArr = sortedMonths.map(function(k) { return { month: k, count: monthlyResolved[k] || 0 }; });
    // Win rate: Won / (Won + Denied + Settled + Withdrawn + Resolved + Closed)
    var closedTotal = wonCount + deniedCount + settledCount + withdrawnCount;
    // Also count 'resolved' and 'closed' from byStatus
    var resolvedCount = byStatus['resolved'] || 0;
    var closedStatusCount = byStatus['closed'] || 0;
    closedTotal += resolvedCount + closedStatusCount;
    var winRate = closedTotal > 0 ? Math.round(((wonCount + settledCount) / closedTotal) * 100) : null;
    var favorableCount = wonCount + settledCount;

    // Extended metrics pass — resolution time, deadline compliance, steward performance,
    // recurrence, step escalation, contract articles, response SLA, YoY (v4.31.5)
    var resolutionDays = [];
    var deadlineMet = 0;
    var deadlineTotal = 0;
    var stewardResolution = {};   // steward -> [days]
    var memberGrievanceCounts = {}; // email -> count
    var stepProgression = { s1: 0, s2: 0, s3: 0, s1to2: 0, s2to3: 0 };
    var articleCounts = {};       // article -> count
    var step1SlaMetCount = 0;
    var step1SlaTotal = 0;
    var yearBuckets = {};         // year -> { filed, closed, won, settled, denied }

    for (var ri = 1; ri < data.length; ri++) {
      try {
        var rRec = _buildGrievanceRecord(data[ri], colMap);
        var rStatus = rRec.status;
        var CLOSED = ['won', 'denied', 'settled', 'resolved', 'closed', 'withdrawn'];

        // Resolution time
        if (rRec.filedTimestamp && rRec.closedTimestamp && rRec.closedTimestamp > rRec.filedTimestamp) {
          var days = Math.round((rRec.closedTimestamp - rRec.filedTimestamp) / 86400000);
          if (days >= 0 && days < 3650) {
            resolutionDays.push(days);
            // Per-steward resolution
            var stw = rRec.steward || 'unassigned';
            if (!stewardResolution[stw]) stewardResolution[stw] = [];
            stewardResolution[stw].push(days);
          }
        }

        // Deadline compliance
        if (rRec.deadlineDays !== null && CLOSED.indexOf(rStatus) >= 0) {
          deadlineTotal++;
          if (rStatus !== 'overdue') deadlineMet++;
        }

        // Recurrence — count grievances per member email
        if (rRec.memberEmail) {
          memberGrievanceCounts[rRec.memberEmail] = (memberGrievanceCounts[rRec.memberEmail] || 0) + 1;
        }

        // Step escalation
        var stepVal = String(rRec.step).trim();
        if (stepVal === '1' || stepVal.toLowerCase() === 'step 1' || stepVal.toLowerCase() === 'i') stepProgression.s1++;
        else if (stepVal === '2' || stepVal.toLowerCase() === 'step 2' || stepVal.toLowerCase() === 'ii') { stepProgression.s2++; stepProgression.s1to2++; }
        else if (stepVal === '3' || stepVal.toLowerCase() === 'step 3' || stepVal.toLowerCase() === 'iii' || stepVal.toLowerCase().indexOf('arb') >= 0) { stepProgression.s3++; stepProgression.s2to3++; }

        // Contract articles violated
        var articles = String(_getVal(data[ri], colMap, HEADERS.grievanceArticles, '')).trim();
        if (articles) {
          articles.split(/[,;]+/).forEach(function(a) {
            var art = a.trim();
            if (art) articleCounts[art] = (articleCounts[art] || 0) + 1;
          });
        }

        // Step 1 response SLA — was Step I Rcvd before Step I Due?
        var s1Due = _getVal(data[ri], colMap, HEADERS.grievanceStep1Due, null);
        var s1Rcvd = _getVal(data[ri], colMap, HEADERS.grievanceStep1Rcvd, null);
        if (s1Due && s1Rcvd) {
          var dueDate = s1Due instanceof Date ? s1Due : new Date(s1Due);
          var rcvdDate = s1Rcvd instanceof Date ? s1Rcvd : new Date(s1Rcvd);
          if (!isNaN(dueDate.getTime()) && !isNaN(rcvdDate.getTime())) {
            step1SlaTotal++;
            if (rcvdDate <= dueDate) step1SlaMetCount++;
          }
        }

        // Year-over-year
        if (rRec.filedTimestamp) {
          var fYear = new Date(rRec.filedTimestamp).getFullYear();
          if (!yearBuckets[fYear]) yearBuckets[fYear] = { filed: 0, closed: 0, won: 0, settled: 0, denied: 0 };
          yearBuckets[fYear].filed++;
        }
        if (rRec.closedTimestamp) {
          var cYear = new Date(rRec.closedTimestamp).getFullYear();
          if (!yearBuckets[cYear]) yearBuckets[cYear] = { filed: 0, closed: 0, won: 0, settled: 0, denied: 0 };
          yearBuckets[cYear].closed++;
          if (rStatus === 'won') yearBuckets[cYear].won++;
          else if (rStatus === 'settled') yearBuckets[cYear].settled++;
          else if (rStatus === 'denied') yearBuckets[cYear].denied++;
        }
      } catch (_re) { /* skip */ }
    }

    // Compute derived stats
    var avgResolutionDays = resolutionDays.length > 0 ? Math.round(resolutionDays.reduce(function(a, b) { return a + b; }, 0) / resolutionDays.length) : null;
    var medianResolutionDays = null;
    if (resolutionDays.length > 0) {
      resolutionDays.sort(function(a, b) { return a - b; });
      var mid = Math.floor(resolutionDays.length / 2);
      medianResolutionDays = resolutionDays.length % 2 === 0 ? Math.round((resolutionDays[mid - 1] + resolutionDays[mid]) / 2) : resolutionDays[mid];
    }
    var deadlineComplianceRate = deadlineTotal > 0 ? Math.round((deadlineMet / deadlineTotal) * 100) : null;

    // Steward resolution averages (anonymized — numbered not named)
    var stewardAvgResolution = [];
    for (var sw in stewardResolution) {
      var sDays = stewardResolution[sw];
      var sAvg = Math.round(sDays.reduce(function(a, b) { return a + b; }, 0) / sDays.length);
      var sLabel = sw;
      if (sw && sw !== 'unassigned') {
        var sUser = findUserByEmail(sw);
        sLabel = (sUser && sUser.name) ? sUser.name : sw;
      }
      stewardAvgResolution.push({ label: sLabel, avgDays: sAvg, cases: sDays.length });
    }
    stewardAvgResolution.sort(function(a, b) { return a.avgDays - b.avgDays; });

    // Recurrence rate
    var totalMembers = Object.keys(memberGrievanceCounts).length;
    var repeatGrievants = 0;
    for (var me in memberGrievanceCounts) {
      if (memberGrievanceCounts[me] >= 2) repeatGrievants++;
    }
    var recurrenceRate = totalMembers > 0 ? Math.round((repeatGrievants / totalMembers) * 100) : 0;

    // Step escalation rates
    var totalFiled = stepProgression.s1 + stepProgression.s2 + stepProgression.s3;
    var escalation1to2 = totalFiled > 0 ? Math.round((stepProgression.s1to2 / totalFiled) * 100) : null;
    var escalation2to3 = stepProgression.s1to2 > 0 ? Math.round((stepProgression.s2to3 / Math.max(stepProgression.s1to2, 1)) * 100) : null;

    // Top violated articles (sorted desc, top 10)
    var articleList = [];
    for (var art in articleCounts) articleList.push({ article: art, count: articleCounts[art] });
    articleList.sort(function(a, b) { return b.count - a.count; });
    articleList = articleList.slice(0, 15);

    // Step 1 response SLA rate
    var step1SlaRate = step1SlaTotal > 0 ? Math.round((step1SlaMetCount / step1SlaTotal) * 100) : null;

    // Year-over-year — sorted array
    var yoyData = [];
    var yoyYears = Object.keys(yearBuckets).sort();
    for (var yi = 0; yi < yoyYears.length; yi++) {
      var yb = yearBuckets[yoyYears[yi]];
      var yClosed = yb.closed || 1;
      yoyData.push({
        year: yoyYears[yi],
        filed: yb.filed,
        closed: yb.closed,
        winRate: yClosed > 0 ? Math.round(((yb.won + yb.settled) / yClosed) * 100) : 0
      });
    }

    return {
      available: true, total: total,
      byStatus: byStatus, byStep: byStep, byUnit: byUnit, byCategory: byCategory,
      monthly: monthlyArr, monthlyResolved: monthlyResolvedArr,
      openCount: openCount, wonCount: wonCount, deniedCount: deniedCount, settledCount: settledCount, withdrawnCount: withdrawnCount,
      overdueCount: overdueCount, dueSoonCount: dueSoonCount,
      // Outcome & resolution metrics (v4.31.4)
      winRate: winRate, favorableCount: favorableCount, closedTotal: closedTotal,
      avgResolutionDays: avgResolutionDays, medianResolutionDays: medianResolutionDays,
      resolvedCount: resolutionDays.length, deadlineComplianceRate: deadlineComplianceRate,
      // Deep analytics (v4.31.5)
      stewardAvgResolution: stewardAvgResolution,
      recurrenceRate: recurrenceRate, repeatGrievants: repeatGrievants, uniqueGrievants: totalMembers,
      escalation1to2: escalation1to2, escalation2to3: escalation2to3, stepProgression: stepProgression,
      articleHeatmap: articleList,
      step1SlaRate: step1SlaRate,
      yoyData: yoyData,
    };
  }

  /**
   * Returns units with 3+ grievances as hotspots, sorted by count descending.
   * @returns {Object[]} Array of { location: string, count: number }
   */
  function getGrievanceHotSpots() {
    // Use combined active + archive for complete hotspot analysis
    var cached = _getAllGrievanceData();
    if (!cached) return [];
    var data = cached.data;
    if (data.length <= 1) return [];
    var colMap = cached.colMap;
    var unitCol = _findColumn(colMap, HEADERS.grievanceUnit);
    if (unitCol === -1) return [];

    var counts = {};
    for (var i = 1; i < data.length; i++) {
      var unit = String(data[i][unitCol]).trim() || 'Unknown';
      counts[unit] = (counts[unit] || 0) + 1;
    }
    var spots = [];
    for (var u in counts) {
      if (counts[u] >= 3) spots.push({ location: u, count: counts[u] });
    }
    spots.sort(function(a, b) { return b.count - a.count; });
    return spots;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Membership Stats (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns anonymized membership statistics (by unit, location, dues, tenure, and hire trends).
   * @returns {Object} { available: boolean, total: number, byUnit, byLocation, byDues, byTenure, ... }
   */
  function getMembershipStats() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return { available: false };
    var data = cached.data;
    if (data.length <= 1) return { available: false };
    var total = data.length - 1;
    if (total < 20) return { available: false, count: total, threshold: 20 };

    var colMap = cached.colMap;
    var byUnit = {};
    var byLocation = {};
    var byDues = {};
    var byHireMonth = {};
    var newMembersLast90 = 0;
    var now = new Date();
    var ninetyDaysAgo = now.getTime() - 90 * 24 * 60 * 60 * 1000;
    var twelveMonthsAgo = new Date(now.getFullYear() - 1, now.getMonth(), 1);
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      var unit = rec.unit || 'Unknown';
      byUnit[unit] = (byUnit[unit] || 0) + 1;
      var loc = rec.workLocation || 'Unknown';
      byLocation[loc] = (byLocation[loc] || 0) + 1;
      var dues = rec.duesStatus || 'Unknown';
      byDues[dues] = (byDues[dues] || 0) + 1;
      // Hire date tracking
      if (rec.hireDate) {
        var hd = new Date(rec.hireDate);
        if (!isNaN(hd.getTime())) {
          if (hd.getTime() >= ninetyDaysAgo) newMembersLast90++;
          if (hd.getTime() >= twelveMonthsAgo.getTime()) {
            var hKey = hd.getFullYear() + '-' + ('0' + (hd.getMonth() + 1)).slice(-2);
            byHireMonth[hKey] = (byHireMonth[hKey] || 0) + 1;
          }
        }
      }
    }
    // Tenure distribution (v4.31.4)
    var byTenure = { '< 1 year': 0, '1-3 years': 0, '3-5 years': 0, '5-10 years': 0, '10+ years': 0 };
    for (var ti = 1; ti < data.length; ti++) {
      var tRec = _buildUserRecord(data[ti], colMap);
      if (tRec.hireDate) {
        var thd = new Date(tRec.hireDate);
        if (!isNaN(thd.getTime())) {
          var yearsEmployed = (now.getTime() - thd.getTime()) / (365.25 * 86400000);
          if (yearsEmployed < 1) byTenure['< 1 year']++;
          else if (yearsEmployed < 3) byTenure['1-3 years']++;
          else if (yearsEmployed < 5) byTenure['3-5 years']++;
          else if (yearsEmployed < 10) byTenure['5-10 years']++;
          else byTenure['10+ years']++;
        }
      }
    }

    return { available: true, total: total, byUnit: byUnit, byLocation: byLocation, byDues: byDues, newMembersLast90: newMembersLast90, byHireMonth: byHireMonth, byTenure: byTenure };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Upcoming Events via CalendarApp (v4.12.0)
  // ═══════════════════════════════════════

  /**
   * Returns upcoming calendar events (next 90 days) with timeline fallback.
   * @param {number} [limit=10] - Max events to return
   * @returns {Object[]|Object} Array of event objects or status object
   */
  function getUpcomingEvents(limit) {
    limit = limit || 10;
    try {
      var config = ConfigReader.getConfig();
      if (!config.calendarId) {
        // No calendar configured — try _Timeline_Events fallback
        var fallback = _getTimelineEvents(limit);
        if (fallback.length > 0) return fallback;
        return { _notConfigured: true, events: [] };
      }
      var cache = CacheService.getScriptCache();
      var cacheKey = 'events_' + config.calendarId;
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);

      var cal = null;
      try { cal = CalendarApp.getCalendarById(config.calendarId); } catch (_e) { cal = null; }
      if (!cal) {
        // Calendar not accessible — try _Timeline_Events fallback
        var fallback2 = _getTimelineEvents(limit);
        if (fallback2.length > 0) return fallback2;
        return { _calNotFound: true };
      }
      var now = new Date();
      var future = new Date(now.getTime() + 90 * 86400000);
      var events = cal.getEvents(now, future);
      var result = [];
      for (var i = 0; i < Math.min(events.length, limit); i++) {
        var ev = events[i];
        result.push({
          title: ev.getTitle(),
          startTime: ev.getStartTime().toISOString(),
          endTime: ev.getEndTime().toISOString(),
          location: ev.getLocation() || '',
          description: (ev.getDescription() || '').substring(0, 300),
        });
      }
      // If calendar returned no events, try timeline fallback
      if (result.length === 0) {
        var fallback3 = _getTimelineEvents(limit);
        if (fallback3.length > 0) result = fallback3;
      }
      if (result.length > 0) cache.put(cacheKey, JSON.stringify(result), 900);
      return result;
    } catch (e) {
      log_('getUpcomingEvents error', e.message);
      // Last resort: try timeline fallback
      try { var fb = _getTimelineEvents(limit); if (fb.length > 0) return fb; } catch (_e2) { log_('_e2', (_e2.message || _e2)); }
      return [];
    }
  }

  /**
   * Reads upcoming events from _Timeline_Events sheet (seeded by DevTools).
   * Returns array of {title, startTime, endTime, location, description} sorted by date.
   */
  function _getTimelineEvents(limit) {
    var ss = _getSS();
    if (!ss) return [];
    var tlSheet = ss.getSheetByName(SHEETS.TIMELINE_EVENTS);
    if (!tlSheet || tlSheet.getLastRow() <= 1) return [];
    var data = tlSheet.getDataRange().getValues();
    var now = new Date();
    var results = [];
    // Schema: ID(0), Title(1), Date(2), Description(3), Type(4), ..., EndTime(8)
    for (var i = 1; i < data.length; i++) {
      var evtDate = data[i][2];
      if (!(evtDate instanceof Date)) continue;
      if (evtDate < now) continue; // skip past events
      var endStr = data[i][8] ? String(data[i][8]) : '';
      var endDate = endStr ? new Date(endStr) : new Date(evtDate.getTime() + 3600000);
      results.push({
        title: String(data[i][1] || ''),
        startTime: evtDate.toISOString(),
        endTime: endDate instanceof Date && !isNaN(endDate) ? endDate.toISOString() : new Date(evtDate.getTime() + 3600000).toISOString(),
        location: '',
        description: String(data[i][3] || '').substring(0, 300),
      });
    }
    results.sort(function(a, b) { return new Date(a.startTime) - new Date(b.startTime); });
    return results.slice(0, limit || 10);
  }

  // ─── Execution Time Guard ────────────────────────────────
  var _execStart = Date.now();
  /**
   * Checks elapsed execution time. Throws if approaching GAS web app
   * timeout (30s for doGet/doPost, 6min for server calls).
   */
  function _checkExecTime(label) {
    var elapsed = Date.now() - _execStart;
    if (elapsed > 25000) {
      log_('SLOW_EXEC', label + ' at ' + elapsed + 'ms');
      throw new Error('Request is taking too long. Please try again.');
    }
  }

  // ═══════════════════════════════════════
  // SHEET DATA CACHE — avoids redundant full-sheet reads within a single
  // server execution (multiple methods reading the same sheet).
  // Uses CacheService with a short TTL for cross-request caching.
  // ═══════════════════════════════════════

  var _sheetDataCache = {};
  var SHEET_CACHE_TTL = 120; // 2 minutes
  var _sheetHealthLog = {}; // { sheetName: rowCount } — for health reporting

  /**
   * Stores a large JSON string in CacheService using chunked keys.
   * Each chunk is max 90KB to stay under the 100KB/key limit.
   */
  function _putChunkedCache(cache, key, json, ttl) {
    var CHUNK_SIZE = 90000;
    if (json.length <= CHUNK_SIZE) {
      cache.put(key, json, ttl);
      cache.put(key + '_n', '1', ttl);
      return;
    }
    var numChunks = Math.ceil(json.length / CHUNK_SIZE);
    var pairs = {};
    pairs[key + '_n'] = String(numChunks);
    for (var i = 0; i < numChunks; i++) {
      pairs[key + '_' + i] = json.substr(i * CHUNK_SIZE, CHUNK_SIZE);
    }
    cache.putAll(pairs, ttl);
  }

  /**
   * Retrieves a chunked JSON string from CacheService.
   * Returns null on miss or partial eviction.
   */
  function _getChunkedCache(cache, key) {
    var n = cache.get(key + '_n');
    if (!n) return null;
    n = parseInt(n, 10);
    if (n === 1) return cache.get(key);
    var keys = [key + '_n'];
    var i;
    for (i = 0; i < n; i++) keys.push(key + '_' + i);
    var all = cache.getAll(keys);
    var parts = [];
    for (i = 0; i < n; i++) {
      var chunk = all[key + '_' + i];
      if (!chunk) return null; // partial eviction
      parts.push(chunk);
    }
    return parts.join('');
  }

  /**
   * Returns cached sheet data (headers + rows) or reads from the sheet
   * and caches the result. Uses in-memory cache for same-execution reuse
   * and CacheService for cross-request reuse.
   * @param {string} sheetName
   * @returns {{ data: Array[], colMap: Object }|null}
   */
  function _getCachedSheetData(sheetName) {
    // In-memory cache (same Apps Script execution)
    if (_sheetDataCache[sheetName]) return _sheetDataCache[sheetName];

    // CacheService cache (cross-request, short TTL)
    var cacheKey = 'SD_' + sheetName.replace(/\s/g, '_');
    try {
      var cache = CacheService.getScriptCache();
      var cached = _getChunkedCache(cache, cacheKey);
      if (cached) {
        var parsed = JSON.parse(cached);
        // Log cache age for observability
        if (parsed._cachedAt) {
          log_('Sheet cache hit', sheetName + ' (age: ' + (Date.now() - parsed._cachedAt) + 'ms)');
        }
        // Restore Date objects from serialized __d markers
        parsed.data = parsed.data.map(function(row) {
          return row.map(function(cell) {
            return (cell && typeof cell === 'object' && cell.__d) ? new Date(cell.__d) : cell;
          });
        });
        parsed.colMap = _buildColumnMap(parsed.data[0]);
        _sheetDataCache[sheetName] = parsed;
        return parsed;
      }
    } catch (_e) { log_('_getCachedSheetData', 'Error reading cache: ' + (_e.message || _e)); }

    var sheet = _getSheet(sheetName);
    if (!sheet) return null;

    var data = executeWithRetry(function() { return sheet.getDataRange().getValues(); }, { maxRetries: 2, baseDelay: 500 });
    var colMap = _buildColumnMap(data[0]);
    var result = { data: data, colMap: colMap };

    // ─── Health Monitor — log scale warnings ───
    var rowCount = data.length - 1; // exclude header
    _sheetHealthLog[sheetName] = rowCount;
    if (typeof SCALE_THRESHOLDS !== 'undefined') {
      if (rowCount >= SCALE_THRESHOLDS.CRITICAL_ROWS) {
        log_('SCALE_CRITICAL', sheetName + ' has ' + rowCount + ' rows — manual intervention recommended');
      } else if (rowCount >= SCALE_THRESHOLDS.THROTTLE_ROWS) {
        log_('SCALE_THROTTLE', sheetName + ' has ' + rowCount + ' rows — paginated mode active');
      } else if (rowCount >= SCALE_THRESHOLDS.WARN_ROWS) {
        log_('SCALE_WARN', sheetName + ' has ' + rowCount + ' rows');
      }
    }

    // Store in memory
    _sheetDataCache[sheetName] = result;

    // Smart TTL: extend cache lifetime for large sheets to reduce re-reads
    var smartTTL = SHEET_CACHE_TTL;
    if (typeof SCALE_THRESHOLDS !== 'undefined') {
      if (rowCount >= SCALE_THRESHOLDS.THROTTLE_ROWS) {
        smartTTL = 300; // 5 min — large sheet, expensive to re-read
      } else if (rowCount >= SCALE_THRESHOLDS.WARN_ROWS) {
        smartTTL = 180; // 3 min
      }
    }

    // Store in CacheService (serialize dates as ISO strings for JSON)
    try {
      var writeCache = CacheService.getScriptCache();
      var serializable = data.map(function(row) {
        return row.map(function(cell) {
          return cell instanceof Date ? { __d: cell.toISOString() } : cell;
        });
      });
      var json = JSON.stringify({ data: serializable, _cachedAt: Date.now() });
      _putChunkedCache(writeCache, cacheKey, json, smartTTL);
    } catch (_e) { log_('_getCachedSheetData', 'Error writing cache: ' + (_e.message || _e)); }

    return result;
  }

  /**
   * Invalidates the sheet data cache for a specific sheet.
   * Call after writes to ensure fresh reads.
   */
  function _invalidateSheetCache(sheetName) {
    delete _sheetDataCache[sheetName];
    // H1: Invalidate email index so stale lookups don't persist in batch ops
    if (sheetName === MEMBER_SHEET) _emailIndex = null;
    try {
      var cacheKey = 'SD_' + sheetName.replace(/\s/g, '_');
      var cache = CacheService.getScriptCache();
      var n = cache.get(cacheKey + '_n');
      cache.remove(cacheKey);
      cache.remove(cacheKey + '_n');
      if (n) {
        var num = parseInt(n, 10);
        for (var i = 0; i < num; i++) cache.remove(cacheKey + '_' + i);
      }
    } catch (_e) { log_('_invalidateSheetCache', 'Error: ' + (_e.message || _e)); }
  }

  // ═══════════════════════════════════════
  // GRIEVANCE ARCHIVE HELPERS (v4.30.0)
  // ═══════════════════════════════════════

  /**
   * Returns cached archive grievance data with a longer TTL (10 min).
   * Archive contains only closed/resolved cases — data is stable.
   * @returns {Object|null} { data: Array[], colMap: Object } or null
   */
  function _getCachedArchiveData() {
    var archiveName = SHEETS.GRIEVANCE_ARCHIVE;
    // Check in-memory first
    if (_sheetDataCache[archiveName]) return _sheetDataCache[archiveName];

    var sheet = _getSheet(archiveName);
    if (!sheet || sheet.getLastRow() <= 1) return null;

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
    var result = { data: data, colMap: colMap };
    _sheetDataCache[archiveName] = result;

    // Cache with 10-min TTL (archive data is stable)
    try {
      var cache = CacheService.getScriptCache();
      var cacheKey = 'SD_' + archiveName.replace(/\s/g, '_');
      var serialized = JSON.stringify({ data: data, _cachedAt: Date.now() });
      _putChunkedCache(cache, cacheKey, serialized, 600);
    } catch (_e) { log_('_getCachedArchiveData cache write', (_e.message || _e)); }

    return result;
  }

  /**
   * Returns combined active + archive grievance data for functions
   * that need a complete view (stats, hotspots, history).
   * @returns {Object|null} { data: Array[], colMap: Object }
   */
  function _getAllGrievanceData() {
    var active = _getCachedSheetData(GRIEVANCE_SHEET);
    var archive = _getCachedArchiveData();

    // If no archive, just return active data
    if (!archive || !archive.data || archive.data.length <= 1) return active;
    if (!active || !active.data || active.data.length <= 1) return archive;

    // Validate column alignment: archive headers must match active headers.
    // If columns were added/reordered after archive creation, skip archive
    // to avoid misaligned data (stats would be inaccurate for archived rows).
    var activeHeaders = active.data[0];
    var archiveHeaders = archive.data[0];
    if (activeHeaders.length !== archiveHeaders.length) {
      log_('_getAllGrievanceData', 'column count mismatch (active=' + activeHeaders.length + ', archive=' + archiveHeaders.length + ') — using active only');
      return active;
    }
    for (var h = 0; h < activeHeaders.length; h++) {
      if (String(activeHeaders[h]).trim() !== String(archiveHeaders[h]).trim()) {
        log_('_getAllGrievanceData', 'header mismatch at col ' + (h + 1) + ' ("' + activeHeaders[h] + '" vs "' + archiveHeaders[h] + '") — using active only');
        return active;
      }
    }

    // Merge: skip header row from archive, concatenate data rows
    var merged = active.data.concat(archive.data.slice(1));
    return { data: merged, colMap: active.colMap };
  }

  // ═══════════════════════════════════════
  // BATCH DATA — single round-trip for SPA init
  // Aggregates all data needed for the initial view in one call.
  // ═══════════════════════════════════════

  /**
   * Returns all data needed for the initial page render in one round trip.
   * @param {string} email - User email
   * @param {string} role - 'member' or 'steward'
   * @returns {Object} Batch data payload
   */
  function getBatchData(email, role) {
    _checkExecTime('getBatchData');
    if (!email) return {};
    email = String(email).trim().toLowerCase();

    if (role === 'steward') {
      return _getStewardBatchData(email);
    }
    return _getMemberBatchData(email);
  }

  /**
   * Aggregates all member-view data in one call: grievances, history, steward info, survey, events, tasks.
   * @param {string} email - Member email
   * @returns {Object} Batch payload for member view init
   */
  function _getMemberBatchData(email) {
    // Pre-warm cache for sheets we'll read multiple times
    try { _getCachedSheetData(GRIEVANCE_SHEET); } catch (_e) { log_('_getMemberBatchData', 'Error pre-warming grievance cache: ' + (_e.message || _e)); }
    try { _getCachedSheetData(MEMBER_SHEET); } catch (_e) { log_('_getMemberBatchData', 'Error pre-warming member cache: ' + (_e.message || _e)); }

    _checkExecTime('_getMemberBatchData:start');
    var grievances = [];
    try { grievances = getMemberGrievances(email); } catch (_e) { handleError(_e, '_getMemberBatchData.getMemberGrievances', ERROR_LEVEL.WARNING); }
    _checkExecTime('_getMemberBatchData:grievances');
    var history = { success: false, history: [] };
    try { history = getMemberGrievanceHistory(email); } catch (_e) { handleError(_e, '_getMemberBatchData.getMemberGrievanceHistory', ERROR_LEVEL.WARNING); }
    _checkExecTime('_getMemberBatchData:history');
    var stewardInfo = null;
    try { stewardInfo = getAssignedStewardInfo(email); } catch (_e) { handleError(_e, '_getMemberBatchData.getAssignedStewardInfo', ERROR_LEVEL.WARNING); }
    _checkExecTime('_getMemberBatchData:stewardInfo');
    var surveyStatus = null;
    try { surveyStatus = getMemberSurveyStatus(email); } catch (_e) { handleError(_e, '_getMemberBatchData.getMemberSurveyStatus', ERROR_LEVEL.WARNING); }
    _checkExecTime('_getMemberBatchData:surveyStatus');
    var events = [];
    try { events = getUpcomingEvents(5); } catch (_e) { handleError(_e, '_getMemberBatchData.getUpcomingEvents', ERROR_LEVEL.WARNING); }
    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, 'member');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { log_('_getMemberBatchData', 'Error getting notification count: ' + (_e.message || _e)); }

    var memberTaskCount = 0;
    try {
      var openTasks = getMemberTasks(email, 'not-completed');
      memberTaskCount = openTasks.length;
    } catch (_e) { log_('_getMemberBatchData', 'Error getting member tasks: ' + (_e.message || _e)); }

    return {
      grievances: grievances,
      history: history,
      stewardInfo: stewardInfo,
      surveyStatus: surveyStatus,
      events: events,
      notificationCount: notifCount,
      memberTaskCount: memberTaskCount,
      activeMeeting: _getActiveMeetingForCheckIn(email),
    };
  }

  /**
   * Aggregates all steward-view data in one call: cases, KPIs, members, tasks, badges, Q&A.
   * @param {string} email - Steward email
   * @returns {Object} Batch payload for steward view init
   */
  function _getStewardBatchData(email) {
    // Pre-warm cache for sheets we'll read multiple times
    try { _getCachedSheetData(GRIEVANCE_SHEET); } catch (_e) { log_('_getStewardBatchData', 'Error pre-warming grievance cache: ' + (_e.message || _e)); }
    try { _getCachedSheetData(MEMBER_SHEET); } catch (_e) { log_('_getStewardBatchData', 'Error pre-warming member cache: ' + (_e.message || _e)); }
    try { _getCachedSheetData(SHEETS.STEWARD_TASKS); } catch (_e) { log_('_getStewardBatchData', 'Error pre-warming tasks cache: ' + (_e.message || _e)); }

    _checkExecTime('_getStewardBatchData:start');
    // Read cases once and compute KPIs from same data (avoids double sheet read)
    var cases = [];
    try {
      cases = getStewardCases(email);
    } catch (_e) {
      handleError(_e, '_getStewardBatchData.getStewardCases', ERROR_LEVEL.WARNING);
    }
    var kpis = _computeKPIsFromCases(cases);

    _checkExecTime('_getStewardBatchData:cases');
    // Member counts — Member Directory already cached from getStewardCases call above.
    // getStewardMembers falls back to getAllMembers() when no members are assigned,
    // so memberCount is always meaningful.
    var memberCount = 0;
    try {
      memberCount = getStewardMembers(email).length;
      if (memberCount === 0) memberCount = getAllMembers().length;
    } catch (_e) { log_('_getStewardBatchData', 'Error getting member count: ' + (_e.message || _e)); }

    _checkExecTime('_getStewardBatchData:members');
    // Task counts — open tasks only; derive overdue from dueDays < 0.
    // No extra sheet read if _Steward_Tasks was already read this execution.
    var taskCount = 0;
    var overdueTaskCount = 0;
    try {
      var openTasks = getTasks(email, 'open');
      taskCount = openTasks.length;
      for (var t = 0; t < openTasks.length; t++) {
        if (openTasks[t].dueDays !== null && openTasks[t].dueDays < 0) overdueTaskCount++;
      }
    } catch (_e) { log_('_getStewardBatchData', 'Error getting tasks: ' + (_e.message || _e)); }

    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, 'steward');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { log_('_getStewardBatchData', 'Error getting notification count: ' + (_e.message || _e)); }

    var qaUnansweredCount = 0;
    try {
      if (typeof QAForum !== 'undefined') {
        // Use lightweight count — avoids building 999 full question objects
        qaUnansweredCount = QAForum.getUnansweredCount();
      }
    } catch (_e) { log_('_getStewardBatchData', 'Error getting QA count: ' + (_e.message || _e)); }

    return {
      cases: cases,
      kpis: kpis,
      memberCount: memberCount,
      taskCount: taskCount,
      overdueTaskCount: overdueTaskCount,
      notificationCount: notifCount,
      qaUnansweredCount: qaUnansweredCount,
      activeMeeting: _getActiveMeetingForCheckIn(email),
    };
  }

  /**
   * Computes KPIs from an already-fetched cases array (no extra sheet read).
   */
  function _computeKPIsFromCases(cases) {
    var total = cases.length;
    var overdue = 0, dueSoon = 0, resolved = 0, active = 0;
    for (var i = 0; i < cases.length; i++) {
      var status = String(cases[i].status).toLowerCase();
      if (_closedStatusesLower.indexOf(status) !== -1) {
        resolved++;
      } else {
        active++;
        if (status === 'overdue') overdue++;
        if (cases[i].deadlineDays !== null && cases[i].deadlineDays <= 7 && cases[i].deadlineDays >= 0) dueSoon++;
      }
    }
    return { totalCases: total, activeCases: active, overdue: overdue, dueSoon: dueSoon, resolved: resolved };
  }

  /**
   * S2: Returns all badge counts in a single call (replaces 3 serial calls).
   * Used by client-side _refreshNavBadges to update notification, task, and Q&A badges.
   * @param {string} email - User email
   * @param {string} role - 'steward' or 'member'
   * @returns {Object} { notificationCount, taskCount, overdueTaskCount, qaUnansweredCount }
   */
  function getBadgeCounts(email, role) {
    if (!email) return { notificationCount: 0, taskCount: 0, overdueTaskCount: 0, qaUnansweredCount: 0 };
    email = String(email).trim().toLowerCase();

    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, role || 'steward');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { log_('getBadgeCounts', 'Error getting notification count: ' + (_e.message || _e)); }

    var taskCount = 0;
    var overdueTaskCount = 0;
    if (role === 'steward') {
      try {
        var openTasks = getTasks(email, 'open');
        taskCount = openTasks.length;
        for (var t = 0; t < openTasks.length; t++) {
          if (openTasks[t].dueDays !== null && openTasks[t].dueDays < 0) overdueTaskCount++;
        }
      } catch (_e) { log_('getBadgeCounts', 'Error getting tasks: ' + (_e.message || _e)); }
    }

    var qaUnansweredCount = 0;
    if (role === 'steward') {
      try {
        if (typeof QAForum !== 'undefined') {
          qaUnansweredCount = QAForum.getUnansweredCount();
        }
      } catch (_e) { log_('getBadgeCounts', 'Error getting QA count: ' + (_e.message || _e)); }
    }

    return {
      notificationCount: notifCount,
      taskCount: taskCount,
      overdueTaskCount: overdueTaskCount,
      qaUnansweredCount: qaUnansweredCount,
    };
  }

  // getMyFeedback removed v4.52.0 (Feedback sheet removed)

  // ═══════════════════════════════════════
  // NOTE v4.24.0: Legacy FlashPolls functions (getActivePolls, submitPollVote, addPoll)
  // removed here. All poll functionality is now in 24_WeeklyQuestions.gs via wq* wrappers.
  // FlashPolls/PollResponses sheets also removed from 23_PortalSheets.gs.
  // ═══════════════════════════════════════

  // ═══════════════════════════════════════
  // HELPERS: Meeting Minutes (v4.52.0)
  // ═══════════════════════════════════════

  /**
   * Resolves the Drive folder ID for meeting minutes documents.
   * Checks Config sheet first, falls back to ScriptProperties.
   * @returns {string} Folder ID or empty string if not configured.
   */
  function resolveMinutesFolderId_() {
    var folderId = '';
    try {
      if (typeof getConfigValue_ === 'function' && typeof CONFIG_COLS !== 'undefined' && CONFIG_COLS.MINUTES_FOLDER_ID) {
        folderId = getConfigValue_(CONFIG_COLS.MINUTES_FOLDER_ID) || '';
      }
      if (!folderId && typeof PropertiesService !== 'undefined') {
        folderId = PropertiesService.getScriptProperties().getProperty('MINUTES_FOLDER_ID') || '';
      }
    } catch (_e) { log_('resolveMinutesFolderId_', _e.message || _e); }
    return folderId;
  }

  /**
   * Creates a formatted Google Doc for meeting minutes and moves it to the Minutes folder.
   * @param {string} title - Meeting title (will be sanitized)
   * @param {Date} meetingDate - Meeting date
   * @param {string} createdBy - Steward email
   * @param {string} bullets - Key points (newline-separated)
   * @param {string} fullMinutes - Detailed notes
   * @returns {{ docUrl: string }} URL of created doc, or empty string on failure
   */
  function createMinutesDoc_(title, meetingDate, createdBy, bullets, fullMinutes) {
    var docUrl = '';
    try {
      // Sanitize title for Doc name: strip control chars, enforce length
      var safeTitle = String(title)
        .substring(0, MINUTES_LIMITS.TITLE_MAX)
        .replace(/[\x00-\x1F\x7F]/g, '')
        .trim() || '(Untitled)';
      var tz = Session.getScriptTimeZone();
      var docTitle = safeTitle + ' \u2014 ' + Utilities.formatDate(meetingDate, tz, 'yyyy-MM-dd');

      var doc = DocumentApp.create(docTitle);
      var body = doc.getBody();
      body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('Recorded by: ' + createdBy);
      body.appendParagraph('Date: ' + Utilities.formatDate(meetingDate, tz, 'MMMM d, yyyy'));
      body.appendParagraph('');

      if (bullets) {
        body.appendParagraph('Key Points').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        String(bullets).split('\n').forEach(function(line) {
          if (line.trim()) body.appendListItem(line.replace(/^[-\u2022*]\s*/, '').trim());
        });
        body.appendParagraph('');
      }
      if (fullMinutes) {
        body.appendParagraph('Full Minutes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(String(fullMinutes));
      }
      doc.saveAndClose();

      // Move to Minutes folder using file.moveTo() (replaces deprecated addFile/removeFile)
      var docFile = DriveApp.getFileById(doc.getId());
      var minutesFolderId = resolveMinutesFolderId_();
      if (minutesFolderId) {
        try {
          docFile.moveTo(DriveApp.getFolderById(minutesFolderId));
        } catch (moveErr) {
          log_('createMinutesDoc_', 'could not move doc to Minutes folder: ' + moveErr.message);
        }
      }
      docUrl = docFile.getUrl();
    } catch (docErr) {
      log_('createMinutesDoc_', 'Drive doc creation failed (non-fatal): ' + docErr.message);
    }
    return { docUrl: docUrl };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Meeting Minutes (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns meeting minutes, most recent first.
   * @param {number} limit
   * @returns {Object[]}
   */
  function getMeetingMinutes(limit) {
    try {
    limit = limit || 20;
    var sheet = (typeof getOrCreateMinutesSheet === 'function') ? getOrCreateMinutesSheet() : null;
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getDataRange().getValues();
    var minutes = [];
    for (var i = 1; i < data.length; i++) {
      var meetingDate = data[i][PORTAL_MINUTES_COLS.MEETING_DATE];
      var dateStr = meetingDate instanceof Date ? _formatDate(meetingDate) : String(meetingDate || '');
      var dateTs = meetingDate instanceof Date ? meetingDate.getTime() : 0;

      minutes.push({
        id: String(data[i][PORTAL_MINUTES_COLS.ID] || ''),
        meetingDate: dateStr,
        meetingDateTs: dateTs,
        title: String(data[i][PORTAL_MINUTES_COLS.TITLE] || ''),
        bullets: String(data[i][PORTAL_MINUTES_COLS.BULLETS] || ''),
        fullMinutes: String(data[i][PORTAL_MINUTES_COLS.FULL_MINUTES] || ''),
        createdBy: String(data[i][PORTAL_MINUTES_COLS.CREATED_BY] || ''),
        createdDate: data[i][PORTAL_MINUTES_COLS.CREATED_DATE] instanceof Date
          ? _formatDate(data[i][PORTAL_MINUTES_COLS.CREATED_DATE]) : '',
        driveDocUrl: String(data[i][PORTAL_MINUTES_COLS.DRIVE_DOC_URL] || ''),
        attachmentUrl: String(data[i][PORTAL_MINUTES_COLS.ATTACHMENT_URL] || ''),
        attachmentName: String(data[i][PORTAL_MINUTES_COLS.ATTACHMENT_NAME] || '')
      });
    }
    minutes.sort(function(a, b) { return (b.meetingDateTs || 0) - (a.meetingDateTs || 0); });
    minutes.forEach(function(m) { delete m.meetingDateTs; });
    return minutes.slice(0, limit);
    } catch (e) {
      log_('getMeetingMinutes error', e.message + '\n' + (e.stack || ''));
      return [];
    }
  }

  /**
   * Adds new meeting minutes with optional file attachment (steward-only).
   * Three-phase design: validate outside lock, fast sheet write inside lock,
   * slow Drive operations outside lock (idempotent).
   *
   * @param {string} stewardEmail
   * @param {Object} minutesData - { title, meetingDate, bullets, fullMinutes, attachmentData?, attachmentName? }
   * @param {string} [idemKey] - Idempotency key (10-min TTL)
   * @returns {Object}
   */
  function addMeetingMinutes(stewardEmail, minutesData, idemKey) {
    // ── Phase 1: Validate + prepare (outside lock) ──────────────────────────
    if (!stewardEmail || !minutesData || !minutesData.title || !String(minutesData.title).trim()) {
      return { success: false, message: 'Missing required fields.' };
    }

    // Attachment validation (before any expensive work)
    var ALLOWED_EXT = /\.(pdf|docx?|xlsx?|pptx?|txt|rtf|csv|png|jpe?g|gif)$/i;
    var hasAttachment = !!(minutesData.attachmentData && minutesData.attachmentName);
    if (minutesData.attachmentData && !minutesData.attachmentName) {
      return { success: false, message: 'Attachment filename is required.' };
    }
    if (hasAttachment) {
      if (!ALLOWED_EXT.test(minutesData.attachmentName)) {
        return { success: false, message: 'File type not supported. Allowed: PDF, Word, Excel, PowerPoint, images, text.' };
      }
      if (minutesData.attachmentData.length * 0.75 > MINUTES_LIMITS.ATTACHMENT_MAX_BYTES) {
        return { success: false, message: 'File too large. Maximum size is 10MB.' };
      }
    }

    var sheet = (typeof getOrCreateMinutesSheet === 'function') ? getOrCreateMinutesSheet() : null;
    if (!sheet) return { success: false, message: 'Minutes sheet unavailable.' };

    var id = 'MIN_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 5);

    // Date parsing: append T12:00:00 to YYYY-MM-DD to avoid UTC midnight timezone shift
    var rawDate = minutesData.meetingDate;
    var meetingDate = rawDate
      ? new Date(typeof rawDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate + 'T12:00:00' : rawDate)
      : new Date();
    if (isNaN(meetingDate.getTime())) meetingDate = new Date();

    // Escape + truncate text fields
    var safeTitle = escapeForFormula(String(minutesData.title).substring(0, MINUTES_LIMITS.TITLE_MAX));
    var safeBullets = escapeForFormulaPreserveNewlines(String(minutesData.bullets || '').substring(0, MINUTES_LIMITS.BULLETS_MAX));
    var safeFullMinutes = escapeForFormulaPreserveNewlines(String(minutesData.fullMinutes || '').substring(0, MINUTES_LIMITS.FULL_MINUTES_MAX));

    // Decode attachment blob (before lock, so decode time doesn't hold the lock)
    var attachmentBlob = null;
    if (hasAttachment) {
      try {
        var bytes = Utilities.base64Decode(minutesData.attachmentData);
        attachmentBlob = Utilities.newBlob(bytes, '', minutesData.attachmentName);
      } catch (decodeErr) {
        return { success: false, message: 'Failed to decode attachment: ' + decodeErr.message };
      }
    }

    // ── Phase 2: Sheet write (inside lock — fast, ~100ms) ───────────────────
    var rowNum;
    try {
      rowNum = withScriptLock_(function() {
        // Idempotency check (inside lock for atomicity with the write)
        if (idemKey) {
          var idemCache = CacheService.getScriptCache();
          if (idemCache.get('IDEM_' + idemKey)) return -1; // signal: duplicate
          idemCache.put('IDEM_' + idemKey, '1', 600);
        }
        sheet.appendRow([
          id,
          meetingDate,
          safeTitle,
          safeBullets,
          safeFullMinutes,
          String(stewardEmail).trim().toLowerCase(),
          new Date(),
          '',  // driveDocUrl — filled in Phase 3
          '',  // attachmentUrl — filled in Phase 3
          ''   // attachmentName — filled in Phase 3
        ]);
        return sheet.getLastRow();
      });
    } catch (_lockErr) {
      return { success: false, message: 'Could not save — another operation in progress. Please try again.' };
    }

    if (rowNum === -1) {
      return { duplicate: true, message: 'Duplicate request ignored' };
    }

    // ── Phase 3: Drive operations (outside lock — slow but idempotent) ──────
    var driveDocUrl = '';
    var attachmentUrl = '';
    var attachmentName = '';

    // Create Google Doc
    var docResult = createMinutesDoc_(minutesData.title, meetingDate, stewardEmail, minutesData.bullets || '', minutesData.fullMinutes || '');
    driveDocUrl = docResult.docUrl;

    // Upload attachment
    if (attachmentBlob) {
      try {
        var attFile = DriveApp.createFile(attachmentBlob);
        var minutesFolderId = resolveMinutesFolderId_();
        if (minutesFolderId) {
          try { attFile.moveTo(DriveApp.getFolderById(minutesFolderId)); } catch (_mv) { log_('addMeetingMinutes', 'attach move failed: ' + _mv.message); }
        }
        attachmentUrl = attFile.getUrl();
        attachmentName = minutesData.attachmentName;
      } catch (attErr) {
        log_('addMeetingMinutes', 'Attachment upload failed (non-fatal): ' + attErr.message);
      }
    }

    // Update row with Drive URLs
    if (driveDocUrl || attachmentUrl) {
      try {
        if (driveDocUrl) sheet.getRange(rowNum, PORTAL_MINUTES_COLS.DRIVE_DOC_URL + 1).setValue(escapeForFormula(driveDocUrl));
        if (attachmentUrl) {
          sheet.getRange(rowNum, PORTAL_MINUTES_COLS.ATTACHMENT_URL + 1).setValue(escapeForFormula(attachmentUrl));
          sheet.getRange(rowNum, PORTAL_MINUTES_COLS.ATTACHMENT_NAME + 1).setValue(escapeForFormula(attachmentName));
        }
      } catch (updateErr) {
        log_('addMeetingMinutes', 'Row URL update failed (non-fatal): ' + updateErr.message);
      }
    }

    // Audit
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MINUTES_ADDED', { steward: stewardEmail, title: minutesData.title, driveDocUrl: driveDocUrl, attachmentName: attachmentName });
    }

    var msg = 'Minutes added.';
    if (driveDocUrl) msg += ' Google Doc saved.';
    if (attachmentUrl) msg += ' Attachment uploaded.';
    return { success: true, message: msg, id: id, driveDocUrl: driveDocUrl, attachmentUrl: attachmentUrl, attachmentName: attachmentName };
  }

  // addPoll removed v4.24.0 — was FlashPolls-based. Use wqSetStewardQuestion() instead.

  // ═══════════════════════════════════════
  // Steward Performance (v4.18.0)
  // ═══════════════════════════════════════

  /**
   * Returns performance metrics for a single steward.
   * @param {string} email - Steward email
   * @returns {Object} Performance data or empty object
   */
  function getStewardPerformance(email) {
    try {
      if (!email) return {};
      email = String(email).trim().toLowerCase();

      var ss = _getSS();
      if (!ss) return {};
      var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
      if (!sheet || sheet.getLastRow() <= 1) return {};

      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowEmail = String(col_(data[i], STEWARD_PERF_COLS.STEWARD)).trim().toLowerCase();
        if (rowEmail !== email) continue;
        return {
          steward: rowEmail,
          totalCases: Number(col_(data[i], STEWARD_PERF_COLS.TOTAL_CASES)) || 0,
          active: Number(col_(data[i], STEWARD_PERF_COLS.ACTIVE)) || 0,
          closed: Number(col_(data[i], STEWARD_PERF_COLS.CLOSED)) || 0,
          won: Number(col_(data[i], STEWARD_PERF_COLS.WON)) || 0,
          winRate: Number(col_(data[i], STEWARD_PERF_COLS.WIN_RATE)) || 0,
          avgDays: Number(col_(data[i], STEWARD_PERF_COLS.AVG_DAYS)) || 0,
          overdue: Number(col_(data[i], STEWARD_PERF_COLS.OVERDUE)) || 0,
          dueThisWeek: Number(col_(data[i], STEWARD_PERF_COLS.DUE_THIS_WEEK)) || 0,
          performanceScore: Number(col_(data[i], STEWARD_PERF_COLS.PERFORMANCE_SCORE)) || 0,
        };
      }
      return {};
    } catch (_e) {
      log_('getStewardPerformance error', _e.message);
      return {};
    }
  }

  /**
   * Returns performance metrics for all stewards.
   * @returns {Object[]}
   */
  function getAllStewardPerformance() {
    try {
      var ss = _getSS();
      if (!ss) return [];
      var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
      if (!sheet || sheet.getLastRow() <= 1) return [];

      var data = sheet.getDataRange().getValues();
      var results = [];
      for (var i = 1; i < data.length; i++) {
        var steward = String(col_(data[i], STEWARD_PERF_COLS.STEWARD)).trim();
        if (!steward) continue;
        results.push({
          steward: steward.toLowerCase(),
          totalCases: Number(col_(data[i], STEWARD_PERF_COLS.TOTAL_CASES)) || 0,
          active: Number(col_(data[i], STEWARD_PERF_COLS.ACTIVE)) || 0,
          closed: Number(col_(data[i], STEWARD_PERF_COLS.CLOSED)) || 0,
          won: Number(col_(data[i], STEWARD_PERF_COLS.WON)) || 0,
          winRate: Number(col_(data[i], STEWARD_PERF_COLS.WIN_RATE)) || 0,
          avgDays: Number(col_(data[i], STEWARD_PERF_COLS.AVG_DAYS)) || 0,
          overdue: Number(col_(data[i], STEWARD_PERF_COLS.OVERDUE)) || 0,
          dueThisWeek: Number(col_(data[i], STEWARD_PERF_COLS.DUE_THIS_WEEK)) || 0,
          performanceScore: Number(col_(data[i], STEWARD_PERF_COLS.PERFORMANCE_SCORE)) || 0,
        });
      }
      return results;
    } catch (_e) {
      log_('getAllStewardPerformance error', _e.message);
      return [];
    }
  }

  // ═══════════════════════════════════════
  // Case Checklist (v4.18.0)
  // ═══════════════════════════════════════

  /**
   * Returns checklist items for a given case.
   * @param {string} caseId
   * @returns {Object[]}
   */
  function getCaseChecklist(caseId) {
    try {
      if (!caseId) return [];
      caseId = String(caseId).trim();

      var ss = _getSS();
      if (!ss) return [];
      var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
      if (!sheet || sheet.getLastRow() <= 1) return [];

      var data = sheet.getDataRange().getValues();
      var items = [];
      for (var i = 1; i < data.length; i++) {
        if (String(col_(data[i], CHECKLIST_COLS.CASE_ID)).trim() !== caseId) continue;
        var dueDate = col_(data[i], CHECKLIST_COLS.DUE_DATE);
        var completedDate = col_(data[i], CHECKLIST_COLS.COMPLETED_DATE);
        items.push({
          id: String(col_(data[i], CHECKLIST_COLS.CHECKLIST_ID) || ''),
          caseId: caseId,
          actionType: String(col_(data[i], CHECKLIST_COLS.ACTION_TYPE) || ''),
          itemText: String(col_(data[i], CHECKLIST_COLS.ITEM_TEXT) || ''),
          category: String(col_(data[i], CHECKLIST_COLS.CATEGORY) || ''),
          required: isTruthyValue(col_(data[i], CHECKLIST_COLS.REQUIRED)),
          completed: isTruthyValue(col_(data[i], CHECKLIST_COLS.COMPLETED)),
          completedBy: String(col_(data[i], CHECKLIST_COLS.COMPLETED_BY) || ''),
          completedDate: completedDate instanceof Date ? _formatDate(completedDate) : '',
          dueDate: dueDate instanceof Date ? _formatDate(dueDate) : String(dueDate || ''),
          notes: String(col_(data[i], CHECKLIST_COLS.NOTES) || ''),
          sortOrder: Number(col_(data[i], CHECKLIST_COLS.SORT_ORDER)) || 0,
        });
      }
      items.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
      return items;
    } catch (_e) {
      log_('getCaseChecklist error', _e.message);
      return [];
    }
  }

  /**
   * Returns completion progress for a case checklist.
   * @param {string} caseId
   * @returns {Object} { total, completed, required, requiredCompleted, percent }
   */
  function getCaseChecklistProgress(caseId) {
    try {
      var items = getCaseChecklist(caseId);
      var total = items.length;
      var completed = 0;
      var required = 0;
      var requiredCompleted = 0;
      for (var i = 0; i < items.length; i++) {
        if (items[i].completed) completed++;
        if (items[i].required) {
          required++;
          if (items[i].completed) requiredCompleted++;
        }
      }
      return {
        total: total,
        completed: completed,
        required: required,
        requiredCompleted: requiredCompleted,
        percent: total > 0 ? Math.round((completed / total) * 100) : 0,
      };
    } catch (_e) {
      log_('getCaseChecklistProgress error', _e.message);
      return { total: 0, completed: 0, required: 0, requiredCompleted: 0, percent: 0 };
    }
  }

  /**
   * Toggles a checklist item's completed state.
   * @param {string} checklistId
   * @param {boolean} completed
   * @param {string} email - Who completed it
   * @returns {Object} { success, message }
   */
  function toggleChecklistItem(checklistId, completed, email) {
    try {
      if (!checklistId) return { success: false, message: 'Missing checklist ID.' };

      var ss = _getSS();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.CASE_CHECKLIST);
      if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Checklist not found.' };

      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(col_(data[i], CHECKLIST_COLS.CHECKLIST_ID)).trim() !== String(checklistId).trim()) continue;
        var rowNum = i + 1;
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED).setValue(completed ? true : false);
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED_BY).setValue(completed ? String(email || '').toLowerCase().trim() : '');
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED_DATE).setValue(completed ? new Date() : '');

        if (typeof logAuditEvent === 'function') {
          logAuditEvent('CHECKLIST_TOGGLED', { checklistId: checklistId, completed: completed, by: email });
        }
        return { success: true, message: completed ? 'Item completed.' : 'Item unchecked.' };
      }
      return { success: false, message: 'Checklist item not found.' };
    } catch (_e) {
      log_('toggleChecklistItem error', _e.message);
      return { success: false, message: 'Failed to update checklist.' };
    }
  }

  // ═══════════════════════════════════════
  // Member Meetings (v4.18.0)
  // ═══════════════════════════════════════

  /**
   * Returns meetings a member has checked into.
   * @param {string} email
   * @returns {Object[]}
   */
  function getMemberMeetings(email) {
    try {
      if (!email) return [];
      email = String(email).trim().toLowerCase();

      var ss = _getSS();
      if (!ss) return [];
      var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
      if (!sheet || sheet.getLastRow() <= 1) return [];

      var data = sheet.getDataRange().getValues();
      var meetings = [];
      for (var i = 1; i < data.length; i++) {
        var rowEmail = String(col_(data[i], MEETING_CHECKIN_COLS.EMAIL)).trim().toLowerCase();
        if (rowEmail !== email) continue;
        var meetingDate = col_(data[i], MEETING_CHECKIN_COLS.MEETING_DATE);
        var checkinTime = col_(data[i], MEETING_CHECKIN_COLS.CHECKIN_TIME);
        var rawDuration = col_(data[i], MEETING_CHECKIN_COLS.MEETING_DURATION);
        var durationStr = rawDuration ? String(rawDuration) : null;
        if (durationStr && !isNaN(parseFloat(durationStr))) {
          var hrs = parseFloat(durationStr);
          durationStr = hrs === 1 ? '1 hour' : hrs + ' hours';
        }
        meetings.push({
          meetingId: String(col_(data[i], MEETING_CHECKIN_COLS.MEETING_ID) || ''),
          meetingName: String(col_(data[i], MEETING_CHECKIN_COLS.MEETING_NAME) || ''),
          meetingDate: meetingDate instanceof Date ? _formatDate(meetingDate) : String(meetingDate || ''),
          meetingType: String(col_(data[i], MEETING_CHECKIN_COLS.MEETING_TYPE) || ''),
          checkinTime: checkinTime instanceof Date ? _formatDate(checkinTime) : String(checkinTime || ''),
          duration: durationStr,
          notesUrl: String(col_(data[i], MEETING_CHECKIN_COLS.NOTES_DOC_URL) || '') || null,
          agendaUrl: String(col_(data[i], MEETING_CHECKIN_COLS.AGENDA_DOC_URL) || '') || null,
          minutesTitle: null,
          minutesBullets: null,
        });
      }
      meetings.reverse(); // newest first
      return meetings;
    } catch (_e) {
      log_('getMemberMeetings error', _e.message);
      return [];
    }
  }

  /**
   * Returns aggregate meeting statistics.
   * @returns {Object} { totalMeetings, totalCheckins, uniqueAttendees, avgAttendance }
   */
  function getMeetingStats() {
    try {
      var ss = _getSS();
      if (!ss) return { totalMeetings: 0, totalCheckins: 0, uniqueAttendees: 0, avgAttendance: 0 };
      var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
      if (!sheet || sheet.getLastRow() <= 1) {
        return { totalMeetings: 0, totalCheckins: 0, uniqueAttendees: 0, avgAttendance: 0 };
      }

      var data = sheet.getDataRange().getValues();
      var meetingIds = {};
      var attendees = {};
      for (var i = 1; i < data.length; i++) {
        var mid = String(col_(data[i], MEETING_CHECKIN_COLS.MEETING_ID)).trim();
        var aemail = String(col_(data[i], MEETING_CHECKIN_COLS.EMAIL)).trim().toLowerCase();
        if (mid) meetingIds[mid] = true;
        if (aemail) attendees[aemail] = true;
      }
      var totalMeetings = Object.keys(meetingIds).length;
      var totalCheckins = data.length - 1;
      var uniqueAttendees = Object.keys(attendees).length;
      return {
        totalMeetings: totalMeetings,
        totalCheckins: totalCheckins,
        uniqueAttendees: uniqueAttendees,
        avgAttendance: totalMeetings > 0 ? Math.round(totalCheckins / totalMeetings) : 0,
      };
    } catch (_e) {
      log_('getMeetingStats error', _e.message);
      return { totalMeetings: 0, totalCheckins: 0, uniqueAttendees: 0, avgAttendance: 0 };
    }
  }

  // ═══════════════════════════════════════
  // Satisfaction (v4.18.0)
  // ═══════════════════════════════════════

  /**
   * Returns satisfaction survey trend data (aggregated averages).
   * @returns {Object} { overallSat, stewardRating, stewardAccess, chapter, leadership, contract, representation, communication, memberVoice, valueAction }
   */
  /**
   * Returns satisfaction trends for the Insights page.
   * v4.23.1: Fully rewritten — uses getSatisfactionSummary() (dynamic col map,
   * section-key grouping). Returns the shape steward_view.html expects:
   *   { overall, responseCount, categories: [{ name, avg }] }
   *
   * Replaces the old positional SATISFACTION_COLS reads, which broke when the
   * v4.23.0 dynamic schema removed fixed column positions.
   */
  function getSatisfactionTrends() {
    try {
      var summary = getSatisfactionSummary();
      if (!summary || !summary.sections) return { overall: 0, responseCount: 0, categories: [] };

      var secs = summary.sections;

      // Overall = OVERALL_SAT section average
      var overall = 0;
      if (secs['OVERALL_SAT'] && secs['OVERALL_SAT'].avg !== null) {
        overall = secs['OVERALL_SAT'].avg;
      }

      // Categories dynamically from summary — skip sections with no data
      var categories = [];
      Object.keys(secs).forEach(function(key) {
        var s = secs[key];
        if (s && s.avg !== null && s.count > 0) {
          categories.push({ name: s.name, avg: s.avg });
        }
      });

      return {
        overall:       Math.round(overall * 10) / 10,
        responseCount: summary.responseCount || 0,
        categories:    categories
      };
    } catch (_e) {
      log_('getSatisfactionTrends error', _e.message);
      return { overall: 0, responseCount: 0, categories: [] };
    }
  }

  // submitFeedback removed v4.52.0 (Feedback sheet removed)

  // ═══════════════════════════════════════
  // PAGINATED MEMBERS — server-side pagination for large member directories
  // ═══════════════════════════════════════

  /**
   * Returns a page of members with optional search/filter.
   * Avoids transferring 8K+ records to the frontend.
   * @param {string} stewardEmail - Authenticated steward email
   * @param {Object} opts - { page: 1, pageSize: 50, search: '', filter: { steward, location, hasGrievance } }
   * @returns {{ items: Object[], total: number, page: number, pageSize: number, totalPages: number }}
   */
  function getMembersPaginated(stewardEmail, opts) {
    opts = opts || {};
    var page = Math.max(1, opts.page || 1);
    var pageSize = Math.min(100, Math.max(10, opts.pageSize || 50));
    var search = String(opts.search || '').trim().toLowerCase();
    var filter = opts.filter || {};

    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return { items: [], total: 0, page: page, pageSize: pageSize, totalPages: 0 };

    var data = cached.data;
    var colMap = cached.colMap;
    var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
    var nameCol = _findColumn(colMap, HEADERS.memberName);
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    var locCol = _findColumn(colMap, HEADERS.memberWorkLocation);

    // Build filtered index (row indices that match)
    var matches = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = emailCol !== -1 ? String(data[i][emailCol]).trim() : '';
      if (!rowEmail) continue;

      // Search filter — match name, email, or location
      if (search) {
        var name = nameCol !== -1 ? String(data[i][nameCol]).toLowerCase() : '';
        var loc = locCol !== -1 ? String(data[i][locCol]).toLowerCase() : '';
        if (name.indexOf(search) === -1 && rowEmail.toLowerCase().indexOf(search) === -1 && loc.indexOf(search) === -1) continue;
      }

      // Steward filter
      if (filter.steward) {
        var assignedTo = stewardCol !== -1 ? String(data[i][stewardCol]).trim().toLowerCase() : '';
        if (assignedTo !== String(filter.steward).toLowerCase()) continue;
      }

      // Location filter
      if (filter.location) {
        var memberLoc = locCol !== -1 ? String(data[i][locCol]).trim() : '';
        if (memberLoc !== filter.location) continue;
      }

      // Open grievance filter
      if (filter.hasGrievance) {
        var rec = _buildUserRecord(data[i], colMap);
        if (!rec.hasOpenGrievance) continue;
      }

      // Unit filter
      if (filter.unit) {
        var unitCol = _findColumn(colMap, HEADERS.memberUnit);
        var memberUnit = unitCol !== -1 ? String(data[i][unitCol]).trim() : '';
        if (memberUnit !== filter.unit) continue;
      }

      // Dues status filter
      if (filter.duesPaying !== undefined && filter.duesPaying !== '') {
        var recForDues = _buildUserRecord(data[i], colMap);
        var wantPaying = filter.duesPaying === true || filter.duesPaying === 'true';
        if (wantPaying ? recForDues.duesPaying === false : recForDues.duesPaying !== false) continue;
      }

      // Office days filter (partial text match)
      if (filter.officeDays) {
        var odCol = _findColumn(colMap, HEADERS.memberOfficeDays);
        var memberOd = odCol !== -1 ? String(data[i][odCol]).trim().toLowerCase() : '';
        if (memberOd.indexOf(String(filter.officeDays).toLowerCase()) === -1) continue;
      }

      matches.push(i);
    }

    var total = matches.length;
    var totalPages = Math.ceil(total / pageSize) || 1;
    var startIdx = (page - 1) * pageSize;
    var pageIndices = matches.slice(startIdx, startIdx + pageSize);

    var items = pageIndices.map(function(rowIdx) {
      var rec = _buildUserRecord(data[rowIdx], colMap);
      return {
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        cubicle: rec.cubicle,
        hireDate: rec.hireDate,
        duesStatus: rec.duesStatus,
        hasOpenGrievance: rec.hasOpenGrievance,
        assignedSteward: stewardCol !== -1 ? String(data[rowIdx][stewardCol]).trim() : ''
      };
    });

    return { items: items, total: total, page: page, pageSize: pageSize, totalPages: totalPages };
  }

  /**
   * Returns member count and active-grievance count (lightweight — no full records).
   * @returns {{ total: number, withGrievances: number }}
   */
  function getMemberCount() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return { total: 0, withGrievances: 0 };
    var data = cached.data;
    var colMap = cached.colMap;
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    if (emailCol === -1) return { total: 0, withGrievances: 0 };
    var total = 0;
    var withGrievances = 0;
    for (var i = 1; i < data.length; i++) {
      if (!String(data[i][emailCol]).trim()) continue;
      total++;
      var rec = _buildUserRecord(data[i], colMap);
      if (rec.hasOpenGrievance) withGrievances++;
    }
    return { total: total, withGrievances: withGrievances };
  }

  /**
   * Returns distinct values for member filter dropdowns.
   * Avoids client-side computation on large member lists.
   * @returns {{ locations: string[], units: string[], stewards: string[] }}
   */
  function getFilterDropdownValues() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return { locations: [], units: [], stewards: [] };
    var data = cached.data;
    var colMap = cached.colMap;
    var locCol = _findColumn(colMap, HEADERS.memberWorkLocation);
    var unitCol = _findColumn(colMap, HEADERS.memberUnit);
    var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    var locs = {}, units = {}, stewards = {};
    for (var i = 1; i < data.length; i++) {
      if (emailCol !== -1 && !String(data[i][emailCol]).trim()) continue;
      if (locCol !== -1) { var l = String(data[i][locCol]).trim(); if (l) locs[l] = true; }
      if (unitCol !== -1) { var u = String(data[i][unitCol]).trim(); if (u) units[u] = true; }
      if (stewardCol !== -1) { var s = String(data[i][stewardCol]).trim(); if (s) stewards[s] = true; }
    }
    return { locations: Object.keys(locs).sort(), units: Object.keys(units).sort(), stewards: Object.keys(stewards).sort() };
  }

  /**
   * Returns row counts and scale status for key sheets.
   * Frontend uses this to decide between bulk-load and paginated mode.
   * @returns {{ members: Object, grievances: Object }}
   */
  function getSheetHealth() {
    // Trigger cache reads to populate _sheetHealthLog
    _getCachedSheetData(MEMBER_SHEET);
    _getCachedSheetData(GRIEVANCE_SHEET);
    var thresholds = (typeof SCALE_THRESHOLDS !== 'undefined') ? SCALE_THRESHOLDS : { WARN_ROWS: 5000, THROTTLE_ROWS: 7000, CRITICAL_ROWS: 8000 };

    function _status(count) {
      if (count >= thresholds.CRITICAL_ROWS) return 'critical';
      if (count >= thresholds.THROTTLE_ROWS) return 'throttle';
      if (count >= thresholds.WARN_ROWS) return 'warn';
      return 'ok';
    }

    var memberRows = _sheetHealthLog[MEMBER_SHEET] || 0;
    var grievanceRows = _sheetHealthLog[GRIEVANCE_SHEET] || 0;

    return {
      members: { rows: memberRows, status: _status(memberRows) },
      grievances: { rows: grievanceRows, status: _status(grievanceRows) }
    };
  }

  // ═══════════════════════════════════════
  // GRIEVANCE FEEDBACK (v4.32.0)
  // ═══════════════════════════════════════

  /**
   * Returns the _Grievance_Feedback hidden sheet, auto-creating with headers if absent.
   * @returns {Sheet|null}
   */
  function _ensureGrievanceFeedback() {
    var ss = _getSS();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FEEDBACK);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.GRIEVANCE_FEEDBACK);
      sheet.getRange(1, 1, 1, 10).setValues([[
        'ID', 'Grievance ID', 'Member Email', 'Steward Email',
        'Satisfaction', 'Communication', 'Timeliness', 'Fairness',
        'Comment', 'Created'
      ]]);
      sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Checks if the member has any recently closed grievances (within 14 days)
   * that haven't been rated yet.
   * @param {string} email - Member email
   * @returns {Object|null} { grievanceId, steward, stewardName, step, closedDate, issueCategory } or null
   */
  function getPendingGrievanceFeedback(email) {
    if (!email) return null;
    email = email.toLowerCase().trim();

    // 1. Get closed grievances for this member — use combined active+archive so
    //    cases that were recently auto-archived (archiveClosedGrievances runs daily)
    //    still generate a feedback prompt within the 14-day window.
    var cached = _getAllGrievanceData();
    if (!cached) return null;
    var data = cached.data;
    var colMap = cached.colMap;
    var closedStatuses = _closedStatusesLower;

    var now = new Date();
    var fourteenDaysAgo = now.getTime() - 14 * 86400000;
    var candidates = [];

    for (var i = 0; i < data.length; i++) {
      var rec = _buildGrievanceRecord(data[i], colMap);
      if (rec.memberEmail !== email) continue;
      if (closedStatuses.indexOf(rec.status) === -1) continue;
      if (rec.closedTimestamp < fourteenDaysAgo || rec.closedTimestamp === 0) continue;
      candidates.push(rec);
    }

    if (candidates.length === 0) return null;

    // 2. Check which ones already have feedback
    var fbSheet = _getSS().getSheetByName(SHEETS.GRIEVANCE_FEEDBACK);
    var ratedIds = {};
    if (fbSheet && fbSheet.getLastRow() >= 2) {
      var fbData = fbSheet.getRange(2, 1, fbSheet.getLastRow() - 1, 3).getValues();
      for (var fi = 0; fi < fbData.length; fi++) {
        if (String(fbData[fi][2]).toLowerCase().trim() === email) {
          ratedIds[String(fbData[fi][1]).trim()] = true;
        }
      }
    }

    // 3. Find first unrated
    for (var ci = 0; ci < candidates.length; ci++) {
      var c = candidates[ci];
      if (!ratedIds[c.id]) {
        var stewardName = c.steward;
        if (c.steward && c.steward !== 'unassigned') {
          var sUser = findUserByEmail(c.steward);
          if (sUser && sUser.name) stewardName = sUser.name;
        }
        return {
          grievanceId: c.id,
          steward: c.steward,
          stewardName: stewardName,
          step: c.step,
          closedDate: c.closedTimestamp ? Utilities.formatDate(new Date(c.closedTimestamp), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
          issueCategory: c.issueCategory,
          status: c.status
        };
      }
    }

    return null;
  }

  /**
   * Submits grievance feedback from a member.
   * @param {string} email - Member email (server-verified)
   * @param {string} grievanceId - The grievance ID being rated
   * @param {Object} ratings - { satisfaction, communication, timeliness, fairness } (1-5 each)
   * @param {string=} comment - Optional comment
   * @returns {Object} { success: true/false, message }
   */
  function submitGrievanceFeedback(email, grievanceId, ratings, comment) {
    if (!email || !grievanceId || !ratings) return { success: false, message: 'Missing fields.' };

    // Validate ratings are 1-5
    var fields = ['satisfaction', 'communication', 'timeliness', 'fairness'];
    for (var fi = 0; fi < fields.length; fi++) {
      var val = parseInt(ratings[fields[fi]], 10);
      if (isNaN(val) || val < 1 || val > 5) return { success: false, message: 'Ratings must be 1-5.' };
    }

    // Verify this grievance belongs to this member and is closed.
    // Use combined active+archive so members can still submit feedback
    // on cases that have been auto-archived after the 90-day threshold.
    var cached = _getAllGrievanceData();
    if (!cached) return { success: false, message: 'Could not read grievances.' };
    var data = cached.data;
    var colMap = cached.colMap;
    var closedStatuses = _closedStatusesLower;
    var matchedRec = null;

    for (var i = 0; i < data.length; i++) {
      var rec = _buildGrievanceRecord(data[i], colMap);
      if (rec.id === grievanceId && rec.memberEmail === email.toLowerCase().trim()) {
        if (closedStatuses.indexOf(rec.status) !== -1) matchedRec = rec;
        break;
      }
    }

    if (!matchedRec) return { success: false, message: 'Grievance not found or not closed.' };

    // Check for duplicate feedback
    var fbSheet = _ensureGrievanceFeedback();
    if (!fbSheet) return { success: false, message: 'Could not access feedback sheet.' };
    if (fbSheet.getLastRow() >= 2) {
      var existing = fbSheet.getRange(2, 1, fbSheet.getLastRow() - 1, 3).getValues();
      for (var ei = 0; ei < existing.length; ei++) {
        if (String(existing[ei][1]).trim() === grievanceId &&
            String(existing[ei][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
          return { success: false, message: 'Feedback already submitted for this grievance.' };
        }
      }
    }

    // Write feedback
    var id = 'GF-' + Utilities.getUuid().substring(0, 8);
    var safeComment = comment ? String(comment).substring(0, 500).trim() : '';
    fbSheet.appendRow([
      id,
      grievanceId,
      email.toLowerCase().trim(),
      matchedRec.steward || '',
      parseInt(ratings.satisfaction, 10),
      parseInt(ratings.communication, 10),
      parseInt(ratings.timeliness, 10),
      parseInt(ratings.fairness, 10),
      escapeForFormula(safeComment),
      new Date()
    ]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('GRIEVANCE_FEEDBACK', { grievanceId: grievanceId, member: email });
    }

    return { success: true, message: 'Thank you for your feedback.' };
  }

  /**
   * Returns aggregate grievance feedback stats for Union Stats.
   * @returns {Object|null} { totalFeedback, avgSatisfaction, avgCommunication, avgTimeliness, avgFairness,
   *                          avgOverall, bySteward: [{ label, avgScore, count }], trend: [{ month, avg }] }
   */
  function getGrievanceFeedbackStats() {
    var ss = _getSS();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FEEDBACK);
    if (!sheet || sheet.getLastRow() < 2) return { totalFeedback: 0 };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var sums = { satisfaction: 0, communication: 0, timeliness: 0, fairness: 0 };
    var count = 0;
    var bySteward = {}; // stewardEmail -> { total, count }
    var byMonth = {};   // 'YYYY-MM' -> { total, count }

    for (var i = 0; i < data.length; i++) {
      var sat = parseFloat(data[i][4]) || 0;
      var com = parseFloat(data[i][5]) || 0;
      var tim = parseFloat(data[i][6]) || 0;
      var fair = parseFloat(data[i][7]) || 0;
      if (sat < 1 || sat > 5) continue; // skip invalid rows

      var avg = (sat + com + tim + fair) / 4;
      sums.satisfaction += sat;
      sums.communication += com;
      sums.timeliness += tim;
      sums.fairness += fair;
      count++;

      // Per-steward aggregation
      var stw = String(data[i][3]).toLowerCase().trim();
      if (stw) {
        if (!bySteward[stw]) bySteward[stw] = { total: 0, count: 0 };
        bySteward[stw].total += avg;
        bySteward[stw].count++;
      }

      // Monthly trend
      var created = data[i][9];
      if (created instanceof Date) {
        var monthKey = created.getFullYear() + '-' + String(created.getMonth() + 1).padStart(2, '0');
        if (!byMonth[monthKey]) byMonth[monthKey] = { total: 0, count: 0 };
        byMonth[monthKey].total += avg;
        byMonth[monthKey].count++;
      }
    }

    if (count === 0) return { totalFeedback: 0 };

    // Build steward scores with real names
    var stewardScores = [];
    for (var sw in bySteward) {
      var sLabel = sw;
      var sUser = findUserByEmail(sw);
      if (sUser && sUser.name) sLabel = sUser.name;
      stewardScores.push({
        label: sLabel,
        avgScore: Math.round((bySteward[sw].total / bySteward[sw].count) * 10) / 10,
        count: bySteward[sw].count
      });
    }
    stewardScores.sort(function(a, b) { return b.avgScore - a.avgScore; });

    // Build monthly trend
    var trend = [];
    var months = Object.keys(byMonth).sort();
    for (var mi = 0; mi < months.length; mi++) {
      var m = months[mi];
      trend.push({
        month: m,
        avg: Math.round((byMonth[m].total / byMonth[m].count) * 10) / 10
      });
    }

    return {
      totalFeedback: count,
      avgSatisfaction: Math.round((sums.satisfaction / count) * 10) / 10,
      avgCommunication: Math.round((sums.communication / count) * 10) / 10,
      avgTimeliness: Math.round((sums.timeliness / count) * 10) / 10,
      avgFairness: Math.round((sums.fairness / count) * 10) / 10,
      avgOverall: Math.round(((sums.satisfaction + sums.communication + sums.timeliness + sums.fairness) / (count * 4)) * 10) / 10,
      bySteward: stewardScores,
      trend: trend
    };
  }

  /**
   * Returns feedback summary for a specific steward.
   * @param {string} stewardEmail
   * @returns {Object|null} { totalFeedback, avgSatisfaction, avgCommunication, avgTimeliness, avgFairness,
   *                          avgOverall, recentComments: [{ date, comment, grievanceId }] }
   */
  function getStewardFeedbackSummary(stewardEmail) {
    if (!stewardEmail) return null;
    stewardEmail = stewardEmail.toLowerCase().trim();
    var ss = _getSS();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FEEDBACK);
    if (!sheet || sheet.getLastRow() < 2) return { totalFeedback: 0 };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var sums = { satisfaction: 0, communication: 0, timeliness: 0, fairness: 0 };
    var count = 0;
    var comments = [];

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase().trim() !== stewardEmail) continue;
      var sat = parseFloat(data[i][4]) || 0;
      if (sat < 1 || sat > 5) continue;

      sums.satisfaction += sat;
      sums.communication += (parseFloat(data[i][5]) || 0);
      sums.timeliness += (parseFloat(data[i][6]) || 0);
      sums.fairness += (parseFloat(data[i][7]) || 0);
      count++;

      var comment = String(data[i][8] || '').trim();
      if (comment) {
        var created = data[i][9];
        comments.push({
          date: created instanceof Date ? Utilities.formatDate(created, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
          comment: comment,
          grievanceId: String(data[i][1]).trim()
        });
      }
    }

    if (count === 0) return { totalFeedback: 0 };

    // Most recent 5 comments
    comments.reverse();
    comments = comments.slice(0, 5);

    return {
      totalFeedback: count,
      avgSatisfaction: Math.round((sums.satisfaction / count) * 10) / 10,
      avgCommunication: Math.round((sums.communication / count) * 10) / 10,
      avgTimeliness: Math.round((sums.timeliness / count) * 10) / 10,
      avgFairness: Math.round((sums.fairness / count) * 10) / 10,
      avgOverall: Math.round(((sums.satisfaction + sums.communication + sums.timeliness + sums.fairness) / (count * 4)) * 10) / 10,
      recentComments: comments
    };
  }

  // v4.46.0 — Leader Hub helpers
  /**
   * Returns members filtered by unit.
   * @param {string} unit
   * @returns {Object[]} [{ name, email, unit }]
   */
  function getUnitMembers(unit) {
    if (!unit) return [];
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];
    var members = [];
    for (var i = 1; i < cached.data.length; i++) {
      var rec = _buildUserRecord(cached.data[i], cached.colMap);
      if (rec.unit === unit && rec.email) {
        members.push({ name: rec.name, email: rec.email, unit: rec.unit });
      }
    }
    return members;
  }

  /**
   * Sends a broadcast email to members in a specific unit.
   * @param {string} senderEmail - Leader/steward email (used as replyTo)
   * @param {string} senderName - Display name for subject fallback
   * @param {string} unit - Target unit name
   * @param {string} message - Email body
   * @param {string} [subject] - Optional email subject
   * @returns {Object} { success, sentCount, failedCount, message }
   */
  function sendUnitBroadcast(senderEmail, senderName, unit, message, subject) {
    var unitMembers = getUnitMembers(unit);
    // Exclude the sender from the recipient list
    unitMembers = unitMembers.filter(function(m) { return m.email.toLowerCase() !== senderEmail.toLowerCase(); });
    if (unitMembers.length === 0) return { success: false, sentCount: 0, failedCount: 0, message: 'No members found in unit.' };

    // v4.51.1: Rate limiting — max 3 broadcasts per leader per hour (BUG-11-001)
    var rateKey = 'BROADCAST_RATE_' + senderEmail.toLowerCase().trim();
    var cache = CacheService.getScriptCache();
    var rateCount = parseInt(cache.get(rateKey) || '0', 10);
    if (rateCount >= 3) {
      return { success: false, sentCount: 0, failedCount: 0, message: 'Broadcast rate limit reached (3 per hour). Please try again later.' };
    }
    cache.put(rateKey, String(rateCount + 1), 3600);

    // v4.51.1: Sanitize subject and message for formula injection (BUG-11-003)
    var safeSubj = escapeForFormula((subject && subject.trim()) ? subject.trim() : (senderName + ' \u2014 ' + unit + ' Update'));
    var safeMessage = escapeForFormula(message.trim());
    var sentCount = 0, failedCount = 0;
    for (var j = 0; j < unitMembers.length; j++) {
      // v4.51.1: Use safeSendEmail_ for quota protection (BUG-11-001)
      var sendResult = typeof safeSendEmail_ === 'function'
        ? safeSendEmail_({ to: unitMembers[j].email, subject: safeSubj, body: safeMessage, replyTo: senderEmail })
        : (function() { try { MailApp.sendEmail({ to: unitMembers[j].email, subject: safeSubj, body: safeMessage, replyTo: senderEmail }); return { success: true }; } catch (e) { return { success: false, error: e.message }; } })();
      if (sendResult && sendResult.success) {
        sentCount++;
      } else {
        secureLog('sendUnitBroadcast', 'send failed', { email: unitMembers[j].email, error: (sendResult && sendResult.error) || 'unknown' });
        failedCount++;
      }
    }
    return { success: true, sentCount: sentCount, failedCount: failedCount, message: 'Sent to ' + sentCount + ' member(s).' };
  }

  // ── Non-Member Contacts (v4.48.0) ──
  function getNonMemberContacts() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, NMC_HEADER_MAP_.length).getValues();
    return data.map(function(row) {
      return {
        contactId:    String(col_(row, NMC_COLS.CONTACT_ID) || ''),
        firstName:    String(col_(row, NMC_COLS.FIRST_NAME) || ''),
        lastName:     String(col_(row, NMC_COLS.LAST_NAME) || ''),
        jobTitle:     String(col_(row, NMC_COLS.JOB_TITLE) || ''),
        workLocation: String(col_(row, NMC_COLS.WORK_LOCATION) || ''),
        unit:         String(col_(row, NMC_COLS.UNIT) || ''),
        unionName:    String(col_(row, NMC_COLS.UNION_NAME) || ''),
        shirtSize:    String(col_(row, NMC_COLS.SHIRT_SIZE) || ''),
        isSteward:    String(col_(row, NMC_COLS.IS_STEWARD) || ''),
        email:        String(col_(row, NMC_COLS.EMAIL) || ''),
        phone:        String(col_(row, NMC_COLS.PHONE) || ''),
        category:     String(col_(row, NMC_COLS.CATEGORY) || ''),
        notes:        String(col_(row, NMC_COLS.NOTES) || '')
      };
    }).filter(function(c) { return c.firstName || c.lastName; });
  }

  function addNonMemberContact(contactData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS);
    if (!sheet) sheet = ensureNonMemberContactsSheet_(ss);
    var lastRow = sheet.getLastRow();
    var nextNum = 1;
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, NMC_COLS.CONTACT_ID, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        var m = String(ids[i][0]).match(/NMC-(\d+)/);
        if (m) nextNum = Math.max(nextNum, parseInt(m[1], 10) + 1);
      }
    }
    var contactId = 'NMC-' + String(nextNum).padStart(3, '0');
    var row = [];
    setCol_(row, NMC_COLS.CONTACT_ID, contactId);
    setCol_(row, NMC_COLS.FIRST_NAME, escapeForFormula(contactData.firstName || ''));
    setCol_(row, NMC_COLS.LAST_NAME, escapeForFormula(contactData.lastName || ''));
    setCol_(row, NMC_COLS.JOB_TITLE, escapeForFormula(contactData.jobTitle || ''));
    setCol_(row, NMC_COLS.WORK_LOCATION, escapeForFormula(contactData.workLocation || ''));
    setCol_(row, NMC_COLS.UNIT, escapeForFormula(contactData.unit || ''));
    setCol_(row, NMC_COLS.UNION_NAME, escapeForFormula(contactData.unionName || ''));
    setCol_(row, NMC_COLS.SHIRT_SIZE, escapeForFormula(contactData.shirtSize || ''));
    setCol_(row, NMC_COLS.IS_STEWARD, escapeForFormula(contactData.isSteward || 'No'));
    setCol_(row, NMC_COLS.EMAIL, escapeForFormula(contactData.email || ''));
    setCol_(row, NMC_COLS.PHONE, escapeForFormula(contactData.phone || ''));
    setCol_(row, NMC_COLS.CATEGORY, escapeForFormula(contactData.category || 'Other'));
    setCol_(row, NMC_COLS.NOTES, escapeForFormula(contactData.notes || ''));
    sheet.appendRow(row);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('NMC_CREATED', { contactId: contactId });
    }
    return { success: true, contactId: contactId };
  }

  function updateNonMemberContact(contactId, updateData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS);
    if (!sheet) return { success: false, error: 'Sheet not found.' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Contact not found.' };
    var ids = sheet.getRange(2, NMC_COLS.CONTACT_ID, lastRow - 1, 1).getValues();
    var rowIdx = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === contactId) { rowIdx = i + 2; break; }
    }
    if (rowIdx === -1) return { success: false, error: 'Contact not found.' };
    if (updateData.firstName !== undefined)    sheet.getRange(rowIdx, NMC_COLS.FIRST_NAME).setValue(escapeForFormula(updateData.firstName));
    if (updateData.lastName !== undefined)     sheet.getRange(rowIdx, NMC_COLS.LAST_NAME).setValue(escapeForFormula(updateData.lastName));
    if (updateData.jobTitle !== undefined)     sheet.getRange(rowIdx, NMC_COLS.JOB_TITLE).setValue(escapeForFormula(updateData.jobTitle));
    if (updateData.workLocation !== undefined) sheet.getRange(rowIdx, NMC_COLS.WORK_LOCATION).setValue(escapeForFormula(updateData.workLocation));
    if (updateData.unit !== undefined)         sheet.getRange(rowIdx, NMC_COLS.UNIT).setValue(escapeForFormula(updateData.unit));
    if (updateData.unionName !== undefined)    sheet.getRange(rowIdx, NMC_COLS.UNION_NAME).setValue(escapeForFormula(updateData.unionName));
    if (updateData.shirtSize !== undefined)    sheet.getRange(rowIdx, NMC_COLS.SHIRT_SIZE).setValue(escapeForFormula(updateData.shirtSize));
    if (updateData.isSteward !== undefined)    sheet.getRange(rowIdx, NMC_COLS.IS_STEWARD).setValue(escapeForFormula(updateData.isSteward));
    if (updateData.email !== undefined)        sheet.getRange(rowIdx, NMC_COLS.EMAIL).setValue(escapeForFormula(updateData.email));
    if (updateData.phone !== undefined)        sheet.getRange(rowIdx, NMC_COLS.PHONE).setValue(escapeForFormula(updateData.phone));
    if (updateData.category !== undefined)     sheet.getRange(rowIdx, NMC_COLS.CATEGORY).setValue(escapeForFormula(updateData.category));
    if (updateData.notes !== undefined)        sheet.getRange(rowIdx, NMC_COLS.NOTES).setValue(escapeForFormula(updateData.notes));
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('NMC_UPDATED', { contactId: contactId, fields: Object.keys(updateData) });
    }
    return { success: true };
  }

  function deleteNonMemberContact(contactId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.NON_MEMBER_CONTACTS);
    if (!sheet) return { success: false, error: 'Sheet not found.' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Contact not found.' };
    var ids = sheet.getRange(2, NMC_COLS.CONTACT_ID, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === contactId) {
        sheet.deleteRow(i + 2);
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('NMC_DELETED', { contactId: contactId });
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Contact not found.' };
  }

  // Public API
  return {
    findUserByEmail: findUserByEmail,
    getUserRole: getUserRole,
    getStewardCases: getStewardCases,
    getStewardKPIs: getStewardKPIs,
    getMemberGrievances: getMemberGrievances,
    getMemberGrievanceHistory: getMemberGrievanceHistory,
    getStewardContact: getStewardContact,
    getUnits: getUnits,
    getFullMemberProfile: getFullMemberProfile,
    updateMemberProfile: updateMemberProfile,
    updateMemberBySteward: updateMemberBySteward,
    getAssignedStewardInfo: getAssignedStewardInfo,
    getAvailableStewards: getAvailableStewards,
    assignStewardToMember: assignStewardToMember,
    getMemberGrievanceDriveUrl: getMemberGrievanceDriveUrl,
    getMemberSurveyStatus: getMemberSurveyStatus,
    logResourceClick: logResourceClick,
    getResourceClickTotal: getResourceClickTotal,
    getResourceStats: getResourceStats,
    logTabVisit: logTabVisit,
    getUsageStats: getUsageStats,
    getOrgLinks: getOrgLinks,
    getStewardMembers: getStewardMembers,
    getAllMembers: getAllMembers,
    getStewardSurveyTracking: getStewardSurveyTracking,
    // Drive contact log helpers — exposed for use by top-level wrappers (e.g. dataSendDirectMessage)
    getOrCreateMemberContactFolderPublic: getOrCreateMemberContactFolder_,
    getOrCreateMemberContactSheetPublic: getOrCreateMemberContactSheet_,
    sendBroadcastMessage: sendBroadcastMessage,
    startGrievanceDraft: startGrievanceDraft,
    createGrievanceDriveFolder: createGrievanceDriveFolder,
    getSurveyResults: getSurveyResults,
    // v4.12.0
    logMemberContact: logMemberContact,
    getMemberContactHistory: getMemberContactHistory,
    getStewardContactLog: getStewardContactLog,
    createTask: createTask,
    getTasks: getTasks,
    updateTask: updateTask,
    completeTask: completeTask,
    getStewardMemberStats: getStewardMemberStats,
    getStewardDirectory: getStewardDirectory,
    getGrievanceStats: getGrievanceStats,
    getGrievanceHotSpots: getGrievanceHotSpots,
    getMembershipStats: getMembershipStats,
    getUpcomingEvents: getUpcomingEvents,
    // v4.15.0
    isChiefSteward: isChiefSteward,
    getChiefStewardTaskView: getChiefStewardTaskView,
    // v4.17.0 - Member Task Assignment
    createMemberTask: createMemberTask,
    getMemberTasks: getMemberTasks,
    completeMemberTask: completeMemberTask,
    stewardCompleteMemberTask: stewardCompleteMemberTask,
    getStewardAssignedMemberTasks: getStewardAssignedMemberTasks,
    // v4.17.0 - Minutes (Polls removed v4.24.0 — use wq* wrappers)
    getMeetingMinutes: getMeetingMinutes,
    addMeetingMinutes: addMeetingMinutes,
    // v4.52.0 - Minutes helpers (exposed for BACKFILL wrapper)
    createMinutesDoc_: createMinutesDoc_,
    // v4.18.0 - Performance, Checklists, Meetings, Satisfaction, Feedback
    getStewardPerformance: getStewardPerformance,
    getAllStewardPerformance: getAllStewardPerformance,
    getCaseChecklist: getCaseChecklist,
    getCaseChecklistProgress: getCaseChecklistProgress,
    toggleChecklistItem: toggleChecklistItem,
    getMemberMeetings: getMemberMeetings,
    getMeetingStats: getMeetingStats,
    getSatisfactionTrends: getSatisfactionTrends,
    // Perf: batch + cache
    getBatchData: getBatchData,
    getBadgeCounts: getBadgeCounts,
    _invalidateSheetCache: _invalidateSheetCache,
    _resetSSCache: _resetSSCache,
    // Scale: paginated members + health monitor
    getMembersPaginated: getMembersPaginated,
    getMemberCount: getMemberCount,
    getFilterDropdownValues: getFilterDropdownValues,
    getSheetHealth: getSheetHealth,
    // v4.32.0 — Grievance Feedback
    getPendingGrievanceFeedback: getPendingGrievanceFeedback,
    submitGrievanceFeedback: submitGrievanceFeedback,
    getGrievanceFeedbackStats: getGrievanceFeedbackStats,
    getStewardFeedbackSummary: getStewardFeedbackSummary,
    // v4.46.0 — Leader Hub
    getUnitMembers: getUnitMembers,
    sendUnitBroadcast: sendUnitBroadcast,
    // v4.48.0 — Non-Member Contacts
    getNonMemberContacts: getNonMemberContacts,
    addNonMemberContact: addNonMemberContact,
    updateNonMemberContact: updateNonMemberContact,
    deleteNonMemberContact: deleteNonMemberContact,
  };

})();
