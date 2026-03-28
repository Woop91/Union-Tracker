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
    grievanceNotes:  ['notes', 'description', 'summary'],
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
      Logger.log('DataService: Steward column not found in Grievance Log');
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

      if (status === 'resolved') {
        resolved++;
      } else if (status === 'overdue') {
        overdue++;
      } else {
        active++;
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

    var closedStatuses = GRIEVANCE_CLOSED_STATUSES.map(function(s) { return s.toLowerCase(); });
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

    var closedStatuses = GRIEVANCE_CLOSED_STATUSES.map(function(s) { return s.toLowerCase(); });
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
      jobTitle: user.jobTitle,
      joined: user.joined,
      memberId: user.memberId || '',
      memberAdminFolderUrl: user.memberAdminFolderUrl || '',
      cubicle: user.cubicle || '',
      employeeId: user.employeeId || '',
      hireDate: user.hireDate || '',
      openRate: user.openRate || '',
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
      street:       HEADERS.memberStreet,
      city:         HEADERS.memberCity,
      state:        HEADERS.memberState,
      zip:          HEADERS.memberZip,
      workLocation: HEADERS.memberWorkLocation,
      officeDays:   HEADERS.memberOfficeDays,
      sharePhone:   HEADERS.memberSharePhone,   // steward opt-in: phone visible to members
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
      return { success: true, message: 'Profile updated.' };
    }

    return { success: false, message: 'Member not found.' };
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
      idemCache.put('IDEM_' + idemKey, '1', 300);
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
      Logger.log('DataService.createGrievanceDriveFolder error: ' + e.message);
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
      Logger.log('DataService.getMemberSurveyStatus error: ' + e.message);
    }
    return { hasCompleted: false, lastCompleted: null };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Resource Click Tracking (v4.31.1)
  // ═══════════════════════════════════════

  /**
   * Logs a resource click/view event. Uses ScriptProperties for durable,
   * lightweight counting. Each resource gets its own property key:
   *   RC_<resourceId> = cumulative click count
   *   RC_TOTAL        = grand total across all resources
   *
   * @param {string} email  — caller email (server-resolved)
   * @param {string} resourceId — e.g. 'RES-001' or 'quick:calendar'
   * @returns {Object} { success: boolean }
   */
  function logResourceClick(email, resourceId) {
    if (!email || !resourceId) return { success: false };
    try {
      var props = PropertiesService.getScriptProperties();
      var rcKey = 'RC_' + String(resourceId).replace(/[^A-Za-z0-9_:-]/g, '');

      // Increment per-resource counter
      var current = parseInt(props.getProperty(rcKey) || '0', 10) || 0;
      props.setProperty(rcKey, String(current + 1));

      // Increment grand total
      var total = parseInt(props.getProperty('RC_TOTAL') || '0', 10) || 0;
      props.setProperty('RC_TOTAL', String(total + 1));

      return { success: true };
    } catch (e) {
      Logger.log('logResourceClick error: ' + e.message);
      return { success: false };
    }
  }

  /**
   * Returns the total resource click count across all resources.
   * @returns {number}
   */
  function getResourceClickTotal() {
    try {
      var props = PropertiesService.getScriptProperties();
      return parseInt(props.getProperty('RC_TOTAL') || '0', 10) || 0;
    } catch (_e) {
      return 0;
    }
  }

  /**
   * Returns detailed resource usage stats for Union Stats > Resources sub-tab.
   * Reads per-resource click counts from ScriptProperties (RC_<id>) and
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
          var rId = String(resData[ri][RESOURCES_COLS.RESOURCE_ID - 1]).trim();
          var visible = String(resData[ri][RESOURCES_COLS.VISIBLE - 1]).trim().toLowerCase();
          if (rId && visible === 'yes') {
            resourceMap[rId] = {
              title: String(resData[ri][RESOURCES_COLS.TITLE - 1]).trim(),
              category: String(resData[ri][RESOURCES_COLS.CATEGORY - 1]).trim()
            };
          }
        }
      }

      // Read per-resource click counts from ScriptProperties
      var props = PropertiesService.getScriptProperties();
      var allProps = props.getProperties();
      var clickMap = {}; // resourceId -> clickCount
      var totalViews = 0;
      var uniqueViewed = 0;
      for (var key in allProps) {
        if (key.indexOf('RC_') === 0 && key !== 'RC_TOTAL') {
          var resId = key.substring(3); // strip 'RC_'
          var clicks = parseInt(allProps[key], 10) || 0;
          if (clicks > 0) {
            clickMap[resId] = clicks;
            totalViews += clicks;
            uniqueViewed++;
          }
        }
      }

      // Use RC_TOTAL as authoritative grand total (may include quick-link clicks)
      var rcTotal = parseInt(allProps['RC_TOTAL'] || '0', 10) || 0;
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
      Logger.log('getResourceStats error: ' + e.message);
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
      Logger.log('logTabVisit error: ' + e.message);
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
      Logger.log('getUsageStats error: ' + e.message);
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
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
      Logger.log('getOrCreateSheetFolder_ error: ' + e.message);
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
      if (!rec.email) continue;
      members.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        cubicle: rec.cubicle,
        hireDate: rec.hireDate,
        hasOpenGrievance: rec.hasOpenGrievance,
        duesStatus: rec.duesStatus,
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
      Logger.log('getStewardSurveyTracking: survey pre-load error: ' + e.message);
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
    var auth = checkWebAppAuthorization('steward'); if (!auth || !auth.isAuthorized) return { success: false, sentCount: 0, message: 'Unauthorized' };
    if (!stewardEmail || !message) return { success: false, sentCount: 0, message: 'Missing required fields.' };

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

    for (var i = 0; i < filtered.length; i++) {
      try {
        MailApp.sendEmail(filtered[i].email, subject, message);
        sentCount++;
      } catch (e) {
        failedCount++;
        failures.push({ email: filtered[i].email, error: e.message });
        Logger.log('Broadcast send error for ' + filtered[i].email + ': ' + e.message);
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
      Logger.log('DataService.getSurveyResults error: ' + e.message);
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
      Logger.log('DataService: getActiveSpreadsheet() returned null for sheet "' + name + '"');
      return null;
    }
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('DataService: Sheet "' + name + '" not found.');
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
      if (colMap.hasOwnProperty(key)) {
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
        var NON_PAYING = ['past due', 'inactive', 'delinquent', 'lapsed', 'non-paying', 'no'];
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
      hasOpenGrievance: hasGrievance === 'yes' || hasGrievance === 'true' || hasGrievance === '1',
      street: String(_getVal(row, colMap, HEADERS.memberStreet, '')).trim(),
      city: String(_getVal(row, colMap, HEADERS.memberCity, '')).trim(),
      state: String(_getVal(row, colMap, HEADERS.memberState, '')).trim(),
      zip: String(_getVal(row, colMap, HEADERS.memberZip, '')).trim(),
      supervisor: String(_getVal(row, colMap, HEADERS.memberSupervisor, '')).trim(),
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
      Logger.log('_getMemberAdminFolder_: getOrCreateMemberAdminFolder not available');
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
      Logger.log('getOrCreateMemberContactSheet_ error: ' + e.message);
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
    sheet.appendRow([id, stewardEmail.toLowerCase().trim(), memberEmail.toLowerCase().trim(), contactType, new Date(), (notes || '').substring(0, 500), duration || '', new Date(), resolvedName]);
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
          var emailIdx          = MEMBER_COLS.EMAIL - 1;
          var recentContactIdx  = MEMBER_COLS.RECENT_CONTACT_DATE - 1;
          var contactStewardIdx = MEMBER_COLS.CONTACT_STEWARD - 1;
          var contactNotesIdx   = MEMBER_COLS.CONTACT_NOTES - 1;
          if (emailIdx >= 0) {
            var mEmailNorm = memberEmail.toLowerCase().trim();
            for (var r = 1; r < mData.length; r++) {
              if (String(mData[r][emailIdx]).toLowerCase().trim() === mEmailNorm) {
                if (recentContactIdx  !== -1) memberDir.getRange(r + 1, recentContactIdx  + 1).setValue(new Date());
                if (contactStewardIdx !== -1) memberDir.getRange(r + 1, contactStewardIdx + 1).setValue(sName);
                if (contactNotesIdx   !== -1 && notes) memberDir.getRange(r + 1, contactNotesIdx + 1).setValue(String(notes).substring(0, 500));
                break;
              }
            }
          }
        }
      }
    } catch (wbErr) {
      Logger.log('logMemberContact writeback error: ' + wbErr.message);
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
          logSheet.appendRow([new Date(), sName, contactType, (notes || '').substring(0, 500), duration || '']);
        }
      }
    } catch (driveErr) {
      Logger.log('logMemberContact Drive sheet error: ' + driveErr.message);
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
      idemCache.put('IDEM_' + idemKey, '1', 300);
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
    sheet.appendRow([id, ownerEmail, title.substring(0, 200), (description || '').substring(0, 500), (memberEmail || '').toLowerCase().trim(), priority || 'medium', 'open', dueDate || '', new Date(), '']);
    _invalidateSheetCache(SHEETS.STEWARD_TASKS);
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
      id, memberEmail.toLowerCase().trim(), title.substring(0, 200),
      (desc || '').substring(0, 500), memberEmail.toLowerCase().trim(),
      priority || 'medium', 'open', dueDate || '', new Date(), '',
      'member', stewardEmail.toLowerCase().trim()
    ]);
    _invalidateSheetCache(SHEETS.STEWARD_TASKS);
    if (typeof logAuditEvent === 'function') logAuditEvent('MEMBER_TASK_CREATED', 'Task ' + id + ' assigned to ' + memberEmail + ' by ' + stewardEmail);
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
        Logger.log('getGrievanceStats: skipped row ' + i + ' — ' + rowErr.message);
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
      Logger.log('getUpcomingEvents error: ' + e.message);
      // Last resort: try timeline fallback
      try { var fb = _getTimelineEvents(limit); if (fb.length > 0) return fb; } catch (_e2) { Logger.log('_e2: ' + (_e2.message || _e2)); }
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
      Logger.log('SLOW_EXEC: ' + label + ' at ' + elapsed + 'ms');
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
          Logger.log('Sheet cache hit: ' + sheetName + ' (age: ' + (Date.now() - parsed._cachedAt) + 'ms)');
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
        Logger.log('SCALE_CRITICAL: ' + sheetName + ' has ' + rowCount + ' rows — manual intervention recommended');
      } else if (rowCount >= SCALE_THRESHOLDS.THROTTLE_ROWS) {
        Logger.log('SCALE_THROTTLE: ' + sheetName + ' has ' + rowCount + ' rows — paginated mode active');
      } else if (rowCount >= SCALE_THRESHOLDS.WARN_ROWS) {
        Logger.log('SCALE_WARN: ' + sheetName + ' has ' + rowCount + ' rows');
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
    } catch (_e) { Logger.log('_getCachedArchiveData cache write: ' + (_e.message || _e)); }

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
      Logger.log('_getAllGrievanceData: column count mismatch (active=' + activeHeaders.length + ', archive=' + archiveHeaders.length + ') — using active only');
      return active;
    }
    for (var h = 0; h < activeHeaders.length; h++) {
      if (String(activeHeaders[h]).trim() !== String(archiveHeaders[h]).trim()) {
        Logger.log('_getAllGrievanceData: header mismatch at col ' + (h + 1) + ' ("' + activeHeaders[h] + '" vs "' + archiveHeaders[h] + '") — using active only');
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
    try { _getCachedSheetData(GRIEVANCE_SHEET); } catch (_) { Logger.log('_: ' + (_.message || _)); }
    try { _getCachedSheetData(MEMBER_SHEET); } catch (_) { Logger.log('_: ' + (_.message || _)); }

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

    var memberTaskCount = 0;
    try {
      var openTasks = getMemberTasks(email, 'not-completed');
      memberTaskCount = openTasks.length;
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
    try { _getCachedSheetData(GRIEVANCE_SHEET); } catch (_) { Logger.log('_: ' + (_.message || _)); }
    try { _getCachedSheetData(MEMBER_SHEET); } catch (_) { Logger.log('_: ' + (_.message || _)); }
    try { _getCachedSheetData(SHEETS.STEWARD_TASKS); } catch (_) { Logger.log('_: ' + (_.message || _)); }

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, 'steward');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

    var qaUnansweredCount = 0;
    try {
      if (typeof QAForum !== 'undefined') {
        // Use lightweight count — avoids building 999 full question objects
        qaUnansweredCount = QAForum.getUnansweredCount();
      }
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
      if (status === 'resolved') resolved++;
      else if (status === 'overdue') overdue++;
      else {
        active++;
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

    var taskCount = 0;
    var overdueTaskCount = 0;
    if (role === 'steward') {
      try {
        var openTasks = getTasks(email, 'open');
        taskCount = openTasks.length;
        for (var t = 0; t < openTasks.length; t++) {
          if (openTasks[t].dueDays !== null && openTasks[t].dueDays < 0) overdueTaskCount++;
        }
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }

    var qaUnansweredCount = 0;
    if (role === 'steward') {
      try {
        if (typeof QAForum !== 'undefined') {
          qaUnansweredCount = QAForum.getUnansweredCount();
        }
      } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    }

    return {
      notificationCount: notifCount,
      taskCount: taskCount,
      overdueTaskCount: overdueTaskCount,
      qaUnansweredCount: qaUnansweredCount,
    };
  }

  /**
   * Returns feedback items submitted by a specific user, newest first.
   * @param {string} email
   * @returns {Object[]} Array of feedback records
   */
  function getMyFeedback(email) {
    if (!email) return [];
    email = String(email).trim().toLowerCase();

    var ss = _getSS();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.FEEDBACK);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][FEEDBACK_COLS.SUBMITTED_BY - 1]).trim().toLowerCase();
      if (rowEmail !== email) continue;

      var ts = data[i][FEEDBACK_COLS.TIMESTAMP - 1];
      items.push({
        date: ts instanceof Date ? _formatDate(ts) : String(ts || ''),
        category: String(data[i][FEEDBACK_COLS.CATEGORY - 1] || ''),
        type: String(data[i][FEEDBACK_COLS.TYPE - 1] || ''),
        priority: String(data[i][FEEDBACK_COLS.PRIORITY - 1] || ''),
        title: String(data[i][FEEDBACK_COLS.TITLE - 1] || ''),
        description: String(data[i][FEEDBACK_COLS.DESCRIPTION - 1] || ''),
        status: String(data[i][FEEDBACK_COLS.STATUS - 1] || 'New'),
        resolution: String(data[i][FEEDBACK_COLS.RESOLUTION - 1] || ''),
      });
    }
    items.reverse(); // newest first
    return items;
  }

  // ═══════════════════════════════════════
  // NOTE v4.24.0: Legacy FlashPolls functions (getActivePolls, submitPollVote, addPoll)
  // removed here. All poll functionality is now in 24_WeeklyQuestions.gs via wq* wrappers.
  // FlashPolls/PollResponses sheets also removed from 23_PortalSheets.gs.
  // ═══════════════════════════════════════

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
        driveDocUrl: String(data[i][PORTAL_MINUTES_COLS.DRIVE_DOC_URL] || '')  // col 8 — Google Doc in Minutes/ folder
      });
    }
    minutes.sort(function(a, b) { return (b.meetingDateTs || 0) - (a.meetingDateTs || 0); });
    minutes.forEach(function(m) { delete m.meetingDateTs; });
    return minutes.slice(0, limit);
    } catch (e) {
      Logger.log('getMeetingMinutes error: ' + e.message + '\n' + (e.stack || ''));
      return [];
    }
  }

  /**
   * Adds new meeting minutes (steward-only).
   * @param {string} stewardEmail
   * @param {Object} data - { title, meetingDate, bullets, fullMinutes }
   * @returns {Object}
   */
  function addMeetingMinutes(stewardEmail, minutesData, idemKey) {
    if (idemKey) {
      var idemCache = CacheService.getScriptCache();
      if (idemCache.get('IDEM_' + idemKey)) return { duplicate: true, message: 'Duplicate request ignored' };
      idemCache.put('IDEM_' + idemKey, '1', 300);
    }
    if (!stewardEmail || !minutesData || !minutesData.title) {
      return { success: false, message: 'Missing required fields.' };
    }
    var sheet = (typeof getOrCreateMinutesSheet === 'function') ? getOrCreateMinutesSheet() : null;
    if (!sheet) return { success: false, message: 'Minutes sheet unavailable.' };

    var id = 'MIN_' + Date.now().toString(36);
    // Append T12:00:00 to YYYY-MM-DD strings to avoid UTC midnight → previous-day shift in America/New_York
    var rawDate = minutesData.meetingDate;
    var meetingDate = rawDate
      ? new Date(typeof rawDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate + 'T12:00:00' : rawDate)
      : new Date();
    if (isNaN(meetingDate.getTime())) meetingDate = new Date();

    // ── Save Google Doc to Minutes/ Drive folder ─────────────────────────────
    // Stores a formatted Google Doc alongside the spreadsheet record so stewards
    // can share a clean link. Falls back gracefully if Drive setup hasn't run.
    var driveDocUrl = '';
    try {
      var minutesFolderId = (typeof getConfigValue_ === 'function')
        ? getConfigValue_(CONFIG_COLS.MINUTES_FOLDER_ID)
        : '';
      if (!minutesFolderId && typeof PropertiesService !== 'undefined') {
        minutesFolderId = PropertiesService.getScriptProperties().getProperty('MINUTES_FOLDER_ID') || '';
      }

      var docTitle = minutesData.title + ' — ' +
        Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      var doc = DocumentApp.create(docTitle);
      var body = doc.getBody();
      body.appendParagraph(docTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('Recorded by: ' + stewardEmail);
      body.appendParagraph('Date: ' + Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'MMMM d, yyyy'));
      body.appendParagraph('');

      if (minutesData.bullets) {
        body.appendParagraph('Key Points').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        String(minutesData.bullets).split('\n').forEach(function(line) {
          if (line.trim()) body.appendListItem(line.replace(/^[-•*]\s*/, '').trim());
        });
        body.appendParagraph('');
      }
      if (minutesData.fullMinutes) {
        body.appendParagraph('Full Minutes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(String(minutesData.fullMinutes));
      }
      doc.saveAndClose();

      // Move the doc into the Minutes/ subfolder if it exists
      var docFile = DriveApp.getFileById(doc.getId());
      if (minutesFolderId) {
        try {
          var minutesFolder = DriveApp.getFolderById(minutesFolderId);
          minutesFolder.addFile(docFile);
          DriveApp.getRootFolder().removeFile(docFile);
        } catch (moveErr) {
          Logger.log('addMeetingMinutes: could not move doc to Minutes folder: ' + moveErr.message);
        }
      }
      driveDocUrl = docFile.getUrl();
    } catch (driveErr) {
      Logger.log('addMeetingMinutes: Drive doc creation failed (non-fatal): ' + driveErr.message);
    }

    // ── Write to sheet (8 columns: add DriveDocUrl at end) ──────────────────
    sheet.appendRow([
      id,
      meetingDate,
      String(minutesData.title).substring(0, 200),
      String(minutesData.bullets || '').substring(0, 2000),
      String(minutesData.fullMinutes || '').substring(0, 5000),
      String(stewardEmail).trim().toLowerCase(),
      new Date(),
      driveDocUrl
    ]);

    // Ensure DriveDocUrl header exists in col 8
    try {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers[7]) sheet.getRange(1, 8).setValue('DriveDocUrl');
    } catch (_he) { Logger.log('_he: ' + (_he.message || _he)); }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MINUTES_ADDED', { steward: stewardEmail, title: minutesData.title, driveDocUrl: driveDocUrl });
    }
    return { success: true, message: 'Minutes added.' + (driveDocUrl ? ' Google Doc saved.' : ''), id: id, driveDocUrl: driveDocUrl };
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
        var rowEmail = String(data[i][STEWARD_PERF_COLS.STEWARD - 1]).trim().toLowerCase();
        if (rowEmail !== email) continue;
        return {
          steward: rowEmail,
          totalCases: Number(data[i][STEWARD_PERF_COLS.TOTAL_CASES - 1]) || 0,
          active: Number(data[i][STEWARD_PERF_COLS.ACTIVE - 1]) || 0,
          closed: Number(data[i][STEWARD_PERF_COLS.CLOSED - 1]) || 0,
          won: Number(data[i][STEWARD_PERF_COLS.WON - 1]) || 0,
          winRate: Number(data[i][STEWARD_PERF_COLS.WIN_RATE - 1]) || 0,
          avgDays: Number(data[i][STEWARD_PERF_COLS.AVG_DAYS - 1]) || 0,
          overdue: Number(data[i][STEWARD_PERF_COLS.OVERDUE - 1]) || 0,
          dueThisWeek: Number(data[i][STEWARD_PERF_COLS.DUE_THIS_WEEK - 1]) || 0,
          performanceScore: Number(data[i][STEWARD_PERF_COLS.PERFORMANCE_SCORE - 1]) || 0,
        };
      }
      return {};
    } catch (_e) {
      Logger.log('getStewardPerformance error: ' + _e.message);
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
        var steward = String(data[i][STEWARD_PERF_COLS.STEWARD - 1]).trim();
        if (!steward) continue;
        results.push({
          steward: steward.toLowerCase(),
          totalCases: Number(data[i][STEWARD_PERF_COLS.TOTAL_CASES - 1]) || 0,
          active: Number(data[i][STEWARD_PERF_COLS.ACTIVE - 1]) || 0,
          closed: Number(data[i][STEWARD_PERF_COLS.CLOSED - 1]) || 0,
          won: Number(data[i][STEWARD_PERF_COLS.WON - 1]) || 0,
          winRate: Number(data[i][STEWARD_PERF_COLS.WIN_RATE - 1]) || 0,
          avgDays: Number(data[i][STEWARD_PERF_COLS.AVG_DAYS - 1]) || 0,
          overdue: Number(data[i][STEWARD_PERF_COLS.OVERDUE - 1]) || 0,
          dueThisWeek: Number(data[i][STEWARD_PERF_COLS.DUE_THIS_WEEK - 1]) || 0,
          performanceScore: Number(data[i][STEWARD_PERF_COLS.PERFORMANCE_SCORE - 1]) || 0,
        });
      }
      return results;
    } catch (_e) {
      Logger.log('getAllStewardPerformance error: ' + _e.message);
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
        if (String(data[i][CHECKLIST_COLS.CASE_ID - 1]).trim() !== caseId) continue;
        var dueDate = data[i][CHECKLIST_COLS.DUE_DATE - 1];
        var completedDate = data[i][CHECKLIST_COLS.COMPLETED_DATE - 1];
        items.push({
          id: String(data[i][CHECKLIST_COLS.CHECKLIST_ID - 1] || ''),
          caseId: caseId,
          actionType: String(data[i][CHECKLIST_COLS.ACTION_TYPE - 1] || ''),
          itemText: String(data[i][CHECKLIST_COLS.ITEM_TEXT - 1] || ''),
          category: String(data[i][CHECKLIST_COLS.CATEGORY - 1] || ''),
          required: String(data[i][CHECKLIST_COLS.REQUIRED - 1]).toLowerCase() === 'true',
          completed: String(data[i][CHECKLIST_COLS.COMPLETED - 1]).toLowerCase() === 'true',
          completedBy: String(data[i][CHECKLIST_COLS.COMPLETED_BY - 1] || ''),
          completedDate: completedDate instanceof Date ? _formatDate(completedDate) : '',
          dueDate: dueDate instanceof Date ? _formatDate(dueDate) : String(dueDate || ''),
          notes: String(data[i][CHECKLIST_COLS.NOTES - 1] || ''),
          sortOrder: Number(data[i][CHECKLIST_COLS.SORT_ORDER - 1]) || 0,
        });
      }
      items.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
      return items;
    } catch (_e) {
      Logger.log('getCaseChecklist error: ' + _e.message);
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
      Logger.log('getCaseChecklistProgress error: ' + _e.message);
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
        if (String(data[i][CHECKLIST_COLS.CHECKLIST_ID - 1]).trim() !== String(checklistId).trim()) continue;
        var rowNum = i + 1;
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED).setValue(completed ? 'true' : 'false');
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED_BY).setValue(completed ? String(email || '').toLowerCase().trim() : '');
        sheet.getRange(rowNum, CHECKLIST_COLS.COMPLETED_DATE).setValue(completed ? new Date() : '');

        if (typeof logAuditEvent === 'function') {
          logAuditEvent('CHECKLIST_TOGGLED', { checklistId: checklistId, completed: completed, by: email });
        }
        return { success: true, message: completed ? 'Item completed.' : 'Item unchecked.' };
      }
      return { success: false, message: 'Checklist item not found.' };
    } catch (_e) {
      Logger.log('toggleChecklistItem error: ' + _e.message);
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
        var rowEmail = String(data[i][MEETING_CHECKIN_COLS.EMAIL - 1]).trim().toLowerCase();
        if (rowEmail !== email) continue;
        var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
        var checkinTime = data[i][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1];
        meetings.push({
          meetingId: String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || ''),
          meetingName: String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || ''),
          meetingDate: meetingDate instanceof Date ? _formatDate(meetingDate) : String(meetingDate || ''),
          meetingType: String(data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || ''),
          checkinTime: checkinTime instanceof Date ? _formatDate(checkinTime) : String(checkinTime || ''),
        });
      }
      meetings.reverse(); // newest first
      return meetings;
    } catch (_e) {
      Logger.log('getMemberMeetings error: ' + _e.message);
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
        var mid = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1]).trim();
        var aemail = String(data[i][MEETING_CHECKIN_COLS.EMAIL - 1]).trim().toLowerCase();
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
      Logger.log('getMeetingStats error: ' + _e.message);
      return { totalMeetings: 0, totalCheckins: 0, uniqueAttendees: 0, avgAttendance: 0 };
    }
  }

  // ═══════════════════════════════════════
  // Satisfaction & Feedback (v4.18.0)
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
      Logger.log('getSatisfactionTrends error: ' + _e.message);
      return { overall: 0, responseCount: 0, categories: [] };
    }
  }

  /**
   * Submits feedback from a member.
   * @param {string} email - Submitter email
   * @param {Object} data - { category, type, priority, title, description }
   * @returns {Object} { success, message }
   */
  function submitFeedback(email, feedbackData, idemKey) {
    if (idemKey) {
      var idemCache = CacheService.getScriptCache();
      if (idemCache.get('IDEM_' + idemKey)) return { duplicate: true, message: 'Duplicate request ignored' };
      idemCache.put('IDEM_' + idemKey, '1', 300);
    }
    try {
      if (!email || !feedbackData || !feedbackData.title) {
        return { success: false, message: 'Missing required fields.' };
      }
      email = String(email).trim().toLowerCase();

      var ss = _getSS();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.FEEDBACK);
      if (!sheet) return { success: false, message: 'Feedback sheet not found.' };

      var row = [
        new Date(),
        email,
        String(feedbackData.category || 'General').substring(0, 100),
        // TYPE column removed v4.24.1 — Category covers same ground
        String(feedbackData.priority || 'Medium').substring(0, 20),
        String(feedbackData.title).substring(0, 200),
        String(feedbackData.description || '').substring(0, 2000),
        'New',
        '',
        '',
        '',
      ];
      sheet.appendRow(row);

      if (typeof logAuditEvent === 'function') {
        logAuditEvent('FEEDBACK_SUBMITTED', { email: email, title: feedbackData.title.substring(0, 100) });
      }
      return { success: true, message: 'Feedback submitted.' };
    } catch (_e) {
      Logger.log('submitFeedback error: ' + _e.message);
      return { success: false, message: 'Failed to submit feedback.' };
    }
  }

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
    var closedStatuses = GRIEVANCE_CLOSED_STATUSES.map(function(s) { return s.toLowerCase(); });

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
          closedDate: c.closedTimestamp ? new Date(c.closedTimestamp).toLocaleDateString() : '',
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
    var closedStatuses = GRIEVANCE_CLOSED_STATUSES.map(function(s) { return s.toLowerCase(); });
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
      safeComment,
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
          date: created instanceof Date ? created.toLocaleDateString() : '',
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
    // v4.17.0 - Feedback, Minutes (Polls removed v4.24.0 — use wq* wrappers)
    getMyFeedback: getMyFeedback,
    getMeetingMinutes: getMeetingMinutes,
    addMeetingMinutes: addMeetingMinutes,
    // v4.18.0 - Performance, Checklists, Meetings, Satisfaction, Feedback
    getStewardPerformance: getStewardPerformance,
    getAllStewardPerformance: getAllStewardPerformance,
    getCaseChecklist: getCaseChecklist,
    getCaseChecklistProgress: getCaseChecklistProgress,
    toggleChecklistItem: toggleChecklistItem,
    getMemberMeetings: getMemberMeetings,
    getMeetingStats: getMeetingStats,
    getSatisfactionTrends: getSatisfactionTrends,
    submitFeedback: submitFeedback,
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
  };

})();
// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

/**
 * Resolves the caller's verified email server-side (never trust client-supplied email).
 * Priority: (1) Session.getActiveUser() for Google SSO, (2) sessionToken for magic link auth.
 * @param {string=} sessionToken - Optional client-supplied session token
 * @returns {string} Verified email or empty string
 * @private
 */
function _resolveCallerEmail(sessionToken) {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return email.toLowerCase().trim();
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  // Fallback: verify session token server-side — never trust plain email from client
  if (sessionToken && typeof Auth !== 'undefined' && typeof Auth.resolveEmailFromToken === 'function') {
    var tokenEmail = Auth.resolveEmailFromToken(sessionToken);
    if (tokenEmail) return tokenEmail.toLowerCase().trim();
  }
  return '';
}

/**
 * Resolves the caller's email and verifies steward role.
 * Use for all steward-only operations.
 * @param {string=} sessionToken - Optional session token for non-SSO auth
 * @returns {string|null} Steward's email if authorized, null otherwise.
 * @private
 */
function _requireStewardAuth(sessionToken) {
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
function dataGetStewardContact(sessionToken, stewardEmail) { if (_isPINSession(sessionToken)) return _PIN_RESTRICTED_RESPONSE; var e = _resolveCallerEmail(sessionToken); if (!e) { Logger.log('dataGetStewardContact: auth failed'); return null; } return DataService.getStewardContact(stewardEmail || e); }

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
function dataGetStewardSurveyTracking(sessionToken, scope) { var s = _requireStewardAuth(sessionToken); if (!s) return { total: 0, completed: 0, members: [] }; try { return DataService.getStewardSurveyTracking(s, scope); } catch (e) { Logger.log('dataGetStewardSurveyTracking error: ' + e.message + '\n' + (e.stack || '')); return { total: 0, completed: 0, members: [] }; } }
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
function dataGetMemberContactHistory(sessionToken, memberEmail) { var s = _requireStewardAuth(sessionToken); if (!s) { Logger.log('dataGetMemberContactHistory: auth failed'); return []; } return DataService.getMemberContactHistory(s, memberEmail); }
/** @param {string} sessionToken @returns {Object[]} Full contact log for a steward. Requires steward auth. */
function dataGetStewardContactLog(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) { Logger.log('dataGetStewardContactLog: auth failed'); return []; } return DataService.getStewardContactLog(s); }

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
      Logger.log('dataSendDirectMessage Drive log error: ' + driveErr.message);
      // Non-fatal — email already sent
    }

    return { success: true, message: 'Message sent.' };
  } catch (e) {
    Logger.log('dataSendDirectMessage error: ' + e.message);
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
    Logger.log('dataGetMemberCaseFolderUrl error: ' + e.message);
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
function dataGetStewardMemberStats(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return {}; try { return DataService.getStewardMemberStats(e); } catch (err) { Logger.log('dataGetStewardMemberStats error: ' + err.message + '\n' + (err.stack || '')); return { total: 0, byLocation: {}, byDues: {} }; } }
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
    Logger.log('dataGetStewardDirectory error: ' + err.message + '\n' + (err.stack || ''));
    return [];
  }
}
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
  try { result.stats = DataService.getGrievanceStats(); } catch (_e) { result.stats = { available: false }; Logger.log('InsightsBatch stats: ' + _e.message); }
  try { result.hotSpots = DataService.getGrievanceHotSpots(); } catch (_e) { result.hotSpots = []; Logger.log('InsightsBatch hotSpots: ' + _e.message); }
  try { result.perf = DataService.getAllStewardPerformance(); } catch (_e) { result.perf = []; Logger.log('InsightsBatch perf: ' + _e.message); }
  try { result.sat = DataService.getSatisfactionTrends(); } catch (_e) { result.sat = { categories: [] }; Logger.log('InsightsBatch sat: ' + _e.message); }
  try { result.memberStats = DataService.getMembershipStats(); } catch (_e) { result.memberStats = null; Logger.log('InsightsBatch memberStats: ' + _e.message); }
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
  try { kpis = DataService.getStewardKPIs(s); } catch (_e) { Logger.log('RefreshNavData kpis: ' + _e.message); }
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
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
  } catch (_re) { Logger.log('_re: ' + (_re.message || _re)); }

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
        } catch (_mv) { Logger.log('_mv: ' + (_mv.message || _mv)); }
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
      Logger.log('BACKFILL_MINUTES_DRIVE_DOCS row ' + i + ': ' + docErr.message);
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
  Logger.log('BACKFILL_MINUTES_DRIVE_DOCS: processed=' + processed + ' skipped=' + skipped + ' errors=' + errors);
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
    Logger.log('Auto-init sheets warning: ' + err.message);
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
      Logger.log('Auto-init ' + pair[0] + ' failed: ' + err.message);
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    return {
      locations: Object.keys(locations).sort(),
      officeDays: Object.keys(officeDays).sort(),
      hasDuesPayingColumn: hasDuesPayingColumn,
      broadcastScopeAll: broadcastScopeAll,
      totalMembers: members.length
    };
  } catch (e) {
    Logger.log('dataGetBroadcastFilterOptions error: ' + e.message + '\n' + (e.stack || ''));
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
    Logger.log('dataLogResourceClick error: ' + e.message);
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
 *   resourceDownloads    — total resource click count from ScriptProperties (RC_TOTAL)
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
    } catch (_se) { Logger.log('_se: ' + (_se.message || _se)); }

    // ── Weekly question votes ───────────────────────────────────────────────
    var weeklyQuestionVotes = 0;
    try {
      var wqRows = _rows(SHEETS.WEEKLY_RESPONSES || '_Weekly_Responses');
      weeklyQuestionVotes = wqRows.length;
    } catch (_we) { Logger.log('_we: ' + (_we.message || _we)); }

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
    } catch (_ce) { Logger.log('_ce: ' + (_ce.message || _ce)); }

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
    } catch (_ge) { Logger.log('_ge: ' + (_ge.message || _ge)); }

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
    } catch (_cl) { Logger.log('_cl: ' + (_cl.message || _cl)); }

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
    } catch (_te) { Logger.log('_te: ' + (_te.message || _te)); }

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
    } catch (_mt) { Logger.log('_mt: ' + (_mt.message || _mt)); }

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
    } catch (_cl2) { Logger.log('_cl2: ' + (_cl2.message || _cl2)); }

    // ── Satisfaction trends ──────────────────────────────────────────────
    var satisfactionData = null;
    try {
      satisfactionData = DataService.getSatisfactionTrends();
    } catch (_sat) { Logger.log('_sat: ' + (_sat.message || _sat)); }

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
    } catch (_tk) { Logger.log('_tk: ' + (_tk.message || _tk)); }

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
    } catch (_ni) { Logger.log('_ni: ' + (_ni.message || _ni)); }

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
    } catch (_qa) { Logger.log('_qa: ' + (_qa.message || _qa)); }

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
    } catch (_sc) { Logger.log('_sc: ' + (_sc.message || _sc)); }

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
    Logger.log('dataGetEngagementStats error: ' + e.message + '\n' + (e.stack || ''));
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
    Logger.log('dataGetMyEngagementScore error: ' + err.message);
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
    Logger.log('dataGetResourceStats error: ' + e.message);
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
    Logger.log('dataLogTabVisit error: ' + err.message);
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
    Logger.log('dataGetUsageStats error: ' + e.message);
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
  try { return DataService.getMemberCount(); } catch (e) { Logger.log('dataGetMemberCount error: ' + e.message); return { total: 0, withGrievances: 0 }; }
}

/** Returns distinct filter dropdown values for the Members finder panel. Requires steward auth. */
function dataGetFilterDropdownValues(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { locations: [], units: [], stewards: [] };
  try { return DataService.getFilterDropdownValues(); } catch (e) { Logger.log('dataGetFilterDropdownValues error: ' + e.message); return { locations: [], units: [], stewards: [] }; }
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
  try { return DataService.getMembersPaginated(s, opts || {}); } catch (e) { Logger.log('dataGetMembersPaginated error: ' + e.message); return { members: [], total: 0, page: 1, pageSize: 25 }; }
}

/**
 * Returns row counts and scale status for key sheets. Requires steward auth.
 * @param {string} sessionToken
 * @returns {Object} { members: { rows, status }, grievances: { rows, status } }
 */
function dataGetSheetHealth(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, authError: true, message: 'Steward access required.' };
  try { return DataService.getSheetHealth(); } catch (e) { Logger.log('dataGetSheetHealth error: ' + e.message); return { members: { rows: 0, status: 'ok' }, grievances: { rows: 0, status: 'ok' } }; }
}

/**
 * Returns live workload summary stats from the Workload Vault sheet. Requires auth.
 * @param {string} sessionToken
 * @returns {Object|null} { avgCaseload, highCaseloadPct, submissionRate, trendDirection }
 */
function dataGetWorkloadSummaryStats(sessionToken) {
  var _caller = _resolveCallerEmail(sessionToken);
  if (!_caller) return null;
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
    } catch (_sr) { Logger.log('_sr: ' + (_sr.message || _sr)); }

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
    } catch (_td) { Logger.log('_td: ' + (_td.message || _td)); }

    return {
      avgCaseload:      Math.round(avgCaseload * 10) / 10,
      highCaseloadPct:  highCaseloadPct,
      submissionRate:   submissionRate,
      trendDirection:   trendDirection,
    };
  } catch (e) {
    Logger.log('dataGetWorkloadSummaryStats error: ' + e.message + '\n' + (e.stack || ''));
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
      { label: 'Check Deadlines', icon: '\u23F0', action: 'deadlines' },
      { label: 'Member Directory', icon: '\uD83D\uDC65', action: 'members' },
      { label: 'Manage Tasks', icon: '\u2705', action: 'tasks' },
    ];
  } else {
    quickActions = [
      { label: 'View My Cases', icon: '\uD83D\uDCCB', action: 'cases' },
      { label: 'Update Contact Info', icon: '\uD83D\uDC64', action: 'profile' },
      { label: 'Check Resources', icon: '\uD83D\uDCDA', action: 'resources' },
      { label: 'Contact ' + (ConfigReader.getConfig().stewardLabel || 'Steward'), icon: '\uD83D\uDCAC', action: 'contact' },
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
    Logger.log('dataSetDefaultView error: ' + err.message);
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
    Logger.log('dataGetWebAppSearchResults error: ' + e.message);
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
    Logger.log('dataUndoToIndex error: ' + e.message);
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
    Logger.log('dataExportUndoHistory error: ' + e.message);
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
    Logger.log('dataGetUndoHistory error: ' + e.message);
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
    Logger.log('dataWebCheckInMember error: ' + e.message);
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
    Logger.log('dataGetDeadlineCalendarData error: ' + e.message);
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
    Logger.log('dataGetCorrelationAlerts error: ' + e.message);
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
    Logger.log('dataGetCorrelationSummary error: ' + e.message);
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
      Logger.log('Bulk email failed for ' + memberEmails[i] + ': ' + e.message);
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
  if (!e) return [];
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
    Logger.log('dataLogUsageEvents error: ' + err.message);
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
    Logger.log('dataGetUsageAnalytics error: ' + err.message);
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
    } catch (_e) { Logger.log('_getActiveMeetingForCheckIn alreadyCheckedIn check: ' + _e.message); }

    return {
      meetingId: meeting.id,
      meetingName: meeting.name,
      meetingTime: meeting.time || '',
      meetingType: meeting.type || '',
      alreadyCheckedIn: alreadyCheckedIn
    };
  } catch (_e) {
    Logger.log('_getActiveMeetingForCheckIn error: ' + _e.message);
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
