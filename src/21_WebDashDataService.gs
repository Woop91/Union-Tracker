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

var DataService = (function () {

  // Sheet names — read from SHEETS constants if available, fallback to defaults
  var MEMBER_SHEET = (typeof SHEETS !== 'undefined' && SHEETS.MEMBER_DIR) ? SHEETS.MEMBER_DIR : 'Member Directory';
  var GRIEVANCE_SHEET = (typeof SHEETS !== 'undefined' && SHEETS.GRIEVANCE_LOG) ? SHEETS.GRIEVANCE_LOG : 'Grievance Log';

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
    memberJoined:    ['joined', 'join date', 'member since', 'date joined'],
    memberDuesStatus:['dues status', 'dues', 'status'],
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

    // Grievance Log
    grievanceId:     ['grievance id', 'id', 'case id', 'gr id'],
    grievanceMemberEmail: ['member email', 'email', 'filed by email', 'grievant email'],
    grievanceMemberFirstName: ['first name', 'first'],
    grievanceMemberLastName:  ['last name', 'last'],
    grievanceStatus: ['status', 'grievance status', 'case status'],
    grievanceStep:   ['step', 'current step', 'grievance step'],
    grievanceDeadline: ['deadline', 'next deadline', 'due date'],
    grievanceFiled:  ['filed', 'filed date', 'date filed', 'created'],
    grievanceSteward:['assigned steward', 'steward', 'steward email', 'assigned to'],
    grievanceUnit:   ['unit', 'workplace unit', 'work location', 'location'],
    grievancePriority: ['priority', 'urgency'],
    grievanceNotes:  ['notes', 'description', 'summary'],
    grievanceIssueCategory: ['issue category', 'category', 'issue type'],
    grievanceResolution: ['resolution', 'outcome', 'result'],
    grievanceDateClosed: ['date closed', 'closed date', 'closed', 'resolved date'],
  };

  // ═══════════════════════════════════════
  // PUBLIC: User Lookup
  // ═══════════════════════════════════════

  /**
   * Finds a user in the Member Directory by email.
   * Returns sanitized user record or null.
   * @param {string} email
   * @returns {Object|null}
   */
  function findUserByEmail(email) {
    if (!email) return null;
    email = String(email).trim().toLowerCase();

    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return null;

    var data = cached.data;
    var colMap = cached.colMap;

    var emailCol = _findColumn(colMap, HEADERS.memberEmail);
    if (emailCol === -1) return null;

    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][emailCol]).trim().toLowerCase();
      if (rowEmail === email) {
        return _buildUserRecord(data[i], colMap);
      }
    }

    return null;
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
        if (c.deadlineDays !== null && c.deadlineDays <= 7 && c.deadlineDays > 0) {
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

    var grievances = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][memberCol]).trim().toLowerCase();
      if (rowEmail === memberEmail) {
        grievances.push(_buildGrievanceRecord(data[i], colMap));
      }
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

    var cached = _getCachedSheetData(GRIEVANCE_SHEET);
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
   * @param {string} email
   * @returns {string|null} Drive folder URL or null
   */
  function getMemberGrievanceDriveUrl(email) {
    // Convention: folder named by grievance ID in the org shared drive
    var grievances = getMemberGrievances(email);
    if (!grievances || grievances.length === 0) return null;
    var activeG = grievances.find(function(g) { return g.status !== 'resolved'; });
    if (!activeG) return null;
    // Drive URL would come from a column or Drive API lookup.
    // For now, return null — steward can provide the link.
    return null;
  }

  /**
   * Creates a grievance draft for a member (member-initiated).
   * Writes to the Grievance Log with status 'Draft' so steward can review.
   * @param {string} email — member email (server-resolved)
   * @param {Object} data — { title, category, description }
   * @returns {Object} { success: boolean, message: string }
   */
  function startGrievanceDraft(email, data) {
    if (!email || !data) return { success: false, message: 'Missing required data.' };
    if (!data.title || !data.description) return { success: false, message: 'Title and description are required.' };

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG) || ss.getSheetByName('Grievance Log');
      if (!sheet) return { success: false, message: 'Grievance sheet not found.' };

      // Resolve member identity dynamically from directory
      var memberId = '';
      var memberFirstName = '';
      var memberLastName = '';
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR) || ss.getSheetByName('Member Directory');
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
      Logger.log('DataService.startGrievanceDraft error: ' + e.message);
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG) || ss.getSheetByName('Grievance Log');
      if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'No grievances found.' };

      var emailLower = String(email).toLowerCase().trim();

      // ── Resolve memberId dynamically from Member Directory ────────────────
      var memberId = '';
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR) || ss.getSheetByName('Member Directory');
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
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
      } catch (_e) {
        // Stored ID invalid, fall through to create
      }
    }

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
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
        surveyFormUrl: config.satisfactionFormUrl || '',
        orgWebsite: config.orgWebsite || '',
      };
    } catch (_e) {
      return { calendarUrl: '', driveFolderUrl: '', surveyFormUrl: '', orgWebsite: '' };
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
  function getAllMembers() {
    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
    var members = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      if (!rec.email) continue;
      var stewardCol = _findColumn(colMap, HEADERS.memberAssignedSteward);
      members.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
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

    // Pre-load _Survey_Tracking once and build an email → status map
    // to avoid one sheet read per member (N+1 pattern).
    var surveyMap = {};
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) throw new Error('Spreadsheet unavailable');
      var trackSheet = ss.getSheetByName(HIDDEN_SHEETS.SURVEY_TRACKING);
      if (trackSheet && trackSheet.getLastRow() > 1) {
        var tData = trackSheet.getDataRange().getValues();
        var tColMap = _buildColumnMap(tData[0]);
        var tEmailCol = _findColumn(tColMap, ['email', 'member email']);
        var tStatusCol = _findColumn(tColMap, ['status', 'completion status', 'completed']);
        var tDateCol = _findColumn(tColMap, ['completed date', 'date completed', 'completion date']);
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
          }
        }
      }
    } catch (e) {
      Logger.log('getStewardSurveyTracking: survey pre-load error: ' + e.message);
    }

    var tracking = [];
    var completedCount = 0;

    for (var i = 0; i < members.length; i++) {
      var memberEmail = String(members[i].email || '').trim().toLowerCase();
      var status = surveyMap[memberEmail] || { hasCompleted: false, lastCompleted: null };
      tracking.push({
        name: members[i].name,
        email: members[i].email,
        completed: status.hasCompleted,
        lastCompleted: status.lastCompleted,
      });
      if (status.hasCompleted) completedCount++;
    }

    return { total: members.length, completed: completedCount, members: tracking };
  }

  /**
   * Sends a broadcast email to filtered members assigned to a steward.
   * @param {string} stewardEmail - The steward sending the broadcast
   * @param {Object} filter - { location, officeDays }
   * @param {string} message - The message body
   * @returns {Object} { success, sentCount, message }
   */
  function sendBroadcastMessage(stewardEmail, filter, message) {
    var auth = checkWebAppAuthorization('steward'); if (!auth || !auth.isAuthorized) return { success: false, sentCount: 0, message: 'Unauthorized' };
    if (!stewardEmail || !message) return { success: false, sentCount: 0, message: 'Missing required fields.' };

    // Use all members when no specific steward filter, otherwise use assigned members
    var members = getAllMembers();
    if (members.length === 0) return { success: false, sentCount: 0, message: 'No members found.' };

    // Apply filters
    var filtered = members.filter(function(m) {
      if (filter && filter.location && m.workLocation) {
        if (m.workLocation.toLowerCase() !== filter.location.toLowerCase()) return false;
      }
      if (filter && filter.officeDays && m.officeDays) {
        var filterDays = filter.officeDays.toLowerCase().split(',').map(function(d) { return d.trim(); });
        var memberDays = m.officeDays.toLowerCase();
        var hasMatch = filterDays.some(function(d) { return memberDays.indexOf(d) !== -1; });
        if (!hasMatch) return false;
      }
      return true;
    });

    if (filtered.length === 0) return { success: false, sentCount: 0, message: 'No members match the filter.' };

    var sentCount = 0;
    var config = ConfigReader.getConfig();
    var subject = config.orgAbbrev + ' - Message from your ' + config.stewardLabel;

    for (var i = 0; i < filtered.length; i++) {
      try {
        MailApp.sendEmail(filtered[i].email, subject, message);
        sentCount++;
      } catch (e) {
        Logger.log('Broadcast send error for ' + filtered[i].email + ': ' + e.message);
      }
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('BROADCAST_SENT', {
        steward: stewardEmail,
        recipientCount: sentCount,
        filter: JSON.stringify(filter || {}),
      });
    }

    return { success: true, sentCount: sentCount, message: 'Sent to ' + sentCount + ' member(s).' };
  }

  /**
   * Returns aggregated quarterly survey results.
   * Privacy threshold: only returns data if 30+ responses.
   * @returns {Object} { available, count, threshold, sections }
   */
  function getSurveyResults() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  function _getSheet(name) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
   * Safe getter — returns cell value or default
   */
  function _getVal(row, colMap, aliases, defaultVal) {
    var col = _findColumn(colMap, aliases);
    if (col === -1 || col >= row.length) return defaultVal !== undefined ? defaultVal : '';
    var val = row[col];
    return (val === null || val === undefined) ? (defaultVal !== undefined ? defaultVal : '') : val;
  }

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
      if (role === 'member') role = 'steward';  // Upgrade role if only 'member'
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
    };
  }

  /**
   * Finds a user by full name (case-insensitive).
   * Used as fallback when assignedSteward contains a name instead of email.
   * @param {string} name - Full name to search for
   * @returns {Object|null} User record or null
   */
  function _findUserByName(name) {
    if (!name) return null;
    name = String(name).trim().toLowerCase();

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
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
    };
  }

  function _formatDate(date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  }

  // ═══════════════════════════════════════
  // PUBLIC: Contact Log (v4.12.0)
  // ═══════════════════════════════════════

  function _ensureContactLog() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.CONTACT_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.CONTACT_LOG);
      sheet.getRange(1, 1, 1, 8).setValues([['ID', 'Steward Email', 'Member Email', 'Contact Type', 'Date', 'Notes', 'Duration', 'Created']]);
      sheet.hideSheet();
    }
    return sheet;
  }

  function logMemberContact(stewardEmail, memberEmail, contactType, notes, duration) {
    if (!stewardEmail || !memberEmail || !contactType) return { success: false, message: 'Missing fields.' };
    var sheet = _ensureContactLog();
    var id = 'CL_' + Date.now().toString(36);
    sheet.appendRow([id, stewardEmail.toLowerCase().trim(), memberEmail.toLowerCase().trim(), contactType, new Date(), (notes || '').substring(0, 500), duration || '', new Date()]);
    if (typeof logAuditEvent === 'function') logAuditEvent('CONTACT_LOG', { steward: stewardEmail, member: memberEmail, type: contactType });

    // Writeback: update Member Directory snapshot columns so WorkloadTracker/KPIs stay current
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        var memberDir = ss.getSheetByName(MEMBER_SHEET);
        if (memberDir) {
          var mData = memberDir.getDataRange().getValues();
          var mHeaders = mData[0];
          var emailIdx = -1, recentContactIdx = -1, contactStewardIdx = -1, contactNotesIdx = -1;
          for (var h = 0; h < mHeaders.length; h++) {
            var hLow = String(mHeaders[h]).toLowerCase().trim();
            if (hLow === 'email') emailIdx = h;
            else if (hLow === 'recent contact date') recentContactIdx = h;
            else if (hLow === 'contact steward') contactStewardIdx = h;
            else if (hLow === 'contact notes') contactNotesIdx = h;
          }
          if (emailIdx !== -1) {
            var mEmailNorm = memberEmail.toLowerCase().trim();
            for (var r = 1; r < mData.length; r++) {
              if (String(mData[r][emailIdx]).toLowerCase().trim() === mEmailNorm) {
                // Resolve steward display name — Contact Steward stores name, not email
                var sRecord = (typeof findUserByEmail === 'function') ? findUserByEmail(stewardEmail) : null;
                var sName = (sRecord && sRecord.name) ? sRecord.name : stewardEmail;
                if (recentContactIdx !== -1)   memberDir.getRange(r + 1, recentContactIdx + 1).setValue(new Date());
                if (contactStewardIdx !== -1)  memberDir.getRange(r + 1, contactStewardIdx + 1).setValue(sName);
                if (contactNotesIdx !== -1 && notes) memberDir.getRange(r + 1, contactNotesIdx + 1).setValue(String(notes).substring(0, 500));
                break;
              }
            }
          }
        }
      }
    } catch (wbErr) {
      Logger.log('logMemberContact writeback error: ' + wbErr.message);
      // Non-fatal — contact log row was already written; snapshot update failed silently
    }

    return { success: true, message: 'Contact logged.' };
  }

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
          notes: data[i][5], duration: data[i][6] });
      }
    }
    // H-16: sort by numeric timestamp — string comparison of formatted dates is not chronological
    results.sort(function(a, b) { return (b._ts || 0) - (a._ts || 0); });
    results.forEach(function(r) { delete r._ts; });
    return results;
  }

  function getStewardContactLog(stewardEmail) {
    var sheet = _ensureContactLog();
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === sEmail) {
        var rawDate2 = data[i][4];
        results.push({ id: data[i][0], memberEmail: data[i][2], type: data[i][3],
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

  function _ensureStewardTasks() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.STEWARD_TASKS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.STEWARD_TASKS);
      sheet.getRange(1, 1, 1, 10).setValues([['ID', 'Steward Email', 'Title', 'Description', 'Member Email', 'Priority', 'Status', 'Due Date', 'Created', 'Completed']]);
      sheet.hideSheet();
    }
    return sheet;
  }

  function createTask(stewardEmail, title, description, memberEmail, priority, dueDate, assignToEmail) {
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

  function createMemberTask(stewardEmail, memberEmail, title, desc, priority, dueDate) {
    if (!stewardEmail || !memberEmail || !title) return { success: false, message: 'Missing required fields.' };
    var sheet = _ensureStewardTasks();
    // Migrate headers if needed
    var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 12)).getValues()[0];
    if (!headers[10] || String(headers[10]).trim() !== 'Assignee Type') {
      sheet.getRange(1, 11).setValue('Assignee Type');
      sheet.getRange(1, 12).setValue('Assigned By');
    }
    var id = 'MT_' + Date.now().toString(36);
    sheet.appendRow([
      id, memberEmail.toLowerCase().trim(), title.substring(0, 200),
      (desc || '').substring(0, 500), memberEmail.toLowerCase().trim(),
      priority || 'medium', 'open', dueDate || '', new Date(), '',
      'member', stewardEmail.toLowerCase().trim()
    ]);
    _invalidateSheetCache(SHEETS.STEWARD_TASKS);
    logAuditEvent('MEMBER_TASK_CREATED', 'Task ' + id + ' assigned to ' + memberEmail + ' by ' + stewardEmail);
    return { success: true, message: 'Task assigned to member.', taskId: id };
  }

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
      if (statusFilter && status !== statusFilter) continue;
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
        logAuditEvent('MEMBER_TASK_COMPLETED', 'Task ' + taskId + ' completed by ' + memberEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
  }

  function stewardCompleteMemberTask(stewardEmail, taskId) {
    // Allows the assigning steward to mark a member task complete on the member's behalf.
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
        logAuditEvent('MEMBER_TASK_COMPLETED_BY_STEWARD', 'Task ' + taskId + ' marked complete by steward ' + stewardEmail);
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found or not yours to complete.' };
  }

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
        var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  function getStewardDirectory() {
    var cached = _getCachedSheetData(MEMBER_SHEET);
    if (!cached) return [];
    var data = cached.data;
    var colMap = cached.colMap;
    var stewards = [];
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      if (!rec.isSteward) continue;
      stewards.push({
        name: rec.name,
        email: rec.email,
        workLocation: rec.workLocation,
        officeDays: rec.officeDays,
        phone: rec.phone,
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

  function getGrievanceStats() {
    var cached = _getCachedSheetData(GRIEVANCE_SHEET);
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
    var openCount = 0, wonCount = 0, deniedCount = 0, settledCount = 0;
    var overdueCount = 0, dueSoonCount = 0;
    for (var i = 1; i < data.length; i++) {
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
      else if (s !== 'resolved' && s !== 'withdrawn' && s !== 'closed') openCount++;
      // Deadline-based counts for org KPI cards
      if (s === 'overdue') {
        overdueCount++;
      } else if (rec.deadlineDays !== null && rec.deadlineDays > 0 && rec.deadlineDays <= 7 &&
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
    }
    // Merge month keys from both maps, sort, take last 12
    var allMonths = {};
    Object.keys(monthly).forEach(function(k) { allMonths[k] = true; });
    Object.keys(monthlyResolved).forEach(function(k) { allMonths[k] = true; });
    var sortedMonths = Object.keys(allMonths).sort().slice(-12);
    var monthlyArr = sortedMonths.map(function(k) { return { month: k, count: monthly[k] || 0 }; });
    var monthlyResolvedArr = sortedMonths.map(function(k) { return { month: k, count: monthlyResolved[k] || 0 }; });
    return {
      available: true, total: total,
      byStatus: byStatus, byStep: byStep, byUnit: byUnit, byCategory: byCategory,
      monthly: monthlyArr, monthlyResolved: monthlyResolvedArr,
      openCount: openCount, wonCount: wonCount, deniedCount: deniedCount, settledCount: settledCount,
      overdueCount: overdueCount, dueSoonCount: dueSoonCount,
    };
  }

  function getGrievanceHotSpots() {
    var cached = _getCachedSheetData(GRIEVANCE_SHEET);
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
    for (var i = 1; i < data.length; i++) {
      var rec = _buildUserRecord(data[i], colMap);
      var unit = rec.unit || 'Unknown';
      byUnit[unit] = (byUnit[unit] || 0) + 1;
      var loc = rec.workLocation || 'Unknown';
      byLocation[loc] = (byLocation[loc] || 0) + 1;
      var dues = rec.duesStatus || 'Unknown';
      byDues[dues] = (byDues[dues] || 0) + 1;
    }
    return { available: true, total: total, byUnit: byUnit, byLocation: byLocation, byDues: byDues };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Upcoming Events via CalendarApp (v4.12.0)
  // ═══════════════════════════════════════

  function getUpcomingEvents(limit) {
    limit = limit || 10;
    try {
      var config = ConfigReader.getConfig();
      if (!config.calendarId) return { _notConfigured: true, events: [] };
      var cache = CacheService.getScriptCache();
      var cacheKey = 'events_' + config.calendarId;
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);

      var cal = CalendarApp.getCalendarById(config.calendarId);
      if (!cal) return [];
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
      cache.put(cacheKey, JSON.stringify(result), 900);
      return result;
    } catch (e) {
      Logger.log('getUpcomingEvents error: ' + e.message);
      return [];
    }
  }

  // ═══════════════════════════════════════
  // SHEET DATA CACHE — avoids redundant full-sheet reads within a single
  // server execution (multiple methods reading the same sheet).
  // Uses CacheService with a short TTL for cross-request caching.
  // ═══════════════════════════════════════

  var _sheetDataCache = {};
  var SHEET_CACHE_TTL = 120; // 2 minutes

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
      var cached = cache.get(cacheKey);
      if (cached) {
        var parsed = JSON.parse(cached);
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
    } catch (_e) { /* cache miss or parse error — read from sheet */ }

    var sheet = _getSheet(sheetName);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
    var result = { data: data, colMap: colMap };

    // Store in memory
    _sheetDataCache[sheetName] = result;

    // Store in CacheService (serialize dates as ISO strings for JSON)
    try {
      var serializable = data.map(function(row) {
        return row.map(function(cell) {
          return cell instanceof Date ? { __d: cell.toISOString() } : cell;
        });
      });
      var json = JSON.stringify({ data: serializable });
      // CacheService limit is 100KB per key — skip if too large
      if (json.length < 95000) {
        cache.put(cacheKey, json, SHEET_CACHE_TTL);
      }
    } catch (_e) { /* non-fatal — in-memory cache still works */ }

    return result;
  }

  /**
   * Invalidates the sheet data cache for a specific sheet.
   * Call after writes to ensure fresh reads.
   */
  function _invalidateSheetCache(sheetName) {
    delete _sheetDataCache[sheetName];
    try {
      var cacheKey = 'SD_' + sheetName.replace(/\s/g, '_');
      CacheService.getScriptCache().remove(cacheKey);
    } catch (_e) { /* non-fatal */ }
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
    if (!email) return {};
    email = String(email).trim().toLowerCase();

    if (role === 'steward') {
      return _getStewardBatchData(email);
    }
    return _getMemberBatchData(email);
  }

  function _getMemberBatchData(email) {
    var grievances = getMemberGrievances(email);
    var history = getMemberGrievanceHistory(email);
    var stewardInfo = getAssignedStewardInfo(email);
    var surveyStatus = getMemberSurveyStatus(email);
    var events = getUpcomingEvents(5);
    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, 'member');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { /* non-fatal */ }

    return {
      grievances: grievances,
      history: history,
      stewardInfo: stewardInfo,
      surveyStatus: surveyStatus,
      events: events,
      notificationCount: notifCount,
    };
  }

  function _getStewardBatchData(email) {
    // Read cases once and compute KPIs from same data (avoids double sheet read)
    var cases = getStewardCases(email);
    var kpis = _computeKPIsFromCases(cases);

    // Member counts — Member Directory already cached from getStewardCases call above.
    // getStewardMembers falls back to getAllMembers() when no members are assigned,
    // so memberCount is always meaningful.
    var memberCount = 0;
    try {
      memberCount = getStewardMembers(email).length;
      if (memberCount === 0) memberCount = getAllMembers().length;
    } catch (_e) { /* non-fatal */ }

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
    } catch (_e) { /* non-fatal */ }

    var notifCount = 0;
    try {
      if (typeof getWebAppNotificationCount === 'function') {
        var nc = getWebAppNotificationCount(email, 'steward');
        notifCount = (nc && nc.count) || 0;
      }
    } catch (_e) { /* non-fatal */ }

    return {
      cases: cases,
      kpis: kpis,
      memberCount: memberCount,
      taskCount: taskCount,
      overdueTaskCount: overdueTaskCount,
      notificationCount: notifCount,
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
        if (cases[i].deadlineDays !== null && cases[i].deadlineDays <= 7 && cases[i].deadlineDays > 0) dueSoon++;
      }
    }
    return { totalCases: total, activeCases: active, overdue: overdue, dueSoon: dueSoon, resolved: resolved };
  }

  function getMyFeedback(email) {
    if (!email) return [];
    email = String(email).trim().toLowerCase();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  // PUBLIC: Polls (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns active polls with vote status for the user.
   * @param {string} email
   * @returns {Object[]}
   */
  function getActivePolls(email) {
    try {
    email = email ? String(email).trim().toLowerCase() : '';

    var pollSheet = (typeof getOrCreatePollsSheet === 'function') ? getOrCreatePollsSheet() : null;
    if (!pollSheet || pollSheet.getLastRow() <= 1) return [];

    var respSheet = (typeof getOrCreatePollResponsesSheet === 'function') ? getOrCreatePollResponsesSheet() : null;

    // Build response index
    var myVotes = {};
    var voteCounts = {};
    if (respSheet && respSheet.getLastRow() > 1) {
      var respData = respSheet.getDataRange().getValues();
      for (var r = 1; r < respData.length; r++) {
        var pollId = String(respData[r][PORTAL_POLL_RESPONSE_COLS.POLL_ID]);
        var voter = String(respData[r][PORTAL_POLL_RESPONSE_COLS.EMAIL]).trim().toLowerCase();
        var resp = String(respData[r][PORTAL_POLL_RESPONSE_COLS.RESPONSE]);

        if (!voteCounts[pollId]) voteCounts[pollId] = {};
        voteCounts[pollId][resp] = (voteCounts[pollId][resp] || 0) + 1;

        if (voter === email) myVotes[pollId] = resp;
      }
    }

    var pollData = pollSheet.getDataRange().getValues();
    var polls = [];
    for (var i = 1; i < pollData.length; i++) {
      var active = String(pollData[i][PORTAL_POLL_COLS.ACTIVE]).trim().toLowerCase();
      if (active !== 'true' && active !== 'yes') continue;

      var id = String(pollData[i][PORTAL_POLL_COLS.ID]);
      var optionsRaw = String(pollData[i][PORTAL_POLL_COLS.OPTIONS] || '');
      var options = optionsRaw.split(',').map(function(o) { return o.trim(); }).filter(function(o) { return o; });

      var results = {};
      var totalVotes = 0;
      options.forEach(function(o) {
        var c = (voteCounts[id] && voteCounts[id][o]) || 0;
        results[o] = c;
        totalVotes += c;
      });

      polls.push({
        id: id,
        question: String(pollData[i][PORTAL_POLL_COLS.QUESTION] || ''),
        options: options,
        hasVoted: myVotes.hasOwnProperty(id),
        myVote: myVotes[id] || null,
        results: results,
        totalVotes: totalVotes,
        unit: String(pollData[i][PORTAL_POLL_COLS.UNIT] || ''),
      });
    }
    return polls;
    } catch (e) {
      Logger.log('getActivePolls error: ' + e.message + '\n' + (e.stack || ''));
      return [];
    }
  }

  /**
   * Submits a poll vote, guarding against double-voting.
   * @param {string} email
   * @param {string} pollId
   * @param {string} response
   * @returns {Object}
   */
  function submitPollVote(email, pollId, response) {
    if (!email || !pollId || !response) return { success: false, message: 'Missing fields.' };
    email = String(email).trim().toLowerCase();

    var respSheet = (typeof getOrCreatePollResponsesSheet === 'function') ? getOrCreatePollResponsesSheet() : null;
    if (!respSheet) return { success: false, message: 'Poll system unavailable.' };

    // Check for existing vote
    if (respSheet.getLastRow() > 1) {
      var data = respSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][PORTAL_POLL_RESPONSE_COLS.POLL_ID]) === pollId &&
            String(data[i][PORTAL_POLL_RESPONSE_COLS.EMAIL]).trim().toLowerCase() === email) {
          return { success: false, message: 'Already voted on this poll.' };
        }
      }
    }

    respSheet.appendRow([pollId, email, response, new Date()]);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_VOTE', { email: email, pollId: pollId });
    }
    return { success: true, message: 'Vote recorded.' };
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
  function addMeetingMinutes(stewardEmail, minutesData) {
    if (!stewardEmail || !minutesData || !minutesData.title) {
      return { success: false, message: 'Missing required fields.' };
    }
    var sheet = (typeof getOrCreateMinutesSheet === 'function') ? getOrCreateMinutesSheet() : null;
    if (!sheet) return { success: false, message: 'Minutes sheet unavailable.' };

    var id = 'MIN_' + Date.now().toString(36);
    var meetingDate = minutesData.meetingDate ? new Date(minutesData.meetingDate) : new Date();
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
    } catch (_he) {}

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MINUTES_ADDED', { steward: stewardEmail, title: minutesData.title, driveDocUrl: driveDocUrl });
    }
    return { success: true, message: 'Minutes added.' + (driveDocUrl ? ' Google Doc saved.' : ''), id: id, driveDocUrl: driveDocUrl };
  }

  /**
   * Creates a new poll (steward-only).
   * @param {string} stewardEmail
   * @param {string} question
   * @param {string} options - comma-separated
   * @param {string} unit - target unit or empty for all
   * @returns {Object}
   */
  function addPoll(stewardEmail, question, options, unit) {
    if (!stewardEmail || !question || !options) return { success: false, message: 'Missing fields.' };
    var sheet = (typeof getOrCreatePollsSheet === 'function') ? getOrCreatePollsSheet() : null;
    if (!sheet) return { success: false, message: 'Polls sheet unavailable.' };

    var id = 'POLL_' + Date.now().toString(36);
    sheet.appendRow([id, question.substring(0, 500), options.substring(0, 500), 'true', (unit || '').substring(0, 100), String(stewardEmail).trim().toLowerCase(), new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_CREATED', { steward: stewardEmail, question: question.substring(0, 100) });
    }
    return { success: true, message: 'Poll created.', id: id };
  }

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

      var ss = SpreadsheetApp.getActiveSpreadsheet();
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
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

      var ss = SpreadsheetApp.getActiveSpreadsheet();
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

      var ss = SpreadsheetApp.getActiveSpreadsheet();
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

      var ss = SpreadsheetApp.getActiveSpreadsheet();
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
      var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  function getSatisfactionTrends() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return {};
      var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
      if (!sheet || sheet.getLastRow() <= 1) return {};

      var data = sheet.getDataRange().getValues();
      // Read from summary area if available (pre-computed by sheet formulas)
      if (data.length > 1 && data[1][SATISFACTION_COLS.AVG_OVERALL_SAT - 1]) {
        return {
          overallSat: Number(data[1][SATISFACTION_COLS.AVG_OVERALL_SAT - 1]) || 0,
          stewardRating: Number(data[1][SATISFACTION_COLS.AVG_STEWARD_RATING - 1]) || 0,
          stewardAccess: Number(data[1][SATISFACTION_COLS.AVG_STEWARD_ACCESS - 1]) || 0,
          chapter: Number(data[1][SATISFACTION_COLS.AVG_CHAPTER - 1]) || 0,
          leadership: Number(data[1][SATISFACTION_COLS.AVG_LEADERSHIP - 1]) || 0,
          contract: Number(data[1][SATISFACTION_COLS.AVG_CONTRACT - 1]) || 0,
          representation: Number(data[1][SATISFACTION_COLS.AVG_REPRESENTATION - 1]) || 0,
          communication: Number(data[1][SATISFACTION_COLS.AVG_COMMUNICATION - 1]) || 0,
          memberVoice: Number(data[1][SATISFACTION_COLS.AVG_MEMBER_VOICE - 1]) || 0,
          valueAction: Number(data[1][SATISFACTION_COLS.AVG_VALUE_ACTION - 1]) || 0,
        };
      }

      // Fallback: compute from individual responses (Q6-Q16 = satisfaction questions)
      var sums = { sat: 0, steward: 0, access: 0 };
      var count = 0;
      for (var i = 1; i < data.length; i++) {
        var sat = Number(data[i][SATISFACTION_COLS.Q6_SATISFIED_REP - 1]);
        if (!sat) continue;
        count++;
        sums.sat += sat;
        sums.steward += Number(data[i][SATISFACTION_COLS.Q7_TRUST_UNION - 1]) || 0;
        sums.access += Number(data[i][SATISFACTION_COLS.Q18_KNOW_CONTACT - 1]) || 0;
      }
      if (count === 0) return {};
      return {
        overallSat: Math.round((sums.sat / count) * 10) / 10,
        stewardRating: Math.round((sums.steward / count) * 10) / 10,
        stewardAccess: Math.round((sums.access / count) * 10) / 10,
        responseCount: count,
      };
    } catch (_e) {
      Logger.log('getSatisfactionTrends error: ' + _e.message);
      return {};
    }
  }

  /**
   * Submits feedback from a member.
   * @param {string} email - Submitter email
   * @param {Object} data - { category, type, priority, title, description }
   * @returns {Object} { success, message }
   */
  function submitFeedback(email, feedbackData) {
    try {
      if (!email || !feedbackData || !feedbackData.title) {
        return { success: false, message: 'Missing required fields.' };
      }
      email = String(email).trim().toLowerCase();

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
      var sheet = ss.getSheetByName(SHEETS.FEEDBACK);
      if (!sheet) return { success: false, message: 'Feedback sheet not found.' };

      var row = [
        new Date(),
        email,
        String(feedbackData.category || 'General').substring(0, 100),
        String(feedbackData.type || 'Suggestion').substring(0, 50),
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
    getOrgLinks: getOrgLinks,
    getStewardMembers: getStewardMembers,
    getAllMembers: getAllMembers,
    getStewardSurveyTracking: getStewardSurveyTracking,
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
    getStewardAssignedMemberTasks: getStewardAssignedMemberTasks,
    // v4.17.0 - Feedback, Polls, Minutes
    getMyFeedback: getMyFeedback,
    getActivePolls: getActivePolls,
    submitPollVote: submitPollVote,
    getMeetingMinutes: getMeetingMinutes,
    addMeetingMinutes: addMeetingMinutes,
    addPoll: addPoll,
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
    _invalidateSheetCache: _invalidateSheetCache,
  };

})();


// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

/**
 * CR-01: Resolves the caller's email server-side. Never trust client-supplied email.
 * @returns {string} The authenticated caller's email, or empty string if unavailable.
 * @private
 */
function _resolveCallerEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email ? email.toLowerCase().trim() : '';
  } catch (_e) {
    return '';
  }
}

/**
 * Resolves the caller's email and verifies steward role.
 * Use for all steward-only operations.
 * @returns {string|null} Steward's email if authorized, null otherwise.
 * @private
 */
function _requireStewardAuth() {
  var auth = checkWebAppAuthorization('steward');
  if (!auth.isAuthorized) return null;
  return (auth.email || '').toLowerCase().trim();
}

// ═══════════════════════════════════════
// AUTHENTICATED DATA SERVICE WRAPPERS
// ═══════════════════════════════════════
// Security model (CR-AUTH-3):
//   - Steward ops: _requireStewardAuth() verifies steward role, uses server email
//   - Member self-service: _resolveCallerEmail() provides server-verified identity
//   - Public reads: no auth required (aggregate/non-PII data only)

function dataGetStewardCases() { var s = _requireStewardAuth(); return s ? DataService.getStewardCases(s) : []; }
function dataGetStewardKPIs() { var s = _requireStewardAuth(); return s ? DataService.getStewardKPIs(s) : {}; }
function dataGetMemberGrievances() { var e = _resolveCallerEmail(); return e ? DataService.getMemberGrievances(e) : []; }
function dataGetMemberGrievanceHistory() { var e = _resolveCallerEmail(); return e ? DataService.getMemberGrievanceHistory(e) : { success: false, message: 'Not authenticated.' }; }
function dataGetStewardContact() { var e = _resolveCallerEmail(); return e ? DataService.getStewardContact(e) : null; }

// v4.11.0 — data service wrappers (CR-AUTH-3: server-side identity + role checks)
// Steward: view any member's full profile; Member: view own profile only
function dataGetFullProfile(email) {
  var caller = _resolveCallerEmail();
  if (!caller) return { success: false, message: 'Not authenticated.' };
  var isSteward = checkWebAppAuthorization('steward').isAuthorized;
  // Members may only fetch their own profile
  var targetEmail = (isSteward && email) ? email : caller;
  return DataService.getFullMemberProfile(targetEmail);
}
// Member self-service: update own safe fields (address, workLocation, officeDays only)
// Stewards can also update member profiles; both paths use updateMemberProfile's field allowlist
function dataUpdateProfile(ignoredEmail, updates) {
  var e = _resolveCallerEmail();
  if (!e) return { success: false, message: 'Not authenticated.' };
  var isSteward = checkWebAppAuthorization('steward').isAuthorized;
  // Members can only update their own record; stewards can pass a target email via updates._targetEmail
  var targetEmail = (isSteward && updates && updates._targetEmail) ? updates._targetEmail : e;
  if (updates && updates._targetEmail) delete updates._targetEmail; // strip internal routing field
  return DataService.updateMemberProfile(targetEmail, updates);
}
function dataGetAssignedSteward() { var e = _resolveCallerEmail(); return e ? DataService.getAssignedStewardInfo(e) : null; }
function dataGetAvailableStewards() { var e = _resolveCallerEmail(); return e ? DataService.getAvailableStewards(e) : []; }
function dataAssignSteward(memberEmail, stewardEmail) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.assignStewardToMember(memberEmail, stewardEmail); }
function dataStartGrievanceDraft(ignoredEmail, data) { var e = _resolveCallerEmail(); return e ? DataService.startGrievanceDraft(e, data) : { success: false, message: 'Not authenticated.' }; }
function dataCreateGrievanceDrive() { var e = _resolveCallerEmail(); return e ? DataService.createGrievanceDriveFolder(e) : { success: false, message: 'Not authenticated.' }; }
function dataGetSurveyStatus() { var e = _resolveCallerEmail(); return e ? DataService.getMemberSurveyStatus(e) : null; }
function dataGetAllMembers() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getAllMembers(); }
function dataGetStewardSurveyTracking(ignoredEmail, scope) { var s = _requireStewardAuth(); if (!s) return { total: 0, completed: 0, members: [] }; try { return DataService.getStewardSurveyTracking(s, scope); } catch (e) { Logger.log('dataGetStewardSurveyTracking error: ' + e.message + '\n' + (e.stack || '')); return { total: 0, completed: 0, members: [] }; } }
function dataSendBroadcast(ignoredEmail, filter, msg) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.sendBroadcastMessage(s, filter, msg); }
function dataGetSurveyResults() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getSurveyResults(); }
// v4.21.0 — Native survey engine wrappers
function dataGetSurveyQuestions() { var e = _resolveCallerEmail(); return e ? getSurveyQuestions() : []; }
function dataSubmitSurveyResponse(ignoredEmail, responses) { var e = _resolveCallerEmail(); return e ? submitSurveyResponse(e, responses) : { success: false, message: 'Not authenticated.' }; }
// dataGetPendingSurveyMembers, dataGetSatisfactionSummary, dataOpenNewSurveyPeriod are in 08e_SurveyEngine.gs
function dataLogMemberContact(ignoredStewardEmail, memberEmail, type, notes, duration) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.logMemberContact(s, memberEmail, type, notes, duration); }
function dataGetMemberContactHistory(ignoredStewardEmail, memberEmail) { var s = _requireStewardAuth(); if (!s) return []; return DataService.getMemberContactHistory(s, memberEmail); }
function dataGetStewardContactLog() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getStewardContactLog(s); }

// Send a direct email notification to a single member (steward-only)
function dataSendDirectMessage(ignoredEmail, memberEmail, subject, body) {
  var s = _requireStewardAuth();
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!memberEmail || !subject || !body) return { success: false, message: 'Missing required fields.' };
  try {
    var config = (typeof ConfigReader !== 'undefined') ? ConfigReader.getConfig() : {};
    var fullSubject = (config.orgAbbrev || 'Union') + ' — ' + subject;
    MailApp.sendEmail(memberEmail.trim(), fullSubject, body);
    if (typeof logAuditEvent === 'function') logAuditEvent('DIRECT_MESSAGE_SENT', { steward: s, member: memberEmail, subject: subject });
    return { success: true, message: 'Message sent.' };
  } catch (e) {
    Logger.log('dataSendDirectMessage error: ' + e.message);
    return { success: false, message: 'Failed to send: ' + e.message };
  }
}

// Returns the Drive folder URL for a member's active (non-resolved) grievance (steward-only)
function dataGetMemberCaseFolderUrl(ignoredEmail, memberEmail) {
  var s = _requireStewardAuth();
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
function dataCreateTask(ignoredStewardEmail, title, desc, memberEmail, priority, dueDate, assignToEmail) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assignToEmail || ''); }
function dataCreateTaskForSteward(ignoredAssignerEmail, assigneeEmail, title, desc, memberEmail, priority, dueDate) {
  var s = _requireStewardAuth();
  if (!s) return { success: false, message: 'Steward access required.' };
  if (!DataService.isChiefSteward(s)) return { success: false, message: 'Not authorized.' };
  return DataService.createTask(s, title, desc, memberEmail, priority, dueDate, assigneeEmail);
}
function dataGetTasks(ignoredStewardEmail, statusFilter) { var s = _requireStewardAuth(); if (!s) return []; return DataService.getTasks(s, statusFilter); }
function dataCompleteTask(ignoredStewardEmail, taskId) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.completeTask(s, taskId); }
function dataGetStewardMemberStats() { var e = _resolveCallerEmail(); if (!e) return {}; try { return DataService.getStewardMemberStats(e); } catch (err) { Logger.log('dataGetStewardMemberStats error: ' + err.message + '\n' + (err.stack || '')); return { total: 0, byLocation: {}, byDues: {} }; } }
function dataGetStewardDirectory() { var e = _resolveCallerEmail(); if (!e) return []; try { return DataService.getStewardDirectory(); } catch (err) { Logger.log('dataGetStewardDirectory error: ' + err.message + '\n' + (err.stack || '')); return []; } }
function dataGetGrievanceStats() { var s = _requireStewardAuth(); if (!s) return { available: false }; return DataService.getGrievanceStats(); }
function dataGetGrievanceHotSpots() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getGrievanceHotSpots(); }
function dataGetMembershipStats() { var e = _resolveCallerEmail(); return e ? DataService.getMembershipStats() : null; }
function dataGetUpcomingEvents(limit) { var e = _resolveCallerEmail(); return e ? DataService.getUpcomingEvents(limit) : []; }
// dataGetSurveyQuestions and dataSubmitSurveyResponse are defined in the v4.21.0 block above (single canonical definition)
function dataIsChiefSteward() { var e = _resolveCallerEmail(); return e ? DataService.isChiefSteward(e) : false; }
// dataGetAgencyGrievanceStats — alias removed; frontend uses dataGetGrievanceStats directly

// v4.17.0 — member task assignment wrappers (CR-AUTH-3: server-side identity)
function dataCreateMemberTask(ignoredStewardEmail, memberEmail, title, desc, priority, dueDate) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.createMemberTask(s, memberEmail, title, desc, priority, dueDate); }
function dataGetMemberTasks(ignoredEmail, statusFilter) { var e = _resolveCallerEmail(); return e ? DataService.getMemberTasks(e, statusFilter) : []; }
function dataCompleteMemberTask(ignoredEmail, taskId) { var e = _resolveCallerEmail(); return e ? DataService.completeMemberTask(e, taskId) : { success: false, message: 'Not authenticated.' }; }
function dataGetStewardAssignedMemberTasks() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getStewardAssignedMemberTasks(s); }
// BUG-TASKS-03: steward completing a member task on the member's behalf
function dataStaffCompleteMemberTask(taskId) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.stewardCompleteMemberTask(s, taskId); }

// v4.16.0 — unwired sheet wrappers (CR-AUTH-3: server-side identity + role checks)
function dataUpdateTask(ignoredEmail, taskId, updates) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.updateTask(s, taskId, updates); }
function dataGetAllStewardPerformance() { var s = _requireStewardAuth(); if (!s) return []; return DataService.getAllStewardPerformance(); }
function dataGetCaseChecklist(caseId) { var e = _resolveCallerEmail(); return e ? DataService.getCaseChecklist(caseId) : []; }
function dataToggleChecklistItem(checklistId, completed) { var e = _resolveCallerEmail(); return e ? DataService.toggleChecklistItem(checklistId, completed, e) : { success: false, message: 'Not authenticated.' }; }
function dataGetMemberMeetings() { var e = _resolveCallerEmail(); return e ? DataService.getMemberMeetings(e) : []; }
function dataGetSatisfactionTrends() { var s = _requireStewardAuth(); if (!s) return { categories: [] }; return DataService.getSatisfactionTrends(); }
function dataSubmitFeedback(ignoredEmail, data) { var e = _resolveCallerEmail(); return e ? DataService.submitFeedback(e, data) : { success: false, message: 'Not authenticated.' }; }
function dataGetMyFeedback() { var e = _resolveCallerEmail(); return e ? DataService.getMyFeedback(e) : []; }
// v4.23.0: Portal Polls system deprecated — replaced by unified wq* poll system (24_WeeklyQuestions.gs)
// These stubs kept so any stale clients that call them get a graceful empty response
function dataGetActivePolls() { return []; }
function dataSubmitPollVote() { return { success: false, message: 'Polls system updated — please refresh.' }; }
function dataAddPoll() { return { success: false, message: 'Polls system updated — please refresh.' }; }
function dataGetMeetingMinutes(limit) { var e = _resolveCallerEmail(); return e ? DataService.getMeetingMinutes(limit) : []; }
function dataAddMeetingMinutes(ignoredStewardEmail, data) { var s = _requireStewardAuth(); if (!s) return { success: false, message: 'Steward access required.' }; return DataService.addMeetingMinutes(s, data); }

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
  try { ui = SpreadsheetApp.getUi(); } catch (_e) {}

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
  } catch (_re) {}

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
        } catch (_mv) {}
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


// Batch data fetch — single round-trip for SPA init (CR-AUTH-3: server-side identity + role)
// Role is re-verified server-side from the Member Directory; client-supplied role is ignored.
function dataGetBatchData(ignoredEmail, ignoredRole) {
  var e = _resolveCallerEmail();
  if (!e) return {};
  // Re-derive role from directory — never trust the client-supplied value
  var serverRole = DataService.getUserRole(e) || 'member';
  // Normalize 'both' → steward view (steward functions are a superset)
  if (serverRole === 'both' || serverRole === 'admin') serverRole = 'steward';
  return DataService.getBatchData(e, serverRole);
}

// Broadcast filter options (CR-AUTH-3: steward auth required)
function dataGetBroadcastFilterOptions() {
  var s = _requireStewardAuth();
  if (!s) return { locations: [], officeDays: [], duesStatuses: [], totalMembers: 0 };
  try {
    var members = DataService.getAllMembers();
    var locations = {};
    var officeDays = {};
    var duesStatuses = {};
    members.forEach(function(m) {
      if (m.workLocation) locations[m.workLocation] = true;
      var days = String(m.officeDays || '');
      if (days) {
        days.split(/[,;]/).forEach(function(d) {
          var day = d.trim();
          if (day) officeDays[day] = true;
        });
      }
      if (m.duesStatus) duesStatuses[m.duesStatus] = true;
    });
    return {
      locations: Object.keys(locations).sort(),
      officeDays: Object.keys(officeDays).sort(),
      duesStatuses: Object.keys(duesStatuses).sort(),
      totalMembers: members.length
    };
  } catch (e) {
    Logger.log('dataGetBroadcastFilterOptions error: ' + e.message + '\n' + (e.stack || ''));
    return { locations: [], officeDays: [], duesStatuses: [], totalMembers: 0 };
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
 *   resourceDownloads    — not currently tracked; returns 0
 *   membershipTrends     — monthly total/new member counts from Member Directory HIRE_DATE (last 6 mo)
 */
function dataGetEngagementStats() {
  var _caller = _resolveCallerEmail();
  if (!_caller) return null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    } catch(_se) {}

    // ── Weekly question votes ───────────────────────────────────────────────
    var weeklyQuestionVotes = 0;
    try {
      var wqRows = _rows(SHEETS.WEEKLY_RESPONSES || '_Weekly_Responses');
      weeklyQuestionVotes = wqRows.length;
    } catch(_we) {}

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
    } catch(_ce) {}

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
    } catch(_ge) {}

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
    } catch(_cl) {}

    // ── Membership trends (last 6 months, by hire date) ─────────────────────
    var membershipTrends = [];
    try {
      if (memberSheet && hireIdx >= 0) {
        var now = new Date();
        var mData2 = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).getValues();
        var monthMap = {};
        for (var ti = 5; ti >= 0; ti--) {
          var d = new Date(now.getFullYear(), now.getMonth() - ti, 1);
          var key = (d.getMonth() + 1) + '/' + (d.getFullYear() - 2000);
          monthMap[key] = { month: key, total: 0, new: 0 };
        }
        var firstKey = Object.keys(monthMap)[0];
        var firstDate = new Date(now.getFullYear(), now.getMonth() - 5, 1);
        for (var ri = 0; ri < mData2.length; ri++) {
          var hire = mData2[ri][hireIdx];
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
    } catch(_te) {}

    return {
      surveyParticipation:  surveyParticipation,
      weeklyQuestionVotes:  weeklyQuestionVotes,
      eventAttendance:      eventAttendance,
      grievanceFilingRate:  grievanceFilingRate,
      stewardContactRate:   stewardContactRate,
      resourceDownloads:    0,  // Not currently tracked — reserved for future resource click-tracking
      membershipTrends:     membershipTrends,
    };
  } catch (e) {
    Logger.log('dataGetEngagementStats error: ' + e.message + '\n' + (e.stack || ''));
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
function dataGetWorkloadSummaryStats() {
  var _caller = _resolveCallerEmail();
  if (!_caller) return null;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    } catch(_sr) {}

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
    } catch(_td) {}

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
function dataGetWelcomeData() {
  // CR-AUTH-3: Use server-side identity instead of client-supplied email
  var email = _resolveCallerEmail();
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
function dataMarkWelcomeDismissed() {
  // CR-AUTH-3: Use server-side identity instead of client-supplied email
  var email = _resolveCallerEmail();
  if (!email) return { success: false };
  var emailHash = _welcomeEmailHash(email);
  var propKey = 'WELCOME_DISMISSED_' + emailHash;
  var props = PropertiesService.getScriptProperties();
  props.setProperty(propKey, new Date().toISOString());
  return { success: true };
}

/**
 * Simple hash of email for property key (avoids special chars in key names).
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
