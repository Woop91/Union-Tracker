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

  // Sheet names — read from Config if available, fallback to defaults
  var MEMBER_SHEET = 'Member Directory';
  var GRIEVANCE_SHEET = 'Grievance Log';

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
    grievanceStatus: ['status', 'grievance status', 'case status'],
    grievanceStep:   ['step', 'current step', 'grievance step'],
    grievanceDeadline: ['deadline', 'next deadline', 'due date'],
    grievanceFiled:  ['filed', 'filed date', 'date filed', 'created'],
    grievanceSteward:['steward', 'assigned steward', 'steward email', 'assigned to'],
    grievanceUnit:   ['unit', 'workplace unit'],
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

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);

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

    var sheet = _getSheet(GRIEVANCE_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);

    var stewardCol = _findColumn(colMap, HEADERS.grievanceSteward);
    if (stewardCol === -1) {
      Logger.log('DataService: Steward column not found in Grievance Log');
      return [];
    }

    var cases = [];
    for (var i = 1; i < data.length; i++) {
      var assignedTo = String(data[i][stewardCol]).trim().toLowerCase();
      if (assignedTo === stewardEmail) {
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

    var now = new Date();
    var sevenDaysMs = 7 * 24 * 60 * 60 * 1000;

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

    var sheet = _getSheet(GRIEVANCE_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);

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

    var sheet = _getSheet(GRIEVANCE_SHEET);
    if (!sheet) return { success: true, history: [] };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colMap = _buildColumnMap(headers);

    var memberCol = _findColumn(colMap, HEADERS.grievanceMemberEmail);
    if (memberCol === -1) return { success: true, history: [] };

    var closedStatuses = ['settled', 'won', 'denied', 'withdrawn', 'closed'];
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
      for (var field in updates) {
        if (!editableFields[field]) continue;
        var col = _findColumn(colMap, editableFields[field]);
        if (col === -1) continue;
        var val = String(updates[field] || '').trim().substring(0, 255);
        sheet.getRange(rowNum, col + 1).setValue(val);
      }

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

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
    var isStewardCol = _findColumn(colMap, HEADERS.memberIsSteward);
    var roleCol = _findColumn(colMap, HEADERS.memberRole);
    var locationCol = _findColumn(colMap, HEADERS.memberWorkLocation);

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
   * Returns survey completion status for a member.
   * @param {string} email
   * @returns {Object} { hasCompleted: boolean, lastCompleted: string|null }
   */
  function getMemberSurveyStatus(email) {
    if (!email) return { hasCompleted: false, lastCompleted: null };
    email = String(email).trim().toLowerCase();

    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var trackSheet = ss.getSheetByName('_Survey_Tracking');
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
   * Returns org-wide links (calendar, drive, survey form, etc.)
   * @returns {Object}
   */
  function getOrgLinks() {
    try {
      var config = ConfigReader.getConfig();
      return {
        calendarUrl: config.calendarId ? 'https://calendar.google.com/calendar/embed?src=' + encodeURIComponent(config.calendarId) : '',
        driveFolderUrl: config.driveFolderId ? 'https://drive.google.com/drive/folders/' + config.driveFolderId : '',
        surveyFormUrl: config.satisfactionFormUrl || '',
        orgWebsite: config.orgWebsite || '',
      };
    } catch (e) {
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

    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
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
   * Returns survey completion tracking for a steward's assigned members.
   * Does NOT return actual survey responses — only completion status.
   * @param {string} stewardEmail
   * @returns {Object} { total, completed, members: [{name, email, completed}] }
   */
  function getStewardSurveyTracking(stewardEmail) {
    var members = getStewardMembers(stewardEmail);
    if (members.length === 0) return { total: 0, completed: 0, members: [] };

    var tracking = [];
    var completedCount = 0;

    for (var i = 0; i < members.length; i++) {
      var status = getMemberSurveyStatus(members[i].email);
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
    if (!stewardEmail || !message) return { success: false, sentCount: 0, message: 'Missing required fields.' };

    var members = getStewardMembers(stewardEmail);
    if (members.length === 0) return { success: false, sentCount: 0, message: 'No assigned members found.' };

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
    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
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
        var now = new Date();
        var diff = parsed.getTime() - now.getTime();
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
      var parsed = new Date(filedRaw);
      if (!isNaN(parsed.getTime())) {
        filedFormatted = _formatDate(parsed);
        filedTimestamp = parsed.getTime();
      }
    }

    var status = String(_getVal(row, colMap, HEADERS.grievanceStatus, 'new')).trim().toLowerCase();

    // Auto-detect overdue if deadline is past and status is active
    if (deadlineDays !== null && deadlineDays < 0 && status !== 'resolved') {
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
      status: status,
      step: String(_getVal(row, colMap, HEADERS.grievanceStep, '')).trim(),
      deadline: deadlineFormatted,
      deadlineDays: deadlineDays,
      filed: filedFormatted,
      filedTimestamp: filedTimestamp,
      closedTimestamp: closedTimestamp,
      steward: String(_getVal(row, colMap, HEADERS.grievanceSteward, '')).trim().toLowerCase(),
      unit: String(_getVal(row, colMap, HEADERS.grievanceUnit, '')).trim(),
      priority: String(_getVal(row, colMap, HEADERS.grievancePriority, 'medium')).trim().toLowerCase(),
      notes: String(_getVal(row, colMap, HEADERS.grievanceNotes, '')).trim(),
      issueCategory: String(_getVal(row, colMap, HEADERS.grievanceIssueCategory, '')).trim(),
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
        results.push({ id: data[i][0], type: data[i][3], date: data[i][4] instanceof Date ? _formatDate(data[i][4]) : String(data[i][4]), notes: data[i][5], duration: data[i][6] });
      }
    }
    results.sort(function(a, b) { return b.date > a.date ? 1 : -1; });
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
        results.push({ id: data[i][0], memberEmail: data[i][2], type: data[i][3], date: data[i][4] instanceof Date ? _formatDate(data[i][4]) : String(data[i][4]), notes: data[i][5], duration: data[i][6] });
      }
    }
    results.sort(function(a, b) { return b.date > a.date ? 1 : -1; });
    return results.slice(0, 100);
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Tasks (v4.12.0)
  // ═══════════════════════════════════════

  function _ensureStewardTasks() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    return { success: true, message: 'Task created.', taskId: id };
  }

  function getTasks(stewardEmail, statusFilter) {
    var sheet = _ensureStewardTasks();
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var tasks = [];
    var sEmail = stewardEmail.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() !== sEmail) continue;
      var status = String(data[i][6]).toLowerCase().trim();
      if (statusFilter && status !== statusFilter) continue;
      var dueDateRaw = data[i][7];
      var dueStr = dueDateRaw instanceof Date ? _formatDate(dueDateRaw) : String(dueDateRaw || '');
      var dueDays = null;
      if (dueDateRaw instanceof Date) {
        dueDays = Math.ceil((dueDateRaw.getTime() - Date.now()) / (86400000));
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
        if (updates.status) sheet.getRange(i + 1, 7).setValue(updates.status);
        if (updates.priority) sheet.getRange(i + 1, 6).setValue(updates.priority);
        if (updates.title) sheet.getRange(i + 1, 3).setValue(updates.title.substring(0, 200));
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
        return { success: true };
      }
    }
    return { success: false, message: 'Task not found.' };
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
        var cfgSheet = ss.getSheetByName('Config');
        if (cfgSheet) email = String(cfgSheet.getRange(2, CONFIG_COLS.CHIEF_STEWARD_EMAIL).getValue() || '').toLowerCase().trim();
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
    if (members.length === 0) return { total: 0, byLocation: {}, byDues: {} };
    var byLocation = {};
    var byDues = {};
    members.forEach(function(m) {
      var loc = m.workLocation || 'Unknown';
      byLocation[loc] = (byLocation[loc] || 0) + 1;
      var dues = m.duesStatus || 'Unknown';
      byDues[dues] = (byDues[dues] || 0) + 1;
    });
    return { total: members.length, byLocation: byLocation, byDues: byDues };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward Directory (v4.12.0)
  // ═══════════════════════════════════════

  function getStewardDirectory() {
    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var colMap = _buildColumnMap(data[0]);
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
    stewards.sort(function(a, b) { return (a.name || '').localeCompare(b.name || ''); });
    return stewards;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Grievance Stats (v4.12.0) — anonymized
  // ═══════════════════════════════════════

  function getGrievanceStats() {
    var sheet = _getSheet(GRIEVANCE_SHEET);
    if (!sheet) return { available: false };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { available: false };
    var colMap = _buildColumnMap(data[0]);
    var total = data.length - 1;
    if (total < 10) return { available: false, count: total, threshold: 10 };

    var byStatus = {};
    var byStep = {};
    var byUnit = {};
    var byCategory = {};
    var monthly = {};
    var monthlyResolved = {};
    var openCount = 0, wonCount = 0, deniedCount = 0, settledCount = 0;
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
    };
  }

  function getGrievanceHotSpots() {
    var sheet = _getSheet(GRIEVANCE_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var colMap = _buildColumnMap(data[0]);
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
    var sheet = _getSheet(MEMBER_SHEET);
    if (!sheet) return { available: false };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { available: false };
    var total = data.length - 1;
    if (total < 20) return { available: false, count: total, threshold: 20 };

    var colMap = _buildColumnMap(data[0]);
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
      if (!config.calendarId) return [];
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
  // PUBLIC: Steward Performance (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns performance data for a single steward by email.
   * Reads _Steward_Performance_Calc and cross-refs Member Directory for name→email.
   * @param {string} email
   * @returns {Object|null}
   */
  function getStewardPerformance(email) {
    if (!email) return null;
    email = String(email).trim().toLowerCase();

    var user = findUserByEmail(email);
    if (!user) return null;
    var stewardName = user.name;
    if (!stewardName) return null;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
    if (!sheet || sheet.getLastRow() <= 1) return null;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][STEWARD_PERF_COLS.STEWARD - 1]).trim();
      if (rowName.toLowerCase() === stewardName.toLowerCase()) {
        return {
          steward: rowName,
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
    }
    return null;
  }

  /**
   * Returns all steward performance rows (for chief steward / insights).
   * @returns {Object[]}
   */
  function getAllStewardPerformance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][STEWARD_PERF_COLS.STEWARD - 1]).trim();
      if (!name) continue;
      results.push({
        steward: name,
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
    results.sort(function(a, b) { return b.performanceScore - a.performanceScore; });
    return results;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Case Checklist (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns checklist items for a case. Wraps 12_Features.gs function.
   * @param {string} caseId
   * @returns {Object[]}
   */
  function getCaseChecklist(caseId) {
    if (!caseId) return [];
    if (typeof getChecklistItems === 'function') return getChecklistItems(caseId);
    return [];
  }

  /**
   * Returns checklist progress for a case. Wraps 12_Features.gs function.
   * @param {string} caseId
   * @returns {Object}
   */
  function getCaseChecklistProgress(caseId) {
    if (!caseId) return { completed: 0, total: 0, percentage: 0, display: 'No checklist' };
    if (typeof getChecklistProgress === 'function') return getChecklistProgress(caseId);
    return { completed: 0, total: 0, percentage: 0, display: 'No checklist' };
  }

  /**
   * Toggles a checklist item. Wraps 12_Features.gs function.
   * @param {string} checklistId
   * @param {boolean} completed
   * @param {string} userEmail
   * @returns {Object}
   */
  function toggleChecklistItem(checklistId, completed, userEmail) {
    if (typeof setChecklistItemCompleted === 'function') {
      return setChecklistItemCompleted(checklistId, completed, userEmail);
    }
    return { success: false, message: 'Checklist feature unavailable.' };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Member Meetings (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns meeting check-in records for a member.
   * @param {string} email
   * @returns {Object[]}
   */
  function getMemberMeetings(email) {
    if (!email) return [];
    email = String(email).trim().toLowerCase();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getDataRange().getValues();
    var meetings = [];
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][MEETING_CHECKIN_COLS.EMAIL - 1]).trim().toLowerCase();
      if (rowEmail !== email) continue;

      var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
      var dateStr = meetingDate instanceof Date ? _formatDate(meetingDate) : String(meetingDate || '');
      var dateTs = meetingDate instanceof Date ? meetingDate.getTime() : 0;

      meetings.push({
        meetingId: String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || ''),
        meetingName: String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || ''),
        meetingDate: dateStr,
        meetingDateTs: dateTs,
        meetingType: String(data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || ''),
        checkInTime: data[i][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1] instanceof Date
          ? _formatDate(data[i][MEETING_CHECKIN_COLS.CHECKIN_TIME - 1]) : '',
        duration: String(data[i][MEETING_CHECKIN_COLS.MEETING_DURATION - 1] || ''),
        notesUrl: String(data[i][MEETING_CHECKIN_COLS.NOTES_DOC_URL - 1] || ''),
        agendaUrl: String(data[i][MEETING_CHECKIN_COLS.AGENDA_DOC_URL - 1] || ''),
      });
    }
    meetings.sort(function(a, b) { return (b.meetingDateTs || 0) - (a.meetingDateTs || 0); });
    meetings.forEach(function(m) { delete m.meetingDateTs; });
    return meetings;
  }

  /**
   * Returns aggregate meeting stats.
   * @returns {Object}
   */
  function getMeetingStats() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() <= 1) return { total: 0, byType: {}, recent: [] };

    var data = sheet.getDataRange().getValues();
    var byType = {};
    var meetingIds = {};
    for (var i = 1; i < data.length; i++) {
      var type = String(data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || 'Other').trim();
      byType[type] = (byType[type] || 0) + 1;
      var mid = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      if (mid) meetingIds[mid] = true;
    }

    return { total: Object.keys(meetingIds).length, attendanceCount: data.length - 1, byType: byType };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Satisfaction Trends (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Returns satisfaction trend data from the AVG columns (BT-CD).
   * @returns {Object}
   */
  function getSatisfactionTrends() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.SATISFACTION);
    if (!sheet || sheet.getLastRow() <= 1) return { categories: [], overall: 0, responseCount: 0 };

    var data = sheet.getDataRange().getValues();
    var responseCount = data.length - 1;
    if (responseCount < 10) return { categories: [], overall: 0, responseCount: responseCount };

    var avgCols = [
      { key: 'AVG_OVERALL_SAT',    name: 'Overall Satisfaction', col: SATISFACTION_COLS.AVG_OVERALL_SAT },
      { key: 'AVG_STEWARD_RATING', name: 'Steward Ratings',     col: SATISFACTION_COLS.AVG_STEWARD_RATING },
      { key: 'AVG_STEWARD_ACCESS', name: 'Steward Access',      col: SATISFACTION_COLS.AVG_STEWARD_ACCESS },
      { key: 'AVG_CHAPTER',        name: 'Chapter Effectiveness', col: SATISFACTION_COLS.AVG_CHAPTER },
      { key: 'AVG_LEADERSHIP',     name: 'Local Leadership',    col: SATISFACTION_COLS.AVG_LEADERSHIP },
      { key: 'AVG_CONTRACT',       name: 'Contract Enforcement', col: SATISFACTION_COLS.AVG_CONTRACT },
      { key: 'AVG_REPRESENTATION', name: 'Representation',      col: SATISFACTION_COLS.AVG_REPRESENTATION },
      { key: 'AVG_COMMUNICATION',  name: 'Communication',       col: SATISFACTION_COLS.AVG_COMMUNICATION },
      { key: 'AVG_MEMBER_VOICE',   name: 'Member Voice',        col: SATISFACTION_COLS.AVG_MEMBER_VOICE },
      { key: 'AVG_VALUE_ACTION',   name: 'Value & Action',      col: SATISFACTION_COLS.AVG_VALUE_ACTION },
      { key: 'AVG_SCHEDULING',     name: 'Scheduling',          col: SATISFACTION_COLS.AVG_SCHEDULING },
    ];

    var categories = [];
    var overallSum = 0;
    var overallCount = 0;

    for (var c = 0; c < avgCols.length; c++) {
      var colIdx = avgCols[c].col - 1; // 0-indexed
      if (colIdx < 0) continue;
      var sum = 0;
      var count = 0;
      for (var r = 1; r < data.length; r++) {
        if (colIdx < data[r].length) {
          var val = Number(data[r][colIdx]);
          if (!isNaN(val) && val > 0) {
            sum += val;
            count++;
          }
        }
      }
      var avg = count > 0 ? Math.round((sum / count) * 10) / 10 : 0;
      categories.push({ key: avgCols[c].key, name: avgCols[c].name, avg: avg, count: count });
      overallSum += sum;
      overallCount += count;
    }

    return {
      categories: categories,
      overall: overallCount > 0 ? Math.round((overallSum / overallCount) * 10) / 10 : 0,
      responseCount: responseCount,
    };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Feedback (v4.16.0)
  // ═══════════════════════════════════════

  /**
   * Submits a feedback entry.
   * @param {string} email
   * @param {Object} data - { category, type, priority, title, description }
   * @returns {Object}
   */
  function submitFeedback(email, feedbackData) {
    if (!email || !feedbackData || !feedbackData.title) {
      return { success: false, message: 'Missing required fields.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.FEEDBACK);
    if (!sheet) return { success: false, message: 'Feedback sheet not found.' };

    sheet.appendRow([
      new Date(),
      String(email).trim().toLowerCase(),
      String(feedbackData.category || 'Other').substring(0, 50),
      String(feedbackData.type || '').substring(0, 50),
      String(feedbackData.priority || 'Medium').substring(0, 20),
      String(feedbackData.title).substring(0, 200),
      String(feedbackData.description || '').substring(0, 1000),
      'New',
      '',
      '',
      '',
    ]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('FEEDBACK_SUBMITTED', { email: email, category: feedbackData.category });
    }
    return { success: true, message: 'Feedback submitted.' };
  }

  /**
   * Returns user's submitted feedback.
   * @param {string} email
   * @returns {Object[]}
   */
  function getMyFeedback(email) {
    if (!email) return [];
    email = String(email).trim().toLowerCase();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
      });
    }
    minutes.sort(function(a, b) { return (b.meetingDateTs || 0) - (a.meetingDateTs || 0); });
    minutes.forEach(function(m) { delete m.meetingDateTs; });
    return minutes.slice(0, limit);
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

    sheet.appendRow([
      id,
      meetingDate,
      String(minutesData.title).substring(0, 200),
      String(minutesData.bullets || '').substring(0, 2000),
      String(minutesData.fullMinutes || '').substring(0, 5000),
      String(stewardEmail).trim().toLowerCase(),
      new Date(),
    ]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MINUTES_ADDED', { steward: stewardEmail, title: minutesData.title });
    }
    return { success: true, message: 'Minutes added.', id: id };
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
    getStewardSurveyTracking: getStewardSurveyTracking,
    sendBroadcastMessage: sendBroadcastMessage,
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
    // v4.16.0 — unwired sheets
    getStewardPerformance: getStewardPerformance,
    getAllStewardPerformance: getAllStewardPerformance,
    getCaseChecklist: getCaseChecklist,
    getCaseChecklistProgress: getCaseChecklistProgress,
    toggleChecklistItem: toggleChecklistItem,
    getMemberMeetings: getMemberMeetings,
    getMeetingStats: getMeetingStats,
    getSatisfactionTrends: getSatisfactionTrends,
    submitFeedback: submitFeedback,
    getMyFeedback: getMyFeedback,
    getActivePolls: getActivePolls,
    submitPollVote: submitPollVote,
    getMeetingMinutes: getMeetingMinutes,
    addMeetingMinutes: addMeetingMinutes,
    addPoll: addPoll,
  };

})();


// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

function dataGetStewardCases(email) { return DataService.getStewardCases(email); }
function dataGetStewardKPIs(email) { return DataService.getStewardKPIs(email); }
function dataGetMemberGrievances(email) { return DataService.getMemberGrievances(email); }
function dataGetMemberGrievanceHistory(email) { return DataService.getMemberGrievanceHistory(email); }
function dataGetStewardContact(email) { return DataService.getStewardContact(email); }
function dataGetUserRole(email) { return DataService.getUserRole(email); }
function dataGetUserProfile(email) { return DataService.findUserByEmail(email); }
function dataGetUnits() { return DataService.getUnits(); }

// v4.11.0 — new data service wrappers
function dataGetFullProfile(email) { return DataService.getFullMemberProfile(email); }
function dataUpdateProfile(email, updates) { return DataService.updateMemberProfile(email, updates); }
function dataGetAssignedSteward(email) { return DataService.getAssignedStewardInfo(email); }
function dataGetAvailableStewards(email) { return DataService.getAvailableStewards(email); }
function dataAssignSteward(memberEmail, stewardEmail) { return DataService.assignStewardToMember(memberEmail, stewardEmail); }
function dataGetGrievanceDriveUrl(email) { return DataService.getMemberGrievanceDriveUrl(email); }
function dataGetSurveyStatus(email) { return DataService.getMemberSurveyStatus(email); }
function dataGetOrgLinks() { return DataService.getOrgLinks(); }
function dataGetStewardMembers(email) { return DataService.getStewardMembers(email); }
function dataGetStewardSurveyTracking(email) { return DataService.getStewardSurveyTracking(email); }
function dataSendBroadcast(email, filter, msg) { return DataService.sendBroadcastMessage(email, filter, msg); }
function dataGetSurveyResults() { return DataService.getSurveyResults(); }

// v4.12.0 — new data service wrappers
function dataLogMemberContact(stewardEmail, memberEmail, type, notes, duration) { return DataService.logMemberContact(stewardEmail, memberEmail, type, notes, duration); }
function dataGetMemberContactHistory(stewardEmail, memberEmail) { return DataService.getMemberContactHistory(stewardEmail, memberEmail); }
function dataGetStewardContactLog(stewardEmail) { return DataService.getStewardContactLog(stewardEmail); }
function dataCreateTask(stewardEmail, title, desc, memberEmail, priority, dueDate) { return DataService.createTask(stewardEmail, title, desc, memberEmail, priority, dueDate); }
function dataCreateTaskForSteward(assignerEmail, assigneeEmail, title, desc, memberEmail, priority, dueDate) { return DataService.createTask(assignerEmail, title, desc, memberEmail, priority, dueDate, assigneeEmail); }
function dataGetTasks(stewardEmail, statusFilter) { return DataService.getTasks(stewardEmail, statusFilter); }
function dataUpdateTask(stewardEmail, taskId, updates) { return DataService.updateTask(stewardEmail, taskId, updates); }
function dataCompleteTask(stewardEmail, taskId) { return DataService.completeTask(stewardEmail, taskId); }
function dataGetStewardMemberStats(stewardEmail) { return DataService.getStewardMemberStats(stewardEmail); }
function dataGetStewardDirectory() { return DataService.getStewardDirectory(); }
function dataGetGrievanceStats() { return DataService.getGrievanceStats(); }
function dataGetGrievanceHotSpots() { return DataService.getGrievanceHotSpots(); }
function dataGetMembershipStats() { return DataService.getMembershipStats(); }
function dataGetUpcomingEvents(limit) { return DataService.getUpcomingEvents(limit); }
function dataGetSurveyQuestions() { return getSurveyQuestions(); }
function dataSubmitSurveyResponse(email, responses) { return submitSurveyResponse(email, responses); }
function dataIsChiefSteward(email) { return DataService.isChiefSteward(email); }
function dataGetChiefStewardTaskView(email) { return DataService.getChiefStewardTaskView(email); }
function dataGetAgencyGrievanceStats() { return DataService.getGrievanceStats(); }

// v4.16.0 — unwired sheet wrappers
function dataGetStewardPerformance(email) { return DataService.getStewardPerformance(email); }
function dataGetAllStewardPerformance() { return DataService.getAllStewardPerformance(); }
function dataGetCaseChecklist(caseId) { return DataService.getCaseChecklist(caseId); }
function dataGetCaseChecklistProgress(caseId) { return DataService.getCaseChecklistProgress(caseId); }
function dataToggleChecklistItem(checklistId, completed, userEmail) { return DataService.toggleChecklistItem(checklistId, completed, userEmail); }
function dataGetMemberMeetings(email) { return DataService.getMemberMeetings(email); }
function dataGetMeetingStats() { return DataService.getMeetingStats(); }
function dataGetSatisfactionTrends() { return DataService.getSatisfactionTrends(); }
function dataSubmitFeedback(email, data) { return DataService.submitFeedback(email, data); }
function dataGetMyFeedback(email) { return DataService.getMyFeedback(email); }
function dataGetActivePolls(email) { return DataService.getActivePolls(email); }
function dataSubmitPollVote(email, pollId, response) { return DataService.submitPollVote(email, pollId, response); }
function dataGetMeetingMinutes(limit) { return DataService.getMeetingMinutes(limit); }
function dataAddMeetingMinutes(stewardEmail, data) { return DataService.addMeetingMinutes(stewardEmail, data); }
function dataAddPoll(stewardEmail, question, options, unit) { return DataService.addPoll(stewardEmail, question, options, unit); }

// Batch data fetch — single round-trip for SPA init
function dataGetBatchData(email, role) { return getWebDashBatchData(email, role); }

// Broadcast filter options — returns unique values from steward's members
function dataGetBroadcastFilterOptions(stewardEmail) {
  var members = DataService.getStewardMembers(stewardEmail);
  var locations = {};
  var officeDays = {};
  var duesStatuses = {};
  members.forEach(function(m) {
    if (m.workLocation) locations[m.workLocation] = true;
    if (m.officeDays) {
      m.officeDays.split(/[,;]/).forEach(function(d) {
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
}

// Engagement stats — reads seeded union stats from Script Properties
function dataGetEngagementStats() {
  try {
    var json = PropertiesService.getScriptProperties().getProperty('SEEDED_UNION_STATS');
    if (!json) return null;
    var stats = JSON.parse(json);
    return {
      surveyParticipation: stats.engagement.surveyParticipation,
      weeklyQuestionVotes: stats.engagement.weeklyQuestionVotes,
      eventAttendance: stats.engagement.eventAttendance,
      grievanceFilingRate: stats.engagement.grievanceFilingRate,
      stewardContactRate: stats.engagement.stewardContactRate,
      resourceDownloads: stats.engagement.resourceDownloads,
      membershipTrends: stats.membershipTrends || [],
    };
  } catch (e) { Logger.log('dataGetEngagementStats error: ' + e.message); return null; }
}

// Workload summary stats — reads seeded workload aggregate from Script Properties
function dataGetWorkloadSummaryStats() {
  try {
    var json = PropertiesService.getScriptProperties().getProperty('SEEDED_UNION_STATS');
    if (!json) return null;
    var stats = JSON.parse(json);
    return stats.workloadSummary || null;
  } catch (e) { Logger.log('dataGetWorkloadSummaryStats error: ' + e.message); return null; }
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
function dataGetWelcomeData(email) {
  if (!email) return { isFirstVisit: false, userName: '', role: 'member', quickActions: [] };
  email = String(email).trim().toLowerCase();

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
function dataMarkWelcomeDismissed(email) {
  if (!email) return { success: false };
  email = String(email).trim().toLowerCase();
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
