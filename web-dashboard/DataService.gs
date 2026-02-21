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
    memberEmail:     ['email', 'email address', 'member email'],
    memberName:      ['name', 'full name', 'member name'],
    memberFirstName: ['first name', 'first'],
    memberLastName:  ['last name', 'last'],
    memberRole:      ['role', 'member role', 'type'],
    memberUnit:      ['unit', 'workplace unit', 'department'],
    memberPhone:     ['phone', 'phone number', 'cell', 'mobile'],
    memberJoined:    ['joined', 'join date', 'member since', 'date joined'],
    memberDuesStatus:['dues status', 'dues', 'status'],
    memberId:        ['member id', 'id', 'member number'],
    
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
    };
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
    
    return {
      id: String(_getVal(row, colMap, HEADERS.grievanceId, '')).trim(),
      memberEmail: String(_getVal(row, colMap, HEADERS.grievanceMemberEmail, '')).trim().toLowerCase(),
      status: status,
      step: String(_getVal(row, colMap, HEADERS.grievanceStep, '')).trim(),
      deadline: deadlineFormatted,
      deadlineDays: deadlineDays,
      filed: filedFormatted,
      filedTimestamp: filedTimestamp,
      steward: String(_getVal(row, colMap, HEADERS.grievanceSteward, '')).trim().toLowerCase(),
      unit: String(_getVal(row, colMap, HEADERS.grievanceUnit, '')).trim(),
      priority: String(_getVal(row, colMap, HEADERS.grievancePriority, 'medium')).trim().toLowerCase(),
      notes: String(_getVal(row, colMap, HEADERS.grievanceNotes, '')).trim(),
    };
  }
  
  function _formatDate(date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  }
  
  // Public API
  return {
    findUserByEmail: findUserByEmail,
    getUserRole: getUserRole,
    getStewardCases: getStewardCases,
    getStewardKPIs: getStewardKPIs,
    getMemberGrievances: getMemberGrievances,
    getStewardContact: getStewardContact,
    getUnits: getUnits,
  };
  
})();


// ═══════════════════════════════════════
// GLOBAL FUNCTIONS (callable from client via google.script.run)
// ═══════════════════════════════════════

function dataGetStewardCases(email) { return DataService.getStewardCases(email); }
function dataGetStewardKPIs(email) { return DataService.getStewardKPIs(email); }
function dataGetMemberGrievances(email) { return DataService.getMemberGrievances(email); }
function dataGetStewardContact(email) { return DataService.getStewardContact(email); }
function dataGetUserRole(email) { return DataService.getUserRole(email); }
function dataGetUserProfile(email) { return DataService.findUserByEmail(email); }
function dataGetUnits() { return DataService.getUnits(); }
