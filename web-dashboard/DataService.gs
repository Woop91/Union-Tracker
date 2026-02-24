/**
 * DataService.gs
 * Data access layer matching the ACTUAL sheet schema.
 * 
 * MEMBER DIRECTORY COLUMNS:
 *   Member ID, First Name, Last Name, Job Title, Work Location, Unit,
 *   Cubicle, Office Days, Email, Phone, Open Rate %, Preferred Communication,
 *   Best Time to Contact, Supervisor, Manager, Is Steward (Yes/No),
 *   Committees, Assigned Steward, Last Virtual Mtg, Last In-Person Mtg,
 *   Volunteer Hours, Interest: Local/Chapter/Allied, Recent Contact Date,
 *   Contact Steward, Contact Notes, Has Open Grievance?, Grievance Status,
 *   Days to Deadline, Start Grievance, Actions, PIN Hash, Employee ID,
 *   Department, Hire Date
 * 
 * GRIEVANCE LOG COLUMNS:
 *   Grievance ID, Member ID, First Name, Last Name, Status, Current Step,
 *   Incident Date, Filing Deadline, Date Filed, Step I Due, Step I Rcvd,
 *   Step II Appeal Due, Step II Appeal Filed, Step II Due, Step II Rcvd,
 *   Step III Appeal Due, Step III Appeal Filed, Date Closed, Days Open,
 *   Next Action Due, Days to Deadline, Articles Violated, Issue Category,
 *   Member Email, Work Location, Assigned Steward, Resolution,
 *   Message Alert, Coordinator Message, Acknowledged By, Acknowledged Date,
 *   Drive Folder ID, Drive Folder URL, Actions, Action Type,
 *   Checklist Progress, Reminder 1/2 Date/Note, Last Updated
 * 
 * ALL COLUMN LOOKUPS BY HEADER NAME — NEVER BY INDEX.
 */

var DataService = (function () {
  
  var MEMBER_SHEET = 'Member Directory';
  var GRIEVANCE_SHEET = 'Grievance Log';
  
  // Header aliases — first match wins
  var MH = {
    memberId:      ['member id'],
    firstName:     ['first name'],
    lastName:      ['last name'],
    jobTitle:      ['job title'],
    workLocation:  ['work location'],
    unit:          ['unit'],
    cubicle:       ['cubicle'],
    officeDays:    ['office days'],
    email:         ['email'],
    phone:         ['phone'],
    openRate:      ['open rate %', 'open rate'],
    prefComm:      ['preferred communication'],
    bestTime:      ['best time to contact'],
    supervisor:    ['supervisor'],
    manager:       ['manager'],
    isSteward:     ['is steward'],
    committees:    ['committees'],
    assignedSteward: ['assigned steward'],
    lastVirtual:   ['last virtual mtg'],
    lastInPerson:  ['last in-person mtg'],
    volunteerHrs:  ['volunteer hours'],
    intLocal:      ['interest: local'],
    intChapter:    ['interest: chapter'],
    intAllied:     ['interest: allied'],
    recentContact: ['recent contact date'],
    contactSteward:['contact steward'],
    contactNotes:  ['contact notes'],
    hasOpenGrievance: ['has open grievance?', 'has open grievance'],
    grievanceStatus:  ['grievance status'],
    daysToDeadline:   ['days to deadline'],
    pinHash:       ['pin hash'],
    employeeId:    ['employee id'],
    department:    ['department'],
    hireDate:      ['hire date'],
  };
  
  var GH = {
    grievanceId:   ['grievance id'],
    memberId:      ['member id'],
    firstName:     ['first name'],
    lastName:      ['last name'],
    status:        ['status'],
    currentStep:   ['current step'],
    incidentDate:  ['incident date'],
    filingDeadline:['filing deadline'],
    dateFiled:     ['date filed'],
    stepIDue:      ['step i due'],
    stepIRcvd:     ['step i rcvd'],
    stepIIAppealDue:  ['step ii appeal due'],
    stepIIAppealFiled:['step ii appeal filed'],
    stepIIDue:     ['step ii due'],
    stepIIRcvd:    ['step ii rcvd'],
    stepIIIAppealDue: ['step iii appeal due'],
    stepIIIAppealFiled:['step iii appeal filed'],
    dateClosed:    ['date closed'],
    daysOpen:      ['days open'],
    nextActionDue: ['next action due'],
    daysToDeadline:['days to deadline'],
    articlesViolated:['articles violated'],
    issueCategory: ['issue category'],
    memberEmail:   ['member email'],
    workLocation:  ['work location'],
    assignedSteward:['assigned steward'],
    resolution:    ['resolution'],
    messageAlert:  ['message alert'],
    coordMessage:  ['coordinator message'],
    ackedBy:       ['acknowledged by'],
    ackedDate:     ['acknowledged date'],
    driveFolderId: ['drive folder id'],
    driveFolderUrl:['drive folder url'],
    actionType:    ['action type'],
    checklistProgress:['checklist progress'],
    reminder1Date: ['reminder 1 date'],
    reminder1Note: ['reminder 1 note'],
    reminder2Date: ['reminder 2 date'],
    reminder2Note: ['reminder 2 note', 'eminder 2 note'], // handles typo in sheet
    lastUpdated:   ['last updated'],
  };
  
  // ═══════════════════════════════════════
  // COLUMN HELPERS
  // ═══════════════════════════════════════
  
  function _colMap(headers) {
    var m = {};
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h) m[h] = i;
    }
    return m;
  }
  
  function _col(cm, aliases) {
    for (var i = 0; i < aliases.length; i++) {
      var k = aliases[i].toLowerCase();
      if (cm.hasOwnProperty(k)) return cm[k];
    }
    return -1;
  }
  
  function _val(row, cm, aliases, def) {
    var c = _col(cm, aliases);
    if (c === -1 || c >= row.length) return def !== undefined ? def : '';
    var v = row[c];
    return (v === null || v === undefined) ? (def !== undefined ? def : '') : v;
  }
  
  function _str(row, cm, aliases, def) { return String(_val(row, cm, aliases, def || '')).trim(); }
  function _date(row, cm, aliases) {
    var v = _val(row, cm, aliases, null);
    if (!v) return null;
    if (v instanceof Date) return v;
    var d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  function _fmtDate(d) {
    if (!d) return '';
    var m = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return m[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
  }
  function _daysUntil(d) {
    if (!d) return null;
    return Math.ceil((d.getTime() - Date.now()) / 86400000);
  }
  
  // ═══════════════════════════════════════
  // USER LOOKUP
  // ═══════════════════════════════════════
  
  function findUserByEmail(email) {
    if (!email) return null;
    email = String(email).trim().toLowerCase();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var ec = _col(cm, MH.email);
    if (ec === -1) return null;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ec]).trim().toLowerCase() === email) {
        return _buildUser(data[i], cm);
      }
    }
    return null;
  }
  
  function getUserRole(email) {
    var u = findUserByEmail(email);
    if (!u) return null;
    // "Is Steward" is Yes/No. Stewards are always also members.
    return u.isSteward ? 'both' : 'member';
  }
  
  function _buildUser(row, cm) {
    var isStewardVal = _str(row, cm, MH.isSteward, 'no').toLowerCase();
    var isSteward = (isStewardVal === 'yes' || isStewardVal === 'true' || isStewardVal === '1');
    
    return {
      memberId:      _str(row, cm, MH.memberId),
      firstName:     _str(row, cm, MH.firstName),
      lastName:      _str(row, cm, MH.lastName),
      name:          (_str(row, cm, MH.firstName) + ' ' + _str(row, cm, MH.lastName)).trim(),
      email:         _str(row, cm, MH.email).toLowerCase(),
      phone:         _str(row, cm, MH.phone) || null,
      jobTitle:      _str(row, cm, MH.jobTitle),
      workLocation:  _str(row, cm, MH.workLocation),
      unit:          _str(row, cm, MH.unit),
      department:    _str(row, cm, MH.department),
      officeDays:    _str(row, cm, MH.officeDays),
      supervisor:    _str(row, cm, MH.supervisor),
      manager:       _str(row, cm, MH.manager),
      isSteward:     isSteward,
      committees:    _str(row, cm, MH.committees),
      assignedSteward: _str(row, cm, MH.assignedSteward),
      hireDate:      _str(row, cm, MH.hireDate),
      employeeId:    _str(row, cm, MH.employeeId),
      duesStatus:    'Current', // Not in your schema — placeholder
      hasOpenGrievance: _str(row, cm, MH.hasOpenGrievance).toLowerCase() === 'yes',
      grievanceStatus: _str(row, cm, MH.grievanceStatus),
      prefComm:      _str(row, cm, MH.prefComm),
    };
  }
  
  // ═══════════════════════════════════════
  // STEWARD: CASES
  // ═══════════════════════════════════════
  
  function getStewardCases(stewardIdentifier) {
    // "Assigned Steward" in grievance log could be name or email
    var id = String(stewardIdentifier).trim().toLowerCase();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GRIEVANCE_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var sc = _col(cm, GH.assignedSteward);
    if (sc === -1) return [];
    
    var cases = [];
    for (var i = 1; i < data.length; i++) {
      var assigned = String(data[i][sc]).trim().toLowerCase();
      // Match on email OR name (in case column stores names)
      if (assigned === id || assigned.indexOf(id.split('@')[0]) > -1) {
        cases.push(_buildGrievance(data[i], cm));
      }
    }
    
    cases.sort(function(a, b) {
      if (a.status === 'overdue' && b.status !== 'overdue') return -1;
      if (b.status === 'overdue' && a.status !== 'overdue') return 1;
      return (a.daysToDeadline || 999) - (b.daysToDeadline || 999);
    });
    
    return cases;
  }
  
  function getStewardKPIs(stewardIdentifier) {
    var cases = getStewardCases(stewardIdentifier);
    var r = { totalCases: cases.length, activeCases: 0, overdue: 0, dueSoon: 0, resolved: 0 };
    cases.forEach(function(c) {
      var s = c.status.toLowerCase();
      if (s === 'resolved' || s === 'closed') r.resolved++;
      else if (s === 'overdue') r.overdue++;
      else {
        r.activeCases++;
        if (c.daysToDeadline !== null && c.daysToDeadline <= 7 && c.daysToDeadline > 0) r.dueSoon++;
      }
    });
    return r;
  }
  
  // ═══════════════════════════════════════
  // MEMBER: GRIEVANCES
  // ═══════════════════════════════════════
  
  function getMemberGrievances(memberEmail) {
    var email = String(memberEmail).trim().toLowerCase();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GRIEVANCE_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var ec = _col(cm, GH.memberEmail);
    if (ec === -1) return [];
    
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ec]).trim().toLowerCase() === email) {
        out.push(_buildGrievance(data[i], cm));
      }
    }
    out.sort(function(a, b) { return (b.dateFiledTs || 0) - (a.dateFiledTs || 0); });
    return out;
  }
  
  // ═══════════════════════════════════════
  // STEWARD CONTACT
  // ═══════════════════════════════════════
  
  function getStewardContact(stewardIdentifier) {
    if (!stewardIdentifier) return null;
    var id = String(stewardIdentifier).trim().toLowerCase();
    
    // Try email lookup first
    var user = findUserByEmail(id);
    if (user) return { name: user.name, email: user.email, phone: user.phone };
    
    // If not email, search by name
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var fnCol = _col(cm, MH.firstName);
    var lnCol = _col(cm, MH.lastName);
    
    for (var i = 1; i < data.length; i++) {
      var fullName = (String(data[i][fnCol] || '') + ' ' + String(data[i][lnCol] || '')).trim().toLowerCase();
      if (fullName === id) {
        var u = _buildUser(data[i], cm);
        return { name: u.name, email: u.email, phone: u.phone };
      }
    }
    return null;
  }
  
  // ═══════════════════════════════════════
  // BUILD GRIEVANCE RECORD
  // ═══════════════════════════════════════
  
  function _buildGrievance(row, cm) {
    var nextActionDue = _date(row, cm, GH.nextActionDue);
    var dtd = _val(row, cm, GH.daysToDeadline, null);
    var daysToDeadline = (dtd !== null && dtd !== '') ? parseInt(dtd, 10) : _daysUntil(nextActionDue);
    
    var status = _str(row, cm, GH.status, 'new').toLowerCase();
    if (daysToDeadline !== null && daysToDeadline < 0 && status !== 'resolved' && status !== 'closed') {
      status = 'overdue';
    }
    
    var dateFiled = _date(row, cm, GH.dateFiled);
    
    // Build timeline from actual step dates
    var timeline = [];
    
    if (dateFiled) timeline.push({ date: _fmtDate(dateFiled), event: 'Grievance filed', done: true });
    
    var s1Due = _date(row, cm, GH.stepIDue);
    var s1Rcvd = _date(row, cm, GH.stepIRcvd);
    if (s1Due) timeline.push({ date: _fmtDate(s1Due), event: 'Step I response due', done: !!s1Rcvd });
    if (s1Rcvd) timeline.push({ date: _fmtDate(s1Rcvd), event: 'Step I response received', done: true });
    
    var s2AppDue = _date(row, cm, GH.stepIIAppealDue);
    var s2AppFiled = _date(row, cm, GH.stepIIAppealFiled);
    if (s2AppDue) timeline.push({ date: _fmtDate(s2AppDue), event: 'Step II appeal due', done: !!s2AppFiled });
    if (s2AppFiled) timeline.push({ date: _fmtDate(s2AppFiled), event: 'Step II appeal filed', done: true });
    
    var s2Due = _date(row, cm, GH.stepIIDue);
    var s2Rcvd = _date(row, cm, GH.stepIIRcvd);
    if (s2Due) timeline.push({ date: _fmtDate(s2Due), event: 'Step II response due', done: !!s2Rcvd });
    if (s2Rcvd) timeline.push({ date: _fmtDate(s2Rcvd), event: 'Step II response received', done: true });
    
    var s3AppDue = _date(row, cm, GH.stepIIIAppealDue);
    var s3AppFiled = _date(row, cm, GH.stepIIIAppealFiled);
    if (s3AppDue) timeline.push({ date: _fmtDate(s3AppDue), event: 'Step III appeal due', done: !!s3AppFiled });
    if (s3AppFiled) timeline.push({ date: _fmtDate(s3AppFiled), event: 'Step III appeal filed', done: true });
    
    var closed = _date(row, cm, GH.dateClosed);
    if (closed) timeline.push({ date: _fmtDate(closed), event: 'Case closed', done: true });
    
    // Next action
    if (nextActionDue && !closed) {
      var exists = timeline.some(function(t) { return t.date === _fmtDate(nextActionDue) && !t.done; });
      if (!exists) {
        timeline.push({ date: _fmtDate(nextActionDue), event: 'Next action due', done: false });
      }
    }
    
    return {
      id:             _str(row, cm, GH.grievanceId),
      memberId:       _str(row, cm, GH.memberId),
      memberName:     (_str(row, cm, GH.firstName) + ' ' + _str(row, cm, GH.lastName)).trim(),
      memberEmail:    _str(row, cm, GH.memberEmail).toLowerCase(),
      status:         status,
      step:           _str(row, cm, GH.currentStep),
      incidentDate:   _fmtDate(_date(row, cm, GH.incidentDate)),
      filingDeadline: _fmtDate(_date(row, cm, GH.filingDeadline)),
      filed:          _fmtDate(dateFiled),
      dateFiledTs:    dateFiled ? dateFiled.getTime() : 0,
      nextActionDue:  _fmtDate(nextActionDue),
      daysToDeadline: isNaN(daysToDeadline) ? null : daysToDeadline,
      daysOpen:       _val(row, cm, GH.daysOpen, ''),
      articlesViolated: _str(row, cm, GH.articlesViolated),
      issueCategory:  _str(row, cm, GH.issueCategory),
      workLocation:   _str(row, cm, GH.workLocation),
      steward:        _str(row, cm, GH.assignedSteward),
      resolution:     _str(row, cm, GH.resolution),
      dateClosed:     _fmtDate(closed),
      checklistProgress: _str(row, cm, GH.checklistProgress),
      driveFolderUrl: _str(row, cm, GH.driveFolderUrl),
      timeline:       timeline,
    };
  }
  
  function getUnits() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var uc = _col(cm, MH.unit);
    if (uc === -1) return [];
    var u = {};
    for (var i = 1; i < data.length; i++) { var v = String(data[i][uc]).trim(); if (v) u[v] = true; }
    return Object.keys(u).sort();
  }
  
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

// Global functions (callable from client)
function dataGetStewardCases(email) { return DataService.getStewardCases(email); }
function dataGetStewardKPIs(email) { return DataService.getStewardKPIs(email); }
function dataGetMemberGrievances(email) { return DataService.getMemberGrievances(email); }
function dataGetStewardContact(id) { return DataService.getStewardContact(id); }
function dataGetUserRole(email) { return DataService.getUserRole(email); }
function dataGetUserProfile(email) { return DataService.findUserByEmail(email); }
function dataGetUnits() { return DataService.getUnits(); }
