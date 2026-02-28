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
      hearingUrl:     (function() { var u = _str(row, cm, GH.coordMessage); return (u && /^https?:\/\//i.test(u)) ? u : ''; })(),
      timeline:       timeline,
    };
  }

  // ═══════════════════════════════════════
  // MEETING INJECTOR — set hearing URL on active grievance
  // ═══════════════════════════════════════

  function setHearingUrl(grievanceId, url) {
    if (!url || !/^https?:\/\//i.test(url)) return { success: false, error: 'URL must start with https://' };
    if (url.length > 500) return { success: false, error: 'URL too long' };
    grievanceId = String(grievanceId || '').trim().toUpperCase();
    if (!grievanceId) return { success: false, error: 'Grievance ID required' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GRIEVANCE_SHEET);
    if (!sheet) return { success: false, error: 'Grievance sheet not found' };
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var ic = _col(cm, GH.grievanceId);
    var cc = _col(cm, GH.coordMessage);
    if (ic === -1) return { success: false, error: 'Grievance ID column not found' };
    if (cc === -1) return { success: false, error: 'Coordinator message column not found' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ic]).trim().toUpperCase() === grievanceId) {
        sheet.getRange(i + 1, cc + 1).setValue(url);
        return { success: true, message: 'Hearing URL set for ' + grievanceId };
      }
    }
    return { success: false, error: 'Grievance ' + grievanceId + ' not found' };
  }

  // ═══════════════════════════════════════
  // WEINGARTEN REQUEST — emergency steward alert
  // ═══════════════════════════════════════

  function submitWeingartenRequest(memberEmail, workLocation) {
    memberEmail = String(memberEmail || '').trim().toLowerCase();
    if (!memberEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(memberEmail)) {
      return { success: false, error: 'Valid email required' };
    }
    var user = findUserByEmail(memberEmail);
    if (!user) return { success: false, error: 'Member not found' };
    var steward = getStewardContact(user.assignedSteward);
    if (steward && steward.email) {
      try {
        MailApp.sendEmail({
          to: steward.email,
          subject: 'WEINGARTEN REQUEST — Immediate Representation Needed',
          body: [
            'WEINGARTEN RIGHTS REQUEST',
            '',
            'Member: ' + user.name + ' (' + memberEmail + ')',
            'Location: ' + (workLocation || user.workLocation || 'Not specified'),
            'Time: ' + new Date().toLocaleString(),
            '',
            'This member is requesting union representation for an investigatory meeting.',
            'Please contact them immediately.',
            '',
            '— DDS Dashboard Automated Alert'
          ].join('\n')
        });
      } catch (e) {
        Logger.log('Weingarten email error: ' + e.message);
      }
    }
    return { success: true, message: 'Your steward has been alerted. They will contact you shortly.' };
  }

  // ═══════════════════════════════════════
  // ENGAGEMENT LEADERBOARD
  // ═══════════════════════════════════════

  function getEngagementLeaderboard() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var officeMap = {};
    for (var i = 1; i < data.length; i++) {
      var loc = _str(data[i], cm, MH.workLocation) || 'Unknown';
      var openRate = parseFloat(_val(data[i], cm, MH.openRate, 0)) || 0;
      var lastContact = _date(data[i], cm, MH.recentContact);
      var intLocal = _str(data[i], cm, MH.intLocal).toLowerCase() === 'yes';
      var intChapter = _str(data[i], cm, MH.intChapter).toLowerCase() === 'yes';
      var intAllied = _str(data[i], cm, MH.intAllied).toLowerCase() === 'yes';
      var volunteerHrs = parseFloat(_val(data[i], cm, MH.volunteerHrs, 0)) || 0;
      if (!officeMap[loc]) {
        officeMap[loc] = { location: loc, total: 0, engagedPts: 0, openRateSum: 0, interested: 0, volunteers: 0 };
      }
      var o = officeMap[loc];
      o.total++;
      o.openRateSum += openRate;
      if (openRate >= 50) o.engagedPts++;
      if (intLocal || intChapter || intAllied) { o.engagedPts++; o.interested++; }
      if (volunteerHrs > 0) { o.engagedPts++; o.volunteers++; }
      if (lastContact) {
        var daysSince = Math.ceil((Date.now() - lastContact.getTime()) / 86400000);
        if (daysSince <= 90) o.engagedPts++;
      }
    }
    var result = [];
    for (var loc in officeMap) {
      if (officeMap.hasOwnProperty(loc)) {
        var o = officeMap[loc];
        var maxPts = o.total * 4;
        var engagementPct = maxPts > 0 ? Math.min(Math.round((o.engagedPts / maxPts) * 100), 100) : 0;
        var avgOpenRate = o.total > 0 ? Math.round(o.openRateSum / o.total) : 0;
        result.push({
          location: o.location,
          name: o.location,
          total: o.total,
          memberCount: o.total,
          engagementPct: engagementPct,
          score: engagementPct,
          avgOpenRate: avgOpenRate,
          interested: o.interested,
          activeCount: o.interested,
          volunteers: o.volunteers,
          status: engagementPct >= 60 ? 'green' : engagementPct >= 30 ? 'yellow' : 'red'
        });
      }
    }
    result.sort(function(a, b) { return b.engagementPct - a.engagementPct; });
    return result;
  }

  // ═══════════════════════════════════════
  // DIGITAL CLIPBOARD — field visit tool
  // ═══════════════════════════════════════

  function getMembersForClipboard() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var memberId = _str(data[i], cm, MH.memberId);
      if (!memberId) continue;
      if (_str(data[i], cm, MH.isSteward).toLowerCase() === 'yes') continue;
      var lastContact = _date(data[i], cm, MH.recentContact);
      var daysSince = lastContact ? Math.ceil((Date.now() - lastContact.getTime()) / 86400000) : 999;
      result.push({
        memberId: memberId,
        name: (_str(data[i], cm, MH.firstName) + ' ' + _str(data[i], cm, MH.lastName)).trim(),
        jobTitle: _str(data[i], cm, MH.jobTitle),
        workLocation: _str(data[i], cm, MH.workLocation),
        unit: _str(data[i], cm, MH.unit),
        email: _str(data[i], cm, MH.email),
        phone: _str(data[i], cm, MH.phone),
        contactNotes: _str(data[i], cm, MH.contactNotes),
        daysSinceContact: daysSince,
        needsFollowup: daysSince > 60
      });
    }
    result.sort(function(a, b) { return b.daysSinceContact - a.daysSinceContact; });
    return result;
  }

  function logClipboardAction(memberId, action, notes) {
    if (!memberId || !action) return { success: false, error: 'Member ID and action required' };
    var validActions = ['Signed Petition', 'Needs Follow-up', 'Contact Info Updated', 'In-Person Contact', 'Declined'];
    if (validActions.indexOf(action) === -1) return { success: false, error: 'Invalid action' };
    memberId = String(memberId).trim();
    var safeNotes = notes ? String(notes).substring(0, 300) : '';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return { success: false, error: 'Member sheet not found' };
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var idCol = _col(cm, MH.memberId);
    var notesCol = _col(cm, MH.contactNotes);
    var dateCol = _col(cm, MH.recentContact);
    if (idCol === -1) return { success: false, error: 'Member ID column not found' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === memberId) {
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yy');
        var newNote = '[' + timestamp + '] ' + action.replace(/'/g, '') + (safeNotes ? ': ' + safeNotes : '');
        var existing = notesCol !== -1 ? String(data[i][notesCol] || '') : '';
        var combined = existing ? (newNote + '\n' + existing).substring(0, 500) : newNote;
        if (notesCol !== -1) sheet.getRange(i + 1, notesCol + 1).setValue(combined);
        if (dateCol !== -1) sheet.getRange(i + 1, dateCol + 1).setValue(new Date());
        return { success: true, message: action + ' logged' };
      }
    }
    return { success: false, error: 'Member not found' };
  }

  // ═══════════════════════════════════════
  // PRECEDENT & WIN ARCHIVE
  // ═══════════════════════════════════════

  function getPrecedentArchive(articleFilter) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GRIEVANCE_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var status = _str(data[i], cm, GH.status).toLowerCase();
      if (status !== 'resolved' && status !== 'closed') continue;
      var resolution = _str(data[i], cm, GH.resolution);
      if (!resolution) continue;
      var articles = _str(data[i], cm, GH.articlesViolated);
      if (articleFilter && String(articleFilter).trim() &&
          articles.toLowerCase().indexOf(String(articleFilter).toLowerCase()) === -1) continue;
      var closed = _date(data[i], cm, GH.dateClosed);
      results.push({
        id: _str(data[i], cm, GH.grievanceId),
        articlesViolated: articles,
        issueCategory: _str(data[i], cm, GH.issueCategory),
        resolution: resolution,
        workLocation: _str(data[i], cm, GH.workLocation),
        dateClosed: closed ? _fmtDate(closed) : '',
        step: _str(data[i], cm, GH.currentStep)
      });
    }
    results.sort(function(a, b) { return a.articlesViolated.localeCompare(b.articlesViolated); });
    return results;
  }

  // ═══════════════════════════════════════
  // NUDGE ENGINE
  // ═══════════════════════════════════════

  function getSurveyNonResponders() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var email = _str(data[i], cm, MH.email);
      if (!email) continue;
      if (_str(data[i], cm, MH.isSteward).toLowerCase() === 'yes') continue;
      var openRate = parseFloat(_val(data[i], cm, MH.openRate, 0)) || 0;
      var lastContact = _date(data[i], cm, MH.recentContact);
      var daysSince = lastContact ? Math.ceil((Date.now() - lastContact.getTime()) / 86400000) : 999;
      if (openRate < 30 || daysSince > 90) {
        results.push({
          memberId: _str(data[i], cm, MH.memberId),
          name: (_str(data[i], cm, MH.firstName) + ' ' + _str(data[i], cm, MH.lastName)).trim(),
          email: email,
          workLocation: _str(data[i], cm, MH.workLocation),
          openRate: openRate,
          daysSinceContact: daysSince
        });
      }
    }
    results.sort(function(a, b) { return b.daysSinceContact - a.daysSinceContact; });
    return results.slice(0, 100);
  }

  function sendSurveyNudge(memberIds, customMessage) {
    if (!memberIds || !memberIds.length) return { success: false, error: 'No members selected' };
    if (memberIds.length > 50) return { success: false, error: 'Maximum 50 per batch' };
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBER_SHEET);
    if (!sheet) return { success: false, error: 'Member sheet not found' };
    var data = sheet.getDataRange().getValues();
    var cm = _colMap(data[0]);
    var idCol = _col(cm, MH.memberId);
    var emailCol = _col(cm, MH.email);
    var fnCol = _col(cm, MH.firstName);
    if (idCol === -1 || emailCol === -1) return { success: false, error: 'Required columns not found' };
    var defaultMsg = customMessage
      ? String(customMessage).substring(0, 500)
      : 'We want to hear from you! Please take a few minutes to complete our member satisfaction survey. Your feedback helps us fight for better conditions for everyone.';
    var idSet = {};
    memberIds.forEach(function(id) { idSet[String(id).trim()] = true; });
    var sent = 0;
    var errors = [];
    for (var i = 1; i < data.length; i++) {
      var mId = String(data[i][idCol] || '').trim();
      if (!idSet[mId]) continue;
      var email = String(data[i][emailCol] || '').trim();
      if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;
      var firstName = fnCol !== -1 ? String(data[i][fnCol] || 'Member') : 'Member';
      try {
        MailApp.sendEmail({
          to: email,
          subject: 'Your Voice Matters — Quick Union Survey',
          body: 'Hi ' + firstName + ',\n\n' + defaultMsg + '\n\n— Your Union Steward Team'
        });
        sent++;
      } catch (e) {
        errors.push(email + ': ' + e.message);
      }
    }
    return { success: true, sent: sent, errors: errors,
             message: 'Nudge sent to ' + sent + ' member' + (sent !== 1 ? 's' : '') };
  }

  // ═══════════════════════════════════════
  // UPCOMING MEETINGS — for event calendar
  // ═══════════════════════════════════════

  function getUpcomingMeetings() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('\uD83D\uDCDD Meeting Check-In Log') ||
                ss.getSheetByName('Meeting Check-In Log') ||
                ss.getSheetByName('_Meeting_Check-In_Log');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var hm = {};
    for (var h = 0; h < headers.length; h++) {
      hm[String(headers[h]).trim().toLowerCase()] = h;
    }
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var cutoff = new Date(today.getTime() + 30 * 86400000);
    var seen = {};
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var midCol = hm['meeting id'];
      var meetingId = midCol !== undefined ? String(data[i][midCol] || '') : '';
      if (!meetingId || seen[meetingId]) continue;
      seen[meetingId] = true;
      var dateCol = hm['meeting date'];
      var dateVal = dateCol !== undefined ? data[i][dateCol] : null;
      if (!(dateVal instanceof Date)) continue;
      var d = new Date(dateVal); d.setHours(0, 0, 0, 0);
      if (d < today || d > cutoff) continue;
      var statusCol = hm['event status'];
      if (statusCol !== undefined && String(data[i][statusCol] || '') === 'Completed') continue;
      var nameCol = hm['meeting name'];
      var typeCol = hm['meeting type'];
      var timeCol = hm['meeting time'];
      var agendaCol = hm['agenda doc url'];
      var name = nameCol !== undefined ? String(data[i][nameCol] || '') : '';
      var type = typeCol !== undefined ? String(data[i][typeCol] || '') : '';
      var time = timeCol !== undefined ? String(data[i][timeCol] || '') : '';
      var agendaUrl = agendaCol !== undefined ? String(data[i][agendaCol] || '') : '';
      results.push({
        id: meetingId,
        name: name,
        date: d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }),
        dateISO: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        time: time,
        type: type,
        agendaUrl: agendaUrl && /^https:\/\/docs\.google\.com\//.test(agendaUrl) ? agendaUrl : ''
      });
    }
    results.sort(function(a, b) { return a.dateISO < b.dateISO ? -1 : a.dateISO > b.dateISO ? 1 : 0; });
    return results.slice(0, 10);
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
    setHearingUrl: setHearingUrl,
    submitWeingartenRequest: submitWeingartenRequest,
    getEngagementLeaderboard: getEngagementLeaderboard,
    getMembersForClipboard: getMembersForClipboard,
    logClipboardAction: logClipboardAction,
    getPrecedentArchive: getPrecedentArchive,
    getSurveyNonResponders: getSurveyNonResponders,
    sendSurveyNudge: sendSurveyNudge,
    getUpcomingMeetings: getUpcomingMeetings,
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
function dataSetHearingUrl(grievanceId, url) { return DataService.setHearingUrl(grievanceId, url); }
function dataSubmitWeingartenRequest(memberEmail, workLocation) { return DataService.submitWeingartenRequest(memberEmail, workLocation); }
function dataGetEngagementLeaderboard() { return DataService.getEngagementLeaderboard(); }
function dataGetMembersForClipboard() { return DataService.getMembersForClipboard(); }
function dataLogClipboardAction(memberId, action, notes) { return DataService.logClipboardAction(memberId, action, notes); }
function dataGetPrecedentArchive(articleFilter) { return DataService.getPrecedentArchive(articleFilter || ''); }
function dataGetSurveyNonResponders() { return DataService.getSurveyNonResponders(); }
function dataSendSurveyNudge(memberIds, customMessage) { return DataService.sendSurveyNudge(memberIds, customMessage || ''); }
function dataGetUpcomingMeetings() { return DataService.getUpcomingMeetings(); }
