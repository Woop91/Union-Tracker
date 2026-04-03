/**
 * ============================================================================
 * 33_NewFeatureServices.gs — Feature Services Module
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Contains multiple IIFE service modules for new dashboard features:
 *   - HandoffService     — Steward shift handoff notes per case
 *   - MentorshipService  — Steward mentorship pairing and tracking
 *   - CommunicationLogService — Member outreach/communication tracking
 *   - KnowledgeBaseService — Contract article reference / knowledge base
 *   - DigestService      — Smart notification digest batching
 *   - DocumentChecklistService — Per-step document checklist with Drive scan
 *   - EscalationEngine   — Automated case escalation recommendations
 *   - ReportService      — Automated report generation
 *   - SMSService         — Twilio SMS integration (v4.36.0)
 *   - RSVPService        — Meeting RSVP invitation tracking (v4.36.0)
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Consolidates 8 small feature services into one file to stay within
 *   the 50-file GAS deployment limit. Each service follows the established
 *   IIFE closure pattern (same as QAForum, TimelineService, FailsafeService).
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   New feature tabs/panels in the SPA degrade gracefully — handoff notes
 *   won't display, mentorship pairings unavailable, etc. Core grievance
 *   tracking, auth, and existing features are unaffected.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS), 00_Security.gs (escapeForFormula),
 *   21_WebDashDataService.gs (_resolveCallerEmail, _requireStewardAuth)
 *
 * @version 4.35.0
 */

// ============================================================================
// HANDOFF SERVICE — Steward Shift Handoff Notes
// ============================================================================

var HandoffService = (function () {

  var HEADERS = ['ID', 'Case ID', 'From Steward', 'To Steward', 'Note Text', 'Created', 'Status'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.HANDOFF_NOTES);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.HANDOFF_NOTES);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  function getHandoffNotes(caseId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.HANDOFF_NOTES);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(caseId).trim() && String(data[i][6]).trim() !== 'archived') {
        results.push({
          id: String(data[i][0]),
          caseId: String(data[i][1]),
          fromSteward: String(data[i][2]),
          toSteward: String(data[i][3]),
          noteText: String(data[i][4]),
          created: data[i][5],
          status: String(data[i][6])
        });
      }
    }
    results.sort(function (a, b) {
      return new Date(b.created) - new Date(a.created);
    });
    return results;
  }

  function addHandoffNote(stewardEmail, caseId, noteText, toSteward) {
    var sheet = initSheet();
    var id = Utilities.getUuid();
    var now = new Date();
    var safeNote = typeof escapeForFormula === 'function' ? escapeForFormula(noteText) : noteText;
    var safeTo = typeof escapeForFormula === 'function' ? escapeForFormula(toSteward || '') : (toSteward || '');
    sheet.appendRow([id, String(caseId), stewardEmail, safeTo, safeNote, now, 'active']);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('HANDOFF_NOTE_ADDED', { caseId: caseId, steward: stewardEmail });
    }
    return { success: true, id: id };
  }

  function archiveHandoffNote(noteId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.HANDOFF_NOTES);
    if (!sheet) return { success: false, message: 'Sheet not found.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(noteId)) {
        sheet.getRange(i + 1, 7).setValue('archived');
        return { success: true };
      }
    }
    return { success: false, message: 'Note not found.' };
  }

  return {
    initSheet: initSheet,
    getHandoffNotes: getHandoffNotes,
    addHandoffNote: addHandoffNote,
    archiveHandoffNote: archiveHandoffNote
  };
})();

// ============================================================================
// MENTORSHIP SERVICE — Steward Mentorship Pairing
// ============================================================================

var MentorshipService = (function () {

  var HEADERS = ['ID', 'Mentor Email', 'Mentee Email', 'Case Types', 'Status', 'Started', 'Notes'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.MENTORSHIP);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.MENTORSHIP);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  function getPairings() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.MENTORSHIP);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim() !== 'closed') {
        results.push({
          id: String(data[i][0]),
          mentorEmail: String(data[i][1]),
          menteeEmail: String(data[i][2]),
          caseTypes: String(data[i][3]),
          status: String(data[i][4]),
          started: data[i][5],
          notes: String(data[i][6])
        });
      }
    }
    return results;
  }

  function createPairing(mentorEmail, menteeEmail, caseTypes) {
    if (!mentorEmail || !menteeEmail) return { success: false, message: 'Mentor and mentee emails are required.' };
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(mentorEmail) || !emailPattern.test(menteeEmail)) {
      return { success: false, message: 'Invalid email format.' };
    }
    if (mentorEmail.toLowerCase().trim() === menteeEmail.toLowerCase().trim()) {
      return { success: false, message: 'Mentor and mentee cannot be the same person.' };
    }

    // Check for duplicate active pairing
    var existing = getPairings();
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].mentorEmail.toLowerCase() === mentorEmail.toLowerCase().trim() &&
          existing[i].menteeEmail.toLowerCase() === menteeEmail.toLowerCase().trim()) {
        return { success: false, message: 'This pairing already exists.' };
      }
    }

    var sheet = initSheet();
    var id = Utilities.getUuid();
    var now = new Date();
    var esc = typeof escapeForFormula === 'function' ? escapeForFormula : function (v) { return v; };
    sheet.appendRow([id, esc(mentorEmail.trim()), esc(menteeEmail.trim()), esc(caseTypes || ''), 'active', now, '']);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MENTORSHIP_CREATED', { mentor: mentorEmail, mentee: menteeEmail });
    }
    if (typeof _refreshNavBadges === 'function') _refreshNavBadges();
    return { success: true, id: id };
  }

  function updatePairingNotes(pairingId, notes) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.MENTORSHIP);
    if (!sheet) return { success: false, message: 'Sheet not found.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(pairingId)) {
        var safeNotes = typeof escapeForFormula === 'function' ? escapeForFormula(notes) : notes;
        sheet.getRange(i + 1, 7).setValue(safeNotes);
        if (typeof _refreshNavBadges === 'function') _refreshNavBadges();
        return { success: true };
      }
    }
    return { success: false, message: 'Pairing not found.' };
  }

  function closePairing(pairingId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'No spreadsheet.' };
    var sheet = ss.getSheetByName(SHEETS.MENTORSHIP);
    if (!sheet) return { success: false, message: 'Sheet not found.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(pairingId)) {
        sheet.getRange(i + 1, 5).setValue('closed');
        if (typeof logAuditEvent === 'function') {
          logAuditEvent('MENTORSHIP_CLOSED', { pairingId: pairingId, mentor: String(data[i][1]), mentee: String(data[i][2]) });
        }
        if (typeof _refreshNavBadges === 'function') _refreshNavBadges();
        return { success: true };
      }
    }
    return { success: false, message: 'Pairing not found.' };
  }

  function suggestPairings() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet) return [];
    var data = memberSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var roleCol = -1, emailCol = -1, joinedCol = -1;
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (h === 'role') roleCol = c;
      if (h === 'email' || h === 'email address') emailCol = c;
      if (h === 'joined' || h === 'date joined') joinedCol = c;
    }
    if (roleCol === -1 || emailCol === -1) return [];

    var stewards = [];
    var sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);

    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][roleCol]).toLowerCase().trim();
      if (role === 'steward' || role === 'both') {
        var joined = joinedCol >= 0 ? new Date(data[i][joinedCol]) : null;
        stewards.push({
          email: String(data[i][emailCol]).trim().toLowerCase(),
          isNew: joined && joined > sixMonthsAgo,
          joined: joined
        });
      }
    }

    // Find existing pairings to avoid duplicates
    var existing = getPairings();
    var pairedEmails = {};
    existing.forEach(function (p) {
      pairedEmails[p.menteeEmail.toLowerCase()] = true;
    });

    var suggestions = [];
    var experienced = stewards.filter(function (s) { return !s.isNew; });
    var newStewards = stewards.filter(function (s) { return s.isNew && !pairedEmails[s.email]; });

    // Track mentor load (existing active pairings + new suggestions)
    var mentorLoad = {};
    experienced.forEach(function (s) { mentorLoad[s.email] = 0; });
    existing.forEach(function (p) {
      var me = p.mentorEmail.toLowerCase();
      if (mentorLoad.hasOwnProperty(me)) mentorLoad[me]++;
    });

    newStewards.forEach(function (mentee) {
      if (experienced.length > 0) {
        // Pick mentor with lowest current load
        var bestMentor = experienced[0];
        var bestLoad = mentorLoad[bestMentor.email] || 0;
        for (var m = 1; m < experienced.length; m++) {
          var load = mentorLoad[experienced[m].email] || 0;
          if (load < bestLoad) { bestMentor = experienced[m]; bestLoad = load; }
        }
        mentorLoad[bestMentor.email] = bestLoad + 1;
        suggestions.push({
          mentorEmail: bestMentor.email,
          menteeEmail: mentee.email,
          reason: 'New steward (joined ' + (mentee.joined ? mentee.joined.toLocaleDateString() : 'recently') + ')'
        });
      }
    });

    return suggestions;
  }

  return {
    initSheet: initSheet,
    getPairings: getPairings,
    createPairing: createPairing,
    updatePairingNotes: updatePairingNotes,
    closePairing: closePairing,
    suggestPairings: suggestPairings
  };
})();

// ============================================================================
// COMMUNICATION LOG SERVICE — Member Outreach Tracking
// ============================================================================

var CommunicationLogService = (function () {

  var HEADERS = ['ID', 'Member Email', 'Steward Email', 'Type', 'Subject', 'Notes', 'Timestamp'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.COMMUNICATION_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.COMMUNICATION_LOG);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  function logCommunication(stewardEmail, memberEmail, type, subject, notes) {
    var sheet = initSheet();
    var id = Utilities.getUuid();
    var now = new Date();
    var safeSubject = typeof escapeForFormula === 'function' ? escapeForFormula(subject || '') : (subject || '');
    var safeNotes = typeof escapeForFormula === 'function' ? escapeForFormula(notes || '') : (notes || '');
    sheet.appendRow([id, escapeForFormula(memberEmail), escapeForFormula(stewardEmail), type || 'other', safeSubject, safeNotes, now]);
    return { success: true, id: id };
  }

  function getCommunicationLog(memberEmail) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.COMMUNICATION_LOG);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var results = [];
    var target = String(memberEmail).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === target) {
        results.push({
          id: String(data[i][0]),
          memberEmail: String(data[i][1]),
          stewardEmail: String(data[i][2]),
          type: String(data[i][3]),
          subject: String(data[i][4]),
          notes: String(data[i][5]),
          timestamp: data[i][6]
        });
      }
    }
    results.sort(function (a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    return results;
  }

  function getStewardCommunicationSummary(stewardEmail) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { total: 0, byType: {} };
    var sheet = ss.getSheetByName(SHEETS.COMMUNICATION_LOG);
    if (!sheet) return { total: 0, byType: {} };
    var data = sheet.getDataRange().getValues();
    var total = 0;
    var byType = {};
    var target = String(stewardEmail).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === target) {
        total++;
        var type = String(data[i][3]) || 'other';
        byType[type] = (byType[type] || 0) + 1;
      }
    }
    return { total: total, byType: byType };
  }

  return {
    initSheet: initSheet,
    logCommunication: logCommunication,
    getCommunicationLog: getCommunicationLog,
    getStewardCommunicationSummary: getStewardCommunicationSummary
  };
})();

// ============================================================================
// KNOWLEDGE BASE SERVICE — Contract Article Reference
// ============================================================================

var KnowledgeBaseService = (function () {

  var HEADERS = ['ID', 'Article Number', 'Title', 'Summary', 'Full Text', 'Related Grievance Types', 'Tags'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_BASE);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.KNOWLEDGE_BASE);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  function searchArticles(query) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_BASE);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var q = String(query).toLowerCase().trim();
    if (!q) return _getAllArticles(data);
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var searchable = (String(data[i][1]) + ' ' + String(data[i][2]) + ' ' +
        String(data[i][3]) + ' ' + String(data[i][6])).toLowerCase();
      if (searchable.indexOf(q) !== -1) {
        results.push(_buildArticle(data[i]));
      }
    }
    return results;
  }

  function getArticle(articleId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_BASE);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(articleId)) {
        return _buildArticle(data[i], true);
      }
    }
    return null;
  }

  function getRelatedArticles(grievanceType) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_BASE);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var target = String(grievanceType).toLowerCase().trim();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var related = String(data[i][5]).toLowerCase();
      if (related.indexOf(target) !== -1) {
        results.push(_buildArticle(data[i]));
      }
    }
    return results;
  }

  function addArticle(articleNumber, title, summary, fullText, relatedTypes, tags) {
    var sheet = initSheet();
    var id = Utilities.getUuid();
    var esc = typeof escapeForFormula === 'function' ? escapeForFormula : function (v) { return v; };
    sheet.appendRow([id, esc(articleNumber), esc(title), esc(summary), esc(fullText), esc(relatedTypes || ''), esc(tags || '')]);
    return { success: true, id: id };
  }

  function _buildArticle(row, includeFullText) {
    var article = {
      id: String(row[0]),
      articleNumber: String(row[1]),
      title: String(row[2]),
      summary: String(row[3]),
      relatedGrievanceTypes: String(row[5]),
      tags: String(row[6])
    };
    if (includeFullText) {
      article.fullText = String(row[4]);
    }
    return article;
  }

  function _getAllArticles(data) {
    var results = [];
    for (var i = 1; i < data.length; i++) {
      results.push(_buildArticle(data[i]));
    }
    return results;
  }

  return {
    initSheet: initSheet,
    searchArticles: searchArticles,
    getArticle: getArticle,
    getRelatedArticles: getRelatedArticles,
    addArticle: addArticle
  };
})();

// ============================================================================
// DIGEST SERVICE — Smart Notification Digest
// ============================================================================

var DigestService = (function () {

  var HEADERS = ['Email', 'Frequency', 'Types', 'Last Sent'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATION_PREFS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.NOTIFICATION_PREFS);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  function getPreferences(email) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { frequency: 'immediate', types: 'all' };
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATION_PREFS);
    if (!sheet) return { frequency: 'immediate', types: 'all' };
    var data = sheet.getDataRange().getValues();
    var target = String(email).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === target) {
        return {
          frequency: String(data[i][1]) || 'immediate',
          types: String(data[i][2]) || 'all',
          lastSent: data[i][3]
        };
      }
    }
    return { frequency: 'immediate', types: 'all' };
  }

  function setPreferences(email, frequency, types) {
    var sheet = initSheet();
    var data = sheet.getDataRange().getValues();
    var target = String(email).trim().toLowerCase();
    var validFreqs = ['immediate', 'daily', 'weekly'];
    var freq = validFreqs.indexOf(frequency) !== -1 ? frequency : 'immediate';
    var safeTypes = typeof escapeForFormula === 'function' ? escapeForFormula(types || 'all') : (types || 'all');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === target) {
        sheet.getRange(i + 1, 2).setValue(freq);
        sheet.getRange(i + 1, 3).setValue(safeTypes);
        return { success: true };
      }
    }
    // New entry
    sheet.appendRow([email, freq, safeTypes, '']);
    return { success: true };
  }

  function buildDigestContent(email) {
    // Aggregate notifications for a user
    var digest = {
      newCases: 0,
      approachingDeadlines: 0,
      overdueItems: 0,
      qaActivity: 0,
      surveyResponses: 0,
      items: []
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return digest;

    // Count approaching deadlines from Grievance Log
    var grievSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievSheet) {
      var grievData = grievSheet.getDataRange().getValues();
      if (grievData.length > 1) {
        var headers = grievData[0];
        var deadlineCol = -1, statusCol = -1;
        for (var c = 0; c < headers.length; c++) {
          var h = String(headers[c]).toLowerCase().trim();
          if (h === 'deadline' || h === 'next deadline') deadlineCol = c;
          if (h === 'status') statusCol = c;
        }
        if (deadlineCol >= 0) {
          var now = new Date();
          var sevenDays = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
          for (var i = 1; i < grievData.length; i++) {
            var status = String(grievData[i][statusCol] || '').toLowerCase();
            if (status === 'resolved' || status === 'closed' || status === 'denied' || status === 'settled') continue;
            var dl = grievData[i][deadlineCol];
            if (dl instanceof Date) {
              if (dl < now) digest.overdueItems++;
              else if (dl <= sevenDays) digest.approachingDeadlines++;
            }
          }
        }
      }
    }

    // Count unanswered Q&A
    if (typeof QAForum !== 'undefined' && typeof QAForum.getUnansweredCount === 'function') {
      try { digest.qaActivity = QAForum.getUnansweredCount(); } catch (_) { log_('_', (_.message || _)); }
    }

    return digest;
  }

  function sendScheduledDigests() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATION_PREFS);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var today = now.toDateString();
    var sent = 0;

    for (var i = 1; i < data.length; i++) {
      var email = String(data[i][0]).trim();
      var freq = String(data[i][1]).trim();
      var lastSent = data[i][3] ? new Date(data[i][3]).toDateString() : '';

      if (freq === 'immediate') continue;
      if (lastSent === today) continue; // Already sent today

      if (freq === 'weekly') {
        // Only send on Mondays
        if (now.getDay() !== 1) continue;
      }

      var digest = buildDigestContent(email);
      var total = digest.newCases + digest.approachingDeadlines + digest.overdueItems + digest.qaActivity;
      if (total === 0) continue; // Nothing to report

      // Build and send digest email
      var subject = 'Dashboard Digest: ' + total + ' item(s) need attention';
      var body = '<h2>Your Dashboard Digest</h2>';
      if (digest.overdueItems > 0) body += '<p>⚠️ <strong>' + digest.overdueItems + '</strong> overdue case(s)</p>';
      if (digest.approachingDeadlines > 0) body += '<p>📅 <strong>' + digest.approachingDeadlines + '</strong> deadline(s) approaching (within 7 days)</p>';
      if (digest.qaActivity > 0) body += '<p>❓ <strong>' + digest.qaActivity + '</strong> unanswered question(s)</p>';
      body += '<p><em>Sent by Dashboard Digest Service</em></p>';

      try {
        if (typeof safeSendEmail_ === 'function') {
          safeSendEmail_(email, subject, body);
        } else if (MailApp.getRemainingDailyQuota() > 5) {
          MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
        }
        sheet.getRange(i + 1, 4).setValue(now);
        sent++;
      } catch (e) {
        log_('sendScheduledDigests', 'Digest send failed for ' + email + ': ' + e.message);
      }
    }
    log_('Digest service', 'sent ' + sent + ' digest email(s).');
  }

  return {
    initSheet: initSheet,
    getPreferences: getPreferences,
    setPreferences: setPreferences,
    buildDigestContent: buildDigestContent,
    sendScheduledDigests: sendScheduledDigests
  };
})();

// ============================================================================
// DOCUMENT CHECKLIST SERVICE — Per-Step Document Requirements
// ============================================================================

var DocumentChecklistService = (function () {

  // Default required documents per grievance step
  var STEP_REQUIREMENTS = {
    'Step I': ['Grievance Form', 'Written Statement'],
    'Step II': ['Step I Decision', 'Appeal Letter', 'Supporting Evidence'],
    'Step III': ['Step II Decision', 'Escalation Request', 'Case Summary'],
    'Arbitration': ['Step III Decision', 'Arbitration Demand', 'Full Case File', 'Witness Statements']
  };

  function getChecklist(caseId, currentStep, driveFolderId) {
    var requirements = STEP_REQUIREMENTS[currentStep] || [];
    var checklist = requirements.map(function (doc) {
      return { name: doc, required: true, found: false, fileUrl: null };
    });

    if (!driveFolderId) return checklist;

    // Scan Drive folder for existing files
    try {
      var folder = DriveApp.getFolderById(driveFolderId);
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var fileName = file.getName().toLowerCase();
        for (var j = 0; j < checklist.length; j++) {
          if (fileName.indexOf(checklist[j].name.toLowerCase()) !== -1) {
            checklist[j].found = true;
            checklist[j].fileUrl = file.getUrl();
          }
        }
      }
    } catch (e) {
      log_('DocumentChecklist', 'folder scan failed for ' + driveFolderId + ': ' + e.message);
    }

    return checklist;
  }

  function getCompletionStatus(caseId, currentStep, driveFolderId) {
    var checklist = getChecklist(caseId, currentStep, driveFolderId);
    var total = checklist.length;
    var found = checklist.filter(function (item) { return item.found; }).length;
    return {
      total: total,
      found: found,
      complete: total > 0 && found === total,
      percentage: total > 0 ? Math.round((found / total) * 100) : 0
    };
  }

  return {
    getChecklist: getChecklist,
    getCompletionStatus: getCompletionStatus
  };
})();

// ============================================================================
// ESCALATION ENGINE — Case Escalation Recommendations
// ============================================================================

var EscalationEngine = (function () {

  var STEP_ORDER = ['Step I', 'Step II', 'Step III', 'Arbitration'];

  function getRecommendation(caseId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    var headers = data[0];
    var cols = {};
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (h === 'id' || h === 'grievance id') cols.id = c;
      if (h === 'status') cols.status = c;
      if (h === 'step' || h === 'current step') cols.step = c;
      if (h === 'deadline' || h === 'next deadline') cols.deadline = c;
      if (h === 'issue category' || h === 'category') cols.category = c;
      if (h === 'unit') cols.unit = c;
    }

    // Find the target case
    var targetCase = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cols.id]).trim() === String(caseId).trim()) {
        targetCase = {
          row: i,
          status: String(data[i][cols.status] || ''),
          step: String(data[i][cols.step] || ''),
          deadline: data[i][cols.deadline],
          category: String(data[i][cols.category] || ''),
          unit: String(data[i][cols.unit] || '')
        };
        break;
      }
    }
    if (!targetCase) return null;

    var reasons = [];
    var confidence = 'low';
    var currentStepIdx = STEP_ORDER.indexOf(targetCase.step);
    var suggestedStep = currentStepIdx >= 0 && currentStepIdx < STEP_ORDER.length - 1
      ? STEP_ORDER[currentStepIdx + 1] : null;

    // Check 1: Deadline proximity
    if (targetCase.deadline instanceof Date) {
      var now = new Date();
      var daysRemaining = Math.ceil((targetCase.deadline - now) / (1000 * 60 * 60 * 24));
      if (daysRemaining < 0) {
        reasons.push('Deadline has passed (' + Math.abs(daysRemaining) + ' day(s) overdue)');
        confidence = 'high';
      } else if (daysRemaining <= 3) {
        reasons.push('Deadline approaching (' + daysRemaining + ' day(s) remaining)');
        if (confidence === 'low') confidence = 'medium';
      }
    }

    // Check 2: Denial pattern — look at status
    var statusLower = targetCase.status.toLowerCase();
    if (statusLower === 'denied' || statusLower === 'rejected') {
      reasons.push('Case was denied at ' + targetCase.step + ' — escalation is the standard next action');
      confidence = 'high';
    }

    // Check 3: Similar past cases — look at outcomes for same category/unit
    if (targetCase.category && cols.category !== undefined) {
      var similarWins = 0;
      var similarTotal = 0;
      for (var j = 1; j < data.length; j++) {
        if (j === targetCase.row) continue;
        var cat = String(data[j][cols.category] || '').toLowerCase();
        var st = String(data[j][cols.status] || '').toLowerCase();
        if (cat === targetCase.category.toLowerCase() && (st === 'resolved' || st === 'settled' || st === 'won' || st === 'denied' || st === 'lost')) {
          similarTotal++;
          if (st === 'resolved' || st === 'settled' || st === 'won') similarWins++;
        }
      }
      if (similarTotal >= 3) {
        var winRate = Math.round((similarWins / similarTotal) * 100);
        if (winRate >= 60) {
          reasons.push('Similar cases (' + targetCase.category + ') have ' + winRate + '% win rate — escalation historically favorable');
          if (confidence === 'low') confidence = 'medium';
        }
      }
    }

    return {
      shouldEscalate: reasons.length > 0,
      confidence: confidence,
      reasons: reasons,
      suggestedStep: suggestedStep,
      currentStep: targetCase.step
    };
  }

  return {
    getRecommendation: getRecommendation
  };
})();

// ============================================================================
// REPORT SERVICE — Automated Report Generation
// ============================================================================

var ReportService = (function () {

  function generateMonthlyReport(dateRange) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'No spreadsheet.' };

    var now = new Date();
    var startDate = dateRange && dateRange.start ? new Date(dateRange.start) : new Date(now.getFullYear(), now.getMonth() - 1, 1);
    var endDate = dateRange && dateRange.end ? new Date(dateRange.end) : new Date(now.getFullYear(), now.getMonth(), 0);

    var report = {
      title: 'Monthly Grievance Summary',
      period: startDate.toLocaleDateString() + ' — ' + endDate.toLocaleDateString(),
      generated: now.toISOString(),
      summary: { totalCases: 0, newCases: 0, resolved: 0, pending: 0, overdue: 0 },
      byCategory: {},
      bySteward: {},
      byUnit: {}
    };

    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) return { success: true, report: report };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, report: report };

    var headers = data[0];
    var cols = {};
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).toLowerCase().trim();
      if (h === 'status') cols.status = c;
      if (h === 'filed' || h === 'date filed') cols.filed = c;
      if (h === 'date closed') cols.closed = c;
      if (h === 'deadline' || h === 'next deadline') cols.deadline = c;
      if (h === 'issue category' || h === 'category') cols.category = c;
      if (h === 'assigned steward' || h === 'steward email') cols.steward = c;
      if (h === 'unit') cols.unit = c;
    }

    for (var i = 1; i < data.length; i++) {
      var filed = cols.filed !== undefined ? data[i][cols.filed] : null;
      var status = String(data[i][cols.status] || '').toLowerCase();
      var category = String(data[i][cols.category] || 'Uncategorized');
      var steward = String(data[i][cols.steward] || 'Unassigned');
      var unit = String(data[i][cols.unit] || 'Unknown');

      report.summary.totalCases++;

      if (filed instanceof Date && filed >= startDate && filed <= endDate) {
        report.summary.newCases++;
      }

      if (status === 'resolved' || status === 'settled' || status === 'won') {
        report.summary.resolved++;
      } else if (status === 'closed' || status === 'denied' || status === 'lost') {
        // Count but don't add to resolved
      } else {
        report.summary.pending++;
        var dl = cols.deadline !== undefined ? data[i][cols.deadline] : null;
        if (dl instanceof Date && dl < now) report.summary.overdue++;
      }

      // Aggregate by category
      report.byCategory[category] = (report.byCategory[category] || 0) + 1;

      // Aggregate by steward
      report.bySteward[steward] = (report.bySteward[steward] || 0) + 1;

      // Aggregate by unit
      report.byUnit[unit] = (report.byUnit[unit] || 0) + 1;
    }

    return { success: true, report: report };
  }

  function generateReportHtml(report) {
    if (!report) return '<p>No report data available.</p>';
    var html = '<h1>' + (typeof escapeHtml === 'function' ? escapeHtml(report.title) : report.title) + '</h1>';
    html += '<p>Period: ' + (typeof escapeHtml === 'function' ? escapeHtml(report.period) : report.period) + '</p>';
    html += '<h2>Summary</h2>';
    html += '<table border="1" cellpadding="8" style="border-collapse:collapse;">';
    html += '<tr><td>Total Cases</td><td>' + report.summary.totalCases + '</td></tr>';
    html += '<tr><td>New Cases</td><td>' + report.summary.newCases + '</td></tr>';
    html += '<tr><td>Resolved</td><td>' + report.summary.resolved + '</td></tr>';
    html += '<tr><td>Pending</td><td>' + report.summary.pending + '</td></tr>';
    html += '<tr><td>Overdue</td><td>' + report.summary.overdue + '</td></tr>';
    html += '</table>';

    if (Object.keys(report.byCategory).length > 0) {
      html += '<h2>By Category</h2><ul>';
      for (var cat in report.byCategory) {
        html += '<li>' + (typeof escapeHtml === 'function' ? escapeHtml(cat) : cat) + ': ' + report.byCategory[cat] + '</li>';
      }
      html += '</ul>';
    }

    return html;
  }

  return {
    generateMonthlyReport: generateMonthlyReport,
    generateReportHtml: generateReportHtml
  };
})();

// ============================================================================
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ============================================================================

// --- Handoff Notes (steward-only) ---
/** @param {string} sessionToken @param {string} caseId @returns {Array} Handoff notes for case. */
function dataGetHandoffNotes(sessionToken, caseId) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return HandoffService.getHandoffNotes(caseId); }
/** @param {string} sessionToken @param {string} caseId @param {string} noteText @param {string} [toSteward] @returns {Object} Result. */
function dataAddHandoffNote(sessionToken, caseId, noteText, toSteward) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return HandoffService.addHandoffNote(s, caseId, noteText, toSteward); }
/** @param {string} sessionToken @param {string} noteId @returns {Object} Archive result. */
function dataArchiveHandoffNote(sessionToken, noteId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return HandoffService.archiveHandoffNote(noteId); }
/** @returns {void} Initializes handoff notes sheet. Steward-only. */
function dataInitHandoffNotes(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return HandoffService.initSheet(); }

// --- Mentorship (steward-only) ---
/** @param {string} sessionToken @returns {Array} Active mentorship pairings. */
function dataGetMentorshipPairings(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return MentorshipService.getPairings(); }
/** @param {string} sessionToken @param {string} mentorEmail @param {string} menteeEmail @param {string} [caseTypes] @returns {Object} Result. */
function dataCreateMentorshipPairing(sessionToken, mentorEmail, menteeEmail, caseTypes) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return MentorshipService.createPairing(mentorEmail, menteeEmail, caseTypes); }
/** @param {string} sessionToken @param {string} pairingId @param {string} notes @returns {Object} Result. */
function dataUpdateMentorshipNotes(sessionToken, pairingId, notes) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return MentorshipService.updatePairingNotes(pairingId, notes); }
/** @param {string} sessionToken @param {string} pairingId @returns {Object} Result. */
function dataCloseMentorshipPairing(sessionToken, pairingId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return MentorshipService.closePairing(pairingId); }
/** @param {string} sessionToken @returns {Array} Suggested mentor-mentee pairings. */
function dataGetMentorshipSuggestions(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return MentorshipService.suggestPairings(); }

// --- Leader Mentorship (steward-only) ---
/** @param {string} sessionToken @returns {Array} All Member Leaders with mentor info. */
function dataGetMemberLeaders(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return [];
  var leaders = (typeof loadMemberData_ === 'function') ? loadMemberData_().leaders : [];
  // Enrich with mentorship pairing info
  var pairings = (typeof MentorshipService !== 'undefined' && typeof MentorshipService.getPairings === 'function') ? MentorshipService.getPairings() : [];
  var mentorMap = {};
  for (var i = 0; i < pairings.length; i++) {
    if (pairings[i].menteeEmail) mentorMap[pairings[i].menteeEmail.toLowerCase()] = pairings[i];
  }
  return leaders.map(function(l) {
    var pairing = mentorMap[l.email.toLowerCase()] || null;
    return {
      name: l.name,
      email: l.email,
      unit: l.unit,
      location: l.location,
      mentorEmail: pairing ? pairing.mentorEmail : null,
      mentorPairingId: pairing ? pairing.id : null,
      mentorStarted: pairing ? pairing.started : null
    };
  });
}

// --- Communication Log (steward-only) ---
/** @param {string} sessionToken @param {string} memberEmail @param {string} type @param {string} subject @param {string} notes @returns {Object} Result. */
function dataLogCommunication(sessionToken, memberEmail, type, subject, notes) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return CommunicationLogService.logCommunication(s, memberEmail, type, subject, notes); }
/** @param {string} sessionToken @param {string} memberEmail @returns {Array} Communication history. */
function dataGetCommunicationLog(sessionToken, memberEmail) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return CommunicationLogService.getCommunicationLog(memberEmail); }

// --- Knowledge Base (any authenticated user) ---
/** @param {string} sessionToken @param {string} query @returns {Array} Matching articles. */
function dataSearchKnowledgeBase(sessionToken, query) { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return KnowledgeBaseService.searchArticles(query); }
/** @param {string} sessionToken @param {string} articleId @returns {Object|null} Full article. */
function dataGetKnowledgeBaseArticle(sessionToken, articleId) { var e = _resolveCallerEmail(sessionToken); if (!e) return null; return KnowledgeBaseService.getArticle(articleId); }
/** @param {string} sessionToken @param {string} grievanceType @returns {Array} Related articles. */
function dataGetRelatedArticles(sessionToken, grievanceType) { var e = _resolveCallerEmail(sessionToken); if (!e) return []; return KnowledgeBaseService.getRelatedArticles(grievanceType); }
/** @param {string} sessionToken @param {string} articleNumber @param {string} title @param {string} summary @param {string} fullText @param {string} relatedTypes @param {string} tags @returns {Object} Result (steward-only). */
function dataAddKnowledgeBaseArticle(sessionToken, articleNumber, title, summary, fullText, relatedTypes, tags) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return KnowledgeBaseService.addArticle(articleNumber, title, summary, fullText, relatedTypes, tags); }

// --- Digest Preferences (any authenticated user for own prefs) ---
/** @param {string} sessionToken @returns {Object} Notification preferences. */
function dataGetDigestPreferences(sessionToken) { var e = _resolveCallerEmail(sessionToken); if (!e) return { frequency: 'immediate', types: 'all' }; return DigestService.getPreferences(e); }
/** @param {string} sessionToken @param {string} frequency @param {string} types @returns {Object} Result. */
function dataSetDigestPreferences(sessionToken, frequency, types) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; return DigestService.setPreferences(e, frequency, types); }
/** Scheduled trigger — sends digest emails to all users with daily/weekly preference. */
function dataSendScheduledDigests() { return DigestService.sendScheduledDigests(); }

// --- Document Checklist (steward-only) ---
/** @param {string} sessionToken @param {string} caseId @param {string} currentStep @param {string} driveFolderId @returns {Array} Document checklist. */
function dataGetDocumentChecklist(sessionToken, caseId, currentStep, driveFolderId) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return DocumentChecklistService.getChecklist(caseId, currentStep, driveFolderId); }

// --- Escalation Recommendations (steward-only) ---
/** @param {string} sessionToken @param {string} caseId @returns {Object|null} Escalation recommendation. */
function dataGetEscalationRecommendation(sessionToken, caseId) { var s = _requireStewardAuth(sessionToken); if (!s) return null; return EscalationEngine.getRecommendation(caseId); }

// --- Report Generation (steward-only) ---
/** @param {string} sessionToken @param {Object} [dateRange] @returns {Object} Generated report data. */
function dataGenerateMonthlyReport(sessionToken, dateRange) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return ReportService.generateMonthlyReport(dateRange); }
/** @param {string} sessionToken @param {Object} report @returns {string} HTML report string (steward-only). */
function dataGenerateReportHtml(sessionToken, report) { var s = _requireStewardAuth(sessionToken); if (!s) return ''; return ReportService.generateReportHtml(report); }

// ============================================================================
// TWO-FACTOR AUTHENTICATION SERVICE (v4.36.0)
// ============================================================================

var TwoFactorService = (function () {

  var CODE_LENGTH = 6;
  var CODE_TTL_SECONDS = 300; // 5 minutes
  var RATE_LIMIT_MAX = 3;
  var RATE_LIMIT_WINDOW = 900; // 15 minutes

  /**
   * Generates a random 6-digit numeric code.
   * @returns {string} 6-digit code
   */
  function _generateCode() {
    var code = '';
    for (var i = 0; i < CODE_LENGTH; i++) {
      code += Math.floor(Math.random() * 10).toString();
    }
    return code;
  }

  /**
   * Hashes a code for secure storage (SHA-256).
   * @param {string} code - The plaintext code
   * @param {string} email - Used as salt
   * @returns {string} Hex hash
   */
  function _hashCode(code, email) {
    var raw = email.toLowerCase().trim() + ':' + code;
    var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    return hash.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
  }

  /**
   * Checks rate limit for code generation.
   * @param {string} email - User email
   * @returns {boolean} True if rate limit exceeded
   */
  function _isRateLimited(email) {
    var cache = CacheService.getScriptCache();
    var key = '2fa_rate_' + email.toLowerCase().trim();
    var count = parseInt(cache.get(key) || '0', 10);
    return count >= RATE_LIMIT_MAX;
  }

  /**
   * Increments rate limit counter.
   * @param {string} email - User email
   */
  function _incrementRateLimit(email) {
    var cache = CacheService.getScriptCache();
    var key = '2fa_rate_' + email.toLowerCase().trim();
    var count = parseInt(cache.get(key) || '0', 10);
    cache.put(key, String(count + 1), RATE_LIMIT_WINDOW);
  }

  /**
   * Generates and sends a 2FA verification code to the user's email.
   * @param {string} email - User's email address
   * @returns {Object} { success, message, error }
   */
  function generateCode(email) {
    if (!email) return { success: false, error: 'Email required.' };

    if (_isRateLimited(email)) {
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('2FA_RATE_LIMITED', SECURITY_SEVERITY.MEDIUM,
          '2FA code request rate limited',
          { email: typeof maskEmail === 'function' ? maskEmail(email) : email });
      }
      return { success: false, error: 'Too many verification requests. Please wait 15 minutes.' };
    }

    var code = _generateCode();
    var hash = _hashCode(code, email);

    // Store hashed code in CacheService with 5-minute TTL
    var cache = CacheService.getScriptCache();
    var cacheKey = '2fa_code_' + email.toLowerCase().trim();
    cache.put(cacheKey, hash, CODE_TTL_SECONDS);

    // Reset failure counter
    cache.remove('2fa_fail_' + email.toLowerCase().trim());

    _incrementRateLimit(email);

    // Send code via email
    var orgName = '';
    try {
      if (typeof getConfigValue_ === 'function') orgName = getConfigValue_(CONFIG_COLS.ORG_NAME) || '';
    } catch (_) { /* ignore */ }

    var subject = (orgName ? orgName + ' — ' : '') + 'Verification Code: ' + code;
    var htmlBody = '<div style="font-family:-apple-system,BlinkMacSystemFont,sans-serif;max-width:400px;margin:0 auto;padding:24px;text-align:center">'
      + '<h2 style="margin:0 0 16px;font-size:18px">Verification Code</h2>'
      + '<div style="font-size:32px;letter-spacing:8px;font-weight:700;color:#4f46e5;padding:16px;background:#f5f3ff;border-radius:12px;margin:0 0 16px">' + code + '</div>'
      + '<p style="font-size:14px;color:#666;margin:0 0 8px">Enter this code to confirm your action.</p>'
      + '<p style="font-size:12px;color:#999">This code expires in 5 minutes. If you didn\'t request this, ignore this email.</p>'
      + '</div>';

    try {
      if (typeof safeSendEmail_ === 'function') {
        var result = safeSendEmail_({ to: email, subject: subject, htmlBody: htmlBody });
        if (!result.success) return { success: false, error: result.error || 'Failed to send email.' };
      } else {
        MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
      }
    } catch (_sendErr) {
      return { success: false, error: 'Failed to send verification email.' };
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('2FA_CODE_SENT', { email: typeof maskEmail === 'function' ? maskEmail(email) : email });
    }

    return { success: true, message: 'Verification code sent to your email.' };
  }

  /**
   * Verifies a 2FA code.
   * @param {string} email - User's email address
   * @param {string} code - The code entered by the user
   * @returns {Object} { success, error }
   */
  function verifyCode(email, code) {
    if (!email || !code) return { success: false, error: 'Email and code are required.' };

    var cache = CacheService.getScriptCache();
    var cacheKey = '2fa_code_' + email.toLowerCase().trim();
    var failKey = '2fa_fail_' + email.toLowerCase().trim();

    // Check failure count
    var failures = parseInt(cache.get(failKey) || '0', 10);
    if (failures >= 5) {
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('2FA_LOCKOUT', SECURITY_SEVERITY.HIGH,
          '2FA verification locked out after 5 failures',
          { email: typeof maskEmail === 'function' ? maskEmail(email) : email });
      }
      return { success: false, error: 'Too many failed attempts. Request a new code.' };
    }

    var storedHash = cache.get(cacheKey);
    if (!storedHash) {
      return { success: false, error: 'Code expired or not requested. Please request a new code.' };
    }

    var inputHash = _hashCode(String(code).trim(), email);
    if (inputHash !== storedHash) {
      cache.put(failKey, String(failures + 1), CODE_TTL_SECONDS);
      if (failures + 1 >= 3 && typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('2FA_FAILED', SECURITY_SEVERITY.HIGH,
          '2FA verification failed ' + (failures + 1) + ' times',
          { email: typeof maskEmail === 'function' ? maskEmail(email) : email });
      }
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('2FA_VERIFY_FAILED', { email: typeof maskEmail === 'function' ? maskEmail(email) : email, attempts: failures + 1 });
      }
      return { success: false, error: 'Invalid code. ' + (4 - failures) + ' attempts remaining.' };
    }

    // Success — delete code to prevent replay
    cache.remove(cacheKey);
    cache.remove(failKey);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('2FA_VERIFIED', { email: typeof maskEmail === 'function' ? maskEmail(email) : email });
    }

    // Grant a short-lived 2FA session (10 minutes) so the user can complete their action
    var sessionKey = '2fa_session_' + email.toLowerCase().trim();
    cache.put(sessionKey, 'verified', 600);

    return { success: true };
  }

  /**
   * Checks if a user has a valid 2FA session (verified within last 10 minutes).
   * @param {string} email - User's email address
   * @returns {boolean} True if 2FA session is active
   */
  function hasValidSession(email) {
    if (!email) return false;
    var cache = CacheService.getScriptCache();
    return cache.get('2fa_session_' + email.toLowerCase().trim()) === 'verified';
  }

  return {
    generateCode: generateCode,
    verifyCode: verifyCode,
    hasValidSession: hasValidSession
  };
})();

// ============================================================================
// TWO-FACTOR GLOBAL WRAPPERS (v4.36.0)
// ============================================================================

/** @param {string} sessionToken @returns {Object} Result with message. */
function dataRequest2FACode(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, error: 'Not authenticated.' };
  return TwoFactorService.generateCode(e);
}

/** @param {string} sessionToken @param {string} code @returns {Object} Verification result. */
function dataVerify2FACode(sessionToken, code) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, error: 'Not authenticated.' };
  return TwoFactorService.verifyCode(e, code);
}

// ============================================================================
// SMS SERVICE — Twilio Integration (v4.36.0)
// ============================================================================

var SMSService = (function () {

  var HEADERS = ['ID', 'Recipient Email', 'Phone Number', 'Message', 'Status', 'Twilio SID', 'Sent At', 'Error'];

  // PropertiesService keys for Twilio credentials
  var PROP_ACCOUNT_SID  = 'TWILIO_ACCOUNT_SID';
  var PROP_AUTH_TOKEN   = 'TWILIO_AUTH_TOKEN';
  var PROP_FROM_NUMBER  = 'TWILIO_FROM_NUMBER';

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.SMS_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.SMS_LOG);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Configures Twilio credentials. Must be called once by admin.
   * @param {string} accountSid - Twilio Account SID
   * @param {string} authToken  - Twilio Auth Token
   * @param {string} fromNumber - Twilio phone number (E.164 format: +1XXXXXXXXXX)
   * @returns {Object} { success, message }
   */
  function configureProvider(accountSid, authToken, fromNumber) {
    if (!accountSid || !authToken || !fromNumber) {
      return { success: false, message: 'All three Twilio fields are required.' };
    }
    if (!/^\+\d{10,15}$/.test(fromNumber)) {
      return { success: false, message: 'From number must be E.164 format (e.g. +15551234567).' };
    }
    var props = PropertiesService.getScriptProperties();
    props.setProperties({
      TWILIO_ACCOUNT_SID: String(accountSid).trim(),
      TWILIO_AUTH_TOKEN:  String(authToken).trim(),
      TWILIO_FROM_NUMBER: String(fromNumber).trim()
    });
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('SMS_PROVIDER_CONFIGURED', { provider: 'Twilio', fromNumber: fromNumber });
    }
    return { success: true, message: 'Twilio configured successfully.' };
  }

  /**
   * Returns true if Twilio credentials are set.
   * @returns {boolean}
   */
  function isConfigured() {
    var props = PropertiesService.getScriptProperties();
    return !!(props.getProperty(PROP_ACCOUNT_SID) && props.getProperty(PROP_AUTH_TOKEN) && props.getProperty(PROP_FROM_NUMBER));
  }

  /**
   * Returns Twilio config (SID + from number, never the auth token).
   * @returns {Object} { configured, accountSid, fromNumber }
   */
  function getProviderStatus() {
    var props = PropertiesService.getScriptProperties();
    var sid = props.getProperty(PROP_ACCOUNT_SID) || '';
    return {
      configured: !!(sid && props.getProperty(PROP_AUTH_TOKEN) && props.getProperty(PROP_FROM_NUMBER)),
      accountSid: sid ? sid.substring(0, 8) + '...' : '',
      fromNumber: props.getProperty(PROP_FROM_NUMBER) || ''
    };
  }

  /**
   * Sends an SMS via Twilio.
   * @param {string} toPhone - Recipient phone in E.164 format
   * @param {string} message - SMS body (max 1600 chars)
   * @returns {Object} { success, sid, error }
   */
  function _sendViaTwilio(toPhone, message) {
    var props = PropertiesService.getScriptProperties();
    var sid   = props.getProperty(PROP_ACCOUNT_SID);
    var token = props.getProperty(PROP_AUTH_TOKEN);
    var from  = props.getProperty(PROP_FROM_NUMBER);
    if (!sid || !token || !from) {
      return { success: false, error: 'Twilio not configured.' };
    }

    var url = 'https://api.twilio.com/2010-04-01/Accounts/' + encodeURIComponent(sid) + '/Messages.json';
    var payload = {
      To:   toPhone,
      From: from,
      Body: String(message).substring(0, 1600)
    };

    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'post',
        payload: payload,
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token)
        },
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();
      var body = JSON.parse(response.getContentText());

      if (code === 201 || code === 200) {
        return { success: true, sid: body.sid || '' };
      }
      return { success: false, error: (body.message || 'Twilio error ' + code) };
    } catch (fetchErr) {
      return { success: false, error: fetchErr.message };
    }
  }

  /**
   * Looks up a member's phone number from the Member Directory.
   * @param {string} email - Member email
   * @returns {string|null} E.164 phone number or null
   */
  function _lookupPhone(email) {
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getDataRange().getValues();
    var emailLower = String(email).toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(col_(data[i], MEMBER_COLS.EMAIL) || '').toLowerCase().trim() === emailLower) {
        var phone = String(col_(data[i], MEMBER_COLS.PHONE) || '').trim();
        if (!phone) return null;
        // Normalize to E.164 if it looks like a US number
        var digits = phone.replace(/\D/g, '');
        if (digits.length === 10) return '+1' + digits;
        if (digits.length === 11 && digits.charAt(0) === '1') return '+' + digits;
        if (phone.charAt(0) === '+') return phone;
        return null; // Can't reliably format
      }
    }
    return null;
  }

  /**
   * Sends an SMS to a member by email address.
   * @param {string} recipientEmail - Member email (phone looked up from directory)
   * @param {string} message - SMS text
   * @returns {Object} { success, message, error }
   */
  function sendSMS(recipientEmail, message) {
    if (!recipientEmail || !message) {
      return { success: false, error: 'Recipient and message are required.' };
    }
    if (!isConfigured()) {
      return { success: false, error: 'SMS not configured. Set up Twilio in Admin Settings.' };
    }

    var phone = _lookupPhone(recipientEmail);
    if (!phone) {
      return { success: false, error: 'No valid phone number found for this member.' };
    }

    // Rate limit: max 5 SMS per recipient per day
    var cacheKey = 'sms_count_' + recipientEmail.toLowerCase().trim() + '_' + new Date().toISOString().substring(0, 10);
    var cache = CacheService.getScriptCache();
    var count = parseInt(cache.get(cacheKey) || '0', 10);
    if (count >= 5) {
      return { success: false, error: 'Daily SMS limit reached for this recipient.' };
    }

    var result = _sendViaTwilio(phone, message);

    // Log to SMS_LOG sheet
    try {
      var sheet = initSheet();
      sheet.appendRow([
        Utilities.getUuid().substring(0, 8),
        escapeForFormula(recipientEmail),
        escapeForFormula(phone),
        escapeForFormula(String(message).substring(0, 200)),
        result.success ? 'Sent' : 'Failed',
        result.sid || '',
        new Date(),
        result.error || ''
      ]);
    } catch (logErr) {
      log_('SMSService.sendSMS log error', logErr.message);
    }

    if (result.success) {
      cache.put(cacheKey, String(count + 1), 86400);
      if (typeof logAuditEvent === 'function') {
        logAuditEvent('SMS_SENT', { recipient: typeof maskEmail === 'function' ? maskEmail(recipientEmail) : recipientEmail });
      }
    }

    return {
      success: result.success,
      message: result.success ? 'SMS sent successfully.' : undefined,
      error: result.error
    };
  }

  /**
   * Sends a pre-formatted grievance status update SMS.
   * @param {string} recipientEmail - Member email
   * @param {string} grievanceId - Grievance ID
   * @param {string} newStatus - New status text
   * @returns {Object} { success, message, error }
   */
  function sendStatusUpdate(recipientEmail, grievanceId, newStatus) {
    var orgName = '';
    try {
      if (typeof getConfigValue_ === 'function') orgName = getConfigValue_(CONFIG_COLS.ORG_ABBREV) || getConfigValue_(CONFIG_COLS.ORG_NAME) || '';
    } catch (_) { /* ignore */ }
    var text = (orgName ? orgName + ': ' : '') + 'Your grievance ' + grievanceId + ' status has been updated to: ' + newStatus + '. Log in for details.';
    return sendSMS(recipientEmail, text);
  }

  /**
   * Sends a test SMS to verify Twilio configuration.
   * @param {string} toPhone - Phone number in E.164 format
   * @returns {Object} { success, sid, error }
   */
  function testSMS(toPhone) {
    if (!isConfigured()) {
      return { success: false, error: 'Twilio not configured.' };
    }
    var orgName = typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Dashboard') : 'Dashboard';
    return _sendViaTwilio(toPhone, 'Test message from ' + orgName + '. If you received this, SMS is working!');
  }

  /**
   * Returns recent SMS log entries.
   * @param {number} [limit=50] - Max entries to return
   * @returns {Array} SMS log entries (newest first)
   */
  function getLog(limit) {
    limit = limit || 50;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.SMS_LOG);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var results = [];
    for (var i = data.length - 1; i >= 0 && results.length < limit; i--) {
      results.push({
        id: data[i][0],
        recipient: data[i][1],
        phone: data[i][2],
        message: data[i][3],
        status: data[i][4],
        twilioSid: data[i][5],
        sentAt: data[i][6],
        error: data[i][7]
      });
    }
    return results;
  }

  return {
    initSheet: initSheet,
    configureProvider: configureProvider,
    isConfigured: isConfigured,
    getProviderStatus: getProviderStatus,
    sendSMS: sendSMS,
    sendStatusUpdate: sendStatusUpdate,
    testSMS: testSMS,
    getLog: getLog
  };
})();

// ============================================================================
// RSVP SERVICE — Meeting Invitation & Attendance Tracking (v4.36.0)
// ============================================================================

var RSVPService = (function () {

  var HEADERS = ['ID', 'Meeting ID', 'Member Email', 'Member Name', 'RSVP Status', 'RSVP Time', 'Token', 'Invited At', 'Attended'];

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.RSVP_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.RSVP_LOG);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Generates a unique RSVP token for a meeting + member pair.
   * @param {string} meetingId - Meeting ID
   * @param {string} email - Member email
   * @returns {string} Token (16-char hex)
   */
  function _generateToken(meetingId, email) {
    var token = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    var cache = CacheService.getScriptCache();
    var key = 'rsvp_' + token;
    cache.put(key, JSON.stringify({
      meetingId: meetingId,
      email: email.toLowerCase().trim(),
      created: Date.now()
    }), 86400); // 24h TTL
    return token;
  }

  /**
   * Validates an RSVP token. Returns token data if valid, null if expired/invalid.
   * @param {string} token - RSVP token
   * @returns {Object|null} { meetingId, email } or null
   */
  function _validateToken(token) {
    if (!token || token.length !== 16) return null;
    var cache = CacheService.getScriptCache();
    var key = 'rsvp_' + token;
    var raw = cache.get(key);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch (_) {
      return null;
    }
  }

  /**
   * Sends RSVP invitations for a meeting to all active members (or filtered set).
   * @param {string} meetingId - Meeting ID
   * @param {Object} [options] - { unit, smsEnabled }
   * @returns {Object} { success, invitedCount, smsCount, errors }
   */
  function sendInvitations(meetingId, options) {
    options = options || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, invitedCount: 0, errors: ['Spreadsheet unavailable.'] };

    // Look up the meeting
    var meetingSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!meetingSheet || meetingSheet.getLastRow() < 2) {
      return { success: false, invitedCount: 0, errors: ['Meeting not found.'] };
    }
    var meetingData = meetingSheet.getDataRange().getValues();
    var meeting = null;
    for (var mi = 1; mi < meetingData.length; mi++) {
      if (String(meetingData[mi][0]) === meetingId && !meetingData[mi][4]) {
        meeting = {
          id: meetingData[mi][0],
          name: meetingData[mi][1],
          date: meetingData[mi][2],
          type: meetingData[mi][3],
          time: meetingData[mi][8]
        };
        break;
      }
    }
    if (!meeting) return { success: false, invitedCount: 0, errors: ['Meeting ID not found: ' + meetingId] };

    // Get all active members
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet || memberSheet.getLastRow() < 2) {
      return { success: false, invitedCount: 0, errors: ['No members found.'] };
    }
    var members = memberSheet.getDataRange().getValues();
    var rsvpSheet = initSheet();
    var webAppUrl = '';
    try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (_) { /* ignore */ }

    var invitedCount = 0;
    var smsCount = 0;
    var errors = [];
    var smsEnabled = options.smsEnabled && SMSService.isConfigured();

    // Get existing invitations for this meeting to avoid duplicates
    var existingRsvps = {};
    if (rsvpSheet.getLastRow() >= 2) {
      var rsvpData = rsvpSheet.getRange(2, 1, rsvpSheet.getLastRow() - 1, HEADERS.length).getValues();
      for (var ri = 0; ri < rsvpData.length; ri++) {
        if (String(rsvpData[ri][1]) === meetingId) {
          existingRsvps[String(rsvpData[ri][2]).toLowerCase().trim()] = true;
        }
      }
    }

    // Org info for email
    var orgName = '';
    try {
      if (typeof getConfigValue_ === 'function') orgName = getConfigValue_(CONFIG_COLS.ORG_NAME) || '';
    } catch (_) { /* ignore */ }

    for (var i = 1; i < members.length; i++) {
      var email = String(col_(members[i], MEMBER_COLS.EMAIL) || '').trim();
      if (!email) continue;
      if (existingRsvps[email.toLowerCase()]) continue;

      // Optional unit filter
      if (options.unit) {
        var memberUnit = String(col_(members[i], MEMBER_COLS.UNIT) || '').trim();
        if (memberUnit.toLowerCase() !== String(options.unit).toLowerCase()) continue;
      }

      var firstName = String(col_(members[i], MEMBER_COLS.FIRST_NAME) || '').trim();
      var lastName = String(col_(members[i], MEMBER_COLS.LAST_NAME) || '').trim();
      var memberName = (firstName + ' ' + lastName).trim();

      // Generate RSVP token and links
      var token = _generateToken(meetingId, email);
      var acceptUrl = webAppUrl + '?page=rsvp&token=' + token + '&response=accept';
      var declineUrl = webAppUrl + '?page=rsvp&token=' + token + '&response=decline';

      // Format meeting date
      var meetingDateStr = '';
      try {
        var md = new Date(meeting.date);
        meetingDateStr = md.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
      } catch (_) {
        meetingDateStr = String(meeting.date);
      }

      // Send email invitation
      var subject = (orgName ? orgName + ' — ' : '') + 'Meeting Invitation: ' + meeting.name;
      var htmlBody = '<div style="font-family:-apple-system,BlinkMacSystemFont,sans-serif;max-width:500px;margin:0 auto;padding:24px">'
        + '<h2 style="margin:0 0 16px;font-size:20px">You\'re Invited</h2>'
        + '<p style="margin:0 0 8px;font-size:15px"><strong>' + escapeHtml(meeting.name) + '</strong></p>'
        + '<p style="margin:0 0 4px;font-size:14px;color:#555">' + escapeHtml(meetingDateStr) + (meeting.time ? ' at ' + escapeHtml(String(meeting.time)) : '') + '</p>'
        + '<p style="margin:0 0 4px;font-size:14px;color:#555">Type: ' + escapeHtml(meeting.type) + '</p>'
        + '<div style="margin:24px 0;text-align:center">'
        + '<a href="' + acceptUrl + '" style="display:inline-block;padding:12px 28px;background:#22c55e;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px;margin-right:12px">Accept</a>'
        + '<a href="' + declineUrl + '" style="display:inline-block;padding:12px 28px;background:#ef4444;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;font-size:14px">Decline</a>'
        + '</div>'
        + '<p style="font-size:12px;color:#999">This link expires in 24 hours.</p>'
        + '</div>';

      try {
        if (typeof safeSendEmail_ === 'function') {
          var emailResult = safeSendEmail_({ to: email, subject: subject, htmlBody: htmlBody });
          if (!emailResult.success) {
            errors.push(email + ': ' + (emailResult.error || 'email failed'));
            continue;
          }
        } else {
          MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
        }
      } catch (emailErr) {
        errors.push(email + ': ' + emailErr.message);
        continue;
      }

      // Optionally send SMS
      if (smsEnabled) {
        var smsResult = SMSService.sendSMS(email, 'Meeting: ' + meeting.name + ' on ' + meetingDateStr + '. RSVP at: ' + acceptUrl);
        if (smsResult.success) smsCount++;
      }

      // Write RSVP row
      rsvpSheet.appendRow([
        Utilities.getUuid().substring(0, 8),
        escapeForFormula(meetingId),
        escapeForFormula(email),
        escapeForFormula(memberName),
        RSVP_STATUS.INVITED,
        '',
        token,
        new Date(),
        ''
      ]);
      invitedCount++;
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('RSVP_INVITATIONS_SENT', { meetingId: meetingId, invitedCount: invitedCount, smsCount: smsCount });
    }

    return { success: true, invitedCount: invitedCount, smsCount: smsCount, errors: errors };
  }

  /**
   * Processes an RSVP response from a token-authenticated deep link.
   * @param {string} token - RSVP token
   * @param {string} response - 'accept' or 'decline'
   * @returns {Object} { success, meetingName, memberName, rsvpStatus, error }
   */
  function processRSVP(token, response) {
    var tokenData = _validateToken(token);
    if (!tokenData) {
      if (typeof recordSecurityEvent === 'function') {
        recordSecurityEvent('RSVP_INVALID_TOKEN', SECURITY_SEVERITY.LOW, 'Invalid or expired RSVP token', { token: token ? token.substring(0, 4) + '...' : 'null' });
      }
      return { success: false, error: 'This RSVP link has expired or is invalid.' };
    }

    var newStatus = (response === 'accept') ? RSVP_STATUS.ACCEPTED : RSVP_STATUS.DECLINED;

    // Update the RSVP row
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'System unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.RSVP_LOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'RSVP record not found.' };

    var data = sheet.getDataRange().getValues();
    var meetingName = '';
    var memberName = '';
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]) === token) {
        sheet.getRange(i + 1, 5).setValue(newStatus); // RSVP Status
        sheet.getRange(i + 1, 6).setValue(new Date()); // RSVP Time
        meetingName = data[i][1]; // Meeting ID — we'll look up the name
        memberName = data[i][3];
        break;
      }
    }

    // Look up meeting name
    if (meetingName) {
      var meetingSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
      if (meetingSheet && meetingSheet.getLastRow() >= 2) {
        var mData = meetingSheet.getDataRange().getValues();
        for (var mi = 1; mi < mData.length; mi++) {
          if (String(mData[mi][0]) === String(tokenData.meetingId)) {
            meetingName = mData[mi][1];
            break;
          }
        }
      }
    }

    // Delete token to prevent reuse
    CacheService.getScriptCache().remove('rsvp_' + token);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('RSVP_RESPONSE', {
        meetingId: tokenData.meetingId,
        email: typeof maskEmail === 'function' ? maskEmail(tokenData.email) : tokenData.email,
        status: newStatus
      });
    }

    return {
      success: true,
      meetingName: meetingName || tokenData.meetingId,
      memberName: memberName,
      rsvpStatus: newStatus
    };
  }

  /**
   * Returns RSVP summary for a meeting.
   * @param {string} meetingId - Meeting ID
   * @returns {Object} { meetingId, invited, accepted, declined, noResponse, attended, members }
   */
  function getRSVPSummary(meetingId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { meetingId: meetingId, invited: 0, accepted: 0, declined: 0, noResponse: 0, attended: 0, members: [] };
    var sheet = ss.getSheetByName(SHEETS.RSVP_LOG);
    if (!sheet || sheet.getLastRow() < 2) return { meetingId: meetingId, invited: 0, accepted: 0, declined: 0, noResponse: 0, attended: 0, members: [] };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var counts = { invited: 0, accepted: 0, declined: 0, noResponse: 0, attended: 0 };
    var members = [];

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]) !== meetingId) continue;
      var status = String(data[i][4]).trim();
      counts.invited++;
      if (status === RSVP_STATUS.ACCEPTED) counts.accepted++;
      else if (status === RSVP_STATUS.DECLINED) counts.declined++;
      else if (status === RSVP_STATUS.ATTENDED) { counts.attended++; counts.accepted++; }
      else counts.noResponse++;

      members.push({
        email: data[i][2],
        name: data[i][3],
        status: status,
        rsvpTime: data[i][5],
        attended: data[i][8] === true || data[i][8] === 'Yes'
      });
    }

    return {
      meetingId: meetingId,
      invited: counts.invited,
      accepted: counts.accepted,
      declined: counts.declined,
      noResponse: counts.noResponse,
      attended: counts.attended,
      members: members
    };
  }

  /**
   * Reconciles RSVP records with actual meeting check-in data.
   * Updates RSVP status to 'Attended' for members who checked in.
   * @param {string} meetingId - Meeting ID
   * @returns {Object} { success, reconciled, expectedVsActual }
   */
  function reconcileAttendance(meetingId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, reconciled: 0 };

    // Get actual check-ins from meeting log
    var checkInSheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!checkInSheet || checkInSheet.getLastRow() < 2) return { success: false, reconciled: 0 };
    var checkInData = checkInSheet.getDataRange().getValues();
    var attendees = {};
    for (var ci = 1; ci < checkInData.length; ci++) {
      if (String(checkInData[ci][0]) === meetingId && checkInData[ci][4]) {
        attendees[String(checkInData[ci][7]).toLowerCase().trim()] = true;
      }
    }

    // Update RSVP records
    var rsvpSheet = ss.getSheetByName(SHEETS.RSVP_LOG);
    if (!rsvpSheet || rsvpSheet.getLastRow() < 2) return { success: false, reconciled: 0 };
    var rsvpData = rsvpSheet.getDataRange().getValues();
    var reconciled = 0;

    for (var i = 1; i < rsvpData.length; i++) {
      if (String(rsvpData[i][1]) !== meetingId) continue;
      var email = String(rsvpData[i][2]).toLowerCase().trim();
      if (attendees[email]) {
        rsvpSheet.getRange(i + 1, 5).setValue(RSVP_STATUS.ATTENDED);
        rsvpSheet.getRange(i + 1, 9).setValue('Yes');
        reconciled++;
      }
    }

    var summary = getRSVPSummary(meetingId);
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('RSVP_RECONCILED', { meetingId: meetingId, reconciled: reconciled, expected: summary.accepted, actual: reconciled });
    }

    return {
      success: true,
      reconciled: reconciled,
      expectedVsActual: { expected: summary.accepted, actual: reconciled }
    };
  }

  return {
    initSheet: initSheet,
    sendInvitations: sendInvitations,
    processRSVP: processRSVP,
    getRSVPSummary: getRSVPSummary,
    reconcileAttendance: reconcileAttendance
  };
})();

// ============================================================================
// SMS & RSVP GLOBAL WRAPPERS (v4.36.0)
// ============================================================================

// --- SMS (steward-only) ---
/** @param {string} sessionToken @param {string} accountSid @param {string} authToken @param {string} fromNumber @returns {Object} Result. */
function dataConfigureTwilio(sessionToken, accountSid, authToken, fromNumber) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return SMSService.configureProvider(accountSid, authToken, fromNumber); }
/** @param {string} sessionToken @param {string} recipientEmail @param {string} message @returns {Object} Result. */
function dataSendSMS(sessionToken, recipientEmail, message) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return SMSService.sendSMS(recipientEmail, message); }
/** @param {string} sessionToken @param {number} [limit] @returns {Array} SMS log. */
function dataGetSMSLog(sessionToken, limit) { var s = _requireStewardAuth(sessionToken); if (!s) return []; return SMSService.getLog(limit); }
/** @param {string} sessionToken @param {string} toPhone @returns {Object} Test result. */
function dataTestSMS(sessionToken, toPhone) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return SMSService.testSMS(toPhone); }
/** @param {string} sessionToken @returns {Object} Provider status (never exposes auth token). */
function dataGetSMSStatus(sessionToken) { var s = _requireStewardAuth(sessionToken); if (!s) return { configured: false }; return SMSService.getProviderStatus(); }

// --- RSVP (steward for management, token-auth for responding) ---
/** @param {string} sessionToken @param {string} meetingId @param {Object} [options] @returns {Object} Invitation result. */
function dataSendMeetingInvitations(sessionToken, meetingId, options) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return RSVPService.sendInvitations(meetingId, options); }
/** @param {string} sessionToken @param {string} meetingId @returns {Object} RSVP summary. */
function dataGetRSVPSummary(sessionToken, meetingId) { var s = _requireStewardAuth(sessionToken); if (!s) return { meetingId: meetingId, invited: 0, accepted: 0, declined: 0, noResponse: 0, attended: 0, members: [] }; return RSVPService.getRSVPSummary(meetingId); }
/** @param {string} sessionToken @param {string} meetingId @returns {Object} Reconciliation result. */
function dataReconcileAttendance(sessionToken, meetingId) { var s = _requireStewardAuth(sessionToken); if (!s) return { success: false, message: 'Steward access required.' }; return RSVPService.reconcileAttendance(meetingId); }
/** @param {string} token @param {string} response @returns {Object} RSVP result. Token-authenticated (no session needed). */
function dataProcessRSVP(token, response) { return RSVPService.processRSVP(token, response); }

// --- Sheet Initialization (no auth — setup only) ---
/** Initializes all new feature sheets. Call once from Apps Script editor or setup flow. */
function initNewFeatureSheets() {
  HandoffService.initSheet();
  MentorshipService.initSheet();
  CommunicationLogService.initSheet();
  KnowledgeBaseService.initSheet();
  DigestService.initSheet();
  SMSService.initSheet();
  RSVPService.initSheet();
}
