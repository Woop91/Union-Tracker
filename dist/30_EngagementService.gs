/**
 * ============================================================================
 * 30_EngagementService.gs — Engagement Score Service
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Computes per-member engagement scores from 5 dimensions: survey
 *   participation, meeting attendance, Q&A activity, workload submissions,
 *   and steward contact freshness. Provides a steward-facing scoreboard
 *   and a member-facing report card.
 *
 * SCORING (0-100 per dimension, weighted composite):
 *   Survey Participation  — 25%  (surveys completed / total available)
 *   Meeting Attendance    — 25%  (meetings attended / total, 90-day window)
 *   Q&A Activity          — 15%  (min(100, posts*10 + answers*15))
 *   Workload Submissions  — 15%  (submissions in last 90 days, scaled)
 *   Contact Freshness     — 20%  (days since last steward contact)
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Helps stewards identify disengaged members before they drop off.
 *   Scores are informational, not punitive — no inter-member ranking
 *   is visible to members. Pre-computed batch runs reduce real-time load.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Engagement tab and Report Card show empty/error state. Core features
 *   (grievances, auth, meetings) are unaffected.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, MEMBER_COLS),
 *   00_Security.gs (escapeForFormula, maskEmail)
 *
 * @version 4.51.0
 */

var EngagementService = (function () {

  var HEADERS = ['Member Email', 'Member Name', 'Unit', 'Survey Score', 'Meeting Score',
    'QA Score', 'Workload Score', 'Contact Score', 'Composite Score', 'Last Computed', 'Trend'];

  var WEIGHTS = { survey: 25, meeting: 25, qa: 15, workload: 15, contact: 20 };

  function initSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken.');
    var sheet = ss.getSheetByName(SHEETS.ENGAGEMENT_SCORES);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.ENGAGEMENT_SCORES);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  /**
   * Helper: read sheet data safely.
   */
  function _sheetRows(ss, sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  }

  /**
   * Computes survey participation score for a member.
   * @returns {number} 0-100
   */
  function _computeSurveyScore(email, surveyData) {
    if (!surveyData || surveyData.length === 0) return 50; // default if no survey system
    var emailLower = email.toLowerCase().trim();
    var total = 0;
    var completed = 0;
    for (var i = 0; i < surveyData.length; i++) {
      var rowEmail = String(surveyData[i][2] || '').toLowerCase().trim(); // EMAIL col per _Survey_Assignments schema (col C, index 2)
      if (rowEmail === emailLower) {
        total++;
        var status = String(surveyData[i][5] || '').toLowerCase().trim(); // CURRENT_STATUS per _Survey_Assignments schema (col F, index 5)
        if (status === 'completed') completed++;
      }
    }
    if (total === 0) return 50;
    return Math.round((completed / total) * 100);
  }

  /**
   * Computes meeting attendance score for a member (90-day window).
   * @returns {number} 0-100
   */
  function _computeMeetingScore(email, checkInData, totalMeetings90d) {
    if (totalMeetings90d === 0) return 50;
    var emailLower = email.toLowerCase().trim();
    var attended = 0;
    var now = new Date();
    var day90ago = new Date(now.getTime() - 90 * 86400000);

    for (var i = 0; i < checkInData.length; i++) {
      var ciEmail = String(checkInData[i][7] || '').toLowerCase().trim(); // EMAIL col per _Meeting_Check_In schema (col H, index 7)
      if (ciEmail !== emailLower) continue;
      var ciDate = checkInData[i][2]; // MEETING_DATE per _Meeting_Check_In schema (col C, index 2)
      if (ciDate instanceof Date && ciDate >= day90ago) attended++;
    }
    return Math.min(100, Math.round((attended / Math.max(totalMeetings90d, 1)) * 100));
  }

  /**
   * Computes Q&A activity score.
   * @returns {number} 0-100
   */
  function _computeQAScore(email, qaData, answerData) {
    var emailLower = email.toLowerCase().trim();
    var posts = 0;
    var answers = 0;

    if (qaData) {
      for (var i = 0; i < qaData.length; i++) {
        if (String(qaData[i][1] || '').toLowerCase().trim() === emailLower) posts++; // Author Email per _QA_Forum schema (col B, index 1)
      }
    }
    if (answerData) {
      for (var j = 0; j < answerData.length; j++) {
        if (String(answerData[j][2] || '').toLowerCase().trim() === emailLower) answers++; // Author Email per _QA_Answers schema (col C, index 2)
      }
    }
    return Math.min(100, posts * 10 + answers * 15);
  }

  /**
   * Computes workload submission score (90-day window).
   * @returns {number} 0-100
   */
  function _computeWorkloadScore(email, workloadData) {
    if (!workloadData || workloadData.length === 0) return 50;
    var emailLower = email.toLowerCase().trim();
    var submissions = 0;
    var now = new Date();
    var day90ago = new Date(now.getTime() - 90 * 86400000);

    for (var i = 0; i < workloadData.length; i++) {
      var wEmail = String(workloadData[i][1] || '').toLowerCase().trim(); // EMAIL per _Workload_Submissions schema (col B, index 1)
      if (wEmail !== emailLower) continue;
      var wDate = workloadData[i][0]; // TIMESTAMP per _Workload_Submissions schema (col A, index 0)
      if (wDate instanceof Date && wDate >= day90ago) submissions++;
    }
    // ~13 weeks in 90 days, 1 per week is excellent
    return Math.min(100, Math.round((submissions / 13) * 100));
  }

  /**
   * Computes contact freshness score based on days since last steward contact.
   * @returns {number} 0-100
   */
  function _computeContactScore(email, contactData) {
    if (!contactData || contactData.length === 0) return 0;
    var emailLower = email.toLowerCase().trim();
    var latestContact = null;

    for (var i = 0; i < contactData.length; i++) {
      var cEmail = String(contactData[i][2] || '').toLowerCase().trim(); // MEMBER_EMAIL per _Contact_Log schema (col C, index 2)
      if (cEmail !== emailLower) continue;
      var cDate = contactData[i][0]; // TIMESTAMP per _Contact_Log schema (col A, index 0)
      if (cDate instanceof Date && (!latestContact || cDate > latestContact)) {
        latestContact = cDate;
      }
    }

    if (!latestContact) return 0;
    var daysSince = Math.floor((new Date() - latestContact) / 86400000);
    if (daysSince < 14) return 100;
    if (daysSince < 30) return 75;
    if (daysSince < 60) return 50;
    if (daysSince < 90) return 25;
    return 0;
  }

  /**
   * Computes engagement score for a single member.
   * @param {string} email
   * @param {Object} cachedData - Pre-loaded sheet data for batch efficiency
   * @returns {Object} { email, scores: {survey, meeting, qa, workload, contact}, composite }
   */
  function computeScoreForMember(email, cachedData) {
    cachedData = cachedData || {};
    var scores = {
      survey: _computeSurveyScore(email, cachedData.surveyData),
      meeting: _computeMeetingScore(email, cachedData.checkInData, cachedData.totalMeetings90d || 0),
      qa: _computeQAScore(email, cachedData.qaData, cachedData.answerData),
      workload: _computeWorkloadScore(email, cachedData.workloadData),
      contact: _computeContactScore(email, cachedData.contactData)
    };

    var composite = Math.round(
      (scores.survey * WEIGHTS.survey +
       scores.meeting * WEIGHTS.meeting +
       scores.qa * WEIGHTS.qa +
       scores.workload * WEIGHTS.workload +
       scores.contact * WEIGHTS.contact) / 100
    );

    return { email: email, scores: scores, composite: composite };
  }

  /**
   * Batch computes scores for all active members and writes to _Engagement_Scores.
   * @returns {Object} { computed, avgScore }
   */
  function computeAllScores() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { computed: 0, avgScore: 0 };

    // Load all member data
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet || memberSheet.getLastRow() < 2) return { computed: 0, avgScore: 0 };
    var memberData = memberSheet.getDataRange().getValues();

    // Pre-load all data sources once (batch efficiency)
    var cachedData = {
      surveyData: _sheetRows(ss, SHEETS.SURVEY_TRACKING || '_Survey_Tracking'),
      checkInData: _sheetRows(ss, SHEETS.MEETING_CHECKIN_LOG),
      qaData: _sheetRows(ss, SHEETS.QA_FORUM),
      answerData: _sheetRows(ss, SHEETS.QA_ANSWERS),
      workloadData: _sheetRows(ss, SHEETS.WORKLOAD_VAULT),
      contactData: _sheetRows(ss, SHEETS.CONTACT_LOG)
    };

    // Count unique meetings in last 90 days
    var now = new Date();
    var day90ago = new Date(now.getTime() - 90 * 86400000);
    var meetingIds = {};
    for (var ci = 0; ci < cachedData.checkInData.length; ci++) {
      var mDate = cachedData.checkInData[ci][2];
      if (mDate instanceof Date && mDate >= day90ago) {
        meetingIds[String(cachedData.checkInData[ci][0])] = true;
      }
    }
    cachedData.totalMeetings90d = Object.keys(meetingIds).length;

    // Read existing scores for trend comparison
    var existingScores = {};
    var scoreSheet = initSheet();
    if (scoreSheet.getLastRow() >= 2) {
      var oldData = scoreSheet.getRange(2, 1, scoreSheet.getLastRow() - 1, HEADERS.length).getValues();
      for (var oi = 0; oi < oldData.length; oi++) {
        existingScores[String(oldData[oi][0]).toLowerCase().trim()] = oldData[oi][8] || 0; // Composite
      }
    }

    // Compute scores for each active member
    var results = [];
    var totalScore = 0;
    for (var i = 1; i < memberData.length; i++) {
      var email = String(col_(memberData[i], MEMBER_COLS.EMAIL) || '').trim();
      if (!email) continue;
      var dues = String(col_(memberData[i], MEMBER_COLS.DUES_STATUS) || '').trim().toLowerCase();
      if (dues === 'inactive') continue;

      var firstName = String(col_(memberData[i], MEMBER_COLS.FIRST_NAME) || '').trim();
      var lastName = String(col_(memberData[i], MEMBER_COLS.LAST_NAME) || '').trim();
      var unit = String(col_(memberData[i], MEMBER_COLS.UNIT) || '').trim();

      var result = computeScoreForMember(email, cachedData);
      var oldScore = existingScores[email.toLowerCase().trim()] || 0;
      var trend = result.composite > oldScore ? 'up' : (result.composite < oldScore ? 'down' : 'stable');

      results.push([
        escapeForFormula(email),
        escapeForFormula((firstName + ' ' + lastName).trim()),
        escapeForFormula(unit),
        result.scores.survey,
        result.scores.meeting,
        result.scores.qa,
        result.scores.workload,
        result.scores.contact,
        result.composite,
        new Date(),
        trend
      ]);
      totalScore += result.composite;
    }

    // Write new data first, then clear any remaining old rows
    if (results.length > 0) {
      scoreSheet.getRange(2, 1, results.length, HEADERS.length).setValues(results);
    }
    // Clear remaining old rows beyond new data
    var lastRow = scoreSheet.getLastRow();
    if (lastRow > results.length + 1) {
      scoreSheet.getRange(results.length + 2, 1, lastRow - results.length - 1, HEADERS.length).clearContent();
    }

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('ENGAGEMENT_SCORES_COMPUTED', { members: results.length, avgScore: results.length > 0 ? Math.round(totalScore / results.length) : 0 });
    }

    return {
      computed: results.length,
      avgScore: results.length > 0 ? Math.round(totalScore / results.length) : 0
    };
  }

  /**
   * Returns the engagement scoreboard (paginated, sortable, filterable).
   * @param {Object} [options] - { page, pageSize, sortBy, filterUnit, searchTerm }
   * @returns {Object} { items, totalRows, page, pageSize, avgScore, topUnit, lowCount }
   */
  function getScoreboard(options) {
    options = options || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { items: [], totalRows: 0, page: 0, pageSize: 20 };
    var sheet = ss.getSheetByName(SHEETS.ENGAGEMENT_SCORES);
    if (!sheet || sheet.getLastRow() < 2) return { items: [], totalRows: 0, page: 0, pageSize: 20 };

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var page = Math.max(0, parseInt(options.page, 10) || 0);
    var pageSize = Math.min(100, Math.max(10, parseInt(options.pageSize, 10) || 20));
    var sortBy = options.sortBy || 'composite';
    var filterUnit = options.filterUnit ? String(options.filterUnit).toLowerCase().trim() : '';
    var searchTerm = options.searchTerm ? String(options.searchTerm).toLowerCase().trim() : '';

    // Build items
    var items = [];
    var totalScore = 0;
    var unitScores = {};
    var lowCount = 0;

    for (var i = 0; i < data.length; i++) {
      var email = String(data[i][0]).trim();
      var name = String(data[i][1]).trim();
      var unit = String(data[i][2]).trim();
      var composite = parseInt(data[i][8], 10) || 0;

      totalScore += composite;
      if (composite < 30) lowCount++;

      if (!unitScores[unit]) unitScores[unit] = { total: 0, count: 0 };
      unitScores[unit].total += composite;
      unitScores[unit].count++;

      if (filterUnit && unit.toLowerCase() !== filterUnit) continue;
      if (searchTerm && (name + ' ' + email + ' ' + unit).toLowerCase().indexOf(searchTerm) === -1) continue;

      items.push({
        email: email,
        name: name,
        unit: unit,
        survey: parseInt(data[i][3], 10) || 0,
        meeting: parseInt(data[i][4], 10) || 0,
        qa: parseInt(data[i][5], 10) || 0,
        workload: parseInt(data[i][6], 10) || 0,
        contact: parseInt(data[i][7], 10) || 0,
        composite: composite,
        trend: String(data[i][10] || 'stable')
      });
    }

    // Sort
    var sortCol = { composite: 'composite', survey: 'survey', meeting: 'meeting', qa: 'qa', workload: 'workload', contact: 'contact', name: 'name' }[sortBy] || 'composite';
    items.sort(function(a, b) {
      if (sortCol === 'name') return a.name.localeCompare(b.name);
      return (b[sortCol] || 0) - (a[sortCol] || 0);
    });

    // Find top unit
    var topUnit = '';
    var topUnitAvg = 0;
    var unitNames = Object.keys(unitScores);
    for (var ui = 0; ui < unitNames.length; ui++) {
      var avg = unitScores[unitNames[ui]].total / unitScores[unitNames[ui]].count;
      if (avg > topUnitAvg) { topUnitAvg = avg; topUnit = unitNames[ui]; }
    }

    var totalRows = items.length;
    var start = page * pageSize;
    var pageItems = items.slice(start, start + pageSize);

    return {
      items: pageItems,
      totalRows: totalRows,
      page: page,
      pageSize: pageSize,
      avgScore: data.length > 0 ? Math.round(totalScore / data.length) : 0,
      topUnit: topUnit,
      lowCount: lowCount
    };
  }

  /**
   * Returns engagement scores aggregated by unit.
   * @returns {Array} [{ unit, avgScore, memberCount }]
   */
  function getScoreByUnit() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName(SHEETS.ENGAGEMENT_SCORES);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var units = {};
    for (var i = 0; i < data.length; i++) {
      var unit = String(data[i][2] || 'Unknown').trim();
      var composite = parseInt(data[i][8], 10) || 0;
      if (!units[unit]) units[unit] = { total: 0, count: 0 };
      units[unit].total += composite;
      units[unit].count++;
    }

    var result = [];
    var unitNames = Object.keys(units).sort();
    for (var u = 0; u < unitNames.length; u++) {
      result.push({
        unit: unitNames[u],
        avgScore: Math.round(units[unitNames[u]].total / units[unitNames[u]].count),
        memberCount: units[unitNames[u]].count
      });
    }
    return result;
  }

  /**
   * Returns a single member's score breakdown.
   * @param {string} email
   * @returns {Object|null} Score object or null
   */
  function getMemberScore(email) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.ENGAGEMENT_SCORES);
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    var emailLower = String(email).toLowerCase().trim();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === emailLower) {
        return {
          email: data[i][0],
          name: data[i][1],
          unit: data[i][2],
          scores: {
            survey: parseInt(data[i][3], 10) || 0,
            meeting: parseInt(data[i][4], 10) || 0,
            qa: parseInt(data[i][5], 10) || 0,
            workload: parseInt(data[i][6], 10) || 0,
            contact: parseInt(data[i][7], 10) || 0
          },
          composite: parseInt(data[i][8], 10) || 0,
          lastComputed: data[i][9],
          trend: String(data[i][10] || 'stable')
        };
      }
    }
    return null;
  }

  /**
   * Returns a member's full report card with activity details.
   * Members can only see their own data.
   * @param {string} email
   * @returns {Object} Report card data
   */
  function getMemberReportCard(email) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { scores: null };

    var score = getMemberScore(email);
    if (!score) {
      // Compute on-the-fly if no pre-computed score exists
      var cachedData = {
        surveyData: _sheetRows(ss, SHEETS.SURVEY_TRACKING || '_Survey_Tracking'),
        checkInData: _sheetRows(ss, SHEETS.MEETING_CHECKIN_LOG),
        qaData: _sheetRows(ss, SHEETS.QA_FORUM),
        answerData: _sheetRows(ss, SHEETS.QA_ANSWERS),
        workloadData: _sheetRows(ss, SHEETS.WORKLOAD_VAULT),
        contactData: _sheetRows(ss, SHEETS.CONTACT_LOG),
        totalMeetings90d: 0
      };
      var computed = computeScoreForMember(email, cachedData);
      score = { email: email, name: '', unit: '', scores: computed.scores, composite: computed.composite, trend: 'stable' };
    }

    var emailLower = email.toLowerCase().trim();

    // Survey completion dates
    var surveyHistory = [];
    try {
      var surveyRows = _sheetRows(ss, SHEETS.SURVEY_TRACKING || '_Survey_Tracking');
      for (var si = 0; si < surveyRows.length; si++) {
        if (String(surveyRows[si][2] || '').toLowerCase().trim() === emailLower) {
          var status = String(surveyRows[si][5] || '').toLowerCase().trim();
          if (status === 'completed' && surveyRows[si][6] instanceof Date) {
            surveyHistory.push(surveyRows[si][6].toISOString().substring(0, 10));
          }
        }
      }
    } catch (_) { /* ignore */ }

    // Meeting attendance count
    var meetingCount = { last90: 0, total: 0 };
    try {
      var checkInRows = _sheetRows(ss, SHEETS.MEETING_CHECKIN_LOG);
      var now = new Date();
      var day90ago = new Date(now.getTime() - 90 * 86400000);
      for (var ci = 0; ci < checkInRows.length; ci++) {
        if (String(checkInRows[ci][7] || '').toLowerCase().trim() === emailLower && checkInRows[ci][4]) {
          meetingCount.total++;
          if (checkInRows[ci][2] instanceof Date && checkInRows[ci][2] >= day90ago) {
            meetingCount.last90++;
          }
        }
      }
    } catch (_) { /* ignore */ }

    // Q&A post count
    var qaPostCount = 0;
    try {
      var qaRows = _sheetRows(ss, SHEETS.QA_FORUM);
      for (var qi = 0; qi < qaRows.length; qi++) {
        if (String(qaRows[qi][1] || '').toLowerCase().trim() === emailLower) qaPostCount++;
      }
    } catch (_) { /* ignore */ }

    // Last steward contact date
    var lastContactDate = null;
    try {
      var contactRows = _sheetRows(ss, SHEETS.CONTACT_LOG);
      for (var cti = 0; cti < contactRows.length; cti++) {
        if (String(contactRows[cti][2] || '').toLowerCase().trim() === emailLower) {
          var cDate = contactRows[cti][0];
          if (cDate instanceof Date && (!lastContactDate || cDate > lastContactDate)) {
            lastContactDate = cDate;
          }
        }
      }
    } catch (_) { /* ignore */ }

    // Grievance summary
    var grievanceSummary = { open: 0, resolved: 0, won: 0 };
    try {
      var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      if (gSheet && gSheet.getLastRow() >= 2) {
        var gData = gSheet.getDataRange().getValues();
        var gHeaders = gData[0];
        var gEmailCol = -1, gStatusCol = -1;
        for (var gc = 0; gc < gHeaders.length; gc++) {
          var gh = String(gHeaders[gc]).toLowerCase().trim();
          if (gh === 'member email') gEmailCol = gc;
          if (gh === 'status') gStatusCol = gc;
        }
        if (gEmailCol >= 0 && gStatusCol >= 0) {
          for (var gi = 1; gi < gData.length; gi++) {
            if (String(gData[gi][gEmailCol] || '').toLowerCase().trim() !== emailLower) continue;
            var gStatus = String(gData[gi][gStatusCol] || '').toLowerCase().trim();
            if (gStatus === 'settled') { grievanceSummary.resolved++; grievanceSummary.won++; }
            else if (gStatus === 'resolved' || gStatus === 'withdrawn' || gStatus === 'denied') grievanceSummary.resolved++;
            else grievanceSummary.open++;
          }
        }
      }
    } catch (_) { /* ignore */ }

    return {
      scores: score,
      surveyHistory: surveyHistory,
      meetingCount: meetingCount,
      qaPostCount: qaPostCount,
      lastContactDate: lastContactDate ? lastContactDate.toISOString() : null,
      grievanceSummary: grievanceSummary
    };
  }

  return {
    initSheet: initSheet,
    computeScoreForMember: computeScoreForMember,
    computeAllScores: computeAllScores,
    getScoreboard: getScoreboard,
    getScoreByUnit: getScoreByUnit,
    getMemberScore: getMemberScore,
    getMemberReportCard: getMemberReportCard
  };
})();

// ============================================================================
// ENGAGEMENT SERVICE GLOBAL WRAPPERS (v4.36.0)
// ============================================================================

/** @param {string} sessionToken @param {Object} [options] @returns {Object} Scoreboard data. */
function dataGetEngagementScoreboard(sessionToken, options) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { items: [], totalRows: 0 };
  return EngagementService.getScoreboard(options);
}

/** @param {string} sessionToken @returns {Object} Computation summary. */
function dataComputeEngagementScores(sessionToken) {
  var s = _requireStewardAuth(sessionToken);
  if (!s) return { success: false, message: 'Steward access required.' };
  var result = EngagementService.computeAllScores();
  return { success: true, computed: result.computed, avgScore: result.avgScore };
}

/** @param {string} sessionToken @returns {Object} Member's own report card. */
function dataGetMyReportCard(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, authError: true, message: 'Not authenticated.' };
  return EngagementService.getMemberReportCard(e);
}


