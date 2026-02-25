/**
 * 24_WeeklyQuestions.gs — Weekly Questions Engagement System
 *
 * Two questions per week: one set by a steward, one drawn from the member pool.
 * Responses are anonymous (email hashed). Results show live after responding.
 *
 * Sheets (hidden):
 *   _Weekly_Questions  — active and past questions
 *   _Weekly_Responses  — hashed-email responses
 *   _Question_Pool     — member-submitted question candidates
 *
 * @version 4.11.0
 * @requires 01_Core.gs (SHEETS)
 * @requires 06_Maintenance.gs (logAuditEvent)
 */

var WeeklyQuestions = (function() {

  // ═══════════════════════════════════════
  // SHEET SETUP
  // ═══════════════════════════════════════

  function initWeeklyQuestionSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    _ensureSheet(ss, SHEETS.WEEKLY_QUESTIONS, [
      'ID', 'Text', 'Source', 'Submitted By', 'Week Start', 'Active', 'Created'
    ]);
    _ensureSheet(ss, SHEETS.WEEKLY_RESPONSES, [
      'ID', 'Question ID', 'Email Hash', 'Response', 'Timestamp'
    ]);
    _ensureSheet(ss, SHEETS.QUESTION_POOL, [
      'ID', 'Text', 'Submitted By Hash', 'Status', 'Created'
    ]);
  }

  function _ensureSheet(ss, name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.hideSheet();
    }
    return sheet;
  }

  // ═══════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════

  function _hashEmail(email) {
    var raw = String(email).trim().toLowerCase();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    return digest.map(function(b) {
      return ('0' + (b & 0xFF).toString(16)).slice(-2);
    }).join('');
  }

  function _getWeekStart(date) {
    date = date || new Date();
    var d = new Date(date);
    d.setHours(0, 0, 0, 0);
    var day = d.getDay();
    var diff = d.getDate() - day + (day === 0 ? -6 : 1); // Monday
    d.setDate(diff);
    return d;
  }

  function _generateId() {
    return 'WQ_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 4);
  }

  function _getSheet(name) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      // Lazy-init: auto-create missing sheets
      initWeeklyQuestionSheets();
      sheet = ss.getSheetByName(name);
    }
    return sheet;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Get Active Questions
  // ═══════════════════════════════════════

  /**
   * Returns this week's active questions with user response status + live stats.
   * @param {string} email
   * @returns {Object} { questions: [{id, text, source, hasResponded, stats}] }
   */
  function getActiveQuestions(email) {
    var weekStart = _getWeekStart();
    var weekStr = weekStart.toISOString().split('T')[0];
    var emailHash = _hashEmail(email);

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() <= 1) return { questions: [] };

    var qData = qSheet.getDataRange().getValues();
    var activeQs = [];

    for (var i = 1; i < qData.length; i++) {
      var qWeek = qData[i][4]; // Week Start column
      if (qWeek instanceof Date) qWeek = qWeek.toISOString().split('T')[0];
      var active = String(qData[i][5]).toLowerCase();
      if (String(qWeek) === weekStr && (active === 'true' || active === 'yes' || active === '1')) {
        activeQs.push({
          id: qData[i][0],
          text: qData[i][1],
          source: qData[i][2],
        });
      }
    }

    // Check responses and get stats
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    var responses = rSheet && rSheet.getLastRow() > 1 ? rSheet.getDataRange().getValues() : [];

    activeQs.forEach(function(q) {
      q.hasResponded = false;
      var counts = {};
      var total = 0;
      for (var r = 1; r < responses.length; r++) {
        if (responses[r][1] !== q.id) continue;
        total++;
        var resp = responses[r][3];
        counts[resp] = (counts[resp] || 0) + 1;
        if (responses[r][2] === emailHash) q.hasResponded = true;
      }
      q.stats = { total: total, counts: counts };
    });

    return { questions: activeQs };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Submit Response
  // ═══════════════════════════════════════

  /**
   * Submits a response to a weekly question. Returns updated stats.
   * @param {string} email
   * @param {string} questionId
   * @param {string} response
   * @returns {Object} { success, stats }
   */
  function submitResponse(email, questionId, response) {
    if (!email || !questionId || !response) {
      return { success: false, message: 'Missing required fields.' };
    }

    var emailHash = _hashEmail(email);

    // Check for duplicate
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    if (!rSheet) return { success: false, message: 'System not initialized.' };

    if (rSheet.getLastRow() > 1) {
      var existing = rSheet.getDataRange().getValues();
      for (var i = 1; i < existing.length; i++) {
        if (existing[i][1] === questionId && existing[i][2] === emailHash) {
          return { success: false, message: 'Already responded.' };
        }
      }
    }

    // Write response
    var id = _generateId();
    rSheet.appendRow([id, questionId, emailHash, response, new Date()]);

    // Calculate updated stats
    var allResp = rSheet.getDataRange().getValues();
    var counts = {};
    var total = 0;
    for (var r = 1; r < allResp.length; r++) {
      if (allResp[r][1] !== questionId) continue;
      total++;
      var resp = allResp[r][3];
      counts[resp] = (counts[resp] || 0) + 1;
    }

    return { success: true, stats: { total: total, counts: counts } };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Submit Pool Question
  // ═══════════════════════════════════════

  /**
   * Submits a candidate question to the pool.
   * @param {string} email
   * @param {string} questionText
   * @returns {Object} { success, message }
   */
  function submitPoolQuestion(email, questionText) {
    if (!email || !questionText) return { success: false, message: 'Missing fields.' };

    var sheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!sheet) return { success: false, message: 'System not initialized.' };

    var id = _generateId();
    var emailHash = _hashEmail(email);
    sheet.appendRow([id, questionText.trim().substring(0, 500), emailHash, 'pending', new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WQ_POOL_SUBMIT', { emailHash: emailHash.substring(0, 8) });
    }

    return { success: true, message: 'Question submitted.' };
  }

  // ═══════════════════════════════════════
  // PUBLIC: Question Stats
  // ═══════════════════════════════════════

  function getQuestionStats(questionId) {
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    if (!rSheet || rSheet.getLastRow() <= 1) return { total: 0, counts: {} };

    var data = rSheet.getDataRange().getValues();
    var counts = {};
    var total = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] !== questionId) continue;
      total++;
      var resp = data[i][3];
      counts[resp] = (counts[resp] || 0) + 1;
    }
    return { total: total, counts: counts };
  }

  function hasUserResponded(email, questionId) {
    var emailHash = _hashEmail(email);
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    if (!rSheet || rSheet.getLastRow() <= 1) return false;

    var data = rSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === questionId && data[i][2] === emailHash) return true;
    }
    return false;
  }

  // ═══════════════════════════════════════
  // PUBLIC: Steward — Set Question & Pool Management
  // ═══════════════════════════════════════

  /**
   * Steward sets their weekly question.
   * @param {string} stewardEmail
   * @param {string} text
   * @returns {Object} { success, message }
   */
  function setStewardQuestion(stewardEmail, text) {
    if (!stewardEmail || !text) return { success: false, message: 'Missing fields.' };

    var sheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!sheet) return { success: false, message: 'System not initialized.' };

    var weekStart = _getWeekStart();
    var weekStr = weekStart.toISOString().split('T')[0];
    var id = _generateId();

    sheet.appendRow([id, text.trim().substring(0, 500), 'steward', stewardEmail, weekStr, 'TRUE', new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('WQ_STEWARD_SET', { steward: stewardEmail, weekStart: weekStr });
    }

    return { success: true, message: 'Question set for this week.' };
  }

  /**
   * Returns pending pool questions for steward review.
   * @returns {Object[]} Array of { id, text, status, created }
   */
  function getPoolQuestions() {
    var sheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getDataRange().getValues();
    var questions = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase() === 'pending') {
        questions.push({
          id: data[i][0],
          text: data[i][1],
          status: data[i][3],
          created: data[i][4],
        });
      }
    }
    return questions;
  }

  /**
   * Selects a random question from the pending pool and makes it active.
   * Intended for a weekly time-driven trigger.
   * @returns {Object} { success, questionId }
   */
  function selectRandomPoolQuestion() {
    var poolSheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!poolSheet || poolSheet.getLastRow() <= 1) return { success: false, message: 'No pool questions.' };

    var data = poolSheet.getDataRange().getValues();
    var pending = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase() === 'pending') {
        pending.push({ row: i + 1, id: data[i][0], text: data[i][1] });
      }
    }

    if (pending.length === 0) return { success: false, message: 'No pending questions.' };

    var selected = pending[Math.floor(Math.random() * pending.length)];

    // Mark as used in pool
    poolSheet.getRange(selected.row, 4).setValue('used');

    // Add to active questions
    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet) return { success: false, message: 'Questions sheet missing.' };

    var weekStart = _getWeekStart();
    var weekStr = weekStart.toISOString().split('T')[0];
    var newId = _generateId();

    qSheet.appendRow([newId, selected.text, 'pool', '', weekStr, 'TRUE', new Date()]);

    return { success: true, questionId: newId, text: selected.text };
  }

  // ═══════════════════════════════════════
  // PUBLIC: History (v4.12.0)
  // ═══════════════════════════════════════

  function getHistory(email, page, pageSize) {
    page = page || 1;
    pageSize = pageSize || 10;
    var emailHash = _hashEmail(email);

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() <= 1) return { questions: [], hasMore: false };

    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    var responses = rSheet && rSheet.getLastRow() > 1 ? rSheet.getDataRange().getValues() : [];

    var qData = qSheet.getDataRange().getValues();
    var allQs = [];
    for (var i = 1; i < qData.length; i++) {
      allQs.push({ id: qData[i][0], text: qData[i][1], source: qData[i][2], weekStart: qData[i][4] });
    }
    allQs.sort(function(a, b) {
      var da = a.weekStart instanceof Date ? a.weekStart.getTime() : 0;
      var db = b.weekStart instanceof Date ? b.weekStart.getTime() : 0;
      return db - da;
    });

    var start = (page - 1) * pageSize;
    var paged = allQs.slice(start, start + pageSize);

    paged.forEach(function(q) {
      var counts = {};
      var total = 0;
      for (var r = 1; r < responses.length; r++) {
        if (responses[r][1] !== q.id) continue;
        total++;
        var resp = responses[r][3];
        counts[resp] = (counts[resp] || 0) + 1;
      }
      q.stats = { total: total, counts: counts };
      q.weekStr = q.weekStart instanceof Date ? q.weekStart.toLocaleDateString() : String(q.weekStart || '');
    });

    return { questions: paged, hasMore: start + pageSize < allQs.length };
  }

  // Public API
  return {
    initWeeklyQuestionSheets: initWeeklyQuestionSheets,
    getActiveQuestions: getActiveQuestions,
    submitResponse: submitResponse,
    submitPoolQuestion: submitPoolQuestion,
    getQuestionStats: getQuestionStats,
    hasUserResponded: hasUserResponded,
    setStewardQuestion: setStewardQuestion,
    getPoolQuestions: getPoolQuestions,
    selectRandomPoolQuestion: selectRandomPoolQuestion,
    getHistory: getHistory,
  };

})();


// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client)
// ═══════════════════════════════════════

function wqGetActiveQuestions(email) { return WeeklyQuestions.getActiveQuestions(email); }
function wqSubmitResponse(email, questionId, response) { return WeeklyQuestions.submitResponse(email, questionId, response); }
function wqSubmitPoolQuestion(email, text) { return WeeklyQuestions.submitPoolQuestion(email, text); }
function wqSetStewardQuestion(stewardEmail, text) { return WeeklyQuestions.setStewardQuestion(stewardEmail, text); }
function wqGetPoolQuestions() { return WeeklyQuestions.getPoolQuestions(); }
function wqInitSheets() { return WeeklyQuestions.initWeeklyQuestionSheets(); }
function wqGetHistory(email, page, pageSize) { return WeeklyQuestions.getHistory(email, page, pageSize); }
