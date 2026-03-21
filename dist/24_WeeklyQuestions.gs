/**
 * 24_WeeklyQuestions.gs — Unified Polls System (v4.23.0)
 *
 * WHAT THIS FILE DOES:
 *   Unified weekly polls system (v4.23.0, replaced separate "Weekly Questions"
 *   + "Polls"). Two polls active each week:
 *     (1) Steward Poll — manually created by any steward, resets weekly
 *     (2) Community Poll — randomly drawn from member-submitted pool every
 *         Monday
 *   Fully anonymous: only SHA-256 hashed email stored, never plaintext.
 *   Results visible to all AFTER voting (members) or always (stewards). No
 *   "myVote" field ever returned.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   SHA-256 email hashing ensures anonymity while preventing double-voting
 *   (same email always produces same hash). Question pool validation enforces
 *   non-leading question text, 2-5 distinct options, and single-concept
 *   questions. Three hidden sheets:
 *     _Weekly_Questions  — active and past polls
 *       Cols: ID | Text | Options (JSON) | Source | SubmittedBy | WeekStart | Active | Created
 *     _Weekly_Responses  — hashed-email responses
 *       Cols: ID | QuestionID | EmailHash | Response | Timestamp
 *     _Question_Pool     — member-submitted poll candidates
 *       Cols: ID | Text | Options (JSON) | SubmittedByHash | Status | Created
 *   0-indexed column constants (Q_COLS, R_COLS, P_COLS) match the portal
 *   convention.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   No polls are shown in the SPA. Voting fails. If hashing breaks, votes are
 *   stored with plaintext emails (privacy violation). If the weekly reset
 *   fails, stale polls persist. The question pool stops accepting member
 *   submissions.
 *
 * DEPENDENCIES:
 *   Depends on: 01_Core.gs (SHEETS, escapeForFormula),
 *               Utilities.computeDigest (GAS SHA-256).
 *   Used by: SPA poll views and weekly trigger for auto-rotation.
 */

var WeeklyQuestions = (function () {

  // ── Column indices (0-based) ─────────────────────────────────────────────
  var Q_COLS = { ID: 0, TEXT: 1, OPTIONS: 2, SOURCE: 3, SUBMITTED_BY: 4, WEEK_START: 5, ACTIVE: 6, CREATED: 7 };
  var R_COLS = { ID: 0, QUESTION_ID: 1, EMAIL_HASH: 2, RESPONSE: 3, TIMESTAMP: 4 };
  var P_COLS = { ID: 0, TEXT: 1, OPTIONS: 2, SUBMITTED_BY_HASH: 3, STATUS: 4, CREATED: 5 };

  // ── Sheet bootstrap ──────────────────────────────────────────────────────

  function initWeeklyQuestionSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet unavailable.');
    _ensureSheet(ss, SHEETS.WEEKLY_QUESTIONS,  ['ID','Text','Options','Source','Submitted By','Week Start','Active','Created']);
    _ensureSheet(ss, SHEETS.WEEKLY_RESPONSES,  ['ID','Question ID','Email Hash','Response','Timestamp']);
    _ensureSheet(ss, SHEETS.QUESTION_POOL,     ['ID','Text','Options','Submitted By Hash','Status','Created']);
  }

  function _ensureSheet(ss, name, headers) {
    var sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.hideSheet();
    return sheet;
  }

  // ── Helpers ──────────────────────────────────────────────────────────────

  function _hashEmail(email) {
    var raw = String(email).trim().toLowerCase();
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    return digest.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
  }

  /**
   * Returns the poll frequency setting stored in ScriptProperties.
   * Defaults to 'weekly' if not set.
   * @returns {'weekly'|'biweekly'|'monthly'}
   */
  function _getFrequency() {
    try {
      var val = PropertiesService.getScriptProperties().getProperty('POLL_FREQUENCY');
      if (val === 'biweekly' || val === 'monthly') return val;
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    return 'weekly';
  }

  /**
   * Returns the canonical start date of the current poll period.
   * weekly   -> Monday of the current week
   * biweekly -> Monday of the even-numbered ISO week (weeks 2,4,6... start new period)
   * monthly  -> 1st of the current month
   */
  function _getPeriodStart(date) {
    var d = new Date(date || new Date());
    d.setHours(0, 0, 0, 0);
    var freq = _getFrequency();

    if (freq === 'monthly') {
      return new Date(d.getFullYear(), d.getMonth(), 1);
    }

    // Get Monday of this week
    var day = d.getDay();
    var monday = new Date(d);
    monday.setDate(d.getDate() - (day === 0 ? 6 : day - 1));

    if (freq === 'biweekly') {
      // ISO week number
      var jan4 = new Date(monday.getFullYear(), 0, 4);
      var weekNum = Math.ceil(((monday - jan4) / 86400000 + jan4.getDay() + 1) / 7);
      // If odd week, step back one week to land on even-week anchor
      if (weekNum % 2 !== 0) {
        monday.setDate(monday.getDate() - 7);
      }
    }

    return monday;
  }

  function _periodKey(date) {
    return _getPeriodStart(date).toISOString().split('T')[0];
  }

  function _id() {
    return 'PL_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 4);
  }

  function _parseOptions(raw) {
    if (!raw) return [];
    try { return JSON.parse(raw); } catch (_e) {
      // Legacy comma-separated fallback
      return String(raw).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
  }

  function _validateOptions(options) {
    if (!Array.isArray(options) || options.length < 2 || options.length > 5) {
      return 'Polls require 2–5 answer options.';
    }
    var seen = {};
    for (var i = 0; i < options.length; i++) {
      var o = String(options[i]).trim();
      if (!o) return 'Options cannot be blank.';
      if (o.length > 100) return 'Each option must be 100 characters or fewer.';
      if (seen[o.toLowerCase()]) return 'Options must be unique.';
      seen[o.toLowerCase()] = true;
    }
    return null; // valid
  }

  function _getSheet(name) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(name);
    if (!sheet) { initWeeklyQuestionSheets(); sheet = ss.getSheetByName(name); }
    return sheet;
  }

  // ── Public: Get Active Polls ─────────────────────────────────────────────

  /**
   * Returns this week's active polls with anonymous aggregate stats.
   * hasResponded = true if the caller already voted (prevents double-vote).
   * options = [] always returned (needed before voting).
   * counts/total = aggregate only; no individual mapping.
   */
  function getActiveQuestions(email) {
    var thisPeriod = _periodKey();
    var emailHash = _hashEmail(email);

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() <= 1) return { questions: [] };

    var qData = qSheet.getDataRange().getValues();
    var active = [];
    for (var i = 1; i < qData.length; i++) {
      var wk = qData[i][Q_COLS.WEEK_START];
      if (wk instanceof Date) wk = wk.toISOString().split('T')[0];
      var isActive = String(qData[i][Q_COLS.ACTIVE]).toLowerCase();
      if (String(wk) === thisPeriod && (isActive === 'true' || isActive === 'yes' || isActive === '1')) {
        active.push({
          id:      String(qData[i][Q_COLS.ID]),
          text:    String(qData[i][Q_COLS.TEXT]),
          options: _parseOptions(qData[i][Q_COLS.OPTIONS]),
          source:  String(qData[i][Q_COLS.SOURCE] || 'steward'),
        });
      }
    }

    // Build response stats in one pass
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    var responses = rSheet && rSheet.getLastRow() > 1 ? rSheet.getDataRange().getValues() : [];

    active.forEach(function(q) {
      var counts = {};
      var total = 0;
      q.options.forEach(function(o) { counts[o] = 0; });
      q.hasResponded = false;
      for (var r = 1; r < responses.length; r++) {
        if (String(responses[r][R_COLS.QUESTION_ID]) !== q.id) continue;
        total++;
        var resp = String(responses[r][R_COLS.RESPONSE]);
        counts[resp] = (counts[resp] || 0) + 1;
        if (responses[r][R_COLS.EMAIL_HASH] === emailHash) q.hasResponded = true;
      }
      q.stats = { total: total, counts: counts };
    });

    return { questions: active };
  }

  // ── Public: Submit Response ──────────────────────────────────────────────

  /**
   * Records an anonymous vote. Returns updated aggregate stats only.
   * Never returns which option the caller chose.
   */
  function submitResponse(email, questionId, response) {
    if (!email || !questionId || !response) return { success: false, message: 'Missing fields.' };

    var emailHash = _hashEmail(email);
    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    if (!rSheet) return { success: false, message: 'System not initialized.' };

    // Dedup check
    if (rSheet.getLastRow() > 1) {
      var existing = rSheet.getDataRange().getValues();
      for (var i = 1; i < existing.length; i++) {
        if (String(existing[i][R_COLS.QUESTION_ID]) === questionId &&
            existing[i][R_COLS.EMAIL_HASH] === emailHash) {
          return { success: false, message: 'Already voted.' };
        }
      }
    }

    rSheet.appendRow([_id(), questionId, emailHash, String(response), new Date()]);

    // Return updated aggregate stats
    var all = rSheet.getDataRange().getValues();
    var counts = {};
    var total = 0;
    for (var r = 1; r < all.length; r++) {
      if (String(all[r][R_COLS.QUESTION_ID]) !== questionId) continue;
      total++;
      var resp = String(all[r][R_COLS.RESPONSE]);
      counts[resp] = (counts[resp] || 0) + 1;
    }
    return { success: true, stats: { total: total, counts: counts } };
  }

  // ── Public: Steward — Create Poll ────────────────────────────────────────

  /**
   * Steward creates this week's poll. Replaces any existing steward poll for
   * the current week (one steward poll per week maximum).
   * @param {string} stewardEmail - verified server-side by wrapper
   * @param {string} text         - question text
   * @param {string[]} options    - 2–5 answer strings
   */
  function setStewardQuestion(stewardEmail, text, options) {
    if (!stewardEmail || !text) return { success: false, message: 'Missing fields.' };

    text = text.trim().substring(0, 500);
    if (!text) return { success: false, message: 'Question text required.' };

    var optErr = _validateOptions(options);
    if (optErr) return { success: false, message: optErr };

    // Role check — belt-and-suspenders (wrapper already calls _requireStewardAuth)
    var callerEmail = '';
    try { callerEmail = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
    if (!callerEmail) return { success: false, message: 'Unable to verify identity.' };

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet) return { success: false, message: 'System not initialized.' };

    var thisPeriod = _periodKey();

    // Deactivate any existing steward poll this week
    if (qSheet.getLastRow() > 1) {
      var data = qSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var wk = data[i][Q_COLS.WEEK_START];
        if (wk instanceof Date) wk = wk.toISOString().split('T')[0];
        if (String(wk) === thisPeriod && String(data[i][Q_COLS.SOURCE]) === 'steward') {
          qSheet.getRange(i + 1, Q_COLS.ACTIVE + 1).setValue('FALSE');
        }
      }
    }

    var id = _id();
    qSheet.appendRow([id, escapeForFormula(text), JSON.stringify(options), 'steward',
                      callerEmail, thisPeriod, 'TRUE', new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_STEWARD_CREATE', { steward: callerEmail, weekStart: thisPeriod });
    }
    return { success: true, message: 'Poll created for this week.', id: id };
  }

  // ── Public: Member — Submit to Pool ─────────────────────────────────────

  /**
   * Member submits a poll question candidate to the pool.
   * Submitter email is hashed — no PII stored.
   * @param {string} email
   * @param {string} text    - question text
   * @param {string[]} options - 2–5 answer strings
   */
  var POOL_SUBMIT_LIMIT = 3; // max submissions per user per poll period

  function submitPoolQuestion(email, text, options) {
    if (!email || !text) return { success: false, message: 'Missing fields.' };

    text = text.trim().substring(0, 500);
    if (!text) return { success: false, message: 'Question text required.' };

    var optErr = _validateOptions(options);
    if (optErr) return { success: false, message: optErr };

    var sheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!sheet) return { success: false, message: 'System not initialized.' };

    var emailHash = _hashEmail(email);

    // Rate limit: max POOL_SUBMIT_LIMIT submissions per user per period
    if (sheet.getLastRow() > 1) {
      var poolData = sheet.getDataRange().getValues();
      var periodStart = _getPeriodStart();
      var userCount = 0;
      for (var i = 1; i < poolData.length; i++) {
        if (poolData[i][P_COLS.SUBMITTED_BY_HASH] === emailHash) {
          var created = poolData[i][P_COLS.CREATED];
          if (created instanceof Date && created >= periodStart) userCount++;
        }
      }
      if (userCount >= POOL_SUBMIT_LIMIT) {
        return { success: false, message: 'You can submit up to ' + POOL_SUBMIT_LIMIT + ' polls per period. Please wait for the next period.' };
      }
    }

    var id = _id();
    sheet.appendRow([id, escapeForFormula(text), JSON.stringify(options), emailHash, 'pending', new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_POOL_SUBMIT', { emailHash: emailHash.substring(0, 8) });
    }
    return { success: true, message: 'Your poll has been added to the community pool. It may be drawn in a future week.' };
  }

  // ── Public: Random Pool Draw (time trigger) ──────────────────────────────

  /**
   * Selects ONE random pending question from the pool and activates it for
   * the current week as the community poll track.
   * Called by a weekly Monday time trigger — NOT by stewards.
   * Replaces any existing community poll for the current week.
   * @returns {Object} { success, questionId, text }
   */
  function selectRandomPoolQuestion() {
    var poolSheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!poolSheet || poolSheet.getLastRow() <= 1) {
      Logger.log('selectRandomPoolQuestion: pool is empty.');
      return { success: false, message: 'No pool questions.' };
    }

    var poolData = poolSheet.getDataRange().getValues();
    var pending = [];
    for (var i = 1; i < poolData.length; i++) {
      if (String(poolData[i][P_COLS.STATUS]).toLowerCase() === 'pending') {
        pending.push({ row: i + 1, id: poolData[i][P_COLS.ID], text: poolData[i][P_COLS.TEXT], options: poolData[i][P_COLS.OPTIONS] });
      }
    }
    if (pending.length === 0) return { success: false, message: 'No pending questions.' };

    var selected = pending[Math.floor(Math.random() * pending.length)];
    poolSheet.getRange(selected.row, P_COLS.STATUS + 1).setValue('used');

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet) return { success: false, message: 'Questions sheet missing.' };

    var thisPeriod = _periodKey();

    // Deactivate existing community poll for this week
    if (qSheet.getLastRow() > 1) {
      var existing = qSheet.getDataRange().getValues();
      for (var j = 1; j < existing.length; j++) {
        var wk = existing[j][Q_COLS.WEEK_START];
        if (wk instanceof Date) wk = wk.toISOString().split('T')[0];
        if (String(wk) === thisPeriod && String(existing[j][Q_COLS.SOURCE]) === 'community') {
          qSheet.getRange(j + 1, Q_COLS.ACTIVE + 1).setValue('FALSE');
        }
      }
    }

    var newId = _id();
    qSheet.appendRow([newId, escapeForFormula(String(selected.text)), selected.options, 'community',
                      '', thisPeriod, 'TRUE', new Date()]);

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_COMMUNITY_DRAW', { weekStart: thisPeriod, poolId: selected.id });
    }
    return { success: true, questionId: newId, text: selected.text };
  }

  // ── Public: Steward — Close a Poll ──────────────────────────────────────

  /**
   * Deactivates a poll by ID. Steward-only.
   */
  function closePoll(stewardEmail, pollId) {
    if (!stewardEmail || !pollId) return { success: false, message: 'Missing fields.' };
    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() <= 1) return { success: false, message: 'No polls found.' };
    var data = qSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][Q_COLS.ID]) === pollId) {
        qSheet.getRange(i + 1, Q_COLS.ACTIVE + 1).setValue('FALSE');
        return { success: true, message: 'Poll closed.' };
      }
    }
    return { success: false, message: 'Poll not found.' };
  }

  // ── Public: History (paginated) ──────────────────────────────────────────

  function getHistory(email, page, pageSize) {
    page = Math.max(1, page || 1);
    pageSize = Math.min(20, pageSize || 10);

    var qSheet = _getSheet(SHEETS.WEEKLY_QUESTIONS);
    if (!qSheet || qSheet.getLastRow() <= 1) return { questions: [], hasMore: false };

    var rSheet = _getSheet(SHEETS.WEEKLY_RESPONSES);
    var responses = rSheet && rSheet.getLastRow() > 1 ? rSheet.getDataRange().getValues() : [];

    var qData = qSheet.getDataRange().getValues();
    var allQs = [];
    for (var i = 1; i < qData.length; i++) {
      allQs.push({
        id:        String(qData[i][Q_COLS.ID]),
        text:      String(qData[i][Q_COLS.TEXT]),
        options:   _parseOptions(qData[i][Q_COLS.OPTIONS]),
        source:    String(qData[i][Q_COLS.SOURCE] || 'steward'),
        weekStart: qData[i][Q_COLS.WEEK_START],
      });
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
      q.options.forEach(function(o) { counts[o] = 0; });
      for (var r = 1; r < responses.length; r++) {
        if (String(responses[r][R_COLS.QUESTION_ID]) !== q.id) continue;
        total++;
        var resp = String(responses[r][R_COLS.RESPONSE]);
        counts[resp] = (counts[resp] || 0) + 1;
      }
      q.stats = { total: total, counts: counts };
      q.weekStr = q.weekStart instanceof Date
        ? q.weekStart.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })
        : String(q.weekStart || '');
    });

    return { questions: paged, hasMore: start + pageSize < allQs.length };
  }

  // ── Public: Pool questions (count only, for display) ────────────────────

  function getPoolCount() {
    var sheet = _getSheet(SHEETS.QUESTION_POOL);
    if (!sheet || sheet.getLastRow() <= 1) return 0;
    var data = sheet.getDataRange().getValues();
    var n = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][P_COLS.STATUS]).toLowerCase() === 'pending') n++;
    }
    return n;
  }

  // ── Public: Frequency settings ──────────────────────────────────────────

  /**
   * Returns current poll frequency: 'weekly' | 'biweekly' | 'monthly'.
   * Safe to call from any context (member or steward).
   */
  function getPollFrequency() {
    return _getFrequency();
  }

  /**
   * Sets the poll frequency. Steward-only (enforced by wrapper).
   * Updates the Monday time trigger interval accordingly.
   * @param {string} freq - 'weekly' | 'biweekly' | 'monthly'
   */
  function setPollFrequency(stewardEmail, freq) {
    if (freq !== 'weekly' && freq !== 'biweekly' && freq !== 'monthly') {
      return { success: false, message: 'Invalid frequency. Use weekly, biweekly, or monthly.' };
    }
    try {
      PropertiesService.getScriptProperties().setProperty('POLL_FREQUENCY', freq);
    } catch (e) {
      return { success: false, message: 'Failed to save setting: ' + e.message };
    }
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_FREQ_CHANGE', { steward: stewardEmail, frequency: freq });
    }
    return { success: true, frequency: freq, message: 'Poll frequency set to ' + freq + '.' };
  }

  // Public API
  return {
    initWeeklyQuestionSheets: initWeeklyQuestionSheets,
    getActiveQuestions:       getActiveQuestions,
    submitResponse:           submitResponse,
    setStewardQuestion:       setStewardQuestion,
    submitPoolQuestion:       submitPoolQuestion,
    selectRandomPoolQuestion: selectRandomPoolQuestion,
    closePoll:                closePoll,
    getHistory:               getHistory,
    getPoolCount:             getPoolCount,
    getPollFrequency:         getPollFrequency,
    setPollFrequency:         setPollFrequency,
    Q_COLS:                   Q_COLS,  // v4.24.4 — exposed so autoSelectCommunityPoll avoids duplicating indices
  };

})();
// ═══════════════════════════════════════
// GLOBAL WRAPPERS
// ═══════════════════════════════════════

function wqGetActiveQuestions(sessionToken) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { questions: [] };
  // v4.28.1: server-side dues gate — non-paying members cannot view active polls
  if (typeof DataService !== 'undefined') {
    var rec = DataService.findUserByEmail(e);
    if (rec && rec.duesPaying === false) return { questions: [] };
  }
  return WeeklyQuestions.getActiveQuestions(e);
}

function wqSubmitResponse(sessionToken, questionId, response) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  // v4.28.1: server-side dues gate — non-paying members cannot vote
  if (typeof DataService !== 'undefined') {
    var rec = DataService.findUserByEmail(e);
    if (rec && rec.duesPaying === false) return { success: false, message: 'Polls require active dues membership.' };
  }
  return WeeklyQuestions.submitResponse(e, questionId, response);
}

// v4.23.0: options param added (array of 2–5 strings)
function wqSetStewardQuestion(sessionToken, text, options) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Steward access required.' };
  return WeeklyQuestions.setStewardQuestion(e, text, options);
}

function wqSubmitPoolQuestion(sessionToken, text, options) {
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  // v4.28.1: server-side dues gate — non-paying members cannot submit to pool
  if (typeof DataService !== 'undefined') {
    var rec = DataService.findUserByEmail(e);
    if (rec && rec.duesPaying === false) return { success: false, message: 'Polls require active dues membership.' };
  }
  return WeeklyQuestions.submitPoolQuestion(e, text, options);
}

function wqClosePoll(sessionToken, pollId) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Steward access required.' };
  return WeeklyQuestions.closePoll(e, pollId);
}

function wqGetHistory(sessionToken, page, pageSize) {
  var e = _resolveCallerEmail(sessionToken);
  return e ? WeeklyQuestions.getHistory(e, page, pageSize) : { questions: [], hasMore: false };
}

function wqGetPoolCount() { return WeeklyQuestions.getPoolCount(); }
function wqInitSheets() { return WeeklyQuestions.initWeeklyQuestionSheets(); }
function wqGetPollFrequency() { return WeeklyQuestions.getPollFrequency(); }
function wqSetPollFrequency(sessionToken, freq) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Steward access required.' };
  return WeeklyQuestions.setPollFrequency(e, freq);
}

/**
 * v4.24.2 — Steward-triggered community poll draw.
 * Bypasses the day-of-week checks in autoSelectCommunityPoll so stewards
 * can release the community poll manually without waiting for the Monday trigger.
 * Steward-only. Calls selectRandomPoolQuestion() directly.
 * Returns { success, message, questionId?, text? }
 */
function wqManualDrawCommunityPoll(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Steward access required.' };
  try {
    var result = WeeklyQuestions.selectRandomPoolQuestion();
    if (typeof logAuditEvent === 'function') {
      logAuditEvent('POLL_COMMUNITY_MANUAL_DRAW', { steward: e, result: result.success ? 'ok' : 'failed' });
    }
    return result;
  } catch (err) {
    return { success: false, message: 'Draw failed: ' + err.message };
  }
}

// Time trigger — runs every Monday but internally respects frequency.
// For biweekly/monthly, it checks whether the period has actually changed
// before drawing a new community poll (avoids drawing on off-weeks).
// v4.24.3: Also skips if a community poll was already manually drawn this period.
function autoSelectCommunityPoll() {
  try {
    var freq = WeeklyQuestions.getPollFrequency();
    var today = new Date();

    // Biweekly: only draw on even-week Mondays
    if (freq === 'biweekly') {
      var jan4 = new Date(today.getFullYear(), 0, 4);
      var weekNum = Math.ceil(((today - jan4) / 86400000 + jan4.getDay() + 1) / 7);
      if (weekNum % 2 !== 0) {
        Logger.log('autoSelectCommunityPoll: biweekly — odd week ' + weekNum + ', skipping draw.');
        return;
      }
    }

    // Monthly: only draw on the 1st of the month
    if (freq === 'monthly') {
      if (today.getDate() !== 1) {
        Logger.log('autoSelectCommunityPoll: monthly — not the 1st, skipping draw.');
        return;
      }
    }

    // v4.24.3: Skip if a community poll already exists and is active for this period.
    // Prevents the trigger from overwriting a poll that was manually released by a steward.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var qSheet = ss.getSheetByName(SHEETS.WEEKLY_QUESTIONS);
      if (qSheet && qSheet.getLastRow() > 1) {
        // Derive current period key (Monday of this week, or period start per frequency)
        var d = new Date(today); d.setHours(0, 0, 0, 0);
        var day = d.getDay();
        var periodStart = new Date(d);
        if (freq === 'monthly') {
          periodStart = new Date(d.getFullYear(), d.getMonth(), 1);
        } else {
          periodStart.setDate(d.getDate() - (day === 0 ? 6 : day - 1)); // Monday
          if (freq === 'biweekly') {
            var jan4b = new Date(periodStart.getFullYear(), 0, 4);
            var wn = Math.ceil(((periodStart - jan4b) / 86400000 + jan4b.getDay() + 1) / 7);
            if (wn % 2 !== 0) periodStart.setDate(periodStart.getDate() - 7);
          }
        }
        var thisPeriod = periodStart.toISOString().split('T')[0];

        var rows = qSheet.getDataRange().getValues();
        // Use WeeklyQuestions.Q_COLS — single source of truth, no duplicated indices.
        var qc = WeeklyQuestions.Q_COLS;
        for (var i = 1; i < rows.length; i++) {
          var wk = rows[i][qc.WEEK_START];
          if (wk instanceof Date) wk = wk.toISOString().split('T')[0];
          var src    = String(rows[i][qc.SOURCE]);
          var active = String(rows[i][qc.ACTIVE]).toLowerCase();
          if (String(wk) === thisPeriod && src === 'community' && (active === 'true' || active === 'yes' || active === '1')) {
            Logger.log('autoSelectCommunityPoll: community poll already active for period ' + thisPeriod + ' — skipping draw.');
            return;
          }
        }
      }
    }

    var result = WeeklyQuestions.selectRandomPoolQuestion();
    Logger.log('autoSelectCommunityPoll: ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('autoSelectCommunityPoll error: ' + e.message);
  }
}

/**
 * Installs the Monday community poll draw trigger.
 * Run once after deployment via menu or Apps Script editor.
 */
function setupCommunityPollTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoSelectCommunityPoll') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('autoSelectCommunityPoll')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  Logger.log('setupCommunityPollTrigger: Community poll draw trigger installed (Mondays at 7 AM).');
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('Community poll draw trigger installed — fires every Monday at 7 AM.', 'Polls', 5);
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
}
