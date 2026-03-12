/**
 * ============================================================================
 * 26_QAForum.gs - Q&A Forum Module
 * ============================================================================
 *
 * Member-steward Q&A system with anonymous posting, upvoting, and moderation.
 *
 * Sheets:
 *   _QA_Forum   — member questions (11 columns)
 *   _QA_Answers  — answers to questions (8 columns)
 *
 * @fileoverview Q&A Forum IIFE module
 * @version 4.17.0
 * @requires 01_Core.gs, 06_Maintenance.gs
 */

var QAForum = (function () {

  // ═══════════════════════════════════════
  // Sheet Setup
  // ═══════════════════════════════════════

  function initQAForumSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken — getActiveSpreadsheet() returned null.');

    // _QA_Forum: ID | Author Email | Author Name | Is Anonymous | Question Text | Status | Upvote Count | Upvoters | Answer Count | Created | Updated
    var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!forumSheet) {
      forumSheet = ss.insertSheet(SHEETS.QA_FORUM);
      forumSheet.getRange(1, 1, 1, 11).setValues([[
        'ID', 'Author Email', 'Author Name', 'Is Anonymous', 'Question Text',
        'Status', 'Upvote Count', 'Upvoters', 'Answer Count', 'Created', 'Updated'
      ]]);
      // GAS-02: Use very-hidden so users cannot unhide PII-containing system sheets via menu
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(forumSheet); else forumSheet.hideSheet();
    }

    // _QA_Answers: ID | Question ID | Author Email | Author Name | Is Steward | Answer Text | Status | Created
    var answerSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!answerSheet) {
      answerSheet = ss.insertSheet(SHEETS.QA_ANSWERS);
      answerSheet.getRange(1, 1, 1, 8).setValues([[
        'ID', 'Question ID', 'Author Email', 'Author Name', 'Is Steward',
        'Answer Text', 'Status', 'Created'
      ]]);
      // GAS-02: Use very-hidden so users cannot unhide PII-containing system sheets via menu
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(answerSheet); else answerSheet.hideSheet();
    }
  }

  // ═══════════════════════════════════════
  // Cached Sheet Read Helper
  // ═══════════════════════════════════════

  // PERF-01: Cache full-sheet reads via CacheService to reduce redundant getDataRange() calls.
  // TTL defaults to 60s — acceptable staleness for read-heavy Q&A list loads.
  function _getCachedSheetData(sheetName, maxAgeSec) {
    maxAgeSec = maxAgeSec || 60;
    try {
      var cache = CacheService.getScriptCache();
      var cacheKey = 'qa_sheet_' + sheetName;
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (_e) { /* cache miss or parse error — fall through to fresh read */ }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var data = sheet.getDataRange().getValues();
    try {
      cache.put(cacheKey, JSON.stringify(data), maxAgeSec);
    } catch (_e) { /* CacheService has 100KB per-key limit — fail silently */ }
    return data;
  }

  // ═══════════════════════════════════════
  // Questions
  // ═══════════════════════════════════════

  function getQuestions(email, page, pageSize, sort) {
    page = page || 1;
    pageSize = pageSize || 20;
    sort = sort || 'recent';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { questions: [], total: 0, page: page, pageSize: pageSize };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet) {
      initQAForumSheets();
      sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    }
    if (!sheet || sheet.getLastRow() <= 1) return { questions: [], total: 0, page: page, pageSize: pageSize };

    var data = _getCachedSheetData(SHEETS.QA_FORUM, 60) || sheet.getDataRange().getValues();
    var questions = [];
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][5] || 'active').toLowerCase().trim();
      if (status === 'deleted') continue;
      questions.push({
        id: data[i][0],
        authorName: isTruthyValue(data[i][3]) ? 'Anonymous' : String(data[i][2] || 'Member'),
        isAnonymous: isTruthyValue(data[i][3]),
        questionText: String(data[i][4] || ''),
        status: status,
        upvoteCount: parseInt(data[i][6], 10) || 0,
        answerCount: parseInt(data[i][8], 10) || 0,
        created: data[i][9] instanceof Date ? _fmtDate(data[i][9]) : String(data[i][9] || ''),
        isOwner: email && String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim()
      });
    }

    // Sort
    if (sort === 'popular') {
      questions.sort(function (a, b) { return b.upvoteCount - a.upvoteCount; });
    } else {
      questions.sort(function (a, b) {
        var da = new Date(a.created), db = new Date(b.created);
        return db.getTime() - da.getTime();
      });
    }

    var total = questions.length;
    var start = (page - 1) * pageSize;
    var paged = questions.slice(start, start + pageSize);
    return { questions: paged, total: total, page: page, pageSize: pageSize };
  }

  function getQuestionDetail(email, questionId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return null;

    var data = _getCachedSheetData(SHEETS.QA_FORUM, 60) || sheet.getDataRange().getValues();
    var question = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === questionId) {
        var status = String(data[i][5] || 'active').toLowerCase().trim();
        if (status === 'deleted') return null;
        question = {
          id: data[i][0],
          authorName: isTruthyValue(data[i][3]) ? 'Anonymous' : String(data[i][2] || 'Member'),
          isAnonymous: isTruthyValue(data[i][3]),
          questionText: String(data[i][4] || ''),
          status: status,
          upvoteCount: parseInt(data[i][6], 10) || 0,
          answerCount: parseInt(data[i][8], 10) || 0,
          created: data[i][9] instanceof Date ? _fmtDate(data[i][9]) : String(data[i][9] || ''),
          isOwner: email && String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim()
        };
        break;
      }
    }
    if (!question) return null;

    // Fetch answers
    var ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    var answers = [];
    if (ansSheet && ansSheet.getLastRow() > 1) {
      var ansData = _getCachedSheetData(SHEETS.QA_ANSWERS, 60) || ansSheet.getDataRange().getValues();
      for (var j = 1; j < ansData.length; j++) {
        if (ansData[j][1] === questionId && String(ansData[j][6] || 'active').toLowerCase().trim() !== 'deleted') {
          answers.push({
            id: ansData[j][0],
            authorName: String(ansData[j][3] || 'Member'),
            isSteward: ansData[j][4] === true || ansData[j][4] === 'TRUE',
            answerText: String(ansData[j][5] || ''),
            status: String(ansData[j][6] || 'active'),
            created: ansData[j][7] instanceof Date ? _fmtDate(ansData[j][7]) : String(ansData[j][7] || '')
          });
        }
      }
    }
    question.answers = answers;
    return question;
  }

  function submitQuestion(email, name, text, isAnonymous) {
    if (!email || !text || !text.trim()) return { success: false, message: 'Question text is required.' };
    text = _sanitize(text.trim().substring(0, 2000));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet) { initQAForumSheets(); sheet = ss.getSheetByName(SHEETS.QA_FORUM); }

    // Rate limit: max 5 questions per hour per user
    var cache = CacheService.getScriptCache();
    var cacheKey = 'qa_rate_' + email.toLowerCase().trim();
    var count = parseInt(cache.get(cacheKey) || '0', 10);
    if (count >= 5) return { success: false, message: 'Rate limit reached. Please wait before posting again.' };
    cache.put(cacheKey, String(count + 1), 3600);

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'QA_' + Date.now().toString(36);
      var now = new Date();
      sheet.appendRow([
        id, email.toLowerCase().trim(), name || 'Member',
        isAnonymous ? true : false, text,
        'active', 0, '', 0, now, now
      ]);
      logAuditEvent('QA_QUESTION_SUBMITTED', 'Question ' + id + ' by ' + (isAnonymous ? 'anonymous' : maskEmail(email)));

      // Notify all stewards of the new unanswered question
      var preview = text.substring(0, 120) + (text.length > 120 ? '...' : '');
      var authorLabel = isAnonymous ? 'A member' : (name || 'A member');
      _createNotificationInternal_(
        'All Stewards',
        'Q&A Forum',
        'New Question in Q&A Forum',
        authorLabel + ' posted: "' + preview + '"'
      );

      return { success: true, questionId: id };
    } finally {
      lock.releaseLock();
    }
  }

  function submitAnswer(email, name, questionId, text, isSteward) {
    if (!email || !questionId || !text || !text.trim()) return { success: false, message: 'Answer text is required.' };
    if (!isSteward) return { success: false, message: 'Only stewards can post answers.' };
    text = _sanitize(text.trim().substring(0, 2000));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!ansSheet) { initQAForumSheets(); ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS); }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'ANS_' + Date.now().toString(36);
      ansSheet.appendRow([
        id, questionId, email.toLowerCase().trim(), name || 'Steward',
        true, text, 'active', new Date()
      ]);

      // Increment answer count on question and get author email for notification
      var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
      var questionAuthorEmail = null;
      var questionText = '';
      if (forumSheet && forumSheet.getLastRow() > 1) {
        var data = forumSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === questionId) {
            var currentCount = parseInt(data[i][8], 10) || 0;
            forumSheet.getRange(i + 1, 9).setValue(currentCount + 1);
            forumSheet.getRange(i + 1, 11).setValue(new Date());
            questionAuthorEmail = String(data[i][1] || '').toLowerCase().trim();
            questionText = String(data[i][4] || '').substring(0, 80);
            break;
          }
        }
      }

      // Notify the question author that their question received an answer
      if (questionAuthorEmail) {
        var preview = questionText + (questionText.length >= 80 ? '...' : '');
        _createNotificationInternal_(
          questionAuthorEmail,
          'Q&A Forum',
          'Your Question Got an Answer',
          (name || 'A steward') + ' answered your question: "' + preview + '"'
        );
      }

      logAuditEvent('QA_ANSWER_SUBMITTED', 'Answer ' + id + ' on question ' + questionId);
      return { success: true, answerId: id };
    } finally {
      lock.releaseLock();
    }
  }

  // ═══════════════════════════════════════
  // Upvoting
  // ═══════════════════════════════════════

  function upvoteQuestion(email, questionId) {
    if (!email || !questionId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Question not found.' };

    // Hash email for dedup (privacy-preserving)
    var emailHash = _hashEmail(email);

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === questionId) {
          var upvoters = String(data[i][7] || '');
          var upvoterList = upvoters ? upvoters.split(',') : [];
          var alreadyVoted = upvoterList.indexOf(emailHash) !== -1;

          if (alreadyVoted) {
            // Toggle off
            upvoterList.splice(upvoterList.indexOf(emailHash), 1);
          } else {
            upvoterList.push(emailHash);
          }

          sheet.getRange(i + 1, 7).setValue(upvoterList.length);
          sheet.getRange(i + 1, 8).setValue(upvoterList.join(','));
          return { success: true, upvoted: !alreadyVoted, newCount: upvoterList.length };
        }
      }
      return { success: false, message: 'Question not found.' };
    } finally {
      lock.releaseLock();
    }
  }

  // ═══════════════════════════════════════
  // Moderation (steward-only)
  // ═══════════════════════════════════════

  function moderateQuestion(stewardEmail, questionId, action) {
    if (!stewardEmail || !questionId || !action) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Question not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === questionId) {
        var newStatus = action === 'delete' ? 'deleted' : action === 'flag' ? 'flagged' : 'active';
        sheet.getRange(i + 1, 6).setValue(newStatus);
        sheet.getRange(i + 1, 11).setValue(new Date());
        logAuditEvent('QA_QUESTION_MODERATED', 'Question ' + questionId + ' ' + action + 'd by ' + maskEmail(stewardEmail));
        return { success: true };
      }
    }
    return { success: false, message: 'Question not found.' };
  }

  function moderateAnswer(stewardEmail, answerId, action) {
    if (!stewardEmail || !answerId || !action) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Answer not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === answerId) {
        var newStatus = action === 'delete' ? 'deleted' : action === 'flag' ? 'flagged' : 'active';
        sheet.getRange(i + 1, 7).setValue(newStatus);
        logAuditEvent('QA_ANSWER_MODERATED', 'Answer ' + answerId + ' ' + action + 'd by ' + maskEmail(stewardEmail));
        return { success: true };
      }
    }
    return { success: false, message: 'Answer not found.' };
  }

  function getFlaggedContent(stewardEmail) {
    if (!stewardEmail) return { questions: [], answers: [] };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { questions: [], answers: [] };
    var flaggedQ = [];
    var flaggedA = [];

    var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (forumSheet && forumSheet.getLastRow() > 1) {
      var qData = forumSheet.getDataRange().getValues();
      for (var i = 1; i < qData.length; i++) {
        if (String(qData[i][5]).toLowerCase().trim() === 'flagged') {
          flaggedQ.push({
            id: qData[i][0],
            authorName: qData[i][3] === true || qData[i][3] === 'TRUE' ? 'Anonymous' : String(qData[i][2] || 'Member'),
            questionText: String(qData[i][4] || '').substring(0, 200),
            created: qData[i][9] instanceof Date ? _fmtDate(qData[i][9]) : ''
          });
        }
      }
    }

    var ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (ansSheet && ansSheet.getLastRow() > 1) {
      var aData = ansSheet.getDataRange().getValues();
      for (var j = 1; j < aData.length; j++) {
        if (String(aData[j][6]).toLowerCase().trim() === 'flagged') {
          flaggedA.push({
            id: aData[j][0],
            questionId: aData[j][1],
            authorName: String(aData[j][3] || 'Member'),
            answerText: String(aData[j][5] || '').substring(0, 200),
            created: aData[j][7] instanceof Date ? _fmtDate(aData[j][7]) : ''
          });
        }
      }
    }

    return { questions: flaggedQ, answers: flaggedA };
  }

  // ═══════════════════════════════════════
  // Resolve
  // ═══════════════════════════════════════

  /**
   * Marks a question as resolved. Allowed by: question owner OR any steward.
   * @param {string} email - Caller's email (server-resolved)
   * @param {string} questionId
   * @param {boolean} isSteward - Whether caller is a steward
   */
  function resolveQuestion(email, questionId, isSteward) {
    if (!email || !questionId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Question not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === questionId) {
        var isOwner = String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim();
        if (!isOwner && !isSteward) return { success: false, message: 'Not authorized to resolve this question.' };
        var current = String(data[i][5] || '').toLowerCase().trim();
        if (current === 'deleted') return { success: false, message: 'Question not found.' };
        sheet.getRange(i + 1, 6).setValue('resolved');
        sheet.getRange(i + 1, 11).setValue(new Date());
        logAuditEvent('QA_QUESTION_RESOLVED', 'Question ' + questionId + ' resolved by ' + maskEmail(email));
        return { success: true };
      }
    }
    return { success: false, message: 'Question not found.' };
  }

  // ═══════════════════════════════════════
  // Helpers
  // ═══════════════════════════════════════

  /**
   * Internal notification writer — bypasses steward auth for system-generated notifications.
   * Writes directly to the Notifications sheet using the same schema as sendWebAppNotification().
   * @private
   */
  function _createNotificationInternal_(recipient, type, title, message) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) return;
      var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
      if (!sheet) {
        if (typeof createNotificationsSheet === 'function') sheet = createNotificationsSheet(ss);
        else return;
      }
      var allData = sheet.getDataRange().getValues();
      var maxNum = 0;
      var C = NOTIFICATIONS_COLS;
      for (var i = 1; i < allData.length; i++) {
        var existId = String(allData[i][C.NOTIFICATION_ID - 1] || '');
        var match = existId.match(/NOTIF-(\d+)/);
        if (match) { var num = parseInt(match[1], 10); if (num > maxNum) maxNum = num; }
      }
      var nextId = 'NOTIF-' + String(maxNum + 1).padStart(3, '0');
      var tz = Session.getScriptTimeZone();
      var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
      sheet.appendRow([
        nextId,
        escapeForFormula ? escapeForFormula(recipient) : recipient,
        escapeForFormula ? escapeForFormula(type) : type,
        escapeForFormula ? escapeForFormula(title) : title,
        escapeForFormula ? escapeForFormula(message) : message,
        'Normal', 'system', 'Q&A Forum', today, '', '', 'Active', 'Dismissible'
      ]);
    } catch (e) {
      Logger.log('QAForum._createNotificationInternal_ error: ' + e.message);
    }
  }

  // Delegate to shared helpers in 01_Core.gs (eliminates duplicate definitions)
  function _fmtDate(date) { return fmtDateShort_(date); }
  function _hashEmail(email) { return hashEmail_(email); }

  function _sanitize(text) {
    if (typeof escapeForFormula === 'function') text = escapeForFormula(text);
    return text;
  }

  // ═══════════════════════════════════════
  // Public API
  // ═══════════════════════════════════════

  return {
    initQAForumSheets: initQAForumSheets,
    getQuestions: getQuestions,
    getQuestionDetail: getQuestionDetail,
    submitQuestion: submitQuestion,
    submitAnswer: submitAnswer,
    upvoteQuestion: upvoteQuestion,
    moderateQuestion: moderateQuestion,
    moderateAnswer: moderateAnswer,
    getFlaggedContent: getFlaggedContent,
    resolveQuestion: resolveQuestion
  };

})();


// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════

function qaGetQuestions(sessionToken, page, pageSize, sort) { var e = _resolveCallerEmail(sessionToken); return QAForum.getQuestions(e, page, pageSize, sort); }
function qaGetQuestionDetail(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); return QAForum.getQuestionDetail(e, questionId); }
function qaSubmitQuestion(sessionToken, name, text, isAnonymous) { var e = _resolveCallerEmail(sessionToken); return QAForum.submitQuestion(e, name, text, isAnonymous); }
function qaSubmitAnswer(sessionToken, name, questionId, text, isSteward) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.submitAnswer(e, name, questionId, text, true); }
function qaUpvoteQuestion(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); return QAForum.upvoteQuestion(e, questionId); }
function qaModerateQuestion(sessionToken, questionId, action) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.moderateQuestion(e, questionId, action); }
function qaModerateAnswer(sessionToken, answerId, action) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.moderateAnswer(e, answerId, action); }
function qaGetFlaggedContent(sessionToken) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.', items: [] }; return QAForum.getFlaggedContent(e); }
function qaResolveQuestion(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; var isSteward = false; try { var auth = checkWebAppAuthorization('steward', sessionToken); isSteward = auth.isAuthorized; } catch(_) {} return QAForum.resolveQuestion(e, questionId, isSteward); }
function qaInitSheets() { return QAForum.initQAForumSheets(); }
