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
      forumSheet.hideSheet();
    }

    // _QA_Answers: ID | Question ID | Author Email | Author Name | Is Steward | Answer Text | Status | Created
    var answerSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!answerSheet) {
      answerSheet = ss.insertSheet(SHEETS.QA_ANSWERS);
      answerSheet.getRange(1, 1, 1, 8).setValues([[
        'ID', 'Question ID', 'Author Email', 'Author Name', 'Is Steward',
        'Answer Text', 'Status', 'Created'
      ]]);
      answerSheet.hideSheet();
    }
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

    var data = sheet.getDataRange().getValues();
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

    var data = sheet.getDataRange().getValues();
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
      var ansData = ansSheet.getDataRange().getValues();
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
      logAuditEvent('QA_QUESTION_SUBMITTED', 'Question ' + id + ' by ' + (isAnonymous ? 'anonymous' : email));
      return { success: true, questionId: id };
    } finally {
      lock.releaseLock();
    }
  }

  function submitAnswer(email, name, questionId, text, isSteward) {
    if (!email || !questionId || !text || !text.trim()) return { success: false, message: 'Answer text is required.' };
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
        id, questionId, email.toLowerCase().trim(), name || 'Member',
        isSteward ? true : false, text, 'active', new Date()
      ]);

      // Increment answer count on question
      var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
      if (forumSheet && forumSheet.getLastRow() > 1) {
        var data = forumSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === questionId) {
            var currentCount = parseInt(data[i][8], 10) || 0;
            forumSheet.getRange(i + 1, 9).setValue(currentCount + 1);
            forumSheet.getRange(i + 1, 11).setValue(new Date());
            break;
          }
        }
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
        logAuditEvent('QA_QUESTION_MODERATED', 'Question ' + questionId + ' ' + action + 'd by ' + stewardEmail);
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
        logAuditEvent('QA_ANSWER_MODERATED', 'Answer ' + answerId + ' ' + action + 'd by ' + stewardEmail);
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
  // Helpers
  // ═══════════════════════════════════════

  function _fmtDate(date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  }

  function _hashEmail(email) {
    var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email.toLowerCase().trim());
    return hash.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 12);
  }

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
    getFlaggedContent: getFlaggedContent
  };

})();


// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════

function qaGetQuestions(email, page, pageSize, sort) { var e = _resolveCallerEmail() || email; return QAForum.getQuestions(e, page, pageSize, sort); }
function qaGetQuestionDetail(email, questionId) { var e = _resolveCallerEmail() || email; return QAForum.getQuestionDetail(e, questionId); }
function qaSubmitQuestion(email, name, text, isAnonymous) { var e = _resolveCallerEmail() || email; return QAForum.submitQuestion(e, name, text, isAnonymous); }
function qaSubmitAnswer(email, name, questionId, text, isSteward) { var e = _resolveCallerEmail() || email; return QAForum.submitAnswer(e, name, questionId, text, isSteward); }
function qaUpvoteQuestion(email, questionId) { var e = _resolveCallerEmail() || email; return QAForum.upvoteQuestion(e, questionId); }
function qaModerateQuestion(stewardEmail, questionId, action) { var e = _requireStewardAuth(); if (!e) return null; return QAForum.moderateQuestion(e, questionId, action); }
function qaModerateAnswer(stewardEmail, answerId, action) { var e = _requireStewardAuth(); if (!e) return null; return QAForum.moderateAnswer(e, answerId, action); }
function qaGetFlaggedContent(stewardEmail) { var e = _requireStewardAuth(); if (!e) return null; return QAForum.getFlaggedContent(e); }
function qaInitSheets() { return QAForum.initQAForumSheets(); }
