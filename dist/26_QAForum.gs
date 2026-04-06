/**
 * ============================================================================
 * 26_QAForum.gs - Q&A Forum Module
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Q&A forum for member-steward communication. Members post questions
 *   (optionally anonymous), stewards answer. Features: upvoting, question
 *   status (open/resolved/reopened), answer moderation, unanswered question
 *   count badge. Two hidden sheets:
 *     _QA_Forum   — questions (11 columns)
 *     _QA_Answers — answers (8 columns)
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Only stewards can post answers — this ensures members get authoritative
 *   responses about union procedures. Anonymous posting encourages questions
 *   about sensitive topics (workplace issues, rights). Sheets are very-hidden
 *   (setSheetVeryHidden_) to protect PII (author emails). Unanswered count
 *   drives the badge indicator in the SPA navigation, ensuring stewards
 *   notice pending questions.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Q&A tab in the SPA shows no questions/answers. Members can't ask
 *   questions. Stewards can't respond. The unanswered badge shows 0
 *   (stewards won't notice pending questions). Existing Q&A data is
 *   preserved in sheets but inaccessible through the UI.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS), 06_Maintenance.gs (logAuditEvent).
 *   Used by SPA Q&A views and the navigation badge system.
 *
 * @version 4.51.0
 */

var QAForum = (function () {

  // ═══════════════════════════════════════
  // Sheet Setup
  // ═══════════════════════════════════════

  /**
   * Initializes the _QA_Forum and _QA_Answers hidden sheets if they do not exist.
   * @returns {void}
   */
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

  /**
   * Returns cached sheet data from CacheService, falling back to a live read.
   * @param {string} sheetName - Name of the sheet to read.
   * @param {number} [maxAgeSec=60] - Cache TTL in seconds.
   * @returns {Array[]|null} 2D array of sheet values, or null if empty/missing.
   */
  function _getCachedSheetData(sheetName, maxAgeSec) {
    maxAgeSec = maxAgeSec || 60;
    try {
      var cache = CacheService.getScriptCache();
      var cacheKey = 'qa_sheet_' + sheetName;
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (_e) { log_('QAForum._getCachedSheetData', 'Error reading cache: ' + (_e.message || _e)); }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var data = sheet.getDataRange().getValues();
    try {
      var jsonStr = JSON.stringify(data);
      if (jsonStr.length < 90000) { // CacheService 100KB limit per key — leave margin
        cache.put(cacheKey, jsonStr, maxAgeSec);
      }
      // If too large, skip caching — data will be read fresh each time
    } catch (_e) { log_('QAForum._getCachedSheetData', 'Error writing cache: ' + (_e.message || _e)); }
    return data;
  }

  // ═══════════════════════════════════════
  // Questions
  // ═══════════════════════════════════════

  /**
   * Returns a paginated, sorted list of Q&A questions.
   * @param {string} email - Caller's email for ownership detection.
   * @param {number} [page=1] - Page number (1-based).
   * @param {number} [pageSize=20] - Results per page.
   * @param {string} [sort='recent'] - Sort order: 'recent' or 'popular'.
   * @param {boolean} [showResolved] - Whether to include resolved questions.
   * @returns {{questions: Object[], total: number, page: number, pageSize: number}}
   */
  function getQuestions(email, page, pageSize, sort, showResolved) {
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
      if (!showResolved && status === 'resolved') continue;
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

  /**
   * Returns a single question with its answers, or null if not found/deleted.
   * @param {string} email - Caller's email for ownership detection.
   * @param {string} questionId - The question ID to look up.
   * @returns {Object|null} Question object with nested answers array, or null.
   */
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
            isSteward: isTruthyValue(ansData[j][4]),
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

  /**
   * Submits a new question to the Q&A forum with rate limiting (5/hour).
   * @param {string} email - Author's email.
   * @param {string} name - Author's display name.
   * @param {string} text - Question body (max 2000 chars, sanitized).
   * @param {boolean} isAnonymous - Whether to hide author identity.
   * @returns {{success: boolean, questionId?: string, message?: string}}
   */
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
    if (!lock.tryLock(10000)) return { success: false, message: 'System busy. Please try again in a moment.' };
    try {
      var id = 'QA_' + Date.now().toString(36);
      var now = new Date();
      sheet.appendRow([
        id, email.toLowerCase().trim(), escapeForFormula(name || 'Member'),
        isAnonymous ? true : false, text,
        'active', 0, '', 0, now, now
      ]);
      _invalidateCache(SHEETS.QA_FORUM);
      logAuditEvent('QA_QUESTION_SUBMITTED', 'Question ' + id + ' by ' + (isAnonymous ? 'anonymous' : maskEmail(email)));

      // Notify all stewards of the new unanswered question
      var preview = text.substring(0, 120) + (text.length > 120 ? '...' : '');
      var authorLabel = isAnonymous ? 'A member' : (name || 'A member');
      var notificationSent = false;
      try {
        _createNotificationInternal_(
          'All Stewards',
          'Q&A Forum',
          'New Question in Q&A Forum',
          authorLabel + ' posted: "' + preview + '"'
        );
        notificationSent = true;
      } catch (notifErr) {
        log_('QA submitQuestion', 'steward notification failed: ' + notifErr.message);
      }

      return { success: true, questionId: id, notificationSent: notificationSent };
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Submits a steward answer to a question and increments the answer count.
   * @param {string} email - Steward's email.
   * @param {string} name - Steward's display name.
   * @param {string} questionId - Target question ID.
   * @param {string} text - Answer body (max 2000 chars, sanitized).
   * @param {boolean} isSteward - Must be true; non-stewards are rejected.
   * @returns {{success: boolean, answerId?: string, message?: string}}
   */
  function submitAnswer(email, name, questionId, text, isSteward) {
    if (!email || !questionId || !text || !text.trim()) return { success: false, message: 'Answer text is required.' };
    if (!isSteward) return { success: false, message: 'Only stewards can post answers.' };
    text = _sanitize(text.trim().substring(0, 2000));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!ansSheet) { initQAForumSheets(); ansSheet = ss.getSheetByName(SHEETS.QA_ANSWERS); }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'System busy. Please try again in a moment.' };
    try {
      var id = 'ANS_' + Date.now().toString(36);
      ansSheet.appendRow([
        id, questionId, email.toLowerCase().trim(), escapeForFormula(name || 'Steward'),
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

      _invalidateCache(SHEETS.QA_FORUM);
      _invalidateCache(SHEETS.QA_ANSWERS);
      logAuditEvent('QA_ANSWER_SUBMITTED', 'Answer ' + id + ' on question ' + questionId);
      return { success: true, answerId: id, _authorEmail: questionAuthorEmail, _questionText: questionText, _stewardName: name };
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Wraps submitAnswer to send a notification to the question author after the lock is released.
   * @param {string} email - Steward's email.
   * @param {string} name - Steward's display name.
   * @param {string} questionId - Target question ID.
   * @param {string} text - Answer body.
   * @param {boolean} isSteward - Must be true.
   * @returns {{success: boolean, answerId?: string, message?: string}}
   */
  function submitAnswerWithNotify(email, name, questionId, text, isSteward) {
    var result = submitAnswer(email, name, questionId, text, isSteward);
    if (result.success && result._authorEmail) {
      try {
        var preview = result._questionText + (result._questionText.length >= 80 ? '...' : '');
        _createNotificationInternal_(
          result._authorEmail,
          'Q&A Forum',
          'Your Question Got an Answer',
          (result._stewardName || 'A steward') + ' answered your question: "' + preview + '"'
        );
      } catch (notifErr) {
        log_('QA submitAnswer', 'author notification failed: ' + notifErr.message);
      }
      // Strip internal fields from response
      delete result._authorEmail;
      delete result._questionText;
      delete result._stewardName;
    }
    return result;
  }

  // ═══════════════════════════════════════
  // Upvoting
  // ═══════════════════════════════════════

  /**
   * Toggles an upvote on a question using hashed email for deduplication.
   * @param {string} email - Voter's email.
   * @param {string} questionId - Question to upvote/un-upvote.
   * @returns {{success: boolean, upvoted?: boolean, newCount?: number, message?: string}}
   */
  function upvoteQuestion(email, questionId) {
    if (!email || !questionId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Question not found.' };

    // Hash email for dedup (privacy-preserving)
    var emailHash = _hashEmail(email);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: 'System busy. Please try again in a moment.' };
    try {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === questionId) {
          // v4.51.1: Block self-upvoting (BUG-4-01)
          var questionAuthor = String(data[i][1] || '').toLowerCase().trim();
          if (questionAuthor === email.toLowerCase().trim()) {
            return { success: false, message: 'You cannot upvote your own question.' };
          }
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
          _invalidateCache(SHEETS.QA_FORUM);
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

  /**
   * Moderates a question by setting its status to deleted, flagged, or active.
   * @param {string} stewardEmail - Moderating steward's email.
   * @param {string} questionId - Question to moderate.
   * @param {string} action - 'delete', 'flag', or 'restore'.
   * @returns {{success: boolean, message?: string}}
   */
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
        _invalidateCache(SHEETS.QA_FORUM);
        logAuditEvent('QA_QUESTION_MODERATED', 'Question ' + questionId + ' ' + action + 'd by ' + maskEmail(stewardEmail));
        return { success: true };
      }
    }
    return { success: false, message: 'Question not found.' };
  }

  /**
   * Moderates an answer and adjusts the parent question's answer count if visibility changes.
   * @param {string} stewardEmail - Moderating steward's email.
   * @param {string} answerId - Answer to moderate.
   * @param {string} action - 'delete', 'flag', or 'restore'.
   * @returns {{success: boolean, message?: string}}
   */
  function moderateAnswer(stewardEmail, answerId, action) {
    if (!stewardEmail || !answerId || !action) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_ANSWERS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Answer not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === answerId) {
        var oldStatus = String(data[i][6] || 'active').toLowerCase().trim();
        var newStatus = action === 'delete' ? 'deleted' : action === 'flag' ? 'flagged' : 'active';
        sheet.getRange(i + 1, 7).setValue(newStatus);
        _invalidateCache(SHEETS.QA_ANSWERS);

        // Adjust answer count on parent question when visibility changes
        var wasVisible = oldStatus !== 'deleted';
        var isVisible = newStatus !== 'deleted';
        if (wasVisible !== isVisible) {
          var questionId = data[i][1];
          var forumSheet = ss.getSheetByName(SHEETS.QA_FORUM);
          if (forumSheet && forumSheet.getLastRow() > 1) {
            var qData = forumSheet.getDataRange().getValues();
            for (var j = 1; j < qData.length; j++) {
              if (qData[j][0] === questionId) {
                var currentCount = parseInt(qData[j][8], 10) || 0;
                var delta = isVisible ? 1 : -1;
                forumSheet.getRange(j + 1, 9).setValue(Math.max(0, currentCount + delta));
                forumSheet.getRange(j + 1, 11).setValue(new Date());
                _invalidateCache(SHEETS.QA_FORUM);
                break;
              }
            }
          }
        }

        logAuditEvent('QA_ANSWER_MODERATED', 'Answer ' + answerId + ' ' + action + 'd by ' + maskEmail(stewardEmail));
        return { success: true };
      }
    }
    return { success: false, message: 'Answer not found.' };
  }

  /**
   * Returns all flagged questions and answers for steward moderation review.
   * @param {string} stewardEmail - Requesting steward's email.
   * @returns {{questions: Object[], answers: Object[]}}
   */
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
            authorName: isTruthyValue(qData[i][3]) ? 'Anonymous' : String(qData[i][2] || 'Member'),
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
        _invalidateCache(SHEETS.QA_FORUM);
        logAuditEvent('QA_QUESTION_RESOLVED', 'Question ' + questionId + ' resolved by ' + maskEmail(email));
        return { success: true };
      }
    }
    return { success: false, message: 'Question not found.' };
  }

  // ═══════════════════════════════════════
  // Reopen
  // ═══════════════════════════════════════

  /**
   * Reopens a resolved question. Allowed by: question owner OR any steward.
   * @param {string} email - Caller's email (server-resolved)
   * @param {string} questionId
   * @param {boolean} isSteward - Whether caller is a steward
   */
  function reopenQuestion(email, questionId, isSteward) {
    if (!email || !questionId) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Question not found.' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === questionId) {
        var isOwner = String(data[i][1]).toLowerCase().trim() === email.toLowerCase().trim();
        if (!isOwner && !isSteward) return { success: false, message: 'Not authorized to reopen this question.' };
        var current = String(data[i][5] || '').toLowerCase().trim();
        if (current !== 'resolved') return { success: false, message: 'Only resolved questions can be reopened.' };
        sheet.getRange(i + 1, 6).setValue('active');
        sheet.getRange(i + 1, 11).setValue(new Date());
        _invalidateCache(SHEETS.QA_FORUM);
        logAuditEvent('QA_QUESTION_REOPENED', 'Question ' + questionId + ' reopened by ' + maskEmail(email));
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
      log_('QAForum._createNotificationInternal_ error', e.message);
    }
  }

  /**
   * Removes a sheet's cached data from CacheService to force a fresh read.
   * @param {string} sheetName - Name of the sheet whose cache entry to remove.
   * @returns {void}
   */
  function _invalidateCache(sheetName) {
    try {
      var cache = CacheService.getScriptCache();
      cache.remove('qa_sheet_' + sheetName);
    } catch (_e) { log_('QAForum._invalidateCache', 'Error: ' + (_e.message || _e)); }
  }

  /**
   * Formats a Date as a short string by delegating to fmtDateShort_ in 01_Core.gs.
   * @param {Date} date - Date to format.
   * @returns {string} Formatted date string.
   */
  function _fmtDate(date) { return fmtDateShort_(date); }

  /**
   * Hashes an email for privacy-preserving deduplication via hashEmail_ in 01_Core.gs.
   * @param {string} email - Email address to hash.
   * @returns {string} Hashed email string.
   */
  function _hashEmail(email) { return hashEmail_(email); }

  /**
   * Sanitizes user input text using escapeForFormula to prevent formula injection.
   * @param {string} text - Raw user input.
   * @returns {string} Sanitized text.
   */
  function _sanitize(text) {
    if (typeof escapeForFormula === 'function') text = escapeForFormula(text);
    return text;
  }

  /**
   * Lightweight count of unanswered questions — avoids building full question objects.
   * Used by batch data to skip the expensive getQuestions(email, 1, 999, ...) call.
   */
  function getUnansweredCount() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return 0;
    var sheet = ss.getSheetByName(SHEETS.QA_FORUM);
    if (!sheet || sheet.getLastRow() <= 1) return 0;

    var data = _getCachedSheetData(SHEETS.QA_FORUM, 60) || sheet.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][5] || 'active').toLowerCase().trim();
      if (status === 'deleted' || status === 'resolved') continue;
      var answerCount = parseInt(data[i][8], 10) || 0;
      if (answerCount === 0) count++;
    }
    return count;
  }

  /**
   * Returns questions visible across all units (org-wide collaboration).
   * All questions are shared — filtering is done client-side by unit badge.
   * @param {string} userEmail - Caller's email for ownership detection.
   * @param {number} [page] - Page number (1-based).
   * @param {number} [pageSize] - Results per page.
   * @param {string} [sort] - Sort order: 'recent' or 'popular'.
   * @returns {{questions: Object[], total: number, page: number, pageSize: number}}
   */
  function getOrgWideQuestions(userEmail, page, pageSize, sort, showResolved) {
    // Reuse existing getQuestions — pass through user's showResolved preference
    return getQuestions(userEmail, page, pageSize, sort, showResolved);
  }

  // ═══════════════════════════════════════
  // Public API
  // ═══════════════════════════════════════

  return {
    initQAForumSheets: initQAForumSheets,
    getQuestions: getQuestions,
    getOrgWideQuestions: getOrgWideQuestions,
    getQuestionDetail: getQuestionDetail,
    getUnansweredCount: getUnansweredCount,
    submitQuestion: submitQuestion,
    submitAnswer: submitAnswerWithNotify,
    upvoteQuestion: upvoteQuestion,
    moderateQuestion: moderateQuestion,
    moderateAnswer: moderateAnswer,
    getFlaggedContent: getFlaggedContent,
    resolveQuestion: resolveQuestion,
    reopenQuestion: reopenQuestion
  };

})();
// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════

/** @param {string} sessionToken @param {number} [page] @param {number} [pageSize] @param {string} [sort] @param {boolean} [showResolved] @returns {Object} Paginated question list. */
function qaGetQuestions(sessionToken, page, pageSize, sort, showResolved) { var e = _resolveCallerEmail(sessionToken); if (!e) return { questions: [], total: 0, page: 1, pageSize: pageSize || 20 }; return QAForum.getQuestions(e, page, pageSize, sort, showResolved); }
/** @param {string} sessionToken @param {string} questionId @returns {Object|null} Question with answers. */
function qaGetQuestionDetail(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); if (!e) return null; return QAForum.getQuestionDetail(e, questionId); }
/** @param {string} sessionToken @param {string} name @param {string} text @param {boolean} isAnonymous @returns {Object} Submission result. */
function qaSubmitQuestion(sessionToken, name, text, isAnonymous) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; return QAForum.submitQuestion(e, name, text, isAnonymous); }
/** @param {string} sessionToken @param {string} name @param {string} questionId @param {string} text @param {boolean} isSteward @returns {Object} Submission result (steward-only). */
function qaSubmitAnswer(sessionToken, name, questionId, text, isSteward) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.submitAnswer(e, name, questionId, text, true); }
/** @param {string} sessionToken @param {string} questionId @returns {Object} Toggle result with new count. */
function qaUpvoteQuestion(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; return QAForum.upvoteQuestion(e, questionId); }
/** @param {string} sessionToken @param {string} questionId @param {string} action @returns {Object} Moderation result (steward-only). */
function qaModerateQuestion(sessionToken, questionId, action) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.moderateQuestion(e, questionId, action); }
/** @param {string} sessionToken @param {string} answerId @param {string} action @returns {Object} Moderation result (steward-only). */
function qaModerateAnswer(sessionToken, answerId, action) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.' }; return QAForum.moderateAnswer(e, answerId, action); }
/** @param {string} sessionToken @returns {Object} Flagged questions and answers (steward-only). */
function qaGetFlaggedContent(sessionToken) { var e = _requireStewardAuth(sessionToken); if (!e) return { success: false, message: 'Steward access required.', items: [] }; return QAForum.getFlaggedContent(e); }
/** @param {string} sessionToken @param {string} questionId @returns {Object} Resolve result (owner or steward). */
function qaResolveQuestion(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; var isSteward = false; try { var auth = checkWebAppAuthorization('steward', sessionToken); isSteward = auth.isAuthorized; } catch (_e) { log_('qaResolveQuestion', 'Error checking steward auth: ' + (_e.message || _e)); } return QAForum.resolveQuestion(e, questionId, isSteward); }
/** @param {string} sessionToken @param {string} questionId @returns {Object} Reopen result (owner or steward). */
function qaReopenQuestion(sessionToken, questionId) { var e = _resolveCallerEmail(sessionToken); if (!e) return { success: false, message: 'Not authenticated.' }; var isSteward = false; try { var auth = checkWebAppAuthorization('steward', sessionToken); isSteward = auth.isAuthorized; } catch (_e) { log_('qaReopenQuestion', 'Error checking steward auth: ' + (_e.message || _e)); } return QAForum.reopenQuestion(e, questionId, isSteward); }
/** @param {string} sessionToken @param {number} [page] @param {number} [pageSize] @param {string} [sort] @param {boolean} [showResolved] @returns {Object} Org-wide paginated question list. */
function qaGetOrgWideQuestions(sessionToken, page, pageSize, sort, showResolved) { var e = _resolveCallerEmail(sessionToken); if (!e) return { questions: [], total: 0, page: 1, pageSize: pageSize || 20 }; return QAForum.getOrgWideQuestions(e, page, pageSize, sort, showResolved); }
/** @returns {void} Initializes Q&A forum sheets (no auth required — setup only). */
function qaInitSheets() { return QAForum.initQAForumSheets(); }
