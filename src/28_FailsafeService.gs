/**
 * ============================================================================
 * 28_FailsafeService.gs - Data Failsafe / Recovery System
 * ============================================================================
 *
 * Scheduled email digests and Google Drive CSV backups for data resilience.
 *
 * Sheet:
 *   _Failsafe_Config — member digest preferences (7 columns)
 *
 * @fileoverview Failsafe Service IIFE module
 * @version 4.22.8
 * @requires 01_Core.gs, 06_Maintenance.gs
 */

var FailsafeService = (function () {

  var BACKUP_FOLDER_NAME = 'DDS_Dashboard_Backups';
  var MAX_BACKUP_FILES = 52; // ~1 year of weekly backups per sheet

  // ═══════════════════════════════════════
  // Sheet Setup
  // ═══════════════════════════════════════

  function initFailsafeSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet binding broken — getActiveSpreadsheet() returned null.');
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.FAILSAFE_CONFIG);
      sheet.getRange(1, 1, 1, 7).setValues([[
        'Email', 'Digest Enabled', 'Digest Frequency', 'Last Digest Sent',
        'Include Grievances', 'Include Workload', 'Include Tasks'
      ]]);
      sheet.hideSheet();
    }
    return sheet;
  }

  // ═══════════════════════════════════════
  // Digest Configuration
  // ═══════════════════════════════════════

  function getDigestConfig(email) {
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
    }

    var data = sheet.getDataRange().getValues();
    var eml = email.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === eml) {
        return {
          enabled: data[i][1] === true || data[i][1] === 'TRUE',
          frequency: String(data[i][2] || 'weekly').toLowerCase(),
          lastSent: data[i][3] instanceof Date ? _fmtDate(data[i][3]) : '',
          includeGrievances: data[i][4] !== false && data[i][4] !== 'FALSE',
          includeWorkload: data[i][5] !== false && data[i][5] !== 'FALSE',
          includeTasks: data[i][6] !== false && data[i][6] !== 'FALSE'
        };
      }
    }
    return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
  }

  function updateDigestConfig(email, config) {
    if (!email || !config) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet) { initFailsafeSheet(); sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG); }

    var eml = email.toLowerCase().trim();
    var frequency = config.frequency === 'monthly' ? 'monthly' : 'weekly';

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (sheet.getLastRow() > 1) {
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][0]).toLowerCase().trim() === eml) {
            sheet.getRange(i + 1, 2).setValue(config.enabled ? true : false);
            sheet.getRange(i + 1, 3).setValue(frequency);
            sheet.getRange(i + 1, 5).setValue(config.includeGrievances !== false);
            sheet.getRange(i + 1, 6).setValue(config.includeWorkload !== false);
            sheet.getRange(i + 1, 7).setValue(config.includeTasks !== false);
            return { success: true, message: 'Digest settings updated.' };
          }
        }
      }
      // New row
      sheet.appendRow([eml, config.enabled ? true : false, frequency, '', config.includeGrievances !== false, config.includeWorkload !== false, config.includeTasks !== false]);
      return { success: true, message: 'Digest settings saved.' };
    } finally {
      lock.releaseLock();
    }
  }

  // ═══════════════════════════════════════
  // Scheduled Digest Processing
  // ═══════════════════════════════════════

  function processScheduledDigests() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { processed: 0 };
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet || sheet.getLastRow() <= 1) return { processed: 0 };

    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var processed = 0;

    for (var i = 1; i < data.length; i++) {
      var enabled = data[i][1] === true || data[i][1] === 'TRUE';
      if (!enabled) continue;

      var email = String(data[i][0]).toLowerCase().trim();
      if (!email) continue;

      var frequency = String(data[i][2] || 'weekly').toLowerCase();
      var lastSent = data[i][3];
      var daysSinceLast = lastSent instanceof Date ? Math.ceil((now.getTime() - lastSent.getTime()) / 86400000) : 999;

      var shouldSend = false;
      if (frequency === 'weekly' && daysSinceLast >= 7) shouldSend = true;
      if (frequency === 'monthly' && daysSinceLast >= 30) shouldSend = true;

      if (!shouldSend) continue;

      // Check MailApp quota
      var remaining = MailApp.getRemainingDailyQuota();
      if (remaining < 5) {
        Logger.log('Email quota low (' + remaining + '), stopping digest processing.');
        break;
      }

      try {
        var config = {
          includeGrievances: data[i][4] !== false && data[i][4] !== 'FALSE',
          includeWorkload: data[i][5] !== false && data[i][5] !== 'FALSE',
          includeTasks: data[i][6] !== false && data[i][6] !== 'FALSE'
        };
        var body = _composeMemberDigest(email, config);
        if (body) {
          var sendLock = LockService.getScriptLock();
          // Try to acquire lock — if another execution already sent, skip
          if (!sendLock.tryLock(5000)) {
            Logger.log('Could not acquire send lock for ' + email + ', skipping to avoid duplicate.');
            continue;
          }
          try {
            // Re-read lastSent inside the lock to guard against concurrent executions
            var freshRow = sheet.getRange(i + 1, 1, 1, 4).getValues()[0];
            var freshLastSent = freshRow[3];
            var freshDays = freshLastSent instanceof Date ? Math.ceil((now.getTime() - freshLastSent.getTime()) / 86400000) : 999;
            var stillDue = (frequency === 'weekly' && freshDays >= 7) || (frequency === 'monthly' && freshDays >= 30);
            if (!stillDue) { continue; } // Another execution already sent
            MailApp.sendEmail({
              to: email,
              subject: 'Your Union Dashboard Digest',
              htmlBody: body,
              noReply: true
            });
            sheet.getRange(i + 1, 4).setValue(now);
            processed++;
          } finally {
            sendLock.releaseLock();
          }
        }
      } catch (err) {
        Logger.log('Digest send error for ' + email + ': ' + err.message);
      }
    }

    logAuditEvent('FAILSAFE_DIGESTS_PROCESSED', processed + ' digests sent');
    return { processed: processed };
  }

  function _composeMemberDigest(email, config) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sections = [];

    // Grievances
    if (config.includeGrievances && typeof DataService !== 'undefined') {
      try {
        var grievances = DataService.getMemberGrievances(email);
        if (grievances && grievances.length > 0) {
          var gHtml = '<h3>Your Grievances (' + grievances.length + ')</h3><ul>';
          grievances.forEach(function (g) {
            gHtml += '<li><strong>' + escapeHtml(g.status || 'Open') + '</strong> — ' + escapeHtml(g.issueType || g.type || 'Case') + ' (Filed: ' + escapeHtml(g.dateFiled || '') + ')</li>';
          });
          gHtml += '</ul>';
          sections.push(gHtml);
        }
      } catch (e) { Logger.log('Digest grievance error: ' + e.message); }
    }

    // Workload
    if (config.includeWorkload) {
      try {
        var wSheet = ss.getSheetByName(SHEETS.WORKLOAD_VAULT);
        if (wSheet && wSheet.getLastRow() > 1) {
          var wData = wSheet.getDataRange().getValues();
          var wHeaders = wData[0];
          var memberEntries = [];
          var emailCol = -1;
          for (var h = 0; h < wHeaders.length; h++) {
            if (String(wHeaders[h]).toLowerCase().indexOf('email') !== -1) { emailCol = h; break; }
          }
          if (emailCol !== -1) {
            for (var w = 1; w < wData.length; w++) {
              if (String(wData[w][emailCol]).toLowerCase().trim() === email.toLowerCase().trim()) {
                memberEntries.push(wData[w]);
              }
            }
          }
          if (memberEntries.length > 0) {
            sections.push('<h3>Workload History (' + memberEntries.length + ' entries)</h3><p>You have ' + memberEntries.length + ' workload submissions on file.</p>');
          }
        }
      } catch (e) { Logger.log('Digest workload error: ' + e.message); }
    }

    // Tasks
    if (config.includeTasks && typeof DataService !== 'undefined') {
      try {
        var tasks = DataService.getMemberTasks(email);
        if (tasks && tasks.length > 0) {
          var openTasks = tasks.filter(function (t) { return t.status !== 'completed'; });
          var tHtml = '<h3>Your Assigned Tasks (' + openTasks.length + ' open)</h3><ul>';
          openTasks.forEach(function (t) {
            tHtml += '<li><strong>' + escapeHtml(t.title) + '</strong> — ' + escapeHtml(t.priority) + ' priority' + (t.dueDate ? ', due ' + escapeHtml(t.dueDate) : '') + '</li>';
          });
          tHtml += '</ul>';
          sections.push(tHtml);
        }
      } catch (e) { Logger.log('Digest tasks error: ' + e.message); }
    }

    if (sections.length === 0) return null;

    var html = '<html><body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">';
    html += '<h2 style="color: #1a73e8;">Union Dashboard Digest</h2>';
    html += '<p style="color: #666;">Here is your latest data summary from the Union Dashboard.</p><hr>';
    html += sections.join('<hr>');
    html += '<hr><p style="font-size: 12px; color: #999;">This is an automated digest. Manage your preferences in the Dashboard under Profile &gt; Email Digest Settings.</p>';
    html += '</body></html>';
    return html;
  }

  // ═══════════════════════════════════════
  // Bulk Export
  // ═══════════════════════════════════════

  function triggerBulkExport(stewardEmail) {
    if (!stewardEmail) return { success: false, message: 'Not authorized.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet || memberSheet.getLastRow() <= 1) return { success: false, message: 'No member data.' };

    var members = memberSheet.getDataRange().getValues();
    var headers = members[0];
    var emailCol = -1;
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).toLowerCase().indexOf('email') !== -1) { emailCol = h; break; }
    }
    if (emailCol === -1) return { success: false, message: 'Email column not found.' };

    var sent = 0;
    for (var i = 1; i < members.length; i++) {
      var email = String(members[i][emailCol]).toLowerCase().trim();
      if (!email || email.indexOf('@') === -1) continue;

      var remaining = MailApp.getRemainingDailyQuota();
      if (remaining < 5) break;

      try {
        var config = { includeGrievances: true, includeWorkload: true, includeTasks: true };
        var body = _composeMemberDigest(email, config);
        if (body) {
          MailApp.sendEmail({
            to: email,
            subject: 'Your Personal Union Data Summary',
            htmlBody: body,
            noReply: true
          });
          sent++;
        }
      } catch (e) { Logger.log('Bulk export error for ' + email + ': ' + e.message); }
    }

    logAuditEvent('FAILSAFE_BULK_EXPORT', sent + ' member summaries emailed by ' + stewardEmail);
    return { success: true, sent: sent };
  }

  // ═══════════════════════════════════════
  // Drive Backup
  // ═══════════════════════════════════════

  function backupCriticalSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var folder = _getOrCreateBackupFolder();
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');

    var sheetsToBackup = [
      SHEETS.MEMBER_DIR,
      SHEETS.GRIEVANCE_LOG,
      SHEETS.WORKLOAD_VAULT,
      SHEETS.STEWARD_TASKS,
      SHEETS.CONTACT_LOG
    ];

    // Add v4.17.0 sheets if they exist
    if (SHEETS.QA_FORUM) sheetsToBackup.push(SHEETS.QA_FORUM);
    if (SHEETS.TIMELINE_EVENTS) sheetsToBackup.push(SHEETS.TIMELINE_EVENTS);

    var backedUp = 0;
    for (var i = 0; i < sheetsToBackup.length; i++) {
      var sheetName = sheetsToBackup[i];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) continue;

      try {
        var data = sheet.getDataRange().getValues();
        var csv = _toCsv(data);
        var fileName = sheetName.replace(/[^a-zA-Z0-9_-]/g, '') + '_' + dateStr + '.csv';
        folder.createFile(fileName, csv, MimeType.CSV);
        backedUp++;
      } catch (e) {
        Logger.log('Backup error for ' + sheetName + ': ' + e.message);
      }
    }

    // Prune old backups — keep most recent MAX_BACKUP_FILES per sheet
    _pruneOldBackups(folder, sheetsToBackup);

    logAuditEvent('FAILSAFE_BACKUP_CREATED', backedUp + ' sheets backed up to Drive');
    return { success: true, backedUp: backedUp, folderName: BACKUP_FOLDER_NAME };
  }

  // ═══════════════════════════════════════
  // Trigger Management
  // ═══════════════════════════════════════

  function setupFailsafeTriggers() {
    // Remove existing triggers first
    removeFailsafeTriggers();

    // Daily digest check (runs at 7 AM)
    ScriptApp.newTrigger('fsProcessScheduledDigests')
      .timeBased()
      .atHour(7)
      .everyDays(1)
      .create();

    // Weekly backup (runs Sunday at 2 AM)
    ScriptApp.newTrigger('fsBackupCriticalSheets')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(2)
      .create();

    logAuditEvent('FAILSAFE_TRIGGERS_INSTALLED', 'Daily digest + weekly backup triggers installed');
    return { success: true, message: 'Triggers installed: daily digest check + weekly backup.' };
  }

  function removeFailsafeTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    for (var i = 0; i < triggers.length; i++) {
      var handler = triggers[i].getHandlerFunction();
      if (handler === 'fsProcessScheduledDigests' || handler === 'fsBackupCriticalSheets') {
        ScriptApp.deleteTrigger(triggers[i]);
        removed++;
      }
    }
    if (removed > 0) logAuditEvent('FAILSAFE_TRIGGERS_REMOVED', removed + ' triggers removed');
    return { success: true, removed: removed };
  }

  // ═══════════════════════════════════════
  // Helpers
  // ═══════════════════════════════════════

  function _pruneOldBackups(folder, sheetNames) {
    sheetNames.forEach(function(sheetName) {
      var prefix = sheetName.replace(/[^a-zA-Z0-9_-]/g, '');
      // Collect all files matching this sheet prefix
      var allFiles = [];
      var iter = folder.getFiles();
      while (iter.hasNext()) {
        var f = iter.next();
        var name = f.getName();
        if (name.indexOf(prefix + '_') === 0 && name.slice(-4) === '.csv') {
          allFiles.push({ file: f, name: name });
        }
      }
      // Sort by name descending (name contains yyyy-MM-dd_HHmm so lexical = chronological)
      allFiles.sort(function(a, b) { return b.name < a.name ? -1 : b.name > a.name ? 1 : 0; });
      // Delete everything beyond MAX_BACKUP_FILES
      for (var i = MAX_BACKUP_FILES; i < allFiles.length; i++) {
        try { allFiles[i].file.setTrashed(true); } catch (_e) { /* ignore */ }
      }
    });
  }

  function _getOrCreateBackupFolder() {
    var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(BACKUP_FOLDER_NAME);
  }

  function _toCsv(data) {
    return data.map(function (row) {
      return row.map(function (cell) {
        var val = cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(cell || '');
        if (val.indexOf(',') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
          val = '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      }).join(',');
    }).join('\n');
  }

  function _fmtDate(date) {
    var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[date.getMonth()] + ' ' + date.getDate() + ', ' + date.getFullYear();
  }

  // ═══════════════════════════════════════
  // Public API
  // ═══════════════════════════════════════

  return {
    initFailsafeSheet: initFailsafeSheet,
    getDigestConfig: getDigestConfig,
    updateDigestConfig: updateDigestConfig,
    processScheduledDigests: processScheduledDigests,
    triggerBulkExport: triggerBulkExport,
    backupCriticalSheets: backupCriticalSheets,
    setupFailsafeTriggers: setupFailsafeTriggers,
    removeFailsafeTriggers: removeFailsafeTriggers
  };

})();


// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════

// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════
// Auth model:
//   fsGetDigestConfig / fsUpdateDigestConfig — server-resolved identity only; client email param ignored
//   fsTriggerBulkExport / fsSetupTriggers / fsRemoveTriggers / fsBackupCriticalSheets / fsInitSheets — steward-only
//
// sessionToken: client passes SESSION_TOKEN (from PAGE_DATA.sessionToken) for magic link / session auth.
// Server validates the token via Auth.resolveEmailFromToken() — the raw email is never trusted.

function fsGetDigestConfig(sessionToken) {
  // Server-resolved identity: SSO first, then verified session token
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
  return FailsafeService.getDigestConfig(e);
}

function fsUpdateDigestConfig(sessionToken, config) {
  // Server-resolved identity only — sessionToken verified server-side
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  return FailsafeService.updateDigestConfig(e, config);
}

function fsProcessScheduledDigests() {
  // Scheduled trigger — no interactive auth; runs as script owner
  return FailsafeService.processScheduledDigests();
}

function fsTriggerBulkExport(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.triggerBulkExport(e);
}

function fsBackupCriticalSheets(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.backupCriticalSheets();
}

function fsSetupTriggers(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.setupFailsafeTriggers();
}

function fsRemoveTriggers(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.removeFailsafeTriggers();
}

function fsInitSheets(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.initFailsafeSheet();
}

/**
 * Ensures ALL required hidden/optional sheets exist.
 * Callable from the SPA (no Sheets UI needed).
 * Equivalent to the sheet-creation portions of CREATE_DASHBOARD.
 * Non-destructive: skips sheets that already exist, only creates missing ones.
 */
/**
 * Ensures ALL required hidden/optional sheets exist.
 * Callable from the SPA (steward auth required).
 * Delegates to _ensureAllSheetsInternal() — single source of truth.
 */
function fsEnsureAllSheets(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };

  var result = _ensureAllSheetsInternal();
  return {
    success: true,
    created: result.created,
    failed: result.failed,
    message: 'Initialized ' + result.created.length + ' sheet group(s).' + (result.failed.length > 0 ? ' ' + result.failed.length + ' failed.' : '')
  };
}

/**
 * Diagnostic function — checks which sheets exist and which are missing.
 * Helps identify why tabs fail with "Something went wrong" errors.
 * Call from SPA Failsafe page to see system health.
 */
function fsDiagnostic(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: 'Spreadsheet binding broken.' };

  var allSheetNames = ss.getSheets().map(function(s) { return s.getName(); });
  var results = [];

  // Check critical sheets
  var required = {
    'Member Directory': SHEETS.MEMBER_DIR,
    'Grievance Log': SHEETS.GRIEVANCE_LOG,
    'Config': SHEETS.CONFIG,
  };
  var optional = {
    'Notifications': SHEETS.NOTIFICATIONS,
    'Resources': SHEETS.RESOURCES,
    'Resource Config': SHEETS.RESOURCE_CONFIG,
    'Contact Log': SHEETS.CONTACT_LOG,
    'Steward Tasks': SHEETS.STEWARD_TASKS,
    'Timeline Events': SHEETS.TIMELINE_EVENTS,
    'QA Forum': SHEETS.QA_FORUM,
    'QA Answers': SHEETS.QA_ANSWERS,
    'Feedback': SHEETS.FEEDBACK,
    'Case Checklist': SHEETS.CASE_CHECKLIST,
    'Weekly Questions': SHEETS.WEEKLY_QUESTIONS,
    'Weekly Responses': SHEETS.WEEKLY_RESPONSES,
    'Question Pool': SHEETS.QUESTION_POOL,
    'Survey Tracking': SHEETS.SURVEY_TRACKING,
    'Survey Vault': SHEETS.SURVEY_VAULT,
    'Member Satisfaction': SHEETS.SATISFACTION,
    'Survey Questions': SHEETS.SURVEY_QUESTIONS,
    'Failsafe Config': SHEETS.FAILSAFE_CONFIG,
    'Workload Vault': SHEETS.WORKLOAD_VAULT,
    'Workload Reporting': SHEETS.WORKLOAD_REPORTING,
    'Audit Log': SHEETS.AUDIT_LOG,
    'Meeting Check-In Log': SHEETS.MEETING_CHECKIN_LOG,
    'Steward Performance': SHEETS.STEWARD_PERFORMANCE_CALC,
  };

  for (var label in required) {
    var sheetName = required[label];
    results.push({
      label: label,
      sheet: sheetName,
      exists: allSheetNames.indexOf(sheetName) !== -1,
      critical: true
    });
  }
  for (var label2 in optional) {
    var sheetName2 = optional[label2];
    results.push({
      label: label2,
      sheet: sheetName2,
      exists: allSheetNames.indexOf(sheetName2) !== -1,
      critical: false
    });
  }

  // Also check auth
  var authOk = false;
  try {
    var testAuth = checkWebAppAuthorization('steward', sessionToken);
    authOk = testAuth && testAuth.isAuthorized;
  } catch (_) {}

  return {
    success: true,
    totalSheets: allSheetNames.length,
    results: results,
    authOk: authOk,
    email: e,
    missing: results.filter(function(r) { return !r.exists; }).map(function(r) { return r.label + ' (' + r.sheet + ')' + (r.critical ? ' ⚠️ CRITICAL' : ''); })
  };
}
