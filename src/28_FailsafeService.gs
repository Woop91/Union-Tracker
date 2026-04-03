/**
 * ============================================================================
 * 28_FailsafeService.gs - Data Failsafe / Recovery System
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   Data resilience system with scheduled email digests and Google Drive CSV
 *   backups. Members can opt into periodic email digests of their grievance/
 *   workload/task data. Automatic weekly Drive backups export key sheets as
 *   CSV files to a SolidBase_Dashboard_Backups folder. Maintains maximum 52 backup
 *   files (~1 year of weekly backups per sheet).
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Failsafe protects against data loss from accidental sheet deletion,
 *   corruption, or access revocation. Email digests give members a personal
 *   copy of their data outside of Google Sheets. Drive CSV backups are simple
 *   and universal — any spreadsheet tool can read them. The 52-file cap
 *   prevents Drive storage bloat while maintaining a full year of recovery
 *   points. Sheet reads are cached (120s TTL) to avoid redundant
 *   getDataRange() calls.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Email digests stop sending — members lose their periodic data summaries.
 *   Drive backups stop — the last backup becomes the most recent recovery
 *   point. If the sheet is corrupted after the failsafe stops, there's no
 *   automated recovery. Existing backups in Drive are unaffected.
 *
 * DEPENDENCIES:
 *   Depends on 01_Core.gs (SHEETS), DriveApp, MailApp (GAS built-ins),
 *   CacheService. Used by weekly backup trigger and member digest
 *   preferences in the SPA.
 *
 * @version 4.43.1
 */

var FailsafeService = (function () {

  var BACKUP_FOLDER_NAME = 'SolidBase_Dashboard_Backups';
  var MAX_BACKUP_FILES = 52; // ~1 year of weekly backups per sheet

  // ═══════════════════════════════════════
  // Sheet Setup
  // ═══════════════════════════════════════

  /**
   * Creates the _Failsafe_Config hidden sheet if it does not exist.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The failsafe config sheet.
   */
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
      // GAS-02: Use very-hidden so users cannot unhide PII-containing system sheets via menu
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(sheet); else sheet.hideSheet();
    }
    return sheet;
  }

  // ═══════════════════════════════════════
  // Cached Sheet Read Helper
  // ═══════════════════════════════════════

  /**
   * Returns cached sheet data from CacheService, falling back to a live read.
   * @param {string} sheetName - Name of the sheet to read.
   * @param {number} [maxAgeSec=120] - Cache TTL in seconds.
   * @returns {Array[]|null} 2D array of sheet values, or null if empty/missing.
   */
  function _getCachedSheetData(sheetName, maxAgeSec) {
    maxAgeSec = maxAgeSec || 120;
    try {
      var cache = CacheService.getScriptCache();
      var cacheKey = 'fs_sheet_' + sheetName;
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (_e) { log_('_e', (_e.message || _e)); }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var data = sheet.getDataRange().getValues();
    try {
      cache.put(cacheKey, JSON.stringify(data), maxAgeSec);
    } catch (_e) { log_('_e', (_e.message || _e)); }
    return data;
  }

  // ═══════════════════════════════════════
  // Digest Configuration
  // ═══════════════════════════════════════

  /**
   * Retrieves a member's email digest preferences, returning defaults if not configured.
   * @param {string} email - Member's email address.
   * @returns {Object} Digest config with enabled, frequency, and include* flags.
   */
  function getDigestConfig(email) {
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
    }

    var data = _getCachedSheetData(SHEETS.FAILSAFE_CONFIG, 120) || sheet.getDataRange().getValues();
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

  /**
   * Creates or updates a member's email digest preferences under a script lock.
   * @param {string} email - Member's email address.
   * @param {Object} config - Digest settings (enabled, frequency, include* flags).
   * @returns {{success: boolean, message: string}}
   */
  function updateDigestConfig(email, config) {
    if (!email || !config) return { success: false, message: 'Missing data.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG);
    if (!sheet) { initFailsafeSheet(); sheet = ss.getSheetByName(SHEETS.FAILSAFE_CONFIG); }

    var eml = email.toLowerCase().trim();
    var frequency = config.frequency === 'monthly' ? 'monthly' : 'weekly';

    // M10: Acquire lock BEFORE reading sheet to prevent stale-read race
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return { success: false, message: 'Could not acquire lock. Please try again.' };
    }
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

  /**
   * Iterates enabled digest configs and sends due email digests with double-checked locking.
   * @returns {{processed: number}} Count of digests sent.
   */
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
        log_('processScheduledDigests', 'Email quota low (' + remaining + '), stopping digest processing.');
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
            log_('processScheduledDigests', 'Could not acquire send lock for ' + email + ', skipping to avoid duplicate.');
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
        log_('processScheduledDigests', 'Digest send error for ' + email + ': ' + err.message);
      }
    }

    logAuditEvent('FAILSAFE_DIGESTS_PROCESSED', processed + ' digests sent');
    return { processed: processed };
  }

  /**
   * Composes an HTML email body summarizing a member's grievances, workload, and tasks.
   * @param {string} email - Member's email address.
   * @param {Object} config - Include flags (includeGrievances, includeWorkload, includeTasks).
   * @returns {string|null} HTML string, or null if no data sections were generated.
   */
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
      } catch (e) { log_('Digest grievance error', e.message); }
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
      } catch (e) { log_('Digest workload error', e.message); }
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
      } catch (e) { log_('Digest tasks error', e.message); }
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
  // Drive Backup
  // ═══════════════════════════════════════

  /**
   * Exports critical sheets as CSV files to the Drive backup folder and prunes old backups.
   * @returns {{success: boolean, backedUp?: number, folderName?: string, message?: string}}
   */
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
        log_('backupCriticalSheets', 'Backup error for ' + sheetName + ': ' + e.message);
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

  /**
   * Installs daily digest (7 AM) and weekly backup (Sunday 2 AM) time-based triggers.
   * @returns {{success: boolean, message: string}}
   */
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

  /**
   * Deletes all project triggers whose handler is a failsafe function.
   * @returns {{success: boolean, removed: number}}
   */
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

  /**
   * Trashes backup CSV files beyond MAX_BACKUP_FILES per sheet prefix.
   * @param {GoogleAppsScript.Drive.Folder} folder - Backup folder to prune.
   * @param {string[]} sheetNames - Sheet names whose backups to prune.
   * @returns {void}
   */
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
        try { allFiles[i].file.setTrashed(true); } catch (_e) { log_('_e', (_e.message || _e)); }
      }
    });
  }

  /**
   * Returns the Drive backup folder, creating it with private sharing if it does not exist.
   * @returns {GoogleAppsScript.Drive.Folder} The backup folder.
   */
  function _getOrCreateBackupFolder() {
    var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
    if (folders.hasNext()) return folders.next();
    // SEC-07: Restrict sharing so only the script owner can access PII-containing backups
    var folder = executeWithRetry(function() { return DriveApp.createFolder(BACKUP_FOLDER_NAME); }, { maxRetries: 2, baseDelay: 500 });
    try {
      folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      folder.setDescription('Automated backups — contains member PII. Do not share.');
    } catch (_e) {
      log_('Could not restrict backup folder sharing', _e.message);
    }
    return folder;
  }

  /**
   * Converts a 2D array to a CSV string with proper quoting and date formatting.
   * @param {Array[]} data - 2D array of cell values.
   * @returns {string} CSV-formatted string.
   */
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

  /**
   * Formats a Date as a short string by delegating to fmtDateShort_ in 01_Core.gs.
   * @param {Date} date - Date to format.
   * @returns {string} Formatted date string.
   */
  function _fmtDate(date) { return fmtDateShort_(date); }

  // ═══════════════════════════════════════
  // Drive Backup Restore
  // ═══════════════════════════════════════

  /**
   * Restores sheet data from a Drive CSV backup file.
   * Creates a pre-restore backup before overwriting.
   * @param {string} fileId - Google Drive file ID of the CSV backup
   * @param {string} sheetName - Target sheet name to restore into
   * @param {boolean} confirmed - Must be true to proceed (safety gate)
   * @returns {Object} { success, message, rowsRestored }
   */
  function restoreFromDriveBackup(fileId, sheetName, confirmed) {
    if (!confirmed) return { success: false, message: 'Restore requires explicit confirmation (confirmed=true).' };
    if (!fileId || !sheetName) return { success: false, message: 'fileId and sheetName are required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet "' + sheetName + '" not found.' };

    // Read CSV from Drive
    var file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (e) {
      return { success: false, message: 'Could not access file: ' + e.message };
    }
    var csv = file.getBlob().getDataAsString();
    var rows = Utilities.parseCsv(csv);
    if (!rows || rows.length === 0) return { success: false, message: 'CSV file is empty.' };

    // Validate column count matches
    var currentCols = sheet.getLastColumn();
    if (currentCols > 0 && rows[0].length !== currentCols) {
      return { success: false, message: 'Column mismatch: backup has ' + rows[0].length + ' columns, sheet has ' + currentCols + '. Restore aborted.' };
    }

    // Create pre-restore backup (very-hidden sheet)
    var backupName = '_PreRestore_' + sheetName.replace(/\s/g, '_') + '_' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss');
    try {
      var backup = sheet.copyTo(ss).setName(backupName);
      if (typeof setSheetVeryHidden_ === 'function') setSheetVeryHidden_(backup); else backup.hideSheet();
    } catch (backupErr) {
      return { success: false, message: 'Could not create pre-restore backup: ' + backupErr.message };
    }

    // Restore data
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return { success: false, message: 'Could not acquire lock for restore.' };
    var restoreErr = null;
    try {
      sheet.clearContents();
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    } catch (e) {
      restoreErr = e;
    } finally {
      lock.releaseLock();
    }
    if (restoreErr) {
      return { success: false, message: 'Restore write failed: ' + restoreErr.message + '. Pre-restore backup saved as "' + backupName + '".' };
    }

    logAuditEvent('FAILSAFE_RESTORE', 'Restored ' + sheetName + ' from Drive file ' + fileId + ' (' + rows.length + ' rows). Pre-restore backup: ' + backupName);
    return { success: true, message: 'Restored ' + rows.length + ' rows to "' + sheetName + '". Pre-restore backup: "' + backupName + '".', rowsRestored: rows.length };
  }

  // ═══════════════════════════════════════
  // Public API
  // ═══════════════════════════════════════

  return {
    initFailsafeSheet: initFailsafeSheet,
    getDigestConfig: getDigestConfig,
    updateDigestConfig: updateDigestConfig,
    processScheduledDigests: processScheduledDigests,
    backupCriticalSheets: backupCriticalSheets,
    setupFailsafeTriggers: setupFailsafeTriggers,
    removeFailsafeTriggers: removeFailsafeTriggers,
    restoreFromDriveBackup: restoreFromDriveBackup
  };

})();
// ═══════════════════════════════════════
// GLOBAL WRAPPERS (callable from client via google.script.run)
// ═══════════════════════════════════════
// Auth model:
//   fsGetDigestConfig / fsUpdateDigestConfig — server-resolved identity only; client email param ignored
//   fsSetupTriggers / fsRemoveTriggers / fsBackupCriticalSheets / fsInitSheets — steward-only
//
// sessionToken: client passes SESSION_TOKEN (from PAGE_DATA.sessionToken) for magic link / session auth.
// Server validates the token via Auth.resolveEmailFromToken() — the raw email is never trusted.

/**
 * Global wrapper: returns the caller's email digest configuration.
 * @param {string} sessionToken - Session token for server-resolved identity.
 * @returns {Object} Digest config object.
 */
function fsGetDigestConfig(sessionToken) {
  // Server-resolved identity: SSO first, then verified session token
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { enabled: false, frequency: 'weekly', includeGrievances: true, includeWorkload: true, includeTasks: true };
  return FailsafeService.getDigestConfig(e);
}

/**
 * Global wrapper: updates the caller's email digest preferences.
 * @param {string} sessionToken - Session token for server-resolved identity.
 * @param {Object} config - Digest settings to save.
 * @returns {{success: boolean, message: string}}
 */
function fsUpdateDigestConfig(sessionToken, config) {
  // Server-resolved identity only — sessionToken verified server-side
  var e = _resolveCallerEmail(sessionToken);
  if (!e) return { success: false, message: 'Not authenticated.' };
  return FailsafeService.updateDigestConfig(e, config);
}

/**
 * Global wrapper: processes scheduled digests (trigger-invoked, no interactive auth).
 * @returns {{processed: number}}
 */
function fsProcessScheduledDigests() {
  // Scheduled trigger — no interactive auth; runs as script owner
  return FailsafeService.processScheduledDigests();
}

/**
 * Global wrapper: backs up critical sheets to Drive as CSV (steward-only).
 * @param {string} sessionToken - Session token for steward auth.
 * @returns {{success: boolean, backedUp?: number, message?: string}}
 */
function fsBackupCriticalSheets(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.backupCriticalSheets();
}

/**
 * Global wrapper: installs failsafe time-based triggers (steward-only).
 * @param {string} sessionToken - Session token for steward auth.
 * @returns {{success: boolean, message: string}}
 */
function fsSetupTriggers(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.setupFailsafeTriggers();
}

/**
 * Global wrapper: removes failsafe time-based triggers (steward-only).
 * @param {string} sessionToken - Session token for steward auth.
 * @returns {{success: boolean, removed: number}}
 */
function fsRemoveTriggers(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.removeFailsafeTriggers();
}

/**
 * Global wrapper: initializes the failsafe config sheet (steward-only).
 * @param {string} sessionToken - Session token for steward auth.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function fsInitSheets(sessionToken) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.initFailsafeSheet();
}

/**
 * Global wrapper: restores a sheet from a Drive CSV backup (steward-only).
 * @param {string} sessionToken - Session token for steward auth.
 * @param {string} fileId - Drive file ID of the CSV backup.
 * @param {string} sheetName - Target sheet name to restore into.
 * @param {boolean} confirmed - Must be true to proceed.
 * @returns {{success: boolean, message: string, rowsRestored?: number}}
 */
function fsRestoreFromBackup(sessionToken, fileId, sheetName, confirmed) {
  var e = _requireStewardAuth(sessionToken);
  if (!e) return { success: false, message: 'Not authorized.' };
  return FailsafeService.restoreFromDriveBackup(fileId, sheetName, confirmed);
}

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
    'Workload Archive': SHEETS.WORKLOAD_ARCHIVE,
    'Workload Reminders': SHEETS.WORKLOAD_REMINDERS,
    'Workload UserMeta': SHEETS.WORKLOAD_USERMETA,
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
  } catch (_) { log_('_', (_.message || _)); }

  return {
    success: true,
    totalSheets: allSheetNames.length,
    results: results,
    authOk: authOk,
    email: e,
    missing: results.filter(function(r) { return !r.exists; }).map(function(r) { return r.label + ' (' + r.sheet + ')' + (r.critical ? ' ⚠️ CRITICAL' : ''); })
  };
}
