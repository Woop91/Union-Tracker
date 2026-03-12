/**
 * ============================================================================
 * 05_Integrations.gs - External Service Integration
 * ============================================================================
 *
 * This module handles all interactions with external Google services:
 * - Google Drive folder management for grievance documents
 * - Google Calendar deadline synchronization
 * - Email notifications
 * - External API calls
 *
 * SEPARATION OF CONCERNS: Isolating external dependencies ensures that if
 * one service (e.g., Drive) has an outage, core spreadsheet functionality
 * remains responsive.
 *
 * @fileoverview External service integrations
 * @version 4.7.0
 * @requires 01_Core.gs
 */

// ============================================================================
// CALENDAR CONFIGURATION
// ============================================================================

/**
 * Calendar configuration for deadline tracking
 * @const {Object}
 */
var CALENDAR_CONFIG = {
  CALENDAR_NAME: 'Grievance Deadlines',
  MEETING_CALENDAR_NAME: 'Union Meetings',
  REMINDER_DAYS: [7, 3, 1],
  MEETING_DEACTIVATE_HOURS: 3  // Hours after meeting end to deactivate check-in
};

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

// ─── DRIVE FOLDER SETUP (v4.20.17) ─────────────────────────────────────────
/**
 * Creates the full DashboardTest Drive folder hierarchy during CREATE_DASHBOARD.
 *
 * Folder structure:
 *   DashboardTest/               ← PRIVATE — only explicitly shared users can access
 *     ├── Grievances/            ← individual case folders go here
 *     ├── Resources/
 *     ├── Minutes/
 *     └── Event Check-In/
 *
 * All folder IDs are written back to Config sheet row 3 so the system
 * can reference them dynamically at runtime.
 *
 * Safe to re-run — existing folders are found by stored ID first, then by
 * name search, then created fresh.
 *
 * @returns {Object} { success, rootFolderId, rootFolderUrl, grievancesFolderId,
 *                     resourcesFolderId, minutesFolderId, eventCheckinFolderId }
 */
function setupDashboardDriveFolders() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  // ── 1. Get or create DashboardTest root ─────────────────────────────────
  var rootFolder = _getOrCreateNamedFolder_(
    getDriveRootFolderName_(),
    'DASHBOARD_ROOT_FOLDER_ID',
    props,
    configSheet,
    null // no parent = root of My Drive
  );
  if (!rootFolder) throw new Error('Could not create or access the root Dashboard folder');

  // Lock it down: PRIVATE — nobody can discover it outside explicit sharing
  try {
    rootFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (shareErr) {
    Logger.log('setupDashboardDriveFolders: could not set root folder to PRIVATE: ' + shareErr.message);
  }

  // ── 2. Create required subfolders inside DashboardTest ───────────────────
  var subDefs = [
    { name: DRIVE_CONFIG.GRIEVANCES_SUBFOLDER,    propKey: 'GRIEVANCE_ROOT_FOLDER_ID',    cfgKey: 'GRIEVANCES_FOLDER_ID' },
    { name: DRIVE_CONFIG.RESOURCES_SUBFOLDER,     propKey: 'RESOURCES_FOLDER_ID',         cfgKey: 'RESOURCES_FOLDER_ID' },
    { name: DRIVE_CONFIG.MINUTES_SUBFOLDER,       propKey: 'MINUTES_FOLDER_ID',           cfgKey: 'MINUTES_FOLDER_ID' },
    { name: DRIVE_CONFIG.EVENT_CHECKIN_SUBFOLDER, propKey: 'EVENT_CHECKIN_FOLDER_ID',     cfgKey: 'EVENT_CHECKIN_FOLDER_ID' },
    { name: DRIVE_CONFIG.MEMBERS_SUBFOLDER,       propKey: 'MEMBERS_FOLDER_ID',           cfgKey: 'MEMBERS_FOLDER_ID' },
  ];

  var result = {
    success: true,
    rootFolderId:        rootFolder.getId(),
    rootFolderUrl:       rootFolder.getUrl(),
    rootFolderName:      getDriveRootFolderName_(),
  };

  subDefs.forEach(function(def) {
    var sub = _getOrCreateNamedFolder_(def.name, def.propKey, props, configSheet, rootFolder);
    if (!sub) { Logger.log('Could not create subfolder: ' + def.name); return; }
    result[def.propKey] = sub.getId();
    // Minutes/ folder: anyone with link can view (members use the link to browse docs)
    // All other subfolders inherit the root PRIVATE setting.
    if (def.name === DRIVE_CONFIG.MINUTES_SUBFOLDER) {
      try {
        sub.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Logger.log('Minutes folder set to view-only public link.');
      } catch (mShareErr) {
        Logger.log('Could not set Minutes folder sharing: ' + mShareErr.message);
      }
    }
    // Also write to Config sheet using the config-specific key
    _writeConfigFolderId_(configSheet, CONFIG_COLS[def.cfgKey], sub.getId(), sub.getUrl());
    Logger.log('Drive folder ready: ' + def.name + ' (' + sub.getId() + ')');
  });

  // Write root folder IDs to Config
  _writeConfigFolderId_(configSheet, CONFIG_COLS.DASHBOARD_ROOT_FOLDER_ID, rootFolder.getId(), rootFolder.getUrl());
  _writeConfigFolderId_(configSheet, CONFIG_COLS.DRIVE_FOLDER_ID, rootFolder.getId(), rootFolder.getUrl());

  return result;
}

/**
 * Gets a folder by stored Script Property ID, falling back to a name search
 * within the given parent, then creating it fresh.
 * @private
 */
function _getOrCreateNamedFolder_(name, propKey, props, configSheet, parentFolder) {
  // Try stored ID first
  var storedId = props.getProperty(propKey);
  if (storedId) {
    try {
      var found = DriveApp.getFolderById(storedId);
      if (found) return found;
    } catch (_e) { /* stale ID */ }
  }

  // Try name search inside parent (or all of Drive)
  var iter = parentFolder ? parentFolder.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (iter.hasNext()) {
    var existing = iter.next();
    props.setProperty(propKey, existing.getId());
    return existing;
  }

  // Create fresh
  var newFolder = parentFolder ? parentFolder.createFolder(name) : DriveApp.createFolder(name);
  props.setProperty(propKey, newFolder.getId());

  newFolder.setDescription('Union Dashboard — auto-managed. Do not move or rename.');
  return newFolder;
}

/**
 * Writes a folder ID (and optionally URL) to a Config sheet column in row 3.
 * Safe — skips if the column is 0/undefined (config key not mapped).
 * @private
 */
function _writeConfigFolderId_(configSheet, colIndex, folderId, folderUrl) {
  if (!configSheet || !colIndex || colIndex === 0) return;
  try {
    configSheet.getRange(3, colIndex).setValue(folderId || '');
  } catch (e) {
    Logger.log('_writeConfigFolderId_: could not write col ' + colIndex + ': ' + e.message);
  }
}

// ─── CALENDAR SETUP (v4.20.17) ──────────────────────────────────────────────
/**
 * Creates the union events calendar and writes its ID to Config.
 * Safe to re-run — finds existing calendar by stored ID first, then by name.
 *
 * @returns {Object} { success, calendarId, calendarName }
 */
function setupDashboardCalendar() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  // Derive calendar name from org name
  var orgName = getSystemName_();
  var calendarName = (orgName && orgName !== 'Union Dashboard') ? orgName + ' Events' : 'Union Events';

  // Try stored ID
  var storedId = props.getProperty('UNION_CALENDAR_ID');
  if (storedId) {
    try {
      var existingCal = CalendarApp.getCalendarById(storedId);
      if (existingCal) {
        _writeCalendarIdToConfig_(configSheet, storedId);
        return { success: true, calendarId: storedId, calendarName: existingCal.getName() };
      }
    } catch (_e) { /* stale ID */ }
  }

  // Search by name
  var byName = CalendarApp.getCalendarsByName(calendarName);
  if (byName.length > 0) {
    var cal = byName[0];
    var calId = cal.getId();
    props.setProperty('UNION_CALENDAR_ID', calId);
    _writeCalendarIdToConfig_(configSheet, calId);
    return { success: true, calendarId: calId, calendarName: calendarName };
  }

  // Create fresh
  var newCal = CalendarApp.createCalendar(calendarName, {
    summary: 'Union events — Auto-managed by Dashboard',
    color: CalendarApp.Color.BLUE,
    timeZone: Session.getScriptTimeZone()
  });

  var newId = newCal.getId();
  props.setProperty('UNION_CALENDAR_ID', newId);
  _writeCalendarIdToConfig_(configSheet, newId);

  Logger.log('Created union events calendar: ' + calendarName + ' (' + newId + ')');
  return { success: true, calendarId: newId, calendarName: calendarName };
}

/**
 * Writes calendar ID to Config sheet row 3, CONFIG_COLS.CALENDAR_ID.
 * @private
 */
function _writeCalendarIdToConfig_(configSheet, calendarId) {
  if (!configSheet || !CONFIG_COLS.CALENDAR_ID) return;
  try {
    configSheet.getRange(3, CONFIG_COLS.CALENDAR_ID).setValue(calendarId || '');
  } catch (e) {
    Logger.log('_writeCalendarIdToConfig_: ' + e.message);
  }
}

// ─── ROOT FOLDER ACCESS ──────────────────────────────────────────────────────
/**
 * Gets or creates the Grievances subfolder inside DashboardTest.
 * This is used by setupDriveFolderForGrievance() to place case folders
 * inside DashboardTest/Grievances/ rather than in a standalone root.
 * @return {Folder} The Grievances folder
 */
function getOrCreateRootFolder() {
  // Grievance case folders live in DashboardTest/Grievances/.
  // Try stored Grievances subfolder ID first (fastest path, set by setupDashboardDriveFolders).
  var props = PropertiesService.getScriptProperties();
  var storedId = props.getProperty('GRIEVANCE_ROOT_FOLDER_ID');
  if (storedId) {
    try {
      return DriveApp.getFolderById(storedId);
    } catch (_e) { /* stale — fall through */ }
  }

  // Find or create the DashboardTest root, then the Grievances subfolder inside it.
  var rootFolder = _getOrCreateNamedFolder_(
    getDriveRootFolderName_(),
    'DASHBOARD_ROOT_FOLDER_ID',
    props,
    null, // configSheet — not needed here
    null  // no parent = My Drive root
  );

  var grievancesFolder = _getOrCreateNamedFolder_(
    DRIVE_CONFIG.GRIEVANCES_SUBFOLDER,
    'GRIEVANCE_ROOT_FOLDER_ID',
    props,
    null,
    rootFolder
  );

  // Lock root if it's new (best-effort)
  try {
    rootFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (_e) {}

  grievancesFolder.setDescription('Individual case folders — Auto-managed by Union Dashboard');

  logAuditEvent(AUDIT_EVENTS.FOLDER_CREATED, {
    folderId:  grievancesFolder.getId(),
    folderName: DRIVE_CONFIG.GRIEVANCES_SUBFOLDER,
    type: 'GRIEVANCES_ROOT',
    createdBy: Session.getActiveUser().getEmail()
  });

  return grievancesFolder;
}

/**
 * Sets up a Drive folder for a specific grievance
 * Folder naming format: LastName, FirstName - YYYY-MM-DD
 * @param {string} grievanceId - The grievance ID
 * @return {Object} Result with folder URL or error
 */
// ============================================================================
// PER-MEMBER ADMIN FOLDER
// ============================================================================

/**
 * Gets or creates the per-member master admin folder under [Root]/Members/.
 * Structure: [Dashboard Root]/Members/LastName, FirstName/
 *              ├── Contact Log — LastName, FirstName  (Sheet — added by DataService)
 *              └── Grievances/                         (subfolder for case folders)
 *
 * Sharing model:
 *   - Members/ parent      → steward-only (PRIVATE, never shared)
 *   - Per-member folder    → steward-only (PRIVATE, never shared with member directly)
 *   - Individual grievance case folders get member editor access (see setupDriveFolderForGrievance)
 *
 * Writes the master folder URL to Member Directory "Member Admin Folder URL" column (non-fatal).
 * Caches MEMBERS_FOLDER_ID in Script Properties.
 *
 * @param {string} memberEmail - The member's email address
 * @returns {{ masterFolder: Folder, grievancesFolder: Folder }|null}
 */
function getOrCreateMemberAdminFolder(memberEmail) {
  if (!memberEmail) return null;
  memberEmail = String(memberEmail).trim().toLowerCase();

  try {
    var props = PropertiesService.getScriptProperties();

    // ── 1. Get or locate the Members/ parent folder ─────────────────────────
    var membersRoot = null;
    var storedMembersId = props.getProperty('MEMBERS_FOLDER_ID') || '';
    if (storedMembersId) {
      try { membersRoot = DriveApp.getFolderById(storedMembersId); } catch (_e) { membersRoot = null; }
    }
    if (!membersRoot) {
      // Resolve via dashboard root
      var dashRootId = props.getProperty('DASHBOARD_ROOT_FOLDER_ID') || '';
      if (dashRootId) {
        try {
          var dashRoot = DriveApp.getFolderById(dashRootId);
          var mIter = dashRoot.getFoldersByName(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
          membersRoot = mIter.hasNext() ? mIter.next() : dashRoot.createFolder(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
        } catch (_e) { membersRoot = null; }
      }
      if (!membersRoot) {
        var fallbackIter = DriveApp.getFoldersByName(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
        membersRoot = fallbackIter.hasNext() ? fallbackIter.next() : DriveApp.createFolder(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
      }
      props.setProperty('MEMBERS_FOLDER_ID', membersRoot.getId());
    }

    // ── 2. Resolve member display name from Member Directory ─────────────────
    var memberFolderName = '';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet) {
      var mData = mSheet.getDataRange().getValues();
      var mHeaders = mData[0];
      var mEmailIdx = -1, mFirstIdx = -1, mLastIdx = -1;
      for (var mh = 0; mh < mHeaders.length; mh++) {
        var mhl = String(mHeaders[mh]).toLowerCase().trim();
        if (mhl === 'email')      mEmailIdx = mh;
        else if (mhl === 'first name') mFirstIdx = mh;
        else if (mhl === 'last name')  mLastIdx  = mh;
      }
      if (mEmailIdx !== -1) {
        for (var mr = 1; mr < mData.length; mr++) {
          if (String(mData[mr][mEmailIdx]).toLowerCase().trim() === memberEmail) {
            var fn = mFirstIdx !== -1 ? String(mData[mr][mFirstIdx] || '').trim() : '';
            var ln = mLastIdx  !== -1 ? String(mData[mr][mLastIdx]  || '').trim() : '';
            if (ln && fn) memberFolderName = sanitizeFolderName(ln) + ', ' + sanitizeFolderName(fn);
            else if (fn)  memberFolderName = sanitizeFolderName(fn);
            break;
          }
        }
      }
    }
    // Fallback to email prefix
    if (!memberFolderName) {
      memberFolderName = sanitizeFolderName(memberEmail.split('@')[0].replace(/[._]/g, ' ').trim() || memberEmail);
    }

    // ── 3. Find or create per-member master folder ───────────────────────────
    var masterFolder = null;
    var isNewMaster   = false;
    var existingIter  = membersRoot.getFoldersByName(memberFolderName);
    if (existingIter.hasNext()) {
      masterFolder = existingIter.next();
    } else {
      masterFolder = membersRoot.createFolder(memberFolderName);
      isNewMaster  = true;
    }

    // ── 4. Find or create Grievances/ subfolder ──────────────────────────────
    var grievancesFolder = null;
    var gIter = masterFolder.getFoldersByName('Grievances');
    if (gIter.hasNext()) {
      grievancesFolder = gIter.next();
    } else {
      grievancesFolder = masterFolder.createFolder('Grievances');
    }

    // ── 5. Write master folder URL to Member Directory (non-fatal) ───────────
    if (isNewMaster) {
      try {
        var folderUrl = masterFolder.getUrl();
        if (mSheet) {
          // Reload fresh (mData may be stale after loop above)
          var mData2    = mSheet.getDataRange().getValues();
          var mHeaders2 = mData2[0];
          var mEmailIdx2 = -1, mAdminIdx2 = -1;
          for (var h2 = 0; h2 < mHeaders2.length; h2++) {
            var hl2 = String(mHeaders2[h2]).toLowerCase().trim();
            if (hl2 === 'email')                  mEmailIdx2 = h2;
            else if (hl2 === 'member admin folder url') mAdminIdx2 = h2;
          }
          if (mEmailIdx2 !== -1 && mAdminIdx2 !== -1) {
            for (var r2 = 1; r2 < mData2.length; r2++) {
              if (String(mData2[r2][mEmailIdx2]).toLowerCase().trim() === memberEmail) {
                mSheet.getRange(r2 + 1, mAdminIdx2 + 1).setValue(folderUrl);
                break;
              }
            }
          }
        }
      } catch (writeErr) {
        Logger.log('getOrCreateMemberAdminFolder URL writeback error: ' + writeErr.message);
      }
    }

    return { masterFolder: masterFolder, grievancesFolder: grievancesFolder };

  } catch (err) {
    Logger.log('getOrCreateMemberAdminFolder error for ' + memberEmail + ': ' + err.message);
    return null;
  }
}

/**
 * Creates or retrieves a Google Drive folder for a grievance case.
 * Folder is nested under [Root]/Members/LastName, FirstName/Grievances/{id} - {date}/
 * Member receives EDITOR access on the case folder so they can upload evidence.
 * Steward receives EDITOR access (in addition to inherited steward team access).
 *
 * @param {string} grievanceId - The grievance ID
 * @return {Object} { success, folderId, folderUrl, message } or errorResponse
 */
function setupDriveFolderForGrievance(grievanceId) {
  try {
    const grievance = getGrievanceById(grievanceId);
    if (!grievance) {
      return errorResponse('Grievance not found', 'setupDriveFolderForGrievance');
    }

    const firstName   = grievance['First Name'] || grievance.firstName  || '';
    const lastName    = grievance['Last Name']  || grievance.lastName   || '';
    const memberEmail = String(grievance['Member Email'] || grievance.memberEmail || '').trim().toLowerCase();
    const dateFiled   = grievance['Date Filed'] || grievance.dateFiled  || new Date();

    const dateStr = Utilities.formatDate(
      new Date(dateFiled),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    // Case folder name: {grievanceId} - {date}   (parent already scoped to the member)
    const caseFolderName = grievanceId + ' - ' + dateStr;

    // ── Resolve member admin folder ──────────────────────────────────────────
    // Try by email first; fall back to name-only lookup if no email
    var folderParent = null;  // the Grievances/ subfolder to create under
    if (memberEmail) {
      var adminResult = getOrCreateMemberAdminFolder(memberEmail);
      if (adminResult) folderParent = adminResult.grievancesFolder;
    }
    if (!folderParent) {
      // Fallback: find/create member folder by name under Members/
      var props        = PropertiesService.getScriptProperties();
      var membersRootId = props.getProperty('MEMBERS_FOLDER_ID') || '';
      var membersRoot   = null;
      if (membersRootId) {
        try { membersRoot = DriveApp.getFolderById(membersRootId); } catch (_e) {}
      }
      if (!membersRoot) membersRoot = getOrCreateRootFolder(); // ultimate fallback to dashboard root
      var memberFolderName = (lastName && firstName)
        ? (sanitizeFolderName(lastName) + ', ' + sanitizeFolderName(firstName))
        : (sanitizeFolderName(firstName || lastName || 'Unknown'));
      var mfIter  = membersRoot.getFoldersByName(memberFolderName);
      var mFolder = mfIter.hasNext() ? mfIter.next() : membersRoot.createFolder(memberFolderName);
      var gfIter  = mFolder.getFoldersByName('Grievances');
      folderParent = gfIter.hasNext() ? gfIter.next() : mFolder.createFolder('Grievances');
    }

    // ── Check for existing case folder ───────────────────────────────────────
    var existingIter = folderParent.getFoldersByName(caseFolderName);
    if (existingIter.hasNext()) {
      var existing = existingIter.next();
      return { success: true, folderId: existing.getId(), folderUrl: existing.getUrl(), message: 'Folder already exists' };
    }

    // ── Create case folder + standard subfolders ─────────────────────────────
    var caseFolder = folderParent.createFolder(caseFolderName);
    caseFolder.createFolder('Step 1 - Informal');
    caseFolder.createFolder('Step 2 - Written');
    caseFolder.createFolder('Step 3 - Review');
    caseFolder.createFolder('Supporting Documents');

    // ── Share: member gets EDITOR (can upload evidence) ──────────────────────
    if (memberEmail) {
      try {
        caseFolder.addEditor(memberEmail);
      } catch (shareErr) {
        Logger.log('setupDriveFolderForGrievance: could not share with member ' + memberEmail + ': ' + shareErr.message);
      }
    }

    // ── Share: assigned steward gets EDITOR ──────────────────────────────────
    var stewardEmail = typeof grievance.stewardEmail === 'string'
      ? grievance.stewardEmail
      : (grievance['Assigned Steward'] ? String(grievance['Assigned Steward']).trim() : '');
    if (stewardEmail && stewardEmail.indexOf('@') !== -1) {
      try {
        caseFolder.addEditor(stewardEmail);
      } catch (shareErr) {
        Logger.log('setupDriveFolderForGrievance: could not share with steward ' + stewardEmail + ': ' + shareErr.message);
      }
    }

    // ── Write URL back to Grievance Log ─────────────────────────────────────
    updateGrievanceFolderLink(grievanceId, caseFolder.getUrl());

    if (typeof logAuditEvent === 'function') {
      logAuditEvent(AUDIT_EVENTS.FOLDER_CREATED, {
        grievanceId:  grievanceId,
        folderId:     caseFolder.getId(),
        folderName:   caseFolderName,
        memberEmail:  memberEmail,
        createdBy:    Session.getActiveUser().getEmail()
      });
    }

    return { success: true, folderId: caseFolder.getId(), folderUrl: caseFolder.getUrl(), message: 'Folder created successfully' };

  } catch (error) {
    Logger.log('setupDriveFolderForGrievance error: ' + error.message);
    return errorResponse(error.message, 'setupDriveFolderForGrievance');
  }
}

/**
 * Scans the root grievance Drive folder for empty sub-folders whose
 * corresponding grievance is resolved, and removes them.
 * Designed to run on a weekly time-based trigger.
 * @returns {Object} Cleanup results
 */
function cleanupEmptyDriveFolders() {
  var rootFolder;
  try {
    rootFolder = getOrCreateRootFolder();
  } catch (_e) {
    return { success: false, reason: 'Could not access root folder' };
  }

  var removed = 0;
  var skipped = 0;
  var subFolders = rootFolder.getFolders();

  // Build a set of resolved grievance IDs for quick lookup
  var resolvedIds = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.STATUS).getValues();
      for (var i = 0; i < data.length; i++) {
        var status = String(data[i][GRIEVANCE_COLS.STATUS - 1] || '').toLowerCase();
        if (status === 'resolved' || status === 'closed' || status === 'withdrawn') {
          var gId = String(data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '');
          if (gId) resolvedIds[gId] = true;
        }
      }
    }
  } catch (sheetErr) {
    Logger.log('cleanupEmptyDriveFolders: sheet read error: ' + sheetErr.message);
  }

  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    try {
      // Only remove if folder tree is empty (no files in any sub-level)
      if (_isFolderTreeEmpty(folder)) {
        // Check if folder name contains a resolved grievance ID
        var folderName = folder.getName();
        // M-49: Extract grievance ID from folder name using known naming patterns,
        // then do O(1) direct lookup instead of O(n) scan through all resolved IDs.
        // Simple template: "{grievanceId} - {date}" => ID is everything before " - "
        var isResolved = false;
        var dashIdx = folderName.indexOf(' - ');
        if (dashIdx > 0) {
          var candidateId = folderName.substring(0, dashIdx).trim();
          if (resolvedIds[candidateId]) {
            isResolved = true;
          }
        }
        // Fallback for non-standard naming: linear scan (rare path)
        if (!isResolved) {
          var resolvedKeys = Object.keys(resolvedIds);
          for (var rk = 0; rk < resolvedKeys.length; rk++) {
            if (folderName.indexOf(resolvedKeys[rk]) >= 0) {
              isResolved = true;
              break;
            }
          }
        }
        if (isResolved) {
          folder.setTrashed(true);
          removed++;
        } else {
          skipped++;
        }
      } else {
        skipped++;
      }
    } catch (folderErr) {
      Logger.log('cleanupEmptyDriveFolders: error on folder: ' + folderErr.message);
      skipped++;
    }
  }

  if (removed > 0) {
    logAuditEvent(AUDIT_EVENTS.SYSTEM_ACTION || 'SYSTEM_ACTION', {
      action: 'cleanupEmptyDriveFolders',
      removed: removed,
      skipped: skipped
    });
  }

  return { success: true, removed: removed, skipped: skipped };
}

/**
 * Recursively checks whether a folder tree contains any files.
 * @param {Folder} folder
 * @returns {boolean} true if folder and all sub-folders contain no files
 * @private
 */
function _isFolderTreeEmpty(folder) {
  if (folder.getFiles().hasNext()) return false;
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    if (!_isFolderTreeEmpty(subs.next())) return false;
  }
  return true;
}

/**
 * Creates Drive folders for multiple grievances in batches
 * @param {string[]} grievanceIds - Array of grievance IDs
 * @return {Object} Result with success count
 */
function batchCreateGrievanceFolders(grievanceIds) {
  const results = {
    success: 0,
    failed: 0,
    errors: []
  };

  const startTime = new Date().getTime();

  for (let i = 0; i < grievanceIds.length; i++) {
    // Check time limit
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
      results.errors.push('Time limit reached, some folders not created');
      break;
    }

    // Batch pause
    if (i > 0 && i % BATCH_LIMITS.MAX_API_CALLS_PER_BATCH === 0) {
      Utilities.sleep(BATCH_LIMITS.PAUSE_BETWEEN_BATCHES_MS);
    }

    const result = setupDriveFolderForGrievance(grievanceIds[i]);

    if (result.success) {
      results.success++;
    } else {
      results.failed++;
      results.errors.push(`${grievanceIds[i]}: ${result.error}`);
    }
  }

  return {
    success: true,
    created: results.success,
    failed: results.failed,
    errors: results.errors,
    message: `Created ${results.success} folders, ${results.failed} failed`
  };
}

/**
 * Menu wrapper: Setup Drive folder for the currently selected grievance
 * Gets the grievance ID from the active row in Grievance Log
 */
function setupFolderForSelectedGrievance() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Ensure sheet has enough columns for Drive folder columns
  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var range = ss.getActiveRange();
  var row = range.getRow();

  // Validate selection
  var validationError = null;
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    validationError = 'Please select a grievance row in the Grievance Log sheet first.';
  } else if (range.getRow() < 2) {
    validationError = 'Please select a data row (not the header).';
  } else if (!sheet.getRange(range.getRow(), GRIEVANCE_COLS.GRIEVANCE_ID).getValue()) {
    validationError = 'No Grievance ID found in the selected row.';
  }

  if (validationError) {
    ui.alert('📁 Setup Grievance Folder', validationError, ui.ButtonSet.OK);
    return;
  }

  var grievanceId = sheet.getRange(range.getRow(), GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  // Check if folder already exists
  var existingUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
  if (existingUrl) {
    var response = ui.alert('📁 Folder Already Exists',
      'A folder already exists for grievance ' + grievanceId + '.\n\n' +
      'Would you like to open it?',
      ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      var html = HtmlService.createHtmlOutput(
        '<script>window.open(' + JSON.stringify(existingUrl) + ', "_blank"); google.script.host.close();</script>'
      ).setWidth(1).setHeight(1);
      ui.showModalDialog(html, 'Opening folder...');
    }
    return;
  }

  // Create the folder
  ss.toast('Creating folder for ' + grievanceId + '...', '📁 Drive', 3);
  var result = setupDriveFolderForGrievance(grievanceId);

  if (result.success) {
    ui.alert('✅ Folder Created',
      'Folder created for grievance ' + grievanceId + '.\n\n' +
      'The folder URL has been saved to the Grievance Log.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Error',
      'Failed to create folder: ' + result.error,
      ui.ButtonSet.OK);
  }
}

/**
 * Menu wrapper: Batch create folders for all grievances missing folders
 */
function batchCreateAllMissingFolders() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    ui.alert('Error', 'Grievance Log not found.', ui.ButtonSet.OK);
    return;
  }

  ensureMinimumColumns(sheet, getGrievanceHeaders().length);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('📁 Batch Create Folders', 'No grievances found.', ui.ButtonSet.OK);
    return;
  }

  // Get grievance IDs and folder URLs
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(GRIEVANCE_COLS.GRIEVANCE_ID, GRIEVANCE_COLS.DRIVE_FOLDER_URL)).getValues();

  var missingFolders = [];
  for (var i = 0; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var folderUrl = data[i][GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1];

    if (grievanceId && !folderUrl) {
      missingFolders.push(grievanceId);
    }
  }

  if (missingFolders.length === 0) {
    ui.alert('📁 Batch Create Folders',
      'All grievances already have folders!',
      ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert('📁 Batch Create Folders',
    'Found ' + missingFolders.length + ' grievances without folders.\n\n' +
    'Create folders for all of them?\n\n' +
    'This may take a few minutes for large numbers.',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Creating ' + missingFolders.length + ' folders...', '📁 Drive', 10);

  var result = batchCreateGrievanceFolders(missingFolders);

  ui.alert('📁 Batch Create Complete',
    'Created: ' + result.created + ' folders\n' +
    'Failed: ' + result.failed + '\n\n' +
    (result.errors.length > 0 ? 'Errors:\n' + result.errors.slice(0, 5).join('\n') : ''),
    ui.ButtonSet.OK);
}

/**
 * Updates the Drive folder link in grievance record
 * @param {string} grievanceId - The grievance ID
 * @param {string} folderUrl - The folder URL
 */
function updateGrievanceFolderLink(grievanceId, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // CR-09: Use canonical SHEETS.GRIEVANCE_LOG instead of SHEETS.GRIEVANCE_LOG
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // H-05: Null check after getSheetByName
  if (!sheet) {
    Logger.log('updateGrievanceFolderLink: Grievance Log sheet not found');
    return;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      sheet.getRange(i + 1, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folderUrl);
      break;
    }
  }
}

/**
 * Opens the Drive folder for selected grievance
 */
function openGrievanceFolder() {
  const sheet = SpreadsheetApp.getActiveSheet();
  // CR-09: Use canonical SHEETS.GRIEVANCE_LOG instead of SHEETS.GRIEVANCE_LOG
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    showAlert('Please select a grievance in the Grievance Tracker', 'Wrong Sheet');
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) return;

  const folderUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();

  if (folderUrl) {
    const html = HtmlService.createHtmlOutput(
      `<script>window.open(${JSON.stringify(folderUrl)}, '_blank'); google.script.host.close();</script>`
    ).setWidth(100).setHeight(50);
    SpreadsheetApp.getUi().showModalDialog(html, 'Opening folder...');
  } else {
    if (showConfirmation('No folder exists. Create one now?', 'Create Folder')) {
      const grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
      const result = setupDriveFolderForGrievance(grievanceId);
      if (result.success) {
        const html = HtmlService.createHtmlOutput(
          `<script>window.open(${JSON.stringify(result.folderUrl)}, '_blank'); google.script.host.close();</script>`
        ).setWidth(100).setHeight(50);
        SpreadsheetApp.getUi().showModalDialog(html, 'Opening folder...');
      }
    }
  }
}

/**
 * Sanitizes a string for use as a folder name
 * @param {string} name - The name to sanitize
 * @return {string} Sanitized folder name
 */
function sanitizeFolderName(name) {
  if (!name) return 'Unknown';
  return name
    .trim()
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, '_')
    .substring(0, 50);
}

// ============================================================================
// MEETING CALENDAR INTEGRATION
// ============================================================================

/**
 * Gets or creates the union meetings calendar
 * @return {Calendar} The meetings calendar
 */
function getOrCreateMeetingsCalendar() {
  var calendarName = CALENDAR_CONFIG.MEETING_CALENDAR_NAME;
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  }
  var newCalendar = CalendarApp.createCalendar(calendarName, {
    summary: 'Union meeting events - Auto-managed by Union Dashboard',
    color: CalendarApp.Color.GREEN
  });
  return newCalendar;
}

/**
 * Creates a Google Calendar event for a meeting
 * @param {Object} meetingData - { name, date, time, duration, type }
 * @returns {string} Calendar event ID or empty string on failure
 */
function createMeetingCalendarEvent(meetingData) {
  try {
    var calendar = getOrCreateMeetingsCalendar();
    // M-55: Append script timezone offset to ensure correct date parsing
    var tz = Session.getScriptTimeZone();
    var tzOffset = Utilities.formatDate(new Date(), tz, 'XXX'); // e.g., "-05:00"
    var meetingDate = new Date(meetingData.date + 'T00:00:00' + tzOffset);
    var startTime = meetingData.time || '09:00';
    var durationHours = parseFloat(meetingData.duration) || 1;
    var timeParts = startTime.split(':');
    var startHour = parseInt(timeParts[0], 10) || 9;
    var startMin = parseInt(timeParts[1], 10) || 0;

    var start = new Date(meetingDate);
    start.setHours(startHour, startMin, 0, 0);

    var end = new Date(start.getTime() + durationHours * 60 * 60 * 1000);

    var eventTitle = '[MTG] ' + meetingData.name + ' (' + (meetingData.type || 'In-Person') + ')';
    var event;
    try {
      event = calendar.createEvent(eventTitle, start, end, {
        description: 'Meeting ID: ' + (meetingData.meetingId || 'TBD') + '\n' +
                     'Type: ' + (meetingData.type || 'In-Person') + '\n' +
                     'Check-in opens on the day of the event.\n\n' +
                     'Auto-generated by Union Dashboard'
      });
    } catch (calErr) {
      // F31: Quota-aware error handling for calendar creation
      var msg = String(calErr.message || '');
      if (msg.indexOf('quota') !== -1 || msg.indexOf('limit') !== -1 || msg.indexOf('rate') !== -1) {
        Logger.log('Calendar quota exceeded, skipping event creation: ' + msg);
      } else {
        Logger.log('Error creating calendar event: ' + msg);
      }
      return '';
    }

    // Add email reminders
    event.removeAllReminders();
    event.addEmailReminder(24 * 60);  // 1 day before
    event.addEmailReminder(60);        // 1 hour before

    return event.getId();
  } catch (error) {
    Logger.log('Error creating meeting calendar event: ' + error.message);
    return '';
  }
}

/**
 * Deletes a meeting calendar event by ID
 * @param {string} eventId - Calendar event ID
 */
function deleteMeetingCalendarEvent(eventId) {
  if (!eventId) return;
  try {
    var calendar = getOrCreateMeetingsCalendar();
    var event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
    }
  } catch (error) {
    Logger.log('Error deleting meeting calendar event: ' + error.message);
  }
}

/**
 * Emails the attendance report for a meeting to specified stewards
 * @param {string} meetingId - Meeting ID to report on
 * @param {string} recipientEmails - Comma-separated steward emails
 * @returns {Object} { success, error }
 */
function emailMeetingAttendanceReport(meetingId, recipientEmails) {
  // M-56: Authorization check — attendance reports contain PII, restrict to steward/admin
  if (typeof getUserRole_ === 'function') {
    var callerEmail = Session.getActiveUser().getEmail();
    var callerRole = getUserRole_(callerEmail);
    if (callerRole !== 'admin' && callerRole !== 'steward') {
      return errorResponse('Unauthorized: only stewards and admins may send attendance reports');
    }
  }

  if (!meetingId || !recipientEmails) {
    return errorResponse('Meeting ID and recipient emails are required');
  }

  // Validate all recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',');
  for (var e = 0; e < emails.length; e++) {
    if (!emailRegex.test(emails[e].trim())) {
      return errorResponse('Invalid email address: ' + emails[e].trim());
    }
  }

  try {
    var result = getMeetingAttendees(meetingId);
    if (!result.success) {
      return errorResponse(result.error || 'Could not retrieve attendance data');
    }

    // Find meeting details from the check-in log
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() <= 1) {
      return errorResponse('No meeting check-in data found');
    }
    var data = sheet.getDataRange().getValues();
    var meetingName = '';
    var meetingDate = '';
    var meetingType = '';

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '') === meetingId) {
        meetingName = data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '';
        meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
        meetingType = data[i][MEETING_CHECKIN_COLS.MEETING_TYPE - 1] || '';
        break;
      }
    }

    var dateStr = meetingDate instanceof Date ? meetingDate.toLocaleDateString() : String(meetingDate);

    // Build email body
    var body = '<h2>Meeting Attendance Report</h2>' +
      '<table style="border-collapse:collapse;margin:10px 0">' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Meeting:</td><td style="padding:4px 12px">' + escapeHtml(meetingName) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Date:</td><td style="padding:4px 12px">' + escapeHtml(dateStr) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Type:</td><td style="padding:4px 12px">' + escapeHtml(meetingType) + '</td></tr>' +
      '<tr><td style="padding:4px 12px;font-weight:bold">Total Attendees:</td><td style="padding:4px 12px">' + escapeHtml(String(result.count)) + '</td></tr>' +
      '</table>';

    if (result.attendees.length > 0) {
      body += '<h3>Attendees</h3>' +
        '<table style="border-collapse:collapse;border:1px solid #ddd">' +
        '<tr style="background:#059669;color:white">' +
        '<th style="padding:8px;border:1px solid #ddd">#</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Member ID</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Name</th>' +
        '<th style="padding:8px;border:1px solid #ddd">Check-In Time</th>' +
        '</tr>';

      for (var j = 0; j < result.attendees.length; j++) {
        var a = result.attendees[j];
        body += '<tr>' +
          '<td style="padding:6px;border:1px solid #ddd">' + (j + 1) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.memberId)) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.name)) + '</td>' +
          '<td style="padding:6px;border:1px solid #ddd">' + escapeHtml(String(a.time)) + '</td>' +
          '</tr>';
      }
      body += '</table>';
    } else {
      body += '<p><em>No members checked in to this meeting.</em></p>';
    }

    body += '<br><p style="font-size:12px;color:#666">Auto-generated by Union Dashboard</p>';

    var emailResult = safeSendEmail_({
      to: recipientEmails,
      subject: 'Meeting Attendance: ' + meetingName + ' (' + dateStr + ')',
      htmlBody: body,
      name: 'Union Dashboard'
    });
    if (!emailResult.success) return errorResponse(emailResult.error);

    return { success: true, message: 'Attendance report emailed to ' + recipientEmails };
  } catch (error) {
    Logger.log('Error emailing attendance report: ' + error.message);
    return errorResponse(error.message);
  }
}

// ============================================================================
// MEETING DOCS (Notes & Agenda) INTEGRATION
// ============================================================================

/**
 * Gets or creates the Meeting Notes folder under the root Drive folder
 * @return {Folder} The Meeting Notes folder
 */
function getOrCreateMeetingNotesFolder() {
  var rootFolder = getOrCreateRootFolder();
  var folderName = 'Meeting Notes';
  var folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var newFolder = rootFolder.createFolder(folderName);
  newFolder.setDescription('Meeting Notes - Auto-managed by Union Dashboard');
  return newFolder;
}

/**
 * Gets or creates the Meeting Agenda folder under the root Drive folder
 * @return {Folder} The Meeting Agenda folder
 */
function getOrCreateMeetingAgendaFolder() {
  var rootFolder = getOrCreateRootFolder();
  var folderName = 'Meeting Agenda';
  var folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var newFolder = rootFolder.createFolder(folderName);
  newFolder.setDescription('Meeting Agenda - Auto-managed by Union Dashboard (Steward access only)');
  return newFolder;
}

/**
 * Creates Meeting Notes and Meeting Agenda Google Docs for a meeting
 * Notes: shared view-only with members (day after meeting)
 * Agenda: shared with stewards only (3 days prior), NOT shared with members
 * @param {Object} meetingData - { meetingId, name, date }
 * @returns {Object} { notesUrl, agendaUrl }
 */
function createMeetingDocs(meetingData) {
  var result = { notesUrl: '', agendaUrl: '' };
  try {
    var dateStr = meetingData.date || '';
    var meetingName = meetingData.name || 'Meeting';

    // Create Meeting Notes doc
    var notesFolder = getOrCreateMeetingNotesFolder();
    var notesTitle = 'Meeting Notes - ' + meetingName + ' - ' + dateStr;
    var notesDoc = DocumentApp.create(notesTitle);
    var notesBody = notesDoc.getBody();
    notesBody.appendParagraph(notesTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    notesBody.appendParagraph('Meeting ID: ' + (meetingData.meetingId || ''));
    notesBody.appendParagraph('Date: ' + dateStr);
    notesBody.appendParagraph('');
    notesBody.appendParagraph('Notes:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    notesBody.appendParagraph('');
    notesDoc.saveAndClose();

    // Move to Meeting Notes folder
    // L-39: Use file.moveTo() instead of deprecated parent.removeFile()
    var notesFile = DriveApp.getFileById(notesDoc.getId());
    notesFile.moveTo(notesFolder);
    result.notesUrl = notesDoc.getUrl();

    // Create Meeting Agenda doc
    var agendaFolder = getOrCreateMeetingAgendaFolder();
    var agendaTitle = 'Meeting Agenda - ' + meetingName + ' - ' + dateStr;
    var agendaDoc = DocumentApp.create(agendaTitle);
    var agendaBody = agendaDoc.getBody();
    agendaBody.appendParagraph(agendaTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    agendaBody.appendParagraph('Meeting ID: ' + (meetingData.meetingId || ''));
    agendaBody.appendParagraph('Date: ' + dateStr);
    agendaBody.appendParagraph('');
    agendaBody.appendParagraph('Agenda Items:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    agendaBody.appendParagraph('1. ');
    agendaBody.appendParagraph('2. ');
    agendaBody.appendParagraph('3. ');
    agendaBody.appendParagraph('');
    agendaBody.appendParagraph('Action Items:').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    agendaBody.appendParagraph('');
    agendaDoc.saveAndClose();

    // Move to Meeting Agenda folder
    // L-39: Use file.moveTo() instead of deprecated parent.removeFile()
    var agendaFile = DriveApp.getFileById(agendaDoc.getId());
    agendaFile.moveTo(agendaFolder);
    result.agendaUrl = agendaDoc.getUrl();

    if (typeof logAuditEvent === 'function') {
      logAuditEvent('MEETING_DOCS_CREATED', {
        meetingId: meetingData.meetingId,
        notesDocUrl: result.notesUrl,
        agendaDocUrl: result.agendaUrl
      });
    }
  } catch (error) {
    Logger.log('Error creating meeting docs: ' + error.message);
  }
  return result;
}

/**
 * Sets a Google Doc to view-only for anyone with the link
 * Used to make meeting notes viewable by members after the meeting
 * @param {string} docUrl - The Google Doc URL
 */
function setDocViewOnlyByLink(docUrl) {
  try {
    // Extract file ID from URL
    var match = docUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return;
    var fileId = match[1];
    var file = DriveApp.getFileById(fileId);
    // M-57: DOMAIN_WITH_LINK fails for non-Workspace (consumer) orgs.
    // Try domain sharing first, fall back to ANYONE_WITH_LINK.
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (domainError) {
      Logger.log('Domain sharing unavailable, falling back to ANYONE_WITH_LINK: ' + domainError.message);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (error) {
    Logger.log('Error setting doc to view-only: ' + error.message);
  }
}

/**
 * Sends meeting document links to stewards via email
 * @param {string} meetingName - Name of the meeting
 * @param {string} meetingDate - Date of the meeting
 * @param {string} docUrl - The document URL to share
 * @param {string} docType - 'notes' or 'agenda'
 * @param {string} recipientEmails - Comma-separated steward emails
 */
function emailMeetingDocLink(meetingName, meetingDate, docUrl, docType, recipientEmails) {
  if (!recipientEmails || !docUrl) return;

  // Validate all recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',');
  for (var e = 0; e < emails.length; e++) {
    if (!emailRegex.test(emails[e].trim())) {
      Logger.log('Invalid email address in recipient list: ' + emails[e].trim());
      return;
    }
  }

  try {
    var typeLabel = docType === 'agenda' ? 'Meeting Agenda' : 'Meeting Notes';
    var safeDocUrl = /^https:\/\/docs\.google\.com\//.test(docUrl) ? docUrl : '';
    var body = '<h2>' + escapeHtml(typeLabel) + '</h2>' +
      '<p><strong>Meeting:</strong> ' + escapeHtml(meetingName) + '</p>' +
      '<p><strong>Date:</strong> ' + escapeHtml(meetingDate) + '</p>' +
      '<p>Click the link below to access the ' + escapeHtml(typeLabel.toLowerCase()) + ':</p>' +
      (safeDocUrl ? '<p><a href="' + escapeHtml(safeDocUrl) + '" style="background:#1a73e8;color:white;padding:10px 20px;border-radius:4px;text-decoration:none;display:inline-block">' +
      'Open ' + escapeHtml(typeLabel) + '</a></p>' : '<p><em>Document link unavailable.</em></p>') +
      '<br><p style="font-size:12px;color:#666">Auto-generated by Union Dashboard</p>';

    safeSendEmail_({
      to: recipientEmails,
      subject: typeLabel + ': ' + meetingName + ' (' + meetingDate + ')',
      htmlBody: body,
      name: 'Union Dashboard'
    });
  } catch (error) {
    Logger.log('Error emailing meeting doc link: ' + error.message);
  }
}

/**
 * Gets all steward emails from the Member Directory
 * @returns {string} Comma-separated steward emails
 */
function getAllStewardEmails_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet || sheet.getLastRow() < 2) return '';

    var data = sheet.getDataRange().getValues();
    var emails = [];
    for (var i = 1; i < data.length; i++) {
      var isSteward = data[i][MEMBER_COLS.IS_STEWARD - 1];
      if (isTruthyValue(isSteward)) {
        var email = String(data[i][MEMBER_COLS.EMAIL - 1] || '').trim();
        if (email) emails.push(email);
      }
    }
    return emails.join(', ');
  } catch (e) {
    Logger.log('Error getting all steward emails: ' + e.message);
    return '';
  }
}

/**
 * Sends scheduled meeting doc notifications from dailyTrigger
 * - Agenda link: 3 days before -> selected stewards (AGENDA_STEWARDS column)
 * - Agenda link: 1 day before -> ALL stewards (from Member Directory)
 * - Notes link: 1 day before -> attendance notification stewards (NOTIFY_STEWARDS column)
 * - Sets notes to view-only: 1 day after meeting (for members)
 * @returns {Object} { agendaSent, notesSent, notesPublished }
 */
function processMeetingDocNotifications() {
  var result = { agendaSent: 0, notesSent: 0, notesPublished: 0 };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEETING_CHECKIN_LOG);
    if (!sheet || sheet.getLastRow() < 2) return result;

    var data = sheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var processed = {};  // Track processed meeting IDs
    var allStewardEmails = null;  // Lazy-loaded

    for (var i = 1; i < data.length; i++) {
      var meetingId = String(data[i][MEETING_CHECKIN_COLS.MEETING_ID - 1] || '');
      if (!meetingId || processed[meetingId]) continue;
      processed[meetingId] = true;

      var meetingDate = data[i][MEETING_CHECKIN_COLS.MEETING_DATE - 1];
      if (!(meetingDate instanceof Date)) continue;

      var meetingDay = new Date(meetingDate);
      meetingDay.setHours(0, 0, 0, 0);
      var diffDays = Math.round((meetingDay.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));

      var meetingName = String(data[i][MEETING_CHECKIN_COLS.MEETING_NAME - 1] || '');
      var notifyEmails = String(data[i][MEETING_CHECKIN_COLS.NOTIFY_STEWARDS - 1] || '');
      var notesUrl = String(data[i][MEETING_CHECKIN_COLS.NOTES_DOC_URL - 1] || '');
      var agendaUrl = String(data[i][MEETING_CHECKIN_COLS.AGENDA_DOC_URL - 1] || '');
      var agendaStewards = String(data[i][MEETING_CHECKIN_COLS.AGENDA_STEWARDS - 1] || '');
      var dateStr = meetingDate.toLocaleDateString();

      // 3 days before: send agenda link to SELECTED stewards only
      if (diffDays === 3 && agendaUrl && agendaStewards) {
        emailMeetingDocLink(meetingName, dateStr, agendaUrl, 'agenda', agendaStewards);
        result.agendaSent++;
      }

      // 1 day before: send agenda link to ALL stewards, notes link to notify stewards
      if (diffDays === 1) {
        if (agendaUrl) {
          // Lazy-load all steward emails only when needed
          if (allStewardEmails === null) {
            allStewardEmails = getAllStewardEmails_();
          }
          if (allStewardEmails) {
            emailMeetingDocLink(meetingName, dateStr, agendaUrl, 'agenda', allStewardEmails);
            result.agendaSent++;
          }
        }
        if (notesUrl && notifyEmails) {
          emailMeetingDocLink(meetingName, dateStr, notesUrl, 'notes', notifyEmails);
          result.notesSent++;
        }
      }

      // 1 day after: set notes to view-only (available to members)
      if (diffDays === -1 && notesUrl) {
        setDocViewOnlyByLink(notesUrl);
        result.notesPublished++;
      }
    }
  } catch (error) {
    Logger.log('Error processing meeting doc notifications: ' + error.message);
  }
  return result;
}

// ============================================================================
// MEMBER DRIVE FOLDER
// ============================================================================

/**
 * Creates or retrieves a Google Drive folder for a member
 * If the member has an existing grievance with a folder, reuses that folder
 * Otherwise creates a new folder under the root folder
 * @param {string} memberId - The member ID
 * @returns {Object} { success, folderUrl, folderId, message, error }
 */
/**
 * @deprecated Use getOrCreateMemberAdminFolder(memberEmail) instead.
 * Kept for backwards-compatibility with steward_view "Create Member Folder" button.
 * Resolves the member's email from their memberId, then delegates.
 */
function setupDriveFolderForMember(memberId) {
  try {
    if (!memberId) return errorResponse('Member ID is required');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberSheet) return errorResponse('Member Directory not found');

    var memberData = memberSheet.getDataRange().getValues();
    var emailCol   = -1;
    for (var h = 0; h < memberData[0].length; h++) {
      if (String(memberData[0][h]).toLowerCase().trim() === 'email') { emailCol = h; break; }
    }
    var memberEmail = '';
    for (var i = 1; i < memberData.length; i++) {
      if (String(memberData[i][MEMBER_COLS.MEMBER_ID - 1]) === String(memberId)) {
        memberEmail = emailCol !== -1 ? String(memberData[i][emailCol]).trim() : '';
        break;
      }
    }
    if (!memberEmail) return errorResponse('Member not found or has no email: ' + memberId);

    var result = getOrCreateMemberAdminFolder(memberEmail);
    if (!result) return errorResponse('Could not create member admin folder');
    return { success: true, folderId: result.masterFolder.getId(), folderUrl: result.masterFolder.getUrl(), message: 'Member admin folder ready' };
  } catch (err) {
    return errorResponse(err.message);
  }
}

// ============================================================================
// GRIEVANCE CALENDAR INTEGRATION
// ============================================================================

/**
 * Gets or creates the grievance deadlines calendar
 * @return {Calendar} The deadlines calendar
 */
function getOrCreateDeadlinesCalendar() {
  const calendarName = CALENDAR_CONFIG.CALENDAR_NAME;

  // Check owned calendars
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  }

  // Create new calendar
  const newCalendar = CalendarApp.createCalendar(calendarName, {
    summary: 'Grievance deadline tracking - Auto-managed by Union Dashboard',
    color: CalendarApp.Color.RED
  });

  return newCalendar;
}

/**
 * Syncs all grievance deadlines to calendar
 * @return {Object} Result with sync count
 */
function syncDeadlinesToCalendar() {
  try {
    const calendar = getOrCreateDeadlinesCalendar();
    const openGrievances = getOpenGrievances();

    let synced = 0;
    let skipped = 0;
    const startTime = new Date().getTime();

    for (const grievance of openGrievances) {
      // Check time limit
      if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - 30000) {
        break;
      }

      const result = syncGrievanceDeadlinesToCalendar(
        grievance,
        calendar
      );

      if (result.synced) {
        synced++;
      } else {
        skipped++;
      }

      // Rate limiting pause
      if (synced % 20 === 0) {
        Utilities.sleep(200);
      }
    }

    logAuditEvent(AUDIT_EVENTS.CALENDAR_SYNCED, {
      synced: synced,
      skipped: skipped,
      syncedBy: Session.getActiveUser().getEmail()
    });

    return {
      success: true,
      synced: synced,
      skipped: skipped,
      message: `Synced ${synced} grievances to calendar`
    };

  } catch (error) {
    console.error('Error syncing to calendar:', error);
    return errorResponse(error.message);
  }
}

/**
 * Syncs a single grievance's deadlines to calendar
 * @param {Object} grievance - Grievance data object
 * @param {Calendar} calendar - Target calendar
 * @return {Object} Sync result
 */
function syncGrievanceDeadlinesToCalendar(grievance, calendar) {
  const grievanceId = grievance['Grievance ID'];
  const memberName = grievance['Member Name'];
  const currentStep = grievance['Current Step'];

  // Get the deadline for current step
  let deadline;
  switch (currentStep) {
    case 'Step I':
    case 'Informal':
      deadline = grievance['Step 1 Due'];
      break;
    case 'Step II':
      deadline = grievance['Step 2 Due'];
      break;
    case 'Step III':
    case 'Arbitration':
      deadline = grievance['Step 3 Due'];
      break;
    default:
      return { synced: false, reason: 'No applicable deadline' };
  }

  if (!(deadline instanceof Date)) {
    return { synced: false, reason: 'Invalid deadline date' };
  }

  // Check if deadline is in the past
  if (deadline < new Date()) {
    return { synced: false, reason: 'Deadline already passed' };
  }

  // Create event title
  const eventTitle = `[GRV] ${grievanceId} - Step ${currentStep} Due (${memberName})`;

  // Dedup via stored event IDs in ScriptProperties (keyed by grievance ID)
  var props = PropertiesService.getScriptProperties();
  var calEventKey = 'CAL_EVENT_' + grievanceId;
  var storedEventId = props.getProperty(calEventKey);
  var description = 'Grievance: ' + grievanceId + '\n' +
                    'Member: ' + memberName + '\n' +
                    'Step: ' + currentStep + '\n' +
                    'Action Required: Response deadline\n\n' +
                    'Auto-generated by Union Dashboard';

  if (storedEventId) {
    try {
      var existingEvent = calendar.getEventById(storedEventId);
      if (existingEvent) {
        existingEvent.setTitle(eventTitle);
        existingEvent.setAllDayDate(deadline);
        existingEvent.setDescription(description);
        return { synced: true, updated: true };
      }
    } catch (_lookupErr) {
      // Stored ID stale — fall through to create new
    }
  }

  // Fallback: search by day in case event exists without stored ID
  var existingEvents = calendar.getEventsForDay(deadline, { search: grievanceId });
  if (existingEvents.length > 0) {
    var found = existingEvents[0];
    found.setTitle(eventTitle);
    props.setProperty(calEventKey, found.getId());
    return { synced: true, updated: true };
  }

  // Create all-day event
  var event = calendar.createAllDayEvent(eventTitle, deadline, {
    description: description
  });
  props.setProperty(calEventKey, event.getId());

  // Set reminders
  event.removeAllReminders();
  CALENDAR_CONFIG.REMINDER_DAYS.forEach(days => {
    event.addEmailReminder(days * 24 * 60); // Convert days to minutes
  });

  return { synced: true, created: true };
}

// Note: syncSingleGrievanceToCalendar() is defined in MobileQuickActions.gs

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Quota-safe email wrapper. Checks MailApp daily quota before sending.
 * @param {Object} options - MailApp.sendEmail options (to, subject, body, etc.)
 * @returns {boolean} true if sent, false if quota exceeded
 */
function safeSendEmail(options) {
  if (MailApp.getRemainingDailyQuota() < 1) {
    Logger.log('Email quota exceeded, skipping: ' + (options.subject || 'no subject'));
    return false;
  }
  MailApp.sendEmail(options);
  return true;
}

/**
 * Sends deadline reminder email for upcoming grievances
 * @param {number} daysAhead - Days to look ahead
 * @return {Object} Result object
 */
function sendDeadlineReminders(daysAhead) {
  try {
    const deadlines = getUpcomingDeadlines(daysAhead || 7);
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(userEmail)) {
      return errorResponse('Could not determine a valid email address for the current user');
    }

    if (deadlines.length === 0) {
      return { success: true, sent: false, message: 'No upcoming deadlines' };
    }

    // Build email body
    let body = `<h2>Upcoming Grievance Deadlines</h2>`;
    body += `<p>The following grievances have deadlines in the next ${escapeHtml(String(daysAhead))} days:</p>`;
    body += `<table style="border-collapse: collapse; width: 100%;">`;
    body += `<tr style="background: #f0f0f0;">
               <th style="padding: 10px; border: 1px solid #ddd;">Grievance</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Member</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Step</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Due Date</th>
               <th style="padding: 10px; border: 1px solid #ddd;">Days Left</th>
             </tr>`;

    deadlines.forEach(d => {
      const urgent = d.daysLeft <= 3 ? 'style="background: #fee2e2;"' : '';
      body += `<tr ${urgent}>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.grievanceId)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.memberName)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.step)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(d.date)}</td>
                 <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(String(d.daysLeft))}</td>
               </tr>`;
    });

    body += `</table>`;
    body += `<p style="margin-top: 20px; color: #666; font-size: 12px;">
               This is an automated reminder from the Union Dashboard.
             </p>`;

    // Send email
    safeSendEmail_({
      to: userEmail,
      subject: `[Union Dashboard] ${deadlines.length} Upcoming Grievance Deadline${deadlines.length > 1 ? 's' : ''}`,
      htmlBody: body
    });

    return {
      success: true,
      sent: true,
      count: deadlines.length,
      message: `Sent reminder for ${deadlines.length} deadlines`
    };

  } catch (error) {
    console.error('Error sending reminders:', error);
    return errorResponse(error.message);
  }
}

/**
 * Sends email to a member
 * @param {string} memberId - The member ID
 * @param {string} subject - Email subject
 * @param {string} body - Email body (HTML)
 * @return {Object} Result object
 */
function sendEmailToMember(memberId, subject, body) {
  try {
    // CR-21: Authorization check — only stewards and admins may send emails to members
    var callerEmail = Session.getActiveUser().getEmail();
    if (typeof getUserRole_ === 'function') {
      var role = getUserRole_(callerEmail);
      if (role !== 'admin' && role !== 'steward') {
        return errorResponse('Unauthorized: only stewards and admins may send emails to members');
      }
    }

    const member = getMemberById(memberId);
    if (!member) {
      return errorResponse('Member not found');
    }

    const email = member['Email'] || member.email;
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim())) {
      return errorResponse('Invalid email address');
    }

    // Strip HTML tags from subject — email subjects are plain text only
    var safeSubject = String(subject || '').replace(/<[^>]*>/g, '').trim();
    var safeBody = String(body || '');

    safeSendEmail_({
      to: email,
      subject: safeSubject,
      htmlBody: safeBody
    });

    return {
      success: true,
      message: `Email sent to ${email}`
    };

  } catch (error) {
    console.error('Error sending email:', error);
    return errorResponse(error.message);
  }
}

// ============================================================================
// PDF SIGNATURE ENGINE (Strategic Command Center)
// ============================================================================

/**
 * Gets or creates an archive folder for a specific member
 * Used for storing grievance PDFs and documents
 * @param {string} name - Member's name
 * @param {string} id - Member's ID
 * @returns {Folder} The member's archive folder
 */
function getOrCreateMemberFolder(name, id) {
  // Get archive folder ID from Config or COMMAND_CONFIG
  var archiveFolderId = getConfigValue_(CONFIG_COLS.ARCHIVE_FOLDER_ID) || COMMAND_CONFIG.ARCHIVE_FOLDER_ID;

  // L-04: Declare folderName once at the top instead of re-declaring in each branch
  var folderName = name + ' (' + id + ')';

  if (!archiveFolderId) {
    // Fall back to creating in root folder
    var rootFolder = getOrCreateRootFolder();
    var folders = rootFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
  }

  try {
    var parentFolder = DriveApp.getFolderById(archiveFolderId);
    folderName = name + ' (' + id + ')';
    folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  } catch (e) {
    Logger.log('Archive folder not found, using root: ' + e.message);
    rootFolder = getOrCreateRootFolder();
    folderName = name + ' (' + id + ')';
    folders = rootFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : rootFolder.createFolder(folderName);
  }
}

/**
 * Creates a signature-ready PDF from a grievance template
 * Merges data and adds signature blocks
 * @param {Folder} folder - Target folder for the PDF
 * @param {Object} data - Grievance data object
 * @returns {File} The created PDF file
 */
function createSignatureReadyPDF(folder, data) {
  // Get template ID from Config or COMMAND_CONFIG
  var templateId = getConfigValue_(CONFIG_COLS.TEMPLATE_ID) || COMMAND_CONFIG.TEMPLATE_ID;

  if (!templateId) {
    throw new Error('Document template ID not configured. Please set TEMPLATE_ID in Config sheet.');
  }

  try {
    // Copy template
    var temp = DriveApp.getFileById(templateId).makeCopy('SIGNATURE_REQUIRED_' + data.name, folder);
    var doc = DocumentApp.openById(temp.getId());
    var body = doc.getBody();

    // M-09: Helper to escape $ in replacement values — replaceText() uses regex
    // internally, so $ would be interpreted as backreference syntax
    function escapeReplacement_(val) {
      return String(val).replace(/\$/g, '$$$$');
    }

    // Replace placeholders with data
    body.replaceText('{{MemberName}}', escapeReplacement_(data.name || 'Unknown'));
    body.replaceText('{{MemberID}}', escapeReplacement_(data.id || '000'));
    body.replaceText('{{Date}}', escapeReplacement_(new Date().toLocaleDateString()));
    body.replaceText('{{Details}}', escapeReplacement_(data.details || 'No details provided.'));
    body.replaceText('{{GrievanceID}}', escapeReplacement_(data.grievanceId || ''));
    body.replaceText('{{Articles}}', escapeReplacement_(data.articles || ''));
    body.replaceText('{{Status}}', escapeReplacement_(data.status || ''));
    body.replaceText('{{Unit}}', escapeReplacement_(data.unit || ''));
    body.replaceText('{{Location}}', escapeReplacement_(data.location || ''));
    body.replaceText('{{Steward}}', escapeReplacement_(data.steward || ''));

    // Append legal signature block
    body.appendParagraph(COMMAND_CONFIG.PDF.SIGNATURE_BLOCK ||
      '\n\n__________________________\nMember Signature\n\n__________________________\nSteward Signature\n\n__________________________\nDate');

    doc.saveAndClose();

    // Convert to PDF
    var pdf = folder.createFile(temp.getAs(MimeType.PDF))
                    .setName('Grievance_UNSIGNED_' + data.name + '_' + new Date().toISOString().slice(0,10) + '.pdf');

    // Remove the temp document
    temp.setTrashed(true);

    return pdf;

  } catch (e) {
    Logger.log('Error creating signature PDF: ' + e.message);
    throw new Error('Failed to create PDF: ' + e.message);
  }
}

/**
 * Creates a PDF for the currently selected grievance
 * Saves to member's Drive folder and optionally emails to member
 * Accessible from the Command menu
 */
function createPDFForSelectedGrievance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('Please select a grievance row in the Grievance Log sheet');
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    ui.alert('Please select a grievance row (not the header)');
    return;
  }

  // Get grievance data including member email
  var data = {
    grievanceId: sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue(),
    name: sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
          sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue(),
    id: sheet.getRange(row, GRIEVANCE_COLS.MEMBER_ID).getValue(),
    status: sheet.getRange(row, GRIEVANCE_COLS.STATUS).getValue(),
    articles: sheet.getRange(row, GRIEVANCE_COLS.ARTICLES).getValue(),
    details: sheet.getRange(row, GRIEVANCE_COLS.RESOLUTION).getValue() || 'Pending',
    location: sheet.getRange(row, GRIEVANCE_COLS.LOCATION).getValue(),
    steward: sheet.getRange(row, GRIEVANCE_COLS.STEWARD).getValue(),
    memberEmail: sheet.getRange(row, GRIEVANCE_COLS.MEMBER_EMAIL).getValue()
  };

  var response = ui.alert(
    'Create Signature PDF',
    'Create a signature-ready PDF for grievance ' + data.grievanceId + '?\n\n' +
    'Member: ' + data.name + '\n' +
    'Status: ' + data.status + '\n' +
    'Email: ' + (data.memberEmail || 'Not on file'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    ss.toast('Creating PDF...', COMMAND_CONFIG.SYSTEM_NAME, 5);

    // Get or create member folder
    var folder = getOrCreateMemberFolder(data.name, data.id);

    // Create the PDF
    var pdf = createSignatureReadyPDF(folder, data);

    // Update grievance record with PDF link and folder URL
    if (GRIEVANCE_COLS.DRIVE_FOLDER_URL) {
      sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folder.getUrl());
    }

    // Ask if user wants to email the PDF to member
    var emailResponse = ui.alert(
      'Email PDF to Member?',
      'PDF created successfully!\n\n' +
      'File: ' + pdf.getName() + '\n' +
      'Saved to: ' + folder.getName() + '\n\n' +
      (data.memberEmail
        ? 'Would you like to email this PDF to ' + data.memberEmail + '?'
        : 'No email on file for this member. Add email to column X to enable this feature.'),
      data.memberEmail ? ui.ButtonSet.YES_NO : ui.ButtonSet.OK
    );

    if (emailResponse === ui.Button.YES && data.memberEmail) {
      sendGrievancePdfEmail_(data, pdf);
      ss.toast('PDF emailed to ' + data.memberEmail, COMMAND_CONFIG.SYSTEM_NAME, 5);
    }

    // Open folder in new tab
    var html = HtmlService.createHtmlOutput(
      '<script>window.open(' + JSON.stringify(folder.getUrl()) + ', "_blank"); google.script.host.close();</script>'
    ).setWidth(100).setHeight(50);
    ui.showModalDialog(html, 'Opening folder...');

  } catch (e) {
    ui.alert('Error', 'Failed to create PDF: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Sends grievance PDF to member via email
 * @param {Object} data - Grievance data object with memberEmail
 * @param {File} pdf - The PDF file to attach
 * @private
 */
function sendGrievancePdfEmail_(data, pdf) {
  if (!data.memberEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.memberEmail).trim())) {
    throw new Error('Invalid or missing member email address');
  }

  var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Grievance Form - ' + data.grievanceId;

  var body = 'Dear ' + data.name + ',\n\n' +
    'Please find attached your grievance form for case ' + data.grievanceId + '.\n\n' +
    'GRIEVANCE DETAILS:\n' +
    '─────────────────────────────────\n' +
    'Grievance ID: ' + data.grievanceId + '\n' +
    'Status: ' + data.status + '\n' +
    'Articles: ' + (data.articles || 'N/A') + '\n' +
    'Unit: ' + (data.unit || 'N/A') + '\n' +
    'Location: ' + (data.location || 'N/A') + '\n' +
    'Assigned Steward: ' + (data.steward || 'N/A') + '\n' +
    '─────────────────────────────────\n\n' +
    'NEXT STEPS:\n' +
    '1. Review the attached form for accuracy\n' +
    '2. Sign where indicated\n' +
    '3. Return the signed form to your steward\n\n' +
    'If you have any questions, please contact your steward.\n' +
    COMMAND_CONFIG.EMAIL.FOOTER;

  safeSendEmail_({
    to: data.memberEmail,
    subject: subject,
    body: body,
    attachments: [pdf.getAs(MimeType.PDF)],
    name: COMMAND_CONFIG.SYSTEM_NAME || 'Union Grievance System'
  });

  // Use secureLog to mask PII in logs
  if (typeof secureLog === 'function') {
    secureLog('EmailPDF', 'Grievance PDF emailed', { recipientMasked: typeof maskEmail === 'function' ? maskEmail(data.memberEmail) : '[REDACTED]' });
  }
}

/**
 * Handles form submission for grievance intake forms
 * Creates PDF and links it back to the log
 * @param {Object} e - Form submission event object
 */
function onGrievanceFormSubmit(e) {
  try {
    var responses = e.namedValues;
    var data = {
      name: responses['Member Name'] ? responses['Member Name'][0] : 'Unknown',
      id: responses['Member ID'] ? responses['Member ID'][0] : '000',
      details: responses['Details'] ? responses['Details'][0] : 'No details provided.',
      grievanceId: responses['Grievance ID'] ? responses['Grievance ID'][0] : '',
      articles: responses['Articles'] ? responses['Articles'][0] : ''
    };

    // Create member folder and PDF
    var memberFolder = getOrCreateMemberFolder(data.name, data.id);
    createSignatureReadyPDF(memberFolder, data);

    // Link PDF back to the grievance log
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (sheet) {
      // Use the event range to find the exact row the form submission added to
      var targetRow = e.range ? e.range.getRow() : sheet.getLastRow();
      // M-42: Only write to DRIVE_FOLDER_URL if the column is currently empty,
      // to avoid overwriting an existing folder URL with the PDF URL.
      // Write the PDF URL to a separate column if available, otherwise use folder URL
      // column only as a fallback when it's empty.
      if (GRIEVANCE_COLS.DRIVE_FOLDER_URL) {
        var existingUrl = sheet.getRange(targetRow, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
        if (!existingUrl) {
          // Store the member folder URL (not the PDF URL) so the link points to the folder
          sheet.getRange(targetRow, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(memberFolder.getUrl());
        }
      }
    }

    // Use secureLog to avoid logging PII
    if (typeof secureLog === 'function') {
      secureLog('FormSubmission', 'Form submission processed - PDF created', {});
    }

  } catch (err) {
    Logger.log('Error processing form submission: ' + err.message);
  }
}

// ============================================================================
// UI DIALOGS FOR INTEGRATIONS
// ============================================================================

/**
 * Shows calendar sync dialog
 */
function showCalendarSyncDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
      <style>
        .sync-container { padding: 20px; }
        .sync-option { margin: 15px 0; padding: 15px; border: 1px solid #ddd; border-radius: 8px; }
        .sync-option h4 { margin: 0 0 8px 0; }
        .sync-option p { margin: 0; color: #666; font-size: 13px; }
        .action-buttons { margin-top: 20px; display: flex; gap: 10px; justify-content: flex-end; }
        .status { margin-top: 15px; padding: 10px; background: #f0f0f0; border-radius: 4px; display: none; }
      </style>
    </head>
    <body>
      <div class="sync-container">
        <div class="sync-option">
          <h4>Sync All Open Grievances</h4>
          <p>Creates calendar events for all deadlines of open grievances</p>
          <button class="btn btn-primary" onclick="syncAll()" style="margin-top: 10px;">
            Sync All Deadlines
          </button>
        </div>

        <div id="status" class="status"></div>

        <div class="action-buttons">
          <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
        </div>
      </div>

      <script>
        function showStatus(msg, isError) {
          const el = document.getElementById('status');
          el.style.display = 'block';
          el.style.background = isError ? '#fee2e2' : '#d1fae5';
          el.textContent = msg;
        }

        function syncAll() {
          showStatus('Syncing...', false);
          google.script.run
            .withSuccessHandler(function(r) {
              showStatus(r.success ? r.message : 'Error: ' + r.error, !r.success);
            })
            .withFailureHandler(function(err) {
              showStatus('Sync failed: ' + err.message, true);
            })
            .syncDeadlinesToCalendar();
        }
      </script>
    </body>
    </html>
  `).setWidth(450).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'Calendar Sync');
}

/**
 * Shows upcoming deadlines dialog
 */
function showUpcomingDeadlines() {
  const deadlines = getUpcomingDeadlines(14);

  let tableRows = '';
  if (deadlines.length === 0) {
    tableRows = '<tr><td colspan="4" style="text-align:center; padding:20px;">No upcoming deadlines</td></tr>';
  } else {
    deadlines.forEach(d => {
      const urgentStyle = d.daysLeft <= 3 ? 'background:#fee2e2;' : '';
      tableRows += `<tr style="${urgentStyle}">
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.grievanceId))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.memberName))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.step))}</td>
        <td style="padding:8px; border-bottom:1px solid #ddd;">${escapeHtml(String(d.date))} (${escapeHtml(String(d.daysLeft))} days)</td>
      </tr>`;
    });
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      ${getCommonStyles()}
    </head>
    <body style="padding: 20px;">
      <h3 style="margin-bottom: 15px;">Upcoming Deadlines (Next 14 Days)</h3>
      <table style="width:100%; border-collapse: collapse;">
        <tr style="background:#f0f0f0;">
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Grievance</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Member</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Step</th>
          <th style="padding:10px; text-align:left; border-bottom:2px solid #ddd;">Due</th>
        </tr>
        ${tableRows}
      </table>
      <div style="margin-top:20px; text-align:right;">
        <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
      </div>
    </body>
    </html>
  `).setWidth(600).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Upcoming Deadlines');
}
/**
 * ============================================================================
 * WEB APP DEPLOYMENT FOR MOBILE ACCESS
 * ============================================================================
 * This file enables the dashboard to be deployed as a standalone web app
 * that can be accessed directly via URL on mobile devices.
 *
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Go to Extensions → Apps Script
 * 2. Click "Deploy" → "New deployment"
 * 3. Select "Web app" as the deployment type
 * 4. Set "Execute as" to your account
 * 5. Set "Who has access" to your organization or anyone
 * 6. Click "Deploy" and copy the URL
 * 7. Bookmark this URL on your mobile device for easy access
 */

// Dead code removed: doGetLegacy() (207 lines) — replaced by web-dashboard/WebApp.gs doGet()

// ============================================================================
// BOTTOM NAV HELPERS (M-DUP-1: extracted from 8 duplicate instances)
// ============================================================================

/** Returns the shared bottom-nav CSS block used by all legacy mobile pages. */
function _getBottomNavCSS_() {
  return '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}';
}

/**
 * Builds bottom-nav HTML from a nav items array.
 * @param {string} baseUrl — web app base URL
 * @param {Array<{icon:string,label:string,page:string}>} items — nav items (page='' for home)
 * @param {string} activePage — which page to highlight as active
 * @returns {string} HTML string
 */
function _getBottomNavHTML_(baseUrl, items, activePage) {
  var html = '<nav class="bottom-nav">';
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var href = item.page ? escapeHtml(baseUrl) + '?page=' + escapeHtml(item.page) : escapeHtml(baseUrl);
    var cls = item.page === activePage ? 'nav-item active' : 'nav-item';
    html += '<a class="' + cls + '" href="' + href + '"><span class="nav-icon">' + item.icon + '</span>' + escapeHtml(item.label) + '</a>';
  }
  return html + '</nav>';
}

/** Standard steward nav: Home, Search, Cases, Members, Links */
var _STEWARD_NAV_ = [
  { icon: '\uD83D\uDCCA', label: 'Home', page: '' },
  { icon: '\uD83D\uDD0D', label: 'Search', page: 'search' },
  { icon: '\uD83D\uDCCB', label: 'Cases', page: 'grievances' },
  { icon: '\uD83D\uDC65', label: 'Members', page: 'members' },
  { icon: '\uD83D\uDD17', label: 'Links', page: 'links' }
];

/** Member nav: Home, Check In, Learn, My Info, Links */
var _MEMBER_NAV_ = [
  { icon: '\uD83D\uDCCA', label: 'Home', page: '' },
  { icon: '\u2705', label: 'Check In', page: 'checkin' },
  { icon: '\uD83D\uDCDA', label: 'Learn', page: 'resources' },
  { icon: '\uD83D\uDC64', label: 'My Info', page: 'selfservice' },
  { icon: '\uD83D\uDD17', label: 'Links', page: 'links' }
];

/** Deadline nav: Home, Cases, Deadlines, Search, Links */
var _DEADLINE_NAV_ = [
  { icon: '\uD83D\uDCCA', label: 'Home', page: '' },
  { icon: '\uD83D\uDCCB', label: 'Cases', page: 'grievances' },
  { icon: '\uD83D\uDCC5', label: 'Deadlines', page: 'deadlines' },
  { icon: '\uD83D\uDD0D', label: 'Search', page: 'search' },
  { icon: '\uD83D\uDD17', label: 'Links', page: 'links' }
];

/**
 * Returns the search page HTML for web app
 */
function getWebAppSearchHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Search - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);margin-bottom:12px;text-align:center}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:14px 14px 14px 45px;border:none;border-radius:12px;font-size:16px;background:white;-webkit-appearance:none}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#666}' +
    '.clear-btn{position:absolute;right:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#999;background:none;border:none;cursor:pointer;display:none}' +

    // Tabs
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0;position:sticky;top:76px;z-index:99}' +
    '.tab{flex:1;padding:14px;text-align:center;font-size:14px;font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:48px}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED}' +

    // Results
    '.results{padding:15px}' +
    '.result-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px}' +
    '.result-type{font-size:12px;color:#7C3AED;font-weight:600;text-transform:uppercase;margin-bottom:6px}' +
    '.result-title{font-size:17px;font-weight:600;color:#333;margin-bottom:4px}' +
    '.result-detail{font-size:14px;color:#666;margin-top:4px}' +
    '.result-badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;margin-top:8px}' +
    '.badge-open{background:#FEE2E2;color:#DC2626}' +
    '.badge-pending{background:#FEF3C7;color:#D97706}' +
    '.badge-resolved{background:#D1FAE5;color:#059669}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +
    '.empty-text{font-size:16px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav (M-DUP-1: extracted to helper)
    _getBottomNavCSS_() +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔍 Search</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search members or grievances..." oninput="handleSearch(this.value)" autocomplete="off" autocapitalize="off">' +
    '<button class="clear-btn" id="clearBtn" onclick="clearSearch()">✕</button>' +
    '</div></div>' +

    '<div class="tabs">' +
    '<button class="tab active" data-tab="all" onclick="setTab(\'all\',this)">All</button>' +
    '<button class="tab" data-tab="members" onclick="setTab(\'members\',this)">Members</button>' +
    '<button class="tab" data-tab="grievances" onclick="setTab(\'grievances\',this)">Grievances</button>' +
    '</div>' +

    '<div class="results" id="results">' +
    '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">Type to search members or grievances</div></div>' +
    '</div>' +

    // Bottom Navigation (M-DUP-1: extracted to helper)
    _getBottomNavHTML_(baseUrl, _STEWARD_NAV_, 'search') +

    '<script>' +
    getClientSideEscapeHtml() +
    'var currentTab="all";' +
    'var searchTimeout=null;' +
    'var lastQuery="";' +

    'function setTab(tab,btn){' +
    '  currentTab=tab;' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  if(lastQuery.length>=2)performSearch(lastQuery);' +
    '}' +

    'function handleSearch(q){' +
    '  lastQuery=q;' +
    '  document.getElementById("clearBtn").style.display=q?"block":"none";' +
    '  if(searchTimeout)clearTimeout(searchTimeout);' +
    '  if(!q||q.length<2){' +
    '    showEmpty("Type at least 2 characters to search");' +
    '    return;' +
    '  }' +
    '  showLoading();' +
    '  searchTimeout=setTimeout(function(){performSearch(q)},300);' +
    '}' +

    'function performSearch(q){' +
    '  google.script.run.withSuccessHandler(renderResults).withFailureHandler(showError).getWebAppSearchResults(q,currentTab);' +
    '}' +

    'function clearSearch(){' +
    '  document.getElementById("searchInput").value="";' +
    '  document.getElementById("clearBtn").style.display="none";' +
    '  lastQuery="";' +
    '  showEmpty("Type to search members or grievances");' +
    '}' +

    'function showEmpty(msg){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">🔍</div><div class=\\"empty-text\\">"+msg+"</div></div>";' +
    '}' +

    'function showLoading(){' +
    '  document.getElementById("results").innerHTML="<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:15px\\">Searching...</div></div>";' +
    '}' +

    'function showError(err){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error: "+escapeHtml(err.message||"Unknown error")+"</div></div>";' +
    '}' +

    'function getBadgeClass(status){' +
    '  if(!status)return"";' +
    '  var s=status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"badge-open";' +
    '  if(s.indexOf("pending")>=0)return"badge-pending";' +
    '  if(s.indexOf("resolved")>=0||s.indexOf("closed")>=0||s.indexOf("withdrawn")>=0)return"badge-resolved";' +
    '  return"";' +
    '}' +

    'function renderResults(data){' +
    '  var c=document.getElementById("results");' +
    '  if(!data||data.length===0){' +
    '    showEmpty("No results found");' +
    '    return;' +
    '  }' +
    '  c.innerHTML=data.map(function(r){' +
    '    var badge=r.status?"<span class=\\"result-badge "+getBadgeClass(r.status)+"\\">"+escapeHtml(r.status)+"</span>":"";' +
    '    return"<div class=\\"result-card\\">"+"<div class=\\"result-type\\">"+(r.type==="member"?"👤 Member":"📋 Grievance")+"</div>"+"<div class=\\"result-title\\">"+escapeHtml(r.title)+"</div>"+"<div class=\\"result-detail\\">"+escapeHtml(r.subtitle)+"</div>"+(r.detail?"<div class=\\"result-detail\\">"+escapeHtml(r.detail)+"</div>":"")+badge+"</div>";' +
    '  }).join("");' +
    '}' +

    // Auto-focus search on load
    'document.getElementById("searchInput").focus();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the grievance list HTML for web app (enhanced with Overdue filter, expandable details)
 */
function getWebAppGrievanceListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Grievances - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px 15px 12px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +

    // Filter pills with Overdue
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 16px;border-radius:20px;font-size:13px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +
    '.filter-pill.danger{background:#DC2626;color:white}' +
    '.filter-pill.danger.active{background:#FEE2E2;color:#DC2626}' +

    // List with expandable cards
    '.grievance-list{padding:15px}' +
    '.grievance-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.grievance-card:active{transform:scale(0.99)}' +
    '.grievance-card.overdue{border-left:4px solid #DC2626}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}' +
    '.grievance-id{font-size:15px;font-weight:700;color:#7C3AED}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600}' +
    '.status-open{background:#FEE2E2;color:#DC2626}' +
    '.status-pending{background:#FEF3C7;color:#D97706}' +
    '.status-resolved{background:#D1FAE5;color:#059669}' +
    '.status-overdue{background:#DC2626;color:white;animation:pulse 2s infinite}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}' +
    '.grievance-name{font-size:16px;font-weight:500;color:#333;margin-bottom:4px}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:4px}' +
    '.grievance-step{display:inline-block;padding:3px 8px;background:#E0E7FF;color:#4F46E5;border-radius:6px;font-size:11px;font-weight:500;margin-top:8px}' +

    // Expandable details
    '.grievance-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.grievance-card.expanded .grievance-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#333;font-weight:500}' +
    '.detail-value.danger{color:#DC2626}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Bottom nav (M-DUP-1: extracted to helper)
    _getBottomNavCSS_() +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>📋 Grievances</h2>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="open" onclick="setFilter(\'open\',this)">Open</button>' +
    '<button class="filter-pill" data-filter="pending" onclick="setFilter(\'pending\',this)">Pending</button>' +
    '<button class="filter-pill danger" data-filter="overdue" onclick="setFilter(\'overdue\',this)">⚠️ Overdue</button>' +
    '<button class="filter-pill" data-filter="resolved" onclick="setFilter(\'resolved\',this)">Resolved</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="grievance-list" id="grievanceList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading grievances...</div></div>' +
    '</div>' +

    // Bottom Navigation (M-DUP-1: extracted to helper)
    _getBottomNavHTML_(baseUrl, _STEWARD_NAV_, 'grievances') +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=25;' +
    'var currentPage=0;' +
    'var dataCache={data:null,timestamp:0};' +
    'var CACHE_TTL=300000;' +

    // Check URL for filter parameter
    'var urlParams=new URLSearchParams(window.location.search);' +
    'var initialFilter=urlParams.get("filter");' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  currentPage=0;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  if(btn)btn.classList.add("active");' +
    '  renderList();' +
    '}' +

    'function getStatusClass(g){' +
    '  if(g.isOverdue)return"status-overdue";' +
    '  if(!g.status)return"";' +
    '  var s=g.status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"status-open";' +
    '  if(s.indexOf("pending")>=0)return"status-pending";' +
    '  return"status-resolved";' +
    '}' +

    'function getStatusText(g){' +
    '  if(g.isOverdue)return"⚠️ OVERDUE";' +
    '  return g.status||"";' +
    '}' +

    'function matchesFilter(g){' +
    '  if(currentFilter==="all")return true;' +
    '  if(currentFilter==="overdue")return g.isOverdue;' +
    '  if(!g.status)return false;' +
    '  var s=g.status.toLowerCase();' +
    '  if(currentFilter==="open")return s.indexOf("open")>=0;' +
    '  if(currentFilter==="pending")return s.indexOf("pending")>=0;' +
    '  if(currentFilter==="resolved")return s.indexOf("resolved")>=0||s.indexOf("withdrawn")>=0||s.indexOf("closed")>=0;' +
    '  return true;' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function renderList(){' +
    '  var filtered=allData.filter(function(g){return matchesFilter(g)});' +
    '  var totalPages=Math.ceil(filtered.length/PAGE_SIZE);' +
    '  var start=currentPage*PAGE_SIZE;' +
    '  var paged=filtered.slice(start,start+PAGE_SIZE);' +
    '  document.getElementById("countBadge").textContent="Showing "+(start+1)+"-"+Math.min(start+PAGE_SIZE,filtered.length)+" of "+filtered.length;' +
    '  var c=document.getElementById("grievanceList");' +
    '  if(filtered.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div>No grievances found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=paged.map(function(g){' +
    '    var cardClass="grievance-card"+(g.isOverdue?" overdue":"");' +
    '    var daysInfo=g.isOverdue?"<span class=\\"detail-value danger\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?"<span class=\\"detail-value\\">"+g.daysToDeadline+" days</span>":"<span class=\\"detail-value\\">N/A</span>");' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"grievance-header\\">"+"<span class=\\"grievance-id\\">"+escapeHtml(g.id)+"</span>"+"<span class=\\"grievance-status "+getStatusClass(g)+"\\">"+getStatusText(g)+"</span>"+"</div>"+"<div class=\\"grievance-name\\">"+escapeHtml(g.name)+"</div>"+(g.category?"<div class=\\"grievance-detail\\">"+escapeHtml(g.category)+"</div>":"")+(g.step?"<span class=\\"grievance-step\\">"+escapeHtml(g.step)+"</span>":"")+"<div class=\\"grievance-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+escapeHtml(g.filedDate)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+escapeHtml(g.incidentDate)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span>"+daysInfo+"</div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+escapeHtml(g.daysOpen)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+escapeHtml(g.location)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+escapeHtml(g.articles)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+escapeHtml(g.steward)+"</span></div>"+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+escapeHtml(g.resolution)+"</span></div>":"")+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(totalPages>1){html+="<div style=\\"display:flex;justify-content:center;gap:10px;padding:15px\\"><button onclick=\\"prevPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage>0?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage===0?"disabled":"")+">← Prev</button><span style=\\"display:flex;align-items:center\\">Page "+(currentPage+1)+" of "+totalPages+"</span><button onclick=\\"nextPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage<totalPages-1?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage>=totalPages-1?"disabled":"")+">Next →</button></div>"}' +
    '  c.innerHTML=html;' +
    '}' +
    'function prevPage(){if(currentPage>0){currentPage--;renderList();window.scrollTo(0,0)}}' +
    'function nextPage(){var filtered=allData.filter(function(g){return matchesFilter(g)});if(currentPage<Math.ceil(filtered.length/PAGE_SIZE)-1){currentPage++;renderList();window.scrollTo(0,0)}}' +

    'function loadData(){' +
    '  if(!navigator.onLine){document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📡</div><div>You appear to be offline</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:10px 20px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";return}' +
    '  if(dataCache.data&&(Date.now()-dataCache.timestamp)<CACHE_TTL){console.log("Using cached data");allData=dataCache.data;if(initialFilter){currentFilter=initialFilter;var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}}renderList();return}' +
    '  console.log("Loading grievance data...");' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    console.log("Data received:",data?data.length:0,"items");' +
    '    allData=data||[];' +
    '    dataCache={data:allData,timestamp:Date.now()};' +
    '    if(initialFilter){' +
    '      currentFilter=initialFilter;' +
    '      var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");' +
    '      if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}' +
    '    }' +
    '    renderList();' +
    '  }).withFailureHandler(function(err){' +
    '    console.error("Failed to load data:",err);' +
    '    document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:10px 20px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button><div style=\\"font-size:11px;color:#999;margin-top:8px\\">"+String(err||"Unknown error")+"</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'window.addEventListener("online",loadData);' +
    'window.addEventListener("pagehide",function(){allData=null;dataCache={data:null,timestamp:0}});' +
    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the member list HTML for web app
 */
function getWebAppMemberListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Members - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:none;border-radius:12px;font-size:15px;background:white}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:18px;color:#666}' +

    // Filter pills
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:8px 0 2px;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 14px;border-radius:20px;font-size:12px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Member list
    '.member-list{padding:15px}' +
    '.member-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.member-card:active{transform:scale(0.99)}' +
    '.member-card.has-grievance{border-left:4px solid #F97316}' +
    '.member-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px}' +
    '.member-name{font-size:16px;font-weight:600;color:#333}' +
    '.member-id{font-size:12px;color:#7C3AED;font-weight:500}' +
    '.member-title{font-size:14px;color:#666;margin-bottom:4px}' +
    '.member-location{font-size:13px;color:#888}' +
    '.member-badges{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}' +
    '.badge{padding:3px 8px;border-radius:12px;font-size:11px;font-weight:500}' +
    '.badge-steward{background:#DDD6FE;color:#7C3AED}' +
    '.badge-grievance{background:#FEF3C7;color:#D97706}' +

    // Expandable details
    '.member-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.member-card.expanded .member-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:80px}' +
    '.detail-value{color:#333;font-weight:500}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav (M-DUP-1: extracted to helper)
    _getBottomNavCSS_() +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>👥 Members</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search by name, ID, title..." oninput="debounceSearch()">' +
    '</div>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="steward" onclick="setFilter(\'steward\',this)">Stewards</button>' +
    '<button class="filter-pill" data-filter="grievance" onclick="setFilter(\'grievance\',this)">With Grievance</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="member-list" id="memberList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading members...</div></div>' +
    '</div>' +

    // Bottom Navigation (M-DUP-1: extracted to helper)
    _getBottomNavHTML_(baseUrl, _STEWARD_NAV_, 'members') +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=25;' +
    'var currentPage=0;' +
    'var dataCache={data:null,timestamp:0};' +
    'var CACHE_TTL=300000;' +
    'var searchTimeout=null;' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  currentPage=0;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterMembers();' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function debounceSearch(){clearTimeout(searchTimeout);searchTimeout=setTimeout(function(){currentPage=0;filterMembers()},300)}' +

    'function filterMembers(){' +
    '  var query=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allData.filter(function(m){' +
    '    var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0||(m.title||"").toLowerCase().indexOf(query)>=0||(m.location||"").toLowerCase().indexOf(query)>=0;' +
    '    var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);' +
    '    return matchesQuery&&matchesFilter;' +
    '  });' +
    '  var totalPages=Math.ceil(filtered.length/PAGE_SIZE);' +
    '  var start=currentPage*PAGE_SIZE;' +
    '  var paged=filtered.slice(start,start+PAGE_SIZE);' +
    '  document.getElementById("countBadge").textContent="Showing "+(filtered.length>0?(start+1)+"-"+Math.min(start+PAGE_SIZE,filtered.length):"0")+" of "+filtered.length;' +
    '  renderList(paged,totalPages);' +
    '}' +

    'function renderList(data,totalPages){' +
    '  var c=document.getElementById("memberList");' +
    '  if(!data||data.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div>No members found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=data.map(function(m){' +
    '    var cardClass="member-card"+(m.hasOpenGrievance?" has-grievance":"");' +
    '    var badges="";' +
    '    if(m.isSteward)badges+="<span class=\\"badge badge-steward\\">🛡️ Steward</span>";' +
    '    if(m.hasOpenGrievance)badges+="<span class=\\"badge badge-grievance\\">⚠️ Open Grievance</span>";' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"member-header\\"><span class=\\"member-name\\">"+escapeHtml(m.name)+"</span><span class=\\"member-id\\">"+escapeHtml(m.id)+"</span></div>"+"<div class=\\"member-title\\">"+escapeHtml(m.title)+"</div>"+"<div class=\\"member-location\\">📍 "+escapeHtml(m.location)+"</div>"+(badges?"<div class=\\"member-badges\\">"+badges+"</div>":"")+"<div class=\\"member-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+escapeHtml(m.email||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📞 Phone:</span><span class=\\"detail-value\\">"+escapeHtml(m.phone||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+escapeHtml(m.unit)+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">👔 Supervisor:</span><span class=\\"detail-value\\">"+escapeHtml(m.supervisor)+"</span></div>"+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(totalPages>1){html+="<div style=\\"display:flex;justify-content:center;gap:10px;padding:15px\\"><button onclick=\\"prevPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage>0?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage===0?"disabled":"")+">← Prev</button><span style=\\"display:flex;align-items:center\\">Page "+(currentPage+1)+" of "+totalPages+"</span><button onclick=\\"nextPage()\\" style=\\"padding:10px 20px;border:none;border-radius:8px;background:"+(currentPage<totalPages-1?"#7C3AED":"#ccc")+";color:white;cursor:pointer\\" "+(currentPage>=totalPages-1?"disabled":"")+">Next →</button></div>"}' +
    '  c.innerHTML=html;' +
    '}' +
    'function prevPage(){if(currentPage>0){currentPage--;filterMembers();window.scrollTo(0,0)}}' +
    'function nextPage(){var query=(document.getElementById("searchInput").value||"").toLowerCase();var filtered=allData.filter(function(m){var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0;var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);return matchesQuery&&matchesFilter});if(currentPage<Math.ceil(filtered.length/PAGE_SIZE)-1){currentPage++;filterMembers();window.scrollTo(0,0)}}' +

    'function loadData(){' +
    '  if(!navigator.onLine){document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📡</div><div>You appear to be offline</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:8px 16px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";return}' +
    '  if(dataCache.data&&(Date.now()-dataCache.timestamp)<CACHE_TTL){allData=dataCache.data;filterMembers();return}' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allData=data||[];' +
    '    dataCache={data:allData,timestamp:Date.now()};' +
    '    filterMembers();' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><button onclick=\\"loadData()\\" style=\\"margin-top:12px;padding:8px 16px;background:#7C3AED;color:white;border:none;border-radius:8px;cursor:pointer\\">Retry</button></div>";' +
    '  }).getWebAppMemberList();' +
    '}' +

    'window.addEventListener("online",loadData);' +
    'window.addEventListener("pagehide",function(){allData=null;dataCache={data:null,timestamp:0}});' +
    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the links/resources page HTML for web app
 */
function getWebAppLinksHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Links - Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px)}' +
    '.header .subtitle{font-size:13px;opacity:0.9;margin-top:5px}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Link cards
    '.link-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}' +
    '.link-card{background:white;padding:20px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-decoration:none;display:flex;flex-direction:column;align-items:center;text-align:center;gap:10px;transition:all 0.2s}' +
    '.link-card:active{transform:scale(0.96);background:#f8f4ff}' +
    '.link-icon{font-size:32px}' +
    '.link-label{font-weight:600;color:#333;font-size:14px}' +
    '.link-desc{font-size:12px;color:#666}' +

    // Full-width link
    '.link-card.full{grid-column:span 2;flex-direction:row;padding:16px;justify-content:flex-start;text-align:left}' +
    '.link-card.full .link-icon{font-size:28px}' +
    '.link-card.full .link-content{flex:1}' +

    // GitHub special styling
    '.link-card.github{background:linear-gradient(135deg,#24292e,#1a1e22);color:white}' +
    '.link-card.github .link-label{color:white}' +
    '.link-card.github .link-desc{color:rgba(255,255,255,0.7)}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav (M-DUP-1: extracted to helper)
    _getBottomNavCSS_() +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔗 Links & Resources</h2>' +
    '<div class="subtitle">Quick access to forms and tools</div>' +
    '</div>' +

    '<div class="container" id="linksContent">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading links...</div></div>' +
    '</div>' +

    // Bottom Navigation (M-DUP-1: extracted to helper)
    _getBottomNavHTML_(baseUrl, _STEWARD_NAV_, 'links') +

    '<script>' +
    getClientSideEscapeHtml() +
    'function safeUrl(u){if(!u)return"#";u=String(u);return/^https?:\\/\\//i.test(u)?u:"#";}' +
    'function loadLinks(){' +
    '  google.script.run.withSuccessHandler(function(links){' +
    '    renderLinks(links);' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("linksContent").innerHTML="<div class=\\"loading\\">⚠️ Error loading links</div>";' +
    '  }).getWebAppResourceLinks();' +
    '}' +

    'function renderLinks(links){' +
    '  var html="";' +

    '  // Forms section' +
    '  html+="<div class=\\"section-title\\">📝 Forms</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  if(links.grievanceForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.grievanceForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📋</span><span class=\\"link-label\\">Grievance Form</span><span class=\\"link-desc\\">File a grievance</span></a>";}' +
    '  if(links.contactForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.contactForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">✉️</span><span class=\\"link-label\\">Contact Form</span><span class=\\"link-desc\\">Send a message</span></a>";}' +
    '  if(links.satisfactionForm){html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.satisfactionForm))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Satisfaction Survey</span><span class=\\"link-desc\\">Give feedback</span></a>";}' +
    '  if(!links.grievanceForm&&!links.contactForm&&!links.satisfactionForm){html+="<div class=\\"link-card full\\"><span class=\\"link-icon\\">ℹ️</span><div class=\\"link-content\\"><span class=\\"link-label\\">No Forms Configured</span><span class=\\"link-desc\\">Add form URLs to Config sheet</span></div></div>";}' +
    '  html+="</div>";' +

    '  // Resources section' +
    '  html+="<div class=\\"section-title\\">🔧 Resources</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  html+="<a class=\\"link-card\\" href=\\""+escapeHtml(safeUrl(links.spreadsheetUrl))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Spreadsheet</span><span class=\\"link-desc\\">Open full dashboard</span></a>";' +
    '  html+="<a class=\\"link-card github\\" href=\\""+escapeHtml(safeUrl(links.githubRepo))+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📦</span><span class=\\"link-label\\">GitHub Repo</span><span class=\\"link-desc\\">Source code</span></a>";' +
    '  html+="</div>";' +

    '  document.getElementById("linksContent").innerHTML=html;' +
    '}' +

    'loadLinks();' +
    '</script>' +

    '</body></html>';
}

/**
 * API function to get search results for web app
 * @param {string} query - Search query
 * @param {string} tab - Tab filter (all, members, grievances)
 * @returns {Array} Search results
 */
function getWebAppSearchResults(query, tab) {
  return getMobileSearchData(query, tab);
}

/**
 * API function to get grievance list for web app (full fields like Interactive Dashboard)
 * @returns {Array} Grievance data with all fields
 */
function getWebAppGrievanceList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppGrievanceList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) {
      Logger.log('getWebAppGrievanceList: Grievance Log sheet not found');
      return [];
    }

    ensureMinimumColumns(sheet, getGrievanceHeaders().length);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppGrievanceList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!isGrievanceId_(grievanceId)) return null;

      var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
      var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

      return {
        id: grievanceId,
        memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
        name: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
        status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
        step: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
        category: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
        articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
        filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
        incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
        nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
        daysToDeadline: daysToDeadline,
        isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
        daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
        location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
        steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
        resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
      };
    }).filter(function(g) { return g !== null; }).slice(0, 100);

    Logger.log('getWebAppGrievanceList: Returning ' + result.length + ' grievances');
    return result;
  } catch (e) {
    Logger.log('getWebAppGrievanceList error: ' + e.toString());
    throw new Error('Failed to load grievances: ' + e.message);
  }
}

/**
 * API function to get member list for web app
 * @returns {Array} Member data
 */
function getWebAppMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppMemberList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) {
      Logger.log('getWebAppMemberList: Member Directory sheet not found');
      return [];
    }

    ensureMinimumColumns(sheet, getMemberHeaders().length);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppMemberList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();

    var result = data.map(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!isMemberId_(memberId)) return null;

      return {
        id: memberId,
        firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
        name: ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim(),
        title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
        location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
        unit: row[MEMBER_COLS.UNIT - 1] || 'N/A',
        email: row[MEMBER_COLS.EMAIL - 1] || '',
        phone: row[MEMBER_COLS.PHONE - 1] || '',
        isSteward: isTruthyValue(row[MEMBER_COLS.IS_STEWARD - 1]),
        supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
        hasOpenGrievance: isTruthyValue(row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1])
      };
    }).filter(function(m) { return m !== null; }).slice(0, 100);

    Logger.log('getWebAppMemberList: Returning ' + result.length + ' members');
    return result;
  } catch (e) {
    Logger.log('getWebAppMemberList error: ' + e.toString());
    throw new Error('Failed to load members: ' + e.message);
  }
}

/**
 * API function to get resource links for web app
 * @returns {Object} Resource links
 */
function getWebAppResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: '',
    orgWebsite: '',
    githubRepo: '',
    resourcesFolderUrl: ''  // v4.20.18: Drive Resources/ folder URL for steward uploads
  };

  if (!ss) return links;
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: '',
    githubRepo: '',
    resourcesFolderUrl: ''
  };

  // Get form URLs from Config sheet using CONFIG_COLS constants (data in row 3)
  if (configSheet && configSheet.getLastRow() >= 3) {
    try {
      // satisfactionForm removed v4.22.7 — survey is native webapp
      var configRow = configSheet.getRange(3, 1, 1, CONFIG_COLS.ORG_WEBSITE).getValues()[0];
      links.grievanceForm = configRow[CONFIG_COLS.GRIEVANCE_FORM_URL - 1] || '';
      links.contactForm = configRow[CONFIG_COLS.CONTACT_FORM_URL - 1] || '';
      links.orgWebsite = configRow[CONFIG_COLS.ORG_WEBSITE - 1] || '';
    } catch (_e) {
      // Ignore errors reading config
    }
  }

  // Resolve Resources/ folder URL from stored ID (v4.20.18)
  try {
    var resFolderId = (typeof getConfigValue_ === 'function' && CONFIG_COLS.RESOURCES_FOLDER_ID)
      ? getConfigValue_(CONFIG_COLS.RESOURCES_FOLDER_ID)
      : '';
    if (!resFolderId) {
      resFolderId = PropertiesService.getScriptProperties().getProperty('RESOURCES_FOLDER_ID') || '';
    }
    if (resFolderId) {
      links.resourcesFolderUrl = DriveApp.getFolderById(resFolderId).getUrl();
    }
  } catch (_re) {}

  return links;
}

/**
 * API function to get dashboard stats with win rate for web app
 * @returns {Object} Dashboard statistics
 */
function getWebAppDashboardStats() {
  var stats = getMobileDashboardStats();

  // Calculate win rate
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return stats;
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (sheet && sheet.getLastRow() > 1) {
    var resolutions = sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, sheet.getLastRow() - 1, 1).getValues();
    var won = 0, total = 0;
    resolutions.forEach(function(row) {
      var res = (row[0] || '').toString().toLowerCase();
      if (res) {
        total++;
        if (res.indexOf('won') >= 0 || res.indexOf('favorable') >= 0) {
          won++;
        }
      }
    });
    stats.winRate = total > 0 ? Math.round((won / total) * 100) + '%' : 'N/A';
  } else {
    stats.winRate = 'N/A';
  }

  // Also get total members
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
    var validMembers = memberIds.filter(function(row) {
      var id = row[0] || '';
      return isMemberId_(id);
    }).length;
    stats.totalMembers = validMembers;
  } else {
    stats.totalMembers = 0;
  }

  return stats;
}

/**
 * Menu function to show the deployed mobile dashboard URL
 */
function showWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var url = ScriptApp.getService().getUrl();

  if (url) {
    ui.alert(
      '📱 Mobile Dashboard URL',
      'Your mobile dashboard URL:\n\n' + url + '\n\n' +
      'Open this URL on your phone and add it to your home screen for quick access!',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '📱 Mobile Dashboard URL',
      'No web app deployment found.\n\n' +
      'To deploy:\n' +
      '1. Go to Extensions → Apps Script\n' +
      '2. Click "Deploy" → "New deployment"\n' +
      '3. Select type: "Web app"\n' +
      '4. Set "Who has access" to your preference\n' +
      '5. Click "Deploy" and copy the URL',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Add Mobile Dashboard link to Config sheet for easy mobile access
 * Creates a clickable hyperlink cell that works on mobile devices
 */
function addMobileDashboardLinkToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    ui.alert('Error', 'Config sheet not found', ui.ButtonSet.OK);
    return;
  }

  // Prompt user to enter the URL from Manage deployments
  var response = ui.prompt(
    '📱 Add Mobile Dashboard Link',
    'To get your web app URL:\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click "Deploy" → "Manage deployments"\n' +
    '3. Copy the Web app URL\n\n' +
    'Paste the URL below:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var url = response.getResponseText().trim();
  if (!url || !url.match(/^https:\/\/script\.google\.com/)) {
    ui.alert(
      'Invalid URL',
      'Please enter a valid Google Apps Script web app URL.\n\n' +
      'It should look like:\nhttps://script.google.com/macros/s/XXXXX/exec',
      ui.ButtonSet.OK
    );
    return;
  }

  // Use the canonical column constant — never hardcode column numbers.
  var targetCol = CONFIG_COLS.MOBILE_DASHBOARD_URL;

  // Write data starting at row 3 (first data row).
  // Rows 1-2 are section/column headers managed by createConfigSheet — don't touch them.
  var linkCell = configSheet.getRange(3, targetCol);
  linkCell.setFormula('=HYPERLINK(' + JSON.stringify(url) + ', "📱 Tap to Open Dashboard")');
  linkCell.setFontSize(14);
  linkCell.setFontWeight('bold');
  linkCell.setFontColor('#1a73e8');
  linkCell.setBackground('#e8f0fe');

  // Also add plain URL below for copying
  var urlCell = configSheet.getRange(4, targetCol);
  urlCell.setValue(url);
  urlCell.setFontSize(10);
  urlCell.setWrap(true);

  // Add instructions
  var instructionCell = configSheet.getRange(5, targetCol);
  instructionCell.setValue('Open Google Sheets on your phone, navigate to Config tab, and tap the blue link above to access the dashboard.');
  instructionCell.setFontSize(9);
  instructionCell.setFontColor('#666666');
  instructionCell.setWrap(true);

  // Set column width
  configSheet.setColumnWidth(targetCol, 300);

  SpreadsheetApp.getUi().alert(
    '📱 Mobile Dashboard Link Added!',
    'A clickable link has been added to the "📱 Mobile Dashboard URL" column of the Config sheet.\n\n' +
    'To access on mobile:\n' +
    '1. Open this spreadsheet in Google Sheets mobile app\n' +
    '2. Go to the Config tab\n' +
    '3. Scroll to the Mobile Dashboard section\n' +
    '4. Tap the blue "Tap to Open Dashboard" link\n\n' +
    'URL: ' + url,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// CONSTANT CONTACT V3 API INTEGRATION — Read-Only Engagement Metrics
// ============================================================================

/**
 * Constant Contact API configuration
 * @const {Object}
 */
var CC_CONFIG = {
  API_BASE: 'https://api.cc.email/v3',
  AUTH_URL: 'https://authz.constantcontact.com/oauth2/default/v1/authorize',
  TOKEN_URL: 'https://authz.constantcontact.com/oauth2/default/v1/token',
  CONTACTS_ENDPOINT: '/contacts',
  ACTIVITY_SUMMARY_ENDPOINT: '/reports/contact_reports/{contact_id}/activity_summary',
  RATE_LIMIT_PER_SECOND: 4,
  RATE_LIMIT_DELAY_MS: 300,
  PAGE_LIMIT: 500,
  ACTIVITY_LOOKBACK_DAYS: 365,
  PROP_API_KEY: 'CC_API_KEY',
  PROP_API_SECRET: 'CC_API_SECRET',
  PROP_ACCESS_TOKEN: 'CC_ACCESS_TOKEN',
  PROP_REFRESH_TOKEN: 'CC_REFRESH_TOKEN',
  PROP_TOKEN_EXPIRY: 'CC_TOKEN_EXPIRY'
};

/**
 * Shows dialog for entering Constant Contact API credentials.
 * Stores API key and secret in Script Properties (encrypted at rest).
 * One-time setup required before syncing engagement metrics.
 */
function showConstantContactSetup() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var existingKey = props.getProperty(CC_CONFIG.PROP_API_KEY);

  var statusMsg = existingKey
    ? 'Current status: API key is configured (ends in ...' + existingKey.slice(-4) + ')\n\n'
    : 'Current status: Not configured\n\n';

  var response = ui.prompt(
    '📧 Constant Contact Setup',
    statusMsg +
    'Enter your Constant Contact API key (client ID).\n\n' +
    'To get your API key:\n' +
    '1. Go to app.constantcontact.com/pages/dma/portal\n' +
    '2. Click "New Application"\n' +
    '3. Copy your API Key (Client ID)\n\n' +
    'API Key:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText().trim()) {
    return;
  }

  var apiKey = response.getResponseText().trim();

  var secretResponse = ui.prompt(
    '📧 Constant Contact Setup (Step 2)',
    'Enter your Client Secret.\n\n' +
    'Click "Generate Client Secret" in your CC app settings.\n' +
    'Important: Copy it immediately — it only appears once.\n\n' +
    'Client Secret:',
    ui.ButtonSet.OK_CANCEL
  );

  if (secretResponse.getSelectedButton() !== ui.Button.OK || !secretResponse.getResponseText().trim()) {
    return;
  }

  var apiSecret = secretResponse.getResponseText().trim();

  props.setProperty(CC_CONFIG.PROP_API_KEY, apiKey);
  props.setProperty(CC_CONFIG.PROP_API_SECRET, apiSecret);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'API credentials saved. Next step: authorize your account.',
    '✅ Constant Contact Configured', 5
  );

  Logger.log('Constant Contact API credentials configured');
}

/**
 * Initiates the OAuth2 authorization flow for Constant Contact.
 * Opens a dialog with the authorization URL for the user to grant access.
 * After granting access, the user pastes back the authorization code.
 */
function authorizeConstantContact() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);

  if (!apiKey) {
    ui.alert('⚠️ Setup Required',
      'Please run "Setup API Credentials" first.',
      ui.ButtonSet.OK);
    return;
  }

  // Check if already authorized
  var existingToken = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  if (existingToken && expiry && new Date(expiry) > new Date()) {
    var reauth = ui.alert('Already Authorized',
      'You already have a valid access token.\nExpires: ' + new Date(expiry).toLocaleString() +
      '\n\nRe-authorize?',
      ui.ButtonSet.YES_NO);
    if (reauth !== ui.Button.YES) return;
  }

  // Build authorization URL — using server flow (authorization code grant)
  var redirectUri = 'https://localhost';
  var scope = 'contact_data offline_access';
  var state = Utilities.getUuid();

  var authUrl = CC_CONFIG.AUTH_URL +
    '?response_type=code' +
    '&client_id=' + encodeURIComponent(apiKey) +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=' + encodeURIComponent(scope) +
    '&state=' + encodeURIComponent(state);

  // Show the URL for the user to visit
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px;max-width:500px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.step{background:#f0f7ff;padding:15px;margin:12px 0;border-radius:8px;border-left:4px solid #1a73e8}' +
    '.step-num{font-weight:bold;color:#1a73e8}' +
    'a{color:#1a73e8;word-break:break-all}' +
    'input{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;margin:8px 0;font-size:14px}' +
    'button{padding:12px 24px;background:#1a73e8;color:white;border:none;border-radius:4px;cursor:pointer;font-size:14px}' +
    'button:hover{background:#1557b0}' +
    '.note{font-size:12px;color:#666;margin-top:12px}' +
    '</style>' +
    '<div class="container">' +
    '<h2>Authorize Constant Contact</h2>' +
    '<div class="step"><span class="step-num">Step 1:</span> Click the link below to authorize:<br><br>' +
    '<a href="' + authUrl + '" target="_blank">Open Constant Contact Authorization</a></div>' +
    '<div class="step"><span class="step-num">Step 2:</span> Log in and click "Allow"</div>' +
    '<div class="step"><span class="step-num">Step 3:</span> You\'ll be redirected to a URL like:<br>' +
    '<code>https://localhost?code=XXXX&state=...</code><br><br>' +
    'Copy the entire URL from your browser address bar and paste it below:<br>' +
    '<input type="text" id="callbackUrl" placeholder="Paste the full redirect URL here...">' +
    '<button onclick="submitCode()">Submit</button></div>' +
    '<div class="note">Your browser may show an error page — that\'s expected. ' +
    'Just copy the URL from the address bar.</div>' +
    '</div>' +
    '<script>' +
    'function submitCode(){' +
    '  var url=document.getElementById("callbackUrl").value.trim();' +
    '  if(!url){alert("Please paste the redirect URL");return;}' +
    '  var match=url.match(/[?&]code=([^&]+)/);' +
    '  if(!match){alert("Could not find authorization code in that URL. Make sure you copied the full URL.");return;}' +
    '  google.script.run' +
    '    .withSuccessHandler(function(msg){' +
    '      var c=document.querySelector(".container");c.innerHTML="";var h=document.createElement("h2");h.textContent="\\u2705 "+msg;var p=document.createElement("p");p.textContent="You can close this dialog.";c.appendChild(h);c.appendChild(p);' +
    '    })' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);})' +
    '    .exchangeConstantContactCode(match[1]);' +
    '}' +
    '</script>'
  ).setWidth(550).setHeight(520);

  ui.showModalDialog(html, '📧 Authorize Constant Contact');
}

/**
 * Exchanges an authorization code for access and refresh tokens.
 * Called from the authorization dialog after user grants access.
 * @param {string} code - The authorization code from the OAuth callback
 * @returns {string} Success message
 */
function exchangeConstantContactCode(code) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var apiSecret = props.getProperty(CC_CONFIG.PROP_API_SECRET);

  if (!apiKey || !apiSecret) {
    throw new Error('API credentials not configured. Run setup first.');
  }

  var payload = {
    grant_type: 'authorization_code',
    code: code,
    redirect_uri: 'https://localhost'
  };

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret)
    },
    payload: Object.keys(payload).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(payload[k]);
    }).join('&'),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(CC_CONFIG.TOKEN_URL, options);
    var responseCode = response.getResponseCode();
    var body = JSON.parse(response.getContentText());

    if (responseCode !== 200) {
      var errorMsg = body.error_description || body.error || 'Unknown error';
      throw new Error('Token exchange failed (' + responseCode + '): ' + errorMsg);
    }

    // Store tokens
    storeConstantContactTokens_(body);

    Logger.log('Constant Contact authorized successfully');
    return 'Authorized Successfully!';
  } catch (e) {
    Logger.log('CC token exchange error: ' + e.message);
    throw e;
  }
}

/**
 * Stores OAuth tokens from a token response.
 * @param {Object} tokenResponse - The token endpoint response
 * @private
 */
function storeConstantContactTokens_(tokenResponse) {
  var props = PropertiesService.getScriptProperties();

  props.setProperty(CC_CONFIG.PROP_ACCESS_TOKEN, tokenResponse.access_token);

  if (tokenResponse.refresh_token) {
    props.setProperty(CC_CONFIG.PROP_REFRESH_TOKEN, tokenResponse.refresh_token);
  }

  // Calculate expiry (tokens last ~2 hours; subtract 5 min buffer)
  var expiresIn = (tokenResponse.expires_in || 7200) - 300;
  var expiry = new Date(Date.now() + expiresIn * 1000).toISOString();
  props.setProperty(CC_CONFIG.PROP_TOKEN_EXPIRY, expiry);
}

/**
 * Proactive Constant Contact token health check.
 * Validates that the stored token is still valid and shows a toast if not.
 * Called from onOpenDeferred_() so the spreadsheet opens without blocking.
 */
function checkConstantContactHealth() {
  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
    if (!token) return; // CC not configured — nothing to check

    var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
    if (expiry && new Date(expiry) > new Date()) return; // Still valid

    // Token expired — try refresh
    var refreshed = getConstantContactToken_();
    if (!refreshed) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Constant Contact token expired. Re-authorize via Admin > Constant Contact.',
        'CC Token Warning',
        10
      );
    }
  } catch (_e) {
    // Non-blocking — swallow errors
    Logger.log('CC health check skipped: ' + _e.message);
  }
}

/**
 * Gets a valid access token, refreshing if expired.
 * @returns {string|null} The access token, or null if not authorized
 * @private
 */
function getConstantContactToken_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  var refreshToken = props.getProperty(CC_CONFIG.PROP_REFRESH_TOKEN);

  if (!token) return null;

  // Check if token is still valid
  if (expiry && new Date(expiry) > new Date()) {
    return token;
  }

  // Token expired — try to refresh
  if (!refreshToken) return null;

  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var apiSecret = props.getProperty(CC_CONFIG.PROP_API_SECRET);

  if (!apiKey || !apiSecret) return null;

  var payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken
  };

  var options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(apiKey + ':' + apiSecret)
    },
    payload: Object.keys(payload).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(payload[k]);
    }).join('&'),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(CC_CONFIG.TOKEN_URL, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('CC token refresh failed: ' + response.getContentText());
      return null;
    }

    var body = JSON.parse(response.getContentText());
    storeConstantContactTokens_(body);
    return body.access_token;
  } catch (e) {
    Logger.log('CC token refresh error: ' + e.message);
    return null;
  }
}

/**
 * Makes an authenticated GET request to the Constant Contact v3 API.
 * Handles token refresh and rate limiting.
 * @param {string} endpoint - The API endpoint path (e.g., '/contacts')
 * @param {Object} [params] - Optional query parameters
 * @returns {Object|null} Parsed JSON response, or null on failure
 * @private
 */
function ccApiGet_(endpoint, params) {
  var token = getConstantContactToken_();
  if (!token) return null;

  var url = CC_CONFIG.API_BASE + endpoint;

  if (params) {
    var queryParts = [];
    for (var key in params) {
      if (params.hasOwnProperty(key) && params[key] !== undefined && params[key] !== null) {
        queryParts.push(encodeURIComponent(key) + '=' + encodeURIComponent(params[key]));
      }
    }
    if (queryParts.length > 0) {
      url += '?' + queryParts.join('&');
    }
  }

  var options = {
    method: 'get',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 401) {
      // Token may have just expired — force refresh and retry once
      var props = PropertiesService.getScriptProperties();
      props.deleteProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
      token = getConstantContactToken_();
      if (!token) return null;

      options.headers['Authorization'] = 'Bearer ' + token;
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
    }

    if (code === 429) {
      // Rate limited — wait and retry once
      Utilities.sleep(1000);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
    }

    if (code !== 200) {
      Logger.log('CC API error (' + code + '): ' + response.getContentText());
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('CC API request error: ' + e.message);
    return null;
  }
}

/**
 * Fetches all contacts from Constant Contact, handling pagination.
 * Returns a map of lowercase email -> contact_id for matching.
 * @returns {Object} Map of email -> contact_id
 * @private
 */
function fetchCCContacts_() {
  var emailToContactId = {};
  var endpoint = CC_CONFIG.CONTACTS_ENDPOINT;
  var params = {
    limit: CC_CONFIG.PAGE_LIMIT,
    include: 'email_address'
  };

  var pageCount = 0;
  var maxPages = 20; // Safety limit

  while (pageCount < maxPages) {
    var data = ccApiGet_(endpoint, params);
    if (!data || !data.contacts) break;

    for (var i = 0; i < data.contacts.length; i++) {
      var contact = data.contacts[i];
      var contactId = contact.contact_id;
      var emailAddresses = contact.email_address;

      if (emailAddresses && emailAddresses.address) {
        emailToContactId[emailAddresses.address.toLowerCase()] = contactId;
      }
    }

    // Check for next page
    if (data._links && data._links.next && data._links.next.href) {
      // The next link is a full path — extract just the path + query
      var nextUrl = data._links.next.href;
      // Parse the path portion
      var pathMatch = nextUrl.match(/\/v3(\/contacts.*)/);
      if (pathMatch) {
        endpoint = pathMatch[1].split('?')[0];
        // Parse query params from the next URL
        var queryString = pathMatch[1].split('?')[1] || '';
        params = {};
        queryString.split('&').forEach(function(part) {
          var kv = part.split('=');
          if (kv.length === 2) {
            params[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1]);
          }
        });
      } else {
        break;
      }
    } else {
      break;
    }

    pageCount++;
    Utilities.sleep(CC_CONFIG.RATE_LIMIT_DELAY_MS);
  }

  return emailToContactId;
}

/**
 * Fetches engagement metrics (activity summary) for a single CC contact.
 * @param {string} contactId - The CC contact UUID
 * @returns {Object|null} Object with openRate and lastActivityDate, or null
 * @private
 */
function fetchCCContactEngagement_(contactId) {
  var now = new Date();
  var start = new Date(now.getTime() - CC_CONFIG.ACTIVITY_LOOKBACK_DAYS * 24 * 60 * 60 * 1000);

  var endpoint = CC_CONFIG.ACTIVITY_SUMMARY_ENDPOINT.replace('{contact_id}', contactId);
  var params = {
    start: start.toISOString(),
    end: now.toISOString()
  };

  var data = ccApiGet_(endpoint, params);
  if (!data || !data.campaign_activities) {
    return null;
  }

  var totalSends = 0;
  var totalOpens = 0;
  var lastActivityDate = null;

  for (var i = 0; i < data.campaign_activities.length; i++) {
    var activity = data.campaign_activities[i];

    totalSends += (activity.em_sends || 0);
    totalOpens += (activity.em_opens || 0);

    // Track the most recent activity date
    var dates = [activity.em_sends_date, activity.em_opens_date, activity.em_clicks_date]
      .filter(function(d) { return d; });

    for (var j = 0; j < dates.length; j++) {
      var d = new Date(dates[j]);
      if (!isNaN(d.getTime()) && (!lastActivityDate || d > lastActivityDate)) {
        lastActivityDate = d;
      }
    }
  }

  var openRate = totalSends > 0 ? Math.round((totalOpens / totalSends) * 100) : 0;

  return {
    openRate: openRate,
    lastActivityDate: lastActivityDate
  };
}

/**
 * Syncs Constant Contact email engagement metrics to the Member Directory.
 * Read-only: pulls open rates and last activity dates from CC, never writes to CC.
 *
 * Matches CC contacts to members by email address (case-insensitive).
 * Updates OPEN_RATE and RECENT_CONTACT_DATE columns.
 *
 * Call from menu: Admin > Data Sync > Sync CC Engagement
 */
function syncConstantContactEngagement() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    ui.alert('❌ Error', 'Member Directory sheet not found.', ui.ButtonSet.OK);
    return;
  }

  // Check authorization
  var token = getConstantContactToken_();
  if (!token) {
    ui.alert('⚠️ Not Authorized',
      'Constant Contact is not connected.\n\n' +
      'Go to Admin > Data Sync > Constant Contact Setup to configure your API credentials, ' +
      'then use "Authorize Constant Contact" to connect your account.',
      ui.ButtonSet.OK);
    return;
  }

  ss.toast('Fetching contacts from Constant Contact...', '📧 CC Sync', 10);

  // Step 1: Fetch all CC contacts (email -> contact_id map)
  var emailToContactId = fetchCCContacts_();
  var ccContactCount = Object.keys(emailToContactId).length;

  if (ccContactCount === 0) {
    ui.alert('⚠️ No Contacts Found',
      'No contacts were returned from Constant Contact.\n\n' +
      'Make sure your CC account has contacts and your API key has the correct permissions.',
      ui.ButtonSet.OK);
    return;
  }

  ss.toast('Found ' + ccContactCount + ' CC contacts. Matching to members...', '📧 CC Sync', 10);

  // Step 2: Read Member Directory emails
  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('No members found in Member Directory.', '⚠️ CC Sync', 3);
    return;
  }

  var memberData = memberSheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.RECENT_CONTACT_DATE).getValues();

  // Step 3: Match members to CC contacts and fetch engagement
  var openRateUpdates = [];
  var contactDateUpdates = [];
  var matchCount = 0;
  var processedCount = 0;

  for (var i = 0; i < memberData.length; i++) {
    var memberEmail = (memberData[i][MEMBER_COLS.EMAIL - 1] || '').toString().toLowerCase().trim();
    var contactId = memberEmail ? emailToContactId[memberEmail] : null;

    if (contactId) {
      matchCount++;

      // Rate limit: pause between API calls
      if (processedCount > 0 && processedCount % CC_CONFIG.RATE_LIMIT_PER_SECOND === 0) {
        Utilities.sleep(1000);
      }

      var engagement = fetchCCContactEngagement_(contactId);
      processedCount++;

      if (engagement) {
        openRateUpdates.push([engagement.openRate]);
        contactDateUpdates.push([engagement.lastActivityDate || '']);
      } else {
        openRateUpdates.push([numericField_(memberData[i][MEMBER_COLS.OPEN_RATE - 1])]);
        contactDateUpdates.push([memberData[i][MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '']);
      }

      // Progress update every 25 contacts
      if (processedCount % 25 === 0) {
        ss.toast('Processing... ' + processedCount + '/' + matchCount + ' contacts', '📧 CC Sync', 5);
      }
    } else {
      // No CC match — preserve existing values
      openRateUpdates.push([numericField_(memberData[i][MEMBER_COLS.OPEN_RATE - 1])]);
      contactDateUpdates.push([memberData[i][MEMBER_COLS.RECENT_CONTACT_DATE - 1] || '']);
    }
  }

  // Step 4: Write updates to Member Directory
  if (openRateUpdates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.OPEN_RATE, openRateUpdates.length, 1).setValues(openRateUpdates);
    memberSheet.getRange(2, MEMBER_COLS.RECENT_CONTACT_DATE, contactDateUpdates.length, 1).setValues(contactDateUpdates);
  }

  // Step 5: Report results
  var summary = 'Constant Contact Sync Complete!\n\n' +
    '• CC contacts found: ' + ccContactCount + '\n' +
    '• Members matched by email: ' + matchCount + '\n' +
    '• Engagement data updated: ' + processedCount + '\n' +
    '• Members without CC match: ' + (memberData.length - matchCount);

  ui.alert('✅ CC Sync Complete', summary, ui.ButtonSet.OK);
  Logger.log('CC sync: ' + matchCount + ' matches out of ' + memberData.length + ' members');
}

/**
 * Shows the current Constant Contact connection status.
 * Displays API key info, token status, and last sync details.
 */
function showConstantContactStatus() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var apiKey = props.getProperty(CC_CONFIG.PROP_API_KEY);
  var token = props.getProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  var expiry = props.getProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);
  var refreshToken = props.getProperty(CC_CONFIG.PROP_REFRESH_TOKEN);

  var status = '📧 Constant Contact Connection Status\n\n';

  if (!apiKey) {
    status += '❌ API Key: Not configured\n';
    status += '❌ Authorization: Not connected\n\n';
    status += 'To get started:\n';
    status += '1. Admin > Data Sync > CC Setup: API Credentials\n';
    status += '2. Admin > Data Sync > CC Authorize Account';
  } else {
    status += '✅ API Key: Configured (ends in ...' + apiKey.slice(-4) + ')\n';

    if (token && expiry) {
      var expiryDate = new Date(expiry);
      if (expiryDate > new Date()) {
        status += '✅ Access Token: Valid (expires ' + expiryDate.toLocaleString() + ')\n';
      } else {
        status += '⚠️ Access Token: Expired\n';
      }
    } else {
      status += '❌ Access Token: Not authorized\n';
    }

    status += refreshToken ? '✅ Refresh Token: Available\n' : '❌ Refresh Token: Missing\n';
    status += '\nReady to sync: ' + (token && refreshToken ? 'Yes' : 'No — re-authorize');
  }

  ui.alert('Constant Contact Status', status, ui.ButtonSet.OK);
}

/**
 * Removes all stored Constant Contact credentials and tokens.
 * Use this to disconnect CC or before re-configuring with a different account.
 */
function disconnectConstantContact() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ Disconnect Constant Contact',
    'This will remove all stored API credentials and tokens.\n\n' +
    'You will need to re-run the setup to reconnect.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(CC_CONFIG.PROP_API_KEY);
  props.deleteProperty(CC_CONFIG.PROP_API_SECRET);
  props.deleteProperty(CC_CONFIG.PROP_ACCESS_TOKEN);
  props.deleteProperty(CC_CONFIG.PROP_REFRESH_TOKEN);
  props.deleteProperty(CC_CONFIG.PROP_TOKEN_EXPIRY);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Constant Contact disconnected. All credentials removed.',
    '🔌 Disconnected', 5
  );
  Logger.log('Constant Contact credentials removed');
}

// ============================================================================
// v4.11.0: EDUCATIONAL RESOURCES HUB
// ============================================================================

/**
 * Returns active resource categories from 📚 Resource Config sheet, sorted by Sort Order.
 * Auto-creates the sheet with defaults if missing.
 * Used by the steward manage form and any UI that needs the live category list.
 * @returns {string[]} Category name strings, e.g. ['Contract Article', 'FAQ', ...]
 */
function getWebAppResourceCategories() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCE_CONFIG);

    // Auto-create with defaults if missing
    if (!sheet) {
      if (typeof createResourceConfigSheet === 'function') {
        sheet = createResourceConfigSheet(ss);
      }
      if (!sheet) return _defaultResourceCategories_();
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return _defaultResourceCategories_();

    var data = sheet.getRange(2, 1, lastRow - 1, RESOURCE_CONFIG_COLS.NOTES).getValues();

    var categories = data
      .filter(function(row) {
        var setting = String(row[RESOURCE_CONFIG_COLS.SETTING - 1] || '').trim();
        var active  = String(row[RESOURCE_CONFIG_COLS.ACTIVE - 1]  || '').toLowerCase();
        var value   = String(row[RESOURCE_CONFIG_COLS.VALUE - 1]   || '').trim();
        return setting === 'Category' && active === 'yes' && value !== '';
      })
      .map(function(row) {
        return {
          name:  String(row[RESOURCE_CONFIG_COLS.VALUE - 1]).trim(),
          order: Number(row[RESOURCE_CONFIG_COLS.SORT_ORDER - 1]) || 999
        };
      });

    categories.sort(function(a, b) { return a.order - b.order; });
    return categories.map(function(c) { return c.name; });

  } catch (e) {
    logError_('getWebAppResourceCategories', e);
    return _defaultResourceCategories_();
  }
}

/**
 * Fallback category list — used only if the sheet is missing and createResourceConfigSheet
 * cannot run (e.g. permissions issue). Matches the default rows in createResourceConfigSheet.
 * @private
 */
function _defaultResourceCategories_() {
  return [
    'Contract Article', 'Know Your Rights', 'Grievance Process',
    'Forms & Templates', 'FAQ', 'Guide', 'Policy', 'Contact Info', 'Link', 'General'
  ];
}

/**
 * API function to get resources list for web app.
 * Reads from the 📚 Resources sheet, returns visible items sorted by Sort Order.
 * @param {string} [audience] - Filter by audience: 'All', 'Members', 'Stewards'. Defaults to 'All'.
 * @returns {Array} Resource objects
 */
function getWebAppResourcesList(audience) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);

    // Auto-create if missing
    if (!sheet) {
      if (typeof createResourcesSheet === 'function') {
        sheet = createResourcesSheet(ss);
      }
      if (!sheet) return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, RESOURCES_COLS.ADDED_BY).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var visible = String(row[RESOURCES_COLS.VISIBLE - 1] || '').toLowerCase();
      if (visible !== 'yes') return null;

      var rowAudience = row[RESOURCES_COLS.AUDIENCE - 1] || 'All';
      if (audience && audience !== 'All' && rowAudience !== 'All' && rowAudience !== audience) return null;

      var dateAdded = row[RESOURCES_COLS.DATE_ADDED - 1];

      return {
        id: row[RESOURCES_COLS.RESOURCE_ID - 1] || '',
        title: row[RESOURCES_COLS.TITLE - 1] || '',
        category: row[RESOURCES_COLS.CATEGORY - 1] || 'General',
        summary: row[RESOURCES_COLS.SUMMARY - 1] || '',
        content: row[RESOURCES_COLS.CONTENT - 1] || '',
        url: row[RESOURCES_COLS.URL - 1] || '',
        icon: row[RESOURCES_COLS.ICON - 1] || '📄',
        sortOrder: row[RESOURCES_COLS.SORT_ORDER - 1] || 999,
        audience: rowAudience,
        dateAdded: dateAdded instanceof Date ? Utilities.formatDate(dateAdded, tz, 'MMM d, yyyy') : (dateAdded || '')
      };
    }).filter(function(r) { return r !== null; });

    // Sort by sortOrder
    result.sort(function(a, b) { return (a.sortOrder || 999) - (b.sortOrder || 999); });

    return result;
  } catch (e) {
    Logger.log('getWebAppResourcesList error: ' + e.toString());
    return [];
  }
}

/**
 * Get all resources including hidden ones (for steward manage tab).
 * Same as getWebAppResourcesList but without Visible=Yes filtering.
 * @returns {Object[]} All resources with visible status
 */
function getWebAppResourcesListAll() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, RESOURCES_COLS.ADDED_BY).getValues();
    var tz = Session.getScriptTimeZone();

    return data.map(function(row) {
      var dateAdded = row[RESOURCES_COLS.DATE_ADDED - 1];
      return {
        id: row[RESOURCES_COLS.RESOURCE_ID - 1] || '',
        title: row[RESOURCES_COLS.TITLE - 1] || '',
        category: row[RESOURCES_COLS.CATEGORY - 1] || 'General',
        summary: row[RESOURCES_COLS.SUMMARY - 1] || '',
        content: row[RESOURCES_COLS.CONTENT - 1] || '',
        url: row[RESOURCES_COLS.URL - 1] || '',
        icon: row[RESOURCES_COLS.ICON - 1] || '\uD83D\uDCC4',
        sortOrder: row[RESOURCES_COLS.SORT_ORDER - 1] || 999,
        visible: String(row[RESOURCES_COLS.VISIBLE - 1] || '').toLowerCase() === 'yes',
        audience: row[RESOURCES_COLS.AUDIENCE - 1] || 'All',
        dateAdded: dateAdded instanceof Date ? Utilities.formatDate(dateAdded, tz, 'MMM d, yyyy') : (dateAdded || ''),
        addedBy: row[RESOURCES_COLS.ADDED_BY - 1] || ''
      };
    }).filter(function(r) { return r.id; });
  } catch (e) {
    logError_('getWebAppResourcesListAll', e);
    return [];
  }
}


/**
 * Add a new resource to the 📚 Resources sheet.
 * @param {Object} data — { title, category, summary, content, url, icon, sortOrder, visible, audience }
 * @returns {Object} { success: boolean, resourceId: string, message: string }
 */
function addWebAppResource(sessionToken, data) {
  try {
    // CR-AUTH-6: Verify steward role before allowing resource creation
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (!sheet) {
      if (typeof createResourcesSheet === 'function') {
        sheet = createResourcesSheet(ss);
      }
      if (!sheet) return { success: false, message: 'Resources sheet not found' };
    }

    if (!data.title) return { success: false, message: 'Title is required' };

    // Generate next ID
    var allData = sheet.getDataRange().getValues();
    var maxNum = 0;
    for (var i = 1; i < allData.length; i++) {
      var existId = String(allData[i][RESOURCES_COLS.RESOURCE_ID - 1] || '');
      var match = existId.match(/RES-(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    var nextId = 'RES-' + String(maxNum + 1).padStart(3, '0');

    var tz = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var addedBy = auth.email || 'unknown';

    // CR-FORMULA: Escape all user-supplied fields to prevent formula injection
    var newRow = [
      nextId,
      escapeForFormula(data.title),
      escapeForFormula(data.category || 'General'),
      escapeForFormula(data.summary || ''),
      escapeForFormula(data.content || ''),
      escapeForFormula(data.url || ''),
      escapeForFormula(data.icon || '\uD83D\uDCC4'),
      data.sortOrder || 999,
      escapeForFormula(data.visible || 'Yes'),
      escapeForFormula(data.audience || 'All'),
      today,
      addedBy
    ];

    sheet.appendRow(newRow);
    return { success: true, resourceId: nextId, message: 'Resource added' };
  } catch (e) {
    logError_('addWebAppResource', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Update an existing resource by ID.
 * @param {string} resourceId — e.g. "RES-001"
 * @param {Object} data — fields to update (title, category, summary, content, url, icon, sortOrder, visible, audience)
 * @returns {Object} { success: boolean, message: string }
 */
function updateWebAppResource(sessionToken, resourceId, data) {
  try {
    // CR-AUTH-6: Verify steward role before allowing resource updates
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (!sheet) return { success: false, message: 'Resources sheet not found' };

    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][RESOURCES_COLS.RESOURCE_ID - 1] || '').trim() === resourceId) {
        // M-PERF: Batch write — modify row copy in-memory, write back in single call
        var rowData = allData[i].slice();
        // CR-FORMULA: Escape all user-supplied fields to prevent formula injection
        if (data.title !== undefined)     rowData[RESOURCES_COLS.TITLE - 1] = escapeForFormula(data.title);
        if (data.category !== undefined)  rowData[RESOURCES_COLS.CATEGORY - 1] = escapeForFormula(data.category);
        if (data.summary !== undefined)   rowData[RESOURCES_COLS.SUMMARY - 1] = escapeForFormula(data.summary);
        if (data.content !== undefined)   rowData[RESOURCES_COLS.CONTENT - 1] = escapeForFormula(data.content);
        if (data.url !== undefined)       rowData[RESOURCES_COLS.URL - 1] = escapeForFormula(data.url);
        if (data.icon !== undefined)      rowData[RESOURCES_COLS.ICON - 1] = escapeForFormula(data.icon);
        if (data.sortOrder !== undefined) rowData[RESOURCES_COLS.SORT_ORDER - 1] = data.sortOrder;
        if (data.visible !== undefined)   rowData[RESOURCES_COLS.VISIBLE - 1] = escapeForFormula(data.visible);
        if (data.audience !== undefined)  rowData[RESOURCES_COLS.AUDIENCE - 1] = escapeForFormula(data.audience);
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, message: 'Resource updated' };
      }
    }
    return { success: false, message: 'Resource not found' };
  } catch (e) {
    logError_('updateWebAppResource', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Soft-delete a resource (set Visible=No).
 * @param {string} resourceId — e.g. "RES-001"
 * @returns {Object} { success: boolean, message: string }
 */
function deleteWebAppResource(sessionToken, resourceId) {
  try {
    // CR-AUTH-6: Verify steward role before allowing resource deletion
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (!sheet) return { success: false, message: 'Resources sheet not found' };

    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][RESOURCES_COLS.RESOURCE_ID - 1] || '').trim() === resourceId) {
        sheet.getRange(i + 1, RESOURCES_COLS.VISIBLE).setValue('No');
        return { success: true, message: 'Resource hidden' };
      }
    }
    return { success: false, message: 'Resource not found' };
  } catch (e) {
    logError_('deleteWebAppResource', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Restore a soft-deleted resource (set Visible=Yes).
 * @param {string} resourceId — e.g. "RES-001"
 * @returns {Object} { success: boolean, message: string }
 */
function restoreWebAppResource(sessionToken, resourceId) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.authorized) return { success: false, message: auth.message || 'Unauthorized' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.RESOURCES);
    if (!sheet) return { success: false, message: 'Resources sheet not found' };

    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][RESOURCES_COLS.RESOURCE_ID - 1] || '').trim() === resourceId) {
        sheet.getRange(i + 1, RESOURCES_COLS.VISIBLE).setValue('Yes');
        return { success: true, message: 'Resource restored' };
      }
    }
    return { success: false, message: 'Resource not found' };
  } catch (e) {
    logError_('restoreWebAppResource', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Returns the educational resources hub HTML page.
 * Warm, welcoming design with categorized content cards.
 * Content is fully dynamic — reads from 📚 Resources sheet.
 */
function getWebAppResourcesHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  // Read org name from config
  var orgName = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet && configSheet.getLastRow() >= 3) {
      orgName = configSheet.getRange(3, CONFIG_COLS.ORG_NAME).getValue() || '';
    }
  } catch (_e) { /* ignore */ }

  return '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Know Your Rights</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:wght@600;700;800&display=swap" rel="stylesheet">' +
    '<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:"DM Sans",sans-serif;background:#fafaf9;color:#1c1917;min-height:100vh;padding-bottom:80px}' +

    // Header — warm gradient, not cold purple
    '.header{background:linear-gradient(135deg,#1e3a5f 0%,#2d5a87 100%);color:white;padding:24px 20px 28px;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-family:"Fraunces",serif;font-size:clamp(22px,5vw,28px);font-weight:700;margin-bottom:6px;letter-spacing:-0.02em}' +
    '.header .subtitle{font-size:14px;opacity:0.85;line-height:1.4}' +

    // Search
    '.search-wrap{padding:0 16px;margin-top:-16px;position:relative;z-index:2}' +
    '.search-bar{width:100%;padding:14px 14px 14px 44px;border:2px solid #e7e5e4;border-radius:14px;font-size:15px;font-family:"DM Sans",sans-serif;background:white;box-shadow:0 2px 8px rgba(0,0,0,0.06);transition:border-color 0.2s}' +
    '.search-bar:focus{outline:none;border-color:#2d5a87;box-shadow:0 2px 12px rgba(45,90,135,0.15)}' +
    '.search-icon{position:absolute;left:30px;top:50%;transform:translateY(-50%);color:#a8a29e;font-size:20px}' +

    // Category pills
    '.cat-pills{display:flex;gap:8px;overflow-x:auto;padding:20px 16px 8px;-webkit-overflow-scrolling:touch}' +
    '.cat-pill{flex-shrink:0;padding:8px 16px;border-radius:24px;font-size:13px;font-weight:600;border:2px solid #e7e5e4;cursor:pointer;background:white;color:#57534e;transition:all 0.2s;white-space:nowrap}' +
    '.cat-pill.active{background:#1e3a5f;color:white;border-color:#1e3a5f}' +
    '.cat-pill:hover{border-color:#1e3a5f}' +

    // Resource cards
    '.resources-list{padding:12px 16px}' +
    '.res-card{background:white;border-radius:16px;padding:20px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;cursor:pointer;transition:all 0.25s}' +
    '.res-card:active{transform:scale(0.99)}' +
    '.res-card.expanded{box-shadow:0 4px 16px rgba(0,0,0,0.1)}' +
    '.res-icon{font-size:28px;margin-bottom:10px;display:block}' +
    '.res-title{font-family:"Fraunces",serif;font-size:17px;font-weight:700;color:#1c1917;margin-bottom:4px;line-height:1.3}' +
    '.res-category{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:0.05em;color:#2d5a87;margin-bottom:8px}' +
    '.res-summary{font-size:14px;color:#57534e;line-height:1.5}' +
    '.res-content{display:none;margin-top:16px;padding-top:16px;border-top:1px solid #f5f5f4;font-size:14px;color:#44403c;line-height:1.7;white-space:pre-line}' +
    '.res-card.expanded .res-content{display:block}' +
    '.res-link{display:inline-flex;align-items:center;gap:6px;margin-top:12px;padding:8px 16px;background:#eff6ff;color:#1e3a5f;border-radius:8px;font-size:13px;font-weight:600;text-decoration:none;transition:background 0.2s}' +
    '.res-link:hover{background:#dbeafe}' +
    '.res-meta{font-size:11px;color:#a8a29e;margin-top:8px}' +
    '.expand-hint{font-size:12px;color:#a8a29e;display:flex;align-items:center;gap:4px;margin-top:8px}' +
    '.expand-hint i{font-size:16px;transition:transform 0.2s}' +
    '.res-card.expanded .expand-hint i{transform:rotate(180deg)}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 24px}' +
    '.empty-icon{font-size:56px;margin-bottom:16px;opacity:0.6}' +
    '.empty-text{font-size:16px;color:#78716c;line-height:1.5}' +
    '.empty-sub{font-size:13px;color:#a8a29e;margin-top:8px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#78716c}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:28px;height:28px;border:3px solid #e7e5e4;border-top-color:#1e3a5f;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#1e3a5f}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h1>📚 Know Your Rights</h1>' +
    '<div class="subtitle">Your guide to the union contract, grievance process, and workplace protections' + (orgName ? ' at ' + escapeHtml(orgName) : '') + '</div>' +
    '</div>' +

    '<div class="search-wrap">' +
    '<i class="material-icons search-icon">search</i>' +
    '<input type="text" class="search-bar" id="searchInput" placeholder="Search articles, rights, FAQ..." oninput="filterResources()">' +
    '</div>' +

    '<div class="cat-pills" id="catPills"></div>' +

    '<div class="resources-list" id="resourcesList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:12px">Loading resources...</div></div>' +
    '</div>' +

    _getBottomNavHTML_(baseUrl, _MEMBER_NAV_, 'resources') +

    '<script>' +
    getClientSideEscapeHtml() +
    'var allResources=[];var currentCat="all";' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function filterResources(){' +
    '  var q=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allResources.filter(function(r){' +
    '    var matchesCat=currentCat==="all"||r.category===currentCat;' +
    '    var matchesQ=!q||q.length<2||r.title.toLowerCase().indexOf(q)>=0||r.summary.toLowerCase().indexOf(q)>=0||(r.content||"").toLowerCase().indexOf(q)>=0||r.category.toLowerCase().indexOf(q)>=0;' +
    '    return matchesCat&&matchesQ;' +
    '  });' +
    '  renderList(filtered);' +
    '}' +

    'function setCat(cat,btn){' +
    '  currentCat=cat;' +
    '  document.querySelectorAll(".cat-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterResources();' +
    '}' +

    'function renderList(items){' +
    '  var c=document.getElementById("resourcesList");' +
    '  if(!items||items.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📖</div><div class=\\"empty-text\\">No resources found</div><div class=\\"empty-sub\\">Try a different search or category</div></div>";' +
    '    return;' +
    '  }' +
    '  c.innerHTML=items.map(function(r){' +
    '    var hasContent=r.content&&r.content.trim().length>0;' +
    '    var hasUrl=r.url&&r.url.trim().length>0;' +
    '    return "<div class=\\"res-card\\"' + ('" onclick=\\"toggleCard(this)\\"') + '>"' +
    '      +"<span class=\\"res-icon\\">"+escapeHtml(r.icon)+"</span>"' +
    '      +"<div class=\\"res-category\\">"+escapeHtml(r.category)+"</div>"' +
    '      +"<div class=\\"res-title\\">"+escapeHtml(r.title)+"</div>"' +
    '      +"<div class=\\"res-summary\\">"+escapeHtml(r.summary)+"</div>"' +
    '      +(hasContent?"<div class=\\"expand-hint\\"><i class=\\"material-icons\\">expand_more</i>Tap to read more</div>":"")' +
    '      +(hasContent?"<div class=\\"res-content\\">"+escapeHtml(r.content).replace(/\\\\n/g,"\\n")+"</div>":"")' +
    '      +(hasUrl?"<a class=\\"res-link\\" href=\\""+escapeHtml(r.url)+"\\" target=\\"_blank\\" onclick=\\"event.stopPropagation()\\"><i class=\\"material-icons\\" style=\\"font-size:16px\\">open_in_new</i>Open Resource</a>":"")' +
    '      +(r.dateAdded?"<div class=\\"res-meta\\">Added "+escapeHtml(r.dateAdded)+"</div>":"")' +
    '      +"</div>";' +
    '  }).join("");' +
    '}' +

    'function buildCatPills(items){' +
    '  var cats={};items.forEach(function(r){cats[r.category]=(cats[r.category]||0)+1});' +
    '  var el=document.getElementById("catPills");' +
    '  var html="<button class=\\"cat-pill active\\" onclick=\\"setCat(\\x27all\\x27,this)\\">All ("+items.length+")</button>";' +
    '  Object.keys(cats).sort().forEach(function(c){' +
    '    html+="<button class=\\"cat-pill\\" onclick=\\"setCat("+JSON.stringify(c).replace(/"/g,"&quot;")+",this)\\">"+escapeHtml(c)+" ("+cats[c]+")</button>";' +
    '  });' +
    '  el.innerHTML=html;' +
    '}' +

    'google.script.run.withSuccessHandler(function(data){' +
    '  allResources=data||[];' +
    '  buildCatPills(allResources);' +
    '  renderList(allResources);' +
    '}).withFailureHandler(function(err){' +
    '  document.getElementById("resourcesList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error loading resources</div><div class=\\"empty-sub\\">"+String(err||"")+"</div></div>";' +
    '}).getWebAppResourcesList();' +
    '</script>' +

    '</body></html>';
}

// ============================================================================
// v4.11.0: MEETING CHECK-IN WEB ROUTE
// ============================================================================

/**
 * Returns a standalone web page for meeting check-in.
 * Reuses the existing check-in logic from 14_MeetingCheckIn.gs.
 * Mobile-optimized for kiosk mode (shared device) or individual phone check-in.
 */
function getWebAppCheckInHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Meeting Check-In</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Fraunces:wght@600;700&display=swap" rel="stylesheet">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:"DM Sans",sans-serif;background:#fafaf9;min-height:100vh;padding-bottom:80px}' +

    '.header{background:linear-gradient(135deg,#065f46 0%,#047857 100%);color:white;padding:24px 20px 28px;text-align:center;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-family:"Fraunces",serif;font-size:clamp(22px,5vw,28px);font-weight:700;margin-bottom:4px}' +
    '.header .subtitle{font-size:14px;opacity:0.85}' +

    '.container{max-width:480px;margin:0 auto;padding:20px 16px}' +

    '.card{background:white;border-radius:16px;padding:24px;box-shadow:0 1px 4px rgba(0,0,0,0.06);border:1px solid #f5f5f4;margin-bottom:16px}' +
    '.card h2{font-family:"Fraunces",serif;font-size:18px;color:#065f46;margin-bottom:16px;text-align:center}' +

    '.field{margin-bottom:18px}' +
    '.field label{display:block;margin-bottom:8px;font-weight:600;color:#1c1917;font-size:14px}' +
    '.field input,.field select{width:100%;padding:14px;border:2px solid #e7e5e4;border-radius:12px;font-size:16px;font-family:"DM Sans",sans-serif;transition:border-color 0.2s}' +
    '.field input:focus,.field select:focus{outline:none;border-color:#047857;box-shadow:0 0 0 3px rgba(4,120,87,0.1)}' +

    '.btn{width:100%;padding:16px;border:none;border-radius:12px;font-size:17px;font-weight:700;font-family:"DM Sans",sans-serif;cursor:pointer;transition:all 0.2s}' +
    '.btn-checkin{background:#047857;color:white}' +
    '.btn-checkin:hover{background:#065f46}' +
    '.btn-checkin:disabled{background:#d6d3d1;cursor:not-allowed}' +

    '.error-msg{color:#dc2626;font-size:14px;text-align:center;padding:12px;background:#fef2f2;border-radius:10px;margin-top:12px;display:none}' +

    '.success-banner{text-align:center;padding:28px;background:#d1fae5;border-radius:16px;margin-bottom:16px;display:none}' +
    '.success-banner .checkmark{font-size:56px;margin-bottom:8px}' +
    '.success-banner .name{font-family:"Fraunces",serif;font-size:22px;font-weight:700;color:#065f46;margin-bottom:4px}' +
    '.success-banner .msg{color:#047857;font-size:14px}' +

    '.attendee-count{text-align:center;padding:14px;background:#ecfdf5;border-radius:12px;color:#047857;font-weight:600;margin-top:12px;font-size:14px}' +

    '.no-meetings{text-align:center;padding:40px 20px;color:#78716c}' +
    '.no-meetings-icon{font-size:48px;margin-bottom:12px;opacity:0.5}' +

    '.meeting-option{padding:14px;border:2px solid #e7e5e4;border-radius:12px;margin-bottom:10px;cursor:pointer;transition:all 0.2s}' +
    '.meeting-option:hover{border-color:#047857}' +
    '.meeting-option.selected{border-color:#047857;background:#ecfdf5}' +
    '.meeting-name{font-weight:600;color:#1c1917;margin-bottom:2px}' +
    '.meeting-time{font-size:13px;color:#78716c}' +

    '.loading{text-align:center;padding:40px;color:#78716c}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:28px;height:28px;border:3px solid #e7e5e4;border-top-color:#047857;border-radius:50%;animation:spin 0.8s linear infinite}' +

    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#065f46}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h1>✅ Meeting Check-In</h1>' +
    '<div class="subtitle">Enter your email and PIN to check in</div>' +
    '</div>' +

    '<div class="container">' +

    '<div id="successBanner" class="success-banner">' +
    '<div class="checkmark">✓</div>' +
    '<div class="name" id="successName"></div>' +
    '<div class="msg">Checked in successfully!</div>' +
    '</div>' +

    '<div class="card" id="meetingCard">' +
    '<h2>Today\'s Meetings</h2>' +
    '<div id="meetingList"><div class="loading"><div class="spinner"></div><div style="margin-top:12px">Loading meetings...</div></div></div>' +
    '</div>' +

    '<div class="card" id="checkinForm" style="display:none">' +
    '<h2>Check In</h2>' +
    '<div class="field"><label>Email Address</label>' +
    '<input type="email" id="email" placeholder="your.email@example.com" autocomplete="email"></div>' +
    '<div class="field"><label>PIN</label>' +
    '<input type="password" id="pin" inputmode="numeric" maxlength="8" placeholder="Enter your PIN"></div>' +
    '<button class="btn btn-checkin" id="checkinBtn" onclick="doCheckIn()">Check In</button>' +
    '<div class="error-msg" id="errorMsg"></div>' +
    '<div class="attendee-count" id="attendeeCount" style="display:none"></div>' +
    '</div>' +

    '</div>' +

    _getBottomNavHTML_(baseUrl, _MEMBER_NAV_, 'checkin') +

    '<script>' +
    'var selectedMeetingId=null;' +

    'google.script.run.withSuccessHandler(function(meetings){' +
    '  var el=document.getElementById("meetingList");' +
    '  if(!meetings||meetings.length===0){' +
    '    el.innerHTML="<div class=\\"no-meetings\\"><div class=\\"no-meetings-icon\\">📅</div><div>No active meetings right now</div><div style=\\"font-size:13px;margin-top:8px\\">Check-in opens on the day of a scheduled meeting</div></div>";' +
    '    return;' +
    '  }' +
    '  if(meetings.length===1){' +
    '    selectedMeetingId=meetings[0].id;' +
    '    el.innerHTML="<div class=\\"meeting-option selected\\"><div class=\\"meeting-name\\">"+meetings[0].name+"</div><div class=\\"meeting-time\\">"+meetings[0].time+"</div></div>";' +
    '    document.getElementById("checkinForm").style.display="block";' +
    '  }else{' +
    '    var html="<div style=\\"font-size:14px;color:#78716c;margin-bottom:12px\\">Select your meeting:</div>";' +
    '    meetings.forEach(function(m){' +
    '      html+="<div class=\\"meeting-option\\" onclick=\\"selectMeeting(\\x27"+m.id+"\\x27,this)\\"><div class=\\"meeting-name\\">"+m.name+"</div><div class=\\"meeting-time\\">"+m.time+"</div></div>";' +
    '    });' +
    '    el.innerHTML=html;' +
    '  }' +
    '}).withFailureHandler(function(err){' +
    '  document.getElementById("meetingList").innerHTML="<div class=\\"no-meetings\\"><div class=\\"no-meetings-icon\\">⚠️</div><div>Error loading meetings</div></div>";' +
    '}).getCheckInEligibleMeetings();' +

    'function selectMeeting(id,el){' +
    '  selectedMeetingId=id;' +
    '  document.querySelectorAll(".meeting-option").forEach(function(o){o.classList.remove("selected")});' +
    '  el.classList.add("selected");' +
    '  document.getElementById("checkinForm").style.display="block";' +
    '  document.getElementById("email").focus();' +
    '}' +

    'function doCheckIn(){' +
    '  var email=document.getElementById("email").value.trim();' +
    '  var pin=document.getElementById("pin").value.trim();' +
    '  var errEl=document.getElementById("errorMsg");' +
    '  errEl.style.display="none";' +
    '  if(!selectedMeetingId){errEl.textContent="Please select a meeting";errEl.style.display="block";return}' +
    '  if(!email){errEl.textContent="Please enter your email";errEl.style.display="block";return}' +
    '  if(!pin||pin.length<4){errEl.textContent="Please enter your PIN (4-8 digits)";errEl.style.display="block";return}' +
    '  var btn=document.getElementById("checkinBtn");' +
    '  btn.disabled=true;btn.textContent="Checking in...";' +
    '  google.script.run.withSuccessHandler(function(result){' +
    '    btn.disabled=false;btn.textContent="Check In";' +
    '    if(result&&result.success){' +
    '      document.getElementById("successName").textContent=result.memberName||"Member";' +
    '      document.getElementById("successBanner").style.display="block";' +
    '      if(result.attendeeCount){var ac=document.getElementById("attendeeCount");ac.textContent="✓ "+result.attendeeCount+" checked in so far";ac.style.display="block"}' +
    '      document.getElementById("email").value="";document.getElementById("pin").value="";' +
    '      setTimeout(function(){document.getElementById("successBanner").style.display="none"},4000);' +
    '    }else{' +
    '      errEl.textContent=(result&&result.message)||"Check-in failed. Verify your email and PIN.";errEl.style.display="block";' +
    '    }' +
    '  }).withFailureHandler(function(err){' +
    '    btn.disabled=false;btn.textContent="Check In";' +
    '    errEl.textContent="Error: "+String(err||"Unknown");errEl.style.display="block";' +
    '  }).processMeetingCheckIn(selectedMeetingId,email,pin);' +
    '}' +

    'document.getElementById("pin").addEventListener("keypress",function(e){if(e.key==="Enter")doCheckIn()});' +
    '</script>' +

    '</body></html>';
}


// ============================================================================
// DEADLINE CALENDAR VIEW (v4.13.0 — PHASE2)
// ============================================================================
// Visual calendar/list of grievance deadlines for stewards.
// Route: ?page=deadlines (requires steward auth via sensitivePages).
// ============================================================================

/**
 * Returns deadline data for all open grievances with upcoming due dates.
 * Reads from the Grievance Log and computes urgency based on days remaining.
 * @returns {Object} { deadlines: [{ grievanceId, memberName, currentStep, dueDate, daysRemaining, urgency, steward, issueCategory }] }
 */
function getDeadlineCalendarData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return { deadlines: [] };
  }

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var deadlines = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][GRIEVANCE_COLS.STATUS - 1] || '').trim();

    // Only include open/active grievances (skip Closed, Won, Withdrawn, Denied, Settled)
    var closedStatuses = ['Closed', 'Won', 'Withdrawn', 'Denied', 'Settled'];
    if (closedStatuses.indexOf(status) !== -1) continue;

    var grievanceId = String(data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '').trim();
    if (!grievanceId) continue;

    var firstName = String(data[i][GRIEVANCE_COLS.FIRST_NAME - 1] || '').trim();
    var lastName = String(data[i][GRIEVANCE_COLS.LAST_NAME - 1] || '').trim();
    var memberName = (firstName + ' ' + lastName).trim() || 'Unknown Member';
    var currentStep = String(data[i][GRIEVANCE_COLS.CURRENT_STEP - 1] || '').trim();
    var steward = String(data[i][GRIEVANCE_COLS.STEWARD - 1] || '').trim();
    var issueCategory = String(data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').trim();

    // Determine the relevant due date based on the current step
    var dueDate = null;
    var nextActionDue = data[i][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];

    // Prefer NEXT_ACTION_DUE if it's a valid date
    if (nextActionDue instanceof Date && !isNaN(nextActionDue.getTime())) {
      dueDate = nextActionDue;
    } else {
      // Fall back to step-specific due dates
      var stepDueCols = [
        GRIEVANCE_COLS.FILING_DEADLINE,
        GRIEVANCE_COLS.STEP1_DUE,
        GRIEVANCE_COLS.STEP2_APPEAL_DUE,
        GRIEVANCE_COLS.STEP2_DUE,
        GRIEVANCE_COLS.STEP3_APPEAL_DUE
      ];

      // Walk backwards to find the latest relevant due date
      for (var s = stepDueCols.length - 1; s >= 0; s--) {
        var candidate = data[i][stepDueCols[s] - 1];
        if (candidate instanceof Date && !isNaN(candidate.getTime())) {
          dueDate = candidate;
          break;
        }
      }
    }

    if (!dueDate) continue;

    // Calculate days remaining
    var dueDateClean = new Date(dueDate);
    dueDateClean.setHours(0, 0, 0, 0);
    var diffMs = dueDateClean.getTime() - today.getTime();
    var daysRemaining = Math.ceil(diffMs / (1000 * 60 * 60 * 24));

    // Determine urgency color
    var urgency = 'green';
    if (daysRemaining < 3) {
      urgency = 'red';
    } else if (daysRemaining <= 7) {
      urgency = 'orange';
    }

    // Also use pre-computed DAYS_TO_DEADLINE if available and dueDate wasn't from NEXT_ACTION_DUE
    var daysToDeadline = data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    if (typeof daysToDeadline === 'number' && !isNaN(daysToDeadline) && !(nextActionDue instanceof Date)) {
      daysRemaining = daysToDeadline;
      urgency = daysRemaining < 3 ? 'red' : (daysRemaining <= 7 ? 'orange' : 'green');
    }

    // Format due date as YYYY-MM-DD for the calendar
    var yyyy = dueDateClean.getFullYear();
    var mm = ('0' + (dueDateClean.getMonth() + 1)).slice(-2);
    var dd = ('0' + dueDateClean.getDate()).slice(-2);
    var dueDateStr = yyyy + '-' + mm + '-' + dd;

    deadlines.push({
      grievanceId: escapeHtml(grievanceId),
      memberName: escapeHtml(memberName),
      currentStep: escapeHtml(currentStep || 'N/A'),
      dueDate: dueDateStr,
      daysRemaining: daysRemaining,
      urgency: urgency,
      steward: escapeHtml(steward),
      issueCategory: escapeHtml(issueCategory),
      status: escapeHtml(status)
    });
  }

  // Sort by due date ascending (most urgent first)
  deadlines.sort(function(a, b) {
    return a.daysRemaining - b.daysRemaining;
  });

  return { deadlines: deadlines };
}

/**
 * Returns a self-contained HTML page with a deadline calendar view.
 * Steward-only page with month grid and list views, color-coded urgency.
 * @returns {string} Complete HTML document
 */
function getDeadlineCalendarHtml() {
  var baseUrl = '';
  try { baseUrl = ScriptApp.getService().getUrl(); } catch (_e) { /* ignore */ }

  // Pre-fetch deadline data server-side
  var calData = { deadlines: [] };
  try {
    calData = getDeadlineCalendarData();
  } catch (_e) {
    calData = { deadlines: [] };
  }
  var deadlinesJson = JSON.stringify(calData.deadlines || []);

  // Status colors from COMMAND_CONFIG
  var statusColorsJson = '{}';
  try {
    if (typeof COMMAND_CONFIG !== 'undefined' && COMMAND_CONFIG.STATUS_COLORS) {
      statusColorsJson = JSON.stringify(COMMAND_CONFIG.STATUS_COLORS);
    }
  } catch (_e) { /* ignore */ }

  var html = '<!DOCTYPE html>' +
    '<html lang="en"><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=5.0,user-scalable=yes">' +
    '<title>Deadline Calendar | ' + escapeHtml(typeof getConfigValue_ === 'function' ? (getConfigValue_(CONFIG_COLS.ORG_NAME) || 'Union') : 'Union') + '</title>' +
    '<style>' +

    // Reset
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +

    // Body
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;' +
    'background:#fafaf9;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#1e40af 0%,#1d4ed8 100%);color:white;' +
    'padding:24px 20px 32px;text-align:center;position:relative;overflow:hidden}' +
    '.header::after{content:"";position:absolute;bottom:-20px;left:-20px;right:-20px;' +
    'height:40px;background:#fafaf9;border-radius:50% 50% 0 0}' +
    '.header h1{font-size:clamp(20px,4.5vw,26px);font-weight:700;margin-bottom:4px}' +
    '.header .subtitle{font-size:clamp(12px,2.5vw,14px);opacity:0.85}' +

    // Container
    '.container{max-width:960px;margin:0 auto;padding:16px}' +

    // Summary stats bar
    '.stats-bar{display:flex;gap:10px;margin-bottom:16px;overflow-x:auto;padding-bottom:4px}' +
    '.stat-pill{flex:0 0 auto;padding:10px 18px;border-radius:12px;text-align:center;' +
    'font-weight:700;font-size:14px;min-width:100px}' +
    '.stat-pill.red{background:#fef2f2;color:#dc2626;border:1px solid #fecaca}' +
    '.stat-pill.orange{background:#fff7ed;color:#ea580c;border:1px solid #fed7aa}' +
    '.stat-pill.green{background:#f0fdf4;color:#16a34a;border:1px solid #bbf7d0}' +
    '.stat-pill .num{font-size:22px;display:block}' +

    // Toggle bar
    '.toggle-bar{display:flex;background:white;border-radius:12px;padding:4px;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.08);margin-bottom:16px;border:1px solid #e5e7eb}' +
    '.toggle-btn{flex:1;padding:10px;border:none;border-radius:10px;font-size:14px;' +
    'font-weight:600;cursor:pointer;background:transparent;color:#6b7280;transition:all 0.2s}' +
    '.toggle-btn.active{background:#1d4ed8;color:white}' +

    // ---- LIST VIEW ----
    '.deadline-list{display:flex;flex-direction:column;gap:10px}' +
    '.deadline-card{background:white;border-radius:14px;padding:16px;' +
    'box-shadow:0 1px 3px rgba(0,0,0,0.06);border:1px solid #f5f5f4;' +
    'display:flex;align-items:center;gap:14px;cursor:pointer;transition:all 0.2s}' +
    '.deadline-card:hover{box-shadow:0 3px 10px rgba(0,0,0,0.1);transform:translateY(-1px)}' +
    '.urgency-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}' +
    '.urgency-dot.red{background:#dc2626}' +
    '.urgency-dot.orange{background:#ea580c}' +
    '.urgency-dot.green{background:#16a34a}' +
    '.deadline-info{flex:1;min-width:0}' +
    '.deadline-case{font-weight:700;color:#1f2937;font-size:14px}' +
    '.deadline-member{font-size:13px;color:#6b7280;margin-top:2px}' +
    '.deadline-step{font-size:12px;color:#9ca3af;margin-top:1px}' +
    '.deadline-days{text-align:right;flex-shrink:0}' +
    '.days-num{font-size:22px;font-weight:700;line-height:1}' +
    '.days-num.red{color:#dc2626}' +
    '.days-num.orange{color:#ea580c}' +
    '.days-num.green{color:#16a34a}' +
    '.days-label{font-size:10px;color:#9ca3af;text-transform:uppercase;text-align:center}' +
    '.overdue-badge{font-size:11px;font-weight:700;color:#dc2626;background:#fef2f2;' +
    'padding:2px 8px;border-radius:6px;display:inline-block}' +

    // ---- MONTH VIEW ----
    '.month-nav{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}' +
    '.month-title{font-size:18px;font-weight:700;color:#1f2937}' +
    '.month-btn{padding:8px 14px;border:1px solid #e5e7eb;border-radius:8px;background:white;' +
    'cursor:pointer;font-size:16px;transition:background 0.2s}' +
    '.month-btn:hover{background:#f3f4f6}' +
    '.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;background:white;' +
    'border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.06);border:1px solid #e5e7eb}' +
    '.cal-header{padding:10px 4px;text-align:center;font-size:12px;font-weight:600;' +
    'color:#6b7280;background:#f9fafb}' +
    '.cal-day{min-height:70px;padding:4px;border:1px solid #f3f4f6;position:relative;cursor:default}' +
    '.cal-day.today{background:#eff6ff}' +
    '.cal-day.other-month{background:#fafafa;color:#d1d5db}' +
    '.cal-day-num{font-size:12px;font-weight:500;color:#374151;margin-bottom:2px}' +
    '.cal-day.other-month .cal-day-num{color:#d1d5db}' +
    '.cal-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin:1px}' +
    '.cal-dot.red{background:#dc2626}.cal-dot.orange{background:#ea580c}.cal-dot.green{background:#16a34a}' +
    '.cal-event{font-size:10px;padding:2px 4px;border-radius:4px;margin-top:1px;' +
    'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}' +
    '.cal-event.red{background:#fef2f2;color:#dc2626}' +
    '.cal-event.orange{background:#fff7ed;color:#ea580c}' +
    '.cal-event.green{background:#f0fdf4;color:#16a34a}' +

    // ---- DETAIL PANEL ----
    '.detail-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);' +
    'display:none;align-items:center;justify-content:center;z-index:1000;padding:20px}' +
    '.detail-overlay.show{display:flex}' +
    '.detail-panel{background:white;border-radius:16px;padding:24px;max-width:420px;width:100%;' +
    'box-shadow:0 20px 40px rgba(0,0,0,0.15);animation:slideUp 0.25s ease}' +
    '@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}' +
    '.detail-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}' +
    '.detail-id{font-size:18px;font-weight:700;color:#1f2937}' +
    '.detail-close{width:32px;height:32px;border:none;background:#f3f4f6;border-radius:8px;' +
    'cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center}' +
    '.detail-close:hover{background:#e5e7eb}' +
    '.detail-row{display:flex;justify-content:space-between;padding:10px 0;border-bottom:1px solid #f3f4f6}' +
    '.detail-label{font-size:13px;color:#6b7280;font-weight:500}' +
    '.detail-value{font-size:13px;color:#1f2937;font-weight:600;text-align:right;max-width:60%}' +
    '.detail-urgency{display:inline-block;padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700}' +
    '.detail-urgency.red{background:#fef2f2;color:#dc2626}' +
    '.detail-urgency.orange{background:#fff7ed;color:#ea580c}' +
    '.detail-urgency.green{background:#f0fdf4;color:#16a34a}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#9ca3af}' +
    '.empty-state-icon{font-size:48px;margin-bottom:12px;opacity:0.5}' +

    // Bottom navigation
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;' +
    'display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));' +
    'box-shadow:0 -2px 10px rgba(0,0,0,0.08);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;' +
    'text-decoration:none;color:#a8a29e;font-size:10px;font-weight:500;min-width:60px;transition:color 0.2s}' +
    '.nav-item.active{color:#1d4ed8}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    // Responsive
    '@media(min-width:768px){' +
    '.container{padding:20px 24px}' +
    '.cal-day{min-height:90px;padding:6px}' +
    '.cal-event{font-size:11px}' +
    '}' +

    '</style></head><body>' +

    // Detail overlay
    '<div id="detailOverlay" class="detail-overlay" onclick="closeDetail(event)">' +
    '<div class="detail-panel" id="detailPanel"></div>' +
    '</div>' +

    // Header
    '<div class="header">' +
    '<h1>Deadline Calendar</h1>' +
    '<div class="subtitle">Grievance deadlines at a glance</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats bar
    '<div class="stats-bar" id="statsBar"></div>' +

    // View toggle
    '<div class="toggle-bar">' +
    '<button class="toggle-btn active" id="btnList" onclick="showView(\\x27list\\x27)">List View</button>' +
    '<button class="toggle-btn" id="btnMonth" onclick="showView(\\x27month\\x27)">Month View</button>' +
    '</div>' +

    // List view container
    '<div id="listView" class="deadline-list"></div>' +

    // Month view container (hidden by default)
    '<div id="monthView" style="display:none"></div>' +

    '</div>' +

    // Bottom navigation (M-DUP-1: extracted to helper)
    _getBottomNavHTML_(baseUrl, _DEADLINE_NAV_, 'deadlines') +

    '<script>' +

    // Client-side escapeHtml (canonical helper from 00_Security.gs)
    getClientSideEscapeHtml() +

    // Data
    'var deadlines=' + deadlinesJson + ';' +
    'var statusColors=' + statusColorsJson + ';' +
    'var currentMonth=new Date().getMonth();' +
    'var currentYear=new Date().getFullYear();' +

    // Render stats bar
    'function renderStats(){' +
    '  var red=0,orange=0,green=0,overdue=0;' +
    '  deadlines.forEach(function(d){' +
    '    if(d.daysRemaining<0)overdue++;' +
    '    if(d.urgency==="red")red++;' +
    '    else if(d.urgency==="orange")orange++;' +
    '    else green++;' +
    '  });' +
    '  var el=document.getElementById("statsBar");' +
    '  el.innerHTML=' +
    '    "<div class=\\"stat-pill red\\"><span class=\\"num\\">"+red+"</span>Critical (&lt;3d)</div>"' +
    '    +"<div class=\\"stat-pill orange\\"><span class=\\"num\\">"+orange+"</span>Soon (3-7d)</div>"' +
    '    +"<div class=\\"stat-pill green\\"><span class=\\"num\\">"+green+"</span>On Track (&gt;7d)</div>";' +
    '}' +
    'renderStats();' +

    // ---- LIST VIEW ----
    'function renderList(){' +
    '  var el=document.getElementById("listView");' +
    '  if(deadlines.length===0){' +
    '    el.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">&#x2705;</div>' +
    '<div style=\\"font-size:18px;margin-bottom:8px\\">No upcoming deadlines</div>' +
    '<div style=\\"font-size:13px\\">All open grievances are on track</div></div>";' +
    '    return;' +
    '  }' +
    '  var h="";' +
    '  deadlines.forEach(function(d,i){' +
    '    h+="<div class=\\"deadline-card\\" onclick=\\"showDetail("+i+")\\">"' +
    '      +"<div class=\\"urgency-dot "+d.urgency+"\\"></div>"' +
    '      +"<div class=\\"deadline-info\\">"' +
    '        +"<div class=\\"deadline-case\\">"+escapeHtml(d.grievanceId)+"</div>"' +
    '        +"<div class=\\"deadline-member\\">"+escapeHtml(d.memberName)+"</div>"' +
    '        +"<div class=\\"deadline-step\\">"+escapeHtml(d.currentStep)+" &middot; Due "+escapeHtml(d.dueDate)+"</div>"' +
    '      +"</div>"' +
    '      +"<div class=\\"deadline-days\\">";' +
    '    if(d.daysRemaining<0){' +
    '      h+="<div class=\\"overdue-badge\\">"+Math.abs(d.daysRemaining)+"d overdue</div>";' +
    '    }else if(d.daysRemaining===0){' +
    '      h+="<div class=\\"days-num red\\">Today</div>";' +
    '    }else{' +
    '      h+="<div class=\\"days-num "+d.urgency+"\\">"+d.daysRemaining+"</div>"' +
    '        +"<div class=\\"days-label\\">days</div>";' +
    '    }' +
    '    h+="</div></div>";' +
    '  });' +
    '  el.innerHTML=h;' +
    '}' +
    'renderList();' +

    // ---- MONTH VIEW ----
    'function renderMonth(){' +
    '  var el=document.getElementById("monthView");' +
    '  var monthNames=["January","February","March","April","May","June",' +
    '    "July","August","September","October","November","December"];' +
    '  var dayNames=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];' +

    '  var firstDay=new Date(currentYear,currentMonth,1).getDay();' +
    '  var daysInMonth=new Date(currentYear,currentMonth+1,0).getDate();' +
    '  var today=new Date();today.setHours(0,0,0,0);' +

    // Build deadline lookup by date
    '  var dateMap={};' +
    '  deadlines.forEach(function(d,i){' +
    '    var key=d.dueDate;' +
    '    if(!dateMap[key])dateMap[key]=[];' +
    '    dateMap[key].push({idx:i,d:d});' +
    '  });' +

    // Month navigation
    '  var h="<div class=\\"month-nav\\">"' +
    '    +"<button class=\\"month-btn\\" onclick=\\"prevMonth()\\">&#x25C0;</button>"' +
    '    +"<div class=\\"month-title\\">"+monthNames[currentMonth]+" "+currentYear+"</div>"' +
    '    +"<button class=\\"month-btn\\" onclick=\\"nextMonth()\\">&#x25B6;</button>"' +
    '    +"</div>";' +

    // Calendar grid header
    '  h+="<div class=\\"cal-grid\\">";' +
    '  dayNames.forEach(function(dn){h+="<div class=\\"cal-header\\">"+dn+"</div>"});' +

    // Days from previous month
    '  var prevMonthDays=new Date(currentYear,currentMonth,0).getDate();' +
    '  for(var p=firstDay-1;p>=0;p--){' +
    '    h+="<div class=\\"cal-day other-month\\"><div class=\\"cal-day-num\\">"+(prevMonthDays-p)+"</div></div>";' +
    '  }' +

    // Days of current month
    '  for(var d=1;d<=daysInMonth;d++){' +
    '    var dateObj=new Date(currentYear,currentMonth,d);' +
    '    var isToday=(dateObj.getTime()===today.getTime());' +
    '    var mm=("0"+(currentMonth+1)).slice(-2);' +
    '    var dd=("0"+d).slice(-2);' +
    '    var key=currentYear+"-"+mm+"-"+dd;' +
    '    var events=dateMap[key]||[];' +

    '    h+="<div class=\\"cal-day"+(isToday?" today":"")+"\\"><div class=\\"cal-day-num\\">"+d+"</div>";' +
    '    events.forEach(function(ev){' +
    '      h+="<div class=\\"cal-event "+ev.d.urgency+"\\" onclick=\\"showDetail("+ev.idx+")\\" title=\\""+escapeHtml(ev.d.grievanceId)+"\\">"+escapeHtml(ev.d.grievanceId)+"</div>";' +
    '    });' +
    '    h+="</div>";' +
    '  }' +

    // Fill remaining cells to complete grid
    '  var totalCells=firstDay+daysInMonth;' +
    '  var remainder=totalCells%7;' +
    '  if(remainder>0){' +
    '    for(var r=1;r<=7-remainder;r++){' +
    '      h+="<div class=\\"cal-day other-month\\"><div class=\\"cal-day-num\\">"+r+"</div></div>";' +
    '    }' +
    '  }' +

    '  h+="</div>";' +
    '  el.innerHTML=h;' +
    '}' +
    'renderMonth();' +

    'function prevMonth(){' +
    '  currentMonth--;if(currentMonth<0){currentMonth=11;currentYear--}' +
    '  renderMonth();' +
    '}' +
    'function nextMonth(){' +
    '  currentMonth++;if(currentMonth>11){currentMonth=0;currentYear++}' +
    '  renderMonth();' +
    '}' +

    // ---- VIEW TOGGLE ----
    'function showView(view){' +
    '  document.getElementById("listView").style.display=view==="list"?"flex":"none";' +
    '  document.getElementById("monthView").style.display=view==="month"?"block":"none";' +
    '  document.getElementById("btnList").classList.toggle("active",view==="list");' +
    '  document.getElementById("btnMonth").classList.toggle("active",view==="month");' +
    '}' +

    // ---- DETAIL PANEL ----
    'function showDetail(idx){' +
    '  var d=deadlines[idx];if(!d)return;' +
    '  var urgLabel=d.urgency==="red"?"Critical":d.urgency==="orange"?"Approaching":"On Track";' +
    '  var daysText=d.daysRemaining<0?Math.abs(d.daysRemaining)+" days overdue"' +
    '    :d.daysRemaining===0?"Due today":d.daysRemaining+" days remaining";' +
    '  var h="<div class=\\"detail-header\\">"' +
    '    +"<div class=\\"detail-id\\">"+escapeHtml(d.grievanceId)+"</div>"' +
    '    +"<button class=\\"detail-close\\" onclick=\\"closeDetail(event,true)\\">&#x2715;</button>"' +
    '    +"</div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Member</span><span class=\\"detail-value\\">"+escapeHtml(d.memberName)+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Current Step</span><span class=\\"detail-value\\">"+escapeHtml(d.currentStep)+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Due Date</span><span class=\\"detail-value\\">"+escapeHtml(d.dueDate)+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Time Remaining</span><span class=\\"detail-value\\">"+escapeHtml(daysText)+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Urgency</span><span class=\\"detail-urgency "+d.urgency+"\\">"+urgLabel+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Status</span><span class=\\"detail-value\\">"+escapeHtml(d.status)+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Steward</span><span class=\\"detail-value\\">"+(d.steward||"Unassigned")+"</span></div>"' +
    '    +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Category</span><span class=\\"detail-value\\">"+(d.issueCategory||"N/A")+"</span></div>";' +
    '  document.getElementById("detailPanel").innerHTML=h;' +
    '  document.getElementById("detailOverlay").classList.add("show");' +
    '}' +

    'function closeDetail(e,force){' +
    '  if(force||e.target.id==="detailOverlay"){' +
    '    document.getElementById("detailOverlay").classList.remove("show");' +
    '  }' +
    '}' +

    // Keyboard: Escape closes detail
    'document.addEventListener("keydown",function(e){if(e.key==="Escape")closeDetail(e,true)});' +

    '</script></body></html>';

  return html;
}


// ============================================================================
// NOTIFICATIONS API (v4.12.0)
// ============================================================================
// Steward-composed, member-dismissable notifications.
// Data lives in 📢 Notifications sheet. Shown in member web view.
// Persist until Expires date set by steward OR dismissed by member.
// Steward composes via separate form accessible from steward dashboard.
// ============================================================================

/**
 * Get active notifications for a user.
 * Filters: Status=Active, not expired, not dismissed by this user.
 * Matches: direct email recipient, "All Members", "All Stewards", "Everyone"
 * @param {string} userEmail — the logged-in user's email
 * @param {string} [userRole] — "steward" or "member" for audience matching
 * @returns {Object[]} array of notification objects
 */
function getWebAppNotifications(sessionToken, userRole) {
  try {
    // CR-AUTH: resolve identity server-side — never trust client-supplied email
    var userEmail = _resolveCallerEmail(sessionToken);
    if (!userEmail) return [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = NOTIFICATIONS_COLS;
    var now = new Date();
    var results = [];
    userEmail = userEmail.toLowerCase().trim();
    // CR-AUTH: derive role from server-verified identity, never trust client-supplied role param
    var auth = checkWebAppAuthorization(null, sessionToken);
    var serverRole = (auth && auth.role) ? auth.role.toLowerCase() : 'member';
    // Treat 'both' (dual-role) and 'admin' as steward for notification targeting
    userRole = (serverRole === 'steward' || serverRole === 'both' || serverRole === 'admin') ? 'steward' : 'member';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Status must be Active
      var status = String(row[C.STATUS - 1] || '').trim();
      if (status !== 'Active') continue;

      // Check expiry
      var expires = row[C.EXPIRES_DATE - 1];
      if (expires && expires instanceof Date && expires < now) continue;
      if (expires && typeof expires === 'string' && expires.trim()) {
        var parsed = new Date(expires);
        if (!isNaN(parsed.getTime()) && parsed < now) continue;
      }

      // Check dismissed
      var dismissed = String(row[C.DISMISSED_BY - 1] || '');
      var dismissedList = dismissed.split(',').map(function(e) { return e.trim().toLowerCase(); });
      if (userEmail && dismissedList.indexOf(userEmail) !== -1) continue;

      // Check recipient match
      var recipient = String(row[C.RECIPIENT - 1] || '').trim().toLowerCase();
      var matches = false;
      if (recipient === 'everyone') matches = true;
      else if (recipient === 'all members') matches = true;
      else if (recipient === 'all stewards' && userRole === 'steward') matches = true;
      else if (recipient === userEmail) matches = true;
      if (!matches) continue;

      results.push({
        id: String(row[C.NOTIFICATION_ID - 1] || ''),
        recipient: String(row[C.RECIPIENT - 1] || ''),
        type: String(row[C.TYPE - 1] || 'System'),
        title: String(row[C.TITLE - 1] || ''),
        message: String(row[C.MESSAGE - 1] || ''),
        priority: String(row[C.PRIORITY - 1] || 'Normal'),
        sentBy: String(row[C.SENT_BY_NAME - 1] || ''),
        createdDate: row[C.CREATED_DATE - 1] ? Utilities.formatDate(
          row[C.CREATED_DATE - 1] instanceof Date ? row[C.CREATED_DATE - 1] : new Date(row[C.CREATED_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        expiresDate: row[C.EXPIRES_DATE - 1] ? Utilities.formatDate(
          row[C.EXPIRES_DATE - 1] instanceof Date ? row[C.EXPIRES_DATE - 1] : new Date(row[C.EXPIRES_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        // v4.22.0: 'Dismissible' (default) | 'Timed'. Empty cells on existing sheets default to Dismissible.
        dismissMode: String(row[C.DISMISS_MODE - 1] || 'Dismissible'),
        rowIndex: i + 1
      });
    }

    // FIX (v4.22.0): reverse first so newest is at index 0, THEN stable-sort Urgent to top.
    // Previous code reversed AFTER sorting, which sent Urgent to the bottom.
    results.reverse();
    results.sort(function(a, b) {
      if (a.priority === 'Urgent' && b.priority !== 'Urgent') return -1;
      if (b.priority === 'Urgent' && a.priority !== 'Urgent') return 1;
      return 0;
    });

    return results;
  } catch (e) {
    logError_('getWebAppNotifications', e);
    return [];
  }
}


/**
 * Lightweight notification count for SPA bell badge.
 * Reuses same filtering logic as getWebAppNotifications but returns only count.
 * @param {string} userEmail
 * @param {string} userRole — 'member' or 'steward'
 * @returns {Object} { count: number }
 */
function getWebAppNotificationCount(sessionToken, userRole) {
  try {
    var results = getWebAppNotifications(sessionToken, userRole);
    return { count: results.length };
  } catch (e) {
    logError_('getWebAppNotificationCount', e);
    return { count: 0 };
  }
}


/**
 * Dismiss a notification for a specific user.
 * Appends user's email to the Dismissed_By column (comma-separated).
 * @param {string} notificationId — e.g. "NOTIF-001"
 * @param {string} userEmail — the user dismissing
 * @returns {Object} { success: boolean, message: string }
 */
function dismissWebAppNotification(sessionToken, notificationId) {
  try {
    // CR-AUTH-4: Use server-side identity — never trust client-supplied email
    var userEmail = _resolveCallerEmail(sessionToken);
    if (!userEmail) return { success: false, message: 'Authentication required' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return { success: false, message: 'Notifications sheet not found' };

    var data = sheet.getDataRange().getValues();
    var C = NOTIFICATIONS_COLS;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][C.NOTIFICATION_ID - 1] || '').trim();
      if (rowId === notificationId) {
        var existing = String(data[i][C.DISMISSED_BY - 1] || '').trim();
        var emails = existing ? existing.split(',').map(function(e) { return e.trim().toLowerCase(); }) : [];
        if (emails.indexOf(userEmail) === -1) {
          emails.push(userEmail);
          sheet.getRange(i + 1, C.DISMISSED_BY).setValue(emails.join(', '));
        }
        return { success: true, message: 'Notification dismissed' };
      }
    }

    return { success: false, message: 'Notification not found' };
  } catch (e) {
    logError_('dismissWebAppNotification', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Send a new notification (steward form submission).
 * Creates a new row in the Notifications sheet.
 * @param {Object} data — { recipient, type, title, message, priority, expiresDate }
 * @returns {Object} { success: boolean, notificationId: string, message: string }
 */
function sendWebAppNotification(sessionToken, data) {
  try {
    // CR-AUTH-6: Verify steward role before allowing notification creation
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) {
      sheet = createNotificationsSheet(ss);
    }

    // Validate required fields
    if (!data.title || !data.message) {
      return { success: false, message: 'Title and message are required' };
    }
    if (!data.recipient) {
      return { success: false, message: 'Recipient is required' };
    }

    // Get steward info from session
    var stewardEmail = auth.email || 'unknown';
    var stewardName = data.senderName || stewardEmail.split('@')[0];

    // Generate next ID
    var allData = sheet.getDataRange().getValues();
    var maxNum = 0;
    var C = NOTIFICATIONS_COLS;
    for (var i = 1; i < allData.length; i++) {
      var existId = String(allData[i][C.NOTIFICATION_ID - 1] || '');
      var match = existId.match(/NOTIF-(\d+)/);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
    var nextId = 'NOTIF-' + String(maxNum + 1).padStart(3, '0');

    var tz = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Build row (must match NOTIFICATIONS_HEADER_MAP_ order)
    // CR-FORMULA: Escape all user-supplied fields to prevent formula injection
    var newRow = [
      nextId,
      escapeForFormula(data.recipient || 'All Members'),
      escapeForFormula(data.type || 'Steward Message'),
      escapeForFormula(data.title),
      escapeForFormula(data.message),
      escapeForFormula(data.priority || 'Normal'),
      stewardEmail,
      escapeForFormula(stewardName),
      today,
      escapeForFormula(data.expiresDate || ''),
      '',
      'Active',
      escapeForFormula(data.dismissMode || 'Dismissible')  // v4.22.0
    ];

    sheet.appendRow(newRow);

    return { success: true, notificationId: nextId, message: 'Notification sent' };
  } catch (e) {
    logError_('sendWebAppNotification', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Returns ALL notifications from the sheet regardless of recipient/expiry.
 * Steward-only — used by the Manage sub-tab to give stewards a full ledger view.
 * Includes dismissedCount so stewards can see reach/engagement.
 * @returns {Object[]} array of notification objects with status + dismissedCount
 */
function getAllWebAppNotifications(sessionToken) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = NOTIFICATIONS_COLS;
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var id = String(row[C.NOTIFICATION_ID - 1] || '').trim();
      if (!id) continue;

      // Count dismissals for engagement metric
      var dismissedRaw = String(row[C.DISMISSED_BY - 1] || '').trim();
      var dismissedCount = dismissedRaw
        ? dismissedRaw.split(',').filter(function(e) { return e.trim(); }).length
        : 0;

      results.push({
        id: id,
        recipient: String(row[C.RECIPIENT - 1] || ''),
        type: String(row[C.TYPE - 1] || 'System'),
        title: String(row[C.TITLE - 1] || ''),
        message: String(row[C.MESSAGE - 1] || ''),
        priority: String(row[C.PRIORITY - 1] || 'Normal'),
        sentBy: String(row[C.SENT_BY_NAME - 1] || ''),
        createdDate: row[C.CREATED_DATE - 1] ? Utilities.formatDate(
          row[C.CREATED_DATE - 1] instanceof Date ? row[C.CREATED_DATE - 1] : new Date(row[C.CREATED_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        expiresDate: row[C.EXPIRES_DATE - 1] ? Utilities.formatDate(
          row[C.EXPIRES_DATE - 1] instanceof Date ? row[C.EXPIRES_DATE - 1] : new Date(row[C.EXPIRES_DATE - 1]),
          Session.getScriptTimeZone(), 'MMM d, yyyy') : '',
        status: String(row[C.STATUS - 1] || 'Active'),
        dismissMode: String(row[C.DISMISS_MODE - 1] || 'Dismissible'),
        dismissedCount: dismissedCount,
        rowIndex: i + 1
      });
    }

    // Newest first
    results.reverse();
    return results;
  } catch (e) {
    logError_('getAllWebAppNotifications', e);
    return [];
  }
}


/**
 * Archives a notification (sets Status = 'Archived').
 * Steward-only. Archived notifications are excluded from all member views
 * because getWebAppNotifications() filters Status = 'Active' only.
 * Non-destructive — row stays in sheet, data is preserved.
 * @param {string} notificationId — e.g. 'NOTIF-003'
 * @returns {Object} { success, message }
 */
function archiveWebAppNotification(sessionToken, notificationId) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
    if (!sheet) return { success: false, message: 'Notifications sheet not found.' };

    var data = sheet.getDataRange().getValues();
    var C = NOTIFICATIONS_COLS;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][C.NOTIFICATION_ID - 1] || '').trim();
      if (rowId === notificationId) {
        sheet.getRange(i + 1, C.STATUS).setValue('Archived');
        return { success: true, message: 'Notification archived.' };
      }
    }

    return { success: false, message: 'Notification not found.' };
  } catch (e) {
    logError_('archiveWebAppNotification', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}


/**
 * Get member list for notification recipient picker (steward use).
 * Returns simplified list: name + email for dropdown.
 * @returns {Object[]} [{ name, email }]
 */
function getNotificationRecipientList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = MEMBER_COLS;
    var results = [];
    // Preset groups first
    results.push({ name: '📢 All Members', email: 'All Members' });
    results.push({ name: '🛡️ All Stewards', email: 'All Stewards' });
    results.push({ name: '🌐 Everyone', email: 'Everyone' });

    for (var i = 1; i < data.length; i++) {
      var firstName = String(data[i][C.FIRST_NAME - 1] || '').trim();
      var lastName = String(data[i][C.LAST_NAME - 1] || '').trim();
      var email = String(data[i][C.EMAIL - 1] || '').trim();
      if (!email) continue;
      var fullName = (firstName + ' ' + lastName).trim();
      results.push({ name: fullName || email, email: email });
    }

    return results;
  } catch (e) {
    logError_('getNotificationRecipientList', e);
    return [];
  }
}


/**
 * Get full member list with directory columns for filtering/sorting.
 * Used by steward notification compose form recipient picker.
 * Returns: name, email, location, department, jobTitle for each member.
 * @returns {Object[]}
 */
function getNotificationRecipientListFull() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var C = MEMBER_COLS;
    var results = [];

    for (var i = 1; i < data.length; i++) {
      var firstName = String(data[i][C.FIRST_NAME - 1] || '').trim();
      var lastName = String(data[i][C.LAST_NAME - 1] || '').trim();
      var email = String(data[i][C.EMAIL - 1] || '').trim();
      if (!email) continue;

      var fullName = (firstName + ' ' + lastName).trim();
      var location = '';
      var department = '';
      var jobTitle = '';

      if (C.WORK_LOCATION) location = String(data[i][C.WORK_LOCATION - 1] || '').trim();
      if (C.DEPARTMENT) department = String(data[i][C.DEPARTMENT - 1] || '').trim();
      if (C.JOB_TITLE) jobTitle = String(data[i][C.JOB_TITLE - 1] || '').trim();

      results.push({
        name: fullName || email,
        email: email,
        location: location,
        department: department,
        jobTitle: jobTitle
      });
    }

    results.sort(function(a, b) { return a.name.localeCompare(b.name); });
    return results;
  } catch (e) {
    logError_('getNotificationRecipientListFull', e);
    return [];
  }
}


// NOTE (v4.22.0): getWebAppNotificationsHtml() removed — standalone ?page=notifications
// route was never wired in doGet(). Notifications fully handled by SPA (index.html).


// ─── ONE-TIME MIGRATIONS ────────────────────────────────────────────────────

/**
 * ONE-TIME MIGRATION (v4.22.0) — Add Dismiss_Mode column to existing Notifications sheet.
 * Safe to re-run: checks for column before modifying anything.
 * After running successfully, this function can be deleted from the codebase.
 *
 * HOW TO RUN: Open Apps Script editor → select this function → click Run.
 * Expected output in Logs: "Migration complete" or "Already has Dismiss_Mode column".
 */
function MIGRATE_ADD_DISMISS_MODE_COLUMN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.NOTIFICATIONS);
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_e) {}

  if (!sheet) {
    var msg = 'Notifications sheet not found. Nothing to migrate.';
    Logger.log(msg);
    if (ui) ui.alert(msg);
    return;
  }

  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if column already exists (case-insensitive)
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]).toLowerCase() === 'dismiss_mode') {
      var alreadyMsg = 'Already has Dismiss_Mode column at col ' + (i + 1) + '. No changes made.';
      Logger.log(alreadyMsg);
      if (ui) ui.alert(alreadyMsg);
      return;
    }
  }

  // Append as next column after current last column
  var newColIndex = sheet.getLastColumn() + 1;

  // Header cell
  var headerCell = sheet.getRange(1, newColIndex);
  headerCell.setValue('Dismiss_Mode')
    .setBackground(COLORS.HEADER_BG || '#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(newColIndex, 120);

  // Data validation on data rows
  var dismissModeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Dismissible', 'Timed'])
    .setAllowInvalid(false)
    .build();
  var lastDataRow = sheet.getLastRow();
  if (lastDataRow >= 2) {
    sheet.getRange(2, newColIndex, lastDataRow - 1).setDataValidation(dismissModeRule);
    // Backfill all existing rows with 'Dismissible' (safe default — preserves legacy behaviour)
    sheet.getRange(2, newColIndex, lastDataRow - 1).setValue('Dismissible');
  }

  var successMsg = 'Migration complete. Dismiss_Mode column added at col ' + newColIndex +
    '. All ' + (lastDataRow - 1) + ' existing rows backfilled with "Dismissible".';
  Logger.log(successMsg);
  if (ui) ui.alert(successMsg);
}

// ─── MANUAL RE-RUN WRAPPERS (v4.20.17) ─────────────────────────────────────
/**
 * Standalone wrapper — run from Apps Script editor or menu to
 * re-create/repair the DashboardTest Drive folder structure.
 */
function SETUP_DRIVE_FOLDERS() {
  var result = setupDashboardDriveFolders();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_e) {}
  var msg = result.success
    ? '✅ Drive folders ready!\n\n' + result.rootFolderName + ': ' + result.rootFolderUrl +
      '\n\nAll folder IDs have been saved to the Config sheet.' +
      '\n\n📄 Minutes/ folder: set to "Anyone with link can view" (members can browse docs).' +
      '\n🔒 All other subfolders: PRIVATE.'
    : '❌ Drive folder setup failed — check Apps Script logs.';
  if (ui) ui.alert('📁 Drive Folder Setup', msg, ui.ButtonSet.OK);
  else Logger.log('SETUP_DRIVE_FOLDERS: ' + JSON.stringify(result));
  return result;
}

/**
 * Standalone wrapper — run from Apps Script editor or menu to
 * re-create/repair the union events calendar.
 */
function SETUP_CALENDAR() {
  var result = setupDashboardCalendar();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_e) {}
  var msg = result.success
    ? '✅ Calendar ready!\n\nCalendar: ' + result.calendarName +
      '\nID: ' + result.calendarId +
      '\n\nCalendar ID has been saved to the Config sheet.'
    : '❌ Calendar setup failed — check Apps Script logs.';
  if (ui) ui.alert('📅 Calendar Setup', msg, ui.ButtonSet.OK);
  else Logger.log('SETUP_CALENDAR: ' + JSON.stringify(result));
  return result;
}
