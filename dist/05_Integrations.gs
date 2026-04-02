/**
 * ============================================================================
 * 05_Integrations.gs - External Service Integration
 * ============================================================================
 *
 * WHAT THIS FILE DOES:
 *   All external Google service integrations — Drive (folder hierarchy for
 *   grievance documents), Calendar (deadline/meeting sync), Email
 *   (notifications, PDF delivery), and Constant Contact API (member
 *   engagement tracking). Also contains CALENDAR_CONFIG constants.
 *
 * WHY IT EXISTS / DESIGN DECISIONS:
 *   Isolates all external service dependencies so that if Drive/Calendar/email
 *   has an outage, core spreadsheet functionality stays responsive. Drive
 *   folders follow a strict hierarchy (DashboardTest/ -> Grievances/,
 *   Resources/, Minutes/, Event Check-In/) with PRIVATE access controls.
 *   Calendar sync creates events for grievance deadlines with configurable
 *   reminders at 7, 3, and 1 day.
 *
 * WHAT HAPPENS IF THIS FILE BREAKS:
 *   Grievance documents won't be filed to Drive. Calendar deadlines won't
 *   sync. Email notifications stop. Constant Contact engagement data goes
 *   stale.
 *
 * DEPENDENCIES:
 *   Depends on DriveApp, CalendarApp, MailApp/GmailApp (GAS built-ins),
 *   01_Core.gs (SHEETS, GRIEVANCE_COLS, CONFIG_COLS). Used by menu items
 *   in 03_, trigger functions in 10_, and the SPA.
 *
 * @fileoverview External service integrations
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
  if (!configSheet) throw new Error('Config sheet not found');

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
  // H6: Treat sharing failure as fatal — folder must be private before proceeding
  try {
    rootFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (shareErr) {
    Logger.log('setupDashboardDriveFolders: FATAL — could not set root folder to PRIVATE: ' + shareErr.message);
    throw new Error('Could not set root folder sharing to PRIVATE. Aborting setup to prevent data exposure. Error: ' + shareErr.message);
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
    if (folderUrl) {
      configSheet.getRange(3, colIndex + 1).setValue(folderUrl);
    }
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
  if (!configSheet) return { success: false, message: 'Config sheet not found' };

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
  } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
    if (!ss) return null;
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
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      throw new Error('Could not acquire lock for folder creation');
    }
    try {
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
    } finally {
      lock.releaseLock();
    }

    // ── 5. Write master folder URL to Member Directory (non-fatal) ───────────
    if (isNewMaster) {
      try {
        var folderUrl = masterFolder.getUrl();
        if (mSheet) {
          // M8: Reuse cached data — only the URL column is written, headers don't change
          var mData2    = mData;
          var mHeaders2 = mHeaders;
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
        try { membersRoot = DriveApp.getFolderById(membersRootId); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
      }
      if (!membersRoot) {
        // Resolve Members/ folder explicitly — do NOT fall back to Grievances/ root.
        // Same resolution logic as getOrCreateMemberAdminFolder: dashboard root → create Members/.
        var dashRootId = props.getProperty('DASHBOARD_ROOT_FOLDER_ID') || '';
        if (dashRootId) {
          try {
            var dashRoot = DriveApp.getFolderById(dashRootId);
            var mIter = dashRoot.getFoldersByName(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
            membersRoot = mIter.hasNext() ? mIter.next() : dashRoot.createFolder(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
          } catch (_e) { Logger.log('setupDriveFolderForGrievance fallback: ' + _e.message); }
        }
        if (!membersRoot) {
          var fallbackIter = DriveApp.getFoldersByName(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
          membersRoot = fallbackIter.hasNext() ? fallbackIter.next() : DriveApp.createFolder(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
        }
      }
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
  // Since v4.20+ case folders live under Members/<LastName>/Grievances/<CaseId>/,
  // not under the legacy Grievances/ root. Scan the Members/ folder instead.
  var rootFolder;
  try {
    var props = PropertiesService.getScriptProperties();
    var membersId = props.getProperty('MEMBERS_FOLDER_ID') || '';
    if (membersId) {
      try { rootFolder = DriveApp.getFolderById(membersId); } catch (_e) { rootFolder = null; }
    }
    if (!rootFolder) {
      // Resolve Members/ via dashboard root (same approach as getOrCreateMemberAdminFolder)
      var dashRootId = props.getProperty('DASHBOARD_ROOT_FOLDER_ID') || '';
      if (dashRootId) {
        try {
          var dashRoot = DriveApp.getFolderById(dashRootId);
          var mIter = dashRoot.getFoldersByName(DRIVE_CONFIG.MEMBERS_SUBFOLDER);
          if (mIter.hasNext()) rootFolder = mIter.next();
        } catch (_e) { /* fall through */ }
      }
    }
    if (!rootFolder) {
      // Ultimate fallback: legacy Grievances/ root for pre-v4.20 installs
      rootFolder = getOrCreateRootFolder();
    }
  } catch (_e) {
    return { success: false, reason: 'Could not access root folder' };
  }

  var removed = 0;
  var skipped = 0;
  // Walk member folders → Grievances/ subfolders → case folders
  var subFolders = rootFolder.getFolders();

  // Build a set of resolved grievance IDs for quick lookup
  var resolvedIds = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet');
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

  // Structure: Members/<LastName, FirstName>/Grievances/<CaseId - Date>/
  // Walk each member folder, then into its Grievances/ subfolder, then check case folders.
  while (subFolders.hasNext()) {
    var memberFolder = subFolders.next();
    try {
      var grievancesIter = memberFolder.getFoldersByName('Grievances');
      if (!grievancesIter.hasNext()) { skipped++; continue; }
      var grievancesFolder = grievancesIter.next();
      var caseFolders = grievancesFolder.getFolders();

      while (caseFolders.hasNext()) {
        var folder = caseFolders.next();
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
          Logger.log('cleanupEmptyDriveFolders: error on case folder: ' + folderErr.message);
          skipped++;
        }
      }
    } catch (memberErr) {
      Logger.log('cleanupEmptyDriveFolders: error on member folder: ' + memberErr.message);
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
function _isFolderTreeEmpty(folder, depth) {
  depth = depth || 0;
  if (depth > 10) return false;
  if (folder.getFiles().hasNext()) return false;
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    if (!_isFolderTreeEmpty(subs.next(), depth + 1)) return false;
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
    if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - BATCH_LIMITS.EXECUTION_BUFFER_MS) {
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
  // CR-09: Use canonical SHEETS.GRIEVANCE_LOG (alias GRIEVANCE_TRACKER removed in v4.25.9)
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

  // L3: Validate and filter recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',').filter(function(e) {
    return e.trim() && emailRegex.test(e.trim());
  });
  if (emails.length === 0) {
    return errorResponse('No valid email addresses provided');
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
    result.success = false;
    result.error = error.message;
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

  // L3: Validate and filter recipient email addresses
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var emails = String(recipientEmails).split(',').filter(function(e) {
    return e.trim() && emailRegex.test(e.trim());
  });
  if (emails.length === 0) {
    Logger.log('No valid email addresses in recipient list');
    return;
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
      var dateStr = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');

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
      if (new Date().getTime() - startTime > BATCH_LIMITS.MAX_EXECUTION_TIME_MS - BATCH_LIMITS.EXECUTION_BUFFER_MS) {
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
    Logger.log('Error syncing to calendar: ' + error);
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
  const memberName = ((grievance['First Name'] || '') + ' ' + (grievance['Last Name'] || '')).trim();
  const currentStep = grievance['Current Step'];

  // Get the deadline for current step
  let deadline;
  switch (currentStep) {
    case 'Step I':
    case 'Informal':
    case '1':
      deadline = grievance['Step I Due'];
      break;
    case 'Step II':
    case '2':
      deadline = grievance['Step II Due'];
      break;
    case 'Step III':
    case 'Arbitration':
    case '3':
      deadline = grievance['Step III Appeal Due'];
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
    } catch (_lookupErr) { Logger.log('_lookupErr: ' + (_lookupErr.message || _lookupErr)); }
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
    Logger.log('Error sending reminders: ' + error);
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
    // Sanitize HTML body — strip tags if not explicitly HTML, or escape dangerous patterns
    var safeBody = String(body || '').replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '');

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
    Logger.log('Error sending email: ' + error);
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
    folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  } catch (e) {
    Logger.log('Archive folder not found, using root: ' + e.message);
    rootFolder = getOrCreateRootFolder();
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
 * Menu wrapper: creates a template-based PDF for the selected grievance row.
 * Uses createGrievancePDF() from 11_CommandHub.gs (Docs template with placeholders).
 */
function createTemplatePDFForSelectedGrievance() {
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

  var data = {
    name: sheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue() + ' ' +
          sheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue(),
    id: sheet.getRange(row, GRIEVANCE_COLS.MEMBER_ID).getValue(),
    details: sheet.getRange(row, GRIEVANCE_COLS.RESOLUTION).getValue() || 'Pending'
  };

  var response = ui.alert(
    'Create Template PDF',
    'Create a Docs-template PDF for ' + data.name + '?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    ss.toast('Creating template PDF...', 'System', 5);
    var folder = getOrCreateMemberFolder(data.name, data.id);
    var pdf = createGrievancePDF(folder, data);
    ss.toast('Template PDF created: ' + pdf.getName(), 'System', 5);
  } catch (e) {
    ui.alert('Error', 'Failed to create template PDF: ' + e.message, ui.ButtonSet.OK);
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
  if (!e || !e.namedValues) return;
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
  if (!configSheet) return links;

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
    } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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
  } catch (_re) { Logger.log('_re: ' + (_re.message || _re)); }

  return links;
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
  linkCell.setFontColor(SHEET_COLORS.LINK_PRIMARY);
  linkCell.setBackground(SHEET_COLORS.BG_LINK_BLUE);

  // Also add plain URL below for copying
  var urlCell = configSheet.getRange(4, targetCol);
  urlCell.setValue(url);
  urlCell.setFontSize(10);
  urlCell.setWrap(true);

  // Add instructions
  var instructionCell = configSheet.getRange(5, targetCol);
  instructionCell.setValue('Open Google Sheets on your phone, navigate to Config tab, and tap the blue link above to access the dashboard.');
  instructionCell.setFontSize(9);
  instructionCell.setFontColor(SHEET_COLORS.TEXT_MED_GRAY);
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
    'h2{color:' + SHEET_COLORS.DIALOG_ACCENT + ';margin-top:0}' +
    '.step{background:#f0f7ff;padding:15px;margin:12px 0;border-radius:8px;border-left:4px solid ' + SHEET_COLORS.DIALOG_ACCENT + '}' +
    '.step-num{font-weight:bold;color:' + SHEET_COLORS.DIALOG_ACCENT + '}' +
    'a{color:' + SHEET_COLORS.DIALOG_ACCENT + ';word-break:break-all}' +
    'input{width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;margin:8px 0;font-size:14px}' +
    'button{padding:12px 24px;background:' + SHEET_COLORS.DIALOG_ACCENT + ';color:white;border:none;border-radius:4px;cursor:pointer;font-size:14px}' +
    'button:hover{background:' + SHEET_COLORS.DIALOG_ACCENT_DARK + '}' +
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
    // Steward-only: returns hidden resources + addedBy emails
    var auth = checkWebAppAuthorization('steward');
    if (!auth.isAuthorized) return [];

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

    return withScriptLock_(function() {
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
    });
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

    return withScriptLock_(function() {
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
    });
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

    return withScriptLock_(function() {
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
    });
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
    if (!auth.isAuthorized) return { success: false, message: auth.message || 'Unauthorized' };

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

// ============================================================================
// KNOWLEDGE ENGINE API (v4.41.0)
// ============================================================================
// Server-side functions for the Knowledge Engine sheet. Provides cached
// retrieval of educational content by placement, type, and audience.
// All content is read from the 📚 Knowledge Engine sheet.
// ============================================================================

/**
 * Returns active knowledge content filtered by placement and audience.
 * Uses CacheService for 5-minute server-side caching to reduce sheet reads.
 * Called by the webapp HTML files to replace hardcoded content arrays.
 *
 * @param {string} [placement] - Filter by placement: 'Auth Screen', 'Home Widget', 'Sidebar', 'Hub Section', 'Inline Quote'
 * @param {string} [audience]  - Filter by audience: 'All', 'Members', 'Stewards'
 * @param {string} [type]      - Filter by type: 'Quote', 'Tip', 'Concept', 'Mini-Lesson', 'Manifesto', 'Negotiation Set'
 * @returns {Array<Object>} Knowledge content objects
 */
function getKnowledgeContent(placement, audience, type) {
  try {
    // Build cache key from filter params
    var cacheKey = 'ke_' + (placement || 'all') + '_' + (audience || 'all') + '_' + (type || 'all');
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (_) { /* cache corrupt, continue */ }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);

    // Auto-create if missing
    if (!sheet) {
      if (typeof createKnowledgeEngineSheet === 'function') {
        sheet = createKnowledgeEngineSheet(ss);
      }
      if (!sheet) return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, KNOWLEDGE_COLS.ADDED_BY).getValues();
    var now = new Date();
    var result = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // Filter: Active
      var active = String(row[KNOWLEDGE_COLS.ACTIVE - 1] || '').toLowerCase();
      if (active !== 'yes') continue;

      // Filter: Placement
      if (placement) {
        var rowPlacement = row[KNOWLEDGE_COLS.PLACEMENT - 1] || '';
        if (rowPlacement !== placement) continue;
      }

      // Filter: Audience
      if (audience && audience !== 'All') {
        var rowAudience = row[KNOWLEDGE_COLS.AUDIENCE - 1] || 'All';
        if (rowAudience !== 'All' && rowAudience !== audience) continue;
      }

      // Filter: Type
      if (type) {
        var rowType = row[KNOWLEDGE_COLS.TYPE - 1] || '';
        if (rowType !== type) continue;
      }

      // Filter: Date scheduling
      var startDate = row[KNOWLEDGE_COLS.START_DATE - 1];
      if (startDate instanceof Date && startDate > now) continue;
      var endDate = row[KNOWLEDGE_COLS.END_DATE - 1];
      if (endDate instanceof Date && endDate < now) continue;

      // Parse bullets from pipe-delimited string
      var bulletsRaw = row[KNOWLEDGE_COLS.BULLETS - 1] || '';
      var bullets = bulletsRaw ? String(bulletsRaw).split('|').map(function(b) { return b.trim(); }).filter(function(b) { return b; }) : [];

      result.push({
        id: row[KNOWLEDGE_COLS.CONTENT_ID - 1] || '',
        type: row[KNOWLEDGE_COLS.TYPE - 1] || '',
        category: row[KNOWLEDGE_COLS.CATEGORY - 1] || 'General',
        title: row[KNOWLEDGE_COLS.TITLE - 1] || '',
        content: row[KNOWLEDGE_COLS.CONTENT - 1] || '',
        attribution: row[KNOWLEDGE_COLS.ATTRIBUTION - 1] || '',
        bullets: bullets,
        audience: row[KNOWLEDGE_COLS.AUDIENCE - 1] || 'All',
        placement: row[KNOWLEDGE_COLS.PLACEMENT - 1] || '',
        priority: row[KNOWLEDGE_COLS.PRIORITY - 1] || 999
      });
    }

    // Sort by priority
    result.sort(function(a, b) { return (a.priority || 999) - (b.priority || 999); });

    // Cache for 5 minutes (300 seconds)
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch (_) { /* cache write failure non-critical */ }

    return result;
  } catch (e) {
    Logger.log('getKnowledgeContent error: ' + e.toString());
    return [];
  }
}

/**
 * Returns ALL knowledge content for steward management (includes inactive).
 * Steward-only access.
 * @returns {Array<Object>} All knowledge content objects
 */
function getKnowledgeContentAll() {
  try {
    var auth = checkWebAppAuthorization('steward');
    if (!auth.isAuthorized) return [];

    var cache = CacheService.getScriptCache();
    var cacheKey = 'ke_all_steward';
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (_) { /* cache corrupt */ }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, KNOWLEDGE_COLS.ADDED_BY).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var bulletsRaw = row[KNOWLEDGE_COLS.BULLETS - 1] || '';
      var dateAdded = row[KNOWLEDGE_COLS.DATE_ADDED - 1];
      return {
        id: row[KNOWLEDGE_COLS.CONTENT_ID - 1] || '',
        type: row[KNOWLEDGE_COLS.TYPE - 1] || '',
        category: row[KNOWLEDGE_COLS.CATEGORY - 1] || 'General',
        title: row[KNOWLEDGE_COLS.TITLE - 1] || '',
        content: row[KNOWLEDGE_COLS.CONTENT - 1] || '',
        attribution: row[KNOWLEDGE_COLS.ATTRIBUTION - 1] || '',
        bullets: bulletsRaw ? String(bulletsRaw).split('|').map(function(b) { return b.trim(); }).filter(function(b) { return b; }) : [],
        bulletsRaw: String(bulletsRaw),
        audience: row[KNOWLEDGE_COLS.AUDIENCE - 1] || 'All',
        placement: row[KNOWLEDGE_COLS.PLACEMENT - 1] || '',
        active: String(row[KNOWLEDGE_COLS.ACTIVE - 1] || '').toLowerCase() === 'yes',
        priority: row[KNOWLEDGE_COLS.PRIORITY - 1] || 999,
        startDate: row[KNOWLEDGE_COLS.START_DATE - 1] instanceof Date ? Utilities.formatDate(row[KNOWLEDGE_COLS.START_DATE - 1], tz, 'yyyy-MM-dd') : '',
        endDate: row[KNOWLEDGE_COLS.END_DATE - 1] instanceof Date ? Utilities.formatDate(row[KNOWLEDGE_COLS.END_DATE - 1], tz, 'yyyy-MM-dd') : '',
        dateAdded: dateAdded instanceof Date ? Utilities.formatDate(dateAdded, tz, 'MMM d, yyyy') : (dateAdded || ''),
        addedBy: row[KNOWLEDGE_COLS.ADDED_BY - 1] || ''
      };
    }).filter(function(r) { return r.id; });

    // Cache for 2 minutes
    try { cache.put(cacheKey, JSON.stringify(result), 120); } catch (_) {}
    return result;
  } catch (e) {
    logError_('getKnowledgeContentAll', e);
    return [];
  }
}

/**
 * Adds a new knowledge content item. Steward-only.
 * @param {string} sessionToken - Auth session token
 * @param {Object} data - Content fields
 * @returns {Object} { success, contentId, message }
 */
function addKnowledgeContent(sessionToken, data) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    // ── Input validation ──
    var validTypes = ['Quote', 'Tip', 'Concept', 'Mini-Lesson', 'Manifesto', 'Negotiation Set'];
    var validCategories = ['Negotiation', 'Rights', 'Grievance', 'Safety', 'Labor', 'General'];
    var validAudiences = ['All', 'Members', 'Stewards'];
    var validPlacements = ['Auth Screen', 'Home Widget', 'Sidebar', 'Hub Section', 'Inline Quote'];

    if (!data || typeof data !== 'object') return { success: false, message: 'Invalid data' };
    if (!data.content && !data.title) return { success: false, message: 'Title or content is required' };
    if (data.type && validTypes.indexOf(data.type) === -1) return { success: false, message: 'Invalid type' };
    if (data.category && validCategories.indexOf(data.category) === -1) return { success: false, message: 'Invalid category' };
    if (data.audience && validAudiences.indexOf(data.audience) === -1) return { success: false, message: 'Invalid audience' };
    if (data.placement && validPlacements.indexOf(data.placement) === -1) return { success: false, message: 'Invalid placement' };
    if ((data.title || '').length > 500) return { success: false, message: 'Title too long (max 500 chars)' };
    if ((data.content || '').length > 10000) return { success: false, message: 'Content too long (max 10,000 chars)' };
    var priority = parseInt(data.priority, 10) || 999;
    if (priority < 1 || priority > 9999) return { success: false, message: 'Priority must be 1–9999' };
    if (data.startDate && data.endDate) {
      var s = new Date(data.startDate), e = new Date(data.endDate);
      if (!isNaN(s.getTime()) && !isNaN(e.getTime()) && e < s) return { success: false, message: 'End date must be after start date' };
    }

    return withScriptLock_(function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);
      if (!sheet) {
        if (typeof createKnowledgeEngineSheet === 'function') {
          sheet = createKnowledgeEngineSheet(ss);
        }
        if (!sheet) return { success: false, message: 'Knowledge Engine sheet not found' };
      }

      // Generate next ID
      var allData = sheet.getDataRange().getValues();
      var maxNum = 0;
      for (var i = 1; i < allData.length; i++) {
        var existId = String(allData[i][KNOWLEDGE_COLS.CONTENT_ID - 1] || '');
        var match = existId.match(/KE-(\d+)/);
        if (match) {
          var num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }
      }
      var nextId = 'KE-' + String(maxNum + 1).padStart(3, '0');

      var tz = Session.getScriptTimeZone();
      var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
      var addedBy = auth.email || 'unknown';

      // Normalize bullets: accept array or pipe-delimited string
      var bullets = '';
      if (Array.isArray(data.bullets)) {
        bullets = data.bullets.join('|');
      } else if (data.bullets) {
        bullets = String(data.bullets);
      }

      var newRow = [
        nextId,
        escapeForFormula(data.type || 'Tip'),
        escapeForFormula(data.category || 'General'),
        escapeForFormula(data.title || ''),
        escapeForFormula(data.content || ''),
        escapeForFormula(data.attribution || ''),
        escapeForFormula(bullets),
        escapeForFormula(data.audience || 'All'),
        escapeForFormula(data.placement || 'Home Widget'),
        escapeForFormula(data.active || 'Yes'),
        data.priority || 999,
        data.startDate || '',
        data.endDate || '',
        today,
        addedBy
      ];

      sheet.appendRow(newRow);

      // Invalidate cache
      _invalidateKnowledgeCache_();

      return { success: true, contentId: nextId, message: 'Content added' };
    });
  } catch (e) {
    logError_('addKnowledgeContent', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}

/**
 * Updates an existing knowledge content item. Steward-only.
 * @param {string} sessionToken - Auth session token
 * @param {string} contentId - The KE-xxx ID to update
 * @param {Object} data - Fields to update
 * @returns {Object} { success, message }
 */
function updateKnowledgeContent(sessionToken, contentId, data) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    // ── Input validation (partial — only validate provided fields) ──
    var validTypes = ['Quote', 'Tip', 'Concept', 'Mini-Lesson', 'Manifesto', 'Negotiation Set'];
    var validCategories = ['Negotiation', 'Rights', 'Grievance', 'Safety', 'Labor', 'General'];
    var validAudiences = ['All', 'Members', 'Stewards'];
    var validPlacements = ['Auth Screen', 'Home Widget', 'Sidebar', 'Hub Section', 'Inline Quote'];

    if (!data || typeof data !== 'object') return { success: false, message: 'Invalid data' };
    if (data.type !== undefined && validTypes.indexOf(data.type) === -1) return { success: false, message: 'Invalid type' };
    if (data.category !== undefined && validCategories.indexOf(data.category) === -1) return { success: false, message: 'Invalid category' };
    if (data.audience !== undefined && validAudiences.indexOf(data.audience) === -1) return { success: false, message: 'Invalid audience' };
    if (data.placement !== undefined && validPlacements.indexOf(data.placement) === -1) return { success: false, message: 'Invalid placement' };
    if (data.title !== undefined && data.title.length > 500) return { success: false, message: 'Title too long (max 500 chars)' };
    if (data.content !== undefined && data.content.length > 10000) return { success: false, message: 'Content too long (max 10,000 chars)' };
    if (data.priority !== undefined) {
      var pri = parseInt(data.priority, 10);
      if (isNaN(pri) || pri < 1 || pri > 9999) return { success: false, message: 'Priority must be 1–9999' };
    }
    if (data.startDate !== undefined && data.endDate !== undefined && data.startDate && data.endDate) {
      var s = new Date(data.startDate), e = new Date(data.endDate);
      if (!isNaN(s.getTime()) && !isNaN(e.getTime()) && e < s) return { success: false, message: 'End date must be after start date' };
    }

    return withScriptLock_(function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);
      if (!sheet) return { success: false, message: 'Knowledge Engine sheet not found' };

      var allData = sheet.getDataRange().getValues();
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][KNOWLEDGE_COLS.CONTENT_ID - 1] || '').trim() === contentId) {
          var rowData = allData[i].slice();
          if (data.type !== undefined)        rowData[KNOWLEDGE_COLS.TYPE - 1] = escapeForFormula(data.type);
          if (data.category !== undefined)    rowData[KNOWLEDGE_COLS.CATEGORY - 1] = escapeForFormula(data.category);
          if (data.title !== undefined)       rowData[KNOWLEDGE_COLS.TITLE - 1] = escapeForFormula(data.title);
          if (data.content !== undefined)     rowData[KNOWLEDGE_COLS.CONTENT - 1] = escapeForFormula(data.content);
          if (data.attribution !== undefined) rowData[KNOWLEDGE_COLS.ATTRIBUTION - 1] = escapeForFormula(data.attribution);
          if (data.bullets !== undefined) {
            var bullets = Array.isArray(data.bullets) ? data.bullets.join('|') : String(data.bullets);
            rowData[KNOWLEDGE_COLS.BULLETS - 1] = escapeForFormula(bullets);
          }
          if (data.audience !== undefined)    rowData[KNOWLEDGE_COLS.AUDIENCE - 1] = escapeForFormula(data.audience);
          if (data.placement !== undefined)   rowData[KNOWLEDGE_COLS.PLACEMENT - 1] = escapeForFormula(data.placement);
          if (data.active !== undefined)      rowData[KNOWLEDGE_COLS.ACTIVE - 1] = escapeForFormula(data.active);
          if (data.priority !== undefined)    rowData[KNOWLEDGE_COLS.PRIORITY - 1] = data.priority;
          if (data.startDate !== undefined)   rowData[KNOWLEDGE_COLS.START_DATE - 1] = data.startDate;
          if (data.endDate !== undefined)     rowData[KNOWLEDGE_COLS.END_DATE - 1] = data.endDate;
          sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);

          _invalidateKnowledgeCache_();
          return { success: true, message: 'Content updated' };
        }
      }
      return { success: false, message: 'Content not found' };
    });
  } catch (e) {
    logError_('updateKnowledgeContent', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}

/**
 * Soft-deletes an knowledge content item by setting Active to 'No'.
 * @param {string} sessionToken - Auth session token
 * @param {string} contentId - The KE-xxx ID to deactivate
 * @returns {Object} { success, message }
 */
function deleteKnowledgeContent(sessionToken, contentId) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    return withScriptLock_(function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);
      if (!sheet) return { success: false, message: 'Knowledge Engine sheet not found' };

      var allData = sheet.getDataRange().getValues();
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][KNOWLEDGE_COLS.CONTENT_ID - 1] || '').trim() === contentId) {
          sheet.getRange(i + 1, KNOWLEDGE_COLS.ACTIVE).setValue('No');
          _invalidateKnowledgeCache_();
          return { success: true, message: 'Content deactivated' };
        }
      }
      return { success: false, message: 'Content not found' };
    });
  } catch (e) {
    logError_('deleteKnowledgeContent', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}

/**
 * Restores a deactivated knowledge content item.
 * @param {string} sessionToken - Auth session token
 * @param {string} contentId - The KE-xxx ID to reactivate
 * @returns {Object} { success, message }
 */
function restoreKnowledgeContent(sessionToken, contentId) {
  try {
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return { success: false, message: 'Steward access required.' };

    return withScriptLock_(function() {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEETS.KNOWLEDGE_ENGINE);
      if (!sheet) return { success: false, message: 'Knowledge Engine sheet not found' };

      var allData = sheet.getDataRange().getValues();
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][KNOWLEDGE_COLS.CONTENT_ID - 1] || '').trim() === contentId) {
          sheet.getRange(i + 1, KNOWLEDGE_COLS.ACTIVE).setValue('Yes');
          _invalidateKnowledgeCache_();
          return { success: true, message: 'Content restored' };
        }
      }
      return { success: false, message: 'Content not found' };
    });
  } catch (e) {
    logError_('restoreKnowledgeContent', e);
    return { success: false, message: 'Error: ' + String(e) };
  }
}

/**
 * Invalidates all Knowledge Engine cache entries.
 * @private
 */
function _invalidateKnowledgeCache_() {
  try {
    var cache = CacheService.getScriptCache();
    // Remove known cache keys — placement x audience combos
    var placements = ['Auth Screen', 'Home Widget', 'Sidebar', 'Hub Section', 'Inline Quote', 'all'];
    var audiences = ['All', 'Members', 'Stewards', 'all'];
    var types = ['Quote', 'Tip', 'Concept', 'Mini-Lesson', 'Manifesto', 'Negotiation Set', 'all'];
    var keys = [];
    for (var p = 0; p < placements.length; p++) {
      for (var a = 0; a < audiences.length; a++) {
        for (var t = 0; t < types.length; t++) {
          keys.push('ke_' + placements[p] + '_' + audiences[a] + '_' + types[t]);
        }
      }
    }
    cache.removeAll(keys);
    cache.remove('ke_all_steward');
  } catch (_) { /* non-critical */ }
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
// getNotificationRecipientList removed v4.29.0 — dead code, superseded by getNotificationRecipientListFull (auth-gated)

/**
 * Get full member list with directory columns for filtering/sorting.
 * Used by steward notification compose form recipient picker.
 * Returns: name, email, location, department, jobTitle for each member.
 * @returns {Object[]}\n */
function getNotificationRecipientListFull(sessionToken) {
  try {
    // CR-AUTH: Steward-only — member email enumeration is a privacy risk (v4.29.0)
    var auth = checkWebAppAuthorization('steward', sessionToken);
    if (!auth.isAuthorized) return [];

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
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }

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
    .setFontColor(SHEET_COLORS.BG_WHITE)
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
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
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

// ═══════════════════════════════════════
// GRIEVANCE FORM — PDF GENERATION & INITIATION  (v4.32.0)
// ═══════════════════════════════════════

/**
 * Generates a populated grievance form PDF from member/grievance data.
 * Reads src/grievance_form.html template, injects field values, converts to PDF blob.
 * @param {Object} formData - All form fields (grievant, jobtitle, region, articles, etc.)
 * @returns {Blob|null} PDF blob or null on error
 * @private
 */
function generateGrievancePDF_(formData) {
  try {
    var html = HtmlService.createHtmlOutputFromFile('grievance_form').getContent();

    // Build value injection map — each key is an element ID in the form
    var valueMap = {
      'grievant':   formData.grievant   || '',
      'jobtitle':   formData.jobtitle   || '',
      'startdate':  formData.startdate  || '',
      'agency':     formData.agency     || '',
      'region':     formData.region     || '',
      'workloc':    formData.workloc    || '',
      'articles':   formData.articles   || '',
      'step':       formData.step       || '1',
      'statement':  formData.statement  || '',
      'statement2': formData.statement2 || '',
      'remedy1':    formData.remedy1    || '',
      'remedy2':    formData.remedy2    || ''
    };

    // Inject values into input/textarea/select fields by ID
    for (var id in valueMap) {
      if (!valueMap.hasOwnProperty(id)) continue;
      var val = escapeHtml(String(valueMap[id]));

      // Handle <input id="X" ... value="Y"> — inject/replace value attribute
      var inputRx = new RegExp('(<(?:input|INPUT)[^>]*id=["\']' + id + '["\'][^>]*?)(/?>)', 'g');
      html = html.replace(inputRx, function(match, before, close) {
        // Remove any existing value attribute
        before = before.replace(/\s+value="[^"]*"/i, '');
        return before + ' value="' + val + '"' + close;
      });

      // Handle <textarea id="X">...</textarea> — inject content
      var taRx = new RegExp('(<textarea[^>]*id=["\']' + id + '["\'][^>]*>)(</textarea>)', 'gi');
      html = html.replace(taRx, '$1' + val + '$2');

      // Handle <select id="X"> — mark matching <option> as selected
      if (val) {
        var selRx = new RegExp('(<select[^>]*id=["\']' + id + '["\'][^>]*>[\\s\\S]*?</select>)', 'gi');
        html = html.replace(selRx, function(selectBlock) {
          // Remove any existing selected attributes
          selectBlock = selectBlock.replace(/ selected/gi, '');
          // Add selected to matching option
          var optRx = new RegExp('(<option[^>]*value=["\']' + val.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '["\'])');
          return selectBlock.replace(optRx, '$1 selected');
        });
      }
    }

    // Handle managers selects (special — value is name text, not a simple match)
    if (formData.managers) {
      var mgrVal = escapeHtml(String(formData.managers));
      var mgrRx = /(<select[^>]*id=['"]managers['"][^>]*>[\s\S]*?<\/select>)/gi;
      html = html.replace(mgrRx, function(block) {
        block = block.replace(/ selected/gi, '');
        var optRx = new RegExp('(<option[^>]*>)(' + mgrVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')(</option>)');
        return block.replace(optRx, function(_m, p1, p2, p3) { return p1 + ' selected' + p2 + p3; });
      });
    }

    // Convert HTML to PDF blob via Drive
    var blob = Utilities.newBlob(html, 'text/html', 'grievance_form.html');
    var pdfBlob = blob.getAs('application/pdf');
    var grievantName = formData.grievant || 'Unknown';
    pdfBlob.setName('Grievance Form - ' + grievantName + '.pdf');

    return pdfBlob;
  } catch (e) {
    Logger.log('generateGrievancePDF_ error: ' + e.message + '\n' + (e.stack || ''));
    return null;
  }
}

// ============================================================================
// E-SIGNATURE SYSTEM  (v4.33.0)
// ============================================================================

/**
 * Generates a SHA-256 document hash for a grievance to bind the signature to the content.
 * The hash covers all substantive grievance fields so any change after signing is detectable.
 * @param {Object} grievanceRow - Key/value object of grievance data
 * @returns {string} Hex-encoded SHA-256 hash
 */
function generateDocumentHash_(grievanceRow) {
  var fields = [
    String(grievanceRow.grievanceId  || ''),
    String(grievanceRow.memberId     || ''),
    String(grievanceRow.firstName    || ''),
    String(grievanceRow.lastName     || ''),
    String(grievanceRow.articles     || ''),
    String(grievanceRow.issueCategory|| ''),
    String(grievanceRow.description  || ''),
    String(grievanceRow.remedy       || ''),
    String(grievanceRow.step         || ''),
    String(grievanceRow.incidentDate || ''),
    String(grievanceRow.dateFiled    || '')
  ];
  var payload = fields.join('|');
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload, Utilities.Charset.UTF_8);
  return rawHash.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2); }).join('');
}

/**
 * Generates a one-time signature token for a grievance.
 * Stored in the Grievance Log row and used in the signing URL.
 * @param {string} grievanceId
 * @returns {string} UUID-based token
 */
function generateSignatureToken_(grievanceId) {
  // Use two full UUIDs concatenated for 256-bit entropy — prevents brute-force enumeration
  return 'SIG-' + Utilities.getUuid() + '-' + Utilities.getUuid();
}

/**
 * Generates a QR code image URL via Google Charts API.
 * The QR encodes the grievance ID + document hash for verification.
 * @param {string} grievanceId
 * @param {string} documentHash
 * @returns {string} URL to QR code image
 */
function getQRCodeUrl_(grievanceId, documentHash) {
  var qrData = 'GRIEVANCE:' + grievanceId + '|HASH:' + documentHash.substring(0, 16);
  return 'https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=' + encodeURIComponent(qrData);
}

/**
 * Generates a draft grievance PDF with a QR code footer.
 * The QR contains the grievance ID + document hash for tamper detection.
 * @param {Object} formData - Same shape as generateGrievancePDF_
 * @param {string} grievanceId
 * @param {string} documentHash
 * @returns {Blob|null} PDF blob
 */
function generateDraftGrievancePDF_(formData, grievanceId, documentHash) {
  try {
    var html = HtmlService.createHtmlOutputFromFile('grievance_form').getContent();

    // Build value injection map
    var valueMap = {
      'grievant':   formData.grievant   || '',
      'jobtitle':   formData.jobtitle   || '',
      'startdate':  formData.startdate  || '',
      'agency':     formData.agency     || '',
      'region':     formData.region     || '',
      'workloc':    formData.workloc    || '',
      'articles':   formData.articles   || '',
      'step':       formData.step       || '1',
      'statement':  formData.statement  || '',
      'statement2': formData.statement2 || '',
      'remedy1':    formData.remedy1    || '',
      'remedy2':    formData.remedy2    || ''
    };

    // Inject values into form fields
    for (var id in valueMap) {
      if (!valueMap.hasOwnProperty(id)) continue;
      var val = escapeHtml(String(valueMap[id]));
      var inputRx = new RegExp('(<(?:input|INPUT)[^>]*id=["\']' + id + '["\'][^>]*?)(/?>)', 'g');
      html = html.replace(inputRx, function(match, before, close) {
        before = before.replace(/\s+value="[^"]*"/i, '');
        return before + ' value="' + val + '"' + close;
      });
      var taRx = new RegExp('(<textarea[^>]*id=["\']' + id + '["\'][^>]*>)(</textarea>)', 'gi');
      html = html.replace(taRx, '$1' + val + '$2');
      if (val) {
        var selRx = new RegExp('(<select[^>]*id=["\']' + id + '["\'][^>]*>[\\s\\S]*?</select>)', 'gi');
        html = html.replace(selRx, function(selectBlock) {
          selectBlock = selectBlock.replace(/ selected/gi, '');
          var optRx = new RegExp('(<option[^>]*value=["\']' + val.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '["\'])');
          return selectBlock.replace(optRx, '$1 selected');
        });
      }
    }

    // Handle managers selects
    if (formData.managers) {
      var mgrVal = escapeHtml(String(formData.managers));
      var mgrRx = /(<select[^>]*id=['"]managers['"][^>]*>[\s\S]*?<\/select>)/gi;
      html = html.replace(mgrRx, function(block) {
        block = block.replace(/ selected/gi, '');
        var optRx = new RegExp('(<option[^>]*>)(' + mgrVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')(</option>)');
        return block.replace(optRx, function(_m, p1, p2, p3) { return p1 + ' selected' + p2 + p3; });
      });
    }

    // ── Inject QR code + document reference footer ──
    var qrUrl = getQRCodeUrl_(grievanceId, documentHash);
    var shortHash = documentHash.substring(0, 16);
    var footerHtml = '<div style="margin-top:20px;padding-top:12px;border-top:1px solid #ccc;display:flex;align-items:flex-end;justify-content:space-between;font-size:8pt;color:#888;">' +
      '<div style="flex:1;"><span style="font-size:7pt;letter-spacing:.5px;">DRAFT — AWAITING MEMBER SIGNATURE</span><br>' +
      '<span style="font-family:Courier New,monospace;font-size:7pt;">Ref: ' + escapeHtml(grievanceId) + ' | ' + shortHash + '</span></div>' +
      '<img src="' + escapeHtml(qrUrl) + '" width="80" height="80" style="margin-left:12px;" alt="QR">' +
      '</div>';

    // Insert footer before closing </div><!-- /.page --> or </body>
    html = html.replace(/<\/div>\s*<div id="toast">/, footerHtml + '</div><div id="toast">');

    var blob = Utilities.newBlob(html, 'text/html', 'grievance_draft.html');
    var pdfBlob = blob.getAs('application/pdf');
    return pdfBlob;
  } catch (e) {
    if (typeof secureLog === 'function') secureLog('esign', 'generateDraftGrievancePDF_ error', { error: e.message });
    else Logger.log('generateDraftGrievancePDF_ error: ' + e.message);
    return null;
  }
}

/**
 * Generates the final signed grievance PDF with embedded signature image + QR code.
 * The signature is rendered as part of the HTML so it cannot be extracted/reused.
 * @param {Object} formData - Form field values
 * @param {string} grievanceId
 * @param {string} documentHash
 * @param {string} sigBase64 - Base64-encoded PNG of the drawn signature
 * @param {string} memberName - Printed name for the signature block
 * @param {string} signedDate - ISO date string
 * @returns {Blob|null} PDF blob
 */
function generateSignedGrievancePDF_(formData, grievanceId, documentHash, sigBase64, memberName, signedDate) {
  try {
    var html = HtmlService.createHtmlOutputFromFile('grievance_form').getContent();

    // Inject form values (same as draft)
    var valueMap = {
      'grievant':   formData.grievant   || '',
      'jobtitle':   formData.jobtitle   || '',
      'startdate':  formData.startdate  || '',
      'agency':     formData.agency     || '',
      'region':     formData.region     || '',
      'workloc':    formData.workloc    || '',
      'articles':   formData.articles   || '',
      'step':       formData.step       || '1',
      'statement':  formData.statement  || '',
      'statement2': formData.statement2 || '',
      'remedy1':    formData.remedy1    || '',
      'remedy2':    formData.remedy2    || ''
    };

    for (var id in valueMap) {
      if (!valueMap.hasOwnProperty(id)) continue;
      var val = escapeHtml(String(valueMap[id]));
      var inputRx = new RegExp('(<(?:input|INPUT)[^>]*id=["\']' + id + '["\'][^>]*?)(/?>)', 'g');
      html = html.replace(inputRx, function(match, before, close) {
        before = before.replace(/\s+value="[^"]*"/i, '');
        return before + ' value="' + val + '"' + close;
      });
      var taRx = new RegExp('(<textarea[^>]*id=["\']' + id + '["\'][^>]*>)(</textarea>)', 'gi');
      html = html.replace(taRx, '$1' + val + '$2');
      if (val) {
        var selRx = new RegExp('(<select[^>]*id=["\']' + id + '["\'][^>]*>[\\s\\S]*?</select>)', 'gi');
        html = html.replace(selRx, function(selectBlock) {
          selectBlock = selectBlock.replace(/ selected/gi, '');
          var optRx = new RegExp('(<option[^>]*value=["\']' + val.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '["\'])');
          return selectBlock.replace(optRx, '$1 selected');
        });
      }
    }

    if (formData.managers) {
      var mgrVal = escapeHtml(String(formData.managers));
      var mgrRx = /(<select[^>]*id=['"]managers['"][^>]*>[\s\S]*?<\/select>)/gi;
      html = html.replace(mgrRx, function(block) {
        block = block.replace(/ selected/gi, '');
        var optRx = new RegExp('(<option[^>]*>)(' + mgrVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')(</option>)');
        return block.replace(optRx, function(_m, p1, p2, p3) { return p1 + ' selected' + p2 + p3; });
      });
    }

    // ── Replace the first signature box with the drawn signature ──
    // Find the first sig-box div and replace with the signature image bound to this document
    var sigBlockHtml =
      '<div style="position:relative;border-bottom:2px solid #000;min-height:60px;padding:4px;">' +
      '<img src="data:image/png;base64,' + sigBase64 + '" style="max-height:55px;max-width:100%;" alt="Signature">' +
      '</div>';
    // Replace first grievant signature box
    html = html.replace(/<div class="sig-box"><\/div>(\s*<span class="sig-lbl">Grievant Signature\(s\)<\/span>)/, sigBlockHtml + '$1');

    // ── Fill in the grievant date field for the first sig block ──
    var dateRx = /(<input[^>]*id=['"]sig1gdate['"][^>]*?)(\/?>\s*)/;
    html = html.replace(dateRx, function(match, before, close) {
      before = before.replace(/\s+value="[^"]*"/i, '');
      return before + ' value="' + escapeHtml(signedDate) + '"' + close;
    });

    // ── Inject signed footer with QR code ──
    var qrUrl = getQRCodeUrl_(grievanceId, documentHash);
    var shortHash = documentHash.substring(0, 16);
    var footerHtml =
      '<div style="margin-top:20px;padding-top:12px;border-top:1px solid #ccc;display:flex;align-items:flex-end;justify-content:space-between;font-size:8pt;color:#888;">' +
      '<div style="flex:1;">' +
      '<span style="font-size:8pt;font-weight:bold;color:#059669;">&#10003; SIGNED</span><br>' +
      '<span style="font-size:7pt;">Signed by ' + escapeHtml(memberName) + ' on ' + escapeHtml(signedDate) + '</span><br>' +
      '<span style="font-family:Courier New,monospace;font-size:7pt;">Ref: ' + escapeHtml(grievanceId) + ' | Hash: ' + shortHash + '</span>' +
      '</div>' +
      '<img src="' + escapeHtml(qrUrl) + '" width="80" height="80" style="margin-left:12px;" alt="QR">' +
      '</div>';

    html = html.replace(/<\/div>\s*<div id="toast">/, footerHtml + '</div><div id="toast">');

    var blob = Utilities.newBlob(html, 'text/html', 'grievance_signed.html');
    var pdfBlob = blob.getAs('application/pdf');
    return pdfBlob;
  } catch (e) {
    if (typeof secureLog === 'function') secureLog('esign', 'generateSignedGrievancePDF_ error', { error: e.message });
    else Logger.log('generateSignedGrievancePDF_ error: ' + e.message);
    return null;
  }
}

/**
 * Builds a standardized PDF filename for a grievance.
 * Pattern: {GrievanceID}_{LastName}-{FirstName}_Grievance_{Status}_{DateFiled}.pdf
 * @param {string} grievanceId
 * @param {string} firstName
 * @param {string} lastName
 * @param {string} status - 'Draft' or 'Signed'
 * @param {Date|string} dateFiled
 * @returns {string}
 */
function buildGrievancePdfName_(grievanceId, firstName, lastName, status, dateFiled) {
  var dateStr = '';
  try {
    dateStr = Utilities.formatDate(new Date(dateFiled), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (_e) {
    dateStr = 'unknown';
  }
  var safeLast  = (lastName  || 'Unknown').replace(/[^a-zA-Z0-9]/g, '');
  var safeFirst = (firstName || ''       ).replace(/[^a-zA-Z0-9]/g, '');
  var namePart = safeLast + (safeFirst ? '-' + safeFirst : '');
  return grievanceId + '_' + namePart + '_Grievance_' + status + '_' + dateStr + '.pdf';
}

/**
 * Retrieves grievance data for the e-signature page.
 * Validates the signature token and returns a summary of the grievance.
 * @param {string} sigToken - The unique signature token from the signing URL
 * @returns {Object} { success, grievanceId, documentHash, articles, ... } or { success: false, message }
 */
function getGrievanceForSigning(sigToken) {
  try {
    if (!sigToken) return { success: false, message: 'Invalid signature link.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Service temporarily unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Service temporarily unavailable.' };

    var data = sheet.getDataRange().getValues();
    var C = GRIEVANCE_COLS;
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      var token = String(data[i][C.SIGNATURE_TOKEN - 1] || '').trim();
      if (token === sigToken) {
        targetRow = i;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Invalid or expired signature link.' };

    var row = data[targetRow];
    var sigStatus = String(row[C.SIGNATURE_STATUS - 1] || '').trim();
    var grievanceId = String(row[C.GRIEVANCE_ID - 1] || '');

    // Already signed?
    if (sigStatus === 'Signed') {
      return {
        success: true,
        alreadySigned: true,
        grievanceId: grievanceId,
        signedDate: String(row[C.SIGNED_DATE - 1] || '')
      };
    }

    // Build grievance data object for hash + display
    var grievanceObj = {
      grievanceId:   grievanceId,
      memberId:      String(row[C.MEMBER_ID - 1]       || ''),
      firstName:     String(row[C.FIRST_NAME - 1]      || ''),
      lastName:      String(row[C.LAST_NAME - 1]       || ''),
      articles:      String(row[C.ARTICLES - 1]        || ''),
      issueCategory: String(row[C.ISSUE_CATEGORY - 1]  || ''),
      description:   String(row[C.RESOLUTION - 1]      || ''),
      remedy:        '',  // stored in resolution field combined
      step:          String(row[C.CURRENT_STEP - 1]     || ''),
      incidentDate:  row[C.INCIDENT_DATE - 1] ? String(row[C.INCIDENT_DATE - 1]) : '',
      dateFiled:     row[C.DATE_FILED - 1] ? String(row[C.DATE_FILED - 1]) : ''
    };

    var documentHash = generateDocumentHash_(grievanceObj);

    // Resolve steward name
    var stewardEmail = String(row[C.STEWARD - 1] || '');
    var stewardName = stewardEmail.split('@')[0];
    try {
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberDir && memberDir.getLastRow() > 1) {
        var mData = memberDir.getDataRange().getValues();
        for (var j = 1; j < mData.length; j++) {
          var sEmail = String(mData[j][MEMBER_COLS.EMAIL - 1] || '').toLowerCase().trim();
          if (sEmail === stewardEmail.toLowerCase().trim()) {
            var sFirst = String(mData[j][MEMBER_COLS.FIRST_NAME - 1] || '').trim();
            var sLast  = String(mData[j][MEMBER_COLS.LAST_NAME - 1]  || '').trim();
            if (sFirst || sLast) stewardName = (sFirst + ' ' + sLast).trim();
            break;
          }
        }
      }
    } catch (_e) { /* fallback to email prefix */ }

    var dateFiledStr = '';
    try {
      if (row[C.DATE_FILED - 1]) {
        dateFiledStr = Utilities.formatDate(new Date(row[C.DATE_FILED - 1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    } catch (_e) { dateFiledStr = String(row[C.DATE_FILED - 1] || ''); }

    return {
      success: true,
      alreadySigned: false,
      grievanceId: grievanceId,
      memberName: grievanceObj.firstName + ' ' + grievanceObj.lastName,
      stewardName: stewardName,
      dateFiled: dateFiledStr,
      articles: grievanceObj.articles,
      issueCategory: grievanceObj.issueCategory,
      description: grievanceObj.description,
      remedy: grievanceObj.remedy,
      step: grievanceObj.step,
      documentHash: documentHash
    };
  } catch (e) {
    if (typeof secureLog === 'function') secureLog('esign', 'getGrievanceForSigning error', { error: e.message });
    else Logger.log('getGrievanceForSigning error: ' + e.message);
    return { success: false, message: 'Error loading grievance.' };
  }
}

/**
 * Processes a submitted signature for a grievance.
 * - Validates signature token
 * - Generates document hash and verifies integrity
 * - Generates signed PDF with embedded signature + QR code
 * - Saves signed PDF to Drive (replaces draft)
 * - Updates Grievance Log with signature metadata
 * - Notifies member: "Grievance Filed" + canned message
 * - Notifies steward: signature completed
 * @param {string} sigToken - Signature token
 * @param {string} sigBase64 - Base64 PNG of drawn signature
 * @returns {Object} { success, message, driveFolderUrl }
 */
function submitGrievanceSignature(sigToken, sigBase64) {
  try {
    if (!sigToken || !sigBase64) return { success: false, message: 'Missing signature data.' };

    // Validate sigBase64: must be pure base64 characters, max 500KB
    if (sigBase64.length > 500000) return { success: false, message: 'Signature data too large.' };
    if (!/^[A-Za-z0-9+/=]+$/.test(sigBase64)) return { success: false, message: 'Invalid signature data format.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };
    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Grievance Log not found.' };

    var data = sheet.getDataRange().getValues();
    var C = GRIEVANCE_COLS;
    var targetRow = -1;

    for (var i = 1; i < data.length; i++) {
      var token = String(data[i][C.SIGNATURE_TOKEN - 1] || '').trim();
      if (token === sigToken) {
        targetRow = i;
        break;
      }
    }

    if (targetRow === -1) return { success: false, message: 'Invalid or expired signature link.' };

    var row = data[targetRow];
    var sheetRow = targetRow + 1; // 1-indexed sheet row

    // Already signed guard
    if (String(row[C.SIGNATURE_STATUS - 1] || '').trim() === 'Signed') {
      return { success: false, message: 'This grievance has already been signed.' };
    }

    var grievanceId = String(row[C.GRIEVANCE_ID - 1] || '');
    var firstName   = String(row[C.FIRST_NAME - 1]   || '');
    var lastName    = String(row[C.LAST_NAME - 1]     || '');
    var memberEmail = String(row[C.MEMBER_EMAIL - 1]  || '').trim().toLowerCase();
    var stewardEmail = String(row[C.STEWARD - 1]      || '').trim();
    var dateFiled   = row[C.DATE_FILED - 1] || new Date();
    var driveFolderUrl = String(row[C.DRIVE_FOLDER_URL - 1] || '');
    var driveFolderId  = String(row[C.DRIVE_FOLDER_ID - 1]  || '');

    // Build grievance data for hash
    var grievanceObj = {
      grievanceId:   grievanceId,
      memberId:      String(row[C.MEMBER_ID - 1]       || ''),
      firstName:     firstName,
      lastName:      lastName,
      articles:      String(row[C.ARTICLES - 1]        || ''),
      issueCategory: String(row[C.ISSUE_CATEGORY - 1]  || ''),
      description:   String(row[C.RESOLUTION - 1]      || ''),
      remedy:        '',
      step:          String(row[C.CURRENT_STEP - 1]     || ''),
      incidentDate:  row[C.INCIDENT_DATE - 1] ? String(row[C.INCIDENT_DATE - 1]) : '',
      dateFiled:     row[C.DATE_FILED - 1] ? String(row[C.DATE_FILED - 1]) : ''
    };

    var documentHash = generateDocumentHash_(grievanceObj);
    var tz = Session.getScriptTimeZone();
    var signedDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var memberName = (firstName + ' ' + lastName).trim();

    // ── Generate signed PDF ──
    var hireDateStr = '';
    // Look up hire date from member directory
    try {
      var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
      if (memberDir && memberDir.getLastRow() > 1) {
        var mAll = memberDir.getDataRange().getValues();
        for (var m = 1; m < mAll.length; m++) {
          if (String(mAll[m][MEMBER_COLS.EMAIL - 1] || '').toLowerCase().trim() === memberEmail) {
            var hd = mAll[m][MEMBER_COLS.HIRE_DATE - 1];
            if (hd) {
              try { hireDateStr = Utilities.formatDate(new Date(hd), tz, 'yyyy-MM-dd'); } catch (_e) { hireDateStr = String(hd); }
            }
            break;
          }
        }
      }
    } catch (_e) { /* skip */ }

    var regionVal = '';
    var workLoc = String(row[C.LOCATION - 1] || '');
    if (workLoc) {
      var locLower = workLoc.toLowerCase();
      if (locLower.indexOf('everett') !== -1) regionVal = 'Everett';
      else if (locLower.indexOf('worcester') !== -1) regionVal = 'Worcester';
    }

    var pdfData = {
      grievant:   memberName,
      jobtitle:   '', // from member dir if available
      startdate:  hireDateStr,
      agency:     '',
      region:     regionVal,
      workloc:    workLoc,
      articles:   grievanceObj.articles,
      step:       grievanceObj.step,
      statement:  grievanceObj.description,
      statement2: '',
      remedy1:    grievanceObj.remedy,
      remedy2:    '',
      managers:   ''
    };

    var signedPdfBlob = generateSignedGrievancePDF_(pdfData, grievanceId, documentHash, sigBase64, memberName, signedDate);

    // ── Save signed PDF to Drive ──
    if (signedPdfBlob && driveFolderId) {
      try {
        var signedFileName = buildGrievancePdfName_(grievanceId, firstName, lastName, 'Signed', dateFiled);
        signedPdfBlob.setName(signedFileName);
        var caseFolder = DriveApp.getFolderById(driveFolderId);

        // Remove all draft PDFs (handles retries that may have created duplicates)
        var draftPrefix = grievanceId + '_';
        var files = caseFolder.getFiles();
        while (files.hasNext()) {
          var f = files.next();
          if (f.getName().indexOf(draftPrefix) === 0 && f.getName().indexOf('_Draft_') !== -1) {
            f.setTrashed(true);
          }
        }

        caseFolder.createFile(signedPdfBlob);
      } catch (driveErr) {
        if (typeof secureLog === 'function') secureLog('esign', 'Drive save error', { error: driveErr.message });
        else Logger.log('submitGrievanceSignature: Drive save error: ' + driveErr.message);
      }
    }

    // ── Update Grievance Log row ──
    if (C.SIGNATURE_STATUS) sheet.getRange(sheetRow, C.SIGNATURE_STATUS).setValue('Signed');
    if (C.SIGNATURE_HASH)   sheet.getRange(sheetRow, C.SIGNATURE_HASH).setValue(documentHash);
    if (C.SIGNED_DATE)      sheet.getRange(sheetRow, C.SIGNED_DATE).setValue(new Date());
    if (C.LAST_UPDATED)     sheet.getRange(sheetRow, C.LAST_UPDATED).setValue(new Date());

    // Signature image is embedded in the signed PDF (authoritative copy in Drive).
    // No separate storage of raw signature data — avoids PII persistence in Script Properties.

    // ── Clear signature token (prevent reuse) ──
    if (C.SIGNATURE_TOKEN) sheet.getRange(sheetRow, C.SIGNATURE_TOKEN).setValue('');

    // ── Notify member: Grievance Filed ──
    if (typeof pushNotification === 'function' && memberEmail) {
      pushNotification(memberEmail, {
        title: 'Grievance Filed (' + grievanceId + ')',
        body: 'Your grievance (' + grievanceId + ') has been signed and officially filed. ' +
              'Your steward will keep you updated on the progress. ' +
              (driveFolderUrl ? 'Your case folder: ' + driveFolderUrl : ''),
        type: 'Grievance Filed'
      });
    }

    // ── Notify steward: Signature received ──
    if (typeof pushNotification === 'function' && stewardEmail) {
      pushNotification(stewardEmail, {
        title: 'Signature Received (' + grievanceId + ')',
        body: memberName + ' has signed grievance ' + grievanceId + '. The signed PDF is ready in the case folder. You may now email the document.',
        type: 'Signature Received'
      });
    }

    // ── Refresh badges ──
    if (typeof _refreshNavBadges === 'function') _refreshNavBadges();

    // ── Audit log ──
    if (typeof logAuditEvent === 'function') {
      logAuditEvent(AUDIT_EVENTS.GRIEVANCE_SIGNED || 'GRIEVANCE_SIGNED', {
        grievanceId: grievanceId,
        memberEmail: memberEmail,
        documentHash: documentHash,
        signedDate: signedDate
      });
    }

    return {
      success: true,
      message: 'Grievance signed successfully.',
      driveFolderUrl: driveFolderUrl
    };
  } catch (e) {
    if (typeof secureLog === 'function') secureLog('esign', 'submitGrievanceSignature error', { error: e.message });
    else Logger.log('submitGrievanceSignature error: ' + e.message);
    return { success: false, message: 'Error processing signature.' };
  }
}

/**
 * Initiates a grievance on behalf of a member. Steward-only.
 * - Creates Grievance Log entry
 * - Generates populated PDF form
 * - Saves PDF to member's grievance Drive folder
 * - Notifies member on webapp
 *
 * @param {string} stewardEmail - Authenticated steward email
 * @param {Object} data - { memberEmail, step, articles, issueCategory, description, remedy,
 *                          incidentDate, managers, formOverrides }
 * @param {string} [idemKey] - Idempotency key to prevent duplicate submissions
 * @returns {Object} { success, grievanceId, driveFolderUrl, message }
 */
function initiateGrievance(stewardEmail, data, idemKey) {
  try {
    // Idempotency guard
    if (idemKey) {
      var idemCache = CacheService.getScriptCache();
      if (idemCache.get('IDEM_' + idemKey)) return { duplicate: true, message: 'Duplicate request ignored.' };
      idemCache.put('IDEM_' + idemKey, '1', 300);
    }

    if (!stewardEmail || !data || !data.memberEmail) {
      return { success: false, message: 'Steward email and member email are required.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: 'Spreadsheet unavailable.' };

    // ── 1. Look up member in Member Directory ────────────────────────────────
    var memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!memberDir || memberDir.getLastRow() < 2) {
      return { success: false, message: 'Member Directory not found or empty.' };
    }

    var memberData = null;
    var allMembers = memberDir.getDataRange().getValues();
    var emailLower = data.memberEmail.toLowerCase().trim();
    for (var i = 1; i < allMembers.length; i++) {
      var rowEmail = String(allMembers[i][MEMBER_COLS.EMAIL - 1] || '').toLowerCase().trim();
      if (rowEmail === emailLower) {
        memberData = {
          memberId:     String(allMembers[i][MEMBER_COLS.MEMBER_ID - 1]    || '').trim(),
          firstName:    String(allMembers[i][MEMBER_COLS.FIRST_NAME - 1]   || '').trim(),
          lastName:     String(allMembers[i][MEMBER_COLS.LAST_NAME - 1]    || '').trim(),
          email:        emailLower,
          phone:        String(allMembers[i][MEMBER_COLS.PHONE - 1]        || '').trim(),
          jobTitle:     String(allMembers[i][MEMBER_COLS.JOB_TITLE - 1]    || '').trim(),
          workLocation: String(allMembers[i][MEMBER_COLS.WORK_LOCATION - 1]|| '').trim(),
          unit:         String(allMembers[i][MEMBER_COLS.UNIT - 1]         || '').trim(),
          supervisor:   String(allMembers[i][MEMBER_COLS.SUPERVISOR - 1]   || '').trim(),
          manager:      String(allMembers[i][MEMBER_COLS.MANAGER - 1]      || '').trim(),
          hireDate:     allMembers[i][MEMBER_COLS.HIRE_DATE - 1] || '',
          streetAddress:String(allMembers[i][MEMBER_COLS.STREET_ADDRESS - 1]|| '').trim(),
          city:         String(allMembers[i][MEMBER_COLS.CITY - 1]         || '').trim(),
          state:        String(allMembers[i][MEMBER_COLS.STATE - 1]        || '').trim(),
          zipCode:      String(allMembers[i][MEMBER_COLS.ZIP_CODE - 1]     || '').trim()
        };
        break;
      }
    }

    if (!memberData) {
      return { success: false, message: 'Member not found in directory.' };
    }

    // ── 2. Create Grievance Log entry ───────────────────────────────────────
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!grievanceSheet) return { success: false, message: 'Grievance Log sheet not found.' };

    var grievanceId = getNextGrievanceId(grievanceSheet);
    var filingDate = new Date();
    var deadlines = calculateInitialDeadlines(filingDate);

    var totalCols = getGrievanceHeaders().length;
    var rowData = new Array(totalCols).fill('');

    rowData[GRIEVANCE_COLS.GRIEVANCE_ID - 1]   = grievanceId;
    rowData[GRIEVANCE_COLS.MEMBER_ID - 1]       = escapeForFormula(memberData.memberId);
    rowData[GRIEVANCE_COLS.FIRST_NAME - 1]      = escapeForFormula(memberData.firstName);
    if (GRIEVANCE_COLS.LAST_NAME) {
      rowData[GRIEVANCE_COLS.LAST_NAME - 1]     = escapeForFormula(memberData.lastName);
    }
    rowData[GRIEVANCE_COLS.STATUS - 1]          = GRIEVANCE_STATUS.OPEN;
    rowData[GRIEVANCE_COLS.CURRENT_STEP - 1]    = data.step || 1;
    if (data.incidentDate) {
      rowData[GRIEVANCE_COLS.INCIDENT_DATE - 1] = new Date(data.incidentDate);
    }
    rowData[GRIEVANCE_COLS.DATE_FILED - 1]      = filingDate;
    rowData[GRIEVANCE_COLS.STEP1_DUE - 1]       = deadlines.step1Due;
    rowData[GRIEVANCE_COLS.ARTICLES - 1]        = escapeForFormula(data.articles || '');
    rowData[GRIEVANCE_COLS.ISSUE_CATEGORY - 1]  = escapeForFormula(data.issueCategory || '');
    rowData[GRIEVANCE_COLS.MEMBER_EMAIL - 1]    = escapeForFormula(memberData.email);
    rowData[GRIEVANCE_COLS.LOCATION - 1]        = escapeForFormula(memberData.workLocation);
    rowData[GRIEVANCE_COLS.STEWARD - 1]         = escapeForFormula(stewardEmail);
    rowData[GRIEVANCE_COLS.RESOLUTION - 1]      = escapeForFormula(data.description || '');
    rowData[GRIEVANCE_COLS.LAST_UPDATED - 1]    = new Date();
    if (GRIEVANCE_COLS.ACTION_TYPE) {
      rowData[GRIEVANCE_COLS.ACTION_TYPE - 1]   = 'Grievance';
    }

    // ── Signature tracking fields ──
    var sigToken = generateSignatureToken_(grievanceId);
    if (GRIEVANCE_COLS.SIGNATURE_STATUS) rowData[GRIEVANCE_COLS.SIGNATURE_STATUS - 1] = 'Pending';
    if (GRIEVANCE_COLS.SIGNATURE_TOKEN)  rowData[GRIEVANCE_COLS.SIGNATURE_TOKEN - 1]  = sigToken;

    grievanceSheet.appendRow(rowData);

    // Programmatic writes don't trigger onEdit, so bidirectional sync
    // (syncDropdownToConfig_) never fires. Explicitly sync dropdown values
    // to Config so they appear in future grievance form dropdowns.
    try {
      var _grDDFields = [
        { val: data.issueCategory,       col: CONFIG_COLS.ISSUE_CATEGORY },
        { val: data.step,                col: CONFIG_COLS.GRIEVANCE_STEP },
        { val: memberData.workLocation,  col: CONFIG_COLS.OFFICE_LOCATIONS }
      ];
      // Articles may be comma-separated (multi-select) — sync each individually
      if (data.articles) {
        var _artParts = String(data.articles).split(',');
        for (var _ap = 0; _ap < _artParts.length; _ap++) {
          var _artVal = _artParts[_ap].trim();
          if (_artVal && CONFIG_COLS.ARTICLES) addToConfigDropdown_(CONFIG_COLS.ARTICLES, _artVal);
        }
      }
      for (var _gd = 0; _gd < _grDDFields.length; _gd++) {
        var _gv = String(_grDDFields[_gd].val || '').trim();
        if (_gv && _grDDFields[_gd].col) addToConfigDropdown_(_grDDFields[_gd].col, _gv);
      }
    } catch (_grSyncErr) {
      Logger.log('initiateGrievance config sync: ' + _grSyncErr.message);
    }

    logAuditEvent(AUDIT_EVENTS.GRIEVANCE_CREATED, {
      grievanceId: grievanceId,
      memberId: memberData.memberId,
      memberEmail: memberData.email,
      createdBy: stewardEmail
    });

    // ── 3. Setup Drive folder ───────────────────────────────────────────────
    var driveResult = setupDriveFolderForGrievance(grievanceId);
    var driveFolderUrl = (driveResult && driveResult.success) ? driveResult.folderUrl : '';

    // ── 4. Generate and save grievance PDF ──────────────────────────────────
    var hireDateStr = '';
    if (memberData.hireDate) {
      try {
        hireDateStr = Utilities.formatDate(new Date(memberData.hireDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } catch (_e) {
        hireDateStr = String(memberData.hireDate);
      }
    }

    // Determine region from work location
    var regionVal = '';
    if (memberData.workLocation) {
      var locLower = memberData.workLocation.toLowerCase();
      if (locLower.indexOf('everett') !== -1) regionVal = 'Everett';
      else if (locLower.indexOf('worcester') !== -1) regionVal = 'Worcester';
    }

    var formOverrides = data.formOverrides || {};
    var pdfData = {
      grievant:   memberData.firstName + ' ' + memberData.lastName,
      jobtitle:   formOverrides.jobtitle || memberData.jobTitle,
      startdate:  formOverrides.startdate || hireDateStr,
      agency:     formOverrides.agency || '',
      region:     formOverrides.region || regionVal,
      workloc:    formOverrides.workloc || memberData.workLocation,
      managers:   formOverrides.managers || memberData.manager || memberData.supervisor,
      managers2:  formOverrides.managers2 || '',
      articles:   formOverrides.articles || data.articles,
      step:       formOverrides.step || String(data.step || '1'),
      statement:  formOverrides.statement || data.description,
      statement2: formOverrides.statement2 || '',
      remedy1:    formOverrides.remedy1 || data.remedy,
      remedy2:    formOverrides.remedy2 || ''
    };

    // Generate document hash for draft PDF
    // Note: remedy is not stored in a separate sheet column, so it must be '' here
    // to match the hash reconstructed at signing time from sheet data.
    var grievanceObj = {
      grievanceId:   grievanceId,
      memberId:      memberData.memberId,
      firstName:     memberData.firstName,
      lastName:      memberData.lastName,
      articles:      data.articles || '',
      issueCategory: data.issueCategory || '',
      description:   data.description || '',
      remedy:        '',
      step:          String(data.step || '1'),
      incidentDate:  data.incidentDate || '',
      dateFiled:     String(filingDate)
    };
    var documentHash = generateDocumentHash_(grievanceObj);

    var pdfBlob = generateDraftGrievancePDF_(pdfData, grievanceId, documentHash);
    var pdfSaved = false;

    if (pdfBlob && driveResult && driveResult.folderId) {
      try {
        var draftFileName = buildGrievancePdfName_(grievanceId, memberData.firstName, memberData.lastName, 'Draft', filingDate);
        pdfBlob.setName(draftFileName);
        var caseFolder = DriveApp.getFolderById(driveResult.folderId);
        caseFolder.createFile(pdfBlob);
        pdfSaved = true;
      } catch (pdfErr) {
        Logger.log('initiateGrievance: PDF save failed: ' + pdfErr.message);
      }
    }

    // ── 5. Notify member: Signature Needed ────────────────────────────────
    var stewardName = stewardEmail.split('@')[0]; // fallback
    try {
      // Try to resolve steward's display name
      for (var j = 1; j < allMembers.length; j++) {
        var sEmail = String(allMembers[j][MEMBER_COLS.EMAIL - 1] || '').toLowerCase().trim();
        if (sEmail === stewardEmail.toLowerCase().trim()) {
          var sFirst = String(allMembers[j][MEMBER_COLS.FIRST_NAME - 1] || '').trim();
          var sLast  = String(allMembers[j][MEMBER_COLS.LAST_NAME - 1] || '').trim();
          if (sFirst || sLast) stewardName = (sFirst + ' ' + sLast).trim();
          break;
        }
      }
    } catch (_e) { /* use fallback name */ }

    var notifTitle = 'Signature Needed (' + grievanceId + ')';
    var notifBody = 'Your steward, ' + stewardName + ', has prepared a grievance (' + grievanceId + ') on your behalf and needs your signature. ' +
      'Please open the dashboard and go to your grievance to review and sign the document.';

    if (typeof pushNotification === 'function') {
      pushNotification(memberData.email, {
        title: notifTitle,
        body: notifBody,
        type: 'Signature Needed'
      });
    }

    // ── 6. Refresh badges ───────────────────────────────────────────────────
    if (typeof _refreshNavBadges === 'function') _refreshNavBadges();

    return {
      success: true,
      grievanceId: grievanceId,
      driveFolderUrl: driveFolderUrl,
      pdfSaved: pdfSaved,
      memberName: memberData.firstName + ' ' + memberData.lastName,
      message: 'Grievance ' + grievanceId + ' created for ' + memberData.firstName + ' ' + memberData.lastName + '.'
    };

  } catch (error) {
    Logger.log('initiateGrievance error: ' + error.message + '\n' + (error.stack || ''));
    return errorResponse(error.message, 'initiateGrievance');
  }
}

/**
 * Returns grievance form field options from Config sheet.
 * Used by the steward intake form to populate dropdowns dynamically.
 * @returns {Object} { articles, issueCategories, steps, managers, stewards, coordinators }
 */
function getGrievanceFormOptions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return {};

    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    var result = {
      articles: [],
      issueCategories: [],
      steps: [],
      managers: [],
      stewards: [],
      coordinators: [],
      jobTitles: [],
      officeLocations: []
    };

    if (!configSheet) return result;

    var configData = configSheet.getDataRange().getValues();

    // Read config columns (data starts row 3, index 2)
    var colMap = {
      articles:         CONFIG_COLS.ARTICLES,
      issueCategories:  CONFIG_COLS.ISSUE_CATEGORY,
      steps:            CONFIG_COLS.GRIEVANCE_STEP,
      managers:         CONFIG_COLS.MANAGERS,
      stewards:         CONFIG_COLS.STEWARDS,
      coordinators:     CONFIG_COLS.GRIEVANCE_COORDINATORS,
      jobTitles:        CONFIG_COLS.JOB_TITLES,
      officeLocations:  CONFIG_COLS.OFFICE_LOCATIONS
    };

    for (var key in colMap) {
      if (!colMap.hasOwnProperty(key) || !colMap[key]) continue;
      var colIdx = colMap[key] - 1;
      for (var r = 2; r < configData.length; r++) {
        var val = String(configData[r][colIdx] || '').trim();
        if (val) result[key].push(val);
      }
    }

    return result;
  } catch (e) {
    Logger.log('getGrievanceFormOptions error: ' + e.message);
    return {};
  }
}

/**
 * Standalone wrapper — run from Apps Script editor or menu to
 * re-create/repair the union events calendar.
 */
function SETUP_CALENDAR() {
  var result = setupDashboardCalendar();
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_e) { Logger.log('_e: ' + (_e.message || _e)); }
  var msg = result.success
    ? '✅ Calendar ready!\n\nCalendar: ' + result.calendarName +
      '\nID: ' + result.calendarId +
      '\n\nCalendar ID has been saved to the Config sheet.'
    : '❌ Calendar setup failed — check Apps Script logs.';
  if (ui) ui.alert('📅 Calendar Setup', msg, ui.ButtonSet.OK);
  else Logger.log('SETUP_CALENDAR: ' + JSON.stringify(result));
  return result;
}
