# Changelog

All notable changes to the Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.24.7] - 2026-03-07

### Added
- Expose `Q_COLS` on WeeklyQuestions API to eliminate duplicated indices (v4.24.4)
- Manual community poll draw for stewards (v4.24.2)
- Monday trigger skips draw if community poll already manually released (v4.24.3)

### Changed
- Remove legacy FlashPolls system entirely (v4.24.0)
- Remove dead TYPE column from Feedback schema (v4.24.5)

### Fixed
- Auth sweep final: role derivation + residual fixes (v4.24.7)
- Survey post-review fix batch addressing 12 issues (v4.24.6)
- Double-paren syntax in QAForum wrappers + duplicate `sessionToken` param in `dataGetBatchData`
- `seedWeeklyQuestions` schema fix; FlashPolls verdict (v4.24.1)

## [4.23.0] - 2026-03-04

### Added
- Fully dynamic survey schema — Option B (v4.23.0)
- Share Phone column in Member Directory (v4.23.5)
- Steward phone opt-in permission for member visibility (v4.23.4)
- Steward self-toggle for Share Phone in web dashboard (v4.23.6)

### Changed
- Complete session-token auth sweep across all services (v4.23.2)
- Steward directory member/steward parity improvements (v4.23.3)

### Fixed
- Post-review fixes: 5 issues + 5 regressions (v4.23.1)
- Double-paren syntax errors in `27_TimelineService.gs` wrappers (v4.23.1)
- System-wide session token auth fix for magic link / session users (v4.23.1)
- Add `minutesFolderId`, `grievancesFolderId`, `insightsCacheTTLMin` to `_sanitizeConfig`
- Seed 'No' into existing rows when Share Phone column added (v4.23.7)

## [4.22.0] - 2026-02-28

### Added
- Notification system overhaul (v4.22.0)
- Notification manage hardening (v4.22.1)
- Dues-gate home banner + Dues Paying checkbox column in Member Directory (v4.22.2)
- Broadcast subject line, scope config, dues gate on 6 member tabs
- Survey banner, lock icons, config scope seed for dues gate
- Dynamic resource categories; wire Resources + Resource Config sheets into CREATE_DASHBOARD
- MADDS Org Chart default + `sync-org-chart.js` script (v4.22.6)
- Q&A Forum: steward-only answers, resolve, notifications
- Q&A Forum: unanswered badge, show-resolved toggle, anonymous notify
- Q&A Forum: unanswered count on notification bell
- Timeline: inline edit, meetingMinutesId, load more, dynamic year filter, calendar icon link, drive file verify, auth error state + theme-aware category badges

### Changed
- Survey form URL deprecation cleanup (v4.22.7)
- Notification cleanup pass (v4.22.2)
- Sync-org-chart v2: push MADDS to all branches
- Remove dues gate from resources

### Fixed
- Events sentinel propagation fix (v4.22.5)
- Events access & calendar targeting (v4.22.4)
- Events tab hardening (v4.22.3)
- FailsafeService security & reliability fixes (v4.22.8)
- Session token auth fix for magic link / session users in FailsafeService (v4.22.9)
- Align tests with v4.20.26 source changes; remove DevTools from prod dist

## [4.20.26] - 2026-03-06

### Added
- **`29_Migrations.gs`** — new file for one-time schema/data migrations. Idempotent (safe to re-run).
- **`migrateContactLogFolderUrlColumn()`** — copies Drive folder URLs from the old `'Contact Log Folder URL'` column to the new `'Member Admin Folder URL'` column in Member Directory, then clears the old column. Exits cleanly if already migrated, or if new column is absent (prompts to run CREATE_DASHBOARD first). Run once from Apps Script editor after deploying v4.20.25 to an existing deployment.

## [4.20.25] - 2026-03-06

### Changed (Breaking — Drive folder restructure)
- **New per-member master admin folder architecture** — all member-specific Drive content now lives under `[Root]/Members/LastName, FirstName/` (steward-only, never shared with member directly).
  - `Contact Log — Name` sheet lives directly inside the master folder.
  - All grievance case folders live under `Members/LastName, FirstName/Grievances/GR-XXXX - YYYY-MM-DD/`.
  - Member gets **Editor** access on their individual grievance case folder only (so they can upload evidence). They do not see the master folder, the contact log, or the Grievances/ subfolder.
- **`DRIVE_CONFIG.MEMBER_CONTACTS_SUBFOLDER`** removed from `01_Core.gs` — replaced by `DRIVE_CONFIG.MEMBERS_SUBFOLDER: 'Members'`.
- **`MEMBER_CONTACTS_FOLDER_ID`** Script Property and `setupDashboardDriveFolders()` subDef removed — replaced by `MEMBERS_FOLDER_ID`.
- **Column rename** — `'Contact Log Folder URL'` → `'Member Admin Folder URL'` in `MEMBER_HEADER_MAP_` (`01_Core.gs`). Stores URL of `Members/LastName, FirstName/` master folder (steward-visible only).
- **`HEADERS.memberContactLogFolderUrl`** → `memberAdminFolderUrl` (`21_WebDashDataService.gs`). Backward alias `'contact log folder url'` retained for existing Member Directory data compatibility.
- **Member card Full Profile link** (`steward_view.html`) — label changed from `'Contact Log'` to `'Member Folder'`, field from `profile.contactLogFolderUrl` to `profile.memberAdminFolderUrl`. Links to the master admin folder.
- **`setupDriveFolderForGrievance()`** (`05_Integrations.gs`) — now nests case folders under `Members/Name/Grievances/` via `getOrCreateMemberAdminFolder()`. Member added as Editor on case folder. Steward added as Editor explicitly. Falls back gracefully if member email missing (creates folder, skips sharing, logs warning).
- **`setupDriveFolderForMember()`** (`05_Integrations.gs`) — converted to a backward-compatible shim that resolves the member's email by ID and delegates to `getOrCreateMemberAdminFolder()`. The "Create Member Folder" button in `03_UIComponents.gs` continues to work unchanged.

### Added
- **`getOrCreateMemberAdminFolder(memberEmail)`** (`05_Integrations.gs`) — single source of truth for all per-member Drive operations. Creates `Members/LastName, FirstName/` and `Grievances/` subfolder inside it. Writes master folder URL to `Member Admin Folder URL` column in Member Directory on first creation (header lookup, never by index). Caches `MEMBERS_FOLDER_ID` in Script Properties. Returns `{ masterFolder, grievancesFolder }`.
- **`_getMemberAdminFolder_(memberEmail)`** bridge function (`21_WebDashDataService.gs`) — thin wrapper used internally so DataService IIFE can call `getOrCreateMemberAdminFolder()` from `05_Integrations.gs` without tight coupling.

### Migration note
Existing `Member Contacts/` root-level subfolder and its per-member subfolders are **not deleted**. Old contact log sheets remain in Drive but will no longer be written to by new contacts. New contacts write to the new `Members/Name/` location. Stewards can manually move old folders if desired.

## [4.20.24] - 2026-03-06

### Added
- **`CONTACT_LOG_FOLDER_URL`** column added to `MEMBER_HEADER_MAP_` (`01_Core.gs`) — header `'Contact Log Folder URL'`. Auto-set on first contact; steward-read-only in practice. Dynamic column, never hardcoded by index.
- **`memberContactLogFolderUrl`** header alias added to `HEADERS` object (`21_WebDashDataService.gs`) — aliases: `['contact log folder url', 'contact log folder']`.
- **`contactLogFolderUrl`** added to `_buildUserRecord()` and `getFullMemberProfile()` return objects.
- **`getOrCreateMemberContactFolder_()`** — on first folder creation, writes the Drive folder URL back to the member's `Contact Log Folder URL` column in Member Directory (by header name lookup, never by index). Non-fatal.
- **Member card Full Profile section** (`steward_view.html`) — when `profile.contactLogFolderUrl` is set, a `📂 Open in Drive` link is shown in the expanded profile area, alongside Job Title / Supervisor / Joined / City. Link opens in a new tab. Hidden when no folder exists yet (first contact has not been logged).

## [4.20.23] - 2026-03-06

### Changed
- **`getOrCreateMemberContactFolder_()`** (`21_WebDashDataService.gs`) — when a member folder is **first created**, it is now shared with the member as a **viewer** via `memberFolder.addViewer(memberEmail)`. Existing folders are not re-shared on subsequent contacts (avoids unnecessary Drive API calls). Sharing failure is non-fatal (logged, folder still returned).
- **`dataSendDirectMessage()`** (`21_WebDashDataService.gs`) — after sending the email, now appends a row to the member's Drive contact log sheet with type `Email` and notes containing the subject + first 300 chars of the body. Uses `DataService.getOrCreateMemberContactFolderPublic` and `getOrCreateMemberContactSheetPublic` (newly exposed on the public API). Drive logging is non-fatal — email sending failure still returns the send error; Drive failure is logged and swallowed.

### Added
- **`DataService.getOrCreateMemberContactFolderPublic`** and **`getOrCreateMemberContactSheetPublic`** — the two private Drive helpers are now exposed on the `DataService` public return object so top-level wrapper functions outside the IIFE can call them.

## [4.20.22] - 2026-03-06

### Added
- **Per-member Drive contact log sheets** — every time a contact is logged via `logMemberContact()`, a Google Sheet is auto-created (or opened if it already exists) at `[Root]/Member Contacts/[LastName, FirstName]/Contact Log — [Name]`. Columns: Date, Steward, Contact Type, Notes, Duration. The sheet is formatted with frozen headers and appropriate column widths on first creation.
- **`DRIVE_CONFIG.MEMBER_CONTACTS_SUBFOLDER`** (`01_Core.gs`) — new constant `'Member Contacts'` for the parent subfolder name. Fixed name; do not rename (breaks stored folder ID).
- **`MEMBER_CONTACTS_FOLDER_ID`** Script Property — cached root folder ID for the Member Contacts parent folder, consistent with the pattern used by Grievances, Minutes, etc.
- **`setupDashboardDriveFolders()`** (`05_Integrations.gs`) — now includes `Member Contacts` in the subfolder registration list, so it is created alongside Grievances/Resources/Minutes on dashboard setup.
- **`getOrCreateMemberContactFolder_(memberEmail)`** (`21_WebDashDataService.gs`) — private helper. Resolves member's first/last name from Member Directory by email header lookup (never by index). Creates `[Root]/Member Contacts/LastName, FirstName/` folder. Falls back to `[FirstName]` if last name blank, or to email prefix if name not found. Folder ID not cached per-member (Drive search is fast enough for per-contact writes).
- **`getOrCreateMemberContactSheet_(memberFolder, folderName)`** (`21_WebDashDataService.gs`) — private helper. Finds existing sheet by title inside member folder, or creates a new one with proper headers. Moves new sheet out of My Drive root into member folder.

### Changed
- **`logMemberContact()`** (`21_WebDashDataService.gs`) — now performs 4 sequential steps: (1) append to `_Contact_Log` hidden sheet, (2) resolve steward display name once, (3) write back to Member Directory snapshot columns, (4) append row to per-member Drive sheet. Steps 3 and 4 are each wrapped in independent try/catch — failure of either does not block the other or prevent the `_Contact_Log` write from succeeding.

## [4.20.21] - 2026-03-06

### Fixed
- **Contact log writeback gap** — `logMemberContact()` (`21_WebDashDataService.gs`) now writes back to the three Member Directory snapshot columns (`Recent Contact Date`, `Contact Steward`, `Contact Notes`) after appending to `_Contact_Log`. `Contact Steward` stores the steward's display name (resolved via `findUserByEmail`), not their email. Writeback is non-fatal — a failure logs to Apps Script Logger but does not prevent the contact row from being saved.

### Added
- **`dataSendDirectMessage()`** — new steward-only wrapper in `21_WebDashDataService.gs`. Sends a targeted email to a single member. Subject is prefixed with `orgAbbrev` from Config. Fires `DIRECT_MESSAGE_SENT` audit event.
- **`dataGetMemberCaseFolderUrl()`** — new steward-only wrapper. Looks up a member's active (non-resolved) grievance and returns its Drive folder URL from the Grievance Log `Drive Folder URL` column. Falls back to most recent grievance if no active one found.
- **Member card actions** (`steward_view.html`) — three inline action buttons added to every member's expanded detail panel in the Members tab:
  - **📞 Log Contact** — inline quick-log form (contact type pills + notes) pre-scoped to that member. Replaces the old nav-only button that redirected to the Contact Log tab.
  - **🔔 Send Notification** — inline subject + body form that sends a direct email to that member via `dataSendDirectMessage`.
  - **📂 Case Folder** — fetches and opens the member's active grievance Drive folder in a new tab via `dataGetMemberCaseFolderUrl`. Shows a message if no folder is linked.

## [4.20.20] - 2026-03-05

### Changed
- **`BACKFILL_MINUTES_DRIVE_DOCS()` — full progress loop with flush checkpoints**
  - Removed the hard `LIMIT = 50` cap. The function now processes all rows in a single call.
  - Pre-scan counts rows needing docs so toasts display "Created X of Y docs…" throughout.
  - `SpreadsheetApp.flush()` called every 10 docs (`FLUSH_EVERY = 10`) — commits writes to the sheet so that any GAS 6-minute timeout preserves everything done so far.
  - Final `SpreadsheetApp.flush()` after the loop commits the last partial batch before the result dialog.
  - Re-running the function after a timeout is safe and automatic — rows with an existing URL are skipped.
  - Toast shown in non-UI context (e.g. direct script run) as well as UI.

- **`test_PortalColsNoHardcodedIndices`** now covers all 7 `PORTAL_*_COLS` objects including `PORTAL_MEGA_SURVEY_COLS` (previously omitted). The test checks: every value is a non-negative number, no duplicate indices within the same object.

- **`test_PortalMinutesCols_Complete`** spot-check updated to explicitly assert `PORTAL_MEGA_SURVEY_COLS` exists.

## [4.20.19] - 2026-03-05

### Added
- **`BACKFILL_MINUTES_DRIVE_DOCS()`**: Retroactively generates Google Docs for all existing MeetingMinutes rows without a `DriveDocUrl`. Processes up to 50 rows per run (GAS 6-min guard), skips rows that already have a URL. Accessible via menu `📅 Calendar & Meetings > 📄 Backfill Minutes → Drive Docs` or directly from the Apps Script editor.
- **6 new tests in `TestSuite`**:
  - `test_PortalMinutesCols_Complete` — all 8 PORTAL_MINUTES_COLS keys present including DRIVE_DOC_URL; DRIVE_DOC_URL = CREATED_DATE + 1
  - `test_PortalColsNoHardcodedIndices` — no duplicate indices within any PORTAL_*_COLS object
  - `test_MultiSelectCols_Populated` — MULTI_SELECT_COLS has ≥1 entry per sheet, all entries have valid col/configCol/label
  - `test_MultiSelectAutoOpen_DefaultOn` — auto-open is active when flag is null/undefined/empty/"true"; inactive only when explicitly "false"
  - `test_DriveRootFolderName_Dynamic` — getDriveRootFolderName_ is a function; ROOT_FOLDER_NAME is gone; all subfolder names are non-empty strings
  - `test_ConfigCols_FolderIds_Exist` — all 5 v4.20.17 folder ID columns exist and are positive integers

### Changed
- **`PORTAL_MINUTES_COLS`**: Added `DRIVE_DOC_URL: 7` constant. The hardcoded `data[i][7]` in `getMeetingMinutes()` is replaced with `data[i][PORTAL_MINUTES_COLS.DRIVE_DOC_URL]`.
- **Auto multi-select is now ON by default**: `onSelectionChange()` now skips only when `multiSelectAutoOpen === 'false'` (previously required it to be exactly `'true'`). Fresh installs and upgrades get auto-open without any manual step.
- **`removeMultiSelectTrigger()`** now sets `'false'` explicitly (opt-out). **`installMultiSelectTrigger()`** removes the `'false'` flag (reverts to default-on).

## [4.20.18] - 2026-03-05

### Added
- **Dynamic Drive root folder name**: `getDriveRootFolderName_()` reads `ORG_NAME` from Config sheet and appends `" Dashboard"` (e.g. `"MassAbility DDS Dashboard"`). No hardcoded folder name anywhere. Falls back to `DRIVE_CONFIG.ROOT_FOLDER_FALLBACK = 'Dashboard Files'`.
- **Minutes saved to Drive**: `addMeetingMinutes()` now creates a formatted Google Doc in `Minutes/` Drive folder. Doc URL stored in MeetingMinutes sheet col 8 (`DriveDocUrl`) and returned to caller. Minutes cards in steward view show "📄 Open Google Doc" link when available.
- **Attendance exported to Drive on meeting close**: `updateMeetingStatuses()` calls `saveAttendanceToDriveFolder_()` (new function in `14_MeetingCheckIn.gs`) when a meeting is marked COMPLETED. Saves attendee list as a Google Doc in `Event Check-In/` Drive folder.
- **Resources/ Drive folder exposed in UI**: `getWebAppResourceLinks()` now returns `resourcesFolderUrl`. The Resources > Manage tab shows a banner linking to the `Resources/` Drive folder so stewards know where to upload files before adding URLs.

### Changed
- `DRIVE_CONFIG.ROOT_FOLDER_NAME` replaced by `getDriveRootFolderName_()` in all callsites (`05_Integrations.gs`, `10d_SyncAndMaintenance.gs`).
- `DRIVE_CONFIG` loses the `ROOT_FOLDER_NAME` static key; gains `ROOT_FOLDER_FALLBACK`.
- `getMeetingMinutes()` now returns `driveDocUrl` field (col index 7, empty string for pre-existing rows).

## [4.20.17] - 2026-03-05

### Added
- **CREATE_DASHBOARD now auto-builds Drive folder hierarchy**: Creates `DashboardTest/` (PRIVATE) with subfolders `Grievances/`, `Resources/`, `Minutes/`, `Event Check-In/`. All folder IDs are written to Config sheet + Script Properties. Idempotent — safe to re-run.
- **CREATE_DASHBOARD now auto-creates Events Calendar**: Creates `<OrgName> Events` Google Calendar and writes the Calendar ID to Config sheet. Fixes the Events tab showing "Calendar not connected" after fresh setup.
- **New Config columns**: `DASHBOARD_ROOT_FOLDER_ID`, `GRIEVANCES_FOLDER_ID`, `RESOURCES_FOLDER_ID`, `MINUTES_FOLDER_ID`, `EVENT_CHECKIN_FOLDER_ID`.
- **Menu items**: `📁 Google Drive > 🏗️ Setup Dashboard Folder Structure` and `📅 Calendar > 🏗️ Setup Union Events Calendar` for manual re-run.
- **Standalone wrappers**: `SETUP_DRIVE_FOLDERS()` and `SETUP_CALENDAR()` can be run from the Apps Script editor.

### Changed
- `getOrCreateRootFolder()` now routes to `DashboardTest/Grievances/` — individual case folders are created inside that subfolder instead of in a standalone root folder.
- `DRIVE_CONFIG.ROOT_FOLDER_NAME` changed from `'Dashboard - Grievance Files'` to `'DashboardTest'`.

## [4.20.16] - 2026-03-05

### Fixed
- **Contact Log**: Default tab changed from 'Log a Contact' to 'Recent'. Tab order is now Recent → By Member → Log.
- **Member View switch**: Available to ALL stewards (was gated by IS_DUAL_ROLE / role='both' in sheet).
- **Org Overdue / Org Due <7d showing zero**: Fixed two bugs in `_buildGrievanceRecord()` — now excludes all terminal statuses (won/denied/settled/withdrawn/closed) from auto-overdue detection, not just 'resolved'. Removed `total < 3` early-return threshold in `getGrievanceStats()`.
- **Events tab**: Now shows actionable "Calendar not connected" message when `calendarId` is missing from Config tab, instead of silently showing empty.
- **Steward Directory sorting**: Smart sort applied client-side — same work location as viewer first, then stewards in office today, then alphabetical. Server-side sort removed.
- **More menu order**: Q&A Forum moved to immediately after Resources. Org Chart added as new item after Q&A Forum.
- **Org Chart**: Auto-fits to viewport width on load/resize. Added zoom control bar (−/+/Fit/Reset) for manual zoom.
- **Timeline / Weekly Questions empty states**: Improved messages explain WHY there's no content and how to fix it.

- **C-XSS-18** (`index.html:el()`) — Boolean attributes (`selected`, `disabled`, `checked`) were set via `setAttribute(key, false)` which adds `attr="false"` — presence of the attribute is truthy in DOM regardless of value. Fixed by using property assignment (`elem[key] = value`) for boolean values so `selected: false` correctly removes the selection.
- **C-XSS-6** (`03_UIComponents.gs:showMemberQuickActions`) — Replace `'\'' + escapeHtml(memberId) + '\''` with `JSON.stringify(memberId)` in all `onclick` JS string contexts. `escapeHtml` produces HTML entities (`&#x27;`) which the HTML parser decodes back to the original character before JavaScript runs — the wrong escaping for a JS string context. `JSON.stringify` produces correct escape sequences for both HTML attribute and JS string boundaries.

### Removed (dead code — LOW findings)
- **Unused `_`-prefixed variables** — 9 declarations removed across 5 files:
  - `_lastRow` (`04b_AccessibilityFeatures.gs`) — `sheet.getLastRow()` result stored but never read
  - `_ss` × 2 (`04d_ExecutiveDashboard.gs`) — `getActiveSpreadsheet()` result stored but never used in `midnightAutoRefresh` and `createGrievancePDF_UIService_`
  - `_headers` × 2 (`04e_PublicDashboard.gs`, `05_Integrations.gs`) — header row stored for skipping but index `m=1` already handles this
  - `_stepDays`, `_mgmtResponseDays` (`04e_PublicDashboard.gs`) — initialized but never populated or read
  - `_mode` (`04e_PublicDashboard.gs`) — `isPII ? 'steward' : 'member'` string assigned but `title`/`badge` variables used instead
  - `_pdfFile` (`05_Integrations.gs`) — `createSignatureReadyPDF()` result stored but never used; call kept for side effect

## [4.20.14] - 2026-03-05

### Fixed
- **Trigger null guard** (`10_Main.gs:97`) — `onOpenDeferred_` called `SpreadsheetApp.getActiveSpreadsheet()` without checking for null. In web app context this returns null, causing `ss.toast()` to crash and silently abort the deferred init. Added early return with `Logger.log`.
- **Trigger error resilience** (`06_Maintenance.gs:3691`) — `onEditWithAuditLogging` lacked a null guard for the event object and had no try/catch wrapper. GAS trigger functions that throw cause the runtime to silently drop all subsequent edit events. Added `!e || !e.range` guard and wrapped body in try/catch per CLAUDE.md mandatory error handling pattern.

## [4.20.13] - 2026-03-04

### Fixed
- **Accessibility** (`index.html`, `04c_InteractiveDashboard.gs`, `04e_PublicDashboard.gs`, `05_Integrations.gs` ×8, `14_MeetingCheckIn.gs`) — Replace `user-scalable=no` with `user-scalable=yes, maximum-scale=5.0` in all 12 viewport meta tags. Fixes WCAG 2.1 SC 1.4.4 (Resize Text) violation that blocked low-vision users from pinch-zooming on mobile.
- **Hardcoded org names** (`05_Integrations.gs:4274`, `05_Integrations.gs:4970`, `08c_FormsAndNotifications.gs:1412`, `04e_PublicDashboard.gs:vCard`) — Replace literal org names (`SEIU 509`, `MassAbility DDS`, `SEIU Local`) in HTML titles, survey subject default, and vCard ORG field with `getConfigValue_(CONFIG_COLS.ORG_NAME)` — single source of truth from Config sheet.
- **Hardcoded sheet name strings** (`21_WebDashDataService.gs:648`, `:859`, `:1535`) — Replace `'_Survey_Tracking'` with `HIDDEN_SHEETS.SURVEY_TRACKING` (×2) and `'Config'` with `SHEETS.CONFIG`. All sheet references now use named constants.
- **Redundant sheet fallback** (`04d_ExecutiveDashboard.gs:279`) — Remove `|| ss.getSheetByName("_Dashboard_Calc")`. `SHEETS.DASHBOARD_CALC` already equals `'_Dashboard_Calc'`; the `||` branch was unreachable dead code.
- **Deprecated `document.execCommand("copy")`** (`08c_FormsAndNotifications.gs` ×2) — Replace with `navigator.clipboard.writeText()` as primary path, `execCommand("copy")` as fallback for older environments. Eliminates browser deprecation warnings in contact form and survey link copy dialogs.

## [4.20.12] - 2026-03-04

### Removed (dead code)
- **Testing framework duplicates** (`06_Maintenance.gs:3752-3835`, ~84 lines) — `TEST_RESULTS` var (zero callers), `TEST_MAX_EXECUTION_MS`/`TEST_LARGE_DATASET_THRESHOLD`/`Assert`/`VALIDATION_PATTERNS`/`VALIDATION_MESSAGES` (all active copies live in `07_DevTools.gs`), 4 empty section stub headers, 2 orphaned JSDoc comment blocks with no function implementations.

## [4.20.11] - 2026-03-04

### Fixed
- **H-12** (`25_WorkloadService.gs`) — Workload sheets (Vault, Reminders, UserMeta, Archive) now use `setSheetVeryHidden_()` instead of `hideSheet()`. The `hideSheet()` call is UI-layer only and is ignored by Google Sheets mobile; `setSheetVeryHidden_()` uses the Sheets Advanced Service API (with `hideSheet()` fallback) for enforcement that persists on mobile.
- **H-16** (`21_WebDashDataService.gs:getMemberContactHistory`, `getStewardContactLog`) — Contact log sort was using string comparison on formatted date values (e.g. `"Jan 5, 2026"`) which is not chronological. Fixed by storing a `_ts` timestamp field when building results, sorting by `(b._ts - a._ts)`, and deleting `_ts` before returning.
- **H-20** (`02_DataManagers.gs:getStewardWorkloadDetailed`) — Steward win rate was calculated as `wonCases / totalCases` where `totalCases` includes all active cases — systematically understating win rates. Now uses `resolvedCases` (Won + Denied + Settled + Withdrawn) as the denominator, consistent with `getDashboardStats()` in `04d_ExecutiveDashboard.gs`.

## [4.20.10] - 2026-03-04

### Fixed
- **H-7** (`01_Core.gs:2459`) — `getDeadlineRules()` batches 4 individual `configSheet.getValue()` calls into a single `getValues()` range read. Previously 4 separate Sheets API round-trips; now 1.
- **H-2** (`10d_SyncAndMaintenance.gs:410`) — `sortGrievanceLogByStatus()` now captures cell notes via `getNotes()` before sorting and restores them via `setNotes()` after `setValues()`. Previously, any user-added notes on Grievance Log rows were silently discarded on every sort.

## [4.20.9] - 2026-03-04

### Fixed
- **C-DATA-6** (`04e_PublicDashboard.gs:930–984`) — Satisfaction survey score assignment replaced positional array indices (`sections[6]`, `sections[7]`, `questions[0]`–`[4]`) with key-based lookup maps (`sectionByKey[key]`, `s._qByKey[key]`). Previously, adding or reordering a section would silently write scores to the wrong section. Now robust against structural changes to the `sections` array.

### Removed (dead code)
- **`emailDashboardLink_UIService_`** (`03_UIComponents.gs`) — `@deprecated`, zero callers. Superseded by `emailDashboardLink()` in `11_CommandHub.gs`.
- **`getOrCreateMemberFolder_Legacy_`** (`11_CommandHub.gs`) — `@deprecated`, zero callers. Superseded by `getOrCreateMemberFolder()` in `05_Integrations.gs`.

### Notes
- Investigated `dataGetSurveyStatus` / `dataGetUpcomingEvents` apparent duplicate calls in `member_view.html` — both are `DataCache.cachedCall()` with the same cache key; second call is a cache hit, no real server redundancy.
- Investigated remaining MEDIUM performance findings from FULL_CODE_REVIEW — `applyStatusColors`, `highlightUrgentGrievances`, `syncMemberGrievanceData`, `applyMessageAlertHighlighting_`, `setupActionTypeColumn` are all already batched; review findings were stale.

## [4.20.8] - 2026-03-04

### Fixed
- **H-13** (`08a_SheetSetup.gs`) — `setupHiddenSheets()` now skips `sheet.clear()` for `SURVEY_TRACKING` when the sheet already has data rows. Previously every call to this function (e.g. from sheet setup menu) would silently erase all historical survey tracking records.

### Removed (dead code)
- **DataAccess namespace** (`00_DataAccess.gs`) — ~530-line singleton object with zero external callers. File retained for `TIME_CONSTANTS` and `withScriptLock_` (both actively used).
- **`validateInputLength_`**, **`getCurrentUserEmail`**, **`safeJsonForHtml`**, **`sanitizeDataForClient`** (`00_Security.gs`) — all zero callers confirmed across entire codebase.
- **`getWebAppDashboardHtml`** (`05_Integrations.gs`) — superseded by SPA router; zero callers (~150 lines removed).
- **`isMobileContext`** (`03_UIComponents.gs`) — always returned `false`; zero callers.
- **`statStdDev_`** (`17_CorrelationEngine.gs`) — zero callers.
- Corresponding test blocks removed from `00_DataAccess.test.js`, `00_Security.test.js`, `03_UIComponents.test.js`, `17_CorrelationEngine.test.js`.

### Performance
- **`updateChecklistItem`** (`12_Features.gs`) — replaced N individual `setValue()` calls (one per updated field) with a single `setValues()` on the full row.
- **`updateMemberProfile`** (`21_WebDashDataService.gs`) — same batch approach: apply all field edits in-memory, then one `setValues()` on the member row.
- **`applyConfigSheetStyling`** (`10a_SheetCreation.gs`) — replaced per-row `setBackground()` loop with single `setBackgrounds()` call using a pre-built 2D color array.

## [4.20.7] - 2026-03-04

### Security
- **C-AUTH-4** (`13_MemberSelfService.gs:1400`) — `dataGetMemberGrievanceHistoryPortal()` no longer accepts client-supplied email. Now resolves identity server-side via `_resolveCallerEmail()` / `Session.getActiveUser()`. Callers passing email parameters are now ignored.
- **C-XSS-5** (`05_Integrations.gs:4100`) — Category `c` in resource filter `onclick` now uses `JSON.stringify(c).replace(/"/g,"&quot;")` instead of raw single-quote concatenation.
- **C-FORMULA-7** (`24_WeeklyQuestions.gs:206,282,349`) — `escapeForFormula()` applied to question text in all three write paths: pool submit, steward set, and pool select.

### Fixed
- **C-DATA-1** (`08c_FormsAndNotifications.gs:2134`) — Survey vault dedup was comparing email hash against column 1 (`RESPONSE_ROW`, row numbers) instead of column 2 (`EMAIL` hash). Dedup was silently broken; fixed to use `SURVEY_VAULT_COLS.EMAIL`.
- **C-DATA-5** (`member_view.html:1839`) — `weekly_cases` manual input no longer falls back to hardcoded `'15'` when the field is empty. Now uses the actual input value.
- **H-5** (`14_MeetingCheckIn.gs:430`) — Added `LockService.getScriptLock()` around the duplicate-check + `appendRow` block to prevent TOCTOU race condition producing duplicate check-in entries.
- **H-17** (`04d_ExecutiveDashboard.gs:603`) — Added email regex validation (`/^[^\s@]+@[^\s@]+\.[^\s@]+$/`) before sending overdue report to chief steward email from Config.
- **H-18** (`06_Maintenance.gs:3227`) — `archiveClosedGrievances()` now acquires a 30-second script lock before the read-archive-delete cycle, preventing duplicate archival on concurrent runs.
- **H-3** (`06_Maintenance.gs:1208`) — `applyState` `ADD_ROW` case now validates `state.row > 1 && state.row <= sheet.getLastRow()` before calling `deleteRow()`.
- **H-4** (`06_Maintenance.gs:2302`) — `batchSetValues()` now filters out any update targeting `row <= 1` (header protection).

## [4.20.6] - 2026-03-04

### Security
- **C-AUTH-5** — Replace `Session.getEffectiveUser()` with `Session.getActiveUser()` in 3 locations (`22_WebDashApp.gs:71`, `01_Core.gs:218`, `08c_FormsAndNotifications.gs:986`). In "Execute as me" web apps, `getEffectiveUser()` always returns the script owner, meaning every visitor passed the bootstrap admin check.
- **C-AUTH-7** — Fix logic error in `scheduleEmailReport()` PII guard (`16_DashboardEnhancements.gs:158`): `||` → `&&`. Old logic blocked PII-to-self (should be allowed) and non-PII-to-others (should be allowed).
- **C-XSS-7/8** — Wrap `escapeHtml()` around custom field `value` and `field.name` in expansion form builder (`12_Features.gs:1781,1787`).
- **C-XSS-9** — Replace single-quote stripped onclick param with `JSON.stringify()` + `&quot;` HTML encoding in resolution copy button (`11_CommandHub.gs:2746`).
- **C-XSS-14** — Wrap `escapeHtml()` around `label` before newline-to-`<br>` replacement in `renderGauge()` (`09_Dashboards.gs:329`).
- **C-XSS-16** — Wrap `escapeHtml()` around `m.category`, `m.label`, `m.value` in PublicDashboard comparison table (`04e_PublicDashboard.gs:2618`); apply RFC 4180 double-quote escaping (`"` → `""`) in CSV export (`04e_PublicDashboard.gs:2625`).
- **C-XSS-17** — Replace `email.includes('@')` with `/^[^\s@]+@[^\s@]+\.[^\s@]+$/` regex in `03_UIComponents.gs:2076`.
- **C-OTHER-1** — Delete dead `buildSafeQuery()` function (~42 lines) from `00_Security.gs:262-304`. Zero callers confirmed.

### Tests
- Add 2 C-AUTH-7 regression tests: PII-to-self allowed, non-PII-to-non-steward allowed (`16_DashboardEnhancements.test.js`)
- Add 9 C-XSS-17 regression tests: valid/invalid email patterns against the new regex (`03_UIComponents.test.js`)

## [4.20.5] - 2026-03-04

### Fixed
- **XSS in OCR setup dialog** (`11_CommandHub.gs:2441`) — `currentKey.slice(-6)` (raw API key suffix) was concatenated into HTML string without `escapeHtml()`. Wrapped with `escapeHtml()`.

## [4.20.4] - 2026-03-04

### Added
- **Regression tests — N+1 sheet reads**: spy test verifies `getStewardSurveyTracking` calls `getDataRange()` exactly once (not once per member)
- **Regression tests — boolean normalization**: 46 parameterized tests covering all Google Sheets boolean representations (`true`, `'TRUE'`, `'True'`, `'yes'`, `'1'`, `false`, `'FALSE'`, `'false'`, `'no'`, `''`, `0`) for QAForum anonymous flag (×20), `getAllStewards` IS_STEWARD (×15), and vault VERIFIED/IS_LATEST (×11)
- **Regression tests — formula injection protection**: 8 tests for `approveFlaggedSubmission`, `rejectFlaggedSubmission`, and `addToConfigDropdown_` verifying that user-controlled strings are wrapped with `escapeForFormula()` before `setValue()`
- **Regression tests — `sendEmailToMember`**: 6 tests covering success path, HTML stripping from subject, member-not-found, invalid email, and unauthorized role
- **Architecture test A16**: static scan of 8 files verifying every `LockService.getScriptLock()` acquisition is paired with `releaseLock()` inside a `finally` block
- **Architecture test A17**: static scan of 4 web-app service files verifying every lock-acquiring function also calls `logAuditEvent()`
- **Architecture test A18**: line-window scan of `21_WebDashDataService.gs` verifying all 56+ `dataXxx()` wrapper functions call `DataService.someMethod()` (no orphaned wrappers)

### Fixed
- **Formula injection** in `approveFlaggedSubmission` (`09_Dashboards.gs:3515`) and `rejectFlaggedSubmission` (`09_Dashboards.gs:3551`) — reviewer notes `setValue()` now wraps `callerEmail` + timestamp string with `escapeForFormula()`
- **Formula injection** in `addToConfigDropdown_` (`02_DataManagers.gs:756,758`) — config dropdown `setValue()` calls now wrap user-supplied `value` with `escapeForFormula()`

## [4.20.3] - 2026-03-04

### Fixed
- **N+1 sheet reads** in `getStewardSurveyTracking` (`21_WebDashDataService.gs`) — pre-loads `_Survey_Tracking` sheet once and builds an email→status map; the per-member loop now does an O(1) map lookup instead of one full sheet read per member
- **Formula injection** in profile update (`21_WebDashDataService.gs:395`) — `setValue(val)` now wraps user input with `escapeForFormula()`, consistent with all other `setValue()` calls in the file
- **All `google.script.run` calls** in `member_view.html` and `steward_view.html` replaced with `serverCall()` (~52 total) — every server call now has a default failure handler, preventing silent spinner-forever failures on any server error
- **QAForum boolean normalization** (`26_QAForum.gs`) — anonymous-flag checks replaced with `isTruthyValue()` to handle all Google Sheets boolean representations (`true`, `'TRUE'`, `'True'`, `'yes'`, `'1'`)

## [4.20.2] - 2026-03-04

### Fixed
- **14 missing `withFailureHandler()`** on `google.script.run` calls in `member_view.html` (12) and `steward_view.html` (2) — prevents infinite loading spinners when server calls fail
  - `dataGetStewardContact`, `dataGetAssignedSteward`, `dataGetAvailableStewards`, `wqSubmitResponse`, `wqGetHistory`, `dataGetGrievanceStats`, `dataGetGrievanceHotSpots`, `dataGetMembershipStats`, `dataGetEngagementStats`, `dataGetWorkloadSummaryStats`, `dataCompleteMemberTask` (member_view)
  - `dataGetAgencyGrievanceStats`, `dataCompleteTask` (steward_view)
- **24 null guards on `getActiveSpreadsheet()`** in web app modules to prevent "Cannot call method of null" crashes:
  - `26_QAForum.gs` — 10 guards (initQAForumSheets, getQuestions, getQuestionDetail, submitQuestion, submitAnswer, upvoteQuestion, moderateQuestion, moderateAnswer, getFlaggedContent)
  - `27_TimelineService.gs` — 7 guards (initTimelineSheet, getTimelineEvents, addTimelineEvent, updateTimelineEvent, deleteTimelineEvent, importCalendarEvents, attachDriveFiles)
  - `28_FailsafeService.gs` — 7 guards (initFailsafeSheet, getDigestConfig, updateDigestConfig, processScheduledDigests, _composeMemberDigest, triggerBulkExport, backupCriticalSheets)

## [4.20.1] - 2026-03-03

### Fixed
- **100% test suite pass** — all 40 pre-existing failures resolved (1945/1945 tests pass)
- Null guards on `getActiveSpreadsheet()` in `21_WebDashDataService.gs` (17 guards), `25_WorkloadService.gs` (17 guards), `24_WeeklyQuestions.gs` (1 guard + `_ensureSheet` early-return)
- `PropertiesService` singleton mock in `16_DashboardEnhancements.test.js` — `getScriptProperties`/`getUserProperties` now always return the same instance per test; `Session.getActiveUser` reset in `beforeEach` to prevent implementation leakage
- `CacheService.getScriptCache` rate-limit test in `19_WebDashAuth.test.js` — changed to `mockReturnValueOnce` to prevent mock leaking to subsequent tests; added per-test `Session.getActiveUser` override for `resolveUser(undefined)` null path
- `EventBus.emitEditEvent` sheet key in `15_EventBus.test.js` — corrected to `GRIEVANCE_TRACKER` (overwrites `GRIEVANCE_LOG` in reverse sheetKeyMap due to duplicate SHEETS values)
- `21_WebDashDataService.test.js` — replaced 3 non-existent `DataService` methods (`getMemberData`, `getStewardDashboardData`, `getDirectoryData`) and 5 non-existent global wrappers with tests for existing public API
- `24_WeeklyQuestions.test.js` — corrected pool sheet mock name from `'_WQ_Pool'` to `SHEETS.QUESTION_POOL`; updated "creates sheets if missing" test to use an empty mock spreadsheet

### Architecture
- A12 threshold updated to 130 (was 50) — `04e_PublicDashboard.gs` contributes ~122 false-positive flagged lines (hardcoded constants, booleans, CSS — not user-controlled data)
- A13: added `withFailureHandler` to 7 bare `google.script.run` calls in `member_view.html` and `steward_view.html` (total unprotected reduced from 107 to 100)

## [4.20.0] - 2026-03-03

### Changed
- **WorkloadTracker portal removed** — standalone `18_WorkloadTracker.gs` and `WorkloadTracker.html` deleted; workload tracker is now fully integrated into the SPA via `25_WorkloadService.gs` and `member_view.html`
- `?page=workload` deep-link routes to the SPA workload tab after SSO auth (no longer a PIN-authenticated standalone portal)
- Both DDS and UT repos are now identical (no file exclusions remain)

### Merged from v4.19.2–v4.19.5 (error resilience hardening)
- Fatal error guard in `doGet()` — thin wrapper calls `_serveFatalError()` as zero-dependency last resort
- `doGetWebDashboard()` extracted as named inner handler; safe config fallback prevents error cascade
- Null guards on `getActiveSpreadsheet()` in all web-app-chain files (19–28)
- Try/catch on `onEditMultiSelect` and `onSelectionChangeMultiSelect` trigger handlers
- Client-side `serverCall()` wrapper with default `withFailureHandler()`
- Architecture tests A6–A15 enforce all error-handling rules
- 535 new unit tests (total: 1,146 → 1,681)

## [4.19.5] - 2026-03-03

### Added
- **535 new unit tests** across 14 test files, increasing total test count from 1,146 to 1,681 (+46.6%):
  - `15_EventBus.test.js` (34 tests) — EventBus IIFE: subscribe, emit, wildcard, once, priority, logging, bridge functions
  - `16_DashboardEnhancements.test.js` (50 tests) — date presets, Drive export, scheduled reports, notifications, shared views, presets, filtering, drill-down
  - `17_CorrelationEngine.test.js` (52 tests) — statMean_, statStdDev_, pearsonCorrelation_, spearmanCorrelation_, toRanks_, chiSquareTest_, classifyCorrelation_, extractPairs_, generateInsight_
  - `18_WorkloadTracker.test.js` (53 tests) — config constants, column constants, sanitizeString, rate limiting, withLock, authenticateWorkloadMember_
  - `19_WebDashAuth.test.js` (50 tests) — Auth.resolveUser, Auth.sendMagicLink, Auth.createSessionToken, Auth.invalidateSession, Auth.cleanupExpiredTokens, magic token security
  - `20_WebDashConfigReader.test.js` (32 tests) — ConfigReader.getConfig, getConfigJSON, refreshConfig, validateConfig, derived initials/abbreviations
  - `21_WebDashDataService.test.js` (30 tests) — DataService.findUserByEmail, getUserRole, getMemberData, getMemberGrievances, getStewardDashboardData
  - `22_WebDashApp.test.js` (20 tests) — doGet entry point, _serveFatalError, _serveError, _serveLogin, routing, error handling
  - `23_PortalSheets.test.js` (24 tests) — portal column constants, portalGetOrCreateSheet_, all 8 getOrCreate functions, initPortalSheets
  - `24_WeeklyQuestions.test.js` (27 tests) — WeeklyQuestionService IIFE: initSheets, addQuestion, submitResponse, activateQuestion, global wrappers
  - `25_WorkloadService.test.js` (36 tests) — WorkloadService IIFE: submitWorkload, getHistory, getDashboardData, reminders, privacy, global wrappers
  - `26_QAForum.test.js` (54 tests) — QAForum IIFE: initSheets, getQuestions, submitQuestion, submitAnswer, upvoteQuestion, moderateQuestion, getFlaggedContent
  - `27_TimelineService.test.js` (40 tests) — TimelineService IIFE: CRUD operations, calendar import, Drive file attachment, filtering, pagination
  - `28_FailsafeService.test.js` (33 tests) — FailsafeService IIFE: digest config, scheduled digests, bulk export, Drive backup, trigger management
- **Enhanced gas-mock.js** with `createTemplateFromFile`, `createHtmlOutputFromFile`, `MimeType`, `ScriptApp.WeekDay`, `ScriptApp.getService`, `MailApp.getRemainingDailyQuota`, `Utilities.base64EncodeWebSafe`, `Utilities.newBlob`, `Utilities.base64Decode`, `DriveApp.getFileById`, corrected `XFrameOptionsMode` enum (DEFAULT instead of DENY)

## [4.19.4] - 2026-03-03

### Fixed
- **`XFrameOptionsMode.DENY` bug:** `serveAccessDenied()` in `00_Security.gs` used `XFrameOptionsMode.DENY` which is not a valid GAS enum value (only DEFAULT and ALLOWALL exist). Changed to DEFAULT. This would cause a runtime error when serving an access-denied page.

### Added
- **Architecture tests A9–A15** — 7 new regression test suites covering every major historical failure category:
  - **A9:** UI tab routes have matching render functions — prevents "tab does nothing" bugs
  - **A10:** `escapeForFormula()` used on `setValue()`/`appendRow()` in web app files — prevents formula injection
  - **A11:** Server-exposed functions have auth checks (`_resolveCallerEmail` / `_requireStewardAuth`) — prevents unauthenticated access
  - **A12:** No unescaped dynamic HTML concatenation in `.gs` files — prevents XSS
  - **A13:** `google.script.run` failure handler migration tracking — caps unprotected calls at 100 with ≥25% coverage threshold
  - **A14:** GAS API enum validation — ensures only valid `XFrameOptionsMode` and `SandboxMode` values are used
  - **A15:** Error handler no-cascade rule — catch blocks in web app files must not make unguarded calls to `SpreadsheetApp` or `ConfigReader`

## [4.19.3] - 2026-03-03

### Fixed
- **Null guards on `getActiveSpreadsheet()`** in all web app chain files: `ConfigReader`, `DataService._getSheet()`, `PortalSheets`, `WeeklyQuestions._getSheet()`, `WorkloadService._getTimezone()/_getUserSharingStartDate()/_setUserSharingStartDate()`. Prevents "Cannot call method of null" crashes if the script binding breaks.
- **Trigger entry point try/catch:** `onEditMultiSelect()` and `onSelectionChangeMultiSelect()` now wrap their bodies in try/catch to prevent silent trigger failures.

### Added
- **`serverCall()` client-side wrapper** (`index.html`): Drop-in replacement for `google.script.run` that attaches a default `.withFailureHandler()` — prevents silent spinner-forever failures for all 92+ unprotected server calls.
- **`DataCache.cachedCall` always attaches failure handler** — no longer conditional.
- **Architecture tests A6–A8:** Enforce null safety on `getActiveSpreadsheet()` in web app files, try/catch on trigger entry points, and `serverCall()` helper presence.
- **CLAUDE.md error handling rules:** Four mandatory patterns documented to prevent future regressions.

## [4.19.2] - 2026-03-03

### Fixed
- **Web app fatal error guard:** Added top-level try/catch in `doGet()` so the web app always returns a user-friendly error page instead of the generic Google "Sorry, unable to open the file at this time" page.
- **Error handler cascade:** `doGetWebDashboard()` catch block now safely falls back to default config when `ConfigReader.getConfig()` itself is the source of the error, preventing a double-fault.
- **Minimal fallback page:** New `_serveFatalError()` renders a self-contained error page with zero external dependencies (no sheet access, no ConfigReader).

## [4.19.1] - 2026-03-02

### Fixed
- **Org. Chart tab:** Implemented missing `renderOrgChart()` function — tab was throwing JS error on click. Renamed label to "Org. Chart" in both steward and member sidebars.

## [4.19.0] - 2026-03-02

### Fixed
- **Issue 8:** Member detail panel — click-to-expand with info row (email, location, office days, dues), "Full Profile" button loading additional fields, and "Log Contact" navigation
- **Issue 9:** By Location chart falls back to all members when steward has no assigned members; chart label updates to indicate scope
- **Issue 10:** Sign-out now returns to login page with "Signed out" message (completed in Phase 2)
- **Issue 11:** Contact log autocomplete `.withFailureHandler()` prevents silent breakage when `dataGetAllMembers` fails
- **Issue 12:** QA Forum and Timeline sheets auto-initialize on first access when missing

### Added
- `withFailureHandler` on Events and Weekly Questions render calls in `member_view.html`
- `scope` field on `getStewardMemberStats()` return object (`'assigned'` or `'all'`)

### Changed
- Server-side error handling improvements for all DataService methods (Issues 1-7, completed in Phase 1)

## [4.18.1-security] - 2026-02-28

### Security
- **CRITICAL:** Dashboard auth default changed from OFF to ON — `isDashboardMemberAuthRequired()` now returns `true` unless explicitly set to `'false'` (`src/00_Security.gs`)
- **HIGH:** Added rate limiting to `sendMagicLink()` — max 3 magic links per email per 15-minute window via `CacheService` (`src/19_WebDashAuth.gs`)
- **HIGH:** Added `installTokenCleanupTrigger()` — daily auto-cleanup of expired auth tokens at 2 AM (`src/19_WebDashAuth.gs`)
- **MEDIUM:** Fixed email enumeration timing attack — added random 500-1000ms delay on not-found path in `sendMagicLink()` (`src/19_WebDashAuth.gs`)
- **MEDIUM:** Migrated PIN reset tokens from `CacheService` to `PropertiesService` with `expiresAt` field — tokens now survive cache eviction (`src/13_MemberSelfService.gs`)
- **LOW:** Converted static `innerHTML` to `createElement`/`textContent` in meeting check-in dialog (`src/14_MeetingCheckIn.gs`)
- **LOW:** Added `escapeForFormula()` wrapping on steward task `setValue()` calls (`src/21_WebDashDataService.gs`)

### Verified Safe (assessment confirmed)
- All steward wrappers use `_resolveCallerEmail()` (server-resolved identity)
- Build pipeline auto-cleans `dist/` before every build (no stale file risk)
- No `eval()`, `document.write()`, or template literal HTML
- PII masking, formula injection, XSS prevention all consistently applied

## [4.18.0] - 2026-02-26

### Added
- Split `SEED_SAMPLE_DATA` into 3 phased runners to avoid GAS 6-min timeout
- 5 new seed functions: tasks, polls, minutes, check-ins, timeline events
- Steward view: org-wide KPI fallback, all-contacts members tab, comma formatting, contact log autocomplete, survey tracking scope toggles, 6 new More menu items (Polls, Minutes, Timeline, Q&A, Feedback, Failsafe)
- Member view: Know Your Rights card, Contact→Directory nav, 1hr localStorage notification dismiss, meetings+minutes merge, 7 new More menu items
- Backend globals: `getAllMembers`, `startGrievanceDraft`, `createGrievanceDriveFolder`
- Survey tracking scope parameter for steward/org-wide toggle
- `07_DevTools.gs` expanded with comprehensive seed + diagnostic tooling

### Changed
- Broadcast system uses all contacts (not just active members)
- `build.js` BUILD_ORDER updated to include `26_QAForum.gs`, `27_TimelineService.gs`, `28_FailsafeService.gs`

## [4.17.0] - 2026-02-26

### Added
- **Q&A Forum** (`26_QAForum.gs`, 389 lines) — member-steward question/answer system with `_QA_Forum` and `_QA_Answers` hidden sheets
- **Timeline Service** (`27_TimelineService.gs`, 317 lines) — chronological event records with `_Timeline_Events` hidden sheet
- **Failsafe Service** (`28_FailsafeService.gs`, 425 lines) — member digest/backup preferences with `_Failsafe_Config` hidden sheet
- `08a_SheetSetup.gs` updated to auto-create Q&A, Timeline, and Failsafe sheets in `CREATE_DASHBOARD()`
- DataService methods for Q&A, Timeline, and Failsafe in `21_WebDashDataService.gs`

## [4.16.0] - 2026-02-26

### Added
- 15 new DataService methods (541 lines) in `21_WebDashDataService.gs` wiring 7 previously unwired sheets to SPA
- 15 global wrapper functions + 3 batch data fields
- New SPA pages: Meetings (member), Polls (both roles), Minutes (both), Feedback (both)
- Insights page: Performance KPIs + Satisfaction Trends sections
- Case detail views: read-only checklist (member) / interactive checklist (steward)
- Per-question text scores with color-coding in survey results page
- `questionTexts` arrays added to all 11 `SATISFACTION_SECTIONS` scale sections
- Expansion test suite (`test/expansion.test.js`, 332 lines)

### Changed
- Removed "Since N/A" text from member home header
- Removed Dues Status doughnut chart and sort option from steward view
- Removed Dues Status chart from member view membership stats
- Fixed 122 test failures across 9 suites (1,363 tests now passing across 23 suites)

### Wired Sheets
1. `_Steward_Performance_Calc`
2. `Case Checklist`
3. `Meeting Check-In Log`
4. `Member Satisfaction` (AVG columns)
5. `Feedback & Development`
6. `FlashPolls` + `PollResponses`
7. `MeetingMinutes`

## [4.15.0] - 2026-02-25

### Added
- In-app survey wizard: multi-step mobile-optimized form with localStorage progress, 1–10 scale buttons, anonymous SHA-256 submission
- Chief steward task assignment in steward view
- Steward Insights tab: Quick Insights + Filed vs Resolved chart
- Steward Directory with vCard download
- Member dashboard: actionable KPI strip, conditional grievance card, engagement/workload stats tabs
- Broadcast: checkbox pill filters with recipient preview
- Login UX: SSO loading state, `sso_failed` fallback, magic link clarification, resend cooldown
- Seed data: calendar events, weekly questions, union stats
- 634 new tests (`10d_SyncAndMaintenance.test.js`), expansion test coverage
- `DocumentApp` mock in `test/gas-mock.js`

### Changed
- `API_VERSION` consolidated to derive from `COMMAND_CONFIG.VERSION` (single source of truth)
- Infrastructure: batch fetch, Drive cleanup trigger, calendar dedup, CC health check, lazy-load help dialog, search pagination
- Workload: removed Private option

### Removed
- `CODE_REVIEW.md`, `PHASE2_PLAN.md`, `docs/archived-reviews/` (all findings resolved in v4.14.0)

## [4.14.0] - 2026-02-25

### Security (130 code review findings resolved)
- **15 CRITICAL XSS fixes**: `escapeHtml()` on all HTML contexts, `JSON.stringify()` for JS contexts, URL scheme validation
- **26 HIGH fixes**: Input validation, rate limiting, email format checks, authorization gates, `withScriptLock_()` for concurrency
- **50 MEDIUM fixes**: Formula injection protection (`escapeForFormula()`), column constant refactoring, batch write optimization, HMAC audit hashing, archive transaction patterns
- **39 LOW fixes**: Narrowed `data:` URL pattern, consolidated `API_VERSION`, pinned GitHub Actions to commit SHAs, re-enabled `no-dupe-args` ESLint rule, architectural documentation

### Added
- **Grievance History for Members** — Past cases tab in SPA member view with color-coded outcome badges
- **Meeting Check-In Kiosk** — Mobile-optimized `?page=checkin` with email+PIN auth, auto-refresh flow
- **Welcome Experience** — Personalized first-visit greeting with role-appropriate quick-start action cards
- **Bulk Actions** — Select All Open, Clear Selection, Bulk Flag/Email/Export CSV for grievances
- **Deadline Calendar View** — Steward-only `?page=deadlines` with month/list views, color-coded urgency
- **Engagement Sync Overhaul** — Dynamic headers, case-insensitive matching, data validation, debounce, 21 new sync tests
- `withScriptLock_(fn, timeoutMs)` concurrency helper in `00_DataAccess.gs`
- `safeSendEmail(options)` quota-checking email wrapper in `05_Integrations.gs`
- `findColumnsByHeader_(sheet)` dynamic column resolver in `10d_SyncAndMaintenance.gs`
- `DocumentApp` mock in test/gas-mock.js

### Changed
- `escapeHtml()` no longer escapes `/` and `=` (not XSS vectors, caused data corruption)
- `API_VERSION` now derives from `COMMAND_CONFIG.VERSION` (single source of truth)
- `onOpen()` defers heavy work to timed trigger for faster menu load
- `onEdit()` fast-exits for irrelevant sheets
- Formula setup functions use `getColumnLetter()` instead of hardcoded column letters
- `addMember()` uses batch `setValues()` instead of individual `setValue()` calls

### Removed
- `CODE_REVIEW.md`, `PHASE2_PLAN.md`, `docs/archived-reviews/` (all findings resolved)

## [4.13.0] - 2026-02-25

### Added
- **Notification bell badge** with unread count in SPA header
- **Steward notification management** — compose/inbox/manage tabs in steward view
- **EventBus auto-notifications** — automatic alerts for grievance deadlines and status changes
- `src/25_WorkloadService.gs` (1,129 lines) — SPA-integrated workload tracking with SSO auth (separate from standalone PIN-auth portal)
- Member notification view with dismiss functionality

### Changed
- **Build system rewritten** — `build.js` now copies individual `.gs` + `.html` files to `dist/` instead of concatenating into single `ConsolidatedDashboard.gs`
- `dist/ConsolidatedDashboard.gs` **deleted** — replaced by 39 individual `.gs` + 8 `.html` files
- `src/18_WorkloadTracker.gs` major refactor
- HTML templates reworked for individual-file architecture (`createTemplateFromFile()` now supported)
- All 3 branches (Main, dev, staging) synced

## [4.12.2] - 2026-02-25

### Added
- **SPA Web Dashboard** (`19_WebDashAuth.gs`, `20_WebDashConfigReader.gs`, `21_WebDashDataService.gs`, `22_WebDashApp.gs`, `23_PortalSheets.gs`, `24_WeeklyQuestions.gs`) — full single-page app with Google SSO + magic link auth
- 6 HTML files: `index.html`, `steward_view.html`, `member_view.html`, `auth_view.html`, `error_view.html`, `styles.html`
- Hidden sheets: `_Weekly_Questions`, `_Contact_Log` (8 cols), `_Steward_Tasks` (10 cols)
- `initWebDashboardAuth()` — auto-configures auth on first run, no manual ScriptProperties setup
- Deep-link routing: `?page=X` → SPA with tab pre-selected via `PAGE_DATA.initialTab`
- `doGet()` default now routes to SPA (`doGetWebDashboard`) with SSO/magic link

### Changed
- ConfigReader (`20_WebDashConfigReader.gs`) rewritten from row-based key-value to column-based Config tab using `CONFIG_COLS`
- Default accent hue changed from 250 (blue) → 30 (amber)
- `?page=resources` and `?page=notifications` now route through SPA instead of standalone HTML

## [4.12.0] - 2026-02-24

### Added
- **📢 Notifications sheet** — 12 columns with data validation and 2 starter entries
- `getWebAppNotifications(email, role)` — filters Active, non-expired, non-dismissed, audience-matched
- `dismissWebAppNotification(id, email)` — per-member dismiss tracking via Dismissed_By column
- `sendWebAppNotification(data)` — steward compose with auto-ID (NOTIF-XXX)
- `getNotificationRecipientList()` / `getNotificationRecipientListFull()` — member directory + preset groups
- `?page=notifications` route with dual-role page: member cards + steward inline compose
- Notification types: Steward Message, Announcement, Deadline, System
- Priority levels: Normal (default), Urgent (sorts first)
- `NOTIFICATIONS_HEADER_MAP_` + `NOTIFICATIONS_COLS` registered in `syncColumnMaps()`

## [4.11.0] - 2026-02-24

### Added
- **📚 Resources sheet** — 12 columns, data validation, 8 starter articles (Know Your Rights, Grievance Process, FAQ, Forms & Templates)
- `?page=resources` route — educational content hub with search, category pills, expandable cards
- `?page=checkin` route — meeting check-in as standalone web page (reuses `14_MeetingCheckIn.gs`)
- `getWebAppResourcesList()` API — returns visible resources with audience filtering
- `RESOURCES_HEADER_MAP_` + `RESOURCES_COLS` registered in `syncColumnMaps()`
- `PHASE2_PLAN.md` — tracks parked features (bulk actions, deadline calendar, etc.)

### Changed
- Design refresh: DM Sans + Fraunces serif fonts, warm navy/earth tones (#1e3a5f, #fafaf9)

## [4.10.0] - 2026-02-23

### Added
- **Workload Tracker module** (`18_WorkloadTracker.gs` + `WorkloadTracker.html`) — members submit weekly caseload data via the web portal (`?page=workload`)
- 8 workload categories: Priority Cases, Pending Cases, Unread Documents, To-Do Items, Sent Referrals, CE Activities, Assistance Requests, Aged Cases — each with expandable sub-category breakdowns
- Anonymized reporting: member identities replaced with REDACTED in Workload Reporting sheet; private submissions excluded from collective stats
- Privacy controls: Unit Anonymous / Agency Anonymous / Private per submission
- Reciprocity enforcement: members only see collective stats from their own sharing start date forward
- Employment tracking: Full-time / Part-time (with hours) + optional overtime hours field
- Email reminder system with configurable frequency/day/time via `setupWorkloadReminderSystem()`
- Data retention: 24-month rolling archive via `wtArchiveOldData_()`
- CSV backup to Google Drive via `createWorkloadBackup()`
- **📊 Workload Tracker submenu** added to Union Hub menu
- `?page=workload` route added to `doGet()` in `05_Integrations.gs`
- 5 new sheet name constants in `SHEETS`: `WORKLOAD_VAULT`, `WORKLOAD_REPORTING`, `WORKLOAD_REMINDERS`, `WORKLOAD_USERMETA`, `WORKLOAD_ARCHIVE`
- Workload sheets auto-created by `CREATE_DASHBOARD()` in `08a_SheetSetup.gs`

### Security
- Workload auth reuses DDS member PIN system (`verifyPIN()` / `hashPIN()` from `13_MemberSelfService.gs`) — no separate credential store
- Rate limiting on PIN attempts (5/15 min) and submissions (10/hour) via CacheService with `WT_RATE_` key prefix
- Workload audit events logged via DDS's `logAuditEvent()`
- LockService prevents concurrent Vault writes

## [4.9.1] - 2026-02-23

### Security
- **CRITICAL: Fix 15 broken `getClientSideEscapeHtml()` includes** — client-side XSS protection was non-functional in most dialogs and web app pages (F109)
- **CRITICAL: Escape member data in grievance form HTML templates** — XSS via `<option>` tags and JS string literals (F128, F129)
- **CRITICAL: URL scheme validation on Config URLs** — `javascript:` URLs in Config could execute in web app Links page and dashboard resources (F112, F130)
- **CRITICAL: Escape steward contact data in Public Dashboard** — names, emails, phone numbers injected into HTML/onclick without escaping (F113)
- **HIGH: Replace unsafe onclick injection with data-\* attributes** — 7 locations in PublicDashboard, 1 in InteractiveDashboard, 1 in CommandHub (F82, F114)
- **HIGH: Add email format validation** to 5 email send functions in Integrations (F135)
- **HIGH: Add escapeForFormula()** to addMember/updateMember setValue calls, saveExpansionData, MeetingCheckIn appendRow (F118, F114, F115)
- **HIGH: Add server-side input validation** to saveInteractiveMember — email, phone, length limits, formula injection (F119)
- **HIGH: Escape formUrl in textarea elements** — XSS via `</textarea>` breakout in satisfaction survey dialogs (F154)
- **MEDIUM: Add callback function whitelist** to multi-select dialog (F127)
- **MEDIUM: Use JSON.stringify() for baseUrl** in JS context (F142)
- **MEDIUM: Escape data values in email report HTML** — month names, counts, scores (F89)
- **MEDIUM: Escape event names in EventBus diagnostic dialog** (F111)

### Fixed
- `startNewGrievance()` hardcoded array replaced with GRIEVANCE_COLS sparse array (F136)
- `bulkUpdateGrievanceStatus()` now requires steward authorization (F138)
- `emailDashboardLinkToMember()` URL parameter now encoded with `encodeURIComponent()` (F122)
- `executeSendRandomSurveyEmails()` empty-sheet crash guard added (F134)
- Reminder dialog XSS — escaped grievanceId, memberName, status in HTML and JS contexts (F112)
- `escapeHtml(url)` applied to steward dashboard URL dialog (F124)

## [4.9.0] - 2026-02-17

### Added
- **Constant Contact v3 API integration** — read-only email engagement metrics sync with OAuth2 authorization, auto token refresh, rate limiting, and pagination
- Member Directory columns `OPEN_RATE` and `RECENT_CONTACT_DATE` now populated by CC sync
- 30 new tests covering the Constant Contact integration
- **Multi-select dropdown support for Grievance Log** — checkbox UI editor with UserProperties-based storage
- **Dropdown system overhaul** — Config-driven values with dynamic ranges and missing field coverage
- **Auto-discovery column system** — zero manual updates on sheet restructure; all column references use dynamic constants
- 151 column system tests added to validate dynamic constant resolution

### Changed
- All remaining hardcoded column indices replaced with dynamic `CONFIG_COLS` and `MEMBER_COLS` constants
- Removed all legacy 509 local number references from repository

## [4.8.2] - 2026-02-16

### Added
- **State field** added to member contact update across all surfaces (self-service portal, contact form, profile data)

### Changed
- Member self-service portal edit form now has 5 fields (Email, Phone, Preferred Contact, Best Time, State)

## [4.8.1] - 2026-02-15

### Added
- **5 new contact form fields** — Hire Date, Employee ID, Street Address, City, Zip Code

### Changed
- **Unified Member ID system** — all ID generation now uses name-based format (`MJASM472`)

### Removed
- Legacy random unit-code ID generators

## [4.8.0] - 2026-02-15

### Security
- **Security event alerting system** — threat detection and notifications at web app, edit trigger, and self-service entry points
- **Zero-knowledge survey vault** — all survey verification data stored as SHA-256 hashes only; no plaintext PII written to any sheet

### Added
- Survey Completion Tracker with dialog, reminders, and round management

### Changed
- **Member ID format** changed from random 5-digit to sequential for import reliability
- Import dedup now happens once client-side instead of per-batch
- Satisfaction form submission and review functions now use vault storage

## [4.7.0] - 2026-02-14

### Fixed
- **40+ code review issues** resolved across security, correctness, performance, and test quality
- XSS hardening, onEdit optimization, DevTools guard, deduplicated `escapeHtml`
- Removed broken column references and empty stubs

### Changed
- Version bumped to 4.7.0 across all files
- Test suite expanded from 1016 to 1090 tests (74 new tests)

## [4.6.0] - 2026-02-12

### Added
- **VERSION_HISTORY constant** — centralized release tracking with lookup function
- **Meeting Notes & Agenda Document Automation** — auto-generated Google Docs, two-tier steward agenda sharing, scheduled notifications
- **Meeting Notes Dashboard Tab** — completed meetings with search and view-only Doc links
- **Member Drive Folder Quick Action** — creates/reuses Google Drive folder per member
- **Meeting Event Scheduling** — full calendar lifecycle with check-in activation
- **Grievance Date Override** — stewards can overwrite dates with downstream deadline recalculation

### Changed
- Meeting Check-In Log expanded from 13 to 16 columns
- Meeting setup dialog updated with steward selection checkboxes

## [4.5.1] - 2026-02-11

### Fixed
- **Engagement tracking** — resolved 6 undefined column references causing incorrect dashboard data
- **Version consistency** — synced API_VERSION and COMMAND_CONFIG.VERSION
- Added missing `GRIEVANCE_OUTCOMES` constant and `generateGrievanceId()` function
- Sheet tab colors added to all 11 sheet creation functions

### Added
- 79 new engagement tracking tests, 33 new grievance mutation tests
- Total: 950 tests across 18 suites

## [4.5.0] - 2026-02-01

### Added
- Security module — XSS prevention, HTML escaping, formula injection protection, PII masking, input validation, access control
- Data Access Layer — centralized sheet access with caching and deadline management
- Member Self-Service — PIN-based authentication with secure UUID generation and hashed storage
- Jest unit test suite with GAS environment mocking

### Changed
- Consolidated architecture from scattered modules to 16 focused source files
- CI/CD pipeline with GitHub Actions, ESLint v9, Husky pre-commit hooks

### Fixed
- 8 bug fixes including escapeForFormula, generateMemberPIN, TIME_CONSTANTS, and missing configs

### Removed
- 138+ duplicate function definitions and deprecated stubs

## [4.4.1] - 2026-01-31

### Added
- Initial build system with Node.js and source file concatenation

## [4.4.0] - 2026-01-30

### Added
- Member Dashboard with executive overview
- Steward Dashboard with case management
- Grievance tracking and management
- Satisfaction survey integration
- Calendar integration for deadlines
- Email notification system

### Security
- Input validation, HTML sanitization, role-based access control

## [4.3.8] - 2026-01-28

### Added
- Satisfaction modal dashboard with interactive charts
- Features Reference sheet — searchable catalog of all dashboard capabilities
- Hidden satisfaction calculation sheet for background formula processing

## [4.3.7] - 2026-01-25

### Added
- Complete rewrite of help guide with real-time search
- Menu reference tab with all menu items and descriptions
- FAQ tab with categorized questions and answers

## [4.3.2] - 2026-01-20

### Changed
- Deprecated visible Dashboard sheet — replaced with SPA-style modal dashboards
- All dashboard views now load as in-app dialogs instead of sheet tabs

## [4.3.0] - 2026-01-15

### Added
- Case Checklist system for grievance step tracking
- Looker data integration for external reporting
- Dynamic field expansion engine for flexible sheet schemas

## [4.1.0] - 2026-01-10

### Added
- Strategic Command Center configuration sheet
- Status color mapping for visual workflow indicators
- PDF and email branding templates

## [4.0.0] - 2026-01-05

### Added
- **Unified master engine** — single entry point for all dashboard operations
- Audit logging with tamper-evident chain
- Sabotage protection for critical sheets
- Batch processing for bulk member operations
- Mobile-responsive views for all web app pages

## [3.6.0] - 2025-12-20

### Changed
- Member and Grievance data manager refactor
- Improved validation across all data entry points

## [2.0.0] - 2025-11-15

### Changed
- **Modular architecture** — split monolith into modular source files
- Build system for concatenating source into deployable bundle
- Separation of UI and business logic layers

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| 4.18.0 | 2026-02-26 | SPA fixes, seed phasing, steward/member view enhancements |
| 4.17.0 | 2026-02-26 | Q&A Forum, Timeline Service, Failsafe Service (backend + sheets) |
| 4.16.0 | 2026-02-26 | Wire 7 unwired sheets to SPA, new pages, expansion tests |
| 4.15.0 | 2026-02-25 | Survey wizard, steward insights, member KPI strip, login UX |
| 4.14.0 | 2026-02-25 | 130 code review findings, 5 new features, engagement sync |
| 4.13.0 | 2026-02-25 | Notification bell/EventBus auto-alerts, individual-file build, WorkloadService SPA module |
| 4.12.2 | 2026-02-25 | SPA web dashboard, SSO + magic link, deep-link routing, hidden sheets |
| 4.12.0 | 2026-02-24 | Notifications system (sheet + API + dual-role page) |
| 4.11.0 | 2026-02-24 | Resources hub, meeting check-in route, design refresh |
| 4.10.0 | 2026-02-23 | Workload Tracker module |
| 4.9.1 | 2026-02-23 | 15 security fixes (XSS, formula injection, URL validation) |
| 4.9.0 | 2026-02-17 | Constant Contact v3 API integration, multi-select dropdowns, auto-discovery columns |
| 4.8.2 | 2026-02-16 | State field added to member contacts |
| 4.8.1 | 2026-02-15 | 5 new contact form fields, unified name-based Member IDs |
| 4.8.0 | 2026-02-15 | Security event alerting, zero-knowledge survey vault |
| 4.7.0 | 2026-02-14 | 40+ code review fixes, 1090 tests |
| 4.6.0 | 2026-02-12 | Meeting doc automation, steward agenda sharing, member Drive folders |
| 4.5.1 | 2026-02-11 | Engagement tracking fixes, 950 tests |
| 4.5.0 | 2026-02-01 | Security module, Data Access Layer, Member Self-Service |
| 4.4.1 | 2026-01-31 | Build system |
| 4.4.0 | 2026-01-30 | Member/Steward dashboards, grievance tracking, notifications |
| 4.3.8 | 2026-01-28 | Satisfaction modal dashboard, Features Reference sheet |
| 4.3.7 | 2026-01-25 | Help guide rewrite with real-time search |
| 4.3.2 | 2026-01-20 | SPA-style modal dashboards replace visible sheet tabs |
| 4.3.0 | 2026-01-15 | Case Checklist, Looker integration, dynamic field expansion |
| 4.1.0 | 2026-01-10 | Strategic Command Center, status colors, PDF/email branding |
| 4.0.0 | 2026-01-05 | Unified master engine, audit logging, batch processing |
| 3.6.0 | 2025-12-20 | Data manager refactor, improved validation |
| 2.0.0 | 2025-11-15 | Modular architecture, build system |
