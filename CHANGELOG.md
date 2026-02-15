# Changelog

All notable changes to the 509 Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.8.0] - 2026-02-15

### Security
- **Zero-knowledge survey vault** — all survey verification data (email, member ID) stored as SHA-256 hashes only
- No plaintext PII is ever written to any sheet — raw email exists in memory only during form submission
- `_Survey_Vault` sheet is hidden and sheet-protected (script owner only)
- Satisfaction sheet contains zero identifying data — cryptographically impossible to link answers to members
- Flagged submissions review now shows "Anonymous submission #N" instead of email
- `_Looker_Satisfaction` export no longer contains member IDs or verification status

### Added
- `hashForVault_(value)` — salted SHA-256 hash function for vault storage
- `SURVEY_VAULT_COLS` constants for 8-column vault structure
- Survey Completion Tracker with dialog, reminders, and round management
- FAQ section for survey completion tracking (7 Q&As)
- Features Reference entries for survey tracking (6 entries)
- Function Checklist entries for vault functions

### Changed
- `onSatisfactionFormSubmit()` — writes anonymous answers to Satisfaction sheet, hashed PII to vault
- `approveFlaggedSubmission()` / `rejectFlaggedSubmission()` — now write to vault, not Satisfaction sheet
- `syncSatisfactionValues()`, `getPublicSurveyData()`, `getSatisfactionByUnit()` — read verification status from vault via `getVaultDataMap_()`
- `SATISFACTION_COLS.EMAIL` through `REVIEWER_NOTES` deprecated (set to -1)
- Verification columns (CE-CK) removed from Satisfaction sheet entirely

## [4.7.0] - 2026-02-14

### Fixed
- **40+ code review issues resolved** across security, correctness, performance, and test quality
- **XSS hardening** — all Critical security issues from comprehensive code review addressed
- **onEdit optimization** — reduced redundant processing in edit trigger handlers
- **DevTools guard** — added production environment check to prevent accidental DevTools execution
- **Deduplicated `escapeHtml`** — consolidated duplicate implementations into single canonical function in `00_Security.gs`
- **Removed empty stubs** — cleaned up placeholder functions that had no implementation (A3 architecture item)
- **Broken `GRIEVANCE_COLS.UNIT` and `MEMBER_COLS.HOME_TOWN` references** — removed 9 references to non-existent column mappings
- **Data safety improvements** — 14 critical corrections for data integrity, UX, and reliability
- **Dashboard footer version** — updated from v4.4.0 to v4.7.0
- **Unused variable lint warning** — removed unused `hiddenSheetNames` in `08d_AuditAndFormulas.gs`

### Changed
- Version bumped to 4.7.0 across all files (VERSION_INFO, package.json, API_VERSION)
- Test suite expanded from 1016 to 1090 tests across 20 suites (was 19)
- 74 new architecture and happy-path tests added
- CODE_REVIEW.md updated with resolution status for all 69 identified issues
- Updated README.md, SECURITY_REVIEW.md version references to 4.7.0

## [4.6.0] - 2026-02-12

### Added
- **VERSION_HISTORY constant** (`01_Core.gs`) - Centralized array tracking every release with version number, date, codename, and key changes. Includes `getVersionDate(ver)` lookup function so any code can resolve a version string to its release date
- **Features sheet now driven by VERSION_HISTORY** (`10b_SurveyDocSheets.gs`) - Replaced hardcoded version table with dynamic rendering from `VERSION_HISTORY`, ensuring dates are always present and consistent
- **getVersionInfo() updated** (`10_Main.gs`) - Now returns current VERSION_INFO fields plus full VERSION_HISTORY instead of stale 2.0.0 data
- **Meeting Notes & Agenda Document Automation** - When a meeting is created, Google Docs for Meeting Notes and Meeting Agenda are auto-generated in dedicated Drive folders (`Meeting Notes/`, `Meeting Agenda/`). URLs stored in Meeting Check-In Log columns N-O
- **Two-Tier Agenda Steward Selection** - Meeting setup dialog lets the organizer select which stewards receive the agenda 3 days before the meeting; ALL stewards receive it at least 1 day before. Agenda is never shared with members
- **Meeting Notes Dashboard Tab** - New "Meeting Notes" tab in Member Dashboard shows completed meetings chronologically with search and view-only Google Doc links. Notes auto-publish (view-only) 1 day after each meeting
- **Scheduled Meeting Document Notifications** - `processMeetingDocNotifications()` runs in `dailyTrigger()`: agenda 3 days before (selected stewards), agenda 1 day before (all stewards), notes 1 day before (notification stewards), view-only publish 1 day after
- **Member Drive Folder Quick Action** - "Create Member Folder" button in Member Quick Actions dialog creates/reuses a Google Drive folder for the member, checking for existing grievance folders first
- **Meeting Event Scheduling** - Full calendar lifecycle for meetings: creates Google Calendar events, activates/deactivates check-in based on event status
- **Grievance Date Override** - Stewards can overwrite grievance dates with downstream deadline recalculation
- **Steward Checkboxes in Meeting Setup** - Dynamic steward list loaded from Member Directory with Select All / Clear All, separate "Email Attendance Report To" and "Send Agenda Early To" sections
- Meeting Check-In Log expanded from 13 columns (A-M) to 16 columns (A-P): `NOTES_DOC_URL` (N), `AGENDA_DOC_URL` (O), `AGENDA_STEWARDS` (P)

### Changed
- Version bumped to 4.6.0 across all files (VERSION_INFO, COMMAND_CONFIG, API_VERSION)
- Meeting setup dialog height increased from 580px to 720px to accommodate steward selection
- `dailyTrigger()` now includes `processMeetingDocNotifications()` with audit logging for agenda sent, notes sent, and notes published counts

### New Functions
- `createMeetingDocs(meetingData)` - Creates Meeting Notes and Agenda Google Docs in dedicated folders
- `getOrCreateMeetingNotesFolder()` / `getOrCreateMeetingAgendaFolder()` - Gets or creates Drive folders for meeting documents
- `emailMeetingDocLink()` - Sends HTML email with meeting document link to stewards
- `setDocViewOnlyByLink(docUrl)` - Sets a Google Doc to view-only sharing via link
- `getAllStewardEmails_()` - Retrieves all steward emails from Member Directory
- `processMeetingDocNotifications()` - Two-tier notification scheduler for meeting documents
- `setupDriveFolderForMember(memberId)` - Creates/reuses Google Drive folder for a member
- `getStewardEmailsForMeetingSetup()` - Returns steward names/emails for the meeting setup dialog

## [4.5.1] - 2026-02-11

### Fixed
- **Engagement tracking: `MEMBER_COLS.FULL_NAME` undefined** - Member names in unified dashboard always showed "Unknown". Now builds name from `FIRST_NAME` + `LAST_NAME`
- **Engagement tracking: `MEMBER_COLS.LAST_UPDATED` undefined** - Directory trends (recent updates, stale contacts) never tracked. Now uses `RECENT_CONTACT_DATE`
- **Engagement tracking: `GRIEVANCE_COLS.CATEGORY` undefined** - Grievance categories always showed "Other". Now uses `ISSUE_CATEGORY`
- **Engagement tracking: `GRIEVANCE_COLS.MEMBER_NAME` undefined** - Grievance member names always showed "Unknown". Now builds from `FIRST_NAME` + `LAST_NAME`
- **Engagement tracking: step denial rates always 0%** - `STEP_1_DATE`/`STEP_2_DATE`/`STEP_3_DATE` were undefined. Now uses `STEP1_RCVD`, `STEP2_RCVD`, `STEP2_APPEAL_FILED`, `STEP3_APPEAL_FILED`
- **Engagement tracking: openRate=0 excluded** - Members with 0% email open rate were skipped because `0` is falsy in JS. Now uses explicit empty checks
- **Interest field comparison** - Added lowercase `'true'` to union interest value checks
- **Version consistency** - Synced `API_VERSION` (4.5.0→4.5.1) and `COMMAND_CONFIG.VERSION` (4.5.0→4.5.1) to match `VERSION_INFO.CURRENT`
- `GRIEVANCE_OUTCOMES` - added missing constant that caused "GRIEVANCE_OUTCOMES is not defined" runtime error, preventing tabs from populating
- `generateGrievanceId()` - added missing function called by `getNextGrievanceId()`, which caused `startNewGrievance()` to silently fail
- Sheet tab colors - added `.setTabColor()` to all 11 sheet creation functions
- Hardcoded hex colors replaced with `COLORS` constants in Getting Started, FAQ, and Config Guide sheet creation

### Added
- 79 new engagement tracking tests (`test/04e_PublicDashboard.test.js`) covering: member counting, engagement metrics, participation rates, hot spot detection, PII handling, grievance processing, directory trends, satisfaction data, date range filtering, HTML generation, edge cases
- Engagement Tracking FAQ category (7 questions) in FAQ sheet
- Total: 950 tests across 18 suites (up from 871 across 17)
- 33 new tests covering core grievance mutation paths:
  - `startNewGrievance()` - validates row data includes `GRIEVANCE_OUTCOMES.PENDING`, audit logging, error handling
  - `resolveGrievance()` - validates outcome/status updates, audit logging, timestamped notes
  - `advanceGrievanceStep()` - validates step 1→2→3→arbitration transitions, status updates, boundary errors
  - `setupCalcFormulasSheet()` - validates `GRIEVANCE_OUTCOMES` and `GRIEVANCE_STATUS` are written to formula sheet
  - `GRIEVANCE_OUTCOMES` constant existence guard - asserts all expected keys are defined

## [4.5.0] - 2026-02-01

### Added
- Security module (`00_Security.gs`) - XSS prevention, HTML escaping, formula injection protection, PII masking, input validation, access control
- Data Access Layer (`00_DataAccess.gs`) - Centralized sheet access with caching, TIME_CONSTANTS for deadline management, deadline urgency calculations
- Member Self-Service (`13_MemberSelfService.gs`) - PIN-based member authentication with secure UUID-based PIN generation and hashed storage
- Jest unit test suite with GAS environment mocking
- jest.config.js and test infrastructure (gas-mock.js, load-source.js)

### Changed
- Consolidated architecture from scattered modules to 16 focused source files
- CI/CD pipeline with GitHub Actions
- ESLint v9 flat config for code quality
- Husky pre-commit hooks
- Version bumped to 4.5.0 across all files (VERSION_INFO, COMMAND_CONFIG, API_VERSION, package.json)

### Fixed
- `escapeForFormula()` - was prefixing formula chars mid-string, now only at start
- `generateMemberPIN()` - replaced insecure Math.random with Utilities.getUuid()
- `TIME_CONSTANTS.DEADLINE_DAYS` - corrected values to match DEADLINE_RULES (7/7/14/10)
- `API_VERSION` - synced with VERSION_INFO (4.5.0)
- `CALENDAR_CONFIG` - re-added missing config to Integrations module
- `MEMBER_COLUMNS.PIN_HASH` - added missing column constant and header
- `initializeDashboard()` - fixed dead stub, now delegates to CREATE_509_DASHBOARD()
- `COMMAND_CONFIG.VERSION` - synced with VERSION_INFO.CURRENT

### Removed
- 138+ duplicate function definitions
- Deprecated function stubs

## [4.4.1] - 2026-01-31

### Added
- Initial build system with Node.js
- Source file concatenation for deployment
- Basic project structure

## [4.4.0] - 2026-01-30

### Added
- Member Dashboard with executive overview
- Steward Dashboard with case management
- Grievance tracking and management
- Satisfaction survey integration
- Calendar integration for deadlines
- Email notification system

### Security
- Input validation for all user inputs
- HTML sanitization for XSS prevention
- Role-based access control

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| 4.6.0 | 2026-02-12 | Meeting Notes & Agenda doc automation, two-tier steward agenda sharing, Meeting Notes dashboard tab, member Drive folders, meeting event scheduling |
| 4.5.1 | 2026-02-11 | Engagement tracking fixes, 950 Jest tests, GRIEVANCE_OUTCOMES/generateGrievanceId fixes |
| 4.5.0 | 2026-02-01 | Security module, Data Access Layer, Member Self-Service, consolidated to 16 source files |
| 4.4.1 | 2026-01-31 | Build system |
| 4.4.0 | 2026-01-30 | Initial release |
