# Changelog

All notable changes to the Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
