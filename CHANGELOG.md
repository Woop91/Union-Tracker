# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [4.25.13] - 2026-03-13

### Fixed
- XSS vulnerability in serverCall error handler — error messages now use textContent instead of raw innerHTML
- `.spinner` class was `display:none` — now renders as a visible rotating ring for auth/survey submit states
- Duplicate `.skeleton-card` CSS definitions consolidated — C1 skeleton scoped under `.skeleton-wrap`

### Added
- Cohesive loading indicator system with contextual variants: `showLoading(container, 'list'|'form'|'kpi')`
- Light-mode support for skeleton-pulse loading indicators (`.skeleton-row`, `.skeleton-card`)
- 30-second loading timeout fallback prevents infinite spinners with reload prompt
- `clearLoading()` helper to cancel loading timeouts when content arrives
- `prefers-reduced-motion` support for skeleton and spinner animations
- `.loading-timeout` CSS class for timeout error state styling

## [4.25.12] - 2026-03-12

### Changed
- Loading spinner replaced with pulsing dots animation for a cleaner loading state

## [4.25.11] - 2026-03-12

### Added
- 9 new GAS-native web app test suites in `31_WebAppTests.gs` — independently triggerable via suite filter, each completes under 3 minutes
  - `webapp_` — doGet routing, template rendering, URL resolution, diagnoseWebApp health check
  - `configrd_` — ConfigReader module completeness, validation, JSON output, drive/auth fields
  - `portal_` — PortalSheets 0-indexed column constants validation, no duplicate indices, setup functions
  - `weeklyq_` — WeeklyQuestions module API, Q_COLS exposure, poll frequency/pool count reads
  - `workload_` — WorkloadService module, SUB_CATEGORIES/CATEGORY_LABELS exposure, sub-categories read
  - `qaforum_` — QAForum module API, question retrieval, pagination defaults, flagged content gating
  - `timeline_` — TimelineService module, event retrieval, category validation, write auth gating
  - `failsafe_` — FailsafeService module, digest config shape, diagnostic/ensureAllSheets existence
  - `endpoints_` — Comprehensive data/wq/qa/tl/fs wrapper existence checks + write endpoint null-token rejection
- Total GAS-native tests: ~170 across 20 suites (up from ~92 across 11 suites)
- `seedQAForumData()` — seeds 10 realistic Q&A Forum questions with 15 answers. Added to `SEED_PHASE_3`
- Individual member participation progress bars in Survey Tracking (snapped to 5% chunks, color-coded)
- Skeleton placeholder UI for loading state (pulsing cards/rows)

### Changed
- Member seed data now assigns 2-3 random office days (comma-separated multi-select)

### Fixed
- **Chrome cache preventing web app from loading** — Added `Cache-Control: no-cache, no-store, must-revalidate`, `Pragma: no-cache`, and `Expires: 0` meta tags to `index.html`. Chrome was aggressively caching the page, serving stale session state in regular browsing mode (worked fine in incognito where cache/localStorage start empty).
- **Stale localStorage token redirect loop** — Added redirect-loop guard (max 2 attempts) to prevent infinite redirect when an expired session token persists in `localStorage`. After 2 failed validation attempts the stale token is automatically cleared.
- **Magic link "Failed to send email" error too generic** — Added step-level tracking (`cache`, `lookup`, `config`, `token`, `url`, `build-email`, `send`) to `sendMagicLink()` outer catch. Server logs now show exactly which step failed; client receives actionable error messages instead of the generic fallback.

## [4.25.10] - 2026-03-11

### Added
- `diagnoseWebApp()` — 14-step diagnostic function in `22_WebDashApp.gs` for debugging app loading issues (checks deployment, permissions, doGet, view rendering, config, auth)
- `memberId` included in profile data (`21_WebDashDataService.gs`)
- Workload sheet diagnostics in `28_FailsafeService.gs`

### Changed
- **Parallel view rendering** in `index.html` — steward and member views now load concurrently for faster startup
- Nav tab reorder: Feedback moved to Admin section; `stewarddirectory` removed from member sidebar
- Enhanced error reporting: null-guard on `fatalErr.stack`, actual error message displayed in bootstrap screen instead of generic failure

## [4.25.7] - 2026-03-10

### Fixed
- **Critical: `onOpen` broken deferred trigger** (`10_Main.gs`) — Two compounding bugs prevented `onOpenDeferred_` from ever running on sheet open:
  1. `ScriptApp.getProjectTriggers()` is not permitted in GAS simple triggers — throws silently, falling through to inline `onOpenDeferred_()` call which also failed for the same reason
  2. `finally` block called `cleanUpOnOpenTrigger_()` synchronously, deleting the 1-second deferred trigger before it could fire
- `onOpen` now does only menu creation + cache clear (correct GAS simple trigger pattern)

### Added
- `setupOpenDeferredTrigger()` (`08e_SurveyEngine.gs`) — installs `onOpenDeferred_` as a proper installable `onOpen` trigger via `ScriptApp.newTrigger().forSpreadsheet().onOpen()`
- Menu item **🔓 Install onOpen Deferred Trigger** under Admin → ⏱️ Triggers
- `menuInstallSurveyTriggers()` now also calls `setupOpenDeferredTrigger()` — running "Install ALL" covers this fix



### Fixed
- Test runner GAS 6-minute execution timeout crash when running all 82 tests
- Timeout guard lowered from 5min to 3.5min (2.5min safety margin vs 1min before)
- Cached spreadsheet reference (`_getCachedSS()`) eliminates 12 redundant `SpreadsheetApp.getActiveSpreadsheet()` network round-trips in test functions

### Added
- Timeout warning banner in SPA test runner UI with guidance to use suite filter
- Skipped test count summary card
- Suite headers now show skipped count
- Failure handler detects "execution time" errors and shows helpful message instead of raw GAS error

## [4.25.5] - 2026-03-09

### Removed
- "Directory" tab (`id: contact`) from member SPA sidebar — members should not have a general directory
- `case 'contact'` routing entry and `contact` color mapping in `index.html`

### Changed
- Deep-links to `#contact` for members now fall through to Home (default)
- Updated README.md, FEATURES.md member tab listing, INTERACTIVE_DASHBOARD_GUIDE.md tab count (10→9)

### Note
- Steward Directory utility link (`id: stewarddirectory`) retained — used for finding/selecting an assigned steward

## [4.25.4] - 2026-03-09

### Added
- Config columns: Step III Appeal Days, Step III Response Days, Arbitration Demand Days
- getDeadlineRules() now reads Step III + Arbitration from Config (was hardcoded)
- New test: deadline config column completeness (7 assertions)

### Fixed
- COMMAND_CONFIG.VERSION stuck at "4.24.4" — corrected to "4.25.3"
- Test positions for CHIEF_STEWARD_EMAIL (+3), ESCALATION_STATUSES (+3), ESCALATION_STEPS (+3)

## [4.25.2] - 2026-03-09

### Added
- Email notifications on scheduled test failures (daily trigger only)
- New Config column: `Test Runner Notify Email` — set to any email to receive failure alerts
- Styled HTML email body (dark theme) with plain-text fallback
- MailApp quota guard (skips if remaining < 5)

### Changed
- `runScheduledTests()` now checks results and calls `_sendTestFailureEmail()` on failure
- Manual runs (menu/SPA) still use toast only — no email spam

## [4.25.1] - 2026-03-09

### Added
- `dataservice` suite (10 tests): DataService CRUD — module existence, public API completeness, findUserByEmail shape, invalid email handling
- `authsweep` suite (6 tests): endpoint auth rejection — steward/member endpoints reject null token, no data leaks, poll stubs safe, test runner endpoints gated
- `configlive` suite (8 tests): live Config completeness — sheet headers exist, column constants don't exceed sheet width, Config row 3 populated
- `survey` suite (10 tests): survey engine integrity — hidden sheets, period/question cols, getSurveyQuestions shape, period management, tracking sheet
- SPA suite filter dropdown updated with 4 new options
- Total: 10 suites, 82 GAS-native tests

## [4.25.0] - 2026-03-09

### Added
- GAS-native test runner framework (`30_TestRunner.gs`) — runs inside Apps Script runtime
- 6 test suites: config, colmap, auth, grievance, security, system (48 tests)
- SPA dashboard panel (steward-only `testrunner` tab) with run-all, per-suite filter, expandable results
- Sheets menu item: 🛠️ Admin → 🧪 Test Runner
- Daily trigger support (6 AM scheduled runs)
- Server endpoints: `dataRunTests`, `dataGetTestResults`, `dataManageTestTrigger` (steward-auth gated)
- Results stored in ScriptProperties for async SPA retrieval

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

## [4.20.18] - 2026-02-27

### Added
- Per-member master admin folder architecture (v4.20.25)
- Auto-migrate missing columns via `_addMissingGrievanceHeaders_` (v4.20.26)
- Auto-migrate missing columns via `_addMissingMemberHeaders_` (v4.20.26)
- Auto-share Minutes folder on creation; auto-migrate DriveDocUrl header (v4.20.18)

### Changed
- One-time migration for Contact Log Folder URL to Member Admin Folder URL (v4.20.26)

### Fixed
- Minutes: schema fix, pagination, date pre-fill, auto-refresh, folder warning, grievance CTA (v4.20.18)
- Minutes: fix 3 bugs in member minutes CTA + improve SETUP_DRIVE_FOLDERS UX (v4.20.18a)
- `dataSendDirectMessage` Drive log broken by folder return shape change (v4.20.25)

## [4.13.0] - 2026-02-25

See git log for detailed changes prior to this changelog.
### Added
- 9 new GAS-native web app test suites in `31_WebAppTests.gs` — independently triggerable via suite filter, each completes under 3 minutes
  - `webapp_` — doGet routing, template rendering, URL resolution, diagnoseWebApp health check
  - `configrd_` — ConfigReader module completeness, validation, JSON output, drive/auth fields
  - `portal_` — PortalSheets 0-indexed column constants validation, no duplicate indices, setup functions
  - `weeklyq_` — WeeklyQuestions module API, Q_COLS exposure, poll frequency/pool count reads
  - `workload_` — WorkloadService module, SUB_CATEGORIES/CATEGORY_LABELS exposure, sub-categories read
  - `qaforum_` — QAForum module API, question retrieval, pagination defaults, flagged content gating
  - `timeline_` — TimelineService module, event retrieval, category validation, write auth gating
  - `failsafe_` — FailsafeService module, digest config shape, diagnostic/ensureAllSheets existence
  - `endpoints_` — Comprehensive data/wq/qa/tl/fs wrapper existence checks + write endpoint null-token rejection
- Total GAS-native tests: ~170 across 20 suites (up from ~92 across 11 suites)
- `seedQAForumData()` — seeds 10 realistic Q&A Forum questions with 15 answers. Added to `SEED_PHASE_3`
- Individual member participation progress bars in Survey Tracking (snapped to 5% chunks, color-coded)
- Skeleton placeholder UI for loading state (pulsing cards/rows)

### Changed
- Member seed data now assigns 2-3 random office days (comma-separated multi-select)

### Fixed
- **Chrome cache preventing web app from loading** — Added `Cache-Control: no-cache, no-store, must-revalidate`, `Pragma: no-cache`, and `Expires: 0` meta tags to `index.html`. Chrome was aggressively caching the page, serving stale session state in regular browsing mode (worked fine in incognito where cache/localStorage start empty).
- **Stale localStorage token redirect loop** — Added redirect-loop guard (max 2 attempts) to prevent infinite redirect when an expired session token persists in `localStorage`. After 2 failed validation attempts the stale token is automatically cleared.
- **Magic link "Failed to send email" error too generic** — Added step-level tracking (`cache`, `lookup`, `config`, `token`, `url`, `build-email`, `send`) to `sendMagicLink()` outer catch. Server logs now show exactly which step failed; client receives actionable error messages instead of the generic fallback.

## [4.25.10] - 2026-03-11

### Added
- `diagnoseWebApp()` — 14-step diagnostic function in `22_WebDashApp.gs` for debugging app loading issues (checks deployment, permissions, doGet, view rendering, config, auth)
- `memberId` included in profile data (`21_WebDashDataService.gs`)
- Workload sheet diagnostics in `28_FailsafeService.gs`

### Changed
- **Parallel view rendering** in `index.html` — steward and member views now load concurrently for faster startup
- Nav tab reorder: Feedback moved to Admin section; `stewarddirectory` removed from member sidebar
- Enhanced error reporting: null-guard on `fatalErr.stack`, actual error message displayed in bootstrap screen instead of generic failure

## [4.25.7] - 2026-03-10

### Fixed
- **Critical: `onOpen` broken deferred trigger** (`10_Main.gs`) — Two compounding bugs prevented `onOpenDeferred_` from ever running on sheet open:
  1. `ScriptApp.getProjectTriggers()` is not permitted in GAS simple triggers — throws silently, falling through to inline `onOpenDeferred_()` call which also failed for the same reason
  2. `finally` block called `cleanUpOnOpenTrigger_()` synchronously, deleting the 1-second deferred trigger before it could fire
- `onOpen` now does only menu creation + cache clear (correct GAS simple trigger pattern)

### Added
- `setupOpenDeferredTrigger()` (`08e_SurveyEngine.gs`) — installs `onOpenDeferred_` as a proper installable `onOpen` trigger via `ScriptApp.newTrigger().forSpreadsheet().onOpen()`
- Menu item **🔓 Install onOpen Deferred Trigger** under Admin → ⏱️ Triggers
- `menuInstallSurveyTriggers()` now also calls `setupOpenDeferredTrigger()` — running "Install ALL" covers this fix



### Fixed
- Test runner GAS 6-minute execution timeout crash when running all 82 tests
- Timeout guard lowered from 5min to 3.5min (2.5min safety margin vs 1min before)
- Cached spreadsheet reference (`_getCachedSS()`) eliminates 12 redundant `SpreadsheetApp.getActiveSpreadsheet()` network round-trips in test functions

### Added
- Timeout warning banner in SPA test runner UI with guidance to use suite filter
- Skipped test count summary card
- Suite headers now show skipped count
- Failure handler detects "execution time" errors and shows helpful message instead of raw GAS error

## [4.25.5] - 2026-03-09

### Removed
- "Directory" tab (`id: contact`) from member SPA sidebar — members should not have a general directory
- `case 'contact'` routing entry and `contact` color mapping in `index.html`

### Changed
- Deep-links to `#contact` for members now fall through to Home (default)
- Updated README.md, FEATURES.md member tab listing, INTERACTIVE_DASHBOARD_GUIDE.md tab count (10→9)

### Note
- Steward Directory utility link (`id: stewarddirectory`) retained — used for finding/selecting an assigned steward

## [4.25.4] - 2026-03-09

### Added
- Config columns: Step III Appeal Days, Step III Response Days, Arbitration Demand Days
- getDeadlineRules() now reads Step III + Arbitration from Config (was hardcoded)
- New test: deadline config column completeness (7 assertions)

### Fixed
- COMMAND_CONFIG.VERSION stuck at "4.24.4" — corrected to "4.25.3"
- Test positions for CHIEF_STEWARD_EMAIL (+3), ESCALATION_STATUSES (+3), ESCALATION_STEPS (+3)

## [4.25.2] - 2026-03-09

### Added
- Email notifications on scheduled test failures (daily trigger only)
- New Config column: `Test Runner Notify Email` — set to any email to receive failure alerts
- Styled HTML email body (dark theme) with plain-text fallback
- MailApp quota guard (skips if remaining < 5)

### Changed
- `runScheduledTests()` now checks results and calls `_sendTestFailureEmail()` on failure
- Manual runs (menu/SPA) still use toast only — no email spam

## [4.25.1] - 2026-03-09

### Added
- `dataservice` suite (10 tests): DataService CRUD — module existence, public API completeness, findUserByEmail shape, invalid email handling
- `authsweep` suite (6 tests): endpoint auth rejection — steward/member endpoints reject null token, no data leaks, poll stubs safe, test runner endpoints gated
- `configlive` suite (8 tests): live Config completeness — sheet headers exist, column constants don't exceed sheet width, Config row 3 populated
- `survey` suite (10 tests): survey engine integrity — hidden sheets, period/question cols, getSurveyQuestions shape, period management, tracking sheet
- SPA suite filter dropdown updated with 4 new options
- Total: 10 suites, 82 GAS-native tests

## [4.25.0] - 2026-03-09

### Added
- GAS-native test runner framework (`30_TestRunner.gs`) — runs inside Apps Script runtime
- 6 test suites: config, colmap, auth, grievance, security, system (48 tests)
- SPA dashboard panel (steward-only `testrunner` tab) with run-all, per-suite filter, expandable results
- Sheets menu item: 🛠️ Admin → 🧪 Test Runner
- Daily trigger support (6 AM scheduled runs)
- Server endpoints: `dataRunTests`, `dataGetTestResults`, `dataManageTestTrigger` (steward-auth gated)
- Results stored in ScriptProperties for async SPA retrieval

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

## [4.20.18] - 2026-02-27

### Added
- Per-member master admin folder architecture (v4.20.25)
- Auto-migrate missing columns via `_addMissingGrievanceHeaders_` (v4.20.26)
- Auto-migrate missing columns via `_addMissingMemberHeaders_` (v4.20.26)
- Auto-share Minutes folder on creation; auto-migrate DriveDocUrl header (v4.20.18)

### Changed
- One-time migration for Contact Log Folder URL to Member Admin Folder URL (v4.20.26)

### Fixed
- Minutes: schema fix, pagination, date pre-fill, auto-refresh, folder warning, grievance CTA (v4.20.18)
- Minutes: fix 3 bugs in member minutes CTA + improve SETUP_DRIVE_FOLDERS UX (v4.20.18a)
- `dataSendDirectMessage` Drive log broken by folder return shape change (v4.20.25)

## [4.13.0] - 2026-02-25

See git log for detailed changes prior to this changelog.
