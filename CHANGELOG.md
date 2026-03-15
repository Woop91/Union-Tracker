# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [4.28.5] - 2026-03-14

### Fixed
- **Gmail scope test false failure** — `test_emailsend_gmailAppAccessible` used `GmailApp.getAliases()` which requires `gmail.readonly` or broader scope, but the project only has `gmail.send`. Replaced with `GmailApp.createDraft()` + immediate trash, which correctly tests the `gmail.send` scope.
- **Tab switching now truly instant on revisit** — Replaced broken DOM snapshot cache (only saved firstChild, node got destroyed by innerHTML='') with a **show/hide pane system**. Each tab renders into its own `page-layout-content` div. On tab switch, the old pane is hidden (display:none) and the cached pane is shown — zero DOM rebuild, zero server calls, zero re-rendering. Tabs that contain forms/mutations (`notifications`, `feedback`, `profile`, `broadcast`, `failsafe`, `testrunner`) always render fresh. Hidden orphan panes from non-cached tabs are garbage-collected automatically. Steward `home`/`cases` share one pane; `weeklyq` normalizes to `polls`.
- **renderPageLayout no longer destroys content on tab switch** — Instead of `innerHTML = ''` + `contentFn()` which destroyed the previous tab's DOM, now creates a new content pane and hides the old one, allowing the pane cache to restore it instantly.

## [4.28.4] - 2026-03-14

### Fixed
- **serverCall() now retries transient errors** — Refactored `serverCall()` to use Proxy-based interception so ALL server calls (including Insights tab's 6 parallel calls) get automatic retry with exponential backoff on HTTP 0/NetworkError. Previously only `DataCache.cachedCall` had retry.
- **Concurrent call throttle** — Added `_throttledServerCall()` queue that limits in-flight `google.script.run` calls to 4 max. Prevents connection exhaustion when Insights (6 calls) or Union Stats (4 calls) fire simultaneously. Queued calls drain automatically as slots free.
- **QA Forum 999-question scan in batch data** — Replaced `QAForum.getQuestions(email, 1, 999, 'recent')` + full-object iteration in `_getStewardBatchData` with new `QAForum.getUnansweredCount()` that counts directly from raw sheet data without building question objects. Saves ~200-400ms per batch call.

## [4.28.3] - 2026-03-14

### Fixed
- **Slow tab switching (4+ sec)** — Added tab DOM cache that preserves rendered content on tab switches; revisiting a tab now restores instantly from cached DOM instead of re-rendering + re-fetching
- **Loading skeleton shown on cache hits** — Members tab (and others via DataCache.has()) now skip the skeleton animation when data is already cached, eliminating 400-500ms of unnecessary visual delay
- **Redundant stats server call** — `dataGetStewardMemberStats` now routed through DataCache (5-min TTL) instead of firing a fresh `serverCall()` on every members tab visit
- **HTTP 0 / NetworkError crashes** — Added `_serverCallWithRetry()` with exponential backoff (1.5s, 3s) for transient network failures; `DataCache.cachedCall` uses retry automatically
- **Poor error messages on network failure** — `serverCall()` default handler now distinguishes transient network errors ("Connection lost") from server errors ("Something went wrong"), with a "Retry" button for network issues
- **Members tab error UX** — `_onMembersFailed` now shows network-aware error message with an inline Retry button (invalidates cache + re-renders) instead of a generic "Try refreshing" message

### Changed
- Architecture test A8 updated to accept `_serverCallWithRetry` as valid failure handler delegation pattern

## [4.28.2] - 2026-03-14

### Fixed
- **dataAssignSteward auth mismatch** — member_view.html called steward-only `dataAssignSteward`; created `dataMemberAssignSteward` using `_resolveCallerEmail` (members can only assign steward to themselves, preventing privilege escalation)

### Added
- **8 new structural guard tests** (G17–G24) in `test/union-stats-guards.test.js`:
  - G17: member_view.html must never call steward-only endpoints (auto-scans all `_requireStewardAuth` functions)
  - G18: Stats renderers must use all non-reserved backend data fields
  - G19: All stats pages must have client-side caching (AppState cache pattern)
  - G20: Sub-tab renderers must separate fetch from content render (enables caching)
  - G21: Member-safe endpoint pairs verified (steward version + member version both exist)
  - G22: Untracked metrics conditionally displayed (no dead zero KPIs)
  - G23: No redundant sheet reads in engagement stats
  - G24: showLoading skeleton type consistency across sub-tabs

## [4.28.1] - 2026-03-14

### Fixed
- **Polls: null sessionToken in 5 client calls** — `wqGetHistory`, `wqSubmitPoolQuestion` (member), `wqClosePoll`, `wqSetStewardQuestion`, `wqSetPollFrequency`, `wqGetHistory` (steward) now pass `SESSION_TOKEN` instead of `null`, fixing broken auth for magic-link users
- **Polls: redundant page variable** in member and steward poll history pagination removed; nav buttons now call `loadPage()` directly
- **POMS tab broken for stewards** — routing guard in `_handleTabNav` was `role === 'member'` only; stewards clicking POMS were silently redirected to dashboard. Now a shared tab for both roles.
- **POMS description incorrect** — "Postal Operations Manual" corrected to "Program Operations Manual System" in both `steward_view.html` and `member_view.html`.
- **Union Stats auth mismatch** — Grievances and Hot Spots sub-tabs were silently broken for non-steward members; created `dataGetMemberGrievanceStats` / `dataGetMemberGrievanceHotSpots` endpoints using `_resolveCallerEmail` (any authenticated member)
- **Engagement KPI** — hidden Resource Downloads card when value is 0 (not currently tracked)
- **Redundant sheet read** — `dataGetEngagementStats` no longer re-reads Member Directory for membership trends
- **showLoading** — workload sub-tab now uses `showLoading(container, 'kpi')` for consistent skeleton
- **Workload Tracker: crash-safe reporting refresh** — `_refreshReportingData()` now writes new data first, then clears stale rows (previously cleared all content before writing, risking empty sheet on timeout)
- **Workload Tracker: atomic rate limiting** — `_checkAndRecordRateLimit()` replaces separate check+record functions, eliminating race condition where concurrent submissions could both pass the limit check
- **Workload Tracker: bar chart dynamic scaling** — Stats tab category averages bar chart now scales to the actual data max instead of hardcoded 100 (values over 100 previously all showed 100% bars)
- **Workload Tracker: sub-category label display** — History and Stats tabs now show human-readable labels (e.g., "Priority Cases") instead of raw keys (e.g., "priority") via `WT_CAT_KEY_LABELS` lookup

### Added
- **Server-side dues-paying gate** on `wqGetActiveQuestions`, `wqSubmitResponse`, `wqSubmitPoolQuestion` — non-paying members blocked server-side (previously client-only)
- **Rate limiting on pool submissions** — max 3 per user per poll period via `POOL_SUBMIT_LIMIT`
- **Confirmation dialog** on steward community poll draw button (prevents accidental replacement)
- **30+ new tests** for `24_WeeklyQuestions`: closePoll, getPoolCount with pending/used filtering, getHistory pagination + pageSize cap, frequency settings validation, option validation (length/blank/duplicate), response dedup, selectRandomPoolQuestion empty/no-pending, all global wrapper existence
- **Membership sub-tab enriched** — now renders Total Members & New (Last 90 Days) KPIs, By Unit chart, By Dues Status doughnut, and New Hires by Month bar chart (previously only showed By Location)
- **Client-side caching** — Union Stats page mirrors steward Insights caching pattern (`AppState.unionStatsCache`, 5-min TTL); revisiting sub-tabs serves cached data instantly
- **Monthly resolved trend** — grievance trend chart now shows both Filed and Resolved lines
- **Workload Tracker: auto-save draft** — Form data is persisted to localStorage on every input change (400ms debounce) and restored on page reload; cleared on successful submission
- **Workload Tracker: last submitted indicator** — Shows "Last submitted: [date]" banner at top of Submit tab, stored in localStorage
- **Tests: 44 new workload tests** — Expanded from 19 to 63 tests covering rate limiting (atomic, per-email, case-insensitive), history shape/filtering/sorting, sub-category JSON parsing, CSV export, structural invariants
- **Tests: G14 deploy guard** — Source-level checks ensuring crash-safe refresh pattern and atomic rate limiting are maintained
- **Tests: G18 SPA integrity** — 10 frontend invariant tests: WT_CAT_KEY_LABELS definition, bar chart dynamic max, auto-save/draft/restore functions, last-submitted indicator, category alignment

### Security
- **Auth check added to `getPOMSReferenceHtml()`** — now verifies `Session.getActiveUser().getEmail()` before serving content.

### Changed
- Frequency button active-state update uses class toggle instead of re-applying inline styles

## [4.28.0] - 2026-03-14

### Added
- **Speed optimizations (OPT-1 through OPT-10)** — webapp tab load reduced from ~3-4s to ~1-1.5s:
  - OPT-1: `dataGetStewardDashboardInit()` — single batch endpoint replaces 3+ sequential server calls on fallback path
  - OPT-2: DataCache wrapping for steward Members tab (5-min TTL, avoids re-fetch on revisit)
  - OPT-4: Layout shell preservation on tab switch — sidebar/nav structure persists, only content area rebuilds
  - OPT-5: Case list pagination (25 items per page with "Show more" button)
  - OPT-6: Members tab `dataGetAllMembers` + `dataGetStewardMemberStats` fired in parallel (saves 0.5-1s)
  - OPT-9: DocumentFragment for all list rendering (case list, member list) — single DOM write
  - OPT-10: 150ms debounce on member filter/sort/search re-renders

### Changed
- OPT-7: Glow-bar CSS animation changed from `infinite` to single iteration (eliminates GPU jank)
- OPT-8: Removed `backdrop-filter: blur()` from `.card-glass` (GPU perf, kept on modals/nav only)

### Fixed
- **Reliability fixes (remaining items from v4.27.2 audit):**
  - `cleanVault()` — `console.error` → `Logger.log`; write-failure recovery after `clearContents()`
  - QA Forum `submitQuestion()` / `submitAnswer()` — notification calls wrapped in try/catch (prevents silent swallow)
  - `doGetWebDashboard()` — config loaded once at top of function (eliminates redundant `ConfigReader.getConfig()` in error path)
  - Org chart / POMS script re-execution wrapped in per-script try/catch
  - PropertiesService quota monitoring in `cleanupExpiredTokens()` (warns at 400KB/500KB)

## [4.27.2] - 2026-03-14

### Fixed
- **Webapp reliability hardening** — 8 fixes to prevent silent failures and blank pages:
  - `Auth.createSessionToken()` — wrapped in try/catch with fallback cookie duration; prevents page load failure when ConfigReader or PropertiesService is unavailable
  - `getUserRole_()` — added null-guard on `getActiveSpreadsheet()`; prevents steward auth from crashing in web app context
  - `dataGetFullProfile()` — returns `{success:false}` instead of raw `null` when member not found; prevents client-side crashes
  - `sendBroadcastMessage()` — added outer try/catch; prevents unhandled throws from member lookup or config reads
  - `_serveAuth()` / `_serveDashboard()` — `getUrl()` wrapped in null-safe helper; prevents `null` webAppUrl in page data
  - `getOrgChartHtml()` / `getPOMSReferenceHtml()` — wrapped in try/catch with graceful fallback HTML
  - `_getMemberBatchData()` / `_getStewardBatchData()` — all sub-calls individually wrapped; one failing section no longer takes down the entire batch

## [4.27.1] - 2026-03-13

### Changed
- **Seed phase consolidation** — merged Phases 1+2 into single Phase 1 (Config + 500 members + 300 grievances + script owner); member count reduced from 1,000 to 500
- Old Phase 3 (ancillary data) renumbered to Phase 2
- `SEED_SAMPLE_DATA` orchestrator and confirmation dialogs updated for 3-phase layout

### Added
- **New Phase 3: Webapp extras** — seeds SPA features not covered by standard seed:
  - `seedMemberTasksData()` — 8 sample member tasks (12-col schema, `Assignee Type = member`, `MT_SEED_` prefix)
  - `seedCaseChecklistData()` — checklist items for up to 5 seeded grievances using `CHECKLIST_TEMPLATES`, with partial completion for realism
  - `seedSurveyQuestionsData()` — delegates to existing `createSurveyQuestionsSheet()` (non-destructive on re-run)
- Nuke cleanup for new seed data: member tasks (`MT_SEED_` prefix) and case checklist items (`CL-SEED-` prefix) removed during `NUKE_SEEDED_DATA`

## [4.27.0] - 2026-03-13

### Added
- **Dev-Only Quick Deploy Menu** (`src/DevMenu.gs`) — consolidated "⚡ Dev Tools" top-level menu with 27 wrapped actions across 4 groups (Initialize, Refresh, Triggers, Scheduled) plus 4 master "run all" buttons with per-item error isolation and summary alerts
- Build gating: `DevMenu.gs` added to `PROD_EXCLUDE` in `build.js`; `onOpen()` uses `typeof buildDevMenu === 'function'` guard so production builds are unaffected

## [4.26.0] - 2026-03-13

### Fixed
- TestRunner controls card contrast — accent border, card background, stronger shadow for high visibility
- Workload Archive sheet now detected by System Diagnostics (removed from skipKeys in DIAGNOSE_SETUP)
- Repair Dashboard creates missing Workload Archive via existing setupHiddenSheets() path

### Changed
- Removed deprecated "Case Analytics" menu item (called deprecated showInteractiveDashboardTab → showStewardDashboard)

### Added
- **Quick Setup (All Init/Sync)** consolidated admin menu — groups all initialize, trigger, sync, refresh, and setup functions in one place for easy onboarding and system repair

## [4.25.15] - 2026-03-13

### Fixed
- TestRunner controls card contrast in dark mode — replaced `--raised`/`--border` with `--surface`/`--muted`, increased shadow opacity

### Changed
- Survey Tracking default scope now 'location' (My Location) instead of 'assigned' (My Members) for proximity-first workflow
- "Other Pending" survey section now collapsible (matches Completed section pattern)
- Removed duplicate "Quick Setup & Sync" admin menu (~35 lines) — all items already exist in dedicated submenus
- Removed duplicate "Force Global Refresh" from Automation menu (available in Appearance > Refresh All Visuals)
- Survey member lists use Prev/Next pagination instead of "Show all" buttons

### Added
- "New Completed" stat card in survey participation stats — shows new member survey completion as fraction with color coding
- "Declining" stat card now shows average participation rate of declining members
- Proximity badge pills on nearby survey members: "Same Location", "Day Overlap", "Same Floor"
- `_surveyPaginate()` helper for paginated survey member lists (20 per page)
- `withdrawnCount` in `getGrievanceStats()` server-side — fixes missing count on Insights drill-down
- `newMembersLast90` and `byHireMonth` in `getMembershipStats()` — enables new hire trend tracking
- "OVERDUE CASES" insight card with due-this-week sub-text
- "NEW MEMBERS" insight card (last 90 days) with monthly hire trend drill-down
- Overtime average and employment mix (FT/PT ratio) cards in Workload Insights
- Hash-based navigation for Insights detail views — browser back button returns to detail panel
- `scrollIntoView` on insight detail open for mobile UX
- Tests: `withdrawnCount` in grievance stats return shape contract, `getMembershipStats` return shape tests

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
