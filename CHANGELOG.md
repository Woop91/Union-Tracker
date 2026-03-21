# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [4.32.1] - 2026-03-20

### Added
- **Resource click tracking** — New `_Resource_Click_Log` hidden sheet and `dataLogResourceClick()` server function. Resource card expands in member view now fire a tracking event (once per card per session). The `resourceDownloads` metric in `dataGetEngagementStats` is now live, reading total clicks from the log instead of returning hardcoded 0. Engagement stats KPI grid always shows the "Resource Views" card.
- **Grievance Drive folder URL wiring** — `getMemberGrievanceDriveUrl()` now reads `DRIVE_FOLDER_URL` from the Grievance Log via the data record instead of returning `null`. Added `grievanceDriveFolderUrl` to the HEADERS lookup map and `driveFolderUrl` to `_buildGrievanceRecord()`, which also fixes `dataGetMemberCaseFolderUrl()`.

### Removed
- **Dead code: Looker integration** — Removed ~1,100 lines of orphaned Looker integration code from `12_Features.gs` (16 functions, never called).
- **Dead code: deprecated stubs** — Removed 5 deprecated redirect stubs from `04d_ExecutiveDashboard.gs`.
- **Dead code: deprecated configs** — Removed `SHEET_NAMES` alias, `COMMAND_CENTER_CONFIG`, `GEMINI_CONFIG`, `rebuildDashboard()`.
- **Hardcoded fallbacks** — Removed unnecessary `typeof SHEETS` guards in `21_WebDashDataService.gs`.

### Fixed
- **BUG-TASKS-03: Steward task completion** — `createMemberTask` was writing `memberEmail` to column 2 instead of `stewardEmail`.
- **Tab stacking: panels showing side-by-side** — Added `_hideAllVisiblePanes()` helper.
- **Switch to Member: dead sidebar button** — Resets cached state before view switch.
- **Comic theme: bold text letter-spacing** — Increased from `0.5px` to `1.5px`.
- **getMyFeedback: dead FEEDBACK_COLS.TYPE reference** — Now aliases to CATEGORY.
- **Mobile: Missing theme options** — Added dark/light toggle and color picker to mobile More menu.
- **Null guards: getSheetByName()** — Added guards to ~40 unguarded calls across 16 files.
- **Race conditions: LockService** — Wrapped 5 unprotected data-write functions in `withScriptLock_()`.
- **getLastRow() safety** — Added empty-sheet guards.
- **Hardcoded fallback removal** — Direct constant references in `12_Features.gs`.

### Changed
- **dataToggleChecklistItem: added withScriptLock_** — Lock-protected checklist toggle.
- **Badge refresh on writes** — `dataUpdateTask`, `dataToggleChecklistItem`, `dataSubmitFeedback` call `_refreshNavBadges()`.
- **Tests updated** — Updated test suites for deprecated code removal and new behavior.

## [4.30.2] - 2026-03-17

### Fixed
- **Mobile: Horizontal wiggle** — Replaced FAB `100vw` calc with simple `right: 16px`; added `overflow-x: auto` to `.sub-tabs` so Union Stats and Workload sub-tab bars scroll horizontally instead of overflowing the viewport on narrow screens.
- **Workload Tracker: Failed to open** — Added try-catch and `CURRENT_USER` null guard inside `renderWorkloadTracker()` so errors display inline instead of throwing an alert via the More menu catch handler.
- **Org Chart: Cells not expanding** — `initDesktop()` was bound to `DOMContentLoaded` which never fires when org chart HTML is injected via `innerHTML` + script re-execution. Now calls `initDesktop()` immediately when `document.readyState` is not `loading`, plus a safety-net call from `renderOrgChart()`.

### Changed
- **Contact Tab: Steward pagination** — Steward directory in member view now shows 15 stewards per page with Previous/Next navigation. Prioritizes stewards at the same work location and in-office today (existing sort preserved). Reduces DOM weight on large directories.
- **Webapp performance: conditional view inclusion** — Member-only users no longer receive steward_view.html (~300KB), reducing initial page payload by ~35% for members.
- **Webapp performance: production minification** — `build:prod` now enables `--minify` by default, reducing HTML/CSS/JS payload by ~150KB (13-26% per file).
- **Webapp performance: non-blocking font loading** — Google Fonts stylesheet loads asynchronously via `media="print"` pattern, eliminating render-blocking CSS.
- **Webapp performance: spreadsheet singleton** — DataService caches `getActiveSpreadsheet()` per execution, eliminating ~19 redundant IPC calls per server request.
- **Webapp performance: lazy-view HTML caching** — `getOrgChartHtml()` now caches rendered HTML in CacheService (6-hour, version-keyed TTL), avoiding repeated `HtmlService.createHtmlOutputFromFile()` calls.

## [4.30.1] - 2026-03-17

### Added
- **Contact Log: Member Name storage & display** — Contact log now stores the member's full name (column 9) alongside their email. Name is auto-resolved from the Member Directory when selected via autocomplete, or looked up server-side as fallback. Recent Contacts and By Member views display the member name prominently.
- **Contact Log: Improved autocomplete** — Typeahead now triggers after 1 character (previously 2) for faster matching. By Member search tab also gets full autocomplete support for name/email matching.

## [4.29.1] - 2026-03-15

### Fixed
- Test registry consistency check now skips `typeof` guard patterns in `fn:` entries
- `renderBottomNav` test updated to look in `index.html` after v4.26.1 consolidation
- **Split View: secondary panel content stacking** — Switching tabs in the right panel now properly replaces content instead of appending below. Old panes are removed from DOM instead of just hidden, preventing scroll accumulation.
- **Split View: GC safety** — Tab pane garbage collection now scopes to the primary panel only, preventing accidental removal of secondary panel content.

## [4.29.0] - 2026-03-15

### Added
- **Split View (dual-panel tabs)** — On tablet and desktop, users can enable split-view mode from the sidebar to display two tabs side-by-side. Features include: pinnable primary/secondary panels, draggable resize divider (25%-75% range), panel target selector (Left/Right), close button on secondary panel, and persistent split ratio via localStorage. Automatically disabled on mobile. Sidebar shows L/R badges indicating which panel each tab occupies.

## [4.28.8] - 2026-03-15

### Added
- **Unified color theme system** — Sheet themes and webapp accent colors now linked. Each `THEME_PRESETS` entry carries an `accentHue` (and emoji). Selecting a color theme from the webapp sidebar updates both the webapp accent palette instantly and persists the sheet theme server-side via `dataApplyColorTheme()`. On next page load, `_sanitizeConfig()` reads the user's saved `accentHue` from UserProperties to override the default.
- **Color Theme picker in webapp sidebar** — New expandable panel (rainbow icon) lists all 8 color themes with name, hue, and header color swatch. Clicking applies instantly (no reload) and saves to server.
- **`ThemeEngine.updatePalette(hue)`** — New method regenerates accent palette from a hue value and re-applies the current theme. Used by the color theme picker for instant visual feedback.
- **`getUserColorTheme()` / `getColorThemeList()`** — Server functions to retrieve the user's current color theme and list of available themes. Theme list passed in page data for sidebar rendering.
- **`dataApplyColorTheme(sessionToken, themeKey)`** — Auth-gated webapp endpoint for persisting color theme selection.

## [4.28.7] - 2026-03-15

### Fixed
- **OLED dark mode** — Dark mode backgrounds changed to pure black (`#000000`) for both steward and member roles for true OLED compatibility (saves battery, deeper blacks)
- **Dark mode text contrast** — Eliminated all `#000` (black) text in dark mode contexts: badge notification counts on accent/warning backgrounds now use `#fff`
- **Notification type badges** — Added dark mode overrides for system/announcement/steward/deadline notification badges (were using light backgrounds with dark text, now use translucent colored backgrounds with light text)
- **Missing CSS variables** — Added `--text-secondary`, `--card`, and `--bg-tertiary` to ThemeEngine for all 4 theme variants (steward/member × dark/light), eliminating poor-contrast fallback values
- **Dark mode not persisted** — `AppState.isDark` toggle now saved/read from `localStorage` across all three views (index, steward, member). Previously reset to dark on every page reload.
- **`applyDashboardTheme()` didn't apply** — Menu dialog theme selector saved preference but never called `APPLY_SYSTEM_THEME()`. Sheets kept old styling until manual re-apply.
- **Two conflicting sheet theme systems** — `THEME_CONFIG.THEMES` (04b) used uppercase keys (`LIGHT`/`DARK`/`PURPLE`/`GREEN`) that conflicted with `THEME_PRESETS` (03) lowercase keys. Unified `THEME_CONFIG` to use `THEME_PRESETS` keys; `showThemeManager()` now calls `applyThemePreset()` for consistent behavior.
- **`--surfaceAlt` CSS variable never set** — Added `surfaceAlt` to all 4 ThemeEngine theme definitions (steward dark/light, member dark/light). Previously relied on CSS fallback values.

## [4.28.6] - 2026-03-15

### Fixed
- **Gmail scope test false failure (correct fix)** — `test_emailsend_gmailAppAccessible` was failing because both `GmailApp.getAliases()` (needs `gmail.readonly`) and `GmailApp.createDraft()` (needs `gmail.compose`) require scopes broader than `gmail.send`. The `gmail.send` scope ONLY supports `sendEmail()` — no side-effect-free probe method exists. Replaced with structural verification: GmailApp service availability + `ScriptApp.getOAuthToken()` token check. Runtime send verification deferred to `testAuthEmailSend()`.
- **Authsweep testRunnerEndpointsGated false failure** — `test_authsweep_testRunnerEndpointsGated` expected `dataRunTests(null)` to reject with `success:false`, but SSO via `Session.getActiveUser().getEmail()` is active when the TestRunner runs in the web app, so `_resolveCallerEmail(null)` returns the steward's email via SSO, bypassing the null token. Rewritten to verify auth gate existence structurally instead of testing null-token rejection (impossible when SSO is active).

### Added
- **4 new authsweep tests** — `test_authsweep_wqEndpointsGated` (WeeklyQuestions: 8 gated + 2 utility), `test_authsweep_qaEndpointsGated` (QA Forum: 9 gated), `test_authsweep_tlEndpointsGated` (Timeline: 10 gated), `test_authsweep_fsEndpointsGated` (Failsafe: 9 gated + 1 trigger-only). Total authsweep coverage: 10 tests, ~100 endpoints verified.

### Verified (Auth Sweep)
- All 10 OAuth scopes in `appsscript.json` match actual API usage — no missing or unused scopes
- All `data*` (50+), `wq*` (11), `qa*` (10), `tl*` (10), `fs*` (9) endpoints have auth gates
- No simple trigger violations (ScriptApp calls correctly deferred to installable triggers)
- Email sending: dual-path (GmailApp primary, MailApp fallback) with quota guards verified

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


---

For versions prior to 4.25.0, see `docs/CHANGELOG_ARCHIVE.md`.
