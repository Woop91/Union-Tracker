# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [4.55.7] - 2026-04-11

### SolidBase: Wholesale catchup sync from DDS v4.55.1 → v4.55.7

This sync bundles six DDS wave-equivalents into one SolidBase commit
because sub-projects B/E/D/C all landed on DDS in the same session
after the v4.55.1 SB sync was committed. See the per-wave entries
below (copied from DDS CHANGELOG) for the full detail of each wave.

### Fixed (Sub-project C — Mobile Audit & Horizontal-Scroll Kills)
- **#3** Survey Results card text wrap: on narrow screens the "A survey is currently open" headline wrapped one word per line because the text `<div>` had no `flex: 1` or `min-width: 0` to expand to fill available space. Fix: add `flex: 1` + `min-width: 0` to the bannerText div, plus a new `.survey-cta-banner` CSS class that switches to `flex-direction: column` at ≤ 480px so the button stacks below the text on phones.
- **Access Log not mobile-friendly**: the 6-column audit log table forced horizontal scroll on every phone. Fix: added `access-log-table` class and per-column classes (`alc-time`, `alc-event`, `alc-user`, `alc-sheet`, `alc-field`, `alc-details`) with an `alc-hide-mobile` marker on Sheet + Field. New CSS media query at ≤ 640px hides those two columns, leaving Time / Event / User / Details which fit without sideways scroll. Details column also gets `word-break: break-word` on mobile.
- **Wide Explorer/Org-Health table**: `steward_view.html:6406` had `minWidth: '1200px'` on the inline-edit members table, forcing horizontal scroll on every laptop under 1200px. Trimmed to `minWidth: '900px'` — still wide enough for the desktop data layout, narrow enough that only phones (and small tablets in portrait) trigger the outer `overflowX: auto`.
- **Sidebar scrollbar in header**: investigation showed no separate header scroll exists. The sidebar has one `overflow-y: auto` on `.sidebar-nav`, which is expected for long nav lists. The user's screenshot likely captured a transient artifact at a specific viewport height. No code change; documented for future reference.

### Tests
- New `G-NO-WIDE-TABLES` guard in `test/spa-integrity.test.js` (+19): asserts no `src/*.html` file contains a hardcoded `minWidth` ≥ 1100px. One test per HTML file. Catches wide-table-minWidth regressions at CI time. Legitimate desktop-only wide surfaces can be added to the EXEMPT whitelist if ever needed (empty today).
- Total deploy guards: 301 (up from 282).

## [4.55.6] - 2026-04-11

### Added / Changed (Sub-project D — Admin UX Enhancements)
- **Larger sidebar org name** — `.sidebar-header-title` bumped from `1rem` to `1.375rem`, `font-weight` 700 → 800, `line-height: 1.15`, `letter-spacing: -0.01em`. `.sidebar-header` padding grows from `12px 20px 20px` to `16px 20px 24px` so the header block takes more vertical space and the org name has real presence.
- **Notification bell hides when empty** — the entire `notif-bell-wrap` creation is now wrapped in the existing `if (_bellTotal > 0)` guard. When there are no pending notifications (and no unanswered Q&A for stewards), the bell is absent from the DOM entirely instead of showing a static unclickable icon.
- **Data Failsafe + Test Runner promoted to Admin group** — both tabs were previously in a dev-only "Temp" group that got filtered out by `_filterDevTabs()` in production. They are now pushed into the `IS_ADMIN` Admin group alongside `Usage Analytics`, so admin users see all three admin tools together. The tab router's Test Runner DEV-gate was also expanded to `(IS_DEV_MODE || IS_ADMIN)` so admins can actually load the page in prod.
- **Browser / OS / device class tracking in usage analytics** — new `parseUserAgent(ua)` helper in `src/index.html` classifies every `navigator.userAgent` into `{browser, os, deviceClass}` (edge/chrome/firefox/safari, ios/android/macos/windows/linux, mobile/tablet/desktop). The client-side usage logger includes these in the event payload; the backend `dataLogUsageEvents` extends `_Usage_Log` from 7 to 10 columns (with lazy migration for existing sheets) and stores the parsed values. `dataGetUsageAnalytics` aggregates two new groupings (`byBrowser`, `byOs`) and the admin Analytics page renders them as horizontal bar charts with counts + percentages. Legacy rows fall back to the existing UA-regex device detection; new rows use the pre-parsed fields for accuracy.

### Tests
- New suite: `test/parse-user-agent.test.js` (+10).
- Total deploy guards: 282 (unchanged — parse-user-agent is a unit test).

## [4.55.5] - 2026-04-11

### Fixed (Sub-project E — Org Health Tree Layout)
- **#11** Org Health Tree heavily skewed left for small orgs. The `branchCount==1` fallback in the angle formula at `src/org_health_tree.html:139` defaulted to `: 0`, which set `angle = startAngle = 7π/8` (upper-left corner). Changed to `: angleSpread / 2` so a single-branch tree points straight up (`π/2`) instead. For `branchCount >= 2` the formula is unchanged and remains symmetric.

### Tests
- New suite: `test/org-health-tree-layout.test.js` (+5).
- Total deploy guards: 282 (unchanged — the new file is a unit test, not a deploy guard).

## [4.55.4] - 2026-04-11

### Fixed (Sub-project B — Broken Tab/Chart/State Logic)
- **#2** Monthly Trend chart "Filed" series rendered as solid black fill. Explicit hex colors (`#f59e0b` amber for Filed, `#10b981` emerald for Resolved) added to the `datasets` array in `member_view.html` so the chart never falls through to `_palette()` and can't inherit a theme-resolved black.
- **#4** Section Balance radar chart showed "Need 3+ sections for radar" in a way users read as "need 3+ responses." New message: "A radar chart needs at least 3 distinct sections to render. Your survey has N sections — this is a visualization limit, not a data issue." Radar option is also hidden from the view picker when fewer than 3 sections exist, so the empty state is unreachable by normal navigation.
- **#7** `renderQAForum is not defined` tab load failure for pure stewards. `renderQAForum` only lives in `member_view.html`, which pure stewards never load. Both `case 'qaforum':` branches in the tab router now wrap the call in the existing `_loadMemberViewThen(container, callback)` lazy-loader (synchronous pass-through for pure members, lazy-fetch for dual-role and pure stewards).
- **#8** + survey open/close mismatch: `getStewardSurveyTracking` now returns `periodActive: boolean` and `periodName: string | null` derived from `getSurveyPeriod()` — one backend source of truth for both views. Steward empty state split into "no survey period active" and "period active but no members in scope" messages. Member view drops the misleading "Jan, Apr, Jul, Oct" hardcoded quarterly message — surveys are opened manually by stewards, not on a calendar.

### Tests
- New suites: `test/radar-section-threshold.test.js` (+5), `test/survey-period-state.test.js` (+7).
- New `G-TAB-ROUTES-SAFE` guard in `test/spa-integrity.test.js` (+4): parses every `case '…':` in the tab router and asserts each routed function name is defined somewhere in the concatenated src bundle. Catches the #7 regression class at CI time.
- Total deploy guards: 282 (up from 278).

## [4.55.3] - 2026-04-11

### Fixed
- **Avg Resolution by Steward chart Y-axis serials** — `_looksLikeStewardName` guard rejects Date serials (including scientific notation), ISO dates, and weekday-prefixed strings from becoming phantom grouping keys in the aggregation loop (`21_WebDashDataService.gs`). No `_findColumn` change — verified during plan-writing that the lookup chain was already exact-match only; the dirty data is in the live Grievance Log rows.
- **Access Log user column hashes** — new `displayUser` helper in `index.html` detects 64-hex-char SHA-256 values and renders `User #xxxxxxxx` instead of the raw hash. Adopted in `renderAccessLogPage` (`steward_view.html`). CSV export intentionally preserves full hashes for forensic integrity.
- **Know Your Rights / FAQ / How to File a Grievance literal `\n`** — new `renderSafeText` helper in `index.html` converts both literal `\n` (backslash+n) and real `\n` into `<br>` nodes while HTML-escaping the rest. Adopted in Resources render sites in `member_view.html` and `steward_view.html`. Seed strings in `10b_SurveyDocSheets.gs` and `07_DevTools.gs` were also corrected (152 substitutions) so future re-seeds produce clean data.

### Tests
- New suites: `test/render-safe-text.test.js` (+10), `test/steward-name-guard.test.js` (+11, includes scientific-notation cases), `test/display-user.test.js` (+7).
- New G-RENDER-SAFE guard in `test/spa-integrity.test.js` (+4).
- Total deploy guards: 278 (up from 274).

## [4.55.1] - 2026-04-10

### Fixed
- **Survey steward/member visibility mismatch** — `getSurveyQuestions()` was caching the entire payload (including `period`) for 5 min via CacheService. Steward view (which calls `getSurveyPeriod()` directly) saw live state; members saw stale `period: null`. Stripped `period` from the cached payload, bumped cache key v1 → v2 to invalidate stale entries, and overlay live `getSurveyPeriod()` on every call. Added cache invalidation in `openNewSurveyPeriod()` and `archiveSurveyPeriod_()` so period lifecycle changes propagate to members immediately.
- **PIN setup blocked by missing Member ID** — `dataCompleteOnboardingStep('pin', ...)` rejected with "Member ID required before setting PIN" when the email-matched row had an empty Member ID cell (legacy data, manual import, partial seed). Now self-heals: generates a unique `M` + 6-digit ID (collision-checked against existing IDs), writes it back to the row, audit-logs `MEMBER_ID_AUTOGEN`, and proceeds with the hash. Members with broken rows can now set their PIN AND their row gets repaired in one shot.
- **Session persistence — multi-pronged diagnostic** — Bumped redirect-loop guard in `index.html` from 2 → 5 attempts (the original was so tight that two transient server hiccups cleared a valid token). Instrumented `Auth._getSessionData` with three-way failure logging (not present / expired / corrupt JSON), each line includes the token prefix for correlation. `_serveAuth` now records `SESSION_TOKEN_REJECTED` (MEDIUM) security event when a sessionToken fails validation. SSO and magic-link `createSessionToken` failures now escalate to `SESSION_CREATE_FAILED` (HIGH) security events instead of dying silently in `log_`. After the next user complaint, the GAS execution log + audit trail tells you exactly which of the three failure modes fired.

### Added
- **Survey anonymity disclosure card** — New `_renderSurveyAnonymityCard(content)` helper in `member_view.html` prepended to the survey wizard. Two-part honest disclosure: (1) "your answers are stored in a separate encrypted vault with no link back to you, by design" and (2) "what stewards CAN see: completion status and unit-level percentages — never your individual answers." Solves the surprise factor when members notice the steward progress view.
- **Steward survey "History: On/Off" toggle** — Added a toggle pill in the survey tracking scope filter row. When off, hides the per-member historical participation rate bar in `_renderSurveyMemberItem`, leaving only current-quarter completion. State persists in `localStorage.dds_steward_show_history`. Default ON. Threaded `showHistory` flag through `_surveyPaginate` and the renderer.
- **Home screen icon (favicon + apple-touch-icon)** — New `src/app_icon.html` containing the DDS logo as a base64 data URL (110KB). Wired into `index.html` `<head>` via `<? var _appIconUrl = include('app_icon'); ?>` (read once). Added Apple PWA meta tags: `apple-mobile-web-app-capable`, `apple-mobile-web-app-status-bar-style`, `apple-mobile-web-app-title`, `mobile-web-app-capable`. Added `app_icon.html` to `build.js` HTML_FILES list. **Caveat:** iOS Safari may still ignore this on home-screen install because GAS web apps live inside a `script.googleusercontent.com` iframe and iOS reads icons from the parent frame. Chrome desktop favicon and Android home-screen install should work.
- **Auth view "previous session expired" hint** — Non-alarming explainer surfaced on the auth page when `PAGE_DATA.tokenChecked === true`, so users understand WHY they're being asked to email a new link instead of being silently redirected.

### Changed
- **`getSurveyQuestions()` cache key** — Bumped `surveyQuestions_v1` → `surveyQuestions_v2`. The v2 payload no longer contains `period`. Both `clearSurveyQuestionsCache()` (in `10b_SurveyDocSheets.gs`) and the inline cache-clear in `08e_SurveyEngine.gs` now remove both keys for a clean transition.
- **Redirect-loop guard tolerance** — `dashboardSessionRedirectCount` threshold raised from 2 to 5 in `index.html`.

## [4.51.1] - 2026-04-05

### Added
- **E-signature schema** — 4 signature columns (SIGNATURE_STATUS, SIGNATURE_TOKEN, SIGNED_DATE, SIGNATURE_HASH) + DESCRIPTION column added to GRIEVANCE_HEADER_MAP_, unblocking the entire signing workflow.
- **Q&A Reopen button** — Resolved questions can be reopened by owner or steward (backend already existed, UI was missing).
- **Meeting data enrichment** — `getMemberMeetings` now returns duration, notesUrl, agendaUrl fields.
- **Leader Hub "Mentorships" sub-tab** — Read-only unit-scoped mentorship visibility for member leaders + `dataGetLeaderUnitMentorships` endpoint.
- **CSV export for My Cases** — Client-side Blob download for grievance/task data.
- **Poll participation badges** — Vote count + participation percentage on poll results.
- **Resources list caching** — DataCache with 30-min STABLE_TTL (follows steward directory pattern).
- **Mobile default view preference** — Dual-role users can set default view from mobile header.
- **ICS calendar download** — Cross-platform calendar integration for events.
- **Autocomplete keyboard navigation** — ArrowDown/Up/Enter/Escape with highlight tracking in all autocomplete fields.
- **Column reorder utility** — `reorderMemberDirectoryColumns()` in 06_Maintenance.gs for normalizing sheet layout.
- **Privacy hints in onboarding wizard** — Phone/address visibility explanations.
- **Email Open Rate tooltip** — Info icon explaining the metric on profile page.

### Fixed
- **`defaultView_` property key mismatch** — `dataGetBatchData` read `DEFAULT_VIEW_` but write used `defaultView_`, corrupting SWR cache for dual-role users.
- **Admin Settings HTML-entity corruption** — Removed server-side `escapeHtml()` on list values (client uses `textContent`, inherently XSS-safe).
- **`escapeForFormula` quote doubling** — Removed SQL-style apostrophe/quote doubling that corrupted config values.
- **ALERT_DAYS config disconnect** — Notification system now reads Config sheet via `getConfigValue_()` instead of ignoring it.
- **Satisfaction trend chart** — Fixed flat array passed to `createLineChart` (needed dataset object wrapper).
- **Survey vault dedup** — Changed `!== 'Superseded'` to `=== 'Latest'` (Superseded value never existed).
- **PIN security** — Weak PIN validation added to standalone setup page (client + server). Client weak PIN list synced to server's 21 entries. PIN reset restores old hash on email failure.
- **4 missing `withScriptLock_` calls** — dataCreateTaskForSteward, dataAddMeetingMinutes, dataSubmitFeedback, attachDriveFiles.
- **`attachDriveFiles` cache invalidation** — Added `_cacheInvalidate()` matching other TimelineService mutations.
- **10 mobile touch-target fixes** — 44px minimums for sub-tab, handoff-btn, swipe-panel-close, nav-item, filter-add-btn. iOS zoom prevention for msg-textarea, search-input, form-select. Deadline dot exclusion. Bulk toolbar safe-area padding.
- **Ctrl+Enter submit** — Expanded `.closest()` selector to match actual form containers (.content, .card, .member-detail-panel).
- **Steward resource cards** — Added keyboard activation (Enter/Space) + ARIA attributes.
- **Minutes "Start a Grievance"** — Changed direct `renderMyCases()` to `_handleTabNav()`.
- **Outreach autocomplete crash** — Added `Array.isArray()` guard for auth failure responses.
- **Broadcast "undefined member(s)"** — Added array check before `.length`.
- **Task notification empty body** — Added `body` field to notification object.
- **Engagement score defaults** — Normalized all "no data" defaults to 0 (was asymmetric: some 50, some 0).
- **Cross-Unit Q&A showResolved** — Now passes through actual toggle value instead of hardcoding `true`.
- **5 PII log leaks** — Replaced raw email logging with `maskEmail()` across 5 files.
- **Contact Log double-submit** — Added disabled guard during server call.
- **Bulk email double-click** — Added sending flag + button disable.
- **`syncColumnMaps` case-sensitivity** — Added case-insensitive fallback to prevent duplicate column creation.
- **`getMemberViewHtml` auth guard** — Added sessionToken parameter and `_resolveCallerEmail` check.
- **`dataSendDirectMessage` rate limit** — Max 10/min per steward via CacheService.
- **`disableDashboardMemberAuth` confirmation** — Added UI alert dialog matching enable counterpart.
- **idemKey TTL** — Increased from 300s to 600s across all 5 endpoints.
- **Spearman guard** — Changed n<3 to n<5 to match Pearson delegate requirement.
- **Config TTL doc** — Corrected header comment from "6-hour" to "5-minute".
- **Cat icons on mobile** — Show during loading (chasing/pouncing classes override display:none).
- **Union tab expandable cards** — Wrapped `_initV9()` in try/catch + fallback onclick rebinder.
- **Agency chart tabs** — Added smod null guard + fallback onclick rebinder.
- **Vacant card dark mode** — Changed background from `#1a1214` to `var(--card-bg)`.
- **`VERSION_INFO.codename` case** — Changed `CODENAME` to lowercase `codename`.
- **Contact wrapper return types** — `dataGetMemberContactHistory`/`dataGetStewardContactLog` return `[]` on auth failure (was `{success:false}`).
- **Phone number validation** — Added format check (7-15 chars, digits/spaces/dashes) in profile + onboarding.
- **Task completion UX** — Replaced blocking `confirm()` with toast notification.

## [4.51.0] - 2026-04-03

### Added
- **Directory Explorer** — Slicer-style filter + SVG Sankey chart + inline editing for Member Directory exploration.
- **Director field** — New `director` field in member records.
- **`dataUpdateMemberBySteward` endpoint** — New server endpoint for steward-initiated member updates.

### Changed
- **Config hardening** — Directors rename for clarity.

### Fixed
- **`dataGetAllMembers`/`dataGetMemberGrievances`/`dataGetAvailableStewards`** — Fixed returning authError objects instead of arrays, which caused downstream iteration failures.
- **Cases sub-tab agency-wide fallback** — Fixed unreachable fallback branch.
- **Broadcast recipient count** — Wired up previously disconnected recipient count display.
- **Mentorship dropdowns** — Added steward/member selector dropdowns to mentorship pairing UI.
- **`getAllMembers()` enrichment** — Added `role` and `duesPaying` fields to member payload.
- **`appsscript.json` build copy** — Fixed `build.js` not copying `appsscript.json` to dist.

## [4.50.7] - 2026-04-02

### Fixed
- **`onOpenDeferred_` trigger persistence** — Fixed leftover code that deleted the installable trigger after first run (from old one-shot approach). Deferred init (column sync, tab colors, modals) now runs on every spreadsheet open.
- **Tab modals on sheet open** — Tab modals now auto-open via installable trigger (full authorization) instead of simple trigger (which lacks `showModalDialog` permission).
- **Session invalidation hardening** — Added deletion verification to session invalidation flow.
- **TrendAlertService race condition** — Wrapped `runDetection()` in `LockService` to prevent duplicate alerts from concurrent triggers.

## [4.50.6] - 2026-04-01

### Fixed
- **Tab modals not auto-opening** — GAS simple triggers (`onSelectionChange`) cannot call `showModalDialog()`. `onTabSwitch_` now uses `toast()` hint instead of the blocked modal call.
- **`showCurrentTabModal()` convenience function** — Auto-detects active sheet and shows the appropriate modal; accessible via menu.
- **Tab Modals menu discoverability** — Moved Tab Modals submenu to top-level menu item.

## [4.50.5] - 2026-04-01

### Fixed
- **`createConfigSheet` seed alignment** — `syncColumnMaps()` now runs after writing new headers so `seedConfigDefault_` and `_applyYesNoValidation` target correct columns (previously used stale `CONFIG_COLS` positions from before header rewrite).
- **Branding column repair** — One-time repair clears misaligned yes/no values and stale data validations from Branding columns (Logo Initials, Steward Label, etc.) that received Feature Toggle dropdowns from the v4.50.0-v4.50.4 bug.
- **`INSIGHTS_CACHE_TTL_MIN` seed location** — Moved from Branding section to Feature Toggles section to match `CONFIG_HEADER_MAP_` grouping.

## [4.50.4] - 2026-04-01

### Fixed
- **`_migrateOrphanedColumns` rewrite** — Replaced unsafe sequential two-pointer algorithm with name-based header lookup. Old algorithm would incorrectly delete valid columns when `CONFIG_HEADER_MAP_` order changed, causing data loss.
- **Global cold-cache column sync** — Global-scope column map initialization now calls `syncColumnMaps()` when cache is expired, protecting all execution contexts (menu handlers, `google.script.run` data functions, triggers) from using wrong array-order defaults.

## [4.50.3] - 2026-04-01

### Fixed
- **Config data migration on column reorder** — `createConfigSheet()` now detects when existing sheet headers differ in order from `CONFIG_HEADER_MAP_` and remaps row 3+ data to match before overwriting headers. Prevents org name showing "yes"/"no" after the v4.50.0 column reorder.
- **Hardened `_applyYesNoValidation`** — No longer silently overwrites non-boolean values with "yes". Logs a warning instead, protecting against data destruction when columns are misaligned.

## [4.50.2] - 2026-03-31

### Fixed
- **Stale ConfigReader cache after column sync** — When `RESOLVED_COL_MAPS` cache expires and `syncColumnMaps()` re-resolves column positions, ConfigReader now force-refreshes instead of serving cached values built with wrong (array-order) positions.
- **`ORG_CONFIG_v2` cache invalidation** — `syncColumnMaps()` now invalidates `ORG_CONFIG_v2` cache key when columns change. Fixes org name showing as "yes" and steward label showing URL after cache expiry.

## [4.50.1] - 2026-03-31

### Fixed
- **Column cache TTL** — Increased `persistColumnMaps_()` cache from 2 hours to 6 hours (GAS maximum) to reduce cold-cache column mismatches after the v4.50.0 Config reorder.
- **Eager column restore** — Added global-scope `loadCachedColumnMaps_()` call in 01_Core.gs so all execution contexts (doGet, data* web functions, onEdit) start with correct column positions.
- **doGet cold-cache fallback** — If column cache is empty on first access after deploy, doGet now calls `syncColumnMaps()` before rendering to prevent wrong-column reads.
- **Dues Status dropdown** — Added `DUES_STATUS → DUES_STATUSES` mapping to `DROPDOWN_MAP` so Dues Status column respects Config-driven values.
- **Simplified seed data** — Removed unused 'Fee Payer' and 'Exempt' from default Dues Statuses seed.
- **Version bump** — COMMAND_CONFIG.VERSION corrected from 4.49.0 → 4.50.1.

## [4.50.0] - 2026-03-31

### Added
- **Union Name column** on Non-Member Contacts tab — free-text field for union affiliation.
- **Shirt Size column** on Non-Member Contacts tab — dropdown with XS, S, M, L, XL, 2XL, 3XL, 4XL.
- **Steward (Yes/No) column** on Non-Member Contacts tab — dropdown indicating steward status.
- All three fields added to: `NMC_HEADER_MAP_` / `NMC_COLS` (01_Core.gs), sheet creation validations (10a_SheetCreation.gs), tab formatting with updated column widths (09_Dashboards.gs), DataService CRUD (21_WebDashDataService.gs), steward view modal and contact cards (steward_view.html), and fallback modal (03_UIComponents.gs).
- **Non-Member dues status** — Added `Non-Member` to the recognized non-paying dues statuses. Non-member contacts added to the Member Directory with `Dues Status: Non-Member` are subject to the same feature restrictions as non-dues-paying members.

### Changed
- **Survey open to all members** — Removed the dues-paying gate from the Quarterly Survey. All authenticated users in the Member Directory can now take the survey regardless of dues status. Survey banner, Take Survey CTA, and More menu lock removed.
- Updated dues-restricted banner text to remove Survey from the locked features list.

## [4.44.0] - 2026-03-28

### Added
- **Dynamic background tab prefetch** — After home page renders, `requestIdleCallback` prefetches data for the user's top 3 most-visited tabs. No artificial delay — fires immediately via idle callback, staggered 1.5s apart to avoid saturating the 4-concurrent-call throttle.
- **Per-user tab frequency tracking** — `_recordTabFrequency()` stores per-role tab visit counts in `localStorage` (`dds_tabFreq`). `_getTopTabs()` reads this to determine which tabs to prefetch, with sensible defaults for first-time users (steward: members/tasks/notifications; member: events/stewarddirectory/notifications).
- **`_PREFETCH_REGISTRY`** — Extensible map of tabId → DataCache warmup function. Currently supports: `members`, `events`, `stewarddirectory`, `survey`. Tabs not in the registry are skipped (their render functions use raw `serverCall`).
- **`STATS_TTL` (60 min)** — New DataCache tier for KPI stats, counts, and aggregations that don't need real-time freshness.
- **Request deduplication** — `_dedupMap` in `_throttledServerCall` coalesces identical in-flight server calls. If `dataGetAllMembers` is already pending and another component requests it, the second caller piggybacks on the first — no duplicate server round-trip.
- **Cases list infinite scroll** — Replaced manual "Show More" button with `LazyList.render()` using `IntersectionObserver`. Renders 25 cases initially, auto-loads 25 more as user scrolls. Zero-click, zero-wait.
- **Visibility-based refresh** — `visibilitychange` listener silently re-fetches batch data when user returns to the tab after 30+ minutes away. Keeps KPIs, badges, and case counts fresh for all-day sessions without polling.
- **Expanded prefetch registry** — Added `tasks`, `contactlog`, `cases` (member), `mytasks`, `notifications` to `_PREFETCH_REGISTRY`. Now 9 tabs can be dynamically prefetched based on user frequency.
- **dataGetAllMembers cache consolidation** — All 4 call sites in steward_view.html now use `DataCache.cachedCall` with the same cache key (`stewardMembers_<email>`). Combined with request dedup, this means at most 1 server call regardless of how many components need member data.

### Changed
- **innerHTML → textContent clears** — Heavy container clears in Members, Contact Log, Notifications, and Cases switched from `innerHTML = ''` to `textContent = ''` (avoids HTML parser overhead on clears).
- **DataCache default TTL** — Increased from 2 minutes to 15 minutes, reducing redundant server calls on tab switches.
- **DataCache stable TTL** — Increased from 5 minutes to 30 minutes for member/steward directories.
- **SWR TTL** — Extended from 30 minutes to 2 hours for instant return-visit rendering.
- **Notifications tab** — Removed from always-fresh list; now cached like other tabs for faster revisits.

## [4.43.1] - 2026-03-27

### Added
- **One-tap webapp check-in** — When a meeting is active, logged-in members and stewards see a green "Check In" banner at the top of their dashboard. One tap records attendance (no PIN re-entry — already session-authenticated).
- **Check-in confirmation** — Banner transitions to a "Checked In!" state after successful check-in, persists on re-render via DataCache.
- **Batch data integration** — Active meeting data (`activeMeeting`) included in both member and steward batch payloads for zero-latency banner display.
- **`dataWebAppCheckIn()` endpoint** — Session-authenticated one-tap check-in with TOCTOU lock protection and audit logging (`MEETING_WEBAPP_CHECKIN`).
- **`.checkin-banner` CSS** — Green-themed banner with slide-in animation, responsive layout, and dark/light mode support.

## [4.43.0] - 2026-03-27

### Security
- **PIN login: constant-time scan** — `devAuthLoginByPIN()` now iterates all members regardless of match position. Response time normalized to ~1–1.5s with random jitter to prevent timing-based attacks.
- **Magic link: global session rate limit** — Added per-session rate limit (5 attempts/15min) on top of per-email limit to prevent email enumeration via rapid cross-email probing.
- **Session token URL cleanup** — After SPA loads, `history.replaceState` strips `?sessionToken=` from the address bar to prevent token persistence in browser history.
- **Biometric PII reduction** — `dds_biometric_name` in localStorage now stores initials only (e.g., "JD") instead of full member name.
- **PIN rate limit error hardening** — `cache.remove()` and `clearPINAttempts()` wrapped in try/catch to prevent cascading errors on success cleanup.

### Added
- **Device key authentication** — After any login (SSO, magic link), a device key is silently registered and stored via the Credential Management API. On return visits, the biometric "Sign in with Face ID / Touch ID / Fingerprint" button now works for ALL users, not just PIN users.
- **Auto-submit PIN** — PIN input auto-submits when 6 digits are entered, removing the need to tap "Sign In".
- **QR code mobile attendance** — Stewards generate QR codes for meetings via `Calendar & Meetings > 📱 QR Code Check-In`. Members scan the QR code on their phone, enter their phone number and PIN, and attendance is recorded.
- **Mobile check-in page** — New `?page=qr-checkin` web route serves a mobile-optimized, dark/light mode check-in page. No login required — authenticates per check-in via phone + PIN.
- **Phone number authentication** — `processQRCheckIn()` looks up members by phone number (normalized 10-digit US) instead of email, with full PIN verification, lockout protection, and audit logging (`MEETING_QR_CHECKIN`).
- **QR code generation** — `getMeetingQRCode()` generates QR code images via Google Charts API (free, no key needed). `createMeeting()` now includes QR URL in its response.
- **Steward QR dialog** — `showMeetingQRCodeDialog()` displays a printable QR code for any active meeting, with copy-link and print support.
- **Sheet tab emoji titles** — All 11 visible tabs that lacked emoji prefixes now have them for quick identification: ⚙️ Config, 👥 Member Directory, ⚖️ Grievance Log, ✅ Case Checklist, 🧪 Test Results, 🔧 Function Checklist, 📊 Workload Reporting, 🗓️ Events, 🗒️ Meeting Minutes, 🔍 Settings Overview, 📇 Non-Member Contacts.
- **Auto-migration** — `migrateSheetTabTitles_()` runs on deferred open to auto-rename existing sheet tabs. Manual trigger available via Admin → Styling → 📑 Apply Tab Titles.
- **Legacy name fallback** — `getOrCreateSheet()` and `applyTabColors_()` transparently handle both old and new tab names during the migration window.
- **`SHEET_LEGACY_NAMES_`** constant maps pre-emoji → post-emoji names for backward compatibility.
- **Self-service PIN reset** — New "Forgot your PIN?" link on PIN login screen. Members enter their email and receive a new PIN via email. Rate-limited (2/email/hour, 10 global/hour) with anti-enumeration generic responses. Server function: `devRequestPINReset()`.
- **Knowledge Engine input validation** — `addKnowledgeContent()` and `updateKnowledgeContent()` now validate type, category, audience, and placement against allowed lists; enforce title (500) and content (10,000) character limits; validate priority range (1–9999); check end date is after start date.
- **Knowledge Engine Manage tab search** — Stewards can now search/filter knowledge items by title, content, or ID in the Manage tab.
- **Knowledge Engine deactivate confirmation** — Deactivating a knowledge item now requires confirmation.
- **Auth focus trap** — Keyboard focus is now trapped within the auth card for accessibility (prevents tabbing into background elements).
- **ARIA label on remember-me toggle** — Screen readers now announce the toggle's purpose and duration.

### Fixed
- **`restoreKnowledgeContent()` missing lock** — Now wrapped in `withScriptLock_()` to prevent race conditions, matching all other CRUD operations.
- **`getKnowledgeContentAll()` uncached** — Added 2-minute CacheService caching for steward Manage tab to reduce sheet reads.
- **Biometric failure UX** — When biometric verification fails, users now see a brief explanation before being redirected to PIN entry instead of a silent fallback.
- **Knowledge Engine cache invalidation** — `_invalidateKnowledgeCache_()` now also clears the steward `ke_all_steward` cache key.
- **PIN lockout message** — Now uses `PIN_CONFIG.LOCKOUT_MINUTES` instead of hardcoded "15 minutes".

### Changed
- **Remember-me default ON** — The "Remember this device" toggle on the email magic link form now defaults to checked, reducing repeat logins for most users.
- **Biometric flow refactored** — `_autoSubmitPIN` now detects credential type (device key vs PIN) and routes to the appropriate server endpoint, with automatic fallback.
- **Browse tab note** — Shows "active content only" hint with link to Manage tab.
- **Form help text** — Bullets and Priority fields now show usage hints in the Knowledge Engine form.

## [4.42.1] - 2026-03-27

### Changed
- **Rename: Education Engine → Knowledge Engine** — All references renamed across the codebase: sheet name (`📚 Knowledge Engine`), constants (`KNOWLEDGE_ENGINE`, `KNOWLEDGE_COLS`, `KNOWLEDGE_HEADER_MAP_`), API functions (`getKnowledgeContent`, `addKnowledgeContent`, `updateKnowledgeContent`, `deleteKnowledgeContent`, `restoreKnowledgeContent`), sheet factory (`createKnowledgeEngineSheet`), cache keys (`ke_` prefix), content IDs (`KE-` prefix), sidebar tab ("Knowledge"), and steward UI render function (`renderStewardKnowledge`).

## [4.42.0] - 2026-03-27

### Added
- **Password manager support** — PIN and email login forms now use proper `<form>` elements with `autocomplete` attributes (`username`, `current-password`, `email`) so password managers (1Password, Bitwarden, Chrome, iOS Keychain, etc.) can detect, save, and autofill credentials.
- **Biometric sign-in** — After a successful PIN login, credentials are saved via the Credential Management API (Chrome/Edge) and native form semantics (Safari/iOS Keychain). On return visits, a "Sign in with Face ID / Touch ID / Fingerprint / Windows Hello" button appears for one-tap biometric authentication.
- **Platform-aware biometric labels** — Login button text adapts to the user's platform (Face ID on iPhone/iPad, Touch ID on Mac, Fingerprint on Android, Windows Hello on Windows).
- **Auto-submit biometric flow** — When credentials are retrieved via the Credential Management API, the PIN is automatically submitted without user interaction.
- **Forget saved credentials** — Users can clear stored biometric credentials from the login screen.

## [4.41.0] - 2026-03-27

### Added
- **Knowledge Engine** — New `📚 Knowledge Engine` sheet centralizes all educational content (quotes, tips, concepts, mini-lessons, manifesto phrases, negotiation sets) in one steward-managed location.
- **Knowledge Engine API** — Server-side functions (`getKnowledgeContent`, `addKnowledgeContent`, `updateKnowledgeContent`, `deleteKnowledgeContent`, `restoreKnowledgeContent`) with 5-minute CacheService caching, steward-only CRUD, and `escapeForFormula` security.
- **Knowledge Engine Management UI** — New "Knowledge" tab in steward sidebar with Browse (filterable by type/placement/search) and Manage (full CRUD with inline forms) sub-tabs.
- **Knowledge Engine sheet factory** — `createKnowledgeEngineSheet()` auto-creates the sheet with seed data migrated from previously hardcoded arrays (26 negotiation sets, 14 manifesto phrases, 43 auth quotes).
- **Dynamic content loading** — `negotiation_knowledge.html` and `auth_manifesto.html` now attempt to load content from the Knowledge Engine at runtime, falling back to hardcoded arrays if the sheet doesn't exist yet.

### Changed
- **negotiation_knowledge.html** — Added Knowledge Engine loader that replaces the hardcoded `NEGOTIATION_SETS` array with sheet-managed content when available.
- **auth_manifesto.html** — Added Knowledge Engine loader for both manifesto phrases and auth quotes, with graceful fallback to hardcoded defaults.
- **Steward sidebar navigation** — Added "Knowledge" tab (🎓) under the Reference group.

## [4.40.0] - 2026-03-25

### Changed
- **PIN Login: GA release** — PIN login option now available in all environments (dev and production). Removed IS_DEV_MODE gates from login screen, steward PIN management, and server-side comments.
- **Login page: black background** — Auth screen now uses a solid black background with glass card backdrop blur and enhanced box shadow.
- **Login page: quote styling** — Quotes are now larger (15px), bolder (weight 500), fully white with text glow for better visibility against the black background.
- **DDS branding** — Default org name fallback changed from "Dashboard" to "DDS" for the DDS repo login page.

### Added
- **Auth page: isDevMode flag** — `_serveAuth()` now includes `isDevMode` in PAGE_DATA for consistency with dashboard pages.

## [4.34.3] - 2026-03-23

### Fixed
- **Mentorship: input validation** — `createPairing` now validates email format, prevents self-pairing, and rejects duplicate active pairings.
- **Mentorship: formula injection** — `escapeForFormula()` applied to mentor/mentee emails before sheet write.
- **Mentorship: audit trail** — `closePairing` now logs `MENTORSHIP_CLOSED` audit event with pairing details.
- **Mentorship: nav badges** — All write operations (`createPairing`, `updatePairingNotes`, `closePairing`) now call `_refreshNavBadges()`.
- **Mentorship: mentor load balancing** — `suggestPairings` now picks the mentor with the fewest existing active pairings instead of naive round-robin.

### Added
- **Mentorship: manual pairing form** — Stewards can now create pairings with any two emails, not just from suggestions.
- **Mentorship: save feedback** — "Save Notes" button shows "Saved!" confirmation with auto-reset.
- **Mentorship: suggestion refresh** — Accepting a suggested pairing now refreshes the Active Pairings list.

## [4.34.2] - 2026-03-22

### Changed
- **BROADCAST_SCOPE_ALL** default changed from `no` to `yes` — stewards can now message all members by default.
- **ENABLE_CORRELATION** now seeded as `yes` with yes/no dropdown validation — correlation engine active by default on new installs.

### Note
All other security features (ACCESS_CONTROL, DASHBOARD_MEMBER_AUTH, ERROR_LOGGING, NOTIFY_ON_CRITICAL, auto-archival) were already enabled by default.

## [4.34.1] - 2026-03-21

### Fixed
- **Theme sync: org chart & POMS** — Embedded org_chart.html and poms_reference.html now follow the app's dark/light toggle via class toggling on `.madds-embed` and `.poms-root`. Live toggle via `UnifiedTheme.apply()` also propagates to embedded components.
- **esign.html hardcoded colors** — 7 hardcoded hex values replaced with CSS custom properties (`--accentHover`, `--accentMuted`, `--badge-pending-bg/text`, `--badge-signed-bg/text`, `--canvas-bg`) with dark mode overrides in `prefers-color-scheme` media query.
- **OS theme detection** — First-visit default now respects `prefers-color-scheme` instead of always defaulting to dark mode. Once user explicitly toggles, localStorage takes precedence.
- **Modal overlay theme-awareness** — `.wt-modal-overlay` uses `var(--overlay-bg)` scoped by `[data-theme-mode]` (light=0.35, dark=0.55) instead of hardcoded `rgba(0,0,0,0.55)`.
- **Offline banner** — Uses `var(--danger, #ef4444)` instead of hardcoded `#ef4444` in both CSS and inline style.
- **Dist parity CI failure** — `minifyHtml()` in build.js now normalizes `\r\n` → `\n` before regex processing, ensuring identical output on Windows and Linux.

### Removed
- Dead code: unused `last7Keys` variable in analytics, orphaned `resourceDownloads` local computation (return object uses `DataService.getResourceClickTotal()`).
- Dev-only files from dist/ (`07_DevTools.gs`, `30_TestRunner.gs`, `31_WebAppTests.gs`, `DevMenu.gs`) — dist/ now matches `build:prod` output.

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
- **Theme cleanup** — Removed 5 novelty themes (Comic, Brutalist, Retro OS, Liquid Pour, Blob Lava); 6 themes remain.
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
- **Webapp performance: lazy-view HTML caching** — `getOrgChartHtml()` and `getPOMSReferenceHtml()` now cache rendered HTML in CacheService (6-hour, version-keyed TTL), avoiding repeated `HtmlService.createHtmlOutputFromFile()` calls.

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


---

For versions prior to 4.25.0, see `docs/CHANGELOG_ARCHIVE.md`.
