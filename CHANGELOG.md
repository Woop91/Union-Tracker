# Changelog

All notable changes to the Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
