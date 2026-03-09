# AI REFERENCE DOCUMENT — DDS-Dashboard
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-03-07
# Consolidation note: On 2026-02-25, duplicate sections were replaced with
# pointers to their canonical source files. No information was removed.

---

## 📖 HOW TO USE THIS FILE

Read these files **in this order** when onboarding to this codebase:

| Order | File | What it covers |
|-------|------|----------------|
| 1 | **CLAUDE.md** | Critical rules, column constants, security patterns, config write paths, coding conventions, git conventions |
| 2 | **This file (AI_REFERENCE.md)** | Project overview, architecture map, LLM-specific context, error log, protected code |
| 3 | **SYNC-LOG.md** | DDS↔Union-Tracker sync flow, Workload Tracker exclusion registry |
| 4 | **CHANGELOG.md** | Full version history (Keep a Changelog format) |
| 5 | **FEATURES.md** | Detailed feature documentation |
| 6 | **COLUMN_ISSUES_LOG.md** | Recurring column bugs — READ if touching column-related code |
| 7 | **CODE_REVIEW.md** | Canonical security/code review |
| 8 | **DEVELOPER_GUIDE.md** | Developer onboarding |

**Do NOT duplicate content from those files here.** If you need to add context an LLM would need that doesn't fit those files, add it here.

---

## 🏗️ PROJECT OVERVIEW

**What:** Google Apps Script application for union steward grievance tracking, member management, and reporting.
**Repo:** `Woop91/DDS-Dashboard` (private). Default branch: `Main` (capital M).
**Mirror:** `Woop91/Union-Tracker-` (public). See SYNC-LOG.md for exclusion rules.
**Deployed via:** CLASP (`clasp push`) to Google Apps Script, bound to a Google Sheet.
**Target users:** Union stewards (power users) and members (casual users) at MassAbility DDS (SEIU 509).
**Architecture:** 42 source `.gs` files + 7 `.html` files in `src/` → copied individually to `dist/` via `node build.js`.
**Current build:** 42 `.gs` + 7 `.html` files in `dist/` (individual file mode, NOT consolidated).
**Web App:** Served via `doGet()` using inline HTML (`HtmlService.createHtmlOutput()`). Does NOT use `createTemplateFromFile()`.
**DDS Apps Script ID:** `[REDACTED]`
**UT Apps Script ID:** `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`

### ⚠️ Key Reminders
- **Critical rules** (dynamic-only, 1-indexed columns, escapeHtml, etc.) → **See CLAUDE.md**
- **Sync rules & WT exclusions** → **See SYNC-LOG.md**
- **`dist/` files are auto-generated.** Never edit directly. Edit `src/*.gs` files, then run `npm run build`.
- **`dist/ConsolidatedDashboard.gs` is DELETED** as of v4.13.0. Build now copies individual files.
- **`web-dashboard/` folder is LEGACY/ORPHANED.** Do not deploy or integrate it.
- **CLASP rootDir:** `./dist` — only `dist/` contents go to Apps Script.
- **Deploy:** `npm run deploy` (lint + test + build:prod + clasp push). Must run locally (requires Google OAuth).
- **After any merge to Main:** Remind user to run `npm run deploy`. Agent cannot run clasp remotely.

---

## 📁 ARCHITECTURE MAP

### Build Pipeline
```
src/*.gs (42 files) + src/*.html (7 files)
    → build.js (copy-files mode) → dist/*.gs + dist/*.html (individual files)
                                    dist/appsscript.json (manifest)
    → clasp push → Google Apps Script project
```
**Note:** As of v4.13.0, `dist/ConsolidatedDashboard.gs` is deleted. Build copies individual files to `dist/`. GAS needs separate `.html` files for `createTemplateFromFile()` and `createHtmlOutputFromFile()`.

### Source File Load Order
| Prefix | Layer | Key Files |
|--------|-------|-----------|
| `00_` | Foundation | `DataAccess.gs`, `Security.gs` |
| `01_` | Core | `Core.gs` — constants, config, utilities |
| `02_` | Data | `DataManagers.gs` — CRUD operations |
| `03_` | UI | `UIComponents.gs` — menus, dialogs |
| `04a-e` | UI modules | Menus, accessibility, interactive/executive/public dashboards |
| `05_` | Integrations | Drive, Calendar, Email, Web App API functions |
| `06_` | Maintenance | Admin tools, undo/redo, audit |
| `07_` | DevTools | Dev-only utilities (**excluded in prod build**) |
| `08a-e` | Sheet utils | Setup, search, forms, audit formulas, survey engine |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Main entry, sheet creation, forms, sync |
| `11_-17_` | Features | CommandHub, self-service, meetings, events, correlation |
| `19_-25_` | Web Dashboard SPA | Auth, config reader, data service, app entry, portal sheets, workload service |
| `26_` | Q&A Forum | Steward-member Q&A with steward-only answers, resolve/reopen |
| `27_` | Timeline Service | Activity feed with inline edit, pagination, calendar links |
| `28_` | Failsafe Service | Security & reliability guardrails |
| `29_` | Migrations | Column auto-migration for schema changes |

### HTML Files in src/
| File | Purpose |
|------|---------|
| `index.html` | SPA entry point (unified dashboard) |
| `steward_view.html` | Steward command center |
| `member_view.html` | Member dashboard |
| `auth_view.html` | Login/auth page |
| `error_view.html` | Error display page |
| `styles.html` | Shared CSS |
| `org_chart.html` | Organizational chart view |

### Web App Routing (doGet)
```
doGet(e)
├── Default → SPA (doGetWebDashboard) with SSO/magic link auth
├── ?mode=steward  → Steward Command Center (requires auth, shows PII)
├── ?mode=member   → Member Dashboard (no PII)
├── ?page=search   → Search (requires steward auth)
├── ?page=grievances → Grievance list (requires steward auth)
├── ?page=members  → Member list (requires steward auth)
├── ?page=links    → Links page
├── ?page=selfservice → Member self-service (Google auth or PIN)
├── ?page=portal   → Public portal
├── ?page=workload → Workload tracker
├── ?page=checkin  → Meeting check-in (v4.11.0)
├── ?page=resources → SPA with resources tab pre-selected (v4.11.0)
├── ?page=notifications → SPA with notifications tab pre-selected (v4.12.0)
├── ?page=qa-forum → Q&A Forum (v4.22.6)
├── ?page=timeline → Timeline activity feed (v4.22.9)
├── ?page=org-chart → Org chart view (v4.22.6)
└── Deep-link: ?page=X → SPA reads PAGE_DATA.initialTab → _handleTabNav()
```

### Authentication System
- **Steward access:** Google account email matched against authorized list via `checkWebAppAuthorization('steward')`
- **Member access:** Google account email OR PIN-based login
- **SPA auth:** Google SSO + magic link (19_WebDashAuth.gs)
- **Dashboard auth toggle:** `isDashboardMemberAuthRequired()` — when enabled, all dashboard pages require member login
- **Auth config:** `ScriptProperties` (no manual setup required — `initWebDashboardAuth()` handles first-time)

### HTML Serving Method
As of v4.13.0, the build outputs individual `.html` files to `dist/`, enabling both `HtmlService.createHtmlOutput()` (inline strings) and `createTemplateFromFile()`/`createHtmlOutputFromFile()` (file-based). The SPA modules (19-25) use file-based HTML. Legacy modules (04-13) still use inline HTML strings.

---

## 📊 SHEET STRUCTURE

### Member Directory (31 columns)
- Headers read dynamically via `MEMBER_HEADER_MAP_` → `MEMBER_COLS`
- Key columns: Email, Role, Name, Unit, Phone, Is Steward

### Grievance Log (28 columns)
- Headers read dynamically via `GRIEVANCE_HEADER_MAP_` → `GRIEVANCE_COLS`
- Key columns: Grievance ID, Member Email, Status, Current Step, Deadline, Assigned Steward

### Config Tab
- Single source of truth for org-specific settings
- Columns built via `CONFIG_HEADER_MAP_` → `CONFIG_COLS`
- Row 1 = section headers, Row 2 = column headers, Data starts at row 3
- **See CLAUDE.md for detailed write-path rules**

### Resources Sheet (12 columns, v4.11.0)
- Headers via `RESOURCES_HEADER_MAP_` → `RESOURCES_COLS`
- Auto-creates with 8 starter articles on first access

### Notifications Sheet (12 columns, v4.12.0)
```
Notification ID | Recipient | Type | Title | Message | Priority
Sent By | Sent By Name | Created Date | Expires Date | Dismissed By | Status
```
- Auto-ID: NOTIF-XXX
- Types: Steward Message, Announcement, Deadline, System
- Priority: Normal, Urgent
- Dismissed_By: comma-separated emails (per-member tracking)
- Active until: Expires Date passes, member dismisses, or steward archives

### Hidden Sheets (v4.12.2+)
- `_Weekly_Questions` — weekly engagement questions
- `_Contact_Log` — steward-member interaction log (8 columns)
- `_Steward_Tasks` — steward task management (10 columns)
- `_QA_Forum` — Q&A Forum questions and answers (v4.22.6)
- `_Timeline` — Timeline activity feed entries (v4.22.9)
- `_Surveys` — Dynamic survey schema and responses (v4.23.0)

### Architectural Rule: No Formulas in Visible Sheets
All visible sheets (Dashboard, Member Satisfaction, Feedback) contain only VALUES, never formulas. Data is recomputed by JavaScript on each change. No broken formula references, no circular dependencies, no formula chains.

---

## ⚠️ PROTECTED CODE — DO NOT MODIFY

The following code sections are USER APPROVED and must NOT be modified or removed without explicit user permission:

### Dashboard Modal Popup — `src/04c_InteractiveDashboard.gs`
Protected functions: `showInteractiveDashboardTab()`, `getInteractiveDashboardHtml()`, `getInteractiveOverviewData()`, `getInteractiveGrievanceData()`, `getInteractiveMemberData()`, `getMyStewardCases()`
Tabs: Overview, Grievances, Members, Analytics, My Cases (steward-only)

### Member Satisfaction Dashboard — `src/04c_InteractiveDashboard.gs`
Protected functions: `showSatisfactionDashboard()`, `getSatisfactionDashboardHtml()`, `getSatisfactionOverviewData()`, `getSatisfactionResponseData()`, `getSatisfactionSectionData()`, `getSatisfactionAnalyticsData()`, `getSatisfactionTrendData()`, `getSatisfactionLocationDrill()`
Tabs: Overview, Trends, Responses, Sections, Insights

### Secure Member Dashboard — `src/11_CommandHub.gs`
Protected functions: `showPublicMemberDashboard()`, `showStewardPerformanceModal()`, `safetyValveScrub()`, `getSecureGrievanceStats_()`, `getSecureAllStewards_()`, `getSecureSatisfactionStats_()`, `getStewardWorkload()`
Features: Material Design UI, Weingarten Rights, live steward search, PII scrubbing

---

## 🐛 CONSOLIDATED ERRORS & FIXES LOG

| Date | Error | Cause | Fix | Status |
|------|-------|-------|-----|--------|
| 2026-03-06 | Add-to-Calendar URL malformed for events with ms ≠ .000 | `.replace('.000Z','Z')` is a string replace, not regex — silently fails for `.100Z`, `.500Z` etc | Changed to `.replace(/\.\d+Z$/,'Z')` regex in home widget + Events page (2 locations in member_view.html) | ✅ |
| 2026-03-06 | Typo in Calendar ID shows "No upcoming events" (wrong error) | `CalendarApp.getCalendarById()` returns null → was returning `[]` same as empty calendar | Returns `{_calNotFound:true}` sentinel; frontend shows diagnostic message with fix instructions | ✅ |
| 2026-03-06 | Add-to-Calendar creates event with no description | `&details=` param omitted from Google Calendar URL builder | Added `(ev.description ? '&details='+encodeURIComponent(ev.description) : '')` in both locations | ✅ |
| 2026-03-06 | Home widget KPI shows "undefined" + forEach crash when calendar sentinel returned | `events.length` on a plain object → `undefined`; `if (!events \|\| events.length === 0)` passes for object; sentinel cached via DataCache.set | Added `Array.isArray(events)` guard before render and before `DataCache.set` | ✅ |
| 2026-03-06 | Events dues-gated when it should be auth-only | `_isDuesPaying()` check on `renderEventsPage` blocked non-dues members | Gate removed; any authenticated session can view Events. `locked:true` removed from More menu item | ✅ |
| 2026-02-22 | `clasp push` → "must include manifest" | `appsscript.json` not in `dist/` | Copied manifest to `dist/` | ✅ |
| 2026-02-22 | `clasp push` → "User has not enabled Apps Script API" | API disabled | Enabled at script.google.com/home/usersettings | ✅ |
| 2026-02-23 | GitHub Actions CI not triggering | Workflow triggers `main` (lowercase) but branch is `Main` | Updated `.github/workflows/build.yml` | ✅ |
| 2026-02-23 | GitHub token `ghp_FTE8...` expired | Token rotated | User generated new token | ✅ |
| 2026-02-23 | Token `ghp_q3Zd...` lacked `repo` scope | Wrong scope selected | User generated third token with correct scope | ✅ |
| 2026-02-25 | Memory had DDS default branch as `staging` | Incorrect memory entry | Corrected to `Main` | ✅ |
| 2026-02-25 | Expired token `ghp_FTE8...` still in memory | Token rotated but memory stale | Updated memory to `ghp_7MY0...` | ✅ |

---

## 🔑 DESIGN DECISIONS LOG

Records **why** architectural choices were made, so future LLMs don't undo them.

| Date | Decision | Reasoning |
|------|----------|-----------|
| 2026-02-24 | Resources page uses Fraunces serif font | Conveys authority/trust — it's a union tool, not a SaaS product |
| 2026-02-24 | Navy + earth tones (#1e3a5f, #fafaf9) | Avoid generic AI purple gradients |
| 2026-02-24 | Check-in page uses green theme | Differentiate from other pages visually |
| 2026-02-24 | Notifications in separate sheet (not Member Directory column) | Notifications are ephemeral; don't pollute member data |
| 2026-02-24 | Dismissed_By as comma-separated in single cell | Avoids per-member rows, scales to thousands |
| 2026-02-24 | Notification auto-ID scans existing IDs for max number | Gap-safe even if rows deleted |
| 2026-02-25 | ConfigReader reads column-based Config (not row-based key-value) | Aligns with existing CONFIG_COLS system; old reader was incompatible |
| 2026-02-25 | Default accent hue changed from 250 (blue) → 30 (amber) | Warm palette matches union identity, distinguishes from generic dashboards |
| 2026-02-25 | SPA deep-links (?page=X → initialTab) with standalone HTML fallback | Consistent SPA experience, but graceful degradation if SPA unavailable |
| 2026-02-25 | `initWebDashboardAuth()` auto-configures on first run | No manual ScriptProperties setup required — reduces deployment friction |
| 2026-02-25 | Switched from consolidated single-file build to individual-file build | GAS needs separate `.html` files for `createTemplateFromFile()` and `createHtmlOutputFromFile()`. Individual files also easier to debug in GAS editor. |
| 2026-02-25 | Added `25_WorkloadService.gs` alongside `18_WorkloadTracker.gs` | 25_ is SPA-integrated (SSO auth), 18_ was standalone portal (PIN auth). 18_ later removed; 25_ is the sole workload module. |

---

## 📝 AGENT ACTIVITY LOG

Records what each AI agent changed, when, and in which files.

| Date | Agent | Action | Files Changed |
|------|-------|--------|---------------|
| 2026-02-23 | Claude (claude.ai) | Deployment audit — fixed CI triggers, verified doGet, synced branches | `.github/workflows/build.yml`, `AI_REFERENCE.md`, `web-dashboard/AI_REFERENCE.md` |
| 2026-02-24 | Claude (claude.ai) | v4.11.0: Resources Hub + Meeting Check-In | `01_Core.gs`, `05_Integrations.gs`, `10b_SurveyDocSheets.gs`, `PHASE2_PLAN.md`, `dist/` |
| 2026-02-24 | Claude (claude.ai) | v4.12.0: Notifications system (sheet + API + page) | `01_Core.gs`, `05_Integrations.gs`, `10b_SurveyDocSheets.gs`, `dist/` |
| 2026-02-25 | Claude (claude.ai) | v4.12.0 continued: Notifications page + branch sync | `05_Integrations.gs`, `dist/`, `AI_REFERENCE.md` |
| 2026-02-25 | Claude (claude.ai) | Created SYNC-LOG.md, appended sync rules | `SYNC-LOG.md`, `AI_REFERENCE.md` |
| 2026-02-25 | Claude (claude.ai) | v4.12.2: SPA port — 6 new GS modules + 6 HTML files | `19-24_*.gs`, `*.html`, `dist/` |
| 2026-02-25 | Claude (claude.ai) | v4.12.2b: UT feature port — config reader, auth, deep-link routing | `01_Core.gs`, `05_Integrations.gs`, `08a_SheetSetup.gs`, `10a_SheetCreation.gs`, `19-22_*.gs`, `index.html` |
| 2026-02-25 | Claude (claude.ai) | Consolidated AI_REFERENCE.md — removed duplication with CLAUDE.md, SYNC-LOG.md, CHANGELOG.md | `AI_REFERENCE.md`, `CHANGELOG.md` |
| 2026-02-25 | Claude (claude.ai) | Merged staging→Main: v4.13.0 SPA overhaul + notification bell/EventBus + individual-file build. Synced all 3 branches. | All `src/`, `dist/`, `build.js`, `CLAUDE.md` |
| 2026-02-28 | Claude Code (Opus 4.6) | v4.18.1-security: Full security assessment + 7 remediations — auth default ON, magic link rate limiting, token cleanup trigger, timing attack fix, PIN token migration to PropertiesService, innerHTML→textContent, escapeForFormula | `00_Security.gs`, `19_WebDashAuth.gs`, `13_MemberSelfService.gs`, `14_MeetingCheckIn.gs`, `21_WebDashDataService.gs`, `CODE_REVIEW.md`, `CHANGELOG.md`, `FEATURES.md` |

---

## 🚀 PARKED FEATURES (ranked by priority)

1. Bulk actions (flag/email/export)
2. Deadline calendar view
3. Grievance history for members
4. Welcome/landing page
5. Events page with Join Virtual button

See `PHASE2_PLAN.md` for details.

---

## 📝 NOTES FOR FUTURE LLMs

1. **Read CLAUDE.md first** — it has the most critical rules including column constant patterns, config write paths, and security patterns.
2. **Read this file second** — it has architecture context, error history, and protected code.
3. **Read SYNC-LOG.md if touching UT** — full exclusion registry with line numbers.
4. **The `web-dashboard/` folder is dead code.** Do not deploy or integrate it.
5. **Never edit `dist/` files directly.** Edit `src/*.gs` and run `npm run build` (copies individual files to dist).
6. **Test with `npm run ci`** before pushing.
7. **Deploy with `npm run deploy`** (lint + test + prod build + clasp push).
8. **Build is 39 `.gs` + 8 `.html` individual files in `dist/`.** GAS has a 6MB per-file limit. Monitor individual file sizes.
9. **The `doGet()` source is in `src/05_Integrations.gs`** (routes) and `src/22_WebDashApp.gs` (SPA entry).
10. **Do not duplicate information across reference docs.** Each doc has one canonical purpose (see table at top).

---

## 2026-02-28 — Branch Cleanup & Version Alignment (v4.18.1)

### Actions Taken
- **Deleted stale branches:** `dev` (1 behind Main, 0 ahead), `staging` (1 behind Main, 0 ahead). No unique content lost.
- **Version alignment:** `package.json` updated from 4.10.0 → 4.18.1 to match `VERSION_INFO` in `01_Core.gs`.
- **CLAUDE.md updated:** Replaced multi-branch workflow with single-branch `Main` policy. Added mandatory version tagging, parity enforcement, and no-assumptions policy.
- **Sync flow simplified:** `DDS Main → UT Main` (direct, no intermediate staging).

### Errors Found & Fixed
- **package.json drift:** Was stuck at 4.10.0 while code was at 4.18.1. Root cause: version bump in `01_Core.gs` and `CHANGELOG.md` but `package.json` was never updated after v4.10.0. Fixed by updating to 4.18.1.
- **Branch accumulation:** `dev` and `staging` branches existed on remote but were just behind `Main` with no unique content. Deleted to enforce single-branch policy.

### Version History
- v4.18.1 — Current. Branch cleanup, version alignment, CLAUDE.md overhaul.

### Reminders for LLMs
- **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values.
- **Single branch: Main only.** Never create dev/staging/feature branches.
- **Version bump is mandatory** on every code change: `VERSION_INFO` + `package.json` + `CHANGELOG.md` + git tag.
- **DDS Script ID must NEVER appear in Union-Tracker** (public repo).
- **Read before act.** Never assume repo state, file contents, or function behavior.

## 2026-02-28 — Workflow Correction (v4.18.2)

### Error Found & Corrected
- **Previous action (v4.18.1) incorrectly deleted UT staging and dev branches.** The user's intended workflow requires UT `staging` as the Claude-managed sync target, with `dev` and `Main` being user-managed. UT staging and dev branches were recreated from current Main (in full parity).
- **Root cause:** Claude assumed "ALL REPOS IN PARITY" meant single-branch everywhere. User clarification revealed the correct flow: `DDS Main → UT staging → [user] → UT dev → UT Main`.

### Actions Taken
- Recreated UT `staging` and `dev` branches (both starting from current Main = v4.18.1, so in parity)
- Updated CLAUDE.md in both repos with correct sync flow diagram
- Added Code Review rules: no false "FIXED" claims, full codebase pattern search, no inflated scores
- Clarified: DDS = single branch (Main only), UT = 3 branches (staging=Claude, dev/Main=user)
- Added fallback handoff protocol: timestamped notes for follow-on agents

## 2026-02-28 — Final Branch Simplification

### Actions Taken
- Deleted UT `staging` and `dev` branches (confirmed Main was not behind either — 0 behind staging, 1 ahead of dev)
- Reverted sync flow to direct: `DDS Main → UT Main`
- Updated CLAUDE.md in both repos to reflect single-branch policy on both repos

## 2026-03-05 — Dashboard Bug Fixes (v4.18.2 batch)

### Issues Fixed
1. **Contact Log default tab** — Was defaulting to 'log'. Changed to 'recent'. Tab order changed to: Recent → By Member → Log.
   - File: `src/steward_view.html`, function `renderContactLog()`

2. **Member View switch for stewards** — Was gated behind `IS_DUAL_ROLE` check (only stewards with role='both' in sheet). Removed gate — all stewards can now switch to member view.
   - File: `src/steward_view.html`, function `_stewardHeader()`
   - Note: Sidebar (desktop) switch in `src/index.html` still uses `IS_DUAL_ROLE` — consistent behavior on desktop.

3. **Org Overdue / Org Due <7d showing zero** — Two bugs:
   a. `_buildGrievanceRecord()` was auto-setting status to 'overdue' when `status !== 'resolved'` — but 'won', 'denied', 'settled', 'withdrawn', 'closed' are also terminal states. Fixed to check against full TERMINAL_STATUSES array.
   b. `getGrievanceStats()` had a `total < 3` early-return threshold — removed (now works with 1+ case).
   - File: `src/21_WebDashDataService.gs`

4. **Events tab dead** — Root cause: `getUpcomingEvents()` silently returned empty array when `calendarId` not set in Config. Now returns `{ _notConfigured: true }` sentinel. Frontend shows "Calendar not connected — admin needs to set Calendar ID in Config tab."
   - Files: `src/21_WebDashDataService.gs`, `src/member_view.html`

5. **Steward Directory sorting** — Was purely alphabetical (server-side). Now client-side smart sort:
   a. Stewards at same work location as the viewing steward appear first
   b. Then stewards in office today (based on officeDays field matching today's day name)
   c. Then alphabetical
   - File: `src/steward_view.html`, function `renderList()` in `renderStewardDirectoryPage()`
   - File: `src/21_WebDashDataService.gs`, `getStewardDirectory()` — removed server-side sort

6. **More menu reorder** — Q&A Forum moved to immediately after Resources. Org Chart added as new menu item after Q&A Forum.
   - File: `src/steward_view.html`, function `renderStewardMore()`

7. **Org Chart width** — Added viewport-fit wrapper, zoom controls (−/+/Fit/Reset), and auto-fit on load/resize via JS `ocZoomFit()`. Chart now scales to fit available width automatically.
   - File: `src/org_chart.html`

8. **Timeline/Weekly Questions empty states** — Improved empty state messages to explain WHY there's no content (not just "No events") and point to how to fix it.
   - File: `src/member_view.html`

### Known Remaining Issues (not fixed in this batch)
- Weekly Questions and Timeline pages appear "dead" to users — they ARE functional but depend on data being present in their respective sheets. Sheet setup instructions needed.
- Events tab also depends on `calendarId` being configured in Config tab.
- Desktop sidebar (index.html) member-view switch still gated by IS_DUAL_ROLE — intentional for now.

### Dynamic Reminder
- EVERYTHING must remain dynamic. No hardcoded sheet names, column positions, user names, or config values.
- Column identification is by header name only via `_findColumn(colMap, HEADERS.xxx)`

## 2026-03-05 — Drive Folder + Calendar Auto-Setup (v4.20.17)

### Feature: CREATE_DASHBOARD now builds full Drive folder hierarchy + Events Calendar

#### Drive Folder Structure (DashboardTest/)
```
DashboardTest/               ← PRIVATE (DriveApp.Access.PRIVATE / Permission.NONE)
  ├── Grievances/            ← all individual grievance case folders live here
  ├── Resources/
  ├── Minutes/
  └── Event Check-In/
```
- All folder IDs written to Config sheet row 3 and Script Properties on creation.
- `getOrCreateRootFolder()` now returns `DashboardTest/Grievances/` — all new case folders go inside it.
- `_getOrCreateNamedFolder_()` is the shared idempotent helper: tries stored ID → name search → creates.
- `_writeConfigFolderId_()` writes a folder ID to a Config column, skips safely if col is 0.
- `SETUP_DRIVE_FOLDERS()` — public wrapper for manual/menu re-run.

#### Config Sheet Columns Added
| Key | Header |
|-----|--------|
| DASHBOARD_ROOT_FOLDER_ID | Dashboard Root Folder ID |
| GRIEVANCES_FOLDER_ID | Grievances Folder ID |
| RESOURCES_FOLDER_ID | Resources Folder ID |
| MINUTES_FOLDER_ID | Minutes Folder ID |
| EVENT_CHECKIN_FOLDER_ID | Event Check-In Folder ID |

#### Calendar Setup
- `setupDashboardCalendar()` creates `<OrgName> Events` calendar (name derived from Config ORG_NAME).
- Calendar ID written to Config row 3 at `CONFIG_COLS.CALENDAR_ID`.
- `SETUP_CALENDAR()` — public wrapper for manual/menu re-run.
- Both functions are idempotent: find existing by stored ID → name search → create.

#### CREATE_DASHBOARD Flow (v4.20.17+)
1. `setupDashboardDriveFolders()` — Drive folder tree
2. `setupDashboardCalendar()` — Events calendar
3. `createConfigSheet()` ... (all prior sheet setup)
4. `WeeklyQuestions.initWeeklyQuestionSheets()` — Weekly Q sheets
5. `TimelineService.initTimelineSheet()` — Timeline sheet
... etc.

#### Menu Additions
- `📁 Google Drive > 🏗️ Setup Dashboard Folder Structure` → `SETUP_DRIVE_FOLDERS()`
- `📅 Calendar & Meetings > 🏗️ Setup Union Events Calendar` → `SETUP_CALENDAR()`

#### Why DashboardTest is PRIVATE
- Setting `DriveApp.Access.PRIVATE + Permission.NONE` means only the script owner
  and explicitly added editors/viewers can access it.
- Users cannot navigate to or discover any files outside the shared folders/subfolders
  they've been explicitly granted access to.

### Dynamic Reminder
- DRIVE_CONFIG.ROOT_FOLDER_NAME = 'DashboardTest' — do not hardcode elsewhere.
- All folder IDs must be read from Config sheet or Script Properties, never hardcoded.

## 2026-03-05 — Dynamic Drive Root + Folder Integrations (v4.20.18)

### Change: ROOT_FOLDER_NAME is now dynamic
- `DRIVE_CONFIG.ROOT_FOLDER_NAME` removed. Use `getDriveRootFolderName_()` everywhere.
- `getDriveRootFolderName_()` reads `Config!row3[CONFIG_COLS.ORG_NAME]` + appends `" Dashboard"`.
- Memoized per execution (`_cachedDriveRootName_`). Falls back to `DRIVE_CONFIG.ROOT_FOLDER_FALLBACK`.

### Feature: Minutes → Drive Doc
- `addMeetingMinutes()` (`21_WebDashDataService.gs`) creates a Google Doc in `Minutes/` folder.
- Doc is created in root, then moved into `Minutes/` folder (using stored `MINUTES_FOLDER_ID`).
- URL stored as col 8 (`DriveDocUrl`) in MeetingMinutes sheet.
- `getMeetingMinutes()` returns `driveDocUrl` field (empty string for rows without it — backward compatible).
- Steward minutes cards show "📄 Open Google Doc" link when `driveDocUrl` is present and valid HTTPS.

### Feature: Attendance → Drive Doc on Meeting Close
- `updateMeetingStatuses()` (`14_MeetingCheckIn.gs`) calls `saveAttendanceToDriveFolder_(meetingId, metaRow)` when a meeting reaches COMPLETED status.
- `saveAttendanceToDriveFolder_()` reads `EVENT_CHECKIN_FOLDER_ID` from Config → Script Properties.
- Creates a Google Doc with full attendee list, moves it into `Event Check-In/` folder.
- Fails silently — Drive errors never interrupt check-in functionality.

### Feature: Resources/ folder linked in Manage tab
- `getWebAppResourceLinks()` now resolves `resourcesFolderUrl` from `RESOURCES_FOLDER_ID` Config col.
- Steward view Resources > Manage tab shows a blue banner: "📂 Open Resources/ Drive Folder" when configured.
- This tells stewards where to upload files before adding their URLs as resources.

### Reminder: All folder IDs must be read from Config sheet or Script Properties
- Never hardcode folder IDs, folder names, or calendar IDs.

## 2026-03-05 — Minutes Backfill + Multi-Select Default-On + Portal Col Tests (v4.20.19)

### Feature: BACKFILL_MINUTES_DRIVE_DOCS()
- Scans all MeetingMinutes rows with empty PORTAL_MINUTES_COLS.DRIVE_DOC_URL (col 8).
- Creates a formatted Google Doc per row, moves it to Minutes/ folder, writes URL back to sheet.
- Max 50 rows/run to stay within GAS 6-min timeout. Re-runnable — skips rows with existing URL.
- Menu: `📅 Calendar & Meetings > 📄 Backfill Minutes → Drive Docs`

### Fix: PORTAL_MINUTES_COLS.DRIVE_DOC_URL = 7
- Constant added to `src/23_PortalSheets.gs`.
- Hardcoded `data[i][7]` in getMeetingMinutes() replaced with PORTAL_MINUTES_COLS.DRIVE_DOC_URL.
- Rule: Never use raw integer literals for array column access — always use the _COLS constant.

### Change: Auto Multi-Select is ON by default
- Previously: required `multiSelectAutoOpen === 'true'` (opt-in). Fresh installs got nothing.
- Now: requires `multiSelectAutoOpen === 'false'` to disable (opt-out). Default is ON.
- `removeMultiSelectTrigger()` → sets 'false' (disables).
- `installMultiSelectTrigger()` → deletes 'false' (re-enables to default-on).
- `onSelectionChange()` logic: `if (autoOpen === 'false') return;`

### Tests Added (TestSuite — 6 new)
| Test | What it catches |
|------|----------------|
| test_PortalMinutesCols_Complete | Missing DRIVE_DOC_URL, wrong offset |
| test_PortalColsNoHardcodedIndices | Duplicate col indices (copy-paste errors) |
| test_MultiSelectCols_Populated | Empty arrays, non-number col, missing label |
| test_MultiSelectAutoOpen_DefaultOn | Regression to opt-in behavior |
| test_DriveRootFolderName_Dynamic | ROOT_FOLDER_NAME reintroduced as static string |
| test_ConfigCols_FolderIds_Exist | Missing folder ID config cols after schema change |

## 2026-03-05 — Backfill Progress Flush + MegaSurvey Col Test (v4.20.20)

### Change: BACKFILL_MINUTES_DRIVE_DOCS — full loop with progress
- REMOVED: hard `LIMIT = 50` cap that forced multiple manual re-runs.
- ADDED: pre-scan to count rows needing docs; opening toast shows total.
- ADDED: `SpreadsheetApp.flush()` every 10 docs (`FLUSH_EVERY` constant). Commits writes mid-loop so GAS 6-min timeout preserves partial progress.
- ADDED: final `SpreadsheetApp.flush()` after loop before result dialog.
- Re-run is safe and idempotent — rows with a URL in PORTAL_MINUTES_COLS.DRIVE_DOC_URL are skipped.
- Toast fallback for non-UI (direct script execution) context.

### Change: test_PortalColsNoHardcodedIndices — covers all 7 PORTAL_*_COLS
- Added `PORTAL_MEGA_SURVEY_COLS` to colObjects array.
- Now validates all 7 portal column map objects for: non-negative numbers only, no duplicate indices.

## 2026-03-05 — End-to-End Function Audit + Critical Survey Bug Fix

### Audit Scope
Full systematic audit of all 95 `google.script.run` functions called from frontend HTML files:
- `index.html` (routing + auth)
- `auth_view.html` (magic link)
- `steward_view.html` (50 server calls)
- `member_view.html` (56 server calls)

### Finding 1: CRITICAL — 08e_SurveyEngine.gs Missing from build.js (BUG)
**Severity: Critical — Entire survey flow broken in dist**
- `08e_SurveyEngine.gs` existed in `src/` but was NOT in `BUILD_ORDER` in `build.js`.
- It was never copied to `dist/` on `node build.js`, meaning it was never deployed.
- **Functions that failed at runtime:**
  - `dataGetPendingSurveyMembers()` — steward survey tracking panel
  - `dataGetSatisfactionSummary()` — steward satisfaction overview
  - `dataOpenNewSurveyPeriod()` — steward opens new survey
  - `dataGetSurveyPeriod()` — period info wrapper
  - `getSurveyPeriod()` — helper called by `getSurveyQuestions()` and `submitSurveyResponse()` in 08c → the entire member survey form broke
- **Fix:** Added `'08e_SurveyEngine.gs'` to `BUILD_ORDER` between `08d_AuditAndFormulas.gs` and `09_Dashboards.gs` in `build.js`.
- **Verification:** `node build.js` now produces `dist/08e_SurveyEngine.gs` ✅

### Finding 2: Duplicate Function Definitions in 21_WebDashDataService.gs (BUG)
- `dataGetSurveyQuestions` appeared at lines 2738 AND 2760 (identical behavior).
- `dataSubmitSurveyResponse` appeared at lines 2740 AND 2761 (identical behavior).
- In GAS, last definition wins — functionally harmless but risk of future divergence.
- **Fix:** Removed second copies (lines 2760–2761); added comment pointing to canonical location.

### Finding 3: dataGetEngagementStats / dataGetWorkloadSummaryStats (Design Note)
- These two functions only read from `SEEDED_UNION_STATS` script property.
- Returns `null` if property is not set (only set via `SEED_SAMPLE_DATA`).
- Frontend handles `null` gracefully (shows "unavailable" message).
- **Not a bug** — intended for anonymized aggregate display; live data path is an open design question.

### Audit Result: All 95 public functions confirmed present in backend (✅ 0 missing signatures)

### REMINDER — ALWAYS DYNAMIC
- Sheet names, column indices, Config tab values: NEVER hardcode. Always pull from Config.
- `08e_SurveyEngine.gs` must always be included in build.js BUILD_ORDER.

## 2026-03-05 — Live Engagement Stats + Survey Triggers (v4.22.0)

### Change: dataGetEngagementStats — SEEDED stub → live sheet reads
- **Before:** Read from `SEEDED_UNION_STATS` script property; returned null on fresh deploys.
- **After:** Reads directly from live sheets every call.
- **Sources per metric:**
  | Metric | Source |
  |--------|--------|
  | surveyParticipation | `_Survey_Tracking` col 5 (CURRENT_STATUS='Completed') / active members |
  | weeklyQuestionVotes | row count of `_Weekly_Responses` |
  | eventAttendance | unique email count in `Meeting Check-In Log` col 7 (EMAIL) |
  | grievanceFilingRate | unique member emails in `Grievance Log` / total active members |
  | stewardContactRate | unique member emails in `_Contact_Log` col 2 / total active members |
  | resourceDownloads | 0 (not tracked — reserved) |
  | membershipTrends | Member Directory HIRE_DATE grouped into last 6 calendar months |
- **Active member definition:** Dues Status not blank and not 'inactive'.
- Returns `null` if Member Directory has zero active members.
- Each metric wrapped in its own try/catch; partial data is returned on per-metric failure.

### Change: dataGetWorkloadSummaryStats — SEEDED stub → live Workload Vault reads
- **Before:** Read from `SEEDED_UNION_STATS` script property.
- **After:** Reads `Workload Vault` sheet directly.
- **Calculations:**
  | Field | Logic |
  |-------|-------|
  | avgCaseload | avg PRIORITY_CASES from most-recent submission per steward |
  | highCaseloadPct | % stewards with PRIORITY_CASES > 5 |
  | submissionRate | % of Member Directory IS_STEWARD='Yes' who have submitted |
  | trendDirection | avg PRIORITY_CASES last 4 weeks vs prior 4 weeks; ±0.5 threshold |
- Returns `null` if Workload Vault is empty.

### Change: Survey triggers — added combined installer + menu items
- **New function:** `menuInstallSurveyTriggers()` in `08e_SurveyEngine.gs`
  - Calls `setupQuarterlyTrigger()` + `setupWeeklyReminderTrigger()` in one shot.
  - Shows confirmation dialog listing both triggers' schedules.
- **Menu locations:**
  - Union Hub → 📋 Survey Engine → ✅ Install ALL Survey Triggers
  - Union Hub → 🏗️ Setup → ✅ Install ALL Survey Triggers
- **⚠️ REQUIRED ONE-TIME ACTION:** After next deploy, run `menuInstallSurveyTriggers` once from the Spreadsheet menu to activate quarterly auto-open and weekly reminder emails.

### REMINDER — ALWAYS DYNAMIC
- Engagement stats and workload stats read column indices by header name at runtime.
- Never hardcode column numbers or sheet names in these functions.

## 2026-03-05 — Auth Bug Fix: Member Profile Self-Service (v4.22.1)

### Bug: dataGetFullProfile — members could never view their own profile settings page
- **Root cause:** Wrapper used `_requireStewardAuth()` — any non-steward call returned 403.
- **Impact:** Member Settings page (`renderMemberSettings`) loaded infinitely / errored silently.
- **Fix:** Caller resolved via `_resolveCallerEmail()`. Stewards can pass any email; members are locked to their own `caller` email regardless of the `email` argument.

### Bug: dataUpdateProfile — members could never save profile changes
- **Root cause:** Wrapper used `_requireStewardAuth()` — member self-service blocked.
- **Impact:** "Save Changes" in Member Settings always returned `{ success: false, message: 'Steward access required.' }`.
- **Fix:** Auth now uses `_resolveCallerEmail()` for both roles. Members always update their own record. Stewards can target another member by passing `updates._targetEmail` (stripped before forwarding to `updateMemberProfile`).
- **Field security unchanged:** `updateMemberProfile()` still only allows `street`, `city`, `state`, `zip`, `workLocation`, `officeDays` — no role escalation or PII exposure possible.

### Audit note: Full function audit complete — all 95 public functions verified end-to-end

## 2026-03-05 — Unified Polls System (v4.23.0)

### Change: Merged "Weekly Questions" + "Polls" into single unified Polls system

**Why merged:** The two systems were functionally redundant — both collected community input on a recurring basis. Polls adds multiple-choice structure; Weekly Questions adds anonymity and community sourcing. The merged system takes the best of both.

**Architecture:**
- Two polls active per week: one **Steward Poll** (steward-created), one **Community Poll** (random draw from member-submitted pool)
- **All votes are anonymous** — only SHA-256 hashed email stored; `myVote` is never returned to any client; only aggregate counts and percentages
- **Custom options** per poll (2–5 choices defined at creation time)
- **Community pool drawn automatically** — Monday time trigger (`autoSelectCommunityPoll`) picks randomly; stewards cannot select or approve pool entries
- **Poll guide** shown inline at question creation (both member and steward): neutral framing, single concept, balanced options, examples, anti-patterns

**Backend changes (24_WeeklyQuestions.gs):**
- `_Weekly_Questions` schema updated: added `Options` column (JSON array), changed `Source` values to `'steward'` | `'community'`
- `setStewardQuestion(email, text, options[])` — replaces current week's steward poll; validates options
- `submitPoolQuestion(email, text, options[])` — member submits to pool with options; email hashed
- `closePoll(stewardEmail, pollId)` — deactivates any poll by ID
- `getPoolCount()` — returns pending pool count (no question text exposed to stewards)
- `selectRandomPoolQuestion()` — random draw; marks pool entry as 'used'; called by trigger only
- `getHistory()` — returns options[] in each record now
- Deprecated: `Portal Polls` sheet path (`dataGetActivePolls`, `dataAddPoll`, `dataSubmitPollVote` return stubs)

**Frontend changes:**
- `renderWeeklyQuestionsPage()` removed — replaced by `renderPollsPage()` (member)
- `renderManagePolls()` rewritten with three sub-tabs: This Week | Create Poll | History
- `renderPollsPage()` (member) has three sub-tabs: This Week | Past Polls | Submit a Poll
- `_renderPollGuide()` and `_buildOptionInputs()` — shared helpers in member_view.html (same JS scope as steward_view.html via index.html include)
- `weeklyq` deep-link redirected to `polls` for backwards compatibility
- All votes show results as horizontal bar chart by percentage with vote count

**Menu/trigger changes (03_UIComponents.gs):**
- Removed "Weekly Questions" from both More menus
- Added `🗳️ Polls` submenu: Install Community Poll Draw Trigger | Draw Now (manual) | Initialize Poll Sheets
- Added `⏱️ Install Community Poll Draw Trigger` to Setup menu

**One-time setup required after deploy:**
Run `Union Hub → Polls → Install Community Poll Draw Trigger` once to activate automatic Monday draws.
Existing `_Weekly_Responses` data is compatible — no migration needed (column structure unchanged).

**Invariant:** Stewards see pool count only, never individual pool submissions. No mechanism exists to pick a specific member's question.

## 2026-03-05 — Org Chart + Poll Frequency (v4.23.1)

### Change 1: Org Chart added to member More menu
- Member More menu now includes `🏛 Org Chart` → calls `_handleTabNav('member', 'orgchart')`
- Routing was already shared (`if (tabId === 'orgchart') renderOrgChart(app, role)` runs before role split)
- Deep-link `?page=orgchart` works for members as it did for stewards

### Change 2: Steward-configurable poll frequency
**Storage:** `PropertiesService.getScriptProperties().getProperty('POLL_FREQUENCY')`
Valid values: `'weekly'` (default) | `'biweekly'` | `'monthly'`

**Backend (24_WeeklyQuestions.gs):**
- `_getFrequency()` — reads ScriptProperty, falls back to `'weekly'`
- `_getPeriodStart(date)` — replaces `_getWeekStart(date)`:
  - weekly → Monday of current week
  - biweekly → Monday of even ISO week number (weeks align to 2,4,6…)
  - monthly → 1st of current month
- `_periodKey(date)` — replaces `_weekKey(date)` — ISO date string of period start
- All internal period comparisons now use `thisPeriod = _periodKey()`
- `getPollFrequency()` / `setPollFrequency(stewardEmail, freq)` — public API
- `wqGetPollFrequency()` / `wqSetPollFrequency(null, freq)` — server wrappers
- `autoSelectCommunityPoll()` — skips draw on off-weeks (biweekly = odd ISO weeks, monthly = non-1st days)

**Frontend:**
- Steward "Create Poll" tab: frequency selector card (3 toggle buttons, highlights current setting)
- Member "This Week" tab: empty-state copy reads actual frequency ("each Monday" / "every other Monday" / "on the 1st of each month")

**No migration needed:** existing `_Weekly_Questions` rows have `Week Start` dates; `_periodKey()` comparisons still match those ISO date strings for weekly cadence.

## 2026-03-05 — Contact Tab Redesign + Syntax Bugfix (v4.23.2)

### Bug fixed: Orphan braces from v4.23.0 cleanup
Two stray `}` braces were left in member_view.html (lines ~2864, ~2873) from the _renderQuestionSubmission removal during the Polls merge. These caused `Unexpected token 'function'` at runtime, breaking the entire member JS. Fixed by removing the dead comment block they were attached to.

### Contact tab redesign (member)
`renderStewardContact` now serves as the single Contact page — assigned steward card at top, full directory below.

**Structure:**
1. **Your [Steward] section** — if assigned: name/avatar, email (mailto), phone (tel), in-office days, vCard download, Change button. If not assigned: prompt card pointing down to directory.
2. **All [Steward]s section** — full filterable directory with:
   - Search input (name or location)
   - Location filter pills (top 5 locations, All default)
   - Smart sort: same-location first → in-office today → alphabetical
   - Per-card: office days with live "● In today" indicator, email, phone, Save Contact (vCard), Select as my [Steward]

**Functions changed:**
- `renderStewardContact(appContainer)` — fully rewritten
- `_renderStewardPicker(content)` — removed, replaced by `_showStewardPicker(targetSection)` (inline, no page reload)
- `_renderMemberDirectory(target)` — new shared helper for the directory section
- `renderStewardDirectory(appContainer)` — still exists (used by steward view); member More menu now routes to renderStewardContact

**Routing:**
- `case 'contact'` bottom nav → `renderStewardContact` (was `renderStewardDirectory`)
- More menu "Steward Directory" → `renderStewardContact` (was `renderStewardDirectory`)
- `?page=stewarddirectory` deep-link still routes to `renderStewardDirectory` (steward-specific page unchanged)

---

## 2026-03-06 — Cases System Bug Fixes (v4.18.x patch)

### Bugs Fixed

**BUG-CASES-01 — getStewardCases: Zero cases returned (CRITICAL)**
- Root cause: "Assigned Steward" column in Grievance Log stores names (e.g. "Jane Smith"), not emails. Prior code compared `assignedTo === stewardEmail` — always false.
- Fix: Dual-match logic. Resolves steward's display name from Member Directory via `findUserByEmail`, then matches `assignedTo` against BOTH the email AND the name (case-insensitive). File: `21_WebDashDataService.gs → getStewardCases`.

**BUG-CASES-02 — role 'both' blocked from steward functions (CRITICAL)**
- Root cause 1: `getUserRole_()` in `00_Security.gs` never returned `'both'` — only checked `IS_STEWARD` column, not `Role` column.
- Root cause 2: `checkWebAppAuthorization('steward')` condition excluded `role === 'both'`.
- Fix: `getUserRole_` now checks `Role` column for `'both'`/`'steward/member'` before checking `IS_STEWARD`. Auth check now accepts `'both'` as steward-authorized.

**BUG-CASES-03 — sendBroadcastMessage always returns Unauthorized**
- Root cause: Typo — checked `auth.authorized` (undefined), should be `auth.isAuthorized`.
- Fix: Property name corrected in `21_WebDashDataService.gs → sendBroadcastMessage`.

**BUG-CASES-04 — startGrievanceDraft uses hardcoded column indices (violates dynamic rule)**
- Root cause: Used `MEMBER_COLS.EMAIL - 1`, `GRIEVANCE_COLS.GRIEVANCE_ID - 1`, etc. directly instead of dynamic lookup.
- Fix: Fully rewritten using `_buildColumnMap` + `_findColumn` matching the rest of DataService. Writes First Name and Last Name separately (not a combined Member Name) matching the Grievance Log schema.

**BUG-CASES-05 — Active filter drops legitimate open cases**
- Root cause: `filter === 'active'` only matched `status === 'active' || status === 'new'`. Statuses like `'pending'`, `'open'`, `'step i'`, `'in progress'` were silently dropped.
- Fix: Inverted logic — active = NOT in TERMINAL_STATUSES AND NOT overdue. Terminal list: `['resolved', 'won', 'denied', 'settled', 'withdrawn', 'closed']`. File: `steward_view.html → renderCaseList`.

**BUG-CASES-06 — Case cards show email username instead of member name**
- Root cause: `_buildGrievanceRecord` did not include member name. Card rendered `email.split('@')[0]`.
- Fix: Added `grievanceMemberFirstName`/`grievanceMemberLastName` aliases to `HEADERS`. `_buildGrievanceRecord` now returns `memberName` (First + Last). Card uses `c.memberName` with email-username fallback.

**BUG-CASES-07 — grievanceSteward alias order caused column miss**
- Root cause: First alias was `'steward'` — matches generic column. `'assigned steward'` (actual header) should be first.
- Fix: Reordered alias list: `['assigned steward', 'steward', 'steward email', 'assigned to']`. Also corrected `grievanceUnit` to include `'work location'` and `'location'` as aliases matching actual sheet header.

### Rule reminder
- Everything dynamic. NEVER use `MEMBER_COLS.X - 1` or `GRIEVANCE_COLS.X - 1` in `21_WebDashDataService.gs`. Use `_findColumn(colMap, HEADERS.*)` exclusively.
- `checkWebAppAuthorization('steward')` returns `isAuthorized` (not `authorized`).
- "Assigned Steward" in Grievance Log = person's name. Always dual-match by name AND email.

---

## 2026-03-06 — Cases follow-up fixes (patch 2)

**BUG-CASES-08 — createGrievanceDriveFolder used hardcoded MEMBER_COLS / GRIEVANCE_COLS**
- Fix: Fully rewritten using `_buildColumnMap` + `_findColumn`. Lookup chain: email match first (most reliable), memberId match as secondary. No hardcoded indices.
- Bonus: previously returned "Member not found" if memberId was blank — now proceeds via email match alone, so members without a memberId in the directory still work.

**BUG-CASES-09 — dataGetStewardCases / dataGetStewardKPIs used weak auth**
- Root cause: `_resolveCallerEmail()` lets any authenticated user call steward endpoints.
- Fix: Both wrappers now use `_requireStewardAuth()` — requires `role === 'steward'`, `'both'`, or `'admin'`. Returns `[]` / `{}` for unauthorized callers instead of empty-but-successful result.

---

## 2026-03-06 — Cases follow-up fixes (patch 3)

**BUG-CASES-10 — dataGetBatchData trusted client-supplied role**
- Risk: a member could pass `role='steward'` and receive the steward batch payload (cases, KPIs, member list).
- Fix: `ignoredRole` parameter is discarded. Role is re-derived server-side via `DataService.getUserRole(email)`. `'both'` and `'admin'` normalize to `'steward'` view.

**SCAN — MEMBER_COLS / GRIEVANCE_COLS in 21_WebDashDataService.gs**
- Full grep scan returned zero results. No remaining hardcoded column indices in DataService.

---

## 2026-03-06 — Eager batch counts + nav badges

**FEAT: _getStewardBatchData — memberCount, taskCount, overdueTaskCount**
- Added three lightweight counts to the steward batch payload.
- `memberCount` — from `getStewardMembers` (falls back to `getAllMembers` if none assigned). Member Directory already cached from `getStewardCases` call so no extra sheet read.
- `taskCount` — open tasks only from `getTasks(email, 'open')`.
- `overdueTaskCount` — derived from same open task array (dueDays < 0); no second call.
- These are stored in `AppState` on batch init, so nav and More screen render counts immediately without any additional server round trips.

**FEAT: Nav badge on Cases tab — overdue count (red)**
- Renders when `AppState.kpis.overdue > 0`. Caps at "9+" display.

**FEAT: Nav badge on Members tab — member count (accent)**
- Renders when `AppState.memberCount > 0`. Caps at "99+" display.

**FEAT: More menu Tasks entry — overdue badge (red) or open count (accent)**
- Overdue tasks take priority: shows red "N overdue" pill.
- If no overdue but open tasks exist: shows accent "N open" pill.
- All three badges render from eagerly-loaded AppState — zero additional server calls.

---

## 2026-03-06 — Fallback path badge parity + CSS confirmation

**FIX: Fallback fetch chain now populates memberCount / taskCount / overdueTaskCount**
- Two parallel non-blocking `serverCall()` chains added alongside the existing cases/KPIs chain.
- `dataGetStewardMembers` → `AppState.memberCount`
- `dataGetTasks(null)` → derives `AppState.taskCount` + `AppState.overdueTaskCount` from open tasks
- Failure handlers zero the counts so badges stay hidden rather than stale.
- Counts arrive after initial render (async), but badges appear on next nav redraw (tab switch or page refresh). This is acceptable — the batch path is the normal path; fallback is a degraded state.

**CSS CONFIRMED: .more-menu-item already has display:flex + align-items:center**
- Badge pill aligns correctly without any CSS change.
- `flex: 1` added to text div in previous commit ensures badge pushes flush right.

---

## 2026-03-06 — Badge refresh + fallback wrapper fix (patch 5)

**BUG: fallback path called non-existent dataGetStewardMembers**
- Correct wrapper is `dataGetAllMembers`. Fixed in steward_view.html fallback chain.

**FEAT: _refreshNavBadges() — live badge updates within session**
- New function in steward_view.html. Fires two parallel non-blocking serverCalls:
  1. `dataGetStewardKPIs` → updates `AppState.kpis` (case overdue count)
  2. `dataGetTasks(null)` → updates `AppState.taskCount` + `AppState.overdueTaskCount`
- After task counts arrive, swaps the existing `.bottom-nav` DOM node in place — no full page reload.
- Uses `AppState.activeTab` to highlight the correct tab after swap.

**FEAT: AppState.activeTab tracking**
- `_handleTabNav` in index.html now sets `AppState.activeTab = tabId` on every navigation. Required by `_refreshNavBadges` to re-render nav with correct active state.

**FEAT: Badge refresh wired to all mutation points**
- Task complete (`dataCompleteTask`) → calls `_refreshNavBadges()` before re-rendering tasks page.
- Task create — all three paths (`dataCreateTask`, `dataCreateMemberTask`, `dataCreateTaskForSteward`) → call `_refreshNavBadges()` on success.
- All calls use `typeof _refreshNavBadges === 'function'` guard for safety.

---

## 2026-03-06 — Badge refresh hardening (patch 6)

**FIX: _refreshNavBadges querySelector scoped to #app**
- Was: `document.querySelector('.bottom-nav')` — would match first nav on page globally.
- Now: `(document.getElementById('app') || document).querySelector('.bottom-nav')` — scoped to the app container, safe if any nav-like element exists outside #app.

**FEAT: _refreshNavBadges wired to member task completion (member_view.html)**
- `dataCompleteMemberTask` success handler now calls `_refreshNavBadges()` before re-rendering.
- `_refreshNavBadges` is defined in steward_view.html. On the member view the function won't exist (members don't load steward_view.html), so the `typeof` guard prevents any error.
- This keeps the steward's task badges accurate when a member completes an assigned task in the same session — without requiring the steward to manually refresh.

---

## 2026-03-06 — Badge refresh debounce + CLAUDE.md standing rule (patch 7)

**FIX: _refreshNavBadges debounced at 600ms**
- `_refreshNavBadgesTimer` module-level var holds the pending timeout.
- Each call cancels the previous timer before setting a new one.
- 600ms chosen to exceed typical GAS round-trip so back-to-back mutations collapse.
- No behavior change for single mutations — fires as before.

**DOCS: CLAUDE.md standing rule added**
- Section: "Standing Rule — Badge Refresh on Case/Task Mutations"
- Documents which mutation types require `_refreshNavBadges()`, the `typeof` guard pattern, and the debounce guarantee.
- Explicitly calls out the future case-status-change UI as a required wire point.

---

## 2026-03-06 — Insights caching + auth fixes + duplicate removal (patch 8)

**FEAT: Insights page caching (AppState.insightsCache, 5-min TTL)**
- On first visit: shows spinner, fires all four calls, stores result in AppState.insightsCache with timestamp.
- On revisit within 5 min: renders instantly from cache, zero server calls.
- On revisit after 5 min (stale): renders stale data immediately (no spinner), fires all four calls in background, replaces content when they resolve.
- On server failure: falls back to stale cache value rather than showing blank/empty.
- `_renderInsightsContent()` extracted as standalone function to support cache-render and fresh-render paths.

**BUG-INSIGHTS-01: dataGetGrievanceStats / dataGetGrievanceHotSpots / dataGetSatisfactionTrends had no auth**
- Any authenticated user (member) could call these and receive grievance statistics, hot spot locations, satisfaction trends, and steward performance data.
- Fixed: `dataGetGrievanceStats` and `dataGetGrievanceHotSpots` now call `_requireStewardAuth()`.
- Fixed: `dataGetSatisfactionTrends` now calls `_requireStewardAuth()`. Returns `{ categories: [] }` for unauthorized callers.

**BUG-INSIGHTS-02: Duplicate wrapper — dataGetAgencyGrievanceStats**
- `dataGetAgencyGrievanceStats` and `dataGetGrievanceStats` both called `DataService.getGrievanceStats()`.
- Alias removed. Frontend updated to call `dataGetGrievanceStats` directly.

---

## 2026-03-06 — Dynamic TTL + dataGetMembershipStats auth (patch 9)

**FEAT: Insights cache TTL now driven by CONFIG**
- `CONFIG.insightsCacheTTLMin` added to ConfigReader (default: 5 minutes).
- New header: `'Insights Cache TTL (Minutes)'` in Config tab → `CONFIG_COLS.INSIGHTS_CACHE_TTL_MIN`.
- Frontend reads `(CONFIG.insightsCacheTTLMin || 5) * 60 * 1000` — hardcoded `5 * 60 * 1000` removed.
- Admins can change the TTL in the Config sheet without a code deploy.
- Falls back to 5 min if the Config tab cell is blank or missing.

**BUG-AUTH-01: dataGetMembershipStats had no auth**
- Any unauthenticated caller could retrieve org membership statistics.
- Members legitimately need this data (used on member union-stats page), so steward-only auth would break them.
- Fixed: `_resolveCallerEmail()` required — returns `null` for unauthenticated callers. Member view already handles null response gracefully via failure handler.

---

## 2026-03-06 — dataGetUpcomingEvents auth + sheet setup auto-seeding (patch 10)

**BUG-AUTH-02: dataGetUpcomingEvents had no auth**
- Unauthenticated callers could read org calendar events.
- Members and stewards both use this, so minimum auth (_resolveCallerEmail) applied — same pattern as dataGetMembershipStats.
- Returns [] for unauthenticated callers. Member view DataCache handles empty gracefully.

**FIX: Insights Cache TTL now auto-seeded by createConfigSheet**
- CONFIG_HEADER_MAP_ already had INSIGHTS_CACHE_TTL_MIN added in patch 9 — headers were always auto-derived from the map, so the column header writes automatically.
- Added seedConfigDefault_(sheet, CONFIG_COLS.INSIGHTS_CACHE_TTL_MIN, [5], isExistingSheet) to 10a_SheetCreation.gs.
- Re-running CREATE_DASHBOARD on existing sheets will add the column header and seed the value 5 without overwriting user changes.
- No manual Config tab editing required.

---

## 2026-03-06 — Full data* wrapper auth audit (patch 11)

**AUDIT: All function data* wrappers scanned — 7 unguarded found and fixed**

All fixed with `_resolveCallerEmail()` minimum auth (both steward and member roles use these):

| Wrapper | Return on unauth |
|---|---|
| dataGetStewardDirectory | [] |
| dataGetCaseChecklist | [] |
| dataGetSurveyQuestions | [] |
| dataGetSurveyPeriod | null |
| dataGetMeetingMinutes | [] |
| dataGetEngagementStats | null |
| dataGetWorkloadSummaryStats | null |

**CONFIRMED CLEAN:** dataGetFullProfile, dataUpdateProfile, dataCreateTaskForSteward, dataGetBatchData, dataGetBroadcastFilterOptions, dataGetWelcomeData, dataMarkWelcomeDismissed — all verified guarded internally.

**STANDING RULE (addendum to CLAUDE.md):** Every new `function data*` wrapper MUST include either `_resolveCallerEmail()` or `_requireStewardAuth()` as its first statement. No exceptions.

---

## 2026-03-06 — Full task system fixes (patch 13)

**BUG-TASKS-01/06: _Steward_Tasks caching**
- getTasks, getMemberTasks, getStewardAssignedMemberTasks use _getCachedSheetData. Same-execution + 2-min cross-request cache.
- All write functions (createTask, completeTask, updateTask, createMemberTask, completeMemberTask, stewardCompleteMemberTask) call _invalidateSheetCache on write.
- getTasks now excludes member tasks (col 11 = member) from steward list.

**BUG-TASKS-02: dataUpdateTask wrapper + inline edit UI**
- New wrapper: dataUpdateTask(ignoredEmail, taskId, updates) with steward auth.
- updateTask backend supports dueDate changes.
- Frontend: open task cards have Edit button toggling inline form (title, priority, due date). No page reload on save.

**BUG-TASKS-03: Steward can complete member tasks**
- New backend: stewardCompleteMemberTask validates task ownership (col 11=member, col 12=stewardEmail).
- New wrapper: dataStaffCompleteMemberTask(taskId) with steward auth.
- Frontend: Member Tasks sub-tab shows Complete for member button on open tasks. Confirm dialog guards it. Audit logged.

**BUG-TASKS-04: IIFE closures for task loop captures**
- Complete and edit button callbacks now use IIFEs to correctly capture task reference in loop.

**BUG-TASKS-05: dataCreateTask exposes assignToEmail**
- Wrapper extended with optional assignToEmail param passed through to backend.

---

## 2026-03-06 — Contact log writeback + member card actions (v4.20.21)

### Problem fixed: two contact tracking systems were not synced
`_Contact_Log` sheet (rolling history, written by web dashboard) and Member Directory snapshot columns (`Recent Contact Date`, `Contact Steward`, `Contact Notes`) were completely independent. Dashboard KPIs and WorkloadTracker read from Member Directory snapshot only, so they showed stale data for any member whose contact was logged via the web UI.

### Fix: `logMemberContact()` writeback (`21_WebDashDataService.gs`)
After writing to `_Contact_Log`, the function now:
1. Opens `MEMBER_SHEET` (dynamic — resolved from `SHEETS.MEMBER_DIR` or `'Member Directory'`)
2. Finds all column indices by header name (never by index)
3. Locates the member's row by email match (case-insensitive)
4. Writes:
   - `Recent Contact Date` → `new Date()`
   - `Contact Steward` → steward's display name via `findUserByEmail()` (falls back to email if not found)
   - `Contact Notes` → notes text (only if notes provided; existing notes not cleared on a contact with no notes)
5. Writeback is wrapped in try/catch — non-fatal; failure logged to Apps Script Logger only

### New backend wrappers
- `dataSendDirectMessage(ignoredEmail, memberEmail, subject, body)` — steward-only. Sends email via MailApp. Prefixes subject with `orgAbbrev`. Fires `DIRECT_MESSAGE_SENT` audit event.
- `dataGetMemberCaseFolderUrl(ignoredEmail, memberEmail)` — steward-only. Returns `{ success, url, grievanceId }`. Prefers active (non-resolved/closed/withdrawn/denied) grievance. Falls back to most recent grievance. Returns `{ success: false, url: null, message }` if no Drive folder URL stored in Grievance Log.

### Member card changes (`steward_view.html` — `_showMemberDetail`)
Added shared `actionArea` div below the button row. All 4 buttons render inline forms into `actionArea` (clears previous form on each click):
- **Full Profile** — unchanged behavior (loads extra profile fields inline)
- **📞 Log Contact** — inline form: contact type pills (Phone/Email/In Person/Text) + notes textarea. Calls `dataLogMemberContact`. Shows ✓ confirmation on success.
- **🔔 Send Notification** — inline form: subject input + body textarea. Calls `dataSendDirectMessage`. Shows ✓ confirmation on success.
- **📂 Case Folder** — no form. Calls `dataGetMemberCaseFolderUrl`, opens Drive URL in new tab on success, or shows message in `actionArea` if no folder found.

### RULES REMINDER
- Everything dynamic — column lookups always by header name, never by index
- `Contact Steward` = display name, NOT email
- `dataSendDirectMessage` and `dataGetMemberCaseFolderUrl` both call `_requireStewardAuth()` as first statement

---

## 2026-03-06 — Per-member Drive folder architecture (v4.20.22 → v4.20.25)

### Current Drive structure (v4.20.25 — authoritative)
```
[Dashboard Root]/                        ← PRIVATE — steward/admin only
  Members/                               ← MEMBERS_SUBFOLDER; ID cached as MEMBERS_FOLDER_ID Script Property
    LastName, FirstName/                 ← per-member master admin folder — steward-only, NEVER shared with member
      Contact Log — LastName, FirstName  ← Google Sheet — steward-only, full rolling history
      Grievances/                        ← steward-only subfolder — members do NOT see this level
        GR-XXXX - YYYY-MM-DD/            ← case folder — member added as EDITOR (can upload evidence)
          Step 1 - Informal/
          Step 2 - Written/
          Step 3 - Review/
          Supporting Documents/
  Resources/
  Minutes/
  Event Check-In/
  Meeting Notes/
  Meeting Agenda/
```

### Sharing model — exact rules
| Folder | Member access | Steward access |
|---|---|---|
| `Members/` | None | Inherited from Dashboard Root |
| `Members/LastName, FirstName/` | None | Inherited |
| `Contact Log` sheet | None | Inherited |
| `Grievances/` subfolder | None | Inherited |
| `GR-XXXX` case folder | **Editor** (explicitly granted) | **Editor** (explicitly granted) |

### Key function: `getOrCreateMemberAdminFolder(memberEmail)` — `05_Integrations.gs`
- Single source of truth for all per-member Drive work
- Resolves `Members/` root via cached Script Property `MEMBERS_FOLDER_ID`, fallback to `DASHBOARD_ROOT_FOLDER_ID`, fallback to DriveApp search
- Resolves member display name from Member Directory by header-name lookup (email column → first/last name). Falls back to email prefix.
- Creates `Members/LastName, FirstName/` and `Grievances/` subfolder inside it on first call
- On first creation: writes master folder URL to `Member Admin Folder URL` column in Member Directory (by header name, never by column index)
- Returns `{ masterFolder: Folder, grievancesFolder: Folder }` or `null` on failure

### `setupDriveFolderForGrievance(grievanceId)` — `05_Integrations.gs`
- Reads `MEMBER_EMAIL` from Grievance Log (`GRIEVANCE_COLS.MEMBER_EMAIL`)
- Calls `getOrCreateMemberAdminFolder(memberEmail)` → gets `grievancesFolder`
- Creates `GR-XXXX - YYYY-MM-DD/` case folder under `grievancesFolder`
- Adds member as Editor on case folder (non-fatal if missing email — folder created, sharing skipped, warning logged)
- Adds assigned steward as Editor (non-fatal)
- Falls back to name-based member folder lookup if no email available

### `setupDriveFolderForMember(memberId)` — `05_Integrations.gs`
- **Shim only** — resolves email from memberId, delegates to `getOrCreateMemberAdminFolder()`
- Kept for backward compatibility with "Create Member Folder" button in `03_UIComponents.gs`

### `logMemberContact()` execution order (4 steps) — `21_WebDashDataService.gs`
1. Append row to `_Contact_Log` hidden sheet (fast — must not fail)
2. Resolve steward display name via `findUserByEmail()` (reused in steps 3+4)
3. Writeback to Member Directory snapshot columns — non-fatal try/catch
4. Append to per-member Drive contact sheet via `getOrCreateMemberAdminFolder()` + `getOrCreateMemberContactSheet_()` — non-fatal

### Member Directory column
- Header: `'Member Admin Folder URL'`
- HEADERS alias: `memberAdminFolderUrl: ['member admin folder url', 'member admin folder', 'contact log folder url']`
  - `'contact log folder url'` alias retained for backward compat with existing Member Directory data
- Surfaced in `_buildUserRecord()` → `user.memberAdminFolderUrl`
- Surfaced in `getFullMemberProfile()` → `profile.memberAdminFolderUrl`
- Steward-visible only (Full Profile section → `📂 Member Folder` link). Never expose to member-facing view.

### RULES — do not violate
- Never share `Members/`, `Members/Name/`, `Contact Log`, or `Grievances/` with the member — these are internal steward workspace
- Only the individual grievance case folder (`GR-XXXX`) is shared with the member, as Editor
- `MEMBERS_SUBFOLDER = 'Members'` — do NOT rename this folder in Drive; the stored Script Property ID will go stale
- `Member Admin Folder URL` column is populated by code (non-fatal). Never hardcode the URL.
- All folder/column lookups must be by header name, never by column index

### dataSendDirectMessage Drive logging — `21_WebDashDataService.gs`
- After email sent: appends `[Date, stewardName, 'Email', 'Subject: X | bodyPreview', '']` to Drive contact sheet
- Body preview capped at 300 chars
- Uses `DataService.getOrCreateMemberContactFolderPublic` + `getOrCreateMemberContactSheetPublic` (private helpers exposed on DataService return object for cross-IIFE access)
- Non-fatal — email send failure returns error; Drive failure is logged and swallowed

### RULE: public exposure of private helpers
Private helpers needed outside the DataService IIFE must be exposed on the return object with a `Public` suffix. Do NOT make them top-level functions.

### Migration note (existing deployments)
Old `Member Contacts/` subfolder and its per-member folders are **not deleted**. Old contact log sheets remain readable in Drive but will no longer receive new writes. New contacts/grievances write to the new `Members/` hierarchy.

## 2026-03-06 — Notification System Overhaul (v4.22.0)

### Bugs Fixed
1. **Sort bug** (`05_Integrations.gs` `getWebAppNotifications()`): `results.reverse()` was called AFTER `sort()`, undoing the Urgent-first ordering. Now: `reverse()` first (newest at index 0), then `sort()` for Urgent promotion.
2. **Manage tab wrong data** (`steward_view.html` `_renderNotifManage()`): was calling `getWebAppNotifications(steward_email, 'steward')` which returns only what the steward received. Now calls new `getAllWebAppNotifications()` which returns every row in the sheet.
3. **Bell badge DOM stale after dismiss** (`member_view.html`): `AppState.notificationCount` was decremented but DOM badge element not updated. Fixed: after card removal, find `.notif-bell-badge` and update or remove it.

### Features Added
- **`DISMISS_MODE` column** (col 13, `Dismissible` | `Timed`) in `NOTIFICATIONS_HEADER_MAP_` (`01_Core.gs`)
  - `Dismissible`: member can permanently dismiss (writes to `Dismissed_By` column via `dismissWebAppNotification()`)
  - `Timed`: notification shows until `Expires_Date`, no dismiss button, "Auto-expires" badge shown
- **Compose form dismiss mode toggle** (`steward_view.html`): two buttons (✕ Dismissible / ⏰ Timed). Timed enforces expiry date.
- **Permanent member dismiss** (`member_view.html`): replaced 1-hour localStorage TTL with `dismissWebAppNotification()` backend call.
- **`getAllWebAppNotifications()`** (`05_Integrations.gs`): steward-auth-gated, returns all rows with `dismissedCount`, `status`, `dismissMode`. Used by Manage tab.

### Dead Code Removed
- `getUserNotifications()` — ScriptProperties-based, orphaned since v4.13.0 Notifications sheet
- `markNotificationRead()` — same
- `broadcastStewardNotification()` — no active callers
- `getWebAppNotificationsHtml()` — 345 lines of standalone HTML, never routed in `doGet()`
- `pushNotification()` rerouted to Notifications sheet (still called by `saveSharedView()`)

### Files Changed
- `src/01_Core.gs` — NOTIFICATIONS_HEADER_MAP_, v4.22.0 changelog entry
- `src/05_Integrations.gs` — sort fix, dismissMode in results, sendWebAppNotification, getAllWebAppNotifications, dead HTML removed
- `src/10b_SurveyDocSheets.gs` — createNotificationsSheet: DISMISS_MODE col width, validation, starter rows
- `src/16_DashboardEnhancements.gs` — dead functions removed, pushNotification rerouted
- `src/steward_view.html` — Manage tab fix, compose dismiss mode toggle
- `src/member_view.html` — permanent dismiss, Timed mode badge, bell badge DOM fix

## 2026-03-06 — Notification Manage Hardening (v4.22.1)

### Features Added
1. **`MIGRATE_ADD_DISMISS_MODE_COLUMN()`** (`05_Integrations.gs`)
   - One-time migration for existing Notifications sheets
   - Checks for column before touching anything (safe to re-run)
   - Appends `Dismiss_Mode` as next column after current last col
   - Backfills all existing rows with `'Dismissible'` (safe legacy default)
   - Shows UI alert on success/skip; logs to Apps Script Logger
   - **Run from Apps Script editor → can be deleted after confirmed success**

2. **`archiveWebAppNotification(notificationId)`** (`05_Integrations.gs`)
   - Steward-auth-gated
   - Sets `Status = 'Archived'` on matching row — non-destructive, row preserved
   - Archived rows excluded from all member views automatically (getWebAppNotifications filters `Status = 'Active'`)

3. **Archive button in Manage tab** (`steward_view.html`)
   - Shown only on `Active` rows
   - Calls `archiveWebAppNotification()` on click
   - On success: updates status pill in-place, removes itself — no full re-render needed
   - Expired/Archived rows remain visible for audit purposes

### Files Changed
- `src/05_Integrations.gs` — MIGRATE_ADD_DISMISS_MODE_COLUMN, archiveWebAppNotification
- `src/steward_view.html` — Archive button in _renderNotifManage
- `src/01_Core.gs` — v4.22.1 changelog entry

---

## 2026-03-06 — Subject line, scope config, dues gate on 6 tabs

### Broadcast: Custom subject line
- `dataSendBroadcast(ignoredEmail, filter, msg, subject)` — 4th param added
- `sendBroadcastMessage(stewardEmail, filter, message, customSubject)` — if customSubject non-empty, uses it; else falls back to auto-generated `orgAbbrev + ' - Message from your ' + stewardLabel`
- UI: subject input field above message textarea; placeholder shows auto-generated default

### Broadcast: Scope toggle via Config (not hardcoded)
- New Config tab column: `Broadcast: Allow All Members Scope` (key: `BROADCAST_SCOPE_ALL`)
- Set value to `yes` in Config tab to show the My Members / All Members toggle in Broadcast UI
- Leave blank or set to anything else to hide the toggle (steward can only send to their assigned members)
- `dataGetBroadcastFilterOptions()` now returns `broadcastScopeAll: boolean`
- `_sanitizeConfig()` now includes `broadcastScopeAll: boolean` (22_WebDashApp.gs)
- RULE: scope toggle visibility must always come from config — never hardcode as always-visible

### Dues paying: null = paying fix
- `filter.duesPaying === 'paying'` now only blocks members where `duesPaying === false` (not null)
- `filter.duesPaying === 'nonpaying'` only includes members where `duesPaying === false`
- null means column is absent — always treated as dues paying (benefit of the doubt)

### safeUser: duesPaying added
- `_serveDashboard()` now includes `duesPaying` in `safeUser` sent to client
- Value: true=paying, false=not paying, null=column absent (22_WebDashApp.gs)
- CURRENT_USER.duesPaying is available client-side in all view scripts

### Dues gate: 6 member tabs blocked for non-dues-paying members
**`member_view.html`**
- `_isDuesPaying()` → returns `CURRENT_USER.duesPaying !== false` (null = paying)
- `_renderDuesGate(appContainer, featureName, backFn)` → standard locked-feature wall with 🔒 icon
- Gated functions (each checks `_isDuesPaying()` at entry, returns gate if false):
  - `renderMemberResources` — Member Resources tab
  - `renderSurveyFormPage` — Quarterly Survey
  - `renderEventsPage` — Events (member role only; stewards bypass the check)
  - `renderUnionStatsPage` — Union Stats / Insights
  - `renderPollsPage` — Polls (includes weekly questions)

**`index.html`**
- `_handleTabNav` org chart case: checks role === 'member' && !_isDuesPaying() before rendering
- Uses typeof guards for _isDuesPaying/_renderDuesGate to prevent errors if member_view not loaded

### RULES
- Stewards are NEVER gated by dues status — checks are role-gated (role === 'member' only)
- Dues gate functions must remain in member_view.html (not index.html) so they have access to member-specific layout/header helpers
- The org chart gate in index.html uses typeof guards since it lives outside member_view scope
- EVERYTHING IS DYNAMIC: dues gate checks CURRENT_USER.duesPaying which comes from live sheet data

---

## 2026-03-06 — Survey banner gate, More menu lock icons, broadcast scope Config seeding

### Survey banner (member_view.html — renderMemberHome)
- Banner condition changed from `status && !status.hasCompleted` to `status && !status.hasCompleted && _isDuesPaying()`
- Non-dues-paying members no longer see the survey banner on the home page
- Consistent with the gate on renderSurveyFormPage itself

### More menu lock indicators (member_view.html — renderMemberMore)
- `isDuesPaying` computed once at the top of renderMemberMore (not per item)
- Menu item schema extended: optional `locked: true` flag on gated items
- Gated items: Events, Union Stats, Org Chart, Resources, Polls, Take Survey
- When `isLocked = m.locked && !isDuesPaying`:
  - Item renders at 60% opacity
  - 🔒 badge (10px, absolute-positioned) overlaid on bottom-right of icon
  - "DUES REQUIRED" amber pill appended next to label text
  - Click still routes through the normal action → hits the gate wall inside the render function
- Non-locked items render identically to before (no style changes)
- RULE: locked flag list must stay in sync with the 6 functions that have `_isDuesPaying()` entry guards

### Broadcast scope Config seeding (10a_SheetCreation.gs)
- `seedConfigDefault_(sheet, CONFIG_COLS.BROADCAST_SCOPE_ALL, ['no'], isExistingSheet)` — default 'no'
- `SpreadsheetApp.newDataValidation().requireValueInList(['yes', 'no'])` applied to row 3 of that column
  — admins get a dropdown picker, not a free-text cell
- Help text: "yes = stewards can broadcast to all members. no = stewards can only broadcast to their assigned members."
- `'── BROADCAST ──'` section header added to sectionHeaders array
- CONFIG_COLS.BROADCAST_SCOPE_ALL is only defined after buildColsFromMap_ runs — validation block wrapped in `if (CONFIG_COLS.BROADCAST_SCOPE_ALL)` guard

### ConfigReader (20_WebDashConfigReader.gs)
- Added `broadcastScopeAll: _readCell(sheet, CONFIG_COLS.BROADCAST_SCOPE_ALL) || 'no'`
- config.broadcastScopeAll now resolves from live Config sheet data
- _sanitizeConfig (22_WebDashApp.gs) converts it to boolean: `String(...).toLowerCase() === 'yes'`
- dataGetBroadcastFilterOptions (21_WebDashDataService.gs) reads it and returns as `broadcastScopeAll: boolean`
- UI reads opts.broadcastScopeAll to conditionally show scope toggle

### REMINDER — everything must be dynamic
- Broadcast scope is controlled by Config tab — never hardcode as always-visible
- Lock icon visibility is driven by CURRENT_USER.duesPaying (from live sheet at login)
- Dues gate checks CURRENT_USER.duesPaying which comes from _serveDashboard safeUser

---

## 2026-03-06 — Dues gate: home page banner + Member Directory checkbox column

### Persistent home page banner (member_view.html — renderMemberHome)
- Inserted synchronously after the welcome slot — no server call, uses CURRENT_USER.duesPaying
- Condition: `!_isDuesPaying()` (false = not paying, null = column absent → not shown)
- Amber color scheme (rgba orange background + border) to match lock icons in More menu
- Lists all 6 gated features explicitly: Events, Resources, Polls, Survey, Org Chart, Union Stats
- Includes "Contact your steward if you believe this is an error" so member has a clear path
- RULE: banner must stay in sync with the locked:true items in renderMemberMore and the
  _isDuesPaying() guards in individual render functions

### Member Directory: Dues Paying column (10a_SheetCreation.gs — createMemberDirectory)
- `sheet.getRange(2, MEMBER_COLS.DUES_PAYING, 4999, 1).insertCheckboxes()` — checked = paying
- Column width set to 100px
- Guarded with `if (MEMBER_COLS.DUES_PAYING)` — safe to run on existing sheets without the column
- Conditional formatting added to existing rules array:
  - TRUE (checked) → green background (#e8f5e9) + dark green text (#2e7d32)
  - FALSE (unchecked) → amber background (#fff8e1) + dark amber text (#f57f17)
  - Unchecked-but-empty cells get no color (checkbox default = FALSE in Sheets,
    so all pre-existing rows will appear amber until explicitly checked)
- IMPORTANT: After running CREATE_DASHBOARD on an existing sheet, all members will appear
  as "not dues paying" until manually checked. Recommend bulk-checking all current paying
  members after column is added.

### REMINDER — everything must be dynamic
- Dues Paying column detected at login via _buildUserRecord (reads HEADERS.memberDuesPaying)
- CURRENT_USER.duesPaying flows: sheet → _buildUserRecord → _serveDashboard safeUser → client
- Home banner, More menu locks, and gate walls all read CURRENT_USER.duesPaying — one source

---

## 2026-03-06 — Member Directory column auto-migration

### _addMissingMemberHeaders_(sheet) — new function (10a_SheetCreation.gs)
- Reads current row-1 headers from the live sheet (case-insensitive match)
- Compares against `getMemberHeaders()` (derived from MEMBER_HEADER_MAP_)
- Appends any missing headers into the next available columns, styled with purple bg/white bold
- Per-column post-setup block inside the loop handles column-specific formatting:
  - `'dues paying'` → `insertCheckboxes()` + width 100px + green/amber conditional format rules
  - Extend this block for any future column that needs special formatting on first-add
- Returns array of added header names (empty if nothing added)
- Never deletes, moves, or overwrites existing columns or data

### createMemberDirectory() migration path
- `getLastRow() <= 1` → new sheet: writes all headers, Dues Paying setup done inline
  (MEMBER_COLS.DUES_PAYING not yet resolved → resolve from `headers.indexOf('Dues Paying')`)
  Rules stored in `sheet._duesCfRules`, merged into final `setConditionalFormatRules()` call
- `getLastRow() > 1` → existing sheet: calls `_addMissingMemberHeaders_()`, shows toast
- Toast message lists added column names so admins know what changed

### Pattern rule for future columns
To add a new column to the Member Directory that auto-migrates to existing sheets:
1. Add `{ key: 'YOUR_KEY', header: 'Your Header' }` to `MEMBER_HEADER_MAP_` in `01_Core.gs`
2. If column needs special formatting (checkbox, dropdown, date format, etc.), add a case
   to the `if (normalised === '...')` block inside `_addMissingMemberHeaders_()` in `10a_SheetCreation.gs`
3. For new sheets, add the same setup after the header write block in `createMemberDirectory()`
   using index-based column resolution (not MEMBER_COLS.YOUR_KEY — not yet resolved)
4. syncColumnMaps() will pick up the new column automatically at next onOpen/execution

### REMINDER — everything must be dynamic
- Column detection always uses header name matching, never hardcoded column index
- MEMBER_COLS constants are resolved at runtime by syncColumnMaps/resolveColumnsFromSheet_
- _addMissingMemberHeaders_ follows the same header-name matching pattern

---

## 2026-03-06 — Grievance Log column auto-migration

### _addMissingGrievanceHeaders_(sheet) — new function (10a_SheetCreation.gs)
- Exact parallel to _addMissingMemberHeaders_ — same logic, same safety rules
- Compares row-1 headers against getGrievanceHeaders() (from GRIEVANCE_HEADER_MAP_)
- Appends missing headers far-right into next empty columns with purple/white bold styling
- Per-column post-setup block included but empty today — MESSAGE_ALERT and QUICK_ACTIONS
  already exist in all live sheets. Add cases here when new Grievance columns need
  special formatting (checkboxes, date format, conditional rules, etc.)
- Returns array of added header names (empty if nothing added)

### createGrievanceLog() migration path
- getLastRow() <= 1 → new sheet: writes all headers (unchanged)
- getLastRow() > 1 → existing sheet: calls _addMissingGrievanceHeaders_(), shows toast
- Toast message lists added column names with 📋 Grievance Log label

### Pattern rule for future Grievance columns
1. Add { key: 'YOUR_KEY', header: 'Your Header' } to GRIEVANCE_HEADER_MAP_ in 01_Core.gs
2. Add per-column formatting case to the post-setup block in _addMissingGrievanceHeaders_()
3. For new-sheet path: add same setup in createGrievanceLog() after header write,
   using index-based resolution (GRIEVANCE_COLS.YOUR_KEY not yet resolved on new sheets)

### Both sheets now covered
- Member Directory → _addMissingMemberHeaders_()
- Grievance Log    → _addMissingGrievanceHeaders_()
- Both functions follow identical contract: header-name match only, no data mutation,
  per-column hook, toast on change, Logger output

---

## 📋 CHANGE LOG — v4.22.x Resources Tab Overhaul

### Changes made (session: 2026-03-06)

**Issues fixed:**
1. ✅ Dues gate removed from `renderMemberResources()` — labor rights content (Weingarten, Just Cause, etc.) must be accessible to all bargaining unit members regardless of dues status. Union-Tracker parity restored.
2. ✅ Category mismatch resolved — steward form dropdown was using a completely different and incompatible list vs. the 📚 Resources sheet data validation.
3. ✅ Category list is now fully dynamic — driven from new `📚 Resource Config` sheet, never hardcoded in client code.
4. ✅ Steward manage list now sorted by Sort Order — previously showed items in insertion (sheet row) order.

**Files changed:**
- `src/01_Core.gs` — Added `SHEETS.RESOURCE_CONFIG`, `RESOURCE_CONFIG_HEADER_MAP_`, `RESOURCE_CONFIG_COLS`, and registered in dynamic column refresh registry.
- `src/10b_SurveyDocSheets.gs` — Added `createResourceConfigSheet()` with 10 default categories matching 📚 Resources sheet validation.
- `src/05_Integrations.gs` — Added `getWebAppResourceCategories()` (reads from 📚 Resource Config, auto-creates if missing, falls back to `_defaultResourceCategories_()` on error).
- `src/member_view.html` — Removed `_isDuesPaying()` gate from `renderMemberResources()`.
- `src/steward_view.html` — Replaced hardcoded category dropdown with live `getWebAppResourceCategories()` call; added sort on manage list.

**New sheet: 📚 Resource Config**
| Column | Purpose |
|---|---|
| Setting | Row type — currently only 'Category' |
| Value | The category name |
| Sort Order | Display order (numeric) |
| Active | Yes / No — hides a category without deleting it |
| Notes | Steward notes (not shown in UI) |

**Default categories (in order):**
Contract Article → Know Your Rights → Grievance Process → Forms & Templates → FAQ → Guide → Policy → Contact Info → Link → General

**⚠️ CRITICAL RULES (do not violate):**
- `RESOURCE_CONFIG_COLS` is built from `RESOURCE_CONFIG_HEADER_MAP_` via `buildColsFromMap_()` — never hardcode column numbers.
- `createResourceConfigSheet()` checks for existing sheet first — NEVER overwrites manually entered data.
- Categories read by UI at form-open time via `getWebAppResourceCategories()` — never cached in config bootstrap.
- `_defaultResourceCategories_()` is a last-resort fallback only — the live sheet is the source of truth.

---

## v4.22.6 — MADDS Org Chart Default (2026-03-06)

**Change summary:** Replaced `src/org_chart.html` with the MADDS (Main Internal Breakout) chart sourced from the `Woop91/509d` repository (`org-chart/MADDS.html`, last updated 2026-03-01). This is now the default Org Chart view for both DDS-Dashboard and Union-Tracker.

**Source:** `509d/org-chart/MADDS.html` — Full SEIU Local 509 organizational chart including: President → Officers → Chief of Staff → Directors → Coordinators → Public Sector Chapters (MassAbility expanded by default) → Other Chapters. Also contains Role Descriptions, Financial Overview, Staff Directory & Compensation, Career Paths, and Chapter Advisors & Internal Organizers sections.

**Embedding approach (critical — do not revert):**
The MADDS chart is a standalone HTML page. To embed it in the GAS SPA, the following conversions were made:
- `<!DOCTYPE html>`, `<html>`, `<head>`, `<body>` tags stripped; content wrapped in `<div class="madds-embed">`
- CSS `:root {}` → `.madds-embed {}` (scopes CSS variables, prevents collision with app vars)
- `html, body {}` → `.madds-embed {}`; `body {}` → `.madds-embed {}`
- `body.light` → `.madds-embed.light` (light/dark mode toggle scoped to container)
- `#mode-toggle` → `#madds-mode-toggle` (avoids ID collision with SPA)
- `toggleMode()` → `maddstoggleMode()` (avoids function name collision)
- `document.body.classList` → `document.querySelector('.madds-embed').classList`
- `position: fixed` on mode-toggle → `position: sticky` (fixed breaks in embedded context)
- Google Fonts `<link>` loaded dynamically via inline JS injector (static `<link>` tags don't load when injected via `.innerHTML`)

**No server-side changes required.** The existing `getOrgChartHtml()` in `22_WebDashApp.gs` continues to serve `org_chart` via `HtmlService.createHtmlOutputFromFile('org_chart').getContent()`. The `renderOrgChart()` function in `index.html` is unchanged.

**Files changed:**
- `src/org_chart.html` — Replaced with MADDS chart fragment (3,266 lines)
- `src/01_Core.gs` — Added v4.22.6 changelog entry; updated VERSION to "4.22.6"

**⚠️ RULES:**
- Do NOT rename `org_chart.html` — the GAS server function is hardcoded to `'org_chart'` filename.
- If updating MADDS.html from the 509d repo, re-run the scoping conversions above — do NOT paste the raw standalone HTML directly.
- Keep `.madds-embed` as the wrapper class and `#madds-mode-toggle` as the toggle ID.
- The 509d repo is the source of truth for the chart content; the converted fragment lives in DDS-Dashboard and Union-Tracker `src/`.

---

## [2026-03-07] Org Chart Sync Script Added

### Change
Added `scripts/sync-org-chart.js` — a Node.js script that fetches `MADDS.html`
from the private `Woop91/509d` repo (GitHub Contents API) and transforms it into
the SPA-compatible `src/org_chart.html` used by DDS-Dashboard.

### Wire-up
- `package.json` → `"sync-org-chart": "node scripts/sync-org-chart.js"` (manual, not pre-deploy)
- Requires `.env` at repo root with `GITHUB_509D_TOKEN=ghp_...` (gitignored)

### Problem Solved
`org_chart.html` was a manually synced copy of `509d/MADDS.html` with no automation.
Any update to the 509d chart required a manual copy + transform + redeploy.

### Transformations Applied (in order)
1. Strip leading block comment
2. Strip `<!DOCTYPE>` / `<html>` opening tags
3. Strip preamble inside `<head>` (`<meta>`, `<title>`, Google Fonts `<link>`)
   ⚠️ NOTE: In MADDS.html, `<head>` wraps the entire `<style>` block (closes AFTER `</style>`).
   Must NOT use `<head>...</head>` greedy match — would delete all CSS.
   Strip individual tags instead, then strip `</head>` separately.
4. CSS scope: `:root {` → `.madds-embed {`
5. CSS scope: `html, body {` → `.madds-embed {`
6. CSS scope: `body.light` → `.madds-embed.light` (all)
7. Replace `<body>` open + first button  **[before global #mode-toggle rename]**
8. JS fn: `toggleMode` → `maddstoggleMode`, retarget `document.body` → `.madds-embed`  **[before rename]**
9. CSS/HTML: rename all remaining `#mode-toggle` → `#madds-mode-toggle`
10. CSS: `position: fixed` → `position: sticky` in `#madds-mode-toggle` block
11. Replace closing `</script></body></html>` → `</script></div><font-loader script>`  **[before html strip]**
12. Strip residual `</html>` / `</body>`

### Verification Checks (12 checks — script aborts if any fail)
Output must NOT contain: `:root {`, `body.light`, `id="mode-toggle"`, `onclick="toggleMode`, 
`function toggleMode`, `<body>`, `<!DOCTYPE`
Output MUST contain: `<div class="madds-embed">`, `maddstoggleMode`, `DM+Serif`, `</div>`, `position: sticky`

---

## Q&A Forum — Changes (2026-03-07)

### Files Modified
- `src/26_QAForum.gs`
- `src/member_view.html`
- `dist/26_QAForum.gs` (rebuilt)
- `dist/member_view.html` (rebuilt)

### Backend Changes (`26_QAForum.gs`)
| Change | Reason |
|---|---|
| `submitAnswer()` — added steward-only guard: `if (!isSteward) return error` | Members should not be able to post answers |
| `submitQuestion()` — added `_createNotificationInternal_()` call to notify 'All Stewards' | Stewards get alerted on new unanswered questions |
| `submitAnswer()` — added `_createNotificationInternal_()` call to notify question author | Author gets alerted when their question receives an answer |
| `resolveQuestion(email, questionId, isSteward)` — new function | Either question owner OR steward can close a thread |
| `_createNotificationInternal_(recipient, type, title, message)` — new private helper | Writes to Notifications sheet without steward session requirement (system-generated) |
| `qaSubmitAnswer()` global wrapper — now calls `_requireStewardAuth()` instead of `_resolveCallerEmail()` | Enforces steward-only at API boundary, not just inside module |
| `qaResolveQuestion()` — new global wrapper | Auto-detects steward role via `checkWebAppAuthorization` |
| `resolveQuestion` added to public API return object | Expose new function |

### Frontend Changes (`member_view.html`)
| Change | Reason |
|---|---|
| Answer form wrapped in `if (isSteward)` | Members cannot post answers |
| "Mark Resolved" button: visible when `(q.isOwner || isSteward) && status !== 'resolved'` | Owner or steward can close thread |
| ✓ Resolved badge added to question detail view | Visual confirmation of resolved state |
| ✓ Resolved badge added to browse list question cards | Visible in list without opening detail |
| Upvote/Flag/Resolve buttons section uses `flexWrap: 'wrap'` | Prevents button overflow on small screens |

### Notification Schema Note
`_createNotificationInternal_()` writes 13 columns matching NOTIFICATIONS_HEADER_MAP_ order:
`ID | Recipient | Type | Title | Message | Priority | SenderEmail | SenderName | Date | ExpiresDate | DismissedBy | Status | DismissMode`
Sender email = 'system', Sender name = 'Q&A Forum'.

### Key Design Decisions
- **Steward-only answers**: Enforced at BOTH backend module level AND global wrapper level (double guard)
- **Notification routing**: New questions → 'All Stewards'; new answers → specific author email (NOT anonymous — if anonymous, the author email is still stored internally and gets notified)
- **Resolved by**: Either party — question owner or steward. No separate "accepted answer" concept; resolved = closed thread
- **`_createNotificationInternal_`** bypasses `sendWebAppNotification()` auth check intentionally — it is called in GAS execution context where the session belongs to the member/submitter, not a steward


### [2026-03-07] sync-org-chart.js v2 — Multi-repo, All Branches

Updated to commit org_chart.html to all branches of both repos:
- `DDS-Dashboard` (Main, staging): `src/org_chart.html`, `dist/org_chart.html`
- `Union-Tracker` (Main, staging): `src/org_chart.html`, `dist/org_chart.html`

MADDS is now the sole org chart. Skip logic prevents empty commits when file is unchanged.
Uses git clone to temp dir (GitHub Contents API returns 403 on write for these repos).

---

## Q&A Forum — Follow-up Changes (2026-03-07 session 2)

### Files Modified
- `src/21_WebDashDataService.gs` — added qaUnansweredCount to steward batch
- `src/steward_view.html` — badge wiring, More menu badge, _refreshNavBadges update
- `src/member_view.html` — showResolved toggle, filtering, nav badge refresh on action

### Item 1 — Anonymous poster notifications
No code change required. `submitAnswer()` reads author email from `_QA_Forum` col B
(always stored internally even for anonymous posts). Notification is sent to that email.
Anonymity only affects the *displayed* author name, not the stored email.

### Item 2 — Steward unanswered Q&A badge
| Location | What | Detail |
|---|---|---|
| `_getStewardBatchData()` | `qaUnansweredCount` field | Counts questions where answerCount === 0 AND status not in (resolved, deleted) |
| `initStewardView()` | `AppState.qaUnansweredCount` | Wired from batch.qaUnansweredCount |
| `renderBottomNav()` | Amber badge on `more_steward` tab | Amber (warning color), max display '9+' |
| `renderStewardMore()` | Amber 'N unanswered' pill on Q&A Forum row | Same amber scheme as nav badge |
| `_refreshNavBadges()` | Re-fetches qaGetQuestions(1,999) after task fetch | Recomputes count, re-renders nav so badge decrements immediately after steward answers |

### Item 3 — Show Resolved toggle
- `showResolved` boolean state variable, default `false`, scoped inside `renderQAForum()`
- Persists across page/sort changes within the same session; resets on re-navigation
- Server returns all non-deleted questions (including resolved). Filter applied client-side
- Empty state message is context-aware: if hiding resolved and no open questions remain,
  hints the user to toggle "Show Resolved"
- `_refreshNavBadges()` called on `qaSubmitAnswer` success and `qaResolveQuestion` success
  so the More tab badge and Q&A Forum row badge decrement without requiring page reload

### Design Note — Client-side vs server-side filtering for resolved
Chose client-side filtering (not a server param) because:
- Server already paginates; adding a second filter param adds complexity
- Page size is 20 and resolved threads accumulate slowly — no performance concern
- Keeps `getQuestions()` API simple and single-purpose


---

## Q&A Forum — Follow-up Changes (2026-03-07 session 2)

### Files Modified
- `src/21_WebDashDataService.gs` — added qaUnansweredCount to steward batch
- `src/steward_view.html` — badge wiring, More menu badge, _refreshNavBadges update
- `src/member_view.html` — showResolved toggle, filtering, nav badge refresh on action

### Item 1 — Anonymous poster notifications
No code change required. submitAnswer() reads author email from _QA_Forum col B
(always stored internally even for anonymous posts). Notification sent to that email.
Anonymity only affects the displayed author name, not the stored email.

### Item 2 — Steward unanswered Q&A badge
_getStewardBatchData(): qaUnansweredCount = questions where answerCount === 0 AND status not in (resolved, deleted)
AppState.qaUnansweredCount wired from batch on steward init.
renderBottomNav(): amber badge on more_steward tab, max display '9+'.
renderStewardMore(): amber 'N unanswered' pill on Q&A Forum row.
_refreshNavBadges(): re-fetches qaGetQuestions(1,999) after task fetch, recomputes, re-renders nav.

### Item 3 — Show Resolved toggle
showResolved boolean state var, default false, scoped inside renderQAForum().
Server returns all non-deleted questions including resolved. Filter is client-side.
Empty state message context-aware: hints user to toggle if open list is empty.
_refreshNavBadges() called on qaSubmitAnswer and qaResolveQuestion success for live badge decrement.

### Design Note — Client-side filtering for resolved
Chose client-side (not server param) because page size is 20 and resolved threads accumulate slowly.
Keeps getQuestions() API simple.


---

## Q&A Forum — Bell Badge (2026-03-07 session 3)

### Files Modified
- `src/index.html` — sidebar bell badge total computation
- `src/steward_view.html` — dismiss handler bell update
- `src/member_view.html` — dismiss handler comment clarification

### Bell Badge Logic
Bell total = notificationCount + qaUnansweredCount (steward only).
Member bell = notificationCount only.
The two values are intentionally kept separate in AppState so each can
be managed independently. Composite total computed only at render time.

### Steward dismiss handler
After dismissing a notification, bell badge re-renders with composite total.
Handles edge case: if badge span doesn't exist (notifs were 0 but qa count > 0),
creates and appends the badge span to .notif-bell-wrap.


---

## Steward Directory — Parity Fixes (2026-03-07)

### Files Modified
- `src/member_view.html`

### Changes Made

#### 1. Unit added to search — `_renderMemberDirectory` (line ~1068)
Search in the member-facing steward list now includes `unit` field.
Before: searched name + workLocation only.
After: searches name + workLocation + unit.
Matches parity with `renderStewardDirectoryPage` in `steward_view.html`.

#### 2. Orphaned `renderStewardDirectory` removed (line ~3206, v4.12.0)
This function was dead code — no nav path, no call site referenced it.
All member-side directory rendering routes through `renderStewardContact` → `_renderMemberDirectory`.
Replaced with tombstone comment: "renderStewardDirectory removed v4.23.3"

### Confirmed Parity: member `_renderMemberDirectory` vs steward `renderStewardDirectoryPage`
| Feature | Member | Steward |
|---------|--------|---------|
| Location pills | ✅ | ✅ |
| Smart sort (location → in-office → alpha) | ✅ | ✅ |
| Search: name | ✅ | ✅ |
| Search: location | ✅ | ✅ |
| Search: unit | ✅ (fixed) | ✅ |
| vCard download | ✅ | ✅ |
| officeDays + in-today indicator | ✅ | ✅ |
| Phone shown | ✅ (intentional) | ✅ |

### Phone Data Note
Both roles receive phone numbers from the same `getStewardDirectory()` payload.
This is intentional per user decision (2026-03-07). If this needs to change,
add role-aware filtering in `getStewardDirectory()` in `21_WebDashDataService.gs`.

---

## Steward Phone Permission — Opt-In for Member Visibility (2026-03-07)

### Files Modified
- `src/21_WebDashDataService.gs`

### Problem
Phone numbers were included in `getStewardDirectory()` unconditionally.
Both stewards and members received the same payload — members could see
every steward's phone regardless of steward preference.

### Solution: Opt-In Column + Role-Aware Redaction

#### New HEADERS alias added
```
memberSharePhone: ['share phone', 'share phone number', 'phone visible', 'public phone', 'share contact']
```
Column absent or blank → defaults to `false` (opt-in required; no accidental exposure).

#### New `sharePhone` field in `_buildUserRecord`
Truthy values: `yes`, `true`, `1`. All others → `false`.

#### `getStewardDirectory(callerIsSteward)` — updated signature
- `callerIsSteward = true` → phone always returned (stewards see all peers)
- `callerIsSteward = false` → phone only returned if `rec.sharePhone === true`

#### `dataGetStewardDirectory()` wrapper — updated
Resolves caller email → looks up their record → passes `callerIsSteward` flag.
No change to the client-side HTML — phone field is already guarded with `if (s.phone)`.

### Backward Compatibility
- Existing sheets with no "Share Phone" column → all phones hidden from members by default.
- Admins add the column and set `Yes` for stewards who consent.
- No migration needed; purely additive.

### Data Flow
```
dataGetStewardDirectory()
  → _resolveCallerEmail() → e
  → DataService.findUserByEmail(e) → callerRec
  → callerIsSteward = callerRec.isSteward
  → DataService.getStewardDirectory(callerIsSteward)
    → for each steward rec:
        phone = (callerIsSteward || rec.sharePhone) ? rec.phone : null
```

---

## Share Phone Column — Sheet Definition (2026-03-07)

### Files Modified
- `src/01_Core.gs` — MEMBER_HEADER_MAP_
- `src/10a_SheetCreation.gs` — createMemberDirectory, _addMissingMemberHeaders_

### Column Added to MEMBER_HEADER_MAP_
```
{ key: 'SHARE_PHONE', header: 'Share Phone' }
```
Position: immediately after `IS_STEWARD`. Accessible as `MEMBER_COLS.SHARE_PHONE`.

### Sheet Behavior
- **New sheets**: Yes/No dropdown validation applied at column creation with help text.
- **Existing sheets**: `_addMissingMemberHeaders_()` detects 'share phone' is missing,
  appends it, and applies the same dropdown + column width (110px) automatically.
  No manual migration required — re-running CREATE_DASHBOARD or any menu action
  that calls createMemberDirectory will add the column.

### Dropdown values
`Yes` = steward opts in; phone shown to members.
`No` (or blank/absent) = phone hidden from members.

### Column width: 110px. Not hidden. No conditional formatting needed.

---

## Share Phone Column — Sheet Definition (2026-03-07)

### Files Modified
- `src/01_Core.gs` — MEMBER_HEADER_MAP_
- `src/10a_SheetCreation.gs` — createMemberDirectory, _addMissingMemberHeaders_

### Column Added to MEMBER_HEADER_MAP_
```
{ key: 'SHARE_PHONE', header: 'Share Phone' }
```
Position: immediately after `IS_STEWARD`. Accessible as `MEMBER_COLS.SHARE_PHONE`.

### Sheet Behavior
- **New sheets**: Yes/No dropdown validation applied at column creation with help text.
- **Existing sheets**: `_addMissingMemberHeaders_()` detects 'share phone' is missing,
  appends it, and applies the same dropdown + column width (110px) automatically.
  No manual migration required — re-running CREATE_DASHBOARD or any menu action
  that calls createMemberDirectory will add the column.

### Dropdown values
`Yes` = steward opts in; phone shown to members.
`No` (or blank/absent) = phone hidden from members.

### Column width: 110px. Not hidden. No conditional formatting needed.

---

## Share Phone — Steward Self-Toggle (2026-03-07)

### Files Modified
- `src/21_WebDashDataService.gs` — updateMemberProfile, getFullMemberProfile
- `src/22_WebDashApp.gs` — _serveDashboard safeUser
- `src/steward_view.html` — renderStewardMore

### How It Works
1. `_serveDashboard`: `sharePhone` added to `safeUser` → available as `CURRENT_USER.sharePhone` on page load.
2. `updateMemberProfile`: `sharePhone` added to `editableFields` allowlist, mapped to `HEADERS.memberSharePhone`.
   Writes `'Yes'` or `'No'` string to the "Share Phone" column. Existing field allowlist protects
   all other columns from being modified by this path.
3. UI: Toggle switch card at top of `renderStewardMore()` (the "More" tab).
   - Reflects current `CURRENT_USER.sharePhone` state on render.
   - On click: optimistic UI update → `dataUpdateProfile({ sharePhone: 'Yes'|'No' })` → 
     success confirms, failure reverts visuals + shows error.
   - Updates `CURRENT_USER.sharePhone` in-memory on success (no page reload needed).
   - 3-second auto-dismiss on status message.

### Write path
```
toggleTrack click
  → dataUpdateProfile(email, { sharePhone: 'Yes'|'No' })
    → DataService.updateMemberProfile(email, { sharePhone: 'Yes'|'No' })
      → writes to 'Share Phone' column in Member Directory
```

### Security
dataUpdateProfile uses `_resolveCallerEmail()` — steward can only update their own row
unless they pass `updates._targetEmail` (admin override path). sharePhone update never
passes _targetEmail, so it always writes to the caller's own row.

---

## Share Phone — Steward Self-Toggle (2026-03-07)

### Files Modified
- `src/21_WebDashDataService.gs` — updateMemberProfile
- `src/22_WebDashApp.gs` — _serveDashboard safeUser
- `src/steward_view.html` — renderStewardMore

### How It Works
1. `_serveDashboard`: `sharePhone` added to `safeUser` → `CURRENT_USER.sharePhone` on load.
2. `updateMemberProfile`: `sharePhone` in `editableFields` → writes `'Yes'`/`'No'` to column.
3. UI: Toggle card at top of renderStewardMore. Optimistic update, success/failure feedback, 3s dismiss.

### Write path
```
toggleTrack click → dataUpdateProfile({ sharePhone: 'Yes'|'No' })
  → DataService.updateMemberProfile → 'Share Phone' column in Member Directory
```

### Security
Steward can only update their own row (no _targetEmail passed). Allowlist in
updateMemberProfile prevents any other column being modified via this path.

---

## Share Phone — Default 'No' Seeding (2026-03-07)

### File Modified
- `src/10a_SheetCreation.gs` — `_addMissingMemberHeaders_`

### Change
When the 'Share Phone' column is auto-added to an existing sheet,
all current data rows are seeded with `'No'` before the dropdown validation is applied.
This makes the default state explicit in the sheet (not blank) and ensures
admins reading the sheet see the intent clearly.

New sheets have no data rows at creation time, so seeding is not needed there.
Blank cells in the column are still treated as `false` by the backend as a belt-and-suspenders fallback.

---

## 2026-03-08 — Steward Dual-Role Toggle Fix

### Issue
Stewards could not switch to Member View. The "Switch to Member/Steward" toggle in the sidebar and mobile headers was gated behind `IS_DUAL_ROLE`, which was only `true` when `role === 'both'`. Since stewards are assigned `role = 'steward'` from the Member Directory sheet, the toggle never appeared for them.

### Root Cause
`src/22_WebDashApp.gs` line 180: `isDualRole: role === 'both'` — excluded regular stewards.

### Fix
Changed to `isDualRole: role === 'steward' || role === 'both'` — since all stewards are members, every steward is inherently dual-role.

### Files Changed
- `src/22_WebDashApp.gs` — one-line fix (isDualRole condition)

### How It Works
- **Sidebar (desktop/tablet)**: `index.html` ~line 469 — shows "Switch to [Member/Steward]" when `IS_DUAL_ROLE` is true
- **Steward mobile header**: `steward_view.html` ~line 94 — always shows switch button (not gated)
- **Member mobile header**: `member_view.html` ~line 69 — shows switch button when `IS_DUAL_ROLE` is true
- Labels come from `CONFIG.memberLabel` / `CONFIG.stewardLabel` (dynamic, from Config tab)
- Toggle calls `initStewardView(app)` or `initMemberView(app)` and updates `AppState.activeRole`
