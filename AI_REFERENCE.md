# AI REFERENCE DOCUMENT — DDS-Dashboard
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-03-09 (v4.25.5)
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
**DDS Apps Script ID:** `18hHHX-4E_ykGCqu_EDwKCwqY9ycyRgPtOmguacsxnVZ4YsRh-YETODiu`
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

---

## 📋 CHANGE LOG — v4.24.1 FlashPolls Verdict + seedWeeklyQuestions Schema Fix

**Date:** 2026-03-07
**Version:** v4.24.1
**Agent:** Claude (claude.ai)

### FlashPolls Sheets — Live Spreadsheet Verdict

**Conclusion: Sheets almost certainly exist (empty, headers only). No PII present.**

Evidence chain:
1. `FlashPolls`/`PollResponses` were introduced in v4.12.2 (2026-02-25) and created by `initPortalSheets()`.
2. `initPortalSheets()` is called from `CREATE_DASHBOARD()`, which must have been run to set up the system.
3. `seedPollsData()` (the only function that would write rows) was added in v4.18.0 (2026-02-26) as part of the Phase 3 seeder — a **separate manual function**, not called by `CREATE_DASHBOARD`.
4. There is zero evidence in CHANGELOG, AI_REFERENCE, git commit messages, or code comments that Phase 3 seeder was ever run on the live script.
5. `submitPollVote()` (the live-vote function) was **never wired to any frontend** — the webapp always used `wq*` wrappers. No real votes could have been recorded.

**Result:** `FlashPolls` and `PollResponses` contain only headers. Still delete them — orphaned infrastructure.

### Bug Fixed: seedWeeklyQuestions — Wrong Column Schema

The existing `seedWeeklyQuestions()` function had two schema bugs causing silent data corruption:

| Sheet | Bug | Fix |
|---|---|---|
| `_Question_Pool` | Missing `Options` column (wrote 5 cols, schema needs 6) | Added `JSON.stringify(options)` as col 3 |
| `_Weekly_Questions` | Missing `Options` column (wrote 7 cols, schema needs 8) | Added `JSON.stringify(options)` as col 3 |
| `_Weekly_Responses` | WQ-002 responses were numeric strings (`'8'`, `'7'`) — not valid option values | Replaced with actual option strings from the question |

**Root cause:** Seed was written before `24_WeeklyQuestions.gs` added the `Options` column to both sheets. Schema drift was never caught because seeding was never confirmed run on the live script.

**Schema (current, from 24_WeeklyQuestions.gs Q_COLS/P_COLS constants):**
- `_Weekly_Questions`: `ID | Text | Options(JSON) | Source | Submitted By | Week Start | Active | Created`
- `_Question_Pool`: `ID | Text | Options(JSON) | Submitted By Hash | Status | Created`
- `_Weekly_Responses`: `ID | Question ID | Email Hash | Response | Timestamp`

### Pool Questions Added (12 seed candidates)
All with properly structured `Options` arrays. Topics: scheduling, policy clarity, contract priorities, caseload, steward accessibility, grievance process, meeting times, union confidence, communication, respect, teamwork.

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

## 📋 CHANGE LOG — v4.24.2 Manual Community Poll Draw

**Date:** 2026-03-07
**Version:** v4.24.2
**Agent:** Claude (claude.ai)

### Feature: Steward Manual Community Poll Draw

**Problem:** The community poll slot only fills via a Monday time trigger (`autoSelectCommunityPoll`). That trigger has never been installed on the live script, so the community poll track has never worked. No steward had any way to release it.

**Solution:** New GAS wrapper + steward UI button.

#### Backend — `24_WeeklyQuestions.gs`
`wqManualDrawCommunityPoll()` — steward-only wrapper. Calls `WeeklyQuestions.selectRandomPoolQuestion()` directly, bypassing `autoSelectCommunityPoll`'s day-of-week guards. Logs to audit trail. Returns `{ success, message, questionId?, text? }`.

#### Frontend — `steward_view.html` — `_renderCreateStewardPoll()`
The static pool count info line replaced with a full **Community Poll Pool card**:
- Pool count display
- **"Draw Community Poll"** button — disabled (with tooltip) when pool is empty
- Inline status: shows drawn question text on success, error message on failure
- Button disabled during in-flight request (prevents double-draw)
- Empty pool guard fires client-side before the server call

#### Why not a toggle
A toggle implies an on/off persistent state. The draw is a one-time action per period — toggling it "on" would be meaningless since there's nothing to turn off. A button is the correct affordance.

#### Note
`setupCommunityPollTrigger()` should still be run eventually to automate future draws. The manual draw button is a first-release tool and a fallback for periods where the trigger misses.

---

### v4.20.18a — Minutes CTA bug fixes (2026-03-07)

#### `member_view.html` — `renderMinutesPage()`
Three bugs fixed:

| Bug | Root cause | Fix |
|-----|-----------|-----|
| CTA never showed | `!CONFIG.grievancesFolderId === false` — `!x === false` evaluates as `!!x` due to precedence; only true when folder IS set | Removed condition entirely; CTA always shown |
| CTA click crashed | Called `renderMyCasesPage()` which does not exist | Corrected to `renderMyCases()` |
| CTA disappeared after load | `showLoading(content)` clears `content.innerHTML`; CTA was appended to `content` before `showLoading` | CTA now appended to `container`; spinner lives in separate `content` div; success handler does `content.innerHTML = ''` safely |

#### `03_UIComponents.gs`
Menu item renamed: `Setup Dashboard Folder Structure` → `Setup / Repair Drive Folder Structure`
Rationale: re-running this on an existing deployment sets Minutes/ to view-only sharing.

#### `05_Integrations.gs` — `SETUP_DRIVE_FOLDERS()`
Alert now explicitly confirms Minutes/ folder sharing status.

---

## 📋 CHANGE LOG — v4.24.3 Auto-Draw Skip When Manually Released

**Date:** 2026-03-07
**Version:** v4.24.3
**Agent:** Claude (claude.ai)

### Feature: Monday Trigger Skips If Community Poll Already Active

**Problem:** `autoSelectCommunityPoll` (Monday time trigger) would overwrite a community poll that a steward had already manually released via the v4.24.2 draw button.

**Fix:** Added a pre-draw check in `autoSelectCommunityPoll` only (not in `selectRandomPoolQuestion` — manual steward draw via button can still overwrite intentionally).

**Logic:** Before calling `selectRandomPoolQuestion()`, the trigger now reads `_Weekly_Questions` and checks if any row has `SOURCE = 'community'`, `ACTIVE = TRUE`, and `WEEK_START = current period`. If found → log and return without drawing.

**Column indices used** (mirrors Q_COLS inside the IIFE — kept in sync as named constants):
- `COL_SOURCE = 3`, `COL_WEEK_START = 5`, `COL_ACTIVE = 6`

**⚠️ Schema dependency:** If `24_WeeklyQuestions.gs` Q_COLS indices ever change, `autoSelectCommunityPoll`'s named constants must be updated to match.


---

## 📋 CHANGE LOG — v4.24.4 Eliminate Duplicated Q_COLS Indices

**Date:** 2026-03-07
**Version:** v4.24.4
**Agent:** Claude (claude.ai)

### Fix: Single Source of Truth for Q_COLS

**Problem (v4.24.3):** `autoSelectCommunityPoll` used hardcoded `COL_SOURCE=3`, `COL_WEEK_START=5`, `COL_ACTIVE=6` constants copied from `Q_COLS` inside the IIFE. Schema drift risk — two places to update.

**Fix:** Exposed `Q_COLS` through the `WeeklyQuestions` public API return object. `autoSelectCommunityPoll` now reads `WeeklyQuestions.Q_COLS.SOURCE`, `.WEEK_START`, `.ACTIVE` directly. One definition, one place.

**Change in `24_WeeklyQuestions.gs`:**
- Public return object: added `Q_COLS: Q_COLS`
- `autoSelectCommunityPoll`: replaced `COL_*` locals with `var qc = WeeklyQuestions.Q_COLS`


---

## 📋 CHANGE LOG — v4.24.5 Remove Dead TYPE Column from Feedback

**Date:** 2026-03-07
**Version:** v4.24.5
**Agent:** Claude (claude.ai)

### Fix: Removed `Type` column from Feedback schema

**Problem:** `FEEDBACK_HEADER_MAP_` had a `TYPE` key (column "Type") that was never surfaced in the submit form. The backend default `feedbackData.type || 'Suggestion'` silently wrote `'Suggestion'` to every row because the frontend never sent `type`. `Category` already covers the same semantic ground.

**Files changed:**
- `src/01_Core.gs` — Removed `{ key: 'TYPE', header: 'Type' }` from `FEEDBACK_HEADER_MAP_`
- `src/21_WebDashDataService.gs` — Removed `type` slot from `submitFeedback` row array
- `src/10b_SurveyDocSheets.gs` — Removed `setColumnWidth(FEEDBACK_COLS.TYPE)` and Type dropdown `setDataValidation` block
- `src/steward_view.html` — Removed stale `CURRENT_USER.email` arg from `.dataGetMyFeedback()` call (email is resolved server-side; the arg was silently ignored)

**Also documented (no code change):** Stewards see the same two-tab Feedback page as members — no management panel to update Status/Resolution. Flagged for future enhancement.

---

## v4.24.6 — Survey Post-Review Fix Batch (2026-03-07)

Applied to remote base at v4.24.4. See commit log for full details.
12 issues fixed — see CHANGELOG entry for v4.24.6 in 01_Core.gs.

### Key Rules
- `buildSatisfactionColsShim_()` maps all SATISFACTION_COLS key names to dynamic positions
- `dataGetSatisfactionSummary(sessionToken)` is member+steward accessible (aggregate anon data)  
- `window._surveyDraft` replaces localStorage for survey progress in GAS iframe context
- 28 double-paren syntax errors in 21_WebDashDataService.gs all fixed
- `initSurveyEngine()` has NOT been run on live sheet — do not run until owner reviews Questions sheet


---

## 📋 CHANGE LOG — v4.24.8 Critical Auth Fix: Missing sessionToken Parameters + Bootstrap Recovery

**Date:** 2026-03-07
**Version:** v4.24.8
**Agent:** Claude Opus 4.6 (claude.ai)

### Reported Symptoms
1. **White screen** in normal browser (non-incognito)
2. **"Email me a sign-in link"** → "Something went wrong / Failed to send email"
3. **"Continue with Google"** → error about needing Google account (expected in incognito)

### Root Cause
The v4.23.1 auth sweep (session token migration) rewrote function bodies to call `_resolveCallerEmail(sessionToken)` or `_requireStewardAuth(sessionToken)` but **did not add `sessionToken` to 6 function parameter lists**. The client correctly passed `SESSION_TOKEN`, but the server functions declared `()` (no params), so `sessionToken` was `undefined` inside. This caused `ReferenceError` on every call for magic-link/session-token users (SSO users were unaffected because `_resolveCallerEmail` tries SSO first).

### Bugs Fixed

**Bug 1 — 6 wrapper functions missing `sessionToken` parameter** (21_WebDashDataService.gs):
- `dataGetBroadcastFilterOptions()` → `dataGetBroadcastFilterOptions(sessionToken)`
- `dataGetStewardDirectory()` → `dataGetStewardDirectory(sessionToken)`
- `dataGetEngagementStats()` → `dataGetEngagementStats(sessionToken)`
- `dataGetWorkloadSummaryStats()` → `dataGetWorkloadSummaryStats(sessionToken)`
- `dataGetWelcomeData()` → `dataGetWelcomeData(sessionToken)`
- `dataMarkWelcomeDismissed()` → `dataMarkWelcomeDismissed(sessionToken)`

**Bug 2 — `dataGetBatchData` client-server parameter mismatch** (index.html + 21_WebDashDataService.gs):
- Client called `.dataGetBatchData(CURRENT_USER.email, CURRENT_VIEW)` but server expected `(sessionToken)`
- Fixed client to call `.dataGetBatchData(SESSION_TOKEN)`
- Server derives email via `_resolveCallerEmail(sessionToken)` and role from Member Directory

**Bug 3 — `sendMagicLink` no top-level try/catch** (19_WebDashAuth.gs):
- Unhandled exceptions before the `MailApp.sendEmail()` try/catch caused `google.script.run` failure handler to fire with generic error
- Added outer try/catch with `Logger.log` for server-side diagnostics

**Bug 4 — Client bootstrap crash = white screen** (index.html):
- `DOMContentLoaded` handler had no error recovery — any JS exception during init left a blank page
- Added try/catch with fallback error UI including "Clear & Reload" button that clears stale localStorage tokens

**Bug 5 — localStorage redirect could throw** (index.html):
- `localStorage.getItem()` / `removeItem()` can throw in restricted contexts (Safari ITP, certain iframes)
- Wrapped session-restore block in try/catch

### Files Changed
| File | Change |
|------|--------|
| `src/21_WebDashDataService.gs` | Added `sessionToken` param to 6 global wrapper functions |
| `src/index.html` | Fixed `dataGetBatchData` call, added bootstrap error recovery, guarded localStorage |
| `src/19_WebDashAuth.gs` | Added outer try/catch to `sendMagicLink` |
| `dist/*` | Rebuilt from src via `node build.js` |

### Key Rules (reminder for future agents)
- ALL `google.script.run` wrapper functions that call `_resolveCallerEmail(sessionToken)` or `_requireStewardAuth(sessionToken)` MUST declare `sessionToken` as a parameter
- Client calls MUST pass `SESSION_TOKEN` (not email or view) as the first argument to any auth-wrapped server function
- The `dist/` directory is the deployment source — always rebuild after editing `src/`
- `dataGetBatchData` resolves both email AND role server-side — client only sends `SESSION_TOKEN`

### Deployment Required
After pushing, Wardis must:
1. `clasp push` from local machine (rootDir: `./dist`)
2. Update the active web app deployment to a new version in Apps Script editor


---

## 📋 CHANGE LOG — v4.24.9 White Screen Fix + MailApp Scope

**Date:** 2026-03-07
**Version:** v4.24.9
**Agent:** Claude Opus 4.6 (claude.ai)

### Root Cause — White Screen
**steward_view.html line 3359** had an unescaped apostrophe in a single-quoted JS string:
```javascript
'Create This Week's '   // ← apostrophe in Week's terminates the string
```
This is a **fatal syntax error** that kills the ENTIRE `<script>` block in steward_view.html, making `initStewardView()` undefined. When the main script calls it → `ReferenceError` → page stays blank.

**Fix:** Changed outer quotes to double quotes: `"Create This Week's "`

### Root Cause — Magic Link "Failed to send email"
**appsscript.json** had `gmail.send` scope (for `GmailApp`) but NOT `script.send_mail` (for `MailApp`). Auth module uses `MailApp.sendEmail()`, which requires the `script.send_mail` scope. These are different APIs with different OAuth scopes.

**Fix:** Added `https://www.googleapis.com/auth/script.send_mail` to oauthScopes in both `appsscript.json` and `dist/appsscript.json`.

### Root Cause — "Continue with Google" error in incognito
Expected behavior. In incognito mode, users aren't logged into Google, so SSO can't resolve their identity. The error message correctly directs them to use the email link option instead.

### Key Rules (for future agents)
- **ALWAYS syntax-check HTML view files after editing** using: `node -e "const vm=require('vm'); const fs=require('fs'); new vm.Script(fs.readFileSync('file.html','utf8').replace(/<script>/,'').replace(/<\/script>/,''))"`
- **Never use unescaped apostrophes in single-quoted JS strings** — use double quotes or `\\`
- **MailApp needs `script.send_mail`**, GmailApp needs `gmail.send` — they are DIFFERENT scopes
- A syntax error in ANY included HTML view file silently kills ALL functions in that file's `<script>` block

### Files Changed
| File | Change |
|------|--------|
| `src/steward_view.html` | Fixed unescaped apostrophe in `Week's` string |
| `appsscript.json` | Added `script.send_mail` OAuth scope |
| `dist/appsscript.json` | Added `script.send_mail` OAuth scope |
| `dist/*` | Rebuilt from src |

### Deployment Note
After `clasp push`, Google will prompt the script owner to **re-authorize** the script because a new OAuth scope was added. Accept the authorization prompt.


---

## 📋 CHANGE LOG — v4.25.0 Deploy Guards + 26 Client-Server Auth Mismatches Fixed

**Date:** 2026-03-07
**Version:** v4.25.0
**Agent:** Claude Opus 4.6 (claude.ai)

### Deploy Guard Tests (test/deploy-guards.test.js)
New automated test suite (64 tests) that catches every class of bug from the v4.24.8/9 outage:

| Guard | What It Catches | How |
|-------|----------------|-----|
| G1 | JS syntax errors in HTML view files | Parses every `<script>` block with `vm.Script` |
| G2 | Missing `sessionToken` parameter on server wrappers | Scans function bodies for `sessionToken` refs vs param list |
| G3 | OAuth scope gaps | Maps API calls (MailApp, GmailApp, etc.) to required scopes |
| G4 | Client-server parameter count mismatches | Compares `google.script.run` call arg counts to server signatures |
| G5 | Unescaped apostrophes in single-quoted JS strings | Regex scan for `'word's'` patterns |
| G6 | Stale dist/ (not rebuilt after src changes) | Byte-compares src/ vs dist/ |
| G7 | Syntax errors in .gs files | Parses all .gs with `vm.Script` |

**Run:** `npm run test:guards` or automatically via `npm run deploy` / `npm test`

### 26 Client-Server Parameter Mismatches Fixed

**QA Forum (26_QAForum.gs) — removed redundant `email`/`stewardEmail` params:**
The v4.23.1 auth sweep added `sessionToken` as first param but left the legacy `email` param in position 2. Client correctly stopped sending email, but server still expected it positionally → all QA args shifted by one → pagination, upvotes, moderation were broken.

Affected: `qaGetQuestions`, `qaGetQuestionDetail`, `qaSubmitQuestion`, `qaSubmitAnswer`, `qaUpvoteQuestion`, `qaModerateQuestion`, `qaModerateAnswer`, `qaGetFlaggedContent`, `qaResolveQuestion`

**Client calls missing SESSION_TOKEN (steward_view.html + member_view.html):**
- `dataStaffCompleteMemberTask(t.id)` → `(SESSION_TOKEN, t.id)`
- `dataGetCaseChecklist(caseId)` → `(SESSION_TOKEN, caseId)` (both views)
- `wqManualDrawCommunityPoll()` → `(SESSION_TOKEN)`
- `dataGetMeetingMinutes(100)` → `(SESSION_TOKEN, 100)` (both views)
- `dataGetMyFeedback()` → `(SESSION_TOKEN)`
- `dataGetUpcomingEvents(20)` → `(SESSION_TOKEN, 20)` (member_view)
- `DataCache.cachedCall` for events: `[10]` → `[SESSION_TOKEN, 10]`

**`dataAssignSteward` — missing member email (member_view.html):**
Client sent `(SESSION_TOKEN, stewardEmail)` but server expects `(sessionToken, memberEmail, stewardEmail)`. Added `CURRENT_USER.email` as second arg.

### Key Rules for Future Agents
- **Run `npm run test:guards` before ANY deployment** — it catches parameter mismatches, syntax errors, scope gaps, and stale dist
- **When adding `sessionToken` to a server function, update ALL client calls** — search all HTML view files
- **When removing a parameter from a server function, update ALL client calls** — positional args shift
- **Never deploy without rebuilding dist/** — `node build.js` then verify G6 passes

### Files Changed
| File | Change |
|------|--------|
| `test/deploy-guards.test.js` | NEW — 64 automated deploy guard tests |
| `src/26_QAForum.gs` | Removed redundant email/stewardEmail from 9 wrapper functions |
| `src/steward_view.html` | Added SESSION_TOKEN to 5 client calls |
| `src/member_view.html` | Added SESSION_TOKEN to 5 client calls + CURRENT_USER.email to dataAssignSteward |
| `package.json` | Added `test:guards` script, wired into deploy + lint-staged |
| `dist/*` | Rebuilt from src |


---

## 🏗️ FUTURE ARCHITECTURE — Extract JS from HTML View Files

**Status:** Planned, not scheduled
**Priority:** Medium — reduces risk class, improves DX
**Effort:** Major refactor

### Problem
`steward_view.html` (~3,900 lines) and `member_view.html` (~4,400 lines) each contain a single `<script>` block with thousands of lines of JavaScript. ESLint cannot lint JavaScript embedded in HTML files (the standard `eslint-plugin-html` fails here because GAS template tags like `<?!= include('styles') ?>` produce parse errors). This means ~8,000 lines of client-side JS are invisible to the linter — the exact code where the v4.24.9 fatal syntax error lived.

Current mitigation: `build.js` validates HTML `<script>` blocks with `vm.Script` at build time, and `test/deploy-guards.test.js` G1 does the same at test time. These catch syntax errors but not logic bugs, unused variables, or style issues that ESLint would flag.

### Proposed Fix
Extract the JavaScript from each HTML view file into separate `.gs` files (e.g., `steward_view_client.gs`, `member_view_client.gs`) that GAS can serve via `HtmlService.createHtmlOutputFromFile()` or inline via `<?!= include() ?>`. The HTML files would then contain only markup + a `<script src>` or `<?!= include('steward_view_client') ?>` tag.

**Benefits:**
- Full ESLint coverage on all client-side JS
- Easier code navigation (IDE support, go-to-definition)
- Smaller diffs on HTML-only or JS-only changes
- Could enable future bundling/minification

**Risks:**
- GAS `include()` returns raw HTML — extracted .gs files would need to be `.html` files containing only `<script>` blocks (current pattern already works this way for `auth_view.html` and `error_view.html`)
- Variable scoping: all view JS currently shares a single `<script>` scope — splitting requires careful attention to shared state (`AppState`, `CONFIG`, `CURRENT_USER`, etc.)
- Testing: deploy-guards G4 (param matching) would need updated file paths

**Recommendation:** Do this when a major feature requires significant edits to either view file, so the refactor cost is amortized against planned work.

---

### 2026-03-08 — POMS Smart Search Integration (v4.24.8)
- **Action**: Added POMS Reference as shared tab in web app (both steward + member roles)
- **New Files**:
  - `src/poms_reference.html` — 78 POMS sections, 17 flowcharts, keyword search, star ratings, bookmarks, notes (78KB, CSS-scoped under `.poms-root`)
  - `dist/poms_reference.html` — copy
- **Modified Files**:
  - `src/index.html` / `dist/index.html` — Added `{ id: 'poms', icon: '📘', label: 'POMS Ref.' }` to both steward and member tab arrays. Added `renderPOMSReference()` function (lazy-load pattern identical to org chart). Added `if (tabId === 'poms')` shared tab handler.
  - `src/22_WebDashApp.gs` / `dist/22_WebDashApp.gs` — Added `getPOMSReferenceHtml()` server function (returns poms_reference.html content)
  - `src/01_Core.gs` / `dist/01_Core.gs` — Version entry 4.24.8
- **Architecture**: Same lazy-load pattern as org chart — `google.script.run.getPOMSReferenceHtml()` fetches HTML, injected into container, `<script>` tags re-executed. Container set to `height:calc(100vh-56px)` for full-page POMS app experience.
- **CSS Scoping**: All POMS styles prefixed with `p-` or `poms-` to avoid SPA conflicts. No `:root` vars. No `body` selectors.
- **Security**: No DDS Script ID in poms_reference.html (verified grep)
- **Sync**: DDS Main → UT staging (commit 1945f90). DDS Script ID check clean.

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

---

## 2026-03-08 — Member List Clickability UX Fix

### Issue
Members in the steward's Members tab appeared non-interactive — no hover state, no cursor indicator, no visual hint they could be clicked to expand details.

### Root Cause
CSS for `.member-list-item` had no `cursor: pointer`, no hover/active states, and no expand indicator. The click handlers existed in JS but the visual affordance was absent.

### Fix
1. **CSS** (`styles.html`): Added `cursor: pointer`, `border-radius: 8px`, hover background, active scale transform, chevron rotation animation, and panel slide-in animation
2. **JS** (`steward_view.html`): Added `▶` chevron span to each member row. Click handler now toggles `.expanded` class for chevron rotation. Siblings get `.expanded` removed on new selection.

### Files Changed
- `src/styles.html` — `.member-list-item` hover/active states, `.member-list-chevron` rotation, `.member-detail-panel` animation
- `src/steward_view.html` — chevron element + expanded class toggle in click handler

### How Member Click Works
1. Steward opens Members tab → `renderStewardMembers()` calls `dataGetAllMembers(SESSION_TOKEN)`
2. Members render as `div.member-list-item` rows with avatar, name, location, badges, chevron
3. Click → `_toggleMemberDetail()` → either collapses (if same member) or expands detail panel
4. Detail panel shows: email, location, office days, dues status + action buttons (Full Profile, Log Contact, Send Notification, Case Folder)
5. "Full Profile" button calls `dataGetFullProfile(email)` for additional fields (job title, supervisor, joined date, member admin folder link)

---

## 2026-03-08 — POMS Missing from Mobile More Menus + Members Tab Error Feedback

### Issue 1: POMS Reference not visible in steward view (mobile)
POMS was in sidebar tabs (`_getSidebarTabs`) and routing (`_handleTabNav`) but was **missing from both the steward and member More menus** on mobile. Mobile users couldn't access POMS at all.

### Fix
Added `{ icon: '📘', label: 'POMS Ref.', ... }` to:
- `steward_view.html` → `renderStewardMore()` menu items (after Org Chart)
- `member_view.html` → `renderMemberMore()` menu items (after Org Chart)
Both route through `_handleTabNav(role, 'poms')` for consistent behavior.

### Issue 2: Members tab — silent failure on auth
`dataGetAllMembers(SESSION_TOKEN)` returns `[]` when `_requireStewardAuth` fails (expired/invalid session). The success handler treats this as "no members" with no indication of the real problem.

### Fix
Improved the empty-state message to suggest refreshing if no members are found unexpectedly.

### Files Changed
- `src/steward_view.html` — POMS in More menu + improved members empty-state message
- `src/member_view.html` — POMS in More menu

---

## 2026-03-08 — Member Click Debug Tracing

### Issue
Steward Members tab loads and displays members, but clicking a member row does nothing — no panel expansion, no error visible.

### Investigation
Code review shows click handler (`addEventListener`), `_toggleMemberDetail`, and `_showMemberDetail` are all wired correctly. Panel creation and `insertBefore` DOM insertion logic is sound. CSS variables (`--raised`, `--accent`, etc.) are defined by ThemeEngine. No `overflow: hidden` or `pointer-events: none` blocking the content area.

### Diagnosis Approach
Added try-catch with visible error display to three functions:
1. Click handler in `renderFilteredMembers` — catches and shows red error below clicked row
2. `_toggleMemberDetail` — catches and logs to console
3. `_showMemberDetail` — catches and shows red error panel below clicked row

If clicks do nothing AND no red error appears → the click event isn't firing (CSS overlay / z-index issue)
If red error appears → the error message will reveal the root cause
If panel appears but is invisible → theme/CSS variable issue

### Files Changed
- `src/steward_view.html` — try-catch wrappers with visible error feedback

---

## 2026-03-08 — Error Visibility + System Diagnostic

### Issue
Multiple tabs (Timeline, potentially others) show "Something went wrong. Please try again." with no indication of what the actual error is. The default `serverCall()` failure handler logs to console but replaces spinners with a generic message, making debugging impossible without F12.

### Root Cause (likely)
Server-side functions are throwing exceptions — most likely because required hidden sheets don't exist in the user's spreadsheet. Many features (Timeline, Q&A Forum, Feedback, Tasks, etc.) depend on auto-created hidden sheets. If auto-creation fails or sheets were never initialized, the feature breaks.

### Fixes
1. **Improved serverCall() failure handler** (`index.html`): Now shows the actual error message in the UI with a ⚠️ icon and Reload button, instead of the opaque "Something went wrong."
2. **System Diagnostic function** (`28_FailsafeService.gs`): New `fsDiagnostic(sessionToken)` function checks all 25+ required/optional sheets, auth status, and returns a health report.
3. **Diagnostic UI** (`steward_view.html`): New "System Diagnostic" card on the Failsafe page with "Run Diagnostic" button. Shows ✅/❌ for every sheet, highlights missing ones.

### How to Use
Steward sidebar → Data Failsafe → Run Diagnostic → see which sheets are missing → missing sheets likely explain which tabs are broken.

### Files Changed
- `src/index.html` — serverCall() failure handler now shows actual error
- `src/28_FailsafeService.gs` — new fsDiagnostic() server function
- `src/steward_view.html` — diagnostic card in Failsafe page

---

## 2026-03-08 — Auto-Init Missing Sheets on SPA Load

### Design
Per user feedback: "why not just have createDashboard auto-create all the sheets?" — CREATE_DASHBOARD already calls all init functions, but it requires Sheets UI (menu bar + confirmation dialog). It can't run from the SPA. If CREATE_DASHBOARD was run on an older code version, newer sheets are missing.

### Solution: Two-layer approach
1. **Auto-init on first load** (`21_WebDashDataService.gs`): `dataGetBatchData()` now calls `_ensureAllSheetsInternal()` on the first SPA load per version. Uses Script Property key `sheetsInitialized_{VERSION}` — re-runs automatically after each code update. No auth requirement (runs server-side). Non-destructive.
2. **Manual button** (`28_FailsafeService.gs` + `steward_view.html`): "Initialize Missing Sheets" button on Failsafe page calls `fsEnsureAllSheets()` with detailed success/failure feedback per sheet group.

### What gets auto-created (17 groups)
Hidden calculation sheets, Contact Log, Steward Tasks, QA Forum + Answers, Timeline Events, Failsafe Config, Weekly Questions + Responses + Pool, Portal sheets, Workload Tracker (4 sheets), Resources, Resource Config, Survey Questions, Member Satisfaction, Feedback, Case Checklist, Notifications, Audit Log.

### Key design decisions
- **Version-keyed flag**: `sheetsInitialized_{VERSION}` ensures re-run after code updates but not on every page load.
- **No auth gate on internal function**: `_ensureAllSheetsInternal()` runs for any user (member or steward) because sheet creation is a system concern, not a role concern.
- **try/catch per init**: Each sheet group is independent — one failure doesn't block others.
- **typeof guards**: Every init call checks `typeof` first — if a module isn't loaded, it silently skips.

### Files Changed
- `src/21_WebDashDataService.gs` — auto-init in dataGetBatchData() + _ensureAllSheetsInternal()
- `src/28_FailsafeService.gs` — fsEnsureAllSheets() with detailed tracking
- `src/steward_view.html` — "Initialize Missing Sheets" button on Failsafe page

---

## 2026-03-08 — Code Review Fixes (latency, debug cleanup, DRY)

### Issue 1: Auto-init on critical load path
**Problem**: `_ensureAllSheetsInternal()` ran inside `dataGetBatchData()` — the initial data fetch. Added latency on first load per version.
**Fix**: Moved to separate `dataEnsureSheetsIfNeeded()` endpoint. Client fires it as fire-and-forget from `DOMContentLoaded` AFTER `initApp()` dispatches. Never blocks initial render.

### Issue 2: Debug scaffolding in production
**Problem**: Commit 0be801c added try-catch wrappers with visible red DOM error boxes in member click handlers — diagnostic code not appropriate for release.
**Fix**: Removed red error DOM insertion from click handler, `_toggleMemberDetail`, and `_showMemberDetail`. Click handler retains a minimal try-catch with console.error only. The other two functions are unwrapped entirely since their callers already catch.

### Issue 3: DRY violation
**Problem**: `fsEnsureAllSheets()` (28_FailsafeService.gs) and `_ensureAllSheetsInternal()` (21_WebDashDataService.gs) contained nearly identical 17-item init lists.
**Fix**: `_ensureAllSheetsInternal()` is now the single source of truth — returns `{ created[], failed[] }`. `fsEnsureAllSheets()` delegates to it with a steward auth wrapper. Zero duplication.

### Architecture after fixes
```
Client (index.html DOMContentLoaded)
  ├── initApp() → dataGetBatchData() → fast, no init work
  └── fire-and-forget → dataEnsureSheetsIfNeeded()
                            └── checks version flag
                            └── calls _ensureAllSheetsInternal() if needed

Failsafe page button
  └── fsEnsureAllSheets(token)
        └── _requireStewardAuth()
        └── _ensureAllSheetsInternal()   ← same function
```

### Files Changed
- `src/21_WebDashDataService.gs` — dataGetBatchData cleaned, dataEnsureSheetsIfNeeded added, _ensureAllSheetsInternal returns results
- `src/28_FailsafeService.gs` — fsEnsureAllSheets delegates to _ensureAllSheetsInternal
- `src/index.html` — fire-and-forget call after initApp
- `src/steward_view.html` — debug scaffolding removed from member click handlers

## Tab Reorder & Group Headers (2026-03-09)

### Changes Made
- **Sidebar group headers**: Added `{ group: 'Label' }` support to `_getSidebarTabs()` in `src/index.html`. Renderer at ~line 459 now handles `tab.group` alongside `tab.divider`.
- **CSS**: Added `.sidebar-group-label` style in `src/styles.html` (10px uppercase, letter-spacing 1.2px, `--text-secondary` color).
- **Three named groups** introduced across both roles:
  - **DDS**: POMS Ref., Workload Tracker
  - **Resources**: Insights (steward) / Union Stats (member), Resources, Org. Chart, Q&A Forum
  - **Activity**: Tasks / My Tasks, Events, Timeline
- **Steward Workload Tracker**: Added `workload` tab to steward sidebar and steward switch case in `_handleTabNav()`. `renderWorkloadTracker()` in `src/member_view.html` made role-aware via `AppState.activeRole`, using `_stewardHeader` or `_memberHeader` dynamically.
- **More menus**: Both `src/steward_view.html` and `src/member_view.html` `menuItems[]` reordered to match group structure. Group/divider rendering added to both `forEach` loops.

### Files Changed
1. `src/styles.html` — Added `.sidebar-group-label` CSS class
2. `src/index.html` — `_getSidebarTabs()` reordered with group headers; sidebar renderer updated; steward switch added `workload` case
3. `src/steward_view.html` — `menuItems[]` reordered with `{ group }` and `{ divider }` entries; renderer updated
4. `src/member_view.html` — `menuItems[]` reordered with `{ group }` and `{ divider }` entries; renderer updated; `renderWorkloadTracker()` made role-aware

### ⚠️ Union-Tracker Note
- WorkloadTracker typeof guards still needed in UT for 3 files (pre-existing issue, not introduced here).
- These changes must be synced DDS Main → UT staging per standard flow.

## Tab Regroup v2 — Intent-Based Groups (2026-03-09)

### Supersedes
Previous DDS/Resources/Activity grouping replaced same day with intent-based groups.

### New Group Structure (both roles)
- **Ungrouped top**: Core daily work (Cases, Members/Directory, Contact Log)
- **Manage**: Tasks, Workload Tracker, Insights/Union Stats — tracking & analysis
- **Reference**: POMS Ref., Resources, Org. Chart — read-only lookup tools
- **Community**: Q&A Forum, Polls, Feedback, Minutes — two-way engagement
- **Comms**: Broadcast (steward only), Notifications, Events, Timeline — scheduling & messaging
- **Ungrouped bottom**: Admin/settings (Survey Tracking, Failsafe for steward; Steward Dir, Profile, Survey Results for member)

### Reasoning
Groups map to workflow modes (track work → look up info → engage community → communicate → admin). Previous "DDS" label was confusing (whole app is DDS). Previous "Activity" group separated Tasks from Cases which breaks workflow proximity.

### Files Changed (same 3 as v1)
1. `src/index.html` — `_getSidebarTabs()` reordered
2. `src/steward_view.html` — `menuItems[]` reordered
3. `src/member_view.html` — `menuItems[]` reordered

## Sheet Tab Order & Tab Colors Update (2026-03-09)

### Changes
- Added `SHEETS.MEETING_CHECKIN_LOG` to `reorderSheetsToStandard()` in `src/08a_SheetSetup.gs` (position: after Meeting Attendance, before Workload Reporting)
- Added `SHEETS.MEETING_CHECKIN_LOG` to yellow (Activity tracking) color group in `applyTabColors_()` in `src/11_CommandHub.gs`

### Final Sheet Tab Order (16 tabs)
1-3: Member Directory, Grievance Log, Case Checklist (🟣 Purple — core data)
4-5: Resources, Resource Config (🟢 Green — reference)
6-8: Survey Questions, Member Satisfaction, Feedback (🔵 Blue — community)
9-11: Meeting Attendance, Meeting Check-In Log, Workload Reporting (🟡 Yellow — activity)
12-13: Getting Started, FAQ (🔴 Red — onboarding)
14-16: Function Checklist, Config Guide, Config (🟠 Orange — admin)

### Menu Structure Confirmed (already restructured in prior session)
- **Union Hub**: Dashboards, Search, Members, Cases, Analytics, Quick Actions, Help
- **Tools**: Outreach, Surveys & Polls, Calendar & Meetings, Drive, Workload, View & Display, Multi-Select, OCR, Web App
- **Admin**: Diagnostics, Data Sync, Validation, Automation, Cache, Security, Maintenance, Styling, Setup, Demo Data

## Sheet Tabs & Menus Restructure v3 (2026-03-09)

### Supersedes
All prior tab order/color/menu sections. This is the canonical layout.

### Sheet Tab Order (4 color groups, 17 tabs)
🟣 Purple — **Core Data** (daily): Member Directory, Grievance Log, Case Checklist, Workload Reporting
🟢 Green — **Reference** (look-up): Resources, Getting Started, FAQ
🔵 Blue — **Engagement** (community): Meeting Attendance, Meeting Check-In Log, Member Satisfaction, Feedback & Dev, Notifications
🟠 Orange — **Config & Admin** (owner-editable): Survey Questions, Resource Config, Function Checklist, Config Guide, Config

Reasoning: 4 groups simpler than 6. Workload Reporting promoted to Core (stewards check daily). Getting Started/FAQ moved to Reference (look-up, not prime position). Survey Questions/Resource Config moved to Admin (config sheets, not data sheets). Notifications added (Blue — engagement data).

### Menu Structure
**📊 Union Hub** (unchanged): Dashboards, Search, Members, Cases, Analytics, Quick Actions, Help
- Removed: Workload duplicate, Calendar/Drive shortcuts (now in Tools)

**🔧 Tools** (restructured):
- Email & Notifications (merged Communication + Notifications)
- Calendar & Meetings (removed setup items → Admin)
- Google Drive (removed setup items → Admin)
- Workload Tracker (single location, no duplicate)
- Surveys & Polls (merged, removed triggers → Admin)
- ── separator ──
- Sheet Tools (merged View & Display + Multi-Select, OCR removed)
- Themes (extracted from Admin for steward access)
- ── separator ──
- Get Mobile App URL (shortcut from Web App & Portal)

**🛠️ Admin** (restructured — broke up 27-item Setup):
- Diagnostics, Repair, Settings (top-level)
- Data Sync, Validation, Automation, Cache, Security (unchanged)
- ── separator ──
- Initial Setup (one-time items: dashboard, calendar, drive, survey, polls, workload, hidden sheets)
- Triggers (all trigger installations in one place)
- Maintenance (refresh, backfill, restore, create sheets, cleanup)
- Styling (config styling, tab colors, theme columns)
- Web App & Portal (moved from Tools — admin work + OCR buried here)

### Files Changed
1. `src/08a_SheetSetup.gs` — `reorderSheetsToStandard()` — new 4-group order with Notifications
2. `src/11_CommandHub.gs` — `applyTabColors_()` — 4 colors (dropped Yellow, Red)
3. `src/03_UIComponents.gs` — `createDashboardMenu()` — full menu restructure

---

## 🔧 FIX LOG — 2026-03-09: Org KPI Cards Not Populating (Steward Cases Tab)

### Problem
When a steward has 0 assigned cases, the Cases tab falls back to org-wide KPIs. The "Org Overdue" and "Org Due <7d" cards stayed as "..." (placeholder) and never populated.

### Root Causes Found
1. **Missing failure handler** (`steward_view.html:150`): The `dataGetGrievanceStats` call had no `.withFailureHandler` for the KPI cards. Server errors caused the generic handler to target `.loading-spinner` elements only — KPI `.kpi-value` elements were untouched.
2. **Silent early return** (`steward_view.html:151`): When `stats.available === false`, the success handler returned without updating cards. Cards stayed "...".
3. **Cases due today uncounted** (`21_WebDashDataService.gs`): `deadlineDays === 0` fell through both overdue (`< 0`) and dueSoon (`> 0`) checks. Three locations affected: `getGrievanceStats`, `getStewardKPIs`, `_computeKPIsFromCases`.
4. **No try/catch in `getGrievanceStats` loop** (`21_WebDashDataService.gs:1890`): One malformed row could crash the entire function.

### Files Changed
- `src/steward_view.html` — Added `.withFailureHandler` to org KPI `dataGetGrievanceStats` call; handle `available:false` by showing "0" instead of "..."
- `src/21_WebDashDataService.gs` — Changed `deadlineDays > 0` to `deadlineDays >= 0` in 3 locations; wrapped `getGrievanceStats` loop in try/catch

### How Org KPI Fallback Works
- `renderStewardDashboard()` checks `(kpis.totalCases || 0) === 0`
- If true → renders "Org Cases / Org Overdue / Org Due <7d / Org Resolved" with "..." placeholder
- Async call to `dataGetGrievanceStats(SESSION_TOKEN)` fills in real values
- `getGrievanceStats()` iterates ALL grievance rows (not steward-filtered) and counts overdue/dueSoon
- Overdue: `_buildGrievanceRecord` auto-derives from `deadlineDays < 0` on non-terminal cases
- DueSoon: `deadlineDays >= 0 && deadlineDays <= 7` on non-terminal cases
- Deadline column matched by aliases: `['deadline', 'next deadline', 'due date']`

### ⚠️ IMPORTANT: Deadline Column Must Match
If the Grievance Log's deadline column header doesn't match one of `['deadline', 'next deadline', 'due date']`, all deadline-based counts will be 0 (not an error, just 0). The column name is case-insensitive.

### Testing Added (2026-03-09)
**24 new unit tests** in `test/21_WebDashDataService.test.js`:
- **Deadline boundary conditions** (12 tests): Covers `deadlineDays` values of -30, -1, 0, 1, 7, 8, null. Tests overdue auto-detection, terminal status protection, and mixed scenarios.
- **KPI computation parity** (4 tests): Verifies `getStewardKPIs` and `getBatchData.kpis` (via `_computeKPIsFromCases`) produce identical counts for the same data.
- **Error resilience** (5 tests): Malformed rows (short rows, null dates, invalid date strings, empty sheet, missing sheet) — none should crash `getGrievanceStats`.
- **Return shape contracts** (3 tests): Ensures `getGrievanceStats` and `getStewardKPIs` always return the keys the frontend depends on.

**2 new deploy guards** in `test/deploy-guards.test.js`:
- **G8 — Failure handler ratchet**: Counts `serverCall()` chains with `.withSuccessHandler` but no `.withFailureHandler`. Current baseline: steward_view=30, member_view=38, index=0. Count must never INCREASE; decrease as debt is addressed.
- **G9 — KPI data contract**: Statically scans `getGrievanceStats`, `getStewardKPIs`, and `_computeKPIsFromCases` return statements for required keys (overdueCount, dueSoonCount, total, etc.). Catches silent key renames/deletions.
---

## 2026-03-09 — SPA Integrity Test Suite (G8–G14)

### Why
Every bug from the 2026-03-08/09 session would have been caught by automated tests:
- POMS missing from mobile → G9 (More menu parity)
- Stewards can't switch to member view → G10 (isDualRole gate)
- Tabs showing "Something went wrong" → G11 (sheet init completeness) + G13 (error visibility)
- Members not clickable → G14 (CSS affordance)
- Dead tabs without routing → G8 (route handler completeness)

### Tests Added (test/spa-integrity.test.js)

| Guard | What it catches | How |
|-------|----------------|-----|
| G8 | Sidebar tab with no route handler | Extracts all tab IDs from `_getSidebarTabs`, checks each has a `case` or shared `if` in `_handleTabNav` |
| G9 | Feature only on desktop, missing on mobile | Extracts sidebar tab IDs, subtracts bottom-nav tabs, checks remainder appears in More menu |
| G10 | Steward toggle gate regression | Asserts `isDualRole` condition includes `role === 'steward'`, not just `role === 'both'` |
| G11 | New sheet added to SHEETS constant but not to auto-init | Maps 17 feature sheet constants to init functions, checks each is covered in `_ensureAllSheetsInternal` |
| G12 | Client calls non-existent server function | Scans HTML for `.funcName()` on serverCall chains, checks each name exists as a `function` in .gs files |
| G13 | serverCall hides actual errors | Asserts failure handler references `err.message` or `err`, not just a generic string |
| G14 | Clickable elements without visual affordance | Checks `.member-list-item` has `cursor: pointer`, `:hover` state, and `addEventListener('click')` |

### Integration
- `npm run test:guards` now runs both `deploy-guards.test.js` and `spa-integrity.test.js`
- Pre-commit hook (`.husky/pre-commit`) runs guards as step 3 — blocks commits that break routing, auth, or sheet init
- `npm run deploy` pipeline: lint → guards → unit tests → build → clasp push

### Files Changed
- `test/spa-integrity.test.js` — new file, 12 tests
- `package.json` — test:guards script updated
- `.husky/pre-commit` — step 3 added for guards

---

## v4.25.0 — GAS-Native Test Runner (2026-03-09)

### What Changed
- **New file**: `src/30_TestRunner.gs` — GAS-native integration test framework
- **Modified**: `build.js` — added `30_TestRunner.gs` to BUILD_ORDER
- **Modified**: `src/index.html` — added `testrunner` tab to steward sidebar + routing
- **Modified**: `src/steward_view.html` — added `renderTestRunnerPage()` function + More menu entry
- **Modified**: `src/03_UIComponents.gs` — added 🧪 Test Runner submenu under Admin
- **Modified**: `src/01_Core.gs` — version bump to 4.25.0 + VERSION_HISTORY entry
- **Modified**: `package.json` — version bump to 4.25.0
- **Modified**: `CHANGELOG.md` — v4.25.0 entry
- **Modified**: `test/architecture.test.js` — added `30_TestRunner.gs` to BUILD_ORDER
- **Modified**: `test/spa-integrity.test.js` — added `renderTestRunnerPage` to renderMap

### How It Works
- **TestRunner IIFE module** in `30_TestRunner.gs` with assert helpers (assertEquals, assertTrue, assertNotNull, assertType, assertContains, assertThrows, assertGreaterThan, assertHasKey, assertDeepEquals, assertFalsy)
- **Test discovery** via `_getTestRegistry()` — explicit registry pattern (GAS doesn't support `Object.keys(this)`)
- **Naming convention**: `test_SUITE_description` — suite extracted from second segment
- **6 suites / 48 tests**:
  - `config` (10): SHEETS constant, sheet existence, ConfigReader shape, CONFIG_COLS, DEFAULT_CONFIG, DEADLINE_DEFAULTS
  - `colmap` (9): GRIEVANCE_COLS/MEMBER_COLS/CONFIG_COLS defined, required keys present, all positive, no duplicates
  - `auth` (8): Auth module, resolveUser, DataService, findUserByEmail, _resolveCallerEmail, _requireStewardAuth, checkWebAppAuthorization, session token functions
  - `grievance` (10): GRIEVANCE_STATUS, all statuses, closed statuses subset, priority coverage, deadline rules, step list, isValidGrievanceId, outcomes
  - `security` (6): escapeHtml, escapeForFormula, XSS blocking, null/number handling
  - `system` (5): VERSION_INFO, semver format, spreadsheet binding, EventBus, HIDDEN_SHEETS
- **All tests are READ-ONLY** — never write to sheets
- **Results stored** in ScriptProperties as JSON (key: `TEST_RUNNER_RESULTS`)
- **SPA panel**: steward-only tab with summary cards, per-suite collapsibles, auto-expand failures, monospace error display
- **Triggers**: manual menu item (`runTestsFromMenu`), daily 6AM (`runScheduledTests`)
- **Server endpoints**: `dataRunTests(sessionToken, filterSuite)`, `dataGetTestResults(sessionToken)`, `dataManageTestTrigger(sessionToken, action)` — all require `_requireStewardAuth`

### Key Design Decisions
1. **GAS-native over Jest mocks**: Tests run inside real GAS runtime hitting real Sheets/Config — catches integration bugs that mocked tests miss (column drift, Config tab misreads, permission issues)
2. **Registry pattern over discovery**: GAS V8 can't enumerate global functions reliably — explicit `_getTestRegistry()` is safer
3. **Read-only tests**: Zero risk of corrupting production data
4. **Steward-only access**: Test results could reveal system internals — gated behind `_requireStewardAuth`
5. **ScriptProperties storage**: Lightweight, persists across sessions, no additional sheet tab needed
6. **Auth on all endpoints**: Follows CLAUDE.md rule — every `data*` function begins with auth check

---

## v4.25.1 — Test Runner Expansion (2026-03-09)

### What Changed
- **Modified**: `src/30_TestRunner.gs` — added 4 new test suites (34 tests → 82 total)
- **Modified**: `src/steward_view.html` — added 4 new suite filter options to SPA dropdown
- **Modified**: `src/01_Core.gs` — version bump to 4.25.1 + VERSION_HISTORY entry
- **Modified**: `package.json`, `CHANGELOG.md` — version bump

### New Test Suites

**dataservice (10 tests):**
- DataService module exists as IIFE with expected public API
- `findUserByEmail` — callable, returns null for nonexistent/empty email (not throws)
- `getUserRole`, `getStewardCases`, `getAllMembers`, `getUnits`, `getBatchData` — all callable
- Public API completeness: 33 expected methods verified

**authsweep (6 tests):**
- 8 steward endpoints reject null token (return safe-empty `[]`, `{}`, etc.)
- 8 member endpoints reject null token (return `[]` or `null`)
- `dataGetBatchData(null)` does not leak member data
- Legacy poll stubs (`dataGetActivePolls`, `dataSubmitPollVote`) return safe values
- Test runner's own 3 endpoints (`dataRunTests`, `dataGetTestResults`, `dataManageTestTrigger`) reject null

**configlive (8 tests):**
- Config, Member Dir, Grievance Log sheets exist and have headers
- CONFIG_COLS, MEMBER_COLS, GRIEVANCE_COLS — no position exceeds actual sheet width (catches col drift)
- `syncColumnMaps` function exists
- Config row 3 ORG_NAME cell is not empty

**survey (10 tests):**
- `HIDDEN_SHEETS.SURVEY_TRACKING/VAULT/PERIODS` defined with `_` prefix
- `SURVEY_PERIODS_COLS` has PERIOD_ID and STATUS
- `SURVEY_QUESTIONS_COLS` has QUESTION_ID, QUESTION_TEXT, TYPE, ACTIVE
- `getSurveyQuestions()` returns array with valid question shape (`id`, `text`, `type` in known set)
- `getSurveyPeriod()` callable, returns null or object
- `submitSurveyResponse` and `SATISFACTION_COLS` exist (backward compat)
- Survey tracking sheet has header if it exists

### Design Rationale
- **Auth sweep tests real endpoints with null tokens** — catches any function that forgot `_resolveCallerEmail`/`_requireStewardAuth` at the top. Accepts either safe-empty return OR throw (both are valid rejection).
- **Config completeness checks sheet width vs constant values** — catches the common bug where a column is added to the header map but not to the actual sheet (or vice versa).
- **Survey tests read-only** — calls `getSurveyQuestions()` and `getSurveyPeriod()` but never `submitSurveyResponse()` or `openNewSurveyPeriod()`.

---

## v4.25.2 — Test Failure Notifications (2026-03-09)

### What Changed
- **Modified**: `src/01_Core.gs` — added `TEST_NOTIFY_EMAIL` to `CONFIG_HEADER_MAP_`, version bump
- **Modified**: `src/30_TestRunner.gs` — `runScheduledTests()` now emails on failure, added `_getTestNotifyEmail()` and `_sendTestFailureEmail(results)`
- **Modified**: `package.json`, `CHANGELOG.md` — version bump

### How It Works
1. **Config tab setup**: Add a header `Test Runner Notify Email` in the Config tab (auto-discovered by `syncColumnMaps`). Put an email address in row 3.
2. **Daily trigger fires** → `runScheduledTests()` → `TestRunner.runAll()` → if failures > 0 → `_sendTestFailureEmail(results)`
3. **Email content**: Subject line shows org name + failure count. Body lists each failed test grouped by suite with error messages. Both plain-text and styled HTML (dark theme).
4. **Quota guard**: Skips email if `MailApp.getRemainingDailyQuota() < 5` to avoid burning quota.
5. **No email on success**: Failure-only. Manual runs use toast only — no email.

### Config Column
- Key: `CONFIG_COLS.TEST_NOTIFY_EMAIL`
- Header: `Test Runner Notify Email`
- Position: After `Broadcast: Allow All Members Scope` (last column in CONFIG_HEADER_MAP_)
- Value: Any valid email address. Leave blank to disable notifications.

### Design Rationale
- **Failure-only**: Daily success emails are noise. If you want proof tests ran, check the SPA panel.
- **Config tab over hardcoded**: Follows "everything dynamic" rule. Admins can change the email without code changes.
- **Quota guard**: MassAbility DDS uses MailApp for other notifications too — don't compete for the 100/day limit.
- **HTML + plain-text**: HTML renders in Gmail/Outlook; plain-text is the fallback for text-only clients.

## v4.25.3 — Deadline Config Completeness (2026-03-09)

### Changes
- **3 new Config columns added** to CONFIG_HEADER_MAP_ (01_Core.gs):
  - `STEP3_APPEAL_DAYS` → header: "Step III Appeal Days" (default: 10)
  - `STEP3_RESPONSE_DAYS` → header: "Step III Response Days" (default: 21)
  - `ARBITRATION_DEMAND_DAYS` → header: "Arbitration Demand Days" (default: 30)
- **getDeadlineRules()** (01_Core.gs): Now reads Step III and Arbitration values from Config instead of hardcoded DEADLINE_DEFAULTS. Falls back to defaults if empty/NaN.
- **createConfigSheet** (10a_SheetCreation.gs): DEADLINES section header expanded from 4→7 cols. 3 new seedConfigDefault_ calls added.
- **COMMAND_CONFIG.VERSION** fixed from stale "4.24.4" to match actual version "4.25.3".

### Bug found
- COMMAND_CONFIG.VERSION was stuck at "4.24.4" while VERSION_INFO showed 4.25.2. Fixed to 4.25.3.

### Test updates
- modules.test.js: CHIEF_STEWARD_EMAIL position 42→45, ESCALATION_STATUSES 45→48, ESCALATION_STEPS 46→49 (shifted +3).
- New test: "has deadline config columns including Step III and Arbitration" (7 assertions).
- Total: 2405 tests pass.

### Files modified
1. `src/01_Core.gs` — CONFIG_HEADER_MAP_, VERSION_INFO, VERSION_HISTORY, COMMAND_CONFIG.VERSION, getDeadlineRules()
2. `src/10a_SheetCreation.gs` — sectionHeaders, seedConfigDefault_ calls
3. `test/modules.test.js` — position updates, new test
4. `AI_REFERENCE.md` — this entry

## v4.25.4 — Stability: TestRunner Timeout Guard + Trigger Audit (2026-03-09)

### Problem diagnosed
- **Web app unstable**: Slow load → test timeouts → HTTP errors → complete failure to load.
- **Root cause**: 82 GAS-native tests run without any timeout guard. Combined with 3.15 MB of .gs code (71,414 lines) that GAS must parse on every execution, plus 558 KB HTML payload evaluated server-side via `<?!= include() ?>`, the 6-minute GAS execution limit was exceeded.
- **Cascading failure**: Timed-out test runs + 20+ installed triggers (daily security digest, hourly enforceHiddenSheets, daily dashboard refresh, backup triggers, etc.) exhausted the daily execution quota. Once quota is depleted, `doGet()` itself fails → app won't load at all until quota resets (24 hours).

### Changes made

#### 1. TestRunner timeout guard (src/30_TestRunner.gs)
- **Global timeout**: `MAX_RUNTIME_MS = 300000` (5 min) — bails before hitting GAS 6-min limit.
- **Per-test slow detection**: `PER_TEST_MAX_MS = 30000` (30s) — flags slow tests in results.
- **Graceful skip**: When timeout is hit, remaining tests are counted as `skipped` (not failed). Status set to `'timeout'` instead of `'complete'`.
- **New fields in results**: `results.timedOut` (boolean), `suiteResult.skipped` (count), `testResult.slow` (boolean flag on slow tests).

#### 2. Trigger audit utility (src/06_Maintenance.gs)
- **`auditAllTriggers()`**: READ-ONLY. Logs all installed triggers with handler names, event types, sources. Detects duplicates (same handler installed multiple times). Added to 🛡️ Data Integrity menu.
- **`cleanupDuplicateTriggers(dryRun)`**: Removes duplicate triggers (keeps one per handler). Safe default: `dryRun=true`.
- **`dataAuditTriggers(sessionToken)`**: SPA endpoint for steward-only access.

### Key metrics (for future reference)
| Metric | Value |
|--------|-------|
| .gs file count | 42 |
| .gs total size | 3,153,281 bytes (3.15 MB) |
| .gs total lines | 71,414 |
| HTML payload (all views inlined) | 558 KB |
| Test count | 82 |
| Installable trigger sources | 20+ locations across codebase |

### Known issue: HTML payload bloat
`index.html` inlines ALL views via `<?!= include() ?>` regardless of user role:
- `steward_view.html` = 225 KB (sent to members too)
- `member_view.html` = 241 KB (sent to stewards too)
- `styles.html` = 38 KB
This 558 KB total is near HtmlService limits. **P1 fix needed**: lazy-load only the relevant view.

### Files modified
1. `src/30_TestRunner.gs` — timeout guard in `runAll()`, new constants `MAX_RUNTIME_MS`, `PER_TEST_MAX_MS`
2. `src/06_Maintenance.gs` — `auditAllTriggers()`, `cleanupDuplicateTriggers()`, `dataAuditTriggers()`, menu item
3. `AI_REFERENCE.md` — this entry

---

## Nav Theme System — Phase 1 Implementation (March 2026)

### What was added
5 switchable navigation visual themes, stored per-user in localStorage.

### Theme definitions
| Theme ID | Label | Type | Visual |
|----------|-------|------|--------|
| `default` | Default | default | Standard nav — follows dark/light mode |
| `blobLava` | Blob Lava | default | Morphing lava-lamp blob tracks active tab |
| `cyberpunk` | Cyberpunk | extra | Grid lines, scanlines, hard neon glow per tab |
| `shatter` | Shatter | extra | Colored shard blocks per tab, flash-on-select |
| `liquidPour` | Liquid Pour | extra | Floating bubble-style nav items with per-tab color |

### Architecture
1. **NavThemes engine** (`index.html`): Singleton object providing theme registry, per-tab color mapping, localStorage persistence (`dds_navTheme` key), body class management (`nav-theme-{id}`).
2. **Per-tab accent colors** (`NavThemes.TAB_COLORS`): Every sidebar/bottom tab ID mapped to a distinct color — used by blob indicator, shatter blocks, and liquid bubbles.
3. **CSS themes** (`styles.html`): 5 theme rulesets scoped via `.nav-theme-{id}` body class. Includes CSS animations: `blobMorph`, `blobPulse`, `cpGridMove`, `cpScanline`, `cpFlicker`, `shatterGlow`, `bubbleFloat1/2`, `bubblePop`.
4. **Blob indicator** (`steward_view.html` `renderBottomNav`): For blobLava and cyberpunk themes, a `.nav-blob-indicator` div is added to `.bottom-nav` and positioned via `requestAnimationFrame` to sit behind the active tab. CSS vars `--blob-color`, `--blob-glow` drive coloring.
5. **Theme picker** (`index.html` `renderSidebarItems`): Expandable panel in sidebar under "Nav Style" item. Shows all 5 themes with emoji, label, description, DEFAULT/EXTRA badge, and checkmark for active.

### How it works — dynamic, never hardcoded
- Tab colors come from `NavThemes.TAB_COLORS` map — adding new tabs auto-inherits default color
- Theme preference persisted in localStorage, read on init
- Body class controls all visual styling — zero inline style overrides for themes
- Sidebar active items get `--blob-color` CSS var set dynamically per tab

### Files modified
1. `src/index.html` — AppState.navTheme, NavThemes engine, NavThemes.init() in initApp(), sidebar theme picker, sidebar active item color vars
2. `src/styles.html` — 5 theme CSS rulesets (~200 lines), theme picker UI styles
3. `src/steward_view.html` — renderBottomNav rewritten with blob indicator, per-tab colors, theme-aware rendering

### IMPORTANT rules preserved
- Everything dynamic, nothing hardcoded
- Column identification by header name (not affected)
- No deletion/overwriting of manually entered data (not affected)
- Config tab remains single source of truth (not affected)

---

## Chapter-Based Survey UI — Phase 2 Implementation (March 2026)

### What was added
Enhanced survey wizard with chapter intro screens, animated progress, mini celebrations between sections, and confetti on completion.

### Architecture — 3-phase section flow
Each survey section now cycles through three phases:
1. **Intro phase**: Full-screen chapter card with emoji, section title, tagline, question count, estimated time, and "Start/Continue" button
2. **Questions phase**: Original multi-type question rendering (slider-10, radio, checkbox, dropdown, paragraph) — preserved intact with all branching logic
3. **Celebration phase**: Mini celebration screen with confetti, completion stats (answered/progress/time), and "Next Section" button

### Visual enhancements
- **Animated progress bar**: CSS glow animation (`surveyGlowBar`), shimmer effect, smooth width transition
- **Step dots**: Active dot expands, completed dots fill with accent color, pending dots are dimmed
- **Confetti**: 30-particle CSS animation on section completion and final submission
- **Section visuals**: Each section gets an emoji, color, and tagline from `_sectionVisuals` map
- **Celebration stats**: Animated count-up cards showing answered count, % progress, elapsed time
- **Thank you screen**: Animated spinning checkmark, pulse rings, stat cards, confetti explosion

### CSS animations added (styles.html)
`surveySlideUp`, `surveyPopIn`, `surveyBounce`, `surveyFloat`, `surveyGlowBar`, `surveyShimmer`, `surveyCelebSpin`, `surveyPulseRing`, `surveyConfetti`, `surveyCountUp`, `surveyEmojiWiggle`

### Integration preserved
- All server calls untouched: `dataGetSurveyQuestions`, `dataGetSurveyStatus`, `dataSubmitSurveyResponse`
- Branch logic (`_sectionBranchRules`, `getActiveSectionKeys`) fully preserved
- Draft saving (`window._surveyDraft`) fully preserved
- All question types (slider-10, radio, radio-branch, checkbox, dropdown, paragraph) untouched
- Validation per section untouched
- Dues-paying gate untouched
- Error handling and retry logic untouched

### Files modified
1. `src/member_view.html` — `_renderSurveyWizard` enhanced with phase system, chapter intros, celebrations, confetti helper, animated thank-you
2. `src/styles.html` — ~150 lines of survey animation CSS

---

## Survey Results 5-View Dashboard — Phase 3 Implementation (March 2026)

### What was added
Replaced single-view survey results with 5-tab dashboard: Overview, Heatmap, Sections, Balance (Radar), and Detail.

### View descriptions
| Tab | ID | What it shows |
|-----|----|---------------|
| 📖 Overview | story | Hero score, scrollable ranked section strip, strongest/weakest insight cards, all-sections list with mini bars |
| 🔥 Heatmap | heatmap | Color-intensity cells per section (opacity scales with score), legend |
| 📈 Sections | sections | Conic progress ring for overall, expandable cards per section with score bars |
| 🎯 Balance | radar | SVG radar/spider chart with data polygon, axis lines, grid rings, emoji labels, ranked list below |
| 🔍 Detail | detail | Filterable section list with chapter pills, per-section cards with score bars |

### Architecture
- Tab bar using existing `.sub-tab` CSS classes for consistency
- All 5 views render inside a shared `viewContainer` div, swapped on tab click
- Data source: `dataGetSatisfactionSummary(SESSION_TOKEN)` — unchanged
- Section data dynamically normalized into `sectionList` array with emoji/color per section
- `scoreColor()` function maps 1-10 scale to 5 color tiers
- SVG radar chart rendered via `document.createElementNS` for full browser compat
- All animations use survey CSS from Phase 2 (`surveySlideUp`, etc.)

### Server integration
- No server-side changes — uses existing `dataGetSatisfactionSummary` endpoint
- Response threshold check (10 minimum) preserved
- Privacy badge preserved
- Error handling preserved
- Future: when `dataGetSurveyHistory` endpoint is added, Heatmap and Sections views can show cross-round data

### Files modified
1. `src/member_view.html` — `renderSurveyResultsPage` rewritten (~220 lines, replaces ~80 lines)

---

## Survey Results 4 Extra Views — Phase 4 Implementation (March 2026)

### What was added
4 additional tabs in the survey results dashboard: Alerts, Participation/Turnout, Compare, and Meeting View.

### View descriptions
| Tab | ID | What it shows |
|-----|----|---------------|
| 🚨 Alerts | alerts | Auto-flagged issues: critical (below 5.0), warning (below 6.5), gap detection, low response warning. Summary badges at top, color-coded cards per alert. |
| 👥 Turnout | participation | Response count hero, conic progress ring for participation rate (uses AppState.memberCount), per-section response counts, "Boost Participation" tips card. |
| 📊 Compare | compare | Side-by-side section comparison with dropdown pickers. Shows dual score cards, difference analysis (modest/significant/major gap), visual bar comparison. |
| 📋 Meeting | meeting | Projector-ready: large hero score (56px font), ranked section bars with rank badges, key takeaways list, auto-generated discussion points based on scores. |

### Architecture
- All 4 views use the same `dataGetSatisfactionSummary` data — no new server endpoints
- Alerts auto-generate from score thresholds (configurable: `lowThreshold = 5.0`, `warnThreshold = 6.5`)
- Participation uses `AppState.memberCount` for rate calculation (gracefully hides if unavailable)
- Compare uses native `<select>` dropdowns for section picker
- Meeting view uses larger fonts/spacing optimized for screen projection
- All animations reuse Phase 2 survey CSS (`surveySlideUp`, `surveyCountUp`, etc.)

### Total tab count in Survey Results: 9
📖 Overview, 🔥 Heatmap, 📈 Sections, 🎯 Balance, 🔍 Detail, 🚨 Alerts, 👥 Turnout, 📊 Compare, 📋 Meeting

### Files modified
1. `src/member_view.html` — 4 new view functions (~320 lines), TABS array expanded, renderView switch expanded

---

## 🚫 Directory Tab Removed from Member View (2026-03-09)

### What changed
- Removed the "Directory" tab (`id: 'contact'`) from the member sidebar in `src/index.html`
- Removed the corresponding `case 'contact'` switch entry in the member routing logic
- Removed the `contact` color mapping from `TAB_COLORS`
- Any existing deep-links to `#contact` for members will now fall through to the default (Home)

### What was NOT removed
- **Steward Directory** (`id: 'stewarddirectory'`) — remains in member sidebar as a utility link for finding/selecting a steward. Labeled dynamically as `CONFIG.stewardLabel + ' Directory'`, not "Directory".
- **"Find a Steward" button** on Home tab (`member_view.html` line ~246) — contextual action when no steward is assigned, not a standalone tab.
- **`renderStewardContact()`** function in `member_view.html` — still used by Steward Directory link.

### Reasoning
Members should not have a general "Directory" tab. The Steward Directory (which shows steward contacts for member selection) is a separate, role-appropriate utility.

### Files modified
1. `src/index.html` — 3 edits: sidebar tab removed, switch case removed, color mapping removed

---

## 🔧 FIX LOG — 2026-03-09 (Session 2): Org Overdue / Org Due <7d Still Zero

### Problem
"Org Overdue" and "Org Due <7d" KPI cards still showed "..." or 0 after previous fix session. The async `dataGetGrievanceStats` call was running and returning `available: true`, but `overdueCount` and `dueSoonCount` were always 0.

### Root Cause
`HEADERS.grievanceDeadline` aliases (`['deadline', 'next deadline', 'due date']`) **do not match any actual column header** in the Grievance Log sheet. The sheet uses `Next Action Due`, `Filing Deadline`, and `Days to Deadline` — none of which were in the alias list.

Result: `_getVal(row, colMap, HEADERS.grievanceDeadline, null)` returned `null` on every row → `deadlineDays` was always `null` → overdue auto-detection never fired → status stayed as whatever the sheet Status column said (e.g. "Step I", "Open") → `getGrievanceStats` `overdueCount` stayed 0.

### Fix Applied (src/21_WebDashDataService.gs)

**Change 1 — Line 60: Expand `grievanceDeadline` aliases**
```
// OLD:
grievanceDeadline: ['deadline', 'next deadline', 'due date'],
// NEW:
grievanceDeadline: ['deadline', 'next deadline', 'due date', 'next action due', 'filing deadline'],
```
`Next Action Due` is a Date column — existing date parsing logic handles it correctly.

**Change 2 — After deadline date block in `_buildGrievanceRecord`:**
Added `Days to Deadline` numeric fallback. If no date column matched, reads the computed `Days to Deadline` column:
- If value is the text `"Overdue"` → `deadlineDays = -1`
- If value is numeric (e.g. `-5`, `3`) → `deadlineDays = Math.ceil(value)`

This two-layer approach ensures coverage whether a case has `Next Action Due` populated or only the computed `Days to Deadline` column.

### Priority Order for Deadline Resolution
1. `Next Action Due` date → compute `deadlineDays` from today
2. `Filing Deadline` date (fallback) → compute `deadlineDays` from today
3. `Days to Deadline` number/text → use directly

### ⚠️ Reminder
Column header matching is case-insensitive via `_buildColumnMap`. `'next action due'` matches `Next Action Due` in the sheet.

---

## 🔬 AUDIT LOG — 2026-03-09: Full Field-to-Column Match Audit

### Purpose
Verified every `HEADERS` alias in `21_WebDashDataService.gs` against actual column headers in `MEMBER_HEADER_MAP_` and `GRIEVANCE_HEADER_MAP_` (both in `01_Core.gs`).

### Method
Extracted all aliases → compared against normalized actual headers → identified misses → traced each miss through `_buildUserRecord` and frontend rendering to determine functional impact.

### Results

#### Member Aliases
| Alias Key | Status | Notes |
|---|---|---|
| memberEmail | ✅ Match | `Email` |
| memberName | ✅ Safe | No "Name" column — falls back to First+Last concat |
| memberFirstName | ✅ Match | `First Name` |
| memberLastName | ✅ Match | `Last Name` |
| memberRole | ✅ Safe | No "Role" column — `Is Steward` fallback covers steward detection |
| memberUnit | ✅ Match | `Unit` |
| memberPhone | ✅ Match | `Phone` |
| memberJoined | ⚠️ Fixed | Added `hire date` alias → now maps to `Hire Date` column |
| memberDuesStatus | ✅ Match | `Dues Status` |
| memberDuesPaying | 🔴 Fixed | No "Dues Paying" column — was returning null → gate bypassed |
| memberId | ✅ Match | `Member ID` |
| memberWorkLocation | ✅ Match | `Work Location` |
| memberOfficeDays | ✅ Match | `Office Days` |
| memberAssignedSteward | ✅ Match | `Assigned Steward` |
| memberIsSteward | ✅ Match | `Is Steward` |
| memberStreet/City/State/Zip | ✅ Match | all exact |
| memberHasOpenGrievance | ✅ Match | `Has Open Grievance?` |
| memberSupervisor | ✅ Match | `Supervisor` |
| memberJobTitle | ✅ Match | `Job Title` |
| memberAdminFolderUrl | ✅ Match | `Member Admin Folder URL` |
| memberSharePhone | ✅ Match | `Share Phone` |

#### Grievance Aliases
| Alias Key | Status | Notes |
|---|---|---|
| grievanceId | ✅ Match | `Grievance ID` |
| grievanceMemberEmail | ✅ Match | `Member Email` |
| grievanceMemberFirstName/LastName | ✅ Match | |
| grievanceStatus | ✅ Match | `Status` |
| grievanceStep | ✅ Match | `Current Step` |
| grievanceDeadline | ✅ Fixed (prev session) | Now maps to `Next Action Due` |
| grievanceFiled | ✅ Match | `Date Filed` |
| grievanceSteward | ✅ Match | `Assigned Steward` |
| grievanceUnit | ✅ Match | `Work Location` |
| grievancePriority | ⚪ No column | Not displayed in grievance cards — no action |
| grievanceNotes | ⚪ No column | Not displayed in grievance cards — no action |
| grievanceIssueCategory | ✅ Match | `Issue Category` |
| grievanceResolution | ✅ Match | `Resolution` |
| grievanceDateClosed | ✅ Match | `Date Closed` |

### Fixes Applied (this session)

**Fix 1 — `memberDuesPaying` gate bypass (HIGH)**
No `Dues Paying` column exists → `duesPaying` was always `null` → treated as paying → all members bypassed dues gate for Org Chart, Survey, Polls, Union Stats, POMS, Resources.

Added fallback derivation from `Dues Status` column (which exists):
- `'current'`, `'active'`, `'paid'`, or any unrecognized → `true` (paying)
- `'past due'`, `'inactive'`, `'delinquent'`, `'lapsed'`, `'non-paying'`, `'no'` → `false` (not paying)
- Column absent entirely → `null` (benefit of the doubt, treated as paying)

File: `src/21_WebDashDataService.gs`, `_buildUserRecord()`, duesPaying block

**Fix 2 — `memberJoined` shows "—" (LOW)**
Added `'hire date'` as alias for `memberJoined` → now maps to `Hire Date` column in sheet.
File: `src/21_WebDashDataService.gs`, line 35

### Columns With No Alias (intentional — not read by DataService)
These exist in the sheet but are read directly by non-DataService code (GS functions, formulas) and not needed in the webapp frontend: `Contact Steward`, `Contact Notes`, `Cubicle`, `Days to Deadline` (member sheet), `Employee ID`, `Grievance Status` (member sheet), `Interest: *`, `Last Virtual/In-Person Mtg`, `Manager`, `Open Rate %`, `PIN Hash`, `Preferred Communication`, `Recent Contact Date`, `Start Grievance`, `Volunteer Hours`, `⚡ Actions`. Grievance sheet: all step date columns, reminder columns, acknowledgment columns, drive folder columns — read by GS backend only.
## ⏱️ Test Runner Timeout Fix (2026-03-09, v4.25.6)

### Root cause
GAS has a hard 6-minute execution limit. The test runner had a 5-minute soft guard that checked BETWEEN tests, but:
1. 82 tests × ~4s each = ~5.5 min cumulative
2. 16 calls to `SpreadsheetApp.getActiveSpreadsheet()` (each a network round-trip)
3. `authsweep` suite calls 16+ real endpoint functions, each doing `checkWebAppAuthorization()` → `getUserRole_()` → sheet read
4. A single slow test at minute 5:01 would push past 6:00 before the next guard check fired

### Fixes applied
1. **Timeout lowered: 5min → 3.5min** — gives 2.5min safety margin instead of 1min
2. **Cached spreadsheet reference** — `_getCachedSS()` replaces 12 individual `SpreadsheetApp.getActiveSpreadsheet()` calls in test functions. Infrastructure functions (menu handlers, email sender) still use direct calls since they run in separate execution contexts.
3. **Cache reset** — `_testSS_ = null` at start of `runAll()` prevents stale references between runs
4. **SPA timeout handling** — timeout warning banner with guidance to use suite filter, skipped count card, suite headers show skipped count, failure handler detects "execution time" error and shows helpful message

### Files modified
1. `src/30_TestRunner.gs` — timeout constant, SS cache, cache reset in runAll
2. `src/steward_view.html` — timeout banner, skipped card, suite skipped count, failure message

## 🐛 onOpen Deferred Trigger Bug Fix (2026-03-10, v4.25.7)

### Root Cause
Two compounding bugs in `onOpen()` (`10_Main.gs`) caused `onOpenDeferred_()` to never run on sheet open:

**Bug 1 — ScriptApp not allowed in simple triggers:**
`ScriptApp.getProjectTriggers()` requires authorization that GAS simple triggers do not have. The call threw silently, fell into the catch block, and attempted to call `onOpenDeferred_()` inline — but that function also calls `ScriptApp.getProjectTriggers()` internally (to self-clean), which failed for the same reason.

**Bug 2 — finally block race condition:**
Even if Bug 1 didn't exist, the `finally { cleanUpOnOpenTrigger_() }` block ran synchronously immediately after trigger creation — deleting the 1-second deferred trigger before it could fire. GAS timer triggers need to be allowed to tick; `finally` does not wait.

**Practical impact:** `syncColumnMaps()`, `enforceHiddenSheets()`, tab colors, and the "Dashboard loaded" toast never executed on sheet open.

### Fixes Applied
1. **`onOpen()` simplified** — now only does cache clear + `createDashboardMenu()`. No ScriptApp calls, no finally block. Correct GAS simple trigger pattern.
2. **`setupOpenDeferredTrigger()`** added to `08e_SurveyEngine.gs` — installs `onOpenDeferred_` as an installable `onOpen` trigger via `ScriptApp.newTrigger().forSpreadsheet().onOpen().create()`. Safe to re-run (removes existing first).
3. **`menuInstallSurveyTriggers()`** updated — now also calls `setupOpenDeferredTrigger()` so "Install ALL" covers this fix.
4. **Menu item added** — Admin → ⏱️ Triggers → 🔓 Install onOpen Deferred Trigger

### ⚠️ Action Required After Deploy
Run **Admin → ⏱️ Triggers → ✅ Install ALL Survey Triggers** (or 🔓 Install onOpen Deferred Trigger alone) ONCE per Google account that uses the sheet. Installable triggers are per-user.

### Files Modified
1. `src/10_Main.gs` — `onOpen()` stripped to menu-only
2. `src/08e_SurveyEngine.gs` — `setupOpenDeferredTrigger()` added, `menuInstallSurveyTriggers()` updated
3. `src/03_UIComponents.gs` — new menu item in Triggers submenu
4. `dist/` — all three files mirrored
