# AI REFERENCE DOCUMENT — SolidBase
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-03-16
# Consolidation note: On 2026-03-16, 97 historical sections (v4.18–v4.25)
# were archived to docs/AI_REFERENCE_ARCHIVE.md. No information was deleted.

---

## 📖 HOW TO USE THIS FILE

Read these files **in this order** when onboarding to this codebase:

| Order | File | What it covers |
|-------|------|----------------|
| 1 | **CLAUDE.md** | Critical rules, column constants, security patterns, config write paths, coding conventions, git conventions |
| 2 | **This file (AI_REFERENCE.md)** | Project overview, architecture map, LLM-specific context, error log, protected code |
| 3 | **SYNC-LOG.md** | SolidBase sync sync flow, Workload Tracker exclusion registry |
| 4 | **CHANGELOG.md** | Full version history (Keep a Changelog format) |
| 5 | **FEATURES.md** | Detailed feature documentation |
| 6 | **COLUMN_ISSUES_LOG.md** | Recurring column bugs — READ if touching column-related code |
| 7 | **CODE_REVIEW.md** | Canonical security/code review |
| 8 | **DEVELOPER_GUIDE.md** | Developer onboarding |

**Do NOT duplicate content from those files here.** If you need to add context an LLM would need that doesn't fit those files, add it here.

---

## 🏗️ PROJECT OVERVIEW

**What:** Google Apps Script application for union steward grievance tracking, member management, and reporting.
**Repo:** `Woop91/SolidBase` (private). Default branch: `Main` (capital M).
**Mirror:** `Woop91/Union-Tracker-` (public). See SYNC-LOG.md for exclusion rules.
**Deployed via:** CLASP (`clasp push`) to Google Apps Script, bound to a Google Sheet.
**Target users:** Union stewards (power users) and members (casual users) at Organization (Union).
**Architecture:** 42 source `.gs` files + 7 `.html` files in `src/` → copied individually to `dist/` via `node build.js`.
**Current build:** 42 `.gs` + 7 `.html` files in `dist/` (individual file mode, NOT consolidated).
**Web App:** Served via `doGet()` using inline HTML (`HtmlService.createHtmlOutput()`). Does NOT use `createTemplateFromFile()`.
**Apps Script ID:** `[REDACTED]
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
| `04a-d` | UI modules | Menus, accessibility, interactive/executive dashboards |
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

### ~~Dashboard Modal Popup — `src/04c_InteractiveDashboard.gs`~~ REMOVED v4.31.0
All Interactive Dashboard code removed. Functionality consolidated into Steward Dashboard (04d) and web app.

### Member Satisfaction Dashboard — `src/09_Dashboards.gs`
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
| 2026-02-25 | Memory had default branch as `staging` | Incorrect memory entry | Corrected to `Main` | ✅ |
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

## 📂 HISTORICAL DEVELOPMENT LOG

All historical entries (v4.18.1–v4.25.9 bug fixes, feature implementations, migration notes)
have been archived to `docs/AI_REFERENCE_ARCHIVE.md`. Refer there for past development context.

---

## SESSION: 2026-03-21 — Orphaned Function Audit & Archive Fix

### Changes Made

**v4.33.1 — Function Audit & Daily Trigger Wiring**

**Files modified:**
- `src/10_Main.gs` — `dailyTrigger()` and `handleSecurityAudit_()`
- `src/21_WebDashDataService.gs` — `getPendingGrievanceFeedback()`, `submitGrievanceFeedback()`

---

### 1. `dailyTrigger()` — 4 orphaned functions wired (`10_Main.gs`)

**Problem:** Four maintenance functions existed and were fully implemented but were never called from any trigger. They ran only if manually invoked.

| Function | Origin | Effect of being orphaned |
|---|---|---|
| `archiveClosedGrievances(90)` | `06_Maintenance.gs` v4.30.0 | Closed grievances accumulated in active sheet forever |
| `dailyAuditArchive()` | `06_Maintenance.gs` v4.30.0 | Audit log grew unbounded; no Drive CSV backups |
| `sendDailySecurityDigest()` | `00_Security.gs` v4.8.1 | Security digest never sent; CRITICAL events silently discarded |
| `authCleanupExpiredTokens()` | `19_WebDashAuth.gs` v4.22.9 | Expired session tokens accumulated in PropertiesService |

**Fix:** Each function now called inside `dailyTrigger()` with `typeof` guard + `try/catch`.

---

### 2. Mass deletion email alert fixed (`10_Main.gs` — `handleSecurityAudit_()`)

**Problem:** `onEdit` is a GAS simple trigger. `MailApp.sendEmail()` is unavailable in this context. `recordSecurityEvent(CRITICAL)` called `sendSecurityAlertEmail_()` which silently failed. The CRITICAL event was logged to audit but **never emailed**. Additionally, `queueSecurityDigestEvent_()` only runs for HIGH severity — CRITICAL events never entered the queue at all.

**Fix:** `handleSecurityAudit_()` now explicitly calls `queueSecurityDigestEvent_()` before `recordSecurityEvent()`. PropertiesService (used by the queue) works in simple trigger context. Admin receives email within 24h when `sendDailySecurityDigest()` runs in `dailyTrigger()`.

---

### 3. Archive data accessible to webapp statistics (`21_WebDashDataService.gs`)

**Already using `_getAllGrievanceData()` (active + archive merged):**
- `getMemberGrievanceHistory()` ✅
- `getDashboardStats()` ✅
- `getGrievanceHotspots()` ✅

**Intentionally active-only (correct by design):**
- `getStewardCases()` — shows open workload; closed/archived cases excluded intentionally
- `getMemberGrievances()` — filters out closed statuses anyway; archive (closed-only) adds nothing

**Fixed to use `_getAllGrievanceData()`:**
- `getPendingGrievanceFeedback()` — needed archive because cases auto-archived within 14-day feedback window would silently drop the prompt
- `submitGrievanceFeedback()` — needed archive to validate ownership of an archived closed case; without this, `'Grievance not found or not closed'` error for any archived case

---

### CORE RULE REMINDER
- Everything dynamic — never hardcode sheet names, column indices, or config values
- Never delete or overwrite manually entered data
- Column identification by header name via `_findColumn(colMap, HEADERS.xxx)`
- All mutations go to ACTIVE sheet only; statistics/reads use `_getAllGrievanceData()` for complete view
- `archiveClosedGrievances()` threshold = 90 days (configurable via `daysOld` param)

---

## SESSION: 2026-03-21 (continued) — Config-driven retention thresholds

### Changes Made

**v4.33.2 — Dynamic archive thresholds via Config tab**

**Files modified:**
- `src/01_Core.gs` — `CONFIG_HEADER_MAP_`: added `GRIEVANCE_ARCHIVE_DAYS` + `AUDIT_ARCHIVE_DAYS`
- `src/10a_SheetCreation.gs` — `createConfigSheet`: seeded both keys with default `90`
- `src/10_Main.gs` — `dailyTrigger()`: reads both thresholds from Config before calling archive functions

### How it works

`dailyTrigger()` reads `CONFIG_COLS.GRIEVANCE_ARCHIVE_DAYS` and `CONFIG_COLS.AUDIT_ARCHIVE_DAYS`
via `getConfigValue_()`. Falls back to 90 if blank, zero, or non-numeric.
Passes values to `archiveClosedGrievances(grievanceArchiveDays)` and `archiveOldAuditLogs_(auditArchiveDays)` respectively.
Admins change retention by editing the Config tab — no code change required.

### Config tab entries

| Header | Key | Default | Purpose |
|---|---|---|---|
| `Grievance Archive Days` | `GRIEVANCE_ARCHIVE_DAYS` | 90 | Days after closure before grievance moves to `_Archive_Grievances` |
| `Audit Log Archive Days` | `AUDIT_ARCHIVE_DAYS` | 90 | Days before audit log entries are exported to Drive CSV and pruned |

### CORE RULE REMINDER
Everything dynamic — config values always read at runtime from Config tab.
Never hardcode thresholds, deadlines, or durations.
