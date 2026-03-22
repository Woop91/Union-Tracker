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
| 3 | **SYNC-LOG.md** | SolidBase sync history and exclusion rules |
| 4 | **CHANGELOG.md** | Full version history (Keep a Changelog format) |
| 5 | **FEATURES.md** | Detailed feature documentation |
| 6 | **COLUMN_ISSUES_LOG.md** | Recurring column bugs — READ if touching column-related code |
| 7 | **CODE_REVIEW.md** | Canonical security/code review |
| 8 | **DEVELOPER_GUIDE.md** | Developer onboarding |

**Do NOT duplicate content from those files here.** If you need to add context an LLM would need that doesn't fit those files, add it here.

---

## 🏗️ PROJECT OVERVIEW

**What:** Google Apps Script application for union steward grievance tracking, member management, and reporting.
**Repo:** `Woop91/SolidBase` (public). Default branch: `Main` (capital M).
**Deployed via:** CLASP (`clasp push`) to Google Apps Script, bound to a Google Sheet.
**Target users:** Union stewards (power users) and members (casual users).
**Architecture:** 45 source `.gs` files + 8 `.html` files in `src/` → copied individually to `dist/` via `node build.js`. Production build excludes DevTools + DevMenu (43 .gs + 8 .html).
**Current build:** 43 `.gs` + 8 `.html` files in `dist/` production (individual file mode, NOT consolidated).
**Web App:** Served via `doGet()` using inline HTML (`HtmlService.createHtmlOutput()`). Does NOT use `createTemplateFromFile()`.
**Apps Script ID:** `[REDACTED — see .clasp.json]`

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
src/*.gs (45 files) + src/*.html (8 files)
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
| `19_-24_` | Web Dashboard SPA | Auth, config reader, data service, app entry, portal sheets, weekly questions |
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
As of v4.13.0, the build outputs individual `.html` files to `dist/`, enabling both `HtmlService.createHtmlOutput()` (inline strings) and `createTemplateFromFile()`/`createHtmlOutputFromFile()` (file-based). The SPA modules (19-24, 26-29) use file-based HTML. Legacy modules (04-13) still use inline HTML strings.

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
| 2026-03-19 | Survey Section 13 always active (not toggled like RTO_CHANGE) | Workforce mobility is a permanent strategic metric, not a time-limited event |
| 2026-03-19 | Q85 (transfer awareness) folded into Q83 "Limited Advancement or Transfer Opportunities" | Reduces question count; awareness is captured as a leaving reason — more actionable |
| 2026-03-19 | Q82 maxSel=2, Q83 maxSel=3 (checkbox) | Multi-select enforced client-side by survey wizard; maxSel col read by branch logic |
| 2026-03-19 | Workforce Retention Avg = q84 only (not q80/q81) | q80/q81 are categorical radios — not numeric. Averaging them is meaningless. q84 is the sole slider-10 in the section |
| 2026-03-19 | getSatisfactionSummary picks up WORKFORCE_RETENTION automatically | Iterates slider-10 questions dynamically by sectionKey — new sections need no code changes |
| 2026-02-24 | Notifications in separate sheet (not Member Directory column) | Notifications are ephemeral; don't pollute member data |
| 2026-02-24 | Dismissed_By as comma-separated in single cell | Avoids per-member rows, scales to thousands |
| 2026-02-24 | Notification auto-ID scans existing IDs for max number | Gap-safe even if rows deleted |
| 2026-02-25 | ConfigReader reads column-based Config (not row-based key-value) | Aligns with existing CONFIG_COLS system; old reader was incompatible |
| 2026-02-25 | Default accent hue changed from 250 (blue) → 30 (amber) | Warm palette matches union identity, distinguishes from generic dashboards |
| 2026-02-25 | SPA deep-links (?page=X → initialTab) with standalone HTML fallback | Consistent SPA experience, but graceful degradation if SPA unavailable |
| 2026-02-25 | `initWebDashboardAuth()` auto-configures on first run | No manual ScriptProperties setup required — reduces deployment friction |
| 2026-02-25 | Switched from consolidated single-file build to individual-file build | GAS needs separate `.html` files for `createTemplateFromFile()` and `createHtmlOutputFromFile()`. Individual files also easier to debug in GAS editor. |
| 2026-02-25 | Added `25_WorkloadService.gs` alongside `18_WorkloadTracker.gs` | 25_ was SPA-integrated (SSO auth), 18_ was standalone portal (PIN auth). Both removed from SolidBase (org-specific feature). |

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
| 2026-03-19 | Claude (claude.ai) | v4.32.0: Workforce Mobility & Retention survey section added — Sections 13 + 13A, q80–q86 (q85 removed/folded), Looker columns, getSatisfactionSummary auto-includes via dynamic section key | `dist/10b_SurveyDocSheets.gs`, `src/10b_SurveyDocSheets.gs`, `dist/12_Features.gs`, `src/12_Features.gs`, `dist/01_Core.gs`, `src/01_Core.gs`, `AI_REFERENCE.md` |

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
3. **Read SYNC-LOG.md** — SolidBase sync history and exclusion rules.
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

## REFACTOR LOG — ADHD → ComfortView rename (2026-03-21)

**Reason:** Internal function names using `ADHD` were inconsistent with the user-facing "Comfort View" branding.

**Files changed:** `src/03_UIComponents.gs`, `src/04b_AccessibilityFeatures.gs`, `src/07_DevTools.gs`, `src/10b_SurveyDocSheets.gs`, `dist/`

**Rename map:**
| Old | New |
|---|---|
| `getADHDSettings` | `getComfortViewSettings` |
| `getDefaultADHDSettings_` | `getDefaultComfortViewSettings_` |
| `saveADHDSettings` | `saveComfortViewSettings` |
| `applyADHDSettings` | `applyComfortViewSettings` |
| `resetADHDSettings` | `resetComfortViewSettings` |
| `toggleGridlinesADHD` | `toggleGridlinesComfortView` |
| `setupADHDDefaults` | `setupComfortViewDefaults` |
| `showADHDControlPanel` | `showComfortViewControlPanel` |
| `'adhdSettings'` (prop key) | `'comfortViewSettings'` |

**DATA MIGRATION NOTE:** Any existing users who had `adhdSettings` stored in PropertiesService will lose their saved Comfort View preferences on next load. They will fall back to `getDefaultComfortViewSettings_()`. This is acceptable one-time cost — defaults are safe.

---

## DEFERRED FEATURE — Grievance Module Toggle (2026-03-21)

**Status:** NOT IMPLEMENTED — documented for future build. No code has been changed.

**Purpose:** Allow the grievance module to be toggled on/off via the Config tab without touching code. Useful for orgs that use SolidBase but do not have a grievance tracking workflow.

**Precedent:** Follows the exact same pattern as `ENABLE_CORRELATION` (`17_CorrelationEngine.gs` + `CONFIG_HEADER_MAP_`).

---

### Config Change

**File:** `src/01_Core.gs`
**Location:** End of `CONFIG_HEADER_MAP_` array, after `AUDIT_ARCHIVE_DAYS` entry.
**What to add:**
```js
// ENABLE_GRIEVANCES: set to 'no' to hide all grievance features from the SPA.
// Default blank or 'yes' = enabled (fully backward-compatible).
{ key: 'ENABLE_GRIEVANCES', header: 'Enable Grievances' }
```

---

### ConfigReader Change

**File:** `src/20_WebDashConfigReader.gs`
**Location:** Inside `getConfig()`, in the `config = { ... }` object.
**What to add:**
```js
enableGrievances: _readRow(CONFIG_COLS.ENABLE_GRIEVANCES) !== 'no',
```
**Why:** Exposes flag to the SPA client as a boolean. Blank or `'yes'` → `true`. Only explicit `'no'` disables it.

---

### SPA Sidebar — `index.html`

**File:** `src/index.html`
**Location:** `_getSidebarTabs(role)` function (~line 2027).
**Steward tabs to gate:** `cases` and `newgrievance`
**Member tabs to gate:** `cases`
**What to add:**
```js
// Steward list — wrap cases + newgrievance:
...(CONFIG.enableGrievances ? [
  { id: 'cases',        icon: '📋', label: 'Cases' },
  { id: 'newgrievance', icon: '📝', label: 'New Grievance' },
] : []),

// Member list — wrap cases:
...(CONFIG.enableGrievances ? [
  { id: 'cases', icon: '📋', label: 'My Cases' },
] : []),
```

---

### Steward View — `steward_view.html`

Three guards needed:

1. **Grievance stats block** on the home dashboard (~line 228, `dataGetGrievanceStats` call).
   Wrap with: `if (CONFIG.enableGrievances) { ... }`

2. **"Start Grievance" button** in member profile panel (~line 1254).
   Wrap with: `if (CONFIG.enableGrievances) { ... }`

3. **"Has Open Case" filter option** in member list search (~line 679 and ~line 800).
   Wrap with: `if (CONFIG.enableGrievances) { ... }`

---

### Member View — `member_view.html`

Two guards needed:

1. **Active case KPI chip** in the member dashboard KPI strip.
   Wrap with: `if (CONFIG.enableGrievances) { ... }`

2. **Open grievance badge** on member cards (~`hasOpenGrievance` references).
   Wrap with: `if (CONFIG.enableGrievances) { ... }`

---

### Server-Side Gate — `21_WebDashDataService.gs`

All `dataGet*Grievance*` and `dataStartGrievance*` / `dataAdvanceGrievance*` wrapper functions should early-return when disabled:
```js
if (getConfigValue_(CONFIG_COLS.ENABLE_GRIEVANCES) === 'no') {
  return { error: 'Grievance module is disabled.' };
}
```
**Note:** Return a soft error object (not a `403` throw) so the SPA can handle gracefully without crashing the page.

---

### Unresolved Questions (must decide before implementing)

1. When `cases` tab is hidden for members — should deep-links to `#cases` redirect to `home`, or show a "not available" message?
2. Should server-side endpoints return `{ error: '...' }` or `{ grievances: [] }` when disabled? (Soft error vs. empty data)

---

### Version Tag for When Implemented
Suggest codename: **"Grievance Module Toggle"**, version bump: minor (e.g. `4.33.2` or next available).
