# AI REFERENCE DOCUMENT — DDS-Dashboard
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-02-25
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
**Architecture:** 36 source `.gs` files + 8 `.html` files in `src/` → built into single `dist/ConsolidatedDashboard.gs` via `node build.js`.
**Current build:** ~70,586 lines / ~2,907 KB (limit: 6MB).
**Web App:** Served via `doGet()` using inline HTML (`HtmlService.createHtmlOutput()`). Does NOT use `createTemplateFromFile()`.
**DDS Apps Script ID:** `18hHHX-4E_ykGCqu_EDwKCwqY9ycyRgPtOmguacsxnVZ4YsRh-YETODiu`
**UT Apps Script ID:** `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`

### ⚠️ Key Reminders
- **Critical rules** (dynamic-only, 1-indexed columns, escapeHtml, etc.) → **See CLAUDE.md**
- **Sync rules & WT exclusions** → **See SYNC-LOG.md**
- **`dist/ConsolidatedDashboard.gs` is auto-generated.** Never edit directly.
- **`web-dashboard/` folder is LEGACY/ORPHANED.** Do not deploy or integrate it.
- **CLASP rootDir:** `./dist` — only `dist/` contents go to Apps Script.
- **Deploy:** `npm run deploy` (lint + test + build:prod + clasp push). Must run locally (requires Google OAuth).
- **After any merge to Main:** Remind user to run `npm run deploy`. Agent cannot run clasp remotely.

---

## 📁 ARCHITECTURE MAP

### Build Pipeline
```
src/*.gs (36 files) + src/*.html (8 files)
    → build.js → dist/ConsolidatedDashboard.gs (single file)
                  dist/appsscript.json (manifest)
    → clasp push → Google Apps Script project
```

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
| `08a-d` | Sheet utils | Setup, search, forms, audit formulas |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Main entry, sheet creation, forms, sync |
| `11_-17_` | Features | CommandHub, self-service, meetings, events, correlation |
| `18_` | Workload Tracker | **DDS ONLY** — excluded from Union-Tracker |
| `19_-24_` | Web Dashboard SPA | Auth, config reader, data service, app entry, portal sheets, weekly questions |

### HTML Files in src/
| File | Purpose |
|------|---------|
| `index.html` | SPA entry point (unified dashboard) |
| `steward_view.html` | Steward command center |
| `member_view.html` | Member dashboard |
| `auth_view.html` | Login/auth page |
| `error_view.html` | Error display |
| `styles.html` | Shared CSS |
| `MultiSelectDialog.html` | Multi-select dropdown UI |
| `WorkloadTracker.html` | WT portal (**DDS only**) |

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
├── ?page=workload → Workload tracker (DDS only)
├── ?page=checkin  → Meeting check-in (v4.11.0)
├── ?page=resources → SPA with resources tab pre-selected (v4.11.0)
├── ?page=notifications → SPA with notifications tab pre-selected (v4.12.0)
└── Deep-link: ?page=X → SPA reads PAGE_DATA.initialTab → _handleTabNav()
```

### Authentication System
- **Steward access:** Google account email matched against authorized list via `checkWebAppAuthorization('steward')`
- **Member access:** Google account email OR PIN-based login
- **SPA auth:** Google SSO + magic link (19_WebDashAuth.gs)
- **Dashboard auth toggle:** `isDashboardMemberAuthRequired()` — when enabled, all dashboard pages require member login
- **Auth config:** `ScriptProperties` (no manual setup required — `initWebDashboardAuth()` handles first-time)

### HTML Serving Method
The consolidated file uses `HtmlService.createHtmlOutput()` with **inline HTML strings** built by functions like `getUnifiedDashboardHtml()`, `getWebAppDashboardHtml()`, etc. It does NOT use `createTemplateFromFile()` or separate `.html` files (that's the orphaned `web-dashboard/` architecture).

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

### Hidden Sheets (v4.12.2)
- `_Weekly_Questions` — weekly engagement questions
- `_Contact_Log` — steward-member interaction log (8 columns)
- `_Steward_Tasks` — steward task management (10 columns)

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
5. **Never edit `dist/ConsolidatedDashboard.gs` directly.** Edit `src/*.gs` and run `npm run build`.
6. **Test with `npm run ci`** before pushing.
7. **Deploy with `npm run deploy`** (lint + test + prod build + clasp push).
8. **Build is ~2.9MB / 6MB limit.** Monitor growth — you have ~50% headroom.
9. **The `doGet()` source is in `src/05_Integrations.gs`** (routes) and `src/22_WebDashApp.gs` (SPA entry).
10. **Do not duplicate information across reference docs.** Each doc has one canonical purpose (see table at top).
