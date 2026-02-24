# AI REFERENCE DOCUMENT — DDS-Dashboard
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-02-23

---

## 🏗️ PROJECT OVERVIEW

**What:** Google Apps Script application for union steward grievance tracking, member management, and reporting.
**Repo:** `Woop91/DDS-Dashboard` (private)
**Deployed via:** CLASP (`clasp push`) to Google Apps Script, bound to a Google Sheet.
**Target users:** Union stewards (power users) and members (casual users) at MassAbility DDS (SEIU 509).
**Architecture:** 30 source `.gs` files in `src/` → built into single `dist/ConsolidatedDashboard.gs` (~60K lines, ~2.4MB) via `node build.js`.
**Web App:** Served via `doGet()` in `ConsolidatedDashboard.gs` using inline HTML (NOT template files).

---

## 🔴 CRITICAL RULES — READ FIRST

1. **EVERYTHING MUST BE DYNAMIC.** Never hardcode org names, unit names, column positions, sheet names, or any data that lives in the spreadsheet. Always read from the sheet or the Config tab.
2. **Column constants are 1-indexed** (matching Google Sheets `getRange()`). For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`.
3. **All HTML must use `escapeHtml()`** for dynamic values. Defined in `src/00_Security.gs`. No exceptions.
4. **`dist/ConsolidatedDashboard.gs` is auto-generated.** Never edit it directly. Edit `src/*.gs` files, then run `npm run build`.
5. **The `web-dashboard/` folder is LEGACY/ORPHANED code.** It uses a different architecture (template files + `include()`) that is NOT deployed. The consolidated build is the live system.
6. **Branches:** `Main` (primary, uppercase M), `staging`, `dev`. All currently in sync. Push to `Main` directly (branch protection disabled).
7. **Deploy command:** `npm run deploy` (lint + test + build:prod + clasp push).
8. **CLASP rootDir:** `./dist` — only files in `dist/` are pushed to Google Apps Script.
9. **Google Apps Script file size limit:** ~6MB. Current build is ~2.4MB. Monitor growth.

---

## 📁 ARCHITECTURE

### Build Pipeline
```
src/*.gs (30 files) → build.js → dist/ConsolidatedDashboard.gs (single file)
                                  dist/appsscript.json (manifest)
                    → clasp push → Google Apps Script project
```

### Source File Load Order (numbered prefixes)
| Prefix | Layer | Key Files |
|--------|-------|-----------|
| `00_` | Foundation | `DataAccess.gs`, `Security.gs` |
| `01_` | Core | `Core.gs` — constants, config, utilities |
| `02_` | Data | `DataManagers.gs` — CRUD operations |
| `03_` | UI | `UIComponents.gs` — menus, dialogs |
| `04a-e` | UI modules | Menus, accessibility, interactive/executive/public dashboards |
| `05_` | Integrations | Drive, Calendar, Email |
| `06_` | Maintenance | Admin tools, undo/redo, audit |
| `07_` | DevTools | Dev-only utilities (excluded in prod build) |
| `08a-d` | Sheet utils | Setup, search, forms, audit formulas |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Main entry, sheet creation, forms, sync |
| `11_-17_` | Features | CommandHub, self-service, meetings, events, correlation |

### Web App Routing (doGet)
```
doGet(e)
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
├── ?page=resources → Educational content hub (v4.11.0)
└── (default)      → Unified member dashboard
```

### Authentication System
- **Steward access:** Google account email matched against authorized list via `checkWebAppAuthorization('steward')`
- **Member access:** Google account email OR PIN-based login
- **Dashboard auth toggle:** `isDashboardMemberAuthRequired()` — when enabled, all dashboard pages require member login
- **Auth config stored in:** `ScriptProperties`

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
- See CLAUDE.md for detailed write-path documentation

---

## 🔄 CHANGE LOG

### 2026-02-23 — Deployment Audit & Fixes (by Claude, claude.ai)
**Issues Found & Fixed:**
1. ❌ GitHub token `ghp_FTE8...` in user preferences was expired → User generated new token `ghp_xlP7...`
2. ❌ Second token `ghp_q3Zd...` lacked `repo` scope → User generated third token with correct scope
3. ❌ CI workflow triggers on `main` (lowercase) but branch is `Main` (uppercase) → **FIXED**: Updated `.github/workflows/build.yml`
4. ⚠️ `web-dashboard/` folder contains orphaned code from deleted branch → **FIXED**: Added deprecation banner to `web-dashboard/AI_REFERENCE.md`
5. ✅ `doGet()` exists in `ConsolidatedDashboard.gs` at line 20313 with full routing — verified all 8 dependencies present
6. ✅ `appsscript.json` manifest identical between root and `dist/`
7. ✅ OAuth scopes correct including `userinfo.email`
8. ✅ All branches (Main, staging, dev) synced to same commit
9. ✅ Build is clean — `npm run build` produces identical `dist/` output
10. ✅ All 1295 tests pass (21 suites)
11. ✅ File size 2.4MB / 6MB GAS limit — healthy headroom

**Files Changed:**
- `.github/workflows/build.yml` — Added `Main`, `staging`, `dev` to CI trigger branches
- `AI_REFERENCE.md` — Created (this file)
- `web-dashboard/AI_REFERENCE.md` — Added deprecation notice (folder is orphaned legacy code)

### 2026-02-24 — v4.11.0: Resources Hub + Meeting Check-In Web Route (by Claude, claude.ai)
**New Features:**
1. ✅ `?page=resources` — Educational content hub: Know Your Rights, Grievance Process, FAQ, Forms & Templates
2. ✅ `?page=checkin` — Meeting check-in as standalone web page (reuses existing 14_MeetingCheckIn.gs logic)
3. ✅ `📚 Resources` sheet — Steward-managed content with 12 columns, data validation, 8 starter articles
4. ✅ `RESOURCES_HEADER_MAP_` + `RESOURCES_COLS` — Dynamic column system, registered in `syncColumnMaps()`
5. ✅ `getWebAppResourcesList()` API — Returns visible resources with audience filtering
6. ✅ Design refresh: DM Sans + Fraunces serif fonts, warm navy/earth tones (not generic purple)
7. ✅ `PHASE2_PLAN.md` — Tracks parked features (bulk actions, deadline calendar, etc.)

**Files Changed:**
- `src/01_Core.gs` — Added `SHEETS.RESOURCES`, `RESOURCES_HEADER_MAP_`, `RESOURCES_COLS`, registered in syncColumnMaps
- `src/05_Integrations.gs` — Added `case 'checkin'` and `case 'resources'` to doGet switch + 3 new functions: `getWebAppResourcesList()`, `getWebAppResourcesHtml()`, `getWebAppCheckInHtml()`
- `src/10b_SurveyDocSheets.gs` — Added `createResourcesSheet()` with starter content and validation
- `PHASE2_PLAN.md` — Created
- `dist/ConsolidatedDashboard.gs` — Rebuilt (62,121 lines / 2,569 KB)

**Design Decisions:**
- Resources page uses warm serif typography (Fraunces) to convey authority/trust — it's a union tool
- Navy + earth tones (#1e3a5f, #fafaf9) instead of generic purple gradients
- Check-in page uses green theme to differentiate from other pages
- All existing routes, tabs, and pages completely untouched — additive only
- Resources sheet auto-creates with starter content when first accessed

**Parked for later (ranked):**
1. Bulk actions (flag/email/export)
2. Deadline calendar view
3. Grievance history for members
4. Welcome/landing page
5. Events page with Join Virtual button

---

## 🐛 ERRORS & FIXES LOG

### 2026-02-22 — CLASP Setup
- **Error:** `clasp push` → "Project contents must include a manifest file named appsscript"
- **Fix:** Copied `appsscript.json` into `dist/` folder (CLASP rootDir)
- **Error:** `clasp push` → "User has not enabled the Apps Script API"
- **Fix:** Enabled at https://script.google.com/home/usersettings

### 2026-02-23 — CI Not Triggering on Main
- **Error:** GitHub Actions workflow only triggered on `main` (lowercase), but branch is `Main` (uppercase)
- **Fix:** Updated `.github/workflows/build.yml` to include `Main` in push/PR triggers

---

## ✅ FEATURES & HOW THEY WORK

### Web App Dashboard (doGet)
- Served via Google Apps Script web app deployment
- URL format: `https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec`
- Mobile-optimized with viewport meta tags and responsive CSS
- Dark theme with Inter font family, Material Icons, Chart.js

### Steward Command Center (?mode=steward)
- Requires Google account authorization
- Shows PII (member contact info, full grievance details)
- 4-up KPI bar, filterable case list, bulk operations

### Member Dashboard (?mode=member)
- No PII visible to other members
- Members see only their own grievances
- Countdown timers, timeline, resources

### Member Self-Service Portal (?page=selfservice)
- Dual auth: Google account verification → PIN fallback
- Members can view their own info and file grievances

### Authentication Toggle
- `enableDashboardMemberAuth()` / `disableDashboardMemberAuth()` — admin menu toggle
- When enabled: all dashboard pages require member login
- When disabled: dashboard pages are openly accessible

### Build System
- `npm run build` — concatenates `src/*.gs` → `dist/ConsolidatedDashboard.gs`
- `npm run build:prod` — same but excludes `07_DevTools.gs` (removes test data seeding)
- `npm run deploy` — lint + test + build:prod + clasp push
- `npm run ci` — clean + lint + build + test:unit

### Security
- `escapeHtml()` for all dynamic HTML content
- `escapeForFormula()` for spreadsheet cell values
- `secureLog()` masks PII before logging
- `validateRole()` for role-based access control
- `validateWebAppRequest()` sanitizes URL parameters
- Security events logged to audit sheet

---

## 📝 NOTES FOR FUTURE LLMs

1. **Read CLAUDE.md first** — it has the most detailed architectural documentation including column constant rules, config write paths, and security patterns.
2. **Read this file second** — it has the deployment context, change history, and known issues.
3. **The `web-dashboard/` folder is dead code.** Do not try to deploy it or integrate it. The consolidated build handles everything.
4. **Never edit `dist/ConsolidatedDashboard.gs` directly.** Edit `src/*.gs` files and run `npm run build`.
5. **Test with `npm run ci`** before pushing.
6. **Deploy with `npm run deploy`** (includes lint + test + prod build + clasp push).
7. **Current file size is 2.4MB / 6MB limit.** If adding major features, monitor growth.
8. **The `doGet()` function is in `src/04e_PublicDashboard.gs`** (or check `src/` files for the source location).
