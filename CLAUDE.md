# CLAUDE.md — DDS-Dashboard

## ⛔ CRITICAL RULES — READ FIRST

1. **No function may delete or overwrite manually entered data.** Manually entered = imported via import function OR typed into cells by user. Functions may only: append new data, write to system-generated fields, clear auto-generated content.
2. **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values. Always read from Config tab or constants. Config tab = single source of truth.
3. **Column constants are 1-indexed** (matching `getRange()`). For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`.
4. **All HTML must use `escapeHtml()`** for dynamic values. Defined in `00_Security.gs:130`. No exceptions. Do not redefine it.
5. **`dist/` contains individual `.gs` + `.html` files for clasp push.** Never edit dist/ directly. Edit `src/`, then build. GAS needs separate `.html` files for `createTemplateFromFile()` / `createHtmlOutputFromFile()`. `--prod` flag strips `07_DevTools.gs`.
6. **After any merge/work to Main:** remind user to run `clasp push`. Agent cannot run clasp remotely.
7. **⚠️ Adding a scope to `appsscript.json` REQUIRES re-authorization of the deployed web app.** GAS does NOT auto-authorize new scopes in existing deployments. After `clasp push`, the admin must: Apps Script editor → Deploy → Manage Deployments → create a new version → authorize all scopes. Failure to do this causes the new scope's services to throw auth errors silently (they reach the catch block and return `{success:false}`). After any scope change, run the `emailsend` test suite via the TestRunner tab to verify. The CI `scope-change-guard` job will also flag this on every push to Main.

## Tech Stack & Version

- **Current version:** 4.25.10 (check `package.json` and `src/01_Core.gs:VERSION_INFO` for latest)
- **Runtime:** Google Apps Script V8 (ES2020), timezone `America/New_York`
- **Build:** Node.js >=18, ESLint v9 (flat config), Jest 29, Husky 9
- **Deployment:** Google CLASP (`@google/clasp`), GitHub Actions CI
- **Source:** 42 `.gs` files + 8 `.html` files (~94K lines)
- **Tests:** 59 Jest test files + GAS-native TestRunner (6 suites, 48+ tests)
- **OAuth Scopes:** 10 (spreadsheets, drive, docs, gmail, calendar, external requests, userinfo, scriptapp, container.ui, send_mail)

## Permissions

Claude has **full, pre-authorized access** to read, edit, create, delete any files in both **DDS-Dashboard** and **Union-Tracker** — local and GitHub remotes. This includes all file operations, git operations, clasp push, build/lint/test/deploy scripts. No permission needed.

For anything **outside** these two repos, ask user first.

## Repos & Sync

- **DDS-Dashboard** (private): Primary repo. Branches: `Main` (production), `staging` (pre-deploy mirror).
- **Union-Tracker** (public): Mirror of DDS. Branches: `Main`, `staging`.
- DDS Apps Script ID: `[REDACTED-DDS-SCRIPT-ID]`
- UT Apps Script ID: `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`
- **DDS Script ID must NEVER appear in UT** (public repo).
- DDS and UT are **identical**. No file exclusions remain (standalone WT removed in v4.20.0, now integrated into SPA via `25_WorkloadService.gs`). `typeof` guards in `index.html`, `member_view.html`, `03_UIComponents.gs` kept as defensive coding. If drift is found, it's a bug.

### Branch Commit Order — MANDATORY

All commits go to `Main` first. **NEVER commit directly to `staging`.**

```
1. git checkout Main → commit → push origin Main
2. git checkout staging → merge Main --no-edit → push origin staging
3. Repeat for UT (redact DDS Script ID from AI_REFERENCE.md if re-copied)
```

Committing to `staging` first creates divergent histories. If it happens: merge `origin/staging` into `Main`, push `Main`, fast-forward `staging`.

### Sync Flow
```
DDS Main → DDS staging
DDS Main → UT Main → UT staging
```
Every DDS Main commit syncs to UT Main, then both stagings. Fetch origin before any push.

### Version Tracking — MANDATORY
- Semver on every meaningful commit. Update `VERSION_INFO` in `src/01_Core.gs`, `package.json`, `CHANGELOG.md` together.
- Tag the commit (e.g., `git tag v4.18.2`). Push tags: `git push origin --tags`.
- No version bump = incomplete commit.

### Parity Enforcement
- After every push: `git status` + `git log --oneline -3` to verify local = remote.
- Unresolvable merge conflict → `TODO-MERGE` in `AI_REFERENCE.md` with timestamp. Resolve ASAP.
- DDS GitHub, DDS local, and GAS must all match after every operation.

### No-Assumptions Policy
- Never assume file contents, function behavior, config values, or repo state. **Read first, then act.**
- Prior agent assumptions → re-examine, correct, document in `AI_REFERENCE.md`.
- Every repo-altering action must be logged: what, why, resulting version.

### Fallback (GitHub unreachable)
Work locally. Add timestamped note in `AI_REFERENCE.md`: what was done, what needs pushing/syncing, enough context for a follow-on agent to continue.

## Commands

```bash
npm run build          # src/*.gs + src/*.html → dist/ (individual files)
npm run build:prod     # Production build (excludes DevTools)
npm run lint           # ESLint all src/*.gs
npm run lint:fix       # ESLint with auto-fix
npm run test           # lint + build + jest (full pipeline)
npm run test:unit      # Jest only
npm run test:guards    # Deploy guards + SPA integrity checks only
npm run coverage       # Jest with coverage report
npm run ci             # clean + lint + build + test:unit
npm run deploy         # lint + test:guards + test:unit + build:prod + clasp push
npm run deploy:dev     # lint + build + clasp push --force (dev shortcut)
npm run watch          # Watch mode (rebuild on src/ changes)
npm run clean          # Remove dist/ directory
npm run sync-org-chart # Sync org chart from Google Drive
npm run check:scopes   # Check for OAuth scope changes in appsscript.json
```

## Architecture

### File Load Order

Files numbered to control GAS execution order:

| Prefix | Layer | Purpose |
|--------|-------|---------|
| `00_` | Foundation | DataAccess, Security |
| `01_` | Core | Constants, config, utilities |
| `02_` | Data | CRUD operations |
| `03_` | UI | Menus, dialogs |
| `04a-e` | UI modules | Menus, accessibility, dashboards |
| `05_` | Integrations | Drive, Calendar, Email, Web App |
| `06_` | Maintenance | Admin, undo/redo, snapshots, audit |
| `07_` | DevTools | Dev-only (excluded in prod) |
| `08a-e` | Sheet utils | Setup, search, forms, audit formulas, survey engine |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Entry, sheet creation, survey docs, forms, sync |
| `11_-17_` | Features | CommandHub, self-service, meetings, events, correlation engine |
| `19_-28_` | Web Dashboard | Auth, config, data service, app, portals, workload, QA, timeline, failsafe |
| `29_` | Migrations | Column schema auto-migration |
| `30_` | TestRunner | GAS-native test runner framework (in-spreadsheet tests) |

### Column Constants

| Constant | Source | Notes |
|----------|--------|-------|
| `SHEETS` / `SHEET_NAMES` | `01_Core.gs` | Identical alias |
| `GRIEVANCE_COLS` | `buildColsFromMap_(GRIEVANCE_HEADER_MAP_)` | 1-indexed |
| `MEMBER_COLS` | `buildColsFromMap_(MEMBER_HEADER_MAP_)` | 1-indexed |
| `CONFIG_COLS` | `buildColsFromMap_(CONFIG_HEADER_MAP_)` | 1-indexed |

### HTML Files (SPA & Views)

| File | Purpose |
|------|---------|
| `index.html` | SPA entry point — unified steward dashboard |
| `steward_view.html` | Steward command center |
| `member_view.html` | Member self-service dashboard |
| `auth_view.html` | Login/authentication page |
| `error_view.html` | Error display fallback |
| `styles.html` | Shared CSS (included via `createTemplateFromFile()`) |
| `org_chart.html` | Organizational chart view |
| `poms_reference.html` | POMS reference documentation |

### Build System (`build.js`)

- Validates `.gs` syntax and `<script>` blocks in HTML before copying
- Copies individual files from `src/` → `dist/` (not consolidated)
- `--prod` excludes `07_DevTools.gs`; warns if prod build exceeds 52 files (limit 60)
- Auto-cleans `dist/` before build to prevent orphaned files
- Also copies `appsscript.json` and `AI_REFERENCE.md` to `dist/`

## Error Handling — MANDATORY

Enforced by architecture tests A5–A8.

1. **GAS entry points (`doGet`, `onOpen`, `onEdit`, triggers):** Always wrap in try/catch. Unhandled throw → generic Google error or silent failure. Catch must call `_serveFatalError()` (zero-dependency fallback).

2. **No cascading errors in catch blocks.** Never call a function in catch that could re-throw. Wrap secondary calls in their own try/catch or use hardcoded fallbacks.

## GAS Simple Trigger Constraints — MANDATORY

`onOpen`, `onEdit`, `onChange`, `onFormSubmit` are **simple triggers**. They run with restricted authorization. Violating these rules causes **silent failure** — no error shown, no stack trace.

### ⛔ NEVER call in a simple trigger:
| Forbidden | Why |
|---|---|
| `ScriptApp.getProjectTriggers()` | Requires auth not granted to simple triggers |
| `ScriptApp.newTrigger()` | Same |
| `ScriptApp.deleteTrigger()` | Same |
| `MailApp.sendEmail()` | Requires auth |
| `GmailApp.*` | Requires auth |
| Any OAuth service | Requires auth |

### ✅ Simple triggers may ONLY:
- Call `SpreadsheetApp.getUi()` to build menus
- Call `SpreadsheetApp.getActiveSpreadsheet()` for read-only sheet access
- Call pure utility functions that do not touch ScriptApp or external services

### ✅ Pattern: deferred init via installable trigger
If `onOpen` needs to do anything beyond menu creation, install a **separate installable onOpen trigger** once via a setup function:
```js
// Run once from menu after deploy — NOT from onOpen itself
function setupOpenDeferredTrigger() {
  // Remove existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onOpenDeferred_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onOpenDeferred_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}
```

### ✅ Race condition rule: never use `finally` with async-like patterns
`finally` in GAS runs **synchronously and immediately** — it does not wait for time-based triggers. Never create a trigger and then clean it up in `finally`; the trigger will be deleted before it fires.

   ```js
   // GOOD: catch (err) { try { cfg = ConfigReader.getConfig(); } catch(_) { cfg = {orgName:'Dashboard'}; } }
   // BAD:  catch (err) { return _serveError(ConfigReader.getConfig(), err.message); }
   ```

3. **`getActiveSpreadsheet()` null-guard** in web app files (19_–28_). Always check before chaining: `var ss = SpreadsheetApp.getActiveSpreadsheet(); if (!ss) return null;`

4. **`google.script.run` must have a failure handler.** Use `serverCall()` (in `index.html`) or explicit `.withFailureHandler()`. Missing handler = infinite spinner on error.

## Security Patterns

- `escapeHtml(value)` — XSS prevention for HTML output (`00_Security.gs:130`)
- `escapeForFormula(value)` — formula injection prevention
- `secureLog(context, message, data)` — masks PII before logging
- `validateRole(email, requiredRole)` — role-based access control

## Config Sheet Write Paths — CRITICAL

1. Always use `CONFIG_COLS.*`, never hardcoded column numbers.
2. Always use `addToConfigDropdown_()` to add values. Never `getRange(lastRow+1, col).setValue()`.
3. Use `DROPDOWN_MAP` and `MULTI_SELECT_COLS` (rebuilt by `syncColumnMaps()`). Not `JOB_METADATA_FIELDS` for writes.
4. Rows 1-2 are structure (section headers, column headers). Data starts row 3.

## Coding Conventions

- `var` vs `const`/`let`: Match existing file style.
- `camelCase` functions, `UPPER_CASE` destructive admin, trailing `_` private helpers.
- Return values: `{ success: true, ... }` / `errorResponse(message)`. Try/catch with `alert()` for UI.
- Audit: `logAuditEvent()` for security, `logIntegrityEvent()` for data.

## Testing

### Jest (local, CI)
- Jest 29, custom GAS mocks in `test/gas-mock.js` + `test/load-source.js`
- 59 test files covering all source modules
- Coverage thresholds: 70% lines, 60% branches, 70% functions, 70% statements
- Key structural tests:
  - `architecture.test.js` — entry point error handling, auth checks (A1–A8)
  - `modules.test.js` — module structure invariants
  - `columns.test.js` — column constant consistency
  - `deploy-guards.test.js` — deployment safety checks
  - `spa-integrity.test.js` — SPA HTML/JS integrity
  - `simple-trigger-lint.test.js` — simple trigger restriction enforcement
  - `auth-denial.test.js` — auth denial path coverage

### GAS-Native TestRunner (in-spreadsheet)
- `30_TestRunner.gs` — 6 suites, 48+ tests run inside Google Apps Script
- Accessible via the TestRunner tab in the SPA dashboard
- Supports daily trigger (6 AM scheduled runs) via `setupOpenDeferredTrigger()`

## CI/CD

### GitHub Actions (`.github/workflows/build.yml`)
Three jobs run on every push/PR:
1. **lint-and-build** — ESLint, build, `npm audit`, Jest, coverage upload, size report (warns >3MB, limit ~6MB)
2. **code-quality** — file/function counts, duplicate function detection (fails on dupes), TODO/FIXME tracking
3. **scope-change-guard** — runs on push to `Main` only; **fails if OAuth scopes are added** (requires re-auth)

### Husky Git Hooks
| Hook | Runs | Purpose |
|------|------|---------|
| `pre-commit` | `lint-staged` + `npm run build` + deploy guards | Lint staged `.gs` files, verify build, run deploy/SPA integrity tests |
| `commit-msg` | `commitlint` | Enforce conventional commit format |
| `pre-push` | `check-scope-change.js` | Block push if OAuth scopes changed |

## Git Conventions

- Conventional Commits enforced by commitlint: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`
- Pre-commit: `.husky/pre-commit` runs lint-staged + build verification + deploy guards
- Both repos use capital `Main`. Branches: `Main` + `staging`.
- Every code-change commit must include version bump.

## Code Review — STRICT

Canonical: `CODE_REVIEW.md`. Archived reviews in `docs/archived-reviews/` are **outdated — do not reference**.

1. Never claim "FIXED" without proof — cite exact file, line, corrected code.
2. Search entire codebase for every vulnerability pattern — not just the first file.
3. No inflated scores. Broken = say broken.

## Standing Rules

### Badge Refresh on Mutations
Any frontend action writing to Grievance Log or Steward Tasks **must** call `_refreshNavBadges()` on success. Debounced at 600ms. Always guard: `if (typeof _refreshNavBadges === 'function') _refreshNavBadges();`

### Auth on Every `data*` Wrapper
Every `function data*` in `21_WebDashDataService.gs` must begin with auth. No exceptions.
```js
var e = _resolveCallerEmail(); if (!e) return <safe_empty>;   // any user
var s = _requireStewardAuth(); if (!s) return <safe_empty>;   // steward-only
```
No auth check = open anonymous endpoint.

## Reference Documents

- `AI_REFERENCE.md` — LLM context (never delete, only append)
- `SYNC-LOG.md` — DDS↔UT sync history
- `CODE_REVIEW.md` — canonical security review
- `COLUMN_ISSUES_LOG.md` — recurring column bugs (READ before column work)
- `FEATURES.md` — feature docs
- `DEVELOPER_GUIDE.md` — developer onboarding
