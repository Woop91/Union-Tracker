# CLAUDE.md — DDS-Dashboard

## ⛔ CRITICAL RULES — READ FIRST

1. **No function may delete or overwrite manually entered data.** Manually entered = imported via import function OR typed into cells by user. Functions may only: append new data, write to system-generated fields, clear auto-generated content.
2. **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values. Always read from Config tab or constants. Config tab = single source of truth.
3. **Column constants are 1-indexed** (matching `getRange()`). For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`.
4. **All HTML must use `escapeHtml()`** for dynamic values. Defined in `00_Security.gs:130`. No exceptions. Do not redefine it.
5. **`dist/` contains individual `.gs` + `.html` files for clasp push.** Never edit dist/ directly. Edit `src/`, then build. GAS needs separate `.html` files for `createTemplateFromFile()` / `createHtmlOutputFromFile()`. `--prod` flag strips `07_DevTools.gs`.
6. **After any merge/work to Main:** remind user to run `clasp push`. Agent cannot run clasp remotely.

## Permissions

Claude has **full, pre-authorized access** to read, edit, create, delete any files in both **DDS-Dashboard** and **Union-Tracker** — local and GitHub remotes. This includes all file operations, git operations, clasp push, build/lint/test/deploy scripts. No permission needed.

For anything **outside** these two repos, ask user first.

## Repos & Sync

- **DDS-Dashboard** (private): Primary repo. Branches: `Main` (production), `staging` (pre-deploy mirror).
- **Union-Tracker** (public): Mirror of DDS. Branches: `Main`, `staging`.
- DDS Apps Script ID: `[REDACTED — SEE DDS-DASHBOARD REPO]`
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
npm run coverage       # Jest with coverage report
npm run ci             # clean + lint + build + test:unit
npm run deploy         # lint + test:unit + build:prod + clasp push
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
| `08a-d` | Sheet utils | Setup, search, forms, audit formulas |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Entry, sheet creation, forms, sync |
| `11_-17_` | Features | CommandHub, self-service, meetings, events, etc. |
| `19_-28_` | Web Dashboard | Auth, config, data service, app, portals, workload, QA, timeline, failsafe |

### Column Constants

| Constant | Source | Notes |
|----------|--------|-------|
| `SHEETS` / `SHEET_NAMES` | `01_Core.gs` | Identical alias |
| `GRIEVANCE_COLS` | `buildColsFromMap_(GRIEVANCE_HEADER_MAP_)` | 1-indexed |
| `MEMBER_COLS` | `buildColsFromMap_(MEMBER_HEADER_MAP_)` | 1-indexed |
| `CONFIG_COLS` | `buildColsFromMap_(CONFIG_HEADER_MAP_)` | 1-indexed |

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

- Jest 29, custom GAS mocks in `test/`
- `architecture.test.js` + `modules.test.js` — structural invariants
- `columns.test.js` — column constant consistency

## Git Conventions

- Conventional Commits: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`
- Pre-commit: `.husky/pre-commit` runs lint-staged + build verification
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
