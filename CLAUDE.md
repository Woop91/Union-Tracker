# CLAUDE.md — DDS-Dashboard

## ⛔ CRITICAL RULES — READ FIRST

1. **No function may delete or overwrite manually entered data.** Manually entered = imported via import function OR typed into cells by user. System-generated = anything written by code functions (except import function output). Functions may only: append new data, write to system-generated fields, clear auto-generated content.
2. **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values. Always read from Config tab or constants. The Config tab is the single source of truth.
3. **Column constants are 1-indexed** (matching `getRange()`). For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`.
4. **All HTML must use `escapeHtml()`** for dynamic values. Defined in `00_Security.gs`. No exceptions.
5. **`dist/` contains individual `.gs` + `.html` files for clasp push.** Never edit dist/ directly. Edit `src/*.gs` and `src/*.html`, then copy to dist/. **Do NOT use the consolidated single-file build** (`ConsolidatedDashboard.gs`) — clasp must push individual files because GAS needs separate `.html` files for `HtmlService.createTemplateFromFile()` and `createHtmlOutputFromFile()`.
6. **After any merge/work to Main:** remind user to run `clasp push` from the repo root. Agent cannot run clasp remotely.

## Permissions

Claude has **full, pre-authorized access** to read, edit, create, and delete any files in both the **DDS-Dashboard** and **Union-Tracker** repositories — local and GitHub remotes. Claude will not ask for permission for any operation within these two repos, including:
- Editing, creating, or deleting source files
- Git operations (commit, push, pull, branch, merge, etc.)
- Running `clasp push` to deploy either repo to Google Apps Script
- Running build, lint, test, and deploy scripts

For **any file or directory outside** the DDS-Dashboard and Union-Tracker repos, Claude **must ask the user for permission** before reading or modifying.

## Repos & Sync

- **DDS-Dashboard** (private): Primary repo. **Single branch: `Main` only.** All work happens here first.
- **Union-Tracker** (public): Mirror minus Workload Tracker. **Single branch: `Main` only.**
- DDS Apps Script ID: `18hHHX-4E_ykGCqu_EDwKCwqY9ycyRgPtOmguacsxnVZ4YsRh-YETODiu`
- UT Apps Script ID: `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`
- **DDS Script ID must NEVER appear in UT** (public repo).

### Default Working Directory
DDS `Main` on GitHub is always the default working branch. All code changes start here.

### Sync Flow
```
DDS Main (GitHub) → UT Main (GitHub)
         ↕
  Local DDS repo (fallback)
```
- Every commit to DDS Main must be synced to UT Main with Workload Tracker exclusions applied.
- Before any push, check the local DDS repo for changes not yet on GitHub. If found, pull or merge them first.

### Repo Differences
DDS and UT are **identical** except:
- DDS includes `src/18_WorkloadTracker.gs` and `src/WorkloadTracker.html` (Workload Tracker module).
- UT excludes these files and uses `typeof` guards where they're referenced.
- No other differences should exist. If drift is found, it's a bug — fix it.

### Fallback (GitHub unreachable)
Work on local DDS repo. **Mandatory:** add a timestamped note in `AI_REFERENCE.md` explaining:
- What was done and why
- What still needs pushing to GitHub
- What still needs syncing to UT Main
- Enough context for a follow-on agent to pick up exactly where work stopped

### Version Tracking — MANDATORY
- **Semver tagging on every meaningful commit.** Update `VERSION_INFO` in `src/01_Core.gs`, `package.json`, and `CHANGELOG.md` together.
- **Tag the commit** (e.g., `git tag v4.18.2`). Push tags: `git push origin --tags`.
- If version is not updated, the commit is incomplete.

### Parity Enforcement
- After every push: verify GitHub remote and local clone are identical (`git status`, `git log --oneline -3`).
- If a merge conflict cannot be resolved immediately: add a `TODO-MERGE` entry in `AI_REFERENCE.md` with conflict description and timestamp. Resolve at the earliest possible point.
- **DDS GitHub, DDS local, and Google Apps Script must reflect the same state after every operation.**
- **UT Main must match DDS Main** (minus Workload Tracker exclusions) after every sync.

### No-Assumptions Policy
- Never assume file contents, function behavior, config values, or repo state. **Read first, then act.**
- If prior work (by this agent or a previous one) contains assumptions: re-examine, correct, and document in `AI_REFERENCE.md`.
- Every repo-altering action must be logged: what changed, why, and resulting version.

### Workload Tracker Exclusion (UT only)
- **EXCLUDE from UT:** `src/18_WorkloadTracker.gs`, `src/WorkloadTracker.html`
- **MODIFY for UT (add typeof guards):** `src/03_UIComponents.gs`, `src/index.html`, `src/member_view.html`
- **Already safe:** `src/05_Integrations.gs`, `src/08a_SheetSetup.gs` (typeof guards exist)
- **Keep all "steward workload" references** — those are grievance case counts, not the WT module.
- If excluding WT would cause errors in UT, include the minimum needed to prevent breakage.

### Fallback
If GitHub is unreachable, work on local repo. Add a note in `AI_REFERENCE.md`: what was done, what needs pushing. Resolve on reconnection.

## Commands

```bash
npm run build          # Copy src/*.gs + src/*.html → dist/ (individual files)
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

### File Naming & Load Order

Files are numbered to control Google Apps Script's execution order:

| Prefix | Layer | Files |
|--------|-------|-------|
| `00_` | Foundation | `DataAccess.gs`, `Security.gs` |
| `01_` | Core | `Core.gs` — constants, config, utility functions |
| `02_` | Data | `DataManagers.gs` — CRUD operations |
| `03_` | UI | `UIComponents.gs` — menus, dialogs |
| `04a-e` | UI modules | Menus, accessibility, dashboards |
| `05_` | Integrations | Drive, Calendar, Email, Web App |
| `06_` | Maintenance | Admin tools, undo/redo, snapshots, audit |
| `07_` | DevTools | Development-only utilities (excluded in prod) |
| `08a-d` | Sheet utils | Setup, search, forms, audit formulas |
| `09_` | Dashboards | Dashboard rendering |
| `10_-10d` | Business logic | Main entry, sheet creation, forms, sync |
| `11_-17_` | Features | CommandHub, features, self-service, meetings, events, enhancements, correlation |
| `18_` | Workload Tracker | **DDS ONLY** — member caseload reporting |
| `19_-24_` | Web Dashboard | Auth, config reader, data service, app, portal sheets, weekly questions |

### Build Pipeline

`dist/` is a flat copy of `src/*.gs` + `src/*.html` — individual files, not consolidated. Clasp pushes individual files so GAS can resolve `createTemplateFromFile()` / `createHtmlOutputFromFile()`. The `--prod` flag strips `07_DevTools.gs`.

### Column Constants

| Constant | Source | Notes |
|----------|--------|-------|
| `SHEETS` / `SHEET_NAMES` | `01_Core.gs` | `SHEET_NAMES = SHEETS` — identical alias |
| `GRIEVANCE_COLS` | `buildColsFromMap_(GRIEVANCE_HEADER_MAP_)` | 1-indexed for `getRange()` |
| `MEMBER_COLS` | `buildColsFromMap_(MEMBER_HEADER_MAP_)` | 1-indexed for `getRange()` |
| `CONFIG_COLS` | `buildColsFromMap_(CONFIG_HEADER_MAP_)` | 1-indexed for `getRange()` |

**Rule:** For `getRange()` calls, use `*_COLS` directly. For array access, subtract 1.

## Security Patterns

### HTML Output — MANDATORY

```javascript
// GOOD: use escapeHtml() for all dynamic values
'<td>' + escapeHtml(String(userName)) + '</td>'
'<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>'

// BAD — XSS via string concatenation
'<td>' + memberName + '</td>'
```

`escapeHtml()` is in `src/00_Security.gs:130`. Do not redefine it.

### Other Security Functions
- `escapeForFormula(value)` — prevents formula injection
- `secureLog(context, message, data)` — masks PII before logging
- `validateRole(email, requiredRole)` — role-based access control

## Config Sheet Write Paths — CRITICAL

Multiple code paths write to Config. When modifying any Config-writing code, check all paths:

1. **Always use `CONFIG_COLS.*`**, never hardcoded column numbers.
2. **Always use `addToConfigDropdown_()`** for adding values. Never `getRange(lastRow + 1, col).setValue()`.
3. **Use `DROPDOWN_MAP` and `MULTI_SELECT_COLS`** (rebuilt by `syncColumnMaps()`). Do not use `JOB_METADATA_FIELDS` for writes.
4. **Rows 1-2 are structure.** Row 1 = section headers, Row 2 = column headers. Data starts at row 3.

## Coding Conventions

- `var` vs `const`/`let`: Match the style of the file you're editing.
- `camelCase` for functions, `UPPER_CASE` for destructive admin functions, trailing `_` for private helpers.
- Error handling: `{ success: true, ... }` / `errorResponse(message)` for return values. Try/catch with `alert()` for UI calls.
- Audit: `logAuditEvent()` for security events, `logIntegrityEvent()` for data events.

## Testing

- Jest 29 with custom GAS mocks in `test/`
- `test/architecture.test.js` + `test/modules.test.js` verify structural invariants
- `test/columns.test.js` validates column constant consistency

## Git Conventions

- Conventional Commits: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`
- Pre-commit hook: `.husky/pre-commit` runs lint-staged + build verification
- **Both repos: Single branch `Main` (capital M). No feature/dev/staging branches.**
- Every commit must include version bump if code changed.

## Code Review — STRICT

Canonical review: `CODE_REVIEW.md`. Archived reviews in `docs/archived-reviews/` are **outdated — do not reference them**.

Rules:
1. **Never claim "FIXED" without proof.** Cite exact file, line number, and the corrected code. If you can't point to it, it's not fixed.
2. **Search the entire codebase** for every vulnerability pattern found — not just the file where it was first spotted.
3. **No inflated scores.** State findings plainly. If something is broken, say it's broken.

## Reference Documents

- `AI_REFERENCE.md` — LLM context doc (never delete, only append)
- `SYNC-LOG.md` — DDS↔UT sync history and exclusion registry
- `CODE_REVIEW.md` — canonical security/code review
- `COLUMN_ISSUES_LOG.md` — recurring column bugs and fixes (READ if working on column-related code)
- `FEATURES.md` — feature documentation
- `DEVELOPER_GUIDE.md` — developer onboarding
