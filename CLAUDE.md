# CLAUDE.md — Union-Tracker

## ⛔ CRITICAL RULES — READ FIRST

1. **No function may delete or overwrite manually entered data.** Manually entered = imported via import function OR typed into cells by user. System-generated = anything written by code functions (except import function output). Functions may only: append new data, write to system-generated fields, clear auto-generated content.
2. **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values. Always read from Config tab or constants.
3. **Column constants are 1-indexed** (matching `getRange()`). For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`.
4. **All HTML must use `escapeHtml()`** for dynamic values. Defined in `00_Security.gs`. No exceptions.
5. **`dist/` contains individual `.gs` + `.html` files for clasp push.** Never edit dist/ directly. Edit `src/*.gs` and `src/*.html`, then copy to dist/. **Do NOT use the consolidated single-file build** (`ConsolidatedDashboard.gs`) — clasp must push individual files because GAS needs separate `.html` files for `HtmlService.createTemplateFromFile()` and `createHtmlOutputFromFile()`.
6. **This repo is PUBLIC.** No credentials, tokens, API keys, or DDS Script ID may ever appear here.

## Repos & Sync

- **Union-Tracker** (public): This repo. Receives syncs from DDS `Main` into UT `staging`.
- **DDS-Dashboard** (private): Primary/source repo. Default branch: `Main`.
- Sync flow: `DDS Main → UT staging → (user manages) → UT dev → UT main`
- UT Apps Script ID: `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`
- **DDS Script ID must NEVER appear in this repo.**

### What's Excluded from UT
- `src/18_WorkloadTracker.gs` — DDS-only member caseload module
- `src/WorkloadTracker.html` — DDS-only portal template
- Files referencing WT use `typeof` guards to handle its absence gracefully.
- All "steward workload" references (grievance case counts per steward) are kept — core functionality.
- See `SYNC-LOG.md` for full exclusion registry.

### Fallback
If GitHub is unreachable, work on local repo. Add a note in AI-REFERENCE.md: what was done, what needs pushing, which branches affected.

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
| `19_-24_` | Web Dashboard | Auth, config reader, data service, app, portal sheets, weekly questions |

**Note:** There is no `18_` in this repo (Workload Tracker is DDS-only).

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
// GOOD
'<td>' + escapeHtml(String(userName)) + '</td>'
'<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>'

// BAD — XSS
'<td>' + memberName + '</td>'
```

`escapeHtml()` is in `src/00_Security.gs:130`. Do not redefine it.

### Other Security Functions
- `escapeForFormula(value)` — prevents formula injection
- `secureLog(context, message, data)` — masks PII before logging
- `validateRole(email, requiredRole)` — role-based access control

## Config Sheet Write Paths — CRITICAL

1. **Always use `CONFIG_COLS.*`**, never hardcoded column numbers.
2. **Always use `addToConfigDropdown_()`** for adding values.
3. **Use `DROPDOWN_MAP` and `MULTI_SELECT_COLS`** (rebuilt by `syncColumnMaps()`).
4. **Rows 1-2 are structure.** Data starts at row 3.

## Coding Conventions

- `var` vs `const`/`let`: Match the style of the file you're editing.
- `camelCase` for functions, `UPPER_CASE` for destructive admin functions, trailing `_` for private helpers.
- Error handling: `{ success: true, ... }` / `errorResponse(message)`.
- Audit: `logAuditEvent()` for security events, `logIntegrityEvent()` for data events.

## Testing

- Jest 29 with custom GAS mocks in `test/`
- `test/architecture.test.js` + `test/modules.test.js` verify structural invariants
- `test/columns.test.js` validates column constant consistency

## Git Conventions

- Conventional Commits: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`
- Main branch: `Main` (capital M)

## Code Review

Canonical review: `CODE_REVIEW.md`. Never claim "FIXED" without citing proof.

## Reference Documents

- `AI_REFERENCE.md` — LLM context doc (never delete, only append)
- `SYNC-LOG.md` — DDS↔UT sync history and exclusion registry
- `CODE_REVIEW.md` — canonical security/code review
