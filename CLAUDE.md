# CLAUDE.md — Union Steward Dashboard

## Project Overview

Google Apps Script application (30 `.gs` source files, ~59K lines) for union steward grievance tracking, member management, and reporting. Deployed to Google Sheets via `clasp`.

## Commands

```bash
npm run build          # Concatenate src/ → dist/ConsolidatedDashboard.gs
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

### Build Pipeline

`build.js` concatenates all `src/*.gs` files (in sorted order) into `dist/ConsolidatedDashboard.gs`. The `--prod` flag strips `07_DevTools.gs`.

**Important:** When editing `src/` files, the `dist/` file must also be kept in sync. Either re-run `npm run build` or manually apply the same changes to both locations.

### Column Constants

All column constants are **1-indexed** (matching Google Sheets `getRange()` conventions):

| Constant | Source | Notes |
|----------|--------|-------|
| `SHEETS` / `SHEET_NAMES` | `01_Core.gs` | `SHEET_NAMES = SHEETS` — identical alias |
| `GRIEVANCE_COLS` | `buildColsFromMap_(GRIEVANCE_HEADER_MAP_)` | 1-indexed for `getRange()` |
| `MEMBER_COLS` | `buildColsFromMap_(MEMBER_HEADER_MAP_)` | 1-indexed for `getRange()` |
| `CONFIG_COLS` | `buildColsFromMap_(CONFIG_HEADER_MAP_)` | 1-indexed for `getRange()` |

**Rule:** For `getRange()` calls, use `*_COLS` values directly. For array access on `getValues()` data, subtract 1: `row[GRIEVANCE_COLS.STATUS - 1]`. The legacy 0-indexed constants (`GRIEVANCE_COLUMNS`, `MEMBER_COLUMNS`) have been removed.

## Security Patterns

### HTML Output — MANDATORY

Every dynamic value inserted into HTML **must** be escaped. No exceptions.

```javascript
// In HTML strings: use escapeHtml() (defined in 00_Security.gs)
'<td>' + escapeHtml(String(userName)) + '</td>'

// In window.open() URLs: use JSON.stringify()
'<script>window.open(' + JSON.stringify(url) + ', "_blank");</script>'

// In <a href>: validate URL scheme + escapeHtml()
var safeUrl = /^https:\/\/docs\.google\.com\//.test(url) ? url : '';
'<a href="' + escapeHtml(safeUrl) + '">Link</a>'
```

**Never** do this:
```javascript
// BAD — XSS via string concatenation
'<td>' + memberName + '</td>'
'<script>window.open("' + url + '");</script>'
```

### `escapeHtml()` Location

Defined once in `src/00_Security.gs:130`. For HTML served in web apps, a minified inline version is emitted by `getSecurityScript()` at `00_Security.gs:664`. Do **not** redefine it in other files.

### Other Security Functions

- `escapeForFormula(value)` — prevents formula injection in spreadsheet cells
- `secureLog(context, message, data)` — masks PII (emails, names, phones) before logging
- `validateRole(email, requiredRole)` — role-based access control for web app pages

## Coding Conventions

- **`var` vs `const`/`let`**: The codebase uses a mix. Older code uses `var`; newer code uses `const`/`let`. Match the style of the file you're editing.
- **Function naming**: `camelCase` for regular functions, `UPPER_CASE` for destructive admin functions (e.g., `NUCLEAR_WIPE_GRIEVANCES`), trailing underscore for private helpers (e.g., `buildColsFromMap_()`, `getConfigValue_()`).
- **Error handling**: Functions that return results use `{ success: true, ... }` / `errorResponse(message)`. Functions called from UI use try/catch with `SpreadsheetApp.getUi().alert()`.
- **Audit logging**: Use `logAuditEvent(eventType, details)` (in `06_Maintenance.gs:1463`) for security/admin events. Use `logIntegrityEvent(eventType, details)` for data integrity events. These two functions have conflicting schemas on the same sheet — this is known debt (see CODE_REVIEW.md Finding 21).

## Config Sheet Write Paths — CRITICAL

The Config sheet has **multiple code paths** that write to it. This is a known source of recurring bugs. When modifying any Config-writing code, you **must** check all paths for consistency.

### Rules

1. **Column lookup**: Always use `CONFIG_COLS.*` constants, never hardcoded column numbers. These constants may be updated at runtime by `syncColumnMaps()`.
2. **Write method**: Always use `addToConfigDropdown_(configCol, value)` for adding dropdown/list values. Never use `getRange(lastRow + 1, col).setValue()` — it scatters data past the end of other columns.
3. **Dynamic maps only**: Use `DROPDOWN_MAP` and `MULTI_SELECT_COLS` (rebuilt by `syncColumnMaps()`) for column lookups. Do **not** use `JOB_METADATA_FIELDS` for Config writes — it captures column numbers at load time and goes stale when `syncColumnMaps()` shifts positions.
4. **Rows 1-2 are structure**: Row 1 = section headers, Row 2 = column headers. Never overwrite these with data or formulas. Data starts at row 3.

### Canonical Write Paths

| Function | File | Lookup Mechanism | Purpose |
|----------|------|-----------------|---------|
| `addToConfigDropdown_()` | `02_DataManagers.gs` | Direct `configCol` param | **Core helper** — all dropdown writes go through here |
| `syncDropdownToConfig_()` | `10_Main.gs` | `DROPDOWN_MAP` + `MULTI_SELECT_COLS` | onEdit bidirectional sync |
| `populateConfigFromSheetData()` | `10a_SheetCreation.gs` | `DROPDOWN_MAP` + `MULTI_SELECT_COLS` | Bulk backfill from sheet data |
| `seedConfigDefault_()` | `10a_SheetCreation.gs` | Direct `CONFIG_COLS.*` param | Initial setup defaults |
| `restoreConfigFromSheetData_()` | `07_DevTools.gs` | `DROPDOWN_MAP` + `MULTI_SELECT_COLS` | Restore after data loss |
| `syncStewardStatus()` | `02_DataManagers.gs` | `CONFIG_COLS.STEWARDS` | Steward↔Member bidirectional |

### Yes/No Validation — Hardcoded, Not Config-Driven

The `YES_NO` Config column was **removed** to eliminate a contamination risk (multiple columns sharing one Config source). All Yes/No columns now use hardcoded `['Yes', 'No']` validation applied directly in `setupDataValidations()` (`08a_SheetSetup.gs`):

- `MEMBER_COLS.IS_STEWARD`
- `MEMBER_COLS.INTEREST_LOCAL`
- `MEMBER_COLS.INTEREST_CHAPTER`
- `MEMBER_COLS.INTEREST_ALLIED`

These columns are **not** in `DROPDOWN_MAP` — they are deliberately excluded from bidirectional Config sync. `migrateRemoveYesNoColumn_()` in `10a_SheetCreation.gs` handles deleting the orphaned column on existing sheets when `CREATE_DASHBOARD` is re-run.

### Known Debt

- `syncNewValueToConfig()` in `09_Dashboards.gs` is a legacy wrapper that now delegates to `syncDropdownToConfig_()`. Do not add logic to it — use `syncDropdownToConfig_()` directly.
- `seedConfigData()` in `07_DevTools.gs` is a DevTools-only seeder — it is excluded from production builds.

## Testing

- **Framework**: Jest 29 with custom GAS mocks
- **Test location**: `test/` directory, files named `XX_ModuleName.test.js`
- **Run**: `npm run test:unit` (Jest only) or `npm run test` (lint + build + Jest)
- **Architecture tests**: `test/architecture.test.js` and `test/modules.test.js` verify structural invariants (e.g., escapeHtml not redefined, all modules present in build)
- **Column tests**: `test/columns.test.js` validates column constant consistency

## Git Conventions

- **Commits**: Conventional Commits enforced by commitlint (`feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`)
- **Pre-commit hook**: `.husky/pre-commit` runs lint-staged + build verification
- **Main branch**: `Main` (not `master`)
- **Feature branches**: `claude/<description>-<id>`

---

## Code Review Methodology

**The canonical review document is `CODE_REVIEW.md`.** Archived reviews in `docs/archived-reviews/` are outdated and must not be used.

### Why These Rules Exist

Past reviews produced contradictory results: one said XSS was "FIXED" with Security 9/10 while the code still had 13 critical XSS vulnerabilities. This happened because reviews checked patterns superficially, assumed fixes applied everywhere, and wrote optimistic summaries without line-by-line verification. These rules exist to prevent that from happening again.

### Mandatory Review Process

#### 1. Never Claim Something Is Fixed Without Verifying

- **Do not** mark a finding as "FIXED" based on the existence of a helper function (e.g., `escapeHtml` existing does not mean it's used everywhere).
- **Do not** mark a finding as "FIXED" because a previous review said so.
- For every "FIXED" claim, cite the **exact file, line number, and code** that proves it.

#### 2. Line-by-Line Verification Required

For every file reviewed, you must:
- Read the actual code (not just function signatures or JSDoc)
- Check every location where dynamic data enters HTML, SQL, URLs, or sheet formulas
- Check every `getRange()` call for correct column constant usage (1-indexed vs 0-indexed)
- Check every `getLastRow() - 1` pattern for empty-sheet crashes
- Verify error handling exists in catch blocks

#### 3. Cross-Reference Across Files

The same pattern often appears in multiple files. When you find a vulnerability:
- Search the entire codebase for the same pattern (`grep` for `innerHTML`, `window.open`, `'<td>' +`, etc.)
- Report **every** instance, not just the first one
- Do not assume other files are clean because one file was fixed

#### 4. Severity Ratings Must Be Evidence-Based

| Severity | Criteria |
|----------|----------|
| CRITICAL | Exploitable security vulnerability (XSS, injection, auth bypass) with a concrete attack vector |
| HIGH | Bug that causes data corruption, crashes, or security weakness under realistic conditions |
| MEDIUM | Bug that causes incorrect behavior, performance issues, or maintainability problems |
| LOW | Code quality, style, dead code, minor inconsistencies |

**Do not** rate something CRITICAL without describing how it could be exploited. **Do not** rate something LOW if it causes data corruption.

#### 5. No Inflated Scores

- Do not assign numeric scores (e.g., "Security 9/10") — they create false confidence
- State findings plainly: "X critical issues remain, Y were fixed since last review"
- If you're uncertain whether something is fixed, say so — "unable to verify" is better than "FIXED"

#### 6. Review Report Structure

Every review must include:
1. **Date, version, scope** — what was reviewed
2. **Findings table** — every finding with file, line, severity, category
3. **For each finding**: description, vulnerable code snippet, suggested fix
4. **Unfixed items from previous reviews** — explicitly list what is still open from `CODE_REVIEW.md`
5. **Verification notes** — for any "FIXED" claim, cite the proof

#### 7. Update CODE_REVIEW.md

After a review is complete:
- Update `CODE_REVIEW.md` with new findings
- Mark genuinely fixed items with the commit hash that fixed them
- Do not create separate review documents — there is one canonical review file
