# CLAUDE.md — SolidBase

## ⛔ CRITICAL RULES — READ FIRST

1. **No function may delete or overwrite manually entered data.** Functions may only: append new data, write to system-generated fields, clear auto-generated content.
2. **Everything must be dynamic.** Never hardcode sheet names, column positions, org names, unit names, or config values. Always read from Config tab or constants.
3. **Column constants are 1-indexed** (matching `getRange()`). For array access: `row[GRIEVANCE_COLS.STATUS - 1]`.
4. **All HTML must use `escapeHtml()`** (`00_Security.gs:130`). No exceptions. Do not redefine it.
5. **Never edit `dist/` directly.** Edit `src/`, then `npm run build`. `--prod` strips `07_DevTools.gs`.
6. **After any merge to Main:** remind user to run `clasp push`.
7. **Adding a scope to `appsscript.json` REQUIRES re-authorization.** Admin must: Apps Script → Deploy → new version → authorize. CI `scope-change-guard` flags this.

## Tech Stack

- **Version:** 4.32.0 (`package.json` + `src/01_Core.gs:VERSION_INFO`)
- **Runtime:** Google Apps Script V8 (ES2020), `America/New_York`
- **Build:** Node >=18, ESLint v9, Jest 29, Husky 9
- **Deploy:** CLASP, GitHub Actions CI
- **Source:** 42 `.gs` + 8 `.html` (~94K lines) | **Tests:** 59 Jest + GAS TestRunner (48+ tests)
- **Scopes:** 10 (spreadsheets, drive, docs, gmail, calendar, external requests, userinfo, scriptapp, container.ui, send_mail)

## Permissions

Claude has **full access** to the **SolidBase** repo (local + remote). For anything outside, ask user.

## Repo

- **SolidBase** (public): `Main` (production), `staging` (pre-deploy mirror)
- SB Script ID: `1V6vzrczxUSYuiobdkKE64mbsZYznZHZwcI51juAtqQojy5Tz8q5zbiTl`

**Commit order:** Main first → merge to staging. Never commit directly to staging.

**Version tracking:** Semver on every commit. Update `VERSION_INFO` + `package.json` + `CHANGELOG.md` + git tag.

## Commands

```bash
npm run build          # src/ → dist/ (individual files)
npm run build:prod     # Production (excludes DevTools)
npm run lint           # ESLint src/*.gs
npm run test           # lint + build + jest
npm run test:unit      # Jest only
npm run ci             # clean + lint + build + test:unit
npm run deploy         # lint + tests + build:prod + clasp push
npm run deploy:dev     # lint + build + clasp push --force
```

## Architecture

See `AI_REFERENCE.md` for full architecture map, file load order, routing, and sheet structure.

**Key constants** (all 1-indexed, built via `buildColsFromMap_()`):
- `GRIEVANCE_COLS`, `MEMBER_COLS`, `CONFIG_COLS` — from header maps in `01_Core.gs`
- `SHEETS` — sheet name constants

## Error Handling — MANDATORY

1. **GAS entry points** (`doGet`, `onOpen`, `onEdit`, triggers): Always try/catch. Catch calls `_serveFatalError()`.
2. **No cascading errors in catch blocks.** Wrap secondary calls in their own try/catch.
3. **`getActiveSpreadsheet()` null-guard** in web app files (19_–28_): `var ss = SpreadsheetApp.getActiveSpreadsheet(); if (!ss) return null;`
4. **`google.script.run` must have a failure handler.** Use `serverCall()` or `.withFailureHandler()`.

## Simple Trigger Constraints

`onOpen`, `onEdit`, `onChange`, `onFormSubmit` run with restricted auth. Violations = **silent failure**.

**NEVER** call in a simple trigger: `ScriptApp.*Trigger*()`, `MailApp`, `GmailApp`, any OAuth service.
**MAY** call: `getUi()` for menus, `getActiveSpreadsheet()` read-only, pure utility functions.

For deferred work, use an **installable trigger** set up via a separate setup function (not from onOpen itself).

**Never use `finally` to clean up triggers** — it runs synchronously before the trigger fires.

## Security Patterns

- `escapeHtml(value)` — XSS prevention (`00_Security.gs:130`)
- `escapeForFormula(value)` — formula injection prevention
- `secureLog(context, message, data)` — masks PII before logging
- `validateRole(email, requiredRole)` — RBAC

## Config Sheet Write Paths — CRITICAL

1. Use `CONFIG_COLS.*`, never hardcoded column numbers.
2. Use `addToConfigDropdown_()` to add values. Never `getRange(lastRow+1, col).setValue()`.
3. Use `DROPDOWN_MAP` and `MULTI_SELECT_COLS`. Not `JOB_METADATA_FIELDS` for writes.
4. Rows 1-2 = structure. Data starts row 3.

## Coding Conventions

- `var`/`const`/`let`: Match existing file style
- `camelCase` functions, `UPPER_CASE` destructive admin, trailing `_` private helpers
- Returns: `{ success: true, ... }` / `errorResponse(message)`
- Audit: `logAuditEvent()` for security, `logIntegrityEvent()` for data

## Testing

- Jest 29 with GAS mocks (`test/gas-mock.js` + `test/load-source.js`)
- Coverage: 70% lines/functions/statements, 60% branches
- GAS TestRunner: `30_TestRunner.gs`, 6 suites, accessible via SPA TestRunner tab

## CI/CD

- **GitHub Actions:** lint-and-build, code-quality (dupe detection), scope-change-guard (Main only)
- **Husky hooks:** pre-commit (lint-staged + build + guards), commit-msg (commitlint), pre-push (scope check)
- **Conventional Commits** enforced: `feat:`, `fix:`, `docs:`, `refactor:`, `test:`, `chore:`

## Standing Rules

**Badge Refresh:** Any write to Grievance Log or Steward Tasks must call `_refreshNavBadges()` on success. Guard: `if (typeof _refreshNavBadges === 'function') _refreshNavBadges();`

**Auth on `data*` wrappers:** Every `function data*` in `21_WebDashDataService.gs` must begin with auth:
```js
var e = _resolveCallerEmail(); if (!e) return <safe_empty>;   // any user
var s = _requireStewardAuth(); if (!s) return <safe_empty>;   // steward-only
```

## Code Review — STRICT

Canonical: `CODE_REVIEW.md`. Archived reviews in `docs/archived-reviews/` are outdated.
1. Never claim "FIXED" without proof — cite file, line, code.
2. Search entire codebase for vulnerability patterns.
3. No inflated scores.

## Reference Documents

- `AI_REFERENCE.md` — Architecture, LLM context, protected code
- `CODE_REVIEW.md` — Canonical security review
- `COLUMN_ISSUES_LOG.md` — Recurring column bugs (READ before column work)
- `CHANGELOG.md` — Version history
- `FEATURES.md` — Feature docs
- `DEVELOPER_GUIDE.md` — Developer onboarding
- `SYNC-LOG.md` — Sync history
